"""
Auction Catalog Parser using Google Document AI OCR Processor

Supports two completely separate parsing pipelines:
  - PRINTED lot numbers  → OCR-based lot number detection  (has_handwriting=False)
  - HANDWRITTEN lot numbers → year-based sequential numbering (has_handwriting=True)

Usage:
    from auction_parser import AuctionCatalogParser
    parser = AuctionCatalogParser(project_id, location, processor_id)
    df = parser.process_catalog_batch(CATALOGS, output_excel=OUTPUT_EXCEL)

PAGE FIELD
----------
  Page_PDF : 1-indexed page number within the PDF file.

IMAGE FIELD
-----------
  Image : "Yes" if a coin-image block (large blank region consistent with two
           circular coin photos side-by-side) is detected immediately above the
           lot on that PDF page; blank otherwise.
"""

import re
import bisect
from typing import List, Dict, Optional, Tuple
from pathlib import Path

from google.cloud import documentai_v1 as documentai
import pandas as pd


OUTPUT_COLUMN_ORDER = [
    'Catalog_Source',
    'Lot_No',
    'Page_PDF',
    'Catalog_Section',
    'Headline',
    'Image',
    'Year',
    'Grade',
    'Grading_Service',
    'Variety',
    'Rarity',
    'Short_Description',
    'Long_Description',
    'Pedigree',
    'Sale_Price',
    'Sold_To',
]

IMAGE_GAP_THRESHOLD = 0.10
IMAGE_PROXIMITY     = 0.08
TWO_COL_MIN_COUNT   = 3


class AuctionCatalogParser:
    """
    Parser for auction catalog PDFs using Google Document AI OCR.

    Two fully independent parsing pipelines:
        _pipeline_printed()     – catalogs with printed (OCR-readable) lot numbers
        _pipeline_handwritten() – catalogs where lot numbers are handwritten in
                                   the margins; OCR is unreliable for those numbers
    """

    def __init__(self, project_id: str, location: str, processor_id: str):
        self.project_id     = project_id
        self.location       = location
        self.processor_id   = processor_id
        self.client         = documentai.DocumentProcessorServiceClient()
        self.processor_name = self.client.processor_path(project_id, location, processor_id)

    # -----------------------------------------------------------------------
    # Public API
    # -----------------------------------------------------------------------

    def process_catalog_batch(
        self,
        catalogs: List[Tuple[str, bool]],
        output_excel: str = 'auction_catalog_parsed.xlsx',
    ) -> pd.DataFrame:
        all_lots: List[Dict] = []

        for pdf_path, has_handwriting in catalogs:
            catalog_name = Path(pdf_path).stem
            label        = "HANDWRITTEN" if has_handwriting else "PRINTED"
            print(f"\n{'='*70}")
            print(f"Catalog : {catalog_name}")
            print(f"Pipeline: {label} lot numbers")
            print(f"{'='*70}")

            document = self._ocr_document(pdf_path)

            if has_handwriting:
                lots = self._pipeline_handwritten(document, catalog_name)
            else:
                lots = self._pipeline_printed(document, catalog_name)

            print(f"-> Extracted {len(lots)} lots from '{catalog_name}'")
            all_lots.extend(lots)

        df = pd.DataFrame(all_lots)

        if df.empty:
            print("\nWARNING: No lots were extracted from any catalog!")
            return df

        existing_cols = [c for c in OUTPUT_COLUMN_ORDER if c in df.columns]
        df = df[existing_cols]
        df = df.sort_values(['Catalog_Source', 'Lot_No']).reset_index(drop=True)
        df.to_excel(output_excel, index=False)
        print(f"\nSaved {len(df)} total lots to: {output_excel}")
        return df

    # -----------------------------------------------------------------------
    # OCR
    # -----------------------------------------------------------------------

    def _ocr_document(self, file_path: str) -> documentai.Document:
        with open(file_path, 'rb') as fh:
            raw = fh.read()
        request = documentai.ProcessRequest(
            name=self.processor_name,
            raw_document=documentai.RawDocument(content=raw, mime_type='application/pdf'),
            skip_human_review=True,
        )
        print(f"  Sending to Document AI: {file_path}")
        return self.client.process_document(request=request).document

    # =======================================================================
    # PIPELINE A – PRINTED LOT NUMBERS  (from Z3)
    # =======================================================================

    def _pipeline_printed(
        self,
        document: documentai.Document,
        catalog_name: str,
    ) -> List[Dict]:
        """
        For catalogs with typeset lot numbers.

        Uses document.text so that text_anchor offsets used by the spatial
        index align correctly with regex match positions.  Two-column layout
        is detected per page; text is rebuilt in proper visual reading order.
        Sale_Price and Sold_To are always blank for printed catalogs.
        """
        print("  [PRINTED PIPELINE] Starting...")

        spatial_index = self._build_spatial_index(document)
        image_gaps    = self._compute_image_gaps(document)
        bold_ranges   = self._extract_bold_ranges(document)

        full_text, page_markers = self._build_ordered_text(document)

        if not full_text:
            print(f"  [PRINTED PIPELINE] ERROR: No text from '{catalog_name}'")
            return []
        print(f"  [PRINTED PIPELINE] Extracted {len(full_text):,} chars.")
        print(f"\n  [PRINTED PIPELINE] First 500 chars:")
        print(repr(full_text[:500]))

        lot_pattern = re.compile(r'^\s*(\d{1,4})(?:\s+|$)', re.MULTILINE)
        raw_matches = list(lot_pattern.finditer(full_text))
        print(f"  [PRINTED PIPELINE] Found {len(raw_matches)} candidate lot-number lines.")
        if raw_matches:
            print(f"  [PRINTED PIPELINE] First 10: {[m.group(1) for m in raw_matches[:10]]}")

        candidates: List[Tuple[str, str, int, int]] = []
        for idx, match in enumerate(raw_matches):
            lot_no      = match.group(1)
            desc_start  = match.end()
            desc_end    = raw_matches[idx + 1].start() if idx + 1 < len(raw_matches) else len(full_text)
            description = full_text[desc_start:desc_end].strip()
            description = re.sub(r'\s*\n\s*', ' ', description)

            if self._is_valid_lot(description):
                candidates.append((lot_no, description, desc_start, match.start()))

        print(f"  [PRINTED PIPELINE] After filtering: {len(candidates)} valid lots.")

        lots: List[Dict] = []
        for lot_no, description, desc_start, lot_start in candidates:
            lot_data = self._parse_lot_fields(lot_no, description, bold_ranges, 0)
            lot_data['Catalog_Source'] = catalog_name
            lot_data['Page_PDF']       = self._page_from_markers(lot_start, page_markers)
            lot_data['Image']          = 'Yes' if self._has_image_above_ordered(
                lot_start, page_markers, image_gaps) else None
            lot_data['Sale_Price']     = None
            lot_data['Sold_To']        = None
            lots.append(lot_data)

        return lots

    # =======================================================================
    # PIPELINE B – HANDWRITTEN LOT NUMBERS  (from v2 — mid-paragraph splitting)
    # =======================================================================

    def _pipeline_handwritten(
        self,
        document: documentai.Document,
        catalog_name: str,
    ) -> List[Dict]:
        """
        For catalogs where lot numbers are handwritten in the margins.

        Processes each PDF page independently with per-page two-column detection.
        Splits on EVERY year occurrence in the text (not just line starts), since
        OCR often merges multiple lots into one paragraph.  Sequential lot numbers
        are assigned in page + visual reading order.
        """
        print("  [HANDWRITTEN PIPELINE] Starting...")

        image_gaps  = self._compute_image_gaps(document)
        bold_ranges = self._extract_bold_ranges(document)

        # Match a year anywhere in text (not just line starts).
        # Requires at least 5 chars of text after the year.
        year_split = re.compile(
            r'(?<!\d)(1[6-9]\d{2}|20\d{2})(?!\d)\s+(?=\S.{4,})'
        )

        # Broad coin keyword set: American + foreign coins + general numismatic terms
        coin_kw = re.compile(
            r'\b(Mass\.|Rosa|Cent|Dollar|Token|Half|Penny|Shilling|Crown|Bust|Eagle|Liberty|'
            r'Wood|Copper|Silver|Gold|Proof|Rev\.|Obv\.|Fine|Uncirculated|Very|Good|Fair|Poor|'
            r'Franc|Pfennig|Thaler|Gulden|Florin|Mark|Lira|Peseta|Real|Reale|Ducat|Sovereign|'
            r'Grosch|Groschen|Denier|Ecu|Sou|Solidus|Denarius|Sestertius|Aureus|'
            r'Drachm|Obol|Stater|Tetradrachm|Didrachm|'
            r'Piastre|Dirhem|Dinar|Rupee|Anna|Cash|'
            r'Riksdaler|Skilling|Ore|Krone|Kreuzer|Heller|Batz|Rappen|Schilling|'
            r'Stuiver|Guilder|Albertin|Patagon|'
            r'Bronze|Lead|Billon|Electrum|Nickel|Brass|'
            r'Medal|Medallion|Pattern|Restrike|Uniface|Bracteate|'
            r'Milled|Cast|Struck|Mint|Rare|Unique|'
            r'Obverse|Reverse|Type|Variety)\b',
            re.IGNORECASE,
        )

        lots        = []
        lot_counter = 1

        for page_idx, page in enumerate(getattr(document, 'pages', [])):
            pdf_page   = page_idx + 1
            page_text  = self._extract_single_page_text(document, page)
            page_gaps  = image_gaps.get(pdf_page, [])
            para_ytops = self._get_page_para_ytops(page)

            if not page_text.strip():
                continue

            if page_idx < 2:
                print(f"  [HANDWRITTEN PIPELINE] Page {pdf_page} text (first 400 chars):")
                print(repr(page_text[:400]))

            year_positions = [(m.start(), m.group(1)) for m in year_split.finditer(page_text)]

            if not year_positions:
                print(f"  [HANDWRITTEN PIPELINE] Page {pdf_page}: no year anchors found, skipping.")
                continue

            print(f"  [HANDWRITTEN PIPELINE] Page {pdf_page}: {len(year_positions)} year anchors found.")

            for i, (pos, _year) in enumerate(year_positions):
                end_pos     = year_positions[i + 1][0] if i + 1 < len(year_positions) else len(page_text)
                description = page_text[pos:end_pos].strip()
                description = re.sub(r'\s*\n\s*', ' ', description)

                if len(description) < 10:
                    continue
                if re.match(r'^[A-Z\s\.\-]{20,}$', description[:80]):
                    continue
                if not coin_kw.search(description):
                    continue

                lot_data = self._parse_lot_fields(str(lot_counter), description, bold_ranges, 0)
                lot_data['Lot_No']         = lot_counter
                lot_data['Catalog_Source'] = catalog_name
                lot_data['Page_PDF']       = pdf_page
                lot_data['Image']          = 'Yes' if self._has_image_above_page(
                    pos, page_text, para_ytops, page_gaps) else None

                # Attempt to recover handwritten sale price from stray OCR tokens
                prefix_zone = page_text[max(0, pos - 20): pos + 30]
                pm = re.search(r'\b(\d{1,4}(?:\.\d{1,2})?)\b', prefix_zone)
                if pm and not lot_data.get('Sale_Price'):
                    try:
                        val = float(pm.group(1))
                        if 0.25 <= val <= 50_000 and not (1600 <= val <= 2099):
                            lot_data['Sale_Price'] = f"${pm.group(1)}"
                    except ValueError:
                        pass

                lots.append(lot_data)
                lot_counter += 1

        print(f"  [HANDWRITTEN PIPELINE] Extracted {len(lots)} lots (1 – {lot_counter - 1}).")
        return lots

    # =======================================================================
    # TEXT EXTRACTION — per-page with per-page column detection
    # =======================================================================

    def _extract_single_page_text(self, document: documentai.Document, page) -> str:
        """Extract text for a single page in correct visual reading order."""
        paras = getattr(page, 'paragraphs', None)
        if not paras:
            lines = getattr(page, 'lines', None)
            if lines:
                return '\n'.join(t for t in (self._layout_text(document, ln) for ln in lines) if t)
            tokens = getattr(page, 'tokens', [])
            return ' '.join(t for t in (self._layout_text(document, tk) for tk in tokens) if t)

        items = []
        for para in paras:
            bp = getattr(getattr(para, 'layout', None), 'bounding_poly', None)
            t  = self._layout_text(document, para)
            if not t:
                continue
            if bp and bp.vertices and len(bp.vertices) >= 4:
                v  = bp.vertices
                xc = (v[0].x + v[1].x) / 2
                yc = (v[0].y + v[2].y) / 2
            else:
                xc = yc = 0
            items.append({'text': t, 'x': xc, 'y': yc})

        if not items:
            return ''

        if self._is_page_two_column(items):
            xs    = sorted(p['x'] for p in items)
            split = xs[len(xs) // 2]
            left  = sorted([p for p in items if p['x'] <  split], key=lambda p: p['y'])
            right = sorted([p for p in items if p['x'] >= split], key=lambda p: p['y'])
            return '\n'.join(p['text'] for p in left + right)
        else:
            items.sort(key=lambda p: p['y'])
            return '\n'.join(p['text'] for p in items)

    def _is_page_two_column(self, items: List[Dict]) -> bool:
        if len(items) < 6:
            return False
        xs     = sorted(p['x'] for p in items)
        median = xs[len(xs) // 2]
        left   = sum(1 for p in items if p['x'] < median * 0.9)
        right  = sum(1 for p in items if p['x'] > median * 1.1)
        return left >= TWO_COL_MIN_COUNT and right >= TWO_COL_MIN_COUNT

    def _build_ordered_text(
        self,
        document: documentai.Document,
    ) -> Tuple[str, List[Tuple[int, int]]]:
        """
        Build visually-ordered full-document text with per-page column detection.
        Returns (full_text, page_markers) where page_markers is a list of
        (char_offset, pdf_page) noting where each page starts in full_text.
        """
        full_text    = ''
        page_markers: List[Tuple[int, int]] = []

        for page_idx, page in enumerate(getattr(document, 'pages', [])):
            page_markers.append((len(full_text), page_idx + 1))
            page_text = self._extract_single_page_text(document, page)
            if page_text:
                full_text += page_text + '\n'

        return full_text, page_markers

    def _page_from_markers(
        self,
        char_offset: int,
        page_markers: List[Tuple[int, int]],
    ) -> Optional[int]:
        if not page_markers:
            return None
        offsets = [m[0] for m in page_markers]
        idx     = bisect.bisect_right(offsets, char_offset) - 1
        return page_markers[max(idx, 0)][1]

    # =======================================================================
    # SPATIAL INDEX  (used by printed pipeline)
    # =======================================================================

    def _build_spatial_index(self, document: documentai.Document) -> List[Dict]:
        index: List[Dict] = []
        for page_idx, page in enumerate(getattr(document, 'pages', [])):
            pdf_page = page_idx + 1
            dim      = getattr(page, 'dimension', None)
            page_h   = getattr(dim, 'height', None) or 1.0

            for para in getattr(page, 'paragraphs', []):
                layout = getattr(para, 'layout', None)
                if not layout:
                    continue
                anchor = getattr(layout, 'text_anchor', None)
                if not anchor or not anchor.text_segments:
                    continue
                segs      = anchor.text_segments
                doc_start = min(int(getattr(s, 'start_index', 0)) for s in segs)
                doc_end   = max(int(getattr(s, 'end_index',   0)) for s in segs)

                bp    = getattr(layout, 'bounding_poly', None)
                y_top = y_bot = None
                if bp and bp.vertices and len(bp.vertices) >= 4:
                    v     = bp.vertices
                    y_top = min(v[0].y, v[1].y) / page_h
                    y_bot = max(v[2].y, v[3].y) / page_h

                index.append({
                    'pdf_page':  pdf_page,
                    'doc_start': doc_start,
                    'doc_end':   doc_end,
                    'y_top':     y_top,
                    'y_bot':     y_bot,
                })

        index.sort(key=lambda x: x['doc_start'])
        return index

    # =======================================================================
    # IMAGE DETECTION
    # =======================================================================

    def _compute_image_gaps(
        self,
        document: documentai.Document,
    ) -> Dict[int, List[Tuple[float, float]]]:
        """Per-page vertical gap detection for coin image regions."""
        gaps: Dict[int, List[Tuple[float, float]]] = {}

        for page_idx, page in enumerate(getattr(document, 'pages', [])):
            pdf_page = page_idx + 1
            dim      = getattr(page, 'dimension', None)
            page_h   = getattr(dim, 'height', None) or 1.0

            boxes: List[Tuple[float, float]] = []
            for para in getattr(page, 'paragraphs', []):
                bp = getattr(getattr(para, 'layout', None), 'bounding_poly', None)
                if bp and bp.vertices and len(bp.vertices) >= 4:
                    v = bp.vertices
                    boxes.append((min(v[0].y, v[1].y) / page_h,
                                  max(v[2].y, v[3].y) / page_h))

            if len(boxes) < 2:
                continue
            boxes.sort()

            page_gaps = [
                (boxes[i-1][1], boxes[i][0])
                for i in range(1, len(boxes))
                if boxes[i][0] - boxes[i-1][1] > IMAGE_GAP_THRESHOLD
            ]
            if page_gaps:
                gaps[pdf_page] = page_gaps

        return gaps

    def _has_image_above_ordered(
        self,
        char_offset: int,
        page_markers: List[Tuple[int, int]],
        image_gaps: Dict[int, List[Tuple[float, float]]],
    ) -> bool:
        """Image detection for the printed pipeline (ordered full_text)."""
        if not page_markers or not image_gaps:
            return False
        pdf_page = self._page_from_markers(char_offset, page_markers)
        if pdf_page is None:
            return False

        offsets    = [m[0] for m in page_markers]
        idx        = bisect.bisect_right(offsets, char_offset) - 1
        if idx < 0:
            return False
        page_start = page_markers[idx][0]
        page_end   = page_markers[idx + 1][0] if idx + 1 < len(page_markers) else page_start + 1
        y_approx   = (char_offset - page_start) / max(page_end - page_start, 1)

        for gap_top, gap_bot in image_gaps.get(pdf_page, []):
            if gap_top < y_approx and (y_approx - gap_bot) <= IMAGE_PROXIMITY:
                return True
        return False

    def _get_page_para_ytops(self, page) -> List[float]:
        """Sorted normalised y_top values for all paragraphs on a page."""
        dim    = getattr(page, 'dimension', None)
        page_h = getattr(dim, 'height', None) or 1.0
        ytops  = []
        for para in getattr(page, 'paragraphs', []):
            bp = getattr(getattr(para, 'layout', None), 'bounding_poly', None)
            if bp and bp.vertices and len(bp.vertices) >= 4:
                v = bp.vertices
                ytops.append(min(v[0].y, v[1].y) / page_h)
        ytops.sort()
        return ytops

    def _has_image_above_page(
        self,
        block_start: int,
        page_text: str,
        para_ytops: List[float],
        page_image_gaps: List[Tuple[float, float]],
    ) -> bool:
        """Image detection for the handwritten pipeline (page-by-page)."""
        if not page_image_gaps:
            return False
        y_approx = block_start / max(len(page_text), 1)
        for gap_top, gap_bot in page_image_gaps:
            if gap_top < y_approx and (y_approx - gap_bot) <= IMAGE_PROXIMITY:
                return True
        return False

    # =======================================================================
    # BOLD TEXT
    # =======================================================================

    def _extract_bold_ranges(self, document: documentai.Document) -> List[Tuple[int, int]]:
        bold_ranges: List[Tuple[int, int]] = []
        try:
            for page in getattr(document, 'pages', []):
                for token in getattr(page, 'tokens', []):
                    is_bold = False  # hook for future font-weight support
                    if is_bold:
                        anchor = getattr(getattr(token, 'layout', None), 'text_anchor', None)
                        if anchor:
                            for seg in anchor.text_segments:
                                bold_ranges.append((
                                    int(getattr(seg, 'start_index', 0)),
                                    int(getattr(seg, 'end_index',   0)),
                                ))
        except Exception:
            pass
        if not bold_ranges:
            print("  No bold style info available (normal for OCR processor).")
        return bold_ranges

    # =======================================================================
    # LOT VALIDATION FILTER (printed pipeline)
    # =======================================================================

    def _is_valid_lot(self, description: str) -> bool:
        if len(description) < 10:
            return False
        if re.match(r'^(Fair|Good|Fine|Poor|Very|Rare)\.?\s*$', description, re.IGNORECASE):
            return False
        if re.search(
            r'\b(Session|Friday|Monday|Tuesday|Wednesday|Thursday|Saturday|Sunday)\b.*\b([ap]\.m\.)\b',
            description, re.IGNORECASE,
        ):
            return False
        if re.search(r"^(The|Heritage|Superior|Stack's|Bowers)", description) and len(description) < 100:
            return False
        has_year    = bool(re.search(r'\b(1[6-9]\d{2}|20\d{2})\b', description))
        has_coin_kw = bool(re.search(
            r'\b(Mass\.|Rosa|Cent|Dollar|Token|Half|Penny|Shilling|Crown|Bust|Eagle|Liberty|'
            r'Wood|Copper|Silver|Gold|Proof)\b', description, re.IGNORECASE,
        ))
        has_text    = len(re.findall(r'[A-Za-z]{4,}', description)) >= 2
        return has_year or has_coin_kw or has_text

    # =======================================================================
    # SHARED LOT-FIELD PARSING
    # =======================================================================

    def _parse_lot_fields(
        self,
        lot_no: str,
        description: str,
        bold_ranges: List[Tuple[int, int]],
        desc_start_pos: int = 0,
        existing: Optional[Dict] = None,
    ) -> Dict:
        lot_data: Dict = existing or {
            'Lot_No':            lot_no,
            'Page_PDF':          None,
            'Catalog_Section':   None,
            'Headline':          None,
            'Image':             None,
            'Short_Description': None,
            'Long_Description':  description,
            'Pedigree':          None,
            'Sale_Price':        None,
            'Sold_To':           None,
            'Year':              None,
            'Grade':             None,
            'Grading_Service':   None,
            'Variety':           None,
            'Rarity':            None,
        }
        lot_data['Long_Description'] = description

        # Year
        ym = re.search(r'\b(1[7-9]\d{2}|20[0-2]\d)\b', description)
        if ym:
            lot_data['Year'] = ym.group(1)

        # Grade
        for pat in [
            r'\b(Very Fine|Extremely Fine|About Uncirculated|Mint State|'
            r'Choice Uncirculated|Gem Uncirculated)\s*(\d{2,3})?\b',
            r'\b(MS|AU|XF|EF|VF|PR|Proof)\s*[-]?\s*(\d{2,3})\b',
            r'\b(Uncirculated|Choice|Gem|Superb|Fine|Good|Fair|Poor)\b',
        ]:
            gm = re.search(pat, description, re.IGNORECASE)
            if gm:
                lot_data['Grade'] = gm.group(0).strip()
                break

        # Grading service
        sm = re.search(r'\b(PCGS|NGC|ANACS|ICG|Hallmark)\b', description)
        if sm:
            lot_data['Grading_Service'] = sm.group(1)

        # Variety
        variety_patterns = [
            r'Newcomb[- .]?\d+',    r'Sheldon[- .]?\d+',
            r'Breen[- .]?\d+',      r'Br[- .]?\d+',
            r'Cohen[- .]?\d+',      r'C[- .]?\d+',
            r'Logan-McCloskey[- .]?\d+', r'LM[- .]?\d+',
            r'John Reich[- .]?\d+', r'JR[- .]?\d+',
            r'Browning[- .]\d+',    r'Haseltine[- .]\d+',
            r'Beistle[- .]\d+',     r'Overton[- .]\d+',
            r'Tompkins[- .]\d+',    r'Bolender[- .]\d+',
            r'Bowers-Borckardt[- .]\d+', r'BB[- .]\d+',
            r'Gilbert[- .]\d+',     r'G[- .]\d+',
            r'B[- ]\d+',            r'O[- ]\d+',
            r'H[- ]\d+',            r'T[- ]\d+',
        ]
        variety_special = [
            (r'(?<!M)(?<!P)V[- ]?\d+', 'V'),
            (r'(?<!P)R[- ]?\d+',       'R'),
            (r'(?<!M)S[- ]?\d+',       'S'),
        ]
        for pat in variety_patterns:
            vm = re.search(pat, description)
            if vm:
                lot_data['Variety'] = vm.group(0)
                break
        if not lot_data.get('Variety'):
            for pat, _ in variety_special:
                vm = re.search(pat, description)
                if vm:
                    txt = vm.group(0)
                    ctx = description[max(0, vm.start() - 5): vm.start() + len(txt) + 5]
                    if not re.search(r'\b(FR|FA|AG|VG|VF|XF|EF|AU|MS|PR)\s*' + re.escape(txt), ctx):
                        lot_data['Variety'] = txt
                        break

        # Rarity
        rm = re.search(r'Rarity[- ]?(\d+)', description, re.IGNORECASE)
        if rm:
            lot_data['Rarity'] = f"R-{rm.group(1)}"

        # Pedigree
        pm = re.search(
            r'(from|purchased at|ex\.?|originally)\s+([^.]+(?:collection|sale))',
            description, re.IGNORECASE,
        )
        if pm:
            lot_data['Pedigree'] = pm.group(0).strip()

        # Headline
        lot_data['Headline'] = self._extract_headline(description)

        # Short_Description from bold ranges
        short_desc = None
        if bold_ranges:
            desc_end = desc_start_pos + len(description)
            for bs, be in bold_ranges:
                if bs >= desc_start_pos and bs < desc_end:
                    rs  = max(0, bs - desc_start_pos)
                    re_ = min(len(description), be - desc_start_pos)
                    bt  = description[rs:re_].strip()
                    if rs < 50 and 5 < len(bt) < 200:
                        if not lot_data['Headline'] or bt.lower() != lot_data['Headline'].lower():
                            short_desc = bt
                            break
        lot_data['Short_Description'] = short_desc

        return lot_data

    def _extract_headline(self, description: str) -> Optional[str]:
        description = description.strip()
        lines = description.split('\n')
        if not lines:
            return None
        first = lines[0].strip()

        if len(first) >= 8 and re.match(r'^[A-Z0-9\s\-]{8,}$', first):
            if first not in {'ONE CENT', 'HALF CENT', 'UNITED STATES', 'DOLLARS'}:
                if not re.match(r'^\d{4}\s+[A-Z]{2,4}\s*\d{0,3}$', first):
                    return first

        if len(lines) >= 2:
            joined = (lines[0] + ' ' + lines[1]).strip()
            if len(joined) >= 8 and re.match(r'^[A-Z0-9\s\-]{8,}$', joined):
                if joined not in {'ONE CENT', 'HALF CENT', 'UNITED STATES'}:
                    if not re.match(r'^\d{4}\s+[A-Z]{2,4}\s*\d{0,3}', joined):
                        return joined

        keywords = ['FIERY', 'SUPERB', 'CHOICE', 'GEM', 'RARE', 'EXTREMELY RARE',
                    'VERY RARE', 'IMPORTANT', 'BEAUTIFUL', 'SPECTACULAR']
        for kw in keywords:
            if first.upper().startswith(kw) and 8 <= len(first) <= 100:
                return first

        return None

    @staticmethod
    def _layout_text(document: documentai.Document, element) -> str:
        try:
            segs = element.layout.text_anchor.text_segments
            return ''.join(
                document.text[
                    int(getattr(s, 'start_index', 0)):
                    int(getattr(s, 'end_index',   0))
                ]
                for s in segs
            )
        except Exception:
            return ''
