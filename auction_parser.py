import os
import re
from typing import List, Dict, Optional
from pathlib import Path

from google.cloud import documentai_v1 as documentai
import pandas as pd


class AuctionCatalogParser:
    
    def __init__(self, project_id: str, location: str, processor_id: str):
        # Initializer
        self.project_id = project_id
        self.location = location
        self.processor_id = processor_id
        
        # Initialize Document AI client
        self.client = documentai.DocumentProcessorServiceClient()
        
        # Construct processor name
        self.processor_name = self.client.processor_path(
            project_id, location, processor_id
        )
    
    def process_document(self, file_path: str) -> documentai.Document:
        
        # Process a single PDF via OCR Processor; return Document AI Document object
        # Read the file
        with open(file_path, "rb") as document:
            document_content = document.read()
        
        # Create the request
        request = documentai.ProcessRequest(
            name=self.processor_name,
            raw_document=documentai.RawDocument(
                content=document_content,
                mime_type="application/pdf"
            ),
            skip_human_review=True
        )
        
        # Process the document
        print(f"Processing: {file_path}")
        result = self.client.process_document(request=request)
        
        return result.document
    
    # Extract structured lot data from processed document and returns dictionaries w/ lot info
    def extract_lot_data(self, document: documentai.Document, catalog_name: str) -> List[Dict]:
        
        # Detect if document has two-column layout
        is_two_column = self._detect_column_layout(document)
        
        # Extract text using appropriate method
        if is_two_column:
            print("Using column-aware text extraction...")
            full_text = self._extract_text_by_columns(document)
        else:
            print("Using standard text extraction...")
            full_text = self._extract_text_from_document(document)
        
        if not full_text:
            print(f"ERROR: Could not extract any text from {catalog_name}")
            return []
        
        print(f"Extracted {len(full_text)} characters of text")
        
        # Try to extract bold text information
        bold_text_info = self._extract_bold_text_ranges(document)
        
        # Process the full text to extract lots
        lots = self._parse_full_text(full_text, catalog_name, bold_text_info)
        
        return lots
    
    # Extract text from Document AI response and return full text from all pages
    def _extract_text_from_document(self, document: documentai.Document) -> str:
        full_text = ""
        
        # Try document.text first (might be populated)
        if hasattr(document, 'text') and document.text:
            return document.text
        
        # Extract from pages
        if not hasattr(document, 'pages') or not document.pages:
            return ""
        
        # Try to extract from paragraphs (best structure)
        for page in document.pages:
            if hasattr(page, 'paragraphs') and page.paragraphs:
                for para in page.paragraphs:
                    if hasattr(para, 'layout') and hasattr(para.layout, 'text_anchor'):
                        # Get text from text_anchor
                        text_anchor = para.layout.text_anchor
                        if hasattr(text_anchor, 'text_segments'):
                            for segment in text_anchor.text_segments:
                                start = int(segment.start_index) if hasattr(segment, 'start_index') else 0
                                end = int(segment.end_index) if hasattr(segment, 'end_index') else 0
                                if hasattr(document, 'text'):
                                    full_text += document.text[start:end] + "\n"
        
        # If paragraphs didn't work, try lines
        if not full_text:
            for page in document.pages:
                if hasattr(page, 'lines') and page.lines:
                    for line in page.lines:
                        if hasattr(line, 'layout') and hasattr(line.layout, 'text_anchor'):
                            text_anchor = line.layout.text_anchor
                            if hasattr(text_anchor, 'text_segments'):
                                for segment in text_anchor.text_segments:
                                    start = int(segment.start_index) if hasattr(segment, 'start_index') else 0
                                    end = int(segment.end_index) if hasattr(segment, 'end_index') else 0
                                    if hasattr(document, 'text'):
                                        full_text += document.text[start:end] + "\n"
        
        # If still nothing, try tokens as last resort
        if not full_text:
            for page in document.pages:
                if hasattr(page, 'tokens') and page.tokens:
                    for token in page.tokens:
                        if hasattr(token, 'layout') and hasattr(token.layout, 'text_anchor'):
                            text_anchor = token.layout.text_anchor
                            if hasattr(text_anchor, 'text_segments'):
                                for segment in text_anchor.text_segments:
                                    start = int(segment.start_index) if hasattr(segment, 'start_index') else 0
                                    end = int(segment.end_index) if hasattr(segment, 'end_index') else 0
                                    if hasattr(document, 'text'):
                                        full_text += document.text[start:end] + " "
        
        return full_text.strip()
    
    # Detect if document has two-column layout
    def _detect_column_layout(self, document: documentai.Document) -> bool:
    
        try:
            if not hasattr(document, 'pages') or not document.pages:
                return False
            
            # Sample first page to detect layout
            page = document.pages[0]
            
            if not hasattr(page, 'paragraphs') or not page.paragraphs:
                return False
            
            # Collect X-coordinates of paragraph centers
            x_positions = []
            
            for paragraph in page.paragraphs[:20]:  # Sample first 20 paragraphs
                if hasattr(paragraph, 'layout') and hasattr(paragraph.layout, 'bounding_poly'):
                    vertices = paragraph.layout.bounding_poly.vertices
                    if len(vertices) >= 2:
                        # Calculate center X position
                        x_center = (vertices[0].x + vertices[1].x) / 2
                        x_positions.append(x_center)
            
            if len(x_positions) < 10:
                return False  # Not enough data
            
            # Check if X positions cluster into two distinct groups
            # Sort positions
            x_sorted = sorted(x_positions)
            
            # Calculate median
            median_x = x_sorted[len(x_sorted) // 2]
            
            # Count how many are left vs right of median
            left_count = sum(1 for x in x_positions if x < median_x * 0.9)
            right_count = sum(1 for x in x_positions if x > median_x * 1.1)
            
            # If we have significant paragraphs on both sides, it's two-column
            is_two_column = left_count >= 3 and right_count >= 3
            
            if is_two_column:
                print(f"Two-column layout detected (left: {left_count}, right: {right_count})")
            else:
                print(f"Single-column layout detected")
            
            return is_two_column
            
        except Exception as e:
            print(f"Column detection failed: {e}")
            return False  # Default to single column
    
    # If there's a two column layout, read left and then right
    def _extract_text_by_columns(self, document: documentai.Document) -> str:
        try:
            if not hasattr(document, 'pages') or not document.pages:
                return ""
            
            full_text = ""
            
            for page in document.pages:
                if not hasattr(page, 'paragraphs') or not page.paragraphs:
                    continue
                
                # Collect paragraphs with their positions
                paragraphs_data = []
                
                for paragraph in page.paragraphs:
                    if not hasattr(paragraph, 'layout') or not hasattr(paragraph.layout, 'text_anchor'):
                        continue
                    
                    text_anchor = paragraph.layout.text_anchor
                    
                    # Get bounding box
                    if hasattr(paragraph.layout, 'bounding_poly') and paragraph.layout.bounding_poly.vertices:
                        vertices = paragraph.layout.bounding_poly.vertices
                        
                        # Calculate center position
                        x_center = (vertices[0].x + vertices[1].x) / 2
                        y_center = (vertices[0].y + vertices[2].y) / 2
                        
                        # Extract text
                        para_text = ""
                        for segment in text_anchor.text_segments:
                            start = int(segment.start_index) if hasattr(segment, 'start_index') else 0
                            end = int(segment.end_index) if hasattr(segment, 'end_index') else 0
                            para_text += document.text[start:end]
                        
                        paragraphs_data.append({
                            'text': para_text,
                            'x': x_center,
                            'y': y_center
                        })
                
                # Find column boundary (median X position)
                if paragraphs_data:
                    x_positions = [p['x'] for p in paragraphs_data]
                    column_boundary = sorted(x_positions)[len(x_positions) // 2]
                    
                    # Separate into left and right columns
                    left_column = [p for p in paragraphs_data if p['x'] < column_boundary]
                    right_column = [p for p in paragraphs_data if p['x'] >= column_boundary]
                    
                    # Sort each column by Y position (top to bottom)
                    left_column.sort(key=lambda p: p['y'])
                    right_column.sort(key=lambda p: p['y'])
                    
                    # Add left column text
                    for para in left_column:
                        full_text += para['text'] + "\n"
                    
                    # Add right column text
                    for para in right_column:
                        full_text += para['text'] + "\n"
            
            return full_text.strip()
            
        except Exception as e:
            print(f"Column-aware extraction failed: {e}")
            # Fallback to regular extraction
            return self._extract_text_from_document(document)
    
    # Find bold text (for Short_Description)
    def _extract_bold_text_ranges(self, document: documentai.Document) -> List[tuple]:
        bold_ranges = []
        
        try:
            if not hasattr(document, 'pages') or not document.pages:
                return bold_ranges
            
            # Check if Document AI provides text style information
            for page in document.pages:
                # Try to get style information from tokens
                if hasattr(page, 'tokens') and page.tokens:
                    for token in page.tokens:
                        # Check if token has text_anchor with style info
                        if hasattr(token, 'layout') and hasattr(token.layout, 'text_anchor'):
                            text_anchor = token.layout.text_anchor
                            
                            # Check for font/style information
                            # Different processors may store this differently
                            is_bold = False
                            
                            # Check if there's a detected_break with font weight
                            if hasattr(token, 'detected_break'):
                                break_info = token.detected_break
                                if hasattr(break_info, 'type'):
                                    # Some processors indicate bold via break type
                                    pass
                            
                            # Check layout for text style
                            if hasattr(token.layout, 'confidence'):
                                # Font weight might be in layout properties
                                layout = token.layout
                                if hasattr(layout, 'bounding_poly'):
                                    # Check vertices or properties
                                    pass
                            
                            # If we determine text is bold, add its range
                            if is_bold and hasattr(text_anchor, 'text_segments'):
                                for segment in text_anchor.text_segments:
                                    start = int(segment.start_index) if hasattr(segment, 'start_index') else 0
                                    end = int(segment.end_index) if hasattr(segment, 'end_index') else 0
                                    bold_ranges.append((start, end))
        
        except Exception as e:
            # If bold extraction fails, just continue without it
            pass
        
        if bold_ranges:
            print(f"Found {len(bold_ranges)} bold text ranges")
        else:
            print("No bold text style information available from Document AI")
        
        return bold_ranges
    
    # Parse document for all lots
    def _parse_full_text(self, text: str, catalog_name: str, bold_text_info: List[tuple] = None) -> List[Dict]:
        lots = []
        
        # Show first 500 chars for debugging
        print(f"\nFirst 500 chars of text:")
        print(repr(text[:500]))
        
        # Pattern to find lot numbers at start of line
        # Accept 1-4 digits (some catalogs use small lot numbers like 2, 3, 15)
        # Followed by space and description OR on its own line
        lot_pattern = re.compile(r'^\s*(\d{1,4})(?:\s+|$)', re.MULTILINE)
        
        # Find all potential lot numbers
        lot_matches = list(lot_pattern.finditer(text))
        print(f"Found {len(lot_matches)} potential lot numbers (1-4 digits)")
        
        if lot_matches:
            print(f"First 10 numbers: {[m.group(1) for m in lot_matches[:10]]}")
        
        # Extract each lot's full description
        filtered_lots = []
        
        for i, match in enumerate(lot_matches):
            lot_no = match.group(1)
            
            # Get description from after lot number to before next lot number
            start_pos = match.end()
            
            if i + 1 < len(lot_matches):
                end_pos = lot_matches[i + 1].start()
            else:
                end_pos = len(text)
            
            description = text[start_pos:end_pos].strip()
            
            # Remove line breaks from description (replace with spaces)
            # This makes multi-line descriptions continuous
            description = re.sub(r'\s*\n\s*', ' ', description)
            
            # Skip if description is too short (likely page number or price)
            if len(description) < 10:
                continue
            
            # Skip if description is JUST a grade word with no other content
            if re.match(r'^(Fair|Good|Fine|Poor|Very|Rare)\.?\s*$', description, re.IGNORECASE):
                continue
            
            # FILTER OUT HEADERS/FOOTERS
            # Session headers like "Session 2 Friday, September 20 • 7:00 p.m."
            if re.search(r'\b(Session|Friday|Monday|Tuesday|Wednesday|Thursday|Saturday|Sunday)\b.*\b(p\.m\.|a\.m\.)\b', description, re.IGNORECASE):
                continue
            
            # Page headers/footers with catalog names
            if re.search(r'^(The|Heritage|Superior|Stack\'s|Bowers)', description) and len(description) < 100:
                continue
            
            # FILTER: Lot must have substantive content
            # Real lots should have either:
            # - A year (1600-2099) 
            # - A coin keyword (Mass., Rosa, Cent, Dollar, Token, etc.)
            # - Substantive text (2+ words of 4+ letters)
            # This filters out page numbers, prices, and random numbers
            
            has_year = bool(re.search(r'\b(1[6-9]\d{2}|20\d{2})\b', description))
            has_coin_keyword = bool(re.search(
                r'\b(Mass\.|Rosa|Cent|Dollar|Token|Half|Penny|Shilling|Crown|Bust|Eagle|Liberty|Wood|Copper|Silver|Gold|Proof)\b',
                description, re.IGNORECASE
            ))
            has_substantive_text = len(re.findall(r'[A-Za-z]{4,}', description)) >= 2
            
            if not (has_year or has_coin_keyword or has_substantive_text):
                continue
            
            # Store lot position for price extraction
            filtered_lots.append((lot_no, description, start_pos, match.start()))
        
        print(f"After filtering: {len(filtered_lots)} lots")
        if filtered_lots:
            print(f"First 5 lot numbers: {[lot[0] for lot in filtered_lots[:5]]}")
        
        # Parse the filtered lots
        for lot_no, description, desc_start, lot_start in filtered_lots:
            lot_data = self._parse_lot_description(lot_no, description, bold_text_info, desc_start)
            
            # Try to extract sale price from text BEFORE the lot number
            # Format: handwritten price like "1.20", "2.5", ".50" before lot number
            price_search_start = max(0, lot_start - 50)  # Look back up to 50 chars
            price_context = text[price_search_start:lot_start]
            
            # Look for price pattern: optional dollar sign, digits, optional decimal
            price_match = re.search(r'\$?\s*(\d+\.?\d*)\s*$', price_context)
            if price_match:
                price_str = price_match.group(1)
                # Only accept if it looks like a reasonable price (not a lot number)
                try:
                    price_val = float(price_str)
                    if 0.01 <= price_val <= 10000:  # Reasonable price range
                        lot_data['Sale_Price'] = f"${price_str}"
                except:
                    pass
            
            lot_data['Catalog_Source'] = catalog_name
            lots.append(lot_data)
        
        return lots
    
    # Parse each lot into its fields
    def _parse_lot_description(self, lot_no: str, description: str, bold_text_info: List[tuple] = None, desc_start_pos: int = 0) -> Dict:
    
        lot_data = {
            'Lot_No': lot_no,
            'Page_Start': None,
            'Page_End': None,
            'Catalog_Section': None,
            'Headline': None,
            'Image_Link': None,
            'Short_Description': None,
            'Long_Description': description,
            'Pedigree': None,
            'Sale_Price': None,
            'Sold_To': None,
            'Year': None,
            'Grade': None,
            'Rarity': None,
            'Grading_Service': None,
            'Variety': None
        }
        
        # Extract year (4 digits, likely 1700-2099)
        year_match = re.search(r'\b(1[7-9]\d{2}|20[0-2]\d)\b', description)
        if year_match:
            lot_data['Year'] = year_match.group(1)
        
        # Extract grade - IMPROVED to capture multi-word grades and numbers
        grade_patterns = [
            # Two-word grades with optional numbers
            r'\b(Very Fine|Extremely Fine|About Uncirculated|Mint State|Choice Uncirculated|Gem Uncirculated)\s*(\d{2,3})?\b',
            # Single word with optional number/hyphen
            r'\b(MS|AU|XF|EF|VF|PR|Proof)\s*[-]?\s*(\d{2,3})\b',
            # Single word grades
            r'\b(Uncirculated|Choice|Gem|Superb|Fine|Good|Fair|Poor)\b',
        ]
        
        for pattern in grade_patterns:
            grade_match = re.search(pattern, description, re.IGNORECASE)
            if grade_match:
                # Get the full match (includes word + number if present)
                full_grade = grade_match.group(0).strip()
                lot_data['Grade'] = full_grade
                break
        
        # Extract grading service
        service_match = re.search(r'\b(PCGS|NGC|ANACS|ICG|Hallmark)\b', description)
        if service_match:
            lot_data['Grading_Service'] = service_match.group(1)
        
        # Extract die marriage patterns
        variety_patterns = [
            r'Breen[- ]?\d+',
            r'Br[- ]?\d+',
            r'Gilbert[- ]?\d+',
            r'G[- ]?\d+',
            r'Cohen[- ]?\d+',
            r'C[- ]?\d+',
            r'Sheldon[- ]?\d+',
            r'S[- ]?\d+',
            r'Newcomb[- ]?\d+',
            r'N[- ]?\d+',
            r'Logan-McCloskey[- ]?\d+',
            r'LM[- ]?\d+',
            r'John Reich[- ]?\d+',
            r'JR[- ]?\d+',
            r'Browning[- ]\d+',
            r'B[- ]\d+',
            r'Haseltine[- ]\d+',
            r'H[- ]\d+',
            r'Beistle[- ]\d+',
            r'Overton[- ]\d+',
            r'O[- ]\d+',
            r'Tompkins[- ]\d+',
            r'T[- ]\d+',
            r'Bolender[- ]\d+',
            r'Bowers-Borckardt[- ]\d+',
            r'BB[- ]\d+',
            r'Breen[. ]?\d+',
            r'Br[. ]?\d+',
            r'Cohen[. ]?\d+',
            r'C[. ]?\d+',
            r'Sheldon[. ]?\d+',
            r'S[. ]?\d+',
            r'Newcomb[. ]?\d+',
            r'N[. ]?\d+',
            r'Logan-McCloskey[. ]?\d+',
            r'LM[. ]?\d+',
            r'John Reich[. ]?\d+',
            r'JR[. ]?\d+',
            r'Browning[. ]\d+',
            r'B[. ]\d+',
            r'Haseltine[. ]\d+',
            r'H[. ]\d+',
            r'Beistle[. ]\d+',
            r'Overton[. ]\d+',
            r'O[. ]\d+',
            r'Tompkins[. ]\d+',
            r'T[. ]\d+',
            r'Bolender[. ]\d+',
            r'Bowers-Borckardt[. ]\d+',
            r'BB[. ]\d+',
        ]
        
        # Sometimes grades were counted as die marriages; this avoids that
        variety_special_patterns = [
            (r'(?<!M)(?<!P)V[- ]?\d+', 'V'),  # Not after M (MV) or P (PV)
            (r'(?<!P)R[- ]?\d+', 'R'),  # Not after P (PR 63 -> R 63)
            (r'(?<!M)S[- ]?\d+', 'S'),  # Not after M (MS 65 -> S 65)
        ]
        
        # Try standard variety patterns first
        for pattern in variety_patterns:
            variety_match = re.search(pattern, description)
            if variety_match:
                lot_data['Variety'] = variety_match.group(0)
                break
        
        # If no match, try special patterns with negative lookbehind
        if not lot_data['Variety']:
            for pattern, letter in variety_special_patterns:
                variety_match = re.search(pattern, description)
                if variety_match:
                    # Double-check it's not part of a grade
                    match_text = variety_match.group(0)
                    # Get context around match
                    match_start = variety_match.start()
                    context_start = max(0, match_start - 5)
                    context = description[context_start:match_start + len(match_text) + 5]
                    
                    # Skip if it's part of a known grade pattern
                    if not re.search(r'\b(FR|FA|AG|VG|VF|XF|EF|AU|MS|PR)\s*' + re.escape(match_text), context):
                        lot_data['Variety'] = match_text
                        break
        
        # Extract rarity
        rarity_match = re.search(r'Rarity[- ]?(\d+)', description, re.IGNORECASE)
        if rarity_match:
            lot_data['Rarity'] = f"R-{rarity_match.group(1)}"
        
        # Extract pedigree information
        pedigree_match = re.search(
            r'(from|purchased at|ex\.?|originally)\s+([^.]+(?:collection|sale))',
            description,
            re.IGNORECASE
        )
        if pedigree_match:
            lot_data['Pedigree'] = pedigree_match.group(0).strip()
        
        # Extract headline (all-caps or keyword-based)
        headline = self._extract_headline(description)
        lot_data['Headline'] = headline
        
        # Short_Description: Extract BOLD text at the start of description
        # If Document AI provided bold text ranges, use them
        # Otherwise, leave Short_Description as None
        short_desc = None
        
        if bold_text_info and len(bold_text_info) > 0:
            # Check if any bold text falls within this description
            desc_end_pos = desc_start_pos + len(description)
            
            for bold_start, bold_end in bold_text_info:
                # Check if this bold range overlaps with our description
                if bold_start >= desc_start_pos and bold_start < desc_end_pos:
                    # Extract the bold text (relative to description start)
                    rel_start = max(0, bold_start - desc_start_pos)
                    rel_end = min(len(description), bold_end - desc_start_pos)
                    
                    bold_text = description[rel_start:rel_end].strip()
                    
                    # Only use if it's at/near the start and not too long
                    if rel_start < 50 and len(bold_text) > 5 and len(bold_text) < 200:
                        # Make sure it's not the same as headline
                        if not headline or bold_text.lower() != headline.lower():
                            short_desc = bold_text
                            break
        
        # If no bold text detected, Short_Description is none and everything is included in Long_Description
        lot_data['Short_Description'] = short_desc
        
        return lot_data
    
    # Expensive lots have a headline, which is above the description (and there's almost always an image for such lots; its above that too)
    def _extract_headline(self, description: str) -> Optional[str]:
        # Remove leading/trailing whitespace
        description = description.strip()
        
        # Split into lines
        lines = description.split('\n')
        if not lines:
            return None
        
        # Pattern 1: All-caps headline on first line(s)
        # Check first line for all-caps (at least 8 characters to avoid "ONE CENT")
        first_line = lines[0].strip()
        
        # All-caps check: only uppercase letters, digits, spaces, hyphens
        if len(first_line) >= 8 and re.match(r'^[A-Z0-9\s\-]{8,}$', first_line):
            # Filter out generic denominations
            if first_line not in ['ONE CENT', 'HALF CENT', 'UNITED STATES', 'DOLLARS']:
                # Filter out things that are just years and grades
                if not re.match(r'^\d{4}\s+[A-Z]{2,4}\s*\d{0,3}$', first_line):
                    return first_line
        
        # Check if first TWO lines together form a headline
        if len(lines) >= 2:
            first_two = (lines[0] + ' ' + lines[1]).strip()
            if len(first_two) >= 8 and re.match(r'^[A-Z0-9\s\-]{8,}$', first_two):
                if first_two not in ['ONE CENT', 'HALF CENT', 'UNITED STATES']:
                    if not re.match(r'^\d{4}\s+[A-Z]{2,4}\s*\d{0,3}', first_two):
                        return first_two
        
        # Pattern 2: Keyword-based headlines (even if not all-caps)
        # These are special promotional text
        headline_keywords = [
            'FIERY', 'SUPERB', 'CHOICE', 'GEM',
            'RARE', 'EXTREMELY RARE', 'VERY RARE',
            'IMPORTANT', 'BEAUTIFUL', 'SPECTACULAR'
        ]
        
        first_line_upper = first_line.upper()
        for keyword in headline_keywords:
            if first_line_upper.startswith(keyword):
                # Take the entire first line as headline
                if len(first_line) >= 8 and len(first_line) <= 100:
                    return first_line
        
        return None
    
    def process_catalog_batch(self, pdf_files: List[str], output_excel: str = 'auction_catalog_parsed.xlsx') -> pd.DataFrame:
        
        # Process multiple catalog PDFs, combine results, and return DataFrame with all extracted lots
        all_lots = []
        
        for pdf_file in pdf_files:
            # Extract catalog name from filename
            catalog_name = Path(pdf_file).stem
            
            # Process the document
            document = self.process_document(pdf_file)
            
            # Extract lot data
            lots = self.extract_lot_data(document, catalog_name)
            all_lots.extend(lots)
            
            print(f"Extracted {len(lots)} lots from {catalog_name}\n")
        
        # Create DataFrame
        df = pd.DataFrame(all_lots)
        
        if len(df) == 0:
            print("\nWARNING: No lots were extracted!")
            return df
        
        # Reorder columns - Catalog_Source first, then Lot_No, then the rest
        column_order = [
            'Catalog_Source',
            'Lot_No',
            'Page_Start',
            'Page_End',
            'Catalog_Section',
            'Headline',
            'Image_Link',
            'Short_Description',
            'Long_Description',
            'Pedigree',
            'Sale_Price',
            'Sold_To',
            'Year',
            'Grade',
            'Grading_Service',
            'Variety',
            'Rarity'
        ]
        
        # Reorder columns (only include columns that exist)
        existing_cols = [col for col in column_order if col in df.columns]
        df = df[existing_cols]
        
        # Sort by catalog source, then lot number
        df = df.sort_values(['Catalog_Source', 'Lot_No']).reset_index(drop=True)
        
        # Save to Excel
        df.to_excel(output_excel, index=False)
        print(f"\nSaved {len(df)} total lots to {output_excel}")
        
        return df
