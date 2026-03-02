"""
Auction Catalog Parser — Configuration
=======================================

Fill in your Google Cloud credentials and PDF paths, then run:

    python config.py

or import from another script:

    from config import CATALOGS, OUTPUT_EXCEL, PROJECT_ID, LOCATION, PROCESSOR_ID
    from auction_parser import AuctionCatalogParser
    parser = AuctionCatalogParser(PROJECT_ID, LOCATION, PROCESSOR_ID)
    df = parser.process_catalog_batch(CATALOGS, output_excel=OUTPUT_EXCEL)

CATALOG FORMAT
--------------
Each entry is a tuple:  (pdf_path, has_handwriting)

    has_handwriting = False  →  PRINTED pipeline
                                Lot numbers are typeset; OCR reads them directly.
                                Covers: no-annotation catalogs, image-interspersed
                                catalogs, complex-layout catalogs.

    has_handwriting = True   →  HANDWRITTEN pipeline
                                Lot numbers are written in the margin; OCR is
                                unreliable for them.  Lots are instead found by
                                locating lines that begin with a printed year and
                                are assigned sequential numbers (1, 2, 3 …).
                                Handwritten prices that OCR happens to capture are
                                also recovered.
                                Covers: Norris (1894), Galpin (1883).
"""

import os
import sys

# ---------------------------------------------------------------------------
# Google Cloud credentials
# ---------------------------------------------------------------------------

# Google Cloud credentials
os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = ('/Users/garrettziss/Documents/3. ECON 2824 Ferrara/pivotal-stacker-426622-c1-474370a8b91d.json')

# Google Document AI processor details
PROJECT_ID = 'pivotal-stacker-426622-c1'
LOCATION = 'us'
PROCESSOR_ID = 'e4894db7abb03eb1'

sys.path.append('/Users/garrettziss/Documents/3. ECON 2824 Ferrara/')

# ---------------------------------------------------------------------------
# Catalogs to process
# (pdf_path, has_handwriting)
# ---------------------------------------------------------------------------

CATALOGS = [
    # ── PRINTED lot numbers (no handwriting) ────────────────────────────────
    ('/Users/garrettziss/Documents/3. ECON 2824 Ferrara/1. No Annotations (Bidwell and Cottier, 1885).pdf',   False),
    ('/Users/garrettziss/Documents/3. ECON 2824 Ferrara/4. Interspersed Images.pdf',                           False),
    ('/Users/garrettziss/Documents/3. ECON 2824 Ferrara/5a. Complex Layouts Superior.pdf',                     False),
    ('/Users/garrettziss/Documents/3. ECON 2824 Ferrara/5b. Complex Layouts Heritagex.pdf',                    False),

    # ── HANDWRITTEN lot numbers ──────────────────────────────────────────────
    ('/Users/garrettziss/Documents/3. ECON 2824 Ferrara/2. Priced, Not Named (Norris, 1894).pdf',              True),
    ('/Users/garrettziss/Documents/3. ECON 2824 Ferrara/3. Named and Priced (Galpin Foreign Coins and Medals, 5.1883).pdf', True),
]

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

OUTPUT_EXCEL = '/Users/garrettziss/Downloads/auction_catalog_parser_results.xlsx'


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def test_connection() -> bool:
    """Quick sanity-check that Google Cloud credentials work."""
    try:
        from google.cloud import documentai_v1 as documentai
        documentai.DocumentProcessorServiceClient()
        print(f"✓ Authentication successful  (project: {PROJECT_ID})")
        return True
    except Exception as exc:
        print(f"✗ Authentication failed: {exc}")
        print()
        print("Make sure you have:")
        print("  1. Downloaded a JSON service-account key file")
        print("  2. Set GOOGLE_APPLICATION_CREDENTIALS to its path above")
        print("  3. Enabled the Document AI API in your project")
        return False


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    print('=' * 70)
    print('AUCTION CATALOG PARSER')
    print('=' * 70)

    # 1. Credentials check
    print('\n[1] Testing Google Cloud connection…')
    if not test_connection():
        print('\n✗ Fix the authentication issues above before continuing.')
        sys.exit(1)

    # 2. Load parser
    print('\n[2] Loading parser…')
    from combined_auction_parser import AuctionCatalogParser

    # 3. Initialise
    print('\n[3] Initialising Document AI client…')
    parser = AuctionCatalogParser(
        project_id   = PROJECT_ID,
        location     = LOCATION,
        processor_id = PROCESSOR_ID,
    )

    # 4. Process
    printed_count     = sum(1 for _, hw in CATALOGS if not hw)
    handwritten_count = sum(1 for _, hw in CATALOGS if hw)
    print(f'\n[4] Processing {len(CATALOGS)} catalog(s):')
    print(f'    • {printed_count} with PRINTED lot numbers   → OCR-based pipeline')
    print(f'    • {handwritten_count} with HANDWRITTEN lot numbers → year-based sequential pipeline')
    print('    This may take several minutes…')

    try:
        df = parser.process_catalog_batch(CATALOGS, output_excel=OUTPUT_EXCEL)

        print()
        print('=' * 70)
        print('✓ SUCCESS')
        print('=' * 70)
        print(f'✓ Extracted {len(df)} lots in total')
        print(f'✓ Results saved to: {OUTPUT_EXCEL}')
        print()
        print('Preview (first 10 rows):')
        preview_cols = ['Catalog_Source', 'Lot_No', 'Year', 'Grade', 'Long_Description']
        print(df[[c for c in preview_cols if c in df.columns]].head(10).to_string(index=False))

    except Exception as exc:
        print()
        print('=' * 70)
        print('✗ ERROR')
        print('=' * 70)
        print(f'Error: {exc}')
        print()
        print('Check:')
        print('  1. All PDF paths in CATALOGS are correct')
        print('  2. PROCESSOR_ID is correct')
        print('  3. The Document AI API is enabled in your project')
        raise
