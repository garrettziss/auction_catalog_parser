"""
Auction Catalog Parser Configuration
Fill in your values and run this file in Spyder to process your PDFs
"""

import os
import sys

# ============================================================================
# GOOGLE STUFF
# ============================================================================

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = '/path/pivotal-stacker-XXXXXX-c1-XXXXXXXXXXXX.json'

PROJECT_ID = "pivotal-stacker-XXXXXX-c1"
LOCATION = "us"

PROCESSOR_ID = "XXXXXXXXXXXXXXXX"

# ============================================================================
# PDFs TO PARSE
# ============================================================================

PDF_FILES = [
    "/path/1. No Annotations (Bidwell and Cottier, 1885).pdf",
    "/path/2. Priced, Not Named (Norris, 1894).pdf",
    "/path/3. Named and Priced (Galpin Foreign Coins and Medals, 5.1883).pdf",
    "/path/4. Interspersed Images.pdf",
    "/path/5a. Complex Layouts Superior.pdf",
    "/path/5b. Complex Layouts Heritagex.pdf",
]

# ============================================================================
# OUTPUT FILE
# ============================================================================

OUTPUT_EXCEL = "auction_catalog_parser_results.xlsx"

# ============================================================================
# Test if credentials are working
# ============================================================================

def test_connection():
    try:
        from google.cloud import documentai_v1 as documentai
        client = documentai.DocumentProcessorServiceClient()
        print("✓ Authentication successful!")
        print(f"✓ Project: {PROJECT_ID}")
        return True
    except Exception as e:
        print(f"✗ Authentication failed: {e}")
        print("\nMake sure you:")
        print("1. Downloaded the JSON key file")
        print("2. Updated the GOOGLE_APPLICATION_CREDENTIALS path above")
        print("3. Enabled the Document AI API")
        return False

# ============================================================================
# RUN THE PARSER
# ============================================================================

if __name__ == "__main__":
    print("=" * 80)
    print("AUCTION CATALOG PARSER")
    print("=" * 80)
    
    # Test connection first
    print("\n1. Testing connection...")
    if not test_connection():
        print("\n✗ Fix authentication issues above before continuing")
        sys.exit(1)
    
    # Import the parser
    print("\n2. Loading parser...")
    from auction_parser import AuctionCatalogParser
    
    # Create parser instance
    print("\n3. Initializing Document AI...")
    parser = AuctionCatalogParser(
        project_id=PROJECT_ID,
        location=LOCATION,
        processor_id=PROCESSOR_ID
    )
    
    # Process the PDFs
    print(f"\n4. Processing {len(PDF_FILES)} PDF files...")
    print("   This may take a few minutes...")
    
    try:
        df = parser.process_catalog_batch(PDF_FILES, output_excel=OUTPUT_EXCEL)
        
        print("\n" + "=" * 80)
        print("✓ SUCCESS!")
        print("=" * 80)
        print(f"✓ Processed {len(df)} lots")
        print(f"✓ Results saved to: {OUTPUT_EXCEL}")
        print("\nSummary:")
        print(df[['Lot_No', 'Year', 'Grade', 'Short_Description']].head(10))
        
    except Exception as e:
        print("\n" + "=" * 80)
        print("✗ ERROR")
        print("=" * 80)
        print(f"Error: {e}")
        print("\nCheck:")
        print("1. PDF file paths are correct")
        print("2. Processor ID is correct")
        print("3. Document AI API is enabled")
