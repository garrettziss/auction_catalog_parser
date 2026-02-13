"""
Auction Catalog Parser Configuration
Fill in your values and run this file in Spyder to process your PDFs
"""

import os
import sys

# ============================================================================
# GOOGLE STUFF
# ============================================================================

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = '/Users/garrettziss/Documents/3. ECON 2824 Ferrara/pivotal-stacker-426622-c1-474370a8b91d.json'

PROJECT_ID = "pivotal-stacker-426622-c1"
LOCATION = "us"

PROCESSOR_ID = "e4894db7abb03eb1"

# ============================================================================
# PDFs TO PARSE
# ============================================================================

PDF_FILES = [
    "/Users/garrettziss/Documents/3. ECON 2824 Ferrara/1. No Annotations (Bidwell and Cottier, 1885).pdf",
    "/Users/garrettziss/Documents/3. ECON 2824 Ferrara/2. Priced, Not Named (Norris, 1894).pdf",
    "/Users/garrettziss/Documents/3. ECON 2824 Ferrara/3. Named and Priced (Galpin Foreign Coins and Medals, 5.1883).pdf",
    "/Users/garrettziss/Documents/3. ECON 2824 Ferrara/4. Interspersed Images.pdf",
    "/Users/garrettziss/Documents/3. ECON 2824 Ferrara/5a. Complex Layouts Superior.pdf",
    "/Users/garrettziss/Documents/3. ECON 2824 Ferrara/5b. Complex Layouts Heritagex.pdf",
]

# ============================================================================
# OUTPUT FILE
# ============================================================================

OUTPUT_EXCEL = "ferrara_auction_results.xlsx"

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
