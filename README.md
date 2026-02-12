# auction_catalog_parser

### Purpose

A significant number of numismatic auction catalogs have been scanned and uploaded to the internet, primarily on the Newman Numismatic Portal, but they are just in PDF form with OCR-based transcription for search purposes.  However, this transcription loses the tabular format of auction catalogs, and sometimes divorces the lot numbers from the lot descriptions, and always prevents easy searching for certain types of coins.  With the power of AI, a parser can be used to extract this information from the PDFs and arrange it in a searchable, tabular format.

### Types of Catalogs

To ensure broad applicability, many different types of catalogs from the 160 year history of American numismatics were tested:

##### 1. No Annotations
Example Catalog: The Bidwell Collection of Ancient, Foreign and American Coins and Medals, and The Cottier Collection of Cents and Half Cents (S.H. & H. Chapman, 6/1885)
Link to Original: https://archive.org/details/catalogueofcolle00chap_21/page/29/mode/1up

This is typical of early auction catalogs (Chapman catalogs, the premier numismatic auction house of the day, is used here). This just includes typed text, with lot numbers and item descriptions.

##### 2. Priced, Not Named
Example Catalog: Catalogue of the Collection of United States Coins of E. S. Norris, Esq., Boston, Massachusetts (S.H. & H. Chapman, 5/1894)
Link to Original: https://archive.org/details/catalogueofcolle00chap_17/page/3/mode/1up

The format is the same as the previous one, but it is "priced," meaning that an attendee of the sale wrote the prices realized for each lot into the catalog during the sale.  This is very useful to numismatists today for pedigree research.  It also adds significant complications for the parser, as it now has to differentiate the handwritten prices from the lot numbers, both of which are numbers, and has to do this when the prices are not always in the same location (though in this example, they are).

##### 3. Named and Priced
Example Catalog: Collection of Foreign Coins and Medals of Mr. A. Galpin (S.H. & H. Chapman, 5/1883)
Link to Original: https://archive.org/details/collectionoffore1883chap/page/13/mode/1up

The format is the same as the previous one, but it is "named and priced," meaning that an attendee of the sale wrote both the prices realized and the names of the purchasers into the catalog.  This significantly increases the value to numismatic researchers for pedigree purposes, but also significantly increases the difficulty for the parser, due to the 19th century handwriting and especially because of its rushed nature in a fast-paced auction.

##### 4. Interspersed Images
Example Catalog: The Amon G. Carter, Jr. Family Collection of United States Gold, Silver & Copper Coins and Foreign Coins (Stack's, 1/1984)
Link to Original: https://archive.org/details/amongcarterjrfam1984stac/amongcarterjrfam1984stac/page/112/mode/1up

The Chapman's format lasted until the 1930s or 1940s, and then gave way to this format, which lasted into the 1980s or 1990s, depending on the auction house. This format is similar to the previous, but due to advances in photographic technology and decreases in price, now include images directly above the description of that item, rather than in a dedicated photographic plate like the Chapmans had to do.  However, these photographs were not included for every lot, providing a challenge for the scraper.  The lots with images also include bolded elements, which represent a shortened version of the description to identify the item at a quick glance (Short_Description).  The most important lots also include a Headline atop of it, drawing attention to the lot, which can be multiple lines, bold, capitalized, or some combination of them.  The era of named and priced sales had passed by this point, due to changes in the hobby and the rise of mail bids and auction agents that obscured the bidder's identity, so this feature will not be included from here on out.

##### 5. Complex Layouts
Example Catalog #5a: The Jascha Heifetz Collection Sale (Superior, 10/1989)
Link to Original #5a: https://archive.org/details/jaschaheifetzcol1989supe/page/96/mode/1up
Example Catalog #5b: The Long Beach Sale (Heritage, 9/1996)
Link to Original #5b: https://archive.org/details/september1996lon1996heri/page/63/mode/1up

While some auction companies continued with the previous format for a long period of time, others had two-columns of lots, but with the same general format (including interspersed images), starting in the 1970s or 1980s. Short_Descriptions are nearly universal, and Headlines are provided for almost every imaged lot, with images becoming more and more common as time goes by.  Two files exist: one for Superior Auctions (5a), and one for Heritage Auctions (5b). The latter was included to see if the model made for the former was generalizable, and it was.

### Output

The model outputs an Excel file, auction_catalog_parser.xlsx. The file example_auction_catalog_parser contains a hand-typed version of the ideal.

### Limitations
