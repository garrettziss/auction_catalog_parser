# Auction Catalog Parser

## Purpose

A significant number of numismatic auction catalogs have been scanned and uploaded to the internet, primarily on the Newman Numismatic Portal, but they are just in PDF form with OCR-based transcription for search purposes.  However, this transcription loses the tabular format of auction catalogs, and sometimes divorces the lot numbers from the lot descriptions, and always prevents easy searching for certain types of coins.  With the power of AI, a parser can be used to extract this information from the PDFs and arrange it in a searchable, tabular format.

## Types of Catalogs

To ensure broad applicability, many different types of catalogs from the 160 year history of American numismatics were tested:

#### 1. No Annotations
Example Catalog: The Bidwell Collection of Ancient, Foreign and American Coins and Medals, and The Cottier Collection of Cents and Half Cents (S.H. & H. Chapman, 6/1885)  
Link to Original: https://archive.org/details/catalogueofcolle00chap_21/page/29/mode/1up

This is typical of early auction catalogs (Chapman catalogs, the premier numismatic auction house of the day, is used here). This just includes typed text, with lot numbers and item descriptions.

#### 2. Priced, Not Named
Example Catalog: Catalogue of the Collection of United States Coins of E. S. Norris, Esq., Boston, Massachusetts (S.H. & H. Chapman, 5/1894)  
Link to Original: https://archive.org/details/catalogueofcolle00chap_17/page/3/mode/1up

The format is the same as the previous one, but it is "priced," meaning that an attendee of the sale wrote the prices realized for each lot into the catalog during the sale.  This is very useful to numismatists today for pedigree research.  It also adds significant complications for the parser, as it now has to differentiate the handwritten prices from the lot numbers, both of which are numbers, and has to do this when the prices are not always in the same location (though in this example, they are).

#### 3. Named and Priced
Example Catalog: Collection of Foreign Coins and Medals of Mr. A. Galpin (S.H. & H. Chapman, 5/1883)  
Link to Original: https://archive.org/details/collectionoffore1883chap/page/13/mode/1up

The format is the same as the previous one, but it is "named and priced," meaning that an attendee of the sale wrote both the prices realized and the names of the purchasers into the catalog.  This significantly increases the value to numismatic researchers for pedigree purposes, but also significantly increases the difficulty for the parser, due to the 19th century handwriting and especially because of its rushed nature in a fast-paced auction.

#### 4. Interspersed Images
Example Catalog: The Amon G. Carter, Jr. Family Collection of United States Gold, Silver & Copper Coins and Foreign Coins (Stack's, 1/1984)  
Link to Original: https://archive.org/details/amongcarterjrfam1984stac/amongcarterjrfam1984stac/page/112/mode/1up

The Chapman's format lasted until the 1930s or 1940s, and then gave way to this format, which lasted into the 1980s or 1990s, depending on the auction house. This format is similar to the previous, but due to advances in photographic technology and decreases in price, now include images directly above the description of that item, rather than in a dedicated photographic plate like the Chapmans had to do.  However, these photographs were not included for every lot, providing a challenge for the scraper.  The lots with images also include bolded elements, which represent a shortened version of the description to identify the item at a quick glance (Short_Description).  The most important lots also include a Headline atop of it, drawing attention to the lot, which can be multiple lines, bold, capitalized, or some combination of them.  The era of named and priced sales had passed by this point, due to changes in the hobby and the rise of mail bids and auction agents that obscured the bidder's identity, so this feature will not be included from here on out.

#### 5. Complex Layouts
Example Catalog #5a: The Jascha Heifetz Collection Sale (Superior, 10/1989)  
Link to Original #5a: https://archive.org/details/jaschaheifetzcol1989supe/page/96/mode/1up  
Example Catalog #5b: The Long Beach Sale (Heritage, 9/1996)  
Link to Original #5b: https://archive.org/details/september1996lon1996heri/page/63/mode/1up

While some auction companies continued with the previous format for a long period of time, others had two-columns of lots, but with the same general format (including interspersed images), starting in the 1970s or 1980s. Short_Descriptions are nearly universal, and Headlines are provided for almost every imaged lot, with images becoming more and more common as time goes by.  Two files exist: one for Superior Auctions (5a), and one for Heritage Auctions (5b). The latter was included to see if the model made for the former was generalizable (and thus was not transcribed), and it was generalizable.

## Process

Download the two Python files, `config.py` and `auction_parser.py`. Open both, and run `config.py`. It will obtain all of the information it needs from `auction_parser.py` without having to run the latter file. The user will have to set up their own instance of the Layout Parser at https://console.cloud.google.com/ai/document-ai.  Once at that link, follow these steps:
1. Click "CREATE PROCESSOR"
2. Select "Layout Parser" (under Document OCR)
3. Choose your region
4. Name it "auction-catalog-parser"
5. Click "CREATE"
6. **Copy the Processor ID** from the details page

## Output

The model outputs an Excel file, `auction_catalog_parser.xlsx`. The file `example_auction_catalog_parser.xlsx` contains a hand-typed version of the ideal.  However, it excludes file 5b, since as mentioned previously, that is used to test the generalizability of the scraper using the most complex possible format.

Here is a Data Dictionary for the outputted file:

| Column | Description |
| --- | --- |
| Catalog_Source    | The name of the parsed file that the lot came from. |
| Lot_No            | Lot Number |
| Page_Start        | Unimplemented - Page of catalog |
| Page_End          | Unimplemented - Page of catalog |
| Catalog_Section   | Unimplemented - Section of catalog (Half Cents, Early Half Dollars, etc.) |
| Headline          | For expensive lots, headline above the image and description |
| Image_Link        | Unimplemented - Plan to crop images and automatically upload them to Google Cloud |
| Short_Description | For newer catalogs, the bolded portion at the start of the lot description |
| Long_Description  | Non-bolded item description |
| Pedigree          | For newer catalogs, the italicized section after the description that indicates prior ownership or sale appearances |
| Sale_Price        | For named and priced catalogs, the price the coin sold for |
| Sold_To           | For named and priced catalogs, the successful bidder, as annotated in the catalog |
| Year              | Year of the coin |
| Grade             | Grade of the coin. For non-certified coins, this is as described in the Long_Description, Sheldon Scale or not. |
| Grading_Service   | For post-1986 sales of certified coins, the grading service that certified it. |
| Variety           | For pre-1840 U.S. coins, the die marriage (such as "Newcomb-1", etc.) |
| Rarity            | Tha rarity of the die marriage |

## Accuracy

#### Lot_No

The parser picks up the following number of lots from each catalog:

| Catalog | Parsed Lots | Actual Lots | Accuracy |
| --- | --- | --- | --- |
| 1. No Annotations (Bidwell and Cottier, 1885)                 |  | 67 |  |
| 2. Priced, Not Named (Norris, 1894)                           |  | 43 |  |
| 3. Named and Priced (Galpin Foreign Coins and Medals, 5.1883) |  | 69 |  |
| 4. Interspersed Images                                        |  | 21 |  |
| 5a. Complex Layouts Superior                                  |  | 76 |  |
| 5b. Complex Layouts Heritagex                                 |  | 60 |  |

(False alarms)

#### Page_PDF



#### Catalog_Section



#### Headline



#### Image



#### Year



#### Grade



#### Grading Service



#### Variety and Rarity



#### Short_Description



#### Long_Description



#### Pedigree



#### Sale_Price and Sold_To

For these columns, only the Named and/or Priced catalogs (#2 and 3) are in play, as the others have no way of having this information available in-line with the lots.  The code ignores these columns for non-named and/or priced sales, which solves an additional problem of the program mistaking lot numbers for prices realized.

| Catalog | Parsed Lots | Actual Lots | Accuracy |
| --- | --- | --- | --- |
| 2. Priced, Not Named (Norris, 1894)                           |  | 43 |  |
| 3. Named and Priced (Galpin Foreign Coins and Medals, 5.1883) |  | 69 |  |

## Limitations

This project has been more successful than the author initially intended, but there are still significant shortcomings:

* The scraper still has issues differentiating the handwritten text from the typed text in named and priced sales. This leads to a number of issues, including nonsense characters and mistaking the years of the coins for lot numbers. The author has created versions that work slightly better for non-named and priced sales, but fall flat on their face for named and priced sales.
* The bold-text Short Descriptions are not reliably differentiated from the remaining text.  This is in part due to a limitation of Google's Layour Parser, which cannot natively identify and differentiate bold text.
* Lot numbers for the next lot are sometimes included as prices realized for the current lot, even when the next lot is properly parsed.
* There is no framework to account for group lots. A few do exist in the data, and the formatting differences should be minor. But due to their rarity (both in these catalogs and in general), a comprehensive evaluation of the failures relating to group lots has not been undertaken.
* There are still various errors here and there.  One, which happens only occasionally now, is that multi-word grades are not captured properly for old catalogs from before there was a standardized coin grading system (such as just capturing "Fine" instead of "Extremely Fine" or "Ex. Fine"; even though these terms are hard-coded into the list of grades to account for).
