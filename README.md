# bookstore-ebooks
Scripts and instructions for combining metadata from Alma, Alma Analytics, and another source (for example, a spreadsheet of assigned course books from a campus bookstore) to create a spreadsheet of library-owned course-assigned ebooks with proxied URLs.

## Scripts

**alma_sru_sn.py**

This script searches a list of ISBNs in Alma's standard number index via SRU, 
and generates an output file of found ISBNs with MMS IDs. Its output is used to
feed an Alma Analytics report (to determine e-book availability for ISBNs returned as found),
and to provide a lookup for the merge of Bookstore and Alma data produced by bookstore_file_merge.py.

**bookstore_file_merge.py**

This script produces a multiple-worksheet Excel .xlsx file of book metadata with proxied URLs derived from
four input files: a .csv file of library-owned or licensed ebooks with Portfolio IDs from Alma Analytics;
a .csv file of Portfolio IDs and proxied URLs from Alma's Export URLs job; a .txt file to serve as a
concordance between Alma MMS IDs and bookstore-supplied ISBNs; and an Excel spreadsheet supplied
by the bookstore (or other partner) of course books assigned for a given semester.

## Instructions

1. Download and open the source spreadsheet. Save the full list of ISBNs from bookstore as either a .txt file or a single-column .csv file to use as input for alma_sru_sn.py.
2. Rename the ISBN column in the spreadsheet to "ISBN" (if the column name isn't an exact match, the final match/merge process won't work. There may be other inconsistencies in your source spreadsheet; be prepared to make changes if needed).
3. Run alma_sru_sn.py with ISBN .txt or .csv file as input.
4. When that's finished, open Alma Analytics. Find the report "Bookstore print ISBNs with E inventory - template" (path: /Shared Folders/Community/Reports/Institutions/UMinnesota) and make a copy. Edit the MMS ID filter in the copied report to contain the list of MMS IDs from isbns_found.txt.
5. Save that report and export as .csv. Move it to the working directory.
6. In Alma, create an itemized Electronic Portfolios set from the Portfolio IDs in the Analytics report.
7. In Alma, run the Export URLs job on the portfolios set. Download the resulting file and move it to the working directory.
8. Run bookstore_file_merge.py. The final spreadsheet should match the Excel mashup template spreadsheet in this repo. Some minimal manual wrangling may be required. 



