# Checks completed

## Data

1. The 17 supplied files were checked.
2. One exact duplicate was left out, leaving 16 statements and 269 entries.
3. Running the cleaner twice gives the same transaction file.
4. The rewritten cleaner gives the same file as the earlier code, including every RecordID.
5. The NVIDIA split gives 12.0017 units still held and CAD 374.71 of remaining cost.
6. Contributions total CAD 19,475. Dividends total CAD 339.67.
7. Calculated cash is CAD 586.35 and matches the final statement.

## Prices

1. The rewritten price script retrieved prices for all 16 open investments in a separate test folder.
2. Ten investments with no units were skipped.
3. Five tests checked the right company, currency, CDR listing, USD conversion and skipping sold investments.
4. A failed price stays blank or is marked as needing attention. It is not entered as zero.

## Workbook

1. Python checked 2,152 transaction calculations against the saved Excel results.
2. Holdings, values, gains, cash and cover counts matched.
3. Investment weights added to 100%.
4. Four charts were found in the workbook.
5. No saved Excel formula errors were found.
6. The workbook file was not changed while the code was simplified.

## Buttons

The six button actions and their worksheet links were kept. The VBA source was checked for missing blocks, broken lines and changes to the button setup.

The code has not been run inside desktop Excel as part of these checks. Use the small contribution test in VBA_GUIDE.md before relying on the buttons for new records.

The earlier workbook test showed that a CAD 100 contribution adds CAD 100 to cash and contributions. At unchanged prices, value rises by CAD 100 and gain stays the same. The file used to check Excel formulas needed its table references refreshed for that test. Desktop Excel still needs its own test.
