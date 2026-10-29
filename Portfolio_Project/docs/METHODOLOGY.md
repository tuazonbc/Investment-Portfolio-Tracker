# How the calculations work

## Where the data came from

The project uses CSV statements from a third party account. The files cover July 2024 to October 2025.

There were 17 files. Two November files were exactly the same, so one copy was left out. The other 16 files contain 269 entries.

The original CSV contents were kept. The account number was removed from the filenames. Each clean entry keeps its source filename and row number, so it can be checked against the original.

## The NVIDIA split

Some statements call the investment Nvidia Corp. Later statements call it Nvidia CDR. These records appear to refer to the same Canadian CDR, rather than US shares.

[CIBC announced a four for one CDR split](https://cibc.mediaroom.com/2024-07-04-CIBC-Announces-Nvidia-Canadian-Depositary-Receipts-Split). Trading at the new unit count started on July 10, 2024. The statement has a reorganization entry on that date.

The account had one unit before the split. After the split, it had four units. The total amount paid stayed at CAD 120.27.

The later entries also fit this reading. Four units plus another 15 purchased units make the 19 units that were then sold.

The code uses this correction for these particular records. It does not assume that every investment with the NVDA ticker is a CDR. The original descriptions remain in the CSVs, and the changes are listed in `data/normalization_log.csv`.

## Purchases and average cost

A purchase adds units and the amount paid to the investment's cost.

For example, buying two units for CAD 20 and then two more for CAD 30 gives four units with a total cost of CAD 50. The average cost is CAD 12.50 per unit.

The project uses this moving average method. Amounts are taken from the actual CAD cash entries in the statements, including costs already included by the brokerage.

## Sales and gains

A sale removes units at their average cost.

Using the example above, selling one unit for CAD 16 removes CAD 12.50 of cost. The gain on that sale is CAD 3.50.

The workbook keeps gains from sold units separate from gains on units still held. The original purchase and sale records stay in the file even when an investment is fully sold.

## Current value

The value of an investment is its recorded units multiplied by the share price. If the quote is in US dollars, the result is also multiplied by the CAD conversion rate.

The portfolio value is the sum of the investment values and the recorded cash balance.

Total gain is that value less the money contributed. It includes changes in share values, gains from sales, dividends, the referral credit and withholding tax.

The investment weights add to 100%. Cash is kept outside those investment weights.

## Share prices

Python gets the latest available prices from Yahoo Finance. It checks the listing symbol, name and currency before accepting a price. It also checks that a CDR has not been mixed up with the original company shares.

Only investments still held need a current price. Sold investments stay in the history without requesting a new quote.

The Yahoo connection is unofficial. Quotes may be delayed, and requests can fail. The project saves the price dates and errors. It asks for another refresh if a price or exchange rate is more than seven days old.

## Checking the totals

The records show CAD 19,475 in contributions and CAD 339.67 in dividends.

Adding all cash entries gives CAD 586.35. This matches the final statement balance.

Three earlier balances shown by the statements differ slightly from the running cash calculation. The differences are CAD 0.07, CAD 0.04 and negative CAD 0.05. They are listed in the Review sheet. The transaction amounts were not changed to force a match.

## What the results mean

The workbook uses the historical units and cash in the files with the saved market prices. It does not assume there were no later trades in the actual account. The result is an example valuation, not a confirmed current account balance.

These are portfolio tracking calculations, not a tax report. The dashboard shows dollar gains. It does not calculate an annual return or compare performance with a market index.

## Adding more records

Each imported row has a RecordID. Importing the same clean file again skips IDs already in Excel.

If a trade was entered manually and later arrives in a broker CSV, it needs to be checked so it is not counted twice. An older trade also needs to be placed in the right order because it can change the average cost used by later sales.
