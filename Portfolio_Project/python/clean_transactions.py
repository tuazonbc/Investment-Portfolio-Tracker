"""Combine the monthly statements into one clean transaction file.

Run this file with: python python/clean_transactions.py
The original statements and the Excel workbook are not changed.
"""

import csv
import hashlib
import json
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SPLIT_SOURCE = "https://cibc.mediaroom.com/2024-07-04-CIBC-Announces-Nvidia-Canadian-Depositary-Receipts-Split"

HEADERS = [
    "RecordID", "Date", "TradeDate", "SecurityID", "Type", "Quantity",
    "AmountCAD", "Currency", "StatementBalance", "SourceFile", "SourceRow",
    "Description", "Status"
]


def save_csv(folder, name, headers, rows):
    """Save a CSV that Excel can read, including names with special characters."""
    with (folder / name).open("w", newline="", encoding="utf-8-sig") as file:
        writer = csv.writer(file)
        writer.writerow(headers)
        writer.writerows(rows)


def get_security_details(description):
    """Descriptions normally start with a ticker, then the investment name."""
    if " - " not in description or ":" not in description:
        return "", ""

    ticker, rest = description.split(" - ", 1)
    name = rest.split(":", 1)[0].strip()
    return ticker, name


def get_quantity(description):
    """Find the number between Bought or Sold and the word shares."""
    share_text = ""
    if "Bought " in description:
        share_text = description.split("Bought ", 1)[1]
    elif "Sold " in description:
        share_text = description.split("Sold ", 1)[1]

    if " shares" not in share_text:
        return Decimal(0), False

    number = share_text.split(" shares", 1)[0].replace(",", "")
    return Decimal(number), True


def transaction_order(transaction):
    # Keep the original row order when several entries have the same date.
    return transaction["Date"], transaction["SourceFile"], transaction["SourceRow"]


def clean(root=ROOT):
    data_folder = root / "data"
    data_folder.mkdir(exist_ok=True)

    seen_files = {}
    transactions = []
    import_log = []
    issues = []
    securities = {}
    changes = []

    # Read each statement. Identical complete files are imported only once.
    for statement in sorted((root / "raw").glob("*.csv")):
        file_hash = hashlib.sha256(statement.read_bytes()).hexdigest()
        with statement.open(encoding="utf-8-sig", newline="") as file:
            raw_rows = list(csv.DictReader(file))

        duplicate_name = seen_files.get(file_hash, "")
        action = "Imported"
        if duplicate_name:
            action = "Skipped identical file"
        import_log.append([statement.name, len(raw_rows), action, duplicate_name, file_hash])

        if duplicate_name:
            continue
        seen_files[file_hash] = statement.name

        # Row 1 contains the CSV headings, so the first transaction is row 2.
        for row_number, raw in enumerate(raw_rows, start=2):
            description = raw["description"]
            transaction_type = raw["transaction"].strip().upper()
            ticker, name = get_security_details(description)

            security_id = ""
            if ticker:
                if "CDR" in name:
                    security_id = ticker + "_CDR"
                else:
                    security_id = ticker + "_ORD"

            # This correction applies only to the known historical sample.
            # The issuer split notice and later sale units identify this as the CDR.
            if ticker == "NVDA" and name == "Nvidia Corp":
                if raw["date"] <= "2025-04-11" and "FX Rate" not in description:
                    security_id = "NVDA_CDR"
                    name = "Nvidia CDR (CAD Hedged)"
                    changes.append([
                        statement.name, row_number,
                        "Normalize legacy NVDA description to CDR",
                        "Issuer 4-for-1 split and 4+15=19 unit reconciliation",
                        SPLIT_SOURCE
                    ])

            if security_id and security_id not in securities:
                asset_type = "Equity"
                currency = ""
                if "CDR" in name:
                    asset_type = "CDR"
                    currency = "CAD"
                elif "ETF" in name or "Fund" in name:
                    asset_type = "ETF"
                    currency = "CAD"

                securities[security_id] = [
                    security_id, ticker, name, asset_type, "Other", "", currency,
                    None, None, None, ""
                ]

            quantity, quantity_found = get_quantity(description)
            trade_date = raw["date"]
            if "executed at " in description:
                trade_date = description.split("executed at ", 1)[1][:10]

            status = "Imported"
            if transaction_type == "STKREORG":
                status = "Corporate action requires details"
            elif transaction_type in ("BUY", "SELL") and not quantity_found:
                status = "REVIEW missing quantity"
                issues.append([statement.name, row_number, status, description])

            # The same source file and row always receive the same ID.
            id_text = file_hash + ":" + str(row_number)
            record_id = hashlib.sha256(id_text.encode()).hexdigest()[:20]

            transaction = {
                "RecordID": record_id,
                "Date": raw["date"],
                "TradeDate": trade_date,
                "SecurityID": security_id,
                "Type": transaction_type,
                "Quantity": float(quantity),
                "AmountCAD": float(Decimal(raw["amount"])),
                "Currency": raw["currency"],
                "StatementBalance": float(Decimal(raw["balance"])),
                "SourceFile": statement.name,
                "SourceRow": row_number,
                "Description": description,
                "Status": status
            }
            transactions.append(transaction)

    transactions.sort(key=transaction_order)

    # Decimal keeps the running money calculations from losing small amounts.
    positions = {}
    for transaction in transactions:
        security_id = transaction["SecurityID"]
        if security_id not in positions:
            positions[security_id] = {"units": Decimal(0), "cost": Decimal(0)}

        position = positions[security_id]
        units = position["units"]
        cost = position["cost"]
        quantity = Decimal(str(transaction["Quantity"]))
        amount = Decimal(str(transaction["AmountCAD"]))
        transaction_type = transaction["Type"]

        if transaction_type == "STKREORG":
            if security_id == "NVDA_CDR" and transaction["Date"] == "2024-07-10":
                # One old CDR became four. Cost and cash did not change.
                transaction["Quantity"] = float(units * 3)
                transaction["Status"] = "Applied 4-for-1 CDR split"
                position["units"] = units * 4
                changes.append([
                    transaction["SourceFile"], transaction["SourceRow"],
                    "Apply issuer 4-for-1 split",
                    f"{units} units became {units * 4}; CAD cost unchanged",
                    SPLIT_SOURCE
                ])
            else:
                issues.append([
                    transaction["SourceFile"], transaction["SourceRow"],
                    "Corporate action requires details", transaction["Description"]
                ])

        elif transaction_type == "BUY":
            position["units"] = units + quantity
            position["cost"] = cost - amount

        elif transaction_type == "SELL":
            if quantity > units + Decimal("0.000001"):
                issues.append([
                    transaction["SourceFile"], transaction["SourceRow"],
                    "Sale exceeds available units", transaction["Description"]
                ])

            average_cost = Decimal(0)
            if units != 0:
                average_cost = cost / units
            position["units"] = units - quantity
            position["cost"] = cost - quantity * average_cost

    cash = Decimal(0)
    contributions = Decimal(0)
    dividends = Decimal(0)

    for transaction in transactions:
        amount = Decimal(str(transaction["AmountCAD"]))
        cash = cash + amount
        difference = Decimal(str(transaction["StatementBalance"])) - cash

        if abs(difference) > Decimal("0.011"):
            issues.append([
                transaction["SourceFile"], transaction["SourceRow"],
                "Cash balance differs from running cash by " + str(difference),
                transaction["Description"]
            ])
        if transaction["Type"] == "CONT":
            contributions = contributions + amount
        if transaction["Type"] == "DIV":
            dividends = dividends + amount

    # Use the same column order that the VBA importer expects.
    transaction_rows = []
    for transaction in transactions:
        row = []
        for heading in HEADERS:
            row.append(transaction[heading])
        transaction_rows.append(row)

    security_rows = sorted(securities.values())
    save_csv(data_folder, "transactions_clean.csv", HEADERS, transaction_rows)
    save_csv(data_folder, "securities_seed.csv", [
        "SecurityID", "Ticker", "Name", "AssetType", "Category", "Exchange",
        "QuoteCurrency", "Price", "FXtoCAD", "PriceDate", "PriceSource"
    ], security_rows)
    save_csv(data_folder, "import_log.csv", [
        "File", "Rows", "Action", "DuplicateOf", "SHA256"
    ], import_log)
    save_csv(data_folder, "normalization_log.csv", [
        "SourceFile", "SourceRow", "Change", "Evidence", "SourceURL"
    ], changes)
    save_csv(data_folder, "data_quality.csv", [
        "SourceFile", "SourceRow", "Issue", "Description"
    ], issues)

    # Keep the chosen symbols, but update which investments still need prices.
    symbol_file = data_folder / "price_symbols.csv"
    if symbol_file.exists():
        with symbol_file.open(encoding="utf-8-sig", newline="") as file:
            reader = csv.DictReader(file)
            headings = list(reader.fieldnames)
            mappings = list(reader)

        if "ActiveFlag" not in headings:
            headings.append("ActiveFlag")

        for mapping in mappings:
            position = positions.get(mapping["SecurityID"], {"units": Decimal(0)})
            if position["units"] > Decimal("0.000001"):
                mapping["ActiveFlag"] = "1"
            else:
                mapping["ActiveFlag"] = "0"

        with symbol_file.open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.DictWriter(file, fieldnames=headings)
            writer.writeheader()
            writer.writerows(mappings)

    summary = {
        "files_received": len(import_log),
        "files_imported": len(seen_files),
        "rows": len(transactions),
        "issues": len(issues),
        "net_contributions": float(contributions),
        "dividends": float(dividends),
        "cash": float(cash),
        "last_statement_balance": transactions[-1]["StatementBalance"],
        "first_date": transactions[0]["Date"],
        "last_date": transactions[-1]["Date"]
    }

    print(json.dumps(summary, indent=2))
    return summary


if __name__ == "__main__":
    clean()
