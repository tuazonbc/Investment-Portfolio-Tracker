"""Check the supplied Excel results with separate Python calculations.

Run: python python/validate_workbook.py
Use the original sample workbook, without added test transactions.
"""

import csv
import json
import xml.etree.ElementTree as ET
import zipfile
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
XML_NAMES = {"excel": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def read_saved_cells(workbook_file):
    """Read saved Excel values without needing Excel or an extra package.

    An xlsx file is a ZIP containing XML files. This part only reads them.
    It does not edit the workbook.
    """
    shared_strings = []
    if "xl/sharedStrings.xml" in workbook_file.namelist():
        text_file = ET.fromstring(workbook_file.read("xl/sharedStrings.xml"))
        for item in text_file:
            shared_strings.append("".join(item.itertext()))

    sheet_files = {}
    links = ET.fromstring(workbook_file.read("xl/_rels/workbook.xml.rels"))
    for link in links:
        sheet_files[link.attrib["Id"]] = link.attrib["Target"]

    workbook = ET.fromstring(workbook_file.read("xl/workbook.xml"))
    sheet_list = workbook.findall("excel:sheets/excel:sheet", XML_NAMES)
    all_sheets = {}

    for sheet in sheet_list:
        sheet_name = sheet.attrib["name"]
        link_name = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        sheet_link = sheet.attrib[link_name]
        file_name = sheet_files[sheet_link]

        if file_name.startswith("/"):
            file_name = file_name.lstrip("/")
        else:
            file_name = "xl/" + file_name

        sheet_xml = ET.fromstring(workbook_file.read(file_name))
        cells = {}

        for cell in sheet_xml.findall(".//excel:c", XML_NAMES):
            address = cell.attrib["r"]
            cell_type = cell.attrib.get("t")
            assert cell_type != "e", ("Excel error", sheet_name, address)

            saved_value = cell.find("excel:v", XML_NAMES)
            if saved_value is not None:
                value = saved_value.text
                if cell_type == "s":
                    value = shared_strings[int(value)]
                cells[address] = value

        all_sheets[sheet_name] = cells

    return all_sheets


def close_enough(actual, expected):
    # Allow only a very small difference from Excel's decimal rounding.
    return abs(Decimal(actual) - expected) < Decimal("0.00001")


def main():
    with (ROOT / "data/transactions_clean.csv").open(encoding="utf-8-sig") as file:
        transactions = list(csv.DictReader(file))

    price_file = json.loads((ROOT / "data/prices_latest.json").read_text())
    prices = {}
    for price in price_file["quotes"]:
        prices[price["SecurityID"]] = price

    with zipfile.ZipFile(ROOT / "Portfolio_Dashboard.xlsx") as workbook_file:
        sheets = read_saved_cells(workbook_file)
        charts = []
        for file_name in workbook_file.namelist():
            if "/charts/chart" in file_name and file_name.endswith(".xml"):
                charts.append(file_name)

    transaction_cells = sheets["Transactions"]
    security_cells = sheets["Securities"]
    dashboard = sheets["Dashboard"]
    assert Decimal(sheets["Cover"]["C22"]) == 269
    assert Decimal(sheets["Cover"]["C23"]) == 16

    positions = {}
    realized_gains = Decimal(0)
    checks_done = 0

    # Work through the transactions in the same order as the workbook.
    for excel_row, transaction in enumerate(transactions, start=6):
        security_id = transaction["SecurityID"]
        if security_id not in positions:
            positions[security_id] = [Decimal(0), Decimal(0)]

        units, cost = positions[security_id]
        quantity = Decimal(transaction["Quantity"])
        amount = Decimal(transaction["AmountCAD"])
        transaction_type = transaction["Type"]

        average_cost = Decimal(0)
        if units > Decimal("0.000001"):
            average_cost = cost / units

        unit_change = Decimal(0)
        cost_change = Decimal(0)
        gain = Decimal(0)

        if transaction_type == "BUY":
            unit_change = quantity
            cost_change = -amount
        elif transaction_type == "SELL":
            assert quantity <= units + Decimal("0.000001"), "Sale exceeds units"
            unit_change = -quantity
            cost_change = -quantity * average_cost
            gain = amount - quantity * average_cost
        elif transaction_type == "STKREORG":
            unit_change = quantity

        # These are the calculation columns on the Transactions sheet.
        expected_cells = {
            "O": unit_change,
            "P": units,
            "Q": cost,
            "R": average_cost,
            "S": cost_change,
            "T": gain,
            "U": units + unit_change,
            "V": cost + cost_change
        }
        for column, expected_value in expected_cells.items():
            actual = transaction_cells[column + str(excel_row)]
            assert close_enough(actual, expected_value), (excel_row, column)
            checks_done = checks_done + 1

        positions[security_id] = [units + unit_change, cost + cost_change]
        realized_gains = realized_gains + gain

    holdings_value = Decimal(0)
    remaining_cost = Decimal(0)
    total_weight = Decimal(0)
    open_positions = 0

    for address, security_id in security_cells.items():
        if not address.startswith("B"):
            continue
        row_number = address[1:]
        if not row_number.isdigit() or int(row_number) < 6:
            continue

        units, cost = positions[security_id]
        assert abs(Decimal(security_cells["M" + row_number]) - units) < Decimal("0.000001")

        if units > Decimal("0.000001"):
            price = prices[security_id]
            assert price["Status"] == "OK"
            value = units * Decimal(str(price["Price"])) * Decimal(str(price["FXtoCAD"]))
            assert close_enough(security_cells["R" + row_number], value)
            assert security_cells["U" + row_number] == "Current"

            holdings_value = holdings_value + value
            remaining_cost = remaining_cost + cost
            total_weight = total_weight + Decimal(security_cells["T" + row_number])
            open_positions = open_positions + 1
        else:
            assert security_cells["U" + row_number] == "Closed"

    cash = Decimal(0)
    contributions = Decimal(0)
    income = Decimal(0)
    for transaction in transactions:
        amount = Decimal(transaction["AmountCAD"])
        cash = cash + amount
        if transaction["Type"] == "CONT":
            contributions = contributions + amount
        if transaction["Type"] in ("DIV", "REFER", "NRT"):
            income = income + amount

    total_value = holdings_value + cash
    total_gain = total_value - contributions
    expected_dashboard = {
        "B6": total_value,
        "F6": contributions,
        "J6": total_gain,
        "D10": cash,
        "L10": realized_gains
    }
    for address, expected_value in expected_dashboard.items():
        assert close_enough(dashboard[address], expected_value), address

    unrealized_gains = holdings_value - remaining_cost
    assert abs(total_gain - (unrealized_gains + realized_gains + income)) < Decimal("0.00001")
    assert abs(total_weight - 1) < Decimal("0.00000001")
    assert positions["NVDA_CDR"][0] == Decimal("12.0017")
    assert close_enough(positions["NVDA_CDR"][1], Decimal("374.71"))
    assert open_positions == 16
    assert len(charts) == 4

    print("PASS:", checks_done, "ledger comparisons, holdings, weights, cash, gains and cover counts.")


if __name__ == "__main__":
    main()
