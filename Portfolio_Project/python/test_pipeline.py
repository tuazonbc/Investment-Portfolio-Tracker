"""Check the sample data without changing the original project files.

Run: python python/test_pipeline.py
"""

import csv
import hashlib
import shutil
import tempfile
from decimal import Decimal
from pathlib import Path

from clean_transactions import ROOT, clean


def output_hash(folder):
    file = folder / "data/transactions_clean.csv"
    return hashlib.sha256(file.read_bytes()).hexdigest()


# A temporary folder lets the tests copy and change files safely.
with tempfile.TemporaryDirectory() as folder_name:
    test_folder = Path(folder_name)
    shutil.copytree(ROOT / "raw", test_folder / "raw")

    first_run = clean(test_folder)
    first_hash = output_hash(test_folder)
    second_run = clean(test_folder)

    assert first_hash == output_hash(test_folder), "Running twice changed the clean data"
    assert first_run["files_received"] == 17
    assert first_run["files_imported"] == 16
    assert first_run["rows"] == 269
    assert Decimal(str(first_run["net_contributions"])) == Decimal("19475.0")
    assert Decimal(str(first_run["dividends"])) == Decimal("339.67")
    assert Decimal(str(first_run["cash"])) == Decimal("586.35")
    assert first_run["cash"] == first_run["last_statement_balance"]

    clean_file = test_folder / "data/transactions_clean.csv"
    with clean_file.open(encoding="utf-8-sig") as file:
        transactions = list(csv.DictReader(file))

    record_ids = set()
    nvidia_ids = set()
    split_rows = []

    for transaction in transactions:
        record_ids.add(transaction["RecordID"])
        if transaction["SecurityID"].startswith("NVDA"):
            nvidia_ids.add(transaction["SecurityID"])
        if transaction["Type"] == "STKREORG":
            split_rows.append(transaction)

    assert len(record_ids) == 269
    assert nvidia_ids == {"NVDA_CDR"}
    assert len(split_rows) == 1
    assert Decimal(split_rows[0]["Quantity"]) == 3

    # A second copy of a statement must not add more transactions.
    first_statement = next((test_folder / "raw").glob("*.csv"))
    shutil.copy(first_statement, test_folder / "raw/duplicate_test.csv")
    duplicate_run = clean(test_folder)
    assert duplicate_run["rows"] == 269

print("PASS: clean data, duplicate files, record IDs, NVIDIA split and cash totals.")
