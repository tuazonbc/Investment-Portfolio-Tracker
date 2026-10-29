"""Get market prices for the investments that are still held.

Run this file with: python python/refresh_prices.py
Excel can also run it through the Refresh prices button.
"""

import argparse
import csv
import json
import math
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

FIELDS = [
    "SecurityID", "ProviderSymbol", "ProviderName", "Currency", "Exchange",
    "Price", "FXtoCAD", "PriceDate", "QuoteTimestamp", "FXDate", "Status", "Error"
]


def get_price(symbol):
    """Ask Yahoo Finance for one symbol and check the price."""
    safe_symbol = urllib.parse.quote(symbol, safe="")
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" + safe_symbol
    url = url + "?interval=1d&range=5d"

    request = urllib.request.Request(
        url,
        headers={"User-Agent": "Mozilla/5.0"}
    )

    with urllib.request.urlopen(request, timeout=12) as response:
        answer = json.load(response)

    results = answer.get("chart", {}).get("result")
    if not results:
        error = answer.get("chart", {}).get("error", "No price returned")
        raise ValueError(str(error))

    details = results[0]["meta"]
    price = details.get("regularMarketPrice")
    price_time = details.get("regularMarketTime")

    if not isinstance(price, (int, float)):
        raise ValueError("The price is missing")
    if not math.isfinite(price) or price <= 0:
        raise ValueError("The price must be a positive number")
    if not price_time:
        raise ValueError("The price date is missing")

    price_date = datetime.fromtimestamp(price_time, timezone.utc)
    age_seconds = (datetime.now(timezone.utc) - price_date).total_seconds()
    if age_seconds < -86400 or age_seconds > 7 * 86400:
        raise ValueError("The price date is outside the allowed seven days")

    return details, price_date


def check_security_price(security, exchange_rate):
    """Check the listing before using its price in the workbook."""
    security_id = security["SecurityID"]
    symbols = security.get("ProviderSymbol", "")

    # A sold investment does not need a current price.
    if security.get("ActiveFlag") == "0":
        return [
            security_id, symbols, "", security.get("ExpectedCurrency", ""),
            "", "", "", "", "", "", "NOT_REQUIRED",
            "Closed position; no current quote needed"
        ]

    last_error = "No provider symbol configured"

    # Some listings have a second symbol that can be tried if the first fails.
    for symbol in symbols.split(";"):
        symbol = symbol.strip()
        if symbol == "":
            continue

        try:
            details, price_date = get_price(symbol)
            name = details.get("longName") or details.get("shortName") or ""
            name_check = security.get("ExpectedNameToken", "").strip()

            if name_check == "":
                raise ValueError("Enter a name to check for this investment")
            if name_check.casefold() not in name.casefold():
                raise ValueError("The returned company name did not match: " + name)

            currency = details.get("currency")
            if currency != security["ExpectedCurrency"]:
                raise ValueError("The returned currency did not match")

            # A CDR and the original company shares are different investments.
            full_name = (name + " " + details.get("shortName", "")).upper()
            if security.get("AssetType") == "CDR":
                if "CDR" not in full_name and "DEPOSITARY" not in full_name:
                    raise ValueError("A CDR was expected, but another listing was returned")

            if currency == "CAD":
                cad_rate = 1
                fx_date = price_date.date().isoformat()
            elif currency == "USD":
                if isinstance(exchange_rate, Exception):
                    raise ValueError("The USD to CAD rate was not available: " + str(exchange_rate))
                fx_details, fx_time = exchange_rate
                cad_rate = fx_details["regularMarketPrice"]
                fx_date = fx_time.date().isoformat()
            else:
                raise ValueError("Only CAD and USD prices are supported")

            return [
                security_id, symbol, name, currency, details.get("exchangeName", ""),
                details["regularMarketPrice"], cad_rate, price_date.date().isoformat(),
                price_date.isoformat(), fx_date, "OK", ""
            ]

        except Exception as error:
            last_error = type(error).__name__ + ": " + str(error)

    # Leave missing prices blank. A zero would give the wrong portfolio value.
    return [
        security_id, symbols, "", "", "", "", "", "", "", "", "FAILED", last_error
    ]


def update_prices(symbol_file, output_file):
    """Read the investment list, get prices and save the results."""
    with symbol_file.open(encoding="utf-8-sig", newline="") as file:
        securities = list(csv.DictReader(file))

    try:
        exchange_rate = get_price("CAD=X")
    except Exception as error:
        exchange_rate = error

    rows = []

    # Up to five requests can run together so the Excel refresh is quicker.
    with ThreadPoolExecutor(max_workers=5) as price_requests:
        waiting = []
        for security in securities:
            request = price_requests.submit(check_security_price, security, exchange_rate)
            waiting.append(request)
        for request in waiting:
            rows.append(request.result())

    output_file.parent.mkdir(parents=True, exist_ok=True)
    temporary_file = output_file.with_suffix(".tmp")
    with temporary_file.open("w", encoding="utf-8-sig", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(FIELDS)
        writer.writerows(rows)

    # Only replace the old output after the new file has finished writing.
    temporary_file.replace(output_file)

    prices_updated = 0
    prices_failed = 0
    closed_positions = 0
    saved_quotes = []

    for row in rows:
        if row[10] == "OK":
            prices_updated = prices_updated + 1
        elif row[10] == "FAILED":
            prices_failed = prices_failed + 1
        elif row[10] == "NOT_REQUIRED":
            closed_positions = closed_positions + 1
        saved_quotes.append(dict(zip(FIELDS, row)))

    summary = {
        "retrieved_utc": datetime.now(timezone.utc).isoformat(),
        "ok": prices_updated,
        "failed": prices_failed,
        "closed": closed_positions,
        "source": "Yahoo Finance chart endpoint (unofficial); latest available quote, potentially delayed",
        "quotes": saved_quotes
    }

    output_file.with_suffix(".json").write_text(
        json.dumps(summary, indent=2), encoding="utf-8"
    )
    print("Prices updated:", prices_updated)
    print("Prices needing attention:", prices_failed)
    print("Closed positions skipped:", closed_positions)
    return summary


if __name__ == "__main__":
    # Excel uses these options to tell Python where to read and save the files.
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", type=Path, default=ROOT / "data/price_symbols.csv")
    parser.add_argument("--output", type=Path, default=ROOT / "data/prices_latest.csv")
    settings = parser.parse_args()
    update_prices(settings.manifest, settings.output)
