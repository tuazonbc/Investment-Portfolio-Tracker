"""Check the price rules without making any internet requests.

Run: python python/test_prices.py
"""

import unittest
from datetime import datetime, timezone
from unittest.mock import patch

from refresh_prices import check_security_price


class PriceChecks(unittest.TestCase):
    def setUp(self):
        # Each test starts with the same example CDR and price.
        self.security = {
            "SecurityID": "AAPL_CDR",
            "ProviderSymbol": "AAPL.TO",
            "ExpectedCurrency": "CAD",
            "ExpectedNameToken": "Apple",
            "AssetType": "CDR"
        }
        self.price = {
            "longName": "Apple Inc.",
            "shortName": "APPLE CDR (CAD HEDGED)",
            "currency": "CAD",
            "regularMarketPrice": 49.02,
            "exchangeName": "TOR"
        }
        self.price_time = datetime.now(timezone.utc)

    def test_cdr_and_cad_price(self):
        # patch supplies an example response instead of using the internet.
        with patch("refresh_prices.get_price", return_value=(self.price, self.price_time)):
            result = check_security_price(self.security, None)
        self.assertEqual(result[10], "OK")
        self.assertEqual(result[6], 1)

    def test_wrong_company(self):
        self.price["longName"] = "Another company"
        with patch("refresh_prices.get_price", return_value=(self.price, self.price_time)):
            result = check_security_price(self.security, None)
        self.assertEqual(result[10], "FAILED")

    def test_wrong_currency(self):
        self.price["currency"] = "USD"
        with patch("refresh_prices.get_price", return_value=(self.price, self.price_time)):
            result = check_security_price(self.security, None)
        self.assertEqual(result[10], "FAILED")

    def test_closed_position(self):
        self.security["ActiveFlag"] = "0"
        with patch("refresh_prices.get_price") as price_request:
            result = check_security_price(self.security, None)
        self.assertEqual(result[10], "NOT_REQUIRED")
        price_request.assert_not_called()

    def test_usd_needs_exchange_rate(self):
        self.security["SecurityID"] = "AAPL_US"
        self.security["ProviderSymbol"] = "AAPL"
        self.security["ExpectedCurrency"] = "USD"
        self.security["AssetType"] = "Equity"
        self.price["currency"] = "USD"

        with patch("refresh_prices.get_price", return_value=(self.price, self.price_time)):
            result = check_security_price(self.security, ValueError("FX unavailable"))
            self.assertEqual(result[10], "FAILED")

            example_rate = ({"regularMarketPrice": 1.4}, self.price_time)
            result = check_security_price(self.security, example_rate)
            self.assertEqual(result[10], "OK")
            self.assertEqual(result[6], 1.4)


if __name__ == "__main__":
    unittest.main()
