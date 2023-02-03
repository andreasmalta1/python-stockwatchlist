import yfinance as yf
import gspread


FAST_INFO_KEYS = [
    "regular_market_previous_close",
    "year_high",
    "year_low",
    "market_cap",
]
INFO_KEYS = [
    "totalRevenue",
    "netIncomeToCommon",
    "trailingEps",
    "dividendRate",
    "sector",
]


def main():
    sa = gspread.service_account(filename="service_account.json")
    sh = sa.open("Stocks Watchlist")
    whs = sh.worksheet("Sheet1")

    ticker_list = whs.col_values(2)
    last_row = len(ticker_list)

    information = {
        "ticker": [],
        "regular_market_previous_close": [],
        "year_high": [],
        "year_low": [],
        "market_cap": [],
        "totalRevenue": [],
        "netIncomeToCommon": [],
        "trailingEps": [],
        "dividendRate": [],
        "sector": [],
    }

    for i in range(2, len(ticker_list) + 1):
        information["ticker"].append([yf.Ticker(ticker_list[i - 1])])

    for ticker in information["ticker"]:
        ticker = ticker[0]
        print(ticker)
        for key in information:
            if key == "ticker":
                continue

            if key in FAST_INFO_KEYS:
                try:
                    information[key].append(ticker.fast_info[key])
                except KeyError:
                    information[key].append([""])

            if key in INFO_KEYS:
                try:
                    information[key].append(ticker.info[key])
                except KeyError:
                    information[key].append([""])

    whs.batch_update(
        [
            {
                "range": "D2" + ":D" + str(last_row),
                "values": information["regular_market_previous_close"],
            },
            {
                "range": "E2" + ":E" + str(last_row),
                "values": information["year_low"],
            },
            {
                "range": "F2" + ":F" + str(last_row),
                "values": information["year_high"],
            },
            {
                "range": "H2" + ":H" + str(last_row),
                "values": information["market_cap"],
            },
            {
                "range": "I2" + ":I" + str(last_row),
                "values": information["totalRevenue"],
            },
            {
                "range": "J2" + ":J" + str(last_row),
                "values": information["netIncomeToCommon"],
            },
            {
                "range": "K2" + ":K" + str(last_row),
                "values": information["trailingEps"],
            },
            {
                "range": "M2" + ":M" + str(last_row),
                "values": information["dividendRate"],
            },
            {
                "range": "O2" + ":O" + str(last_row),
                "values": information["sector"],
            },
        ]
    )


main()
