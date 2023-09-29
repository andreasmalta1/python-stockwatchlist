from yahooquery import Ticker
import gspread

close_price = []
price_low = []
price_high = []
market_cap = []
yearly_revenue = []
yearly_earnings = []
trailing_eps = []
trailing_pe = []
dividend_rate = []
stock_sector = []


def main():
    sa = gspread.service_account(filename="service_account.json")
    sh = sa.open("Stocks Watchlist")
    whs = sh.worksheet("Sheet1")

    ticker_list = whs.col_values(2)
    ticker_list.pop(0)
    last_row = len(ticker_list)
    last_row += 1

    for ticker in ticker_list:
        print(ticker)
        ticker_obj = Ticker(ticker)

        close_price_value = ticker_obj.summary_detail[ticker].get(
            "regularMarketPreviousClose"
        )
        if not close_price_value:
            close_price.append([None])
        else:
            close_price.append([close_price_value])

        price_low_value = ticker_obj.summary_detail[ticker].get("fiftyTwoWeekLow")
        if not price_low_value:
            price_low.append([None])
        else:
            price_low.append([price_low_value])

        price_high_value = ticker_obj.summary_detail[ticker].get("fiftyTwoWeekHigh")
        if not price_low_value:
            price_high.append([None])
        else:
            price_high.append([price_high_value])

        market_cap_value = ticker_obj.summary_detail[ticker].get("marketCap")
        if not market_cap_value:
            market_cap.append([None])
        else:
            market_cap.append([market_cap_value])

        try:
            yearly_revenue_value = ticker_obj.earnings[ticker]["financialsChart"][
                "yearly"
            ][-1].get("revenue")
            if not yearly_revenue_value:
                yearly_revenue.append([None])
            else:
                yearly_revenue.append([yearly_revenue_value])

        except TypeError:
            yearly_revenue.append([None])

        try:
            yearly_earnings_value = ticker_obj.earnings[ticker]["financialsChart"][
                "yearly"
            ][-1].get("earnings")
            if not yearly_earnings_value:
                yearly_earnings.append([None])
            else:
                yearly_earnings.append([yearly_earnings_value])

        except TypeError:
            yearly_earnings.append([None])

        try:
            trailing_eps_value = ticker_obj.key_stats[ticker].get("trailingEps")
            if not trailing_eps_value:
                trailing_eps.append([None])
            else:
                trailing_eps.append([trailing_eps_value])
        except AttributeError:
            trailing_eps.append([None])
        except TypeError:
            trailing_eps.append([None])

        trailing_pe_value = ticker_obj.summary_detail[ticker].get("trailingPE")
        if not trailing_pe_value:
            trailing_pe.append([None])
        else:
            trailing_pe.append([trailing_pe_value])

        dividend_rate_value = ticker_obj.summary_detail[ticker].get("dividendRate")
        if not dividend_rate_value:
            dividend_rate.append([None])
        else:
            dividend_rate.append([dividend_rate_value])

        try:
            stock_sector_value = ticker_obj.summary_detail[ticker].get("sector")
            if not stock_sector_value:
                stock_sector.append([None])
            else:
                stock_sector.append([stock_sector_value])
        except AttributeError:
            stock_sector.append([None])
        except TypeError:
            stock_sector.append([None])

    whs.batch_update(
        [
            {
                "range": "D2" + ":D" + str(last_row),
                "values": close_price,
            },
            {
                "range": "E2" + ":E" + str(last_row),
                "values": price_low,
            },
            {
                "range": "F2" + ":F" + str(last_row),
                "values": price_high,
            },
            {
                "range": "H2" + ":H" + str(last_row),
                "values": market_cap,
            },
            {
                "range": "I2" + ":I" + str(last_row),
                "values": yearly_revenue,
            },
            {
                "range": "J2" + ":J" + str(last_row),
                "values": yearly_earnings,
            },
            {
                "range": "K2" + ":K" + str(last_row),
                "values": trailing_eps,
            },
            {
                "range": "L2" + ":K" + str(last_row),
                "values": trailing_pe,
            },
            {
                "range": "M2" + ":M" + str(last_row),
                "values": dividend_rate,
            },
            {
                "range": "O2" + ":O" + str(last_row),
                "values": stock_sector,
            },
        ]
    )


main()
