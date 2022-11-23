import yfinance as yf
import gspread

def main():

    # Make for all sections

    sa = gspread.service_account(filename='service_account.json')
    sh = sa.open('Stocks Watchlist Updated')
    whs = sh.worksheet('Sheet1')
    
    ticker_list = whs.col_values(1)

    information = {
        'ticker': [],
        'regularMarketPrice': [],
        'fiftyTwoWeekLow': [],
        'fiftyTwoWeekHigh': [],
        'marketCap': [],
        'totalRevenue': [],
        'netIncomeToCommon': [],
        'trailingEps': [],
        'dividendRate': [],
        'sector': []}
    
    moneybase_tickers = []
    watchlist_tickers = []

    for etoro_counter, ticker in enumerate(ticker_list):
        if ticker == 'MoneyBase':
            break
        if ticker == 'Ticker' or ticker == 'EToro':
            continue
        information['ticker'].append([yf.Ticker(ticker)])

    for ticker in information['ticker']:
        ticker = ticker[0]
        for key in information:
            if key == 'ticker':
                continue
            try:
                information[key].append([ticker.info[key]])
            except KeyError:
                information[key].append([""])

    whs.batch_update([
        {'range': 'C3:C' + str(etoro_counter),
        'values': information['regularMarketPrice']},
        {'range': 'D3:D' + str(etoro_counter),
        'values': information['fiftyTwoWeekLow']},
        {'range': 'E3:E' + str(etoro_counter),
        'values': information['fiftyTwoWeekHigh']},
        {'range': 'G3:G' + str(etoro_counter),
        'values': information['marketCap']},
        {'range': 'I3:I' + str(etoro_counter),
        'values': information['totalRevenue']},
        {'range': 'J3:J' + str(etoro_counter),
        'values': information['netIncomeToCommon']},
        {'range': 'K3:K' + str(etoro_counter),
        'values': information['trailingEps']},
        {'range': 'P3:P' + str(etoro_counter),
        'values': information['dividendRate']},
        {'range': 'R3:R' + str(etoro_counter),
        'values': information['sector']}
    ])

main()
