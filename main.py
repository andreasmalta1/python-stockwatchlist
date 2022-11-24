import yfinance as yf
import gspread

def main():

    sa = gspread.service_account(filename='service_account.json')
    sh = sa.open('Stocks Watchlist')
    whs = sh.worksheet('Sheet1')
    
    ticker_list = whs.col_values(1)
    starting_row = 3

    for k in range(0, 3):
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
    
    

        for i in range(starting_row, len(ticker_list) + 1):
            if ticker_list[i-1] == 'MoneyBase' or ticker_list[i-1] == 'Watchlist':
                break
            if ticker_list[i-1] == 'Ticker' or ticker_list[i-1] == 'EToro':
                continue
            information['ticker'].append([yf.Ticker(ticker_list[i-1])])

        for ticker in information['ticker']:
            ticker = ticker[0]
            for key in information:
                if key == 'ticker':
                    continue
                try:
                    information[key].append([ticker.info[key]])
                except KeyError:
                    information[key].append([""])

        if k == 2:
            i += 1

        whs.batch_update([
            {'range': 'C' + str(starting_row) + ':C' + str(i-1),
            'values': information['regularMarketPrice']},
            {'range': 'D' + str(starting_row) + ':D' + str(i-1),
            'values': information['fiftyTwoWeekLow']},
            {'range': 'E' + str(starting_row) + ':E' + str(i-1),
            'values': information['fiftyTwoWeekHigh']},
            {'range': 'G' + str(starting_row) + ':G' + str(i-1),
            'values': information['marketCap']},
            {'range': 'H' + str(starting_row) + ':H' + str(i-1),
            'values': information['totalRevenue']},
            {'range': 'I' + str(starting_row) + ':I' + str(i-1),
            'values': information['netIncomeToCommon']},
            {'range': 'J' + str(starting_row) + ':J' + str(i-1),
            'values': information['trailingEps']},
            {'range': 'L' + str(starting_row) + ':L' + str(i-1),
            'values': information['dividendRate']},
            {'range': 'N' + str(starting_row) + ':N' + str(i-1),
            'values': information['sector']}
        ])

main()
