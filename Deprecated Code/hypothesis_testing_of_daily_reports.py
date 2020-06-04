import datetime
import os
import xlsxwriter
import xlrd
from iexfinance.stocks import Stock
from iexfinance.stocks import get_historical_data


def get_original_stock_price(ticker, date):
    print(ticker)
    df = get_historical_data(ticker, date)
    return float(df[f'{date.year}-{date.month:02}-{date.day:02}']['close'])

def get_current_price(ticker):
    print(ticker)
    price = Stock(ticker)
    stockPrice = price.get_price()
    return float(stockPrice)

class SPY:
    def __init__(self, purchase_date):
        self.ticker = 'SPY'
        self.original_price = get_original_stock_price(self.ticker, purchase_date)
        self.current_price = get_current_price(self.ticker)
        self.return_percentage = (self.current_price - self.original_price) / self.original_price




class Option:
    def __init__(self, ticker, type, exp_date, total_price, purchase_date):
        self.ticker = ticker
        self.type = type
        self.exp_date = exp_date
        self.total_price = total_price
        self.purchase_date = purchase_date
        self.original_stock_price = get_original_stock_price(ticker, purchase_date)
        self.current_stock_price = get_current_price(ticker)
        self.return_percentage = (self.current_stock_price - self.original_stock_price) / self.original_stock_price


options_list = []
for file in os.listdir('data_for_validation'):
    read_workbook = xlrd.open_workbook(f'data_for_validation/{file}')
    read_worksheet = read_workbook.sheet_by_index(3)

    raw_purchase_date = file.split(' ')[4]
    purchase_date_splits = raw_purchase_date.split('.')
    purchase_date = datetime.date(int(purchase_date_splits[2]), int(purchase_date_splits[0]), int(purchase_date_splits[1]))

    for row in range(1, read_worksheet.nrows):
        raw_exp_date = read_worksheet.cell_value(row, 2)
        exp_date_splits = raw_exp_date.split('/')
        exp_date = datetime.date(int(exp_date_splits[2]), int(exp_date_splits[0]), int(exp_date_splits[1]))
        options_list.append(Option(read_worksheet.cell_value(row, 0).replace('$', ''), read_worksheet.cell_value(row, 1).lower(),
                                   exp_date, float(read_worksheet.cell_value(row, 8)), purchase_date))


list_of_calls = [option for option in options_list if option.type == 'call']
list_of_puts = [option for option in options_list if option.type == 'put']

def sort_by_exp_date(option):
    return option.exp_date

list_of_calls = sorted(list_of_calls, key=sort_by_exp_date)
list_of_puts = sorted(list_of_puts, key=sort_by_exp_date)

spy = SPY(purchase_date)

def WA_of_option_list(option_list):
    numerator = 0
    denominator = 0
    for option in option_list:
        numerator += option.return_percentage * option.total_price
        denominator += option.total_price
    return numerator / denominator

def simple_average(option_list):
    numerator = 0
    denominator = 0
    for option in option_list:
        numerator += option.return_percentage
        denominator += 1
    return numerator / denominator

write_workbook = xlsxwriter.Workbook('output/output.xlsx')
sheet_name = f'Data for {purchase_date.month}.{purchase_date.day}.{purchase_date.year}'
write_worksheet = write_workbook.add_worksheet(sheet_name)

write_worksheet.write(0, 0, "Calls:")

row = 1
write_worksheet.write(row, 0, "Ticker")
write_worksheet.write(row, 1, "Expiration Date")
write_worksheet.write(row, 2, "Cost of Trade")
write_worksheet.write(row, 3, "Original Stock Price")
write_worksheet.write(row, 4, "Current Stock Price")
write_worksheet.write(row, 5, "Percent Change")
row += 1

for option in list_of_calls:
    write_worksheet.write(row, 0, option.ticker)
    write_worksheet.write(row, 1, f'{option.exp_date.month}/{option.exp_date.day}/{option.exp_date.year}')
    write_worksheet.write(row, 2, option.total_price)
    write_worksheet.write(row, 3, option.original_stock_price)
    write_worksheet.write(row, 4, option.current_stock_price)
    write_worksheet.write(row, 5, option.return_percentage)
    row += 1

write_worksheet.write(row, 0, "Simple Average")
write_worksheet.write(row, 5, simple_average(list_of_calls))
row += 1
write_worksheet.write(row, 0, "Weighted Average")
write_worksheet.write(row, 5, WA_of_option_list(list_of_calls))
row += 1

write_worksheet.write(row, 0, "SPY")
write_worksheet.write(row, 3, spy.original_price)
write_worksheet.write(row, 4, spy.current_price)
write_worksheet.write(row, 5, spy.return_percentage)
row += 2

write_worksheet.write(row, 0, "Puts:")
row += 1

write_worksheet.write(row, 0, "Ticker")
write_worksheet.write(row, 1, "Expiration Date")
write_worksheet.write(row, 2, "Cost of Trade")
write_worksheet.write(row, 3, "Original Stock Price")
write_worksheet.write(row, 4, "Current Stock Price")
write_worksheet.write(row, 5, "Percent Change")
row += 1

for option in list_of_puts:
    write_worksheet.write(row, 0, option.ticker)
    write_worksheet.write(row, 1, f'{option.exp_date.month}/{option.exp_date.day}/{option.exp_date.year}')
    write_worksheet.write(row, 2, option.total_price)
    write_worksheet.write(row, 3, option.original_stock_price)
    write_worksheet.write(row, 4, option.current_stock_price)
    write_worksheet.write(row, 5, option.return_percentage)
    row += 1

write_worksheet.write(row, 0, "Simple Average")
write_worksheet.write(row, 5, simple_average(list_of_puts))
row += 1
write_worksheet.write(row, 0, "Weighted Average")
write_worksheet.write(row, 5, WA_of_option_list(list_of_puts))
row += 1

write_worksheet.write(row, 0, "SPY")
write_worksheet.write(row, 3, spy.original_price)
write_worksheet.write(row, 4, spy.current_price)
write_worksheet.write(row, 5, spy.return_percentage)

write_workbook.close()