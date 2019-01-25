import os
import datetime
from iexfinance.stocks import Stock
import pandas as pd
import xlsxwriter


def get_current_stock_price(ticker):
    print(ticker)
    price = Stock(ticker)
    stockPrice = price.get_price()
    return float(stockPrice)

def sort_by_exp_date(option):
    return option.exp_date

class __Stock:
    def __init__(self, ticker):
        self.price = get_current_stock_price(ticker)
        self.ticker = ticker
        self.call_spreads = {}
        self.call_spreads['itm'] = []
        self.call_spreads['otm'] = []
        self.put_spreads = {}
        self.put_spreads['itm'] = []
        self.put_spreads['otm'] = []

    def sort_spreads(self):
        self.call_spreads['itm'] = sorted(self.call_spreads['itm'], key=sort_by_exp_date)
        self.call_spreads['otm'] = sorted(self.call_spreads['otm'], key=sort_by_exp_date)
        self.put_spreads['itm'] = sorted(self.put_spreads['itm'], key=sort_by_exp_date)
        self.put_spreads['otm'] = sorted(self.put_spreads['otm'], key=sort_by_exp_date)



class Spread:
    def __init__(self, type, ticker, last, max_loss, max_proft, break_even, probability, exp_date, leg1_strike, leg2_strike):
        self.type = type
        self.ticker = ticker
        self.last = float(last)
        self.max_loss = float(max_loss)
        self.max_profit = float(max_proft)
        self.break_even = float(break_even)
        self.probability = probability
        self.exp_date = exp_date
        self.leg1_strike = float(leg1_strike)
        self.leg2_strike = float(leg2_strike)





stock_dict = {}
for file in os.listdir('spread_data'):
    df = pd.read_csv(f'spread_data/{file}')
    df.drop(df.tail(1).index, inplace=True)
    if 'call' in file:
        type = 'call_credit'
    else:
        type = 'put_credit'
    for i in range(1, len(df.index)):
        raw_exp_date_list = df['Exp Date'][i].split('/')
        exp_date = datetime.date(int(raw_exp_date_list[2]), int(raw_exp_date_list[0]), int(raw_exp_date_list[1]))
        spread = Spread(type, df['Symbol'][i], df['Last'][i], df['Max Loss'][i], df['Max Profit'][i], df['Break Even'][i], df['Probability'][i], exp_date, df['Leg1 Strike'][i], df['Leg2 Strike'][i])
        if not spread.ticker in stock_dict:
            stock_dict[spread.ticker] = __Stock(spread.ticker)
        if spread.type == 'call_credit':
            if spread.max_loss > spread.max_profit:
                stock_dict[spread.ticker].call_spreads['otm'].append(spread)
            else:
                stock_dict[spread.ticker].call_spreads['itm'].append(spread)
        else:
            if spread.max_loss > spread.max_profit:
                stock_dict[spread.ticker].put_spreads['otm'].append(spread)
            else:
                stock_dict[spread.ticker].put_spreads['itm'].append(spread)


for stock in stock_dict.values():
    stock.sort_spreads()

def sort_by_number_of_spreads(stock):
    num_of_spreads = 0
    num_of_spreads += len(stock.call_spreads['itm'])
    num_of_spreads += len(stock.call_spreads['otm'])
    num_of_spreads += len(stock.put_spreads['itm'])
    num_of_spreads += len(stock.put_spreads['otm'])
    return num_of_spreads

stock_list = sorted(stock_dict.values(), key=sort_by_number_of_spreads, reverse=True)

workbook = xlsxwriter.Workbook('spread_output/spread_output.xlsx')
worksheet = workbook.add_worksheet('Spread Data')

worksheet.write(0, 0, "Expiration Date")
worksheet.write(0, 1, "Premium")
worksheet.write(0, 2, "Max Loss")
worksheet.write(0, 3, "Max Profit")
worksheet.write(0, 4, "Break Even")
worksheet.write(0, 5, "Probability")
worksheet.write(0, 6, "Leg 1 Strike")
worksheet.write(0, 7, "Leg 2 Strike")
worksheet.write(0, 8, "Current Price")
row = 1

def write_option_data(option, row, stock_price):
    worksheet.write(row, 0, f'{option.exp_date.month}/{option.exp_date.day}/{option.exp_date.year}')
    worksheet.write(row, 1, f'${option.last:,.2f}')
    worksheet.write(row, 2, f'${option.max_loss:,.2f}')
    worksheet.write(row, 3, f'${option.max_profit:,.2f}')
    worksheet.write(row, 4, f'${option.break_even:,.2f}')
    worksheet.write(row, 5, option.probability)
    worksheet.write(row, 6, f'${option.leg1_strike:,.2f}')
    worksheet.write(row, 7, f'${option.leg2_strike:,.2f}')
    worksheet.write(row, 8, f'${stock_price:,.2f}')

for stock in stock_list:
    worksheet.write(row, 0 , stock.ticker)
    row += 1
    if len(stock.call_spreads['itm']) > 0:
        worksheet.write(row, 0, "ITM Call Credit Spread")
        worksheet.write(row, 1, len(stock.call_spreads['itm']))
        row += 1
        for option in stock.call_spreads['itm']:
            write_option_data(option, row, stock.price)
            row += 1
    if len(stock.call_spreads['otm']) > 0:
        worksheet.write(row, 0, "OTM Call Credit Spread")
        worksheet.write(row, 1, len(stock.call_spreads['otm']))
        row += 1
        for option in stock.call_spreads['otm']:
            write_option_data(option, row, stock.price)
            row += 1
    if len(stock.put_spreads['itm']) > 0:
        worksheet.write(row, 0, "ITM Put Credit Spread")
        worksheet.write(row, 1, len(stock.put_spreads['itm']))
        row += 1
        for option in stock.put_spreads['itm']:
            write_option_data(option, row, stock.price)
            row += 1
    if len(stock.put_spreads['otm']) > 0:
        worksheet.write(row, 0, "OTM Put Credit Spread")
        worksheet.write(row, 1, len(stock.put_spreads['otm']))
        row += 1
        for option in stock.put_spreads['otm']:
            write_option_data(option, row, stock.price)
            row += 1
    row += 1

workbook.close()