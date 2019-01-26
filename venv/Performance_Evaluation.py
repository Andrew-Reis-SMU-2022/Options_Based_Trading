import os
import datetime
from iexfinance.stocks import get_historical_data
import pandas as pd
import xlsxwriter


def get_stock_price(ticker, historical_date):
    print(ticker)
    df = get_historical_data(ticker, historical_date)
    return float(df[f'{historical_date.year}-{historical_date.month:02}-{historical_date.day:02}']['close'])


class __Stock:
    def __init__(self, ticker):
        self.ticker = ticker
        self.price = get_stock_price(ticker, testing_date)
        self.options = {}
        self.options['calls'] = []
        self.options['puts'] = []

    def calc_total_monies(self):
        self.put_money = 0
        self.call_money = 0
        for option in self.options['calls']:
            self.call_money += option.price
        for option in self.options['puts']:
            self.put_money += option.price
        self.total_money = self.put_money + self.call_money
        if self.put_money > self.call_money:
            self.consensus_type = 'Put'
        elif self.call_money > self.put_money:
            self.consensus_type = 'Call'
        else:
            self.consensus_type = None

    def calc_consensus_weighted_average_strike(self):
        numerator = 0
        denominator = 0
        if self.consensus_type == 'Call':
            for option in self.options['calls']:
                numerator += option.strike * option.volume
                denominator += option.volume
        else:
            for option in self.options['puts']:
                numerator += option.strike * option.volume
                denominator += option.volume
        self.WA_strike = numerator / denominator

    def calc_total_volumes(self):
        self.put_volume = 0
        self.call_volume = 0
        for option in self.options['calls']:
            self.call_volume += option.volume
        for option in self.options['puts']:
            self.put_volume += option.volume
        self.total_volume = self.put_volume + self.call_volume

    def calc_profit_of_consensus(self):
        self.consensus_profit = 0
        if self.call_money > self.put_money:
            for option in self.options['calls']:
                self.consensus_profit += option.volume * 100.0 * (self.price - option.strike) - option.price
            if self.consensus_profit < -self.call_money:
                self.consensus_profit = -self.call_money
        elif self.put_money > self.call_money:
            for option in self.options['puts']:
                self.consensus_profit += option.volume * 100.0 * (option.strike - self.price) - option.price
            if self.consensus_profit < -self.put_money:
                self.consensus_profit = -self.put_money
        else:
            self.consensus_profit = None

    def calc_WA_stock_purchase_price(self):
        numerator = 0
        denominator = 0
        if self.call_money > self.put_money:
            for option in self.options['calls']:
                numerator += option.stock_purchase_price * option.volume
                denominator += option.volume
        else:
            for option in self.options['puts']:
                numerator += option.stock_purchase_price * option.volume
                denominator += option.volume
        self.WA_stock_beginning_price = numerator / denominator



class Option:
    def __init__(self, ticker, type, strike, premium, volume, purchase_date):
        self.ticker = ticker
        self.type = type
        self.strike = strike
        self.premium = premium
        self.volume = volume
        self.price = volume * premium * 100.0
        self.purchase_date = purchase_date
        self.stock_purchase_price = get_stock_price(ticker, purchase_date)




raw_input = input("Enter Expiration Date: ")
raw_splits = raw_input.split('/')
testing_date = datetime.date(int(raw_splits[2]), int(raw_splits[0]), int(raw_splits[1]))

stock_dict = {}
for file in os.listdir('data_for_validation'):
    df = pd.read_csv(f'data_for_validation/{file}')
    df.drop(df.tail(1).index, inplace=True)
    raw_buy_date_splits = file.split('-')[4:7]
    buy_date = datetime.date(int(raw_buy_date_splits[2].split('.')[0]), int(raw_buy_date_splits[0]), int(raw_buy_date_splits[1]))
    for i in range(1, len(df.index)):
        raw_exp_splits = df['Exp Date'][i].split('/')
        exp_date = datetime.date(int(raw_exp_splits[2]) + 2000, int(raw_exp_splits[0]), int(raw_exp_splits[1]))
        if testing_date == exp_date:
            option = Option(df['Symbol'][i], df['Type'][i].lower(), float(df['Strike'][i]), float(df['Last'][i]), int(df['Volume'][i]), buy_date)
            if not option.ticker in stock_dict:
                stock_dict[option.ticker] = __Stock(option.ticker)
            if option.type == 'call':
                stock_dict[option.ticker].options['calls'].append(option)
            else:
                stock_dict[option.ticker].options['puts'].append(option)


for stock in stock_dict.values():
    stock.calc_total_monies()
    stock.calc_total_volumes()
    stock.calc_consensus_weighted_average_strike()
    stock.calc_profit_of_consensus()
    stock.calc_WA_stock_purchase_price()


consensus_stocks = [stock for stock in stock_dict.values() if stock.total_money > 1000000.0 and
                    ((stock.call_money > 3.0/4.0 * stock.total_money and stock.call_volume > 0.5 * stock.total_volume)
                     or (stock.put_money > 3.0/4.0 * stock.total_money and stock.put_volume > 0.5 * stock.total_volume))]

def sort_by_consensus_money(stock):
    if stock.call_money > stock.put_money:
        return stock.call_money
    else:
        return stock.put_money

consensus_stocks = sorted(consensus_stocks, key=sort_by_consensus_money, reverse=True)

out_dict = {'Ticker': [], 'Type': [], 'WA Beginning Stock Price': [], 'Final Stock Price': [], 'Percent Chagne': [],
            'WA Strike': [], 'Consensus Volume': [], 'Initial Investment': [], 'Profit': [], 'ROI': []}

for stock in consensus_stocks:
    out_dict['Ticker'].append(stock.ticker)
    out_dict['Type'].append(stock.consensus_type)
    out_dict['WA Beginning Stock Price'].append(stock.WA_stock_beginning_price)
    out_dict['Final Stock Price'].append(stock.price)
    out_dict['Percent Change'].append((stock.price - stock.WA_stock_beginning_price) / stock.WA_stock_beginning_price)
    out_dict['WA Strike'].append(stock.WA_strike)
    if stock.consensus_type == 'Call':
        out_dict['Consensus Volume'].append(stock.call_volume)
        out_dict['Initial Investment'].append(stock.call_money)
    else:
        out_dict['Consensus Volume'].append(stock.put_volume)
        out_dict['Initial Investment'].append(stock.put_money)
    out_dict['Profit'].append(stock.consensus_profit)
    out_dict['ROI'].append( out_dict['Profit'][-1] / out_dict['Initial Investment'][-1])

output_df = pd.DataFrame(out_dict)
writer = pd.ExcelWriter('output/output.xlsx', engine='xlsxwriter')
output_df.to_excel(writer, sheet_name='Performance Evaluation')
writer.save()



