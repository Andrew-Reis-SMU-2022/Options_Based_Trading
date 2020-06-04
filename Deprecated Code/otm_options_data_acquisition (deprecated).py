import os
import csv
import datetime
import xlsxwriter
from iexfinance.stocks import Stock

workbook = xlsxwriter.Workbook('output/output.xlsx')

def write_to_formated_excel(stocks_list, worksheet_name):
    worksheet = workbook.add_worksheet(worksheet_name)

    worksheet.write(0, 0, "Ticker/Expiration Date")
    worksheet.write(0, 1, "Current Price")
    worksheet.write(0, 2, "Predicted Price")
    worksheet.write(0, 3, "Percent Change")
    worksheet.write(0, 4, "Total Volume")
    worksheet.write(0, 5, "Total Money Traded")
    worksheet.write(0, 6, "Number of Option Trades")
    worksheet.write(0, 7, "Calls")
    worksheet.write(0, 8, "Puts")

    row = 1

    for stock in stocks_list:
        worksheet.write(row, 0, stock.ticker)
        row += 1
        for exp_date_obj in stock.exp_dates:
            worksheet.write(row, 0,
                            f'{exp_date_obj.exp_date.month}/{exp_date_obj.exp_date.day}/{exp_date_obj.exp_date.year}')
            worksheet.write(row, 1, stock.price)
            worksheet.write(row, 2, '{:.2f}'.format(exp_date_obj.predicted_price))
            worksheet.write(row, 3, '{:.2f}%'.format(exp_date_obj.percent_change * 100))
            worksheet.write(row, 4, exp_date_obj.total_volume)
            worksheet.write(row, 5, '${:11,.2f}'.format(exp_date_obj.total_money_traded))
            worksheet.write(row, 6, exp_date_obj.num_of_options)
            worksheet.write(row, 7, exp_date_obj.num_of_calls)
            worksheet.write(row, 8, exp_date_obj.num_of_puts)
            row += 1
        row += 1


def write_to_raw_excel(stocks_list):
    worksheet = workbook.add_worksheet('Raw Data For Manipulation')

    worksheet.write(0, 0, "Ticker")
    worksheet.write(0, 1, "Current Price")
    worksheet.write(0, 2, "Expiration Date")
    worksheet.write(0, 3, "Predicted Price")
    worksheet.write(0, 4, "Percent Change")
    worksheet.write(0, 5, "Volume")
    worksheet.write(0, 6, "Total Money Traded")
    worksheet.write(0, 7, "Number of Option Trades")
    worksheet.write(0, 8, "Calls")
    worksheet.write(0, 9, "Puts")

    row = 1

    for stock in stocks_list:
        for exp_date_obj in stock.exp_dates:
            worksheet.write(row, 0, stock.ticker)
            worksheet.write(row, 1, stock.price)
            worksheet.write(row, 2,
                            f'{exp_date_obj.exp_date.month}/{exp_date_obj.exp_date.day}/{exp_date_obj.exp_date.year}')
            worksheet.write(row, 3, exp_date_obj.predicted_price)
            worksheet.write(row, 4, exp_date_obj.percent_change)
            worksheet.write(row, 5, exp_date_obj.total_volume)
            worksheet.write(row, 6, exp_date_obj.total_money_traded)
            worksheet.write(row, 7, exp_date_obj.num_of_options)
            worksheet.write(row, 8, exp_date_obj.num_of_calls)
            worksheet.write(row, 9, exp_date_obj.num_of_puts)
            row += 1


def get_current_stock_price(ticker):
    print(ticker)
    price = Stock(ticker)
    stockPrice = price.get_price()
    return float(stockPrice)


class ExpDate:
    def __init__(self, exp_date):
        self.exp_date = exp_date
        self.options = []
        self.num_of_calls = 0
        self.num_of_puts = 0
        self.num_of_options = 0

    def calc_predicted_price(self):
        # Weightd Average of strike price with volume
        numerator = 0
        denominator = 0
        for option in self.options:
            numerator += option.strike * option.volume
            denominator += option.volume

        self.predicted_price = numerator / denominator

    def calc_percent_change(self, current_stock_price):
        self.percent_change = (self.predicted_price - current_stock_price) / current_stock_price

    def calc_total_volume(self):
        self.total_volume = 0
        for option in self.options:
            self.total_volume += option.volume

    def calc_total_price(self):
        self.total_money_traded = 0
        for option in self.options:
            self.total_money_traded += option.total_cost


class __Stock:
    def __init__(self, ticker):
        self.ticker = ticker
        self.price = get_current_stock_price(ticker)
        self.exp_dates = []


class OptionTrade:
    def __init__(self, ticker, strike, exp_date, last, volume, type):
        self.ticker = ticker
        self.strike = strike
        self.exp_date = exp_date
        self.last = last
        self.volume = volume
        self.total_cost = last * volume * 100.0
        self.type = type


start = datetime.datetime.now()  # For Optimization Purposes

stocks_dict = {}
for file in os.listdir('data'):
    with open(f"data/{file}", 'r') as csv_file:
        csv_reader = csv.reader(csv_file)
        next(csv_reader)
        for line in csv_reader:
            if len(line) > 2:
                if (line[2].lower() == 'call' and float(line[3]) > float(line[1])) or (
                        line[2].lower() == 'put' and float(line[3]) < float(line[1])):  # for otm calls and puts
                    mdy_list = line[4].split('/')
                    date = datetime.date(int(mdy_list[2]) + 2000, int(mdy_list[0]), int(mdy_list[1])) # +2000 because years are just listed as 19 or 20 in the csv
                    if date > datetime.date.today():
                        option = OptionTrade(line[0], float(line[3]), date, float(line[9]), int(line[10]),
                                             line[2].lower())
                        if not option.ticker in stocks_dict:
                            stocks_dict[option.ticker] = __Stock(option.ticker)
                        list_exp_dates = []
                        for exp_date_obj in stocks_dict[option.ticker].exp_dates:
                            list_exp_dates.append(exp_date_obj.exp_date)
                        if not option.exp_date in list_exp_dates:
                            stocks_dict[option.ticker].exp_dates.append(ExpDate(option.exp_date))
                        for exp_date_obj in stocks_dict[option.ticker].exp_dates:
                            if option.exp_date == exp_date_obj.exp_date:
                                exp_date_obj.options.append(option)
                                if option.type == 'call':
                                    exp_date_obj.num_of_calls += 1
                                else:
                                    exp_date_obj.num_of_puts += 1
                                exp_date_obj.num_of_options += 1


def sorting_exp_dates(exp_date_obj):
    return exp_date_obj.exp_date


for key, stock in stocks_dict.items():
    for exp_date in stock.exp_dates:
        exp_date.calc_predicted_price()
        exp_date.calc_percent_change(stock.price)
        exp_date.calc_total_volume()
        exp_date.calc_total_price()

    stock.exp_dates = sorted(stock.exp_dates, key=sorting_exp_dates)


def sorting_stocks_percentage(stock_obj):
    return abs(stock_obj.exp_dates[0].percent_change)


stocks_sorted_list = sorted(stocks_dict.values(), key=sorting_stocks_percentage, reverse=True)
write_to_formated_excel(stocks_sorted_list, "Sorted by Percent Change")


def sorting_stocks_total_money(stock_obj):
    largest_total_money = 0
    for exp_date_obj in stock_obj.exp_dates:
        if exp_date_obj.total_money_traded > largest_total_money:
            largest_total_money = exp_date_obj.total_money_traded
    return largest_total_money


stocks_sorted_list = sorted(stocks_dict.values(), key=sorting_stocks_total_money, reverse=True)
write_to_formated_excel(stocks_sorted_list, "Sorted by Total Money Traded")

write_to_raw_excel(stocks_dict.values())

workbook.close()

end = datetime.datetime.now()
print(end - start)