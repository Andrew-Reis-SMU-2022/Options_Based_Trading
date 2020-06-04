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
    worksheet.write(0, 7, "Call Trades")
    worksheet.write(0, 8, "Call Volume")
    worksheet.write(0, 9, "Call Money")
    worksheet.write(0, 10, "Put Trades")
    worksheet.write(0, 11, "Put Volume")
    worksheet.write(0, 12, "Put Money")

    row = 1

    for stock in stocks_list:
        worksheet.write(row, 0, f'${stock.ticker}')
        if stock.cvpv != None:
            worksheet.write(row, 1, "CVPV Ratio:")
            worksheet.write(row, 2, f'{stock.cvpv:.3f}')
        if stock.cmpm != None:
            worksheet.write(row, 3, "CMPM Ratio:")
            worksheet.write(row, 4, f'{stock.cmpm:.3f}')
        row += 1
        for exp_date_obj in stock.exp_dates:
            worksheet.write(row, 0,
                            f'{exp_date_obj.exp_date.month}/{exp_date_obj.exp_date.day}/{exp_date_obj.exp_date.year}')
            worksheet.write(row, 1, '${:.2f}'.format(stock.price))
            worksheet.write(row, 2, '${:.2f}'.format(exp_date_obj.predicted_price))
            worksheet.write(row, 3, '{:.2f}%'.format(exp_date_obj.percent_change * 100))
            worksheet.write(row, 4, exp_date_obj.total_volume)
            worksheet.write(row, 5, '${:11,.2f}'.format(exp_date_obj.total_money_traded))
            worksheet.write(row, 6, exp_date_obj.num_of_options)
            worksheet.write(row, 7, exp_date_obj.num_of_calls)
            worksheet.write(row, 8, exp_date_obj.call_volume)
            worksheet.write(row, 9, '${:11,.2f}'.format(exp_date_obj.money_in_calls))
            worksheet.write(row, 10, exp_date_obj.num_of_puts)
            worksheet.write(row, 11, exp_date_obj.put_volume)
            worksheet.write(row, 12, '${:11,.2f}'.format(exp_date_obj.money_in_puts))
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
    worksheet.write(0, 8, "Call Trades")
    worksheet.write(0, 9, "Call Volume")
    worksheet.write(0, 10, "Call Money")
    worksheet.write(0, 11, "Put Trades")
    worksheet.write(0, 12, "Put Volume")
    worksheet.write(0, 13, "Put Money")

    row = 1

    for stock in stocks_list:
        for exp_date_obj in stock.exp_dates:
            worksheet.write(row, 0, f'${stock.ticker}')
            worksheet.write(row, 1, stock.price)
            worksheet.write(row, 2,
                            f'{exp_date_obj.exp_date.month}/{exp_date_obj.exp_date.day}/{exp_date_obj.exp_date.year}')
            worksheet.write(row, 3, exp_date_obj.predicted_price)
            worksheet.write(row, 4, exp_date_obj.percent_change)
            worksheet.write(row, 5, exp_date_obj.total_volume)
            worksheet.write(row, 6, exp_date_obj.total_money_traded)
            worksheet.write(row, 7, exp_date_obj.num_of_options)
            worksheet.write(row, 8, exp_date_obj.num_of_calls)
            worksheet.write(row, 9, exp_date_obj.call_volume)
            worksheet.write(row, 10, exp_date_obj.money_in_calls)
            worksheet.write(row, 11, exp_date_obj.num_of_puts)
            worksheet.write(row, 12, exp_date_obj.put_volume)
            worksheet.write(row, 13, exp_date_obj.money_in_puts)
            row += 1


def write_to_individual_trades(stocks_list):
    worksheet = workbook.add_worksheet("Individual Option Trades")
    worksheet.write(0, 0, "Ticker/Expiration Date")
    worksheet.write(0, 1, "Stock Price")
    worksheet.write(0, 2, "Strike Price")
    worksheet.write(0, 3, "Premium")
    worksheet.write(0, 4, "Break Even Price")
    worksheet.write(0, 5, "Volume")
    worksheet.write(0, 6, "Total Price Paid")

    row = 1

    for stock in stocks_list:
        worksheet.write(row, 0, f'${stock.ticker}')
        row += 1

        if len(stock.list_of_calls) > 0:
            worksheet.write(row, 0, "Calls:")
            row += 1
            for option in stock.list_of_calls:
                worksheet.write(row, 0, f'{option.exp_date.month}/{option.exp_date.day}/{option.exp_date.year}')
                worksheet.write(row, 1, '${:.2f}'.format(stock.price))
                worksheet.write(row, 2, '${:.2f}'.format(option.strike))
                worksheet.write(row, 3, '${:.2f}'.format(option.last))
                worksheet.write(row, 4, '${:.2f}'.format(option.break_even))
                worksheet.write(row, 5, option.volume)
                worksheet.write(row, 6, '${:,.2f}'.format(option.total_cost))
                row += 1

        if len(stock.list_of_puts) > 0:
            worksheet.write(row, 0, "Puts:")
            row += 1
            for option in stock.list_of_puts:
                worksheet.write(row, 0, f'{option.exp_date.month}/{option.exp_date.day}/{option.exp_date.year}')
                worksheet.write(row, 1, '${:.2f}'.format(stock.price))
                worksheet.write(row, 2, '${:.2f}'.format(option.strike))
                worksheet.write(row, 3, '${:.2f}'.format(option.last))
                worksheet.write(row, 4, '${:.2f}'.format(option.break_even))
                worksheet.write(row, 5, option.volume)
                worksheet.write(row, 6, '${:,.2f}'.format(option.total_cost))
                row += 1
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
        # Weightd Average of break even with volume
        numerator = 0
        denominator = 0
        for option in self.options:
            numerator += option.break_even * option.volume
            denominator += option.volume

        self.predicted_price = numerator / denominator

    def calc_percent_change(self, current_stock_price):
        self.percent_change = (self.predicted_price - current_stock_price) / current_stock_price

    def calc_total_volumes(self):
        self.total_volume = 0
        self.call_volume = 0
        self.put_volume = 0
        for option in self.options:
            self.total_volume += option.volume
            if option.type == 'call':
                self.call_volume += option.volume
            elif option.type == 'put':
                self.put_volume += option.volume


    def calc_total_price(self):
        self.total_money_traded = 0
        for option in self.options:
            self.total_money_traded += option.total_cost

    def calc_put_money(self):
        self.money_in_puts = 0
        for option in self.options:
            if option.type == 'put':
                self.money_in_puts += option.total_cost

    def calc_call_money(self):
        self.money_in_calls = 0
        for option in self.options:
            if option.type == 'call':
                self.money_in_calls += option.total_cost


def option_list_sorter(option):
    return option.exp_date

class __Stock:
    def __init__(self, ticker):
        self.ticker = ticker
        self.price = get_current_stock_price(ticker)
        self.exp_dates = []

    def construct_option_list(self):
        self.list_of_calls = []
        self.list_of_puts = []
        for exp_date_obj in self.exp_dates:
            for option in exp_date_obj.options:
                if option.type == 'call':
                    self.list_of_calls.append(option)
                elif option.type == 'put':
                    self.list_of_puts.append(option)
        self.list_of_calls = sorted(self.list_of_calls, key=option_list_sorter)
        self.list_of_puts = sorted(self.list_of_puts, key=option_list_sorter)

    def calc_ratios(self):
        total_call_volume = 0
        total_put_volume = 0
        total_call_money = 0
        total_put_money = 0
        for exp_date_obj in self.exp_dates:
            total_call_volume += exp_date_obj.call_volume
            total_put_volume += exp_date_obj.put_volume
            total_call_money += exp_date_obj.money_in_calls
            total_put_money += exp_date_obj.money_in_puts

        if total_put_volume == 0 or total_call_volume == 0:
            self.cvpv = None
        else:
            self.cvpv = total_call_volume / total_put_volume
        if total_put_money == 0 or total_call_money == 0:
            self.cmpm = None
        else:
            self.cmpm = total_call_money / total_put_money




class OptionTrade:
    def __init__(self, ticker, strike, exp_date, last, volume, type):
        self.ticker = ticker
        self.strike = strike
        self.exp_date = exp_date
        self.last = last
        self.volume = volume
        self.total_cost = last * volume * 100.0
        self.type = type
        if self.type == 'call':
            self.break_even = strike + last
        if self.type == 'put':
            self.break_even = strike - last



start = datetime.datetime.now()  # For Optimization Purposes

stocks_dict = {}
for file in os.listdir('data'):
    with open(f"data/{file}", 'r') as csv_file:
        csv_reader = csv.reader(csv_file)
        next(csv_reader)
        for line in csv_reader:
            if len(line) > 2:
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
        exp_date.calc_total_volumes()
        exp_date.calc_total_price()
        exp_date.calc_call_money()
        exp_date.calc_put_money()
    stock.calc_ratios()
    stock.construct_option_list()

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


def alphabetical_order(stock):
    return stock.ticker

stocks_sorted_list = sorted(stocks_dict.values(), key=alphabetical_order)
write_to_individual_trades(stocks_sorted_list)

workbook.close()

end = datetime.datetime.now()
print(end - start)