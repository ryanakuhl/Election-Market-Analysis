import os
import csv
import time, random
import datetime
from datetime import timedelta
import xlsxwriter
import iexfinance
from iexfinance import Stock
import pandas as pd
import pandas_datareader as pdr
from datetime import datetime
from datetime import timedelta


election_dict = {}
election_years = []
all_yahoo_data = {}
stock_symbols = ['SPY', 'NASDAQ', 'NYSE', 'DOW']


def write_to_formatted_excel():
    row = 0
    df1 = pd.DataFrame(columns=['SPY', 'NASDAQ', 'NYSE', 'DOW'])
    data = {'SPY': [0, 0, 0, 0, 0],
            'NASDAQ': [0, 0, 0, 0, 0],
            'NYSE': [0, 0, 0, 0, 0],
            'DOW': [0, 0, 0, 0, 0],
            }
    df3 = pd.DataFrame(data)
    writer_object = pd.ExcelWriter('pandas_line_chart.xlsx', engine='xlsxwriter')
    df1.to_excel(writer_object, sheet_name='PercentChange')
    df1.to_excel(writer_object, sheet_name='ElectionData')
    df3.to_excel(writer_object, sheet_name='AverageChange')
    workbook_object = writer_object.book
    second_sheet = writer_object.sheets['ElectionData']
    second_sheet.set_column('A:E', 20)

    times_added = 0
    for i in range(1,len(lets_process.presidents)):
        second_sheet.write(row, 0, "Election Number")
        second_sheet.write(row, 1, "House Dem")
        second_sheet.write(row, 2, "House Rep")
        second_sheet.write(row, 3, "Senate Dem")
        second_sheet.write(row, 4, "Senate Rep")
        row += 1
        second_sheet.write(row, 0, lets_process.this_year[i])
        row += 1
        second_sheet.write(row, 0, lets_process.presidents[i])
        second_sheet.write(row, 1, lets_process.house_dems[i])
        second_sheet.write(row, 2, lets_process.house_repubs[i])
        second_sheet.write(row, 3, lets_process.senate_dems[i])
        second_sheet.write(row, 4, lets_process.senate_repubs[i])
        row += 1
        if lets_process.house_held[i] < 1.00:
            second_sheet.write(row, 1, 'Republican Majority')
            second_sheet.write(row, 2, (1 - lets_process.house_repubs[i] / lets_process.house_dems[i]) * -100)
        else:
            second_sheet.write(row, 1, 'Democrat Majority')
            second_sheet.write(row, 2, (1 - lets_process.house_dems[i] / lets_process.house_repubs[i]) * -100)
        if lets_process.senate_held[i] < 1.00:
            second_sheet.write(row, 3, 'Republican Majority')
            second_sheet.write(row, 4, (1 - lets_process.senate_repubs[i] / lets_process.senate_dems[i]) * -100)
        else:
            second_sheet.write(row, 3, 'Democrat Majority')
            second_sheet.write(row, 4, (1 - lets_process.senate_dems[i] / lets_process.senate_repubs[i]) * -100)
        row += 2
        df2 = pd.DataFrame(columns=['SPY', 'NASDAQ', 'NYSE', 'DOW'])
        stats = [lets_process.this_year[i] + '_' + s for s in stock_symbols if lets_process.this_year[i] + '_' + s in all_yahoo_data.keys()]
        for s in stats:
            single_stock = pd.DataFrame(all_yahoo_data.get(s).Close / all_yahoo_data.get(s).Open, columns=[s.split('_')[1]])
            df2[s.split('_')[1]] = single_stock[s.split('_')[1]]
            df3[getattr(df3, s.split('_')[1]).name] += single_stock[getattr(df3, s.split('_')[1]).name].values
            times_added += .25
        if df2.values.any():
            df1 = df1.append(df2)
    row = 0
    worksheet_object = writer_object.sheets['PercentChange']
    worksheet_object.set_column('A:E', 20)
    df1.to_excel(writer_object, sheet_name='PercentChange', startrow=row)
    worksheet_object = writer_object.sheets['AverageChange']
    df3 = df3 / times_added
    df3.to_excel(writer_object, sheet_name='AverageChange', startrow=row)
    chart_object = workbook_object.add_chart({'type': 'line'})
    newlist = [round(float(x), 3) for x in list(df3['SPY'])]
    newlist.insert(0, 'AverageChange')
    chart_object.add_series({
        'name': '=AverageChange!$B$1',
        'categories': '=AverageChange!$A$2:$A$6',
        'values': '=AverageChange!$B$2:$B$6',
         })
    chart_object.add_series({
        'name': '=AverageChange!$C$1',
        'categories': '=AverageChange!$A$2:$A$6',
        'values': '=AverageChange!$C$2:$C$6',
         })
    chart_object.add_series({
        'name': '=AverageChange!$D$1',
        'categories': '=AverageChange!$A$2:$A$6',
        'values': '=AverageChange!$D$2:$D$6',
        })
    chart_object.add_series({
        'name': '=AverageChange!$E$1',
        'categories': '=AverageChange!$A$2:$A$6',
        'values': '=AverageChange!$E$2:$E$6',
        })
    chart_object.set_title({'name': 'Combined Mean'})
    chart_object.set_x_axis({'name': 'Markets'})
    chart_object.set_y_axis({'name': '% Change'})
    worksheet_object.insert_chart('G2', chart_object, {'x_offset': 20, 'y_offset': 0})
    writer_object.save()


class ProcessData:

    def __init__(self):
        self.number_of_elections = 0
        self.elections = []
        self.election_values = []
        self.election_keys = []
        self.house_dems = []
        self.house_repubs = []
        self.senate_repubs = []
        self.senate_dems = []
        self.presidents = []
        self.this_year = []
        self.congress_date = []
        self.house_held = []
        self.senate_held = []
        self.stock_market = []

    def all_elections(self):
        self.elections += election_dict
        self.election_values += election_dict.values()
        self.election_keys += election_dict.keys()
        for each_election in self.elections:
            self.this_year += [each_election]
            self.house_dems += [election_dict.get(each_election).house_dem]
            self.house_repubs += [election_dict.get(each_election).house_repub]
            self.senate_repubs += [election_dict.get(each_election).senate_repub]
            self.senate_dems += [election_dict.get(each_election).senate_dem]
            self.presidents += [election_dict.get(each_election).president]
            self.congress_date += [election_dict.get(each_election).congress_date]
            self.house_held += [election_dict.get(each_election).house_held]
            self.senate_held += [election_dict.get(each_election).senate_held]


class USA:
    def __init__(self, president, president_election, president_party, presidential_date, congress, house_total, house_dem, house_repub, senate_total, senate_dem, senate_repub, congress_date):

        self.house_dem = ''
        self.house_repub = ''
        self.senate_repub = ''
        self.senate_dem = ''
        self.president = president#Col B
        self.president_election = president_election #Col A
        self.president_party = president_party#Col C
        self.president_date = presidential_date #Col D if Y/N
        self.congress_date = congress_date #Col G
        self.congress = congress #Col F
        self.house_total = house_total #Col U
        if len(house_dem) > 1:
            self.house_dem = int(house_dem) #Col V
        else:
            self.house_dem = 1
        if len(house_repub) > 1:
            self.house_repub = int(house_repub) #Col W
        else:
            self.house_repub = 1
        self.senate_total = senate_total #Col P
        if len(senate_dem) > 1:
            self.senate_dem = int(senate_dem) #Col Q
        else:
            self.senate_dem = 1
        if len(senate_repub) > 1:
            self.senate_repub = int(senate_repub) #Col R
        else:
            self.senate_repub = 1
        self.house_held = self.house_dem / self.house_repub
        self.senate_held = self.senate_dem / self.senate_repub
        self.week_start = ''
        self.week_end = ''

    def start_of_week(self):
        self.week_start = ''
        if len(self.congress_date) > 3:
            hopefully_tuesday = datetime.strptime(self.congress_date, '%B %d, %Y')
            beginning_of_week = hopefully_tuesday.weekday()
            self.week_start = hopefully_tuesday - timedelta(beginning_of_week)
            self.week_end = self.week_start + timedelta(4)

    def get_stock_week_of(self):
        stock_market = {
                        'SPY': ('^GSPC','1990-1-2'),#goes/ went back to 70
                        'NASDAQ': ('^IXIC','1990-2-1'),#goes/went back to 71
                        'DOW': ('^DJI','1990-1-29'),#used to go back to 85
                        'NYSE': ('^NYA','1990-1-2')#goes/went back to 70
                        }
        within_stock_market_range = [key for key in stock_market if len(self.congress_date) > 3 and self.week_start > datetime.strptime(stock_market.get(key)[1], '%Y-%m-%d')]
        for stock in within_stock_market_range:
            stock_val = pdr.get_data_yahoo(stock_market.get(stock)[0], start=self.week_start.strftime('%Y-%m-%d'),end=self.week_end.strftime('%Y-%m-%d'))
            all_yahoo_data.update({self.congress+'_'+stock : stock_val})
            time.sleep(2)


with open('elections.csv', 'r') as csv_file:
    csv_reader = csv.reader(csv_file)
    next(csv_reader)
    for line in csv_reader:
        if len(line) > 2:
            election = USA(line[1],line[0],line[2],line[3],line[5],line[20],line[21],line[22],line[15],line[16],line[17],line[6])
            election.start_of_week()
            election.get_stock_week_of()
            if election.president_date not in election_dict:
                election_dict[election.congress] = election
    lets_process = ProcessData()
    lets_process.all_elections()
    # collect object dataframes into one for analysis
    write_to_formatted_excel()
csv_file.close()


"""

Next:

Start learning charts to create data cuts

"""

