import numpy as np
import pandas as pd
import math, requests, openpyxl
import equal_weights as ew
from secrets1 import IEX_CLOUD_API_TOKEN
from statistics import mean, median


def batch_requests(rfpath, fd):
    stocks = pd.read_csv(rfpath)
    symbol_grp = list(ew.chunks(stocks['Ticker'], 100))
    symbol_strings = []
    for grp in symbol_grp:
        symbol_strings.append(",".join(grp))
    for ss in symbol_strings:
        batch_api_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={ss}&types=stats,price&token={IEX_CLOUD_API_TOKEN}'
        data = requests.get(batch_api_url).json()
        for symbol in ss.split(","):
            fd.loc[len(fd)] = [symbol, data[symbol]['price'], 'N/A',
                                                data[symbol]['stats']['year1ChangePercent'], 'N/A',
                                                data[symbol]['stats']['month6ChangePercent'], 'N/A',
                                                data[symbol]['stats']['month3ChangePercent'], 'N/A',
                                                data[symbol]['stats']['month1ChangePercent'], 'N/A'
                                                ]
    time_periods = ['One-Year', 'Six-Month', 'Three-Month', 'One-Month']
    for time_period in time_periods:
        each_col = f'{time_period} Return Percentile'
        calc_col = f'{time_period} Price Return'
        fd[each_col] = (fd[calc_col].rank()/len(fd.index))


def mean_median_momentum(fd, request_type):
    time_periods = ['One-Year', 'Six-Month', 'Three-Month', 'One-Month']
    for row in fd.index:
        momentum_percentiles = []
        for time_period in time_periods:
            momentum_percentiles.append(fd.loc[row, f'{time_period} Return Percentile'])
        if request_type == 'm' or request_type == 'mean':
            fd.loc[row, 'HQM Score'] = mean(momentum_percentiles)
        else: fd.loc[row, 'HQM Score'] = median(momentum_percentiles)
    

if __name__ == "__main__":
    rfpath = "e:\\Python Learning\\Quant Finance\\algorithmic-trading-python-master\\starter_files\\sp_500_stocks.csv"
    dbfpath = "e:\\Python Learning\\Quant Finance\\algorithmic-trading-python-master\\starter_files\\data.xlsx"
    to_read = openpyxl.load_workbook(dbfpath)
    ws = to_read.active
    lrow = len(list(ws.rows))
    my_columns = ['Ticker', 'Price', 'Number of Shares to Buy', 'One-Year Price Return',
    'One-Year Return Percentile', 'Six-Month Price Return', 'Six-Month Return Percentile',
    'Three-Month Price Return', 'Three-Month Return Percentile', 'One-Month Price Return',
    'One-Month Return Percentile']
    final_dataframe = pd.DataFrame(columns = my_columns)
    request_type = ws.cell(row = lrow, column = 5).value
    portfolio_size = ws.cell(row = lrow, column = 4).value
    val = float(portfolio_size)
    batch_requests(rfpath, fd=final_dataframe)
    request_type = ws.cell(row = lrow, column = 5).value
    mean_median_momentum(final_dataframe, request_type)
        
    copied_data = final_dataframe.copy(deep=True)
    copied_data.sort_values('HQM Score', ascending = False, inplace = True)
    final_dataframe = copied_data.iloc[:50].copy(deep=True)
    final_dataframe.reset_index(inplace=True, drop = True)
    position_size = val/len(final_dataframe.index)
    for i in range (0, len(final_dataframe.index)):
        final_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/final_dataframe.loc[i, 'Price'])
    fpath = "e:/Python Learning/Quant Finance/algorithmic-trading-python-master/starter_files/Momentum recommendation.xlsx"
    ew.save_excel_file(fpath, request_type, final_dataframe)