import requests, math, openpyxl, os
import numpy as np
import pandas as pd
import equal_weights as ew
from secrets1 import IEX_CLOUD_API_TOKEN
from statistics import mean, median


def val_requests(rfpath, rv_dataframe):
    stocks = pd.read_csv(rfpath)
    symbol_grp = list(ew.chunks(stocks['Ticker'], 100))
    symbol_strings = []
    for symbol in symbol_grp:
        symbol_strings.append(','.join(symbol))
    for ss in symbol_strings:
        batch_api_call_url = f'https://sandbox.iexapis.com/stable/stock/market/batch?symbols={ss}&types=quote,advanced-stats&token={IEX_CLOUD_API_TOKEN}'
        data = requests.get(batch_api_call_url).json()
        for symbol in ss.split(","):
            enterprise_value = data[symbol]['advanced-stats']['enterpriseValue']
            ebitda = data[symbol]['advanced-stats']['EBITDA']
            gross_profit = data[symbol]['advanced-stats']['grossProfit']
            try:
                ev_to_ebitda = enterprise_value/ebitda
            except TypeError:
                ev_to_ebitda = 0
            try:
                ev_to_gp = enterprise_value/gross_profit
            except TypeError:
                ev_to_gp = 0
            rv_dataframe.loc[len(rv_dataframe)] = [symbol, data[symbol]['quote']['latestPrice'], 'N/A', 
                                            data[symbol]['quote']['peRatio'], 'N/A',
                                            data[symbol]['advanced-stats']['priceToBook'], 'N/A',
                                            data[symbol]['advanced-stats']['priceToSales'], 'N/A',
                                            ev_to_ebitda, 'N/A', ev_to_gp, 'N/A','N/A'
                                            ]
    for column in ['Price-to-Earnings Ratio', 'Price-to-Book Ratio', 'Price-to-Sales Ratio', 'EV/EBITDA Ratio', 'EV/GP Ratio']:
        rv_dataframe[column].fillna(rv_dataframe[column].mean(), inplace = True)
    metrics = ['Price-to-Earnings', 'Price-to-Book', 'Price-to-Sales', 'EV/EBITDA', 'EV/GP']
    for m in metrics:
        rv_dataframe[f'{m} Percentile'] = rv_dataframe[f'{m} Ratio'].rank()/len(rv_dataframe.index)


def mean_median_value(fd, request_type):
    metrics = ['Price-to-Earnings', 'Price-to-Book', 'Price-to-Sales', 'EV/EBITDA', 'EV/GP']
    for row in rv_dataframe.index:
        value_percentiles = []
        for metric in metrics:
            value_percentiles.append(rv_dataframe.loc[row, f'{metric} Percentile'])
        if request_type == 'mean' or request_type == 'm':
            rv_dataframe.loc[row, 'RV Score'] = mean(value_percentiles)
        else: rv_dataframe.loc[row, 'RV Score'] = median(value_percentiles)


if __name__ == "__main__":
    rfpath = "e:\\Python Learning\\Quant Finance\\algorithmic-trading-python-master\\starter_files\\sp_500_stocks.csv"
    rv_columns = ['Ticker', 'Price', 'Number of Shares to Buy', 'Price-to-Earnings Ratio', 
    'Price-to-Earnings Percentile', 'Price-to-Book Ratio', 'Price-to-Book Percentile',
    'Price-to-Sales Ratio', 'Price-to-Sales Percentile', 'EV/EBITDA Ratio', 
    'EV/EBITDA Percentile', 'EV/GP Ratio', 'EV/GP Percentile', 'RV Score']
    rv_dataframe = pd.DataFrame(columns = rv_columns)
    val_requests(rfpath, rv_dataframe)
    dbfpath = "e:\\Python Learning\\Quant Finance\\algorithmic-trading-python-master\\starter_files\\data.xlsx"
    to_read = openpyxl.load_workbook(dbfpath)
    ws = to_read.active
    lrow = len(list(ws.rows))
    portfolio_size = ws.cell(row = lrow, column = 4).value
    val = float(portfolio_size)
    request_type = ws.cell(row = lrow, column = 5).value
    mean_median_value(rv_dataframe, request_type)    
    copied_data = rv_dataframe.copy(deep=True)
    copied_data.sort_values('RV Score', ascending = False, inplace = True)
    rv_dataframe = copied_data.iloc[:50].copy(deep=True)
    rv_dataframe.reset_index(inplace=True, drop = True)
    position_size = val/len(rv_dataframe.index)
    for i in range (0, len(rv_dataframe.index)):
        rv_dataframe.loc[i, 'Number of Shares to Buy'] = math.floor(position_size/rv_dataframe.loc[i, 'Price'])
    fpath = "e:/Python Learning/Quant Finance/algorithmic-trading-python-master/starter_files/Value_recommendation.xlsx"
    ew.save_excel_file(fpath, request_type, rv_dataframe)