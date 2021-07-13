# -*- coding: utf-8 -*-
"""
Created on Sat Jun 19 16:41:28 2021

@author: Hamzah
"""
import yfinance as yf
import xlsxwriter
import matplotlib.pyplot as plt
import pandas as pd
import xlrd
import os
import matplotlib.pyplot as plt
import numpy as np
import scipy
import random
import itertools
import datetime
from datetime import date
import time
from time import sleep
from math import isnan
from matplotlib.backends.backend_pdf import PdfPages
from textwrap import wrap

# Directory for files
os.chdir(r"C:\Users\Hamzah\Documents\Python files")
wb = xlrd.open_workbook(r'C:\Users\Hamzah\Documents\Python files\final tickers.xlsx')

wb_tickers=wb.sheet_by_index(0)
N=wb_tickers.nrows
pd.set_option('display.max_columns', None)  

#tickers
master_tickers=[]
for i in range(1,N):
    x=wb_tickers.row_values(i)[0]
    master_tickers.append(x)
    
#Set dates
start_date="2015-11-01"
end_date="2020-12-31"



div_threshold=7
pick_threshold=0.76
stock_threshold=0.5
cap_threshold=0

frequency_stock=30

pp = PdfPages('dividends above 7 v4.pdf')
possibles=[]
exceptions=[]
final_data={}
for ticker in master_tickers:
    try:
        test=pd.DataFrame()
        stock=yf.Ticker(ticker)
        divs=stock.dividends
        divs=divs.groupby([divs.index.year]).sum()
        data=yf.download(ticker,start_date,end_date)
        data=data.Close
        price=data
        test[ticker," price"]=price
        #test[ticker]=price.pct_change().fillna(0)
        num=len(data)
        test[ticker," dividend yield"]=price.pct_change(periods=frequency_stock).fillna(0)
        for i in range(frequency_stock,num):
            test[ticker," dividend yield"][i]=(divs.loc[data.index[i].year]/data[i])
        test[ticker," dividend yield"]=test[ticker," dividend yield"]*100
        test[ticker," moving average"]=test[ticker," price"].rolling(window=frequency_stock).mean()
        above_threshold=[]
        above_threshold_stock=[]
        for i in range(frequency_stock,num):
            if test[ticker," price"][i]>=test[ticker," moving average"][i]:
                above_threshold_stock.append(1)
        above_stock_n=len(above_threshold_stock)
        for i in range(frequency_stock,num):
            if test[ticker," dividend yield"][i]>=div_threshold:
                above_threshold.append(1)
        above_n=len(above_threshold)
        if above_n/num >= pick_threshold and above_stock_n/num >= cap_threshold:
            with pd.ExcelWriter("dividend data above 7.xlsx",engine='openpyxl',mode='a') as writer:
                test.to_excel(writer,sheet_name=ticker)
                divs.to_excel(writer,sheet_name=(ticker+"_divs"))
            possibles.append(ticker)
            final_data[ticker]=test
            b=stock.info
            TITLE=(b['longName'],b['industry'],b['industry'])
            a=test.plot.line(title=TITLE,subplots=True,grid=True,figsize=(26,18),lw=4,fontsize=24)
            a[0].figure.savefig(pp,format="pdf")
    except:
        exceptions.append(ticker)
pp.close()




# Directory for files for correlation
os.chdir(r"C:\Users\Hamzah\Documents\Python files")
wb = xlrd.open_workbook(r'C:\Users\Hamzah\Documents\Python files\possible tickers.xlsx')

wb_tickers=wb.sheet_by_index(0)
N=wb_tickers.nrows
pd.set_option('display.max_columns', None)  


stocks=[]
for i in range(1,N):
    x=wb_tickers.row_values(i)[0]
    stocks.append(x)

stocks_data=pd.DataFrame()

for stock in stocks:
    data=yf.download(stock,start_date,end_date)
    data=data.Close
    stocks_data[stock]=data


buys=['MMLP','CSQ']
po=pd.DataFrame()
for buy in buys:
    data=yf.download(buy,start_date,end_date)
    data=data.Close
    po[buy]=data
    
    

    
