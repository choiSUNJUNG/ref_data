# 5일 평균 거래량 30만주 이상 종목 추출하기
from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import openpyxl
import datetime
import time

yf.pdr_override()
wb = openpyxl.Workbook()
# wb.save('watch_data.xlsx')
sheet = wb.active
# cell name : date, simbol, company name, upper or lower or narrow band 
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'industry', 'vol_avr'])
wb.save('step2_300k_day_coms.xlsx')
#회사 데이터 읽기
df_com1 = pd.read_excel("nasdaq.xlsx")
df_com2 = pd.read_excel("nyse.xlsx")
df_com3 = pd.read_excel("amex.xlsx")
# print(df_com)
# print(len(df_com))
now = datetime.datetime.now()
i = 1
for i in range(len(df_com1)):
    # df = pdr.get_data_yahoo(df_com1.iloc[i]['Symbol'], period = '1mo')  # 기간 1month
    df = pdr.get_data_yahoo(df_com1.iloc[i]['Symbol'], period = '10d')  # 기간 10일
    df['vol_avr'] = df['Volume'].rolling(window=5).mean()
    df_v = df.iloc[-2]['vol_avr']
    print(df_com1.iloc[i]['Symbol'])
    if df_v > 500000 :  # 일평균 30만주 이상 거래 종목 추출
        sheet.append([now, 'NASDAQ', df_com1.iloc[i]['Symbol'], df_com1.iloc[i]['IndustryCode'], \
            df_com1.iloc[i]['Name'], df_com1.iloc[i]['Industry'], df_v])
    i += 1   
wb.save('step2_300k_day_coms.xlsx')
i = 1
for i in range(len(df_com2)):
    # df = pdr.get_data_yahoo(df_com2.iloc[i]['Symbol'], period = '1mo')  # 기간 1month
    df = pdr.get_data_yahoo(df_com2.iloc[i]['Symbol'], period = '10d')  # 기간 10일
    try:
        df['vol_avr'] = df['Volume'].rolling(window=5).mean()
        df_v = df.iloc[-2]['vol_avr']
        print(df_com2.iloc[i]['Symbol'])
        if df_v > 500000 :
            sheet.append([now, 'NYSE', df_com2.iloc[i]['Symbol'], df_com2.iloc[i]['IndustryCode'], \
                df_com2.iloc[i]['Name'], df_com2.iloc[i]['Industry'], df_v])
    except:
        print("no data found")
    i += 1 
wb.save('step2_300k_day_coms.xlsx')

i = 1
for i in range(len(df_com3)):
    # df = pdr.get_data_yahoo(df_com3.iloc[i]['Symbol'], period = '1mo')  # 기간 1month
    df = pdr.get_data_yahoo(df_com3.iloc[i]['Symbol'], period = '10d')  # 기간 10일

    df['vol_avr'] = df['Volume'].rolling(window=5).mean()
    # print(df)
    # print(df['Symbol'])
    # print(df['vol_avr'])
    df_v = df.iloc[-2]['vol_avr']
    print(df_com3.iloc[i]['Symbol'])
    if df_v > 500000 :
        sheet.append([now, 'AMEX', df_com3.iloc[i]['Symbol'], df_com3.iloc[i]['IndustryCode'], \
            df_com3.iloc[i]['Name'], df_com3.iloc[i]['Industry'], df_v])
    i += 1 
wb.save('step2_300k_day_coms.xlsx')