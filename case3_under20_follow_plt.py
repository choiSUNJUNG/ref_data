# 일 평균 30만주이상 거래되는 nasdaq, newyork, amex 중(300k_day_coms.xlsx)
# 20일 이평선 아래에서 최저점을 지나 상승하는 종목 추출(case3_under20_follow.xlsx)
# 30$ 이하 종목으로 turn around 종목(20 이평선이 60 이평선 아래 있는 경우)
# 일 1회 가동 
from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import openpyxl
import datetime
import time

yf.pdr_override()
# wb = openpyxl.Workbook()
# now = datetime.datetime.now()
# filename = datetime.datetime.now().strftime("%Y-%m-%d")

# # wb.save('watch_data.xlsx')
# sheet = wb.active
# # cell name : date, simbol, company name, upper or lower or narrow band 
# sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'bol_high', 'bol_low', 'bol_gap(%)', 'rise_margin(%)', 'monthly_rise(%)', 'MA20', 'open', 'close', 'volume', 'industry', 'trade'])
# wb.save('case1_bollinger_follow_'+filename+'.xlsx')

#회사 데이터 읽기
# df_com = pd.read_excel("step2_300k_day_coms.xlsx")
# df_com = pd.read_excel("case3_under20_follow_2021-07-19.xlsx")
df_com = pd.read_excel("case3_over20_follow_2021-07-19.xlsx")
i = 1
for i in range(len(df_com)):
    # now = datetime.datetime.now()
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '70d')
    try :
        df['MA20'] = df['Close'].rolling(window=20).mean()
        df['MA60'] = df['Close'].rolling(window=60).mean()
        df['stddev'] = df['Close'].rolling(window=20).std()
        df['upper'] = df['MA20'] + (df['stddev']*2)
        df['lower'] = df['MA20'] - (df['stddev']*2)
        # df['vol_avr'] = df['Volume'].rolling(window=5).mean()
        # df['gap'] = df['upper'] - df['lower']
        # df['rise_margin'] = (df['upper'] - df['Close']) / df['Close'] * 100
       
        plt.figure(figsize=(9, 7))
        plt.subplot(2, 1, 1)
        plt.plot(df['upper'], color='green', label='Bollinger upper')
        plt.plot(df['lower'], color='brown', label='Bollinger lower')
        plt.plot(df['MA20'], color='red', label='MA20')
        plt.plot(df['MA60'], color='black', label='MA60')
        plt.plot(df['Close'], color='blue', label='Price')
        
        # plt.text(df_com.iloc[i]['company_name'])     
        plt.title(df_com.iloc[i]['symbol']+'stock price')
        plt.xlabel('time')
        plt.xticks(rotation = 45)
        plt.ylabel('stock price')
        plt.legend()
            
        plt.subplot(2, 1, 2)
        plt.plot(df['Volume'], color='blue', label='Volume')
        plt.ylabel('Volume')
        plt.xlabel('time')
        plt.xticks(rotation = 45)
        plt.legend()
        plt.show() 
        # df = df[['Open', 'High', 'Close', 'Volume']] 
                  
    except Exception as e:
        print(e)
        print('error', df_com.iloc[i]['symbol'])
    print(df_com.iloc[i]['symbol'] )
    print(i)
    print(len(df_com))
    i += 1   

