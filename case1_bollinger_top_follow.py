# 일 평균 30만주이상 거래되는 nasdaq, newyork, amex 중(step2_300k_day_coms.xlsx)
# 최근 3개월간 10%(연간 100%) 이상 상승한 종목 중(step3_3mon_10p_up.xlsx)  <-- <월봉 양호한 종목 추출 목적>
# 볼린저밴드 상단 접근 종목(밴드 상단의 -20%선 이상) 중 시가가 기준선 위에 있고 종가가 상단선 80% 이상이며, 전일 시가, 종가 갭의 2배이상 상승한 종목
# 일일 1회 가동 
from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import mplfinance as mpf
import openpyxl
import datetime
import time

yf.pdr_override()
wb = openpyxl.Workbook()
now = datetime.datetime.now()
filename = datetime.datetime.now().strftime("%Y-%m-%d")

# wb.save('watch_data.xlsx')
sheet = wb.active
# cell name : date, simbol, company name, upper or lower or narrow band 
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'bol_high', 'bol_low', 'bol_gap(%)', 'rise_margin(%)', 'monthly_rise(%)', 'MA20', 'open', 'close', 'volume', 'industry', 'trade'])
wb.save('case1_bollinger_follow_'+filename+'.xlsx')

#회사 데이터 읽기
# df_com = pd.read_excel("step2_300k_day_coms.xlsx")
df_com = pd.read_excel("step3_3mon_10p_up_2021-07-08.xlsx")
i = 1
for i in range(len(df_com)):
    # now = datetime.datetime.now()
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '60d')
    try :
        df['MA20'] = df['Close'].rolling(window=20).mean()
        # df['MA60'] = df['Close'].rolling(window=60).mean()
        df['stddev'] = df['Close'].rolling(window=20).std()
        df['upper'] = df['MA20'] + (df['stddev']*2)
        df['lower'] = df['MA20'] - (df['stddev']*2)
        # df['vol_avr'] = df['Volume'].rolling(window=5).mean()
        df['gap'] = df['upper'] - df['lower']
        df['rise_margin'] = (df['upper'] - df['Close']) / df['Close'] * 100
       
        # df = df[19:]
        cur_close = df.iloc[-1]['Close']
        cur_open = df.iloc[-1]['Open']
        today_gap = cur_close - cur_open
        pre_open_price = df.iloc[-2]['Open']
        pre_close_price = df.iloc[-2]['Close']
        pre_gap = pre_close_price - pre_open_price
        # a1 = max([df.iloc[-2]['Open'], df.iloc[-2]['Close'], df.iloc[-3]['Open'], df.iloc[-3]['Close']])
        # a2 = max([df.iloc[-2]['Open'], df.iloc[-2]['Close']])
        # a3 = max([df.iloc[-3]['Open'], df.iloc[-3]['Close']])
        # df_u = df.iloc[-1]['upper']
        # df_l1 = df.iloc[-1]['lower']
        # df_l2 = df.iloc[-2]['lower']
        # df_l3 = df.iloc[-3]['lower']
        # df_g1 = df.iloc[-1]['gap']
        # df_g2 = df.iloc[-2]['gap']
        # df_g3 = df.iloc[-3]['gap']
        # df_v = df.iloc[-2]['vol_avr']
        # print(df_com.iloc[i]['simbol'])
        # print('볼린저 : ',df.iloc[-1]['MA20'], df.iloc[-1]['upper'], df.iloc[-1]['lower'])
        # print('볼린저 밴드폭 : ',df.iloc[-1]['bandwidth'], '%')


        if cur_close >= df.iloc[-1]['upper'] * 0.8 and cur_close > cur_open >= df.iloc[-1]['MA20'] and today_gap >= pre_gap * 2 and df.iloc[-1]['Volume'] >= 300000 :
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], df_com.iloc[i]['company_name'], \
                df.iloc[-1]['upper'], df.iloc[-1]['lower'], df.iloc[-1]['gap'], df.iloc[-1]['rise_margin'], df_com.iloc[i]['month_rise(%)'], df.iloc[-1]['MA20'], df.iloc[-1]['Open'], \
                    df.iloc[-1]['Close'], df.iloc[-1]['Volume'], df_com.iloc[i]['industry'],'buy'])
            wb.save('case1_bollinger_follow_'+filename+'.xlsx')
            print('buy', df_com.iloc[i]['symbol'])
            plt.figure(figsize=(9, 7))
            plt.subplot(2, 1, 1)
            plt.plot(df['upper'], color='green', label='Bollinger upper')
            plt.plot(df['lower'], color='brown', label='Bollinger lower')
            plt.plot(df['MA20'], color='black', label='MA20')
            plt.plot(df['Close'], color='blue', label='Price')
             
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
            df = df[['Open', 'High', 'Close', 'Volume']] 
                  
    except Exception as e:
        print(e)
        print('error', df_com.iloc[i]['symbol'])
    print(df_com.iloc[i]['symbol'])
    i += 1   

df_1 = pd.read_excel('case1_bollinger_follow_'+filename+'.xlsx') 
# df_b_f = df_1.sort_values(by = 'rise_margin(%)', ascending= False) # gap_close_ratio(%) 기준 올림차순으로 sorting
df_b_f = df_1.sort_values(by = 'rise_margin(%)')
df_b_f.to_excel('case1_bollinger_follow_'+filename+'.xlsx')
# df_b_f.to_excel('imsi_trend_bollinger_follow_sorted_'+filename+'.xlsx')
    #         time.sleep(0.1)
    # except Exception as e:
    #     print(e)
    #     time.sleep(0.1)

