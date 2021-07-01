# 일 평균 30만주이상 거래되는 nasdaq, newyork, amex 중(step2_300k_day_coms.xlsx)
# 일일 상승률 10% 이상 종목 추출해서 순서대로 나열하기(step3_p1_daily_10p_up.xlsx)  <-- <상승 시작한 종목 추출 목적>

from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import matplotlib.pyplot as plt
import openpyxl
import datetime
import time

yf.pdr_override()
wb = openpyxl.Workbook()
now = datetime.datetime.now()
filename = datetime.datetime.now().strftime("%Y-%m-%d")

# wb.save('watch_data.xlsx')
sheet = wb.active
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'prev_price', 'today_price', 'daily_rise(%)', 'industry'])
wb.save('step4_p2_daily_10p_up_'+filename+'.xlsx')

#회사 데이터 읽기
# df_com = pd.read_excel("step2_300k_day_coms.xlsx")
df_com = pd.read_excel("step3_3mon_10p_up_2021-07-01.xlsx")
# while True:
#     try:
i = 1
for i in range(len(df_com)):
# now = datetime.datetime.now()
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '10d')
    try :
        df.start_price = df.iloc[-2]['Close']
        df.end_price = df.iloc[-1]['Close']
        df.daily_rise = (df.end_price / df.start_price - 1) * 100
        print(df_com.iloc[i]['symbol'])
        if df.daily_rise >= 10 :
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], df_com.iloc[i]['company_name'], \
                df.start_price, df.end_price, df.daily_rise, df_com.iloc[i]['industry']])
            wb.save('step4_p2_daily_10p_up_'+filename+'.xlsx')
    except Exception as e:
        print(e)
        print('error', df_com.iloc[i]['symbol'])
    i += 1
   
df_1 = pd.read_excel('step4_p2_daily_10p_up_'+filename+'.xlsx') 
df_b_f = df_1.sort_values(by = 'daily_rise(%)', ascending= False) # gap_close_ratio(%) 기준 올림차순으로 sorting
df_b_f.to_excel('step4_p2_daily_10p_up_'+filename+'.xlsx')

