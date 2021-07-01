# 일 평균 30만주이상 거래되는 nasdaq, newyork, amex 중(step2_300k_day_coms.xlsx)
# 최근 3개월간 10%(연간 100%) 이상 상승한 종목 추출해서 상승률 순서대로 나열하기(step3_3mon_10p_up.xlsx)  <-- <월봉 양호한 종목 추출 목적>
# 월 1회 가동 

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
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'start_price', 'end_price', 'month_rise(%)', 'industry'])
wb.save('step3_3mon_10p_up_'+filename+'.xlsx')

#회사 데이터 읽기
df_com = pd.read_excel("step2_300k_day_coms.xlsx")
i = 1
for i in range(len(df_com)):
    # now = datetime.datetime.now()
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '3mo')
    try :
        df.start_price = df.iloc[1]['Close']
        df.end_price = df.iloc[-1]['Close']
        df.month_rise = (df.end_price / df.start_price -1) * 100
        if df.month_rise >= 10:
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], df_com.iloc[i]['company_name'], \
                    df.start_price, df.end_price, df.month_rise, df_com.iloc[i]['industry']])
            wb.save('step3_3mon_10p_up_'+filename+'.xlsx')        
            print('10%이상', df_com.iloc[i]['symbol'])
        else :
            print('10%이하', df_com.iloc[i]['symbol'])
    except Exception as e:
        print(e)
        print('error', df_com.iloc[i]['symbol'])
    i += 1   

df_1 = pd.read_excel('step3_3mon_10p_up_'+filename+'.xlsx') 
df_b_f = df_1.sort_values(by = 'month_rise(%)', ascending= False) # gap_close_ratio(%) 기준 올림차순으로 sorting
df_b_f.to_excel('step3_3mon_10p_up_'+filename+'.xlsx')

