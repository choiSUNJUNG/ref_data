# 일 평균 30만주이상 거래되는 nasdaq, newyork, amex 중(300k_day_coms.xlsx)
# 60일 거래량 평균 대비 60일내 일일 최대 거래량이 500% 이상인 종목 추출(5_volume_follow.xlsx)
# 주1회 가동(매주말) 

from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import openpyxl
import datetime
import time


yf.pdr_override()
wb = openpyxl.Workbook()
now = datetime.datetime.now()
filename = datetime.datetime.now().strftime("%Y-%m-%d")
sheet = wb.active
# cell name 생성
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'price', 'vol_max', 'vol_mean', 'industry'])
wb.save('step4_p3_volume_5times_up_'+filename+'.xlsx')

#회사 데이터 읽기
# df_com = pd.read_excel("step2_300k_day_coms.xlsx")
df_com = pd.read_excel("step3_3mon_10p_up_2021-07-01.xlsx")
now = datetime.datetime.now()
i = 1
for i in range(len(df_com)):
    # df = pdr.get_data_yahoo(df_com1.iloc[i]['Symbol'], period = '1mo')  # 기간 1month
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '70d')  # 기간 10일
    try :
        df['vol_mean'] = df['Volume'].rolling(window=60).mean()
        df['vol_max'] = df['Volume'].rolling(window=60).max()
    
        # print(df)
  
        if df.iloc[-1]['vol_max'] > df.iloc[-1]['vol_mean'] * 5 : # 일일 최대 거래량이 60일 평균 거래량보다 500% 이상인 경우
            k = 1
            for k in range(len(df)):
                if df.iloc[-1]['vol_max'] == df.iloc[k]['Volume'] and df.iloc[k]['Close'] >= df.iloc[k]['Open'] :
                    sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
                        df_com.iloc[i]['company_name'], df.iloc[-1]['Close'], df.iloc[-1]['vol_max'], df.iloc[-1]['vol_mean'], \
                            df_com.iloc[i]['industry']])
            wb.save('step4_p3_volume_5times_up_'+filename+'.xlsx')
            print('세력', df_com.iloc[i]['symbol'])
            k += 1
    except Exception as e:
        print(e)
        print('error', df_com.iloc[i]['symbol'])    
    i += 1   
   