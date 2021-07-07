# 일 평균 30만주이상 거래되는 nasdaq, newyork, amex(300k_day_coms.xlsx) 중
# 최근 3개월간 10%(연간 100%) 이상 상승한 종목 추출해서 상승률 순서대로 나열한 종목(step3_3mon_10p_up.xlsx) 중  <-- <월봉 양호한 종목 추출 목적>
# 60일 거래량 평균 대비 60일내 일일 최대 거래량이 500% 이상인 종목(step4_p3_volume_5times_up.xlsx)과
# 볼린저밴드 하단 근처 종목(밴드갭의 +-20%선 이내) 중 당일 종가가 전일 시작가 또는 종가 중 큰 값보다 높은 종목과
# 일일 상승률 10% 이상 종목(step3_p1_daily_10p_up.xlsx)을 통합하여
# 현 주가가 12주(60일선)에 접근중인 종목 추출(step6_final_60line.xlsx)
# 
from pandas_datareader import data as pdr
import yfinance as yf
import pandas as pd
import openpyxl
import datetime
import time

yf.pdr_override()
now = datetime.datetime.now()
filename = datetime.datetime.now().strftime("%Y-%m-%d")
wb = openpyxl.Workbook()
sheet = wb.active
# cell name 생성
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'price', 'vol_max', 'vol_mean', 'industry', 'macd', 'sto', 'adx'])
wb.save('step6_final_60line_'+filename+'.xlsx')

#회사 데이터 읽기

df_com = pd.read_excel("step4_p3_volume_5times_up_2021-07-06.xlsx") 

i = 1
for i in range(len(df_com)):
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '70d')  # 기간 70일
    try :
        # df['high_max'] = df['Close'].rolling(window=20).max()
        df['ma60'] = df['Close'].rolling(window=60).mean()
        # df['min_14'] = df['Low'].rolling(14).min()
        # df['max_14'] = df['High'].rolling(14).max()
        # df.min = df['min_14']
        # df.max = df['max_14']
        # df['sto_K_14'] = (df['Close'] - df.min) / (df.max - df.min) * 100   #stochastics_fast의 %K
        # df['sto_D_5'] = df['sto_K_14'].rolling(5).mean()    #stochastics_slow의 %K
        # df['sto_DS_3'] = df['sto_D_5'].rolling(3).mean()    #stochastics_slow의 %D
   
        # print(df)
        # 오늘 종가(현재가)가 20 고점(전고점) 보다 높고 60이평선 위에 있으며 60 이평선 또한 상승하는 경우
        if df.iloc[-1]['Close'] < df.iloc[-1]['ma60'] * 1.1 :  
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
                df_com.iloc[i]['company_name'], df_com.iloc[i]['price'], df_com.iloc[i]['vol_max'], df_com.iloc[i]['vol_mean'], \
                    df_com.iloc[i]['industry'], 'macd', 'sto', 'adx'])
            wb.save('step6_final_60line_'+filename+'.xlsx')
            print('매수발생 : ', df_com.iloc[i]['symbol'])
    except Exception as e:
        print(e)
        print('error', df_com.iloc[i]['symbol'])
    i += 1  
