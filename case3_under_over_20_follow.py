# 일 평균 30만주이상 거래되는 nasdaq, newyork, amex 중(300k_day_coms.xlsx)
# 20일 이평선 아래에서 최저점을 지나 상승하는 종목 추출(case3_under20_follow.xlsx)
# 30$ 이하 종목으로 turn around 종목(20 이평선이 60 이평선 아래 있는 경우) 또는 
# 30$ 이하 종목으로 20이평선 돌파 종목(case3_over20_follow.xlsx)
# 일 1회 가동 

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
sheet.append(['time', 'market', 'symbol', 'code', 'company_name', 'close_min', 'close', 'MA20', 'MA60', 'gap', 'industry'])
wb.save('case3_under20_follow_'+filename+'.xlsx')
wb.save('case3_over20_follow_'+filename+'.xlsx')

#회사 데이터 읽기
df_com = pd.read_excel("step2_300k_day_coms.xlsx")

i = 1
for i in range(len(df_com)):
    # df = pdr.get_data_yahoo(df_com1.iloc[i]['Symbol'], period = '1mo')  # 기간 1month
    df = pdr.get_data_yahoo(df_com.iloc[i]['symbol'], period = '70d')  # 기간 10일
    try : 
        df['MA20'] = df['Close'].rolling(window=20).mean()
        df['MA60'] = df['Close'].rolling(window=60).mean()
        # df['vol_mean'] = df['Volume'].rolling(window=10).mean()
        # df['vol_max'] = df['Volume'].rolling(window=10).max()
        df['close_min'] = df['Close'].rolling(window=20).min()
        gap = df.iloc[-1]['Close'] - df.iloc[-1]['MA20']   
        # df['close_mean'] = df['Close'].rolling(window=5).max()
    
        # print(df)
  
        # if df.iloc[-2]['close_min'] <= df.iloc[-1]['Close'] <= df.iloc[-1]['MA20'] <= df.iloc[-1]['MA60'] < 50 :
        #     sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
        #         df_com.iloc[i]['company_name'], df.iloc[-2]['close_min'], df.iloc[-1]['Close'], df.iloc[-1]['MA20'], df.iloc[-1]['MA60'], \
        #             gap, df_com.iloc[i]['industry']])
        #     wb.save('case3_under20_follow_'+filename+'.xlsx')
        #     print('20일선 접근', df_com.iloc[i]['symbol'])
        if df.iloc[-2]['Close'] <= df.iloc[-2]['MA20'] and df.iloc[-1]['Close'] >= df.iloc[-1]['MA20'] :
            sheet.append([now, df_com.iloc[i]['market'], df_com.iloc[i]['symbol'], df_com.iloc[i]['code'], \
                df_com.iloc[i]['company_name'], df.iloc[-2]['close_min'], df.iloc[-1]['Close'], df.iloc[-1]['MA20'], df.iloc[-1]['MA60'], \
                    gap, df_com.iloc[i]['industry']])
            wb.save('case3_over20_follow_'+filename+'.xlsx')
            print('20일선 돌파', df_com.iloc[i]['symbol']) 
            
    except Exception as e:
        print(e)
        print('error', df_com.iloc[i]['symbol'])    
    i += 1   
df_1 = pd.read_excel('case3_over20_follow_'+filename+'.xlsx') 
df_b_f = df_1.sort_values(by = 'gap') # gap 기준 올림차순으로 sorting
df_b_f.to_excel('case3_over20_follow_'+filename+'.xlsx')
# df_1 = pd.read_excel('case3_under20_follow_'+filename+'.xlsx') 
# df_b_f = df_1.sort_values(by = 'gap') # gap 기준 올림차순으로 sorting
# df_b_f.to_excel('case3_under20_follow_'+filename+'.xlsx')