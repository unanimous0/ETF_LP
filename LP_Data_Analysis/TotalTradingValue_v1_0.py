# -*- coding: utf-8 -*-
"""
Created on Wed Nov 24 22:01:22 2021

@author: BAHAMA
"""

#%%
import logging
import numpy as np
import pandas as pd
import xlwings as xw
from datetime import datetime

logger = logging.getLogger()

#%%
# 파일 경로 확인 및 raw data 데이터프레임 생성
# input_date = input()
input_date = '20211119'
file_dir = "D:/고은환 공유폴더/고은환/LP데이터/"
#file_dir = "C:/Users/Simons/Downloads/"
#file_dir = "/Users/bahama/Downloads/"
#file_name = f'[ETF]회원사별 LP 거래실적(1일)(주문번호-843-1)_{input_date}.csv'
file_name = "[ETF]회원사별 LP 거래실적(1일)(주문번호-843-1)_19_23_취합본.csv"
file_path = file_dir + file_name

raw_book = xw.Book(file_path)
raw_data_df = raw_book.sheets(1).used_range.options(pd.DataFrame, index=False).value


#%%
# 거래일자 데이터를 날짜 형식으로 변경
raw_data_df['거래일자'] = raw_data_df['거래일자'].astype(int).astype(str)
raw_data_df['거래일자'] = pd.to_datetime(raw_data_df['거래일자']).dt.date

# 중복제거된 거래일자 리스트
date_list = list(raw_data_df['거래일자'].drop_duplicates())


#%%
# 파일 인코딩 확인
#import chardet
#rawdata = open(file_path, 'rb').read()
#result = chardet.detect(rawdata)
#charenc = result['encoding']
#print(charenc)


#%%
# 딕셔너리 {종목코드:종목명} 생성
etf_codes = raw_data_df['종목코드'].drop_duplicates()
etf_names = raw_data_df['종목명'].drop_duplicates()
etf_names.index = etf_codes

etf_code_name_dict = etf_names.to_dict()


#%%
# raw data 데이터프레임 재구성
df_with_7cols = raw_data_df[['거래일자', '종목코드', '종목명', '회원사명', 'LP매도거래대금', 'LP매수거래대금']]

# 총거래대금 column 추가 (순매수/순매도대금 아닌 총거래대금)
df_with_7cols['LP총거래대금'] = df_with_7cols['LP매도거래대금'] + df_with_7cols['LP매수거래대금']
df_with_7cols['LP총거래대금']

# KB증권을 포함한 주요 회원사명 리스트 생성
lpcompanies = ['KB증권', 'NH투자증권', '신한투자', '메리츠', '키움증권']

# 주요 회원사 데이터만 추출
main_company_condition = df_with_7cols['회원사명'].isin(lpcompanies)
main_lpcomp_df = df_with_7cols.loc[main_company_condition]

# 주요 회원사의 총거래대금(TOTAL TRADING VALUE) 변수 생성
KB_TOTAL_TRADING_VALUE = 0
NH_TOTAL_TRADING_VALUE = 0
SH_TOTAL_TRADING_VALUE = 0
MZ_TOTAL_TRADING_VALUE = 0
KW_TOTAL_TRADING_VALUE = 0

# # 속도 향상을 위해 데이터프레임 numpy array로 변형
# np_df = df.to_numpy()


#%%
code_lpcomp_date_value_dict = {}

for etf_code in etf_code_name_dict.keys():
    one_etf_df = main_lpcomp_df[main_lpcomp_df['종목코드'] == etf_code]
   
    lpcomp_date_value_dict = {}
   
    for lpcomp in lpcompanies:
        lpcomp_date_list = one_etf_df[one_etf_df['회원사명'] == lpcomp]['거래일자'].values
        lpcomp_value_list = one_etf_df[one_etf_df['회원사명'] == lpcomp]['LP총거래대금'].values
       
        lpcomp_date_value_dict[lpcomp] = {x : y for x, y in zip(lpcomp_date_list, lpcomp_value_list)}
       
    code_lpcomp_date_value_dict[etf_code] = lpcomp_date_value_dict
   
    del lpcomp_date_value_dict
   

#%%
# 컬럼과 각 컬럼에 해당하는 데이터가 섞이지 않도록 딕셔너리 생성 (회원사명 순서 바뀌는 오류 방지)
lpcomp_value_dict = {
                      lpcompanies[0] : KB_TOTAL_TRADING_VALUE,
                      lpcompanies[1] : NH_TOTAL_TRADING_VALUE,
                      lpcompanies[2] : SH_TOTAL_TRADING_VALUE,
                      lpcompanies[3] : MZ_TOTAL_TRADING_VALUE,
                      lpcompanies[4] : KW_TOTAL_TRADING_VALUE
                      }

# 새로운 열을 갖는 데이터프레임 생성
total_value_df = pd.DataFrame(columns=['거래일자', '종목코드', '종목명', lpcompanies[0], lpcompanies[1], lpcompanies[2], lpcompanies[3], lpcompanies[4]])    
   
# test_df = pd.DataFrame(code_lpcomp_date_value_dict).T

# # 종목이 KR7069500007인 ETF의 LP사 중 KB의 2021-11-19 날짜의 총 거래대금
# test_df.loc['KR7069500007']['KB증권'][datetime.date(2021,11,19)]


"""
    # 위에서 정의한 date_list나, 나중에 파일명 읽어올 때 하나씩 추가되도록 만들면될듯
    datelist = ['20211119', '20211122', '20211123']    
   
    # str 형식의 날짜를 datetime 형태로 변경
    datelist = [datetime.date(datetime.strptime(x, "%Y%m%d")) for x in datelist]
"""


#%%
# 행별 회사의 총LP거래대금 데이터 가져오기
def getTotalValue(etf_code, lpcomp, date) -> int:
    dict_set_1 = code_lpcomp_date_value_dict[etf_code]
    dict_set_2 = dict_set_1[lpcomp]
   
    # TODO dict_set_2 딕셔너리에서 LP회원사 키가 갖는 값의 딕셔너리에 날짜가 없는 경우 (-> {'KB증권': {}, ~} 이런 식으로 값이 없는 경우에 대한 예외처리 필요)
    # 날짜 키로 조회하면서 해당 날짜에 총거래대금이 없는 경우 0 으로 값 설정
    if date in dict_set_2.keys():
        last_dict_value = dict_set_2[date]
    else:
        last_dict_value = 0.0
   
    return int(last_dict_value)      # 굳이 정수로 안바꿔도 되긴함


#%%
# 아래 for문을 감싸는 날짜 기준으로 도는 for문이 있어야할듯
for date in date_list:
    for etf_code in code_lpcomp_date_value_dict.keys():
        # date = date                                            # 거래일자
        # etf_code = etf_code                                    # 종목코드
        etf_name = etf_code_name_dict[etf_code]                  # 종목명
       
        # LP총거래대금
        # KB증권(KB_TOTAL_TRADING_VALUE)
        lpcomp_value_dict[lpcompanies[0]] = getTotalValue(etf_code, lpcompanies[0], date)
       
        # NH투자증권(NH_TOTAL_TRADING_VALUE)
        lpcomp_value_dict[lpcompanies[1]] = getTotalValue(etf_code, lpcompanies[1], date)
       
        # 신한투자(SH_TOTAL_TRADING_VALUE)
        lpcomp_value_dict[lpcompanies[2]] = getTotalValue(etf_code, lpcompanies[2], date)
       
        # 메리츠(MZ_TOTAL_TRADING_VALUE)
        lpcomp_value_dict[lpcompanies[3]] = getTotalValue(etf_code, lpcompanies[3], date)
       
        # 키움증권(KW_TOTAL_TRADING_VALUE)
        lpcomp_value_dict[lpcompanies[4]] = getTotalValue(etf_code, lpcompanies[4], date)
       
        # total_value_df에 넣기 위한 행을 리스트로 만들기 (만들면 df.append를 통해 그 행을 total_value_df에 추가)
        raw_set = [date, etf_code, etf_name, lpcomp_value_dict[lpcompanies[0]], lpcomp_value_dict[lpcompanies[1]], lpcomp_value_dict[lpcompanies[2]], lpcomp_value_dict[lpcompanies[3]], lpcomp_value_dict[lpcompanies[4]]]
        
        raw_append_dict = {
                    total_value_df.columns[0] : date,                                      # 거래일자
                    total_value_df.columns[1] : etf_code,                                  # 종목코드
                    total_value_df.columns[2] : etf_name,                                  # 종목명
                    total_value_df.columns[3] : lpcomp_value_dict[lpcompanies[0]],        # KB증권
                    total_value_df.columns[4] : lpcomp_value_dict[lpcompanies[1]],        # NH투자증권
                    total_value_df.columns[5] : lpcomp_value_dict[lpcompanies[2]],        # 신한투자
                    total_value_df.columns[6] : lpcomp_value_dict[lpcompanies[3]],        # 메리츠
                    total_value_df.columns[7] : lpcomp_value_dict[lpcompanies[4]]         # 키움증권   
                }
        
        total_value_df = total_value_df.append(raw_append_dict, ignore_index=True)


#%%
class InputValueError(Exception):
    def __init__(self, msg="입력된 값이 입력 조건에 부합하지 않습니다. 입력 조건에 맞게 다시 입력하세요."):
        self.msg = msg
        
    def __str__(self):
        return self.msg

# 데이터프레임은 만들어졌고 날짜/기간과 종목코드로 조회했을 때, 총거래대금을 나타낼 수 있도록 작성
# 입력받은 날짜
# input_date2 = '20211119'
# condition_date = datetime.date(datetime.strptime(input_date2, "%Y%m%d"))

# 입력받은 기간
# TODO 기간은 어떻게 받아야하나 -> 컴퓨터내 기간으로 오면 주말 제거애야하니, 엑셀 파일명에서 받아온 날짜(date_list 변수 활용)로 하는게 좋을듯
# condition_period = "20211119 ~ 20211123"

# 입력받은 종목 코드
# condition_etf_code = 'KR7069500007'


# 입력받은 조건에 따라 데이터프레임 생성
def makeConditionalDataFrame(start_date, end_date, etf_code) -> pd.DataFrame():
    _start_date = datetime.date(datetime.strptime(start_date, "%Y%m%d"))
    _end_date = datetime.date(datetime.strptime(end_date, "%Y%m%d"))

    if _start_date == _end_date:
        _condition_df = total_value_df[(total_value_df['거래일자'] == _start_date) & (total_value_df['종목코드'] == etf_code)]
        # _condition_df2 = total_value_df.query("거래일자 == _start_date & 종목코드 == etf_code")

    elif _start_date < _end_date:
        # 위 조건인 경우, 시작일과 종료일을 date_list에서 시작과 끝으로 지정해서 그 사이의 날짜들 다 가져오면 됨
#        start_date_index = date_list.index(_start_date)
#        end_date_index = date_list.index(_start_date)
#        condition_period_list = date_list[start_date_index : end_date_index + 1]

        _condition_df = total_value_df[(total_value_df['거래일자'] >= _start_date) & (total_value_df['거래일자'] <= _end_date) & (total_value_df['종목코드'] == etf_code)]
        
    else:
        raise InputValueError()

    return _condition_df


# TODO 기간을 시작일과 종료일 입력받을 수도 있게 만들고 & T일로부터 1W, 1M, 1Y 전의 기간도 출력하도록 만들 것
# TODO 1W, 1M 등 기간 선택 방법 두가지 1) 시작일과 종료일을 구한 후 범위로 DF에서 가져오기 2) 시작일과 종료일 인덱스로 해당되는 날짜들을 DF에서 가져오기 -> 1) 선택

#def setPeriod(period, start_date, etf_code) -> datetime():
#
#    start_date_index = date_list.index(start_date)
#
#    # date_list의 날짜는 데이터 제공업체에서 주는 날짜를 기준으로 작성되었음 (영업일 기준)
#    # 1일 기준 -> 시작일 당일
#    one_day_start_date = start_date
#    one_day_end_date = start_date
#
#    # 1주 기준 -> 시작일 포함 5일 (T-4, T-3, T-2, T-1, T)
#    one_week = date_list[start_date_index - 4 : start_date_index + 1]
#    one_week_start_date = one_week[0]
#    one_week_end_date = one_week[-1]
#
#    # 1달 기준 -> 시작일 포함 25일 (T-24 ~ T)
#    one_month = date_list[start_date_index - 24 : start_date_index + 1]
#    one_month_start_date = one_month[0]
#    one_month_end_date = one_month[-1]
#    
#    # 1분기 기준 -> 시작일 포함 70일 (T-69 ~ T)
#    one_quarter = date_list[start_date_index - 69 : start_date_index + 1]
#    one_quarter_start_date = one_quarter[0]
#    one_quarter_end_date = one_quarter[-1]
#
#    # 1년 기준 -> 시작일 포함 255일 (T-254 ~ T)
#    # one_year = date_list[start_date_index - 254 : start_date_index + 1]
#    # one_year_start_date = one_year[0]
#    # one_year_end_date = one_year[-1]
#
#    if period == '1D':
#        _start_date = one_day_start_date
#        _end_date = one_day_end_date
#
#    elif period == '1W':
#        _start_date = one_week_start_date
#        _end_date = one_week_end_date
#        
#    elif period == '1M':
#        _start_date = one_month_start_date
#        _end_date = one_month_end_date
#
#    elif period == '1Q':
#        _start_date = one_quarter_start_date
#        _end_date = one_quarter_end_date
#
#    # elif period == '1Y':
#    #     _start_date = one_year_start_date
#    #     _end_date = one_year_end_date
#
#    return (_start_date, _end_date)


while True:    
    try:
        # input_start_date = input("시작일(yyyymmdd): ")
        # input_end_date = input("종료일(yyyymmdd): ")
        # input_etf_code = input("ETF 종목 코드: ")
        input_start_date = '20211119'
        input_end_date = '20211123'
        input_etf_code = 'KR7069500007'
        
        if (len(input_start_date) != 8) | (len(input_end_date) != 8):
            raise InputValueError()

        condition_df = makeConditionalDataFrame(input_start_date, input_end_date, input_etf_code)

        break
        
    except Exception as e:
        logger.exception(e)


# TODO 결과물 condition_df의 인덱스랑 거래일자 Datetime이 시분초까지 나오는 것 같은데 -> 확인
print("END")

# TODO 엑셀 파일 만들기
# TODO 클래스화
# TODO 날짜 입력 버튼 생성
# TODO Plotly로 출력
# TODO DB화 (MySQL)



# %%
# from tkinter import *
# root = Tk()
# root.geometry("600x400")
# root.title("TESTING")
# root.option_add("*Font", "맑은고딕 25")
# root.resizable(False, False)

# def btnpress():
#     btn.config(text = ent.get())

# ent = Entry(root)
# ent.pack()

# btn = Button(root)
# btn.config(text='버튼')
# btn.config(width=10)
# btn.config(command=btnpress)
# btn.pack()

# root.mainloop()

# "https://yeachan.tistory.com/6"
