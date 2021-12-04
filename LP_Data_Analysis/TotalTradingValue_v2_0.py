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
# import TTVDefSet_v2_0 as tvds
from datetime import datetime

logger = logging.getLogger()

"""
    # 파일 인코딩 확인
    import chardet
    rawdata = open(file_path, 'rb').read()
    result = chardet.detect(rawdata)
    charenc = result['encoding']
    print(charenc)
"""

class ETFLPTradingValueDF:
    # def __init__(self, input_start_date='20211119', input_etf_code='KR7069500007'):
    def __init__(self):
        # 조회하고자 입력받은 시작일/기준일
        # self.input_start_date = input_start_date

        # # 조회하고자 입력받은 ETF 종목 코드
        # self.input_etf_code = input_etf_code

        # # 조회하고자 입력받은 조회 기간
        # self.input_period = '1D'      # '1W', '1M', '1Q', '1Y'

        # 파일 경로 확인 및 raw data 데이터프레임 생성
        # self.file_path = self.setPath(self.input_start_date)        # 취합본 또는 DB 사용하면 경로설정을 위한 날짜는 필요없음
        self.file_path = self.setPath()

        # 초기 데이터프레임 생성
        self.raw_data_df = getRawDF(self.file_path)

        # 중복제거된 거래일자 리스트
        self.date_list = getTradingDateList(self.raw_data_df)

        # 딕셔너리 {종목코드:종목명} 생성
        self.etf_code_name_dict = getETFCodeNameDict(self.raw_data_df)

        # raw data 데이터프레임 재구성 및 총거래대금 column 추가 (순매수/순매도대금 아닌 총거래대금)
        self.df_with_7cols = rebuildRawDF(self.raw_data_df)

        # KB증권을 포함한 주요 회원사명 리스트 생성
        self.lpcompanies = ['KB증권', 'NH투자증권', '신한투자', '메리츠', '키움증권']

        # 주요 회원사 데이터만 추출
        self.main_lpcomp_df = getMainLPCompDF(self.df_with_7cols, self.lpcompanies)

        # 컬럼과 각 컬럼에 해당하는 데이터가 섞이지 않도록 딕셔너리 생성 (회원사명 순서 바뀌는 오류 방지)
        self.lpcomp_value_dict = getLPCompValueDict(self.lpcompanies)

        # 속도 향상을 위해 데이터프레임 numpy array로 변형
        # np_df = df.to_numpy()

        # 최종 결과 출력을 위해 필요한 팩터들을 다 갖고 있는 딕셔너리 생성
        self.code_lpcomp_date_value_dict = tvds.getCodeLPCompDateValueDict(self.etf_code_name_dict, 
                                                                           self.main_lpcomp_df,
                                                                           self.lpcompanies)

        self.empty_total_value_df = getEmptyTotalTradingValueDF(self.lpcompanies)

        # 이 fillTotalTraidngValueDF 함수까지가 취합본 또는 DB에서 읽어온 데이터를 팀장님이 원하는 형식의 데이터프레임으로 만드는 과정 
        # 이 함수가 이 클래스의 목적
        # 이 함수 이후에 원하는 날짜와 종목코드를 입력하는 작업이 진행됨 (클래스 밖에서) (이 클래스 함수까진 날짜나 종목코드 필요 없음)
        self.total_value_df = fillTotalTradingValueDF(self.date_list, 
                                                           self.code_lpcomp_date_value_dict, 
                                                           self.etf_code_name_dict, 
                                                           self.lpcomp_value_dict, 
                                                           self.lpcompanies, 
                                                           self.empty_total_value_df)

        def returnTotalTradingValueDF(self):
            return self.total_value_df

        
if __name__ == "__main__":
    # input_start_date = input("시작일(yyyymmdd): ")
    # input_end_date = input("종료일(yyyymmdd): ")
    # input_etf_code = input("ETF 종목 코드: ")
    # input_start_date = '20211119'
    # input_end_date = '20211123'
    # input_etf_code = 'KR7069500007'

    tvdf = ETFLPTradingValueDF()
    total_value_df = tvdf.returnTotalTradingValueDF()


    """
    # 1. 시작일과 종료일로 입력받은 경우의 함수 호출
    # 2. 시작일과 그 시작일로부터의 기간(1D, 1W, 1M, 1Q, 1Y 등)을 입력받은 경우의 함수 호출
    # 지금은 단순히 1번과 2번으로 입력받아서 구분하지만, 
    # 나중엔 한 화면에서 이벤트가 기간 또는 기준일 중 어디에 발생했을 때, 
    # 그에 맞는 콜백함수가 호출되도록 구현해야함
"""

    while True:
        try:
            input_etf_code = input("조회하고자하는 ETF 종목 코드를 입력해주세요. -> ")

            if len(input_etf_code) != 12:
                raise InputValueError("ETF 종목 코드를 올바르게 입력하세요.")

            search_date_condition = input("1. 기간 조회, 2. 기준일 조회 -> 1과 2중 하나를 입력하세요. -> ")

            if search_date_condition == '1':
                print("하루만 조회하고 싶은 경우에도 종료일을 입력해주세요.")
                input_start_date = input("시작일 입력 [yyyymmdd] -> ")
                input_end_date = input("종료일 입력 [yyyymmdd] -> ")
                # TODO ETFLPTradingValueDF 클래스 객체 생성
                condition_df = makeConditionalDF(start_date=input_start_date, end_date=input_end_date, etf_code=input_etf_code, total_value_df=total_value_df)

                break

            elif search_date_condition == '2':
                input_start_date = input("시작일/기준일 입력 [yyyymmdd] -> ")
                
                input_period = input("""
                기준일로부터 조회하고자하는 기간을 선택하세요
                    1. 1 Day
                    2. 1 Week
                    3. 1 Month
                    4. 1 Quarter
                    5. 1 Year (Not yet)
                    ->
                """)

                if input_period == '1':
                    period = '1D'
                elif input_period == '2':
                    period = '1W'
                elif input_period == '3':
                    period = '1M'
                elif input_period == '4':
                    period = '1Q'
                elif input_period == '5':
                    period = '1Y'
                else:
                    raise InputValueError("1부터 5중의 숫자 중 하나를 올바르게 입력하세요.")     


                # TODO ETFLPTradingValueDF 클래스 객체 생성
                condition_df = makeConditionalDF(start_date=input_start_date, etf_code=input_etf_code, total_value_df=total_value_df, period=period)

                break

            else:
                raise InputValueError("It needs nothing except 1 or 2.")

        except Exception as e:
            logger.exception(e)

    
    print("/n+++++ 최종 결과물 condition_df 반환 완료 +++++/n")
