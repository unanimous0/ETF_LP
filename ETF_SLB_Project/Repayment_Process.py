# -*- coding: utf-8 -*-
"""
Created on Thu Feb 10 16:28:50 2022

@author: SE21297
"""

# 리콜 상환과 자체 상환에 따른 Matched Data와 비교

from distutils.command import check
import os
import time
import ctypes
import xlwings as xw
import pandas as pd

class Repayment_Process:
    def __init__(self):
        self.app = xw.App(visible=False)

        self.matched_table   = xw.Book("matched9088건.xlsx").sheets(1).used_range.options(pd.DataFrame, index=False).value
        # self.unmatched_table = xw.Book("unmatched1199건.xlsx").sheets(1).used_range.options(pd.DataFrame, index=False).value
        self.office_5264     = xw.Book("5264화면_20220126.xls").sheets(1).used_range.options(pd.DataFrame, index=False).value
    

    def __del__(self):
        self.app.kill()
    
    ######################################## 리콜 상환 - 세이프 "Recall조회" 화면 ########################################

    def df_safe_recall_search_repayment(self):
        self.path_dir = os.getcwd()
        self.file_list = os.listdir(self.path_dir)

        # file_list에 리콜 상환 파일이 없는 경우
        if "세이프 Recall조회 양식.xlsx" not in self.file_list:
            ctypes.windll.user32.MessageBoxW(None, "리콜 상환할 파일(\"세이프 Recall조회\")이 없습니다.", "알림", 0)

        for file_name in self.file_list:
            if file_name == "세이프 Recall조회 양식.xlsx":
                file_dir = self.path_dir + '/' + file_name
                df_safe_repayment1 = self.read_safe_recall_search_repayment_file(file_dir)
                
                # 리콜 상환 양식 - 거래유형 확인
                for i, val in enumerate(df_safe_repayment1["거래유형"]):
                    if val != "지정":
                        ctypes.windll.user32.MessageBoxW(None, "리콜 들어온 종목의 거래유형이 '지정거래'가 아닙니다.", "알림", 0)            
                        break                    
                
                # 리콜 상환 양식 - 종목코드 6자리로 변경
                for i, val in enumerate(df_safe_repayment1["종목코드"]):
                    if val[:2] == "KR":
                        val = val[3:9]
                        df_safe_repayment1["종목코드"].iloc[i] = val
                        
                return df_safe_repayment1
                
                
    def read_safe_recall_search_repayment_file(self, file_dir):
        _safe_repayment_file1 = xw.Book(file_dir)
        _df_safe_repayment1 = _safe_repayment_file1.sheets(1).used_range.options(pd.DataFrame, index=False).value
        
        return _df_safe_repayment1
    
    ####################################################################################################################
    
   
    
  
    ####################################### 리콜 상환 - 세이프 "상환예정내역" 화면 #######################################

    def df_safe_prearranged_repayment(self):
        self.path_dir = os.getcwd()
        self.file_list = os.listdir(self.path_dir)

        # file_list에 리콜 상환 파일이 없는 경우
        if "세이프 상환예정내역 양식.xlsx" not in self.file_list:
            ctypes.windll.user32.MessageBoxW(None, "리콜 상환할 파일(\"세이프 상환예정내역\")이 없습니다.", "알림", 0)

        for file_name in self.file_list:
            if file_name == "세이프 상환예정내역 양식.xlsx":
                file_dir = self.path_dir + '/' + file_name
                df_safe_repayment2 = self.read_safe_prearranged_repayment_file(file_dir)
                
                # 리콜 상환 양식 - 거래유형 확인
                for i, val in enumerate(df_safe_repayment2["거래유형"]):
                    if val != "지정":
                        ctypes.windll.user32.MessageBoxW(None, "리콜 들어온 종목의 거래유형이 '지정거래'가 아닙니다.", "알림", 0)            
                        break                    
                
                # 리콜 상환 양식 - 종목코드 6자리로 변경
                for i, val in enumerate(df_safe_repayment2["종목코드"]):
                    if val[:2] == "KR":
                        val = val[3:9]
                        df_safe_repayment2["종목코드"].iloc[i] = val
                        
                return df_safe_repayment2
                
                
    def read_safe_prearranged_repayment_file(self, file_dir):
        _safe_repayment_file2 = xw.Book(file_dir)
        _df_safe_repayment2 = _safe_repayment_file2.sheets(1).used_range.options(pd.DataFrame, index=False).value
        
        return _df_safe_repayment2
    
    ##################################################################################################################
    
    
    
       
    #################################################### 자체 상환 ####################################################
    
    def df_self_repayment(self):
        self.path_dir = os.getcwd()
        self.file_list = os.listdir(self.path_dir)

        # file_list에 리콜 상환 파일이 없는 경우
        # if "자체 상환 양식.xlsx" not in self.file_list:
        if self.file_list[0] not in self.file_list:
            ctypes.windll.user32.MessageBoxW(None, "리콜 상환할 파일(\"자체상환파일\")이 없습니다.", "알림", 0)

        for file_name in self.file_list:
            if file_name == "자체 상환 양식.xlsx":
                file_dir = self.path_dir + '/' + file_name
                df_self_repayment = self.read_self_repayment_file(file_dir)
                
                # 자체 상환 양식 - 종목코드 6자리로 변경 (A 있는 경우 A 삭제)
                for i, val in enumerate(df_self_repayment["종목코드"]):
                    if val[0] == 'A':
                        val = val[1:7]
                        df_self_repayment["종목코드"].iloc[i] = val
                        
                return df_self_repayment
            
        # file_list에 리콜 상환 파일이 없는 경우
        ctypes.windll.user32.MessageBoxW(None, "자체 상환할 파일이 없습니다.", "알림", 0)
                
                
    def read_self_repayment_file(self, file_dir):
        _safe_repayment_file = xw.Book(file_dir)
        _df_safe_repayment = _safe_repayment_file.sheets(1).used_range.options(pd.DataFrame, index=False).value
        
        return _df_safe_repayment
    
    
    def process_self_repayment(self, df_self_repayment):
        # 상환가능여부 및 기타정보 추가하기 위한 데이터프레임 생성
        # 아래에서 상환가능하면, 상환요청양식에 Matched Table의 ["체결일", "체결번호", "대여자계좌", "대여자펀드코드", "대여자펀드명"] 과 ["상환가능여부"] 달아주기
        self.check_repayment = pd.DataFrame(columns=["펀드코드", "펀드명", "종목코드", "종목명", "상환수량", "상환가능여부", "체결일", "체결번호", "대여자계좌", "대여자펀드코드", "대여자펀드명"])

        """
            # 자체상환양식에 있는 종목과 수량이 Matched Table의 SUM값과 비교해서 상환가능한지부터 파악 후, 가능하면 하나씩 체결일/체결번호를 할당해야함
            # 이때 SUM값은 1. 같은 펀드 -> 1번에서 False가 나오면 2. 같은 운용사내 다른 펀드의 SUM값을 비교해서, 하나씩 체결일/체결번호를 할당해야함
            # 이렇게 Matched Table에서 해당 종목의 수량을 비교하는 1번과 2번 기능을 위한 함수 따로 생성
        """

        # 1. 펀드코드와 종목코드로 동일 펀드내에서 상환가능한지 확인
        fnd_stck_code_TF = self.check_fnd_stck_code(df_self_repayment)
        

        # 자체상환에는 펀드코드가 지정되어있으므로 펀드코드와 종목코드로 비교
        for i, row in df_self_repayment.iterrows():
            filter_fnd_stck_code = self.matched_table.loc[(self.matched_table["오피스펀드코드"] == row["펀드코드"]) & (self.matched_table["종목코드"] == row["종목코드"])]
            
            remained_cnt = filter_fnd_stck_code["차입잔여수량"].iloc[i]
            
            """
                * 자체 상환은 이미 잔고에 있는 수량을 상환하므로, 바로 Matched Table의 차입잔여수량과 비교하면 됨
                # TODO 단순히 비교 불가 -> Matched Table에서 하나의 행에서는 같은 펀드에 같은 종목이 있어도 수량이 부족할 수 있음
                # TODO              -> 그런데 같은 펀드에서 동일 종목을 다른 날짜에 차입해왔다면 그 수량을 합하면 상환가능함
                # TODO              -> 따라서 하나의 행만 비교하면 안되고 SUM 값으로 비교해야함
                # TODO              -> 그리고 하나씩 상환하고 다음 행으로 넘어갈 때는 상환수량의 수가 그 행의 차입수량만큼 줄어들어야함
                # TODO              -> 반대로 SUM값이 상환하려는 수량보다 크다면 마지막 행의 차입잔여수량은 상환수량만큼 줄어들어야함
                # TODO              -> 물론 차입잔여수량이 작은 행들은 상환되면서 사라지게 만들면 됨 (이게 모두 Matched Table 업데이트 과정임)
                # TODO 이렇게 SUM 값으로 봐도 없는 경우에만 동일 운용사, 다른 펀드로 넘어가면됨
                # TODO              -> 동일 운용사를 돌때도 동일하게 SUM 값으로 비교하고 넘어가면서 차입수량 빼고 해야함
            """
            
            # 1. 상환하려는 수량이 모두 동일 펀드 내에 있는 경우
            if (remained_cnt >= row["상환수량"]):
                print("[상환 가능합니다.")
                var_check_repayment = "상환가능"

                # check_repayment 데이터프레임에 추가
                repayment_possible = [row["펀드코드", row["펀드명"], row["종목코드"], row["종목명"], row["상환수량"], var_check_repayment, self.matched_table["체결일"], \
                                      self.matched_table["체결번호"], self.matched_table["대여자계좌"], self.matched_table["대여자펀드코드"], self.matched_table["대여자펀드명"]]]
                self.check_repayment.append(repayment_possible)
                
            # 2. 상환하려는 수량이 동일 운용사의 다른 펀드에 있는 경우
            # TODO 동일 펀드에 부분수량이 있고 나머진 다른 펀드에 있다면, 동일 펀드에서 상환가능한 부분은 상환해야함)   
            elif (True):
                pass

            # 3. 상환하려는 수량이 동일 운용사 펀드들의 합보다 큰 경우 -> ERROR
            # TODO 그래도 가능한 만큼은 상환해야함 -> 상환가능수량 및 부족수량 모두 보여줄것)   
            else:
                # print("ERROR - 잔고의 수량이 Matched Table의 차입잔여수량보다 많습니다. (차입잔여수량 < 잔고수량)")
                ctypes.windll.user32.MessageBoxW(None, "ERROR - 잔고의 수량이 Matched Table의 차입잔여수량보다 많습니다. (차입잔여수량 < 잔고수량)", "알림", 0)
                break

        return self.check_repayment

    
    def check_fnd_stck_code(self, df_self_repayment):
        # 상환가능목록을 담기위한 빈 리스트
        input_list = []

        for i, self_rp_row in df_self_repayment.iterrows():
            filter_fnd_stck_code = self.matched_table.loc[(self.matched_table["오피스펀드코드"] == self_rp_row["펀드코드"]) & (self.matched_table["종목코드"] == self_rp_row["종목코드"])]
            filter_fsc_sum = filter_fnd_stck_code["차입잔여수량"].sum()

            # 1. 상환신청수량이 동일 펀드에 모두 있는 경우
            if filter_fsc_sum >= self_rp_row["상환수량"]:
                print(f"{self_rp_row['종목명']} 은(는) 동일 펀드 내에서 상환 가능합니다.")
                check_result = True

                filter_fnd_stck_code = filter_fnd_stck_code.sort_values(by=["차입잔여수량", "체결일", "체결번호"], ascending=[True, True, True])
                # print(filter_fnd_stck_code[["오피스펀드코드", "오피스펀드명", "종목코드", "종목명", "차입잔여수량", "체결일", "체결번호", "대여자계좌", "대여자펀드명"]])
                """
                    * 위와 같이 필터링했을 때 염두해야할 점
                      - 오피스펀드코드와 세이프대여자펀드명이 다를 수 있음
                      - 당연한 이유: 처음 차입할 때 차입하는 종목을 운용사가 꼭 동일 펀드에서 빌려주지는 않으므로
                      - 처음엔 이런 케이스가 있다면 같은 펀드 내에서 상환하는게 좋다고 생각했으나, 빌려온 곳에 그대로 갚는게 중요하므로 오피스펀드코드와 세이프대여자펀드명이 달라도 그대로 상환
                """


                # 1-A. 상환이 가능하면 차입잔여수량이 작은 내역부터 체결일/체결번호 할당
                for j, ffsc_row in filter_fnd_stck_code.iterrows():

                    if self_rp_row["상환수량"] == 0:
                        break

                    # 1-A-a. 한 번에 상환이 안되는 경우
                    if self_rp_row["상환수량"] > ffsc_row["차입잔여수량"]:
                        repayment_partial_possible = {
                                                        "펀드코드"      : self_rp_row["펀드코드"], 
                                                        "펀드명"        : self_rp_row["펀드명"],
                                                        "종목코드"      : self_rp_row["종목코드"],
                                                        "종목명"        : self_rp_row["종목명"],
                                                        "상환수량"      : int(ffsc_row["차입잔여수량"]),
                                                        "체결일"        : ffsc_row["체결일"],
                                                        "체결번호"      : int(ffsc_row["체결번호"]),
                                                        "대여자계좌"     : ffsc_row["대여자계좌"],
                                                        "대여자펀드코드"  : ffsc_row["대여자펀드코드"],
                                                        "대여자펀드명"   : ffsc_row["대여자펀드명"]
                                                     }
                        
                        input_list.append(repayment_partial_possible)

                        self_rp_row["상환수량"] -= ffsc_row["차입잔여수량"] 

                        # TODO Matched Table의 row는 상환가능한 케이스이므로 해당 row 삭제

                    # 1-A-b. 한 번에 상환이 가능한 경우 or 부분 상환 후 상환수량이 0이 될 수 있는 경우
                    elif self_rp_row["상환수량"]  <= ffsc_row["차입잔여수량"]:
                        repayment_possible = {
                                                "펀드코드"      : self_rp_row["펀드코드"], 
                                                "펀드명"        : self_rp_row["펀드명"],
                                                "종목코드"      : self_rp_row["종목코드"],
                                                "종목명"        : self_rp_row["종목명"],
                                                "상환수량"      : int(self_rp_row["상환수량"]),
                                                "체결일"        : ffsc_row["체결일"],
                                                "체결번호"      : int(ffsc_row["체결번호"]),
                                                "대여자계좌"     : ffsc_row["대여자계좌"],
                                                "대여자펀드코드"  : ffsc_row["대여자펀드코드"],
                                                "대여자펀드명"   : ffsc_row["대여자펀드명"]
                                             }

                        input_list.append(repayment_possible)

                        self_rp_row["상환수량"] -= self_rp_row["상환수량"]

                        # TODO Matched Table의 row는 상환가능한 케이스이므로 해당 row 삭제 또는 상환신청수량 만큼 차입잔여수량 차감
                        # ffsc_row["차입잔여수량"] -= self_rp_row["상환수량"]       # 이건 할 필요없음 -> 할거면 Matched_Table에 적용해야함

                # 1-B. 차입잔여수량이 같은 내역이 있다면 체결일이 오래된 내역의 체결일/체결번호 할당

                # 1-C. 차입잔여수량과 체결일까지 같은 내역이 있다면 체결번호가 작은 내역의 체결일/체결번호 할당

            else:
                # 2. 상환신청수량이 동일 펀드에 부분만 있는 경우
                pass

                # 3. 상환신청수량이 아예 동일 펀드에 없는 경우

        check_repayment = pd.DataFrame(input_list, columns=["펀드코드", "펀드명", "종목코드", "종목명", "상환수량", "체결일", "체결번호", \
                                                            "대여자계좌", "대여자펀드코드", "대여자펀드명"])
        print(check_repayment)

        
    
    ################################################################################################################## 
    
    
if __name__ == "__main__":
    start_time = time.time()

    rp = Repayment_Process();

    # 세이프 Recall조회에서 상환 진행
    # df_safe_recall_search_repayment = rp.df_safe_recall_search_repayment()
    
    # 세이프 상환예정내역에서 상환 진행
    # df_safe_prearranged_repayment = rp.df_safe_prearranged_repayment()

    # 자체 상환 진행
    df_self_repayment = rp.df_self_repayment()
    check_repayment = rp.process_self_repayment(df_self_repayment)

    # 객체 제거
    del rp

    end_time = time.time()
    running_time = end_time - start_time

    print(f"Total Runnig Time : {running_time:.2f} Seconds")