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
            ctypes.windll.user32.MessageBoxW(None, "리콜 상환할 파일(세이프 Recall조회)이 없습니다.", "알림", 0)

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
            ctypes.windll.user32.MessageBoxW(None, "리콜 상환할 파일(세이프 상환예정내역)이 없습니다.", "알림", 0)

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
    
    def build_df_self_repayment(self) -> pd.DataFrame():
        self.path_dir = os.getcwd()
        self.file_list = os.listdir(self.path_dir)

        # file_list에 리콜 상환 파일이 없는 경우
        # if "자체 상환 양식.xlsx" not in self.file_list:
        if "자체 상환 양식.xlsx" not in self.file_list:
            ctypes.windll.user32.MessageBoxW(None, "리콜 상환할 파일(자체상환파일)이 없습니다.", "알림", 0)
            
        else:
            df_self_repayment = pd.DataFrame()
            for file_name in self.file_list:
                if file_name == "자체 상환 양식.xlsx":
                    file_dir = self.path_dir + "\\" + file_name
                    df_self_repayment = self.read_self_repayment_file(file_dir)
                    
                    # 자체 상환 양식 - 종목코드 6자리로 변경 (A 있는 경우 A 삭제)
                    for i, val in enumerate(df_self_repayment["종목코드"]):
                        if val[0] == 'A':
                            val = val[1:7]
                            df_self_repayment["종목코드"].iloc[i] = val
                
        return df_self_repayment
                
                
    def read_self_repayment_file(self, file_dir) -> pd.DataFrame():
        _safe_repayment_file = xw.Book(file_dir)
        _df_safe_repayment = _safe_repayment_file.sheets(1).used_range.options(pd.DataFrame, index=False).value
        
        return _df_safe_repayment

    
    def process_self_repayment(self, df_self_repayment) -> pd.DataFrame():
        # 상환가능목록을 담기위한 빈 리스트
        input_list = []
    
        print("\n==================================== 동일 펀드에서의 상환가능여부를 탐색합니다. ====================================\n")

        for i, self_rp_row in df_self_repayment.iterrows():
            filter_fnd_stck_code = self.matched_table.loc[(self.matched_table["오피스펀드코드"] == self_rp_row["펀드코드"]) & (self.matched_table["종목코드"] == self_rp_row["종목코드"])]
            
            filter_fnd_stck_code = filter_fnd_stck_code.sort_values(by=["차입잔여수량", "체결일", "체결번호"], ascending=[True, True, True])
            # print(filter_fnd_stck_code[["오피스펀드코드", "오피스펀드명", "종목코드", "종목명", "차입잔여수량", "체결일", "체결번호", "대여자계좌", "대여자펀드명"]])
            """
                * 위와 같이 필터링했을 때 염두해야할 점
                    - 오피스펀드코드와 세이프대여자펀드명이 다를 수 있음
                    - 당연한 이유: 처음 차입할 때 차입하는 종목을 운용사가 꼭 동일 펀드에서 빌려주지는 않으므로
                    - 처음엔 이런 케이스가 있다면 같은 펀드 내에서 상환하는게 좋다고 생각했으나, 빌려온 곳에 그대로 갚는게 중요하므로 오피스펀드코드와 세이프대여자펀드명이 달라도 그대로 상환
            """
            
            filter_fsc_sum = filter_fnd_stck_code["차입잔여수량"].sum()

            # 1. 상환신청수량이 동일 펀드에 모두 있는 경우
            if filter_fsc_sum >= self_rp_row["상환신청수량"]:
                print(f"{self_rp_row['종목명']} 은(는) 동일 펀드 내에서 상환 가능합니다.\n")

                # 1-A. 상환이 가능하면 차입잔여수량이 작은 내역부터 체결일/체결번호 할당
                for j, ffsc_row in filter_fnd_stck_code.iterrows():
                    if self_rp_row["상환신청수량"] == 0:
                        break

                    # 1-A-a. 한 번에 상환이 안되는 경우
                    if self_rp_row["상환신청수량"] > ffsc_row["차입잔여수량"]:
                        repayment_partial_possible = {
                                                        "펀드코드"       : self_rp_row["펀드코드"], 
                                                        "펀드명"         : self_rp_row["펀드명"],
                                                        "종목코드"       : self_rp_row["종목코드"],
                                                        "종목명"         : self_rp_row["종목명"],
                                                        "상환수량"       : int(ffsc_row["차입잔여수량"]),
                                                        "체결일"         : ffsc_row["체결일"],
                                                        "체결번호"       : int(ffsc_row["체결번호"]),
                                                        "대여자계좌"     : ffsc_row["대여자계좌"],
                                                        "대여자펀드코드" : ffsc_row["대여자펀드코드"],
                                                        "대여자펀드명"   : ffsc_row["대여자펀드명"]
                                                     }
                        
                        input_list.append(repayment_partial_possible)

                        self_rp_row["상환신청수량"] -= ffsc_row["차입잔여수량"]
                        df_self_repayment.at[i, "상환신청수량"] = self_rp_row["상환신청수량"]

                        # TODO Matched Table의 row는 상환가능한 케이스이므로 해당 row 삭제

                    # 1-A-b. 한 번에 상환이 가능한 경우 or 부분 상환 후 상환신청수량이 0이 될 수 있는 경우
                    elif self_rp_row["상환신청수량"]  <= ffsc_row["차입잔여수량"]:
                        repayment_possible = {
                                                "펀드코드"       : self_rp_row["펀드코드"], 
                                                "펀드명"         : self_rp_row["펀드명"],
                                                "종목코드"       : self_rp_row["종목코드"],
                                                "종목명"         : self_rp_row["종목명"],
                                                "상환수량"       : int(self_rp_row["상환신청수량"]),
                                                "체결일"         : ffsc_row["체결일"],
                                                "체결번호"       : int(ffsc_row["체결번호"]),
                                                "대여자계좌"     : ffsc_row["대여자계좌"],
                                                "대여자펀드코드" : ffsc_row["대여자펀드코드"],
                                                "대여자펀드명"   : ffsc_row["대여자펀드명"]
                                             }

                        input_list.append(repayment_possible)

                        self_rp_row["상환신청수량"] -= self_rp_row["상환신청수량"]
                        df_self_repayment.at[i, "상환신청수량"] = self_rp_row["상환신청수량"]

                        # TODO Matched Table의 row는 상환가능한 케이스이므로 해당 row 삭제 또는 상환신청수량 만큼 차입잔여수량 차감
                        # ffsc_row["차입잔여수량"] -= self_rp_row["상환신청수량"]       # 이건 할 필요없음 -> 할거면 Matched_Table에 적용해야함

                # TODO 1-B. 차입잔여수량이 같은 내역이 있다면 체결일이 오래된 내역의 체결일/체결번호 할당
                # TODO 1-C. 차입잔여수량과 체결일까지 같은 내역이 있다면 체결번호가 작은 내역의 체결일/체결번호 할당
                #           -> filter_fnd_stck_code 을 생성할 때 정렬 순서를 위와 같이 만듦

                print(f"{self_rp_row['종목명']} 의 수량은 모두 동일 펀드 내에서 상환되었습니다.\n")

            elif filter_fsc_sum < self_rp_row["상환신청수량"]:
                # 2. 상환신청수량이 동일 펀드에 부분만 있는 경우 -> 부분만 상환 후 나머지는 동일 운용사, 다른 펀드에서 상환 (또는 차입 후 상환)
                if filter_fsc_sum != 0:
                    print(f"{self_rp_row['종목명']} 은(는) 동일 펀드 내에서 일부 수량만 상환 가능합니다.\n")
                    
                    # 2-A. 상환이 일부만 가능하면 동일 펀드에서 차입잔여수량이 작은 내역부터 체결일/체결번호 할당 후, 다른 펀드에서도 동일하게 할당
                    for j, ffsc_row in filter_fnd_stck_code.iterrows():
                        repayment_partial_possible = {
                                                    "펀드코드"       : self_rp_row["펀드코드"], 
                                                    "펀드명"         : self_rp_row["펀드명"],
                                                    "종목코드"       : self_rp_row["종목코드"],
                                                    "종목명"         : self_rp_row["종목명"],
                                                    "상환수량"       : int(ffsc_row["차입잔여수량"]),
                                                    "체결일"         : ffsc_row["체결일"],
                                                    "체결번호"       : int(ffsc_row["체결번호"]),
                                                    "대여자계좌"     : ffsc_row["대여자계좌"],
                                                    "대여자펀드코드" : ffsc_row["대여자펀드코드"],
                                                    "대여자펀드명"   : ffsc_row["대여자펀드명"]
                                                    }
                        
                        input_list.append(repayment_partial_possible)

                        self_rp_row["상환신청수량"] -= ffsc_row["차입잔여수량"]
                        df_self_repayment.at[i, "상환신청수량"] = self_rp_row["상환신청수량"]
                        
                        # 동일 펀드에 있는 부분 수량을 상환한 후, 그 부분 수량을 동일 펀드의 전체 부분 수량의 합에서 차감
                        # 계속 차감되다가 동일 펀드의 전체 부분 수량의 합이 0이 되면 동일 운용사, 다른 펀드로 넘어가도록 진행
                        filter_fsc_sum -= ffsc_row["차입잔여수량"]
                        if filter_fsc_sum == 0:
                            print(f"동일 펀드에 있는 {self_rp_row['종목명']} 의 수량은 모두 상환되었습니다. 동일 운용사, 다른 펀드에서 나머지 상환가능수량을 계속 탐색합니다.\n")
                            break

                # 3. 상환신청수량이 아예 동일 펀드에 없는 경우 -> 전수량을 동일 운용사, 다른 펀드에서 상환 (또는 차입 후 상환)
                else:
                    print(f"{self_rp_row['종목명']} 은(는) 동일 펀드 내에 수량이 없으므로 상환이 불가능합니다. 동일 운용사, 다른 펀드에서 나머지 상환가능수량을 계속 탐색합니다.\n")


        # 같은 펀드 내에서 상환 가능한 내역
        same_funds_repayment = pd.DataFrame(input_list, columns=["펀드코드", "펀드명", "종목코드", "종목명", "상환수량", "체결일", "체결번호", \
                                                            "대여자계좌", "대여자펀드코드", "대여자펀드명"])
        
        # 같은 운용사, 다른 펀드에서 상환 가능한 내역
        df_self_rp_other_funds = df_self_repayment[df_self_repayment["상환신청수량"] != 0]
        list_other_funds_rp, df_self_rp_error, other_funds_repayment = self.process_other_funds_repayment(df_self_rp_other_funds)
        for _dict in list_other_funds_rp:
            input_list.append(_dict)
            
        # 상환 가능한 전체 내역
        total_funds_repayment = pd.DataFrame(input_list, columns=["펀드코드", "펀드명", "종목코드", "종목명", "상환수량", "체결일", "체결번호", \
                                                            "대여자계좌", "대여자펀드코드", "대여자펀드명"])
        total_funds_repayment = total_funds_repayment.sort_values(by=["종목코드", "상환수량", "체결일"], ascending=[True, True, True])

        # 결과 출력
        print("\n========================================== 동일 펀드에서 상환 가능한 내역 ===========================================\n")
        print(same_funds_repayment)
        print("=====================================================================================================================\n")
        
        print("\n================================== 동일 펀드에서 상환하고 난 후의 자체상환 신청내역 =================================\n")
        print(df_self_repayment)
        print("=====================================================================================================================\n")

        print("\n==================================== 동일 운용사, 다른 펀드에서 상환 가능한 내역 ====================================\n")
        print(other_funds_repayment)
        print("=====================================================================================================================\n")
        
        print("\n=========================== 동일 운용사, 다른 펀드에서 상환하고 난 후의 자체상환 신청내역 ===========================\n")
        print(df_self_rp_other_funds)
        print("=====================================================================================================================\n")
        
        print("\n================================= Matched Table에서 상환 불가능한 자체상환 신청내역 =================================\n")
        print(df_self_rp_error)
        print("=====================================================================================================================\n")

        print("\n============================================== 자체상환 신청내역 결과 ===============================================\n")
        print(total_funds_repayment)
        print("=====================================================================================================================\n")
        
        
        if len(df_self_rp_error.index) != 0:
            ctypes.windll.user32.MessageBoxW(None, "[경고] Matched Table에서 상환이 안되는 종목이 있습니다.", "알림", 0)
            print("[Error] 자체상환 신청내역 중 상환 불가능한 내역이 있습니다.\n")
        else:
            print("[완 료] 자체상환 신청내역의 전 종목이 상환 가능하여, 모든 내역에 체결일/체결번호를 할당했습니다.\n")
        
        return same_funds_repayment
    
    
    def process_other_funds_repayment(self, df_self_rp_other_funds) -> dict():
         # 상환가능목록을 담기위한 빈 리스트
        input_list = []
        
        print("\n============================== 동일 운용사, 다른 펀드에서의 상환가능여부를 탐색합니다. =============================\n")
        
        # TODO 자체상환신청받은 내역에서 오피스펀드코드를 보고 펀드맵핑 클래스를 통해 어느 운용사인지 확인해야함 -> Matched Table의 대여자계좌의 앞 4자리 숫자 필요

        for i, self_rp_row in df_self_rp_other_funds.iterrows():
            filter_fnd_stck_code = self.matched_table.loc[(self.matched_table["오피스펀드코드"] != self_rp_row["펀드코드"]) & (self.matched_table["종목코드"] == self_rp_row["종목코드"]) \
                                                            & (self.matched_table["대여자계좌"].str.contains("3020"))]
            
            filter_fnd_stck_code = filter_fnd_stck_code.sort_values(by=["차입잔여수량", "체결일", "체결번호"], ascending=[True, True, True])
            
            filter_fsc_sum = filter_fnd_stck_code["차입잔여수량"].sum()
            
            # 1. 상환신청수량이 동일 운용사에 모두 있는 경우
            if filter_fsc_sum >= self_rp_row["상환신청수량"]:
                # 1-A. 상환이 가능하면 차입잔여수량이 작은 내역부터 체결일/체결번호 할당
                for j, ffsc_row in filter_fnd_stck_code.iterrows():
                    if self_rp_row["상환신청수량"] == 0:
                        break

                    # 1-A-a. 한 번에 상환이 안되는 경우
                    if self_rp_row["상환신청수량"] > ffsc_row["차입잔여수량"]:
                        repayment_partial_possible = {
                                                        "펀드코드"       : self_rp_row["펀드코드"], 
                                                        "펀드명"         : self_rp_row["펀드명"],
                                                        "종목코드"       : self_rp_row["종목코드"],
                                                        "종목명"         : self_rp_row["종목명"],
                                                        "상환수량"       : int(ffsc_row["차입잔여수량"]),
                                                        "체결일"         : ffsc_row["체결일"],
                                                        "체결번호"       : int(ffsc_row["체결번호"]),
                                                        "대여자계좌"     : ffsc_row["대여자계좌"],
                                                        "대여자펀드코드" : ffsc_row["대여자펀드코드"],
                                                        "대여자펀드명"   : ffsc_row["대여자펀드명"]
                                                     }
                        
                        input_list.append(repayment_partial_possible)

                        self_rp_row["상환신청수량"] -= ffsc_row["차입잔여수량"]
                        df_self_rp_other_funds.at[i, "상환신청수량"] = self_rp_row["상환신청수량"]

                        # TODO Matched Table의 row는 상환가능한 케이스이므로 해당 row 삭제

                    # 1-A-b. 한 번에 상환이 가능한 경우 or 부분 상환 후 상환신청수량이 0이 될 수 있는 경우
                    elif self_rp_row["상환신청수량"]  <= ffsc_row["차입잔여수량"]:
                        repayment_possible = {
                                                "펀드코드"       : self_rp_row["펀드코드"], 
                                                "펀드명"         : self_rp_row["펀드명"],
                                                "종목코드"       : self_rp_row["종목코드"],
                                                "종목명"         : self_rp_row["종목명"],
                                                "상환수량"       : int(self_rp_row["상환신청수량"]),
                                                "체결일"         : ffsc_row["체결일"],
                                                "체결번호"       : int(ffsc_row["체결번호"]),
                                                "대여자계좌"     : ffsc_row["대여자계좌"],
                                                "대여자펀드코드" : ffsc_row["대여자펀드코드"],
                                                "대여자펀드명"   : ffsc_row["대여자펀드명"]
                                             }

                        input_list.append(repayment_possible)

                        self_rp_row["상환신청수량"] -= self_rp_row["상환신청수량"]
                        df_self_rp_other_funds.at[i, "상환신청수량"] = self_rp_row["상환신청수량"]

                        # TODO Matched Table의 row는 상환가능한 케이스이므로 해당 row 삭제 또는 상환신청수량 만큼 차입잔여수량 차감
                        
                print(f"{self_rp_row['종목명']} 의 수량은 모두 동일 운용사, 다른 펀드 내에서 상환되었습니다.\n")
                
            elif filter_fsc_sum < self_rp_row["상환신청수량"]:
                # 2. 상환신청수량이 동일 운용사에 부분만 있는 경우 -> 부분만 상환 후 나머지는 동일 운용사, 다른 펀드에서 상환 (또는 차입 후 상환)
                if filter_fsc_sum != 0:
                    print(f"{self_rp_row['종목명']} 은(는) 동일 운용사 내에서 일부 수량만 상환 가능합니다.\n")
                    
                    # 2-A. 상환이 일부만 가능하면 동일 펀드에서 차입잔여수량이 작은 내역부터 체결일/체결번호 할당 후, 다른 펀드에서도 동일하게 할당
                    for j, ffsc_row in filter_fnd_stck_code.iterrows():
                        repayment_partial_possible = {
                                                    "펀드코드"       : self_rp_row["펀드코드"], 
                                                    "펀드명"         : self_rp_row["펀드명"],
                                                    "종목코드"       : self_rp_row["종목코드"],
                                                    "종목명"         : self_rp_row["종목명"],
                                                    "상환수량"       : int(ffsc_row["차입잔여수량"]),
                                                    "체결일"         : ffsc_row["체결일"],
                                                    "체결번호"       : int(ffsc_row["체결번호"]),
                                                    "대여자계좌"     : ffsc_row["대여자계좌"],
                                                    "대여자펀드코드" : ffsc_row["대여자펀드코드"],
                                                    "대여자펀드명"   : ffsc_row["대여자펀드명"]
                                                    }
                        
                        input_list.append(repayment_partial_possible)

                        self_rp_row["상환신청수량"] -= ffsc_row["차입잔여수량"]
                        df_self_repayment.at[i, "상환신청수량"] = self_rp_row["상환신청수량"]
                        
                        # 동일 펀드에 있는 부분 수량을 상환한 후, 그 부분 수량을 동일 펀드의 전체 부분 수량의 합에서 차감
                        # 계속 차감되다가 동일 펀드의 전체 부분 수량의 합이 0이 되면 동일 운용사, 다른 펀드로 넘어가도록 진행
                        filter_fsc_sum -= ffsc_row["차입잔여수량"]
                        if filter_fsc_sum == 0:
                            print(f"동일 운용사, 다른 펀드에 있는 {self_rp_row['종목명']} 의 수량은 모두 상환되었습니다.\n")
                            # self.process_other_funds_repayment(df_self_repayment)
                            break

                # 3. 상환신청수량이 아예 동일 펀드에 없는 경우 -> 전수량을 동일 운용사, 다른 펀드에서 상환 (또는 차입 후 상환)
                else:
                    print(f"{self_rp_row['종목명']} 은(는) 동일 운용사, 다른 펀드 내에 수량이 없으므로 상환이 불가능합니다. Unmatched Table을 탐색하거나 다른 운용사에서 차입한 내역을 확인하세요. (또는 차입후 상환)\n")
        
        
        df_self_rp_error = df_self_rp_other_funds[df_self_rp_other_funds["상환신청수량"] != 0]
        
        other_funds_repayment = pd.DataFrame(input_list, columns=["펀드코드", "펀드명", "종목코드", "종목명", "상환수량", "체결일", "체결번호", \
                                                            "대여자계좌", "대여자펀드코드", "대여자펀드명"])
        
        return (input_list, df_self_rp_error, other_funds_repayment)

    ################################################################################################################## 
    
    
if __name__ == "__main__":
    start_time = time.time()

    rp = Repayment_Process();

    # 세이프 Recall조회에서 상환 진행
    # df_safe_recall_search_repayment = rp.df_safe_recall_search_repayment()
    
    # 세이프 상환예정내역에서 상환 진행
    # df_safe_prearranged_repayment = rp.df_safe_prearranged_repayment()

    # 자체 상환 진행
    df_self_repayment = rp.build_df_self_repayment()
    check_repayment = rp.process_self_repayment(df_self_repayment)

    # 객체 제거
    del rp

    end_time = time.time()
    running_time = end_time - start_time

    print(f"Total Runnig Time : {running_time:.2f} Seconds")
