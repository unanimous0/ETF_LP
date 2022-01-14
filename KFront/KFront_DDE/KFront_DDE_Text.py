import os
import time
import pandas as pd
import xlwings as xw

class KFront_DDE_Text:
    
    def __init__(self):
        pass

    def Make_Dict(self):
        self.app = xw.App(visible=False)
        self.wb = xw.Book("C:\\Users\\MP\\Downloads\\코스피_코스닥_종목코드_종목명.xlsx")
        self.sheet = self.wb.sheets[0]

        self.df_code_name = self.sheet.range('A1').options(pd.DataFrame, index=False, expand='table').value
        
        # 데이터프레임 -> 딕셔너리
        # self.df_code_name.set_index("종목코드")
        # self.dict_code_name = dict([(x, y) for x, y in zip(self.df_code_name["종목코드"], self.df_code_name['종목명'])])

        # 데이터프레임 -> 딕셔너리 (더 심플하게)
        self._dict_code_name = dict(self.df_code_name.values.tolist())
        
        self.app.kill()

        return self._dict_code_name


    def Make_DDE_Text(self, dict_code_name):
        self.dde_text_ticker = {
                                "현재가" :      "LAST",
                                "전일종가" :    "CLOSE_YST",
                                "매도1호가" :   "SELL1",
                                "매수1호가" :   "BUY1",
                                "시가" :        "OPEN",
                                "고가" :        "HIGH",
                                "저가" :        "LOW",
                                "전일대비" :    "CHANGE",
                                "등락률" :      "CHANGE_RATE",
                                "NAV" :        "NAV",
                                "전일NAV" :    "YST_NAV"
                             }

        self.dde_df = pd.DataFrame(columns=self.dde_text_ticker.values())
        
        for key in dict_code_name.keys():        
                
            self.dde_dict = {
                    "LAST"          : f"=KFront|DATA!'{key}:LAST'",
                    "CLOSE_YST"     : f"=KFront|DATA!'{key}:CLOSE_YST'",
                    "SELL1"         : f"=KFront|DATA!'{key}:SELL1'",
                    "BUY1"          : f"=KFront|DATA!'{key}:BUY1'",
                    "OPEN"          : f"=KFront|DATA!'{key}:OPEN'",
                    "HIGH"          : f"=KFront|DATA!'{key}:HIGH'",
                    "LOW"           : f"=KFront|DATA!'{key}:LOW'",
                    "CHANGE"        : f"=KFront|DATA!'{key}:CHANGE'",
                    "CHANGE_RATE"   : f"=KFront|DATA!'{key}:CHANGE_RATE'",
                    "NAV"           : f"=KFront|DATA!'{key}:NAV'",
                    "YST_NAV"       : f"=KFront|DATA!'{key}:YST_NAV'"
            }
            
            
            self.dde_df = self.dde_df.append(self.dde_dict, ignore_index=True)

        self.dde_df.index = dict_code_name.keys()

        return self.dde_df
    
    
    def Make_Excel(self, dde_df):
        dde_df.to_excel("KFRONT_DDE_CODE.xlsx")



if __name__ == "__main__":
    start_time = time.time()

    kdt = KFront_DDE_Text()
    dict_code_name = kdt.Make_Dict()
    dde_df = kdt.Make_DDE_Text(dict_code_name)
    kdt.Make_Excel(dde_df)
    
    end_time = time.time()
    running_time = end_time - start_time

    print(f"Total Runnig Time : {running_time:.2f} Seconds")

        


