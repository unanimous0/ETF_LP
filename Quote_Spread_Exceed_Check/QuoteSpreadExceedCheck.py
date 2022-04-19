import os
import time
from datetime import datetime
from numpy import empty
import xlwings as xw

path_dir = os.getcwd()

file_name = "호가 스프레드 초과 확인.xlsm"

book = xw.Book(path_dir + '/' + file_name)

sheet_sp_check = book.sheets["신고 스프레드 확인"]
sheet_records = book.sheets["기록"]
sheet_ongoing = book.sheets["40분 이상 초과"]

recording_dict = {}
comparison_list = []

ONGOING_CNT = 1
RECORDING_CNT = 0

# import threading
# TIME_CNT = 0
# # time_cnt 증가 (40분마다 증가)
# def increase_time_cnt():
#     global TIME_CNT
#     TIME_CNT += 1
#     # print(TIME_CNT)
#     threading.Timer(7, increase_time_cnt).start()
# increase_time_cnt()

while 1:    
    try:        
        comparison_list.clear()
        
        OVER_SPREAD = sheet_sp_check.range("M2:R450").value
        
        """
            * 컬럼명 정보
            기준 초과 = [0][0]
            단축 코드 = [0][1]
            종목 이름 = [0][2]
            신고 호가 = [0][3]
            현재 호가 = [0][4]
            초과 정도 = [0][5]
            
            * 나머지 데이터는 행 번호만 다름
        """
        
        # TODO 문제 1. 데이터의 값이 뜨는 영역이 계속 바뀜
        # TODO 문제 2. 초과된 경우의 시간이 뜨게하는 것으로 끝나는 건 의미 없음 -> 기록이 되어야함
        for j, row in enumerate(OVER_SPREAD):
            if row[0] != '':
                # print("Current Time --> ", datetime.now().strftime("%H:%M:%S"))
                # print(row)
                # sheet_sp_check.range(f"S{i+2}").value = datetime.now().strftime("%H:%M:%S")
                
                # 밑에서 비교 작업을 위해 한 회의 반복문에 있는 종목코드는 모두 리스트에 저장
                comparison_list.append(row[1])
                
                # dict로 현상태 저장
                if row[1] not in recording_dict:
                    recording_dict[row[1]] = datetime.now()
                    # recording_dict[row[1]] = datetime.now().strftime("%H:%M:%S")
                    # sheet_sp_check.range(f"S{j+2}").value = datetime.now().strftime("%H:%M:%S")    # 엑셀에는 처음 들어온 시간만 표시
                    
                else:
                    pass
                
            else:
                # 마지막 행 데이터면, 비교 작업 시작 (비교작업: dict에는 있는데 다음 반복문에서 해당 종목이 안들어오면 종료로 인식)
                # if j == len(OVER_SPREAD)-1:
                for etf in list(recording_dict.keys()):
                    if etf in comparison_list:
                        # TODO 해당 종목이 사라지진 않았지만, 기준을 초과하는 시간이 길어지는 경우도 표시해줘야함
                        #      예를 들어 초과된 시간이 30분이 지났지만 계속 초과된 상태라 아래 else문에 들어가지 않는 경우가 가능
                        #      따라서 일정 시간동안 초과되어있으면 해당 종목도 표시해줘야함
                        # TODO 일정 시간동안 초과된 종목을 표시해줬는데, 40분이 넘으면 40분 이후의 모든 반복문에서 엑셀에 기록함
                        #      따라서 "40분 * TIME_CNT"를 통해 40분 간격으로만 표시하도록 변경
                        # if (((datetime.now() - recording_dict[etf]).seconds)/60) >= 0.1 * TIME_CNT:      # 기준 스프레드를 초과한 시간이 40분 이상인 경우
                        if ((datetime.now() - recording_dict[etf]).seconds != 0) and ((datetime.now() - recording_dict[etf]).seconds % 2400 == 0) :
                            sheet_ongoing.range(f"A{ONGOING_CNT+1}").value = etf
                            sheet_ongoing.range(f"C{ONGOING_CNT+1}").value = recording_dict[etf].strftime("%H:%M:%S")
                            sheet_ongoing.range(f"D{ONGOING_CNT+1}").value = datetime.now().strftime("%H:%M:%S")
                            sheet_ongoing.range(f"E{ONGOING_CNT+1}").value = ((datetime.now() - recording_dict[etf]).seconds)/60

                            ONGOING_CNT += 1
                            # continue
                            
                    else:
                        # 이 경우가 dict에는 있는데 이번 반복문에서 해당 종목이 없는 경우 -> 종료시간 입력 후 전체 시간 확인 & 해당 종목 dict에서 제거
                        # 로그에 기록
                        sheet_records.range(f"A{RECORDING_CNT+2}").value = etf
                        sheet_records.range(f"C{RECORDING_CNT+2}").value = recording_dict[etf].strftime("%H:%M:%S")
                        sheet_records.range(f"D{RECORDING_CNT+2}").value = datetime.now().strftime("%H:%M:%S")
                        sheet_records.range(f"E{RECORDING_CNT+2}").value = ((datetime.now() - recording_dict[etf]).seconds)/60
                        # sheet_records.range(f"D{RECORDING_CNT+2}").value = sheet_records.range(f"C{RECORDING_CNT+2}").value - recording_dict[etf]
                        """
                            * 윗윗줄의 코드 VS 주석처리된 바로 위 코드
                              - 엄밀히 따지면 윗윗줄의 코드보다 주석처리된 바로 위 코드가 더 맞는 코드
                              - 왜냐하면 두 줄 모두 datetime.now()를 쓰고 있으므로, 엄밀히 따지면 두 개의 시간은 서로 같은 시간이 아니기 때문 (마이크로 초의 차이가 발생)
                              - 그러나 주석처리된 바로 위 코드는 실행 시, 오류 문제, 과부하 문제 등이 있고, 사용목적상 마이크로 초의 차이는 의미 없으므로 윗윗줄의 코드를 사용
                        """
                        
                        recording_dict.pop(etf)
                        # recording_dict.pop(etf, None)
                        
                        RECORDING_CNT += 1
                        
                break                    
        
        time.sleep(1)
        
        # print("WHILE문 끝")
        
    except Exception as e:
        print("CODE RAISED THE EXCEPTION AS FOLLOWS - ", str(e))