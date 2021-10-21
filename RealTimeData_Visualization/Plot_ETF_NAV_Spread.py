# -*- coding: utf-8 -*-
"""
Created on Mon Oct 18 13:51:51 2021
@author: Eunhwan Koh
Ref url: https://training.xlwings.org/courses/270054/lectures/4299644
"""

#%%
import xlwings as xw

# sheet = xw.books.active.sheets.active         # 현재 열려있는 엑셀 창을 인식

"""
    * 엑셀 창을 열지 않고 xlwings 를 사용하는 코드 패턴
    * python xlwings 라이브러리를 이용해서 엑셀에 데이터를 읽고 쓸 때, 기본적인 세팅은 엑셀 창이 열리도록 되어있음
    * 반면 엑셀 창을 띄워서 작업하지 않는 상황도 필요 (ex. 간단히 엑셀 파일에서 데이터만 읽어오면 되는 상황 등)
    * * xlwings의 객체는 크게 네 가지 유형이 존재 (객체의 계층 구조 순서대로 나열하면 다음과 같음)
    * * 1. App (엑셀 인스턴스)
    * * 2. Book
    * * 3. Sheet
    * * 4. Range
    * * * * 보통의 경우 xw.Book(파일명) 의 방식으로 엑셀 파일을 읽어왔으나, 엑셀 창 없이 읽기 위해서는 App 객체부터 생성해야함
    출처: https://hogni.tistory.com/58
"""

# 엑셀 인스턴스(App) 생성
app = xw.App(visible=False)

# 엑셀 파일 불러오기
wb = xw.Book("C:\\Users\\Check\\Desktop\\Check_ETF_NAV_RTD.xlsx")

# 첫 번째 시트 읽어오기
sheet = wb.sheets[0]

# # 데이터프레임 형태로 엑셀 시트 읽어오기
# df = sheet.ranne('A1').options(pd.DataFrame, index=False, expand='table').value

# 인스턴스 종료
app.kill()      # app을 kill 하지 않는 경우 "피호출자가 호출을 거부했습니다." 라는 에러가 발생함
                # # 예상 원인 1: 엑셀이 정품이 아닌 경우
                # # 예상 원인 2: 엑셀 변환 중에는 엑셀 문서 작업을 하지 않아야 하는데, kill 하지 않는 경우 내부적으로 엑셀 작업이 진행될 가능성 존재


#%%
import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation
from datetime import datetime


#%%
plt.style.use('bmh')

fig, ax = plt.subplots(3, 1)

ax2 = ax[0].twinx()     # animate_1에서 이중 y축 사용을 위해 먼저 설정  

xdata1 = []         # subplot에서 x축은 모두 공통

# animate_1
ydata1 = []
ydata11 = []

# animate_2
# xdata2 = []       # x축은 동일함
ydata2 = []
ydata22 = []

# animate_3
ydata3 = []


#%%
# ETF iNAV와 현재가를 두 개의 Y축에 대하여 나타냄
# # 이렇게 나타내면 스프레드를 볼 수 없음
def animate_1(i):
    y = sheet['B6'].value       # ETF iNAV
    y2 = sheet['B4'].value      # ETF 현재가

    # xdata1.append(i)
    xdata1.append(datetime.now().strftime("%H:%M:%S.%f")[:-3])      # 시:분:초:ms(3자리까지)
    ydata1.append(y)
    ydata11.append(y2)
    
    ax[0].cla()
    ax2.cla()
    
    ax[0].plot(xdata1, ydata1, color='red')

    ax2.plot(xdata1, ydata11, color='blue')
    ax2.grid(None)                               # twinx로 y축을 2개 쓸 경우 gird가 중복되므로 하나 제거

    ax[0].tick_params(axis='x', rotation=45)     # x축 레이블 각도 조절 (axis='x' 없으면 x축과 y축 레이블 모두 각도가 바뀜)


ani_1 = FuncAnimation(plt.gcf(), animate_1, interval=0.001)


#%%
# ETF iNAV와 현재가를 하나의 Y축에 대하여 나타냄
# # 이렇게 해야 스프레드를 정확히 볼 수 있음
# # 첫 번째 방식으로하면 현재가 대비 ETF NAV의 움직임을 정확히 파악할 수 없음 (스프레드가 아니므로)
def animate_2(i):
    y = sheet['B6'].value       # ETF iNAV
    y2 = sheet['B4'].value      # ETF 현재가

    ydata2.append(y)
    ydata22.append(y2)
    
    ax[1].cla()
    
    ax[1].plot(xdata1, ydata2, color='red')
    ax[1].plot(xdata1, ydata22, color='blue')

    ax[1].tick_params(axis='x', rotation=45)

ani_2 = FuncAnimation(plt.gcf(), animate_2, interval=0.001)


#%%
# ETF iNAV와 현재가의 차이, 즉 스프레드를 나타냄

def animate_3(i):
    y = sheet['B7'].value       # ETF iNAV

    ydata3.append(y)
    
    ax[2].cla()
    
    ax[2].plot(xdata1, ydata3, color='orange')

    ax[2].tick_params(axis='x', rotation=45)


ani_3 = FuncAnimation(plt.gcf(), animate_3, interval=0.001)


#%%
fig.tight_layout()
plt.show()
