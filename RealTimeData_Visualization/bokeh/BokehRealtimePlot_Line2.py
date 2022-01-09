# Streaming Data

import numpy as np
import pandas as pd
import xlwings as xw
from math import radians    # rotate axis ticks
from datetime import datetime
from bokeh.io import curdoc
from bokeh.layouts import layout
from bokeh.plotting import figure, output_file
from bokeh.models import ColumnDataSource, DatetimeTickFormatter, Select


# 엑셀 인스턴스(App) 생성
app = xw.App(visible=False)

# 엑셀 파일 불러오기
wb = xw.Book("C:\\Users\\Simons\\Desktop\\test_excel.xlsx")

# 첫 번째 시트 읽어오기
sheet = wb.sheets[0]

# 인스턴스 종료
app.kill()      # app을 kill 하지 않는 경우 "피호출자가 호출을 거부했습니다." 라는 에러가 발생함
                # 예상 원인 1: 엑셀이 정품이 아닌 경우
                # 예상 원인 2: 엑셀 변환 중에는 엑셀 문서 작업을 하지 않아야 하는데, 
                #              kill 하지 않는 경우 내부적으로 엑셀 작업이 진행될 가능성 존재

# Create Figure
p = figure(x_axis_type="datetime", width=900, height=350)

# Generate Data
def create_value1():
    # 데이터프레임 형태로 엑셀 시트 읽어오기
    val = sheet['A1'].value
    # np_df = df.to_numpy()
       
    return val

def create_value2():
    # 데이터프레임 형태로 엑셀 시트 읽어오기
    val = sheet['B1'].value
    # np_df = df.to_numpy()
       
    return val

# Create data source
source = ColumnDataSource(data=dict(
    x=[],
    y1=[],
    y2=[]
    )
)

# p.circle(x="x", y="y", color="firebrick", line_color="firebrick", source=source)
# p.line(x="x", y="y", source=source)
p.vline_stack(['y1', 'y2'], x='x', source = source)

# Create Periodic Function
def update():
    new_data = dict(x=[datetime.now()], y1=[create_value1()], y2=[create_value2()])
    print(new_data)
    # source.stream(new_data, rollover=200)
    source.stream(new_data)
    p.title.text = "Now Streaming %s Data" % select.value

# Callback Function (Update Function)
def update_intermed(attrname, old, new):
    source.data = dict(x=[], y1=[], y2=[])
    update()
    
# date_pattern = ["%Y-%m-%d\n%H:%M:%S"]
date_pattern = ["%H:%M:%S"]

p.xaxis.formatter = DatetimeTickFormatter(
    seconds = date_pattern,
    minsec = date_pattern,
    minutes = date_pattern,
    hourmin = date_pattern,
    hours  = date_pattern,
    days = date_pattern,
    months = date_pattern,
    years = date_pattern
)

p.xaxis.major_label_orientation = radians(50)
p.xaxis.axis_label = "Date"
p.yaxis.axis_label = "Value"

# Create Selection Widget
options = [("Stock1", "Stock One"), ("Stock2", "Stock Two")]
select = Select(title = "Market Name", value = "Stock1", options = options)
select.on_change("value", update_intermed)

# Configure Layout
lay_out = layout([[p], [select]])
curdoc().theme = "dark_minimal"
curdoc().title = "Streaming Stock Data Example"
curdoc().add_root(lay_out)
curdoc().add_periodic_callback(update, 100)

# bokeh serve --show bokeh_test.py
