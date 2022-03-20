import PySimpleGUI as sg
import pandas as pd
import datetime



sg.theme('Green')   # 设置当前主题
layout = [            
            [sg.FileBrowse('选择日报文件',key = '日报路径'),sg.Text()],
            [sg.CalendarButton('选择查询日期：'),sg.Input(key = '日期')],
            [sg.Multiline('查询姓名（每行一个，双击按钮生效！）\n打开软件初次查询会加载日报\nby F.W.Yue',size=(80,15), key='Names')],
            [sg.Button('查询排班')],
            [sg.Output(size=(80, 20))],
        ]
# 创造窗口
window = sg.Window('排班表查询 beta 版', layout)

#读取排班表：
def get_paiban():
    event, values = window.read()
    ribaopath = values['日报路径']
    df_paiban = pd.read_excel(ribaopath,sheet_name=4)
    return df_paiban
    
def get_banci(df_paiban,NameList):
    event, values = window.read()
    date = values['日期']
    date = date.split(' ')[0]
    Date = datetime.datetime.strptime(date, "%Y-%m-%d")
    for name in NameList:
        DF_paiban_row = df_paiban[df_paiban['姓名'].isin([name])]
        try:
            banci = DF_paiban_row[Date].iloc[0]
        except:
            banci = '/'
        print(name,',',banci)

df_paiban = get_paiban()

while True:
    event, values = window.read()
    names = values['Names']
    nameList = names.split('\n')
    get_banci(df_paiban,nameList)
    print('*'*8)
