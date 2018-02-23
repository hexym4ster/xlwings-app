#号码列表整理函数
import os
import xlwings as xw
def numlist():
    global yuangong_numlist,dianxin_numlist,yuangong_xlpath,dianxin_xlpath
    yuangong_xlpath=input('请输入员工表格地址')
    dianxin_xlpath=input('请输入电信表格地址')
    app=xw.App(visible=False,add_book=False)
    app.display_alerts=False
    app.screen_updating=False
    yuangong_wb=app.books.open(yuangong_xlpath)
    yuangong_numlist=yuangong_wb.sheets['sheet1'].range('B2').expand('down').value
    dianxin_wb=app.books.open(dianxin_xlpath)
    dianxin_numlist=dianxin_wb.sheets['sheet1'].range('A3').expand('down').value
    yuangong_wb.close()
    dianxin_wb.close()
    app.quit()
    