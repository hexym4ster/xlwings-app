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
    
numlist()

def listcompare():
    global correct_list
    correct_list=[]
    for i in dianxin_numlist:
        if i in yuangong_numlist:
            correct_list.append(i)
            
listcompare()

def new_excel():
    app=xw.App(visible=False,add_book=False)
    new_wb=app.books.open(yuangong_xlpath)
    chongfunum_list=[]
    for i in yuangong_numlist:
        if i in correct_list:
            chongfunum_list.append('F'+str(yuangong_numlist.index(i)+2))
    for i in chongfunum_list:
        new_wb.sheets['sheet1'].range(i).value='佣金已返'
    
    new_path='e:\work\%s(加备注).xlsx'%os.path.split(yuangong_xlpath)[1][:-5]
    new_wb.save(new_path)
    new_wb.close()
    app.quit()                                                                                                                                                                                                                                         
    
new_excel()
print('完成！')