import os #导入os模块获取最后级别的文件名，用来给新生成的excel命名
import xlwings as xw #导入xlwings模块

def numlist(): #生成两个列表：员工放号和电信返佣金号码列表
    global yuangong_numlist,dianxin_numlist,yuangong_xlpath,dianxin_xlpath #定义全局变量
    yuangong_xlpath=input('请输入员工表格地址')
    dianxin_xlpath=input('请输入电信表格地址')
    app=xw.App(visible=False,add_book=False)
    app.display_alerts=False
    app.screen_updating=False
    yuangong_wb=app.books.open(yuangong_xlpath)
    yuangong_numlist=yuangong_wb.sheets['sheet1'].range('B2').expand('down').value #从B2开始向下员工放号号码存入列表
    dianxin_wb=app.books.open(dianxin_xlpath)
    dianxin_numlist=dianxin_wb.sheets['sheet1'].range('A3').expand('down').value #从A3开始向下电信返佣金号号码存入列表
    yuangong_wb.close()
    dianxin_wb.close()
    app.quit()

numlist()

def listcompare(): #生成重复正确返佣金号码的列表，简称正确列表
    global correct_list #定义全局变量
    correct_list=[]
    for i in dianxin_numlist:
        if i in yuangong_numlist:
            correct_list.append(i)
            
listcompare()

def new_excel(): #生成新excel文件
    app=xw.App(visible=False,add_book=False)
    dianxin_wb=app.books.open(dianxin_xlpath) 
    new_wb=app.books.open(yuangong_xlpath) #新excel文件先应用员工模板化excel文件
    chongfunum_yuangong_list=[] #需要填入是否返佣金项的range号列表，F
    chongfunum_yuangong_list2=[] #需要填入佣金金额的range号列表，G

    dianxinyongjin_dict={} #佣金字典
    num_dict={} #号码字典
    
    for i in correct_list:
            chongfunum_yuangong_list.append('F'+str(yuangong_numlist.index(i)+2)) #对比员工放号列表和正确列表生成，F range
            num_dict['G'+str(yuangong_numlist.index(i)+2)]=i
            chongfunum_yuangong_list2.append('G'+str(yuangong_numlist.index(i)+2)) #对比员工放号列表和正确列表生成，G range
        
    for i in chongfunum_yuangong_list:
        new_wb.sheets['sheet1'].range(i).value='佣金已返' 
    
    for i in correct_list:
        dianxinyongjin_dict[i]=dianxin_wb.sheets['sheet1'].range('B'+str(dianxin_numlist.index(i)+3)).value
    
    for i in chongfunum_yuangong_list2:
        new_wb.sheets['sheet1'].range(i).value=dianxinyongjin_dict[num_dict[i]]


    
    new_path='e:\work\%s(加备注).xlsx'%os.path.split(yuangong_xlpath)[1][:-5] #通过os模块获取员工excel文件名生成新文件名
    new_wb.save(new_path)
    new_wb.close()
    dianxin_wb.close()
    app.quit()                                                                                                                                                                                                                                         

new_excel()
print('完成！')