from tkinter import *
from tkinter.filedialog import *
import os
import xlwings as xw



root=Tk()
dianxin_path=StringVar()
yuangong_path=StringVar()
root.title("电信佣金对账App 测试版")
root.geometry("400x400+100+100")

def selectPath_dianxin():
    global dianxin_xlpath
    path_=askopenfilename()
    dianxin_xlpath=path_
    dianxin_path.set(path_)

def selectPath_yuangong():
    global yuangong_xlpath
    path_=askopenfilename()
    yuangong_xlpath=path_
    yuangong_path.set(path_)



def numlist():
    global yuangong_numlist,dianxin_numlist,yuangong_xlpath,dianxin_xlpath

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



def listcompare(): 
    global correct_list
    correct_list=[]
    for i in dianxin_numlist:
        if i in yuangong_numlist:
            correct_list.append(i)
            
def getDesktop():
    return os.path.join(os.path.expanduser("~"), 'Desktop')

def newexcel():
    app=xw.App(visible=False,add_book=False)
    dianxin_wb=app.books.open(dianxin_xlpath) 
    new_wb=app.books.open(yuangong_xlpath)
    chongfunum_yuangong_list=[]
    chongfunum_yuangong_list2=[]

    dianxinyongjin_dict={}
    num_dict={}
    
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


    
    new_path=r'%s\%s(加备注).xlsx'%(getDesktop(),os.path.split(yuangong_xlpath)[1][:-5]) 
    new_wb.save(new_path)
    new_wb.close()
    dianxin_wb.close()
    app.quit()                                                                                                                                                                                                                                         

def create():
    numlist()
    listcompare()
    newexcel()

Label(root,text="员工文件路径选择:").grid(row=0,column=0)
Label(root,text="电信文件路径选择:").grid(row=1,column=0)
Entry(root,textvariable=yuangong_path).grid(row=0,column=1)
Entry(root,textvariable=dianxin_path).grid(row=1,column=1) 
Button(root,text="选择",command=selectPath_yuangong).grid(row=0,column=2)
Button(root,text="选择",command=selectPath_dianxin).grid(row=1,column=2)
Button(root,text="生成",command=create).grid(row=2,column=2)    
    
root.mainloop()
