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