#列表对比函数
def listcompare():
    global correct_list
    correct_list=[]
    for i in dianxin_numlist:
        if i in yuangong_numlist:
            correct_list.append(i)
            