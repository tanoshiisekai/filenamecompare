import os

from pyexcel_xls import get_data as get_data_xls
from pyexcel_xlsx import get_data as get_data_xlsx

dirname = "C:/datacompare/pic"
tablefile = "C:/datacompare/学籍表1.xls"


tablename1 = "学生基础信息"         # 文件1中要参照的表名
col1name = ['A', 'E']       # 文件1中的有效列（注意排序对结果有影响）
row1 = 3                    # 文件1从第几行数据开始读取，从0开始计数
rowcount1 = 42

unwantedchar = ["\u3000", " "]

mainkeycol1 = 0
mainkeycol2 = 1

def getcolnum(colname):
    resultlist = []
    for col in colname:
        thesum = 0
        length = len(col)
        loop = length - 1
        while loop >= 0:
            thesum = thesum + (ord(col[length-loop-1])-ord('A') + 1) * (26 ** loop)
            loop = loop - 1
        resultlist.append(thesum - 1)
    return resultlist


def listequal(list1, list2, keycol1, keycol2):
    if len(list1) != len(list2):
        return None
    else:
        for x in range(0, len(list1)):
            if list1[x] != list2[x]:
                if list1[keycol1] == list2[keycol1] or list1[keycol2] == list2[keycol2]:
                    return tuple(((list1[keycol1], list1[keycol2], list1[x]),
                                 (list2[keycol1], list2[keycol2], list2[x])))
                return tuple((False, list1[x], list2[x]))
        return True


col1 = getcolnum(col1name)

if tablefile.endswith(".xls"):
    f1 = get_data_xls(tablefile)[tablename1][row1:]
elif tablefile.endswith(".xlsx"):
    f1 = get_data_xlsx(tablefile)[tablename1][row1:]
table1list = []
for i in range(0, rowcount1):
    table1list.append([][:])
for c in col1:
    for d in range(0, rowcount1):
        table1list[d].append(f1[d][c])


def splitintoparts(filename):
    splitposi = 0
    for i in range(0, len(filename)):
        if filename[i].isdigit():
            splitposi = i
            break
    return (filename[0:i], filename[i:])

filelist = os.listdir(dirname)
filelist = [x.split(".")[0] for x in filelist]
for char in unwantedchar:
    filelist = [x.replace(char, "") for x in filelist]
filelist = [splitintoparts(x) for x in filelist]

reflist = table1list
aimlist = filelist

passcount = 0
for aim in aimlist:
    iflag = 0
    for ref in reflist:
        flag = listequal(aim, ref, mainkeycol1, mainkeycol2)
        if flag is True:
            iflag = 1
            passcount = passcount + 1
            break
        if iflag == 0:
            if flag[0] is not False:
                print("E: ", "\n", aim, "\n", ref, "\n", "===>", flag)
if passcount == len(aimlist):
    print("All Passed.")
print("Pass Count:", passcount)