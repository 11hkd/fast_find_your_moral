import operator
import sys
import xlwings as xw
import os
import traceback
import time

ls = []
namels = ["陈千里","彭健桉","罗奇国","范孜贺","庞宇豪","韩凯迪","刘俊杰","孙天泽","赵传宇","张少龙"\
          ,"孙嘉庆","吴可凡","王力冉","周保臣","陶虹润","张世鹏","秦鹏博","张增阳","孟令民","王文铖",\
            "马浩杰","吴雨深","刘昕月","张硕","柳启枝","孙明悦","吴晓艳","张诗棋","张雨鑫","汪小宇",\
            "汪超","唐玉梅","李坤龙","毛凯晨","谢鸣星","陈秋樾","王羽菲","刘国玉"]

pt = 0
print("正在加载,请确保此目录下的表格在此时没有被打开，否则可能出现意料之外的错误")
time.sleep(1.5)
print("正在查找表格··· ···")
time.sleep(0.5)

def dirimg(path):
    files = os.listdir(path)
    for i in files:
        file_d = os.path.join(path, i)
        ls.append(file_d)


def nchange(x):
    clist = ['壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖']
    danwei = ['拾', '佰', '元', '角', '分']
    num = 0
    if x is None:
        return 0
    if len(x) > 1:
        for i in range(len(x)):
            for j in range(len(clist)):
                for k in range(len(danwei)):
                    if x[i] == clist[j] and x[i + 1] == danwei[k]:
                        if k == 0:
                            num += 10 * (j + 1)
                        if k == 1:
                            num += 100 * (j + 1)
                        if k == 2:
                            num += (j + 1)
                        if k == 3:
                            num += 0.1 * (j + 1)
                        if k == 4:
                            num += 0.01 * (j + 1)
    if len(x) == 1:
        for i in range(len(clist)):
            if x == clist[i]:
                return i + 1
    return num


dirimg('.')

pathls,sheet,index,name,res,grd,tim,wb,app,idls = [],[],[],[],[],[],[],[],[],[]

for i in range(len(ls)):
    f, e = os.path.splitext(ls[i])
    if (e == ".xlsx" or e == ".xls") and (f[2] != "~" or f[1] != "~"):
        pathls.append(ls[i])
        print("找到表格：", ls[i])
        sheet.append("")
        index.append("")
        name.append("")
        res.append("")
        grd.append("")
        tim.append("")
        wb.append("")
        app.append("")
        idls.append("")
print("查找结束！正在初始化中。。。")
for i in range(len(pathls)):
    try:
        app[i] = xw.App(visible=False, add_book=False)
        wb[i] = app[i].books.open(pathls[i])
        sheet[i] = wb[i].sheets.active
        index[i] = sheet[i].api.UsedRange.Row + sheet[i].api.UsedRange.Rows.Count - 1
        name[i] = sheet[i].range(f'C4:C{index[i]}').value
        res[i] = sheet[i].range(f'D4:D{index[i]}').value
        grd[i] = sheet[i].range(f'F4:F{index[i]}').value
        tim[i] = sheet[i].range(f'G4:G{index[i]}').value
        idls[i] = sheet[i].range(f'B4:B{index[i]}').value
        app[i].kill() 
        for j in range(len(idls[i])):
            if idls[i][j] == None:
                break
            idls[i][j] = int(idls[i][j])
        print("初始化进度："+str(i+1)+"/"+str(len(pathls)))
    except Exception as e:
        print("查找时出现了错误，这有可能是未找到表格，也有可能是读取到了临时文件，如要解决此项问题，请关闭表格或重启电脑后再试。发生错误的表格：",pathls[i],"第",j,"行")
        sys.exit()
time.sleep(0.5)
flag0 = input("查询个人加分输入1，查询全班加分表选2\n")
dic = {}
if flag0 == "1":
    while True:
        idf = -1
        if pt == 0:
            print("初始化完毕！请选择读取模式：\n[0]：全部扫描")
            for i in range(len(pathls)):
                print("[" + str(i + 1) + "]：仅读取", pathls[i])
            flag = input()
            pt = -1
            try:
                flag = int(flag)
                if flag > len(pathls):
                    print("你看看有第",flag,"个选项吗")
                    pt = 0
            except:
                print("好好输个数字不行吗（恼，溜了")
                sys.exit()
        elif flag == 0:
            print("全部读取：")
            while True:
                nm = input("请输入你的名字或学号,若输入0则重新进入模式选择\n")
                print("")
                if nm == '0':
                    pt = 0
                    break
                sum = 0
                for i in range(len(pathls)):
                    f = 0
                    print(" 表格",pathls[i],"读取到：")
                    for j in range(index[i] - 4):
                        if name[i][j] == nm or str(idls[i][j]) == nm:
                            try:
                                int(grd[i][j])
                            except:
                                grd[i][j] = nchange(grd[i][j])
                            print("  ",res[i][j], " ", grd[i][j], " ", tim[i][j])
                            f += 1
                            sum += int(grd[i][j])
                    if f == 0:
                        print("  未在这个表格内找到该姓名所对应的加分项！")
                        print("")
                    else:
                        print("")
                if sum != 0:
                    print("总分为", sum)
                    print("")
                else:
                    print("未在表中找到这个名字")
                    print("")
        else:
            flag -= 1
            print("读取：",pathls[flag])
            while True:
                nm = input("请输入你的名字或学号,若输入0则重新进入模式选择\n")
                print("")
                if nm == '0':
                    pt = 0
                    break
                print("")
                sum = 0
                for j in range(index[flag] - 5):
                    if name[flag][j] == nm or str(idls[flag][j]) == nm:
                        try:
                            int(grd[flag][j])
                        except:
                            grd[flag][j] = nchange(grd[flag][j])
                        print(res[flag][j], " ", grd[flag][j], " ", tim[flag][j])
                        sum += int(grd[flag][j])
                if sum != 0:
                    print("这个表中的总分为", sum)
                    print("")
                else:
                    print("未找到这个名字")
                    print("")
elif flag0 == "2":
    while True:
        idf = -1
        if pt == 0:
            print("初始化完毕！请选择读取模式：\n[0]：全部扫描")
            for i in range(len(pathls)):
                print("[" + str(i + 1) + "]：仅读取", pathls[i])
            flag = input()
            pt = -1
            try:
                flag = int(flag)
                if flag > len(pathls):
                    print("你看看有第",flag,"个选项吗")
                    pt = 0
            except:
                print("好好输个数字不行吗（恼，溜了")
                sys.exit()
        elif flag == 0:
            print("全部读取：")
            while True:
                nm = input("按任意键后回车以继续,若输入0则重新进入模式选择\n")
                print("")
                if nm == '0':
                    pt = 0
                    break
                for k in namels:
                    sum = 0
                    for i in range(len(pathls)):
                        f = 0
                        for j in range(index[i] - 4):
                            if name[i][j] == k or str(idls[i][j]) == k:
                                try:
                                    int(grd[i][j])
                                except:
                                    grd[i][j] = nchange(grd[i][j])
                                sum += int(grd[i][j])
                    dic.update({k:sum})
                sort = sorted(dic.items(),key=operator.itemgetter(1),reverse=True)
                for k in sort:
                    print(k)
        else:
            flag -= 1
            print("读取：",pathls[flag])
            while True:
                nm = input("按任意键后回车以继续,若输入0则重新进入模式选择\n")
                print("")
                if nm == '0':
                    pt = 0
                    break
                print("")
                for k in namels:
                    sum = 0
                    for j in range(index[flag] - 5):
                        if name[flag][j] == k or str(idls[flag][j]) == k:
                            try:
                                int(grd[flag][j])
                            except:
                                grd[flag][j] = nchange(grd[flag][j])
                            sum += int(grd[flag][j])
                    dic.update({k:sum})
                sort = sorted(dic.items(),key=operator.itemgetter(1),reverse=True)
                for k in sort:
                    print(k)
else:
    print("不要乱输好伐")