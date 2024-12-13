# coding=gbk
import pandas as pd
import os
import xlwings as xw
import time


def if_not_in(a1, b1):
    for I in b1:
        if a1 in I:
            return 0
    return 1


def sort_key(key):
    return key[1], key[2]


read = [1, 1]
r = ["1", "2"]
color1 = (165, 42, 42)
color2 = (255, 0, 255)
sheet_keyword = ""
while 1:
    keyword = str(input("请输入关键字（输入“？”进入设置）：\n"))
    if keyword == "?":
        print("Copyright ? 2021 HarveyZhang. All rights reserved.\nGmail:HarveyZhangJY@gmail.com")
        print("=========================================================")
        while 1:
            print("1-----设置需筛选的sheet的关键字（当前关键字：", sheet_keyword, "）\n2-----设置颜色（当前设置：", color1, color2,
                  "）\n3-----设置需修改哪些表（当前设置：", r, ")\n0-----退出设置")
            choose = int(input())
            if choose == 0:
                break
            elif choose == 1:
                sheet_keyword = str(input("请输入筛选sheet的关键字：\n"))
            elif choose == 2:
                color1 = str(input("请输入表1差异行需要标记颜色的rgb值：\n")).split(" ")
                color2 = str(input("请输入表2差异行需要标记颜色的rgb值：\n")).split(" ")
                for i in range(3):
                    color1[1] = int(color1[i])
                    color2[1] = int(color2[i])
                color1 = tuple(color1)
                color2 = tuple(color2)
            elif choose == 3:
                r = str(input("请输入需要修改哪些表：\n")).split(" ")
                if r == ["1"]:
                    read = [1, 0]
                elif r == ["2"]:
                    read = [0, 1]
                elif not r:
                    read = [0, 0]
                else:
                    read = [1, 1]
        keyword = str(input("请输入需对比文件的关键字：\n"))
    if "\\" in keyword:
        a = os.popen("dir " + keyword.split(" ")[0] + "*" + keyword.split(" ")[-1] + "*.xls* /B/s")
    else:
        a = os.popen(r"dir *" + keyword + r"*.xls* /B/s")
    b = []
    for line in a.readlines():
        b.append(line.replace("\n", ""))
    if len(b) >= 2:
        for i in range(len(b)):
            if "\\" in keyword:
                print(i + 1, b[i].replace("\\".join(keyword.split("\\")[0:-1]), ""))
            else:
                print(i + 1, b[i].replace(os.getcwd(), ""))
        files = str(input("请选择两张表，用空格隔开：\n")).split(" ")
        while len(files) != 2 or files[0] == "" or files[1] == "" or int(files[0]) > len(b) or int(files[1]) > len(b):
            print("输入有误，请重新输入")
            files = str(input("请选择两张表，用空格隔开：\n")).split(" ")

        content1 = []
        content2 = []
        need_judge1 = []
        need_judge2 = []
        diff1 = []
        diff2 = []
        name1 = b[int(files[0]) - 1]
        name2 = b[int(files[1]) - 1]
        df1 = pd.read_excel(name1, sheet_name=None)
        df2 = pd.read_excel(name2, sheet_name=None)

        cols = str(input("请输入需要对比的列号，可输入多列，用空格隔开（如A BB C等）：\n")).split(" ")
        cols2 = cols.copy()
        c = cols.copy()
        col = []
        if cols != [""]:
            for i in range(len(cols)):
                for j in cols[i]:
                    col.append(str(ord(j) - 64))
                print(int("".join(col), 26))
                col = []
        else:
            cols = []
        sheet_name1 = []
        sheet_name2 = []
        sheet_n1 = list(df1.keys())
        sheet_n2 = list(df2.keys())
        for l, k in df1.items():
            if sheet_keyword in l:
                sheet_name1.append(l)
                content1.append([])
                for i in range(k.shape[0]):
                    content1[-1].append([])
                    for j in range(k.shape[1]):
                        if pd.notnull(k.iloc[i, j]):
                            content1[-1][i].append(
                                str(k.iloc[i, j]).replace("\n", "").replace(" ", "").replace(",", "").replace("-",
                                                                                                              "").replace(
                                    ".", "").replace("_", "").replace("Dx000", "").replace("x000D", ""))
                        else:
                            content1[-1][i].append("")
                count1 = 0
                for j in range(len(content1[-1])):
                    if not content1[-1][count1]:
                        del content1[-1][count1]
                    else:
                        count1 += 1
        for l, k in df2.items():
            if sheet_keyword in l:
                sheet_name2.append(l)
                content2.append([])
                for i in range(k.shape[0]):
                    content2[-1].append([])
                    for j in range(k.shape[1]):
                        if pd.notnull(k.iloc[i, j]):
                            content2[-1][i].append(
                                str(k.iloc[i, j]).replace("\n", "").replace(" ", "").replace(",", "").replace("-",
                                                                                                              "").replace(
                                    ".", "").replace("_", "").replace("Dx000", "").replace("x000D", ""))
                        else:
                            content2[-1][i].append("")
                count2 = 0
                for j in range(len(content2[-1])):
                    if not content2[-1][count2]:
                        del content2[-1][count2]
                    else:
                        count2 += 1
        if cols:
            for i in content1:
                need_judge1.append([])
                for j in i:
                    need_judge1[-1].append([])
                    for k in range(len(j)):
                        if k in cols:
                            need_judge1[-1][-1].append(j[k])
            for i in content2:
                need_judge2.append([])
                for j in i:
                    need_judge2[-1].append([])
                    for k in range(len(j)):
                        if k in cols:
                            need_judge2[-1][-1].append(j[k])
        else:
            need_judge1 = content1
            need_judge2 = content2
        for i in range(len(need_judge1)):
            diff1.append([])
            for j in range(len(need_judge1[i])):
                if if_not_in(need_judge1[i][j], need_judge2):
                    diff1[-1].append([j] + content1[i][j])
        for i in range(len(need_judge2)):
            diff2.append([])
            for j in range(len(need_judge2[i])):
                if if_not_in(need_judge2[i][j], need_judge1):
                    diff2[-1].append([j] + content2[i][j])
        diff_write1 = []
        diff_write2 = []
        for i in range(len(diff1)):
            for j in range(len(diff1[i])):
                if diff1[i][j]:
                    diff_write1.append([1, sheet_name1[i], diff1[i][j][0]])
        for i in range(len(diff2)):
            for j in range(len(diff2[i])):
                if diff2[i][j]:
                    diff_write2.append([2, sheet_name2[i], diff2[i][j][0]])
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        wb = app.books.add()
        if read[0] == 1:
            wb1 = app.books.open(name1)
            sheets1 = wb1.sheets
            for i in diff_write1:
                if cols:
                    for j in cols1:
                        sheets1[sheet_n1.index(i[1])].range(j + str(i[2] + 2)).color = color1
                else:
                    row1 = sheet1[sheet_n1.index(i[1])].api.UsedRange.Columns.count
                    sheets1[sheet_n1.index(i[1])].range((1,i[2] + 2),(row,i[2] + 2)).color = color1
        if read[1] == 1:
            wb2 = app.books.open(name2)
            sheets2 = wb2.sheets
            for i in diff_write2:
                if cols:
                    for j in cols2:
                        sheets2[sheet_n2.index(i[1])].range(j + str(i[2] + 2)).color = color2
                else:
                    row2 = sheets2[sheet_n2.index(i[1])].api.UsedRange.Columns.count
                    sheets2[sheet_n2.index(i[1])].range((1,i[2] + 2),(row,i[2] + 2)).color = color2
        wb1.save(
            name1[0:-5].split("\\")[-1] + "_" + keyword.replace("*", "") + "_" + "_".join(c) + str(
                time.time()) + ".xlsx")
        wb2.save(
            name2[0:-5].split("\\")[-1] + "_" + keyword.replace("*", "") + "_" + "_".join(c) + str(
                time.time()) + ".xlsx")
        wb.save('1.xlsx')
        # os.system("start " + name1[0:-5] + "_" + keyword + "_" + "_".join(c) + str(time.time()) + ".xlsx")
        # os.system("start " + name2[0:-5] + "_" + keyword + "_" + "_".join(c) + str(time.time()) + ".xlsx")
        wb1.close()
        wb2.close()
        app.quit()
        print("文件1发现", len(diff_write1), "条差异，保存至",
              name1[0:-5].split("\\")[-1] + "_" + keyword.replace("*", "") + "_" + "_".join(c) + str(
                  time.time()) + ".xlsx")
        print("文件2发现", len(diff_write2), "条差异，保存至",
              name2[0:-5].split("\\")[-1] + "_" + keyword.replace("*", "") + "_" + "_".join(c) + str(
                  time.time()) + ".xlsx")
        back = str(input("是否需要对比其他文件：y|n\n"))
        if back != "y" and back != "Y":
            break
    else:
        print("至少需要找到两个文件，请重新输入。")
