#!/usr/bin/env python

import xlrd
import xlwt
import os
import shutil
import sqlite3

#
# for fpathe,dirs,fs in os.walk('result'):
#     for f in fs:
#         print(os.path.join(fpathe,f))
#
#
# resultWorkBook = xlwt.Workbook();
#
# sheet1 = resultWorkBook.add_sheet('sheet1')
# sheet1.write_merge(0, 0, 0, 1, 'Long Cell')
# sheet1.write(1, 0, 1)
# sheet1.write(1, 1, 2)
# resultWorkBook.save("测试大表.xls");

if os.path.exists('jx2017_step2.db'):
    os.remove('jx2017_step2.db')
conn = sqlite3.connect('jx2017_step2.db')

createRawResults = '''
    create table rawrecords (
      name text,
      pfname text,
      pfType text,
      pfContent text,
      pfInfo  text,
      pfCategory text,
      relval     int DEFAULT 0
    );
'''

createRecords = '''
    create table records (
      content text,
      rule    text,
      maxvalue    text,
      ruletype  text,
      pfCategory text,
      tableName text
    );
'''

createUsers = '''create table user (
  id INT ,
  name text,
  unit text,
  jxtype text,
  mail
);'''

conn.execute(createRawResults)
conn.execute(createUsers)
conn.execute(createRecords)


def readConfigExcel():
    filepath = '编程对应的人员组别.xlsx'
    dataWorkBook = xlrd.open_workbook(filepath)

    # 第二个表
    sheet1 = dataWorkBook.sheet_by_index(1)
    rowscont = sheet1.nrows
    print("编程对应的人员组别.xlsx 中用户表共用户数 %s", rowscont)
    for i in range(rowscont):
        id = sheet1.row_values(i)[0]
        name = sheet1.row_values(i)[1]
        unit = sheet1.row_values(i)[2]
        jxtype = sheet1.row_values(i)[3]
        mail = sheet1.row_values(i)[4]
        sqlInsert = "insert into user (id,name,unit,jxtype,mail) VALUES (:id,:name,:unit,:jxtype,:mail);"
        sqlresut = conn.execute(sqlInsert, {'id': id, 'name': name, 'unit': unit, 'jxtype': jxtype, 'mail': mail})


def readRecordRunExcel():
    filepath = '考核表汇总.xlsx'
    dataWorkBook = xlrd.open_workbook(filepath)
    pageCount = len(dataWorkBook.sheets())
    for i in range(pageCount):
        sheet = dataWorkBook.sheet_by_index(i)
        tableName = sheet.name
        rowcount = sheet.nrows
        for j in range(rowcount):
            if j > 2:
                content = sheet.row_values(j)[1]
                rule = sheet.row_values(j)[2]
                maxvalue = sheet.row_values(j)[3]
                ruletype = sheet.row_values(j)[4]
                pfCategory = sheet.row_values(j)[6]
                insertSql = "insert into records(content,rule,maxvalue,ruletype,pfCategory,tableName) VALUES (:content,:rule,:maxvale,:ruletype,:pfc,:tablename)"
                conn.execute(insertSql,
                             {'content': content, 'rule': rule, 'maxvale': maxvalue, 'ruletype': ruletype,
                              'pfc': pfCategory, 'tablename': tableName})


def initDatabase():
    readConfigExcel()
    readRecordRunExcel()


def readRawData():
    for fpathe, dirs, fs in os.walk('result'):
        for f in fs:
            dataPath = os.path.join(fpathe, f)
            dataWorkBook = xlrd.open_workbook(dataPath)
            sheet1 = dataWorkBook.sheet_by_index(0)
            print("录入结果： %s", dataPath)
            rowscont = sheet1.nrows
            for i in range(rowscont):
                name = sheet1.row_values(i)[0]
                pfType = sheet1.row_values(i)[2]
                pfContent = sheet1.row_values(i)[3]
                pfInfo = sheet1.row_values(i)[4]
                pfCategory = sheet1.row_values(i)[7]
                pfName = sheet1.row_values(i)[8]
                value = sheet1.row_values(i)[9]
                if "石晨起" == name:
                    print(rowscont)

                if "得分" != value:
                    insertSql = " insert into rawrecords(name,pfname,pfType,pfContent,pfInfo,pfCategory,relval)" \
                                " VALUES (:name,:pfname,:pfType,:pfContent,:pfInfo,:pfCategory,:relval)"
                    intValue = 0
                    if '' != value:
                        intValue = float(value)
                    sqlresut = conn.execute(insertSql,
                                            {
                                                'name': name,
                                                'pfname': pfName,
                                                'pfType': pfType,
                                                'pfCategory': pfCategory,
                                                'relval': intValue,
                                                'pfContent': pfContent,
                                                'pfInfo': pfInfo
                                            })


def finalBigReport():
    dirPaht = "final"
    if os.path.exists(dirPaht):
        shutil.rmtree(dirPaht)

    os.makedirs(dirPaht)

    fileName = "原始数据.xls"
    filepath = os.path.join(dirPaht, fileName);

    resultWorkBook = xlwt.Workbook();
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    borders.bottom_colour = 0x3A

    style = xlwt.XFStyle()
    tall_style = xlwt.easyxf('font:height 240;')  # 36pt,类型小初的字号

    sheet1 = resultWorkBook.add_sheet('sheet1')
    sheet1.write(0, 0, '被评价人', style)
    sheet1.write(0, 1, '考核类型', style)
    sheet1.write(0, 2, '考核内容', style)
    sheet1.write(0, 3, '评分细则', style)
    sheet1.write(0, 4, '评分策略', style)
    sheet1.write(0, 5, '评分人', style)
    sheet1.write(0, 6, '得分', style)
    sheet1.col(0).width = 256 * 30
    sheet1.col(1).width = 256 * 30
    sheet1.col(2).width = 256 * 30
    sheet1.col(3).width = 256 * 30
    sheet1.col(4).width = 256 * 30
    sheet1.col(5).width = 256 * 30
    sheet1.col(6).width = 256 * 30

    sheet1.row(0).set_style(tall_style)

    pfRecords = conn.execute(
        "SELECT name,pfname,pfType,pfCategory,relval,pfContent,pfInfo FROM rawrecords ORDER BY name,pfType,pfCategory");
    for pfRecord in pfRecords:
        rowsCount = len(sheet1.rows) + 1
        if rowsCount < 10:
            print("routecount %s", rowsCount)

        sheet1.write(rowsCount, 0, pfRecord[0], style)
        sheet1.write(rowsCount, 1, pfRecord[5], style)
        sheet1.write(rowsCount, 2, pfRecord[2], style)
        sheet1.write(rowsCount, 3, pfRecord[6], style)
        sheet1.write(rowsCount, 4, pfRecord[3], style)
        sheet1.write(rowsCount, 5, pfRecord[1], style)
        sheet1.write(rowsCount, 6, pfRecord[4], style)

    resultWorkBook.save(filepath)


def finalPersonReport():
    dirPaht = "final\\persons"
    os.makedirs(dirPaht)

    # 从用户表中查出所有的人，如果查不到该人的评分记录，需要报警一下
    users = conn.execute("SELECT DISTINCT name,jxtype,unit FROM user")

    baseborder = xlwt.Borders()
    baseborder.left = 1
    baseborder.right = 1
    baseborder.top = 1
    baseborder.bottom = 1
    baseborder.bottom_colour = 0x3A

    for user in users:
        # 没人生成个表项

        name = user[0]
        pfType = user[1]
        unit = user[2]
        # 根据个人的评分类型，得到评分策略
        items = conn.execute("SELECT content,rule,maxvalue,ruletype,pfCategory FROM records WHERE tableName=:tbname",
                             {
                                 "tbname": pfType
                             })
        items = items.fetchall()
        rowcount = len(items)
        if rowcount == 0:
            print("error : %s 的考核表 %s 错误了", name, pfType)

        # 写表头
        fileName = name + ".xls"
        filepath = os.path.join(dirPaht, fileName)

        resultWorkBook = xlwt.Workbook()

        sheet1 = resultWorkBook.add_sheet('sheet1')

        # 首行
        if True:
            style = xlwt.XFStyle()
            fnt = xlwt.Font()
            fnt.bold = True
            fnt.name = u'宋体'
            # 10.5 * 20
            fnt.height = 210

            alignBase = xlwt.Alignment()
            alignBase.horz = xlwt.Alignment.HORZ_CENTER
            alignBase.vert = xlwt.Alignment.VERT_CENTER
            style.alignment = alignBase
            style.font = fnt
            style.borders = baseborder
            sheet1.write_merge(0, 0, 0, 6, "合伙人综合管理考核评价表（" + pfType + "）", style)

        # 次行
        if True:
            style = xlwt.XFStyle()
            fnt = xlwt.Font()
            fnt.bold = True
            fnt.name = u'宋体'
            # 10 * 20
            fnt.height = 200

            alignBase = xlwt.Alignment()
            alignBase.horz = xlwt.Alignment.HORZ_LEFT
            alignBase.vert = xlwt.Alignment.VERT_CENTER
            style.alignment = alignBase
            style.font = fnt
            style.borders = baseborder
            sheet1.write_merge(1, 1, 0, 6, "姓名：" + name +
                               "            单位：" + unit +
                               "            职务：" + pfType +
                               "            考核年度：2017 ", style)

        # 三行 title
        if True:
            style = xlwt.XFStyle()
            fnt = xlwt.Font()
            fnt.bold = True
            fnt.name = u'宋体'
            # 10 * 20
            fnt.height = 200

            alignBase = xlwt.Alignment()
            alignBase.horz = xlwt.Alignment.HORZ_CENTER
            alignBase.vert = xlwt.Alignment.VERT_CENTER
            style.alignment = alignBase
            style.font = fnt
            style.borders = baseborder

            sheet1.write(2, 0, '序号', style)
            sheet1.write(2, 1, '评价内容', style)
            sheet1.write(2, 2, '评价标准', style)
            sheet1.write(2, 3, '考核分值', style)
            sheet1.write(2, 4, '考核方式', style)
            sheet1.write(2, 5, '评分', style)
            sheet1.write(2, 6, '评分人', style)

        baseline = 2

        fnt = xlwt.Font()
        fnt.name = u'宋体'
        # 10 * 20
        fnt.height = 200

        alignCenter = xlwt.Alignment()
        alignCenter.horz = xlwt.Alignment.HORZ_CENTER
        alignCenter.vert = xlwt.Alignment.VERT_CENTER
        alignCenter.wrap = True

        alignLeft = xlwt.Alignment()
        alignLeft.horz = xlwt.Alignment.HORZ_LEFT
        alignLeft.vert = xlwt.Alignment.VERT_CENTER
        alignLeft.wrap = True

        styleNumber = xlwt.XFStyle()
        styleNumber.alignment = alignCenter
        styleNumber.font = fnt
        styleNumber.borders = baseborder

        styleText = xlwt.XFStyle()
        styleText.alignment = alignLeft
        styleText.font = fnt
        styleText.borders = baseborder

        totleMax = 0
        totleAvg = 0

        for item in items:
            baseline += 1
            avgValue = 0
            pyCategory = item[4]
            maxValue = item[2]
            content = item[0]
            pfInfo = item[1]
            try:
                totleMax += int(round(float(maxValue)))
            except ValueError as e:
                print(maxValue)

            valueRows = conn.execute(
                "SELECT round(avg(relval),1) AS avgvalue FROM rawrecords WHERE name = :name AND  pfType = :pfType AND pfCategory = :pfc AND pfContent=:content AND pfInfo = :pfInfo",
                {
                    "name": name,
                    "pfType": pfType,
                    "pfc": pyCategory,
                    "content": content,
                    "pfInfo": pfInfo
                })
            valueRows = valueRows.fetchall()
            rowcount = len(valueRows)
            if rowcount > 0:
                for valueRow in valueRows:
                    if None != valueRow[0]:
                        avgValue = valueRow[0]

            totleAvg += avgValue
            sheet1.write(baseline, 0, baseline - 2, styleNumber)
            sheet1.write(baseline, 1, item[0], styleText)
            sheet1.write(baseline, 2, item[1], styleText)
            sheet1.write(baseline, 3, maxValue, styleNumber)
            sheet1.write(baseline, 4, item[3], styleText)
            sheet1.write(baseline, 5, avgValue, styleNumber)
            sheet1.write(baseline, 6, item[4], styleNumber)

        if True:
            baseline += 1
            sheet1.write_merge(baseline, baseline, 0, 2, "合计", styleText)
            sheet1.write(baseline, 3, totleMax, styleNumber)
            sheet1.write(baseline, 4, "", styleNumber)
            sheet1.write(baseline, 5, totleAvg, styleNumber)
            sheet1.write(baseline, 6, "", styleNumber)

        sheet1.col(0).width = 256 * 10
        sheet1.col(1).width = 256 * 50
        sheet1.col(2).width = 256 * 50
        sheet1.col(3).width = 256 * 10
        sheet1.col(4).width = 256 * 25
        sheet1.col(5).width = 256 * 10
        sheet1.col(6).width = 256 * 25

        resultWorkBook.save(filepath)


def group_person_record():
    dirpath = "final"
    fileName = "汇总.xls"
    filepath = os.path.join(dirpath, fileName)
    resultWorkBook = xlwt.Workbook()
    sheet1 = resultWorkBook.add_sheet('sheet1')

    # 次行
    if True:
        style = xlwt.XFStyle()
        fnt = xlwt.Font()
        fnt.bold = True
        fnt.name = u'宋体'
        # 10 * 20
        fnt.height = 200

        baseborder = xlwt.Borders()
        baseborder.left = 1
        baseborder.right = 1
        baseborder.top = 1
        baseborder.bottom = 1
        baseborder.bottom_colour = 0x3A

        alignBase = xlwt.Alignment()
        alignBase.horz = xlwt.Alignment.HORZ_LEFT
        alignBase.vert = xlwt.Alignment.VERT_CENTER
        style.alignment = alignBase
        style.font = fnt
        style.borders = baseborder
        sheet1.write(0, 0, "姓名", style)
        sheet1.write(0, 1, "总分", style)

    baseline = 0
    valueRows = conn.execute(
        "SELECT name,sum(avgvalue) AS v FROM ( SELECT name,round(avg(relval),1) AS avgvalue FROM rawrecords GROUP BY name,pfType,pfCategory,pfContent,pfInfo ) GROUP BY name;")
    for record in valueRows:
        baseline += 1
        sheet1.write(baseline, 0, record[0], style)
        sheet1.write(baseline, 1, record[1], style)
    resultWorkBook.save(filepath)


def finalCollectReprot():
    dirpath = "final"
    fileName = "考核汇总表.xls"
    filepath = os.path.join(dirpath, fileName)
    resultWorkBook = xlwt.Workbook()

    style = xlwt.XFStyle()
    fnt = xlwt.Font()
    fnt.bold = True
    fnt.name = u'宋体'
    # 10 * 20
    fnt.height = 200

    baseborder = xlwt.Borders()
    baseborder.left = 1
    baseborder.right = 1
    baseborder.top = 1
    baseborder.bottom = 1
    baseborder.bottom_colour = 0x3A

    alignBase = xlwt.Alignment()
    alignBase.horz = xlwt.Alignment.HORZ_LEFT
    alignBase.vert = xlwt.Alignment.VERT_CENTER

    style.alignment = alignBase
    style.font = fnt
    style.borders = baseborder

    # 查出所有考核大类，每个考核类写一个sheet
    types = conn.execute("SELECT DISTINCT jxtype FROM user;")
    for type in types:
        # 针对每一个类型，新建一个sheets
        subsheet = resultWorkBook.add_sheet(type[0])

        # 找出各种考核类型
        subsheet.write(0, 0, "评分部门", style)
        subsheet.write(0, 1, "", style)
        subsheet.write(0, 2, "", style)

        subsheet.write(1, 0, "考核分值", style)
        subsheet.write(1, 1, "", style)
        subsheet.write(1, 2, "", style)

        subsheet.write(2, 0, "评分内容", style)
        subsheet.write(2, 1, "", style)
        subsheet.write(2, 2, "", style)

        subsheet.write(3, 0, "序号", style)
        subsheet.write(3, 1, "姓名", style)
        subsheet.write(3, 2, "单位", style)

        # 找出各种考核项目
        basecol = 3
        pfItems = conn.execute(
            "select distinct pfCategory,maxvalue,content,rule from records where tableName = :p0;",
            {
                "p0": type[0]
            })

        for pfitem in pfItems:
            subsheet.write(0, basecol, pfitem[0], style)
            subsheet.write(1, basecol, pfitem[1], style)
            subsheet.write(2, basecol, pfitem[2], style)
            subsheet.write(3, basecol, pfitem[3], style)
            basecol += 1

        # 找出符合这个考核的各个人
        users = conn.execute("select name,unit from user where jxtype = :jxtype order by unit,name;",{
            "jxtype" : type[0]
        })



        baserow = 4
        basecol = 3
        userindex = 1
        for user in users :
            username = user[0]
            print("写分 --- ", username)
            unit = user[1]
            subsheet.write(baserow, 0, userindex, style)
            subsheet.write(baserow, 1, username, style)
            subsheet.write(baserow, 2, unit, style)

            pfItems2 = conn.execute(
                "select distinct pfCategory,maxvalue,content,rule from records where tableName = :p0;",
                {
                    "p0": type[0]
                })
            for pfitem2 in pfItems2 :
                content = pfitem2[2]
                info = pfitem2[3]
                rawvalue = conn.execute("select round(avg(relval),1)  from rawrecords where name = :p0 and pfContent = :p1 and pfInfo = :p2;",
                {
                    "p0" : username,
                    "p1" : content,
                    "p2" : info
                })

                valueRows = conn.execute("select round(avg(relval),1)  from rawrecords where name = :p0 and pfContent = :p1 and pfInfo = :p2;",
                                         {
                                             "p0" : username,
                                             "p1" : content,
                                             "p2" : info
                                         })

                valueRows = valueRows.fetchall()
                rowcount = len(valueRows)
                print("写分 ===  ", rowcount)
                for relval0 in rawvalue :
                    val = relval0[0]
                    subsheet.write(baserow, basecol, val, style)
                    print("写分行 %s 列 %s", baserow,basecol)

                basecol += 1

            baserow += 1
            userindex += 1
            basecol = 3

            print("写分 >>>  ", baserow,userindex)


    resultWorkBook.save(filepath)



def main():
    initDatabase()
    readRawData()
    # 出最终报表
    # 出个大的报表，供快速筛查
    finalBigReport()
    # 根据每个人出个人表格
    finalPersonReport()
    group_person_record()

    # 最终的汇总表
    finalCollectReprot()


main()

conn.commit()
conn.close()
