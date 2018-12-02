import os
import shutil

import xlwt

from common import define
from dao.step1_collection_dao import collection_dao


# 筛选出需要评分的人
# 需要被评分的人的评分表是什么
# 找到评分表的每一行元素
# 找到元素的的策略是什么
# 根据策略找到评分人的集合
# 以评分人的名字创建文件夹，以被评人为文件名创建excel
# 将评分表的行元素写入的这个文件的尾部

class RuleWriter:
    def __init__(self):
        self.dao = collection_dao()

    def __writeToFile(self):
        dirPaht = define.dirPaht
        if os.path.exists(dirPaht):
            shutil.rmtree(dirPaht)

        os.makedirs(dirPaht)
        scroers = self.dao.get_allscroers()
        for pfName in scroers:
            filepath = os.path.join(dirPaht, "附件3 " + pfName + ".xls")
            resultWorkBook = xlwt.Workbook();
            borders = xlwt.Borders()
            borders.left = 1
            borders.right = 1
            borders.top = 1
            borders.bottom = 1
            borders.bottom_colour = 0x3A

            alignBase = xlwt.Alignment()
            alignBase.horz = xlwt.Alignment.HORZ_LEFT
            alignBase.vert = xlwt.Alignment.VERT_CENTER
            alignBase.wrap = True

            style = xlwt.XFStyle()
            style.alignment = alignBase
            style.borders = borders

            tall_style = xlwt.easyxf('font:height 240;')  # 36pt,类型小初的字号

            sheet1 = resultWorkBook.add_sheet('sheet1')
            sheet1.write(0, 0, '被评价人', style)
            sheet1.write(0, 1, '被评价人单位', style)
            sheet1.write(0, 2, '考核类型', style)
            sheet1.write(0, 3, '评价内容', style)
            sheet1.write(0, 4, '评价标准', style)
            sheet1.write(0, 5, '考核分值', style)
            sheet1.write(0, 6, '考核方式', style)
            sheet1.write(0, 7, '评分策略', style)
            sheet1.write(0, 8, '评分人', style)
            sheet1.write(0, 9, '得分', style)
            sheet1.write(0, 10, '支持材料', style)
            sheet1.col(0).width = 256 * 15
            sheet1.col(1).width = 256 * 15
            sheet1.col(2).width = 256 * 15
            sheet1.col(3).width = 256 * 60
            sheet1.col(4).width = 256 * 60
            sheet1.col(5).width = 256 * 10
            sheet1.col(6).width = 256 * 30
            sheet1.col(7).width = 256 * 10
            sheet1.col(8).width = 256 * 10
            sheet1.col(9).width = 256 * 10
            sheet1.col(10).width = 256 * 10
            sheet1.row(0).set_style(tall_style)

            rowsCount = 0
            rawRecordPos = self.dao.get_allruleByScroer(pfName)

            for rawRecordPo in rawRecordPos:
                rowsCount = rowsCount + 1
                userinfopos = self.dao.get_userinfo_byusername(rawRecordPo.scroee)
                for userinfopo in userinfopos:
                    sheet1.write(rowsCount, 0, userinfopo.name, style)
                    sheet1.write(rowsCount, 1, userinfopo.dep, style)
                    sheet1.write(rowsCount, 2, userinfopo.reporttype, style)
                    sheet1.write(rowsCount, 3, rawRecordPo.content, style)
                    sheet1.write(rowsCount, 4, rawRecordPo.rule, style)
                    sheet1.write(rowsCount, 5, rawRecordPo.maxvalue, style)
                    sheet1.write(rowsCount, 6, rawRecordPo.ruletype, style)
                    sheet1.write(rowsCount, 7, rawRecordPo.pfCategory, style)
                    sheet1.write(rowsCount, 8, pfName, style)
                    sheet1.write(rowsCount, 9, "", style)
                    sheet1.write(rowsCount, 10, "", style)
                    sheet1.row(rowsCount).set_style(tall_style)

            resultWorkBook.save(filepath)

    def __findInScroeGroup(self, grouanme, scroee):
        groupnames = self.dao.get_scroegroupnames()
        scroers = []
        if grouanme in groupnames:
            # 获取该组下所有的人
            scroers = self.dao.get_allscroers_by_groupname(grouanme, scroee)
        return scroers

    def __findInScroePairs(self, scroee, category):
        scroers = self.dao.get_scroepairs(scroee, category)
        if None == scroers:
            scroers = []
        return scroers

    def __findInUsers(self, name, depatment):
        scroers = self.dao.get_scroeInUserTable(name, depatment)
        if None == scroers:
            scroers = []
        return scroers

    def __findInAlldepatmentAdmins(self, name):
        scroers = self.dao.get_all_depatmentadmins(name)
        if None == scroers:
            scroers = []
        return scroers

    def __findScrers(self, user, rule):
        users = []
        catetorys = rule.pfCategory.partition(",")
        for catetory in catetorys:
            if "" == catetory or "---" == catetory:
                continue
            # 覆盖：管委会成员，8部们，管理总部领导，管委会主席、主任会计师，。
            subusers = self.__findInScroeGroup(catetory, user.name)
            users.extend(subusers)

            subusers = self.__findInScroePairs(user.name, catetory)
            users.extend(subusers)

            if '所在单位负责人' == catetory:
                subusers = self.__findInScroeGroup(user.dep, user.name)
                users.extend(subusers)

            if catetory.find('所在单位合伙人') > -1:
                subusers = self.__findInUsers(user.name, user.dep)
                users.extend(subusers)

            if '单位负责人' == catetory:
                subusers = self.__findInAlldepatmentAdmins(user.name)
            users.extend(subusers)

        return users

    def write(self):
        users = self.dao.get_alluserinfo()
        for user in users:
            # 根据绩效类型获取该绩效类型下所有的评分项
            rules = self.dao.get_allrules_byjxtype(user.reporttype)
            for rule in rules:
                scrers = self.__findScrers(user, rule)
                if scrers.__len__() == 0:
                    print(
                        "[WARNING]" + user.name + "with:" + user.reporttype + " has not find record with " + rule.pfCategory)
                    continue
                # 如果有多个人，需要生成多个评分记录
                for scroer in scrers:
                    self.dao.add_rawrule_info(user.name, scroer, rule)

        self.__writeToFile()
        self.dao.commit()
        self.dao.close()
