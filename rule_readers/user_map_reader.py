# 读取所有人员的影射关系
import xlrd

from common import define
from common.scroe_group_po import ScroeGroupPo
from common.scroe_pair_po import ScroePairPo
from common.userinfo_po import UserInfoPo
from dao.step1_collection_dao import collection_dao


class UserMapReader:
    def __init__(self):
        filepath = define.input_user_map_filename
        self.dataWorkBook = xlrd.open_workbook(filepath)
        self.dao = collection_dao()

    # 加载评价组
    def read_scroegroup_table(self):
        sheet1 = self.dataWorkBook.sheet_by_name("评价组")
        for i in range(sheet1.nrows):
            if 0 == i:
                # 跳过标题行
                continue
            # 逐条保存个人信息
            scroegroup_po = ScroeGroupPo(
                (sheet1.row_values(i)[0]),
                (sheet1.row_values(i)[1]),
                (sheet1.row_values(i)[2])
            )
            self.dao.add_scroegroup_info(scroegroup_po)

    # 加载特殊评价关系评
    def read_scroepair_table(self):
        sheet1 = self.dataWorkBook.sheet_by_name("特殊评价关系")
        for i in range(sheet1.nrows):
            if 0 == i:
                # 跳过标题行
                continue
            # 逐条保存个人信息
            po = ScroePairPo(
                (sheet1.row_values(i)[0]),
                (sheet1.row_values(i)[1]),
                (sheet1.row_values(i)[2]),
                (''),
            )
            self.dao.add_scroepair_info(po)

    # 读取考核人的人员信息
    # 考核的类型必须是预制信息，并且每个考核类型对应一个报表输出
    def read_userinfo_table(self):
        sheet1 = self.dataWorkBook.sheet_by_name("考核名单")
        for i in range(sheet1.nrows):
            if 0 == i:
                # 跳过标题行
                continue
            # 逐条保存个人信息
            userinfo_po = UserInfoPo(
                (sheet1.row_values(i)[0]),
                (sheet1.row_values(i)[1]),
                (sheet1.row_values(i)[2]),
                (sheet1.row_values(i)[3]),
                ''
            )
            self.dao.add_userinfo(userinfo_po)

    def read(self):
        self.read_scroegroup_table()
        self.read_scroepair_table()
        self.read_userinfo_table()
        self.dao.commit()
