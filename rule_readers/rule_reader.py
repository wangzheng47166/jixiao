# 读取所有人员的影射关系
import xlrd

from common import define
from common.record_po import RecordItemPo
from common.scroe_group_po import ScroeGroupPo
from common.scroe_pair_po import ScroePairPo
from common.userinfo_po import UserInfoPo
from dao.step1_collection_dao import collection_dao


class RulesReader:
    def __init__(self):
        filepath = define.rule_config_filename
        self.dataWorkBook = xlrd.open_workbook(filepath)
        self.dao = collection_dao()

    # 加载评价组
    def read_rule_table(self):
        pageCount = len(self.dataWorkBook.sheets())
        for i in range(pageCount):
            sheet = self.dataWorkBook.sheet_by_index(i)
            tableName = sheet.name
            rowcount = sheet.nrows
            for j in range(rowcount):
                # 跳过三行表头
                if j > 2:
                    # content, rule, maxvalue, ruletype, pfCategory, tableName
                    po = RecordItemPo(
                        (sheet.row_values(j)[1]),
                        (sheet.row_values(j)[2]),
                        (sheet.row_values(j)[3]),
                        (sheet.row_values(j)[4]),
                        (sheet.row_values(j)[6]),
                        tableName
                    )

                    if not ('' == po.pfCategory or '---' == po.pfCategory):
                        self.dao.add_ruleinfo(po)

    def read(self):
        self.read_rule_table()
        self.dao.commit()
