from utils import DataTimeUtils

prefix = "d:"
# prefix = "C:/Date_D"

def get_step1_db_name():
    return prefix + "/test/jixiao_step1_prepair_" + str(DataTimeUtils.get_current_year()) + ".db"


def get_step2_db_name():
    return prefix + "/test/jixiao_step1_calculation_" + str(DataTimeUtils.get_current_year()) + ".db"


rule_config_filename = prefix + '''/test/考核表汇总_2.xlsx'''

input_user_map_filename = prefix +  '''/test/编程对应的人员组别_1.xlsx'''

dirPaht = prefix + "/test/result"

resultPath = prefix + "/test/final/person"
finalPath = prefix + "/test/final"
