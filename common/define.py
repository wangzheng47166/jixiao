from utils import DataTimeUtils


def get_step1_db_name():
    return "C:/Date_D/test/jixiao_step1_prepair_" + str(DataTimeUtils.get_current_year()) + ".db"


def get_step2_db_name():
    return "C:/Date_D/test/jixiao_step1_calculation_" + str(DataTimeUtils.get_current_year()) + ".db"


rule_config_filename = '''C:/Date_D/test/考核表汇总_2.xlsx'''

input_user_map_filename = '''C:/Date_D/test/编程对应的人员组别_1.xlsx'''

dirPaht = "C:/Date_D/test/result"
