from utils import DataTimeUtils
from common import define

print("当前年份 - %s", DataTimeUtils.get_current_year())

print("数据库1 - %s", define.get_step1_db_name())

print("数据库2 - %s", define.get_step2_db_name())
