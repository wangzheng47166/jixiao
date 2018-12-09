import os
import sqlite3

from common import define
from common.rawrecord_po import ScroeTaskPo
from common.record_po import RecordItemPo
from common.userinfo_po import UserInfoPo


class collection_dao:
    conn = None

    def __init__(self):
        self.initConn()
        return

    def initConn(self):
        if None == collection_dao.conn:
            dbname = define.get_step1_db_name()
            if os.path.exists(dbname):
                os.remove(dbname)
                # conn1 = sqlite3.connect(dbname)
            collection_dao.conn = sqlite3.connect(dbname)
            self.intTables()

    def intTables(self):
        # 创建 打分组表
        createScoreGroup = '''create table score_group (
          id integer  PRIMARY KEY autoincrement,
          scroegroup text,
          user text,
          desc text default null 
        );'''

        # 创建 特殊评价关系
        createScorePaire = '''
        create table score_pair (
            id integer  PRIMARY KEY autoincrement,
            scoree text not null , 
            scorer text not null ,
            scorekey text not null ,
            desc text default null 
        );
        '''
        createUsers = '''create table user (
          id INT ,
          name text,
          unit text,
          jxtype text,
          mail
        );'''

        createRecords = '''
            create table records (
              record_index varchar(64),
              content text,
              rule    text,
              maxvalue    text,
              ruletype  text,
              pfCategory text,
              tableName text,
              unique (record_index)
            );
        '''

        createRawResults = '''
            create table rawrecords (
                  record_index varchar(64),
              name text,
              pfname text,
              content text,
              rule    text,
              maxvalue    text,
              ruletype  text,
              pfCategory text,
              relval     text,
              unique (record_index,name,pfname)
            );
        '''

        collection_dao.conn.execute(createUsers)
        collection_dao.conn.execute(createScoreGroup)
        collection_dao.conn.execute(createScorePaire)
        collection_dao.conn.execute(createRecords)
        collection_dao.conn.execute(createRawResults)

    def close(self):
        collection_dao.conn.commit()
        collection_dao.conn.close()

    def commit(self):
        collection_dao.conn.commit()

    def add_scroegroup_info(self, scroegrouppo):
        self.initConn()
        sql_insert = "insert into score_group(scroegroup,user,desc) VALUES (:scroegroup,:user,:desc);"
        collection_dao.conn.execute(sql_insert, {
            'scroegroup': scroegrouppo.groupname,
            'user': scroegrouppo.username,
            'desc': scroegrouppo.desc
        })

    def add_scroepair_info(self, scropairpo):
        self.initConn()
        sql = "insert into score_pair(scoree,scorer,scorekey,desc) VALUES (:scoree,:scorer,:scorekey,:desc);"
        collection_dao.conn.execute(sql, {
            'scoree': (scropairpo.scroee),
            'scorer': (scropairpo.scroer),
            'scorekey': (scropairpo.scroekey),
            'desc': (scropairpo.desc),
        })

    def add_userinfo(self, userinfopo):
        self.initConn()
        sql_insert = "insert into user (id,name,unit,jxtype,mail) VALUES (:id,:name,:unit,:jxtype,:mail);"
        collection_dao.conn.execute(sql_insert, {
            'id': userinfopo.id,
            'name': (userinfopo.name),
            'unit': (userinfopo.dep),
            'jxtype': (userinfopo.reporttype),
            'mail': (userinfopo.email)
        })

    def add_ruleinfo(self, record_po):
        self.initConn()
        sql = "insert into records" \
              "(record_index,content,rule,maxvalue,ruletype,pfCategory,tableName) " \
              "VALUES (:recordindex,:content,:rule,:maxvalue,:ruletype,:pfc,:tablename)"
        collection_dao.conn.execute(sql, {
            "recordindex": record_po.record_index,
            "content": (record_po.content),
            "rule": (record_po.rule),
            "maxvalue": (record_po.maxvalue),
            "ruletype": (record_po.ruletype),
            "pfc": (record_po.pfCategory),
            "tablename": (record_po.tableName)
        })

    def add_rawrule_info(self, scroee, scroer, record_po):
        if None == scroer:
            print("[WARNING]" + scroee + " has no scroer " + "" + " with " + record_po.pfCategory)
            return

        exits = collection_dao.conn.execute("SELECT * FROM rawrecords WHERE "
                                            "name=:name AND "
                                            "pfname=:pfname AND "
                                            "record_index=:index ",
                                            {
                                                'name': scroee,
                                                'pfname': scroer,
                                                'index': record_po.record_index
                                            })

        rowcount = len(exits.fetchall())
        if rowcount > 0:
            print("name->" + scroee + ";pfname->" + scroer + ";c-> " + (record_po.pfCategory) + " 重复，被跳过,找到行数： ",
                  rowcount)
            return

        insetStr = "insert into rawrecords " \
                   "(record_index, name,pfname,content,rule,maxvalue,ruletype,pfCategory) " \
                   "values (:recordindex,:name,:pfname,:content,:rule,:maxvalue,:ruletype,:pfc)"

        collection_dao.conn.execute(insetStr, {
            "recordindex": record_po.record_index,
            'name': scroee,
            'content': (record_po.content),
            'rule': (record_po.rule),
            'maxvalue': (record_po.maxvalue),
            'ruletype': (record_po.ruletype),
            'pfname': scroer,
            'pfc': (record_po.pfCategory)})

    def get_alluserinfo(self):
        self.initConn()
        users = []
        rows = collection_dao.conn.execute("SELECT name,jxtype,unit FROM user;")
        for row in rows:
            user = UserInfoPo(
                0, row[0], row[2], row[1], ''
            )
            users.append(user)
        return users

    def get_allrules_byjxtype(self, jxtype):
        self.initConn()
        rules = []
        rows = collection_dao.conn.execute(
            "SELECT record_index,content,rule,maxvalue,ruletype,pfCategory "
            "FROM records "
            "WHERE tableName=:tablename", {
                "tablename": jxtype
            })
        for row in rows:
            rule = RecordItemPo(
                row[0], row[1], row[2], row[3], row[4], row[5], ''
            )
            rules.append(rule)
        return rules

    def get_scroegroupnames(self):
        self.initConn()
        groupnames = []
        sql = "select distinct scroegroup from score_group "
        rows = collection_dao.conn.execute(sql)
        for row in rows:
            groupnames.append(row[0])

        return groupnames

    def get_allscroers_by_groupname(self, grouanme, scroee):
        self.initConn()
        users = []
        sql = "select distinct user from score_group where scroegroup = :scroegroup and user<>:scroee ";
        rows = collection_dao.conn.execute(sql, {
            "scroegroup": grouanme,
            "scroee": scroee
        })
        for row in rows:
            users.append(row[0])
        return users

    def get_scroepairs(self, scroee, category):
        self.initConn()
        users = []
        sql = "select distinct scorer from score_pair where scoree=:scroee and scorekey=:category"
        rows = collection_dao.conn.execute(sql, {
            "category": category,
            "scroee": scroee
        })
        for row in rows:
            users.append(row[0])
        return users

    # 在同一个分所的其他人
    def get_scroeInUserTable(self, username, depatment):
        self.initConn()
        users = []
        sql = "select distinct name from user where name<>:username and unit=:depatment"
        rows = collection_dao.conn.execute(sql, {
            "username": username,
            "depatment": depatment
        })
        for row in rows:
            users.append(row[0])
        return users

    # 获取除了自己之外的分所其他人
    def get_all_depatmentadmins(self, scroee):
        self.initConn()
        users = []
        sql = "select distinct user from score_group where desc = :desc and user<>:scroee "
        rows = collection_dao.conn.execute(sql, {
            "desc": '单位负责人',
            "scroee": scroee
        })
        for row in rows:
            users.append(row[0])
        return users

    # 获取所有评分人的信息
    def get_allscroers(self):
        self.initConn()
        users = []
        rows = collection_dao.conn.execute("SELECT DISTINCT pfname FROM rawrecords")
        for row in rows:
            users.append(row[0])
        return users

    # 获取评分人的所有评分任务
    def get_allruleByScroer(self, pfname):
        self.initConn()
        sql = "SELECT name,content,rule,maxvalue,ruletype,pfCategory FROM rawrecords WHERE pfname=:pfname"
        rows = collection_dao.conn.execute(sql, {
            "pfname": pfname
        })
        pos = []
        for row in rows:
            po = ScroeTaskPo(
                row[0],
                pfname,
                row[1],
                row[2],
                row[3],
                row[4],
                row[5],
                ''
            )
            pos.append(po)
        return pos

    def get_userinfo_byusername(self, username):
        self.initConn()
        sql = "SELECT name,unit,jxtype FROM user WHERE name=:username"
        rows = collection_dao.conn.execute(sql, {
            "username": username
        })
        pos = []
        for row in rows:
            po = UserInfoPo(
                0, row[0], row[1], row[2], ''
            )
            pos.append(po)
        return pos
