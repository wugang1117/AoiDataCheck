import pymssql

host = '192.168.0.36'
server = 'DACENTER\PRODMSSQL'
db = 'InetDbBack'
user = 'sa'
passwd = 'v5yIT9YmfaOIDF3kitIcHuZa&aKu&B$5'
# user = 'data13'
# passwd = '&iYiG6bW^U1%a6jDuo!sFTn1Ee%x3I'

# 封装类
class MysqlHelp(object):
    # 构造
    def __init__(self):
        self.host = host
        self.server = server
        self.user = user
        self.passwd = passwd
        self.db = db

    # 创建连接
    def open_con(self):
        self.con = pymssql.connect(host=self.host,server=self.server, database=self.db, user=self.user, password=self.passwd)
        self.cursor = self.con.cursor()

    # 关闭连接
    def close(self):
        self.cursor.close()
        self.con.cursor()

    # 调用语句
    def insert_delete_update(self, sql):
        try:
            self.open_con()
            self.cursor.execute(sql, params=[])
            self.con.commit()
            self.close()
        except Exception as error:
            print(error)

    # 查询 接收全部的返回结果行
    def select_fetchall(self, sql, params=()):
        try:
            self.open_con()

            self.cursor.execute(sql, params)

            results = self.cursor.fetchall()

            self.con.commit()

            self.close()
            return results

        except Exception as error:
            print(error)

    # 调用存储过程语句
    def exec_sp(self, sql,param=[]):
        try:
            self.open_con()
            self.con.autocommit(True)
            self.cursor.execute(sql, params=param)
            self.close()
        except Exception as error:
            print(error)

    # 调用存储过程语句
    def select_exec_sp(self, sql):
        try:
            self.open_con()
            self.con.autocommit(True)
            self.cursor.execute(sql, params=[])
            results = self.cursor.fetchall()
            self.con.commit()
            self.close()
            return results
        except Exception as error:
            print(error)