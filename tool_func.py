import pymysql

conn = pymysql.connect(
    host="192.168.0.101",
    port=3306,
    user="root",
    password="xxx",
    db="devops",
    charset="utf8"
)
# 执行sql的光标
cursor = conn.cursor()

# 执行sql
sql = "SELECT PROJECT_ID, PROJECT_NAME FROM `dps_pjm_project`"
cursor.execute(sql)

# 获取查询数据
result = cursor.fetchall()

# 执行完毕后关闭光标
cursor.close()
# 关闭数据库
conn.close()


class MySQLHandler(pymysql.connections.Connection):
    def __init__(self):
        super(MySQLHandler, self).__init__()


