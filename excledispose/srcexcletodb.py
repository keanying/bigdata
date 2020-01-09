import xlrd
import pymssql


class SQLConnect:
    def __init__(self, server, user, password, database, port):
        self.server = server
        self.user = user
        self.password = password
        self.database = database
        self.port = port

    # 获取连接
    def sqlConnect(self):
        if not self.database:
            raise (NameError, "No databases infomation .")
        self.conn = pymssql.connect(self.server, self.user, self.password, self.database, self.port, charset="utf8")
        cursor = self.conn.cursor()
        if not cursor:
            raise (NameError, "Connection mysql error .")
        else:
            return cursor

    # 查询
    def query(self, sql):
        cursor = self.sqlConnect()
        cursor.execute(sql)
        resList = cursor.fetchall()
        self.conn.close()
        return resList

    # 批量添加
    def insert(self, sql, values):
        cursor = self.sqlConnect()
        cursor.executemany(sql, values)
        # self.conn.commit() #commit() 对于oracle才会执行commit提交
        self.conn.close()


class queryExcle():
    def excle(self):
        book = None
        try:
            book = xlrd.open_workbook(
                "D:\\WeChat Files\\WeChat Files\txt\\FileStorage\\File\\2019-11\\example.xlsx")  # excle文件地址
        except Exception as e:
            print('EXCLE file path err: %s' % (e))
        sheet = book.sheet_by_name("20191104-multipurpose")  # excle表的sheet标识位
        sql = "INSERT INTO PostData_API_Multipurpose_test(annotation,operationsteps,expect,result,post_value,return_value) VALUES (%s,%s,%s,%s,%s,%s)"

        for r in range(1, sheet.nrows):  # 跳过excle第一行，如果第一行为标题信息
            values = []
            annotation = sheet.cell(r, 1).value  # 获取行的第2列
            operationsteps = sheet.cell(r, 2).value  # 获取行的第3列
            expect = sheet.cell(r, 3).value  # 获取行的第4列
            result = sheet.cell(r, 4).value  # 获取行的第5列
            post_value = sheet.cell(r, 5).value  # 获取行的第6列
            return_value = sheet.cell(r, 6).value  # 获取行的第7列

            # 将需要的excle对应到mysql表列字段位置，构建元祖，位置必须与mysql表字段位置对应
            num = (str(annotation), str(operationsteps), str(expect), str(result), str(post_value), str(return_value))
            values.append(num)  # 将元组添加到列表
            nummbers = sheet.cell(r, 0).value  # 获取行的第1列，excle内的ID标识，记录执行到哪里
            try:
                self.run_preferential(sql, values)
            except Exception as e:
                print(e)
                print("wrong with the perform to id %s ." % (str(nummbers)))
            continue

    def run_preferential(self, sql, values):
        server = "10.001.001.111"
        user = "username"
        password = "password"
        database = "databases"
        port = 1433
        connection = SQLConnect(server, user, password, database, port)
        connection.insert(sql=sql, values=values)

if __name__ == '__main__':
    print("begain start app ...")
    start = queryExcle()
    start.excle()
    print("application execution completed .")
