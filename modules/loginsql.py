import sqlite3


# 登录检查
def check(name, password):
    conn = sqlite3.connect('./database/user.db')
    c = conn.cursor()
    try:
        passwordr = c.execute('SELECT password FROM USER WHERE username=? ', [name])
        a = ''
        for i in passwordr:
            a = str(i[0])
        if a == password and a is not '':
            return 1
        else:
            return 0
    except sqlite3.OperationalError:
        return 0
    finally:
        conn.close()


# 注册事件
def insert(name, password):
    if name == '' or password == '':
        return 1
    conn = sqlite3.connect('./database/user.db')
    c = conn.cursor()
    c.execute('INSERT INTO USER (username,password) VALUES(?,?)', (name, password))
    conn.commit()
    conn.close()
    return 0


# 建表事件，用于测试
def create():
    conn = sqlite3.connect('./database/user.db')
    c = conn.cursor()
    c.execute(
        'CREATE TABLE USER (username VARCHAR PRIMARY KEY NOT NULL,password VARCHAR NOT NULL);')
    conn.close()


# 删除事件，用于测试
def delete(name):
    conn = sqlite3.connect('./database/user.db')
    c = conn.cursor()
    c.execute('DELETE FROM USER WHERE name=?', [(name)])


# 查表事件，用于测试
def show_table():
    conn = sqlite3.connect('./database/user.db')
    c = conn.cursor()
    passwordr = c.execute('''SELECT password,username FROM USER ''')
    for i in passwordr:
        print(i[0], i[1])
    conn.close()
