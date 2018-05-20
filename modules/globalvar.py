# 无需互斥变量
USER_NAME = ''
LOGIN_VALUE = 0


def get_loginvalue():
    global LOGIN_VALUE
    return LOGIN_VALUE


def set_loginvalue(value):
    global LOGIN_VALUE
    LOGIN_VALUE = value


def get_username():
    global USER_NAME
    return USER_NAME


def set_username(value):
    global USER_NAME
    USER_NAME = value
