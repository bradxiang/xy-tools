import wx
import time

from modules import loginsql
from modules import globalvar


# 登录对话框
class LoginDialog(wx.Dialog):

    def __init__(self):
        time.sleep(2)
        text = u'欢迎使用xy工具箱!'
        wx.Dialog.__init__(self, None, -1, u'登录/注册', pos=(450, 330))
        # 对话框部件设置
        self.SetBackgroundColour('CORAL')
        texts = wx.StaticText(self, -1, text)
        name = wx.StaticText(self, -1, u'用户名')
        password = wx.StaticText(self, -1, u'密码')
        self.namet = wx.TextCtrl(self)
        self.passwordt = wx.TextCtrl(self, style=wx.TE_PASSWORD)
        self.namet.SetBackgroundColour('white')
        self.passwordt.SetBackgroundColour('white')
        self.blank = wx.StaticText(self, -1, '')
        login = wx.Button(self, id=-1, label=u'登录')
        sign = wx.Button(self, id=-1, label=u'注册')
        login.SetBackgroundColour('sky blue')
        sign.SetBackgroundColour('sky blue')
        fgs = wx.FlexGridSizer(2, 2, 5, 5)
        fgs.Add(name, 0, wx.ALIGN_RIGHT)
        fgs.Add(self.namet, 0, wx.EXPAND)
        fgs.Add(password, 0, wx.ALIGN_RIGHT)
        fgs.Add(self.passwordt, 0, wx.EXPAND)
        fgs.AddGrowableCol(1)
        # sizer设置
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(texts, 0, wx.ALL, 5)
        sizer.Add(wx.StaticLine(self), 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(fgs, 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(self.blank, 0, wx.ALIGN_CENTER, 5)
        sizer.Add(login, 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(sign, 0, wx.EXPAND | wx.ALL, 5)
        self.SetSizer(sizer)
        # 按钮绑定
        self.Bind(wx.EVT_BUTTON, self.login_handler, login)
        self.Bind(wx.EVT_BUTTON, self.sign_handler, sign)
        self.Bind(wx.EVT_CLOSE, self.on_close)

    # 登录事件
    def login_handler(self, event):
        final = loginsql.check(self.namet.GetValue(),
                               self.passwordt.GetValue())
        if not final:
            self.blank.SetLabelText(u'用户名或密码错误')
        else:
            globalvar.set_username(self.namet.GetValue())
            self.Close(True)
            globalvar.set_loginvalue(0)
            event.Skip()

    # 注册事件
    def sign_handler(self, event):
        # returns=loginsql.insert(self.namet.GetValue(),self.passwordt.GetValue())
        # if returns:
        #     self.blank.SetLabelText(u'用户名或密码不能为空')
        self.blank.SetLabelText(u'不允许注册！！！')

    # 关闭事件
    @staticmethod
    def on_close(event):
        globalvar.set_loginvalue(1)
        event.Skip()
