import wx
import wx.lib.buttons as buttons
from wx import adv
from modules import globalvar
from modules import notepaddialog
from modules import login
from modules import home
from modules import manager
from modules.apkpi import ApKpi
from modules import pspapagingreport
from modules import checkhold
from modules import grirreport


# 主框架
class Frame(wx.Frame):

    def __init__(self):
        wx.Frame.__init__(self, None, -1, u'xy工具箱', size=(1080, 730), pos=(350, 20))
        self.sp = wx.SplitterWindow(self, style=wx.SP_LIVE_UPDATE)
        self.panel1 = wx.Panel(self.sp, -1, style=wx.SUNKEN_BORDER)
        self.panel2 = wx.Panel(self.sp, -1)
        self.statusbar = self.CreateStatusBar()
        self.toolbar = self.CreateToolBar()
        self.menubar = wx.MenuBar()
        self.SetMinSize((1080, 730))
        self.SetMaxSize((1080, 730))
        # self.cursorinit()
        self.set_splitter_window()
        self.add_panel1button()
        self.set_statusbar()
        self.home_handler(None)
        self.set_menubar()
        self.set_toolbar()
        icon = wx.Icon('./myicon.ico', wx.BITMAP_TYPE_ICO, 32, 32)
        self.SetIcon(icon)

    # 分割窗口
    def set_splitter_window(self):
        self.panel1.SetBackgroundColour((36, 169, 225))
        # self.panel2.SetBackgroundColour('AQUAMARINE')
        self.sp.SplitVertically(self.panel1, self.panel2, 150)
        self.sp.SetMinimumPaneSize(150)

    # 状态栏
    def set_statusbar(self):
        self.statusbar.SetFieldsCount(3)
        self.panel2.Bind(wx.EVT_MOTION, self.on_motion)

    # 工具栏
    def set_toolbar(self):
        self.toolbar.SetExtraStyle(wx.TB_FLAT | wx.TB_DOCKABLE)
        self.toolbar.AddTool(
            100, "a", wx.Image("./pic/homepage.png", wx.BITMAP_TYPE_PNG).Scale(32, 32).ConvertToBitmap())
        self.toolbar.AddSeparator()
        self.toolbar.AddTool(
            101, "b", wx.Image("./pic/manage.png", wx.BITMAP_TYPE_PNG).Scale(32, 32).ConvertToBitmap())
        self.toolbar.Realize()

    # 状态栏坐标显示
    def on_motion(self, event):
        self.statusbar.SetStatusText(u'光标坐标:  ' + str(event.GetPosition()), 1)

    # 菜单栏
    def set_menubar(self):
        menu1 = wx.Menu()
        menu1.Append(100, "&主页", "返回主页")
        self.Bind(wx.EVT_MENU, self.home_handler, id=100)
        menu1.Append(101, "&功能管理", "用户功能管理")
        self.Bind(wx.EVT_MENU, self.manage_handler, id=101)
        self.menubar.Append(menu1, u'文件')
        menu2 = wx.Menu()
        menu2.Append(200, "&Copy", "Copy in status bar")
        menu2.Append(201, "C&ut", "")
        menu2.Append(202, "Paste", "")
        menu2.AppendSeparator()
        menu2.Append(203, "&Options...", "Display Options")
        self.menubar.Append(menu2, "&编辑")
        menu3 = wx.Menu()
        self.menubar.Append(menu3, u'其他')
        self.SetMenuBar(self.menubar)

    # 自定义光标
    def init_cursor(self):
        pic_cursor = wx.Image('./pic/cursor.png', wx.BITMAP_TYPE_PNG)
        cursor = wx.CursorFromImage(pic_cursor)
        self.SetCursor(cursor)

    # button图片
    @staticmethod
    def set_button_pic():
        return [['./pic/homepage.png', u'主页'],
                ['./pic/manage.png', u'用户管理'],
                ['./pic/CheckHold.png', u'CheckHold'],
                ['./pic/GRIRReport.png', u'GRIRReport'],
                ['./pic/KPI-AP-aging.png', u'KPI-AP-aging'],
                ['./pic/PSPAPAgingReport.png', u'PSPAPAgingReport']]

    # 初始化按钮
    def init_button(self, handler, pic_index):
        pic = wx.Image(self.set_button_pic()[pic_index][0], wx.BITMAP_TYPE_PNG).Scale(
            80, 80).ConvertToBitmap()
        button = buttons.GenBitmapButton(
            self.panel1, -1, pic, size=(150, 90))
        button.SetBezelWidth(7)
        button.SetBackgroundColour('sky blue')
        button.SetToolTip(self.set_button_pic()[pic_index][1])
        self.Bind(wx.EVT_BUTTON, handler, button)
        return button

    # panel1按钮添加
    def add_panel1button(self):
        button3 = self.init_button(self.checkhold_handler, 2)
        button4 = self.init_button(self.grirreport_handler, 3)
        button5 = self.init_button(self.kpiapaging_handler, 4)
        button6 = self.init_button(self.pspapagingreport_handler, 5)
        sizer = wx.FlexGridSizer(rows=0, cols=1, hgap=5, vgap=5)
        sizer.Add(button3, 0, wx.EXPAND)
        sizer.Add(button4, 0, wx.EXPAND)
        sizer.Add(button5, 0, wx.EXPAND)
        sizer.Add(button6, 0, wx.EXPAND)
        sizer.AddGrowableCol(0, proportion=1)
        sizer.Layout()
        self.panel1.SetSizer(sizer)
        self.panel1.Fit()

    # 主页事件
    def home_handler(self, event):
        self.statusbar.SetStatusText(u'欢迎使用xy工具箱！', 0)
        home.Index(self.panel2)

    # 功能管理事件
    def manage_handler(self, event):
        self.statusbar.SetStatusText(u'这是我的密码本', 0)
        manager.Table(self.panel2, globalvar.get_username())

    # 按钮事件
    def checkhold_handler(self, event):
        self.statusbar.SetStatusText(u'这是我的CheckHold', 0)
        handler = checkhold.CheckHold()
        dialog = notepaddialog.InputNotepadDialog("CheckHold", handler, ('AutoHold_var', 'TC_HOLD_AMOUNT_var'))
        dialog.show_modal()

    # 按钮事件
    def grirreport_handler(self, event):
        self.statusbar.SetStatusText(u'这是我的GRIRReporthandler', 0)
        grirreport.GRIRReport()

    # 按钮事件
    def kpiapaging_handler(self, event):
        self.statusbar.SetStatusText(u'这是我的KPIAPAGING', 0)
        handler = ApKpi()
        dialog = notepaddialog.NotepadDialog("KPIAPAging", handler)
        dialog.show_modal()

    # 按钮事件
    def pspapagingreport_handler(self, event):
        self.statusbar.SetStatusText(u'这是我的PSPAPAgingReport', 0)
        pspapagingreport.PSPAPAgingReport()


# 启动画面
def splash_screen():
    pic = wx.Image('./pic/splashscreen.png', type=wx.BITMAP_TYPE_PNG).ConvertToBitmap()
    adv.SplashScreen(pic, adv.SPLASH_CENTER_ON_SCREEN | adv.SPLASH_TIMEOUT, 1000, None, -1)


if __name__ == '__main__':
    app = wx.App(False)
    # 启动画面
    splash_screen()
    # 登录对话框
    dlg = login.LoginDialog()
    dlg.ShowModal()
    dlg.Destroy()
    if globalvar.get_loginvalue():
        exit()
    # 主框架
    frame = Frame()
    frame.Show()
    app.MainLoop()
