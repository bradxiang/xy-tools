import wx
import wx.grid
import shelve
import wx.lib.buttons as buttons


class Table(object):

    def __init__(self, panel, user_name):
        panel.DestroyChildren()
        self.table = None
        self.database = ""
        self.panel = panel
        self.user_name = user_name
        self.shelve_file(user_name)
        self.column_labels = [u'项目名', u'用户名', u'密码', u'注释']
        self.grid = wx.grid.Grid(
            parent=self.panel, id=-1, pos=(0, 0),
            size=(900, 571), style=wx.WANTS_CHARS,)
        self.add_button()
        self.init_table()
        self.panel.Bind(wx.EVT_CLOSE, self.on_close)

    # 表格初始化
    def init_table(self):
        self.grid.CreateGrid(100, 4)
        # self.grid.SetColSize(1,180)
        self.grid.SetDefaultColSize(200, True)
        self.grid.SetDefaultRowSize(30, True)
        self.grid.SetDefaultCellAlignment(wx.CENTRE, wx.CENTRE)
        self.grid.SetDefaultCellTextColour('black')
        self.grid.SetDefaultCellBackgroundColour('sky blue')
        self.grid.SetDefaultCellFont(
            wx.Font(11, wx.SWISS, wx.NORMAL, wx.NORMAL))
        for row in range(4):
            self.grid.SetColLabelValue(row, self.column_labels[row])
        self.show_table()

    # 按钮数据
    @staticmethod
    def set_button_pic():
        return [['./pic/save.png', u'保存', (800, 572)],
                ['./pic/add.png', u'新建', (850, 572)]]

    # 初始化按钮
    def init_button(self, handler, pic_index):
        pic = wx.Image(self.set_button_pic()[pic_index][0], wx.BITMAP_TYPE_PNG).Scale(
            30, 30).ConvertToBitmap()
        button = buttons.GenBitmapButton(
            self.panel, -1, pic, size=(32, 32),
            pos=self.set_button_pic()[pic_index][2])
        button.SetBackgroundColour('CORAL')
        button.SetToolTip(self.set_button_pic()[pic_index][1])
        button.Bind(wx.EVT_BUTTON, handler)
        return button

    # 添加按钮
    def add_button(self):
        self.init_button(self.save_handler, 0)
        self.init_button(self.new_handler, 1)

    # 数据库查询
    def shelve_file(self, user_name):
        self.database = shelve.open('./database/database.dat')
        try:
            self.table = self.database[user_name]
            self.database.close()
        except Exception as e:
            print(e)
            self.database[self.user_name] = [
                ["" for i in range(4)] for j in range(5)]

    # 显示表
    def show_table(self):
        self.database = shelve.open('./database/database.dat')
        self.table = self.database[self.user_name]
        for row in range(len(self.database[self.user_name])):
            for col in range(4):
                temp = self.database[self.user_name][row][col]
                if temp is not None:
                    self.grid.SetCellValue(row, col, '%s' % self.database[
                                           self.user_name][row][col])
                else:
                    self.grid.SetCellValue(row, col, '')
        self.grid.ForceRefresh()
        self.database.close()

    # 保存事件
    def save_handler(self, event):
        temp = list()
        for i in range(self.grid.GetNumberRows()):
            if self.grid.GetCellValue(i, 0) == "":
                break
            else:
                temp.append([ self.grid.GetCellValue(i, x) for x in range(self.grid.GetNumberCols())])
        self.database = shelve.open('./database/database.dat')
        self.table = self.database[self.user_name]
        self.database[self.user_name] = temp
        self.database.close()
        dlg = wx.MessageDialog(None, u"数据保存成功", u"标题信息", wx.ICON_QUESTION)
        if dlg.ShowModal() == wx.ID_YES:
            self.Close(True)
        dlg.Destroy()
        self.show_table()

    # 新建行
    def new_handler(self, event):
        self.grid.AppendRows(numRows=1)
        self.grid.ForceRefresh()

    def on_close(self, event):
        pass
