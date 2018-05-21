import wx
import threading
import os
from modules import inputdialog


class NotepadDialog(object):
    """docstring for notepad"""

    def __init__(self, name, target):
        """
        name dialog及线程名称
        target 执行对象
        """
        self.app = wx.App()
        self.win = wx.Frame(
            None,
            title=name,
            size=(600, 335))
        self.win.Bind(wx.EVT_CLOSE, self.on_close)
        self.bkg = wx.Panel(self.win)
        self.openBtn = wx.Button(self.bkg, label='open&&run')
        self.openBtn.Bind(wx.EVT_BUTTON, self.open_file)

        self.saveBtn = wx.Button(self.bkg, label='save')
        self.saveBtn.Bind(wx.EVT_BUTTON, self.save_file)

        self.filename = wx.TextCtrl(self.bkg, style=wx.TE_READONLY)
        self.contents = wx.TextCtrl(self.bkg, style=wx.TE_MULTILINE)

        self.hbox = wx.BoxSizer()
        self.hbox.Add(self.openBtn, proportion=0,
                      flag=wx.LEFT | wx.ALL, border=5)
        self.hbox.Add(self.filename, proportion=1, flag=wx.EXPAND |
                                                        wx.TOP | wx.BOTTOM, border=5)
        self.hbox.Add(self.saveBtn, proportion=0,
                      flag=wx.LEFT | wx.ALL, border=5)

        self.bbox = wx.BoxSizer(wx.VERTICAL)
        self.bbox.Add(self.hbox, proportion=0, flag=wx.EXPAND | wx.ALL)
        self.bbox.Add(
            self.contents,
            proportion=1,
            flag=wx.EXPAND | wx.LEFT | wx.BOTTOM | wx.RIGHT,
            border=5)
        self.wildcard = "text 文件 (*.txt)|*.txt|" \
                        "All files (*.*)|*.*"
        self.name = name
        self.target = target

    def show_modal(self):
        self.bkg.SetSizer(self.bbox)
        self.win.Show()
        self.app.MainLoop()

    def open_file(self, evt):
        dlg = wx.FileDialog(
            self.win,
            "Open",
            "",
            "",
            "All files (*.*)|*.*",
            wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:
            file_path = dlg.GetPath()
            self.filename.SetValue(file_path)
            args = (file_path, self.contents)
        else:
            return
        work_thread = threading.Thread(target=self.target.run,
                                       args=args,
                                       name=self.name)
        work_thread.setDaemon(True)
        work_thread.start()

    def save_file(self, evt):
        dlg = wx.FileDialog(
            None, message="保存文件", defaultDir=os.getcwd(), defaultFile="", wildcard=self.wildcard, style=wx.FD_SAVE)
        if dlg.ShowModal() == wx.ID_OK:
            self.filename.SetValue(dlg.GetPath())
            file = open(dlg.GetPath(), 'w')
            file.write(self.contents.GetValue())
            file.close()
            self.contents.AppendText("log保存成功\n")

    def on_close(self, event):
        dial = wx.MessageDialog(None, '程序数据还没有处理完成，确定退出?', 'Question',
                                wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION)
        num = len(threading.enumerate())
        if num == 1:
            self.win.Destroy()
        else:
            flag = True
            for n in threading.enumerate():
                if n.name == self.name:
                    flag = False
                    ret = dial.ShowModal()
                    if ret == wx.ID_YES:
                        self.win.Destroy()
                        break
                    else:
                        event.Veto()
                        break
            if flag is True:
                self.win.Destroy()


class InputNotepadDialog(NotepadDialog):
    """docstring for notepad"""

    def __init__(self, name, target, target_args):
        """
        name dialog及线程名称
        target 执行函数
        target_args 执行函数的参数
        """
        super().__init__(name, target)
        self.openBtn.SetLabel('open')
        self.target_args = target_args

    def open_file(self, evt):
        dlg = wx.FileDialog(
            self.win,
            "Open",
            "",
            "",
            "All files (*.*)|*.*",
            wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        if dlg.ShowModal() == wx.ID_OK:
            file_path = dlg.GetPath()
            self.filename.SetValue(file_path)
        else:
            return
        args1 = {'in_file_name': file_path, 'contents': self.contents}
        dlg = inputdialog.InputDialog(self.target_args)
        if dlg.ShowModal() == 0:
            args2 = dlg.results
            args = {}
            args.update(args1)
            args.update(args2)
            work_thread = threading.Thread(target=self.target.run,
                                           args=(args,),
                                           name=self.name)
            work_thread.setDaemon(True)
            work_thread.start()

# if __name__ == "__main__":
#     from apkpi import ApKpi
#     handler = ApKpi()
#     dialog = NotepadDialog("KPIAPAging", handler)
#     dialog.show_modal()

# from apkpi import ApKpi
# handler = ApKpi()
# dialog = InputNotepadDialog("KPIAPAging", handler, ('参数1', '参数2'))
# dialog.show_modal()
