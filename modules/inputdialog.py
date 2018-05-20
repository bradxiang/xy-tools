#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'''
In this code example, we create a
custom dialog.

author: xy
last modified: july 2018
'''

import wx


class InputDialog(wx.Dialog):

    def __init__(self, name=()):
        wx.Dialog.__init__(self, None)
        if len(name) < 1:
            print("参数为零")
        self.text_ctrls =list()
        self.results = {}
        self.args = name
        self.init_ui()
        self.SetSize((250, 230))
        self.SetTitle("InputDialog")

    def init_ui(self):
        pnl = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)
        sb = wx.StaticBox(pnl, label='Setting values:')
        sbs = wx.StaticBoxSizer(sb, orient=wx.VERTICAL)

        # sbs.Add(wx.RadioButton(pnl, label='256 Colors',
        #     style=wx.RB_GROUP))
        # sbs.Add(wx.RadioButton(pnl, label='128 Colors'))
        # sbs.Add(wx.TextCtrl(pnl), flag=wx.LEFT, border=5)
        for v in self.args:
            hbox = wx.BoxSizer(wx.HORIZONTAL)
            static_text = wx.StaticText(pnl, -1, v+"：", style=wx.ALIGN_CENTER)
            hbox.Add(static_text)
            text_ctrl = wx.TextCtrl(pnl)
            self.text_ctrls.append(text_ctrl)
            hbox.Add(text_ctrl, flag=wx.CENTRE, border=5)
            sbs.Add(hbox, 0, wx.EXPAND | wx.ALL, 3)
        pnl.SetSizer(sbs)

        hbox = wx.BoxSizer(wx.HORIZONTAL)
        ok_button = wx.Button(self,  label='运行')
        close_button = wx.Button(self, wx.ID_CANCEL, label='取消')
        hbox.Add(ok_button)
        hbox.Add(close_button, flag=wx.LEFT, border=5)

        vbox.Add(pnl, proportion=1,
                 flag=wx.ALL | wx.EXPAND, border=5)
        vbox.Add(hbox, flag=wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, border=10)
        self.SetSizer(vbox)
        self.Bind(wx.EVT_BUTTON, self.on_close, ok_button)

    def on_close(self, e):
        for i in range(len(self.args)):
            self.results[self.args[i]] = self.text_ctrls[i].GetValue()
        self.Destroy()


class Example(wx.Frame):

    def __init__(self, *args, **kw):
        super(Example, self).__init__(*args, **kw)

        self.InitUI()

    def InitUI(self):
        tb = self.CreateToolBar()
        tb.AddTool(toolId=wx.ID_ANY, label='', bitmap=wx.Bitmap('../pic/add.png'))

        tb.Realize()

        tb.Bind(wx.EVT_TOOL, self.OnChangeDepth)

        self.SetSize((350, 250))
        self.SetTitle('Custom dialog')
        self.Centre()

    def OnChangeDepth(self, e):
        cdDialog = InputDialog(('Change Color Depth',))
        cdDialog.ShowModal()
        cdDialog.Destroy()


# if __name__ == '__main__':
#     app = wx.App()
#     ex = Example(None)
#     ex.Show()
#     app.MainLoop()
