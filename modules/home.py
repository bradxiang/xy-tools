import wx


class Index(object):

    def __init__(self, panel2):
        panel2.DestroyChildren()
        panel_pic = wx.Image(
            './pic/panel.jpg',
            wx.BITMAP_TYPE_JPEG,).Scale(930, 730).ConvertToBitmap()
        wx.StaticBitmap(
            parent=panel2, pos=(0, 0), bitmap=panel_pic)
