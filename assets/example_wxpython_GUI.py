# let's import packages to use
import numpy as np
import math
import sqlite3
import pandas as pd
import datetime as dt
from sqlalchemy import create_engine
import sqlalchemy
from sqlite3 import OperationalError
import time
import os
from dateutil.relativedelta import relativedelta as delta
import sys
import re
from pandas import ExcelWriter
from pandas.api.types import CategoricalDtype
from pyxlsb import open_workbook as open_xlsb
import locale
from openpyxl import load_workbook
import wx
import wx.adv
from pubsub import pub
import calendar


###############################################################################
###############################################################################


class MyDialog(wx.MessageDialog):
    def __init__(self, parent, title):
        super(MyDialog, self).__init__(parent, title=title, size=(250, 150))

        self.Bind(wx.EVT_BUTTON, self.OnButtonClicked)

    def OnButtonClicked(self, e):

        # print('event reached panel class')
        e.Skip()


class MyPanel(wx.Panel):

    def __init__(self, *args, **kw):
        super(MyPanel, self).__init__(*args, **kw)

        self.Bind(wx.EVT_BUTTON, self.OnButtonClicked)

    def OnButtonClicked(self, e):

        # print('event reached panel class')
        e.Skip()


class MyButton(wx.Button):

    def __init__(self, *args, **kw):
        super(MyButton, self).__init__(*args, **kw)

        self.Bind(wx.EVT_BUTTON, self.OnButtonClicked)

    def OnButtonClicked(self, e):

        # print('event reached button class')
        e.Skip()


#''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


class Cal_Frame(wx.Frame):

    def __init__(self, parent, title):
        super(Cal_Frame, self).__init__(parent, title=title)

        self.tables_sql = parent.tables_out
        self.cols_sql = parent.cols_out
        self.d1 = parent.mm_dts[0]
        self.d2 = parent.mm_dts[1]
        print("Cal_Frame ", self.d1)
        print("Cal_Frame ", self.d2)
        self.initUI()

    def initUI(self):

        # force modality of the window
        self.MakeModal(True)
        # create panel
        self.panel = wx.Panel(self)

        self.early = None
        self.late = None

        # add sizer
        hbox = wx.BoxSizer(wx.HORIZONTAL)

        # create buttons
        buttonClose = wx.Button(self.panel, label="Frühdatum")
        buttonOk = wx.Button(self.panel, label="Spätdatum")
        # add buttons to sizer
        hbox.Add(buttonClose, 0, flag=wx.ALL, border=5)
        hbox.Add(buttonOk, 0, flag=wx.ALL, border=5)

        # create sizer for buttons
        buttonbox = wx.BoxSizer(wx.HORIZONTAL)
        # create buttons
        buttonClose = wx.Button(self.panel, label="Abbrechen")
        buttonOk = wx.Button(self.panel, label="Bestätigen")
        # add buttons to sizer
        buttonbox.Add(buttonClose, 0, flag=wx.ALL, border=5)
        buttonbox.Add(buttonOk, 0, flag=wx.ALL, border=5)

        # create topsizer
        topSizer = wx.BoxSizer(wx.VERTICAL)
        # add selection boxes and buttons to topsizer
        topSizer.Add(hbox, 0, wx.CENTER)
        topSizer.Add(buttonbox, 0, wx.ALIGN_RIGHT)
        # compute the panels size based on the topSizer
        self.panel.SetSizer(topSizer)
        # fit sizer
        topSizer.Fit(self)

        # event binding
        pub.subscribe(self.myListener_early, "early")
        pub.subscribe(self.myListener_late, "late")
        self.Bind(wx.EVT_BUTTON, self.OnButtonClicked)

        # show frame
        self.Center()
        self.Show()

    def OnButtonClicked(self, e):

        if e.GetEventObject().GetLabel() == "Frühdatum":
            Calendar_early(self, "Kalendar")

        elif e.GetEventObject().GetLabel() == "Spätdatum":
            Calendar_late(self, "Kalendar")

        elif e.GetEventObject().GetLabel() == "Bestätigen":

            # print("Action")

            if self.late == self.early == None:

                print("Du hast nichst ausgewählt. Tschüß")
                print(self.early)
                print(self.late)
                self.Destroy()

            elif self.late == None and self.early != None:

                print("not in between")
                print("early ", self.early)
                print("late ", self.late)
                self.Destroy()

                with wx.FileDialog(
                    self,
                    "Abruf abspeichern",
                    wildcard="Excel files (*.xlsx)|*.xlsx",
                    style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
                ) as fileDialog:

                    if fileDialog.ShowModal() == wx.ID_CANCEL:
                        return  # the user changed their mind

                    # save the current contents in the file
                    pathname = fileDialog.GetPath()

                    # make call to SQL type 1
                    Update([]).SQL_Abfrage1(
                        self.tables_sql, self.cols_sql, self.early, pathname
                    )

            elif self.late >= self.early:

                print("in between")
                print(self.early)
                print(self.late)
                self.Destroy()

                with wx.FileDialog(
                    self,
                    "Abruf abspeichern",
                    wildcard="Excel files (*.xlsx)|*.xlsx",
                    style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT,
                ) as fileDialog:

                    if fileDialog.ShowModal() == wx.ID_CANCEL:
                        return  # the user changed their mind

                    # save the current contents in the file
                    pathname = fileDialog.GetPath()

                    # make call to SQL type 1
                    Update([]).SQL_Abfrage2(
                        self.tables_sql, self.cols_sql, self.early, self.late, pathname
                    )

        elif e.GetEventObject().GetLabel() == "Abbrechen":

            print("Abbgebrochen")

            self.Destroy()

    def myListener_early(self, message):
        """
        Listener function
        """
        self.early = message

    def myListener_late(self, message):
        """
        Listener function
        """
        self.late = message

    def MakeModal(self, modal=True):
        if modal and not hasattr(self, "_disabler"):
            self._disabler = wx.WindowDisabler(self)
        if not modal and hasattr(self, "_disabler"):
            del self._disabler


#'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


class Calendar_early(wx.Frame):

    def __init__(self, parent, title):
        super(Calendar_early, self).__init__(parent, title=title)

        self.d1 = parent.d1
        self.d2 = parent.d2
        self.initUI()
        self.parent = parent

    def initUI(self):

        # force modality of the window
        self.MakeModal(True)
        # create panel
        self.panel = wx.Panel(self)
        # create calendar widget

        # add sizer
        hbox = wx.BoxSizer(wx.HORIZONTAL)

        cal = wx.adv.CalendarCtrl(self.panel, -1, style=wx.adv.CAL_MONDAY_FIRST)
        cal.SetDateRange(
            lowerdate=wx.DateTime(
                calendar.monthrange(self.d1.year, self.d1.month)[1],
                self.d1.month - 1,
                self.d1.year,
            ),
            upperdate=wx.DateTime(
                calendar.monthrange(self.d2.year, self.d2.month)[1],
                self.d2.month - 1,
                self.d2.year,
            ),
        )

        hbox.Add(cal, 0, flag=wx.ALL, border=5)

        # create topsizer
        topSizer = wx.BoxSizer(wx.VERTICAL)
        # add selection boxes and buttons to topsizer
        topSizer.Add(hbox, 0, wx.CENTER)
        # compute the panels size based on the topSizer
        self.panel.SetSizer(topSizer)
        # fit sizer
        topSizer.Fit(self)

        # event binding
        self.Bind(wx.adv.EVT_CALENDAR, self.OnDateSelected)
        self.Bind(wx.EVT_CLOSE, self.onClose)

        # show frame
        self.Center()
        self.Show()

    def OnDateSelected(self, e):

        date2send = e.GetDate()

        pub.sendMessage("early", message=date2send)
        self.Destroy()

    def onClose(self, e):

        self.Destroy()

    def MakeModal(self, modal=True):
        if modal and not hasattr(self, "_disabler"):
            self._disabler = wx.WindowDisabler(self)
        if not modal and hasattr(self, "_disabler"):
            del self._disabler


#''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


class Calendar_late(wx.Frame):

    def __init__(self, parent, title):
        super(Calendar_late, self).__init__(parent, title=title)

        self.d1 = parent.d1
        self.d2 = parent.d2
        self.initUI()
        self.parent = parent

    def initUI(self):

        # force modality of the window
        self.MakeModal(True)
        # create panel
        self.panel = wx.Panel(self)
        # create calendar widget

        # add sizer
        hbox = wx.BoxSizer(wx.HORIZONTAL)

        cal = wx.adv.CalendarCtrl(self.panel, -1, style=wx.adv.CAL_MONDAY_FIRST)
        cal.SetDateRange(
            lowerdate=wx.DateTime(
                calendar.monthrange(self.d1.year, self.d1.month)[1],
                self.d1.month - 1,
                self.d1.year,
            ),
            upperdate=wx.DateTime(
                calendar.monthrange(self.d2.year, self.d2.month)[1],
                self.d2.month - 1,
                self.d2.year,
            ),
        )

        hbox.Add(cal, 0, flag=wx.ALL, border=5)

        # create topsizer
        topSizer = wx.BoxSizer(wx.VERTICAL)
        # add selection boxes and buttons to topsizer
        topSizer.Add(hbox, 0, wx.CENTER)
        # compute the panels size based on the topSizer
        self.panel.SetSizer(topSizer)
        # fit sizer
        topSizer.Fit(self)

        # event binding
        self.Bind(wx.adv.EVT_CALENDAR, self.OnDateSelected)
        self.Bind(wx.EVT_CLOSE, self.onClose)

        # show frame
        self.Center()
        self.Show()

    def OnDateSelected(self, e):

        pub.sendMessage("late", message=e.GetDate())
        self.Destroy()

    def onClose(self, e):

        self.Destroy()

    def MakeModal(self, modal=True):
        if modal and not hasattr(self, "_disabler"):
            self._disabler = wx.WindowDisabler(self)
        if not modal and hasattr(self, "_disabler"):
            del self._disabler


#''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


class Median_GUI(wx.Frame):

    def __init__(self, *args, **kw):
        super(Median_GUI, self).__init__(*args, **kw)

        self.InitUI()

    def InitUI(self):

        if os.path.isdir("C:\\mypth"):
            pth_base = "C:\\mypth"

        else:
            sys.exit("unknown base path")

        pth = pth_base + "Der_Hammer\\"

        mpnl = MyPanel(self)

        font = wx.SystemSettings.GetFont(wx.SYS_SYSTEM_FONT)
        font.SetPointSize(7)

        vbox = wx.BoxSizer(wx.VERTICAL)

        bit1 = wx.StaticBitmap(
            self, wx.ID_ANY, wx.Bitmap(pth + "MEDIAN_Logo.png", wx.BITMAP_TYPE_ANY)
        )
        btn1 = MyButton(self, 1, "Datenbank aktualisieren")
        btn2 = MyButton(self, 2, "Wirtschaftsplan Dateien sammeln")
        btn3 = MyButton(self, 3, "Monatsberichte erstellen")
        btn4 = MyButton(self, 4, "Forecast Dateien sammeln")
        btn5 = MyButton(self, 5, "Abfrage")
        btn6 = MyButton(self, 6, "Allgemeine Datei Sammler")

        ##self was vbox before in the above code 6718-6723

        vbox.Add(bit1, 0, wx.EXPAND | wx.ALL, 5)
        vbox.Add(btn1, 0, wx.EXPAND | wx.ALL, 5)
        vbox.Add(btn2, 0, wx.EXPAND | wx.ALL, 5)
        vbox.Add(btn3, 0, wx.EXPAND | wx.ALL, 5)
        vbox.Add(btn4, 0, wx.EXPAND | wx.ALL, 5)
        vbox.Add(btn5, 0, wx.EXPAND | wx.ALL, 5)
        vbox.Add(btn6, 0, wx.EXPAND | wx.ALL, 5)

        vbox.SetSizeHints(self)
        self.SetSizer(vbox)

        self.Bind(wx.EVT_BUTTON, self.OnButtonClicked, id=btn1.GetId())
        self.Bind(wx.EVT_BUTTON, self.OnButtonClicked, id=btn2.GetId())
        self.Bind(wx.EVT_BUTTON, self.OnButtonClicked, id=btn3.GetId())
        self.Bind(wx.EVT_BUTTON, self.OnButtonClicked, id=btn4.GetId())
        self.Bind(wx.EVT_BUTTON, self.OnButtonClicked, id=btn5.GetId())
        self.Bind(wx.EVT_BUTTON, self.OnButtonClicked, id=btn6.GetId())

        self.SetTitle("Werkzeugskiste")
        self.Centre()

    def ShowMessage(self, e):
        txt = "Das wird etwa 40 Minuten dauern. Willst du fortfahren?"
        dial = wx.MessageDialog(
            None,
            txt,
            "WP_2019 Dateien sammlen?",
            wx.YES_NO | wx.NO_DEFAULT | wx.ICON_INFORMATION,
        )
        val = dial.ShowModal()

        if val == wx.ID_YES:

            # run WP_2019 method
            Update([]).WP_2019()

        elif val == wx.ID_NO:

            print("no")

    def OnButtonClicked(self, e):

        if e.GetId() == 1:

            print("Datenbank aktualisieren was pressed")
            DB_Akt(None, "Datenbank aktualisieren")

        elif e.GetId() == 2:

            print("Wirtschaftsplan Dateien sammeln was pressed")
            self.ShowMessage(e)

        elif e.GetId() == 3:

            print("Monatsberichte erstellen was pressed")

        elif e.GetId() == 4:

            print("Forecast Dateien sammeln was pressed")
            Update([]).Caller("FC_xxx")

        elif e.GetId() == 5:

            print("Abfrage was pressed")
            Abfrage(None, "Abfrage")

        elif e.GetId() == 6:

            print("Allgemeine Datei Sammler was pressed")

        e.Skip()


def main():

    app = wx.App()
    gui = Median_GUI(None)
    gui.Show()
    app.MainLoop()


if __name__ == "__main__":
    main()
