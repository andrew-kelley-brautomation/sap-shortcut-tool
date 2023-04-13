import pathlib
import re
from tkinter import messagebox, simpledialog
import pythoncom
import win32com.client
import win32gui


def newTicket():
    pythoncom.CoInitialize()
    session = openSAP()
    if session is None:
        return
    session.sendCommand("/n*IW51 RIWO00-QMART=P1")
    while session.findById("wnd[0]").Text != "Create Service Notification: Initial Screen":
        pass
    win32gui.SetForegroundWindow(win32gui.FindWindow(None, 'please choose partners for notification'))
    # messagebox.showinfo("waiting for action", "Please select customer")
    # print("Messagebox passed")

def recordMail(timeSpent, attach):
    outlookObj = win32com.client.Dispatch('Outlook.Application')
    try:
        outlookItem = outlookObj.ActiveInspector().CurrentItem
    except:
        try:
            outlookItem = outlookObj.ActiveExplorer().Selection.Item(1)
        except:
            outlookItem = None
    if outlookItem is None or outlookItem.Class != 43:
        messagebox.showerror("SAP Shortcut Error", "Current Outlook Item Not An Email")
        return
    tickets = re.findall('400\d{6}', outlookItem.Subject)
    if len(tickets) < 1:
        messagebox.showerror("SAP Shortcut Error", "No ticket in subject line")
        return
    filepath = "C:/Temp/SAPEmails/"
    pathlib.Path(filepath).mkdir(parents=True, exist_ok=True)
    outlookItem.SaveAs(filepath + "emailForTicket.msg", 3)
    pythoncom.CoInitialize()
    session = openSAP()
    if session is None:
        return
    for ticket in tickets:
        session.SendCommand("/n*IW52 RIWO00-QMNUM=" + ticket)
        session.findById("wnd[0]/shellcont/shell").clickLink("MAIL", "Column01")
        session.findById("wnd[1]/usr/txtN_QMMA-MATXT").text = "L1 <> Customer"
        session.findById("wnd[1]/usr/cntlMAIL/shell").text = outlookItem.Body
        if attach:
            session.findById("wnd[1]/tbar[0]/btn[14]").press()
            session.findById("wnd[2]/usr/btnATTACH_INSERT").press()
            session.findById("wnd[3]/usr/txtDY_PATH").text = filepath
            session.findById("wnd[3]/usr/txtDY_FILENAME").text = "emailForTicket.msg"
            session.findById("wnd[3]/tbar[0]/btn[0]").press()
            session.findById("wnd[2]/tbar[0]/btn[13]").press()
        session.findById("wnd[1]/tbar[0]/btn[13]").press()
        session.findById("wnd[1]/usr/tblSAPLZCATS_UITC_CATS_TD/txtGS_ZSUPPORT_INPUT-ZSUP_MINUTES[3,0]").text = timeSpent
        session.findById("wnd[1]/tbar[0]/btn[15]").press()
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.EndTransaction()
    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    pathlib.Path(filepath + "emailForTicket.msg").unlink()


def recordMail_Andrew(timeSpent, attach):
    outlookObj = win32com.client.Dispatch('Outlook.Application')
    try:
        outlookItem = outlookObj.ActiveInspector().CurrentItem
    except:
        try:
            outlookItem = outlookObj.ActiveExplorer().Selection.Item(1)
        except:
            outlookItem = None
    if outlookItem is None or outlookItem.Class != 43:
        messagebox.showerror("SAP Shortcut Error", "Current Outlook Item Not An Email")
        return
    tickets = re.findall('400\d{6}', outlookItem.Subject)
    if len(tickets) < 1:
        messagebox.showerror("SAP Shortcut Error", "No ticket in subject line")
        return
    filepath = "C:/Temp/SAPEmails/"
    pathlib.Path(filepath).mkdir(parents=True, exist_ok=True)
    outlookItem.SaveAs(filepath + "emailForTicket.msg", 3)
    session = openSAP()
    if session is None:
        return
    for ticket in tickets:
        session.SendCommand("/n*IW52 RIWO00-QMNUM=" + ticket)
        session.findById("wnd[0]/shellcont/shell").clickLink("MAIL", "Column01")
        session.findById("wnd[1]/usr/txtN_QMMA-MATXT").text = "L1 <> Customer"
        session.findById("wnd[1]/usr/cntlMAIL/shell").text = outlookItem.Body
        session.findById("wnd[1]/tbar[0]/btn[13]").press()
        session.findById("wnd[1]/usr/tblSAPLZCATS_UITC_CATS_TD/txtGS_ZSUPPORT_INPUT-ZSUP_MINUTES[3,0]").text = timeSpent
        session.findById("wnd[1]/tbar[0]/btn[15]").press()
        session.findById("wnd[0]/tbar[0]/btn[11]").press()
        if attach:
            session.SendCommand("/n*IW52 RIWO00-QMNUM=" + ticket)
            session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem("ATAD", "Column01")
            session.findById("wnd[0]/shellcont/shell").clickLink("ATAD", "Column01")
            session.findById("wnd[1]/usr/chk[2,7]").selected = True
            session.findById("wnd[1]/tbar[0]/btn[18]").press()
            session.findById("wnd[2]/usr/btnATTACH_INSERT").press()
            session.findById("wnd[3]/usr/txtDY_PATH").text = filepath
            session.findById("wnd[3]/usr/txtDY_FILENAME").text = "emailForTicket.msg"
            session.findById("wnd[3]/tbar[0]/btn[0]").press()
            session.findById("wnd[2]/tbar[0]/btn[13]").press()
            session.findById("wnd[1]/tbar[0]/btn[13]").press()
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.EndTransaction()
    session.findById("wnd[0]/tbar[0]/btn[15]").press()
    pathlib.Path(filepath + "emailForTicket.msg").unlink()


def trackTime():
    pythoncom.CoInitialize()
    session = openSAP()
    if session is None:
        return
    session.startTransaction("CAT2")
    session.findById("wnd[0]/tbar[1]/btn[7]").press()
    session.findById("wnd[1]/usr/tabsTS_PROFILE/tabpPRGE/ssubPROFILE:SAPLCATS:3100/chkTCATS-TARGETROW").selected = True
    session.findById("wnd[1]/usr/tabsTS_PROFILE/tabpPRGE/ssubPROFILE:SAPLCATS:3100/chkTCATS-SUMROW").selected = True
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[0]/tbar[1]/btn[5]").press()
    while session.findById("wnd[0]").Text != "Time Sheet: Data Entry View":
        pass
    win32gui.SetForegroundWindow(win32gui.FindWindow(None, 'Time Sheet: Data Entry View'))


def displayTicket():
    ticketNum = simpledialog.askstring("SAP Shortcut Input", "Ticket Number:")
    if ticketNum is None:
        return
    session = openSAP()
    if session is None:
        return
    session.SendCommand("/n*IW53 RIWO00-QMNUM=" + ticketNum)
    session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem("CHRO", "Column01")
    session.findById("wnd[0]/shellcont/shell").clickLink("CHRO", "Column01")

def changeTicket():
    ticketNum = simpledialog.askstring("SAP Shortcut Input", "Ticket Number:")
    if ticketNum is None:
        return
    session = openSAP()
    if session is None:
        return
    session.SendCommand("/n*IW52 RIWO00-QMNUM=" + ticketNum)


def mm03():
    modelNum = simpledialog.askstring("SAP Shortcut Input", "Model Number:")
    if modelNum is None:
        return
    session = openSAP()
    if session is None:
        return
    session.SendCommand("/n*MM03 RMMG1-MATNR=" + modelNum)
    session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").GetAbsoluteRow(0).Selected = True
    session.findById("wnd[1]/tbar[0]/btn[0]").press()


def addTicketSolution(ticket, solution):
    session = openSAP()
    if session is None:
        return
    session.SendCommand("/n*IW52 RIWO00-QMNUM=" + ticket)
    session.findById("wnd[0]/shellcont/shell").clickLink("LOVO", "Column01")
    session.findById("wnd[1]/usr/txtN_QMSM-MATXT").text = "Solution"
    session.findById("wnd[1]/usr/cntlLOESUNG/shell").text = solution
    session.findById("wnd[1]/tbar[0]/btn[13]").press()
    session.findById("wnd[1]/usr/tblSAPLZCATS_UITC_CATS_TD/txtGS_ZSUPPORT_INPUT-ZSUP_MINUTES[3,0]").text = 5
    session.findById("wnd[1]/tbar[0]/btn[15]").press()
    session.findById("wnd[0]/tbar[0]/btn[11]").press()
    session.SendCommand("/n*IW52 RIWO00-QMNUM=" + ticket)
    session.findById("wnd[0]/shellcont/shell").clickLink("ABGE", "Column01")
    if session.Children.Count > 1:
        session.findById("wnd[1]/usr/btnBUTTON_1").press()
    session.findById("wnd[0]/tbar[0]/btn[11]").press()


def zsupl4():
    session = openSAP()
    if session is None:
        return
    session.StartTransaction("ZSUPL4")
    session.findById("wnd[0]/usr/btn%_SO_INGRP_%_APP_%-VALU_PUSH").press()
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "465"
    session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "407"
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[8]").press()


def openSAP():
    pythoncom.CoInitialize()
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except Exception as e:
        print(str(e))
        messagebox.showerror('SAP Shortcut Error', 'Please log in to SAP')
        return None
    if not type(SapGuiAuto) == win32com.client.CDispatch:
        return None
    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        return None
    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
        return None
    numSessions = connection.Children.Count
    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
        return None
    if connection.Children.Count > 5:
        messagebox.showerror('SAP Shortcut Error', 'Too many sessions open.\nPlease close an unneeded window')
        return None
    session.CreateSession()
    while not (connection.Children.Count > numSessions):
        pass
    session = connection.Children(connection.Children.Count - 1)
    return session


if __name__ == "__main__":
    zsupl4()
