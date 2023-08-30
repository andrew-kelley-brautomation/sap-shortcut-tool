import pathlib
import re
import time
from tkinter import messagebox, simpledialog
import pythoncom
import win32com.client
import win32gui
import os
import parseConfig


def newTicket():
    try:
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
    except Exception as e:
        if session.findById("wnd[0]/sbar").Text != "":
            messagebox.showerror("SAP Tool", session.findById("wnd[0]/sbar").Text)
        else:
            messagebox.showerror("SAP Tool", "An error occurred processing this request.")


def recordMail(subject, timeSpent, attach, type, internal, separate, emailBodyText):
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
    if len(emailBodyText) <= 1:
        emailBody = outlookItem.Body
    else:
        emailBody = emailBodyText
    if len(tickets) < 1:
        messagebox.showerror("SAP Shortcut Error", "No ticket in subject line")
        return
    filepath = "C:/Temp/SAPEmails/"
    pathlib.Path(filepath).mkdir(parents=True, exist_ok=True)
    outlookItem.SaveAs(filepath + "emailForTicket.msg", 3)
    try:
        session = openSAP()
        if session is None:
            return
        for ticket in tickets:
            session.SendCommand("/n*IW52 RIWO00-QMNUM=" + ticket)
            if internal:
                session.findById("wnd[0]/shellcont/shell").clickLink("IDAT", "Column01")
            else:
                session.findById("wnd[0]/shellcont/shell").clickLink("MAIL", "Column01")
            session.findById("wnd[1]/usr/txtN_QMMA-MATXT").text = subject
            if internal:
                session.findById("wnd[1]/usr/cntlINTERN_DATA/shell").text = emailBody
            else:
                session.findById("wnd[1]/usr/cntlMAIL/shell").text = emailBody
            session.findById("wnd[1]/tbar[0]/btn[13]").press()
            session.findById("wnd[1]/usr/tblSAPLZCATS_UITC_CATS_TD/txtGS_ZSUPPORT_INPUT-ZSUP_MINUTES[3,0]").text = timeSpent
            if type != "00":
                session.findById("wnd[1]/usr/cmbZCATS_TS_EVAL_NOTIFICATION-ZEVAL_TYPE").Key = type
            session.findById("wnd[1]/tbar[0]/btn[15]").press()
            session.findById("wnd[0]/tbar[0]/btn[11]").press()
            if session.Children.Count > 1:
                session.findById("wnd[1]/usr/btnBUTTON_1").press()
            if attach:
                session.SendCommand("/n*IW52 RIWO00-QMNUM=" + ticket)
                session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem("ATAD", "Column01")
                session.findById("wnd[0]/shellcont/shell").clickLink("ATAD", "Column01")
                if internal:
                    maxPosition = session.findById("wnd[1]/usr").VerticalScrollbar.Maximum
                    position = 0
                    foundComm = False
                    while not foundComm:
                        labels = session.findById("wnd[1]/usr").Children
                        for label in labels:
                            beginIndex = label.ID.index(",")
                            row = label.ID[beginIndex + 1:len(label.ID) - 1]
                            if label.text == "internal communication created":
                                session.findById("wnd[1]/usr/chk[2," + row + "]").selected = True
                                foundComm = True
                                break
                        if position > maxPosition:
                            break
                        session.findById("wnd[1]/usr").VerticalScrollbar.Position += session.findById(
                            "wnd[1]/usr").VerticalScrollbar.PageSize
                        position = session.findById("wnd[1]/usr").VerticalScrollbar.Position
                else:
                    session.findById("wnd[1]/usr/chk[2,7]").selected = True
                session.findById("wnd[1]/tbar[0]/btn[18]").press()
                session.findById("wnd[2]/usr/btnATTACH_INSERT").press()
                session.findById("wnd[3]/usr/txtDY_PATH").text = filepath
                session.findById("wnd[3]/usr/txtDY_FILENAME").text = "emailForTicket.msg"
                session.findById("wnd[3]/tbar[0]/btn[0]").press()
                session.findById("wnd[2]/tbar[0]/btn[13]").press()
                session.findById("wnd[1]/tbar[0]/btn[13]").press()
                session.findById("wnd[0]/tbar[0]/btn[11]").press()
                if session.Children.Count > 1:
                    session.findById("wnd[1]/usr/btnBUTTON_1").press()
            if separate:
                numAttachments = outlookItem.Attachments.Count
                filenames = '"'
                files = []
                for num in range(1, numAttachments + 1):
                    attachment = outlookItem.Attachments.Item(num)
                    image = re.findall('image\d{3}', attachment.DisplayName)
                    if len(image) == 0:
                        files.append(attachment.DisplayName)
                        filenames += attachment.DisplayName + '" "'
                        attachment.SaveAsFile(filepath + attachment.DisplayName)
                if len(files) > 0:
                    session.SendCommand("/n*IW52 RIWO00-QMNUM=" + ticket)
                    session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem("ATAD", "Column01")
                    session.findById("wnd[0]/shellcont/shell").clickLink("ATAD", "Column01")
                    if internal:
                        maxPosition = session.findById("wnd[1]/usr").VerticalScrollbar.Maximum
                        position = 0
                        foundComm = False
                        while not foundComm:
                            labels = session.findById("wnd[1]/usr").Children
                            for label in labels:
                                beginIndex = label.ID.index(",")
                                row = label.ID[beginIndex + 1:len(label.ID) - 1]
                                if label.text == "internal communication created":
                                    session.findById("wnd[1]/usr/chk[2," + row + "]").selected = True
                                    foundComm = True
                                    break
                            if position > maxPosition:
                                break
                            session.findById("wnd[1]/usr").VerticalScrollbar.Position += session.findById(
                                "wnd[1]/usr").VerticalScrollbar.PageSize
                            position = session.findById("wnd[1]/usr").VerticalScrollbar.Position
                    else:
                        session.findById("wnd[1]/usr/chk[2,7]").selected = True
                    session.findById("wnd[1]/tbar[0]/btn[18]").press()
                    session.findById("wnd[2]/usr/btnATTACH_INSERT").press()
                    session.findById("wnd[3]/usr/txtDY_PATH").text = filepath
                    session.findById("wnd[3]/usr/txtDY_FILENAME").text = filenames[:-2]
                    session.findById("wnd[3]/tbar[0]/btn[0]").press()
                    session.findById("wnd[2]/tbar[0]/btn[13]").press()
                    session.findById("wnd[1]/tbar[0]/btn[13]").press()
                    session.findById("wnd[0]/tbar[0]/btn[11]").press()
                    if session.Children.Count > 1:
                        session.findById("wnd[1]/usr/btnBUTTON_1").press()
                    for file in files:
                        pathlib.Path(filepath + file).unlink()
        session.EndTransaction()
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
    except Exception as e:
        if session.findById("wnd[0]/sbar").Text != "":
            messagebox.showerror("SAP Tool", session.findById("wnd[0]/sbar").Text)
        else:
            messagebox.showerror("SAP Tool", "An error occurred processing this request.")
    pathlib.Path(filepath + "emailForTicket.msg").unlink()


def trackTime():
    try:
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
    except Exception as e:
        if session.findById("wnd[0]/sbar").Text != "":
            messagebox.showerror("SAP Tool", session.findById("wnd[0]/sbar").Text)
        else:
            messagebox.showerror("SAP Tool", "An error occurred processing this request.")


def displayTicket():
    try:
        ticketNum = simpledialog.askstring("SAP Shortcut Input", "Ticket Number:")
        if ticketNum is None:
            return
        session = openSAP()
        if session is None:
            return
        session.SendCommand("/n*IW53 RIWO00-QMNUM=" + ticketNum)
        session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem("CHRO", "Column01")
        session.findById("wnd[0]/shellcont/shell").clickLink("CHRO", "Column01")
    except Exception as e:
        if session.findById("wnd[0]/sbar").Text != "":
            messagebox.showerror("SAP Tool", session.findById("wnd[0]/sbar").Text)
        else:
            messagebox.showerror("SAP Tool", "An error occurred processing this request.")

def changeTicket():
    try:
        ticketNum = simpledialog.askstring("SAP Shortcut Input", "Ticket Number:")
        if ticketNum is None:
            return
        session = openSAP()
        if session is None:
            return
        session.SendCommand("/n*IW52 RIWO00-QMNUM=" + ticketNum)
    except Exception as e:
        print("Test")
        if session.findById("wnd[0]/sbar").Text != "":
            messagebox.showerror("SAP Tool", session.findById("wnd[0]/sbar").Text)
        else:
            messagebox.showerror("SAP Tool", "An error occurred processing this request.")


def mm03():
    try:
        modelNum = simpledialog.askstring("SAP Shortcut Input", "Model Number:")
        if modelNum is None:
            return
        session = openSAP()
        if session is None:
            return
        session.SendCommand("/n*MM03 RMMG1-MATNR=" + modelNum)
        session.findById("wnd[1]/usr/tblSAPLMGMMTC_VIEW").GetAbsoluteRow(0).Selected = True
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except Exception as e:
        if session.findById("wnd[0]/sbar").Text != "":
            messagebox.showerror("SAP Tool", session.findById("wnd[0]/sbar").Text)
        else:
            messagebox.showerror("SAP Tool", "An error occurred processing this request.")


def addTicketSolution(ticket, solution, timeSpent, close, addToBody):
    try:
       session = openSAP()
       if session is None:
           return
       session.SendCommand("/n*IW52 RIWO00-QMNUM=" + ticket)
       session.findById("wnd[0]/shellcont/shell").clickLink("LOVO", "Column01")
       session.findById("wnd[1]/usr/txtN_QMSM-MATXT").text = "Solution"
       session.findById("wnd[1]/usr/cntlLOESUNG/shell").text = solution
       session.findById("wnd[1]/tbar[0]/btn[13]").press()
       session.findById("wnd[1]/usr/tblSAPLZCATS_UITC_CATS_TD/txtGS_ZSUPPORT_INPUT-ZSUP_MINUTES[3,0]").text = timeSpent
       session.findById("wnd[1]/tbar[0]/btn[15]").press()
       if addToBody:
           textField = "wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell"
           subjText = ""
           for lineNum in range(session.findById(textField).LineCount + 1):
               subjText += session.findById(textField).GetLineText(lineNum) + "\n"
           subjText += "********************* Solution ******************\n"
           subjText += solution
       session.findById("wnd[0]/tbar[0]/btn[11]").press()
       if session.Children.Count > 1:
           session.findById("wnd[1]/usr/btnBUTTON_1").press()
       if close:
           session.SendCommand("/n*IW52 RIWO00-QMNUM=" + ticket)
           session.findById("wnd[0]/shellcont/shell").clickLink("ABGE", "Column01")
           time.sleep(1)
           if session.Children.Count > 1:
               session.findById("wnd[1]/usr/btnBUTTON_1").press()
           session.findById("wnd[0]/tbar[0]/btn[11]").press()
    except Exception as e:
        if session.findById("wnd[0]/sbar").Text != "":
            messagebox.showerror("SAP Tool", session.findById("wnd[0]/sbar").Text)
        else:
            messagebox.showerror("SAP Tool", "An error occurred processing this request.")


def zsupl4():
    try:
        session = openSAP()
        if session is None:
            return
        session.StartTransaction("ZSUPL4")
        session.findById("wnd[0]/usr/btn%_SO_INGRP_%_APP_%-VALU_PUSH").press()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "465"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "407"
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
    except Exception as e:
        if session.findById("wnd[0]/sbar").Text != "":
            messagebox.showerror("SAP Tool", session.findById("wnd[0]/sbar").Text)
        else:
            messagebox.showerror("SAP Tool", "An error occurred processing this request.")


def openSAP():
    pythoncom.CoInitialize()
    try:

        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except Exception as e:
        # print(str(e))
        # messagebox.showerror('SAP Shortcut Error', 'Please log in to SAP')
        # return None
        if parseConfig.parseConfig()['LOGIN'].getboolean('AUTO_LOGIN', True):
            parseConfig.makeBatch()
            count = 0
            while win32gui.FindWindow(None, 'License Information For Multiple Logons') == 0 and \
                    win32gui.FindWindow(None, 'SAP Easy Access') == 0:
                pass
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
        else:
            messagebox.showerror('SAP Shortcut Error', 'Please log in to SAP')
            return None
    if not type(SapGuiAuto) == win32com.client.CDispatch:
        messagebox.showerror('SAP Shortcut Error', 'Something went wrong')
        return None
    application = SapGuiAuto.GetScriptingEngine
    if not type(application) == win32com.client.CDispatch:
        messagebox.showerror('SAP Shortcut Error', 'Something went wrong')
        return None
    connection = application.Children(0)
    if not type(connection) == win32com.client.CDispatch:
        messagebox.showerror('SAP Shortcut Error', 'Something went wrong')
        return None
    numSessions = connection.Children.Count
    session = connection.Children(0)
    if not type(session) == win32com.client.CDispatch:
        messagebox.showerror('SAP Shortcut Error', 'Something went wrong')
        return None
    if connection.Children.Count > 5:
        messagebox.showerror('SAP Shortcut Error', 'Too many sessions open.\nPlease close an unneeded window')
        return None
    if win32gui.FindWindow(None, 'License Information For Multiple Logons') != 0:
        session.findById('wnd[1]/usr/radMULTI_LOGON_OPT2').select()
        session.findById('wnd[1]/tbar[0]/btn[0]').press()
    else:
        if session.findById("wnd[0]").Text == 'SAP Easy Access':
            return session
        session.CreateSession()
        while connection.Children.Count <= numSessions:
            pass
        ind = 1
        while not session.findById("wnd[0]").Text == 'SAP Easy Access':
            session = connection.Children(connection.Children.Count - ind)
            ind += 1
    return session

if __name__ == "__main__":
    session = openSAP()
    session.StartTransaction("ZSUPPORT")