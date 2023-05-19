import configparser
import subprocess
from tkinter import *
from tkinter import font
from tkinter.ttk import Combobox
from SAPfunctions import *
import math
import parseConfig

root = Tk()

root.title("SAP Shortcuts")

# Thanks to Chris Hairston for recommending the below graphics optimizations
root.geometry("")
graphicsSettings = parseConfig.parseConfig()['GRAPHICS']
defaultFont = font.nametofont("TkDefaultFont")
defaultFontType = defaultFont.actual("family")
defaultFontSize = defaultFont.actual("size")
scaledFontSize = math.ceil(defaultFontSize*float(graphicsSettings.get('SCALING', 1)))
scaledFont = (defaultFontType, scaledFontSize)

def open_button_on_click():
    newTicket()


def mail_button_on_click():

    sapTypes = {
        "Do not change": "00",
        "User Application Bug": "01",
        "Product Request": "02",
        "Hardware Question": "03",
        "Software Question": "04",
        "Guidance (Beratung)": "05",
        "System Software Bug": "06",
        "Hardware Bug (Single Failure)": "07",
        "Hardware Bug (Series Failure)": "08",
        "Documentation Missing, Insufficient, or Bug": "09",
        "Licensing Issue": "10",
        "Refer To Documentation": "11",
        "Prototype / Beta Support": "12",
        "No Rating Possible": "15",
        "Homepage (Downloads, Links)": "16",
        "Q-Reports": "20",
        "Q-Figures": "21",
        "Q-Management": "22",
        "Delivery Issues": "23",
    }
    child = Toplevel(root)
    attach = IntVar()
    internal = IntVar()
    subjLabel = Label(child, font=scaledFont)
    subj = Entry(child, font=scaledFont)
    mailSettings = parseConfig.parseConfig()['MAIL']
    subj.insert(0, mailSettings.get('DEFAULT_SUBJECT', "L1 <> Customer"))
    subjLabel.config(text=f"Subject: ({len(subj.get())}/40)")
    timeLabel = Label(child, text="Time Spent:", font=scaledFont)
    timeAmount = Entry(child, font=scaledFont)
    timeAmount.insert(0, mailSettings.getint('DEFAULT_TIME', 5))
    attachBox = Checkbutton(child, text="Attach Email to Ticket", variable=attach, font=scaledFont)
    internalBox = Checkbutton(child, text="Internal Communication", variable=internal, font=scaledFont)
    internalBox.deselect()
    if mailSettings.getboolean('DEFAULT_ATTACH'):
        attachBox.select()
    subjLabel.grid(column=2, row=2)
    subj.grid(column=2, row=3)
    timeLabel.grid(column=2, row=4)
    timeAmount.grid(column=2, row=5)
    attachBox.grid(column=2, row=6)
    internalBox.grid(column=2, row=7)
    selectedType = StringVar()
    typeSelector = Combobox(child, textvariable=selectedType, width=40)
    print(list(sapTypes.keys()))
    typeSelector['values'] = list(sapTypes.keys())
    typeSelector['state'] = 'readonly'
    typeSelector.current(0)
    typeSelector.grid(column=2, row=8)
    errorLabel = Label(child, fg="red")

    def validate_subject(subject):
        subjLabel.config(text=f"Subject: ({len(subject)}/40)")
        if len(subject) > 40:
            errorLabel.config(text="Subject must be less than 40 characters")
            errorLabel.grid(column=2, row=1)
            if mailSettings.getboolean('STOP_AT_FORTY', False):
                return False
        else:
            errorLabel.grid_remove()
        return True

    # if mailSettings.getboolean('STOP_AT_FORTY', False):
    validation = root.register(validate_subject)
    subj.config(validate="key", validatecommand=(validation, '%P'))

    def cont(event=None):
        try:
            timeSpent = int(timeAmount.get())
            subjectText = subj.get()
            if len(subjectText) <= 40:
                child.destroy()
                recordMail(subjectText, timeSpent, True if attach.get() == 1 else False,
                           sapTypes.get(selectedType.get()), True if internal.get() == 1 else False)
            else:
                errorLabel.config(text=f"Subject must be less than 40 characters (Currently: {len(subjectText)})")
                errorLabel.grid(column=2, row=1)
        except ValueError as e:
            errorLabel.config(text="Please enter an integer time quantity")
            errorLabel.grid(column=2, row=1)

    child.bind("<Return>", cont)
    contButton = Button(child, text="Continue", height=1, width=60, bd=5, command=cont, font=scaledFont)
    contButton.grid(column=2, row=9)


def time_tracking_on_click():
    trackTime()


def display_button_on_click():
    displayTicket()


def change_button_on_click():
    changeTicket()


def mm03_button_on_click():
    mm03()


def zsupl4_button_on_click():
    zsupl4()


def solution_button_on_click():
    child = Toplevel(root)
    solutionSettings = parseConfig.parseConfig()['SOLUTION']
    tktLabel = Label(child, text="Ticket Number:", font=scaledFont)
    tktNum = Entry(child, font=scaledFont)
    timeLabel = Label(child, text="Time Spent:", font=scaledFont)
    timeAmount = Entry(child)
    close = IntVar()
    closeBox = Checkbutton(child, text="Close Ticket", variable=close, font=scaledFont)
    addToBody = IntVar()
    addToBodyBox = Checkbutton(child, text="Add to Ticket Body", variable=close, font=scaledFont)
    if solutionSettings.getboolean('DEFAULT_CLOSE'):
        closeBox.select()
    timeAmount.insert(0, solutionSettings.getint('DEFAULT_TIME', 5))
    solLabel = Label(child, text="Solution:", font=scaledFont)
    solution = Text(child, width=60, height=10, font=scaledFont)
    tktLabel.grid(column=2, row=1)
    tktNum.grid(column=2, row=2)
    timeLabel.grid(column=2, row=3)
    timeAmount.grid(column=2, row=4)
    closeBox.grid(column=2, row=5)
    addToBodyBox.grid(column=2, row=6)
    solLabel.grid(column=2, row=7)
    solution.grid(column=2, row=8)

    def ticket_solution():
        ticketNum = tktNum.get()
        solutionText = solution.get("1.0", END)
        try:
            timeSpent = int(timeAmount.get())
        except ValueError as e:
            timeSpent = 5
        child.destroy()
        addTicketSolution(ticketNum, solutionText, timeSpent,
                          True if close.get() == 1 else False, True if addToBody.get() == 1 else False)

    contButton = Button(child, text="Continue", height=1, width=60, bd=5, command=ticket_solution, font=scaledFont)
    contButton.grid(column=2, row=9)


def settings_button_on_click():
    subprocess.Popen(["notepad.exe", "C:/SAP Shortcut Tool/config.ini"])


buttonWidth = 15
buttonHeight = 1

openBtn = Button(root, text="Open New Ticket", fg="blue", height=buttonHeight,
                 width=buttonWidth, command=open_button_on_click, font=scaledFont)
openBtn.grid(column=2, row=2)

mailBtn = Button(root, text="Record Mail", fg="blue", height=buttonHeight,
                 width=buttonWidth, command=mail_button_on_click, font=scaledFont)
mailBtn.grid(column=3, row=2)

timeBtn = Button(root, text="Track Time", fg="blue", height=buttonHeight,
                 width=buttonWidth, command=time_tracking_on_click, font=scaledFont)
timeBtn.grid(column=4, row=2)

displayBtn = Button(root, text="Display Ticket", fg="blue", height=buttonHeight,
                    width=buttonWidth, command=display_button_on_click, font=scaledFont)
displayBtn.grid(column=2, row=3)

changeBtn = Button(root, text="Change Ticket", fg="blue", height=buttonHeight,
                   width=buttonWidth, command=change_button_on_click, font=scaledFont)
changeBtn.grid(column=3, row=3)

mm03Button = Button(root, text="MM03", fg="blue", height=buttonHeight,
                    width=buttonWidth, command=mm03_button_on_click, font=scaledFont)
mm03Button.grid(column=4, row=3)

solButton = Button(root, text="Solution", fg="blue", height=buttonHeight,
                   width=buttonWidth, command=solution_button_on_click, font=scaledFont)
solButton.grid(column=2, row=4)

zsupl4Button = Button(root, text="Ticket List", fg="blue", height=buttonHeight,
                      width=buttonWidth, command=zsupl4_button_on_click, font=scaledFont)
zsupl4Button.grid(column=3, row=4)

settingsButton = Button(root, text="Edit Settings", fg="blue", height=buttonHeight,
                      width=buttonWidth, command=settings_button_on_click, font=scaledFont)
settingsButton.grid(column=4, row=4)

root.mainloop()
