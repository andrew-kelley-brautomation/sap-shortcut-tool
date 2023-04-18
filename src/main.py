from tkinter import *
from SAPfunctions import *
from settings import *

root = Tk()

root.title("SAP Shortcuts")
root.geometry("348x79")

initSettings = SAP_Settings()



def open_button_on_click():
    newTicket()


def mail_button_on_click():
    child = Toplevel(root)
    attach = IntVar()
    subjLabel = Label(child, text="Subject:")
    subj = Entry(child)
    subj.insert(0, initSettings.defaultRecordMailSubject)
    timeLabel = Label(child, text="Time Spent:")
    timeAmount = Entry(child)
    timeAmount.insert(0, initSettings.defaultRecordMailTime)
    attachBox = Checkbutton(child, text="Attach Email to Ticket", variable=attach)
    attachBox.select()
    subjLabel.grid(column=2, row=2)
    subj.grid(column=2, row=3)
    timeLabel.grid(column=2, row=4)
    timeAmount.grid(column=2, row=5)
    attachBox.grid(column=2, row=6)

    def cont(event):
        try:
            timeSpent = int(timeAmount.get())
            subjectText = subj.get()
            child.destroy()
            recordMail(subjectText, timeSpent, True if attach.get() == 1 else False)
        except ValueError as e:
            errorLabel = Label(child, text="Please enter a time quantity", fg="red")
            errorLabel.grid(column=2, row=1)

    child.bind("<Return>", cont)
    contButton = Button(child, text="Continue", height=1, width=60, bd=5, command=cont)
    contButton.grid(column=2, row=7)



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
    tktLabel = Label(child, text="Ticket Number:")
    tktNum = Entry(child)
    timeLabel = Label(child, text="Time Spent:")
    timeAmount = Entry(child)
    timeAmount.insert(0, initSettings.defaultSolutionTime)
    solLabel = Label(child, text="Solution:")
    solution = Text(child, width=60, height=10)
    tktLabel.grid(column=2, row=1)
    tktNum.grid(column=2, row=2)
    timeLabel.grid(column=2, row=3)
    timeAmount.grid(column=2, row=4)
    solLabel.grid(column=2, row=5)
    solution.grid(column=2, row=6)
    

    def ticket_solution():
        ticketNum = tktNum.get()
        solutionText = solution.get("1.0", END)
        try:
            timeSpent = int(timeAmount.get())
        except ValueError as e:
            timeSpent = 5
        child.destroy()
        addTicketSolution(ticketNum, solutionText, timeSpent)

    contButton = Button(child, text="Continue", height=1, width=60, bd=5, command=ticket_solution)
    contButton.grid(column=2, row=7)


buttonWidth = 15
buttonHeight = 1

openBtn = Button(root, text="Open New Ticket", fg="blue", height=buttonHeight, width=buttonWidth, command=open_button_on_click)
openBtn.grid(column=2, row=2)

mailBtn = Button(root, text="Record Mail", fg="blue", height=buttonHeight, width=buttonWidth, command=mail_button_on_click)
mailBtn.grid(column=3, row=2)

timeBtn = Button(root, text="Track Time", fg="blue", height=buttonHeight, width=buttonWidth, command=time_tracking_on_click)
timeBtn.grid(column=4, row=2)

displayBtn = Button(root, text="Display Ticket", fg="blue", height=buttonHeight, width=buttonWidth, command=display_button_on_click)
displayBtn.grid(column=2, row=3)

changeBtn = Button(root, text="Change Ticket", fg="blue", height=buttonHeight, width=buttonWidth, command=change_button_on_click)
changeBtn.grid(column=3, row=3)

mm03Button = Button(root, text="MM03", fg="blue", height=buttonHeight, width=buttonWidth, command=mm03_button_on_click)
mm03Button.grid(column=4, row=3)

solButton = Button(root, text="Solution", fg="blue", height=buttonHeight, width=buttonWidth, command=solution_button_on_click)
solButton.grid(column=2, row=4)

zsupl4Button = Button(root, text="Ticket List", fg="blue", height=buttonHeight, width=buttonWidth, command=zsupl4_button_on_click)
zsupl4Button.grid(column=3, row=4)

root.mainloop()
