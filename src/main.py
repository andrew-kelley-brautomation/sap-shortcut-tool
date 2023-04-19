from tkinter import *
from tkinter.ttk import Combobox

from SAPfunctions import *

root = Tk()

root.title("SAP Shortcuts")
root.geometry("348x79")


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
    subjLabel = Label(child, text="Subject:")
    subj = Entry(child, width=40)
    subj.insert(0, "L1 <> Customer")
    timeLabel = Label(child, text="Time Spent:")
    timeAmount = Entry(child)
    attachBox = Checkbutton(child, text="Attach Email to Ticket", variable=attach)
    attachBox.select()
    subjLabel.grid(column=2, row=2)
    subj.grid(column=2, row=3)
    timeLabel.grid(column=2, row=4)
    timeAmount.grid(column=2, row=5)
    attachBox.grid(column=2, row=6)
    selectedType = StringVar()
    typeSelector = Combobox(child, textvariable=selectedType, width=40)
    print(list(sapTypes.keys()))
    typeSelector['values'] = list(sapTypes.keys())
    typeSelector['state'] = 'readonly'
    typeSelector.current(0)
    typeSelector.grid(column=2, row=7)

    def cont(event=None):
        try:
            timeSpent = int(timeAmount.get())
            subjectText = subj.get()
            if len(subjectText) <= 40:
                child.destroy()
                recordMail(subjectText, timeSpent, True if attach.get() == 1 else False, sapTypes.get(selectedType.get()))
            else:
                errorLabel = Label(child, text="Subject must be less than 40 characters", fg="red")
                errorLabel.grid(column=2, row=1)
        except ValueError as e:
            errorLabel = Label(child, text="Please enter an integer time quantity", fg="red")
            errorLabel.grid(column=2, row=1)

    child.bind("<Return>", cont)
    contButton = Button(child, text="Continue", height=1, width=60, bd=5, command=cont)
    contButton.grid(column=2, row=8)


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
    solLabel = Label(child, text="Solution:")
    solution = Text(child, width=60, height=10)
    tktLabel.grid(column=2, row=1)
    tktNum.grid(column=2, row=2)
    solLabel.grid(column=2, row=3)
    solution.grid(column=2, row=4)

    def ticket_solution():
        ticketNum = tktNum.get()
        solutionText = solution.get("1.0", END)
        child.destroy()
        addTicketSolution(ticketNum, solutionText)

    contButton = Button(child, text="Continue", height=1, width=60, bd=5, command=ticket_solution)
    contButton.grid(column=2, row=5)


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
