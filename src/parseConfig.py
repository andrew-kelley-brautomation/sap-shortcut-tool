import subprocess
from tkinter import messagebox, simpledialog
import configparser
import os

configSettings = {
    "MAIL": {
        "DEFAULT_SUBJECT": "L1 <> Customer",
        "DEFAULT_TIME": "5",
        "DEFAULT_ATTACH": "True",
        "STOP_AT_FORTY": "True",
    },
    "GRAPHICS": {
        "SCALING": "1",
    },
    "SOLUTION": {
        "DEFAULT_TIME": "5",
        "DEFAULT_CLOSE": "True",
    },
    "LOGIN": {
        "AUTO_LOGIN": "True",
    },
    "HOTKEYS": {
        "NEW_TICKET": "ctrl+shift+q",
        "RECORD_MAIL": "ctrl+shift+w",
        "TRACK_TIME": "ctrl+shift+e",
        "DISPLAY": "ctrl+shift+a",
        "CHANGE": "ctrl+shift+s",
        "MM03": "ctrl+shift+d",
        "SOLUTION": "ctrl+shift+z",
        "TICKET_LIST": "ctrl+shift+x",
        "QUICK_CREATE": "ctrl+shift+c",
    },
    "SAP": {
        "PERSONNEL_NUMBER": "10212",
    },
}

def parseConfig():
    parser = configparser.ConfigParser()
    try:
        configFile = open("C:/SAP Shortcut Tool/config.ini", "r")
        parser.read_file(configFile)
        configFile.close()
    except FileNotFoundError:
        messagebox.showinfo("SAP Tool", "Unable to find config file, created default file.")
        os.makedirs("C:/SAP Shortcut Tool/", exist_ok=True)
    for section, settings in configSettings.items():
        if not parser.has_section(section):
            parser.add_section(section)
        for option, value in settings.items():
            if not parser.has_option(section, option):
                if option == 'PERSONNEL_NUMBER':
                    value = simpledialog.askstring('SAP Tool', 'SAP Personnel Number:')
                parser.set(section, option, value)
    configFile = open("C:/SAP Shortcut Tool/config.ini", "w")
    parser.write(configFile)
    configFile.close()
    return parser

def makeBatch():
    try:
        batchFile = open("C:/SAP Shortcut Tool/openSAP.bat", 'r')
    except FileNotFoundError:
        messagebox.showinfo("SAP Tool", "Unable to find batch file, creating file.")
        os.makedirs("C:/SAP Shortcut Tool/", exist_ok=True)
        username = simpledialog.askstring("SAP Tool", "SAP Username:")
        password = simpledialog.askstring("SAP Tool", "SAP Password:")
        batchFile = open("C:/SAP Shortcut Tool/openSAP.bat", 'w')
        batchFile.write('start C:/"Program Files (x86)"/SAP/FrontEnd/SAPgui/sapshcut -sysname="B&R Production System" '
                        '-client=444 -user=' + username + ' -pw=' + password)
    batchFile.close()
    subprocess.Popen('C:/SAP Shortcut Tool/openSAP.bat')

if __name__ == "__main__":
    parseConfig()
