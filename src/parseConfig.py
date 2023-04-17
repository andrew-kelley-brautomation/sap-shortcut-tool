from tkinter import messagebox
import configparser
import os

configSettings = {
        "MAIL": {
            "DEFAULT_SUBJECT": "L1 <> Customer",
            "DEFAULT_TIME": 5,
            "DEFAULT_ATTACH": True
        }
    }
def parseConfig():
    parser = None
    try:
        parser = configparser.ConfigParser()
        configFile = open("C:/SAP Shortcut Tool/config.ini", "r")
        parser.read_file(configFile)
        configFile.close()

    except FileNotFoundError:
        messagebox.showinfo("SAP Tool", "Unable to find config file, created default file.")
        if parser is None:
            parser = configparser.ConfigParser()
        os.makedirs("C:/SAP Shortcut Tool/", exist_ok=True)
        configFile = open("C:/SAP Shortcut Tool/config.ini", "w")
        configFile.write("[MAIL]\nDEFAULT_SUBJECT = L1 <> Customer\nDEFAULT_TIME: 5\nDEFAULT_ATTACH = True")
        configFile.close()
    return parser


if __name__ == "__main__":
    parseConfig()