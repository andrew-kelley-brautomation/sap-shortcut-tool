import keyboard
from main import *


def new_ticket():
    open_button_on_click()


def record_mail():
    mail_button_on_click()


def time_tracking():
    time_tracking_on_click()


def display():
    display_button_on_click()


def change():
    change_button_on_click()


def mm03():
    mm03_button_on_click()


def zsupl4():
    zsupl4_button_on_click()


def solution():
    solution_button_on_click()


def quick_create():
    quick_button_on_click()


if __name__ == "__main__":
    hotkeySettings = parseConfig.parseConfig()['HOTKEYS']
    keyboard.add_hotkey(hotkeySettings.get("NEW_TICKET", "ctrl+shift+q"), new_ticket)
    keyboard.add_hotkey(hotkeySettings.get("RECORD_MAIL", "ctrl+shift+w"), record_mail)
    keyboard.add_hotkey(hotkeySettings.get("TRACK_TIME", "ctrl+shift+e"), time_tracking)
    keyboard.add_hotkey(hotkeySettings.get("DISPLAY", "ctrl+shift+a"), display)
    keyboard.add_hotkey(hotkeySettings.get("CHANGE", "ctrl+shift+s"), change)
    keyboard.add_hotkey(hotkeySettings.get("MM03", "ctrl+shift+d"), mm03)
    keyboard.add_hotkey(hotkeySettings.get("SOLUTION", "ctrl+shift+z"), solution)
    keyboard.add_hotkey(hotkeySettings.get("TICKET_LIST", "ctrl+shift+x"), zsupl4)
    keyboard.add_hotkey(hotkeySettings.get("QUICK_CREATE", "ctrl+shift+c"), quick_create)
    keyboard.wait()


