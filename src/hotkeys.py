import keyboard
from SAPfunctions import *


def function1():
    newTicket()


# Define your hotkey combination
if __name__ == "__main__":
    hotkey = "ctrl+alt+p"

    # Add the hotkey event listener
    keyboard.add_hotkey(hotkey, function1)

    # Start the event listener loop
    keyboard.wait()


