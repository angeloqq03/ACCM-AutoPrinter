import os
from tkinter import *
import sys
from LegalFormApp import LegalFormApp as LegalFormApplication

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

if __name__ == "__main__":
    root = Tk()
    root.withdraw()
    app = LegalFormApplication(root, resource_path('front_page.png'))
    root.mainloop()