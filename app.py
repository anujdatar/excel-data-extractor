#!/usr/bin/env python
import tkinter as tk

from src import MainAppGui

if __name__ == '__main__':
    root = tk.Tk()
    app_gui = MainAppGui(root)
    root.mainloop()
