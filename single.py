import tkinter as tk

from src import SingleAppGui

if __name__ == '__main__':
    root = tk.Tk()
    app_gui = SingleAppGui(root)
    root.mainloop()
