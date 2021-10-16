import tkinter as tk
from tkinter import ttk

from src import SingleAppGui, BatchAppGui


def destroy_child_widgets(widget_group: tk.Widget) -> None:
    """
    Delete all child widgets within from a tkinter parent object
    :param widget_group: tk.Widget -> Any tkinter widget that can contain child
            elements. Tk, Frame, LabelFrame, etc.
    """
    for widget in widget_group.winfo_children():
        widget.destroy()


class MainAppGui:
    """
    Class for containing all GUI elements.
    Same window to show both single and batch extraction widgets
    """
    def __init__(self, master: tk.Tk):
        self.master = master
        self.master.title('Python Data Extractor')
        self.master.resizable(False, False)

        self.menu_bar = tk.Menu(self.master)
        self.master.config(menu=self.menu_bar)
        self.options_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.set_menu_items()

        self.central_widget = tk.Frame(self.master)
        self.central_widget.grid()

        self.single_gui_button = ttk.Button(
            self.central_widget,
            text='Single File Extraction',
            command=self.launch_single
        )
        self.single_gui_button.grid(row=0, column=0, padx=20, pady=20)

        self.batch_gui_button = ttk.Button(
            self.central_widget,
            text='Batch Extraction',
            command=self.launch_batch
        )
        self.batch_gui_button.grid(row=0, column=1, padx=20, pady=20)

    def set_menu_items(self) -> None:
        """Sets the application menu bar"""
        self.options_menu.add_command(
            label='Single File Extraction',
            command=self.launch_single
        )
        self.options_menu.add_command(
            label='Batch Extraction',
            command=self.launch_batch
        )
        self.options_menu.add_separator()
        self.options_menu.add_command(label='Exit', command=self.master.quit)

        self.menu_bar.add_cascade(label='Options', menu=self.options_menu)

    def launch_single(self) -> None:
        """Show the single file extraction widget in the main window"""
        print('launching single extraction gui')
        self.master.title('Python Data Extractor - Single')

        destroy_child_widgets(self.central_widget)
        SingleAppGui(self.central_widget)

    def launch_batch(self) -> None:
        """Show the batch file extraction widget in the main window"""
        print('launching batch extraction gui')
        self.master.title('Python Data Extractor - Batch')

        destroy_child_widgets(self.central_widget)
        BatchAppGui(self.central_widget)
