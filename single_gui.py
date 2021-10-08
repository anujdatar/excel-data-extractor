import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

from run_data_extraction import single_extract


class MainAppGui(object):
    def __init__(self, master=None):
        self.master = master

        self.source_filename = tk.StringVar()
        self.target_filename = tk.StringVar()

        # self.master.minsize(width=640, height=480)
        self.master.resizable(True, True)
        self.master.title("Python Data Extractor")

        self.source_file_selector_label = ttk.Label(self.master, text="Select Source File")
        self.source_file_selector_label.grid(row=0, column=0, padx=5, pady=5)

        self.source_file_selector_entry = ttk.Entry(self.master, textvariable=self.source_filename)
        self.source_file_selector_entry.grid(row=0, column=1, ipadx=80, padx=5, pady=5)

        self.source_file_selector_button = ttk.Button(
            self.master,
            text="Browse",
            command=self.select_source_file
        )
        self.source_file_selector_button.grid(row=0, column=3, padx=5, pady=5)

        self.target_file_selector_label = ttk.Label(self.master, text="Select Target File")
        self.target_file_selector_label.grid(row=1, column=0, padx=5, pady=5)

        self.target_file_selector_entry = ttk.Entry(self.master, textvariable=self.target_filename)
        self.target_file_selector_entry.grid(row=1, column=1, ipadx=80, padx=5, pady=5)

        self.target_file_selector_button = ttk.Button(
            self.master,
            text="Browse",
            command=self.select_target_file
        )
        self.target_file_selector_button.grid(row=1, column=3, padx=5, pady=5)

        self.execute_button = ttk.Button(
            self.master,
            text="Execute",
            command=self.run_program
        )
        self.execute_button.grid(row=2, column=1, padx=20, pady=20)

    def select_source_file(self):
        """
        Open a file selection dialog box
        """
        filetypes = [
            ('Excel', '.xlsx')
        ]
        source_filename = filedialog.askopenfilename(
            title='Select Source File',
            initialdir='./',
            filetypes=filetypes
        )
        self.source_filename.set(source_filename)

    def select_target_file(self):
        """
        Open a file selection dialog box
        """
        filetypes = [
            ('Excel', '.xlsx')
        ]
        target_filename = filedialog.askopenfilename(
            title='Select Target File',
            initialdir='./',
            filetypes=filetypes
        )
        self.target_filename.set(target_filename)

    def run_program(self):
        # get file and folder name as strings
        source_filename = self.source_filename.get()
        target_filename = self.target_filename.get()

        single_extract(source_filename, target_filename)
