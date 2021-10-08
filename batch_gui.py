import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

from run_data_extraction import batch_extract


class MainAppGui(object):
    def __init__(self, master=None):
        self.master = master

        self.folder_name = tk.StringVar()
        self.filename = tk.StringVar()

        # self.master.minsize(width=640, height=480)
        self.master.resizable(True, True)
        self.master.title("Python Data Extractor")

        self.folder_selector = ttk.Label(self.master, text="Select Data")
        self.folder_selector.grid(row=0, column=0, padx=5, pady=5)

        self.folder_selector_entry = ttk.Entry(self.master, textvariable=self.folder_name)
        self.folder_selector_entry.grid(row=0, column=1, ipadx=80, padx=5, pady=5)

        self.folder_selector_button = ttk.Button(
            self.master,
            text="Select Folder",
            command=self.select_source_folder
        )
        self.folder_selector_button.grid(row=0, column=3, padx=5, pady=5)

        self.file_selector = ttk.Label(self.master, text="Select Target File")
        self.file_selector.grid(row=1, column=0, padx=5, pady=5)

        self.file_selector_entry = ttk.Entry(self.master, textvariable=self.filename)
        self.file_selector_entry.grid(row=1, column=1, ipadx=80, padx=5, pady=5)

        self.file_selector_button = ttk.Button(
            self.master,
            text="Select File",
            command=self.select_target_file
        )
        self.file_selector_button.grid(row=1, column=3, padx=5, pady=5)

        self.execute_button = ttk.Button(
            self.master,
            text="Execute",
            command=self.run_program
        )
        self.execute_button.grid(row=2, column=1, padx=20, pady=20)

    def select_source_folder(self):
        """
        Open a folder selection dialog box
        """
        folder_name = filedialog.askdirectory(
            title='Select Folder',
            initialdir='./'
        )
        self.folder_name.set(folder_name)

    def select_target_file(self):
        """
        Open a file selection dialog box
        """
        filetypes = [
            ('Excel', '.xlsx')
        ]
        filename = filedialog.askopenfilename(
            title='Select Target File',
            initialdir='./',
            filetypes=filetypes
        )
        self.filename.set(filename)

    def run_program(self):
        # get file and folder name as strings
        data_folder_name = self.folder_name.get()
        target_filename = self.filename.get()

        batch_extract(data_folder_name, target_filename)
