import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

from abc import ABC, abstractmethod

from src import batch_extract, single_extract


class MainAppGui(ABC):
    """Base GUI class for the app"""
    def __init__(self, master=None):
        self.master = master
        self.master.resizable(True, True)

        self.source_name = tk.StringVar()
        self.target_name = tk.StringVar()

        # select source file or folder
        # must define self.source_label.text on child class
        self.source_label = ttk.Label(self.master)
        self.source_label.grid(row=0, column=0, padx=5, pady=5)

        self.source_entry = ttk.Entry(self.master, textvariable=self.source_name)
        self.source_entry.grid(row=0, column=1, ipadx=80, padx=5, pady=5)

        self.source_button = ttk.Button(
            self.master,
            text='Browse',
            command=self.select_source
        )
        self.source_button.grid(row=0, column=2, padx=5, pady=5)

        # select target file
        self.target_label = ttk.Label(self.master, text='Select Target File')
        self.target_label.grid(row=1, column=0, padx=5, pady=5)

        self.target_entry = ttk.Entry(self.master, textvariable=self.target_name)
        self.target_entry.grid(row=1, column=1, ipadx=80, padx=5, pady=5)

        self.target_button = ttk.Button(
            self.master,
            text='Browse',
            command=self.select_target
        )
        self.target_button.grid(row=1, column=2, padx=5, pady=5)

        # execute program
        self.execute_button = ttk.Button(
            self.master,
            text='Execute',
            command=self.run_program
        )
        self.execute_button.grid(row=2, column=1, padx=20, pady=20)

    @abstractmethod
    def select_source(self) -> None:
        """Abstract method to select source file or folder"""
        pass

    @abstractmethod
    def run_program(self) -> None:
        """Abstract method for running whatever program is required"""
        pass

    def select_target(self) -> None:
        """Open a file selection dialog to select target excel file"""
        filetypes = [
            ('Excel', '.xlsx')
        ]
        target_filename = filedialog.askopenfilename(
            title='Select Target File',
            initialdir='./',
            filetypes=filetypes
        )
        self.target_name.set(target_filename)


class BatchAppGui(MainAppGui):
    """GUI for batch extraction app"""
    def __init__(self, master=None):
        super().__init__(master)
        self.master.title('Python Data Extractor - Batch')

        self.source_label.config(text='Select Source Folder')

    def select_source(self) -> None:
        source_folder_name = filedialog.askdirectory(
            title='Select Folder',
            initialdir='./'
        )
        self.source_name.set(source_folder_name)

    def run_program(self) -> None:
        source_folder_name = self.source_name.get()
        target_filename = self.target_name.get()

        batch_extract(source_folder_name, target_filename)


class SingleAppGui(MainAppGui):
    """GUI for batch extraction app"""
    def __init__(self, master=None):
        super().__init__(master)
        self.master.title('Python Data Extractor - Single')

        self.source_label.config(text='Select Source File')

    def select_source(self) -> None:
        filetypes = [
            ('Excel', '.xlsx')
        ]
        source_filename = filedialog.askopenfilename(
            title='Select Source File',
            initialdir='./',
            filetypes=filetypes
        )
        self.source_name.set(source_filename)

    def run_program(self) -> None:
        source_filename = self.source_name.get()
        target_filename = self.target_name.get()

        single_extract(source_filename, target_filename)
