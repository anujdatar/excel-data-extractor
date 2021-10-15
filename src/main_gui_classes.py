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

        self.variables_filename = tk.StringVar()
        self.source_name = tk.StringVar()
        self.target_name = tk.StringVar()
        self.target_sheet_name = tk.StringVar()

        # select variables definition csv file
        self.variables_file_label = ttk.Label(self.master, text='Select Variables CSV')
        self.variables_file_label.grid(row=0, column=0, padx=5, pady=5, sticky='W')

        self.variables_file_entry = ttk.Entry(
            self.master,
            textvariable=self.variables_filename
        )
        self.variables_file_entry.grid(row=0, column=1, ipadx=80, padx=5, pady=5)

        self.variables_file_button = ttk.Button(
            self.master,
            text='Browse',
            command=self.select_variables_file
        )
        self.variables_file_button.grid(row=0, column=2, padx=5, pady=5)

        # select source file or folder
        self.source_label = ttk.Label(self.master)
        self.source_label.grid(row=1, column=0, padx=5, pady=5, sticky='W')
        self.set_source_label()

        self.source_entry = ttk.Entry(self.master, textvariable=self.source_name)
        self.source_entry.grid(row=1, column=1, ipadx=80, padx=5, pady=5)

        self.source_button = ttk.Button(
            self.master,
            text='Browse',
            command=self.select_source
        )
        self.source_button.grid(row=1, column=2, padx=5, pady=5)

        # select target file
        self.target_label = ttk.Label(self.master, text='Select Target File')
        self.target_label.grid(row=2, column=0, padx=5, pady=5, sticky='W')

        self.target_entry = ttk.Entry(self.master, textvariable=self.target_name)
        self.target_entry.grid(row=2, column=1, ipadx=80, padx=5, pady=5)

        self.target_button = ttk.Button(
            self.master,
            text='Browse',
            command=self.select_target
        )
        self.target_button.grid(row=2, column=2, padx=5, pady=5)

        # enter target worksheet name
        self.target_sheet_label = ttk.Label(
            self.master,
            text='Enter Target Sheet Name'
        )
        self.target_sheet_label.grid(row=3, column=0, padx=5, pady=5, sticky='W')

        self.target_sheet_entry = ttk.Entry(
            self.master,
            textvariable=self.target_sheet_name
        )
        self.target_sheet_entry.grid(row=3, column=1, ipadx=40, padx=5, pady=5, sticky='W')

        # execute program
        self.execute_button = ttk.Button(
            self.master,
            text='Execute',
            command=self.run_program
        )
        self.execute_button.grid(row=4, column=1, padx=20, pady=20)

    @abstractmethod
    def set_source_label(self) -> None:
        """Abstract method, forces setting source label in child class"""
        pass

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

    def select_variables_file(self) -> None:
        """Open a file selection dialog to select variables definition csv"""
        filetypes = [
            ('CSV', '.csv')
        ]
        variables_filename = filedialog.askopenfilename(
            title='Select Variables Definition File',
            initialdir='./',
            filetypes=filetypes
        )
        self.variables_filename.set(variables_filename)

    def clear_input_fields(self) -> None:
        """Clear all input/entry fields"""
        self.variables_filename.set('')
        self.source_name.set('')
        self.target_name.set('')
        self.target_sheet_name.set('')


class BatchAppGui(MainAppGui):
    """GUI for batch extraction app"""
    def __init__(self, master=None):
        super().__init__(master)
        self.master.title('Python Data Extractor - Batch')

    def set_source_label(self):
        self.source_label.config(text='Select Source Folder')

    def select_source(self) -> None:
        """Open tk file dialogue to select source directory"""
        source_folder_name = filedialog.askdirectory(
            title='Select Folder',
            initialdir='./'
        )
        self.source_name.set(source_folder_name)

    def run_program(self) -> None:
        """Execute batch data extraction for all files in source folder"""
        variables_filename = self.variables_filename.get()
        source_folder_name = self.source_name.get()
        target_filename = self.target_name.get()
        target_sheet = self.target_sheet_name.get()

        batch_extract(
            variables_filename,
            source_folder_name,
            target_filename,
            target_sheet
        )
        self.clear_input_fields()


class SingleAppGui(MainAppGui):
    """GUI for batch extraction app"""
    def __init__(self, master=None):
        super().__init__(master)
        self.master.title('Python Data Extractor - Single')

    def set_source_label(self):
        self.source_label.config(text='Select Source File')

    def select_source(self) -> None:
        """Open tk file dialogue to select source excel file"""
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
        """Execute data extraction on selected source file"""
        variables_filename = self.variables_filename.get()
        source_filename = self.source_name.get()
        target_filename = self.target_name.get()
        target_sheet = self.target_sheet_name.get()

        single_extract(
            variables_filename,
            source_filename,
            target_filename,
            target_sheet
        )
        self.clear_input_fields()
