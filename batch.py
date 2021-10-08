import tkinter as tk

from batch_gui import MainAppGui
# from participant_class import Participant
# from write_to_target_sheet import write_to_target_sheet

datasheet_path = './SampleDataSheet.xlsx'
target_file_path = './Hebert_Data Variables.xlsx'

if __name__ == '__main__':
    root = tk.Tk()
    app_gui = MainAppGui(root)
    root.mainloop()
