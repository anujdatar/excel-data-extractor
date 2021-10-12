import os
import pathlib

from src import Participant
from src import write_to_target_sheet


def batch_extract(source_folder_name: str, target_filename: str):
    """
    Handle extraction in batch mode, from source folder
    :param source_folder_name: str -> absolute path of source folder containing excel files
    :param target_filename: str -> absolute path of target excel file
    """
    if not source_folder_name:
        print('please select source data folder')

    if not target_filename:
        print('please select target file')

    for file in os.listdir(source_folder_name):
        if pathlib.Path(file).suffix == '.xlsx':
            if '~$' not in pathlib.Path(file).name:
                filepath = source_folder_name + '/' + file
                extract(filepath, target_filename)


def single_extract(source_filename: str, target_filename: str):
    """
    Handle extraction in single file mode.
    :param source_filename: str -> absolute path of source excel file
    :param target_filename: str -> absolute path of target excel file
    """
    if not source_filename:
        print('please select source data file')

    if not target_filename:
        print('please select target file')

    if pathlib.Path(source_filename).suffix == '.xlsx':
        if '~$' not in pathlib.Path(source_filename).name:
            extract(source_filename, target_filename)


def extract(source_file: str, target_file: str):
    """
    Extract data from source excel file and write to target file
    :param source_file: str -> absolute path of source excel file
    :param target_file: str -> absolute path of target excel file
    """
    # fetch data
    participant_data = Participant(source_file)
    write_to_target_sheet(target_file, participant_data)
