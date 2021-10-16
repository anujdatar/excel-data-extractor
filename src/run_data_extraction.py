import os
import pathlib

from src import extract


def batch_extract(
        vars_file: str,
        source_folder_name: str,
        target_filename: str,
        target_sheet: str) -> None:
    """
    Handle extraction in batch mode, from source folder
    :param vars_file: str -> absolute path to variable definition csv
    :param source_folder_name: str -> absolute path of source folder containing excel files
    :param target_filename: str -> absolute path of target excel file
    :param target_sheet: str -> name of worksheet in target excel workbook
    """
    if not source_folder_name:
        print('please select source data folder')

    if not target_filename:
        print('please select target file')

    for file in os.listdir(source_folder_name):
        if pathlib.Path(file).suffix == '.xlsx':
            if '~$' not in pathlib.Path(file).name:
                filepath = source_folder_name + '/' + file
                extract(vars_file, filepath, target_filename, target_sheet)

    print('Data Extraction Complete')


def single_extract(
        vars_file: str,
        source_filename: str,
        target_filename: str,
        target_sheet: str) -> None:
    """
    Handle extraction in single file mode.
    :param vars_file: str -> absolute path to variable definition csv
    :param source_filename: str -> absolute path of source excel file
    :param target_filename: str -> absolute path of target excel file
    :param target_sheet: str -> name of worksheet in target excel workbook
    """
    if not source_filename:
        print('please select source data file')

    if not target_filename:
        print('please select target file')

    if pathlib.Path(source_filename).suffix == '.xlsx':
        if '~$' not in pathlib.Path(source_filename).name:
            extract(vars_file, source_filename, target_filename, target_sheet)

    print('Data Extraction Complete')
