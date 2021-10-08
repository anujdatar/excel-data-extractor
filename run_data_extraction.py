import os
import pathlib
from participant_class import Participant
from write_to_target_sheet import write_to_target_sheet


def batch_extract(source_folder_name, target_filename):
    if not source_folder_name:
        print('please select source data folder')

    if not target_filename:
        print('please select target file')

    for file in os.listdir(source_folder_name):
        if pathlib.Path(file).suffix == '.xlsx':
            if '~$' not in pathlib.Path(file).name:
                filepath = source_folder_name + '/' + file
                extract(filepath, target_filename)


def single_extract(source_filename, target_filename):
    if not source_filename:
        print('please select source data file')

    if not target_filename:
        print('please select target file')

    if pathlib.Path(source_filename).suffix == '.xlsx':
        if '~$' not in pathlib.Path(source_filename).name:
            extract(source_filename, target_filename)


def extract(source_file, target_file):
    # fetch data
    participant_data = Participant(source_file)
    write_to_target_sheet(target_file, participant_data)