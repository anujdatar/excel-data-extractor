# Excel Data Extractor

Simple batch data extraction tool written in python

## Usage

1. Install python 3.9
2. Navigate to repository folder
3. Install virtual environment
    ```sh
    # on ubuntu you might have to install python3-virtualenv
    python -m venv venv
    ```
4. Activate venv
    ```sh
    # on Windows
    source ./venv/Scripts/activate
   
    # on Linux/Unix
    ./venv/bin/activate
    ```
5. Install `pip` dependencies
    ```sh
    pip install openpyxl
    ```

6. Run program
    ```sh
    # for processing all files in a folder as a batch
    ./batch.py
    
    # for processing one excel file at a time
    ./single.py
    ```
7. Define data to be extracted from source Excel sheets and the target cells
in a CSV **(\*.csv)** file

    | Variable   | Sheet Name | Cell Address | data type | destination column |
    | ---------- | ---------- | ------------ | --------- | ------------------ |
    | variable 1 | Sheet 1    |  a1          | str       | 1                  |
    | variable 2 | Sheet 3    |  g4          | int       | 2                  |
    | variable 3 | Sheet 3    |  b1          | int       | 3                  |

    Note: 
    - variable names are for user reference only, not used by program
    - Sheet Names should be exact, they are case-sensitive
    - Cell address is for the source data, not case-sensitive
    - data type is for reference right now, validation needs to be implemented
    - should be a number, column A=1, B=2, C=3, ... etc.

8. In the GUI,
    - select the variable definition CSV file, source file/folder, target file
    - enter the name of the worksheet in the target Excel file. Should be exact, 
    is case-sensitive

## Note
This is for a very specific application at the moment extracting pre-defined
cells from source and populating pre-defined cells in the target file. 

## TODO
Make a more generic implementation that is able to take any input file(s) and
some sort of cell definitions and perform the extraction.
