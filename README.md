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
    python batch.py
    
    # for processing one excel file at a time
    python single.py
    ```

## Note
This is for a very specific application at the moment extracting pre-defined
cells from source and populating pre-defined cells in the target file. 

# TODO
Make a more generic implementation that is able to take any input file(s) and
some sort of cell definitions and perform the extraction.
