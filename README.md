## This is a test project to compose an Excel report (.xlsx) with a lot of pictures using python

### How to install:

#### via Pip:

`$ pip install requirements.txt`

#### via Pipenv:

`$ pipenv install`

### Run tests:

`$ python -m unittest`

### TODO List:

* ~~catch `PermissionError: [Errno 13] Permission denied: 'workbook.xlsx'` when trying to open already opened file~~ Done
* 

### Project Structure

#### `/tests/` folder contains test cases for unittest

#### `/tests/output/` folder to keep automatically created spreadsheets

#### `/tests/input/` folder to keep test data

#### `/scripts/` folder to keep one use scripts for gathering test data