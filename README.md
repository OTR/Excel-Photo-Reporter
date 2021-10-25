## This is a test project to compose an Excel report (.xlsx) with a lot of pictures using python

### How to install:

#### via Pip:

`$ pip install requirements.txt`

#### via Pipenv:

`$ pipenv install`

### Run tests:

Run all tests:

`$ python -m unittest`

To run exact TestCase:

`$ python -m unittest tests.test_2nd_create_amazon_books_report.TestCreateAmazonBooksReport
`

### TODO List:

* ~~catch `PermissionError: [Errno 13] Permission denied: 'workbook.xlsx'` when trying to open already opened file~~ Done
* 

### Project Structure

#### `/tests/` folder contains test cases for unittest

#### `/tests/output/` folder to keep automatically created spreadsheets

#### `/tests/input/` folder to keep test data

#### `/scripts/` folder to keep one use scripts for gathering test data

### Test Cases:

#### 1st Test Case:

Create a simple spreadsheet
  
* fill it up with some numbers
  
* save it on a disk

* read it from a disk

* make sure numbers are right

#### 2nd Test Case:

Create a report about existing books on amazon by given keyword (by default keyword is `python`)

1. Gather a dataset about books from amazon using `urllib` and `lxml`

2. Save it on a disk as CSV file (separator is a semicolon)

3. Load it from a disk in memory

4. Create and fill up a spreadsheet with that dataset

5. Widen each row up to 165 pt (in Excel measures) except the 1st row (row with titles) which is 57 pt

6. Widen rows according to the following (in Excel measures):

    * Position:                  5 pt
    * Title:                    36 pt
    * Author(s):                11 pt
    * Price for paperback, $:   11 pt
    * Date of publishing:       11 pt
    * Cover:                    21 pt
    
7. Change the 1st row:
    * Background color to light blue RGB(155, 194, 230)
    * Horizontally align text to center
    * Vertically alighn text to top
    * Make a font style bold
    
8. Change other rows:
    * Horizontally align text to left
    * Vertically align text to top

9. Change background color of the "Cover" column to gray RGB(123, 123, 123)

10. Make a thin black border around each cell

11. Insert an image into each cell in "Cover" column:
    * Make it's padding to 10% of cell width
    * Make an image being auto-resized with a cell
  
12. Make a worksheet protected from modifying
