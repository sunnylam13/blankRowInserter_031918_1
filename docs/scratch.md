# Scratch Notes and Log

## Monday, March 19, 2018 11:31 AM

The 2 integers, one is *N* and the other *M*...

starting at row *N*, the program inserts *M* blank rows into the sheet...

	python3 blankRowInserter.py 3 2 salesNumbers.xlsx

the program should read the spreadsheet...

then it writes out a new spreadsheet using a `for` loop to copy the first *N* lines...

then it adds *M* blank rows

then it copies the remaining lines of the spreadsheet...

test data

	file:///Users/sunnyair/Dropbox/python_projects/blankRowInserter_031918_1/tests/updatedProduceSales.xlsx

	tests/updatedProduceSales.xlsx

	../tests/updatedProduceSales.xlsx # assuming you're leaving the folder of blankRowInserter.py

for testing

	python3 blankRowInserter.py 3 2 ../tests/updatedProduceSales.xlsx