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


## Monday, March 19, 2018 12:48 PM

We have a for loop to go through each row...

Then we have a for loop within that row to go through each column...

## Tuesday, March 20, 2018 8:49 AM

Finding that using the list to do the storage of data is limited as once you start looping you have no idea which data belongs to which row...

	https://stackoverflow.com/questions/1024847/add-new-keys-to-a-dictionary

Instead I'll try using a dict with the key being the row and the values being the column data...

## Tuesday, March 20, 2018 9:27 AM

If using the dict method once you add the blank rows to the `values_afterM_dict` positions, then setting values in the new sheet becomes nothing more than a loop through `value_uptoN_dict` and `values_afterM_dict`...

Since the exact cell to put it in is already stored in the dict...  accounting for blank rows to be inserted (`M`)...

Seems a lot less work then using a list...


## Tuesday, March 20, 2018 9:40 AM

To print key value pairs:


	Your existing code just needs a little tweak. i is the key, so you would just need to use it:

	for i in d:
	    print i, d[i]

	You can also get an iterator that contains both keys and values. 

	In Python 2, d.items() returns a list of (key, value) tuples, while d.iteritems() returns an iterator that provides the same:

	for k, v in d.iteritems():
	    print k, v

	In Python 3, d.items() returns the iterator; to get a list, you need to pass the iterator to list() yourself.

	for k, v in d.items():
	    print(k, v)

https://stackoverflow.com/questions/26660654/how-do-i-print-the-key-value-pairs-of-a-dictionary-in-python/26660785

