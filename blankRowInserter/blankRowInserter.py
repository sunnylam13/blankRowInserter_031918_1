# -*- coding: utf-8 -*-

#! python3

import openpyxl, sys

from openpyxl.styles import Font
from openpyxl.styles import Color, Fill
from openpyxl.cell import Cell

try:
	from openpyxl.cell import column_index_from_string,get_column_letter
except ImportError:
	from openpyxl.utils import column_index_from_string,get_column_letter

import logging
logging.basicConfig(level=logging.DEBUG, format=" %(asctime)s - %(levelname)s - %(message)s")
# logging.disable(logging.CRITICAL)

# parse the command line

position_row_N = sys.argv[1]
position_row_N = 3 # for testing
logging.debug('The row where we start insertion is:  %s' % (position_row_N))

blank_row_num_M = sys.argv[2]
blank_row_num_M = 2 # for testing
logging.debug('The number of blank rows to insert is:  %s' % (blank_row_num_M))

sheet_to_process = sys.argv[3]
sheet_to_process = "../tests/updatedProduceSales.xlsx" # set it for testing purposes, disable in final product
logging.debug('The spreadsheet to process is:  %s' % (sheet_to_process))

# access the workbook

wb = openpyxl.load_workbook(sheet_to_process) # create the workbook
sheet = wb.active # switch to the active sheet, there should only be one

logging.debug('Testing to see that sheet loaded right, giving a value:  ')
logging.debug(sheet['A2'])
logging.debug(sheet['A2'].value)

# find out the max rows and max columns so we can set upper ends for loops

upper_row_max = sheet.max_row
logging.debug('The maximum number of rows in the sheet is %i' % (upper_row_max))

upper_col_max = sheet.max_column
logging.debug('The maximum number of columns in the sheet is %i' % (upper_col_max))

# copy all of the data up to the N row and its cells, insertion of M happens after N

values_list = [] # to store the values from the spreadsheet

def row_analyzer(min_value,max_value):
	
	for rowValue in range(min_value,max_value): # +1 because we're not starting at 0
		# do not set the upper limit to upper_row_max+1 as normal, instead set it to n + 1 as we will be inserting the gap at that point
		# we still add + 1 because range() stops 1 point before position_row_N normally, we want it to stop exactly at position_row_N
		
		# within this specific rowValue we go through each colValue
		for colValue in range(min_value,upper_col_max+1): # +1 because we're not starting at 0
			# convert the colValue into a letter coordinate
			column_letter = get_column_letter(colValue)
			# combine the column coordinate and row coordinate
			# get the cell coordinate's value and push it into the values_list
			cell_value = sheet[column_letter + str(rowValue)].value
			values_list.append(cell_value)
			logging.debug('The value for %s has been stored in the values_list' % (column_letter + str(rowValue)))
			logging.debug('The value for %s' % (column_letter + str(rowValue)) )
			logging.debug(cell_value)

row_analyzer(1,position_row_N + 1)

# for rowValue in range(1,position_row_N + 1): # +1 because we're not starting at 0
# 	# do not set the upper limit to upper_row_max+1 as normal, instead set it to n + 1 as we will be inserting the gap at that point
# 	# we still add + 1 because range() stops 1 point before position_row_N normally, we want it to stop exactly at position_row_N
	
# 	# within this specific rowValue we go through each colValue
# 	for colValue in range(1,upper_col_max+1): # +1 because we're not starting at 0
# 		# convert the colValue into a letter coordinate
# 		column_letter = get_column_letter(colValue)
# 		# combine the column coordinate and row coordinate
# 		# get the cell coordinate's value and push it into the values_list
# 		cell_value = sheet[column_letter + str(rowValue)].value
# 		values_list.append(cell_value)
# 		logging.debug('The value for %s has been stored in the values_list' % (column_letter + str(rowValue)))
# 		logging.debug('The value for %s' % (column_letter + str(rowValue)) )
# 		logging.debug(cell_value)


logging.debug('The values list is')
logging.debug(values_list)

# save the final sheet

# wb.save('multiplicationTable.xlsx')
# logging.debug('Spreadsheet file saved.')