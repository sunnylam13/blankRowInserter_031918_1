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

# save the final sheet

# wb.save('multiplicationTable.xlsx')
# logging.debug('Spreadsheet file saved.')