# Prototype Python script to generate workscope exhibits
#
# NOTES: 
#	1. This script was written to run under Python 2.7.x
#	2. This script relies upon the OpenPyXl, Beautiful Soup (version 4),
#      and wxPython libraries being installed
# 		OpenPyXl is used to read and navigate the input .xlsx workbook
#       BeautifulSoup is used to 'pretty print' (i.e, format) the gerated HTML
#       wxPython is used for the GUI
#	3. To install OpenPyXl, Beautiful Soup (version 4), and wxPython under Python 2.7.x:
#     	<Python_installation_folder>/python.exe -m pip install openpyxl
#     	<Python_installation_folder>/python.exe -m pip install beautifulsoup4
#		<Python_installation_folder>/python.exe -m pip install wxPython
#
# Author: Benjamin Krepp
# Date: 23-27 July 2018
#
# Requirements on the input .xlsx spreadsheet
# ===========================================
#
# This script relies upon the workscope exhibit template .xlsx file having been created
# with 'defined_names' for various locations of interest in it.
# Although 'defined names' are often used to specify ranges of cells, here they are used
# to identify specific cells, whose row and/or column index is obtained for use in the
# generation of the HTML for a workscope exhibit.
#
# 1. The worksheet containing the workscope exhibits MUST be named 'workscope_exhibits'.
#    Other worksheets may be present; their contents are ignored by this script.
#
# 2. These defined names MUST be present in the workbook and defined as follows:
#
#	project_name_cell - cell containing the project's name
#	direct_salary_cell - cell containing the total direct salary and overhead
#   odc_cell - cell containing the total of other direct costs
#   total_cost_cell - cell containing the total cost of the project
#   task_list_top - any cell in line immediately preceeding list of tasks;
#                   only the row index of this cell is used
#   task_list_bottom - any cell in line immediately following list of tasks;
#                      only the row index of this cell is used
#   funding list_top - any cell in line immediately preceeding list of funding
#                      sources; only the row index of this cell is used
#   funding list_bottom - any cell in line immediately following list of funding
#                         sources; only the row index of this cell is used
#   task_number_column - any cell in column containing the task numbers in
#                        the cost table; only the column index of this cell is used
#   task_name_column - any cell in column containint the task name in the 
#                      cost table; only the column index of this cell is used
#   m1_column -    cost table header cell containing the text 'M-1'
#   p5_column -    cost table header cell containing the text 'P-5'
#   p4_column -    cost table header cell containing the text 'P-4' 
#   p3_column -    cost table header cell containing the text 'P-3'
#   p2_column -    cost table header cell containing the text 'P-2'
#   p1_column -    cost table header cell containing the text 'P-1'
#   sp3_column -   cost table header cell containing the text 'SP-3'
#   sp1_column -   cost table header cell containing the text 'SP-1'
#   temp_column -  cost table header cell containing the text 'Temp'
#   total_column - cost table header cell containing the text 'Total'
#	direct_salary_column - cost table header cell containg the string
#                          'Salary' (2nd line of 'Direct Salary')
#	overhead_column - cost table header cell containing the overhead
#                     rate (as a string); used when we want to access the column
#   overhead_cell -   ditto; used when we want to access the cell itself
#   total_cost_column - cost table header cell containing the text
#                       'Cost' (2nd line of 'Total Cost'
#	total_line - any cell in row containing person-hour totals in 
#                the cost table; only the row index of this cell is used
#   odc_travel_line - any cell on line containing Other Direct Costs: travel;
#                     only the row index of this cell is used
#	odc_office_equipment_line - any cell on line containing Other Direct Costs:
#                               general office equipment; only the row index of 
#                               this cell is used
#   odc_dp_equipment_line - any cell on line containing Other Direct Costs:
#						    data processing equipment; only the row index of 
#                           this cell is used
# 	odc_consultants_line - any cell on line containing Other Direct Costs:
#						   consultants; only the row index of this cell is used
#	odc_printing_line -	any cell on line containing Other Direct Costs:
#                       printing; only the row index of this cell is used
# 	odc_other_line - any cell on line containing Ohter Direct Costs: other;
#					 only the row index of this cell is used
#	
# 
# Internals of this Script: Top-level Functions
# =============================================
#
# initialize - Reads a completed .xlsx workscope exhibit template; extracts the row-
#              and colum-indices (and a couple of other things) of interest/use, 
#              which are stored in a dictionary object. This object is subsequently
#              used throughout the rest of this script to extract data from cells
#              of interest in the spreadsheet. It is the most important data structure
#              in this program.
#
# write_exhibit_2 - driver routine for generating Exhibit 2;
#                   calls write_exhibit_2_initial_boilerplate,
#               	write_exhibit_2_body, and write_exhibit_2_final_boilerplate
#
# write_exhibit_2_initial_boilerplate - writes boilerplate HTML at beginning of
#                                       Exhibit 2
#
# write_exhibit_2_final_boilerplate - writes boilerplate HTML at end of Exhibit 2
#
# write_exhibit_2_body - driver routine for producing HTML for the body of Exhibit 2;
#                        calls 	write_direct_salary_div, write_salary_cost_table_div,
#                        write_other_direct_costs_div, write_total_direct_costs_div, 
#                        and write_funding_div
#
# write_direct_salary_div - writes "one-line div" containing total direct salary and
#                           overhead cost
#
# write_other_direct_costs_div - writes "one-line div" containing total of other
#                                direct costs
#
# write_total_direct_costs_div - writes "one-line div" containing total cost
#
# write_funding_div - writes div with list of funding source(s)
#
# write_salary_cost_table_div - writes the div containing the salary cost table;
#                               calls write_task_tr. This is the driver routine
#                               for most of the work done by this program.
#
# write_task_tr - writes row for a given task in the work scope
#
# Internals of this Script: Utility Functions
# ===========================================
#
# get_column_index - return the column index for a defined name assigned to a single cell
#
# get_row_index - return the row index for a defined name assigned to a single cell
#
# get_cell_contents - returns the contents of a cell, given a worksheet name, row index,
#                     and column index
#
# format_person_weeks - formats a value indicating a quantity of person weeks (a float)
#                       as a string one decimal place of precision
#
# format_dollars - formats a value indicating a quantity of dollars (a float) as a
#                  string with zero decimal places of precision (i.e., an integer),
#                  using the ',' symbol as the thousands delimeter
#
#
# 'Ultra-quck Quick Start Guide to Using the OpenPyXl Library
# ============================================================
#
# Open an .xlsx workbook:
#	wb = openpyxl.load_workbook(full_path_to_workbook_file, data_only=True)
#
# Get list of worksheets in workbook:
#	ws_list = wb.sheetnames
#
# Get a named worksheet:
#	ws = wb['workscope_template']
#
# Get list of defined names in workbook:
# 	list_of_dns = wb.defined_names
#
# Get value of a defined name, e.g., 'foobar'
# 	dn_val = wb.defined_names['foobar'].value
#
# Get the worksheet and cell indices for a defined name,
# and get the value of the cell it refers to
#	temp = dn.split('!')
#   # temp[0] is the worksheet name; temp[1] is the cell reference
#   cell = ws[temp[1]]
#	row_index = cell.row
#	column_index = cell.col_idx
# 	cell_value = ws.cell(row_index,column_index).value

import os
import openpyxl
from bs4 import BeautifulSoup
import wx

# Var in which we accumulate all HTML generated.
accumulatedHTML = ''

# *** TBD: Come up with a better name for this fn!
def blah(s):
	global accumulatedHTML
	print s
	accumulatedHTML = accumulatedHTML + s
# end_def blah()


# Return the column index for a defined name assigned to A SINGLE CELL.
# Note: In Excel, the scope of 'defined names' is the entire workBOOK, not a particular workSHEET.
def get_column_index(wb, name):
	ws = wb['workscope_exhibits']
	x = wb.defined_names[name].value
	temp = x.split('!')
	# temp[0] is the worksheet reference, temp[1] is the cell reference
	cell = ws[temp[1]]
	col_ix = cell.col_idx
	return col_ix
# end_def get_column_index()

# Return the row index for a defined name assigned to A SINGLE CELL.
# Note: In Excel, the scope of 'defined names' is the workBOOK, not a particular workSHEET.
def get_row_index(wb, name):
	ws = wb['workscope_exhibits']
	x = wb.defined_names[name].value
	temp = x.split('!')
	# temp[0] is the worksheet reference, temp[1] is the cell reference
	cell = ws[temp[1]]
	row = cell.row
	return row
# end_def get_row_index()

# Return the contents of a cell.
# If OpenPyXl cell accessor raises exception OR value returned by OpenPyXl accessor is None, return the empty string.
def get_cell_contents(ws, row_ix, col_ix):
	try:
		temp = ws.cell(row_ix, col_ix).value
	except:
		temp = ''
	if temp == None:
		retval = ' '
	else:
		retval = temp
	return retval
# end_def get_cell_contents()

# Person weeks are formatted as a floating point number with one digit of precision.
def format_person_weeks(person_weeks):
	retval = "%.1f" % person_weeks
	return retval
# end_def format_person_weeks()

# Dollars are formatted as a floating point number with NO digits of precision,
# i.e., as an integer, with commas as the thousands delimiter.
# Note: This function does NOT prepend a '$' symbol to the string returned.
def format_dollars(dollars):
	retval = '{0:,.0f}'.format(dollars)
	return retval
# end_def format_dollars()


# Open the workbook (.xlsx file) inidicated by the "fullpath" parameter.
# Return a dictionary containing all row and column inidices of interest,
# as well as entries for the workbook itself and the worksheet containing
# the workscope exhibit data.
# 
def initialize(fullpath):
	# retval dictionary
	retval = {}
	# Workbook MUST be opened with data_only parameter set to True.
	# This ensures that we read the computed value in cells containing a formula, not the formula itself.
	wb = openpyxl.load_workbook(fullpath, data_only=True)
	retval['wb'] = wb
	# 
	# N.B. The worksheet containing the workscope exhibits is named 'workscope_exhibits'.
	ws = wb['workscope_exhibits']
	retval['ws'] = ws
	
	# Collect row and column indices for cells of interest
	#
	try:
		retval['project_name_cell_row_ix'] = get_row_index(wb, 'project_name_cell')
		retval['project_name_cell_col_ix'] = get_column_index(wb, 'project_name_cell')
	except:
		retval['project_name_cell_row_ix'] = None
		retval['project_name_cell_col_ix'] = None
	try:
		retval['direct_salary_cell_row_ix'] = get_row_index(wb, 'direct_salary_cell')
		retval['direct_salary_cell_col_ix'] = get_column_index(wb, 'direct_salary_cell')
	except:
		retval['direct_salary_cell_row_ix'] = None
		retval['direct_salary_cell_col_ix'] = None
	try:
		retval['odc_cell_row_ix'] = get_row_index(wb, 'odc_cell')
		retval['odc_cell_col_ix'] = get_column_index(wb, 'odc_cell')
	except:
		retval['odc_cell_row_ix'] = None
		retval['odc_cell_col_ix'] = None
	try:
		retval['total_cost_cell_row_ix'] = get_row_index(wb, 'total_cost_cell')
		retval['total_cost_cell_col_ix'] = get_column_index(wb, 'total_cost_cell')
	except:
		retval['total_cost_cell_row_ix'] = None
		retval['total_cost_cell_col_ix'] = None
	# Overhead rate cell.
	try:
		retval['overhead_cell_row_ix'] = get_row_index(wb, 'overhead_cell')
		retval['overhead_cell_col_ix'] = get_column_index(wb, 'overhead_cell')
	except:
		retval['overhead_cell_row_ix'] = None
		retval['overhead_cell_col_ix'] = None
	#		
	# Collect useful row indices
	#
	try:
		retval['task_list_top_row_ix'] = get_row_index(wb, 'task_list_top')
	except:
		retval['task_list_top_row_ix'] = None
	try:
		retval['task_list_bottom_row_ix'] = get_row_index(wb, 'task_list_bottom')
	except:
		retval['task_list_bottom_row_ix'] = None
	try:
		retval['total_line_row_ix'] = get_row_index(wb, 'total_line')	
	except:
		retval['total_line_row_ix'] = None
	# Rows containing other direct costs
	try:
		retval['odc_travel_line_ix'] =  get_row_index(wb, 'odc_travel_line')
	except:
		retval['odc_travel_line_ix'] = None
	try:
		retval['odc_office_equipment_line_ix'] = get_row_index(wb, 'odc_office_equipment_line')
	except:
		retval['odc_office_equipment_line_ix'] = None
	try:
		retval['odc_dp_equipment_line_ix'] = get_row_index(wb, 'odc_dp_equipment_line')
	except:
		retval['odc_dp_equipment_line_ix'] = None
	try:
		retval['odc_consultants_line_ix'] = get_row_index(wb, 'odc_consultants_line')
	except:
		retval['odc_consultants_line_ix'] = None
	try:
		retval['odc_printing_line_ix'] = get_row_index(wb, 'odc_printing_line')
	except:
		retval['odc_printing_line_ix'] = None
	try:
		retval['odc_other_line_ix'] = get_row_index(wb, 'odc_other_line')	
	except:
		retval['odc_other_line_ix'] =  None
	# Rows containing info on funding source(s)
	try:
		retval['funding_list_top_row_ix'] = get_row_index(wb, 'funding_list_top')
	except:
		retval['funding_list_top_row_ix'] = None
	try:
		retval['funding_list_bottom_row_ix'] = get_row_index(wb, 'funding_list_bottom')
	except:
		retval['funding_list_bottom_row_ix'] = None
	#
	# Collect useful column indices
	#
	try:
		retval['task_number_col_ix'] = get_column_index(wb, 'task_number_column')
	except:
		retval['task_number_col_ix'] = None
	try:
		retval['task_name_col_ix'] = get_column_index(wb, 'task_name_column')
	except:
		retval['task_name_col_ix'] = None		
	try:
		retval['m1_col_ix'] = get_column_index(wb, 'm1_column')
	except:
		retval['m1_col_ix'] = None
	try:
		retval['p5_col_ix'] = get_column_index(wb, 'p5_column')
	except:
		retval['p5_col_ix'] = None
	try:	
		retval['p4_col_ix'] = get_column_index(wb, 'p4_column')
	except:
		retval['p4_col_ix'] = None
	try:
		retval['p3_col_ix'] = get_column_index(wb, 'p3_column')
	except:
		retval['p3_col_ix'] = None
	try:
		retval['p2_col_ix'] = get_column_index(wb, 'p2_column')
	except:
		retval['p2_col_ix'] = None
	try:
		retval['p1_col_ix'] = get_column_index(wb, 'p1_column')
	except:
		retval['p1_col_ix'] = None
	try:
		retval['sp3_col_ix'] = get_column_index(wb, 'sp3_column')
	except:
		retval['sp3_col_ix'] = None
	try:
		retval['sp1_col_ix'] = get_column_index(wb, 'sp1_column')
	except:
		retval['sp1_col_ix'] = None
	try:
		retval['temp_col_ix'] = get_column_index(wb, 'temp_column')
	except:
		retval['temp_col_ix'] = None
	# The following statement refers to the column for total labor cost before overhead
	try:
		retval['total_col_ix'] = get_column_index(wb, 'total_column')
	except:
		retval['total_col_ix'] = None
	try:
		retval['direct_salary_col_ix'] = get_column_index(wb, 'direct_salary_column')
	except:
		retval['direct_salary_col_ix'] = None
	try:
		retval['overhead_col_ix'] = get_column_index(wb, 'overhead_column')
	except:
		retval['total_col_ix'] = None
	try:
		retval['total_cost_col_ix'] = get_column_index(wb, 'total_cost_column')
	except:
		retval['total_cost_col_ix'] = None
	#
	# C'est un petit hacque: The column index for funding source names is the same as that for task names.
	#
	retval['funding_source_name_col_ix'] = retval['task_name_col_ix']
	return retval
# end_def initialize()

# Write initial "boilerplate" HTML for Exhibit 2.
# This includes all content from DOCTYPE, the <html> tag, and everything in the <head>.
def write_exhibit_2_initial_boilerplate():
	s = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
	blah(s)
	s = '<html xmlns="http://www.w3.org/1999/xhtml" lang="en"><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">'
	blah(s)
	s = '<title>CTPS Work Scope Exhibit 2</title>'
	blah(s)
	s = '<link rel="stylesheet" type="text/css" href="./ctps_work_scope_print.css">'
	blah(s)
	s = '</head>'
	blah(s)
# end_def write_exhibit_2_initial_boilerplate()

# This writes the final "boilerplate" HTML for Exhibit 2: the closing </body> and </html> tags.
def write_exhibit_2_final_boilerplate():
	s = '</body>' 
	blah(s)
	s = '</html>'
	blah(s)
# end_def write_exhibit_2_final_boilerplate()


def write_direct_salary_div(xlsInfo):
	s = '<div id="directSalaryDiv" class="barH2">'
	blah(s)
	s = '<h2>Direct Salary and Overhead</h2>'
	blah(s)
	t1 = '<div class="h2AmtDiv">'
	t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['direct_salary_cell_row_ix'], xlsInfo['direct_salary_cell_col_ix']))
	t3 = '</div>'
	s = t1 + t2 + t3
	blah(s)
	s = '</div>'
	blah(s)
# end_def write_direct_salary_div()


######################################################################################################
# Helper function to generate <tr> (and its contents) for one task in the salary cost table.
# This function is called only from write_salary_cost_table_div, which it is LOGICALLY nested within.
# In order to expedite development/prototyping, however, it is currently defined here at scope-0.
# When the tool has become stable, move it within the def of salary_cost_table_div.
#
def write_task_tr(task_num, task_row_ix, xlsInfo, real_cols_info):
	# Open <tr> element
	t1 = '<tr id='
	tr_id = 'taskHeader' + str(task_num)
	t2 = tr_id + '>'
	s = t1 + t2
	blah(s)
	
	# <td> for task number and task name
	# Note: This contains 3 divs organized thus: <div> <div></div> <div></div> </div>
	t1 = '<td headers="taskTblHdr" scope="row" '
	if task_num == 1:
		t2  = 'class="firstTaskTblCell">'
	else:
		t2 = 'class="taskTblCell">'
	# end_if
	s = t1 + t2 
	blah(s)
	# Open outer div
	s = '<div class="taskTblCellDiv">'
	blah(s)
	# First inner div
	t1 = '<div class="taskNumDiv">'
	t2 = get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['task_number_col_ix'])
	t3 = '</div>'
	s = t1 + t2 + t3
	blah(s)
	# Second inner div
	t1 = '<div class="taskNameDiv">'
	t2 = get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['task_name_col_ix'])
	t3 = '</div>'
	s = t1 + t2 + t3
	blah(s)
	# Close outer div, and close <td>
	s = '</div>'
	blah(s)
	s = '</td>'
	blah(s)
	
	# Generate the <td>s for all the salary grades used in this work scope exhibit
	for col_info in real_cols_info:
		t1 = '<td headers="' + tr_id + ' personWeekTblHdr ' + col_info['col_header_id'] + '"'
		t2 = ' class="rightPaddedTblCell">'
		t3 = format_person_weeks(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo[col_info['col_ix']]))
		t4 = '</td>'
		s = t1 + t2 + t3 + t4
		blah(s)
	# end_for
	
	# Generate the <td>s for 'Total [person weeks]', 'Direct Salary', 'Overhead', and 'Total Cost'.
	#
	# Total [person weeks]
	t1 = '<td headers="' + tr_id + ' personWeekTblHdr personWeekTotalTblHdr" class="rightPaddedTblCell">'
	t2 = format_person_weeks(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['total_col_ix']))
	t3 = '</td>'
	s = t1 + t2 + t3
	blah(s)
	#
	# Direct Salary
	t1 = '<td headers="' + tr_id + ' salaryTblHdr" class="rightPaddedTblCell">'
	t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['direct_salary_col_ix']))
	t3 = '</td>'
	s = t1 + t2 + t3
	blah(s)
	#
	# Overhead
	t1 = '<td headers="' + tr_id + ' overheadTblHdr" class="rightPaddedTblCell">'
	t2 = '$' +  format_dollars(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['overhead_col_ix']))
	t3 = '</td>'
	s = t1 + t2 + t3
	blah(s)		
	#
	# Total Cost
	t1 = '<td headers="' + tr_id + ' totalTblHdr" class="rightPaddedTblCell">'
	t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['total_cost_col_ix']))
	t3 = '</td>'
	s = t1 + t2 + t3
	blah(s)		
	
	s = '</tr>'
	blah(s)
# end_def write_task_tr()

##################################################################
# Top-level routine for generating HTML for salary cost table div.
# Calls end_def write_task_tr as a helper function.
def write_salary_cost_table_div(xlsInfo):
	s = '<div class="costTblDiv">'
	blah(s)
	s = '<table id="ex2Tbl" summary="Breakdown of staff time by task in column one, expressed in person weeks for each implicated pay grade in the middle columns,'
	s = s + 'together with resulting total salary and associated overhead costs in the last columns.">'
	blah(s)
	
	# The table header (<thead>) element and its contents
	#
	s = '<thead>'
	blah(s)
	
	# <thead> contents
	# Most of this is invariant bolierplate. The exceptions are the number of "real" columns and the overhead rate.
	
	# First row of <thead> contents
	# 
	s = '<tr>'
	blah(s)
	s = '<th id="taskTblHdr" class="colTblHdr" rowspan="2" scope="col"><br>Task</th>'
	blah(s)
	# 
	# Get actual number of columns to use for "colspan".
	# Determine which columns contain non-zero data: it's sufficent to check the total row for this.
	# Accumulate the result in real_col_ixs, and then use real_col_ixs to create real_cols_info, 
	# which is re-used when generating <tr>s for individual tasks.
	#
	all_cols = ['m1_col_ix', 'p5_col_ix', 'p4_col_ix', 'p3_col_ix', 'p2_col_ix', 'p1_col_ix', 'sp3_col_ix', 'sp1_col_ix', 'temp_col_ix']
	real_col_ixs = []
	for col in all_cols:
		val = get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo[col])
		if val != 0:
			real_col_ixs.append(col)
		# end_if
	# end_for
	
	# Number of columns containing person-week data in output table is equal to the length of "real_col_ixs" + 1 (for the "Total" column.)
	n_real_cols = len(real_col_ixs) + 1
	
	# Although not needed until later, create and populate real_cols_info now.
	real_cols_info = []
	info = {}
	t1, t2 = '', ''
	for col_ix in real_col_ixs:
		info = {}
		info['col_ix'] = col_ix
		# *** Next line is a temp hack!
		# *** TBD: Need 'named range' for row containing job classification abbreviations.
		t1 = get_cell_contents(xlsInfo['ws'], xlsInfo['task_list_top_row_ix']-1, xlsInfo[col_ix])
		info['col_header_with_dash'] = t1
		t2 = t1.replace('-',' ')
		info['col_header_wo_dash'] = t2
		info['col_header_id'] = (t2.replace(' ','')).lower()
		real_cols_info.append(info)
	# end_for
		
	t1 = '<th id="personWeekTblHdr" class="colTblHdr" colspan="'
	t2 = n_real_cols
	t3 = '" abbr="Person Weeks" scope="colgroup">Person-Weeks</th>'
	s = t1 + str(t2) + t3
	blah(s)
	s = '<th id="salaryTblHdr" class="colTblHdr" rowspan="2" scope="col" abbr="Direct Salary">Direct<br>Salary</th>'
	blah(s)
	t1 = '<th id="overheadTblHdr" class="colTblHdr" rowspan="2" scope="col" abbr="Overhead">Overhead<br>'
	t2 = get_cell_contents(xlsInfo['ws'], xlsInfo['overhead_cell_row_ix'], xlsInfo['overhead_cell_col_ix'])
	t2 = t2.replace('@ ', '')
	t3 = '</th>'
	s = t1 + t2 + t3 
	blah(s)
	s = '<th id="totalTblHdr" class="colTblHdr" rowspan="2" scope="col" abbr="Total Cost">Total<br>Cost</th>'
	blah(s)
	s = '</tr>'
	blah(s)
	
	# Second row of <thead> contents
	#
	s = '<tr>'
	blah(s)
	# Column headers for all columns for job classifications used in this work scope
	#
	for col_info in real_cols_info:
		t1 = '<th id="'
		t2 = col_info['col_header_id']
		t3 = '" class="personWKTblHdr" scope="col" abbr="'
		t4 = col_info['col_header_wo_dash']
		t5 = '">'
		t6 = col_info['col_header_with_dash']
		t7 = '</th>'
		s = t1 + t2 + t3 + t4 + t5 + t6 + t7
		blah(s)
	# end_for
	# Second: column header for Total column
	s = '<th id="personWeekTotalTblHdr" scope="col">Total</th>'
	blah(s)
	s = '</tr>'
	blah(s)
	
	# Close <thead> 
	s = '</thead>'
	blah(s)	
	
	# The table body <tbody> element and its contents.
	#
	s = '<tbody>'
	blah(s)
	
	# <tbody> contents.
	#
	# Write <tr>s for each task in the task list.
	i = 0
	for task_row_ix in range(xlsInfo['task_list_top_row_ix']+1,xlsInfo['task_list_bottom_row_ix']):
		i = i + 1
		write_task_tr(i, task_row_ix, xlsInfo, real_cols_info)
	# end_for
	
	# The 'Total' row
	#
	s = '<tr>'
	blah(s)
	s = '<td headers="taskTblHdr" id="totalRowTblHdr" class="taskTblCell" scope="row" abbr="Total All Tasks">'
	blah(s)
	s = '<div class="taskTblCellDiv">'
	blah(s)
	# Total row, task number column (empty)
	s = '<div class="taskNumDiv"> </div>'
	blah(s)
	# Total row, "task name" colum - which contains the pseudo task name 'Total'
	s = '<div class="taskNameDiv">Total</div>'
	blah(s)
	s = '</div>'
	blah(s)
	s = '</td>'
	blah(s)
	
	# Total row: columns for salary grades used in this workscope
	for col_info in real_cols_info:
		t1 = '<td headers="totalRowTblHdr personWeekTblHdr ' + col_info['col_header_id'] + '" class="totalRowTblCell">'
		t2 = format_person_weeks(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo[col_info['col_ix']]))
		t3 = '</td>'
		s = t1 + t2 + t3
		blah(s)
	# end_for
	
	# Total row: Total [person weeks] column
	t1 = '<td id="personWeeksTotalRowTblCell" headers="totalRowTblHdr personWeekTblHdr personWeekTotalTblHdr" class="totalRowTblCell">'
	t2 = format_person_weeks(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo['total_col_ix']))
	t3 = '</td>'
	s = t1 + t2 + t3
	blah(s)
	# Total row, direct salary column
	t1 = '<td id="directSalaryTotalRowTblCell" headers="totalRowTblHdr salaryTblHdr" class="totalRowTblCell">'
	t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo['direct_salary_col_ix']))
	t3 = '</td>'
	s = t1 + t2 + t3
	blah(s)
	# Total row, overhead column
	t1 = '<td id="overheadTotalRowTblCell" headers="totalRowTblHdr overheadTblHdr" class="totalRowTblCell">'
	t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo['overhead_col_ix']))
	t3 = '</td>'
	s = t1 + t2 + t3
	blah(s)
	# Total row, total cost column
	t1 = '<td id="totalTotalRowTblCell" headers="totalRowTblHdr totalTblHdr" class="totalRowTblCell">'
	t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo['total_cost_col_ix']))
	t3 = '</td>'
	s = t1 + t2 + t3
	blah(s)
	# Close <tr> for Total row
	s = '</tr>'
	blah(s)
	
	# Close <tbody>, <table>, and <div>
	s = '</tbody>'
	blah(s)
	s = '</table>'
	blah(s)
	s = '</div>'
	blah(s)
# end_def write_salary_cost_table_div()


def write_other_direct_costs_div(xlsInfo):
	s = '<div id="otherDirectDiv" class="barH2">'
	blah(s)
	s = '<h2>Other Direct Costs</h2>'
	blah(s)
	t1 = '<div class="h2AmtDiv">'
	odc_total = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_cell_row_ix'], xlsInfo['odc_cell_col_ix'])
	t2 = '$' + format_dollars(odc_total)
	t3 = '</div>'
	s = t1 + t2 + t3
	blah(s)
	s = '</div>'
	blah(s)
	# Write the divs for the specific other direct costs and a wrapper div around all of them (even if there are none.)
	#
	# <div> for wrapper
	s = '<div class="costTblDiv">'
	blah(s)
	
	# Utiltiy function to write HTML for one kind of 'other direct cost.'
	def write_odc(name, cost):
		s = '<div class="otherExpDiv">'
		blah(s)
		s = '<div class="otherExpDescDiv">' + name + '</div>'
		blah(s)
		s = '<div class="otherExpAmtDiv">' + '$' + format_dollars(cost) + '</div>'
		blah(s)
		s = '</div>'
		blah(s)
	# end_def write_odc()
	
	# Travel
	travel = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_travel_line_ix'], xlsInfo['total_cost_col_ix'])
	if travel != 0:
		write_odc('Travel', travel)
	
	# General office equipment
	general_office_equipment = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_office_equipment_line_ix'], xlsInfo['total_cost_col_ix'])
	if general_office_equipment != 0:
		write_odc('General Office Equipment', general_office_equipment)
	
	# Data processing equipment
	dp_equipment = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_dp_equipment_line_ix'], xlsInfo['total_cost_col_ix'])
	if dp_equipment != 0:
		write_odc('Data Processing Equipent', dp_equipment)
	
	# Consultant(s)
	consultants = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_consultants_line_ix'], xlsInfo['total_cost_col_ix'])
	if consultants != 0:
		write_odc('Consultants', consultants)
	
	# Printing
	printing = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_printing_line_ix'], xlsInfo['total_cost_col_ix'])
	if printing != 0:
		write_odc('Printing', printing)
	
	# Other 
	other = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_other_line_ix'], xlsInfo['total_cost_col_ix'])
	if other != 0:
		desc = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_other_line_ix'], xlsInfo['task_name_col_ix'])
		write_odc(desc, other)
	
	# </div> for wrapper
	s = '</div>'
	blah(s)
# end_def write_other_direct_costs_div()

def write_total_direct_costs_div(xlsInfo):
	s = '<div id="totalDirectDiv" class="barH2">'
	blah(s)
	s = '<h2>TOTAL COST</h2>'
	blah(s)
	t1 = '<div class="h2AmtDiv">'
	t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['total_cost_cell_row_ix'], xlsInfo['total_cost_cell_col_ix']))
	t3 = '</div>'
	s = t1 + t2 + t3
	blah(s)
	s = '</div>'
	blah(s)
# end_def write_total_direct_costs_div()

def write_funding_div(xlsInfo):
	s = '<div id="fundingDiv">'
	blah(s)
	s = '<div id="fundingHdrDiv">'
	blah(s)
	s = 'Funding'
	blah(s)
	s = '</div>'
	blah(s)
	s =	'<div id="fundingListDiv">'
	blah(s)
	#
	kount = 0
	for fs_row in range(xlsInfo['funding_list_top_row_ix']+1,xlsInfo['funding_list_bottom_row_ix']):
		kount = kount + 1
		# Emit <br> before funding source name except for first funding source.
		s = ''
		if kount != 1:
			s = s + '<br>'
		# end_if
		s = s + get_cell_contents(xlsInfo['ws'], fs_row, xlsInfo['task_name_col_ix'])
		blah(s)
	# end_for		
	s = '</div>'
	blah(s)
	s = '</div>'
	blah(s)
# end_def write_funding_div()

# This writes the HTML for the entire <body> of Exhibit 2, including:
#	the opening <body> tag
#	initial content (e.g., project name, etc.)
# 	the div for the "Direct Salary and Overhead" line
#	the div for the salary cost table
#	the div forthe "Other Direct Costs" line
#	the div for funding source(s)
def write_exhibit_2_body(xlsInfo):
	s = '<body style="text-align:center;margin:0pt;padding:0pt;">'
	blah(s)
	s = '<div id="exhibit2">'
	blah(s)
	s = '<div class="exhibitPageLayoutDiv1"><div class="exhibitPageLayoutDiv2">'
	blah(s)
	s = '<h1>'
	blah(s)
	s = 'Exhibit 2<br>'
	blah(s)
	s = 'ESTIMATED COST<br>'
	blah(s)
	s = str(get_cell_contents(xlsInfo['ws'], xlsInfo['project_name_cell_row_ix'], xlsInfo['project_name_cell_col_ix']))
	s = s + '<br>'
	blah(s)
	s = '</h1>'
	blah(s)
	#
	write_direct_salary_div(xlsInfo)
	write_salary_cost_table_div(xlsInfo)
	write_other_direct_costs_div(xlsInfo)
	write_total_direct_costs_div(xlsInfo)
	write_funding_div(xlsInfo)
# end_def write_exhibit_2_body()

def write_exhibit_2(xlsInfo):
	write_exhibit_2_initial_boilerplate()
	write_exhibit_2_body(xlsInfo)
	write_exhibit_2_final_boilerplate()
# end_def write_exhibit_2()

# Main driver routine.
# Currently only generates HTML for Exhibit 2.
# TBD:
#	1. Generate HTML for Exhibit 1
#	2. Generate PDF for Exhibit 1
#	3. Generate PDF for Exhibit 2
def main():
	app = wx.App()
	frame = wx.Frame(None, -1, 'win.py')
	# In previous versions of wxPython, the following line would have been written: "frame.SetDimensions(0,0,200,500)
	frame.DoSetSize(0,0,200,500, wx.SIZE_AUTO)
	# Create open file dialog
	openFileDialog = wx.FileDialog(frame, "Select workscope exhibit spreadsheet", "", "", 
										  "Excel files (*.xlsx)|*.xlsx", 
										   wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
	openFileDialog.ShowModal()
	fullpath = openFileDialog.GetPath()
	openFileDialog.Destroy()
	
	t1 = os.path.split(fullpath)
	in_dir = t1[0]
	in_fn = t1[1]
	in_fn_wo_suffix = os.path.splitext(in_fn)[0]
	
	ex_1_out_html_fn = in_dir + '\\' + in_fn_wo_suffix + '_Exhibit_1.html'
	ex_1_out_pdf_fn =  in_dir + '\\' + in_fn_wo_suffix + '_Exhibit_1.pdf'
	ex_2_out_html_fn = in_dir + '\\' + in_fn_wo_suffix + '_Exhibit_2.html'
	ex_2_out_pdf_fn =  in_dir + '\\' + in_fn_wo_suffix + '_Exhibit_2.pdf'
	
	# Collect 'navigation' information from input .xlsx file
	xlsInfo = initialize(fullpath)
	
	# TBD: Generate exhibit 1 HTML
	pass
	
	# TBD: Generate exhbit 1 PDF
	pass
	
	# Generate Exhibit 2 HTML
	write_exhibit_2(xlsInfo)
	soup = BeautifulSoup(accumulatedHTML, 'html.parser')
	ex_2_html = soup.prettify() + '\n'
	o = open(ex_2_out_html_fn, 'w')
	# NOTE: We need to encode the output as UTF-8 because it may contain non-ASCII characters,
	# e.g., the  "section" symbol used to identify funding sources such as <section>5303 ..
	o.write(ex_2_html.encode("UTF-8"))
	o.close()
	
	# TBD: Generate exhibit 2 PDF
	pass

# end_def main()
