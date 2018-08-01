# Prototype Python script to generate workscope exhibits
#
# NOTES: 
#   1. This script was written to run under Python 2.7.x
#   2. This script relies upon the OpenPyXl, Beautiful Soup (version 4),
#      and wxPython libraries being installed
#       OpenPyXl is used to read and navigate the input .xlsx workbook
#       BeautifulSoup is used to 'pretty print' (i.e, format) the gerated HTML
#       wxPython is used for the GUI
#   3. To install OpenPyXl, Beautiful Soup (version 4), and wxPython under Python 2.7.x:
#       <Python_installation_folder>/python.exe -m pip install openpyxl
#       <Python_installation_folder>/python.exe -m pip install beautifulsoup4
#       <Python_installation_folder>/python.exe -m pip install wxPython
#
# Author: Benjmin Krepp
# Date: 23-31 July 2018
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
#   project_name_cell - cell containing the project's name
#   direct_salary_cell - cell containing the total direct salary and overhead
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
#   direct_salary_column - cost table header cell containg the string
#                          'Salary' (2nd line of 'Direct Salary')
#   overhead_column - cost table header cell containing the overhead
#                     rate (as a string); used when we want to access the column
#   overhead_cell -   ditto; used when we want to access the cell itself
#   total_cost_column - cost table header cell containing the text
#                       'Cost' (2nd line of 'Total Cost'
#   total_line - any cell in row containing person-hour totals in 
#                the cost table; only the row index of this cell is used
#   odc_travel_line - any cell on line containing Other Direct Costs: travel;
#                     only the row index of this cell is used
#   odc_office_equipment_line - any cell on line containing Other Direct Costs:
#                               general office equipment; only the row index of 
#                               this cell is used
#   odc_dp_equipment_line - any cell on line containing Other Direct Costs:
#                           data processing equipment; only the row index of 
#                           this cell is used
#   odc_consultants_line - any cell on line containing Other Direct Costs:
#                          consultants; only the row index of this cell is used
#   odc_printing_line - any cell on line containing Other Direct Costs:
#                       printing; only the row index of this cell is used
#   odc_other_line - any cell on line containing Ohter Direct Costs: other;
#                    only the row index of this cell is used
#   
# 
# Internals of this Script: Top-level Functions
# =============================================
#
# main - main driver routine for this program
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
#                   write_exhibit_2_body, and write_exhibit_2_final_boilerplate
#
# write_exhibit_2_initial_boilerplate - writes boilerplate HTML at beginning of
#                                       Exhibit 2
#
# write_exhibit_2_final_boilerplate - writes boilerplate HTML at end of Exhibit 2
#
# write_exhibit_2_body - driver routine for producing HTML for the body of Exhibit 2;
#                        calls  write_ex2_direct_salary_div, write_ex2_salary_cost_table_div,
#                        write_ex2_other_direct_costs_div, write_ex2_total_direct_costs_div, 
#                        and write_ex2_funding_div
#
# write_ex2_direct_salary_div - writes "one-line div" containing total direct salary and
#                           overhead cost
#
# write_ex2_other_direct_costs_div - writes "one-line div" containing total of other
#                                direct costs
#
# write_ex2_total_direct_costs_div - writes "one-line div" containing total cost
#
# write_ex2_funding_div - writes div with list of funding source(s)
#
# write_ex2_salary_cost_table_div - writes the div containing the salary cost table;
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
# A Guide for the Perplexed (with apologies to Maimonides):
# An 'Ultra-quck' Quick Start Guide to Using the OpenPyXl Library
# ===============================================================
#
# Open an .xlsx workbook:
#   wb = openpyxl.load_workbook(full_path_to_workbook_file, data_only=True)
#
# Get list of worksheets in workbook:
#   ws_list = wb.sheetnames
#
# Get a named worksheet:
#   ws = wb['workscope_template']
#
# Get list of defined names in workbook:
#   list_of_dns = wb.defined_names
#
# Get value of a defined name, e.g., 'foobar'
#   dn_val = wb.defined_names['foobar'].value
#
# Get the worksheet and cell indices for a defined name,
# and get the value of the cell it refers to
#   temp = dn.split('!')
#   # temp[0] is the worksheet name; temp[1] is the cell reference
#   cell = ws[temp[1]]
#   row_index = cell.row
#   column_index = cell.col_idx
#   cell_value = ws.cell(row_index,column_index).value
#
###############################################################################

import os
import sys
import openpyxl
from bs4 import BeautifulSoup
import wx, wx.html

# Gross global var in which we accumulate all HTML generated.
accumulatedHTML = ''

# Append string to gross globabl variable accumulatedHTML 
def appendHTML(s):
    global accumulatedHTML
    # print s
    accumulatedHTML += s
# end_def appendHTML()

######################################################################################
# I would really like to manage the collection of HTML using the following function,
# but have decided against this (at least for the time being) in order to make the
# code easier to understand for people who are unfamiliar with closures in general 
# (and closures in Python 2.x in particular) and functional programming.
# If you know what you're doing, modifying the code to use "functional_output_manager" 
# will be straightforward. I leave "functional_output_manager" here as a teaser for
# those who might enjoy the opportunity to work with functional code. 
# -- BK 7/27/2018
def functional_output_manager():
    my_vars = {}
    my_vars['accumulatedHTML'] = ''
    def append(s):
        my_vars['accumulatedHTML'] += s
    # end_def
    def clear():
        my_vars['accumulatedHTML'] = ''
    # end_def
    def get():
        return my_vars['accumulatedHTML']
    # end_def
    retval = {}
    retval['append'] = append
    retval['clear'] = clear
    retval['get'] = get
    return retval
# end_def output_manager()



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


# The following routine is under development
def write_ex1_schedule_table(xlsInfo):
    s = '<table id="ex1Tbl"'
    s += 'summary="Breakdown of schedule by tasks in column one and calendar time ranges and deliverable dates in column two.">'
    appendHTML(s)
    
    s = '<thead>'
    # First row of column header, first column: 'Task'
    appendHTML(s)
    s = '<tr>'
    appendHTML(s)
    s = '<th id="ex1taskTblHdr" class="colTblHdr" rowspan="2"><br>Task</th>'
    appendHTML(s)
    
    # First row of column header, second column: time unit used in table: 'Day' | 'Week' | 'Month'
    #
    # Values of the 2 following vars are placeholders!
    colspan = 23 
    time_unit  = 'Week'
    t1 = '<th id="ex1weekTblHeader" class="colTblHdr"'
    t2 = 'colspan="' + colspan + '">' + time_unit + '</th>'
    s = t1 + t2 
    appendHTML(s)
    s = '</tr>'
    appendHTML(s)
    
    # Second row of column header: numbers of individual time units in schedule
    s = '<tr>'
    appendHTML(s)
    
    # *** TBD: Here generate the <th>s for the second row of column headers
    
    
    s = '</tr>'
    appendHTML(s)
    # Close table header
	s = '</thead>'
    appendHTML(s)
    
    # Call subordinate routine to do the heavy lifting: generate Exhibit 2 table body
    write_ex1_schedule_table_body(xlsInfo)
    
# end_def write_ex1_schedule_table()


# The following routine is under development
def write_ex1_milestone_div(xlsInfo):
    s = '<div id="milestoneDiv">'
    appendHTML(s)
    s = '<div id="milestoneHdrDiv">'
    appendHTML(s)
    s = 'Products/Milestones'
    appendHTML(s)
    s = '</div>'
    appendHTML(s)
    s = '<div id="milestoneListDiv">'
    appendHMTL(s)
    # The general form of the 'list' (but it's not an HTML <list>) of deliverables is:
    #   <span class="label"> LETTER_CODE_FOR_DELIVERABLE </span> NAME_OF_DELIVERABLE <br>
    # Example:
    #   <span class="label"> A: </span> Memo to MPO with initial findings <br>
    
    s = '</div>'
    appendHTML(s)
    s = '</div>'
    appendHTML(s)
# end_def write_ex1_milestone_div()


def write_exhibit_1_body(xlsInfo):
    pass
    s = '<body style="text-align:center;padding:0pt;margin:0pt;">'
    appendHTML(s)
    s = '<div id="exhibit1">'
    appendHTML(s)
    s = 'class="exhibitPageLayoutDiv1"><div class="exhibitPageLayoutDiv2">'
    appendHTML(s)
    s = '<h1>'
    appendHTML(s)
    s = 'Exhibit 1<br>'
    appendHTML(s)
    s = 'ESTIMATED SCHEDULE<br>'
    appendHTML(s)
    # Project name
    s = str(get_cell_contents(xlsInfo['ws'], xlsInfo['project_name_cell_row_ix'], xlsInfo['project_name_cell_col_ix']))
    s = s + '<br>'
    appendHTML(s)
    s = '</h1>'
    appendHTML(s)
    #
    write_ex1_schedule_table(xlsInfo)
    write_ex1_milestone_div(xlsInfo)
# end_def 

# TBD: Combine this and write_exhibit_2_body into a single, parameterized,  routine.
def write_exhibit_1_initial_boilerplate():
     s = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
    appendHTML(s)
    s = '<html xmlns="http://www.w3.org/1999/xhtml" lang="en"><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">'
    appendHTML(s)
    s = '<title>CTPS Work Scope Exhibit 1</title>'
    appendHTML(s)
    s = '<link rel="stylesheet" type="text/css" href="./ctps_work_scope_print.css">'
    appendHTML(s)
    s = '</head>'
    appendHTML(s)
# end_def write_exhibit_1_initial_boilerplate()

# Shares 100% code with write_exhibit_1_final_boilerplate. 
# TBD: Combine these two routines.
# Write the final "boilerplate" HTML for Exhibit 1: the closing </body> and </html> tags.
def write_exhibit_1_final_boilerplate():
    s = '</body>' 
    appendHTML(s)
    s = '</html>'
    appendHTML(s)
# end_def write_exhibit_1_final_boilerplate()


def write_exhibit_1(xlsInfo):
    write_exhibit_1_initial_boilerplate()
    write_exhibit_1_body(xlsInfo)
    write_exhibit_1_final_boilerplate()
# end_def write_exhibit_1()

# Write initial "boilerplate" HTML for Exhibit 2.
# This includes all content from DOCTYPE, the <html> tag, and everything in the <head>.
def write_exhibit_2_initial_boilerplate():
    s = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
    appendHTML(s)
    s = '<html xmlns="http://www.w3.org/1999/xhtml" lang="en"><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">'
    appendHTML(s)
    s = '<title>CTPS Work Scope Exhibit 2</title>'
    appendHTML(s)
    s = '<link rel="stylesheet" type="text/css" href="./ctps_work_scope_print.css">'
    appendHTML(s)
    s = '</head>'
    appendHTML(s)
# end_def write_exhibit_2_initial_boilerplate()

# This writes the final "boilerplate" HTML for Exhibit 2: the closing </body> and </html> tags.
def write_exhibit_2_final_boilerplate():
    s = '</body>' 
    appendHTML(s)
    s = '</html>'
    appendHTML(s)
# end_def write_exhibit_2_final_boilerplate()

def write_ex2_direct_salary_div(xlsInfo):
    s = '<div id="directSalaryDiv" class="barH2">'
    appendHTML(s)
    s = '<h2>Direct Salary and Overhead</h2>'
    appendHTML(s)
    t1 = '<div class="h2AmtDiv">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['direct_salary_cell_row_ix'], xlsInfo['direct_salary_cell_col_ix']))
    t3 = '</div>'
    s = t1 + t2 + t3
    appendHTML(s)
    s = '</div>'
    appendHTML(s)
# end_def write_ex2_direct_salary_div()

######################################################################################################
# Helper function to generate <tr> (and its contents) for one task in the salary cost table.
# This function is called only from write_ex2_salary_cost_table_div, which it is LOGICALLY nested within.
# In order to expedite development/prototyping, however, it is currently defined here at scope-0.
# When the tool has become stable, move it within the def of salary_cost_table_div.
#
def write_task_tr(task_num, task_row_ix, xlsInfo, real_cols_info):
    # Open <tr> element
    t1 = '<tr id='
    tr_id = 'taskHeader' + str(task_num)
    t2 = tr_id + '>'
    s = t1 + t2
    appendHTML(s)
    
    # <td> for task number and task name
    # Note: This contains 3 divs organized thus: <div> <div></div> <div></div> </div>
    t1 = '<td headers="taskTblHdr" scope="row" '
    if task_num == 1:
        t2  = 'class="firstTaskTblCell">'
    else:
        t2 = 'class="taskTblCell">'
    # end_if
    s = t1 + t2 
    appendHTML(s)
    # Open outer div
    s = '<div class="taskTblCellDiv">'
    appendHTML(s)
    # First inner div
    t1 = '<div class="taskNumDiv">'
    t2 = get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['task_number_col_ix'])
    t3 = '</div>'
    s = t1 + t2 + t3
    appendHTML(s)
    # Second inner div
    t1 = '<div class="taskNameDiv">'
    t2 = get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['task_name_col_ix'])
    t3 = '</div>'
    s = t1 + t2 + t3
    appendHTML(s)
    # Close outer div, and close <td>
    s = '</div>'
    appendHTML(s)
    s = '</td>'
    appendHTML(s)
    
    # Generate the <td>s for all the salary grades used in this work scope exhibit
    for col_info in real_cols_info:
        t1 = '<td headers="' + tr_id + ' personWeekTblHdr ' + col_info['col_header_id'] + '"'
        t2 = ' class="rightPaddedTblCell">'
        t3 = format_person_weeks(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo[col_info['col_ix']]))
        t4 = '</td>'
        s = t1 + t2 + t3 + t4
        appendHTML(s)
    # end_for
    
    # Generate the <td>s for 'Total [person weeks]', 'Direct Salary', 'Overhead', and 'Total Cost'.
    #
    # Total [person weeks]
    t1 = '<td headers="' + tr_id + ' personWeekTblHdr personWeekTotalTblHdr" class="rightPaddedTblCell">'
    t2 = format_person_weeks(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['total_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    appendHTML(s)
    #
    # Direct Salary
    t1 = '<td headers="' + tr_id + ' salaryTblHdr" class="rightPaddedTblCell">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['direct_salary_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    appendHTML(s)
    #
    # Overhead
    t1 = '<td headers="' + tr_id + ' overheadTblHdr" class="rightPaddedTblCell">'
    t2 = '$' +  format_dollars(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['overhead_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    appendHTML(s)       
    #
    # Total Cost
    t1 = '<td headers="' + tr_id + ' totalTblHdr" class="rightPaddedTblCell">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['total_cost_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    appendHTML(s)       
    
    s = '</tr>'
    appendHTML(s)
# end_def write_task_tr()

############################################################################
# Top-level routine for generating HTML for Exhibit 2 salary cost table div.
# Calls end_def write_ex2_task_tr as a helper function.
def write_ex2_salary_cost_table_div(xlsInfo):
    s = '<div class="costTblDiv">'
    appendHTML(s)
    s = '<table id="ex2Tbl" summary="Breakdown of staff time by task in column one, expressed in person weeks for each implicated pay grade in the middle columns,'
    s = s + 'together with resulting total salary and associated overhead costs in the last columns.">'
    appendHTML(s)
    
    # The table header (<thead>) element and its contents
    #
    s = '<thead>'
    appendHTML(s)
    
    # <thead> contents
    # Most of this is invariant bolierplate. The exceptions are the number of "real" columns and the overhead rate.
    
    # First row of <thead> contents
    # 
    s = '<tr>'
    appendHTML(s)
    s = '<th id="taskTblHdr" class="colTblHdr" rowspan="2" scope="col"><br>Task</th>'
    appendHTML(s)
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
    appendHTML(s)
    s = '<th id="salaryTblHdr" class="colTblHdr" rowspan="2" scope="col" abbr="Direct Salary">Direct<br>Salary</th>'
    appendHTML(s)
    t1 = '<th id="overheadTblHdr" class="colTblHdr" rowspan="2" scope="col" abbr="Overhead">Overhead<br>'
    t2 = get_cell_contents(xlsInfo['ws'], xlsInfo['overhead_cell_row_ix'], xlsInfo['overhead_cell_col_ix'])
    t2 = t2.replace('@ ', '')
    t3 = '</th>'
    s = t1 + t2 + t3 
    appendHTML(s)
    s = '<th id="totalTblHdr" class="colTblHdr" rowspan="2" scope="col" abbr="Total Cost">Total<br>Cost</th>'
    appendHTML(s)
    s = '</tr>'
    appendHTML(s)
    
    # Second row of <thead> contents
    #
    s = '<tr>'
    appendHTML(s)
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
        appendHTML(s)
    # end_for
    # Second: column header for Total column
    s = '<th id="personWeekTotalTblHdr" scope="col">Total</th>'
    appendHTML(s)
    s = '</tr>'
    appendHTML(s)
    
    # Close <thead> 
    s = '</thead>'
    appendHTML(s)   
    
    # The table body <tbody> element and its contents.
    #
    s = '<tbody>'
    appendHTML(s)
    
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
    appendHTML(s)
    s = '<td headers="taskTblHdr" id="totalRowTblHdr" class="taskTblCell" scope="row" abbr="Total All Tasks">'
    appendHTML(s)
    s = '<div class="taskTblCellDiv">'
    appendHTML(s)
    # Total row, task number column (empty)
    s = '<div class="taskNumDiv"> </div>'
    appendHTML(s)
    # Total row, "task name" colum - which contains the pseudo task name 'Total'
    s = '<div class="taskNameDiv">Total</div>'
    appendHTML(s)
    s = '</div>'
    appendHTML(s)
    s = '</td>'
    appendHTML(s)
    
    # Total row: columns for salary grades used in this workscope
    for col_info in real_cols_info:
        t1 = '<td headers="totalRowTblHdr personWeekTblHdr ' + col_info['col_header_id'] + '" class="totalRowTblCell">'
        t2 = format_person_weeks(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo[col_info['col_ix']]))
        t3 = '</td>'
        s = t1 + t2 + t3
        appendHTML(s)
    # end_for
    
    # Total row: Total [person weeks] column
    t1 = '<td id="personWeeksTotalRowTblCell" headers="totalRowTblHdr personWeekTblHdr personWeekTotalTblHdr" class="totalRowTblCell">'
    t2 = format_person_weeks(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo['total_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    appendHTML(s)
    # Total row, direct salary column
    t1 = '<td id="directSalaryTotalRowTblCell" headers="totalRowTblHdr salaryTblHdr" class="totalRowTblCell">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo['direct_salary_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    appendHTML(s)
    # Total row, overhead column
    t1 = '<td id="overheadTotalRowTblCell" headers="totalRowTblHdr overheadTblHdr" class="totalRowTblCell">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo['overhead_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    appendHTML(s)
    # Total row, total cost column
    t1 = '<td id="totalTotalRowTblCell" headers="totalRowTblHdr totalTblHdr" class="totalRowTblCell">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo['total_cost_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    appendHTML(s)
    # Close <tr> for Total row
    s = '</tr>'
    appendHTML(s)
    
    # Close <tbody>, <table>, and <div>
    s = '</tbody>'
    appendHTML(s)
    s = '</table>'
    appendHTML(s)
    s = '</div>'
    appendHTML(s)
# end_def write_ex2_salary_cost_table_div()

def write_ex2_other_direct_costs_div(xlsInfo):
    s = '<div id="otherDirectDiv" class="barH2">'
    appendHTML(s)
    s = '<h2>Other Direct Costs</h2>'
    appendHTML(s)
    t1 = '<div class="h2AmtDiv">'
    odc_total = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_cell_row_ix'], xlsInfo['odc_cell_col_ix'])
    t2 = '$' + format_dollars(odc_total)
    t3 = '</div>'
    s = t1 + t2 + t3
    appendHTML(s)
    s = '</div>'
    appendHTML(s)
    # Write the divs for the specific other direct costs and a wrapper div around all of them (even if there are none.)
    #
    # <div> for wrapper
    s = '<div class="costTblDiv">'
    appendHTML(s)
    
    # Utiltiy function to write HTML for one kind of 'other direct cost.'
    def write_odc(name, cost):
        s = '<div class="otherExpDiv">'
        appendHTML(s)
        s = '<div class="otherExpDescDiv">' + name + '</div>'
        appendHTML(s)
        s = '<div class="otherExpAmtDiv">' + '$' + format_dollars(cost) + '</div>'
        appendHTML(s)
        s = '</div>'
        appendHTML(s)
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
    appendHTML(s)
# end_def write_ex2_other_direct_costs_div()

def write_ex2_total_direct_costs_div(xlsInfo):
    s = '<div id="totalDirectDiv" class="barH2">'
    appendHTML(s)
    s = '<h2>TOTAL COST</h2>'
    appendHTML(s)
    t1 = '<div class="h2AmtDiv">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['total_cost_cell_row_ix'], xlsInfo['total_cost_cell_col_ix']))
    t3 = '</div>'
    s = t1 + t2 + t3
    appendHTML(s)
    s = '</div>'
    appendHTML(s)
# end_def write_ex2_total_direct_costs_div()

def write_ex2_funding_div(xlsInfo):
    s = '<div id="fundingDiv">'
    appendHTML(s)
    s = '<div id="fundingHdrDiv">'
    appendHTML(s)
    s = 'Funding'
    appendHTML(s)
    s = '</div>'
    appendHTML(s)
    s = '<div id="fundingListDiv">'
    appendHTML(s)
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
        appendHTML(s)
    # end_for       
    s = '</div>'
    appendHTML(s)
    s = '</div>'
    appendHTML(s)
# end_def write_ex2_funding_div()

# This writes the HTML for the entire <body> of Exhibit 2, including:
#   the opening <body> tag
#   initial content (e.g., project name, etc.)
#   the div for the "Direct Salary and Overhead" line
#   the div for the salary cost table
#   the div forthe "Other Direct Costs" line
#   the div for funding source(s)
def write_exhibit_2_body(xlsInfo):
    s = '<body style="text-align:center;margin:0pt;padding:0pt;">'
    appendHTML(s)
    s = '<div id="exhibit2">'
    appendHTML(s)
    s = '<div class="exhibitPageLayoutDiv1"><div class="exhibitPageLayoutDiv2">'
    appendHTML(s)
    s = '<h1>'
    appendHTML(s)
    s = 'Exhibit 2<br>'
    appendHTML(s)
    s = 'ESTIMATED COST<br>'
    appendHTML(s)
    # Project name
    s = str(get_cell_contents(xlsInfo['ws'], xlsInfo['project_name_cell_row_ix'], xlsInfo['project_name_cell_col_ix']))
    s = s + '<br>'
    appendHTML(s)
    s = '</h1>'
    appendHTML(s)
    #
    write_ex2_direct_salary_div(xlsInfo)
    write_ex2_salary_cost_table_div(xlsInfo)
    write_ex2_other_direct_costs_div(xlsInfo)
    write_ex2_total_direct_costs_div(xlsInfo)
    write_ex2_funding_div(xlsInfo)
# end_def write_exhibit_2_body()

def write_exhibit_2(xlsInfo):
    write_exhibit_2_initial_boilerplate()
    write_exhibit_2_body(xlsInfo)
    write_exhibit_2_final_boilerplate()
# end_def write_exhibit_2()

# Main driver routine - this function does NOT launch a GUI.
# Currently only generates HTML for Exhibit 2.
# TBD: Generate HTML for Exhibit 1
def main(fullpath):
    global accumulatedHTML # Yeech
    t1 = os.path.split(fullpath)
    in_dir = t1[0]
    in_fn = t1[1]
    in_fn_wo_suffix = os.path.splitext(in_fn)[0]
    ex_1_out_html_fn = in_dir + '\\' + in_fn_wo_suffix + '_Exhibit_1.html'
    ex_2_out_html_fn = in_dir + '\\' + in_fn_wo_suffix + '_Exhibit_2.html'
    
    # Collect 'navigation' information from input .xlsx file
    xlsInfo = initialize(fullpath)
    
    # TBD: Generate exhibit 1 HTML; write_exhibit_1() is currently just a stub.
    accumulatedHTML = ''
    write_exhibit_1(xlsInfo)
    
    # Generate Exhibit 2 HTML
    accumulatedHTML = ''
    write_exhibit_2(xlsInfo)
    soup = BeautifulSoup(accumulatedHTML, 'html.parser')
    ex_2_html = soup.prettify() + '\n'
    o = open(ex_2_out_html_fn, 'w')
    # NOTE: We need to encode the output as UTF-8 because it may contain non-ASCII characters,
    # e.g., the "section" symbol used to identify funding sources such as <section>5303 ..
    o.write(ex_2_html.encode("UTF-8"))
    o.close()   
# end_def main()

# Code for the application's GUI begins here.
#
aboutText = """<p>Help text for this program is TBD.<br>
This program is running on version %(wxpy)s of <b>wxPython</b> and %(python)s of <b>Python</b>.
See <a href="http://wiki.wxpython.org">wxPython Wiki</a></p>""" 

class HtmlWindow(wx.html.HtmlWindow):
    def __init__(self, parent, id, size=(600,400)):
        wx.html.HtmlWindow.__init__(self,parent, id, size=size)
        if "gtk2" in wx.PlatformInfo:
            self.SetStandardFonts()
    # end_def __init__()

    def OnLinkClicked(self, link):
        wx.LaunchDefaultBrowser(link.GetHref())
    # end_def OnLinkClicked()
# end_class HtmlWindow

class AboutBox(wx.Dialog):
    def __init__(self):
        wx.Dialog.__init__(self, None, -1, "About the Workscope Exhibit Tool",
                           style=wx.DEFAULT_DIALOG_STYLE|wx.RESIZE_BORDER|wx.TAB_TRAVERSAL)
        hwin = HtmlWindow(self, -1, size=(400,200))
        vers = {}
        vers["python"] = sys.version.split()[0]
        vers["wxpy"] = wx.VERSION_STRING
        hwin.SetPage(aboutText % vers)
        btn = hwin.FindWindowById(wx.ID_OK)
        irep = hwin.GetInternalRepresentation()
        hwin.SetSize((irep.GetWidth()+25, irep.GetHeight()+10))
        self.SetClientSize(hwin.GetSize())
        self.CentreOnParent(wx.BOTH)
        self.SetFocus()
    # end_def __init__()
# end_class AboutBox

# This is the class for the main GUI itself.
class Frame(wx.Frame):
    xlsxFileName = ''
    def __init__(self, title):
        wx.Frame.__init__(self, None, title=title, pos=(150,150), size=(600,250),
                          style=wx.SYSTEM_MENU | wx.CAPTION | wx.CLOSE_BOX)
        self.Bind(wx.EVT_CLOSE, self.OnClose)

        menuBar = wx.MenuBar()
        menu = wx.Menu()
        m_exit = menu.Append(wx.ID_EXIT, "E&xit\tAlt-X", "Close window and exit program.")
        self.Bind(wx.EVT_MENU, self.OnClose, m_exit)
        menuBar.Append(menu, "&File")
        menu = wx.Menu()
        m_about = menu.Append(wx.ID_ABOUT, "&About", "Information about this program")
        self.Bind(wx.EVT_MENU, self.OnAbout, m_about)
        menuBar.Append(menu, "&Help")
        self.SetMenuBar(menuBar)
        
        self.statusbar = self.CreateStatusBar()

        panel = wx.Panel(self)
        box = wx.BoxSizer(wx.VERTICAL)
        box.AddSpacer(20)
              
        m_select_file = wx.Button(panel, wx.ID_ANY, "Select Excel workbook")
        m_select_file.Bind(wx.EVT_BUTTON, self.OnSelectFile)
        box.Add(m_select_file, 0, wx.CENTER)
        box.AddSpacer(20)
        
        m_generate = wx.Button(panel, wx.ID_ANY, "Generate HTML for Exhibits")
        m_generate.Bind(wx.EVT_BUTTON, self.OnGenerate)
        box.Add(m_generate, 0, wx.CENTER)
 
        # Placeholder for name of selected .xlsx file; it is populated in OnSelectFile(). 
        self.m_text = wx.StaticText(panel, -1, " ")
        self.m_text.SetFont(wx.Font(10, wx.SWISS, wx.NORMAL, wx.NORMAL))
        self.m_text.SetSize(self.m_text.GetBestSize())
        box.Add(self.m_text, 0, wx.ALL, 10)      
        
        panel.SetSizer(box)
        panel.Layout()
    # end_def __init__()
        
    def OnClose(self, event):
        dlg = wx.MessageDialog(self, 
            "Do you really want to close this application?",
            "Confirm Exit", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_OK:
            self.Destroy()
    # end_def OnClose()

    def OnSelectFile(self, event):
        frame = wx.Frame(None, -1, 'win.py')
        frame.SetSize(0,0,200,50)
        openFileDialog = wx.FileDialog(frame, "Select workscope exhibit spreadsheet", "", "", 
                                       "Excel files (*.xlsx)|*.xlsx", 
                                       wx.FD_OPEN | wx.FD_FILE_MUST_EXIST)
        openFileDialog.ShowModal()
        self.xlsxFileName = openFileDialog.GetPath()
        self.m_text.SetLabel("Selected .xlsx file: " + self.xlsxFileName)
        openFileDialog.Destroy()
    # end_def OnSelectFile()
    
    def OnGenerate(self, event):
        dlg = wx.MessageDialog(self, 
            "Do you really want to run the HTML generation tool?",
            "Confirm: OK/Cancel", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_OK:
            main(self.xlsxFileName)
            message = "HTML for workscope exhibits generated."
            caption = "Work Scope Exhibit Tool"
            dlg = wx.MessageDialog(None, message, caption, wx.OK | wx.ICON_INFORMATION)
            dlg.ShowModal()
            dlg.Destroy()
            self.Destroy()
        else:
            self.Destroy()
    # end_def OnGenerate()

    def OnAbout(self, event):
        dlg = AboutBox()
        dlg.ShowModal()
        dlg.Destroy() 
    # end_def OnAbout()
# end_class Frame

# The code for the GUI'd application itself begins here.
#
app = wx.App(redirect=True)   # Error messages go to popup window
top = Frame("Workscope Exhibit Tool")
top.Show()
app.MainLoop()
