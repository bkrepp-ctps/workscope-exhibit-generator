# Python module to generate workscope exhibits
#
# NOTES: 
#   1. This module was written to run under Python 2.7.x
#   2. This module relies upon the OpenPyXl and Beautiful Soup (version 4),
#      libraries being installed
#       OpenPyXl is used to read and navigate the input .xlsx workbook
#       BeautifulSoup is used to 'pretty print' (i.e, format) the gerated HTML
#   3. To install OpenPyXl, Beautiful Soup (version 4), and wxPython under Python 2.7.x:
#       <Python_installation_folder>/python.exe -m pip install openpyxl
#       <Python_installation_folder>/python.exe -m pip install beautifulsoup4
#   4. This module relies upon the 'excelFileManager.py' module to read the 
#      input .xlsx file from which the workscope exhibits are generated.
#      PLEASE READ THE DOCUMENTION FOR THE excelFileManager MODULE THOROUGHLY!
#
# This script is a 'port' of a CFML application to Python. The code has been written
# in such a way as to make correlation of a given section of Python code that produces
# an HTML fragment as easy as possible  to correlate with the corresponding segment 
# of CFML code. As there is no functional spec for the original CFML application 
# (the CFML code IS the functional spec for the app, so to speak), this was essential
# in order to ensure functional correctness and debug-ability. As a side-effect of this,
# this code is neither particularly efficient nor particularly idiomatic Python. 
#
# Author: Benjamin Krepp
# Date: 23 July - 8 August 2018
#   
# Internals of this Module: Top-level Functions
# =============================================
#
# main - main driver routine for this program
#
# gen_exhibit_2 - driver routine for generating Exhibit 2;
#                 calls gen_exhibit_2_initial_boilerplate,
#                 gen_exhibit_2_body, and gen_exhibit_2_final_boilerplate
#
# gen_exhibit_2_initial_boilerplate - generates boilerplate HTML at beginning of
#                                     Exhibit 2
#
# gen_exhibit_2_final_boilerplate - generates boilerplate HTML at end of Exhibit 2
#
# gen_exhibit_2_body - driver routine for producing HTML for the body of Exhibit 2;
#                      calls  gen_ex2_direct_salary_div, gen_ex2_salary_cost_table_div,
#                      gen_ex2_other_direct_costs_div, gen_ex2_total_direct_costs_div, 
#                      and gen_ex2_funding_div
#
# gen_ex2_direct_salary_div - generates "one-line div" containing total direct salary and
#                             overhead cost
#
# gen_ex2_other_direct_costs_div - generates "one-line div" containing total of other
#                                  direct costs
#
# gen_ex2_total_direct_costs_div - generates "one-line div" containing total cost
#
# gen_ex2_funding_div - generates div with list of funding source(s)
#
# gen_ex2_salary_cost_table_div - generates the div containing the salary cost table;
#                                 calls gen_task_tr. This is the driver routine
#                                 for most of the work done by this program.
#
# gen_task_tr - generates row for a given task in the work scope
#
# Internals of this Module: Utility Functions
# ===========================================
#
# format_person_weeks - formats a value indicating a quantity of person weeks (a float)
#                       as a string one decimal place of precision
#
# format_dollars - formats a value indicating a quantity of dollars (a float) as a
#                  string with zero decimal places of precision (i.e., an integer),
#                  using the ',' symbol as the thousands delimeter
#
###############################################################################

import os
import sys
import math
import re
import openpyxl
from bs4 import BeautifulSoup
from excelFileManager import initExcelFile, get_column_index, get_row_index, get_cell_contents, get_sched_col_info, get_last_used_sched_column, MAGIC_FILL_STYLE

debug_flags = {}
debug_flags['dump_sched_elements'] = False

# Global pseudo-constants
# *** TBD: Verify that the following should be a float value
WEEKS_PER_MONTH = 52.0/12.0
WEEKS_PER_QUARTER = 52/4
BAR_UNIT_IN_POINTS = 34.6875

# Gross global var in which we accumulate all HTML generated
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



def gen_ex1_task_tr_2nd_td(task_num, task_row_ix, xlsInfo, sched_col_info):
    global debug_flags
    t1 = '<td colspan="' + str(sched_col_info['schedule_columns']) + ' '
    # *** TBD: 'timeUnit1' seems to ALWAYS be incuded as a header. Is this right?
    t2 = 'headers ="row' + str(task_num) + ' timeUnit1" '
    # *** TBD: Why is this class set according to whether or not it's the first task (i.e., row)?
    if task_num == 1:
        t3 = 'class="firstSchedColCell">'
    else:
        t3 = 'class="schedColCell">'
    # end_if
    s = t1 + t2 + t3
    appendHTML(s)
    
    # The guts of 2nd <td> in schedule row.
    # This may contain an arbitrary number of chart 'bars' and an arbitrary number
    # of 'milestones'. Each of these is placed in a <div> of its own,  generated in 
    # ascending chronological order.
    #
    # In order to this, we build (1) a list of 'bars' and (2) a list of 'milestones,'
    # both in ascending chronological (i.e., column) order. We use a common dictionary 
    # data structure for each element of these two lists:
    #     'type'      : 'bar' or 'milestone'
    #     'start'     : start column index
    #     'end'       : end column index   ('start' == 'end' for milestones)
    #     'milestone' : if item is a milestone, the milestone letter, otherwise ''
    # We then merge the two lists, maintaining ascending chronological order.  
    # Having done this, we will then be in a position to generate the relevant HTML.
    
    # Generate the list of 'bars'
    # Prep work: Generate a string of 0's and 1's indicating the cells in the schedule
    # bar chart that have been 'filled in' with the magic fill pattern 'gray125'.
    # Logically, we're creating a bit vector; it's implemented, however as a vector
    # of '0' and '1' characters in a string.
    my_pseudo_bv = ''
    ws = xlsInfo['ws']
    for col in range(xlsInfo['first_schedule_col_ix'],xlsInfo['last_schedule_col_ix']):
        cell = ws.cell(task_row_ix, col)
        fill = cell.fill
        patternType = fill.patternType
        my_pseudo_bv += '1' if patternType == MAGIC_FILL_STYLE else '0'
    # end_for
    
    # The list of 'bars'
    bars = []
    my_re = re.compile('1+')
    my_iter = my_re.finditer(my_pseudo_bv)
    for match in my_iter:
        my_span = match.span()
        # To get the actual column indices of the first and last cell, bias the indices
        # in the bitvector by the index of the first column in the schedule table
        start = my_span[0] + xlsInfo['first_schedule_col_ix']
        end = my_span[1] + xlsInfo['first_schedule_col_ix'] - 1
        # Debug:
        # print "task #" + str(task_num) + " start: " + str(start) + " end: " + str(end)
        temp = {}
        temp['type'] = 'bar'
        temp['start'] = start
        temp['end'] = end
        temp['milestone'] = ''
        bars.append(temp)
    # end_for
    
    # Generate the list of 'milestones'
    milestones = []
    for col in range(xlsInfo['first_schedule_col_ix'],xlsInfo['last_schedule_col_ix']):
        val = get_cell_contents(xlsInfo['ws'], task_row_ix, col)
        if str(val).isupper():
            # Debug:
            # print "milestone : " + str(val) + " start: " + str(col)
            temp = {}
            temp['type'] = 'milestone'
            temp['start'] = col
            temp['end'] = col
            temp['milestone'] = val
            milestones.append(temp)
        # end_if
    # end_for
    
    # Combine 'bars' and 'milestones' into a single list, and sort the result 
    # on (1) the 'start' value and (2) the 'type'.
    # N.B. 'bar' appears before 'milestone' in the sort order for 'type', 
    #      so thereby ensure that the HTML for task bars are generated before
    #      the HTML for any milesones occuring within them.
    big_list = bars + milestones
    big_list_sorted = sorted(big_list, key=lambda x: (x['start'], x['type']))
    # Debug:
    if debug_flags['dump_sched_elements'] == True:
        for thing in big_list_sorted:
            print '*** ' + thing['type'] + ' ' + str(thing['start']) + ' ' + str(thing['end']) + ' ' + thing['milestone']
        # end_if
    # end_for
    

    
    
    # Close seond <td> in row
    s = '</td>'
    appendHTML(s)
# end_def gen_ex1_task_tr_2nd_td()

# The following routine is under development
def gen_ex1_task_tr(task_num, task_row_ix, xlsInfo, sched_col_info):
    s = '<tr>'
    appendHTML(s)
      
    # First <td> in row: task number and task name
    t1 = '<td id="row' + str(task_num) + '" headers="ex1taskTblHdr" '
    if task_num == 1:
        t2 = 'class="firstTaskTblCell">'
    else:
        t2 = 'class="taskTblCell">'
    # end_if
    s = t1 + t2
    appendHTML(s)
    
    t1 = '<div class="taskNumDiv">'
    # *** TBD: Fetch task number from cell in  Excel file rather than using task_num
    t2 = str(task_num) + '.'
    t3 = '</div>'
    s = t1 + t2 + t3
    appendHTML(s)
    
    t1 = '<div class="taskNameDiv">'
    #  *** TBD: This currently gets the task name from its cell in the cost table
    t2 = get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['task_name_col_ix'])
    t3 = '</div>'
    s = t1 + t2 + t3
    appendHTML(s)
    # Close first <td> in row
    s = '</td>'
    appendHTML(s)
    
    # Second <td> in row: schedule bar(s) and deliverable(s), (if any)
    gen_ex1_task_tr_2nd_td(task_num, task_row_ix, xlsInfo, sched_col_info)
    
    # Close <tr>
    s = '</tr>'
    appendHTML(s)
# end_def gen_ex1_task_tr()

def gen_ex1_schedule_table_body(xlsInfo, sched_col_info):
    # Open <tbody>
    s = '<tbody>'
    appendHTML(s)
    # Write the <tr>s in the table body
    i = 0
    for task_row_ix in range(xlsInfo['task_list_top_row_ix']+1,xlsInfo['task_list_bottom_row_ix']):
        i = i + 1
        gen_ex1_task_tr(i, task_row_ix, xlsInfo, sched_col_info)
    # end_for
    # Close <tbody>
    s = '</tbody>'
    appendHTML(s)
    # Close <table>
    s = '</table>'
    appendHTML(s)
# end_def gen_ex1_schedule_table_body()

# Collect and compute information on columnar organization
# of the schedule chart in Exhibit 1
def get_sched_col_info(xlsInfo):
    rv = {}
    
    # ***  How to get 'last_week'? Harvest from .xlsx file????
    # Placeholder, for now.
    rv['last_week'] = 48
    
    # Placeholder, for now.
    # *** TBD: Harvest this value from input .xlsx and store in xlsInfo in initialize()
    rv['sched_unit'] = 'w' 
    
    # Comments from CFML code:
    # if the project will span 12 weeks or less, or 6-12 months, the column width will be such that
	# 12 columns would fit in the space alloted (wide columns), other so that 24 columns would fit (narrow)
    
    if rv['last_week'] <= 13:
        rv['sched_col_width_basis'] = 12
        rv['column_unit'] = 'w'
    elif rv['last_week'] <= 25:
        if rv['schedule_unit'] == 'months':
            rv['schedule_col_width_basis'] = 12
            rv['column_unit'] = 'm'
        else:
            rv['schedule_col_width_basis'] = 24
            rv['column_unit'] = 'w'
        # end_if
    elif rv['last_week'] <= 53:
        # Project is one year or less, but moe than six months
        rv['sched_col_width_basis'] = 12
        rv['column_unit'] = 'm'
    elif rv['last_week'] <= 105:
        # Project is two years or less, but more than one year
        rv['sched_col_width_basis'] = 24
        rv['column_unit'] = 'm'
    elif rv['last_week'] < 157:
        # Project is three years or less, but more than two years
        rv['sched_col_width_basis'] = 12
        rv['column_unit'] = 'q'
    else:
        # Project is more than three years long
        # Note, however, that projects more than four years long just won't fit
        rv['sched_col_width_basis'] = 24
        rv['column_unit'] = 'q'
    # end_if
    
    # Each bar unit is a column, and if the column headings are in months, 
    # then each bar unit is slightly over 4 weeks
    if rv['column_unit'] == 'm':
        rv['weeks_per_bar_unit'] = WEEKS_PER_MONTH
    elif column_unit == 'q':
        rv['weeks_per_bar_unit'] = WEEKS_PER_QUARTER
    else:
        rv['weeks_per_bar_unit'] = 1
    # end_if
    
    # Set 'schedule_columns'
    # CFML statment: schedule_columns = Ceiling(Round(100*(last_week - 1)/weeks_per_bar_unit)/100)
    rv['schedule_columns'] = int(math.ceil(round(100*(rv['last_week'] - 1)/rv['weeks_per_bar_unit'],0)/100))
    
    # Debug:
    print 'schedule_columns = ' + str(rv['schedule_columns'])

    # This is now hardwired
    rv['cell_border_unit'] = 'PixBorder'    
    
    return rv
# end_def get_sched_col_info()

def gen_ex1_schedule_table(xlsInfo):
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
    
    # Prep for generation of 2nd row of table <head> column header
    # and the table <body>
    sched_col_info = get_sched_col_info(xlsInfo)
    
    # First row of column header, second column: time unit used in table, e.g., 'Week'
    #

    t1 = '<th id="ex1weekTblHeader" class="colTblHdr"'
    t2 = 'colspan="' + str(sched_col_info['schedule_columns']) + '">' 
    if sched_col_info['column_unit'] == 'm':
        t3 = 'Month'
    elif sched_col_info['column_unit'] == '1':
        t3 = 'Quarter'
    else:
        t3 = 'Week'
    # end_if
    t4 = '</th>'
    s = t1 + t2 +t3 + t4
    appendHTML(s)
    s = '</tr>'
    appendHTML(s)
    

    # Second row of column header: numbers of individual time units in schedule
    s = '<tr>'
    appendHTML(s)
    # The <th>s for the second row of column headers
    # *** TBD: Chekc that we're using the right value here.
    for i in range(1,sched_col_info['schedule_columns']+1):
        t1 = '<th id='
        t2 = '"timeUnit' + str(i) + '"'
        if sched_col_info['schedule_columns'] == 23:
            t3 = ' class="scheduleColHdr24PixBorder" '
        else:
            # Only alternative is 12, right?
            t3 = ' class="scheduleColHdr12PixBorder" '
        # end_if
        t3 += ' abbr="Schedule range">'
        t4 = str(i) + '</th>'
        s = t1 + t2 + t3 + t4
        appendHTML(s)
    # end_for
    # Close the 2nd row of column headers
    s = '</tr>'
    appendHTML(s)
    
    # Close table header
    s = '</thead>'
    appendHTML(s)
  
    # Call subordinate routine to do the heavy lifting: generate the <table> body for Exhibit 1
    gen_ex1_schedule_table_body(xlsInfo, sched_col_info)
# end_def gen_ex1_schedule_table()


def gen_ex1_milestone_div(xlsInfo):
    s = '<div id="milestoneDiv">'
    appendHTML(s)
    s = '<div id="milestoneHdrDiv">'
    appendHTML(s)
    s = 'Products/Milestones'
    appendHTML(s)
    s = '</div>'
    appendHTML(s)
    s = '<div id="milestoneListDiv">'
    appendHTML(s)
    # The general form of the 'list' (but it's not an HTML <list>) of deliverables is:
    #   <span class="label"> LETTER_CODE_FOR_DELIVERABLE </span> NAME_OF_DELIVERABLE <br>
    # Example:
    #   <span class="label"> A: </span> Memo to MPO with initial findings <br>
    
    # Dumb little predicate to return True if string is empty or only contains blanks, False otherwise.
    def is_empty(s):
        return s.strip() == ''
    # end_def is_empty()
    
    first_milestone_ix = xlsInfo['milestones_list_first_row_ix']
    # Find the last row of the milestones list: crawl down milestone_label_column
    # until the first row containing an 'empty' cell is found. 
    last_milestone_ix = first_milestone_ix + 1
    while is_empty(get_cell_contents(xlsInfo['ws'], last_milestone_ix, xlsInfo['milestone_label_col_ix'])) == False:
        last_milestone_ix += 1
    # end_while
    
    for milestone_ix in range(first_milestone_ix, last_milestone_ix):
        t1 = '<span class="label">'
        t2 = get_cell_contents(xlsInfo['ws'], milestone_ix, xlsInfo['milestone_label_col_ix'])
        t3 = '</span>'
        t4 = get_cell_contents(xlsInfo['ws'], milestone_ix, xlsInfo['milestone_name_col_ix'])
        t5 = '<br>'
        s = t1 + t2 + t3 + t4 + t5
        appendHTML(s)
    # end_for
    
    s = '</div>'
    appendHTML(s)
    s = '</div>'
    appendHTML(s)
# end_def gen_ex1_milestone_div()


def gen_exhibit_1_body(xlsInfo):
    pass
    s = '<body style="text-align:center;padding:0pt;margin:0pt;">'
    appendHTML(s)
    s = '<div id="exhibit1">'
    appendHTML(s)
    s = '<div class="exhibitPageLayoutDiv1"><div class="exhibitPageLayoutDiv2">'
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
    gen_ex1_schedule_table(xlsInfo)
    gen_ex1_milestone_div(xlsInfo)
# end_def 

# TBD: Combine this and gen_exhibit_2_body into a single, parameterized,  routine.
def gen_exhibit_1_initial_boilerplate():
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
# end_def gen_exhibit_1_initial_boilerplate()

# Shares 100% code with gen_exhibit_1_final_boilerplate. 
# TBD: Combine these two routines.
# Write the final "boilerplate" HTML for Exhibit 1: the closing </body> and </html> tags.
def gen_exhibit_1_final_boilerplate():
    s = '</body>' 
    appendHTML(s)
    s = '</html>'
    appendHTML(s)
# end_def gen_exhibit_1_final_boilerplate()


def gen_exhibit_1(xlsInfo):
    gen_exhibit_1_initial_boilerplate()
    gen_exhibit_1_body(xlsInfo)
    gen_exhibit_1_final_boilerplate()
# end_def gen_exhibit_1()

# Write initial "boilerplate" HTML for Exhibit 2.
# This includes all content from DOCTYPE, the <html> tag, and everything in the <head>.
def gen_exhibit_2_initial_boilerplate():
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
# end_def gen_exhibit_2_initial_boilerplate()

# This writes the final "boilerplate" HTML for Exhibit 2: the closing </body> and </html> tags.
def gen_exhibit_2_final_boilerplate():
    s = '</body>' 
    appendHTML(s)
    s = '</html>'
    appendHTML(s)
# end_def gen_exhibit_2_final_boilerplate()

def gen_ex2_direct_salary_div(xlsInfo):
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
# end_def gen_ex2_direct_salary_div()

######################################################################################################
# Helper function to generate <tr> (and its contents) for one task in the salary cost table.
# This function is called only from gen_ex2_salary_cost_table_div, which it is LOGICALLY nested within.
# In order to expedite development/prototyping, however, it is currently defined here at scope-0.
# When the tool has become stable, move it within the def of salary_cost_table_div.
#
def gen_task_tr(task_num, task_row_ix, xlsInfo, real_cols_info):
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
# end_def gen_task_tr()

############################################################################
# Top-level routine for generating HTML for Exhibit 2 salary cost table div.
# Calls end_def gen_ex2_task_tr as a helper function.
def gen_ex2_salary_cost_table_div(xlsInfo):
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
        gen_task_tr(i, task_row_ix, xlsInfo, real_cols_info)
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
# end_def gen_ex2_salary_cost_table_div()

def gen_ex2_other_direct_costs_div(xlsInfo):
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
    def gen_odc(name, cost):
        s = '<div class="otherExpDiv">'
        appendHTML(s)
        s = '<div class="otherExpDescDiv">' + name + '</div>'
        appendHTML(s)
        s = '<div class="otherExpAmtDiv">' + '$' + format_dollars(cost) + '</div>'
        appendHTML(s)
        s = '</div>'
        appendHTML(s)
    # end_def gen_odc()
    
    # Travel
    travel = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_travel_line_ix'], xlsInfo['total_cost_col_ix'])
    if travel != 0:
        gen_odc('Travel', travel)
    
    # General office equipment
    general_office_equipment = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_office_equipment_line_ix'], xlsInfo['total_cost_col_ix'])
    if general_office_equipment != 0:
        gen_odc('General Office Equipment', general_office_equipment)
    
    # Data processing equipment
    dp_equipment = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_dp_equipment_line_ix'], xlsInfo['total_cost_col_ix'])
    if dp_equipment != 0:
        gen_odc('Data Processing Equipent', dp_equipment)
    
    # Consultant(s)
    consultants = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_consultants_line_ix'], xlsInfo['total_cost_col_ix'])
    if consultants != 0:
        gen_odc('Consultants', consultants)
    
    # Printing
    printing = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_printing_line_ix'], xlsInfo['total_cost_col_ix'])
    if printing != 0:
        gen_odc('Printing', printing)
    
    # Other 
    other = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_other_line_ix'], xlsInfo['total_cost_col_ix'])
    if other != 0:
        desc = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_other_line_ix'], xlsInfo['task_name_col_ix'])
        gen_odc(desc, other)
    
    # </div> for wrapper
    s = '</div>'
    appendHTML(s)
# end_def gen_ex2_other_direct_costs_div()

def gen_ex2_total_direct_costs_div(xlsInfo):
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
# end_def gen_ex2_total_direct_costs_div()

def gen_ex2_funding_div(xlsInfo):
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
# end_def gen_ex2_funding_div()

# This writes the HTML for the entire <body> of Exhibit 2, including:
#   the opening <body> tag
#   initial content (e.g., project name, etc.)
#   the div for the "Direct Salary and Overhead" line
#   the div for the salary cost table
#   the div forthe "Other Direct Costs" line
#   the div for funding source(s)
def gen_exhibit_2_body(xlsInfo):
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
    gen_ex2_direct_salary_div(xlsInfo)
    gen_ex2_salary_cost_table_div(xlsInfo)
    gen_ex2_other_direct_costs_div(xlsInfo)
    gen_ex2_total_direct_costs_div(xlsInfo)
    gen_ex2_funding_div(xlsInfo)
# end_def gen_exhibit_2_body()

def gen_exhibit_2(xlsInfo):
    gen_exhibit_2_initial_boilerplate()
    gen_exhibit_2_body(xlsInfo)
    gen_exhibit_2_final_boilerplate()
# end_def gen_exhibit_2()

# Pretty-formats HTML and saves it to specified filename.
def write_html_to_file(html, filename):
    soup = BeautifulSoup(html, 'html.parser')
    pretty_html = soup.prettify() + '\n'
    o = open(filename, 'w')
    # NOTE: We need to encode the output as UTF-8 because it may contain non-ASCII characters,
    # e.g., the "section" symbol used to identify funding sources such as <section>5303 ...
    o.write(pretty_html.encode("UTF-8"))
    o.close()
# end_def write_html_to_file()

# Main driver routine - this function does NOT launch a GUI.
def main(fullpath):
    global accumulatedHTML # Yeech
    t1 = os.path.split(fullpath)
    in_dir = t1[0]
    in_fn = t1[1]
    in_fn_wo_suffix = os.path.splitext(in_fn)[0]
    ex_1_out_html_fn = in_dir + '\\' + in_fn_wo_suffix + '_Exhibit_1.html'
    ex_2_out_html_fn = in_dir + '\\' + in_fn_wo_suffix + '_Exhibit_2.html'
    
    # Collect 'navigation' information from input .xlsx file
    xlsInfo = initExcelFile(fullpath)
    
    # Generate Exhibit 1 HTML, and save it to disk
    # NOTE: gen_exhibit_1() is currently a work-in-progress
    accumulatedHTML = ''
    gen_exhibit_1(xlsInfo)
    write_html_to_file(accumulatedHTML, ex_1_out_html_fn)
    
    # Generate Exhibit 2 HTML, and save it to disk
    accumulatedHTML = ''
    gen_exhibit_2(xlsInfo)
    write_html_to_file(accumulatedHTML, ex_2_out_html_fn)
# end_def main()
