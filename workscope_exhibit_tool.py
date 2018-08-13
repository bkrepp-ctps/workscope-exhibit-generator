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
# in such a way as to make it as simple as possible to correlate a given section of 
# Python code that generates an HTML fragment with the HTML output produced by the
# CFML application and the relevant segment of CFML code  As there is no functional 
# spec for the CFML application, this was essential in order to ensure functional 
# correctness and debug-ability. Please note that as a consequence of this, 
# the code is neither particularly efficient nor particularly idiomatic Python. 
# It was, however, intended to be as easy to understand by a 'newbie' as possible.
#
# Author: Benjamin Krepp
# Date: 23-27 July, 30-31 July, 6-10 August 2018, 13 August 2018
#   
# Internals of this Module: Top-level Functions
# =============================================
#
# main - main driver routine for this program
#
# gen_exhibit_1 - driver routine for generating Exhibit 1;
#                 calls gen_exhibit_1_initial_boilerplate,
#                 gen_exhibit_1_body, and gen_exhibit_1_final_boilerplate
#
# gen_exhibit_1_initial_boilerplate - generates boilerplate HTML at beginning of
#                                     Exhibit 1
#
# gen_exhibit_1_final_boilerplate - generates boilerplate HTML at end of Exhibit 1
#
# gen_exhibit_1_body - driver routine for producing HTML for the body of Exhibit 1;
#                      calls gen_ex1_schedule_table and  gen_ex1_milestone_div
#
# gen_ex1_schedule_table - driver routine for generating HTML for the schedule
#                          <table> in Exhibit 1
#
# gen_ex1_schedule_table_body - driver routine for generating the <body> of the
#                               HTML <table> schedule in Exhibit 1; calls
#                               gen_ex1_task_tr
#
# gen_ex1_task_tr - driver routeine for generating the <tr> for a given task in 
#                   the schedule in Exhibit 1; calls gen_ex1_task_tr_2nd_td
#
# gen_ex1_task_tr_2nd_td - routine responsible for generating the second <td>
#                          in the <tr> for a given task in the schedule in 
#                          Exhibit 1. This routine bears close reading.
#
# gen_ex1_milestone_div - generates HTML for the milestones/deliverables <div>
#                         of Exhibit 1
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
# col_ix_to_temporal_string - maps a column index in the schedule portion of the input 
#                             .xlsx file to a text string that expresses the point in 
#                             time indicated by the input column index in terms of the 
#                             major- and minor-units of the schedule.
#                             Example: Map column index X to "Month 3, Week 1"
#
###############################################################################

import os
import sys
import math
import re
import openpyxl
from bs4 import BeautifulSoup
from excelFileManager import initExcelFile, get_column_index, get_row_index, get_cell_contents, \
                             get_last_used_sched_column, MAGIC_FILL_STYLE, \
                             dump_xlsInfo
from stringAccumulator import stringAccumulator

debug_flags = {}
debug_flags['dump_sched_elements'] = False

# Global pseudo-constants:
# Width of table HEADER cells in the schedule table
SCHED_HEADER_CELL_WITDH_IN_PTS_12PX_BORDER = 33.9375
SCHED_HEADER_CELL_WITDH_IN_PX_12PX_BORDER = (SCHED_HEADER_CELL_WITDH_IN_PTS_12PX_BORDER * 1.3333)
SCHED_HEADER_CELL_WIDTH_IN_PTS_24PX_BORDER = 16.59375
SCHED_HEADER_CELL_WIDTH_IN_PX_24PX_BORDER = (SCHED_HEADER_CELL_WIDTH_IN_PTS_24PX_BORDER * 1.3333)


# Person weeks are formatted as a floating point number with one digit of precision.
def format_person_weeks(person_weeks):
    retval = "%.1f" % person_weeks
    return retval
# end_def format_person_weeks()

# Dollars are formatted as a floating point number with NO digits of precision,
# i.e., as an integer, but also with commas as the thousands delimiter.
# Note: This function does NOT prepend a '$' symbol to the string returned.
def format_dollars(dollars):
    retval = '{0:,.0f}'.format(dollars)
    return retval
# end_def format_dollars()

# Map a column index in the schedule portion of the input .xlsx file to
# a text string that expresses the point in time indicated by the 
# input column index in terms of the major- and minor-units of the schedule.
# Example: Map column index X to "Month 3, Week 1"
def col_ix_to_temporal_string(col_ix, xlsInfo):
    retval = ''
    maj_unit = xlsInfo['sched_major_units']
    min_unit = xlsInfo['sched_minor_units']
    num_subdivisions = xlsInfo['num_sched_subdivisions']
    
    # The trick  here is to remember that after 'unbiasing' the input column index
    # by the index of the first column in the schedule, the result will be 0-based,
    # whereas human beings think of the first <time unit> of a schedule as <time unit> 1.    
    start_abs = (col_ix - xlsInfo['first_schedule_col_ix']) + num_subdivisions
    maj_abs = start_abs / num_subdivisions
    # The same principle applies to the minor schedule units
    min_abs = (start_abs % num_subdivisions) + 1
    retval = maj_unit + ' ' + str(maj_abs) + ', ' + min_unit + ' ' + str(min_abs)
        
    # Debug
    # print '*** retval: ' + retval
    return retval
# end_def_col_ix_to_temporal_string()

def gen_ex1_task_tr_2nd_td(htmlAcc, task_num, task_row_ix, xlsInfo):
    global SCHED_HEADER_CELL_WITDH_IN_PX_12PX_BORDER, SCHED_HEADER_CELL_WIDTH_IN_PX_24PX_BORDER
    global debug_flags

    t1 = '<td colspan="' + str(xlsInfo['num_sched_col_header_cells']) + '" '
    # *** TBD: 'timeUnit1' seems to ALWAYS be incuded as a header. Is this right?
    t2 = 'headers ="row' + str(task_num) + ' timeUnit1" '
    if task_num == 1:
        t3 = 'class="firstSchedColCell">'
    else:
        t3 = 'class="schedColCell">'
    # end_if
    s = t1 + t2 + t3
    htmlAcc.append(s)
    
    # The guts of 2nd <td> in schedule row.
    # This may contain an arbitrary number of chart 'bars' and an arbitrary number
    # of 'milestones'. Each of these is placed in a <div> of its own,  generated in 
    # ascending chronological order.
    #
    # In order to this, we build (1) a list of 'bars' and (2) a list of 'milestones,'
    # both in ascending chronological (i.e., column) order. We use a common dictionary 
    # data structure for the elements of these two lists:
    #     'type'      : 'bar' or 'milestone'
    #     'start'     : start column index
    #     'end'       : end column index   ('start' == 'end' for milestones)
    #     'milestone' : if item is a milestone, the milestone letter, otherwise ''
    # We then merge the two lists, maintaining ascending chronological order.  
    # Having done this, we will then be in a position to generate the relevant HTML.
    
    # Build the list of 'bars'
    # Prep work: Generate a string of 0's and 1's indicating the cells in the schedule
    # bar chart that have been 'filled in' with the magic fill pattern 'gray125'.
    # Logically, we're creating a bit vector; it's implemented, however as a vector
    # of '0' and '1' characters in a string.
    my_pseudo_bv = ''
    ws = xlsInfo['ws']
    for col in range(xlsInfo['first_schedule_col_ix'],xlsInfo['last_used_schedule_col_ix']+1):
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
        # in the (logical) bitvector by the index of the first column in the schedule table
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
    
    # Build the list of 'milestones'
    milestones = []
    for col in range(xlsInfo['first_schedule_col_ix'],xlsInfo['last_used_schedule_col_ix']+1):
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
    #      and so thereby ensures that the HTML for task bars are generated before
    #      the HTML for any milesones occuring within them.
    big_list = bars + milestones
    big_list_sorted = sorted(big_list, key=lambda x: (x['start'], x['type']))

    if debug_flags['dump_sched_elements']:
        for thing in big_list_sorted:
            print '*** ' + thing['type'] + ' ' + str(thing['start']) + ' ' + str(thing['end']) + ' ' + thing['milestone']
        # end_if
    # end_for
    
    # The schedule table BODY in the output HTML does not consist of "real" <td> cells.
    # Rather, we generate <div>s whose left offset and width are determined by:
    #     1. the index of the first cell of the relevant schedule item in the INPUT .xlsx file
    #     2. the number of cells for the relevant schedule item in the INPUT .xlsx file
    #     3. the width (in pixels) of the schedule table HEADER cells in the output HTML
    #     4. the number of minor schedule units per major schedule unit in the input .xlsx file
    
    if xlsInfo['num_sched_col_header_cells'] <= 12:
        hdr_cell_width = SCHED_HEADER_CELL_WITDH_IN_PX_12PX_BORDER
    else:
        hdr_cell_width = SCHED_HEADER_CELL_WITDH_IN_PX_24PX_BORDER
    # end_if

    # Width of 'virtual' cell for one subdivision of the major schedule unit
    minor_cell_width  = hdr_cell_width / float(xlsInfo['num_sched_subdivisions'])
    
    # Debug
    # print 'Header cell width = ' + str(hdr_cell_width)
    # print 'Minor cell width = ' + str(minor_cell_width)
    
    # Generation of the <divs> for the schedule bars and milestones
    for item in big_list_sorted:
        left = float(item['start'] - xlsInfo['first_schedule_col_ix']) *  minor_cell_width
        if item['type'] == 'bar':
            num_subdivisions = item['end'] - item['start'] + 1
            width = num_subdivisions * minor_cell_width
            
            # Debug
            # print '*** Task #' + str(task_num) +  ' start: ' + str(item['start']) + ' end: ' + str(item['end']) + ' ' + ' left = ' + str(left) + ' width = ' + str(width)
            
            t1 = '<div class="schedElemDiv">'
            t2 = '<div class="scheduleBar" style="'
            t2 += 'left:' + str(left) + 'px;'
            t2 += 'width:' + str(width) + 'px'
            t2 += '">'
            # Stuff for screen reader
            t3 = '<div class="overflowHiddenTextDiv">'
            t4 = 'From ' + col_ix_to_temporal_string(item['start'], xlsInfo)
            t5 = ' to ' + col_ix_to_temporal_string(item['end'], xlsInfo) + '.'
            t6 = '</div>'
            # End of stuff for screen reader
            # Close <div> with class=schedBar 
            t7 = '</div>'
            # Close <div> with class=schedElemDiv
            t8 = '</div>'
            s = t1 + t2 + t3 + t4 + t5 + t6 + t7 + t8
            htmlAcc.append(s)
        else:
            # Must be a 'milestone'
            # Debug
            # print '*** Milestone: ' + item['milestone'] + ' start: ' + str(item['start']) +  ' ' + ' left = ' + str(left)
            t1 = '<div class="schedElemDiv">'
            t2 = '<div class="deliverableCodeDiv" style="'
            t2 += 'left:' + str(left) + 'px;">'
            # Firt bunch of stuff for screen reader
            t3 = '<div class="overflowHiddenTextDiv">Deliverable</div>'
            # The name of the milestone/deliverable
            t4 = item['milestone']
            
            # Second bunch of stuff for screen reader - text of when the milestone/deliverable will arrive
            t5 = '<div class="overflowHiddenTextDiv">'
            t6 = 'Delivered by ' + col_ix_to_temporal_string(item['start'], xlsInfo) + '.'
            t7 = '</div>'
            
            # Close the remaining two <div>s
            t8 = '</div></div>'
            s = t1 + t2 + t3 + t4 + t5 + t6 + t7 + t8
            htmlAcc.append(s)
        # end_if
    #end_for
    
    # Close seond <td> in row
    s = '</td>'
    htmlAcc.append(s)
# end_def gen_ex1_task_tr_2nd_td()

def gen_ex1_task_tr(htmlAcc, task_num, task_row_ix, xlsInfo):
    s = '<tr>'
    htmlAcc.append(s)
      
    # First <td> in row: task number and task name
    t1 = '<td id="row' + str(task_num) + '" headers="ex1taskTblHdr" '
    if task_num == 1:
        t2 = 'class="firstTaskTblCell">'
    else:
        t2 = 'class="taskTblCell">'
    # end_if
    s = t1 + t2
    htmlAcc.append(s)
    
    t1 = '<div class="taskNumDiv">'
    # *** TBD: Fetch task number from cell in Excel file rather than using task_num
    t2 = str(task_num) + '.'
    t3 = '</div>'
    s = t1 + t2 + t3
    htmlAcc.append(s)
    
    t1 = '<div class="taskNameDiv">'
    #  *** TBD: This currently gets the task name from its cell in the cost table
    t2 = get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['task_name_col_ix'])
    t3 = '</div>'
    s = t1 + t2 + t3
    htmlAcc.append(s)
    # Close first <td> in row
    s = '</td>'
    htmlAcc.append(s)
    
    # Second <td> in row: schedule bar(s) and deliverable(s), (if any)
    gen_ex1_task_tr_2nd_td(htmlAcc, task_num, task_row_ix, xlsInfo)
    
    # Close <tr>
    s = '</tr>'
    htmlAcc.append(s)
# end_def gen_ex1_task_tr()

def gen_ex1_schedule_table_body(htmlAcc, xlsInfo):
    # Open <tbody>
    s = '<tbody>'
    htmlAcc.append(s)
    # Write the <tr>s in the table body
    i = 0
    for task_row_ix in range(xlsInfo['task_list_top_row_ix']+1,xlsInfo['task_list_bottom_row_ix']):
        i = i + 1
        gen_ex1_task_tr(htmlAcc, i, task_row_ix, xlsInfo)
    # end_for
    # Close <tbody>
    s = '</tbody>'
    htmlAcc.append(s)
    # Close <table>
    s = '</table>'
    htmlAcc.append(s)
# end_def gen_ex1_schedule_table_body()

def gen_ex1_schedule_table(htmlAcc, xlsInfo):
    s = '<table id="ex1Tbl"'
    s += 'summary="Breakdown of schedule by tasks in column one and calendar time ranges and deliverable dates in column two.">'
    htmlAcc.append(s)
    
    s = '<thead>'
    # First row of column header, first column: 'Task'
    htmlAcc.append(s)
    s = '<tr>'
    htmlAcc.append(s)
    s = '<th id="ex1taskTblHdr" class="colTblHdr" rowspan="2"><br>Task</th>'
    htmlAcc.append(s)
    
    # First row of table header, second column: name of MAJOR time unit used in table,
    # i.e., either 'Quarter', 'Month' or 'Week'
    #

    t1 = '<th id="ex1weekTblHeader" class="colTblHdr"'
    t2 = 'colspan="' + str(xlsInfo['num_sched_col_header_cells']) + '">' 
    t3 = xlsInfo['sched_major_units']
    t4 = '</th>'
    s = t1 + t2 +t3 + t4
    htmlAcc.append(s)
    s = '</tr>'
    htmlAcc.append(s)
    
    # Second row of table header: the numbers of the individual MAJOR time units in schedule
    s = '<tr>'
    htmlAcc.append(s)
    # The <th>s for the second row of headers,
    # the numbers of the MAJOR schedule units actually used in the schedule
    
    if xlsInfo['num_sched_col_header_cells'] <= 12:
        sched_header_cell_class_string = ' class="scheduleColHdr12PixBorder" '
    else:
        sched_header_cell_class_string = ' class="scheduleColHdr24PixBorder" '
    # end_if
    
    for i in range(1,xlsInfo['num_sched_col_header_cells']+1):
        t1 = '<th id='
        t2 = '"timeUnit' + str(i) + '"'
        t3 = sched_header_cell_class_string + ' abbr="Schedule range">'
        t4 = str(i) + '</th>'
        s = t1 + t2 + t3 + t4
        htmlAcc.append(s)
    # end_for
    
    # Close the 2nd row of column headers
    s = '</tr>'
    htmlAcc.append(s)
    
    # Close table header
    s = '</thead>'
    htmlAcc.append(s)
  
    # Call subordinate routine to do the heavy lifting: generate the <table> body for Exhibit 1
    gen_ex1_schedule_table_body(htmlAcc, xlsInfo)
# end_def gen_ex1_schedule_table()


def gen_ex1_milestone_div(htmlAcc, xlsInfo):
    s = '<div id="milestoneDiv">'
    htmlAcc.append(s)
    s = '<div id="milestoneHdrDiv">'
    htmlAcc.append(s)
    s = 'Products/Milestones'
    htmlAcc.append(s)
    s = '</div>'
    htmlAcc.append(s)
    s = '<div id="milestoneListDiv">'
    htmlAcc.append(s)
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
        htmlAcc.append(s)
    # end_for
    
    s = '</div>'
    htmlAcc.append(s)
    s = '</div>'
    htmlAcc.append(s)
# end_def gen_ex1_milestone_div()


def gen_exhibit_1_body(htmlAcc, xlsInfo):
    pass
    s = '<body style="text-align:center;padding:0pt;margin:0pt;">'
    htmlAcc.append(s)
    s = '<div id="exhibit1">'
    htmlAcc.append(s)
    s = '<div class="exhibitPageLayoutDiv1"><div class="exhibitPageLayoutDiv2">'
    htmlAcc.append(s)
    s = '<h1>'
    htmlAcc.append(s)
    s = 'Exhibit 1<br>'
    htmlAcc.append(s)
    s = 'ESTIMATED SCHEDULE<br>'
    htmlAcc.append(s)
    # Project name
    s = str(get_cell_contents(xlsInfo['ws'], xlsInfo['project_name_cell_row_ix'], xlsInfo['project_name_cell_col_ix']))
    s = s + '<br>'
    htmlAcc.append(s)
    s = '</h1>'
    htmlAcc.append(s)
    #
    gen_ex1_schedule_table(htmlAcc, xlsInfo)
    gen_ex1_milestone_div(htmlAcc, xlsInfo)
# end_def 

# TBD: Combine this and gen_exhibit_2_body into a single, parameterized,  routine.
def gen_exhibit_1_initial_boilerplate(htmlAcc):
    s = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
    htmlAcc.append(s)
    s = '<html xmlns="http://www.w3.org/1999/xhtml" lang="en"><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">'
    htmlAcc.append(s)
    s = '<title>CTPS Work Scope Exhibit 1</title>'
    htmlAcc.append(s)
    s = '<link rel="stylesheet" type="text/css" href="./ctps_work_scope_print.css">'
    htmlAcc.append(s)
    s = '</head>'
    htmlAcc.append(s)
# end_def gen_exhibit_1_initial_boilerplate()

# Shares 100% code with gen_exhibit_1_final_boilerplate. 
# TBD: Combine these two routines.
# Write the final "boilerplate" HTML for Exhibit 1: the closing </body> and </html> tags.
def gen_exhibit_1_final_boilerplate(htmlAcc):
    s = '</body>' 
    htmlAcc.append(s)
    s = '</html>'
    htmlAcc.append(s)
# end_def gen_exhibit_1_final_boilerplate()


def gen_exhibit_1(htmlAcc, xlsInfo):
    gen_exhibit_1_initial_boilerplate(htmlAcc)
    gen_exhibit_1_body(htmlAcc, xlsInfo)
    gen_exhibit_1_final_boilerplate(htmlAcc)
# end_def gen_exhibit_1()

# Write initial "boilerplate" HTML for Exhibit 2.
# This includes all content from DOCTYPE, the <html> tag, and everything in the <head>.
def gen_exhibit_2_initial_boilerplate(htmlAcc):
    s = '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">'
    htmlAcc.append(s)
    s = '<html xmlns="http://www.w3.org/1999/xhtml" lang="en"><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">'
    htmlAcc.append(s)
    s = '<title>CTPS Work Scope Exhibit 2</title>'
    htmlAcc.append(s)
    s = '<link rel="stylesheet" type="text/css" href="./ctps_work_scope_print.css">'
    htmlAcc.append(s)
    s = '</head>'
    htmlAcc.append(s)
# end_def gen_exhibit_2_initial_boilerplate()

# This writes the final "boilerplate" HTML for Exhibit 2: the closing </body> and </html> tags.
def gen_exhibit_2_final_boilerplate(htmlAcc):
    s = '</body>' 
    htmlAcc.append(s)
    s = '</html>'
    htmlAcc.append(s)
# end_def gen_exhibit_2_final_boilerplate()

def gen_ex2_direct_salary_div(htmlAcc, xlsInfo):
    s = '<div id="directSalaryDiv" class="barH2">'
    htmlAcc.append(s)
    s = '<h2>Direct Salary and Overhead</h2>'
    htmlAcc.append(s)
    t1 = '<div class="h2AmtDiv">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['direct_salary_cell_row_ix'], xlsInfo['direct_salary_cell_col_ix']))
    t3 = '</div>'
    s = t1 + t2 + t3
    htmlAcc.append(s)
    s = '</div>'
    htmlAcc.append(s)
# end_def gen_ex2_direct_salary_div()

######################################################################################################
# Helper function to generate <tr> (and its contents) for one task in the salary cost table.
# This function is called only from gen_ex2_salary_cost_table_div, which it is LOGICALLY nested within.
# In order to expedite development/prototyping, however, it is currently defined here at scope-0.
# When the tool has become stable, move it within the def of salary_cost_table_div.
#
def gen_task_tr(htmlAcc, task_num, task_row_ix, xlsInfo, real_cols_info):
    # Open <tr> element
    t1 = '<tr id='
    tr_id = 'taskHeader' + str(task_num)
    t2 = tr_id + '>'
    s = t1 + t2
    htmlAcc.append(s)
    
    # <td> for task number and task name
    # Note: This contains 3 divs organized thus: <div> <div></div> <div></div> </div>
    t1 = '<td headers="taskTblHdr" scope="row" '
    if task_num == 1:
        t2  = 'class="firstTaskTblCell">'
    else:
        t2 = 'class="taskTblCell">'
    # end_if
    s = t1 + t2 
    htmlAcc.append(s)
    # Open outer div
    s = '<div class="taskTblCellDiv">'
    htmlAcc.append(s)
    # First inner div
    t1 = '<div class="taskNumDiv">'
    t2 = get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['task_number_col_ix'])
    t3 = '</div>'
    s = t1 + t2 + t3
    htmlAcc.append(s)
    # Second inner div
    t1 = '<div class="taskNameDiv">'
    t2 = get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['task_name_col_ix'])
    t3 = '</div>'
    s = t1 + t2 + t3
    htmlAcc.append(s)
    # Close outer div, and close <td>
    s = '</div>'
    htmlAcc.append(s)
    s = '</td>'
    htmlAcc.append(s)
    
    # Generate the <td>s for all the salary grades used in this work scope exhibit
    for col_info in real_cols_info:
        t1 = '<td headers="' + tr_id + ' personWeekTblHdr ' + col_info['col_header_id'] + '"'
        t2 = ' class="rightPaddedTblCell">'
        t3 = format_person_weeks(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo[col_info['col_ix']]))
        t4 = '</td>'
        s = t1 + t2 + t3 + t4
        htmlAcc.append(s)
    # end_for
    
    # Generate the <td>s for 'Total [person weeks]', 'Direct Salary', 'Overhead', and 'Total Cost'.
    #
    # Total [person weeks]
    t1 = '<td headers="' + tr_id + ' personWeekTblHdr personWeekTotalTblHdr" class="rightPaddedTblCell">'
    t2 = format_person_weeks(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['total_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    htmlAcc.append(s)
    #
    # Direct Salary
    t1 = '<td headers="' + tr_id + ' salaryTblHdr" class="rightPaddedTblCell">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['direct_salary_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    htmlAcc.append(s)
    #
    # Overhead
    t1 = '<td headers="' + tr_id + ' overheadTblHdr" class="rightPaddedTblCell">'
    t2 = '$' +  format_dollars(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['overhead_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    htmlAcc.append(s)       
    #
    # Total Cost
    t1 = '<td headers="' + tr_id + ' totalTblHdr" class="rightPaddedTblCell">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], task_row_ix, xlsInfo['total_cost_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    htmlAcc.append(s)       
    
    s = '</tr>'
    htmlAcc.append(s)
# end_def gen_task_tr()

############################################################################
# Top-level routine for generating HTML for Exhibit 2 salary cost table div.
# Calls end_def gen_ex2_task_tr as a helper function.
def gen_ex2_salary_cost_table_div(htmlAcc, xlsInfo):
    s = '<div class="costTblDiv">'
    htmlAcc.append(s)
    s = '<table id="ex2Tbl" summary="Breakdown of staff time by task in column one, expressed in person weeks for each implicated pay grade in the middle columns,'
    s = s + 'together with resulting total salary and associated overhead costs in the last columns.">'
    htmlAcc.append(s)
    
    # The table header (<thead>) element and its contents
    #
    s = '<thead>'
    htmlAcc.append(s)
    
    # <thead> contents
    # Most of this is invariant bolierplate. The exceptions are the number of "real" columns and the overhead rate.
    
    # First row of <thead> contents
    # 
    s = '<tr>'
    htmlAcc.append(s)
    s = '<th id="taskTblHdr" class="colTblHdr" rowspan="2" scope="col"><br>Task</th>'
    htmlAcc.append(s)
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
    htmlAcc.append(s)
    s = '<th id="salaryTblHdr" class="colTblHdr" rowspan="2" scope="col" abbr="Direct Salary">Direct<br>Salary</th>'
    htmlAcc.append(s)
    t1 = '<th id="overheadTblHdr" class="colTblHdr" rowspan="2" scope="col" abbr="Overhead">Overhead<br>'
    t2 = get_cell_contents(xlsInfo['ws'], xlsInfo['overhead_cell_row_ix'], xlsInfo['overhead_cell_col_ix'])
    t2 = t2.replace('@ ', '')
    t3 = '</th>'
    s = t1 + t2 + t3 
    htmlAcc.append(s)
    s = '<th id="totalTblHdr" class="colTblHdr" rowspan="2" scope="col" abbr="Total Cost">Total<br>Cost</th>'
    htmlAcc.append(s)
    s = '</tr>'
    htmlAcc.append(s)
    
    # Second row of <thead> contents
    #
    s = '<tr>'
    htmlAcc.append(s)
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
        htmlAcc.append(s)
    # end_for
    # Second: column header for Total column
    s = '<th id="personWeekTotalTblHdr" scope="col">Total</th>'
    htmlAcc.append(s)
    s = '</tr>'
    htmlAcc.append(s)
    
    # Close <thead> 
    s = '</thead>'
    htmlAcc.append(s)   
    
    # The table body <tbody> element and its contents.
    #
    s = '<tbody>'
    htmlAcc.append(s)
    
    # <tbody> contents.
    #
    # Write <tr>s for each task in the task list.
    i = 0
    for task_row_ix in range(xlsInfo['task_list_top_row_ix']+1,xlsInfo['task_list_bottom_row_ix']):
        i = i + 1
        gen_task_tr(htmlAcc, i, task_row_ix, xlsInfo, real_cols_info)
    # end_for
    
    # The 'Total' row
    #
    s = '<tr>'
    htmlAcc.append(s)
    s = '<td headers="taskTblHdr" id="totalRowTblHdr" class="taskTblCell" scope="row" abbr="Total All Tasks">'
    htmlAcc.append(s)
    s = '<div class="taskTblCellDiv">'
    htmlAcc.append(s)
    # Total row, task number column (empty)
    s = '<div class="taskNumDiv"> </div>'
    htmlAcc.append(s)
    # Total row, "task name" colum - which contains the pseudo task name 'Total'
    s = '<div class="taskNameDiv">Total</div>'
    htmlAcc.append(s)
    s = '</div>'
    htmlAcc.append(s)
    s = '</td>'
    htmlAcc.append(s)
    
    # Total row: columns for salary grades used in this workscope
    for col_info in real_cols_info:
        t1 = '<td headers="totalRowTblHdr personWeekTblHdr ' + col_info['col_header_id'] + '" class="totalRowTblCell">'
        t2 = format_person_weeks(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo[col_info['col_ix']]))
        t3 = '</td>'
        s = t1 + t2 + t3
        htmlAcc.append(s)
    # end_for
    
    # Total row: Total [person weeks] column
    t1 = '<td id="personWeeksTotalRowTblCell" headers="totalRowTblHdr personWeekTblHdr personWeekTotalTblHdr" class="totalRowTblCell">'
    t2 = format_person_weeks(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo['total_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    htmlAcc.append(s)
    # Total row, direct salary column
    t1 = '<td id="directSalaryTotalRowTblCell" headers="totalRowTblHdr salaryTblHdr" class="totalRowTblCell">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo['direct_salary_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    htmlAcc.append(s)
    # Total row, overhead column
    t1 = '<td id="overheadTotalRowTblCell" headers="totalRowTblHdr overheadTblHdr" class="totalRowTblCell">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo['overhead_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    htmlAcc.append(s)
    # Total row, total cost column
    t1 = '<td id="totalTotalRowTblCell" headers="totalRowTblHdr totalTblHdr" class="totalRowTblCell">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['total_line_row_ix'], xlsInfo['total_cost_col_ix']))
    t3 = '</td>'
    s = t1 + t2 + t3
    htmlAcc.append(s)
    # Close <tr> for Total row
    s = '</tr>'
    htmlAcc.append(s)
    
    # Close <tbody>, <table>, and <div>
    s = '</tbody>'
    htmlAcc.append(s)
    s = '</table>'
    htmlAcc.append(s)
    s = '</div>'
    htmlAcc.append(s)
# end_def gen_ex2_salary_cost_table_div()

def gen_ex2_other_direct_costs_div(htmlAcc, xlsInfo):
    s = '<div id="otherDirectDiv" class="barH2">'
    htmlAcc.append(s)
    s = '<h2>Other Direct Costs</h2>'
    htmlAcc.append(s)
    t1 = '<div class="h2AmtDiv">'
    odc_total = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_cell_row_ix'], xlsInfo['odc_cell_col_ix'])
    t2 = '$' + format_dollars(odc_total)
    t3 = '</div>'
    s = t1 + t2 + t3
    htmlAcc.append(s)
    s = '</div>'
    htmlAcc.append(s)
    # Write the divs for the specific other direct costs and a wrapper div around all of them (even if there are none.)
    #
    # <div> for wrapper
    s = '<div class="costTblDiv">'
    htmlAcc.append(s)
    
    # Utiltiy function to write HTML for one kind of 'other direct cost.'
    def gen_odc(htmlAcc, name, cost):
        s = '<div class="otherExpDiv">'
        htmlAcc.append(s)
        s = '<div class="otherExpDescDiv">' + name + '</div>'
        htmlAcc.append(s)
        s = '<div class="otherExpAmtDiv">' + '$' + format_dollars(cost) + '</div>'
        htmlAcc.append(s)
        s = '</div>'
        htmlAcc.append(s)
    # end_def gen_odc()
    
    # Travel
    travel = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_travel_line_ix'], xlsInfo['total_cost_col_ix'])
    if travel != 0:
        gen_odc(htmlAcc, 'Travel', travel)
    
    # General office equipment
    general_office_equipment = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_office_equipment_line_ix'], xlsInfo['total_cost_col_ix'])
    if general_office_equipment != 0:
        gen_odc(htmlAcc, 'General Office Equipment', general_office_equipment)
    
    # Data processing equipment
    dp_equipment = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_dp_equipment_line_ix'], xlsInfo['total_cost_col_ix'])
    if dp_equipment != 0:
        gen_odc(htmlAcc, 'Data Processing Equipent', dp_equipment)
    
    # Consultant(s)
    consultants = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_consultants_line_ix'], xlsInfo['total_cost_col_ix'])
    if consultants != 0:
        gen_odc(htmlAcc, 'Consultants', consultants)
    
    # Printing
    printing = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_printing_line_ix'], xlsInfo['total_cost_col_ix'])
    if printing != 0:
        gen_odc(htmlAcc, 'Printing', printing)
    
    # Other 
    other = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_other_line_ix'], xlsInfo['total_cost_col_ix'])
    if other != 0:
        desc = get_cell_contents(xlsInfo['ws'], xlsInfo['odc_other_line_ix'], xlsInfo['task_name_col_ix'])
        gen_odc(htmlAcc, desc, other)
    
    # </div> for wrapper
    s = '</div>'
    htmlAcc.append(s)
# end_def gen_ex2_other_direct_costs_div()

def gen_ex2_total_direct_costs_div(htmlAcc, xlsInfo):
    s = '<div id="totalDirectDiv" class="barH2">'
    htmlAcc.append(s)
    s = '<h2>TOTAL COST</h2>'
    htmlAcc.append(s)
    t1 = '<div class="h2AmtDiv">'
    t2 = '$' + format_dollars(get_cell_contents(xlsInfo['ws'], xlsInfo['total_cost_cell_row_ix'], xlsInfo['total_cost_cell_col_ix']))
    t3 = '</div>'
    s = t1 + t2 + t3
    htmlAcc.append(s)
    s = '</div>'
    htmlAcc.append(s)
# end_def gen_ex2_total_direct_costs_div()

def gen_ex2_funding_div(htmlAcc, xlsInfo):
    s = '<div id="fundingDiv">'
    htmlAcc.append(s)
    s = '<div id="fundingHdrDiv">'
    htmlAcc.append(s)
    s = 'Funding'
    htmlAcc.append(s)
    s = '</div>'
    htmlAcc.append(s)
    s = '<div id="fundingListDiv">'
    htmlAcc.append(s)
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
        htmlAcc.append(s)
    # end_for       
    s = '</div>'
    htmlAcc.append(s)
    s = '</div>'
    htmlAcc.append(s)
# end_def gen_ex2_funding_div()

# This writes the HTML for the entire <body> of Exhibit 2, including:
#   the opening <body> tag
#   initial content (e.g., project name, etc.)
#   the div for the "Direct Salary and Overhead" line
#   the div for the salary cost table
#   the div forthe "Other Direct Costs" line
#   the div for funding source(s)
def gen_exhibit_2_body(htmlAcc, xlsInfo):
    s = '<body style="text-align:center;margin:0pt;padding:0pt;">'
    htmlAcc.append(s)
    s = '<div id="exhibit2">'
    htmlAcc.append(s)
    s = '<div class="exhibitPageLayoutDiv1"><div class="exhibitPageLayoutDiv2">'
    htmlAcc.append(s)
    s = '<h1>'
    htmlAcc.append(s)
    s = 'Exhibit 2<br>'
    htmlAcc.append(s)
    s = 'ESTIMATED COST<br>'
    htmlAcc.append(s)
    # Project name
    s = str(get_cell_contents(xlsInfo['ws'], xlsInfo['project_name_cell_row_ix'], xlsInfo['project_name_cell_col_ix']))
    s = s + '<br>'
    htmlAcc.append(s)
    s = '</h1>'
    htmlAcc.append(s)
    #
    gen_ex2_direct_salary_div(htmlAcc, xlsInfo)
    gen_ex2_salary_cost_table_div(htmlAcc, xlsInfo)
    gen_ex2_other_direct_costs_div(htmlAcc, xlsInfo)
    gen_ex2_total_direct_costs_div(htmlAcc, xlsInfo)
    gen_ex2_funding_div(htmlAcc, xlsInfo)
# end_def gen_exhibit_2_body()

def gen_exhibit_2(htmlAcc, xlsInfo):
    gen_exhibit_2_initial_boilerplate(htmlAcc)
    gen_exhibit_2_body(htmlAcc, xlsInfo)
    gen_exhibit_2_final_boilerplate(htmlAcc)
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
    htmlAcc = stringAccumulator()
    t1 = os.path.split(fullpath)
    in_dir = t1[0]
    in_fn = t1[1]
    in_fn_wo_suffix = os.path.splitext(in_fn)[0]
    ex_1_out_html_fn = in_dir + '\\' + in_fn_wo_suffix + '_Exhibit_1.html'
    ex_2_out_html_fn = in_dir + '\\' + in_fn_wo_suffix + '_Exhibit_2.html'
    
    # Collect 'navigation' information from input .xlsx file
    xlsInfo = initExcelFile(fullpath)
    if xlsInfo['errors'] == '':
        # Generate Exhibit 1 HTML, and save it to disk
        gen_exhibit_1(htmlAcc, xlsInfo)
        write_html_to_file(htmlAcc.get(), ex_1_out_html_fn)
        # Generate Exhibit 2 HTML, and save it to disk
        htmlAcc.re_init()
        gen_exhibit_2(htmlAcc, xlsInfo)
        write_html_to_file(htmlAcc.get(), ex_2_out_html_fn)
    else:
        print 'HTML generation aborted.\nErrors found when reading ' + fullpath + ':\n'
        print xlsInfo['errors']        
    # end_if
# end_def main()

# If this module has been invoked from the command line, the following statement
# ensures that the function "main" is called with the first parameter that was 
# passed on the command line, e.g.,
#     c:\Python27\python.exe -m workscope_exhibit_generator full_path_to_xlsx_file
if __name__== "__main__":
    main(sys.argv[1])
