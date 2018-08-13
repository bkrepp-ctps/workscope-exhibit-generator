# Python 'string accumulator' module
#
# NOTE: 
#   This module was written to run under Python 2.7.x
#
# Author: Benjamin Krepp
# Date: 10, 13 August 2018

class stringAccumulator:
    def __init__(self):
        self.accum = ''
    def append(self, x):
        self.accum += x
    def get(self):
        return self.accum
    def re_init(self):
        self.accum = ''
# end_class

######################################################################################
# I would really like to have managed the collection of HTML using the following 
# function, but have decided against this (at least for the time being) in order to 
# make the code easier to understand for people who are unfamiliar with closures in 
# general (and closures in Python 2.x in particular) and functional programming.
# If you're so interested, modifying the code to use "functional_stringAccumulator" 
# should be straightforward. I leave "functional_stringAccumulator" here as a brain 
# teaser for those who might enjoy the opportunity to work with functional code. 
# -- BK 7/27/2018, 8/13/2018
def functional_stringAccumulator():
   # N.B. dict required to hold local vars, due to Python 2.7 idosyncracy
    my_vars = {}
    my_vars['accum'] = ''
    def append(s):
        my_vars['accum'] += s
    # end_def
    def get():
        return my_vars['accum']
    # end_def
    def re_init():
        my_vars['accum'] = ''
    # end_def    
    retval = {}
    retval['append'] = append
    retval['get'] = get
    retval['re_init'] = re_init
    return retval
# end_def functional_stringAccumulator()