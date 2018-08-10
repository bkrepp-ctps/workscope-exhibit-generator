# Python 'string accumulator' module
#
# NOTE: 
#   This module was written to run under Python 2.7.x
#
# Author: Benjamin Krepp
# Date: 10 August 2018

class stringAccumulator:
	accum = ''
	def append(self, x):
		self.accum += x
	def get(self):
		return self.accum
    def re_init(self):
        self.accum = ''
# end_class