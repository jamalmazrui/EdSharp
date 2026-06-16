#!/usr/bin/python
import re
import sys

nargs = len(sys.argv)

if nargs > 1:

	myfile = sys.argv[1]
	output = open( myfile )

	regex = re.compile( r'\r' )

	lines = output.readlines()
	for line in lines:
		formatted = regex.sub( r'', line )
		print formatted,

	output.close()

else:
	print 'I\'ve seen Tron eight times',
