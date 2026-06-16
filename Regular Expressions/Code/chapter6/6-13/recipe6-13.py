#!/usr/bin/python
import re
import sys

nargs = len(sys.argv)

if nargs > 1:

	myfile = sys.argv[1]
	output = open( myfile )
	regex = re.compile( r'(CREATE\s+PROCEDURE\s+)(?!\[?[A-Z]+\]?\.)(\w+)', re.I )
	lines = output.readlines()
	for line in lines:
		formatted = regex.sub( r'\1[dbo].[\2]', line )
		print formatted,

	output.close()

else:
	print 'Supply a parameter!',
