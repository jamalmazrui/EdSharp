#!/usr/bin/python
import re
import sys

nargs = len(sys.argv)

if nargs > 1:

	myfile = sys.argv[1]
	output = open( myfile )
	regex = re.compile( r'^([\d.]*?)\s.*GET\s(\/\S*)\sHTTP\/\d\.\d\"\s404\s\d{3}$' )

	lines = output.readlines()
	for line in lines:
		if regex.match( line ):
			formatted = regex.sub( r'\1:  \2', line )
			print formatted,

	output.close()

else:
	print 'Please supply a parameter!',
