#!/usr/bin/python
import re
import sys
from os import popen

nargs = len(sys.argv)

if nargs > 1:

	myfile = sys.argv[1]
	output = open( myfile )
	regex = re.compile( r'^([A-z0-9][-A-z0-9]+)\s+(IN)\s+(A|CNAME)\s+([A-z0-9][-A-z0-9.]+\.|[\d.]+)$' )

	lines = output.readlines()
	for line in lines:
		formatted = regex.sub( r'\1\t\2\t\3\t\4', line )
		print formatted,

	output.close()

else:
	print 'Please supply a parameter!',
