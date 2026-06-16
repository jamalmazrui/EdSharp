#!/usr/bin/python
import re
import sys

nargs = len(sys.argv)

if nargs > 1:

	myfile = sys.argv[1]
	output = open( myfile )
	regex = re.compile( r'^.+\s+pppd\[\d+\]:\s+local\s+IP\s+address\s+([\d.]+)$' )

	lines = output.readlines()
	for line in lines:
		if regex.match( line ):
			formatted = regex.sub( r'Local IP Address is:  \1', line )
			print formatted,

	output.close()

else:
	print 'Please supply a parameter!',
