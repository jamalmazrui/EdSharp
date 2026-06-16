#!/usr/bin/python
import re
import sys
from os import popen

nargs = len(sys.argv)

if nargs > 1:

	myfile = sys.argv[1]
	output = open( myfile )
	regex = re.compile( r'^[ >]*Subject:' )

	lines = output.readlines()
	for line in lines:
		if regex.match( line ):
			print line,

	output.close()

else:
	print 'Please supply a parameter!',
