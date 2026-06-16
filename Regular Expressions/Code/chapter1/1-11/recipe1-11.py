#!/usr/bin/python
import re
import sys

nargs = len(sys.argv)

if nargs > 1:

	mystr = sys.argv[1]

	r = re.compile(r'\b(\w+)\s+\1\b', re.M )

	if r.match( open( mystr ).read( ) ) :
		print "Found double words",
	else:
		print "No match here.",

else:
	print "I came here for an argument.",
