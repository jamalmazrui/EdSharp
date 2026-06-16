#!/usr/bin/python
import re
import sys

nargs = len(sys.argv)

if nargs > 1:
	mystr = sys.argv[1]
	r = re.compile(r'\t', re.M)

	returnstr = r.sub( ',', open( mystr ).read( ) )

	print returnstr,

else:
	print 'Come again?',
