#!/usr/bin/python
import re
import sys

nargs = len(sys.argv)

if nargs > 1:

	mystr = sys.argv[1]

	r = re.compile( r'\n', re.M )

	newstr = r.sub( ', ', open( mystr ).read( ) )

	print newstr,

else:
	print 'Filename?  Anyone?  Anyone?',
