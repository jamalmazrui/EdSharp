#!/usr/bin/python
import re
import sys

nargs = len(sys.argv)

if nargs > 1:

	mystr = sys.argv[1]

	r = re.compile( r',\s*', re.M )

	newstr = r.sub( ',\n', open( mystr ).read( ) )

	print newstr,

else:
	print 'Filename?  Anyone?  Anyone?',
