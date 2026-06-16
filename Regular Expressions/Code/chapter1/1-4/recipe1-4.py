#!/usr/bin/python
import re
import sys

nargs = len(sys.argv)

if nargs > 1:
	mystr = sys.argv[1]

	r = re.compile(r'Joh?n(athan)? Doe', re.M)
	if r.search(mystr):
		print 'Here\'s Johnny!',
	else: 
		print 'Who?',

else:
	print 'I came here for an argument!',
