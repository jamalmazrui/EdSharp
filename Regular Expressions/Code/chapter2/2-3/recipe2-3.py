#!/usr/bin/python
import re
import sys

nargs = len(sys.argv)

if nargs > 1:

	mystr = sys.argv[1]

	r = re.compile( r'^https?:\/\/([^\/]+)\/?.*$', re.I )

	hostname = r.sub( r'\1', mystr )

	print 'Hostname is: ' + hostname,

else:
	print 'Filename?  Anyone?  Anyone?',
