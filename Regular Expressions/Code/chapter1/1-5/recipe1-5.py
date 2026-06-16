#!/usr/bin/python
import re
import sys

nargs = len(sys.argv)

if nargs > 1:
	mystr = sys.argv[1]

	r = re.compile(r'\b[bcm]at\b', re.M)
	if r.search(mystr):
		print 'I spy a bat, a cat, or a mat',
	else: 
		print 'I don\'t spy nuttin\', honey',

else:
	print 'Come again?',
