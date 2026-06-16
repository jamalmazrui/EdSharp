#!/usr/bin/python
import re
from os import popen

output = popen( 'du -hs /Users/myusername/*' )
regex = re.compile( r'^(?:(?:[2-9][0-9]{2}M)|(?:[0-9.]+G))\s+(.*)$' )

lines = output.readlines()

for line in lines:
	if regex.match( line ):
		formatted = regex.sub( r'You have more than 200MB in directory: \1', line )
		print formatted,

output.close()
