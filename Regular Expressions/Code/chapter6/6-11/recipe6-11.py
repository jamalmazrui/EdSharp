#!/usr/bin/python
import re
from os import popen

output = popen( 'netstat -p tcp' )
regex = re.compile( r'^(?:.*\s+)([-\w.:]+)\s+([-\w.:]+)(?:\s+ESTABLISHED)$' )

lines = output.readlines()

for line in lines:
	if regex.match( line ):
		formatted = regex.sub( r'Found established connection:  \1 -> \2', line )
		print formatted,

output.close()
