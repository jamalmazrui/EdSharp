#!/usr/bin/python
import re
from os import popen

output = popen( 'ps aux' )
regex = re.compile( r'^(root)\s+(\d+)\s+.*\s+(\/.*syslogd)(?:\s+-?\w+)*$' )

lines = output.readlines()
for line in lines:
	if regex.match( line ):
		formatted = regex.sub( r'\2:\3 (\1)', line )
		print formatted,

output.close()
