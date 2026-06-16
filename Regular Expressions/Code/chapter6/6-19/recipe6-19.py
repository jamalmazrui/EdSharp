#!/usr/bin/python
import re
from os import popen

output = popen( 'df' )
regex = re.compile( r'^.*\s(([5-9][0-9])|(100))%\s.*$' )

lines = output.readlines()

for line in lines:
	if regex.match( line ):
		print line,

output.close()
