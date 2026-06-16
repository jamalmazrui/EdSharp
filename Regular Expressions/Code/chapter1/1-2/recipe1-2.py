import re

r = re.compile(r'word',re.M )
if r.search( open( 'sample.txt' ).read( ) ) : 
	print "I finally found what I'm looking for." ,
else: 
	print "\"word\"s not here, man." ,
