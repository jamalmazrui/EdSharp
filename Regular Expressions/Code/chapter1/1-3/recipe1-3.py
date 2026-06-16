import re

r = re.compile(r'\s+((moo)|(oink))\s+')
if r.search( open( 'sample.txt' ).read( ) ) : 
	print "I spy a cow or pig.",
else: 
	print "Ah, there's no cow or pig here.",
