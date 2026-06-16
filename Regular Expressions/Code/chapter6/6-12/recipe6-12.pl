#!/usr/bin/perl -w
use strict;

while (<STDIN>)
{ 
	print $_ if ( s/^(?:(?:[2-9][0-9]{2}M)|(?:[0-9.]+G))\s+(.*)$/Directory larger than 200M: $1/g );
}
