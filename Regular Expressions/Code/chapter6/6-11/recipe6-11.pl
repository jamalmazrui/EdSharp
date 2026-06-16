#!/usr/bin/perl -w
use strict;

while (<STDIN>)
{ 
	print $_ if ( s/^(?:.*\s+)([-\w.:]+)\s+([-\w.:]+)(?:\s+ESTABLISHED)$/Found established connection:  $1 -> $2/g );
}
