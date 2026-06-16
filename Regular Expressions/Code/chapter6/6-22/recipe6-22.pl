#!/usr/bin/perl -w
use strict;

my $fn = $ARGV[0] || die "Please supply a paramter\n";

open( FILE, $fn ) || die "Could not open file\n";

while ( <FILE> )
{ 
	print $_ if ( s/^([^=]+?)=(.*)$/Key: '$1'; Value: '$2'/ );
}

close( FILE );
