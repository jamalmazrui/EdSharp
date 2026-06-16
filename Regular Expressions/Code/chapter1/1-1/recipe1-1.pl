#!/usr/bin/perl -w
use strict;

open( FILE, $ARGV[0] ) || die "Cannot open file!";

my $i = 0;

while ( <FILE> )
{
	$i++;
	next unless /^\s*$/;
	print "Found blank line at line $i\n";
}

close( FILE );
