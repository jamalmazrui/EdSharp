#!/usr/bin/perl -w
use strict;

my $input = $ARGV[0];

if ( $input =~  /^[A-z]+$/ )
{
	print "Thanks for the input:  '" . $input . "'\n";
} else {
	print "Alpha only, please!\n";
}
