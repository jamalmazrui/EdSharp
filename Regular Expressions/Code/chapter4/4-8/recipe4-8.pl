#!/usr/bin/perl -w
use strict;

my $input = $ARGV[0] || die "Parameter required.";

if ( $input =~  /^\$\d+\.\d\d$/ )
{
	print "Thanks for the input:  '" . $input . "'\n";
} else {
	print "How much?\n";
}
