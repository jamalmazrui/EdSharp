#!/usr/bin/perl -w
use strict;

my $input = $ARGV[0];

if ( $input =~  /^.{0,15}$/ )
{
	print "Input '" . $input . "' is valid\n";
} else {
	print "No more than 15 characters!\n";
}
