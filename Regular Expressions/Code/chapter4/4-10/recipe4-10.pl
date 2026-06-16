#!/usr/bin/perl -w
use strict;

my $input = $ARGV[0];

if ( $input =~  /^(([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5]).){3}([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$/ ) 
{ 
	print "IP '" . $input . "' is valid\n";
} else {
	print "Please enter a valid IP address!\n";
}
