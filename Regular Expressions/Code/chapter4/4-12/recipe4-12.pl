#!/usr/bin/perl -w
use strict;

my $input = $ARGV[0];

if ( $input =~  /^(?:(?:http|ftp)s?):\/\/(?:[A-z0-9][-A-z0-9]+\.)+[A-z]{2,4}$/ )
{
	print "URL '" . $input . "' is valid\n";
} else {
	print "Please enter a valid URL!\n";
}
