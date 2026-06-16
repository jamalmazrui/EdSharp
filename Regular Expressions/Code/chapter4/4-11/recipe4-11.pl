#!/usr/bin/perl -w
use strict;

my $input = $ARGV[0];

if ( $input =~  /^[-\w.]+@([A-z0-9][-A-z0-9]+\.)+[A-z]{2,4}$/ )
{
	print "Email address '" . $input . "' is valid\n";
} else {
	print "Please enter a valid e-mail address!\n";
}
