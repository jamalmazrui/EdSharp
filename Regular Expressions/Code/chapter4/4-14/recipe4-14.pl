#!/usr/bin/perl -w
use strict;

my $input = $ARGV[0] || die "Please supply a valid social security number.";

if ( $input =~  /^(00[1-9]|0[1-9]\d|[1-5]\d{2}|6[0-5]\d|66[0-5]|66[7-9]|6[7-8]\d|690|7[0-2]\d|73[0-3]|750|76[4-9]|77[0-2])-(?!00)\d{2}-(?!0000)\d{4}$/ )
{
	print "The number you entered is valid!\n";
} 
else 
{
	print "Please enter a valid number!\n";
}
