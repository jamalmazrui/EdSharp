#!/usr/bin/perl -w
use strict;

my $date = $ARGV[0];

if ( $date =~ /^(?:0?[1-9]|1[0-2])\/(?:0?[1-9]|[1-2]\d|3[0-1])\/(?:\d{4})$/ )
{
	print "Looks good to me...\n";
} else {
	print "Please enter a VALID date!\n";
}

