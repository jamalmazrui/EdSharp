#!/usr/bin/perl -w
use strict;

my $date = $ARGV[0];

if ( $date =~ /^\d{5}(?:-\d{4})?$/ )
{
	print "Looks good to me...\n";
} else {
	print "I need a VALID postal code!\n";
}

