#!/usr/bin/perl -w
use strict;

my $number = $ARGV[0];

if ( $number =~ s/^\+(\d{1,3})([- ]?.*)$/$1/ )
{
	print "Your dialing code is '". $number ."'\n";
} else {
	print "I didn't find a dialing code.\n";
}

