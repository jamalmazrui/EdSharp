#!/usr/bin/perl -w
use strict;

my $time = $ARGV[0] || die "No parameter given.";

if ( $time =~ /^(?:0?[1-9]|1[0-2]):(?:[0-5][0-9])(?::[0-5][0-9])? [PA]\.?M\.?$/ )
{
	print "Looks good to me...\n";
} else {
	print "What time is it?\n";
}

