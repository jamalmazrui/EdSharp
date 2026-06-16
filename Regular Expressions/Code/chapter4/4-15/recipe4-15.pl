#!/usr/bin/perl -w
use strict;

my $ccnumber = $ARGV[0] || die "No parameter given.";

if ( $ccnumber =~ /^(?:\d{4}[- ]?){4}$/ )
{
	print "Looks good to me...\n";
} else {
	print "You can't buy anything with that...\n";
}

