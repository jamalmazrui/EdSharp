#!/usr/bin/perl -w
use strict;

my $input = $ARGV[0];

if ( $input =~  /^(?:P\.?O\.?\s)?(?:BOX)\b/i )
{
	print "The address provided looks like a post office box\n";
} else {
	print "Shipping to '" . $input . "'\n";
}
