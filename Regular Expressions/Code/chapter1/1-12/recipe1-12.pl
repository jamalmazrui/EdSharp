#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "You must supply a parameter!\n";

if ( $mystr =~ /^Word\b/ )
{
	print "In the beginning was the Word.\n";
}
else
{
	print "Nope.  Not here.\n";
}
