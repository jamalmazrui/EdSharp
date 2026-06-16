#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "You must supply a parameter!\n";
if ($mystr =~ /\b[bcm]at\b/)
{
	print "Found a word rhyming with hat\n";
}
else
{
	print "That's a negatory, big buddy.\n";
}
