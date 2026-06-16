#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "You must supply a parameter!\n";

if ( $mystr =~ /\bword$/ )
{
	print "The line ends in word.\n";
}
else
{
	print "I see no line ending in word.\n";
}
