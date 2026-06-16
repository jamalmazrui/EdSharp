#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "Please supply a parameter";

if ( $mystr =~ s/^[^?]*\?(.*)$/$1/ )
{
	print $mystr . "\n";
}
else
{
	print "No query string here.\n";
}
