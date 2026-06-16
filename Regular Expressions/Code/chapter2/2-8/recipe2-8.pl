#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "Please supply a parameter";

if ( $mystr =~ s/^.+\.(\w{2,4})$/$1/ )
{
	print "'" . $mystr  . "'\n";
}
else
{
	print "No file extension found!\n";
}
