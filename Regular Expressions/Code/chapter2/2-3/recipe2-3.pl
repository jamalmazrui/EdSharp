#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "Please supply a parameter";

if ( $mystr =~ s/^https?\:\/\/([^\/]+)\/?.*$/$1/ )
{
	print "The domain is '" . $mystr . "'\n";
}
else
{
	print "No hostname or anything like that in this string.\n";
}
