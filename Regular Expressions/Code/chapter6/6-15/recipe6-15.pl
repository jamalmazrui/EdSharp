#!/usr/bin/perl -w
use strict;

open ( FILE, $ARGV[0] ) || die "Cannot open file!";

while (<FILE>)
{ 
	my $line = $_;
	$line =~ s/\(\s*(\w+)\s*([=!]?=)\s*null\s*\)/( null $2 $1 )/g;
	print $line;
}

close( FILE );
