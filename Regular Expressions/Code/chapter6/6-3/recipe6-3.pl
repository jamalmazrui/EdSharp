#!/usr/bin/perl -w
use strict;

open ( FILE, $ARGV[0] ) || die "Cannot open file!";

while (<FILE>)
{ 
	my $line = $_;
	$line =~ s/MyMethod\s*\(\s*(\"?\w+\"?),\s*(\"?\w+\"?)\s*\);/MyMethod( $2, $1 );/g;
	print $line;
}

close( FILE );
