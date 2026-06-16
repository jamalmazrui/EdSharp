#!/usr/bin/perl -w
use strict;

my $fn = $ARGV[0] || die "Please supply a filename!";
open ( FILE, $ARGV[0] ) || die "Cannot open file.";

while (<FILE>)
{ 
	my $line = $_;
	$line =~ s/^\/\/ //g;
	print $line;
}

close( FILE );
