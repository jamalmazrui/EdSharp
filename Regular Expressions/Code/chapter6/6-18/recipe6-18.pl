#!/usr/bin/perl -w
use strict;

open ( FILE, $ARGV[0] ) || die "Cannot open file!";

while (<FILE>)
{ 
	my $line = $_;
	$line =~ s/^([A-z0-9][-A-z0-9]+)\s+(IN)\s+(A|CNAME)\s+([A-z0-9][-A-z0-9.]+\.|[\d.]+)$/$1\t$2\t$3\t$4/;
	print $line;
}

close( FILE );
