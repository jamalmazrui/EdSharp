#!/usr/bin/perl -w
use strict;

open( FILE, $ARGV[0] ) || die "Cannot open file!";

while ( <FILE> )
{
	# print the filtered line
	my $line = $_;
	$line =~ s/\b([a-z])(\w+)\b/\u$1$2/g;
	print $line;
}

close( FILE );
