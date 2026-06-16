#!/usr/bin/perl -w
use strict;

open( FILE, $ARGV[0] ) || die "Cannot open file!";

my $i = 0;

while ( <FILE> )
{	
	my $line = $_;
	$line =~ s/,(?=(?:[^"]*$)|(?:[^"]*"[^"]*"[^"]*)*$)/\t/go;
	print $line;
}

close( FILE );
