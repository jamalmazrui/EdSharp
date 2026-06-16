#!/usr/bin/perl -w
use strict;

open( FILE, $ARGV[0] ) || die "Cannot open file!";

while ( <FILE> )
{
	my $line = $_;
	$line =~ s/\n/, /mg;
	print $line;
}

print "\n";

close( FILE );
