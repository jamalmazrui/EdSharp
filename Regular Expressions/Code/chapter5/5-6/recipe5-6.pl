#!/usr/bin/perl -w

use strict;

my $fn = $ARGV[0] || die "Please supply a parameter!\n";

open( FILE, $fn ) || die "Cannot open file!\n";

while ( <FILE> )
{
	my $line = reverse $_;
	$line =~ s/>(?![^><]+?\/?<)/;tl&/g;
	$line = reverse $line;
	print $line;
}

close( FILE );
