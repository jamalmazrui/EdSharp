#!/usr/bin/perl -w
use strict;

open( FILE, $ARGV[0] ) || die "Cannot open file!";

my $i = 0;

while ( <FILE> )
{
	$i++;
	next unless /^[^\t]+(\t[^\t]+){4}$/;
	print $_;
}

close( FILE );
