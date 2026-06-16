#!/usr/bin/perl -w
use strict;

open( FILE, $ARGV[0] ) || die "Cannot open file!";

my $i = 0; 

while ( <FILE> )
{
  $i++;
	next unless /\bword\b/;
	print "Found word at line $i\n";
}

close( FILE );
