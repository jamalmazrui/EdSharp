#!/usr/bin/perl -w
use strict;

open( FILE, $ARGV[0] ) || die "Cannot open file!";

my $i = 0; 

while ( <FILE> )
{
  $i++;
	next unless /\s+((moo)|(oink))\s+/;
	print "Found animal sound at line $i\n";
}

close( FILE );
