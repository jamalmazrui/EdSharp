#!/usr/bin/perl -w 
use strict; 

my $mystr = $ARGV[0] || die "You must supply a parameter!\n"; 

if ( $mystr =~ /^(?=.*[A-Z])(?=.*[a-z])(?=.*[0-9]).{7,15}$/ ) 
{ 
	print "Good password!\n"; 
} 
else 
{ 
	print "Dude!  That's way too easy.\n"; 
} 
