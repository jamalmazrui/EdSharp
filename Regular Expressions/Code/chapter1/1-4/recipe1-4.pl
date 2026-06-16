#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "You must supply a parameter!\n";
if ( $mystr =~ /Joh?n(athan)? Doe/) 
{
	print "Hello!  I've been looking for you!\n" 
}
