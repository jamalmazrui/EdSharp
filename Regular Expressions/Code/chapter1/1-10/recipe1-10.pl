#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "You must supply a parameter!\n";

if ( $mystr =~ /\b(\w+)\s\1\b/ )
{
	print "I think I'm seeing double.\n";
}
else
{
	print "Looks okay to me.\n";
}
