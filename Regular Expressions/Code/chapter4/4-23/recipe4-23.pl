#!/usr/bin/perl -w
use strict;

my $str = $ARGV[0];

if ( $str =~ /^(t(rue)?|y(es)?)$/i )
{
	print "Yes!!\n";
}
else
{
	print "Uh, no.\n";
}

exit 0;
