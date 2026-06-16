#!/usr/bin/perl -w
use strict;

my $html = $ARGV[0];

if ( $html =~ /\<[^\/>]+\ssummary=['\"][^'\"]*?['\"][^\/>]*?\/?\>/ )
{
	print "Found match in '" . $html . "'\n";
}
else
{
	print "Didn't find match in string '" . $html . "'\n";
}
