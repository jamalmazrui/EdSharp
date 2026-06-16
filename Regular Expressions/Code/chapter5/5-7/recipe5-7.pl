#!/usr/bin/perl -w
use strict;

my $xml = $ARGV[0];

if ( $xml =~ /\<([^> \/]+)[^>]*?\>(?:.*?)\<\/\1\>/ )
{
	print "Looks good!  Original input is:  '" . $xml . "'\n";
}
else
{
	print "Suspect unclosed XML tag in '" . $xml . "'\n";
}
