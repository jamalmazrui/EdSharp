#!/usr/bin/perl -w
use strict;

my $html = $ARGV[0];

$html =~  s/(?:(?<=\>)|(?<=\/\>))(\s+)(?=\<\/?)//g;
print "'" . $html . "'\n";
