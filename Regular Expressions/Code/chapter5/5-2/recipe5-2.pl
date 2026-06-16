#!/usr/bin/perl -w
use strict;

my $html = $ARGV[0];

$html =~  s/(?<=[:;\w{])\s+(?=[}\w;:])//g;
print "'" . $html . "'\n";
