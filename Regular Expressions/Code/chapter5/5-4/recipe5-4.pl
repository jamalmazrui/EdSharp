#!/usr/bin/perl -w
use strict;

my $html = $ARGV[0];

$html =~  s/(?:(?<=\<)|(?<=\<\/))([^\/>]+)(?:\sstyle=\"[^\"]+?\")([^\/>]*)(?=\/?\>| )/$1$2/g;
print "'" . $html . "'\n";
