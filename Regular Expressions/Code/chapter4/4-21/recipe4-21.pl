#!/usr/bin/perl -w
use strict;

my $name = $ARGV[0];

$name =~ s/^(.*)\s([A-z']+)$/$2, $1/;
print "'" . $name ."'\n";
