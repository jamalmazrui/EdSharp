#!/usr/bin/perl -w
use strict;
my $line = $ARGV[0];
$line =~ s/,\s*/,\n/g;
print $line. "\n";
