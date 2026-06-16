#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "Please supply a parameter";

$mystr =~ s/\b(\w+@[\w.]+)\b/<a href="mailto:$1">$1<\/a>/;
print $mystr . "\n";
