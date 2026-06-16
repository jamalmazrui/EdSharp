#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "Please supply a parameter";

$mystr =~ s/\b(http:\/\/\S+[^-.,\"\';: ])\b/<a href="$1">$1<\/a>/g;
print $mystr . "\n";
