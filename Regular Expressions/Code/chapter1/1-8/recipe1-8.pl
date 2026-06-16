#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "You must supply a parameter!\n";

$mystr =~ s/\t/,/g;

print $mystr . "\n";
