#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "Please supply a parameter";

$mystr =~ s/^(https?:\/\/(?:[A-z0-9][-A-z0-9]+\.)+[A-z]{2,6}(?::[0-9]+)?\/[-\w.\/]*)(?:\?[-\w:@=&\$.+!*'()"]+)$/$1?new=parameter/g;
print "'" . $mystr  . "'\n";
