#!/usr/bin/perl -w
use strict;

my $date = $ARGV[0] || die "Please supply a parameter!";

$date =~  s/^(\d{1,2})[-\/.]?(\d{1,2})[-\/.]?((?:\d{2}|\d{4}))$/$1-$2-$3/;
print "'" . $date . "'\n";
