#!/usr/bin/perl -w
use strict;

my $negativenumber = $ARGV[0];

$negativenumber =~  s/^-([\d,.]+)$/($1)/g;
print "'" . $negativenumber . "'\n";
