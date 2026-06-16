#!/usr/bin/perl -w
use strict;

my $phonenumber = $ARGV[0];

$phonenumber =~  s/^\(?(\d{3})\)?[- .]?(\d{3})[- .]?(\d{4})$/($1) $2-$3/;
print "'" . $phonenumber . "'\n";
