#!/usr/bin/perl -w
use strict;

my $largenumber = $ARGV[0] || die "Supply a parameter!";

$largenumber =~ s/(?<=\d)(?=(\d{3})+(?!\d))/,/g;

print "'" . $largenumber . "'\n";
