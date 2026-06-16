#!/usr/bin/perl -w
use strict;
my $digit = $ARGV[0] || die "No, no, no.  I came here for an argument."; 
$digit =~  s/(?:(?<=^)|(?<=[^\d]))(\d)(?=[^\d])/0$1/g;
print "'" . $digit . "'\n";
