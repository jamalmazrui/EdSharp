#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "Please supply a parameter";
$mystr =~ s/^(https?:\/\/)(?:((25[0-5]|2[0-4][0-9]|1?[0-9]{2}|[0-9])\.){3}(25[0-5]|2[0-4][0-9]|1?[0-9]{2}|[0-9]))/$1host.example.com/;
print "'" . $mystr  . "'\n";
