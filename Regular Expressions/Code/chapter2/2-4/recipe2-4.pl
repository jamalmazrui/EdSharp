#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "Please supply a parameter";

# find strings that start with www and append http://
$mystr =~ s/\b(www\.\S+)\b/http:\/\/$1/g;
print $mystr . "\n";
