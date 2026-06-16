#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "Please supply a parameter";

# $mystr =~ s/ /%20/g;
$mystr =~ s/(^https?:\/\/(?:[A-z0-9][-A-z0-9]+\.)+[A-z]{2,4}(?::[0-9]+)?\/)(?:[-\w.\/]+)(\?[-\w:@=&\$.+!*'()"]+)$/$1different\/path\/file.php$2/g;
print "'" . $mystr  . "'\n";
