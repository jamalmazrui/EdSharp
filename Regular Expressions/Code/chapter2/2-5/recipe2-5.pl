#!/usr/bin/perl -w

use strict;

my $path = $ARGV[0] || die "Whatchew talkin' 'bout, Willis?";

$path =~ s/\/cygdrive\/([a-z])\//\U$1:\\/;
$path =~ s/\//\\/g;

print $path;

exit 0;
