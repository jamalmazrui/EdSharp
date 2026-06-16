#!/usr/bin/perl -w
use strict;

my $username = $ARGV[0];

if ( $username =~ s/^([^@]+)(@.*)$/$1/ )
{
	print "$username\n";
} else {
	print "I couldn't determine your hostname...\n";
}

