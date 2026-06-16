#!/usr/bin/perl -w
use strict;

my $input = $ARGV[0] || die "These aren't the droids you're looking for.";

if ( $input =~  /^(?:1[- .]?)?\(?\d{3}\)?[- .]?\d{3}[- .]?\d{4}$/ )
{
	print "Phone number '" . $input . "' is valid\n";
} else {
	print "Please enter a valid phone number!\n";
}
