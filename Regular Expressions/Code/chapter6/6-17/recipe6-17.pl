#!/usr/bin/perl -w
use strict;

open ( FILE, $ARGV[0] ) || die "Cannot open file!";

while (<FILE>)
{ 
	my $line = $_;
	print $line if ( $line =~ /^[ >]*Subject:/ );
}

close( FILE );
