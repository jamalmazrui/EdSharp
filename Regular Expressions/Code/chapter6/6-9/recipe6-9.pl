#!/usr/bin/perl -w
use strict;

open ( FILE, $ARGV[0] ) || die "Cannot open file!";

while (<FILE>)
{ 
	my $line = $_;
	if ( $line =~ m/^(?:\/\*(?:(?!\*\/).)*|\/\/.*?)WORD/ )
	{
		print $line;
	}
}

close( FILE );
