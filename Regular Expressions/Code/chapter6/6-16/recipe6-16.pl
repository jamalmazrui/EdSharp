#!/usr/bin/perl -w
use strict;

open ( FILE, $ARGV[0] ) || die "Cannot open file!";

while (<FILE>)
{ 
	my $line = $_;
	if ( $line =~ s/\r//g )
	{
		print $line;
	}
}

close( FILE );
