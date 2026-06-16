#!/usr/bin/perl -w
use strict;

open ( FILE, $ARGV[0] ) || die "Cannot open file!";

while (<FILE>)
{ 
	my $line = $_;
	if ( $line =~ s/^([\d.]*?)\s.*GET\s(\/\S*)\sHTTP\/\d\.\d\"\s404\s\d{3}$/$1:  $2/g )
	{
		print $line;
	}
}

close( FILE );
