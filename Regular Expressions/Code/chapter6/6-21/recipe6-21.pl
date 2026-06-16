#!/usr/bin/perl -w
use strict;

open ( FILE, $ARGV[0] ) || die "Cannot open file!";

while (<FILE>)
{ 
	my $line = $_;
	# Jul 25 10:30:09 rhes-pe2400-01 pppd[1396]: local  IP address 209.98.170.206
	if ( $line =~ s/^.+\s+pppd\[\d+\]:\s+local\s+IP\s+address\s+([\d.]+)$/$1/ )
	{
		print $line;
	}
}

close( FILE );
