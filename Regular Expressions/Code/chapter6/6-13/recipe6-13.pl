#!/usr/bin/perl -w
use strict;

open ( FILE, $ARGV[0] ) || die "Cannot open file!";

while (<FILE>)
{ 
	my $line = $_;
	$line =~ s/(CREATE\s+PROCEDURE\s+)(?!\[?\w+\]?\.)(\w+)/$1\[dbo\].\[$2\]/g;
	print $line;
}

close( FILE );
