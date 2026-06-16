#!/usr/bin/perl -w
use strict;

open ( FILE, $ARGV[0] ) || die "Cannot open file!";

while (<FILE>)
{ 
	my $line = $_;
	$line =~ s/\bMyMethod\s*\(/MyNewMethod(/g;
	print $line;
}

close( FILE );
