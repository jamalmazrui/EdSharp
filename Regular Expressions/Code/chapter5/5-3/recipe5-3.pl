#!/usr/bin/perl -w
use strict;

open ( FILE, $ARGV[0] ) || die "Cannot open file!";

while (<FILE>)
{ 
	my $line = $_;
	$line =~ s/(?:(?<=\<)|(?<=\<\/))([^\/ >]+)(?=\/?\>| )/\L$1/g;
	print $line;
}

close( FILE );
