#!/usr/bin/perl -w
use strict;

my $fn = $ARGV[0] || die "Please supply a filename!";
open ( FILE, $ARGV[0] ) || die "Cannot open file.";

while (<FILE>)
{ 
	my $line = $_;
	print $line if ( $line =~ /^(?:(?:(?!\/\/|\/\*).)*|(?:(?!\/\*).)*\/\*.*?\*\/(?:(?!\/\*).)*)(public|private)?\s+int\s+myvar\s+=/ );
}

close( FILE );
