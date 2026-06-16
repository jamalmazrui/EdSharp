#!/usr/bin/perl -w
use strict;

open( FILE, $ARGV[0] ) || die "Cannot open file!";

while ( <FILE> )
{
	# print the filtered line
	my $line = $_;
	$line =~ s/\b(bleep|beep|blankity)\b/\$\@#!/gi;
	print $line;
}

close( FILE );
