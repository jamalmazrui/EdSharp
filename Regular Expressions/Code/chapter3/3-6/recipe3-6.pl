#!/usr/bin/perl -w
use strict;

open( FILE, $ARGV[0] ) || die "Cannot open file!";

my $i = 0;

while ( <FILE> )
{	
	my $field = $_;
	$field =~ s/^(?:[^\t]+\t)([^\t]+)(?:\t.*)$/$1/;
	print $field;
}

close( FILE );
