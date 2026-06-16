#!/usr/bin/perl -w
use strict;

opendir( CURDIR, '.' ) || die "The ship can't take much more of this, captain!\n";

my @files = readdir CURDIR;
closedir CURDIR;

foreach my $file ( @files )
{
	if ( $file =~ /log[-]200406(0[1-9]|2[0-9])\.log/ )
	{
		print "$file\n";
	}
}
