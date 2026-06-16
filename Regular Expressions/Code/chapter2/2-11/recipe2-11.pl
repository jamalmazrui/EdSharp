#!/usr/bin/perl -w
use strict;

opendir( CURDIR, '.' ) || die "The ship can't take much more of this, captain!\n";

my @files = readdir CURDIR;
closedir CURDIR;

foreach my $file ( @files )
{
	if ( $file =~ /^[-A-z0-9._]+\.(te?xt|TXT)$/ )
	{
		print "$file\n";
	}
}
