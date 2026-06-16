#!/usr/bin/perl -w
use strict;

opendir( CURDIR, '.' ) || die "The ship can't take much more of this, captain!\n";

my @files = readdir CURDIR;
closedir CURDIR;

foreach my $file ( @files )
{
	if ( $file =~ /^[-A-z0-9_.]*\.(?:txt|TXT|text)$/ )
	{
		# rename the file	
		my $newfilename = $file;
		$newfilename =~ s/^([-A-z0-9_.].*)\.(?:te?xt|TXT)/$1\.txt/;
		rename $file, $newfilename || die "Could not rename file '" . 
			$file . "'\n";
	}
}
