#!/usr/bin/perl -w
use strict;

my $mystr = $ARGV[0] || die "You must supply a parameter!\n";
if ( $mystr =~ /\b[Ss]+\s*[Pp]+\s*[Aa4]+\s*[Mm]*\b/ )
{
	print "They found me.  I don't know how, but they found me.\n";
}
