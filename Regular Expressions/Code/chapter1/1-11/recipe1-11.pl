#!/usr/bin/perl -w
use strict;

# my $mystr = $ARGV[0] || die "You must supply a parameter!\n";

my $mystr = "my word\nword";

print $mystr;

if ( $mystr =~ /\b(\w+)[\s\n]\1\b/mi )
{
	print "I think I'm seeing double.\n";
}
else
{
	print "Looks okay to me.\n";
}
