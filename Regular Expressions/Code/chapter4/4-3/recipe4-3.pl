#!/usr/bin/perl -w
use strict;

my $date = $ARGV[0] || die "What!?!";

if ( $date =~ /^(?:(?:(?:0[1-9]|[12][0-9]|30)-(?:Sep|Apr|Jun|Nov))|(?:(?:0[1-9]|[12][0-9])-Feb)|(?:(?:0[1-9]|[12][0-9]|3[01])-(?:Jan|Mar|May|Jul|Aug|Oct|Nov|Dec)))-\d{4}$/ ) {
	print "Found valid date:  " . $date . "\n";
} else { 
	print "Found invalid date:  " . $date . "\n";
}

exit 0;
