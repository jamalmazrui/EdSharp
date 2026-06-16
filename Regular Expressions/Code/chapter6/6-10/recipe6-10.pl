#!/usr/bin/perl -w
use strict;

while (<STDIN>)
{ 
	print $_ if ( s/^(root)\s+(\d+)\s+.*\s+(\/.*syslogd)(?:\s+-?\w+)*$/$2:$3 ($1)/ );
}
