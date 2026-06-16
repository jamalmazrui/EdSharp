#!/usr/bin/perl 
use strict;

while ( <STDIN> )
{ 
    print $_ if ( /\s(([5-9][0-9])|(100))\%\s/ );
}
