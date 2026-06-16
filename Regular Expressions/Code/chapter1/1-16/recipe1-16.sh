#!/bin/sh

sed 's/\<\(bleep\|beep\|blankity\)/$@#!/g' sample.txt
