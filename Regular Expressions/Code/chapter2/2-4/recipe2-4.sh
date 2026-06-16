#!/bin/sh

sed 's/\(^\|[[:space:]]\)\(www\..*\)\([[:space:]]\|\$\)/\1http:\/\/\2\3/g' sample.txt
