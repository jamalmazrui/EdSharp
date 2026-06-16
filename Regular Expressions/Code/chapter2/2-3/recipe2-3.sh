#!/bin/sh

echo $1 | sed 's/^https\?:\/\/\([^\/]\+\)\/\?.*$/\1/' 
