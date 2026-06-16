#!/bin/sh

sed 's/\<\(http:\/\/[^[:space:]]\+[^-",.:; ]\)\>/<a href=\"\1\">\1<\/a>/g' sample.txt
