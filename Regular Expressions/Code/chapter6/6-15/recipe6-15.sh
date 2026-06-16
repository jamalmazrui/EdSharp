#!/bin/bash

sed 's/( *\([a-z0-9_]\+\) *\([=!]\?=\) *null *)/( null \2 \1 )/g' sample.txt 
