#!/bin/bash

sed 's/MyMethod *( *\(.*\), *\(.*\) *) *;/MyMethod( \2, \1 );/g' sample.txt
