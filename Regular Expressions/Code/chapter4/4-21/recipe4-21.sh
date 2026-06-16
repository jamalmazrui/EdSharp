#!/bin/bash

sed -r 's/^(.*)[[:space:]]([A-z]+)$/\2, \1/g' sample.txt
