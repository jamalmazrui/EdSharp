#!/bin/bash

sed 's/\(<\/\?[^\/>]\+\)[[:space:]]style=[\'"][^\'"]*[\'"]\(.*\/\?>\| \)/\1\2/g' sample.html 
