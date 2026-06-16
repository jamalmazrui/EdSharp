#!/bin/bash

sed 's/\(\/\?>\)[[:space:]]\+\(<\/\?\)/\1\2/g' sample.html
