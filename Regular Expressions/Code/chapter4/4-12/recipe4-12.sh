#!/bin/bash

grep -E '^((http|ftp)s?)://([A-z0-9][-A-z0-9]+\.)+[A-z]{2,4}$' sample.txt 
