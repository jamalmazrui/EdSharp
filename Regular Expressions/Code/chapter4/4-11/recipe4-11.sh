#!/bin/bash

grep -E '^[-A-z0-9_.]+@([A-z0-9][-A-z0-9]+\.)+[A-z]{2,4}$' sample.txt
