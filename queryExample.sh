#!/bin/bash

# * Example query for PyMedWrapper

./pubmed.py \
    ~/Desktop/myPyMedWrapperTestOutput.xlsx \
    --tool "myCLIpubmedQuery" \
    --email "my.email@address.com" \
    --terms Alzheimer brain MRI cerebellum \
    --maxResults 200 \
    --pubSinceYear 2010

exit

