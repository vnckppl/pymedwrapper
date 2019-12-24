# PyMed Wrapper #

This script is a wrapper for [pymed](https://github.com/gijswobben/pymed). 

It allows quick batch searching Pubmed and storing results in an Excel sheet.

**Installation**

pip install -r requirements.txt 

**Usage**
```
usage: pubmed.py [-h] --tool TOOL --email EMAIL [--terms TERMS [TERMS ...]]
                 [--maxResults MAXRESULTS] [--pubSinceYear PUBSINCEYEAR]
                 [--pubSinceLast PUBSINCELAST]
                 oFile

positional arguments:
  oFile                 Output file name. Path must exist.

optional arguments:
  -h, --help            show this help message and exit
  --tool TOOL           Your project name. This information is passed onto the
                        Pubmed API.
  --email EMAIL         Your email address. This information is passed onto
                        the Pubmed API.
  --terms TERMS [TERMS ...]
                        List of search terms
  --maxResults MAXRESULTS
                        Maximum number of results
  --pubSinceYear PUBSINCEYEAR
                        Only results that have been published since <year>.
                        Default=1980
  --pubSinceLast PUBSINCELAST
                        Only results that have been published since the last
                        <x> years (default=not set). This takes precedence
                        over --pubSinceYear.
```

**Examples**
``` sh
./pubmed.py \
    ~/Desktop/myPyMedWrapperTestOutput.xlsx \
    --tool "myCLIpubmedQuery" \
    --email "my.email@address.com" \
    --terms Alzheimer brain MRI cerebellum \
    --maxResults 200 \
    --pubSinceYear 2010
```

