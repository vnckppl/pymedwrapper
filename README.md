# PyMed Wrapper #

This script is a wrapper for [pymed](https://github.com/gijswobben/pymed). 

It allows quick batch searching Pubmed and storing results in an Excel sheet.

**Installation**

pip install -r requirements.txt 

**Usage**
```
usage: pubmed.py [-h] --tool TOOL --email EMAIL
                 [--author1 AUTHOR1 [AUTHOR1 ...]] [--authors AUTHORS]
                 [--title TITLE [TITLE ...]] [--terms TERMS [TERMS ...]]
                 [--userquery USERQUERY [USERQUERY ...]]
                 [--maxResults MAXRESULTS] [--pubSinceYear PUBSINCEYEAR]
                 [--pubSinceLast PUBSINCELAST]
                 oFile

This script takes in a list of search terms for Pubmed and then uses the pymed
library to connect to the Pubmed API to obtain search results.

positional arguments:
  oFile                 Output file name. Path must exist.

optional arguments:
  -h, --help            show this help message and exit
  --tool TOOL           Your project name. This information is passed onto the
                        Pubmed API.
  --email EMAIL         Your email address. This information is passed onto
                        the Pubmed API.
  --author1 AUTHOR1 [AUTHOR1 ...]
                        First author. Format: Last#First or
                        Last#FirstInitials.
  --authors AUTHORS     Authors. Format: Last#First or Last#FirstInitials.
  --title TITLE [TITLE ...]
                        Title search words.
  --terms TERMS [TERMS ...]
                        List of search terms without Pubmed labels such as
                        [auth] or [ti].
  --userquery USERQUERY [USERQUERY ...]
                        List of search terms that can include Pubmed labels
                        such as [auth] and [ti].
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
```shell
./pubmed.py \
    ~/Desktop/myPyMedWrapperTestOutput_Example1.xlsx \
    --tool "myCLIpubmedQuery" \
    --email "my.email@address.com" \
    --terms Alzheimer brain MRI cerebellum \
    --maxResults 200 \
    --pubSinceYear 2010
```

```shell
./pubmed.py \
    ~/Desktop/myPyMedWrapperTestOutput_Example2.xlsx \
    --tool "myCLIpubmedQuery" \
    --email "my.email@address.com" \
    --authors 'Verghese#J Leeuw' \
    --title pain older \
    --maxResults 10 \
    --pubSinceYear 2000
```

**Known Bugs**
Searching for an author with 'Last Full First Name' (e.g., Doe John) results may not show up. In this case, use 'Last First Initial(s)' (e.g., Doe J).
