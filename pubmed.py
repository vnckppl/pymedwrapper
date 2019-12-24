#!/usr/bin/env python3

# * Pubmed Queries via Python Pubmed API Wrapper
# 2019-12-23
# Vincent Koppelmans

# * Background
# PyMed: https://github.com/gijswobben/pymed
# Script based on: https://github.com/gijswobben/pymed/blob/master/examples/advanced_search/main.py

# * Libraries
import argparse
from pymed import PubMed
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
import xlsxwriter

# * Arguments
if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        description='This script takes in a list of search terms '
        'for Pubmed and then uses the pymed library to connect '
        'to the Pubmed API to obtain search results.')
    # Positional arguments
    parser.add_argument('oFile',
                        help='Output file name. Path must exist.'
    )
    # Flags for information requested by Pubmed API
    parser.add_argument('--tool',
                        help='Your project name. This information is passed onto the Pubmed API.',
                        required=True
    )
    parser.add_argument('--email',
                        help='Your email address. This information is passed onto the Pubmed API.',
                        required=True
    )
    # All other flags used to build the query
    parser.add_argument('--terms',
                        nargs='+',
                        help='List of search terms'
    )
    parser.add_argument('--maxResults',
                        default=50,
                        help='Maximum number of results'
    )
    parser.add_argument('--pubSinceYear',
                        default=1980,
                        help='Only results that have been published since <year>. '
                        ' Default=1980'
    )
    parser.add_argument('--pubSinceLast',
                        help='Only results that have been published since the last <x> years '
                        '(default=not set). This takes precedence over --pubSinceYear.'
    )

args = parser.parse_args()

# * Class for Synonyms
class query(object):

    # * Store Flags
    def __init__(self):
        self.oFile = args.oFile
        self.terms = args.terms
        self.maxResults = args.maxResults
        self.psYear = args.pubSinceYear
        self.psLast = args.pubSinceLast
        self.email = args.yourEmail
        self.tool = args.projectName
                        
    # * Build Object
    # Create a PubMed object that GraphQL can use to query
    def buildQuery(self):
        # Build Object and send some info to PubMed by their request
        # Note that the parameters below are not required but kindly requested by PubMed Central
        # https://www.ncbi.nlm.nih.gov/pmc/tools/developers/
        self.pubmed = PubMed(tool=self.tool, email=self.email)

        # Create a GraphQL query in plain text
        # Example: query:
        # '(("2018/05/01"[Date - Create] : "3000"[Date - Create])) AND (Xiaoying Xian[Author] OR diabetes)'
        self.query = ""
        for item in self.terms:
            self.query=self.query+' AND '+item

        # Calculate what the start dat is for articles to be included based on user settings
        if self.psLast is not None:
            # Only include articles published in the last <x> years
            self.dYa = datetime.now() - relativedelta(years=int(self.psLast))
            self.dayYearsAgo = str(self.dYa).split(' ')[0].replace('-','/')
            self.dYaQuery = '('+self.dayYearsAgo+'[Date - Create] : "3000"[Date - Create])'

        else:
            self.dYaQuery = '("'+self.psYear+'/01/01"[Date - Create] : "3000"[Date - Create])'
        self.query=self.dYaQuery + self.query

    def runQuery(self):    
        # Execute the query against the API
        self.results = self.pubmed.query(self.query, max_results=int(self.maxResults)+1)

        # Make dictionary to store data
        self.output={}

        # Loop over the retrieved articles
        self.nResults=0
        for result in self.results:
            self.nResults=self.nResults+1

        # Check if there are more than 200 results
        if self.nResults > int(self.maxResults):
            # Show warning
            print('More than '+str(self.maxResults)+' results found')
        elif self.nResults == 0:
            # Show warning
            print('No results found')            
        else:
            # Print number of results
            print(str(self.nResults)+' result(s) obtained.')
            
            # Loop over the retrieved articles
            self.results = self.pubmed.query(self.query, max_results=int(self.maxResults))
            for article in self.results:

                # Extract and format information from the article
                article_id = article.pubmed_id.split()[0]
                title = article.title
                authors = article.authors
                # if article.keywords:
                #     if None in article.keywords:
                #         article.keywords.remove(None)
                #     keywords = '", "'.join(article.keywords)
                publication_date = article.publication_date
                abstract = article.abstract
                if hasattr(article, 'journal'):
                    journal = article.journal
                else:
                    journal = 'NA'

                # Reshape author list
                authorString=''
                for author in authors:
                    last=author['lastname']
                    first=author['firstname']
                    if last is None:
                        last='NA'
                    if first is None:
                        first='NA'
                    authorString=authorString+' '+last+', '+first+';'

                # Add results to the dictionary
                self.output[article_id] = [article_id, title, authorString, journal, publication_date, abstract]

            # Put data in a dataframe after extraction
            self.DF = pd.DataFrame.from_dict(self.output)
            self.DF = self.DF.T
            self.DF = self.DF.reset_index(drop=True) # Remove row names
            self.DF.columns = ["PMID","Title","Authors","Journal","PubDate","Abstract"]

            # Save to Excel
            self.writer = pd.ExcelWriter(self.oFile, engine='xlsxwriter')
            self.DF.to_excel(self.writer, sheet_name='PMquery', index=False)
            self.workbook=self.writer.book
            self.worksheet = self.writer.sheets['PMquery']

            # Formatting
            self.format = self.workbook.add_format({
                'text_wrap': True,
                'align': 'top'
            })
            self.worksheet.set_column('A:A', 9,  self.format)
            self.worksheet.set_column('B:C', 22, self.format)
            self.worksheet.set_column('D:E', 11, self.format)
            self.worksheet.set_column('F:F', 58, self.format)
            self.writer.save()

# * Run Query
myQuery=query()
myQuery.buildQuery()
myQuery.runQuery()
