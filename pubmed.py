#!/usr/bin/env python3

# * Pubmed Queries via Python Pubmed API Wrapper
# 2019-12-25
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
    parser.add_argument('--author1',
                        nargs='+',
                        help='First author. Format: Last#First or Last#FirstInitials.'
    )
    parser.add_argument('--authors',
                        help='Authors. Format: Last#First or Last#FirstInitials.'
    )
    parser.add_argument('--title',
                        nargs='+',
                        help='Title search words.'
    )
    parser.add_argument('--terms',
                        nargs='+',
                        help='List of search terms without Pubmed labels such as [auth] or [ti].'
    )
    parser.add_argument('--userquery',
                        nargs='+',
                        help='List of search terms that can include Pubmed labels such as [auth] and [ti].'
    )
    parser.add_argument('--maxResults',
                        default=50,
                        help='Maximum number of results'
    )
    parser.add_argument('--pubSinceYear',
                        default=str(1980),
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
        # Positional Arguments
        self.oFile = args.oFile
        # Flags for information requested by Pubmed API
        self.email = args.email
        self.tool = args.tool
        # All other flags used to build the query
        self.author1 = args.author1
        self.authors = args.authors
        self.title = args.title
        self.terms = args.terms
        self.userquery = args.userquery
        self.psYear = args.pubSinceYear
        self.psLast = args.pubSinceLast
        self.maxResults = args.maxResults
        
    # * Build Object
    # Create a PubMed object that GraphQL can use to query
    def buildQuery(self):
        # Build Object and send some info to PubMed by their request
        # Note that the parameters below are not required but kindly requested by PubMed Central
        # https://www.ncbi.nlm.nih.gov/pmc/tools/developers/
        self.pubmed = PubMed(tool=self.tool, email=self.email)

        # * Create query to feed into Pubmed
        self.query = ""
        # First author
        if self.author1 is not None:
            if '#' in str(self.author1):
                self.author1 = str(self.author1).replace('#',' ')
            self.query = self.query+str(self.author1)[2:-2]+' [1au] AND '
        # Authors
        if self.authors is not None:
            for author in self.authors.split(' '):
                if '#' in author:
                    author = author.replace('#',' ')
                self.query = self.query+author+' [auth] AND '
        # Title
        if self.title is not None:
            for tword in self.title:
                self.query = self.query+tword+' [ti] AND '
        # Terms
        if self.terms is not None:
            for item in self.terms.split(' '):
                self.query=self.query+item+' AND '
        # User query
        if self.userquery is not None:
            userquery = str(self.userquery)[2:-2]
            self.query=self.query+userquery+' AND '

        
        # Calculate what the start date is for articles to be included based on user settings
        if self.psLast is not None:
            # Only include articles published in the last <x> years
            self.dYa = datetime.now() - relativedelta(years=int(self.psLast))
            self.dayYearsAgo = str(self.dYa).split(' ')[0].replace('-','/')
            self.dYaQuery = '('+self.dayYearsAgo+'[Date - Create] : "3000"[Date - Create])'

        else:
            self.dYaQuery = '("'+self.psYear+'/01/01"[Date - Create] : "3000"[Date - Create])'
        self.query=self.query + self.dYaQuery 

        # Announce created query for verification:
        print(f'''
        This is your query:
        {self.query}
        ''')
        
    def runQuery(self):    
        # Execute the query against the API
        self.results = self.pubmed.query(self.query, max_results=int(self.maxResults)+1)

        # Make dictionary to store data
        self.output={}

        # Loop over the retrieved articles
        self.nResults=0
        for result in self.results:
            self.nResults=self.nResults+1

        # Check if there are more than <n> results
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
