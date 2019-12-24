#!/usr/bin/env python3

# * Pubmed Queries via Python Pubmed API Wrapper
# 2019-12-23
# Vincent Koppelmans

# * Background
# Requests Guillaume:
# 1) Perform several searches based on items from a list
# 2) Limit the results to those published in the last 5 years
# 3) Ignore if the results is > 200. In this case, flag these search terms.

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
        'to the Pubmed API to obtain search results. It limits '
        'search results to those publications published in the last '
        '5 years. It only returns results if there are fewer '
        'than 200 results. If there are more than 200 results, '
        'this script will return a warning.')
    parser.add_argument('oFile',
                        help='Output file name. Path must exist.'
    )
    parser.add_argument('--terms',
                        nargs='+',
                        help='List of search terms'
    )    
args = parser.parse_args()

# * Class for Synonyms
class query(object):

    # * Store Terms
    def __init__(self):
        self.oFile = args.oFile
        self.terms = args.terms
        
    # * Build Object
    # Create a PubMed object that GraphQL can use to query
    def buildQuery(self):
        # Build Object and send some info to PubMed by their request
        # Note that the parameters below are not required but kindly requested by PubMed Central
        # https://www.ncbi.nlm.nih.gov/pmc/tools/developers/
        self.pubmed = PubMed(tool="", email="")

        # Create a GraphQL query in plain text
        # Example: query:
        # '(("2018/05/01"[Date - Create] : "3000"[Date - Create])) AND (Xiaoying Xian[Author] OR diabetes)'
        self.query = ""
        for item in self.terms:
            self.query=self.query+' AND '+item
        # Only include articles published in the last five years
        self.dFyA = datetime.now() - relativedelta(years=5)
        self.dayFiveYearsAgo = str(self.dFyA).split(' ')[0].replace('-','/')
        self.dFyAquery = '('+self.dayFiveYearsAgo+'[Date - Create] : "3000"[Date - Create])'
        self.query=self.dFyAquery + self.query

    def runQuery(self):    
        # Execute the query against the API
        self.results = self.pubmed.query(self.query, max_results=201)

        # Make dictionary to store data
        self.output={}

        # Loop over the retrieved articles
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
            journal = article.journal

            # Reshape author list
            authorString=''
            for author in authors:
                last=author['lastname']
                first=author['firstname']
                authorString=authorString+' '+last+', '+first+';'

            # Add results to the dictionary
            self.output[article_id] = [article_id, title, authorString, journal, publication_date, abstract]

            # Show information about the article
            # print(
            #     f'''{article_id} - {publication_date} - {title}\n{abstract}\n'''
            # )

        # Put data in a dataframe after extraction
        self.DF = pd.DataFrame.from_dict(self.output)
        self.DF = self.DF.T
        self.DF = self.DF.reset_index(drop=True) # Remove row names
        self.DF.columns = ["PMID","Title","Authors","Journal","PubDate","Abstract"]
        # Check if there are more than 200 results
        if len(self.DF) > 200:
            # Show warning
            print('More than 200 results found')
        elif len(self.DF) == 0:
            # Show warning
            print('No results found')            
        else:
            # Print number of results
            print(str(len(self.DF))+' result(s) obtained.')

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
