from bs4 import BeautifulSoup
import requests
from datetime import datetime
import pymsgbox
import webbrowser

from portfolio import Portfolio


class SEC:

    def __init__(self):
        # Base url of the SEC website
        self.base_url = r"https://www.sec.gov"

        # Official SEC url that contains all the CIK x Ticker data
        self.cik_url = self.base_url + r'/include/ticker.txt'

        # Get ticker from Portfolio sheet
        with Portfolio() as p:
            self.ticker = p.get_ticker_selection()

        # Get CIK number from SEC website
        self.cik = self.get_cik_number()

        # Base url for finding company filings
        self.edgar_url = self.base_url + "/cgi-bin/browse-edgar"

        # Define search parameters for the SEC EDGAR browser
        self.param_dict = {'action': 'getcompany',
                           'CIK': str(self.cik),
                           'dateb': datetime.today().date().strftime("%Y%m%d"),
                           'start': '',
                           'output': 'atom',
                           'count': '100'}

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def get_cik_number(self):
        # Getting data from SEC website by scraping
        txt_content = requests.get(self.cik_url).text

        # Parsing CIK x Ticker data into a dictionary
        mapping_dict = dict()
        for mapping in txt_content.split('\n'):
            company_ticker = mapping.split('\t')[0]
            company_cik = mapping.split('\t')[1]

            mapping_dict[company_ticker] = company_cik

        # Returning CIK number based on ticker
        try:
            cik_number = mapping_dict[self.ticker.lower()]
            return cik_number
        except KeyError:
            pymsgbox.alert(
                "No CIK number found for selected ticker. Only companies trading in US have SEC document filing.")
            exit()

    def open_report(self, bool_annual=True):
        # Define type of search for 10K annual reports on SEC EDGAR
        self.param_dict['type'] = '10-K'

        # If bool annual is False the type of report requested is quarterly
        if bool_annual is False:
            self.param_dict['type'] = '10-Q'

        # Request the url, and parse the response
        response = requests.get(url=self.edgar_url, params=self.param_dict)
        soup = BeautifulSoup(response.content, features='lxml')

        # Find all entry tags
        entries = soup.find_all('entry')

        # Loop through each found entry
        filing_detail_link = None
        for entry in entries:
            # Grab the link to the filing detail site of the first entry
            filing_detail_link = entry.find('filing-href').text
            break

        # Check if the filing detail link is not blank
        if filing_detail_link is not None:
            # Extract new data from link
            response = requests.get(filing_detail_link)
            soup = BeautifulSoup(response.content, features='lxml')

            # Get first 10-K/10-Q document from table
            # Get table
            doc_format_table = soup.find('table', summary='Document Format Files')

            # Get all rows from table
            all_rows = doc_format_table.find_all('tr')

            # Loop through all rows to get correct link
            for row in all_rows:
                xbrl_marker = row.find('span')

                if xbrl_marker is not None:
                    file_link = self.base_url + row.find('a').get('href')

                    # Open the link in the browser
                    webbrowser.open(file_link)
                    break
