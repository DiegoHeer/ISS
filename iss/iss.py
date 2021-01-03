import xlwings as xw
import pandas as pd
import os
import pymsgbox

from update import UpdateTable


def update_screener():
    update = UpdateTable('AAPL')
    update.table_to_ticker_list('screener')
    update.extract_fs_data()
    update.data_to_table()

# TODO: Make functions below
def update_watchlist():
    pass


def move_to_watchlist():
    pass


def move_to_portfolio():
    pass


def show_user_instructions(sheet_name):
    # Select the correct sheet where the extraction of tickers will happen
    if sheet_name.lower() == 'watchlist':
        pymsgbox.confirm('test')
    else:
        pymsgbox.confirm('test')
