import pymsgbox

import handler
from handler import FSHandler


def update_screener():
    update = FSHandler('screener')
    update.rule1_data_to_table()


def update_watchlist():
    update = FSHandler('watchlist')
    update.rule1_data_to_table()


def move_to_watchlist():
    ticker = handler.ask_ticker_to_user()

    with FSHandler('screener') as mover:
        mover.delete_ticker_from_table(ticker)
        mover.add_ticker_to_table(ticker, sheet_name='watchlist')


def move_to_portfolio():
    # TODO: Make functions below
    pass


def show_user_instructions(sheet_name):
    # Select the correct sheet where the extraction of tickers will happen
    if sheet_name.lower() == 'watchlist':
        pymsgbox.confirm('test')
    else:
        pymsgbox.confirm('test')


def access_financial_statement():
    ticker = handler.ask_ticker_to_user()

    # Create an object to have access to the functions required to open the excel file
    access = FSHandler('screener')
    access.open_fs_excel_file(ticker)
