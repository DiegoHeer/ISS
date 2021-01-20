import pymsgbox

import handler
from handler import FSHandler


def update_screener():
    update = FSHandler('screener')
    update.rule1_data_to_table()


def update_watchlist():
    update = FSHandler('watchlist')
    update.rule1_data_to_table()
    update.ta_to_watchlist()


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


def see_technical_analysis_chart():
    ticker = handler.ask_ticker_to_user()

    # TODO: create the see_technical_analysis_chart()


def see_trading_view_chart():
    ticker = handler.ask_ticker_to_user()

    # TODO: create the see_trading_view_chart()


def remove_non_approved_tickers():
    user_answer = pymsgbox.confirm("Remove also tickers were the MOAT needs to be checked?", "Ticker Removal",
                                   (pymsgbox.YES_TEXT, pymsgbox.NO_TEXT))
    remover = FSHandler('screener')
    if user_answer == pymsgbox.YES_TEXT:
        remover.dumb_non_approved_tickers(True)
    else:
        remover.dumb_non_approved_tickers(False)
