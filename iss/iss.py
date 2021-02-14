import pymsgbox
import webbrowser
import xlwings as xw

import handler
from handler import FSHandler
import portfolio

import pandas as pd
import json


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
    # TODO: Make functions for move_to_portfolio
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
    handler.gen_technical_analysis_chart(ticker)


def see_trading_view_chart(sheet_name=None):
    if sheet_name is not None:
        wb = xw.Book.caller()
        ws = wb.sheets[sheet_name].api
        ticker = ws.OLEObjects("TickerBox").Object.Value
    else:
        ticker = handler.ask_ticker_to_user()

    # Home page url on tradingview.com about the stock
    base_url = rf"https://www.tradingview.com/symbols/{ticker}/"

    # get chart link and open it in the browser
    chart_link = handler.get_full_featured_tradingview_chart(base_url)
    webbrowser.open(chart_link)


def remove_non_approved_tickers():
    user_answer = pymsgbox.confirm("Remove also tickers were the MOAT needs to be checked?", "Ticker Removal",
                                   (pymsgbox.YES_TEXT, pymsgbox.NO_TEXT))
    remover = FSHandler('screener')
    if user_answer == pymsgbox.YES_TEXT:
        remover.dump_non_approved_tickers(True)
    else:
        remover.dump_non_approved_tickers(False)


def tester():
    portfolio.tester()


def portfolio_ticker_selection():
    # TODO: Make functions for portfolio_ticker_selection
    pass


def portfolio_new_entry():
    test = portfolio.Portfolio()
    test.get_portfolio_dict()
    pymsgbox.confirm(test.dict_to_df())

    # TODO: Make functions for portfolio_new_entry
    pass


def portfolio_update_all():
    # TODO: Make functions for portfolio_update_all
    pass


def portfolio_buy():
    # TODO: Make functions for portfolio_buy
    pass


def portfolio_sell():
    # TODO: Make functions for portfolio_sell
    pass


def portfolio_update_ta():
    # TODO: Make functions for portfolio_update_ta
    pass


def portfolio_update_rule1():
    # TODO: Make functions for portfolio_update_rule1
    pass


def portfolio_open_last_annual_report():
    # TODO: Make functions for portfolio_open_last_annual_report
    pass


def portfolio_open_last_quarterly_report():
    # TODO: Make functions for portfolio_open_last_quarterly_report
    pass


def portfolio_still_to_be_defined():
    # TODO: Make functions for portfolio_still_to_be_defined
    pass
