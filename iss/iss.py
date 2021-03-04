import pymsgbox
import webbrowser

import handler
import portfolio
from handler import FSHandler
from portfolio import Portfolio
from sec import SEC


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
    if sheet_name == 'Portfolio':
        access = Portfolio()
        ticker = access.get_ticker_selection()

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
    ticker_change = Portfolio()

    # Get temporary data from portfolio backend json
    ticker_change.update_backend_excel()

    # Get the Technical Analysis chart already made for the selected ticker
    ticker_change.get_ta_chart(bool_update=False)


def portfolio_new_entry():
    with Portfolio() as entry:
        # Get the equity list based on the transaction log list
        entry.get_equity_list()

        # Ask new ticker to user
        ticker = handler.ask_ticker_to_user()

        # Update data validation of ticker selection in Portfolio sheet with new ticker
        entry.update_ticker_selection_combo_box(new_ticker=ticker)

        # Show buy transaction form to user
        entry.transaction_entrybox("Buy")

    # Update complete portfolio sheet
    portfolio_update_all()


def portfolio_update_all():
    update_all = Portfolio()

    # Get important data for updates
    rule1_data = update_all.get_rule1_data()

    # Update all main blocks
    update_all.fill_in_summary_block()
    update_all.fill_in_general_info_block(rule1_data)
    update_all.fill_in_stock_price_block()
    update_all.fill_in_capital_block()
    update_all.fill_in_time_block()
    update_all.fill_in_status_block()
    update_all.fill_in_balance_block()
    update_all.fill_in_profits_block()
    update_all.fill_in_ta_block()
    update_all.fill_in_rule1_analysis_block(rule1_data)

    # Update TA chart
    update_all.get_ta_chart()

    # Update Portfolio chart
    update_all.get_portfolio_chart()

    # Store new data in the portfolio backend json file
    update_all.save_backend_dict()


def portfolio_buy():
    buy = Portfolio()
    buy.transaction_entrybox("Buy")


def portfolio_sell():
    sell = Portfolio()
    sell.transaction_entrybox("Sell")


def portfolio_update_ta():
    update_ta = Portfolio()

    # Updates the Technical Analysis text block
    update_ta.fill_in_ta_block()

    # Update the TA Chart
    update_ta.get_ta_chart()

    # Update Portfolio Chart
    update_ta.get_portfolio_chart()

    # Store new data in the portfolio backend json file
    update_ta.save_backend_dict()


def portfolio_update_rule1():
    update_rule1 = Portfolio()

    # Fill in Rule #1 block
    rule1_data = update_rule1.get_rule1_data()
    update_rule1.fill_in_rule1_analysis_block(rule1_data)

    # Store new data in the portfolio backend json file
    update_rule1.save_backend_dict()


def portfolio_open_last_annual_report():
    report = SEC()
    report.open_report(bool_annual=True)


def portfolio_open_last_quarterly_report():
    report = SEC()
    report.open_report(bool_annual=False)


def portfolio_still_to_be_defined():
    # TODO: Make functions for portfolio_still_to_be_defined
    pass
