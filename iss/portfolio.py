import pandas as pd
import xlwings as xw
import json
import pymsgbox
from yahoofinancials import YahooFinancials
import yfinance
from yahoo_fin import stock_info as si
from datetime import datetime, timedelta
import pandas_market_calendars as mcal
from yahoo_earnings_calendar import YahooEarningsCalendar
import dateutil.parser
import calendar
from os.path import join
import matplotlib.pyplot as plt
import pathlib

import handler
from technical_analysis.ta import TA


def get_min_indicator(rule1_data, indicator):
    i = 0
    min_value = 0
    for key, value in rule1_data.items():
        if indicator.upper() in key.upper():
            if "%" in value:
                value = float(value.replace("%", "")) / 100

            if i == 0:
                min_value = value
            else:
                if value < min_value:
                    min_value = value

            i += 1

    return round(min_value, 4)


class Portfolio:

    def __init__(self):
        self.database_path = join(pathlib.Path(__file__).parent.absolute(), "data", "portfolio_database.json")
        self.backend_path = join(pathlib.Path(__file__).parent.absolute(), "data", "portfolio_backend.json")
        self.main_sheet_name = "Portfolio"
        self.log_sheet_name = "Portfolio_Log"
        self.backend_sheet_name = "Portfolio_Backend"
        self.equity_sheet_name = "Portfolio_Equities"

        self.wb = None
        self.wb_path = None
        self.ws = None
        self.ws_non_api = None
        self.ticker = None

        # Database parameters
        self.db_df = None
        self.db_dict = None

        # Backend parameters
        self.bk_df = None
        self.bk_dict = None

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        pass

    def initialize_worksheet(self, sheet_type):
        if sheet_type == "log":
            sheet_name = self.log_sheet_name
        elif sheet_type == "backend":
            sheet_name = self.backend_sheet_name
        elif sheet_type == "equities":
            sheet_name = self.equity_sheet_name
        else:
            sheet_name = self.main_sheet_name

        self.wb = xw.Book.caller()
        self.wb_path = self.wb.fullname
        self.ws = self.wb.sheets[sheet_name].api
        self.ws_non_api = self.wb.sheets[sheet_name]

    def get_ticker_selection(self):
        self.initialize_worksheet(sheet_type="main")

        combo_box = self.ws.OLEObjects("TickerBox").Object
        self.ticker = combo_box.Value

        return self.ticker

    def excel_log_to_df(self):
        self.db_df = pd.read_excel(self.wb_path, sheet_name=self.log_sheet_name, engine='openpyxl')
        self.db_df["Transaction Date"] = self.db_df["Transaction Date"].astype(str)

        return self.db_df

    def excel_equities_to_df(self):
        self.db_df = pd.read_excel(self.wb_path, sheet_name=self.equity_sheet_name, engine='openpyxl')

        return self.db_df

    def df_to_dict(self, index_bool=False, dict_type='database'):
        if dict_type == 'database':
            if index_bool is True:
                self.db_dict = self.db_df.to_dict(orient='index')
            else:
                self.db_dict = self.db_df.to_dict()

            return self.db_dict
        else:
            if index_bool is True:
                self.bk_dict = self.bk_df.to_dict(orient='index')
            else:
                self.bk_dict = self.bk_df.to_dict()

            return self.bk_dict

    def save_dicts_to_json(self):
        # TODO: this function needs to be tested
        if self.db_df is None and self.db_dict is None and self.bk_dict is None:
            pymsgbox.alert("No input data to save in database/backend.")
            return
        elif self.db_df is not None:
            self.db_dict = self.df_to_dict(index_bool=True)

        if self.bk_dict is not None:
            with open(self.backend_path, 'w') as file:
                json.dump(self.bk_dict, file, indent=4, sort_keys=True)

        with open(self.database_path, 'w') as file:
            json.dump(self.db_dict, file, indent=4, sort_keys=True)

    def get_portfolio_dict(self):
        with open(self.database_path, 'r') as file:
            self.db_dict = json.load(file)

        return self.db_dict

    def get_backend_dict(self):
        with open(self.backend_path, 'r') as file:
            self.bk_dict = json.load(file)

        return self.bk_dict

    def save_backend_dict(self):
        # TODO: this function needs to be tested
        # Get dataframe with backend data
        self.bk_df = pd.read_excel(self.wb_path, sheet_name=self.backend_sheet_name, engine='openpyxl')

        # Change backend dataframe to a dictionary
        new_bk_data = self.df_to_dict(index_bool=True, dict_type='backend')

        # Get the current dictionary from the json file
        self.get_backend_dict()

        # Change the backend dictionary to include the new data
        self.get_ticker_selection()
        self.bk_dict[self.ticker] = new_bk_data

        # Save updated dictionary
        self.save_dicts_to_json()

    def update_backend_excel(self):
        # TODO: this function needs to be tested
        # Get backend dictionary
        self.get_backend_dict()

        # Extract the data related to the selected ticker
        self.get_ticker_selection()
        ticker_data = self.bk_dict[self.ticker]

        # Get an iss translation list for the backend sheet
        result_dict = handler.translate_dict_keys(ticker_data, 'Backend')

        # initialize xlwings worksheet
        self.initialize_worksheet(sheet_type="backend")

        # Enter all required data using a for loop based on the named ranges
        for key, value in result_dict.items():
            self.ws.Range(key).Value = ticker_data[value]

    def dict_to_df(self):
        self.db_df = pd.DataFrame(self.db_dict).transpose()

        return self.db_df

    def get_rule1_data(self):
        self.get_ticker_selection()
        # Get Rule#1 data
        with handler.FSHandler(self.main_sheet_name) as get_data:
            rule1_data = get_data.extract_rule1_metrics_data(self.ticker)

        return rule1_data

    def fill_in_summary_block(self):
        self.get_ticker_selection()
        self.initialize_worksheet(sheet_type="backend")

        self.ws.Range("market_cap").Value = YahooFinancials(self.ticker).get_market_cap()
        self.ws.Range("ttm_eps").Value = YahooFinancials(self.ticker).get_earnings_per_share()
        self.ws.Range("ttm_pe").Value = YahooFinancials(self.ticker).get_pe_ratio()
        self.ws.Range("current_vol").Value = YahooFinancials(self.ticker).get_current_volume()
        self.ws.Range("av_vol_10_day").Value = YahooFinancials(self.ticker).get_ten_day_avg_daily_volume()
        self.ws.Range("av_vol_3_months").Value = YahooFinancials(self.ticker).get_three_month_avg_daily_volume()

    def get_general_info(self):
        general_info_dict = dict()
        yfinance_info = yfinance.Ticker(self.ticker).info
        general_info_dict["company"] = yfinance_info["shortName"]
        general_info_dict["currency"] = yfinance_info["currency"]
        general_info_dict["country"] = yfinance_info["country"]
        general_info_dict["exchange"] = YahooFinancials(self.ticker).get_stock_exchange()
        general_info_dict["sector"] = yfinance_info["sector"]
        general_info_dict["industry"] = yfinance_info["industry"]

        return general_info_dict

    def fill_in_general_info_block(self, rule1_dict):
        self.get_ticker_selection()
        self.initialize_worksheet(sheet_type="backend")

        self.ws.Range("company").Value = rule1_dict["Company Name"]
        self.ws.Range("currency").Value = rule1_dict["Currency"]
        self.ws.Range("country").Value = rule1_dict["Country"]
        self.ws.Range("exchange").Value = rule1_dict["Exchange"]
        self.ws.Range("sector").Value = rule1_dict["Sector"]
        self.ws.Range("industry").Value = rule1_dict["Industry"]

    def fill_in_stock_price_block(self):
        self.get_ticker_selection()
        self.initialize_worksheet(sheet_type="backend")

        self.ws.Range("stock_price").Value = si.get_live_price(self.ticker)

    def get_log_total_sum(self, transaction_type, sum_column, start_date=None, end_date=None):
        # Takes the portfolio dataframe, applies two filters, and sums the total column data
        if self.ticker is None:
            self.get_ticker_selection()

        self.get_portfolio_dict()
        self.dict_to_df()

        sum_df = self.db_df.loc[(self.db_df['Ticker'] == self.ticker) & (self.db_df['Type'] == transaction_type)]

        if start_date is not None and end_date is None:
            sum_df = sum_df.loc[sum_df['Transaction Date'] >= start_date]
            pass
        elif start_date is None and end_date is not None:
            sum_df = sum_df.loc[sum_df['Transaction Date'] <= end_date]
            pass
        elif start_date is not None and end_date is not None:
            sum_df = sum_df.loc[sum_df['Transaction Date'] >= start_date]
            pass

        return sum_df[sum_column].sum()

    def get_capital_balance(self, start_date=None, end_date=None):
        shares_bought_value = self.get_log_total_sum('Buy', 'Value', start_date, end_date)
        shares_sold_value = self.get_log_total_sum('Sell', 'Value', start_date, end_date)

        return shares_bought_value - shares_sold_value

    def get_shares_balance(self, start_date=None, end_date=None):
        shares_bought = self.get_log_total_sum('Buy', 'Shares', start_date, end_date)
        shares_sold = self.get_log_total_sum('Sell', 'Shares', start_date, end_date)

        return shares_bought - shares_sold

    def fill_in_capital_block(self):
        capital_balance = self.get_capital_balance()
        shares_balance = self.get_shares_balance()

        self.initialize_worksheet(sheet_type='backend')
        stock_price = self.ws.Range("stock_price").Value

        total_capital = stock_price * shares_balance

        self.ws.Range("capital").Value = total_capital
        self.ws.Range("earnings").Value = total_capital - capital_balance

    def fill_in_time_block(self):
        self.get_ticker_selection()
        self.initialize_worksheet(sheet_type="backend")

        today = datetime.today().date()
        self.ws.Range('last_update').Value = today.strftime("%d/%m/%Y")

        # Gets the functioning dates in a a dataframe format of the stock exchange of the stock
        stock_exchange = self.ws.Range("exchange").Value
        exchange_calendar = None
        for exchange_name in mcal.get_calendar_names():
            if stock_exchange.upper() in exchange_name.upper() or exchange_name.upper() in stock_exchange.upper():
                exchange_calendar = mcal.get_calendar(exchange_name)
                break

        if stock_exchange is None:
            pymsgbox.alert("No valid calendar could be found of the stock exchange linked to the selected ticker")
            return

        self.ws.Range('last_market_close_date').Value = datetime.date(
            exchange_calendar.valid_days(start_date=today - timedelta(days=7), end_date=today)[-1]).strftime("%x %X")
        self.ws.Range('next_market_open_date').Value = datetime.date(
            exchange_calendar.valid_days(start_date=today, end_date=today + timedelta(days=7))[0]).strftime("%x %X")

        # Gets the next earning date for the specific stock (AKA Quarter Report release date)
        yec = YahooEarningsCalendar()
        self.ws.Range('next_quarter_report_date').Value = datetime.utcfromtimestamp(
            yec.get_next_earnings_date(self.ticker)).strftime("%x %X")

        # Gets the last earning date of a specific stock
        earnings_df = pd.DataFrame(yec.get_earnings_of(self.ticker))
        earnings_df['report_date'] = earnings_df['startdatetime'].apply(lambda x: dateutil.parser.isoparse(x).date())
        last_earnings_df = earnings_df.loc[earnings_df['report_date'] < today]

        self.ws.Range('last_quarter_report_date').Value = last_earnings_df['report_date'].iloc[0].strftime("%x %X")

        # Gets the date where you first start investing in this stock, and last buy and sell dates
        self.get_portfolio_dict()
        self.dict_to_df()
        stock_df = self.db_df.loc[self.db_df['Ticker'] == self.ticker]
        inv_start_date = datetime.strptime(stock_df['Transaction Date'].iloc[0], '%Y-%m-%d').strftime("%x %X")

        buy_df = stock_df.loc[stock_df['Type'] == 'Buy']
        sell_df = stock_df.loc[stock_df['Type'] == 'Sell']

        last_buy_date = datetime.strptime(buy_df['Transaction Date'].iloc[-1], '%Y-%m-%d').strftime("%x %X")
        last_sell_date = datetime.strptime(sell_df['Transaction Date'].iloc[-1], '%Y-%m-%d').strftime("%x %X")

        self.ws.Range('investment_start_date').Value = inv_start_date
        self.ws.Range('last_buy_date').Value = last_buy_date
        self.ws.Range('last_sell_date').Value = last_sell_date

    def fill_in_status_block(self):
        shares_balance = self.get_shares_balance()
        shares_bought = self.get_log_total_sum('Buy', 'Shares')
        shares_sold = self.get_log_total_sum('Sell', 'Shares')

        self.initialize_worksheet(sheet_type="backend")
        self.ws.Range('active_shares').Value = shares_balance
        self.ws.Range('shares_bought').Value = shares_bought
        self.ws.Range('shares_sold').Value = shares_sold

    def fill_in_balance_block(self):
        capital_balance = self.get_capital_balance()
        shares_bought_value = self.get_log_total_sum('Buy', 'Value')
        shares_sold_value = self.get_log_total_sum('Sell', 'Value')

        self.initialize_worksheet(sheet_type="backend")
        self.ws.Range('invested_capital').Value = capital_balance
        self.ws.Range('shares_bought_value').Value = shares_bought_value
        self.ws.Range('shares_sold_value').Value = shares_sold_value

    def get_profits(self, start_date, end_date):
        # Calculation of profits
        start_date = start_date.strftime("%Y-%m-%d")
        end_date = end_date.strftime("%Y-%m-%d")
        today = datetime.today().strftime("%Y-%m-%d")

        if start_date <= today <= end_date:
            # Get historical stock price data
            self.get_ticker_selection()
            historical_data = YahooFinancials(self.ticker).get_historical_price_data(start_date=start_date,
                                                                                     end_date=end_date,
                                                                                     time_interval="daily")
            start_price = historical_data[self.ticker]['prices'][0]['adjclose']
            end_price = historical_data[self.ticker]['prices'][-1]['adjclose']

            # Calculate earnings in the beginning of the period and at the end
            start_earnings = start_price * self.get_shares_balance(end_date=start_date) - self.get_capital_balance(
                end_date=start_date)
            end_earnings = end_price * self.get_shares_balance(end_date=end_date) - self.get_capital_balance(
                end_date=end_date)

            return end_earnings - start_earnings
        else:
            return 0

    def fill_in_profits_block(self):
        today = datetime.today()

        # Calculation of annual profits
        start_date = datetime(year=today.year, month=1, day=1)
        end_date = datetime(year=today.year, month=12, day=31)
        profits_annual = self.get_profits(start_date, end_date)

        # Calculation of 1°quarter profits
        start_date = datetime(year=today.year, month=1, day=1)
        end_date = datetime(year=today.year, month=3, day=31)
        profits_1quarter = self.get_profits(start_date, end_date)

        # Calculation of 2°quarter profits
        start_date = datetime(year=today.year, month=4, day=1)
        end_date = datetime(year=today.year, month=6, day=30)
        profits_2quarter = self.get_profits(start_date, end_date)

        # Calculation of 3°quarter profits
        start_date = datetime(year=today.year, month=7, day=1)
        end_date = datetime(year=today.year, month=9, day=30)
        profits_3quarter = self.get_profits(start_date, end_date)

        # Calculation of 4°quarter profits
        start_date = datetime(year=today.year, month=10, day=1)
        end_date = datetime(year=today.year, month=12, day=31)
        profits_4quarter = self.get_profits(start_date, end_date)

        # Calculation of month profits
        start_date = datetime(year=today.year, month=today.month, day=1)
        end_date = datetime(year=today.year, month=today.month, day=calendar.monthrange(today.year, today.month)[1])
        profits_month = self.get_profits(start_date, end_date)

        # Calculation of week profits
        start_date = today - timedelta(days=today.weekday())
        end_date = start_date + timedelta(days=6)
        profits_week = self.get_profits(start_date, end_date)

        # Calculation of day profits
        start_date = datetime.today() - timedelta(days=1)
        end_date = datetime.today()
        profits_day = self.get_profits(start_date, end_date)

        self.initialize_worksheet(sheet_type="backend")
        self.ws.Range('profits_annual').Value = profits_annual
        self.ws.Range('profits_1quarter').Value = profits_1quarter
        self.ws.Range('profits_2quarter').Value = profits_2quarter
        self.ws.Range('profits_3quarter').Value = profits_3quarter
        self.ws.Range('profits_4quarter').Value = profits_4quarter
        self.ws.Range('profits_month').Value = profits_month
        self.ws.Range('profits_week').Value = profits_week
        self.ws.Range('profits_day').Value = profits_day

    def fill_in_ta_block(self):
        # Get technical analysis data
        self.get_ticker_selection()
        ta_data = TA(self.ticker)
        ta_data.get_price_history()
        ta_data.get_indicators()

        self.initialize_worksheet(sheet_type="backend")
        self.ws.Range('ma_10_day').Value = ta_data.get_ma10_buy_sell().upper()
        self.ws.Range('macd_8_17_9').Value = ta_data.get_macd_buy_sell().upper()
        self.ws.Range('stoch_14_5').Value = ta_data.get_stoch_buy_sell().upper()

    def fill_in_rule1_analysis_block(self, rule1_data):
        self.get_ticker_selection()

        # Get info from portfolio equities sheet
        self.initialize_worksheet(sheet_type='equities')
        self.excel_equities_to_df()
        self.db_df = self.db_df.set_index('Ticker')
        self.df_to_dict(index_bool=True)

        meaning_type = self.db_dict[self.ticker]['Meaning Type']
        moat_type = self.db_dict[self.ticker]['Moat Type']

        # Get info from Rule #1 dataset
        self.initialize_worksheet(sheet_type="backend")
        self.ws.Range('meaning_type').Value = meaning_type
        self.ws.Range('moat_type').Value = moat_type
        self.ws.Range('min_roic').Value = get_min_indicator(rule1_data, 'roic')
        # self.ws.Range('min_roe').Value = min_roe
        self.ws.Range('min_equity_growth').Value = get_min_indicator(rule1_data, 'equity')
        self.ws.Range('min_eps_growth').Value = get_min_indicator(rule1_data, 'eps')
        self.ws.Range('min_sales_growth').Value = get_min_indicator(rule1_data, 'sales')
        self.ws.Range('min_fcf_growth').Value = get_min_indicator(rule1_data, 'fcf')
        self.ws.Range('min_ocf_growth').Value = get_min_indicator(rule1_data, 'ocf')
        self.ws.Range('payoff_debt_possible').Value = rule1_data['Debt - Payoff Possible']
        self.ws.Range('sticker_price').Value = round(rule1_data['Sticker Price'], 2)
        self.ws.Range('mos_price').Value = round(rule1_data['MOS Price'], 2)
        # self.ws.Range('payback_time').Value = payback_time
        # self.ws.Range('payback_price').Value = payback_price

        # pymsgbox.confirm(rule1_data)
        # TODO: After correcting Rule #1 metrics calculations, activate lines in this function.
        #  Also include it in the ISS translation list.

    def get_ta_chart(self):
        # Update the ta chart for the selected ticker
        self.get_ticker_selection()
        handler.gen_technical_analysis_chart(self.ticker, show_fig=False)

        # Get technical analysis chart path
        chart_storage_path = join(pathlib.Path(__file__).parent.absolute(), "data", "ta_charts")
        chart_path = join(chart_storage_path, self.ticker + ".png")

        # Insert the new picture into current picture location in the portfolio sheet
        self.initialize_worksheet(sheet_type="portfolio")
        self.ws_non_api.pictures['ta_chart'].update(chart_path)

    def get_portfolio_chart(self):
        # Get first list of tickers (no duplicates)
        self.get_portfolio_dict()
        self.dict_to_df()
        ticker_list = self.db_df['Ticker'].tolist()
        ticker_list = list(set(ticker_list))

        # Get total capitalization of tickers
        total_cap_dict = dict()
        for ticker in ticker_list:
            self.ticker = ticker
            stock_price = si.get_live_price(self.ticker)
            total_cap_dict[ticker] = self.get_shares_balance() * stock_price

        # Storage path for portfolio chart picture
        chart_directory = join(pathlib.Path(__file__).parent.absolute(), "data", "portfolio_charts")
        chart_path = join(chart_directory, "portfolio_chart.png")

        # Create a 'donut' chart using the total capitalization dictionary
        fig = plt.figure()
        fig.patch.set_facecolor('black')
        fig.patch.set_alpha(0.0)
        plt.rcParams['text.color'] = 'white'

        chart_circle = plt.Circle((0, 0), 0.7, color='black')
        plt.pie([value for key, value in total_cap_dict.items()], labels=[key for key, value in total_cap_dict.items()],
                autopct="%1.0f%%", pctdistance=0.6)
        p = plt.gcf()
        p.gca().add_artist(chart_circle)

        # Save new figure created
        fig.savefig(chart_path, bbox_inches="tight")

        # Insert the new picture into current picture location in the portfolio sheet
        self.initialize_worksheet(sheet_type="portfolio")
        self.ws_non_api.pictures['portfolio_chart'].update(chart_path)


def tester():
    test = Portfolio()
    # test.get_portfolio_chart()
    # test.get_portfolio_chart()
    test.fill_in_capital_block()
