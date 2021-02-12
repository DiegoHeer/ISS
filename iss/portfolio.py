import pandas as pd
import xlwings as xw
import json
from os.path import join
import pymsgbox
from yahoofinancials import YahooFinancials
import yfinance
import handler


class Portfolio:
    def __init__(self):
        self.database_path = r"D:\PythonProjects\iss\iss\data\portfolio_database.json"
        self.main_sheet_name = "Portfolio"
        self.log_sheet_name = "Portfolio_Log"
        self.backend_sheet_name = "Portfolio_Backend"
        self.equity_sheet_name = "Portfolio_Equities"

        self.wb = None
        self.wb_path = None
        self.ws = None
        self.ticker = None

        self.database_df = None
        self.database_dict = None

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

    def get_ticker_selection(self):
        self.initialize_worksheet(sheet_type="main")

        combo_box = self.ws.OLEObjects("TickerBox").Object
        self.ticker = combo_box.Value

        return self.ticker

    def excel_log_to_df(self):
        self.database_df = pd.read_excel(self.wb_path, sheet_name=self.log_sheet_name)
        self.database_df["Transaction Date"] = self.database_df["Transaction Date"].astype(str)

        return self.database_df

    def df_to_dict(self):
        return self.database_df.to_dict(orient='index')

    def save_portfolio_dict(self):
        if self.database_df is None and self.database_dict is None:
            pymsgbox.alert("No input data to save in database")
        elif self.database_df is not None:
            self.database_dict = self.df_to_dict()

        with open(self.database_path, 'w') as file:
            json.dump(self.database_dict, file, indent=4, sort_keys=True)

    def get_portfolio_dict(self):
        with open(self.database_path, 'r') as file:
            self.database_dict = json.load(file)

        return self.database_dict

    def dict_to_df(self):
        self.database_df = pd.DataFrame(self.database_dict).transpose()

        return self.database_df

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
        pass

    def fill_in_earnings_block(self):
        pass

    def fill_in_time_block(self):
        pass

    def fill_in_status_block(self):
        pass

    def fill_in_balance_block(self):
        pass

    def fill_in_profits_block(self):
        pass

    def fill_in_ta_block(self):
        pass

    def fill_in_rule1_analysis_block(self, rule1_data):
        pass


def tester():
    test = Portfolio()
    rule1_data = test.get_rule1_data()
    test.fill_in_general_info_block(rule1_data)
    # pymsgbox.confirm(test.get_rule1_data())
