import openpyxl
import pickle
import os
import pandas
from pprint import pprint

all_companies_metrics = []
cols = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']

def extract_net_profit(DataSheet):
    net_profit_values = []
    for col in cols:
        net_profit_values.append(round((DataSheet[col + '30'].value / DataSheet[col + '17'].value) * 100, 2))
    return net_profit_values


def extract_price(DataSheet):
    price_values = []
    for col in cols:
        price_values.append(DataSheet[col + '90'].value)
    return price_values


def extract_ROE(DataSheet):
    ROE_values = []
    for col in cols:
        total_equity = DataSheet[col + '57'].value + DataSheet[col + '58'].value
        if total_equity > 0:
            ROE_values.append(round((DataSheet[col + '30'].value / total_equity)*100))
        else:
            ROE_values.append(None)
    return ROE_values


def extract_trailing_price(DataSheet):
    trailing_price = DataSheet['B8'].value
    return trailing_price


def extract_report_dates(DataSheet):
    report_dates = []
    for col in cols:
        report_dates.append(DataSheet[col +  '16'].value.year)
    return report_dates


def create_metrics_sheet(company):
    metrics = {}
    workbook = openpyxl.load_workbook('../Data/' + company)
    DataSheet = workbook.get_sheet_by_name('Data Sheet')
    metrics['Company_Name'] = company.replace('.xlsx', '')
    metrics['Net Profit Margin'] = extract_net_profit(DataSheet)
    metrics['Price']  = extract_price(DataSheet)
    metrics['Return On Equity'] = extract_ROE(DataSheet)
    metrics['Trailing_price'] = extract_trailing_price(DataSheet)
    metrics['Report Dates'] = extract_report_dates(DataSheet)
    all_companies_metrics.append(metrics)
    return all_companies_metrics


def create_all_companies_metrics():
    files  =  os.listdir('../Data/')
    for file in files:
        create_metrics_sheet(file)
    pprint(all_companies_metrics)
    with open('all_companies_metrics_data.pickle','wb') as handle:
        pickle.dump(all_companies_metrics, handle, protocol=pickle.HIGHEST_PROTOCOL)


print(create_all_companies_metrics())