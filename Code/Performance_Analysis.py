import pickle
from pprint import pprint
import pandas as pd

def check_all_companies_performance(ROE_threshold = 0, NPM_threshold = 0):
    with open('all_companies_metrics_data.pickle', 'rb') as handle:
        all_companies_metrics = pickle.load(handle)

    all_companies_holdings = []

    for company in all_companies_metrics:
        if not check_price_performance(company, ROE_threshold, NPM_threshold):
            continue
        else:
            company_holding = {}

            buying_year, cost_price, selling_year, selling_price, price_performance = check_price_performance\
            (company, ROE_threshold, NPM_threshold)

            company_holding['Company_Name'] = company['Company_Name']
            company_holding['Buying_Year'] = buying_year
            company_holding['Selling_Year'] = selling_year
            company_holding['Selling_Price'] = selling_price
            company_holding['Cost_Price'] = cost_price
            company_holding['Price_Performance'] = price_performance

            all_companies_holdings.append(company_holding)

    overall_price_performance = 0
    for company in all_companies_holdings:
        overall_price_performance += company['Price_Performance']

    all_companies_holdings_df = pd.DataFrame(all_companies_holdings)
    all_companies_holdings_df.to_excel('all_companies_holdings_'+ str(ROE_threshold) + '_' + str(NPM_threshold) + '.xlsx')

    return all_companies_holdings, NPM_threshold, ROE_threshold , overall_price_performance



def check_price_performance(company, ROE_threshold = 0, NPM_threshold = 0):
    NPM = company['Net Profit Margin']
    ROE = company['Return On Equity']
    Price = company['Price']
    Year  = company['Report Dates']

    for i in range(len(NPM)):
        if Price[i] != None:
            if NPM[i] > NPM_threshold and ROE[i] > ROE_threshold:
                cost_price = Price[i]
                buying_year = Year[i]
                break
            else:
                return None

    for j in range(1, len(NPM)):
        if Price[j]:
            if NPM[j] > NPM_threshold and ROE[j] > ROE_threshold:
                selling_price = Price[j]
                selling_year = Year[j]
            else:
                selling_price = Price[j-1]
                selling_year = Year[j-1]
                break


    if selling_year == 2017:
        selling_price = company['Trailing_price']
    if selling_year != buying_year:
        price_performance = (selling_price - cost_price) / \
                            (cost_price * (selling_year - buying_year)) * 100
    else:
        return None

    return buying_year, cost_price, selling_year, selling_price, price_performance



