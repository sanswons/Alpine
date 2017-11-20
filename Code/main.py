from Performance_Analysis import *
from pprint import pprint

if __name__ == '__main__':
    # change the threshold for ROE and Net Profit Margin
    ROE_threshold = 0
    NPM_threshold = 0
    all_companies_holdings, NPM_threshold, ROE_threshold,  overall_price_performance = check_all_companies_performance(ROE_threshold, NPM_threshold)
    pprint(all_companies_holdings)
    print(' \n Net Profit Margin : {0}\n ROE : {1}\n OVERALL PRICE PERFORMANCE : {2}'.format(NPM_threshold, ROE_threshold, overall_price_performance / len(all_companies_holdings)))
