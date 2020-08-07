from flask import Flask, request, render_template, redirect
import csv
import copy
import random
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font
from datetime import datetime

app = Flask(__name__)


#extract relevant data from the financial statements

def extract_data(financial_statement, name, row_index = True, column_index = True):
    
    #if both row index and column index are set
    if row_index != True and column_index != True:
        return financial_statement.filter(regex='(?i){}$'.format(name), axis=0).T.iloc[row_index,column_index]
    
    #if only row index is set
    elif row_index != True:
        return financial_statement.filter(regex='(?i){}$'.format(name), axis=0).T.iloc[row_index,:]
    
    #if only column index is set
    elif column_index != True:
        return financial_statement.filter(regex='(?i){}$'.format(name), axis=0).T.iloc[:,column_index]
    
    #if both are not set
    else:
        return financial_statement.filter(regex='(?i){}$'.format(name), axis=0).T

#check bad debt

def check_bad_debt(balance_sheet):
    
    LT_Debt = extract_data(balance_sheet, 'long-term debt')
    AR = extract_data(balance_sheet, 'accounts receivable')
    
    Bad_Debt = pd.DataFrame().reindex_like(AR).rename(columns = {'Accounts receivable':'Bad Debt'})
    Bad_Debt['Bad Debt'] = (LT_Debt.iloc[:,-1]/AR.iloc[:,-1])*100
    
    return Bad_Debt

#extract Inventory Turnover Rate
def extract_inventory_turnover_rate(balance_sheet, income_statement):
    
    COGS = extract_data(income_statement, 'total cost of goods sold')
    Inventories = extract_data(balance_sheet, 'inventories')
    
    Inventory_Turnover_Rate = pd.DataFrame().reindex_like(Inventories).rename(columns = {"Inventories":"Inventory Turnover Rate"})
    Inventory_Turnover_Rate['Inventory Turnover Rate'] = COGS.iloc[:,-1] / Inventories.iloc[:,-1]

    return Inventory_Turnover_Rate

#check inventory turnover rate

def check_inventory_turnover(Inventory_Turnover_Rate, Industry_Average):
    '''
    Returns inventory_turnover, mean, res
    '''
    inventory_turnover = Inventory_Turnover_Rate.iloc[-1,-1]
    inventory_turnover_res = "The company's inventory turnover rate is " + str(inventory_turnover)
    print("The company's inventory turnover rate is {}".format(inventory_turnover))
        
    mean = Industry_Average['Inventory Turnover Rate'].iloc[-1]
    mean_res = "The industry's average inventory turnover rate is " + str(mean)
    print("The company's average inventory turnover rate is {}".format(mean))
    
    res = ""
    if (abs(inventory_turnover - mean)/mean)*100 > 30:
        #print("More than 30% deviation in inventory turnover rate.")
        res = "More than 30% deviation in inventory turnover rate."
        print(res)
        
    else:
        #print("The deviation in inventory turnover rate is less than 30%.")
        res = "The deviation in inventory turnover rate is less than 30%."
        print(res)
    
    return inventory_turnover_res, mean_res, res


#check EBITDA

def check_EBITDA(income_statement, Industry_Average):
    '''
    Returns Latest_EBITDA, mean, res
    '''
    Latest_EBITDA = extract_data(income_statement, 'EBITDA', -1, -1)
    print("The company's latest EBITDA is {}".format(Latest_EBITDA))
    Latest_EBITDA_res = "The company's latest EBITDA is " + str(Latest_EBITDA)
        
    mean = Industry_Average['EBITDA'].iloc[-1]
    print("The industry's mean EBITDA is {}".format(mean))
    mean_res = "The industry's mean EBITDA is " + str(mean)
    
    res = ""
    if (abs(Latest_EBITDA - mean)/mean)*100 > 20:
        #print("More than 20% deviation in EBITDA.")
        res = "More than 20% deviation in EBITDA."
        print(res)
    else:
        res = "The deviation in EBITDA is less than 20%."
        print(res)

    return Latest_EBITDA_res, mean_res, res


#extract Debt Service Coverage Ratio

def extract_DSCR(balance_sheet, income_statement):
    
    Current_Liabilities = extract_data(balance_sheet, 'total current liabilities')
    Curr_Port_LTDebt = extract_data(balance_sheet, 'current portion of lt debt')
    Operating_Income = extract_data(income_statement, 'operating income')
    
    DSCR = pd.DataFrame().reindex_like(Current_Liabilities).rename(columns = {"Total Current Liabilities":"Debt Service Coverage Ratio"})
    DSCR['Debt Service Coverage Ratio'] = Operating_Income.iloc[:,-1] / (Current_Liabilities.iloc[:,-1] + Curr_Port_LTDebt.iloc[:,-1])
    
    return DSCR


#check Debt Service Coverage Ratio

def check_DSCR(balance_sheet, income_statement):
    '''
    Commented out new debt user input
    returns res
    '''
    loop = True
    
    while loop:
        
        #new_debt = float(input("Please enter a new debt value:"))
        new_debt = 123

        Latest_Current_Liabilities = extract_data(balance_sheet, 'total current liabilities', -1, -1)
        Latest_Curr_Port_LTDebt = extract_data(balance_sheet, 'current portion of lt debt', -1, -1)
        Latest_Operating_Income = extract_data(income_statement, 'operating income', -1, -1)

        New_Curr_Port_LTDebt = Latest_Curr_Port_LTDebt + new_debt

        DSCR = Latest_Operating_Income/(Latest_Current_Liabilities + New_Curr_Port_LTDebt)
        
        res = "With a new debt value of 123, "
        if DSCR < 1:
            res += "DSCR < 1. System rejects the new debt."
            print(res)
        elif DSCR < 1.2:
            res += "1 <= DSCR < 1.2. Further consideration required."
            print(res)
        else:
            res += "DSCR >= 1.2. System approves the new debt."
            print(res)
            
        #choice = input("Do you want to enter another new debt value? (Y/N):")
        
        #if choice == 'N':
        #    loop = False
        break

    return res


#extract Debt to EBITDA

def extract_Debt_to_EBITDA(balance_sheet, income_statement):
    
    Liabilities = extract_data(balance_sheet, 'total liabilities')
    EBITDA = extract_data(income_statement, 'EBITDA')
    
    Debt_to_EBITDA = pd.DataFrame().reindex_like(EBITDA).rename(columns = {"EBITDA":"Debt to EBITDA Ratio"})
    Debt_to_EBITDA['Debt to EBITDA Ratio'] = Liabilities.iloc[:,-1] / EBITDA.iloc[:,-1]

    return Debt_to_EBITDA


#check Debt to EBITDA

def check_debt_to_EBITDA(Debt_to_EBITDA, Debt_to_EBITDA_threshold):
    
    res = ""
    if Debt_to_EBITDA > (Debt_to_EBITDA_threshold * 1.15):
        res = "Exceeds threshold x 115%. Company is rejected."
        print(res)
    elif Debt_to_EBITDA > Debt_to_EBITDA_threshold and Debt_to_EBITDA <= (Debt_to_EBITDA_threshold * 1.15):
        res = "Exceeds threshold by less than 15%. Further consideration required."
        print(res)
    else:
        res = "Company's EBITDA is less than threshold value."
        print(res)
    
    return res


#extract Gearing ratio
def extract_gearing(balance_sheet):
    
    Liabilities = extract_data(balance_sheet, 'total liabilities')
    Shareholders_Equity = extract_data(balance_sheet, "Total Shareholders' Equity")
    
    Gearing = pd.DataFrame().reindex_like(Liabilities).rename(columns = {"Total Liabilities":"Gearing Ratio"})
    Gearing['Gearing Ratio'] = Liabilities.iloc[:,-1] / Shareholders_Equity.iloc[:,-1]

    return Gearing


#check Gearing ratio

def check_gearing(Gearing):

    if Gearing > 1:
        print("Company is highly leveraged.")
    else:
        print("Gearing ratio less than 1.")


#extract Net Worth
def extract_net_worth(balance_sheet):
    
    Total_Assets = extract_data(balance_sheet, 'total assets')
    Total_Liabilities = extract_data(balance_sheet, 'total liabilities')
    
    Net_Worth = pd.DataFrame().reindex_like(Total_Assets).rename(columns = {"Total Assets":"Net Worth"})
    Net_Worth['Net Worth'] = Total_Assets.iloc[:,-1] - Total_Liabilities.iloc[:,-1]

    return Net_Worth


#check Net Worth

def check_net_worth(Net_Worth):
    
    rising_net_worth = True
    prev_net_worth = Net_Worth['Net Worth'].iloc[0]

    for net_worth in Net_Worth['Net Worth']:
        if prev_net_worth > net_worth:
            rising_net_worth = False
            break
        else:
            prevv_net_worth = net_worth

    res = ""
    if rising_net_worth == True:
        res = "Consistently rising net worth."
        print(res)
    else:
        res = "Company's net worth is not consistently rising. Not recommended for a loan."
        print(res)
    
    return res


#calculate Scores

def calculate_score(Scores, name, company_value, industry_average):
    
    if name == 'Bad Debt' or name == 'Debt to EBITDA' or name == 'Gearing':
        score = (industry_average - company_value)/industry_average * 100
        
    else:
        score = (company_value - industry_average)/industry_average * 100
        
    if score < -100:
        score = 0
    elif score > 100:
        score = 100
    else:
        #Scale to 0 - 100
        score = (score + 100)/2
        
    Scores[name] = score
    
    return score


def store_data(count, SSIC, BS, IS, Database):
    
    #Bad Debt
    bad_debt = check_bad_debt(BS)
    
    #Inventory Turnover Rate
    inventory_turnover = extract_inventory_turnover_rate(BS, IS)
    
    #EBITDA
    EBITDA = extract_data(IS, 'EBITDA')
    
    #Net Worth
    net_worth = extract_net_worth(BS)
    
    #Gearing
    Gearing = extract_gearing(MCD_BS)
    
    if count == 1:
        Industry_Average = pd.DataFrame()
        Industry_Average = bad_debt
        Industry_Average['Inventory Turnover Rate'] = inventory_turnover['Inventory Turnover Rate']
        Industry_Average['EBITDA'] = EBITDA['EBITDA']
        Industry_Average['Net Worth'] = net_worth['Net Worth']
        Industry_Average['Gearing Ratio'] = Gearing['Gearing Ratio']
        
    else:
        Industry_Average = pd.read_excel(Database,str(SSIC),index_col=0)
        Industry_Average['Bad Debt'] = (Industry_Average['Bad Debt'].iloc[:] + bad_debt.iloc[:,-1])/count
        Industry_Average['Inventory Turnover Rate'] = (Industry_Average['Inventory Turnover Rate'].iloc[:] + inventory_turnover.iloc[:,-1])/count
        Industry_Average['EBITDA'] = (Industry_Average['EBITDA'].iloc[:] + EBITDA.iloc[:,-1])/count
        Industry_Average['Net Worth'] = (Industry_Average['Net Worth'].iloc[:] + net_worth.iloc[:,-1])/count
        Industry_Average['Gearing Ratio'] = (Industry_Average['Gearing Ratio'].iloc[:] + Gearing.iloc[:,-1])/count
        
    return Industry_Average


def plot_graph(count, Data, Industry_Average):
    
    calculation = Data.columns[0]
    
    if count == 1:
        plt.title('MCDonald\'s' + calculation)
        plt.plot(Data)
        plt.legend(['MCDonald\'s'])
        plt.savefig("{}.png".format(calculation))
        
    else:
        plt.title('MCDonald\'s {} vs Industry Average'.format(calculation))
        plt.plot(Data)
        plt.plot(Industry_Average['{}'.format(calculation)])
        plt.legend(['MCDonald\'s','Industry Average'])
        plt.savefig("{}.png".format(calculation))


#Check if the Company has been inputted to the system using UEN

def isFirstRead(UEN, List_of_Companies):
    
    if UEN in List_of_Companies['UEN'].values:
        return False

    else:
        return True


def extract_uen_and_ssic_from_BS(MCD_BS):
    '''
    Replace parameter with balance sheet from input
    returns tuple (UEN, SSIC)
    '''
    #Extract the company's UEN
    UEN = extract_data(MCD_BS, "UEN", 0, 0)

    #Extract the company's SSIC
    SSIC = extract_data(MCD_BS, "SSIC", 0, 0)

    return UEN, SSIC


def retrieve_industry_averages(UEN, List_of_Companies, SSIC, MCD_BS, MCD_IS, Database):
    '''
    Returns count and industry averages
    '''
    if isFirstRead(UEN, List_of_Companies):
    
        #Storing in List of Companies
        List_of_Companies.loc[len(List_of_Companies)+1] = ["Mcdonald's Corporation",UEN,SSIC,datetime.today().strftime('%Y-%m-%d')]
    
        #Counting the number of companies in the same industry
        count = List_of_Companies['SSIC'].value_counts().get(SSIC)
    
        #Retrieve the new industry average
        Industry_Average = store_data(count, SSIC, MCD_BS, MCD_IS, Database)
    
    else:
    
        #Counting the number of companies in the same industry
        count = List_of_Companies['SSIC'].value_counts().get(SSIC)
    
        #Retrieve the industry average
        Industry_Average = pd.read_excel(Database,str(SSIC),index_col=0)

    return count, Industry_Average


def get_bad_debt_score(MCD_BS, Industry_Average, Scores):
    '''
    PLOT_GRAPH FUNCTIONS ARE COMMENTED OUT
    '''
    bad_debt = check_bad_debt(MCD_BS)
    #plot_graph(count, bad_debt, Industry_Average)
    latest_bad_debt = bad_debt.iloc[-1,-1]
    latest_industry_average = Industry_Average['Bad Debt'].iloc[-1]
    score = calculate_score(Scores, 'Bad Debt', latest_bad_debt, latest_industry_average)
    bd_score = "The company's bad debt score is " + str(score)

    return bd_score


def get_inventory_turnover_rate_score(MCD_BS, MCD_IS, Industry_Average, Scores):
    '''
    PLOT_GRAPH FUNCTIONS ARE COMMENTED OUT
    not sure if you want the printed statements from check_inventory_turnover
    '''
    inventory_turnover = extract_inventory_turnover_rate(MCD_BS, MCD_IS)
    inv_turnover, mean, res = check_inventory_turnover(inventory_turnover, Industry_Average)
    #plot_graph(count, inventory_turnover, Industry_Average)

    latest_inventory_turnover = inventory_turnover.iloc[-1,-1]
    latest_industry_average = Industry_Average['Inventory Turnover Rate'].iloc[-1]
    score = calculate_score(Scores, 'Inventory Turnover Rate', latest_inventory_turnover, latest_industry_average)
    it_score = "The company's inventory turnover score is " + str(score)

    return it_score, inv_turnover, mean, res


def get_ebidta_score(MCD_IS, Industry_Average, Scores):
    '''
    PLOT_GRAPH FUNCTIONS ARE COMMENTED OUT
    not sure if you want the printed statements from check_EBIDTA
    '''
    lat_ebitda, mean, res = check_EBITDA(MCD_IS, Industry_Average)
    EBITDA = extract_data(MCD_IS, 'EBITDA')
    latest_EBITDA = EBITDA.iloc[-1,-1]
    latest_industry_average = Industry_Average['EBITDA'].iloc[-1]
    score = calculate_score(Scores, 'EBITDA', latest_EBITDA, latest_industry_average)
    eb_score = "The company's EBIDTA score is " + str(score)

    return eb_score, lat_ebitda, mean, res


def get_dscr(MCD_BS, MCD_IS, Scores):
    '''
    PLOT GRAPH FUNCTIONS ARE COMMENTED OUT
    USER INPUTS ARE ALSO COMMENTED OUT AND REPLACED WITH SOME VALUE
    '''
    DSCR = extract_DSCR(MCD_BS, MCD_IS)
    latest_DSCR = DSCR.iloc[-1,-1]
    print("The company's DSCR is {}".format(latest_DSCR))
    latest_DSCR_res = "The company's DSCR is " + str(latest_DSCR)
    res = check_DSCR(MCD_BS, MCD_IS)

    #lower_limit_DSCR = float(input("Please enter a lower limit DSCR:"))
    lower_limit_DSCR = 1.0
    score = calculate_score(Scores, 'DSCR', latest_DSCR, lower_limit_DSCR)
    print(score)
    dscr_score = "With a lower limit DSCR of 1.0, the company's DSCR score is " + str(score)

    return dscr_score, latest_DSCR_res, res


def get_debt_to_ebitda(MCD_BS, MCD_IS, Scores):
    Debt_to_EBITDA = extract_Debt_to_EBITDA(MCD_BS, MCD_IS)
    latest_Debt_to_EBITDA = Debt_to_EBITDA.iloc[-1,-1]
    print("The company's Debt to EBITDA ratio is {}".format(latest_Debt_to_EBITDA))
    latest_Debt_to_EBITDA_res = "The company's Debt to EBITDA ratio is " + str(latest_Debt_to_EBITDA)

    #Debt_to_EBITDA_threshold = float(input("Please enter a threshold for Debt to EBITDA ratio:"))
    Debt_to_EBITDA_threshold = 1.8

    res = check_debt_to_EBITDA(latest_Debt_to_EBITDA, Debt_to_EBITDA_threshold)

    score = calculate_score(Scores, 'Debt to EBITDA', latest_Debt_to_EBITDA, Debt_to_EBITDA_threshold)
    print(score)
    debt_to_ebitda_score = "With a Debt to EBITDA threshold of 1.8, the debt to EBITDA score is " + str(score)

    return debt_to_ebitda_score, latest_Debt_to_EBITDA_res, res


def get_gearing_ratio(MCD_BS, Industry_Average, Scores):
    Gearing = extract_gearing(MCD_BS)
    latest_gearing = Gearing.iloc[-1,-1]
    print("The company's gearing is {}".format(latest_gearing))
    latest_gearing_res = "The company's gearing is " + str(latest_gearing)

    latest_industry_average = Industry_Average['Gearing Ratio'].iloc[-1]
    score = calculate_score(Scores, 'Gearing', latest_gearing, latest_industry_average)
    print(score)
    gear_score = "The company's gearing ratio is " + str(score)

    return gear_score, latest_gearing_res


def get_net_worth(MCD_BS, Industry_Average, Scores):
    '''
    PLOT GRAPH FUNCTIONS ARE ALL COMMENTED OUT
    '''
    net_worth = extract_net_worth(MCD_BS)
    res = check_net_worth(net_worth)

    #plot_graph(count, net_worth, Industry_Average)

    latest_net_worth = net_worth.iloc[-1,-1]
    latest_industry_average = Industry_Average['Net Worth'].iloc[-1]
    score = calculate_score(Scores, 'Net Worth', latest_net_worth, latest_industry_average)
    print(score)
    net_worth_score = "The company's net worth is " + str(score)

    return net_worth_score, res


def get_overall_score(Scores):
    '''
    Returns overall bankability score.
    '''
    print(Scores)

    Scores_df = pd.DataFrame.from_dict(Scores, orient = 'index', columns = ['Score'] )
    Scores_df['Weightage (%)'] = 100/len(Scores)

    bankability = (Scores_df['Score'] * Scores_df['Weightage (%)']/100).sum()

    print("MCDonald's overall bankability score is {}%".format(bankability))

    overall_score = "The company's bankability is: " + str(bankability)

    return overall_score


@app.route("/")
def index():
    return render_template('index.html')


@app.route("/data", methods=['GET', 'POST'])
def data():
    if request.method == 'POST':
        csvfile1 = request.files['csvfile1']
        csvfile2 = request.files['csvfile2']
        file = csvfile2.filename
        file_details = file[:-4] + "'s results:"
        database = pd.ExcelFile(csvfile1)
        comp = pd.ExcelFile(csvfile2)
        #database = pd.ExcelFile(request.files['csvfile1'])
        #comp = pd.ExcelFile(request.files['csvfile2'])
        #print(comp)
        MCD_BS = pd.read_excel(comp,'Balance Sheet',index_col=0)
        MCD_IS = pd.read_excel(comp,'Income Statement',index_col=0)
        
        Scores = {}
        
        List_of_Companies = pd.read_excel(database,'List of Companies',index_col=0)

        UEN, SSIC = extract_uen_and_ssic_from_BS(MCD_BS)

        count, Industry_Average = retrieve_industry_averages(UEN, List_of_Companies, SSIC, MCD_BS, MCD_IS, database)

        bad_debt_score = get_bad_debt_score(MCD_BS, Industry_Average, Scores)

        inv_turnover_score, inv_turnover, inv_turnover_mean, inv_turnover_res = get_inventory_turnover_rate_score(MCD_BS, MCD_IS, Industry_Average, Scores)

        ebidta_score, lat_ebitda, ebidta_mean, ebidta_res = get_ebidta_score(MCD_IS, Industry_Average, Scores)

        dscr_score, latest_dscr, dscr_res = get_dscr(MCD_BS, MCD_IS, Scores)

        debt_to_ebitda_score, latest_Debt_to_EBITDA_res, debt_to_ebitda_res = get_debt_to_ebitda(MCD_BS, MCD_IS, Scores)   

        gearing_ratio_score, latest_gearing = get_gearing_ratio(MCD_BS, Industry_Average, Scores)     

        net_worth_score, net_worth_res = get_net_worth(MCD_BS, Industry_Average, Scores)

        bankability = get_overall_score(Scores)
        
        return render_template(
            'data.html',
            bad_debt_score=bad_debt_score,
            inv_turnover_score=inv_turnover_score,
            inv_turnover=inv_turnover,
            inv_turnover_mean=inv_turnover_mean, 
            inv_turnover_res=inv_turnover_res,
            ebidta_score=ebidta_score, 
            lat_ebitda=lat_ebitda, 
            ebidta_mean=ebidta_mean, 
            ebidta_res=ebidta_res,
            dscr_score=dscr_score,
            latest_dscr=latest_dscr,
            dscr_res=dscr_res,
            debt_to_ebitda_score=debt_to_ebitda_score, 
            latest_Debt_to_EBITDA_res=latest_Debt_to_EBITDA_res,
            debt_to_ebitda_res=debt_to_ebitda_res,
            gearing_ratio_score=gearing_ratio_score, 
            latest_gearing=latest_gearing,
            net_worth_score=net_worth_score, 
            net_worth_res=net_worth_res,
            bankability=bankability,
            file_details=file_details
        )


if __name__ == '__main__':
    app.run(debug = True)
