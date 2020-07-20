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


@app.route("/")
def index():
    return render_template('index.html')


@app.route("/data", methods=['GET', 'POST'])
def data():
    if request.method == 'POST':
        database = pd.ExcelFile(request.form['csvfile1'])
        comp = pd.ExcelFile(request.form['csvfile2'])
        MCD_BS = pd.read_excel(comp,'Balance Sheet',index_col=0)
        MCD_IS = pd.read_excel(comp,'Income Statement',index_col=0)


        return render_template('data.html')

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
    
    inventory_turnover = Inventory_Turnover_Rate.iloc[-1,-1]
    print("The company's inventory turnover rate is {}".format(inventory_turnover))
        
    mean = Industry_Average['Inventory Turnover Rate'].iloc[-1]
    print("The industry's average inventory turnover rate is {}".format(mean))
    
    if (abs(inventory_turnover - mean)/mean)*100 > 30:
        print("More than 30% deviation in inventory turnover rate.")
    else:
        print("The deviation in inventory turnover rate is less than 30%.")

#check EBITDA

def check_EBITDA(income_statement, Industry_Average):
    
    Latest_EBITDA = extract_data(income_statement, 'EBITDA', -1, -1)
    print("The company's latest EBITDA is {}".format(Latest_EBITDA))
        
    mean = Industry_Average['EBITDA'].iloc[-1]
    print("The industry's mean EBITDA is {}".format(mean))
    
    if (abs(Latest_EBITDA - mean)/mean)*100 > 20:
        print("More than 20% deviation in EBITDA.")
    else:
        print("The deviation in EBITDA is less than 20%.")

#extract Debt Service Coverage Ratio

def extract_DSCR(balance_sheet, income_statement):
    
    Current_Liabilities = extract_data(balance_sheet, 'total current liabilities')
    Curr_Port_LTDebt = extract_data(balance_sheet, 'current portion of lt debt')
    Operating_Income = extract_data(income_statement, 'operating income')
    
    DSCR = pd.DataFrame().reindex_like(Current_Liabilities).rename(columns = {"Total Current Liabilities":"Debt Service Coverage Ratio"})
    DSCR['Debt Service Coverage Ratio'] = Operating_Income.iloc[:,-1] / (Current_Liabilities.iloc[:,-1] + Curr_Port_LTDebt.iloc[:,-1])
    
    return DSCR

if __name__ == '__main__':
    app.run(debug = True)
