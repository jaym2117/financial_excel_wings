import xlwings as xw
import sqlConfig as sql 
import pandas as pd

@xw.func
def get_income_statement_amount(companyDB, startDate, endDate, accounts):
    str(accounts)
    account_list = accounts.split(",")

    sql_statement = f'''
        SELECT	SUM(a.Debit - a.Credit) as Actual 
        FROM JDT1 a
        JOIN OACT b 
            ON a.Account = b.AcctCode
        WHERE b.Levels = 4 and a.RefDate >= '{startDate}' and a.RefDate <= '{endDate}' and (
    '''
    
    for account in account_list: 
        sql_statement += f"LEFT(FormatCode, 5) = {account} or "
    
    sql_statement = sql_statement[:-4]
    sql_statement += ")"

    dls = sql.SQL_Config('dlssap', companyDB, 'jason', 'Knights17?')  
    df = dls.sqlStatement(sql_statement)
    return df.iloc[0][0]

@xw.func
def get_balance_sheet_amount(companyDB, endDate, accounts):
    str(accounts)
    account_list = accounts.split(",")

    sql_statement = f'''
        SELECT	SUM(a.Debit - a.Credit) as Actual 
        FROM JDT1 a
        JOIN OACT b 
            ON a.Account = b.AcctCode
        WHERE b.Levels = 4 and a.RefDate <= '{endDate}' and (
    '''
    
    for account in account_list: 
        sql_statement += f"LEFT(FormatCode, 5) = {account} or "
    
    sql_statement = sql_statement[:-4]
    sql_statement += ")"

    dls = sql.SQL_Config('dlssap', companyDB, 'jason', 'Knights17?')  
    df = dls.sqlStatement(sql_statement)
    return df.iloc[0][0]


@xw.func
def get_retained_earnings(companyDB, endDate):
    sql_statement = f'''
        SELECT	SUM(a.Debit - a.Credit) as Actual 
        FROM JDT1 a
        JOIN OACT b 
            ON a.Account = b.AcctCode
        WHERE b.Levels = 4 and a.RefDate <= '{endDate}' and LEFT(FormatCode, 5) >= 40000 
    '''
    dls = sql.SQL_Config('dlssap', companyDB, 'jason', 'Knights17?')  
    df = dls.sqlStatement(sql_statement)
    return df.iloc[0][0]

@xw.func
def get_cash_flow_amount(companyDB, startDate, endDate, accounts):
    str(accounts)
    account_list = accounts.split(",")

    sql_statement = f'''
        SELECT	SUM(a.Debit - a.Credit) as Actual 
        FROM JDT1 a
        JOIN OACT b 
            ON a.Account = b.AcctCode
        WHERE b.Levels = 4 and a.RefDate >= '{startDate}' and a.RefDate <= '{endDate}' and (
    '''
    
    for account in account_list: 
        sql_statement += f"LEFT(FormatCode, 5) = {account} or "
    
    sql_statement = sql_statement[:-4]
    sql_statement += ")"

    dls = sql.SQL_Config('dlssap', companyDB, 'jason', 'Knights17?')  
    df = dls.sqlStatement(sql_statement)
    return df.iloc[0][0]

@xw.func
def get_net_income_amount(companyDB, startDate, endDate):

    sql_statement = f'''
        SELECT	SUM(a.Debit - a.Credit) as Actual 
        FROM JDT1 a
        JOIN OACT b 
            ON a.Account = b.AcctCode
        WHERE b.Levels = 4 and a.RefDate >= '{startDate}' and a.RefDate <= '{endDate}' and LEFT(FormatCode, 5) >= 40000
    '''

    dls = sql.SQL_Config('dlssap', companyDB, 'jason', 'Knights17?')  
    df = dls.sqlStatement(sql_statement)
    return df.iloc[0][0]
