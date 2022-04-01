# -*- coding: utf-8 -*-
"""
Created on Tue Mar 22 11:49:00 2022

@author: dludwinski
"""
import pandas as pd
import datetime as dt
import os, sys, getpass


file_dir = os.path.dirname(__file__)
sys.path.append(file_dir)
user = getpass.getuser()

##############################################################################
# Do not edit anything above this line


file_location = f'C:/Users/{user}/Desktop' # Used for both Import & Export

report_name = 'DelinquencyReport'    # Name of the saved Paylink report


# Do not edit anything above below line
##############################################################################


def import_paylink_delinquents(file_location=file_location,
                               report_name=report_name):
    if report_name == '':
        report_name = 'DelinquencyReport'
    else:
        report_name = report_name
    delinquents = pd.read_csv(f'{file_location}\{report_name}.csv',
                              keep_default_na=True)
    delinquents.fillna('', inplace=True)
    delinquents.drop(columns=['Customer'], inplace=True)
    delinquents = days_delinquent_calc(delinquents)   
    return delinquents

def import_paylink_delinquents_gui(report_name=''):    
    delinquents = pd.read_csv(f'{report_name}',
                              keep_default_na=True)
    delinquents.fillna('', inplace=True)
    delinquents.drop(columns=['Customer'], inplace=True)
    delinquents = days_delinquent_calc(delinquents)   
    return delinquents

def days_delinquent_calc(df):
    df.insert(13, 'Today', dt.date.today())
    df_col_to_datetime('Today', df)
    df_col_to_datetime('Next Due Date', df)
    df.insert(14, 'Days Delq.', (df['Today'] - df['Next Due Date']))
    return df

def df_col_to_datetime(column, df):
    column = str(column)
    df[column] = pd.to_datetime(df[column])
    
def double_digit_nums(number=int()):
    if number <= 9:
        number = f'0{number}'
    return number  
    
def sort_companies(df):
    audit = df.copy()  
    aegis = audit.loc[
        audit[
            'Seller Name'] == 'AEGIS WARRANTY GROUP'] 
    olympus = audit.loc[
        audit[
            'Seller Name'] == 'OLYMPUS PROTECTION'] 
    nvps = audit.loc[
        audit[
            'Seller Name'] != 'AEGIS WARRANTY GROUP' ]  
    nvps = nvps.loc[
        nvps[
            'Seller Name'] != 'OLYMPUS PROTECTION' ] 
    return nvps, aegis, olympus

def remove_accounts(df):
    df = df.copy()
    df = df.loc[df['Payment Pending'] == False]
    six_days = df.loc[
        df['Days Delq.'] <= pd.to_timedelta("6day")
        ] 
    df = df.loc[
        df['Days Delq.'] > pd.to_timedelta("6day")
        ] 
    df = df.loc[df['Customer Requested Cancel'] == 'N']
    df = pd.concat([df, six_days], axis=0)
    return df
  
def sort_by_days_delq(df):
    days_group = df.copy()
    six_days = days_group.loc[
        days_group['Days Delq.'] <= pd.to_timedelta("6day")
        ] 
    seven_to_24 = days_group.loc[
        (days_group['Days Delq.'] > pd.to_timedelta("6day"))
        & (days_group['Days Delq.'] <= pd.to_timedelta("24day"))
        ]
    
    pending_cnp = days_group.loc[
        (days_group['Days Delq.'] > pd.to_timedelta("24day"))
        ]  
    new_df = pd.concat([six_days, pending_cnp, seven_to_24])
    return new_df
   
def delinquent_audit_function(file_location=file_location,
                               report_name=''):    
    try:
        all_delinquents = import_paylink_delinquents(
            file_location, report_name)
    except:
        all_delinquents = import_paylink_delinquents_gui(report_name)
    clean_delinquents = remove_accounts(all_delinquents)
    clean_delinquents = sort_by_days_delq(clean_delinquents)
    return clean_delinquents

def format_dates(column, df):
    column = str(column)
    df[column] = pd.to_datetime(df[column], format="%Y/%m/%d")
    df[column] = df[column].dt.strftime("%m/%d/%Y")

    
    

def export_audits(file_location=file_location,
                  report_name=report_name):
    if report_name == '':
        report_name = 'DelinquencyReport'
    else:
        report_name = report_name
    try:
        all_delinquents = import_paylink_delinquents(
            file_location, report_name)
    except:
        all_delinquents = import_paylink_delinquents_gui(report_name)
    format_dates('Next Due Date', all_delinquents)
    format_dates('Today', all_delinquents)
    clean_delq = delinquent_audit_function()
    sorted_comp = sort_companies(clean_delq)
    nvps, aegis, olympus = sorted_comp[0], sorted_comp[1], sorted_comp[2]
    
    today = dt.date.today()
    year = today.year
    day = today.day
    day = double_digit_nums(day)
    month = today.month
    month = double_digit_nums(month)

    nvps.to_csv(
        f'{file_location}\DelinquencyReport_NVPS{year}{month}{day}.csv',
        index=False)
    aegis.to_csv(
        f'{file_location}\DelinquencyReport_AEGIS{year}{month}{day}.csv',
        index=False)
    olympus.to_csv(
        f'{file_location}\DelinquencyReport_OLYMPUS{year}{month}{day}.csv',
        index=False)
    
    with pd.ExcelWriter(
            f'{file_location}\All_Delinquents {today}.xlsx') as writer:
        
        all_delinquents.to_excel(
            writer,sheet_name="All Discounts", index=False
            )
        clean_delq.to_excel(
            writer, sheet_name="Rep Monthly Totals", index=False
            )  

def phone_num_check():
    try:
        clean_delq = import_paylink_delinquents(
            file_location, report_name)
        
    except:
        clean_delq = import_paylink_delinquents_gui(report_name)
    if clean_delq['Phone'].dtypes == object:
        print('\nMISSING PHONE NUMBER!\n\n') 
    elif 1111111111 > clean_delq['Phone'].min():
        print('\nPHONE MISSING DIGITS!\n\n')
    elif clean_delq['Phone'].max() > 11999999999:
        print('\nPHONE NUMBER WITH OVER 10 DIGITS!\n\n')
    else:
        return True
        

if __name__ == '__main__':
    if phone_num_check():
        export_audits()

