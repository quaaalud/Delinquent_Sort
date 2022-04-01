# -*- coding: utf-8 -*-
"""
Created on Fri Apr  1 12:30:02 2022

@author: dludwinski
"""

import os, sys, getpass
file_dir = os.path.dirname(__file__)
sys.path.append(file_dir)
user = getpass.getuser()

import pandas as pd
import datetime as dt
import PySimpleGUI as sg

pre_set_save = f'C:/Users/{user}/Downloads' # Used for exporting reports

def import_paylink_delinquents_gui(report_name):    
    delinquents = pd.read_csv(report_name,
                              keep_default_na=True)
    delinquents.fillna('', inplace=True)
    delinquents.drop(columns=['Customer'], inplace=True)
    delinquents = days_delinquent_calc(delinquents) 
    format_dates('Next Due Date', delinquents)
    format_dates('Today', delinquents)
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
   
def delinquent_audit_function(report_name):    

    all_delinquents = import_paylink_delinquents_gui(report_name)
    clean_delinquents = remove_accounts(all_delinquents)
    clean_delinquents = sort_by_days_delq(clean_delinquents)
    return clean_delinquents

def format_dates(column, df):
    column = str(column)
    df[column] = pd.to_datetime(df[column], format="%Y/%m/%d")
    df[column] = df[column].dt.strftime("%m/%d/%Y")

def export_audits(file_location, report_name):

    all_delinquents = import_paylink_delinquents_gui(report_name)
    clean_delq = delinquent_audit_function(report_name)
    nvps, aegis, olympus = sort_companies(clean_delq)
    
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

def phone_num_check(report_name):

    clean_delq = import_paylink_delinquents_gui(report_name)
    if clean_delq['Phone'].dtypes == object:
        print('\nMISSING PHONE NUMBER!\n\n') 
        return 1
    
    elif 1111111111 > clean_delq['Phone'].min():
        print('\nPHONE MISSING DIGITS!\n\n')
        return 2
    
    elif clean_delq['Phone'].max() > 11999999999:
        print('\nPHONE NUMBER WITH OVER 10 DIGITS!\n\n')
        return 3
    
    else:
        return 4

sg.theme('DarkAmber')

pre_set_save = f'C:/Users/{user}/Downloads'

def delinquent_app_gui():
    layout = [
        [sg.Text('DELINQUENT ACCOUNT SORT', justification='center')],
        [sg.Text('Select a file to audit:', size=(15, 1), auto_size_text=False,
                 justification='left'),sg.InputText(
                     '', key='-FILE-'),
         sg.FileBrowse()],
        [sg.Text('Save location:', size=(15, 1), auto_size_text=False,
                 justification='left'),sg.InputText(
                     '', key='-SAVE-'),
         sg.FolderBrowse()],
        [[sg.Button('Submit')], [sg.Button('Cancel')]],
        ]
    
    window = sg.Window('DELINQUENT ACCOUNT SORT APP', layout)
    
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancel':
            break
        if event == 'Submit':
            import_file = values['-FILE-']
            if values['-SAVE-'] == '':
                f_path = pre_set_save
            else:
                f_path = values['-SAVE-']
            try:
                if phone_num_check(import_file) == 1:
                    sg.Popup('\nPHONE NUMBER(S) MISSING IN THE REPORT!\n\n')
                if phone_num_check(import_file) == 2:
                    sg.Popup('\nPHONE NUMBER(S) MISSING DIGITS FOUND!\n\n')
                if phone_num_check(import_file) == 3:
                    sg.Popup('\nPHONE NUMBER(S) WITH OVER 10 DIGITS FOUND!\n\n')
                if phone_num_check(import_file) == 4:
                    export_audits(f_path, import_file)   
                    sg.Popup('\nYOUR SORTED FILES ARE READY!\n\n')
                
            except FileNotFoundError:
                sg.Popup('\nYOU DID NOT SELECT A FILE SILLY!\n\nTRY AGAIN!\n\n')
            except UnicodeDecodeError:
                file_name = os.path.split(import_file)
                sg.Popup(f'THE AUDIT FAILED!\n\nCHECK THE FILE:\n\n\
    "{file_name[-1].upper()}"\n\n\
    AND RETRY!\n')        
            values['-FILE-'] = ''
            values['-SAVE-'] = ''
            window.refresh()
    
    window.close()

delinquent_app_gui()