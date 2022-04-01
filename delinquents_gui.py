# -*- coding: utf-8 -*-
"""
Created on Fri Mar 25 13:42:39 2022

@author: dludwinski
"""

import PySimpleGUI as sg
import os, sys, getpass
file_dir = os.path.dirname(__file__)
sys.path.append(file_dir)
import new_delinquents_script as nd
user = getpass.getuser()

sg.theme('DarkAmber')

pre_set_save = f'C:/Users/{user}/Downloads'

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
            if nd.phone_num_check():
                nd.export_audits(file_location=f_path,
                                 report_name=import_file)                
            else:
                sg.Popup('\nBAD PHONE NUMBER(S) IN THE REPORT!\n\n')
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