#! python3
# -*- coding: utf-8 -*-

import os, glob, xlrd, openpyxl, constants
from shutil import copy as shucpy
from shutil import move as shmove
from datetime import datetime

# Standardizing inputs
def convert_to_xlsx(file):
    print('Converting input to xlsx...')
    file = xlrd.open_workbook(file)
    converted = openpyxl.Workbook()

    for i in range(0, file.nsheets):
        sheet = file.sheet_by_index(i)
        active = converted.active if i==0 else converted.create_sheet()
        active.title = sheet.name

        for row in range(0, sheet.nrows):
            for col in range(0, sheet.ncols):
                active.cell(row=row+1, column=col+1).value = sheet.cell_value(row, col)

    print('Converted.')
    return converted

def getExcelFiles(input_path, dest_path):
    print('Opening input files...')

    new_report = openpyxl.load_workbook(dest_path)
    print('Opened '+ dest_path)

    input = convert_to_xlsx(input_path) if input_path.endswith('s') else openpyxl.load_workbook(input_path)
    print('Opened '+ input_path)

    return (input, new_report)


def main():
    path = os.getcwd()

    #get newest input file
    input_path = max(glob.iglob(os.path.join(path, 'Input', '*.xls*')), key=os.path.getctime)
    template_path = os.path.join(path, 'Templates', 'Template.xlsx')
    dest_path = os.path.join(path, 'Output', 'INTLStatusReport.xlsx')
    input_files = os.listdir(os.path.join(path, 'Input'))
    backup_path = os.path.join(path, 'Backups')

    # create a working copy of the template
    print('Copying template to workspace...')
    shucpy(template_path, dest_path)

    print('New spreadsheet: ' + dest_path)
    input, new_report = getExcelFiles(input_path, dest_path)

    new_sheet = new_report.active
    input_sheet = input.active
    # print('Populating new spreadsheet...')

    for i in range(2,90):
        # sanitizing
        country = input_sheet.cell(row=i, column=2).value.replace(' ', '_').replace('.', '')
        task = input_sheet.cell(row=i, column=3).value.replace(' ', '_').replace('-', '_')

        # working with two possible data types--excel serial date or None
        tmp = input_sheet.cell(row=i, column=4).value

        if not tmp:
            start_time = ''
        elif type(tmp) is float:
            start_time = datetime(*(xlrd.xldate.xldate_as_tuple(tmp, 0))).strftime('%H:%M')
        else:
            start_time = tmp.strftime('%H:%M')

        tmp = input_sheet.cell(row=i, column=5).value
        if not tmp:
            end_time = ''
        elif type(tmp) is float:
            end_time = datetime(*(xlrd.xldate.xldate_as_tuple(tmp, 0))).strftime('%H:%M')
        else:
            end_time = tmp.strftime('%H:%M')

        # ignoring invalid input (+0.5 for more accurate rounding)
        tmp = float(input_sheet.cell(row=i, column=6).value)
        elapsed = int(tmp+0.5) if (tmp < 1000 and tmp > 0) else ''


        print('Country: ', country, ', start: ', start_time, ', end: ', end_time, 'elapsed: ', elapsed)

        #grabs proper index dictionary from constants.py
        dict=getattr(constants,task)

        # writing information to excel sheet
        new_sheet.cell(row=dict[country], column=3).value = start_time
        new_sheet.cell(row=dict[country], column=4).value = end_time
        new_sheet.cell(row=dict[country], column=5).value = elapsed

    print('Finished populating.')

    print('Saving files...')
    new_report.save(dest_path)
    print('Saved: '+dest_path)
    input.save(os.path.join(path, 'Input', 'Input.xlsx'))
    print('Saved: '+ os.path.join(path, 'Input', 'Input.xlsx'))
    print('Spreadsheet conversion complete.')
    print('Moving files....')

    for item in input_files:
        shmove(item, backup_path)


if __name__ == '__main__':
    main()
