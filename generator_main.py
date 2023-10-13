# -*- coding: utf-8 -*-
"""
Created on Thu Jan 12 15:11:47 2023

@author: Kang Liew Bei
"""

#from fpdf import FPDF
#from datetime import date
#import csv
import sys
#import obj_reminder
import read_data


""" Main """
if __name__ == '__main__':

    """ Get the arguments from the command line """
    try:
        print('Running script ', sys.argv[0])
        year_level = sys.argv[1]
        data_filename = sys.argv[2]
        print(f'Year level: {year_level}, data file: {data_filename}')
        template_filename = 'Template_Reminder.xlsx'
        #data_filename = 'Sample_P5.xlsx'
        #year_level = 'P5'

    except IndexError:
        print('run script_name year_level filename, e.g.')
        print('run generator_main.py P1 Data_P1.xlsx')

    """ Read the Excel """
    """ Read the year level """
    """ Read the filename """
    """ Structure of list"""
    try:
        data_file = read_data.DataFile(data_filename, template_filename, year_level)
        data_row = data_file.read_row_per_student()

        data_file.save_file()

    except FileNotFoundError:
        print(f'File not found: {data_filename} and {template_filename}')