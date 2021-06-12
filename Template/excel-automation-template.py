#!/usr/bin/env python

# AUTHOR: SVD
# DATE:   01 June 2021
# DESC:   This is a python template to open excel files

import openpyxl
import os
import math

# Execute for each python file
def run(fileName): 
  global wb
  global sheet 
  global intervals

  # Open excel file
  wb = openpyxl.load_workbook(fileName)
  sheet = wb.active
  max_rows = sheet.max_row

  print(f'\nFINISHED: {fileName}')

# Main, start of the program
if __name__ == "__main__":
  while True:
    # User input
    print('\n============================================')
    print('\nPlease make sure the excel files are closed. ')
    filePath = input('\nEnter full file path / folder name: ')
    try:
      os.chdir(filePath)
    except:
      print('Folder does not exist')
      break

    # Open the folder location and find the excel files
    excelFiles = os.listdir('.')
    print(f'\n{len(excelFiles)} excel files found.')

    # For each excel file
    for i in range(0, len(excelFiles)):
      if not excelFiles[i].endswith('.xlsx'):
        print('\tNot a valid folder for .xlsx files')
        break
      run(excelFiles[i])
      
    os.chdir("..")
