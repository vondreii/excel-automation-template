#!/usr/bin/env python

# AUTHOR: SVD
# DATE:   01 June 2021
# DESC:   This is a python template to open excel files

import openpyxl
import os
import math

def clearTable(fileName, cols):
  for letter in cols:
    for row in range(1,max_rows+1):
      sheet[f'{letter}{row}'].value = ''
  wb.save(fileName)

# Execute for each python file
def run(fileName): 
  global wb
  global sheet 
  global intervals
  global max_rows

  # Open excel file
  wb = openpyxl.load_workbook(fileName)
  sheet = wb.active
  max_rows = sheet.max_row

  print(str(fileName)+": READING")

  # Car Location Data
  cols = ['E','F','G','H','I','J','K','L','M']
  clearTable(fileName, cols)
  print(str(fileName) + ": Cleaned Player Movement data")
  
  # Car Location Data
  cols = ['P','Q','S','T','U','V']
  clearTable(fileName, cols)
  print(str(fileName) + ": Cleaned Car Crash data")

  # Car rb Data
  cols = ['Y','Z','AA','AB']
  clearTable(fileName, cols)
  print(str(fileName) + ": Cleaned RB Crash data")

  # Response Time
  cols = ['AE','AF','AG','AJ']
  clearTable(fileName, cols)
  print(str(fileName) + ": Cleaned Response Time data")

  # Everything Else
  cols = ['AM','AN','AO','AP','AR','AS','AT','AV','AW','AX','AY','BA','BB','BC','BE','BF','BG','BH','BI','BJ','BK','BL','BN','BO']
  clearTable(fileName, cols)
  print(str(fileName) + ": Cleaned Obstacle Spawning Data")

  print(str(fileName)+": FINISHED")

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
