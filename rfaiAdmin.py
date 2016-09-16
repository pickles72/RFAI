#! /usr/bin/env python3

# Script to administrate RFAI database
# last revision: Sep 2016, Michael Fruhnert

## INSTRUCTIONS:
## 1)  use along with rfai.py and rfaitoolbox.py
##
## Path to Modify category: modify question - modify category

#import all necessary libraries
from openpyxl import load_workbook
from os import system
from rfaitoolbox import selectQuestionMaster, rearrangeCategories
from rfaitoolbox import replaceQuestion, collectRFAI, normalizeMaster
from rfaitoolbox import findRFAIFolder
import basicUtils as basics

doMore = True
system('clear')
rfaiFolder = findRFAIFolder(True) # assume that we work one at a time
if (rfaiFolder[0] != ''):
    print(rfaiFolder[1] + ' set as default.\n')

while(doMore):
    print('This is the RFAI administration tool. ' \
            'Use with extreme caution!\n')

    doThis = input("Choose from options below:\n"\
                "c: re-arrange categories\n" \
                "q: modify questions / answers\n" \
                "n: normalize master (ASCII conformity)\n" \
                "r: replace question (mark for deletion)\n" \
                "w: collect files and write responses\n" \
                "0: exit\n" \
                "\n")
    
    system('clear')
    if (doThis == 'c'):
        rearrangeCategories()
    elif (doThis == 'q'):
        selectQuestionMaster()
    elif (doThis == 'n'):
        normalizeMaster()
    elif (doThis == 'r'):
        replaceQuestion(rfaiFolder[0])
    elif (doThis == 'w'):
        collectRFAI(rfaiFolder[0])
    elif (doThis == '0'):
        print("RFAI administration tool shut-down...\n")
        doMore = False
    else:
        print("Invalid input. Enter 0 to exit.\n")
