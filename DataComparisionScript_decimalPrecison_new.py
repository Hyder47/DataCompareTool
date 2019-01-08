import numpy as np
import pandas as pd
import xlsxwriter as xl
import os
import sys
import re
from collections import OrderedDict

dirName = os.path.dirname(os.path.realpath(__file__))
sasPath = os.path.join(dirName,'Input\\SAS_File\\')
pythonPath = os.path.join(dirName,'Input\\Python_File\\')
outputPath = os.path.join(dirName,'Output\\')

#Get the list of SAS Files
os.chdir(sasPath)
sasFiles = os.listdir(sasPath)
#Get the list of Python Files
os.chdir(pythonPath)
pythonFiles = os.listdir(pythonPath)

#Matching SAS and Python files
commonFiles = np.intersect1d(sasFiles, pythonFiles)

# function to check both SAS and Python files are present.
def fileChecker(fileName):
    returnValue = False
    os.chdir(sasPath)
    sasFile = os.path.exists(fileName)
    os.chdir(pythonPath)
    pythonFile = os.path.exists(fileName)
    if (sasFile == True) & (pythonFile == True):
        returnValue = True
    
    return returnValue

# function to colour False values.
def colour_func(val):
    color = 'red' if val == False else 'black'
    return 'color: %s' % color

for i in commonFiles:
    fileName = i
    fileCheck = fileChecker(i)
    if fileCheck == True:
        print ('Comparing ' + fileName)
        dfSas = pd.read_excel(sasPath + fileName)
        dfPython = pd.read_excel(pythonPath + fileName)
        cols = list(dfSas.columns)
        dfSas.sort_values(by = [cols[0]], inplace = True)
        dfSas.reset_index(drop = True, inplace = True)
        dfPython.sort_values(by = [cols[0]], inplace = True)
        dfPython.reset_index(drop = True, inplace = True)
           
        finalDF = pd.DataFrame()
        output = OrderedDict() # this ordered dictionary stores the comparision data. 

        if dfSas.shape == dfPython.shape:
            print ('Level 1 match passed.')
            if dfSas.columns.all() == dfPython.columns.all():
                print ('Level 2 match passed.')
                cols = dfSas.columns
                for i in cols:
                    if (dfSas[i].dtype == 'float64') & (dfPython[i].dtype == 'float64'):
                        dfSas[i] = dfSas[i].astype(str)
                        dfPython[i] = dfPython[i].astype(str)
                        matchSas = re.findall('\.\d{1,8}', str(dfSas[i]))
                        matchPython = re.findall('\.\d{1,8}', str(dfPython[i]))
                        val_compare = [True if matchSas[x]== matchPython[x] else False for x in range(len(matchSas))]
                        output[i] = val_compare
                                                                                      
                    else:
                        value_compare = list(dfSas[i] == dfPython[i])
                        output[i] = value_compare

                finalDF = pd.DataFrame(output)
                finalDF = finalDF.style.applymap(colour_func)
                workbook = xl.Workbook(outputPath + fileName +'.xlsx')
                worksheet1 = workbook.add_worksheet('SAS')
                worksheet2 = workbook.add_worksheet('Python')
                worksheet3 = workbook.add_worksheet('Compare')

                print ('Excel schema generated.')

                writer = pd.ExcelWriter(outputPath + fileName +'.xlsx')
                dfSas.to_excel(writer,sheet_name='SAS',index = False)
                dfPython.to_excel(writer,sheet_name='Python',index = False)
                finalDF.to_excel(writer,sheet_name='Compare', index = False)
                workbook.close()
                writer.save()
                
                print(fileName + ' comparision completed.')

            else:
                print (fileName + ' Level 1 match failed.')
        else:
            print (fileName + ' Level 2 match failed.')
            
    else:
        sys.exit('Terminating program. Files not found. Place the files and re run the script.')
