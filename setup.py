import subprocess
import sys
import os
import shutil

try:
    print ('Installing required modules...')
    choice = input('Do you wish to continue [y/n] :')
    if choice != 'y':
        sys.exit('Terminating the process')
    else: 

        subprocess.call('easy_install pip', shell=True)
        subprocess.call('pip install pandas', shell=True)
        subprocess.call('pip install numpy', shell=True)
        subprocess.call('pip install xlsxwriter', shell=True)
        subprocess.call('pip install xlrd', shell=True)
        subprocess.call('pip install xlwt', shell=True)
        subprocess.call('pip install jinja2', shell=True)
        subprocess.call('pip list', shell=True)

        print ('Installation completed successfully...')

        subprocess.call('pip list', shell=True)
             
        print ('Creating the folder structure..')

        currDirName = os.path.dirname(os.path.realpath(__file__))
        mainFolder =os.path.join(currDirName, 'Data_Comparison_Tool')
        inputFolder = os.path.join(mainFolder, 'Input')
        outputFolder = os.path.join(mainFolder, 'Output')
        sasFileFoler = os.path.join(inputFolder, 'SAS_File')
        pythonFileFoler = os.path.join(inputFolder, 'Python_File')

        try:
            os.mkdir(mainFolder)
            os.mkdir(inputFolder)
            os.mkdir(outputFolder)
            os.mkdir(sasFileFoler)
            os.mkdir(pythonFileFoler)

            print ('Directories created.')

            print ('Placing the master script inside the parent folder..')
            os.chdir(currDirName)
            print ('the current directory is ' + currDirName)
        except:
            print ('Folder structure already present')

        try:
            print ('Placing the master script inside the parent folder..')
            os.chdir(currDirName)
            print ('the current directory is ' + currDirName)
            shutil.move('DataComparisionScript.py', mainFolder)
            print ('File moved..')
        except:
            print ('Could not place the file..kindly copy it to ' + mainFolder)

        print ('DataComaprisonTool setup completed.')


except:
           print ('Ending program.... setup aborted')
