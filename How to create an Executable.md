# validationTool

How to create exe using PyInstaller

1. Install PyInstaller using pip
    pip install pyinstaller
    
2. Create a folder to keep the generated exe and all the intermediate files. 
3. Goto the folder created above and open cmd 
5. run the PyInstaller from cmd prompt
    pyinstaller --onefile path/to/the/source file/main.py
    where main.py is the file from where execution of your script starts.
    This will create a 2 folders 'build' and 'dist' along with a file, main.spec
    
    You may have to repeat this step many times until you get the exe to run. Use the batch file included, PyInstaller.bat.
    Right click and 'Run as Administrator' to run the pyinstaller. 
    Runing this batch will delete all the contents of the folder and generate the folders 'build', 'dist' and the *.spec file
    Remember to change the folder path in the batch file, when you change your folder. 
6. Once the batch file has been run and the exe has been generated, you are only half way through. The exe will not run as of now,
   because of some missing dependencies.
7. To fix that you need to make some changes in the .spec file and re-run the pyinstaller
8. This is how you do it. Open the .spec file, goto the line hiddenimports= [] and change it to this
   hiddenimports=['six','pkg_resources._vendor.packaging','pkg_resources._vendor.packaging.version',
                  'pkg_resources._vendor.packaging.specifiers','pkg_resources._vendor.packaging.requirements',
                  'pyparsing'],
   These are the missing dependencies. The pyinstaller is unable to resolve these dependencies, and it needs to be resolved manually.
9. Once the changes have been made, goto the folder where the .spec file is present, delete the 'build' and 'dist' folder and 
    run the pyinstaller, now with the spec file
    pyinstaller main.spec
10. The exe generated now should run without any issues.
11. You will find your exe in the dist folder
