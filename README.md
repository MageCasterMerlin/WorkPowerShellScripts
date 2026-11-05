How to Install:

1. Select Code > download zip
2. Extract Files
3. Create Folder in Downloads Folder named "IX440 Report"
4. Put IX440 Script.ps1 in the folder
5. Right click on desktop and make a new shortcut 
6. Paste the following line into the field, MAKE SURE TO UPDATE TO YOUR USERNAME: 

powershell.exe -noexit -ExecutionPolicy Bypass -File "C:\Users\USERNAME\Downloads\IX440 Report\IX440 Script.ps1"


How To Use:

1. Download IX440 SPD, STORE and PX110 as .csv
2. Rename the 2 IX440 Files as "IX440 SPD" and "IX440 STORE"
3. Put All .csv files "IX440 Report" folder
4. Open previous IX440 in app from SharePoint (Needed for VLOOKUP)
5. Run script using the shortcut


What This Script Does:
1. Copies data from CSV files into new Excell workbook
2. Deletes unnecessary columns 
3. Moves columns to proper place
4. Runs text to column
5. Inserts VLOOKUP data
6. Saves IX440 report with todays date  
