# Excel Spreadsheet Automation

## This Repo will hold excel automation projects
#### Currently listed is a spreadsheet audit
##### This code:
*	First checks for duplicates within each sheet of the Audit Spreadsheet 
*	Then check for duplicates across each sheet
*	Then checks for consistency between Resonant database and the excel sheet
*	Then checks for consistency between LifeSuite database and the excel sheet
*	Returns each of these checks to their own excel file with an if/then condition to determine if there were records found

#### All confidential variables have been removed as this code works with sensitive data
#### This code is not meant to be run, only to be used as a reference to showcase understanding and ability

# Function details
#### Spreadsheet_Audit: This is the main bracket for the code, the code will be executed from within here
#### SelectFile: This function opens a pop-up window that allows you to select your input file. In this case it is the MU_Audit.xlsx file
#### SelectBeginAndRunDate: This function calls the begin and run dates from the input box
#### GetPaths: This function sets the output directories needed, it does this by allowing the user to select them themselves.
#### GlobalVariables: his function defines variables as global since we are using variables created from different functions. We then define them below.
#### DefineTables: This function defines the tables that we will create an use across the code.
#### DupesBySheet: This function takes the MU Audit spreadsheet (the original excel file) and finds the duplicates within each sheet.
#### AllDupes: This function takes the MU Audit spreadsheet (the same original excel file) and finds the duplicates across sheets.
#### Connection: To connect to SQL Server, we need to create pyodbc connections, this will allow us to read the stored procs
#### StoredProcResults: To use the data that we just got from the stored procs, we need to call it
#### AllExcelPols: This function finds every single policy number in the original MU Audit excel file
#### PolsInLifeNotInExcel: This function uses SQL to create a new table based on left joins. i.e. what exists in table 1 (lifesuite stored proc) that does not exists in table 2 (MU Audit spreadsheet)
#### PolsInResNotInExcel: This is a replication of the above function but for resonant this time
