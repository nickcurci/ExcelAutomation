# -- =============================================
# -- Author:		<Curci, Nicholas>
# -- Create date: <06/27/2021>
# -- Description:	<Runs the Spreadsheet Audit in it's entirety
#                   including; checking for duplicates within each sheet
#                   checking for duplicates across sheets
#                   checking for consistency from Resonant to the Excel sheet
#                   checking for consistency from LifeSuite to the Excel Sheet>
# -- JIRA Ticket #: xxxxxxxxxxxxxxxxxxxxx
# -- SSRS Report: ?MU Audit?
# -- Notes: Need to make a master variable for each table as a raw input
#           so that I do not have to use pd.read_excel every time(i.e. MIB)
#           Creating a raw variable will make life easier for every function
#           Will likely reduce runtime and code length. Also reduce chance of failure
#           # get rid of misc folders with mkdir
# -- =============================================

## These are the inputs needed to run the code, the warnings have been turned off for ease of reading in the output
## If not installed already please install pip, pandas, os, pyodbc, pandasql, numpy, tkinter
import pandas as pd
import os
import pyodbc
import pandasql as ps
import numpy as np
from tkinter.filedialog import askopenfilename
from tkinter import *
import tkinter as tk
from tkinter import simpledialog
import datetime
import warnings
warnings.filterwarnings("ignore")

## This is the main function that will be run, additional functions will be run from within
def Spreadsheet_Audit():

    ## This function opens a pop-up window that allows you to select your input file.
    ## In this case it is the MU_Audit.xlsx file
    def selectFile():
            ## Declare filename
            global filename # The global name before a variable makes it useable across functions
            Tk().withdraw()  # we don't want a full GUI, so keep the root window from appearing
            filename = askopenfilename()  # show an "Open" dialog box and return the path to the selected file
            print("This excel file has been selected", filename)
    selectFile() ## This piece runs the function


    ## This function calls the begin and run dates from the input box
    def selectBeginAndRunDate():
        global beginDate_input
        global runDate_input
        global today
        global current_day
        current_day = datetime.date.today() #set the current day
        today = datetime.date.strftime(current_day, "%m/%d/%Y") # chenge the format of the day to match what is expected in SQL Server
        ROOT = tk.Tk() # sets the root for tkinter
        ROOT.withdraw() # allows us to withdraw from tkinter

        # this piece, which is repeated for the other variable, is an input box for a string, which is the format we need for the date
        beginDate_input = simpledialog.askstring(title="Begin Date",
                                           prompt="Enter the begin date in MM/DD/YYY format:",
                                                 initialvalue="01/01/2020")
        runDate_input = simpledialog.askstring(title="Run Date",
                                               prompt="Enter the start date in MM/DD/YYY format:",
                                               initialvalue=today)
    selectBeginAndRunDate()


    ## This function sets the output directories needed, it does this by allowing the user to select them themselves.
    def getPaths():
        global outputDir_input # setting of global variables again
        global miscDir_input
        global procDir_input
        root = tk.Tk() # setting of root and withdraw again
        root.withdraw()
        # these three lines below open file explorer windows where folders can be selected as directories
        outputDir_input = tk.filedialog.askdirectory(title="Select where you would like to store the Final output files")
        miscDir_input = tk.filedialog.askdirectory(title="Select where you would like to store misc files")
        procDir_input = tk.filedialog.askdirectory(title="Select where you would like to store the output of the SSRS Stored Procedures")
    getPaths()


    ## This function defines variables as global since we are using varaibles created from different functions
    ## We then define them below.
    def GlobalVariables():
        global inputDirectory
        global DuplicatesFoundInSameSheet
        global DuplicatesFoundAcrossSheets
        global ResonantAndLifeSuite
        global AllExcel
        global AllExcel_Res
        global All_LifeSuite
        global RunDate
        global BeginDate
        global outputDirectory
        global miscDirectory
        global writer
        global procDirectory
        # below is the just setting of variable names or locations
        miscDirectory = miscDir_input
        procDirectory = procDir_input
        inputDirectory = filename
        DuplicatesFoundInSameSheet = r'DuplicatesFoundInSameSheet.xlsx'
        DuplicatesFoundAcrossSheets = r'DuplicatesFoundAcrossSheets.xlsx'
        ResonantAndLifeSuite = r'ResonantAndLifeSuite.xlsx'
        AllExcel = 'AllExcel.xlsx'
        AllExcel_Res = 'AllExcel_Res.xlsx'
        All_LifeSuite = 'All_Lifesuite.xlsx'
        RunDate = runDate_input
        BeginDate = beginDate_input
        outputDirectory = outputDir_input
        # the writer function below allows us to write to excel, we will need it in the future, it is a global variable
        writer = pd.ExcelWriter(os.path.join(outputDirectory, ResonantAndLifeSuite))
    GlobalVariables()

    ## This functuon defines the tables that we will create an use across the code
    def defineTables():
        global MIB
        MIB = pd.read_excel(os.path.join(inputDirectory),  # opens the workbook
                            sheet_name='xxxxxxxxxxxx',  # specifies the correct sheer
                            usecols='A',  # specifies which column to look at
                            header=1)  # sets the header equal to a specific row
        MIB['SheetName'] = 'xxxxxxxxxxxx'  # names the sheet in the to be created dataframe
        MIB = pd.DataFrame(MIB, index=None) # creates the dataframe

        global IN4655
        IN4655 = pd.read_excel(os.path.join(inputDirectory),
                            sheet_name= 'xxxxxxxxxxxx',
                            usecols='E',
                            header=1)
        IN4655['SheetName'] = 'xxxxxxxxxxxx'
        IN4655 = pd.DataFrame(IN4655, index=None)

        global Exceptions
        Exceptions = pd.read_excel(os.path.join(inputDirectory),
                            sheet_name= 'xxxxxxxxxxxx',
                            usecols='E',
                            header=1)
        Exceptions['SheetName'] = 'xxxxxxxxxxxx'
        Exceptions = pd.DataFrame(Exceptions, index=None)

        global LT2020_MU_Input
        LT2020_MU_Input = pd.read_excel(os.path.join(inputDirectory),
                            sheet_name= 'xxxxxxxxxxxx',
                            usecols='E',
                            header=1)
        LT2020_MU_Input['SheetName'] = 'xxxxxxxxxxxx'
        LT2020_MU_Input = pd.DataFrame(LT2020_MU_Input, index=None)

        global MU_Input
        MU_Input = pd.read_excel(os.path.join(inputDirectory),
                            sheet_name= 'xxxxxxxxxxxx',
                            usecols='E',
                            header=2)
        MU_Input['SheetName'] = 'MU_Input'
        MU_Input = pd.DataFrame(MU_Input, index=None)

        global WL2020_Input
        WL2020_Input = pd.read_excel(os.path.join(inputDirectory),
                            sheet_name= 'xxxxxxxxxxxx',
                            usecols='E',
                            header=1)
        WL2020_Input['SheetName'] = 'xxxxxxxxxxxx'
        WL2020_Input = pd.DataFrame(WL2020_Input, index=None)

        global Withdrawn
        Withdrawn = pd.read_excel(os.path.join(inputDirectory),
                            sheet_name= 'xxxxxxxxxxxx',
                            usecols='E',
                            header=1)
        Withdrawn['SheetName'] = 'xxxxxxxxxxxx'
        Withdrawn = pd.DataFrame(Withdrawn, index=None)
    defineTables()

    ## This function takes the MU Audit spreadsheet (the original excel file) and finds the duplicates within each sheet.
    def DupesBySheet():
        # The following code block is similar for each sheet
        MIB_dupes = MIB[MIB.duplicated()] # creates a new dataframe MIB_dupes with only records that are duplicated in the sheet
        MIB_dupes = pd.DataFrame(MIB_dupes, index = None) # sends this new data to a dataframe to be interacted with later

        IN4655_dupes = IN4655[IN4655.duplicated()]
        IN4655_dupes = pd.DataFrame(IN4655_dupes, index = None)


        Exceptions_dupes = Exceptions[Exceptions.duplicated()]
        Exceptions_dupes = pd.DataFrame(Exceptions_dupes, index = None)


        LT2020_MU_Input_dupes = LT2020_MU_Input[LT2020_MU_Input.duplicated()]
        LT2020_MU_Input_dupes = pd.DataFrame(LT2020_MU_Input_dupes, index = None)


        MU_Input_dupes = MU_Input[MU_Input.duplicated()]
        MU_Input_dupes = pd.DataFrame(MU_Input_dupes, index=None)


        WL2020_Input_dupes = WL2020_Input[WL2020_Input.duplicated()].dropna()
        WL2020_Input_dupes = pd.DataFrame(WL2020_Input_dupes, index = None)


        Withdrawn_dupes = Withdrawn[Withdrawn.duplicated()]
        Withdrawn_dupes = pd.DataFrame(Withdrawn_dupes, index = None)

        # after performing the ETL on each sheet, the writer opens and sets the path to Duplicates Found In The Same Sheet
        writer = pd.ExcelWriter(os.path.join(outputDirectory,DuplicatesFoundInSameSheet))
        # Since this is a file we need, we send it to the output directory
        # The lines below send the data to the workbook "writer" and to the specific sheet mentioned afterwards
        # This way we have all of the data in one workbook but on multiple sheets for ease of use outside of python
        MIB_dupes.to_excel(writer, 'xxxxxxxxxxxx')
        IN4655_dupes.to_excel(writer, 'xxxxxxxxxxxx')
        Exceptions_dupes.to_excel(writer, 'xxxxxxxxxxxx')
        LT2020_MU_Input_dupes.to_excel(writer, 'xxxxxxxxxxxx')
        MU_Input_dupes.to_excel(writer, 'xxxxxxxxxxxx')
        WL2020_Input_dupes.to_excel(writer, 'xxxxxxxxxxxx')
        Withdrawn_dupes.to_excel(writer, 'xxxxxxxxxxxx')
        # This command saves the workbook
        writer.save()
        # just a print statement urging the user to check that the file exists and is accurate since it cannot be produced inside the console
        print("Please check the Duplicates Found In Same Sheet excel file to determine if duplicates were found within the same sheet")
    DupesBySheet()


    ## This function takes the MU Audit spreadsheed(the same original excel file) and finds the duplicates across sheets.
    def AllDupes():
        # Each of these code blocks are similar in the way they ETL data
        # This block reads the data from the OG workbook, then specifies the sheet name, column letter, and header line
        # we dont really need to re-read this in, since we already have the dataframes already
        # but for speed, this way was faster, especially when testing in different files

    #############################################################################
    ### This is where the fun stuff happens
    ### Here we will merge each dataframe that we just created together
    ### We will do this one by one until each table is merged to to each other
    ### kind of like a tree
    ### we will use the pd.merge method called on policy number as an inner join
    ### the inner join is used so we only get matches between tables
    #############################################################################

        MIB_To_IN46 = pd.merge(MIB, IN4655, on='xxxxxxxxxxxx Number', how='inner')
        # print(MIB_To_IN46)

        MIB_To_Ex = pd.merge(MIB, Exceptions, on='xxxxxxxxxxxx Number', how='inner')
        # print(MIB_To_Ex)

        MIB_To_LT = pd.merge(MIB, LT2020_MU_Input, on='xxxxxxxxxxxx Number', how='inner')
        # print(MIB_To_LT)

        MIB_To_MU = pd.merge(MIB, MU_Input, on='xxxxxxxxxxxx Number', how='inner')
        # print(MIB_To_MU)

        MIB_To_WL = pd.merge(MIB, WL2020_Input, on='xxxxxxxxxxxx Number', how='inner')
        # print(MIB_To_WL)

        MIB_To_WD = pd.merge(MIB, Withdrawn, on='xxxxxxxxxxxx Number', how='inner')
        # print(MIB_To_WD)

        ###

        IN46_To_Ex = pd.merge(IN4655, Exceptions, on='xxxxxxxxxxxx Number', how='inner')
        # print(IN46_To_Ex)

        IN46_To_LT = pd.merge(IN4655, LT2020_MU_Input, on='xxxxxxxxxxxx Number', how='inner')
        # print(IN46_To_LT)

        IN46_To_MU = pd.merge(IN4655, MU_Input, on='xxxxxxxxxxxx Number', how='inner')
        # print(IN46_To_MU)

        IN46_To_WL = pd.merge(IN4655, WL2020_Input, on='xxxxxxxxxxxx Number', how='inner')
        # print(IN46_To_WL)

        IN46_To_WD = pd.merge(IN4655, Withdrawn, on='xxxxxxxxxxxx Number', how='inner')
        # print(IN46_To_WD)

        ###

        Ex_To_MU = pd.merge(Exceptions, MU_Input, on='xxxxxxxxxxxx Number', how='inner')
        # print(Ex_To_MU)

        Ex_To_WL = pd.merge(Exceptions, WL2020_Input, on='xxxxxxxxxxxx Number', how='inner')
        #  print(Ex_To_WL)

        Ex_To_WD = pd.merge(Exceptions, Withdrawn, on='xxxxxxxxxxxx Number', how='inner')
        # print(Ex_To_WD)

        Ex_To_LT = pd.merge(Exceptions, LT2020_MU_Input, on='xxxxxxxxxxxx Number', how='inner')
        # print(Ex_To_LT)

        ###

        LT_To_WL = pd.merge(LT2020_MU_Input, WL2020_Input, on='xxxxxxxxxxxx Number', how='inner')
        # print(LT_To_WL)

        LT_To_WD = pd.merge(LT2020_MU_Input, Withdrawn, on='xxxxxxxxxxxx Number', how='inner')
        # print(LT_To_WD)

        LT_To_MU = pd.merge(LT2020_MU_Input, MU_Input, on='xxxxxxxxxxxx Number', how='inner')
        # print(LT_To_MU)


        ###

        MU_To_WD = pd.merge(MU_Input, Withdrawn, on='xxxxxxxxxxxx Number', how='inner')
        # print(MU_To_WD)

        MU_To_WL  = pd.merge(MU_Input, WL2020_Input, on='xxxxxxxxxxxx Number', how='inner')
        # print(MU_To_WL)

        ###

        WL_To_WD = pd.merge(WL2020_Input, Withdrawn, on='xxxxxxxxxxxx Number', how='inner')
        # print(WL_To_WD)

        #################################################################
        ######## Here we concat every merge table we just made into one
        ######## This will show us every policy number match across all the
        ######## the sheets in the original excel file
        ######## we then send this table to a dataframe called AllDupes
        #################################################################
        global AllDupes
        AllDupes = pd.concat([MIB_To_WL,MIB_To_LT,MIB_To_WD, MIB_To_MU,  MIB_To_Ex, MIB_To_IN46, IN46_To_MU,
                              IN46_To_WL,IN46_To_WD,IN46_To_LT,IN46_To_Ex, Ex_To_LT, Ex_To_WD, Ex_To_WL, Ex_To_MU,
                              LT_To_MU,LT_To_WD, LT_To_WL, MU_To_WL, MU_To_WD, WL_To_WD], axis=0, join='outer')
        AllDupes = pd.DataFrame(AllDupes, index = None)

        # After the dataframe of matches is created, we search to see if it is empty
        # if it is empty we know that there are no matches between sheets
        if AllDupes.empty:
            print("There were no duplicate records found across sheets")
        else:
            print("There were duplicate records found across sheets. They can be found in the Duplicates Found Across Sheets excel file: \n"
                  "An example of some the records found are: \n", AllDupes.head())
        # we call the writer again to write the results of the concatenated dataframe to excel
        # we send it to the output directory and call it Duplicates Across Sheets
        # we then save the workbook in the writer.save() command
        writer = pd.ExcelWriter(os.path.join(outputDirectory,DuplicatesFoundAcrossSheets))
        AllDupes.to_excel(writer, 'DuplicatesAcrossSheets')
        writer.save()
    AllDupes()


    ## To connect to SQL Server, we need to create pyodbc connections
    ## this will allow us to read the stored procs
    def Connection():
        # Here the connection is created by setting up the driver, server, database, and a trusted connection since I do not have specific credentials
        # There is a spearate connection for lifesuite and resonant since they are stored procs in different databases
        LS_connection = pyodbc.connect(driver='{SQL Server}', server='xxxxxxxxxxxxxxx', database='xxxxxxxxxxxxxxxxx', trusted_connection='yes')
        Res_connection = pyodbc.connect(driver='{SQL Server}', server='xxxxxxxxxxxxxxx', database='xxxxxxxxxxxxxxxxxxxxx', trusted_connection='yes')

        # Here we call the stored procedures, the ? represent the parameters, we define them below
        LS = "{call dbo.xxxxxxxxxxxx(?,?)}"
        Res = "{call dbo.xxxxxxxxxxxx(?)}"

        # here we define the parameters for each stored proc
        LS_params = (RunDate,BeginDate,)
        Res_params = (RunDate,)

        # global variables again since we need them in different functions
        global MU_Stored_Proc_LS
        global MU_Stored_Proc_Res

        # here is where we actually call the stored procs
        # we use the read sql command where the sql, con, and params = our variables we created above
        # we use a different line for each stored proc.
        # the results of these stored procs are saved into the variables below
        # we will create dataframes out of them next
        MU_Stored_Proc_LS = pd.read_sql(sql=LS, con=LS_connection, params=LS_params)
        MU_Stored_Proc_Res = pd.read_sql(sql=Res, con=Res_connection, params=Res_params)
    Connection()


    # To use the data that we just gathered from the stored procs, we need to call it
    def StoredProcResults():
        global LS # global variables again
        global Res
        # here is the file that we would like to save these results into, we have not specified a filepath yet
        ResProcResults = 'ResProcResults.xlsx'
        LSProcResults = 'LSProcResults.xlsx'

        # here is where we turnt he lifesuite stored proc results into a dataframe
        # we set LS = to the stored proc results and only use the policy number column
        LS = MU_Stored_Proc_LS['policy_number']
        LS = pd.DataFrame(LS, index = None) # we create the dataframe
        LS['row_num'] = np.arange(len(LS)) # we set an arbitrary colum = to sequential numbers for the length of the dataframe
        LS = LS.set_index(['row_num']) # we use these number as the index to ensure that it is formatted correctly
        LS['policy_number'] = LS['policy_number'].str.replace(" ", "") # we replace any spaces with nothing to ensure data continuity
        LS.to_excel(os.path.join(procDirectory,LSProcResults)) # we send the results to excel

        # the same process is repeated for the Resonant stored proc results
        Res = MU_Stored_Proc_Res['xxxxxxxxxxxx']
        Res = pd.DataFrame(Res, index = None)
        Res['row_num'] = np.arange(len(Res))
        Res = Res.set_index(['row_num'])
        Res['xxxxxxxxxxxx'] = Res['xxxxxxxxxxxx'].str.replace(" ", "")
        Res.to_excel(os.path.join(procDirectory,ResProcResults))
    StoredProcResults()


    # This function finds every single policy number in the original MU Audit excel file
    def AllExcelPols():
        # Similar to above, we call each sheet into its own dataframe.
        # we dont really need to re-read this in, since we already have the dataframes already
        # but for speed, this way was faster, especially when testing in different files

        MIB_AEP = MIB['xxxxxxxxxxxx Number']
        MIB_AEP = pd.DataFrame(MIB_AEP, index = None)

        IN4655_AEP = IN4655['xxxxxxxxxxxx Number']
        IN4655_AEP = pd.DataFrame(IN4655_AEP, index = None)

        Exceptions_AEP = Exceptions['xxxxxxxxxxxx Number']
        Exceptions_AEP = pd.DataFrame(Exceptions_AEP, index = None)

        LT2020_MU_Input_AEP = LT2020_MU_Input['xxxxxxxxxxxx Number']
        LT2020_MU_Input_AEP = pd.DataFrame(LT2020_MU_Input_AEP, index=None)

        MU_Input_AEP = MU_Input['xxxxxxxxxxxx Number']
        MU_Input_AEP = pd.DataFrame(MU_Input_AEP, index=None)

        WL2020_Input_AEP = WL2020_Input['xxxxxxxxxxxx Number']
        WL2020_Input_AEP = pd.DataFrame(WL2020_Input_AEP, index=None)

        Withdrawn_AEP = Withdrawn['xxxxxxxxxxxx Number']
        Withdrawn_AEP = pd.DataFrame(Withdrawn_AEP, index=None)


        ## setting global variables
        global ExcelPols
        global EP
        global EPRes
        global ExcelPols_LS
        # creating one dataframe out of the tables created from the sheets
        ExcelPols= pd.concat([MIB_AEP, IN4655_AEP,Exceptions_AEP,LT2020_MU_Input_AEP,MU_Input_AEP,WL2020_Input_AEP,Withdrawn_AEP])
        ExcelPols = pd.DataFrame(ExcelPols, index=None)
        # resetting the index
        ExcelPols.reset_index()
        # renaming the column for consistency purposes
        ExcelPols.columns = ["xxxxxxxxxxxx"]
        ExcelPols = ExcelPols.dropna() # dropping null values
        ExcelPols_Res = ExcelPols
        ExcelPols_Res.to_excel(os.path.join(miscDirectory,AllExcel_Res)) # sending this to excel
        # reading it back in, no idea why i had to do this but i did and it worked
        # can probably go back and fix this
        ExcelPols_Res = pd.read_excel(os.path.join(miscDirectory,AllExcel_Res),
                                  sheet_name='Sheet1',
                                  usecols='B',
                                  header=0)
        ExcelPols_Res = pd.DataFrame(ExcelPols_Res, index=None) # creates the dataframe
        EPRes=ExcelPols_Res # renames it with a copy


        #########################################
        ExcelPols['row_num'] = np.arange(len(ExcelPols)) # adds the arbitrary row number
        ExcelPols = ExcelPols.set_index(['row_num']) # sets that as the index
        ExcelPols['xxxxxxxxxxxx'] = ExcelPols['xxxxxxxxxxxx'].str.replace(" ", "") # replaces spaces with blanks
        ExcelPols.to_excel(os.path.join(miscDirectory,'ExPols.xlsx')) # sents that to excel
        EP = ExcelPols # copies it
    AllExcelPols()


    # This function uses SQL to create a new table based on left joins
    # i.e. what exists in table 1 (lifesuite stored proc) that does not exists in table 2 (MU Audit spreadsheet)
    def PolsInLifeNotInExcel():
        global ExistsInLifeAndNotInExcel
        # sets a variable = to the result of this SQL which is a left join where the value is null in table 2
        ExistsInLifeAndNotInExcel = ps.sqldf("""SELECT * FROM LS t1
                                                LEFT JOIN EP t2 on t1.xxxxxxxxxxxx = t2.xxxxxxxxxxxx
                                                WHERE t2.xxxxxxxxxxxx IS NULL""")

        ExistsInLifeAndNotInExcel = pd.DataFrame(ExistsInLifeAndNotInExcel, index = None) # sends results to a dataframe
        ExistsInLifeAndNotInExcel = ExistsInLifeAndNotInExcel['xxxxxxxxxxxx'] # makes the new data = only to the one column

        ExistsInLifeAndNotInExcel.to_excel(writer, 'ExistsInLifeAndNotInExcel')  # sends that data to excel

        # if the table is empty and there are no misses then
        if ExistsInLifeAndNotInExcel.empty:
            print("There are no policy numbers that exist in Lifesuite that do not exist in Excel")

        else:
            print("There are policy numbers that exist in Lifesuite but not in excel. They are: \n", ExistsInLifeAndNotInExcel)
    PolsInLifeNotInExcel()


    # This is a replication of the above function but for resonant this time
    def PolsInResNotInExcel():
        global ExistsInResAndNotInExcel
        ExistsInResAndNotInExcel = ps.sqldf("""SELECT * FROM Res t1 
                                                LEFT JOIN EPRes t2 ON t2.xxxxxxxxxxxx = t1.xxxxxxxxxxxx 
                                                WHERE t2.xxxxxxxxxxxx IS NULL""")
        ExistsInResAndNotInExcel = pd.DataFrame(ExistsInResAndNotInExcel, index = None)
        ExistsInResAndNotInExcel = ExistsInResAndNotInExcel['xxxxxxxxxxxx']
        ExistsInResAndNotInExcel.to_excel(writer, 'ExistsInResAndNotInExcel')
        writer.save()
        if ExistsInResAndNotInExcel.empty:
            print("There are no policy numbers that exist in Resonant that do not exist in Excel")

        else:
            print("There are policy numbers that exist in resonant but not in excel. They are: ", ExistsInResAndNotInExcel)
    PolsInResNotInExcel()

# This is the end
Spreadsheet_Audit()
