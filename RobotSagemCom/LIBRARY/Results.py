######################################################################
#######         Developper : Mansouri WISSEM
#######         This script is developped to used in robot framework script
#######         To create and insert data into an xlsx file
#######         There is many way to use this script
#######         
######################################################################

import time
import  os
import  sys
from openpyxl.styles import Color, PatternFill, Font, Border, Side , Alignment
from openpyxl.worksheet.dimensions  import Dimension
from openpyxl import load_workbook , Workbook
import string

class Results():
    # Initialise the class
    def __init__(self , fileName=None , sheetName=None):
        # initialise same data
        self.fileName   =   fileName
        self.sheetName  =   sheetName
        print ("Call Result Class")
        # config sheet file :
        # add same style
        self.bold_font      = Font(bold=True)
        self.thin_border    = Border(left=Side(style='thin'), 
            right=Side(style='thin'), 
            top=Side(style='thin'), 
            bottom=Side(style='thin')
            )
        self.defaultFill    = PatternFill(start_color='d3d3d3', 
            end_color='d3d3d3', 
            fill_type='solid'
            )
        self.greenFill      = PatternFill(start_color='0cff0c',
            end_color='0cff0c',
            fill_type='solid'
            )
        self.center_aligned_text = Alignment(horizontal="center")
        self.listAlpha   =   string.ascii_uppercase
        self.numberList   =   ['0' , '1' , '2' , '3' , '4' , '5','6', '7' , '8' , '9']
    #create or open file
    def create_file(self):
        #if file exist open it
        if os.path.isfile(self.fileName):
            print ("Open an Existing File")
            myBook = load_workbook(filename=self.fileName)
        # create new file
        else:
            print ("Create a new file")
            myBook = Workbook()
        #Get all sheet in the file
        sheetsNames =   myBook.sheetnames
        print (sheetsNames)

        # crate the specific sheet file
        if self.sheetName == None:
            mySheet = myBook.create_sheet("Results")
        # Test if sheet in argv
        else:
            #test if sheet exist
            #open the existing sheet
            if self.sheetName in sheetsNames:
                print ("open existing sheet")
                mySheet = myBook[self.sheetName]
            #create a new sheet file
            else:
                mySheet = myBook.create_sheet(self.sheetName)

        #Delete default sheet name
        self.remove_sheet_by_name("Sheet" , myBook)

        # Create Sheet And Workbook
        self.mySheet    =   mySheet
        self.myBook     =   myBook

        #Save file
        self.save_data_in_file()

    # Function Save file
    def save_data_in_file(self):
        self.myBook.save(filename=self.fileName)

    # Function to delete a specific sheet in xlsx file
    def remove_sheet_by_name(self , sheetFineName , workbookName):
        sheetsNames =   workbookName.sheetnames
        if sheetFineName in sheetsNames:
            print ("Delete Sheet")
            workbookName.remove(workbookName[sheetFineName])

    # This function used to adjust a specific cell (col) in th e active sheet file
    def adjust_cell(self , cellName , cellSize):
        if self.mySheet[cellName] != None:
            self.mySheet.column_dimensions[self.mySheet[cellName].column_letter].width=cellSize
            self.save_data_in_file()

    # this function used to create the sheet header
    def create_header_file(self, *argsList):
        # Example usage: create_header_file("TestSuite" , "Summary"  , "Test Case" , "Results")
        key         =   0
        for value in argsList:
            print ("{}".format(value))
            strKey     =   self.listAlpha[key]+"1"
            self.insert_cell_header_data(strKey , value)
            key +=1
            

    # insert data in header cell
    def insert_cell_header_data(self, cellName , data):
        # There is a 3 methods to call this function
        # 1 :using full cell name example:   insert_cell_header_data("A1" , "data")
        # 2 :using first string col name example:   insert_cell_header_data("A" , "data")
        # 3 :using int cell name example:   insert_cell_header_data(1 , "data")

        
        strCelName  =   str(cellName)
        if len(strCelName) == 1 and strCelName[0] in self.listAlpha:
            cellNameKey =   strCelName+"1"
        elif strCelName[0] in self.listAlpha and strCelName[1] in self.numberList:
            cellNameKey =   strCelName
        else:
            key         =   cellName
            cellNameKey =   self.listAlpha[key]+"1"
            print (cellNameKey)
        
        self.mySheet[cellNameKey]              =   data
        self.mySheet[cellNameKey].border       =   self.thin_border
        self.mySheet[cellNameKey].fill         =   self.defaultFill
        self.mySheet[cellNameKey].font         =   self.bold_font
        self.mySheet[cellNameKey].alignment    =   self.center_aligned_text
        self.save_data_in_file()
        
    # insert data in specific cell
    def insert_cell_data(self, cellName , data):
        # insert data in specific cell
        # The cellName can be
        # complete name string  : like A1 , B5 
        # row : col : like      1:5         22:14
        # Example usage : ResultsObject.insert_cell_data("A5" , "OK")
        #                 ResultsObject.insert_cell_data("1:5" , "OK")
        # import : in the second way the numbere begin from 0
        strCellName =   str(cellName)

        if strCellName[0] in self.listAlpha:
            self.mySheet[strCellName]       =   data
        else:
            arrKey                          =   strCellName.split(':')
            firstKey                        =   int(arrKey[0])
            secondKey                       =   int(arrKey[1]) + 1
            cellData                        =   self.listAlpha[firstKey]+str(secondKey)
            self.mySheet[cellData]          =   data
        self.save_data_in_file()

    # insert a data by col name or col value
    def insert_col_data(self, cellName , data , appendData="False"):
        # insert a data by col name or col value
        # The cellName can be
        # string  : like A , B
        # integer : 1  ,   14 , 0
        # Example usage : ResultsObject.insert_col_data("A" , "OK")
        #                 ResultsObject.insert_col_data("1" , "OK")
        # import : in the second way the numbere begin from 0
        strCellName =   str(cellName)
        if strCellName in self.listAlpha:
            key             =   self.get_empty_cell(strCellName)
        else:
            strCellName     =   self.listAlpha[int(cellName)]
            key             =   self.get_empty_cell(strCellName)

        print (key)
        
        self.mySheet[key]   =   data
        self.save_data_in_file()

    # This function used to get the next cell to insert data
    def get_empty_cell(self, cel):
        strCel  =   str(cel)
        Testkey     =   1
        if len(strCel) == 1  and strCel in self.listAlpha:
            celName =   strCel
        elif strCel[0] in self.listAlpha and strCel[1] in self.numberList:
            celName =   strCel[0]
        else:
            intCel  =   int(cel)
            celName =   self.listAlpha[intCel]
            celName =   celName
        
        while True:
            celNameValue =   celName+str(Testkey)
            if self.mySheet[celNameValue].value == None:
                break
            Testkey +=1
        return celNameValue

    #insert data by header name
    def insert_conName_data(self, colNameData):
        pass


    ################## FU Tests ##########
    ##  FU : Boot
    ##  FU : Periodic Check
    ######################################
    # Create insert function in FU Test
    def write_head_fu_test_file(self):
        self.mySheet['A1']="TestSuite"
        self.mySheet['B1']="TestCase"
        self.mySheet['C1']="Summary"
        #self.mySheet['D1']="Preconditions"
        self.mySheet['D1']="Step"
        #self.mySheet['F1']="ExpectedResults"
        self.mySheet['E1']="Result"
        self.mySheet['F1']="Comments"
        #Set Border style , header Style
        for header in self.mySheet['A1:F1']:
            for col in header:
                col.border      =   self.thin_border
                col.fill        =   self.defaultFill
                col.font        =   self.bold_font
                col.alignment   =   self.center_aligned_text
        self.save_data_in_file()


    def adjust_test_file(self):
        self.adjust_cell('A1' , 20)
        self.adjust_cell('B1' , 20)
        self.adjust_cell('C1' , 20)
        self.adjust_cell('D1' , 45)
        self.adjust_cell('E1' , 15)
        self.adjust_cell('F1' , 45)

    def create_fu_test_result_file(self):
        self.write_head_fu_test_file()
        self.adjust_test_file()
        #self.save_data_in_file()

    def insert_fu_result(self):
        pass

if __name__ == "__main__":
    currentPath     =   os.getcwd()
    currentPath     =   currentPath.replace('\\' , '/')
    print (currentPath)

    pathFolder      =   currentPath+"/RobotSagemCom/testLibrary/"
    firstFile       =   pathFolder+"testInPythonScript.xlsx"
    firstResults    =   Results(firstFile , "Results")
    firstResults.create_file()


    #firstResults.insert_col_data('A' , 'Test col data')
    #firstResults.insert_col_data(2 , 'Test col data')
    #firstResults.insert_cell_data('1:2' , 'Test')
    #firstResults.insert_cell_header_data("A1" , "data")
    #firstResults.insert_cell_header_data("B" , "data")
    #firstResults.insert_cell_header_data(2 , "data")
    #keyHeader      =   {"A1":"First" , "B1" : "Second"}
    # 1 :using full cell name example:   insert_cell_header_data("A1" , "data")
    # 2 :using first string col name example:   insert_cell_header_data("A" , "data")
    # 3 :using int cell name example:   insert_cell_header_data(1 , "data")
    #firstResults.create_header_file("TestSuite" , "Summary"  , "Test Case" , "Results")

    """
    firstResults.create_file()
    firstResults.create_fu_test_result_file()
    firstResults.create_header_file("First" , "Second" , "Third" ,"Fourth")
    """