######################################################################
#######         Developper : Mansouri WISSEM
#######         This script is developped to used in robot framework script
#######         To create and insert data into an xlsx file
#######         There is many way to use this script
#######         
######################################################################

import  os
import  sys
from openpyxl.styles import Color, PatternFill, Font, Border, Side , Alignment
from openpyxl.worksheet.dimensions  import Dimension
from openpyxl import load_workbook , Workbook


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
    def create_header_file(self, **kwargs):
        for key , value in kwargs.items():
            print ("{} == {}".format(key , value))
            strKey     =   str(key)
            self.insert_data_in_cell(strKey , value)

    # insert data in specific cell
    def insert_data_in_cell(self, cellName , data):
        self.mySheet[cellName]          =   data
        self.mySheet[cellName].border   =   self.thin_border
        self.save_data_in_file()

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
    #keyHeader      =   {"A1":"First" , "B1" : "Second"}

    firstResults.create_file()
    firstResults.create_fu_test_result_file()
    #firstResults.create_header_file(A1="First" , B1="Second")