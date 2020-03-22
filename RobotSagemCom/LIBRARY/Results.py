import  os
import  sys
from openpyxl.styles import Color, PatternFill, Font, Border, Side
from openpyxl.worksheet.dimensions  import Dimension
from openpyxl import load_workbook , Workbook


class Results():
    def __init__(self , fileName=None , sheetName=None):
        self.fileName   =   fileName
        self.sheetName  =   sheetName
        print ("Call Result Class")

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
        self.removeSheetByName("Sheet" , myBook)

        # Create Sheet And Workbook
        self.mySheet    =   mySheet
        self.myBook     =   myBook
        
        #Save file
        self.saveDataInFile()
    
    # Function Save file
    def saveDataInFile(self):
        self.myBook.save(filename=self.fileName)
    
    def removeSheetByName(self , sheetFineName , workbookName):
        sheetsNames =   workbookName.sheetnames
        if sheetFineName in sheetsNames:
            print ("Delete Sheet")
            workbookName.remove(workbookName[sheetFineName])

if __name__ == "__main__":
    pathFolder  = "C:/Users/MANSOURI/Desktop/workspace/RobotSagemCom/testLibrary/"
    firstFile   =   pathFolder+"firstResult.xlsx"
    firstResults = Results(firstFile , "Results")
    firstResults.create_file()