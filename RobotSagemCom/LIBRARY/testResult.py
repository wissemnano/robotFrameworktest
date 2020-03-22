import  xlwt
import  xlrd

pathFolder  =   "C:/Users/MANSOURI/Desktop/workspace/RobotSagemCom/testLibrary/xlsxfile.xlsx"
myBook      =   xlrd.open_workbook(pathFolder)
mySheet     =   myBook.sheet_by_index(0)
print (mySheet)