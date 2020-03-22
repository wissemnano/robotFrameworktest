*** Settings ***
Library         Selenium2Library
Library         C:\\Users\\MANSOURI\\Desktop\\workspace\\RobotSagemCom\\MyLibrary.py    WITH NAME   FirstLib
Library         C:\\Users\\MANSOURI\\Desktop\\workspace\\RobotSagemCom\\writeResults.py
#Library         MyLibrary.py     ${test}  
*** Variable ***
#${test}                 MyLibrary
${fileName}             test.log
${Step}                 step1 , step2 jjfjvfvjf , hferff , rfirefirf ,freferf
*** Test Case ***
Test
    Log                             this is a test
    writeTestSuiteInfo              xlsxfile=xlsxfile.xlsx      TestSuite=FU    TestCase=Case       Summary=None
    writeTestSuiteInfo              xlsxfile=xlsxfile.xlsx      Step=${Step}
    
*** Keywords ***
