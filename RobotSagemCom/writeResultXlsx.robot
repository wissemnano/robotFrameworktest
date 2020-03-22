*** Settings ***


*** Variable ***
${xlsxLibrary}                      C:\\Users\\MANSOURI\\Desktop\\workspace\\RobotSagemCom\\LIBRARY\\Results.py
${firstFile}                        C:\\Users\\MANSOURI\\Desktop\\workspace\\RobotSagemCom\\testLibrary\\firstTestLib.xlsx
${secondFile}                       C:\\Users\\MANSOURI\\Desktop\\workspace\\RobotSagemCom\\testLibrary\\secondTestLib.xlsx
*** Test Case ***
Example
    Import Library                  ${xlsxLibrary}      ${firstFile}        WITH NAME   FirstLib
    Log To Console                  Create object results
    Log                             Create object results
    Import Library                  ${xlsxLibrary}      ${secondFile}       WITH NAME   SecondLib
    Log To Console                  Create object results
    Log                             Create object results

    FirstLib.Create File            
    SecondLib.Create File           

*** Keywords ***

