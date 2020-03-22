*** Settings ***


*** Variable ***
${xlsxLibrary}                      LIBRARY\\Results.py
${firstFile}                        testLibrary\\firstTestLib.xlsx
${secondFile}                       testLibrary\\secondTestLib.xlsx
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

