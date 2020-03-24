*** Settings ***


*** Variable ***
${xlsxLibrary}                      C:\\Users\\MANSOURI\\Desktop\\workspace\\robotFrameworktest\\RobotSagemCom\\LIBRARY\\Results.py
${firstFile}                        C:\\Users\\MANSOURI\\Desktop\\workspace\\robotFrameworktest\\RobotSagemCom\\testLibrary\\firstTestLib.xlsx
${secondFile}                       C:\\Users\\MANSOURI\\Desktop\\workspace\\robotFrameworktest\\RobotSagemCom\\testLibrary\\secondTestLib.xlsx
${firstLib}                         FirstLib
${secondLib}                        SecondLib

*** Test Case ***
Example
    Init Xlsx Result File           ${firstFile}                ${firstLib}
    
*** Keywords ***

Init Xlsx Result File
    [Arguments]                     ${fileName}                 ${firstLib}
    Import Library                  ${xlsxLibrary}              ${fileName}        WITH NAME   ${firstLib}
    Log To Console                  Create object results
    Log                             Create object results
    ${firstLib}                     Create File          