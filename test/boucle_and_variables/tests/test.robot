*** Settings ***
Library             BuiltIn
Library             Collections
*** Variables ***
${var1}             "wissem"


*** Test Cases ***
Test01
    Log         ${var1}
    Log         ${CURDIR}
Create Liste
    ${list} =       Create List         a   b   c   d 
    ${tmp} =        Get From List       ${list}             0
    Log             ${tmp}
Create Dictionnaire
    ${dict}=        Create Dictionary       a   1   b   2
    Log             ${dict}
    ${tmp} =        Get From Dictionary           ${dict}         a
    Log             ${tmp}
*** Keywords ***