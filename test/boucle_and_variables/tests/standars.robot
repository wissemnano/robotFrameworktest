*** Settings ***
Library     BuiltIn
Library     Collections
*** Variables ***
*** Test Cases ***
Test Case 01
    Log         "Test"
    Comment     "Comment Test"
    Sleep       2s
Test Case 02
    :FOR        ${number}           IN RANGE            1   10  1
    \           Run Keyword IF      ${number}==5        Log To Console      ${number}
    ...     ELSE        Log To Console      "Test"

*** Keywords ***