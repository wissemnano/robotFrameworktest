*** Settings ***
Library             custom.py
Library             SeleniumLibrary
Library             BuiltIn
Library             Collections
*** Variable ***

*** Test Case ***
My Test
    hello       wissem
    ${first} =  FunctionName    10  10
    Log To Console              ${first}

*** Keywords ***
