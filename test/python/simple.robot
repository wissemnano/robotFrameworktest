*** Settings ***
Library             BuiltIn
Library             Collections
Variables            test.py
*** Variable ***
${test_var}             simple texte

*** Test Case ***
Log Var Test
    Log To Console                 ${test_var}
    Log To Console                 ${text}

*** Keywords ***
