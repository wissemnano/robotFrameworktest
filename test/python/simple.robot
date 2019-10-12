*** Settings ***
Library             BuiltIn
Library             Collections

*** Variable ***
${test_var}             simple texte

*** Test Case ***
Log Var Test
    Log                 ${test_var}

*** Keywords ***
