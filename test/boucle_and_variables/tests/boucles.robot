*** Settings ***
Library         BuiltIn
Library         Collections


*** Variables ***

*** Test Cases ***
Test Case 01
    My Keyword 01
    My Keyword 02

*** Keywords ***
My Keyword 01
    :FOR    ${animal}   IN      Chat Chien
    \       Log         ${animal}
    \       Log         Animal Enregistrer
My Keyword 02  
    ${ma_liste} =       Create List         Chat    Chien
    :FOR        ${item}         IN      @{ma_liste}*
    \           Log             ${item}