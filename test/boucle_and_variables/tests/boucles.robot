*** Settings ***
Library         BuiltIn
Library         Collections


*** Variables ***

*** Test Cases ***
Test Case 01
    [Tags]      First Test
    My Keyword 01
    My Keyword 02
Test Case 02
    [Tags]      Second Test
    Boucle_Range
Test Case 03
    [Tags]      Test Boucle If
    Run_If_Test_case

*** Keywords ***
My Keyword 01
    :FOR    ${animal}   IN      Chat Chien
    \       Log         ${animal}
    \       Log         Animal Enregistrer
My Keyword 02  
    ${ma_liste} =       Create List         Chat    Chien
    :FOR        ${item}         IN      @{ma_liste}*
    \           Log             ${item}

Boucle_Range
    :FOR       ${number}        IN RANGE    2   10   1
    \          Log              ${number}


Run_If_Test_case
    [Documentation]     Run an If Test Boucle
    :FOR                ${number}           IN RANGE        1   10  1
    \                   Run Keyword IF      ${number}/2 == 0    Log         ${number}