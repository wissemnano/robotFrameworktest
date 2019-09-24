*** Settings ***
Library     Selenium2Library
*** Variables ***
${URL}      https://www.google.com
${BROWSER}          chrome
${txtSearch}        robot framework
${InputText}        //*[@id="tsf"]/div[2]/div[1]/div[1]/div/div[2]/input
${btnSearch}        //*[@id="tsf"]/div[2]/div[1]/div[3]/center/input[1]

*** Test Cases ***
Search Test Text
    Open Chrome
    Searching For Text
    CLose Chrome Browser
    
*** Keywords ***
Open Chrome
    [Documentation]     Test Case more Info here
    [Tags]      Acceptances
    open browser    ${URL}  ${BROWSER}
    Maximize Browser Window
Searching For Text
    Input Text      xpath=${InputText}  ${txtSearch}
    Click Element        xpath=${btnSearch}
    sleep       1s
CLose Chrome Browser
    close browser