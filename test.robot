*** Settings ***
Library         Selenium2Library
*** Variable ***
${URL}                      https://opensource-demo.orangehrmlive.com/
${login_user}               Admin
${login_password}           admin123
${login_input_text}         xpath=//*[@id="txtUsername"]
${password_input_text}      xpath=//*[@id="txtPassword"]
${login_button}             xpath=//*[@id="btnLogin"]

*** Test Case ***
Test Case
    Open Browser                        ${URL}                          chrome
    Maximize Browser Window
    Sleep                               2s
    Input Text                          ${login_input_text}             ${login_user}
    Input Text                          ${password_input_text}          ${login_password}
    Click Element                       ${login_button}
    Sleep                               3s
    Click Element                       xpath=//*[@id="welcome"]
    Sleep                               1s
    Click Element                       xpath=//*[@id="welcome-menu"]/ul/li[2]/a
    Sleep                               2s
    Close Browser

*** Keywords ***
