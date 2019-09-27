*** Settings ***
Library     Selenium2Library

*** Variables ***

${URL}                  https://opensource-demo.orangehrmlive.com/
${userPath}             //*[@id="txtUsername"]
${userNameCorrectly}    Admin
${userNameFailed}       test
${loginPath}            //*[@id="txtPassword"]
${loginFailed}          admin
${loginCorrectly}       admin123
${logBtn}               //*[@id="btnLogin"]
${orangeDashboard}      https://opensource-demo.orangehrmlive.com/index.php/dashboard
${FailedUrl}            https://opensource-demo.orangehrmlive.com/index.php/auth/validateCredentials


*** Test Cases ***
Log In Test Orange
    Open Orange Test Web Page
    Login Should Be Failed
    Login Should Run Successfly
    Close Orange Page


*** Keywords ***
Open Orange Test Web Page
    Open Browser        ${URL}   chrome
    Maximize Browser window

Login Should Run Successfly
    Input Text              xpath=${userPath}       ${userNameCorrectly}
    Input Text              xpath=${loginPath}      ${loginCorrectly}
    Click Element           xpath=${logBtn}
    Location Should Be      ${orangeDashboard}

Login Should Be Failed
    Input Text              xpath=${userPath}       ${userNameFailed}
    Input Text              xpath=${loginPath}      ${loginFailed}
    Click Element           xpath=${logBtn}
    Location Should Be      ${FailedUrl}

Close Orange Page
    sleep       2s
    close browser