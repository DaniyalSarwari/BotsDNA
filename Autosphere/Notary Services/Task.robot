*** Settings ***
Library     Autosphere.Browser.Playwright  auto_close=${FALSE}
Library     Autosphere.HTTP

*** Keyword ***
Goto Website And Download Excel File
    Set Download Directory  ${CURDIR}
    Download  url=https://botsdna.com/notaries/AP-ADVOCATES.xlsx  overwrite=True  target_file=${CURDIR}
    Open Available Browser  url=https://botsdna.com/notaries/
    Maximize Browser Window
    Wait Until Page Contains Element    //h1[contains(text(),"FILL NOTARY ADVOCATE Details")]  timeout=15s

*** Keyword ***
Read Excel and Fill Notary Details

*** Tasks ***
Notary Services
    Goto Website And Download Excel File
    Read Excel and Fill Notary Details