*** Settings ***
Library     Autosphere.Browser.Playwright  auto_close=${FALSE}
Library     Autosphere.HTTP
Library    Autosphere.Excel.Files
Library    String

*** Keyword ***
Goto Website And Download Excel File
    Set Download Directory  ${CURDIR}
    Download  url=https://botsdna.com/notaries/AP-ADVOCATES.xlsx  overwrite=True  target_file=${CURDIR}
    Open Available Browser  url=https://botsdna.com/notaries/
    Maximize Browser Window
    Wait Until Page Contains Element    //h1[contains(text(),"FILL NOTARY ADVOCATE Details")]  timeout=15s

*** Keyword ***
Read Excel and Fill Notary Details
    Open Workbook  ${CURDIR}//AP-ADVOCATES.xlsx
    ${district}=  Set Variable  ${EMPTY}
    ${Advocate_name}=  Set Variable  ${EMPTY}
    ${area}=  Set Variable  ${EMPTY}

    ${empty_row}=  Find Empty Row
    FOR    ${row}    IN RANGE    2    ${empty_row}
        Log    ${row}
        ${s_no}=  Get Cell Value    ${row}    A
        ${type}=  Evaluate    type($s_no)
        ${is string}=   Evaluate     isinstance($s_no, str)
        IF    ${is string}
            ${district}=  Set Variable  ${s_no}
            IF    "DIST" in """${district}"""
                 ${district}=  Replace String  ${district}  DIST  ${EMPTY}
            END
            Strip String    ${district}
        ELSE IF  "${district}" != "${EMPTY}"
            ${Advocate_name}=  Get Cell Value    ${row}    B
            ${area}=  Get Cell Value    ${row}    C
        END

        LOG  ${district}
        LOG  ${Advocate_name}
        LOG  ${area}
    END
#    Input Text  //*[@id="notary"]
#    Input Text  //*[@id="area"]
    Close Workbook
*** Tasks ***
Notary Services
    Goto Website And Download Excel File
    Read Excel and Fill Notary Details