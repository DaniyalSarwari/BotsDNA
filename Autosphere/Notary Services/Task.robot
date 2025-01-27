*** Settings ***
Library     Autosphere.Browser.Playwright  auto_close=${FALSE}
Library     Autosphere.HTTP
Library    Autosphere.Excel.Files
Library    String

*** Variables ***
${EXCEL_FILE}   AP-ADVOCATES.xlsx


*** Keyword ***
Goto Website And Download Excel File
    Set Download Directory  ${CURDIR}
    #Download  url=https://botsdna.com/notaries/AP-ADVOCATES.xlsx  overwrite=True  target_file=${CURDIR}
    Open Browser  url=https://botsdna.com/notaries/
    Maximize Browser Window
    Wait Until Page Contains Element    //h1[contains(text(),"FILL NOTARY ADVOCATE Details")]  timeout=15s

*** Keyword ***
Fill Notary Detail
    [Arguments]  ${dist}  ${adv_name}  ${area}
    Input Text  //*[@id="notary"]  ${adv_name}
    Input Text  //*[@id="area"]  ${area}
    Select From List By Label    //select[@id="DIST"]  ${dist}
    ${selected_value}=  Get Value    //select[@id="DIST"]
    Click Element    //*[@value="Submit Notary"]
    Wait Until Page Contains Element    //*[contains(text(),"Transaction Number")]  timeout=15s
    ${transaction_number}=  Get Text    //*[contains(text(),"Transaction Number")]//p[@id="TransNo"]
    Go Back
    Wait Until Page Contains Element    //h1[contains(text(),"FILL NOTARY ADVOCATE Details")]  timeout=15s
    [Return]  ${transaction_number}

*** Keyword ***
Read Excel and Fill Notary Details
    Open Workbook  ${CURDIR}\\${EXCEL_FILE}
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
            ${district}=  Strip String    ${district}
            Log  District Row found. Skipping Row
            Continue For Loop
        ELSE IF  "${district}" != "${EMPTY}"
            ${Advocate_name}=  Get Cell Value    ${row}    B
            ${area}=  Get Cell Value    ${row}    C
        END
        
        ${district}=  Strip String    ${district}
        ${Advocate_name}=  Strip String    ${Advocate_name}
        ${area}=  Strip String    ${area}
        LOG  ${district}
        LOG  ${Advocate_name}
        LOG  ${area}
        
        IF    ("""${district}""" != "${EMPTY}") and ("""${Advocate_name}""" != "${EMPTY}") and ("""${area}""" != "${EMPTY}")
            ${trans_num}=  Fill Notary Detail  ${district}  ${Advocate_name}  ${area}
            Set Cell Value    ${row}    D    ${trans_num}
            Save Workbook
        END
    END
#    Input Text  //*[@id="notary"]
#    Input Text  //*[@id="area"]
    Close Workbook
*** Tasks ***
Notary Services
    Goto Website And Download Excel File
    Read Excel and Fill Notary Details