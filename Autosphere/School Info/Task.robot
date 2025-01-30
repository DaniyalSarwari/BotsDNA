*** Settings ***
Library    Autosphere.Browser.Playwright
Library    Autosphere.Excel.Files
Library    Autosphere.FileSystem
Library    Autosphere.HTTP
Library    OperatingSystem
Library    String

*** Variables ***
${URL}      https://botsdna.com/school/
${DOWNLOAD_URL}     https://botsdna.com/school/Master%20Template.xlsx
${DOWNLOAD_DIRECTORY}       ${CURDIR}\\File
${MASTER_FILE}      Master Template.xlsx

*** Keywords ***
Open Portal And Download Master File
    Open Browser  url=${URL}
    Maximize Browser Window
    Wait Until Page Contains Element    //h1[contains(text(),"School Database")]
    Wait Until Element Is Visible    //*[@id="SchoolCode"]

    ${file_status}  Run Keyword And Return Status   File Should Exist  ${DOWNLOAD_DIRECTORY}\\${MASTER_FILE}
    Run Keyword If    not ${file_status}    Download Master File

Download Master File
    Create Directory  ${DOWNLOAD_DIRECTORY}
    Empty Directory    ${DOWNLOAD_DIRECTORY}
    Download    url=${DOWNLOAD_URL}  target_file=${DOWNLOAD_DIRECTORY}
    ${status1}=  Run Keyword And Return Status    Wait Until Keyword Succeeds  2  2s  File Should Exist    ${DOWNLOAD_DIRECTORY}\\${MASTER_FILE}
    IF    ${status1} == ${False}
        ${status2}=  Run Keyword And Return Status    Wait Until Keyword Succeeds  2  2s  File Should Exist    ${DOWNLOAD_DIRECTORY}\\Master%20Template.xlsx
        IF    ${status2} == ${True}
             Move File    ${DOWNLOAD_DIRECTORY}\\Master%20Template.xlsx    ${DOWNLOAD_DIRECTORY}\\${MASTER_FILE}
        END
    END


*** Keywords ***
Get School Code and Fill School Data
    Open Workbook  ${DOWNLOAD_DIRECTORY}\\${MASTER_FILE}
    ${last_row}=  Find Empty Row
    FOR    ${row}    IN RANGE    2    ${last_row}
        ${school_code}=  Get Cell Value    ${row}    A
        ${type}=  Evaluate    type(${school_code})
        IF    "str" in "${type}"
             ${school_code}  Strip String  ${school_code}
             ${school_code}  Convert To Integer    ${school_code}
        END
        Wait Until Element Is Visible    //*[@id="SchoolCode"]
        Input Text    //*[@id="SchoolCode"]    ${school_code}
        Click Element    //*[@id="SearchSchool"]
        Sleep    3s

        Switch To specific Tab  1
        ${status}=  Run Keyword And Return Status  Wait Until Element Is Visible    //table
#        IF    ${status}
#             Fill School Data in Excel
#
#        END
        Close Window
        Switch To specific Tab  0
        Sleep  1s

    END
    Close Workbook

Switch To specific Tab
    [Arguments]     ${tab}
    @{handles}  Get Window Handles
    ${len}=  Get Length    ${handles}
    Switch Window    ${handles}[${tab}]


*** Keywords ***
Fill School Data in Excel


*** Tasks ***
School Database
    Open Portal And Download Master File
    Get School Code and Fill School Data
