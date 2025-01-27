*** Settings ***
Library     Autosphere.Browser.Playwright   auto_close=${False}
Library     Autosphere.Excel.Files
Library     Autosphere.HTTP
Library     OperatingSystem
Library    String


*** Variables ***
${WEBSITE}   https://botsdna.com/server/
${EXCEL_FILE}  input.xlsx
${DOWNLOAD_LINK}    https://botsdna.com/server/input.xlsx
${DOWNLOAD_DIRECTORY}   ${CURDIR}\\File



*** Keywords ***
Open website and Download File
    Open Browser  url=${WEBSITE}
    Maximize Browser Window
    Wait Until Page Contains Element    //*[contains(text(),"Server Setup")]
#    Set Download Directory      ${DOWNLOAD_DIRECTORY}
    ${file_status}=  Run Keyword And Return Status  File Should Exist   ${DOWNLOAD_DIRECTORY}\\input.xlsx
    Run Keyword If    ${file_status} == False
    ...    Download    url=${DOWNLOAD_LINK}     target_file=${DOWNLOAD_DIRECTORY}      overwrite=True
#    Pause Execution


Read Data And Create Server
    Open Workbook    ${CURDIR}\\File\\${EXCEL_FILE}
    ${Active Sheet}=  Get Active Worksheet
    ${data}=  Read Worksheet As Table   name=${Active Sheet}      header=True
    LOG  ${data}
    FOR    ${row}    IN    @{data}
        ${request_id}=  Set Variable  ${row}[RequestID]
        ${os}=  Set Variable  ${row}[OS]
        ${os}=  Strip String    ${os}
        ${ram}=  Set Variable  ${row}[RAM]
        ${ram}=  Strip String    ${ram}
        ${hdd}=  Set Variable  ${row}[HDD]
        ${hdd}=  Strip String    ${hdd}
        ${application}=  Set Variable  ${row}[Applications]
        ${application}=  Strip String    ${application}
        @{apps}=  Split String  ${application}  separator=,
        ${len}=  Get Length    ${apps}

        Create Server  ${request_id}  ${os}  ${ram}  ${hdd}  ${apps}


    END
    Close Workbook


Create Server
    [Arguments]  ${request_id}  ${os}  ${ram}  ${hdd}  ${apps}
    LOG  ${request_id}
    LOG  ${os}
    LOG  ${ram}
    LOG  ${hdd}
    LOG  ${apps}
    ${os_selection_status}=  Set Variable  ${False}
    ${ram_selection_status}=  Set Variable  ${False}
    ${hdd_selection_status}=  Set Variable  ${False}
    ${apps_selection_status}=  Set Variable  ${False}

    Wait Until Page Contains Element    //*[contains(text(),"Server Setup")]
    Select From List By Label    //*[@id="os"]   ${os}
    ${selected_os}=  Get Value    //*[@id="os"]
    IF    '${selected_os}' == '${os}'
         ${os_selection_status}=  Set Variable  ${True}
    END
    
    Select From List By Label    //*[@id="Ram"]  ${ram}
    ${selected_ram}=  Get Value    //*[@id="Ram"]
    IF    '${selected_ram}' == '${ram}'
         ${ram_selection_status}=  Set Variable  ${True}
    END

    ${hdd_selection_status}=  Select Size  ${hdd}  ${hdd_selection_status}
    ${apps_selection_status}=  Select Applications   ${apps}   ${apps_selection_status}

    IF    ('${os_selection_status}' == '${True}') and ('${ram_selection_status}' == '${True}') and ('${hdd_selection_status}' == '${True}') and ('${apps_selection_status}' == '${True}')
         Click Element    (//*[(@type="button") and (@value="Create Server")])
         Wait Until Page Contains Element    //*[@id="serverIP"]   timeout=15s
         Go Back
         Wait Until Page Contains Element    //*[contains(text(),"Server Setup")]
         Reload Page
         Wait Until Page Contains Element    //*[contains(text(),"Server Setup")]


    END
Select Size
    [Arguments]    ${hdd}       ${hdd_status}
    ${radio_buttons_count}=  Get Element Count    //*[@id="hdd"]
    FOR    ${counter}    IN RANGE    1    ${radio_buttons_count}+1
        ${option}=  Get Text  (//*[@id="hdd"])[${counter}]//following-sibling::label[1]
        ${option}=  Strip String    ${option}
        IF    '${option}' == '${hdd}'
             Click Element    (//*[@id="hdd"])[${counter}]
             ${hdd_status}=  Set Variable  ${True}
             #Exit For Loop
        END

    END

    [Return]        ${hdd_status}

Select Applications
    [Arguments]     ${apps}     ${apps_status}
    ${len}=  Get Length    ${apps}
    ${checkbox_count}=  Get Element Count  (//*[@type="checkbox"])

    FOR    ${counter}    IN RANGE    0    ${len}
        ${app}=  Set Variable   ${apps}[${counter}]
        ${app}=  Strip String    ${app}
        LOG    ${app}
        FOR    ${count_2}    IN RANGE    1    ${checkbox_count} + 1
            Log    ${app}
            ${picked_app}=  Get Text    (//*[@type="checkbox"])[${count_2}]//following-sibling::label[1]
            ${picked_app}=  Strip String    ${picked_app}
            IF    '${picked_app}' == '${app}'
                 Click Element    (//*[@type="checkbox"])[${count_2}]
                 ${apps_status}=  Set Variable  ${True}
            END

        END

    END
    [Return]        ${apps_status}

*** Tasks ***
Server Creation
    Open website and Download File
    Read Data And Create Server