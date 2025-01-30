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
        IF    ${status}
             Fill School Data in Excel  ${row}
        END
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
    [Arguments]  ${row}

    ${school_name}=  Get Text    //center//h1
    ${school_address}=  Get Text    //td//b[contains(text(),"School Address")]//..//following-sibling::td
    ${phone_number}=  Get Text    //td//b[contains(text(),"School Phonenumber")]//..//following-sibling::td
    ${strength}=  Get Text    //td//b[contains(text(),"Student's Strenth")]//..//following-sibling::td
    ${principle_name}=  Get Text    //td//b[contains(text(),"Prncipal Name")]//..//following-sibling::td  
    ${teaching_staff}=  Get Text    //td//b[contains(text(),"Number of TeachingStaff")]//..//following-sibling::td
    ${non_teaching_staff}=  Get Text    //td//b[contains(text(),"Number of Non-TeachingStaff")]//..//following-sibling::td
    ${numof_school_buses}=  Get Text    //td//b[contains(text(),"Number of School buses")]//..//following-sibling::td
    ${playground}=  Get Text    //td//b[contains(text(),"School Playground")]//..//following-sibling::td
    ${facilities}=  Get Text    //td//b[contains(text(),"Facilities")]//..//following-sibling::td
    ${accrediation}=  Get Text    //td//b[contains(text(),"School Accrediation")]//..//following-sibling::td
    ${hostel}=  Get Text    //td//b[contains(text(),"School Hostel")]//..//following-sibling::td
    ${canteen}=  Get Text    //td//b[contains(text(),"School Canteen")]//..//following-sibling::td
    ${stationary}=  Get Text    //td//b[contains(text(),"School Stationary")]//..//following-sibling::td
    ${teaching_method}=  Get Text    //td//b[contains(text(),"School Teaching method's")]//..//following-sibling::td
    ${school_timing}=  Get Text    //td//b[contains(text(),"School Timing")]//..//following-sibling::td
    ${achievements}=  Get Text    //td//b[contains(text(),"School Achivements")]//..//following-sibling::td
    ${awards}=  Get Text    //td//b[contains(text(),"School Awards")]//..//following-sibling::td
    ${uniform}=  Get Text    //td//b[contains(text(),"School Uniform")]//..//following-sibling::td
    ${school_type}=  Get Text    //td//b[contains(text(),"School type")]//..//following-sibling::td

    Set Cell Value    ${row}    B    ${school_name}
    Set Cell Value    ${row}    C    ${school_address}
    Set Cell Value    ${row}    D    ${phone_number}
    Set Cell Value    ${row}    E    ${strength}
    Set Cell Value    ${row}    F    ${principle_name}
    Set Cell Value    ${row}    G    ${teaching_staff}
    Set Cell Value    ${row}    H    ${non_teaching_staff}
    Set Cell Value    ${row}    I    ${numof_school_buses}
    Set Cell Value    ${row}    J    ${playground}
    Set Cell Value    ${row}    K    ${facilities}
    Set Cell Value    ${row}    L    ${accrediation}
    Set Cell Value    ${row}    M    ${hostel}
    Set Cell Value    ${row}    N    ${canteen}
    Set Cell Value    ${row}    O    ${stationary}
    Set Cell Value    ${row}    P    ${teaching_method}
    Set Cell Value    ${row}    Q    ${school_timing}
    Set Cell Value    ${row}    R    ${achievements}
    Set Cell Value    ${row}    S    ${awards}
    Set Cell Value    ${row}    T    ${uniform}
    Set Cell Value    ${row}    U    ${school_type}

    Save Workbook

*** Tasks ***
School Database
    Open Portal And Download Master File
    Get School Code and Fill School Data
