*** Settings ***
Library     Autosphere.Browser.Playwright  auto_close=${FALSE}
Library     Autosphere.Excel.Files
Library     Autosphere.FileSystem
Library     Autosphere.HTTP
Library     OperatingSystem
Library     String
Library     Autosphere.Archive

*** Keyword ***
Goto Website And Download Excel File
    Set Download Directory  ${CURDIR}\\Loan Data
    OperatingSystem.Empty Directory    ${CURDIR}\\Loan Data
    Download  url=https://botsdna.com/ActiveLoans/input.xlsx  overwrite=True  target_file=${CURDIR}\\Loan Data
    Open Available Browser  url=https://botsdna.com/ActiveLoans/
    Maximize Browser Window
    Wait Until Page Contains Element    //h1[contains(text(),"Active Loans")]  timeout=15s

*** Keyword ***
Extract Loan Code
    [Documentation]  This keyword accept account number and extract last four digit which is a loan code
    [Arguments]  ${acct_num}
    Strip String  ${acct_num}
    ${length}=  Get Length    ${acct_num}
    ${last_four_digit}=  Get Substring    ${acct_num}   ${length-4}  ${length}
    [Return]  ${last_four_digit}

*** Keyword ***
Get Loan Status
    [Arguments]  ${code}
    ${status}=  Get Text  (//a[contains(text(),"${code}")]/../preceding-sibling::td)[1]
    [Return]  ${status}

*** Keyword ***
Get Pan Number
    [Arguments]  ${code}
    ${pan}=  Get Text  (//a[contains(text(),"${code}")]/../following-sibling::td)[1]
    [Return]  ${pan}

*** Keyword ***
Download And Unzip File
    [Documentation]     This Keyword only download, extract and remove zip file
    [Arguments]     ${acct_num}  ${code}
    Click Element    (//a[contains(text(),"${code}")])[1]
    ${file_status}=  Set Variable  ${False}
    FOR    ${counter}    IN RANGE    1    100
        ${files}=  OperatingSystem.List Files In Directory    ${CURDIR}\\Loan Data
        LOG  ${files}
        ${file_status}=  Does File Exist  ${CURDIR}\\Loan Data\\${acct_num}.zip
        IF    ${file_status}
             LOG  File downloaded successfully
             LOG  ${files}
             Exit For Loop
        ELSE
            Continue For Loop
        END
    END
    IF    ${file_status}
         Extract Archive  ${CURDIR}\\Loan Data\\${acct_num}.zip  ${CURDIR}\\Loan Data
         Sleep    0.3s
         OperatingSystem.Remove Files  ${CURDIR}\\Loan Data\\${acct_num}.zip
    END


*** Keyword ***
DownLoad And Extract Data from Zip File
    [Documentation]     This keyword download the zip file, unzip, extract required data from text file and return values in dictionary
    [Arguments]     ${acct_num}  ${code}
    &{detail}=  Create Dictionary
    ${extract_status}=  Run Keyword And Return Status  Download And Unzip File  ${acct_num}  ${code}
    IF    ${extract_status}
         #This keyword read data from text file
    END




*** Keyword ***
Perform Excel Work
    [Documentation]     This is the main keyword In this keyword bot get the values from website
    ...                 Download and extract the data from text file
    ...                 And write in excel file
    
    Open Workbook  ${CURDIR}//Loan Data//input.xlsx
    ${empty_row}=  Find Empty Row
    FOR    ${row}    IN RANGE    2    ${empty_row}
        Log    ${row}

        ${account_number}=  Get Cell Value    ${row}    A
        ${loan_code}=  Extract Loan Code  ${account_number}
        LOG  ${loan_code}

        ${status}=  Get Loan Status  ${loan_code}
        LOG  ${status}

        ${pan_number}=  Get Pan Number  ${loan_code}
        LOG  ${pan_number}

        DownLoad And Extract Data from Zip File  ${account_number}  ${loan_code}


#        //a[contains(text(),"7325")]
    END
    Close Workbook

*** Task ***
Active Loans
    Goto Website And Download Excel File
    Perform Excel Work
