*** Settings ***
Library     Autosphere.Browser.Playwright  auto_close=${FALSE}
Library     Autosphere.Excel.Files
Library     Autosphere.FileSystem
Library     Autosphere.Archive
Library     Autosphere.HTTP
Library     OperatingSystem
Library     Collections
Library     String


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
DownLoad And Extract Data from File
    [Documentation]     This keyword download the zip file,
    ...                 Unzip, Extract required data from text file
    ...                 And return values in dictionary

    [Arguments]     ${acct_num}  ${code}
    &{detail}=  Create Dictionary
    Set To Dictionary  ${detail}  Bank            ${EMPTY}
    Set To Dictionary  ${detail}  Branch          ${EMPTY}
    Set To Dictionary  ${detail}  Loan Date     ${EMPTY}
    Set To Dictionary  ${detail}  Amount          ${EMPTY}
    Set To Dictionary  ${detail}  EMI             ${EMPTY}

    ${extract_status}=  Run Keyword And Return Status  Download And Unzip File  ${acct_num}  ${code}
    IF    ${extract_status}
        LOG  Text file extract successfully. Now get required data from text file
        ${content}=  Get File    ${CURDIR}\\Loan Data\\${acct_num}.txt
        LOG  ${content}
        @{lines}=  Split To Lines    ${content}
        LOG  ${lines}
        FOR    ${line}    IN    @{lines}
            Log     ${line}
            IF    "Bank" in """${line}"""
                 @{values}=  Split String    ${line}  separator=:
                 ${bank_name}=  Set Variable  ${values}[1]
                 ${bank_name}=  Strip String    ${bank_name}
                 Set To Dictionary    ${detail}  Bank=${bank_name}
                 LOG  ${detail}
            END

            IF    "Branch" in """${line}"""
                 @{values}=  Split String    ${line}  separator=:
                 ${branch_name}=  Set Variable  ${values}[1]
                 ${branch_name}=  Strip String    ${branch_name}
                 Set To Dictionary    ${detail}  Branch=${branch_name}
                 LOG  ${detail}
            END

            IF    "Loan Taken On" in """${line}"""
                 @{values}=  Split String    ${line}  separator=:
                 ${loan_date}=  Set Variable  ${values}[1]
                 ${loan_date}=  Strip String    ${loan_date}
                 Set To Dictionary    ${detail}  Loan Date=${loan_date}
                 LOG  ${detail}
            END

            IF    "Amount" in """${line}"""
                 @{values}=  Split String    ${line}  separator=:
                 ${amount}=  Set Variable  ${values}[1]
                 ${amount}=  Strip String    ${amount}
                 Set To Dictionary    ${detail}  Amount=${amount}
                 LOG  ${detail}
            END

            IF    "EMI" in """${line}"""
                 @{values}=  Split String    ${line}  separator=:
                 ${emi}=  Set Variable  ${values}[1]
                 ${emi}=  Strip String    ${emi}
                 Set To Dictionary    ${detail}  EMI=${emi}
                 LOG  ${detail}
            END
        END
        OperatingSystem.Remove Files  ${CURDIR}\\Loan Data\\${acct_num}.txt
    END
    LOG  ${detail}
    [Return]    ${detail}


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

        &{account_detail}=  DownLoad And Extract Data from File  ${account_number}  ${loan_code}
        LOG  ${account_detail}

        Set Cell Value    ${row}    B    ${account_detail}[Bank]
        Set Cell Value    ${row}    C    ${account_detail}[Branch]
        Set Cell Value    ${row}    D    ${account_detail}[Loan Date]
        Set Cell Value    ${row}    E    ${account_detail}[Amount]
        Set Cell Value    ${row}    F    ${account_detail}[EMI]
        Set Cell Value    ${row}    G    ${pan_number}
        Set Cell Value    ${row}    H    ${status}

        Save Workbook
    END
    Close Workbook

*** Task ***
Active Loans
    Goto Website And Download Excel File
    Perform Excel Work
