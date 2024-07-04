*** Settings ***
Library     Autosphere.Browser.Playwright  auto_close=${FALSE}
Library     Autosphere.Excel.Files
Library     Autosphere.HTTP
Library     String

*** Keyword ***
Goto Website And Download Excel File
    Set Download Directory  ${CURDIR}//Loan Data
    Download  url=https://botsdna.com/ActiveLoans/input.xlsx  overwrite=True  target_file=${CURDIR}//Loan Data
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
Perform Excel Work
    Open Workbook  ${CURDIR}//Loan Data//input.xlsx
    ${empty_row}=  Find Empty Row
    FOR    ${row}    IN RANGE    2    ${empty_row}
        Log    ${row}

        ${account_number}=  Get Cell Value    ${row}    A
        ${loan_code}=  Extract Loan Code  ${account_number}
        LOG  ${loan_code}

        ${status}=  Get Loan Status  ${loan_code}
        LOG  ${status}



#        //a[contains(text(),"7325")]
    END
    Close Workbook

*** Task ***
Active Loans
    Goto Website And Download Excel File
    Perform Excel Work
