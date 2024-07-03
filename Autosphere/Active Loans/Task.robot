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
Perform Excel Work
    Open Workbook  ${CURDIR}\\Loan Data\\input.xlsx
    ${empty_row}=  Find Empty Row
    FOR    ${row}    IN RANGE    2    ${empty_row}
        Log    ${row}
        ${account_number}=  Get Cell Value    ${row}    A
        Strip String  ${account_number}
    END
    Close Workbook

*** Task ***
Active Loans
    Goto Website And Download Excel File
    Perform Excel Work
