*** Settings ***
Documentation   Locator Challange
Library  Autosphere.Browser.Playwright
Library  Autosphere.Excel.Files
# Library  Autosphere.MSExcel
Library  Autosphere.HTTP

*** Keyword ***
Download Excel File
    Wait Until Page Contains Element  //a[contains(text(),"Output.xlsx")]
    Download  https://botsdna.com/locator/Locator-Output.xlsx  target_file=${CURDIR}  overwrite=True


*** Keyword ***
Navigate to Locator Challenge
    Open Available Browser   url=https://www.botsdna.com/   maximized=True
    Wait Until Page Contains Element  //td//p
    Wait Until Element Is Visible  //td//p
    Click Element  ((//*[@id="rcorners2"])[1]//..)[1]
    Wait Until Page Contains Element  //h1[contains(text(),"Customer Locator")]  timeout=10s


*** Keyword ***
Extract and write in Excel
    Open Workbook  Locator-Output.xlsx
    ${default_sheet}=  Get Active Worksheet  
    
    ${no_of_countries}=  Get Element Count  //table[@style='']//th
    ${total_rows}=  Get Element Count  (//table[@style='']//tr)
    
    FOR  ${country}  IN RANGE  2  ${no_of_countries}+1
    
        ${country_name}=  Get Text  (//table[@style='']//th)[${country}]
        Create Worksheet  ${country_name}
        Set Active Worksheet  ${country_name}
        Set Cell Value  1  A  CustomerName
        Set Cell Value  1  B  Number of Locations
        
        FOR  ${row}  IN RANGE  2  ${total_rows}+1
            ${no_of_location}=  Get Text  (//table[@style='']//tr)[${row}]//td[${country}]
            IF  '${no_of_location}' != '0'
                ${customer_name}=  Get Text  (//table[@style='']//tr)[${row}]//td[1]
                
                LOG  ${no_of_location}
                LOG  ${customer_name}
                ${next}=    Find empty row
                Set Cell Value  ${next}  A  ${customer_name}
                Set Cell Value  ${next}  B  ${no_of_location}
            END
        END
        
        Save Workbook
    END
    Remove Worksheet  ${default_sheet}
    # Set Active Worksheet  
    Save Workbook
    
    Close Workbook
    

*** Tasks ***
Customer Locator
    Navigate to Locator Challenge
    Download Excel File
    Extract and write in Excel
    
    
