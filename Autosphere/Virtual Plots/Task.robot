*** Settings ***
Library    Autosphere.Browser.Playwright
Library    Autosphere.Excel.Files
Library    OperatingSystem
Library    Autosphere.HTTP
Library    String

*** Variables ***
${URL}  https://botsdna.com/vitrualplots/
${FILE_PATH}  ${CURDIR}\\File
${DOWNLOAD_URL}  https://botsdna.com/vitrualplots/input.xlsx
${FILE_NAME}  input.xlsx

*** Keywords ***
Open Website And Download Required File
    Open Browser  url=${URL}
    Maximize Browser Window
    Wait Until Element Is Visible    (//tbody)[1]

    ${file_status}=  Run Keyword And Return Status  File Should Exist  ${FILE_PATH}\\${FILE_NAME}
    IF    not ${file_status}
        Create Directory    ${FILE_PATH}
        Download  url=${DOWNLOAD_URL}  target_file=${FILE_PATH}
        Wait Until Keyword Succeeds    2    2s    File Should Exist  ${FILE_PATH}\\${FILE_NAME}
    END

*** Keywords ***
Booking Land Virtual
    Open Workbook    ${FILE_PATH}\\${FILE_NAME}
    @{details}=  Read Worksheet  header=${True}
    LOG  ${details}
    FOR    ${detail}    IN    @{details}
        Log    ${detail}
        ${seller_number}=  Set Variable  ${detail}[Seller Mobile]
        ${valid_seller_number}=  Validated Number  ${seller_number}
        LOG  ${valid_seller_number}

        ${seller_email}=  Set Variable  ${detail}[Seller Email]
        ${buyer_number}=  Set Variable  ${detail}[Buyer Mobile]
        ${valid_buyer_number}=  Validated Number  ${buyer_number}
        LOG  ${valid_buyer_number}

        ${buyer_email}=  Set Variable  ${detail}[Buyer Email]
        ${plot_number}=  Set Variable  ${detail}[Plot No]
        ${area}=  Set Variable  ${detail}[Sqft]

        ${status}  ${transaction_number}  ${buyer_name}  ${seller_name}  Perform Booking Over Website  ${valid_seller_number}  ${valid_buyer_number}  ${plot_number}  ${area}
        IF    ${status}
            LOG  Plot Booked Successfully ${\n} Buyer Name: ${buyer_name} ${\n} Seller Name: ${seller_name} ${\n} Transaction Number: ${transaction_number}   INFO

        ELSE
            LOG  Problem Occurred While Booking Plot
        END

    END



    Close Workbook

*** Keywords ***
Perform Booking Over Website
    [Documentation]     This Keyword take buyer,seller numbers and plot number and area
    ...                 And return buyer and seller name with transaction number
    ...                 And overall status as True/False to show either booking is success or fail
    [Arguments]  ${seller_num}  ${buyer_num}  ${plot_num}  ${covered_area}
    LOG  ${seller_num}
    LOG  ${buyer_num}
    LOG  ${plot_num}
    LOG  ${covered_area}
    
    ${booking_status}=  Set Variable  ${False}
    ${tran_number}=  Set Variable  ${None}
    ${buyer_name}=  Set Variable  ${EMPTY}
    ${seller_name}=  Set Variable  ${EMPTY}

    Wait Until Page Contains Element    (//h1)[1]

    Scroll Element Into View    //td[contains(text(),"${seller_num}")]
    ${seller_name}=  Get Text  //td[contains(text(),"${seller_num}")]//preceding-sibling::td[1]
    Click Element    //td[contains(text(),"${seller_num}")]//preceding-sibling::td[2]//input
    LOG  "${seller_name}" selected as seller   INFO

    Scroll Element Into View    //td[contains(text(),"${buyer_num}")]
    ${buyer_name}=  Get Text  //td[contains(text(),"${buyer_num}")]//preceding-sibling::td[1]
    Click Element    //td[contains(text(),"${buyer_num}")]//preceding-sibling::td[3]//input
    LOG  "${buyer_name}" selected as buyer   INFO

    LOG  Inputting Plot Number Below
    Scroll Element Into View    //input[@value="Submit"]
    Input Text    //td[contains(text(),"Plot No")]//following-sibling::td//input    ${plot_num}

    LOG  Inputting Area Below
    Input Text  //td[contains(text(),"Sqft")]//following-sibling::td//input  ${covered_area}
    
    Click Element    //input[@value="Submit"]

    Wait Until Element Is Visible    //*[@id="TransNo"]  timeout=15s
    ${tran_number}=  Get Text    //*[@id="TransNo"]
    LOG  ${tran_number}
    Go Back
    Reload Page
    
    IF    ("${seller_name}" != "${EMPTY}") and ("${buyer_name}" != "${EMPTY}") and ("${tran_number}" != "${None}")
         ${booking_status}=  Set Variable  ${True}
    END

    [Return]  ${booking_status}  ${tran_number}  ${buyer_name}  ${seller_name}
    
    
*** Keywords ***
Validated Number
    [Documentation]     This Keyword take number, convert to string if it is not
    ...                 And return either last 10 digit or False (if number is less than 10)
    [Arguments]  ${number}
    ${number}=  Convert To String    ${number}
    ${number}=  Remove String    ${number}  +  -  ${SPACE}
#    ${type}  Evaluate    type(${number})
#    IF    "str" not in "${type}"
#         ${number}=  Convert To String    ${number}
#    END

    ${result}=  Set Variable  ${False}
    ${len}=  Get Length    ${number}
    IF    ${len} >= 10
        ${number}=  Strip String  ${number}
        ${result}=  Get Substring    ${number}    -10
    END

    [Return]  ${result}

*** Tasks ***
Virtual Plots
    Open Website And Download Required File
    Booking Land Virtual