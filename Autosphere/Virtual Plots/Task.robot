*** Settings ***
Library    Autosphere.Browser.Playwright
Library    Autosphere.Excel.Files
Library    OperatingSystem
Library    Autosphere.HTTP
Library    String
Library    DateTime
Library    Autosphere.Email.ImapSmtp

*** Variables ***
${URL}  https://botsdna.com/vitrualplots/
${FILE_PATH}  ${CURDIR}\\File
${DOWNLOAD_URL}  https://botsdna.com/vitrualplots/input.xlsx
${FILE_NAME}  input.xlsx
&{SMTP_DETAIL}  server=smtp.gmail.com  smtp_port=587  username=danysheikh@gmail.com  password=[APP_PASSWORD]

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
    ${row}=  Set Variable  2
    FOR    ${detail}    IN    @{details}
        LOG    ${row}
        Log    ${detail}
        ${excel_status}=  Set Variable  Not Success
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
            LOG  ${SMTP_DETAIL}
            IF  ("${seller_email}" != "None") and ("${buyer_email}" != "None")
                ${email_status}  Run Keyword And Return Status   Send Email  ${SMTP_DETAIL}  ${detail}  ${buyer_name}  ${seller_name}  ${transaction_number}
                IF    ${email_status}
                     ${excel_status}=  Set Variable    ${transaction_number}
                ELSE
                     ${excel_status}=  Set Variable    Problem Sending Email
                END
            ELSE
                ${excel_status}=  Set Variable    Email Not Valid
            END

        ELSE
            LOG  Problem Occurred While Booking Plot
            ${excel_status}=  Set Variable    Problem Occurred While Booking Plot
        END

        Set Cell Value    ${row}    G    ${excel_status}
        ${row}=  Evaluate    ${row} + 1
        Save Workbook

    END

    Close Workbook

*** Keywords ***
Send Email
    [Documentation]   This Keyword send an email as per required format
    [Arguments]     ${SMTP_DETAIL}  ${detail}  ${buyer_name}  ${seller_name}  ${transaction_number}
    LOG  ${buyer_name}
    LOG  ${seller_name}
    LOG  ${transaction_number}
    LOG  ${detail}
    LOG  ${SMTP_DETAIL}
    ${date}=  Get Current Date
    ${subject}=  Set Variable  Plot has booked Successfully - ${transaction_number}
    ${receiver}=  Set Variable  ${detail}[Buyer Email],${detail}[Seller Email]
    ${body}=  Catenate
    ...      <html>
    ...      <head>
    ...        <style>
    ...          body { font-family: Arial, sans-serif; font-size: 14px; color: #333; }
    ...        </style>
    ...      </head>
    ...      <body>
    ...        <p>Dear <b>${seller_name}</b> & <b>${buyer_name}</b>,</p>
    ...        <p>New Plot (Plot Number: <b>${detail}[Plot No]</b> ) with No.of.Sqft <b>${detail}[Sqft]</b> has been Booked successfully on ${date}</p>
    ...        <p>Here you can find Booking Details...</p>
    ...        <table border='1' style='border-collapse:collapse'>
    ...            <tr>
    ...                <td>Booking Number</td>
    ...                <td>${transaction_number}</td>
    ...            </tr>
    ...            <tr>
    ...                <td>Buyer Name</td>
    ...                <td>${buyer_name}</td>
    ...            </tr>
    ...            <tr>
    ...                <td>Buyer Phone Number</td>
    ...                <td>${detail}[Buyer Mobile]</td>
    ...            </tr>
    ...            <tr>
    ...                <td>Seller Name</td>
    ...                <td>${seller_name}</td>
    ...            </tr>
    ...            <tr>
    ...                <td>Seller Phone Number</td>
    ...                <td>${detail}[Seller Mobile]</td>
    ...            </tr>
    ...        </table><br>
    ...        <p>Thanks |</p>
    ...      </body>
    ...    </html>
    LOG  ${body}
#    Authorize  account=${SMTP_DETAIL}[username]  password=${SMTP_DETAIL}[password]  smtp_server=${SMTP_DETAIL}[server]  smtp_port=${SMTP_DETAIL}[smtp_port]
#    Send Message  sender=${SMTP_DETAIL}[username]  recipients=${receiver}  subject=${subject}  body=${body}  html=True
    LOG  Authorize and Send Message keyword is commented in Code



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