*** Settings ***
Library    Autosphere.Browser.Playwright
Library    OperatingSystem
Library    Autosphere.HTTP


*** Variables ***
${WEBSITE}      https://botsdna.com/BGV/
${DOWNLOAD_LINK}        https://botsdna.com/BGV/Employee%20Documents.zip
${DOCUMENT_PATH}        ${CURDIR}\\Documents
${DOCUMENT_NAME}        Employee Documents.zip

*** Keywords ***
Open Website And Download Documents
    Open Browser    url=${WEBSITE}
    Maximize Browser Window
    Create Directory    ${DOCUMENT_PATH}

    # If Document file does not exist then it will download it.
    ${file_status}=  Run Keyword And Return Status  File Should Exist    ${DOCUMENT_PATH}\\${DOCUMENT_NAME}
    IF    '${file_status}' != 'True'
        Empty Directory    ${DOCUMENT_PATH}
        Download    url=${DOWNLOAD_LINK}        target_file=${DOCUMENT_PATH}
        Sleep    3s
    END

    # Check file name if it is with %20 then replace it with actual file name to remove that %20
    ${check_file}=  Run Keyword And Return Status  File Should Exist    ${DOCUMENT_PATH}\\Employee%20Documents.zip
    IF    ${check_file}
         Move File    ${DOCUMENT_PATH}\\Employee%20Documents.zip    ${DOCUMENT_PATH}\\${DOCUMENT_NAME}
    END


*** Tasks ***
Background Verification
    Open Website And Download Documents