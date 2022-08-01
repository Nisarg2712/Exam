*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.Excel.Files
Library             RPA.HTTP
Library             RPA.PDF
Library             RPA.Tables
Library             Collections
Library             RPA.FileSystem
Library             RPA.RobotLogListener
Library             RPA.Archive
Library             RPA.Dialogs


*** Variables ***
${SFile}    C:/Users/npatel.WMP/Documents/Robots/exambot/Screenshot
${HFile}    C:/Users/npatel.WMP/Documents/Robots/exambot/HPDF
${RFile}    C:/Users/npatel.WMP/Documents/Robots/exambot/RFile


*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open The Intranet Website
    Input Dialog
    Download The CSV File
    Fill The Form Using The Data From The Excel File
    Make Zip
    Close All Browsers
    [Teardown]    No Operation


*** Keywords ***
Open The Intranet Website
    Open Available Browser    url=https://robotsparebinindustries.com/#/robot-order    maximized=True

Input Dialog
    Add heading    Please Enter URL For CSV
    Add text input    URL    label=CSV Download URL
    ${result}=    Run dialog

Download The CSV File
   Download    https://robotsparebinindustries.com/orders.csv    overwrite=True

Fill The Form Using The Data From The Excel File
    ${index}=    Set Variable    0
    ${order_data}=    Read table from CSV    orders.csv    header=True
    FOR    ${order_data}    IN    @{order_data}
        ${index}=    Evaluate    ${index} + 1
        Fill And Submit The Form For One Person    ${order_data}    ${index}
    END

Fill And Submit The Form For One Person
    [Arguments]    ${order_data}    ${index}
    Click Element    alias:Button
    Select From List By Value    head    ${order_data}[Head]
    Input Text    address    ${order_data}[Address]
    Select Radio Button    body    ${order_data}[Body]
    Input Text    alias:Input    ${order_data}[Legs]    clear=False
    Click Button    Preview
    Wait Until Keyword Succeeds    5x    1    Get Order    ${index}

Get Order
    [Arguments]    ${index}
    Click Button    Order
    ${RHTML}=    Get Element Attribute    id:receipt    outerHTML
    Html To Pdf    ${RHTML}    ${CURDIR}${/}HPDF${/}${index}.pdf
    Screenshot    locator=id:robot-preview-image    filename=${CURDIR}${/}Screenshot${/}${index}.png
    ${RPNG}=    Create List    ${CURDIR}${/}HPDF${/}${index}.pdf    ${CURDIR}${/}Screenshot${/}${index}.png
    ${RPDF}=    Open Pdf    ${CURDIR}${/}HPDF${/}${index}.pdf
    Add Files To Pdf    files=${RPNG}    target_document=${CURDIR}${/}RFile${/}${index}.pdf
    Close Pdf    ${RPDF}
    Click Button    Order another robot

Make Zip
    Archive Folder With Zip    ${RFile}${/}    ${OUTPUT_DIR}${/}Reciept.zip
