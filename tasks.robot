*** Settings ***
Documentation       Orders robots from RobotSpareBin Industries Inc.
...                 Saves the order HTML receipt as a PDF file.
...                 Saves the screenshot of the ordered robot.
...                 Embeds the screenshot of the robot to the PDF receipt.
...                 Creates ZIP archive of the receipts and the images.

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.PDF
Library             RPA.Robocloud.Secrets
Library             RPA.Tables
Library             RPA.Archive
Library             RPA.RobotLogListener


*** Tasks ***
Insert The Sales Data For The Week And Export It As A Pdf
    [Documentation]    Inserts the sales data for the week into the intranet and exports it as a PDF file.
    Open the robot order website
    Download The Excel File
    Fill The Form Using The Data From The Excel File
    Archive Folder With Zip    %{ROBOT_ROOT}${/}output    Receipts.zip    ${TRUE}    exclude=output.xml
    [Teardown]    Logout And Close The Browser


*** Keywords ***
Open the robot order website
    Open Available Browser    https://robotsparebinindustries.com/#/robot-order

Download The Excel File
    Download    https://robotsparebinindustries.com/orders.csv    overwrite=TRUE

Close The Annoying Modal
    Click Element When Visible    css:button[class="btn btn-dark"]

Fill the form
    [Arguments]    ${order}
    Close The Annoying Modal
    Select From List By Value    head    ${order}[Head]
    Select Radio Button    body    ${order}[Body]
    Input Text    xpath:/html/body/div/div/div[1]/div/div[1]/form/div[3]/input    ${order}[Legs]
    Input Text    address    ${order}[Address]

Fill The Form Using The Data From The Excel File
    Mute Run On Failure	Run Keyword
    ${sales_orders}=    Read table from CSV    %{ROBOT_ROOT}${/}orders.csv
    FOR    ${order}    IN    @{sales_orders}
        Fill the form    ${order}
        Click Element When Visible    id:preview
        Order the robot
        ${robot_image}=    Get Element Attribute    id:receipt    outerHTML
        Html To Pdf    ${robot_image}    ${OUTPUT_DIR}${/}Receipts${/}receipt${order}[Order number].pdf
        Screenshot    id:robot-preview-image    ${OUTPUT_DIR}${/}Receipts${/}image${order}[Order number].png
        ${receiptPDF}=    Open Pdf    ${OUTPUT_DIR}${/}Receipts${/}receipt${order}[Order number].pdf
        ${robotPNG}=    Create List
        ...    ${OUTPUT_DIR}${/}Receipts${/}image${order}[Order number].png
        ...    ${OUTPUT_DIR}${/}Receipts${/}receipt${order}[Order number].pdf
        Add Files To Pdf    ${robotPNG}    ${OUTPUT_DIR}${/}Receipts${/}receipt.pdf    ${TRUE}
        Close Pdf    ${receiptPDF}
        Click Element When Visible    id:order-another
        Log    ${order}
        Archive Folder With Zip    ${OUTPUT_DIR}${/}Receipts    Receipts.zip    ${TRUE}
    END

END
    Close Workbook

Order the robot
    Click Button    id:order
    FOR    ${i}    IN RANGE    9999999
        ${success}=    Is Element Visible    id:receipt
        IF    ${success}    BREAK
        Click Button    id:order
    END

Logout And Close The Browser
    Close Browser
