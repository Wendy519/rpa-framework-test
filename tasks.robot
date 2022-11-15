*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF

Library             RPA.Browser.Selenium
Library             Dialogs
Library             RPA.HTTP
Library             RPA.Tables
Library             RPA.PDF
Library             RPA.Archive
Library             RPA.FileSystem


*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open Intranet Website
    Log in
    Download the CSV file
    Fill in sales data from CSV
    Collect the results
    Collect sales data in PDF
    Log out
    Save created files in an Archive
    Keep browser open


*** Keywords ***
Open Intranet Website
    Open Available Browser    https://robotsparebinindustries.com/

Log in
    Input Text    username    maria
    Input Password    password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Download the CSV file
    Download    https://robotsparebinindustries.com/SalesData.xlsx

Enter data for each person
    [Arguments]    ${salesrep}
    Input Text    firstname    ${salesrep}[First Name]
    Input Text    lastname    ${salesrep}[Last Name]
    Input Text    salesresult    ${salesrep}[Sales]
    Select From List By Value    salestarget    ${salesrep}[Sales Target]
    Click Button    Submit

Fill in sales data from CSV
    ${salesreps}=    Read table from CSV    SalesData.csv
    FOR    ${salesrep}    IN    @{salesreps}
        Enter data for each person    ${salesrep}
    END

Collect the results
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}files${/}sales_summary.png

Collect sales data in PDF
    Wait Until Element Is Visible    id:sales-results
    ${sales_results_html}=    Get Element Attribute    id:sales-results    outerHTML
    Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}files${/}sales_results.pdf
    Print To Pdf    ${OUTPUT_DIR}${/}files${/}other.pdf
    ${image}=    Set Variable    ${OUTPUT_DIR}${/}files${/}sales_summary.png
    ${TEMPLATE}=    Set Variable    devdata/test.template
    ${PDF}=    Set Variable    ${OUTPUT_DIR}${/}combined.pdf
    ${DATA}=    Create Dictionary
    ...    image=${image}
    ...    sales_results_html=${sales_results_html}
    Template HTML to PDF
    ...    template=${TEMPLATE}
    ...    output_path=${PDF}
    ...    variables=${DATA}

Log out
    Click Button    Log out

Save created files in an Archive
    Archive Folder With Zip    ${OUTPUT_DIR}${/}files    ${OUTPUT_DIR}${/}MyFiles.zip
    Remove Directory    ${OUTPUT_DIR}${/}files    recursive=${True}

Keep browser open
    Pause Execution
