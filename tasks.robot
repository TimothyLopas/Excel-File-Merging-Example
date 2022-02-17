*** Settings ***
Documentation     Template robot main suite.
Library           RPA.Excel.Files
Library           Collections
Library           OperatingSystem

*** Variables ***
${FILE_NAME}=     Test_Workbook.xlsx
${MANY_EXCEL_PATH}=    ${CURDIR}${/}ExcelFolder

*** Keywords ***
Multiple Sheets into One Excel Sheet
    ${WORKSHEET_DICT}=    Create Dictionary
    ${i}=    Set Variable    ${0}
    Open Workbook    ${CURDIR}${/}${FILE_NAME}
    @{sheets}=    List Worksheets
    FOR    ${sheet}    IN    @{sheets}
        ${worksheet_data}=    Read Worksheet    ${sheet}    header=${FALSE}
        Set To Dictionary    ${WORKSHEET_DICT}    name_${i}=${sheet}    data_${i}=${worksheet_data}
        ${i}=    Evaluate    ${i}+1
    END
    Close Workbook
    Create Workbook    fmt=xlsx
    FOR    ${y}    IN RANGE    ${i}
        Append Rows To Worksheet    ${WORKSHEET_DICT.data_${y}}
    END
    Save Workbook    ${OUTPUT_DIR}${/}MultipleSheetsToOne.xlsx

Multiple Files into one Excel Sheet
    ${WORKSHEET_DICT}=    Create Dictionary
    @{list}=    List Files In Directory    ${MANY_EXCEL_PATH}    *.xlsx
    ${i}=    Set Variable    ${0}
    FOR    ${workbook}    IN    @{list}
        Open Workbook    ${MANY_EXCEL_PATH}${/}${workbook}
        @{sheets}=    List Worksheets
        ${worksheet_data}=    Read Worksheet    ${sheets[0]}    header=${FALSE}
        Set To Dictionary    ${WORKSHEET_DICT}    name_${i}=${sheets[0]}    data_${i}=${worksheet_data}
        ${i}=    Evaluate    ${i}+1
    END
    Close Workbook
    Create Workbook    fmt=xlsx
    FOR    ${y}    IN RANGE    ${i}
        Append Rows To Worksheet    ${WORKSHEET_DICT.data_${y}}
    END
    Save Workbook    ${OUTPUT_DIR}${/}MultipleWorkbooksToOne.xlsx

*** Tasks ***
Excel Combination Tasks
    Multiple Sheets into One Excel Sheet
    Multiple Files into one Excel Sheet
