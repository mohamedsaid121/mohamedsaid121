function main(workbook: ExcelScript.Workbook, sheetName: string): string {
    const sourceFileName = workbook.getName().replace(/\.xlsx$/i, "");
    const SourceColumns = [
        "Staff ID",
        "Full Name",
        "Number of Sick Days",
        "Diagnosis",
        "SHW Feedback "
    ];

    let sheet = workbook.getWorksheet(sheetName);
    if (!sheet){
        console.log(`sheet ${sheetName} not found`);
        return;
    }

    let usedRange = sheet.getUsedRange();
    if (!usedRange) {
        console.log(`No data found in ${sheetName}`);
        return;
    }

    let table = sheet.getTables()[0];
    if(!table){
        console.log(`data is Not in table format`);
        table = sheet.addTable(usedRange.getAddress(), true);
    }

    let rawData = usedRange.getValues() as string[][];
    let headers = rawData[0].map(header => header.trim().toLowerCase());
    let trimmedSourceColumnNames = SourceColumns.map(colName => colName.trim().toLowerCase());

    // columns indices in excel file
    const columnIndices = trimmedSourceColumnNames.map(colName =>
        headers.findIndex(header => header === colName)
    );
    

    let filteredData = rawData.slice(1)
                              .filter(row => row.some(cell =>
        cell !== "" && cell !== null && cell !== undefined
    ));

    let extractedData = filteredData.map(row =>
        columnIndices.map(index => index !== -1 ? row[index]: "")
    );

    const result = {
        sourceFileName: sourceFileName,
        sourceColumns: SourceColumns,
        data: extractedData
    };

    return JSON.stringify(result);
}
