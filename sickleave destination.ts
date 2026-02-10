function main(workbook: ExcelScript.Workbook, sourceDataJson: string): void {
    // CONFIGURATION
    let columnMapping: { [key: string]: string } = {
        "Staff ID": "Staff ID",
        "Emp Name": "Full Name",
        "Diagnose": "SHW Feedback ",
        "Number of Sick Days": "Number of Sick Days",
        "Original Diagnose": "Diagnosis"
    };
    interface SourceData {
        sourceFileName: string;
        sourceColumns: string[];
        data: string[][];
    }

    const sourceColumnName = "Source";

    // Parsing json source data
    const parsed: SourceData = JSON.parse(sourceDataJson) as SourceData;
    const sourceFileName: string = parsed.sourceFileName;
    const sourceColumns: string[] = parsed.sourceColumns;
    const rowsToAppend: string[][] = parsed.data;

    if (!rowsToAppend || rowsToAppend.length === 0) {
        console.log("No data to append");
        return;
    }

    let sheet = workbook.getActiveWorksheet();
    let table = sheet.getTable("Table1");
    

    
    let UsedRange = table.getRange();
    let tableData = UsedRange.getValues() as string[][];
    let headers = tableData[0];
    

    // Prepare data for appending
    let newRows: string[][] = [];

    for (let i = 0; i < rowsToAppend.length; i++) {
        let sourceRow = rowsToAppend[i];
        let newRow: string[] = new Array(headers.length).fill("");

        // Filling Destination rows with source rows
        for (const destCol in columnMapping) {
            const sourceColName: string = columnMapping[destCol];
            const sourceIndex = sourceColumns.indexOf(sourceColName);
            const destIndex = headers.indexOf(destCol);

            if (sourceIndex !== -1 && destIndex !== -1) {
                newRow[destIndex] = sourceRow[sourceIndex] || "";
            }
        }

        // Source file name
        let sourceIndex = headers.indexOf(sourceColumnName);
        if (sourceIndex !== -1) {
            newRow[sourceIndex] = sourceFileName;
        }

        newRows.push(newRow);
    }

    
    if (newRows.length > 0) {
        table.addRows(-1, newRows);
    }

    
    sheet.getUsedRange()?.getFormat().autofitColumns();
}
