function main(workbook: ExcelScript.Workbook, sourceDataJson: string): void {
    // CONFIGURATION
    const columnMapping: { [key: string]: string } = {
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

    // Parse with type assertion
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
    

    // Get table data
    const UsedRange = table.getRange();
    const tableData = UsedRange.getValues() as string[][];
    const headers = tableData[0];

    // Prepare data for appending
    const newRows: string[][] = [];

    for (let i = 0; i < rowsToAppend.length; i++) {
        const sourceRow = rowsToAppend[i];
        const newRow: string[] = new Array(headers.length).fill("");

        // columns Mapping
        for (const destCol in columnMapping) {
            const sourceColName: string = columnMapping[destCol];
            const sourceIndex = sourceColumns.indexOf(sourceColName);
            const destIndex = headers.indexOf(destCol);

            if (sourceIndex !== -1 && destIndex !== -1) {
                newRow[destIndex] = sourceRow[sourceIndex] || "";
            }
        }

        // Add file name of each source
        const sourceIndex = headers.indexOf(sourceColumnName);
        if (sourceIndex !== -1) {
            newRow[sourceIndex] = sourceFileName;
        }

        newRows.push(newRow);
    }

    
    if (newRows.length > 0) {
        table.addRows(-1, newRows);
    }

    // Auto-fit
    sheet.getUsedRange()?.getFormat().autofitColumns();
}
