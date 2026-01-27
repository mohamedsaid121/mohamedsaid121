function main(
    workbook: ExcelScript.Workbook, 
    sheetName: string, 
    in_sourceColumns: string[], 
    in_data: string[][],
    sourceFileName: string
) {
    console.log("Script 2 starting...");
    
    let sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
        sheet = workbook.addWorksheet(sheetName);
    }
    
    // Your column mapping
    const columnMapping: {[key: string]: string} = {
        'Staff ID': 'Staff ID',
        'Emp Name': 'Full Name',
        'Number of Sick Days': 'Number of Sick Days',
        'Original Diagnose': 'Diagnosis',
        'Diagnose': 'SHW Feedback '
    };
    
    const sourceFileColumn = "Source File";
    
    // Get or create table
    let table = sheet.getTables()[0];
    if (!table) {
        const headers = [...Object.keys(columnMapping), sourceFileColumn];
        sheet.getRange("A1").getResizedRange(0, headers.length - 1).setValues([headers]);
        table = sheet.addTable(sheet.getRange("A1").getResizedRange(0, headers.length - 1), true);
        table.setName("ConsolidatedData");
    }
    
    const tableHeaders = table.getHeaderRowRange().getValues()[0] as string[];
    
    // Find maximum rows in source data
    const maxRows = Math.max(...in_data.map(col => col.length), 0);
    
    if (maxRows === 0) {
        console.log("No data to append");
        return { message: "No data" };
    }
    
    // Prepare rows
    const rowsToAdd: string[][] = [];
    const sourceFileIndex = tableHeaders.indexOf(sourceFileColumn);
    
    for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
        const newRow: string[] = new Array(tableHeaders.length).fill("");
        
        // Map each source column to destination
        in_sourceColumns.forEach((sourceCol, colIndex) => {
            const trimmedSourceCol = sourceCol.trim();
            
            // Find destination column for this source column
            for (const [destCol, mappedSourceCol] of Object.entries(columnMapping)) {
                if (mappedSourceCol.trim() === trimmedSourceCol) {
                    const destIndex = tableHeaders.indexOf(destCol);
                    if (destIndex !== -1 && rowIndex < in_data[colIndex].length) {
                        newRow[destIndex] = String(in_data[colIndex][rowIndex]);
                    }
                    break;
                }
            }
        });
        
        // Add source file name
        if (sourceFileIndex !== -1) {
            newRow[sourceFileIndex] = sourceFileName;
        }
        
        rowsToAdd.push(newRow);
    }
    
    // Add rows to table
    if (rowsToAdd.length > 0) {
        table.addRows(-1, rowsToAdd);
        console.log(`Added ${rowsToAdd.length} rows from ${sourceFileName}`);
    }
    
    return {
        message: `Success: ${rowsToAdd.length} rows appended`
    };
}
