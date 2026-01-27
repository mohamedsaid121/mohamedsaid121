function main(
    workbook: ExcelScript.Workbook, 
    sheetName: string, 
    in_sourceColumns: string[], 
    in_data: string[][],
    sourceFileName: string  // NEW: Add source file name parameter
) {
    console.log("=== Script 2 Started ===");
    console.log(`Source file: ${sourceFileName}`);
    console.log(`Received ${in_sourceColumns.length} columns`);
    console.log("Source columns:", JSON.stringify(in_sourceColumns));
    console.log(`Data structure: ${in_data.length} columns, max rows: ${Math.max(...in_data.map(col => col.length))}`);
    
    // Debug: Show data sample
    in_data.forEach((colData, idx) => {
        console.log(`Column ${in_sourceColumns[idx]}: ${colData.length} values`);
        if (colData.length > 0) {
            console.log(`  Sample: ${colData.slice(0, 3).join(", ")}...`);
        }
    });
    
    let sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
        console.log(`Sheet "${sheetName}" not found - creating it`);
        sheet = workbook.addWorksheet(sheetName);
    }
    
    // Define mapping (destination: source)
    const columnMapping: {[key: string]: string} = {
        'Staff ID': 'Staff ID',
        'Emp Name': 'Full Name',
        'Number of Sick Days': 'Number of Sick Days',
        'Original Diagnose': 'Diagnosis',
        'Diagnose': 'SHW Feedback '
    };
    
    // Source file column name
    const sourceFileColumn = "Source File";
    
    // Get or create table
    let table = sheet.getTables()[0];
    let destinationColumns = Object.keys(columnMapping);
    
    if (!table) {
        console.log("No table found - creating new table");
        // Create headers including source file column
        const headers = [...destinationColumns, sourceFileColumn];
        const headerRange = sheet.getRange("A1").getResizedRange(0, headers.length - 1);
        headerRange.setValues([headers]);
        
        table = sheet.addTable(headerRange, true);
        table.setName("ConsolidatedData");
        console.log(`Created new table with headers: ${headers.join(", ")}`);
    } else {
        console.log("Found existing table");
    }
    
    // Get table headers
    const tableHeaders = table.getHeaderRowRange().getValues()[0] as string[];
    console.log("Table headers:", tableHeaders);
    
    // Prepare mapping info
    interface ColumnInfo {
        sourceCol: string;
        destCol: string;
        sourceIndex: number;
        destIndex: number;
        sourceData: string[];
    }
    
    const columnInfo: ColumnInfo[] = [];
    
    // Build column information
    in_sourceColumns.forEach((sourceCol, sourceIndex) => {
        const trimmedSourceCol = sourceCol.trim();
        
        // Find matching destination column
        for (const [destCol, mappedSourceCol] of Object.entries(columnMapping)) {
            if (mappedSourceCol.trim() === trimmedSourceCol) {
                const destIndex = tableHeaders.indexOf(destCol);
                if (destIndex !== -1) {
                    columnInfo.push({
                        sourceCol: trimmedSourceCol,
                        destCol: destCol,
                        sourceIndex: sourceIndex,
                        destIndex: destIndex,
                        sourceData: in_data[sourceIndex] || []
                    });
                    console.log(`Mapped: ${sourceCol} → ${destCol} (table col ${destIndex})`);
                } else {
                    console.log(`Warning: Destination column "${destCol}" not found in table`);
                }
                break;
            }
        }
    });
    
    // Get source file column index
    const sourceFileIndex = tableHeaders.indexOf(sourceFileColumn);
    if (sourceFileIndex === -1) {
        console.log(`Warning: Source file column "${sourceFileColumn}" not found in table`);
    } else {
        console.log(`Source file column at index: ${sourceFileIndex}`);
    }
    
    // Determine number of rows to add
    const maxRows = Math.max(...columnInfo.map(info => info.sourceData.length));
    console.log(`Will add ${maxRows} rows`);
    
    if (maxRows === 0) {
        console.log("No data to append");
        return { message: "No data to append" };
    }
    
    // Prepare rows to add
    const rowsToAdd: string[][] = [];
    
    for (let rowIndex = 0; rowIndex < maxRows; rowIndex++) {
        const newRow: string[] = new Array(tableHeaders.length).fill("");
        
        // Fill data for each mapped column
        columnInfo.forEach(info => {
            if (rowIndex < info.sourceData.length && info.destIndex !== -1) {
                newRow[info.destIndex] = String(info.sourceData[rowIndex]);
            }
        });
        
        // Add source file name
        if (sourceFileIndex !== -1) {
            newRow[sourceFileIndex] = sourceFileName;
        }
        
        rowsToAdd.push(newRow);
    }
    
    console.log(`Prepared ${rowsToAdd.length} rows to add`);
    
    // Add rows to table
    if (rowsToAdd.length > 0) {
        try {
            table.addRows(-1, rowsToAdd);
            console.log(`✅ Successfully appended ${rowsToAdd.length} rows from ${sourceFileName}`);
            
            // Auto-fit columns
            sheet.getUsedRange()?.getFormat().autofitColumns();
        } catch (error) {
            console.log(`❌ Error adding rows: ${error}`);
            // Fallback: Try direct range writing
            try {
                const lastRow = table.getRange().getLastRow().getRowIndex();
                const startRow = lastRow + 1;
                const startCol = table.getRange().getColumnIndex();
                
                const targetRange = sheet.getRange(startRow, startCol)
                    .getResizedRange(rowsToAdd.length - 1, tableHeaders.length - 1);
                targetRange.setValues(rowsToAdd);
                
                console.log(`✅ Added ${rowsToAdd.length} rows via range (fallback)`);
            } catch (fallbackError) {
                console.log(`❌ Fallback also failed: ${fallbackError}`);
            }
        }
    }
    
    console.log("=== Script 2 Completed ===");
    
    return {
        message: `Appended ${rowsToAdd.length} rows from ${sourceFileName}`,
        rowsAppended: rowsToAdd.length
    };
}