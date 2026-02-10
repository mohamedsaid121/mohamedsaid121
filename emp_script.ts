function main(workbook: ExcelScript.Workbook, sourceDataJson: string): void {
    // CONFIGURATION
    let columnMapping: { [key: string]: string } = {
        "Staff ID": "Staff ID",
        "Emp Name": "Full Name",
        "Diagnose": "SHW Feedback ",
        "Number of Sick Days": "Number of Sick Days",
        "Original Diagnose": "Diagnosis",
        "Department": "Department"
    };
    
    interface SourceData {
        sourceFileName: string;
        sourceColumns: string[];
        data: string[][];
    }

    const sourceColumnName = "Source";
    
    // === LOAD EMPLOYEES FROM SAME WORKBOOK ===
    console.log("Loading employee departments from EmployeesData sheet...");
    const departmentMap: {[staffId: string]: string} = {};
    
    try {
        const employeesSheet = workbook.getWorksheet("EmployeesData");
        
        if (employeesSheet) {
            const employeesRange = employeesSheet.getUsedRange();
            
            if (employeesRange) {
                const allEmployees = employeesRange.getValues() as string[][];
                console.log(`Found ${allEmployees.length} rows in EmployeesData sheet`);
                
                const empHeaders = allEmployees[0].map(h => 
                    h === null || h === undefined ? "" : String(h).trim()
                );
                
                // Find column indices - using more flexible matching
                const staffIdIndex = empHeaders.findIndex(h => 
                    h.toLowerCase().includes("staff") || 
                    h.toLowerCase().includes("id") ||
                    h.toLowerCase().includes("sys")
                );
                
                const deptIndex = empHeaders.findIndex(h => 
                    h.toLowerCase().includes("department") || 
                    h.toLowerCase().includes("dept")
                );
                
                console.log(`Staff ID column index: ${staffIdIndex}, column name: "${empHeaders[staffIdIndex]}"`);
                console.log(`Department column index: ${deptIndex}, column name: "${empHeaders[deptIndex]}"`);
                
                if (staffIdIndex !== -1) {
                    let loadedCount = 0;
                    for (let i = 1; i < allEmployees.length; i++) {
                        const row = allEmployees[i];
                        const staffId = (row[staffIdIndex] || "").toString().trim();
                        if (staffId) {
                            const department = deptIndex !== -1 ? 
                                (row[deptIndex] || "").toString().trim() : "Unknown";
                            departmentMap[staffId] = department;
                            loadedCount++;
                            console.log(`Mapped Staff ID: ${staffId} -> Department: ${department}`);
                        }
                    }
                    console.log(`Loaded ${loadedCount} employee departments into memory`);
                    console.log("Sample mappings:", Object.keys(departmentMap).slice(0, 5).map(k => `${k}:${departmentMap[k]}`));
                } else {
                    console.log("ERROR: Could not find Staff ID column in EmployeesData sheet");
                    console.log("Available columns:", empHeaders);
                }
            } else {
                console.log("ERROR: EmployeesData sheet is empty");
            }
        } else {
            console.log("WARNING: EmployeesData sheet not found in workbook");
            console.log("Available sheets:", workbook.getWorksheets().map(s => s.getName()));
        }
    } catch (error) {
        console.log(`ERROR loading employee data: ${error}`);
        console.log("Continuing without department data");
    }

    // Parsing json source data
    const parsed: SourceData = JSON.parse(sourceDataJson) as SourceData;
    const sourceFileName: string = parsed.sourceFileName;
    const sourceColumns: string[] = parsed.sourceColumns;
    const rowsToAppend: string[][] = parsed.data;

    console.log("Source columns:", sourceColumns);
    console.log("Number of rows to append:", rowsToAppend.length);
    console.log("First row sample:", rowsToAppend[0]);

    if (!rowsToAppend || rowsToAppend.length === 0) {
        console.log("No data to append");
        return;
    }

    let sheet = workbook.getActiveWorksheet();
    let table = sheet.getTable("Table1");
    
    if (!table) {
        console.log("Table 'Table1' not found");
        return;
    }
    
    let headers = table.getHeaderRowRange().getValues()[0] as string[];
    console.log("Table headers:", headers);

    // Prepare data for appending
    let newRows: string[][] = [];

    for (let i = 0; i < rowsToAppend.length; i++) {
        let sourceRow = rowsToAppend[i];
        let newRow: string[] = new Array(headers.length).fill("");

        // Get Staff ID from source row using the source column name
        const sourceStaffIdIndex = sourceColumns.indexOf("Staff ID");
        const staffId = sourceStaffIdIndex !== -1 ? 
            (sourceRow[sourceStaffIdIndex] || "").toString().trim() : "";
        
        console.log(`Row ${i}: Staff ID from source = "${staffId}"`);

        // Filling Destination rows with source rows
        for (const destCol in columnMapping) {
            const sourceColName: string = columnMapping[destCol];
            const sourceIndex = sourceColumns.indexOf(sourceColName);
            const destIndex = headers.indexOf(destCol);

            if (destIndex !== -1) {
                if (destCol === "Department") {
                    // Lookup department using the Staff ID
                    const department = departmentMap[staffId] || "Not Found";
                    newRow[destIndex] = department;
                    console.log(`Row ${i}: Department lookup for ${staffId} = "${department}"`);
                } else if (sourceIndex !== -1) {
                    newRow[destIndex] = sourceRow[sourceIndex] || "";
                }
            } else {
                console.log(`Warning: Destination column "${destCol}" not found in table`);
            }
        }

        // Source file name
        let sourceIndex = headers.indexOf(sourceColumnName);
        if (sourceIndex !== -1) {
            newRow[sourceIndex] = sourceFileName;
        }

        console.log(`Row ${i} prepared:`, newRow.filter(cell => cell !== ""));
        newRows.push(newRow);
    }

    
    if (newRows.length > 0) {
        table.addRows(-1, newRows);
        console.log(`Added ${newRows.length} rows with department lookup`);
    }

    
    sheet.getUsedRange()?.getFormat().autofitColumns();
    
    // Debug: Show a few rows from the table to verify
    const tableRange = table.getRangeBetweenHeaderAndTotal();
    const tableValues = tableRange.getValues();
    console.log("First 3 rows from table after update:");
    for (let i = 0; i < Math.min(3, tableValues.length); i++) {
        console.log(`Row ${i}:`, tableValues[i]);
    }
}
