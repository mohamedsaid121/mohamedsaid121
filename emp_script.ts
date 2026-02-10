function main(workbook: ExcelScript.Workbook, sourceDataJson: string): void {
    // CONFIGURATION
    let columnMapping: { [key: string]: string } = {
        "Staff ID": "Staff ID",           // This maps from Table column to Source column
        "Emp Name": "Full Name",
        "Diagnose": "SHW Feedback ",
        "Number of Sick Days": "Number of Sick Days",
        "Original Diagnose": "Diagnosis",
        "Department": "Department"  // Department doesn't exist in source - will be looked up
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
                
                console.log("Employee sheet headers:", empHeaders);
                
                // Find Staff ID column in EmployeesData - looking for "Users Sys Id"
                const staffIdIndex = empHeaders.findIndex(h => 
                    h.toLowerCase() === "users sys id" || 
                    h.toLowerCase() === "users sysid" ||
                    h.toLowerCase().includes("users") && h.toLowerCase().includes("sys") && h.toLowerCase().includes("id")
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
                            if (loadedCount <= 5) { // Log first 5 for debugging
                                console.log(`Mapped Staff ID: "${staffId}" -> Department: "${department}"`);
                            }
                        }
                    }
                    console.log(`Loaded ${loadedCount} employee departments into memory`);
                } else {
                    console.log("ERROR: Could not find 'Users Sys Id' column in EmployeesData sheet");
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
    const sourceColumns: string[] = parsed.sourceColumns.map(col => col.trim());
    const rowsToAppend: string[][] = parsed.data;

    console.log("Source columns:", sourceColumns);
    console.log("Number of rows to append:", rowsToAppend.length);
    if (rowsToAppend.length > 0) {
        console.log("First row sample:", rowsToAppend[0]);
    }

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

        // Get Staff ID from source row - look for "Staff ID" column
        const sourceStaffIdIndex = sourceColumns.findIndex(col => 
            col.toLowerCase() === "staff id" || 
            col.toLowerCase().includes("staff") && col.toLowerCase().includes("id")
        );
        
        let staffId = "";
        if (sourceStaffIdIndex !== -1) {
            staffId = (sourceRow[sourceStaffIdIndex] || "").toString().trim();
        }
        
        console.log(`Row ${i}: Staff ID from source = "${staffId}"`);

        // Filling Destination rows with source rows
        for (const destCol in columnMapping) {
            const sourceColName: string = columnMapping[destCol];
            const sourceIndex = sourceColumns.indexOf(sourceColName);
            const destIndex = headers.indexOf(destCol);

            if (destIndex !== -1) {
                if (destCol === "Department") {
                    // Lookup department using the Staff ID
                    const department = departmentMap[staffId] || 
                                     (staffId ? "Department Not Found" : "Missing Staff ID");
                    newRow[destIndex] = department;
                    console.log(`Row ${i}: Department lookup for "${staffId}" = "${department}"`);
                } else if (sourceIndex !== -1) {
                    newRow[destIndex] = sourceRow[sourceIndex] || "";
                } else if (destCol === "Staff ID") {
                    // Special case: If Staff ID column in table, use the staffId we extracted
                    newRow[destIndex] = staffId;
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
    const rowsToShow = Math.min(3, tableValues.length);
    console.log(`First ${rowsToShow} rows from table after update:`);
    for (let i = 0; i < rowsToShow; i++) {
        console.log(`Row ${i}:`, tableValues[i]);
    }
}
