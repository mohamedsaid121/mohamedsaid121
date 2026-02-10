function main(workbook: ExcelScript.Workbook, sourceDataJson: string): void {
    // CONFIGURATION
    let columnMapping: { [key: string]: string } = {
        "Staff ID": "Staff ID",
        "Emp Name": "Full Name",
        "Diagnose": "SHW Feedback ",
        "Number of Sick Days": "Number of Sick Days",
        "Original Diagnose": "Diagnosis",
        "Department": "Department"  // ADDED
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
        // Get the EmployeesData sheet from CURRENT workbook
        const employeesSheet = workbook.getWorksheet("EmployeesData");
        
        if (employeesSheet) {
            const employeesRange = employeesSheet.getUsedRange();
            
            if (employeesRange) {
                const allEmployees = employeesRange.getValues() as string[][];
                console.log(`Found ${allEmployees.length} rows in EmployeesData sheet`);
                
                // Get headers (first row)
                const empHeaders = allEmployees[0].map(h => 
                    h === null || h === undefined ? "" : String(h).trim()
                );
                
                // Find column indices
                const staffIdIndex = empHeaders.findIndex(h => 
                    h.toLowerCase().includes("users sys id") || 
                    h.toLowerCase().includes("users sysid") ||
                    h.toLowerCase().includes("sys id")
                );
                
                const deptIndex = empHeaders.findIndex(h => 
                    h.toLowerCase().includes("department") || 
                    h.toLowerCase().includes("dept")
                );
                
                console.log(`Staff ID column index: ${staffIdIndex}`);
                console.log(`Department column index: ${deptIndex}`);
                
                if (staffIdIndex !== -1) {
                    // Load all employees into dictionary
                    let loadedCount = 0;
                    for (let i = 1; i < allEmployees.length; i++) {
                        const row = allEmployees[i];
                        const staffId = (row[staffIdIndex] || "").toString().trim();
                        if (staffId) {
                            const department = deptIndex !== -1 ? 
                                (row[deptIndex] || "").toString().trim() : "";
                            departmentMap[staffId] = department;
                            loadedCount++;
                        }
                    }
                    console.log(`Loaded ${loadedCount} employee departments into memory`);
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
                if (destCol === "Department") {
                    // Lookup department from EmployeesData sheet
                    const staffId = sourceRow[0] || ""; // Staff ID from first column
                    newRow[destIndex] = departmentMap[staffId] || "";
                } else {
                    newRow[destIndex] = sourceRow[sourceIndex] || "";
                }
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
}
