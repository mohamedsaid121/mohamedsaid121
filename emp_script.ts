function main(workbook: ExcelScript.Workbook, sourceDataJson: string): void {
    // CONFIGURATION
    let columnMapping: { [key: string]: string } = {
        "Staff ID": "Staff ID",
        "Emp Name": "Full Name",
        "Diagnose": "SHW Feedback ",
        "Number of Sick Days": "Number of Sick Days",
        "Original Diagnose": "Diagnosis",
        <span style="color:blue;font-weight:bold">"Department": "Department"</span>  // ADDED
    };
    
    interface SourceData {
        sourceFileName: string;
        sourceColumns: string[];
        data: string[][];
    }

    const sourceColumnName = "Source";
    
    <span style="color:blue;font-weight:bold">// === EMPLOYEES FILE CONFIGURATION ===
    const EMPLOYEES_FILE_PATH = "/sites/YOURSITE/Shared Documents/Employees.xlsx"; // UPDATE THIS PATH
    const EMPLOYEES_STAFFID_COLUMN = "Users Sys Id"; // Column name in Employees file
    const EMPLOYEES_DEPT_COLUMN = "Department"; // Column name in Employees file</span>

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
    
    <span style="color:blue;font-weight:bold">// === LOAD EMPLOYEES FILE AND CREATE DEPARTMENT MAPPING ===
    console.log("Loading employee departments from: " + EMPLOYEES_FILE_PATH);
    const departmentMap: {[staffId: string]: string} = {};
    
    try {
        // Open Employees workbook
        const employeesWorkbook = ExcelScript.Workbook.open(EMPLOYEES_FILE_PATH);
        const employeesSheet = employeesWorkbook.getActiveWorksheet();
        const employeesRange = employeesSheet.getUsedRange();
        
        if (employeesRange) {
            const allEmployees = employeesRange.getValues() as string[][];
            console.log(`Employees file has ${allEmployees.length} rows`);
            
            // Find Staff ID and Department columns (case-insensitive)
            const empHeaders = allEmployees[0].map(h => 
                h === null || h === undefined ? "" : String(h).trim()
            );
            
            // Look for "Users Sys Id" column (exact match with case-insensitive)
            const staffIdIndex = empHeaders.findIndex(h => 
                h.toLowerCase() === EMPLOYEES_STAFFID_COLUMN.toLowerCase()
            );
            
            const deptIndex = empHeaders.findIndex(h => 
                h.toLowerCase() === EMPLOYEES_DEPT_COLUMN.toLowerCase()
            );
            
            console.log(`Staff ID column ("${EMPLOYEES_STAFFID_COLUMN}") found at index: ${staffIdIndex}`);
            console.log(`Department column ("${EMPLOYEES_DEPT_COLUMN}") found at index: ${deptIndex}`);
            
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
                console.log(`ERROR: Could not find "${EMPLOYEES_STAFFID_COLUMN}" column in Employees file`);
                console.log("Available columns:", empHeaders);
            }
        } else {
            console.log("ERROR: Employees file is empty");
        }
    } catch (error) {
        console.log(`ERROR loading Employees file: ${error}`);
        console.log("Continuing without department data");
    }</span>

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
                <span style="color:blue;font-weight:bold">if (destCol === "Department") {
                    // Lookup department from Employees file
                    const staffId = sourceRow[0] || ""; // Staff ID from first column
                    newRow[destIndex] = departmentMap[staffId] || "";
                } else {</span>
                    newRow[destIndex] = sourceRow[sourceIndex] || "";
                <span style="color:blue;font-weight:bold">}</span>
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
        <span style="color:blue;font-weight:bold">console.log(`Added ${newRows.length} rows with department lookup`);</span>
    }

    
    sheet.getUsedRange()?.getFormat().autofitColumns();
}