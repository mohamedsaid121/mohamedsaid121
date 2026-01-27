function main(workbook: ExcelScript.Workbook, sourceDataJson: string): void {
  // CONFIGURATION
  const columnMapping: { [key: string]: string } = {
    "Customer ID": "ClientID",
    "Customer Name": "ClientName",
    "Order Date": "DateOfOrder",
    "Amount": "TotalValue",
    "Product SKU": "ProductCode"
  };
  
  const tableName = "ConsolidatedTable";
  const sourceColumnName = "Source File";
  
  // Parse with type assertion
  const parsed = JSON.parse(sourceDataJson) as any;
  const sourceFileName: string = parsed.sourceFileName;
  const sourceColumns: string[] = parsed.sourceColumns;
  const rowsToAppend: string[][] = parsed.data;
  
  if (!rowsToAppend || rowsToAppend.length === 0) {
    console.log("No data to append");
    return;
  }
  
  // Get worksheet
  let sheet = workbook.getWorksheet("Consolidated Data");
  if (!sheet) {
    sheet = workbook.addWorksheet("Consolidated Data");
  }
  
  // Get or create table
  let table = sheet.getTable(tableName);
  if (!table) {
    // Create headers
    const headers: string[] = Object.keys(columnMapping).concat(sourceColumnName);
    sheet.getRange("A1").getResizedRange(0, headers.length - 1).setValues([headers]);
    
    // Create table
    table = sheet.addTable(
      sheet.getRange("A1").getResizedRange(0, headers.length - 1),
      true
    );
    table.setName(tableName);
  }
  
  // Get table data
  const tableRange = table.getRange();
  const tableValues = tableRange.getValues() as string[][];
  const tableHeaders = tableValues[0];
  
  // Prepare data for appending
  const newRows: string[][] = [];
  
  for (let i = 0; i < rowsToAppend.length; i++) {
    const sourceRow = rowsToAppend[i];
    const newRow: string[] = new Array(tableHeaders.length).fill("");
    
    // Map columns
    for (const destCol in columnMapping) {
      const sourceColName: string = columnMapping[destCol];
      const sourceIndex = sourceColumns.indexOf(sourceColName);
      const destIndex = tableHeaders.indexOf(destCol);
      
      if (sourceIndex !== -1 && destIndex !== -1) {
        newRow[destIndex] = sourceRow[sourceIndex] || "";
      }
    }
    
    // Add source file name
    const sourceIndex = tableHeaders.indexOf(sourceColumnName);
    if (sourceIndex !== -1) {
      newRow[sourceIndex] = sourceFileName;
    }
    
    newRows.push(newRow);
  }
  
  // Add rows to table
  if (newRows.length > 0) {
    table.addRows(-1, newRows);
  }
  
  // Auto-fit
  sheet.getUsedRange()?.getFormat().autofitColumns();
}