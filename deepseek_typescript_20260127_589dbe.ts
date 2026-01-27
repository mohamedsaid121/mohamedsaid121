function main(
  workbook: ExcelScript.Workbook,
  sourceDataJson: string
): void {
  
  // HARDCODED CONFIGURATION - Edit these 3 lines only
  const columnMapping = {
    "Customer ID": "ClientID",
    "Customer Name": "ClientName",
    "Order Date": "DateOfOrder", 
    "Amount": "TotalValue",
    "Product SKU": "ProductCode"
  };
  const tableName = "ConsolidatedTable";
  const sourceColumnName = "Source File";
  // END OF CONFIGURATION
  
  // Parse source data
  const sourceData = JSON.parse(sourceDataJson);
  const sourceFileName = sourceData.sourceFileName;
  const sourceColumns = sourceData.sourceColumns;
  const rowsToAppend = sourceData.data;
  
  if (rowsToAppend.length === 0) {
    console.log("No data to append");
    return;
  }
  
  // Get or create destination worksheet
  let destSheet = workbook.getWorksheet("Consolidated Data");
  if (!destSheet) {
    destSheet = workbook.addWorksheet("Consolidated Data");
  }
  
  // Get or create the table
  let table = destSheet.getTable(tableName);
  
  if (!table) {
    // Create new table with headers
    const destinationColumns = Object.keys(columnMapping);
    const allHeaders = [...destinationColumns, sourceColumnName];
    
    // Write headers to first row
    const headerRange = destSheet.getRange("A1").getResizedRange(0, allHeaders.length - 1);
    headerRange.setValues([allHeaders]);
    
    // Create table starting at A1
    table = destSheet.addTable(headerRange, true);
    table.setName(tableName);
    
    console.log(`Created new table: ${tableName}`);
  }
  
  // Get table headers
  const tableHeaders = table.getHeaderRowRange().getValues()[0] as string[];
  
  // Prepare data transformation
  const destinationColumns = Object.keys(columnMapping);
  
  // Map source to destination columns
  const columnMappingIndices = destinationColumns.map(destCol => {
    const sourceColName = columnMapping[destCol];
    const sourceIndex = sourceColumns.indexOf(sourceColName);
    const tableIndex = tableHeaders.indexOf(destCol);
    return { sourceIndex, tableIndex };
  });
  
  const sourceColumnTableIndex = tableHeaders.indexOf(sourceColumnName);
  
  // Transform each row using map
  const rowsToAdd = rowsToAppend.map(sourceRow => {
    const newRow = new Array(tableHeaders.length).fill("");
    
    // Map data from source to destination columns
    columnMappingIndices.forEach(({ sourceIndex, tableIndex }) => {
      if (tableIndex !== -1 && sourceIndex !== -1) {
        newRow[tableIndex] = sourceRow[sourceIndex] || "";
      }
    });
    
    // Add source file name
    if (sourceColumnTableIndex !== -1) {
      newRow[sourceColumnTableIndex] = sourceFileName;
    }
    
    return newRow;
  });
  
  // Add all rows to the table
  if (rowsToAdd.length > 0) {
    table.addRows(-1, rowsToAdd);
    console.log(`Added ${rowsToAdd.length} rows from ${sourceFileName}`);
  }
  
  // Auto-fit columns
  destSheet.getUsedRange()?.getFormat().autofitColumns();
}