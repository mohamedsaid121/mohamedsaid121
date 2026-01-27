function main(workbook: ExcelScript.Workbook): string {
  // Your actual source column names here
  const expectedSourceColumns = [
    "ClientID",
    "ClientName", 
    "DateOfOrder",
    "TotalValue",
    "ProductCode"
  ];
  
  const sheet = workbook.getActiveWorksheet();
  const usedRange = sheet.getUsedRange();
  
  if (!usedRange) {
    return JSON.stringify({
      sourceFileName: workbook.getName().replace(".xlsx", ""),
      sourceColumns: expectedSourceColumns,
      data: []
    });
  }
  
  const allData = usedRange.getValues() as string[][];
  const headers = allData[0];
  
  // Get column indices
  const columnIndices = expectedSourceColumns.map(colName => 
    headers.findIndex(h => h.toLowerCase() === colName.toLowerCase())
  );
  
  // Extract data
  const extractedData = allData.slice(1).map(row => 
    columnIndices.map(index => index !== -1 ? row[index] || "" : "")
  );
  
  const result = {
    sourceFileName: workbook.getName().replace(/\.xlsx$/i, ""),
    sourceColumns: expectedSourceColumns,
    data: extractedData
  };
  
  return JSON.stringify(result);
}