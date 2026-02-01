function main(workbook: ExcelScript.Workbook, sheetName: string): string {
    const sourceFileName = workbook.getName().replace(/\.xlsx$/i, "");
    const SourceColumns = [
        "Certificate NBR",
        "Individual Name"
    ];

    let ExpectedOutputColumns = [
        "Certificate NBR",
        "Individual Name",
        "Diagnosis"
    ]

    let sheet = workbook.getWorksheet(sheetName);
    if (!sheet){
        console.log(`sheet ${sheetName} not found`);
        return;
    }

    let usedRange = sheet.getUsedRange();
    if (!usedRange) {
        console.log(`No data found in ${sheetName}`);
        return;
    }

    let table = sheet.getTables()[0];
    if(!table){
        console.log(`data is Not in table format`);
        table = sheet.addTable(usedRange.getAddress(), true);
    }

    let rawData = usedRange.getValues() as string[][];
    let headers = rawData[0].map(header => header.trim().toLowerCase());
    let trimmedSourceColumnNames = SourceColumns.map(colName => colName.trim().toLowerCase());

    // source columns indices in excel file without diagnosis columns
    let columnIndices = trimmedSourceColumnNames.map(colName =>
        headers.findIndex(header => header === colName)
    );

    let diagnosisColumnIndices: number[] = [];
    let diagnosisColumnNames: string [] = [];

    for (let i = 0; i < headers.length; i++){
        let headerName = headers[i];
        if (headerName && /^Disease Description(\s*\d*)?$/i.test(headerName)) {
            diagnosisColumnIndices.push(i);
            diagnosisColumnNames.push(headerName);
        }
    }

    let filteredData = rawData.slice(1)
                              .filter(row => row.some(cell =>
        cell !== "" || cell !== null || cell !== undefined
    ));


    let duplicateStaffId = new Set<string>();
    let extractedData: string[][] = [];

    for (let row in filteredData){
        let staffId = row[columnIndices[0]];
        let fulName = row[columnIndices[1]];
        let diagnoses: string[] = [];

        if (staffId && duplicateStaffId.has(staffId)){
            continue;
        }
        if (staffId){
            duplicateStaffId.add(staffId);
        }

        if (diagnosisColumnIndices.length > 0){
            for (let diagIndex of diagnosisColumnIndices){
                if(diagIndex < row.length){
                    let diagnosis = row[diagIndex] || "";
                    if (diagnosis.trim() !== ""){
                        diagnoses.push(diagnosis.trim());
                    }
                }
            }
        }
        if (diagnoses.length > 0){
            for (let diagnosis of diagnoses){
                extractedData.push([staffId, fulName, diagnosis]);
            }
        } else {
            extractedData.push([staffId, fulName, ""]);
        }
    }

    const result = {
        sourceFileName: sourceFileName,
        sourceColumns: ExpectedOutputColumns,
        data: extractedData
    };
    console.log(result);

    return JSON.stringify(result);
}







"result": "{\"sourceFileName\":\"New Microsoft Excel Worksheet\",\"sourceColumns\":[\"Certificate NBR\",\"Individual Name\",\"Diagnosis\"],\"data\":[[null,null,\"\"],[null,null,\"\"],[null,null,\"\"],[null,null,\"\"],[null,null,\"\"],[null,null,\"\"],[null,null,\"\"],[null,null,\"\"],[null,null,\"\"],[null,null,\"\"],[\"0\",null,\"\"],[\"1\",null,\"\"],[\"2\",null,\"\"],[\"3\",null,\"\"],[\"4\",null,\"\"],[\"5\",null,\"\"],[\"6\",null,\"\"],[\"7\",null,\"\"],[\"8\",null,\"\"],[\"9\",null,\"\"]]}
