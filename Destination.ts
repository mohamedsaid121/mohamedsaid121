
function main(workbook: ExcelScript.Workbook, sheetName: string, in_sourceColumns: string[], in_data: string [][]) {

    console.log("script 2 received");
    console.log("type of in_sourcecolumns", typeof in_sourceColumns);
    console.log("in_sourcecolumns:", JSON.stringify(in_sourceColumns));
    console.log("type of in_data", typeof in_data);
    console.log("in_data:", JSON.stringify(in_data));

    let sheet = workbook.getWorksheet(sheetName);
    if (!sheet){
        console.log(`sheet "${sheetName}" not found`);
        return;
    }

    const columnMapping: {[key: string]: string} = {
        'Staff ID':'Staff ID',
        'Full Name':'Emp Name',
        'Number of Sick Days':'Number of Sick Days',
        'Diagnosis':'Original Diagnose',
        'SHW Feedback':'Diagnose'
    };

    let table = sheet.getTables()[0];
    if(!table){
        console.log('No table found in the destination sheet');
        return;
    }

    in_sourceColumns.map((sourceColumnName, index) => {

        let trimmedSourceColumn = sourceColumnName.trim()
        let destinationColumnName = columnMapping[trimmedSourceColumn];

        if(!destinationColumnName){
            console.log(`No mapping for "${trimmedSourceColumn}"`)
            return;
        }

        if(!in_data[index] || in_data[index].length == 0){
            console.log(`No data for "${trimmedSourceColumn}"`);
            return;
        }

        let column = table.getColumnByName(destinationColumnName);
        if (column){
            let startRow = column.getRangeBetweenHeaderAndTotal().getRowCount();
            let dataRange: (string | number | boolean)[][] = in_data[index].map(value => [value]);

            column.getRangeBetweenHeaderAndTotal()
                .getCell(startRow, 0)
                .getResizedRange(dataRange.length -1, 0)
                .setValues(dataRange);

            console.log(`${dataRange.length} rows have been written to ${destinationColumnName}`); 
        }
    });
    return {
        message: "Data appended successfully"
    };
}
