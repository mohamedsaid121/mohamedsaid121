
function main(workbook: ExcelScript.Workbook, sheetName: string) {
    let sheet = workbook.getWorksheet(sheetName);

    if (!sheet){
        console.log(`sheet "${sheetName}" not found`);
        return;
    }
    let table = sheet.getTables()[0];
    const sourceColumns = ['Staff ID', 'Full Name', 'Number of Sick Days', 'Diagnosis', 'SHW Feedback '];

    if (!table)
    {
        let usedRange = sheet.getUsedRange();

        if(!usedRange){
            console.log("No data found in sheet");
            return;
        }
        //let UsedRangeWithoutSheet = usedRange.split('!')[1];
        //console.log(usedRange);

        table = sheet.addTable(usedRange.getAddress(), true);
    }

    let columnsData: (string | number | boolean)[][] = sourceColumns.map(columnName => {
        try {
            let column = table.getColumnByName(columnName);
            if (column){
                let data: (string | number | boolean)[][] = column.getRangeBetweenHeaderAndTotal().getValues();
                return data.map(row => row[0]);
            } 
            else
            {
                console.log(`column "${columnName}" not found`);
                return [];
            }

        }catch (error) {
            console.log(`error find column "${columnName}" : ${error}`);
            return [];
        }
        
    });

    return {
        sourceColumns: sourceColumns,
        data: columnsData
    };

}
