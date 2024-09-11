const ExcelJs = require('exceljs');

async function writeExcelTest(searchText, replaceText, filePath){
    const workbook = new ExcelJs.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.getWorksheet('Sheet1'); // This needs to be declared before calling readExcel

    // Call readExcel and wait for its completion
    const output = await readExcel(worksheet, searchText);

    if (output.row !== -1 && output.column !== -1) {
        const cell = worksheet.getCell(output.row, output.column);
        cell.value = replaceText;
        await workbook.xlsx.writeFile(filePath);
        console.log(`Replaced '${searchText}' with '${replaceText}' at row ${output.row}, column ${output.column}.`);
    } else {
        console.log(`Text '${searchText}' not found in the file.`);
    }
}

async function readExcel(worksheet, searchText) {
    let output = { row: -1, column: -1 };

    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            if (cell.value === searchText) {
                output.row = rowNumber;
                output.column = colNumber;
            }
        });
    });

    return output;
}

writeExcelTest("Banana", "Republic", "./files/excellDownloadTest.xlsx");
