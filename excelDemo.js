const ExcelJs = require('exceljs');


async function excelTest(){

    let output = {row: -1, column: -1};

    const workbook = new ExcelJs.Workbook();
    await workbook.xlsx.readFile("./files/excellDownloadTest.xlsx");

    const worksheet = workbook.getWorksheet('Sheet1');
    worksheet.eachRow((row, rowNumber) =>
        {
            row.eachCell((cell, colNumber)=>
                {
                    // // print all excel content
                    // console.log(cell.value);
                    
                    // //identify item
                    if(cell.value === "Apple"){
                        console.log(rowNumber) //3
                        console.log(colNumber) //2

                        output.row = rowNumber;
                        output.column = colNumber;
                    }
                })
        })

    const cell = worksheet.getCell(output.row, output.column);
    cell.value = "Iphone";
    await workbook.xlsx.writeFile("./files/excellDownloadTest.xlsx");

}

excelTest();