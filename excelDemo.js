const ExcelJs = require('excel.js');

const workbook = new ExcelJs.Workbook();
const worksheet = workbook.getWorksheet('Sheet1');
worksheet.eachRow((row, rowNumber) =>
    {
        row.eachCell((cell, colNumber)=>
            {
                console.log(cell.value);
            })
    })