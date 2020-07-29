const ExcelJS = require('exceljs');
const faker = require('faker');
const min_value = 40;
const max_value = 80;
const seed_count = 5000;

function generateRandomValue(){
    return Math.floor(Math.random() * (max_value - min_value) + min_value);
}

(async function(){

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Me';
    workbook.lastModifiedBy = 'Her';
    workbook.created = new Date(1985, 8, 30);
    workbook.modified = new Date();
    workbook.lastPrinted = new Date(2016, 9, 27);    

    const sheet = workbook.addWorksheet('Hasil RMIB');
    sheet.addRow([ 'Nama Peserta',	'OUT.',	'ME.',	'COMP.', 'SCI.', 'PERS.', 'AESTH.', 'LIT.', 'MUS.', 'S.S.', 'CLER.', 'PRAC.', 'MED.' ])    

    for(let i = 0; i < seed_count; i++){
        sheet.addRow([
            faker.name.findName(),
            generateRandomValue(), generateRandomValue(), generateRandomValue(), generateRandomValue(),
            generateRandomValue(), generateRandomValue(), generateRandomValue(), generateRandomValue(),
            generateRandomValue(), generateRandomValue(), generateRandomValue(), generateRandomValue()
        ])
        sheet.getRow(Number(i)+2).eachCell((cell, cellNumber) => {
            cell.border = {
                top: {style:'thin', color: {argb:'00000000'}},
                left: {style:'thin', color: {argb:'00000000'}},
                bottom: {style:'thin', color: {argb:'00000000'}},
                right: {style:'thin', color: {argb:'00000000'}}
            };
        })         
    }

    const col = sheet.getColumn(1);
    const row = sheet.getRow(1);
    col.width = 20;
    row.height = 30;
    row.eachCell(function(cell, cellNumber){
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = {
            top: {style:'thin', color: {argb:'00000000'}},
            left: {style:'thin', color: {argb:'00000000'}},
            bottom: {style:'thin', color: {argb:'00000000'}},
            right: {style:'thin', color: {argb:'00000000'}}
        };            
    })

    await workbook.xlsx.writeFile('hasil_tes.xlsx');
    console.log("Test result successfuly seeded")
})()