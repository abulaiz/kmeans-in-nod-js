const ExcelJS = require('exceljs');
const kmeans = require('node-kmeans');
let Timeout = null;
let data = [];
let result = null;
let _TIME = {
    START : 0, FINISH_FETCH_DATA : 0, FINISH_CONVERT_DATA : 0, FINISH_CLUSTERING : 0, FINISH_WRITE_FILE : 0
}

function timeToSecond(time){
    return time / 1000;
}

function triggerWhenFinnish(timeout, callback){
    clearTimeout(Timeout);
    Timeout = setTimeout(function(){
        callback();
    }, timeout)
}

async function afterClustered(){
    _TIME.FINISH_CLUSTERING = new Date().getTime();
    const workbook = new ExcelJS.Workbook();
    for(let i in result){
        const sheet = workbook.addWorksheet(`Cluster ${(i+1)}`);
        sheet.addRow([ 'Nama Peserta',  'OUT.', 'ME.',  'COMP.', 'SCI.', 'PERS.', 'AESTH.', 'LIT.', 'MUS.', 'S.S.', 'CLER.', 'PRAC.', 'MED.' ]) 
        for(let j in result[i].clusterInd){
            sheet.addRow(data[result[i].clusterInd[j]])
            sheet.getRow(Number(j)+2).eachCell((cell, cellNumber) => {
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
    }

    await workbook.xlsx.writeFile('hasil_clustering.xlsx');
    _TIME.FINISH_WRITE_FILE = new Date().getTime();
    console.log(`Membaca dataset : ${timeToSecond( _TIME.FINISH_FETCH_DATA - _TIME.START )} detik`)
    console.log(`Konversi data : ${timeToSecond(_TIME.FINISH_CONVERT_DATA - _TIME.FINISH_FETCH_DATA)} detik`)
    console.log(`Clustering data : ${timeToSecond(_TIME.FINISH_CLUSTERING - _TIME.FINISH_CONVERT_DATA)} detik`)
    console.log(`Menulis hasil : ${timeToSecond(_TIME.FINISH_WRITE_FILE - _TIME.FINISH_CLUSTERING)} detik`)
    console.log(`Waktu total : ${timeToSecond(_TIME.FINISH_WRITE_FILE - _TIME.START)} detik`)
    console.log('Proses clustering selesai, silahkan cek file hasil_clustering.xlsx')
}

function afterDataFethced(){
    data.shift();
    
    _TIME.FINISH_FETCH_DATA = new Date().getTime();

	let vectors = [];
	for (let i in data) {
        vectors.push( JSON.parse( JSON.stringify(data[i]) ) );
        vectors[i].shift();
	}    
    
    _TIME.FINISH_CONVERT_DATA = new Date().getTime();

    kmeans.clusterize(vectors, {k: 3}, (err,res) => {
      result = res;
      afterClustered()
    });    
}

(async function(){
    _TIME.START = new Date().getTime();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile('hasil_tes.xlsx');
    workbook.eachSheet(function(worksheet, sheetId) {
        worksheet.eachRow(function(row, rowNumber) {
            data.push( Object.keys(row.values).map((key) => row.values[key]) )
            triggerWhenFinnish(10 , afterDataFethced)
        });
    });

})()