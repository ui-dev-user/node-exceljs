var excel = require('exceljs');

var workbook = new excel.Workbook();
var filename = './SampleBook.xlsx';
workbook.xlsx.readFile(filename).then(function () {
	console.log('Reading file.');
	var worksheet = workbook.getWorksheet("Sample");
	
	//adding single row
	worksheet.getRow(4).getCell('B').value = 12001;
	worksheet.getRow(4).getCell('C').value = 'Katy Brod';
	worksheet.getRow(4).getCell('D').value = 'Delivery Head';
	worksheet.getRow(4).getCell('E').value = 19;
	worksheet.getRow(4).getCell('F').value = 'Connecticut';
	worksheet.getRow(4).getCell('G').value = 'Havana';
	worksheet.getRow(4).getCell('H').value = 'New';
    
	
	return workbook.xlsx.writeFile(filename).then(() => {
        console.info('Finished writing to excel.' );
    });

});
