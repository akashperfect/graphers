var Excel = require("exceljs");
var validator = require('validator');
var workbook = new Excel.Workbook();
// workbook.creator = "Me";
// workbook.lastModifiedBy = "Her";
// workbook.created = new Date(1985, 8, 30);
// workbook.modified = new Date();
// var sheet = workbook.addWorksheet("My Sheet");

// var worksheet = workbook.getWorksheet(1);
// Add column headers and define column keys and widths 
// Note: these column structures are a workbook-building convenience only, 
// apart from the column width, they will not be fully persisted. 
// worksheet.columns = [
//     { header: "Id", key: "id", width: 10 },
//     { header: "Name", key: "name", width: 32 },
//     { header: "D.O.B.", key: "DOB", width: 10 }
// ];

workbook.xlsx.readFile("./DDW_0200C_08-2011.xlsx")
    .then(function() {
		workbook.xlsx.writeFile("data.xlsx")
		    .then(function() {
		        // done 
		    });
	});