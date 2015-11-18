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

// workbook.xlsx.readFile("./DDW_0200C_08-2011.xlsx")
//     .then(function() {
// 		workbook.xlsx.writeFile("data.xlsx")
// 		    .then(function() {
// 		        // done 
// 		    });
// 	});

/*	Data source is 
 * https://data.gov.in/catalog/educational-level-age-and-sex-population-age-7-and-above-2011-india-and-states
 */

workbook.xlsx.readFile("./data.xlsx")
    .then(function() {
		var worksheet = workbook.getWorksheet(1);
		// var dobCol = worksheet.getColumn(4);
		
		// console.log(worksheet.columns);
		// console.log((dobCol.));
		var row = worksheet.getRow(1);
		data = {};
		data.fields = [];
		i = 1;
		row.eachCell(function(cell, colNumber) {
			var col1 = worksheet.getColumn(colNumber);
			var type = ""
			col1.eachCell(function(cell, rowNumber) {
				if(rowNumber == 1)
					return ;
				if( typeof(cell.value) === "string"){
					type = "string";
				}
				else {
					if(type != "string")
						type = "number";
				}
			});
		    // console.log(cell.address);
			data.fields.push({"id":i,"label":cell.value,"pattern":"","type":type});
			i ++ ;
		});
		data.row = [];
		worksheet.eachRow(function(row, rowNumber) {
		    // console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
	    	var new_row = []
		    row.eachCell(function(cell, colNumber) {
		    	new_row.push({"v":cell.value,"f":null});
			    // console.log("Cell " + colNumber + " = " + cell.value 
			    // 	+ " type = " + validator.isInt(cell.value));
			});
			data.row.push({"c":new_row});
		});
		console.log(data);
		// data.row.push({"c":[{"v":"Mushrooms","f":null},{"v":3,"f":null}]})
		console.log(data.row[1]);

		// workbook.xlsx.writeFile("another.xlsx")
		//     .then(function() {
		//         // done 
		//     });
		// worksheet.commit();
		// workbook.commit();
        // use workbook 
    });


