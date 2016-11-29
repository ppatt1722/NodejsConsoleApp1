
var gpdpFormatterOptions = {
	logfile: '',
	testToReqXlsxFilePath : '',
	scriptsFolderPath : '',
	gpdpOutputPath : ''
}

function gpdpFormatter(options) {
	// prepare constants for execution	
	var valueColumnName = 'Requirements ID';
	var keyColumnName = 'Test ID';
	var winston = require('winston');
	if ( (typeof options.logfile) != 'undefined')
		winston.add(winston.transports.File, { filename: logfile });
	
	var mainOptions = {
		logger				: winston,
		valueColumnName		: valueColumName,
		keyColumnName		: keyColumnName			
	}
	
	if ((typeof options.test) != 'undefined' && options.Test) {
		return {
			// add methods for testing here
			testToReqMapping	: ManyToManyLookup,
			sheetToCSV			: SheetToCSV
		}
	}
	else {
		return {
			// public methods and properties
			formatTestResults	: Main
		}
	}
	
	function Main() {
	}
	
	function ManyToManyLookup(workbookPath, options) {
		if (typeof require !== 'undefined') XLSX = require('xlsx');
		var dict = [];
		var workbook = XLSX.readFile(workbookPath);
		var aintThatSomeSheet = workbook.SheetNames[0];
		var worksheet = workbook.Sheets[aintThatSomeSheet];
		// get cell range
		var range = XLSX.utils.decode_range(worksheet["!ref"]);
		var val, R, C, keyColumn = 0, valueColumn = 1, key, value, lastRow = 1;
		
		// iterate rows
		for (R = range.s.r; R <= range.e.r; ++R) {
			// iterate columns
			for (C = range.s.c; C <= range.e.c; ++C) {
				// get current cell value
				val = worksheet[XLSX.utils.encode_cell({ c: C, r: R })];
				// if header row get index of key and value columns
				if (R == 0) {
					if (val.v == options.keyColumnName) keyColumn = C;
					if (val.v == options.valueColumnName) valueColumn = C;
				} else { // not header row
					key = worksheet[XLSX.utils.encode_cell({ c: keyColumn, r: R })].v;
					value = worksheet[XLSX.utils.encode_cell({ c: valueColumn, r: R })].v;
					// initialize value with empty array if not already
					if (dict[key] == null)
						dict[key] = [];
					// push value as one-of-many mapped values for given key
					dict[key].push(value);
					break;
				}
			}
		}
		return dict;
	}
	
	function SheetToCSV(workbookPath, sheet_name, options) {
		var XLSX = require('xlsx');
		var workbook = XLSX.readFile(workbookPath);
		
		var aintThatSomeSheet = sheet_name;
		if (sheet_name == null)
			aintThatSomeSheet = workbook.SheetNames[0];
		
		/* Get worksheet */
		var worksheet = workbook.Sheets[aintThatSomeSheet];
		if (worksheet == null) console.log("The specified worksheet does not exist.");
		
		var csv = XLSX.utils.sheet_to_csv(worksheet);
		
		return csv;
	}
}


function gpdpFormatterTester(testOptions) {
	var map = gpdpFormatter(testOptions.xlsxPath, testOptions);
	
	for (var i in map)
		for (var j in map[i])
			console.log(i + ' => ' + map[i][j]);
}


var testOptions = {
	isFinite		: true,
	xlsxPath		: 'C:\\Users\\pspattillo\\Documents\\Doc\\Stafford Project\\sample output from tool.xlsx',
	keyColumnName	: 'Test ID',
	valueColumnName	: 'Requirements ID'
}

gpdpFormatter(testOptions);