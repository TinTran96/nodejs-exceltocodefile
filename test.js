const Excel = require('exceljs');
const fs = require('fs');
const colors = require('colors');
const { exit } = require('process');
const { stringify } = require('querystring');
const filename = './file.xlsx';
var workbook = new Excel.Workbook();
const config = loadConfigFile();
const dir = config.dir;
const fileExt = config.fileExt;
const xlsxFile = config.xlsxFile;

try {
	if (!fs.existsSync(xlsxFile)) {
		console.log((xlsxFile + ' not exist. Please check config file or make sure that file at the right location.').red);
		exit();
	}
	fs.writeFileSync('./exam/emptyExamLog.txt', '');
	workbook.xlsx.readFile(xlsxFile)
		.then(function() {
			var worksheet = workbook.getWorksheet('Sheet1');
			worksheet.eachRow({ includeEmpty: false }, function(row, rowNumber) {
				// console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
				if(rowNumber != 1)
				{
					let mssv = row.values[1];
					let content = [row.values[2],row.values[3]];
					writeCFile(mssv, content);
				}
			});
		});
  } catch(err) {
	console.error(err);
  }

function writeCFile(mssv, content){
	let studentFolderDir = dir + '/' + mssv;
	if (!fs.existsSync(dir)){
	    fs.mkdirSync(dir);
	}
	if (!fs.existsSync(studentFolderDir))
	{
		fs.mkdirSync(studentFolderDir);
	}
	let i = 1;
	content.forEach(exam => {
		let output = studentFolderDir + '/cau' + i + fileExt;
		if(exam === undefined){
			console.log((mssv + ' > '+ output + ' Empty ').red);
			fs.appendFileSync('./exam/emptyExamLog.txt', mssv + ' > '+ output + '\n');
		}else{
			fs.writeFile(output, exam, function (err) {
				if (err) return console.log(('Error' + err).red);
				console.log((mssv + ' > '+ output).green);
			});
		}
		i++;
	});
}

function loadConfigFile(){
	let configData = fs.readFileSync('config.json');
	let config = JSON.parse(configData);
	return config;
}
