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
	workbook.xlsx.readFile(xlsxFile)
		.then(function() {
			var worksheet = workbook.getWorksheet('Sheet1');
			worksheet.eachRow({ includeEmpty: false }, function(row, rowNumber) {
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
		fs.writeFile(output, exam, function (err) {
			if (err) return console.log(('Error' + err).red);
			console.log((mssv + ' > '+ output).green);
		  });
		i++;
	});
}

function loadConfigFile(){
	let configData = fs.readFileSync('config.json');
	let config = JSON.parse(configData);
	return config;
}
