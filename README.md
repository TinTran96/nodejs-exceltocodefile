
# nodejs-exceltocodefile

Generate exam code file (`.c`, `.cpp`,...) based on Google Form XLXS file.


## Authors

- [@TinTran](https://github.com/TinTran96)


## Installation
Dev on `Node v14.16.1`
```bash
  npm install
```

Setup config file
```
{ 
    "xlsxFile": "./test.xlsx",
	"dir": "./exam",
	"fileExt": ".c"
}
```
`xlsxFile` is xlxs file name.

`dir` is directory name contain exam code file. 

`fileExt` is file extension, here is `.c` file.

Then run the script

```bash
  npm run start
```
![alt text](./demo/xlsx.png 'xlsx example')
![alt text](./demo/result.png 'result')

    