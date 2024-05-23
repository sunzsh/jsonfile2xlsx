const fs = require('fs') 
const ExcelJS = require('exceljs')
const path = require("path")

const args = process.argv.slice(2);
if (!args || args.length < 2) {
    console.log('缺少参数！');
    process.exit(1);
}

const conversionType = args[0]; // New argument for conversion direction

function translateColumnName(num) {
    let columnName = '';
    do {
        let every = num % 26;
        columnName = String.fromCharCode(every + 64) + columnName;
        num -= every
        num = num / 26
    } while (num > 0)
    
    return columnName;
}

args.slice(1).forEach(filePath => {
    doConvert(filePath, conversionType)
});

function soildBorderForRow(row) {
    row._cells.forEach(cell => {
        cell.border = {
            top: {style:'thin'},
            left: {style:'thin'},
            bottom: {style:'thin'},
            right: {style:'thin'}
        };
    });
}

function buildTitle(sheet, array) {

    const columns = [];

    const existsInColumns = function(k) {
        return columns.findIndex((n) => n.key == k) >= 0;
    }

    array.forEach(item => {
        Object.keys(item).forEach(key => {
            if (existsInColumns(key)) {
                return;
            }
            columns.push({ header: key, key: key, width: 20 });
        });
    });
    sheet.columns = columns;

    const titleRow = sheet.getRow(1);
    titleRow.font = {  size: 14,  bold: true };
    titleRow.height = 30;
    titleRow.alignment = { vertical: 'middle', horizontal: 'center' };
    soildBorderForRow(titleRow);
}

function convertXlsxToJson(filePath) {
    const workbook = new ExcelJS.Workbook();
    return workbook.xlsx.readFile(filePath)
        .then(() => {
            const worksheet = workbook.getWorksheet(1);
            const json = [];
            worksheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
                const rowValue = row.values;
                rowValue.shift(); // Remove the first element which is undefined due to ExcelJS indexing
                json.push(rowValue);
            });
            return json;
        });
}

function doConvert(filePath, conversionType) {
    if (conversionType === 'json2xlsx') {
        let array = readArray(filePath)
        if (!array || array.length == 0) {
            return ;
        }

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Sheet1', {views:[{xSplit: 1}]});
        buildTitle(sheet, array);
        
        array.forEach(item => {
            const row = sheet.addRow(item);
            row.font = {  size: 14 };
            row.alignment = { vertical: 'middle', horizontal: 'center' };
            row.height = 30;
            soildBorderForRow(row);
        });

        // 自动筛选器
        sheet.autoFilter = {
            from: 'A1',
            to: {
              row: 1,
              column: sheet.columns.length
            }
          }
        sheet.properties.defaultRowHeight = 30;

        // 设置条纹背景
        sheet.addConditionalFormatting({
            ref: `A1:${translateColumnName(sheet.columns.length)}${array.length}`,
            rules: [
                {
                type: 'expression',
                formulae: ['MOD(ROW(),2)=0'],
                style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'CECECE'}}},
                }
            ]
        })

        let xlsxFilePath = path.join(path.dirname(filePath), path.basename(filePath, ".txt") + ".xlsx", );
        workbook.xlsx.writeFile(xlsxFilePath);
    } else if (conversionType === 'xlsx2json') {
        convertXlsxToJson(filePath).then(json => {
            const jsonFilePath = path.join(path.dirname(filePath), path.basename(filePath, ".xlsx") + ".json");
            fs.writeFileSync(jsonFilePath, JSON.stringify(json, null, 2));
        });
    }
}


function readArray(filePath) {
    if(!fs.existsSync(filePath)) {
        console.log(`文件不存在：${filePath}`);
        return null;
    }
    const data = fs.readFileSync(filePath, 'utf8')

    let contentArray;
    try {
        contentArray = (new Function("return " + data))();
    } catch (e) {
        console.log('文件格式不符合JSON标准 !');
        return null;
    }
    return contentArray
}
