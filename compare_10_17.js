const XLSX = require('xlsx');

const path10 = './Merged_five_266.xlsx';
const path17 = './266 0217';

// 读取 Excel 文件
function readExcel(filePath, fileName) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet).map(item => ({ ...item, "来源": fileName }));

    return data;
}

// 读取2个 Excel 文件
const path10Data = readExcel(path10, "path10");
const path17Data = readExcel(path17, "path17");