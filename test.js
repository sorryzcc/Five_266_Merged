const XLSX = require('xlsx');

// 文件路径定义
const excelpath = './key_266文本_ms文本_266负责人_ms负责人_266来源.xlsm';

// 读取 Excel 文件
function readExcel(filePath, fileName) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet).map(item => ({ ...item, "来源": fileName }));

    return data;
}

const MSData = readExcel(excelpath, "excelpath");

console.log(MSData,'MSData');
