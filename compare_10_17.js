const XLSX = require('xlsx');

const path10 = './Merged_five_266.xlsx';
const path17 = './266 0217.xlsx';

// 读取 Excel 文件
function readExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet);
}

// 提取所有Key字段的值
function getAllKeys(dataArray) {
    let keysSet = new Set();
    dataArray.forEach(item => {
        if (item.Key !== undefined) {
            keysSet.add(item.Key);
        }
    });
    return Array.from(keysSet);
}

// 对比两个数组对象的Key字段值，找到第二个数组中相对于第一个数组新增的Key值
function findNewKeys(oldKeys, newKeys) {
    return newKeys.filter(key => !oldKeys.includes(key));
}

// 读取数据
const path10Data = readExcel(path10);
const path17Data = readExcel(path17);

console.log(`Path10 数据条数: ${path10Data.length}`);
console.log(`Path17 数据条数: ${path17Data.length}`);

// 获取所有的Key字段值
const oldKeys = getAllKeys(path10Data);
const newKeys = getAllKeys(path17Data);

console.log(`Path10 中发现的Key数量: ${oldKeys.length}`);
console.log(`Path17 中发现的Key数量: ${newKeys.length}`);

// 找到path17Data中新增的Key值
const additionalKeys = findNewKeys(oldKeys, newKeys);

console.log(`新增的Key数量为：${additionalKeys.length}`);
if (additionalKeys.length === 0) {
    console.log("未发现新增的Key。");
} else {
    console.log(`新增的Key为：\n${additionalKeys.join('\n')}`);
}

// 根据新增的Key值从path17Data中提取对应的记录
const newData = path17Data.filter(item => additionalKeys.includes(item.Key));

// 创建新的工作簿和工作表
const newWorkbook = XLSX.utils.book_new();
const newWorksheet = XLSX.utils.json_to_sheet(newData);

// 将工作表添加到工作簿并保存
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Sheet1');
XLSX.writeFile(newWorkbook, './compare_10_17.xlsx');

console.log(`已生成包含新增Key的Excel文件: compare_10_17.xlsx`);