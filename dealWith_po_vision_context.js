const XLSX = require('xlsx');

// 文件路径定义
const updatedMergedPath = './compare_10_17.xlsx';

// 读取 Excel 文件
function readExcel(filePath) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet);
}

// 定义一个函数来解析 ToolRemark 并返回 po, version 和 Context
function parseToolRemark(toolRemark) {
    // 确保 toolRemark 是字符串
    if (typeof toolRemark !== 'string') {
        toolRemark = '';
    }

    // 正则表达式匹配场景、版本和负责人
    const match = toolRemark.match(/场景：(.*?) 使用版本：(.*?) 负责人：(.*)/);

    let context, po, version;

    if (match && match.length >= 4) { // 匹配成功且至少有三个捕获组
        context = match[1].trim();
        version = match[2].trim();
        po = match[3].trim();
    } else {
        // 如果没有匹配到预期格式，则尝试直接获取值
        if (toolRemark.includes('场景：')) {
            const parts = toolRemark.split('使用版本：');
            context = parts[0].replace('场景：', '').trim();
            if (parts[1] && parts[1].includes('负责人：')) {
                version = parts[1].split('负责人：')[0].trim();
                po = parts[1].split('负责人：')[1].trim();
            } else {
                version = '';
                po = '';
            }
        } else {
            // 如果不包含 场景：，则整个字段都是 Context
            context = toolRemark.trim();
            version = '';
            po = '';
        }
    }

    return { po, version, context };
}

// 读取 Updated_Merged.xlsx 文件
const updatedMergedData = readExcel(updatedMergedPath);

// 遍历 updatedMergedData，添加 po, version, context 和 origin 字段
const processedData = updatedMergedData.map(item => {
    const { po, version, context } = parseToolRemark(item.ToolRemark || '');
    // 添加 Origin 字段，值为 Translate 字段的值
    return { ...item, po, version, context, Origin: item.Translate ? item.Translate : '' };
});

console.log(processedData, 'processedData');

// 将 JSON 数据转换为工作表
const worksheet = XLSX.utils.json_to_sheet(processedData);

// 创建一个新的工作簿并添加工作表
const workbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

// 写入 Excel 文件
XLSX.writeFile(workbook, 'dealWith_po_vision_context.xlsx');

console.log(`已生成包含新增Key和解析信息的Excel文件: dealWith_po_vision_context.xlsx`);