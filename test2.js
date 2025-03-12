const XLSX = require('xlsx');

// 文件路径定义
const Mappath = `D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本关卡配置表@MapTranslationConfiguration.xlsx`;
const Totalpath = 'D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本配置表@TotalTranslationConfiguration.xlsx';
const Systempath = 'D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本系统配置表@SystemTranslationConfiguration.xlsx';
const Opspath = 'D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本运营配置表@OpsEvenTranslationConfiguration.xlsx';
const Battlepath = 'D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本战斗配置表@BattleTranslationConfiguration.xlsx';

const MSpath = './20250312_130419.xlsm';

// 读取 Excel 文件
function readExcel(filePath, fileName) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(worksheet).map(item => ({ ...item, "来源": fileName }));
}

// 解析 ToolRemark 或 ms负责人 字段以提取负责人名字
function parsePersonInCharge(fieldValue) {
    if (typeof fieldValue !== 'string') {
        fieldValue = '';
    }
    const match = fieldValue.match(/场景：(.*?) 使用版本：(.*?) 负责人：(.*)/);

    if (match && match.length >= 4) {
        return match[3].trim();
    } else {
        if (fieldValue.includes('负责人：')) {
            const parts = fieldValue.split('负责人：');
            return parts[1].trim();
        } else {
            return '';
        }
    }
}

// 读取5个 Excel 文件
const MapData = readExcel(Mappath, "MapTranslationConfiguration");
const TotalData = readExcel(Totalpath, "TotalTranslationConfiguration");
const SystemData = readExcel(Systempath, "SystemTranslationConfiguration");
const OpsData = readExcel(Opspath, "OpsEvenTranslationConfiguration");
const BattleData = readExcel(Battlepath, "BattleTranslationConfiguration");
const MSData = readExcel(MSpath, "MSpath");

let combinedData = [...TotalData, ...MapData, ...SystemData, ...OpsData, ...BattleData];

// 对比 MSData 与 combinedData 并仅保留 Translate 与 Simp_TIMI 不同的项
function findMatchingEntries(combinedData, MSData) {
    const result = [];

    combinedData.forEach(item => {
        const match = MSData.find(msItem => msItem.Label === item.Key);

        if (match && item.Translate !== match.Simp_TIMI) {
            const newObj = {
                key: item.Key,
                '266文本': item.Translate,
                'ms文本': match.Simp_TIMI,
                '266负责人': parsePersonInCharge(item.ToolRemark),
                'ms负责人': parsePersonInCharge(match['Simp_TIMI (Comment)']),
                '266来源': item.来源, // 确保此行存在
                '266_ToolRemark': item.ToolRemark,
                'MS_Simp_TIMI (Comment)': match['Simp_TIMI (Comment)'],
                'MS_Simp_Chinese': match['Simp_Chinese']
            };

            result.push(newObj);
        }
    });

    return result;
}

const comparisonResult = findMatchingEntries(combinedData, MSData);

// 创建新的工作簿和工作表
const newWorkbook = XLSX.utils.book_new();
const worksheetData = comparisonResult.map(item => ({
    'Label（key）': item.key,
    'Simp_TIMI': item['ms文本'],
    'Simp_Chinese': item['MS_Simp_Chinese'],
    'Simp_TIMI (Comment)': item['MS_Simp_TIMI (Comment)'],
    'ms负责人': item['ms负责人'],
    '266文本': item['266文本'],
    '266备注': item['266_ToolRemark'],
    '266负责人': item['266负责人'],
    '266来源': item['266来源'] // 添加这一行来包含'266来源'
}));

const ws = XLSX.utils.json_to_sheet(worksheetData);

// 将工作表添加到工作簿中
XLSX.utils.book_append_sheet(newWorkbook, ws, 'ComparisonResult');

// 写入文件
const outputPath = './test2.xlsx'; // 输出文件路径
XLSX.writeFile(newWorkbook, outputPath);

console.log(`对比结果已保存至 ${outputPath}`);