const XLSX = require('xlsx');

// 文件路径定义
const Mappath = `D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本关卡配置表@MapTranslationConfiguration.xlsx`;
const Totalpath = 'D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本配置表@TotalTranslationConfiguration.xlsx';
const Systempath = 'D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本系统配置表@SystemTranslationConfiguration.xlsx';
const Opspath = 'D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本运营配置表@OpsEvenTranslationConfiguration.xlsx';
const Battlepath = 'D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本战斗配置表@BattleTranslationConfiguration.xlsx';

const MSpath = './20250306_161140.xlsm';

// 读取 Excel 文件
function readExcel(filePath, fileName) {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet).map(item => ({ ...item, "来源": fileName }));

    return data;
}

// 读取5个 Excel 文件
const MapData = readExcel(Mappath, "MapTranslationConfiguration");
const TotalData = readExcel(Totalpath, "TotalTranslationConfiguration");
const SystemData = readExcel(Systempath, "SystemTranslationConfiguration");
const OpsData = readExcel(Opspath, "OpsEvenTranslationConfiguration");
const BattleData = readExcel(Battlepath, "BattleTranslationConfiguration");
const MSData = readExcel(MSpath, "MSpath");

// 合并数据
let combinedData = [...TotalData, ...MapData, ...SystemData, ...OpsData, ...BattleData];

// 对比 MSData 与 combinedData 并仅保留 Translate 与 Simp_TIMI 不同的项
function findMatchingEntries(combinedData, MSData) {
    const result = [];

    combinedData.forEach(item => {
        // 寻找 MSData 中 Label 与 combinedData 中 Key 匹配的项
        const match = MSData.find(msItem => msItem.Label === item.Key);

        if (match && item.Translate !== match.Simp_TIMI) { // 确保 Translate 与 Simp_TIMI 不相同
            // 创建新对象
            const newObj = {
                key: item.Key,
                '266文本': item.Translate,
                'ms文本': match.Simp_TIMI,
                '266负责人': item.ToolRemark,
                'ms负责人': match['Simp_TIMI (Comment)'],
                '266来源': item.来源  // 添加 "266来源" 字段
            };

            result.push(newObj);
        }
    });

    return result;
}

// 使用函数处理数据
const comparisonResult = findMatchingEntries(combinedData, MSData);

// 创建新的工作簿和工作表
const newWorkbook = XLSX.utils.book_new();
const worksheetData = comparisonResult.map(item => ({
    key: item.key,
    '266文本': item['266文本'],
    'ms文本': item['ms文本'],
    '266负责人': item['266负责人'],
    'ms负责人': item['ms负责人'],
    '266来源': item['266来源']  // 添加到输出数据中
}));
const ws = XLSX.utils.json_to_sheet(worksheetData);

// 将工作表添加到工作簿中
XLSX.utils.book_append_sheet(newWorkbook, ws, 'ComparisonResult');

// 写入文件
const outputPath = './key_266文本_ms文本_266负责人_ms负责人_266来源.xlsx'; // 输出文件路径
XLSX.writeFile(newWorkbook, outputPath);

console.log(`对比结果已保存至 ${outputPath}`);