import pandas as pd

# 文件路径定义
file_paths = {
    "MapTranslationConfiguration": 'D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本关卡配置表@MapTranslationConfiguration.xlsx',
    "TotalTranslationConfiguration": 'D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本配置表@TotalTranslationConfiguration.xlsx',
    "SystemTranslationConfiguration": 'D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本系统配置表@SystemTranslationConfiguration.xlsx',
    "OpsEvenTranslationConfiguration": 'D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本运营配置表@OpsEvenTranslationConfiguration.xlsx',
    "BattleTranslationConfiguration": 'D:/PM_Mainland_Trunk_20230321_r552586/PMGameClient/Tables/ResXlsx/266.国内文本战斗配置表@BattleTranslationConfiguration.xlsx',
    "MSpath": './筛选过后的0419表格.xlsx'
}

# 读取 Excel 文件并添加来源列
def read_excel_with_source(file_path, source_name):
    df = pd.read_excel(file_path)
    df['来源'] = source_name  # 添加来源列
    return df

# 解析负责人字段
def parse_person_in_charge(field_value):
    if isinstance(field_value, str) and '负责人：' in field_value:
        return field_value.split('负责人：')[-1].strip()
    return ''

# 读取所有需要的 Excel 表格
data_frames = [
    (read_excel_with_source(file_paths["TotalTranslationConfiguration"], "TotalTranslationConfiguration"), "TotalTranslationConfiguration"),
    (read_excel_with_source(file_paths["MapTranslationConfiguration"], "MapTranslationConfiguration"), "MapTranslationConfiguration"),
    (read_excel_with_source(file_paths["SystemTranslationConfiguration"], "SystemTranslationConfiguration"), "SystemTranslationConfiguration"),
    (read_excel_with_source(file_paths["OpsEvenTranslationConfiguration"], "OpsEvenTranslationConfiguration"), "OpsEvenTranslationConfiguration"),
    (read_excel_with_source(file_paths["BattleTranslationConfiguration"], "BattleTranslationConfiguration"), "BattleTranslationConfiguration")
]

# 合并 combinedData
combined_data = pd.concat([df for df, _ in data_frames], ignore_index=True)

# 读取 MSData
ms_data = pd.read_excel(file_paths["MSpath"])

# 确保 Key 和 Label 都存在
common_keys = set(combined_data['Key'].dropna()).intersection(set(ms_data['Label'].dropna()))

# 创建一个新的 DataFrame 来存储结果
comparison_result = []

for key in common_keys:
    combined_row = combined_data[combined_data['Key'] == key].iloc[0]
    ms_row = ms_data[ms_data['Label'] == key].iloc[0]

    # 检查是否包含所需列
    if 'Simp_TIMI' not in ms_row or 'Simp_Chinese' not in ms_row:
        print(f"Warning: Label {key} does not have Simp_TIMI or Simp_Chinese in MSData.")
        continue

    if combined_row['Translate'] != ms_row['Simp_TIMI']:
        comparison_result.append({
            'Label（key）': key,
            'Simp_TIMI': ms_row['Simp_TIMI'],
            'Simp_Chinese': ms_row['Simp_Chinese'],
            'Simp_TIMI (Comment)': ms_row.get('Simp_TIMI (Comment)', ''),
            'ms负责人': parse_person_in_charge(ms_row.get('Simp_TIMI (Comment)', '')),
            '266文本': combined_row['Translate'],
            '266备注': combined_row['ToolRemark'],
            '266负责人': parse_person_in_charge(combined_row['ToolRemark']),
            '266来源': combined_row['来源']
        })

# 将结果保存到新的 Excel 文件中
output_df = pd.DataFrame(comparison_result)
output_df.to_excel('./test3.xlsx', index=False, sheet_name='ComparisonResult')

print("对比结果已保存至 ./test3.xlsx")