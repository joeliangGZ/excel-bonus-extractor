import os
import pandas as pd

# 配置常量
INPUT_FOLDER = "input_bonus"
OUTPUT_FILE = "奖金汇总表.xlsx"
REQUIRED_COLUMNS = ["ID", "姓名"]
BONUS_COLUMN_KEYWORD = "实发奖金"

def get_all_excel_files(folder_path):
    """获取文件夹内所有Excel文件路径"""
    excel_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.lower().endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
                excel_files.append(os.path.join(root, file))
    return excel_files

def extract_data_from_sheet(df):
    """从单个工作表中提取所需数据"""
    # 检查是否存在必填列
    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_columns:
        return None, f"缺少必填列: {', '.join(missing_columns)}"
    
    # 查找实发奖金列
    bonus_columns = [col for col in df.columns if BONUS_COLUMN_KEYWORD in str(col)]
    if not bonus_columns:
        return None, "未找到包含'实发奖金'的列"
    
    # 取第一个匹配的奖金列
    bonus_col = bonus_columns[0]
    
    # 提取所需列
    extracted_df = df[REQUIRED_COLUMNS + [bonus_col]].copy()
    # 重命名奖金列
    extracted_df.columns = ["ID", "姓名", "实发奖金"]
    
    return extracted_df, None

def main():
    print(f"开始处理Excel文件，输入文件夹: {INPUT_FOLDER}")
    
    # 获取所有Excel文件
    excel_files = get_all_excel_files(INPUT_FOLDER)
    if not excel_files:
        print(f"错误: {INPUT_FOLDER} 文件夹中没有找到Excel文件")
        return
    
    print(f"找到 {len(excel_files)} 个Excel文件")
    
    all_data = []
    error_logs = []
    
    # 处理每个文件
    for file_path in excel_files:
        print(f"\n处理文件: {os.path.basename(file_path)}")
        try:
            # 获取所有工作表名
            xls = pd.ExcelFile(file_path)
            sheet_names = xls.sheet_names
            print(f"文件包含 {len(sheet_names)} 个工作表")
            
            # 处理每个工作表
            for sheet_name in sheet_names:
                print(f"  处理工作表: {sheet_name}")
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                
                extracted_data, error = extract_data_from_sheet(df)
                if error:
                    error_msg = f"文件 {os.path.basename(file_path)} 工作表 {sheet_name}: {error}"
                    error_logs.append(error_msg)
                    print(f"    警告: {error}")
                    continue
                
                if not extracted_data.empty:
                    all_data.append(extracted_data)
                    print(f"    成功提取 {len(extracted_data)} 条数据")
                
        except Exception as e:
            error_msg = f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}"
            error_logs.append(error_msg)
            print(f"  错误: {str(e)}")
    
    if not all_data:
        print("\n错误: 没有提取到任何有效数据")
        return
    
    # 合并所有数据
    combined_df = pd.concat(all_data, ignore_index=True)
    print(f"\n数据汇总完成，共 {len(combined_df)} 条记录")
    
    # 导出到Excel
    try:
        combined_df.to_excel(OUTPUT_FILE, index=False)
        print(f"成功生成汇总文件: {OUTPUT_FILE}")
    except Exception as e:
        print(f"导出Excel文件时出错: {str(e)}")
        return
    
    # 输出错误日志
    if error_logs:
        print("\n处理过程中出现以下问题:")
        for error in error_logs:
            print(f"- {error}")
    
    print("\n处理完成!")

if __name__ == "__main__":
    main()
