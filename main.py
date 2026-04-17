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

def find_header_row(df):
    """查找表头所在行，返回行索引和列名映射"""
    for idx, row in df.iterrows():
        # 将行内容转为字符串，查找是否包含必填列
        row_values = [str(val).strip() for val in row.values]
        
        # 检查是否包含ID和姓名列
        has_id = any(col == "ID" for col in row_values)
        has_name = any(col == "姓名" for col in row_values)
        has_bonus = any(BONUS_COLUMN_KEYWORD in col for col in row_values)
        
        if has_id and has_name and has_bonus:
            # 构建列名映射
            column_mapping = {}
            for col_idx, col_name in enumerate(row_values):
                column_mapping[col_idx] = col_name
            return idx, column_mapping
    
    return None, None

def extract_data_from_sheet(df):
    """从单个工作表中提取所需数据，自动查找表头位置"""
    # 先检查列名本身是否已经是表头（正常情况，表头在第一行）
    column_names = [str(col).strip() for col in df.columns]
    has_id = any(col == "ID" for col in column_names)
    has_name = any(col == "姓名" for col in column_names)
    has_bonus = any(BONUS_COLUMN_KEYWORD in col for col in column_names)
    
    if has_id and has_name and has_bonus:
        # 列名本身就是表头
        id_col_idx = next(i for i, name in enumerate(column_names) if name == "ID")
        name_col_idx = next(i for i, name in enumerate(column_names) if name == "姓名")
        bonus_col_idx = next(i for i, name in enumerate(column_names) if BONUS_COLUMN_KEYWORD in name)
        
        data_rows = df
    else:
        # 在行数据中查找表头
        header_row_idx, column_mapping = find_header_row(df)
        if header_row_idx is None:
            return None, "未找到包含ID、姓名和实发奖金的表头行"
        
        # 获取列索引
        id_col_idx = next(i for i, name in column_mapping.items() if name == "ID")
        name_col_idx = next(i for i, name in column_mapping.items() if name == "姓名")
        bonus_col_idx = next(i for i, name in column_mapping.items() if BONUS_COLUMN_KEYWORD in name)
        
        # 从表头下一行开始提取数据
        data_rows = df.iloc[header_row_idx + 1:]
    
    # 提取所需列
    extracted_data = []
    for _, row in data_rows.iterrows():
        id_val = row.iloc[id_col_idx]
        name_val = row.iloc[name_col_idx]
        bonus_val = row.iloc[bonus_col_idx]
        
        # 过滤无效行：ID或姓名为空的行（包括总计行、空行）
        if pd.isna(id_val) or pd.isna(name_val):
            continue
            
        # 转换为字符串，去除前后空格
        id_str = str(id_val).strip()
        name_str = str(name_val).strip()
        bonus_str = str(bonus_val).strip() if pd.notna(bonus_val) else ""
        
        # 过滤表头/子表头行：包含这些关键字的行都不是有效数据
        invalid_keywords = ["ID", "姓名", "实发奖金", "合计", "总计", "汇总", "总得分", "得分", "完成率", "目标", "备注", "说明", "字段"]
        if (any(keyword in id_str for keyword in invalid_keywords) or 
            any(keyword in name_str for keyword in invalid_keywords) or
            any(keyword in bonus_str for keyword in invalid_keywords)):
            # 遇到新的表头行，说明后面是其他区域数据，停止处理当前工作表所有后续行
            break
            
        # 过滤ID不是有效数字/编号的情况（ID应该是数字或者字母加数字，不会是纯中文）
        if len(id_str) > 0 and all('\u4e00' <= c <= '\u9fff' for c in id_str):
            continue
            
        extracted_data.append({
            "ID": id_val,
            "姓名": name_str,
            "实发奖金": bonus_val
        })
    
    extracted_df = pd.DataFrame(extracted_data)
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
