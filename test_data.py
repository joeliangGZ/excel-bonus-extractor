import pandas as pd
import os

# 创建测试用Excel文件
test_folder = "input_bonus"

# 测试文件1
df1_sheet1 = pd.DataFrame({
    "ID": [1, 2, 3, 4],
    "姓名": ["张三", "李四", "王五", "赵六"],
    "部门": ["技术部", "市场部", "财务部", "人事部"],
    "实发奖金": [5000, 6000, 4500, 5500]
})

df1_sheet2 = pd.DataFrame({
    "ID": [3, 4, 5, 6],
    "姓名": ["王五", "赵六", "孙七", "周八"],
    "加班天数": [3, 2, 5, 1],
    "季度实发奖金": [1200, 800, 2000, 500]
})

with pd.ExcelWriter(os.path.join(test_folder, "2024年Q1奖金表.xlsx")) as writer:
    df1_sheet1.to_excel(writer, sheet_name="技术部", index=False)
    df1_sheet2.to_excel(writer, sheet_name="其他部门", index=False)

# 测试文件2
df2_sheet1 = pd.DataFrame({
    "ID": [1, 2, 7, 8],
    "姓名": ["张三", "李四", "吴九", "郑十"],
    "入职日期": ["2022-01-15", "2023-03-20", "2021-05-10", "2024-02-01"],
    "实发奖金（年终）": [15000, 18000, 12000, 3000]
})

df2_sheet2 = pd.DataFrame({
    "ID": [9, 10, 1, 2],
    "姓名": ["冯十一", "陈十二", "张三", "李四"],
    "岗位": ["经理", "主管", "工程师", "工程师"],
    "实发奖金": [8000, 7000, 2000, 2500]
})

with pd.ExcelWriter(os.path.join(test_folder, "2024年年终奖金表.xlsx")) as writer:
    df2_sheet1.to_excel(writer, sheet_name="正式员工", index=False)
    df2_sheet2.to_excel(writer, sheet_name="管理团队", index=False)

print("测试文件创建成功，已放入input_bonus文件夹")
print("测试数据说明:")
print("- 共2个Excel文件，4个工作表")
print("- ID为1的张三出现了3次，ID为2的李四出现了3次")
print("- 奖金列表有不同的列名，都包含'实发奖金'关键字")
print("- 最终汇总应该有 4 + 4 + 4 + 4 = 16 条记录")
