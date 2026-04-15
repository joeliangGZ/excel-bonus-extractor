import pandas as pd
import os

# 创建表头不在第一行的测试文件
test_folder = "input_bonus"

# 先清空原来的测试文件
for file in os.listdir(test_folder):
    os.remove(os.path.join(test_folder, file))

# 测试情况1：表头在第4行（索引3）
df1 = pd.DataFrame([
    ["2024年Q2奖金表", "", "", ""],
    ["制表人：财务部门", "", "", ""],
    ["日期：2024-06-30", "", "", ""],
    ["ID", "姓名", "部门", "实发奖金"],
    [1, "张三", "技术部", 5500],
    [2, "李四", "市场部", 6500],
    [3, "王五", "财务部", 4800],
    ["", "", "", ""],
    [4, "赵六", "人事部", 5800]
])

with pd.ExcelWriter(os.path.join(test_folder, "表头在第四行.xlsx")) as writer:
    df1.to_excel(writer, sheet_name="Sheet1", index=False, header=False)

# 测试情况2：表头在第2行，列顺序不固定
df2 = pd.DataFrame([
    ["销售部奖金明细", "", "", ""],
    ["姓名", "岗位", "ID", "本月实发奖金", "备注"],
    ["孙七", "销售主管", 5, 7200, "绩效A"],
    ["周八", "销售代表", 6, 6800, "绩效B"],
    ["吴九", "销售代表", 7, 5900, "绩效C"]
])

with pd.ExcelWriter(os.path.join(test_folder, "表头在第二行列顺序不同.xlsx")) as writer:
    df2.to_excel(writer, sheet_name="销售部", index=False, header=False)

# 测试情况3：原来的正常格式（表头在第一行），确保兼容
df3 = pd.DataFrame({
    "ID": [8, 9, 10],
    "姓名": ["郑十", "冯十一", "陈十二"],
    "实发奖金": [4500, 8200, 7300]
})

with pd.ExcelWriter(os.path.join(test_folder, "正常表头格式.xlsx")) as writer:
    df3.to_excel(writer, sheet_name="Sheet1", index=False)

print("测试文件创建完成，包含3种情况:")
print("1. 表头在第4行，中间有空行")
print("2. 表头在第2行，列顺序不固定")
print("3. 正常表头格式（兼容测试）")
print("\n总共应该提取的数据: 4 + 3 + 3 = 10条记录")
