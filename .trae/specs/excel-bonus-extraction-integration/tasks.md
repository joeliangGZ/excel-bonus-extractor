# Tasks
- [x] Task 1: 项目初始化与环境配置
  - [x] SubTask 1.1: 创建固定输入文件夹（如：input_bonus）
  - [x] SubTask 1.2: 确认并安装所需依赖包（pandas、openpyxl等Excel处理库）
- [x] Task 2: 实现文件与工作表遍历功能
  - [x] SubTask 2.1: 编写代码遍历输入文件夹内所有.xlsx/.xls格式文件
  - [x] SubTask 2.2: 实现读取单个Excel文件所有工作表的功能
- [x] Task 3: 实现数据提取逻辑
  - [x] SubTask 3.1: 编写代码提取ID、姓名两列固定数据
  - [x] SubTask 3.2: 实现匹配列名包含「实发奖金」列的功能
  - [x] SubTask 3.3: 处理缺失列的异常情况
- [x] Task 4: 实现数据整合与规范输出
  - [x] SubTask 4.1: 将所有提取的数据汇总到单个数据结构中
  - [x] SubTask 4.2: 规范输出列名为：ID、姓名、实发奖金
  - [x] SubTask 4.3: 实现导出汇总数据到新Excel文件的功能
- [x] Task 5: 功能测试与验证
  - [x] SubTask 5.1: 准备测试用Excel文件，包含多文件多工作表场景
  - [x] SubTask 5.2: 运行程序验证数据提取、整合、输出是否符合要求
  - [x] SubTask 5.3: 验证相同ID数据原样保留规则是否生效

# Task Dependencies
- Task 2 依赖于 Task 1 完成
- Task 3 依赖于 Task 2 完成
- Task 4 依赖于 Task 3 完成
- Task 5 依赖于 Task 4 完成
