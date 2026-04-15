# Excel 奖金数据提取整合工具 Spec

## Why
解决手动从多个Excel奖金表中提取和整合数据的低效问题，实现自动化批量处理，减少人工操作错误。

## What Changes
- 创建Python自动化处理程序
- 新增固定输入文件夹用于存放待处理Excel文件
- 新增输出逻辑自动生成汇总后的Excel文件
- 实现多文件多工作表遍历读取功能
- 实现指定列数据提取功能
- 实现数据整合与规范输出功能

## Impact
- Affected specs: 无
- Affected code: 核心Python程序文件、工具依赖配置文件

## ADDED Requirements
### Requirement: Excel 数据提取整合功能
系统 SHALL 提供自动提取多个Excel文件中奖金数据并整合到单个文件的功能。

#### Scenario: 成功用例
- **WHEN** 用户将待处理Excel奖金表放入指定输入文件夹并运行程序
- **THEN** 程序自动读取所有Excel文件及所有工作表，提取ID、姓名和实发奖金列数据，汇总生成新的Excel文件，新文件包含ID、姓名、实发奖金三列，相同ID数据原样保留不合并。

### Requirement: 数据处理规则
系统 SHALL 严格遵循以下数据规则：相同ID的数据直接原样复制，不合并、不去重、不修改。

#### Scenario: 数据规则验证
- **WHEN** 存在多个相同ID的数据记录
- **THEN** 所有记录全部原样保留在最终输出文件中，无任何合并或修改操作。
