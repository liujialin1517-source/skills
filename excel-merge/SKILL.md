---
name: excel-merge
description: 合并两个Excel表格。当用户需要合并Excel文件、将两个Excel的指定列合并到一个文件中、合并Excel数据时使用此插件。默认将第一个Excel的第一列和第二个Excel的第二列合并,输出到桌面。
---

# Excel表格合并插件

## 技能概述

此插件用于将两个Excel文件中的指定列合并到一个新的Excel文件中。默认情况下,会将第一个Excel文件的第一列和第二个Excel文件的第二列合并,并输出到桌面。

## 使用场景

当用户请求以下操作时使用此插件:
- "合并两个Excel表格"
- "将两个Excel文件合并"
- "把Excel的第一列和第二个Excel的第二列合并"
- "合并Excel数据"
- "Excel表格融合"

## 使用方法

### 基本用法

当用户要求合并Excel时,WorkBuddy会:
1. 询问第一个Excel文件的路径
2. 询问第二个Excel文件的路径
3. 询问要合并的列号(可选,默认文件1第1列,文件2第2列)
4. 执行合并操作
5. 将结果保存到桌面,默认文件名为 `merged_excel.xlsx`

### 高级用法

可以指定要合并的列号:
- 文件1的列号: 从1开始,默认为1
- 文件2的列号: 从1开始,默认为2
- 输出文件名: 可以自定义,默认为 `merged_excel.xlsx`

### 使用步骤

#### 步骤 1: 获取Excel文件路径

用户需要提供两个Excel文件的完整路径,例如:
- `C:\Users\用户名\Desktop\data1.xlsx`
- `C:\Users\用户名\Desktop\data2.xlsx`

#### 步骤 2: 确定要合并的列

默认配置:
- 第一个Excel: 第1列
- 第二个Excel: 第2列

用户可以指定其他列号。

#### 步骤 3: 执行合并脚本

运行 `scripts/merge_excel.py` 脚本:

```bash
python scripts/merge_excel.py <文件1路径> <文件2路径> [列1] [列2] [输出文件名]
```

示例:
```bash
# 使用默认设置(文件1第1列 + 文件2第2列)
python scripts/merge_excel.py "C:\Users\Desktop\data1.xlsx" "C:\Users\Desktop\data2.xlsx"

# 自定义列号(文件1第1列 + 文件2第3列)
python scripts/merge_excel.py data1.xlsx data2.xlsx 1 3

# 指定输出文件名
python scripts/merge_excel.py data1.xlsx data2.xlsx 1 2 result.xlsx
```

#### 步骤 4: 查看结果

合并后的文件会自动保存到桌面,文件名为 `merged_excel.xlsx`(或自定义的文件名)。

## 输出格式

合并后的Excel文件包含:
- 第1行: 表头("文件1_第1列" 和 "文件2_第2列")
- 第2行开始: 合并的数据
- 列宽: 自动调整为30,方便查看

如果两个文件的行数不一致,会以较长的文件为准,较短文件不足的部分留空。

## 功能特性

✅ **自动合并**: 一键合并两个Excel文件的指定列
✅ **灵活配置**: 可以指定任意列号
✅ **美化输出**: 自动设置表头样式和列宽
✅ **输出到桌面**: 方便快速找到结果文件
✅ **错误处理**: 友好的错误提示信息

## 脚本说明

- `scripts/merge_excel.py`: Python脚本,使用openpyxl库处理Excel文件
  - 支持读取指定工作表的指定列
  - 自动处理数据对齐
  - 生成格式化的Excel输出文件

## 依赖

- Python 3.x
- openpyxl库 (用于处理Excel文件)

### 安装依赖

```bash
pip install openpyxl
```

## 注意事项

1. **文件格式**: 支持.xlsx格式的Excel文件
2. **列号从1开始**: 不是从0开始
3. **输出位置**: 默认保存到桌面
4. **文件覆盖**: 如果输出文件已存在,会被覆盖
5. **空行处理**: 如果某行为空,会保留空字符串

## 示例场景

### 场景1: 合并学生姓名和成绩

- 文件1: `students.xlsx` - 学生名单在第1列
- 文件2: `scores.xlsx` - 成绩在第2列
- 合并后: 学生姓名和成绩在同一行

### 场景2: 合并产品信息

- 文件1: `products.xlsx` - 产品名称在第1列
- 文件2: `prices.xlsx` - 价格在第2列
- 合并后: 产品名称和价格在同一行

### 场景3: 合并联系方式

- 文件1: `names.xlsx` - 姓名在第1列
- 文件2: `phones.xlsx` - 电话号码在第2列
- 合并后: 姓名和电话号码在同一行
