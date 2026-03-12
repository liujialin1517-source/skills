# Excel表格合并插件

## 插件简介

这是一个用于合并两个Excel表格的WorkBuddy插件。可以将两个Excel文件中的指定列合并到一个新的Excel文件中,并自动输出到桌面。

## 功能特性

- ✅ 支持合并两个Excel文件的任意列
- ✅ 自动美化输出格式(表头样式、列宽)
- ✅ 智能处理不同行数的数据
- ✅ 自动输出到桌面,方便查找
- ✅ 友好的错误提示
- ✅ 支持自定义输出文件名

## 安装方法

### 通过WorkBuddy插件市场安装

1. 在WorkBuddy插件市场中添加仓库: `https://github.com/liujialin1517-source/skills`
2. 刷新插件列表
3. 找到"Excel表格合并"插件并点击安装

### 手动安装

```bash
# 克隆仓库
git clone https://github.com/liujialin1517-source/skills.git

# 复制插件到WorkBuddy目录
cp -r skills/excel-merge ~/.workbuddy/plugins/
```

## 使用方法

### 方法1: 通过WorkBuddy使用

在WorkBuddy中输入以下任一语句:

- "合并两个Excel表格"
- "将两个Excel文件合并"
- "把Excel的第一列和第二个Excel的第二列合并"
- "合并Excel数据"

WorkBuddy会自动识别并调用此插件,然后会询问:
1. 第一个Excel文件的路径
2. 第二个Excel文件的路径
3. 要合并的列号(可选)

### 方法2: 直接运行脚本

```bash
# 基本用法(使用默认设置)
python scripts/merge_excel.py "C:\Users\Desktop\data1.xlsx" "C:\Users\Desktop\data2.xlsx"

# 自定义列号
python scripts/merge_excel.py data1.xlsx data2.xlsx 1 3

# 指定输出文件名
python scripts/merge_excel.py data1.xlsx data2.xlsx 1 2 result.xlsx
```

## 参数说明

| 参数 | 说明 | 必需 | 默认值 |
|------|------|------|--------|
| 文件1路径 | 第一个Excel文件的完整路径 | 是 | - |
| 文件2路径 | 第二个Excel文件的完整路径 | 是 | - |
| 列1 | 第一个文件的列号(从1开始) | 否 | 1 |
| 列2 | 第二个文件的列号(从1开始) | 否 | 2 |
| 输出文件名 | 输出的Excel文件名 | 否 | merged_excel.xlsx |

## 使用示例

### 示例1: 合并学生姓名和成绩

**文件1: students.xlsx**
```
A列
张三
李四
王五
```

**文件2: scores.xlsx**
```
      B列
      90
      85
      95
```

**命令:**
```bash
python scripts/merge_excel.py students.xlsx scores.xlsx 1 2
```

**输出: merged_excel.xlsx**
```
A列          | B列
张三          | 90
李四          | 85
王五          | 95
```

### 示例2: 合并产品信息

**文件1: products.xlsx**
```
A列
苹果
香蕉
橙子
西瓜
```

**文件2: prices.xlsx**
```
      C列
      5.5
      3.0
      4.0
      2.5
```

**命令:**
```bash
python scripts/merge_excel.py products.xlsx prices.xlsx 1 3
```

**输出: merged_excel.xlsx**
```
A列          | B列
苹果          | 5.5
香蕉          | 3.0
橙子          | 4.0
西瓜          | 2.5
```

### 示例3: 处理不同行数

**文件1: names.xlsx (3行)**
```
A列
张三
李四
王五
```

**文件2: phones.xlsx (5行)**
```
      B列
      13800138000
      13800138001
      13800138002
      13800138003
      13800138004
```

**命令:**
```bash
python scripts/merge_excel.py names.xlsx phones.xlsx
```

**输出: merged_excel.xlsx**
```
A列          | B列
张三          | 13800138000
李四          | 13800138001
王五          | 13800138002
              | 13800138003
              | 13800138004
```

## 输出格式

合并后的Excel文件包含:

1. **表头行**: 自动生成表头
   - A列: "文件1_第X列"
   - B列: "文件2_第Y列"

2. **数据行**: 从第2行开始
   - 列宽: 自动设置为30字符
   - 表头样式: 蓝色背景、白色粗体字、居中对齐

3. **输出位置**: 桌面 (Desktop)

## 技术要求

### 必需组件

- Python 3.6 或更高版本
- openpyxl库

### 安装依赖

```bash
# 使用pip安装openpyxl
pip install openpyxl

# 或使用国内镜像加速
pip install openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple
```

### 验证安装

```bash
# 检查Python版本
python --version

# 检查openpyxl安装
python -c "import openpyxl; print(openpyxl.__version__)"
```

## 常见问题

### Q1: 提示找不到openpyxl模块

**A:** 需要先安装openpyxl库:
```bash
pip install openpyxl
```

### Q2: 提示文件不存在

**A:** 请确保:
- 文件路径正确(使用完整路径或相对路径)
- 文件存在且可访问
- 文件是.xlsx格式(不是.xls)

### Q3: 合并后数据为空

**A:** 请检查:
- 指定的列号是否正确(从1开始)
- Excel文件中该列是否有数据
- 是否跳过了表头行

### Q4: 输出文件在哪里?

**A:** 默认输出到桌面:
- Windows: `C:\Users\用户名\Desktop\merged_excel.xlsx`
- macOS: `/Users/用户名/Desktop/merged_excel.xlsx`

### Q5: 可以合并多个Excel文件吗?

**A:** 当前版本只支持合并两个Excel文件。如果需要合并多个,可以分步进行:
```bash
# 先合并文件1和文件2
python scripts/merge_excel.py file1.xlsx file2.xlsx 1 1 temp1.xlsx

# 再将结果与文件3合并
python scripts/merge_excel.py temp1.xlsx file3.xlsx 1 1 final.xlsx
```

## 高级用法

### 自定义输出位置

修改脚本中的 `get_desktop_path()` 函数,或直接指定完整输出路径:

```python
result = create_merged_excel(
    file1_path="data1.xlsx",
    file2_path="data2.xlsx",
    output_path="C:/MyFolder",  # 自定义输出目录
    output_filename="custom_name.xlsx"
)
```

### 在Python代码中调用

```python
from scripts.merge_excel import create_merged_excel

# 合并Excel
result = create_merged_excel(
    file1_path="data1.xlsx",
    file2_path="data2.xlsx",
    file1_column=1,
    file2_column=2,
    output_filename="result.xlsx"
)

if result:
    print(f"合并成功: {result}")
```

## 错误处理

插件包含完善的错误处理:

- ✅ 文件不存在: 清晰的错误提示
- ✅ 文件格式错误: 自动检测并提示
- ✅ 列号超出范围: 提示正确的列号范围
- ✅ 权限问题: 提示需要访问权限
- ✅ 依赖缺失: 提示安装openpyxl

## 版本历史

### v1.0.0 (2026-03-12)
- ✅ 初始版本发布
- ✅ 支持合并两个Excel文件的指定列
- ✅ 自动美化输出格式
- ✅ 智能处理不同行数的数据

## 贡献指南

欢迎贡献代码、报告问题或提出建议!

1. Fork本仓库
2. 创建特性分支 (`git checkout -b feature/AmazingFeature`)
3. 提交更改 (`git commit -m 'Add some AmazingFeature'`)
4. 推送到分支 (`git push origin feature/AmazingFeature`)
5. 提交Pull Request

## 许可证

MIT License - 详见 LICENSE 文件

## 作者

- 作者: liujialin1517
- GitHub: https://github.com/liujialin1517-source

## 致谢

感谢以下项目:
- [openpyxl](https://openpyxl.readthedocs.io/) - Excel文件处理库
- [WorkBuddy](https://codebuddy.ai/) - AI开发助手

## 联系方式

如有问题或建议,请:
- 提交GitHub Issue
- 发送邮件至: [你的邮箱]

---

**享受使用Excel表格合并插件!** 🎉
