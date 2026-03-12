# 桌面文件计数插件

## 插件介绍

这是一个用于统计桌面文件数量的 WorkBuddy 插件。它可以快速扫描桌面目录，统计所有文件的数量、类型分布和总大小。

## 功能特性

- ✅ 统计桌面文件总数（包括子目录）
- ✅ 统计目录数量
- ✅ 显示文件类型分布（按扩展名）
- ✅ 显示总文件大小
- ✅ 识别最大的文件
- ✅ 友好的可视化输出

## 使用方法

### 通过 WorkBuddy 使用

直接在 WorkBuddy 中输入：
- "统计桌面有多少个文件"
- "查看桌面文件数量"
- "计算桌面文件总数"

WorkBuddy 会自动识别并调用此技能。

### 直接运行脚本

```bash
python C:\Users\wuwuc\.workbuddy\plugins\marketplaces\local-user-marketplace\plugins\desktop-file-counter\scripts\count_desktop_files.py
```

或指定自定义路径：

```bash
python C:\Users\wuwuc\.workbuddy\plugins\marketplaces\local-user-marketplace\plugins\desktop-file-counter\scripts\count_desktop_files.py "D:\我的文档"
```

## 技术要求

- Python 3.x
- 无需额外依赖（使用 Python 标准库）
