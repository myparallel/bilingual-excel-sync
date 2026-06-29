# 双语Excel表格同步技能术语表

## 基本术语

### 双语表格 (Bilingual Table)
包含中文和英文两种语言的表格，通常中文在上，英文在下。

### 纯英文表 (English-Only Table)
只包含英文内容的表格，通常作为SHEET2。

### 翻译字典 (Translation Dictionary)
包含中文和英文对应关系的字典，用于自动翻译。

### 同步更新 (Sync Update)
当SHEET1内容更新后，自动更新SHEET2的内容。

## 技术术语

### openpyxl
Python库，用于读写Excel文件。

### 工作表 (Worksheet)
Excel文件中的一个页面，也称为Sheet。

### 单元格 (Cell)
表格中的一个格子，包含数据或公式。

### 合并单元格 (Merged Cells)
多个单元格合并成一个大的单元格。

### 自动换行 (Wrap Text)
单元格中的文本自动换行显示。

### 页面设置 (Page Setup)
设置打印页面的大小、方向等。

## 翻译术语

### 精确匹配 (Exact Match)
翻译字典的key必须与文件中的字符串完全一致。

### 模糊匹配 (Fuzzy Match)
翻译字典的key与文件中的字符串近似匹配。

### 专业术语 (Professional Terms)
特定领域的专业词汇，如传感器、变送器等。

### 通用术语 (Common Terms)
日常使用的通用词汇，如项目名称、状态等。

## 文件术语

### 输入文件 (Input File)
需要翻译的原始中文Excel文件。

### 输出文件 (Output File)
翻译后生成的双语Excel文件。

### 分析文件 (Analysis File)
包含文件内容分析结果的文本文件。

### 验证文件 (Verification File)
包含翻译验证结果的文本文件。

## 功能术语

### 双语转换 (Bilingual Conversion)
将中文内容转换为"中文在上，英文在下"的双语格式。

### 英文提取 (English Extraction)
从双语文本中提取英文部分。

### 格式保持 (Format Preservation)
保持原有单元格样式、合并单元格、公式等。

### 批量处理 (Batch Processing)
同时处理多个文件。

## 命令术语

### AI对话命令 (AI Chat Command)
在AI对话中使用的命令，如"请将文件.xlsx制作成双语表格"。

### Python API命令 (Python API Command)
在Python代码中使用的命令，如`syncer.convert()`。

### 命令行命令 (Command Line Command)
在命令行中使用的命令，如`python scripts/bilingual_excel_sync.py convert`。

## 状态术语

### 已付 (Paid)
付款状态，表示已经付款。

### 待付 (Pending)
付款状态，表示等待付款。

### 已确认 (Confirmed)
订单状态，表示已经确认。

### 待确认 (Pending Confirmation)
订单状态，表示等待确认。

### 已发货 (Shipped)
物流状态，表示已经发货。

### 有库存 (In Stock)
库存状态，表示有库存。

### 缺货 (Out of Stock)
库存状态，表示没有库存。

## 单位术语

### 台 (Unit/Set)
计量单位，通常用于设备。

### 套 (Set)
计量单位，通常用于套装。

### 个 (Piece)
计量单位，通常用于单个物品。

### 包 (Pack)
计量单位，通常用于包装。

### 盒 (Box)
计量单位，通常用于盒装。

## 文件格式术语

### .xlsx
Excel 2007+文件格式。

### .xlsm
启用宏的Excel工作簿格式。

### .xls
Excel 97-2003文件格式。

### .csv
逗号分隔值文件格式。

### .tsv
制表符分隔值文件格式。

## 错误术语

### FileNotFoundError
文件不存在错误。

### PermissionError
权限错误，文件被占用或没有写入权限。

### ValueError
值错误，参数不正确。

### UnicodeDecodeError
编码错误，通常出现在Windows PowerShell中。

### MemoryError
内存错误，内存不足。

## 优化术语

### 缓存 (Cache)
临时存储计算结果，提高性能。

### 批量处理 (Batch Processing)
同时处理多个项目，提高效率。

### 并行处理 (Parallel Processing)
同时执行多个任务，提高效率。

### 流式处理 (Streaming Processing)
逐个处理项目，减少内存使用。

## 测试术语

### 单元测试 (Unit Testing)
测试单个函数或方法。

### 集成测试 (Integration Testing)
测试多个组件的协作。

### 性能测试 (Performance Testing)
测试程序的性能。

### 验证测试 (Validation Testing)
测试程序的正确性。

## 部署术语

### 安装 (Installation)
将程序安装到系统中。

### 配置 (Configuration)
设置程序的参数。

### 部署 (Deployment)
将程序部署到生产环境。

### 维护 (Maintenance)
更新和修复程序。

## 文档术语

### README
项目说明文件。

### API参考 (API Reference)
应用程序编程接口参考文档。

### 用户指南 (User Guide)
用户使用指南。

### 开发者指南 (Developer Guide)
开发者指南。

### 故障排除 (Troubleshooting)
问题排查指南。

## 版本术语

### 版本号 (Version Number)
程序的版本标识。

### 更新日志 (Changelog)
记录程序的更新历史。

### 发布 (Release)
发布新版本。

### 补丁 (Patch)
修复问题的更新。

## 许可证术语

### MIT许可证 (MIT License)
一种开源许可证。

### 开源 (Open Source)
源代码公开的软件。

### 免费 (Free)
不需要付费。

### 贡献 (Contribute)
为项目做出贡献。

## 社区术语

### GitHub
代码托管平台。

### Issue
问题报告。

### Pull Request
代码合并请求。

### Fork
复制项目。

### Star
收藏项目。

## 工具术语

### pip
Python包管理器。

### conda
Python环境管理器。

### virtualenv
Python虚拟环境。

### pytest
Python测试框架。

### black
Python代码格式化工具。

## 概念术语

### 自动化 (Automation)
自动执行任务。

### 国际化 (Internationalization)
支持多种语言。

### 本地化 (Localization)
适应特定地区的需求。

### 标准化 (Standardization)
遵循统一的标准。

### 最佳实践 (Best Practice)
推荐的做法。