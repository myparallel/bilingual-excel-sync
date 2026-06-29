# 双语Excel表格同步技能使用指南

## 目录

1. [快速开始](#快速开始)
2. [基本用法](#基本用法)
3. [高级功能](#高级功能)
4. [常见问题](#常见问题)
5. [最佳实践](#最佳实践)

## 快速开始

### 安装

```bash
pip install openpyxl
```

### 基本使用

1. 准备中文Excel文件
2. 告诉AI您的需求
3. 等待AI处理
4. 查看结果

### 示例

```python
from scripts.bilingual_excel_sync import BilingualExcelSync

# 创建同步器
syncer = BilingualExcelSync()

# 加载翻译
translations = {
    '项目名称': 'Project Name',
    '状态': 'Status',
}
syncer.load_translations(translations)

# 执行转换
syncer.convert('中文表格.xlsx', '双语表格.xlsx')
```

## 基本用法

### 1. 转换单个文件

**AI对话方式：**
```
请将"发货清单.xlsx"制作成双语表格
```

**Python代码方式：**
```python
syncer.convert('发货清单.xlsx')
```

**命令行方式：**
```bash
python scripts/bilingual_excel_sync.py convert 发货清单.xlsx
```

### 2. 自定义翻译

**AI对话方式：**
```
请将以下中文翻译成英文：
- 土壤传感器 → Soil Sensor
- 振动传感器 → Vibration Sensor
```

**Python代码方式：**
```python
translations = {
    '土壤传感器': 'Soil Sensor',
    '振动传感器': 'Vibration Sensor',
}
syncer.load_translations(translations)
```

### 3. 同步更新

**AI对话方式：**
```
我更新了SHEET1中的"项目名称"字段，请同步更新SHEET2
```

**Python代码方式：**
```python
syncer.sync_english_sheet()
```

### 4. 验证翻译

**AI对话方式：**
```
请验证双语表格的翻译完整性
```

**Python代码方式：**
```python
syncer.verify('双语表格.xlsx', 'verification.txt')
```

## 高级功能

### 1. 批量处理

```python
files = ['文件1.xlsx', '文件2.xlsx', '文件3.xlsx']
for file in files:
    syncer.convert(file)
```

### 2. 自定义工作表名称

```python
syncer.add_sheet_translation('Sheet1', '双语版 Bilingual')
syncer.add_sheet_translation('Sheet2', 'English Version')
```

### 3. 分析文件内容

```python
syncer.analyze('输入文件.xlsx', 'analysis.txt')
```

### 4. 调整布局

```python
syncer.adjust_layout()
```

## 常见问题

### Q: 为什么翻译不准确？

A: 可能的原因：
1. 内置翻译字典不包含您的专业术语
2. 需要提供自定义翻译字典

**解决方案：**
```
请使用以下翻译字典：
- 您的术语1 → English 1
- 您的术语2 → English 2
```

### Q: 如何处理特殊格式？

A: 技能会自动处理：
1. 合并单元格
2. 单元格样式
3. 公式
4. 注释

### Q: 支持哪些Excel版本？

A: 支持：
- Excel 2007+ (.xlsx)
- Excel 97-2003 (.xls)
- 启用宏的工作簿 (.xlsm)

### Q: 如何提高翻译质量？

A: 建议：
1. 提供完整的翻译字典
2. 使用准确的术语
3. 保持一致性

## 最佳实践

### 1. 准备阶段

- 备份原始文件
- 整理中文内容
- 准备翻译字典

### 2. 转换阶段

- 使用清晰的命名
- 验证翻译结果
- 检查格式保持

### 3. 后期处理

- 同步更新SHEET2
- 验证数据完整性
- 保存最终版本

### 4. 维护阶段

- 更新翻译字典
- 记录修改历史
- 定期备份

## 示例代码

### 完整示例

```python
from scripts.bilingual_excel_sync import BilingualExcelSync

# 1. 创建同步器
syncer = BilingualExcelSync()

# 2. 加载翻译字典
translations = {
    # 列标题
    '项目名称': 'Project Name',
    '状态': 'Status',
    '数量': 'Quantity',
    
    # 状态值
    '已付': 'Paid',
    '待付': 'Pending',
    
    # 专业术语
    '土壤传感器': 'Soil Sensor',
    '振动传感器': 'Vibration Sensor',
}
syncer.load_translations(translations)

# 3. 执行转换
input_file = '中文表格.xlsx'
output_file = '双语表格.xlsx'
syncer.convert(input_file, output_file)

# 4. 验证结果
syncer.verify(output_file, 'verification.txt')
```

### 命令行示例

```bash
# 分析文件
python scripts/bilingual_excel_sync.py analyze 中文表格.xlsx analysis.txt

# 执行转换
python scripts/bilingual_excel_sync.py convert 中文表格.xlsx 双语表格.xlsx

# 验证结果
python scripts/bilingual_excel_sync.py verify 双语表格.xlsx verification.txt
```

### AI对话示例

```
用户：请将"发货清单.xlsx"制作成双语表格
AI：好的，我将为您处理这个文件...
```

## 性能优化

### 1. 使用缓存

```python
# 启用缓存
syncer.enable_cache()
```

### 2. 批量处理

```python
# 批量处理多个文件
syncer.batch_convert(['文件1.xlsx', '文件2.xlsx'])
```

### 3. 并行处理

```python
# 使用并行处理
syncer.parallel_convert(['文件1.xlsx', '文件2.xlsx'])
```

## 故障排除

### 问题1：翻译不准确

**症状：** SHEET2中还有中文字符

**原因：** 翻译字典不完整

**解决：** 提供完整的翻译字典

### 问题2：格式丢失

**症状：** 单元格样式改变

**原因：** 样式复制失败

**解决：** 检查原始文件格式

### 问题3：文件损坏

**症状：** 无法打开输出文件

**原因：** 文件保存错误

**解决：** 检查文件权限和磁盘空间

## 技术支持

如有问题，请：
1. 查看本文档
2. 运行测试脚本
3. 检查错误日志
4. 联系技术支持

## 更新日志

### v1.0.0
- 初始版本
- 支持双语转换
- 支持同步生成SHEET2
- 内置常用翻译字典