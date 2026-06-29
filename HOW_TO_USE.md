# 如何使用双语Excel表格同步技能

## 第一步：了解技能功能

这个技能可以帮您：
1. 将中文Excel表格转换为中英双语版本
2. 自动生成纯英文的SHEET2工作表
3. 保持原有格式和样式
4. 支持自定义翻译字典

## 第二步：准备中文Excel表格

创建一个包含中文内容的Excel文件，例如：

**示例表格.xlsx**

| 项目名称 | 数量 | 状态 |
|---------|------|------|
| 传感器 | 10 | 已付 |
| 变送器 | 5 | 待付 |

## 第三步：告诉AI您的需求

在AI对话中，使用以下任一方式：

### 方式1：直接说明需求

```
请将"示例表格.xlsx"制作成双语表格，并同步生成纯英文的SHEET2
```

### 方式2：详细描述需求

```
我需要将这个Excel表格翻译成中英双语版本，同时要有一个单独的纯英文工作表。
要求：
1. 中文在上，英文在下
2. 保持原有格式
3. 自动生成纯英文的SHEET2
```

### 方式3：使用关键词

```
双语表格、中英双语、翻译Excel、Excel翻译、表格翻译
```

## 第四步：等待AI处理

AI会自动：
1. 分析Excel文件中的所有中文内容
2. 使用内置翻译字典进行翻译
3. 生成双语版本（中文在上，英文在下）
4. 同步生成纯英文的SHEET2工作表
5. 保持原有格式和样式

## 第五步：查看结果

输出文件：`原文件名_中英双语版.xlsx`

### SHEET1（双语版）

| 项目名称 | 数量 | 状态 |
|---------|------|------|
| 项目名称<br>Project Name | 数量<br>Quantity | 状态<br>Status |
| 传感器<br>Sensor | 10 | 已付<br>Paid |
| 变送器<br>Transmitter | 5 | 待付<br>Pending |

### SHEET2（纯英文版）

| Project Name | Quantity | Status |
|--------------|----------|--------|
| Project Name | Quantity | Status |
| Sensor | 10 | Paid |
| Transmitter | 5 | Pending |

## 第六步：自定义翻译（可选）

如果内置翻译不满足需求，可以告诉AI：

```
请将以下中文翻译成英文：
- 土壤传感器 → Soil Sensor
- 振动传感器 → Vibration Sensor
- 太阳能储能电池 → Solar Energy Storage Battery
```

## 第七步：同步更新（可选）

当您更新SHEET1中的中文内容后，告诉AI：

```
我更新了SHEET1中的"项目名称"字段，请同步更新SHEET2
```

## 第八步：验证翻译（可选）

如果需要验证翻译的完整性：

```
请验证双语表格的翻译完整性
```

## 常见问题

### Q: 为什么SHEET2中还有中文字符？

A: 可能的原因：
1. 翻译字典中没有对应的翻译
2. 翻译字典的key与文件中的字符串不完全一致
3. 字符串包含特殊字符或空格

**解决方案**：告诉AI需要翻译的具体内容。

### Q: 如何添加新的翻译？

A: 两种方式：
1. 在对话中直接告诉AI新的翻译
2. 提供自定义翻译字典

### Q: 如何保持原有格式？

A: 技能会自动：
1. 保留原有单元格样式（字体、颜色、边框）
2. 保留合并单元格信息
3. 保留公式（如SUM、=D7+D14等）

### Q: 支持哪些Excel格式？

A: 支持：
- .xlsx（Excel 2007+）
- .xlsm（启用宏的工作簿）
- .xls（Excel 97-2003）

## 使用技巧

### 技巧1：批量处理

如果您有多个文件需要处理，可以告诉AI：

```
我有多个Excel文件需要翻译成双语版本，请帮我处理：
1. 发货清单.xlsx
2. 产品列表.xlsx
3. 订单信息.xlsx
```

### 技巧2：自定义翻译字典

如果您有特定的翻译需求，可以提供翻译字典：

```
请使用以下翻译字典：
- 项目名称 → Project Name
- 状态 → Status
- 已付 → Paid
- 待付 → Pending
```

### 技巧3：同步更新

当您更新SHEET1后，可以告诉AI：

```
我更新了SHEET1中的"状态"字段，请同步更新SHEET2
```

### 技巧4：验证翻译

如果需要验证翻译的完整性：

```
请验证双语表格的翻译完整性，并告诉我哪些字段没有翻译
```

## 高级用法

### 1. 使用Python API

```python
from scripts.bilingual_excel_sync import BilingualExcelSync

# 创建同步器实例
syncer = BilingualExcelSync()

# 加载翻译字典
translations = {
    '项目名称': 'Project Name',
    '状态': 'Status',
}

syncer.load_translations(translations)

# 执行转换
syncer.convert('示例表格.xlsx', '双语版_示例表格.xlsx')
```

### 2. 使用命令行

```bash
# 分析文件
python scripts/bilingual_excel_sync.py analyze 示例表格.xlsx analysis.txt

# 执行转换
python scripts/bilingual_excel_sync.py convert 示例表格.xlsx 双语版_示例表格.xlsx

# 验证结果
python scripts/bilingual_excel_sync.py verify 双语版_示例表格.xlsx verification.txt
```

### 3. 运行演示

```bash
python demo.py
```

### 4. 运行测试

```bash
python test_skill.py
python test_cases.py
```

## 获取帮助

如有问题，请：
1. 查看[USAGE.md](USAGE.md)详细说明
2. 查看[API_REFERENCE.md](API_REFERENCE.md)API参考
3. 运行`python demo.py`查看演示
4. 运行`python test_skill.py`验证功能

## 更多信息

- 详细使用说明：[USAGE.md](USAGE.md)
- 安装指南：[INSTALL.md](INSTALL.md)
- 故障排除：[TROUBLESHOOTING.md](TROUBLESHOOTING.md)
- 示例代码：[examples.py](examples.py)
- API参考：[API_REFERENCE.md](API_REFERENCE.md)