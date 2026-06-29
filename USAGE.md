# 双语Excel表格同步技能使用说明

## 快速开始

### 1. 直接使用技能

当您需要将中文Excel表格转换为双语版本时，只需告诉AI：

```
请将"发货清单.xlsx"制作成双语表格，并同步生成纯英文的SHEET2
```

或者：

```
我需要将这个Excel表格翻译成中英双语版本，同时要有一个单独的纯英文工作表
```

### 2. 技能会自动完成

技能会自动：
1. 分析Excel文件中的所有中文内容
2. 使用内置翻译字典进行翻译
3. 生成双语版本（中文在上，英文在下）
4. 同步生成纯英文的SHEET2工作表
5. 保持原有格式和样式

## 内置翻译字典

技能已经内置了常用翻译，包括：

### 通用字段
- 项目名称 → Project Name
- 状态 → Status
- 数量 → Quantity
- 单位 → Unit
- 备注 → Remarks

### 外贸术语
- 发货清单 → Shipping List
- 装箱单 → Packing List
- 发票 → Invoice
- 合同 → Contract

### 传感器术语
- 压力监测变送器 → Pressure Monitoring Transmitter
- 液位监测变送器 → Level Monitoring Transmitter
- 温度监测变送器 → Temperature Monitoring Transmitter

## 自定义翻译

如果内置翻译不满足需求，您可以：

### 1. 提供自定义翻译字典

```
请将以下中文翻译成英文：
- 土壤传感器 → Soil Sensor
- 振动传感器 → Vibration Sensor
- 太阳能储能电池 → Solar Energy Storage Battery
```

### 2. 使用技能的翻译字典功能

技能支持批量加载翻译字典：

```python
from bilingual_excel_sync import BilingualExcelSync

syncer = BilingualExcelSync()

# 加载自定义翻译
custom_translations = {
    '自定义字段1': 'Custom Field 1',
    '自定义字段2': 'Custom Field 2',
}

syncer.load_translations(custom_translations)
```

## 同步更新

### 自动同步

当您更新SHEET1中的中文内容后，技能会自动同步更新SHEET2的英文内容：

```
我更新了SHEET1中的"项目名称"字段，请同步更新SHEET2
```

### 手动同步

如果需要手动同步：

```python
from bilingual_excel_sync import BilingualExcelSync

syncer = BilingualExcelSync()
syncer.load_workbook('双语版表格.xlsx')

# 同步更新SHEET2
syncer.sync_english_sheet()

# 保存文件
syncer.save('双语版表格.xlsx')
```

## 验证结果

### 检查翻译完整性

```
请验证双语表格的翻译完整性
```

技能会：
1. 检查SHEET1是否包含双语内容
2. 检查SHEET2是否只包含英文内容
3. 统计已翻译和未翻译的单元格数量

### 输出验证报告

技能会生成验证报告，包括：
- 总单元格数
- 已翻译单元格数
- 未翻译单元格列表

## 文件结构

```
bilingual-excel-sync/
├── SKILL.md                 # 技能说明文件
├── README.md               # 项目说明文件
├── USAGE.md               # 使用说明文件
├── test_skill.py          # 测试脚本
├── evals/
│   └── evals.json         # 测试用例
└── scripts/
    └── bilingual_excel_sync.py  # 主程序
```

## 常见问题

### Q: 为什么SHEET2中还有中文字符？

A: 可能的原因：
1. 翻译字典中没有对应的翻译
2. 翻译字典的key与文件中的字符串不完全一致
3. 字符串包含特殊字符或空格

**解决方案**：检查翻译字典，确保key与文件中的字符串完全一致。

### Q: 如何添加新的翻译？

A: 两种方式：
1. 在对话中直接告诉AI新的翻译
2. 修改脚本中的翻译字典

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

## 最佳实践

### 1. 备份原始文件

在转换前，建议备份原始文件：
```
请先备份"原始表格.xlsx"，然后制作双语版本
```

### 2. 使用有意义的文件名

输出文件会自动命名为：`原文件名_中英双语版.xlsx`

### 3. 定期同步更新

当SHEET1内容更新后，及时同步更新SHEET2：
```
我更新了SHEET1，请同步更新SHEET2
```

### 4. 验证翻译结果

转换完成后，建议验证翻译结果：
```
请验证双语表格的翻译完整性
```

## 技术支持

如有问题或建议，请：
1. 检查本文档的常见问题部分
2. 查看SKILL.md中的详细说明
3. 运行test_skill.py测试脚本

## 更新日志

### v1.0.0
- 初始版本
- 支持双语转换
- 支持同步生成纯英文SHEET2
- 内置常用翻译字典