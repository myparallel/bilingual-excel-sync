# 双语Excel表格同步技能

将中文Excel表格自动转换为中英双语版本，并同步生成纯英文SHEET2工作表。

## 功能特点

- **双语转换**：将中文内容转换为"中文在上，英文在下"的双语格式
- **同步更新**：SHEET1双语表更新时，自动同步SHEET2纯英文表
- **智能翻译**：根据上下文自动翻译专业术语和常用表达
- **格式保持**：保持原有单元格样式、合并单元格、公式等

## 使用方法

### 1. 分析文件内容

```bash
python bilingual_excel_sync.py analyze <input_file> [output_file]
```

分析Excel文件中的所有唯一字符串值，输出到文本文件。

### 2. 执行转换

```bash
python bilingual_excel_sync.py convert <input_file> [output_file]
```

将中文Excel表格转换为中英双语版本，并同步生成纯英文SHEET2。

### 3. 验证结果

```bash
python bilingual_excel_sync.py verify <bilingual_file> [output_file]
```

验证转换结果，检查未翻译的字段。

## 翻译字典

技能内置了常用翻译字典，包括：

- **通用字段**：项目名称、状态、日期格式等
- **外贸术语**：发货清单、装箱单、发票、合同等
- **传感器术语**：各种传感器类型的专业翻译

## 输出文件

- **文件名**：`原文件名_中英双语版.xlsx`
- **位置**：与原文件相同目录
- **工作表**：
  - `Sheet1 双语版 Bilingual`：中英双语内容
  - `Sheet2 English Version`：纯英文内容

## 示例

### 输入文件

| 项目名称 | 状态 | 数量 |
|---------|------|------|
| 传感器 | 已付 | 10 |
| 变送器 | 待付 | 5 |

### 输出文件（SHEET1双语版）

| 项目名称 | 状态 | 数量 |
|---------|------|------|
| 传感器<br>Sensor | 已付<br>Paid | 10 |
| 变送器<br>Transmitter | 待付<br>Pending | 5 |

### 输出文件（SHEET2纯英文版）

| Project Name | Status | Quantity |
|--------------|--------|----------|
| Sensor | Paid | 10 |
| Transmitter | Pending | 5 |

## 注意事项

1. **编码问题**：Windows PowerShell中文显示乱码，建议将结果输出到文件
2. **翻译匹配**：使用精确匹配，翻译字典的key必须与文件中的字符串完全一致
3. **格式保持**：保留原有单元格样式、合并单元格、公式等
4. **同步逻辑**：SHEET1更新后必须同步更新SHEET2

## 依赖库

- openpyxl：用于读写Excel文件

## 安装依赖

```bash
pip install openpyxl
```

## 许可证

MIT License