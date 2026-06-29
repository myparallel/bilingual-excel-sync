# 双语Excel表格同步技能 | Bilingual Excel Sync Tool

[中文](#中文) | [English](#english)

---

## 中文

将中文Excel表格自动转换为中英双语版本，并同步生成纯英文SHEET2工作表。

### 功能特点

- **双语转换**：将中文内容转换为"中文在上，英文在下"的双语格式
- **同步更新**：SHEET1双语表更新时，自动同步SHEET2纯英文表
- **智能翻译**：根据上下文自动翻译专业术语和常用表达
- **格式保持**：保持原有单元格样式、合并单元格、公式等

### 使用方法

```bash
# 分析文件
python bilingual_excel_sync.py analyze <input_file> [output_file]

# 执行转换
python bilingual_excel_sync.py convert <input_file> [output_file]

# 验证结果
python bilingual_excel_sync.py verify <bilingual_file> [output_file]
```

### 翻译字典

技能内置了常用翻译字典，包括：
- **通用字段**：项目名称、状态、日期格式等
- **外贸术语**：发货清单、装箱单、发票、合同等
- **传感器术语**：各种传感器类型的专业翻译

### 输出文件

- **文件名**：`原文件名_中英双语版.xlsx`
- **工作表**：
  - `Sheet1 双语版 Bilingual`：中英双语内容
  - `Sheet2 English Version`：纯英文内容

### 示例

**输入：**

| 项目名称 | 状态 | 数量 |
|---------|------|------|
| 传感器 | 已付 | 10 |

**输出（SHEET1双语版）：**

| 项目名称 | 状态 | 数量 |
|---------|------|------|
| 传感器<br>Sensor | 已付<br>Paid | 10 |

**输出（SHEET2纯英文版）：**

| Project Name | Status | Quantity |
|--------------|--------|----------|
| Sensor | Paid | 10 |

### 安装依赖

```bash
pip install openpyxl
```

---

## English

Automatically convert Chinese Excel spreadsheets to bilingual Chinese-English versions, with synchronized English-only SHEET2.

### Features

- **Bilingual Conversion**: Convert Chinese content to bilingual format with Chinese on top and English below
- **Synchronized Update**: Automatically sync SHEET2 English version when SHEET1 bilingual sheet is updated
- **Smart Translation**: Automatically translate professional terms and common expressions based on context
- **Format Preservation**: Maintain original cell styles, merged cells, formulas, etc.

### Usage

```bash
# Analyze file
python bilingual_excel_sync.py analyze <input_file> [output_file]

# Execute conversion
python bilingual_excel_sync.py convert <input_file> [output_file]

# Verify results
python bilingual_excel_sync.py verify <bilingual_file> [output_file]
```

### Translation Dictionary

Built-in translation dictionaries include:
- **Common Fields**: Project name, status, date formats, etc.
- **Trade Terms**: Shipping list, packing list, invoice, contract, etc.
- **Sensor Terms**: Professional translations for various sensor types

### Output Files

- **Filename**: `原文件名_中英双语版.xlsx`
- **Worksheets**:
  - `Sheet1 双语版 Bilingual`: Chinese-English bilingual content
  - `Sheet2 English Version`: English-only content

### Example

**Input:**

| 项目名称 | 状态 | 数量 |
|---------|------|------|
| 传感器 | 已付 | 10 |

**Output (SHEET1 Bilingual):**

| 项目名称 | 状态 | 数量 |
|---------|------|------|
| 传感器<br>Sensor | 已付<br>Paid | 10 |

**Output (SHEET2 English Only):**

| Project Name | Status | Quantity |
|--------------|--------|----------|
| Sensor | Paid | 10 |

### Install Dependencies

```bash
pip install openpyxl
```

---

## 许可证 | License

MIT License
