---
name: bilingual-excel-sync
description: 将中文Excel表格转换为中英双语版本，并同步生成纯英文SHEET2。当用户提到"双语表格"、"中英双语"、"翻译Excel"、"Excel翻译"、"表格翻译"、"双语版Excel"、"英文翻译"、"中英文对照"、"同步英文表"、"纯英文表"时触发。适用于外贸、商务、技术文档等需要中英文对照的场景。
---

# 双语Excel表格同步技能

将中文Excel表格自动转换为中英双语版本，并同步生成纯英文SHEET2工作表。

## 核心功能

1. **双语转换**：将中文内容转换为"中文在上，英文在下"的双语格式
2. **同步更新**：SHEET1双语表更新时，自动同步SHEET2纯英文表
3. **智能翻译**：根据上下文自动翻译专业术语和常用表达
4. **格式保持**：保持原有单元格样式、合并单元格、公式等

## 工作流程

### 第一步：分析文件内容

使用openpyxl加载Excel文件，提取所有唯一字符串值：

```python
import openpyxl

def analyze_excel(file_path):
    wb = openpyxl.load_workbook(file_path)
    unique_values = set()
    
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        for row in sheet.iter_rows(values_only=True):
            for cell in row:
                if cell and isinstance(cell, str):
                    unique_values.add(cell)
    
    return unique_values
```

### 第二步：构建翻译字典

根据分析结果，构建精确匹配的翻译字典。翻译字典应包含：

1. **列标题**：如"项目名称" → "Project Name"
2. **状态值**：如"已付" → "Paid", "待付" → "Pending"
3. **单位**：如"台" → "Unit/Set", "套" → "Set"
4. **专业术语**：根据行业领域翻译专业词汇
5. **常用表达**：如"标准盒装" → "Standard Box"

### 第三步：执行双语转换

```python
from openpyxl.styles import Alignment

def convert_to_bilingual(wb, translations):
    for sheet_name in wb.sheetnames:
        sheet = wb[sheet_name]
        
        # 更新工作表名称
        if sheet_name == 'Sheet1':
            sheet.title = 'Sheet1 双语版 Bilingual'
        
        # 处理单元格
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if cell.value in translations:
                        # 添加英文翻译到中文下方
                        cell.value = cell.value + '\n' + translations[cell.value]
                        cell.alignment = Alignment(wrap_text=True)
    
    return wb
```

### 第四步：同步SHEET2纯英文表

```python
import re
from copy import copy

def contains_chinese(text):
    """检查文本是否包含中文字符"""
    chinese_pattern = re.compile(r'[\u4e00-\u9fff]')
    return bool(chinese_pattern.search(text))

def extract_english(text):
    """从双语文本中提取英文部分"""
    if text is None:
        return None
    if isinstance(text, (int, float)):
        return text
    text_str = str(text)
    if '\n' in text_str:
        lines = text_str.split('\n')
        # 找到第一个不包含中文字符的行
        for i, line in enumerate(lines):
            if not contains_chinese(line):
                # 返回从这一行开始的所有行
                return '\n'.join(lines[i:])
    return text_str

def sync_english_sheet(wb):
    """同步生成纯英文SHEET2"""
    sheet1 = wb['Sheet1 双语版 Bilingual']
    
    # 删除旧的SHEET2（如果存在）
    if 'Sheet2 English Version' in wb.sheetnames:
        del wb['Sheet2 English Version']
    
    # 创建新的SHEET2
    sheet2 = wb.create_sheet('Sheet2 English Version')
    
    # 复制SHEET1内容到SHEET2，提取英文部分
    for row_idx, row in enumerate(sheet1.iter_rows(), 1):
        for col_idx, cell in enumerate(row, 1):
            new_cell = sheet2.cell(row=row_idx, column=col_idx)
            
            if cell.value is not None:
                english_value = extract_english(cell.value)
                new_cell.value = english_value
            
            # 复制样式
            if cell.font:
                new_cell.font = copy(cell.font)
            if cell.alignment:
                new_cell.alignment = copy(cell.alignment)
            if cell.border:
                new_cell.border = copy(cell.border)
            if cell.fill:
                new_cell.fill = copy(cell.fill)
    
    # 复制列宽
    for col_letter, col_dim in sheet1.column_dimensions.items():
        sheet2.column_dimensions[col_letter].width = col_dim.width
    
    # 复制行高
    for row_num, row_dim in sheet1.row_dimensions.items():
        sheet2.row_dimensions[row_num].height = row_dim.height
    
    # 设置页面为A4横向
    sheet2.page_setup.paperSize = sheet2.PAPERSIZE_A4
    sheet2.page_setup.orientation = 'landscape'
    
    return wb
```

### 第五步：验证结果

```python
def verify_translation(wb):
    """验证翻译结果"""
    sheet1 = wb['Sheet1 双语版 Bilingual']
    sheet2 = wb['Sheet2 English Version']
    
    # 检查SHEET1是否包含双语内容
    bilingual_count = 0
    for row in sheet1.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str) and '\n' in cell.value:
                bilingual_count += 1
    
    # 检查SHEET2是否只包含英文内容
    english_only = True
    for row in sheet2.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                if contains_chinese(cell.value):
                    english_only = False
                    break
    
    return {
        'bilingual_count': bilingual_count,
        'english_only': english_only
    }
```

## 使用示例

### 完整转换流程

```python
import openpyxl
from bilingual_excel_sync import BilingualExcelSync

# 创建同步器实例
syncer = BilingualExcelSync()

# 加载翻译字典
translations = {
    '项目名称': 'Project Name',
    '支付时间': 'Payment Date',
    '金额（USD)': 'Amount (USD)',
    '状态': 'Status',
    '已付': 'Paid',
    '待付': 'Pending',
    '小计：': 'Subtotal:',
    '总计：': 'Total:',
}

# 执行转换
input_file = '原始表格.xlsx'
output_file = '双语版表格.xlsx'

syncer.convert(input_file, output_file, translations)
print(f'转换完成！输出文件: {output_file}')
```

### 同步更新流程

当SHEET1内容更新后，同步更新SHEET2：

```python
# 更新SHEET1中的翻译
syncer.update_translations(new_translations)

# 同步更新SHEET2
syncer.sync_english_sheet()

# 保存文件
syncer.save()
```

## 翻译字典模板

### 通用字段

```python
COMMON_TRANSLATIONS = {
    # 列标题
    'NO.': 'NO.',
    '项目名称': 'Project Name',
    '项目描述': 'Project Description',
    '金额（USD)': 'Amount (USD)',
    '状态': 'Status',
    '传感器类型': 'Sensor Type',
    '传感器型号': 'Sensor Model',
    '基础型号': 'Base Model',
    
    # 状态值
    '已付': 'Paid',
    '待付': 'Pending',
    '未付': 'Unpaid',
    '已签约': 'Signed',
    '待签约': 'To Be Signed',
    '谈判中': 'In Negotiation',
    '待谈判': 'To Be Negotiated',
    '未签约': 'Not Signed',
    
    # 汇总行
    '小计：': 'Subtotal:',
    '总计：': 'Total:',
    
    # 日期格式
    '（预估）': '(Estimated)',
    '预计': 'Expected',
}
```

### 外贸常用术语

```python
TRADE_TRANSLATIONS = {
    # 贸易术语
    '发货清单': 'Shipping List',
    '装箱单': 'Packing List',
    '发票': 'Invoice',
    '合同': 'Contract',
    '订单': 'Order',
    '样品': 'Sample',
    
    # 产品相关
    '产品名称': 'Product Name',
    '产品描述': 'Product Description',
    '规格': 'Specification',
    '数量': 'Quantity',
    '单位': 'Unit',
    '单价': 'Unit Price',
    '总价': 'Total Price',
    
    # 物流相关
    '发货': 'Shipment',
    '收货': 'Receipt',
    '运输': 'Transport',
    '仓储': 'Warehousing',
    '报关': 'Customs Declaration',
    
    # 付款相关
    '付款方式': 'Payment Method',
    '预付款': 'Advance Payment',
    '尾款': 'Balance Payment',
    '信用证': 'Letter of Credit',
    '电汇': 'Wire Transfer',
}
```

## 注意事项

### 编码问题

- **不要**依赖控制台输出验证结果（Windows PowerShell中文乱码）
- **必须**将分析和验证结果写入文件（UTF-8编码）
- 使用 `with open('xxx.txt', 'w', encoding='utf-8')` 写入

### 翻译匹配

- 使用**精确匹配**，不要模糊匹配
- 翻译字典的key必须与文件中的字符串**完全一致**
- 注意空格、标点、括号等细节

### 格式保持

- 保留原有单元格样式（字体、颜色、边框）
- 保留合并单元格信息
- 保留公式（如SUM、=D7+D14等）

### 同步逻辑

- SHEET1双语表更新后，必须同步更新SHEET2纯英文表
- SHEET2只包含英文内容，不包含任何中文字符
- 两个工作表的列宽、行高、页面设置保持一致

## 输出文件

- **文件名**：`原文件名_中英双语版.xlsx`
- **位置**：与原文件相同目录
- **工作表**：
  - `Sheet1 双语版 Bilingual`：中英双语内容
  - `Sheet2 English Version`：纯英文内容

## 快速使用

```bash
# 分析文件
python bilingual_excel_sync.py analyze <input_file> [output_file]

# 执行转换
python bilingual_excel_sync.py convert <input_file> [output_file]

# 验证结果
python bilingual_excel_sync.py verify <bilingual_file> [output_file]
```