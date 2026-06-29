# 双语Excel表格同步技能API参考

## BilingualExcelSync类

### 导入

```python
from scripts.bilingual_excel_sync import BilingualExcelSync
```

### 构造函数

```python
syncer = BilingualExcelSync()
```

创建一个新的BilingualExcelSync实例。

### 方法

#### load_workbook(file_path)

加载Excel文件。

**参数：**
- `file_path` (str): Excel文件路径

**返回：**
- `Workbook`: openpyxl工作簿对象

**异常：**
- `FileNotFoundError`: 文件不存在时抛出

**示例：**
```python
syncer.load_workbook('发货清单.xlsx')
```

#### load_translations(translations)

批量加载翻译字典。

**参数：**
- `translations` (dict): 翻译字典，格式为 `{中文: 英文}`

**示例：**
```python
translations = {
    '项目名称': 'Project Name',
    '状态': 'Status',
}
syncer.load_translations(translations)
```

#### add_sheet_translation(chinese, english)

添加工作表名称翻译。

**参数：**
- `chinese` (str): 中文工作表名称
- `english` (str): 英文工作表名称

**示例：**
```python
syncer.add_sheet_translation('Sheet1', 'Sheet1 双语版 Bilingual')
```

#### contains_chinese(text)

检查文本是否包含中文字符。

**参数：**
- `text` (str): 要检查的文本

**返回：**
- `bool`: 如果包含中文字符返回True，否则返回False

**示例：**
```python
result = syncer.contains_chinese('项目名称')  # True
result = syncer.contains_chinese('Project Name')  # False
```

#### extract_english(text)

从双语文本中提取英文部分。

**参数：**
- `text` (str): 双语文本，格式为 `中文\n英文`

**返回：**
- `str`: 英文部分

**示例：**
```python
text = '项目名称\nProject Name'
result = syncer.extract_english(text)  # 'Project Name'
```

#### analyze(input_file, output_file)

分析文件内容，输出所有唯一字符串到文件。

**参数：**
- `input_file` (str): 输入Excel文件路径
- `output_file` (str): 输出文本文件路径

**返回：**
- `str`: 输出文件路径

**示例：**
```python
syncer.analyze('发货清单.xlsx', 'analysis.txt')
```

#### convert_to_bilingual()

将SHEET1转换为双语格式。

**返回：**
- `Workbook`: 更新后的工作簿对象

**示例：**
```python
syncer.load_workbook('发货清单.xlsx')
syncer.load_translations(translations)
syncer.convert_to_bilingual()
```

#### sync_english_sheet()

同步生成纯英文SHEET2。

**返回：**
- `Workbook`: 更新后的工作簿对象

**示例：**
```python
syncer.sync_english_sheet()
```

#### adjust_layout()

调整布局为A4横向版面。

**返回：**
- `Workbook`: 更新后的工作簿对象

**示例：**
```python
syncer.adjust_layout()
```

#### convert(input_file, output_file=None)

执行完整转换流程。

**参数：**
- `input_file` (str): 输入Excel文件路径
- `output_file` (str, optional): 输出Excel文件路径，默认为 `原文件名_中英双语版.xlsx`

**返回：**
- `str`: 输出文件路径

**示例：**
```python
syncer.convert('发货清单.xlsx', '双语版_发货清单.xlsx')
```

#### verify(bilingual_file, output_file)

验证转换结果，检查未翻译的字段。

**参数：**
- `bilingual_file` (str): 双语Excel文件路径
- `output_file` (str): 验证结果输出文件路径

**返回：**
- `str`: 输出文件路径

**示例：**
```python
syncer.verify('双语版_发货清单.xlsx', 'verification.txt')
```

#### save(output_path)

保存双语版Excel文件。

**参数：**
- `output_path` (str): 输出文件路径

**返回：**
- `str`: 输出文件路径

**异常：**
- `ValueError`: 工作簿未加载时抛出

**示例：**
```python
syncer.save('双语版_发货清单.xlsx')
```

## 内置翻译字典

### COMMON_TRANSLATIONS

通用字段翻译字典。

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
    '传感类型': 'Sensor Type',
    
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

### TRADE_TRANSLATIONS

外贸常用术语翻译字典。

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

### SENSOR_TRANSLATIONS

传感器/工业术语翻译字典。

```python
SENSOR_TRANSLATIONS = {
    '压力监测变送器': 'Pressure Monitoring Transmitter',
    '液位监测变送器': 'Level Monitoring Transmitter',
    '微压差变送器': 'Micro Differential Pressure Transmitter',
    '温度监测变送器': 'Temperature Monitoring Transmitter',
    '温湿度监测变送器': 'Temperature & Humidity Monitoring Transmitter',
    'PH监测变送器': 'PH Monitoring Transmitter',
    '电导率监测变送器': 'Conductivity Monitoring Transmitter',
    '溶解氧变送器': 'Dissolved Oxygen Transmitter',
    '浊度变送器': 'Turbidity Transmitter',
    '双通道压力监测变送器': 'Dual-Channel Pressure Monitoring Transmitter',
    '水浸监测变送器': 'Water Immersion Monitoring Transmitter',
    '土壤温湿度电导率变送器': 'Soil Temp/Humidity/Conductivity Transmitter',
    '防爆压力变送器': 'Explosion-Proof Pressure Transmitter',
    '防爆液位变送器': 'Explosion-Proof Level Transmitter',
    '防爆深潜投入式液位变送器': 'Explosion-Proof Submersible Level Transmitter',
    '氨气传感变送器': 'Ammonia Gas Sensor Transmitter',
    '雷达物位变送器': 'Radar Level Transmitter',
    '液位尺传感变送器': 'Level Ruler Sensor Transmitter',
    '二氧化碳传感变送器': 'CO2 Sensor Transmitter',
    '甲醛传感变送器': 'Formaldehyde Sensor Transmitter',
    '多普勒流量计': 'Doppler Flow Meter',
    '甲烷传感变送器': 'Methane Sensor Transmitter',
    '一氧化碳传感变送器': 'CO Sensor Transmitter',
    '氧气传感变送器': 'Oxygen Sensor Transmitter',
    '超声波液位传感变送器': 'Ultrasonic Level Sensor Transmitter',
    '防爆温度变送器': 'Explosion-Proof Temperature Transmitter',
    '盐度传感变送器': 'Salinity Sensor Transmitter',
    '疏水阀传感器': 'Steam Trap Sensor',
    '多功能检测仪': 'Multi-function Detector',
    '三通道压力及甲烷泄露监测传感变送器': '3-Channel Pressure & Methane Leak Monitor Transmitter',
    '电导率变送器': 'Conductivity Transmitter',
    '硫化氢传感变送器': 'H2S Sensor Transmitter',
}
```

## 命令行接口

### 分析文件

```bash
python scripts/bilingual_excel_sync.py analyze <input_file> [output_file]
```

**参数：**
- `input_file`: 输入Excel文件路径
- `output_file`: 输出文本文件路径（默认为 `analysis.txt`）

**示例：**
```bash
python scripts/bilingual_excel_sync.py analyze 发货清单.xlsx analysis.txt
```

### 执行转换

```bash
python scripts/bilingual_excel_sync.py convert <input_file> [output_file]
```

**参数：**
- `input_file`: 输入Excel文件路径
- `output_file`: 输出Excel文件路径（默认为 `原文件名_中英双语版.xlsx`）

**示例：**
```bash
python scripts/bilingual_excel_sync.py convert 发货清单.xlsx 双语版_发货清单.xlsx
```

### 验证结果

```bash
python scripts/bilingual_excel_sync.py verify <bilingual_file> [output_file]
```

**参数：**
- `bilingual_file`: 双语Excel文件路径
- `output_file`: 验证结果输出文件路径（默认为 `verification.txt`）

**示例：**
```bash
python scripts/bilingual_excel_sync.py verify 双语版_发货清单.xlsx verification.txt
```

## 错误处理

### FileNotFoundError

文件不存在时抛出。

```python
try:
    syncer.load_workbook('不存在的文件.xlsx')
except FileNotFoundError as e:
    print(f"文件不存在: {e}")
```

### ValueError

工作簿未加载时抛出。

```python
try:
    syncer.save('输出.xlsx')
except ValueError as e:
    print(f"错误: {e}")
```

## 依赖库

- `openpyxl`: 用于读写Excel文件
- `re`: 用于正则表达式匹配
- `os`: 用于文件路径操作
- `sys`: 用于系统操作
- `copy`: 用于对象复制

## 版本信息

- 版本: 1.0.0
- 作者: OpenCode
- 许可证: MIT