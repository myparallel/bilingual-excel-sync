# 双语Excel表格同步技能使用示例

## 示例1：基本使用

### 输入文件

**原始表格.xlsx**

| NO. | 项目名称 | 数量 | 单位 | 状态 |
|-----|---------|------|------|------|
| 1 | 土壤传感器 | 20 | 台 | 已确认 |
| 2 | 振动传感器 | 2 | 台 | 待确认 |
| 3 | 太阳能储能电池 | 3 | 台 | 已发货 |

### 输出文件

**双语版表格.xlsx**

#### SHEET1（双语版）

| NO. | 项目名称 | 数量 | 单位 | 状态 |
|-----|---------|------|------|------|
| NO. | 项目名称<br>Project Name | 数量<br>Quantity | 单位<br>Unit | 状态<br>Status |
| 1 | 土壤传感器<br>Soil Sensor | 20 | 台<br>Unit/Set | 已确认<br>Confirmed |
| 2 | 振动传感器<br>Vibration Sensor | 2 | 台<br>Unit/Set | 待确认<br>Pending Confirmation |
| 3 | 太阳能储能电池<br>Solar Energy Storage Battery | 3 | 台<br>Unit/Set | 已发货<br>Shipped |

#### SHEET2（纯英文版）

| NO. | Project Name | Quantity | Unit | Status |
|-----|--------------|----------|------|--------|
| NO. | Project Name | Quantity | Unit | Status |
| 1 | Soil Sensor | 20 | Unit/Set | Confirmed |
| 2 | Vibration Sensor | 2 | Unit/Set | Pending Confirmation |
| 3 | Solar Energy Storage Battery | 3 | Unit/Set | Shipped |

## 示例2：发货清单

### 输入文件

**发货清单.xlsx**

| ITEM名称 | 规格 | 数量 | 单位 | 备注 |
|---------|------|------|------|------|
| 土壤传感器（DGL定制） | DG-214-4G-P2 | 20 | 台 | 标准盒装：主机+探头+背装卡架 |
| 振动传感器（DLG定制） | PG-T920-4G-P2 | 2 | 台 | 标准盒装：主机+探头 |
| 太阳能储能电池 | - | 3 | 台 | 上海发货 |

### 输出文件

**发货清单_中英双语版.xlsx**

#### SHEET1（双语版）

| ITEM名称 | 规格 | 数量 | 单位 | 备注 |
|---------|------|------|------|------|
| ITEM名称<br>Item Name | 规格<br>Specification | 数量<br>Quantity | 单位<br>Unit | 备注<br>Remarks |
| 土壤传感器（DGL定制）<br>Soil Sensor (DGL Customized) | DG-214-4G-P2 | 20 | 台<br>Unit/Set | 标准盒装：主机+探头+背装卡架<br>Standard Box: Main Unit + Probe + Back-Mount Card Rack |
| 振动传感器（DLG定制）<br>Vibration Sensor (DLG Customized) | PG-T920-4G-P2 | 2 | 台<br>Unit/Set | 标准盒装：主机+探头<br>Standard Box: Main Unit + Probe |
| 太阳能储能电池<br>Solar Energy Storage Battery | - | 3 | 台<br>Unit/Set | 上海发货<br>Shanghai Shipment |

#### SHEET2（纯英文版）

| Item Name | Specification | Quantity | Unit | Remarks |
|-----------|---------------|----------|------|---------|
| Item Name | Specification | Quantity | Unit | Remarks |
| Soil Sensor (DGL Customized) | DG-214-4G-P2 | 20 | Unit/Set | Standard Box: Main Unit + Probe + Back-Mount Card Rack |
| Vibration Sensor (DLG Customized) | PG-T920-4G-P2 | 2 | Unit/Set | Standard Box: Main Unit + Probe |
| Solar Energy Storage Battery | - | 3 | Unit/Set | Shanghai Shipment |

## 示例3：产品列表

### 输入文件

**产品列表.xlsx**

| 产品编号 | 产品名称 | 产品描述 | 单价（USD） | 库存数量 | 状态 |
|---------|---------|---------|------------|---------|------|
| P001 | 压力传感器 | 用于工业压力监测 | 150.00 | 100 | 有库存 |
| P002 | 温度传感器 | 用于温度监测 | 120.00 | 50 | 有库存 |
| P003 | 液位传感器 | 用于液位监测 | 180.00 | 0 | 缺货 |

### 输出文件

**产品列表_中英双语版.xlsx**

#### SHEET1（双语版）

| 产品编号 | 产品名称 | 产品描述 | 单价（USD） | 库存数量 | 状态 |
|---------|---------|---------|------------|---------|------|
| 产品编号<br>Product Code | 产品名称<br>Product Name | 产品描述<br>Product Description | 单价（USD）<br>Unit Price (USD) | 库存数量<br>Stock Quantity | 状态<br>Status |
| P001 | 压力传感器<br>Pressure Sensor | 用于工业压力监测<br>For industrial pressure monitoring | 150.00 | 100 | 有库存<br>In Stock |
| P002 | 温度传感器<br>Temperature Sensor | 用于温度监测<br>For temperature monitoring | 120.00 | 50 | 有库存<br>In Stock |
| P003 | 液位传感器<br>Level Sensor | 用于液位监测<br>For level monitoring | 180.00 | 0 | 缺货<br>Out of Stock |

#### SHEET2（纯英文版）

| Product Code | Product Name | Product Description | Unit Price (USD) | Stock Quantity | Status |
|--------------|--------------|---------------------|------------------|----------------|--------|
| Product Code | Product Name | Product Description | Unit Price (USD) | Stock Quantity | Status |
| P001 | Pressure Sensor | For industrial pressure monitoring | 150.00 | 100 | In Stock |
| P002 | Temperature Sensor | For temperature monitoring | 120.00 | 50 | In Stock |
| P003 | Level Sensor | For level monitoring | 180.00 | 0 | Out of Stock |

## 示例4：订单信息

### 输入文件

**订单信息.xlsx**

| 订单编号 | 客户名称 | 产品名称 | 数量 | 单价（USD） | 总价（USD） | 订单状态 | 交货日期 |
|---------|---------|---------|------|------------|------------|---------|---------|
| ORD-2024-001 | ABC公司 | 压力传感器 | 10 | 150.00 | 1500.00 | 已确认 | 2024-03-15 |
| ORD-2024-002 | XYZ公司 | 温度传感器 | 5 | 120.00 | 600.00 | 待确认 | 2024-03-20 |
| ORD-2024-003 | DEF公司 | 液位传感器 | 8 | 180.00 | 1440.00 | 已发货 | 2024-03-10 |

### 输出文件

**订单信息_中英双语版.xlsx**

#### SHEET1（双语版）

| 订单编号 | 客户名称 | 产品名称 | 数量 | 单价（USD） | 总价（USD） | 订单状态 | 交货日期 |
|---------|---------|---------|------|------------|------------|---------|---------|
| 订单编号<br>Order Number | 客户名称<br>Customer Name | 产品名称<br>Product Name | 数量<br>Quantity | 单价（USD）<br>Unit Price (USD) | 总价（USD）<br>Total Price (USD) | 订单状态<br>Order Status | 交货日期<br>Delivery Date |
| ORD-2024-001 | ABC公司<br>ABC Company | 压力传感器<br>Pressure Sensor | 10 | 150.00 | 1500.00 | 已确认<br>Confirmed | 2024-03-15 |
| ORD-2024-002 | XYZ公司<br>XYZ Company | 温度传感器<br>Temperature Sensor | 5 | 120.00 | 600.00 | 待确认<br>Pending Confirmation | 2024-03-20 |
| ORD-2024-003 | DEF公司<br>DEF Company | 液位传感器<br>Level Sensor | 8 | 180.00 | 1440.00 | 已发货<br>Shipped | 2024-03-10 |

#### SHEET2（纯英文版）

| Order Number | Customer Name | Product Name | Quantity | Unit Price (USD) | Total Price (USD) | Order Status | Delivery Date |
|--------------|---------------|--------------|----------|------------------|-------------------|--------------|---------------|
| Order Number | Customer Name | Product Name | Quantity | Unit Price (USD) | Total Price (USD) | Order Status | Delivery Date |
| ORD-2024-001 | ABC Company | Pressure Sensor | 10 | 150.00 | 1500.00 | Confirmed | 2024-03-15 |
| ORD-2024-002 | XYZ Company | Temperature Sensor | 5 | 120.00 | 600.00 | Pending Confirmation | 2024-03-20 |
| ORD-2024-003 | DEF Company | Level Sensor | 8 | 180.00 | 1440.00 | Shipped | 2024-03-10 |

## 示例5：自定义翻译

### 输入文件

**自定义表格.xlsx**

| 自定义字段1 | 自定义字段2 | 自定义字段3 |
|-----------|-----------|-----------|
| 值1 | 值2 | 值3 |

### 自定义翻译字典

```python
translations = {
    '自定义字段1': 'Custom Field 1',
    '自定义字段2': 'Custom Field 2',
    '自定义字段3': 'Custom Field 3',
    '值1': 'Value 1',
    '值2': 'Value 2',
    '值3': 'Value 3',
}
```

### 输出文件

**自定义表格_中英双语版.xlsx**

#### SHEET1（双语版）

| 自定义字段1 | 自定义字段2 | 自定义字段3 |
|-----------|-----------|-----------|
| 自定义字段1<br>Custom Field 1 | 自定义字段2<br>Custom Field 2 | 自定义字段3<br>Custom Field 3 |
| 值1<br>Value 1 | 值2<br>Value 2 | 值3<br>Value 3 |

#### SHEET2（纯英文版）

| Custom Field 1 | Custom Field 2 | Custom Field 3 |
|----------------|----------------|----------------|
| Custom Field 1 | Custom Field 2 | Custom Field 3 |
| Value 1 | Value 2 | Value 3 |

## 示例6：同步更新

### 初始状态

**SHEET1（双语版）**

| 项目名称 | 状态 |
|---------|------|
| 项目名称<br>Project Name | 状态<br>Status |
| 传感器<br>Sensor | 已付<br>Paid |
| 变送器<br>Transmitter | 待付<br>Pending |

**SHEET2（纯英文版）**

| Project Name | Status |
|--------------|--------|
| Project Name | Status |
| Sensor | Paid |
| Transmitter | Pending |

### 更新SHEET1

更新第2行第2列的值：

| 项目名称 | 状态 |
|---------|------|
| 项目名称<br>Project Name | 状态<br>Status |
| 传感器<br>Sensor | 已付<br>Paid |
| 变送器<br>Transmitter | 待付<br>Pending |
| 新项目<br>New Project | 新状态<br>New Status |

### 同步更新SHEET2

| Project Name | Status |
|--------------|--------|
| Project Name | Status |
| Sensor | Paid |
| Transmitter | Pending |
| New Project | New Status |

## 示例7：验证翻译

### 输入文件

**验证表格.xlsx**

| 项目名称 | 状态 |
|---------|------|
| 传感器 | 已付 |
| 变送器 | 待付 |
| 未知项目 | 未知状态 |

### 验证结果

```
验证文件: 验证表格_中英双语版.xlsx

=== 工作表: Sheet1 双语版 Bilingual ===

所有行数据:
Row 1: ('项目名称\nProject Name', '状态\nStatus')
Row 2: ('传感器\nSensor', '已付\nPaid')
Row 3: ('变送器\nTransmitter', '待付\nPending')
Row 4: ('未知项目', '未知状态')

=== 统计 ===
总单元格数: 8
已翻译单元格数: 6
未翻译单元格数: 2

未翻译的字段:
  Row 4, Col 1: "未知项目"
  Row 4, Col 2: "未知状态"
```

## 更多示例

更多使用示例请查看：
- [examples.py](examples.py) - Python代码示例
- [USAGE.md](USAGE.md) - 详细使用说明
- [demo.py](demo.py) - 完整演示代码