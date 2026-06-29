# 双语Excel表格同步技能教程

## 目录

1. [基础教程](#基础教程)
2. [进阶教程](#进阶教程)
3. [高级教程](#高级教程)
4. [实战案例](#实战案例)

## 基础教程

### 教程1：第一次使用

#### 目标
将一个简单的中文Excel表格转换为双语版本。

#### 步骤

**第1步：创建测试文件**

创建一个Excel文件，内容如下：

| 项目名称 | 数量 | 状态 |
|---------|------|------|
| 传感器 | 10 | 已付 |
| 变送器 | 5 | 待付 |

保存为`测试表格.xlsx`。

**第2步：告诉AI您的需求**

```
请将"测试表格.xlsx"制作成双语表格
```

**第3步：等待AI处理**

AI会自动处理并生成`测试表格_中英双语版.xlsx`。

**第4步：查看结果**

打开文件，您会看到：
- SHEET1：中英双语版
- SHEET2：纯英文版

#### 结果

**SHEET1（双语版）**

| 项目名称 | 数量 | 状态 |
|---------|------|------|
| 项目名称<br>Project Name | 数量<br>Quantity | 状态<br>Status |
| 传感器<br>Sensor | 10 | 已付<br>Paid |
| 变送器<br>Transmitter | 5 | 待付<br>Pending |

**SHEET2（纯英文版）**

| Project Name | Quantity | Status |
|--------------|----------|--------|
| Project Name | Quantity | Status |
| Sensor | 10 | Paid |
| Transmitter | 5 | Pending |

### 教程2：使用自定义翻译

#### 目标
使用自定义翻译字典转换表格。

#### 步骤

**第1步：准备翻译字典**

```python
translations = {
    '项目名称': 'Project Name',
    '状态': 'Status',
    '已付': 'Paid',
    '待付': 'Pending',
    '传感器': 'Sensor',
    '变送器': 'Transmitter',
}
```

**第2步：告诉AI使用自定义翻译**

```
请使用以下翻译字典转换"测试表格.xlsx"：
- 项目名称 → Project Name
- 状态 → Status
- 已付 → Paid
- 待付 → Pending
- 传感器 → Sensor
- 变送器 → Transmitter
```

**第3步：等待AI处理**

AI会使用您提供的翻译字典进行转换。

**第4步：查看结果**

结果与教程1相同。

### 教程3：同步更新

#### 目标
更新SHEET1后，同步更新SHEET2。

#### 步骤

**第1步：转换表格**

```
请将"测试表格.xlsx"制作成双语表格
```

**第2步：更新SHEET1**

手动更新SHEET1中的内容，例如添加新行。

**第3步：同步更新SHEET2**

```
我更新了SHEET1，请同步更新SHEET2
```

**第4步：查看结果**

SHEET2会自动更新以匹配SHEET1的内容。

## 进阶教程

### 教程4：处理复杂表格

#### 目标
处理包含多列和复杂内容的表格。

#### 步骤

**第1步：创建复杂表格**

创建一个Excel文件，内容如下：

| NO. | 项目名称 | 规格 | 数量 | 单位 | 备注 |
|-----|---------|------|------|------|------|
| 1 | 土壤传感器 | DG-214 | 20 | 台 | 标准盒装 |
| 2 | 振动传感器 | PG-T920 | 2 | 台 | 特殊定制 |
| 3 | 太阳能储能电池 | - | 3 | 台 | 上海发货 |

**第2步：告诉AI您的需求**

```
请将"复杂表格.xlsx"制作成双语表格，需要翻译所有中文内容
```

**第3步：等待AI处理**

AI会翻译所有中文内容，包括列标题、项目名称、单位、备注等。

**第4步：查看结果**

所有中文内容都会被翻译成英文。

### 教程5：批量处理

#### 目标
批量处理多个Excel文件。

#### 步骤

**第1步：准备多个文件**

准备多个Excel文件：
- 文件1.xlsx
- 文件2.xlsx
- 文件3.xlsx

**第2步：告诉AI批量处理**

```
我有多个Excel文件需要翻译成双语版本：
1. 文件1.xlsx
2. 文件2.xlsx
3. 文件3.xlsx
请帮我批量处理
```

**第3步：等待AI处理**

AI会逐个处理每个文件。

**第4步：查看结果**

每个文件都会生成对应的双语版本。

### 教程6：使用专业术语

#### 目标
使用专业术语进行翻译。

#### 步骤

**第1步：准备专业术语字典**

```python
translations = {
    '压力监测变送器': 'Pressure Monitoring Transmitter',
    '液位监测变送器': 'Level Monitoring Transmitter',
    '温度监测变送器': 'Temperature Monitoring Transmitter',
}
```

**第2步：告诉AI使用专业术语**

```
请使用以下专业术语字典翻译"传感器表格.xlsx"：
- 压力监测变送器 → Pressure Monitoring Transmitter
- 液位监测变送器 → Level Monitoring Transmitter
- 温度监测变送器 → Temperature Monitoring Transmitter
```

**第3步：等待AI处理**

AI会使用专业术语进行翻译。

**第4步：查看结果**

所有专业术语都会被正确翻译。

## 高级教程

### 教程7：使用Python API

#### 目标
使用Python API进行编程控制。

#### 步骤

**第1步：导入模块**

```python
from scripts.bilingual_excel_sync import BilingualExcelSync
```

**第2步：创建同步器实例**

```python
syncer = BilingualExcelSync()
```

**第3步：加载翻译字典**

```python
translations = {
    '项目名称': 'Project Name',
    '状态': 'Status',
}
syncer.load_translations(translations)
```

**第4步：执行转换**

```python
syncer.convert('输入文件.xlsx', '输出文件.xlsx')
```

**第5步：验证结果**

```python
syncer.verify('输出文件.xlsx', 'verification.txt')
```

### 教程8：使用命令行

#### 目标
使用命令行工具进行批量处理。

#### 步骤

**第1步：分析文件**

```bash
python scripts/bilingual_excel_sync.py analyze 输入文件.xlsx analysis.txt
```

**第2步：执行转换**

```bash
python scripts/bilingual_excel_sync.py convert 输入文件.xlsx 输出文件.xlsx
```

**第3步：验证结果**

```bash
python scripts/bilingual_excel_sync.py verify 输出文件.xlsx verification.txt
```

### 教程9：自定义工作表名称

#### 目标
自定义工作表名称。

#### 步骤

**第1步：告诉AI自定义名称**

```
请将"测试表格.xlsx"制作成双语表格，工作表名称使用：
- SHEET1：中英双语版
- SHEET2：English Version
```

**第2步：等待AI处理**

AI会使用自定义的工作表名称。

**第3步：查看结果**

工作表名称会按照您的要求显示。

## 实战案例

### 案例1：外贸订单翻译

#### 背景
您是一家外贸公司的员工，需要将中文订单表格翻译成英文发给国外客户。

#### 步骤

**第1步：准备订单表格**

创建一个包含订单信息的Excel文件。

**第2步：告诉AI您的需求**

```
请将"订单表格.xlsx"翻译成中英双语版本，用于发给国外客户
```

**第3步：等待AI处理**

AI会翻译所有内容。

**第4步：发送给客户**

将生成的双语表格发送给客户。

### 案例2：产品清单翻译

#### 背景
您是一家制造公司的员工，需要将产品清单翻译成英文供国际团队使用。

#### 步骤

**第1步：准备产品清单**

创建一个包含产品信息的Excel文件。

**第2步：告诉AI您的需求**

```
请将"产品清单.xlsx"翻译成中英双语版本，方便国际团队查看
```

**第3步：等待AI处理**

AI会翻译所有内容。

**第4步：分发给团队**

将生成的双语表格分发给国际团队。

### 案例3：发货清单翻译

#### 背景
您是一家物流公司的员工，需要将发货清单翻译成英文用于报关。

#### 步骤

**第1步：准备发货清单**

创建一个包含发货信息的Excel文件。

**第2步：告诉AI您的需求**

```
请将"发货清单.xlsx"翻译成英文，并保持原有格式，用于报关
```

**第3步：等待AI处理**

AI会翻译所有内容并保持格式。

**第4步：用于报关**

将生成的英文表格用于报关。

## 总结

通过本教程，您已经学会了：

1. 基础教程：如何第一次使用技能
2. 进阶教程：如何处理复杂表格和批量处理
3. 高级教程：如何使用Python API和命令行
4. 实战案例：如何在实际工作中应用技能

## 更多资源

- 详细使用说明：[USAGE.md](USAGE.md)
- 快速开始：[QUICK_START.md](QUICK_START.md)
- 使用示例：[example_使用示例.md](example_使用示例.md)
- API参考：[API_REFERENCE.md](API_REFERENCE.md)
- 故障排除：[TROUBLESHOOTING.md](TROUBLESHOOTING.md)

## 获取帮助

如有问题，请：
1. 查看本文档
2. 运行`python demo.py`查看演示
3. 运行`python test_skill.py`验证功能