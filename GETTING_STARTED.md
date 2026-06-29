# 双语Excel表格同步技能入门指南

## 欢迎使用双语Excel表格同步技能！

这个技能可以帮您将中文Excel表格自动转换为中英双语版本，并同步生成纯英文的SHEET2工作表。

## 5分钟快速入门

### 第1步：安装依赖

```bash
pip install openpyxl
```

### 第2步：准备中文Excel表格

创建一个包含中文内容的Excel文件，例如：

**示例表格.xlsx**

| 项目名称 | 数量 | 状态 |
|---------|------|------|
| 传感器 | 10 | 已付 |
| 变送器 | 5 | 待付 |

### 第3步：告诉AI您的需求

在AI对话中说：

```
请将"示例表格.xlsx"制作成双语表格，并同步生成纯英文的SHEET2
```

### 第4步：等待AI处理

AI会自动：
1. 分析Excel文件中的所有中文内容
2. 使用内置翻译字典进行翻译
3. 生成双语版本（中文在上，英文在下）
4. 同步生成纯英文的SHEET2工作表

### 第5步：查看结果

输出文件：`示例表格_中英双语版.xlsx`

- **SHEET1**：中英双语版（中文在上，英文在下）
- **SHEET2**：纯英文版（只包含英文内容）

## 常用场景

### 场景1：外贸订单翻译

```
请将"订单列表.xlsx"翻译成中英双语版本，用于发给国外客户
```

### 场景2：产品清单翻译

```
我需要将"产品清单.xlsx"制作成双语表格，方便国际团队查看
```

### 场景3：发货清单翻译

```
请将"发货清单.xlsx"翻译成英文，并保持原有格式
```

## 自定义翻译

如果内置翻译不满足需求，可以告诉AI：

```
请将以下中文翻译成英文：
- 土壤传感器 → Soil Sensor
- 振动传感器 → Vibration Sensor
- 太阳能储能电池 → Solar Energy Storage Battery
```

## 同步更新

当您更新SHEET1中的中文内容后，告诉AI：

```
我更新了SHEET1中的"项目名称"字段，请同步更新SHEET2
```

## 验证翻译

如果需要验证翻译的完整性：

```
请验证双语表格的翻译完整性
```

## 使用Python API

```python
from scripts.bilingual_excel_sync import BilingualExcelSync

# 创建同步器实例
syncer = BilingualExcelSync()

# 加载翻译字典
translations = {
    '项目名称': 'Project Name',
    '状态': 'Status',
    '已付': 'Paid',
    '待付': 'Pending',
}

syncer.load_translations(translations)

# 执行转换
syncer.convert('示例表格.xlsx', '双语表格.xlsx')
```

## 使用命令行

```bash
# 分析文件
python scripts/bilingual_excel_sync.py analyze 示例表格.xlsx analysis.txt

# 执行转换
python scripts/bilingual_excel_sync.py convert 示例表格.xlsx 双语表格.xlsx

# 验证结果
python scripts/bilingual_excel_sync.py verify 双语表格.xlsx verification.txt
```

## 运行演示

```bash
python demo.py
```

## 运行测试

```bash
python test_skill.py
python test_cases.py
```

## 更多信息

- 详细使用说明：[USAGE.md](USAGE.md)
- 快速开始：[QUICK_START.md](QUICK_START.md)
- 使用示例：[example_使用示例.md](example_使用示例.md)
- API参考：[API_REFERENCE.md](API_REFERENCE.md)
- 故障排除：[TROUBLESHOOTING.md](TROUBLESHOOTING.md)

## 获取帮助

如有问题，请：
1. 查看[USAGE.md](USAGE.md)详细说明
2. 运行`python demo.py`查看演示
3. 运行`python test_skill.py`验证功能

## 贡献

欢迎贡献代码和改进建议！请查看[CONTRIBUTING.md](CONTRIBUTING.md)了解详情。

## 许可证

本项目采用MIT许可证。请查看[LICENSE](LICENSE)了解详情。