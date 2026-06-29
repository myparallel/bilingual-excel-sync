# 双语Excel表格同步技能

[English](README.md) | [中文](README_CN.md)

将中文Excel表格自动转换为中英双语版本，并同步生成纯英文SHEET2工作表。

## 功能特点

- **双语转换**：将中文内容转换为"中文在上，英文在下"的双语格式
- **同步更新**：SHEET1双语表更新时，自动同步SHEET2纯英文表
- **智能翻译**：根据上下文自动翻译专业术语和常用表达
- **格式保持**：保持原有单元格样式、合并单元格、公式等

## 快速开始

### 安装

```bash
pip install openpyxl
```

### 使用

在AI对话中说：

```
请将"发货清单.xlsx"制作成双语表格，并同步生成纯英文的SHEET2
```

### 输出

生成`发货清单_中英双语版.xlsx`，包含：
- `Sheet1 双语版 Bilingual`：中英双语内容
- `Sheet2 English Version`：纯英文内容

## 使用方法

### AI对话使用

```
请将"文件名.xlsx"制作成双语表格
```

### Python API使用

```python
from scripts.bilingual_excel_sync import BilingualExcelSync

syncer = BilingualExcelSync()
syncer.load_translations(translations)
syncer.convert('输入.xlsx', '输出.xlsx')
```

### 命令行使用

```bash
python scripts/bilingual_excel_sync.py convert 输入.xlsx 输出.xlsx
```

## 内置翻译

### 通用字段
- 项目名称 → Project Name
- 状态 → Status
- 数量 → Quantity
- 单位 → Unit
- 备注 → Remarks

### 状态值
- 已付 → Paid
- 待付 → Pending
- 已确认 → Confirmed
- 待确认 → Pending Confirmation
- 已发货 → Shipped

### 单位
- 台 → Unit/Set
- 套 → Set
- 个 → Piece
- 包 → Pack
- 盒 → Box

## 自定义翻译

```
请将以下中文翻译成英文：
- 土壤传感器 → Soil Sensor
- 振动传感器 → Vibration Sensor
- 太阳能储能电池 → Solar Energy Storage Battery
```

## 同步更新

```
我更新了SHEET1中的"项目名称"字段，请同步更新SHEET2
```

## 验证翻译

```
请验证双语表格的翻译完整性
```

## 文档

- [快速开始](QUICK_START.md)
- [使用说明](USAGE.md)
- [API参考](API_REFERENCE.md)
- [常见问题](FAQ.md)
- [故障排除](TROUBLESHOOTING.md)

## 示例

### 输入文件

| 项目名称 | 数量 | 状态 |
|---------|------|------|
| 传感器 | 10 | 已付 |
| 变送器 | 5 | 待付 |

### 输出文件

#### SHEET1（双语版）

| 项目名称 | 数量 | 状态 |
|---------|------|------|
| 项目名称<br>Project Name | 数量<br>Quantity | 状态<br>Status |
| 传感器<br>Sensor | 10 | 已付<br>Paid |
| 变送器<br>Transmitter | 5 | 待付<br>Pending |

#### SHEET2（纯英文版）

| Project Name | Quantity | Status |
|--------------|----------|--------|
| Project Name | Quantity | Status |
| Sensor | 10 | Paid |
| Transmitter | 5 | Pending |

## 测试

```bash
# 运行演示
python demo.py

# 运行测试
python test_skill.py

# 运行示例
python examples.py
```

## 依赖

- Python 3.6+
- openpyxl 3.0+

## 许可证

MIT License

## 贡献

欢迎贡献代码！请查看[贡献指南](CONTRIBUTING.md)。

## 支持

如有问题，请：
1. 查看[常见问题](FAQ.md)
2. 运行`python demo.py`查看演示
3. 创建GitHub Issue

## 更新日志

查看[更新日志](CHANGELOG.md)了解最新更新。

## 相关链接

- [项目说明](README.md)
- [使用说明](USAGE.md)
- [API参考](API_REFERENCE.md)
- [示例代码](examples.py)
- [演示脚本](demo.py)

## 记忆口诀

1. **准备**：创建中文Excel表格
2. **告诉**：告诉AI您的需求
3. **等待**：等待AI处理
4. **查看**：查看生成的双语表格
5. **同步**：更新SHEET1后同步SHEET2