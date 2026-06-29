# 双语Excel表格同步技能快速参考

## 快速开始

### 安装
```bash
pip install openpyxl
```

### 基本使用
```
请将"文件名.xlsx"制作成双语表格
```

## 常用命令

### AI对话命令

| 命令 | 说明 |
|------|------|
| `请将"文件.xlsx"制作成双语表格` | 转换单个文件 |
| `请将以下中文翻译成英文：...` | 自定义翻译 |
| `我更新了SHEET1，请同步更新SHEET2` | 同步更新 |
| `请验证双语表格的翻译完整性` | 验证翻译 |

### Python API命令

```python
from scripts.bilingual_excel_sync import BilingualExcelSync

syncer = BilingualExcelSync()
syncer.load_translations(translations)
syncer.convert('输入.xlsx', '输出.xlsx')
syncer.verify('输出.xlsx', 'verification.txt')
```

### 命令行命令

```bash
# 分析文件
python scripts/bilingual_excel_sync.py analyze 输入.xlsx analysis.txt

# 执行转换
python scripts/bilingual_excel_sync.py convert 输入.xlsx 输出.xlsx

# 验证结果
python scripts/bilingual_excel_sync.py verify 输出.xlsx verification.txt
```

## 翻译字典示例

### 通用字段
```python
translations = {
    '项目名称': 'Project Name',
    '状态': 'Status',
    '数量': 'Quantity',
    '单位': 'Unit',
    '备注': 'Remarks',
}
```

### 状态值
```python
translations = {
    '已付': 'Paid',
    '待付': 'Pending',
    '已确认': 'Confirmed',
    '待确认': 'Pending Confirmation',
    '已发货': 'Shipped',
    '有库存': 'In Stock',
    '缺货': 'Out of Stock',
}
```

### 单位
```python
translations = {
    '台': 'Unit/Set',
    '套': 'Set',
    '个': 'Piece',
    '包': 'Pack',
    '盒': 'Box',
}
```

## 常见问题

### Q: 为什么翻译不准确？
A: 提供自定义翻译字典。

### Q: 如何保持格式？
A: 技能会自动保持格式。

### Q: 支持哪些格式？
A: .xlsx, .xlsm, .xls

### Q: 如何同步更新？
A: 告诉AI"我更新了SHEET1，请同步更新SHEET2"。

## 文件结构

```
bilingual-excel-sync/
├── SKILL.md                 # 技能说明
├── README.md               # 项目说明
├── USAGE.md               # 使用说明
├── QUICK_START.md         # 快速开始
├── QUICK_REFERENCE.md     # 快速参考
├── demo.py               # 演示脚本
├── test_skill.py         # 测试脚本
├── examples.py           # 示例代码
└── scripts/
    └── bilingual_excel_sync.py  # 主程序
```

## 常用文件

| 文件 | 说明 |
|------|------|
| `demo.py` | 运行演示 |
| `test_skill.py` | 运行测试 |
| `examples.py` | 查看示例 |
| `USAGE.md` | 详细说明 |
| `API_REFERENCE.md` | API参考 |

## 快速测试

```bash
# 运行演示
python demo.py

# 运行测试
python test_skill.py

# 运行示例
python examples.py
```

## 获取帮助

1. 查看[USAGE.md](USAGE.md)
2. 运行`python demo.py`
3. 运行`python test_skill.py`

## 更新日志

### v1.0.0
- 初始版本
- 支持双语转换
- 支持同步生成SHEET2
- 内置常用翻译字典