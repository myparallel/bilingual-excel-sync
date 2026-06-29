# 双语Excel表格同步技能速查表

## 一行命令

### 转换文件
```
请将"文件.xlsx"制作成双语表格
```

### 自定义翻译
```
请将以下中文翻译成英文：术语1 → English 1, 术语2 → English 2
```

### 同步更新
```
我更新了SHEET1，请同步更新SHEET2
```

### 验证翻译
```
请验证双语表格的翻译完整性
```

## Python代码

### 基本使用
```python
from scripts.bilingual_excel_sync import BilingualExcelSync

syncer = BilingualExcelSync()
syncer.load_translations({'中文': 'English'})
syncer.convert('输入.xlsx', '输出.xlsx')
```

### 完整示例
```python
from scripts.bilingual_excel_sync import BilingualExcelSync

# 1. 创建同步器
syncer = BilingualExcelSync()

# 2. 加载翻译
translations = {
    '项目名称': 'Project Name',
    '状态': 'Status',
    '已付': 'Paid',
    '待付': 'Pending',
}
syncer.load_translations(translations)

# 3. 执行转换
syncer.convert('中文表格.xlsx', '双语表格.xlsx')

# 4. 验证结果
syncer.verify('双语表格.xlsx', 'verification.txt')
```

## 命令行

### 分析文件
```bash
python scripts/bilingual_excel_sync.py analyze 输入.xlsx analysis.txt
```

### 执行转换
```bash
python scripts/bilingual_excel_sync.py convert 输入.xlsx 输出.xlsx
```

### 验证结果
```bash
python scripts/bilingual_excel_sync.py verify 输出.xlsx verification.txt
```

## 翻译字典

### 通用字段
```python
{
    '项目名称': 'Project Name',
    '状态': 'Status',
    '数量': 'Quantity',
    '单位': 'Unit',
    '备注': 'Remarks',
}
```

### 状态值
```python
{
    '已付': 'Paid',
    '待付': 'Pending',
    '已确认': 'Confirmed',
    '待确认': 'Pending Confirmation',
    '已发货': 'Shipped',
}
```

### 单位
```python
{
    '台': 'Unit/Set',
    '套': 'Set',
    '个': 'Piece',
    '包': 'Pack',
    '盒': 'Box',
}
```

## 常见问题

### 翻译不准确
**解决：** 提供自定义翻译字典

### 格式丢失
**解决：** 技能会自动保持格式

### SHEET2有中文
**解决：** 告诉AI需要翻译的具体内容

### 文件打不开
**解决：** 检查文件权限和磁盘空间

## 快速测试

### 运行演示
```bash
python demo.py
```

### 运行测试
```bash
python test_skill.py
```

### 运行示例
```bash
python examples.py
```

## 文件说明

| 文件 | 说明 |
|------|------|
| `demo.py` | 演示脚本 |
| `test_skill.py` | 测试脚本 |
| `examples.py` | 示例代码 |
| `USAGE.md` | 使用说明 |
| `API_REFERENCE.md` | API参考 |

## 获取帮助

1. 查看[USAGE.md](USAGE.md)
2. 运行`python demo.py`
3. 运行`python test_skill.py`

## 记忆口诀

1. **准备**：创建中文Excel表格
2. **告诉**：告诉AI您的需求
3. **等待**：等待AI处理
4. **查看**：查看生成的双语表格
5. **同步**：更新SHEET1后同步SHEET2