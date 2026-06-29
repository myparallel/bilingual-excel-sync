# 双语Excel表格同步技能故障排除指南

## 常见问题及解决方案

### 1. 安装问题

#### 问题：pip install openpyxl 失败

**错误信息**：
```
ERROR: Could not find a version that satisfies the requirement openpyxl
```

**解决方案**：
```bash
# 方法1：使用国内镜像
pip install openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple

# 方法2：使用conda
conda install openpyxl

# 方法3：使用--user参数
pip install --user openpyxl
```

#### 问题：Python版本不兼容

**错误信息**：
```
SyntaxError: invalid syntax
```

**解决方案**：
检查Python版本：
```bash
python --version
```

确保使用Python 3.6或更高版本。

### 2. 编码问题

#### 问题：Windows PowerShell中文显示乱码

**现象**：
```
�����嵥.xlsx
```

**解决方案**：
1. 将结果输出到文件：
```bash
python scripts/bilingual_excel_sync.py analyze 发货清单.xlsx analysis.txt
```

2. 使用文本编辑器打开analysis.txt查看结果。

3. 或者设置PowerShell编码：
```powershell
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
```

#### 问题：文件名包含中文字符

**错误信息**：
```
FileNotFoundError: [Errno 2] No such file or directory: '发货清单.xlsx'
```

**解决方案**：
1. 使用英文文件名：
```bash
rename 发货清单.xlsx shipping_list.xlsx
python scripts/bilingual_excel_sync.py convert shipping_list.xlsx
```

2. 或者在Python脚本中使用完整路径：
```python
import os
file_path = os.path.join('D:', '工作', '发货清单.xlsx')
```

### 3. 翻译问题

#### 问题：翻译不准确

**现象**：
SHEET2中仍然包含中文字符。

**解决方案**：
1. 检查翻译字典是否包含所有需要翻译的字段。
2. 确保翻译字典的key与文件中的字符串完全一致。
3. 注意空格、标点、括号等细节。

**示例**：
```python
# 错误的翻译字典
translations = {
    '项目名称': 'Project Name',  # 正确
    '项目名称 ': 'Project Name',  # 错误：多了空格
}

# 正确的翻译字典
translations = {
    '项目名称': 'Project Name',
}
```

#### 问题：SHEET2中还有中文字符

**现象**：
验证结果显示未翻译的字段。

**解决方案**：
1. 运行验证脚本查看哪些字段未翻译：
```bash
python scripts/bilingual_excel_sync.py verify 双语版表格.xlsx verification.txt
```

2. 根据验证结果更新翻译字典。

### 4. 格式问题

#### 问题：双语内容显示不完整

**现象**：
单元格中只显示中文，不显示英文。

**解决方案**：
1. 检查单元格是否设置了自动换行：
```python
from openpyxl.styles import Alignment
cell.alignment = Alignment(wrap_text=True)
```

2. 调整行高以适应双语内容：
```python
sheet.row_dimensions[row_num].height = 30
```

#### 问题：公式丢失

**现象**：
转换后公式变成静态值。

**解决方案**：
技能会保留公式，但如果公式引用了其他单元格，可能需要手动调整。

### 5. 同步问题

#### 问题：SHEET2没有同步更新

**现象**：
更新SHEET1后，SHEET2内容没有变化。

**解决方案**：
1. 确保调用了同步函数：
```python
syncer.sync_english_sheet()
```

2. 确保保存了文件：
```python
syncer.save(output_file)
```

3. 重新打开文件查看更新。

### 6. 性能问题

#### 问题：大文件转换很慢

**现象**：
处理大型Excel文件时耗时很长。

**解决方案**：
1. 分批处理大文件。
2. 只处理需要翻译的列。
3. 使用更高效的翻译算法。

### 7. 其他问题

#### 问题：文件被占用

**错误信息**：
```
PermissionError: [Errno 13] Permission denied: '文件名.xlsx'
```

**解决方案**：
1. 关闭Excel程序。
2. 等待几秒后重试。
3. 或者复制文件后处理副本。

#### 问题：内存不足

**错误信息**：
```
MemoryError
```

**解决方案**：
1. 关闭其他程序释放内存。
2. 分批处理大文件。
3. 增加系统虚拟内存。

## 调试技巧

### 1. 启用详细日志

在脚本中添加详细日志：
```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

### 2. 检查中间结果

保存中间结果以便调试：
```python
# 保存分析结果
syncer.analyze(input_file, 'debug_analysis.txt')

# 保存翻译结果
syncer.convert(input_file, 'debug_output.xlsx')
```

### 3. 验证翻译字典

检查翻译字典是否正确：
```python
print(f"翻译字典包含 {len(translations)} 个条目")
for key, value in translations.items():
    print(f"{key} -> {value}")
```

## 获取帮助

### 1. 查看文档

- README.md：项目说明
- USAGE.md：使用说明
- INSTALL.md：安装说明

### 2. 运行测试

```bash
python test_skill.py
python test_cases.py
```

### 3. 检查版本

```bash
python -c "import openpyxl; print(openpyxl.__version__)"
python --version
```

### 4. 提交问题

如果问题仍然存在，请提供以下信息：
1. 错误信息
2. 操作系统版本
3. Python版本
4. openpyxl版本
5. 输入文件示例
6. 期望的输出结果

## 预防措施

### 1. 备份原始文件

在转换前备份原始文件：
```bash
copy 原始文件.xlsx 原始文件_备份.xlsx
```

### 2. 使用版本控制

使用Git管理文件版本：
```bash
git init
git add .
git commit -m "初始版本"
```

### 3. 定期验证

定期验证翻译结果：
```bash
python scripts/bilingual_excel_sync.py verify 双语版表格.xlsx verification.txt
```

### 4. 更新翻译字典

根据实际需求更新翻译字典，确保翻译准确。