# 双语Excel表格同步技能常见问题解答

## 基本问题

### Q: 这个技能是做什么的？
A: 这个技能可以将中文Excel表格自动转换为中英双语版本，并同步生成纯英文的SHEET2工作表。

### Q: 什么时候应该使用这个技能？
A: 当您需要：
- 将中文表格翻译成英文
- 生成中英双语对照表格
- 为国外客户准备英文表格
- 国际团队协作时使用

### Q: 这个技能支持哪些Excel格式？
A: 支持：
- .xlsx（Excel 2007+）
- .xlsm（启用宏的工作簿）
- .xls（Excel 97-2003）

### Q: 这个技能是免费的吗？
A: 是的，这个技能是完全免费的，采用MIT许可证。

## 安装问题

### Q: 如何安装这个技能？
A: 只需要安装Python依赖：
```bash
pip install openpyxl
```

### Q: 需要安装什么Python版本？
A: 需要Python 3.6或更高版本。

### Q: 安装时出现错误怎么办？
A: 尝试以下解决方案：
```bash
pip install --user openpyxl
```
或者使用conda：
```bash
conda install openpyxl
```

## 使用问题

### Q: 如何使用这个技能？
A: 在AI对话中说：
```
请将"文件名.xlsx"制作成双语表格
```

### Q: 如何自定义翻译？
A: 告诉AI您的翻译需求：
```
请将以下中文翻译成英文：
- 术语1 → English 1
- 术语2 → English 2
```

### Q: 如何同步更新SHEET2？
A: 告诉AI：
```
我更新了SHEET1，请同步更新SHEET2
```

### Q: 如何验证翻译的完整性？
A: 告诉AI：
```
请验证双语表格的翻译完整性
```

### Q: 为什么SHEET2中还有中文字符？
A: 可能的原因：
1. 翻译字典中没有对应的翻译
2. 翻译字典的key与文件中的字符串不完全一致
3. 字符串包含特殊字符或空格

**解决方案**：告诉AI需要翻译的具体内容。

### Q: 如何处理特殊格式？
A: 技能会自动处理：
1. 合并单元格
2. 单元格样式
3. 公式
4. 注释

### Q: 如何批量处理多个文件？
A: 告诉AI：
```
我有多个Excel文件需要翻译：
1. 文件1.xlsx
2. 文件2.xlsx
3. 文件3.xlsx
```

## 翻译问题

### Q: 翻译不准确怎么办？
A: 提供自定义翻译字典：
```
请使用以下翻译字典：
- 您的术语1 → English 1
- 您的术语2 → English 2
```

### Q: 如何添加专业术语？
A: 告诉AI专业术语的翻译：
```
请将以下专业术语翻译成英文：
- 压力监测变送器 → Pressure Monitoring Transmitter
- 液位监测变送器 → Level Monitoring Transmitter
```

### Q: 如何保持翻译的一致性？
A: 提供完整的翻译字典，确保相同术语使用相同翻译。

### Q: 内置翻译字典包含哪些内容？
A: 包含：
- 通用字段（项目名称、状态、数量等）
- 外贸术语（发货清单、装箱单、发票等）
- 传感器术语（压力传感器、温度传感器等）

## 格式问题

### Q: 如何保持原有格式？
A: 技能会自动：
1. 保留原有单元格样式（字体、颜色、边框）
2. 保留合并单元格信息
3. 保留公式（如SUM、=D7+D14等）

### Q: 双语内容显示不完整怎么办？
A: 检查单元格是否设置了自动换行：
1. 技能会自动设置自动换行
2. 调整行高以适应双语内容

### Q: 公式丢失了怎么办？
A: 技能会保留公式，但如果公式引用了其他单元格，可能需要手动调整。

## 性能问题

### Q: 处理大文件很慢怎么办？
A: 尝试以下优化：
1. 分批处理大文件
2. 只处理需要翻译的列
3. 使用更高效的翻译算法

### Q: 内存使用过高怎么办？
A: 尝试以下优化：
1. 关闭其他程序释放内存
2. 分批处理大文件
3. 增加系统虚拟内存

### Q: 文件被占用怎么办？
A: 尝试以下解决方案：
1. 关闭Excel程序
2. 等待几秒后重试
3. 或者复制文件后处理副本

## 技术问题

### Q: 如何使用Python API？
A: 示例代码：
```python
from scripts.bilingual_excel_sync import BilingualExcelSync

syncer = BilingualExcelSync()
syncer.load_translations(translations)
syncer.convert('输入.xlsx', '输出.xlsx')
```

### Q: 如何使用命令行？
A: 示例命令：
```bash
python scripts/bilingual_excel_sync.py convert 输入.xlsx 输出.xlsx
```

### Q: 如何运行演示？
A: 运行以下命令：
```bash
python demo.py
```

### Q: 如何运行测试？
A: 运行以下命令：
```bash
python test_skill.py
python test_cases.py
```

## 故障排除

### Q: 出现"ModuleNotFoundError"错误怎么办？
A: 确保已安装openpyxl：
```bash
pip install openpyxl
```

### Q: 出现"FileNotFoundError"错误怎么办？
A: 检查文件路径是否正确，文件是否存在。

### Q: 出现"PermissionError"错误怎么办？
A: 检查文件是否被其他程序占用，或者没有写入权限。

### Q: 出现"UnicodeDecodeError"错误怎么办？
A: 这是Windows PowerShell的编码问题，不影响实际文件。将结果输出到文件即可。

### Q: 中文显示乱码怎么办？
A: 这是Windows PowerShell的编码问题，不影响实际文件。将结果输出到文件即可。

## 获取帮助

### Q: 在哪里可以获取帮助？
A: 可以通过以下方式获取帮助：
1. 查看[USAGE.md](USAGE.md)详细说明
2. 运行`python demo.py`查看演示
3. 运行`python test_skill.py`验证功能

### Q: 如何报告问题？
A: 可以：
1. 创建GitHub Issue
2. 发送邮件至项目维护者

### Q: 如何贡献代码？
A: 请查看[CONTRIBUTING.md](CONTRIBUTING.md)了解详情。

### Q: 如何更新技能？
A: 重新下载最新版本的技能文件即可。

## 更多信息

### Q: 在哪里可以找到更多示例？
A: 查看以下文件：
- [examples.py](examples.py) - Python代码示例
- [example_使用示例.md](example_使用示例.md) - 使用示例
- [demo.py](demo.py) - 完整演示代码

### Q: 在哪里可以找到API参考？
A: 查看[API_REFERENCE.md](API_REFERENCE.md)。

### Q: 在哪里可以找到性能优化指南？
A: 查看[PERFORMANCE.md](PERFORMANCE.md)。

### Q: 在哪里可以找到故障排除指南？
A: 查看[TROUBLESHOOTING.md](TROUBLESHOOTING.md)。