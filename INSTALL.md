# 双语Excel表格同步技能安装说明

## 系统要求

- Python 3.6 或更高版本
- 操作系统：Windows、macOS、Linux

## 安装步骤

### 1. 安装Python依赖

```bash
pip install openpyxl
```

### 2. 下载技能文件

将以下文件下载到本地目录：

```
bilingual-excel-sync/
├── SKILL.md
├── README.md
├── USAGE.md
├── INSTALL.md
├── test_skill.py
├── test_cases.py
├── evals/
│   └── evals.json
└── scripts/
    └── bilingual_excel_sync.py
```

### 3. 配置环境变量（可选）

如果需要在命令行中直接使用技能，可以将技能目录添加到PATH环境变量：

#### Windows

```powershell
# 添加到用户环境变量
[Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\path\to\bilingual-excel-sync", "User")
```

#### macOS/Linux

```bash
# 添加到bashrc或zshrc
export PATH="$PATH:/path/to/bilingual-excel-sync"
```

## 验证安装

### 1. 运行测试脚本

```bash
python test_skill.py
```

如果看到以下输出，说明安装成功：

```
创建测试文件: 测试表格.xlsx
加载文件: 测试表格.xlsx
转换为双语格式...
同步生成英文表...
调整布局...
保存文件: 测试表格_中英双语版.xlsx
转换完成！

=== 验证结果 ===
工作表列表: ['Sheet1 双语版 Bilingual', 'Sheet2 English Version']

SHEET1内容:
行1: ('NO.', 'ITEM名称\nItem Name', '规格\nSpecification', '数量\nQuantity', '单位\nUnit', '备注\nRemarks')
行2: ('1', '土壤传感器\nSoil Sensor', 'DG-214-4G-P2', 20, '台\nUnit/Set', '标准盒装\nStandard Box')
行3: ('2', '振动传感器\nVibration Sensor', 'PG-T920-4G-P2', 2, '台\nUnit/Set', '标准盒装\nStandard Box')

SHEET2内容:
行1: ('NO.', 'Item Name', 'Specification', 'Quantity', 'Unit', 'Remarks')
行2: ('1', 'Soil Sensor', 'DG-214-4G-P2', 20, 'Unit/Set', 'Standard Box')
行3: ('2', 'Vibration Sensor', 'PG-T920-4G-P2', 2, 'Unit/Set', 'Standard Box')

清理测试文件: 测试表格.xlsx

测试完成！输出文件: 测试表格_中英双语版.xlsx
```

### 2. 运行所有测试用例

```bash
python test_cases.py
```

如果看到所有测试用例运行完成，说明技能功能正常。

## 使用技能

### 1. 命令行使用

```bash
# 分析文件
python scripts/bilingual_excel_sync.py analyze 发货清单.xlsx analysis.txt

# 执行转换
python scripts/bilingual_excel_sync.py convert 发货清单.xlsx 双语版_发货清单.xlsx

# 验证结果
python scripts/bilingual_excel_sync.py verify 双语版_发货清单.xlsx verification.txt
```

### 2. 在AI对话中使用

直接告诉AI：

```
请将"发货清单.xlsx"制作成双语表格，并同步生成纯英文的SHEET2
```

AI会自动调用技能完成转换。

## 常见问题

### Q: 安装openpyxl失败

A: 尝试使用以下命令：

```bash
pip install --user openpyxl
```

或者使用conda：

```bash
conda install openpyxl
```

### Q: 运行脚本时提示"ModuleNotFoundError: No module named 'openpyxl'"

A: 确保已安装openpyxl：

```bash
pip list | grep openpyxl
```

如果没有安装，运行：

```bash
pip install openpyxl
```

### Q: 中文显示乱码

A: 这是Windows PowerShell的编码问题，不影响实际文件。将结果输出到文件即可：

```bash
python scripts/bilingual_excel_sync.py analyze 发货清单.xlsx analysis.txt
```

然后用文本编辑器打开analysis.txt查看结果。

### Q: 翻译不准确

A: 技能使用内置翻译字典，如果翻译不准确，可以：

1. 在对话中告诉AI正确的翻译
2. 修改脚本中的翻译字典

## 卸载技能

删除技能目录即可：

```bash
# Windows
rmdir /s /q "C:\path\to\bilingual-excel-sync"

# macOS/Linux
rm -rf /path/to/bilingual-excel-sync
```

## 技术支持

如有问题或建议，请：

1. 查看README.md和USAGE.md
2. 运行test_skill.py验证安装
3. 检查Python和openpyxl版本

## 更新日志

### v1.0.0
- 初始版本
- 支持双语转换
- 支持同步生成纯英文SHEET2
- 内置常用翻译字典