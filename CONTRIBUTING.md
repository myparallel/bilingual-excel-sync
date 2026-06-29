# 双语Excel表格同步技能贡献指南

感谢您对双语Excel表格同步技能的关注！我们欢迎任何形式的贡献。

## 如何贡献

### 1. 报告问题

如果您发现了bug或有功能建议，请：

1. 检查是否已有相关issue
2. 创建新的issue，包含以下信息：
   - 问题描述
   - 复现步骤
   - 期望行为
   - 实际行为
   - 环境信息（操作系统、Python版本等）

### 2. 提交代码

如果您想提交代码修复或新功能：

1. Fork项目
2. 创建您的特性分支：`git checkout -b feature/AmazingFeature`
3. 提交您的更改：`git commit -m 'Add some AmazingFeature'`
4. 推送到分支：`git push origin feature/AmazingFeature`
5. 打开一个Pull Request

### 3. 改进文档

文档改进同样重要：

1. 修复拼写错误
2. 添加使用示例
3. 翻译文档
4. 添加FAQ

### 4. 添加翻译

如果您想添加新的翻译：

1. 编辑翻译字典文件
2. 确保翻译准确
3. 添加测试用例
4. 提交PR

## 开发环境设置

### 1. 克隆项目

```bash
git clone https://github.com/yourusername/bilingual-excel-sync.git
cd bilingual-excel-sync
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 运行测试

```bash
python test_skill.py
python test_cases.py
```

## 代码规范

### 1. Python代码风格

- 遵循PEP 8规范
- 使用4个空格缩进
- 行长度不超过79字符
- 使用有意义的变量名

### 2. 文档字符串

使用Google风格的文档字符串：

```python
def function_name(param1, param2):
    """函数描述。
    
    Args:
        param1: 参数1描述
        param2: 参数2描述
    
    Returns:
        返回值描述
    
    Raises:
        ValueError: 异常描述
    """
    pass
```

### 3. 注释

- 解释为什么，而不是是什么
- 保持注释与代码同步
- 避免过多的注释

### 4. 提交信息

使用清晰的提交信息：

```
feat: 添加新功能
fix: 修复bug
docs: 更新文档
test: 添加测试
refactor: 重构代码
style: 代码格式调整
chore: 构建过程或辅助工具的变动
```

## 测试指南

### 1. 编写测试

为每个新功能编写测试：

```python
def test_new_feature():
    """测试新功能"""
    # 准备测试数据
    test_data = create_test_data()
    
    # 执行测试
    result = new_feature(test_data)
    
    # 验证结果
    assert result == expected_result
```

### 2. 运行测试

```bash
# 运行所有测试
python -m pytest tests/

# 运行特定测试
python -m pytest tests/test_bilingual_sync.py

# 运行带覆盖率的测试
python -m pytest --cov=bilingual_excel_sync tests/
```

### 3. 测试覆盖率

确保测试覆盖率不低于80%：

```bash
coverage run -m pytest tests/
coverage report
coverage html  # 生成HTML报告
```

## 翻译指南

### 1. 添加新翻译

1. 找到需要翻译的中文字符串
2. 确定准确的英文翻译
3. 添加到翻译字典
4. 添加测试用例

### 2. 翻译原则

- 准确性：翻译要准确表达原意
- 一致性：相同术语使用相同翻译
- 简洁性：翻译要简洁明了
- 专业性：使用行业术语

### 3. 翻译字典结构

```python
translations = {
    # 列标题
    '项目名称': 'Project Name',
    
    # 状态值
    '已付': 'Paid',
    
    # 专业术语
    '压力传感器': 'Pressure Sensor',
}
```

## 文档指南

### 1. 文档结构

- README.md：项目概述
- USAGE.md：详细使用说明
- INSTALL.md：安装指南
- TROUBLESHOOTING.md：故障排除
- CONTRIBUTING.md：贡献指南

### 2. 文档风格

- 使用清晰的语言
- 提供代码示例
- 包含截图或图表
- 保持文档更新

### 3. 文档模板

```markdown
# 标题

简短描述

## 功能特点

- 特点1
- 特点2

## 使用方法

### 步骤1

说明和示例

### 步骤2

说明和示例

## 注意事项

- 注意1
- 注意2
```

## 发布流程

### 1. 版本号

使用语义化版本号：

- MAJOR.MINOR.PATCH
- 例如：1.0.0

### 2. 发布步骤

1. 更新版本号
2. 更新CHANGELOG.md
3. 创建Git标签
4. 发布到PyPI（如果适用）

### 3. 发布检查清单

- [ ] 所有测试通过
- [ ] 文档已更新
- [ ] 版本号已更新
- [ ] CHANGELOG已更新
- [ ] 无安全漏洞

## 社区准则

### 1. 行为准则

- 尊重他人
- 接受建设性批评
- 专注于对社区最有利的事情
- 对他人表示同理心

### 2. 沟通方式

- 使用清晰的语言
- 避免专业术语
- 提供上下文
- 保持耐心

### 3. 冲突解决

- 保持冷静
- 寻求理解
- 寻找共同点
- 必要时寻求调解

## 许可证

本项目采用MIT许可证。贡献代码即表示您同意您的贡献将在MIT许可证下发布。

## 联系方式

如有问题，请通过以下方式联系我们：

- 创建GitHub Issue
- 发送邮件至项目维护者

## 致谢

感谢所有贡献者的付出！