from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = [line.strip() for line in fh if line.strip() and not line.startswith("#")]

setup(
    name="bilingual-excel-sync",
    version="1.0.0",
    author="OpenCode",
    author_email="opencode@example.com",
    description="将中文Excel表格转换为中英双语版本，并同步生成纯英文SHEET2工作表",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/bilingual-excel-sync",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Office/Business",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Text Processing :: Linguistic",
    ],
    python_requires=">=3.6",
    install_requires=requirements,
    entry_points={
        "console_scripts": [
            "bilingual-excel-sync=scripts.bilingual_excel_sync:main",
        ],
    },
    keywords="excel, bilingual, translation, english, chinese, spreadsheet, 外贸, 双语, 翻译",
    project_urls={
        "Bug Reports": "https://github.com/yourusername/bilingual-excel-sync/issues",
        "Source": "https://github.com/yourusername/bilingual-excel-sync",
    },
)