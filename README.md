# MySQL 数据库文档生成器 (MySQL DocGen)

这是一个简单易用的 Python 桌面应用程序，旨在帮助开发者快速将 MySQL 数据库的表结构信息导出为格式化的 **Word (`.docx`)**、**HTML (`.html`)** 和 **Markdown (`.md`)** 文档。

![Python Version](https://img.shields.io/badge/python-3.x-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## 🚀 功能特点

- **多格式支持**：支持导出为 Word、HTML 和 Markdown 三种常用格式，满足不同场景需求。
- **连接方便**：支持通过主机、端口、用户名和密码连接到任何 MySQL 数据库。
- **实时过滤**：提供搜索框，可根据表名或注释实时筛选需要导出的数据表。
- **批量导出**：支持全选或多选数据表，一键生成完整的数据库设计文档。
- **文档美观**：导出的文档包含清晰的表格，涵盖字段名、类型、长度、是否允许空、默认值及注释。
- **独立运行**：支持使用 PyInstaller 打包成 EXE 文件，无需 Python 环境即可在 Windows 上运行。

## 🛠️ 技术栈

- **语言**：Python 3
- **界面**：Tkinter (TTK)
- **数据库驱动**：PyMySQL
- **文档处理**：python-docx

## 📦 快速开始

### 1. 克隆或下载项目
确保您的系统已安装 Python 3.x。

### 2. 安装依赖
在项目根目录下运行以下命令安装必要的库：
```bash
pip install -r requirements.txt
```

### 3. 运行程序
```bash
python doc_gen.py
```

## 🔨 打包为 EXE (Windows)

如果您希望生成一个独立的可执行文件，可以使用以下命令：

```bash
pyinstaller MySQL_Doc_Gen.spec
```

打包完成后，在 `dist` 文件夹下即可找到 `MySQL_Doc_Gen.exe`。

## 📖 使用指南

1. **输入连接信息**：填写 MySQL 服务器地址、端口、用户名、密码和数据库名。
2. **连接数据库**：点击“连接数据库”按钮。成功后，中间列表会显示该库下的所有数据表。
3. **选择数据表**：
   - 使用搜索框快速定位。
   - 在列表中点击选择（支持 Ctrl/Shift 多选）。
   - 或者点击“全选”。
4. **生成文档**：根据需要点击 **“导出 Word”**、**“导出 HTML”** 或 **“导出 Markdown”** 按钮，选择保存路径即可。

## 📝 许可证

本项目采用 MIT 许可证。您可以自由使用、修改和分发。

---

  

![image-20260205231232768](images/image-20260205231232768.png)