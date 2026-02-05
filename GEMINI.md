# MySQL 数据库文档生成器 (DocGen)

这是一个基于 Python 开发的桌面工具，旨在帮助开发者快速将 MySQL 数据库的表结构导出为整洁的 Microsoft Word (.docx) 文档。

## 项目概览

- **核心功能**：
  - 支持连接本地或远程 MySQL 数据库。
  - 自动获取数据库中所有表的列表及注释。
  - 支持按表名或注释进行实时搜索过滤。
  - 批量选择表并生成包含详细字段信息的 Word 表格文档。
- **技术栈**：
  - **GUI 框架**：`tkinter` & `ttk` (原生 Python GUI)。
  - **数据库连接**：`pymysql`。
  - **文档处理**：`python-docx`。
  - **构建工具**：`PyInstaller`。

## 构建与运行

### 运行环境准备
项目需要 Python 3.x 环境。首先安装必要的第三方库：
```bash
pip install -r requirements.txt
```

### 直接运行脚本
```bash
python doc_gen.py
```

### 打包为独立可执行文件 (Windows)
项目已配置好 PyInstaller 规范文件，可以使用以下命令打包成无需 Python 环境即可运行的 EXE 文件：
```bash
pyinstaller MySQL_Doc_Gen.spec
```
打包后的文件将位于 `dist/` 目录下。

## 项目结构与开发规范

- `doc_gen.py`: 项目的主入口和核心逻辑。所有的 GUI 布局、数据库查询和文档生成逻辑都封装在 `DBDocGeneratorApp` 类中。
- `MySQL_Doc_Gen.spec`: PyInstaller 配置文件，用于定义打包行为（如窗口模式、名称等）。
- **编码风格**：
  - 遵循 PEP 8 规范。
  - UI 样式通过 `ttk.Style` 进行统一配置，优先使用 `Microsoft YaHei UI` 字体以适配中文 Windows 环境。
  - 数据库查询使用 `information_schema.COLUMNS` 以获取最准确的列定义和注释。
- **注意事项**：
  - 在生成 Word 文档时，使用了 `Microsoft YaHei` 字体以确保中文字符显示正常。
  - 脚本包含 Windows 任务栏图标修复逻辑 (`SetCurrentProcessExplicitAppUserModelID`)。

## 待办事项 / 未来改进
- [ ] 支持导出到 PDF 格式。
- [ ] 支持更多的数据库类型（如 PostgreSQL, SQLite）。
- [ ] 增加自定义 Word 模板功能。
