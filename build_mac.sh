#!/bin/bash

# MySQL DocGen macOS 打包脚本

echo "开始为 macOS 打包 MySQL 数据库文档生成器..."

# 1. 安装依赖
echo "正在检查并安装依赖..."
pip install -r requirements.txt
pip install pyinstaller

# 2. 执行打包
# --onefile: 打包成单个可执行文件
# --windowed: 运行时不显示终端窗口 (macOS 下生成 .app 包)
# --clean: 清理临时文件
echo "正在开始 PyInstaller 打包流程..."
pyinstaller --noconfirm --onefile --windowed 
  --name "MySQL_Doc_Gen" 
  --osx-bundle-identifier "com.tyzou.mysqldocgen" 
  --clean 
  doc_gen.py

echo "================================================="
echo "打包完成！"
echo "可执行文件位于: dist/MySQL_Doc_Gen"
echo "macOS App 位于: dist/MySQL_Doc_Gen.app"
echo "================================================="
