#!/bin/bash

echo "安装Excel Image Reader依赖..."

# 检查Node.js是否安装
if ! command -v node &> /dev/null; then
    echo "错误: 请先安装Node.js"
    exit 1
fi

# 检查npm是否安装
if ! command -v npm &> /dev/null; then
    echo "错误: 请先安装npm"
    exit 1
fi

# 安装依赖
echo "正在安装依赖包..."
npm install

# 编译TypeScript
echo "正在编译TypeScript..."
npm run build

echo "安装完成！"
echo ""
echo "使用方法："
echo "1. 将你的Excel文件重命名为test.xlsx并放在example目录下"
echo "2. 运行: npm run test"
echo "3. 或者运行: node dist/example/example.js"
echo ""
echo "高级示例："
echo "node dist/example/advanced-example.js"
