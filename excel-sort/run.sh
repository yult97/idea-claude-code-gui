#!/bin/bash

# Excel需求排序工具启动脚本

echo "Excel需求排序工具"
echo "=================="

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到Python3，请先安装Python3"
    exit 1
fi

# 检查Node.js是否安装
if ! command -v node &> /dev/null; then
    echo "错误: 未找到Node.js，请先安装Node.js"
    exit 1
fi

# 启动后端服务
echo "正在启动后端服务..."
cd backend

# 检查虚拟环境
if [ ! -d "venv" ]; then
    echo "创建Python虚拟环境..."
    python3 -m venv venv
fi

# 激活虚拟环境
source venv/bin/activate

# 安装依赖
echo "安装Python依赖..."
pip install -r requirements.txt

# 启动后端服务（后台运行）
echo "启动Flask后端服务..."
python app.py &
BACKEND_PID=$!

echo "后端服务已启动 (PID: $BACKEND_PID) - http://localhost:5001"

# 启动前端服务
echo "正在启动前端服务..."
cd ../frontend

# 安装依赖
if [ ! -d "node_modules" ]; then
    echo "安装前端依赖..."
    npm install
fi

# 启动前端服务
echo "启动React前端服务..."
npm start

# 清理函数
cleanup() {
    echo "正在关闭服务..."
    kill $BACKEND_PID 2>/dev/null
    exit 0
}

# 捕获中断信号
trap cleanup SIGINT SIGTERM