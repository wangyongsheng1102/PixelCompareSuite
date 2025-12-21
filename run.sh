#!/bin/bash

# PixelCompareSuite 运行脚本

echo "正在恢复 NuGet 包..."
dotnet restore

if [ $? -ne 0 ]; then
    echo "错误: 包恢复失败"
    exit 1
fi

echo "正在构建项目..."
dotnet build

if [ $? -ne 0 ]; then
    echo "错误: 构建失败"
    exit 1
fi

echo "正在运行应用程序..."
dotnet run

