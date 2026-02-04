#!/bin/bash
# PDF加密邮件发送系统 - Linux版

echo ""
echo "===================================="
echo "  PDF加密邮件发送系统 (Excel配置版)"
echo "===================================="
echo ""

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未检测到Python3，请先安装Python3"
    exit 1
fi

# 获取脚本所在目录
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# 运行脚本
python3 pdf_encrypt_send.py

echo ""
echo "按Enter键退出..."
read
