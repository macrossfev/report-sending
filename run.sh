#!/bin/bash
# PDF加密邮件发送系统 - Linux版

echo ""
echo "===================================="
echo "  PDF加密邮件发送系统 (Excel配置版)"
echo "===================================="
echo ""

# 检查Python是否安装
PYTHON_CMD=""
if command -v python3 >/dev/null 2>&1; then
    PYTHON_CMD="python3"
elif [ -x /usr/bin/python3 ]; then
    PYTHON_CMD="/usr/bin/python3"
else
    echo "[错误] 未检测到Python3，请先安装Python3"
    exit 1
fi

# 获取脚本所在目录
SCRIPT_DIR="$( cd "$( dirname "$0" )" && pwd )"
cd "$SCRIPT_DIR"

# 运行脚本
$PYTHON_CMD pdf_encrypt_send.py

echo ""
echo "按Enter键退出..."
read dummy
