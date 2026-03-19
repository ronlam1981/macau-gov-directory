#!/bin/bash
# 澳門政府部門通訊錄 - 更新資料
# 雙擊此檔案即可自動執行

cd "$(dirname "$0")"
echo ""
echo "========================================="
echo "  澳門政府部門通訊錄 - 更新資料"
echo "========================================="
echo ""

# 檢查 Python3
if ! command -v python3 &> /dev/null; then
    echo "❌ 找不到 python3，請先安裝 Python 3"
    echo ""
    echo "按任意鍵關閉..."
    read -n 1
    exit 1
fi

# 檢查 pandas
python3 -c "import pandas" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "⚠️  正在安裝所需套件 pandas + openpyxl..."
    pip3 install pandas openpyxl
    echo ""
fi

python3 updatedata.py

echo ""
echo "按任意鍵關閉..."
read -n 1
