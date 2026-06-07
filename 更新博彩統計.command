#!/bin/bash
# 澳門 DICJ 博彩統計資料更新工具
# 雙擊此檔案（macOS）即可執行

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "=============================="
echo "  DICJ 博彩統計更新程序"
echo "=============================="

# 檢查 Python3
if ! command -v python3 &>/dev/null; then
    echo "❌ 找不到 Python3，請先安裝 Python3"
    read -p "按 Enter 結束..."
    exit 1
fi

# 安裝依賴
echo ""
echo "📦 檢查 Python 依賴..."
python3 -m pip install --quiet requests beautifulsoup4 openpyxl lxml 2>/dev/null
echo "   ✓ 依賴已就緒"

# 執行更新（--update 只在有新資料時重爬）
echo ""
python3 dicj_gaming_scraper.py --update

# 完成後打開 Excel
OUTPUT="$SCRIPT_DIR/DICJ_博彩統計.xlsx"
if [ -f "$OUTPUT" ]; then
    echo ""
    echo "📂 正在打開 Excel 檔案..."
    if [[ "$OSTYPE" == "darwin"* ]]; then
        open "$OUTPUT"
    elif [[ "$OSTYPE" == "linux-gnu"* ]]; then
        xdg-open "$OUTPUT" 2>/dev/null || echo "請手動打開：$OUTPUT"
    elif [[ "$OSTYPE" == "msys" || "$OSTYPE" == "win32" ]]; then
        start "$OUTPUT"
    fi
fi

echo ""
read -p "按 Enter 結束..."
