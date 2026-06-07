#!/bin/bash
# 澳門 DICJ 博彩統計 — 首次完整下載（所有年份 2002 至今）
# 雙擊此檔案（macOS）即可執行

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "================================================"
echo "  DICJ 博彩統計資料 — 首次完整下載"
echo "  年份範圍：2002 年至今"
echo "================================================"
echo ""
echo "⚠️  首次下載需要較長時間（約10-30分鐘）"
echo "   請確保網絡連線正常"
echo ""
read -p "按 Enter 開始下載，Ctrl+C 取消..."
echo ""

# 安裝依賴
echo "📦 安裝 Python 依賴..."
python3 -m pip install --quiet requests beautifulsoup4 openpyxl lxml 2>/dev/null
echo "   ✓ 完成"
echo ""

# 執行完整爬取
python3 dicj_gaming_scraper.py --year 2002 $(date +%Y)

OUTPUT="$SCRIPT_DIR/DICJ_博彩統計.xlsx"
if [ -f "$OUTPUT" ]; then
    echo ""
    echo "🎉 下載完成！"
    if [[ "$OSTYPE" == "darwin"* ]]; then
        open "$OUTPUT"
    elif [[ "$OSTYPE" == "linux-gnu"* ]]; then
        xdg-open "$OUTPUT" 2>/dev/null || echo "請手動打開：$OUTPUT"
    fi
fi

echo ""
read -p "按 Enter 結束..."
