#!/bin/bash
# 澳門政府部門通訊錄 - 更新程序（更新資料 + 發佈到 GitHub Pages）
# 雙擊此檔案即可自動執行

cd "$(dirname "$0")"
echo ""
echo "========================================="
echo "  澳門政府部門通訊錄 - 更新程序"
echo "  (更新資料 → 發佈到 GitHub Pages)"
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

# 步驟 1：從 gov.mo 下載並更新資料
echo "📥 步驟 1/2：從 gov.mo 下載最新資料..."
echo ""
python3 updatedata.py

# 檢查 updatedata.py 是否成功
if [ $? -ne 0 ]; then
    echo ""
    echo "⚠️  資料更新未完成，不會發佈到網站。"
    echo ""
    echo "按任意鍵關閉..."
    read -n 1
    exit 1
fi

# 步驟 2：發佈到 GitHub Pages
echo ""
echo "========================================="
echo "📤 步驟 2/2：發佈到 GitHub Pages..."
echo "========================================="
echo ""

# 檢查 git
if ! command -v git &> /dev/null; then
    echo "❌ 找不到 git，請先安裝 Git"
    echo ""
    echo "按任意鍵關閉..."
    read -n 1
    exit 1
fi

# 檢查是否有變更
if git diff --quiet index.html 2>/dev/null && git diff --quiet Rawdata/ 2>/dev/null; then
    echo "ℹ️  沒有檔案變更，無需發佈。"
else
    git add index.html Rawdata/
    TODAY=$(date +%Y-%m-%d)
    git commit -m "更新資料庫 $TODAY"

    echo ""
    echo "⬆️  正在推送到 GitHub..."
    git push origin main

    if [ $? -eq 0 ]; then
        echo ""
        echo "✅ 已成功發佈到 GitHub Pages！"
        echo "🌐 網站：https://ronlam1981.github.io/macau-gov-directory/"
        echo ""
        echo "⏳ 網站通常需要 1-2 分鐘完成部署。"
    else
        echo ""
        echo "❌ 推送失敗，請檢查網絡連線或 Git 設定。"
    fi
fi

echo ""
echo "按任意鍵關閉..."
read -n 1
