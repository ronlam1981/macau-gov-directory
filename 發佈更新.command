#!/bin/bash
# ============================================================
#  澳門政府部門通訊錄 - 一鍵發佈更新
#  功能：更新資料 + 發佈/更新 GitHub Pages 網站
#  用法：雙擊此檔案即可自動執行
# ============================================================
#
#  首次使用：初始化 Git → 建立 GitHub 倉庫 → 啟用 Pages → 發佈
#  之後使用：下載最新資料 → 對比更新 → 推送到 GitHub Pages
# ============================================================

cd "$(dirname "$0")"
PROJECT_DIR="$(pwd)"
REPO_NAME="$(basename "$PROJECT_DIR")"

echo ""
echo "╔═════════════════════════════════════════╗"
echo "║   澳門政府部門通訊錄 - 一鍵發佈更新    ║"
echo "╚═════════════════════════════════════════╝"
echo ""

# ─── 環境檢查 ──────────────────────────────────
check_tool() {
    if ! command -v "$1" &> /dev/null; then
        echo "❌ 找不到 $1，$2"
        echo ""
        echo "按任意鍵關閉..."
        read -n 1
        exit 1
    fi
}

check_tool "python3" "請先安裝 Python 3（https://www.python.org）"
check_tool "git" "請先安裝 Git（https://git-scm.com）"

# 檢查/安裝 pandas
python3 -c "import pandas" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "⚠️  正在安裝所需 Python 套件..."
    pip3 install pandas openpyxl
    echo ""
fi

# 檢查 gh CLI
GH_CMD=""
if command -v gh &> /dev/null; then
    GH_CMD="gh"
elif [ -f "/tmp/gh_extract/gh_2.88.1_macOS_amd64/bin/gh" ]; then
    GH_CMD="/tmp/gh_extract/gh_2.88.1_macOS_amd64/bin/gh"
fi

# ─── 步驟 1：更新資料 ────────────────────────────
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "📥 步驟 1/3：從 gov.mo 下載最新資料"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

if [ -f "updatedata.py" ]; then
    python3 updatedata.py
    UPDATE_RESULT=$?
else
    echo "⚠️  找不到 updatedata.py，跳過資料更新"
    UPDATE_RESULT=0
fi

if [ $UPDATE_RESULT -ne 0 ]; then
    echo ""
    echo "⚠️  資料更新步驟失敗或取消。"
    echo ""
    read -p "是否繼續發佈現有版本到 GitHub Pages？(y/n): " CONTINUE
    if [[ "$CONTINUE" != "y" && "$CONTINUE" != "Y" && "$CONTINUE" != "" ]]; then
        echo "❌ 已取消。"
        echo ""
        echo "按任意鍵關閉..."
        read -n 1
        exit 1
    fi
fi

# ─── 步驟 2：初始化 Git（首次使用）──────────────────
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "🔧 步驟 2/3：Git 倉庫設定"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

FIRST_TIME=false

# 初始化 Git（如果尚未初始化）
if [ ! -d ".git" ]; then
    FIRST_TIME=true
    echo "📦 首次發佈：初始化 Git 倉庫..."
    git init

    # 建立 .gitignore
    if [ ! -f ".gitignore" ]; then
        cat > .gitignore << 'GITIGNORE'
.DS_Store
__pycache__/
*.pyc
.claude/
GITIGNORE
    fi

    git add .gitignore
    echo "   ✓ Git 倉庫已初始化"
fi

# 檢查是否有遠端倉庫
REMOTE_URL=$(git remote get-url origin 2>/dev/null)

if [ -z "$REMOTE_URL" ]; then
    FIRST_TIME=true
    echo ""
    echo "🌐 首次發佈：建立 GitHub 遠端倉庫..."

    # 檢查 gh CLI
    if [ -z "$GH_CMD" ]; then
        echo ""
        echo "❌ 找不到 GitHub CLI (gh)。"
        echo "   請先安裝：https://cli.github.com"
        echo "   或執行：brew install gh"
        echo ""
        echo "   安裝後請重新執行此程序。"
        echo ""
        echo "按任意鍵關閉..."
        read -n 1
        exit 1
    fi

    # 檢查 gh 登入狀態
    $GH_CMD auth status &>/dev/null
    if [ $? -ne 0 ]; then
        echo "   需要登入 GitHub..."
        $GH_CMD auth login --web -p https
        if [ $? -ne 0 ]; then
            echo "❌ GitHub 登入失敗。"
            echo ""
            echo "按任意鍵關閉..."
            read -n 1
            exit 1
        fi
    fi

    # 確保 git credential 使用 gh
    $GH_CMD auth setup-git 2>/dev/null

    # Stage 所有需要的檔案
    git add index.html
    [ -f "updatedata.py" ] && git add updatedata.py
    [ -f "updatedata.command" ] && git add updatedata.command
    [ -f "更新程序.py" ] && git add "更新程序.py"
    [ -f "更新程序.command" ] && git add "更新程序.command"
    [ -f "發佈更新.command" ] && git add "發佈更新.command"
    [ -d "Rawdata" ] && git add Rawdata/
    [ -d ".github" ] && git add .github/

    git commit -m "初始版本：$(basename "$PROJECT_DIR")"

    # 建立 GitHub 倉庫
    echo ""
    read -p "GitHub 倉庫名稱 [$REPO_NAME]: " CUSTOM_NAME
    CUSTOM_NAME=${CUSTOM_NAME:-$REPO_NAME}

    $GH_CMD repo create "$CUSTOM_NAME" --public --source=. --push \
        --description "澳門政府部門通訊錄 - 可搜尋的政府部門聯絡資料目錄" 2>&1

    if [ $? -ne 0 ]; then
        echo "❌ 建立 GitHub 倉庫失敗。"
        echo ""
        echo "按任意鍵關閉..."
        read -n 1
        exit 1
    fi

    REMOTE_URL=$(git remote get-url origin 2>/dev/null)
    GH_USER=$($GH_CMD api user --jq .login 2>/dev/null)
    echo "   ✓ 倉庫已建立：$REMOTE_URL"

    # 建立 GitHub Actions workflow
    if [ ! -f ".github/workflows/deploy.yml" ]; then
        mkdir -p .github/workflows
        cat > .github/workflows/deploy.yml << 'WORKFLOW'
name: Deploy to GitHub Pages

on:
  push:
    branches: ["main"]
  workflow_dispatch:

permissions:
  contents: read
  pages: write
  id-token: write

concurrency:
  group: "pages"
  cancel-in-progress: false

jobs:
  deploy:
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
    runs-on: ubuntu-latest
    steps:
      - name: Checkout
        uses: actions/checkout@v4
      - name: Setup Pages
        uses: actions/configure-pages@v5
      - name: Upload artifact
        uses: actions/upload-pages-artifact@v3
        with:
          path: '.'
      - name: Deploy to GitHub Pages
        id: deployment
        uses: actions/deploy-pages@v4
WORKFLOW

        git add .github/workflows/deploy.yml
        git commit -m "加入 GitHub Actions 自動部署"
        git push origin main 2>&1
    fi

    # 啟用 GitHub Pages
    echo ""
    echo "🔄 啟用 GitHub Pages..."
    $GH_CMD api "repos/$GH_USER/$CUSTOM_NAME/pages" -X POST \
        --input - 2>/dev/null <<JSON
{"build_type":"workflow","source":{"branch":"main","path":"/"}}
JSON

    PAGES_URL="https://$GH_USER.github.io/$CUSTOM_NAME/"
    echo "   ✓ GitHub Pages 已啟用"
    echo "   🌐 網址：$PAGES_URL"

else
    echo "   ✓ 遠端倉庫：$REMOTE_URL"

    # 從遠端 URL 提取 Pages URL
    GH_USER=$(echo "$REMOTE_URL" | sed -n 's|https://github.com/\([^/]*\)/.*|\1|p')
    REPO_SLUG=$(echo "$REMOTE_URL" | sed -n 's|https://github.com/[^/]*/\([^/.]*\).*|\1|p')
    PAGES_URL="https://$GH_USER.github.io/$REPO_SLUG/"
fi

# ─── 步驟 3：發佈到 GitHub Pages ──────────────────
echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "📤 步驟 3/3：發佈到 GitHub Pages"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

# Stage 變更的檔案
git add index.html 2>/dev/null
git add Rawdata/ 2>/dev/null
git add "發佈更新.command" 2>/dev/null
git add updatedata.py updatedata.command 2>/dev/null
git add "更新程序.py" "更新程序.command" 2>/dev/null
git add .github/ 2>/dev/null

# 檢查是否有變更
if git diff --cached --quiet 2>/dev/null && [ "$FIRST_TIME" = false ]; then
    echo "ℹ️  沒有檔案變更，無需發佈。"
    echo "   現有網站：$PAGES_URL"
else
    TODAY=$(date +%Y-%m-%d)

    if [ "$FIRST_TIME" = false ]; then
        git commit -m "更新資料庫 $TODAY"
    fi

    echo "⬆️  正在推送到 GitHub..."
    git push origin main 2>&1

    if [ $? -eq 0 ]; then
        echo ""
        echo "╔═════════════════════════════════════════╗"
        echo "║         ✅ 發佈成功！                    ║"
        echo "╚═════════════════════════════════════════╝"
        echo ""
        echo "   🌐 網站：$PAGES_URL"
        echo "   📅 日期：$TODAY"
        echo ""
        echo "   ⏳ 首次部署約需 1-2 分鐘生效"
    else
        echo ""
        echo "❌ 推送失敗！"
        echo "   請檢查網絡連線或 Git 設定。"
        echo ""
        echo "   可嘗試手動執行："
        echo "   cd \"$PROJECT_DIR\" && git push origin main"
    fi
fi

echo ""
echo "按任意鍵關閉..."
read -n 1
