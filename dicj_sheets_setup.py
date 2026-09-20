#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
澳門 DICJ 博彩統計 — Google Sheets 初始設置工具
================================================
一次性工具，用於：
  1. 設置 Google Sheets API 憑證
  2. 建立新試算表並導入歷史數據
  3. 提供逐步設置說明

使用方法：
  python3 dicj_sheets_setup.py --setup     # 首次設置並導入資料
  python3 dicj_sheets_setup.py --update    # 更新現有試算表
  python3 dicj_sheets_setup.py --info      # 顯示設置說明
"""

import os
import sys
import json
import argparse

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# ══════════════════════════════════════════════════════════
# 設置說明
# ══════════════════════════════════════════════════════════

SETUP_INSTRUCTIONS = """
╔══════════════════════════════════════════════════════════════╗
║       Google Sheets API 設置說明（首次使用）              ║
╚══════════════════════════════════════════════════════════════╝

【方法一：Google Apps Script（推薦，無需 API 金鑰）】

  1. 前往 Google Sheets 網站，建立一個新的試算表
     https://sheets.google.com

  2. 點擊「擴充功能」→「Apps Script」

  3. 刪除預設代碼，將 dicj_apps_script.gs 的全部內容貼入

  4. 儲存（Ctrl+S），然後點擊「執行」→「initializeSheets」
     • 系統會要求授權，請點擊「審查權限」→「允許」

  5. 執行完成後，點擊「執行」→「updateAll」
     • 這會抓取所有可用的歷史資料

  6. 設定自動更新：點擊「執行」→「setupMonthlyTrigger」
     • 每月 5 日上午 9 時自動更新

────────────────────────────────────────────────────────────

【方法二：Google Sheets Python API（進階）】

  步驟 1：在 Google Cloud Console 建立專案
  ─────────────────────────────────────────
  a. 前往 https://console.cloud.google.com/
  b. 點擊頂部「選取專案」→「新增專案」
  c. 輸入名稱（例如：DICJ-Statistics），點擊「建立」

  步驟 2：啟用 Google Sheets API
  ────────────────────────────────
  a. 在搜尋欄搜尋「Google Sheets API」
  b. 點擊「啟用」
  c. 同樣搜尋「Google Drive API」並啟用

  步驟 3：建立 OAuth 2.0 憑證
  ─────────────────────────────
  a. 左側選單：「API 和服務」→「憑證」
  b. 點擊「建立憑證」→「OAuth 用戶端 ID」
  c. 應用程式類型：選「桌上型電腦應用程式」
  d. 輸入名稱（例如：DICJ Tool），點擊「建立」
  e. 下載 JSON 檔案，命名為 credentials.json
  f. 將 credentials.json 放到本程式同一目錄

  步驟 4：安裝 Python 套件
  ─────────────────────────
  pip3 install google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client openpyxl

  步驟 5：執行設置
  ─────────────────
  python3 dicj_sheets_setup.py --setup

────────────────────────────────────────────────────────────

【試算表試算表 ID】
  設置完成後，試算表 ID 會儲存在 google_sheet_id.txt
  也可在試算表的網址中找到：
  https://docs.google.com/spreadsheets/d/【這裡是ID】/edit

"""


def show_info():
    print(SETUP_INSTRUCTIONS)


# ══════════════════════════════════════════════════════════
# 檢查依賴
# ══════════════════════════════════════════════════════════

def check_dependencies():
    missing = []
    try:
        import google.auth
    except ImportError:
        missing.append("google-auth")
    try:
        import googleapiclient
    except ImportError:
        missing.append("google-api-python-client")
    try:
        import google_auth_oauthlib
    except ImportError:
        missing.append("google-auth-oauthlib")

    if missing:
        print("❌ 缺少必要套件，請執行以下命令安裝：")
        print(f"   pip3 install {' '.join(missing)} google-auth-httplib2")
        sys.exit(1)


# ══════════════════════════════════════════════════════════
# Google Sheets 連接
# ══════════════════════════════════════════════════════════

def get_service():
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds_file = os.path.join(SCRIPT_DIR, "credentials.json")
    token_file = os.path.join(SCRIPT_DIR, "token.json")

    if not os.path.exists(creds_file):
        print(f"❌ 找不到 credentials.json")
        print("   請依照說明建立 OAuth 憑證並下載至本目錄")
        print("   執行 python3 dicj_sheets_setup.py --info 查看詳細說明")
        sys.exit(1)

    creds = None
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(creds_file, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_file, "w") as f:
            f.write(creds.to_json())

    service = build("sheets", "v4", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    return service, drive_service


# ══════════════════════════════════════════════════════════
# 建立試算表
# ══════════════════════════════════════════════════════════

def create_spreadsheet(service):
    body = {
        "properties": {"title": "DICJ 澳門博彩統計資料庫"},
        "sheets": [
            {"properties": {"title": "每月幸運博彩毛收入", "index": 0}},
            {"properties": {"title": "每季各博彩項目", "index": 1}},
            {"properties": {"title": "年度彙總", "index": 2}},
            {"properties": {"title": "分析報告", "index": 3}},
            {"properties": {"title": "更新日誌", "index": 4}},
        ],
    }
    result = service.spreadsheets().create(body=body).execute()
    sheet_id = result["spreadsheetId"]
    print(f"✓ 試算表已建立，ID：{sheet_id}")

    id_file = os.path.join(SCRIPT_DIR, "google_sheet_id.txt")
    with open(id_file, "w") as f:
        f.write(sheet_id)
    print(f"✓ ID 已儲存至 {id_file}")
    return sheet_id


def get_sheet_id():
    id_file = os.path.join(SCRIPT_DIR, "google_sheet_id.txt")
    if os.path.exists(id_file):
        return open(id_file).read().strip()
    return None


# ══════════════════════════════════════════════════════════
# 讀取 Excel 資料
# ══════════════════════════════════════════════════════════

def read_excel_data():
    excel_file = os.path.join(SCRIPT_DIR, "DICJ_博彩統計.xlsx")
    if not os.path.exists(excel_file):
        print(f"❌ 找不到 {excel_file}")
        print("   請先執行 python3 dicj_gaming_scraper.py 生成 Excel 檔案")
        sys.exit(1)

    try:
        from openpyxl import load_workbook
    except ImportError:
        sys.exit("❌ 缺少 openpyxl：pip3 install openpyxl")

    wb = load_workbook(excel_file)
    sheets_data = {}
    for ws in wb.worksheets:
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append(list(row))
        sheets_data[ws.title] = rows
    return sheets_data


# ══════════════════════════════════════════════════════════
# 上傳資料至 Google Sheets
# ══════════════════════════════════════════════════════════

def upload_sheet_data(service, sheet_id, sheet_name, data):
    # 清除現有內容
    range_name = f"'{sheet_name}'"
    service.spreadsheets().values().clear(
        spreadsheetId=sheet_id,
        range=range_name,
    ).execute()

    # 過濾掉全 None 行
    filtered = []
    for row in data:
        if any(v is not None for v in row):
            filtered.append([str(v) if v is not None else "" for v in row])

    if not filtered:
        return

    # 批次上傳（每次最多 50,000 行）
    BATCH = 5000
    for i in range(0, len(filtered), BATCH):
        batch = filtered[i:i + BATCH]
        start_row = i + 1
        range_spec = f"'{sheet_name}'!A{start_row}"
        body = {"values": batch}
        service.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=range_spec,
            valueInputOption="USER_ENTERED",
            body=body,
        ).execute()
        print(f"   上傳第 {i+1}–{i+len(batch)} 行...")

    print(f"✓ '{sheet_name}' 上傳完成（共 {len(filtered)} 行）")


def apply_formatting(service, sheet_id):
    """為各工作表套用基本格式"""
    # 取得工作表 ID
    meta = service.spreadsheets().get(spreadsheetId=sheet_id).execute()
    sheet_ids = {s["properties"]["title"]: s["properties"]["sheetId"]
                 for s in meta["sheets"]}

    header_color = {"red": 0.122, "green": 0.306, "blue": 0.475}  # #1F4E79
    requests = []

    for sheet_name, sid in sheet_ids.items():
        requests.append({
            "repeatCell": {
                "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1},
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": header_color,
                        "textFormat": {"bold": True, "foregroundColor": {"red": 1, "green": 1, "blue": 1}},
                        "horizontalAlignment": "CENTER",
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)",
            }
        })
        requests.append({
            "updateSheetProperties": {
                "properties": {"sheetId": sid, "gridProperties": {"frozenRowCount": 1}},
                "fields": "gridProperties.frozenRowCount",
            }
        })

    if requests:
        service.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={"requests": requests},
        ).execute()
    print("✓ 格式設置完成")


# ══════════════════════════════════════════════════════════
# 主程序
# ══════════════════════════════════════════════════════════

def run_setup():
    print("\n═══════════════════════════════════════════")
    print("  DICJ 博彩統計 — Google Sheets 初始設置")
    print("═══════════════════════════════════════════\n")

    check_dependencies()
    print("✓ 所有依賴已安裝\n")

    print("📡 連接 Google API...")
    service, drive_service = get_service()
    print("✓ 已授權\n")

    sheet_id = get_sheet_id()
    if sheet_id:
        print(f"📊 找到現有試算表 ID：{sheet_id}")
        ans = input("是否使用此試算表？[Y/n] ").strip().lower()
        if ans == "n":
            sheet_id = None

    if not sheet_id:
        print("\n📊 建立新試算表...")
        sheet_id = create_spreadsheet(service)

    print(f"\n📥 讀取 Excel 資料...")
    sheets_data = read_excel_data()
    print(f"✓ 讀取到 {len(sheets_data)} 個工作表\n")

    sheet_mapping = {
        "每月幸運博彩毛收入": "每月幸運博彩毛收入",
        "每季各博彩項目": "每季各博彩項目",
        "年度彙總": "年度彙總",
    }

    print("📤 上傳資料至 Google Sheets...")
    for excel_name, sheets_name in sheet_mapping.items():
        if excel_name in sheets_data:
            print(f"\n  正在上傳「{sheets_name}」...")
            upload_sheet_data(service, sheet_id, sheets_name, sheets_data[excel_name])

    print("\n🎨 套用格式...")
    apply_formatting(service, sheet_id)

    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit"
    print(f"\n{'═'*50}")
    print(f"✅ 設置完成！")
    print(f"   試算表網址：{url}")
    print(f"{'═'*50}")
    print()
    print("下一步：")
    print("  1. 開啟試算表，點擊「擴充功能」→「Apps Script」")
    print("  2. 貼入 dicj_apps_script.gs 的全部內容")
    print("  3. 執行「setupMonthlyTrigger」設定每月自動更新")
    print()


def run_update():
    check_dependencies()
    service, _ = get_service()
    sheet_id = get_sheet_id()
    if not sheet_id:
        print("❌ 找不到試算表 ID，請先執行 --setup")
        sys.exit(1)

    print(f"📥 讀取 Excel 資料...")
    sheets_data = read_excel_data()

    for name in ["每月幸運博彩毛收入", "每季各博彩項目", "年度彙總"]:
        if name in sheets_data:
            print(f"\n  更新「{name}」...")
            upload_sheet_data(service, sheet_id, name, sheets_data[name])

    print("\n✅ 更新完成！")


# ══════════════════════════════════════════════════════════
# 入口
# ══════════════════════════════════════════════════════════

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="DICJ Google Sheets 設置工具")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--setup",  action="store_true", help="首次設置並導入資料")
    group.add_argument("--update", action="store_true", help="更新現有試算表")
    group.add_argument("--info",   action="store_true", help="顯示設置說明")
    args = parser.parse_args()

    if args.setup:
        run_setup()
    elif args.update:
        run_update()
    elif args.info:
        show_info()
    else:
        show_info()
