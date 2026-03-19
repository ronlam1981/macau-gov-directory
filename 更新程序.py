#!/usr/bin/env python3
"""
澳門政府部門通訊錄 - 更新程序
用法：雙擊「更新程序.command」或執行 python3 更新程序.py
從 Rawdata/ 內的 Excel 檔案重新生成 index.html 的資料庫
（不會從網上下載，只使用本地檔案）
"""
import pandas as pd
import re, os, sys
from datetime import date
from io import BytesIO

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_FILE = os.path.join(SCRIPT_DIR, "index.html")
RAWDATA_DIR = os.path.join(SCRIPT_DIR, "Rawdata")
EXCEL_MAIN = os.path.join(RAWDATA_DIR, "msar-apm-entities-contact-zh.xlsx")
EXCEL_COMM = os.path.join(RAWDATA_DIR, "comissoes.xlsx")

def get_cat(i):
    if i <= 20: return "行政長官"
    if i <= 22: return "行政會"
    if i <= 43: return "行政法務司"
    if i <= 62: return "經濟財政司"
    if i <= 81: return "保安司"
    if i <= 199: return "社會文化司"
    if i <= 223: return "運輸工務司"
    if i <= 228: return "廉政公署"
    if i == 229: return "審計署"
    if i <= 231: return "立法機關"
    if i <= 241: return "功能組織"
    if i <= 264: return "司法機關"
    return "公共事業"

def parse_leaders(text):
    if pd.isna(text) or not str(text).strip():
        return []
    leaders = []
    for line in str(text).strip().split('\n'):
        line = line.strip()
        if not line or '---' in line:
            continue
        if '：' in line:
            parts = line.split('：', 1)
            role = parts[0].strip()
            person = parts[1].strip()
            if person and person != '---':
                leaders.append({"r": role, "p": person})
    return leaders

def clean_str(val):
    if pd.isna(val):
        return ""
    s = str(val).strip()
    if s in ('---', 'nan'):
        return ""
    return s

def js_escape(s):
    return s.replace('\\', '\\\\').replace('"', '\\"').replace('\n', '\\n')

def read_excel_data():
    """從本地 Rawdata/ Excel 檔案讀取所有條目"""
    if not os.path.exists(EXCEL_MAIN):
        print(f"❌ 找不到主檔案：{EXCEL_MAIN}")
        sys.exit(1)

    df1 = pd.read_excel(EXCEL_MAIN, header=None)
    entries = []
    for i in range(len(df1)):
        name = clean_str(df1.iloc[i, 0])
        if not name:
            continue
        entries.append({
            "n": name,
            "c": get_cat(i),
            "ad": clean_str(df1.iloc[i, 1]),
            "ph": clean_str(df1.iloc[i, 2]),
            "fx": clean_str(df1.iloc[i, 3]),
            "em": clean_str(df1.iloc[i, 4]),
            "w": clean_str(df1.iloc[i, 5]),
            "L": parse_leaders(df1.iloc[i, 7])
        })

    if os.path.exists(EXCEL_COMM):
        comm_df = pd.read_excel(EXCEL_COMM, sheet_name=0, header=None)
        existing_names = {e["n"] for e in entries}
        for i in range(len(comm_df)):
            name = clean_str(comm_df.iloc[i, 0])
            if not name or name in existing_names:
                continue
            entries.append({
                "n": name,
                "c": "諮詢組織",
                "ad": clean_str(comm_df.iloc[i, 1]),
                "ph": clean_str(comm_df.iloc[i, 2]),
                "fx": "",
                "em": clean_str(comm_df.iloc[i, 3]),
                "w": clean_str(comm_df.iloc[i, 4]),
                "L": []
            })
    else:
        print(f"⚠️  找不到諮詢組織檔案：{EXCEL_COMM}，跳過")

    return entries

def entries_to_js(entries):
    js_lines = []
    for e in entries:
        leaders_parts = []
        for l in e["L"]:
            leaders_parts.append('{r:"' + js_escape(l["r"]) + '",p:"' + js_escape(l["p"]) + '"}')
        leaders_js = "[" + ",".join(leaders_parts) + "]"
        line = '{n:"' + js_escape(e["n"]) + '",c:"' + js_escape(e["c"]) + '",ph:"' + js_escape(e["ph"]) + '",fx:"' + js_escape(e["fx"]) + '",em:"' + js_escape(e["em"]) + '",ad:"' + js_escape(e["ad"]) + '",w:"' + js_escape(e["w"]) + '",L:' + leaders_js + '}'
        js_lines.append(line)
    return "const D=[\n" + ",\n".join(js_lines) + "\n];"

def main():
    print("=" * 60)
    print("  澳門政府部門通訊錄 - 更新程序（從本地資料重建）")
    print("=" * 60)

    # 1. Read local Excel files
    print("\n📖 從 Rawdata/ 讀取本地 Excel 資料...")
    entries = read_excel_data()
    leader_count = sum(len(e["L"]) for e in entries)
    print(f"   讀取到：{len(entries)} 個部門/機構，{leader_count} 條領導人記錄")

    # 2. Read current HTML
    if not os.path.exists(HTML_FILE):
        print(f"❌ 找不到 HTML 檔案：{HTML_FILE}")
        sys.exit(1)

    with open(HTML_FILE, "r", encoding="utf-8") as f:
        html = f.read()

    # 3. Generate new JS and update HTML
    new_js = entries_to_js(entries)
    today = date.today().isoformat()

    old_match = re.search(r'const D=\[.*?\];', html, re.DOTALL)
    if old_match:
        html = html.replace(old_match.group(0), new_js)
    else:
        print("⚠️  無法找到舊資料區塊，請檢查 HTML 結構")
        sys.exit(1)

    html = re.sub(
        r'資料庫更新日期：\d{4}-\d{2}-\d{2}',
        f'資料庫更新日期：{today}',
        html
    )

    # 4. Confirm
    reply = input(f"\n將使用本地 Rawdata/ 重建資料庫（{len(entries)} 個部門）。確認？(y/n): ").strip().lower()
    if reply not in ('y', 'yes', ''):
        print("❌ 已取消。")
        return

    # 5. Write
    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\n✅ 程序更新完成！")
    print(f"   部門/機構：{len(entries)}")
    print(f"   領導人記錄：{leader_count}")
    print(f"   更新日期：{today}")
    print(f"   檔案：{HTML_FILE}")

if __name__ == "__main__":
    main()
