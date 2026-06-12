#!/usr/bin/env python3
"""
澳門政府部門通訊錄 - 一鍵更新工具
用法：python3 update.py
自動從 gov.mo 下載最新 XLSX，對比現有資料庫，顯示變更並更新 index.html
"""
import pandas as pd
import re, os, sys, tempfile, shutil
from datetime import date
from urllib.request import urlopen, Request
from io import BytesIO

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_FILE = os.path.join(SCRIPT_DIR, "index.html")
RAWDATA_DIR = os.path.join(SCRIPT_DIR, "Rawdata")
EXCEL_MAIN = os.path.join(RAWDATA_DIR, "msar-apm-entities-contact-zh.xlsx")
EXCEL_COMM = os.path.join(RAWDATA_DIR, "comissoes.xlsx")

URL_MAIN = "https://www.gov.mo/zh-hant/apm-entity-page/download-xlsx/"
URL_COMM = "https://www.gov.mo/Comissoes/Excel.ashx"

def download_file(url, desc):
    """從 URL 下載檔案，回傳 bytes"""
    print(f"   ⬇️  下載{desc}...")
    req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        resp = urlopen(req, timeout=30)
        data = resp.read()
        print(f"      ✓ 已下載 ({len(data)//1024} KB)")
        return data
    except Exception as e:
        print(f"      ✗ 下載失敗：{e}")
        return None

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

def read_excel_data(main_bytes=None, comm_bytes=None):
    """從 Excel 資料讀取所有條目。可接受 bytes 或從檔案讀取"""
    if main_bytes:
        df1 = pd.read_excel(BytesIO(main_bytes), header=None)
    elif os.path.exists(EXCEL_MAIN):
        df1 = pd.read_excel(EXCEL_MAIN, header=None)
    else:
        print(f"❌ 找不到主檔案且無下載資料")
        sys.exit(1)

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

    comm_df = None
    if comm_bytes:
        comm_df = pd.read_excel(BytesIO(comm_bytes), sheet_name=0, header=None)
    elif os.path.exists(EXCEL_COMM):
        comm_df = pd.read_excel(EXCEL_COMM, sheet_name=0, header=None)

    if comm_df is not None:
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

def extract_old_data(html):
    m = re.search(r'const D=\[.*?\];', html, re.DOTALL)
    if not m:
        return None, []
    old_block = m.group(0)
    entries = []
    pattern = r'\{n:"([^"\\]*(?:\\.[^"\\]*)*)",c:"([^"\\]*(?:\\.[^"\\]*)*)",ph:"([^"\\]*(?:\\.[^"\\]*)*)",fx:"([^"\\]*(?:\\.[^"\\]*)*)",em:"([^"\\]*(?:\\.[^"\\]*)*)",ad:"([^"\\]*(?:\\.[^"\\]*)*)",w:"([^"\\]*(?:\\.[^"\\]*)*)",L:\[(.*?)\]\}'
    for match in re.finditer(pattern, old_block):
        leaders = []
        for lm in re.finditer(r'\{r:"([^"]*)",p:"([^"]*)"\}', match.group(8)):
            leaders.append({"r": lm.group(1), "p": lm.group(2)})
        entries.append({
            "n": match.group(1),
            "c": match.group(2),
            "ph": match.group(3),
            "fx": match.group(4),
            "em": match.group(5),
            "ad": match.group(6),
            "w": match.group(7),
            "L": leaders
        })
    return old_block, entries

def compare_data(old_entries, new_entries):
    old_map = {e["n"]: e for e in old_entries}
    new_map = {e["n"]: e for e in new_entries}

    added = [n for n in new_map if n not in old_map]
    removed = [n for n in old_map if n not in new_map]

    leader_changes = []
    detail_changes = []
    field_labels = {"c": "類別", "ph": "電話", "fx": "傳真", "em": "電郵", "ad": "地址", "w": "網址"}

    for name in new_map:
        if name not in old_map:
            continue
        old_e = old_map[name]
        new_e = new_map[name]

        old_leaders = {(l["r"], l["p"]) for l in old_e.get("L", [])}
        new_leaders = {(l["r"], l["p"]) for l in new_e.get("L", [])}
        if old_leaders != new_leaders:
            removed_l = old_leaders - new_leaders
            added_l = new_leaders - old_leaders
            leader_changes.append({"name": name, "removed": removed_l, "added": added_l})

        changed_fields = []
        for field, label in field_labels.items():
            old_val = old_e.get(field, "")
            new_val = js_escape(new_e.get(field, ""))
            if old_val != new_val:
                changed_fields.append(label)
        if changed_fields:
            detail_changes.append({"name": name, "fields": changed_fields})

    return added, removed, leader_changes, detail_changes

def main():
    print("=" * 60)
    print("  澳門政府部門通訊錄 - 一鍵更新工具")
    print("=" * 60)

    # 1. Download from gov.mo
    print("\n🌐 從 gov.mo 下載最新資料...")
    main_bytes = download_file(URL_MAIN, "政府部門通訊錄")
    comm_bytes = download_file(URL_COMM, "諮詢組織通訊錄")

    if not main_bytes:
        print("\n❌ 主檔案下載失敗，無法繼續。")
        sys.exit(1)

    # 2. Parse new data
    print("\n📖 解析下載資料...")
    new_entries = read_excel_data(main_bytes, comm_bytes)
    new_leader_count = sum(len(e["L"]) for e in new_entries)
    print(f"   線上資料：{len(new_entries)} 個部門/機構，{new_leader_count} 條領導人記錄")

    # 3. Read old data from HTML
    if not os.path.exists(HTML_FILE):
        print(f"❌ 找不到 HTML 檔案：{HTML_FILE}")
        sys.exit(1)

    with open(HTML_FILE, "r", encoding="utf-8") as f:
        html = f.read()

    old_block, old_entries = extract_old_data(html)
    if old_block is None:
        print("⚠️  無法從現有 HTML 中提取舊資料，將直接寫入新資料")
    else:
        old_leader_count = sum(len(e["L"]) for e in old_entries)
        print(f"   現有資料：{len(old_entries)} 個部門/機構，{old_leader_count} 條領導人記錄")

    # 4. Compare — 以完整資料對比，不只比較數目
    new_js = entries_to_js(new_entries)

    if old_block and new_js == old_block:
        print("\n✅ 線上資料與現有資料完全一致，無需更新。")
        return

    if old_entries:
        added, removed, leader_changes, detail_changes = compare_data(old_entries, new_entries)

        print("\n" + "─" * 60)
        print("  📋 變更摘要")
        print("─" * 60)

        if added:
            print(f"\n  🆕 新增 {len(added)} 個部門/機構：")
            for n in added:
                print(f"     + {n}")

        if removed:
            print(f"\n  🗑️  移除 {len(removed)} 個部門/機構：")
            for n in removed:
                print(f"     - {n}")

        if leader_changes:
            print(f"\n  👤 {len(leader_changes)} 個部門領導人變更：")
            for ch in leader_changes:
                print(f"     📌 {ch['name']}")
                for r, p in ch["removed"]:
                    print(f"        - {r}：{p}")
                for r, p in ch["added"]:
                    print(f"        + {r}：{p}")

        if detail_changes:
            print(f"\n  📝 {len(detail_changes)} 個部門其他資料變更：")
            for ch in detail_changes:
                fields_str = "、".join(ch["fields"])
                print(f"     📌 {ch['name']}（{fields_str}）")

        print("\n" + "─" * 60)
        total_changes = len(added) + len(removed) + len(leader_changes) + len(detail_changes)
        print(f"  共 {total_changes} 項變更")
        print("─" * 60)
    else:
        print("\n⚠️  無舊資料可對比，將直接寫入")

    # 5. Save downloaded files to Rawdata/
    os.makedirs(RAWDATA_DIR, exist_ok=True)
    with open(EXCEL_MAIN, "wb") as f:
        f.write(main_bytes)
    print(f"   💾 已儲存：{EXCEL_MAIN}")
    if comm_bytes:
        with open(EXCEL_COMM, "wb") as f:
            f.write(comm_bytes)
        print(f"   💾 已儲存：{EXCEL_COMM}")

    # 7. Update HTML (new_js already generated in step 4)
    today = date.today().isoformat()

    if old_block:
        html = html.replace(old_block, new_js)
    else:
        html = re.sub(r'const D=\[.*?\];', new_js, html, flags=re.DOTALL)

    html = re.sub(
        r'資料庫更新日期：\d{4}-\d{2}-\d{2}',
        f'資料庫更新日期：{today}',
        html
    )

    with open(HTML_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    tmp_copy = "/tmp/macau-dir/index.html"
    if os.path.exists(os.path.dirname(tmp_copy)):
        with open(tmp_copy, "w", encoding="utf-8") as f:
            f.write(html)

    print(f"\n✅ 更新完成！")
    print(f"   部門/機構：{len(new_entries)}")
    print(f"   領導人記錄：{new_leader_count}")
    print(f"   更新日期：{today}")
    print(f"   檔案：{HTML_FILE}")

if __name__ == "__main__":
    main()
