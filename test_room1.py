"""
第1會議室轉換邏輯端對端測試
用範例 Word（已填寫）模擬轉換，與範例 Excel 逐格比對
"""
import zipfile
import xml.etree.ElementTree as ET
import re
import os
import sys

sys.stdout = open('test_room1_output.txt', 'w', encoding='utf-8')

BASE    = r'C:\Users\y5876\Desktop\VS code\auto_seatingchart\座位表範例'
DOCX    = os.path.join(BASE, '範例1- 第1會議室.docx')
XLSX    = os.path.join(BASE, '範例1- 第1會議室.xlsx')

NS   = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
SS   = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

# ── 與 index.html 完全相同的對映表 ───────────────────────────────────────────
ROOM1_MAP = {
    "21,3": 7,  "21,4": 6,  "21,5": 5,  "21,6": 3,
    "21,7": 2,  "21,8": 4,  "21,9": 8,  "21,10": 9, "21,11": 10,
    "0,12":31,  "1,12":30,  "2,12":29,  "3,12":28,  "4,12":27,
    "5,12":26,  "6,12":25,  "7,12":24,  "8,12":23,  "9,12":22,
    "10,12":21, "11,12":20, "12,12":19, "13,12":18, "14,12":17,
    "15,12":16, "16,12":15, "17,12":14, "18,12":13, "19,12":12, "20,12":11,
    "0,2":52,   "1,2":51,   "2,2":50,   "3,2":49,   "4,2":48,
    "5,2":47,   "6,2":46,   "7,2":45,   "8,2":44,   "9,2":43,
    "10,2":42,  "11,2":41,  "12,2":40,  "13,2":39,  "14,2":38,
    "15,2":37,  "16,2":36,  "17,2":35,  "18,2":34,  "19,2":33,  "20,2":32,
    "6,16":64,  "7,16":63,  "8,16":62,  "9,16":61,  "10,16":60,
    "11,16":59, "12,16":58, "13,16":57, "14,16":56, "15,16":55,
    "16,16":54, "17,16":53,
    "6,0":76,   "7,0":75,   "8,0":74,   "9,0":73,   "10,0":72,
    "11,0":71,  "12,0":70,  "13,0":69,  "14,0":68,  "15,0":67,
    "16,0":66,  "17,0":65,
}

# ── Word 解析（與 match.py 相同邏輯） ────────────────────────────────────────
def get_grid_span(tc):
    tcPr = tc.find('w:tcPr', NS)
    if tcPr is not None:
        gs = tcPr.find('w:gridSpan', NS)
        if gs is not None:
            return int(gs.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 1))
    return 1

def is_vmerge_continue(tc):
    tcPr = tc.find('w:tcPr', NS)
    if tcPr is not None:
        vm = tcPr.find('w:vMerge', NS)
        if vm is not None:
            val = vm.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '')
            return val != 'restart'
    return False

def parse_docx(path):
    with zipfile.ZipFile(path) as z:
        with z.open('word/document.xml') as f:
            root = ET.parse(f).getroot()
    cells = {}
    tbl = root.findall('.//w:tbl', NS)[0]
    for r_idx, tr in enumerate(tbl.findall('w:tr', NS)):
        col = 0
        for tc in tr.findall('w:tc', NS):
            span = get_grid_span(tc)
            if not is_vmerge_continue(tc):
                lines = []
                for p in tc.findall('.//w:p', NS):
                    line = ''.join(t.text or '' for t in p.findall('.//w:t', NS))
                    if line.strip():
                        lines.append(line.strip())
                if lines:
                    cells[f'{r_idx},{col}'] = lines
            col += span
    return cells

# ── 與 index.html processLines() 相同邏輯（v1.4）────────────────────────────
def process_lines(lines):
    if not lines:
        return '', '', ''
    b = lines[0]
    c = ''
    is_proxy = False
    for l in lines[1:]:
        l = l.strip()
        m = re.match(r'^[（(](.+)[）)]$', l)
        if m:
            c = m.group(1).strip()
            is_proxy = True
        elif re.match(r'^[（(]', l):
            c = re.sub(r'^[（(]', '', l).rstrip('）)').strip()
            is_proxy = True
        elif l:
            c = l
            is_proxy = False
    e = b + c if (c and not is_proxy) else b
    return b, c, e

# ── Excel 解析（讀取 B、C、E 欄） ────────────────────────────────────────────
def parse_xlsx(path):
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            with z.open('xl/sharedStrings.xml') as f:
                ss = ET.parse(f).getroot()
            for si in ss.findall('.//s:si', SS):
                shared.append(''.join(t.text or '' for t in si.findall('.//s:t', SS)))
        sheet_files = sorted(n for n in z.namelist() if n.startswith('xl/worksheets/sheet'))
        with z.open(sheet_files[0]) as f:
            ws = ET.parse(f).getroot()

    def get_val(c):
        t_attr = c.get('t', '')
        v_el = c.find('s:v', SS)
        if v_el is None:
            return ''
        raw = v_el.text or ''
        if t_attr == 's':
            try:
                return shared[int(raw)]
            except Exception:
                return raw
        return raw

    rows_data = {}
    for row in ws.findall('.//s:row', SS):
        r = int(row.get('r', 0))
        row_vals = {}
        for c in row.findall('s:c', SS):
            addr = c.get('r', '')
            m = re.match(r'([A-Z]+)', addr)
            if m:
                row_vals[m.group(1)] = get_val(c)
        rows_data[r] = row_vals
    return rows_data

# ── 主測試流程 ───────────────────────────────────────────────────────────────
def main():
    print('=' * 70)
    print('第1會議室 轉換邏輯測試')
    print('=' * 70)

    word_cells = parse_docx(DOCX)
    excel_rows = parse_xlsx(XLSX)

    passed = []
    failed = []
    empty_word = []
    empty_excel = []

    for key, excel_row in sorted(ROOM1_MAP.items(), key=lambda x: x[1]):
        lines = word_cells.get(key)

        # Word 格位空白（範例中本來就空著）
        if not lines:
            excel_b = excel_rows.get(excel_row, {}).get('B', '').strip()
            excel_c = excel_rows.get(excel_row, {}).get('C', '').strip()
            excel_e = excel_rows.get(excel_row, {}).get('E', '').strip()
            empty_word.append((key, excel_row, excel_b + excel_c, excel_e))
            continue

        b, c, e = process_lines(lines)
        tool_display = b + c

        ex = excel_rows.get(excel_row, {})
        excel_b = ex.get('B', '').strip()
        excel_display = excel_b  # 只比對 B 欄（欄位1）

        if not excel_display:
            empty_excel.append((key, excel_row, b + c))
            continue

        match = (b.strip() == excel_display.strip())
        if match:
            passed.append((key, excel_row, tool_display))
        else:
            failed.append((key, excel_row, tool_display, excel_display))

    # ── 輸出結果 ──────────────────────────────────────────────────────────────
    print(f'\n✅ 通過：{len(passed)} 個')
    for key, row, text in passed:
        print(f'   [{key}] → Excel行{row:3d}  {text}')

    if failed:
        print(f'\n❌ 不符：{len(failed)} 個')
        for key, row, got, expected in failed:
            print(f'   [{key}] → Excel行{row:3d}')
            print(f'      工具輸出：{repr(got)}')
            print(f'      Excel原值：{repr(expected)}')

    if empty_word:
        print(f'\n⬜ Word 空白格（範例中未填寫，Excel也空白屬正常）：{len(empty_word)} 個')
        for key, row, excel_val, excel_e in empty_word:
            note = f'  ← Excel有值: {excel_val or excel_e}' if (excel_val or excel_e) else ''
            print(f'   [{key}] → Excel行{row:3d}{note}')

    if empty_excel:
        print(f'\n⚠️  Word有內容但Excel原值空白：{len(empty_excel)} 個')
        for key, row, text in empty_excel:
            print(f'   [{key}] → Excel行{row:3d}  Word: {repr(text)}')

    total = len(ROOM1_MAP)
    print(f'\n{"─"*70}')
    print(f'總計 {total} 個對映格位 | ✅通過 {len(passed)} | ❌不符 {len(failed)} | ⬜Word空白 {len(empty_word)} | ⚠️Excel空白 {len(empty_excel)}')

if __name__ == '__main__':
    main()
    sys.stdout.close()
    print('完成！請查看 test_room1_output.txt', file=sys.stderr)
