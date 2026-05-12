"""
自動配對 Word 表格格位 與 Excel 座位號碼
輸出完整對映表供 HTML 使用
"""
import zipfile
import xml.etree.ElementTree as ET
import re
import sys

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
SS_NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

def get_cell_lines(tc):
    lines = []
    for p in tc.findall('.//w:p', NS):
        line = ''.join(r.text or '' for r in p.findall('.//w:t', NS))
        if line.strip():
            lines.append(line.strip())
    return lines

def normalize(lines):
    """合併多行，去除括號內的代理人資訊"""
    text = ''.join(lines)
    text = re.sub(r'[（(][^）)]*[）)]', '', text)
    return text.strip()

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

def parse_docx_cells(path):
    """回傳 {(row, col): lines} 字典（只含非空格）"""
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
                lines = get_cell_lines(tc)
                if lines:
                    cells[(r_idx, col)] = lines
            col += span
    return cells

def parse_xlsx_seats(path):
    """回傳 {seat_num: {'row': excel_row, 'text': normalized_text, 'lines': [b,c,d]}} 字典"""
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            with z.open('xl/sharedStrings.xml') as f:
                ss = ET.parse(f).getroot()
            for si in ss.findall('.//s:si', SS_NS):
                shared.append(''.join(t.text or '' for t in si.findall('.//s:t', SS_NS)))

        sheet_names = sorted(n for n in z.namelist() if n.startswith('xl/worksheets/sheet'))
        with z.open(sheet_names[0]) as f:
            ws = ET.parse(f).getroot()

    def get_val(c):
        t_attr = c.get('t', '')
        v_el = c.find('s:v', SS_NS)
        if v_el is None:
            return ''
        raw = v_el.text or ''
        if t_attr == 's':
            try: return shared[int(raw)]
            except: return raw
        return raw

    rows_data = {}
    for row in ws.findall('.//s:row', SS_NS):
        r = int(row.get('r', 0))
        row_vals = {}
        for c in row.findall('s:c', SS_NS):
            addr = c.get('r', '')
            col_letter = re.match(r'([A-Z]+)', addr).group(1)
            row_vals[col_letter] = get_val(c)
        rows_data[r] = row_vals

    seats = {}
    for r, vals in rows_data.items():
        if r == 1: continue  # header
        seat_raw = vals.get('A', '')
        if not seat_raw.strip():
            continue
        try:
            seat_num = int(seat_raw)
        except ValueError:
            continue
        b = vals.get('B', '').strip()
        c = vals.get('C', '').strip()
        d = vals.get('D', '').strip()
        combined = (b + c + d).strip()
        seats[seat_num] = {
            'excel_row': r,
            'text': combined,
            'b': b, 'c': c, 'd': d,
        }
    return seats

def find_mapping(docx_path, xlsx_path):
    word_cells = parse_docx_cells(docx_path)
    excel_seats = parse_xlsx_seats(xlsx_path)

    # 建立 Excel normalized text → seat_num 的索引
    excel_index = {}
    for seat_num, info in excel_seats.items():
        norm = info['text']
        if norm:
            excel_index[norm] = seat_num

    mapping = {}  # (row, col) → seat_num
    unmatched_word = []
    unmatched_excel = set(excel_seats.keys())

    for (r, c), lines in sorted(word_cells.items()):
        norm_word = normalize(lines)
        if not norm_word:
            continue

        matched = None
        # 嘗試精確匹配
        if norm_word in excel_index:
            matched = excel_index[norm_word]
        else:
            # 嘗試部分匹配（Word 文字包含 Excel 文字，或反之）
            best_score = 0
            for en, sn in excel_index.items():
                if not en:
                    continue
                # 檢查最長公共子字串長度
                shorter = min(len(norm_word), len(en))
                if norm_word in en or en in norm_word:
                    score = shorter
                    if score > best_score:
                        best_score = score
                        matched = sn

        if matched:
            mapping[(r, c)] = matched
            unmatched_excel.discard(matched)
        else:
            unmatched_word.append(((r, c), lines, norm_word))

    return mapping, unmatched_word, unmatched_excel, excel_seats

def main():
    sys.stdout = open('match_output.txt', 'w', encoding='utf-8')
    base = r'C:\Users\y5876\Desktop\VS code\auto_seatingchart\座位表範例'

    rooms = [
        ('範例1- 第1會議室.docx', '範例1- 第1會議室.xlsx', '第1會議室'),
        ('範例2-第3會議室.docx', '範例2-第3會議室.xlsx', '第3會議室'),
    ]

    for docx_name, xlsx_name, room_name in rooms:
        import os
        docx_path = os.path.join(base, docx_name)
        xlsx_path = os.path.join(base, xlsx_name)

        print(f'\n{"="*60}')
        print(f'{room_name} 座位對映表')
        print('='*60)

        mapping, unmatched_word, unmatched_excel, excel_seats = find_mapping(docx_path, xlsx_path)

        print(f'\n✅ 成功配對 {len(mapping)} 個座位:')
        print(f'{"Word 位置":<15} {"座位號":<8} {"Excel行":<8} {"顯示名稱"}')
        print('-'*60)
        for (r, c), seat_num in sorted(mapping.items()):
            info = excel_seats[seat_num]
            print(f'[{r},{c}]{"":8} {seat_num:<8} {info["excel_row"]:<8} {info["text"]}')

        if unmatched_word:
            print(f'\n⚠️  Word 中未配對到 Excel ({len(unmatched_word)} 個):')
            for pos, lines, norm in unmatched_word:
                print(f'  [{pos[0]},{pos[1]}]: {repr(lines)} → 標準化: {repr(norm)}')

        if unmatched_excel:
            print(f'\n⚠️  Excel 中未被配對的座位 ({len(unmatched_excel)} 個):')
            for sn in sorted(unmatched_excel):
                info = excel_seats[sn]
                print(f'  Seat {sn} (row {info["excel_row"]}): {repr(info["text"])}')

        print(f'\n📋 JavaScript 對映表（可直接貼入 HTML）:')
        print(f'const ROOM_{room_name[1]}_MAP = ' + '{')
        for (r, c), seat_num in sorted(mapping.items()):
            info = excel_seats[seat_num]
            print(f'  "{r},{c}": {{"seat": {seat_num}, "excelRow": {info["excel_row"]}}},')
        print('};')

    sys.stdout.close()
    print("完成！請查看 match_output.txt", file=sys.stderr)

if __name__ == '__main__':
    main()
