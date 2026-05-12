"""
分析範例 Word 與 Excel 檔案，印出表格結構以建立座位對映表
執行方式：python analyze.py
"""
import zipfile
import xml.etree.ElementTree as ET
import os

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

def get_cell_text(tc):
    """取得一個 <w:tc> 內的所有文字（跨 paragraph/run）"""
    parts = []
    for p in tc.findall('.//w:p', NS):
        line = ''.join(r.text or '' for r in p.findall('.//w:t', NS))
        if line:
            parts.append(line)
    return '\n'.join(parts)

def get_grid_span(tc):
    """取得欄合併數，預設 1"""
    tcPr = tc.find('w:tcPr', NS)
    if tcPr is not None:
        gs = tcPr.find('w:gridSpan', NS)
        if gs is not None:
            return int(gs.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 1))
    return 1

def is_vmerge_continue(tc):
    """判斷是否為垂直合併的延續格（不是起始格）"""
    tcPr = tc.find('w:tcPr', NS)
    if tcPr is not None:
        vm = tcPr.find('w:vMerge', NS)
        if vm is not None:
            val = vm.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '')
            return val != 'restart'
    return False

def analyze_docx(path):
    print(f"\n{'='*60}")
    print(f"Word 檔案：{os.path.basename(path)}")
    print('='*60)
    with zipfile.ZipFile(path) as z:
        with z.open('word/document.xml') as f:
            tree = ET.parse(f)
    root = tree.getroot()

    tables = root.findall('.//w:tbl', NS)
    print(f"共找到 {len(tables)} 個表格\n")

    for tbl_idx, tbl in enumerate(tables):
        rows = tbl.findall('w:tr', NS)
        print(f"--- 表格 {tbl_idx+1}（共 {len(rows)} 列）---")
        grid = []
        for r_idx, tr in enumerate(rows):
            cells = tr.findall('w:tc', NS)
            row_data = []
            col = 0
            for tc in cells:
                span = get_grid_span(tc)
                vmerge_cont = is_vmerge_continue(tc)
                text = get_cell_text(tc)
                row_data.append({
                    'col': col,
                    'span': span,
                    'vmerge_cont': vmerge_cont,
                    'text': text,
                })
                col += span
            grid.append(row_data)
            # 印出每格內容
            for cell in row_data:
                display = repr(cell['text']) if cell['text'] else '(空)'
                flag = '[vmerge↓]' if cell['vmerge_cont'] else ''
                span_str = f'[span={cell["span"]}]' if cell['span'] > 1 else ''
                print(f"  [{r_idx},{cell['col']}]{span_str}{flag} {display}")
        print()

def analyze_xlsx(path):
    print(f"\n{'='*60}")
    print(f"Excel 檔案：{os.path.basename(path)}")
    print('='*60)
    with zipfile.ZipFile(path) as z:
        # 列出所有 sheet
        names = [n for n in z.namelist() if n.startswith('xl/worksheets/sheet')]
        print(f"Sheets: {names}")

        # 讀取 sharedStrings（若有）
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            with z.open('xl/sharedStrings.xml') as f:
                ss_tree = ET.parse(f)
            ss_ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for si in ss_tree.findall('.//s:si', ss_ns):
                t_texts = [t.text or '' for t in si.findall('.//s:t', ss_ns)]
                shared.append(''.join(t_texts))

        # 分析第一個 sheet
        sheet_path = names[0]
        with z.open(sheet_path) as f:
            ws_tree = ET.parse(f)
        ws_ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        ws_root = ws_tree.getroot()

        print(f"\n分析 {sheet_path} 中有文字內容的儲存格：")
        rows = ws_root.findall('.//s:row', ws_ns)
        for row in rows:
            for c in row.findall('s:c', ws_ns):
                addr = c.get('r')
                t_attr = c.get('t', '')
                v_el = c.find('s:v', ws_ns)
                if v_el is None:
                    continue
                raw = v_el.text or ''
                if t_attr == 's':
                    # shared string
                    try:
                        val = shared[int(raw)]
                    except (IndexError, ValueError):
                        val = raw
                elif t_attr == 'inlineStr':
                    is_el = c.find('.//s:t', ws_ns)
                    val = is_el.text if is_el is not None else ''
                else:
                    val = raw
                if val.strip():
                    print(f"  {addr}: {repr(val)}")

if __name__ == '__main__':
    import sys
    # 強制輸出為 UTF-8，避免 Windows cp950 錯誤
    sys.stdout = open('analyze_output.txt', 'w', encoding='utf-8')

    base = r'C:\Users\y5876\Desktop\VS code\auto_seatingchart\座位表範例'
    files = [
        ('範例1- 第1會議室.docx', '範例1- 第1會議室.xlsx'),
        ('範例2-第3會議室.docx', '範例2-第3會議室.xlsx'),
    ]
    for docx_name, xlsx_name in files:
        analyze_docx(os.path.join(base, docx_name))
        analyze_xlsx(os.path.join(base, xlsx_name))

    sys.stdout.close()
    print("完成！請查看 analyze_output.txt", file=sys.stderr)
