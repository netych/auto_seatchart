"""
產生 index.html：讀取 Excel/Word 範本轉 base64，並嵌入完整 HTML/JS 邏輯
"""
import base64
import io
import os
import zipfile
import xml.etree.ElementTree as ET

BASE = r'C:\Users\y5876\Desktop\VS code\auto_seatingchart'
SAMPLE_DIR = os.path.join(BASE, '座位表範例')
ROOM1_XLSX = os.path.join(SAMPLE_DIR, '範例1- 第1會議室.xlsx')
ROOM3_XLSX = os.path.join(SAMPLE_DIR, '範例2-第3會議室.xlsx')
ROOM1_DOCX = os.path.join(SAMPLE_DIR, '範例1- 第1會議室.docx')
ROOM3_DOCX = os.path.join(SAMPLE_DIR, '範例2-第3會議室.docx')

W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
ET.register_namespace('w', W_NS)

def to_b64(path):
    with open(path, 'rb') as f:
        return base64.b64encode(f.read()).decode('ascii')

def make_blank_docx_b64(src_path, decolor_cols=None):
    """讀取原始 DOCX，清空所有表格儲存格的文字，回傳 base64 字串"""
    buf = io.BytesIO()
    with zipfile.ZipFile(src_path, 'r') as zin, zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == 'word/document.xml':
                data = _clear_table_text(data, decolor_cols=decolor_cols)
            zout.writestr(item, data)
    return base64.b64encode(buf.getvalue()).decode('ascii')

def _clear_table_text(xml_bytes, decolor_cols=None):
    """解析 document.xml，清空所有 <w:tc> 內的文字 run，每格保留一個空 <w:p>。
    decolor_cols: set of col indices whose <w:shd> fill should be removed."""
    xml_str = xml_bytes.decode('utf-8')
    decl = ''
    if xml_str.startswith('<?xml'):
        decl = xml_str[:xml_str.index('?>') + 2] + '\n'

    root = ET.fromstring(xml_bytes)
    tbl = root.findall(f'.//{{{W_NS}}}tbl')[0]

    for tr in tbl.findall(f'{{{W_NS}}}tr'):
        col = 0
        for tc in tr.findall(f'{{{W_NS}}}tc'):
            # 計算 gridSpan
            tcPr = tc.find(f'{{{W_NS}}}tcPr')
            span = 1
            if tcPr is not None:
                gs = tcPr.find(f'{{{W_NS}}}gridSpan')
                if gs is not None:
                    span = int(gs.get(f'{{{W_NS}}}val', 1))
                # 移除指定欄的填色
                if decolor_cols and col in decolor_cols:
                    shd = tcPr.find(f'{{{W_NS}}}shd')
                    if shd is not None:
                        tcPr.remove(shd)

            # 清空文字 run
            paras = list(tc.findall(f'{{{W_NS}}}p'))
            for p in paras:
                for r in list(p.findall(f'{{{W_NS}}}r')):
                    p.remove(r)
                for tag in ['hyperlink', 'ins', 'del']:
                    for el in list(p.findall(f'{{{W_NS}}}{tag}')):
                        p.remove(el)
            for p in paras[1:]:
                tc.remove(p)

            col += span

    new_xml = ET.tostring(root, encoding='unicode', xml_declaration=False)
    return (decl + new_xml).encode('utf-8')

# 生成資料
room1_xlsx_b64   = to_b64(ROOM1_XLSX)
room3_xlsx_b64   = to_b64(ROOM3_XLSX)
room1_docx_b64   = make_blank_docx_b64(ROOM1_DOCX, decolor_cols={12})
room3_docx_b64   = make_blank_docx_b64(ROOM3_DOCX)
print('Excel 範本和空白 Word 範本已準備好')

# HTML template - use __ROOM1_B64__ and __ROOM3_B64__ as placeholders
HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>座位表自動轉換工具</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: "Microsoft JhengHei", Arial, sans-serif; background: #f5f7fa; margin: 0; padding: 20px; }
  .container { max-width: 640px; margin: 40px auto; background: #fff; border-radius: 12px; box-shadow: 0 4px 20px rgba(0,0,0,.1); padding: 36px; }
  h1 { color: #1a3a5c; font-size: 1.5rem; margin: 0 0 8px; }
  .subtitle { color: #666; font-size: .9rem; margin: 0 0 28px; }
  .section { margin-bottom: 24px; }
  label { display: block; font-weight: 600; color: #333; margin-bottom: 8px; font-size: .95rem; }
  .radio-group { display: flex; gap: 16px; }
  .radio-group label { font-weight: normal; display: flex; align-items: center; gap: 6px; cursor: pointer; }
  .file-input-wrap { border: 2px dashed #b0c4de; border-radius: 8px; padding: 20px; text-align: center; cursor: pointer; transition: border-color .2s, background .2s; }
  .file-input-wrap:hover { border-color: #1a6fc4; }
  .file-input-wrap.drag-over { border-color: #1a6fc4; background: #e8f2ff; }
  .file-input-wrap input { display: none; }
  .file-name { margin-top: 8px; color: #555; font-size: .9rem; }
  .btn { display: block; width: 100%; padding: 14px; background: #1a6fc4; color: #fff; border: none; border-radius: 8px; font-size: 1.05rem; font-weight: 600; cursor: pointer; transition: background .2s; }
  .btn:hover:not(:disabled) { background: #155ba0; }
  .btn:disabled { background: #90b8e0; cursor: not-allowed; }
  .log { margin-top: 20px; background: #f0f4f8; border-radius: 8px; padding: 14px; font-size: .85rem; color: #444; max-height: 260px; overflow-y: auto; white-space: pre-wrap; display: none; }
  .badge { display: inline-block; background: #e3f0ff; color: #1a6fc4; border-radius: 4px; padding: 2px 8px; font-size: .8rem; font-weight: 600; margin-left: 8px; }
  .btn-outline { display: block; width: 100%; padding: 11px; background: #fff; color: #1a6fc4; border: 2px solid #1a6fc4; border-radius: 8px; font-size: .95rem; font-weight: 600; cursor: pointer; transition: background .2s, color .2s; margin-bottom: 10px; }
  .btn-outline:hover { background: #e3f0ff; }
  .divider { border: none; border-top: 1px solid #e8edf2; margin: 20px 0; }
</style>
</head>
<body>
<div class="container">
  <h1>座位表自動轉換工具</h1>
  <p class="subtitle">上傳 Word 座位表 → 自動填入 Excel 表單並下載</p>

  <div class="section">
    <label>選擇會議室</label>
    <div class="radio-group">
      <label><input type="radio" name="room" value="1" checked> 第1會議室</label>
      <label><input type="radio" name="room" value="3"> 第3會議室</label>
    </div>
  </div>

  <hr class="divider">

  <div class="section">
    <label>步驟 1：下載空白 Word 範本</label>
    <button class="btn-outline" onclick="downloadBlankDocx()">
      ⬇ 下載空白座位表範本 (.docx)
    </button>
    <div style="color:#888;font-size:.82rem;margin-top:4px;">填入出席者資訊後，再上傳轉換</div>
  </div>

  <div class="section">
    <label>步驟 2：上傳填寫完成的 Word 座位表 <span class="badge">.docx</span></label>
    <div class="file-input-wrap" id="dropZone" onclick="document.getElementById('docxFile').click()">
      <svg width="36" height="36" viewBox="0 0 24 24" fill="none" stroke="#90a8c8" stroke-width="1.5">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
        <polyline points="17 8 12 3 7 8"/>
        <line x1="12" y1="3" x2="12" y2="15"/>
      </svg>
      <div>點擊或拖曳 .docx 檔案至此</div>
      <div class="file-name" id="fileName">尚未選擇檔案</div>
      <input type="file" id="docxFile" accept=".docx">
    </div>
  </div>

  <button class="btn" id="convertBtn" disabled onclick="convert()">步驟 3：產生 Excel 並下載</button>
  <div class="log" id="log"></div>
</div>

<!-- JSZip (MIT) -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js"></script>
<!-- SheetJS Community Edition (Apache 2.0) -->
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>

<script>
// ── Excel 範本 (base64 嵌入) ─────────────────────────────────────────────────
const TEMPLATES = {
  '1': '__ROOM1_XLSX_B64__',
  '3': '__ROOM3_XLSX_B64__',
};

// ── 空白 Word 範本 (base64 嵌入) ─────────────────────────────────────────────
const BLANK_DOCX = {
  '1': '__ROOM1_DOCX_B64__',
  '3': '__ROOM3_DOCX_B64__',
};

// ── 座位對映表：Word [row,col] → Excel Sheet1 的行號 ────────────────────────
const ROOM1_MAP = {
  // 主桌底列 (Word 第 21 列)
  "21,3": 7,  "21,4": 6,  "21,5": 5,  "21,6": 3,
  "21,7": 2,  "21,8": 4,  "21,9": 8,  "21,10": 9, "21,11": 10,
  // 右側賓客 (第 12 欄，上→下，Excel 行 31→11)
  "0,12":31,  "1,12":30,  "2,12":29,  "3,12":28,  "4,12":27,
  "5,12":26,  "6,12":25,  "7,12":24,  "8,12":23,  "9,12":22,
  "10,12":21, "11,12":20, "12,12":19, "13,12":18, "14,12":17,
  "15,12":16, "16,12":15, "17,12":14, "18,12":13, "19,12":12, "20,12":11,
  // 左側人員 (第 2 欄，上→下，Excel 行 52→32)
  "0,2":52,   "1,2":51,   "2,2":50,   "3,2":49,   "4,2":48,
  "5,2":47,   "6,2":46,   "7,2":45,   "8,2":44,   "9,2":43,
  "10,2":42,  "11,2":41,  "12,2":40,  "13,2":39,  "14,2":38,
  "15,2":37,  "16,2":36,  "17,2":35,  "18,2":34,  "19,2":33,  "20,2":32,
  // 右外側 (第 16 欄，Excel 行 64→53)
  "6,16":64,  "7,16":63,  "8,16":62,  "9,16":61,  "10,16":60,
  "11,16":59, "12,16":58, "13,16":57, "14,16":56, "15,16":55,
  "16,16":54, "17,16":53,
  // 左外側 (第 0 欄，Excel 行 76→65)
  "6,0":76,   "7,0":75,   "8,0":74,   "9,0":73,   "10,0":72,
  "11,0":71,  "12,0":70,  "13,0":69,  "14,0":68,  "15,0":67,
  "16,0":66,  "17,0":65,
};

const ROOM3_MAP = {
  // 主桌 (Word 第 7 列)
  "7,0": 3,   "7,5": 2,   "7,12": 4,
  // 內右 (第 13 欄)
  "1,13": 8,  "2,13": 7,  "3,13": 6,  "4,13": 5,
  // 右區 (第 20 欄)
  "1,20": 12, "2,20": 11, "3,20": 10, "4,20": 9,
  // 左內 (第 8 欄，Excel 行 18→13)
  "0,8": 18,  "1,8": 17,  "2,8": 16,  "3,8": 15,  "4,8": 14,  "5,8": 13,
  // 右外 (第 25 欄，Excel 行 24→19)
  "0,25": 24, "1,25": 23, "2,25": 22, "3,25": 21, "4,25": 20, "5,25": 19,
  // 左區 (第 4 欄，Excel 行 30→25)
  "0,4": 30,  "1,4": 29,  "2,4": 28,  "3,4": 27,  "4,4": 26,  "5,4": 25,
  // 最左外 (第 30 欄，Excel 行 36→31)
  "0,30": 36, "1,30": 35, "2,30": 34, "3,30": 33, "4,30": 32, "5,30": 31,
  // 最外側 (第 0 欄，Excel 行 42→37)
  "0,0": 42,  "1,0": 41,  "2,0": 40,  "3,0": 39,  "4,0": 38,  "5,0": 37,
  // 最右側簡報席 (第 34 欄，Excel 行 47→43，共 5 席)
  "1,34": 47, "2,34": 46, "3,34": 45, "4,34": 44, "5,34": 43,
};

const MAPS = { '1': ROOM1_MAP, '3': ROOM3_MAP };

// ── 下載空白 Word 範本 ───────────────────────────────────────────────────────
function downloadBlankDocx() {
  const room = document.querySelector('input[name="room"]:checked').value;
  const roomName = room === '1' ? '第1會議室' : '第3會議室';
  const b64 = BLANK_DOCX[room];
  // base64 → Uint8Array
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = '座位表空白範本_' + roomName + '.docx';
  a.click();
  URL.revokeObjectURL(a.href);
}

// ── UI 事件 ──────────────────────────────────────────────────────────────────
function setFile(file) {
  if (!file) return;
  if (!file.name.toLowerCase().endsWith('.docx')) {
    alert('請選擇 .docx 格式的 Word 檔案');
    return;
  }
  // 把 File 塞進 input（供 convert() 讀取）
  const dt = new DataTransfer();
  dt.items.add(file);
  document.getElementById('docxFile').files = dt.files;
  document.getElementById('fileName').textContent = file.name;
  document.getElementById('convertBtn').disabled = false;
}

document.getElementById('docxFile').addEventListener('change', function() {
  const name = this.files[0] ? this.files[0].name : '尚未選擇檔案';
  document.getElementById('fileName').textContent = name;
  document.getElementById('convertBtn').disabled = !this.files[0];
});

(function() {
  const zone = document.getElementById('dropZone');
  zone.addEventListener('dragover', function(e) {
    e.preventDefault();
    zone.classList.add('drag-over');
  });
  zone.addEventListener('dragleave', function(e) {
    if (!zone.contains(e.relatedTarget)) zone.classList.remove('drag-over');
  });
  zone.addEventListener('drop', function(e) {
    e.preventDefault();
    zone.classList.remove('drag-over');
    const file = e.dataTransfer.files[0];
    setFile(file);
  });
})();

function logMsg(msg) {
  const el = document.getElementById('log');
  el.style.display = 'block';
  el.textContent += msg + '\n';
  el.scrollTop = el.scrollHeight;
}

// ── DOCX 解析 ────────────────────────────────────────────────────────────────
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function getCellLines(tc) {
  const lines = [];
  const paras = tc.getElementsByTagNameNS(W_NS, 'p');
  for (let i = 0; i < paras.length; i++) {
    const texts = paras[i].getElementsByTagNameNS(W_NS, 't');
    let line = '';
    for (let j = 0; j < texts.length; j++) line += texts[j].textContent;
    if (line.trim()) lines.push(line.trim());
  }
  return lines;
}

function getGridSpan(tc) {
  const tcPr = tc.getElementsByTagNameNS(W_NS, 'tcPr')[0];
  if (tcPr) {
    const gs = tcPr.getElementsByTagNameNS(W_NS, 'gridSpan')[0];
    if (gs) {
      const val = gs.getAttributeNS(W_NS, 'val') || gs.getAttribute('w:val') || '1';
      return parseInt(val);
    }
  }
  return 1;
}

function isVMergeContinue(tc) {
  const tcPr = tc.getElementsByTagNameNS(W_NS, 'tcPr')[0];
  if (tcPr) {
    const vm = tcPr.getElementsByTagNameNS(W_NS, 'vMerge')[0];
    if (vm) {
      const val = vm.getAttributeNS(W_NS, 'val') || vm.getAttribute('w:val') || '';
      return val !== 'restart';
    }
  }
  return false;
}

function parseDocx(xmlText) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(xmlText, 'application/xml');
  const tbls = doc.getElementsByTagNameNS(W_NS, 'tbl');
  if (!tbls.length) return {};

  // 取第一個表格
  const tbl = tbls[0];
  const cells = {};
  const rows = tbl.childNodes;
  let r = 0;
  for (let ri = 0; ri < rows.length; ri++) {
    const row = rows[ri];
    if (row.localName !== 'tr') continue;
    let col = 0;
    for (let ci = 0; ci < row.childNodes.length; ci++) {
      const tc = row.childNodes[ci];
      if (tc.localName !== 'tc') continue;
      const span = getGridSpan(tc);
      if (!isVMergeContinue(tc)) {
        const lines = getCellLines(tc);
        if (lines.length) cells[r + ',' + col] = lines;
      }
      col += span;
    }
    r++;
  }
  return cells;
}

// ── 文字處理：出席委員 → B欄，代理人去括號 → C欄 ───────────────────────────
function processLines(lines) {
  if (!lines.length) return { b: '', c: '' };
  // 第1行：出席委員
  const b = lines[0];
  // 其餘行：找括號行，去除括號後作為代理人；非括號行串接到 b 後（多行正文）
  let c = '';
  for (let i = 1; i < lines.length; i++) {
    const l = lines[i].trim();
    const m = l.match(/^[（(](.+)[）)]$/);
    if (m) {
      c = m[1].trim();  // 去括號，代理人姓名
    }
    // 括號不完整（只開頭有括號但不成對）也嘗試去頭括號
    else if (/^[（(]/.test(l)) {
      c = l.replace(/^[（(]/, '').replace(/[）)]$/, '').trim();
    }
    // 非括號行：屬於正文第二行，視為 b 的延續（與舊邏輯相同）
  }
  return { b: b, c: c };
}

// ── 主流程 ───────────────────────────────────────────────────────────────────
async function convert() {
  const btn = document.getElementById('convertBtn');
  const logEl = document.getElementById('log');
  logEl.textContent = '';
  logEl.style.display = 'block';
  btn.disabled = true;
  btn.textContent = '處理中…';

  try {
    const room = document.querySelector('input[name="room"]:checked').value;
    const file = document.getElementById('docxFile').files[0];
    const map = MAPS[room];

    // 1. 解析 DOCX
    logMsg('📄 解析 Word 檔案中…');
    const docxBuf = await file.arrayBuffer();
    const zip = await JSZip.loadAsync(docxBuf);
    const xmlText = await zip.file('word/document.xml').async('text');
    const cells = parseDocx(xmlText);
    logMsg('   取得 ' + Object.keys(cells).length + ' 個非空格位');

    // 2. 載入 Excel 範本
    logMsg('📊 載入 Excel 範本中…');
    const wb = XLSX.read(TEMPLATES[room], { type: 'base64' });
    const ws = wb.Sheets[wb.SheetNames[0]];

    // 3. 清空 B、C、D、E、F 欄（保留 A 欄座位號碼）
    const toClear = [];
    for (const addr in ws) {
      if (addr[0] === '!') continue;
      const m = addr.match(/^([A-Z]+)(\d+)$/);
      if (!m) continue;
      if (['B','C','D','E'].includes(m[1]) && parseInt(m[2]) > 1) {
        toClear.push(addr);
      }
    }
    toClear.forEach(function(a) { delete ws[a]; });

    // 4. 依對映表填入資料
    let filled = 0, empty = 0;
    for (const key in map) {
      const excelRow = map[key];
      const lines = cells[key];
      if (!lines) { empty++; continue; }

      const processed = processLines(lines);
      if (!processed.b) { empty++; continue; }

      ws['B' + excelRow] = { v: processed.b, t: 's' };
      if (processed.c) ws['C' + excelRow] = { v: processed.c, t: 's' };
      ws['E' + excelRow] = { v: processed.b, t: 's' };

      logMsg('   [' + key + '] → 行' + excelRow + ': ' + processed.b + (processed.c ? ' / ' + processed.c : ''));
      filled++;
    }

    logMsg('\n✅ 完成！填入 ' + filled + ' 個座位，' + empty + ' 個無資料');

    // 5. 下載
    const roomName = room === '1' ? '第1會議室' : '第3會議室';
    const fileName = '座位表_' + roomName + '_已填妥.xlsx';
    XLSX.writeFile(wb, fileName);
    logMsg('💾 已下載：' + fileName);

  } catch (e) {
    logMsg('❌ 錯誤：' + e.message);
    console.error(e);
  } finally {
    btn.disabled = false;
    btn.textContent = '產生 Excel 並下載';
  }
}
</script>
</body>
</html>'''

html = (HTML_TEMPLATE
    .replace('__ROOM1_XLSX_B64__', room1_xlsx_b64)
    .replace('__ROOM3_XLSX_B64__', room3_xlsx_b64)
    .replace('__ROOM1_DOCX_B64__', room1_docx_b64)
    .replace('__ROOM3_DOCX_B64__', room3_docx_b64)
)

out_path = os.path.join(BASE, 'index.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(html)

size_kb = os.path.getsize(out_path) // 1024
print(f'已產生 index.html ({size_kb:,} KB)')
