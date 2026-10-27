/* 訪問患者リスト
 * Excelの患者マスタを取り込み、検索 → 当日リスト作成 → 印刷を行う。
 * データはすべてブラウザ内で処理し、外部へ送信しない。
 */
(function () {
  'use strict';

  const STORAGE_KEY = 'visitPatientList.v1';
  const SEARCH_LIMIT = 300;
  const PREVIEW_LIMIT = 20;
  const HEADER_SCAN_ROWS = 10;

  // 取込先の項目定義。題名（見出し）の自動判別に使うキーワードを持つ。
  // exact: 題名と完全一致、partial: 題名に含まれる、exclude: 含まれていたら部分一致では判別しない
  const FIELDS = [
    {
      key: 'id', label: 'ID',
      exact: ['ID', '患者ID', '患者番号', '患者No', '患者No.', 'カルテ番号', 'カルテNo', 'カルテNo.', 'カルテID', '患者コード', '利用者ID', '利用者番号', '番号', 'No', 'No.', 'コード'],
      partial: ['患者ID', '患者番号', '患者コード', '利用者ID', '利用者番号', 'カルテ', 'ID'],
      exclude: ['保険', '郵便', '電話', '受給', '記号', '医療機関', '施設', 'TEL', 'FAX', 'メール'],
    },
    {
      key: 'kana', label: 'ふりがな',
      exact: ['ふりがな', 'フリガナ', 'かな', 'カナ', 'よみ', 'よみがな', 'カナ氏名', '氏名カナ', '患者名カナ', 'カナ患者名', '患者カナ', '患者名ふりがな', '氏名ふりがな', '氏名(カナ)', '氏名(ふりがな)', '患者名(カナ)', '患者名(ふりがな)'],
      partial: ['ふりがな', 'フリガナ', 'カナ', 'よみ'],
      exclude: ['住所', '家族', '保護者', '緊急', '担当', '医師', 'ケアマネ', '連絡'],
    },
    {
      key: 'name', label: '患者名',
      exact: ['氏名', '患者名', '名前', '患者氏名', 'お名前', '漢字氏名', '氏名漢字', '患者名漢字', '利用者名', '利用者氏名', '氏名(漢字)', '患者名(漢字)'],
      partial: ['氏名', '患者名', '利用者名', '名前'],
      exclude: ['カナ', 'ふりがな', 'よみ', 'ローマ', '英字', '家族', '保護者', '緊急', '担当', '医師', 'ケアマネ', '連絡', '施設'],
    },
    {
      key: 'sex', label: '性別',
      exact: ['性別', '性', '男女', '男女区分', '性別区分'],
      partial: ['性別'],
      exclude: [],
    },
    {
      key: 'birth', label: '生年月日',
      exact: ['生年月日', '誕生日', '生年', '生年月日(西暦)', '生年月日(和暦)'],
      partial: ['生年月日', '誕生'],
      exclude: [],
    },
    {
      key: 'age', label: '年齢',
      exact: ['年齢', '歳', '才'],
      partial: ['年齢'],
      exclude: ['生年月日'],
    },
    {
      key: 'address', label: '住所',
      exact: ['住所', '現住所', '住所1', '所在地', '自宅住所', '患者住所', '住所(1)'],
      partial: ['住所', '所在地'],
      exclude: ['郵便', '〒', 'カナ', 'ふりがな', '緊急', '家族', '勤務', 'メール', 'MAIL', '2', '3', '番地', '建物', '方書'],
    },
    {
      key: 'address2', label: '住所（続き）',
      exact: ['住所2', '住所(2)', '番地', '方書', '建物', '建物名', 'マンション名', '住所(番地)', '住所(建物)'],
      partial: ['住所2', '番地', '建物', '方書'],
      exclude: ['郵便', 'カナ', 'ふりがな'],
    },
    {
      key: 'gh', label: 'グループホーム',
      exact: ['グループホーム', 'グループホーム名', 'GH', 'GH名', '施設', '施設名', '入居施設', '入居施設名', '入所施設', '居住施設', 'ホーム名'],
      partial: ['グループホーム', 'GH', '入居施設', '施設名', 'ホーム'],
      exclude: ['住所', '電話', '郵便', 'カナ', 'ふりがな', 'ID', '番号', 'コード', 'TEL', 'FAX'],
    },
  ];
  // グループホーム列が名称ではなく「○」「有」などの印だけの場合
  const GH_FLAG_VALUES = ['○', '◯', '〇', '●', '◎', '1', '有', 'あり', 'はい', '該当', '入居', 'gh', 'グループホーム', '✓', '✔', 'レ', 'y', 'yes'];
  const GH_NONE_VALUES = ['', '-', '―', 'ー', '×', '✕', '0', '無', 'なし', 'いいえ', '非該当', '居宅', '在宅', 'n', 'no'];
  const GH_FLAG_LABEL = 'グループホーム';

  // 印刷レイアウト（mm / pt）
  const PRINT_MARGIN_MM = 12;
  const PRINT_BASE_PT = 11;
  const PRINT_MIN_PT = 6.5;
  const PRINT_FONT = '"Hiragino Kaku Gothic ProN", "Hiragino Sans", "Yu Gothic UI", "Meiryo", sans-serif';
  const MM_TO_PX = 96 / 25.4;
  const PT_TO_PX = 96 / 72;
  const FIELD_BY_KEY = Object.fromEntries(FIELDS.map((f) => [f.key, f]));

  const ERAS = [
    { name: '令和', abbr: 'R', start: '2019-05-01', base: 2019 },
    { name: '平成', abbr: 'H', start: '1989-01-08', base: 1989 },
    { name: '昭和', abbr: 'S', start: '1926-12-25', base: 1926 },
    { name: '大正', abbr: 'T', start: '1912-07-30', base: 1912 },
    { name: '明治', abbr: 'M', start: '1868-01-25', base: 1868 },
  ];
  const WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];

  // ---------- 状態 ----------
  const state = {
    patients: [],
    importMeta: null,
    today: [],
    todaySort: null, // { key, dir: 1 | -1 }
    settings: { version: 2, visitDate: '', title: '訪問予定患者一覧', birthFormat: 'ad', paper: 'landscape', ghPosition: 'last', listOrder: 'group' },
  };
  let byKey = new Map();
  let pending = null; // 取込確認中のデータ
  let lastWorkbook = null; // 「取込設定を再確認」用（このページを開いている間のみ保持）

  const $ = (sel) => document.querySelector(sel);

  // ---------- 文字列ユーティリティ ----------
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }

  function toHiragana(s) {
    return s.replace(/[ァ-ヶ]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0x60));
  }

  // 検索・判別用の正規化：全角半角統一、カタカナ→ひらがな、小文字化、空白除去
  function normalize(s) {
    return toHiragana(String(s == null ? '' : s).normalize('NFKC').toLowerCase()).replace(/[\s　・･]/g, '');
  }

  function cleanText(v) {
    return String(v == null ? '' : v).normalize('NFKC').replace(/\s+/g, ' ').trim();
  }

  // ---------- 日付 ----------
  function pad2(n) { return String(n).padStart(2, '0'); }

  function makeIso(y, m, d) {
    y = Number(y); m = Number(m); d = Number(d);
    if (!y || !m || !d) return null;
    const dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    return `${y}-${pad2(m)}-${pad2(d)}`;
  }

  function eraBase(token) {
    const t = token.toUpperCase();
    const era = ERAS.find((e) => e.name === token || e.abbr === t);
    return era ? era.base : null;
  }

  // Excelのセル値（シリアル値・文字列・和暦など）を YYYY-MM-DD に変換。解釈できなければ null
  function parseDate(raw) {
    if (raw == null || raw === '') return null;
    if (raw instanceof Date) return makeIso(raw.getFullYear(), raw.getMonth() + 1, raw.getDate());
    if (typeof raw === 'number') {
      if (raw >= 18000101 && raw <= 29991231) {
        const s = String(Math.trunc(raw));
        return makeIso(s.slice(0, 4), s.slice(4, 6), s.slice(6, 8));
      }
      if (raw > 0 && raw < 2958466) {
        const p = XLSX.SSF.parse_date_code(raw);
        return p ? makeIso(p.y, p.m, p.d) : null;
      }
      return null;
    }
    const s = String(raw).normalize('NFKC').replace(/\s/g, '');
    let m = s.match(/^(明治|大正|昭和|平成|令和|[MTSHRmtshr])\.?(\d{1,2}|元)[年.\/\-](\d{1,2})[月.\/\-](\d{1,2})日?$/);
    if (m) {
      const base = eraBase(m[1]);
      const n = m[2] === '元' ? 1 : Number(m[2]);
      return base ? makeIso(base + n - 1, m[3], m[4]) : null;
    }
    m = s.match(/^(\d{4})[年.\/\-](\d{1,2})[月.\/\-](\d{1,2})日?/);
    if (m) return makeIso(m[1], m[2], m[3]);
    m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m) return makeIso(m[1], m[2], m[3]);
    if (/^\d+(\.\d+)?$/.test(s)) return parseDate(Number(s));
    return null;
  }

  function todayIso() {
    const d = new Date();
    return makeIso(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }

  function calcAge(birthIso, atIso) {
    if (!birthIso || !atIso) return null;
    const [by, bm, bd] = birthIso.split('-').map(Number);
    const [ty, tm, td] = atIso.split('-').map(Number);
    let age = ty - by;
    if (tm < bm || (tm === bm && td < bd)) age -= 1;
    return age >= 0 ? age : null;
  }

  function toWareki(iso) {
    const era = ERAS.find((e) => iso >= e.start);
    if (!era) return null;
    const [y, m, d] = iso.split('-').map(Number);
    return { era, year: y - era.base + 1, m, d };
  }

  function formatBirth(p, fmt) {
    if (!p.birth) return p.birthRaw || '';
    const [y, m, d] = p.birth.split('-');
    const ad = `${y}/${m}/${d}`;
    if (fmt === 'ad') return ad;
    const w = toWareki(p.birth);
    if (!w) return ad;
    const jp = `${w.era.abbr}${pad2(w.year)}.${m}.${d}`;
    return fmt === 'jp' ? jp : `${jp} / ${y}`;
  }

  function formatVisitDate(iso) {
    if (!iso) return '';
    const [y, m, d] = iso.split('-').map(Number);
    const wd = WEEKDAYS[new Date(y, m - 1, d).getDay()];
    const w = toWareki(iso);
    const jp = w ? `（${w.era.name}${w.year === 1 ? '元' : w.year}年）` : '';
    return `${y}年${m}月${d}日（${wd}）${jp}`;
  }

  function ageOf(p) {
    const a = calcAge(p.birth, state.settings.visitDate || todayIso());
    if (a != null) return a;
    const n = parseInt(String(p.ageRaw).normalize('NFKC'), 10);
    return Number.isFinite(n) ? n : null;
  }

  function normalizeSex(v) {
    const s = cleanText(v);
    const k = s.toLowerCase();
    if (['1', '男', '男性', 'm', 'male', 'おとこ'].includes(k)) return '男';
    if (['2', '女', '女性', 'f', 'female', 'おんな'].includes(k)) return '女';
    return s;
  }

  // グループホーム名を返す。グループホームでなければ空文字
  function normalizeGh(v) {
    const s = cleanText(v);
    const k = s.toLowerCase();
    if (GH_NONE_VALUES.includes(k)) return '';
    if (GH_FLAG_VALUES.includes(k)) return GH_FLAG_LABEL;
    return s;
  }

  // ---------- 保存 ----------
  function save() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        patients: state.patients,
        importMeta: state.importMeta,
        today: state.today,
        todaySort: state.todaySort,
        settings: state.settings,
      }));
    } catch (e) {
      console.warn(e);
    }
  }

  function load() {
    try {
      const data = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null');
      if (!data) return;
      state.patients = Array.isArray(data.patients) ? data.patients : [];
      state.importMeta = data.importMeta || null;
      state.today = Array.isArray(data.today) ? data.today : [];
      state.todaySort = data.todaySort || null;
      const saved = data.settings || {};
      // 旧版の初期値（A4縦）で保存されている場合は A4横 にする
      if (!saved.version) { delete saved.paper; saved.version = 2; }
      Object.assign(state.settings, saved);
    } catch (e) {
      console.warn(e);
    }
  }

  function rebuildIndex() {
    byKey = new Map(state.patients.map((p) => [p.key, p]));
    state.today = state.today.filter((k) => byKey.has(k));
  }

  // ---------- Excel 読込 ----------
  async function readWorkbook(file) {
    const buf = await file.arrayBuffer();
    if (/\.csv$/i.test(file.name)) {
      let text;
      try {
        text = new TextDecoder('utf-8', { fatal: true }).decode(buf);
      } catch (e) {
        text = new TextDecoder('shift_jis').decode(buf);
      }
      // raw: true で先頭の 0 などを文字列のまま保持する
      return XLSX.read(text.replace(/^﻿/, ''), { type: 'string', raw: true });
    }
    return XLSX.read(buf, { type: 'array' });
  }

  function readSheet(wb, sheetName) {
    const ws = wb.Sheets[sheetName];
    if (!ws || !ws['!ref']) return { fmt: [], raw: [], rowOffset: 0, colOffset: 0 };
    const range = XLSX.utils.decode_range(ws['!ref']);
    const opts = { header: 1, defval: '', blankrows: true };
    return {
      fmt: XLSX.utils.sheet_to_json(ws, Object.assign({ raw: false }, opts)),
      raw: XLSX.utils.sheet_to_json(ws, Object.assign({ raw: true }, opts)),
      rowOffset: range.s.r,
      colOffset: range.s.c,
    };
  }

  // 題名1つに対する項目ごとのスコア
  function scoreHeader(field, header) {
    const h = normalize(header);
    if (!h) return 0;
    if (field.exact.some((k) => normalize(k) === h)) return 1000;
    if (field.exclude.some((k) => h.includes(normalize(k)))) return 0;
    let best = 0;
    for (const k of field.partial) {
      const nk = normalize(k);
      if (h.includes(nk)) best = Math.max(best, 100 + nk.length * 10 - (h.length - nk.length));
    }
    return best;
  }

  // 題名から 項目 → 列番号 を自動判別（1列は1項目にのみ割当）
  function autoMap(headers) {
    const cands = [];
    FIELDS.forEach((f) => headers.forEach((h, col) => {
      const s = scoreHeader(f, h);
      if (s > 0) cands.push({ field: f.key, col, score: s });
    }));
    cands.sort((a, b) => b.score - a.score || a.col - b.col);
    const mapping = {};
    const usedCols = new Set();
    for (const c of cands) {
      if (mapping[c.field] != null || usedCols.has(c.col)) continue;
      mapping[c.field] = c.col;
      usedCols.add(c.col);
    }
    return mapping;
  }

  // 先頭10行の中から、最も多くの項目を判別できた行を見出し行とする
  function guessHeaderRow(rows) {
    let best = -1;
    let bestCount = 0;
    for (let r = 0; r < Math.min(rows.length, HEADER_SCAN_ROWS); r++) {
      const cnt = Object.keys(autoMap(rows[r].map(cleanText))).length;
      if (cnt > bestCount) { best = r; bestCount = cnt; }
    }
    if (best >= 0) return best;
    const firstNonEmpty = rows.findIndex((row) => row.some((v) => cleanText(v) !== ''));
    return Math.max(firstNonEmpty, 0);
  }

  function colLetter(i) {
    return XLSX.utils.encode_col(i);
  }

  // 見出し行・割当に従って患者データを組み立てる
  function buildPatients(p) {
    const { sheet, headerRow, mapping } = p;
    const get = (rowArr, key) => (mapping[key] != null ? rowArr[mapping[key]] : '');
    const out = [];
    for (let r = headerRow + 1; r < sheet.fmt.length; r++) {
      const fmt = sheet.fmt[r] || [];
      const raw = sheet.raw[r] || [];
      const mappedValues = Object.values(mapping).map((c) => cleanText(fmt[c]));
      if (mappedValues.every((v) => v === '')) continue;
      const birthRawVal = get(raw, 'birth');
      const birthText = cleanText(get(fmt, 'birth'));
      const address = [cleanText(get(fmt, 'address')), cleanText(get(fmt, 'address2'))].filter(Boolean).join(' ');
      out.push({
        row: r + sheet.rowOffset + 1,
        id: cleanText(get(fmt, 'id')),
        kana: cleanText(get(fmt, 'kana')),
        name: cleanText(get(fmt, 'name')),
        sex: normalizeSex(get(fmt, 'sex')),
        birth: parseDate(birthRawVal) || parseDate(birthText),
        birthRaw: birthText,
        ageRaw: cleanText(get(fmt, 'age')),
        address,
        addr1: cleanText(get(fmt, 'address')),
        gh: normalizeGh(get(fmt, 'gh')),
      });
    }
    // 重複しないキーを割り当てる（IDがあればID、重複・未設定なら行番号を付加）
    const seen = new Map();
    out.forEach((pt) => seen.set(pt.id, (seen.get(pt.id) || 0) + 1));
    out.forEach((pt) => {
      if (!pt.id) pt.key = `row:${pt.row}`;
      else pt.key = seen.get(pt.id) === 1 ? `id:${pt.id}` : `id:${pt.id}#row${pt.row}`;
    });
    return out;
  }

  // ---------- 取込確認画面 ----------
  async function onFileSelected(e) {
    const file = e.target.files[0];
    e.target.value = '';
    if (!file) return;
    let wb;
    try {
      wb = await readWorkbook(file);
    } catch (err) {
      alert('ファイルを読み込めませんでした。\n' + err.message);
      return;
    }
    lastWorkbook = { wb, fileName: file.name };
    openImportDialog(wb, file.name, null);
  }

  function openImportDialog(wb, fileName, prevMeta) {
    const sheetName = prevMeta && wb.SheetNames.includes(prevMeta.sheetName)
      ? prevMeta.sheetName
      : pickSheet(wb);
    pending = { wb, fileName, manual: new Set() };
    $('#import-file-name').textContent = `ファイル: ${fileName}`;
    const sel = $('#sheet-select');
    sel.innerHTML = wb.SheetNames.map((n) => `<option value="${esc(n)}">${esc(n)}</option>`).join('');
    sel.value = sheetName;
    $('#sheet-label').style.display = wb.SheetNames.length > 1 ? '' : 'none';
    loadSheetIntoPending(sheetName, prevMeta);
    $('#import-dialog').showModal();
  }

  // 複数シートの場合、最も多く項目を判別できたシートを初期選択にする
  function pickSheet(wb) {
    let best = wb.SheetNames[0];
    let bestCount = -1;
    for (const name of wb.SheetNames) {
      const sheet = readSheet(wb, name);
      const hr = guessHeaderRow(sheet.fmt);
      const cnt = Object.keys(autoMap((sheet.fmt[hr] || []).map(cleanText))).length;
      if (cnt > bestCount) { best = name; bestCount = cnt; }
    }
    return best;
  }

  function loadSheetIntoPending(sheetName, prevMeta) {
    pending.sheetName = sheetName;
    pending.sheet = readSheet(pending.wb, sheetName);
    const useSaved = prevMeta && prevMeta.sheetName === sheetName;
    pending.headerRow = useSaved ? prevMeta.headerRow : guessHeaderRow(pending.sheet.fmt);
    $('#header-row').value = pending.headerRow + pending.sheet.rowOffset + 1;
    pending.manual = new Set(useSaved ? prevMeta.manual || [] : []);
    pending.mapping = useSaved ? Object.assign({}, prevMeta.mapping) : autoMap(headersOf());
    renderImportDialog();
  }

  function headersOf() {
    return (pending.sheet.fmt[pending.headerRow] || []).map(cleanText);
  }

  function setHeaderRow(excelRow) {
    const r = excelRow - pending.sheet.rowOffset - 1;
    if (!Number.isInteger(r) || r < 0 || r >= pending.sheet.fmt.length) return;
    pending.headerRow = r;
    pending.manual = new Set();
    pending.mapping = autoMap(headersOf());
    renderImportDialog();
  }

  function renderImportDialog() {
    const sheet = pending.sheet;
    const headers = headersOf();
    const width = Math.max(headers.length, ...sheet.fmt.slice(pending.headerRow, pending.headerRow + 50).map((r) => r.length), 0);
    const colToField = {};
    Object.entries(pending.mapping).forEach(([k, c]) => { colToField[c] = k; });

    const rowsHtml = [];
    for (let c = 0; c < width; c++) {
      const samples = [];
      for (let r = pending.headerRow + 1; r < sheet.fmt.length && samples.length < 3; r++) {
        const v = cleanText((sheet.fmt[r] || [])[c]);
        if (v) samples.push(v);
      }
      if (!headers[c] && samples.length === 0) continue;
      const fk = colToField[c] || '';
      const opts = ['<option value="">（取り込まない）</option>']
        .concat(FIELDS.map((f) => `<option value="${f.key}"${f.key === fk ? ' selected' : ''}>${esc(f.label)}</option>`))
        .join('');
      const badge = fk ? (pending.manual.has(fk) ? '<span class="badge manual">手動</span>' : '<span class="badge auto">自動判別</span>') : '';
      rowsHtml.push(`<tr class="${fk ? 'mapped' : ''}">
        <td>${colLetter(c + sheet.colOffset)}列</td>
        <td>${esc(headers[c]) || '<span class="hint">（題名なし）</span>'}</td>
        <td class="sample" title="${esc(samples.join(' / '))}">${esc(samples.join(' / '))}</td>
        <td><select data-col="${c}">${opts}</select></td>
        <td>${badge}</td>
      </tr>`);
    }
    $('#mapping-table tbody').innerHTML = rowsHtml.join('') || '<tr><td colspan="5" class="empty">データがありません</td></tr>';

    const patients = buildPatients(pending);
    pending.patients = patients;
    renderPreview(patients);
    renderWarnings(patients);
  }

  function renderPreview(patients) {
    const fmt = state.settings.birthFormat;
    const visit = state.settings.visitDate || todayIso();
    $('#preview-info').textContent = `（取込対象 ${patients.length} 件${patients.length > PREVIEW_LIMIT ? `、先頭 ${PREVIEW_LIMIT} 件を表示` : ''}）`;
    const html = patients.slice(0, PREVIEW_LIMIT).map((p) => {
      const age = calcAge(p.birth, visit);
      return `<tr>
        <td class="num">${p.row}</td>
        <td>${esc(p.id)}</td><td>${esc(p.kana)}</td><td>${esc(p.name)}</td><td>${esc(p.sex)}</td>
        <td>${esc(formatBirth(p, fmt))}</td>
        <td class="num">${esc(age != null ? age : p.ageRaw)}</td>
        <td class="addr">${esc(p.address)}</td>
        <td>${esc(p.gh)}</td>
      </tr>`;
    }).join('');
    $('#preview-table tbody').innerHTML = html || '<tr><td colspan="9" class="empty">取込対象の行がありません</td></tr>';
  }

  function renderWarnings(patients) {
    const m = pending.mapping;
    const msgs = [];
    const canImport = (m.id != null || m.name != null) && patients.length > 0;
    if (m.id == null && m.name == null) msgs.push({ error: true, text: '「ID」または「患者名」の列を指定してください。' });
    else if (patients.length === 0) msgs.push({ error: true, text: '見出し行より下に取込対象のデータがありません。見出しの行を確認してください。' });

    const missing = ['id', 'kana', 'name', 'sex', 'birth', 'address', 'gh']
      .filter((k) => m[k] == null).map((k) => FIELD_BY_KEY[k].label);
    if (m.birth == null && m.age == null) missing.push('年齢');
    if (missing.length) msgs.push({ text: `未割当の項目: ${missing.join('、')}（空欄で取り込まれます）` });

    if (m.id != null) {
      const counts = new Map();
      patients.forEach((p) => { if (p.id) counts.set(p.id, (counts.get(p.id) || 0) + 1); });
      const dups = [...counts.entries()].filter(([, n]) => n > 1).map(([id]) => id);
      if (dups.length) msgs.push({ text: `IDが重複しています: ${dups.slice(0, 10).join('、')}${dups.length > 10 ? ` ほか${dups.length - 10}件` : ''}` });
      const noId = patients.filter((p) => !p.id).length;
      if (noId) msgs.push({ text: `IDが空欄の行が ${noId} 件あります。` });
    }
    if (m.birth != null) {
      const bad = patients.filter((p) => !p.birth && p.birthRaw);
      if (bad.length) msgs.push({ text: `生年月日を日付として解釈できない行が ${bad.length} 件あります（例: ${bad.slice(0, 3).map((p) => `${p.row}行目「${p.birthRaw}」`).join('、')}）。そのまま表示します。` });
    }

    $('#mapping-warnings').innerHTML = msgs.map((w) => `<div class="warn${w.error ? ' error' : ''}">${esc(w.text)}</div>`).join('');
    $('#btn-import-ok').disabled = !canImport;
  }

  function onMappingChange(e) {
    const sel = e.target.closest('select[data-col]');
    if (!sel) return;
    const col = Number(sel.dataset.col);
    const field = sel.value;
    Object.keys(pending.mapping).forEach((k) => {
      if (pending.mapping[k] === col || k === field) {
        delete pending.mapping[k];
        pending.manual.delete(k);
      }
    });
    if (field) {
      pending.mapping[field] = col;
      pending.manual.add(field);
    }
    renderImportDialog();
  }

  function confirmImport() {
    const patients = buildPatients(pending);
    state.patients = patients;
    state.importMeta = {
      fileName: pending.fileName,
      sheetName: pending.sheetName,
      headerRow: pending.headerRow,
      mapping: pending.mapping,
      manual: [...pending.manual],
      importedAt: new Date().toISOString(),
    };
    const before = state.today.length;
    rebuildIndex();
    save();
    $('#import-dialog').close();
    pending = null;
    renderAll();
    const dropped = before - state.today.length;
    if (dropped > 0) alert(`新しいマスタに見つからない患者 ${dropped} 名を当日リストから外しました。`);
  }

  // ---------- 検索 ----------
  function searchPatients() {
    const q = $('#search-input').value;
    const target = $('#search-target').value;
    const terms = q.split(/[\s　]+/).map(normalize).filter(Boolean);
    const fieldsFor = target === 'all' ? ['id', 'kana', 'name', 'address', 'gh'] : [target];
    return state.patients.filter((p) => {
      if (!terms.length) return true;
      const hay = fieldsFor.map((f) => normalize(p[f])).join('\n');
      return terms.every((t) => hay.includes(t));
    });
  }

  function renderSearch() {
    const has = state.patients.length > 0;
    $('#search-input').disabled = !has;
    $('#search-target').disabled = !has;
    const tbody = $('#search-table tbody');
    if (!has) {
      $('#search-info').textContent = 'Excelの患者マスタを取り込むと検索できます。';
      tbody.innerHTML = '<tr><td colspan="9" class="empty">患者マスタが取り込まれていません</td></tr>';
      return;
    }
    const results = searchPatients();
    const shown = results.slice(0, SEARCH_LIMIT);
    $('#search-info').textContent = `該当 ${results.length} 件${results.length > SEARCH_LIMIT ? `（先頭 ${SEARCH_LIMIT} 件を表示。条件を絞り込んでください）` : ''}　チェックを入れると当日リストに追加されます。`;
    const inToday = new Set(state.today);
    const fmt = state.settings.birthFormat;
    tbody.innerHTML = shown.map((p) => {
      const checked = inToday.has(p.key);
      const age = ageOf(p);
      return `<tr data-key="${esc(p.key)}" class="${checked ? 'checked' : ''}">
        <td class="col-check"><input type="checkbox" ${checked ? 'checked' : ''} aria-label="当日リストに追加"></td>
        <td>${esc(p.id)}</td><td>${esc(p.kana)}</td><td>${esc(p.name)}</td><td>${esc(p.sex)}</td>
        <td>${esc(formatBirth(p, fmt))}</td><td class="num">${esc(age != null ? age : '')}</td>
        <td class="addr">${esc(p.address)}</td>
        <td class="gh">${esc(p.gh)}</td>
      </tr>`;
    }).join('') || '<tr><td colspan="9" class="empty">該当する患者がいません</td></tr>';
  }

  function onSearchTableClick(e) {
    const tr = e.target.closest('tr[data-key]');
    if (!tr) return;
    const key = tr.dataset.key;
    const cb = tr.querySelector('input[type="checkbox"]');
    if (e.target !== cb) cb.checked = !cb.checked;
    if (cb.checked) {
      if (!state.today.includes(key)) state.today.push(key);
      state.todaySort = null;
    } else {
      state.today = state.today.filter((k) => k !== key);
    }
    tr.classList.toggle('checked', cb.checked);
    save();
    renderToday();
  }

  // ---------- 当日リスト ----------
  const collator = new Intl.Collator('ja', { numeric: true, sensitivity: 'base' });

  function sortValue(p, key) {
    switch (key) {
      case 'age': return ageOf(p);
      case 'birth': return p.birth || null;
      case 'kana': return p.kana ? toHiragana(p.kana.normalize('NFKC')) : null;
      default: return p[key] || null;
    }
  }

  function sortToday(key) {
    const dir = state.todaySort && state.todaySort.key === key ? -state.todaySort.dir : 1;
    state.todaySort = { key, dir };
    state.settings.listOrder = 'group';
    $('#list-order').value = 'group';
    applySort();
    save();
    renderToday();
  }

  function applySort() {
    if (!state.todaySort) return;
    const { key, dir } = state.todaySort;
    state.today = state.today
      .map((k, i) => ({ k, i, v: sortValue(byKey.get(k), key) }))
      .sort((a, b) => {
        // 空欄は昇順・降順どちらでも末尾
        if (a.v == null && b.v == null) return a.i - b.i;
        if (a.v == null) return 1;
        if (b.v == null) return -1;
        const c = typeof a.v === 'number' && typeof b.v === 'number' ? a.v - b.v : collator.compare(String(a.v), String(b.v));
        return c * dir || a.i - b.i;
      })
      .map((x) => x.k);
  }

  // 当日リストを「居宅」と「グループホーム（名称ごと）」に分ける。
  // 並べ替えは各グループの中で行われ、グループどうしが混ざることはない。
  // 訪問ルート順のときは並びをそのまま使い、連続する同じグループホームを1つにまとめる。
  function groupedToday() {
    if (state.settings.listOrder === 'route') {
      const runs = [];
      for (const k of state.today) {
        const g = byKey.get(k).gh || '';
        const last = runs[runs.length - 1];
        if (last && last.gh === g) last.keys.push(k);
        else runs.push({ gh: g, keys: [k] });
      }
      return runs;
    }
    const groups = new Map();
    for (const k of state.today) {
      const g = byKey.get(k).gh || '';
      if (!groups.has(g)) groups.set(g, []);
      groups.get(g).push(k);
    }
    const home = groups.has('') ? [{ gh: '', keys: groups.get('') }] : [];
    const gh = [...groups.keys()].filter(Boolean).sort(collator.compare).map((g) => ({ gh: g, keys: groups.get(g) }));
    return state.settings.ghPosition === 'first' ? gh.concat(home) : home.concat(gh);
  }

  function groupLabel(g) {
    if (!g.gh) return `居宅（${g.keys.length}名）`;
    const name = /グループホーム|GH/i.test(g.gh) ? g.gh : `グループホーム　${g.gh}`;
    return `${name}（${g.keys.length}名）`;
  }

  // 見出し行を出すか。訪問ルート順ではグループホームの直前にだけ出す
  function showGroupHeader(groups, g) {
    if (state.settings.listOrder === 'route') return !!g.gh;
    return groups.some((x) => x.gh);
  }

  function renderToday() {
    const fmt = state.settings.birthFormat;
    $('#today-count').textContent = `${state.today.length}名`;
    document.querySelectorAll('#today-table th.sortable').forEach((th) => {
      th.classList.remove('asc', 'desc');
      if (state.todaySort && th.dataset.key === state.todaySort.key) th.classList.add(state.todaySort.dir > 0 ? 'asc' : 'desc');
    });
    const groups = groupedToday();
    const html = [];
    for (const g of groups) {
      if (showGroupHeader(groups, g)) html.push(`<tr class="group-row${g.gh ? ' gh' : ''}"><td colspan="8">${esc(groupLabel(g))}</td></tr>`);
      const last = g.keys.length - 1;
      g.keys.forEach((k, i) => {
        const p = byKey.get(k);
        const age = ageOf(p);
        html.push(`<tr data-key="${esc(k)}">
          <td>${esc(p.id)}</td><td>${esc(p.kana)}</td><td>${esc(p.name)}</td><td>${esc(p.sex)}</td>
          <td>${esc(formatBirth(p, fmt))}</td><td class="num">${esc(age != null ? age : '')}</td>
          <td class="addr">${esc(p.address)}</td>
          <td class="col-ops"><div class="ops">
            <button type="button" class="btn small" data-op="up" ${i === 0 ? 'disabled' : ''} title="上へ">↑</button>
            <button type="button" class="btn small" data-op="down" ${i === last ? 'disabled' : ''} title="下へ">↓</button>
            <button type="button" class="btn small" data-op="remove" title="リストから外す">外す</button>
          </div></td>
        </tr>`);
      });
    }
    $('#today-table tbody').innerHTML = html.join('') || '<tr><td colspan="8" class="empty">検索結果でチェックした患者がここに表示されます</td></tr>';
    $('#btn-print').disabled = state.today.length === 0;
    $('#gh-position').disabled = state.settings.listOrder === 'route';
    if (window.VisitRoute) window.VisitRoute.refresh();
  }

  function onTodayTableClick(e) {
    const th = e.target.closest('th.sortable');
    if (th) { sortToday(th.dataset.key); return; }
    const btn = e.target.closest('button[data-op]');
    if (!btn) return;
    const key = btn.closest('tr').dataset.key;
    if (btn.dataset.op === 'remove') {
      state.today = state.today.filter((k) => k !== key);
      renderSearch();
    } else {
      // 同じグループ内でのみ入れ替える
      const g = groupedToday().find((x) => x.keys.includes(key));
      const i = g.keys.indexOf(key);
      const j = btn.dataset.op === 'up' ? i - 1 : i + 1;
      if (j < 0 || j >= g.keys.length) return;
      const a = state.today.indexOf(g.keys[i]);
      const b = state.today.indexOf(g.keys[j]);
      [state.today[a], state.today[b]] = [state.today[b], state.today[a]];
      state.todaySort = null;
    }
    save();
    renderToday();
  }

  // ---------- 印刷 ----------
  let measureCtx = null;

  function textWidthPx(text, pt) {
    if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d');
    measureCtx.font = `${pt * PT_TO_PX}px ${PRINT_FONT}`;
    return measureCtx.measureText(text).width;
  }

  // 列幅（mm）。住所は残りの幅すべて
  function printColumns() {
    const paperWidth = state.settings.paper === 'landscape' ? 297 : 210;
    const total = paperWidth - PRINT_MARGIN_MM * 2;
    const cols = [
      { key: 'id', mm: 22 },
      { key: 'kana', mm: 40 },
      { key: 'name', mm: 34 },
      { key: 'sex', mm: 11 },
      { key: 'birth', mm: state.settings.birthFormat === 'both' ? 32 : 24 },
      { key: 'age', mm: 12 },
    ];
    const used = cols.reduce((s, c) => s + c.mm, 0);
    cols.push({ key: 'address', mm: Math.max(total - used, 30) });
    return cols;
  }

  // セル幅に収まるよう文字を縮小する。最小サイズでも収まらなければ折り返す
  function fitCell(text, mm, cls) {
    const avail = (mm - 3) * MM_TO_PX; // 左右の余白と罫線ぶん
    const w = textWidthPx(text, PRINT_BASE_PT);
    let style = '';
    if (text && w > avail) {
      const pt = Math.max(PRINT_MIN_PT, Math.floor((PRINT_BASE_PT * avail / w) * 10) / 10);
      style = `font-size:${pt}pt;`;
      if (textWidthPx(text, pt) > avail) style += 'white-space:normal;';
    }
    return `<td class="${cls || ''}"${style ? ` style="${style}"` : ''}>${esc(text)}</td>`;
  }

  function renderPrint() {
    const fmt = state.settings.birthFormat;
    const cols = printColumns();
    const width = Object.fromEntries(cols.map((c) => [c.key, c.mm]));
    $('#print-h1').textContent = state.settings.title || '訪問予定患者一覧';
    $('#print-meta').textContent = `訪問日: ${formatVisitDate(state.settings.visitDate || todayIso())}　${state.today.length}名`;
    $('#print-cols').innerHTML = cols.map((c) => `<col style="width:${c.mm}mm">`).join('');
    const groups = groupedToday();
    const html = [];
    for (const g of groups) {
      if (showGroupHeader(groups, g)) html.push(`<tr class="group-row"><td colspan="7">${esc(groupLabel(g))}</td></tr>`);
      for (const k of g.keys) {
        const p = byKey.get(k);
        const age = ageOf(p);
        html.push(`<tr>
          ${fitCell(p.id, width.id)}${fitCell(p.kana, width.kana)}${fitCell(p.name, width.name)}${fitCell(p.sex, width.sex)}
          ${fitCell(formatBirth(p, fmt), width.birth)}${fitCell(age != null ? String(age) : '', width.age, 'num')}
          ${fitCell(p.address, width.address, 'addr')}
        </tr>`);
      }
    }
    $('#print-body').innerHTML = html.join('');
    $('#page-style').textContent = `@page { size: A4 ${state.settings.paper === 'landscape' ? 'landscape' : 'portrait'}; margin: ${PRINT_MARGIN_MM}mm; }`;
  }

  // ---------- 全体 ----------
  function renderMasterStatus() {
    const meta = state.importMeta;
    $('#master-status').textContent = meta
      ? `患者マスタ: ${meta.fileName}（${meta.sheetName}）${state.patients.length}件 ／ 取込日時 ${new Date(meta.importedAt).toLocaleString('ja-JP')}`
      : '患者マスタ未取込';
    $('#btn-remap').disabled = !lastWorkbook;
  }

  function renderAll() {
    renderMasterStatus();
    renderSearch();
    renderToday();
  }

  function bindSetting(sel, key, rerender) {
    const el = $(sel);
    el.value = state.settings[key];
    el.addEventListener(el.tagName === 'SELECT' ? 'change' : 'input', () => {
      state.settings[key] = el.value;
      save();
      if (rerender) {
        if (key === 'visitDate' && state.todaySort && state.todaySort.key === 'age') applySort();
        renderSearch();
        renderToday();
      }
    });
  }

  function init() {
    load();
    if (!state.settings.visitDate) state.settings.visitDate = todayIso();
    rebuildIndex();

    $('#file-input').addEventListener('change', onFileSelected);
    $('#btn-remap').addEventListener('click', () => {
      if (lastWorkbook) openImportDialog(lastWorkbook.wb, lastWorkbook.fileName, state.importMeta && state.importMeta.fileName === lastWorkbook.fileName ? state.importMeta : null);
    });
    $('#btn-clear-all').addEventListener('click', () => {
      if (!confirm('このブラウザに保存されている患者マスタと当日リストを消去します。よろしいですか？')) return;
      localStorage.removeItem(STORAGE_KEY);
      if (window.VisitRoute) window.VisitRoute.clearStorage();
      state.patients = [];
      state.importMeta = null;
      state.today = [];
      state.todaySort = null;
      lastWorkbook = null;
      rebuildIndex();
      renderAll();
    });

    $('#sheet-select').addEventListener('change', (e) => loadSheetIntoPending(e.target.value, null));
    $('#header-row').addEventListener('change', (e) => setHeaderRow(Number(e.target.value)));
    $('#mapping-table').addEventListener('change', onMappingChange);
    $('#btn-import-cancel').addEventListener('click', () => { $('#import-dialog').close(); pending = null; });
    $('#btn-import-ok').addEventListener('click', confirmImport);

    $('#search-input').addEventListener('input', renderSearch);
    $('#search-target').addEventListener('change', renderSearch);
    $('#search-table').addEventListener('click', onSearchTableClick);

    $('#today-table').addEventListener('click', onTodayTableClick);
    $('#btn-clear-today').addEventListener('click', () => {
      if (!state.today.length || !confirm('当日リストを空にします。よろしいですか？')) return;
      state.today = [];
      save();
      renderSearch();
      renderToday();
    });
    $('#btn-print').addEventListener('click', () => {
      if (!state.today.length) return;
      renderPrint();
      window.print();
    });
    window.addEventListener('beforeprint', renderPrint);

    bindSetting('#visit-date', 'visitDate', true);
    bindSetting('#print-title', 'title', false);
    bindSetting('#birth-format', 'birthFormat', true);
    bindSetting('#paper', 'paper', false);
    bindSetting('#gh-position', 'ghPosition', true);
    bindSetting('#list-order', 'listOrder', true);

    renderAll();
  }

  // 地図・訪問ルート（route.js）から使う窓口
  window.VisitApp = {
    // 当日リストの患者（表示順）
    todayPatients() {
      return groupedToday().flatMap((g) => g.keys).map((k) => byKey.get(k));
    },
    // 計算した訪問順を当日リストに反映する
    applyRouteOrder(keys) {
      const set = new Set(keys);
      state.today = keys.filter((k) => byKey.has(k)).concat(state.today.filter((k) => !set.has(k)));
      state.todaySort = null;
      state.settings.listOrder = 'route';
      $('#list-order').value = 'route';
      save();
      renderToday();
    },
  };

  document.addEventListener('DOMContentLoaded', init);
})();
