'use strict';
const $ = (id) => document.getElementById(id);
const el = (tag, attrs = {}, ...kids) => {
  const n = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (k === 'class') n.className = v;
    else if (k === 'text') n.textContent = v;
    else if (k.startsWith('on')) n.addEventListener(k.slice(2), v);
    else if (v !== null && v !== undefined) n.setAttribute(k, v);
  }
  kids.flat().forEach(c => n.appendChild(typeof c === 'string' ? document.createTextNode(c) : c));
  return n;
};

let INIT = null;
const KANTOKU_YAKU = ['工事施工管理幹', '副主幹', '技師'];
const HANSHUTSU = ['Co、As殻', '汚泥', '残土', '残土（2）'];

let BUSY = 0;
function busy(on) {
  BUSY += on ? 1 : -1;
  if (BUSY < 0) BUSY = 0;
  const b = document.getElementById('busy');
  if (b) b.classList.toggle('hide', BUSY === 0);
}

async function api(path, body) {
  const opt = body === undefined ? {} :
    { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) };
  busy(true);
  try {
    const r = await fetch(path, opt);
    const j = await r.json();
    if (j && j.error && !j.files) console.warn(path, j.error);
    return j;
  } catch (e) {
    console.error(path, e);
    return { error: '通信に失敗しました（' + e + '）' };
  } finally {
    busy(false);
  }
}

/** ボタンを押したことが分かるよう、処理中は文字を変えて止める */
async function withButton(btn, label, fn) {
  if (!btn) return fn();
  const old = btn.textContent;
  btn.disabled = true;
  btn.textContent = label;
  try {
    return await fn();
  } finally {
    btn.disabled = false;
    btn.textContent = old;
  }
}

/* ============================================================ 初期化 */
async function init() {
  INIT = await api('/api/init');

  // Excel の状態
  const st = $('excelState');
  const ver = INIT.version ? ('  版 ' + INIT.version) : '';
  if (INIT.excel.ok) { st.textContent = 'Excel ' + INIT.excel.version + ' OK' + ver; }
  else { st.textContent = 'Excel が見つかりません' + ver; st.className = 'state ng'; }

  // 発注者・協力会社の候補
  fillList('dl-hacchusha', [...new Set((INIT.company.発注者 || []).map(x => x.名称))]);
  fillList('dl-ka', [...new Set((INIT.company.発注者 || []).map(x => x.課))]);
  fillList('dl-kaisha', (INIT.company.協力会社 || []).map(x => x.会社));

  document.querySelectorAll('select.staff').forEach(fillStaff);
  $('会社名').value = INIT.company.会社名 || '';

  buildItemChips();
  buildKankyo();
  buildTable('tbl-kantoku', 3, (i) => [
    cellInput('kantoku-yaku-' + i, KANTOKU_YAKU[i] || ''),
    cellStaffFree('kantoku-name-' + i)]);
  buildTable('tbl-itaku', 3, (i) => [
    cellInput('itaku-種別-' + i, i === 0 ? '交通誘導員' : ''),
    cellInput('itaku-会社-' + i, '', 'dl-kaisha', () => autoKyoryoku('itaku', i)),
    cellInput('itaku-担当-' + i), cellInput('itaku-電話-' + i)]);
  buildTable('tbl-shitauke', 5, (i) => [
    cellInput('shitauke-工種-' + i),
    cellInput('shitauke-会社-' + i, '', 'dl-kaisha', () => autoKyoryoku('shitauke', i)),
    cellInput('shitauke-担当-' + i), cellInput('shitauke-電話-' + i)]);
  buildHanshutsu();

  (INIT.company.安全方針既定 || []).forEach((v, i) => { const n = $('安全方針' + (i + 1)); if (n) n.value = v; });

  // 設定タブ
  $('set-雛形').value = INIT.settings.雛形フォルダー || '';
  $('set-必携').value = INIT.settings.必携フォルダー || '';
  $('set-親').value = INIT.settings.工事フォルダー親 || '';
  drawStaffTable();
  drawKyoryokuTable();

  // 保存済み工事
  const pp = $('projectPicker');
  INIT.projects.forEach(p => pp.appendChild(el('option', { value: p, text: p })));

  wire();
  updateKouki();
  checkSetup();
}

function checkSetup() {
  const w = $('warn');
  const msgs = [];
  if (!INIT.settings.雛形フォルダー) msgs.push('施工計画書雛形フォルダーが見つかりません。「設定」で場所を指定してください。');
  if (!INIT.excel.ok) msgs.push('Excel が見つかりません。Excel（デスクトップ版）が入っているパソコンで使ってください。');
  if (!msgs.length) { w.classList.add('hide'); return; }
  w.innerHTML = '';
  w.appendChild(el('span', { text: msgs.join(' / ') }));
  if (!INIT.settings.雛形フォルダー) {
    w.appendChild(el('button', {
      text: '設定を開く',
      onclick: () => document.querySelector('#tabs button[data-tab="settei"]').click()
    }));
  }
  w.classList.remove('hide');
}

function fillList(id, values) {
  const dl = $(id); dl.innerHTML = '';
  values.filter(Boolean).forEach(v => dl.appendChild(el('option', { value: v })));
}

function fillStaff(sel) {
  const cur = sel.value;
  sel.innerHTML = '';
  sel.appendChild(el('option', { value: '', text: '―' }));
  (INIT.staff || []).forEach(s => sel.appendChild(el('option', { value: s.氏名, text: s.氏名 })));
  sel.value = cur;
}

function cellInput(id, val = '', list = null, onchange = null) {
  const td = el('td');
  const inp = el('input', { id, value: val });
  if (list) inp.setAttribute('list', list);
  if (onchange) inp.addEventListener('change', onchange);
  td.appendChild(inp); return td;
}

function cellStaffFree(id) {
  const td = el('td');
  const inp = el('input', { id, placeholder: '未定なら空欄' });
  td.appendChild(inp); return td;
}

function buildTable(tid, rows, mk) {
  const tb = $(tid).querySelector('tbody'); tb.innerHTML = '';
  for (let i = 0; i < rows; i++) tb.appendChild(el('tr', {}, mk(i)));
}

function buildHanshutsu() {
  const tb = $('tbl-hanshutsu').querySelector('tbody'); tb.innerHTML = '';
  HANSHUTSU.forEach((name, i) => {
    const first = el('td', {}, name);
    if (name === '残土（2）') first.appendChild(el('input', { id: 'han-内容-' + i, placeholder: '内容（例：工事用道路盛土材）' }));
    tb.appendChild(el('tr', {}, [first,
      cellInput('han-処分先-' + i), cellInput('han-所在地-' + i), cellInput('han-距離-' + i)]));
  });
}

function buildItemChips() {
  const box = $('itemList'); box.innerHTML = '';
  (INIT.items || []).forEach(it => {
    const cb = el('input', { type: 'checkbox', id: 'item-' + it.id });
    cb.checked = !it.needs;            // 取り込みが要る項目は、読み取れたら自動で入る
    const note = el('span', { class: 'note', id: 'need-' + it.id });
    const lab = el('label', { class: cb.checked ? 'checked' : '', id: 'lab-' + it.id },
      cb, it.label, note);
    cb.addEventListener('change', () => {
      cb.dataset.touched = '1';
      lab.classList.toggle('checked', cb.checked);
    });
    box.appendChild(lab);
  });
  syncNeedItems();
}

/** （1）（1-2）（5）（6）は設計書の取り込みが要る。取り込めたら選べるようにする */
function syncNeedItems() {
  const n = { rows: ROWS.length, kikai: KIKAI.length, zairyo: ZAIRYO.length, shousa: SHOUSA ? 1 : 0 };
  (INIT.items || []).filter(it => it.needs).forEach(it => {
    const cb = $('item-' + it.id), lab = $('lab-' + it.id), note = $('need-' + it.id);
    if (!cb) return;
    const c = n[it.needs] || 0;
    cb.disabled = !c;
    if (!c) cb.checked = false;
    else if (!cb.dataset.touched) cb.checked = true;
    if (it.needs === 'shousa') {
      note.textContent = c ? '　指摘 ' + shousaFromTable().findings.filter(x => !x.除).length + '件'
                           : '　← 設計照査が必要';
    } else {
      note.textContent = c ? '　' + c + '件' : '　← 設計書の取り込みが必要';
    }
    lab.classList.toggle('checked', cb.checked);
    lab.classList.toggle('off', !c);
  });
}

function buildKankyo() {
  const item = (INIT.items || []).find(x => x.kankyo);
  const box = $('kankyoList'); box.innerHTML = '';
  if (!item) return;
  item.kankyo.forEach(k => {
    const cb = el('input', { type: 'checkbox', value: String(k.row), class: 'kankyo-cb' });
    box.appendChild(el('label', {}, cb, el('span', { class: 'g', text: k.group }), k.text));
  });
}

function autoKyoryoku(kind, i) {
  const name = $(kind + '-会社-' + i).value.trim();
  const hit = (INIT.company.協力会社 || []).find(x => x.会社 === name);
  if (!hit) return;
  if (!$(kind + '-担当-' + i).value) $(kind + '-担当-' + i).value = hit.担当 || '';
  if (!$(kind + '-電話-' + i).value) $(kind + '-電話-' + i).value = hit.電話 || '';
  if (kind === 'shitauke' && !$('shitauke-工種-' + i).value) $('shitauke-工種-' + i).value = hit.得意工種 || '';
}

/* ============================================================ 工事データ */
const TAISEI_KEYS = ['現場代理人', '現場代理人TEL', '主任技術者', '主任技術者TEL', '事務担当',
  '品質管理担当', '出来形管理担当', '工程管理担当', '安全管理担当',
  '統括安全衛生責任者', '安全管理者', '衛生管理者', '安全巡視員',
  '火気取締責任者', '電気取扱責任者', '車両取扱責任者', '機械取扱責任者',
  '作業指揮者1名称', '作業指揮者1', '作業指揮者2名称', '作業指揮者2',
  '事故対策指揮者', '被災者救出係', '事故現場指揮者', '現場交通管理係',
  '事故原因調査係', '情報連絡係', '災害対策係'];

function collect() {
  const v = (id) => ($(id) ? $(id).value.trim() : '');
  const taisei = {};
  TAISEI_KEYS.forEach(k => taisei[k] = v(k));
  taisei.社内検査員 = [1, 2, 3, 4].map(i => v('社内検査員' + i));

  const prj = {
    工事名: v('工事名'), 発注者: v('発注者'), 発注者課: v('発注者課'), 事業名: v('事業名'),
    工事場所: v('工事場所'), 設計書番号: v('設計書番号'), 会社名: v('会社名'),
    工事略称: v('工事略称'),
    工事概要: $('工事概要').value.replace(/\s+$/, ''),
    工期自: v('工期自'), 工期至: v('工期至'), 出力先: v('出力先'),
    体制: taisei,
    監督員: [0, 1, 2].map(i => ({ 役職: v('kantoku-yaku-' + i), 氏名: v('kantoku-name-' + i) })),
    委託: [0, 1, 2].map(i => ({ 種別: v('itaku-種別-' + i), 会社: v('itaku-会社-' + i), 担当: v('itaku-担当-' + i), 電話: v('itaku-電話-' + i) })),
    下請: [0, 1, 2, 3, 4].map(i => ({ 工種: v('shitauke-工種-' + i), 会社: v('shitauke-会社-' + i), 担当: v('shitauke-担当-' + i), 電話: v('shitauke-電話-' + i) })),
    指定機械: { あり: $('指定機械あり').checked, 文言: v('指定機械文言') },
    安全方針: [1, 2, 3, 4].map(i => v('安全方針' + i)),
    搬出先: {},
    現場環境改善: [...document.querySelectorAll('.kankyo-cb')].filter(c => c.checked).map(c => c.value),
  };
  HANSHUTSU.forEach((name, i) => {
    prj.搬出先[name] = { 処分先: v('han-処分先-' + i), 所在地: v('han-所在地-' + i), 距離: v('han-距離-' + i) };
    if (name === '残土（2）') prj.搬出先[name].内容 = v('han-内容-' + i);
  });
  return prj;
}

function applyProject(p) {
  const set = (id, val) => { if ($(id)) $(id).value = val || ''; };
  ['工事名', '発注者', '発注者課', '事業名', '工事場所', '設計書番号', '工事略称', '会社名',
    '工期自', '工期至', '出力先', '工事概要'].forEach(k => set(k, p[k]));
  const t = p.体制 || {};
  TAISEI_KEYS.forEach(k => set(k, t[k]));
  (t.社内検査員 || []).forEach((n, i) => set('社内検査員' + (i + 1), n));
  (p.監督員 || []).forEach((x, i) => { set('kantoku-yaku-' + i, x.役職); set('kantoku-name-' + i, x.氏名); });
  (p.委託 || []).forEach((x, i) => ['種別', '会社', '担当', '電話'].forEach(f => set('itaku-' + f + '-' + i, x[f])));
  (p.下請 || []).forEach((x, i) => ['工種', '会社', '担当', '電話'].forEach(f => set('shitauke-' + f + '-' + i, x[f])));
  $('指定機械あり').checked = !!(p.指定機械 || {}).あり;
  set('指定機械文言', (p.指定機械 || {}).文言);
  $('指定機械文言').disabled = !$('指定機械あり').checked;
  (p.安全方針 || []).forEach((s, i) => set('安全方針' + (i + 1), s));
  HANSHUTSU.forEach((name, i) => {
    const d = (p.搬出先 || {})[name] || {};
    set('han-処分先-' + i, d.処分先); set('han-所在地-' + i, d.所在地); set('han-距離-' + i, d.距離);
    if (name === '残土（2）') set('han-内容-' + i, d.内容);
  });
  const chosen = new Set(p.現場環境改善 || []);
  document.querySelectorAll('.kankyo-cb').forEach(c => c.checked = chosen.has(c.value));

  // 前に取り込んだ設計書の明細（保存してあれば戻す）
  if ((p._明細 || []).length) {
    ROWS = p._明細;
    drawRows();
    $('parseInfo').textContent = ROWS.length + ' 明細（保存した工事データから復元）';
  }
  if ((p._機械 || []).length || (p._資材 || []).length) {
    KIKAI = p._機械 || []; ZAIRYO = p._資材 || [];
    drawKikaiZairyo();
    $('tankaInfo').textContent = '機械 ' + KIKAI.length + ' 件／資材 ' + ZAIRYO.length +
      ' 件（保存した工事データから復元）';
  }
  if (p._照査) {
    SHOUSA = p._照査;
    drawShousa();
    $('shousaInfo').textContent = (SHOUSA.summary || '') + '（' + (SHOUSA.date || '') +
      ' の照査結果を保存データから復元）';
  }
  updateKouki();
}

/** 取り込んだ明細（工事データに一緒に保存する） */
function torikomiData() {
  const d = {};
  const r = rowsFromTable(), k = kikaiFromTable(), z = zairyoFromTable();
  if (r.length) d._明細 = r;
  if (k.length) d._機械 = k;
  if (z.length) d._資材 = z;
  if (SHOUSA) d._照査 = shousaFromTable();
  return d;
}

function updateKouki() {
  const a = $('工期自').value, b = $('工期至').value;
  $('kouki表示').textContent = (a && b) ? ('工期 ' + wareki(a) + '～' + wareki(b)) : '';
  const out = $('出力先');
  if (!out.value && INIT && INIT.settings.工事フォルダー親 && $('工事名').value) {
    out.placeholder = INIT.settings.工事フォルダー親 + '\\（工事フォルダー）\\施工計画書';
  }
}

function wareki(iso) {
  const d = new Date(iso);
  if (isNaN(d)) return '';
  const y = d.getFullYear() - 2018;
  return 'R' + y + '.' + (d.getMonth() + 1) + '.' + d.getDate();
}

/* ============================================================ 画面の配線 */
function wire() {
  document.querySelectorAll('#tabs button').forEach(b => b.addEventListener('click', () => {
    document.querySelectorAll('#tabs button').forEach(x => x.classList.toggle('on', x === b));
    document.querySelectorAll('.tab').forEach(s => s.classList.toggle('on', s.id === 'tab-' + b.dataset.tab));
  }));

  $('工期自').addEventListener('change', updateKouki);
  $('工期至').addEventListener('change', updateKouki);
  $('工事名').addEventListener('input', updateKouki);

  $('指定機械あり').addEventListener('change', e => $('指定機械文言').disabled = !e.target.checked);

  ['現場代理人', '主任技術者'].forEach(k => $(k).addEventListener('change', () => {
    const s = (INIT.staff || []).find(x => x.氏名 === $(k).value);
    if (s && s.電話 && !$(k + 'TEL').value) $(k + 'TEL').value = s.電話;
  }));

  const pick = (btn, targetId, start) => withButton(btn, '選択中…', async () => {
    const p = await openFolderPicker(start());
    if (p) $(targetId).value = p;
  });
  wireFolderPicker();
  $('btnPickOut').addEventListener('click', e =>
    pick(e.target, '出力先', () => $('出力先').value || INIT.settings.工事フォルダー親 || ''));
  $('btnPickZumen').addEventListener('click', e =>
    pick(e.target, 'zumenFolder', () => $('zumenFolder').value || INIT.settings.工事フォルダー親 || ''));
  document.querySelectorAll('button.pick').forEach(b => b.addEventListener('click', e =>
    pick(e.target, b.dataset.for, () => $(b.dataset.for).value)));

  $('btnSave').addEventListener('click', async () => {
    const p = collect();
    if (!p.工事名) { alert('工事名を入れてください'); return; }
    Object.assign(p, torikomiData());   // 取り込んだ明細も一緒に保存する
    const r = await api('/api/project', p);
    if (r.saved && ![...$('projectPicker').options].some(o => o.value === r.saved)) {
      $('projectPicker').appendChild(el('option', { value: r.saved, text: r.saved }));
    }
    $('projectPicker').value = r.saved;
    flash('保存しました：' + r.saved);
  });

  $('projectPicker').addEventListener('change', async e => {
    if (!e.target.value) return;
    const p = await api('/api/project?name=' + encodeURIComponent(e.target.value));
    applyProject(p);
    flash('読み込みました：' + e.target.value);
  });

  $('btnSaveSettings').addEventListener('click', async () => {
    const s = Object.assign({}, INIT.settings, {
      雛形フォルダー: $('set-雛形').value.trim(),
      必携フォルダー: $('set-必携').value.trim(),
      工事フォルダー親: $('set-親').value.trim(),
    });
    INIT.settings = await api('/api/settings', s);
    checkSetup();
    flash('設定を保存しました');
  });

  $('btnAddStaff').addEventListener('click', () => addStaffRow({ 氏名: '', 電話: '' }));
  $('btnSaveStaff').addEventListener('click', async () => {
    const rows = [...$('tbl-staff').querySelectorAll('tbody tr')].map(tr => ({
      氏名: tr.children[0].querySelector('input').value.trim(),
      電話: tr.children[1].querySelector('input').value.trim(),
    })).filter(x => x.氏名);
    INIT.staff = (await api('/api/staff', { staff: rows })).staff;
    document.querySelectorAll('select.staff').forEach(fillStaff);
    flash('名簿を保存しました（' + INIT.staff.length + '名）');
  });

  $('btnAddKyoryoku').addEventListener('click', () => addKyoryokuRow({}));
  $('btnSaveKyoryoku').addEventListener('click', async () => {
    const rows = [...$('tbl-kyoryoku').querySelectorAll('tbody tr')].map(tr => ({
      会社: tr.children[0].querySelector('input').value.trim(),
      担当: tr.children[1].querySelector('input').value.trim(),
      電話: tr.children[2].querySelector('input').value.trim(),
      得意工種: tr.children[3].querySelector('input').value.trim(),
    })).filter(x => x.会社);
    INIT.company.協力会社 = rows;
    await api('/api/company', INIT.company);
    fillList('dl-kaisha', rows.map(x => x.会社));
    flash('協力会社を保存しました');
  });

  $('btnBuild').addEventListener('click', doBuild);
  $('btnOpenFolder').addEventListener('click', () => api('/api/open', { path: $('出力先').value }));
  $('btnScan').addEventListener('click', e => withButton(e.target, '判別中…', doScan));
  $('btnMakeNaiyou').addEventListener('click', e => withButton(e.target, '作成中…', doMakeNaiyou));
  $('btnMakeKikai').addEventListener('click', e => withButton(e.target, '作成中…', doMakeKikai));
  $('btnMakeShousa').addEventListener('click', e => withButton(e.target, '作成中…', doMakeShousa));
  $('btnRedoShousa').addEventListener('click', e => withButton(e.target, '照査しています…', doShousa));
}

/* ============================================================ フォルダー選択 */
let FP_RESOLVE = null, FP_PARENT = '';

function openFolderPicker(start) {
  return new Promise(resolve => {
    FP_RESOLVE = resolve;
    $('fp').classList.remove('hide');
    fpGo(start || '');
  });
}

function fpClose(value) {
  $('fp').classList.add('hide');
  const r = FP_RESOLVE;
  FP_RESOLVE = null;
  if (r) r(value);
}

async function fpGo(path) {
  const r = await api('/api/listdir', { path });
  $('fp-msg').textContent = r.error || '';
  if (r.path !== undefined && r.path !== '') $('fp-path').value = r.path;
  FP_PARENT = r.parent || '';
  const list = $('fp-list');
  list.innerHTML = '';
  (r.dirs || []).forEach(d => list.appendChild(
    el('div', { class: 'fp-item', text: '📁　' + d.name, onclick: () => fpGo(d.path) })));
  if (!(r.dirs || []).length) {
    list.appendChild(el('div', { class: 'fp-item hint', text: '（この下にフォルダーはありません）' }));
  }
  const dv = $('fp-drives');
  dv.innerHTML = '';
  (r.drives || []).forEach(d => dv.appendChild(
    el('button', { class: 'ghost', text: d, onclick: () => fpGo(d) })));
}

function wireFolderPicker() {
  $('fp-up').addEventListener('click', () => fpGo(FP_PARENT || $('fp-path').value));
  $('fp-open').addEventListener('click', () => fpGo($('fp-path').value));
  $('fp-path').addEventListener('keydown', e => { if (e.key === 'Enter') fpGo($('fp-path').value); });
  $('fp-ok').addEventListener('click', () => fpClose($('fp-path').value.trim()));
  $('fp-cancel').addEventListener('click', () => fpClose(null));
  $('fp').addEventListener('click', e => { if (e.target === $('fp')) fpClose(null); });
  document.addEventListener('keydown', e => {
    if (e.key === 'Escape' && !$('fp').classList.contains('hide')) fpClose(null);
  });
}

function addStaffRow(s) {
  const tb = $('tbl-staff').querySelector('tbody');
  tb.appendChild(el('tr', {},
    el('td', {}, el('input', { value: s.氏名 || '' })),
    el('td', {}, el('input', { value: s.電話 || '' })),
    el('td', {}, el('button', { class: 'ghost', text: '削除', onclick: (e) => e.target.closest('tr').remove() }))));
}
function drawStaffTable() {
  $('tbl-staff').querySelector('tbody').innerHTML = '';
  (INIT.staff || []).forEach(addStaffRow);
}
function addKyoryokuRow(k) {
  const tb = $('tbl-kyoryoku').querySelector('tbody');
  tb.appendChild(el('tr', {},
    el('td', {}, el('input', { value: k.会社 || '' })),
    el('td', {}, el('input', { value: k.担当 || '' })),
    el('td', {}, el('input', { value: k.電話 || '' })),
    el('td', {}, el('input', { value: k.得意工種 || '' })),
    el('td', {}, el('button', { class: 'ghost', text: '削除', onclick: (e) => e.target.closest('tr').remove() }))));
}
function drawKyoryokuTable() {
  $('tbl-kyoryoku').querySelector('tbody').innerHTML = '';
  (INIT.company.協力会社 || []).forEach(addKyoryokuRow);
}

function flash(msg) {
  const st = $('buildState');
  st.textContent = msg;
  setTimeout(() => { if (st.textContent === msg) st.textContent = ''; }, 4000);
}

/* ============================================================ 出力 */
async function doBuild() {
  const prj = collect();
  if (!prj.工事名) { alert('工事名を入れてください'); return; }
  if (!prj.出力先) { alert('出力先フォルダーを選んでください'); return; }
  const ids = (INIT.items || []).filter(it => $('item-' + it.id) && $('item-' + it.id).checked)
    .map(it => it.id);
  if (!ids.length) { alert('作る項目を選んでください'); return; }

  document.querySelectorAll('#tabs button').forEach(x => x.classList.toggle('on', x.dataset.tab === 'shuturyoku'));
  document.querySelectorAll('.tab').forEach(s => s.classList.toggle('on', s.id === 'tab-shuturyoku'));

  $('btnBuild').disabled = true;
  $('log').textContent = '';
  $('shots').innerHTML = '';
  $('buildState').textContent = 'Excel を動かしています…（1〜2分かかります。Excelの画面は出ません）';

  // （1）（1-2）（5）（6）は設計書から取り込んだ明細で作るので、一緒に送る
  const r = await api('/api/build', {
    project: prj, ids, verify: $('verify').checked,
    rows: rowsFromTable(), kikai: kikaiFromTable(), zairyo: zairyoFromTable(),
    shousa: SHOUSA ? shousaFromTable() : null,
  });
  if (r.error) { $('btnBuild').disabled = false; $('buildState').textContent = r.error; return; }
  poll();
}

async function poll() {
  const j = await api('/api/job');
  $('log').textContent = (j.log || []).join('\n');
  $('log').scrollTop = $('log').scrollHeight;
  if (j.running) { setTimeout(poll, 700); return; }

  $('btnBuild').disabled = false;
  if (j.error) {
    $('buildState').innerHTML = '<span class="msg-ng">失敗しました：' + escapeHtml(j.error) + '</span>';
    return;
  }
  const res = j.result || {};
  let msg = '<span class="msg-ok">完了：' + (res.files || []).length + ' ファイルを出力しました</span>';
  (res.warnings || []).forEach(w => { msg += '<br><span class="msg-ng">' + escapeHtml(w) + '</span>'; });
  $('buildState').innerHTML = msg;
  (res.images || []).forEach(p => {
    $('shots').appendChild(el('figure', {},
      el('img', { src: '/api/image?path=' + encodeURIComponent(p), loading: 'lazy' }),
      el('figcaption', { text: p.split('\\').pop() })));
  });
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

/* ============================================================ 設計図書（段階2） */
async function doScan() {
  const folder = $('zumenFolder').value.trim();
  if (!folder) { alert('フォルダーを選んでください'); return; }
  $('scanResult').innerHTML = '<p class="hint">読んでいます…</p>';
  const r = await api('/api/scan', { folder });
  if (r.error) { $('scanResult').innerHTML = '<p class="msg-ng">' + escapeHtml(r.error) + '</p>'; return; }
  SCAN_FILES = r.files || [];

  const box = el('div');

  // 見積参考資料が1つ見つかれば、工種・数量と機械・資材をまとめて取り込む
  const best = (r.files || []).find(f => f.内訳あり && f.単価表あり) ||
               (r.files || []).find(f => f.kind === '見積参考資料' && f.文字あり);
  if (best) {
    box.appendChild(el('div', { class: 'card' },
      el('div', {}, el('b', { text: 'この資料からまとめて取り込む' }),
        el('span', { class: 'hint', text: '　' + best.name })),
      el('div', { class: 'hint', text: '工種・数量（1）（1-2）と 機械・資材（5）（6）を読み取り、続けて設計照査（資料どうしの照合）も行います。1〜2分かかります。' }),
      el('div', { class: 'row' },
        el('button', {
          text: 'まとめて読み取る',
          onclick: (e) => withButton(e.target, '読んでいます…', () => doTorikomi(best.path))
        }),
        el('span', { class: 'hint', id: 'toriInfo' }))));
  }

  if (INIT.shousa && SCAN_FILES.length) {
    box.appendChild(el('div', { class: 'row' },
      el('button', {
        class: 'ghost', text: '設計照査だけ行う',
        onclick: (e) => withButton(e.target, '照査しています…', doShousa)
      }),
      el('span', { class: 'hint', text: '見つかった資料どうしを照合します（数量・単位・明細書の参照・工事名・工期）' })));
  }

  (r.files || []).forEach(f => {
    const card = el('div', { class: 'card' },
      el('div', {}, el('span', { class: 'tag ' + (f.kind === '不明' ? 'warn' : 'ok'), text: f.kind }),
        el('b', { text: f.name }), el('span', { class: 'hint', text: '　' + f.pages + 'ページ' })));
    const btns = el('div', { class: 'row' });
    if (f.内訳あり) {
      btns.appendChild(el('button', {
        class: 'ghost', text: 'この資料から工種・数量を読み取る',
        onclick: (e) => withButton(e.target, '読んでいます…', () => doParse(f.path))
      }));
    }
    if (f.単価表あり || f.文字あり) {
      btns.appendChild(el('button', {
        class: 'ghost',
        text: f.単価表あり ? '施工単価表から機械・資材を拾う' : '機械・資材を拾ってみる',
        onclick: (e) => withButton(e.target, '読んでいます…', () => doTanka(f.path))
      }));
    }
    if (btns.childNodes.length) card.appendChild(btns);
    box.appendChild(card);
  });
  if (r.meta && Object.keys(r.meta).length) {
    const m = r.meta;
    const card = el('div', { class: 'card' }, el('b', { text: '読み取れた工事情報' }));
    Object.entries(m).forEach(([k, v]) => card.appendChild(el('div', { text: k + '：' + v })));
    card.appendChild(el('div', { class: 'row' },
      el('button', { class: 'ghost', text: '工事情報タブに反映する', onclick: () => { applyMeta(m); flash('反映しました'); } })));
    box.appendChild(card);
  }
  $('scanResult').innerHTML = '';
  $('scanResult').appendChild(box);
}

function applyMeta(m) {
  const map = { 工事名: '工事名', 工事場所: '工事場所', 設計書番号: '設計書番号',
                発注者: '発注者', 事業名: '事業名' };
  Object.entries(map).forEach(([k, id]) => { if (m[k] && !$(id).value) $(id).value = m[k]; });
}

/** 見積参考資料1つから（1）（1-2）と（5）（6）の材料をまとめて取り込む */
async function doTorikomi(file) {
  const info = $('toriInfo');
  if (info) info.textContent = '工種・数量を読んでいます…';
  await doParse(file);
  if (info) info.textContent = '機械・資材を読んでいます…';
  await doTanka(file);
  let shousaMsg = '';
  if (INIT.shousa) {
    if (info) info.textContent = '設計照査（資料どうしの照合）をしています…';
    await doShousa();
    if (SHOUSA) shousaMsg = '／設計照査 ' + (SHOUSA.summary || '');
  }
  if (info) {
    info.textContent = '工種・数量 ' + ROWS.length + ' 明細／機械 ' + KIKAI.length +
      ' 件／資材 ' + ZAIRYO.length + ' 件を取り込みました' + shousaMsg +
      '。「作る項目」で（1）（1-2）（5）（6）' + (SHOUSA ? '・設計照査結果' : '') + 'が選べます';
  }
}

/* ============================================================ 設計照査 */
let SCAN_FILES = [], SHOUSA = null;

async function doShousa() {
  if (!SCAN_FILES.length) { alert('先に「判別する」を押してください'); return; }
  $('shousaArea').classList.remove('hide');
  $('shousaInfo').textContent = '照査しています…';
  const r = await api('/api/shousa', { files: SCAN_FILES, project: collect() });
  if (r.error) { $('shousaInfo').textContent = r.error; return; }
  SHOUSA = r;
  $('shousaInfo').textContent = r.summary + '（' + r.docs.length + ' 資料を照合）';
  drawShousa();
}

function drawShousa() {
  if (!SHOUSA) return;
  $('shousaArea').classList.remove('hide');
  const tb = $('tbl-shousa').querySelector('tbody'); tb.innerHTML = '';
  const fs = SHOUSA.findings || [];
  if (!fs.length) {
    tb.appendChild(el('tr', {}, el('td', { colspan: '6', class: 'hint',
      text: '資料の照合では指摘事項はありませんでした。現地・図面は Excel の照査チェックリストで確かめてください' })));
  }
  fs.forEach((x, i) => {
    const cb = el('input', { type: 'checkbox', id: 'sdel-' + i });
    cb.checked = !!x.除;
    cb.addEventListener('change', syncNeedItems);
    tb.appendChild(el('tr', { class: x.重要度 === '要確認' ? 'ng' : '' },
      el('td', {}, cb),
      el('td', {}, el('span', { class: 'tag ' + (x.重要度 === '要確認' ? 'warn' : ''), text: x.重要度 })),
      el('td', { text: x.区分 || '' }),
      el('td', { text: x.対象 || '' }),
      el('td', { text: x.内容 || '' }),
      el('td', { class: 'hint', text: x.資料 || '' })));
  });
  $('shousaChecked').innerHTML = '';
  if ((SHOUSA.checked || []).length) {
    $('shousaChecked').appendChild(el('div', { text: '実施した照合：' }));
    SHOUSA.checked.forEach(t => $('shousaChecked').appendChild(el('div', { text: '・' + t })));
  }
  syncNeedItems();
}

/** 画面で「除」を付けた指摘を反映した照査結果（保存・出力用） */
function shousaFromTable() {
  if (!SHOUSA) return null;
  const r = Object.assign({}, SHOUSA);
  r.findings = (SHOUSA.findings || []).map((x, i) => {
    const y = Object.assign({}, x);
    const cb = $('sdel-' + i);
    if (cb ? cb.checked : x.除) y.除 = true; else delete y.除;
    return y;
  });
  return r;
}

async function doMakeShousa() {
  const prj = collect();
  if (!prj.出力先) { alert('出力先フォルダーを選んでください（工事情報タブ）'); return; }
  if (!SHOUSA) { alert('先に設計照査を行ってください'); return; }
  $('shousaInfo').textContent = '作っています…';
  const r = await api('/api/shousa_xlsx', { project: prj, result: shousaFromTable() });
  $('shousaInfo').textContent = r.error ? ('失敗：' + r.error) : ('作成：' + (r.files || []).join(' / '));
  if (r.error) alert('設計照査結果を作れませんでした： ' + r.error);
}

let ROWS = [];
function drawRows() {
  if (ROWS.length) $('parseArea').classList.remove('hide');
  const tb = $('tbl-rows').querySelector('tbody'); tb.innerHTML = '';
  ROWS.forEach((row, i) => {
    tb.appendChild(el('tr', {},
      el('td', {}, el('input', { type: 'checkbox', id: 'del-' + i })),
      el('td', {}, el('input', { id: 'lv-' + i, value: row.level || '' })),
      el('td', {}, el('input', { id: 'nm-' + i, value: row.name || '' })),
      el('td', {}, el('input', { id: 'sp-' + i, value: row.spec || '' })),
      el('td', {}, el('input', { id: 'qt-' + i, value: row.qty === null || row.qty === undefined ? '' : row.qty })),
      el('td', {}, el('input', { id: 'un-' + i, value: row.unit || '' }))));
  });
  syncNeedItems();
}

async function doParse(file) {
  $('parseArea').classList.remove('hide');   // 先に出しておかないと経過が見えない
  $('parseInfo').textContent = '読んでいます…（大きい資料は1分ほどかかります）';
  const r = await api('/api/parse', { file });
  if (r.error) { $('parseInfo').textContent = r.error; alert(r.error); return; }
  ROWS = r.rows || [];
  $('parseInfo').textContent = r.rows.length + ' 明細（' + (r.source || '') + '）';
  drawRows();
  if (r.meta) applyMeta(r.meta);
}

let KIKAI = [], ZAIRYO = [];
function drawKikaiZairyo() {
  if (KIKAI.length || ZAIRYO.length) $('tankaArea').classList.remove('hide');
  const tk = $('tbl-kikai').querySelector('tbody'); tk.innerHTML = '';
  KIKAI.forEach((x, i) => tk.appendChild(el('tr', {},
    el('td', {}, el('input', { type: 'checkbox', id: 'kdel-' + i })),
    el('td', {}, el('input', { id: 'knm-' + i, value: x.name || '' })),
    el('td', {}, el('input', { id: 'ksp-' + i, value: x.spec || '' })),
    el('td', {}, el('input', { id: 'kun-' + i, value: x.unit || '台' })),
    el('td', {}, el('input', { id: 'kqt-' + i, value: x.qty === null ? '' : x.qty })),
    el('td', {}, el('input', { id: 'kwk-' + i, value: x.work || '' })),
    el('td', {}, el('input', { id: 'kln-' + i, value: x.low_noise || '' })),
    el('td', {}, el('input', { id: 'kem-' + i, value: x.emission || '' })))));

  const tz = $('tbl-zairyo').querySelector('tbody'); tz.innerHTML = '';
  ZAIRYO.forEach((x, i) => tz.appendChild(el('tr', {},
    el('td', {}, el('input', { type: 'checkbox', id: 'zdel-' + i })),
    el('td', {}, el('input', { id: 'znm-' + i, value: x.name || '' })),
    el('td', {}, el('input', { id: 'zsp-' + i, value: x.spec || '' })),
    el('td', {}, el('input', { id: 'zun-' + i, value: x.unit || '' })),
    el('td', {}, el('input', { id: 'zqt-' + i, value: x.qty === null || x.qty === undefined ? '' : x.qty })),
    el('td', {}, el('input', { id: 'zwk-' + i, value: x.work || '' })),
    el('td', { class: 'hint', title: x.trail || '' },
      (x.estimate ? '※ ' : '') + (x.basis || '')))));
  syncNeedItems();
}

async function doTanka(file) {
  $('tankaArea').classList.remove('hide');   // 先に出しておかないと経過が見えない
  $('tankaInfo').textContent = '読んでいます…（大きい資料は1分ほどかかります）';
  const r = await api('/api/tanka', { file });
  if (r.error) { $('tankaInfo').textContent = r.error; alert(r.error); return; }
  KIKAI = r.kikai || []; ZAIRYO = r.zairyo || [];
  $('tankaInfo').textContent = '機械 ' + KIKAI.length + ' 件／資材 ' + ZAIRYO.length + ' 件（' + r.source + '）';
  drawKikaiZairyo();
}

/* 画面の表で直した値を読み出す。出力タブの一括生成と、取り込みタブのボタンで共用 */
function rowsFromTable() {
  return ROWS.map((_, i) => ({
    level: $('lv-' + i).value.trim(), name: $('nm-' + i).value.trim(), spec: $('sp-' + i).value.trim(),
    qty: $('qt-' + i).value.trim(), unit: $('un-' + i).value.trim(), del: $('del-' + i).checked,
  })).filter(r => !r.del && r.name);
}

function kikaiFromTable() {
  return KIKAI.map((_, i) => ({
    name: $('knm-' + i).value.trim(), spec: $('ksp-' + i).value.trim(),
    unit: $('kun-' + i).value.trim(), qty: $('kqt-' + i).value.trim(),
    work: $('kwk-' + i).value.trim(), low_noise: $('kln-' + i).value.trim(),
    emission: $('kem-' + i).value.trim(), del: $('kdel-' + i).checked,
  })).filter(r => !r.del && r.name);
}

function zairyoFromTable() {
  return ZAIRYO.map((x, i) => ({
    name: $('znm-' + i).value.trim(), spec: $('zsp-' + i).value.trim(),
    unit: $('zun-' + i).value.trim(), qty: $('zqt-' + i).value.trim(),
    work: $('zwk-' + i).value.trim(), del: $('zdel-' + i).checked,
    basis: x.basis || '', trail: x.trail || '', estimate: !!x.estimate,   // 算出根拠に残す
  })).filter(r => !r.del && r.name);
}

async function doMakeKikai() {
  const prj = collect();
  if (!prj.出力先) { alert('出力先フォルダーを選んでください（工事情報タブ）'); return; }
  const kikai = kikaiFromTable();
  const zairyo = zairyoFromTable();
  if (!kikai.length && !zairyo.length) { alert('明細がありません'); return; }
  $('tankaInfo').textContent = '作っています…';
  const r = await api('/api/kikai', { project: prj, kikai, zairyo });
  $('tankaInfo').textContent = r.error ? ('失敗：' + r.error) : ('作成：' + (r.files || []).join(' / '));
  if (r.error) alert('（5）（6）を作れませんでした： ' + r.error);
}

async function doMakeNaiyou() {
  const prj = collect();
  if (!prj.出力先) { alert('出力先フォルダーを選んでください（工事情報タブ）'); return; }
  const rows = rowsFromTable();
  if (!rows.length) { alert('明細がありません'); return; }
  $('parseInfo').textContent = '作っています…';
  const r = await api('/api/naiyou', { project: prj, rows, targets: ['1', '1-2'] });
  $('parseInfo').textContent = r.error ? ('失敗：' + r.error) : ('作成：' + (r.files || []).join(' / '));
  if (r.error) alert('（1）（1-2）を作れませんでした： ' + r.error);
}

init();
