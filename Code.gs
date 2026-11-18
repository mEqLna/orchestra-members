/**
 * 演奏メンバー参加登録：Googleスプレッドシート連携用 Apps Script
 *
 * 使い方
 * 1. 対象スプレッドシートで［拡張機能］＞［Apps Script］を開き、このコードを全部貼り付けて保存
 * 2. 関数 setup を選んで［実行］（初回のみ権限の承認が必要）。「登録データ」「集計」シートができます
 * 3. ［デプロイ］＞［新しいデプロイ］＞種類「ウェブアプリ」
 *      次のユーザーとして実行：自分
 *      アクセスできるユーザー：全員
 * 4. 発行された「ウェブアプリのURL」を、HTML側の API_URL に貼り付ける
 *
 * 曲・楽器を変更したら、このファイルとHTMLの両方の SONGS / INSTRUMENTS を同じ内容にそろえてください。
 * 曲を追加したあとは setup を再実行すると、集計シートが作り直されます（登録データは消えません）。
 */

const DATA = '登録データ';
const SUMMARY = '集計';
const SETTINGS = 'グループ設定'; // 楽器のグループ組み替え（全員で共有）
const ORDER = '表示順'; // グループ・楽器の表示順（全員で共有）
const ALIASES = '名称変更'; // 楽器名・グループ名の変更（全員で共有）

const SONGS = [
  { id: 'M1', title: 'A. リード：セカンド・センチュリー' },
  { id: 'M2', title: 'A. リード：アルメニアン・ダンス パートI' },
  { id: 'M3', title: 'A. リード：バラード ～アルト・サクソフォーンとバンドのための～' },
  { id: 'M4', title: 'A. リード：エル・カミーノ・レアル ～ラテン・ファンタジー～' },
  { id: 'M5', title: 'K. J. アルフォード：行進曲「ボギー大佐」' },
  { id: 'M6', title: 'J. F. ワーグナー：行進曲「双頭の鷲の旗の下に」' },
  { id: 'M7', title: '小山清茂：吹奏楽のための木挽歌' },
  { id: 'M8', title: 'ジャパニーズ・グラフィティII ～坂本九メモリアル～' },
  { id: 'M9', title: 'ジブリ名曲集' }
];

const INSTRUMENTS = [
  'ピッコロ（Piccolo）', 'フルート（Flute）', 'オーボエ（Oboe）', 'ファゴット（Bassoon）',
  'E♭クラリネット（E♭ Clarinet）', 'B♭クラリネット（B♭ Clarinet）', 'アルトクラリネット（Alto Clarinet）',
  'バスクラリネット（Bass Clarinet）', 'ソプラノサクソフォン（Soprano Saxophone）', 'アルトサクソフォン（Alto Saxophone）',
  'テナーサクソフォン（Tenor Saxophone）', 'バリトンサクソフォン（Baritone Saxophone）', 'トランペット（Trumpet）',
  'コルネット（Cornet）', 'ホルン（Horn in F / E♭）', 'トロンボーン（Trombone）', 'バストロンボーン（Bass Trombone）',
  'ユーフォニアム（Euphonium）', 'チューバ（Tuba）', '打楽器（Percussion）', 'コントラバス（Contrabass / String Bass）'
];

// 集計シートの階層。items は INSTRUMENTS の番号
const GROUPS = [
  { name: 'フルート系', items: [0, 1] },
  { name: 'ダブルリード', items: [2, 3] }, // オーボエ・ファゴット
  { name: 'クラリネット', items: [4, 5, 6, 7] },
  { name: 'サックス', items: [8, 9, 10, 11] },
  { name: 'トランペット系', items: [12, 13] },
  { name: 'ホルン', items: [14] },
  { name: 'トロンボーン系', items: [15, 16] },
  { name: 'ユーフォニアム', items: [17] },
  { name: 'チューバ', items: [18] },
  { name: '打楽器', items: [19] },
  { name: 'コントラバス', items: [20] }
];

const OTHER_GROUP = 'その他の楽器';
const GROUP_NAMES = GROUPS.map(g => g.name).concat([OTHER_GROUP]);

// ===== 初期設定（シートと集計式を作る） =====
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let d = ss.getSheetByName(DATA);
  if (!d) {
    const first = ss.getSheets()[0];
    if (ss.getSheets().length === 1 && first.getLastRow() === 0) {
      first.setName(DATA);
      d = first;
    } else {
      d = ss.insertSheet(DATA);
    }
  }
  // 列：名前, 主な楽器, 各曲（その曲の担当楽器名。空欄＝不参加）, 更新日時, 区分
  const header = ['名前', '楽器'].concat(SONGS.map(s => s.id), ['更新日時', '区分']);
  d.getRange(1, 1, 1, header.length).setValues([header])
    .setFontWeight('bold').setBackground('#f0ede4');
  // 名前・楽器・各曲の担当楽器は常に文字として扱う（=で始まる入力の誤作動防止）
  d.getRange(1, 1, d.getMaxRows(), 2 + SONGS.length).setNumberFormat('@');
  d.getRange(2, 3 + SONGS.length, d.getMaxRows() - 1, 1).setNumberFormat('yyyy/mm/dd hh:mm');
  d.getRange(1, 4 + SONGS.length, d.getMaxRows(), 1).setNumberFormat('@');
  backfill_(d);
  d.setFrozenRows(1);
  d.setColumnWidth(1, 160);
  d.setColumnWidth(2, 260);

  ensureSettings_(ss);
  ensureOrder_(ss);
  ensureAliases_(ss);
  buildSummary_(ss);
  SpreadsheetApp.getUi().alert('設定が完了しました。続いて［デプロイ］＞［新しいデプロイ］からウェブアプリとして公開してください。');
}

function col_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function backfill_(d) {
  // 区分の列と、各曲の「1/0」→「担当楽器名」への置き換えを行う（既存データ用）
  const nS = SONGS.length;
  const n = d.getLastRow() - 1;
  if (n < 1) return;
  const rng = d.getRange(2, 1, n, 4 + nS);
  const vals = rng.getValues();
  let changed = false;
  vals.forEach(r => {
    const inst = String(r[1]).trim();
    // 旧形式：1＝参加（主な楽器）、0や空欄＝不参加
    for (let i = 0; i < nS; i++) {
      const v = String(r[2 + i]).trim();
      if (v === '1') { r[2 + i] = inst; changed = true; }
      else if (v === '0') { r[2 + i] = ''; changed = true; }
    }
    const custom = inst !== '' && INSTRUMENTS.indexOf(inst) < 0;
    const g = String(r[3 + nS] || '');
    if (custom && g === '') { r[3 + nS] = OTHER_GROUP; changed = true; }
    if (!custom && g !== '') { r[3 + nS] = ''; changed = true; }
  });
  if (changed) rng.setValues(vals);
}

function ensureSettings_(ss) {
  let g = ss.getSheetByName(SETTINGS);
  if (!g) g = ss.insertSheet(SETTINGS);
  g.getRange(1, 1, 1, 3).setValues([['楽器', 'グループ', '更新日時']])
    .setFontWeight('bold').setBackground('#f0ede4');
  g.getRange('A:B').setNumberFormat('@');
  g.setFrozenRows(1);
  g.setColumnWidth(1, 260);
  g.setColumnWidth(2, 160);
}

function buildSummary_(ss) {
  const s = ss.getSheetByName(SUMMARY) || ss.insertSheet(SUMMARY);
  s.clear();
  const nS = SONGS.length;
  const first = col_(3), last = col_(2 + nS);
  const grpCol = col_(4 + nS);
  const songRange = `'${DATA}'!$${first}$2:$${last}`;
  const hdrRange = `'${DATA}'!$${first}$1:$${last}$1`;
  const nameRange = `'${DATA}'!$A$2:$A`;
  const instRange = `'${DATA}'!$B$2:$B`;
  const grpRange = `'${DATA}'!$${grpCol}$2:$${grpCol}`;
  const setRange = `'${SETTINGS}'!$A$2:$B`;
  const aliasInst = `'${ALIASES}'!$A$2:$B`;
  const aliasGrp = `'${ALIASES}'!$D$2:$E`;
  const bold = '#f7f5ee';
  const songCols = SONGS.map((_, j) => 4 + j);
  const lastCol = 3 + nS;

  // 行の配置を先に決める（表どうしが参照し合うため）
  const NEW_SLOTS = 10; // 新しいグループの表示枠
  const gt = 3 + nS + 2;                 // 「グループごとの合計人数」の見出し行
  const hrG = gt + 1;
  const gStart = hrG + 1;
  const gEnd = gStart + GROUP_NAMES.length + NEW_SLOTS - 1;
  const gTotal = gEnd + 1;
  const it = gTotal + 2;                 // 「楽器ごとの人数」の見出し行
  const hrI = it + 1;
  const iStart = hrI + 1;
  const iEnd = iStart + INSTRUMENTS.length - 1;
  const ct = iEnd + 2;                   // 「追加された楽器」の見出し行
  const hrC = ct + 1;
  const dr = hrC + 1;

  const stGrp = `$A$${iStart}:$A$${iEnd}`;
  const stCnt = `$C$${iStart}:$C$${iEnd}`;
  const cuGrp = `$B$${dr}:$B`;
  const cuCnt = `$C$${dr}:$C`;
  const stSong = L => `$${L}$${iStart}:$${L}$${iEnd}`;
  const cuSong = L => `$${L}$${dr}:$${L}`;
  const idxOf = (L, hr) => `INDEX(${songRange},0,MATCH(${L}$${hr},${hdrRange},0))`;
  const songIds = SONGS.map(x => x.id);

  // --- 曲ごとの参加人数 ---
  s.getRange(1, 1).setValue('曲ごとの参加人数').setFontWeight('bold').setFontSize(12);
  s.getRange(2, 1, 1, 3).setValues([['曲ID', '曲名', '参加人数']]).setFontWeight('bold').setBackground('#f0ede4');
  SONGS.forEach((song, i) => {
    const r = 3 + i;
    s.getRange(r, 1, 1, 2).setValues([[song.id, song.title]]);
    s.getRange(r, 3).setFormula(`=COUNTA(INDEX(${songRange},0,MATCH($A${r},${hdrRange},0)))`);
  });
  const sumRow = 3 + nS;
  s.getRange(sumRow, 1).setValue('登録者数');
  s.getRange(sumRow, 3).setFormula(`=COUNTA(${nameRange})`);
  s.getRange(sumRow, 1, 1, 3).setFontWeight('bold').setBackground(bold);

  // --- グループごとの合計人数（組み替え・新しく作られたグループを反映） ---
  s.getRange(gt, 1).setValue('グループごとの合計人数（組み替え・表示順を反映）').setFontWeight('bold').setFontSize(12);
  s.getRange(hrG, 1, 1, lastCol).setValues([['グループ（元の名前）', '表示名', '人数'].concat(songIds)])
    .setFontWeight('bold').setBackground('#f0ede4');
  const groupFormulas = (r) => {
    s.getRange(r, 3).setFormula(`=SUMIF(${stGrp},$A${r},${stCnt})+SUMIF(${cuGrp},$A${r},${cuCnt})`);
    songCols.forEach(c => {
      const L = col_(c);
      s.getRange(r, c).setFormula(`=SUMIF(${stGrp},$A${r},${stSong(L)})+SUMIF(${cuGrp},$A${r},${cuSong(L)})`);
    });
  };
  // グループの並びは「表示順」シートを優先し、載っていないものは標準の順、新しいグループは最後
  const ordG = `'${ORDER}'!$A$2:$A`;
  const flat = '{' + GROUP_NAMES.map(n => '"' + n + '"').join(';') + ';' + stGrp + ';' + cuGrp + '}';
  const groupList =
    `LET(cur,FILTER(${flat},${flat}<>""),o,IFERROR(FILTER(${ordG},ISNUMBER(MATCH(${ordG},cur,0))),""),` +
    `u,UNIQUE({o;cur}),FILTER(u,u<>""))`;
  for (let k = 1; k <= GROUP_NAMES.length + NEW_SLOTS; k++) {
    const r = gStart + k - 1;
    s.getRange(r, 1).setFormula(`=IFERROR(INDEX(${groupList},${k}),"")`);
    s.getRange(r, 2).setFormula(`=IF($A${r}="","",IFERROR(VLOOKUP($A${r},${aliasGrp},2,FALSE),$A${r}))`);
    s.getRange(r, 3).setFormula(
      `=IF($A${r}="","",SUMIF(${stGrp},$A${r},${stCnt})+SUMIF(${cuGrp},$A${r},${cuCnt}))`);
    songCols.forEach(c => {
      const L = col_(c);
      s.getRange(r, c).setFormula(
        `=IF($A${r}="","",SUMIF(${stGrp},$A${r},${stSong(L)})+SUMIF(${cuGrp},$A${r},${cuSong(L)}))`);
    });
  }
  s.getRange(gTotal, 1).setValue('合計');
  for (let c = 3; c <= lastCol; c++) {
    const L = col_(c);
    s.getRange(gTotal, c).setFormula(`=SUM(${L}${gStart}:${L}${gEnd})`);
  }
  s.getRange(gTotal, 1, 1, lastCol).setFontWeight('bold').setBackground('#faf8f2');

  // --- 楽器ごとの人数（標準の楽器。グループは組み替えを反映） ---
  s.getRange(it, 1).setValue('楽器ごとの人数（主な楽器）と曲別の内訳（曲ごとの担当楽器）').setFontWeight('bold').setFontSize(12);
  s.getRange(hrI, 1, 1, lastCol + 1).setValues([['現在のグループ', '楽器', '人数'].concat(songIds, ['表示名'])])
    .setFontWeight('bold').setBackground('#f0ede4');
  let r = iStart;
  GROUPS.forEach(g => {
    g.items.forEach(i => {
      s.getRange(r, 2).setValue(INSTRUMENTS[i]);
      s.getRange(r, 1).setFormula(`=IFERROR(VLOOKUP($B${r},${setRange},2,FALSE),"${g.name}")`);
      s.getRange(r, 3).setFormula(`=COUNTIF(${instRange},$B${r})`);
      s.getRange(r, lastCol + 1).setFormula(`=IFERROR(VLOOKUP($B${r},${aliasInst},2,FALSE),$B${r})`);
      songCols.forEach(c => {
        const L = col_(c);
        s.getRange(r, c).setFormula(`=COUNTIF(${idxOf(L, hrI)},$B${r})`);
      });
      r++;
    });
  });

  // --- 追加された楽器（標準の一覧にない楽器。登録のたびに自動で増減する） ---
  s.getRange(ct, 1).setValue('追加された楽器の人数と曲別の内訳').setFontWeight('bold').setFontSize(12);
  s.getRange(hrC, 1, 1, lastCol + 1).setValues([['楽器', '現在のグループ', '人数'].concat(songIds, ['表示名'])])
    .setFontWeight('bold').setBackground('#f0ede4');
  const stdList = '{' + INSTRUMENTS.map(n => '"' + n + '"').join(';') + '}';
  s.getRange(dr, 1).setFormula(
    `=IFERROR(LET(u,UNIQUE(FILTER(${instRange},${instRange}<>"")),FILTER(u,ISNA(MATCH(u,${stdList},0)))),"")`);
  s.getRange(dr, 2).setFormula(
    `=ARRAYFORMULA(IF($A${dr}:$A="","",IFERROR(VLOOKUP($A${dr}:$A,${setRange},2,FALSE),` +
    `IFERROR(VLOOKUP($A${dr}:$A,{${instRange},${grpRange}},2,FALSE),"${OTHER_GROUP}"))))`);
  s.getRange(dr, 3).setFormula(
    `=ARRAYFORMULA(IF($A${dr}:$A="","",COUNTIF(${instRange},$A${dr}:$A)))`);
  s.getRange(dr, lastCol + 1).setFormula(
    `=ARRAYFORMULA(IF($A${dr}:$A="","",IFERROR(VLOOKUP($A${dr}:$A,${aliasInst},2,FALSE),$A${dr}:$A)))`);
  songCols.forEach(c => {
    const L = col_(c);
    s.getRange(dr, c).setFormula(
      `=ARRAYFORMULA(IF($A${dr}:$A="","",COUNTIF(${idxOf(L, hrC)},$A${dr}:$A)))`);
  });

  // 見た目：人数の列は中央、見出し行は左寄せ＋人数以降は中央
  s.getRange(1, 3, s.getMaxRows(), lastCol - 2).setHorizontalAlignment('center');
  [hrG, hrI, hrC].forEach(h => {
    s.getRange(h, 1, 1, 2).setHorizontalAlignment('left');
    s.getRange(h, 3, 1, lastCol - 2).setHorizontalAlignment('center');
  });
  s.setColumnWidth(1, 160);
  s.setColumnWidth(2, 300);
  s.setColumnWidth(3, 80);
  s.setColumnWidth(lastCol + 1, 220);
}

// ===== ウェブアプリ（HTMLから呼ばれる） =====
function doGet() {
  return json_({ ok: true, members: readAll_(), groups: readOverrides_(), order: readOrder_(), aliases: readAliases_() });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    if (req.action === 'save') save_(req.member || {});
    else if (req.action === 'delete') del_(String(req.name || ''));
    else if (req.action === 'setGroup') setGroup_(String(req.inst || ''), String(req.grp || ''));
    else if (req.action === 'setOrder') setOrder_(String(req.kind || ''), req.list);
    else if (req.action === 'setAlias') setAlias_(String(req.kind || ''), String(req.key || ''), String(req.name || ''));
    else throw new Error('不明な操作です');
    return json_({ ok: true, members: readAll_(), groups: readOverrides_(), order: readOrder_(), aliases: readAliases_() });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function dataSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(DATA);
}

function readAll_() {
  const d = dataSheet_();
  const n = d.getLastRow() - 1;
  if (n < 1) return [];
  const nS = SONGS.length;
  const vals = d.getRange(2, 1, n, 4 + nS).getValues();
  return vals
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({
      name: String(r[0]),
      inst: String(r[1]),
      grp: String(r[3 + nS] || ''),
      songs: SONGS.filter((s, i) => String(r[2 + i]).trim() !== '').map(s => s.id),
      songInst: songInstOf_(r, String(r[1]))
    }));
}

function findRow_(d, name) {
  const last = d.getLastRow();
  if (last < 2) return -1;
  const names = d.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < names.length; i++) {
    if (String(names[i][0]).trim() === name) return i + 2;
  }
  return -1;
}

function save_(m) {
  const name = String(m.name || '').trim().slice(0, 50);
  // 楽器名は標準の一覧にないもの（各自が追加したもの）も受け付ける
  const inst = String(m.inst || '').replace(/[\u0000-\u001f]/g, '').trim().replace(/\s+/g, ' ').slice(0, 60);
  if (!name || !inst) {
    throw new Error('名前または楽器が正しくありません');
  }
  const songs = (Array.isArray(m.songs) ? m.songs : []).map(String);
  // 曲ごとの担当楽器（指定がなければ主な楽器）
  const si = (m.songInst && typeof m.songInst === 'object') ? m.songInst : {};
  const instFor = id => (songs.indexOf(id) < 0) ? '' : (clean_(si[id], 60) || inst);
  // 追加された楽器だけ、入れるグループを記録する（標準の楽器は空欄）
  const isCustom = INSTRUMENTS.indexOf(inst) < 0;
  // グループ名は標準のものに限らず、新しく作られたものも受け付ける
  const g = String(m.grp || '').replace(/[\u0000-\u001f]/g, '').trim().replace(/\s+/g, ' ').slice(0, 20);
  const grp = isCustom ? (g || OTHER_GROUP) : '';
  const row = [name, inst]
    .concat(SONGS.map(s => instFor(s.id)), [new Date(), grp]);
  const d = dataSheet_();
  let target = findRow_(d, name);
  if (target < 0) target = Math.max(d.getLastRow(), 1) + 1;
  d.getRange(target, 1, 1, row.length).setValues([row]);
}

function del_(name) {
  const d = dataSheet_();
  const row = findRow_(d, name.trim());
  if (row > 0) d.deleteRow(row);
}

// ===== 楽器のグループ組み替え（全員で共有。「グループ設定」シートに保存） =====
function settingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(SETTINGS) || (ensureSettings_(ss), ss.getSheetByName(SETTINGS));
}

function readOverrides_() {
  const g = settingsSheet_();
  const n = g.getLastRow() - 1;
  const map = {};
  if (n < 1) return map;
  g.getRange(2, 1, n, 2).getValues().forEach(r => {
    const inst = String(r[0]).trim(), grp = String(r[1]).trim();
    if (inst && grp) map[inst] = grp;
  });
  return map;
}

function clean_(v, max) {
  return String(v || '').replace(/[\u0000-\u001f]/g, '').trim().replace(/\s+/g, ' ').slice(0, max);
}

// grp が空なら組み替えを解除（元のグループに戻す）
function setGroup_(inst, grp) {
  inst = clean_(inst, 60);
  grp = clean_(grp, 20);
  if (!inst) throw new Error('楽器が正しくありません');
  const g = settingsSheet_();
  const last = g.getLastRow();
  let row = -1;
  if (last >= 2) {
    const names = g.getRange(2, 1, last - 1, 1).getValues();
    for (let i = 0; i < names.length; i++) {
      if (String(names[i][0]).trim() === inst) { row = i + 2; break; }
    }
  }
  if (!grp) {
    if (row > 0) g.deleteRow(row);
    return;
  }
  if (row < 0) row = Math.max(last, 1) + 1;
  g.getRange(row, 1, 1, 3).setValues([[inst, grp, new Date()]]);
}

// ===== 表示順（全員で共有。「表示順」シートのA列＝グループ、B列＝楽器） =====
function ensureOrder_(ss) {
  let o = ss.getSheetByName(ORDER);
  if (!o) o = ss.insertSheet(ORDER);
  o.getRange(1, 1, 1, 2).setValues([['グループの並び', '楽器の並び']])
    .setFontWeight('bold').setBackground('#f0ede4');
  o.getRange('A:B').setNumberFormat('@');
  o.setFrozenRows(1);
  o.setColumnWidth(1, 160);
  o.setColumnWidth(2, 300);
}

function orderSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(ORDER) || (ensureOrder_(ss), ss.getSheetByName(ORDER));
}

function readOrder_() {
  const o = orderSheet_();
  const n = o.getLastRow() - 1;
  const out = { groups: [], insts: [] };
  if (n < 1) return out;
  o.getRange(2, 1, n, 2).getValues().forEach(r => {
    const g = String(r[0]).trim(), i = String(r[1]).trim();
    if (g) out.groups.push(g);
    if (i) out.insts.push(i);
  });
  return out;
}

// kind: 'groups'（グループの並び）または 'insts'（楽器の並び）。list は名前の配列
function setOrder_(kind, list) {
  if (kind !== 'groups' && kind !== 'insts') throw new Error('並び順の種類が正しくありません');
  if (!Array.isArray(list) || list.length > 300) throw new Error('並び順が正しくありません');
  const seen = {};
  const names = [];
  list.forEach(v => {
    const t = clean_(v, 60);
    if (t && !seen[t]) { seen[t] = true; names.push(t); }
  });
  const o = orderSheet_();
  const col = kind === 'groups' ? 1 : 2;
  const rows = Math.max(o.getMaxRows() - 1, names.length, 1);
  const vals = [];
  for (let i = 0; i < rows; i++) vals.push([i < names.length ? names[i] : '']);
  o.getRange(2, col, rows, 1).setValues(vals);
}

// ===== 名前の変更（全員で共有。「名称変更」シート：A:B＝楽器、D:E＝グループ） =====
function ensureAliases_(ss) {
  let a = ss.getSheetByName(ALIASES);
  if (!a) a = ss.insertSheet(ALIASES);
  a.getRange(1, 1, 1, 5).setValues([['楽器（元の名前）', '楽器（表示名）', '', 'グループ（元の名前）', 'グループ（表示名）']])
    .setFontWeight('bold').setBackground('#f0ede4');
  a.getRange('A:E').setNumberFormat('@');
  a.setFrozenRows(1);
  [260, 200, 30, 160, 160].forEach((w, i) => a.setColumnWidth(i + 1, w));
}

function aliasSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(ALIASES) || (ensureAliases_(ss), ss.getSheetByName(ALIASES));
}

function readAliasBlock_(a, col) {
  const n = a.getLastRow() - 1;
  const map = {};
  if (n < 1) return map;
  a.getRange(2, col, n, 2).getValues().forEach(r => {
    const k = String(r[0]).trim(), v = String(r[1]).trim();
    if (k && v) map[k] = v;
  });
  return map;
}

function readAliases_() {
  const a = aliasSheet_();
  return { insts: readAliasBlock_(a, 1), groups: readAliasBlock_(a, 4) };
}

// kind: 'inst' または 'group'。name が空なら元の名前に戻す
function setAlias_(kind, key, name) {
  if (kind !== 'inst' && kind !== 'group') throw new Error('変更の種類が正しくありません');
  key = clean_(key, 60);
  name = clean_(name, kind === 'group' ? 20 : 40);
  if (!key) throw new Error('対象が正しくありません');
  const a = aliasSheet_();
  const col = kind === 'inst' ? 1 : 4;
  const map = readAliasBlock_(a, col);
  if (name) map[key] = name; else delete map[key];
  const keys = Object.keys(map);
  const rows = Math.max(a.getMaxRows() - 1, keys.length, 1);
  const vals = [];
  for (let i = 0; i < rows; i++) vals.push(i < keys.length ? [keys[i], map[keys[i]]] : ['', '']);
  a.getRange(2, col, rows, 2).setValues(vals);
}

// 行データから「曲ID → その曲の担当楽器」を作る（主な楽器と同じものは省く）
function songInstOf_(r, inst) {
  const out = {};
  SONGS.forEach((s, i) => {
    const v = String(r[2 + i]).trim();
    if (v && v !== inst) out[s.id] = v;
  });
  return out;
}
