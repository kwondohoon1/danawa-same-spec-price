// 동일스펙 가격비교 — 가격/스펙은 크롤러 저장소에서 실시간으로 읽고,
// 엑셀은 브라우저 안에서만 파싱한다 (서버 전송·저장 없음).
'use strict';

const SRC = 'https://raw.githubusercontent.com/kwondohoon1/danawa-monitor-crawler/main/data';
const CATS = {
  gpu: { label: '그래픽카드', excel: 'VGA', cate: '112753' },
  ssd: { label: 'SSD', excel: 'SSD', cate: '112760' },
  ram: { label: 'RAM', excel: 'RAM', cate: '112752' },
};
const OWN_SELLER = /한성컴퓨터/;          // 다나와 상품명에 붙는 자사 유통사 표기
const EXCLUDE = /중고|해외구매|리퍼|벌크/;
const SHARED_BRAND = /삼성/;               // 여러 판매처가 파는 브랜드 → 경쟁모델 대신 같은 모델의 다나와 최저가만 사용
const TARGET = 0.99;                     // 제안가 = 경쟁 최저가 × 99% (1% 낮게)
const KEEP_HI = 0.995;                   // 경쟁가 대비 -0.5% ~ -1% 이면 유지

// G마켓 조건 기본값 (엑셀 정산가세팅 '마켓 조건'을 찾지 못했을 때)
const G_DEFAULT = { md: 0.02, sel: 0.11, selS: 0.02, dup: 0.08, card: 0.07, cardS: 0.5, capH: 150000, hi: 1200000, capL: 70000, fee: 0.09, pro: 0.02 };
const G_LABELS = { 'MD쿠폰': 'md', '선택쿠폰': 'sel', '선택쿠폰셀러부담': 'selS', '중복쿠폰(셀러부담)': 'dup', '카드즉시할인': 'card', '카드할인셀러부담비율': 'cardS', '카드할인한도(고가)': 'capH', '카드할인고가기준금액': 'hi', '카드할인한도(중저가)': 'capL', 'Cat수수료': 'fee', '프로모션수수료': 'pro' };

const state = { db: {}, rows: [], g: { ...G_DEFAULT }, gFromExcel: false, openRow: null, results: [] };
const $ = (s) => document.querySelector(s);
const won = (n) => (n == null || isNaN(n) ? '-' : Math.round(n).toLocaleString('ko-KR'));
const signed = (n) => (n > 0 ? '+' : n < 0 ? '−' : '±') + Math.abs(Math.round(n)).toLocaleString('ko-KR');
const pctTxt = (p) => (p == null ? '-' : (p > 0 ? '+' : '') + (p * 100).toFixed(2) + '%');
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const fl10 = (n) => Math.floor(n / 10 + 1e-9) * 10;

// ---------- CSV ----------
function parseCSV(text) {
  text = text.replace(/^﻿/, '');
  const rows = []; let row = [], f = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (q) {
      if (c === '"') { if (text[i + 1] === '"') { f += '"'; i++; } else q = false; }
      else f += c;
    } else if (c === '"') q = true;
    else if (c === ',') { row.push(f); f = ''; }
    else if (c === '\n') { row.push(f.replace(/\r$/, '')); rows.push(row); row = []; f = ''; }
    else f += c;
  }
  if (f || row.length) { row.push(f); rows.push(row); }
  const head = rows.shift() || [];
  return { head, rows: rows.filter((r) => r.length > 1) };
}

// ---------- 스펙 파싱 ----------
const norm = (s) => String(s || '').toUpperCase().replace(/\s+/g, ' ').trim();
function gpuChip(s) {
  const m = norm(s).match(/\b(RTX|GTX|RX|ARC)\s*([A-Z]?\d{3,4})\s*(TI SUPER|TI|SUPER|XTX|XT|GRE)?\b/);
  return m ? [m[1], m[2], m[3]].filter(Boolean).join(' ') : '';
}
function gb(s) { const m = String(s || '').match(/(\d+)\s*(TB|GB)/i); return m ? (m[2].toUpperCase() === 'TB' ? +m[1] * 1024 : +m[1]) : null; }
const sizeLabel = (g) => (g >= 1024 ? g / 1024 + 'TB' : g + 'GB');
const mbps = (s) => { const m = String(s || '').replace(/,/g, '').match(/(\d+)\s*MB/i); return m ? +m[1] : 0; };
const years = (s) => { const m = String(s || '').match(/(\d+)\s*년/); return m ? +m[1] : 0; };
const fansOf = (s) => { const m = String(s || '').match(/(\d)\s*팬/); return m ? +m[1] : 0; };
const fansFromName = (n) => (/TRIPLE|3\s*FAN/i.test(n) ? 3 : /DUAL|2\s*FAN/i.test(n) ? 2 : /SINGLE|1\s*FAN/i.test(n) ? 1 : 0);
const clOf = (s) => { const m = String(s || '').match(/CL\s*(\d+)/i); return m ? +m[1] : 0; };

// 상품 속성 (다나와 스펙 → 공통 형태)
function attrsFromSpec(cat, sp, name) {
  if (cat === 'gpu') {
    const chip = gpuChip(sp.chipset), mem = gb(sp.memory_size);
    return { base: chip && mem ? `${chip} · ${mem}GB` : '', fans: fansOf(sp.fans) || fansFromName(name), led: /LED 라이트/.test(sp.full_spec || '') || /\bA?RGB\b/i.test(name) };
  }
  if (cat === 'ssd') {
    const cap = gb(sp.capacity);
    const gen = (String(sp.interface).match(/PCIe\s*(\d)\.0/i) || [])[1];
    const ff = /M\.2/i.test(sp.form_factor) ? 'M.2' : (sp.form_factor || '').split(' ')[0];
    const bus = gen ? `PCIe${gen}.0` : (/SATA/i.test(sp.interface) ? 'SATA' : (sp.interface || '').split(' ')[0]);
    return {
      base: cap ? [ff, bus, sizeLabel(cap)].filter(Boolean).join(' · ') : '',
      nand: /MLC/.test(sp.nand) ? 'MLC' : /TLC/.test(sp.nand) ? 'TLC' : /QLC/.test(sp.nand) ? 'QLC' : '',
      dram: /DRAM 탑재/.test(sp.dram || ''), read: mbps(sp.seq_read), write: mbps(sp.seq_write), warranty: years(sp.warranty),
    };
  }
  if (sp.usage && !/데스크탑/.test(sp.usage)) return { base: '' };
  const gen = norm(sp.generation), spd = (String(sp.speed).match(/(\d{4,5})/) || [])[1];
  const cap = gb(sp.capacity), mods = parseInt(sp.module_count, 10) || ((String(sp.capacity).match(/x\s*(\d)/i) || [])[1] | 0) || 1;
  return { base: gen && spd && cap ? `${gen}-${spd} · ${cap}GB (${mods}개)` : '', cl: clOf(sp.timing) || clOf(name), rgb: /RGB/i.test(sp.led_color || '') || /RGB/i.test(name) };
}

// 엑셀 모델명 → 속성 (다나와 미등록 모델용)
function attrsFromName(cat, name) {
  const n = norm(name);
  if (cat === 'gpu') { const c = gpuChip(n), m = gb(n); return { base: c && m ? `${c} · ${m}GB` : '', fans: fansFromName(n), led: /\bA?RGB\b/.test(n) }; }
  if (cat === 'ram') {
    const gen = (n.match(/DDR(\d)/) || n.match(/\bD(\d)-/) || [])[1];
    const spd = (n.match(/(?:DDR\d|D\d)-(\d{4,5})/) || [])[1];
    const kit = n.match(/(\d+)\s*GB\s*\(\s*(\d+)\s*G?B?\s*X\s*(\d)\s*\)/);
    const cap = kit ? +kit[1] : gb(n), mods = kit ? +kit[3] : 1;
    return { base: gen && spd && cap ? `DDR${gen}-${spd} · ${cap}GB (${mods}개)` : '', cl: clOf(n), rgb: /RGB/.test(n) };
  }
  return { base: '' }; // SSD는 다나와 상품 매칭이 있어야 세대/성능을 알 수 있음
}

// 우리 제품 속성 → 경쟁 상품 조건 (같거나 더 좋은 스펙만)
function makeRule(cat, a) {
  const parts = [a.base], tests = [];
  if (cat === 'gpu') {
    if (a.fans) { parts.push(`${a.fans}팬 이상`); tests.push((b) => b.fans >= a.fans); }
    if (a.led) { parts.push('LED'); tests.push((b) => b.led); }
  } else if (cat === 'ssd') {
    if (a.nand === 'TLC' || a.nand === 'MLC') { parts.push('TLC 이상'); tests.push((b) => b.nand === 'TLC' || b.nand === 'MLC'); }
    if (a.dram) { parts.push('DRAM'); tests.push((b) => b.dram); }
    const r = Math.floor(a.read / 1000) * 1000, w = Math.floor(a.write / 1000) * 1000;
    if (r || w) { parts.push(`${r.toLocaleString()}/${w.toLocaleString()}MB/s 이상`); tests.push((b) => b.read >= r && b.write >= w); }
    if (a.warranty) { parts.push(`${a.warranty}년 이상`); tests.push((b) => b.warranty >= a.warranty); }
  } else {
    if (a.cl) { parts.push(`CL${a.cl} 이하`); tests.push((b) => b.cl && b.cl <= a.cl); }
    if (a.rgb) { parts.push('RGB'); tests.push((b) => b.rgb); }
  }
  return { label: parts.join(' · '), test: (b) => tests.every((t) => t(b)) };
}

// ---------- 데이터 로드 ----------
// 가격은 07~18시 매시 갱신되므로 브라우저·CDN 캐시를 건너뛰고 항상 새로 받는다.
const REFRESH_MS = 5 * 60 * 1000;
const get = (u) => fetch(`${u}?t=${Date.now()}`, { cache: 'no-store' })
  .then((r) => { if (!r.ok) throw new Error(`${u.split('/').pop()} ${r.status}`); return r.text(); });

// 오늘 시간대별 수집 기록에서 가장 최근 수집 시각 (없으면 null)
async function latestCollectedAt() {
  const today = new Date(Date.now() + 9 * 3600 * 1000).toISOString().slice(0, 10);
  try {
    const m = parseCSV(await get(`${SRC}/hourly/${today}/collected.csv`));
    const i = m.head.indexOf('collected_at');
    return m.rows.map((r) => r[i]).filter(Boolean).sort().pop() || null;
  } catch { return null; }
}

async function loadCat(cat) {
  const [p, s] = await Promise.all([get(`${SRC}/latest/${cat}.csv`), get(`${SRC}/specs/${cat}_specs.csv`)]);
  const prices = parseCSV(p), specs = parseCSV(s);
  const dates = prices.head.slice(2);
  const specBy = {};
  for (const r of specs.rows) { const o = {}; specs.head.forEach((h, i) => (o[h] = r[i])); specBy[o.product_code] = o; }
  const items = [];
  for (const r of prices.rows) {
    const code = r[0], name = r[1];
    if (!code || EXCLUDE.test(name)) continue;
    let price = null, date = null;
    for (let i = 2; i < r.length; i++) { const v = parseInt(r[i], 10); if (v > 0) { price = v; date = dates[i - 2]; break; } }
    if (!price) continue;
    const sp = specBy[code] || {};
    const a = attrsFromSpec(cat, sp, name);
    items.push({ code, name, price, date, week: parseInt(r[r.length - 1], 10) || null, sp, a, key: a.base, own: OWN_SELLER.test(name) });
  }
  const groups = {};
  for (const it of items) if (it.key) (groups[it.key] ||= []).push(it);
  for (const k in groups) groups[k].sort((x, y) => x.price - y.price);
  return { items, groups, today: dates[0] };
}

async function loadAll() {
  const st = $('#data-status');
  try {
    const [res, at] = await Promise.all([
      Promise.all(Object.keys(CATS).map(async (c) => [c, await loadCat(c)])),
      latestCollectedAt(),
    ]);
    res.forEach(([c, d]) => (state.db[c] = d));
    state.collectedAt = at;
    const n = res.reduce((a, [, d]) => a + d.items.length, 0);
    const when = at ? `${at.slice(0, 10)} ${at.slice(11, 16)}` : res[0][1].today;
    st.innerHTML = `가격 데이터 <b>${when}</b> 수집 기준 · ${n.toLocaleString()}개 상품 불러옴`;
    $('#foot').innerHTML = `데이터: <a href="https://github.com/kwondohoon1/danawa-monitor-crawler" target="_blank" rel="noopener">danawa-monitor-crawler</a> (다나와 최저가, 07~18시 매시 갱신, 배송비 미포함). 중고·해외구매·리퍼·벌크 제외.`;
    renderBrowseSpecs();
    if (state.rows.length) renderCompare();
  } catch (e) {
    st.innerHTML = `<span class="err">데이터를 불러오지 못했습니다: ${esc(e.message)}</span>`;
  }
}

// ---------- 엑셀 ----------
function readExcel(file) {
  const fs = $('#file-status');
  const reader = new FileReader();
  reader.onload = (ev) => {
    try {
      const wb = XLSX.read(ev.target.result, { type: 'array' });
      const rows = extractRows(wb);
      if (!rows.length) throw new Error("'모델명'과 '노출가 (2차혜택가)' 열이 있는 시트를 찾지 못했습니다");
      state.rows = rows;
      fs.innerHTML = `<b>${esc(file.name)}</b> · ${rows.length}개 모델 읽음${state.gFromExcel ? ' · G마켓 조건 반영' : ' · <span class="err">G마켓 조건을 못 찾아 기본값 사용</span>'} <span class="muted">(브라우저 메모리에만 있음)</span>`;
      renderCompare();
    } catch (e) {
      fs.innerHTML = `<span class="err">엑셀을 읽지 못했습니다: ${esc(e.message)}</span>`;
    }
  };
  reader.readAsArrayBuffer(file);
}

const squash = (v) => String(v ?? '').replace(/\s+/g, '');
function extractRows(wb) {
  const order = [...wb.SheetNames].sort((a, b) => (b === '정산가세팅') - (a === '정산가세팅'));
  for (const name of order) {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' });
    for (let h = 0; h < Math.min(aoa.length, 60); h++) {
      const head = aoa[h].map(squash);
      const iModel = head.indexOf('모델명');
      const iPrice = head.findIndex((v) => /^노출가\(2차/.test(v));
      if (iModel < 0 || iPrice < 0) continue;
      const col = (t) => head.indexOf(t);
      const iCat = col('카테고리'), iBrand = col('브랜드'), iCost = col('원가'), iSettle = col('정산가');
      const out = [];
      for (let r = h + 1; r < aoa.length; r++) {
        const model = String(aoa[r][iModel] || '').trim(), price = Number(aoa[r][iPrice]);
        if (!model || !(price > 0)) continue;
        out.push({ cat: String(aoa[r][iCat] || '').trim().toUpperCase(), brand: String(aoa[r][iBrand] || '').trim(), model, price,
          cost: iCost >= 0 ? Number(aoa[r][iCost]) || null : null, settle: iSettle >= 0 ? Number(aoa[r][iSettle]) || null : null });
      }
      if (out.length) { readGConditions(aoa, h); return out; }
    }
  }
  return [];
}
// '마켓 조건' 표에서 G마켓·옥션 열 값을 읽는다
function readGConditions(aoa, limit) {
  state.g = { ...G_DEFAULT }; state.gFromExcel = false;
  let gCol = -1, found = 0;
  for (let r = 0; r < limit; r++) {
    const row = aoa[r] || [];
    if (gCol < 0) { gCol = row.findIndex((v) => /^G마켓/.test(squash(v))); if (gCol >= 0) continue; }
    const key = G_LABELS[squash(row[0])];
    if (key && gCol >= 0 && row[gCol] !== '' && !isNaN(Number(row[gCol]))) { state.g[key] = Number(row[gCol]); found++; }
  }
  state.gFromExcel = found >= 8;
}

// ---------- G마켓 계산 (엑셀 정산가세팅과 같은 식) ----------
const gSecond = (L, g) => L - fl10(L * g.sel) - fl10(L * g.dup);
function gSettle(L, g) {
  const sa = fl10(L * g.sel), da = fl10(L * g.dup), R = L - sa - da;
  const card = Math.min(fl10(R * g.card), R >= g.hi ? g.capH : g.capL);
  return L - L * g.fee - L * g.pro - (g.sel ? (sa / g.sel) * g.selS : 0) - da - card * g.cardS;
}
function gListFrom2nd(O, g) {
  const L1 = fl10(O / (1 - g.sel - g.dup));
  for (const L of [L1 + 10, L1, L1 - 10]) if (gSecond(L, g) <= O) return L;
  return L1 - 10;
}
const settleAt2nd = (O, g) => gSettle(gListFrom2nd(O, g), g);
function breakeven2nd(cost, g) { // 정산가 ≥ 원가가 되는 최소 2차혜택가
  let lo = 0, hi = Math.ceil(cost * 3 / 10) * 10;
  while (hi - lo > 10) { const mid = fl10((lo + hi) / 2); if (gSettle(mid, g) >= cost) hi = mid; else lo = mid; }
  return gSecond(hi, g);
}

// ---------- 자사 모델 ↔ 다나와 상품 매칭 ----------
const ALIAS = [[/에이서|ACER/g, 'ACER'], [/프레데터|PREDATOR/g, 'PREDATOR'], [/지포스|GEFORCE/g, 'GEFORCE'], [/라데온|RADEON/g, 'RADEON'], [/삼성전자|삼성|SAMSUNG/g, 'SAMSUNG']];
function tokens(s) {
  let n = norm(s).replace(/[()\[\],/]/g, ' ').replace(/\bD(\d)-/g, 'DDR$1-');
  for (const [re, to] of ALIAS) n = n.replace(re, to);
  return n.split(' ').filter((t) => t && !/^(M\.2|NVME|한성컴퓨터|STCOM|제이씨현)$/.test(t));
}
function catOf(row) {
  const c = row.cat;
  if (/VGA|GPU|그래픽/.test(c)) return 'gpu';
  if (/SSD/.test(c)) return 'ssd';
  if (/RAM|메모리/.test(c)) return 'ram';
  if (gpuChip(row.model)) return 'gpu';
  if (/DDR\d|\bD\d-/i.test(row.model)) return 'ram';
  return 'ssd';
}
function findListing(cat, row, base) {
  const d = state.db[cat]; if (!d) return null;
  const want = tokens(row.model), cap = gb(row.model);
  let best = null, bestScore = 0;
  for (const it of base ? d.groups[base] || [] : d.items) {
    const have = new Set(tokens(it.name));
    let hit = 0; for (const t of want) if (have.has(t)) hit++;
    if (cat !== 'gpu' && cap && gb(it.name) !== cap) continue;
    const score = hit / want.length + (it.own ? 0.05 : 0);
    if (score > bestScore) { bestScore = score; best = it; }
  }
  return bestScore >= 0.85 ? best : null;
}

function matchRow(row) {
  const cat = catOf(row), nameA = attrsFromName(cat, row.model);
  return { row, cat, nameA, listing: findListing(cat, row, nameA.base) };
}

function analyze({ row, cat, nameA, listing }, ownCodes) {
  const d = state.db[cat], g = state.g;
  // 속성: 다나와 상품이 있으면 그 스펙, 없으면 모델명. 이름의 RGB/팬 정보는 항상 반영
  const a = listing ? { ...listing.a } : nameA;
  if (cat === 'gpu') { a.led = a.led || nameA.led; a.fans = a.fans || nameA.fans; }
  if (cat === 'ram') { a.rgb = a.rgb || nameA.rgb; a.cl = a.cl || nameA.cl; }
  const group = (d && a.base && d.groups[a.base]) || [];
  // 여러 판매처가 파는 브랜드(삼성 등): 경쟁모델 없이 같은 모델의 다나와 최저가만 비교
  const sameOther = !!(listing && !listing.own && SHARED_BRAND.test(row.brand + ' ' + row.model));
  const rule = sameOther ? { label: '동일모델 다나와 최저가 (경쟁모델 제외)', test: (b) => b === listing.a }
    : a.base ? makeRule(cat, a) : null;
  // 그 외: 엑셀에 있는 자사 모델의 다나와 상품은 경쟁에서 제외
  const others = sameOther ? [listing]
    : group.filter((it) => !it.own && rule.test(it.a) && !ownCodes.has(it.code));
  const low = others[0] || null;
  const diff = low ? row.price - low.price : null;
  const cheaper = others.filter((it) => it.price < row.price).length;

  // 제안 노출가
  let sug = row.price, verdict = '비교 불가', vcls = 'muted', be = null;
  if (low) {
    const ratio = row.price / low.price;
    if (ratio >= TARGET && ratio <= KEEP_HI) { verdict = '유지'; vcls = 'ok'; }
    else {
      sug = fl10(low.price * TARGET);
      verdict = sug < row.price ? '인하' : '인상'; vcls = sug < row.price ? 'down' : 'up';
    }
    if (row.cost) {
      be = breakeven2nd(row.cost, g);
      if (settleAt2nd(sug, g) < row.cost) { sug = row.price; verdict = '손실·보류'; vcls = 'warn'; }
    }
  }
  const sugSettle = settleAt2nd(sug, g);
  return { ...row, catKey: cat, a, rule, listing, group, others, low, diff, pct: low ? diff / low.price : null,
    rank: low ? cheaper + 1 : null, total: others.length + 1, sug, sugDiff: sug - row.price, verdict, vcls, be,
    sugSettle, sugMargin: row.cost ? (sugSettle - row.cost) / row.cost : null, sameOther };
}

// ---------- 렌더: 비교 ----------
const link = (cat, it, text) => `<a href="${danawaUrl(cat, it.code)}" target="_blank" rel="noopener" title="다나와에서 보기">${esc(text ?? it.name)}</a>`;
function renderCompare() {
  if (!Object.keys(state.db).length) return;
  $('#compare').hidden = false;
  const f = $('#cat-filter').value, onlyUp = $('#only-expensive').checked;
  const matched = state.rows.map(matchRow);
  const ownCodes = new Set(matched.filter((m) => m.listing).map((m) => m.listing.code));
  state.results = matched.map((m) => analyze(m, ownCodes));
  state.ownCodes = ownCodes;
  const list = state.results.filter((a) => (!f || CATS[a.catKey].excel === f) && (!onlyUp || a.diff > 0));
  state.list = list; state.sel = null;
  const tb = $('#compare-table tbody');
  // data-v = 복사될 원본 값 (숫자는 쉼표 없이)
  const td = (c, v, html, cls = '') => `<td data-c="${c}" data-v="${esc(v ?? '')}" class="${cls}">${html}</td>`;
  tb.innerHTML = list.map((a, i) => {
    const cls = a.diff > 0 ? 'up' : a.diff < 0 ? 'down' : '';
    const model = a.listing ? link(a.catKey, a.listing, a.model) : esc(a.model);
    const shared = a.sameOther;
    const dnw = a.listing ? `${link(a.catKey, a.listing, won(a.listing.price))}${shared ? '<div class="spec">공용 상품페이지</div>' : ''}` : '<span class="muted">다나와 미등록</span>';
    const low = a.low ? `<b>${won(a.low.price)}</b><div class="spec">${link(a.catKey, a.low)}${shared ? ' <span class="tag">동일모델</span>' : ''}</div>` : '<span class="muted">동일스펙 없음</span>';
    const rule = a.rule ? esc(a.rule.label) : '<span class="err">스펙 판별 불가</span>';
    const sugNote = a.verdict === '손실·보류' ? ` <span class="spec">손익분기 ${won(a.be)}</span>` : a.sugDiff ? ` <span class="spec">${signed(a.sugDiff)}</span>` : '';
    const open = a.group.length ? `<button type="button" class="ghost open" data-open="${i}">${state.openRow === a.model ? '접기' : '보기'}</button>` : '';
    return `<tr data-r="${i}">
      ${td(0, a.model, `<div class="mtop"><span class="cat">${esc(CATS[a.catKey].label)}</span><span>${model}</span>${open}</div><div class="rule">${rule}</div>`, 'left')}
      ${td(1, a.price, `<b>${won(a.price)}</b>`)}
      ${td(2, a.listing ? a.listing.price : '', dnw)}
      ${td(3, a.low ? a.low.price : '', low)}
      ${td(4, a.diff ?? '', `<span class="${cls}">${a.diff == null ? '-' : signed(a.diff)}</span><div class="spec">${pctTxt(a.pct)}${a.rank ? ` · ${a.rank}/${a.total}위` : ''}</div>`)}
      ${td(5, a.sug, `<b>${won(a.sug)}</b>${sugNote}<div><span class="badge ${a.vcls}">${a.verdict}</span></div>`, 'sugc')}
      ${td(6, Math.round(a.sugSettle), `${won(a.sugSettle)}<div class="spec">${a.sugMargin == null ? '' : '마진 ' + pctTxt(a.sugMargin)}</div>`)}
    </tr>${state.openRow === a.model ? detailRow(a) : ''}`;
  }).join('') || '<tr><td colspan="7" class="center muted">표시할 모델이 없습니다</td></tr>';
  tb.querySelectorAll('[data-open]').forEach((b) => b.addEventListener('click', () => {
    const a = list[+b.dataset.open]; state.openRow = state.openRow === a.model ? null : a.model; renderCompare();
  }));

  const all = state.results, cnt = (v) => all.filter((a) => a.verdict === v).length;
  $('#summary').innerHTML = `
    <div class="tile"><b>${all.length}</b><span>엑셀 모델</span></div>
    <div class="tile"><b class="down">${cnt('인하')}</b><span>인하 제안</span></div>
    <div class="tile"><b class="up">${cnt('인상')}</b><span>인상 제안</span></div>
    <div class="tile"><b class="ok">${cnt('유지')}</b><span>유지 (−0.5~−1%)</span></div>
    <div class="tile"><b class="warn">${cnt('손실·보류')}</b><span>손실·보류</span></div>
    <div class="tile"><b>${cnt('비교 불가')}</b><span>비교 대상 없음</span></div>`;
}

function detailRow(a) {
  const rows = a.group.map((it, i) => {
    const same = a.sameOther && it.code === a.listing.code;
    const mine = it.own || (state.ownCodes.has(it.code) && !same);
    const ok = !mine && a.rule.test(it.a);
    const d = a.price - it.price;
    return `<tr class="${mine ? 'own' : ok ? '' : 'dim'}"><td class="center">${i + 1}</td>
      <td class="name">${link(a.catKey, it)}${mine ? '<span class="tag own">자사</span>' : ok ? '' : '<span class="tag">조건 밖</span>'}</td>
      <td class="left spec">${esc(specSummary(a.catKey, it.a))}</td>
      <td>${won(it.price)}</td><td class="${d > 0 ? 'up' : d < 0 ? 'down' : ''}">${signed(d)}</td><td class="muted">${esc(it.date)}</td></tr>`;
  }).join('');
  return `<tr class="detail"><td colspan="7"><div class="spec" style="margin-bottom:6px">경쟁 조건: <b>${esc(a.rule.label)}</b> · 회색 = 조건 밖(기본 스펙만 같음)</div>
    <table class="grid"><thead><tr><th>#</th><th class="left">${esc(a.a.base)} 전체 상품</th><th class="left">스펙</th><th>다나와 최저가</th><th>엑셀 노출가 − 이 상품</th><th>가격일</th></tr></thead><tbody>${rows}</tbody></table></td></tr>`;
}
function specSummary(cat, a) {
  if (cat === 'gpu') return [a.fans ? a.fans + '팬' : '', a.led ? 'LED' : ''].filter(Boolean).join(' · ');
  if (cat === 'ssd') return [a.nand, a.dram ? 'DRAM' : 'DRAM없음', a.read ? `${a.read.toLocaleString()}/${a.write.toLocaleString()}` : '', a.warranty ? a.warranty + '년' : ''].filter(Boolean).join(' · ');
  return [a.cl ? 'CL' + a.cl : '', a.rgb ? 'RGB' : ''].filter(Boolean).join(' · ');
}
function danawaUrl(cat, code) { return `https://prod.danawa.com/info/?pcode=${encodeURIComponent(code)}&cate=${CATS[cat].cate}`; }

// ---------- 엑셀형 셀 선택 · 복사 ----------
const table = () => $('#compare-table');
function paintSel() {
  const s = state.sel;
  table().querySelectorAll('td.sel').forEach((el) => el.classList.remove('sel'));
  if (!s) return;
  const [r0, r1] = [Math.min(s.r0, s.r1), Math.max(s.r0, s.r1)], [c0, c1] = [Math.min(s.c0, s.c1), Math.max(s.c0, s.c1)];
  table().querySelectorAll('tbody tr[data-r]').forEach((tr) => {
    const r = +tr.dataset.r; if (r < r0 || r > r1) return;
    tr.querySelectorAll('td[data-c]').forEach((el) => { const c = +el.dataset.c; if (c >= c0 && c <= c1) el.classList.add('sel'); });
  });
}
function selTSV() {
  const s = state.sel; if (!s) return '';
  const [r0, r1] = [Math.min(s.r0, s.r1), Math.max(s.r0, s.r1)], [c0, c1] = [Math.min(s.c0, s.c1), Math.max(s.c0, s.c1)];
  const lines = [];
  table().querySelectorAll('tbody tr[data-r]').forEach((tr) => {
    const r = +tr.dataset.r; if (r < r0 || r > r1) return;
    const cells = [...tr.querySelectorAll('td[data-c]')].filter((el) => +el.dataset.c >= c0 && +el.dataset.c <= c1);
    lines.push(cells.map((el) => el.dataset.v).join('\t'));
  });
  return lines.join('\n');
}
function toast(msg) { const t = $('#copy-toast'); t.textContent = msg; t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(() => (t.hidden = true), 1800); }
function setupSelection() {
  const tb = table().tBodies[0];
  let drag = false;
  const cellOf = (el) => { const td = el.closest('td[data-c]'), tr = el.closest('tr[data-r]'); return td && tr ? { r: +tr.dataset.r, c: +td.dataset.c } : null; };
  tb.addEventListener('mousedown', (e) => {
    if (e.button !== 0 || e.target.closest('button')) return;
    const p = cellOf(e.target); if (!p) return;
    e.preventDefault();
    if (e.shiftKey && state.sel) { state.sel.r1 = p.r; state.sel.c1 = p.c; }
    else state.sel = { r0: p.r, c0: p.c, r1: p.r, c1: p.c };
    drag = true; paintSel();
  });
  tb.addEventListener('mouseover', (e) => {
    if (!drag) return; const p = cellOf(e.target); if (!p) return;
    state.sel.r1 = p.r; state.sel.c1 = p.c; paintSel();
  });
  document.addEventListener('mouseup', () => (drag = false));
  // 열 제목 클릭 = 열 전체 선택 (Shift = 여러 열)
  table().tHead.addEventListener('click', (e) => {
    const th = e.target.closest('th[data-c]'); if (!th || !state.list?.length) return;
    const c = +th.dataset.c, last = state.list.length - 1;
    state.sel = e.shiftKey && state.sel ? { ...state.sel, r0: 0, r1: last, c1: c } : { r0: 0, r1: last, c0: c, c1: c };
    paintSel();
  });
  document.addEventListener('mousedown', (e) => { if (!e.target.closest('#compare-table')) { state.sel = null; paintSel(); } });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') { state.sel = null; paintSel(); } });
  document.addEventListener('copy', (e) => {
    if (!state.sel || window.getSelection().toString()) return;
    const text = selTSV(); if (!text) return;
    e.clipboardData.setData('text/plain', text); e.preventDefault();
    toast(`복사됨 ✓ ${text.split('\n').length}행`);
  });
}

// ---------- 렌더: 동일스펙 가격표 ----------
function renderBrowseSpecs() {
  const cat = $('#b-cat').value, d = state.db[cat]; if (!d) return;
  const keys = Object.keys(d.groups).filter((k) => d.groups[k].length > 1)
    .sort((a, b) => d.groups[b].length - d.groups[a].length || a.localeCompare(b, 'ko'));
  const sel = $('#b-spec'), prev = sel.value;
  sel.innerHTML = keys.map((k) => `<option value="${esc(k)}">${esc(k)} (${d.groups[k].length})</option>`).join('');
  if (keys.includes(prev)) sel.value = prev;
  renderBrowse();
}
function renderBrowse() {
  const cat = $('#b-cat').value, d = state.db[cat]; if (!d) return;
  const g = d.groups[$('#b-spec').value] || [];
  const q = norm($('#b-search').value);
  const min = g.length ? g[0].price : 0;
  $('#browse-table tbody').innerHTML = g.filter((it) => !q || norm(it.name).includes(q)).map((it, i) => `
    <tr class="${it.own ? 'own' : ''}"><td class="center">${i + 1}</td>
      <td class="name">${link(cat, it)}${it.own ? '<span class="tag own">자사</span>' : ''}</td>
      <td class="left spec">${esc(specSummary(cat, it.a))}</td>
      <td>${won(it.price)}</td><td class="${it.price > min ? 'up' : ''}">${it.price > min ? signed(it.price - min) : '최저'}</td>
      <td class="center muted">${esc(it.sp.registration_month || '')}</td></tr>`).join('')
    || '<tr><td colspan="6" class="center muted">상품이 없습니다</td></tr>';
}

// ---------- 이벤트 ----------
const drop = $('#drop'), input = $('#file');
input.addEventListener('change', () => input.files[0] && readExcel(input.files[0]));
['dragenter', 'dragover'].forEach((e) => drop.addEventListener(e, (ev) => { ev.preventDefault(); drop.classList.add('over'); }));
['dragleave', 'drop'].forEach((e) => drop.addEventListener(e, (ev) => { ev.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', (ev) => { const f = ev.dataTransfer.files[0]; if (f) readExcel(f); });
$('#cat-filter').addEventListener('change', renderCompare);
$('#only-expensive').addEventListener('change', renderCompare);
setupSelection();
$('#clear').addEventListener('click', () => {
  state.rows = []; state.results = []; state.openRow = null; input.value = '';
  $('#compare').hidden = true; $('#file-status').textContent = '엑셀은 이 브라우저 안에서만 읽고, 어디에도 전송·저장하지 않습니다.';
});
$('#b-cat').addEventListener('change', renderBrowseSpecs);
$('#b-spec').addEventListener('change', renderBrowse);
$('#b-search').addEventListener('input', renderBrowse);

loadAll();
// 페이지를 열어둔 동안에도 새 시간대 수집이 올라오면 다시 불러온다 (올린 엑셀 비교도 새 가격으로 다시 계산)
setInterval(async () => {
  const at = await latestCollectedAt();
  if (at && at !== state.collectedAt) loadAll();
}, REFRESH_MS);
