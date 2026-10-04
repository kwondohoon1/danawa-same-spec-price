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
const ODD_DROP = 0.8;                    // 이전 수집일 최저가의 80% 미만이면 '가격 이상?' (판매처 한 곳의 비정상 가격 등)

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
// 그래픽카드 이름의 OC 여부와 메모리 표기 (D7 / D6·D6X / 표기 없음)
const ocOf = (n) => /\bOC\b/i.test(n);
const memOf = (n) => (/\bG?D(?:DR)?6X?\b/i.test(n) ? 'D6' : /\bG?D(?:DR)?7\b/i.test(n) ? 'D7' : '');

// 상품 속성 (다나와 스펙 → 공통 형태)
function attrsFromSpec(cat, sp, name) {
  if (cat === 'gpu') {
    const chip = gpuChip(sp.chipset), mem = gb(sp.memory_size);
    return { base: chip && mem ? `${chip} · ${mem}GB` : '', fans: fansOf(sp.fans) || fansFromName(name), led: /LED 라이트/.test(sp.full_spec || '') || /\bA?RGB\b/i.test(name), oc: ocOf(name), mem: memOf(name) };
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
  if (cat === 'gpu') { const c = gpuChip(n), m = gb(n); return { base: c && m ? `${c} · ${m}GB` : '', fans: fansFromName(n), led: /\bA?RGB\b/.test(n), oc: ocOf(n), mem: memOf(n) }; }
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
    if (a.oc) { parts.push('OC'); tests.push((b) => b.oc); }                       // OC 모델은 OC끼리만
    if (a.mem === 'D7') { parts.push('D6 제외'); tests.push((b) => b.mem !== 'D6'); } // D7 은 D6 표기만 제외, 표기 없음은 포함
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

// ---------- 다나와 필터 URL ----------
// 다나와 목록 '상세검색' 필터 코드 (2026-09-30 기준). 같은 항목 안의 여러 값은 OR, 항목끼리는 AND
const codeMap = (s) => Object.fromEntries(s.split(',').map((x) => { const i = x.lastIndexOf('='); return [x.slice(0, i).toUpperCase().replace(/\s+/g, ''), x.slice(i + 1)]; }));
const DNW = {
  chip: codeMap('RTX 5090=1018234,RTX 5080=1018237,RTX 5070 Ti=1018240,RTX 5070=1018246,RTX 5060 Ti=1035862,RTX 5060=1018243,RTX 5050=1052509,RTX 4080 SUPER=925852,RTX 4070 Ti=823393,RTX 4070 SUPER=925846,RTX 4070=846919,RTX 4060 Ti=863683,RTX 4060=863686,RTX 3090=693490,RTX 3080 Ti=731872,RTX 3080=693451,RTX 3070 Ti=733036,RTX 3070=705349,RTX 3060 Ti=709720,RTX 3060=723391,RTX 3050=761620,GTX 1660 Ti=332302,GTX 1660 SUPER=622280,GTX 1660=338797,GTX 1650 SUPER=624611,GTX 1650=343624,RX 9070 XT=1022905,RX 9070 GRE=1147345,RX 9070=1022908,RX 9060 XT=1050613,RX 9060=1056586,RX 7900 XTX=818815,RX 7800 XT=901393,RX 7700 XT=901396,RX 7600 XT=944461,RX 7600=864124,RX 6900 XT=708859,RX 6800 XT=708862,RX 6800=708865,RX 6700 XT=726172,RX 6600 XT=741613,RX 6600=746104,RX 6500 XT=760483,RX 580=217480'),
  vram: codeMap('32=306823,24=306820,20=765568,16=188705,12=213322,10=693454,8=188704,6=137546,4=110066'),
  fans: [[1, '100040'], [2, '100041'], [3, '100042'], [4, '351085']],
  gddr7: '1018456',
  ssdIf: { 'PCIe5.0': '859759', 'PCIe4.0': '402191', SATA: '88980' },
  ssdCap: [[64, '610790'], [128, '610793'], [256, '610811'], [525, '610814'], [1024, '610817'], [2048, '610820'], [4096, '610823'], [8192, '682156'], [19456, '671123'], [1e9, '713191']],
  ssdRead: [[449, '90174'], [549, '93370'], [1499, '93371'], [2499, '221062'], [3999, '700813'], [5999, '700816'], [7999, '700819'], [11999, '701431'], [1e9, '927916']],
  ssdWrite: [[499, '90175'], [999, '93368'], [1499, '221067'], [1999, '221066'], [2999, '700822'], [3999, '700825'], [4999, '700828'], [5999, '700831'], [6999, '700834'], [8999, '702262'], [1e9, '927925']],
  ramGen: { DDR5: '748099', DDR4: '164333', DDR3: '1217' },
  ramCap: codeMap('128=230128,96=836026,64=157451,48=836023,32=109194,24=109193,16=90210,12=84071,8=1248,4=1246'),
  ramMods: { 1: '1228', 2: '1229', 4: '1231' },
  ramSpd: codeMap('8400=978784,8000=816907,7600=807862,7400=927928,7200=807856,7000=814987,6800=807853,6600=776749,6400=756892,6200=755377,6000=755374,5600=749644,5200=748117,4800=748702,4800D4=336316,4000=203452,3600=183447,3200=168792,3000=159636,2933=184392,2666=131762,2400=43870,2133=31641'),
  ramCl: [[14, '167912'], [15, '164647'], [16, '164473'], [17, '183449'], [18, '195363'], [19, '204922'], [22, '629105'], [26, '1019383'], [28, '804178'], [30, '774730'], [32, '755380'], [34, '749641'], [36, '749920'], [38, '748120'], [40, '748699'], [42, '762271'], [46, '790423'], [48, '844780'], [52, '983833']],
};
// 우리 제품 기준 스펙(경쟁 조건)으로 다나와 목록을 거른 URL. OC 여부는 다나와 필터가 없어 제외
function danawaFilterUrl(cat, a, sameListing) {
  if (sameListing) return danawaUrl(cat, sameListing.code);
  const v = [];
  if (a && a.base && cat === 'gpu') {
    const [chip, mem] = a.base.split(' · ');
    const c = DNW.chip[chip.toUpperCase().replace(/\s+/g, '')]; if (c) v.push(c);
    const m = DNW.vram[parseInt(mem, 10)]; if (m) v.push(m);
    if (a.fans) DNW.fans.filter(([n]) => n >= a.fans).forEach(([, x]) => v.push(x));
    if (a.mem === 'D7') v.push(DNW.gddr7);
  } else if (a && a.base && cat === 'ssd') {
    const cap = gb(a.base.split(' · ').pop());
    v.push(...(/^M\.2/.test(a.base) ? ['202347'] : []));
    const bus = Object.keys(DNW.ssdIf).find((k) => a.base.includes(k)); if (bus) v.push(DNW.ssdIf[bus]);
    if (cap) v.push(DNW.ssdCap.find(([hi]) => cap <= hi)[1]);
    if (a.nand === 'TLC' || a.nand === 'MLC') v.push('213319', '86089');
    if (a.dram) v.push('342157');
    const r = Math.floor(a.read / 1000) * 1000, w = Math.floor(a.write / 1000) * 1000;
    if (r) DNW.ssdRead.filter(([hi]) => hi >= r).forEach(([, x]) => v.push(x));
    if (w) DNW.ssdWrite.filter(([hi]) => hi >= w).forEach(([, x]) => v.push(x));
    if (a.warranty >= 5) v.push('720838'); else if (a.warranty >= 3) v.push('720838', '720841');
  } else if (a && a.base && cat === 'ram') {
    const m = a.base.match(/^(DDR\d)-(\d+) · (\d+)GB \((\d)개\)/);
    if (m) {
      v.push('1223', DNW.ramGen[m[1]] || '');
      const spd = DNW.ramSpd[m[1] === 'DDR4' && m[2] === '4800' ? '4800D4' : m[2]]; if (spd) v.push(spd);
      if (DNW.ramCap[m[3]]) v.push(DNW.ramCap[m[3]]);
      if (DNW.ramMods[m[4]]) v.push(DNW.ramMods[m[4]]);
    }
    if (a.cl) DNW.ramCl.filter(([n]) => n <= a.cl).forEach(([, x]) => v.push(x));
    if (a.rgb) v.push('247310');
  }
  const list = `https://prod.danawa.com/list/?cate=${CATS[cat].cate}`;
  const vv = v.filter(Boolean);
  return vv.length ? `${list}&searchOption=/searchAttributeValue=${vv.join(',')}` : list;
}

// ---------- 데이터 로드 ----------
// 가격은 07~18시 매시 갱신되므로 브라우저·CDN 캐시를 건너뛰고 항상 새로 받는다.
const REFRESH_MS = 5 * 60 * 1000;
const get = (u) => fetch(`${u}?t=${Date.now()}`, { cache: 'no-store' })
  .then((r) => { if (!r.ok) throw new Error(`${u.split('/').pop()} ${r.status}`); return r.text(); });

// 오늘 시간대별 수집 기록에서 가장 최근 수집 시각 (없으면 null)
// 오늘 기록이 아직 없으면(자정~첫 수집 전) 전날 마지막 수집 시각
async function latestCollectedAt() {
  for (const back of [0, 1]) {
    const day = new Date(Date.now() + 9 * 3600 * 1000 - back * 86400 * 1000).toISOString().slice(0, 10);
    try {
      const m = parseCSV(await get(`${SRC}/hourly/${day}/collected.csv`));
      const i = m.head.indexOf('collected_at');
      const at = m.rows.map((r) => r[i]).filter(Boolean).sort().pop();
      if (at) return at;
    } catch { /* 그날 기록 없음 */ }
  }
  return null;
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
    let price = null, date = null, prev = null, prevDate = null, k = -1;
    for (let i = 2; i < r.length; i++) { const v = parseInt(r[i], 10); if (v > 0) { price = v; date = dates[i - 2]; k = i; break; } }
    if (!price) continue;
    for (let i = k + 1; i < r.length; i++) { const v = parseInt(r[i], 10); if (v > 0) { prev = v; prevDate = dates[i - 2]; break; } }
    const sp = specBy[code] || {};
    const a = attrsFromSpec(cat, sp, name);
    items.push({ code, name, price, date, prev, prevDate, odd: !!(prev && price < prev * ODD_DROP),
      week: parseInt(r[r.length - 1], 10) || null, sp, a, key: a.base, own: OWN_SELLER.test(name) });
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
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '', blankrows: true });
    for (let h = 0; h < Math.min(aoa.length, 60); h++) {
      const head = aoa[h].map(squash);
      const iModel = head.indexOf('모델명');
      const iPrice = head.findIndex((v) => /^노출가\(2차/.test(v));
      if (iModel < 0 || iPrice < 0) continue;
      const col = (t) => head.indexOf(t);
      const iCat = col('카테고리'), iBrand = col('브랜드'), iCost = col('원가'), iSettle = col('정산가'), iMax = col('최대혜택가');
      const out = [];
      for (let r = h + 1; r < aoa.length; r++) {
        const model = String(aoa[r][iModel] || '').trim(), price = Number(aoa[r][iPrice]);
        if (!model) break; // 첫 표만 읽는다 (정산가세팅 아래쪽 '자사몰 기준' 표 제외)
        if (!(price > 0)) continue;
        const num = (i) => (i >= 0 ? Number(aoa[r][i]) || null : null);
        out.push({ cat: String(aoa[r][iCat] || '').trim().toUpperCase(), brand: String(aoa[r][iBrand] || '').trim(), model, price,
          cost: num(iCost), settle: num(iSettle), maxb: num(iMax) });
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
const gMaxBenefit = (O, g) => O - Math.min(fl10(O * g.card), O >= g.hi ? g.capH : g.capL); // 2차혜택가 → 최대혜택가
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
  if (cat === 'gpu') { a.led = a.led || nameA.led; a.fans = a.fans || nameA.fans; a.oc = a.oc || nameA.oc; a.mem = a.mem || nameA.mem; }
  if (cat === 'ram') { a.rgb = a.rgb || nameA.rgb; a.cl = a.cl || nameA.cl; }
  const group = (d && a.base && d.groups[a.base]) || [];
  // 여러 판매처가 파는 브랜드(삼성 등): 경쟁모델 없이 같은 모델의 다나와 최저가만 비교
  const sameOther = !!(listing && !listing.own && SHARED_BRAND.test(row.brand + ' ' + row.model));
  const rule = sameOther ? { label: '동일모델 다나와 최저가 (경쟁모델 제외)', test: (b) => b === listing.a }
    : a.base ? makeRule(cat, a) : null;
  // 그 외: 엑셀에 있는 자사 모델의 다나와 상품은 경쟁에서 제외
  let others = sameOther ? [listing]
    : group.filter((it) => !it.own && rule.test(it.a) && !ownCodes.has(it.code));
  // 이상 가격: 이력이 있으면 이전 수집일보다 20%↓, 이력이 없으면(새로 잡힌 상품) 같은 조건 경쟁 상품 중간값보다 20%↓
  const prices = others.map((it) => it.price).sort((x, y) => x - y);
  const med = prices.length >= 5 ? prices[Math.floor(prices.length / 2)] : null;
  const oddCodes = new Set(group.filter((it) => it.odd || (!it.prev && med && it.price < med * ODD_DROP)).map((it) => it.code));
  if (listing && (listing.odd || (!listing.prev && med && listing.price < med * ODD_DROP))) oddCodes.add(listing.code);
  const skip = (it) => state.exclude.has(it.code) || (state.dropOdd && oddCodes.has(it.code));
  const skippedOdd = sameOther ? 0 : others.filter(skip).length;
  if (!sameOther) others = others.filter((it) => !skip(it));
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
      if (sug === row.price) { verdict = '유지'; vcls = 'ok'; }
      else { verdict = sug < row.price ? '인하' : '인상'; vcls = sug < row.price ? 'up' : 'down'; } // 색은 차이 칸과 같게 (인하=우리가 비쌈)
    }
    if (row.cost) {
      be = breakeven2nd(row.cost, g);
      if (settleAt2nd(sug, g) < row.cost) { sug = row.price; verdict = '손실·보류'; vcls = 'warn'; }
    }
  }
  const sugSettle = settleAt2nd(sug, g);
  const curSettle = row.settle || settleAt2nd(row.price, g);
  const filterUrl = danawaFilterUrl(cat, a, sameOther ? listing : null);
  return { ...row, catKey: cat, a, rule, listing, group, others, low, diff, pct: low ? diff / low.price : null,
    rank: low ? cheaper + 1 : null, total: others.length + 1, sug, sugDiff: sug - row.price, verdict, vcls, be,
    sugSettle, sugMargin: row.cost ? (sugSettle - row.cost) / row.cost : null, sameOther, filterUrl, skippedOdd, oddCodes, med,
    curMax: row.maxb || gMaxBenefit(row.price, g), curSettle, sugMax: gMaxBenefit(sug, g) };
}

// ---------- 직접 유지 ----------
// 사용자가 판정 버튼을 눌러 유지로 바꾼 모델 (모델명 기준, 새 가격으로 다시 계산해도 유지)
state.keep = new Set();
state.exclude = new Set();   // '빼기'로 직접 뺀 경쟁 상품 코드
function applyKeep(a) {
  if (!state.keep.has(a.model) || (a.verdict !== '인하' && a.verdict !== '인상')) return a;
  const g = state.g, sugSettle = settleAt2nd(a.price, g);
  return { ...a, verdict: '유지', vcls: 'ok', manual: true, autoVerdict: a.verdict, autoSug: a.sug, sug: a.price, sugDiff: 0,
    sugSettle, sugMax: gMaxBenefit(a.price, g), sugMargin: a.cost ? (sugSettle - a.cost) / a.cost : null };
}

const oddTag = (it, a) => {
  if (!it || !a || !a.oddCodes || !a.oddCodes.has(it.code)) return '';
  const why = it.odd ? `이전 수집일(${esc(it.prevDate)}) ${won(it.prev)}원 → ${won(it.price)}원 (${Math.round((1 - it.price / it.prev) * 100)}% 하락)`
    : `이전 가격 없음 · 같은 조건 경쟁 상품 중간값 ${won(a.med)}원보다 ${Math.round((1 - it.price / a.med) * 100)}% 낮음`;
  return ` <span class="tag odd" title="${why}">가격 이상?</span>`;
};

// ---------- 렌더: 비교 ----------
const link = (cat, it, text) => `<a href="${danawaUrl(cat, it.code)}" target="_blank" rel="noopener" title="다나와에서 보기">${esc(text ?? it.name)}</a>`;
function renderCompare() {
  if (!Object.keys(state.db).length) return;
  $('#compare').hidden = false;
  const f = $('#cat-filter').value, onlyUp = $('#only-expensive').checked;
  const matched = state.rows.map(matchRow);
  const ownCodes = new Set(matched.filter((m) => m.listing).map((m) => m.listing.code));
  state.results = matched.map((m) => applyKeep(analyze(m, ownCodes)));
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
    const dnw = a.listing ? `${link(a.catKey, a.listing, won(a.listing.price))}${oddTag(a.listing, a)}${shared ? '<div class="spec">공용 상품페이지</div>' : ''}` : '<span class="muted">다나와 미등록</span>';
    const xbtn = a.low && !shared ? ` <button type="button" class="xbtn" data-ex="${esc(a.low.code)}" title="이 가격을 빼고 다음으로 싼 상품으로 비교">빼기</button>` : '';
    const low = a.low ? `<b>${won(a.low.price)}</b>${oddTag(a.low, a)}${xbtn}<div class="spec">${link(a.catKey, a.low)}${shared ? ' <span class="tag">동일모델</span>' : ''}</div>${a.skippedOdd ? `<div class="spec">더 싼 가격 ${a.skippedOdd}개 뺌</div>` : ''}` : '<span class="muted">동일스펙 없음</span>';
    const rule = (a.rule ? esc(a.rule.label) : '<span class="err">스펙 판별 불가</span>')
      + ` <a class="flink" href="${esc(a.filterUrl)}" target="_blank" rel="noopener" title="우리 제품 기준 스펙으로 다나와 목록 필터">다나와 필터 ↗</a>`;
    const sugNote = a.verdict === '손실·보류' ? ` <span class="spec">손익분기 ${won(a.be)}</span>` : a.sugDiff ? ` <span class="spec">${signed(a.sugDiff)}</span>` : '';
    const canKeep = a.manual || a.verdict === '인하' || a.verdict === '인상';
    const badge = !canKeep ? `<span class="badge ${a.vcls}">${a.verdict}</span>`
      : a.manual ? `<button type="button" class="badge vbtn manual ${a.vcls}" data-keep="${i}" title="직접 유지로 바꿈 · 다시 누르면 원래 판정(${a.autoVerdict} ${won(a.autoSug)})">유지</button>`
      : `<button type="button" class="badge vbtn ${a.vcls}" data-keep="${i}" title="누르면 유지 (현재 노출가 그대로)">${a.verdict}</button>`;
    const open = a.group.length ? `<button type="button" class="ghost open" data-open="${i}">${state.openRow === a.model ? '접기' : '보기'}</button>` : '';
    return `<tr data-r="${i}">
      ${td(0, a.model, `<div class="mtop"><span class="cat">${esc(CATS[a.catKey].label)}</span><span>${model}</span>${open}</div><div class="rule">${rule}</div>`, 'left')}
      ${td(1, a.price, `<b>${won(a.price)}</b>`)}
      ${td(2, a.listing ? a.listing.price : '', dnw)}
      ${td(3, a.low ? a.low.price : '', low)}
      ${td(4, a.diff ?? '', `<span class="${cls}">${a.diff == null ? '-' : signed(a.diff)}</span><div class="spec">${pctTxt(a.pct)}${a.rank ? ` · ${a.rank}/${a.total}위` : ''}</div>`)}
      ${td(5, a.sug, `<b>${won(a.sug)}</b>${sugNote}<div>${badge}</div>`, 'sugc')}
      ${td(6, Math.round(a.sugSettle), `${won(a.sugSettle)}<div class="spec">${a.sugMargin == null ? '' : '마진 ' + pctTxt(a.sugMargin)}</div>`)}
    </tr>${state.openRow === a.model ? detailRow(a) : ''}`;
  }).join('') || '<tr><td colspan="7" class="center muted">표시할 모델이 없습니다</td></tr>';
  tb.querySelectorAll('[data-ex]').forEach((b) => b.addEventListener('click', () => { state.exclude.add(b.dataset.ex); renderCompare(); }));
  tb.querySelectorAll('[data-keep]').forEach((b) => b.addEventListener('click', () => {
    const m = list[+b.dataset.keep].model;
    if (state.keep.has(m)) state.keep.delete(m); else state.keep.add(m);
    renderCompare();
  }));
  tb.querySelectorAll('[data-open]').forEach((b) => b.addEventListener('click', () => {
    const a = list[+b.dataset.open]; state.openRow = state.openRow === a.model ? null : a.model; renderCompare();
  }));

  const all = state.results, cnt = (v) => all.filter((a) => a.verdict === v).length;
  const oddRows = all.filter((a) => (a.low && a.oddCodes.has(a.low.code)) || (state.dropOdd && a.skippedOdd)).length;
  const ob = $('#drop-odd');
  ob.textContent = state.dropOdd ? `이상 가격 다시 포함 (${oddRows})` : `이상 가격 빼기 (${oddRows})`;
  ob.classList.toggle('on', !!state.dropOdd);
  ob.disabled = !state.dropOdd && !oddRows;
  const ub = $('#undo-ex');
  ub.hidden = !state.exclude.size; ub.textContent = `직접 뺀 가격 ${state.exclude.size}개 되돌리기`;
  $('#summary').innerHTML = `
    <div class="tile"><b>${all.length}</b><span>엑셀 모델</span></div>
    <div class="tile"><b class="up">${cnt('인하')}</b><span>인하 제안</span></div>
    <div class="tile"><b class="down">${cnt('인상')}</b><span>인상 제안</span></div>
    <div class="tile"><b class="ok">${cnt('유지')}</b><span>유지${all.some((a) => a.manual) ? ` (직접 ${all.filter((a) => a.manual).length})` : ' (−0.5~−1%)'}</span></div>
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
      <td class="name">${link(a.catKey, it)}${mine ? '<span class="tag own">자사</span>' : ok ? '' : '<span class="tag">조건 밖</span>'}${oddTag(it, a)}${state.exclude.has(it.code) ? '<span class="tag">뺌</span>' : ''}</td>
      <td class="left spec">${esc(specSummary(a.catKey, it.a))}</td>
      <td>${won(it.price)}</td><td class="${d > 0 ? 'up' : d < 0 ? 'down' : ''}">${signed(d)}</td><td class="muted">${esc(it.date)}</td></tr>`;
  }).join('');
  return `<tr class="detail"><td colspan="7"><div class="spec" style="margin-bottom:6px">경쟁 조건: <b>${esc(a.rule.label)}</b> · 회색 = 조건 밖(기본 스펙만 같음)</div>
    <table class="grid"><thead><tr><th>#</th><th class="left">${esc(a.a.base)} 전체 상품</th><th class="left">스펙</th><th>다나와 최저가</th><th>엑셀 노출가 − 이 상품</th><th>가격일</th></tr></thead><tbody>${rows}</tbody></table></td></tr>`;
}
function specSummary(cat, a) {
  if (cat === 'gpu') return [a.fans ? a.fans + '팬' : '', a.led ? 'LED' : '', a.oc ? 'OC' : '', a.mem].filter(Boolean).join(' · ');
  if (cat === 'ssd') return [a.nand, a.dram ? 'DRAM' : 'DRAM없음', a.read ? `${a.read.toLocaleString()}/${a.write.toLocaleString()}` : '', a.warranty ? a.warranty + '년' : ''].filter(Boolean).join(' · ');
  return [a.cl ? 'CL' + a.cl : '', a.rgb ? 'RGB' : ''].filter(Boolean).join(' · ');
}
function danawaUrl(cat, code) { return `https://prod.danawa.com/info/?pcode=${encodeURIComponent(code)}&cate=${CATS[cat].cate}`; }

// ---------- 엑셀로 받기 ----------
// 단가표 '최저가비교' 시트 양식: 1행 날짜, 2행 구분, 3행 제목, 4행부터 정산가세팅 순서 (ExcelJS 로 색·서식 포함)
const XL = {
  cols: [
    ['다나와코드', 11], ['상품명', 50],
    ['노출가', 11], ['최대혜택가', 11], ['원가', 11], ['마진', 10], ['마진율', 8],
    ['노출가', 11], ['최대혜택가', 11], ['원가', 11], ['마진', 10], ['마진율', 8],
    ['다나와코드', 11], ['상품명', 46], ['노출가', 11], ['가격차이', 10], ['상품 URL', 22], ['필터 URL', 22],
    ['판정', 9],
  ],
  // [시작열, 끝열, 제목, 구분색, 제목행색]
  groups: [[1, 2, '', 'D9D9D9', 'F2F2F2'], [3, 7, '현재세팅', '9BC2E6', 'DDEBF7'], [8, 12, '제안세팅', 'A9D08E', 'E2EFDA'],
    [13, 18, '가격비교', 'F4B084', 'FCE4D6'], [19, 19, '', 'D9D9D9', 'F2F2F2']],
  money: [3, 4, 5, 6, 8, 9, 10, 11, 15], pct: [7, 12], links: [1, 13, 17, 18],
};
const fillOf = (rgb) => ({ type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + rgb } });
const thin = (rgb) => { const s = { style: 'thin', color: { argb: 'FF' + rgb } }; return { top: s, left: s, bottom: s, right: s }; };

function gFormulas(o, cost, g) {
  const rd = (x, m) => `ROUNDDOWN(${x}*${m},-1)`;
  const sec = (x) => `(${x}-${rd(x, g.sel)}-${rd(x, g.dup)})`;              // 등재가 → 2차혜택가
  const a = `ROUNDDOWN(${o}/(1-${g.sel}-${g.dup}),-1)`;                       // 2차혜택가 → 등재가 (정산가세팅 L열과 같은 방식)
  const L = `(IF(${sec(`(${a}+10)`)}<=${o},${a}+10,IF(${sec(a)}<=${o},${a},${a}-10)))`;
  const selPart = g.sel ? `-${rd(L, g.sel)}/${g.sel}*${g.selS}` : '';
  const card = `MIN(ROUNDDOWN(${sec(L)}*${g.card},-1),IF(${sec(L)}>=${g.hi},${g.capH},${g.capL}))*${g.cardS}`;
  const settle = `${L}-${L}*${g.fee}-${L}*${g.pro}${selPart}-${rd(L, g.dup)}-${card}`;
  return {
    max: `${o}-MIN(ROUNDDOWN(${o}*${g.card},-1),IF(${o}>=${g.hi},${g.capH},${g.capL}))`,
    margin: `IF(${cost}="","",${settle}-${cost})`,
    rate: `IF(OR(${cost}="",${cost}=0),"",K{r}/${cost})`,
  };
}

function buildExportWorkbook() {
  const R = Math.round;
  const now = new Date(Date.now() + 9 * 3600 * 1000).toISOString();
  const when = state.collectedAt ? `${state.collectedAt.slice(0, 10)} ${state.collectedAt.slice(11, 16)}` : '';
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('최저가비교', { views: [{ state: 'frozen', xSplit: 2, ySplit: 3 }] });
  const font = { name: '맑은 고딕', size: 10 };
  ws.columns = XL.cols.map(([, width]) => ({ width }));

  ws.mergeCells('A1:B1');                                   // 수집 날짜·시각 (맨 위)
  ws.getCell('A1').value = when ? `수집 ${when}  (다나와 최저가 · 배송비 미포함)` : '수집 시각 확인 불가';
  ws.getCell('A1').font = { ...font, size: 11, bold: true };
  ws.getCell('C1').value = `${state.dropOdd || state.exclude.size ? '이상 가격 뺀 기준 · ' : ''}노란색 = 인하 필요 · 빨간색 = 인상 · 받은 시각 ${now.slice(0, 10)} ${now.slice(11, 16)}`;
  ws.getCell('C1').font = { ...font, color: { argb: 'FF808080' } };

  for (const [c1, c2, title, g, h] of XL.groups) {
    if (c2 > c1) ws.mergeCells(2, c1, 2, c2);
    const top = ws.getCell(2, c1);
    top.value = title; top.fill = fillOf(g); top.font = { ...font, bold: true }; top.alignment = { horizontal: 'center', vertical: 'middle' };
    for (let c = c1; c <= c2; c++) {
      ws.getCell(2, c).border = thin('A6A6A6');
      const hc = ws.getCell(3, c);
      hc.value = XL.cols[c - 1][0]; hc.fill = fillOf(h); hc.font = { ...font, bold: true };
      hc.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }; hc.border = thin('A6A6A6');
    }
  }
  ws.getRow(2).height = 20; ws.getRow(3).height = 20;

  state.results.forEach((a, i) => {
    const r = 4 + i, low = a.low, cost = a.cost || null;
    const curM = cost ? a.curSettle - cost : null, sugM = cost ? a.sugSettle - cost : null;
    const lowUrl = low ? danawaUrl(a.catKey, low.code) : '';
    const vals = [
      a.listing ? a.listing.code : '', a.model,
      a.price, R(a.curMax), cost ?? '', curM == null ? '' : R(curM), curM == null ? '' : curM / cost,
      a.sug, R(a.sugMax), cost ?? '', sugM == null ? '' : R(sugM), sugM == null ? '' : sugM / cost,
      low ? low.code : '', low ? low.name : '', low ? low.price : '', low ? a.price - low.price : '', lowUrl, a.filterUrl,
      a.verdict,
    ];
    const urls = { 1: a.listing ? danawaUrl(a.catKey, a.listing.code) : '', 13: lowUrl, 17: lowUrl, 18: a.filterUrl };
    const mark = a.verdict === '인하' ? 'FFFF00' : a.verdict === '인상' ? 'FF8080' : null;   // 인하 노랑 · 인상 빨강
    vals.forEach((v, j) => {
      const c = j + 1, cell = ws.getCell(r, c);
      cell.value = urls[c] && v !== '' ? { text: String(v), hyperlink: urls[c] } : v;
      cell.font = { ...font };
      cell.border = thin('D9D9D9');
      cell.alignment = [1, 13, 19].includes(c) ? { vertical: 'middle', horizontal: 'center' } : { vertical: 'middle' };
      if (XL.money.includes(c)) cell.numFmt = '#,##0;[Red]-#,##0';
      if (XL.pct.includes(c)) cell.numFmt = '0.00%;[Red]-0.00%';
      if (XL.links.includes(c) && urls[c] && v !== '') cell.font = { ...font, color: { argb: 'FF0563C1' }, underline: true };
      if (c === 16 && typeof v === 'number') {
        cell.numFmt = '+#,##0;-#,##0;0';
        if (v) cell.font = { ...font, bold: true, color: { argb: v > 0 ? 'FFC00000' : 'FF0070C0' } }; // 우리가 비싸면 빨강, 싸면 파랑
      }
      if (mark && c >= 8 && c <= 12) cell.fill = fillOf(mark);                                      // 제안세팅 칸 색
    });
    ws.getCell(r, 8).font = { ...ws.getCell(r, 8).font, bold: true };
    // 제안세팅 I~L 은 수식 (result 는 열자마자 보이도록 미리 계산한 값)
    const f = gFormulas(`H${r}`, `J${r}`, state.g);
    ws.getCell(r, 9).value = { formula: f.max, result: R(a.sugMax) };
    ws.getCell(r, 10).value = { formula: `IF(E${r}="","",E${r})`, result: cost ?? '' };
    ws.getCell(r, 11).value = { formula: f.margin, result: sugM == null ? '' : sugM };
    ws.getCell(r, 12).value = { formula: f.rate.replace('{r}', r), result: sugM == null ? '' : sugM / cost };
  });
  wb.calcProperties = { fullCalcOnLoad: true };
  return { wb, name: `최저가비교_${now.slice(0, 10).replace(/-/g, '')}_${now.slice(11, 16).replace(':', '')}.xlsx` };
}

// 열어 둔 창이 예전 코드면(사이트 업데이트 후 새로고침 안 함) 옛 양식으로 받게 되므로 먼저 확인
async function isStale() {
  try {
    const mine = ([...document.scripts].map((x) => x.src).find((u) => /app\.js\?v=/.test(u)) || '').split('v=')[1];
    const html = await fetch(`index.html?t=${Date.now()}`, { cache: 'no-store' }).then((r) => r.text());
    const live = (html.match(/app\.js\?v=([\w]+)/) || [])[1];
    return !!(mine && live && mine !== live);
  } catch { return false; }
}

async function exportXlsx() {
  if (!state.results.length) return;
  if (await isStale()) { alert('사이트가 업데이트되었습니다. 새로고침(F5) 후 엑셀을 다시 올리고 받아 주세요.'); return; }
  if (!window.ExcelJS) { alert('엑셀 모듈을 불러오지 못했습니다. 새로고침 후 다시 눌러 주세요.'); return; }
  const { wb, name } = buildExportWorkbook();
  const buf = await wb.xlsx.writeBuffer();
  const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  const link = document.createElement('a');
  link.href = url; link.download = name; document.body.appendChild(link); link.click(); link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

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
$('#drop-odd').addEventListener('click', () => { state.dropOdd = !state.dropOdd; renderCompare(); });
$('#undo-ex').addEventListener('click', () => { state.exclude.clear(); renderCompare(); });
$('#export').addEventListener('click', exportXlsx);
$('#only-expensive').addEventListener('change', renderCompare);
setupSelection();
$('#clear').addEventListener('click', () => {
  state.rows = []; state.results = []; state.openRow = null; state.keep.clear(); state.exclude.clear(); input.value = '';
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
