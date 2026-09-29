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

const state = { db: {}, rows: [], openRow: null };
const $ = (s) => document.querySelector(s);
const won = (n) => (n == null || isNaN(n) ? '-' : Math.round(n).toLocaleString('ko-KR'));
const signed = (n) => (n > 0 ? '+' : n < 0 ? '−' : '±') + Math.abs(Math.round(n)).toLocaleString('ko-KR');
const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

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

// ---------- 스펙 키 ----------
const norm = (s) => String(s || '').toUpperCase().replace(/\s+/g, ' ').trim();

function gpuChip(s) {
  const m = norm(s).match(/\b(RTX|GTX|RX|ARC)\s*([A-Z]?\d{3,4})\s*(TI SUPER|TI|SUPER|XTX|XT|GRE)?\b/);
  return m ? [m[1], m[2], m[3]].filter(Boolean).join(' ') : '';
}
function gb(s) { const m = String(s || '').match(/(\d+)\s*(TB|GB)/i); return m ? (m[2].toUpperCase() === 'TB' ? +m[1] * 1024 : +m[1]) : null; }
function sizeLabel(g) { return g >= 1024 ? g / 1024 + 'TB' : g + 'GB'; }

const keyers = {
  gpu: {
    fromSpec: (sp) => { const c = gpuChip(sp.chipset); const m = gb(sp.memory_size); return c && m ? `${c} · ${m}GB` : ''; },
    fromName: (name) => { const c = gpuChip(name); const m = gb(name); return c && m ? `${c} · ${m}GB` : ''; },
  },
  ssd: {
    fromSpec: (sp) => {
      const cap = gb(sp.capacity); if (!cap) return '';
      const gen = (String(sp.interface).match(/PCIe\s*(\d)\.0/i) || [])[1];
      const ff = /M\.2/i.test(sp.form_factor) ? 'M.2' : (sp.form_factor || '').split(' ')[0];
      const bus = gen ? `PCIe${gen}.0` : (/SATA/i.test(sp.interface) ? 'SATA' : (sp.interface || '').split(' ')[0]);
      return [ff, bus, sizeLabel(cap)].filter(Boolean).join(' · ');
    },
    fromName: null, // SSD는 이름만으로 세대를 알 수 없어 다나와 상품 매칭이 필요
  },
  ram: {
    fromSpec: (sp) => {
      if (sp.usage && !/데스크탑/.test(sp.usage)) return '';
      const gen = norm(sp.generation); const spd = (String(sp.speed).match(/(\d{4,5})/) || [])[1];
      const cap = gb(sp.capacity); const mods = parseInt(sp.module_count, 10) || ((String(sp.capacity).match(/x\s*(\d)/i) || [])[1] | 0) || 1;
      return gen && spd && cap ? `${gen}-${spd} · ${cap}GB (${mods}개)` : '';
    },
    fromName: (name) => {
      const n = norm(name);
      const gen = (n.match(/DDR(\d)/) || n.match(/\bD(\d)-/) || [])[1];
      const spd = (n.match(/(?:DDR\d|D\d)-(\d{4,5})/) || [])[1];
      const kit = n.match(/(\d+)\s*GB\s*\(\s*(\d+)\s*G?B?\s*X\s*(\d)\s*\)/);
      const cap = kit ? +kit[1] : gb(n); const mods = kit ? +kit[3] : 1;
      return gen && spd && cap ? `DDR${gen}-${spd} · ${cap}GB (${mods}개)` : '';
    },
  },
};

// ---------- 데이터 로드 ----------
async function loadCat(cat) {
  const [p, s] = await Promise.all([
    fetch(`${SRC}/latest/${cat}.csv`).then((r) => { if (!r.ok) throw new Error(`${cat}.csv ${r.status}`); return r.text(); }),
    fetch(`${SRC}/specs/${cat}_specs.csv`).then((r) => { if (!r.ok) throw new Error(`${cat}_specs.csv ${r.status}`); return r.text(); }),
  ]);
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
    const last = parseInt(r[r.length - 1], 10) || null;
    const sp = specBy[code] || {};
    const key = keyers[cat].fromSpec(sp);
    items.push({ code, name, price, date, week: last, key, sp, own: OWN_SELLER.test(name) });
  }
  const groups = {};
  for (const it of items) if (it.key) (groups[it.key] ||= []).push(it);
  for (const k in groups) groups[k].sort((a, b) => a.price - b.price);
  return { items, groups, today: dates[0] };
}

async function loadAll() {
  const st = $('#data-status');
  try {
    const res = await Promise.all(Object.keys(CATS).map(async (c) => [c, await loadCat(c)]));
    res.forEach(([c, d]) => (state.db[c] = d));
    const today = res[0][1].today;
    const n = res.reduce((a, [, d]) => a + d.items.length, 0);
    st.innerHTML = `가격 데이터 <b>${today}</b> 기준 · ${n.toLocaleString()}개 상품 불러옴`;
    $('#foot').innerHTML = `데이터: <a href="https://github.com/kwondohoon1/danawa-monitor-crawler" target="_blank" rel="noopener">danawa-monitor-crawler</a> (다나와 최저가, 매일 갱신). 중고·해외구매·리퍼·벌크 상품 제외.`;
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
      fs.innerHTML = `<b>${esc(file.name)}</b> · ${rows.length}개 모델 읽음 <span class="muted">(브라우저 메모리에만 있음)</span>`;
      renderCompare();
    } catch (e) {
      fs.innerHTML = `<span class="err">엑셀을 읽지 못했습니다: ${esc(e.message)}</span>`;
    }
  };
  reader.readAsArrayBuffer(file);
}

// 헤더 행에서 '모델명' + '노출가 (2차…)' 열을 찾는다 (정산가세팅 시트 우선)
function extractRows(wb) {
  const order = [...wb.SheetNames].sort((a, b) => (b === '정산가세팅') - (a === '정산가세팅'));
  for (const name of order) {
    const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: '' });
    for (let h = 0; h < Math.min(aoa.length, 60); h++) {
      const head = aoa[h].map((v) => String(v).replace(/\s+/g, ''));
      const iModel = head.indexOf('모델명');
      const iPrice = head.findIndex((v) => /^노출가\(2차/.test(v));
      if (iModel < 0 || iPrice < 0) continue;
      const iCat = head.indexOf('카테고리'), iBrand = head.indexOf('브랜드');
      const out = [];
      for (let r = h + 1; r < aoa.length; r++) {
        const model = String(aoa[r][iModel] || '').trim();
        const price = Number(aoa[r][iPrice]);
        if (!model || !(price > 0)) continue;
        out.push({ cat: String(aoa[r][iCat] || '').trim().toUpperCase(), brand: String(aoa[r][iBrand] || '').trim(), model, price, sheet: name });
      }
      if (out.length) return out;
    }
  }
  return [];
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
// key가 있으면 같은 동일스펙 그룹 안에서만 찾는다 (다른 스펙 상품과 잘못 매칭 방지)
function findListing(cat, row, key) {
  const d = state.db[cat]; if (!d) return null;
  const want = tokens(row.model);
  const cap = gb(row.model);
  let best = null, bestScore = 0;
  for (const it of key ? d.groups[key] || [] : d.items) {
    const have = new Set(tokens(it.name));
    let hit = 0; for (const t of want) if (have.has(t)) hit++;
    if (cat !== 'gpu' && cap && gb(it.name) !== cap) continue;
    const score = hit / want.length + (it.own ? 0.05 : 0);
    if (score > bestScore) { bestScore = score; best = it; }
  }
  return bestScore >= 0.85 ? best : null;
}

function analyze(row) {
  const cat = catOf(row);
  const d = state.db[cat];
  let key = keyers[cat].fromName ? keyers[cat].fromName(row.model) : '';
  const listing = findListing(cat, row, key);
  if (!key) key = (listing && listing.key) || '';
  const group = (d && key && d.groups[key]) || [];
  const others = group.filter((it) => !it.own && (!listing || it.code !== listing.code));
  const low = others[0] || null;
  const diff = low ? row.price - low.price : null;
  const cheaper = others.filter((it) => it.price < row.price).length;
  return { ...row, catKey: cat, key, listing, group, others, low, diff, pct: low ? diff / low.price : null, rank: low ? cheaper + 1 : null, total: others.length + 1 };
}

// ---------- 렌더: 비교 ----------
function renderCompare() {
  if (!Object.keys(state.db).length) return;
  $('#compare').hidden = false;
  const f = $('#cat-filter').value, onlyUp = $('#only-expensive').checked;
  const all = state.rows.map(analyze);
  const list = all.filter((a) => (!f || CATS[a.catKey].excel === f) && (!onlyUp || a.diff > 0));
  const tb = $('#compare-table tbody');
  tb.innerHTML = list.map((a, i) => {
    const cls = a.diff > 0 ? 'up' : a.diff < 0 ? 'down' : '';
    const low = a.low ? `${won(a.low.price)}<div class="spec">${esc(a.low.name)}</div>` : '<span class="muted">동일스펙 없음</span>';
    const miss = !a.key ? '<span class="err">스펙 판별 불가</span>' : esc(a.key);
    return `<tr data-i="${i}">
      <td class="center">${esc(CATS[a.catKey].label)}</td>
      <td class="name">${esc(a.model)}${a.listing ? ` <a class="tag" href="${danawaUrl(a.catKey, a.listing.code)}" target="_blank" rel="noopener">다나와</a>` : ''}</td>
      <td class="left nowrap">${miss}</td>
      <td><b>${won(a.price)}</b></td>
      <td class="name" style="min-width:200px;text-align:right">${low}</td>
      <td class="${cls}">${a.diff == null ? '-' : signed(a.diff)}</td>
      <td class="${cls}">${a.pct == null ? '-' : (a.pct > 0 ? '+' : '') + (a.pct * 100).toFixed(1) + '%'}</td>
      <td class="center">${a.rank ? `${a.rank} / ${a.total}` : '-'}</td>
      <td class="center">${a.group.length ? `<button type="button" class="ghost" data-open="${i}">${state.openRow === a.model ? '접기' : '보기'}</button>` : ''}</td>
    </tr>${state.openRow === a.model ? detailRow(a) : ''}`;
  }).join('') || '<tr><td colspan="9" class="center muted">표시할 모델이 없습니다</td></tr>';
  tb.querySelectorAll('[data-open]').forEach((b) => b.addEventListener('click', () => {
    const a = list[+b.dataset.open]; state.openRow = state.openRow === a.model ? null : a.model; renderCompare();
  }));

  const withLow = all.filter((a) => a.low);
  const up = withLow.filter((a) => a.diff > 0).length, down = withLow.filter((a) => a.diff < 0).length, eq = withLow.length - up - down;
  $('#summary').innerHTML = `
    <div class="tile"><b>${all.length}</b><span>엑셀 모델</span></div>
    <div class="tile"><b class="up">${up}</b><span>최저가보다 비쌈</span></div>
    <div class="tile"><b class="down">${down}</b><span>최저가보다 쌈</span></div>
    <div class="tile"><b>${eq}</b><span>최저가와 같음</span></div>
    <div class="tile"><b>${all.length - withLow.length}</b><span>비교 대상 없음</span></div>`;
}

function detailRow(a) {
  const rows = a.group.map((it, i) => {
    const own = it.own || (a.listing && it.code === a.listing.code);
    const d = a.price - it.price;
    return `<tr class="${own ? 'own' : ''}"><td class="center">${i + 1}</td>
      <td class="name"><a href="${danawaUrl(a.catKey, it.code)}" target="_blank" rel="noopener">${esc(it.name)}</a>${own ? '<span class="tag own">자사</span>' : ''}</td>
      <td>${won(it.price)}</td><td class="${d > 0 ? 'up' : d < 0 ? 'down' : ''}">${signed(d)}</td><td class="muted">${esc(it.date)}</td></tr>`;
  }).join('');
  return `<tr class="detail"><td colspan="9"><table class="grid"><thead><tr><th>#</th><th class="left">동일스펙 상품 (${esc(a.key)})</th><th>다나와 최저가</th><th>우리 노출가 − 이 상품</th><th>가격일</th></tr></thead><tbody>${rows}</tbody></table></td></tr>`;
}

function danawaUrl(cat, code) { return `https://prod.danawa.com/info/?pcode=${encodeURIComponent(code)}&cate=${CATS[cat].cate}`; }

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
  const tb = $('#browse-table tbody');
  tb.innerHTML = g.filter((it) => !q || norm(it.name).includes(q)).map((it, i) => `
    <tr class="${it.own ? 'own' : ''}"><td class="center">${i + 1}</td>
      <td class="name"><a href="${danawaUrl(cat, it.code)}" target="_blank" rel="noopener">${esc(it.name)}</a>${it.own ? '<span class="tag own">자사</span>' : ''}</td>
      <td>${won(it.price)}</td><td class="${it.price > min ? 'up' : ''}">${it.price > min ? signed(it.price - min) : '최저'}</td>
      <td class="muted">${won(it.week)}</td><td class="center muted">${esc(it.sp.registration_month || '')}</td></tr>`).join('')
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
$('#clear').addEventListener('click', () => {
  state.rows = []; state.openRow = null; input.value = '';
  $('#compare').hidden = true; $('#file-status').textContent = '엑셀은 이 브라우저 안에서만 읽고, 어디에도 전송·저장하지 않습니다.';
});
$('#b-cat').addEventListener('change', renderBrowseSpecs);
$('#b-spec').addEventListener('change', renderBrowse);
$('#b-search').addEventListener('input', renderBrowse);

loadAll();
