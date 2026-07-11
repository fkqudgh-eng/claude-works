// ==UserScript==
// @name         CloudCast 메뉴보드 자동기입 (식단표 연동)
// @namespace    local.cloudcast.menuautofill
// @version      4.4
// @description  주간메뉴표 xlsx → 보드명 자동 파싱 → 미리보기(코너별 선택 체크박스) 후 메뉴보드 자동기입. 사이트 프로파일·양식검증·self-test 내장. Save 버튼은 절대 누르지 않음.
// @match        https://apps.cloud-cast.com/menuboard/menu/*
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  // ═══════════════════════════════════════════════════════════════
  // [6] 사이트 프로파일 — 양식·매핑·검증·테스트를 한 객체로 외부화.
  //     다른 사업장/양식은 프로파일만 추가하면 코어 로직 수정 불필요.
  // ═══════════════════════════════════════════════════════════════
  const PROFILES = [{
    id: 'nh-tower',
    label: 'NH농협타워점',
    match: /apps\.cloud-cast\.com\/menuboard/,
    COL: { 월: 'F', 화: 'G', 수: 'H', 목: 'I', 금: 'J' },
    RANGE: {
      조식: { 2: [5, 9],  3: [10, 14], 라면: [15, 16] },
      중식: { 1: [18, 22], 2: [23, 28], 3: [29, 33], 플러스바: [34, 35] },
      석식: { 2: [37, 42], 라면: [43, 44] },
    },
    // [4] 양식 검증 앵커: D열 코너 라벨(줄바꿈 포함될 수 있어 includes로 대조)
    ANCHORS: { D5: '코너2', D10: '코너3', D15: '셀프라면', D18: '코너1', D23: '코너2', D29: '코너3', D37: '코너2', D43: '셀프라면' },
    boardName: {
      day: /\(([월화수목금])\)/,
      meal: [[/조식/, '조식'], [/석식/, '석식'], [/중식|중/, '중식']],
      pos: /\((?:왼쪽|오른쪽)\)/,
      dateTail: /_?\s*\d+\/\d+.*$/,
    },
    ramenKey: '라면',
    // [7] self-test: mock 워크시트 + 기대값 — 실 xlsx 없이 파싱 로직 회귀검증
    test: {
      mock: {
        F18: '열무김치냉국수', F19: '추가밥', F20: '갈비만두', F21: '빵가루마요샐러드', F22: '반달단무지',
        F5: '바지락순두부찌개', F6: '잡곡밥/숭늉', F7: '동그랑땡전', F8: '흑임자무나물', F9: '배추김치',
        F34: '그린샐러드/드레싱2종', F35: '현미밥/청포도에이드',
        F43: '셀프라면/토핑4종/배추김치', F44: '토핑4종(계란,파,치즈,양파)',
      },
      cases: [
        { meal: '중식', day: '월', slot: 1, 이름: '열무김치냉국수', 설명: '열무김치냉국수\n추가밥\n갈비만두\n빵가루마요샐러드\n반달단무지' },
        { meal: '조식', day: '월', slot: 2, 이름: '바지락순두부찌개', 설명: '바지락순두부찌개\n잡곡밥/숭늉\n동그랑땡전\n흑임자무나물\n배추김치' },
        { meal: '석식', day: '월', slot: '라면', 이름: '셀프라면', 설명: '셀프라면\n토핑4종\n배추김치\n토핑4종(계란,파,치즈,양파)' },
        { meal: '중식', day: '월', slot: '샐러드', 이름: '', 설명: '', 지원: false },
      ],
      // 플러스바 공통 필드 검증(코너 아님)
      plusbarCases: [
        { meal: '중식', day: '월', expect: '그린샐러드.드레싱2종.현미밥.청포도에이드' },
        { meal: '석식', day: '월', expect: null }, // 석식엔 플러스바 없음
      ],
    },
  }];

  function activeProfile() {
    return PROFILES.find(p => p.match.test(location.href)) || PROFILES[0];
  }

  // ═══ 상태 ═══
  let WB = null;          // 파싱된 워크북
  let SHEET_NAME = null;  // 본표로 선택된 시트명
  let VALID = null;       // 양식 검증 결과
  const undoStack = [];   // [5] 되돌리기 스냅샷

  // ═══ 공통 유틸 ═══
  function setVal(el, val) {
    if (el == null) return;
    const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, val);
    el.dispatchEvent(new Event('input',  { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }
  function collect() {
    const all = [...document.querySelectorAll('input, textarea')]
      .filter(el => el.tagName === 'TEXTAREA' || el.type === 'text' || el.type === 'number');
    return { board: all.find(el => el.name === 'menuboard_nm') || null, menu: all.filter(el => el.name !== 'menuboard_nm') };
  }
  function cellText(ws, addr) {
    const c = ws && ws[addr];
    if (!c) return '';
    const v = (c.w != null ? c.w : c.v);
    return v == null ? '' : String(v).trim();
  }

  // ═══ 보드명 파싱 (프로파일 규칙 사용) ═══
  function parseBoardName(name, P) {
    P = P || activeProfile();
    const day = (name.match(P.boardName.day) || [])[1] || '';
    let meal = '';
    for (const [re, m] of P.boardName.meal) { if (re.test(name)) { meal = m; break; } }
    let seg = name;
    const pos = name.match(P.boardName.pos);
    if (pos) seg = name.slice(pos.index + pos[0].length);
    seg = seg.replace(P.boardName.dateTail, '');
    const slots = [];
    (seg.match(/\d+/g) || []).forEach(n => slots.push(parseInt(n, 10)));
    if (/라면/.test(seg)) slots.push('라면');
    if (/샐러드/.test(seg)) slots.push('샐러드');
    // 플러스바는 코너(슬롯)가 아니라 각 슬롯 공통 필드 → slots에 넣지 않음
    return { day, meal, slots };
  }

  // ═══ 슬롯 추출 (ws를 인자로 받는 순수함수 — self-test 재사용) ═══
  function extractSlotOn(ws, P, meal, day, slotKey) {
    const rng = P.RANGE[meal] && P.RANGE[meal][slotKey];
    if (!rng) return { 이름: '', 설명: '', 지원: false }; // 샐러드 등 양식 미표기
    const col = P.COL[day];
    if (slotKey === P.ramenKey) {
      // 라면: "셀프라면/토핑4종/배추김치" 슬래시를 줄바꿈으로 분리(원문 재배열)
      const parts = [];
      for (let r = rng[0]; r <= rng[1]; r++) {
        cellText(ws, col + r).split('/').map(s => s.trim()).filter(Boolean).forEach(x => parts.push(x));
      }
      return { 이름: parts[0] || '셀프라면', 설명: parts.join('\n'), 지원: true };
    }
    const vals = [];
    for (let r = rng[0]; r <= rng[1]; r++) { const v = cellText(ws, col + r); if (v) vals.push(v); }
    return { 이름: cellText(ws, col + rng[0]) || (vals[0] || ''), 설명: vals.join('\n'), 지원: true };
  }
  function curWs() { return WB && SHEET_NAME ? WB.Sheets[SHEET_NAME] : null; }

  // 플러스바(공통 필드): 끼니에 플러스바 행이 있으면 R34~R35의 슬래시 항목을 '.'로 연결.
  //   예) '그린샐러드/드레싱2종' + '현미밥/식혜' → '그린샐러드.드레싱2종.현미밥.식혜'
  //   해당 끼니에 플러스바가 없으면(조식/석식) null.
  function plusbarText(ws, P, meal, day) {
    const rng = P.RANGE[meal] && P.RANGE[meal]['플러스바'];
    if (!rng) return null;
    const col = P.COL[day];
    const parts = [];
    for (let r = rng[0]; r <= rng[1]; r++) {
      cellText(ws, col + r).split('/').map(s => s.trim()).filter(Boolean).forEach(x => parts.push(x));
    }
    return parts.join('.');
  }

  // ═══════════════════════════════════════════════════════════════
  // [4] 양식 검증 (앵커 대조) + 본표 시트 자동 선택
  // ═══════════════════════════════════════════════════════════════
  function validateSheet(ws, P) {
    const fail = []; let ok = 0, total = 0;
    for (const [addr, exp] of Object.entries(P.ANCHORS)) {
      total++;
      if (cellText(ws, addr).includes(exp)) ok++; else fail.push(`${addr}≠"${exp}"`);
    }
    return { score: total ? ok / total : 0, ok, total, fail };
  }
  function pickSheet(wb, P) { // 앵커 최고점 시트를 본표로(보존식 등 오선택 방지)
    let best = null;
    wb.SheetNames.forEach(nm => {
      const v = validateSheet(wb.Sheets[nm], P);
      if (!best || v.score > best.v.score) best = { nm, v };
    });
    return best;
  }

  // ═══════════════════════════════════════════════════════════════
  // [7] self-test — mock 워크시트로 extractSlot 회귀검증
  // ═══════════════════════════════════════════════════════════════
  function selfTest() {
    const P = activeProfile();
    const ws = {};
    for (const [k, v] of Object.entries(P.test.mock)) ws[k] = { v };
    const results = P.test.cases.map(c => {
      const r = extractSlotOn(ws, P, c.meal, c.day, c.slot);
      const expName = c.이름 || '', expDesc = c.설명 || '';
      const expSup = c.지원 !== false;
      const pass = r.이름 === expName && r.설명 === expDesc && r.지원 === expSup;
      return { case: `${c.meal}/${c.day}/${c.slot}`, pass, got: r.이름, exp: expName };
    });
    // 플러스바 공통 필드 검증
    (P.test.plusbarCases || []).forEach(c => {
      const got = plusbarText(ws, P, c.meal, c.day);
      results.push({ case: `plusbar/${c.meal}/${c.day}`, pass: got === c.expect, got: String(got), exp: String(c.expect) });
    });
    const passN = results.filter(r => r.pass).length;
    // eslint-disable-next-line no-console
    console.table(results);
    toast(`self-test: ${passN}/${results.length} pass ${passN === results.length ? '✓' : '✗ (콘솔 확인)'}`);
    return results;
  }

  // ═══════════════════════════════════════════════════════════════
  // [5] 계획 생성 → 미리보기 → 적용 → 되돌리기
  // ═══════════════════════════════════════════════════════════════
  function selVal(id, autoVal) {
    const v = document.getElementById(id).value;
    return v === '자동' ? autoVal : v;
  }
  // 기입할 코너(slots)와 끼니/요일 계산 — buildPlan/템플릿 자동선택이 공유
  function computeSlots() {
    const P = activeProfile();
    const { board } = collect();
    if (!WB) return { error: '먼저 주간메뉴표 xlsx 파일을 선택하세요.' };
    if (!board) return { error: '메뉴보드 편집 화면에서 실행하세요.' };
    if (VALID && VALID.score < 0.7) return { error: `양식 불일치(${VALID.ok}/${VALID.total}) — 다른 양식일 수 있습니다: ${VALID.fail.slice(0,3).join(', ')}` };

    const parsed = parseBoardName(board.value || '', P);
    const meal = selVal('__maf_meal', parsed.meal);
    const day  = selVal('__maf_day',  parsed.day);
    if (!meal || !day) return { error: `끼니/요일 인식 실패(끼니/요일 셀렉트로 지정): "${board.value}"` };

    let slots;
    if (document.getElementById('__maf_cornermode').value === '수동') {
      slots = [...document.querySelectorAll('.__maf_slot')].map(s => {
        const v = s.value;
        return v === '' ? '' : (/^\d+$/.test(v) ? parseInt(v, 10) : v);
      }).filter(v => v !== '');
      if (!slots.length) return { error: '직접 지정 모드: 슬롯 코너를 1개 이상 선택하세요.' };
    } else {
      slots = parsed.slots;
      if (!slots.length) return { error: `보드명에서 코너 인식 실패(직접 지정 사용): "${board.value}"` };
    }
    return { P, meal, day, slots };
  }

  // 템플릿 라디오(1메뉴/2메뉴)를 필요한 코너 수에 맞게 자동 선택
  function selectTemplate(n) {
    const r = [...document.querySelectorAll('input[type=radio]')].find(x => x.name === '' && x.value === String(n));
    if (r && !r.checked) { r.click(); return true; } // Vue @change 반영 위해 click
    return false;
  }
  function waitFor(cond, ms) {
    return new Promise(res => {
      const t0 = Date.now();
      (function loop() { if (cond() || Date.now() - t0 > ms) return res(); setTimeout(loop, 100); })();
    });
  }
  // 코너 수 = 1 또는 2면 해당 템플릿으로 전환 후 슬롯 생성 대기
  async function ensureTemplate(need) {
    if (need < 1 || need > 2) return;                       // 1·2메뉴만 자동전환
    const cur = Math.floor(collect().menu.length / 6);
    if (cur === need) return;
    if (selectTemplate(need)) {
      await waitFor(() => Math.floor(collect().menu.length / 6) === need, 3000);
    }
  }

  function buildPlan() {
    const c = computeSlots();
    if (c.error) return c;
    const { P, meal, day, slots } = c;
    const { menu } = collect();
    const slotCount = Math.floor(menu.length / 6);
    if (slotCount === 0) return { error: '메뉴 인풋을 찾지 못했습니다.' };

    const ws = curWs();
    const n = Math.min(slotCount, slots.length);
    const rows = [];
    for (let i = 0; i < n; i++) {
      const key = slots[i];
      const s = extractSlotOn(ws, P, meal, day, key);
      const base = i * 6;
      rows.push({
        idx: i, key, 지원: s.지원,
        신규이름: s.이름, 신규설명: s.설명,
        기존이름: menu[base + 0] ? menu[base + 0].value : '',
        기존설명: menu[base + 2] ? menu[base + 2].value : '',
      });
    }
    const plusbar = plusbarText(ws, P, meal, day); // 공통 플러스바(없으면 null)
    return { meal, day, slots, slotCount, rows, menu, plusbar };
  }

  function applyPlan(plan) {
    const { menu, rows, plusbar, slotCount } = plan;
    undoStack.push(menu.map(el => el.value)); // 스냅샷
    if (undoStack.length > 20) undoStack.shift();
    let done = 0;
    const offSlots = new Set(); // 사용자가 체크 해제한 지원 슬롯(플러스바까지 보존)
    rows.forEach(r => {
      if (!r.지원) return; // 미지원(샐러드)은 이름/설명 보존 · 플러스바는 기존대로 기입
      if (r.selected === false) { offSlots.add(r.idx); return; } // 체크 해제 → 슬롯 전체 보존
      const base = r.idx * 6;
      setVal(menu[base + 0], r.신규이름);
      setVal(menu[base + 2], r.신규설명);
      done++;
    });
    // 플러스바(공통): 코너와 무관하게 모든 슬롯의 플러스바 필드(pos5)에 동일 기입.
    //   단, 사용자가 체크 해제한 슬롯은 해당 슬롯 전체 보존 취지이므로 플러스바도 건너뜀.
    let pb = 0;
    if (plusbar != null && plusbar !== '') {
      for (let i = 0; i < slotCount; i++) { if (offSlots.has(i)) continue; setVal(menu[i * 6 + 5], plusbar); pb++; }
    }
    updateUndoBtn();
    const slotTxt = plan.slots.map(k => (typeof k === 'number' ? '코너' + k : k)).join('·');
    const skip = plan.rows.filter(r => !r.지원).map(r => (typeof r.key === 'number' ? '코너' + r.key : r.key));
    const skipTxt = skip.length ? ` · 미지원(수동): ${skip.join(', ')}` : '';
    const off = plan.rows.filter(r => r.지원 && r.selected === false).map(r => (typeof r.key === 'number' ? '코너' + r.key : r.key));
    const offTxt = off.length ? ` · 제외(보존): ${off.join(', ')}` : '';
    const pbTxt = pb ? ` · 플러스바 공통 ${pb}칸` : '';
    const over = plan.slots.length > plan.slotCount ? ` · 앞 ${plan.slotCount}개만` : '';
    toast(`${plan.meal} ${plan.day}요일 · [${slotTxt}] ${done}칸${pbTxt} (Save 미클릭)${skipTxt}${offTxt}${over}`);
  }

  function undo() {
    const snap = undoStack.pop();
    if (!snap) { toast('되돌릴 내역이 없습니다.'); return; }
    const { menu } = collect();
    snap.forEach((v, i) => { if (menu[i]) setVal(menu[i], v); });
    updateUndoBtn();
    toast('직전 상태로 되돌렸습니다.');
  }
  function updateUndoBtn() {
    const b = document.getElementById('__maf_undo');
    if (b) { b.disabled = undoStack.length === 0; b.style.opacity = undoStack.length ? '1' : '.5'; }
  }

  // 미리보기 모달
  async function showPreview() {
    // 코너 수에 맞춰 템플릿(1메뉴/2메뉴) 자동 선택 후 슬롯 생성 대기
    const c = computeSlots();
    if (c.error) { toast(c.error); return; }
    await ensureTemplate(c.slots.length);

    const plan = buildPlan(); // 템플릿 전환 후 재수집 기준
    if (plan.error) { toast(plan.error); return; }
    document.getElementById('__maf_modal')?.remove();
    // 코너별 기입 선택: 기본 ON. 지원되는 코너만 토글 가능(미지원=항상 보존).
    plan.rows.forEach(r => { if (r.selected === undefined) r.selected = r.지원; });
    const esc = s => (s || '').replace(/</g, '&lt;').replace(/\n/g, '<br>');
    const rowHtml = plan.rows.map((r, i) => {
      const changed = (r.기존이름 || r.기존설명);
      const color = !r.지원 ? '#999' : (changed ? '#b58900' : '#2a2');
      const badge = !r.지원 ? '미지원' : (changed ? '덮어씀⚠' : '신규');
      const chk = r.지원
        ? `<input type="checkbox" class="__maf_pick" data-i="${i}" ${r.selected !== false ? 'checked' : ''} title="이 코너 기입" style="width:16px;height:16px;cursor:pointer">`
        : `<input type="checkbox" disabled title="미지원(항상 보존)" style="width:16px;height:16px">`;
      return `<tr style="border-top:1px solid #eee;color:${color}">
        <td style="padding:4px 6px;text-align:center">${chk}</td>
        <td style="padding:4px 6px;white-space:nowrap">${typeof r.key === 'number' ? '코너' + r.key : r.key}<br><small>${badge}</small></td>
        <td style="padding:4px 6px"><b>${esc(r.신규이름)}</b><br><small>${esc(r.신규설명)}</small></td>
        <td style="padding:4px 6px;color:#999"><small>${esc(r.기존이름)}<br>${esc(r.기존설명)}</small></td></tr>`;
    }).join('');
    // 플러스바 공통 행(있을 때만) — 체크 해제한 슬롯은 플러스바도 보존됨
    const pbRow = (plan.plusbar != null && plan.plusbar !== '')
      ? `<tr style="border-top:2px solid #00838f;color:#00838f;background:#f7fdfe">
           <td style="padding:4px 6px;text-align:center"><small>—</small></td>
           <td style="padding:4px 6px;white-space:nowrap">플러스바<br><small>공통 ${plan.slotCount}칸</small></td>
           <td style="padding:4px 6px" colspan="2"><b>${esc(plan.plusbar)}</b> <small>(선택 슬롯 동일)</small></td></tr>`
      : '';
    const m = document.createElement('div');
    m.id = '__maf_modal';
    m.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);z-index:100001;background:#fff;border-radius:10px;box-shadow:0 8px 30px rgba(0,0,0,.4);width:560px;max-width:92vw;max-height:80vh;overflow:auto;font-size:13px;font-family:sans-serif';
    m.innerHTML = `
      <div style="padding:10px 14px;font-weight:700;background:#00838f;color:#fff;border-radius:10px 10px 0 0">
        미리보기 — ${plan.meal} ${plan.day}요일 · ${plan.slots.map(k => typeof k === 'number' ? '코너' + k : k).join('·')}</div>
      <table style="width:100%;border-collapse:collapse">
        <tr style="background:#f4f4f4">
          <th style="padding:4px 6px;text-align:center" title="전체 선택/해제"><input type="checkbox" id="__maf_pickall" style="width:16px;height:16px;cursor:pointer"></th>
          <th style="padding:4px 6px;text-align:left">코너</th><th style="padding:4px 6px;text-align:left">신규(→기입)</th><th style="padding:4px 6px;text-align:left">기존</th></tr>
        ${rowHtml}${pbRow}</table>
      <div style="padding:10px 14px;text-align:right;border-top:1px solid #eee">
        <button id="__maf_cancel" style="padding:7px 14px;margin-right:6px;border:1px solid #ccc;background:#fff;border-radius:6px;cursor:pointer">취소</button>
        <button id="__maf_apply" style="padding:7px 16px;border:0;background:#00bcd4;color:#fff;border-radius:6px;font-weight:600;cursor:pointer">적용</button></div>`;
    document.body.appendChild(m);

    // 코너별 체크박스 ↔ plan.rows[i].selected 동기화
    const picks = [...m.querySelectorAll('.__maf_pick')];
    const applyBtn = m.querySelector('#__maf_apply');
    const pickAll = m.querySelector('#__maf_pickall');
    const syncPickAll = () => {
      if (!pickAll) return;
      const on = picks.filter(cb => cb.checked).length;
      pickAll.checked = on === picks.length && picks.length > 0;
      pickAll.indeterminate = on > 0 && on < picks.length;
    };
    const syncApplyBtn = () => {
      // 지원 코너가 하나라도 있는 화면에서 전부 해제하면 적용 무의미 → 버튼 비활성
      const anyOn = picks.some(cb => cb.checked);
      if (picks.length) { applyBtn.disabled = !anyOn; applyBtn.style.opacity = anyOn ? '1' : '.5'; applyBtn.style.cursor = anyOn ? 'pointer' : 'not-allowed'; }
    };
    picks.forEach(cb => cb.addEventListener('change', () => {
      plan.rows[+cb.dataset.i].selected = cb.checked;
      syncPickAll(); syncApplyBtn();
    }));
    if (pickAll) pickAll.addEventListener('change', () => {
      picks.forEach(cb => { cb.checked = pickAll.checked; plan.rows[+cb.dataset.i].selected = pickAll.checked; });
      pickAll.indeterminate = false; syncApplyBtn();
    });
    syncPickAll(); syncApplyBtn();

    m.querySelector('#__maf_cancel').onclick = () => m.remove();
    applyBtn.onclick = () => { if (applyBtn.disabled) return; applyPlan(plan); m.remove(); };
  }

  // ═══ 토스트 ═══
  function toast(msg) {
    let t = document.getElementById('__maf_toast');
    if (!t) {
      t = document.createElement('div');
      t.id = '__maf_toast';
      t.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);z-index:100002;background:#333;color:#fff;padding:14px 20px;border-radius:8px;font-size:14px;max-width:400px;text-align:center;box-shadow:0 4px 16px rgba(0,0,0,.35);opacity:0;transition:opacity .2s';
      document.body.appendChild(t);
    }
    t.textContent = msg;
    t.style.opacity = '1';
    clearTimeout(t._h);
    t._h = setTimeout(() => (t.style.opacity = '0'), 5000);
  }

  // ═══ 패널 ═══
  function mountPanel() {
    if (document.getElementById('__maf_panel')) return;
    const p = document.createElement('div');
    p.id = '__maf_panel';
    p.style.cssText = 'position:fixed;bottom:20px;left:20px;z-index:99999;background:#fff;border:1px solid #ccc;border-radius:10px;padding:12px;width:240px;box-shadow:0 4px 14px rgba(0,0,0,.2);font-size:13px;font-family:sans-serif';
    p.innerHTML = `
      <div id="__maf_head" title="더블클릭: self-test" style="font-weight:700;margin:-12px -12px 8px;padding:8px 12px;color:#00838f;cursor:move;user-select:none;border-bottom:1px solid #eee;background:#f7fdfe;border-radius:10px 10px 0 0">메뉴 자동기입 <span style="float:right;color:#bbb;font-weight:400">⠿</span></div>
      <label style="display:block;margin-bottom:6px">주간메뉴표
        <input type="file" id="__maf_file" accept=".xlsx" style="display:block;margin-top:3px;width:100%"></label>
      <div id="__maf_fname" style="color:#888;font-size:11px;margin-bottom:4px;word-break:break-all">파일 미선택</div>
      <div id="__maf_valid" style="font-size:11px;margin-bottom:4px;min-height:14px"></div>
      <div id="__maf_detect" style="color:#00838f;font-size:11px;margin-bottom:8px;min-height:14px"></div>
      <label style="display:block;margin-bottom:6px">끼니
        <select id="__maf_meal" style="width:100%;margin-top:3px;padding:3px">
          <option value="자동" selected>자동(보드명)</option><option>조식</option><option>중식</option><option>석식</option></select></label>
      <label style="display:block;margin-bottom:6px">요일
        <select id="__maf_day" style="width:100%;margin-top:3px;padding:3px">
          <option value="자동" selected>자동(보드명)</option><option>월</option><option>화</option><option>수</option><option>목</option><option>금</option></select></label>
      <label style="display:block;margin-bottom:4px">코너
        <select id="__maf_cornermode" style="width:100%;margin-top:3px;padding:3px">
          <option value="자동" selected>자동(보드명)</option><option value="수동">직접 지정</option></select></label>
      <div id="__maf_slots" style="display:none;margin-bottom:8px"></div>
      <button id="__maf_run" style="width:100%;background:#00bcd4;color:#fff;border:0;padding:9px;border-radius:6px;font-weight:600;cursor:pointer;margin:4px 0 6px">미리보기 → 기입</button>
      <button id="__maf_undo" disabled style="width:100%;background:#eee;color:#333;border:0;padding:7px;border-radius:6px;cursor:pointer;opacity:.5">되돌리기</button>`;
    document.body.appendChild(p);

    document.getElementById('__maf_file').addEventListener('change', e => {
      const f = e.target.files[0]; if (!f) return;
      const fr = new FileReader();
      fr.onload = () => {
        try {
          WB = XLSX.read(new Uint8Array(fr.result), { type: 'array' });
          const P = activeProfile();
          const best = pickSheet(WB, P);
          SHEET_NAME = best.nm; VALID = best.v;
          document.getElementById('__maf_fname').textContent = f.name + ' · 시트: ' + SHEET_NAME;
          renderValid();
          toast('파일 로드: ' + f.name + ' · 양식 ' + best.v.ok + '/' + best.v.total);
        } catch (err) {
          WB = null; SHEET_NAME = null; VALID = null;
          document.getElementById('__maf_fname').textContent = '읽기 실패';
          document.getElementById('__maf_valid').textContent = '';
          toast('xlsx 파싱 실패(SheetJS 로드/파일 확인): ' + err.message);
        }
      };
      fr.readAsArrayBuffer(f);
    });
    document.getElementById('__maf_run').addEventListener('click', showPreview);
    document.getElementById('__maf_undo').addEventListener('click', undo);
    document.getElementById('__maf_head').addEventListener('dblclick', selfTest);
    document.getElementById('__maf_cornermode').addEventListener('change', rebuildSlotSelects);
    document.getElementById('__maf_meal').addEventListener('change', rebuildSlotSelects);
    makeDraggable(p, document.getElementById('__maf_head'));
    refreshDetect(); renderValid(); updateUndoBtn(); rebuildSlotSelects();
  }

  // 코너 직접 지정: 끼니의 코너 목록 + 샐러드를 슬롯 수만큼 드롭다운 생성
  function rebuildSlotSelects() {
    const box = document.getElementById('__maf_slots');
    const modeEl = document.getElementById('__maf_cornermode');
    if (!box || !modeEl) return;
    if (modeEl.value !== '수동') { box.style.display = 'none'; box.innerHTML = ''; return; }
    box.style.display = 'block';
    const P = activeProfile();
    const { board, menu } = collect();
    const parsed = board ? parseBoardName(board.value || '', P) : { meal: '', slots: [] };
    const meal = selVal('__maf_meal', parsed.meal);
    const slotCount = Math.max(1, Math.floor(menu.length / 6) || 2);
    // 플러스바는 코너가 아니라 공통 필드이므로 슬롯 코너 옵션에서 제외
    const keys = (meal && P.RANGE[meal]) ? Object.keys(P.RANGE[meal]).filter(k => k !== '플러스바') : [];
    const label = k => (/^\d+$/.test(k) ? '코너' + k : k);
    const prev = [...box.querySelectorAll('.__maf_slot')].map(s => s.value);
    const optHtml = i => `<option value="">슬롯${i + 1}: 미지정</option>` +
      keys.map(k => `<option value="${k}">슬롯${i + 1}: ${label(k)}</option>`).join('') +
      `<option value="샐러드">슬롯${i + 1}: 샐러드(수동)</option>`;
    let html = '';
    for (let i = 0; i < slotCount; i++) html += `<select class="__maf_slot" style="width:100%;margin:2px 0;padding:3px">${optHtml(i)}</select>`;
    box.innerHTML = html;
    // 이전 선택 + 보드명 자동값으로 초기화
    const slotEls = [...box.querySelectorAll('.__maf_slot')];
    slotEls.forEach((sel, i) => {
      const auto = parsed.slots[i];
      const want = prev[i] || (auto != null ? String(auto) : '');
      if ([...sel.options].some(o => o.value === want)) sel.value = want;
    });
  }

  function renderValid() {
    const el = document.getElementById('__maf_valid'); if (!el) return;
    if (!VALID) { if (el.textContent !== '') el.textContent = ''; return; }
    const good = VALID.score >= 0.7;
    const txt = `양식 ${good ? '✓' : '✗'} ${VALID.ok}/${VALID.total}` + (good ? '' : ' — 다른 양식?');
    if (el.textContent !== txt) { el.textContent = txt; el.style.color = good ? '#2a2' : '#c00'; }
  }
  function refreshDetect() {
    const el = document.getElementById('__maf_detect'); if (!el) return;
    const { board } = collect();
    if (!board || !board.value) { if (el.textContent !== '보드명 없음') el.textContent = '보드명 없음'; return; }
    const p = parseBoardName(board.value);
    const slotTxt = p.slots.map(k => (typeof k === 'number' ? '코너' + k : k)).join('·') || '?';
    const txt = `감지: ${p.meal || '?'} / ${p.day || '?'}요일 / ${slotTxt}`;
    if (el.textContent !== txt) el.textContent = txt;
  }

  // 드래그
  function makeDraggable(panel, handle) {
    let sx = 0, sy = 0, ox = 0, oy = 0;
    handle.addEventListener('mousedown', e => {
      e.preventDefault();
      const r = panel.getBoundingClientRect();
      panel.style.left = r.left + 'px'; panel.style.top = r.top + 'px';
      panel.style.right = 'auto'; panel.style.bottom = 'auto';
      sx = e.clientX; sy = e.clientY; ox = r.left; oy = r.top;
      document.addEventListener('mousemove', onMove); document.addEventListener('mouseup', onUp);
    });
    function onMove(e) {
      const maxX = window.innerWidth - panel.offsetWidth, maxY = window.innerHeight - panel.offsetHeight;
      panel.style.left = Math.max(0, Math.min(maxX, ox + e.clientX - sx)) + 'px';
      panel.style.top  = Math.max(0, Math.min(maxY, oy + e.clientY - sy)) + 'px';
    }
    function onUp() { document.removeEventListener('mousemove', onMove); document.removeEventListener('mouseup', onUp); }
  }

  // SPA(hash 라우팅) 대응: 폴링(MutationObserver는 refreshDetect의 DOM 수정과 피드백 루프)
  mountPanel();
  setInterval(() => { mountPanel(); refreshDetect(); renderValid(); }, 1000);
})();
