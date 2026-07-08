// ==UserScript==
// @name         CloudCast 메뉴보드 자동기입 (식단표 연동)
// @namespace    local.cloudcast.menuautofill
// @version      3.2
// @description  주간메뉴표 xlsx 업로드 → 보드명(요일·끼니·코너)을 자동 파싱해 해당 코너 메뉴를 메뉴보드 인풋에 자동기입. Save 버튼은 절대 누르지 않음.
// @match        https://apps.cloud-cast.com/menuboard/menu/*
// @require      https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js
// @run-at       document-idle
// @grant        none
// ==/UserScript==

(function () {
  'use strict';

  // ─────────────────────────────────────────────────────────────
  // 식단표(주간메뉴표 xlsx) 셀맵 — NH농협타워점 출력양식 기준(4유형 인스턴스로 검증)
  //   요일 = 열(F:월 ~ J:금)
  //   끼니·코너 = 행 범위 [시작(메인), 끝]
  //   이름 = 메인행 값,  설명 = 범위 내 비어있지 않은 셀 전체(메인 포함) 줄바꿈 결합
  // ─────────────────────────────────────────────────────────────
  const COL = { 월: 'F', 화: 'G', 수: 'H', 목: 'I', 금: 'J' };

  // 끼니 → { 코너키: [시작행, 끝행] }.  코너키: 숫자(코너N) 또는 '라면'
  const RANGE = {
    조식: { 2: [5, 9],  3: [10, 14], 라면: [15, 16] },
    중식: { 1: [18, 22], 2: [23, 28], 3: [29, 33], 플러스바: [34, 35] },
    석식: { 2: [37, 42], 라면: [43, 44] },
  };

  let WB = null;

  // ── Vue v-model 반영: native setter 후 input/change 디스패치 ──
  function setVal(el, val) {
    if (el == null) return;
    const proto = el.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
    Object.getOwnPropertyDescriptor(proto, 'value').set.call(el, val);
    el.dispatchEvent(new Event('input',  { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }

  // ── 메뉴 인풋 수집 (CSS [type=text]는 type 미명시 인풋을 놓치므로 property 필터) ──
  function collect() {
    const all = [...document.querySelectorAll('input, textarea')]
      .filter(el => el.tagName === 'TEXTAREA' || el.type === 'text' || el.type === 'number');
    return {
      board: all.find(el => el.name === 'menuboard_nm') || null,
      menu:  all.filter(el => el.name !== 'menuboard_nm'),
    };
  }

  function cellText(ws, addr) {
    const c = ws[addr];
    if (!c) return '';
    const v = (c.w != null ? c.w : c.v);
    return v == null ? '' : String(v).trim();
  }

  // ── 보드명 파싱 → { day, meal, slots } ──
  //   예) "(금)석식_(오른쪽)코너2,라면_7/10" → {day:'금', meal:'석식', slots:[2,'라면']}
  //   예) "(화)중_(왼쪽)_코너1,3 _7/7"      → {day:'화', meal:'중식', slots:[1,3]}
  function parseBoardName(name) {
    const day = (name.match(/\(([월화수목금])\)/) || [])[1] || '';
    const meal = /조식/.test(name) ? '조식'
               : /석식/.test(name) ? '석식'
               : /중식|중/.test(name) ? '중식' : '';
    // 코너 조합 구간: 위치표기(왼쪽/오른쪽) 뒤 ~ 날짜(_7/7) 앞
    let seg = name;
    const pos = name.match(/\((?:왼쪽|오른쪽)\)/);
    if (pos) seg = name.slice(pos.index + pos[0].length);
    seg = seg.replace(/_?\s*\d+\/\d+.*$/, ''); // 뒤쪽 날짜 제거
    const slots = [];
    (seg.match(/\d+/g) || []).forEach(n => slots.push(parseInt(n, 10))); // 코너 숫자들
    if (/라면/.test(seg)) slots.push('라면');
    if (/샐러드/.test(seg)) slots.push('샐러드');
    return { day, meal, slots };
  }

  // ── (끼니, 요일, 슬롯키) → { 이름, 설명, 지원여부 } ──
  function extractSlot(meal, day, slotKey) {
    const rng = RANGE[meal] && RANGE[meal][slotKey];
    if (!rng) return { 이름: '', 설명: '', 지원: false }; // 샐러드 등 양식 미표기
    const col = COL[day];
    const ws = WB.Sheets[WB.SheetNames[0]];

    // 라면: 식단표는 "셀프라면/토핑4종/배추김치"(슬래시) + "토핑4종(…)" 2행 →
    //       슬래시를 줄바꿈으로 분리해 가독성 확보(원문 재배열, 항목 창작 없음)
    if (slotKey === '라면') {
      const parts = [];
      for (let r = rng[0]; r <= rng[1]; r++) {
        cellText(ws, col + r).split('/').map(s => s.trim()).filter(Boolean).forEach(x => parts.push(x));
      }
      return { 이름: parts[0] || '셀프라면', 설명: parts.join('\n'), 지원: true };
    }

    const vals = [];
    for (let r = rng[0]; r <= rng[1]; r++) {
      const v = cellText(ws, col + r);
      if (v) vals.push(v);
    }
    return { 이름: cellText(ws, col + rng[0]) || (vals[0] || ''), 설명: vals.join('\n'), 지원: true };
  }

  // ── 자동기입 ──
  function fill() {
    if (!WB) { toast('먼저 주간메뉴표 xlsx 파일을 선택하세요.'); return; }
    const { board, menu } = collect();
    if (!board) { toast('메뉴보드 편집 화면에서 실행하세요.'); return; }

    const parsed = parseBoardName(board.value || '');
    // 끼니/요일은 보드명 자동값 기준, 셀렉트에서 수동 오버라이드 허용
    const mealSel = document.getElementById('__maf_meal').value;
    const daySel  = document.getElementById('__maf_day').value;
    const meal = mealSel === '자동' ? parsed.meal : mealSel;
    const day  = daySel  === '자동' ? parsed.day  : daySel;

    if (!meal || !day) { toast(`보드명에서 끼니/요일 인식 실패: "${board.value}"`); return; }
    if (!parsed.slots.length) { toast(`보드명에서 코너 인식 실패: "${board.value}"`); return; }

    const slotCount = Math.floor(menu.length / 6); // 메뉴보드 코너 슬롯 수
    if (slotCount === 0) { toast('메뉴 인풋을 찾지 못했습니다.'); return; }

    const n = Math.min(slotCount, parsed.slots.length);
    let done = 0; const skipped = [];
    for (let i = 0; i < n; i++) {
      const key = parsed.slots[i];
      const s = extractSlot(meal, day, key);
      // 미지원 슬롯(샐러드 등)은 식단표에 데이터가 없으므로 기존 입력을 지우지 않고 건너뜀
      if (!s.지원) { skipped.push(typeof key === 'number' ? '코너' + key : String(key)); continue; }
      const base = i * 6; // [이름,가격,설명,영양,맵기,플러스바]
      setVal(menu[base + 0], s.이름);   // 메뉴 이름 = 메인
      setVal(menu[base + 2], s.설명);   // 메뉴 설명 = 코너 전체 구성
      // 가격/영양/맵기/플러스바는 식단표에 데이터가 없어 건드리지 않음
      done++;
    }
    const slotTxt = parsed.slots.map(k => (typeof k === 'number' ? '코너' + k : k)).join('·');
    const skipTxt = skipped.length ? ` · 미지원(식단표無, 수동입력): ${skipped.join(', ')}` : '';
    const overTxt = parsed.slots.length > slotCount ? ` · 코너 ${parsed.slots.length}개 중 앞 ${slotCount}개만(슬롯 부족)` : '';
    toast(`${meal} ${day}요일 · [${slotTxt}] ${done}칸 기입 (Save 미클릭)${skipTxt}${overTxt}`);
  }

  // ── UI ──
  function toast(msg) {
    let t = document.getElementById('__maf_toast');
    if (!t) {
      t = document.createElement('div');
      t.id = '__maf_toast';
      t.style.cssText = 'position:fixed;top:50%;left:50%;transform:translate(-50%,-50%);z-index:100000;background:#333;color:#fff;padding:14px 20px;border-radius:8px;font-size:14px;max-width:380px;text-align:center;box-shadow:0 4px 16px rgba(0,0,0,.35);opacity:0;transition:opacity .2s';
      document.body.appendChild(t);
    }
    t.textContent = msg;
    t.style.opacity = '1';
    clearTimeout(t._h);
    t._h = setTimeout(() => (t.style.opacity = '0'), 5000);
  }

  function mountPanel() {
    if (document.getElementById('__maf_panel')) return;
    const p = document.createElement('div');
    p.id = '__maf_panel';
    p.style.cssText = 'position:fixed;bottom:20px;left:20px;z-index:99999;background:#fff;border:1px solid #ccc;border-radius:10px;padding:12px;width:236px;box-shadow:0 4px 14px rgba(0,0,0,.2);font-size:13px;font-family:sans-serif';
    p.innerHTML = `
      <div id="__maf_head" style="font-weight:700;margin:-12px -12px 8px;padding:8px 12px;color:#00838f;cursor:move;user-select:none;border-bottom:1px solid #eee;background:#f7fdfe;border-radius:10px 10px 0 0">메뉴 자동기입 <span style="float:right;color:#bbb;font-weight:400">⠿</span></div>
      <label style="display:block;margin-bottom:6px">주간메뉴표
        <input type="file" id="__maf_file" accept=".xlsx" style="display:block;margin-top:3px;width:100%">
      </label>
      <div id="__maf_fname" style="color:#888;font-size:11px;margin-bottom:6px;word-break:break-all">파일 미선택</div>
      <div id="__maf_detect" style="color:#00838f;font-size:11px;margin-bottom:8px;min-height:14px"></div>
      <label style="display:block;margin-bottom:6px">끼니
        <select id="__maf_meal" style="width:100%;margin-top:3px;padding:3px">
          <option value="자동" selected>자동(보드명)</option>
          <option value="조식">조식</option>
          <option value="중식">중식</option>
          <option value="석식">석식</option>
        </select>
      </label>
      <label style="display:block;margin-bottom:10px">요일
        <select id="__maf_day" style="width:100%;margin-top:3px;padding:3px">
          <option value="자동" selected>자동(보드명)</option>
          <option>월</option><option>화</option><option>수</option><option>목</option><option>금</option>
        </select>
      </label>
      <button id="__maf_run" style="width:100%;background:#00bcd4;color:#fff;border:0;padding:9px;border-radius:6px;font-weight:600;cursor:pointer">자동기입</button>`;
    document.body.appendChild(p);

    document.getElementById('__maf_file').addEventListener('change', e => {
      const f = e.target.files[0];
      if (!f) return;
      const fr = new FileReader();
      fr.onload = () => {
        try {
          WB = XLSX.read(new Uint8Array(fr.result), { type: 'array' });
          document.getElementById('__maf_fname').textContent = f.name + ' · 시트: ' + WB.SheetNames[0];
          toast('파일 로드 완료: ' + f.name);
          refreshDetect();
        } catch (err) {
          WB = null;
          document.getElementById('__maf_fname').textContent = '읽기 실패';
          toast('xlsx 파싱 실패: ' + err.message);
        }
      };
      fr.readAsArrayBuffer(f);
    });
    document.getElementById('__maf_run').addEventListener('click', fill);
    makeDraggable(p, document.getElementById('__maf_head'));
    refreshDetect();
  }

  // 현재 보드명 파싱 결과를 패널에 표시
  function refreshDetect() {
    const el = document.getElementById('__maf_detect');
    if (!el) return;
    const { board } = collect();
    if (!board || !board.value) { el.textContent = '보드명 없음'; return; }
    const p = parseBoardName(board.value);
    const slotTxt = p.slots.map(k => (typeof k === 'number' ? '코너' + k : k)).join('·') || '?';
    const txt = `감지: ${p.meal || '?'} / ${p.day || '?'}요일 / ${slotTxt}`;
    if (el.textContent !== txt) el.textContent = txt; // 동일값이면 DOM 미변경(루프 방지)
  }

  // 헤더를 잡고 패널 이동 (right/bottom 고정 → left/top 절대좌표로 전환)
  function makeDraggable(panel, handle) {
    let sx = 0, sy = 0, ox = 0, oy = 0;
    handle.addEventListener('mousedown', e => {
      e.preventDefault();
      const r = panel.getBoundingClientRect();
      panel.style.left = r.left + 'px';
      panel.style.top = r.top + 'px';
      panel.style.right = 'auto';
      panel.style.bottom = 'auto';
      sx = e.clientX; sy = e.clientY; ox = r.left; oy = r.top;
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });
    function onMove(e) {
      const maxX = window.innerWidth - panel.offsetWidth;
      const maxY = window.innerHeight - panel.offsetHeight;
      panel.style.left = Math.max(0, Math.min(maxX, ox + e.clientX - sx)) + 'px';
      panel.style.top  = Math.max(0, Math.min(maxY, oy + e.clientY - sy)) + 'px';
    }
    function onUp() {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    }
  }

  // SPA(hash 라우팅) 화면 전환 대응: MutationObserver는 refreshDetect의 DOM 수정과
  // 피드백 루프를 일으키므로 사용하지 않고, 가벼운 폴링으로 패널 유지 + 감지 갱신.
  mountPanel();
  setInterval(() => { mountPanel(); refreshDetect(); }, 1000);
})();
