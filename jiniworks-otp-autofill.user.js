// ==UserScript==
// @name         Jiniworks 관리자 OTP 자동입력 (kihf)
// @namespace    local.jiniworks.otpautofill
// @version      1.0
// @description  kihf.jiniworks.com 관리자 로그인 화면에서 TOTP(시간기반 OTP)를 등록된 시크릿으로 자동 생성해 OTP 입력칸에 채워줌. 시크릿은 브라우저(GM 저장소)에만 보관. 자동제출은 기본 OFF.
// @match        https://kihf.jiniworks.com/_fox/login.do*
// @match        https://kihf.jiniworks.com/_fox/*login*
// @run-at       document-idle
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_deleteValue
// @grant        GM_registerMenuCommand
// ==/UserScript==

/*
 * ────────────────────────────────────────────────────────────────────────────
 *  동작 원리
 *  - OTP(일회용 비밀번호)는 "1회성" 값이라, 스크립트가 스스로 "자동입력"하려면
 *    인증앱(Google Authenticator / OTP 등)과 동일한 TOTP 시크릿을 알고 있어야
 *    현재 시각의 6자리 코드를 직접 계산할 수 있습니다.
 *  - 이 스크립트는 처음 1번만 시크릿을 등록받아 GM 저장소에 보관하고,
 *    이후 로그인 화면이 뜨면 현재 코드(30초 주기)를 계산해 OTP 칸에 채웁니다.
 *
 *  ⚠️ 보안 주의
 *  - 시크릿을 브라우저에 저장하면 사실상 2단계 인증의 "두 번째 요소"가
 *    아이디/비번과 같은 기기에 함께 놓이게 됩니다. 편의를 위한 것이며,
 *    본인 소유의 관리자 계정 · 신뢰하는 개인 PC에서만 사용하세요.
 *  - 시크릿은 GM 저장소(도메인/스크립트 스코프)에만 저장되고 외부로 전송하지
 *    않습니다.
 *
 *  🔧 최초 설정 (딱 1번)
 *  1) Tampermonkey 아이콘 ▸ 이 스크립트 메뉴 ▸ "① OTP 시크릿 등록"
 *  2) OTP 등록 시 받은 Base32 시크릿 또는 otpauth:// URL을 붙여넣기
 *     (인증앱 최초 등록 때의 그 값 — QR 아래 "키를 직접 입력" 문자열)
 *  3) 완료. 로그인 화면 우하단 패널에 현재 코드가 표시되고 자동 입력됩니다.
 *
 *  🔎 OTP 입력칸이 자동 감지되지 않으면
 *  - 메뉴 ▸ "③ OTP 입력칸 선택자 지정"에서 CSS 선택자(예: #otpNo)를 넣으세요.
 * ────────────────────────────────────────────────────────────────────────────
 */

(function () {
  'use strict';

  const KEY = {
    secret: 'otp_secret_b32',
    digits: 'otp_digits',
    period: 'otp_period',
    algo: 'otp_algo',
    selector: 'otp_selector',
    autosubmit: 'otp_autosubmit',
  };

  // ── 설정 로드 ──
  const cfg = () => ({
    secret: (GM_getValue(KEY.secret, '') || '').trim(),
    digits: parseInt(GM_getValue(KEY.digits, 6), 10) || 6,
    period: parseInt(GM_getValue(KEY.period, 30), 10) || 30,
    algo: GM_getValue(KEY.algo, 'SHA-1'),
    selector: (GM_getValue(KEY.selector, '') || '').trim(),
    autosubmit: !!GM_getValue(KEY.autosubmit, false),
  });

  // ────────────────────────── TOTP 계산 (RFC 6238) ──────────────────────────
  function base32ToBytes(b32) {
    const A = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ234567';
    const clean = String(b32).toUpperCase().replace(/=+$/, '').replace(/[^A-Z2-7]/g, '');
    let bits = '';
    for (const c of clean) bits += A.indexOf(c).toString(2).padStart(5, '0');
    const bytes = [];
    for (let i = 0; i + 8 <= bits.length; i += 8) bytes.push(parseInt(bits.slice(i, i + 8), 2));
    return new Uint8Array(bytes);
  }

  async function totp(secretB32, opt) {
    const { digits = 6, period = 30, algo = 'SHA-1', t = Date.now() } = opt || {};
    const counter = Math.floor(t / 1000 / period);
    const key = base32ToBytes(secretB32);
    if (!key.length) throw new Error('시크릿(Base32)이 비어있거나 형식이 올바르지 않습니다');
    const ck = await crypto.subtle.importKey('raw', key, { name: 'HMAC', hash: algo }, false, ['sign']);
    const buf = new ArrayBuffer(8);
    const dv = new DataView(buf);
    dv.setUint32(0, Math.floor(counter / 2 ** 32), false); // 상위 32비트
    dv.setUint32(4, counter >>> 0, false);                 // 하위 32비트 (빅엔디안)
    const h = new Uint8Array(await crypto.subtle.sign('HMAC', ck, buf));
    const off = h[h.length - 1] & 0xf;
    const bin = ((h[off] & 0x7f) << 24) | ((h[off + 1] & 0xff) << 16) | ((h[off + 2] & 0xff) << 8) | (h[off + 3] & 0xff);
    return (bin % 10 ** digits).toString().padStart(digits, '0');
  }

  // otpauth:// URL 또는 순수 Base32 문자열을 파싱해 설정으로 저장
  function parseAndStoreSecret(raw) {
    const input = String(raw).trim();
    if (!input) return false;
    let secret = input, digits = 6, period = 30, algo = 'SHA-1';
    if (/^otpauth:\/\//i.test(input)) {
      try {
        const u = new URL(input);
        const p = u.searchParams;
        secret = (p.get('secret') || '').trim();
        if (p.get('digits')) digits = parseInt(p.get('digits'), 10) || 6;
        if (p.get('period')) period = parseInt(p.get('period'), 10) || 30;
        if (p.get('algorithm')) {
          const a = p.get('algorithm').toUpperCase();
          algo = a === 'SHA256' ? 'SHA-256' : a === 'SHA512' ? 'SHA-512' : 'SHA-1';
        }
      } catch (_) { /* 아래에서 Base32로 취급 */ }
    }
    secret = secret.replace(/\s+/g, '').toUpperCase();
    if (!/^[A-Z2-7]+=*$/.test(secret) || base32ToBytes(secret).length < 10) return false;
    GM_setValue(KEY.secret, secret);
    GM_setValue(KEY.digits, digits);
    GM_setValue(KEY.period, period);
    GM_setValue(KEY.algo, algo);
    return true;
  }

  // ────────────────────────── OTP 입력칸 감지 ──────────────────────────
  const OTP_RE = /(otp|one[\s_-]?time|일회용|인증\s*번호|인증코드|auth[\s_-]?no|authno|otpno|otp[\s_-]?code|otp[\s_-]?num|security\s*code)/i;

  function isVisible(el) {
    return el && el.offsetParent !== null && !el.disabled && !el.readOnly;
  }

  function findOtpInput() {
    const c = cfg();
    // 1) 사용자가 명시한 선택자 최우선
    if (c.selector) {
      try {
        const el = document.querySelector(c.selector);
        if (el && isVisible(el)) return el;
      } catch (_) { /* 잘못된 선택자 무시 */ }
    }
    const inputs = [...document.querySelectorAll('input')]
      .filter(el => el.type !== 'hidden' && el.type !== 'password' && el.type !== 'submit' && el.type !== 'button' && isVisible(el));
    // 2) 속성(id/name/placeholder/label/class)에 OTP 키워드
    const scored = inputs.map(el => {
      const lbl = (el.labels && el.labels[0] && el.labels[0].textContent) || '';
      const hay = `${el.id} ${el.name} ${el.placeholder} ${el.className} ${el.getAttribute('title') || ''} ${el.getAttribute('aria-label') || ''} ${lbl}`;
      return { el, hit: OTP_RE.test(hay) };
    });
    const byKeyword = scored.find(s => s.hit);
    if (byKeyword) return byKeyword.el;
    // 3) maxlength 4~8 의 숫자형/텍스트 입력칸 (아이디·비번 칸 제외됨)
    const byLen = inputs.find(el => {
      const ml = parseInt(el.getAttribute('maxlength') || '0', 10);
      const numeric = el.inputMode === 'numeric' || el.type === 'number' || /^\d*$/.test(el.value);
      return ml >= 4 && ml <= 8 && numeric;
    });
    return byLen || null;
  }

  // React/Vue의 value 프로퍼티 setter를 우회해 값 주입 + 이벤트 발생
  function setVal(el, val) {
    const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
    setter.call(el, val);
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
  }

  // 자동제출: OTP 칸이 속한 폼의 로그인/확인 버튼 클릭 (기본 OFF)
  function trySubmit(otpEl) {
    const scope = otpEl.form || document;
    const btns = [...scope.querySelectorAll('button, input[type=submit], a')].filter(isVisible);
    const re = /(로그인|확인|인증|제출|login|sign\s*in|submit|ok)/i;
    const btn = btns.find(b => re.test((b.innerText || b.value || '').trim()));
    if (btn) { btn.click(); return true; }
    if (otpEl.form && typeof otpEl.form.requestSubmit === 'function') { otpEl.form.requestSubmit(); return true; }
    return false;
  }

  // ────────────────────────── 상태 & 자동입력 루프 ──────────────────────────
  let lastCode = '';
  let lastFilledCounter = -1;
  let lastFilledEl = null;

  async function tick() {
    const c = cfg();
    const codeEl = document.getElementById('__otp_code');
    const barEl = document.getElementById('__otp_bar');
    const statEl = document.getElementById('__otp_stat');

    if (!c.secret) {
      if (codeEl) codeEl.textContent = '시크릿 미등록';
      if (statEl) statEl.textContent = '메뉴 ▸ ① OTP 시크릿 등록';
      return;
    }

    let code;
    try {
      code = await totp(c.secret, c);
    } catch (e) {
      if (codeEl) codeEl.textContent = '오류';
      if (statEl) statEl.textContent = e.message;
      return;
    }

    const now = Date.now();
    const counter = Math.floor(now / 1000 / c.period);
    const remain = c.period - Math.floor(now / 1000) % c.period;
    lastCode = code;

    if (codeEl) codeEl.textContent = code.replace(/(\d{3})(\d+)/, '$1 $2');
    if (statEl) statEl.textContent = `남은 시간 ${remain}s · ${c.digits}자리/${c.period}s`;
    if (barEl) barEl.style.width = Math.round((remain / c.period) * 100) + '%';

    // OTP 칸 감지 → 새 코드 주기가 되었거나 아직 안 채운 칸이면 채움
    const el = findOtpInput();
    if (el) {
      const needFill = el !== lastFilledEl || lastFilledCounter !== counter || el.value !== code;
      if (needFill) {
        setVal(el, code);
        lastFilledEl = el;
        lastFilledCounter = counter;
        if (statEl) statEl.textContent = `입력 완료 · 남은 ${remain}s`;
        if (c.autosubmit) {
          // 코드가 유효한 동안만 제출 (경계 직전 제출 방지: 최소 2초 여유)
          if (remain >= 2) trySubmit(el);
        }
      }
    }
  }

  // ────────────────────────── UI 패널 ──────────────────────────
  function mountPanel() {
    if (document.getElementById('__otp_panel')) return;
    if (!document.body) return;
    const p = document.createElement('div');
    p.id = '__otp_panel';
    p.style.cssText = 'position:fixed;right:16px;bottom:16px;z-index:2147483647;width:210px;background:#fff;border:1px solid #d9dde3;border-radius:10px;box-shadow:0 6px 22px rgba(0,0,0,.16);font:13px/1.4 -apple-system,BlinkMacSystemFont,"Malgun Gothic",sans-serif;color:#222;padding:12px;';
    p.innerHTML = `
      <div id="__otp_head" style="font-weight:700;margin:-12px -12px 10px;padding:8px 12px;color:#1565c0;cursor:move;user-select:none;border-bottom:1px solid #eee;background:#f5f9ff;border-radius:10px 10px 0 0">관리자 OTP 자동입력 <span style="float:right;color:#bbb;font-weight:400">⠿</span></div>
      <div id="__otp_code" style="font-size:26px;font-weight:800;letter-spacing:2px;text-align:center;color:#0d47a1;font-variant-numeric:tabular-nums">──────</div>
      <div style="height:5px;background:#eceff3;border-radius:3px;margin:8px 0 6px;overflow:hidden"><div id="__otp_bar" style="height:100%;width:100%;background:#1e88e5;transition:width .9s linear"></div></div>
      <div id="__otp_stat" style="color:#888;font-size:11px;min-height:14px;text-align:center;margin-bottom:8px"></div>
      <div style="display:flex;gap:6px">
        <button id="__otp_fill" style="flex:1;background:#1e88e5;color:#fff;border:0;padding:8px;border-radius:6px;font-weight:600;cursor:pointer">지금 채우기</button>
        <button id="__otp_copy" style="width:64px;background:#eef2f7;color:#333;border:0;padding:8px;border-radius:6px;font-weight:600;cursor:pointer">복사</button>
      </div>`;
    document.body.appendChild(p);

    document.getElementById('__otp_fill').addEventListener('click', () => {
      const el = findOtpInput();
      if (!el) { flash('OTP 입력칸을 못 찾음 — 메뉴 ③에서 선택자를 지정하세요'); return; }
      if (!lastCode) { flash('코드 미생성 — 시크릿을 먼저 등록하세요'); return; }
      setVal(el, lastCode);
      lastFilledEl = el; lastFilledCounter = -2; // 강제 표시 갱신
      flash('입력 완료');
    });
    document.getElementById('__otp_copy').addEventListener('click', async () => {
      if (!lastCode) { flash('코드 없음'); return; }
      try { await navigator.clipboard.writeText(lastCode); flash('복사됨'); }
      catch (_) { flash('복사 실패(권한)'); }
    });
    makeDraggable(p, document.getElementById('__otp_head'));
  }

  function flash(msg) {
    const s = document.getElementById('__otp_stat');
    if (!s) return;
    const prev = s.textContent;
    s.textContent = msg;
    s.style.color = '#1565c0';
    setTimeout(() => { s.style.color = '#888'; }, 1600);
  }

  function makeDraggable(panel, handle) {
    let sx = 0, sy = 0, ox = 0, oy = 0;
    handle.addEventListener('mousedown', e => {
      e.preventDefault();
      const r = panel.getBoundingClientRect();
      panel.style.left = r.left + 'px'; panel.style.top = r.top + 'px';
      panel.style.right = 'auto'; panel.style.bottom = 'auto';
      sx = e.clientX; sy = e.clientY; ox = r.left; oy = r.top;
      document.addEventListener('mousemove', onMove);
      document.addEventListener('mouseup', onUp);
    });
    function onMove(e) {
      const maxX = window.innerWidth - panel.offsetWidth;
      const maxY = window.innerHeight - panel.offsetHeight;
      panel.style.left = Math.max(0, Math.min(maxX, ox + e.clientX - sx)) + 'px';
      panel.style.top = Math.max(0, Math.min(maxY, oy + e.clientY - sy)) + 'px';
    }
    function onUp() {
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    }
  }

  // ────────────────────────── Tampermonkey 메뉴 ──────────────────────────
  GM_registerMenuCommand('① OTP 시크릿 등록 / 변경', () => {
    const cur = cfg().secret ? '(현재 등록됨 — 새 값 입력 시 교체)' : '';
    const v = prompt(`OTP 시크릿(Base32) 또는 otpauth:// URL 을 붙여넣으세요 ${cur}\n\n예) JBSWY3DPEHPK3PXP\n예) otpauth://totp/…?secret=…`, '');
    if (v == null) return;
    if (parseAndStoreSecret(v)) { flash('시크릿 등록 완료'); tick(); }
    else alert('형식이 올바르지 않습니다. Base32 문자(A–Z,2–7) 또는 otpauth:// URL을 확인하세요.');
  });
  GM_registerMenuCommand('② 자동제출 켜기/끄기', () => {
    const next = !cfg().autosubmit;
    GM_setValue(KEY.autosubmit, next);
    alert('자동제출: ' + (next ? 'ON — 코드 입력 후 로그인/확인 버튼을 자동 클릭합니다' : 'OFF'));
  });
  GM_registerMenuCommand('③ OTP 입력칸 선택자 지정', () => {
    const v = prompt('OTP 입력칸의 CSS 선택자를 입력하세요 (자동감지 실패 시).\n예) #otpNo  ·  input[name="otpNumber"]\n\n비워두면 자동감지 사용.', cfg().selector);
    if (v == null) return;
    GM_setValue(KEY.selector, v.trim());
    flash(v.trim() ? '선택자 지정됨' : '자동감지로 복귀');
  });
  GM_registerMenuCommand('④ 저장된 시크릿 삭제', () => {
    if (confirm('저장된 OTP 시크릿을 삭제할까요?')) {
      GM_deleteValue(KEY.secret);
      alert('삭제되었습니다.');
    }
  });

  // ────────────────────────── 부팅 ──────────────────────────
  mountPanel();
  tick();
  // OTP 화면이 로그인 2단계로 뒤늦게 나타날 수 있으므로 주기 폴링 + 패널 유지
  setInterval(() => { mountPanel(); tick(); }, 1000);
})();
