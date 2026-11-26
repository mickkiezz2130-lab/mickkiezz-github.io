/**
 * ปาเป้าพิชิตเชื้อ 3D — เซิร์ฟเวอร์ลงทะเบียน/สถิติ/อันดับ (Google Apps Script + Google Sheet)
 * หน่วยป้องกันและควบคุมการติดเชื้อ โรงพยาบาลบุรีรัมย์
 *
 * วิธีติดตั้ง (ทำครั้งเดียว):
 * 1) สร้าง Google Sheet ใหม่ → เมนู ส่วนขยาย → Apps Script → ลบโค้ดเดิม วางไฟล์นี้ทั้งหมด → บันทึก
 * 2) เลือกฟังก์ชัน setup ด้านบน → กด เรียกใช้ (Run) → อนุญาตสิทธิ์
 * 3) ทำให้ใช้งานได้ (Deploy) → การทำให้ใช้งานได้รายการใหม่ → ประเภท: เว็บแอป
 *    - เรียกใช้ในฐานะ: ฉัน (Me)
 *    - ผู้มีสิทธิ์เข้าถึง: ทุกคน (Anyone)
 *    → ทำให้ใช้งานได้ → คัดลอก URL ที่ลงท้ายด้วย /exec
 * 4) นำ URL ไปใส่ที่ CONFIG.API_URL ในไฟล์ ic-darts/index.html (หรือในเกม: 📊 สถิติ → ⚙️ ตั้งค่า)
 *
 * แผ่นงานที่ได้:  Players (ผู้เล่น + สถิติสะสม) · Plays (ทุกครั้งที่เล่น) · สรุป (ตารางสรุปอัตโนมัติ)
 * ห้องแข่ง / คนออนไลน์ / คำเชิญ เก็บใน CacheService (ชั่วคราว ไม่ลงชีต)
 * ข้อจำกัด: Apps Script รับคำขอพร้อมกันได้จำกัด เหมาะกับการแข่งกลุ่มเล็ก (ห้องละไม่เกิน 8 คน)
 */
const SHEET_ID = '';            // ว่าง = ใช้ Google Sheet ที่ผูกกับสคริปต์นี้
const SALT = 'icdarts-brh-2569'; // ใช้เข้ารหัส PIN (เปลี่ยนได้ก่อนเริ่มใช้งานจริง)

const HEAD = {
  Players: ['id', 'ชื่อ', 'หน่วยงาน', 'pinHash', 'ลงทะเบียน', 'เล่นล่าสุด', 'จำนวนครั้ง', 'คะแนนรวม(สูงสุดทุกด่าน)', 'ดาว', 'ปาทั้งหมด', 'ปาโดน', 'ตอบถูก', 'ตอบผิด', 'เวลาเล่น(วิ)', 'เหรียญ', 'profile'],
  Plays: ['เวลา', 'id', 'ชื่อ', 'หน่วยงาน', 'level', 'ด่าน', 'ระดับ', 'คะแนน', 'ดาว', 'ตอบถูก', 'ตอบผิด', 'ปา', 'ปาโดน', 'เวลา(วิ)'],
};

function ss_() { return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(name) {
  const s = ss_(); let sh = s.getSheetByName(name);
  if (!sh) {
    sh = s.insertSheet(name); sh.appendRow(HEAD[name]); sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD[name].length).setFontWeight('bold').setBackground('#0e7490').setFontColor('#ffffff');
  }
  return sh;
}

/** เรียกครั้งแรกเพื่อสร้างแผ่นงาน */
function setup() {
  sheet_('Players'); sheet_('Plays');
  const s = ss_(); let sum = s.getSheetByName('สรุป');
  if (!sum) sum = s.insertSheet('สรุป');
  sum.clear();
  sum.getRange('A1').setValue('การเล่นแยกตามด่าน').setFontWeight('bold');
  sum.getRange('A2').setFormula('=IFERROR(QUERY(Plays!A:N,"select F, count(B), avg(H), avg(I) where B is not null group by F order by count(B) desc label F \'ด่าน\', count(B) \'ครั้งที่เล่น\', avg(H) \'คะแนนเฉลี่ย\', avg(I) \'ดาวเฉลี่ย\'",1),"ยังไม่มีข้อมูล")');
  sum.getRange('G1').setValue('การเล่นแยกตามหน่วยงาน').setFontWeight('bold');
  sum.getRange('G2').setFormula('=IFERROR(QUERY(Plays!A:N,"select D, count(B), sum(H), sum(J), sum(K) where B is not null group by D order by count(B) desc label D \'หน่วยงาน\', count(B) \'ครั้งที่เล่น\', sum(H) \'คะแนนรวม\', sum(J) \'ตอบถูก\', sum(K) \'ตอบผิด\'",1),"ยังไม่มีข้อมูล")');
  sum.getRange('A30').setValue('จำนวนผู้เล่นทั้งหมด').setFontWeight('bold');
  sum.getRange('B30').setFormula('=COUNTA(Players!A2:A)');
  sum.getRange('A31').setValue('จำนวนครั้งที่เล่นทั้งหมด').setFontWeight('bold');
  sum.getRange('B31').setFormula('=COUNTA(Plays!B2:B)');
}

function hash_(pin) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, SALT + ':' + pin, Utilities.Charset.UTF_8);
  return Utilities.base64Encode(raw);
}
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function clean_(s, n) { return String(s == null ? '' : s).replace(/[<>]/g, '').replace(/^[=+\-@]+/, '').trim().slice(0, n); }
function num_(x) { return Math.max(0, Math.min(1e7, Math.round(Number(x) || 0))); }
function same_(a, b) { return String(a).trim().toLowerCase() === String(b).trim().toLowerCase(); }

function doGet(e) {
  const a = (e && e.parameter && e.parameter.action) || 'ping';
  if (a === 'board') return out_(board_());
  return out_({ ok: true, app: 'icdarts', time: new Date() });
}

function doPost(e) {
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, err: 'bad', msg: 'ข้อมูลไม่ถูกต้อง' }); }
  if (d.action === 'ping') return out_({ ok: true });
  if (d.action === 'board') return out_(board_());
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    if (d.action === 'register') return out_(register_(d));
    if (d.action === 'login') return out_(login_(d));
    if (d.action === 'submit') return out_(submit_(d));
    if (MP[d.action]) { const me = auth_(d); if (!me) return out_({ ok: false, err: 'auth' }); return out_(MP[d.action](d, me)); }
    return out_({ ok: false, err: 'action' });
  } catch (x) {
    return out_({ ok: false, err: 'server', msg: String(x) });
  } finally {
    try { lock.releaseLock(); } catch (y) {}
  }
}

function rows_() { const sh = sheet_('Players'); const v = sh.getDataRange().getValues(); return { sh: sh, v: v }; }

function register_(d) {
  const name = clean_(d.name, 30), unit = clean_(d.unit, 40), pin = String(d.pin || '');
  if (name.length < 2 || !unit || !/^\d{4}$/.test(pin)) return { ok: false, err: 'bad', msg: 'กรอกข้อมูลไม่ครบ' };
  const P = rows_();
  if (P.v.some((r, i) => i > 0 && same_(r[1], name))) return { ok: false, err: 'exists' };
  const id = 'P' + Date.now().toString(36) + Math.floor(Math.random() * 1e4);
  P.sh.appendRow([id, name, unit, hash_(pin), new Date(), '', 0, 0, 0, 0, 0, 0, 0, 0, 0, '']);
  return { ok: true, id: id, unit: unit };
}

function login_(d) {
  const name = clean_(d.name, 30); const P = rows_();
  const r = P.v.find((x, i) => i > 0 && same_(x[1], name));
  if (!r) return { ok: false, err: 'notfound' };
  if (r[3] !== hash_(String(d.pin || ''))) return { ok: false, err: 'pin' };
  let profile = null; try { profile = r[15] ? JSON.parse(r[15]) : null; } catch (x) {}
  return { ok: true, id: r[0], unit: r[2], profile: profile };
}

function submit_(d) {
  const P = rows_();
  const i = P.v.findIndex((r, k) => k > 0 && String(r[0]) === String(d.id));
  if (i < 1) return { ok: false, err: 'auth' };
  const r = P.v[i];
  if (r[3] !== hash_(String(d.pin || ''))) return { ok: false, err: 'auth' };
  sheet_('Plays').appendRow([new Date(), r[0], r[1], r[2], clean_(d.level, 20), clean_(d.levelName, 80), clean_(d.diff, 10),
    num_(d.score), num_(d.stars), num_(d.correct), num_(d.wrong), num_(d.throws), num_(d.hits), num_(d.dur)]);
  const p = d.profile || {}; const st = p.st || {};
  const row = [r[0], r[1], r[2], r[3], r[4], new Date(), num_(st.plays || d.plays), num_(d.total), num_(d.totalStars),
    num_(st.throws), num_(st.hits), num_(st.correct), num_(st.wrong), num_(st.time), num_(p.coins), JSON.stringify(p).slice(0, 45000)];
  P.sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
  return { ok: true };
}

function board_() {
  const P = rows_();
  const players = P.v.slice(1).filter(r => r[0]).map(r => ({
    id: r[0], name: r[1], unit: r[2], plays: Number(r[6]) || 0, total: Number(r[7]) || 0, stars: Number(r[8]) || 0,
    acc: Number(r[9]) ? Math.round(Number(r[10]) / Number(r[9]) * 100) : 0,
  })).sort((a, b) => b.total - a.total || b.stars - a.stars);
  const um = {};
  players.forEach(x => { const u = um[x.unit] || (um[x.unit] = { unit: x.unit, players: 0, plays: 0, total: 0 }); u.players++; u.plays += x.plays; u.total += x.total; });
  const plays = sheet_('Plays'); const last = plays.getLastRow(); let todayN = 0;
  const tz = Session.getScriptTimeZone(); const td = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  if (last > 1) {
    const from = Math.max(2, last - 3000);
    todayN = plays.getRange(from, 1, last - from + 1, 1).getValues()
      .filter(t => t[0] instanceof Date && Utilities.formatDate(t[0], tz, 'yyyy-MM-dd') === td).length;
  }
  return { ok: true, players: players.slice(0, 200), units: Object.keys(um).map(k => um[k]).sort((a, b) => b.total - a.total),
    summary: { players: players.length, plays: Math.max(0, last - 1), today: todayN } };
}

/* ======================= ผู้เล่นออนไลน์ · คำเชิญ · ห้องแข่ง ======================= */
const TTL = 21600;
function cache_() { return CacheService.getScriptCache(); }
function cget_(k) { const v = cache_().get(k); return v ? JSON.parse(v) : null; }
function cput_(k, o, ttl) { cache_().put(k, JSON.stringify(o), ttl || TTL); }
function auth_(d) {
  const key = 'auth:' + d.id, h = hash_(String(d.pin || ''));
  const c = cget_(key); if (c && c.h === h) return c;
  const r = rows_().v.find((x, i) => i > 0 && String(x[0]) === String(d.id));
  if (!r || r[3] !== h) return null;
  const me = { id: r[0], name: r[1], unit: r[2], h: h }; cput_(key, me); return me;
}
function roomOut_(room) { if (room) room.now = Date.now(); return room; }
function newCode_() { const A = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'; let c; do { c = ''; for (let i = 0; i < 4; i++) c += A[Math.floor(Math.random() * A.length)]; } while (cget_('room:' + c)); return c; }
function player_(me) { return { n: me.name, u: me.unit, s: 0, done: false, t: Date.now() }; }
const MP = {
  hello(d, me) {
    const now = Date.now(); const pr = cget_('presence') || {};
    pr[me.id] = { n: me.name, u: me.unit, s: clean_(d.status, 10) || 'menu', t: now };
    Object.keys(pr).forEach(k => { if (now - pr[k].t > 90000) delete pr[k]; }); cput_('presence', pr, 600);
    const inv = cget_('inv:' + me.id) || []; if (inv.length) cache_().remove('inv:' + me.id);
    const players = Object.keys(pr).filter(k => k !== me.id).map(k => ({ id: k, name: pr[k].n, unit: pr[k].u, status: pr[k].s })).slice(0, 80);
    return { ok: true, online: Object.keys(pr).length, players: players, invites: inv.filter(x => now - x.t < 300000) };
  },
  invite(d, me) {
    const to = String(d.to || ''), code = clean_(d.code, 6); if (!to || !cget_('room:' + code)) return { ok: false, err: 'nf' };
    const k = 'inv:' + to; const inv = (cget_(k) || []).filter(x => x.code !== code);
    inv.push({ from: me.name, unit: me.unit, code: code, t: Date.now() }); cput_(k, inv.slice(-5), 600); return { ok: true };
  },
  room_create(d, me) {
    const code = newCode_();
    const room = { code: code, host: me.id, level: num_(d.level), diff: num_(d.diff), seed: Math.floor(Math.random() * 1e9), pub: !!d.pub, state: 'wait', created: Date.now(), startAt: 0, players: {} };
    room.players[me.id] = player_(me); cput_('room:' + code, room);
    if (room.pub) { const l = cget_('pubrooms') || []; l.push(code); cput_('pubrooms', l.slice(-30), 900); }
    return { ok: true, room: roomOut_(room) };
  },
  room_join(d, me) {
    const room = cget_('room:' + clean_(d.code, 6).toUpperCase()); if (!room) return { ok: false, err: 'nf' };
    if (!room.players[me.id]) {
      if (room.state !== 'wait') return { ok: false, err: 'started' };
      if (Object.keys(room.players).length >= 8) return { ok: false, err: 'full' };
      room.players[me.id] = player_(me); cput_('room:' + room.code, room);
    }
    return { ok: true, room: roomOut_(room) };
  },
  room_quick(d, me) {
    const now = Date.now(); const l = cget_('pubrooms') || []; const keep = [];
    for (const code of l) {
      const room = cget_('room:' + code);
      if (!room || room.state !== 'wait' || now - room.created > 180000) continue; keep.push(code);
      if (Object.keys(room.players).length < 8) {
        room.players[me.id] = room.players[me.id] || player_(me); cput_('room:' + code, room); cput_('pubrooms', keep.concat(l.slice(l.indexOf(code) + 1)), 900);
        return { ok: true, room: roomOut_(room) };
      }
    }
    cput_('pubrooms', keep, 900);
    return MP.room_create(Object.assign({}, d, { pub: true }), me);
  },
  room_get(d, me) {
    const room = cget_('room:' + clean_(d.code, 6)); if (!room) return { ok: false, err: 'nf' };
    if (room.players[me.id]) room.players[me.id].t = Date.now();
    if (room.state === 'wait') Object.keys(room.players).forEach(k => { if (k !== room.host && Date.now() - room.players[k].t > 30000) delete room.players[k]; });
    cput_('room:' + room.code, room); return { ok: true, room: roomOut_(room) };
  },
  room_start(d, me) {
    const room = cget_('room:' + clean_(d.code, 6)); if (!room) return { ok: false, err: 'nf' };
    if (room.host !== me.id) return { ok: false, err: 'host' };
    if (room.state === 'wait') { room.state = 'play'; room.startAt = Date.now() + 4000; cput_('room:' + room.code, room); }
    return { ok: true, room: roomOut_(room) };
  },
  room_score(d, me) {
    const room = cget_('room:' + clean_(d.code, 6)); if (!room || !room.players[me.id]) return { ok: false, err: 'nf' };
    const p = room.players[me.id]; p.s = num_(d.score); p.st = num_(d.stars); p.done = p.done || !!d.done; p.t = Date.now();
    const ids = Object.keys(room.players);
    if (ids.every(k => room.players[k].done || Date.now() - room.players[k].t > 45000) || Date.now() - room.startAt > 600000) room.state = 'done';
    cput_('room:' + room.code, room); return { ok: true, room: roomOut_(room) };
  },
  room_leave(d, me) {
    const room = cget_('room:' + clean_(d.code, 6)); if (!room) return { ok: true };
    if (room.state === 'wait') {
      delete room.players[me.id]; const ids = Object.keys(room.players);
      if (!ids.length) { cache_().remove('room:' + room.code); return { ok: true }; }
      if (room.host === me.id) room.host = ids[0];
    } else if (room.players[me.id]) room.players[me.id].done = true;
    cput_('room:' + room.code, room); return { ok: true };
  },
};
