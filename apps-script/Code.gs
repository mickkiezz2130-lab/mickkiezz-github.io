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
 *                 โจทย์_PPE และ โจทย์_คำถาม — แก้/เพิ่มโจทย์ได้เอง เกมดึงไปใช้ภายใน ~2 นาที
 *                 (ชีต PPE เพิ่มแถวใหม่: ใส่ชื่อโรค ติ๊กอุปกรณ์ เขียนคำอธิบาย · ช่อง ใช้งาน/อีโมจิ เว้นว่างได้)
 *                 (ติ๊ก "ใช้งาน" ออกเพื่อซ่อนข้อนั้น · ต้องมีอย่างน้อย 7 โรค และ 12 คำถาม ไม่งั้นเกมใช้ชุดเดิม)
 * ห้องแข่ง / คนออนไลน์ / คำเชิญ เก็บใน CacheService (ชั่วคราว ไม่ลงชีต)
 * ข้อจำกัด: Apps Script รับคำขอพร้อมกันได้จำกัด เหมาะกับการแข่งกลุ่มเล็ก (ห้องละไม่เกิน 8 คน)
 */
const SHEET_ID = '1LN8tjpUvWqoYjfX8zwnr4mEaU8SDLb4pUlNgMXXAi5c'; // Google Sheet "PPE" ของหน่วย IC (ว่าง = ใช้ชีตที่ผูกกับสคริปต์)
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
  sheet_('Players'); sheet_('Plays'); setupContent_();
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
  if (d.action === 'content') return out_(content_());
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
    const room = { code: code, host: me.id, level: num_(d.level), diff: num_(d.diff), seed: Math.floor(Math.random() * 1e9), sub: clean_(d.sub, 6), pub: !!d.pub, state: 'wait', created: Date.now(), startAt: 0, players: {} };
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

/* ======================= โจทย์ที่แก้ได้จากชีต ======================= */
const DEFAULT_CONTENT = {"quiz": [{"q": "ล้างมือด้วยแอลกอฮอล์เจล (Hand rub) ควรถูให้ทั่วนานเท่าใด", "a": "20–30 วินาที", "w": ["5 วินาที", "2 นาที", "5 นาที"], "x": "Alcohol hand rub ถูให้ทั่วจนแห้ง 20–30 วินาที"}, {"q": "ล้างมือด้วยน้ำและสบู่ ควรใช้เวลาประมาณเท่าใด", "a": "40–60 วินาที", "w": ["10 วินาที", "3 นาที", "15 วินาที"], "x": "ล้างมือด้วยน้ำและสบู่ 7 ขั้นตอน ใช้เวลา 40–60 วินาที"}, {"q": "กรณีใดควรล้างมือด้วยน้ำและสบู่ แทนแอลกอฮอล์เจล", "a": "มือเปื้อนสิ่งสกปรกที่มองเห็น หรือดูแลผู้ป่วย C. difficile", "w": ["ก่อนจับตัวผู้ป่วยทุกครั้ง", "หลังจับราวเตียง", "ก่อนเตรียมยารับประทาน"], "x": "สิ่งสกปรกที่มองเห็นและสปอร์ C. difficile ต้องใช้น้ำและสบู่"}, {"q": "แอลกอฮอล์ล้างมือควรมีความเข้มข้นเท่าใดจึงมีประสิทธิภาพ", "a": "60–80%", "w": ["10–20%", "30–40%", "95–100%"], "x": "แอลกอฮอล์ 60–80% ฆ่าเชื้อได้ดีที่สุด"}, {"q": "หลังสวมหน้ากาก N95 ทุกครั้ง ต้องทำสิ่งใด", "a": "ตรวจการรั่ว (Seal check)", "w": ["ล้างหน้ากากด้วยน้ำ", "พับหน้ากากครึ่งหนึ่ง", "สวมหน้ากากอนามัยทับ"], "x": "ทำ Seal check ทุกครั้งที่สวม และทำ Fit test ตามรอบที่กำหนด"}, {"q": "Standard Precautions ใช้กับผู้ป่วยกลุ่มใด", "a": "ผู้ป่วยทุกราย", "w": ["เฉพาะผู้ป่วย HIV", "เฉพาะผู้ป่วยในห้องแยกโรค", "เฉพาะผู้ป่วยที่มีไข้"], "x": "ถือว่าเลือดและสารคัดหลั่งของทุกคนอาจติดเชื้อได้"}, {"q": "ผู้ป่วยวัณโรคปอดระยะแพร่เชื้อ ควรอยู่ห้องแบบใด", "a": "ห้องแยกความดันลบ", "w": ["ห้องความดันบวก", "ห้องรวมใกล้ประตู", "ห้องพักญาติ"], "x": "Airborne Precautions: ห้องความดันลบ และสวม N95"}, {"q": "ผู้ป่วยไข้หวัดใหญ่ ควรเว้นระยะห่างจากผู้อื่นอย่างน้อยเท่าใด", "a": "1 เมตร", "w": ["10 เซนติเมตร", "30 เซนติเมตร", "ไม่ต้องเว้นระยะ"], "x": "Droplet Precautions: หน้ากากอนามัย และเว้นระยะอย่างน้อย 1 เมตร"}, {"q": "การถอด PPE ตามลำดับ (CDC) ควรถอดสิ่งใดก่อน", "a": "ถุงมือ", "w": ["หน้ากาก", "แว่น/Face shield", "เสื้อกาวน์ก่อนเสมอ"], "x": "ถุงมือปนเปื้อนมากที่สุด จึงถอดก่อน · หน้ากากถอดหลังออกจากห้อง"}, {"q": "หลังถอดถุงมือ ต้องทำอย่างไร", "a": "ล้างมือทุกครั้ง", "w": ["ไม่ต้องล้างมือ เพราะใส่ถุงมือแล้ว", "ล้างเฉพาะเมื่อถุงมือขาด", "เช็ดมือกับเสื้อกาวน์"], "x": "ถุงมืออาจมีรูรั่วและมือปนเปื้อนขณะถอด"}, {"q": "เชื้อใดสร้างสปอร์ที่ทนต่อแอลกอฮอล์", "a": "Clostridioides difficile", "w": ["Staphylococcus aureus", "Escherichia coli", "ไวรัสไข้หวัดใหญ่"], "x": "C. difficile สร้างสปอร์ ต้องล้างมือด้วยน้ำและสบู่"}, {"q": "น้ำยาเช็ดพื้นผิวในห้องผู้ป่วย C. difficile ที่เหมาะสม", "a": "สารประกอบคลอรีน (Hypochlorite)", "w": ["แอลกอฮอล์ 70%", "น้ำเปล่า", "น้ำหอมปรับอากาศ"], "x": "สปอร์ต้องใช้สารที่ฆ่าสปอร์ได้ เช่น Hypochlorite"}, {"q": "ภาชนะทิ้งของมีคม ควรปิดและเปลี่ยนเมื่อบรรจุถึงระดับใด", "a": "3 ใน 4 ของภาชนะ", "w": ["เต็มจนล้น", "1 ใน 4 ของภาชนะ", "เมื่อครบ 1 เดือนเท่านั้น"], "x": "บรรจุไม่เกิน 3/4 เพื่อป้องกันเข็มทิ่มตำ"}, {"q": "เมื่อถูกเข็มที่ใช้กับผู้ป่วยตำ ควรทำสิ่งใดเป็นอันดับแรก", "a": "ล้างแผลด้วยน้ำสะอาดและสบู่ทันที แล้วรายงาน", "w": ["บีบเค้นเลือดออกแรงๆ", "ใช้ปากดูดแผล", "รอเลิกเวรค่อยรายงาน"], "x": "ล้างแผลทันที ไม่บีบเค้น และรายงานเพื่อประเมินความเสี่ยงและรับยา PEP"}, {"q": "ยาป้องกันหลังสัมผัสเชื้อ HIV (PEP) ควรเริ่มเมื่อใด", "a": "เร็วที่สุด ภายใน 72 ชั่วโมง", "w": ["หลังสัมผัส 1 สัปดาห์", "เมื่อมีอาการไข้", "หลังผลเลือดเป็นบวก"], "x": "ยิ่งเร็วยิ่งดี ควรภายใน 1–2 ชม. และไม่เกิน 72 ชม."}, {"q": "ข้อใดเป็นการปฏิบัติที่ \"ไม่ควรทำ\" หลังใช้เข็มฉีดยา", "a": "สวมปลอกเข็มกลับด้วยสองมือ", "w": ["ทิ้งลงภาชนะของมีคมทันที", "ใช้เทคนิคมือเดียวหากจำเป็นต้องสวมปลอก", "ทิ้งทั้งกระบอกโดยไม่ถอดเข็ม"], "x": "การสวมปลอกเข็มสองมือเสี่ยงต่อเข็มตำมือ"}, {"q": "ถุงเก็บปัสสาวะของผู้ป่วยคาสายสวน ควรอยู่ตำแหน่งใด", "a": "ต่ำกว่ากระเพาะปัสสาวะ และไม่วางกับพื้น", "w": ["สูงกว่าระดับเตียง", "วางบนพื้นข้างเตียง", "วางบนตัวผู้ป่วย"], "x": "CAUTI Bundle: ระบบปิด ไหลลงตามแรงโน้มถ่วง ไม่ให้ถุงแตะพื้น"}, {"q": "ควรประเมินความจำเป็นในการคาสายสวนปัสสาวะบ่อยเพียงใด", "a": "ทุกวัน", "w": ["เดือนละครั้ง", "เมื่อสายอุดตันเท่านั้น", "ไม่ต้องประเมิน"], "x": "ถอดสายเร็วที่สุดเมื่อหมดข้อบ่งชี้ ลดความเสี่ยง CAUTI"}, {"q": "การป้องกันปอดอักเสบจากเครื่องช่วยหายใจ (VAP) ควรจัดท่าศีรษะสูงเท่าใด", "a": "30–45 องศา", "w": ["0 องศา (นอนราบ)", "10 องศา", "90 องศาตลอดเวลา"], "x": "VAP Bundle: ศีรษะสูง 30–45 องศา ถ้าไม่มีข้อห้าม"}, {"q": "น้ำยาทำความสะอาดผิวหนังก่อนใส่สายสวนหลอดเลือดดำส่วนกลาง (central line) ที่แนะนำ", "a": "Chlorhexidine > 0.5% ในแอลกอฮอล์", "w": ["น้ำเกลือ", "สบู่ทั่วไป", "น้ำยาล้างจาน"], "x": "CLABSI Bundle: Chlorhexidine-alcohol และรอให้แห้งก่อนแทงเข็ม"}, {"q": "ก่อนผ่าตัด หากจำเป็นต้องกำจัดขน ควรใช้วิธีใด", "a": "ใช้เครื่องตัดขน (Clipper)", "w": ["โกนด้วยมีดโกนคืนก่อนผ่าตัด", "ถอนขนด้วยแหนบ", "ใช้ครีมกำจัดขนบนแผล"], "x": "การโกนทำให้เกิดแผลถลอก เพิ่มความเสี่ยง SSI"}, {"q": "เครื่องมือที่สัมผัสเนื้อเยื่อปราศจากเชื้อ (Critical items) ต้องผ่านกระบวนการใด", "a": "ทำให้ปราศจากเชื้อ (Sterilization)", "w": ["เช็ดด้วยแอลกอฮอล์", "ล้างน้ำเปล่า", "ตากแดด"], "x": "Spaulding: Critical → Sterilization"}, {"q": "หูฟังแพทย์ (Stethoscope) จัดเป็นอุปกรณ์กลุ่มใด", "a": "Non-critical: ทำความสะอาด/ฆ่าเชื้อระดับต่ำ", "w": ["Critical: ต้องนึ่งฆ่าเชื้อ", "Semi-critical: ฆ่าเชื้อระดับสูง", "ใช้ครั้งเดียวทิ้งเสมอ"], "x": "สัมผัสผิวหนังที่ไม่มีแผล → เช็ดด้วยแอลกอฮอล์ 70% ระหว่างผู้ป่วย"}, {"q": "ขยะติดเชื้อ ต้องทิ้งในถุงสีใด", "a": "สีแดง", "w": ["สีดำ", "สีเหลือง", "ถุงใสทั่วไป"], "x": "ขยะติดเชื้อใช้ถุงแดง มีสัญลักษณ์ชีวอันตราย"}, {"q": "วัคซีนใดที่บุคลากรทางการแพทย์ควรได้รับทุกปี", "a": "ไข้หวัดใหญ่", "w": ["บาดทะยัก", "อีสุกอีใส", "ไวรัสตับอักเสบบี"], "x": "ไข้หวัดใหญ่เปลี่ยนสายพันธุ์ ต้องฉีดทุกปี"}, {"q": "ห้องแยกโรคความดันลบ ควรเปลี่ยนอากาศอย่างน้อยกี่รอบต่อชั่วโมง (ACH)", "a": "12 รอบ", "w": ["1 รอบ", "3 รอบ", "100 รอบ"], "x": "แนะนำ ≥ 12 ACH และระบายอากาศออกภายนอกหรือผ่าน HEPA"}, {"q": "ผู้ป่วยติดเชื้อ MRSA ควรได้รับการดูแลแบบใด", "a": "Contact Precautions", "w": ["Airborne Precautions", "ไม่ต้องแยก", "Droplet Precautions เท่านั้น"], "x": "MRSA แพร่ทางสัมผัส: ถุงมือ เสื้อกาวน์ อุปกรณ์เฉพาะราย"}, {"q": "Moment ที่ 2 ของการล้างมือ (5 Moments) คือข้อใด", "a": "ก่อนทำหัตถการสะอาด/ปราศจากเชื้อ", "w": ["หลังสัมผัสผู้ป่วย", "ก่อนเข้าห้องน้ำ", "หลังสัมผัสสิ่งแวดล้อมรอบผู้ป่วย"], "x": "5 Moments: 1 ก่อนสัมผัส · 2 ก่อนหัตถการปราศจากเชื้อ · 3 หลังสารคัดหลั่ง · 4 หลังสัมผัส · 5 หลังสิ่งแวดล้อม"}, {"q": "การล้างมือ 7 ขั้นตอน ขั้นตอนสุดท้ายคือข้อใด", "a": "ถูรอบข้อมือ", "w": ["ฝ่ามือถูฝ่ามือ", "ถูนิ้วหัวแม่มือ", "ปลายนิ้วถูฝ่ามือ"], "x": "ขั้นที่ 7: ถูรอบข้อมือทั้งสองข้าง"}, {"q": "ผ้าปูที่นอนเปื้อนเลือดผู้ป่วย ควรจัดการอย่างไร", "a": "ม้วนเก็บด้านเปื้อนไว้ด้านใน ใส่ถุงผ้าติดเชื้อ ไม่สะบัด", "w": ["สะบัดผ้าก่อนพับ", "วางบนพื้นห้องผู้ป่วย", "ล้างเองในอ่างล้างมือ"], "x": "ลดการฟุ้งกระจายของเชื้อ และส่งซักตามระบบ"}], "ppe": [{"e": "🫁", "n": "วัณโรคปอดระยะแพร่เชื้อ", "ppe": ["n95"], "x": "Airborne: หน้ากาก N95 (Fit test/Seal check) + ห้องแยกความดันลบ"}, {"e": "🔴", "n": "โรคหัด (Measles)", "ppe": ["n95"], "x": "Airborne: หน้ากาก N95 + ห้องแยกความดันลบ"}, {"e": "🤒", "n": "โรคสุกใส (Varicella)", "ppe": ["n95", "glove", "gown"], "x": "Airborne + Contact: N95 ร่วมกับถุงมือและเสื้อกาวน์"}, {"e": "🤧", "n": "ไข้หวัดใหญ่ (Influenza)", "ppe": ["smask"], "x": "Droplet: หน้ากากอนามัย เว้นระยะ ≥ 1 เมตร"}, {"e": "👶", "n": "โรคไอกรน (Pertussis)", "ppe": ["smask"], "x": "Droplet: หน้ากากอนามัย"}, {"e": "🧠", "n": "ไข้กาฬหลังแอ่น (Meningococcal)", "ppe": ["smask"], "x": "Droplet: หน้ากากอนามัย จนได้ยาปฏิชีวนะครบ 24 ชม."}, {"e": "🦠", "n": "แผลติดเชื้อ MRSA", "ppe": ["glove", "gown"], "x": "Contact: ถุงมือ + เสื้อกาวน์ ใช้อุปกรณ์เฉพาะราย"}, {"e": "💩", "n": "ท้องเสียจาก C. difficile", "ppe": ["glove", "gown"], "x": "Contact: ถุงมือ + เสื้อกาวน์ และล้างมือด้วยสบู่และน้ำ"}, {"e": "🧫", "n": "เชื้อดื้อยา CRE", "ppe": ["glove", "gown"], "x": "Contact: ถุงมือ + เสื้อกาวน์ แยกผู้ป่วย/รวมกลุ่ม"}, {"e": "🐛", "n": "โรคหิด (Scabies)", "ppe": ["glove", "gown"], "x": "Contact: ถุงมือ + เสื้อกาวน์"}, {"e": "🦠", "n": "ผู้ป่วยยืนยัน COVID-19", "ppe": ["n95", "glove", "gown", "eye"], "x": "N95 + ถุงมือ + เสื้อกาวน์ + แว่น/Face shield"}, {"e": "🐔", "n": "ไข้หวัดนก H5N1", "ppe": ["n95", "glove", "gown", "eye"], "x": "Airborne + Contact + ป้องกันตา: ครบชุด"}, {"e": "💨", "n": "ดูดเสมหะผู้ป่วยวัณโรค (หัตถการก่อละออง)", "ppe": ["n95", "glove", "gown", "eye"], "x": "หัตถการก่อละอองฝอย: N95 + ถุงมือ + เสื้อกาวน์ + แว่น/Face shield"}]};
const PPE_HEAD = ['ลำดับ', 'โรค / สถานการณ์', 'การแพร่เชื้อ', 'ถุงมือ', 'เสื้อกาวน์', 'หน้ากากอนามัย', 'หน้ากาก N95', 'แว่น/Face shield', 'จำนวนชิ้น', 'คำอธิบาย', 'ใช้งาน', 'อีโมจิ'];
const PPE_COLS = { glove: 'ถุงมือ', gown: 'เสื้อกาวน์', smask: 'หน้ากากอนามัย', n95: 'หน้ากาก N95', eye: 'แว่น/Face shield' };
const PPE_KEYS = ['glove', 'gown', 'smask', 'n95', 'eye'];
const QUIZ_HEAD = ['ใช้งาน', 'คำถาม', 'คำตอบที่ถูก', 'ตัวเลือกผิด 1', 'ตัวเลือกผิด 2', 'ตัวเลือกผิด 3', 'คำอธิบาย (แสดงหลังตอบ)'];
function setupContent_() {
  const s = ss_();
  if (!s.getSheetByName('โจทย์_PPE')) {
    const sh = s.insertSheet('โจทย์_PPE');
    const rows = DEFAULT_CONTENT.ppe.map((r, i) => [i + 1, r.n, ''].concat(PPE_KEYS.map(k => r.ppe.indexOf(k) >= 0)).concat(['=COUNTIF(D' + (i + 2) + ':H' + (i + 2) + ',TRUE)', r.x, true, r.e]));
    sh.getRange(1, 1, 1, PPE_HEAD.length).setValues([PPE_HEAD]).setFontWeight('bold').setBackground('#7c3aed').setFontColor('#ffffff');
    sh.getRange(2, 1, rows.length, PPE_HEAD.length).setValues(rows);
    const cb = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sh.getRange('K2:K200').setDataValidation(cb); sh.getRange('D2:H200').setDataValidation(cb);
    sh.setFrozenRows(1); sh.setColumnWidth(2, 260); sh.setColumnWidth(10, 380);
  }
  if (!s.getSheetByName('โจทย์_คำถาม')) {
    const sh = s.insertSheet('โจทย์_คำถาม');
    const rows = DEFAULT_CONTENT.quiz.map(r => [true, r.q, r.a, r.w[0], r.w[1], r.w[2], r.x]);
    sh.getRange(1, 1, 1, QUIZ_HEAD.length).setValues([QUIZ_HEAD]).setFontWeight('bold').setBackground('#be185d').setFontColor('#ffffff');
    sh.getRange(2, 1, rows.length, QUIZ_HEAD.length).setValues(rows);
    sh.getRange('A2:A500').setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    sh.setFrozenRows(1); sh.setColumnWidth(2, 380); [3, 4, 5, 6].forEach(c => sh.setColumnWidth(c, 220)); sh.setColumnWidth(7, 380);
  }
}
function content_() {
  const c = cget_('content'); if (c) return c;
  const s = ss_(); const out = { ok: true, ppe: [], quiz: [] };
  const sp = s.getSheetByName('โจทย์_PPE');
  if (sp && sp.getLastRow() > 1) {
    // อ่านตามชื่อหัวคอลัมน์ ย้าย/เพิ่มคอลัมน์ได้ ขอให้ชื่อหัวตรงตามเดิม
    const v = sp.getRange(1, 1, sp.getLastRow(), sp.getLastColumn()).getValues(); const h = v[0].map(x => String(x).trim());
    const col = n => h.indexOf(n); const yes = x => x === true || String(x).toUpperCase() === 'TRUE';
    const iN = col('โรค / สถานการณ์'), iX = col('คำอธิบาย'), iOn = col('ใช้งาน'), iE = col('อีโมจิ'), iT = col('การแพร่เชื้อ');
    v.slice(1).forEach(r => {
      const ppe = PPE_KEYS.filter(k => col(PPE_COLS[k]) >= 0 && yes(r[col(PPE_COLS[k])]));
      const on = iOn < 0 || r[iOn] === '' || yes(r[iOn]);
      const tr = iT >= 0 ? String(r[iT]) : '';
      const e = (iE >= 0 && String(r[iE]).trim()) || (/airborne/i.test(tr) ? '🫁' : /droplet/i.test(tr) ? '🤧' : '🦠');
      if (on && iN >= 0 && String(r[iN]).trim() && ppe.length) out.ppe.push({ e: String(e).slice(0, 8), n: clean_(r[iN], 80), ppe: ppe, x: clean_(iX >= 0 ? r[iX] : '', 200) || clean_(tr, 80) });
    });
  }
  const sq = s.getSheetByName('โจทย์_คำถาม');
  if (sq && sq.getLastRow() > 1) sq.getRange(2, 1, sq.getLastRow() - 1, QUIZ_HEAD.length).getValues().forEach(r => {
    const w = [r[3], r[4], r[5]].map(x => clean_(x, 120)).filter(Boolean);
    if ((r[0] === true || String(r[0]).toUpperCase() === 'TRUE' || r[0] === '') && String(r[1]).trim() && String(r[2]).trim() && w.length === 3) out.quiz.push({ q: clean_(r[1], 200), a: clean_(r[2], 120), w: w, x: clean_(r[6], 200) });
  });
  cput_('content', out, 120); return out;
}
