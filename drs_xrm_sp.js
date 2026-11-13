// drs_xrm_sp.js  ตัวแทน Xrm.WebApi ตัวที่สอง: หน้า DR เรียกเหมือน Dataverse แต่ข้อมูลอยู่ใน SharePoint list (ขั้นที่ 1)
// ไฟล์ใหม่ ไม่แตะ DR เดิม  ใช้ได้ทั้ง node (require ไว้เทส ส่ง fetch ปลอมเข้า) และเบราว์เซอร์ (script ธรรมดาหลัง drs_schema.js)
// ในเบราว์เซอร์ ถ้า <html data-drs-site="<ที่อยู่ไซต์>"> ตัวนี้จะตั้งเองทันที: window.DRS_XRM , ตัวกระจายตัวจัดการ on* , ตัวดัก fetch /api/data
//
// รองรับ (ขั้นที่ 1 รหัสตามรายงาน feasibility ข้อ 2):
//   S01 list ใบงาน ($select $orderby หน้า 5000 nextLink , createdby -> Author) , S02 S03 สมาชิกทีม , S06 ตั้งค่า (กรองคีย์ เรียง Modified top 1)
//   S12 อ่านใบเดียวพร้อม @odata.etag (id ตัวเลข หรือ GUID ที่หน้าจอสร้าง -> dr_clientkey) , S15 อีเมลผู้ใช้ปัจจุบัน
//   S18 สร้างใบ (GUID -> dr_clientkey ห้ามซ้ำ , เลขตั๋ว DR-<Id 4 หลัก>) , S19 PATCH ใบด้วย If-Match ผ่านตัวดัก fetch (412 คืนตรง)
//   S22 ถึง S24 สมาชิกทีม , S27 ตั้งค่า , S30 getGlobalContext , S08 ถึง S10 สังเคราะห์ role จากกลุ่ม SharePoint ของผู้ใช้
// ขั้น 2 (11/11/2026) ผู้รับผิดชอบหลายคน: S01 S13 $expand dr_DataRequest_Assignees อ่านจาก dr_assignees ของแถวงานที่เชื่อ ,
//   S28 S29 online.execute Associate Disassociate (ทีมเท่านั้น) MERGE แถวงานด้วย IF-MATCH ชน 412 ลองใหม่ ชื่ออ่านจาก DRS Team Members
// ขั้น 3 (11/11/2026) ไฟล์แนบ (annotation) = ไฟล์แนบของแถว list: S07 รายการ , S14 อ่าน base64 , S20 สร้าง , S21 ลบ
//   ผู้แจ้งแนบที่แถวใบงาน (DRS Requests) ทีมแนบที่แถวงาน (DRS Work) ไฟล์ส่งมอบชื่อขึ้นต้น drs-delivery-- annotationid = UniqueId ของไฟล์
//   ผลวัดก่อนเลือก .claude/evidence/2026-11-11-dr-sp-stage23/probe_att_run1.log
// ยังไม่รองรับ: throw ข้อความขึ้นต้น [DRS ยังไม่รองรับ] พร้อม entity กับ query และจดใน xrm.__drs.unsupported
//   ยกเว้นอ่านคอลัมน์ของขั้นถัดไป (dr_internalnote ขั้น 4) ได้ค่าว่างพร้อม console.warn ขึ้นต้น [DRS ยังไม่รองรับ] (throw แล้วหน้าโหลดไม่ได้ทั้งหน้า)
(function (root) {
  "use strict";
  const MM = "application/json;odata=minimalmetadata";
  const TAG = "[DRS ยังไม่รองรับ]";
  const GUID = /^\{?[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\}?$/i;

  function schemaOf() {
    if (typeof module === "object" && module.exports) return require("./drs_schema.js");
    return root.DRS && root.DRS.schema;
  }

  // ---------- แยก query แบบ OData ของ Dataverse (ระดับบนสุด ไม่ตัดใน วงเล็บ) ----------
  function splitTop(s, sep) {
    const out = []; let depth = 0, cur = "", q = false;
    for (const ch of s) {
      if (ch === "'") q = !q;
      if (!q && ch === "(") depth++;
      if (!q && ch === ")") depth--;
      if (!q && depth === 0 && ch === sep) { out.push(cur); cur = ""; continue; }
      cur += ch;
    }
    if (cur !== "") out.push(cur);
    return out;
  }
  function parseQuery(q) {
    const s = String(q || "").replace(/^\?/, "");
    const o = {};
    for (const part of splitTop(s, "&")) {
      if (!part) continue;
      const i = part.indexOf("=");
      const k = i < 0 ? part : part.slice(0, i), v = i < 0 ? "" : part.slice(i + 1);
      o[k] = v;
    }
    return o;
  }
  // "a,b($select=c,d),e" -> [{name:"a"},{name:"b",select:["c","d"]},...]
  function parseExpand(v) {
    return splitTop(v, ",").map((x) => {
      const m = x.trim().match(/^([A-Za-z0-9_]+)(?:\((.*)\))?$/);
      if (!m) return { name: x.trim(), bad: true };
      const inner = m[2] ? parseQuery(m[2].replace(/;/g, "&")) : {};
      return { name: m[1], select: inner.$select ? inner.$select.split(",").map((c) => c.trim()) : [], rest: Object.keys(inner).filter((k) => k !== "$select") };
    });
  }

  // ---------- ใบงานรวมสองแหล่ง: DRS Requests (ของผู้แจ้ง) + DRS Work (ของทีม) hardening 11/11/2026 ----------
  // กติกาทั้งหมดอยู่ตรงนี้ที่เดียว ทุกบทบาทอ่านด้วยกติกาเดียวกัน ค่าที่ผู้แจ้งปลอมได้มีผลได้ไม่เกินที่จอให้ผู้แจ้งกดเอง
  const verOf = (w) => { const v = parseFloat(String((w && w.OData__UIVersionString) || "0")); return Number.isFinite(v) ? v : 0; };
  const tms = (s) => { const t = Date.parse(s || ""); return Number.isFinite(t) ? t : null; };
  const byId = (a, b) => a.Id - b.Id;
  const rowsOf = (r, ws) => (ws || []).filter((w) => Number(w.dr_requestid) === Number(r.Id));
  // แถวที่เชื่อ: เวอร์ชันเกิน 1.0 (ผู้แจ้งสร้างได้แต่แก้ไม่ได้ เวอร์ชันจึงค้าง 1.0 จนทีมแตะ assume_run1.log A1 A2 A4 A8)
  // หลายแถว: ผู้สร้างแถว = ผู้สร้างใบ ก่อน (แถวนี้ผู้แจ้งอ่านได้) แล้ว Id น้อยสุด
  function pickRead(r, ws) {
    const t = rowsOf(r, ws).filter((w) => verOf(w) > 1).sort(byId);
    return t.find((w) => Number(w.AuthorId) === Number(r.AuthorId)) || t[0] || null;
  }
  // แถวที่ทีมจะเขียน: แถวที่ผู้สร้างใบสร้าง (เวอร์ชันใดก็ได้ ครั้งแรกเขียนใหม่ทั้งแถว) ก่อน แล้วแถวที่เชื่อ ไม่มีเลยค่อยสร้าง
  // แถวที่คนอื่นสร้างให้ใบที่ไม่ใช่ของตัวเองจึงไม่ถูกเลือก ไม่ถูกเชื่อตลอดไป และไม่มีข้อมูลทีมรั่วเข้าไป
  function pickTarget(r, ws) {
    const all = rowsOf(r, ws).sort(byId);
    return all.find((w) => Number(w.AuthorId) === Number(r.AuthorId)) || all.find((w) => verOf(w) > 1) || null;
  }
  function parseAct(raw) {
    return String(raw || "").split("\n").map((l) => l.trim()).filter(Boolean)
      .map((l) => { try { const o = JSON.parse(l); return o && typeof o === "object" && !Array.isArray(o) ? o : null; } catch (e) { return null; } }).filter(Boolean);
  }
  const canon = (o) => JSON.stringify(Object.keys(o).sort().map((k) => [k, o[k]]));
  // ---------- แชทสองแหล่ง (พี่ปริม T7 T7d 11/11/2026) ----------
  // อักขระกำหนดทิศ (T6) ทำให้ชื่อไฟล์โชว์กลับทิศ เช่น รายงาน U+202E gpj.exe ตัดทิ้งก่อนโชว์
  const BIDI = /[\u202A-\u202E\u2066-\u2069]/g;
  const unbidi = (s) => String(s).replace(BIDI, "");
  function parseChat(raw) {
    return String(raw || "").split("\n").map((l) => l.trim()).filter(Boolean)
      .map((l) => { try { const o = JSON.parse(l); return o && typeof o === "object" && !Array.isArray(o) ? o : null; } catch (e) { return null; } }).filter(Boolean);
  }
  const joinChat = (arr) => (arr.length ? arr.slice().sort((a, b) => String(a.t || "").localeCompare(String(b.t || ""))).map((o) => JSON.stringify(o)).join("\n") : null);
  function cleanMsg(o) {
    const x = Object.assign({}, o);
    if (typeof x.filename === "string") x.filename = unbidi(x.filename);
    if (typeof x.msg === "string") x.msg = unbidi(x.msg);
    return x;
  }
  function previewOf(m) {
    if (m.kind === "file") return "📎 " + String(m.filename || "");
    if (m.kind === "link") return "🔗 " + String(m.url || "");
    return String(m.msg || "").slice(0, 120);
  }
  // ข้อความทีม = dr_teamconv ของแถวงานที่เชื่อ (ทีมเขียนฝ่ายเดียว) , ข้อความใน Requests: บรรทัด by team ไม่โชว์ (ปลอม หรือของเดิมที่ย้ายแล้ว)
  // t ซ้ำข้อความทีมไม่โชว์ , ที่เหลือบังคับ by requester ชื่อผู้สร้างใบ ไม่มี delivery , replyTo ชี้ข้อความจริงเท่านั้น
  // ชื่อฝั่งผู้แจ้ง = อีเมลผู้สร้างใบ (Author.EMail ปลอมไม่ได้) เพราะหน้าเขียนและเทียบ msgIsMine ด้วย CURRENT_USER ซึ่งเป็นอีเมล
  // บังคับเป็น Author.Title แล้วผู้แจ้งลบข้อความตัวเองไม่ได้ ข้อความตัวเองขึ้นยังไม่อ่าน (probe_mine_before.log 11/11/2026)
  function mergeChat(wRaw, rRaw, author) {
    const team = parseChat(wRaw).map((o) => cleanMsg(Object.assign({}, o, { by: "team" })));
    const tt = new Set(team.map((o) => String(o.t)));
    const reqs = parseChat(rRaw).filter((o) => o.by !== "team" && !tt.has(String(o.t)))
      .map((o) => { const x = cleanMsg(Object.assign({}, o, { by: "requester", name: author.email || author.name })); delete x.delivery; return x; });
    const all = team.concat(reqs), byT = new Map(all.map((o) => [String(o.t), o]));
    for (const x of reqs) {
      if (!x.replyTo || typeof x.replyTo !== "object") { delete x.replyTo; continue; }
      const ref = x.replyTo.t !== undefined ? byT.get(String(x.replyTo.t)) : null;
      if (ref && ref !== x) x.replyTo = ref.deleted ? { t: ref.t, name: ref.name || "", msg: "", deleted: true } : { t: ref.t, name: ref.name || "", msg: previewOf(ref) };
      else if (x.replyTo.deleted) x.replyTo = Object.assign({ name: "", msg: "", deleted: true }, x.replyTo.t ? { t: x.replyTo.t } : {});
      else delete x.replyTo;
    }
    return joinChat(all);
  }
  // หน้าส่งแชททั้งก้อน แยกเป็นบรรทัดทีมกับบรรทัดอื่น
  function splitChat(raw) {
    const all = parseChat(raw);
    return { team: all.filter((o) => o.by === "team"), other: all.filter((o) => o.by !== "team") };
  }
  // บรรทัดดิบ by team ที่มีอยู่แล้วใน Requests เก็บไว้ตามเดิมทุกครั้งที่เขียน (ADDITIVE ไม่โชว์ ไม่ลบ)
  const legacyTeam = (raw) => parseChat(raw).filter((o) => o.by === "team");
  function reattr(o, a) {
    const x = Object.assign({}, o, { by: a.name });
    if (a.email) x.be = a.email; else delete x.be;
    return x;
  }
  function mergeActivity(wRaw, rRaw, author) {
    const out = [], seen = new Set();
    const add = (o) => { const k = canon(o); if (!seen.has(k)) { seen.add(k); out.push(o); } };
    parseAct(wRaw).forEach(add);
    parseAct(rRaw).forEach((o) => add(reattr(o, author)));
    out.sort((a, b) => String(a.t || "").localeCompare(String(b.t || "")));
    return out.length ? out.map((o) => JSON.stringify(o)).join("\n") : null;
  }
  // ผู้รับผิดชอบในแถวงาน: JSON รายการ {id,name} ค่าที่ไม่ใช่รายการหรือแถวไม่ครบถูกทิ้ง
  function parseAssignees(raw) {
    let a;
    try { a = JSON.parse(String(raw || "[]")); } catch (e) { return []; }
    if (!Array.isArray(a)) return [];
    const seen = new Set();
    return a.filter((x) => x && typeof x === "object" && x.id !== undefined && x.id !== null && String(x.id) !== "" && !seen.has(String(x.id)) && seen.add(String(x.id)))
      .map((x) => ({ id: String(x.id), name: String(x.name || "") }));
  }
  // ---------- บันทึกภายในทีม (ขั้น 4) list DRS Internal Notes ทีม แอดมิน เท่านั้น ----------
  // ตัดบรรทัดซ้ำ (ข้อความดิบตรงกัน) เก็บตัวแรก: หน้าลองใหม่ (copyMessageToInternalNote) ส่งโน้ตเดิมซ้ำมาได้
  function uniqLines(raw) {
    const seen = new Set();
    return String(raw || "").split("\n").map((l) => l.trim()).filter((l) => l && !seen.has(l) && seen.add(l));
  }
  // อ่าน: รวมทุกแถวของใบ (แถวซ้ำเกิดได้ถ้าทีมสองคนสร้างพร้อมกัน) ตัดซ้ำ เรียงตามเวลา t (เรียงคงลำดับเดิมเมื่อเวลาเท่ากัน)
  function notesOf(rows, id) {
    const mine = (rows || []).filter((x) => Number(x.dr_requestid) === Number(id)).sort(byId);
    const all = uniqLines(mine.map((x) => x.dr_internalnote || "").join("\n"));
    const tOf = (l) => { try { const o = JSON.parse(l); return o && typeof o.t === "string" ? o.t : ""; } catch (e) { return ""; } };
    const out = all.map((l, i) => ({ l, i, t: tOf(l) })).sort((a, b) => (a.t < b.t ? -1 : a.t > b.t ? 1 : a.i - b.i)).map((x) => x.l);
    return { text: out.length ? out.join("\n") : null, head: mine[0] || null };
  }

  function mergeRequest(S, r, ws) {
    const P = S.SPLIT, ST = S.STATUS;
    const w = pickRead(r, ws);
    const nz = (v) => (v === undefined ? null : v);
    const base = w && w.dr_status !== null && w.dr_status !== undefined ? Number(w.dr_status) : P.newStatus;
    const it = r.dr_reqstatus === null || r.dr_reqstatus === undefined ? null : Number(r.dr_reqstatus);
    const itAt = tms(r.dr_reqstatuson), wAt = w ? tms(w.dr_statuson) : null;
    const fresh = it !== null && itAt !== null && (wAt === null || itAt > wAt);
    const ok = fresh && ((it === ST.cancelled && P.selfEditable.includes(base)) || (it === ST.closed && base === ST.delivered));
    const status = ok ? it : base;
    const inWin = P.selfEditable.includes(status);
    const shared = (k) => (w && !inWin ? nz(w[k]) : nz(r[k]));
    const wv = (k) => (w ? nz(w[k]) : null);
    const done = P.deliveryDone.includes(status);
    const author = { name: (r.Author && r.Author.Title) || "", email: String((r.Author && r.Author.EMail) || "").toLowerCase() };
    const eff = {
      dr_status: status, dr_priority: shared("dr_priority"), dr_sladue: shared("dr_sladue"),
      dr_teamnote: wv("dr_teamnote"), dr_deliverylink: wv("dr_deliverylink"), dr_checklist: wv("dr_checklist"),
      dr_registeredon: wv("dr_registeredon"), dr_deliveredon: wv("dr_deliveredon"),
      dr_closedon: ok && it === ST.closed ? r.dr_reqstatuson : wv("dr_closedon"),
      dr_activity: mergeActivity(w && w.dr_activity, r.dr_reqactivity, author),
      dr_assignees: wv("dr_assignees"),
      dr_teamconv: wv("dr_teamconv"),
      dr_conversation: mergeChat(w && w.dr_teamconv, r.dr_conversation, author),
      dr_ticket: "DR-" + String(r.Id).padStart(4, "0"),
      dr_deliveryfeedback: done ? nz(r.dr_deliveryfeedback) : null,
      dr_feedbacknote: done ? nz(r.dr_feedbacknote) : null,
    };
    return { w, eff, base, intentOk: ok };
  }

  function makeSpXrm(opts) {
    opts = opts || {};
    const S = opts.schema || schemaOf();
    if (!S) throw new Error("ไม่พบ drs_schema (โหลด drs_schema.js ก่อน)");
    const site = String(opts.siteUrl || "").replace(/\/$/, "");
    if (!/^https?:\/\//.test(site)) throw new Error("ต้องระบุ siteUrl ของไซต์ SharePoint");
    const F = opts.fetch || ((u, o) => root.fetch(u, o));
    const sleep = opts.sleep || ((ms) => new Promise((r) => setTimeout(r, ms)));
    const warn = opts.warn || ((m) => { try { console.warn(m); } catch (e) {} });
    const prefix = opts.prefix === undefined ? S.PREFIX : opts.prefix;
    const N = S.names(prefix), E = S.entities(prefix);
    const state = { unsupported: [], warned: new Set(), pages: new Map(), pageSeq: 0, me: null, groups: [], lazyWork: new Set() };
    const now = opts.now || (() => new Date().toISOString());
    const P = S.SPLIT, ST = S.STATUS;
    // ช่องของใบงานที่ตัวแทนคิดจากสองแหล่ง (ไม่อ่านคอลัมน์ชื่อเดียวกันใน Requests ที่ผู้แจ้งแก้ได้)
    const MERGED = new Set([...P.teamOnly, ...P.shared, ...P.gated, P.activity, ...P.derived, P.chat]);
    // คอลัมน์ของ Requests ที่ต้องมีเสมอเพื่อคิดค่ารวม
    const R_BASE = ["Id", "AuthorId", "Author/Title", "Author/EMail", "Created", "Modified", "dr_priority", "dr_sladue", "dr_reqstatus", "dr_reqstatuson", "dr_reqactivity", "dr_deliveryfeedback", "dr_feedbacknote", "dr_conversation"];
    const W_SEL = () => ["Id", "AuthorId", "OData__UIVersionString", ...E.dr_datarequest.work.columns.map((c) => c[1])].join(",");
    const isTeam = () => state.groups.some((g) => g.title === N.groups.team || g.title === N.groups.admin);
    const ticketOf = (id) => "DR-" + String(id).padStart(4, "0");

    function unsupported(entity, what, soft) {
      const msg = `${TAG} ${entity} ${what}`;
      state.unsupported.push(msg);
      if (soft) { if (!state.warned.has(msg)) { state.warned.add(msg); warn(msg); } return msg; }
      const e = new Error(msg); e.drsUnsupported = true; e.status = 501; e.httpStatusCode = 501; throw e;
    }

    // ---------- HTTP ----------
    let digest = null, digestAt = 0;
    async function getDigest(force) {
      if (digest && !force && Date.now() - digestAt < 20 * 60000) return digest;
      const r = await F(site + "/_api/contextinfo", { method: "POST", headers: { Accept: MM }, credentials: "same-origin" });
      if (!r.ok) throw httpErr(r.status, "ขอ digest ไม่ได้");
      digest = (await r.json()).FormDigestValue; digestAt = Date.now();
      return digest;
    }
    function httpErr(status, message, code) {
      const e = new Error(message || "HTTP " + status); e.status = status; e.httpStatusCode = status; e.errorCode = code; e.code = code; return e;
    }
    async function sp(method, path, body, extra) {
      for (let i = 0, fresh = false; ; i++) {
        const headers = Object.assign({ Accept: MM }, extra || {});
        if (method !== "GET") { headers["X-RequestDigest"] = await getDigest(fresh); if (body !== undefined) headers["Content-Type"] = MM; }
        const url = /^https?:/.test(path) ? path : site + path;
        const r = await F(url, { method, headers, credentials: "same-origin", body: body === undefined ? undefined : JSON.stringify(body) });
        if ((r.status === 429 || r.status === 503) && i < 8) {
          const ra = r.headers && r.headers.get ? r.headers.get("Retry-After") : null;
          const sec = ra !== null && ra !== "" && Number.isFinite(Number(ra)) ? Number(ra) : 5 * 2 ** Math.min(i, 5);
          await sleep(Math.min(300, sec) * 1000);
          continue;
        }
        const text = await r.text();
        if (method !== "GET" && r.status === 403 && !fresh && /security validation/i.test(text)) { fresh = true; continue; }
        let b = null;
        try { b = text ? JSON.parse(text) : null; } catch (e) { b = null; }
        if (!r.ok) {
          const er = b && (b["odata.error"] || b.error);
          const msg = er ? (er.message && (er.message.value || er.message)) : text.slice(0, 200);
          throw httpErr(r.status, String(msg || "HTTP " + r.status), er && er.code);
        }
        return { status: r.status, body: b, etag: r.headers && r.headers.get ? r.headers.get("ETag") : null };
      }
    }
    const listPath = (meta) => `/_api/web/lists/GetByTitle('${encodeURIComponent(meta.list.replace(/'/g, "''"))}')`;

    // ---------- ผู้ใช้ปัจจุบัน + บทบาทจากกลุ่ม ----------
    const ctx = { userSettings: { userId: "", userName: "", securityRoles: [] }, getClientUrl: () => site };
    const ready = (async () => {
      const u = await sp("GET", "/_api/web/currentuser?$select=Id,Title,Email");
      state.me = { id: String(u.body.Id), name: u.body.Title || "", email: u.body.Email || "" };
      ctx.userSettings.userId = "{" + state.me.id + "}";
      ctx.userSettings.userName = state.me.name;
      const g = await sp("GET", "/_api/web/currentuser/groups?$select=Id,Title");
      state.groups = ((g.body && g.body.value) || []).map((x) => ({ id: x.Id, title: x.Title }));
    })();
    ready.catch(() => {});

    function synthRoles() {
      const out = [];
      for (const [k, title] of Object.entries(N.groups)) {
        const g = state.groups.find((x) => x.title === title);
        if (g) out.push({ roleid: "drs-group-" + g.id, name: S.ROLE_NAME_OF_GROUP[k] });
      }
      return out;
    }
    function myIdFrom(v) { return String(v || "").replace(/[{}]/g, "").toLowerCase(); }

    // ---------- แปลงคอลัมน์ ----------
    function colMap(meta) {
      const m = new Map();
      for (const [lg, spn, type] of meta.columns) m.set(lg, { sp: spn, type });
      return m;
    }
    function literal(v, type) {
      const t = v.trim();
      let val;
      if (/^'.*'$/.test(t)) val = { s: t.slice(1, -1).replace(/''/g, "'") };
      else if (/^(true|false)$/i.test(t)) val = { b: /^true$/i.test(t) };
      else if (/^null$/i.test(t)) val = { n: true };
      else if (/^-?\d+(\.\d+)?$/.test(t)) val = { num: t };
      else if (GUID.test(t)) val = { s: t.replace(/[{}]/g, "") };
      else if (/^\d{4}-\d{2}-\d{2}T/.test(t)) val = { s: t };
      else return null;
      if (val.b !== undefined) return val.b ? "1" : "0"; // SharePoint: eq true คืนแถว false เงียบ ๆ ต้องใช้ 1 0 (โน้ต Proven)
      if (val.n) return "null";
      if (val.num !== undefined) return type === "Text" || type === "Note" ? `'${val.num}'` : val.num;
      return `'${val.s.replace(/'/g, "''")}'`;
    }

    function buildItemsQuery(entity, meta, q, maxPage) {
      const cols = colMap(meta);
      const known = Object.keys(q);
      for (const k of known) if (!["$select", "$expand", "$filter", "$orderby", "$top"].includes(k)) unsupported(entity, `ตัวเลือก ${k} (query ${JSON.stringify(q)})`);
      const sel = new Set(["Id"]), out = [], expandSp = new Set();
      const want = q.$select ? q.$select.split(",").map((s) => s.trim()).filter(Boolean) : [];
      const merged = meta.work ? MERGED : null;
      for (const lg of want) {
        if (lg === meta.idKey) { out.push({ lg, kind: "id" }); continue; }
        if (meta.notes && lg === P.internal) { out.push({ lg, kind: "note" }); continue; }
        if (merged && merged.has(lg)) { out.push({ lg, kind: "merged" }); continue; }
        if (cols.has(lg)) { sel.add(cols.get(lg).sp); out.push({ lg, kind: "col", sp: cols.get(lg).sp }); continue; }
        if (S.SYSTEM[lg]) { sel.add(S.SYSTEM[lg]); out.push({ lg, kind: S.SYSTEM[lg] === "AuthorId" ? "owner" : "col", sp: S.SYSTEM[lg] }); continue; }
        if ((meta.legacyNull || []).includes(lg)) continue;
        if (meta.later && meta.later[lg]) { unsupported(entity, `อ่านคอลัมน์ ${lg} (${meta.later[lg]}) ตอบค่าว่าง`, true); continue; }
        const e = new Error(`${TAG} ${entity} ไม่มีคอลัมน์ ${lg} (Could not find a property named '${lg}')`);
        e.status = 400; e.httpStatusCode = 400; e.drsUnsupported = true; state.unsupported.push(e.message); throw e; // แบบเดียวกับ Dataverse ตอบ 400 หน้าจอถอยชั้นเองได้
      }
      if (!want.length) for (const [lg, c] of cols) { if (merged && merged.has(lg)) out.push({ lg, kind: "merged" }); else { sel.add(c.sp); out.push({ lg, kind: "col", sp: c.sp }); } }
      if (!want.length && meta.notes) out.push({ lg: P.internal, kind: "note" });
      if (merged) { R_BASE.forEach((c) => sel.add(c)); expandSp.add("Author"); }
      const expands = [];
      if (q.$expand) {
        for (const x of parseExpand(q.$expand)) {
          if (x.name === "createdby" && !x.bad) {
            sel.add("Author/Title"); sel.add("Author/EMail"); sel.add("Author/Id"); expandSp.add("Author");
            expands.push({ name: "createdby", kind: "author", select: x.select });
          } else if (entity === "dr_datarequest" && x.name === "dr_DataRequest_Assignees" && !x.bad) {
            expands.push({ name: x.name, kind: "assignees", select: x.select });
          } else unsupported(entity, `$expand ${x.name} (query ${JSON.stringify(q)})`);
        }
      }
      const parts = ["$select=" + [...sel].join(",")];
      if (expandSp.size) parts.push("$expand=" + [...expandSp].join(","));
      if (q.$filter) {
        const conds = q.$filter.split(/\s+and\s+/i).map((c) => {
          const m = c.trim().match(/^\(?\s*([A-Za-z0-9_]+)\s+(eq|ne)\s+(.+?)\s*\)?$/);
          if (!m) unsupported(entity, `$filter ${q.$filter}`);
          const [, lg, op, raw] = m;
          let spn, type;
          if (lg === meta.idKey) {
            const t = raw.trim();
            if (/^-?\d+$/.test(t)) return `Id ${op} ${t}`;
            if (meta.clientKey && GUID.test(t)) { spn = meta.clientKey; type = "Text"; }
            else unsupported(entity, `$filter ${q.$filter}`);
          } else if (merged && merged.has(lg)) unsupported(entity, `$filter ช่องรวมสองแหล่ง ${q.$filter}`);
          else if (cols.has(lg)) { spn = cols.get(lg).sp; type = cols.get(lg).type; }
          else if (lg === "_ownerid_value" || lg === "_createdby_value") { const t = raw.trim().replace(/^'|'$/g, ""); if (!/^\d+$/.test(t)) unsupported(entity, `$filter ${q.$filter}`); return `AuthorId ${op} ${t}`; }
          else if (S.SYSTEM[lg]) { spn = S.SYSTEM[lg]; type = "DateTime"; }
          else unsupported(entity, `$filter ${q.$filter}`);
          const lit = literal(raw, type);
          if (lit === null) unsupported(entity, `$filter ${q.$filter}`);
          return `${spn} ${op} ${lit}`;
        });
        parts.push("$filter=" + conds.join(" and "));
      }
      if (q.$orderby) {
        const ob = q.$orderby.split(",").map((x) => {
          const m = x.trim().match(/^([A-Za-z0-9_]+)(?:\s+(asc|desc))?$/i);
          if (!m) unsupported(entity, `$orderby ${q.$orderby}`);
          const lg = m[1], dir = (m[2] || "asc").toLowerCase();
          if (merged && merged.has(lg)) unsupported(entity, `$orderby ช่องรวมสองแหล่ง ${q.$orderby}`);
          const spn = lg === meta.idKey ? "Id" : cols.has(lg) ? cols.get(lg).sp : S.SYSTEM[lg] && S.SYSTEM[lg] !== "AuthorId" ? S.SYSTEM[lg] : null;
          if (!spn) unsupported(entity, `$orderby ${q.$orderby}`);
          return `${spn} ${dir}`;
        });
        parts.push("$orderby=" + ob.join(","));
      }
      const top = q.$top ? Number(q.$top) : maxPage ? Number(maxPage) : 5000;
      if (!Number.isFinite(top) || top < 1) unsupported(entity, `$top ${q.$top}`);
      parts.push("$top=" + Math.min(top, 5000));
      const qs = parts.map((p) => { const i = p.indexOf("="); return p.slice(0, i) + "=" + encodeURIComponent(p.slice(i + 1)); }).join("&");
      return { qs, out, expands, wantAll: !want.length, wantNote: out.some((c) => c.kind === "note") };
    }

    function mapItem(entity, meta, plan, it, ws, nts) {
      const o = {};
      o[meta.idKey] = String(it.Id);
      if (it["odata.etag"] !== undefined) o["@odata.etag"] = it["odata.etag"];
      const m = meta.work ? mergeRequest(S, it, ws) : null;
      // nts = แถวบันทึกภายใน (ทีมที่ขอช่องนี้เท่านั้น ผู้แจ้งเป็น null ไม่ได้ยิง list นี้)
      const nt = nts ? notesOf(nts, it.Id) : null;
      // etag ก้อนเดียวที่หน้าถือ: ของแถวใบ + ของแถวงานที่เชื่อ (+ แถวบันทึก Id ต่ำสุดเมื่ออ่านบันทึก) ตัวแทนแยกใช้ตอนเขียน
      if (m) o["@odata.etag"] = JSON.stringify(Object.assign({ r: it["odata.etag"] || "", w: m.w ? m.w["odata.etag"] || "" : "", wid: m.w ? m.w.Id : null },
        nt ? { n: nt.head ? nt.head["odata.etag"] || "" : "", nid: nt.head ? nt.head.Id : null } : {}));
      for (const c of plan.out) {
        if (c.kind === "note") { o[c.lg] = nt ? nt.text : null; continue; }
        if (c.kind === "merged") o[c.lg] = m.eff[c.lg] === undefined ? null : m.eff[c.lg];
        else if (c.kind === "id") o[c.lg] = String(it.Id);
        else if (c.kind === "owner") o[c.lg] = it.AuthorId === undefined || it.AuthorId === null ? null : String(it.AuthorId);
        else o[c.lg] = it[c.sp] === undefined ? null : it[c.sp];
      }
      for (const x of plan.expands) {
        if (x.kind === "assignees") {
          o[x.name] = parseAssignees(m && m.eff.dr_assignees).map((a) => {
            const full = { dr_teammemberid: a.id, dr_name: a.name };
            if (!x.select || !x.select.length) return full;
            const pick = {}; x.select.forEach((k) => { pick[k] = full[k] === undefined ? null : full[k]; }); return pick;
          });
        }
        if (x.kind === "author") {
          const a = it.Author || {};
          const cb = { systemuserid: a.Id === undefined ? null : String(a.Id), fullname: a.Title || "", internalemailaddress: a.EMail || "" };
          if (x.select && x.select.length) { const pick = {}; x.select.forEach((k) => { pick[k] = cb[k]; }); o.createdby = pick; } else o.createdby = cb;
        }
      }
      return o;
    }

    // ---------- payload ขาเขียน ----------
    function mapPayload(entity, meta, data) {
      const cols = colMap(meta), body = {};
      for (const [k, v] of Object.entries(data || {})) {
        if (k === meta.idKey) {
          if (meta.clientKey && GUID.test(String(v))) { body[meta.clientKey] = String(v).replace(/[{}]/g, "").toLowerCase(); continue; }
          unsupported(entity, `ตั้ง ${k} เอง (${String(v).slice(0, 40)})`);
        }
        if (k.includes("@odata.bind")) unsupported(entity, `ผูกความสัมพันธ์ ${k}`);
        if (meta.later && meta.later[k]) unsupported(entity, `เขียนคอลัมน์ ${k} (${meta.later[k]})`);
        if (!cols.has(k)) unsupported(entity, `เขียนคอลัมน์ ${k} (ไม่มีใน list)`);
        const c = cols.get(k);
        body[c.sp] = c.type === "Number" && v !== null && v !== "" && v !== undefined ? Number(v) : v === "" && c.type === "DateTime" ? null : v;
      }
      return body;
    }

    async function itemByClientKey(meta, guid) {
      const g = String(guid).replace(/[{}]/g, "").toLowerCase();
      const r = await sp("GET", `${listPath(meta)}/items?$select=Id&$filter=${encodeURIComponent(`${meta.clientKey} eq '${g}'`)}&$top=1`);
      const v = (r.body && r.body.value) || [];
      return v.length ? v[0].Id : null;
    }

    // ---------- DRS Work ----------
    const workPath = () => listPath(E.dr_datarequest.work);
    // อ่านแถวงานทุกแถวที่ผู้ใช้นี้เห็น (ทีมเห็นทุกแถว ผู้แจ้งเห็นแถวที่ตัวเองสร้าง) อ่านไม่ได้ = throw ไม่โชว์ค่าเริ่มต้นหลอกว่าสถานะรอคิว
    async function fetchAllWork() {
      let url = `${workPath()}/items?$select=${W_SEL()}&$top=5000`, out = [];
      for (let i = 0; url && i < 100; i++) {
        const r = await sp("GET", url);
        out = out.concat((r.body && r.body.value) || []);
        url = r.body && (r.body["odata.nextLink"] || r.body["@odata.nextLink"]);
      }
      return out;
    }
    // ---------- DRS Internal Notes (ทีม แอดมิน เท่านั้น) ----------
    const notesPath = () => listPath(E.dr_datarequest.notes);
    const N_SEL = "Id,dr_requestid,dr_internalnote";
    // ผู้แจ้ง: ไม่ยิงเลย ได้ null (Dataverse คืน null ให้คอลัมน์ที่ FSP กันไว้) , ทีม: n = เลขใบ หรือ null = ทุกแถว
    // list ยังไม่ติดตั้งหรือทีมไม่มีสิทธิ์ (404 403) = เตือนครั้งเดียว ได้ null เหมือนผู้แจ้ง ไม่ทำให้หน้าล้ม
    async function readNotes(n) {
      if (!isTeam()) return null;
      try {
        let url = `${notesPath()}/items?$select=${N_SEL}` + (n === null ? "&$top=5000" : `&$filter=${encodeURIComponent(`dr_requestid eq ${Number(n)}`)}&$top=100`), out = [];
        for (let i = 0; url && i < 100; i++) {
          const r = await sp("GET", url);
          out = out.concat((r.body && r.body.value) || []);
          url = r.body && (r.body["odata.nextLink"] || r.body["@odata.nextLink"]);
        }
        return out;
      } catch (e) {
        if (e.status !== 403 && e.status !== 404) throw e;
        if (!state.warned.has("notes")) { state.warned.add("notes"); warn(`[DRS] อ่าน list ${E.dr_datarequest.notes.list} ไม่ได้ (${e.status}) บันทึกภายในทีมเป็นค่าว่าง`); }
        return null;
      }
    }
    // ทีมเขียน: แบบต่อท้ายอย่างเดียว (หน้าเขียนบันทึกแค่ต่อท้าย dashboard 7267 13312 ไม่มีลบหรือแก้)
    // อ่านแถว Id ต่ำสุดของใบสดทุกครั้ง รวมของจริงกับที่หน้าส่ง ตัดซ้ำ แล้ว MERGE ด้วย etag ของแถวที่เพิ่งอ่าน ชน 412 อ่านใหม่ลองซ้ำ
    // ไม่พึ่ง etag ของหน้า: หน้าที่อ่านบันทึกไม่ได้ (เห็นค่าว่าง) แล้วเพิ่มบันทึก จะไม่ทับบันทึกเดิม (unit_red_union_run1.log เทส 7)
    async function writeNotes(n, value) {
      for (let attempt = 0; ; attempt++) {
        const rows = (await sp("GET", `${notesPath()}/items?$select=${N_SEL}&$filter=${encodeURIComponent(`dr_requestid eq ${Number(n)}`)}&$top=100`)).body;
        const head = (((rows && rows.value) || []).slice().sort(byId))[0] || null;
        const text = uniqLines((head ? head.dr_internalnote || "" : "") + "\n" + String(value || "")).join("\n") || null;
        if (!head) {
          const c = await sp("POST", `${notesPath()}/items`, { Title: ticketOf(n), dr_requestid: Number(n), dr_internalnote: text });
          return { n: (c.body && c.body["odata.etag"]) || c.etag || "", nid: c.body && c.body.Id };
        }
        try {
          const r = await sp("POST", `${notesPath()}/items(${head.Id})`, { dr_internalnote: text }, MERGEH(head["odata.etag"] || "*"));
          return { n: r.etag || "", nid: head.Id };
        } catch (e) {
          if (e.status === 412 && attempt < 3) continue;
          throw e;
        }
      }
    }

    async function fetchWorkFor(n) {
      const r = await sp("GET", `${workPath()}/items?$select=${W_SEL()}&$filter=${encodeURIComponent(`dr_requestid eq ${Number(n)}`)}&$top=100`);
      return (r.body && r.body.value) || [];
    }
    // ผู้แจ้งที่ยังไม่มีแถวงานของตัวเองสำหรับใบของตัวเอง (สร้างแถวพลาดตอนแจ้ง หรือทีมสร้างแถวไว้ก่อน) สร้างให้เบื้องหลัง ครั้งต่อไปที่ทีมเขียนจะย้ายค่ามาแถวนี้
    function lazyWork(items, ws) {
      if (isTeam()) return;
      const me = Number(state.me.id);
      const missing = items.filter((it) => Number(it.AuthorId) === me && !state.lazyWork.has(it.Id) && !ws.some((w) => Number(w.dr_requestid) === Number(it.Id) && Number(w.AuthorId) === me)).slice(0, 10);
      missing.forEach((it) => state.lazyWork.add(it.Id));
      (async () => { for (const it of missing) { try { await sp("POST", `${workPath()}/items`, { Title: ticketOf(it.Id), dr_requestid: it.Id }); } catch (e) { warn(`[DRS] สร้างแถวงานของใบ ${it.Id} ไม่สำเร็จ ${e.status || ""}`); } } })();
    }
    function parseEtag(etag) {
      if (!etag || etag === "*") return { r: "*", w: "", wid: null };
      try { const o = JSON.parse(etag); if (o && typeof o === "object" && "r" in o) return Object.assign({ r: o.r || "*", w: o.w || "", wid: o.wid === undefined ? null : o.wid }, "n" in o ? { n: o.n || "", nid: o.nid === undefined ? null : o.nid } : {}); } catch (e) { /* etag เดี่ยวของ SharePoint */ }
      return { r: etag, w: "", wid: null };
    }
    const deny = (msg) => { throw httpErr(403, `[DRS] ${msg} (403)`, "DRS_FORBIDDEN"); };
    const MERGEH = (m) => ({ "X-HTTP-Method": "MERGE", "IF-MATCH": m || "*" });
    function workValue(k, v) {
      const c = E.dr_datarequest.work.columns.find((x) => x[0] === k);
      if (!c) unsupported("dr_datarequest", `เขียนคอลัมน์ ${k} ลง DRS Work`);
      return c[2] === "Number" && v !== null && v !== "" && v !== undefined ? Number(v) : v === "" && c[2] === "DateTime" ? null : v === undefined ? null : v;
    }

    // ---------- WebApi ----------
    async function retrieveMultipleRecords(entity, query, maxPageSize) {
      await ready;
      const q0 = String(query || "");
      const cont = q0.match(/[?&]__drsnext=(\d+)/);
      if (cont) {
        const pg = state.pages.get(cont[1]);
        if (!pg) unsupported(entity, `หน้าถัดไปที่ไม่รู้จัก ${q0}`);
        state.pages.delete(cont[1]);
        const r = await sp("GET", pg.url);
        return wrapPage(entity, pg.meta, pg.plan, r.body, pg.ws, pg.nts);
      }
      const q = parseQuery(q0);
      if (entity === "systemuser") {
        const f = (q.$filter || "").match(/^systemuserid eq \{?([^}\s]+)\}?$/);
        if (!f || myIdFrom(f[1]) !== state.me.id) unsupported(entity, `query ${q0}`);
        const ex = q.$expand ? parseExpand(q.$expand).map((x) => x.name) : [];
        if (ex.some((n) => !["systemuserroles_association", "teammembership_association"].includes(n))) unsupported(entity, `query ${q0}`);
        const me = { systemuserid: state.me.id };
        if (ex.includes("systemuserroles_association")) me.systemuserroles_association = synthRoles();
        if (ex.includes("teammembership_association")) me.teammembership_association = []; // บทบาทมาจากกลุ่ม SharePoint ตรง ๆ ไม่มีชั้นทีม
        return { entities: [me], nextLink: undefined };
      }
      if (entity === "role" && /^_roletemplateid_value eq /.test(q.$filter || "")) return { entities: [], nextLink: undefined }; // SharePoint ไม่มี System Administrator ของ Dataverse
      if (entity === "annotation") return listAnnotations(q);
      const meta = E[entity];
      if (!meta) unsupported(entity, `query ${q0}`);
      const plan = buildItemsQuery(entity, meta, q, maxPageSize);
      const r = await sp("GET", `${listPath(meta)}/items?${plan.qs}`);
      let ws = null;
      if (meta.work) { ws = await fetchAllWork(); lazyWork((r.body && r.body.value) || [], ws); }
      const nts = plan.wantNote ? await readNotes(null) : null;
      return wrapPage(entity, meta, plan, r.body, ws, nts);
    }
    function wrapPage(entity, meta, plan, body, ws, nts) {
      const items = (body && body.value) || [];
      const res = { entities: items.map((it) => mapItem(entity, meta, plan, it, ws, nts)) };
      const nl = body && (body["odata.nextLink"] || body["@odata.nextLink"]);
      if (nl) { const id = String(++state.pageSeq); state.pages.set(id, { url: nl, meta, plan, ws, nts }); res.nextLink = `${site}/drs/next?__drsnext=${id}`; }
      return res;
    }

    async function retrieveRecord(entity, id, query) {
      await ready;
      const q = parseQuery(query);
      if (entity === "systemuser") {
        if (myIdFrom(id) !== state.me.id) unsupported(entity, `อ่านผู้ใช้คนอื่น ${id}`);
        const all = { systemuserid: state.me.id, internalemailaddress: state.me.email, fullname: state.me.name };
        if (!q.$select) return all;
        const o = {}; q.$select.split(",").forEach((k) => { k = k.trim(); if (!(k in all)) unsupported(entity, `คอลัมน์ ${k}`); o[k] = all[k]; });
        return o;
      }
      if (entity === "annotation") return getAnnotation(id, q);
      const meta = E[entity];
      if (!meta) unsupported(entity, `id ${id} query ${query || ""}`);
      let n = String(id).replace(/[{}]/g, "");
      if (!/^\d+$/.test(n)) {
        if (meta.clientKey && GUID.test(n)) { const found = await itemByClientKey(meta, n); if (!found) throw httpErr(404, `ไม่พบ ${entity} ${n}`); n = String(found); }
        else unsupported(entity, `id ${id}`);
      }
      const plan = buildItemsQuery(entity, meta, Object.assign({}, q, { $top: undefined }), 0);
      const qs = plan.qs.replace(/&?\$top=\d+/, "");
      const r = await sp("GET", `${listPath(meta)}/items(${n})?${qs}`);
      return mapItem(entity, meta, plan, r.body, meta.work ? await fetchWorkFor(n) : null, plan.wantNote ? await readNotes(n) : null);
    }

    async function createRecord(entity, data) {
      await ready;
      if (entity === "annotation") return createAnnotation(data);
      const meta = E[entity];
      if (!meta) unsupported(entity, `สร้าง ${JSON.stringify(Object.keys(data || {}))}`);
      // หน้าไม่สร้างใบพร้อมบันทึกภายใน (เขียนผ่าน updateRecord หลังมีใบเท่านั้น) เจอ = ล้มดัง ไม่ทิ้งเงียบ
      if (meta.notes && data && Object.prototype.hasOwnProperty.call(data, S.SPLIT.internal)) unsupported(entity, `สร้างใบพร้อม ${S.SPLIT.internal} (เขียนบันทึกภายในหลังสร้างใบ)`);
      const team = !!meta.work && isTeam();
      let plain = data;
      const wInit = {};
      if (meta.work) {
        // ช่องของทีมไม่ลง Requests: ทีมใส่ลงแถวงาน ผู้แจ้งส่งได้แค่สถานะเริ่ม (หน้าใส่รอคิวเอง ตัวแทนไม่ใช้ ค่าเริ่มคิดตอนอ่าน)
        plain = {};
        for (const [k, v] of Object.entries(data || {})) {
          if (P.derived.includes(k)) continue;
          if (k === P.chat) {
            const sp_ = splitChat(v);
            if (team) wInit.dr_teamconv = joinChat(sp_.team);
            plain[k] = joinChat(sp_.other);
            continue;
          }
          if (P.teamOnly.includes(k) || k === P.activity) {
            if (team) wInit[k] = v;
            else if (k !== "dr_status") deny(`ช่อง ${k} เป็นช่องของทีม ผู้แจ้งใส่ตอนสร้างไม่ได้`);
            continue;
          }
          plain[k] = v;
        }
      }
      const body = mapPayload(entity, meta, plain);
      let created;
      try { created = await sp("POST", `${listPath(meta)}/items`, body); }
      catch (e) {
        // คีย์ฝั่งหน้าจอซ้ำ = ใบนี้ถูกสร้างไปแล้วรอบก่อน (ลองใหม่หลังเน็ตสะดุด R2G-5) คืน id เดิม ไม่สร้างซ้ำ
        if (meta.clientKey && body[meta.clientKey] && /Duplicate/i.test(String(e.code || "") + String(e.message || ""))) {
          const old = await itemByClientKey(meta, body[meta.clientKey]);
          if (old) return { entityType: entity, id: String(old) };
        }
        throw e;
      }
      const newId = created.body.Id;
      if (meta.work) {
        // แถวงาน: ผู้สร้างแถว = ผู้สร้างใบ ผู้แจ้งจึงอ่านสถานะใบตัวเองได้ (เลขตั๋วคิดจาก Id ตอนอ่าน ไม่เขียนลงช่องใด)
        state.lazyWork.add(newId);
        try {
          const w = await sp("POST", `${workPath()}/items`, { Title: ticketOf(newId), dr_requestid: newId });
          if (team) {
            // ทีมสร้างเอง (แจ้งแทน): เขียนอีกรอบให้แถวขึ้นเวอร์ชัน 2 จึงถูกเชื่อ
            const full = { dr_status: P.newStatus, dr_statuson: now(), dr_priority: body.dr_priority === undefined ? null : body.dr_priority, dr_sladue: body.dr_sladue === undefined ? null : body.dr_sladue };
            for (const [k, v] of Object.entries(wInit)) full[k] = workValue(k, v);
            await sp("POST", `${workPath()}/items(${w.body.Id})`, full, MERGEH(w.body["odata.etag"]));
          }
        } catch (e) { state.lazyWork.delete(newId); warn(`[DRS] สร้างแถวงานของใบ ${newId} ไม่สำเร็จ ${e.status || ""} ${e.message}`); }
      }
      return { entityType: entity, id: String(newId) };
    }

    // ทีมเขียนแถวงานของใบ n: อ่านใบกับแถวงานสด build(m, r) คืนช่องที่จะเขียน (null = ไม่เขียน)
    // ครั้งแรกที่ทีมเขียนแถว (เวอร์ชัน 1.0) เขียนทุกช่องใหม่จากค่ารวมปัจจุบัน ค่าที่ผู้สร้างแถวใส่มาเองถูกล้าง แถวที่เชื่อแล้วและไม่มีอะไรเขียน = ไม่ยิง
    // pageEtag (etag ที่หน้าถือ) ใช้เมื่อเป็นแถวเดียวกัน ชน 412 ส่งคืนหน้า , ไม่มี pageEtag (ผู้รับผิดชอบ ไฟล์แนบ) ชน 412 อ่านใหม่แล้วลองอีก ไม่เกิน 3 รอบ
    async function teamWork(n, build, pageEtag) {
      const meta = E.dr_datarequest;
      for (let attempt = 0; ; attempt++) {
        const rq = await sp("GET", `${listPath(meta)}/items(${n})?$select=${encodeURIComponent(R_BASE.join(","))}&$expand=Author`);
        const r = rq.body, ws = await fetchWorkFor(n), m = mergeRequest(S, r, ws), target = pickTarget(r, ws);
        const wPay = build(m, r);
        if (wPay === null) return null;
        let row = target, ifm;
        if (!row) { row = (await sp("POST", `${workPath()}/items`, { Title: ticketOf(n), dr_requestid: Number(n) })).body; ifm = row["odata.etag"]; }
        else ifm = pageEtag && pageEtag.wid === row.Id && pageEtag.w ? pageEtag.w : row["odata.etag"];
        let body = wPay;
        if (verOf(row) <= 1) {
          const old = m.w;
          body = Object.assign({ Title: ticketOf(n), dr_requestid: Number(n), dr_status: m.eff.dr_status, dr_statuson: now(), dr_priority: m.eff.dr_priority, dr_sladue: m.eff.dr_sladue,
            dr_teamnote: m.eff.dr_teamnote, dr_deliverylink: m.eff.dr_deliverylink, dr_checklist: m.eff.dr_checklist, dr_activity: old ? old.dr_activity || null : null,
            dr_registeredon: m.eff.dr_registeredon, dr_closedon: m.eff.dr_closedon, dr_deliveredon: m.eff.dr_deliveredon, dr_assignees: m.eff.dr_assignees, dr_teamconv: m.eff.dr_teamconv }, wPay);
        } else if (!Object.keys(wPay).length) return { wEtag: row["odata.etag"] || "", wid: row.Id };
        try {
          const wr = await sp("POST", `${workPath()}/items(${row.Id})`, body, MERGEH(ifm));
          return { wEtag: wr.etag || "", wid: row.Id };
        } catch (e) {
          if (e.status === 412 && !pageEtag && attempt < 3) continue;
          throw e;
        }
      }
    }

    // ใบงาน: แยกเขียนสองแหล่งตามบทบาท (กติกาที่หัว mergeRequest)
    async function updateRequest(n, data, etag) {
      const meta = E.dr_datarequest, et = parseEtag(etag), team = isTeam();
      const has = (k) => Object.prototype.hasOwnProperty.call(data || {}, k);
      // บันทึกภายในทีม: ผู้แจ้งเขียนไม่ได้ (FSP เดิม) ปฏิเสธก่อนเขียนอะไรออกไป
      if (has(P.internal) && !team) deny("บันทึกภายในทีมเป็นของทีม ผู้แจ้งเขียนไม่ได้");
      const plain = {};
      for (const [k, v] of Object.entries(data || {})) if (!P.derived.includes(k) && !P.teamOnly.includes(k) && k !== P.activity && k !== P.chat && k !== P.internal) plain[k] = v; // ช่องผู้แจ้ง + ความสำคัญ วันครบกำหนด ไป Requests
      const chat = has(P.chat) ? splitChat(data[P.chat]) : null;
      const rPay = mapPayload("dr_datarequest", meta, plain);
      let rEtag = et.r, wEtag = et.w, wid = et.wid;
      if (!team) {
        if (has("dr_status")) {
          const v = Number(data.dr_status);
          if (v !== ST.cancelled && v !== ST.closed) deny("ผู้แจ้งเปลี่ยนสถานะนี้เองไม่ได้ ได้แค่ยกเลิกหรือปิดงาน");
          rPay.dr_reqstatus = v; rPay.dr_reqstatuson = now();
        }
        for (const k of P.teamOnly) {
          if (k === "dr_status" || !has(k)) continue;
          if ((k === "dr_closedon" || k === "dr_deliveredon") && has("dr_status")) continue; // วันที่หน้าเติมตอนเปลี่ยนสถานะ ตัวแทนคิดเองตอนอ่าน
          deny(`ช่อง ${k} เป็นช่องของทีม ผู้แจ้งแก้ไม่ได้`);
        }
        if (chat) {
          // ผู้แจ้ง: บรรทัด by team ที่ส่งมา (แก้ ลบ ปลอม ข้อความทีม) ทิ้งทั้งหมด ข้อความทีมจริงอยู่ dr_teamconv ที่ผู้แจ้งแก้ไม่ได้
          const cur = await sp("GET", `${listPath(meta)}/items(${n})?$select=dr_conversation`);
          rPay.dr_conversation = joinChat(legacyTeam(cur.body && cur.body.dr_conversation).concat(chat.other));
        }
        if (has(P.activity)) {
          const w = pickRead({ Id: Number(n), AuthorId: Number(state.me.id) }, await fetchWorkFor(n));
          const wKeys = new Set(parseAct(w && w.dr_activity).map(canon));
          const mine = parseAct(data[P.activity]).filter((o) => !wKeys.has(canon(o)));
          rPay.dr_reqactivity = mine.length ? mine.map((o) => JSON.stringify(o)).join("\n") : null;
        }
      } else {
        const wPay = {};
        for (const k of [...P.teamOnly, ...P.shared]) if (has(k)) wPay[k] = workValue(k, data[k]);
        if (has(P.activity)) wPay[P.activity] = data[P.activity] || null;
        if (chat) wPay.dr_teamconv = joinChat(chat.team);
        let curConv = null;
        if (Object.keys(wPay).length) {
          const w = await teamWork(n, (m, r) => {
            curConv = r.dr_conversation;
            const pay = Object.assign({}, wPay);
            if (has("dr_status")) {
              pay.dr_statuson = now();
              // ออกจากช่วงที่ผู้แจ้งแก้ใบได้: คัดความสำคัญ วันครบกำหนด ของผู้แจ้งตอนนี้มาเก็บ นอกช่วงนี้ค่าที่ผู้แจ้งแก้ทีหลังไม่มีผล
              if (P.selfEditable.includes(m.eff.dr_status) && !P.selfEditable.includes(Number(data.dr_status))) for (const k of P.shared) if (!has(k)) pay[k] = r[k] === undefined ? null : r[k];
            }
            return pay;
          }, et.w ? et : { r: et.r, w: "", wid: -1 });
          wEtag = w.wEtag; wid = w.wid;
        }
        // ทีม: บรรทัดที่ไม่ใช่ทีม (ของผู้แจ้ง ที่ทีมปักหมุดหรือแอดมินลบถาวร) ลง Requests พร้อมบรรทัดดิบ by team เดิม
        if (chat) rPay.dr_conversation = joinChat(legacyTeam(curConv).concat(chat.other));
      }
      if (Object.keys(rPay).length) { const rr = await sp("POST", `${listPath(meta)}/items(${n})`, rPay, MERGEH(et.r)); rEtag = rr.etag || et.r; }
      // บันทึกเขียนหลังสุด (หลังแชท) ถ้าแชทชน 412 หน้าอ่านใหม่แล้วส่งโน้ตเดิมซ้ำ ตัดซ้ำใน writeNotes
      const nt = has(P.internal) ? await writeNotes(n, data[P.internal]) : "n" in et ? { n: et.n, nid: et.nid } : null;
      return { status: 204, etag: JSON.stringify(Object.assign({ r: rEtag, w: wEtag, wid }, nt || {})) };
    }

    // ใช้ทั้ง updateRecord (If-Match * แบบเดียวกับ Xrm.WebApi.updateRecord ที่ไม่เช็คชน) และตัวดัก PATCH ของหน้าที่ส่ง etag มา
    async function update(entity, id, data, etag) {
      await ready;
      const meta = E[entity];
      if (!meta) unsupported(entity, `แก้ id ${id}`);
      const n = String(id).replace(/[{}]/g, "");
      if (!/^\d+$/.test(n)) unsupported(entity, `แก้ id ${id}`);
      if (meta.work) return updateRequest(n, data, etag);
      const body = mapPayload(entity, meta, data);
      const r = await sp("POST", `${listPath(meta)}/items(${n})`, body, { "X-HTTP-Method": "MERGE", "IF-MATCH": etag || "*" });
      return { status: r.status, etag: r.etag };
    }
    async function updateRecord(entity, id, data) { await update(entity, id, data, "*"); return { entityType: entity, id: String(id) }; }
    async function deleteRecord(entity, id) {
      await ready;
      if (entity === "annotation") return deleteAnnotation(id);
      const meta = E[entity];
      if (!meta) unsupported(entity, `ลบ id ${id}`);
      const n = String(id).replace(/[{}]/g, "");
      if (!/^\d+$/.test(n)) unsupported(entity, `ลบ id ${id}`);
      await sp("POST", `${listPath(meta)}/items(${n})`, undefined, { "X-HTTP-Method": "DELETE", "IF-MATCH": "*" });
      return { entityType: entity, id: n };
    }

    // ---------- ขั้น 2 ผู้รับผิดชอบ (S28 S29) ----------
    // ทีมเท่านั้น ผู้แจ้งเรียก = 403 (หน้าเรียกตอนสร้างใบเมื่อแอดมินเปิดช่องเลือกผู้รับผิดชอบให้ผู้ใช้ ได้ข้อความบันทึกไม่สำเร็จ ทีมมอบเองทีหลัง)
    // ผูกซ้ำหรือเอาออกคนที่ไม่มีอยู่ = สำเร็จเฉย ๆ ไม่เขียน (หน้าอ่านสดก่อนรับงานอยู่แล้ว)
    async function setAssignee(entity, recId, tmId, add) {
      await ready;
      if (entity !== "dr_datarequest") unsupported(entity, `ผูกผู้รับผิดชอบกับ ${entity}`);
      const n = String(recId || "").replace(/[{}]/g, ""), t = String(tmId || "").replace(/[{}]/g, "");
      if (!/^\d+$/.test(n) || !/^\d+$/.test(t)) unsupported(entity, `ผู้รับผิดชอบ id ${recId} ${tmId}`);
      if (!isTeam()) deny("ผู้แจ้งเพิ่มหรือเอาผู้รับผิดชอบออกเองไม่ได้ (ทีมเป็นคนมอบ)");
      let name = "";
      if (add) {
        const mm = await sp("GET", `${listPath(E.dr_teammember)}/items(${t})?$select=Id,Title`);
        name = (mm.body && mm.body.Title) || "";
      }
      await teamWork(n, (m) => {
        const cur = parseAssignees(m.eff.dr_assignees), has = cur.some((a) => a.id === t);
        if (add === has) return {};
        const next = add ? cur.concat([{ id: t, name }]) : cur.filter((a) => a.id !== t);
        return { dr_assignees: next.length ? JSON.stringify(next) : null };
      }, null);
      return { ok: true, status: 204 };
    }
    async function execute(req) {
      const md = req && req.getMetadata ? req.getMetadata() : {};
      const op = md.operationName || "?", ent = (req && req.target && req.target.entityType) || "?";
      if (req && req.relationship === "dr_DataRequest_Assignees" && (op === "Associate" || op === "Disassociate")) {
        const tm = op === "Associate" ? ((req.relatedEntities || [])[0] || {}).id : req.relatedEntityId;
        if (op === "Associate" && (req.relatedEntities || []).length !== 1) unsupported(ent, `online.execute Associate หลายแถวพร้อมกัน`);
        return setAssignee(ent, req.target && req.target.id, tm, op === "Associate");
      }
      unsupported(ent, `online.execute ${op} ${req && req.relationship ? req.relationship : ""}`);
    }

    // ---------- ขั้น 3 ไฟล์แนบ (annotation) ----------
    const DELIV = "drs-delivery--", TEAMG = "drs-team--";
    const MIME = { pdf: "application/pdf", png: "image/png", jpg: "image/jpeg", jpeg: "image/jpeg", gif: "image/gif", webp: "image/webp", bmp: "image/bmp",
      txt: "text/plain", csv: "text/csv", json: "application/json", zip: "application/zip",
      xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", xls: "application/vnd.ms-excel",
      docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document", doc: "application/msword",
      pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation", ppt: "application/vnd.ms-powerpoint", pbix: "application/octet-stream" };
    const mimeOf = (name) => { const m = String(name).match(/\.([A-Za-z0-9]+)$/); return (m && MIME[m[1].toLowerCase()]) || "application/octet-stream"; };
    // ชื่อที่โชว์กับ subject: ไฟล์ในแถวงาน (ทีมแนบ ผู้แจ้งแนบที่นี่ไม่ได้ probe C 403) อ่านคำนำหน้า ไฟล์ในแถวใบงาน (ผู้แจ้ง) เป็นไฟล์ทั่วไปเสมอ
    function fileInfo(stored, onWork) {
      let name = stored, delivery = false;
      if (onWork && stored.startsWith(DELIV)) { name = stored.slice(DELIV.length); delivery = true; }
      else if (onWork && stored.startsWith(TEAMG)) name = stored.slice(TEAMG.length);
      name = unbidi(name);
      return { name, subject: delivery ? "delivery" : name === "delivery" ? "delivery (ไฟล์แนบ)" : name };
    }
    const roots = {};
    async function rootOf(meta) {
      if (!roots[meta.list]) roots[meta.list] = (await sp("GET", `${listPath(meta)}/RootFolder?$select=ServerRelativeUrl`)).body.ServerRelativeUrl;
      return roots[meta.list];
    }
    const alias = (v) => "@a='" + encodeURIComponent(String(v).replace(/'/g, "''")) + "'";
    // ไฟล์ในโฟลเดอร์แนบของแถว ยังไม่มีไฟล์ = ไม่มีโฟลเดอร์ 404 ได้รายการว่าง อ่านไม่ได้ (403 ฯลฯ) throw ไม่ตอบว่างหลอก
    async function folderFiles(path) {
      try {
        const r = await sp("GET", `/_api/web/GetFolderByServerRelativePath(DecodedUrl=@a)/Files?$select=Name,UniqueId,Length,TimeCreated&${alias(path)}`);
        return (r.body && r.body.value) || [];
      } catch (e) { if (e.status === 404) return []; throw e; }
    }
    function bytesToB64(buf) {
      const u = new Uint8Array(buf); let str = "";
      for (let i = 0; i < u.length; i += 0x8000) str += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
      return root.btoa(str);
    }
    function b64ToBytes(b64) {
      const str = root.atob(String(b64 || "").replace(/\s/g, "")), u = new Uint8Array(str.length);
      for (let i = 0; i < str.length; i++) u[i] = str.charCodeAt(i);
      return u;
    }
    // คำขอตัวไบต์ (อัปไฟล์ อ่านเนื้อไฟล์) ไม่ผ่าน JSON
    async function spRaw(method, path, bytes) {
      for (let i = 0, fresh = false; ; i++) {
        const headers = { Accept: MM };
        if (method !== "GET") { const dg = await getDigest(fresh); headers["X-RequestDigest"] = dg; }
        const r = await F(site + path, { method, headers, credentials: "same-origin", body: bytes });
        if ([429, 503].includes(r.status) && i < 8) {
          const ra = r.headers && r.headers.get ? r.headers.get("Retry-After") : null;
          await sleep(Math.min(300, ra !== null && ra !== "" && Number.isFinite(Number(ra)) ? Number(ra) : 5 * 2 ** Math.min(i, 5)) * 1000);
          continue;
        }
        if (!r.ok) {
          const text = await r.text();
          if (method !== "GET" && r.status === 403 && !fresh && /security validation/i.test(text)) { fresh = true; continue; }
          let b = null; try { b = JSON.parse(text); } catch (e) { b = null; }
          const er = b && (b["odata.error"] || b.error);
          throw httpErr(r.status, String((er && er.message && (er.message.value || er.message)) || text.slice(0, 200) || "HTTP " + r.status), er && er.code);
        }
        return method === "GET" ? r.arrayBuffer() : r;
      }
    }
    // id ไฟล์ = <เลขใบ>_<UniqueId> (T7c พี่ปริม: แชทใบหนึ่งชี้ไฟล์ของอีกใบได้) รับเฉพาะไฟล์ในโฟลเดอร์แนบของใบนั้น
    // หรือของแถวงานที่เชื่อแล้วของใบนั้น ไฟล์อื่นของไซต์ ไฟล์ใบอื่น แถวงานที่ยังไม่เชื่อ = 403 ไม่อ่านเนื้อ ไม่ลบ
    const annId = (n, uid) => `${n}_${String(uid || "").toLowerCase()}`;
    async function fileById(id) {
      const p = String(id || "").replace(/[{}]/g, "").toLowerCase().match(/^(\d+)_([0-9a-f-]{36})$/);
      if (!p || !GUID.test(p[2])) deny(`id ไฟล์ ${String(id).slice(0, 60)} ไม่มีเลขใบ (รุ่นก่อน) ไม่รับ`);
      const n = p[1], g = p[2];
      const f = (await sp("GET", `/_api/web/GetFileById('${g}')?$select=Name,ServerRelativeUrl,Length,TimeCreated`)).body || {};
      const url = String(f.ServerRelativeUrl || "");
      const at = async (meta) => { const base = (await rootOf(meta)) + "/Attachments/"; return url.startsWith(base) ? url.slice(base.length).match(/^(\d+)\/([^/]+)$/) : null; };
      const r = await at(E.dr_datarequest);
      if (r && r[1] === n) return { n, g, f, meta: E.dr_datarequest, item: r[1], stored: r[2], onWork: false };
      const w = await at(E.dr_datarequest.work);
      if (w && (await fetchWorkFor(n)).some((x) => Number(x.dr_requestid) === Number(n) && verOf(x) > 1 && String(x.Id) === w[1])) return { n, g, f, meta: E.dr_datarequest.work, item: w[1], stored: w[2], onWork: true };
      throw httpErr(403, "[DRS] ไฟล์นี้ไม่ใช่ไฟล์แนบของใบนี้ (403)", "DRS_FORBIDDEN");
    }
    function annotationOf(n, info, f, sel) {
      const full = { annotationid: annId(n, f.UniqueId), filename: info.name, mimetype: mimeOf(info.name), filesize: Number(f.Length || 0),
        subject: info.subject, createdon: f.TimeCreated || null, isdocument: true };
      if (!sel || !sel.length) return full;
      const o = {};
      for (const k of sel) { if (!(k in full)) unsupported("annotation", `คอลัมน์ ${k}`); o[k] = full[k]; }
      return o;
    }
    async function listAnnotations(q) {
      const f = String(q.$filter || "").match(/^_objectid_value eq \{?(\d+)\}?(?: and isdocument eq true)?$/);
      if (!f) unsupported("annotation", `query ${JSON.stringify(q)}`);
      for (const k of Object.keys(q)) if (!["$select", "$filter", "$orderby"].includes(k)) unsupported("annotation", `ตัวเลือก ${k}`);
      if (q.$orderby && !/^createdon( asc)?$/i.test(q.$orderby.trim())) unsupported("annotation", `$orderby ${q.$orderby}`);
      const n = f[1], sel = q.$select ? q.$select.split(",").map((x) => x.trim()).filter(Boolean) : [];
      const out = [];
      for (const x of await folderFiles(`${await rootOf(E.dr_datarequest)}/Attachments/${n}`)) out.push({ f: x, info: fileInfo(x.Name, false) });
      // ไฟล์ทีม: แถวงานที่เชื่อของใบนี้เท่านั้น (ผู้แจ้งเห็นเฉพาะแถวที่ตัวเองสร้าง)
      const wroot = await rootOf(E.dr_datarequest.work);
      for (const w of (await fetchWorkFor(n)).filter((w) => verOf(w) > 1).sort(byId)) for (const x of await folderFiles(`${wroot}/Attachments/${w.Id}`)) out.push({ f: x, info: fileInfo(x.Name, true) });
      out.sort((a, b) => String(a.f.TimeCreated || "").localeCompare(String(b.f.TimeCreated || "")) || String(a.info.name).localeCompare(String(b.info.name)));
      return { entities: out.map((x) => annotationOf(n, x.info, x.f, sel)), nextLink: undefined };
    }
    async function getAnnotation(id, q) {
      const x = await fileById(id);
      const info = fileInfo(x.stored, x.onWork);
      const sel = q.$select ? q.$select.split(",").map((k) => k.trim()).filter(Boolean) : [];
      const o = annotationOf(x.n, info, x.f, sel.filter((k) => k !== "documentbody"));
      if (!sel.length || sel.includes("documentbody")) o.documentbody = bytesToB64(await spRaw("GET", `/_api/web/GetFileById('${x.g}')/$value`));
      return o;
    }
    // ชื่อไฟล์ที่ SharePoint รับ (ตัดอักขระต้องห้าม) และไม่ชนไฟล์เดิมในแถวเดียวกัน (ชน = ต่อท้าย (2) (3) ...)
    function safeName(name) {
      let s = String(name || "file").replace(/["*:<>?\/\\|\u0000-\u001f]/g, "_").replace(/^[\s.]+|[\s.]+$/g, "");
      return s || "file";
    }
    function uniqueName(name, taken) {
      const low = new Set(taken.map((t) => String(t).toLowerCase()));
      if (!low.has(name.toLowerCase())) return name;
      const m = name.match(/^(.*?)(\.[^.]*)?$/), base = m[1], ext = m[2] || "";
      for (let i = 2; i < 1000; i++) { const c = `${base} (${i})${ext}`; if (!low.has(c.toLowerCase())) return c; }
      throw httpErr(409, "[DRS] ชื่อไฟล์ซ้ำมากเกินไป");
    }
    async function createAnnotation(data) {
      const d = data || {};
      const bind = String(d["objectid_dr_datarequest@odata.bind"] || "").match(/^\/?dr_datarequests\(\{?(\d+)\}?\)$/);
      if (!bind) unsupported("annotation", `สร้างโดยไม่ผูกใบงาน ${JSON.stringify(Object.keys(d))}`);
      for (const k of Object.keys(d)) if (!["subject", "filename", "mimetype", "documentbody", "isdocument", "notetext", "objectid_dr_datarequest@odata.bind"].includes(k)) unsupported("annotation", `สร้างพร้อมคอลัมน์ ${k}`);
      const n = bind[1], delivery = d.subject === "delivery", team = isTeam();
      if (delivery && !team) deny("ผู้แจ้งส่งไฟล์ส่งมอบไม่ได้");
      const bytes = b64ToBytes(d.documentbody);
      const clean = safeName(d.filename);
      let meta, item, stored;
      if (team) {
        // แถวงาน: เขียนใหม่ทั้งแถวก่อนถ้ายังเวอร์ชัน 1.0 เพราะการแนบไฟล์ทำให้เวอร์ชันขึ้นเอง (probe D) ค่าที่ผู้แจ้งใส่ตอนสร้างจะกลายเป็นค่าที่เชื่อ
        const w = await teamWork(n, () => ({}), null);
        meta = E.dr_datarequest.work; item = w.wid;
        stored = delivery ? DELIV + clean : clean.startsWith(DELIV) || clean.startsWith(TEAMG) ? TEAMG + clean : clean;
      } else { meta = E.dr_datarequest; item = n; stored = clean; }
      const folder = `${await rootOf(meta)}/Attachments/${item}`;
      stored = uniqueName(stored, (await folderFiles(folder)).map((x) => x.Name));
      await spRaw("POST", `${listPath(meta)}/items(${item})/AttachmentFiles/addUsingPath(DecodedUrl=@a)?${alias(stored)}`, bytes);
      const got = await sp("GET", `/_api/web/GetFileByServerRelativePath(DecodedUrl=@a)?$select=UniqueId&${alias(folder + "/" + stored)}`);
      return { entityType: "annotation", id: annId(n, got.body.UniqueId) };
    }
    async function deleteAnnotation(id) {
      const x = await fileById(id);
      await sp("POST", `${listPath(x.meta)}/items(${x.item})/AttachmentFiles/getByFileName(@a)?${alias(x.stored)}`, undefined, { "X-HTTP-Method": "DELETE" });
      return { entityType: "annotation", id: annId(x.n, x.g) };
    }

    const xrm = {
      __shim: true, __drs: state, ready,
      Utility: { getGlobalContext: () => ctx },
      WebApi: {
        retrieveMultipleRecords, retrieveRecord, createRecord, updateRecord, deleteRecord,
        online: { execute },
      },
      _update: update,
    };
    return xrm;
  }

  // ---------- ตัวดัก fetch: คำขอ /api/data/v9.x/ ของหน้า DR (patchDataRequestRaw , dvFetch) ไม่ออกเน็ต ----------
  function installFetchBridge(win, xrm) {
    const orig = win.fetch.bind(win);
    const R = win.Response;
    const json = (status, obj, headers) => new R(obj === null ? null : JSON.stringify(obj), { status, headers: Object.assign({ "Content-Type": "application/json" }, headers || {}) });
    win.fetch = async function (input, init) {
      const url = typeof input === "string" ? input : input && input.url ? input.url : String(input);
      const m = url.match(/\/api\/data\/v9\.\d+\/(.*)$/);
      if (!m) return orig(input, init);
      init = init || {};
      const method = String(init.method || "GET").toUpperCase();
      const hdr = init.headers || {};
      const get = (k) => (hdr.get ? hdr.get(k) : hdr[k] !== undefined ? hdr[k] : hdr[Object.keys(hdr).find((x) => x.toLowerCase() === k.toLowerCase())]);
      const pm = m[1].match(/^dr_datarequests\(([^)]+)\)$/);
      if (method === "PATCH" && pm) {
        try {
          const r = await xrm._update("dr_datarequest", pm[1], JSON.parse(init.body || "{}"), get("If-Match") || "*");
          return new R(null, { status: 204, headers: r.etag ? { ETag: r.etag } : {} });
        } catch (e) {
          // 412 (etag เก่า) ส่งรหัสเดิมกลับ หน้าแปลงเป็น ConcurrencyConflictError เอง
          return json(e.status && e.status >= 400 ? e.status : 500, { error: { code: e.code || "", message: e.message } });
        }
      }
      const msg = `${TAG} fetch ${method} /api/data/${m[1].split("?")[0]}`;
      xrm.__drs.unsupported.push(msg);
      try { win.console.warn(msg); } catch (e) {}
      return json(501, { error: { code: "DRS_UNSUPPORTED", message: msg } });
    };
    return win.fetch;
  }

  // ---------- ตัวกระจายตัวจัดการ on* (หน้า SharePoint บล็อก onclick ในแท็ก ตัวสร้างเปลี่ยนชื่อเป็น data-drs-on<event>) ----------
  const NON_BUBBLE = ["mouseenter", "mouseleave", "focus", "blur", "load", "error", "scroll"];
  const BUBBLE = ["click", "dblclick", "change", "input", "keydown", "keyup", "keypress", "submit", "mousedown", "mouseup", "contextmenu", "focusin", "focusout", "paste", "dragstart", "drop", "dragover"];
  // รันเฉพาะข้อความที่ตัวสร้างจดไว้ใน drs_handlers.js (window.DRS_HANDLERS) ข้อความอื่นไม่รัน จดไว้ใน blocked() และเตือนครั้งเดียวต่อข้อความ
  // ไม่มีรายการ = ไม่รันอะไรเลย (ปิดไว้ก่อน) hardening 11/11/2026: เดิมรัน data-drs-on* จาก HTML ใดก็ได้ HTML ที่หลุดเข้าหน้าจึงกลายเป็นโค้ด
  function installInlineHandlers(doc, win, registry) {
    const reg = registry === undefined ? win.DRS_HANDLERS || null : registry;
    const exact = {}, pats = {}, verdict = new Map(), blocked = [];
    if (reg) {
      for (const [ev, list] of Object.entries(reg.exact || {})) exact[ev] = new Set(list);
      for (const [ev, list] of Object.entries(reg.patterns || {})) pats[ev] = list.map((src) => new RegExp("^(?:" + src + ")$"));
    }
    const allowed = (ev, code) => {
      const key = ev + "\u0000" + code;
      if (verdict.has(key)) return verdict.get(key);
      const ok = !!reg && ((exact[ev] && exact[ev].has(code)) || (pats[ev] || []).some((re) => re.test(code)));
      verdict.set(key, ok);
      if (!ok) {
        blocked.push(`${ev}: ${code.slice(0, 160)}`);
        try { (win.console || console).warn(`[DRS ตัวจัดการไม่อยู่ในรายการ ไม่รัน] ${ev}: ${code.slice(0, 160)}`); } catch (e) {}
      }
      return ok;
    };
    const cache = new Map();
    const compile = (code) => {
      let fn = cache.get(code);
      if (!fn) { fn = new win.Function("event", "with(this.ownerDocument||document){with((this&&this.form)||{}){with(this){\n" + code + "\n}}}"); cache.set(code, fn); }
      return fn;
    };
    const report = (e) => { if (win.reportError) win.reportError(e); else setTimeout(() => { throw e; }, 0); };
    const run = (el, ev, attr) => {
      const code = el.getAttribute(attr);
      if (code === null) return;
      if (!allowed(attr.slice(11), code)) return; // "data-drs-on".length = 11
      try { const r = compile(code).call(el, ev); if (r === false) ev.preventDefault(); } catch (e) { report(e); }
    };
    for (const t of BUBBLE) {
      const attr = "data-drs-on" + t;
      doc.addEventListener(t, (ev) => {
        const path = ev.composedPath ? ev.composedPath() : [];
        for (const n of path) {
          if (n === doc || n === win) break;
          if (n.nodeType === 1 && n.hasAttribute(attr)) { run(n, ev, attr); if (ev.cancelBubble) break; }
        }
      }, false);
    }
    for (const t of NON_BUBBLE) {
      const attr = "data-drs-on" + t;
      doc.addEventListener(t, (ev) => { const n = ev.target; if (n && n.nodeType === 1 && n.hasAttribute(attr)) run(n, ev, attr); }, true);
    }
    return { bubble: BUBBLE.slice(), nonBubble: NON_BUBBLE.slice(), blocked: () => blocked.slice(), hasRegistry: !!reg };
  }

  const api = { makeSpXrm, installFetchBridge, installInlineHandlers, parseQuery, parseExpand, TAG, mergeRequest, pickRead, pickTarget, parseAssignees, mergeChat, splitChat };
  if (typeof module === "object" && module.exports) module.exports = api;
  else {
    root.DRS = root.DRS || {}; Object.assign(root.DRS, api);
    // ตั้งเองในเบราว์เซอร์เมื่อหน้าใส่ data-drs-site ไว้ (สำเนาหน้า DR จาก build_drs_page.mjs)
    const d = root.document, site = d && d.documentElement && d.documentElement.getAttribute("data-drs-site");
    if (site) {
      root.DRS_DISPATCH = installInlineHandlers(d, root); // รายการจาก drs_handlers.js ที่โหลดก่อนไฟล์นี้ (window.DRS_HANDLERS)
      // ทาง J (13/11/2026) หน้าเปิดจากนอก SharePoint: drs_spa_auth.js ตั้ง DRS_AUTH_FETCH ที่แนบ token ของผู้ใช้ไว้ก่อน
      root.DRS_XRM = makeSpXrm({ siteUrl: site, fetch: typeof root.DRS_AUTH_FETCH === "function" ? root.DRS_AUTH_FETCH : root.fetch.bind(root), prefix: d.documentElement.getAttribute("data-drs-prefix") || undefined });
      installFetchBridge(root, root.DRS_XRM);
    }
  }
})(typeof globalThis !== "undefined" ? globalThis : this);
