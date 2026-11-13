// drs_schema.js  สเปก DR รุ่น SharePoint ขั้นที่ 1 ที่เดียว (ตัวติดตั้ง drs_provision.js กับตัวแทน drs_xrm_sp.js ใช้ร่วมกัน ชื่อคอลัมน์จึงไม่เพี้ยนกัน)
// ไฟล์ใหม่ ไม่แตะ DR เดิม  ใช้ได้ทั้ง node (require) และเบราว์เซอร์ (script ธรรมดา ได้ window.DRS.schema)
// ชื่อคอลัมน์ใน SharePoint ตั้งตรงกับชื่อ logical ของ Dataverse (dr_status ฯลฯ) ยกเว้นชื่อที่ SharePoint มีอยู่แล้ว (Title, Created, Modified, Author, Id)
// ขั้นที่ 1 ตามรายงาน .claude/research/2026-11-11-dr-on-sharepoint-feasibility.md ข้อ 6.3: ใบงาน ตั้งค่า สมาชิกทีม
// ขั้น 2 ผู้รับผิดชอบ (คอลัมน์ dr_assignees ใน DRS Work) กับขั้น 3 ไฟล์แนบ (ไฟล์แนบของแถว list ไม่มีคอลัมน์ใหม่) ทำแล้ว 11/11/2026
// ขั้น 4 บันทึกภายในทีม (list DRS Internal Notes) , ขั้นต่อไป (ยังไม่ทำ): โปรไฟล์
(function (root) {
  "use strict";
  const PREFIX = "DRS ";

  function names(prefix) {
    prefix = prefix === undefined ? PREFIX : prefix;
    return {
      prefix,
      lists: { requests: prefix + "Requests", settings: prefix + "Settings", members: prefix + "Team Members", work: prefix + "Work", notes: prefix + "Internal Notes" },
      groups: { requester: prefix + "Requesters", team: prefix + "Team", admin: prefix + "Admins" },
      roles: { requester: prefix + "Requester no delete", team: prefix + "Team no delete", admin: prefix + "Admin Contribute", requesterAdd: prefix + "Requester add only" },
    };
  }

  // 0x88 = ลบแถว (0x8) กับลบเวอร์ชัน (0x80) , 0x100 = Override List Behaviors ข้ามกติกาอ่านแก้เฉพาะแถวของตัวเอง (โน้ต Proven ท่าเดียวกับ LE)
  const ROLE_SPECS = Object.freeze({
    requester: { from: "Contribute", removeMask: "0x88", addMask: "0", description: "DRS ผู้แจ้ง: สร้างและแก้ใบของตัวเอง ลบไม่ได้" },
    team: { from: "Contribute", removeMask: "0x88", addMask: "0x100", description: "DRS ทีม Data: เห็นและแก้ทุกใบ ลบไม่ได้" },
    admin: { from: "Contribute", removeMask: "0", addMask: "0x100", description: "DRS แอดมิน: เห็นแก้ลบทุกใบ" },
    // DRS Work: ผู้แจ้งสร้างแถวงานของใบตัวเองได้ (เป็นผู้สร้าง จึงอ่านแถวนั้นได้ใน list อ่านเฉพาะของตัวเอง) แต่แก้และลบไม่ได้ 0x2 = AddListItems
    requesterAdd: { from: "Read", removeMask: "0", addMask: "0x2", description: "DRS ผู้แจ้งบน DRS Work: อ่านและเพิ่มแถว แก้ลบไม่ได้" },
  });

  // ชื่อ role ของ Dataverse ที่หน้า DR จับด้วยชื่อ (drRoleKindOf ใน dashboard.html) กลุ่ม SharePoint แต่ละกลุ่มสังเคราะห์เป็น role ชื่อนี้
  const ROLE_NAME_OF_GROUP = Object.freeze({
    requester: "dr_Data Request - Requester",
    team: "dr_Data Request - Team Data",
    admin: "dr_Data Request - Team Admin",
  });

  // [logical, ชื่อใน SharePoint, ชนิด, index, ห้ามซ้ำ]
  const REQUEST_COLUMNS = [
    ["dr_title", "Title", "Text", false, false],
    ["dr_ticket", "dr_ticket", "Text", true, false],
    // GUID ที่หน้าจอสร้างเองตอนแจ้ง (กันสร้างซ้ำตอนลองใหม่ R2G-5) SharePoint ตั้ง Id เองไม่ได้
    // ‼️ ห้ามตั้งห้ามซ้ำ (EnforceUniqueValues): list ที่ผู้ใช้เห็นเฉพาะแถวตัวเอง (ReadSecurity 2) มีคอลัมน์ห้ามซ้ำไม่ได้ ได้ 500
    //    "A list for which users can only view their own items cannot have fields that enforce unique values." (probe_rs500_run1.log 11/11/2026)
    //    กันซ้ำด้วยด่านของหน้าเอง: ลองใหม่ = retrieveRecord(GUID) ก่อนสร้าง ตัวแทนค้นจาก dr_clientkey ที่มี index
    ["dr_clientkey", "dr_clientkey", "Text", true, false],
    ["dr_status", "dr_status", "Number", true, false],
    ["dr_priority", "dr_priority", "Number", false, false],
    ["dr_requesttype", "dr_requesttype", "Number", false, false],
    ["dr_region", "dr_region", "Number", false, false],
    ["dr_deliveryfeedback", "dr_deliveryfeedback", "Number", false, false],
    ["dr_sladue", "dr_sladue", "DateTime", false, false],
    ["dr_registeredon", "dr_registeredon", "DateTime", false, false],
    ["dr_closedon", "dr_closedon", "DateTime", false, false],
    ["dr_deliveredon", "dr_deliveredon", "DateTime", false, false],
    ["dr_requesteremail", "dr_requesteremail", "Text", true, false],
    ["dr_decisionuse", "dr_decisionuse", "Note", false, false],
    ["dr_details", "dr_details", "Note", false, false],
    ["dr_deliverylink", "dr_deliverylink", "Note", false, false],
    ["dr_teamnote", "dr_teamnote", "Note", false, false],
    ["dr_requesterlinks", "dr_requesterlinks", "Note", false, false],
    ["dr_conversation", "dr_conversation", "Note", false, false],
    ["dr_checklist", "dr_checklist", "Note", false, false],
    ["dr_activity", "dr_activity", "Note", false, false],
    ["dr_feedbacknote", "dr_feedbacknote", "Note", false, false],
    // คำขอของผู้แจ้ง (hardening 11/11/2026): ผู้แจ้งเขียนสถานะจริงไม่ได้แล้ว กดยกเลิกหรือปิดงานลงช่องนี้ ตัวแทนตัดสินตอนอ่านตามกติกาเดียวกับจอ
    ["dr_reqstatus", "dr_reqstatus", "Number", false, false],
    ["dr_reqstatuson", "dr_reqstatuson", "DateTime", false, false],
    // ประวัติที่ผู้แจ้งเขียน (ชื่อผู้ทำถูกแทนด้วยผู้สร้างใบตอนอ่าน) ประวัติของทีมอยู่ DRS Work
    ["dr_reqactivity", "dr_reqactivity", "Note", false, false],
  ];
  // ช่องของทีม (hardening 11/11/2026 จากผลพี่ปริม: ผู้แจ้งแก้ 9 ช่องของทีมในใบตัวเองได้ 204 เพราะสิทธิ์ list กันระดับแถว ไม่กันคอลัมน์)
  // หนึ่งแถวต่อใบงาน คีย์ dr_requestid = Id ของใบใน DRS Requests ห้ามซ้ำไม่ได้ (list อ่านเฉพาะของตัวเอง) ตัวแทนเลือกแถวตามกติกาใน drs_xrm_sp.js
  // ขั้น 4 บันทึกภายในทีม (รุ่น Dataverse กันด้วย FSP "DR Team Internal" ผู้แจ้งอ่านไม่ได้แม้ใบตัวเอง) list แยกที่ผู้แจ้งไม่มีสิทธิ์เลย
  // ห้ามเก็บใน Requests หรือ Work: ผู้แจ้งอ่านแถวของใบตัวเองได้ทั้งสองที่ (probe_internal_run2.log ผู้แจ้งได้ 404 403 ทุกทางเมื่อเป็น list แยก)
  const NOTE_COLUMNS = [
    ["dr_notename", "Title", "Text", false, false],
    ["dr_requestid", "dr_requestid", "Number", true, false],
    ["dr_internalnote", "dr_internalnote", "Note", false, false],
  ];

  const WORK_COLUMNS = [
    ["dr_workname", "Title", "Text", false, false],
    ["dr_requestid", "dr_requestid", "Number", true, false],
    ["dr_status", "dr_status", "Number", true, false],
    ["dr_statuson", "dr_statuson", "DateTime", false, false],
    ["dr_priority", "dr_priority", "Number", false, false],
    ["dr_sladue", "dr_sladue", "DateTime", false, false],
    ["dr_teamnote", "dr_teamnote", "Note", false, false],
    ["dr_deliverylink", "dr_deliverylink", "Note", false, false],
    ["dr_checklist", "dr_checklist", "Note", false, false],
    ["dr_activity", "dr_activity", "Note", false, false],
    ["dr_registeredon", "dr_registeredon", "DateTime", false, false],
    ["dr_closedon", "dr_closedon", "DateTime", false, false],
    ["dr_deliveredon", "dr_deliveredon", "DateTime", false, false],
    // ขั้น 2 (11/11/2026): ผู้รับผิดชอบหลายคน JSON [{"id":"<Id ใน DRS Team Members>","name":"<ชื่อ>"}] แทน N:N dr_DataRequest_Assignees
    // อยู่แถวงานเพราะผู้แจ้งอ่านแถวงานของใบตัวเองได้แต่แก้ไม่ได้ ทีมแก้ได้ ค่าเชื่อเมื่อเวอร์ชันเกิน 1.0 เหมือนช่องอื่นของทีม
    ["dr_assignees", "dr_assignees", "Note", false, false],
    // แชทฝั่งทีม (11/11/2026 จากพี่ปริม T7 T7d: ผู้แจ้งเขียนข้อความในนามทีมและแก้ข้อความทีมใน dr_conversation ของใบตัวเองได้ 204)
    // JSON lines แบบเดียวกับ dr_conversation เฉพาะบรรทัด by team ทีมเขียนฝ่ายเดียว ตัวแทนรวมกับฝั่งผู้แจ้งตอนอ่าน
    ["dr_teamconv", "dr_teamconv", "Note", false, false],
  ];
  const SETTING_COLUMNS = [
    ["dr_name", "Title", "Text", false, false],
    ["dr_settingkey", "dr_settingkey", "Text", true, false],
    ["dr_settingvalue", "dr_settingvalue", "Note", false, false],
  ];
  const MEMBER_COLUMNS = [
    ["dr_name", "Title", "Text", false, false],
    ["dr_email", "dr_email", "Text", true, false],
    ["dr_active", "dr_active", "Boolean", false, false],
    ["dr_order", "dr_order", "Number", false, false],
    ["dr_isadmin", "dr_isadmin", "Boolean", false, false],
    ["dr_avatar", "dr_avatar", "Note", false, false],
  ];

  // คอลัมน์ระบบที่ทุก entity มี (Dataverse -> SharePoint)
  const SYSTEM = Object.freeze({ createdon: "Created", modifiedon: "Modified", _ownerid_value: "AuthorId", _createdby_value: "AuthorId" });

  // entity ที่ขั้นที่ 1 รองรับ  later = คอลัมน์ของขั้นถัดไป (อ่านได้ค่าว่างพร้อมเตือน เขียนแล้ว throw)  legacyNull = คอลัมน์เก่าที่ไม่ย้าย อ่านได้ค่าว่างเงียบ ๆ
  function entities(prefix) {
    const N = names(prefix);
    return {
      dr_datarequest: { list: N.lists.requests, idKey: "dr_datarequestid", clientKey: "dr_clientkey", columns: REQUEST_COLUMNS, later: {}, legacyNull: ["_dr_assignee_value"], ticket: true, work: { list: N.lists.work, columns: WORK_COLUMNS }, notes: { list: N.lists.notes, columns: NOTE_COLUMNS } },
      dr_setting: { list: N.lists.settings, idKey: "dr_settingid", columns: SETTING_COLUMNS, later: {}, legacyNull: [] },
      dr_teammember: { list: N.lists.members, idKey: "dr_teammemberid", columns: MEMBER_COLUMNS, later: {}, legacyNull: ["dr_activerequestcount"] },
    };
  }

  // แผนสิทธิ์ราย list: ws/rs 2 = เขียน/อ่านเฉพาะของตัวเอง (คนที่มี Override ข้ามได้)
  // Requests: ผู้แจ้งเห็นแก้เฉพาะใบตัวเอง ทีมเห็นแก้ทุกใบ ลบไม่ได้ แอดมินลบได้ (brief ข้อ a ถึง c)
  // Settings: ทุกคนอ่านได้ ทีมและแอดมินเขียนได้ (Dataverse เดิม dr_setting Create Write = Team Data) ทีมลบไม่ได้
  // Team Members: ทุกคนอ่านได้ แอดมินเท่านั้นเขียนลบ (Dataverse เดิม Team Admin)
  function listPlan(prefix) {
    const N = names(prefix);
    return [
      { key: "requests", title: N.lists.requests, columns: REQUEST_COLUMNS, versioning: true, ws: 2, rs: 2, grants: { requester: "role:requester", team: "role:team", admin: "role:admin" } },
      { key: "settings", title: N.lists.settings, columns: SETTING_COLUMNS, versioning: true, ws: 1, rs: 1, grants: { requester: "Read", team: "role:team", admin: "role:admin" } },
      { key: "members", title: N.lists.members, columns: MEMBER_COLUMNS, versioning: true, ws: 1, rs: 1, grants: { requester: "Read", team: "Read", admin: "role:admin" } },
      // Work: ผู้แจ้งอ่านเฉพาะแถวที่ตัวเองสร้าง สร้างได้ แก้ลบไม่ได้ ทีมเห็นแก้ทุกแถว (Override) ลบไม่ได้ แอดมินลบได้ เวอร์ชันเปิด (ตัวแทนใช้เลขเวอร์ชันตัดสินว่าทีมแตะแถวแล้ว)
      // noCrawl (ฟ้าเคาะ 11/11/2026): คนที่ถูกถอดจาก DRS Team ยังค้นหาเห็นบันทึกทีมกับลิงก์ส่งงานของทุกใบ 7 ถึง 20 นาที (fa_search_run2.log 7/7 , s3_attack.log ของพี่ปริม)
      // หน้า DR ไม่ใช้ /_api/search ปิดแล้วไม่กระทบแอป
      { key: "work", title: N.lists.work, columns: WORK_COLUMNS, versioning: true, ws: 2, rs: 2, noCrawl: true, grants: { requester: "role:requesterAdd", team: "role:team", admin: "role:admin" } },
      // Internal Notes: ทีม แอดมิน อ่านเขียนทุกแถว ลบได้เฉพาะแอดมิน ผู้แจ้งไม่มีสิทธิ์เลย (ตัวติดตั้งถอนกลุ่ม DRS ที่ไม่อยู่ใน grants ออก)
      // noCrawl: ไม่ขึ้นผลค้นหา คนที่เพิ่งออกจากทีมยังค้นเจอเนื้อบันทึกได้เกิน 12 นาที (แคชกลุ่มของการค้นหา probe_search_wait_B.log)
      { key: "notes", title: N.lists.notes, columns: NOTE_COLUMNS, versioning: true, ws: 1, rs: 1, noCrawl: true, grants: { team: "role:team", admin: "role:admin" } },
    ];
  }

  // การแบ่งช่องของใบงานระหว่าง DRS Requests กับ DRS Work (ตัวแทนใช้ทั้งตอนอ่านรวมและตอนแยกเขียน) ค่าสถานะตรงกับ STATUS_VALUE ใน dashboard.html
  const STATUS = Object.freeze({ pending: 121900000, queue: 121900001, doing: 121900002, waitinfo: 121900003, delivered: 121900004, closed: 121900005, cancelled: 121900006 });
  const SPLIT = Object.freeze({
    teamOnly: ["dr_status", "dr_teamnote", "dr_deliverylink", "dr_checklist", "dr_registeredon", "dr_closedon", "dr_deliveredon"], // อยู่ DRS Work เท่านั้น
    shared: ["dr_priority", "dr_sladue"],        // ผู้แจ้งเสนอใน Requests ช่วงที่ยังแก้ใบได้ ทีมคัดไปเก็บใน Work ตอนสถานะออกจากช่วงนั้น
    gated: ["dr_deliveryfeedback", "dr_feedbacknote"], // ของผู้แจ้ง โชว์เฉพาะสถานะจริงส่งมอบแล้วหรือปิดงาน
    activity: "dr_activity",                      // ทีมเขียนใน Work ผู้แจ้งเขียน dr_reqactivity รวมตอนอ่าน
    chat: "dr_conversation",
    internal: "dr_internalnote",                  // บันทึกภายในทีม อยู่ list DRS Internal Notes ทีม แอดมิน อ่านเขียน ผู้แจ้งได้ null                      // ฝั่งผู้แจ้งอยู่ Requests ฝั่งทีมอยู่ Work.dr_teamconv รวมตอนอ่าน (แยกตอนเขียน)
    derived: ["dr_ticket"],                       // สร้างจาก Id ตอนอ่าน
    intents: ["dr_reqstatus", "dr_reqstatuson", "dr_reqactivity"],
    selfEditable: [121900000, 121900001, 121900003], // SELF_EDITABLE_STATUSES ของหน้า (รอรับงาน รอคิว รอข้อมูล)
    deliveryDone: [121900004, 121900005],            // DELIVERY_DONE_STATUSES
    newStatus: 121900001,                             // NEW_TICKET_STATUS รอคิว
  });

  const api = { PREFIX, names, ROLE_SPECS, ROLE_NAME_OF_GROUP, REQUEST_COLUMNS, SETTING_COLUMNS, MEMBER_COLUMNS, WORK_COLUMNS, NOTE_COLUMNS, SYSTEM, STATUS, SPLIT, entities, listPlan };
  if (typeof module === "object" && module.exports) module.exports = api;
  else { root.DRS = root.DRS || {}; root.DRS.schema = api; }
})(typeof globalThis !== "undefined" ? globalThis : this);
