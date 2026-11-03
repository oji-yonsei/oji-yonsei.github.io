/**
 * 오지 탐험대 자료실 업로드 서버 (Google Apps Script 웹 앱)
 * - 대원이 사이트에서 올린 파일을 대표 계정 드라이브의 "오지 탐험대 자료실/학기" 폴더에 저장합니다.
 * - 사이트 자료실 목록은 이 폴더들에서 "링크가 있는 모든 사용자" 공유가 켜진 파일을 읽어 보여줍니다.
 *   (드라이브에서 파일 공유를 끄거나 파일을 지우면 사이트 목록에서도 사라집니다.)
 * - 업로드·수정·삭제에는 대원 코드가 필요합니다. 삭제한 파일은 드라이브 휴지통으로 옮겨집니다(30일 안에 복구 가능).
 * - 대원 코드 설정: 프로젝트 설정 → 스크립트 속성에 UPLOAD_CODE 를 추가하세요.
 * - 공지사항: 대표 드라이브 "오지 탐험대 자료실" 폴더의 "오지 탐험대 공지사항" 시트에 저장됩니다(처음 호출 때 자동 생성).
 *   사이트에서 공지를 쓰고 고치고 지우려면 스크립트 속성에 ADMIN_CODE(운영진 코드)를 추가하세요.
 *   시트에서 직접 고쳐도 사이트에 반영됩니다. 아래쪽 행일수록 최신 공지로 보입니다.
 */
const ROOT_ID = "1GbHabez0EURsNq7VChHDylZ7f4aSlj33"; // 오지 탐험대 자료실 폴더
const KINDS = ["활동 보고서", "활동 사진", "카드뉴스", "발제 자료", "기타"];
const EXTS = ["pdf","jpg","jpeg","png","gif","webp","heic","hwp","hwpx","doc","docx","ppt","pptx","xls","xlsx","mp4","mov","txt","csv"];
const MAX_BYTES = 10 * 1024 * 1024; // 10MB
const NOTICE_CATS = ["소식", "모임", "활동", "독서", "모집", "행사", "자료실", "기타"];
const NOTICE_HEAD = ["id", "날짜", "분류", "제목", "내용", "고정", "링크"];
const NOTICE_SEED = [
  ["2026.03.26", "소식", "오지 탐험대 캐릭터 ‘지금이’를 공개했습니다", "탐험모자를 쓰고 길을 나서는 오지 탐험대의 캐릭터, 지금이를 소개합니다. 모임 소개 페이지에서 지금이를 만나 보세요.", "", ""],
  ["2026 2학기", "소식", "IPAID 산하 학술모임으로 등록했습니다", "오지 탐험대는 연세대학교 빈곤문제국제개발연구원(IPAID) 산하 학술모임으로 등록되어 활동합니다.", "", "https://ipaid.yonsei.ac.kr/ipaid/index.do"],
  ["2026 2학기", "모임", "2학기 정기 모임은 매주 화요일 오후에 진행됩니다", "2학기 정기 모임은 매주 화요일 오후, 약 1시간 내외로 진행됩니다.", "Y", ""],
  ["2026 2학기", "활동", "10주 포트폴리오 과정이 진행 중입니다", "사실 → 해석 → 표현의 3단계로, 경험 정리부터 이력서·자기소개서 초안까지 함께 만듭니다.", "", ""],
  ["2026 2학기", "독서", "『왜 세계의 절반은 굶주리는가?』 주차별 발제가 진행 중입니다", "매주 목차를 나눠 읽고 발제·영상·기사를 함께 봅니다. 활동 페이지에서 주차별 진도를 확인할 수 있어요.", "", ""],
  ["2026.11.03", "자료실", "자료실이 열렸습니다", "활동 보고서, 사진, 카드뉴스를 자료실에서 보고 직접 올릴 수 있어요. 업로드에는 대원 코드가 필요합니다.", "", ""]
];

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function semLabel_(sem) { // "2026-1" -> "2026년 1학기"
  const m = /^(\d{4})-([12])$/.exec(sem);
  return m ? m[1] + "년 " + m[2] + "학기" : null;
}

function semFolder_(sem) {
  const label = semLabel_(sem);
  const root = DriveApp.getFolderById(ROOT_ID);
  const it = root.getFoldersByName(label);
  return it.hasNext() ? it.next() : root.createFolder(label);
}

function doGet() {
  const items = [];
  const folders = DriveApp.getFolderById(ROOT_ID).getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    const m = /^(\d{4})년 ([12])학기$/.exec(folder.getName());
    if (!m) continue;
    const sem = m[1] + "-" + m[2];
    const files = folder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.getSharingAccess() !== DriveApp.Access.ANYONE_WITH_LINK) continue;
      let meta = {};
      try { meta = JSON.parse(f.getDescription() || "{}"); } catch (e) {}
      items.push({
        id: f.getId(),
        title: meta.title || f.getName(),
        kind: KINDS.indexOf(meta.kind) >= 0 ? meta.kind : "기타",
        semester: sem,
        date: Utilities.formatDate(f.getDateCreated(), "Asia/Seoul", "yyyy-MM-dd"),
        url: f.getUrl(),
        mime: f.getMimeType(),
        size: f.getSize(),
        uploader: meta.uploader || ""
      });
    }
  }
  items.sort(function (a, b) { return b.date.localeCompare(a.date); });
  let notices = null;
  try { notices = readNotices_(); } catch (e) {}
  return json_({ ok: true, items: items, notices: notices });
}

function noticeSheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("NOTICE_SHEET_ID");
  if (id) { try { return SpreadsheetApp.openById(id).getSheets()[0]; } catch (e) {} }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const again = props.getProperty("NOTICE_SHEET_ID");
    if (again && again !== id) return SpreadsheetApp.openById(again).getSheets()[0];
    const ss = SpreadsheetApp.create("오지 탐험대 공지사항");
    DriveApp.getFileById(ss.getId()).moveTo(DriveApp.getFolderById(ROOT_ID));
    const sh = ss.getSheets()[0];
    sh.setName("공지");
    sh.getRange("A:G").setNumberFormat("@");
    sh.getRange(1, 1, 1, NOTICE_HEAD.length).setValues([NOTICE_HEAD]).setFontWeight("bold");
    sh.setFrozenRows(1);
    const rows = NOTICE_SEED.map(function (r) { return [Utilities.getUuid().slice(0, 8)].concat(r); });
    sh.getRange(2, 1, rows.length, NOTICE_HEAD.length).setValues(rows);
    sh.setColumnWidth(4, 320); sh.setColumnWidth(5, 420);
    props.setProperty("NOTICE_SHEET_ID", ss.getId());
    return sh;
  } finally { lock.releaseLock(); }
}

function cell_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, "Asia/Seoul", "yyyy.MM.dd");
  return String(v == null ? "" : v).trim();
}

function readNotices_() {
  const sh = noticeSheet_();
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const title = cell_(r[3]);
    if (!title) continue;
    let id = cell_(r[0]);
    if (!id) { id = Utilities.getUuid().slice(0, 8); sh.getRange(i + 1, 1).setValue(id); }
    const pin = cell_(r[5]).toUpperCase();
    out.push({ id: id, order: i, date: cell_(r[1]), category: cell_(r[2]) || "기타", title: title, body: cell_(r[4]),
      pinned: pin === "Y" || pin === "TRUE" || pin === "O" || pin === "고정", link: /^https:\/\//.test(cell_(r[6])) ? cell_(r[6]) : "" });
  }
  return out.reverse();
}

function noticeRow_(sh, id) {
  const ids = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getValues();
  for (let i = 1; i < ids.length; i++) if (cell_(ids[i][0]) === id) return i + 1;
  return 0;
}

function noticeAction_(d) {
  const admin = PropertiesService.getScriptProperties().getProperty("ADMIN_CODE");
  if (!admin || String(d.adminCode || "") !== admin) return { ok: false, error: "admin" };
  const sh = noticeSheet_();
  if (d.action === "notice_delete") {
    const row = noticeRow_(sh, String(d.id || ""));
    if (!row) return { ok: false, error: "notfound" };
    sh.deleteRow(row);
    return { ok: true, notices: readNotices_() };
  }
  const title = String(d.title || "").trim().slice(0, 100);
  if (!title) return { ok: false, error: "title" };
  const cat = NOTICE_CATS.indexOf(d.category) >= 0 ? d.category : "기타";
  const date = String(d.date || "").trim().slice(0, 20) || Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy.MM.dd");
  const body = String(d.body || "").slice(0, 5000);
  const link = /^https:\/\//.test(String(d.link || "").trim()) ? String(d.link).trim().slice(0, 500) : "";
  const pin = d.pinned ? "Y" : "";
  if (d.action === "notice_update") {
    const row = noticeRow_(sh, String(d.id || ""));
    if (!row) return { ok: false, error: "notfound" };
    sh.getRange(row, 2, 1, 6).setValues([[date, cat, title, body, pin, link]]);
  } else {
    sh.appendRow([Utilities.getUuid().slice(0, 8), date, cat, title, body, pin, link]);
  }
  return { ok: true, notices: readNotices_() };
}

function inArchive_(f) {
  const parents = f.getParents();
  while (parents.hasNext()) {
    const p = parents.next();
    if (!/^\d{4}년 [12]학기$/.test(p.getName())) continue;
    const gp = p.getParents();
    while (gp.hasNext()) if (gp.next().getId() === ROOT_ID) return true;
  }
  return false;
}

function itemOf_(f, sem) {
  let meta = {};
  try { meta = JSON.parse(f.getDescription() || "{}"); } catch (e) {}
  return { id: f.getId(), title: meta.title || f.getName(), kind: KINDS.indexOf(meta.kind) >= 0 ? meta.kind : "기타", semester: sem,
    date: Utilities.formatDate(f.getDateCreated(), "Asia/Seoul", "yyyy-MM-dd"), url: f.getUrl(), mime: f.getMimeType(), size: f.getSize(), uploader: meta.uploader || "" };
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (/^notice_(create|update|delete)$/.test(String(d.action || ""))) return json_(noticeAction_(d));
    const code = PropertiesService.getScriptProperties().getProperty("UPLOAD_CODE");
    if (!code || String(d.code || "") !== code) return json_({ ok: false, error: "code" });
    const action = d.action || "upload";

    if (action === "update" || action === "delete") {
      let f;
      try { f = DriveApp.getFileById(String(d.id || "")); } catch (err) { return json_({ ok: false, error: "notfound" }); }
      if (f.isTrashed() || !inArchive_(f)) return json_({ ok: false, error: "notfound" });
      if (action === "delete") {
        f.setTrashed(true); // 드라이브 휴지통으로 이동 (30일 안에 복구 가능)
        return json_({ ok: true, id: f.getId() });
      }
      if (!semLabel_(d.semester)) return json_({ ok: false, error: "semester" });
      const kind = KINDS.indexOf(d.kind) >= 0 ? d.kind : "기타";
      const title = String(d.title || "").trim().slice(0, 80);
      if (!title) return json_({ ok: false, error: "title" });
      let meta = {};
      try { meta = JSON.parse(f.getDescription() || "{}"); } catch (err) {}
      meta.title = title; meta.kind = kind;
      f.setDescription(JSON.stringify(meta));
      const ext = (String(meta.original || f.getName()).split(".").pop() || "").toLowerCase();
      f.setName("[" + kind + "] " + title + (ext ? "." + ext : ""));
      f.moveTo(semFolder_(d.semester));
      return json_({ ok: true, item: itemOf_(f, d.semester) });
    }

    if (!semLabel_(d.semester)) return json_({ ok: false, error: "semester" });
    const kind = KINDS.indexOf(d.kind) >= 0 ? d.kind : "기타";
    const title = String(d.title || "").trim().slice(0, 80);
    if (!title) return json_({ ok: false, error: "title" });
    const fileName = String(d.fileName || "file").replace(/[\\/:*?"<>|]/g, "_").slice(0, 120);
    const ext = (fileName.split(".").pop() || "").toLowerCase();
    if (EXTS.indexOf(ext) < 0) return json_({ ok: false, error: "type" });
    const bytes = Utilities.base64Decode(String(d.data || ""));
    if (!bytes.length) return json_({ ok: false, error: "empty" });
    if (bytes.length > MAX_BYTES) return json_({ ok: false, error: "size" });

    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let f;
    try {
      const blob = Utilities.newBlob(bytes, String(d.mime || "application/octet-stream"), "[" + kind + "] " + title + "." + ext);
      f = semFolder_(d.semester).createFile(blob);
    } finally { lock.releaseLock(); }
    f.setDescription(JSON.stringify({ title: title, kind: kind, uploader: String(d.uploader || "").trim().slice(0, 30), original: fileName }));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return json_({ ok: true, item: itemOf_(f, d.semester) });
  } catch (err) {
    return json_({ ok: false, error: "server" });
  }
}
