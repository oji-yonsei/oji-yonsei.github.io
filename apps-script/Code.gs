/**
 * 오지 탐험대 자료실 업로드 서버 (Google Apps Script 웹 앱)
 * - 대원이 사이트에서 올린 파일을 대표 계정 드라이브의 "오지 탐험대 자료실/학기" 폴더에 저장합니다.
 * - 사이트 자료실 목록은 이 폴더들에서 "링크가 있는 모든 사용자" 공유가 켜진 파일을 읽어 보여줍니다.
 *   (드라이브에서 파일 공유를 끄거나 파일을 지우면 사이트 목록에서도 사라집니다.)
 * - 업로드에는 대원 코드가 필요합니다: 프로젝트 설정 → 스크립트 속성에 UPLOAD_CODE 를 추가하세요.
 */
const ROOT_ID = "1GbHabez0EURsNq7VChHDylZ7f4aSlj33"; // 오지 탐험대 자료실 폴더
const KINDS = ["활동 보고서", "활동 사진", "카드뉴스", "발제 자료", "기타"];
const EXTS = ["pdf","jpg","jpeg","png","gif","webp","heic","hwp","hwpx","doc","docx","ppt","pptx","xls","xlsx","mp4","mov","txt","csv"];
const MAX_BYTES = 10 * 1024 * 1024; // 10MB

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
  return json_({ ok: true, items: items });
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const code = PropertiesService.getScriptProperties().getProperty("UPLOAD_CODE");
    if (!code || String(d.code || "") !== code) return json_({ ok: false, error: "code" });
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
    return json_({ ok: true, item: { id: f.getId(), title: title, kind: kind, semester: d.semester,
      date: Utilities.formatDate(new Date(), "Asia/Seoul", "yyyy-MM-dd"), url: f.getUrl(), mime: f.getMimeType(), size: f.getSize() } });
  } catch (err) {
    return json_({ ok: false, error: "server" });
  }
}
