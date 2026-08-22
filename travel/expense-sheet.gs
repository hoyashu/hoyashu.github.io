/**
 * 여행 지출 — 구글시트 백엔드 (Google Apps Script)  · 여행별 시트 버전
 *
 * 하는 일: travel/<여행>/ 일정표 페이지의 💰 지출 탭이 이 웹앱을 통해
 *          구글시트에 지출을 읽고/쓴다. 폰·노트북 어디서 넣어도 같은 시트에 쌓이고,
 *          같이 가는 사람들이 각자 폰에서 넣어도 하나로 모인다.
 *
 * ── 이전 버전과 달라진 점 ──────────────────────────────────
 * 요청에 `trip` 값을 같이 보내면 **그 이름의 시트**에 기록한다.
 *   trip 없음        → "지출"        시트  (고성 페이지가 쓰던 그대로 · 그래서 안 깨진다)
 *   trip: "yeongheung" → "지출-yeongheung" 시트
 * 여행이 늘어나도 스프레드시트 하나에 탭만 하나씩 늘어난다.
 * ──────────────────────────────────────────────────────────
 *
 * ── 이미 배포해 둔 걸 이 코드로 바꾸는 법 (URL 유지) ────────
 * 1. 기존 스프레드시트 → 확장 프로그램 → Apps Script
 * 2. 코드를 다 지우고 이 파일 내용을 통째로 붙여넣기 → 저장
 * 3. 우측 상단 "배포" → **"배포 관리"** → 기존 배포의 ✏️(연필) →
 *    버전을 **"새 버전"** 으로 바꾸고 → 배포
 *    ※ "새 배포"를 누르면 URL이 바뀝니다. 반드시 "배포 관리 → 편집"으로!
 * 4. URL이 그대로면 페이지는 손댈 것 없이 바로 동작합니다.
 * ──────────────────────────────────────────────────────────
 */

const SECRET      = 'goseong2026';   // ← 페이지 쪽 SYNC_KEY와 동일해야 함
const SHEET_BASE  = '지출';
const MAX_ROWS    = 800;             // 시트 하나당 폭주 방지

function doGet(e)  { return handle(e); }
function doPost(e) { return handle(e); }

function handle(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    let p = {};
    if (e && e.postData && e.postData.contents) {
      p = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      p = e.parameter;
    }

    if (p.key !== SECRET) return json({ ok: false, error: 'unauthorized' });

    const sh = getSheet(p.trip);
    const action = p.action || 'list';

    if (action === 'add') {
      if (sh.getLastRow() >= MAX_ROWS) return json({ ok: false, error: 'too many rows' });
      const name = String(p.name || '').slice(0, 100);
      const memo = String(p.memo || '').slice(0, 200);
      const amount = Math.round(Number(p.amount) || 0);
      if (!name || !amount) return json({ ok: false, error: 'invalid item' });
      sh.appendRow([
        String(p.id || ('u' + Date.now())),
        String(p.date || ''),
        String(p.cat || '기타'),
        name,
        amount,
        memo,
        new Date(),
        ''
      ]);
    } else if (action === 'delete') {
      const v = sh.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) {
        if (String(v[i][0]) === String(p.id)) sh.getRange(i + 1, 8).setValue('Y');
      }
    }

    return json({ ok: true, trip: sheetName(p.trip), items: readItems(sh) });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** trip 값을 시트 이름으로. 값이 없으면 예전 이름("지출")을 그대로 쓴다 */
function sheetName(trip) {
  const t = String(trip || '').replace(/[^A-Za-z0-9가-힣_-]/g, '').slice(0, 24);
  return t ? (SHEET_BASE + '-' + t) : SHEET_BASE;
}

/** 시트가 없으면 헤더와 함께 만든다 */
function getSheet(trip) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const name = sheetName(trip);
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(['id', '날짜', '분류', '항목', '금액(총액)', '메모', '기록시각', '삭제']);
    sh.setFrozenRows(1);
    sh.getRange('A1:H1').setFontWeight('bold');
    sh.setColumnWidth(4, 220);
    sh.setColumnWidth(6, 260);
  }
  return sh;
}

/** 삭제 표시(H열='Y')가 없는 행만 돌려준다 */
function readItems(sh) {
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][7]).toUpperCase() === 'Y') {
      out.push({ id: String(v[i][0]), deleted: true });
      continue;
    }
    if (!v[i][0]) continue;
    out.push({
      id: String(v[i][0]),
      date: String(v[i][1]),
      cat: String(v[i][2]),
      name: String(v[i][3]),
      amount: Number(v[i][4]) || 0,
      memo: String(v[i][5] || '')
    });
  }
  return out;
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
