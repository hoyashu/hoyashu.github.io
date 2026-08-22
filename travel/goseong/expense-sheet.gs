/**
 * 고성 여행 지출 — 구글시트 백엔드 (Google Apps Script)
 *
 * 하는 일: 여행 지출 페이지(https://hoyashu.github.io/goseong/)가 이 웹앱을 통해
 *          구글시트에 지출을 읽고/쓰게 해준다. 폰·노트북 어디서 넣어도 같은 시트에 쌓인다.
 *
 * ── 설치 방법 ──────────────────────────────────────────────
 * 1. 구글 드라이브에서 새 스프레드시트를 만든다 (이름은 아무거나, 예: "고성여행 지출")
 * 2. 상단 메뉴 → 확장 프로그램 → Apps Script
 * 3. 기본으로 있는 코드를 다 지우고 이 파일 내용을 통째로 붙여넣는다
 * 4. 아래 SECRET 값을 원하는 문자열로 바꾼다 (페이지 쪽과 똑같이 맞춰야 함)
 * 5. 저장 → 우측 상단 "배포" → "새 배포"
 *    - 유형: 웹 앱
 *    - 실행 사용자: 나
 *    - 액세스 권한: **모든 사용자**   ← 이걸 꼭 "모든 사용자"로
 * 6. 배포하면 권한 승인 창이 뜬다 → 본인 계정 선택 → "고급" → "안전하지 않음(이동)" → 허용
 *    (내가 만든 스크립트라 나오는 정상 경고입니다)
 * 7. 나오는 **웹 앱 URL**(https://script.google.com/macros/s/..../exec)을 복사해서 알려주면
 *    페이지에 심어서 배포하겠습니다.
 * ──────────────────────────────────────────────────────────
 */

const SECRET     = 'goseong2026';   // ← 페이지와 동일해야 함. 바꾸면 페이지 쪽도 알려주세요.
const SHEET_NAME = '지출';
const MAX_ROWS   = 500;             // 폭주 방지

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

    const sh = getSheet();
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

    return json({ ok: true, items: readItems(sh) });

  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** 시트가 없으면 헤더와 함께 만든다 */
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['id', '날짜', '분류', '항목', '금액(2인)', '메모', '기록시각', '삭제']);
    sh.setFrozenRows(1);
    sh.getRange('A1:H1').setFontWeight('bold');
    sh.setColumnWidth(4, 220);
    sh.setColumnWidth(6, 220);
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
