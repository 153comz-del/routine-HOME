// 루틴 앱 데이터 API (v6 - 저금통 추가)
const SS = SpreadsheetApp.getActiveSpreadsheet();
const DEFAULT_CATS = ['집안일', '식물', '반려동물', '건강', '기타'];

// 카테고리 시트를 가져오고, 없으면 자동으로 만들어요
function catSheet() {
  let sh = SS.getSheetByName('카테고리');
  if (!sh) {
    sh = SS.insertSheet('카테고리');
    sh.getRange(1, 1).setValue('이름');
    sh.getRange(2, 1, DEFAULT_CATS.length, 1).setValues(DEFAULT_CATS.map(c => [c]));
  }
  return sh;
}

// 음식 시트를 가져오고, 없으면 자동으로 만들어요 (v4: '개수' 열 추가)
function foodSheet() {
  let sh = SS.getSheetByName('음식목록');
  if (!sh) {
    sh = SS.insertSheet('음식목록');
    sh.getRange(1, 1, 1, 4).setValues([['id', '이름', '유통기한', '개수']]);
  } else if (sh.getRange(1, 4).getValue() === '') {
    // 기존 시트(3열)에는 개수 헤더만 자동으로 붙여요
    sh.getRange(1, 4).setValue('개수');
  }
  return sh;
}

// 저금통 시트를 가져오고, 없으면 자동으로 만들어요
function bankSheet() {
  let sh = SS.getSheetByName('저금통');
  if (!sh) {
    sh = SS.insertSheet('저금통');
    sh.getRange(1, 1, 1, 4).setValues([['id', '이름', '목표', '현재']]);
  }
  return sh;
}

// 데이터 읽기 (앱이 열릴 때 호출)
function doGet(e) {
  try {
    const taskSheet = SS.getSheetByName('할일목록');
    const logSheet = SS.getSheetByName('완료기록');

    const taskValues = taskSheet.getDataRange().getValues();
    const tasks = taskValues.slice(1).filter(r => r[0] !== '').map(r => ({
      id: String(r[0]),
      name: r[1],
      category: r[2],
      cycleType: r[3],   // 한번 / 매일 / 일 / 주 / 월 / 요일 / 월일 / 메모
      cycleValue: Number(r[4]) || 1,
      nextDate: formatDate(r[5])
    }));

    const logValues = logSheet.getDataRange().getValues();
    const logs = logValues.slice(1).filter(r => r[0] !== '')
      .slice(-100)  // 최근 100건만
      .map(r => ({
        date: formatDate(r[0]),
        name: r[1],
        type: r[2]
      }));

    // 카테고리 읽기 (시트가 없으면 자동 생성)
    const catValues = catSheet().getDataRange().getValues();
    let categories = catValues.slice(1).map(r => String(r[0])).filter(c => c !== '');
    if (categories.length === 0) categories = DEFAULT_CATS;

    // 음식 읽기 (시트가 없으면 자동 생성) — 유통기한이 없으면 빈칸, 개수 기본 1
    const foodValues = foodSheet().getDataRange().getValues();
    const foods = foodValues.slice(1).filter(r => r[0] !== '').map(r => ({
      id: String(r[0]),
      name: r[1],
      expiry: r[2] === '' ? '' : formatDate(r[2]),
      qty: Number(r[3]) || 1
    }));

    // 저금통 읽기 (시트가 없으면 자동 생성)
    const bankValues = bankSheet().getDataRange().getValues();
    const banks = bankValues.slice(1).filter(r => r[0] !== '').map(r => ({
      id: String(r[0]),
      name: r[1],
      target: Number(r[2]) || 0,
      current: Number(r[3]) || 0
    }));

    return jsonResponse({ ok: true, tasks: tasks, logs: logs, categories: categories, foods: foods, banks: banks });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// 데이터 저장 (완료, 미루기, 추가, 삭제, 카테고리/음식 변경 시 호출)
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);  // 동시 저장 충돌 방지 (20초까지 대기)
    const body = JSON.parse(e.postData.contents);

    if (body.action === 'saveTasks') {
      // 할일목록 전체를 새로 씀
      const sheet = SS.getSheetByName('할일목록');
      sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 6).clearContent();
      const rows = body.tasks.map(t => [
        t.id, t.name, t.category, t.cycleType, t.cycleValue, t.nextDate
      ]);
      if (rows.length > 0) {
        sheet.getRange(2, 1, rows.length, 6).setValues(rows);
      }
    }

    if (body.action === 'addLog') {
      // 완료/미룸 기록 한 줄 추가
      SS.getSheetByName('완료기록')
        .appendRow([body.log.date, body.log.name, body.log.type]);
    }

    if (body.action === 'saveCategories') {
      // 카테고리 목록 전체를 새로 씀
      const sh = catSheet();
      sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).clearContent();
      const rows = (body.categories || []).map(c => [c]);
      if (rows.length > 0) {
        sh.getRange(2, 1, rows.length, 1).setValues(rows);
      }
    }

    if (body.action === 'saveFoods') {
      // 음식 목록 전체를 새로 씀 (v4: 4열 - 유통기한 빈칸 허용, 개수)
      const sh = foodSheet();
      sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 4).clearContent();
      const rows = (body.foods || []).map(f => [
        f.id, f.name, f.expiry || '', Number(f.qty) || 1
      ]);
      if (rows.length > 0) {
        sh.getRange(2, 1, rows.length, 4).setValues(rows);
      }
    }

    if (body.action === 'saveBanks') {
      // 저금통 목록 전체를 새로 씀
      const sh = bankSheet();
      sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 4).clearContent();
      const rows = (body.banks || []).map(b => [
        b.id, b.name, Number(b.target) || 0, Number(b.current) || 0
      ]);
      if (rows.length > 0) {
        sh.getRange(2, 1, rows.length, 4).setValues(rows);
      }
    }

    if (body.action === 'deleteLog') {
      // 완료기록에서 일치하는 기록 한 줄을 아래(최신)부터 찾아 삭제
      const sh = SS.getSheetByName('완료기록');
      const values = sh.getDataRange().getValues();
      for (let i = values.length - 1; i >= 1; i--) {
        const r = values[i];
        if (formatDate(r[0]) === String(body.log.date) &&
            String(r[1]) === String(body.log.name) &&
            String(r[2]) === String(body.log.type)) {
          sh.deleteRow(i + 1);
          break;
        }
      }
    }

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function formatDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  const s = String(v).trim();
  if (s === '') return '';
  // 컴퓨터 형식(2026-09-14T15:00:00.000Z)은 한국 시간 기준 날짜로 변환
  if (/^\d{4}-\d{2}-\d{2}T\d{2}:/.test(s)) {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  // 시트에 글자로 적힌 날짜도 알아서 읽어요: 2026.8.27, 2026/8/27, 2026년 8월 27일 등
  const m = s.match(/(\d{4})[.\/\-년\s]+(\d{1,2})[.\/\-월\s]+(\d{1,2})/);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  // 영어 형식(Sat Aug 29 2026 ... GMT+0900)도 복구
  const d = new Date(s);
  if (!isNaN(d.getTime()) && d.getFullYear() >= 2000 && d.getFullYear() <= 2100) {
    return Utilities.formatDate(d, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  return s;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
