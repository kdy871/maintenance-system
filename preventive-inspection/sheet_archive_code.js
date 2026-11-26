/**
 * [Code.gs 교체/업그레이드용] 5개 파트 × (주간 / 월간 / 분기) 전체 시트 주기별 우측 이동 스크립트
 *
 * [동작 방식]
 * 1. 기존 '터빈전기(주간)' 1개 시트(고정 행)만 밀어주던 것을,
 *    시트 안의 모든 블록(10개든 20개든)과 5개 파트 전체 탭을 자동 인식해 이동시킵니다.
 * 2. 시트 탭 이름에 '(주간)' 또는 '주간'이 포함되면 -> 매주 금요일 18시 이후 자동 우측 이동
 *    시트 탭 이름에 '(월간)' 또는 '월간'이 포함되면 -> 매월 말일 18시 이후 자동 우측 이동
 *    시트 탭 이름에 '(분기)' 또는 '분기'가 포함되면 -> 분기 말일(3/6/9/12월 말일) 18시 이후 자동 우측 이동
 * 3. E:F열을 오른쪽(G:H)으로 1칸 깔끔하게 넘긴 뒤, 새 E:F열을 비우고 날짜 칸은 '미점검'으로 초기화합니다.
 *    -> 이후 웹에서 블록별로 저장할 때 실제 입력한 날짜·시간이 해당 블록 E 날짜 칸에 기록됩니다.
 *
 * ※ 주의: WebApp.gs(google_apps_script.js)의 parseInspectionSheet(sheet) 함수를 함께 사용합니다.
 *         반드시 같은 Apps Script 프로젝트 안에 WebApp.gs와 함께 저장하세요.
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('점검표 관리')
    .addItem('현재 시트 수동 마감 (우측으로 이동)', 'manualShiftCurrentSheet')
    .addItem('전체 주간 시트 수동 마감', 'manualShiftAllWeekly')
    .addSeparator()
    .addItem('주기별 자동 마감 트리거 설정', 'setupPeriodicTriggers')
    .addToUi();
}

/**
 * [수동 마감] 현재 보고 있는 시트의 E:F열을 오른쪽(G:H)으로 이동하고 새 입력칸 비우기
 */
function manualShiftCurrentSheet() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const response = ui.alert(
    '수동 마감 확인',
    `[${sheet.getName()}] 시트의 현재 입력 데이터(E:F열)를 오른쪽(G:H열)으로 이동하고 새 입력칸을 비우시겠습니까?`,
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    const shiftedBlocks = shiftSheetBlocksRight_(sheet);
    ui.alert(`[${sheet.getName()}] ${shiftedBlocks}개 블록의 우측 이동 및 초기화가 완료되었습니다.`);
  }
}

/**
 * [수동 마감] 모든 '주간' 시트를 한 번에 우측 이동
 */
function manualShiftAllWeekly() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '전체 주간 시트 마감 확인',
    '모든 주간 시트 탭의 E:F열을 오른쪽으로 이동하고 비우시겠습니까?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let count = 0;
  for (let i = 0; i < sheets.length; i++) {
    if (!sheets[i].isSheetHidden() && sheets[i].getName().indexOf('주간') !== -1) {
      shiftSheetBlocksRight_(sheets[i]);
      count++;
    }
  }
  ui.alert(`총 ${count}개 주간 시트의 마감이 완료되었습니다.`);
}

/**
 * [자동 마감] 시간 기반 트리거가 1시간마다 실행하여 주간/월간/분기 마감 시점이면 1회씩 우측 이동
 */
function autoShiftBySchedule() {
  const now = new Date();
  const day = now.getDay();   // 0:일 ~ 5:금, 6:토
  const hour = now.getHours();
  const month = now.getMonth() + 1; // 1 ~ 12
  const date = now.getDate();
  const lastDayOfMonth = new Date(now.getFullYear(), month, 0).getDate();

  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  // 1. 주간 마감 조건: 금요일 18시 이후 또는 토/일요일
  const isWeeklyWindow = (day === 5 && hour >= 18) || day === 6 || day === 0;
  const weekKey = getWeekKey_(now);
  if (isWeeklyWindow && props.getProperty('LAST_SHIFTED_WEEK_ALL') !== weekKey) {
    let shiftedAny = false;
    for (let i = 0; i < sheets.length; i++) {
      const s = sheets[i];
      if (!s.isSheetHidden() && s.getName().indexOf('주간') !== -1) {
        shiftSheetBlocksRight_(s);
        shiftedAny = true;
      }
    }
    if (shiftedAny) {
      props.setProperty('LAST_SHIFTED_WEEK_ALL', weekKey);
    }
  }

  // 2. 월간 마감 조건: 매월 말일 18시 이후
  const isMonthEndWindow = (date === lastDayOfMonth && hour >= 18);
  const monthKey = `${now.getFullYear()}-${month}`;
  if (isMonthEndWindow && props.getProperty('LAST_SHIFTED_MONTH_ALL') !== monthKey) {
    let shiftedAny = false;
    for (let i = 0; i < sheets.length; i++) {
      const s = sheets[i];
      if (!s.isSheetHidden() && s.getName().indexOf('월간') !== -1) {
        shiftSheetBlocksRight_(s);
        shiftedAny = true;
      }
    }
    if (shiftedAny) {
      props.setProperty('LAST_SHIFTED_MONTH_ALL', monthKey);
    }
  }

  // 3. 분기 마감 조건: 3/6/9/12월 말일 18시 이후
  const isQuarterEndMonth = (month === 3 || month === 6 || month === 9 || month === 12);
  const quarterNo = Math.ceil(month / 3);
  const quarterKey = `${now.getFullYear()}-Q${quarterNo}`;
  if (isQuarterEndMonth && isMonthEndWindow && props.getProperty('LAST_SHIFTED_QUARTER_ALL') !== quarterKey) {
    let shiftedAny = false;
    for (let i = 0; i < sheets.length; i++) {
      const s = sheets[i];
      if (!s.isSheetHidden() && s.getName().indexOf('분기') !== -1) {
        shiftSheetBlocksRight_(s);
        shiftedAny = true;
      }
    }
    if (shiftedAny) {
      props.setProperty('LAST_SHIFTED_QUARTER_ALL', quarterKey);
    }
  }
}

/**
 * 시트 안의 모든 블록을 자동 감지해 E:F열 -> G:H열로 우측 이동 후 E:F열 초기화
 */
function shiftSheetBlocksRight_(sheet) {
  const parsed = parseInspectionSheet(sheet);
  if (!parsed.blocks || parsed.blocks.length === 0) return 0;

  const firstBlock = parsed.blocks[0];
  const valCol = firstBlock.firstValCol || 5; // 기본 E열(5)
  const noteCol = (firstBlock.dates && firstBlock.dates[0] && firstBlock.dates[0].noteCol) || (valCol + 1); // 기본 F열(6)

  // F열(noteCol) 바로 뒤에 2열(새 G:H)을 삽입 (기존 G:H...는 자동으로 I:J...로 밀림)
  sheet.insertColumnsAfter(noteCol, 2);

  const newHistValCol = noteCol + 1;
  const newHistNoteCol = noteCol + 2;

  // 열 너비 맞추기
  sheet.setColumnWidth(newHistValCol, sheet.getColumnWidth(valCol));
  sheet.setColumnWidth(newHistNoteCol, sheet.getColumnWidth(noteCol));

  for (let b = 0; b < parsed.blocks.length; b++) {
    const blk = parsed.blocks[b];
    const startRow = Math.min(blk.dateHeaderRow, blk.subHeaderRow);
    const endRow = blk.dataEndRow;
    const rowSpan = endRow - startRow + 1;
    if (rowSpan <= 0) continue;

    // 1. E:F 전체(날짜 헤더 + 서브헤더 + 데이터)를 새로 삽입된 G:H로 복사 (값 + 서식)
    const srcRange = sheet.getRange(startRow, valCol, rowSpan, 2);
    const dstRange = sheet.getRange(startRow, newHistValCol, rowSpan, 2);
    srcRange.copyTo(dstRange);

    // 과거 이력 열(G열)에 불필요한 드롭다운 유효성 검사 제거 (값은 그대로 유지)
    if (blk.dataEndRow >= blk.dataStartRow) {
      sheet.getRange(blk.dataStartRow, newHistValCol, blk.dataEndRow - blk.dataStartRow + 1, 1).clearDataValidations();
    }

    // 2. 현재 입력칸 E:F 데이터 비우기 (드롭다운과 서식은 그대로 유지!)
    if (blk.dataEndRow >= blk.dataStartRow) {
      sheet.getRange(blk.dataStartRow, valCol, blk.dataEndRow - blk.dataStartRow + 1, 2).clearContent();
    }

    // 3. 새 입력칸 날짜 헤더를 '미점검'으로 설정 (웹에서 저장 시 실제 날짜·시간으로 자동 변경됨)
    sheet.getRange(blk.dateHeaderRow, valCol).setValue('미점검');
  }

  SpreadsheetApp.flush();
  return parsed.blocks.length;
}

/**
 * [트리거 자동 등록] 주간/월간/분기 자동 마감 트리거 설정
 */
function setupPeriodicTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const fn = triggers[i].getHandlerFunction();
    if (fn === 'autoShiftAfterFriday18' || fn === 'autoShiftBySchedule') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('autoShiftBySchedule')
    .timeBased()
    .everyHours(1)
    .create();

  SpreadsheetApp.getUi().alert('주간(금 18시)/월간(말일 18시)/분기(분기말 18시) 자동 마감 트리거가 설정되었습니다.');
}

function getWeekKey_(date) {
  const d = new Date(date.getFullYear(), date.getMonth(), date.getDate());
  d.setDate(d.getDate() + 4 - (d.getDay() || 7));
  const yearStart = new Date(d.getFullYear(), 0, 1);
  const weekNo = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return `${d.getFullYear()}-${weekNo}`;
}
