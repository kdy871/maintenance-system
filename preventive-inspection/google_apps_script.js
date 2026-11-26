/**
 * 발전소 설비 예방점검 Google Apps Script (웹 앱용 - WebApp.gs에 추가)
 *
 * [주의] 기존 시트의 Code.gs (onOpen, manualSaveAndShift, autoShiftAfterFriday18 등)는 지우지 마세요!
 *        Apps Script 왼쪽 [파일 +] -> [스크립트] 추가(이름: WebApp) 후 이 코드를 넣어 함께 사용합니다.
 *
 * [동작 방식]
 * 1. 한 시트 안에 세로로 이어진 여러 점검표(3호기 발전기 주간 점검표, 4호기 발전기 주간 점검표 등)를 정확히 분리 인식합니다.
 * 2. 시트의 첫 번째 데이터 입력 열(E:F열)을 '현재 입력칸'으로 사용하여 저장합니다.
 *    -> 기존 금요일 18시/수동 저장(E:L -> G:N 우측 이동) 기능과 충돌 없이 그대로 이어집니다.
 * 3. 시트에서 항목(행)을 추가/삭제/수정하거나 새 시트 탭(월간/분기 등)을 추가하면 웹 새로고침 시 즉시 반영됩니다.
 */

const SPREADSHEET_ID = '1fmNem3oXLQFqg22ZJ9an55x5gSudjWz5jONRl1f_2b8';

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheets = ss.getSheets().filter(function (s) {
      return !s.isSheetHidden() && s.getName() !== '점검이력_DB';
    });

    if (sheets.length === 0) {
      return jsonResponse({ success: false, error: '사용 가능한 시트 탭이 없습니다.' });
    }

    const requestedSheetName = e && e.parameter && e.parameter.sheet ? String(e.parameter.sheet) : '';
    const sheet = (requestedSheetName && ss.getSheetByName(requestedSheetName)) || sheets[0];
    const sheetNames = sheets.map(function (s) { return s.getName(); });

    const parsed = parseInspectionSheet(sheet);

    return jsonResponse({
      success: true,
      spreadsheetId: SPREADSHEET_ID,
      sheetNames: sheetNames,
      activeSheet: sheet.getName(),
      blocks: parsed.blocks,
      dates: parsed.allDates,
      items: parsed.items,
      updatedAt: new Date().toISOString()
    });
  } catch (err) {
    return jsonResponse({
      success: false,
      error: err.toString(),
      stack: err.stack || ''
    });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const bodyText = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    const payload = JSON.parse(bodyText);

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = payload.sheetName ? String(payload.sheetName) : '';
    const sheet = (sheetName && ss.getSheetByName(sheetName)) || ss.getSheets()[0];

    const dateLabel = String(payload.dateLabel || '').trim();
    const entries = Array.isArray(payload.entries) ? payload.entries : [];

    if (!dateLabel) {
      return jsonResponse({ success: false, error: '저장할 점검 날짜·시간(dateLabel)이 비어 있습니다.' });
    }

    const parsed = parseInspectionSheet(sheet);
    if (!parsed.blocks || parsed.blocks.length === 0) {
      return jsonResponse({ success: false, error: '시트에서 "점검값" 헤더를 찾지 못했습니다.' });
    }

    const blockById = {};
    for (let b = 0; b < parsed.blocks.length; b++) {
      blockById[parsed.blocks[b].id] = parsed.blocks[b];
    }

    const itemByKey = {};
    const itemByRow = {};
    for (let i = 0; i < parsed.items.length; i++) {
      const item = parsed.items[i];
      itemByKey[item.key] = item;
      itemByRow[item.row] = item;
    }

    // 블록별 '현재 입력칸'(첫 번째 점검값/점검내용 열, 예: E:F열)에 저장하고 날짜 헤더 갱신
    const targetColByBlockId = {};

    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i];
      let matchedItem = null;
      if (entry.key && itemByKey[entry.key]) {
        matchedItem = itemByKey[entry.key];
      } else if (entry.row && itemByRow[entry.row]) {
        const candidate = itemByRow[entry.row];
        if (!entry.label || candidate.label === entry.label) {
          matchedItem = candidate;
        }
      }
      if (!matchedItem) continue;

      const blk = blockById[matchedItem.blockId];
      if (!blk || targetColByBlockId[blk.id]) continue;

      let activePair = blk.dates && blk.dates.length > 0 ? blk.dates[0] : null;
      if (activePair) {
        // 기존 E:F 입력칸의 날짜 헤더(예: E1, E46)를 웹에서 저장한 현재 날짜·시간으로 업데이트
        sheet.getRange(blk.dateHeaderRow, activePair.valCol).setValue(dateLabel);
      } else {
        activePair = createFirstDateColumnPairForBlock(sheet, blk, dateLabel);
      }
      targetColByBlockId[blk.id] = activePair;
    }

    let updatedCount = 0;
    const unmatched = [];

    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i];
      let matchedItem = null;

      if (entry.key && itemByKey[entry.key]) {
        matchedItem = itemByKey[entry.key];
      } else if (entry.row && itemByRow[entry.row]) {
        const candidate = itemByRow[entry.row];
        if (!entry.label || candidate.label === entry.label) {
          matchedItem = candidate;
        }
      }

      if (!matchedItem) {
        unmatched.push(entry.label || entry.key || ('row:' + entry.row));
        continue;
      }

      const targetPair = targetColByBlockId[matchedItem.blockId];
      if (!targetPair) {
        unmatched.push(entry.label || entry.key);
        continue;
      }

      const rowNum = matchedItem.row;
      if (entry.value !== undefined) {
        sheet.getRange(rowNum, targetPair.valCol).setValue(entry.value);
      }
      if (entry.note !== undefined && targetPair.noteCol) {
        sheet.getRange(rowNum, targetPair.noteCol).setValue(entry.note);
      }
      updatedCount++;
    }

    SpreadsheetApp.flush();

    // 저장 후 해당 블록의 E:F 데이터가 모두 비워져 있다면 날짜 헤더도 '미점검'으로 정리
    const touchedBlockIds = Object.keys(targetColByBlockId);
    for (let b = 0; b < touchedBlockIds.length; b++) {
      const bId = touchedBlockIds[b];
      const blk = blockById[bId];
      const pair = targetColByBlockId[bId];
      if (!blk || !pair || blk.dataEndRow < blk.dataStartRow) continue;

      const colCount = pair.noteCol ? 2 : 1;
      const vals = sheet.getRange(blk.dataStartRow, pair.valCol, blk.dataEndRow - blk.dataStartRow + 1, colCount).getDisplayValues();
      const hasAnyValue = vals.some(function (r) {
        return r.some(function (cell) { return String(cell || '').trim() !== ''; });
      });
      if (!hasAnyValue) {
        sheet.getRange(blk.dateHeaderRow, pair.valCol).setValue('미점검');
      }
    }

    SpreadsheetApp.flush();

    return jsonResponse({
      success: true,
      activeSheet: sheet.getName(),
      dateLabel: dateLabel,
      updatedCount: updatedCount,
      unmatched: unmatched
    });
  } catch (err) {
    return jsonResponse({
      success: false,
      error: err.toString(),
      stack: err.stack || ''
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 시트 전체를 스캔하여 세로로 여러 개 배치된 예방점검 블록(3호기, 4호기 등)을 정확히 파싱합니다.
 */
function parseInspectionSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 2) {
    return { blocks: [], allDates: [], items: [] };
  }

  const allDisplay = sheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  // 1. 시트 내 모든 "점검값" 서브헤더 행 탐색
  const rawHeaderBlocks = [];
  for (let r = 0; r < lastRow; r++) {
    let firstValColIdx = -1;
    for (let c = 0; c < lastCol; c++) {
      const cell = String(allDisplay[r][c] || '').replace(/\s+/g, '');
      if (cell.indexOf('점검값') !== -1) {
        firstValColIdx = c;
        break;
      }
    }
    if (firstValColIdx !== -1) {
      rawHeaderBlocks.push({
        subHeaderRowIdx: r,
        firstValColIdx: firstValColIdx
      });
    }
  }

  if (rawHeaderBlocks.length === 0) {
    return { blocks: [], allDates: [], items: [] };
  }

  // 2. 병합 셀 맵 구축 (제목 가로 병합 행 포함)
  const fullMergedRanges = sheet.getRange(1, 1, lastRow, lastCol).getMergedRanges();
  const mergeFill = {};
  const horizMergeMeta = {};

  for (let m = 0; m < fullMergedRanges.length; m++) {
    const mr = fullMergedRanges[m];
    const rStart = mr.getRow() - 1;
    const rEnd = mr.getLastRow() - 1;
    const cStart = mr.getColumn() - 1;
    const cEnd = mr.getLastColumn() - 1;
    const val = String(mr.getDisplayValue() || '').trim();

    for (let r = rStart; r <= rEnd; r++) {
      for (let c = cStart; c <= cEnd; c++) {
        if (val) mergeFill[r + '_' + c] = val;
      }
    }

    if (rStart === rEnd && cStart === 0 && (cEnd - cStart + 1) >= 2 && val) {
      horizMergeMeta[rStart] = { text: val, cols: cEnd - cStart + 1 };
    }
  }

  function getCellText(r0, c0) {
    const direct = String(allDisplay[r0][c0] || '').trim();
    if (direct) return direct;
    return mergeFill[r0 + '_' + c0] || '';
  }

  // 3. 각 블록의 제목 행(titleRowIdx), 날짜 행(dateRowIdx), 시작 행(blockStartRowIdx)을 먼저 정확히 계산
  const blockMeta = [];
  for (let b = 0; b < rawHeaderBlocks.length; b++) {
    const hb = rawHeaderBlocks[b];
    const subRowIdx = hb.subHeaderRowIdx;
    const prevSubIdx = b > 0 ? rawHeaderBlocks[b - 1].subHeaderRowIdx : -1;
    const searchFloor = Math.max(prevSubIdx + 1, subRowIdx - 2);

    // 날짜 행: 보통 서브헤더 바로 위(subRowIdx - 1)
    const dateRowIdx = Math.max(0, subRowIdx - 1);

    // 대제목(3호기 발전기 주간 점검표, 4호기 발전기 주간 점검표 등) 탐색:
    // 서브헤더 바로 위(subRowIdx - 1)부터 위로 올라가며 가로 병합된 제목을 우선 찾음
    let blockTitle = '';
    let titleRowIdx = -1;

    for (let tr = subRowIdx - 1; tr >= searchFloor; tr--) {
      if (horizMergeMeta[tr] && horizMergeMeta[tr].text) {
        blockTitle = horizMergeMeta[tr].text;
        titleRowIdx = tr;
        break;
      }
    }

    // 가로 병합이 없는 경우에만 바로 위 1~2행 중 첫 번째 비어있지 않은 제목 행 확인
    if (!blockTitle) {
      for (let tr = subRowIdx - 1; tr >= searchFloor; tr--) {
        // 직전 블록의 마지막 데이터 행(예: No가 숫자이고 E열에 드롭다운이 있거나 일반 항목인 행)을 오인하지 않도록,
        // 오른쪽 첫 번째 날짜칸에 날짜/공백이 있고 A열이 숫자가 아닌 행만 제목 후보로 인정
        const colAText = getCellText(tr, 0);
        const looksLikeNumberedDataRow = colAText !== '' && !isNaN(Number(colAText));
        if (looksLikeNumberedDataRow) continue;

        const rowTexts = [];
        for (let c = 0; c < hb.firstValColIdx; c++) {
          const t = getCellText(tr, c);
          if (t && rowTexts.indexOf(t) === -1) rowTexts.push(t);
        }
        if (rowTexts.length > 0) {
          blockTitle = rowTexts.join(' ');
          titleRowIdx = tr;
          break;
        }
      }
    }

    if (!blockTitle) {
      blockTitle = '예방점검 #' + (b + 1);
    }

    const blockStartRowIdx = titleRowIdx >= 0 ? Math.min(titleRowIdx, dateRowIdx) : dateRowIdx;
    blockMeta.push({
      subRowIdx: subRowIdx,
      firstValColIdx: hb.firstValColIdx,
      dateRowIdx: dateRowIdx,
      titleRowIdx: titleRowIdx,
      blockStartRowIdx: blockStartRowIdx,
      blockTitle: blockTitle
    });
  }

  // 4. 블록별 데이터 행 파싱
  const parsedBlocks = [];
  const allDatesMap = {};
  const allDatesList = [];
  const allItems = [];
  const keyOccurrences = {};

  for (let b = 0; b < blockMeta.length; b++) {
    const bm = blockMeta[b];
    const subRowIdx = bm.subRowIdx;
    const dateRowIdx = bm.dateRowIdx;
    const dataStartRowIdx = subRowIdx + 1;
    const dataEndRowIdx = (b + 1 < blockMeta.length)
      ? (blockMeta[b + 1].blockStartRowIdx - 1)
      : (lastRow - 1);

    const metaColCount = bm.firstValColIdx;
    const metaHeaders = [];
    for (let c = 0; c < metaColCount; c++) {
      const lbl = String(allDisplay[subRowIdx][c] || '').trim();
      metaHeaders.push(lbl || ('열 ' + columnToLetter(c + 1)));
    }

    // 날짜 열 쌍 추출
    const blockDates = [];
    for (let c = bm.firstValColIdx; c < lastCol; c++) {
      const subText = String(allDisplay[subRowIdx][c] || '').replace(/\s+/g, '');
      if (subText.indexOf('점검값') !== -1) {
        let noteCol = null;
        if (c + 1 < lastCol) {
          const nextText = String(allDisplay[subRowIdx][c + 1] || '').replace(/\s+/g, '');
          if (nextText.indexOf('점검내용') !== -1 || nextText === '') {
            noteCol = c + 2; // 1-based
          }
        }
        const rawDateLabel = getCellText(dateRowIdx, c);
        const dateLabel = rawDateLabel || ('__EMPTY_DATE_' + columnToLetter(c + 1));
        const valHeader = String(allDisplay[subRowIdx][c] || '').trim() || '점검값';
        const noteHeader = noteCol ? (String(allDisplay[subRowIdx][noteCol - 1] || '').trim() || '점검내용') : '점검내용';
        const dObj = {
          label: dateLabel,
          hasRealLabel: rawDateLabel !== '',
          valCol: c + 1,
          noteCol: noteCol,
          valHeader: valHeader,
          noteHeader: noteHeader
        };
        blockDates.push(dObj);

        if (rawDateLabel && !allDatesMap[dateLabel]) {
          allDatesMap[dateLabel] = true;
          allDatesList.push(dateLabel);
        }
      }
    }

    const blockId = 'block_' + (b + 1);
    const blockInfo = {
      id: blockId,
      title: bm.blockTitle,
      titleRow: bm.titleRowIdx >= 0 ? bm.titleRowIdx + 1 : null,
      dateHeaderRow: dateRowIdx + 1,
      subHeaderRow: subRowIdx + 1,
      dataStartRow: dataStartRowIdx + 1,
      dataEndRow: dataEndRowIdx + 1,
      metaColCount: metaColCount,
      metaHeaders: metaHeaders,
      valHeader: blockDates.length > 0 ? blockDates[0].valHeader : '점검값',
      noteHeader: blockDates.length > 0 ? blockDates[0].noteHeader : '점검내용',
      firstValCol: bm.firstValColIdx + 1,
      dates: blockDates
    };
    parsedBlocks.push(blockInfo);

    if (dataEndRowIdx < dataStartRowIdx) continue;

    // 드롭다운(Data Validation) 추출 (E열 기준)
    const numRows = dataEndRowIdx - dataStartRowIdx + 1;
    const valColToInspect = blockDates.length > 0 ? blockDates[0].valCol : (bm.firstValColIdx + 1);
    const validations = sheet.getRange(dataStartRowIdx + 1, valColToInspect, numRows, 1).getDataValidations();

    let currentSubSectionTitle = bm.blockTitle;

    for (let r = dataStartRowIdx; r <= dataEndRowIdx; r++) {
      // 중간에 A~D 전체가 가로 병합된 소제목 행이 있으면 소제목으로 분리
      if (horizMergeMeta[r] && horizMergeMeta[r].cols >= metaColCount && metaColCount >= 2) {
        let rightEmpty = true;
        for (let d = 0; d < blockDates.length; d++) {
          if (String(allDisplay[r][blockDates[d].valCol - 1] || '').trim() !== '') {
            rightEmpty = false;
            break;
          }
        }
        if (rightEmpty) {
          currentSubSectionTitle = horizMergeMeta[r].text;
          continue;
        }
      }

      const metaValues = [];
      let hasMeta = false;
      for (let c = 0; c < metaColCount; c++) {
        const val = getCellText(r, c);
        metaValues.push(val);
        if (val) hasMeta = true;
      }

      if (!hasMeta) continue;

      const nonEmptyMeta = metaValues.filter(function (v) { return v !== ''; });
      const label = nonEmptyMeta[nonEmptyMeta.length - 1];

      const baseKey = blockId + '||' + currentSubSectionTitle + '||' + metaValues.join('||');
      keyOccurrences[baseKey] = (keyOccurrences[baseKey] || 0) + 1;
      const uniqueKey = baseKey + '#' + keyOccurrences[baseKey];

      const rule = validations[r - dataStartRowIdx][0];
      const options = extractDropdownOptions(rule);

      const valuesByDate = {};
      for (let d = 0; d < blockDates.length; d++) {
        const bd = blockDates[d];
        const v = String(allDisplay[r][bd.valCol - 1] || '').trim();
        const n = bd.noteCol ? String(allDisplay[r][bd.noteCol - 1] || '').trim() : '';
        valuesByDate[bd.label] = { value: v, note: n };
      }

      allItems.push({
        row: r + 1,
        blockId: blockId,
        blockTitle: bm.blockTitle,
        sectionTitle: currentSubSectionTitle,
        key: uniqueKey,
        label: label,
        meta: metaValues,
        options: options,
        valuesByDate: valuesByDate
      });
    }
  }

  return {
    blocks: parsedBlocks,
    allDates: allDatesList,
    items: allItems
  };
}

function createFirstDateColumnPairForBlock(sheet, blockInfo, dateLabel) {
  const newValCol = blockInfo.firstValCol || (blockInfo.metaColCount + 1);
  const newNoteCol = newValCol + 1;

  const dateRange = sheet.getRange(blockInfo.dateHeaderRow, newValCol, 1, 2);
  dateRange.breakApart();
  dateRange.merge();
  dateRange.setValue(dateLabel);
  dateRange.setHorizontalAlignment('center');

  sheet.getRange(blockInfo.subHeaderRow, newValCol).setValue('점검값');
  sheet.getRange(blockInfo.subHeaderRow, newNoteCol).setValue('점검내용');

  return {
    label: dateLabel,
    hasRealLabel: true,
    valCol: newValCol,
    noteCol: newNoteCol
  };
}

function extractDropdownOptions(rule) {
  if (!rule) return [];
  const type = rule.getCriteriaType();
  const args = rule.getCriteriaValues();

  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    const list = Array.isArray(args[0]) ? args[0] : [];
    return list.map(function (v) { return String(v).trim(); }).filter(Boolean);
  }

  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && args[0]) {
    try {
      const vals = args[0].getDisplayValues();
      const flat = [];
      for (let i = 0; i < vals.length; i++) {
        for (let j = 0; j < vals[i].length; j++) {
          const s = String(vals[i][j] || '').trim();
          if (s && flat.indexOf(s) === -1) flat.push(s);
        }
      }
      return flat;
    } catch (e) {
      return [];
    }
  }

  return [];
}

function columnToLetter(col) {
  let temp;
  let letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
