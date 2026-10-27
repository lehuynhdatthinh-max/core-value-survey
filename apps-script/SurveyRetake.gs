/**
 * KHẢO SÁT VHDN – XEM ĐIỂM + LÀM LẠI, GỘP THEO MSNV TRONG TỪNG THÁNG   (thêm 27/10/2026)
 *
 * Khảo sát làm hằng tháng, nên trong sheet Responses mỗi MSNV có MỘT dòng cho MỖI THÁNG, luôn là lần nộp mới nhất của tháng đó:
 *   - Tháng = tháng của cột Timestamp, theo giờ Việt Nam (múi giờ của dự án Apps Script).
 *   - appendResponse_ (Code.gs) gọi upsertResponseRow_: MSNV đã có dòng CÙNG THÁNG -> ghi đè dòng đó (giữ vị trí),
 *     "Số lần làm" (trong tháng) +1; chưa có dòng tháng này -> thêm dòng mới, "Số lần làm" = 1.
 *     Ví dụ: làm 3 lần trong tháng 8 -> 1 dòng tháng 8 (Số lần làm 3); sang tháng 10 làm lại -> thêm 1 dòng tháng 10.
 *   - Dòng Stress Test (cột "Stress Test" = YES) không bao giờ ghi đè / bị ghi đè.
 *   - getMyResult_ (action GET getMyResult): điểm của THÁNG HIỆN TẠI cho trang khảo sát (tháng mới chưa làm -> null).
 *   - mergeDuplicateResponsesByMsnv(): chạy MỘT LẦN để gộp dòng trùng đã có (theo MSNV + tháng, giữ lần nộp mới nhất).
 *
 * Bản lưu trong repo core-value-survey: apps-script/SurveyRetake.gs.
 */

const RETAKE_ATTEMPTS_HEADER = 'Số lần làm';
const RETAKE_RESULT_CACHE_SECONDS = 600;

function retakeTimeZone_() {
  return Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh';
}

/** "2026-10" theo giờ VN; '' nếu Timestamp trống / không đọc được. */
function retakeMonthOf_(timestampValue) {
  const date = parseDate_(timestampValue);
  return date ? Utilities.formatDate(date, retakeTimeZone_(), 'yyyy-MM') : '';
}

function retakeCurrentMonth_() {
  return Utilities.formatDate(new Date(), retakeTimeZone_(), 'yyyy-MM');
}

function retakeMsnvKey_(value) {
  return normalizeMSNV_(value).replace(/^'+/, '').toUpperCase();
}

/** Khóa gộp: MSNV + tháng. */
function retakeGroupKey_(msnv, timestampValue) {
  const key = retakeMsnvKey_(msnv);
  return key ? key + '|' + retakeMonthOf_(timestampValue) : '';
}

/*
 * Bộ nhớ điểm của tháng (CacheService), để tra điểm nhanh khi nhiều người cùng mở khảo sát:
 * mỗi tháng một bộ 32 "ngăn" JSON { MSNV: [giờ nộp ISO, điểm, tổng, biết, tổng biết, hiểu, tổng hiểu, số lần làm] },
 * dựng một lần từ sheet Responses (chỉ đọc các cột cần, không đọc nội dung bài làm), cập nhật ngay mỗi lần ghi kết quả.
 */
const RESULT_INDEX_BUCKETS = 32;
const RESULT_INDEX_SECONDS = 21600;

function resultIndexPrefix_(month) {
  return 'vhdn_residx_v1_' + month + '_';
}

function resultIndexMetaKey_(month) {
  return resultIndexPrefix_(month) + 'meta';
}

function resultIndexBucketKey_(month, msnvKey) {
  let hash = 0;
  for (let i = 0; i < msnvKey.length; i++) hash = (hash * 31 + msnvKey.charCodeAt(i)) >>> 0;
  return resultIndexPrefix_(month) + 'b' + (hash % RESULT_INDEX_BUCKETS);
}

/** Giữ tên cũ cho Code.gs: xóa khóa này = bỏ bộ nhớ điểm của tháng (lần tra sau dựng lại). */
function retakeResultCacheKey_(msnv, month) {
  return resultIndexMetaKey_(month || retakeCurrentMonth_());
}

/** Đọc một số cột của Responses (gộp các cột liền nhau thành một lần đọc). Trả về { tên cột: mảng giá trị }. */
function readResponseColumns_(sheet, headers, names) {
  const lastRow = sheet.getLastRow();
  const count = Math.max(lastRow - 1, 0);
  const out = {};
  const wanted = names.map(name => ({ name: name, index: headers.indexOf(name) }));
  wanted.filter(w => w.index < 0).forEach(w => { out[w.name] = new Array(count).fill(''); });
  const present = wanted.filter(w => w.index >= 0).sort((a, b) => a.index - b.index);
  let i = 0;
  while (i < present.length) {
    let j = i;
    while (j + 1 < present.length && present[j + 1].index <= present[j].index + 2) j++;
    const from = present[i].index;
    const to = present[j].index;
    const block = count ? sheet.getRange(2, from + 1, count, to - from + 1).getValues() : [];
    for (let k = i; k <= j; k++) out[present[k].name] = block.map(row => row[present[k].index - from]);
    i = j + 1;
  }
  return out;
}

function resultIndexEntryFromRow_(get) {
  const timestamp = parseDate_(get('Timestamp'));
  return [
    timestamp ? timestamp.toISOString() : String(get('Timestamp') || ''),
    number_(get('Điểm - TỔNG'), 0), number_(get('Tổng câu'), 0),
    number_(get('Điểm - BIẾT'), 0), number_(get('Tổng - BIẾT'), 0),
    number_(get('Điểm - HIỂU'), 0), number_(get('Tổng - HIỂU'), 0),
    Math.max(number_(get(RETAKE_ATTEMPTS_HEADER), 1), 1)
  ];
}

const RESULT_INDEX_COLUMNS = ['MSNV', 'Timestamp', 'Stress Test', 'Điểm - BIẾT', 'Tổng - BIẾT', 'Điểm - HIỂU',
  'Tổng - HIỂU', 'Điểm - TỔNG', 'Tổng câu', RETAKE_ATTEMPTS_HEADER];

/** Dựng bộ nhớ điểm của một tháng từ sheet Responses. */
function buildMonthResultIndex_(month) {
  const cache = CacheService.getScriptCache();
  const startedAt = Date.now();
  const sheet = getSpreadsheet_().getSheetByName(CFG.SHEETS.RESPONSES);
  const buckets = {};
  for (let b = 0; b < RESULT_INDEX_BUCKETS; b++) buckets[resultIndexPrefix_(month) + 'b' + b] = {};
  let entries = 0;

  if (sheet && sheet.getLastRow() >= 2) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(h => String(h || '').trim());
    const cols = readResponseColumns_(sheet, headers, RESULT_INDEX_COLUMNS);
    cols.MSNV.forEach((value, i) => {
      if (isStressValue_(cols['Stress Test'][i])) return;
      const key = retakeMsnvKey_(value);
      if (!key || retakeMonthOf_(cols.Timestamp[i]) !== month) return;
      const bucket = buckets[resultIndexBucketKey_(month, key)];
      if (!bucket[key]) entries++;
      bucket[key] = resultIndexEntryFromRow_(name => cols[name][i]); // dòng dưới cùng = mới nhất
    });
  }

  const values = {};
  Object.keys(buckets).forEach(k => { values[k] = JSON.stringify(buckets[k]); });
  values[resultIndexMetaKey_(month)] = JSON.stringify({ month: month, builtAt: startedAt, entries: entries });
  cache.putAll(values, RESULT_INDEX_SECONDS);

  // Có kết quả mới được ghi trong lúc đang dựng -> bỏ bản vừa dựng, lần sau dựng lại
  const lastWrite = Number(cache.get(resultIndexPrefix_(month) + 'lastWriteAt') || 0);
  if (lastWrite >= startedAt) cache.remove(resultIndexMetaKey_(month));
  return entries;
}

/** Tra một MSNV trong bộ nhớ điểm của tháng; tự dựng nếu chưa có. { ready, entry } */
function readMonthResultIndex_(month, msnvKey) {
  const cache = CacheService.getScriptCache();
  const tryRead = () => {
    const got = cache.getAll([resultIndexMetaKey_(month), resultIndexBucketKey_(month, msnvKey)]);
    const meta = got[resultIndexMetaKey_(month)];
    const bucket = got[resultIndexBucketKey_(month, msnvKey)];
    if (!meta || bucket == null) return null;
    return { ready: true, entry: JSON.parse(bucket)[msnvKey] || null };
  };

  let found = tryRead();
  if (found) return found;

  // Nhiều người cùng tra khi bộ nhớ trống: chỉ một người dựng, người khác chờ tối đa ~8 giây
  const buildingKey = resultIndexPrefix_(month) + 'building';
  if (!cache.get(buildingKey)) {
    cache.put(buildingKey, '1', 30);
    try { buildMonthResultIndex_(month); } finally { cache.remove(buildingKey); }
    found = tryRead();
    if (found) return found;
  } else {
    for (let waited = 0; waited < 8000; waited += 400) {
      Utilities.sleep(400);
      found = tryRead();
      if (found) return found;
    }
  }
  return { ready: false, entry: null };
}

/** Ghi kết quả mới vào bộ nhớ điểm của tháng (gọi trong khóa ghi). */
function updateMonthResultIndex_(month, msnvKey, entry) {
  const cache = CacheService.getScriptCache();
  cache.put(resultIndexPrefix_(month) + 'lastWriteAt', String(Date.now()), RESULT_INDEX_SECONDS);
  if (!cache.get(resultIndexMetaKey_(month))) return; // chưa dựng: lần tra sau dựng từ sheet
  const bucketKey = resultIndexBucketKey_(month, msnvKey);
  const raw = cache.get(bucketKey);
  if (raw == null) { cache.remove(resultIndexMetaKey_(month)); return; }
  const bucket = JSON.parse(raw);
  bucket[msnvKey] = entry;
  cache.put(bucketKey, JSON.stringify(bucket), RESULT_INDEX_SECONDS);
}

function resultFromIndexEntry_(entry) {
  const timestamp = parseDate_(entry[0]);
  return {
    submittedAt: timestamp ? Utilities.formatDate(timestamp, retakeTimeZone_(), 'dd/MM/yyyy HH:mm') : String(entry[0] || ''),
    score: entry[1], total: entry[2],
    knowScore: entry[3], totalKnow: entry[4],
    understandScore: entry[5], totalUnderstand: entry[6],
    percent: entry[2] > 0 ? Math.round(entry[1] / entry[2] * 10000) / 100 : 0,
    attempts: entry[7]
  };
}

function isStressValue_(value) {
  return String(value || '').trim().toUpperCase() === 'YES';
}

/**
 * Ghi MỘT LẦN nhiều kết quả vào Responses (processQueue gọi cho cả đợt): mỗi kết quả ghi đè dòng cùng MSNV + cùng tháng,
 * hoặc thêm dòng mới; "Số lần làm" cộng dồn trong tháng. Đọc các cột khóa một lần cho cả đợt, ghi phần thêm mới một lần.
 * Dòng Stress Test / không MSNV luôn thêm mới. Trả về số kết quả đã ghi.
 */
function appendResponsesBatch_(rowObjects) {
  if (!rowObjects.length) return 0;
  const sheet = ensureResponsesSheet_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    // Tiêu đề: thêm các cột còn thiếu (giống appendResponse_)
    const allKeys = [];
    rowObjects.forEach(obj => Object.keys(obj).forEach(k => { if (allKeys.indexOf(k) < 0) allKeys.push(k); }));
    if (allKeys.indexOf(RETAKE_ATTEMPTS_HEADER) < 0) allKeys.push(RETAKE_ATTEMPTS_HEADER);
    let headers = sheet.getLastColumn() > 0
      ? sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(h => String(h || '').trim())
      : [];
    if (sheet.getLastRow() === 0) headers = [];
    const missing = allKeys.filter(k => headers.indexOf(k) < 0);
    if (missing.length) {
      if (sheet.getMaxColumns() < headers.length + missing.length) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length + missing.length - sheet.getMaxColumns());
      }
      sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
      headers = headers.concat(missing);
    }
    const width = headers.length;
    const valuesFor = obj => headers.map(h => Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : '');
    const attemptsIndex = headers.indexOf(RETAKE_ATTEMPTS_HEADER);

    const cols = readResponseColumns_(sheet, headers, ['MSNV', 'Timestamp', 'Stress Test', RETAKE_ATTEMPTS_HEADER]);
    const target = {};   // nhóm -> { row: số dòng trên sheet } hoặc { append: vị trí trong appendRows }
    const attempts = {}; // nhóm -> số lần đã làm trong tháng
    cols.MSNV.forEach((value, i) => {
      if (isStressValue_(cols['Stress Test'][i])) return;
      const group = retakeGroupKey_(value, cols.Timestamp[i]);
      if (!group) return;
      if (!target[group]) target[group] = { row: i + 2 };
      attempts[group] = (attempts[group] || 0) + Math.max(number_(cols[RETAKE_ATTEMPTS_HEADER][i], 1), 1);
    });

    const overwrites = {}; // số dòng -> giá trị
    const appendRows = [];
    const indexUpdates = [];
    rowObjects.forEach(obj => {
      const group = retakeGroupKey_(obj['MSNV'], obj['Timestamp']);
      if (!group || isStressValue_(obj['Stress Test'])) { appendRows.push(valuesFor(obj)); return; }
      const row = Object.assign({}, obj);
      row[RETAKE_ATTEMPTS_HEADER] = (attempts[group] || 0) + 1;
      attempts[group] = row[RETAKE_ATTEMPTS_HEADER];
      const values = valuesFor(row);
      if (target[group] && target[group].row) overwrites[target[group].row] = values;
      else if (target[group]) appendRows[target[group].append] = values;
      else { target[group] = { append: appendRows.length }; appendRows.push(values); }
      indexUpdates.push({
        month: retakeMonthOf_(obj['Timestamp']), key: retakeMsnvKey_(obj['MSNV']),
        entry: resultIndexEntryFromRow_(name => name === RETAKE_ATTEMPTS_HEADER ? row[RETAKE_ATTEMPTS_HEADER] : obj[name])
      });
    });

    // Ghi đè: gom các dòng liền nhau thành một lần ghi
    const rows = Object.keys(overwrites).map(Number).sort((a, b) => a - b);
    for (let i = 0; i < rows.length;) {
      let j = i;
      while (j + 1 < rows.length && rows[j + 1] === rows[j] + 1) j++;
      sheet.getRange(rows[i], 1, j - i + 1, width).setValues(rows.slice(i, j + 1).map(r => overwrites[r]));
      i = j + 1;
    }

    // Thêm mới: một khối ở cuối (chia đợt 200 dòng)
    if (appendRows.length) {
      const start = Math.max(sheet.getLastRow(), 1) + 1;
      const need = start + appendRows.length - 1;
      if (sheet.getMaxRows() < need) sheet.insertRowsAfter(sheet.getMaxRows(), need - sheet.getMaxRows());
      for (let k = 0; k < appendRows.length; k += 200) {
        const chunk = appendRows.slice(k, k + 200);
        sheet.getRange(start + k, 1, chunk.length, width).setValues(chunk);
      }
    }

    indexUpdates.forEach(u => updateMonthResultIndex_(u.month, u.key, u.entry));
    return rowObjects.length;
  } finally {
    lock.releaseLock();
  }
}

/** Ghi một kết quả (giữ cho appendResponse_ trong Code.gs). */
function upsertResponseRow_(sheet, headers, rowObject) {
  return appendResponsesBatch_([rowObject]);
}

/** Action GET getMyResult: điểm lần nộp gần nhất TRONG THÁNG HIỆN TẠI của một MSNV, result = null nếu tháng này chưa làm. */
function getMyResult_(msnv) {
  const key = retakeMsnvKey_(msnv);
  if (!key) return { ok: false, error: 'Thiếu MSNV.' };
  const month = retakeCurrentMonth_();

  const indexed = readMonthResultIndex_(month, key);
  if (indexed.ready) {
    return { ok: true, month: month, result: indexed.entry ? resultFromIndexEntry_(indexed.entry) : null };
  }

  // Dự phòng (bộ nhớ không dùng được): đọc thẳng sheet
  const sheet = getSpreadsheet_().getSheetByName(CFG.SHEETS.RESPONSES);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, month: month, result: null };
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(h => String(h || '').trim());
  const cols = readResponseColumns_(sheet, headers, RESULT_INDEX_COLUMNS);
  let entry = null;
  cols.MSNV.forEach((value, i) => {
    if (isStressValue_(cols['Stress Test'][i]) || retakeMsnvKey_(value) !== key) return;
    if (retakeMonthOf_(cols.Timestamp[i]) !== month) return;
    entry = resultIndexEntryFromRow_(name => cols[name][i]);
  });
  return { ok: true, month: month, result: entry ? resultFromIndexEntry_(entry) : null };
}

/** Dựng sẵn bộ nhớ điểm tháng này (gọi kèm warmPerformanceCache hoặc chạy tay trước giờ cao điểm). */
function warmMonthResultIndex() {
  const entries = buildMonthResultIndex_(retakeCurrentMonth_());
  return log_('Đã dựng bộ nhớ điểm tháng ' + retakeCurrentMonth_() + ': ' + entries + ' MSNV.');
}

const RETAKE_MERGE_TMP_SHEET = 'Responses (đang gộp)';
const RETAKE_MERGE_CHUNK_ROWS = 200;

/**
 * Chạy MỘT LẦN từ trình soạn thảo: gộp các dòng trùng đã có trong Responses, theo MSNV + THÁNG.
 * Mỗi nhóm (MSNV, tháng) giữ VỊ TRÍ dòng đầu tiên, nội dung = lần nộp MỚI NHẤT của tháng đó, "Số lần làm" = tổng số lần
 * trong tháng. Tháng khác nhau không gộp. Dòng Stress Test và dòng không có MSNV giữ nguyên.
 *
 * Cách làm an toàn với sheet lớn: KHÔNG sửa / xóa gì trên tab cũ. Ghi kết quả vào một tab MỚI theo từng đợt nhỏ,
 * rồi đổi tên: tab cũ -> "Responses (sao lưu trước gộp …)" (chính là bản sao lưu), tab mới -> "Responses".
 * Mỗi bước ghi nhật ký để biết dừng ở đâu nếu có lỗi.
 */
function mergeDuplicateResponsesByMsnv() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CFG.SHEETS.RESPONSES);
  if (!sheet || sheet.getLastRow() < 3) return log_('Không có dòng trùng.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let tmp = null;
  try {
    log_('1/5 Đọc tab ' + sheet.getName() + ': ' + sheet.getLastRow() + ' dòng × ' + sheet.getLastColumn() + ' cột.');
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    let headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(header => String(header || '').trim());
    const msnvIndex = headers.indexOf('MSNV');
    if (msnvIndex < 0) throw new Error('Sheet Responses không có cột MSNV ở dòng 1.');
    const timestampIndex = headers.indexOf('Timestamp');
    const stressIndex = headers.indexOf('Stress Test');
    const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

    const groupOf = row => {
      if (stressIndex >= 0 && isStressValue_(row[stressIndex])) return '';
      return retakeGroupKey_(row[msnvIndex], timestampIndex >= 0 ? row[timestampIndex] : '');
    };
    const groups = {};
    data.forEach((row, i) => {
      const group = groupOf(row);
      if (group) (groups[group] = groups[group] || []).push(i);
    });
    const duplicateGroups = Object.keys(groups).filter(group => groups[group].length > 1);
    const needsAttemptsColumn = headers.indexOf(RETAKE_ATTEMPTS_HEADER) < 0;
    log_('2/5 ' + Object.keys(groups).length + ' nhóm MSNV + tháng, ' + duplicateGroups.length + ' nhóm có dòng trùng.');
    if (!duplicateGroups.length && !needsAttemptsColumn) return log_('Không có dòng trùng trong cùng tháng, không cần gộp.');

    if (needsAttemptsColumn) {
      headers = headers.concat(RETAKE_ATTEMPTS_HEADER);
      data.forEach(row => row.push(''));
    }
    const width = headers.length;
    const attemptsIndex = headers.indexOf(RETAKE_ATTEMPTS_HEADER);

    // Dựng bảng mới trong bộ nhớ
    const firstOf = {};
    Object.keys(groups).forEach(group => { firstOf[groups[group][0]] = group; });
    const output = [];
    data.forEach((row, i) => {
      const group = firstOf[i];
      if (group === undefined) {
        if (!groupOf(row)) output.push(row); // Stress Test / không MSNV: giữ nguyên
        return;                               // dòng trùng đứng sau trong cùng tháng: bỏ
      }
      const rows = groups[group];
      const latest = data[rows[rows.length - 1]].slice();
      latest[attemptsIndex] = rows.reduce((sum, r) => sum + Math.max(number_(data[r][attemptsIndex], 1), 1), 0);
      output.push(latest);
    });

    // Tab tạm: xóa bản dở dang của lần chạy trước (nếu có), tạo mới đủ kích thước
    const leftoverTmp = ss.getSheetByName(RETAKE_MERGE_TMP_SHEET);
    if (leftoverTmp) ss.deleteSheet(leftoverTmp);
    tmp = ss.insertSheet(RETAKE_MERGE_TMP_SHEET);
    if (tmp.getMaxColumns() < width) tmp.insertColumnsAfter(tmp.getMaxColumns(), width - tmp.getMaxColumns());
    if (tmp.getMaxRows() < output.length + 1) tmp.insertRowsAfter(tmp.getMaxRows(), output.length + 1 - tmp.getMaxRows());
    tmp.getRange(1, 1, 1, width).setValues([headers]);
    log_('3/5 Ghi ' + output.length + ' dòng sang tab tạm "' + RETAKE_MERGE_TMP_SHEET + '" (mỗi đợt ' + RETAKE_MERGE_CHUNK_ROWS + ' dòng)…');
    for (let start = 0; start < output.length; start += RETAKE_MERGE_CHUNK_ROWS) {
      const chunk = output.slice(start, start + RETAKE_MERGE_CHUNK_ROWS);
      tmp.getRange(2 + start, 1, chunk.length, width).setValues(chunk);
      SpreadsheetApp.flush();
    }
    tmp.setFrozenRows(1);

    // Đổi tên: tab cũ thành bản sao lưu, tab mới thành Responses
    const backupName = 'Responses (sao lưu trước gộp ' +
      Utilities.formatDate(new Date(), retakeTimeZone_(), 'dd-MM-yyyy HH.mm') + ')';
    log_('4/5 Đổi tên tab cũ thành "' + backupName + '", tab mới thành "' + CFG.SHEETS.RESPONSES + '".');
    sheet.setName(backupName);
    tmp.setName(CFG.SHEETS.RESPONSES);
    tmp = null;
    SpreadsheetApp.flush();

    const cache = CacheService.getScriptCache();
    Object.keys(groups).forEach(group => {
      const parts = group.split('|');
      cache.remove(retakeResultCacheKey_(parts[0], parts[1]));
    });

    return log_('5/5 Xong: gộp ' + duplicateGroups.length + ' nhóm MSNV + tháng (giữ lần nộp mới nhất của tháng), bỏ ' +
      (data.length - output.length) + ' dòng trùng, Responses còn ' + output.length + ' dòng. Tab cũ giữ nguyên tên "' +
      backupName + '" để đối chiếu / khôi phục.');
  } catch (error) {
    log_('LỖI khi gộp: ' + errorMessage_(error) + ' — tab Responses cũ KHÔNG bị sửa.' +
      (tmp ? ' Có thể xóa tab "' + RETAKE_MERGE_TMP_SHEET + '".' : ''));
    throw error;
  } finally {
    lock.releaseLock();
  }
}

/** Chạy từ trình soạn thảo nếu gộp vẫn lỗi: in ra kích thước / đặc điểm của tab Responses để tìm nguyên nhân. */
function diagnoseResponsesSheet() {
  const sheet = getSpreadsheet_().getSheetByName(CFG.SHEETS.RESPONSES);
  if (!sheet) return log_('Không có tab ' + CFG.SHEETS.RESPONSES);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const info = {
    lastRow: lastRow, lastColumn: lastColumn, maxRows: sheet.getMaxRows(), maxColumns: sheet.getMaxColumns(),
    frozenRows: sheet.getFrozenRows(), hasFilter: Boolean(sheet.getFilter()),
    protections: sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE).length + sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).length,
    mergedRanges: lastRow && lastColumn ? sheet.getRange(1, 1, lastRow, lastColumn).getMergedRanges().length : 0
  };
  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    let chars = 0, longest = 0;
    values.forEach(row => row.forEach(v => { const n = String(v).length; chars += n; if (n > longest) longest = n; }));
    info.totalCharacters = chars;
    info.longestCell = longest;
  }
  return log_('Responses: ' + JSON.stringify(info));
}

function log_(message) {
  console.log(message);
  return message;
}

/** Thử nhanh từ trình soạn thảo: đổi MSNV rồi bấm Chạy, xem Nhật ký thực thi. */
function testGetMyResult() {
  console.log(JSON.stringify(getMyResult_('622412005')));
}

/** Xóa dữ liệu thử tải của Claude (run ID "claude-load-test") khỏi Queue và Responses — chạy sau khi thử tải xong. */
function cleanupClaudeLoadTest() {
  return log_(JSON.stringify(cleanupVhdnStressTest_('claude-load-test')));
}
