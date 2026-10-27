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

function retakeResultCacheKey_(msnv, month) {
  return 'vhdn_my_result_v2_' + retakeMsnvKey_(msnv) + '_' + (month || retakeCurrentMonth_());
}

/** Các cột MSNV / Timestamp / Stress Test / Số lần làm của mọi dòng dữ liệu (chỉ đọc các cột cần thiết). */
function readResponseKeys_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  const count = Math.max(lastRow - 1, 0);
  const column = name => {
    const index = headers.indexOf(name);
    return index >= 0 && count
      ? sheet.getRange(2, index + 1, count, 1).getValues().map(row => row[0])
      : new Array(count).fill('');
  };
  return {
    msnvs: column('MSNV'),
    timestamps: column('Timestamp'),
    stress: column('Stress Test'),
    attempts: column(RETAKE_ATTEMPTS_HEADER)
  };
}

function isStressValue_(value) {
  return String(value || '').trim().toUpperCase() === 'YES';
}

/**
 * Ghi một kết quả vào Responses: ghi đè dòng cùng MSNV + cùng tháng, hoặc thêm dòng mới.
 * headers = tiêu đề dòng 1 (đã gồm mọi cột của rowObject). Trả về số dòng đã ghi.
 */
function upsertResponseRow_(sheet, headers, rowObject) {
  const valuesFor = object => headers.map(header =>
    Object.prototype.hasOwnProperty.call(object, header) ? object[header] : ''
  );
  const group = retakeGroupKey_(rowObject['MSNV'], rowObject['Timestamp']);

  if (!group || isStressValue_(rowObject['Stress Test'])) {
    const rowNumber = sheet.getLastRow() + 1;
    sheet.getRange(rowNumber, 1, 1, headers.length).setValues([valuesFor(rowObject)]);
    return rowNumber;
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(CFG.LOCK_WAIT_MS);
  try {
    const keys = readResponseKeys_(sheet, headers);
    let existing = -1;
    let previousCount = 0;
    for (let i = 0; i < keys.msnvs.length; i++) {
      if (isStressValue_(keys.stress[i])) continue;
      if (retakeGroupKey_(keys.msnvs[i], keys.timestamps[i]) !== group) continue;
      if (existing < 0) existing = i;
      previousCount += Math.max(number_(keys.attempts[i], 1), 1);
    }

    const row = Object.assign({}, rowObject);
    row[RETAKE_ATTEMPTS_HEADER] = previousCount + 1;

    const rowNumber = existing >= 0 ? existing + 2 : sheet.getLastRow() + 1;
    sheet.getRange(rowNumber, 1, 1, headers.length).setValues([valuesFor(row)]);
    CacheService.getScriptCache().remove(
      retakeResultCacheKey_(rowObject['MSNV'], retakeMonthOf_(rowObject['Timestamp']))
    );
    return rowNumber;
  } finally {
    lock.releaseLock();
  }
}

/** Action GET getMyResult: điểm lần nộp gần nhất TRONG THÁNG HIỆN TẠI của một MSNV, result = null nếu tháng này chưa làm. */
function getMyResult_(msnv) {
  const key = retakeMsnvKey_(msnv);
  if (!key) return { ok: false, error: 'Thiếu MSNV.' };

  const month = retakeCurrentMonth_();
  const cache = CacheService.getScriptCache();
  const cacheKey = retakeResultCacheKey_(key, month);
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const sheet = getSpreadsheet_().getSheetByName(CFG.SHEETS.RESPONSES);
  let response = { ok: true, month: month, result: null };

  if (sheet && sheet.getLastRow() >= 2) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getDisplayValues()[0].map(header => String(header || '').trim());
    const keys = readResponseKeys_(sheet, headers);
    const group = key + '|' + month;

    let latest = -1;
    let rows = 0;
    for (let i = 0; i < keys.msnvs.length; i++) {
      if (isStressValue_(keys.stress[i])) continue;
      if (retakeGroupKey_(keys.msnvs[i], keys.timestamps[i]) !== group) continue;
      latest = i; // dòng dưới cùng = mới nhất (trường hợp còn dòng trùng chưa gộp)
      rows++;
    }

    if (latest >= 0) {
      const values = sheet.getRange(latest + 2, 1, 1, headers.length).getValues()[0];
      const get = name => {
        const index = headers.indexOf(name);
        return index < 0 ? '' : values[index];
      };
      const score = number_(get('Điểm - TỔNG'), 0);
      const total = number_(get('Tổng câu'), 0);
      const timestamp = parseDate_(get('Timestamp'));
      response = {
        ok: true,
        month: month,
        result: {
          submittedAt: timestamp ? Utilities.formatDate(timestamp, retakeTimeZone_(), 'dd/MM/yyyy HH:mm') : String(get('Timestamp') || ''),
          score: score,
          total: total,
          knowScore: number_(get('Điểm - BIẾT'), 0),
          totalKnow: number_(get('Tổng - BIẾT'), 0),
          understandScore: number_(get('Điểm - HIỂU'), 0),
          totalUnderstand: number_(get('Tổng - HIỂU'), 0),
          percent: total > 0 ? Math.round(score / total * 10000) / 100 : 0,
          attempts: Math.max(number_(get(RETAKE_ATTEMPTS_HEADER), rows), rows)
        }
      };
    }
  }

  cache.put(cacheKey, JSON.stringify(response), RETAKE_RESULT_CACHE_SECONDS);
  return response;
}

/**
 * Chạy MỘT LẦN từ trình soạn thảo: gộp các dòng trùng đã có trong Responses, theo MSNV + THÁNG.
 * Mỗi nhóm (MSNV, tháng) giữ VỊ TRÍ dòng đầu tiên, nội dung = lần nộp MỚI NHẤT của tháng đó, "Số lần làm" = tổng số lần
 * trong tháng. Tháng khác nhau không gộp. Dòng Stress Test và dòng không có MSNV giữ nguyên. Tự chép một tab sao lưu.
 * Đọc cả bảng một lần, ghi lại một lần, xóa phần thừa ở cuối một lần.
 */
function mergeDuplicateResponsesByMsnv() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(CFG.SHEETS.RESPONSES);
  if (!sheet || sheet.getLastRow() < 3) return log_('Không có dòng trùng.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const lastRow = sheet.getLastRow();
    let lastColumn = sheet.getLastColumn();
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

    // Gom dòng theo (MSNV, tháng)
    const groups = {};
    data.forEach((row, i) => {
      const group = groupOf(row);
      if (group) (groups[group] = groups[group] || []).push(i);
    });
    const duplicateGroups = Object.keys(groups).filter(group => groups[group].length > 1);
    const needsAttemptsColumn = headers.indexOf(RETAKE_ATTEMPTS_HEADER) < 0;
    if (!duplicateGroups.length && !needsAttemptsColumn) return log_('Không có dòng trùng trong cùng tháng, không cần gộp.');

    // Sao lưu nguyên tab trước khi sửa
    const backupName = 'Responses (sao lưu trước gộp ' +
      Utilities.formatDate(new Date(), retakeTimeZone_(), 'dd-MM-yyyy HH.mm') + ')';
    sheet.copyTo(ss).setName(backupName);

    // Cột "Số lần làm": thêm ở cuối nếu chưa có (mở rộng lưới trước để không vượt kích thước sheet)
    if (needsAttemptsColumn) {
      if (sheet.getMaxColumns() < lastColumn + 1) sheet.insertColumnsAfter(sheet.getMaxColumns(), 1);
      sheet.getRange(1, lastColumn + 1).setValue(RETAKE_ATTEMPTS_HEADER);
      headers = headers.concat(RETAKE_ATTEMPTS_HEADER);
      lastColumn += 1;
      data.forEach(row => row.push(''));
    }
    const attemptsIndex = headers.indexOf(RETAKE_ATTEMPTS_HEADER);

    // Dựng lại bảng: dòng đầu của mỗi nhóm mang nội dung lần mới nhất trong tháng, các dòng sau của nhóm bị bỏ
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

    sheet.getRange(2, 1, output.length, lastColumn).setValues(output);
    const leftover = data.length - output.length;
    if (leftover > 0) sheet.deleteRows(2 + output.length, leftover);
    SpreadsheetApp.flush();

    const cache = CacheService.getScriptCache();
    Object.keys(groups).forEach(group => {
      const parts = group.split('|');
      cache.remove(retakeResultCacheKey_(parts[0], parts[1]));
    });

    return log_('Đã gộp ' + duplicateGroups.length + ' nhóm MSNV + tháng (giữ lần nộp mới nhất của tháng), xóa ' + leftover +
      ' dòng trùng. Còn ' + output.length + ' dòng. Bản sao lưu: tab "' + backupName + '".');
  } finally {
    lock.releaseLock();
  }
}

function log_(message) {
  console.log(message);
  return message;
}

/** Thử nhanh từ trình soạn thảo: đổi MSNV rồi bấm Chạy, xem Nhật ký thực thi. */
function testGetMyResult() {
  console.log(JSON.stringify(getMyResult_('622412005')));
}
