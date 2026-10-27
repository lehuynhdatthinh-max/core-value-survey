/**
 * KHẢO SÁT VHDN – XEM ĐIỂM LẦN TRƯỚC + LÀM LẠI (ĐIỂM MỚI GHI ĐÈ ĐIỂM CŨ)   (thêm 27/10/2026)
 *
 * Mỗi MSNV chỉ có MỘT dòng trong sheet Responses, luôn là lần nộp MỚI NHẤT:
 *   - appendResponse_ (Code.gs) gọi upsertResponseRow_: MSNV đã có dòng -> ghi đè ngay dòng đó (giữ vị trí), cột "Số lần làm" +1;
 *     chưa có -> thêm dòng mới với "Số lần làm" = 1.
 *   - Dòng Stress Test (cột "Stress Test" = YES) không bao giờ ghi đè / bị ghi đè, để dọn stress test vẫn chạy như cũ.
 *   - getMyResult_ (action GET getMyResult): trang khảo sát hiện điểm lần trước sau khi xác nhận MSNV.
 *   - mergeDuplicateResponsesByMsnv(): chạy MỘT LẦN để gộp các dòng trùng MSNV đã có từ trước (giữ lần nộp mới nhất).
 *
 * Bản lưu trong repo core-value-survey: apps-script/SurveyRetake.gs.
 */

const RETAKE_ATTEMPTS_HEADER = 'Số lần làm';
const RETAKE_RESULT_CACHE_SECONDS = 600;

function retakeResultCacheKey_(msnv) {
  return 'vhdn_my_result_v1_' + normalizeMSNV_(msnv).toUpperCase();
}

function retakeMsnvKey_(value) {
  return normalizeMSNV_(value).replace(/^'+/, '').toUpperCase();
}

/** Cột MSNV / Stress Test / Số lần làm + MSNV của mọi dòng dữ liệu (chỉ đọc các cột cần thiết). */
function readResponseKeys_(sheet, headers) {
  const msnvIndex = headers.indexOf('MSNV');
  const stressIndex = headers.indexOf('Stress Test');
  const attemptsIndex = headers.indexOf(RETAKE_ATTEMPTS_HEADER);
  const lastRow = sheet.getLastRow();
  const count = Math.max(lastRow - 1, 0);
  const column = index => index >= 0 && count
    ? sheet.getRange(2, index + 1, count, 1).getValues().map(row => row[0])
    : new Array(count).fill('');
  return {
    msnvIndex: msnvIndex,
    attemptsIndex: attemptsIndex,
    msnvs: column(msnvIndex),
    stress: column(stressIndex),
    attempts: column(attemptsIndex)
  };
}

function isStressValue_(value) {
  return String(value || '').trim().toUpperCase() === 'YES';
}

/**
 * Ghi một kết quả vào Responses: ghi đè dòng cũ của cùng MSNV, hoặc thêm dòng mới.
 * headers = tiêu đề dòng 1 (đã gồm mọi cột của rowObject). Trả về số dòng đã ghi.
 */
function upsertResponseRow_(sheet, headers, rowObject) {
  const valuesFor = object => headers.map(header =>
    Object.prototype.hasOwnProperty.call(object, header) ? object[header] : ''
  );
  const key = retakeMsnvKey_(rowObject['MSNV']);

  if (!key || isStressValue_(rowObject['Stress Test'])) {
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
      if (isStressValue_(keys.stress[i]) || retakeMsnvKey_(keys.msnvs[i]) !== key) continue;
      if (existing < 0) existing = i;
      previousCount += Math.max(number_(keys.attempts[i], 1), 1);
    }

    const row = Object.assign({}, rowObject);
    row[RETAKE_ATTEMPTS_HEADER] = previousCount + 1;

    const rowNumber = existing >= 0 ? existing + 2 : sheet.getLastRow() + 1;
    sheet.getRange(rowNumber, 1, 1, headers.length).setValues([valuesFor(row)]);
    CacheService.getScriptCache().remove(retakeResultCacheKey_(key));
    return rowNumber;
  } finally {
    lock.releaseLock();
  }
}

/** Action GET getMyResult: điểm lần nộp gần nhất của một MSNV, result = null nếu chưa làm. */
function getMyResult_(msnv) {
  const key = retakeMsnvKey_(msnv);
  if (!key) return { ok: false, error: 'Thiếu MSNV.' };

  const cache = CacheService.getScriptCache();
  const cached = cache.get(retakeResultCacheKey_(key));
  if (cached) return JSON.parse(cached);

  const sheet = getSpreadsheet_().getSheetByName(CFG.SHEETS.RESPONSES);
  let response = { ok: true, result: null };

  if (sheet && sheet.getLastRow() >= 2) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getDisplayValues()[0].map(header => String(header || '').trim());
    const keys = readResponseKeys_(sheet, headers);

    let latest = -1;
    let rows = 0;
    for (let i = 0; i < keys.msnvs.length; i++) {
      if (isStressValue_(keys.stress[i]) || retakeMsnvKey_(keys.msnvs[i]) !== key) continue;
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
      const timeZone = Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh';
      response = {
        ok: true,
        result: {
          submittedAt: timestamp ? Utilities.formatDate(timestamp, timeZone, 'dd/MM/yyyy HH:mm') : String(get('Timestamp') || ''),
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

  cache.put(retakeResultCacheKey_(key), JSON.stringify(response), RETAKE_RESULT_CACHE_SECONDS);
  return response;
}

/**
 * Chạy MỘT LẦN từ trình soạn thảo: gộp các dòng trùng MSNV đã có trong Responses.
 * Mỗi MSNV giữ VỊ TRÍ dòng đầu tiên, ghi đè bằng NỘI DUNG lần nộp mới nhất (dòng dưới cùng),
 * "Số lần làm" = tổng số lần, rồi xóa các dòng thừa. Dòng Stress Test và dòng không có MSNV giữ nguyên.
 */
function mergeDuplicateResponsesByMsnv() {
  const sheet = ensureResponsesSheet_();
  if (sheet.getLastRow() < 3) return 'Không có dòng trùng.';

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .getDisplayValues()[0].map(header => String(header || '').trim());
    if (headers.indexOf('MSNV') < 0) throw new Error('Sheet Responses không có cột MSNV.');
    if (headers.indexOf(RETAKE_ATTEMPTS_HEADER) < 0) {
      sheet.getRange(1, headers.length + 1).setValue(RETAKE_ATTEMPTS_HEADER);
      headers = headers.concat(RETAKE_ATTEMPTS_HEADER);
    }
    const attemptsColumn = headers.indexOf(RETAKE_ATTEMPTS_HEADER);
    const keys = readResponseKeys_(sheet, headers);

    const groups = {};
    keys.msnvs.forEach((value, i) => {
      if (isStressValue_(keys.stress[i])) return;
      const key = retakeMsnvKey_(value);
      if (!key) return;
      (groups[key] = groups[key] || []).push(i);
    });

    const toDelete = [];
    let merged = 0;
    const cache = CacheService.getScriptCache();
    Object.keys(groups).forEach(key => {
      const rows = groups[key];
      const attempts = rows.reduce((sum, i) => sum + Math.max(number_(keys.attempts[i], 1), 1), 0);
      const first = rows[0];
      const latest = rows[rows.length - 1];
      if (rows.length > 1) {
        const values = sheet.getRange(latest + 2, 1, 1, headers.length).getValues()[0];
        values[attemptsColumn] = attempts;
        sheet.getRange(first + 2, 1, 1, headers.length).setValues([values]);
        rows.slice(1).forEach(i => toDelete.push(i));
        merged++;
      } else if (keys.attempts[first] === '' || keys.attempts[first] == null) {
        sheet.getRange(first + 2, attemptsColumn + 1).setValue(1);
      }
      cache.remove(retakeResultCacheKey_(key));
    });

    toDelete.sort((a, b) => b - a).forEach(i => sheet.deleteRow(i + 2));
    const message = 'Đã gộp ' + merged + ' MSNV (giữ lần nộp mới nhất), xóa ' + toDelete.length + ' dòng trùng.';
    console.log(message);
    return message;
  } finally {
    lock.releaseLock();
  }
}

/** Thử nhanh từ trình soạn thảo: đổi MSNV rồi bấm Chạy, xem Nhật ký thực thi. */
function testGetMyResult() {
  console.log(JSON.stringify(getMyResult_('622412005')));
}
