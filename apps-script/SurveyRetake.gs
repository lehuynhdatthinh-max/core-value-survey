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
 * Mỗi MSNV giữ VỊ TRÍ dòng đầu tiên, nội dung = lần nộp MỚI NHẤT (dòng dưới cùng), "Số lần làm" = tổng số lần.
 * Dòng Stress Test và dòng không có MSNV giữ nguyên. Tự chép một tab sao lưu trước khi ghi.
 * Đọc cả bảng một lần, ghi lại một lần, xóa phần thừa ở cuối một lần (không xóa từng dòng — dễ gặp lỗi
 * "An unknown error has occurred" khi sheet nhiều dòng).
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
    const stressIndex = headers.indexOf('Stress Test');

    const data = sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();

    // Gom dòng theo MSNV (bỏ qua Stress Test / không có MSNV)
    const groups = {};
    data.forEach((row, i) => {
      if (stressIndex >= 0 && isStressValue_(row[stressIndex])) return;
      const key = retakeMsnvKey_(row[msnvIndex]);
      if (key) (groups[key] = groups[key] || []).push(i);
    });
    const duplicateKeys = Object.keys(groups).filter(key => groups[key].length > 1);
    const needsAttemptsColumn = headers.indexOf(RETAKE_ATTEMPTS_HEADER) < 0;
    if (!duplicateKeys.length && !needsAttemptsColumn) return log_('Không có dòng trùng, không cần gộp.');

    // Sao lưu nguyên tab trước khi sửa
    const backupName = 'Responses (sao lưu trước gộp ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh', 'dd-MM-yyyy HH.mm') + ')';
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

    // Dựng lại bảng: dòng đầu của mỗi MSNV mang nội dung lần mới nhất, các dòng trùng sau bị bỏ
    const firstOf = {};
    Object.keys(groups).forEach(key => { firstOf[groups[key][0]] = key; });
    const output = [];
    data.forEach((row, i) => {
      const key = firstOf[i];
      if (key === undefined) {
        const isGrouped = !(stressIndex >= 0 && isStressValue_(row[stressIndex])) && retakeMsnvKey_(row[msnvIndex]);
        if (!isGrouped) output.push(row); // Stress Test / không MSNV: giữ nguyên
        return;                           // dòng trùng đứng sau: bỏ
      }
      const rows = groups[key];
      const latest = data[rows[rows.length - 1]].slice();
      latest[attemptsIndex] = rows.reduce((sum, r) => sum + Math.max(number_(data[r][attemptsIndex], 1), 1), 0);
      output.push(latest);
    });

    sheet.getRange(2, 1, output.length, lastColumn).setValues(output);
    const leftover = data.length - output.length;
    if (leftover > 0) sheet.deleteRows(2 + output.length, leftover);
    SpreadsheetApp.flush();

    const cache = CacheService.getScriptCache();
    Object.keys(groups).forEach(key => cache.remove(retakeResultCacheKey_(key)));

    return log_('Đã gộp ' + duplicateKeys.length + ' MSNV (giữ lần nộp mới nhất), xóa ' + leftover +
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
