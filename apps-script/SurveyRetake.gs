/**
 * KHẢO SÁT VHDN – XEM ĐIỂM LẦN TRƯỚC + LÀM LẠI (ĐIỂM MỚI GHI ĐÈ ĐIỂM CŨ)   (thêm 27/10/2026)
 *
 * Dán file này vào dự án Apps Script của sheet "phản hồi khảo sát" (cùng dự án với Code.gs đang chạy web app khảo sát).
 * Trang khảo sát (index.html trên GitHub) đã gọi sẵn hành động  getMyResult  sau khi xác nhận MSNV.
 *
 * Cần làm 2 việc trong dự án (xem README trong thư mục apps-script/ của repo):
 *   1) Trong doGet (nơi đang xử lý action 'lookupEmployee'), thêm action 'getMyResult' gọi surveyRetakeGetMyResult_(msnv).
 *      Hàm trả về { ok:true, result:{...} | null } — cùng kiểu với lookupEmployee nên dùng chung cách trả về của lookupEmployee.
 *   2) Chạy setupSurveyRetakeTrigger() MỘT LẦN: tạo trigger 5 phút gộp các dòng trùng MSNV trong sheet Responses —
 *      giữ VỊ TRÍ dòng cũ nhưng ghi đè bằng NỘI DUNG lần nộp mới nhất, rồi xóa các dòng thừa.
 *      Người làm bài luôn thấy đúng điểm mới nhất ngay (getMyResult lấy dòng dưới cùng); trigger chỉ dọn sheet.
 *      (Muốn sheet ghi đè ngay lúc nộp: trong hàm chuyển Queue -> Responses, thay appendRow(values)
 *       bằng surveyRetakeUpsertResponseRow_(sheet, values).)
 *
 * Không đổi cột, không thêm cột: dò dòng tiêu đề và cột theo TÊN tiêu đề (MSNV, Thời gian, Điểm Biết, ...).
 * Viết kiểu ES5 (var / function) cho chạy được cả runtime cũ.
 */

var SURVEY_RETAKE_SHEET_NAMES = ['Responses', 'RESPONSES', 'responses', 'Phản hồi', 'PHẢN HỒI'];
var SURVEY_RETAKE_HEADER_SCAN_ROWS = 10;

// Tên tiêu đề có thể gặp cho từng loại cột (so sánh không dấu, không phân biệt hoa thường, bỏ khoảng trắng / ký tự lạ)
var SURVEY_RETAKE_COLUMNS = {
  msnv: ['msnv', 'masonhanvien', 'manhanvien', 'manv', 'employeeid', 'employeecode'],
  timestamp: ['thoigian', 'timestamp', 'thoigiannop', 'ngaynop', 'ngaygio', 'submittedat', 'createdat', 'thoidiem'],
  knowScore: ['diembiet', 'knowscore', 'phanbiet'],
  totalKnow: ['tongbiet', 'diemtoidabiet', 'totalknow', 'maxbiet'],
  understandScore: ['diemhieu', 'understandscore', 'phanhieu'],
  totalUnderstand: ['tonghieu', 'diemtoidahieu', 'totalunderstand', 'maxhieu'],
  score: ['diemtong', 'tongdiem', 'score', 'diem', 'diemdat'],
  total: ['diemtoida', 'total', 'thangdiem', 'tongdiemtoida', 'maxscore'],
  percent: ['tyle', 'tile', 'percent', 'phantram', 'tyledat']
};

/** Hành động getMyResult: điểm lần nộp gần nhất của một MSNV (null nếu chưa làm). */
function surveyRetakeGetMyResult_(msnv) {
  var key = surveyRetakeNormalizeMsnv_(msnv);
  if (!key) return { ok: false, error: 'Thiếu MSNV.' };

  // Chỉ đọc cột MSNV, rồi đọc đúng 1 dòng cần thiết (nhẹ khi nhiều người cùng mở khảo sát)
  var table = surveyRetakeReadResponses_();
  if (!table) return { ok: true, result: null };

  var latestIndex = -1;
  var attempts = 0;
  for (var i = 0; i < table.msnvs.length; i++) {
    if (surveyRetakeNormalizeMsnv_(table.msnvs[i]) !== key) continue;
    attempts++;
    latestIndex = i; // dòng dưới cùng = lần nộp mới nhất
  }
  if (latestIndex < 0) return { ok: true, result: null };
  var latest = table.sheet.getRange(table.headerRow + 1 + latestIndex, 1, 1, table.width).getValues()[0];

  var c = table.col;
  var num = function (index) {
    if (index < 0) return null;
    var value = latest[index];
    if (value === '' || value === null) return null;
    var n = Number(String(value).replace('%', '').replace(',', '.'));
    return isNaN(n) ? null : n;
  };

  var score = num(c.score);
  var total = num(c.total);
  var knowScore = num(c.knowScore);
  var understandScore = num(c.understandScore);
  // Ô "Điểm tổng" có thể ghi dạng "6/9"
  if (c.score >= 0 && /\//.test(String(latest[c.score]))) {
    var parts = String(latest[c.score]).split('/');
    score = Number(parts[0]);
    if (total === null) total = Number(parts[1]);
  }
  if (score === null && knowScore !== null && understandScore !== null) score = knowScore + understandScore;
  var totalKnow = num(c.totalKnow);
  var totalUnderstand = num(c.totalUnderstand);
  if (total === null && totalKnow !== null && totalUnderstand !== null) total = totalKnow + totalUnderstand;

  var percent = num(c.percent);
  // Ô tỷ lệ định dạng % trong sheet được đọc ra dạng 0.6667 -> đổi về 66.67
  if (percent !== null && percent <= 1 && total && score !== null && Math.abs(percent - score / total) < 0.005) {
    percent = Math.round(percent * 10000) / 100;
  }
  if (percent === null && total && score !== null) percent = Math.round(score / total * 10000) / 100;

  return {
    ok: true,
    result: {
      submittedAt: surveyRetakeFormatTime_(c.timestamp >= 0 ? latest[c.timestamp] : ''),
      score: score,
      total: total,
      knowScore: knowScore,
      totalKnow: totalKnow,
      understandScore: understandScore,
      totalUnderstand: totalUnderstand,
      percent: percent,
      attempts: attempts
    }
  };
}

/**
 * Gộp dòng trùng MSNV trong Responses: với mỗi MSNV có nhiều dòng, ghi nội dung dòng MỚI NHẤT (dưới cùng)
 * vào dòng ĐẦU TIÊN của MSNV đó rồi xóa các dòng còn lại. Dòng không có MSNV giữ nguyên.
 */
function dedupeSurveyResponsesByMsnv() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return 'Đang bận (hàm khác đang ghi sheet), để lần chạy sau.';
  try {
    var table = surveyRetakeReadResponses_();
    if (!table || !table.msnvs.length) return 'Không có dữ liệu.';

    var firstRowByMsnv = {};
    var latestRowByMsnv = {};
    var extraRows = [];
    for (var i = 0; i < table.msnvs.length; i++) {
      var key = surveyRetakeNormalizeMsnv_(table.msnvs[i]);
      if (!key) continue;
      if (firstRowByMsnv[key] === undefined) {
        firstRowByMsnv[key] = i;
      } else {
        extraRows.push(i);
      }
      latestRowByMsnv[key] = i;
    }
    if (!extraRows.length) return 'Không có MSNV trùng.';

    var sheet = table.sheet;
    var width = table.width;
    var firstDataRow = table.headerRow + 1; // số dòng trên sheet của msnvs[0]
    var merged = 0;
    Object.keys(firstRowByMsnv).forEach(function (key) {
      var first = firstRowByMsnv[key];
      var latest = latestRowByMsnv[key];
      if (first === latest) return;
      // Đọc lại bằng getFormulas để không làm mất công thức (nếu có) của dòng mới
      var source = sheet.getRange(firstDataRow + latest, 1, 1, width);
      var formulas = source.getFormulas()[0];
      var values = source.getValues()[0];
      var out = values.map(function (v, j) { return formulas[j] ? formulas[j] : v; });
      sheet.getRange(firstDataRow + first, 1, 1, width).setValues([out]);
      merged++;
    });

    // Xóa từ dưới lên để số dòng phía trên không bị lệch
    extraRows.sort(function (a, b) { return b - a; });
    extraRows.forEach(function (index) { sheet.deleteRow(firstDataRow + index); });

    var message = 'Đã ghi đè ' + merged + ' MSNV bằng lần nộp mới nhất, xóa ' + extraRows.length + ' dòng trùng.';
    console.log(message);
    return message;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Dùng thay cho sheet.appendRow(values) trong hàm chuyển Queue -> Responses (tùy chọn):
 * MSNV đã có dòng -> ghi đè dòng đó; chưa có -> thêm dòng mới. Trả về số dòng đã ghi.
 */
function surveyRetakeUpsertResponseRow_(sheet, values) {
  var table = surveyRetakeReadResponses_(sheet);
  if (table) {
    var key = surveyRetakeNormalizeMsnv_(values[table.col.msnv]);
    if (key) {
      for (var i = table.msnvs.length - 1; i >= 0; i--) {
        if (surveyRetakeNormalizeMsnv_(table.msnvs[i]) === key) {
          var rowNumber = table.headerRow + 1 + i;
          sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
          return rowNumber;
        }
      }
    }
  }
  sheet.appendRow(values);
  return sheet.getLastRow();
}

/** Chạy MỘT LẦN từ trình soạn thảo: trigger 5 phút + gộp ngay các dòng trùng đang có. */
function setupSurveyRetakeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (trigger) {
    if (trigger.getHandlerFunction() === 'dedupeSurveyResponsesByMsnv') ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger('dedupeSurveyResponsesByMsnv').timeBased().everyMinutes(5).create();
  var table = surveyRetakeReadResponses_();
  var info = table
    ? 'Sheet "' + table.sheet.getName() + '", tiêu đề ở dòng ' + table.headerRow + ', cột: ' + JSON.stringify(table.col)
    : 'KHÔNG tìm thấy sheet Responses / cột MSNV — kiểm tra lại tên sheet.';
  console.log(info);
  var result = dedupeSurveyResponsesByMsnv();
  console.log(result);
  return info + '\n' + result;
}

/** Thử nhanh từ trình soạn thảo: đổi MSNV rồi bấm Run, xem kết quả ở Execution log. */
function testSurveyRetakeGetMyResult() {
  console.log(JSON.stringify(surveyRetakeGetMyResult_('622412005')));
}

// ---------------------------------------------------------------- nội bộ

function surveyRetakeReadResponses_(sheetOpt) {
  var sheet = sheetOpt || surveyRetakeFindSheet_();
  if (!sheet) return null;
  var lastRow = sheet.getLastRow();
  var width = sheet.getLastColumn();
  if (lastRow < 1 || width < 1) return null;

  var scanRows = Math.min(SURVEY_RETAKE_HEADER_SCAN_ROWS, lastRow);
  var head = sheet.getRange(1, 1, scanRows, width).getDisplayValues();
  var headerIndex = -1;
  for (var r = 0; r < head.length && headerIndex < 0; r++) {
    for (var k = 0; k < head[r].length; k++) {
      if (SURVEY_RETAKE_COLUMNS.msnv.indexOf(surveyRetakeKey_(head[r][k])) >= 0) { headerIndex = r; break; }
    }
  }
  if (headerIndex < 0) return null;

  var headers = head[headerIndex].map(surveyRetakeKey_);
  var col = {};
  Object.keys(SURVEY_RETAKE_COLUMNS).forEach(function (name) {
    col[name] = -1;
    var names = SURVEY_RETAKE_COLUMNS[name];
    for (var n = 0; n < names.length && col[name] < 0; n++) {
      var at = headers.indexOf(names[n]);
      if (at >= 0) col[name] = at;
    }
  });

  var headerRow = headerIndex + 1;
  var msnvs = lastRow > headerRow
    ? sheet.getRange(headerRow + 1, col.msnv + 1, lastRow - headerRow, 1).getValues().map(function (r) { return r[0]; })
    : [];
  return { sheet: sheet, headerRow: headerRow, width: width, col: col, msnvs: msnvs };
}

function surveyRetakeFindSheet_() {
  var ss = typeof getSpreadsheet_ === 'function' ? getSpreadsheet_() : SpreadsheetApp.getActiveSpreadsheet();
  var configured = PropertiesService.getScriptProperties().getProperty('SURVEY_RESPONSES_SHEET');
  var names = configured ? [configured].concat(SURVEY_RETAKE_SHEET_NAMES) : SURVEY_RETAKE_SHEET_NAMES;
  for (var i = 0; i < names.length; i++) {
    var sheet = ss.getSheetByName(names[i]);
    if (sheet) return sheet;
  }
  return null;
}

function surveyRetakeNormalizeMsnv_(value) {
  return String(value === null || value === undefined ? '' : value).replace(/\s+/g, '').replace(/^'+/, '').toUpperCase();
}

function surveyRetakeKey_(text) {
  return String(text || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/đ/g, 'd').replace(/Đ/g, 'D')
    .toLowerCase()
    .replace(/[^a-z0-9%]/g, '');
}

function surveyRetakeFormatTime_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm');
  }
  return String(value);
}
