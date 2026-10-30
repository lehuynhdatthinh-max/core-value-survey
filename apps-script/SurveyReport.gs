/**
 * BÁO CÁO VHDN (thêm 28/10/2026) — số liệu cho tab "Báo cáo VHDN" trong Dashboard quản trị, giống file
 * "Báo cáo khảo sát VHDN - 2026.pptx":
 *   • BIẾT: theo tháng và theo bộ phận — Tổng nhân sự khảo sát, Nhân sự đạt (đúng TẤT CẢ câu BIẾT), Tỷ lệ.
 *   • HIỂU: theo tháng và theo bộ phận — điểm trung bình phần HIỂU; điểm chuẩn = điểm tối đa phần HIỂU của tháng
 *     (tự lấy từ cột "Tổng - HIỂU"), quản trị có thể đặt tay cho từng tháng. Điểm TB "tất cả bộ phận" = trung bình
 *     điểm TB của các bộ phận (như báo cáo).
 *   • Quý: cộng BIẾT của 3 tháng; HIỂU và điểm chuẩn = trung bình các tháng có số liệu.
 * Số liệu tháng cũ (trước khi khảo sát trực tuyến) được NHẬP từ file PowerPoint / Excel báo cáo, lưu ở tab
 * "Báo cáo VHDN (nhập)". Tháng nào có bài khảo sát trực tuyến thì dùng số liệu trực tuyến.
 * Bộ phận được gộp theo nhóm của báo cáo (KHO 7-8-DEPOT, BẢO VỆ ICD 30, XNK, GNQT…); tên lạ giữ nguyên.
 */

const REPORT_IMPORT_SHEET = 'Báo cáo VHDN (nhập)';
const REPORT_IMPORT_HEADERS = ['Tháng', 'Bộ phận', 'BIẾT - Tổng', 'BIẾT - Đạt', 'HIỂU - Điểm TB', 'HIỂU - Điểm tối đa', 'Nguồn', 'Nhập lúc'];
const REPORT_BENCHMARK_PROP = 'VHDN_REPORT_BENCHMARKS';
const REPORT_MAX_IMPORT_ROWS = 3000;

/** Thứ tự nhóm bộ phận như trong báo cáo. */
const REPORT_GROUP_ORDER = [
  'HCNS', 'P.MUA', 'HSSE', 'BẢO TRÌ', 'KINH DOANH', 'KẾ HOẠCH', 'VẬN TẢI', 'CTTT', 'XNK', 'KẾ TOÁN', 'IT',
  'KHO 1-4', 'KHO 5', 'KHO 6', 'KHO 7-8-DEPOT', 'KHO ST',
  'BẢO VỆ ICD 30', 'BẢO VỆ KHO 6', 'BẢO VỆ KHO 7-8', 'BẢO VỆ KHO ST', 'GNQT'
];

function reportFold_(value) {
  return String(value || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/Đ/g, 'D')
    .replace(/[–—_]/g, '-').replace(/\s*-\s*/g, '-').replace(/\s+/g, ' ').trim();
}

/** Tên bộ phận (trong Responses hoặc trong file nhập) → nhóm của báo cáo. */
function reportGroupOf_(department) {
  const raw = String(department || '').trim();
  const f = reportFold_(raw);
  if (!f) return '(Chưa rõ bộ phận)';
  if (f === 'P.MUA' || f === 'PMUA' || f === 'MUA HANG' || f === 'P. MUA') return 'P.MUA';
  if (/^XNK\b/.test(f)) return 'XNK';
  if (f === 'GNQT' || f === 'QNQT' || /GIAO NHAN QUOC TE|FORWARDING/.test(f)) return 'GNQT';
  if (/^KHO 7-8/.test(f) || f === 'KHO 7-8-DEPOT') return 'KHO 7-8-DEPOT';
  if (/^KHO 1-4/.test(f)) return 'KHO 1-4';
  if (/^BAO VE ICD/.test(f)) return 'BẢO VỆ ICD 30';
  if (/^BAO VE (KHO )?ST$/.test(f)) return 'BẢO VỆ KHO ST';
  if (/^BAO VE KHO 7-8/.test(f)) return 'BẢO VỆ KHO 7-8';
  if (/^BAO VE KHO 6$/.test(f)) return 'BẢO VỆ KHO 6';
  for (let i = 0; i < REPORT_GROUP_ORDER.length; i++) {
    if (reportFold_(REPORT_GROUP_ORDER[i]) === f) return REPORT_GROUP_ORDER[i];
  }
  return raw.toUpperCase();
}

function reportGroupSort_(a, b) {
  const ia = REPORT_GROUP_ORDER.indexOf(a), ib = REPORT_GROUP_ORDER.indexOf(b);
  if (ia >= 0 || ib >= 0) return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib);
  return a.localeCompare(b, 'vi');
}

/** "2026-01", "1/2026", "01/2026", "tháng 1/2026", ngày → "2026-01"; '' nếu không đọc được. */
function reportMonthKey_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, retakeTimeZone_(), 'yyyy-MM');
  const text = String(value || '').trim();
  let m = /^(\d{4})-(\d{1,2})/.exec(text);
  if (m) return m[1] + '-' + ('0' + m[2]).slice(-2);
  m = /(\d{1,2})\s*\/\s*(\d{4})/.exec(text);
  if (m && Number(m[1]) >= 1 && Number(m[1]) <= 12) return m[2] + '-' + ('0' + m[1]).slice(-2);
  return '';
}

function reportRound_(value, digits) {
  const p = Math.pow(10, digits == null ? 2 : digits);
  return Math.round(Number(value || 0) * p) / p;
}

/* ---------- Số liệu từ bài khảo sát trực tuyến ---------- */

function buildLiveReportMonths_() {
  const sheet = getSpreadsheet_().getSheetByName(CFG.SHEETS.RESPONSES);
  const months = {};
  if (!sheet || sheet.getLastRow() < 2) return months;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0].map(h => String(h || '').trim());
  const cols = readResponseColumns_(sheet, headers,
    ['MSNV', 'Timestamp', 'Stress Test', 'Bộ phận', 'Điểm - BIẾT', 'Tổng - BIẾT', 'Điểm - HIỂU', 'Tổng - HIỂU']);

  // Mỗi MSNV một bài mỗi tháng (bài mới nhất) — phòng khi còn dòng trùng cũ chưa gộp.
  const latest = {};
  cols.MSNV.forEach((value, i) => {
    if (isStressValue_(cols['Stress Test'][i])) return;
    const key = retakeMsnvKey_(value);
    const month = retakeMonthOf_(cols.Timestamp[i]);
    if (!key || !month) return;
    const time = (parseDate_(cols.Timestamp[i]) || new Date(0)).getTime();
    const id = key + '|' + month;
    if (!latest[id] || latest[id].time <= time) latest[id] = { i: i, time: time, month: month };
  });

  Object.keys(latest).forEach(id => {
    const item = latest[id], i = item.i;
    const month = months[item.month] || (months[item.month] = { groups: {}, maxCounts: {} });
    const group = reportGroupOf_(cols['Bộ phận'][i]);
    const g = month.groups[group] || (month.groups[group] = { total: 0, pass: 0, understandSum: 0 });
    const know = number_(cols['Điểm - BIẾT'][i], 0), knowMax = number_(cols['Tổng - BIẾT'][i], 0);
    const und = number_(cols['Điểm - HIỂU'][i], 0), undMax = number_(cols['Tổng - HIỂU'][i], 0);
    g.total++;
    if (knowMax > 0 && know >= knowMax) g.pass++;
    g.understandSum += und;
    if (undMax > 0) month.maxCounts[undMax] = (month.maxCounts[undMax] || 0) + 1;
  });

  Object.keys(months).forEach(key => {
    const month = months[key];
    // Điểm tối đa phần HIỂU của tháng: giá trị gặp nhiều nhất (đề trong tháng có thể đổi giữa chừng).
    let best = 0, bestCount = -1;
    Object.keys(month.maxCounts).forEach(v => {
      const c = month.maxCounts[v];
      if (c > bestCount || (c === bestCount && Number(v) > best)) { best = Number(v); bestCount = c; }
    });
    months[key] = {
      month: key,
      source: 'live',
      understandMax: best,
      groups: Object.keys(month.groups).map(name => {
        const g = month.groups[name];
        return { group: name, total: g.total, pass: g.pass, understandAvg: g.total ? reportRound_(g.understandSum / g.total, 2) : 0 };
      })
    };
  });
  return months;
}

/* ---------- Số liệu nhập từ file báo cáo ---------- */

function ensureReportImportSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(REPORT_IMPORT_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(REPORT_IMPORT_SHEET);
    sheet.getRange(1, 1, 1, REPORT_IMPORT_HEADERS.length).setValues([REPORT_IMPORT_HEADERS]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readImportedReportMonths_() {
  const sheet = getSpreadsheet_().getSheetByName(REPORT_IMPORT_SHEET);
  const months = {};
  if (!sheet || sheet.getLastRow() < 2) return months;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, REPORT_IMPORT_HEADERS.length).getValues();
  values.forEach(row => {
    const key = reportMonthKey_(row[0]);
    if (!key || !String(row[1] || '').trim()) return;
    const month = months[key] || (months[key] = { month: key, source: 'import', understandMax: 0, fileName: String(row[6] || ''), importedAt: row[7], groups: [] });
    const max = Number(row[5]) || 0;
    if (max > month.understandMax) month.understandMax = max;
    month.groups.push({
      group: reportGroupOf_(row[1]),
      total: Math.round(Number(row[2]) || 0),
      pass: Math.round(Number(row[3]) || 0),
      // Ô trống = bộ phận không có số liệu HIỂU tháng đó (khác với 0 điểm).
      understandAvg: row[4] === '' || row[4] == null ? null : reportRound_(row[4], 2)
    });
  });
  return months;
}

/**
 * Lưu số liệu nhập. rows = [{ month: "2026-01", group, total, pass, understandAvg, understandMax }].
 * Các tháng có trong lần nhập được thay toàn bộ; tháng khác giữ nguyên.
 */
function saveReportImport_(rows, fileName) {
  if (!Array.isArray(rows) || !rows.length) throw new Error('File không có số liệu báo cáo nào.');
  if (rows.length > REPORT_MAX_IMPORT_ROWS) throw new Error('Quá nhiều dòng (' + rows.length + ').');
  const clean = [];
  rows.forEach(r => {
    const month = reportMonthKey_(r.month);
    const group = String(r.group || '').trim();
    if (!month || !group) return;
    clean.push([month, reportGroupOf_(group),
      Math.max(0, Math.round(Number(r.total) || 0)), Math.max(0, Math.round(Number(r.pass) || 0)),
      r.understandAvg === '' || r.understandAvg == null || !isFinite(Number(r.understandAvg)) ? '' : reportRound_(r.understandAvg, 3),
      reportRound_(r.understandMax, 3), String(fileName || '').slice(0, 200), new Date()]);
  });
  if (!clean.length) throw new Error('Không đọc được tháng / bộ phận nào trong file.');
  const months = Array.from(new Set(clean.map(r => r[0])));

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = ensureReportImportSheet_();
    const last = sheet.getLastRow();
    const kept = last >= 2
      ? sheet.getRange(2, 1, last - 1, REPORT_IMPORT_HEADERS.length).getValues().filter(row => months.indexOf(reportMonthKey_(row[0])) === -1)
      : [];
    const all = kept.concat(clean).sort((a, b) => String(a[0]).localeCompare(String(b[0])) || reportGroupSort_(a[1], b[1]));
    if (last >= 2) sheet.getRange(2, 1, last - 1, REPORT_IMPORT_HEADERS.length).clearContent();
    // Cột Tháng ghi dạng chữ để Sheets không đổi "2026-01" thành ngày.
    sheet.getRange(2, 1, all.length, 1).setNumberFormat('@');
    sheet.getRange(2, 1, all.length, REPORT_IMPORT_HEADERS.length).setValues(all);
  } finally {
    lock.releaseLock();
  }
  return { months: months.sort(), rows: clean.length };
}

function deleteReportImportMonth_(monthValue) {
  const month = reportMonthKey_(monthValue);
  if (!month) throw new Error('Tháng không hợp lệ.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSpreadsheet_().getSheetByName(REPORT_IMPORT_SHEET);
    if (!sheet || sheet.getLastRow() < 2) return { removed: 0 };
    const last = sheet.getLastRow();
    const rows = sheet.getRange(2, 1, last - 1, REPORT_IMPORT_HEADERS.length).getValues();
    const kept = rows.filter(row => reportMonthKey_(row[0]) !== month);
    sheet.getRange(2, 1, last - 1, REPORT_IMPORT_HEADERS.length).clearContent();
    if (kept.length) {
      sheet.getRange(2, 1, kept.length, 1).setNumberFormat('@');
      sheet.getRange(2, 1, kept.length, REPORT_IMPORT_HEADERS.length).setValues(kept);
    }
    return { removed: rows.length - kept.length };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Điểm chuẩn HIỂU đặt tay ---------- */

function readReportBenchmarks_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(REPORT_BENCHMARK_PROP) || '{}') || {}; }
  catch (error) { return {}; }
}

/** value trống / null = bỏ đặt tay, quay về điểm tối đa tự động. */
function saveReportBenchmark_(monthValue, value) {
  const month = reportMonthKey_(monthValue);
  if (!month) throw new Error('Tháng không hợp lệ.');
  const all = readReportBenchmarks_();
  if (value === '' || value == null) delete all[month];
  else {
    const n = Number(value);
    if (!isFinite(n) || n <= 0 || n > 100) throw new Error('Điểm chuẩn phải là số lớn hơn 0.');
    all[month] = reportRound_(n, 2);
  }
  PropertiesService.getScriptProperties().setProperty(REPORT_BENCHMARK_PROP, JSON.stringify(all));
  return all;
}

/* ---------- Gộp cho Dashboard ---------- */

/**
 * { months: [{ month, source: 'live'|'import', hasImport, understandMax, benchmark, benchmarkManual,
 *              totals: { total, pass, rate, understandAvg }, groups: [{ group, total, pass, rate, understandAvg }] }],
 *   rate / understandAvg = null khi không có số liệu
 *   groupOrder }
 */
function getReportData_() {
  const live = buildLiveReportMonths_();
  const imported = readImportedReportMonths_();
  const manual = readReportBenchmarks_();
  const keys = Array.from(new Set(Object.keys(live).concat(Object.keys(imported)))).sort();

  const months = keys.map(key => {
    const base = live[key] || imported[key];
    const groups = base.groups.slice().sort((a, b) => reportGroupSort_(a.group, b.group)).map(g => ({
      group: g.group, total: g.total, pass: g.pass,
      rate: g.total ? reportRound_(g.pass / g.total, 4) : null,
      understandAvg: g.understandAvg
    }));
    const total = groups.reduce((s, g) => s + g.total, 0);
    const pass = groups.reduce((s, g) => s + g.pass, 0);
    const scored = groups.filter(g => g.understandAvg != null);
    const hasManual = Object.prototype.hasOwnProperty.call(manual, key);
    return {
      month: key,
      source: base.source,
      hasImport: Boolean(live[key] && imported[key]),
      fileName: base.fileName || '',
      understandMax: base.understandMax,
      benchmark: hasManual ? manual[key] : base.understandMax,
      benchmarkManual: hasManual,
      totals: {
        total: total, pass: pass, rate: total ? reportRound_(pass / total, 4) : null,
        understandAvg: scored.length ? reportRound_(scored.reduce((s, g) => s + g.understandAvg, 0) / scored.length, 2) : null
      },
      groups: groups
    };
  });
  return { months: months, groupOrder: REPORT_GROUP_ORDER.slice(), generatedAt: new Date().toISOString() };
}

/* ---------- Nhân viên chưa làm khảo sát trong tháng ---------- */

/**
 * Người có trong danh sách nhân sự (Sheet1 — đồng bộ từ SafetyHub) nhưng chưa có bài khảo sát trực tuyến nào trong tháng.
 * Bộ phận lấy theo cột "Bộ phận" của Sheet1, gộp nhóm như báo cáo. Danh sách nhân sự là danh sách HIỆN TẠI.
 * { month, employees, done, missing, groups: [{ group, total, done, missing: [{ msnv, name, position }] }] }
 */
function getMissingEmployees_(monthValue) {
  const month = reportMonthKey_(monthValue);
  if (!month) throw new Error('Tháng không hợp lệ.');

  const doneSet = {};
  const resp = getSpreadsheet_().getSheetByName(CFG.SHEETS.RESPONSES);
  if (resp && resp.getLastRow() >= 2) {
    const headers = resp.getRange(1, 1, 1, resp.getLastColumn()).getDisplayValues()[0].map(h => String(h || '').trim());
    const cols = readResponseColumns_(resp, headers, ['MSNV', 'Timestamp', 'Stress Test']);
    cols.MSNV.forEach((value, i) => {
      if (isStressValue_(cols['Stress Test'][i])) return;
      const key = retakeMsnvKey_(value);
      if (key && retakeMonthOf_(cols.Timestamp[i]) === month) doneSet[key] = true;
    });
  }

  const sheet = getSpreadsheet_().getSheetByName(CFG.SHEETS.EMPLOYEES);
  if (!sheet) throw new Error('Không tìm thấy sheet nhân sự: ' + CFG.SHEETS.EMPLOYEES);
  const header = findEmployeeHeaderInfo_(sheet);
  const lastRow = sheet.getLastRow();
  const groups = {};
  let employees = 0, done = 0;
  if (lastRow >= header.dataStartRow) {
    const cols = [header.msnvColumn, header.nameColumn, header.departmentColumn, header.positionColumn];
    const minCol = Math.min.apply(null, cols), maxCol = Math.max.apply(null, cols);
    const rows = sheet.getRange(header.dataStartRow, minCol, lastRow - header.dataStartRow + 1, maxCol - minCol + 1).getDisplayValues();
    const seen = {};
    rows.forEach(row => {
      const msnv = normalizeMSNV_(row[header.msnvColumn - minCol]);
      const key = retakeMsnvKey_(msnv);
      if (!key || seen[key]) return;
      seen[key] = true;
      const group = reportGroupOf_(row[header.departmentColumn - minCol]);
      const g = groups[group] || (groups[group] = { group: group, total: 0, done: 0, missing: [] });
      g.total++; employees++;
      if (doneSet[key]) { g.done++; done++; return; }
      g.missing.push({
        msnv: msnv,
        name: String(row[header.nameColumn - minCol] || '').trim(),
        position: String(row[header.positionColumn - minCol] || '').trim()
      });
    });
  }
  const list = Object.keys(groups).map(k => groups[k]).sort((a, b) => reportGroupSort_(a.group, b.group));
  list.forEach(g => g.missing.sort((a, b) => String(a.msnv).localeCompare(String(b.msnv), 'vi', { numeric: true })));
  return { month: month, employees: employees, done: done, missing: employees - done, groups: list, generatedAt: new Date().toISOString() };
}

/* ---------- Màn hình kết quả: tỷ lệ đã làm + bảng xếp hạng của bộ phận ---------- */

const BOARD_CACHE_SECONDS = 120;
const BOARD_PASS_RATE = 0.5; // bộ phận ĐẠT khi ít nhất 50% nhân viên đã làm (giống tab Báo cáo VHDN)
const BOARD_TOP = 10;
// "Cần xem lại": không xếp hạng (vẫn tính là đã làm) khi làm quá nhanh, hoặc một máy nộp liên tiếp cho nhiều người.
const BOARD_MIN_SECONDS = 60;          // lần làm đầu dưới 60 giây
const BOARD_SHARED_DEVICE_PEOPLE = 3;  // cùng một máy (UserAgent) nộp cho từ 3 người trở lên…
const BOARD_SHARED_GAP_MINUTES = 10;   // …mỗi bài cách bài trước trên máy đó không quá 10 phút

function boardCacheKey_(month, suffix) { return 'vhdn_board_v2_' + month + '_' + suffix; }

function boardApprovedKey_(month) { return 'VHDN_BOARD_APPROVED_' + month; }

/** MSNV quản trị đã duyệt (vẫn xếp hạng dù bị gắn "Cần xem lại"), lưu trong Script Properties theo tháng. */
function readBoardApproved_(month) {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(boardApprovedKey_(month)) || '{}') || {}; } catch (error) { return {}; }
}

/** Quản trị: duyệt / bỏ duyệt một bài "Cần xem lại", rồi dựng lại bảng xếp hạng. */
function saveBoardReview_(monthValue, msnv, approve) {
  const month = reportMonthKey_(monthValue);
  const key = retakeMsnvKey_(msnv);
  if (!month || !key) throw new Error('Thiếu tháng hoặc MSNV.');
  const approved = readBoardApproved_(month);
  if (approve) approved[key] = new Date().toISOString(); else delete approved[key];
  PropertiesService.getScriptProperties().setProperty(boardApprovedKey_(month), JSON.stringify(approved));
  try { CacheService.getScriptCache().remove(boardCacheKey_(month, 'meta')); } catch (error) {}
  return getAdminLeaderboard_(month);
}

/** Lý do "Cần xem lại" của từng MSNV: f = làm quá nhanh, m = số người nộp liên tiếp trên cùng máy. */
function boardReviewFlags_(entries) {
  const flags = {};
  entries.forEach(e => { if (e.d !== '' && e.d < BOARD_MIN_SECONDS) (flags[e.k] = flags[e.k] || {}).f = true; });
  const byDevice = {};
  entries.forEach(e => { if (e.ua) (byDevice[e.ua] = byDevice[e.ua] || []).push(e); });
  Object.keys(byDevice).forEach(ua => {
    const list = byDevice[ua].sort((a, b) => a.at - b.at);
    let run = [];
    const close = () => { if (run.length >= BOARD_SHARED_DEVICE_PEOPLE) run.forEach(e => { (flags[e.k] = flags[e.k] || {}).m = run.length; }); };
    list.forEach(e => {
      if (run.length && e.at - run[run.length - 1].at > BOARD_SHARED_GAP_MINUTES * 60000) { close(); run = []; }
      run.push(e);
    });
    close();
  });
  return flags;
}

/** Số nhân viên theo nhóm bộ phận (Sheet1). */
function readStaffGroupCounts_() {
  const counts = {};
  const sheet = getSpreadsheet_().getSheetByName(CFG.SHEETS.EMPLOYEES);
  if (!sheet) return counts;
  const header = findEmployeeHeaderInfo_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < header.dataStartRow) return counts;
  const minCol = Math.min(header.msnvColumn, header.departmentColumn), maxCol = Math.max(header.msnvColumn, header.departmentColumn);
  const rows = sheet.getRange(header.dataStartRow, minCol, lastRow - header.dataStartRow + 1, maxCol - minCol + 1).getDisplayValues();
  const seen = {};
  rows.forEach(row => {
    const key = retakeMsnvKey_(row[header.msnvColumn - minCol]);
    if (!key || seen[key]) return;
    seen[key] = true;
    const group = reportGroupOf_(row[header.departmentColumn - minCol]);
    counts[group] = (counts[group] || 0) + 1;
  });
  return counts;
}

/** Xếp hạng: tỷ lệ điểm cao hơn trước; bằng điểm thì làm nhanh hơn trước; không có thời gian xếp sau. */
function boardCompare_(a, b) {
  const ra = a.t ? a.s / a.t : 0, rb = b.t ? b.s / b.t : 0;
  if (rb !== ra) return rb - ra;
  const da = a.d === '' || a.d == null ? Infinity : a.d, db = b.d === '' || b.d == null ? Infinity : b.d;
  if (da !== db) return da - db;
  return (a.at || 0) - (b.at || 0);
}

/** Dựng số liệu bảng xếp hạng của cả tháng (một lần đọc Responses + Sheet1), lưu bộ nhớ đệm 2 phút. */
function buildBoardMonth_(month) {
  const groups = {};
  const members = {};
  const resp = getSpreadsheet_().getSheetByName(CFG.SHEETS.RESPONSES);
  if (resp && resp.getLastRow() >= 2) {
    const headers = resp.getRange(1, 1, 1, resp.getLastColumn()).getDisplayValues()[0].map(h => String(h || '').trim());
    const cols = readResponseColumns_(resp, headers, ['MSNV', 'Timestamp', 'Stress Test', 'Bộ phận', 'Họ và tên', 'Điểm - TỔNG', 'Tổng câu',
      RETAKE_ATTEMPTS_HEADER, 'Thời gian làm (giây)', 'UserAgent'].concat(FIRST_ATTEMPT_HEADERS));
    const latest = {};
    cols.MSNV.forEach((value, i) => {
      if (isStressValue_(cols['Stress Test'][i])) return;
      const key = retakeMsnvKey_(value);
      if (!key || retakeMonthOf_(cols.Timestamp[i]) !== month) return;
      const time = (parseDate_(cols.Timestamp[i]) || new Date(0)).getTime();
      if (!latest[key] || latest[key].time <= time) latest[key] = { i: i, time: time };
    });
    const entries = [];
    Object.keys(latest).forEach(key => {
      const i = latest[key].i;
      const group = reportGroupOf_(cols['Bộ phận'][i]);
      const g = groups[group] || (groups[group] = { done: 0, ranking: [], review: [] });
      g.done++;
      members[key] = group;
      const hasFirst = cols['Điểm lần đầu'][i] !== '' && cols['Điểm lần đầu'][i] != null;
      const single = Math.max(number_(cols[RETAKE_ATTEMPTS_HEADER][i], 1), 1) === 1;
      if (!hasFirst && !single) return; // làm lại từ trước khi có cột "lần đầu": không biết điểm lần đầu
      const d = number_(hasFirst ? cols['Thời gian lần đầu (giây)'][i] : cols['Thời gian làm (giây)'][i], 0);
      entries.push({
        g: group,
        ua: String(cols.UserAgent[i] || '').trim(),
        k: key,
        n: String(cols['Họ và tên'][i] || '').trim(),
        s: number_(hasFirst ? cols['Điểm lần đầu'][i] : cols['Điểm - TỔNG'][i], 0),
        t: number_(hasFirst ? cols['Tổng câu lần đầu'][i] : cols['Tổng câu'][i], 0),
        d: d > 0 ? d : '',
        at: latest[key].time
      });
    });
    const flags = boardReviewFlags_(entries);
    const approved = readBoardApproved_(month);
    entries.forEach(e => {
      const g = groups[e.g];
      const r = { k: e.k, n: e.n, s: e.s, t: e.t, d: e.d, at: e.at };
      const f = flags[e.k];
      if (f) r.x = (f.f ? 'f' : '') + (f.m ? 'm' + f.m : '');
      if (f && approved[e.k]) r.ok = 1;
      (f && !approved[e.k] ? g.review : g.ranking).push(r);
    });
  }
  const staff = readStaffGroupCounts_();
  Object.keys(staff).forEach(group => { if (!groups[group]) groups[group] = { done: 0, ranking: [], review: [] }; });
  const names = Object.keys(groups);
  const payload = {};
  names.forEach((group, index) => {
    const g = groups[group];
    g.ranking.sort(boardCompare_);
    g.review.sort((a, b) => a.at - b.at);
    payload[boardCacheKey_(month, 'g' + index)] = JSON.stringify({ group: group, staff: staff[group] || 0, done: g.done, ranking: g.ranking, review: g.review });
  });
  const memberIndex = {};
  Object.keys(members).forEach(key => { memberIndex[key] = names.indexOf(members[key]); });
  payload[boardCacheKey_(month, 'meta')] = JSON.stringify({ groups: names, members: memberIndex, builtAt: Date.now() });
  try { CacheService.getScriptCache().putAll(payload, BOARD_CACHE_SECONDS); } catch (error) { console.warn('Board cache: ' + error); }
  return payload;
}

function readBoardGroup_(month, cache, payload, index) {
  const raw = payload ? payload[boardCacheKey_(month, 'g' + index)] : cache.get(boardCacheKey_(month, 'g' + index));
  return raw ? JSON.parse(raw) : null;
}

/**
 * GET getDepartmentBoard: bộ phận của người vừa làm — đã làm / tổng nhân viên / ĐẠT khi ≥ 50%,
 * top 10 theo điểm lần làm đầu tiên trong tháng (bằng điểm thì nhanh hơn xếp trên), và hạng của chính người đó.
 * Bài vừa nộp có thể còn trong hàng đợi (xử lý mỗi phút): khi đó trả về pending = true, trang tự xếp hạng tạm.
 */
function getDepartmentBoard_(msnv, departmentHint) {
  const month = retakeCurrentMonth_();
  const key = retakeMsnvKey_(msnv);
  const cache = CacheService.getScriptCache();
  let payload = null;
  let metaRaw = cache.get(boardCacheKey_(month, 'meta'));
  if (!metaRaw) {
    const lock = LockService.getScriptLock();
    const locked = lock.tryLock(8000);
    try {
      metaRaw = cache.get(boardCacheKey_(month, 'meta'));
      if (!metaRaw) { payload = buildBoardMonth_(month); metaRaw = payload[boardCacheKey_(month, 'meta')]; }
    } finally {
      if (locked) lock.releaseLock();
    }
  }
  const meta = JSON.parse(metaRaw);
  let index = key && meta.members[key] != null ? meta.members[key] : -1;
  const pending = index < 0;
  if (index < 0) index = meta.groups.indexOf(reportGroupOf_(departmentHint));
  if (index < 0) return { ok: true, month: month, group: reportGroupOf_(departmentHint), staff: 0, done: 0, rate: null, pass: false, top: [], me: null, all: [], pending: true };

  let g = readBoardGroup_(month, cache, payload, index);
  if (!g) { payload = buildBoardMonth_(month); g = readBoardGroup_(month, cache, payload, index); }
  const rank = key ? g.ranking.findIndex(r => r.k === key) : -1;
  const review = !!(key && (g.review || []).some(r => r.k === key));
  const pack = (r, i) => ({ rank: i + 1, name: r.n, score: r.s, total: r.t, seconds: r.d === '' ? null : r.d, me: r.k === key });
  return {
    ok: true,
    month: month,
    group: g.group,
    staff: g.staff,
    done: g.done,
    rate: g.staff ? Math.round(g.done / g.staff * 1000) / 1000 : null,
    pass: g.staff > 0 && g.done / g.staff >= BOARD_PASS_RATE,
    passRate: BOARD_PASS_RATE,
    top: g.ranking.slice(0, BOARD_TOP).map(pack),
    me: rank >= 0 ? pack(g.ranking[rank], rank) : null,
    ranked: g.ranking.length,
    // Điểm / tổng / thời gian của mọi người (không tên) để trang tự xếp hạng tạm khi bài vừa nộp chưa vào sheet.
    all: g.ranking.map(r => [r.s, r.t, r.d === '' ? null : r.d]),
    pending: pending,
    review: review, // bài của người này đang "Cần xem lại" nên chưa xếp hạng
    minSeconds: BOARD_MIN_SECONDS,
    cachedAt: meta.builtAt
  };
}

/**
 * Quản trị: bảng xếp hạng ĐẦY ĐỦ của mọi bộ phận trong một tháng (có MSNV), và top toàn công ty.
 * Cùng cách xếp như màn hình kết quả: điểm lần làm đầu tiên trong tháng, bằng điểm thì nhanh hơn xếp trên.
 */
function getAdminLeaderboard_(monthValue) {
  const month = reportMonthKey_(monthValue) || retakeCurrentMonth_();
  const payload = buildBoardMonth_(month);
  const meta = JSON.parse(payload[boardCacheKey_(month, 'meta')]);
  const all = [];
  const groups = meta.groups.map((name, index) => {
    const g = JSON.parse(payload[boardCacheKey_(month, 'g' + index)]);
    const ranking = g.ranking.map((r, i) => ({ rank: i + 1, msnv: r.k, name: r.n, score: r.s, total: r.t, seconds: r.d === '' ? null : r.d, flags: boardFlagText_(r.x), approved: !!r.ok }));
    const review = (g.review || []).map(r => ({ msnv: r.k, name: r.n, score: r.s, total: r.t, seconds: r.d === '' ? null : r.d, flags: boardFlagText_(r.x),
      submittedAt: Utilities.formatDate(new Date(r.at), retakeTimeZone_(), 'dd/MM/yyyy HH:mm') }));
    g.ranking.forEach(r => all.push(Object.assign({ group: g.group }, r)));
    return {
      group: g.group, staff: g.staff, done: g.done,
      rate: g.staff ? Math.round(g.done / g.staff * 1000) / 1000 : null,
      pass: g.staff > 0 && g.done / g.staff >= BOARD_PASS_RATE,
      unranked: g.done - ranking.length - review.length,
      ranking: ranking,
      review: review
    };
  }).sort((a, b) => reportGroupSort_(a.group, b.group));
  all.sort(boardCompare_);
  return {
    month: month,
    passRate: BOARD_PASS_RATE,
    groups: groups,
    company: all.slice(0, 20).map((r, i) => ({ rank: i + 1, group: r.group, msnv: r.k, name: r.n, score: r.s, total: r.t, seconds: r.d === '' ? null : r.d })),
    ranked: all.length,
    reviewCount: groups.reduce((n, g) => n + g.review.length, 0),
    rules: { minSeconds: BOARD_MIN_SECONDS, devicePeople: BOARD_SHARED_DEVICE_PEOPLE, deviceGapMinutes: BOARD_SHARED_GAP_MINUTES },
    generatedAt: new Date().toISOString()
  };
}

function boardFlagText_(code) {
  if (!code) return [];
  const out = [];
  if (code.indexOf('f') >= 0) out.push('Làm quá nhanh (dưới ' + BOARD_MIN_SECONDS + ' giây)');
  const m = /m(\d+)/.exec(code);
  if (m) out.push('Một máy nộp liên tiếp cho ' + m[1] + ' người');
  return out;
}

/** Thử từ trình soạn thảo: bảng xếp hạng bộ phận của một MSNV. */
function testDepartmentBoard() {
  const data = getDepartmentBoard_('622412005', '');
  console.log(JSON.stringify(data, null, 2));
  return data;
}

/** Thử từ trình soạn thảo: số người chưa làm khảo sát tháng này theo bộ phận. */
function testMissingEmployees() {
  const data = getMissingEmployees_(retakeCurrentMonth_());
  console.log(data.month + ': ' + data.missing + '/' + data.employees + ' chưa làm. ' +
    data.groups.map(g => g.group + ' ' + g.missing.length + '/' + g.total).join(' · '));
  return data;
}

/** Thử từ trình soạn thảo: xem tóm tắt số liệu báo cáo trong Nhật ký thực thi. */
function testReportData() {
  const data = getReportData_();
  console.log(JSON.stringify(data.months.map(m => ({ month: m.month, source: m.source, total: m.totals.total, pass: m.totals.pass,
    rate: m.totals.rate, hieu: m.totals.understandAvg, chuan: m.benchmark, groups: m.groups.length }))));
  return data;
}

/**
 * Chạy từ trình soạn thảo: liệt kê các bài làm tháng này nhanh hơn `maxSeconds` giây (mặc định 60) vào tab
 * "Kiểm tra làm nhanh yyyy-MM" để xem vì sao: lựa chọn, câu tự luận, thiết bị (UserAgent), các bài cùng thiết bị
 * nộp liền nhau, các câu tự luận giống hệt nhau. Chỉ đọc Responses, không sửa điểm.
 */
function auditFastSubmissions(maxSeconds) {
  const limit = Number(maxSeconds) > 0 ? Number(maxSeconds) : 60;
  const month = retakeCurrentMonth_();
  const ss = getSpreadsheet_();
  const resp = ss.getSheetByName(CFG.SHEETS.RESPONSES);
  if (!resp || resp.getLastRow() < 2) return log_('Chưa có bài làm.');
  const headers = resp.getRange(1, 1, 1, resp.getLastColumn()).getDisplayValues()[0].map(h => String(h || '').trim());
  const cols = readResponseColumns_(resp, headers, ['MSNV', 'Timestamp', 'Stress Test', 'Bộ phận', 'Họ và tên', 'Điểm - TỔNG', 'Tổng câu',
    'Answers JSON', 'UserAgent', RETAKE_ATTEMPTS_HEADER, 'Thời gian làm (giây)'].concat(FIRST_ATTEMPT_HEADERS));
  // Câu tự luận = câu không có đáp án để chọn (theo danh sách câu hỏi).
  const textIds = {};
  try { getQuestionConfig_(true).forEach(q => { textIds[q.id] = !(q.options && q.options.length); }); } catch (err) {}
  const all = [];
  cols.MSNV.forEach((value, i) => {
    if (isStressValue_(cols['Stress Test'][i]) || !retakeMsnvKey_(value) || retakeMonthOf_(cols.Timestamp[i]) !== month) return;
    const hasFirst = cols['Điểm lần đầu'][i] !== '' && cols['Điểm lần đầu'][i] != null;
    const seconds = number_(hasFirst ? cols['Thời gian lần đầu (giây)'][i] : cols['Thời gian làm (giây)'][i], 0);
    let answers = {};
    try { answers = JSON.parse(cols['Answers JSON'][i] || '{}'); } catch (err) {}
    const isText = k => textIds[k] || (!(k in textIds) && typeof answers[k] === 'string' && (answers[k].length > 3 || /\s/.test(answers[k])));
    const text = Object.keys(answers).filter(isText).map(k => String(answers[k] || '').trim()).join(' | ');
    const picks = Object.keys(answers).filter(k => !isText(k)).map(k => k + '=' + [].concat(answers[k]).join(',')).join('; ');
    const time = parseDate_(cols.Timestamp[i]);
    all.push({ msnv: retakeMsnvKey_(value), name: String(cols['Họ và tên'][i] || ''), group: reportGroupOf_(cols['Bộ phận'][i]),
      time: time, seconds: seconds, score: (hasFirst ? cols['Điểm lần đầu'][i] : cols['Điểm - TỔNG'][i]) + '/' + (hasFirst ? cols['Tổng câu lần đầu'][i] : cols['Tổng câu'][i]),
      attempts: number_(cols[RETAKE_ATTEMPTS_HEADER][i], 1), ua: String(cols.UserAgent[i] || ''), text: text, picks: picks });
  });
  const fast = all.filter(r => r.seconds > 0 && r.seconds < limit);
  const count = (list, key) => list.reduce((m, r) => { m[r[key]] = (m[r[key]] || 0) + 1; return m; }, {});
  const sameDevice = count(fast, 'ua');
  const normText = r => r.text.toLowerCase().replace(/\s+/g, ' ').trim();
  fast.forEach(r => { r.textKey = normText(r); });
  const sameText = count(fast, 'textKey');
  // Bài nộp ngay trước đó trên cùng thiết bị (theo thời gian nộp): vài phút một bài = một máy chuyền tay / một người làm hộ.
  const byDevice = {};
  all.slice().sort((a, b) => (a.time || 0) - (b.time || 0)).forEach(r => {
    const prev = byDevice[r.ua];
    r.gap = prev && r.time && prev.time ? Math.round((r.time - prev.time) / 6000) / 10 : '';
    r.prevName = prev ? prev.name + ' (' + prev.msnv + ')' : '';
    byDevice[r.ua] = r;
  });
  fast.sort((a, b) => reportGroupSort_(a.group, b.group) || (a.time || 0) - (b.time || 0));
  const name = 'Kiểm tra làm nhanh ' + month;
  const old = ss.getSheetByName(name); if (old) ss.deleteSheet(old);
  const sheet = ss.insertSheet(name);
  const head = ['Bộ phận', 'MSNV', 'Họ và tên', 'Nộp lúc', 'Giây (lần đầu)', 'Điểm lần đầu', 'Số lần làm', 'Câu tự luận', 'Số ký tự tự luận',
    'Số bài nhanh cùng nội dung tự luận', 'Lựa chọn (lần nộp mới nhất)', 'Số bài nhanh cùng thiết bị', 'Bài trước trên cùng thiết bị', 'Cách bài trước (phút)', 'Thiết bị (UserAgent)'];
  const rows = fast.map(r => [r.group, r.msnv, r.name, r.time ? Utilities.formatDate(r.time, retakeTimeZone_(), 'dd/MM/yyyy HH:mm:ss') : '', r.seconds, r.score, r.attempts,
    r.text, r.text.length, sameText[r.textKey], r.picks, sameDevice[r.ua], r.prevName, r.gap, r.ua]);
  sheet.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (rows.length) sheet.getRange(2, 1, rows.length, head.length).setValues(rows);
  sheet.getRange(2, 2, Math.max(rows.length, 1), 1).setNumberFormat('@');
  sheet.setFrozenRows(1);
  const devices = Object.keys(sameDevice).filter(k => sameDevice[k] > 1).length;
  const texts = Object.keys(sameText).filter(k => sameText[k] > 1);
  return log_(month + ': ' + fast.length + '/' + all.length + ' bài làm dưới ' + limit + ' giây. ' +
    Object.entries(count(fast, 'group')).map(e => e[0] + ' ' + e[1]).join(' · ') + '. ' +
    devices + ' thiết bị có từ 2 bài nhanh trở lên; ' + texts.length + ' câu tự luận lặp lại (' + texts.slice(0, 5).map(t => '“' + t.slice(0, 40) + '” ×' + sameText[t]).join(', ') + '). Xem tab "' + name + '".');
}
