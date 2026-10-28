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

/** Thử từ trình soạn thảo: xem tóm tắt số liệu báo cáo trong Nhật ký thực thi. */
function testReportData() {
  const data = getReportData_();
  console.log(JSON.stringify(data.months.map(m => ({ month: m.month, source: m.source, total: m.totals.total, pass: m.totals.pass,
    rate: m.totals.rate, hieu: m.totals.understandAvg, chuan: m.benchmark, groups: m.groups.length }))));
  return data;
}
