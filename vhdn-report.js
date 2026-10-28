/*
 * Tab "Báo cáo VHDN" của Dashboard quản trị (nạp khi mở tab lần đầu).
 * Hiển thị số liệu như file "Báo cáo khảo sát VHDN": BIẾT (tổng nhân sự, nhân sự đạt, tỷ lệ) và HIỂU (điểm TB so với
 * điểm chuẩn) theo tháng / quý / năm và theo bộ phận; nhập số liệu tháng cũ từ PowerPoint / Excel; xuất PowerPoint / Excel.
 * Dùng các hàm có sẵn của index.html: state.admin, apiPost, setLoading, escapeHtml.
 */
(function () {
  'use strict';

  const LIBS = {
    pptx: 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js',
    xlsx: 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
    jszip: 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js',
    exceljs: 'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js'
  };
  const COLOR = { total: '#3F78B5', pass: '#E45620', bench: '#4D4D4F' };
  const PASS_RATE = 0.5; // bộ phận ĐẠT khi ít nhất 50% nhân viên đã làm khảo sát trong tháng
  const deptRate = g => g.total ? g.done / g.total : 0;
  const deptPass = g => g.total > 0 && deptRate(g) >= PASS_RATE;
  const BAR_SPAN = 84; // thanh dài nhất chiếm 84% bề ngang, chừa chỗ cho số
  const report = { data: null, period: '', pending: null, missing: {}, missingLoading: '' };
  const $ = id => document.getElementById(id);
  const esc = v => (typeof escapeHtml === 'function' ? escapeHtml(v) : String(v == null ? '' : v));
  const fold = v => String(v || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/Đ/g, 'D').replace(/\s+/g, ' ').trim();

  /* ---------- Định dạng ---------- */
  const num = (v, d) => v == null || !isFinite(v) ? '–' : Number(v).toLocaleString('vi-VN', { maximumFractionDigits: d == null ? 1 : d });
  const pct = v => v == null ? '–' : Math.round(v * 100) + '%';
  function monthParts(key) { const [y, m] = key.split('-'); return { y: Number(y), m: Number(m) }; }
  function monthLabel(key) { const p = monthParts(key); return 'tháng ' + p.m + '/' + p.y; }
  function periodLabel(id) {
    if (id.startsWith('Y')) return 'năm ' + id.slice(1);
    if (id.startsWith('Q')) { const [y, q] = id.slice(1).split('-'); const a = (q - 1) * 3 + 1; return 'Quý ' + q + ' (tháng ' + a + '-' + (a + 2) + ')/' + y; }
    return monthLabel(id);
  }

  /* ---------- Gộp nhiều tháng (quý / năm) — như báo cáo: cộng BIẾT, trung bình HIỂU ---------- */
  function monthsOf(id) {
    const all = (report.data && report.data.months) || [];
    if (id.startsWith('Y')) return all.filter(m => m.month.startsWith(id.slice(1) + '-'));
    if (id.startsWith('Q')) {
      const [y, q] = id.slice(1).split('-').map(Number);
      return all.filter(m => { const p = monthParts(m.month); return p.y === y && Math.ceil(p.m / 3) === q; });
    }
    return all.filter(m => m.month === id);
  }
  function aggregate(months) {
    if (months.length === 1) return Object.assign({ months: months }, months[0]);
    const order = (report.data && report.data.groupOrder) || [];
    const byGroup = {};
    months.forEach(m => m.groups.forEach(g => {
      const x = byGroup[g.group] || (byGroup[g.group] = { group: g.group, total: 0, pass: 0, und: [] });
      x.total += g.total; x.pass += g.pass;
      if (g.understandAvg != null) x.und.push(g.understandAvg);
    }));
    const groups = Object.values(byGroup).map(x => ({
      group: x.group, total: x.total, pass: x.pass, rate: x.total ? x.pass / x.total : null,
      understandAvg: x.und.length ? x.und.reduce((a, b) => a + b, 0) / x.und.length : null
    })).sort((a, b) => {
      const ia = order.indexOf(a.group), ib = order.indexOf(b.group);
      return (ia < 0 ? 999 : ia) - (ib < 0 ? 999 : ib) || a.group.localeCompare(b.group, 'vi');
    });
    const total = groups.reduce((s, g) => s + g.total, 0), pass = groups.reduce((s, g) => s + g.pass, 0);
    const scored = groups.filter(g => g.understandAvg != null);
    const benches = months.map(m => m.benchmark).filter(b => b > 0);
    return {
      months: months, groups: groups,
      benchmark: benches.length ? benches.reduce((a, b) => a + b, 0) / benches.length : 0,
      totals: { total, pass, rate: total ? pass / total : null, understandAvg: scored.length ? scored.reduce((s, g) => s + g.understandAvg, 0) / scored.length : null }
    };
  }
  function periodData(id) { const months = monthsOf(id); return months.length ? aggregate(months) : null; }
  function currentMonth() { const d = new Date(); return d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2); }
  function periodIds() {
    // Tháng hiện tại luôn có trong danh sách (để xem ai chưa làm, kể cả khi chưa có bài nào).
    const months = Array.from(new Set(((report.data && report.data.months) || []).map(m => m.month).concat([currentMonth()]))).sort().reverse();
    const out = [];
    const years = Array.from(new Set(months.map(k => k.slice(0, 4))));
    years.forEach(y => {
      out.push({ id: 'Y' + y, label: 'Cả năm ' + y });
      [4, 3, 2, 1].forEach(q => { if (months.some(k => k.startsWith(y) && Math.ceil(monthParts(k).m / 3) === q)) out.push({ id: 'Q' + y + '-' + q, label: 'Quý ' + q + '/' + y }); });
      months.filter(k => k.startsWith(y)).forEach(k => {
        const m = report.data.months.find(x => x.month === k);
        out.push({ id: k, label: 'Tháng ' + monthParts(k).m + '/' + monthParts(k).y + (!m ? ' (chưa có bài)' : m.source === 'import' ? ' (nhập từ file)' : '') });
      });
    });
    return out;
  }

  /* ---------- Tải số liệu ---------- */
  async function load(showLoader) {
    if (showLoader !== false) setLoading(true);
    try {
      const response = await apiPost('adminReport', Object.assign({}, state.admin));
      setData(response.report);
    } catch (error) {
      notice(error.message || String(error), 'error');
    } finally {
      if (showLoader !== false) setLoading(false);
    }
  }
  function setData(data) {
    report.data = data;
    const ids = periodIds().map(p => p.id);
    if (!ids.includes(report.period)) report.period = currentMonth();
    render();
  }
  function notice(text, kind) { const el = $('reportNotice'); el.textContent = text || ''; el.className = 'message' + (kind ? ' ' + kind : ''); }

  /* ---------- Vẽ ---------- */
  function render() {
    const select = $('reportPeriod');
    const periods = periodIds();
    select.innerHTML = periods.length ? periods.map(p => '<option value="' + esc(p.id) + '"' + (p.id === report.period ? ' selected' : '') + '>' + esc(p.label) + '</option>').join('')
      : '<option value="">Chưa có số liệu</option>';
    const d = report.period ? periodData(report.period) : null;
    $('reportBody').classList.toggle('hidden', !d);
    $('reportEmpty').classList.toggle('hidden', Boolean(d));
    renderImports();
    if (!d) { renderMissing(); return; }
    const label = periodLabel(report.period);
    $('reportTitle').textContent = 'Kết quả khảo sát VHDN ' + label;

    const single = d.months.length === 1 ? d.months[0] : null;
    let source = '';
    if (single) source = single.source === 'import' ? 'Số liệu nhập từ file' + (single.fileName ? ' “' + single.fileName + '”' : '') + '.'
      : 'Số liệu từ bài khảo sát trực tuyến (mỗi người một bài mới nhất trong tháng).' + (single.hasImport ? ' Tháng này cũng có số liệu nhập từ file — đang dùng số liệu trực tuyến.' : '');
    else source = 'Gộp ' + d.months.length + ' tháng: ' + d.months.map(m => monthParts(m.month).m + '/' + monthParts(m.month).y).join(', ') + '. BIẾT cộng dồn; HIỂU và điểm chuẩn lấy trung bình các tháng.';
    $('reportSource').textContent = source;

    $('rTotal').textContent = num(d.totals.total, 0);
    $('rPass').textContent = num(d.totals.pass, 0);
    $('rRate').textContent = pct(d.totals.rate);
    $('rBench').textContent = num(d.benchmark, 2);
    $('rUnd').textContent = num(d.totals.understandAvg, 2);
    $('rBenchSub').textContent = single ? (single.benchmarkManual ? 'Đặt tay (tự động: ' + num(single.understandMax, 2) + ')' : 'Tự động = điểm tối đa phần HIỂU') : 'Trung bình các tháng';

    const bench = $('reportBenchEditor');
    bench.classList.toggle('hidden', !single);
    if (single) {
      $('reportBenchInput').value = single.benchmarkManual ? single.benchmark : '';
      $('reportBenchInput').placeholder = 'Tự động: ' + num(single.understandMax, 2);
      $('reportBenchLabel').textContent = 'Điểm chuẩn HIỂU ' + monthLabel(single.month);
    }

    renderKnowChart(d);
    renderUnderstandChart(d);
    renderTrend();
    renderTable(d);
    renderMissing();
  }

  function legend(items) {
    return '<div class="rc-legend">' + items.map(i => '<span><i style="background:' + i.color + (i.line ? ';height:2px;border-radius:0' : '') + '"></i>' + esc(i.label) + '</span>').join('') + '</div>';
  }

  function renderKnowChart(d) {
    const max = Math.max(1, ...d.groups.map(g => g.total));
    const rows = d.groups.filter(g => g.total > 0);
    $('chartKnow').innerHTML = legend([{ color: COLOR.total, label: 'Tổng nhân sự khảo sát' }, { color: COLOR.pass, label: 'Nhân sự đạt (đúng tất cả câu BIẾT)' }]) +
      (rows.length ? '<div class="rc-row rc-head"><div></div><div></div><div class="rc-rate">Tỷ lệ đạt</div></div>' + rows.map(g => {
        const tip = g.group + ': ' + g.pass + '/' + g.total + ' đạt (' + pct(g.rate) + ')';
        return '<div class="rc-row" title="' + esc(tip) + '"><div class="rc-label">' + esc(g.group) + '</div><div class="rc-track">' +
          '<div class="rc-line"><div class="rc-bar" style="width:' + (g.total / max * BAR_SPAN) + '%;background:' + COLOR.total + '"></div><span class="rc-val">' + g.total + '</span></div>' +
          '<div class="rc-line"><div class="rc-bar" style="width:' + (g.pass / max * BAR_SPAN) + '%;background:' + COLOR.pass + '"></div><span class="rc-val">' + g.pass + '</span></div>' +
          '</div><div class="rc-rate">' + pct(g.rate) + '</div></div>';
      }).join('') : '<div class="rc-empty">Không có số liệu BIẾT.</div>');
  }

  function renderUnderstandChart(d) {
    const rows = d.groups.filter(g => g.understandAvg != null);
    const top = Math.max(d.benchmark || 0, ...rows.map(g => g.understandAvg), 1);
    const benchPos = d.benchmark ? d.benchmark / top * BAR_SPAN : null;
    $('chartUnderstand').innerHTML = legend([{ color: COLOR.pass, label: 'Điểm trung bình phần HIỂU' }, { color: COLOR.bench, label: 'Điểm chuẩn (' + num(d.benchmark, 2) + ' điểm)', line: true }]) +
      (rows.length ? '<div class="rc-plot"><div class="rc-row rc-head"><div></div><div></div><div class="rc-rate">So với chuẩn</div></div>' + rows.map(g => {
        const gap = d.benchmark ? g.understandAvg - d.benchmark : null;
        return '<div class="rc-row" title="' + esc(g.group + ': ' + num(g.understandAvg, 2) + ' / chuẩn ' + num(d.benchmark, 2)) + '"><div class="rc-label">' + esc(g.group) + '</div><div class="rc-track single">' +
          (g.understandAvg / top * BAR_SPAN >= 22
            ? '<div class="rc-line"><div class="rc-bar rc-bar-in" style="width:' + (g.understandAvg / top * BAR_SPAN) + '%;background:' + COLOR.pass + '"><span>' + num(g.understandAvg, 2) + '</span></div></div>'
            : '<div class="rc-line"><div class="rc-bar" style="width:' + (g.understandAvg / top * BAR_SPAN) + '%;background:' + COLOR.pass + '"></div><span class="rc-val">' + num(g.understandAvg, 2) + '</span></div>') +
          (benchPos != null ? '<div class="rc-bench" style="left:' + benchPos + '%"></div>' : '') +
          '</div><div class="rc-rate">' + (gap == null ? '' : (gap >= 0 ? '+' : '−') + num(Math.abs(gap), 2)) + '</div></div>';
      }).join('') + '</div>' : '<div class="rc-empty">Không có số liệu HIỂU.</div>');
  }

  function renderTrend() {
    const months = ((report.data && report.data.months) || []).slice().sort((a, b) => a.month.localeCompare(b.month));
    const year = report.period.slice(report.period.startsWith('Y') || report.period.startsWith('Q') ? 1 : 0, (report.period.startsWith('Y') || report.period.startsWith('Q') ? 1 : 0) + 4);
    const list = months.filter(m => m.month.startsWith(year));
    const col = (m, value, max, text, benchValue) => {
      const h = value == null ? 0 : Math.max(2, value / max * 100);
      const b = benchValue ? benchValue / max * 100 : 0;
      const sel = m.month === report.period ? ' sel' : '';
      return '<div class="rt-col' + sel + '" title="' + esc(monthLabel(m.month) + ': ' + text + (benchValue ? ' · chuẩn ' + num(benchValue, 2) : '')) + '"><div class="rt-area">' +
        '<div class="rt-bar" style="height:' + h + '%"></div>' +
        (benchValue ? '<div class="rt-bench" style="bottom:' + b + '%"></div>' : '') +
        '<span class="rt-val" style="bottom:calc(' + Math.max(h, b) + '% + 3px)">' + esc(text) + '</span>' +
        '</div><div class="rt-label">' + monthParts(m.month).m + '</div></div>';
    };
    const undMax = Math.max(1, ...list.map(m => Math.max(m.totals.understandAvg || 0, m.benchmark || 0))) * 1.2;
    $('chartTrend').innerHTML = list.length ? '<div class="rt-grid"><div><h4>Tỷ lệ đạt BIẾT theo tháng (' + year + ')</h4><div class="rt-chart">' +
      list.map(m => col(m, m.totals.rate == null ? null : m.totals.rate * 100, 115, pct(m.totals.rate))).join('') + '</div></div>' +
      '<div><h4>Điểm TB HIỂU theo tháng (' + year + ')</h4>' + legend([{ color: COLOR.pass, label: 'Điểm TB tất cả bộ phận' }, { color: COLOR.bench, label: 'Điểm chuẩn', line: true }]) + '<div class="rt-chart">' +
      list.map(m => col(m, m.totals.understandAvg, undMax, num(m.totals.understandAvg, 2), m.benchmark)).join('') + '</div></div></div>'
      : '<div class="rc-empty">Chưa có số liệu theo tháng.</div>';
  }

  function renderTable(d) {
    $('reportRows').innerHTML = d.groups.map(g => '<tr><td><strong>' + esc(g.group) + '</strong></td><td class="num">' + num(g.total, 0) + '</td><td class="num">' + num(g.pass, 0) +
      '</td><td class="num">' + pct(g.rate) + '</td><td class="num">' + num(g.understandAvg, 2) + '</td><td class="num">' +
      (g.understandAvg == null || !d.benchmark ? '–' : (g.understandAvg >= d.benchmark ? 'Đạt chuẩn' : 'Dưới chuẩn ' + num(d.benchmark - g.understandAvg, 2))) + '</td></tr>').join('') +
      '<tr class="total"><td><strong>Tất cả bộ phận</strong></td><td class="num">' + num(d.totals.total, 0) + '</td><td class="num">' + num(d.totals.pass, 0) + '</td><td class="num">' + pct(d.totals.rate) +
      '</td><td class="num">' + num(d.totals.understandAvg, 2) + '</td><td class="num">Chuẩn ' + num(d.benchmark, 2) + '</td></tr>';
  }

  function renderImports() {
    const imported = ((report.data && report.data.months) || []).filter(m => m.source === 'import' || m.hasImport);
    $('reportImports').innerHTML = imported.length ? imported.map(m => '<li><span>' + esc('Tháng ' + monthParts(m.month).m + '/' + monthParts(m.month).y) +
      (m.source === 'import' ? (m.fileName ? ' · ' + esc(m.fileName) : '') : ' · đang dùng số liệu trực tuyến') + '</span>' +
      '<button class="btn btn-light btn-xs" type="button" data-del="' + esc(m.month) + '">Xóa số liệu nhập</button></li>').join('')
      : '<li class="muted">Chưa nhập tháng nào.</li>';
  }

  /* ---------- Điểm chuẩn ---------- */
  async function saveBenchmark(useAuto) {
    const d = periodData(report.period);
    if (!d || d.months.length !== 1) return;
    const value = useAuto ? '' : $('reportBenchInput').value.trim();
    if (!useAuto && (!value || !(Number(value) > 0))) { notice('Nhập điểm chuẩn lớn hơn 0, hoặc bấm “Dùng tự động”.', 'error'); return; }
    setLoading(true);
    try {
      const response = await apiPost('adminReportBenchmark', Object.assign({}, state.admin, { month: d.months[0].month, value }));
      setData(response.report);
      notice(useAuto ? 'Đã dùng điểm chuẩn tự động.' : 'Đã lưu điểm chuẩn ' + value + '.', 'success');
    } catch (error) { notice(error.message, 'error'); } finally { setLoading(false); }
  }

  /* ---------- Thư viện ngoài (nạp khi cần) ---------- */
  const loaded = {};
  function loadLib(name, globalName) {
    if (window[globalName]) return Promise.resolve(window[globalName]);
    if (!loaded[name]) loaded[name] = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = LIBS[name]; s.async = true;
      s.onload = () => window[globalName] ? resolve(window[globalName]) : reject(new Error('Không nạp được thư viện ' + name + '.'));
      s.onerror = () => { loaded[name] = null; reject(new Error('Không tải được thư viện ' + name + '. Kiểm tra mạng rồi thử lại.')); };
      document.head.appendChild(s);
    });
    return loaded[name];
  }

  /* ---------- Nhập file ---------- */
  function cleanGroup(text) { return String(text || '').split(/\r?\n/)[0].replace(/\s*\(\s*\d+([.,]\d+)?\s*%\s*\)\s*$/, '').trim(); }
  function monthFromText(text) {
    const m = /TH[AÁ]NG\s*(\d{1,2})\s*\/\s*(\d{4})/i.exec(String(text || '').normalize('NFC')) || /\b(\d{1,2})\s*\/\s*(\d{4})\b/.exec(String(text || ''));
    return m && Number(m[1]) >= 1 && Number(m[1]) <= 12 ? m[2] + '-' + ('0' + m[1]).slice(-2) : '';
  }

  async function parsePptx(file) {
    const JSZip = await loadLib('jszip', 'JSZip');
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const charts = Object.keys(zip.files).filter(n => /^ppt\/charts\/chart\d+\.xml$/.test(n));
    const rows = {};
    const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart', A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
    const texts = (node, ns, tag) => Array.from(node.getElementsByTagNameNS(ns, tag)).map(n => n.textContent);
    const points = node => {
      const out = [];
      if (!node) return out;
      Array.from(node.getElementsByTagNameNS(C, 'pt')).forEach(pt => {
        const v = pt.getElementsByTagNameNS(C, 'v')[0];
        out[Number(pt.getAttribute('idx'))] = v ? v.textContent : '';
      });
      return out;
    };
    let used = 0;
    for (const name of charts) {
      const doc = new DOMParser().parseFromString(await zip.file(name).async('string'), 'application/xml');
      const titleNode = doc.getElementsByTagNameNS(C, 'title')[0];
      const title = titleNode ? texts(titleNode, A, 't').join('') : '';
      const t = fold(title);
      if (!/BO PHAN/.test(t) || /\bQU[IY]\b/.test(t)) continue; // chỉ biểu đồ theo bộ phận của từng tháng
      const month = monthFromText(title);
      const kind = /HIEU/.test(t) ? 'und' : /BIET/.test(t) ? 'know' : '';
      if (!month || !kind) continue;
      const series = Array.from(doc.getElementsByTagNameNS(C, 'ser')).map(ser => {
        const tx = ser.getElementsByTagNameNS(C, 'tx')[0];
        const cat = ser.getElementsByTagNameNS(C, 'cat')[0];
        const val = ser.getElementsByTagNameNS(C, 'val')[0];
        return { name: tx ? fold(texts(tx, C, 'v').join('') || tx.textContent) : '', cats: points(cat), vals: points(val) };
      });
      if (!series.length) continue;
      const cats = series[0].cats;
      const find = re => series.find(s => re.test(s.name));
      const get = (s, i) => s && s.vals[i] !== undefined && s.vals[i] !== '' ? Number(s.vals[i]) : null;
      used++;
      if (kind === 'know') {
        const total = find(/^TONG/) || series[0], pass = find(/^DAT\b|NHAN SU DAT/);
        cats.forEach((c, i) => {
          const g = cleanGroup(c); if (!g) return;
          const r = rows[month + '|' + g] || (rows[month + '|' + g] = { month, group: g, total: 0, pass: 0, understandAvg: null, understandMax: 0 });
          r.total = get(total, i) || 0; r.pass = get(pass, i) || 0;
        });
      } else {
        const max = find(/TOI DA|CHUAN/), avg = series.find(s => s !== max) || series[0];
        const maxValue = max ? Number(max.vals.find(v => v !== undefined && v !== '')) || 0 : 0;
        cats.forEach((c, i) => {
          const g = cleanGroup(c); if (!g) return;
          const r = rows[month + '|' + g] || (rows[month + '|' + g] = { month, group: g, total: 0, pass: 0, understandAvg: null, understandMax: 0 });
          r.understandAvg = get(avg, i); r.understandMax = maxValue;
        });
      }
    }
    // Tháng nào cũng mang điểm tối đa HIỂU của tháng đó cho mọi bộ phận.
    const maxByMonth = {};
    Object.values(rows).forEach(r => { if (r.understandMax) maxByMonth[r.month] = r.understandMax; });
    Object.values(rows).forEach(r => { if (!r.understandMax) r.understandMax = maxByMonth[r.month] || 0; });
    return { rows: Object.values(rows), charts: used };
  }

  async function parseExcel(file) {
    const XLSX = await loadLib('xlsx', 'XLSX');
    const book = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: true });
    const sheetName = book.SheetNames.find(n => /DU LIEU/.test(fold(n))) || book.SheetNames[0];
    const grid = XLSX.utils.sheet_to_json(book.Sheets[sheetName], { header: 1, defval: '', raw: true });
    let head = -1, map = {};
    for (let r = 0; r < Math.min(grid.length, 20) && head < 0; r++) {
      const cells = grid[r].map(fold);
      const at = re => cells.findIndex(c => re.test(c));
      const month = at(/^THANG$|^KY$|^THANG\b/), group = at(/^BO PHAN$/);
      if (month >= 0 && group >= 0) {
        head = r;
        map = { month, group, total: at(/^BIET - TONG|^TONG CONG$|^TONG NHAN SU/), pass: at(/^BIET - DAT|^DAT$|^NHAN SU DAT/),
          avg: at(/^HIEU - DIEM TB|^HIEU - DIEM TRUNG BINH|^DIEM TB/), max: at(/^HIEU - DIEM (TOI DA|CHUAN)|^DIEM (TOI DA|CHUAN)/) };
      }
    }
    if (head < 0) throw new Error('Không thấy dòng tiêu đề có cột “Tháng” và “Bộ phận”. Dùng file Excel xuất từ tab này (sheet “Dữ liệu”).');
    const val = (row, i) => i < 0 || row[i] === '' || row[i] == null ? null : Number(String(row[i]).replace(',', '.'));
    const rows = [];
    grid.slice(head + 1).forEach(row => {
      const raw = row[map.month];
      const month = raw instanceof Date ? raw.getFullYear() + '-' + ('0' + (raw.getMonth() + 1)).slice(-2) : (/^\d{4}-\d{1,2}/.test(String(raw)) ? String(raw).slice(0, 4) + '-' + ('0' + String(raw).split('-')[1]).slice(-2) : monthFromText('tháng ' + raw));
      const group = cleanGroup(row[map.group]);
      if (!month || !group || /^TAT CA/.test(fold(group))) return;
      rows.push({ month, group, total: val(row, map.total) || 0, pass: val(row, map.pass) || 0, understandAvg: val(row, map.avg), understandMax: val(row, map.max) || 0 });
    });
    return { rows, charts: 0, sheet: sheetName };
  }

  async function onFile(event) {
    const file = event.target.files && event.target.files[0];
    event.target.value = '';
    if (!file) return;
    setLoading(true);
    try {
      const isPptx = /\.pptx$/i.test(file.name);
      if (!isPptx && !/\.(xlsx|xls|csv)$/i.test(file.name)) throw new Error('Chọn file PowerPoint (.pptx) hoặc Excel (.xlsx).');
      const parsed = isPptx ? await parsePptx(file) : await parseExcel(file);
      if (!parsed.rows.length) throw new Error(isPptx ? 'Không tìm thấy biểu đồ “… của các bộ phận tháng M/YYYY” nào trong file.' : 'File không có dòng số liệu nào.');
      report.pending = { fileName: file.name, rows: parsed.rows };
      renderPreview(parsed);
    } catch (error) { notice(error.message || String(error), 'error'); } finally { setLoading(false); }
  }

  function renderPreview(parsed) {
    const byMonth = {};
    parsed.rows.forEach(r => { (byMonth[r.month] || (byMonth[r.month] = [])).push(r); });
    const existing = ((report.data && report.data.months) || []);
    const keys = Object.keys(byMonth).sort();
    $('reportPreview').innerHTML = '<div class="report-preview"><h4>Xem trước: ' + esc(report.pending.fileName) + '</h4><table><thead><tr><th>Tháng</th><th>Bộ phận</th><th>BIẾT: tổng / đạt</th><th>HIỂU TB</th><th>Điểm tối đa HIỂU</th><th>Ghi chú</th></tr></thead><tbody>' +
      keys.map(k => {
        const rs = byMonth[k], total = rs.reduce((s, r) => s + r.total, 0), pass = rs.reduce((s, r) => s + r.pass, 0);
        const und = rs.filter(r => r.understandAvg != null), max = Math.max(0, ...rs.map(r => r.understandMax || 0));
        const ex = existing.find(m => m.month === k);
        const note = !ex ? 'Mới' : ex.source === 'live' ? 'Có bài trực tuyến — vẫn dùng số liệu trực tuyến' : 'Thay số liệu nhập trước';
        return '<tr><td>' + esc(monthParts(k).m + '/' + monthParts(k).y) + '</td><td class="num">' + rs.length + '</td><td class="num">' + total + ' / ' + pass + (total ? ' (' + pct(pass / total) + ')' : '') +
          '</td><td class="num">' + (und.length ? num(und.reduce((s, r) => s + r.understandAvg, 0) / und.length, 2) : '–') + '</td><td class="num">' + (max ? num(max, 2) : '–') + '</td><td>' + esc(note) + '</td></tr>';
      }).join('') + '</tbody></table><div class="actions"><button class="btn btn-light" type="button" id="reportCancelImport">Hủy</button><button class="btn btn-primary" type="button" id="reportSaveImport">Lưu ' + keys.length + ' tháng</button></div></div>';
    $('reportCancelImport').onclick = () => { report.pending = null; $('reportPreview').innerHTML = ''; };
    $('reportSaveImport').onclick = saveImport;
    notice('Đọc được ' + parsed.rows.length + ' dòng (' + keys.length + ' tháng). Kiểm tra rồi bấm Lưu.', '');
  }

  async function saveImport() {
    if (!report.pending) return;
    setLoading(true);
    try {
      const response = await apiPost('adminReportImport', Object.assign({}, state.admin, { rows: report.pending.rows, fileName: report.pending.fileName }));
      report.pending = null; $('reportPreview').innerHTML = '';
      setData(response.report);
      notice('Đã lưu số liệu ' + response.result.months.length + ' tháng: ' + response.result.months.map(k => monthParts(k).m + '/' + monthParts(k).y).join(', ') + '.', 'success');
    } catch (error) { notice(error.message, 'error'); } finally { setLoading(false); }
  }

  async function deleteImport(month) {
    if (!confirm('Xóa số liệu nhập của tháng ' + monthParts(month).m + '/' + monthParts(month).y + '?')) return;
    setLoading(true);
    try {
      const response = await apiPost('adminReportDeleteImport', Object.assign({}, state.admin, { month }));
      setData(response.report);
      notice('Đã xóa số liệu nhập tháng ' + monthParts(month).m + '/' + monthParts(month).y + '.', 'success');
    } catch (error) { notice(error.message, 'error'); } finally { setLoading(false); }
  }

  /* ---------- Xuất ---------- */
  function exportYear() { return report.period.replace(/^[YQ]/, '').slice(0, 4); }
  /** Các kỳ của năm theo thứ tự báo cáo: tháng 1, 2, 3, Quý 1, tháng 4… */
  function exportItems(year) {
    const months = ((report.data && report.data.months) || []).map(m => m.month).filter(k => k.startsWith(year + '-')).sort();
    const items = [];
    months.forEach((k, i) => {
      items.push({ id: k, label: 'tháng ' + monthParts(k).m + '/' + year, data: periodData(k) });
      const q = Math.ceil(monthParts(k).m / 3), next = months[i + 1];
      if (!next || Math.ceil(monthParts(next).m / 3) !== q) {
        const qd = periodData('Q' + year + '-' + q);
        if (qd && qd.months.length > 1) items.push({ id: 'Q' + year + '-' + q, label: periodLabel('Q' + year + '-' + q), quarter: true, data: qd });
      }
    });
    return items;
  }

  async function exportExcel() {
    if (!report.data || !report.period) return;
    setLoading(true);
    try {
      const XLSX = await loadLib('xlsx', 'XLSX');
      const year = exportYear(), items = exportItems(year);
      const dataRows = [];
      items.filter(i => !i.quarter).forEach(i => i.data.groups.forEach(g => dataRows.push({
        'Tháng': i.id, 'Bộ phận': g.group, 'BIẾT - Tổng': g.total, 'BIẾT - Đạt': g.pass, 'BIẾT - Tỷ lệ': g.rate == null ? '' : Math.round(g.rate * 1000) / 10,
        'HIỂU - Điểm TB': g.understandAvg == null ? '' : Math.round(g.understandAvg * 100) / 100, 'HIỂU - Điểm tối đa': i.data.benchmark
      })));
      const summary = items.map(i => ({ 'Kỳ': i.label, 'Tổng nhân sự khảo sát': i.data.totals.total, 'Nhân sự đạt BIẾT': i.data.totals.pass,
        'Tỷ lệ đạt (%)': i.data.totals.rate == null ? '' : Math.round(i.data.totals.rate * 1000) / 10,
        'Điểm chuẩn HIỂU': Math.round(i.data.benchmark * 100) / 100, 'Điểm TB HIỂU tất cả bộ phận': i.data.totals.understandAvg == null ? '' : Math.round(i.data.totals.understandAvg * 100) / 100 }));
      const book = XLSX.utils.book_new();
      const s1 = XLSX.utils.json_to_sheet(summary); s1['!cols'] = [{ wch: 28 }, { wch: 22 }, { wch: 18 }, { wch: 14 }, { wch: 16 }, { wch: 26 }];
      const s2 = XLSX.utils.json_to_sheet(dataRows); s2['!cols'] = [{ wch: 10 }, { wch: 18 }, { wch: 12 }, { wch: 11 }, { wch: 12 }, { wch: 15 }, { wch: 18 }];
      XLSX.utils.book_append_sheet(book, s1, 'Tổng hợp');
      XLSX.utils.book_append_sheet(book, s2, 'Dữ liệu');
      items.forEach(i => {
        const rows = i.data.groups.map(g => ({ 'Bộ phận': g.group, 'Tổng cộng': g.total, 'Đạt': g.pass, 'Tỷ lệ (%)': g.rate == null ? '' : Math.round(g.rate * 1000) / 10,
          'Điểm TB HIỂU': g.understandAvg == null ? '' : Math.round(g.understandAvg * 100) / 100, 'Điểm chuẩn': Math.round(i.data.benchmark * 100) / 100 }));
        const name = (i.quarter ? 'Quý ' + i.id.split('-')[1] : 'Tháng ' + monthParts(i.id).m) + '-' + year;
        const s = XLSX.utils.json_to_sheet(rows); s['!cols'] = [{ wch: 18 }, { wch: 11 }, { wch: 8 }, { wch: 10 }, { wch: 13 }, { wch: 11 }];
        XLSX.utils.book_append_sheet(book, s, name.slice(0, 31));
      });
      XLSX.writeFile(book, 'Bao-cao-khao-sat-VHDN-' + year + '.xlsx');
      notice('Đã xuất Excel năm ' + year + ' (' + items.length + ' kỳ). Sheet “Dữ liệu” có thể nhập lại vào tab này.', 'success');
    } catch (error) { notice(error.message || String(error), 'error'); } finally { setLoading(false); }
  }

  async function logoDataUrl() {
    try {
      const blob = await (await fetch('images/logo.png', { cache: 'force-cache' })).blob();
      return await new Promise(resolve => { const r = new FileReader(); r.onload = () => resolve(r.result); r.onerror = () => resolve(''); r.readAsDataURL(blob); });
    } catch (error) { return ''; }
  }

  async function exportPptx() {
    if (!report.data || !report.period) return;
    setLoading(true);
    try {
      const PptxGenJS = await loadLib('pptx', 'PptxGenJS');
      const logo = await logoDataUrl();
      const year = exportYear(), items = exportItems(year);
      const pres = new PptxGenJS();
      pres.layout = 'LAYOUT_WIDE'; // 13.33 x 7.5 in
      pres.title = 'Báo cáo khảo sát VHDN ' + year;
      pres.company = 'TBS Logistics';
      const INK = '1A1A1A', MUTED = '6B6B6E', BLUE = '3F78B5', ORANGE = 'E45620', CHAR = '4D4D4F', GRID = 'DADADC';
      const header = (slide, heading) => {
        if (logo) slide.addImage({ data: logo, x: 0.45, y: 0.3, w: 1.03, h: 0.6 });
        slide.addText('KẾT QUẢ KHẢO SÁT VHDN NĂM ' + year, { x: 1.7, y: 0.32, w: 11, h: 0.5, fontFace: 'Arial', fontSize: 20, bold: true, color: CHAR, margin: 0, isTextBox: true });
        slide.addText(heading, { x: 0.45, y: 1.05, w: 12.4, h: 0.55, fontFace: 'Arial', fontSize: 22, bold: true, color: INK, margin: 0, isTextBox: true });
      };
      const axis = { catAxisLabelColor: MUTED, valAxisLabelColor: MUTED, catAxisLabelFontFace: 'Arial', valAxisLabelFontFace: 'Arial', catAxisLabelFontSize: 10, valAxisLabelFontSize: 10,
        valGridLine: { color: GRID, size: 0.75 }, catGridLine: { style: 'none' }, dataLabelColor: INK, dataLabelFontFace: 'Arial', dataLabelFontSize: 11,
        titleFontFace: 'Arial', titleColor: INK, titleFontSize: 13, legendFontFace: 'Arial', legendFontSize: 11, legendColor: MUTED };

      // 1. Trang bìa
      const cover = pres.addSlide();
      cover.background = { color: 'FFFFFF' };
      if (logo) cover.addImage({ data: logo, x: 0.8, y: 0.7, w: 2.06, h: 1.2 });
      cover.addText('Lan toả VHDN TBSL', { x: 0.8, y: 2.7, w: 11.7, h: 1.2, fontFace: 'Arial', fontSize: 48, bold: true, color: ORANGE, margin: 0, isTextBox: true });
      cover.addText('Báo cáo kết quả khảo sát văn hóa doanh nghiệp “BIẾT” và “HIỂU” năm ' + year, { x: 0.8, y: 4.0, w: 11.7, h: 0.8, fontFace: 'Arial', fontSize: 22, color: CHAR, margin: 0, isTextBox: true });
      cover.addText(items.filter(i => !i.quarter).map(i => i.label.replace('tháng ', 'T')).join(' · '), { x: 0.8, y: 4.9, w: 11.7, h: 0.5, fontFace: 'Arial', fontSize: 14, color: MUTED, margin: 0, isTextBox: true });

      // 2. Tổng hợp BIẾT / HIỂU: 3 kỳ mỗi trang như báo cáo
      const chunks = []; for (let i = 0; i < items.length; i += 3) chunks.push(items.slice(i, i + 3));
      const summarySlides = (kind) => chunks.forEach(group => {
        const slide = pres.addSlide();
        header(slide, kind === 'know' ? 'Kết quả khảo sát việc truyền thông “BIẾT” của các bộ phận:' : 'Kết quả khảo sát “HIỂU” — điểm trung bình so với điểm chuẩn:');
        group.forEach((it, k) => {
          const d = it.data, x = 0.45 + k * 4.2;
          const data = kind === 'know'
            ? [{ name: 'Tổng nhân sự khảo sát', labels: [''], values: [d.totals.total] }, { name: 'Nhân sự đạt', labels: [''], values: [d.totals.pass] }]
            : [{ name: 'Điểm chuẩn', labels: [''], values: [Math.round(d.benchmark * 100) / 100] }, { name: 'Điểm TB tất cả bộ phận', labels: [''], values: [Math.round((d.totals.understandAvg || 0) * 100) / 100] }];
          slide.addChart(pres.charts.BAR, data, Object.assign({}, axis, {
            x, y: 1.9, w: 4.0, h: 5.1, barDir: 'col', barGapWidthPct: 60,
            chartColors: kind === 'know' ? [BLUE, ORANGE] : [CHAR, ORANGE],
            showTitle: true, title: (kind === 'know' ? '“BIẾT” ' : '“HIỂU” ') + it.label.toUpperCase() + (kind === 'know' ? ' · TỶ LỆ ĐẠT ' + pct(d.totals.rate) : ''),
            showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: kind === 'know' ? '#,##0' : '0.0#',
            showLegend: true, legendPos: 'b', valAxisMinVal: 0, catAxisHidden: true
          }));
        });
      });
      summarySlides('know');
      summarySlides('und');

      // 3. Theo bộ phận — tên biểu đồ giống báo cáo để có thể nhập lại file này
      const deptKnow = it => {
        const slide = pres.addSlide(), d = it.data, groups = d.groups.filter(g => g.total > 0);
        header(slide, 'Kết quả khảo sát việc truyền thông “BIẾT” của các bộ phận:');
        slide.addChart(pres.charts.BAR, [
          { name: 'Tổng cộng', labels: groups.map(g => g.group + ' (' + pct(g.rate) + ')'), values: groups.map(g => g.total) },
          { name: 'Đạt', labels: groups.map(g => g.group + ' (' + pct(g.rate) + ')'), values: groups.map(g => g.pass) }
        ], Object.assign({}, axis, {
          x: 0.45, y: 1.75, w: 12.45, h: 5.45, barDir: 'col', barGapWidthPct: 40, chartColors: [BLUE, ORANGE],
          showTitle: true, title: 'Kết quả khảo sát việc truyền thông “BIẾT” của các bộ phận ' + it.label + ' · tỷ lệ đạt chung ' + pct(d.totals.rate),
          showValue: true, dataLabelPosition: 'outEnd', dataLabelFontSize: 9, catAxisLabelFontSize: 9, catAxisLabelRotate: -45, showLegend: true, legendPos: 'b', valAxisMinVal: 0
        }));
      };
      const deptUnd = it => {
        const slide = pres.addSlide(), d = it.data, groups = d.groups.filter(g => g.understandAvg != null);
        const b = Math.round(d.benchmark * 100) / 100, labels = groups.map(g => g.group);
        header(slide, 'Kết quả khảo sát “HIỂU” của các bộ phận:');
        slide.addChart([
          { type: pres.charts.BAR, data: [{ name: 'Tổng cộng', labels, values: groups.map(g => Math.round(g.understandAvg * 100) / 100) }],
            options: { barDir: 'col', barGapWidthPct: 50, chartColors: [ORANGE], showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.0#' } },
          { type: pres.charts.LINE, data: [{ name: 'Điểm tối đa (' + b + ' điểm)', labels, values: labels.map(() => b) }],
            options: { chartColors: [CHAR], lineSize: 2, lineDataSymbol: 'none', showValue: false } }
        ], Object.assign({}, axis, {
          x: 0.45, y: 1.75, w: 12.45, h: 5.45, showTitle: true,
          title: 'Điểm trung bình phần “HIỂU” các bộ phận ' + it.label + ' · điểm chuẩn ' + b + ' · TB tất cả bộ phận ' + num(d.totals.understandAvg, 2),
          catAxisLabelFontSize: 9, catAxisLabelRotate: -45, dataLabelFontSize: 9, showLegend: true, legendPos: 'b', valAxisMinVal: 0
        }));
      };
      items.forEach(deptKnow);
      items.forEach(deptUnd);

      const end = pres.addSlide();
      if (logo) end.addImage({ data: logo, x: 5.63, y: 2.2, w: 2.06, h: 1.2 });
      end.addText('THANK YOU', { x: 0.5, y: 3.8, w: 12.33, h: 1, fontFace: 'Arial', fontSize: 40, bold: true, color: ORANGE, align: 'center', margin: 0, isTextBox: true });

      await pres.writeFile({ fileName: 'Bao-cao-khao-sat-VHDN-' + year + '.pptx' });
      notice('Đã xuất PowerPoint năm ' + year + ' (' + items.length + ' kỳ). File này có thể nhập lại vào tab này.', 'success');
    } catch (error) { notice(error.message || String(error), 'error'); } finally { setLoading(false); }
  }


  /* ---------- Nhân viên chưa làm khảo sát trong tháng ---------- */
  function selectedMonth() { return /^\d{4}-\d{2}$/.test(report.period) ? report.period : ''; }

  async function loadMissing(month, force) {
    if (!force && (report.missing[month] || report.missingLoading === month)) return;
    report.missingLoading = month;
    renderMissing();
    try {
      const response = await apiPost('adminReportMissing', Object.assign({}, state.admin, { month }));
      report.missing[month] = response.missing;
    } catch (error) {
      report.missing[month] = { error: error.message || String(error) };
    } finally {
      if (report.missingLoading === month) report.missingLoading = '';
      if (selectedMonth() === month) renderMissing();
    }
  }

  function renderMissing() {
    const month = selectedMonth();
    const box = $('missingList'), summary = $('missingSummary');
    $('missingExport').disabled = true;
    if (!month) {
      $('missingTitle').textContent = 'Nhân viên chưa làm khảo sát';
      summary.textContent = 'Chọn một tháng ở “Kỳ báo cáo” để xem danh sách nhân viên chưa làm khảo sát.';
      box.innerHTML = ''; return;
    }
    $('missingTitle').textContent = 'Nhân viên chưa làm VHDN ' + monthLabel(month);
    const data = report.missing[month];
    if (!data) { summary.textContent = 'Đang tải danh sách…'; box.innerHTML = ''; loadMissing(month); return; }
    if (data.error) { summary.textContent = 'Không tải được danh sách: ' + data.error; box.innerHTML = ''; return; }
    const info = ((report.data && report.data.months) || []).find(m => m.month === month);
    const offline = !info || info.source !== 'live';
    summary.innerHTML = '<strong>' + num(data.missing, 0) + '</strong> / ' + num(data.employees, 0) + ' nhân viên chưa làm (' + pct(data.employees ? data.missing / data.employees : null) +
      '), ' + num(data.done, 0) + ' người đã làm. So theo danh sách nhân sự hiện tại (Sheet1, đồng bộ từ SafetyHub).' +
      '<br><strong>' + data.groups.filter(deptPass).length + ' / ' + data.groups.length + '</strong> bộ phận đạt (ít nhất ' + Math.round(PASS_RATE * 100) + '% nhân viên đã làm)' +
      (data.groups.some(g => !deptPass(g)) ? ' · Chưa đạt: ' + data.groups.filter(g => !deptPass(g)).map(g => esc(g.group)).join(', ') : '') + '.' +
      (offline ? ' <span class="missing-warn">Tháng này chưa có bài khảo sát trực tuyến nên mọi người đều tính là chưa làm.</span>' : '');
    const q = fold($('missingSearch').value);
    const groups = data.groups.map(g => Object.assign({}, g, { shown: q ? g.missing.filter(p => fold(p.msnv + ' ' + p.name).includes(q)) : g.missing }))
      .filter(g => !q || g.shown.length);
    box.innerHTML = groups.length ? groups.map(g => {
      const badge = '<span class="mg-badge ' + (deptPass(g) ? 'pass' : 'fail') + '">' + (deptPass(g) ? 'Đạt' : 'Chưa đạt') + ' · đã làm ' + pct(deptRate(g)) + '</span>';
      const head = '<summary><span class="mg-name">' + esc(g.group) + '</span>' + badge + '<span class="mg-count">' +
        (g.missing.length ? '<strong>' + g.missing.length + '</strong> / ' + g.total + ' chưa làm' : 'Đã làm đủ ' + g.total + ' người') + '</span></summary>';
      if (!g.missing.length) return '<details class="mg done">' + head + '</details>';
      return '<details class="mg"' + (q ? ' open' : '') + '>' + head + '<table><thead><tr><th class="mg-stt">STT</th><th class="mg-msnv">MSNV</th><th>Họ và tên</th></tr></thead><tbody>' +
        g.shown.map((p, i) => '<tr><td class="mg-stt">' + (i + 1) + '</td><td class="mg-msnv">' + esc(p.msnv) + '</td><td>' + esc(p.name) + '</td></tr>').join('') + '</tbody></table></details>';
    }).join('') : '<div class="rc-empty">Không tìm thấy nhân viên phù hợp.</div>';
    $('missingExport').disabled = !data.missing;
  }

  /** Tên sheet Excel: tối đa 31 ký tự, không có : \ / ? * [ ], không trùng. */
  function sheetName(name, used) {
    let base = String(name || 'Bộ phận').replace(/[:\\\/?*\[\]]/g, '-').slice(0, 31).trim() || 'Bộ phận';
    let out = base, n = 2;
    while (used.has(out.toLowerCase())) { const tail = ' (' + n++ + ')'; out = base.slice(0, 31 - tail.length) + tail; }
    used.add(out.toLowerCase());
    return out;
  }

  async function exportMissing() {
    const month = selectedMonth(), data = report.missing[month];
    if (!month || !data || data.error || !data.missing) return;
    setLoading(true);
    try {
      const ExcelJS = await loadLib('exceljs', 'ExcelJS');
      const p = monthParts(month), label = ('0' + p.m).slice(-2) + '/' + p.y;
      const title = 'DANH SÁCH NHÂN VIÊN CHƯA LÀM VHDN THÁNG ' + label;
      const ORANGE = 'FFE45620', CHAR = 'FF4D4D4F', TINT = 'FFFFF6F1', LINE = 'FFDADADC', WHITE = 'FFFFFFFF', INK = 'FF1A1A1A';
      const font = (o) => Object.assign({ name: 'Arial', size: 11, color: { argb: INK } }, o || {});
      const fill = argb => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });
      const border = { top: { style: 'thin', color: { argb: LINE } }, left: { style: 'thin', color: { argb: LINE } }, bottom: { style: 'thin', color: { argb: LINE } }, right: { style: 'thin', color: { argb: LINE } } };
      const stamp = new Date().toLocaleString('vi-VN');

      const book = new ExcelJS.Workbook();
      book.creator = 'TBS Logistics · HSSE'; book.created = new Date();
      const used = new Set();

      // Một sheet tổng hợp số người chưa làm theo bộ phận, rồi mỗi bộ phận một sheet.
      const layout = (ws, lastCol, subtitle, headers, widths) => {
        ws.columns = widths.map(w => ({ width: w }));
        ws.mergeCells(1, 1, 1, lastCol);
        const t = ws.getCell(1, 1);
        t.value = title; t.font = font({ size: 14, bold: true, color: { argb: WHITE } }); t.fill = fill(ORANGE);
        t.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
        ws.getRow(1).height = 32;
        ws.mergeCells(2, 1, 2, lastCol);
        const sub = ws.getCell(2, 1);
        sub.value = subtitle; sub.font = font({ size: 10, italic: true, color: { argb: CHAR } }); sub.alignment = { horizontal: 'center' };
        ws.getRow(2).height = 18;
        const head = ws.getRow(4);
        headers.forEach((h, i) => {
          const c = head.getCell(i + 1);
          c.value = h; c.font = font({ bold: true, color: { argb: WHITE } }); c.fill = fill(CHAR); c.border = border;
          c.alignment = { vertical: 'middle', horizontal: 'center' };
        });
        head.height = 22;
        ws.views = [{ state: 'frozen', ySplit: 4 }];
        ws.pageSetup = { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 0, printTitlesRow: '4:4',
          margins: { left: 0.5, right: 0.5, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 } };
        ws.headerFooter = { oddFooter: '&L&8TBS Logistics · ' + title + '&R&8Trang &P/&N' };
      };
      const bodyRow = (ws, r, values, centerCols) => {
        const row = ws.getRow(r);
        values.forEach((v, i) => {
          const c = row.getCell(i + 1);
          c.value = v; c.font = font(); c.border = border;
          if (r % 2 === 1) c.fill = fill(TINT);
          c.alignment = { vertical: 'middle', horizontal: centerCols.includes(i) ? 'center' : 'left' };
        });
        row.height = 18;
      };

      const sum = book.addWorksheet(sheetName('Tổng hợp', used), { properties: { tabColor: { argb: ORANGE } } });
      const passed = data.groups.filter(deptPass).length;
      layout(sum, 7, 'Cập nhật ' + stamp + ' · ' + data.missing + '/' + data.employees + ' nhân viên chưa làm · ' + passed + '/' + data.groups.length +
        ' bộ phận đạt (≥ ' + Math.round(PASS_RATE * 100) + '% nhân viên đã làm)', ['STT', 'Bộ phận', 'Tổng nhân sự', 'Đã làm', 'Chưa làm', 'Tỷ lệ đã làm', 'Kết quả'], [7, 30, 15, 11, 11, 14, 13]);
      const result = (cell, ok) => {
        cell.value = ok ? 'Đạt' : 'Chưa đạt';
        cell.font = font({ bold: true, color: { argb: ok ? 'FF14532D' : 'FF7F1D1D' } });
        cell.fill = fill(ok ? 'FFE8F6EE' : 'FFFDECEC');
      };
      data.groups.forEach((g, i) => {
        bodyRow(sum, 5 + i, [i + 1, g.group, g.total, g.done, g.missing.length, deptRate(g), ''], [0, 2, 3, 4, 5, 6]);
        sum.getCell(5 + i, 6).numFmt = '0%';
        result(sum.getCell(5 + i, 7), deptPass(g));
      });
      const totalRow = 5 + data.groups.length;
      bodyRow(sum, totalRow, ['', 'Tổng cộng', data.employees, data.done, data.missing, data.employees ? data.done / data.employees : 0, passed + '/' + data.groups.length + ' đạt'], [2, 3, 4, 5, 6]);
      sum.getCell(totalRow, 6).numFmt = '0%';
      sum.getRow(totalRow).eachCell(c => { c.font = font({ bold: true }); c.fill = fill(TINT); });

      data.groups.filter(g => g.missing.length).forEach(g => {
        const ws = book.addWorksheet(sheetName(g.group, used));
        layout(ws, 3, 'Bộ phận: ' + g.group + ' · ' + g.missing.length + '/' + g.total + ' người chưa làm · đã làm ' + Math.round(deptRate(g) * 100) + '% → ' +
          (deptPass(g) ? 'ĐẠT' : 'CHƯA ĐẠT') + ' (cần ≥ ' + Math.round(PASS_RATE * 100) + '%) · cập nhật ' + stamp, ['STT', 'MSNV', 'Họ và tên'], [8, 16, 48]);
        if (!deptPass(g)) ws.properties.tabColor = { argb: 'FFDC2626' };
        g.missing.forEach((person, i) => {
          bodyRow(ws, 5 + i, [i + 1, String(person.msnv), person.name], [0, 1]);
          ws.getCell(5 + i, 2).numFmt = '@'; // MSNV giữ dạng chữ
        });
      });

      const buffer = await book.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'Danh-sach-chua-lam-VHDN-thang-' + p.m + '-' + p.y + '.xlsx';
      document.body.appendChild(a); a.click(); a.remove();
      setTimeout(() => URL.revokeObjectURL(a.href), 4000);
      notice('Đã xuất Excel: ' + data.groups.filter(g => g.missing.length).length + ' bộ phận, ' + data.missing + ' nhân viên chưa làm.', 'success');
    } catch (error) { notice(error.message || String(error), 'error'); } finally { setLoading(false); }
  }

  /* ---------- Gắn sự kiện ---------- */
  function bind() {
    $('reportPeriod').onchange = e => { report.period = e.target.value; render(); };
    $('reportReload').onclick = () => { report.missing = {}; load(); };
    $('reportImportButton').onclick = () => $('reportFile').click();
    $('reportFile').onchange = onFile;
    $('reportExportExcel').onclick = exportExcel;
    $('reportExportPptx').onclick = exportPptx;
    $('reportBenchSave').onclick = () => saveBenchmark(false);
    $('reportBenchAuto').onclick = () => saveBenchmark(true);
    $('reportImports').onclick = e => { const b = e.target.closest('[data-del]'); if (b) deleteImport(b.dataset.del); };
    $('missingSearch').oninput = () => { clearTimeout(renderMissing.t); renderMissing.t = setTimeout(renderMissing, 150); };
    $('missingExport').onclick = exportMissing;
    $('missingReload').onclick = () => { const m = selectedMonth(); if (m) loadMissing(m, true); };
  }

  window.vhdnReport = { open() { if (!report.bound) { bind(); report.bound = true; } return load(); }, _state: report, _parsePptx: parsePptx, _parseExcel: parseExcel };
})();
