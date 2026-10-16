/**
 * Học 當代中文 1 — máy chủ thống kê (Google Apps Script, gắn với 1 file Google Sheets)
 *
 * 1) Đổi MẬT KHẨU giáo viên ở dòng dưới.
 * 2) Chạy hàm `caiDat` một lần (bấm ▶ Chạy) để tạo các trang tính và cấp quyền.
 * 3) Triển khai → Triển khai mới → Loại: Ứng dụng web
 *      - Thực thi dưới dạng: Tôi
 *      - Người có quyền truy cập: Bất kỳ ai
 *    → Sao chép "URL ứng dụng web" gửi cho Claude.
 */
const MAT_KHAU = 'doimatkhau';   // ← ĐỔI MẬT KHẨU GIÁO VIÊN Ở ĐÂY

const SH_HS = 'HocSinh', SH_LOG = 'NhatKy', SH_TK = 'ThongKe';
const H_HS = ['Mã lớp', 'Mã học sinh', 'Họ tên', 'Ngày tạo'];
const H_LOG = ['Thời gian', 'Mã lớp', 'Mã học sinh', 'Bài', 'Lật thẻ', 'Lần nghe', 'Câu trắc nghiệm', 'Trắc nghiệm đúng',
  'Câu luyện nghe', 'Luyện nghe đúng', 'Giây học'];
const H_TK = ['Mã lớp', 'Mã học sinh', 'Họ tên', 'Bài', 'Lật thẻ', 'Thẻ đã xem', 'Lần nghe', 'Câu trắc nghiệm',
  'Trắc nghiệm đúng', 'Câu luyện nghe', 'Luyện nghe đúng', 'Phút học', 'Từ đã nhớ', 'Lần cuối'];

function caiDat() { sheets_(); }

function sheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mk = (name, head) => {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); }
    if (sh.getLastRow() === 0) {
      sh.appendRow(head); sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#E8EFEC');
    }
    return sh;
  };
  const r = { hs: mk(SH_HS, H_HS), log: mk(SH_LOG, H_LOG), tk: mk(SH_TK, H_TK) };
  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 3) ss.deleteSheet(s1);
  return r;
}

const norm_ = v => String(v == null ? '' : v).trim().toUpperCase();
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function roster_(sh) {
  const v = sh.hs.getDataRange().getValues().slice(1);
  return v.filter(r => r[0] !== '' && r[1] !== '').map(r => ({ c: norm_(r[0]), s: norm_(r[1]), n: String(r[2] || '') }));
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const sh = sheets_();
  try {
    if (p.a === 'login') {
      const st = roster_(sh).find(r => r.c === norm_(p.c) && r.s === norm_(p.s));
      return out_(st ? { ok: true, c: st.c, s: st.s, n: st.n } : { ok: false, err: 'Sai mã lớp hoặc mã học sinh' });
    }
    if (p.a === 'stats' || p.a === 'add' || p.a === 'del') {
      if (String(p.pw || '') !== MAT_KHAU) return out_({ ok: false, err: 'Sai mật khẩu' });
      if (p.a === 'add') return out_(addStudents_(sh, p));
      if (p.a === 'del') return out_(delStudent_(sh, p));
      const tk = sh.tk.getDataRange().getValues().slice(1).map(r => r.slice(0, H_TK.length).map(x => x instanceof Date ? x.toISOString() : x));
      return out_({ ok: true, roster: roster_(sh), rows: tk, head: H_TK });
    }
    if (p.a === 'log' && p.d) return out_(log_(sh, JSON.parse(p.d)));
    return out_({ ok: true, msg: 'Máy chủ thống kê Học 當代中文 1 đang chạy.' });
  } catch (err) { return out_({ ok: false, err: String(err) }); }
}

function doPost(e) {
  const sh = sheets_();
  try {
    const d = JSON.parse(e.postData.contents);
    if (d.a === 'log') return out_(log_(sh, d));
    return out_({ ok: false, err: 'unknown' });
  } catch (err) { return out_({ ok: false, err: String(err) }); }
}

/* p.list = "LOP|MAHS|Ho ten\n..." */
function addStudents_(sh, p) {
  const have = new Set(roster_(sh).map(r => r.c + '|' + r.s));
  const rows = [];
  String(p.list || '').split(/\n/).forEach(line => {
    const [c, s, n] = line.split('|').map(x => (x || '').trim());
    if (!c || !s) return;
    const k = norm_(c) + '|' + norm_(s);
    if (have.has(k)) return; have.add(k);
    rows.push([norm_(c), norm_(s), n || '', new Date()]);
  });
  if (rows.length) sh.hs.getRange(sh.hs.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  return { ok: true, added: rows.length, roster: roster_(sh) };
}

function delStudent_(sh, p) {
  const v = sh.hs.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) {
    if (norm_(v[i][0]) === norm_(p.c) && norm_(v[i][1]) === norm_(p.s)) sh.hs.deleteRow(i + 1);
  }
  return { ok: true, roster: roster_(sh) };
}

/* d = {a:'log', c, s, d:{ "<bai>": {f,seen,li,q,qc,lq,lqc,t,kn} } } */
function log_(sh, d) {
  const st = roster_(sh).find(r => r.c === norm_(d.c) && r.s === norm_(d.s));
  if (!st) return { ok: false, err: 'Mã không hợp lệ' };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const now = new Date();
    const data = d.d || {};
    const logRows = [];
    const tkR = sh.tk.getDataRange().getValues();
    const idx = {};
    for (let i = 1; i < tkR.length; i++) idx[norm_(tkR[i][0]) + '|' + norm_(tkR[i][1]) + '|' + tkR[i][3]] = i;
    const add = [];
    Object.keys(data).forEach(bai => {
      const x = data[bai] || {}; const n = k => Math.max(0, Math.min(100000, Number(x[k]) || 0));
      const lesson = bai === '0' ? 'Tổng hợp' : 'Bài ' + bai;
      logRows.push([now, st.c, st.s, lesson, n('f'), n('li'), n('q'), n('qc'), n('lq'), n('lqc'), n('t')]);
      const key = st.c + '|' + st.s + '|' + lesson;
      let row = idx[key] != null ? tkR[idx[key]] : null;
      if (!row) { row = [st.c, st.s, st.n, lesson, 0, 0, 0, 0, 0, 0, 0, 0, 0, now]; add.push(row); idx[key] = -1; }
      row[2] = st.n;
      row[4] = (+row[4] || 0) + n('f');
      if (x.seen != null) row[5] = Math.max(+row[5] || 0, n('seen'));
      row[6] = (+row[6] || 0) + n('li');
      row[7] = (+row[7] || 0) + n('q');
      row[8] = (+row[8] || 0) + n('qc');
      row[9] = (+row[9] || 0) + n('lq');
      row[10] = (+row[10] || 0) + n('lqc');
      row[11] = Math.round(((+row[11] || 0) + n('t') / 60) * 10) / 10;
      if (x.kn != null) row[12] = n('kn');
      row[13] = now;
    });
    if (logRows.length) sh.log.getRange(sh.log.getLastRow() + 1, 1, logRows.length, H_LOG.length).setValues(logRows);
    if (tkR.length > 1) sh.tk.getRange(2, 1, tkR.length - 1, H_TK.length).setValues(tkR.slice(1).map(r => r.slice(0, H_TK.length)));
    if (add.length) sh.tk.getRange(sh.tk.getLastRow() + 1, 1, add.length, H_TK.length).setValues(add);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
