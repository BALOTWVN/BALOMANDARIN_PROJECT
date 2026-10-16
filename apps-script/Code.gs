/**
 * BALO Mandarin · 當代中文課程 1 — máy chủ tài khoản & thống kê
 * (Google Apps Script gắn với 1 file Google Sheets của giáo viên)
 *
 * CÀI ĐẶT (một lần):
 * 1) Đổi tên đăng nhập / mật khẩu QUẢN TRỊ ở 2 dòng bên dưới.
 * 2) Chọn hàm `caiDat` → bấm ▶ Chạy → cấp quyền.
 * 3) Triển khai → Triển khai mới → Ứng dụng web
 *      Thực thi dưới dạng: Tôi · Người có quyền truy cập: Bất kỳ ai
 *    → sao chép "URL ứng dụng web".
 * Khi sửa mã sau này: Triển khai → Quản lý triển khai → ✏️ → Phiên bản: Mới → Triển khai.
 * Phiên bản mã: 2 (gửi dữ liệu qua tham số j, tránh tên tham số 'c' Google giữ riêng)
 */
const ADMIN_USER = 'admin';        // ← tên đăng nhập quản trị
const ADMIN_PASS = 'doimatkhau';   // ← mật khẩu quản trị (chỉ dùng khi chạy caiDat lần đầu)

const SH = {
  gv: ['GiaoVien', ['Tên đăng nhập', 'Họ tên', 'Mật khẩu (mã hóa)', 'Quản trị', 'Ngày tạo']],
  lop: ['Lop', ['Mã lớp', 'Tên lớp', 'Giáo viên', 'Ngày tạo']],
  hs: ['HocSinh', ['Mã lớp', 'Mã học sinh', 'Họ tên', 'Mật khẩu (mã hóa)', 'Ngày tạo']],
  log: ['NhatKy', ['Thời gian', 'Mã lớp', 'Mã học sinh', 'Bài', 'Lật thẻ', 'Lần nghe', 'Câu trắc nghiệm', 'Trắc nghiệm đúng',
    'Câu luyện nghe', 'Luyện nghe đúng', 'Giây học']],
  tk: ['ThongKe', ['Mã lớp', 'Mã học sinh', 'Họ tên', 'Bài', 'Lật thẻ', 'Thẻ đã xem', 'Lần nghe', 'Câu trắc nghiệm',
    'Trắc nghiệm đúng', 'Câu luyện nghe', 'Luyện nghe đúng', 'Phút học', 'Từ đã nhớ', 'Lần cuối']]
};
const NTK = SH.tk[1].length;

function caiDat() {
  const s = sheets_();
  secret_();
  if (!rows_(s.gv).some(r => norm_(r[0]) === norm_(ADMIN_USER)))
    s.gv.appendRow([norm_(ADMIN_USER), 'Quản trị', hash_(ADMIN_PASS), 'x', new Date()]);
}

/* ---------- helpers ---------- */
function sheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), o = {};
  Object.keys(SH).forEach(k => {
    const [name, head] = SH[k];
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(head); sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#E8EFEC');
    }
    o[k] = sh;
  });
  ['Sheet1', 'Trang tính1', 'Trang tính 1'].forEach(n => { const x = ss.getSheetByName(n); if (x && x.getLastRow() === 0 && ss.getSheets().length > 5) ss.deleteSheet(x); });
  return o;
}
const norm_ = v => String(v == null ? '' : v).trim().toUpperCase();
const out_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const rows_ = sh => sh.getDataRange().getValues().slice(1);
function secret_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('SECRET', s); }
  return s;
}
const hex_ = b => b.map(x => ('0' + (x & 255).toString(16)).slice(-2)).join('');
function hash_(pw) { return hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, secret_() + '|' + String(pw))); }
function tok_(c, s) { return hex_(Utilities.computeHmacSha256Signature(norm_(c) + '|' + norm_(s), secret_())).slice(0, 32); }
function delRows_(sh, pred) { const v = sh.getDataRange().getValues(); for (let i = v.length - 1; i >= 1; i--) if (pred(v[i])) sh.deleteRow(i + 1); }

/* ---------- entry points ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.j) { try { return handle_(JSON.parse(p.j)); } catch (err) { return out_({ ok: false, err: String(err) }); } }
  if (p.a) return handle_(p);
  return out_({ ok: true, msg: 'Máy chủ BALO Mandarin đang chạy.' });
}
function doPost(e) {
  try { return handle_(JSON.parse(e.postData.contents)); }
  catch (err) { return out_({ ok: false, err: String(err) }); }
}
function handle_(p) {
  if (p.lop != null && p.c == null) p.c = p.lop; // 'c' là tham số Google giữ riêng
  let s;
  try {
    s = sheets_();
    switch (p.a) {
      case 'login': return out_(stLogin_(s, p));
      case 'chpw': return out_(stChpw_(s, p));
      case 'log': return out_(log_(s, p));
      case 'tlogin': { const t = teacher_(s, p); return out_(t ? { ok: true, u: t.u, n: t.n, admin: t.admin } : { ok: false, err: 'Sai tên đăng nhập hoặc mật khẩu' }); }
    }
    const t = teacher_(s, p);
    if (!t) return out_({ ok: false, err: 'Phiên đăng nhập giáo viên không hợp lệ', auth: true });
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try { return out_(teacherAct_(s, t, p)); } finally { lock.releaseLock(); }
  } catch (err) { return out_({ ok: false, err: String(err) }); }
}

/* ---------- students ---------- */
function findSt_(s, c, st) {
  const v = s.hs.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (norm_(v[i][0]) === norm_(c) && norm_(v[i][1]) === norm_(st)) return { i: i + 1, r: v[i] };
  return null;
}
function stPwOk_(row, pw) { return row[3] ? row[3] === hash_(pw) : norm_(pw) === norm_(row[1]); } // chưa đặt mật khẩu → mật khẩu = mã học sinh
function stLogin_(s, p) {
  const f = findSt_(s, p.c, p.s);
  if (!f || !stPwOk_(f.r, p.p || '')) return { ok: false, err: 'Sai mã lớp, mã học sinh hoặc mật khẩu' };
  const lop = rows_(s.lop).find(r => norm_(r[0]) === norm_(p.c));
  return { ok: true, c: norm_(f.r[0]), s: norm_(f.r[1]), n: String(f.r[2] || ''), cn: lop ? String(lop[1] || '') : '', tok: tok_(f.r[0], f.r[1]), mustChange: !f.r[3] };
}
function stChpw_(s, p) {
  const f = findSt_(s, p.c, p.s);
  if (!f || !stPwOk_(f.r, p.p || '')) return { ok: false, err: 'Mật khẩu hiện tại không đúng' };
  if (String(p.np || '').length < 4) return { ok: false, err: 'Mật khẩu mới cần ít nhất 4 ký tự' };
  s.hs.getRange(f.i, 4).setValue(hash_(p.np));
  return { ok: true };
}

/* ---------- teachers ---------- */
function teacher_(s, p) {
  if (!p.u || !p.p) return null;
  const r = rows_(s.gv).find(r => norm_(r[0]) === norm_(p.u));
  if (!r || r[2] !== hash_(p.p)) return null;
  return { u: norm_(r[0]), n: String(r[1] || r[0]), admin: !!r[3] };
}
function myClasses_(s, t) {
  return rows_(s.lop).filter(r => r[0] !== '' && (t.admin || norm_(r[2]) === t.u))
    .map(r => ({ c: norm_(r[0]), n: String(r[1] || ''), gv: norm_(r[2]), at: r[3] instanceof Date ? r[3].toISOString() : '' }));
}
function own_(s, t, c) { return myClasses_(s, t).some(x => x.c === norm_(c)); }
const cleanCode_ = c => norm_(c).replace(/[^A-Z0-9_-]/g, '').slice(0, 20);

function teacherAct_(s, t, p) {
  switch (p.a) {
    case 'data': {
      const cls = myClasses_(s, t), set = new Set(cls.map(x => x.c));
      const roster = rows_(s.hs).filter(r => set.has(norm_(r[0]))).map(r => ({ c: norm_(r[0]), s: norm_(r[1]), n: String(r[2] || ''), pw: !!r[3] }));
      const rows = rows_(s.tk).filter(r => set.has(norm_(r[0]))).map(r => r.slice(0, NTK).map(x => x instanceof Date ? x.toISOString() : x));
      const res = { ok: true, me: t, classes: cls, roster, rows, head: SH.tk[1] };
      if (t.admin) res.teachers = rows_(s.gv).map(r => ({ u: norm_(r[0]), n: String(r[1] || ''), admin: !!r[3] }));
      return res;
    }
    case 'addclass': {
      const c = cleanCode_(p.c);
      if (c.length < 2) return { ok: false, err: 'Mã lớp cần ít nhất 2 ký tự (chữ, số, - hoặc _)' };
      if (rows_(s.lop).some(r => norm_(r[0]) === c)) return { ok: false, err: 'Mã lớp ' + c + ' đã tồn tại' };
      s.lop.appendRow([c, String(p.n || ''), t.u, new Date()]);
      return { ok: true, c };
    }
    case 'editclass': {
      if (!own_(s, t, p.c)) return { ok: false, err: 'Không có quyền với lớp này' };
      const v = s.lop.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) if (norm_(v[i][0]) === norm_(p.c)) {
        if (p.n != null) s.lop.getRange(i + 1, 2).setValue(String(p.n));
        if (t.admin && p.gv) s.lop.getRange(i + 1, 3).setValue(norm_(p.gv));
      }
      return { ok: true };
    }
    case 'delclass': {
      if (!own_(s, t, p.c)) return { ok: false, err: 'Không có quyền với lớp này' };
      delRows_(s.lop, r => norm_(r[0]) === norm_(p.c));
      delRows_(s.hs, r => norm_(r[0]) === norm_(p.c));
      return { ok: true };
    }
    case 'addst': {
      if (!own_(s, t, p.c)) return { ok: false, err: 'Không có quyền với lớp này' };
      const c = norm_(p.c), have = new Set(rows_(s.hs).filter(r => norm_(r[0]) === c).map(r => norm_(r[1])));
      const add = [], skip = [];
      String(p.list || '').split(/\n/).forEach(line => {
        const parts = line.split('|').map(x => (x || '').trim());
        const code = cleanCode_(parts[0]); if (!code) return;
        if (have.has(code)) { skip.push(code); return; }
        have.add(code);
        add.push([c, code, parts[1] || '', parts[2] ? hash_(parts[2]) : '', new Date()]);
      });
      if (add.length) s.hs.getRange(s.hs.getLastRow() + 1, 1, add.length, 5).setValues(add);
      return { ok: true, added: add.length, skipped: skip };
    }
    case 'editst': {
      if (!own_(s, t, p.c)) return { ok: false, err: 'Không có quyền với lớp này' };
      const f = findSt_(s, p.c, p.s); if (!f) return { ok: false, err: 'Không tìm thấy học sinh' };
      if (p.n != null) s.hs.getRange(f.i, 3).setValue(String(p.n));
      if (p.reset) s.hs.getRange(f.i, 4).setValue(p.np ? hash_(p.np) : '');
      return { ok: true };
    }
    case 'delst': {
      if (!own_(s, t, p.c)) return { ok: false, err: 'Không có quyền với lớp này' };
      delRows_(s.hs, r => norm_(r[0]) === norm_(p.c) && norm_(r[1]) === norm_(p.s));
      return { ok: true };
    }
    case 'tchpw': {
      if (String(p.np || '').length < 6) return { ok: false, err: 'Mật khẩu mới cần ít nhất 6 ký tự' };
      const v = s.gv.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) if (norm_(v[i][0]) === t.u) s.gv.getRange(i + 1, 3).setValue(hash_(p.np));
      return { ok: true };
    }
    case 'addteacher': {
      if (!t.admin) return { ok: false, err: 'Chỉ quản trị mới tạo được tài khoản giáo viên' };
      const u = cleanCode_(p.nu);
      if (u.length < 3) return { ok: false, err: 'Tên đăng nhập cần ít nhất 3 ký tự' };
      if (String(p.np || '').length < 6) return { ok: false, err: 'Mật khẩu cần ít nhất 6 ký tự' };
      if (rows_(s.gv).some(r => norm_(r[0]) === u)) return { ok: false, err: 'Tên đăng nhập đã tồn tại' };
      s.gv.appendRow([u, String(p.nn || u), hash_(p.np), p.nadmin ? 'x' : '', new Date()]);
      return { ok: true };
    }
    case 'delteacher': {
      if (!t.admin) return { ok: false, err: 'Chỉ quản trị mới xóa được' };
      if (norm_(p.du) === t.u) return { ok: false, err: 'Không thể tự xóa tài khoản đang dùng' };
      delRows_(s.gv, r => norm_(r[0]) === norm_(p.du));
      return { ok: true };
    }
    case 'resetteacher': {
      if (!t.admin) return { ok: false, err: 'Chỉ quản trị mới đặt lại được' };
      if (String(p.np || '').length < 6) return { ok: false, err: 'Mật khẩu cần ít nhất 6 ký tự' };
      const v = s.gv.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) if (norm_(v[i][0]) === norm_(p.du)) s.gv.getRange(i + 1, 3).setValue(hash_(p.np));
      return { ok: true };
    }
  }
  return { ok: false, err: 'Yêu cầu không hợp lệ' };
}

/* ---------- usage log: {a:'log', c, s, tok, d:{ "<bai>": {f,seen,li,q,qc,lq,lqc,t,kn} }} ---------- */
function log_(s, d) {
  if (!d.tok || d.tok !== tok_(d.c, d.s)) return { ok: false, err: 'Phiên học sinh không hợp lệ' };
  const f = findSt_(s, d.c, d.s); if (!f) return { ok: false, err: 'Không tìm thấy học sinh' };
  const c = norm_(f.r[0]), st = norm_(f.r[1]), name = String(f.r[2] || '');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const now = new Date(), data = d.d || {}, logRows = [], add = [];
    const tkR = s.tk.getDataRange().getValues(), idx = {};
    for (let i = 1; i < tkR.length; i++) idx[norm_(tkR[i][0]) + '|' + norm_(tkR[i][1]) + '|' + tkR[i][3]] = i;
    Object.keys(data).forEach(bai => {
      const x = data[bai] || {}, n = k => Math.max(0, Math.min(100000, Number(x[k]) || 0));
      const lesson = bai === '0' ? 'Tổng hợp' : 'Bài ' + Number(bai);
      logRows.push([now, c, st, lesson, n('f'), n('li'), n('q'), n('qc'), n('lq'), n('lqc'), n('t')]);
      const key = c + '|' + st + '|' + lesson;
      let row = idx[key] != null && idx[key] >= 0 ? tkR[idx[key]] : null;
      if (!row) { row = [c, st, name, lesson, 0, 0, 0, 0, 0, 0, 0, 0, 0, now]; add.push(row); idx[key] = -1; }
      row[2] = name;
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
    if (logRows.length) s.log.getRange(s.log.getLastRow() + 1, 1, logRows.length, logRows[0].length).setValues(logRows);
    if (tkR.length > 1) s.tk.getRange(2, 1, tkR.length - 1, NTK).setValues(tkR.slice(1).map(r => r.slice(0, NTK)));
    if (add.length) s.tk.getRange(s.tk.getLastRow() + 1, 1, add.length, NTK).setValues(add);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
