/**
 * Code_PipeSpool.gs - Web app nối ứng dụng Pipe Spool với Google Sheet.
 *
 * Ứng dụng gọi bằng GET:  <Web app URL>?action=traCuu|capNhat&code=...&dwg=...&sp=...
 *   traCuu : tra một mã; mã chưa có trong DanhSachChiTiet KHÔNG bị thêm vào, chỉ ghi lịch sử (tô vàng, tự hết tô khi mã được bổ sung).
 *   capNhat: ghi công đoạn (chỉ cho mã đã có trong DanhSachChiTiet) (stageIndex 0..12), pallet, ghi chú vào TrangThaiSanPham + LichSuCapNhat.
 * Trả về JSON: { ok, notInList, data: { master, status } } hoặc { ok:false, error }.
 *
 * CÁCH DÙNG
 *  1. Dán toàn bộ file này vào Apps Script (thay nội dung Mã.gs).
 *  2. Đặt SHEET_ID bên dưới = ID của Google Sheet (đoạn giữa /d/ và /edit trên thanh địa chỉ).
 *     Nếu mở Apps Script từ chính Google Sheet (Tiện ích mở rộng > Apps Script) thì để trống.
 *  3. Chọn hàm kiemTra rồi bấm Chạy, cho phép quyền truy cập khi được hỏi.
 *  4. Triển khai > Triển khai mới > Ứng dụng web: Thực thi bằng "Tôi", Người có quyền truy cập "Bất kỳ ai".
 *  5. Copy "URL ứng dụng web" dán vào app (nút ⚙️ ở góc trên).
 *  Sửa code xong phải Triển khai > Quản lý bản triển khai > Phiên bản mới để URL giữ nguyên.
 */

var SHEET_ID = '';   // <- dán ID Google Sheet vào đây (hoặc để trống nếu script gắn trong Sheet)

var STAGES = ['Cắt chi tiết', 'Gá đính', 'Hàn lót', 'Hàn dở dang', 'Hoàn thiện hàn',
  'Rửa axit', 'Chuyển công đoạn sơn', 'Làm sạch', 'Sơn lót', 'Sơn trung gian', 'Sơn phủ',
  'Hoàn thiện sơn', 'Đóng kiện'];
var PRE_END = 4, ACID_I = 5, PAINT_START = 6, FINAL_I = 12;

var SH_MASTER = 'DanhSachChiTiet', SH_STATUS = 'TrangThaiSanPham', SH_LOG = 'LichSuCapNhat', SH_PAINT = 'BangMauSon';

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}
function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Không thấy sheet "' + name + '" trong Google Sheet.');
  return sh;
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function s_(v) { return v === null || v === undefined ? '' : String(v); }

function doGet(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(25000); } catch (err) { return json_({ ok: false, error: 'Máy chủ đang bận, thử lại sau vài giây.' }); }
  try {
    var p = (e && e.parameter) || {};
    var code = s_(p.code).trim();
    if (!code) throw new Error('Thiếu mã chi tiết.');
    var out;
    if (p.action === 'traCuu') out = traCuu_(p, code);
    else if (p.action === 'capNhat') out = capNhat_(p, code);
    else throw new Error('Hành động không hợp lệ: ' + p.action);
    return json_(out);
  } catch (err) {
    return json_({ ok: false, error: s_(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Tìm dòng theo Mã chi tiết (cột B) ---------- */
function findRow_(sh, code) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var v = sh.getRange(2, 2, last - 1, 1).getValues();
  for (var i = 0; i < v.length; i++) if (s_(v[i][0]).trim() === code) return i + 2;
  return 0;
}
function nextFreeRow_(sh) {   // dòng đầu tiên sau dòng cuối có Mã chi tiết
  var last = sh.getLastRow();
  if (last < 2) return 2;
  var v = sh.getRange(2, 2, last - 1, 1).getValues(), n = 0;
  for (var i = 0; i < v.length; i++) if (s_(v[i][0]).trim() !== '') n = i + 1;
  return n + 2;
}

/* ---------- DanhSachChiTiet ---------- */
function masterObj_(r) {
  var v = sheet_(SH_MASTER).getRange(r, 1, 1, 21).getValues()[0];
  // Chưa nhập Công đoạn cuối nhưng đã có màu sơn cụ thể (số RAL) -> mặc định "Hoàn thiện sơn"
  if (!s_(v[15]).trim() && /RAL\s*\d{4}/.test(s_(v[10]))) v[15] = 'Hoàn thiện sơn';
  return {
    dwg: s_(v[2]), sp: s_(v[3]), qty: s_(v[4]), weight: s_(v[5]), description: s_(v[6]),
    typeMaterial: s_(v[7]), paintType: s_(v[8]), finishColor: s_(v[10]), commodity: s_(v[11]),
    paintSystem: s_(v[12]), ngayGiaoHang: v[13] instanceof Date ? Utilities.formatDate(v[13], 'GMT+7', 'yyyy-MM-dd') : s_(v[13]),
    note: s_(v[14]), congDoanCuoi: s_(v[15]), trangThaiDuLieu: s_(v[16]),
    goi: s_(v[17]), trang: s_(v[18]), file: s_(v[19])
  };
}

/* ---------- Công đoạn được áp dụng theo "Công đoạn cuối" (giống app) ---------- */
function allowedSet_(congDoanCuoi) {
  var ci = STAGES.indexOf(congDoanCuoi);
  if (ci < 0) return null;   // chưa có Công đoạn cuối -> cho mọi công đoạn
  var allow = {}, i;
  for (i = 0; i <= PRE_END; i++) allow[i] = true;
  allow[FINAL_I] = true;
  if (ci === ACID_I) allow[ACID_I] = true;
  else if (ci >= PAINT_START && ci < FINAL_I) for (i = PAINT_START; i <= ci; i++) allow[i] = true;
  return allow;
}

/* ---------- TrangThaiSanPham ---------- */
function lastDoneIndex_(marks) {
  var idx = -1;
  for (var i = 0; i < marks.length; i++) if (marks[i] === 'X') idx = i;
  return idx;
}
function countLog_(code) {
  var sh = sheet_(SH_LOG), last = sh.getLastRow();
  if (last < 2) return 0;
  var v = sh.getRange(2, 2, last - 1, 1).getValues(), n = 0;
  for (var i = 0; i < v.length; i++) if (s_(v[i][0]).trim() === code) n++;
  return n;
}
function statusObj_(code) {
  var sh = sheet_(SH_STATUS), r = findRow_(sh, code);
  if (!r) return null;
  var row = sh.getRange(r, 1, 1, 25).getValues()[0], marks = [];
  for (var i = 0; i < 13; i++) marks.push(s_(row[7 + i]));
  var t = row[24];
  return {
    marks: marks, pallet: s_(row[20]),
    capNhatLanCuoi: t instanceof Date ? t.toISOString() : s_(t),
    soLanTacDong: countLog_(code)
  };
}

/* ---------- LichSuCapNhat ---------- */
var LOG_FLAG_TEXT = 'Chưa có trong danh sách';
function logFlagFormula_(r) {
  return '=IF(B' + r + '="","",IF(COUNTIF(' + SH_MASTER + '!$B:$B,B' + r + ')=0,"' + LOG_FLAG_TEXT + '",""))';
}
function writeLog_(code, dwg, sp, action, stage, pallet, note) {
  var sh = sheet_(SH_LOG);
  sh.appendRow([new Date(), code, dwg, sp, action, stage, pallet || '', note || '']);
  sh.getRange(sh.getLastRow(), 9).setFormula(logFlagFormula_(sh.getLastRow()));   // cột I: tự hết cảnh báo khi mã vào Danh sách chi tiết
}
function lastLogIsRecentLookup_(code) {   // tránh ghi trùng khi app tra cứu liên tiếp trong vài giây
  var sh = sheet_(SH_LOG), last = sh.getLastRow();
  if (last < 2) return false;
  var v = sh.getRange(last, 1, 1, 5).getValues()[0];
  return s_(v[1]) === code && s_(v[4]) === 'Tra cứu' && v[0] instanceof Date && (new Date() - v[0]) < 10000;
}

/* ---------- Hành động: traCuu ---------- */
function traCuu_(p, code) {
  var r = findRow_(sheet_(SH_MASTER), code);
  if (!r) {   // mã chưa có trong danh sách: KHÔNG thêm vào DanhSachChiTiet, chỉ ghi lịch sử (cột I tự đánh dấu)
    if (!lastLogIsRecentLookup_(code)) {
      var dwg0 = s_(p.dwg) || code.replace(/-SP\d+$/, ''), sp0 = s_(p.sp) || (code.match(/SP\d+$/) || [''])[0];
      writeLog_(code, dwg0, sp0, 'Tra cứu', 'Chưa có trong danh sách', '', 'Mã chưa có trong Danh sách chi tiết - chỉ ghi lịch sử, không thêm vào danh sách');
    }
    return { ok: true, notInList: true, data: { master: null, status: null } };
  }
  var master = masterObj_(r);
  var status = statusObj_(code);
  if (!lastLogIsRecentLookup_(code)) {
    var idx = status ? lastDoneIndex_(status.marks) : -1;
    writeLog_(code, master.dwg, master.sp, 'Tra cứu', idx >= 0 ? STAGES[idx] : 'Chưa bắt đầu', status ? status.pallet : '', '');
    if (status) status.soLanTacDong = countLog_(code);
  }
  return { ok: true, notInList: false, data: { master: master, status: status } };
}

/* ---------- Hành động: capNhat ---------- */
function capNhat_(p, code) {
  var idx = parseInt(p.stageIndex, 10);
  if (isNaN(idx) || idx < 0 || idx > FINAL_I) throw new Error('Công đoạn không hợp lệ.');
  var mr = findRow_(sheet_(SH_MASTER), code);
  if (!mr) {
    writeLog_(code, s_(p.dwg), s_(p.sp), 'Cập nhật bị từ chối', 'Chưa có trong danh sách', '', 'Mã chưa có trong Danh sách chi tiết - chưa thể cập nhật công đoạn');
    throw new Error('Mã này chưa có trong Danh sách chi tiết nên chưa thể cập nhật công đoạn. Hãy bổ sung mã vào danh sách trước.');
  }
  var master = masterObj_(mr);
  var allow = allowedSet_(master.congDoanCuoi);
  if (allow && !allow[idx]) throw new Error('Công đoạn "' + STAGES[idx] + '" không áp dụng cho chi tiết này (Công đoạn cuối: ' + master.congDoanCuoi + ').');

  var sh = sheet_(SH_STATUS), r = findRow_(sh, code);
  var prev = statusObj_(code);
  var prevIdx = prev ? lastDoneIndex_(prev.marks) : -1;
  var prevPallet = prev ? prev.pallet : '';
  var pallet = prevPallet;
  if (idx === FINAL_I) {
    pallet = s_(p.pallet).trim();
    if (!pallet) throw new Error('Cần nhập Số Pallet khi Đóng kiện.');
  }
  var note = s_(p.note).trim();
  var isDown = prevIdx >= 0 && idx < prevIdx;
  var isPalletChange = idx === FINAL_I && prevIdx === FINAL_I && prevPallet && pallet !== prevPallet;
  if ((isDown || isPalletChange) && !note) throw new Error('Trường hợp này cần nhập Ghi chú lý do thay đổi.');
  var action = isDown ? 'Hạ công đoạn (sửa lỗi)' : (isPalletChange ? 'Đổi Pallet' : (idx === FINAL_I ? 'Đóng kiện' : 'Cập nhật công đoạn'));
  var stageNote = isPalletChange ? STAGES[idx] + ' (Pallet cũ: ' + prevPallet + ')' : STAGES[idx];

  var marks = [];
  for (var i = 0; i < 13; i++) marks.push(allow && !allow[i] ? '-' : (i <= idx ? 'X' : ''));
  if (!r) r = nextFreeRow_(sh);
  sh.getRange(r, 2, 1, 4).setValues([[code, master.dwg, master.sp, master.congDoanCuoi]]);
  sh.getRange(r, 7).setValue(master.finishColor);
  sh.getRange(r, 8, 1, 13).setValues([marks]);
  sh.getRange(r, 21, 1, 5).setValues([[pallet, (idx + 1) + '/13', STAGES[idx], 0, new Date()]]);
  sh.getRange(r, 1).setFormula('=IF(B' + r + '="","",ROW()-1)');

  writeLog_(code, master.dwg, master.sp, action, stageNote, pallet, note);
  sh.getRange(r, 24).setValue(countLog_(code));
  SpreadsheetApp.flush();
  return { ok: true, notInList: false, data: { master: masterObj_(mr), status: statusObj_(code) } };
}

/* ---------- Chạy tay một lần để cấp quyền và kiểm tra sheet ---------- */
function kiemTra() {
  var need = [SH_MASTER, SH_STATUS, SH_LOG, SH_PAINT], msg = [];
  need.forEach(function (n) {
    var sh = ss_().getSheetByName(n);
    msg.push(n + ': ' + (sh ? (sh.getLastRow() - 1) + ' dòng dữ liệu' : 'KHÔNG CÓ SHEET NÀY'));
  });
  Logger.log(msg.join('\n'));
  return msg.join('\n');
}

/* ---------- Chạy tay MỘT LẦN: tạo cột cảnh báo + tô vàng trong LichSuCapNhat ---------- */
function caiDatCanhBaoLichSu() {
  var sh = sheet_(SH_LOG), last = Math.max(sh.getLastRow(), 2);
  sh.getRange(1, 9).setValue('Kiểm tra mã').setFontWeight('bold');
  var f = [];
  for (var r = 2; r <= last; r++) f.push([logFlagFormula_(r)]);
  sh.getRange(2, 9, f.length, 1).setFormulas(f);
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$I2="' + LOG_FLAG_TEXT + '"')
    .setBackground('#FFE599').setBold(true)
    .setRanges([sh.getRange(2, 1, 5000, 9)]).build();
  sh.setConditionalFormatRules([rule]);   // thay các quy tắc tô màu cũ (nếu có) của sheet này
  SpreadsheetApp.flush();
  Logger.log('Đã cài đặt cảnh báo cho ' + (last - 1) + ' dòng lịch sử.');
}
