/**
 * NHẬN ĐƠN TỪ LANDING PAGE → GHI GOOGLE SHEET + GỬI EMAIL THÔNG BÁO
 *
 * CÁCH CÀI (làm 1 lần):
 * 1. Tạo 1 Google Sheet mới → menu Tiện ích mở rộng (Extensions) → Apps Script.
 * 2. Xoá code mẫu, dán toàn bộ file này vào. Sửa NOTIFY_EMAIL bên dưới. Bấm Lưu.
 * 3. Bấm Triển khai (Deploy) → Tùy chọn triển khai mới (New deployment)
 *      - Loại: Ứng dụng web (Web app)
 *      - Thực thi dưới dạng (Execute as): Tôi (Me)
 *      - Người có quyền truy cập (Who has access): Bất kỳ ai (Anyone)
 *    → Triển khai → Cấp quyền (chọn tài khoản → Nâng cao → Đi tới ... → Cho phép).
 * 4. Copy "URL ứng dụng web" (dạng https://script.google.com/macros/s/.../exec)
 *    dán vào CONFIG.submitUrl trong index.html.
 *
 * Lưu ý: mỗi lần sửa code này phải Triển khai → Quản lý triển khai → Chỉnh sửa →
 * Phiên bản: "Phiên bản mới" thì thay đổi mới có hiệu lực (URL giữ nguyên).
 */

// Email nhận thông báo đơn mới. Nhiều email thì cách nhau dấu phẩy.
const NOTIFY_EMAIL = "datvu06803@gmail.com";
const SHEET_NAME = "Đơn hàng";

const HEADERS = ["Thời gian", "Mã đơn", "Họ tên", "SĐT", "Địa chỉ", "Sản phẩm", "Giá", "Ghi chú"];

const money_ = n => Number(n || 0).toLocaleString("vi-VN") + "đ";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = JSON.parse(e.postData.contents);
    const fullAddress = [d.address, d.ward, d.district, d.province].filter(String).join(", ");
    const product = d.package + (d.gift ? " + " + d.gift : "");
    const price = money_(d.total);

    getSheet_().appendRow([
      d.time || new Date(), d.code, d.name, "'" + d.phone, fullAddress, product, price, d.note
    ]);

    sendMail_(d, fullAddress);
    return json_({ ok: true });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// Mở URL web app trên trình duyệt để kiểm tra script đang chạy.
function doGet() {
  return json_({ ok: true, message: "Order webhook is running" });
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold").setBackground("#fbe9e7");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sendMail_(d, fullAddress) {
  if (!NOTIFY_EMAIL) return;
  const money = money_;
  const esc = s => String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  const rows = [
    ["Mã đơn", d.code],
    ["Họ tên", d.name],
    ["SĐT", d.phone],
    ["Gói", d.package],
    ["Giá", money(d.price)],
    ["Phí ship", d.ship ? money(d.ship) : "Miễn phí"],
    ["Tổng thu (COD)", money(d.total)],
    ["Quà tặng", d.gift || "—"],
    ["Địa chỉ", fullAddress],
    ["Ghi chú", d.note || "—"],
    ["Thời gian", d.time],
    ["Trang", d.page]
  ];
  const html =
    '<div style="font-family:Arial,sans-serif;font-size:14px">' +
    '<h3 style="color:#b3261e;margin:0 0 12px">🛒 Đơn hàng mới ' + esc(d.code) + '</h3>' +
    '<table cellpadding="6" style="border-collapse:collapse">' +
    rows.map(r => '<tr><td style="border:1px solid #e6dccb;background:#faf6ef"><b>' + esc(r[0]) +
      '</b></td><td style="border:1px solid #e6dccb">' + esc(r[1]) + '</td></tr>').join("") +
    '</table></div>';

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: "[Đơn mới] " + d.name + " - " + d.phone + " - " + d.package + " - " + money(d.total),
    htmlBody: html,
    name: "Landing Sách Tiếng Trung"
  });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Chạy hàm này 1 lần trong trình soạn thảo để cấp quyền và thử gửi 1 đơn mẫu.
function testOrder() {
  doPost({ postData: { contents: JSON.stringify({
    code: "TEST01", package: "Combo 2 bộ", price: 389000, ship: 0, total: 389000,
    name: "Nguyễn Văn A", phone: "0912345678", province: "Hà Nội", district: "Huyện Gia Lâm",
    ward: "Xã Kim Sơn", address: "Số 12", note: "Đơn thử", gift: "Tặng 1 bút bay màu + 10 ruột",
    page: "test", time: new Date().toLocaleString("vi-VN")
  }) } });
}
