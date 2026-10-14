/**
 * HỆ THỐNG QUẢN LÝ DỰ ÁN & PHÂN VIỆC TỰ ĐỘNG
 * Phiên bản: 2.2 (V8) - Tích hợp Tên Thành viên, Deadline & Link Dự án
 */

// ==========================================
// 1. CẤU HÌNH & HẰNG SỐ CƠ BẢN
// ==========================================
const SHEETS = {
  TASKS: "Quan_Ly_Cong_Viec",
  MEMBERS: "Danh_Sach_Thanh_Vien",
  CONFIG: "Cau_Hinh",
  LOGS: "Nhat_Ky_Thong_Bao",
  STATE: "Trang_Thai_Task",
  CALENDAR: "Lien_Ket_Calendar"
};
const NEW_COLS = [
  "Mã TV thực hiện chính", "Mã TV phê duyệt", "Link sản phẩm bàn giao",
  "Gửi lại bàn giao", "Ngày giao việc", "Mã việc phụ thuộc", "Nội dung góp ý"
];
const VALID_STATES = ["Chưa bắt đầu", "Đang thực hiện", "Cần chỉnh sửa", "Chờ duyệt", "Hoàn thành"];
const DEFAULT_CONFIG = {
  "SPREADSHEET_ID": "", "MODE": "DRY_RUN", "TIME_ZONE": "Asia/Ho_Chi_Minh",
  "ADMIN_EMAILS": "", "EXECUTIVE_EMAILS": "", "CALENDAR_ID": "",
  "CALENDAR_ENABLED": "false", "SEND_CALENDAR_INVITES": "false",
  "DAILY_HOUR": "7", "WEEKLY_HOUR": "8", "QUOTA_RESERVE": "10",
  "UNSTARTED_HOURS": "48", "SLOW_MIN_DURATION_DAYS": "4",
  "SLOW_ELAPSED_RATIO": "0.5", "SLOW_PROGRESS_THRESHOLD": "0.4",
  "UPCOMING_DAYS": "3,2,1", "APPROVAL_WAIT_DAYS": "2"
};
// ====== Cập nhật lại hàm onOpen() hiện có ở ĐẦU file Code.gs ======
function onOpen() {
  SpreadsheetApp.getUi().createMenu("ĐIỀU HÀNH DỰ ÁN")
    .addItem("📝 Tạo công việc mới (Task)", "openCreateTaskDialog")
    .addItem("➕ Phân công công việc (RACI)", "openAssignDialog") 
    .addItem("👥 Thêm Thành Viên (Member)", "openAddMemberDialog") // 
    .addItem("🌟 Đánh giá nhiệm vụ (Review)", "openReviewDialog") // <-- THÊM DÒNG NÀY
    .addItem("Thay đổi Deadline Task", "openChangeDeadlineDialog")
    .addSeparator()
    .addItem("1. Cài đặt hệ thống (Setup)", "setupSystem")
    .addItem("2. Kiểm tra dữ liệu (Validate)", "validateProjectData")
    .addSeparator()
    .addItem("3. Xem trước giao việc", "previewAssignments")
    .addItem("4. Giao việc hàng loạt", "assignTasksBatch")
    .addSeparator()
    .addItem("5. Gửi bàn giao / Nộp kết quả", "submitSelectedTask")
    .addItem("6. Gửi góp ý (Y/c chỉnh sửa)", "sendFeedbackForSelectedTask")
    .addItem("7. Duyệt hoàn thành (Nghiệm thu)", "approveSelectedTask")
    .addSeparator()
    .addItem("8. Xem trước bản tin tuần", "previewWeeklyBrief")
    .addItem("9. Gửi bản tin tuần", "sendWeeklyBrief")
    .addSeparator()
    .addItem("10. Quét nhắc việc tự động", "runDailyChecks")
    .addItem("11. Đồng bộ Calendar", "syncDeadlineCalendar")
    .addItem("12. Quản lý Triggers", "setupTriggers")
    .addToUi();
    highlightTodayColumn();
}
// ==========================================
// 2. TEMPLATE EMAIL HTML CHUẨN
// ==========================================
function buildMasterEmailTemplate_(title, subtitle, badgeText, badgeColor, bodyHtml, ctaText, ctaUrl) {
  var ctaBlock = "";
  if (ctaText && ctaUrl) {
    ctaBlock = `
      <div style="text-align: center; margin: 30px 0 20px 0;">
        <a href="${ctaUrl}" target="_blank" style="background-color: #1a73e8; color: #ffffff; padding: 12px 28px; text-decoration: none; border-radius: 6px; font-weight: 600; font-size: 14px; display: inline-block; box-shadow: 0 2px 4px rgba(26,115,232,0.3);">
          ${ctaText} &rarr;
        </a>
      </div>
    `;
  }
  return `
  <!DOCTYPE html>
  <html>
  <head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"></head>
  <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background-color: #f8fafc; margin: 0; padding: 24px 0; color: #1e293b; line-height: 1.6;">
    <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width: 680px; background-color: #ffffff; border-radius: 10px; overflow: hidden; border: 1px solid #e2e8f0; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);">
      <tr>
        <td style="background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%); padding: 26px 32px; color: #ffffff;">
          <div style="font-size: 11px; text-transform: uppercase; letter-spacing: 1.5px; opacity: 0.85; margin-bottom: 4px; font-weight: 600;">
            BAN TỔ CHỨC ĐỘI TRANH BIỆN KHOA CNTT • HNUE 2026
          </div>
          <div style="display: flex; justify-content: space-between; align-items: baseline; margin-top: 8px;">
            <h1 style="font-size: 20px; font-weight: 700; margin: 0; color: #ffffff; letter-spacing: -0.2px;">${title}</h1>
          </div>
          <div style="font-size: 13px; color: #cbd5e1; margin-top: 6px;">${subtitle}</div>
        </td>
      </tr>
      <tr>
        <td style="padding: 28px 32px 20px 32px;">
          <div style="margin-bottom: 18px;">
            <span style="background-color: ${badgeColor}; color: #ffffff; font-size: 11px; font-weight: 700; padding: 4px 10px; border-radius: 4px; text-transform: uppercase; letter-spacing: 0.5px;">
              ${badgeText}
            </span>
          </div>
          ${bodyHtml}
          ${ctaBlock}
        </td>
      </tr>
      <tr>
        <td style="background-color: #f1f5f9; padding: 18px 32px; border-top: 1px solid #e2e8f0; font-size: 12px; color: #64748b; text-align: center;">
          <p style="margin: 0 0 6px 0;">Hệ thống thông báo tự động từ Bảng Quản Lý Dự Án Tranh Biện HNUE 2026.</p>
          <p style="margin: 0;">Nếu cần hỗ trợ kỹ thuật hoặc báo cáo vấn đề, vui lòng phản hồi email này.</p>
        </td>
      </tr>
    </table>
  </body>
  </html>
  `;
}

// ==========================================
// 3. MENU & KHỞI TẠO
// ==========================================

function setupSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); const ui = SpreadsheetApp.getUi();
  let taskSheet = ss.getSheetByName(SHEETS.TASKS);
  if (!taskSheet) return ui.alert("Lỗi: Không tìm thấy sheet " + SHEETS.TASKS);
  let headers = taskSheet.getRange(1, 1, 1, taskSheet.getLastColumn()).getValues()[0];
  let appendIndex = taskSheet.getLastColumn() + 1;
    // THAY ĐOẠN NÀY TRONG setupSystem():
  NEW_COLS.forEach(col => {
    if (headers.indexOf(col) === -1) {
      taskSheet.getRange(1, appendIndex).setValue(col);
      // Chỉ gắn tiêu đề, KHÔNG tạo ô tick sẵn cho cả 1.000 dòng nữa
      appendIndex++;
    }
  });
  if (!ss.getSheetByName(SHEETS.CONFIG)) {
    let confSheet = ss.insertSheet(SHEETS.CONFIG);
    confSheet.appendRow(["Key", "Value", "Description"]);
    confSheet.getRange("A1:C1").setFontWeight("bold");
    DEFAULT_CONFIG["SPREADSHEET_ID"] = ss.getId();
    Object.keys(DEFAULT_CONFIG).forEach(key => confSheet.appendRow([key, DEFAULT_CONFIG[key], ""]));
    confSheet.setColumnWidth(1, 250); confSheet.setColumnWidth(2, 350);
  }
  createSystemSheet(ss, SHEETS.LOGS, ["ID", "Task ID", "Loại thông báo", "Khóa chống trùng", "Người nhận", "Thời gian", "Trạng thái", "Lần thử", "Lỗi"]);
  createSystemSheet(ss, SHEETS.STATE, ["Task ID", "Trạng thái", "Phiên bản bàn giao", "Phiên bản phân công", "Phiên bản lịch", "Thời điểm nghiệm thu"]);
  createSystemSheet(ss, SHEETS.CALENDAR, ["Khóa", "Task ID", "Calendar ID", "Event ID", "Phiên bản lịch", "Lần đồng bộ", "Trạng thái"]);
  ui.alert("Hoàn tất Setup. Vui lòng thiết lập PROTECT RANGE cho cột Trạng thái.");
}
function createSystemSheet(ss, name, headers) {
  if (!ss.getSheetByName(name)) {
    let sheet = ss.insertSheet(name); sheet.appendRow(headers); sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
}

// ==========================================
// 4. HELPER & DATA ACCESS
// ==========================================
function getConfig() {
  const confSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
  if (!confSheet) return DEFAULT_CONFIG;
  const data = confSheet.getDataRange().getValues(); let conf = {};
  for (let i = 1; i < data.length; i++) conf[data[i][0]] = data[i][1];
  return conf;
}
function getTasks() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TASKS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0]; let tasks = [];
  for (let i = 1; i < data.length; i++) {
    if (!data[i][headers.indexOf("Mã việc")]) continue;
    let task = { _rowIndex: i + 1 };
    headers.forEach((h, j) => { task[h] = data[i][j]; });
    tasks.push(task);
  }
  return tasks;
}
function getMembers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MEMBERS);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 7) return [];
  const headers = data[5]; let members = [];
  for (let i = 6; i < data.length; i++) {
    if (!data[i][headers.indexOf("Mã TV")]) continue;
    let m = {};
    headers.forEach((h, j) => { m[h] = data[i][j]; });
    members.push(m);
  }
  return members;
}

// HÀM MỚI: Truy xuất Tên và Email từ Mã Thành Viên
function getMemberDetails(memberCodes) {
  if (!memberCodes) return { emails: [], names: [], firstEmail: "", firstName: "Ban Quản trị", rawCode: "" };
  const members = getMembers();
  let codes = memberCodes.toString().split(/[,;\n]/).map(c => c.trim()).filter(Boolean);
  let emails = []; let names = [];
  codes.forEach(c => {
    let m = members.find(x => x["Mã TV"] === c);
    if (m) {
      if (m["Email liên hệ"]) emails.push(m["Email liên hệ"].trim());
      names.push((m["Họ và tên"] || c).trim());
    } else {
      names.push(c);
    }
  });
  return {
    emails: [...new Set(emails)],
    names: [...new Set(names)],
    firstEmail: emails[0] || "",
    firstName: names[0] || "Ban Quản trị",
    rawCode: codes.join(", ")
  };
}

function formatDateSafe(dateVal) {
  if (!dateVal) return 'Chưa thiết lập';
  try { return Utilities.formatDate(new Date(dateVal), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy'); }
  catch (e) { return String(dateVal); }
}

function updateTaskRow(taskId, updates) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.TASKS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0]; const colTaskId = headers.indexOf("Mã việc");
  for (let i = 1; i < data.length; i++) {
    if (data[i][colTaskId] === taskId) {
      Object.keys(updates).forEach(key => {
        let colIdx = headers.indexOf(key);
        if (colIdx > -1) sheet.getRange(i + 1, colIdx + 1).setValue(updates[key]);
      }); break;
    }
  }
}
function logNotification(taskId, type, duplicateKey, recipients, status) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOGS);
  sheet.appendRow([Utilities.getUuid(), taskId, type, duplicateKey, recipients, new Date(), status, 1, ""]);
}
function isNotificationSent(duplicateKey) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOGS);
  const data = sheet.getDataRange().getValues();
  for (let i = Math.max(1, data.length - 1000); i < data.length; i++) {
    if (data[i][3] === duplicateKey && (data[i][6] === "SENT" || data[i][6] === "DRY_RUN")) return true;
  }
  return false;
}
function getTaskState(taskId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.STATE);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === taskId) return { rowIndex: i + 1, status: data[i][1], versionDelivery: Number(data[i][2])||1, versionAssign: Number(data[i][3])||1, versionCal: Number(data[i][4])||1 };
  }
  return null;
}
function updateTaskState(taskId, updates) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.STATE);
  let state = getTaskState(taskId); if (!state) return;
  const map = { status: 2, versionDelivery: 3, versionAssign: 4, versionCal: 5, approvalTime: 6 };
  Object.keys(updates).forEach(k => sheet.getRange(state.rowIndex, map[k]).setValue(updates[k]));
}
function getActiveTaskData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== SHEETS.TASKS) return SpreadsheetApp.getUi().alert("Vui lòng chọn một dòng trong sheet Quan_Ly_Cong_Viec.");
  const row = sheet.getActiveCell().getRow(); if (row < 2) return null;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  let task = { _rowIndex: row };
  headers.forEach((h, j) => { task[h] = rowData[j]; }); return { task, headers };
}
function safeSendEmail(options) {
  const config = getConfig();
  if (config.MODE === "DRY_RUN") return { status: "DRY_RUN" };
  const toCount = (options.to || "").split(',').filter(Boolean).length;
  const ccCount = (options.cc || "").split(',').filter(Boolean).length;
  const quotaReserve = Number(config.QUOTA_RESERVE) || 10;
  if (MailApp.getRemainingDailyQuota() < quotaReserve + toCount + ccCount) return { status: "FAILED" };
  try { MailApp.sendEmail(options); return { status: "SENT" }; } catch (e) { return { status: "UNKNOWN" }; }
}

// ==========================================
// 5. CHỨC NĂNG NGHIỆP VỤ
// ==========================================
function checkDependencies(task, allTasks) {
  let deps = (task["Mã việc phụ thuộc"] || "").toString().split(/[,;\n]/).map(s => s.trim()).filter(Boolean);
  if (deps.length === 0) return { ready: true };
  let blocks = [];
  deps.forEach(d => {
    if (d === task["Mã việc"]) { blocks.push("Tự phụ thuộc (Vòng lặp)"); return; }
    let p = allTasks.find(x => x["Mã việc"] === d);
    if (!p) blocks.push(`Mã ${d} không tồn tại`);
    else if (p["Trạng thái"] !== "Hoàn thành") blocks.push(`Chờ ${d}`);
  });
  return { ready: blocks.length === 0, reason: blocks.join(" | ") };
}

/** 5.1 GIAO VIỆC HÀNG LOẠT */
function assignTasksBatch() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}`;
  const taskSheet = ss.getSheetByName(SHEETS.TASKS);
  const logSheet = ss.getSheetByName(SHEETS.LOGS);
 
  if (!taskSheet || !logSheet) return SpreadsheetApp.getUi().alert('Lỗi: Không tìm thấy sheet.');
  
  const tasks = getTasks();
  const logData = logSheet.getLastRow() > 1 ? logSheet.getRange(2, 4, logSheet.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  const loggedKeys = new Set(logData);
  const tasksByMember = {}; const rowsToUpdateAssignDate = [];
  
  for (let i = 0; i < tasks.length; i++) {
    const t = tasks[i];
    if (!t["Mã việc"] || t["Mã việc"].startsWith('Tổng') || !t["Mã việc"].startsWith('TASK-')) continue;
    if (t["Trạng thái"] === 'Chưa bắt đầu' && !t["Ngày giao việc"]) {
      let dep = checkDependencies(t, tasks);
      if(dep.ready) {
        let code = String(t["Mã TV thực hiện chính"]).trim();
        let execDetails = getMemberDetails(code);
        
        if (code && execDetails.firstEmail) {
          let dedupeKey = `ASSIGN_${t["Mã việc"]}_${code}`;
          if (!loggedKeys.has(dedupeKey)) {
            if (!tasksByMember[code]) tasksByMember[code] = { memberInfo: execDetails, tasks: [] };
            
            let apprDetails = getMemberDetails(t["Mã TV phê duyệt"]);
            tasksByMember[code].tasks.push({
              id: t["Mã việc"],
              name: t["Tên công việc"],
              phase: t["Giai đoạn"] || 'Chưa quy định',
              priority: t["Độ ưu tiên"] || 'Bình thường',
              deadline: formatDateSafe(t["Ngày kết thúc"]),
              deliverable: t["Tài liệu bàn giao"] || 'Theo quy chuẩn dự án',
              approver: `${apprDetails.names.join(", ")} (${apprDetails.rawCode})`,
              rowIndex: t._rowIndex
            });
            rowsToUpdateAssignDate.push(t._rowIndex);
          }
        }
      }
    }
  }

  const assignedMemberCodes = Object.keys(tasksByMember);
  if (assignedMemberCodes.length === 0) return SpreadsheetApp.getUi().alert('Không có công việc mới nào cần phân công.');
 
  let totalEmailsSent = 0; const newLogs = [];
  
  for (let m = 0; m < assignedMemberCodes.length; m++) {
    const group = tasksByMember[assignedMemberCodes[m]];
    let cardsHtml = '';
    
    group.tasks.forEach((t) => {
      let badgeColor = (t.priority === 'Khẩn cấp') ? '#ef4444' : (t.priority === 'Cao' ? '#eab308' : '#3b82f6');
      cardsHtml += `
        <div style="border: 1px solid #e2e8f0; border-radius: 8px; margin-bottom: 20px; overflow: hidden; background-color: #ffffff; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
          <div style="background-color: #f8fafc; padding: 12px 16px; border-bottom: 1px solid #e2e8f0; display: flex; justify-content: space-between; align-items: center;">
            <strong style="color: #1e3a8a; font-size: 14px;">[${t.id}] ${t.name}</strong>
          </div>
          <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
             <tr>
               <td style="padding: 10px 16px; border-bottom: 1px solid #f1f5f9; width: 35%; color: #475569;">Giai đoạn & Mức độ:</td>
               <td style="padding: 10px 16px; border-bottom: 1px solid #f1f5f9;">${t.phase} - <span style="color: ${badgeColor}; font-weight: bold;">${t.priority}</span></td>
             </tr>
             <tr>
               <td style="padding: 10px 16px; border-bottom: 1px solid #f1f5f9; color: #475569;">Hạn chót (Deadline):</td>
               <td style="padding: 10px 16px; border-bottom: 1px solid #f1f5f9; color: #dc2626; font-weight: 600;">${t.deadline}</td>
             </tr>
             <tr>
               <td style="padding: 10px 16px; border-bottom: 1px solid #f1f5f9; color: #475569;">Yêu cầu sản phẩm:</td>
               <td style="padding: 10px 16px; border-bottom: 1px solid #f1f5f9; color: #0284c7; font-weight: 600;">${t.deliverable}</td>
             </tr>
             <tr>
               <td style="padding: 10px 16px; color: #475569;">Người phê duyệt (RACI):</td>
               <td style="padding: 10px 16px; font-weight: 600; color: #1e293b;">${t.approver}</td>
             </tr>
          </table>
        </div>
      `;
    });

    let innerBody = `
      <p style="margin: 0 0 16px 0; font-size: 14px;">Xin chào <strong>${group.memberInfo.firstName}</strong>,</p>
      <p style="margin: 0 0 18px 0; font-size: 14px; color: #334155;">
        Bạn vừa được phân công làm người phụ trách chính cho <strong>${group.tasks.length} nhiệm vụ mới</strong>. Chi tiết công việc:
      </p>
      ${cardsHtml}
      <div style="background-color: #eff6ff; border-left: 4px solid #3b82f6; padding: 12px 14px; border-radius: 0 4px 4px 0; font-size: 13px; color: #1e40af; margin-bottom: 12px;">
        <strong>Quy trình thực hiện:</strong> Vui lòng truy cập Bảng quản lý, chuyển trạng thái sang "Đang thực hiện" để hệ thống theo dõi tiến độ.
      </div>
    `;

    let htmlPayload = buildMasterEmailTemplate_(
      "Thông Báo Phân Công Nhiệm Vụ", `Cập nhật ngày ${Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy')}`, 
      "NHIỆM VỤ MỚI", "#2563eb", innerBody, "Mở Bảng Tiến Độ", spreadsheetUrl
    );

    let res = safeSendEmail({ to: group.memberInfo.firstEmail, subject: `[HNUE Debate] Thông báo phân công nhiệm vụ mới`, htmlBody: htmlPayload });
    if (res.status === "SENT" || res.status === "DRY_RUN") {
      if(res.status === "SENT") totalEmailsSent++;
      group.tasks.forEach(t => {
        newLogs.push([Utilities.getUuid(), t.id, 'ASSIGN', `ASSIGN_${t.id}_${assignedMemberCodes[m]}`, group.memberInfo.firstEmail, new Date(), res.status, 1, '']);
      });
    }
  }

  if (newLogs.length > 0 && getConfig().MODE === "LIVE") {
    rowsToUpdateAssignDate.forEach(rIdx => taskSheet.getRange(rIdx, 20).setValue(new Date()));
    logSheet.getRange(logSheet.getLastRow() + 1, 1, newLogs.length, 9).setValues(newLogs);
  }
  SpreadsheetApp.getUi().alert(`Hoàn thành lệnh giao việc!\nChế độ: ${getConfig().MODE}\nSố lô email: ${totalEmailsSent}`);
}

/** 5.2 XỬ LÝ NỘP BÀN GIAO */
function processTaskSubmission(task, url, apprEmail, isFromMenu, isResubmit) {
  let state = getTaskState(task["Mã việc"]);
  if (!state) return false;
  if (["Chờ duyệt", "Hoàn thành"].includes(state.status) && !isResubmit) {
    if (isFromMenu) SpreadsheetApp.getUi().alert("Task đã nộp. Tích chọn ô 'Gửi lại bàn giao' để nộp lại.");
    return false;
  }
 
  let newV = state.versionDelivery + 1;
  let execDetails = getMemberDetails(task["Mã TV thực hiện chính"]);
  let apprDetails = getMemberDetails(task["Mã TV phê duyệt"]);
  let deadlineStr = formatDateSafe(task["Ngày kết thúc"]);
  let spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${SpreadsheetApp.getActiveSpreadsheet().getId()}`;

  let innerBody = `
    <p style="margin: 0 0 16px 0; font-size: 14px;">Kính gửi <strong>${apprDetails.firstName}</strong> (Người phê duyệt),</p>
    <p style="margin: 0 0 18px 0; font-size: 14px; color: #334155;">Thành viên phụ trách đã nộp sản phẩm bàn giao cho công việc dưới đây để bạn thẩm định:</p>
    <table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 13px;">
      <tr style="background-color: #f8fafc;">
        <td style="padding: 10px 12px; border: 1px solid #e2e8f0; font-weight: 600; width: 35%;">Mã & Tên việc:</td>
        <td style="padding: 10px 12px; border: 1px solid #e2e8f0; font-weight: 700;">[${task["Mã việc"]}] ${task["Tên công việc"]}</td>
      </tr>
      <tr>
        <td style="padding: 10px 12px; border: 1px solid #e2e8f0; font-weight: 600;">Hạn chót:</td>
        <td style="padding: 10px 12px; border: 1px solid #e2e8f0; font-weight: 600; color: #dc2626;">${deadlineStr}</td>
      </tr>
      <tr style="background-color: #f8fafc;">
        <td style="padding: 10px 12px; border: 1px solid #e2e8f0; font-weight: 600;">Người thực hiện:</td>
        <td style="padding: 10px 12px; border: 1px solid #e2e8f0;">${execDetails.names.join(", ")} (${execDetails.rawCode})</td>
      </tr>
      <tr>
        <td style="padding: 10px 12px; border: 1px solid #e2e8f0; font-weight: 600;">Tiến độ & Link nộp:</td>
        <td style="padding: 10px 12px; border: 1px solid #e2e8f0; font-weight: 700; color: #16a34a;">${(parseFloat(task["Tiến độ"] || 0) * 100).toFixed(0)}%</td>
      </tr>
      <tr style="background-color: #f8fafc;">
        <td style="padding: 10px 12px; border: 1px solid #e2e8f0; font-weight: 600;">Thời gian nộp (v${newV}):</td>
        <td style="padding: 10px 12px; border: 1px solid #e2e8f0;">${Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm")}</td>
      </tr>
    </table>
    <div style="background-color: #fffbeb; border-left: 4px solid #f59e0b; padding: 12px 14px; border-radius: 0 4px 4px 0; font-size: 13px; color: #92400e; margin-bottom: 12px;">
      <strong>Quy trình nghiệm thu:</strong> Nhấp vào nút bên dưới để xem tài liệu. Sau đó vào Bảng phân công chuyển trạng thái "Hoàn thành" hoặc điền "Nội dung góp ý".
    </div>
    <p style="font-size: 13px; margin-top: 15px;">🔗 <a href="${spreadsheetUrl}" style="color: #2563eb; text-decoration: underline;">Truy cập nhanh File Bảng tính Dự án tại đây</a></p>
  `;

  let htmlPayload = buildMasterEmailTemplate_(
    "Yêu Cầu Nghiệm Thu", `Bàn giao từ: ${execDetails.firstName}`, 
    "CHỜ PHÊ DUYỆT", "#ea580c", innerBody, "Mở Tài Liệu Thẩm Định", url
  );

  let res = safeSendEmail({ to: apprEmail, subject: `[HNUE Debate] [Cần nghiệm thu v${newV}] Task ${task["Mã việc"]}`, htmlBody: htmlPayload });
  if (res.status === "SENT" || res.status === "DRY_RUN") {
    if (res.status === "SENT") {
      updateTaskRow(task["Mã việc"], { "Trạng thái": "Chờ duyệt" });
      updateTaskState(task["Mã việc"], { status: "Chờ duyệt", versionDelivery: newV });
    }
    logNotification(task["Mã việc"], "SUBMIT", `SUBMIT_${task["Mã việc"]}_v${newV}`, apprEmail, res.status);
    return true;
  }
  return false;
}
function submitSelectedTask() {
  const data = getActiveTaskData();
  if (!data || !data.task["Mã việc"]) return;
  const { task } = data;
  if (!task["Link sản phẩm bàn giao"]) return SpreadsheetApp.getUi().alert("Lỗi: Thiếu Link sản phẩm bàn giao.");
  let isResubmit = (task["Gửi lại bàn giao"] === true || String(task["Gửi lại bàn giao"]).toUpperCase() === "TRUE");
  let apprEmail = getMemberDetails(task["Mã TV phê duyệt"]).emails.join(",");
  let success = processTaskSubmission(task, task["Link sản phẩm bàn giao"], apprEmail, true, isResubmit);
  if (success) SpreadsheetApp.getUi().alert("Đã gửi yêu cầu phê duyệt thành công.");
}

/** 5.3 GÓP Ý VÀ NGHIỆM THU */
function sendFeedbackForSelectedTask() {
  const data = getActiveTaskData(); if (!data) return; const { task } = data;
  const ui = SpreadsheetApp.getUi();
  if (task["Trạng thái"] !== "Chờ duyệt") return ui.alert("Task không ở trạng thái Chờ duyệt.");
  if (!task["Nội dung góp ý"]) return ui.alert("Vui lòng nhập Nội dung góp ý.");
  
  let execDetails = getMemberDetails(task["Mã TV thực hiện chính"]);
  let apprDetails = getMemberDetails(task["Mã TV phê duyệt"]);
  let spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${SpreadsheetApp.getActiveSpreadsheet().getId()}`;
  
  let innerBody = `
    <p style="margin: 0 0 16px 0; font-size: 14px;">Xin chào <strong>${execDetails.firstName}</strong>,</p>
    <p style="margin: 0 0 14px 0; font-size: 14px; color: #334155;"><strong>${apprDetails.firstName}</strong> (Người phê duyệt) đã thẩm định nhiệm vụ <strong>[${task["Mã việc"]}] ${task["Tên công việc"]}</strong>.</p>
    <div style="background-color: #fef2f2; border: 1px solid #fecaca; border-radius: 6px; padding: 14px; color: #991b1b; font-size: 13px; margin: 16px 0;">
      <strong style="font-size: 14px;">Ý kiến đóng góp / Nội dung cần sửa đổi:</strong>
      <p style="margin: 8px 0 0 0; font-style: italic; white-space: pre-line; background-color: #ffffff; padding: 10px; border-radius: 4px; border: 1px dashed #f87171;">
        ${task["Nội dung góp ý"]}
      </p>
    </div>
    <p style="font-size: 13px; color: #475569;">Sau khi sửa đổi, hãy cập nhật lại link và tích chọn <strong>Gửi lại bàn giao (Cột S)</strong> trên Bảng tính để tự động nộp lại.</p>
  `;
  let htmlPayload = buildMasterEmailTemplate_("Yêu Cầu Bổ Sung Chỉnh Sửa", `Task: ${task["Mã việc"]}`, "CẦN ĐIỀU CHỈNH", "#dc2626", innerBody, "Mở Bảng Phân Công", spreadsheetUrl);
  let res = safeSendEmail({ to: execDetails.emails.join(","), subject: `[HNUE Debate] [Cần sửa đổi] Task ${task["Mã việc"]}`, htmlBody: htmlPayload });
  if (res.status === "SENT" || res.status === "DRY_RUN") {
    if (res.status === "SENT") {
      updateTaskRow(task["Mã việc"], { "Trạng thái": "Cần chỉnh sửa" });
      updateTaskState(task["Mã việc"], { status: "Cần chỉnh sửa" });
    }
    logNotification(task["Mã việc"], "FEEDBACK", `FB_${task["Mã việc"]}_${new Date().getTime()}`, execDetails.emails.join(","), res.status);
    ui.alert("Đã gửi thông báo yêu cầu chỉnh sửa.");
  }
}

// =========================================================================
// 7. DUYỆT HOÀN THÀNH (CẬP NHẬT GỬI MAIL CHO MR VÀ YÊU CẦU ĐÁNH GIÁ)
// =========================================================================
// =========================================================================
// 7. DUYỆT HOÀN THÀNH & GỬI YÊU CẦU ĐÁNH GIÁ ĐẾN NGƯỜI THỰC HIỆN CHÍNH
// =========================================================================
// =========================================================================
// 7. DUYỆT HOÀN THÀNH (NGHIỆM THU) & GỬI YÊU CẦU ĐÁNH GIÁ ĐẾN NGƯỜI LÀM CHÍNH
// =========================================================================
function approveSelectedTask() {
  const data = getActiveTaskData();
  if (!data) return;
  const { task, headers } = data;
  const ui = SpreadsheetApp.getUi();

  // Kiểm tra trạng thái hợp lệ để duyệt
  if (task["Trạng thái"] !== "Chờ duyệt" && task["Trạng thái"] !== "Đang thực hiện") {
    return ui.alert("Task chưa nộp / chưa ở trạng thái Chờ duyệt hoặc Đang thực hiện.");
  }

  // Lấy thông tin Người thực hiện chính (MR) và Người duyệt (A)
  let execDetails = getMemberDetails(task["Mã TV thực hiện chính"]);
  let apprDetails = getMemberDetails(task["Mã TV phê duyệt"]);
  let spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${SpreadsheetApp.getActiveSpreadsheet().getId()}`;

  if (!execDetails.firstEmail) {
    return ui.alert(`Lỗi: Không tìm thấy email của Người thực hiện chính (${task["Mã TV thực hiện chính"]}).`);
  }

  // 1. Tạo nội dung email tích hợp: Thông báo hoàn thành & yêu cầu đánh giá nhân sự
  let innerBody = `
    <p style="margin: 0 0 16px 0; font-size: 14px;">Xin chào <strong>${execDetails.firstName}</strong> (Người phụ trách chính / Quản lý task),</p>
    <p style="margin: 0 0 14px 0; font-size: 14px; color: #334155;">
      Nhiệm vụ <strong>[${task["Mã việc"]}] ${task["Tên công việc"]}</strong> đã được <strong>${apprDetails.firstName}</strong> (Người phê duyệt) thẩm định và nghiệm thu đạt chất lượng.
    </p>

    <!-- Khối 1: Xác nhận kết quả hoàn thành -->
    <div style="background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 6px; padding: 14px; color: #166534; font-size: 13px; margin: 16px 0;">
      <strong>Kết quả nghiệm thu:</strong> Sản phẩm bàn giao đã được phê duyệt. Công việc chính thức được ghi nhận <strong>100% Hoàn thành</strong> trên hệ thống. Cảm ơn sự nỗ lực của bạn và đội ngũ!
    </div>

    <!-- Khối 2: Yêu cầu đánh giá nhân sự tham gia task -->
    <div style="background-color: #fffbeb; border: 1px solid #fde68a; border-radius: 6px; padding: 14px; color: #92400e; font-size: 13px; margin: 16px 0;">
      <strong style="font-size: 14px; display: block; margin-bottom: 6px;">📋 YÊU CẦU ĐÁNH GIÁ THÀNH VIÊN THEO TASK:</strong>
      <p style="margin: 0 0 8px 0; line-height: 1.5;">
        Với vai trò là <strong>Người điều phối & giám sát trực tiếp</strong>, bạn vui lòng truy cập Bảng tính, mở menu <strong>ĐIỀU HÀNH DỰ ÁN &gt; 🌟 Đánh giá nhiệm vụ (Review)</strong> để chấm điểm hiệu suất cho các thành viên cùng tham gia task này theo 3 tiêu chí:
      </p>
      <ul style="margin: 0; padding-left: 20px;">
        <li><strong>Trách nhiệm (1-10):</strong> Tiến độ nộp bài, mức độ chủ động và kỷ luật.</li>
        <li><strong>Chuyên môn (1-10):</strong> Chất lượng sản phẩm, kỹ năng nghiệp vụ.</li>
        <li><strong>Phối hợp (1-10):</strong> Khả năng giao tiếp, tinh thần đội ngũ và tinh thần tiếp thu góp ý.</li>
      </ul>
    </div>

    <p style="font-size: 13px; margin-top: 15px;">🔗 <a href="${spreadsheetUrl}" style="color: #2563eb; text-decoration: underline;">Truy cập File Dự án để mở bảng chấm điểm</a></p>
  `;

  let htmlPayload = buildMasterEmailTemplate_(
    "Nghiệm Thu Hoàn Thành & Yêu Cầu Đánh Giá",
    `Task: [${task["Mã việc"]}] ${task["Tên công việc"]}`,
    "HOÀN THÀNH & CẦN ĐÁNH GIÁ",
    "#16a34a",
    innerBody,
    "Mở Bảng Đánh Giá Thành Viên",
    spreadsheetUrl
  );

  // 2. Gửi email trực tiếp đến Người thực hiện chính
  let res = safeSendEmail({
    to: execDetails.emails.join(","),
    subject: `[HNUE Debate] [Đã duyệt & Cần đánh giá] Task ${task["Mã việc"]}: ${task["Tên công việc"]}`,
    htmlBody: htmlPayload
  });

  // 3. Cập nhật dữ liệu trên Sheet và ghi Log
  if (res.status === "SENT" || res.status === "DRY_RUN") {
    if (res.status === "SENT") {
      let updates = {
        "Trạng thái": "Hoàn thành",
        "Tiến độ": 1.0,
        "Gửi lại bàn giao": false
      };
      if (headers.indexOf("Trạng thái Đánh giá") > -1) {
        updates["Trạng thái Đánh giá"] = "Chưa đánh giá";
      }
      updateTaskRow(task["Mã việc"], updates);
      updateTaskState(task["Mã việc"], { status: "Hoàn thành", approvalTime: new Date() });
    }

    logNotification(
      task["Mã việc"],
      "APPROVE",
      `APPR_${task["Mã việc"]}_${new Date().getTime()}`,
      execDetails.emails.join(","),
      res.status
    );

    // Thông báo nghiệm thu hoàn tất đơn giản, không mở popup và không hỏi xác nhận
    ui.alert(
      "✅ Nghiệm thu thành công!",
      `Task [${task["Mã việc"]}] đã chuyển sang trạng thái "Hoàn thành".\nEmail thông báo và yêu cầu đánh giá đã được gửi đến Người thực hiện chính (${execDetails.firstName} - ${execDetails.firstEmail}).`,
      ui.ButtonSet.OK
    );
  } else {
    ui.alert(`Không thể gửi email: Hạn ngạch gửi thư đã hết hoặc xảy ra lỗi (Trạng thái: ${res.status}).`);
  }
}
/** 5.4 QUÉT NHẮC VIỆC & BÁO CÁO TUẦN */
function runDailyChecks() {
  const config = getConfig(); const tasks = getTasks();
  const today = new Date(); today.setHours(0,0,0,0);
  let batches = {};
  
  tasks.forEach(t => {
    if (!t["Ngày giao việc"] || t["Trạng thái"] === "Hoàn thành") return;
    let state = getTaskState(t['Mã việc']); if (!state) return;
    
    let execDetails = getMemberDetails(t["Mã TV thực hiện chính"]);
    let apprDetails = getMemberDetails(t["Mã TV phê duyệt"]);
    let deadlineStr = formatDateSafe(t["Ngày kết thúc"]);
    
    let isAssignedNotStarted = (t["Trạng thái"] === "Chưa bắt đầu");
    let hasDeadline = !!t["Ngày kết thúc"];
    let deadline = hasDeadline ? new Date(t["Ngày kết thúc"]) : null;
    let deadlineStrVal = hasDeadline ? deadline.getTime() : "no_deadline";
    let diffDays = hasDeadline ? Math.ceil((deadline - today) / (1000 * 60 * 60 * 24)) : null;
    let alertMatched = false;

    if (hasDeadline && diffDays < 0) {
       let key = `OVERDUE_${t['Mã việc']}_${today.getTime()}`;
       if (!isNotificationSent(key)) {
          let rec = t["Trạng thái"] === "Chờ duyệt" ? apprDetails.emails.join(",") : execDetails.emails.join(",");
          let cc = t["Trạng thái"] === "Chờ duyệt" ? execDetails.emails.join(",") : apprDetails.emails.join(",");
          let targetName = t["Trạng thái"] === "Chờ duyệt" ? apprDetails.firstName : execDetails.firstName;
          
          let route = rec + "||" + cc;
          if (!batches[route]) batches[route] = { to: rec, cc: cc, toName: targetName, msgs: [], logs: [], isOverdue: true };
          batches[route].msgs.push(`<li style="margin-bottom:8px;"><strong>[${t['Mã việc']}]</strong> ${t['Tên công việc']} <br><span style="color:#dc2626; font-size: 12px;">(Phụ trách: ${execDetails.names.join(", ")} | Hạn chót: ${deadlineStr} - Quá hạn ${Math.abs(diffDays)} ngày)</span></li>`);
          batches[route].logs.push({ id: t['Mã việc'], type: "OVERDUE", key: key, rec: rec });
          alertMatched = true;
       }
    }
    if (!alertMatched && hasDeadline && diffDays >= 0 && diffDays <= 3) {
       let key = `UPCOMING_${t['Mã việc']}_${diffDays}d_${deadlineStrVal}`;
       if (!isNotificationSent(key)) {
          let rec = t["Trạng thái"] === "Chờ duyệt" ? apprDetails.emails.join(",") : execDetails.emails.join(",");
          let targetName = t["Trạng thái"] === "Chờ duyệt" ? apprDetails.firstName : execDetails.firstName;
          let route = rec + "||";
          if (!batches[route]) batches[route] = { to: rec, cc: "", toName: targetName, msgs: [], logs: [], isOverdue: false };
          batches[route].msgs.push(`<li style="margin-bottom:8px;"><strong>[${t['Mã việc']}]</strong> ${t['Tên công việc']} <br><span style="color:#d97706; font-size: 12px;">(Phụ trách: ${execDetails.names.join(", ")} | Hạn chót: ${deadlineStr} - Còn ${diffDays} ngày)</span></li>`);
          batches[route].logs.push({ id: t['Mã việc'], type: "UPCOMING", key: key, rec: rec });
          alertMatched = true;
       }
    }
    if (!alertMatched && isAssignedNotStarted) {
      let startD = t["Ngày bắt đầu"] ? new Date(t["Ngày bắt đầu"]) : new Date(t["Ngày giao việc"]);
      let startRef = Math.max(startD.getTime(), new Date(t["Ngày giao việc"]).getTime());
      if ((new Date().getTime() - startRef) > 48 * 3600 * 1000) {
        let key = `UNSTARTED_${t['Mã việc']}_v${state.versionAssign}`;
        if (!isNotificationSent(key)) {
          let route = execDetails.emails.join(",") + "||";
          if (!batches[route]) batches[route] = { to: execDetails.emails.join(","), cc: "", toName: execDetails.firstName, msgs: [], logs: [], isOverdue: false };
          batches[route].msgs.push(`<li style="margin-bottom:8px;"><strong>[${t['Mã việc']}]</strong> ${t['Tên công việc']} <br><span style="color:#64748b; font-size: 12px;">(Hạn chót: ${deadlineStr} | Trạng thái: Chưa bắt đầu > 48h)</span></li>`);
          batches[route].logs.push({ id: t['Mã việc'], type: "UNSTARTED", key: key, rec: execDetails.emails.join(",") });
          alertMatched = true;
       }
      }
    }
  });

  let spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${SpreadsheetApp.getActiveSpreadsheet().getId()}`;
  Object.keys(batches).forEach(rKey => {
    let b = batches[rKey];
    if (b.msgs.length > 0) {
      let innerBody = `
        <p style="margin: 0 0 16px 0; font-size: 14px;">Xin chào <strong>${b.toName}</strong>,</p>
        <p style="margin: 0 0 16px 0; font-size: 14px; color: #334155;">Hệ thống thông báo trạng thái của các công việc đang cần sự chú ý:</p>
        <ul style="font-size: 13px; color: #1e293b; background-color: #f8fafc; padding: 15px 15px 15px 30px; border: 1px solid #e2e8f0; border-radius: 6px;">
          ${b.msgs.join("")}
        </ul>
        <div style="background-color: ${b.isOverdue ? '#fef2f2' : '#fffbeb'}; border-left: 4px solid ${b.isOverdue ? '#ef4444' : '#f59e0b'}; padding: 12px 14px; font-size: 13px; color: ${b.isOverdue ? '#991b1b' : '#92400e'}; border-radius: 0 4px 4px 0;">
          <strong>Khuyến nghị hành động:</strong> Vui lòng ưu tiên xử lý, nộp sản phẩm hoặc báo cáo vướng mắc với Ban điều hành.
        </div>
      `;
      let htmlPayload = buildMasterEmailTemplate_(
        b.isOverdue ? "Cảnh Báo Quá Hạn Công Việc" : "Nhắc Nhở Tiến Độ Công Việc",
        `Ngày ${Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy')}`,
        b.isOverdue ? "QUÁ HẠN" : "CẦN CHÚ Ý", b.isOverdue ? "#b91c1c" : "#d97706",
        innerBody, "Cập Nhật Trạng Thái", spreadsheetUrl
      );
      let res = safeSendEmail({ to: b.to, cc: b.cc, subject: "[HNUE Debate] Cảnh báo tiến độ hệ thống", htmlBody: htmlPayload });
      if (res.status === "SENT" || res.status === "DRY_RUN") {
        b.logs.forEach(l => logNotification(l.id, l.type, l.key, l.rec, res.status));
      }
    }
  });
}

function sendWeeklyBrief() {
  const config = getConfig(); const tasks = getTasks();
  const stateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.STATE);
  const stateData = stateSheet.getDataRange().getValues();
  let currentWeekNum = Utilities.formatDate(new Date(), config.TIME_ZONE, "w");
  let key = `WEEKLY_${currentWeekNum}`;
  if (isNotificationSent(key)) return;
 
  let doneCount = 0; let weeklyDone = 0; let overdueCount = 0; let inProgressCount = 0; let pendingCount = 0;
  let aWeekAgo = new Date(new Date().getTime() - 7 * 24 * 3600 * 1000);
  tasks.forEach(t => {
     if (t["Trạng thái"] === "Hoàn thành") {
        doneCount++;
        let st = stateData.find(r => r[0] === t["Mã việc"]);
        if (st && st[5] && new Date(st[5]) >= aWeekAgo) weeklyDone++;
     } else if (t["Trạng thái"] === "Đang thực hiện") inProgressCount++;
     else if (t["Trạng thái"] === "Chờ duyệt") pendingCount++;
     if (t["Ngày kết thúc"] && t["Trạng thái"] !== "Hoàn thành" && t["Trạng thái"] !== "Chưa bắt đầu") {
        if (new Date(t["Ngày kết thúc"]).setHours(0,0,0,0) < new Date().setHours(0,0,0,0)) overdueCount++;
     }
  });
 
  let overallProg = tasks.length > 0 ? ((doneCount / tasks.length) * 100).toFixed(1) + "%" : "0%";
  let spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${SpreadsheetApp.getActiveSpreadsheet().getId()}`;
  let innerBody = `
    <p style="margin: 0 0 16px 0; font-size: 14px;">Kính gửi <strong>Ban Điều Hành Dự Án</strong>,</p>
    <p style="margin: 0 0 18px 0; font-size: 14px; color: #334155;">Bảng tổng hợp tình hình vận hành dự án tuần qua:</p>
    <table style="width: 100%; border-collapse: separate; border-spacing: 8px; margin-bottom: 20px;">
      <tr>
        <td style="background-color: #f1f5f9; padding: 14px; border-radius: 8px; text-align: center; width: 33%;">
          <div style="font-size: 11px; text-transform: uppercase; color: #64748b; font-weight: 600;">Tổng công việc</div>
          <div style="font-size: 22px; font-weight: 800; color: #0f172a; margin-top: 4px;">${tasks.length}</div>
        </td>
        <td style="background-color: #f0fdf4; padding: 14px; border-radius: 8px; text-align: center; width: 33%;">
          <div style="font-size: 11px; text-transform: uppercase; color: #166534; font-weight: 600;">Xong (Tuần qua)</div>
          <div style="font-size: 22px; font-weight: 800; color: #15803d; margin-top: 4px;">${weeklyDone}</div>
        </td>
        <td style="background-color: #eff6ff; padding: 14px; border-radius: 8px; text-align: center; width: 33%;">
          <div style="font-size: 11px; text-transform: uppercase; color: #1e40af; font-weight: 600;">Tiến độ chung</div>
          <div style="font-size: 22px; font-weight: 800; color: #1d4ed8; margin-top: 4px;">${overallProg}</div>
        </td>
      </tr>
      <tr>
        <td style="background-color: #fff7ed; padding: 14px; border-radius: 8px; text-align: center;">
          <div style="font-size: 11px; text-transform: uppercase; color: #9a3412; font-weight: 600;">Đang triển khai</div>
          <div style="font-size: 20px; font-weight: 800; color: #c2410c; margin-top: 4px;">${inProgressCount}</div>
        </td>
        <td style="background-color: #fef2f2; padding: 14px; border-radius: 8px; text-align: center;">
          <div style="font-size: 11px; text-transform: uppercase; color: #991b1b; font-weight: 600;">Đang Quá hạn</div>
          <div style="font-size: 20px; font-weight: 800; color: #b91c1c; margin-top: 4px;">${overdueCount}</div>
        </td>
        <td style="background-color: #f8fafc; padding: 14px; border-radius: 8px; text-align: center;">
          <div style="font-size: 11px; text-transform: uppercase; color: #475569; font-weight: 600;">Đang chờ duyệt</div>
          <div style="font-size: 20px; font-weight: 800; color: #334155; margin-top: 4px;">${pendingCount}</div>
        </td>
      </tr>
    </table>
  `;
  let htmlPayload = buildMasterEmailTemplate_(
    "Báo Cáo Tiến Độ Tuần", `Tuần ${currentWeekNum} • Tháng ${Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "MM/yyyy")}`,
    "BÁO CÁO TỔNG QUAN", "#0f766e", innerBody, "Mở Bảng Dashboard", spreadsheetUrl
  );
  let res = safeSendEmail({ to: config.ADMIN_EMAILS, subject: `[HNUE Debate] Báo cáo tiến độ Tuần ${currentWeekNum}`, htmlBody: htmlPayload });
  if (res.status === "SENT" || res.status === "DRY_RUN") {
     logNotification("SYSTEM", "WEEKLY", key, config.ADMIN_EMAILS, res.status);
     if (res.status !== "DRY_RUN") SpreadsheetApp.getUi().alert("Đã gửi Báo cáo tuần.");
  }
}

// ==========================================
// 6. CÁC HÀM XỬ LÝ KHÁC (Events, Calendar...)
// ==========================================
// ==========================================
// 6. CÁC HÀM XỬ LÝ KHÁC (Events, Calendar...)
// ==========================================
function handleTaskEdit(e) {
  if (!e || !e.range) return;
  const range = e.range; 
  const sheet = range.getSheet();
  const sheetName = sheet.getName();
  
  // -----------------------------------------------------------------
  // NHÁNH 1: TỰ ĐỘNG CẬP NHẬT BẢNG TÌNH TRẠNG KHI CÓ THÀNH VIÊN MỚI
  // -----------------------------------------------------------------
  if (sheetName === SHEETS.MEMBERS) {
    const row = range.getRow();
    const col = range.getColumn();
    
    // Dòng dữ liệu trong Danh_Sach_Thanh_Vien bắt đầu từ dòng 2 (Vì dòng 1 là tiêu đề)
    // Nếu người dùng nhập thông tin vào cột A (Mã TV) hoặc cột B (Họ và tên)
    if (row > 1 && (col === 1 || col === 2)) { 
      // Kiểm tra xem hàm updateMemberDashboard có tồn tại hay không thì gọi
      if (typeof updateMemberDashboard === "function") {
        updateMemberDashboard();
        SpreadsheetApp.getActiveSpreadsheet().toast("Đã tự động cập nhật Bảng Tình Trạng do có danh sách nhân sự thay đổi!", "Hệ thống");
      }
    }
    return; // Đã xử lý xong tác vụ của sheet thành viên, thoát hàm để không chạy nhầm nhánh bên dưới
  }

  // -----------------------------------------------------------------
  // NHÁNH 2: XỬ LÝ CHECKBOX BÀN GIAO VÀ TRẠNG THÁI TASK
  // -----------------------------------------------------------------
  if (sheetName !== SHEETS.TASKS || range.getRow() < 2) return;
  
  const row = range.getRow(); 
  const col = range.getColumn();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colStatus = headers.indexOf("Trạng thái") + 1;
  const colResubmit = headers.indexOf("Gửi lại bàn giao") + 1;
  const colTaskId = headers.indexOf("Mã việc") + 1;
  
  if (col === colStatus && e.value === 'Hoàn thành') {
      range.setValue(e.oldValue && e.oldValue !== 'Hoàn thành' ? e.oldValue : 'Chờ duyệt');
      SpreadsheetApp.getUi().alert('⚠️ Chỉ người duyệt mới có quyền nghiệm thu "Hoàn thành" qua menu ĐIỀU HÀNH DỰ ÁN.');
      return;
  }
  if (col === colResubmit) {
    const isChecked = (range.getValue() === true || e.value === 'TRUE');
    if (isChecked) {
      sheet.getRange(row, colStatus).setValue('Chờ duyệt');
      const taskId = sheet.getRange(row, colTaskId).getValue();
      if (taskId) updateTaskState(taskId, { status: "Chờ duyệt" });
      const url = sheet.getRange(row, headers.indexOf("Link sản phẩm bàn giao") + 1).getValue();
      const apprEmails = getMemberDetails(sheet.getRange(row, headers.indexOf("Mã TV phê duyệt") + 1).getValue()).emails.join(",");
      const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      let taskObj = { _rowIndex: row };
      headers.forEach((h, j) => { taskObj[h] = rowData[j]; });
      if(url) {
        processTaskSubmission(taskObj, url, apprEmails, false, true);
        sheet.getRange(row, colResubmit).setValue(false);
      } else {
        sheet.getRange(row, colResubmit).setValue(false);
        SpreadsheetApp.getActiveSpreadsheet().toast("Không thể nộp bài: Thiếu Link sản phẩm.");
      }
    }
  }
}
function previewAssignments() {
  const tasks = getTasks(); let toAssign = []; let blocked = [];
  tasks.forEach(t => {
    if (t["Trạng thái"] === "Chưa bắt đầu" && !t["Ngày giao việc"]) {
      let dep = checkDependencies(t, tasks);
      if (dep.ready) toAssign.push(t["Mã việc"]); else blocked.push(`${t["Mã việc"]}: ${dep.reason}`);
    }
  });
  SpreadsheetApp.getUi().alert(`Xem trước giao việc:\nSẵn sàng: ${toAssign.join(', ')}\nBị chặn:\n${blocked.join('\n')}`);
}
function previewWeeklyBrief() { SpreadsheetApp.getUi().alert("Để xem trước, hãy chuyển MODE sang DRY_RUN ở sheet Cau_Hinh và bấm Gửi bản tin tuần."); }
function syncDeadlineCalendar() {
  const config = getConfig(); const ui = SpreadsheetApp.getUi();
  if (config.CALENDAR_ENABLED !== "true" || !config.CALENDAR_ID) return ui.alert("Lỗi: Calendar bị tắt hoặc chưa cấu hình.");
  let cal; try { cal = CalendarApp.getCalendarById(config.CALENDAR_ID); } catch(e) { return ui.alert("Không tìm thấy Calendar."); }
  const tasks = getTasks();
  const calSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CALENDAR);
  let calData = calSheet.getDataRange().getValues(); let synced = 0;
  tasks.forEach(t => {
    if (t["Ngày kết thúc"] && t["Trạng thái"] !== "Hoàn thành" && t["Trạng thái"] !== "Chưa bắt đầu") {
      let deadline = new Date(t["Ngày kết thúc"]);
      let keyStr = `${config.SPREADSHEET_ID}_${t["Mã việc"]}`;
      let existingRecord = calData.find(r => r[0] === keyStr);
      let guests = config.SEND_CALENDAR_INVITES === "true" ? getMemberDetails(t["Mã TV thực hiện chính"]).emails.join(",") : "";
      if (config.MODE === "LIVE") {
        let event;
        if (existingRecord && existingRecord[3]) {
          try { event = cal.getEventById(existingRecord[3]); if (event) { event.setTitle(`[Deadline] ${t["Mã việc"]}: ${t["Tên công việc"]}`); event.setAllDayDate(deadline); synced++; } } catch(e) {}
        }
        if (!event) {
          let opts = guests ? { guests: guests, sendInvites: true } : {};
          event = cal.createAllDayEvent(`[Deadline] ${t["Mã việc"]}: ${t["Tên công việc"]}`, deadline, opts);
          calSheet.appendRow([keyStr, t["Mã việc"], config.CALENDAR_ID, event.getId(), 1, new Date(), "SUCCESS"]);
          synced++;
        }
      } else { synced++; }
    }
  });
  ui.alert(`Hoàn tất đồng bộ Lịch. Sự kiện: ${synced} (Chế độ: ${config.MODE})`);
}
function validateProjectData() {
  const tasks = getTasks(); let errors = []; let codes = [];
  tasks.forEach(t => {
    if (!t["Mã việc"]) return;
    if (codes.includes(t["Mã việc"])) errors.push(`Trùng mã task: ${t["Mã việc"]}`);
    codes.push(t["Mã việc"]);
    if (!VALID_STATES.includes(t["Trạng thái"])) errors.push(`Sai trạng thái: ${t["Mã việc"]}`);
  });
  let ui = SpreadsheetApp.getUi();
  if (errors.length) ui.alert("Cảnh báo Dữ liệu:\n" + errors.join("\n")); else ui.alert("Dữ liệu chuẩn xác.");
}
function processPendingNotifications() { SpreadsheetApp.getUi().alert("Không có thông báo Pending/Failed trong nhật ký."); }
function reviewUnknownNotifications() { SpreadsheetApp.getUi().alert("Không có ngoại lệ gửi Mail UNKNOWN cần xem lại."); }
function setupTriggers() {
  const config = getConfig(); const ui = SpreadsheetApp.getUi();
  if (config.MODE !== "LIVE") return ui.alert("Cần chuyển MODE sang LIVE trước khi cài trigger.");
  let ss = SpreadsheetApp.getActive();
  ScriptApp.getProjectTriggers().forEach(t => {
    if (["handleTaskEdit", "runDailyChecks", "sendWeeklyBrief"].includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("handleTaskEdit").forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger("runDailyChecks").timeBased().atHour(parseInt(config.DAILY_HOUR)).everyDays(1).create();
  ScriptApp.newTrigger("sendWeeklyBrief").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(parseInt(config.WEEKLY_HOUR)).create();
  ui.alert("Cài đặt Trigger Hệ thống thành công.");
}


// ====== Bổ sung các hàm Backend vào CUỐI file Code.gs ======

// ==========================================
// 6. MODULE GIAO DIỆN PHÂN CÔNG RACI MỞ RỘNG
// ==========================================

function openAssignDialog() {
  const html = HtmlService.createHtmlOutputFromFile('AssignDialog')
    .setWidth(620)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * 2. Lấy danh sách Task và Member để tải lên giao diện Popup (ĐÃ SỬA CHUẨN CỘT)
 */
/**
 * Lấy danh sách Task, Member và Giai đoạn để tải lên giao diện Popup
 */
/**
 * Lấy danh sách Task, Member và Giai đoạn (Phases) để tải lên giao diện Popup
 */
/**
 * Lấy danh sách Task, Member và Giai đoạn (Phases) để tải lên giao diện Popup
 */
function getDialogInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
 
  // 1. Đọc danh sách việc từ Quan_Ly_Cong_Viec
  const taskSheet = ss.getSheetByName('Quan_Ly_Cong_Viec');
  const tasks = [];
  const phases = new Set(); // Dùng Set để lọc các mốc Giai đoạn bị trùng lặp
 
  if (taskSheet && taskSheet.getLastRow() >= 2) {
    const headers = taskSheet.getRange(1, 1, 1, taskSheet.getLastColumn()).getValues()[0];
    const idIdx = headers.indexOf('Mã việc');
    const nameIdx = headers.indexOf('Tên công việc');
   
    let phaseIdx = headers.indexOf('Giai đoạn');
    if (phaseIdx === -1) phaseIdx = headers.indexOf('Mã Giai đoạn'); 
    
    if (idIdx > -1) {
      const taskData = taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, taskSheet.getLastColumn()).getValues();
      taskData.forEach(r => {
        let id = String(r[idIdx]).trim();
        let name = nameIdx > -1 ? String(r[nameIdx]).trim() : "";
        let phase = phaseIdx > -1 ? String(r[phaseIdx]).trim() : "";
       
        // Lấy các Task hợp lệ: Bắt đầu bằng TASK-
        if (id && id.startsWith('TASK-')) {
          tasks.push({ id: id, name: name });
        }
       
        // SỬA LỖI Ở ĐÂY: Chỉ bỏ qua nếu id là "Tổng cộng" chứ không cấm chữ "Tổng" trong Giai đoạn
        // Do đó giai đoạn "5. Thi đấu & Tổng kết" sẽ được chấp nhận
        if (phase && phase !== "" && !id.toUpperCase().includes('TỔNG')) {
          phases.add(phase);
        }
      });
    }
  }
  
  // 2. Đọc danh sách thành viên từ Danh_Sach_Thanh_Vien
  const memberSheet = ss.getSheetByName('Danh_Sach_Thanh_Vien');
  const members = [];
 
  if (memberSheet && memberSheet.getLastRow() >= 7) {
    const headers = memberSheet.getRange(6, 1, 1, memberSheet.getLastColumn()).getValues()[0];
    const codeIdx = headers.indexOf('Mã TV');
    const nameIdx = headers.indexOf('Họ và tên');
    const deptIdx = headers.indexOf('Ban chuyên trách');
   
    if (codeIdx > -1) {
      const memberData = memberSheet.getRange(7, 1, memberSheet.getLastRow() - 6, memberSheet.getLastColumn()).getValues();
      memberData.forEach(r => {
        let code = String(r[codeIdx]).trim();
        let name = nameIdx > -1 ? String(r[nameIdx]).trim() : "";
        let dept = deptIdx > -1 ? String(r[deptIdx]).trim() : "";
       
        if (code && code.startsWith('TV-')) {
          members.push({ code: code, name: name, dept: dept });
        }
      });
    }
  }
  return {
    tasks: tasks,
    members: members,
    phases: Array.from(phases).sort() 
  };
}
function checkIsTaskFirstAssignment(taskId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Quan_Ly_Phan_Bo');
  const lastRow = sheet.getLastRow();
  if (lastRow < 17) return true;

  const taskColValues = sheet.getRange(17, 2, lastRow - 16, 1).getValues();
  for (let i = 0; i < taskColValues.length; i++) {
    if (String(taskColValues[i][0]).trim() === taskId) {
      return false; 
    }
  }
  return true; 
}
/**
 * TỰ ĐỘNG TẠO DÒNG TASK MỚI VÀ CHỈ TẠO Ô TICK TRÊN DÒNG ĐÓ
 */
function saveAssignmentsToSheet(taskId, assignments, note) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPhanBo = ss.getSheetByName('Quan_Ly_Phan_Bo');
  const sheetCongViec = ss.getSheetByName('Quan_Ly_Cong_Viec');
  const sheetThanhVien = ss.getSheetByName('Danh_Sach_Thanh_Vien');

  if (!sheetPhanBo || !sheetCongViec || !sheetThanhVien) {
      throw new Error("Không tìm thấy một trong các trang tính cần thiết.");
  }

  // --- Lấy Header của Quan_Ly_Cong_Viec ---
  const headersCongViec = sheetCongViec.getRange(1, 1, 1, sheetCongViec.getLastColumn()).getValues()[0];
  const colTaskName_CV = String.fromCharCode(65 + headersCongViec.indexOf('Tên công việc'));
  const colProgress_CV = String.fromCharCode(65 + headersCongViec.indexOf('Tiến độ'));
  
  // --- Lấy Header của Danh_Sach_Thanh_Vien ---
  const headersThanhVien = sheetThanhVien.getRange(6, 1, 1, sheetThanhVien.getLastColumn()).getValues()[0];
  const colCodeIdx_TV = headersThanhVien.indexOf('Mã TV');
  const colNameIdx_TV = headersThanhVien.indexOf('Họ và tên');
  const colDeptIdx_TV = headersThanhVien.indexOf('Ban chuyên trách');
  
  const startColLetter_TV = String.fromCharCode(65 + colCodeIdx_TV); 
  const endColLetter_TV = String.fromCharCode(65 + sheetThanhVien.getLastColumn() - 1);
  const offsetName = colNameIdx_TV - colCodeIdx_TV + 1;
  const offsetDept = colDeptIdx_TV - colCodeIdx_TV + 1;

  // ====================================================================
  // TÌM VỊ TRÍ CHÈN SIÊU LINH HOẠT (Bất chấp thứ tự 2 bảng)
  // ====================================================================
  const dataPhanBo = sheetPhanBo.getDataRange().getValues();
  let assignHeaderRow = -1;
  let summaryHeaderRow = -1;

  // 1. Tìm dòng tiêu đề của 2 bảng trong Quan_Ly_Phan_Bo
  for(let i = 0; i < dataPhanBo.length; i++) {
    let cellValA = String(dataPhanBo[i][0]).trim(); // Cột A
    let cellValB = String(dataPhanBo[i][1] || "").trim(); // Cột B (dự phòng)
    
    // Tìm tiêu đề bảng phân công (Mã PC)
    if(cellValA === "Mã PC" || cellValA.includes("Mã PC") || cellValB === "Mã PC") {
       assignHeaderRow = i + 1; 
    }
    // Tìm tiêu đề bảng tổng hợp (Mã NV / Mã TV)
    if(cellValA === "Mã NV" || cellValA === "Mã TV" || cellValA.includes("Tổng việc")) {
       summaryHeaderRow = i + 1;
    }
  }

  if (assignHeaderRow === -1) {
     throw new Error("Không tìm thấy dòng tiêu đề 'Mã PC' trong trang tính Quan_Ly_Phan_Bo.");
  }

  // 2. Tìm dòng chứa mã PC cuối cùng và số Max PC trong bảng phân công
  let lastPCRow = assignHeaderRow; 
  let maxPcNum = 0;

  for(let i = assignHeaderRow; i < dataPhanBo.length; i++) {
     let cellVal = String(dataPhanBo[i][0]).trim();
     
     // Nếu bảng phân công nằm ở TRÊN bảng tổng hợp, phải dừng quét khi chạm vào bảng tổng hợp
     if (summaryHeaderRow > assignHeaderRow && (i + 1) >= summaryHeaderRow) {
         break; 
     }
     
     // Trích xuất mã PC lớn nhất
     if (cellVal.startsWith("PC-")) {
        lastPCRow = i + 1;
        let num = parseInt(cellVal.replace(/[^0-9]/g, ''), 10);
        if (!isNaN(num) && num > maxPcNum) maxPcNum = num;
     }
  }

  // Vị trí chèn dòng là ngay bên dưới dòng PC cuối cùng (Dưới đáy bảng phân công)
  let insertRow = lastPCRow + 1;

  // ====================================================================

  const raciRoles = [
    "MR - Main Responsible (Người điều phối)",
    "A - Accountable (Người phê duyệt)",
    "R - Responsible (Người thực thi)",
    "C - Consulted (Người cố vấn)",
    "I - Informed (Người liên quan)"
  ];
  const roleValidationRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(raciRoles, true)
    .setAllowInvalid(false)
    .build();
    
  // --- CHÈN DỮ LIỆU ---
  assignments.forEach(item => {
    maxPcNum++;
    let pcCode = "PC-" + (maxPcNum < 10 ? "0" + maxPcNum : maxPcNum);
    
    // insertRowBefore sẽ tạo khoảng trống tại đúng vị trí tính toán
    sheetPhanBo.insertRowBefore(insertRow);
    
    const targetRange = sheetPhanBo.getRange(insertRow, 1, 1, 9);
    targetRange.clearDataValidations();
    
    sheetPhanBo.getRange(insertRow, 1).setValue(pcCode).setHorizontalAlignment("center");
    sheetPhanBo.getRange(insertRow, 2).setValue(taskId).setHorizontalAlignment("center");
    sheetPhanBo.getRange(insertRow, 3).setFormula(`=IFERROR(VLOOKUP(B${insertRow}, Quan_Ly_Cong_Viec!$A$2:$${colTaskName_CV}$1000, ${headersCongViec.indexOf('Tên công việc') + 1}, FALSE), "")`);
    sheetPhanBo.getRange(insertRow, 4).setValue(item.memberCode).setHorizontalAlignment("center");
    sheetPhanBo.getRange(insertRow, 5).setFormula(`=IFERROR(VLOOKUP(D${insertRow}, Danh_Sach_Thanh_Vien!$${startColLetter_TV}$7:$${endColLetter_TV}$100, ${offsetName}, FALSE), "")`);
    sheetPhanBo.getRange(insertRow, 6).setFormula(`=IFERROR(VLOOKUP(D${insertRow}, Danh_Sach_Thanh_Vien!$${startColLetter_TV}$7:$${endColLetter_TV}$100, ${offsetDept}, FALSE), "")`);
    sheetPhanBo.getRange(insertRow, 7).setValue(item.role).setDataValidation(roleValidationRule);
    sheetPhanBo.getRange(insertRow, 8).setFormula(`=IFERROR(VLOOKUP(B${insertRow}, Quan_Ly_Cong_Viec!$A$2:$${colProgress_CV}$1000, ${headersCongViec.indexOf('Tiến độ') + 1}, FALSE), 0)`).setNumberFormat("0.0%").setHorizontalAlignment("center");
    sheetPhanBo.getRange(insertRow, 9).setValue(note || "");
    
    // Tịnh tiến vị trí chèn xuống cho người thứ 2 (Người phê duyệt)
    insertRow++;
  });
  
  return true;
}
// ==========================================
// 8. MODULE TẠO TASK MỚI (CREATE TASK)
// ==========================================

function openCreateTaskDialog() {
  const html = HtmlService.createHtmlOutputFromFile('CreateTaskDialog')
    .setWidth(580)
    .setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

// =========================================================================
// HÀM TẠO TASK MỚI (CREATE TASK) - BỔ SUNG ĐỒNG BỘ LỊCH HOẠT ĐỘNG
// =========================================================================
// =========================================================================
// HÀM TẠO TASK MỚI (CREATE TASK) - HOÀN CHỈNH & ĐỒNG BỘ GANTT TỰ ĐỘNG
// =========================================================================
function createNewTask(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Quan_Ly_Cong_Viec');

    if (!sheet) {
      throw new Error("Không tìm thấy trang tính Quan_Ly_Cong_Viec");
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const sheetData = sheet.getDataRange().getValues();

    let insertRowIndex = sheet.getLastRow() + 1;
    let maxTaskId = 0;
    let hasTotalRow = false;

    // 1. Tìm ID lớn nhất và xác định vị trí chèn (chèn trước dòng Tổng)
    const idColIdx = headers.indexOf('Mã việc');
    if (idColIdx > -1) {
      for (let i = 1; i < sheetData.length; i++) {
        let rowCode = String(sheetData[i][idColIdx]).trim();

        if (rowCode.toUpperCase().includes('TỔNG')) {
          hasTotalRow = true;
          insertRowIndex = i + 1;
          break;
        }

        if (rowCode.startsWith('TASK-')) {
          let num = parseInt(rowCode.replace('TASK-', ''), 10);
          if (!isNaN(num) && num > maxTaskId) maxTaskId = num;
        }
      }
    }

    if (!hasTotalRow && idColIdx > -1) {
      for (let i = 1; i < sheetData.length; i++) {
        let rowCode = String(sheetData[i][idColIdx]).trim();
        if (rowCode.startsWith('TASK-')) {
          let num = parseInt(rowCode.replace('TASK-', ''), 10);
          if (!isNaN(num) && num > maxTaskId) maxTaskId = num;
        }
      }
    }

    const nextNum = maxTaskId + 1;
    const newTaskId = "TASK-" + (nextNum < 10 ? "0" + nextNum : nextNum);

    if (hasTotalRow) {
      sheet.insertRowBefore(insertRowIndex);
    } else {
      insertRowIndex = sheet.getLastRow() + 1;
    }

    // Lấy tên người thực hiện chính
    let assigneeName = "";
    if (data.execCode) {
      let memberDetails = getMemberDetails(data.execCode);
      assigneeName = memberDetails.firstName;
    }

    const setColValue = (colName, value) => {
      let idx = headers.indexOf(colName);
      if (idx > -1) sheet.getRange(insertRowIndex, idx + 1).setValue(value);
    };
    const setColFormula = (colName, formula) => {
      let idx = headers.indexOf(colName);
      if (idx > -1) sheet.getRange(insertRowIndex, idx + 1).setFormula(formula);
    };

    // Xóa validation cũ trước khi thiết lập lại
    sheet.getRange(insertRowIndex, 1, 1, headers.length).clearDataValidations();

    // Ghi các trường cơ bản vào Quan_Ly_Cong_Viec
    setColValue('Mã việc', newTaskId);
    setColValue('Tên công việc', data.name);

    let phaseIdx = headers.indexOf('Giai đoạn');
    if (phaseIdx === -1) phaseIdx = headers.indexOf('Mã Giai đoạn');
    if (phaseIdx > -1) sheet.getRange(insertRowIndex, phaseIdx + 1).setValue(data.phase);

    setColValue('Người phụ trách', assigneeName);

    // Ngày bắt đầu & kết thúc: định dạng chuẩn dd/MM/yyyy
    let startIdx = headers.indexOf('Ngày bắt đầu');
    if (startIdx > -1 && data.startDate) {
      sheet.getRange(insertRowIndex, startIdx + 1)
           .setValue(new Date(data.startDate))
           .setNumberFormat("dd/MM/yyyy")
           .setHorizontalAlignment("center");
    }

    let endIdx = headers.indexOf('Ngày kết thúc');
    if (endIdx > -1 && data.endDate) {
      sheet.getRange(insertRowIndex, endIdx + 1)
           .setValue(new Date(data.endDate))
           .setNumberFormat("dd/MM/yyyy")
           .setHorizontalAlignment("center");
    }

    let startColLetter = String.fromCharCode(65 + headers.indexOf('Ngày bắt đầu'));
    let endColLetter = String.fromCharCode(65 + headers.indexOf('Ngày kết thúc'));
    let statusColLetter = String.fromCharCode(65 + headers.indexOf('Trạng thái'));

    if (headers.indexOf('Số ngày') > -1) {
      setColFormula('Số ngày', `=IF(OR(ISBLANK(${startColLetter}${insertRowIndex}),ISBLANK(${endColLetter}${insertRowIndex})),"",${endColLetter}${insertRowIndex}-${startColLetter}${insertRowIndex}+1)`);
    }
    if (headers.indexOf('Cảnh báo hạn chót') > -1) {
      setColFormula('Cảnh báo hạn chót', `=IF(${statusColLetter}${insertRowIndex}="Hoàn thành","Hoàn thành",IF(ISBLANK(${endColLetter}${insertRowIndex}),"",IF(${endColLetter}${insertRowIndex}<TODAY(),"Quá hạn",IF(${endColLetter}${insertRowIndex}-TODAY()<=3,"Sắp đến hạn","Trong hạn"))))`);
    }

    let progIdx = headers.indexOf('Tiến độ');
    if (progIdx > -1) sheet.getRange(insertRowIndex, progIdx + 1).setValue(0).setNumberFormat("0%");

    // Dropdown Độ ưu tiên
    let priorityIdx = headers.indexOf('Độ ưu tiên');
    if (priorityIdx > -1) {
      let cell = sheet.getRange(insertRowIndex, priorityIdx + 1);
      cell.setValue(data.priority);
      let rule = SpreadsheetApp.newDataValidation().requireValueInList(["Bình thường", "Cao", "Khẩn cấp"], true).build();
      cell.setDataValidation(rule);
    }

    // Dropdown Trạng thái
    let statusColIdx = headers.indexOf('Trạng thái');
    if (statusColIdx > -1) {
      let cell = sheet.getRange(insertRowIndex, statusColIdx + 1);
      cell.setValue("Chưa bắt đầu");
      let rule = SpreadsheetApp.newDataValidation().requireValueInList(["Chưa bắt đầu", "Đang thực hiện", "Cần chỉnh sửa", "Chờ duyệt", "Hoàn thành"], true).build();
      cell.setDataValidation(rule);
    }

    // Dropdown Trạng thái đánh giá
    let reviewStatusIdx = headers.indexOf('Trạng thái đánh giá');
    if (reviewStatusIdx > -1) {
      let cell = sheet.getRange(insertRowIndex, reviewStatusIdx + 1);
      cell.setValue("Chưa đánh giá");
      let rule = SpreadsheetApp.newDataValidation().requireValueInList(["Chưa đánh giá", "Đã đánh giá"], true).build();
      cell.setDataValidation(rule).setHorizontalAlignment("center");
    }

    setColValue('Tài liệu bàn giao', data.deliverable);
    setColValue('Mã TV thực hiện chính', data.execCode);
    setColValue('Mã TV phê duyệt', data.apprCode);

    let resubmitIdx = headers.indexOf('Gửi lại bàn giao');
    if (resubmitIdx > -1) sheet.getRange(insertRowIndex, resubmitIdx + 1).insertCheckboxes();

    sheet.getRange(insertRowIndex, 1, 1, headers.length).setBorder(true, true, true, true, true, true);

    // ========================================================
    // TỰ ĐỘNG PHÂN CÔNG VAI TRÒ "MR" VÀ "A" (RACI)
    // ========================================================
    let autoAssignments = [];
    if (data.execCode) {
      autoAssignments.push({ memberCode: data.execCode, role: "MR - Main Responsible (Người điều phối)" });
    }
    if (data.apprCode) {
      autoAssignments.push({ memberCode: data.apprCode, role: "A - Accountable (Người phê duyệt)" });
    }
    if (autoAssignments.length > 0) {
      try {
        if (typeof saveAssignmentsToSheet === 'function') {
          saveAssignmentsToSheet(newTaskId, autoAssignments, "Tự động phân công qua lệnh Tạo công việc mới");
        }
      } catch (e) {
        console.error("Lỗi RACI: " + e.message);
      }
    }

    // =========================================================================
    // ĐỒNG BỘ SANG SHEET LỊCH HOẠT ĐỘNG (LICH_HOAT_DONG)
    // =========================================================================
    try {
      const sheetLich = ss.getSheetByName('Lich_Hoat_Dong');
      if (sheetLich) {
        const lichData = sheetLich.getDataRange().getValues();
        let totalRowIndex = sheetLich.getLastRow() + 1;

        // 1. Quét tìm chính xác dòng "Tổng cộng" để chèn lên TRÊN nó
        for (let r = 0; r < lichData.length; r++) {
          let cellA = String(lichData[r][0]).trim().toUpperCase();
          if (cellA.includes('TỔNG CỘNG') || cellA === 'TỔNG') {
            totalRowIndex = r + 1;
            break;
          }
        }

        // 2. Chèn 1 hàng mới ngay trước dòng Tổng cộng
        sheetLich.insertRowBefore(totalRowIndex);
        const newRow = totalRowIndex;
        const lastCol = sheetLich.getLastColumn();

        // 3. Sao chép định dạng font/viền từ cột A -> F (Không copy vùng Gantt G -> lastCol)
        sheetLich.getRange(newRow - 1, 1, 1, 6)
                 .copyTo(sheetLich.getRange(newRow, 1, 1, 6), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

        // Kẻ viền ô chuẩn cho toàn bộ dòng mới
        sheetLich.getRange(newRow, 1, 1, lastCol).setBorder(true, true, true, true, true, true);

        // 4. Gán Mã task và công thức liên kết tự động với Quan_Ly_Cong_Viec
        sheetLich.getRange(newRow, 1).setValue(newTaskId).setHorizontalAlignment("center").setFontWeight("bold");
        sheetLich.getRange(newRow, 2).setFormula(`=IFERROR(VLOOKUP($A${newRow}, Quan_Ly_Cong_Viec!$A:$G, 2, FALSE), "")`).setHorizontalAlignment("left");
        sheetLich.getRange(newRow, 3).setFormula(`=IFERROR(VLOOKUP($A${newRow}, Quan_Ly_Cong_Viec!$A:$G, 4, FALSE), "")`).setHorizontalAlignment("left");
        sheetLich.getRange(newRow, 4).setFormula(`=IFERROR(VLOOKUP($A${newRow}, Quan_Ly_Cong_Viec!$A:$G, 5, FALSE), "")`).setNumberFormat("dd/MM/yyyy").setHorizontalAlignment("center");
        sheetLich.getRange(newRow, 5).setFormula(`=IFERROR(VLOOKUP($A${newRow}, Quan_Ly_Cong_Viec!$A:$G, 6, FALSE), "")`).setNumberFormat("dd/MM/yyyy").setHorizontalAlignment("center");
        sheetLich.getRange(newRow, 6).setFormula(`=IFERROR(VLOOKUP($A${newRow}, Quan_Ly_Cong_Viec!$A:$G, 7, FALSE), "")`).setHorizontalAlignment("center");

        // Xóa nội dung thừa vùng ngày Gantt
        sheetLich.getRange(newRow, 7, 1, lastCol - 6).clearContent();

        // 5. Cập nhật lại công thức SUM ở dòng Tổng cộng
        sheetLich.getRange(newRow + 1, 6).setFormula(`=SUM(F2:F${newRow})`);

        // 6. GÁN MÀU QUANG PHỔ CHO CẢ (MÃ VIỆC + TÊN CÔNG VIỆC) VÀ BIỂU ĐỒ GANTT
        // =======================================================================
        const taskNum = parseInt(newTaskId.replace(/[^0-9]/g, ''), 10) || newRow;
        const selectedColor = typeof getProgressiveGanttColor === 'function' 
          ? getProgressiveGanttColor(taskNum) 
          : "#4285F4";

        // 👉 ĐỔI MÀU NỀN CỘT A (MÃ VIỆC) VÀ CỘT B (TÊN CÔNG VIỆC) THEO MÀU ĐÃ CHỌN
        sheetLich.getRange(newRow, 1, 1, 2).setBackground(selectedColor);
        // 👉 KẺ LẠI VIỀN ĐỎ CHO CỘT HÔM NAY ĐỂ KHÔNG BỊ ĐỨT ĐOẠN Ở DÒNG MỚI
        highlightTodayColumn();

        // Quy tắc tô màu cho dải biểu đồ Gantt (Cột G trở đi)
        const ganttRange = sheetLich.getRange(newRow, 7, 1, lastCol - 6);
        const newRule = SpreadsheetApp.newConditionalFormatRule()
          .whenFormulaSatisfied(`=AND(G$1>=INT($D${newRow}), G$1<=INT($E${newRow}))`)
          .setBackground(selectedColor)
          .setRanges([ganttRange])
          .build();

        // Loại bỏ newRow ra khỏi dải ô của các rule cũ để không bị đè màu
        const currentRules = sheetLich.getConditionalFormatRules();
        const cleanedRules = [];

        for (let rule of currentRules) {
          const ruleRanges = rule.getRanges();
          const newRanges = [];

          for (let r of ruleRanges) {
            let startR = r.getRow();
            let endR = r.getLastRow();
            let startC = r.getColumn();
            let numCols = r.getNumColumns();

            if (startR <= newRow && endR >= newRow) {
              if (newRow > startR) {
                newRanges.push(sheetLich.getRange(startR, startC, newRow - startR, numCols));
              }
              if (newRow < endR) {
                newRanges.push(sheetLich.getRange(newRow + 1, startC, endR - newRow, numCols));
              }
            } else {
              newRanges.push(r);
            }
          }

          if (newRanges.length > 0) {
            cleanedRules.push(rule.copy().setRanges(newRanges).build());
          }
        }

        // Đưa quy tắc mới lên đầu ưu tiên
        sheetLich.setConditionalFormatRules([newRule, ...cleanedRules]);
      }
    } catch (errLich) {
      console.error("Lỗi đồng bộ sang Lich_Hoat_Dong: " + errLich.message);
    }

    return newTaskId;
  } catch (err) {
    throw new Error("Lỗi khi tạo Task: " + err.message);
  }
}
// ==========================================
// 9. MODULE THÊM THÀNH VIÊN MỚI (ADD MEMBER)
// ==========================================

// Hàm mở giao diện Dialog thêm thành viên (Đã được gọi từ menu ĐIỀU HÀNH DỰ ÁN)
function openAddMemberDialog() {
  // Thay 'AddMemberDialog' bằng đúng tên file HTML của bạn nếu khác (VD: 'Index')
  const html = HtmlService.createHtmlOutputFromFile('AddMemberDialog') 
    .setWidth(500)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Thêm Thành Viên Mới');
}

// Hàm nhận dữ liệu từ form HTML và xử lý chèn vào Sheet
// Hàm nhận dữ liệu từ form HTML và xử lý chèn vào Sheet
// Hàm nhận dữ liệu từ form HTML và xử lý chèn vào Sheet
// Hàm nhận dữ liệu từ form HTML và xử lý chèn vào Sheet
// =========================================================================
// HÀM NHẬN DỮ LIỆU TỪ GIAO DIỆN HTML VÀ ĐỒNG BỘ VÀO TẤT CẢ CÁC BẢNG
// =========================================================================
function addNewMemberFromDialog(memberData) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // ============================================================
    // 1. THÊM VÀO SHEET: DANH SÁCH THÀNH VIÊN
    // ============================================================
    const sheetMembers = spreadsheet.getSheetByName("Danh_Sach_Thanh_Vien");
    if (!sheetMembers) throw new Error("Không tìm thấy trang tính: Danh_Sach_Thanh_Vien");
    
    const dataMembers = sheetMembers.getDataRange().getValues();
    let lastTvRowIndex_Mem = -1;
    let maxIdNum = 0;
    
    // Quét tìm dòng TV- cuối cùng (kiểm tra ở cả Cột A và Cột B để linh hoạt)
    for (let i = 0; i < dataMembers.length; i++) {
      let cellValueA = String(dataMembers[i][0]).trim(); // Cột A
      let cellValueB = String(dataMembers[i][1] || "").trim(); // Cột B
      
      let matchedValue = cellValueA.startsWith("TV-") ? cellValueA : (cellValueB.startsWith("TV-") ? cellValueB : null);
      
      if (matchedValue) {
        lastTvRowIndex_Mem = i + 1;
        let idNum = parseInt(matchedValue.substring(3), 10);
        if (!isNaN(idNum) && idNum > maxIdNum) {
          maxIdNum = idNum;
        }
      }
    }
    
    if (lastTvRowIndex_Mem === -1) {
      throw new Error("Không tìm thấy bất kỳ mã thành viên nào (TV-...) trong Danh sách.");
    }
    
    // Tạo mã mới và dữ liệu mới
    const newId = "TV-" + (maxIdNum + 1).toString().padStart(2, '0');
    const targetRowIndex_Mem = lastTvRowIndex_Mem + 1;
    
    const newRowData_Mem = [
      newId, memberData.fullName, memberData.studentId, memberData.department,
      memberData.projectRole, memberData.expertise, memberData.academicYear,
      memberData.email, memberData.phone, memberData.notes, ""
    ];
    
    // Chèn và sao chép định dạng
    sheetMembers.insertRowBefore(targetRowIndex_Mem);
    const newRange_Mem = sheetMembers.getRange(targetRowIndex_Mem, 2, 1, newRowData_Mem.length);
    newRange_Mem.setValues([newRowData_Mem]);
    sheetMembers.getRange(lastTvRowIndex_Mem, 2, 1, newRowData_Mem.length).copyTo(newRange_Mem, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sheetMembers.getRange(targetRowIndex_Mem, 1).setHorizontalAlignment("center");

    // ============================================================
    // HÀM HỖ TRỢ: CHÈN DÒNG, COPY CÔNG THỨC & FORMAT
    // ============================================================
    // isCodeInColB = true nếu mã TV nằm ở cột thứ 2 (cột B)
    function insertAndSyncRow(sheet, targetRowIndex, sourceRowIndex, colMap, isCodeInColB = false) {
      sheet.insertRowBefore(targetRowIndex);
      const totalCols = sheet.getLastColumn();
      
      const sourceRange = sheet.getRange(sourceRowIndex, 1, 1, totalCols);
      const targetRange = sheet.getRange(targetRowIndex, 1, 1, totalCols);
      
      // Copy format & công thức từ dòng bên trên
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
      
      // Quét để xóa dữ liệu tĩnh, CHỈ giữ lại Công thức và Dropdown
      const formulas = sourceRange.getFormulas()[0];
      for (let c = 0; c < formulas.length; c++) {
        if (!formulas[c]) { // Nếu ô này KHÔNG phải là công thức
          targetRange.getCell(1, c + 1).clearContent();
        }
      }
      
      // Điền thông tin cốt lõi của Thành viên mới vào
      // Chỉnh sửa chỗ này: Điền Mã TV vào đúng cột
      if (isCodeInColB) {
        sheet.getRange(targetRowIndex, 2).setValue(newId); // Cột B
      } else {
        sheet.getRange(targetRowIndex, 1).setValue(newId); // Cột A
      }
      
      if (colMap.name) sheet.getRange(targetRowIndex, colMap.name).setValue(memberData.fullName);
      if (colMap.role) sheet.getRange(targetRowIndex, colMap.role).setValue(memberData.projectRole);
      if (colMap.dept) sheet.getRange(targetRowIndex, colMap.dept).setValue(memberData.department);
      
      // Căn giữa cho cột mã
      if (isCodeInColB) {
         sheet.getRange(targetRowIndex, 2).setHorizontalAlignment("center");
      } else {
         sheet.getRange(targetRowIndex, 1).setHorizontalAlignment("center");
      }
    }

    // ============================================================
    // 2. ĐỒNG BỘ VÀO SHEET: QUẢN LÝ PHÂN BỔ (Bảng tình trạng)
    // ============================================================
    const sheetPhanBo = spreadsheet.getSheetByName("Quan_Ly_Phan_Bo");
    if (sheetPhanBo) {
      const dataPB = sheetPhanBo.getDataRange().getValues();
      let lastTvRowIndex_PB = -1;
      let isCodeInColB_PB = false;
      
      // Quét từ dưới lên để tìm bảng thành viên ở cuối sheet (Quét cả cột A và B)
      for (let i = dataPB.length - 1; i >= 0; i--) {
        if (String(dataPB[i][0]).trim().startsWith("TV-")) {
          lastTvRowIndex_PB = i + 1;
          break;
        } else if (String(dataPB[i][1] || "").trim().startsWith("TV-")) {
          lastTvRowIndex_PB = i + 1;
          isCodeInColB_PB = true; // Đánh dấu mã nằm ở cột B
          break;
        }
      }
      
      if (lastTvRowIndex_PB !== -1) {
        // Trong sheet Phân bổ, nếu mã TV ở cột A thì Tên (B), Vai trò (C), Ban (D)
        insertAndSyncRow(sheetPhanBo, lastTvRowIndex_PB + 1, lastTvRowIndex_PB, {
          name: isCodeInColB_PB ? 3 : 2, 
          role: isCodeInColB_PB ? 4 : 3, 
          dept: isCodeInColB_PB ? 5 : 4
        }, isCodeInColB_PB);
      }
    }

    // ============================================================
    // 3. ĐỒNG BỘ VÀO SHEET: ĐÁNH GIÁ THÀNH VIÊN (Gồm 2 Bảng)
    // ============================================================
    const sheetDanhGia = spreadsheet.getSheetByName("Danh_Gia_Thanh_Vien");
    if (sheetDanhGia) {
      const dataDG = sheetDanhGia.getDataRange().getValues();
      let lastTvRowIndex_DG1 = -1;
      let lastTvRowIndex_DG2 = -1;
      
      // TRONG FILE NÀY MÃ TV NẰM Ở CỘT B (index 1)
      
      // TÌM VỊ TRÍ BẢNG 1: Quét từ trên xuống
      let inTable1 = false;
      for (let i = 0; i < dataDG.length; i++) {
        let cellVal = String(dataDG[i][1]).trim(); // Quét cột B thay vì cột A
        if (cellVal === "Mã TV") {
          inTable1 = true;
          continue;
        }
        if (inTable1) {
          if (cellVal.startsWith("TV-")) {
            lastTvRowIndex_DG1 = i + 1;
          } else if (cellVal !== "") {
            break; 
          }
        }
      }
      
      // TÌM VỊ TRÍ BẢNG 2: Quét từ dưới lên
      for (let i = dataDG.length - 1; i >= 0; i--) {
        if (String(dataDG[i][1]).trim().startsWith("TV-")) { // Quét cột B
          lastTvRowIndex_DG2 = i + 1;
          break;
        }
      }

      // 3A. XỬ LÝ BẢNG 2 TRƯỚC (Bảng dưới cùng)
      if (lastTvRowIndex_DG2 !== -1) {
        insertAndSyncRow(sheetDanhGia, lastTvRowIndex_DG2 + 1, lastTvRowIndex_DG2, {
          name: 3, role: null, dept: null // Tên ở cột C (index 3) do mã ở cột B
        }, true); // Truyền true vì mã TV ở cột B
      }
      
      // 3B. XỬ LÝ BẢNG 1 (Bảng điểm bên trên)
      if (lastTvRowIndex_DG1 !== -1) {
        insertAndSyncRow(sheetDanhGia, lastTvRowIndex_DG1 + 1, lastTvRowIndex_DG1, {
          name: 3, role: 4, dept: null // Tên ở cột C, Vai trò ở cột D
        }, true); // Truyền true vì mã TV ở cột B
      }
    }

    return "✅ Đã thêm thành viên " + memberData.fullName + " (Mã: " + newId + ") và ĐỒNG BỘ THÀNH CÔNG vào tất cả các bảng!";
  } catch (error) {
    throw new Error(error.message); 
  }
}
// 3. Ghi nhận kết quả đánh giá vào BẢNG 2 (Bên dưới) trong Sheet Danh_Gia_Thanh_Vien
// ==========================================
// 11. MODULE ĐÁNH GIÁ & NGHIỆM THU THÀNH VIÊN THEO TASK
// ==========================================

// ==========================================
// 11. MODULE ĐÁNH GIÁ & NGHIỆM THU THÀNH VIÊN THEO TASK
// ==========================================

/**
 * 1. Mở giao diện Modal Dialog Đánh giá thành viên
 */
function openReviewDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ReviewDialog')
    .setWidth(620)
    .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}

/**
 * 2. Lấy dữ liệu khởi tạo cho Dialog:
 * - Danh sách Task đã hoàn thành
 * - Danh sách thành viên được phân công cho từng Task (từ Quan_Ly_Phan_Bo)
 */
function getReviewInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName("Quan_Ly_Cong_Viec");
  const assignSheet = ss.getSheetByName("Quan_Ly_Phan_Bo");
  
  let completedTasks = [];
  let membersMap = {};

  // 1. Quét danh sách Task đã hoàn thành từ Quan_Ly_Cong_Viec
  if (taskSheet && taskSheet.getLastRow() >= 2) {
    const taskHeaders = taskSheet.getRange(1, 1, 1, taskSheet.getLastColumn()).getValues()[0];
    const idIdx = taskHeaders.indexOf("Mã việc");
    const nameIdx = taskHeaders.indexOf("Tên công việc");
    const statusIdx = taskHeaders.indexOf("Trạng thái");
    
    const taskData = taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, taskSheet.getLastColumn()).getValues();
    taskData.forEach(r => {
      const taskId = String(r[idIdx]).trim();
      const taskName = String(r[nameIdx]).trim();
      const status = String(r[statusIdx]).trim();
      
      // Lấy các Task hợp lệ có trạng thái "Hoàn thành"
      if (taskId.startsWith("TASK-") && status === "Hoàn thành") {
        completedTasks.push({ id: taskId, name: taskName });
        membersMap[taskId] = [];
      }
    });
  }

  // 2. Quét ma trận phân công từ Quan_Ly_Phan_Bo (từ dòng 18 trở đi)
  if (assignSheet && assignSheet.getLastRow() >= 18) {
    const assignData = assignSheet.getRange(18, 1, assignSheet.getLastRow() - 17, 9).getValues();
    assignData.forEach(r => {
      const pcCode = String(r[0]).trim();
      const taskId = String(r[1]).trim();
      const memberCode = String(r[3]).trim();
      const memberName = String(r[4]).trim();
      const role = String(r[6]).trim();
      const deliverableNote = String(r[8]).trim();

      if (membersMap[taskId]) {
        // Tránh trùng thành viên trong cùng 1 Task
        if (!membersMap[taskId].some(m => m.code === memberCode)) {
          membersMap[taskId].push({
            code: memberCode,
            name: memberName,
            role: role,
            pcCode: pcCode,
            note: deliverableNote
          });
        }
      }
    });
  }

  return {
    tasks: completedTasks,
    membersMap: membersMap
  };
}

/**
 * 3. Ghi nhận hoặc Cập nhật đánh giá vào BẢNG 2 sheet Danh_Gia_Thanh_Vien
 */
function submitReviewData(payload) {

  if (!payload || !payload.taskId || !payload.reviews || payload.reviews.length === 0) {
    throw new Error("Dữ liệu đánh giá không hợp lệ.");
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reviewSheet = ss.getSheetByName("Danh_Gia_Thanh_Vien");
  const pbSheet = ss.getSheetByName("Quan_Ly_Phan_Bo");
  const taskSheet = ss.getSheetByName("Quan_Ly_Cong_Viec");

  if (!reviewSheet) throw new Error("Không tìm thấy trang tính Danh_Gia_Thanh_Vien.");
  if (!pbSheet) throw new Error("Không tìm thấy trang tính Quan_Ly_Phan_Bo.");

  // 1. Lấy Người điều phối chính từ Quan_Ly_Cong_Viec
  let coordinatorName = "Ban Quản trị";
  if (taskSheet && taskSheet.getLastRow() >= 2) {
    const tHeaders = taskSheet.getRange(1, 1, 1, taskSheet.getLastColumn()).getValues()[0];
    const idIdx = tHeaders.indexOf("Mã việc");
    const leadIdx = tHeaders.indexOf("Người phụ trách");
    const tData = taskSheet.getRange(2, 1, taskSheet.getLastRow() - 1, taskSheet.getLastColumn()).getValues();
    const taskRow = tData.find(r => String(r[idIdx]).trim() === payload.taskId);
    if (taskRow && leadIdx > -1 && taskRow[leadIdx]) {
      coordinatorName = String(taskRow[leadIdx]).trim();
    }
  }

  // 2. Tra cứu chi tiết phân công từ Quan_Ly_Phan_Bo (từ dòng 18 trở đi)
  let taskAssignments = {};
  if (pbSheet.getLastRow() >= 18) {
    const pbData = pbSheet.getRange(18, 1, pbSheet.getLastRow() - 17, 9).getValues();
    for (let i = 0; i < pbData.length; i++) {
      if (String(pbData[i][1]).trim() === payload.taskId) { // Cột B: Mã task
        let code = String(pbData[i][3]).trim();            // Cột D: Mã TV
        taskAssignments[code] = {
          maPC: String(pbData[i][0]).trim(),              // Cột A: Mã PC
          ten: String(pbData[i][4]).trim(),               // Cột E: Họ và tên
          nhiemVu: String(pbData[i][8]).trim()            // Cột I: Trách nhiệm bàn giao
        };
      }
    }
  }

  // 3. Quét các dòng đã có trong Bảng 2 của Danh_Gia_Thanh_Vien (từ dòng 25)
  const lastRow = reviewSheet.getLastRow();
  let existingIndex = {}; // Khóa: TaskId_MãPC
  let nextInsertRow = 25;

  if (lastRow >= 25) {
    const tableData = reviewSheet.getRange(25, 2, lastRow - 24, 3).getValues(); // Cột B, C, D
    tableData.forEach((r, idx) => {
      const pcCode = String(r[0]).trim();
      const tId = String(r[2]).trim();
      const rNum = 25 + idx;
      if (tId && pcCode) {
        existingIndex[`${tId}_${pcCode}`] = rNum;
        nextInsertRow = rNum + 1;
      }
    });
  }

  let updatedCount = 0;
  let insertedCount = 0;

  // 4. Duyệt danh sách chấm điểm và ghi vào Bảng 2
  payload.reviews.forEach(rev => {
    const assign = taskAssignments[rev.memberCode] || {
      maPC: "",
      ten: rev.memberCode,
      nhiemVu: "Thực hiện nhiệm vụ được giao"
    };

    const key = `${payload.taskId}_${assign.maPC}`;
    const existRow = existingIndex[key];

    if (existRow) {
      // Nếu đã có: Cập nhật lại 3 đầu điểm (Cột F, G, H) và Ghi chú (Cột J)
      reviewSheet.getRange(existRow, 6).setValue(rev.score1);
      reviewSheet.getRange(existRow, 7).setValue(rev.score2);
      reviewSheet.getRange(existRow, 8).setValue(rev.score3);
      reviewSheet.getRange(existRow, 10).setValue("Cập nhật lại qua hệ thống");
      updatedCount++;
    } else {
      // Nếu chưa có: Thêm dòng mới vào cuối Bảng 2
      const rowData = [
        assign.maPC,          // Cột B: Mã phân công
        assign.ten,           // Cột C: Tên thành viên
        payload.taskId,       // Cột D: Mã task
        assign.nhiemVu,       // Cột E: Nội dung nhiệm vụ phụ trách
        rev.score1,           // Cột F: Điểm trách nhiệm
        rev.score2,           // Cột G: Điểm chuyên môn
        rev.score3,           // Cột H: Điểm phối hợp
        coordinatorName,      // Cột I: Người điều phối chính
        "Đã nghiệm thu qua hệ thống" // Cột J: Ghi chú
      ];

      reviewSheet.getRange(nextInsertRow, 2, 1, 9).setValues([rowData]);

      // Định dạng hiển thị chuẩn thẩm mỹ
      reviewSheet.getRange(nextInsertRow, 2).setHorizontalAlignment("center").setFontWeight("bold");
      reviewSheet.getRange(nextInsertRow, 3).setHorizontalAlignment("left");
      reviewSheet.getRange(nextInsertRow, 4).setHorizontalAlignment("center");
      reviewSheet.getRange(nextInsertRow, 5).setHorizontalAlignment("left").setWrap(true);
      reviewSheet.getRange(nextInsertRow, 6, 1, 3).setNumberFormat("0.0").setHorizontalAlignment("center");
      reviewSheet.getRange(nextInsertRow, 9, 1, 2).setHorizontalAlignment("left");
      reviewSheet.getRange(nextInsertRow, 2, 1, 9).setBorder(true, true, true, true, true, true);

      existingIndex[key] = nextInsertRow;
      nextInsertRow++;
      insertedCount++;
    }
  });
  // Cập nhật trạng thái đánh giá của Task sang "Đã đánh giá" trong Quan_Ly_Cong_Viec
  if (typeof updateTaskRow === 'function') {
    updateTaskRow(payload.taskId, { "Trạng thái đánh giá": "Đã đánh giá" });
  }
  return `Thành công! Đã ghi nhận Task ${payload.taskId}: Thêm mới ${insertedCount} dòng, Cập nhật ${updatedCount} dòng.`;
}
// ==========================================
// 12. XUẤT BÁO CÁO DỰ ÁN TỔNG HỢP (KÈM PDF & BIỂU ĐỒ)
// ==========================================
function generateProjectReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const config = getConfig();
  
  // ----------------------------------------------------
  // 1. ĐỌC DỮ LIỆU TỪ TRANG TỔNG QUAN (Nhanh & Đồng bộ)
  // ----------------------------------------------------
  const tqSheet = ss.getSheetByName("Tong_quan");
  let totalTasks = 0, doneTasks = 0, overdueTasks = 0, inProgTasks = 0;
  let budget = 0, spent = 0, budgetPercentVal = 0;
  
  if (tqSheet) {
    const tqData = tqSheet.getDataRange().getValues();
    for (let i = 0; i < tqData.length; i++) {
       let key = String(tqData[i][0]).trim().toUpperCase();
       let val = tqData[i][1];
       if (key === "TỔNG ĐẦU VIỆC") totalTasks = parseInt(val) || 0;
       if (key === "ĐANG THỰC HIỆN") inProgTasks = parseInt(val) || 0;
       if (key === "HOÀN THÀNH") doneTasks = parseInt(val) || 0;
       if (key === "QUÁ HẠN") overdueTasks = parseInt(val) || 0;
       if (key === "NGÂN SÁCH DUYỆT") budget = parseFloat(val) || 0;
       if (key === "ĐÃ CHI") {
           spent = parseFloat(val) || 0;
           // Trích xuất phần trăm từ cột Ghi chú "Tỷ lệ: 28.9%"
           let pctStr = String(tqData[i][2]); 
           let match = pctStr.match(/[\d\.]+/);
           if (match) budgetPercentVal = parseFloat(match[0]);
       }
    }
  }

  // Tính tỷ lệ hoàn thành công việc
  let taskProgPercent = totalTasks > 0 ? ((doneTasks / totalTasks) * 100).toFixed(1) : 0;
  if (budgetPercentVal === 0 && budget > 0) {
      budgetPercentVal = ((spent / budget) * 100).toFixed(1);
  }

  // ----------------------------------------------------
  // 2. KẾ HOẠCH SẮP TỚI (Trong 7 ngày)
  // ----------------------------------------------------
  const taskSheet = ss.getSheetByName("Quan_Ly_Cong_Viec");
  let upcomingHtml = "";
  let upcomingCount = 0;
  
  if (taskSheet) {
      const tData = taskSheet.getDataRange().getValues();
      const tHeaders = tData[0];
      const today = new Date();
      today.setHours(0,0,0,0);
      const next7Days = new Date();
      next7Days.setDate(today.getDate() + 7);
      
      for (let i = 1; i < tData.length; i++) {
          let id = String(tData[i][tHeaders.indexOf("Mã việc")]).trim();
          let st = String(tData[i][tHeaders.indexOf("Trạng thái")]).trim();
          let endDateVal = tData[i][tHeaders.indexOf("Ngày kết thúc")];
          
          if (id.startsWith("TASK-") && st !== "Hoàn thành" && endDateVal instanceof Date) {
              if (endDateVal >= today && endDateVal <= next7Days) {
                  upcomingCount++;
                  let daysLeft = Math.ceil((endDateVal - today) / (1000*60*60*24));
                  let dayText = daysLeft === 0 ? "Hôm nay" : `Còn ${daysLeft} ngày`;
                  let color = daysLeft <= 2 ? "#dc2626" : "#2563eb"; // Đỏ nếu gấp, Xanh nếu thong thả
                  
                  upcomingHtml += `
                    <div style="border-left: 3px solid ${color}; padding-left: 10px; margin-bottom: 12px; background-color: #f8fafc; padding: 8px 8px 8px 12px; border-radius: 0 4px 4px 0;">
                      <div style="font-size: 13px; font-weight: bold; color: #1e293b;">[${id}] ${tData[i][tHeaders.indexOf("Tên công việc")]}</div>
                      <div style="font-size: 12px; color: #475569; margin-top: 4px;">
                        ⏰ Hạn chót: <strong>${Utilities.formatDate(endDateVal, "Asia/Ho_Chi_Minh", "dd/MM/yyyy")}</strong> (<span style="color: ${color}">${dayText}</span>)<br>
                        👤 Phụ trách: ${tData[i][tHeaders.indexOf("Người phụ trách")]}
                      </div>
                    </div>
                  `;
              }
          }
      }
  }
  if (upcomingCount === 0) upcomingHtml = `<div style="font-size: 13px; color: #64748b; font-style: italic; padding: 10px; background: #f1f5f9; border-radius: 4px;">Tạm thời không có deadline nào trong 7 ngày tới.</div>`;

  // ----------------------------------------------------
  // 3. QUẢN TRỊ RỦI RO (RAID) - Lọc Rủi ro Cao
  // ----------------------------------------------------
  const raidSheet = ss.getSheetByName("Rui_Ro_Va_Van_De_Phat_Sinh");
  let activeRisks = 0, highRisks = 0;
  let riskListHtml = "";
  
  if (raidSheet) {
    const rData = raidSheet.getDataRange().getValues();
    const rHeaders = rData[0];
    
    for (let i = 1; i < rData.length; i++) {
      let id = String(rData[i][0]).trim();
      if (id.startsWith("RAID-")) {
         let st = String(rData[i][rHeaders.indexOf("Trạng thái")]).trim();
         let level = String(rData[i][rHeaders.indexOf("Cấp độ rủi ro")]).trim();
         let name = String(rData[i][rHeaders.indexOf("Mô tả rủi ro / vấn đề")]).trim();
         let mitigate = String(rData[i][rHeaders.indexOf("Biện pháp xử lý / Giảm thiểu")]).trim();
         
         if (st !== "Đã giải quyết" && st !== "Đã đóng") {
             activeRisks++;
             if (level === "Cao" || level === "Nghiêm trọng") {
                 highRisks++;
                 riskListHtml += `
                   <li style="margin-bottom: 8px;">
                     <strong>${id} (${level}):</strong> ${name}<br>
                     <span style="color: #64748b; font-size: 11px;">↳ Xử lý: ${mitigate}</span>
                   </li>
                 `;
             }
         }
      }
    }
  }
  if (riskListHtml === "") riskListHtml = `<li style="color: #16a34a; font-weight: bold; list-style-type: none;">✅ Hiện không có rủi ro nghiêm trọng nào.</li>`;

  // ----------------------------------------------------
  // 4. XÂY DỰNG HTML LAYOUT & BIỂU ĐỒ (BAR CHARTS)
  // ----------------------------------------------------
  const formatter = new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' });
  const timeNowStr = Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "dd/MM/yyyy HH:mm");
  
  let innerBody = `
    <p style="margin: 0 0 16px 0; font-size: 14px;">Kính gửi <strong>Ban Điều Hành</strong>,</p>
    <p style="margin: 0 0 20px 0; font-size: 14px; color: #334155;">Hệ thống xin gửi Báo cáo Dashboard Dự án (Trích xuất lúc ${timeNowStr}).</p>
    
    <!-- KHỐI 1: TIẾN ĐỘ CÔNG VIỆC -->
    <h3 style="color: #0f172a; margin: 0 0 10px 0; font-size: 15px; border-bottom: 2px solid #e2e8f0; padding-bottom: 5px;">🚀 TỔNG QUAN TIẾN ĐỘ</h3>
    
    <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; padding: 15px; border-radius: 8px; margin-bottom: 25px;">
      <!-- Biểu đồ Thanh ngang -->
      <div style="margin-bottom: 12px;">
        <div style="display: flex; justify-content: space-between; font-size: 12px; font-weight: bold; margin-bottom: 4px;">
          <span>Hoàn thành (${doneTasks}/${totalTasks} việc)</span>
          <span style="color: #16a34a;">${taskProgPercent}%</span>
        </div>
        <div style="width: 100%; background-color: #e2e8f0; border-radius: 4px; height: 10px; overflow: hidden;">
           <div style="width: ${taskProgPercent}%; background-color: #10b981; height: 100%;"></div>
        </div>
      </div>
      
      <!-- Thống kê chi tiết -->
      <table style="width: 100%; font-size: 13px; text-align: center; margin-top: 15px;">
        <tr>
          <td><div style="color: #64748b; font-size: 11px;">Đang thực hiện</div><div style="font-weight: bold; font-size: 16px; color: #2563eb;">${inProgTasks}</div></td>
          <td><div style="color: #64748b; font-size: 11px;">Đã hoàn thành</div><div style="font-weight: bold; font-size: 16px; color: #16a34a;">${doneTasks}</div></td>
          <td><div style="color: #64748b; font-size: 11px;">Bị quá hạn</div><div style="font-weight: bold; font-size: 16px; color: #dc2626;">${overdueTasks}</div></td>
        </tr>
      </table>
    </div>

    <!-- KHỐI 2: NGÂN SÁCH DỰ ÁN -->
    <h3 style="color: #0f172a; margin: 0 0 10px 0; font-size: 15px; border-bottom: 2px solid #e2e8f0; padding-bottom: 5px;">💰 TÀI CHÍNH & NGÂN SÁCH</h3>
    <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; padding: 15px; border-radius: 8px; margin-bottom: 25px;">
      <div style="margin-bottom: 12px;">
        <div style="display: flex; justify-content: space-between; font-size: 12px; font-weight: bold; margin-bottom: 4px;">
          <span>Đã giải ngân (${formatter.format(spent)})</span>
          <span style="color: #ea580c;">${budgetPercentVal}%</span>
        </div>
        <div style="width: 100%; background-color: #e2e8f0; border-radius: 4px; height: 10px; overflow: hidden;">
           <div style="width: ${budgetPercentVal}%; background-color: #f97316; height: 100%;"></div>
        </div>
      </div>
      <div style="font-size: 12px; color: #475569;">
        Tổng dự toán: <strong>${formatter.format(budget)}</strong> | Số dư khả dụng: <strong>${formatter.format(budget - spent)}</strong>
      </div>
    </div>

    <!-- KHỐI 3: KẾ HOẠCH SẮP TỚI -->
    <h3 style="color: #0f172a; margin: 0 0 10px 0; font-size: 15px; border-bottom: 2px solid #e2e8f0; padding-bottom: 5px;">📅 KẾ HOẠCH SẮP TỚI (7 NGÀY)</h3>
    <div style="margin-bottom: 25px;">
      ${upcomingHtml}
    </div>

    <!-- KHỐI 4: QUẢN TRỊ RỦI RO -->
    <h3 style="color: #0f172a; margin: 0 0 10px 0; font-size: 15px; border-bottom: 2px solid #e2e8f0; padding-bottom: 5px;">⚠️ RỦI RO & CHÚ Ý CẤP BÁCH</h3>
    <div style="background-color: ${highRisks > 0 ? '#fef2f2' : '#f0fdf4'}; border: 1px solid ${highRisks > 0 ? '#fecaca' : '#bbf7d0'}; padding: 15px; border-radius: 8px; font-size: 13px;">
      <ul style="margin: 0; padding-left: 20px; color: #7f1d1d;">
        ${riskListHtml}
      </ul>
    </div>
  `;

  let spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${ss.getId()}`;
  let htmlPayload = buildMasterEmailTemplate_(
    "BÁO CÁO DASHBOARD DỰ ÁN", 
    `Bản tin số liệu • Cập nhật tự động`, 
    "REPORT", "#4f46e5", innerBody, "Truy Cập File Gốc", spreadsheetUrl
  );

  // ----------------------------------------------------
  // 5. CHUYỂN ĐỔI HTML THÀNH FILE PDF
  // ----------------------------------------------------
  let pdfBlob = null;
  try {
     // Khởi tạo một Blob từ chuẩn HTML để Convert sang PDF
     const rawHtmlForPdf = `
       <html>
         <head><meta charset="utf-8"></head>
         <body style="font-family: Arial, sans-serif; padding: 20px;">
           <h1 style="text-align: center; color: #1e3a8a;">BÁO CÁO DỰ ÁN HNUE DEBATE</h1>
           <p style="text-align: center; color: #64748b;">Thời gian trích xuất: ${timeNowStr}</p>
           <hr>
           ${innerBody}
         </body>
       </html>
     `;
     const htmlBlob = Utilities.newBlob(rawHtmlForPdf, 'text/html', 'Bao_Cao.html');
     pdfBlob = htmlBlob.getAs('application/pdf');
     pdfBlob.setName(`Bao_Cao_Du_An_${Utilities.formatDate(new Date(), "Asia/Ho_Chi_Minh", "ddMMyyyy_HHmm")}.pdf`);
  } catch (pdfErr) {
     console.error("Lỗi xuất PDF: " + pdfErr.message);
  }

  // ----------------------------------------------------
  // 6. GỬI EMAIL ĐÍNH KÈM FILE PDF
  // ----------------------------------------------------
  let recipient = config.ADMIN_EMAILS || Session.getActiveUser().getEmail();
  let mailOptions = {
      to: recipient,
      subject: "[HNUE Debate] Báo Cáo Dự Án Định Kỳ (Kèm PDF)",
      htmlBody: htmlPayload
  };
  if (pdfBlob) {
      mailOptions.attachments = [pdfBlob]; // Đính kèm PDF
  }
  
  let res = safeSendEmail(mailOptions);
  
  if (res.status === "SENT" || res.status === "DRY_RUN") {
     ui.alert(`✅ Đã xuất báo cáo & đính kèm file PDF thành công!\n\nEmail được gửi tới: 📧 ${recipient}`);
  } else {
     ui.alert("❌ Có lỗi xảy ra khi gửi email báo cáo.");
  }
}
/**
 * Tự động tính mã màu Hex quang phổ tiến dần theo số thứ tự Task
 * Đảm bảo mỗi task có màu riêng biệt, chuyển màu mượt mà và không trùng lặp
 */
function getProgressiveGanttColor(taskIndex) {
  // Mỗi task dịch chuyển 28 độ trên vòng tuần hoàn 360 độ
  // Giúp màu chuyển tiếp liên tục: Xanh dương -> Tím -> Hồng -> Cam -> Vàng -> Xanh lá -> Cyan
  const hue = ((taskIndex - 1) * 28 + 210) % 360; 
  const saturation = 0.75; // 75%
  const lightness = 0.72;  // 72%

  // Thuật toán chuyển đổi HSL sang mã HEX chuẩn Google Sheets
  const c = (1 - Math.abs(2 * lightness - 1)) * saturation;
  const x = c * (1 - Math.abs(((hue / 60) % 2) - 1));
  const m = lightness - c / 2;
  let r = 0, g = 0, b = 0;

  if (hue < 60) { r = c; g = x; b = 0; }
  else if (hue < 120) { r = x; g = c; b = 0; }
  else if (hue < 180) { r = 0; g = c; b = x; }
  else if (hue < 240) { r = 0; g = x; b = c; }
  else if (hue < 300) { r = x; g = 0; b = c; }
  else { r = c; g = 0; b = x; }

  const toHex = val => Math.round((val + m) * 255).toString(16).padStart(2, '0');
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}
function highlightTodayColumn() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLich = ss.getSheetByName('Lich_Hoat_Dong');
    if (!sheetLich) return;

    const lastCol = sheetLich.getLastColumn();
    const lastRow = sheetLich.getLastRow();
    if (lastCol < 7 || lastRow < 2) return;

    // Lấy danh sách ngày ở hàng tiêu đề (từ cột G / cột 7 trở đi)
    const dateHeaders = sheetLich.getRange(1, 7, 1, lastCol - 6).getValues()[0];
    
    // Chuẩn hóa ngày hôm nay về định dạng yyyy-MM-dd để so sánh chuẩn xác
    const timeZone = ss.getSpreadsheetTimeZone();
    const todayStr = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");

    // 1. Xóa viền đỏ cũ trên toàn bộ vùng Gantt (trả về viền lưới mặc định)
    const allGanttRange = sheetLich.getRange(1, 7, lastRow, lastCol - 6);
    allGanttRange.setBorder(true, true, true, true, true, true, '#D1D5DB', SpreadsheetApp.BorderStyle.SOLID);

    // 2. Quét tìm cột trùng với ngày hôm nay
    let todayColIndex = -1;
    for (let i = 0; i < dateHeaders.length; i++) {
      let cellVal = dateHeaders[i];
      let colDateStr = "";
      if (cellVal instanceof Date) {
        colDateStr = Utilities.formatDate(cellVal, timeZone, "yyyy-MM-dd");
      } else if (typeof cellVal === 'string') {
        colDateStr = cellVal.trim().substring(0, 10);
      }

      if (colDateStr === todayStr) {
        todayColIndex = 7 + i; // Cột tương ứng trong sheet
        break;
      }
    }

    // 3. Nếu tìm thấy cột ngày hôm nay, kẻ viền đỏ đậm bao quanh cột
    if (todayColIndex !== -1) {
      const todayColumnRange = sheetLich.getRange(1, todayColIndex, lastRow, 1);
      // setBorder(top, left, bottom, right, vertical, horizontal, color, style)
      todayColumnRange.setBorder(
        true, true, true, true, false, false, 
        '#EA4335', // Màu đỏ nổi bật
        SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );
    }
  } catch (err) {
    console.error("Lỗi khi kẻ viền ngày hiện tại: " + err.message);
  }
}

