/**
 * ============================================================================
 * MODULE: QUẢN LÝ VÀ ĐIỀU CHỈNH DEADLINE TASK DỰ ÁN TRANH BIỆN
 * ============================================================================
 */

// Định nghĩa vị trí cột chuẩn xác trong sheet Quan_Ly_Cong_Viec
const TASK_COL_INDEX = {
  TASK_ID: 1,       // Cột A: Mã việc
  TASK_NAME: 2,     // Cột B: Tên công việc
  PHASE: 3,         // Cột C: Giai đoạn
  ASSIGNEE: 4,      // Cột D: Người phụ trách
  START_DATE: 5,    // Cột E: Ngày bắt đầu
  END_DATE: 6,      // Cột F: Ngày kết thúc (Deadline)
  DURATION: 7,      // Cột G: Số ngày
  PROGRESS: 8,      // Cột H: Tiến độ
  PRIORITY: 9,      // Cột I: Độ ưu tiên
  STATUS: 10,       // Cột J: Trạng thái
  DEADLINE_ALERT: 11// Cột K: Cảnh báo hạn chót
};

// Cấu hình bảng phụ Lịch sử điều chỉnh Deadline (Cột AA:AK -> Cột 27 đến 37)
const HISTORY_CONFIG = {
  START_COL: 27, // Cột AA
  NUM_COLS: 11,  // AA đến AK
  HEADERS: [
    "Mã thay đổi", "Mã task", "Tên công việc", "Deadline cũ", 
    "Deadline mới", "Chênh lệch ngày", "Lý do điều chỉnh", 
    "Người thực hiện", "Thời điểm thay đổi", "Trạng thái đồng bộ", "Ghi chú / lỗi"
  ]
};

/**
 * Hiển thị hộp thoại giao diện HTML
 */
function openChangeDeadlineDialog() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile("ChangeDeadlineDialog")
    .setWidth(560)
    .setHeight(640)
    .setTitle("Điều Chỉnh Deadline Công Việc");
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Điều Chỉnh Deadline Công Việc");
}

/**
 * Kiểm tra phân quyền người dùng (Stub cho giai đoạn tiếp theo)
 */
function canChangeTaskDeadline(userEmail, taskId) {
  // TODO: Tích hợp ma trận phân quyền Admin/PM khi module bảo mật hoàn thiện
  return true;
}

/**
 * Lấy cấu hình hệ thống từ sheet Cau_Hinh
 */
function getSystemConfig_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Cau_Hinh");
  const config = {
    MODE: "DRY_RUN",
    TIME_ZONE: "Asia/Ho_Chi_Minh",
    CALENDAR_ENABLED: false,
    CALENDAR_ID: "",
    ADMIN_EMAILS: ""
  };
  
  if (!configSheet) return config;
  const values = configSheet.getRange("A2:B20").getValues();
  values.forEach(row => {
    const key = String(row[0]).trim();
    const val = String(row[1]).trim();
    if (key === "MODE") config.MODE = val;
    if (key === "TIME_ZONE") config.TIME_ZONE = val || "Asia/Ho_Chi_Minh";
    if (key === "CALENDAR_ENABLED") config.CALENDAR_ENABLED = (val.toUpperCase() === "TRUE");
    if (key === "CALENDAR_ID") config.CALENDAR_ID = val;
    if (key === "ADMIN_EMAILS") config.ADMIN_EMAILS = val;
  });
  return config;
}

/**
 * Lấy danh sách Task hợp lệ để hiển thị trên giao diện
 */
function getTasksForDeadlineChange() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const taskSheet = ss.getSheetByName("Quan_Ly_Cong_Viec");
  if (!taskSheet) throw new Error("Không tìm thấy sheet Quan_Ly_Cong_Viec");

  const lastRow = taskSheet.getLastRow();
  if (lastRow < 2) return [];

  // Giới hạn đọc cố định từ A2:K[lastRow] để không bị ảnh hưởng bởi bảng lịch sử ở cột AA
  const numRows = lastRow - 1;
  const data = taskSheet.getRange(2, 1, numRows, 11).getValues();
  const taskList = [];

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const taskId = String(row[TASK_COL_INDEX.TASK_ID - 1]).trim();
    const taskName = String(row[TASK_COL_INDEX.TASK_NAME - 1]).trim();
    const status = String(row[TASK_COL_INDEX.STATUS - 1]).trim();
    const startDateRaw = row[TASK_COL_INDEX.START_DATE - 1];
    const endDateRaw = row[TASK_COL_INDEX.END_DATE - 1];

    // Bỏ qua dòng trống hoặc dòng tổng hợp
    if (!taskId || taskId.indexOf("TASK-") !== 0) continue;
    // Bỏ qua task đã hoàn thành theo quy tắc nghiệp vụ
    if (status.toLowerCase() === "hoàn thành") continue;

    const startDate = (startDateRaw instanceof Date && !isNaN(startDateRaw)) ? startDateRaw : null;
    const endDate = (endDateRaw instanceof Date && !isNaN(endDateRaw)) ? endDateRaw : null;

    taskList.push({
      taskId: taskId,
      taskName: taskName,
      status: status,
      assignee: String(row[TASK_COL_INDEX.ASSIGNEE - 1] || ""),
      startDateStr: startDate ? Utilities.formatDate(startDate, "GMT+7", "yyyy-MM-dd") : "",
      startDateDisplay: startDate ? Utilities.formatDate(startDate, "GMT+7", "dd/MM/yyyy") : "Chưa đặt",
      endDateStr: endDate ? Utilities.formatDate(endDate, "GMT+7", "yyyy-MM-dd") : "",
      endDateDisplay: endDate ? Utilities.formatDate(endDate, "GMT+7", "dd/MM/yyyy") : "Chưa đặt"
    });
  }

  return taskList;
}

/**
 * Xử lý cập nhật Deadline với cơ chế Lock chống xung đột
 */
function changeTaskDeadline(payload) {
  const lock = LockService.getDocumentLock();
  // Khóa tối đa 15 giây để xử lý giao dịch an toàn
  if (!lock.tryLock(15000)) {
    return { success: false, message: "Hệ thống đang bận do có người khác đang thao tác. Vui lòng thử lại sau vài giây!" };
  }

  try {
    const taskId = String(payload.taskId || "").trim();
    const newDeadlineStr = String(payload.newDeadline || "").trim();
    const clientOldDeadlineStr = String(payload.currentDeadline || "").trim();
    const reason = String(payload.reason || "").trim();

    if (!taskId) return { success: false, message: "Mã task không hợp lệ." };
    if (!newDeadlineStr) return { success: false, message: "Vui lòng chọn deadline mới." };
    if (!reason) return { success: false, message: "Lý do điều chỉnh deadline là bắt buộc." };

    // Kiểm tra quyền
    const currentUserEmail = Session.getActiveUser().getEmail() || "Không xác định";
    if (!canChangeTaskDeadline(currentUserEmail, taskId)) {
      return { success: false, message: "Bạn không có quyền thay đổi deadline của công việc này." };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const taskSheet = ss.getSheetByName("Quan_Ly_Cong_Viec");
    if (!taskSheet) return { success: false, message: "Không tìm thấy sheet Quan_Ly_Cong_Viec." };

    // Xác định dòng của task bằng cách quét cột A
    const lastRow = taskSheet.getLastRow();
    const idRange = taskSheet.getRange(2, TASK_COL_INDEX.TASK_ID, lastRow - 1, 1).getValues();
    let targetRow = -1;

    for (let i = 0; i < idRange.length; i++) {
      if (String(idRange[i][0]).trim() === taskId) {
        targetRow = i + 2; // +2 do offset dòng 1 tiêu đề
        break;
      }
    }

    if (targetRow === -1) {
      return { success: false, message: "Không tìm thấy mã công việc: " + taskId };
    }

    // Đọc lại dữ liệu thực tế tại thời điểm ghi để kiểm tra xung đột
    const currentRowData = taskSheet.getRange(targetRow, 1, 1, 11).getValues()[0];
    const taskName = currentRowData[TASK_COL_INDEX.TASK_NAME - 1];
    const currentStatus = String(currentRowData[TASK_COL_INDEX.STATUS - 1] || "").trim();
    const rawStartDate = currentRowData[TASK_COL_INDEX.START_DATE - 1];
    const rawEndDate = currentRowData[TASK_COL_INDEX.END_DATE - 1];

    if (currentStatus.toLowerCase() === "hoàn thành") {
      return { success: false, message: "Công việc đã ở trạng thái 'Hoàn thành'. Không thể thay đổi deadline." };
    }

    const actualCurrentDateStr = (rawEndDate instanceof Date && !isNaN(rawEndDate)) 
      ? Utilities.formatDate(rawEndDate, "GMT+7", "yyyy-MM-dd") 
      : "";

    // Kiểm tra xem có người khác vừa thay đổi deadline kể từ khi người dùng mở form không
    if (clientOldDeadlineStr && actualCurrentDateStr && clientOldDeadlineStr !== actualCurrentDateStr) {
      return { 
        success: false, 
        conflict: true,
        message: "Deadline công việc này vừa được thay đổi bởi một thành viên khác thành " + 
                 Utilities.formatDate(rawEndDate, "GMT+7", "dd/MM/yyyy") + 
                 ". Vui lòng tải lại dữ liệu!" 
      };
    }

    const newDeadlineDate = new Date(newDeadlineStr + "T00:00:00+07:00");
    if (isNaN(newDeadlineDate.getTime())) {
      return { success: false, message: "Định dạng ngày deadline mới không hợp lệ." };
    }

    // Kiểm tra ràng buộc Ngày bắt đầu <= Deadline mới
    if (rawStartDate instanceof Date && !isNaN(rawStartDate)) {
      const startDateNorm = new Date(Utilities.formatDate(rawStartDate, "GMT+7", "yyyy-MM-dd") + "T00:00:00+07:00");
      if (newDeadlineDate < startDateNorm) {
        return { 
          success: false, 
          message: "Deadline mới không được sớm hơn Ngày bắt đầu (" + 
                   Utilities.formatDate(rawStartDate, "GMT+7", "dd/MM/yyyy") + ")." 
        };
      }
    }

    // Nếu deadline mới giống deadline hiện tại thì từ chối cập nhật
    if (actualCurrentDateStr === newDeadlineStr) {
      return { success: false, message: "Deadline mới trùng với deadline hiện tại. Không có thay đổi nào được thực hiện." };
    }

    // Tính chênh lệch ngày
    let diffDays = 0;
    if (rawEndDate instanceof Date && !isNaN(rawEndDate)) {
      const oldDateNorm = new Date(Utilities.formatDate(rawEndDate, "GMT+7", "yyyy-MM-dd") + "T00:00:00+07:00");
      diffDays = Math.round((newDeadlineDate.getTime() - oldDateNorm.getTime()) / (1000 * 60 * 60 * 24));
    }
    const diffDaysStr = diffDays > 0 ? ("+" + diffDays) : String(diffDays);

    // 1. CẬP NHẬT CỘT F TẠI QUAN_LY_CONG_VIEC
    taskSheet.getRange(targetRow, TASK_COL_INDEX.END_DATE).setValue(newDeadlineDate);

    // 2. TĂNG PHIÊN BẢN LỊCH TẠI TRANG_THAI_TASK
    updateTaskCalendarVersion_(taskId);

    // 3. ĐỒNG BỘ DỊCH VỤ NGOÀI (CALENDAR & NOTIFICATION)
    const syncResult = syncDeadlineServices_(taskId, taskName, rawEndDate, newDeadlineDate, diffDaysStr, reason);

    // 4. GHI BẢN GHI LỊCH SỬ THAY ĐỔI VÀO CỘT AA:AK
    const changeId = "DL-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss");
    appendDeadlineHistory_(taskSheet, {
      changeId: changeId,
      taskId: taskId,
      taskName: taskName,
      oldDeadline: rawEndDate instanceof Date ? rawEndDate : "",
      newDeadline: newDeadlineDate,
      diffDays: diffDaysStr,
      reason: reason,
      user: currentUserEmail,
      timestamp: new Date(),
      syncStatus: syncResult.status,
      note: syncResult.note
    });

    return {
      success: true,
      message: "Cập nhật deadline cho công việc " + taskId + " thành công!",
      taskId: taskId,
      diffDays: diffDaysStr,
      newDeadline: Utilities.formatDate(newDeadlineDate, "GMT+7", "dd/MM/yyyy")
    };

  } catch (err) {
    return { success: false, message: "Lỗi hệ thống khi cập nhật: " + err.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Cập nhật phiên bản lịch tại sheet Trang_Thai_Task
 */
function updateTaskCalendarVersion_(taskId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statusSheet = ss.getSheetByName("Trang_Thai_Task");
  if (!statusSheet) return;

  const lastRow = statusSheet.getLastRow();
  if (lastRow < 2) return;

  const ids = statusSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === taskId) {
      const verCell = statusSheet.getRange(i + 2, 5); // Cột E: Phiên bản lịch
      const curVer = Number(verCell.getValue()) || 1;
      verCell.setValue(curVer + 1);
      break;
    }
  }
}

/**
 * Tự động mở rộng số cột và thiết lập bảng lịch sử tại AA:AK
 */
function ensureDeadlineHistoryTable_(taskSheet) {
  const requiredMaxCols = HISTORY_CONFIG.START_COL + HISTORY_CONFIG.NUM_COLS - 1; // 37 (Cột AK)
  const currentMaxCols = taskSheet.getMaxColumns();

  if (currentMaxCols < requiredMaxCols) {
    taskSheet.insertColumnsAfter(currentMaxCols, requiredMaxCols - currentMaxCols);
  }

  // Kiểm tra ô tiêu đề AA1
  const headerCheck = taskSheet.getRange(1, HISTORY_CONFIG.START_COL).getValue();
  if (!headerCheck || String(headerCheck).trim() === "") {
    const headerRange = taskSheet.getRange(1, HISTORY_CONFIG.START_COL, 1, HISTORY_CONFIG.NUM_COLS);
    headerRange.setValues([HISTORY_CONFIG.HEADERS]);
    headerRange.setBackground("#1A365D")
               .setFontColor("#FFFFFF")
               .setFontWeight("bold")
               .setHorizontalAlignment("center");
  }
}

/**
 * Ghi dòng lịch sử mới vào cột AA:AK
 */
function appendDeadlineHistory_(taskSheet, record) {
  ensureDeadlineHistoryTable_(taskSheet);

  // Tìm dòng trống đầu tiên trong phạm vi bảng lịch sử (cột AA)
  const lastSheetRow = taskSheet.getLastRow();
  let nextRow = 2;

  if (lastSheetRow >= 2) {
    const historyColValues = taskSheet.getRange(2, HISTORY_CONFIG.START_COL, lastSheetRow - 1, 1).getValues();
    for (let i = historyColValues.length - 1; i >= 0; i--) {
      if (historyColValues[i][0] && String(historyColValues[i][0]).trim() !== "") {
        nextRow = i + 3; // +2 do header, +1 dòng kế tiếp
        break;
      }
    }
  }

  const rowData = [
    record.changeId,
    record.taskId,
    record.taskName,
    record.oldDeadline,
    record.newDeadline,
    record.diffDays,
    record.reason,
    record.user,
    record.timestamp,
    record.syncStatus,
    record.note
  ];

  const targetRange = taskSheet.getRange(nextRow, HISTORY_CONFIG.START_COL, 1, HISTORY_CONFIG.NUM_COLS);
  targetRange.setValues([rowData]);

  // Định dạng hiển thị chuẩn
  taskSheet.getRange(nextRow, HISTORY_CONFIG.START_COL + 3).setNumberFormat("dd/MM/yyyy"); // Deadline cũ
  taskSheet.getRange(nextRow, HISTORY_CONFIG.START_COL + 4).setNumberFormat("dd/MM/yyyy"); // Deadline mới
  taskSheet.getRange(nextRow, HISTORY_CONFIG.START_COL + 5).setHorizontalAlignment("center"); // Chênh lệch ngày
  taskSheet.getRange(nextRow, HISTORY_CONFIG.START_COL + 8).setNumberFormat("dd/MM/yyyy HH:mm:ss"); // Thời điểm
}

/**
 * Đồng bộ Google Calendar và gửi Email thông báo
 */
function syncDeadlineServices_(taskId, taskName, oldDate, newDate, diffDaysStr, reason) {
  const config = getSystemConfig_();
  let syncStatus = "Thành công";
  const notes = [];

  // 1. Đồng bộ Calendar
  if (config.CALENDAR_ENABLED && config.CALENDAR_ID) {
    try {
      const cal = CalendarApp.getCalendarById(config.CALENDAR_ID);
      if (cal) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const calLinkSheet = ss.getSheetByName("Lien_Ket_Calendar");
        let eventFound = false;

        if (calLinkSheet && calLinkSheet.getLastRow() >= 2) {
          const links = calLinkSheet.getRange(2, 1, calLinkSheet.getLastRow() - 1, 7).getValues();
          for (let i = 0; i < links.length; i++) {
            if (String(links[i][1]).trim() === taskId) {
              const eventId = String(links[i][3]).trim();
              if (eventId) {
                const event = cal.getEventById(eventId);
                if (event) {
                  event.setAllDayDate(newDate);
                  event.setDescription((event.getDescription() || "") + 
                    "\n[Cập nhật Deadline]: Dời " + diffDaysStr + " ngày. Lý do: " + reason);
                  calLinkSheet.getRange(i + 2, 6).setValue(new Date()); // Lần đồng bộ
                  calLinkSheet.getRange(i + 2, 7).setValue("Đã đồng bộ");
                  eventFound = true;
                  notes.push("Calendar: Đã cập nhật sự kiện");
                  break;
                }
              }
            }
          }
        }
        if (!eventFound) {
          notes.push("Calendar: Chưa có liên kết sự kiện cũ");
        }
      }
    } catch (e) {
      syncStatus = "Lỗi đồng bộ";
      notes.push("Lỗi Calendar: " + e.message);
    }
  } else {
    notes.push("Calendar: Tắt");
  }

  // 2. Ghi nhật ký thông báo Email
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const notifSheet = ss.getSheetByName("Nhat_Ky_Thong_Bao");
    const oldDateStr = (oldDate instanceof Date && !isNaN(oldDate)) 
      ? Utilities.formatDate(oldDate, "GMT+7", "dd/MM/yyyy") 
      : "Chưa đặt";
    const newDateStr = Utilities.formatDate(newDate, "GMT+7", "dd/MM/yyyy");

    if (notifSheet) {
      const notifId = "NTF-" + Utilities.formatDate(new Date(), "GMT+7", "yyyyMMdd-HHmmss");
      const dedupeKey = taskId + "_DEADLINE_" + newDateStr;
      
      notifSheet.appendRow([
        notifId,
        taskId,
        "ĐIỀU CHỈNH DEADLINE",
        dedupeKey,
        config.ADMIN_EMAILS || "Ban Quản trị",
        new Date(),
        config.MODE === "LIVE" ? "Đã gửi" : "Mô phỏng (" + config.MODE + ")",
        1,
        ""
      ]);
    }
    notes.push("Email log: Đã ghi nhận");
  } catch (e) {
    notes.push("Lỗi Email log: " + e.message);
  }

  return { status: syncStatus, note: notes.join(" | ") };
}