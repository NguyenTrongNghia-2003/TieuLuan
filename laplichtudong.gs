var ws = SpreadsheetApp.getActiveSpreadsheet();
var ss = ws.getActiveSheet();
let eventDuration = 3; // thời gian của sự kiện (giờ)

// Hàm chọn tất cả sự kiện
function checkall() {
  for (var i = 2; i <= ss.getLastRow(); i++) {
    ss.getRange(i, 5).setValue("Check");
  }
}

// Hàm bỏ chọn tất cả sự kiện
function uncheckall() {
  for (var i = 2; i <= ss.getLastRow(); i++) {
    ss.getRange(i, 5).setValue("Uncheck");
  }
}

// Hàm thêm màu sắc và nhãn cho sự kiện dựa vào độ quan trọng
function getEventColor(priority) {
  if (priority === "Rất quan trọng") return CalendarApp.EventColor.RED;
  if (priority === "Quan trọng") return CalendarApp.EventColor.ORANGE;
  return CalendarApp.EventColor.GREEN; // Bình thường
}

// Hàm tạo sự kiện và gửi email
function createEvents() {
  var log = []; // Dùng để lưu log
  for (var i = 2; i <= ss.getLastRow(); i++) {
    const status = ss.getRange(i, 5).getValue();
    if (status === "Check") {
      let eventName = ss.getRange(i, 1).getValue();
      let date = ss.getRange(i, 2).getValue();
      let location = ss.getRange(i, 3).getValue();
      let guest = ss.getRange(i, 4).getValue();
      let priority = ss.getRange(i, 6).getValue();

      var startingDate = new Date(date);
      var endingDate = new Date(date);
      endingDate.setHours(startingDate.getHours() + eventDuration);

      try {
        // Tạo sự kiện trong Calendar
        var event = CalendarApp.getDefaultCalendar().createEvent(eventName, startingDate, endingDate, {
          location: location,
          guests: guest,
          sendInvites: true
        });
        event.setColor(getEventColor(priority));

        // Gửi email mời
        MailApp.sendEmail(
          guest,
          eventName,
          "Thư mời Quý vị đến dự: " + eventName +
          "\nBắt đầu vào lúc: " + startingDate +
          "\nKết thúc lúc: " + endingDate +
          "\nSự có mặt của Quý vị sẽ góp phần vào sự thành công của sự kiện!"
        );

        log.push(`Sự kiện \"${eventName}\" đã được tạo thành công.`);
      } catch (error) {
        log.push(`Lỗi khi tạo sự kiện \"${eventName}\": ${error.message}`);
      }
    }
  }

  // Lưu log vào một sheet riêng
  var logSheet = ws.getSheetByName("Log") || ws.insertSheet("Log");
  logSheet.clear(); // Xóa dữ liệu cũ
  logSheet.getRange(1, 1, log.length, 1).setValues(log.map(l => [l]));

  // Hiển thị thông báo hoàn tất
  SpreadsheetApp.getUi().alert("Hoàn tất việc tạo lịch. Kiểm tra sheet 'Log' để biết chi tiết.");
}

// Hàm cập nhật sự kiện
function updateEvents() {
  var calendar = CalendarApp.getDefaultCalendar(); // Sử dụng lịch mặc định của tài khoản
  for (var i = 2; i <= ss.getLastRow(); i++) {
    let eventName = ss.getRange(i, 1).getValue(); // Cột tên sự kiện
    let date = ss.getRange(i, 2).getValue();      // Cột thời gian
    let location = ss.getRange(i, 3).getValue(); // Cột địa điểm
    let guest = ss.getRange(i, 4).getValue();    // Cột email khách mời
    let priority = ss.getRange(i, 6).getValue(); // Cột mức độ quan trọng

    var startingDate = new Date(date);
    var endingDate = new Date(date);
    endingDate.setHours(startingDate.getHours() + eventDuration); // Thêm thời gian sự kiện

    try {
      // Tìm sự kiện dựa vào thời gian và tiêu đề
      var events = calendar.getEventsForDay(startingDate);
      var eventToUpdate = null;

      // Duyệt qua danh sách sự kiện trong ngày để tìm sự kiện khớp tên và thời gian
      for (var j = 0; j < events.length; j++) {
        var event = events[j];
        if (event.getTitle() === eventName && event.getStartTime().getTime() === startingDate.getTime()) {
          eventToUpdate = event;
          break;
        }
      }

      if (eventToUpdate) {
        // Cập nhật thông tin sự kiện
        eventToUpdate.setTitle(eventName);                // Cập nhật tên sự kiện
        eventToUpdate.setTime(startingDate, endingDate);  // Cập nhật thời gian
        eventToUpdate.setLocation(location);             // Cập nhật địa điểm
        eventToUpdate.setGuests(guest);                  // Cập nhật khách mời
        eventToUpdate.setColor(getEventColor(priority));  // Cập nhật màu sắc dựa vào mức độ quan trọng

        Logger.log(`Đã cập nhật sự kiện: ${eventName}`);
      } else {
        Logger.log(`Không tìm thấy sự kiện: ${eventName} vào thời gian ${startingDate}`);
      }
    } catch (error) {
      Logger.log(`Lỗi khi cập nhật sự kiện: ${eventName} - ${error.message}`);
    }
  }

  // Thông báo hoàn tất
  SpreadsheetApp.getUi().alert("Hoàn tất việc cập nhật sự kiện!");
}


// Hàm xóa sự kiện
function deleteEvents() {
  var calendar = CalendarApp.getDefaultCalendar(); // Lịch mặc định của Google Calendar

  for (var i = 2; i <= ss.getLastRow(); i++) {
    const status = ss.getRange(i, 5).getValue(); // Kiểm tra trạng thái tick ở cột 5
    if (status === true || status === "Check") { // Nếu sự kiện được tick
      let eventName = ss.getRange(i, 1).getValue(); // Cột tên sự kiện
      let date = ss.getRange(i, 2).getValue();      // Cột thời gian sự kiện

      var startingDate = new Date(date);

      try {
        // Tìm sự kiện trong ngày
        var events = calendar.getEventsForDay(startingDate);
        var eventToDelete = null;

        // Duyệt qua danh sách sự kiện trong ngày để tìm sự kiện khớp tên và thời gian
        for (var j = 0; j < events.length; j++) {
          var event = events[j];
          if (event.getTitle() === eventName && event.getStartTime().getTime() === startingDate.getTime()) {
            eventToDelete = event;
            break;
          }
        }

        if (eventToDelete) {
          // Xóa sự kiện
          eventToDelete.deleteEvent();
          Logger.log(`Đã xóa sự kiện: ${eventName}`);
        } else {
          Logger.log(`Không tìm thấy sự kiện để xóa: ${eventName}`);
        }
      } catch (error) {
        Logger.log(`Lỗi khi xóa sự kiện: ${eventName} - ${error.message}`);
      }
    }
  }

  // Thông báo hoàn tất
  SpreadsheetApp.getUi().alert("Hoàn tất việc xóa các sự kiện được tick!");
}


// Hàm lọc sự kiện theo ngày, tháng, năm
// Hàm lọc sự kiện theo ngày
function filterEventsByDate() {
  var logSheet = ws.getSheetByName("Filter Log") || ws.insertSheet("Filter Log");
  logSheet.clear();

  const filterDate = Browser.inputBox("Nhập ngày cần lọc (YYYY-MM-DD):");
  const events = [];

  for (var i = 2; i <= ss.getLastRow(); i++) {
    const eventDate = new Date(ss.getRange(i, 2).getValue());
    if (eventDate.toISOString().split('T')[0] === filterDate) {
      events.push([
        ss.getRange(i, 1).getValue(),
        eventDate,
        ss.getRange(i, 3).getValue(),
        ss.getRange(i, 4).getValue()
      ]);
    }
  }

  if (events.length > 0) {
    logSheet.getRange(1, 1, events.length, events[0].length).setValues(events);
    SpreadsheetApp.getUi().alert("Đã lọc xong. Kiểm tra sheet 'Filter Log'.");
  } else {
    SpreadsheetApp.getUi().alert("Không tìm thấy sự kiện nào phù hợp.");
  }
}

// Hàm lọc sự kiện theo tháng
function filterEventsByMonth() {
  var logSheet = ws.getSheetByName("Filter Log") || ws.insertSheet("Filter Log");
  logSheet.clear();

  const filterMonth = Browser.inputBox("Nhập tháng cần lọc (YYYY-MM):");
  const events = [];

  for (var i = 2; i <= ss.getLastRow(); i++) {
    const eventDate = new Date(ss.getRange(i, 2).getValue());
    const eventMonth = eventDate.getFullYear() + "-" + ("0" + (eventDate.getMonth() + 1)).slice(-2);
    if (eventMonth === filterMonth) {
      events.push([
        ss.getRange(i, 1).getValue(),
        eventDate,
        ss.getRange(i, 3).getValue(),
        ss.getRange(i, 4).getValue()
      ]);
    }
  }

  if (events.length > 0) {
    logSheet.getRange(1, 1, events.length, events[0].length).setValues(events);
    SpreadsheetApp.getUi().alert("Đã lọc xong. Kiểm tra sheet 'Filter Log'.");
  } else {
    SpreadsheetApp.getUi().alert("Không tìm thấy sự kiện nào phù hợp.");
  }
}

// Hàm lọc sự kiện theo năm
function filterEventsByYear() {
  var logSheet = ws.getSheetByName("Filter Log") || ws.insertSheet("Filter Log");
  logSheet.clear();

  const filterYear = Browser.inputBox("Nhập năm cần lọc (YYYY):");
  const events = [];

  for (var i = 2; i <= ss.getLastRow(); i++) {
    const eventDate = new Date(ss.getRange(i, 2).getValue());
    if (eventDate.getFullYear().toString() === filterYear) {
      events.push([
        ss.getRange(i, 1).getValue(),
        eventDate,
        ss.getRange(i, 3).getValue(),
        ss.getRange(i, 4).getValue()
      ]);
    }
  }

  if (events.length > 0) {
    logSheet.getRange(1, 1, events.length, events[0].length).setValues(events);
    SpreadsheetApp.getUi().alert("Đã lọc xong. Kiểm tra sheet 'Filter Log'.");
  } else {
    SpreadsheetApp.getUi().alert("Không tìm thấy sự kiện nào phù hợp.");
  }
}

// Trigger tự động tạo sự kiện và gửi email trước 30 phút
function addEmailReminder() {
  const now = new Date();

  for (let i = 2; i <= ss.getLastRow(); i++) {
    const eventDate = new Date(ss.getRange(i, 2).getValue()); // Cột thời gian sự kiện
    const guestEmail = ss.getRange(i, 4).getValue();          // Cột email khách mời
    const eventName = ss.getRange(i, 1).getValue();           // Cột tên sự kiện

    const reminderTime = new Date(eventDate.getTime() - 30 * 60 * 1000);

    if (Math.abs(now.getTime() - reminderTime.getTime()) <= 60 * 1000) { // Kiểm tra thời gian nhắc
      MailApp.sendEmail(
        guestEmail,
        "Lịch nhắc sự kiện: " + eventName,
        "Quý vị có lịch sự kiện: " + eventName +
        "\nThời gian: " + eventDate +
        "\nVui lòng chuẩn bị và tham gia đúng giờ."
      );
    }
  }
}

// Thêm menu vào Google Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("📩 Tác vụ sự kiện")
    .addItem("📋 Lên lịch và gửi email", "createEvents")
    .addItem("✏️ Cập nhật sự kiện", "updateEvents")
    .addItem("🗑️ Xóa sự kiện", "deleteEvents")
    .addSubMenu(
      ui.createMenu("🔎 Lọc sự kiện")
        .addItem("Theo ngày", "filterEventsByDate")
        .addItem("Theo tháng", "filterEventsByMonth")
        .addItem("Theo năm", "filterEventsByYear")
    )
    .addSeparator()
    .addItem("✔️ Chọn tất cả", "checkall")
    .addItem("❌ Bỏ chọn tất cả", "uncheckall")
    .addToUi();
}
