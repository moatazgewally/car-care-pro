function doGet(e) {
  var action = e ? e.parameter.action : "";
  var userId = e ? e.parameter.userId : "";
  var carId = e ? e.parameter.carId : "";
  
  if (action === "getCars") return respondJSON(getUserCars(userId));
  if (action === "getShifts") return respondJSON(getShiftsHistory(userId, carId));
  if (action === "getRules") return respondJSON(getMaintenanceRules(carId));
  
  return respondJSON({ success: true, message: "Car Care API Active" });
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action;
    var res = { success: false, message: "إجراء غير صالح" };

    if (action === "login") res = checkLogin(data.username, data.password);
    else if (action === "register") res = registerUser(data.username, data.password);
    else if (action === "addCar") res = addCar(data.userId, data.carName, data.odometer, data.installment, data.workDays);
    else if (action === "updateKm") res = updateCarKm(data.userId, data.carId, data.newKm);
    else if (action === "saveShift") res = saveShift(data.userId, data.carId, data.km, data.revenue, data.fuel, data.misc);
    else if (action === "addRule") res = addMaintenanceRule(data.carId, data.name, data.intervalKm, data.estCost, data.currentKm);

    return respondJSON(res);
  } catch (err) {
    return respondJSON({ success: false, message: err.message });
  }
}

function respondJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// باقي دوال قواعد البيانات والتشفير والعمليات المالية كما هي...
function getOrCreateSheet(sheetName, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) sheet.appendRow(headers);
  }
  return sheet;
}

function initDatabase() {
  getOrCreateSheet("Users", ["UserId", "Username", "PasswordHash", "CreatedAt"]);
  getOrCreateSheet("Cars", ["CarId", "UserId", "Name", "Odometer", "MonthlyInstallment", "WorkDays", "LastUpdate"]);
  getOrCreateSheet("Shifts", ["ShiftId", "UserId", "CarId", "Date", "Km", "Revenue", "Fuel", "Misc", "SavingTarget", "NetProfit", "InstallmentPart", "MaintPart"]);
  getOrCreateSheet("MaintenanceRules", ["RuleId", "CarId", "Name", "IntervalKm", "EstCost", "LastDoneKm"]);
  getOrCreateSheet("OdometerLogs", ["LogId", "CarId", "UserId", "Date", "OldKm", "NewKm"]);
}

function hashPassword(password) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  var txtHash = "";
  for (var i = 0; i < rawHash.length; i++) {
    var byteVal = rawHash[i];
    if (byteVal < 0) byteVal += 256;
    var byteStr = byteVal.toString(16);
    if (byteStr.length == 1) byteStr = '0' + byteStr;
    txtHash += byteStr;
  }
  return txtHash;
}

function registerUser(username, password) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    initDatabase();
    var cleanUser = String(username || '').trim();
    var cleanPass = String(password || '').trim();

    if (cleanUser.length < 3) return { success: false, message: "اسم المستخدم 3 أحرف على الأقل" };
    if (cleanPass.length < 6) return { success: false, message: "كلمة المرور 6 أحرف على الأقل" };

    var sheet = getOrCreateSheet("Users");
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === cleanUser.toLowerCase()) {
        return { success: false, message: "اسم المستخدم مسجل بالفعل!" };
      }
    }

    var userId = "u_" + Utilities.getUuid().substring(0, 8);
    var passHash = hashPassword(cleanPass);
    sheet.appendRow([userId, cleanUser, passHash, new Date()]);

    return { success: true, message: "تم إنشاء الحساب بنجاح!", data: { userId: userId, username: cleanUser } };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function checkLogin(username, password) {
  try {
    initDatabase();
    var cleanUser = String(username || '').trim().toLowerCase();
    var cleanPass = String(password || '').trim();

    var sheet = getOrCreateSheet("Users");
    var data = sheet.getDataRange().getValues();
    var inputHash = hashPassword(cleanPass);

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === cleanUser && String(data[i][2]) === inputHash) {
        return { success: true, message: "تم تسجيل الدخول بنجاح", data: { userId: data[i][0], username: data[i][1] } };
      }
    }
    return { success: false, message: "اسم المستخدم أو كلمة المرور غير صحيحة." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getUserCars(userId) {
  try {
    initDatabase();
    var sheet = getOrCreateSheet("Cars");
    var data = sheet.getDataRange().getValues();
    var userCars = [];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(userId)) {
        userCars.push({
          id: String(data[i][0]),
          name: String(data[i][2]),
          odometer: Number(data[i][3]) || 0,
          monthlyInstallment: Number(data[i][4]) || 0,
          workDaysPerMonth: Number(data[i][5]) || 20,
          lastUpdate: data[i][6] ? new Date(data[i][6]).toLocaleDateString('ar-EG') : '-'
        });
      }
    }
    return { success: true, data: userCars };
  } catch (e) {
    return { success: false, message: e.message, data: [] };
  }
}

function addCar(userId, carName, odometer, installment, workDays) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    initDatabase();
    var sheet = getOrCreateSheet("Cars");
    var carId = "c_" + Utilities.getUuid().substring(0, 8);
    sheet.appendRow([carId, userId, carName, Number(odometer), Number(installment), Number(workDays), new Date()]);
    return { success: true, message: "تمت إضافة السيارة بنجاح" };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function updateCarKm(userId, carId, newKm) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    initDatabase();
    var km = Number(newKm);
    var sheet = getOrCreateSheet("Cars");
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(carId) && String(data[i][1]) === String(userId)) {
        var oldKm = Number(data[i][3]) || 0;
        sheet.getRange(i + 1, 4).setValue(km);
        sheet.getRange(i + 1, 7).setValue(new Date());

        var logSheet = getOrCreateSheet("OdometerLogs");
        logSheet.appendRow(["log_" + Utilities.getUuid().substring(0, 8), carId, userId, new Date(), oldKm, km]);
        return { success: true, message: "تم تحديث العداد بنجاح" };
      }
    }
    return { success: false, message: "السيارة غير موجودة" };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getMaintenanceRules(carId) {
  try {
    initDatabase();
    var sheet = getOrCreateSheet("MaintenanceRules");
    var data = sheet.getDataRange().getValues();
    var rules = [];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(carId)) {
        rules.push({
          id: String(data[i][0]),
          carId: String(data[i][1]),
          name: String(data[i][2]),
          intervalKm: Number(data[i][3]) || 0,
          estCost: Number(data[i][4]) || 0,
          lastDoneKm: Number(data[i][5]) || 0
        });
      }
    }
    return { success: true, data: rules };
  } catch (e) {
    return { success: false, message: e.message, data: [] };
  }
}

function addMaintenanceRule(carId, name, intervalKm, estCost, currentKm) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    initDatabase();
    var sheet = getOrCreateSheet("MaintenanceRules");
    var ruleId = "m_" + Utilities.getUuid().substring(0, 8);
    sheet.appendRow([ruleId, carId, name, Number(intervalKm), Number(estCost), Number(currentKm)]);
    return { success: true, message: "تم إضافة بند الصيانة بنجاح" };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function saveShift(userId, carId, km, revenue, fuel, misc) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    initDatabase();

    var shiftKm = Number(km) || 0;
    var shiftRev = Number(revenue) || 0;
    var shiftFuel = Number(fuel) || 0;
    var shiftMisc = Number(misc) || 0;

    var carSheet = getOrCreateSheet("Cars");
    var carData = carSheet.getDataRange().getValues();
    var targetCar = null;
    var rowIndex = -1;

    for (var i = 1; i < carData.length; i++) {
      if (String(carData[i][0]) === String(carId) && String(carData[i][1]) === String(userId)) {
        targetCar = {
          odometer: Number(carData[i][3]) || 0,
          installment: Number(carData[i][4]) || 0,
          workDays: Number(carData[i][5]) || 20
        };
        rowIndex = i + 1;
        break;
      }
    }

    if (!targetCar) return { success: false, message: "السيارة غير موجودة" };

    var dailyInstallment = targetCar.workDays > 0 ? (targetCar.installment / targetCar.workDays) : 0;
    var maintSheet = getOrCreateSheet("MaintenanceRules");
    var maintData = maintSheet.getDataRange().getValues();
    var dailyMaintCost = 0;
    var hasRules = false;

    for (var j = 1; j < maintData.length; j++) {
      if (String(maintData[j][1]) === String(carId)) {
        var interval = Number(maintData[j][3]) || 0;
        var cost = Number(maintData[j][4]) || 0;
        if (interval > 0 && cost > 0) {
          dailyMaintCost += (cost / interval) * shiftKm;
          hasRules = true;
        }
      }
    }

    if (!hasRules && shiftKm > 0) dailyMaintCost = shiftKm * 0.35;

    var savingTarget = dailyInstallment + dailyMaintCost;
    var netProfit = shiftRev - (shiftFuel + shiftMisc + savingTarget);

    var shiftSheet = getOrCreateSheet("Shifts");
    var shiftId = "s_" + Utilities.getUuid().substring(0, 8);
    var now = new Date();

    shiftSheet.appendRow([
      shiftId, userId, carId, now, shiftKm, shiftRev, shiftFuel, shiftMisc, 
      Math.round(savingTarget), Math.round(netProfit), Math.round(dailyInstallment), Math.round(dailyMaintCost)
    ]);

    if (shiftKm > 0) {
      var newOdometer = targetCar.odometer + shiftKm;
      carSheet.getRange(rowIndex, 4).setValue(newOdometer);
      carSheet.getRange(rowIndex, 7).setValue(now);

      var logSheet = getOrCreateSheet("OdometerLogs");
      logSheet.appendRow(["log_" + Utilities.getUuid().substring(0, 8), carId, userId, now, targetCar.odometer, newOdometer]);
    }

    return {
      success: true,
      message: "تم حفظ الوردية بنجاح",
      data: {
        savingTarget: Math.round(savingTarget),
        netProfit: Math.round(netProfit),
        dailyInstallment: Math.round(dailyInstallment),
        dailyMaintCost: Math.round(dailyMaintCost)
      }
    };
  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}

function getShiftsHistory(userId, carId) {
  try {
    initDatabase();
    var sheet = getOrCreateSheet("Shifts");
    var data = sheet.getDataRange().getValues();
    var shifts = [];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(userId) && String(data[i][2]) === String(carId)) {
        var rawDate = data[i][3] ? new Date(data[i][3]) : new Date();
        var yyyy = rawDate.getFullYear();
        var mm = String(rawDate.getMonth() + 1).padStart(2, '0');
        var dd = String(rawDate.getDate()).padStart(2, '0');

        shifts.push({
          id: String(data[i][0]),
          date: rawDate.toLocaleDateString('ar-EG'),
          isoDate: yyyy + '-' + mm + '-' + dd,
          km: Number(data[i][4]) || 0,
          revenue: Number(data[i][5]) || 0,
          fuel: Number(data[i][6]) || 0,
          misc: Number(data[i][7]) || 0,
          saving: Number(data[i][8]) || 0,
          net: Number(data[i][9]) || 0,
          installment: Number(data[i][10]) || 0,
          maint: Number(data[i][11]) || 0
        });
      }
    }
    return { success: true, data: shifts };
  } catch (e) {
    return { success: false, message: e.message, data: [] };
  }
}
