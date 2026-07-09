/**
 * ============================================================
 *  ট্যুর ম্যানেজমেন্ট সিস্টেম - Google Apps Script ব্যাকএন্ড (v2)
 * ============================================================
 */

const SHEET_NAMES = {
  PACKAGES: 'Packages',
  BOOKINGS: 'Bookings',
  SEATS: 'Seats',
  EXPENSES: 'Expenses',
  ADMINS: 'Admins'
};

// ---------- ইউটিলিটি ----------

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function generateId(prefix) {
  return prefix + '-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => (obj[h] = row[i]));
      return obj;
    });
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- প্রাথমিক সেটআপ (একবার চালান) ----------

function setupSheets() {
  const pkg = getSheet(SHEET_NAMES.PACKAGES);
  pkg.clear();
  pkg.appendRow([
    'id', 'spotName', 'startDate', 'endDate', 'totalDays',
    'delegateCount', 'delegateFee', 'vehicles', 'status',
    'createdAt', 'slug', 'coverNote'
  ]);

  const bk = getSheet(SHEET_NAMES.BOOKINGS);
  bk.clear();
  bk.appendRow([
    'id', 'packageId', 'serial', 'date', 'name', 'mobile', 'address',
    'seatsJson', 'delegateFee', 'paidAmount', 'dueAmount',
    'status', 'createdAt', 'confirmedBy', 'confirmedAt', 'paymentMethod', 'transactionId', 'receivedBy'
  ]);

  const st = getSheet(SHEET_NAMES.SEATS);
  st.clear();
  st.appendRow(['packageId', 'vehicleName', 'seatNo', 'status', 'bookingId', 'holdNote']);

  const ex = getSheet(SHEET_NAMES.EXPENSES);
  ex.clear();
  ex.appendRow(['id', 'packageId', 'title', 'category', 'amount', 'note', 'date', 'addedBy', 'createdAt']);

  const ad = getSheet(SHEET_NAMES.ADMINS);
  ad.clear();
  ad.appendRow(['username', 'password', 'role', 'name', 'mobile', 'email', 'address', 'photoUrl', 'active']);
  ad.appendRow(['admin1', 'admin123', 'admin', 'এডমিন ১', '01700000001', '', '', '', true]);
  ad.appendRow(['admin2', 'admin123', 'admin', 'এডমিন ২', '01700000002', '', '', '', true]);
  ad.appendRow(['moderator1', 'pass123', 'moderator', 'মডারেটর ১', '01700000003', '', '', '', true]);
  ad.appendRow(['host1', 'pass123', 'host', 'হোস্ট ১', '01700000004', '', '', '', true]);

  SpreadsheetApp.flush();
  return 'Setup সম্পন্ন হয়েছে!';
}

// ---------- ওয়েব API এন্ট্রি পয়েন্ট ----------

function doGet(e) {
  const action = e.parameter.action;
  try {
    switch (action) {
      case 'getPackagesSummary':
        return jsonResponse({ ok: true, data: getPackagesSummary() });
      case 'getPackages':
        return jsonResponse({ ok: true, data: getPackages() });
      case 'getPackage':
        return jsonResponse({ ok: true, data: getPackageDetail(e.parameter.id) });
      case 'getPackageFull':
        return jsonResponse({ ok: true, data: getPackageFull(e.parameter.id) });
      case 'getSeats':
        return jsonResponse({ ok: true, data: getSeatsForPackage(e.parameter.packageId) });
      case 'getBookings':
        return jsonResponse({ ok: true, data: getBookings(e.parameter.packageId) });
      case 'getExpenses':
        return jsonResponse({ ok: true, data: getExpenses(e.parameter.packageId) });
      case 'getReport':
        return jsonResponse({ ok: true, data: getFinanceReport(e.parameter.packageId) });
      case 'getOverallReport':
        return jsonResponse({ ok: true, data: getOverallReport() });
      case 'findTickets':
        return jsonResponse({ ok: true, data: findTickets(e.parameter.name, e.parameter.mobile) });
      case 'getAdminNames':
        return jsonResponse({ ok: true, data: getAdminNames() });
      case 'getTeamContacts':
        return jsonResponse({ ok: true, data: getTeamContacts() });
      case 'getPackageCounts':
        return jsonResponse({ ok: true, data: getPackageCounts() });
      default:
        return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  try {
    switch (action) {
      case 'login':
        return jsonResponse(login(body.username, body.password));
      case 'createPackage':
        return jsonResponse({ ok: true, data: createPackage(body) });
      case 'updatePackage':
        return jsonResponse(updatePackage(body));
      case 'deletePackage':
        return jsonResponse(deletePackage(body.id));
      case 'createBooking':
        return jsonResponse(createBooking(body));
      case 'confirmBooking':
        return jsonResponse(confirmBooking(body.bookingId, body.confirmedBy));
      case 'cancelBooking':
        return jsonResponse(cancelBooking(body.bookingId));
      case 'updatePayment':
        return jsonResponse(updatePayment(body.bookingId, body.paidAmount));
      case 'holdSeats':
        return jsonResponse(holdSeats(body.packageId, body.vehicleName, body.seats, body.holdNote));
      case 'unholdSeats':
        return jsonResponse(unholdSeats(body.packageId, body.vehicleName, body.seats));
      case 'releaseSeat':
        return jsonResponse(releaseSeat(body.packageId, body.vehicleName, body.seatNo));
      case 'addExpense':
        return jsonResponse({ ok: true, data: addExpense(body) });
      case 'deleteExpense':
        return jsonResponse(deleteExpense(body.id));
      case 'manageAdmin':
        return jsonResponse(manageAdmin(body));
      case 'getAdmins':
        return jsonResponse({ ok: true, data: getAdmins() });
      default:
        return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ---------- লগইন / এডমিন ----------

function login(username, password) {
  const admins = sheetToObjects(getSheet(SHEET_NAMES.ADMINS));
  const found = admins.find(a => a.username === username && String(a.password) === String(password) && a.active);
  if (!found) return { ok: false, error: 'ইউজারনেম বা পাসওয়ার্ড ভুল, অথবা একাউন্ট নিষ্ক্রিয়।' };
  return {
    ok: true,
    data: {
      username: found.username, role: found.role, name: found.name,
      mobile: found.mobile, email: found.email, address: found.address, photoUrl: found.photoUrl
    }
  };
}

function getAdmins() {
  return sheetToObjects(getSheet(SHEET_NAMES.ADMINS)).map(a => ({
    username: a.username, role: a.role, name: a.name,
    mobile: a.mobile, email: a.email, address: a.address, photoUrl: a.photoUrl, active: a.active
  }));
}

// পাবলিক ব্যবহারের জন্য - শুধু সক্রিয় এডমিন/মডারেটর/হোস্টদের নাম (বুকিং ফরমে "কার কাছে টাকা দিলেন" সিলেক্ট করতে)
function getAdminNames() {
  return sheetToObjects(getSheet(SHEET_NAMES.ADMINS))
    .filter(a => a.active)
    .map(a => ({ username: a.username, name: a.name, role: a.role }));
}

// পাবলিক "যোগাযোগ" পেজের জন্য - নাম, মোবাইল, ইমেইল, ঠিকানা, ছবি (পাসওয়ার্ড বাদে)
function getTeamContacts() {
  return sheetToObjects(getSheet(SHEET_NAMES.ADMINS))
    .filter(a => a.active)
    .map(a => ({
      name: a.name, role: a.role, mobile: a.mobile,
      email: a.email, address: a.address, photoUrl: a.photoUrl
    }));
}

function manageAdmin(body) {
  const sheet = getSheet(SHEET_NAMES.ADMINS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const usernameCol = headers.indexOf('username');

  if (body.mode === 'add') {
    sheet.appendRow([
      body.username, body.password, body.role, body.name,
      body.mobile || '', body.email || '', body.address || '', body.photoUrl || '', true
    ]);
    SpreadsheetApp.flush();
    return { ok: true };
  }
  for (let i = 1; i < data.length; i++) {
    if (data[i][usernameCol] === body.username) {
      if (body.mode === 'delete') { sheet.deleteRow(i + 1); SpreadsheetApp.flush(); return { ok: true }; }
      if (body.mode === 'edit') {
        if (body.password) sheet.getRange(i + 1, headers.indexOf('password') + 1).setValue(body.password);
        if (body.role) sheet.getRange(i + 1, headers.indexOf('role') + 1).setValue(body.role);
        if (body.name) sheet.getRange(i + 1, headers.indexOf('name') + 1).setValue(body.name);
        if (body.mobile !== undefined) sheet.getRange(i + 1, headers.indexOf('mobile') + 1).setValue(body.mobile);
        if (body.email !== undefined) sheet.getRange(i + 1, headers.indexOf('email') + 1).setValue(body.email);
        if (body.address !== undefined) sheet.getRange(i + 1, headers.indexOf('address') + 1).setValue(body.address);
        if (body.photoUrl !== undefined) sheet.getRange(i + 1, headers.indexOf('photoUrl') + 1).setValue(body.photoUrl);
        if (body.active !== undefined) sheet.getRange(i + 1, headers.indexOf('active') + 1).setValue(body.active);
        SpreadsheetApp.flush();
        return { ok: true };
      }
    }
  }
  return { ok: false, error: 'এডমিন পাওয়া যায়নি' };
}

// ---------- প্যাকেজ ----------

function slugify(text, id) {
  const base = text.toString().trim()
    .replace(/[^\w\u0980-\u09FF\s-]/g, '')
    .replace(/\s+/g, '-')
    .toLowerCase();
  return base + '-' + id.split('-')[1];
}

function createPackage(body) {
  const sheet = getSheet(SHEET_NAMES.PACKAGES);
  const id = generateId('PKG');
  const slug = slugify(body.spotName, id);
  const vehicles = body.vehicles || [];

  sheet.appendRow([
    id, body.spotName, body.startDate, body.endDate, body.totalDays,
    body.delegateCount, body.delegateFee, JSON.stringify(vehicles), 'active',
    new Date().toISOString(), slug, body.coverNote || ''
  ]);

  addSeatsForVehicles(id, vehicles);
  SpreadsheetApp.flush();
  return { id, slug };
}

function addSeatsForVehicles(packageId, vehicles) {
  const seatSheet = getSheet(SHEET_NAMES.SEATS);
  const existing = sheetToObjects(seatSheet).filter(s => s.packageId === packageId);
  const existingVehicleNames = new Set(existing.map(s => s.vehicleName));
  const rows = [];

  vehicles.forEach(v => {
    if (existingVehicleNames.has(v.name)) return; // আগে থেকে থাকা গাড়ির সীট আবার তৈরি হবে না
    const seatNos = generateSeatNumbers(v.seatCount);
    seatNos.forEach(seatNo => {
      rows.push([packageId, v.name, seatNo, 'available', '', '']);
    });
  });
  if (rows.length > 0) {
    seatSheet.getRange(seatSheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }
}

// সীট নম্বর তৈরি: প্রতি সারিতে ৪টি (A1-A4, B1-B4...)। শেষে ১টি সীট বাকি থাকলে
// নতুন সারি (যেমন K1) না বানিয়ে আগের সারিতেই ৫ম সীট (J5) হিসেবে যোগ হবে।
function generateSeatNumbers(seatCount) {
  const rowLetters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  const seatsPerRow = 4;
  const fullRows = Math.floor(seatCount / seatsPerRow);
  const remainder = seatCount % seatsPerRow;
  const seatNos = [];

  let totalRows = fullRows;
  let lastRowExtra = 0;

  if (remainder === 1 && fullRows > 0) {
    // শেষ ১টি বিজোড় সীট আগের সারির সাথে ৫ম সীট হিসেবে যুক্ত হবে
    totalRows = fullRows; // সারি সংখ্যা একই থাকবে
    lastRowExtra = 1;
  } else if (remainder > 0) {
    totalRows = fullRows + 1; // সাধারণ বিজোড় শেষ সারি (২ বা ৩টি সীট)
  }

  for (let r = 0; r < totalRows; r++) {
    const isLastRow = r === totalRows - 1;
    const countInRow = isLastRow ? (remainder === 1 && fullRows > 0 ? seatsPerRow + lastRowExtra : (remainder > 0 && r === fullRows ? remainder : seatsPerRow)) : seatsPerRow;
    for (let c = 1; c <= countInRow; c++) {
      seatNos.push(rowLetters[r] + c);
    }
  }
  return seatNos;
}

function updatePackage(body) {
  const sheet = getSheet(SHEET_NAMES.PACKAGES);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === body.id) {
      const row = i + 1;
      const setIf = (field, value) => {
        if (value !== undefined && value !== null && value !== '') {
          sheet.getRange(row, headers.indexOf(field) + 1).setValue(value);
        }
      };
      setIf('spotName', body.spotName);
      setIf('startDate', body.startDate);
      setIf('endDate', body.endDate);
      setIf('totalDays', body.totalDays);
      setIf('delegateCount', body.delegateCount);
      setIf('delegateFee', body.delegateFee);
      setIf('coverNote', body.coverNote);
      setIf('status', body.status);

      if (body.vehicles && body.vehicles.length > 0) {
        const oldVehicles = JSON.parse(data[i][headers.indexOf('vehicles')] || '[]');
        const merged = mergeVehicles(oldVehicles, body.vehicles);
        sheet.getRange(row, headers.indexOf('vehicles') + 1).setValue(JSON.stringify(merged));
        addSeatsForVehicles(body.id, merged); // নতুন গাড়ির জন্য সীট তৈরি (পুরনো গাড়ি অপরিবর্তিত থাকবে)
      }
      SpreadsheetApp.flush();
      return { ok: true };
    }
  }
  return { ok: false, error: 'প্যাকেজ পাওয়া যায়নি' };
}

function mergeVehicles(oldList, newList) {
  const map = {};
  oldList.forEach(v => (map[v.name] = v));
  newList.forEach(v => (map[v.name] = v)); // নতুন তথ্য দিয়ে ওভাররাইট বা যোগ
  return Object.values(map);
}

function getPackages() {
  return sheetToObjects(getSheet(SHEET_NAMES.PACKAGES)).map(p => {
    p.vehicles = JSON.parse(p.vehicles || '[]');
    return p;
  });
}

// হোমপেজের উপরে দেখানোর জন্য - কতটি চলমান, কতটি সম্পন্ন
function getPackageCounts() {
  const packages = getPackages();
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  let running = 0, completed = 0;
  packages.forEach(p => {
    const endDate = p.endDate ? new Date(p.endDate) : null;
    const isPast = endDate ? endDate < today : false;
    if (isPast) completed++;
    else if (p.status === 'active') running++;
  });
  return { running, completed, total: packages.length };
}

// একটি কল দিয়ে সব প্যাকেজ + সীট সামারি (হোমপেজের জন্য ফাস্ট)
function getPackagesSummary() {
  const packages = getPackages();
  const allSeats = sheetToObjects(getSheet(SHEET_NAMES.SEATS));
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return packages.map(p => {
    const seats = allSeats.filter(s => s.packageId === p.id);
    const taken = seats.filter(s => s.status !== 'available').length;
    p.totalSeats = seats.length;
    p.availableSeats = seats.length - taken;
    const endDate = p.endDate ? new Date(p.endDate) : null;
    p.isPast = endDate ? endDate < today : false;
    return p;
  });
}

function deletePackage(id) {
  // Packages শীট থেকে মুছে ফেলা
  const pkgSheet = getSheet(SHEET_NAMES.PACKAGES);
  const pkgData = pkgSheet.getDataRange().getValues();
  const pkgHeaders = pkgData[0];
  const idCol = pkgHeaders.indexOf('id');
  let found = false;
  for (let i = pkgData.length - 1; i >= 1; i--) {
    if (pkgData[i][idCol] === id) { pkgSheet.deleteRow(i + 1); found = true; break; }
  }
  if (!found) return { ok: false, error: 'প্যাকেজ পাওয়া যায়নি' };

  // সংশ্লিষ্ট Seats, Bookings, Expenses ও মুছে ফেলা (ডেটা পরিষ্কার রাখতে)
  deleteRowsByPackageId(getSheet(SHEET_NAMES.SEATS), id);
  deleteRowsByPackageId(getSheet(SHEET_NAMES.BOOKINGS), id);
  deleteRowsByPackageId(getSheet(SHEET_NAMES.EXPENSES), id);

  SpreadsheetApp.flush();
  return { ok: true };
}

function deleteRowsByPackageId(sheet, packageId) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = headers.indexOf('packageId');
  if (col === -1) return;
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][col] === packageId) sheet.deleteRow(i + 1);
  }
}

function getPackageDetail(idOrSlug) {
  const packages = getPackages();
  const pkg = packages.find(p => p.id === idOrSlug || p.slug === idOrSlug);
  if (!pkg) return null;
  const seats = getSeatsForPackage(pkg.id);
  const total = seats.length;
  const taken = seats.filter(s => s.status !== 'available').length;
  pkg.totalSeats = total;
  pkg.availableSeats = total - taken;
  return pkg;
}

// একটি কল দিয়ে প্যাকেজ + সীট (বুকিং পেজের জন্য ফাস্ট)
function getPackageFull(idOrSlug) {
  const pkg = getPackageDetail(idOrSlug);
  if (!pkg) return null;
  return { package: pkg, seats: getSeatsForPackage(pkg.id) };
}

// ---------- সিট ----------

function getSeatsForPackage(packageId) {
  return sheetToObjects(getSheet(SHEET_NAMES.SEATS)).filter(s => s.packageId === packageId);
}

function findSeatRow(sheet, data, headers, packageId, vehicleName, seatNo) {
  for (let i = 1; i < data.length; i++) {
    if (
      data[i][headers.indexOf('packageId')] === packageId &&
      data[i][headers.indexOf('vehicleName')] === vehicleName &&
      data[i][headers.indexOf('seatNo')] === seatNo
    ) {
      return i + 1;
    }
  }
  return null;
}

function holdSeats(packageId, vehicleName, seats, holdNote) {
  const sheet = getSheet(SHEET_NAMES.SEATS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const results = [];
  seats.forEach(seatNo => {
    const rowIndex = findSeatRow(sheet, data, headers, packageId, vehicleName, seatNo);
    if (rowIndex) {
      const statusCol = headers.indexOf('status') + 1;
      const currentStatus = sheet.getRange(rowIndex, statusCol).getValue();
      if (currentStatus === 'available') {
        sheet.getRange(rowIndex, statusCol).setValue('held');
        sheet.getRange(rowIndex, headers.indexOf('holdNote') + 1).setValue(holdNote || 'এডমিন হোল্ড');
        results.push({ seatNo, ok: true });
      } else {
        results.push({ seatNo, ok: false });
      }
    }
  });
  SpreadsheetApp.flush();
  return { ok: true, results };
}

function unholdSeats(packageId, vehicleName, seats) {
  const sheet = getSheet(SHEET_NAMES.SEATS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const results = [];
  seats.forEach(seatNo => {
    const rowIndex = findSeatRow(sheet, data, headers, packageId, vehicleName, seatNo);
    if (rowIndex) {
      const statusCol = headers.indexOf('status') + 1;
      const currentStatus = sheet.getRange(rowIndex, statusCol).getValue();
      if (currentStatus === 'held') {
        sheet.getRange(rowIndex, statusCol).setValue('available');
        sheet.getRange(rowIndex, headers.indexOf('holdNote') + 1).setValue('');
        sheet.getRange(rowIndex, headers.indexOf('bookingId') + 1).setValue('');
        results.push({ seatNo, ok: true });
      } else {
        results.push({ seatNo, ok: false });
      }
    }
  });
  SpreadsheetApp.flush();
  return { ok: true, results };
}

function releaseSeat(packageId, vehicleName, seatNo) {
  const sheet = getSheet(SHEET_NAMES.SEATS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rowIndex = findSeatRow(sheet, data, headers, packageId, vehicleName, seatNo);
  if (!rowIndex) return { ok: false, error: 'সিট পাওয়া যায়নি' };
  sheet.getRange(rowIndex, headers.indexOf('status') + 1).setValue('available');
  sheet.getRange(rowIndex, headers.indexOf('bookingId') + 1).setValue('');
  sheet.getRange(rowIndex, headers.indexOf('holdNote') + 1).setValue('');
  SpreadsheetApp.flush();
  return { ok: true };
}

// ---------- বুকিং (একাধিক সীট সাপোর্ট সহ) ----------

function getBookings(packageId) {
  const all = sheetToObjects(getSheet(SHEET_NAMES.BOOKINGS)).map(b => {
    b.seats = JSON.parse(b.seatsJson || '[]');
    return b;
  });
  if (packageId) return all.filter(b => b.packageId === packageId);
  return all;
}

// body.seats = [{vehicleName, seatNo}, ...]  একাধিক সীট একসাথে বুক করা যাবে
function createBooking(body) {
  const seatSheet = getSheet(SHEET_NAMES.SEATS);
  const seatData = seatSheet.getDataRange().getValues();
  const seatHeaders = seatData[0];
  const statusCol = seatHeaders.indexOf('status') + 1;
  const bookingCol = seatHeaders.indexOf('bookingId') + 1;

  const requestedSeats = body.seats || [{ vehicleName: body.vehicleName, seatNo: body.seatNo }];

  // প্রথমে সব সীট খালি আছে কিনা যাচাই
  const rowIndexes = [];
  for (const s of requestedSeats) {
    const rowIndex = findSeatRow(seatSheet, seatData, seatHeaders, body.packageId, s.vehicleName, s.seatNo);
    if (!rowIndex) return { ok: false, error: 'সিট ' + s.seatNo + ' পাওয়া যায়নি' };
    const currentStatus = seatData[rowIndex - 1][seatHeaders.indexOf('status')];
    if (currentStatus !== 'available') {
      return { ok: false, error: 'দুঃখিত, সিট ' + s.seatNo + ' ইতিমধ্যে বুক/হোল্ড করা আছে। পেজ রিফ্রেশ করে আবার চেষ্টা করুন।' };
    }
    rowIndexes.push(rowIndex);
  }

  const bookingSheet = getSheet(SHEET_NAMES.BOOKINGS);
  const id = generateId('BK');
  const existingCount = getBookings(body.packageId).length;
  const serial = existingCount + 1;

  const perSeatFee = Number(body.delegateFee) || 0;
  const totalFee = perSeatFee * requestedSeats.length;
  const paidAmount = Number(body.paidAmount) || 0;
  const dueAmount = totalFee - paidAmount;

  bookingSheet.appendRow([
    id, body.packageId, serial, body.date, body.name, "'" + String(body.mobile), body.address,
    JSON.stringify(requestedSeats), totalFee, paidAmount, dueAmount,
    'pending', new Date().toISOString(), '', '', body.paymentMethod || '', body.transactionId || '', body.receivedBy || ''
  ]);

  rowIndexes.forEach(rowIndex => {
    seatSheet.getRange(rowIndex, statusCol).setValue('held');
    seatSheet.getRange(rowIndex, bookingCol).setValue(id);
  });

  SpreadsheetApp.flush();
  return { ok: true, data: { id, serial, dueAmount, seatCount: requestedSeats.length, totalFee } };
}

function confirmBooking(bookingId, confirmedBy) {
  const sheet = getSheet(SHEET_NAMES.BOOKINGS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === bookingId) {
      const row = i + 1;
      sheet.getRange(row, headers.indexOf('status') + 1).setValue('confirmed');
      sheet.getRange(row, headers.indexOf('confirmedBy') + 1).setValue(confirmedBy);
      sheet.getRange(row, headers.indexOf('confirmedAt') + 1).setValue(new Date().toISOString());

      const packageId = data[i][headers.indexOf('packageId')];
      const seats = JSON.parse(data[i][headers.indexOf('seatsJson')] || '[]');

      const seatSheet = getSheet(SHEET_NAMES.SEATS);
      const seatData = seatSheet.getDataRange().getValues();
      const seatHeaders = seatData[0];
      seats.forEach(s => {
        const rowIndex = findSeatRow(seatSheet, seatData, seatHeaders, packageId, s.vehicleName, s.seatNo);
        if (rowIndex) seatSheet.getRange(rowIndex, seatHeaders.indexOf('status') + 1).setValue('booked');
      });
      SpreadsheetApp.flush();
      return { ok: true };
    }
  }
  return { ok: false, error: 'বুকিং পাওয়া যায়নি' };
}

function cancelBooking(bookingId) {
  const sheet = getSheet(SHEET_NAMES.BOOKINGS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === bookingId) {
      const row = i + 1;
      sheet.getRange(row, headers.indexOf('status') + 1).setValue('cancelled');
      const packageId = data[i][headers.indexOf('packageId')];
      const seats = JSON.parse(data[i][headers.indexOf('seatsJson')] || '[]');
      seats.forEach(s => releaseSeat(packageId, s.vehicleName, s.seatNo));
      SpreadsheetApp.flush();
      return { ok: true };
    }
  }
  return { ok: false, error: 'বুকিং পাওয়া যায়নি' };
}

function updatePayment(bookingId, paidAmount) {
  const sheet = getSheet(SHEET_NAMES.BOOKINGS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === bookingId) {
      const row = i + 1;
      const fee = Number(data[i][headers.indexOf('delegateFee')]) || 0;
      const paid = Number(paidAmount) || 0;
      sheet.getRange(row, headers.indexOf('paidAmount') + 1).setValue(paid);
      sheet.getRange(row, headers.indexOf('dueAmount') + 1).setValue(fee - paid);
      SpreadsheetApp.flush();
      return { ok: true, data: { dueAmount: fee - paid } };
    }
  }
  return { ok: false, error: 'বুকিং পাওয়া যায়নি' };
}

// ---------- টিকেট লুকআপ (নাম + মোবাইল দিয়ে) ----------

function findTickets(name, mobile) {
  const bookings = getBookings().filter(b => b.status === 'confirmed');
  const packages = getPackages();
  const pkgMap = {};
  packages.forEach(p => (pkgMap[p.id] = p));

  const matches = bookings.filter(b => {
    const nameMatch = name ? b.name.toString().trim().toLowerCase().includes(name.trim().toLowerCase()) : true;
    const mobileMatch = mobile ? b.mobile.toString().trim() === mobile.trim() : true;
    return nameMatch && mobileMatch;
  });

  return matches.map(b => ({
    ...b,
    packageInfo: pkgMap[b.packageId] ? {
      spotName: pkgMap[b.packageId].spotName,
      startDate: pkgMap[b.packageId].startDate,
      endDate: pkgMap[b.packageId].endDate
    } : null
  }));
}

// ---------- খরচ ----------

function getExpenses(packageId) {
  const all = sheetToObjects(getSheet(SHEET_NAMES.EXPENSES));
  if (packageId) return all.filter(x => x.packageId === packageId);
  return all;
}

function addExpense(body) {
  const sheet = getSheet(SHEET_NAMES.EXPENSES);
  const id = generateId('EXP');
  sheet.appendRow([
    id, body.packageId, body.title, body.category || 'অন্যান্য',
    Number(body.amount) || 0, body.note || '', body.date || new Date().toISOString().split('T')[0],
    body.addedBy || '', new Date().toISOString()
  ]);
  SpreadsheetApp.flush();
  return { id };
}

function deleteExpense(id) {
  const sheet = getSheet(SHEET_NAMES.EXPENSES);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) { sheet.deleteRow(i + 1); SpreadsheetApp.flush(); return { ok: true }; }
  }
  return { ok: false, error: 'পাওয়া যায়নি' };
}

// ---------- রিপোর্ট ----------

function getFinanceReport(packageId) {
  const bookings = getBookings(packageId).filter(b => b.status !== 'cancelled');
  const expenses = getExpenses(packageId);

  const totalIncome = bookings.reduce((sum, b) => sum + (Number(b.paidAmount) || 0), 0);
  const totalDue = bookings.reduce((sum, b) => sum + (Number(b.dueAmount) || 0), 0);
  const totalExpected = bookings.reduce((sum, b) => sum + (Number(b.delegateFee) || 0), 0);
  const totalExpense = expenses.reduce((sum, x) => sum + (Number(x.amount) || 0), 0);

  return {
    packageId,
    totalBookings: bookings.length,
    confirmedBookings: bookings.filter(b => b.status === 'confirmed').length,
    pendingBookings: bookings.filter(b => b.status === 'pending').length,
    totalExpected, totalIncome, totalDue, totalExpense,
    netProfit: totalIncome - totalExpense,
    expenses
  };
}

function getOverallReport() {
  const packages = getPackages();
  const allBookings = getBookings().filter(b => b.status !== 'cancelled');
  const allExpenses = getExpenses();

  const totalIncome = allBookings.reduce((sum, b) => sum + (Number(b.paidAmount) || 0), 0);
  const totalDue = allBookings.reduce((sum, b) => sum + (Number(b.dueAmount) || 0), 0);
  const totalExpense = allExpenses.reduce((sum, x) => sum + (Number(x.amount) || 0), 0);

  const perPackage = packages.map(p => {
    const pBookings = allBookings.filter(b => b.packageId === p.id);
    const pExpenses = allExpenses.filter(x => x.packageId === p.id);
    const income = pBookings.reduce((s, b) => s + (Number(b.paidAmount) || 0), 0);
    const due = pBookings.reduce((s, b) => s + (Number(b.dueAmount) || 0), 0);
    const expense = pExpenses.reduce((s, x) => s + (Number(x.amount) || 0), 0);
    return {
      packageId: p.id, spotName: p.spotName,
      totalIncome: income, totalExpense: expense,
      netProfit: income - expense, totalDue: due, bookings: pBookings.length
    };
  });

  return {
    totalPackages: packages.length, totalBookings: allBookings.length,
    totalIncome, totalDue, totalExpense, netProfit: totalIncome - totalExpense,
    perPackage
  };
}
