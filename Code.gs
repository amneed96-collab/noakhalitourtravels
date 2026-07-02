/**
 * ============================================================
 *  ট্যুর ম্যানেজমেন্ট সিস্টেম - Google Apps Script ব্যাকএন্ড
 * ============================================================
 *
 * সেটআপ নির্দেশনা (README.md ফাইলে বিস্তারিত আছে):
 * 1. একটি নতুন Google Sheet তৈরি করুন
 * 2. Extensions > Apps Script এ যান
 * 3. এই কোডটি Code.gs এ পেস্ট করুন
 * 4. setupSheets() ফাংশনটি একবার রান করুন (শীট ও হেডার তৈরি হবে)
 * 5. Deploy > New deployment > Web app হিসেবে ডিপ্লয় করুন
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. যে URL পাবেন সেটি ফ্রন্টএন্ডের config.js এ বসান
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
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

function generateId(prefix) {
  return prefix + '-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const rows = data.slice(1);
  return rows
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
  // Packages
  const pkg = getSheet(SHEET_NAMES.PACKAGES);
  pkg.clear();
  pkg.appendRow([
    'id', 'spotName', 'startDate', 'endDate', 'totalDays',
    'delegateCount', 'delegateFee', 'vehicles', 'status',
    'createdAt', 'slug', 'coverNote'
  ]);

  // Bookings
  const bk = getSheet(SHEET_NAMES.BOOKINGS);
  bk.clear();
  bk.appendRow([
    'id', 'packageId', 'serial', 'date', 'name', 'mobile', 'address',
    'vehicleName', 'seatNo', 'delegateFee', 'paidAmount', 'dueAmount',
    'status', 'createdAt', 'confirmedBy', 'confirmedAt'
  ]);

  // Seats (প্রতিটি প্যাকেজের প্রতিটি গাড়ির প্রতিটি সিট)
  const st = getSheet(SHEET_NAMES.SEATS);
  st.clear();
  st.appendRow([
    'packageId', 'vehicleName', 'seatNo', 'status', 'bookingId', 'holdNote'
  ]);

  // Expenses
  const ex = getSheet(SHEET_NAMES.EXPENSES);
  ex.clear();
  ex.appendRow([
    'id', 'packageId', 'title', 'category', 'amount', 'note', 'date', 'addedBy', 'createdAt'
  ]);

  // Admins - ডিফল্ট এডমিন (অবশ্যই পরে পাসওয়ার্ড পরিবর্তন করুন)
  const ad = getSheet(SHEET_NAMES.ADMINS);
  ad.clear();
  ad.appendRow(['username', 'password', 'role', 'name', 'active']);
  ad.appendRow(['admin', 'admin123', 'admin', 'প্রধান এডমিন', true]);
  ad.appendRow(['assistant1', 'pass123', 'co-admin', 'সহকারী ১', true]);
  ad.appendRow(['assistant2', 'pass123', 'co-admin', 'সহকারী ২', true]);

  SpreadsheetApp.flush();
  return 'Setup সম্পন্ন হয়েছে!';
}

// ---------- ওয়েব API এন্ট্রি পয়েন্ট ----------

function doGet(e) {
  const action = e.parameter.action;
  try {
    switch (action) {
      case 'getPackages':
        return jsonResponse({ ok: true, data: getPackages() });
      case 'getPackage':
        return jsonResponse({ ok: true, data: getPackageDetail(e.parameter.id) });
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

// ---------- লগইন ----------

function login(username, password) {
  const admins = sheetToObjects(getSheet(SHEET_NAMES.ADMINS));
  const found = admins.find(
    a => a.username === username && String(a.password) === String(password) && a.active
  );
  if (!found) {
    return { ok: false, error: 'ইউজারনেম বা পাসওয়ার্ড ভুল, অথবা একাউন্ট নিষ্ক্রিয়।' };
  }
  return {
    ok: true,
    data: { username: found.username, role: found.role, name: found.name }
  };
}

function getAdmins() {
  return sheetToObjects(getSheet(SHEET_NAMES.ADMINS)).map(a => ({
    username: a.username, role: a.role, name: a.name, active: a.active
  }));
}

function manageAdmin(body) {
  // body: {mode: 'add'|'edit'|'delete', username, password, role, name, active}
  const sheet = getSheet(SHEET_NAMES.ADMINS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const usernameCol = headers.indexOf('username');

  if (body.mode === 'add') {
    sheet.appendRow([body.username, body.password, body.role, body.name, true]);
    return { ok: true };
  }

  for (let i = 1; i < data.length; i++) {
    if (data[i][usernameCol] === body.username) {
      if (body.mode === 'delete') {
        sheet.deleteRow(i + 1);
        return { ok: true };
      }
      if (body.mode === 'edit') {
        if (body.password) sheet.getRange(i + 1, headers.indexOf('password') + 1).setValue(body.password);
        if (body.role) sheet.getRange(i + 1, headers.indexOf('role') + 1).setValue(body.role);
        if (body.name) sheet.getRange(i + 1, headers.indexOf('name') + 1).setValue(body.name);
        if (body.active !== undefined) sheet.getRange(i + 1, headers.indexOf('active') + 1).setValue(body.active);
        return { ok: true };
      }
    }
  }
  return { ok: false, error: 'এডমিন পাওয়া যায়নি' };
}

// ---------- প্যাকেজ ----------

function slugify(text, id) {
  const base = text
    .toString()
    .trim()
    .replace(/[^\w\u0980-\u09FF\s-]/g, '')
    .replace(/\s+/g, '-')
    .toLowerCase();
  return base + '-' + id.split('-')[1];
}

function createPackage(body) {
  const sheet = getSheet(SHEET_NAMES.PACKAGES);
  const id = generateId('PKG');
  const slug = slugify(body.spotName, id);

  // vehicles: [{name: 'Bus-1', type:'বাস', seatCount: 40}, ...]
  const vehicles = body.vehicles || [];

  sheet.appendRow([
    id,
    body.spotName,
    body.startDate,
    body.endDate,
    body.totalDays,
    body.delegateCount,
    body.delegateFee,
    JSON.stringify(vehicles),
    'active',
    new Date().toISOString(),
    slug,
    body.coverNote || ''
  ]);

  // প্রতিটি গাড়ির জন্য সিট তৈরি করা
  const seatSheet = getSheet(SHEET_NAMES.SEATS);
  const rows = [];
  vehicles.forEach(v => {
    const rowLetters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    const seatsPerRow = 4; // A1-A4, B1-B4 প্যাটার্ন অনুযায়ী
    for (let i = 0; i < v.seatCount; i++) {
      const rowIdx = Math.floor(i / seatsPerRow);
      const seatIdx = (i % seatsPerRow) + 1;
      const seatNo = rowLetters[rowIdx] + seatIdx;
      rows.push([id, v.name, seatNo, 'available', '', '']);
    }
  });
  if (rows.length > 0) {
    seatSheet
      .getRange(seatSheet.getLastRow() + 1, 1, rows.length, 6)
      .setValues(rows);
  }

  return { id, slug };
}

function getPackages() {
  return sheetToObjects(getSheet(SHEET_NAMES.PACKAGES)).map(p => {
    p.vehicles = JSON.parse(p.vehicles || '[]');
    return p;
  });
}

function getPackageDetail(idOrSlug) {
  const packages = getPackages();
  const pkg = packages.find(p => p.id === idOrSlug || p.slug === idOrSlug);
  if (!pkg) return null;

  const seats = getSeatsForPackage(pkg.id);
  const totalSeats = seats.length;
  const bookedSeats = seats.filter(s => s.status === 'booked').length;
  const heldSeats = seats.filter(s => s.status === 'held').length;

  pkg.totalSeats = totalSeats;
  pkg.availableSeats = totalSeats - bookedSeats - heldSeats;
  return pkg;
}

// ---------- সিট ----------

function getSeatsForPackage(packageId) {
  return sheetToObjects(getSheet(SHEET_NAMES.SEATS)).filter(s => s.packageId === packageId);
}

function findSeatRow(sheet, packageId, vehicleName, seatNo) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  for (let i = 1; i < data.length; i++) {
    if (
      data[i][headers.indexOf('packageId')] === packageId &&
      data[i][headers.indexOf('vehicleName')] === vehicleName &&
      data[i][headers.indexOf('seatNo')] === seatNo
    ) {
      return { rowIndex: i + 1, headers };
    }
  }
  return null;
}

function holdSeats(packageId, vehicleName, seats, holdNote) {
  const sheet = getSheet(SHEET_NAMES.SEATS);
  const results = [];
  seats.forEach(seatNo => {
    const found = findSeatRow(sheet, packageId, vehicleName, seatNo);
    if (found) {
      const statusCol = found.headers.indexOf('status') + 1;
      const noteCol = found.headers.indexOf('holdNote') + 1;
      const currentStatus = sheet.getRange(found.rowIndex, statusCol).getValue();
      if (currentStatus === 'available') {
        sheet.getRange(found.rowIndex, statusCol).setValue('held');
        sheet.getRange(found.rowIndex, noteCol).setValue(holdNote || 'এডমিন হোল্ড');
        results.push({ seatNo, ok: true });
      } else {
        results.push({ seatNo, ok: false, reason: 'ইতিমধ্যে ' + currentStatus });
      }
    }
  });
  return { ok: true, results };
}

function releaseSeat(packageId, vehicleName, seatNo) {
  const sheet = getSheet(SHEET_NAMES.SEATS);
  const found = findSeatRow(sheet, packageId, vehicleName, seatNo);
  if (!found) return { ok: false, error: 'সিট পাওয়া যায়নি' };
  const statusCol = found.headers.indexOf('status') + 1;
  const bookingCol = found.headers.indexOf('bookingId') + 1;
  const noteCol = found.headers.indexOf('holdNote') + 1;
  sheet.getRange(found.rowIndex, statusCol).setValue('available');
  sheet.getRange(found.rowIndex, bookingCol).setValue('');
  sheet.getRange(found.rowIndex, noteCol).setValue('');
  return { ok: true };
}

// ---------- বুকিং ----------

function getBookings(packageId) {
  const all = sheetToObjects(getSheet(SHEET_NAMES.BOOKINGS));
  if (packageId) return all.filter(b => b.packageId === packageId);
  return all;
}

function createBooking(body) {
  const seatSheet = getSheet(SHEET_NAMES.SEATS);
  const found = findSeatRow(seatSheet, body.packageId, body.vehicleName, body.seatNo);

  if (!found) return { ok: false, error: 'সিট পাওয়া যায়নি' };

  const statusCol = found.headers.indexOf('status') + 1;
  const currentStatus = seatSheet.getRange(found.rowIndex, statusCol).getValue();
  if (currentStatus !== 'available') {
    return { ok: false, error: 'দুঃখিত, সিট ' + body.seatNo + ' ইতিমধ্যে বুক/হোল্ড করা আছে। অন্য সিট বেছে নিন।' };
  }

  const bookingSheet = getSheet(SHEET_NAMES.BOOKINGS);
  const id = generateId('BK');
  const existingCount = getBookings(body.packageId).length;
  const serial = existingCount + 1;

  const delegateFee = Number(body.delegateFee) || 0;
  const paidAmount = Number(body.paidAmount) || 0;
  const dueAmount = delegateFee - paidAmount;

  bookingSheet.appendRow([
    id,
    body.packageId,
    serial,
    body.date,
    body.name,
    body.mobile,
    body.address,
    body.vehicleName,
    body.seatNo,
    delegateFee,
    paidAmount,
    dueAmount,
    'pending', // pending -> confirmed
    new Date().toISOString(),
    '',
    ''
  ]);

  // সিট 'held' করে দেওয়া হলো (কনফার্ম না হওয়া পর্যন্ত)
  const bookingCol = found.headers.indexOf('bookingId') + 1;
  seatSheet.getRange(found.rowIndex, statusCol).setValue('held');
  seatSheet.getRange(found.rowIndex, bookingCol).setValue(id);

  return { ok: true, data: { id, serial, dueAmount } };
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
      const vehicleName = data[i][headers.indexOf('vehicleName')];
      const seatNo = data[i][headers.indexOf('seatNo')];

      const seatSheet = getSheet(SHEET_NAMES.SEATS);
      const found = findSeatRow(seatSheet, packageId, vehicleName, seatNo);
      if (found) {
        seatSheet
          .getRange(found.rowIndex, found.headers.indexOf('status') + 1)
          .setValue('booked');
      }
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
      const vehicleName = data[i][headers.indexOf('vehicleName')];
      const seatNo = data[i][headers.indexOf('seatNo')];
      releaseSeat(packageId, vehicleName, seatNo);
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
      return { ok: true, data: { dueAmount: fee - paid } };
    }
  }
  return { ok: false, error: 'বুকিং পাওয়া যায়নি' };
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
    id,
    body.packageId,
    body.title,
    body.category || 'অন্যান্য',
    Number(body.amount) || 0,
    body.note || '',
    body.date || new Date().toISOString().split('T')[0],
    body.addedBy || '',
    new Date().toISOString()
  ]);
  return { id };
}

function deleteExpense(id) {
  const sheet = getSheet(SHEET_NAMES.EXPENSES);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
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
    totalExpected,
    totalIncome,
    totalDue,
    totalExpense,
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
    const r = getFinanceReport(p.id);
    return {
      packageId: p.id,
      spotName: p.spotName,
      totalIncome: r.totalIncome,
      totalExpense: r.totalExpense,
      netProfit: r.netProfit,
      totalDue: r.totalDue,
      bookings: r.totalBookings
    };
  });

  return {
    totalPackages: packages.length,
    totalBookings: allBookings.length,
    totalIncome,
    totalDue,
    totalExpense,
    netProfit: totalIncome - totalExpense,
    perPackage
  };
}
