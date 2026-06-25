// =====================================================
// Code.gs  —  SM Schedule System (Branch-first v2)
// GAS Backend: branch schedule + 8 slots + custom-time
//              validation + export modes
// =====================================================

const SHEETS = {
  SCHEDULES:  'Schedules',
  EMPLOYEES:  'Employees',
  BRANCHES:   'Branches',
  HOLIDAYS:   'Holidays',
  SHIFT_TYPES:'ShiftTypes',
  EVENTS:     'Events',
  CONFIG:     'Config'
};

const SLOTS_PER_BRANCH = 8;

// ── Entry Point ──────────────────────────────────────
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบตารางรอบ SM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── บังคับขอสิทธิ์ทั้งหมด (รันครั้งเดียวจาก Editor → Run → forceAuthorize) ──
// จะเด้งหน้าขออนุญาต Drive/Sheets ให้กด Allow แล้วค่อย Deploy เวอร์ชันใหม่
function forceAuthorize() {
  SpreadsheetApp.getActiveSpreadsheet();                 // สิทธิ์ Sheets
  const tmp = SpreadsheetApp.create('__auth_check__');   // สร้างไฟล์ (ต้องใช้ Drive)
  const id  = tmp.getId();
  const f   = DriveApp.getFileById(id);                  // อ่านไฟล์ (สิทธิ์ Drive)
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // แชร์ไฟล์
  DriveApp.getFileById(id).setTrashed(true);             // ลบไฟล์ทดสอบทิ้ง
  const msg = '✅ อนุญาตสิทธิ์ครบแล้ว — ขั้นตอนถัดไป: Deploy > Manage deployments > Edit > New version';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
  return msg;
}

// ── Helpers ──────────────────────────────────────────
function getSS()        { return SpreadsheetApp.getActiveSpreadsheet(); }
function getSheet(name) { return getSS().getSheetByName(name); }

function sheetToObjects(name) {
  const sh = getSheet(name);
  if (!sh) return [];
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];
  const headers = vals[0];
  return vals.slice(1)
    .filter(r => r.some(c => c !== '' && c !== null))
    .map(r => Object.fromEntries(headers.map((h, i) => [h, r[i]])));
}

function getConfigValue(key, fallback) {
  const row = sheetToObjects(SHEETS.CONFIG).find(c => c.key === key);
  return row ? row.value : fallback;
}

function getRequiredHours() {
  return Number(getConfigValue('requiredShiftHours', 9)) || 9;
}

// คำนวณจำนวนชั่วโมงของกะแบบ "HH:MM-HH:MM" หรือ "HH.MM-HH.MM"
// คืนค่า null ถ้าไม่ใช่รูปแบบเวลา (เช่น OFF/PH/LEAVE)
// แยกค่า OT แบบ "เวลา|สาเหตุ"
function otSplit(v) {
  if (!v) return { time: '', reason: '' };
  v = String(v);
  const i = v.indexOf('|');
  return i < 0 ? { time: v, reason: '' } : { time: v.slice(0, i), reason: v.slice(i + 1) };
}

function shiftSpanHours(code) {
  if (!code) return null;
  code = String(code).split('|')[0]; // ตัดสาเหตุ OT ออก
  const m = code.trim().match(/^(\d{1,2})[:.](\d{2})\s*-\s*(\d{1,2})[:.](\d{2})$/);
  if (!m) return null;
  let start = Number(m[1]) + Number(m[2]) / 60;
  let end   = Number(m[3]) + Number(m[4]) / 60;
  if (end <= start) end += 24; // ข้ามเที่ยงคืน
  return Math.round((end - start) * 100) / 100;
}

function isWorkingShift(code) {
  if (!code) return false;
  return !['OFF', 'PH', 'LEAVE', 'SICK', 'COMP'].includes(code);
}

function makeSlotId(branchCode, slot) {
  return `${branchCode}-S${slot}`;
}

// ── Master Data ──────────────────────────────────────

function getAMList() {
  const amSet = new Set(
    sheetToObjects(SHEETS.BRANCHES).map(b => b.amCode).filter(Boolean)
  );
  return [...amSet].sort();
}

// คืนรายการสาขาของ AM พร้อมข้อมูลรอบ DC/FDC
function getBranchList(amCode) {
  return sheetToObjects(SHEETS.BRANCHES)
    .filter(b => (!amCode || b.amCode === amCode) &&
                 b.isActive !== false && b.isActive !== 'FALSE')
    .map(b => ({
      branchCode:  String(b.branchCode),
      branchName:  b.branchName,
      workHours:   b.workHours,
      dcSchedule:  b.dcSchedule  || '',
      fdcSchedule: b.fdcSchedule || '',
      smName:      b.smName      || '',
      amCode:      b.amCode,
      rmName:      b.rmName || '',
      dmName:      b.dmName || ''
    }));
}

function getShiftTypes() {
  return sheetToObjects(SHEETS.SHIFT_TYPES);
}

// ── Admin ─────────────────────────────────────────────
function verifyAdminPin(pin) {
  const real = String(getConfigValue('adminPin', '1234'));
  return { success: String(pin) === real };
}

// ── Dashboard (สรุปทุกสาขาในเขต) ──────────────────────
function getDashboard(amCode, year, month) {
  year  = Number(year);
  month = Number(month);
  const branches = getBranchList(amCode);
  const daysInMonth = new Date(year, month, 0).getDate();

  const holidayCount = sheetToObjects(SHEETS.HOLIDAYS).filter(h => {
    if (!h.date) return false;
    const dt = new Date(h.date);
    return dt.getFullYear() === year && (dt.getMonth() + 1) === month;
  }).length;

  const rows = branches.map(b => {
    const data = getBranchSchedule(b.branchCode, year, month);
    let staff = 0, totalWork = 0, totalHours = 0, blanks = 0;
    let anyStatus = 'draft';

    data.slots.forEach(s => {
      const hasName = !!(s.name || s.empId);
      if (!hasName) return;
      staff++;
      const sum = data.summaries[s.slotId] || {};
      totalWork  += sum.totalWork  || 0;
      totalHours += sum.totalHours || 0;
      if (sum.status && sum.status !== 'draft') anyStatus = sum.status;
      // นับช่องว่าง (วันที่ยังไม่กรอก) ของพนักงานคนนี้
      const sh = data.shifts[s.slotId] || {};
      for (let d = 1; d <= daysInMonth; d++) if (!sh[d]) blanks++;
    });

    return {
      branchCode:  b.branchCode,
      branchName:  b.branchName,
      dcSchedule:  b.dcSchedule,
      fdcSchedule: b.fdcSchedule,
      staff, totalWork,
      totalHours: Math.round(totalHours * 10) / 10,
      blanks, status: anyStatus
    };
  });

  return {
    amCode, year, month, daysInMonth, holidayCount,
    branchCount: branches.length,
    totalStaff: rows.reduce((a, r) => a + r.staff, 0),
    rows
  };
}

// ── Branch Schedule (main load — per branch, 8 slots) ─

function getBranchSchedule(branchCode, year, month) {
  branchCode = String(branchCode);
  year  = Number(year);
  month = Number(month);

  const branch = sheetToObjects(SHEETS.BRANCHES)
    .find(b => String(b.branchCode) === branchCode) || {};

  // พนักงานของสาขานี้ (ที่มีอยู่)
  const existing = sheetToObjects(SHEETS.EMPLOYEES)
    .filter(e => String(e.branchCode) === branchCode)
    .reduce((acc, e) => { acc[Number(e.slot)] = e; return acc; }, {});

  // สร้าง 8 slots — เติม placeholder ว่างถ้ายังไม่มี
  const slots = [];
  for (let s = 1; s <= SLOTS_PER_BRANCH; s++) {
    const e = existing[s];
    slots.push({
      slotId:     makeSlotId(branchCode, s),
      slot:       s,
      branchCode: branchCode,
      empId:      e ? String(e.empId || '') : '',
      name:       e ? (e.name || '')        : '',
      nickname:   e ? (e.nickname || '')    : '',
      phone:      e ? String(e.phone || '') : '',
      empType:    e ? (e.empType || '')     : '',
      dayOff:     e ? (e.dayOff || '')      : ''
    });
  }

  // วันหยุดในเดือนนี้
  const holidayDays = sheetToObjects(SHEETS.HOLIDAYS)
    .filter(h => {
      if (!h.date) return false;
      const d = new Date(h.date);
      return d.getFullYear() === year && (d.getMonth() + 1) === month;
    })
    .map(h => ({ day: new Date(h.date).getDate(), name: h.name, isNational: h.isNational }));

  // วันสำคัญในเดือนนี้
  const monthEvents = sheetToObjects(SHEETS.EVENTS)
    .filter(ev => {
      if (!ev.date) return false;
      const d = new Date(ev.date);
      return d.getFullYear() === year && (d.getMonth() + 1) === month;
    })
    .map(ev => ({ day: new Date(ev.date).getDate(), title: ev.title, detail: ev.detail || '' }));

  // ตารางรอบของสาขานี้
  const scheduleRows = sheetToObjects(SHEETS.SCHEDULES).filter(
    r => String(r.branchCode) === branchCode &&
         Number(r.year) === year && Number(r.month) === month
  );

  const shifts = {};      // slotId → {day: shift}
  const otShifts = {};    // slotId → {day: ot time}
  const summaries = {};   // slotId → totals
  scheduleRows.forEach(row => {
    const sid = String(row.slotId);
    const map = {}, otMap = {};
    for (let d = 1; d <= 31; d++) {
      const v = row['d' + d];
      if (v !== undefined && v !== null && v !== '') map[d] = String(v);
      const ov = row['ot' + d];
      if (ov !== undefined && ov !== null && ov !== '') otMap[d] = String(ov);
    }
    shifts[sid] = map;
    otShifts[sid] = otMap;
    summaries[sid] = {
      totalOff:   Number(row.totalOff)   || 0,
      totalAM:    Number(row.totalAM)    || 0,
      totalPM:    Number(row.totalPM)    || 0,
      totalNight: Number(row.totalNight) || 0,
      totalWork:  Number(row.totalWork)  || 0,
      totalHours: Number(row.totalHours) || 0,
      totalOT:    Number(row.totalOT)    || 0,
      status:     row.status || 'draft',
      note:       row.note   || '',
      otNote:     row.otNote || ''
    };
  });

  return {
    branchCode,
    branchName:  branch.branchName  || '',
    workHours:   branch.workHours   || '',
    dcSchedule:  branch.dcSchedule  || '',
    fdcSchedule: branch.fdcSchedule || '',
    smName:      branch.smName      || '',
    amCode:      branch.amCode      || '',
    rmName:      branch.rmName      || '',
    dmName:      branch.dmName      || '',
    year, month,
    slots,
    shifts,
    otShifts,
    summaries,
    holidayDays,
    monthEvents,
    shiftTypes:  sheetToObjects(SHEETS.SHIFT_TYPES),
    requiredHours: getRequiredHours()
  };
}

// ── Save / Edit / Delete Employee Slot ────────────────

// emp: {slotId, branchCode, slot, empId, name, empType, dayOff}
function saveEmployeeSlot(emp) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(20000)) return { success: false, error: 'ระบบกำลังบันทึกข้อมูลอื่นอยู่ กรุณาลองใหม่' };
    _dedupeSheet(SHEETS.EMPLOYEES, ['slotId']);
    _saveEmployeesCore([emp]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ลบทั้งข้อมูลพนักงาน และ ตารางรอบ ของ slot นั้น
function deleteEmployeeSlot(slotId) {
  try {
    // ลบจาก Employees
    const empSh   = getSheet(SHEETS.EMPLOYEES);
    const empVals = empSh.getDataRange().getValues();
    const empCol  = empVals[0].indexOf('slotId');
    for (let i = empVals.length - 1; i >= 1; i--) {
      if (String(empVals[i][empCol]) === String(slotId)) empSh.deleteRow(i + 1);
    }
    // ลบจาก Schedules
    const schSh   = getSheet(SHEETS.SCHEDULES);
    const schVals = schSh.getDataRange().getValues();
    const schCol  = schVals[0].indexOf('slotId');
    for (let i = schVals.length - 1; i >= 1; i--) {
      if (String(schVals[i][schCol]) === String(slotId)) schSh.deleteRow(i + 1);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ── Save ALL (transactional: พนักงาน + ตาราง ใน call เดียว + Lock) ──
// employees: [{slotId, branchCode, slot, empId, name, nickname, phone, dayOff, empType}]
// changes:   [{slotId, empId, name, day, shift, kind?}]
function saveAll(branchCode, year, month, employees, changes) {
  const lock = LockService.getScriptLock();
  try {
    if (!lock.tryLock(30000)) {
      return { success: false, error: 'ระบบกำลังบันทึกข้อมูลอื่นอยู่ กรุณาลองใหม่อีกครั้ง' };
    }

    branchCode = String(branchCode);
    year  = Number(year);
    month = Number(month);
    const required = getRequiredHours();

    // ── ตรวจชั่วโมงกะปกติ (ไม่รวม OT) ──
    const invalid = [];
    (changes || []).forEach(c => {
      if (c.kind !== 'ot' && isWorkingShift(c.shift)) {
        const hrs = shiftSpanHours(c.shift);
        if (hrs !== null && hrs < required) {
          invalid.push({ slotId: c.slotId, day: c.day, shift: c.shift, hours: hrs });
        }
      }
    });
    if (invalid.length) {
      return { success: false, error: `มี ${invalid.length} ช่องที่ชั่วโมงทำงานไม่ครบ ${required} ชม.`, invalid };
    }

    // ป้องกัน/ล้างแถวซ้ำก่อนเสมอ
    _dedupeSheet(SHEETS.EMPLOYEES, ['slotId']);
    _dedupeSheet(SHEETS.SCHEDULES, ['branchCode', 'year', 'month', 'slotId']);

    _saveEmployeesCore(employees || []);
    _saveShiftsCore(branchCode, year, month, changes || []);

    return { success: true };
  } catch (e) {
    console.error(e);
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// (คงไว้เพื่อ backward-compat — เรียก saveAll โดยไม่มีพนักงาน)
function saveBranchShifts(branchCode, year, month, changes) {
  return saveAll(branchCode, year, month, [], changes);
}

// บันทึกพนักงาน (upsert ตาม slotId — ถือว่า dedupe แล้ว มีแถวเดียวต่อ slot)
function _saveEmployeesCore(employees) {
  if (!employees.length) return;
  const sh      = getSheet(SHEETS.EMPLOYEES);
  const vals    = sh.getDataRange().getValues();
  const headers = vals[0];
  const slotCol = headers.indexOf('slotId');
  const rowBySlot = {};
  for (let i = 1; i < vals.length; i++) rowBySlot[String(vals[i][slotCol])] = i;

  const branchMap = {};
  sheetToObjects(SHEETS.BRANCHES).forEach(b => { branchMap[String(b.branchCode)] = b; });

  employees.forEach(emp => {
    const branch = branchMap[String(emp.branchCode)] || {};
    const record = {
      slotId:     emp.slotId,
      branchCode: emp.branchCode,
      slot:       emp.slot,
      empId:      emp.empId || '',
      name:       emp.name  || '',
      nickname:   emp.nickname || '',
      phone:      emp.phone || '',
      branchName: branch.branchName || '',
      empType:    emp.empType || '',
      dayOff:     emp.dayOff  || '',
      amCode:     branch.amCode || '',
      isActive:   true
    };
    const rowData = headers.map(h => (record[h] !== undefined ? record[h] : ''));
    const existing = rowBySlot[String(emp.slotId)];
    if (existing !== undefined) {
      sh.getRange(existing + 1, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sh.appendRow(rowData);
    }
  });
}

// บันทึกกะ/OT (upsert ตาม branch/year/month/slotId — dedupe แล้ว)
function _saveShiftsCore(branchCode, year, month, changes) {
  if (!changes.length) return;
  const sh      = getSheet(SHEETS.SCHEDULES);
  const allVals = sh.getDataRange().getValues();
  const headers = allVals[0];
  const colIdx  = name => headers.indexOf(name);

  const rowKey = {};
  for (let i = 1; i < allVals.length; i++) {
    if (String(allVals[i][colIdx('branchCode')]) === branchCode &&
        Number(allVals[i][colIdx('year')])  === year &&
        Number(allVals[i][colIdx('month')]) === month) {
      rowKey[String(allVals[i][colIdx('slotId')])] = i;
    }
  }

  const bySlot = {};
  changes.forEach(c => {
    const key = String(c.slotId);
    if (!bySlot[key]) bySlot[key] = { meta: c, days: [] };
    bySlot[key].days.push(c);
  });

  Object.entries(bySlot).forEach(([slotId, { meta, days }]) => {
    let rowIdx = rowKey[slotId] !== undefined ? rowKey[slotId] : -1;
    const colFor = c => (c.kind === 'ot' ? 'ot' : 'd') + c.day;

    if (rowIdx === -1) {
      const newRow = new Array(headers.length).fill('');
      newRow[colIdx('branchCode')] = branchCode;
      newRow[colIdx('year')]       = year;
      newRow[colIdx('month')]      = month;
      newRow[colIdx('slotId')]     = slotId;
      newRow[colIdx('empId')]      = meta.empId || '';
      newRow[colIdx('name')]       = meta.name  || '';
      newRow[colIdx('status')]     = 'draft';
      days.forEach(c => { newRow[colIdx(colFor(c))] = c.shift; });
      sh.appendRow(newRow);
      rowIdx = sh.getLastRow() - 1;
    } else {
      if (meta.empId !== undefined) sh.getRange(rowIdx + 1, colIdx('empId') + 1).setValue(meta.empId || '');
      if (meta.name  !== undefined) sh.getRange(rowIdx + 1, colIdx('name')  + 1).setValue(meta.name  || '');
      days.forEach(c => { sh.getRange(rowIdx + 1, colIdx(colFor(c)) + 1).setValue(c.shift); });
    }

    const freshRow = sh.getRange(rowIdx + 1, 1, 1, headers.length).getValues()[0];
    _updateSummary(sh, headers, freshRow, rowIdx + 1, year, month);
  });
}

// ── ล้างแถวซ้ำ (merge เซลล์ที่มีค่าเข้าด้วยกัน) ──
function _dedupeSheet(sheetName, keyColNames) {
  const sh = getSheet(sheetName);
  if (!sh) return 0;
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return 0;
  const headers = vals[0];
  const idx = keyColNames.map(n => headers.indexOf(n));
  if (idx.some(i => i < 0)) return 0;

  const out = [headers];
  const keepByKey = {};
  let removed = 0;

  for (let i = 1; i < vals.length; i++) {
    const blank = idx.every(c => vals[i][c] === '' || vals[i][c] === null);
    if (blank) continue;
    const key = idx.map(c => String(vals[i][c])).join('||');
    if (keepByKey[key] !== undefined) {
      // merge: เติมค่าที่แถวเก็บยังว่างจากแถวซ้ำ
      const krow = out[keepByKey[key]];
      for (let c = 0; c < headers.length; c++) {
        if ((krow[c] === '' || krow[c] === null) && vals[i][c] !== '' && vals[i][c] !== null) {
          krow[c] = vals[i][c];
        }
      }
      removed++;
    } else {
      out.push(vals[i].slice());
      keepByKey[key] = out.length - 1;
    }
  }

  if (removed > 0) {
    sh.getRange(1, 1, sh.getMaxRows(), headers.length).clearContent();
    sh.getRange(1, 1, out.length, headers.length).setValues(out);
  }
  return removed;
}

// เรียกจาก Editor เพื่อล้างแถวซ้ำที่มีอยู่เดิม
function cleanupDuplicates() {
  const lock = LockService.getScriptLock();
  lock.tryLock(30000);
  try {
    const e = _dedupeSheet(SHEETS.EMPLOYEES, ['slotId']);
    const s = _dedupeSheet(SHEETS.SCHEDULES, ['branchCode', 'year', 'month', 'slotId']);
    const msg = `ล้างแถวซ้ำเสร็จ — Employees: ${e} แถว, Schedules: ${s} แถว`;
    try { SpreadsheetApp.getUi().alert(msg); } catch (er) { Logger.log(msg); }
    return { success: true, employeesRemoved: e, schedulesRemoved: s };
  } finally {
    lock.releaseLock();
  }
}

function _updateSummary(sh, headers, row, sheetRow, year, month) {
  const col       = name => headers.indexOf(name) + 1;
  const daysCount = new Date(year, month, 0).getDate();

  let offC = 0, amC = 0, pmC = 0, nightC = 0, workC = 0, hoursC = 0, otC = 0;

  for (let d = 1; d <= daysCount; d++) {
    const v = String(row[headers.indexOf('d' + d)] || '').trim();
    if (v) {
      if (!isWorkingShift(v)) { offC++; }
      else {
        workC++;
        const hrs = shiftSpanHours(v);
        if (hrs !== null) hoursC += hrs;
        const h = parseInt(v.split(/[:.]/)[0]);
        if      (h >= 5  && h <= 10) amC++;
        else if (h >= 13 && h <= 15) pmC++;
        else if (h >= 15)             nightC++;
      }
    }
    // โอที (รวมชั่วโมงไม่จำกัดความยาว)
    const ot = String(row[headers.indexOf('ot' + d)] || '').trim();
    if (ot) {
      const oh = shiftSpanHours(ot);
      otC += (oh !== null ? oh : 0);
    }
  }

  if (col('totalOff')   > 0) sh.getRange(sheetRow, col('totalOff')).setValue(offC);
  if (col('totalAM')    > 0) sh.getRange(sheetRow, col('totalAM')).setValue(amC);
  if (col('totalPM')    > 0) sh.getRange(sheetRow, col('totalPM')).setValue(pmC);
  if (col('totalNight') > 0) sh.getRange(sheetRow, col('totalNight')).setValue(nightC);
  if (col('totalWork')  > 0) sh.getRange(sheetRow, col('totalWork')).setValue(workC);
  if (col('totalHours') > 0) sh.getRange(sheetRow, col('totalHours')).setValue(hoursC);
  if (col('totalOT')    > 0) sh.getRange(sheetRow, col('totalOT')).setValue(Math.round(otC * 10) / 10);
}

// ── Approval Status ───────────────────────────────────

function updateScheduleStatus(branchCode, year, month, slotId, status) {
  try {
    const sh      = getSheet(SHEETS.SCHEDULES);
    const allVals = sh.getDataRange().getValues();
    const headers = allVals[0];
    const col     = name => headers.indexOf(name);
    for (let i = 1; i < allVals.length; i++) {
      if (String(allVals[i][col('branchCode')]) === String(branchCode) &&
          Number(allVals[i][col('year')])   === Number(year)  &&
          Number(allVals[i][col('month')])  === Number(month) &&
          String(allVals[i][col('slotId')]) === String(slotId)) {
        sh.getRange(i + 1, col('status') + 1).setValue(status);
        return { success: true };
      }
    }
    return { success: false, error: 'ไม่พบแถวนี้' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ── Holiday CRUD ──────────────────────────────────────

function getHolidays(year) {
  return sheetToObjects(SHEETS.HOLIDAYS)
    .filter(h => h.date && new Date(h.date).getFullYear() === Number(year))
    .map(h => ({
      date:       Utilities.formatDate(new Date(h.date), 'Asia/Bangkok', 'yyyy-MM-dd'),
      name:       h.name,
      isNational: h.isNational,
      note:       h.note || ''
    }))
    .sort((a, b) => a.date.localeCompare(b.date));
}

function saveHoliday(holiday) {
  try {
    const sh      = getSheet(SHEETS.HOLIDAYS);
    const allVals = sh.getDataRange().getValues();
    const headers = allVals[0];
    const col     = name => headers.indexOf(name);
    const target  = new Date(holiday.date);

    let rowIdx = -1;
    for (let i = 1; i < allVals.length; i++) {
      if (allVals[i][col('date')] &&
          new Date(allVals[i][col('date')]).toDateString() === target.toDateString()) {
        rowIdx = i; break;
      }
    }
    const rowData = [
      target, holiday.name,
      holiday.isNational === true || holiday.isNational === 'true',
      holiday.branchCodes || '', holiday.note || ''
    ];
    if (rowIdx === -1) sh.appendRow(rowData);
    else sh.getRange(rowIdx + 1, 1, 1, rowData.length).setValues([rowData]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteHoliday(dateStr) {
  try {
    const sh      = getSheet(SHEETS.HOLIDAYS);
    const allVals = sh.getDataRange().getValues();
    const dcol    = allVals[0].indexOf('date');
    const target  = new Date(dateStr);
    for (let i = allVals.length - 1; i >= 1; i--) {
      if (allVals[i][dcol] &&
          new Date(allVals[i][dcol]).toDateString() === target.toDateString()) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'ไม่พบวันหยุดนี้' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ── Employee list (for employees page) ────────────────

function getEmployees(amCode) {
  return sheetToObjects(SHEETS.EMPLOYEES)
    .filter(e => (!amCode || e.amCode === amCode) && (e.name || e.empId));
}

// ── วันสำคัญ (Events) + แจ้งเตือน ──────────────────────

function getEvents(year) {
  return sheetToObjects(SHEETS.EVENTS)
    .filter(ev => ev.date && (!year || new Date(ev.date).getFullYear() === Number(year)))
    .map(ev => ({
      date:   Utilities.formatDate(new Date(ev.date), 'Asia/Bangkok', 'yyyy-MM-dd'),
      title:  ev.title,
      detail: ev.detail || ''
    }))
    .sort((a, b) => a.date.localeCompare(b.date));
}

function saveEvent(ev) {
  try {
    const sh      = getSheet(SHEETS.EVENTS);
    const allVals = sh.getDataRange().getValues();
    const headers = allVals[0];
    const c       = name => headers.indexOf(name);
    const target  = new Date(ev.date);

    let rowIdx = -1;
    for (let i = 1; i < allVals.length; i++) {
      if (allVals[i][c('date')] &&
          new Date(allVals[i][c('date')]).toDateString() === target.toDateString() &&
          String(allVals[i][c('title')]) === String(ev.title)) {
        rowIdx = i; break;
      }
    }
    const rowData = [target, ev.title, ev.detail || '', true];
    if (rowIdx === -1) sh.appendRow(rowData);
    else sh.getRange(rowIdx + 1, 1, 1, rowData.length).setValues([rowData]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deleteEvent(dateStr, title) {
  try {
    const sh      = getSheet(SHEETS.EVENTS);
    const allVals = sh.getDataRange().getValues();
    const dc      = allVals[0].indexOf('date');
    const tc      = allVals[0].indexOf('title');
    const target  = new Date(dateStr);
    for (let i = allVals.length - 1; i >= 1; i--) {
      if (allVals[i][dc] && new Date(allVals[i][dc]).toDateString() === target.toDateString() &&
          String(allVals[i][tc]) === String(title)) {
        sh.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'ไม่พบวันสำคัญนี้' };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// วันสำคัญที่กำลังจะมาถึงภายใน N วัน (สำหรับแจ้งเตือน)
function getUpcomingEvents(days) {
  days = Number(days) || 45;
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const limit = new Date(today.getTime() + days * 86400000);
  return sheetToObjects(SHEETS.EVENTS)
    .filter(ev => ev.date)
    .map(ev => {
      const d = new Date(ev.date); d.setHours(0, 0, 0, 0);
      return {
        date:     Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd'),
        title:    ev.title,
        detail:   ev.detail || '',
        daysLeft: Math.round((d.getTime() - today.getTime()) / 86400000)
      };
    })
    .filter(ev => ev.daysLeft >= 0 && new Date(ev.date) <= limit)
    .sort((a, b) => a.daysLeft - b.daysLeft);
}

// ── Export (รายงานมืออาชีพ) ───────────────────────────
// mode: 'branch' = สาขาเดียว, 'sm' = เฉพาะ SM ทุกสาขา, 'all' = ทุกสาขาใน AM
function exportToExcel(amCode, year, month, mode, branchCode) {
  try {
    year  = Number(year);
    month = Number(month);
    mode  = mode || 'branch';

    const THAI_MONTHS = ['','มกราคม','กุมภาพันธ์','มีนาคม','เมษายน','พฤษภาคม','มิถุนายน',
                         'กรกฎาคม','สิงหาคม','กันยายน','ตุลาคม','พฤศจิกายน','ธันวาคม'];
    const DAY_TH = ['อา','จ','อ','พ','พฤ','ศ','ส'];
    const daysInMonth = new Date(year, month, 0).getDate();

    let branches = getBranchList(amCode);
    let titleSuffix = '';
    if (mode === 'branch' && branchCode) {
      branches = branches.filter(b => String(b.branchCode) === String(branchCode));
      titleSuffix = '_' + branchCode;
    } else if (mode === 'sm') {
      titleSuffix = '_SM';
    } else {
      titleSuffix = '_ALL';
    }

    const fileName = `ตารางรอบ_${amCode}${titleSuffix}_${THAI_MONTHS[month]}${year + 543}`;
    const exportSS   = SpreadsheetApp.create(fileName);
    const firstSheet = exportSS.getActiveSheet();

    // สร้าง entry 1 รายการต่อพนักงาน 1 คน (ผูกข้อมูลสาขาของตัวเอง)
    const makeEntry = (branch, data, slot) => ({
      branch,
      slot,
      shifts: data.shifts[slot.slotId] || {},
      ot:     (data.otShifts && data.otShifts[slot.slotId]) || {},
      sum:    data.summaries[slot.slotId] || {},
      holidayDays: data.holidayDays || []
    });

    if (mode === 'sm') {
      // ── รวม SM ทุกสาขาเป็นชีตเดียว เรียงลำดับ 1,2,3,4 ──
      const entries = [];
      branches.forEach(branch => {
        const data = getBranchSchedule(branch.branchCode, year, month);
        data.slots.forEach(s => {
          if ((s.empType || '').toUpperCase() !== 'SM') return;
          if (!(s.name || s.empId)) return;
          entries.push(makeEntry(branch, data, s));
        });
      });
      if (!entries.length) {
        return { success: false, error: 'ไม่พบข้อมูล SM ที่จะ export' };
      }
      firstSheet.setName('SM');
      _buildExportSheet(firstSheet, entries, { mode: 'sm', branch: null }, year, month, daysInMonth, THAI_MONTHS, DAY_TH);
    } else {
      // ── สาขาเดียว / ทุกสาขา (แยกชีตละสาขา) ──
      let firstUsed = false, sheetCount = 0;
      branches.forEach(branch => {
        const data = getBranchSchedule(branch.branchCode, year, month);
        const entries = [];
        data.slots.forEach(s => {
          if (!(s.name || s.empId ||
                (data.shifts[s.slotId] && Object.keys(data.shifts[s.slotId]).length))) return;
          entries.push(makeEntry(branch, data, s));
        });
        if (!entries.length) return;

        const sheet = firstUsed
          ? exportSS.insertSheet(String(branch.branchCode))
          : (firstSheet.setName(String(branch.branchCode)), firstSheet);
        firstUsed = true;
        sheetCount++;
        _buildExportSheet(sheet, entries, { mode, branch }, year, month, daysInMonth, THAI_MONTHS, DAY_TH);
      });
      if (!sheetCount) {
        return { success: false, error: 'ไม่พบข้อมูลที่จะ export ตามเงื่อนไขที่เลือก' };
      }
    }

    SpreadsheetApp.flush();

    // ── ให้เบราว์เซอร์ดาวน์โหลด .xlsx โดยตรง (ไม่ใช้ UrlFetchApp) ──
    const ssId = exportSS.getId();
    // แชร์ + เก็บกวาด (ถ้ายังไม่มีสิทธิ์ Drive ก็ข้ามไป — ไฟล์ของเจ้าของยังโหลดได้)
    try {
      const file = DriveApp.getFileById(ssId);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      _cleanupOldExports(ssId);
    } catch (e) { console.warn('Drive sharing skipped: ' + e); }

    return {
      success: true,
      fileName: fileName + '.xlsx',
      fileId: ssId,
      downloadUrl: 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=xlsx'
    };
  } catch (e) {
    console.error(e);
    return { success: false, error: e.toString() };
  }
}

// ลบไฟล์ export ทิ้งทันทีหลังดาวน์โหลดเสร็จ (เรียกจาก frontend)
function deleteExportFile(fileId) {
  try { DriveApp.getFileById(fileId).setTrashed(true); return { success: true }; }
  catch (e) { return { success: false, error: e.toString() }; }
}

// เก็บไฟล์ export ไว้ในโฟลเดอร์เดียว + ลบไฟล์เก่ากว่า 2 ชม. ทิ้ง (กัน Drive รก)
function _cleanupOldExports(keepId) {
  try {
    const FOLDER = 'SM_Schedule_Exports';
    const it = DriveApp.getFoldersByName(FOLDER);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);

    const f = DriveApp.getFileById(keepId);
    folder.addFile(f);
    try { DriveApp.getRootFolder().removeFile(f); } catch (e) {}

    const cutoff = Date.now() - 2 * 3600 * 1000;
    const files = folder.getFiles();
    while (files.hasNext()) {
      const x = files.next();
      if (x.getId() !== keepId && x.getDateCreated().getTime() < cutoff) {
        try { x.setTrashed(true); } catch (e) {}
      }
    }
  } catch (e) {}
}

// สีประจำรหัสกะพิเศษ (มาร์คสีต่างกันในรีพอร์ต)
const EXPORT_SHIFT_COLORS = {
  OFF:   { bg: '#e0e0e0', fc: '#424242' },  // วันหยุดประจำวัน
  PH:    { bg: '#ffcdd2', fc: '#b71c1c' },  // หยุดนักขัตฤกษ์
  LEAVE: { bg: '#c8e6c9', fc: '#1b5e20' },  // ลาพักร้อน
  SICK:  { bg: '#fff9c4', fc: '#f57f17' },  // ลาป่วย
  COMP:  { bg: '#f8bbd0', fc: '#880e4f' }   // พักก่อน
};
// แสดงในเซลล์เป็นชื่อภาษาไทย (ใช้ code เดียวกับชื่อ)
const CODE_LABEL = {
  OFF:   'วันหยุดประจำวัน',
  PH:    'หยุดนักขัตฤกษ์',
  LEAVE: 'ลาพักร้อน',
  SICK:  'ลาป่วย',
  COMP:  'ชดเชยขัตฤกษ์'
};
function _shiftCellColor(code) {
  if (EXPORT_SHIFT_COLORS[code]) return EXPORT_SHIFT_COLORS[code];
  const h = parseInt(String(code).split(/[:.]/)[0]);
  if (h >= 5  && h <= 10) return { bg: '#fff8e1', fc: '#333' }; // เช้า
  if (h >= 13 && h <= 15) return { bg: '#e3f2fd', fc: '#333' }; // บ่าย
  if (h >= 15)            return { bg: '#ede7f6', fc: '#333' }; // ดึก
  return null;
}

// สร้างชีตแบบมืออาชีพจาก entries (1 entry = พนักงาน 1 คน ผูกสาขาของตัวเอง)
// info = { mode:'branch'|'sm'|'all', branch:<branch|null> }
function _buildExportSheet(sheet, entries, info, year, month, daysInMonth, THAI_MONTHS, DAY_TH) {
  const mode        = info.mode;
  const branch      = info.branch;                 // null เมื่อ mode==='sm'
  const numbered    = true;                        // เรียงลำดับ 1,2,3,4 ต่อเนื่องทุกโหมด
  const holidaySet  = new Set(((entries[0] && entries[0].holidayDays) || []).map(h => h.day));
  const perRowRoute = (mode === 'sm');             // SM: รอบรถ/สาขา รายคน, อื่นๆ: บนหัวครั้งเดียว

  // คอลัมน์คงที่
  const FIXED = perRowRoute
    ? ['ลำดับ','ชื่อสาขา','ตารางรอบขนส่ง DC','ตารางรอบขนส่ง FDC','เวลาทำการ','รหัสพนักงาน','ชื่อ-สกุล','ชื่อเล่น','เบอร์โทร','วันหยุดประจำสัปดาห์']
    : ['ลำดับ','ชื่อสาขา','รหัสพนักงาน','ชื่อ-สกุล','ชื่อเล่น','เบอร์โทร','วันหยุดประจำสัปดาห์'];
  const SUMM  = ['OFF','เช้า','บ่าย','ดึก','วันทำงาน','รวม ชม.','โอที','หมายเหตุ'];
  const nFixed = FIXED.length;
  const nDays  = daysInMonth;
  const nSumm  = SUMM.length;
  const totalCols = nFixed + nDays + nSumm;

  // ── สำคัญ: เพิ่มจำนวนคอลัมน์ให้พอ (ชีตใหม่มีแค่ 26 คอลัมน์) ──
  const needCols = totalCols - sheet.getMaxColumns();
  if (needCols > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), needCols);

  const nameIdx = FIXED.indexOf('ชื่อ-สกุล'); // index 0-based ของคอลัมน์ชื่อ

  const HEADER_ROW = 5;
  const DATA_START = 6;

  sheet.setHiddenGridlines(false);

  const b0 = (entries[0] && entries[0].branch) || {};

  // ── หัวรายงาน ──
  sheet.getRange(1, 1, 1, totalCols).merge();
  sheet.getRange(1, 1).setValue(mode === 'sm' ? 'ตารางรอบการทำงาน SM' : 'ตารางรอบการทำงานสาขา')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center').setFontColor('#1a237e');

  sheet.getRange(2, 1, 1, totalCols).merge();
  const subTitle = (mode === 'sm')
    ? `รายงานตาราง SM ทุกสาขา   |   ประจำเดือน ${THAI_MONTHS[month]} พ.ศ. ${year + 543}`
    : `สาขา ${branch.branchCode} — ${branch.branchName}   |   ประจำเดือน ${THAI_MONTHS[month]} พ.ศ. ${year + 543}`;
  sheet.getRange(2, 1).setValue(subTitle)
    .setFontSize(11).setFontWeight('bold').setHorizontalAlignment('center').setFontColor('#333');

  sheet.getRange(3, 1, 1, totalCols).merge();
  const line3 = (mode === 'sm')
    ? `DD: ${b0.rmName || '-'}   |   AM:${String(b0.amCode || '').replace(/^AM/i, '')} ${b0.dmName || '-'}`
    : `รอบขนส่ง DC: ${branch.dcSchedule || '-'}   |   รอบขนส่ง FDC: ${branch.fdcSchedule || '-'}   |   เวลาทำการ: ${branch.workHours}   |   SM: ${branch.smName}   |   AM:${String(branch.amCode || '').replace(/^AM/i, '')} ${branch.dmName || '-'}`;
  sheet.getRange(3, 1).setValue(line3)
    .setFontSize(9).setHorizontalAlignment('center').setFontColor('#666');

  // ── หัวตาราง ──
  const dayH = [];
  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(year, month - 1, d).getDay();
    dayH.push(`${d}\n${DAY_TH[dow]}`);
  }
  sheet.getRange(HEADER_ROW, 1, 1, totalCols).setValues([[...FIXED, ...dayH, ...SUMM]]);
  sheet.getRange(HEADER_ROW, 1, 1, totalCols)
    .setBackground('#1a237e').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true).setFontSize(9);
  sheet.setRowHeight(HEADER_ROW, 34);

  for (let d = 1; d <= daysInMonth; d++) {
    const dow = new Date(year, month - 1, d).getDay();
    const col = nFixed + d;
    if (holidaySet.has(d))            sheet.getRange(HEADER_ROW, col).setBackground('#c62828');
    else if (dow === 0 || dow === 6)  sheet.getRange(HEADER_ROW, col).setBackground('#3949ab');
  }

  // ── ข้อมูล: 2 แถวต่อคน (กะ + โอที) ──
  const matrix    = [];   // ค่าทั้งหมด
  const meta      = [];   // {kind:'shift'|'ot', isSM, dayCodes:[]}
  const otReasons = {};   // สาเหตุ OT -> รวมชั่วโมง (สรุปท้ายรายงาน)

  entries.forEach((e, idx) => {
    const b         = e.branch;
    const slot      = e.slot;
    const empShifts = e.shifts || {};
    const empOT     = e.ot || {};
    const sum       = e.sum || {};
    const isSM      = (slot.empType || '').toUpperCase() === 'SM';
    const no        = numbered ? (idx + 1) : (isSM ? 'SM' : (idx + 1));

    // แถวกะ
    const fixedVals = perRowRoute
      ? [no, b.branchName, b.dcSchedule || '', b.fdcSchedule || '', b.workHours || '',
         slot.empId || '', slot.name || '', slot.nickname || '', slot.phone || '', slot.dayOff || '']
      : [no, b.branchName, slot.empId || '', slot.name || '', slot.nickname || '', slot.phone || '', slot.dayOff || ''];

    const shiftRow = [...fixedVals];
    const dayCodes = [];
    for (let d = 1; d <= daysInMonth; d++) {
      let v = empShifts[d] || '';
      if (!v && holidaySet.has(d)) v = 'PH';
      shiftRow.push(CODE_LABEL[v] || v);  // แสดงชื่อไทยสำหรับรหัสพิเศษ
      dayCodes.push(v);                   // เก็บ code เดิมไว้ระบายสี
    }
    shiftRow.push(sum.totalOff||0, sum.totalAM||0, sum.totalPM||0, sum.totalNight||0,
                  sum.totalWork||0, sum.totalHours||0, sum.totalOT||0, sum.note||'');
    matrix.push(shiftRow);
    meta.push({ kind: 'shift', isSM, dayCodes });

    // แถวโอที (ใต้ชื่อ ระบุเวลาที่ทำ เช่น 12.00-15.00)
    const otRow = new Array(totalCols).fill('');
    otRow[nameIdx] = 'โอที (OT)';
    for (let d = 1; d <= daysInMonth; d++) {
      const parsed = otSplit(empOT[d] || '');
      otRow[nFixed + d - 1] = parsed.time; // แสดงเฉพาะเวลา
      if (parsed.time) {
        const hrs = shiftSpanHours(parsed.time) || 0;
        const key = parsed.reason || '(ไม่ระบุสาเหตุ)';
        otReasons[key] = (otReasons[key] || 0) + hrs;
      }
    }
    otRow[nFixed + nDays + 6] = sum.totalOT || ''; // คอลัมน์ "โอที" ในสรุป
    matrix.push(otRow);
    meta.push({ kind: 'ot', isSM });
  });

  if (matrix.length) {
    sheet.getRange(DATA_START, 1, matrix.length, totalCols).setValues(matrix);
    const dataRange = sheet.getRange(DATA_START, 1, matrix.length, totalCols);
    dataRange.setFontSize(9).setVerticalAlignment('middle');
    sheet.getRange(DATA_START, 1, matrix.length, 1).setHorizontalAlignment('center');
    sheet.getRange(DATA_START, nFixed + 1, matrix.length, nDays)
      .setHorizontalAlignment('center').setWrap(true).setFontSize(8); // วัน: wrap รองรับชื่อไทย
    sheet.getRange(DATA_START, nFixed + nDays + 1, matrix.length, nSumm).setHorizontalAlignment('center');

    for (let r = 0; r < matrix.length; r++) {
      const rowNum = DATA_START + r;
      const m = meta[r];

      if (m.kind === 'ot') {
        // แถวโอที: พื้นส้มอ่อน + ตัวส้ม
        sheet.getRange(rowNum, 1, 1, totalCols).setBackground('#fff3e0').setFontColor('#e65100').setFontSize(8);
        sheet.getRange(rowNum, nameIdx + 1).setFontWeight('bold');
        sheet.setRowHeight(rowNum, 16);
        continue;
      }

      // แถวกะ — ไฮไลต์ SM เฉพาะตอนปนกับพนักงาน (ไม่ใช่รายงาน SM ล้วน)
      if (m.isSM && !numbered) sheet.getRange(rowNum, 1, 1, totalCols).setBackground('#e8eaf6').setFontWeight('bold');

      // มาร์คสีรหัสกะพิเศษ + กะปกติรายเซลล์
      for (let d = 1; d <= daysInMonth; d++) {
        const code = m.dayCodes[d - 1];
        if (!code) continue;
        const c = _shiftCellColor(code);
        if (c) {
          const cell = sheet.getRange(rowNum, nFixed + d);
          cell.setBackground(c.bg).setFontColor(c.fc);
        }
      }
    }
  }

  // ── เส้นขอบ + ฟอนต์ ──
  const allRange = sheet.getRange(HEADER_ROW, 1, matrix.length + 1, totalCols);
  allRange.setBorder(true, true, true, true, true, true, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
  allRange.setFontFamily('Arial');

  // ── ความกว้างคอลัมน์ ──
  // perRowRoute: ลำดับ,DC,FDC,เวลาทำการ,ชื่อสาขา,รหัส,ชื่อ-สกุล,ชื่อเล่น,เบอร์,วันหยุด
  // branch:      ลำดับ,ชื่อสาขา,รหัส,ชื่อ-สกุล,ชื่อเล่น,เบอร์,วันหยุด
  let cw;
  if (perRowRoute) cw = [45,160,120,120,90,85,160,70,95,120];
  else             cw = [45,170,85,170,70,95,120];
  cw.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
  for (let d = 1; d <= daysInMonth; d++) sheet.setColumnWidth(nFixed + d, 64);
  const sumStart = nFixed + nDays + 1;
  for (let c = 0; c < 7; c++) sheet.setColumnWidth(sumStart + c, 48);
  sheet.setColumnWidth(sumStart + 7, 140);

  // ตรึงเฉพาะแถวหัว — ไม่ตรึงคอลัมน์ (เลี่ยง error: เส้น freeze ตัดผ่านเซลล์ title ที่ merge)
  sheet.setFrozenRows(HEADER_ROW);

  // ── คำอธิบายสัญลักษณ์ + สีตัวอย่าง (code = ชื่อภาษาไทย) ──
  const legendRow = DATA_START + matrix.length + 2;
  sheet.getRange(legendRow, 1).setValue('คำอธิบายสี:')
    .setFontWeight('bold').setFontSize(9).setFontColor('#1a237e');
  const legends = ['OFF', 'PH', 'LEAVE', 'SICK', 'COMP'];
  legends.forEach((code, i) => {
    const col = 2 + i * 2;
    const c = EXPORT_SHIFT_COLORS[code];
    sheet.getRange(legendRow, col).setValue(CODE_LABEL[code]).setBackground(c.bg).setFontColor(c.fc)
      .setFontWeight('bold').setHorizontalAlignment('center').setFontSize(8)
      .setBorder(true, true, true, true, false, false, '#bdbdbd', SpreadsheetApp.BorderStyle.SOLID);
  });
  sheet.getRange(legendRow + 1, 1).setValue('โอที (OT): ระบุช่วงเวลาที่ทำใต้ชื่อ เช่น 12.00-15.00')
    .setFontSize(8).setFontColor('#e65100');

  // ── สรุปสาเหตุการทำ OT (มุมล่างขวา) ──
  const reasonKeys = Object.keys(otReasons);
  if (reasonKeys.length) {
    reasonKeys.sort((a, b) => otReasons[b] - otReasons[a]);
    const boxW   = 4; // กว้าง 4 คอลัมน์ (สาเหตุ 3 + ชม. 1)
    const startC = Math.max(1, totalCols - boxW + 1); // ชิดขวา
    let row = legendRow;

    sheet.getRange(row, startC, 1, boxW).merge();
    sheet.getRange(row, startC).setValue('สรุปสาเหตุการทำ OT')
      .setBackground('#e8590c').setFontColor('#fff').setFontWeight('bold')
      .setHorizontalAlignment('center').setFontSize(9);
    row++;

    sheet.getRange(row, startC, 1, boxW - 1).merge();
    sheet.getRange(row, startC).setValue('สาเหตุ').setFontWeight('bold').setFontSize(8)
      .setBackground('#fff3e0').setHorizontalAlignment('center');
    sheet.getRange(row, startC + boxW - 1).setValue('ชม.').setFontWeight('bold').setFontSize(8)
      .setBackground('#fff3e0').setHorizontalAlignment('center');
    row++;

    let totalOtHours = 0;
    reasonKeys.forEach(rk => {
      const hrs = Math.round(otReasons[rk] * 10) / 10;
      totalOtHours += hrs;
      sheet.getRange(row, startC, 1, boxW - 1).merge();
      sheet.getRange(row, startC).setValue(rk).setFontSize(8).setHorizontalAlignment('left')
        .setWrap(true);
      sheet.getRange(row, startC + boxW - 1).setValue(hrs).setFontSize(8)
        .setFontColor('#e65100').setFontWeight('bold').setHorizontalAlignment('center');
      row++;
    });

    sheet.getRange(row, startC, 1, boxW - 1).merge();
    sheet.getRange(row, startC).setValue('รวมทั้งหมด').setFontWeight('bold').setFontSize(8)
      .setHorizontalAlignment('right').setBackground('#fff3e0');
    sheet.getRange(row, startC + boxW - 1).setValue(Math.round(totalOtHours * 10) / 10)
      .setFontWeight('bold').setFontSize(8).setFontColor('#e65100')
      .setHorizontalAlignment('center').setBackground('#fff3e0');

    sheet.getRange(legendRow, startC, row - legendRow + 1, boxW)
      .setBorder(true, true, true, true, true, true, '#e8590c', SpreadsheetApp.BorderStyle.SOLID)
      .setFontFamily('Arial');
  }

  // ── ลายเซ็น (เว้นให้พ้นกล่องสรุป OT) ──
  const signRow = legendRow + Math.max(5, reasonKeys.length + 5);
  sheet.getRange(signRow, 2).setValue('ผู้จัดทำ (SM): ________________________').setFontSize(9);
  sheet.getRange(signRow + 2, 2).setValue('ผู้อนุมัติ (RM/DM): ________________________').setFontSize(9);
}
