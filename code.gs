const SHEET_ID = '1oqAhb_f0aEKIeF3jgtOdT7Hhu2sO4RTa9To4TJRBVXs';

const EMPLOYEE_HEADERS = [
  'EmpID', 'Email', 'Name', 'Role', 'Department',
  'EmploymentType', 'WorkMode', 'Status',
  'ReportingManager', 'ReportingManagerEmail',
  'Manager', 'ManagerEmail',
  'Phone', 'JoinDate',
  'ContractStartDate', 'ContractEndDate', 'ContractTotalDays',
  'NoticePeriod', 'RenewalNotes',
  'CurrentProject', 'AddedOn'
];

const ATTENDANCE_HEADERS = [
  'RecordID', 'EmpID', 'Email', 'Name', 'Department',
  'CurrentProject', 'Date', 'CheckInTime', 'Location', 'Status', 'AttendanceColor'
];

const ABSENCE_HEADERS = [
  'Date', 'Employee ID', 'Email', 'Name', 'Department', 'Role',
  'Reporting Manager', 'Manager', 'Project', 'Work Mode',
  'Employment Type', 'Status'
];

const MANAGER_HEADERS = [
  'ManagerID', 'ManagerType', 'Name', 'Email', 'Role', 'Department', 'AddedOn'
];

const HR_PROFILE_HEADERS = [
  'HRID', 'Username', 'Password', 'Name', 'Email', 'Role', 'Status', 'AddedOn'
];

const LEAVE_HEADERS = [
  'LeaveID', 'EmpID', 'Email', 'Name', 'Department',
  'LeaveType', 'FromDate', 'ToDate', 'Days', 'Reason',
  'Status', 'AppliedOn', 'UpdatedOn', 'ApprovedBy', 'ApprovedDate',
  'RejectedBy', 'RejectedDate', 'RejectionReason',
  'ReportingManager', 'ReportingManagerEmail', 'Manager', 'ManagerEmail'
];

function doGet(e) {
  try {
    const action = e.parameter.action || '';
    const data = e.parameter || {};
    return jsonResponse(handleAction(action, data));
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function doPost(e) {
  try {
    let data = {};
    let action = '';

    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (parseError) {
        data = e.parameter || {};
        if (data.payload) {
          try {
            data = JSON.parse(data.payload);
          } catch (ignored) {}
        }
      }
    } else {
      data = e.parameter || {};
    }

    action = data.action || '';
    return jsonResponse(handleAction(action, data));
  } catch (err) {
    return jsonResponse({ success: false, message: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleAction(action, data) {
  switch (action) {
    case 'ping':              return { success: true, message: 'pong' };
    case 'testEmail':         return testEmailSystem(data);
    case 'setup':             return setupSheets();
    case 'adminLogin':        return adminLogin(data);
    case 'getHRProfiles':     return getHRProfiles();
    case 'addHRProfile':      return addHRProfile(data);
    case 'updateHRProfile':   return updateHRProfile(data);
    case 'deleteHRProfile':   return deleteHRProfile(data.username);
    case 'getEmployee':       return getEmployee(data.email);
    case 'getEmployees':      return getEmployees();
    case 'addEmployee':       return addEmployee(data);
    case 'updateEmployee':    return updateEmployee(data);
    case 'deleteEmployee':    return deleteEmployee(data.email);
    case 'markAttendance':    return handleMarkAttendance(data);
    case 'recordAbsent':      return handleRecordAbsent(data);
    case 'checkLoginBeforeCutoff': return checkLoginBeforeCutoff(data);
    case 'captureAbsentUsers': return captureAbsentUsers(data);
    case 'getAttendance':     return getAttendance(data.email);
    case 'getAllAttendance':  return getAllAttendance();
    case 'applyLeave':        return applyLeave(data);
    case 'getLeaves':         return getLeaves(data.email);
    case 'getAllLeaves':      return getAllLeaves();
    case 'updateLeaveStatus': return updateLeaveStatus(data);
    case 'getProjects':       return getProjects();
    case 'saveProjects':      return saveProjects(data);
    case 'assignProject':     return assignProject(data);
    case 'getEmployeesByProject': return getEmployeesByProject(data);
    case 'getManagers':       return getManagers();
    case 'validateManagerProfile': return validateManagerProfile(data.email, data.role);
    case 'validateProductivityTrackerProfile':
    case 'validateProductionManagerProfile': return validateProductivityTrackerProfile(data.email);
    case 'debugManagerProfile': return debugManagerProfile(data.email);
    case 'addManager':        return addManager(data);
    case 'updateManager':     return updateManager(data);
    case 'deleteManager':     return deleteManager(data.email);
    default:                  return { success: false, message: 'Unknown action: ' + action };
  }
}

function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const configs = {
    'Employees': EMPLOYEE_HEADERS,
    'Attendance': ATTENDANCE_HEADERS,
    'Leaves': LEAVE_HEADERS,
    'Projects': ['ProjectName'],
    'Managers': MANAGER_HEADERS,
    'HRProfiles': HR_PROFILE_HEADERS,
    'Absences': ABSENCE_HEADERS,
    'AbsentUsers': [
      'Date', 'EmpID', 'Email', 'Name', 'Department',
      'Status', 'RecordedOn'
    ]
  };

  for (const [name, headers] of Object.entries(configs)) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      styleHeaderRow(sheet, headers.length);
      sheet.setFrozenRows(1);
    }
  }

  ensureEmployeeSchema(ss);
  ensureAttendanceSchema(ss);
  ensureManagerSchema(ss);
  ensureLeaveSchema(ss);
  ensureAbsenceSchema(ss);
  return { success: true, message: 'All sheets ready!' };
}

function styleHeaderRow(sheet, headerCount) {
  sheet.getRange(1, 1, 1, headerCount)
    .setBackground('#1a1a2e')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
}

function ensureEmployeeSchema(ss) {
  const sheet = ss.getSheetByName('Employees');
  if (!sheet) return;

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(EMPLOYEE_HEADERS);
    styleHeaderRow(sheet, EMPLOYEE_HEADERS.length);
    sheet.setFrozenRows(1);
    return;
  }

  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const existingMap = {};
  existingHeaders.forEach(function(h, i) {
    if (h) existingMap[h.toString().toLowerCase().trim()] = i + 1;
  });

  EMPLOYEE_HEADERS.forEach(function(header) {
    const key = header.toLowerCase().trim();
    if (!Object.prototype.hasOwnProperty.call(existingMap, key)) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(header);
      sheet.getRange(1, col)
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }
  });
}

function ensureManagerSchema(ss) {
  const sheet = ss.getSheetByName('Managers');
  if (!sheet) return;

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(MANAGER_HEADERS);
    styleHeaderRow(sheet, MANAGER_HEADERS.length);
    sheet.setFrozenRows(1);
    return;
  }

  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const existingMap = {};
  existingHeaders.forEach(function(h, i) {
    if (h) existingMap[h.toString().toLowerCase().trim()] = i + 1;
  });

  MANAGER_HEADERS.forEach(function(header) {
    const key = header.toLowerCase().trim();
    if (!Object.prototype.hasOwnProperty.call(existingMap, key)) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(header);
      sheet.getRange(1, col)
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }
  });
}

function ensureLeaveSchema(ss) {
  const sheet = ss.getSheetByName('Leaves');
  if (!sheet) return;

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(LEAVE_HEADERS);
    styleHeaderRow(sheet, LEAVE_HEADERS.length);
    sheet.setFrozenRows(1);
    return;
  }

  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const existingMap = {};
  existingHeaders.forEach(function(h, i) {
    if (h) existingMap[h.toString().toLowerCase().trim()] = i + 1;
  });

  LEAVE_HEADERS.forEach(function(header) {
    const key = header.toLowerCase().trim();
    if (!Object.prototype.hasOwnProperty.call(existingMap, key)) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(header);
      sheet.getRange(1, col)
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }
  });
}

function ensureAttendanceSchema(ss) {
  const sheet = ss.getSheetByName('Attendance');
  if (!sheet) return;

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(ATTENDANCE_HEADERS);
    styleHeaderRow(sheet, ATTENDANCE_HEADERS.length);
    sheet.setFrozenRows(1);
    return;
  }

  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const existingMap = {};
  existingHeaders.forEach(function(h, i) {
    if (h) existingMap[h.toString().toLowerCase().trim()] = i + 1;
  });

  ATTENDANCE_HEADERS.forEach(function(header) {
    const key = header.toLowerCase().trim();
    if (!Object.prototype.hasOwnProperty.call(existingMap, key)) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(header);
      sheet.getRange(1, col)
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }
  });
}

function ensureAbsenceSchema(ss) {
  let sheet = ss.getSheetByName('Absences');
  if (!sheet) {
    sheet = ss.insertSheet('Absences');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(ABSENCE_HEADERS);
    sheet.getRange(1, 1, 1, ABSENCE_HEADERS.length)
      .setBackground('#6c757d')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
    return;
  }

  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const existingMap = {};
  existingHeaders.forEach(function(h, i) {
    if (h) existingMap[h.toString().toLowerCase().trim()] = i + 1;
  });

  ABSENCE_HEADERS.forEach(function(header) {
    const key = header.toLowerCase().trim();
    if (!Object.prototype.hasOwnProperty.call(existingMap, key)) {
      const col = sheet.getLastColumn() + 1;
      sheet.getRange(1, col).setValue(header);
      sheet.getRange(1, col)
        .setBackground('#6c757d')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    }
  });

  sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), ABSENCE_HEADERS.length))
    .setBackground('#6c757d')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
}

function adminLogin(data) {
  const username = (data.username || '').toString().trim();
  const password = (data.password || '').toString();
  if (!username || !password) return { success: false, message: 'Username and password are required' };
  const defaultCreds = [
    { username: 'admin', password: 'admin@123' },
    { username: 'hr', password: 'hr@123' }
  ];
  const matchedDefault = defaultCreds.some(function(c) {
    return c.username === username && c.password === password;
  });
  if (matchedDefault) return {
    success: true,
    message: 'Login successful',
    profile: { username: username, name: username === 'hr' ? 'HR Administrator' : 'Admin', role: 'HR Administrator' }
  };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('HRProfiles');
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName('HRProfiles');
  }
  if (!sheet || sheet.getLastRow() < 2) return { success: false, message: 'Invalid username or password' };

  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getHRHeaders(sheet);
  const targetUser = username.toLowerCase();

  for (let i = 0; i < rows.length; i++) {
    const rowUser = (getValueByHeader(rows[i], hdr, 'username', 1) || '').toString().trim().toLowerCase();
    const rowPass = (getValueByHeader(rows[i], hdr, 'password', 2) || '').toString();
    const rowStatus = (getValueByHeader(rows[i], hdr, 'status', 6) || 'Active').toString().trim().toLowerCase();
    if (rowUser === targetUser && rowPass === password) {
      if (rowStatus && rowStatus !== 'active') return { success: false, message: 'This HR profile is inactive' };
      return {
        success: true,
        message: 'Login successful',
        profile: {
          username: getValueByHeader(rows[i], hdr, 'username', 1) || username,
          name: getValueByHeader(rows[i], hdr, 'name', 3) || username,
          email: getValueByHeader(rows[i], hdr, 'email', 4) || '',
          role: getValueByHeader(rows[i], hdr, 'role', 5) || 'HR Administrator'
        }
      };
    }
  }

  return { success: false, message: 'Invalid username or password' };
}

function getEmpHeaders(sheet) {
  const raw = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  raw.forEach(function(h, i) {
    if (h) map[h.toString().toLowerCase().trim()] = i;
  });
  return map;
}

function getAttHeaders(sheet) {
  const raw = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  raw.forEach(function(h, i) {
    if (h) map[h.toString().toLowerCase().replace(/\s/g, '')] = i;
  });
  return map;
}

function getManagerHeaders(sheet) {
  const raw = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  raw.forEach(function(h, i) {
    if (h) map[h.toString().toLowerCase().trim()] = i;
  });
  return map;
}

function getHRHeaders(sheet) {
  const raw = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  raw.forEach(function(h, i) {
    if (h) map[h.toString().toLowerCase().trim()] = i;
  });
  return map;
}

function getLeaveHeaders(sheet) {
  const raw = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  raw.forEach(function(h, i) {
    if (h) map[h.toString().toLowerCase().trim()] = i;
  });
  return map;
}

function normalizeEmail(email) {
  return (email || '').toString().trim().toLowerCase();
}

const PRODUCTIVITY_TRACKER_LABEL = 'Productivity Tracker';
const LEGACY_PRODUCTIVITY_TRACKER_LABEL = 'Production Manager';
const ALLOWED_EMAIL_DOMAINS = ['@dashversemail.com', '@dashverse.ai', '@dashtoon.com'];
function isAllowedWorkEmail(email) {
  const value = normalizeEmail(email);
  if (!value || value.indexOf('@') === -1) return false;
  for (let i = 0; i < ALLOWED_EMAIL_DOMAINS.length; i++) {
    if (value.endsWith(ALLOWED_EMAIL_DOMAINS[i])) return true;
  }
  return false;
}

function isProductivityTrackerLabel(value) {
  const normalized = (value || '').toString().trim().toLowerCase();
  return normalized === PRODUCTIVITY_TRACKER_LABEL.toLowerCase() ||
    normalized === LEGACY_PRODUCTIVITY_TRACKER_LABEL.toLowerCase();
}

function hasProductivityTrackerRole(value) {
  const normalized = (value || '').toString().trim().toLowerCase();
  return normalized.indexOf(PRODUCTIVITY_TRACKER_LABEL.toLowerCase()) !== -1 ||
    normalized.indexOf(LEGACY_PRODUCTIVITY_TRACKER_LABEL.toLowerCase()) !== -1;
}

function normalizeProductivityTrackerLabel(value) {
  const trimmed = (value || '').toString().trim();
  return isProductivityTrackerLabel(trimmed) ? PRODUCTIVITY_TRACKER_LABEL : trimmed;
}

function normalizeLeaveStatus(status) {
  const s = (status || '').toString().trim().toLowerCase();
  if (s === 'approved') return 'Approved';
  if (s === 'rejected') return 'Rejected';
  if (s === 'cancelled' || s === 'canceled') return 'Cancelled';
  return 'Pending';
}

function parseTimeTo24Hour(timeStr) {
  if (!timeStr) return null;
  try {
    const str = timeStr.toString().trim().toUpperCase();
    const match24 = str.match(/^(\d{1,2}):(\d{2})(?::\d{2})?$/);
    if (match24) {
      return parseInt(match24[1], 10) + parseInt(match24[2], 10) / 60;
    }
    const match12 = str.match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)$/);
    if (match12) {
      let hours = parseInt(match12[1], 10);
      const minutes = parseInt(match12[2], 10);
      const period = match12[3];
      if (period === 'PM' && hours !== 12) hours += 12;
      if (period === 'AM' && hours === 12) hours = 0;
      return hours + minutes / 60;
    }
    const dateObj = new Date(timeStr);
    if (!isNaN(dateObj.getTime())) {
      return dateObj.getHours() + dateObj.getMinutes() / 60;
    }
    return null;
  } catch (e) {
    return null;
  }
}

function formatSheetDate(value) {
  if (!value) return '';
  try {
    if (value instanceof Date) {
      // Google Sheets Date object - format directly in script timezone
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    }
    const date = new Date(value);
    if (isNaN(date.getTime())) return String(value);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  } catch(e) {
    return String(value);
  }
}

function getValueByHeader(row, hdr, key, fallbackIndex) {
  const index = hdr[key] !== undefined ? hdr[key] : fallbackIndex;
  if (index === -1 || index === undefined) return '';
  return row[index] || '';
}

function hasField(data, fieldName) {
  return Object.prototype.hasOwnProperty.call(data, fieldName) && data[fieldName] !== undefined;
}

function setCellIfProvided(sheet, rowNumber, hdr, headerName, value) {
  const index = hdr[headerName];
  if (index !== undefined && value !== undefined) {
    sheet.getRange(rowNumber, index + 1).setValue(value);
  }
}

function applyEmployeeStatusCellStyle(sheet, rowNumber, hdr, status) {
  const index = hdr['status'];
  if (index === undefined) return;

  const cell = sheet.getRange(rowNumber, index + 1);
  const normalized = (status || 'Active').toString().trim().toLowerCase();
  cell.setBackground('#ffffff').setFontColor('#000000').setFontWeight('normal');

  if (normalized === 'active') {
    cell.setBackground('#d4edda').setFontColor('#155724').setFontWeight('bold');
  } else if (normalized === 'inactive') {
    cell.setBackground('#fff3cd').setFontColor('#856404').setFontWeight('bold');
  } else if (normalized === 'rejected') {
    cell.setBackground('#f8d7da').setFontColor('#721c24').setFontWeight('bold');
  }
}

function getAttendanceColorInfo(dateValue) {
  const checkInDate = dateValue instanceof Date && !isNaN(dateValue.getTime()) ? dateValue : new Date();
  const hour = checkInDate.getHours();
  if (hour < 14) {
    return {
      attendanceColor: 'green',
      rowBackground: '#d4edda',
      timeBackground: '#28a745',
      timeFontColor: '#ffffff'
    };
  }
  if (hour < 17) {
    return {
      attendanceColor: 'yellow',
      rowBackground: '#fff3cd',
      timeBackground: '#ffc107',
      timeFontColor: '#000000'
    };
  }
  return {
    attendanceColor: 'red',
    rowBackground: '#f8d7da',
    timeBackground: '#dc3545',
    timeFontColor: '#ffffff'
  };
}

function getAttendanceColorFromRecordedValue(attendanceColor, checkInValue) {
  const normalized = (attendanceColor || '').toString().trim().toLowerCase();
  if (normalized === 'green' || normalized === 'yellow' || normalized === 'red') {
    return normalized;
  }

  const hour = parseTimeTo24Hour(checkInValue);
  if (hour === null) return 'green';
  if (hour < 14) return 'green';
  if (hour < 17) return 'yellow';
  return 'red';
}

function applyAttendanceRowStyle(sheet, rowNumber, columnCount, colorInfo, timeColumnIndex) {
  if (!sheet || !colorInfo) return;
  sheet.getRange(rowNumber, 1, 1, columnCount).setBackground(colorInfo.rowBackground);
  if (timeColumnIndex !== undefined && timeColumnIndex !== -1) {
    sheet.getRange(rowNumber, timeColumnIndex + 1)
      .setBackground(colorInfo.timeBackground)
      .setFontColor(colorInfo.timeFontColor)
      .setFontWeight('bold');
  }
}

function buildEmployeeObject(row, hdr) {
  const employmentType = getValueByHeader(row, hdr, 'employmenttype', -1) ||
    getValueByHeader(row, hdr, 'type', -1) ||
    'Permanent';
  const role = normalizeProductivityTrackerLabel(row[hdr['role']] || '');

  return {
    id: row[hdr['empid']] || '',
    email: row[hdr['email']] || '',
    name: row[hdr['name']] || '',
    role: role,
    department: row[hdr['department']] || '',
    employmentType: employmentType,
    employeeType: employmentType,
    employment_type: employmentType,
    workMode: getValueByHeader(row, hdr, 'workmode', -1),
    status: getValueByHeader(row, hdr, 'status', -1) || 'Active',
    reportingManager: getValueByHeader(row, hdr, 'reportingmanager', 5),
    reportingManagerEmail: getValueByHeader(row, hdr, 'reportingmanageremail', -1),
    manager: getValueByHeader(row, hdr, 'manager', 6),
    managerEmail: getValueByHeader(row, hdr, 'manageremail', -1),
    phone: getValueByHeader(row, hdr, 'phone', 7),
    joinDate: formatSheetDate(getValueByHeader(row, hdr, 'joindate', 8)),
    contractStartDate: formatSheetDate(getValueByHeader(row, hdr, 'contractstartdate', -1)),
    contractEndDate: formatSheetDate(getValueByHeader(row, hdr, 'contractenddate', -1)),
    contractTotalDays: getValueByHeader(row, hdr, 'contracttotaldays', -1),
    noticePeriod: getValueByHeader(row, hdr, 'noticeperiod', -1),
    renewalNotes: getValueByHeader(row, hdr, 'renewalnotes', -1),
    currentProject: getValueByHeader(row, hdr, 'currentproject', -1)
  };
}

function buildManagerObject(row, hdr) {
  const managerType = normalizeProductivityTrackerLabel(getValueByHeader(row, hdr, 'managertype', 1));
  const role = normalizeProductivityTrackerLabel(getValueByHeader(row, hdr, 'role', 4));
  return {
    id: getValueByHeader(row, hdr, 'managerid', 0),
    managerType: managerType,
    name: getValueByHeader(row, hdr, 'name', 2),
    email: getValueByHeader(row, hdr, 'email', 3),
    role: role,
    department: getValueByHeader(row, hdr, 'department', 5),
    addedOn: getValueByHeader(row, hdr, 'addedon', 6)
  };
}

function buildHRProfileObject(row, hdr) {
  return {
    id: getValueByHeader(row, hdr, 'hrid', 0),
    username: getValueByHeader(row, hdr, 'username', 1),
    password: getValueByHeader(row, hdr, 'password', 2),
    name: getValueByHeader(row, hdr, 'name', 3),
    email: getValueByHeader(row, hdr, 'email', 4),
    role: getValueByHeader(row, hdr, 'role', 5),
    status: getValueByHeader(row, hdr, 'status', 6),
    addedOn: getValueByHeader(row, hdr, 'addedon', 7)
  };
}

function buildLeaveObject(row, hdr) {
  return {
    id: getValueByHeader(row, hdr, 'leaveid', 0),
    empId: getValueByHeader(row, hdr, 'empid', 1),
    email: getValueByHeader(row, hdr, 'email', 2),
    name: getValueByHeader(row, hdr, 'name', 3),
    department: getValueByHeader(row, hdr, 'department', 4),
    leaveType: getValueByHeader(row, hdr, 'leavetype', 5),
    fromDate: formatSheetDate(getValueByHeader(row, hdr, 'fromdate', 6)),
    toDate: formatSheetDate(getValueByHeader(row, hdr, 'todate', 7)),
    days: getValueByHeader(row, hdr, 'days', 8),
    reason: getValueByHeader(row, hdr, 'reason', 9),
    status: normalizeLeaveStatus(getValueByHeader(row, hdr, 'status', 10)),
    appliedOn: formatSheetDate(getValueByHeader(row, hdr, 'appliedon', 11)),
    updatedOn: formatSheetDate(getValueByHeader(row, hdr, 'updatedon', 12)),
    approvedBy: getValueByHeader(row, hdr, 'approvedby', 13),
    approvedDate: formatSheetDate(getValueByHeader(row, hdr, 'approveddate', 14)),
    rejectedBy: getValueByHeader(row, hdr, 'rejectedby', 15),
    rejectedDate: formatSheetDate(getValueByHeader(row, hdr, 'rejecteddate', 16)),
    rejectionReason: getValueByHeader(row, hdr, 'rejectionreason', 17),
    reportingManager: getValueByHeader(row, hdr, 'reportingmanager', -1),
    reportingManagerEmail: getValueByHeader(row, hdr, 'reportingmanageremail', -1),
    manager: getValueByHeader(row, hdr, 'manager', -1),
    managerEmail: getValueByHeader(row, hdr, 'manageremail', -1)
  };
}

function attachEmployeeManagersToLeave(leave, employee) {
  if (!leave || !employee) return leave;
  if (!leave.reportingManager) leave.reportingManager = employee.reportingManager || '';
  if (!leave.reportingManagerEmail) leave.reportingManagerEmail = employee.reportingManagerEmail || '';
  if (!leave.manager) leave.manager = employee.manager || '';
  if (!leave.managerEmail) leave.managerEmail = employee.managerEmail || '';
  return leave;
}

function getEmployeeForLeaveNotification(leaveRecord) {
  const leaveEmail = normalizeEmail(leaveRecord ? leaveRecord.email : '');
  let employee = null;
  if (leaveEmail) {
    try {
      const employeeData = getEmployee(leaveEmail);
      if (employeeData.success && employeeData.employee) {
        employee = employeeData.employee;
      }
    } catch (lookupErr) {
      Logger.log('Employee lookup failed for leave status email: ' + lookupErr.toString());
    }
  }

  if (!employee) {
    employee = {
      email: leaveEmail,
      name: leaveRecord ? leaveRecord.name : '',
      department: leaveRecord ? leaveRecord.department : '',
      manager: leaveRecord ? leaveRecord.manager : '',
      managerEmail: leaveRecord ? leaveRecord.managerEmail : '',
      reportingManager: leaveRecord ? leaveRecord.reportingManager : '',
      reportingManagerEmail: leaveRecord ? leaveRecord.reportingManagerEmail : ''
    };
  } else {
    attachEmployeeManagersToLeave(leaveRecord, employee);
  }
  return employee;
}

function buildAttendanceObject(row, hdr) {
  const rawCheckIn = getValueByHeader(row, hdr, 'checkintime', 7);
  let checkInStr = '';
  if (rawCheckIn) {
    try {
      if (rawCheckIn instanceof Date) {
        checkInStr = Utilities.formatDate(rawCheckIn, Session.getScriptTimeZone(), 'hh:mm a');
      } else {
        const d = new Date(rawCheckIn);
        if (!isNaN(d.getTime())) {
          checkInStr = Utilities.formatDate(d, Session.getScriptTimeZone(), 'hh:mm a');
        } else {
          checkInStr = String(rawCheckIn);
        }
      }
    } catch(e) {
      checkInStr = String(rawCheckIn);
    }
  }
  return {
    recordId: getValueByHeader(row, hdr, 'recordid', 0),
    empId: getValueByHeader(row, hdr, 'empid', 1),
    email: getValueByHeader(row, hdr, 'email', 2),
    name: getValueByHeader(row, hdr, 'name', 3),
    department: getValueByHeader(row, hdr, 'department', 4),
    currentProject: getValueByHeader(row, hdr, 'currentproject', 5),
    date: formatSheetDate(getValueByHeader(row, hdr, 'date', 6)),
    checkIn: checkInStr,
    location: getValueByHeader(row, hdr, 'location', 8),
    status: getValueByHeader(row, hdr, 'status', 9),
    attendanceColor: getValueByHeader(row, hdr, 'attendancecolor', -1)
  };
}

function getEmployee(email) {
  if (!email) return { success: false, message: 'Email is required' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  ensureEmployeeSchema(ss);
  const sheet = ss.getSheetByName('Employees');
  if (!sheet) return { success: false, message: 'Employees sheet not found.' };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'No employees found. Ask admin to add you.' };

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getEmpHeaders(sheet);
  const wantedEmail = normalizeEmail(email);

  for (let i = 0; i < rows.length; i++) {
    const rowEmail = normalizeEmail(getValueByHeader(rows[i], hdr, 'email', 1));
    if (rowEmail === wantedEmail) {
      return { success: true, employee: buildEmployeeObject(rows[i], hdr) };
    }
  }

  return { success: false, message: 'Employee not found. Please contact your admin.' };
}

function getEmployees() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ensureEmployeeSchema(ss);
  const sheet = ss.getSheetByName('Employees');

  if (!sheet || sheet.getLastRow() < 2) return { success: true, employees: [] };

  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getEmpHeaders(sheet);
  const employees = [];

  for (let i = 0; i < rows.length; i++) {
    if (rows[i][hdr['empid']] || rows[i][hdr['email']]) {
      employees.push(buildEmployeeObject(rows[i], hdr));
    }
  }

  return { success: true, employees: employees };
}

function addEmployee(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ensureEmployeeSchema(ss);
    const sheet = ss.getSheetByName('Employees');
    
    if (!sheet) {
      return { success: false, message: 'Employees sheet not found. Please run setup first.' };
    }

    const hdr = getEmpHeaders(sheet);
    const email = normalizeEmail(data.email);
    const reportingManagerEmail = normalizeEmail(data.reportingManagerEmail);
    const managerEmail = normalizeEmail(data.managerEmail);
    const employeeName = (data.name || '').toString().trim();

    if (!email) return { success: false, message: 'Email is required' };
    if (!employeeName) return { success: false, message: 'Employee name is required' };
    if (!isAllowedWorkEmail(email)) {
      return { success: false, message: 'Email domain not allowed. Use @dashversemail.com, @dashverse.ai, or @dashtoon.com' };
    }
    if (reportingManagerEmail && !isAllowedWorkEmail(reportingManagerEmail)) {
      return { success: false, message: 'Reporting Manager email domain is not allowed' };
    }
    if (managerEmail && !isAllowedWorkEmail(managerEmail)) {
      return { success: false, message: 'Manager email domain is not allowed' };
    }
    if (sheet.getLastRow() > 1) {
      const lastRow = sheet.getLastRow();
      const existing = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      for (let i = 0; i < existing.length; i++) {
        const emailIdx = hdr['email'] !== undefined ? hdr['email'] : 1;
        if (normalizeEmail(existing[i][emailIdx]) === email) {
          return { success: false, message: 'Employee with this email already exists!' };
        }
      }
    }

    const empId = 'EMP' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const totalCols = sheet.getLastColumn();
    
    if (totalCols < 1) {
      return { success: false, message: 'Employees sheet has no columns. Please run setup.' };
    }
    const newRow = [];
    for (let col = 0; col < totalCols; col++) {
      newRow[col] = '';
    }
    if (hdr['empid'] !== undefined) newRow[hdr['empid']] = empId;
    if (hdr['email'] !== undefined) newRow[hdr['email']] = email;
    if (hdr['name'] !== undefined) newRow[hdr['name']] = employeeName;
    if (hdr['role'] !== undefined) newRow[hdr['role']] = data.role || '';
    if (hdr['department'] !== undefined) newRow[hdr['department']] = data.department || '';
    if (hdr['employmenttype'] !== undefined) {
      newRow[hdr['employmenttype']] = data.employmentType || data.employeeType || data.employment_type || data.type || 'Permanent';
    }
    if (hdr['workmode'] !== undefined) newRow[hdr['workmode']] = data.workMode || '';
    if (hdr['status'] !== undefined) newRow[hdr['status']] = data.status || 'Active';
    if (hdr['reportingmanager'] !== undefined) newRow[hdr['reportingmanager']] = data.reportingManager || '';
    if (hdr['reportingmanageremail'] !== undefined) newRow[hdr['reportingmanageremail']] = reportingManagerEmail || '';
    if (hdr['manager'] !== undefined) newRow[hdr['manager']] = data.manager || '';
    if (hdr['manageremail'] !== undefined) newRow[hdr['manageremail']] = managerEmail || '';
    if (hdr['phone'] !== undefined) newRow[hdr['phone']] = data.phone || '';
    if (hdr['joindate'] !== undefined) newRow[hdr['joindate']] = data.joinDate || '';
    if (hdr['contractstartdate'] !== undefined) newRow[hdr['contractstartdate']] = data.contractStartDate || '';
    if (hdr['contractenddate'] !== undefined) newRow[hdr['contractenddate']] = data.contractEndDate || '';
    if (hdr['contracttotaldays'] !== undefined) newRow[hdr['contracttotaldays']] = data.contractTotalDays || '';
    if (hdr['noticeperiod'] !== undefined) newRow[hdr['noticeperiod']] = data.noticePeriod || '';
    if (hdr['renewalnotes'] !== undefined) newRow[hdr['renewalnotes']] = data.renewalNotes || '';
    if (hdr['currentproject'] !== undefined) newRow[hdr['currentproject']] = data.currentProject || '';
    if (hdr['addedon'] !== undefined) newRow[hdr['addedon']] = now;
    sheet.appendRow(newRow);
    applyEmployeeStatusCellStyle(sheet, sheet.getLastRow(), hdr, data.status || 'Active');
    SpreadsheetApp.flush();
    
    Logger.log('Employee added successfully: ' + empId + ' - ' + email);
    return {
      success: true,
      message: 'Employee profile created successfully!',
      id: empId,
      empId: empId,
      employeeName: employeeName,
      employeeEmail: email
    };
  } catch (err) {
    Logger.log('Error in addEmployee: ' + err.toString());
    return { success: false, message: 'Error adding employee: ' + err.toString() };
  }
}

function updateEmployee(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ensureEmployeeSchema(ss);
    const sheet = ss.getSheetByName('Employees');
    if (!sheet) return { success: false, message: 'Employees sheet not found' };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'No employees found' };

    if (!data.email) return { success: false, message: 'Email is required for update' };
    if (!isAllowedWorkEmail(data.email)) {
      return { success: false, message: 'Email domain not allowed. Use @dashversemail.com, @dashverse.ai, or @dashtoon.com' };
    }

    if (hasField(data, 'reportingManagerEmail') && data.reportingManagerEmail && !isAllowedWorkEmail(data.reportingManagerEmail)) {
      return { success: false, message: 'Reporting Manager email domain is not allowed' };
    }
    if (hasField(data, 'managerEmail') && data.managerEmail && !isAllowedWorkEmail(data.managerEmail)) {
      return { success: false, message: 'Manager email domain is not allowed' };
    }

    const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const hdr = getEmpHeaders(sheet);
    const targetEmail = normalizeEmail(data.email);
    const reportingManagerEmail = normalizeEmail(data.reportingManagerEmail);
    const managerEmail = normalizeEmail(data.managerEmail);
    for (let i = 0; i < rows.length; i++) {
      const emailIdx = hdr['email'] !== undefined ? hdr['email'] : 1;
      const rowEmail = normalizeEmail(rows[i][emailIdx]);
      if (rowEmail === targetEmail) {
        const r = i + 2;

        if (hasField(data, 'name'))                 setCellIfProvided(sheet, r, hdr, 'name', data.name);
        if (hasField(data, 'role'))                 setCellIfProvided(sheet, r, hdr, 'role', data.role);
        if (hasField(data, 'department'))           setCellIfProvided(sheet, r, hdr, 'department', data.department);
        if (hasField(data, 'employmentType') || hasField(data, 'employeeType') || hasField(data, 'employment_type') || hasField(data, 'type')) {
          setCellIfProvided(sheet, r, hdr, 'employmenttype', data.employmentType || data.employeeType || data.employment_type || data.type || 'Permanent');
        }
        if (hasField(data, 'workMode'))             setCellIfProvided(sheet, r, hdr, 'workmode', data.workMode);
        if (hasField(data, 'status'))               setCellIfProvided(sheet, r, hdr, 'status', data.status || 'Active');
        if (hasField(data, 'reportingManager'))     setCellIfProvided(sheet, r, hdr, 'reportingmanager', data.reportingManager);
        if (hasField(data, 'reportingManagerEmail'))setCellIfProvided(sheet, r, hdr, 'reportingmanageremail', reportingManagerEmail);
        if (hasField(data, 'manager'))              setCellIfProvided(sheet, r, hdr, 'manager', data.manager);
        if (hasField(data, 'managerEmail'))         setCellIfProvided(sheet, r, hdr, 'manageremail', managerEmail);
        if (hasField(data, 'phone'))                setCellIfProvided(sheet, r, hdr, 'phone', data.phone);
        if (hasField(data, 'joinDate'))             setCellIfProvided(sheet, r, hdr, 'joindate', data.joinDate);
        if (hasField(data, 'contractStartDate'))    setCellIfProvided(sheet, r, hdr, 'contractstartdate', data.contractStartDate);
        if (hasField(data, 'contractEndDate'))      setCellIfProvided(sheet, r, hdr, 'contractenddate', data.contractEndDate);
        if (hasField(data, 'contractTotalDays'))    setCellIfProvided(sheet, r, hdr, 'contracttotaldays', data.contractTotalDays);
        if (hasField(data, 'noticePeriod'))         setCellIfProvided(sheet, r, hdr, 'noticeperiod', data.noticePeriod);
        if (hasField(data, 'renewalNotes'))         setCellIfProvided(sheet, r, hdr, 'renewalnotes', data.renewalNotes);
        if (hasField(data, 'currentProject'))       setCellIfProvided(sheet, r, hdr, 'currentproject', data.currentProject);
        if (hasField(data, 'status'))               applyEmployeeStatusCellStyle(sheet, r, hdr, data.status || 'Active');
        SpreadsheetApp.flush();

        Logger.log('Employee updated successfully: ' + data.email);
        return {
          success: true,
          message: 'Employee profile updated successfully!',
          employeeName: hasField(data, 'name') ? (data.name || '').toString().trim() : (rows[i][hdr['name']] || ''),
          employeeEmail: targetEmail
        };
      }
    }

    return { success: false, message: 'Employee not found' };
  } catch (err) {
    Logger.log('Error in updateEmployee: ' + err.toString());
    return { success: false, message: 'Error updating employee: ' + err.toString() };
  }
}

function deleteEmployee(email) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ensureEmployeeSchema(ss);
  const sheet = ss.getSheetByName('Employees');
  if (!sheet) return { success: false, message: 'Sheet not found' };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'No employees found' };

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getEmpHeaders(sheet);

  for (let i = 0; i < rows.length; i++) {
    if (normalizeEmail(rows[i][hdr['email']]) === normalizeEmail(email)) {
      sheet.deleteRow(i + 2);
      Logger.log('Employee deleted: ' + email);
      return { success: true, message: 'Employee profile has been deleted successfully!' };
    }
  }

  return { success: false, message: 'Employee not found' };
}

function getManagers() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Managers');

  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName('Managers');
  }

  ensureManagerSchema(ss);

  if (!sheet || sheet.getLastRow() < 2) return { success: true, managers: [] };

  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getManagerHeaders(sheet);
  const managers = [];

  for (let i = 0; i < rows.length; i++) {
    if (getValueByHeader(rows[i], hdr, 'email', 3)) {
      managers.push(buildManagerObject(rows[i], hdr));
    }
  }

  return { success: true, managers: managers };
}

function debugManagerProfile(email) {
  const normalizedEmail = normalizeEmail(email);
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Managers');
  
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName('Managers');
  }

  ensureManagerSchema(ss);
  
  const hdr = getManagerHeaders(sheet);
  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  
  const allManagers = [];
  for (let i = 0; i < rows.length; i++) {
    const mgrEmail = normalizeEmail(getValueByHeader(rows[i], hdr, 'email', 3));
    const mgrType = (getValueByHeader(rows[i], hdr, 'managertype', 1) || '').toLowerCase().trim();
    const mgrName = getValueByHeader(rows[i], hdr, 'name', 1) || '';
    
    allManagers.push({
      row: i + 2,
      name: mgrName,
      email: mgrEmail,
      emailRaw: getValueByHeader(rows[i], hdr, 'email', 3),
      managerType: mgrType,
      exactMatch: mgrEmail === normalizedEmail
    });
  }
  
  return {
    searchingFor: normalizedEmail,
    allManagers: allManagers,
    totalManagers: allManagers.length,
    headers: hdr
  };
}

function validateManagerProfile(email, role) {
  if (!email) return { success: false, message: 'Email is required' };
  
  const normalizedEmail = normalizeEmail(email);
  if (!normalizedEmail) return { success: false, message: 'Invalid email format' };
  
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Managers');
  
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName('Managers');
  }

  ensureManagerSchema(ss);
  
  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, message: 'Manager profile not found. Please contact HR.' };
  }
  
  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getManagerHeaders(sheet);
  
  // Log for debugging
  Logger.log('Searching for email: ' + normalizedEmail + ' with role: ' + role);
  
  for (let i = 0; i < rows.length; i++) {
    const emailVal = getValueByHeader(rows[i], hdr, 'email', 3);
    const mgrEmail = normalizeEmail(emailVal);
    const mgrTypeVal = getValueByHeader(rows[i], hdr, 'managertype', 1) || '';
    const mgrType = mgrTypeVal.toString().toLowerCase().trim();
    
    Logger.log('Row ' + (i+2) + ': Email=' + mgrEmail + ', Type=' + mgrType);
    
    if (mgrEmail === normalizedEmail) {
      Logger.log('Found matching email! Checking role...');
      const normalizedRole = (role || '').toLowerCase().trim();
      if (!role || mgrType === normalizedRole) {
        Logger.log('Role matches! Access granted.');
        return { success: true, message: 'Profile validated', manager: buildManagerObject(rows[i], hdr) };
      } else {
        Logger.log('Role mismatch. Manager type: ' + mgrType + ', Expected: ' + normalizedRole);
        return { success: false, message: `Your profile is registered as '${mgrType}', but trying to login as '${role}'. Please select the correct role.` };
      }
    }
  }
  
  Logger.log('No matching email found in database');
  return { success: false, message: 'Your profile is not registered as a manager. Please contact HR.' };
}

function validateProductivityTrackerProfile(email) {
  if (!email) return { success: false, message: 'Email is required' };

  const normalizedEmail = normalizeEmail(email);
  if (!normalizedEmail) return { success: false, message: 'Invalid email format' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Managers');
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName('Managers');
  }
  ensureManagerSchema(ss);

  let matchedNonTrackerProfile = null;
  if (sheet && sheet.getLastRow() >= 2) {
    const hdr = getManagerHeaders(sheet);
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

    for (let i = 0; i < rows.length; i++) {
      const rowEmail = normalizeEmail(getValueByHeader(rows[i], hdr, 'email', 3));
      if (rowEmail !== normalizedEmail) continue;

      const managerType = (getValueByHeader(rows[i], hdr, 'managertype', 1) || '').toString().trim().toLowerCase();
      const roleVal = (getValueByHeader(rows[i], hdr, 'role', 4) || '').toString().trim().toLowerCase();
      const nameVal = (getValueByHeader(rows[i], hdr, 'name', 2) || '').toString().trim();
      const isProductivityTracker =
        isProductivityTrackerLabel(managerType) ||
        ((managerType === 'manager' || !managerType) && hasProductivityTrackerRole(roleVal));

      if (!isProductivityTracker) {
        matchedNonTrackerProfile = buildManagerObject(rows[i], hdr);
        continue;
      }

      return {
        success: true,
        message: 'Profile validated',
        manager: buildManagerObject(rows[i], hdr),
        name: nameVal || (normalizedEmail.split('@')[0] || '')
      };
    }
  }

  const employeeProfile = findProductivityTrackerEmployeeProfile(ss, normalizedEmail);
  if (employeeProfile) {
    return {
      success: true,
      message: 'Profile validated',
      manager: {
        id: employeeProfile.id || '',
        managerType: PRODUCTIVITY_TRACKER_LABEL,
        name: employeeProfile.name || '',
        email: employeeProfile.email || normalizedEmail,
        role: normalizeProductivityTrackerLabel(employeeProfile.role) || PRODUCTIVITY_TRACKER_LABEL,
        department: employeeProfile.department || '',
        addedOn: ''
      },
      name: employeeProfile.name || (normalizedEmail.split('@')[0] || '')
    };
  }

  if (matchedNonTrackerProfile) {
    return { success: false, message: 'Your backend profile is not registered as a Productivity Tracker. Please contact HR.' };
  }

  return { success: false, message: 'Productivity Tracker profile not found. Please contact HR.' };
}

function findProductivityTrackerEmployeeProfile(ss, email) {
  ensureEmployeeSchema(ss);
  const sheet = ss.getSheetByName('Employees');
  if (!sheet || sheet.getLastRow() < 2) return null;

  const hdr = getEmpHeaders(sheet);
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < rows.length; i++) {
    const rowEmail = normalizeEmail(getValueByHeader(rows[i], hdr, 'email', 1));
    if (rowEmail !== email) continue;

    const roleVal = (getValueByHeader(rows[i], hdr, 'role', 3) || '').toString().trim().toLowerCase();
    if (hasProductivityTrackerRole(roleVal)) {
      return buildEmployeeObject(rows[i], hdr);
    }
  }
  return null;
}

function normalizeProjectName(name) {
  return (name || '').toString().trim().toLowerCase();
}

function splitProjects(value) {
  if (!value) return [];
  return value
    .toString()
    .split(',')
    .map(function(p) { return p.trim(); })
    .filter(Boolean);
}

function getEmployeesByProject(data) {
  const project = (data.project || data.projectName || '').toString().trim();
  if (!project) return { success: false, message: 'Project is required' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  ensureEmployeeSchema(ss);

  const empSheet = ss.getSheetByName('Employees');
  const attSheet = ss.getSheetByName('Attendance');
  if (!empSheet) return { success: false, message: 'Employees sheet not found' };
  if (!attSheet) return { success: true, project: project, employees: [] };

  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const timeFilter = (data.timeFilter || 'all').toString().trim().toLowerCase();

  const pKey = normalizeProjectName(project);

  // Build employee index by email
  const empHdr = getEmpHeaders(empSheet);
  const empRows = empSheet.getLastRow() > 1
    ? empSheet.getRange(2, 1, empSheet.getLastRow() - 1, empSheet.getLastColumn()).getValues()
    : [];
  const empByEmail = {};
  for (let i = 0; i < empRows.length; i++) {
    const rowEmail = normalizeEmail(getValueByHeader(empRows[i], empHdr, 'email', 1));
    if (!rowEmail) continue;
    empByEmail[rowEmail] = buildEmployeeObject(empRows[i], empHdr);
  }

  // Scan attendance logs for employees who selected this project
  const aHdr = getAttHeaders(attSheet);
  const attRows = attSheet.getLastRow() > 1
    ? attSheet.getRange(2, 1, attSheet.getLastRow() - 1, attSheet.getLastColumn()).getValues()
    : [];

  const iMail = aHdr['email'] !== undefined ? aHdr['email'] : 2;
  const iName = aHdr['name'] !== undefined ? aHdr['name'] : 3;
  const iDept = aHdr['department'] !== undefined ? aHdr['department'] : 4;
  const iProj = aHdr['currentproject'] !== undefined ? aHdr['currentproject'] : 5;
  const iDate = aHdr['date'] !== undefined ? aHdr['date'] : 6;
  const iTime = aHdr['checkintime'] !== undefined ? aHdr['checkintime'] : 7;

  // latestByEmail[email] = { date, time, dateTimeKey }
  const latestByEmail = {};
  for (let r = 0; r < attRows.length; r++) {
    const email = normalizeEmail(attRows[r][iMail]);
    if (!email) continue;

    const projRaw = iProj !== -1 ? attRows[r][iProj] : '';
    const projects = splitProjects(projRaw).map(normalizeProjectName);
    if (projects.indexOf(pKey) === -1) continue;

    const dateStr = formatSheetDate(attRows[r][iDate]) || '';
    if (timeFilter === 'today' && dateStr !== today) continue;
    const timeStr = (attRows[r][iTime] || '').toString().trim();
    const dateTimeKey = (dateStr || '') + ' ' + (timeStr || '');

    const prev = latestByEmail[email];
    if (!prev || dateTimeKey > prev.dateTimeKey) {
      latestByEmail[email] = { date: dateStr, time: timeStr, dateTimeKey: dateTimeKey };
    }

    // If employee record missing, create a minimal one from attendance
    if (!empByEmail[email]) {
      empByEmail[email] = {
        id: '',
        email: email,
        name: attRows[r][iName] || '',
        role: '',
        department: attRows[r][iDept] || '',
        employmentType: '',
        workMode: '',
        reportingManager: '',
        reportingManagerEmail: '',
        manager: '',
        managerEmail: '',
        phone: '',
        joinDate: '',
        contractStartDate: '',
        contractEndDate: '',
        contractTotalDays: '',
        noticePeriod: '',
        renewalNotes: '',
        currentProject: projRaw || ''
      };
    }
  }

  const emails = Object.keys(latestByEmail);
  const employees = emails
    .map(function(email) {
      const emp = empByEmail[email] || { email: email };
      const last = latestByEmail[email];
      return Object.assign({}, emp, {
        lastProjectCheckInDate: last ? last.date : '',
        lastProjectCheckInTime: last ? last.time : ''
      });
    })
    .sort(function(a, b) {
      const ak = (a.lastProjectCheckInDate || '') + ' ' + (a.lastProjectCheckInTime || '');
      const bk = (b.lastProjectCheckInDate || '') + ' ' + (b.lastProjectCheckInTime || '');
      return bk.localeCompare(ak);
    });

  return {
    success: true,
    project: project,
    employees: employees,
    count: employees.length
  };
}

function addManager(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Managers');
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName('Managers');
  }
  ensureManagerSchema(ss);

  const hdr = getManagerHeaders(sheet);
  const email = normalizeEmail(data.email);
  if (!email) return { success: false, message: 'Email is required' };
  if (!isAllowedWorkEmail(email)) {
    return { success: false, message: 'Email domain not allowed. Use @dashversemail.com, @dashverse.ai, or @dashtoon.com' };
  }
  if (!data.name) return { success: false, message: 'Name is required' };
  if (!data.managerType) return { success: false, message: 'Manager type is required' };
  const managerType = normalizeProductivityTrackerLabel(data.managerType);
  const role = normalizeProductivityTrackerLabel(data.role);

  if (sheet.getLastRow() > 1) {
    const lastRow = sheet.getLastRow();
    const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (normalizeEmail(getValueByHeader(rows[i], hdr, 'email', 3)) === email) {
        return { success: false, message: 'Manager with this email already exists!' };
      }
    }
  }

  const managerId = 'MGR' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const totalCols = sheet.getLastColumn();
  const newRow = new Array(totalCols).fill('');

  if (hdr['managerid'] !== undefined) newRow[hdr['managerid']] = managerId;
  if (hdr['managertype'] !== undefined) newRow[hdr['managertype']] = managerType || '';
  if (hdr['name'] !== undefined) newRow[hdr['name']] = data.name || '';
  if (hdr['email'] !== undefined) newRow[hdr['email']] = email;
  if (hdr['role'] !== undefined) newRow[hdr['role']] = role || '';
  if (hdr['department'] !== undefined) newRow[hdr['department']] = data.department || '';
  if (hdr['addedon'] !== undefined) newRow[hdr['addedon']] = now;

  sheet.appendRow(newRow);
  return { success: true, message: 'Manager added successfully!', id: managerId };
}

function updateManager(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Managers');
  if (!sheet) return { success: false, message: 'Managers sheet not found' };
  ensureManagerSchema(ss);

  const hdr = getManagerHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'No managers found' };

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const email = normalizeEmail(data.email);
  if (!isAllowedWorkEmail(email)) {
    return { success: false, message: 'Email domain not allowed. Use @dashversemail.com, @dashverse.ai, or @dashtoon.com' };
  }
  const managerType = normalizeProductivityTrackerLabel(data.managerType);
  const role = normalizeProductivityTrackerLabel(data.role);

  for (let i = 0; i < rows.length; i++) {
    if (normalizeEmail(getValueByHeader(rows[i], hdr, 'email', 3)) === email) {
      const r = i + 2;
      if (hasField(data, 'managerType')) setCellIfProvided(sheet, r, hdr, 'managertype', managerType);
      if (hasField(data, 'name'))        setCellIfProvided(sheet, r, hdr, 'name', data.name);
      if (hasField(data, 'role'))        setCellIfProvided(sheet, r, hdr, 'role', role);
      if (hasField(data, 'department'))  setCellIfProvided(sheet, r, hdr, 'department', data.department);
      return { success: true, message: 'Manager updated successfully!' };
    }
  }

  return { success: false, message: 'Manager not found' };
}

function deleteManager(email) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Managers');
  if (!sheet) return { success: false, message: 'Managers sheet not found' };
  ensureManagerSchema(ss);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'No managers found' };

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getManagerHeaders(sheet);

  for (let i = 0; i < rows.length; i++) {
    if (normalizeEmail(getValueByHeader(rows[i], hdr, 'email', 3)) === normalizeEmail(email)) {
      sheet.deleteRow(i + 2);
      return { success: true, message: 'Manager deleted successfully' };
    }
  }

  return { success: false, message: 'Manager not found' };
}

function getHRProfiles() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('HRProfiles');
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName('HRProfiles');
  }
  if (!sheet || sheet.getLastRow() < 2) return { success: true, profiles: [] };

  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getHRHeaders(sheet);
  const profiles = [];

  for (let i = 0; i < rows.length; i++) {
    if (getValueByHeader(rows[i], hdr, 'username', 1)) {
      profiles.push(buildHRProfileObject(rows[i], hdr));
    }
  }

  return { success: true, profiles: profiles };
}

function addHRProfile(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('HRProfiles');
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName('HRProfiles');
  }

  const hdr = getHRHeaders(sheet);
  const username = (data.username || '').toString().trim();
  const password = (data.password || '').toString();
  if (!username) return { success: false, message: 'Username is required' };
  if (!password) return { success: false, message: 'Password is required' };
  if (data.email && !isAllowedWorkEmail(data.email)) {
    return { success: false, message: 'Email domain not allowed. Use @dashversemail.com, @dashverse.ai, or @dashtoon.com' };
  }

  if (sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const target = username.toLowerCase();
    for (let i = 0; i < rows.length; i++) {
      const rowUser = (getValueByHeader(rows[i], hdr, 'username', 1) || '').toString().trim().toLowerCase();
      if (rowUser === target) return { success: false, message: 'Username already exists' };
    }
  }

  const id = 'HR' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const totalCols = sheet.getLastColumn();
  const newRow = new Array(totalCols).fill('');

  if (hdr['hrid'] !== undefined) newRow[hdr['hrid']] = id;
  if (hdr['username'] !== undefined) newRow[hdr['username']] = username;
  if (hdr['password'] !== undefined) newRow[hdr['password']] = password;
  if (hdr['name'] !== undefined) newRow[hdr['name']] = data.name || '';
  if (hdr['email'] !== undefined) newRow[hdr['email']] = data.email || '';
  if (hdr['role'] !== undefined) newRow[hdr['role']] = data.role || 'HR Administrator';
  if (hdr['status'] !== undefined) newRow[hdr['status']] = data.status || 'Active';
  if (hdr['addedon'] !== undefined) newRow[hdr['addedon']] = now;

  sheet.appendRow(newRow);
  return { success: true, message: 'HR profile created successfully', id: id };
}

function updateHRProfile(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('HRProfiles');
  if (!sheet) return { success: false, message: 'HRProfiles sheet not found' };

  const username = (data.username || '').toString().trim();
  if (!username) return { success: false, message: 'Username is required' };
  if (hasField(data, 'email') && data.email && !isAllowedWorkEmail(data.email)) {
    return { success: false, message: 'Email domain not allowed. Use @dashversemail.com, @dashverse.ai, or @dashtoon.com' };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'No HR profiles found' };

  const hdr = getHRHeaders(sheet);
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const target = username.toLowerCase();

  for (let i = 0; i < rows.length; i++) {
    const rowUser = (getValueByHeader(rows[i], hdr, 'username', 1) || '').toString().trim().toLowerCase();
    if (rowUser === target) {
      const r = i + 2;
      if (hasField(data, 'password')) setCellIfProvided(sheet, r, hdr, 'password', data.password);
      if (hasField(data, 'name')) setCellIfProvided(sheet, r, hdr, 'name', data.name);
      if (hasField(data, 'email')) setCellIfProvided(sheet, r, hdr, 'email', data.email);
      if (hasField(data, 'role')) setCellIfProvided(sheet, r, hdr, 'role', data.role);
      if (hasField(data, 'status')) setCellIfProvided(sheet, r, hdr, 'status', data.status);
      return { success: true, message: 'HR profile updated successfully' };
    }
  }

  return { success: false, message: 'HR profile not found' };
}

function deleteHRProfile(username) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('HRProfiles');
  if (!sheet) return { success: false, message: 'HRProfiles sheet not found' };

  const target = (username || '').toString().trim().toLowerCase();
  if (!target) return { success: false, message: 'Username is required' };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'No HR profiles found' };

  const hdr = getHRHeaders(sheet);
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (let i = 0; i < rows.length; i++) {
    const rowUser = (getValueByHeader(rows[i], hdr, 'username', 1) || '').toString().trim().toLowerCase();
    if (rowUser === target) {
      sheet.deleteRow(i + 2);
      return { success: true, message: 'HR profile deleted successfully' };
    }
  }

  return { success: false, message: 'HR profile not found' };
}

function handleMarkAttendance(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Attendance');
    if (!sheet) {
      setupSheets();
      sheet = ss.getSheetByName('Attendance');
    }

    ensureAttendanceSchema(ss);

    const email = normalizeEmail(data.email);
    if (!email) return { success: false, message: 'Email is required' };

    const tz = Session.getScriptTimeZone();
    const rawCheckIn = data.checkInTime || new Date().toISOString();
    let checkInDate = new Date(rawCheckIn);
    if (isNaN(checkInDate.getTime())) checkInDate = new Date();

    const today = Utilities.formatDate(checkInDate, tz, 'yyyy-MM-dd');
    const timeStr = Utilities.formatDate(checkInDate, tz, 'hh:mm a');
    const colorInfo = getAttendanceColorInfo(checkInDate);

    let currentProject = data.currentProject || '';
    if (!currentProject) {
      const employeeRes = getEmployee(email);
      if (employeeRes.success && employeeRes.employee) {
        currentProject = employeeRes.employee.currentProject || '';
      }
    }

    const aHdr = getAttHeaders(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      const emailIndex = aHdr['email'] !== undefined ? aHdr['email'] : 2;
      const dateIndex = aHdr['date'] !== undefined ? aHdr['date'] : 6;
      const checkInIndex = aHdr['checkintime'] !== undefined ? aHdr['checkintime'] : 7;
      const colorIndex = aHdr['attendancecolor'] !== undefined ? aHdr['attendancecolor'] : -1;

      for (let i = 0; i < rows.length; i++) {
        const rowEmail = normalizeEmail(rows[i][emailIndex]);
        const rowDate = formatSheetDate(rows[i][dateIndex]) || String(rows[i][dateIndex] || '').substring(0, 10);
        if (rowEmail === email && rowDate === today) {
          const recordedTime = rows[i][checkInIndex] || timeStr;
          const recordedColor = colorIndex >= 0 ? rows[i][colorIndex] : '';
          return {
            success: true,
            message: 'Already marked present today',
            time: recordedTime,
            date: today,
            attendanceColor: getAttendanceColorFromRecordedValue(recordedColor, recordedTime)
          };
        }
      }
    }

    const totalCols = sheet.getLastColumn();
    const newRow = new Array(totalCols).fill('');
    if (aHdr['recordid'] !== undefined) newRow[aHdr['recordid']] = 'ATT' + Date.now();
    if (aHdr['empid'] !== undefined) newRow[aHdr['empid']] = data.id || data.empId || '';
    if (aHdr['email'] !== undefined) newRow[aHdr['email']] = email;
    if (aHdr['name'] !== undefined) newRow[aHdr['name']] = data.name || '';
    if (aHdr['department'] !== undefined) newRow[aHdr['department']] = data.department || '';
    if (aHdr['currentproject'] !== undefined) newRow[aHdr['currentproject']] = currentProject;
    if (aHdr['date'] !== undefined) newRow[aHdr['date']] = today;
    if (aHdr['checkintime'] !== undefined) newRow[aHdr['checkintime']] = timeStr;
    if (aHdr['location'] !== undefined) newRow[aHdr['location']] = data.location || 'Office';
    if (aHdr['status'] !== undefined) newRow[aHdr['status']] = 'Present';
    if (aHdr['attendancecolor'] !== undefined) newRow[aHdr['attendancecolor']] = colorInfo.attendanceColor;

    const rowNumber = sheet.getLastRow() + 1;
    sheet.getRange(rowNumber, 1, 1, totalCols).setValues([newRow]);
    applyAttendanceRowStyle(sheet, rowNumber, totalCols, colorInfo, aHdr['checkintime']);

    return {
      success: true,
      message: 'Attendance marked successfully',
      time: timeStr,
      date: today,
      attendanceColor: colorInfo.attendanceColor
    };
  } catch (err) {
    return { success: false, message: 'Error: ' + err.toString() };
  }
}

function markAttendance(data) {
  return handleMarkAttendance(data);
}

function checkLoginBeforeCutoff(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Attendance');
  if (!sheet) {
    return { success: false, loggedBefore15: false };
  }

  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  let loggedBefore15 = false;
  let loggedToday = false;
  let latestCheckIn = '';

  if (sheet.getLastRow() > 1) {
    const lastRow = sheet.getLastRow();
    const existing = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const aHdr = getAttHeaders(sheet);
    const aIMail = aHdr['email'] !== undefined ? aHdr['email'] : 2;
    const aIDate = aHdr['date'] !== undefined ? aHdr['date'] : 6;
    const aICheckIn = aHdr['checkintime'] !== undefined ? aHdr['checkintime'] : 7;

    for (let i = 0; i < existing.length; i++) {
      const rowDate = existing[i][aIDate] ? existing[i][aIDate].toString().substring(0, 10) : '';
      const rowEmail = normalizeEmail(existing[i][aIMail]);
      if (rowEmail === normalizeEmail(data.email) && rowDate === today) {
        loggedToday = true;
        const checkInStr = existing[i][aICheckIn];
        if (checkInStr) latestCheckIn = checkInStr;
        if (checkInStr) {
          const checkInTime = parseTimeTo24Hour(checkInStr);
          if (checkInTime !== null && checkInTime < 15) {
            loggedBefore15 = true;
          }
        }
      }
    }
  }

  return { success: true, loggedBefore15: loggedBefore15, loggedToday: loggedToday, time: latestCheckIn };
}

function captureAbsentUsers(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const empSheet = ss.getSheetByName('Employees');
  const attSheet = ss.getSheetByName('Attendance');
  const absentSheet = ss.getSheetByName('AbsentUsers');
  
  if (!empSheet || !attSheet || !absentSheet) {
    return { success: false, message: 'Required sheets not found' };
  }

  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const recordedOn = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  const empLastRow = empSheet.getLastRow();
  if (empLastRow <= 1) return { success: true, message: 'No employees found', absentCount: 0 };
  
  const empRows = empSheet.getRange(2, 1, empLastRow - 1, empSheet.getLastColumn()).getValues();
  const eHdr = getEmpHeaders(empSheet);
  const iMail = eHdr['email'] !== undefined ? eHdr['email'] : 1;
  const iName = eHdr['name'] !== undefined ? eHdr['name'] : 2;
  const iDept = eHdr['department'] !== undefined ? eHdr['department'] : 4;
  const iEmpId = eHdr['empid'] !== undefined ? eHdr['empid'] : 0;
  const attLastRow = attSheet.getLastRow();
  const attendedEmails = new Set();
  
  if (attLastRow > 1) {
    const attRows = attSheet.getRange(2, 1, attLastRow - 1, attSheet.getLastColumn()).getValues();
    const aHdr = getAttHeaders(attSheet);
    const aIMail = aHdr['email'] !== undefined ? aHdr['email'] : 2;
    const aIDate = aHdr['date'] !== undefined ? aHdr['date'] : 6;
    
    for (let i = 0; i < attRows.length; i++) {
      const rowDate = attRows[i][aIDate] ? attRows[i][aIDate].toString().substring(0, 10) : '';
      if (rowDate === today) {
        attendedEmails.add(normalizeEmail(attRows[i][aIMail]));
      }
    }
  }
  let absentCount = 0;
  const absHdr = getAttHeaders(absentSheet);
  const totalAbsentCols = absentSheet.getLastColumn();
  
  for (let i = 0; i < empRows.length; i++) {
    const empEmail = normalizeEmail(empRows[i][iMail]);
    // Get and parse join date
    let joinDateStr = '';
    if (eHdr['joindate'] !== undefined) {
      joinDateStr = empRows[i][eHdr['joindate']];
    }
    let joinDate = null;
    if (joinDateStr) {
      // Try to parse joinDateStr as yyyy-MM-dd
      joinDate = new Date(joinDateStr);
      if (isNaN(joinDate.getTime())) {
        // Try alternate parsing if needed
        joinDate = new Date(joinDateStr.replace(/\//g, '-'));
      }
    }
    // If joinDate is after today, skip absent marking
    if (joinDate && joinDate > new Date(today)) {
      continue;
    }
    if (!attendedEmails.has(empEmail)) {
      const absentLastRow = absentSheet.getLastRow();
      let alreadyRecorded = false;
      if (absentLastRow > 1) {
        const absentRows = absentSheet.getRange(2, 1, absentLastRow - 1, absentSheet.getLastColumn()).getValues();
        const absMailIdx = absHdr['email'] !== undefined ? absHdr['email'] : 2;
        const absDateIdx = absHdr['date'] !== undefined ? absHdr['date'] : 0;
        for (let j = 0; j < absentRows.length; j++) {
          const absDate = absentRows[j][absDateIdx] ? absentRows[j][absDateIdx].toString().substring(0, 10) : '';
          if (absDate === today && normalizeEmail(absentRows[j][absMailIdx]) === empEmail) {
            alreadyRecorded = true;
            break;
          }
        }
      }
      if (!alreadyRecorded) {
        const newAbsentRow = new Array(totalAbsentCols).fill('');
        if (absHdr['date'] !== undefined) newAbsentRow[absHdr['date']] = today;
        if (absHdr['empid'] !== undefined) newAbsentRow[absHdr['empid']] = empRows[i][iEmpId] || '';
        if (absHdr['email'] !== undefined) newAbsentRow[absHdr['email']] = empRows[i][iMail] || '';
        if (absHdr['name'] !== undefined) newAbsentRow[absHdr['name']] = empRows[i][iName] || '';
        if (absHdr['department'] !== undefined) newAbsentRow[absHdr['department']] = empRows[i][iDept] || '';
        if (absHdr['status'] !== undefined) newAbsentRow[absHdr['status']] = 'Absent';
        if (absHdr['recordedon'] !== undefined) newAbsentRow[absHdr['recordedon']] = recordedOn;
        absentSheet.appendRow(newAbsentRow);
        absentCount++;
      }
    }
  }
  
  return { success: true, message: 'Absent users captured', absentCount: absentCount };
}

function handleRecordAbsent(data) {
  return recordAbsentEmployees();
}

function recordAbsentEmployees() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    ensureEmployeeSchema(ss);
    ensureAttendanceSchema(ss);
    ensureLeaveSchema(ss);
    ensureAbsenceSchema(ss);

    const empSheet = ss.getSheetByName('Employees');
    const attSheet = ss.getSheetByName('Attendance');
    const leaveSheet = ss.getSheetByName('Leaves');
    const absenceSheet = ss.getSheetByName('Absences');

    if (!empSheet || empSheet.getLastRow() < 2) {
      return { success: true, recorded: 0 };
    }

    const tz = Session.getScriptTimeZone();
    const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const empRows = empSheet.getRange(2, 1, empSheet.getLastRow() - 1, empSheet.getLastColumn()).getValues();
    const eHdr = getEmpHeaders(empSheet);

    const presentEmails = new Set();
    if (attSheet && attSheet.getLastRow() > 1) {
      const attRows = attSheet.getRange(2, 1, attSheet.getLastRow() - 1, attSheet.getLastColumn()).getValues();
      const aHdr = getAttHeaders(attSheet);
      const attEmailIndex = aHdr['email'] !== undefined ? aHdr['email'] : 2;
      const attDateIndex = aHdr['date'] !== undefined ? aHdr['date'] : 6;
      for (let i = 0; i < attRows.length; i++) {
        const rowDate = formatSheetDate(attRows[i][attDateIndex]) || String(attRows[i][attDateIndex] || '').substring(0, 10);
        if (rowDate === today) {
          presentEmails.add(normalizeEmail(attRows[i][attEmailIndex]));
        }
      }
    }

    const onLeaveEmails = new Set();
    if (leaveSheet && leaveSheet.getLastRow() > 1) {
      const leaveRows = leaveSheet.getRange(2, 1, leaveSheet.getLastRow() - 1, leaveSheet.getLastColumn()).getValues();
      const lHdr = getLeaveHeaders(leaveSheet);
      for (let i = 0; i < leaveRows.length; i++) {
        const status = normalizeLeaveStatus(getValueByHeader(leaveRows[i], lHdr, 'status', 10));
        if (status !== 'Approved') continue;

        const fromDate = formatSheetDate(getValueByHeader(leaveRows[i], lHdr, 'fromdate', 6));
        const toDate = formatSheetDate(getValueByHeader(leaveRows[i], lHdr, 'todate', 7));
        if (fromDate && toDate && today >= fromDate && today <= toDate) {
          onLeaveEmails.add(normalizeEmail(getValueByHeader(leaveRows[i], lHdr, 'email', 2)));
        }
      }
    }

    const alreadyAbsent = new Set();
    if (absenceSheet && absenceSheet.getLastRow() > 1) {
      const existingAbsences = absenceSheet.getRange(2, 1, absenceSheet.getLastRow() - 1, absenceSheet.getLastColumn()).getValues();
      for (let i = 0; i < existingAbsences.length; i++) {
        const rowDate = formatSheetDate(existingAbsences[i][0]) || String(existingAbsences[i][0] || '').substring(0, 10);
        if (rowDate === today) {
          alreadyAbsent.add(normalizeEmail(existingAbsences[i][2]));
        }
      }
    }

    const absentRows = [];
    for (let i = 0; i < empRows.length; i++) {
      const email = normalizeEmail(getValueByHeader(empRows[i], eHdr, 'email', 1));
      const status = (getValueByHeader(empRows[i], eHdr, 'status', -1) || 'Active').toString().trim().toLowerCase();
      const joinDate = formatSheetDate(getValueByHeader(empRows[i], eHdr, 'joindate', -1));
      if (!email) continue;
      if (status !== 'active') continue;
      if (joinDate && joinDate > today) continue;
      if (presentEmails.has(email)) continue;
      if (onLeaveEmails.has(email)) continue;
      if (alreadyAbsent.has(email)) continue;

      const employmentType = getValueByHeader(empRows[i], eHdr, 'employmenttype', -1) ||
        getValueByHeader(empRows[i], eHdr, 'type', -1) || 'Permanent';

      absentRows.push([
        today,
        getValueByHeader(empRows[i], eHdr, 'empid', 0),
        email,
        getValueByHeader(empRows[i], eHdr, 'name', 2),
        getValueByHeader(empRows[i], eHdr, 'department', 4),
        normalizeProductivityTrackerLabel(getValueByHeader(empRows[i], eHdr, 'role', 3)),
        getValueByHeader(empRows[i], eHdr, 'reportingmanager', -1),
        getValueByHeader(empRows[i], eHdr, 'manager', -1),
        getValueByHeader(empRows[i], eHdr, 'currentproject', -1),
        getValueByHeader(empRows[i], eHdr, 'workmode', -1),
        employmentType,
        'Absent'
      ]);
    }

    if (absentRows.length > 0) {
      const startRow = absenceSheet.getLastRow() + 1;
      absenceSheet.getRange(startRow, 1, absentRows.length, ABSENCE_HEADERS.length).setValues(absentRows);
      absenceSheet.getRange(startRow, 1, absentRows.length, ABSENCE_HEADERS.length).setBackground('#f8d7da');
    }

    return { success: true, recorded: absentRows.length };
  } catch (err) {
    Logger.log('recordAbsentEmployees error: ' + err.toString());
    return { success: false, message: 'Error: ' + err.toString(), error: err.toString() };
  }
}

function setupDailyAbsentTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === 'recordAbsentEmployees') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('recordAbsentEmployees')
    .timeBased()
    .atHour(23)
    .nearMinute(59)
    .everyDays(1)
    .create();

  Logger.log('Daily absent recording trigger set up successfully');
  return { success: true, message: 'Daily absent recording trigger set up successfully' };
}

function getAttendance(email) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ensureAttendanceSchema(ss);
  const sheet = ss.getSheetByName('Attendance');
  if (!sheet) return { success: true, attendance: [] };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, attendance: [] };

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getAttHeaders(sheet);

  const iMail = hdr['email'] !== undefined ? hdr['email'] : 2;
  const iDate = hdr['date'] !== undefined ? hdr['date'] : 6;
  const iCIn = hdr['checkintime'] !== undefined ? hdr['checkintime'] : 7;
  const iLoc = hdr['location'] !== undefined ? hdr['location'] : 8;
  const iSts = hdr['status'] !== undefined ? hdr['status'] : 9;
  const iProj = hdr['currentproject'] !== undefined ? hdr['currentproject'] : -1;
  const iRec = hdr['recordid'] !== undefined ? hdr['recordid'] : 0;
  const iColor = hdr['attendancecolor'] !== undefined ? hdr['attendancecolor'] : -1;

  const wantedEmail = normalizeEmail(email);
  const result = [];
  for (let i = 0; i < rows.length; i++) {
    if (normalizeEmail(rows[i][iMail]) === wantedEmail) {
      result.push({
        recordId: rows[i][iRec] || '',
        date: formatSheetDate(rows[i][iDate]) || '',
        checkIn: rows[i][iCIn] || '',
        location: rows[i][iLoc] || 'Office',
        status: rows[i][iSts] || 'Present',
        currentProject: iProj >= 0 ? (rows[i][iProj] || '') : '',
        attendanceColor: iColor >= 0 ? (rows[i][iColor] || '') : ''
      });
    }
  }

  return { success: true, attendance: result };
}

function getAllAttendance() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ensureAttendanceSchema(ss);
  const sheet = ss.getSheetByName('Attendance');
  if (!sheet || sheet.getLastRow() < 2) return { success: true, attendance: [] };

  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getAttHeaders(sheet);

  const iRec = hdr['recordid'] !== undefined ? hdr['recordid'] : 0;
  const iEId = hdr['empid'] !== undefined ? hdr['empid'] : 1;
  const iMail = hdr['email'] !== undefined ? hdr['email'] : 2;
  const iName = hdr['name'] !== undefined ? hdr['name'] : 3;
  const iDept = hdr['department'] !== undefined ? hdr['department'] : 4;
  const iProj = hdr['currentproject'] !== undefined ? hdr['currentproject'] : -1;
  const iDate = hdr['date'] !== undefined ? hdr['date'] : 6;
  const iCIn = hdr['checkintime'] !== undefined ? hdr['checkintime'] : 7;
  const iLoc = hdr['location'] !== undefined ? hdr['location'] : 8;
  const iSts = hdr['status'] !== undefined ? hdr['status'] : 9;
  const iColor = hdr['attendancecolor'] !== undefined ? hdr['attendancecolor'] : -1;

  const result = [];
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][iRec] || rows[i][iMail]) {
      result.push({
        recordId: rows[i][iRec] || '',
        empId: rows[i][iEId] || '',
        email: rows[i][iMail] || '',
        name: rows[i][iName] || '',
        department: rows[i][iDept] || '',
        currentProject: iProj >= 0 ? (rows[i][iProj] || '') : '',
        date: formatSheetDate(rows[i][iDate]) || '',
        checkIn: rows[i][iCIn] || '',
        location: rows[i][iLoc] || 'Office',
        status: rows[i][iSts] || 'Present',
        attendanceColor: iColor >= 0 ? (rows[i][iColor] || '') : ''
      });
    }
  }

  return { success: true, attendance: result };
}

function applyLeave(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Leaves');
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName('Leaves');
  }
  ensureLeaveSchema(ss);
  sheet = ss.getSheetByName('Leaves');

  const employeeEmail = normalizeEmail(data.email);
  if (!employeeEmail) return { success: false, message: 'Employee email is required' };

  const leaveId = 'LV' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss');
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  let days = 1;
  let employee = null;

  try {
    days = Math.max(1, Math.ceil((new Date(data.toDate) - new Date(data.fromDate)) / 86400000) + 1);
  } catch (e) {}

  try {
    const empData = getEmployee(employeeEmail);
    if (empData.success && empData.employee) {
      employee = empData.employee;
    }
  } catch (lookupErr) {
    Logger.log('Employee lookup failed during applyLeave: ' + lookupErr.toString());
  }

  // Check if employee has completed leaves (already used all available leaves)
  let leaveWarning = '';
  let totalApprovedDays = 0;
  let totalPendingDays = 0;
  
  try {
    if (sheet.getLastRow() > 1) {
      const lastRow = sheet.getLastRow();
      const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      const hdr = getLeaveHeaders(sheet);
      
      for (let i = 0; i < rows.length; i++) {
        const rowEmail = normalizeEmail(getValueByHeader(rows[i], hdr, 'email', 2));
        const rowStatus = normalizeLeaveStatus(getValueByHeader(rows[i], hdr, 'status', 10));
        const rowDays = parseInt(getValueByHeader(rows[i], hdr, 'days', 8)) || 0;
        
        if (rowEmail === employeeEmail) {
          if (rowStatus === 'Approved') {
            totalApprovedDays += rowDays;
          } else if (rowStatus === 'Pending') {
            totalPendingDays += rowDays;
          }
        }
      }
    }
    
    // Check if leaves are exhausted (assuming 30 days annual limit)
    const ANNUAL_LEAVE_LIMIT = 30;
    const totalUsedAndPending = totalApprovedDays + totalPendingDays + days;
    
    if (totalUsedAndPending > ANNUAL_LEAVE_LIMIT) {
      leaveWarning = '⚠️ WARNING: You have completed/exhausted your annual leaves (' + ANNUAL_LEAVE_LIMIT + ' days). ' +
                     'Your salary will be DEDUCTED for this leave application. ' +
                     'Total approved: ' + totalApprovedDays + ' days, Pending: ' + totalPendingDays + ' days, Current request: ' + days + ' days.';
      Logger.log('Leave quota exceeded for ' + employeeEmail + ': ' + leaveWarning);
    }
  } catch (checkErr) {
    Logger.log('Error checking leave balance: ' + checkErr.toString());
  }

  const leavePayload = {
    empId: data.empId || (employee ? employee.id : '') || '',
    email: employeeEmail,
    name: data.name || (employee ? employee.name : '') || '',
    department: data.department || (employee ? employee.department : '') || '',
    leaveType: data.leaveType || '',
    fromDate: data.fromDate || '',
    toDate: data.toDate || '',
    reason: data.reason || '',
    manager: data.manager || (employee ? employee.manager : '') || '',
    managerEmail: normalizeEmail(data.managerEmail || (employee ? employee.managerEmail : '') || ''),
    reportingManager: data.reportingManager || (employee ? employee.reportingManager : '') || '',
    reportingManagerEmail: normalizeEmail(data.reportingManagerEmail || (employee ? employee.reportingManagerEmail : '') || ''),
    warning: leaveWarning || null,
    leavesSummary: {
      approvedDays: totalApprovedDays,
      pendingDays: totalPendingDays,
      newRequestDays: days,
      totalUsedAndPending: totalApprovedDays + totalPendingDays + days,
      annualLimit: 30
    }
  };

  const hdr = getLeaveHeaders(sheet);
  const newRow = [];
  for (let col = 0; col < sheet.getLastColumn(); col++) newRow[col] = '';
  if (hdr['leaveid'] !== undefined) newRow[hdr['leaveid']] = leaveId;
  if (hdr['empid'] !== undefined) newRow[hdr['empid']] = leavePayload.empId;
  if (hdr['email'] !== undefined) newRow[hdr['email']] = leavePayload.email;
  if (hdr['name'] !== undefined) newRow[hdr['name']] = leavePayload.name;
  if (hdr['department'] !== undefined) newRow[hdr['department']] = leavePayload.department;
  if (hdr['leavetype'] !== undefined) newRow[hdr['leavetype']] = leavePayload.leaveType;
  if (hdr['fromdate'] !== undefined) newRow[hdr['fromdate']] = leavePayload.fromDate;
  if (hdr['todate'] !== undefined) newRow[hdr['todate']] = leavePayload.toDate;
  if (hdr['days'] !== undefined) newRow[hdr['days']] = days;
  if (hdr['reason'] !== undefined) newRow[hdr['reason']] = leavePayload.reason;
  if (hdr['status'] !== undefined) newRow[hdr['status']] = 'Pending';
  if (hdr['appliedon'] !== undefined) newRow[hdr['appliedon']] = now;
  if (hdr['updatedon'] !== undefined) newRow[hdr['updatedon']] = now;
  if (hdr['reportingmanager'] !== undefined) newRow[hdr['reportingmanager']] = leavePayload.reportingManager;
  if (hdr['reportingmanageremail'] !== undefined) newRow[hdr['reportingmanageremail']] = leavePayload.reportingManagerEmail;
  if (hdr['manager'] !== undefined) newRow[hdr['manager']] = leavePayload.manager;
  if (hdr['manageremail'] !== undefined) newRow[hdr['manageremail']] = leavePayload.managerEmail;
  sheet.appendRow(newRow);

  let notificationResult = { ok: false, sent: [], issues: [] };
  try {
    if (employee) {
      notificationResult = sendLeaveApplicationEmail(employee, leavePayload, leaveId);
    } else {
      notificationResult.issues.push('Employee record not found for leave notification: ' + employeeEmail);
      Logger.log(notificationResult.issues[notificationResult.issues.length - 1]);
    }
  } catch (emailErr) {
    notificationResult.issues.push('Email send error: ' + emailErr.toString());
    Logger.log(notificationResult.issues[notificationResult.issues.length - 1]);
  }

  const resultMessage = leaveWarning 
    ? 'Leave applied successfully! ' + leaveWarning
    : 'Leave applied successfully!';

  return {
    success: true,
    message: resultMessage,
    warning: leaveWarning || null,
    id: leaveId,
    leavesSummary: {
      approvedDays: totalApprovedDays,
      pendingDays: totalPendingDays,
      newRequestDays: days,
      totalUsedAndPending: totalApprovedDays + totalPendingDays + days,
      annualLimit: 30
    },
    notification: notificationResult
  };
}

function getLeaves(email) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ensureLeaveSchema(ss);
  const sheet = ss.getSheetByName('Leaves');
  if (!sheet) return { success: true, leaves: [] };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, leaves: [] };

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getLeaveHeaders(sheet);
  const wantedEmail = normalizeEmail(email);
  let employee = null;
  try {
    const empData = getEmployee(wantedEmail);
    if (empData.success && empData.employee) employee = empData.employee;
  } catch (ignored) {}

  const result = [];
  for (let i = 0; i < rows.length; i++) {
    const rowEmail = normalizeEmail(getValueByHeader(rows[i], hdr, 'email', 2));
    if (rowEmail === wantedEmail) {
      result.push(attachEmployeeManagersToLeave(buildLeaveObject(rows[i], hdr), employee));
    }
  }

  return { success: true, leaves: result };
}

function getAllLeaves() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ensureLeaveSchema(ss);
  const sheet = ss.getSheetByName('Leaves');
  if (!sheet || sheet.getLastRow() < 2) return { success: true, leaves: [] };

  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getLeaveHeaders(sheet);
  const employeesByEmail = {};
  try {
    const empData = getEmployees();
    if (empData.success && empData.employees) {
      empData.employees.forEach(function(emp) {
        const key = normalizeEmail(emp.email);
        if (key) employeesByEmail[key] = emp;
      });
    }
  } catch (ignored) {}

  const result = [];
  for (let i = 0; i < rows.length; i++) {
    if (getValueByHeader(rows[i], hdr, 'leaveid', 0)) {
      const leave = buildLeaveObject(rows[i], hdr);
      result.push(attachEmployeeManagersToLeave(leave, employeesByEmail[normalizeEmail(leave.email)]));
    }
  }

  return { success: true, leaves: result };
}

function updateLeaveStatus(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ensureLeaveSchema(ss);
  const sheet = ss.getSheetByName('Leaves');
  if (!sheet) return { success: false, message: 'Sheet not found' };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'No leaves found' };

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getLeaveHeaders(sheet);
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const requestedId = (data.id || '').toString().trim();
  if (!requestedId) return { success: false, message: 'Leave ID is required' };
  const requestedStatus = normalizeLeaveStatus(data.status || 'Pending');
  const requestedReason = (data.rejectionReason || '').toString().trim();

  if ((requestedStatus === 'Rejected' || requestedStatus === 'Cancelled') && !requestedReason) {
    return { success: false, message: 'Rejection reason is required' };
  }

  for (let i = 0; i < rows.length; i++) {
    const leaveId = getValueByHeader(rows[i], hdr, 'leaveid', 0);
    if (leaveId && leaveId.toString().trim() === requestedId) {
      const r = i + 2;
      const currentStatus = normalizeLeaveStatus(getValueByHeader(rows[i], hdr, 'status', 10));
      const newStatus = normalizeLeaveStatus(data.status || 'Pending');
      const actorRole = (data.actorRole || '').toString().trim().toUpperCase();

      if (currentStatus !== 'Pending' && newStatus !== currentStatus) {
        const canCancelApproved = currentStatus === 'Approved' && actorRole === 'HR' && newStatus === 'Cancelled';
        if (!canCancelApproved) {
          return { success: false, message: 'This leave request is already processed and cannot be changed' };
        }
      }

      if (hdr['status'] !== undefined) sheet.getRange(r, hdr['status'] + 1).setValue(newStatus);
      if (hdr['updatedon'] !== undefined) sheet.getRange(r, hdr['updatedon'] + 1).setValue(now);

      if (newStatus === 'Approved') {
        if (hdr['approvedby'] !== undefined) {
          sheet.getRange(r, hdr['approvedby'] + 1).setValue(data.approvedBy || '');
        }
        if (hdr['approveddate'] !== undefined) {
          sheet.getRange(r, hdr['approveddate'] + 1).setValue(now);
        }
        if (hdr['rejectedby'] !== undefined) sheet.getRange(r, hdr['rejectedby'] + 1).setValue('');
        if (hdr['rejecteddate'] !== undefined) sheet.getRange(r, hdr['rejecteddate'] + 1).setValue('');
        if (hdr['rejectionreason'] !== undefined) sheet.getRange(r, hdr['rejectionreason'] + 1).setValue('');
      } else if (newStatus === 'Rejected') {
        if (hdr['rejectedby'] !== undefined) {
          sheet.getRange(r, hdr['rejectedby'] + 1).setValue(data.rejectedBy || '');
        }
        if (hdr['rejecteddate'] !== undefined) {
          sheet.getRange(r, hdr['rejecteddate'] + 1).setValue(now);
        }
        if (hdr['rejectionreason'] !== undefined) {
          sheet.getRange(r, hdr['rejectionreason'] + 1).setValue(data.rejectionReason || '');
        }
        if (hdr['approvedby'] !== undefined) sheet.getRange(r, hdr['approvedby'] + 1).setValue('');
        if (hdr['approveddate'] !== undefined) sheet.getRange(r, hdr['approveddate'] + 1).setValue('');
      } else if (newStatus === 'Cancelled') {
        if (hdr['rejectedby'] !== undefined) {
          sheet.getRange(r, hdr['rejectedby'] + 1).setValue(data.rejectedBy || '');
        }
        if (hdr['rejecteddate'] !== undefined) {
          sheet.getRange(r, hdr['rejecteddate'] + 1).setValue(now);
        }
        if (hdr['rejectionreason'] !== undefined) {
          sheet.getRange(r, hdr['rejectionreason'] + 1).setValue(data.rejectionReason || 'Cancelled by HR');
        }
        if (hdr['approvedby'] !== undefined) sheet.getRange(r, hdr['approvedby'] + 1).setValue('');
        if (hdr['approveddate'] !== undefined) sheet.getRange(r, hdr['approveddate'] + 1).setValue('');
      }

      const color = newStatus === 'Approved'
        ? '#d4edda'
        : newStatus === 'Rejected'
          ? '#f8d7da'
          : newStatus === 'Cancelled'
            ? '#fde68a'
            : '#fff3cd';

      const totalCols = sheet.getLastColumn();
      sheet.getRange(r, 1, 1, totalCols).setBackground(color);
      SpreadsheetApp.flush();

      let notificationResult = { ok: false, sent: [], issues: [] };
      try {
        const freshRow = sheet.getRange(r, 1, 1, sheet.getLastColumn()).getValues()[0];
        const leaveRecord = buildLeaveObject(freshRow, hdr);
        leaveRecord.status = newStatus;
        leaveRecord.approvedBy = data.approvedBy || leaveRecord.approvedBy || '';
        leaveRecord.rejectedBy = data.rejectedBy || leaveRecord.rejectedBy || '';
        leaveRecord.rejectionReason = data.rejectionReason || leaveRecord.rejectionReason || '';
        leaveRecord.updatedOn = now;
        const employeeForMail = getEmployeeForLeaveNotification(leaveRecord);
        notificationResult = sendLeaveStatusUpdateEmail(employeeForMail, leaveRecord);
      } catch (emailErr) {
        notificationResult.issues.push('Leave status email error: ' + emailErr.toString());
        Logger.log(notificationResult.issues[notificationResult.issues.length - 1]);
      }

      const employeeMailSent = notificationResult.sent.some(function(item) {
        return item.indexOf('employee:') === 0;
      });
      const messageSuffix = employeeMailSent ? '' : ' Status updated, but employee email was not sent.';

      return {
        success: true,
        message: 'Leave ' + newStatus.toLowerCase() + ' successfully!' + messageSuffix,
        notification: notificationResult
      };
    }
  }

  return { success: false, message: 'Leave record not found' };
}

function getProjects() {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Projects');

    if (!sheet) {
      sheet = ss.insertSheet('Projects');
      sheet.getRange(1, 1).setValue('ProjectName')
        .setBackground('#1a1a2e')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, projects: [] };

    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    const projects = [];

    for (let i = 0; i < data.length; i++) {
      const val = data[i][0] ? data[i][0].toString().trim() : '';
      if (val) projects.push(val);
    }

    return { success: true, projects: projects };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function saveProjects(data) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName('Projects');
    if (!sheet) sheet = ss.insertSheet('Projects');

    let projects = [];
    try {
      projects = JSON.parse(data.projects || '[]');
    } catch (pe) {
      projects = (data.projects || '')
        .split(',')
        .map(function(p) { return p.trim(); })
        .filter(Boolean);
    }

    if (!Array.isArray(projects) || projects.length === 0) {
      return { success: false, message: 'No projects provided' };
    }

    sheet.clearContents();
    sheet.getRange(1, 1).setValue('ProjectName')
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
    sheet.setFrozenRows(1);

    const cleanProjects = projects
      .map(function(p) { return p ? p.toString().trim() : ''; })
      .filter(Boolean);

    if (cleanProjects.length > 0) {
      const writeData = cleanProjects.map(function(p) { return [p]; });
      sheet.getRange(2, 1, writeData.length, 1).setValues(writeData);
    }

    return {
      success: true,
      message: cleanProjects.length + ' projects saved successfully'
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function assignProject(data) {
  try {
    const email = normalizeEmail(data.email);
    const project = (data.project || '').trim();

    if (!email) return { success: false, message: 'Email is required' };
    if (!project) return { success: false, message: 'Project is required' };

    const ss = SpreadsheetApp.openById(SHEET_ID);
    ensureEmployeeSchema(ss);
    const sheet = ss.getSheetByName('Employees');
    if (!sheet) return { success: false, message: 'Employees sheet not found' };

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: false, message: 'No employees found' };

    const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const hdr = getEmpHeaders(sheet);
    const iMail = hdr['email'] !== undefined ? hdr['email'] : 1;
    const iProj = hdr['currentproject'] !== undefined ? hdr['currentproject'] : -1;

    if (iProj === -1) return { success: false, message: 'CurrentProject column not found' };

    for (let i = 0; i < rows.length; i++) {
      if (normalizeEmail(rows[i][iMail]) === email) {
        sheet.getRange(i + 2, iProj + 1).setValue(project);
        return { success: true, message: 'Project assigned successfully' };
      }
    }

    return { success: false, message: 'Employee not found' };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function findManagerEmailByReference(referenceValue, managerType) {
  const ref = normalizeEmail(referenceValue);
  const refText = (referenceValue || '').toString().trim().toLowerCase();
  if (!ref && !refText) return '';

  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Managers');
  if (!sheet || sheet.getLastRow() < 2) {
    Logger.log('findManagerEmailByReference: Managers sheet not found or empty');
    return '';
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const hdr = getManagerHeaders(sheet);
  const wantedType = normalizeProductivityTrackerLabel(managerType).toLowerCase();

  for (let i = 0; i < rows.length; i++) {
    const rowType = normalizeProductivityTrackerLabel(getValueByHeader(rows[i], hdr, 'managertype', 1)).toLowerCase();
    const rowName = (getValueByHeader(rows[i], hdr, 'name', 2) || '').toString().trim().toLowerCase();
    const rowEmail = normalizeEmail(getValueByHeader(rows[i], hdr, 'email', 3));
    
    // Log for debugging
    if ((ref && rowEmail) || refText) {
      Logger.log('Manager lookup: ref=' + ref + ', rowEmail=' + rowEmail + ', refText=' + refText + ', rowName=' + rowName);
    }
    
    if (wantedType && rowType && rowType !== wantedType) continue;
    if ((ref && rowEmail === ref) || (refText && rowName === refText)) {
      Logger.log('Manager found: ' + rowEmail);
      return rowEmail || '';
    }
  }

  Logger.log('No manager found for reference: ' + referenceValue + ', type: ' + managerType);
  return '';
}

function sendLeaveApplicationEmail(employee, data, leaveId) {
  const result = { ok: false, sent: [], issues: [] };
  try {
    const empEmail = normalizeEmail(employee.email);
    const empName = data.name || employee.name || 'Employee';
    const leaveType = data.leaveType || 'Leave';
    const fromDate = data.fromDate || '';
    const toDate = data.toDate || '';
    const reason = data.reason || 'No reason provided';
    
    // Better manager email resolution with multiple fallbacks
    let managerEmail = '';
    if (data.managerEmail) {
      managerEmail = normalizeEmail(data.managerEmail);
    } else if (employee.managerEmail) {
      managerEmail = normalizeEmail(employee.managerEmail);
    } else if (data.manager) {
      // Try to find manager by name from Managers sheet
      managerEmail = findManagerEmailByReference(data.manager, 'manager') || 
                     findManagerEmailByReference(data.manager, '');
    } else if (employee.manager) {
      // Try to find manager by name from Managers sheet
      managerEmail = findManagerEmailByReference(employee.manager, 'manager') || 
                     findManagerEmailByReference(employee.manager, '');
    }
    
    const managerName = data.manager || employee.manager || 'Manager';
    Logger.log('Leave application: Employee=' + empEmail + ', Manager=' + managerEmail);
    
    // Better reporting manager email resolution
    let reportingManagerEmail = '';
    if (data.reportingManagerEmail) {
      reportingManagerEmail = normalizeEmail(data.reportingManagerEmail);
    } else if (employee.reportingManagerEmail) {
      reportingManagerEmail = normalizeEmail(employee.reportingManagerEmail);
    } else if (data.reportingManager) {
      reportingManagerEmail = findManagerEmailByReference(data.reportingManager, 'reporting') || 
                              findManagerEmailByReference(data.reportingManager, '');
    } else if (employee.reportingManager) {
      reportingManagerEmail = findManagerEmailByReference(employee.reportingManager, 'reporting') || 
                              findManagerEmailByReference(employee.reportingManager, '');
    }
    
    const reportingManagerName = data.reportingManager || employee.reportingManager || 'Reporting Manager';
    Logger.log('Reporting Manager lookup: ' + reportingManagerEmail);

    // Build warning section if present
    let warningHtml = '';
    if (data.warning) {
      warningHtml = '<div style="background-color: #fff3cd; padding: 15px; border-left: 4px solid #ff9800; margin: 15px 0; border-radius: 4px;">' +
        '<p style="color: #856404; margin: 0;"><strong>⚠️ WARNING:</strong></p>' +
        '<p style="color: #856404; margin: 5px 0 0 0;">' + data.warning.replace(/⚠️ WARNING: /g, '') + '</p>' +
        '</div>';
    }

    const emailBody = '<html>' +
      '<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">' +
      '<div style="max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f9f9f9; border-radius: 8px;">' +
      '<h2 style="color: #667eea; margin-top: 0;">Leave Application Notification</h2>' +
      '<p>Dear ' + empName + ',</p>' +
      '<p style="color: #666;">You have successfully applied for leave. Here are the details:</p>' +
      '<div style="background-color: #fff; padding: 15px; border-left: 4px solid #667eea; margin: 15px 0;">' +
      '<p><strong>Leave ID:</strong> ' + leaveId + '</p>' +
      '<p><strong>Leave Type:</strong> ' + leaveType + '</p>' +
      '<p><strong>From Date:</strong> ' + fromDate + '</p>' +
      '<p><strong>To Date:</strong> ' + toDate + '</p>' +
      '<p><strong>Reason:</strong> ' + reason + '</p>' +
      '<p><strong>Status:</strong> <span style="color: #f59e0b; font-weight: bold;">Pending Approval</span></p>' +
      '</div>' +
      warningHtml +
      '<p style="color: #666; font-size: 0.9em;">Your leave request has been submitted for approval. Your manager will review and respond shortly.</p>' +
      '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
      '<p style="color: #999; font-size: 0.85em;">AttendPro System | Do not reply to this email</p>' +
      '</div>' +
      '</body>' +
      '</html>';
    
    // Send email to employee
    if (empEmail) {
      try {
        MailApp.sendEmail({
          to: empEmail,
          subject: 'Leave Application Confirmation - ' + empName,
          htmlBody: emailBody
        });
        result.sent.push('employee:' + empEmail);
        Logger.log('Employee notification sent to: ' + empEmail);
      } catch (empEmailErr) {
        result.issues.push('Failed to send employee email to ' + empEmail + ': ' + empEmailErr.toString());
        Logger.log(result.issues[result.issues.length - 1]);
      }
    } else {
      result.issues.push('Employee email missing for leave confirmation');
      Logger.log(result.issues[result.issues.length - 1]);
    }
    
    // Send email to manager
    if (managerEmail && managerEmail.indexOf('@') > 0) {
      const managerEmailBody = '<html>' +
        '<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">' +
        '<div style="max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f9f9f9; border-radius: 8px;">' +
        '<h2 style="color: #667eea; margin-top: 0;">Leave Request for Approval</h2>' +
        '<p>Dear ' + managerName + ',</p>' +
        '<p style="color: #666;"><strong>' + empName + '</strong> has applied for ' + leaveType + '. Please review and approve/reject this request.</p>' +
        '<div style="background-color: #fff; padding: 15px; border-left: 4px solid #667eea; margin: 15px 0;">' +
        '<p><strong>Employee Name:</strong> ' + empName + '</p>' +
        '<p><strong>Employee Email:</strong> ' + empEmail + '</p>' +
        '<p><strong>Leave ID:</strong> ' + leaveId + '</p>' +
        '<p><strong>Leave Type:</strong> ' + leaveType + '</p>' +
        '<p><strong>From Date:</strong> ' + fromDate + '</p>' +
        '<p><strong>To Date:</strong> ' + toDate + '</p>' +
        '<p><strong>Reason:</strong> ' + reason + '</p>' +
        '</div>' +
        warningHtml +
        '<p style="color: #666; margin-top: 15px;"><strong>Action Required:</strong> Please visit the AttendPro Manager Portal to approve or reject this request.</p>' +
        '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
        '<p style="color: #999; font-size: 0.85em;">AttendPro System | Do not reply to this email</p>' +
        '</div>' +
        '</body>' +
        '</html>';

      try {
        MailApp.sendEmail({
          to: managerEmail,
          subject: 'Leave Request for Approval - ' + empName,
          htmlBody: managerEmailBody
        });
        result.sent.push('manager:' + managerEmail);
        Logger.log('Manager notification sent to: ' + managerEmail);
      } catch (mgrEmailErr) {
        result.issues.push('Failed to send manager email to ' + managerEmail + ': ' + mgrEmailErr.toString());
        Logger.log(result.issues[result.issues.length - 1]);
      }
    } else {
      result.issues.push('Manager email missing or invalid for leave notification. Employee: ' + empEmail + ', manager ref: ' + (employee.manager || employee.managerEmail || 'Not found') + ', resolved email: ' + managerEmail);
      Logger.log(result.issues[result.issues.length - 1]);
    }
    
    // Send email to reporting manager
    if (reportingManagerEmail && reportingManagerEmail.indexOf('@') > 0 && reportingManagerEmail !== managerEmail) {
      const reportingManagerEmailBody = '<html>' +
        '<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">' +
        '<div style="max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f9f9f9; border-radius: 8px;">' +
        '<h2 style="color: #667eea; margin-top: 0;">Leave Request for Approval</h2>' +
        '<p>Dear ' + reportingManagerName + ',</p>' +
        '<p style="color: #666;"><strong>' + empName + '</strong> has applied for ' + leaveType + '. Please review and approve or reject this request.</p>' +
        '<div style="background-color: #fff; padding: 15px; border-left: 4px solid #667eea; margin: 15px 0;">' +
        '<p><strong>Employee Name:</strong> ' + empName + '</p>' +
        '<p><strong>Employee Email:</strong> ' + empEmail + '</p>' +
        '<p><strong>Leave ID:</strong> ' + leaveId + '</p>' +
        '<p><strong>Leave Type:</strong> ' + leaveType + '</p>' +
        '<p><strong>From Date:</strong> ' + fromDate + '</p>' +
        '<p><strong>To Date:</strong> ' + toDate + '</p>' +
        '<p><strong>Reason:</strong> ' + reason + '</p>' +
        '</div>' +
        warningHtml +
        '<p style="color: #666; margin-top: 15px;"><strong>Action Required:</strong> Please visit the AttendPro Reporting Manager Portal to approve or reject this request.</p>' +
        '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
        '<p style="color: #999; font-size: 0.85em;">AttendPro System | Do not reply to this email</p>' +
        '</div>' +
        '</body>' +
        '</html>';

      try {
        MailApp.sendEmail({
          to: reportingManagerEmail,
          subject: 'Leave Request Notification - ' + empName,
          htmlBody: reportingManagerEmailBody
        });
        result.sent.push('reporting:' + reportingManagerEmail);
        Logger.log('Reporting manager notification sent to: ' + reportingManagerEmail);
      } catch (reportEmailErr) {
        result.issues.push('Failed to send reporting manager email to ' + reportingManagerEmail + ': ' + reportEmailErr.toString());
        Logger.log(result.issues[result.issues.length - 1]);
      }
    } else if (!reportingManagerEmail || reportingManagerEmail.indexOf('@') === -1) {
      result.issues.push('Reporting manager email missing or invalid for leave notification. Employee: ' + empEmail + ', reporting manager ref: ' + (employee.reportingManager || employee.reportingManagerEmail || 'Not found') + ', resolved email: ' + reportingManagerEmail);
      Logger.log(result.issues[result.issues.length - 1]);
    }

    Logger.log('Leave notification emails completed for Leave ID: ' + leaveId + '. Sent: ' + JSON.stringify(result.sent) + ', Issues: ' + JSON.stringify(result.issues));
    result.ok = result.sent.length > 0;
    return result;
  } catch (err) {
    Logger.log('Error sending leave notification emails: ' + err.toString());
    result.issues.push(err.toString());
    return result;
  }
}

function sendLeaveStatusUpdateEmail(employee, leaveRecord) {
  const result = { ok: false, sent: [], issues: [] };
  try {
    const empEmail = normalizeEmail(employee.email);
    if (!empEmail) {
      result.issues.push('Employee email missing for leave status update');
      Logger.log(result.issues[result.issues.length - 1]);
      return result;
    }

    const empName = employee.name || leaveRecord.name || 'Employee';
    const leaveType = leaveRecord.leaveType || 'Leave';
    const status = normalizeLeaveStatus(leaveRecord.status || '');
    const fromDate = leaveRecord.fromDate || '';
    const toDate = leaveRecord.toDate || '';
    const reason = leaveRecord.reason || 'No reason provided';
    const approver = leaveRecord.approvedBy || leaveRecord.rejectedBy || 'Manager/HR';
    const rejectionReason = leaveRecord.rejectionReason || '';
    const statusColor = status === 'Approved' ? '#16a34a' : status === 'Rejected' ? '#dc2626' : '#d97706';
    let managerEmail = '';
    let reportingManagerEmail = '';

    if (employee.managerEmail) {
      managerEmail = normalizeEmail(employee.managerEmail);
    } else if (employee.manager) {
      managerEmail = findManagerEmailByReference(employee.manager, 'manager') ||
        findManagerEmailByReference(employee.manager, '');
    }

    if (employee.reportingManagerEmail) {
      reportingManagerEmail = normalizeEmail(employee.reportingManagerEmail);
    } else if (employee.reportingManager) {
      reportingManagerEmail = findManagerEmailByReference(employee.reportingManager, 'reporting') ||
        findManagerEmailByReference(employee.reportingManager, '');
    }

    let extraHtml = '';
    if (status === 'Approved') {
      extraHtml = '<p><strong>Approved By:</strong> ' + approver + '</p>';
    } else if (status === 'Rejected') {
      extraHtml = '<p><strong>Rejected By:</strong> ' + approver + '</p>' +
        '<p><strong>Reason:</strong> ' + (rejectionReason || 'Not provided') + '</p>';
    } else if (status === 'Cancelled') {
      extraHtml = '<p><strong>Updated By:</strong> ' + approver + '</p>' +
        '<p><strong>Reason:</strong> ' + (rejectionReason || 'Cancelled') + '</p>';
    }

    const htmlBody = '<html>' +
      '<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">' +
      '<div style="max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f9f9f9; border-radius: 8px;">' +
      '<h2 style="color: ' + statusColor + '; margin-top: 0;">Leave Request ' + status + '</h2>' +
      '<p>Dear ' + empName + ',</p>' +
      '<p>Your leave request has been updated.</p>' +
      '<div style="background-color: #fff; padding: 15px; border-left: 4px solid ' + statusColor + '; margin: 15px 0;">' +
      '<p><strong>Leave ID:</strong> ' + (leaveRecord.leaveId || leaveRecord.id || '') + '</p>' +
      '<p><strong>Leave Type:</strong> ' + leaveType + '</p>' +
      '<p><strong>From Date:</strong> ' + fromDate + '</p>' +
      '<p><strong>To Date:</strong> ' + toDate + '</p>' +
      '<p><strong>Reason:</strong> ' + reason + '</p>' +
      '<p><strong>Status:</strong> <span style="color: ' + statusColor + '; font-weight: bold;">' + status + '</span></p>' +
      extraHtml +
      '</div>' +
      '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
      '<p style="color: #999; font-size: 0.85em;">AttendPro System | Do not reply to this email</p>' +
      '</div>' +
      '</body>' +
      '</html>';
    const plainBody = 'Leave Request ' + status + '\n\n' +
      'Dear ' + empName + ',\n\n' +
      'Your leave request has been updated.\n\n' +
      'Leave ID: ' + (leaveRecord.leaveId || leaveRecord.id || '') + '\n' +
      'Leave Type: ' + leaveType + '\n' +
      'From Date: ' + fromDate + '\n' +
      'To Date: ' + toDate + '\n' +
      'Reason: ' + reason + '\n' +
      'Status: ' + status + '\n' +
      (status === 'Approved' ? 'Approved By: ' + approver + '\n' : '') +
      (status === 'Rejected' ? 'Rejected By: ' + approver + '\nRejection Reason: ' + (rejectionReason || 'Not provided') + '\n' : '') +
      (status === 'Cancelled' ? 'Cancelled By: ' + approver + '\nCancellation Reason: ' + (rejectionReason || 'Cancelled') + '\n' : '');

    // Send email to employee
    try {
      MailApp.sendEmail({
        to: empEmail,
        subject: 'Leave Request ' + status + ' - ' + empName,
        body: plainBody,
        htmlBody: htmlBody
      });
      result.sent.push('employee:' + empEmail);
      Logger.log('Leave status update email sent to employee: ' + empEmail);
    } catch (empEmailErr) {
      result.issues.push('Failed to send employee status email: ' + empEmailErr.toString());
      Logger.log(result.issues[result.issues.length - 1]);
    }
    
    // Also send notification to manager and reporting manager about the decision
    if (status === 'Approved' || status === 'Rejected' || status === 'Cancelled') {
      if (managerEmail && managerEmail.indexOf('@') > 0) {
        const managerNotificationHtml = '<html>' +
          '<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">' +
          '<div style="max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f9f9f9; border-radius: 8px;">' +
          '<h2 style="color: ' + statusColor + '; margin-top: 0;">Leave Request ' + status + ' - ' + empName + '</h2>' +
          '<p>Dear Manager,</p>' +
          '<p>The following leave request has been ' + status.toLowerCase() + ':</p>' +
          '<div style="background-color: #fff; padding: 15px; border-left: 4px solid ' + statusColor + '; margin: 15px 0;">' +
          '<p><strong>Employee Name:</strong> ' + empName + '</p>' +
          '<p><strong>Employee Email:</strong> ' + empEmail + '</p>' +
          '<p><strong>Leave ID:</strong> ' + (leaveRecord.leaveId || leaveRecord.id || '') + '</p>' +
          '<p><strong>Leave Type:</strong> ' + leaveType + '</p>' +
          '<p><strong>From Date:</strong> ' + fromDate + '</p>' +
          '<p><strong>To Date:</strong> ' + toDate + '</p>' +
          '<p><strong>Status:</strong> <span style="color: ' + statusColor + '; font-weight: bold;">' + status + '</span></p>' +
          extraHtml +
          '</div>' +
          '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
          '<p style="color: #999; font-size: 0.85em;">AttendPro System | Do not reply to this email</p>' +
          '</div>' +
          '</body>' +
          '</html>';
        
        try {
          MailApp.sendEmail({
            to: managerEmail,
            subject: 'Leave Request ' + status + ' - ' + empName,
            body: plainBody,
            htmlBody: managerNotificationHtml
          });
          result.sent.push('manager:' + managerEmail);
          Logger.log('Leave status update email sent to manager: ' + managerEmail);
        } catch (mgrEmailErr) {
          result.issues.push('Failed to send manager status email: ' + mgrEmailErr.toString());
          Logger.log(result.issues[result.issues.length - 1]);
        }
      } else {
        Logger.log('Manager email not found for leave status update notification. Employee: ' + empEmail);
      }

      if (reportingManagerEmail && reportingManagerEmail.indexOf('@') > 0 && reportingManagerEmail !== managerEmail) {
        const reportingNotificationHtml = '<html>' +
          '<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">' +
          '<div style="max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f9f9f9; border-radius: 8px;">' +
          '<h2 style="color: ' + statusColor + '; margin-top: 0;">Leave Request ' + status + ' - ' + empName + '</h2>' +
          '<p>Dear Reporting Manager,</p>' +
          '<p>The following leave request has been ' + status.toLowerCase() + ':</p>' +
          '<div style="background-color: #fff; padding: 15px; border-left: 4px solid ' + statusColor + '; margin: 15px 0;">' +
          '<p><strong>Employee Name:</strong> ' + empName + '</p>' +
          '<p><strong>Employee Email:</strong> ' + empEmail + '</p>' +
          '<p><strong>Leave ID:</strong> ' + (leaveRecord.leaveId || leaveRecord.id || '') + '</p>' +
          '<p><strong>Leave Type:</strong> ' + leaveType + '</p>' +
          '<p><strong>From Date:</strong> ' + fromDate + '</p>' +
          '<p><strong>To Date:</strong> ' + toDate + '</p>' +
          '<p><strong>Status:</strong> <span style="color: ' + statusColor + '; font-weight: bold;">' + status + '</span></p>' +
          extraHtml +
          '</div>' +
          '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
          '<p style="color: #999; font-size: 0.85em;">AttendPro System | Do not reply to this email</p>' +
          '</div>' +
          '</body>' +
          '</html>';

        try {
          MailApp.sendEmail({
            to: reportingManagerEmail,
            subject: 'Leave Request ' + status + ' - ' + empName,
            body: plainBody,
            htmlBody: reportingNotificationHtml
          });
          result.sent.push('reporting:' + reportingManagerEmail);
          Logger.log('Leave status update email sent to reporting manager: ' + reportingManagerEmail);
        } catch (reportEmailErr) {
          result.issues.push('Failed to send reporting manager status email: ' + reportEmailErr.toString());
          Logger.log(result.issues[result.issues.length - 1]);
        }
      } else {
        Logger.log('Reporting manager email not found for leave status update notification. Employee: ' + empEmail);
      }
    }
    
    result.ok = result.sent.length > 0;
    Logger.log('Leave status update emails completed. Sent: ' + JSON.stringify(result.sent) + ', Issues: ' + JSON.stringify(result.issues));
    return result;
  } catch (err) {
    Logger.log('Error sending leave status email: ' + err.toString());
    result.issues.push(err.toString());
    return result;
  }
}

function testLeaveMail() {
  const employee = {
    email: 'employee@dashverse.ai',
    name: 'Test Employee',
    manager: 'Test Manager',
    managerEmail: 'manager@dashverse.ai',
    reportingManager: 'Test Reporting Manager',
    reportingManagerEmail: 'reporting@dashverse.ai'
  };

  const data = {
    name: 'Test Employee',
    leaveType: 'Sick Leave',
    fromDate: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    toDate: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    reason: 'Testing leave notification',
    manager: employee.manager,
    managerEmail: employee.managerEmail,
    reportingManager: employee.reportingManager,
    reportingManagerEmail: employee.reportingManagerEmail
  };

  const leaveId = 'LVTEST001';
  const result = sendLeaveApplicationEmail(employee, data, leaveId);
  Logger.log(JSON.stringify(result));
  return result;
}

function testEmailSystem(data) {
  const result = { success: false, message: '', details: {} };
  try {
    // Test MailApp availability
    const testEmail = data.testEmail || 'test@dashverse.ai';
    const subject = 'AttendPro Email System Test - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const htmlBody = '<html><body style="font-family: Arial, sans-serif;">' +
      '<h2 style="color: #667eea;">✅ AttendPro Email System Test</h2>' +
      '<p>This is a test email to verify the email system is working properly.</p>' +
      '<p><strong>Test Time:</strong> ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') + '</p>' +
      '<p><strong>Status:</strong> <span style="color: #16a34a; font-weight: bold;">Email System Operational</span></p>' +
      '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
      '<p style="color: #999; font-size: 0.85em;">AttendPro System | Do not reply to this email</p>' +
      '</body></html>';
    
    MailApp.sendEmail({
      to: testEmail,
      subject: subject,
      htmlBody: htmlBody
    });
    
    result.success = true;
    result.message = 'Test email sent successfully!';
    result.details = {
      sentTo: testEmail,
      timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      mailService: 'MailApp operational'
    };
    Logger.log('Email test successful: ' + JSON.stringify(result));
  } catch (err) {
    result.success = false;
    result.message = 'Email test failed: ' + err.toString();
    result.details = {
      error: err.toString(),
      timestamp: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    };
    Logger.log('Email test failed: ' + err.toString());
  }
  return result;
}
