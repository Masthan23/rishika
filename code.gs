
const SHEET_ID = '1Vtb5IuClsUs9TPVrM_Gc3CkopvtIXbrDZYM7OzCWxko';

const EMPLOYEE_HEADERS = [
  'EmpID', 'Email', 'Name', 'Role', 'Department',
  'EmploymentType', 'WorkMode',
  'ReportingManager', 'ReportingManagerEmail',
  'Manager', 'ManagerEmail',
  'Phone', 'JoinDate',
  'ContractStartDate', 'ContractEndDate', 'ContractTotalDays',
  'NoticePeriod', 'RenewalNotes',
  'CurrentProject', 'AddedOn'
];

const MANAGER_HEADERS = [
  'ManagerID', 'ManagerType', 'Name', 'Email', 'Role', 'Department', 'Team', 'AddedOn'
];

const HR_PROFILE_HEADERS = [
  'HRID', 'Username', 'Password', 'Name', 'Email', 'Role', 'Status', 'AddedOn'
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
    case 'markAttendance':    return markAttendance(data);
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
    case 'getManagers':       return getManagers();
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
    'Attendance': [
      'RecordID', 'EmpID', 'Email', 'Name', 'Department',
      'CurrentProject', 'Date', 'CheckInTime', 'Location', 'Status'
    ],
    'Leaves': [
      'LeaveID', 'EmpID', 'Email', 'Name', 'Department',
      'LeaveType', 'FromDate', 'ToDate', 'Days', 'Reason',
      'Status', 'AppliedOn', 'UpdatedOn', 'ApprovedBy', 'ApprovedDate',
      'RejectedBy', 'RejectedDate', 'RejectionReason'
    ],
    'Projects': ['ProjectName'],
    'Managers': MANAGER_HEADERS,
    'HRProfiles': HR_PROFILE_HEADERS,
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
  if (matchedDefault) return { success: true, message: 'Login successful' };

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
      return { success: true, message: 'Login successful' };
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

const ALLOWED_EMAIL_DOMAINS = ['@dashversemail.com', '@dashverse.ai', '@dashtoon.com'];
function isAllowedWorkEmail(email) {
  const value = normalizeEmail(email);
  if (!value || value.indexOf('@') === -1) return false;
  for (let i = 0; i < ALLOWED_EMAIL_DOMAINS.length; i++) {
    if (value.endsWith(ALLOWED_EMAIL_DOMAINS[i])) return true;
  }
  return false;
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

function buildEmployeeObject(row, hdr) {
  const employmentType = getValueByHeader(row, hdr, 'employmenttype', -1) ||
    getValueByHeader(row, hdr, 'type', -1) ||
    'Permanent';

  return {
    id: row[hdr['empid']] || '',
    email: row[hdr['email']] || '',
    name: row[hdr['name']] || '',
    role: row[hdr['role']] || '',
    department: row[hdr['department']] || '',
    employmentType: employmentType,
    employeeType: employmentType,
    employment_type: employmentType,
    workMode: getValueByHeader(row, hdr, 'workmode', -1),
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
  return {
    id: getValueByHeader(row, hdr, 'managerid', 0),
    managerType: getValueByHeader(row, hdr, 'managertype', 1),
    name: getValueByHeader(row, hdr, 'name', 2),
    email: getValueByHeader(row, hdr, 'email', 3),
    role: getValueByHeader(row, hdr, 'role', 4),
    department: getValueByHeader(row, hdr, 'department', 5),
    team: getValueByHeader(row, hdr, 'team', 6),
    addedOn: getValueByHeader(row, hdr, 'addedon', 7)
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
    rejectionReason: getValueByHeader(row, hdr, 'rejectionreason', 17)
  };
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
    status: getValueByHeader(row, hdr, 'status', 9)
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

function addManager(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Managers');
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName('Managers');
  }

  const hdr = getManagerHeaders(sheet);
  const email = normalizeEmail(data.email);
  if (!email) return { success: false, message: 'Email is required' };
  if (!isAllowedWorkEmail(email)) {
    return { success: false, message: 'Email domain not allowed. Use @dashversemail.com, @dashverse.ai, or @dashtoon.com' };
  }
  if (!data.name) return { success: false, message: 'Name is required' };
  if (!data.managerType) return { success: false, message: 'Manager type is required' };

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
  if (hdr['managertype'] !== undefined) newRow[hdr['managertype']] = data.managerType || '';
  if (hdr['name'] !== undefined) newRow[hdr['name']] = data.name || '';
  if (hdr['email'] !== undefined) newRow[hdr['email']] = email;
  if (hdr['role'] !== undefined) newRow[hdr['role']] = data.role || '';
  if (hdr['department'] !== undefined) newRow[hdr['department']] = data.department || '';
  if (hdr['team'] !== undefined) newRow[hdr['team']] = data.team || '';
  if (hdr['addedon'] !== undefined) newRow[hdr['addedon']] = now;

  sheet.appendRow(newRow);
  return { success: true, message: 'Manager added successfully!', id: managerId };
}

function updateManager(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Managers');
  if (!sheet) return { success: false, message: 'Managers sheet not found' };

  const hdr = getManagerHeaders(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'No managers found' };

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const email = normalizeEmail(data.email);
  if (!isAllowedWorkEmail(email)) {
    return { success: false, message: 'Email domain not allowed. Use @dashversemail.com, @dashverse.ai, or @dashtoon.com' };
  }

  for (let i = 0; i < rows.length; i++) {
    if (normalizeEmail(getValueByHeader(rows[i], hdr, 'email', 3)) === email) {
      const r = i + 2;
      if (hasField(data, 'managerType')) setCellIfProvided(sheet, r, hdr, 'managertype', data.managerType);
      if (hasField(data, 'name'))        setCellIfProvided(sheet, r, hdr, 'name', data.name);
      if (hasField(data, 'role'))        setCellIfProvided(sheet, r, hdr, 'role', data.role);
      if (hasField(data, 'department'))  setCellIfProvided(sheet, r, hdr, 'department', data.department);
      if (hasField(data, 'team'))        setCellIfProvided(sheet, r, hdr, 'team', data.team);
      return { success: true, message: 'Manager updated successfully!' };
    }
  }

  return { success: false, message: 'Manager not found' };
}

function deleteManager(email) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('Managers');
  if (!sheet) return { success: false, message: 'Managers sheet not found' };

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

function markAttendance(data) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName('Attendance');
  if (!sheet) {
    setupSheets();
    sheet = ss.getSheetByName('Attendance');
  }

  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const now = new Date();
  const currentTime = now.getHours() + now.getMinutes() / 60;

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

  if (currentTime >= 15 && !loggedBefore15 && !loggedToday) {
    return { success: false, message: 'Please login before 3:00pm (15:00).' };
  }

  if (loggedToday) {
    return {
      success: true,
      alreadyMarked: true,
      message: 'Already marked present today',
      time: latestCheckIn,
      date: today
    };
  }

  let currentProject = data.currentProject || '';
  if (!currentProject) {
    try {
      ensureEmployeeSchema(ss);
      const empSheet = ss.getSheetByName('Employees');
      if (empSheet) {
        const empLastRow = empSheet.getLastRow();
        const empRows = empSheet.getRange(2, 1, empLastRow - 1, empSheet.getLastColumn()).getValues();
        const eHdr = getEmpHeaders(empSheet);
        const iMail = eHdr['email'] !== undefined ? eHdr['email'] : 1;
        const iProj = eHdr['currentproject'] !== undefined ? eHdr['currentproject'] : -1;
        if (iProj !== -1) {
          for (let i = 0; i < empRows.length; i++) {
            if (normalizeEmail(empRows[i][iMail]) === normalizeEmail(data.email)) {
              currentProject = empRows[i][iProj] || '';
              break;
            }
          }
        }
      }
    } catch (e) {}
  }

  const recordId = 'ATT' + Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmmss');
  const timeStr = Utilities.formatDate(new Date(), tz, 'HH:mm:ss');

  const aHdr = getAttHeaders(sheet);
  const totalC = sheet.getLastColumn();
  const newRow = new Array(totalC).fill('');

  if (aHdr['recordid'] !== undefined) newRow[aHdr['recordid']] = recordId;
  if (aHdr['empid'] !== undefined) newRow[aHdr['empid']] = data.id || '';
  if (aHdr['email'] !== undefined) newRow[aHdr['email']] = data.email || '';
  if (aHdr['name'] !== undefined) newRow[aHdr['name']] = data.name || '';
  if (aHdr['department'] !== undefined) newRow[aHdr['department']] = data.department || '';
  if (aHdr['currentproject'] !== undefined) newRow[aHdr['currentproject']] = currentProject;
  if (aHdr['date'] !== undefined) newRow[aHdr['date']] = today;
  if (aHdr['checkintime'] !== undefined) newRow[aHdr['checkintime']] = timeStr;
  if (aHdr['location'] !== undefined) newRow[aHdr['location']] = data.location || 'Office';
  if (aHdr['status'] !== undefined) newRow[aHdr['status']] = 'Present';

  sheet.appendRow(newRow);
  return { success: true, message: 'Attendance marked!', time: timeStr, date: today };
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

function getAttendance(email) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
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
        currentProject: iProj >= 0 ? (rows[i][iProj] || '') : ''
      });
    }
  }

  return { success: true, attendance: result };
}

function getAllAttendance() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Attendance');
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
        status: rows[i][iSts] || 'Present'
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
    reportingManagerEmail: normalizeEmail(data.reportingManagerEmail || (employee ? employee.reportingManagerEmail : '') || '')
  };

  sheet.appendRow([
    leaveId, leavePayload.empId, leavePayload.email, leavePayload.name, leavePayload.department,
    leavePayload.leaveType, leavePayload.fromDate, leavePayload.toDate, days,
    leavePayload.reason, 'Pending', now, now,
    '', '', '', '', ''
  ]);

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

  return {
    success: true,
    message: 'Leave applied successfully!',
    id: leaveId,
    notification: notificationResult
  };
}

function getLeaves(email) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Leaves');
  if (!sheet) return { success: true, leaves: [] };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: true, leaves: [] };

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getLeaveHeaders(sheet);
  const wantedEmail = normalizeEmail(email);

  const result = [];
  for (let i = 0; i < rows.length; i++) {
    const rowEmail = normalizeEmail(getValueByHeader(rows[i], hdr, 'email', 2));
    if (rowEmail === wantedEmail) {
      result.push(buildLeaveObject(rows[i], hdr));
    }
  }

  return { success: true, leaves: result };
}

function getAllLeaves() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Leaves');
  if (!sheet || sheet.getLastRow() < 2) return { success: true, leaves: [] };

  const lastRow = sheet.getLastRow();
  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getLeaveHeaders(sheet);

  const result = [];
  for (let i = 0; i < rows.length; i++) {
    if (getValueByHeader(rows[i], hdr, 'leaveid', 0)) {
      result.push(buildLeaveObject(rows[i], hdr));
    }
  }

  return { success: true, leaves: result };
}

function updateLeaveStatus(data) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Leaves');
  if (!sheet) return { success: false, message: 'Sheet not found' };

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { success: false, message: 'No leaves found' };

  const rows = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const hdr = getLeaveHeaders(sheet);
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const requestedId = (data.id || '').toString().trim();
  if (!requestedId) return { success: false, message: 'Leave ID is required' };

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

      let notificationResult = { ok: false, sent: [], issues: [] };
      try {
        const employeeData = getEmployee(getValueByHeader(rows[i], hdr, 'email', 2));
        if (employeeData.success && employeeData.employee) {
          const leaveRecord = buildLeaveObject(rows[i], hdr);
          leaveRecord.status = newStatus;
          leaveRecord.approvedBy = data.approvedBy || leaveRecord.approvedBy || '';
          leaveRecord.rejectedBy = data.rejectedBy || leaveRecord.rejectedBy || '';
          leaveRecord.rejectionReason = data.rejectionReason || leaveRecord.rejectionReason || '';
          leaveRecord.updatedOn = now;
          notificationResult = sendLeaveStatusUpdateEmail(employeeData.employee, leaveRecord);
        } else {
          notificationResult.issues.push('Employee record not found for leave status email');
          Logger.log(notificationResult.issues[notificationResult.issues.length - 1]);
        }
      } catch (emailErr) {
        notificationResult.issues.push('Leave status email error: ' + emailErr.toString());
        Logger.log(notificationResult.issues[notificationResult.issues.length - 1]);
      }

      return {
        success: true,
        message: 'Leave ' + newStatus.toLowerCase() + ' successfully!',
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
  if (!sheet || sheet.getLastRow() < 2) return '';

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const hdr = getManagerHeaders(sheet);
  const wantedType = (managerType || '').toString().trim().toLowerCase();

  for (let i = 0; i < rows.length; i++) {
    const rowType = (getValueByHeader(rows[i], hdr, 'managertype', 1) || '').toString().trim().toLowerCase();
    const rowName = (getValueByHeader(rows[i], hdr, 'name', 2) || '').toString().trim().toLowerCase();
    const rowEmail = normalizeEmail(getValueByHeader(rows[i], hdr, 'email', 3));
    if (wantedType && rowType && rowType !== wantedType) continue;
    if ((ref && rowEmail === ref) || (refText && rowName === refText)) {
      return rowEmail || '';
    }
  }

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
    const managerReference = data.managerEmail || employee.managerEmail || data.manager || employee.manager || '';
    const managerEmail = normalizeEmail(data.managerEmail || employee.managerEmail || '') ||
      findManagerEmailByReference(managerReference, 'manager');
    const managerName = data.manager || employee.manager || 'Manager';
    const reportingManagerReference = data.reportingManagerEmail || employee.reportingManagerEmail || data.reportingManager || employee.reportingManager || '';
    const reportingManagerEmail = normalizeEmail(data.reportingManagerEmail || employee.reportingManagerEmail || '') ||
      findManagerEmailByReference(reportingManagerReference, 'reporting');
    const reportingManagerName = data.reportingManager || employee.reportingManager || 'Reporting Manager';

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
      '<p style="color: #666; font-size: 0.9em;">Your leave request has been submitted for approval. Your manager will review and respond shortly.</p>' +
      '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
      '<p style="color: #999; font-size: 0.85em;">AttendPro System | Do not reply to this email</p>' +
      '</div>' +
      '</body>' +
      '</html>';
    if (empEmail) {
      MailApp.sendEmail({
        to: empEmail,
        subject: 'Leave Application Confirmation - ' + empName,
        htmlBody: emailBody
      });
      result.sent.push('employee:' + empEmail);
    } else {
      result.issues.push('Employee email missing for leave confirmation');
    }
    if (managerEmail) {
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
        '<p style="color: #666; margin-top: 15px;"><strong>Action Required:</strong> Please visit the AttendPro Manager Portal to approve or reject this request.</p>' +
        '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
        '<p style="color: #999; font-size: 0.85em;">AttendPro System | Do not reply to this email</p>' +
        '</div>' +
        '</body>' +
        '</html>';

      MailApp.sendEmail({
        to: managerEmail,
        subject: 'Leave Request for Approval - ' + empName,
        htmlBody: managerEmailBody
      });
      result.sent.push('manager:' + managerEmail);
    } else {
      result.issues.push('Manager email missing for leave notification. Employee: ' + empEmail + ', manager ref: ' + (employee.manager || employee.managerEmail || ''));
      Logger.log(result.issues[result.issues.length - 1]);
    }
    if (reportingManagerEmail && reportingManagerEmail !== managerEmail) {
      const reportingManagerEmailBody = '<html>' +
        '<body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">' +
        '<div style="max-width: 600px; margin: 0 auto; padding: 20px; background-color: #f9f9f9; border-radius: 8px;">' +
        '<h2 style="color: #667eea; margin-top: 0;">Leave Request Notification</h2>' +
        '<p>Dear ' + reportingManagerName + ',</p>' +
        '<p style="color: #666;"><strong>' + empName + '</strong> has applied for ' + leaveType + '. This is for your information.</p>' +
        '<div style="background-color: #fff; padding: 15px; border-left: 4px solid #667eea; margin: 15px 0;">' +
        '<p><strong>Employee Name:</strong> ' + empName + '</p>' +
        '<p><strong>Employee Email:</strong> ' + empEmail + '</p>' +
        '<p><strong>Leave ID:</strong> ' + leaveId + '</p>' +
        '<p><strong>Leave Type:</strong> ' + leaveType + '</p>' +
        '<p><strong>From Date:</strong> ' + fromDate + '</p>' +
        '<p><strong>To Date:</strong> ' + toDate + '</p>' +
        '<p><strong>Reason:</strong> ' + reason + '</p>' +
        '</div>' +
        '<p style="color: #666; margin-top: 15px;">This request has been forwarded to the direct manager for approval.</p>' +
        '<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">' +
        '<p style="color: #999; font-size: 0.85em;">AttendPro System | Do not reply to this email</p>' +
        '</div>' +
        '</body>' +
        '</html>';

      MailApp.sendEmail({
        to: reportingManagerEmail,
        subject: 'Leave Request Notification - ' + empName,
        htmlBody: reportingManagerEmailBody
      });
      result.sent.push('reporting:' + reportingManagerEmail);
    } else if (!reportingManagerEmail) {
      result.issues.push('Reporting manager email missing for leave notification. Employee: ' + empEmail + ', reporting manager ref: ' + (employee.reportingManager || employee.reportingManagerEmail || ''));
      Logger.log(result.issues[result.issues.length - 1]);
    }

    Logger.log('Leave notification emails sent successfully for Leave ID: ' + leaveId);
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
      '<p><strong>Leave ID:</strong> ' + (leaveRecord.leaveId || '') + '</p>' +
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

    MailApp.sendEmail({
      to: empEmail,
      subject: 'Leave Request ' + status + ' - ' + empName,
      htmlBody: htmlBody
    });
    result.sent.push('employee:' + empEmail);
    result.ok = true;
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
