// ═══════════════════════════════════════════════════════════════
// TEAM CALENDAR & TASK MANAGER — Google Apps Script Backend v7
// ═══════════════════════════════════════════════════════════════
// SETUP: Run "setupSheets" once after pasting this code.
// If upgrading from v6, run "upgradeToV7" to add Comments sheet
// and new task columns without losing existing data.
// ═══════════════════════════════════════════════════════════════

// ──── CONFIGURATION ────
const SHEET_ID = '1GoVbf0sZJRTy3S-ieWgxCx8OBxGAxqupQ8CR8pkuo20';

// ──── SHEET NAMES ────
const SHEETS = {
  USERS: 'Users',
  EVENTS: 'Events',
  TASKS: 'Tasks',
  COMMENTS: 'Comments'
};

const ROLES = { OWNER: 'owner', EA: 'ea', MEMBER: 'member' };

// ═══════════════════════════════════════
// SETUP — Run this ONCE for new installs
// ═══════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Users sheet
  let uSheet = ss.getSheetByName(SHEETS.USERS);
  if (!uSheet) {
    uSheet = ss.insertSheet(SHEETS.USERS);
    uSheet.appendRow(['username', 'password', 'role', 'createdAt']);
    uSheet.getRange('1:1').setFontWeight('bold').setBackground('#059669').setFontColor('#FFFFFF');
    uSheet.appendRow(['Jitendra', 'admin123', 'owner', new Date().toISOString()]);
    uSheet.setFrozenRows(1);
  }

  // Events sheet
  let eSheet = ss.getSheetByName(SHEETS.EVENTS);
  if (!eSheet) {
    eSheet = ss.insertSheet(SHEETS.EVENTS);
    eSheet.appendRow(['id', 'title', 'person', 'date', 'start', 'end', 'type', 'createdBy', 'updatedAt']);
    eSheet.getRange('1:1').setFontWeight('bold').setBackground('#7C3AED').setFontColor('#FFFFFF');
    eSheet.setFrozenRows(1);
  }

  // Tasks sheet (v7: added start, end columns for scheduling)
  let tSheet = ss.getSheetByName(SHEETS.TASKS);
  if (!tSheet) {
    tSheet = ss.insertSheet(SHEETS.TASKS);
    tSheet.appendRow(['id', 'title', 'description', 'assignedTo', 'assignedBy', 'status', 'priority', 'dueDate', 'start', 'end', 'createdBy', 'createdAt', 'updatedAt']);
    tSheet.getRange('1:1').setFontWeight('bold').setBackground('#F43F5E').setFontColor('#FFFFFF');
    tSheet.setFrozenRows(1);
  }

  // Comments sheet (v7: new)
  let cSheet = ss.getSheetByName(SHEETS.COMMENTS);
  if (!cSheet) {
    cSheet = ss.insertSheet(SHEETS.COMMENTS);
    cSheet.appendRow(['id', 'taskId', 'text', 'author', 'replyTo', 'replyAuthor', 'createdAt']);
    cSheet.getRange('1:1').setFontWeight('bold').setBackground('#F59E0B').setFontColor('#FFFFFF');
    cSheet.setFrozenRows(1);
  }

  return 'Setup complete! Default owner: Jitendra / admin123';
}

// ═══════════════════════════════════════
// UPGRADE — Run this if upgrading from v6
// ═══════════════════════════════════════
function upgradeToV7() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // Add Comments sheet if missing
  let cSheet = ss.getSheetByName(SHEETS.COMMENTS);
  if (!cSheet) {
    cSheet = ss.insertSheet(SHEETS.COMMENTS);
    cSheet.appendRow(['id', 'taskId', 'text', 'author', 'replyTo', 'replyAuthor', 'createdAt']);
    cSheet.getRange('1:1').setFontWeight('bold').setBackground('#F59E0B').setFontColor('#FFFFFF');
    cSheet.setFrozenRows(1);
  }

  // Add start/end columns to Tasks if missing
  const tSheet = ss.getSheetByName(SHEETS.TASKS);
  if (tSheet) {
    const headers = tSheet.getRange(1, 1, 1, tSheet.getLastColumn()).getValues()[0];
    if (headers.indexOf('start') === -1) {
      // Insert start and end columns after dueDate
      const dueDateIdx = headers.indexOf('dueDate');
      if (dueDateIdx >= 0) {
        tSheet.insertColumnAfter(dueDateIdx + 1);
        tSheet.getRange(1, dueDateIdx + 2).setValue('start');
        tSheet.insertColumnAfter(dueDateIdx + 2);
        tSheet.getRange(1, dueDateIdx + 3).setValue('end');
      }
    }
  }

  return 'Upgrade to v7 complete! Comments sheet and task scheduling columns added.';
}

// ═══════════════════════════════════════
// HTTP HANDLERS
// ═══════════════════════════════════════
function doGet(e) {
  let result;
  try {
    const action = e.parameter.action;
    const user = e.parameter.user || '';

    switch (action) {
      case 'ping': result = { status: 'ok', version: '7.0' }; break;
      case 'login': result = handleLogin(user, e.parameter.password || ''); break;
      case 'getUsers': result = getUsers(user); break;
      case 'getEvents': result = getEvents(user); break;
      case 'getTasks': result = getTasks(user); break;
      case 'getComments': result = getComments(e.parameter.taskId || '', user); break;
      default: result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return jsonResponse(result);
}

function doPost(e) {
  let result;
  try {
    const data = JSON.parse(e.postData.contents);
    const user = data.user || '';

    if (!verifyUser(user)) {
      return jsonResponse({ error: 'Unauthorized' });
    }

    switch (data.action) {
      case 'saveEvent': result = saveEvent(data.event, user); break;
      case 'deleteEvent': result = deleteEvent(data.id, user); break;
      case 'saveTask': result = saveTask(data.task, user); break;
      case 'deleteTask': result = deleteTask(data.id, user); break;
      case 'updateTaskStatus': result = updateTaskStatus(data.id, data.status, user); break;
      case 'addUser': result = addUser(data.newUser, user); break;
      case 'editUser': result = editUser(data.targetUser, data.updates, user); break;
      case 'removeUser': result = removeUser(data.targetUser, user); break;
      case 'addComment': result = addComment(data, user); break;
      default: result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { error: err.message };
  }
  return jsonResponse(result);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════
// AUTH & USERS
// ═══════════════════════════════════════
function handleLogin(username, password) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === username.toLowerCase()) {
      if (data[i][1].toString() === password) {
        return { success: true, username: data[i][0], role: data[i][2] };
      } else {
        return { error: 'Wrong password' };
      }
    }
  }
  return { error: 'User not found. Ask your admin to create your account.' };
}

function verifyUser(username) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === username.toLowerCase()) return true;
  }
  return false;
}

function getUsers(requestedBy) {
  const role = getUserRole(requestedBy);
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  const users = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      if (role === ROLES.OWNER) {
        users.push({ username: data[i][0], password: data[i][1], role: data[i][2] });
      } else {
        users.push({ username: data[i][0], role: data[i][2] });
      }
    }
  }
  return { users };
}

function addUser(newUser, requestedBy) {
  const role = getUserRole(requestedBy);
  if (role !== ROLES.OWNER) return { error: 'Only owner can add users' };
  if (!newUser.username || !newUser.password) return { error: 'Username and password required' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === newUser.username.toLowerCase()) {
      return { error: 'User already exists' };
    }
  }

  sheet.appendRow([newUser.username, newUser.password, newUser.role || ROLES.MEMBER, new Date().toISOString()]);
  return { success: true };
}

function editUser(targetUser, updates, requestedBy) {
  const role = getUserRole(requestedBy);
  if (role !== ROLES.OWNER) return { error: 'Only owner can edit users' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === targetUser.toLowerCase()) {
      if (updates.password) sheet.getRange(i + 1, 2).setValue(updates.password);
      if (updates.role) sheet.getRange(i + 1, 3).setValue(updates.role);
      if (updates.username) sheet.getRange(i + 1, 1).setValue(updates.username);
      return { success: true };
    }
  }
  return { error: 'User not found' };
}

function removeUser(targetUser, requestedBy) {
  const role = getUserRole(requestedBy);
  if (role !== ROLES.OWNER) return { error: 'Only owner can remove users' };
  if (targetUser.toLowerCase() === requestedBy.toLowerCase()) return { error: 'Cannot remove yourself' };

  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === targetUser.toLowerCase()) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: 'User not found' };
}

function getUserRole(username) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === username.toLowerCase()) {
      return data[i][2];
    }
  }
  return null;
}

// ═══════════════════════════════════════
// EVENTS (Calendar)
// ═══════════════════════════════════════
function getEvents(user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.EVENTS);
  if (!sheet) return { events: {} };

  const role = getUserRole(user);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const events = {};
  const tz = ss.getSpreadsheetTimeZone();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const event = {};
    headers.forEach((h, j) => {
      let val = row[j];
      if (val instanceof Date) {
        if (h === 'start' || h === 'end') {
          val = Utilities.formatDate(val, tz, 'HH:mm');
        } else if (h === 'date' || h === 'dueDate') {
          val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
        }
      }
      event[h] = val;
    });

    const creator = (event.createdBy || '').toLowerCase();
    const userLower = user.toLowerCase();

    if (role === ROLES.OWNER) {
      events[row[0]] = event;
    } else if (role === ROLES.EA) {
      const ownerName = getOwnerName().toLowerCase();
      if (creator === userLower || creator === ownerName) {
        events[row[0]] = event;
      }
    } else {
      if (creator === userLower) {
        events[row[0]] = event;
      }
    }
  }

  return { events };
}

function getOwnerName() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === ROLES.OWNER) return data[i][0];
  }
  return '';
}

function saveEvent(event, user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.EVENTS);

  if (!event.id) {
    event.id = 'ev_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 5);
  }
  event.createdBy = event.createdBy || user;
  event.updatedAt = new Date().toISOString();

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === event.id) { rowIndex = i + 1; break; }
  }

  const rowData = headers.map(h => event[h] || '');
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  return { success: true, id: event.id };
}

function deleteEvent(id, user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.EVENTS);
  const data = sheet.getDataRange().getValues();
  const role = getUserRole(user);

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const creator = (data[i][7] || '').toLowerCase();
      if (role === ROLES.OWNER || role === ROLES.EA || creator === user.toLowerCase()) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
      return { error: 'Cannot delete others events' };
    }
  }
  return { error: 'Event not found' };
}

// ═══════════════════════════════════════
// TASKS
// ═══════════════════════════════════════
function getTasks(user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.TASKS);
  if (!sheet) return { tasks: {} };

  const role = getUserRole(user);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const tasks = {};
  const tz = ss.getSpreadsheetTimeZone();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const task = {};
    headers.forEach((h, j) => {
      let val = row[j];
      if (val instanceof Date) {
        if (h === 'dueDate' || h === 'date') {
          val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
        } else if (h === 'start' || h === 'end') {
          val = Utilities.formatDate(val, tz, 'HH:mm');
        }
      }
      task[h] = val;
    });

    const assignedTo = (task.assignedTo || '').toLowerCase();
    const assignedBy = (task.assignedBy || '').toLowerCase();
    const createdBy = (task.createdBy || '').toLowerCase();
    const userLower = user.toLowerCase();
    const ownerName = getOwnerName().toLowerCase();

    if (role === ROLES.OWNER) {
      tasks[row[0]] = task;
    } else if (role === ROLES.EA) {
      if (assignedTo === ownerName || assignedBy === ownerName || createdBy === ownerName || assignedTo === userLower || assignedBy === userLower || createdBy === userLower) {
        tasks[row[0]] = task;
      }
    } else {
      if (assignedTo === userLower || assignedBy === userLower || createdBy === userLower) {
        tasks[row[0]] = task;
      }
    }
  }

  // Attach comment counts to each task
  const commSheet = ss.getSheetByName(SHEETS.COMMENTS);
  if (commSheet && commSheet.getLastRow() > 1) {
    const commData = commSheet.getDataRange().getValues();
    const commHeaders = commData[0];
    const taskIdCol = commHeaders.indexOf('taskId');
    const createdAtCol = commHeaders.indexOf('createdAt');
    const authorCol = commHeaders.indexOf('author');

    for (let i = 1; i < commData.length; i++) {
      const tid = commData[i][taskIdCol];
      if (tasks[tid]) {
        tasks[tid].commentCount = (tasks[tid].commentCount || 0) + 1;
        const cat = commData[i][createdAtCol] instanceof Date
          ? Utilities.formatDate(commData[i][createdAtCol], tz, "yyyy-MM-dd'T'HH:mm:ss")
          : (commData[i][createdAtCol] || '');
        const catStr = String(cat);
        if (!tasks[tid].lastCommentAt || catStr > tasks[tid].lastCommentAt) {
          tasks[tid].lastCommentAt = catStr;
          tasks[tid].lastCommentBy = commData[i][authorCol] || '';
        }
      }
    }
  }

  return { tasks };
}

function saveTask(task, user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.TASKS);
  const role = getUserRole(user);

  if (!task.id) {
    task.id = 'task_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 5);
    task.createdAt = new Date().toISOString();
  }
  task.createdBy = task.createdBy || user;
  if (role === ROLES.EA && !task.assignedBy) {
    task.assignedBy = getOwnerName();
  } else {
    task.assignedBy = task.assignedBy || user;
  }
  task.updatedAt = new Date().toISOString();

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let rowIndex = -1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === task.id) { rowIndex = i + 1; break; }
  }

  const rowData = headers.map(h => task[h] !== undefined ? task[h] : '');
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }

  return { success: true, id: task.id };
}

function deleteTask(id, user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.TASKS);
  const data = sheet.getDataRange().getValues();
  const role = getUserRole(user);

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const creator = (data[i][8] || '').toLowerCase();
      if (role === ROLES.OWNER || role === ROLES.EA || creator === user.toLowerCase()) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
      return { error: 'Cannot delete this task' };
    }
  }
  return { error: 'Task not found' };
}

function updateTaskStatus(id, status, user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.TASKS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const role = getUserRole(user);
  const statusCol = headers.indexOf('status') + 1;
  const updatedCol = headers.indexOf('updatedAt') + 1;

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const assignedTo = (data[i][headers.indexOf('assignedTo')] || '').toLowerCase();
      const createdBy = (data[i][headers.indexOf('createdBy')] || '').toLowerCase();
      const ownerName = getOwnerName().toLowerCase();
      const userLower = user.toLowerCase();

      if (role === ROLES.OWNER ||
          (role === ROLES.EA && (assignedTo === ownerName || assignedTo === userLower || createdBy === userLower)) ||
          (assignedTo === userLower || createdBy === userLower)) {
        sheet.getRange(i + 1, statusCol).setValue(status);
        sheet.getRange(i + 1, updatedCol).setValue(new Date().toISOString());
        return { success: true };
      }
      return { error: 'No permission to update this task' };
    }
  }
  return { error: 'Task not found' };
}

// ═══════════════════════════════════════
// COMMENTS (v7 — new)
// ═══════════════════════════════════════
function getComments(taskId, user) {
  if (!taskId) return { comments: [] };
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEETS.COMMENTS);
  if (!sheet || sheet.getLastRow() <= 1) return { comments: [] };

  const tz = ss.getSpreadsheetTimeZone();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const comments = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const comment = {};
    headers.forEach((h, j) => {
      let val = row[j];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, tz, "yyyy-MM-dd'T'HH:mm:ss");
      }
      comment[h] = val;
    });
    if (comment.taskId === taskId) {
      comments.push(comment);
    }
  }

  comments.sort((a, b) => (a.createdAt || '').localeCompare(b.createdAt || ''));
  return { comments };
}

function addComment(data, user) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(SHEETS.COMMENTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEETS.COMMENTS);
    sheet.appendRow(['id', 'taskId', 'text', 'author', 'replyTo', 'replyAuthor', 'createdAt']);
    sheet.getRange('1:1').setFontWeight('bold').setBackground('#F59E0B').setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);
  }

  const id = 'cmt_' + new Date().getTime() + '_' + Math.random().toString(36).substr(2, 5);
  const now = new Date().toISOString();
  sheet.appendRow([id, data.taskId || '', data.text || '', user, data.replyTo || '', data.replyAuthor || '', now]);
  return { success: true, id, createdAt: now };
}
