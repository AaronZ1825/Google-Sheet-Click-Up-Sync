/**
 * Complete two-way synchronization script：
 * - syncAll: merge two functions into one wrapper
 * - syncToClickUp(): Google Sheets → ClickUp (with comments and Logger.log)
 * - syncFromClickUp(): ClickUp → Google Sheets (with comments and Logger.log)
 * Support: Assignees, Tags, 10 Custom Fields (dynamically pull metadata)
 * logger.log is for you to debug in the Execution log
 */

// ---- Constant configuration ----
const SHEET_ID      = 'type in your google sheet ID here';           // Google Sheets ID
const LIST_ID       = 'type in your click up list id here';            // ClickUp List ID
const TOKEN         = 'type in your click up api token';          // ClickUp API Token
const BASE_URL      = 'https://api.clickup.com/api/v2';
const ASSIGNEES_COL = 9;    // I 
const TAGS_COL      = 10;   // J 
const SNAPSHOT_COL  = 8;    // H 
const LAST_SYNC_COL = 21;

// 10 Custom Field column names (corresponding to Google Sheet column titles)
const CUSTOM_FIELD_NAMES = [
  'Approved By',
  'Category',
  'Est. Total Cost (CAD)',
  'Google Sheet Link',
  'Item Link',
  'Ordered Date',
  'Purpose',
  'Quantity',
  'Received Date',
  'Requested By'
];

// Cache Custom Field metadata: { name: { id, type, type_config... } }
let fieldMetaMap = {};

/**
 * Pull and cache custom field metadata
 */
function loadCustomFieldMeta() {
  Logger.log('>> loadCustomFieldMeta');
  const response = UrlFetchApp.fetch(
    `${BASE_URL}/list/${LIST_ID}/field`,
    { headers: { 'Authorization': TOKEN }, muteHttpExceptions: true }
  );
  const code = response.getResponseCode();
  const body = response.getContentText();
  Logger.log(`GET /field return ${code}: ${body}`);
  if (code === 200) {
    const fields = JSON.parse(body).fields;
    fields.forEach(f => fieldMetaMap[f.name] = f);
    Logger.log(`Loaded ${fields.length} custom fields`);
  } else {
    Logger.log(`❌ Failed to load custom fields: ${code}`);
  }
}

// Global Cache
let usernameToId = {};
let idToUsername = {};

function loadMemberMap() {
  Logger.log('>> loadMemberMap');
  const res = UrlFetchApp.fetch(
    `${BASE_URL}/list/${LIST_ID}/member`,
    { headers: { 'Authorization': TOKEN }, muteHttpExceptions: true }
  );
  if (res.getResponseCode() !== 200) {
    Logger.log('❌ loadMemberMap fail: ' + res.getContentText());
    return;
  }
  const members = JSON.parse(res.getContentText()).members || [];
  members.forEach(m => {
    // If m.user does not exist, then m itself is considered to be a user object.
    const u = m.user || m;
    if (u && u.id && u.username) {
      usernameToId[u.username] = u.id;
      idToUsername[u.id]       = u.username;
    }
  });
  Logger.log(`Loaded ${members.length} members`);
}

// —— Global cache of ClickUp tasks for conflict detection —— 
let tasksCache = {};

/**
 * Wrapper of syncToClickUp and syncFromClickUp
 * Set trigger for this function only
 */
function syncAll() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('⚠️ Unable to obtain lock, skip this synchronization');
    return;
  }
  try {
    // —— First fetch and cache all Custom Field metadata —— 
    loadCustomFieldMeta();
    loadMemberMap();
    // —— ① Cache all ClickUp tasks (including closed ones) to memory —— 
    tasksCache = (function fetchAll() {
      const map = {};
      let page = 0;
      while (true) {
        const res = UrlFetchApp.fetch(
          `${BASE_URL}/list/${LIST_ID}/task?include_closed=true&page=${page}`,
          { headers:{ 'Authorization': TOKEN } }
        );
        const list = JSON.parse(res.getContentText()).tasks || [];
        if (!list.length) break;
        list.forEach(t => map[t.id] = t);
        page++;
      }
      return map;
    })();
    syncToClickUp();
    syncFromClickUp();
  } finally {
    lock.releaseLock();
  }
}


/**
 * Sheets → ClickUp Sync
 */
function syncToClickUp() {
  Logger.log('>>> syncToClickUp() Start <<<');
  // ——— ① Script start time, used for conflict detection ———
  const scriptStart = new Date().getTime();

  const sheet    = SpreadsheetApp.openById(SHEET_ID).getSheetByName('type in your google sheet spreadsheet name here');
  const startRow = 2;
  const lastRow  = sheet.getLastRow();
  const numRows  = lastRow - startRow + 1;
  if (numRows < 1) { 
    Logger.log('⚠️ No data row, exit'); 
    return; 
    }

  const customCount = CUSTOM_FIELD_NAMES.length;
  const totalCols   = TAGS_COL + customCount;
  const rows        = sheet.getRange(startRow, 1, numRows, totalCols).getValues();

  const props      = PropertiesService.getScriptProperties();
  const prevIds    = JSON.parse(props.getProperty('knownTaskIds') || '[]');
  const currentIds = rows.map(r => String(r[6] || '').trim()).filter(id => id);

  // 1) Delete removed tasks: only compare ID lists, no dependency on row count changes
  const toDelete = prevIds.filter(id => !currentIds.includes(id));
  // Take the existing deletion list
  const deletedIds = JSON.parse(props.getProperty('deletedTaskIds')||'[]');
  toDelete.forEach(taskId => {
    const url = `${BASE_URL}/task/${taskId}`;
    Logger.log(`DELETE [task remove] ${url}`);
    const resp = UrlFetchApp.fetch(url, {
      method: 'delete',
      headers: { 'Authorization': TOKEN },
      muteHttpExceptions: true
    });
    Logger.log(`→ ${resp.getResponseCode()}: ${resp.getContentText()}`);
    // Record to "Deleted List"
    deletedIds.push(taskId);
  });
  props.setProperty('deletedTaskIds', JSON.stringify(deletedIds));

  // 2) Line-by-line processing
  rows.forEach((row, idx) => {
    const r = startRow + idx;
    // ——— ② Conflict check: If the user has edited it during this sync, skip this line ———
    const cell = sheet.getRange(r, LAST_SYNC_COL).getValue();
    const lastSync = cell ? new Date(cell).getTime() : 0;
    if (lastSync > scriptStart) {
      Logger.log(`❗️ Row ${r} Manually edited during this sync, skipping writing back`);
      return;
    }
    // —— ③ Remote conflict detection (ignore description field): Use the cached 
    //      tasksCache to construct a snapshot and compare it with column H —— 
    const taskId2  = String(row[6] || '').trim();
    const lastHash2 = String(row[7] || '');
    if (taskId2 && tasksCache[taskId2]) {
      const remoteSnap = buildRemoteSnapshot(tasksCache[taskId2]);
      const remoteParts = remoteSnap.split('|');
      const localParts  = lastHash2.split('|');
      // Ignore index 1 (description) and only compare the rest
      const remoteCore = [remoteParts[0]].concat(remoteParts.slice(2)).join('|');
      const localCore  = [localParts[0]].concat(localParts.slice(2)).join('|');

    // —— DEBUG: Print local vs remote ——
    // Logger.log(🔍 localCore:  "${localCore}");
    // Logger.log(🔍 remoteCore: "${remoteCore}");

      if (remoteCore !== localCore) {
        Logger.log(`❗️ Task ${taskId2} The remote has been modified (except for the description), skip the local push, and update the snapshot`);
        // Write the latest remote snapshot back to column H to avoid "misjudgment" of conflicts next time
        sheet.getRange(r, SNAPSHOT_COL).setValue(remoteSnap);
        // Optional: Update the synchronization time to prevent interference from "local change detection"
        sheet.getRange(r, LAST_SYNC_COL).setValue(new Date());
        return;
      }
    }
    let [
      title, desc, startRaw, dueRaw,
      statusRaw, priority, taskId, lastHash,
      assigneesRaw, tagsRaw,
      ...cfRaws
    ] = row;
    const hasTitle = !!String(title).trim();

    // 2.1 State Normalization
    const status = statusRaw
      ? statusRaw.charAt(0).toUpperCase() + statusRaw.slice(1).toLowerCase()
      : '';
    sheet.getRange(r, 5).setValue(status);
    Logger.log(`Row ${r} status normalized to "${status}"`);

    // 2.2 Time format (keep the same as syncFromClickUp yyyy-MM-dd)
    const sd = startRaw
      ? Utilities.formatDate(new Date(startRaw), Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : '';
    const dd = dueRaw
      ? Utilities.formatDate(new Date(dueRaw),   Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : '';
    // —— Redefine startTs/dueTs for payload and coreSnap —— 
    const startTs = startRaw ? new Date(startRaw).getTime() : '';
    const dueTs   = dueRaw   ? new Date(dueRaw).getTime()   : '';

    // 2.3 Parsing Assignees & Tags
    const assigneesList = assigneesRaw
      ? String(assigneesRaw).split(/\s*,\s*/).map(name => {
      // If the ID is written, use it directly; otherwise use usernameToId mapping
      return usernameToId[name] || name;
    }).filter(x => x)
  : [];
    const tagsList = tagsRaw
      ? String(tagsRaw).split(/\s*,\s*/).filter(x => x)
      : [];

    // 2.4 Parsing Custom Fields
    const cfPayloads = [];
    const cfTextVals = [];
    CUSTOM_FIELD_NAMES.forEach((name, i) => {
      const meta   = fieldMetaMap[name];
      const raw    = cfRaws[i] || '';         // Raw Input
      let apiVal   = null;
      let textVal  = raw;                    // Text used to write snapshots
      switch (meta.type) {
        case 'drop_down':
          apiVal = meta.type_config.options.find(o => o.name === raw)?.id || null;
          break;
        case 'multi_select':
          apiVal = raw
            ? String(raw).split(/\s*,\s*/).map(n => meta.type_config.options.find(o => o.name === n)?.id).filter(x => x)
            : null;
          break;
        case 'short_text':
        case 'text':
          apiVal = raw !== '' && raw != null ? String(raw) : null;
          break;
        case 'number':
          apiVal = raw !== '' ? Number(raw) : null;
          break;
        case 'date':
          apiVal  = raw ? new Date(raw).getTime() : null;
          // 格式化成 yyyy-MM-dd
          textVal = raw
            ? Utilities.formatDate(new Date(raw), Session.getScriptTimeZone(), 'yyyy-MM-dd')
            : '';
          break;
        case 'url':
          apiVal = raw || null;
          break;
        case 'currency':
        case 'money':
          apiVal = raw !== '' ? Number(raw) : null;
          break;
      }
      cfTextVals.push(textVal);
      if (apiVal !== null) {
        const payload = { value: apiVal };
        if (meta.type === 'url') payload.type = 'url';
        cfPayloads.push({ id: meta.id, payload });
      }
    });

    // 2.5 Snapshot
    const snapshot = [
      title, desc, sd, dd,
      status, priority,
      assigneesList.join(','), tagsList.join(','),
      ...cfTextVals
    ].join('|');
    Logger.log(`Row ${r} snapshot="${snapshot}"`);

    // 3) Clear title to delete
    if (!hasTitle && taskId) {
      const url = `${BASE_URL}/task/${taskId}`;
      Logger.log(`DELETE [clear title] ${url}`);
      UrlFetchApp.fetch(url, {
        method: 'delete',
        headers: { 'Authorization': TOKEN },
        muteHttpExceptions: true
      });
      sheet.getRange(r, 1, 1, totalCols).clearContent();
      return;
    }

    // 4) Create a new task
    if (hasTitle && !taskId) {
      const payload = {
        name:          title,
        description:   desc != null ? String(desc) : undefined,
        start_date:    startTs || undefined,
        due_date:      dueTs   || undefined,
        status:        status  || undefined,
        priority:      priority ? String(priority) : undefined,
        assignees:     assigneesList.length ? assigneesList : undefined,
        tags:          tagsList.length ? tagsList : undefined,
        custom_fields: cfPayloads.map(cf => ({ id: cf.id, value: cf.payload.value }))
      };
      const url = `${BASE_URL}/list/${LIST_ID}/task`;
      Logger.log(`POST create ${url}`);
      Logger.log(`Payload=${JSON.stringify(payload)}`);
      const resp = UrlFetchApp.fetch(url, {
        method:           'post',
        contentType:      'application/json',
        headers:          { 'Authorization': TOKEN },
        payload:          JSON.stringify(payload),
        muteHttpExceptions: true
      });
      Logger.log(`→ ${resp.getResponseCode()}`);
      if (resp.getResponseCode() >= 200 && resp.getResponseCode() < 300) {
        const newId = JSON.parse(resp.getContentText()).id;
        sheet.getRange(r, 7).setValue(newId);
        sheet.getRange(r, SNAPSHOT_COL).setValue(snapshot);
        Logger.log(`➕ New create success taskId=${newId}, Write Back snapshot`);
      }
      return;
    }

    // 5) Update Tasks
    if (hasTitle && taskId && snapshot !== lastHash) {
      // 5.1 Core Field Updates
      const coreSnap = [title, desc, sd,   dd,   status, priority].join('|');
      const lastCore = lastHash.split('|').slice(0, 6).join('|');
      if (coreSnap !== lastCore) {
        const up = {
          name:       title,
          description:desc != null ? String(desc) : undefined,
          start_date: startTs || undefined,
          due_date:   dueTs   || undefined,
          status:     status  || undefined,
          priority:   priority ? String(priority) : undefined
        };
        const u = `${BASE_URL}/task/${taskId}`;
        Logger.log(`PUT Core fields ${u}`);
        const respUp = UrlFetchApp.fetch(u, {
          method:           'put',
          contentType:      'application/json',
          headers:          { 'Authorization': TOKEN },
          payload:          JSON.stringify(up),
          muteHttpExceptions: true
        });
        Logger.log(`→ ${respUp.getResponseCode()}`);
      }

      // 5.2 Assignees Sync
      const oldAsg = lastHash.split('|')[6] || '';
      const newAsg = assigneesList.join(',');
      if (newAsg !== oldAsg) {
        Logger.log(`Assignees Change: Old=[${oldAsg}], New=[${newAsg}]`);
        const respA = UrlFetchApp.fetch(`${BASE_URL}/task/${taskId}`, {
          headers: { 'Authorization': TOKEN }, muteHttpExceptions: true
        });
        const existA = JSON.parse(respA.getContentText()).assignees.map(a => String(a.id));
        existA.forEach(uid => {
          const du = `${BASE_URL}/task/${taskId}/assignee/${uid}`;
          UrlFetchApp.fetch(du, { method: 'delete', headers: { 'Authorization': TOKEN }, muteHttpExceptions: true });
        });
        assigneesList.forEach(uid => {
          const pu = `${BASE_URL}/task/${taskId}/assignee/${uid}`;
          UrlFetchApp.fetch(pu, { method: 'post', headers: { 'Authorization': TOKEN }, muteHttpExceptions: true });
        });
      }

      // 5.3 Tags Sync
      const oldTgs = lastHash.split('|')[7] || '';
      const newTgs = tagsList.join(',');
      if (newTgs !== oldTgs) {
        Logger.log(`Tags Change: Old=[${oldTgs}], New=[${newTgs}]`);
        const respT = UrlFetchApp.fetch(`${BASE_URL}/task/${taskId}`, {
          headers: { 'Authorization': TOKEN }, muteHttpExceptions: true
        });
        const existT = JSON.parse(respT.getContentText()).tags.map(t => t.name);
        existT.forEach(tag => {
          const du = `${BASE_URL}/task/${taskId}/tag/${encodeURIComponent(tag)}`;
          UrlFetchApp.fetch(du, { method: 'delete', headers: { 'Authorization': TOKEN }, muteHttpExceptions: true });
        });
        tagsList.forEach(tag => {
          const pu = `${BASE_URL}/task/${taskId}/tag/${encodeURIComponent(tag)}`;
          UrlFetchApp.fetch(pu, { method: 'post', headers: { 'Authorization': TOKEN }, muteHttpExceptions: true });
        });
      }

      // 5.4 Custom Fields Sync
      const oldCFs = lastHash.split('|').slice(8);
      CUSTOM_FIELD_NAMES.forEach((name, i) => {
        const meta   = fieldMetaMap[name];
        const oldVal = oldCFs[i] || '';
        const newVal = cfTextVals[i] || '';
        if (newVal !== oldVal) {
          const cf = cfPayloads.find(x => x.id === meta.id);
          if (cf) {
            const ucf = `${BASE_URL}/task/${taskId}/field/${meta.id}`;
            Logger.log(`POST CustomField ${name} ${ucf}`);
            UrlFetchApp.fetch(ucf, {
              method:           'post',
              contentType:      'application/json',
              headers:          { 'Authorization': TOKEN },
              payload:          JSON.stringify(cf.payload),
              muteHttpExceptions: true
            });
          }
        }
      });

      // 5.5 Write back snapshot
      sheet.getRange(r, SNAPSHOT_COL).setValue(snapshot);
      // ——— ③ Write back the synchronization timestamp ———
      sheet.getRange(r, LAST_SYNC_COL).setValue(new Date());
      Logger.log(`snapshot update H${r}`);
    }
  });
  
  // Rescan the entire table column 7 to ensure that all newly added or deleted TaskIds are included
  const finalIds = sheet
    .getRange(startRow, 7, lastRow - startRow + 1)
    .getValues()
    .flat()
    .map(x => String(x || '').trim())
    .filter(x => x);
  props.setProperty('knownTaskIds', JSON.stringify(finalIds));
  Logger.log('>>> syncToClickUp() End <<<');
}

/**
 * ClickUp → Google Sheets Sync
 */
function syncFromClickUp() {
  Logger.log('>>> syncFromClickUp() Start <<<');

  const props = PropertiesService.getScriptProperties();

  // —— 1) Pagination pull include_closed=true ——
  let tasks = [], page = 0;
  while (true) {
    // include_closed=true pull all first (including Complete/Closed)
    const url = `${BASE_URL}/list/${LIST_ID}/task?include_closed=true&page=${page}`;
    const res = UrlFetchApp.fetch(url, {
      headers:          { 'Authorization': TOKEN },
      muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return;
    const list = JSON.parse(res.getContentText()).tasks || [];
    if (!list.length) break;
    tasks.push(...list);
    page++;
  }
  Logger.log(`Pull ${tasks.length} in total`);

  // —— 快照比对 ——  
  const snapArr = tasks
    .map(t => `${t.id}:${t.date_updated||0}`)
    .sort();
  const currentSnap = snapArr.join(',');
  const prevSnap    = props.getProperty('clickupSnapshot');
  if (currentSnap === prevSnap) {
    Logger.log('⏭️ syncFromClickUp skipped — No remote changes');
    return;
  }
  props.setProperty('clickupSnapshot', currentSnap);
  Logger.log('✅ Found remote changes and started syncing to Sheet');

  // —— 2) 构建本地映射 ——  
  const sheet    = SpreadsheetApp.openById(SHEET_ID).getSheetByName('type in your google sheet spreadsheet name here');
  const startRow = 2;
  const lastRow  = sheet.getLastRow();
  const numRows2 = lastRow - startRow + 1;
  const cols2    = TAGS_COL + CUSTOM_FIELD_NAMES.length;
  const data     = numRows2 > 0
    ? sheet.getRange(startRow, 1, numRows2, cols2).getValues()
    : [];
  const map = {};
  data.forEach((r, i) => {
    if (r[6]) map[r[6]] = startRow + i;
  });
  const tz = Session.getScriptTimeZone();

  // —— 2.5) Filter out the Closed tasks that you deleted in the Sheet —— 
  const deletedIds = JSON.parse(props.getProperty('deletedTaskIds')||'[]');
  tasks = tasks.filter(t => !deletedIds.includes(t.id));
  
  // —— 3) Synchronize or append rows ——  
  tasks.forEach(t => {
    const tId     = t.id;
    const title   = t.name || '';
    const apiDesc = t.content?.trim() || t.description?.trim() || '';
    const sd      = t.start_date
      ? Utilities.formatDate(new Date(+t.start_date), tz, 'yyyy-MM-dd')
      : '';
    const dd      = t.due_date
      ? Utilities.formatDate(new Date(+t.due_date), tz, 'yyyy-MM-dd')
      : '';
    const status  = t.status?.status
      ? t.status.status.charAt(0).toUpperCase() + t.status.status.slice(1).toLowerCase()
      : '';
    const prio    = t.priority?.id || '';
    const asgStr = t.assignees
                  ?.map(a => idToUsername[a.id] || a.id)
                  .join(',') || '';

    // New: Assignee ID list for snapshot
    const asgIdStr = t.assignees
    ?.map(a => String(a.id))
    .join(',') || '';

    const tagsStr = t.tags?.map(x => x.name).join(',') || '';

    // Handling Custom Fields Text
    const cfTexts = CUSTOM_FIELD_NAMES.map(name => {
      const meta = fieldMetaMap[name];
      const cf   = t.custom_fields.find(f => f.id === meta.id);
      if (!cf) return '';

      if (cf.value_name) return cf.value_name;

      if (meta.type === 'drop_down' && meta.type_config?.options) {
        if (typeof cf.value === 'number' || /^\d+$/.test(String(cf.value))) {
          const opt = meta.type_config.options[+cf.value];
          if (opt?.name) return opt.name;
        }
        const optById = meta.type_config.options.find(o => o.id === cf.value);
        if (optById?.name) return optById.name;
      }
      if (meta.type === 'multi_select') {
        return cf.value_names?.join(',') || '';
      }
      switch (meta.type) {
        case 'short_text':
        case 'text':      return cf.value || '';
        case 'number':    return cf.value != null ? String(cf.value) : '';
        case 'date':      return cf.value ? Utilities.formatDate(new Date(+cf.value), tz, 'yyyy-MM-dd') : '';
        case 'url':       return cf.value || '';
        case 'currency':
        case 'money':     return cf.value != null ? Number(cf.value).toFixed(2) : '';
        default:          return String(cf.value || '');
      }
    });

    if (map[tId]) {
      // Update existing rows
      const r = map[tId];
      const oldDesc   = sheet.getRange(r, 2).getValue();
      const finalDesc = oldDesc || apiDesc;
      sheet.getRange(r, 1, 1, 6)
           .setValues([[ title, finalDesc, sd, dd, status, prio ]]);
      // If the user deletes the Assignee in the Sheet, but the assignee is still hanging on the server
      // Then call the API to delete all of them, and then clear the cell
      const cellAsg = sheet.getRange(r, ASSIGNEES_COL);
      const sheetAsg = cellAsg.getValue();
      // Take out the assignee part in the last snapshot to determine whether there was an assignee before
      const lastHash = sheet.getRange(r, SNAPSHOT_COL).getValue() || '';
      const prevAsg  = lastHash.split('|')[6] || '';
      cellAsg.setValue(asgStr);
      sheet.getRange(r, TAGS_COL).setValue(tagsStr);
      CUSTOM_FIELD_NAMES.forEach((_, i) =>
        sheet.getRange(r, TAGS_COL + 1 + i).setValue(cfTexts[i])
      );
      Logger.log(`↗️ update row ${r}`);
      // —— Write the latest row-level snapshot —— 
      const newSnapshot = [
        title, finalDesc, sd, dd, status, prio,
        asgIdStr, tagsStr,
        ...cfTexts
      ].join('|');
      sheet.getRange(r, SNAPSHOT_COL).setValue(newSnapshot);
      Logger.log(`snapshot update H${r}`);
    } else {
      // Add a new row
      const newRow = [
        title, apiDesc, sd, dd,
        status, prio, tId, '',
        asgStr, tagsStr,
        ...cfTexts
      ];
      sheet.appendRow(newRow);
      Logger.log(`➕ add new row (taskId=${tId})`);
      // —— new row also write snapshot —— 
      const newR = sheet.getLastRow();
      const newSnapshot = [
        title, apiDesc, sd, dd, status, prio,
        asgStr, tagsStr,
        ...cfTexts
      ].join('|');
      sheet.getRange(newR, SNAPSHOT_COL).setValue(newSnapshot);
      Logger.log(`snapshot Write in H${newR}`);
    }
  });

  // —— 4) Delete the removed task row ——  
  for (let rr = sheet.getLastRow(); rr >= startRow; rr--) {
    const id = sheet.getRange(rr, 7).getValue();
    if (id && !tasks.some(tt => tt.id === id)) {
      sheet.deleteRow(rr);
      Logger.log(`🗑️ delete row ${rr}`);
    }
  }

  Logger.log('>>> syncFromClickUp() End <<<');
}

/**
 * Convert a ClickUp task object into a snapshot string that matches the Google Sheet H column
 */
function buildRemoteSnapshot(t) {
  const tz = Session.getScriptTimeZone();

  // —— Construct the core part (including a "description placeholder") —— 
  const sd = t.start_date
    ? Utilities.formatDate(new Date(+t.start_date), tz, 'yyyy-MM-dd')
    : '';
  const dd = t.due_date
    ? Utilities.formatDate(new Date(+t.due_date),   tz, 'yyyy-MM-dd')
    : '';
  const status = t.status?.status
    ? t.status.status.charAt(0).toUpperCase() + t.status.status.slice(1).toLowerCase()
    : '';
  const prio = t.priority?.id || '';
  const assigneesStr = t.assignees?.map(a => a.id).join(',') || '';
  const tagsStr      = t.tags?.map(x => x.name).join(',') || '';
  // Core fields: 0=Title, 1=Description placeholder, 2=Start, 3=End, 4=Status, 5=Priority, 6=Assignees, 7=Tags
  const core = [
    t.name || '',
    '',
    sd,
    dd,
    status,
    prio,
    assigneesStr,
    tagsStr
  ].join('|');
  // —— Construct the Custom Fields section, 10 paragraphs in total —— 
  const cfs = CUSTOM_FIELD_NAMES.map(name => {
    const meta = fieldMetaMap[name];
    const cf   = t.custom_fields.find(f => f.id === meta.id);
    if (!cf) return '';
    let out = '';
    switch (meta.type) {
      case 'drop_down':
        // 1) Prioritize using API to directly output value_name
        if (cf.value_name) {
          out = cf.value_name;
          break;
        }
        // 2) If cf.value is a number (index), get name directly from the options array
        if (typeof cf.value === 'number' || /^\d+$/.test(String(cf.value))) {
          const idx = Number(cf.value);
          const opt = meta.type_config.options[idx];
          if (opt && opt.name) {
            out = opt.name;
            break;
          }
        }
        // 3) Finally, search by ID
        const optById = meta.type_config.options.find(o => o.id === cf.value);
        out = optById ? optById.name : '';
        break;
      case 'multi_select':
        // Multiple selection: value_names is preferred
        if (cf.value_names) {
          out = cf.value_names.join(',');
        } else if (Array.isArray(cf.value)) {
          out = cf.value
            .map(id => meta.type_config.options.find(o => o.id === id)?.name)
            .filter(n => n)
            .join(',');
        }
        break;
      case 'date':
        out = cf.value
          ? Utilities.formatDate(new Date(+cf.value), tz, 'yyyy-MM-dd')
          : '';
        break;
      case 'currency':
      case 'money':
        out = cf.value != null
          ? Number(cf.value).toFixed(2)
          : '';
        break;
      case 'number':
        out = cf.value != null
          ? String(cf.value)
          : '';
        break;
      // Short text / Normal text / URL
      case 'short_text':
      case 'text':
      case 'url':
        out = cf.value || '';
        break;
      default:
        out = cf.value != null
          ? String(cf.value)
          : '';
    }
    return out;
  }).join('|');
  return `${core}|${cfs}`;
}
