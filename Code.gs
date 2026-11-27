const APP_CONFIG = {
  APP_NAME: 'Dorm 14 Laundry Schedule',
  MENU_NAME: 'Dorm 14 Laundry',
  ADMIN_PAGE: 'AdminPage',
  TZ: 'Europe/Ljubljana',
  SPREADSHEET_ID_KEY: 'SPREADSHEET_ID',
  SCHEMA_VERSION_KEY: 'SCHEMA_VERSION',
  SCHEMA_VERSION: 'dorm14-fresh-v4',
  PUBLIC_BOARD_SIGNATURE_KEY: 'PUBLIC_BOARD_SIGNATURE',
  PUBLIC_BOARD_DIRTY_KEY: 'PUBLIC_BOARD_DIRTY',
  PUBLIC_BOARD_TRIGGER_INSTALLED_KEY: 'PUBLIC_BOARD_TRIGGER_INSTALLED',
  PUBLIC_BOARD_TRIGGER_HANDLER: 'publicBoardDispatcher',
  PUBLIC_BOARD_TRIGGER_EVERY_MINUTES: 1,
  DATA_CACHE_VERSION_KEY: 'DATA_CACHE_VERSION',
  DATA_CACHE_TTL_SECONDS: 40,
  DASHBOARD_CACHE_TTL_SECONDS: 20,
  AUTOMATION_KEYS_PREFIX: 'AUTOMATION',
  RETRYABLE_ERROR_TOKEN: '[RETRYABLE]',
  MUTATION_LOCK_TIMEOUT_MS: 1000,
  CYCLE_SWITCH_LOCK_TIMEOUT_MS: 400,
  CYCLE_SWITCH_SETTLE_TOTAL_MS: 3000,
  CYCLE_SWITCH_SETTLE_STEP_MS: 250,
  LAST_LOGIN_WRITE_THROTTLE_SECONDS: 86400,
  SHEETS: {
    SETTINGS: 'Settings',
    RESIDENTS: 'Residents',
    SLOTS: 'Slots',
    RESERVATIONS: 'Reservations',
    RESERVATIONS_ARCHIVE: 'ReservationsArchive',
    MACHINES: 'Machines',
    MACHINE_REPORTS: 'MachineReports',
    PUBLIC_BOARD: 'PublicBoard',
    AUDIT: 'Audit',
    AUDIT_ARCHIVE: 'AuditArchive',
  },
  HEADERS: {
    SETTINGS: ['key', 'value', 'notes'],
    RESIDENTS: [
      'resident_id',
      'room',
      'display_name',
      'access_code',
      'active',
      'notes',
      'created_at',
      'updated_at',
      'recovery_hash',
      'registered_at',
      'last_login_at',
      'last_activity_at',
      'disabled_at',
      'disabled_reason'
    ],
    MACHINES: [
      'machine_no',
      'status',
      'status_changed_at',
      'status_note'
    ],
    MACHINE_REPORTS: [
      'report_id',
      'machine_no',
      'resident_id',
      'room',
      'display_name',
      'reservation_id',
      'slot_id',
      'reported_at',
      'description',
      'status',
      'resolved_at',
      'resolved_by_resident_id',
      'resolution_note'
    ],
    SLOTS: ['slot_id', 'week_key', 'slot_date', 'day_label', 'slot_label', 'start_iso', 'end_iso', 'capacity', 'is_open', 'cycle_start_date', 'created_at'],
    RESERVATIONS: ['reservation_id', 'resident_id', 'room', 'display_name', 'slot_id', 'week_key', 'status', 'created_at', 'updated_at', 'note', 'machine_no', 'resident_label'],
    PUBLIC_BOARD: ['Day / Dan', 'Date / Datum', 'Machine / Stroj', '2:00 - 6:00', '6:00 - 10:00', '10:00 - 14:00', '14:00 - 18:00', '18:00 - 22:00', '22:00 - 2:00'],
    AUDIT: ['timestamp', 'actor', 'action', 'target_id', 'details'],
  },
  SLOT_TEMPLATES: [
    { startHour: 2, label: '02:00–06:00' },
    { startHour: 6, label: '06:00–10:00' },
    { startHour: 10, label: '10:00–14:00' },
    { startHour: 14, label: '14:00–18:00' },
    { startHour: 18, label: '18:00–22:00' },
    { startHour: 22, label: '22:00–02:00' },
  ],
  SLOT_CAPACITY: 4,
  MACHINE_REPORT_STATUSES: {
    OPEN: 'OPEN',
    DISMISSED: 'DISMISSED',
    MACHINE_OUT_OF_ORDER: 'MACHINE_OUT_OF_ORDER',
    MACHINE_RESTORED: 'MACHINE_RESTORED',
  },
  MACHINE_REPORT_DESCRIPTION_MAX_LENGTH: 500,
  CYCLE_DAYS: 14,
  ACCESS_CODE_LENGTH: 6,
  AUTOMATION_CHECK_EVERY_MINUTES: 10,
  AUTOMATION_WINDOWS: {
    OPEN_START: '19:00',
    OPEN_END: '19:15',
    FINALIZE_START: '02:05',
    FINALIZE_END: '02:20',
    ARCHIVE_START: '03:00',
    ARCHIVE_END: '03:15',
  },
};

let _SS_CACHE = null;
let _SETTINGS_CACHE = null;
let _HEADER_MAP_CACHE = {};
let _SCHEMA_READY_CACHE = false;
let _RESIDENTS_CACHE = null;
let _SLOTS_CACHE = null;
let _RESERVATIONS_CACHE = null;
let _SLOT_BY_ID_CACHE = null;
let _RESERVATION_BY_ID_CACHE = null;
let _CONFIRMED_RESERVATION_COUNTS_BY_SLOT_CACHE = null;
let _CONFIRMED_RESERVATIONS_BY_RESIDENT_CACHE = null;
let _DATA_CACHE_VERSION_CACHE = null;
let _CONFIRMED_RESERVATIONS_BY_SLOT_CACHE = null;
let _MACHINES_CACHE = null;
let _MACHINE_REPORTS_CACHE = null;

function doGet(e) {
  ensureSchema_();

  const page = String((e && e.parameter && e.parameter.page) || '').toLowerCase();
  const fileName = page === 'admin' ? APP_CONFIG.ADMIN_PAGE : 'Index';

  const template = HtmlService.createTemplateFromFile(fileName);
  template.APP_BASE_URL = getAppBaseUrl();

  return template
    .evaluate()
    .setTitle(APP_CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAppBaseUrl() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    throw new Error('Web app URL is unavailable. Deploy this project as a Web app and open the deployed /exec URL.');
  }
  return url;
}

function getAppPageUrl(page) {
  const base = getAppBaseUrl();
  return page
    ? `${base}?page=${encodeURIComponent(String(page).trim())}`
    : base;
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(APP_CONFIG.MENU_NAME)
    .addItem('1. Initialize project', 'initializePrototype')
    .addItem('2. Upgrade existing sheet to current schema', 'upgradePrototypeSchema')
    .addSeparator()
    .addItem('3. Seed demo residents', 'seedDemoResidents')
    .addItem('4. Generate codes for blank residents', 'generateCodesForBlankResidents')
    .addItem('5. Open next 2-week cycle now', 'openNextCycleNow')
    .addItem('6. Finalize carryover display now', 'finalizeCarryoverDisplayNow')
    .addItem('7. Rebuild cycle availability', 'rebuildCycleAvailability')
    .addItem('8. Refresh public board now', 'refreshPublicBoardNow')
    .addItem('9. Archive old data now', 'archiveOldDataNow')
    .addSeparator()
    .addItem('10. Install automation triggers', 'installAutomationTriggers')
    .addItem('11. Remove automation triggers', 'clearAutomationTriggers')
    .addSeparator()
    .addItem('12. Reset resident code by resident ID', 'resetResidentCodePrompt')
    .addItem('13. Simulate next cycle opening now', 'simulateNextCycleSwitchNow')
    .addItem('Open user web app help', 'showWebAppHint')
    .addItem('14. Set admin PIN', 'setAdminPinPrompt')
    .addToUi();
}

// ---------------------------
// Admin entry points
// ---------------------------

function initializePrototype() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Open the bound spreadsheet and run this from there.');
  PropertiesService.getScriptProperties().setProperty(APP_CONFIG.SPREADSHEET_ID_KEY, ss.getId());
  ensureSchema_(true);
  notifyAdmin_(
    `${APP_CONFIG.APP_NAME} initialized.\n\n` +
    'Recommended next steps:\n' +
    '1) Seed demo residents or let residents self-register.\n' +
    '2) Open the next 2-week cycle.\n' +
    '3) Install automation triggers.\n' +
    '4) Deploy the Apps Script web app.'
  );
}

function upgradePrototypeSchema() {
  ensureSchema_(true);
  syncReservationMetadata_();
  rebuildCycleAvailability();
  refreshPublicBoardFull_();
  clearQueuedPublicBoardRefresh_();
  removePublicBoardRefreshTrigger_();
  notifyAdmin_(`${APP_CONFIG.APP_NAME} schema verified and upgraded.`);
}

function seedDemoResidents() {
  ensureSchema_();
  const info = getResidentsSheetInfo_();
  const { sheet, values } = info;
  if (values.length > 1) {
    const ui = getUiOrNull_();
    if (ui) {
      const response = ui.alert(
        'The Residents sheet already contains data. Add demo rows anyway?',
        ui.ButtonSet.YES_NO
      );
      if (response !== ui.Button.YES) return;
    }
  }

  const existingIds = new Set(getAllResidents_().map(r => r.residentId));
  const demos = [
    { residentId: 'RES-001', room: '101', displayName: 'Janez Novak', notes: 'demo resident', recoveryHash: hashRecoveryKey_('demo101') },
    { residentId: 'RES-002', room: '102', displayName: 'Ana Demo', notes: 'demo resident', recoveryHash: hashRecoveryKey_('demo102') },
    { residentId: 'RES-003', room: '103', displayName: 'Boris Demo', notes: 'demo resident', recoveryHash: hashRecoveryKey_('demo103') },
    { residentId: 'RES-004', room: '104', displayName: 'Cene Demo', notes: 'demo resident', recoveryHash: hashRecoveryKey_('demo104') },
  ];

  const rows = demos
    .filter(item => !existingIds.has(item.residentId))
    .map(item => makeResidentRow_(Object.assign({ active: true }, item)));

  if (!rows.length) return;

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  clearExecutionCaches_();
  bumpDataCacheVersion_();
  generateCodesForBlankResidents();
  logAudit_('admin', 'seed_demo_residents', 'Residents', `Added ${rows.length} demo residents`);
}

function generateCodesForBlankResidents() {
  ensureSchema_();
  const info = getResidentsSheetInfo_();
  const { sheet, values, map } = info;
  if (values.length < 2) return;

  const usedCodes = new Set(values.slice(1).map(r => String(getCellValue_(r, map, 'access_code') || '').trim()).filter(Boolean));
  let changed = 0;

  for (let i = 1; i < values.length; i++) {
    const active = normalizeBoolean_(getCellValue_(values[i], map, 'active'));
    const currentCode = String(getCellValue_(values[i], map, 'access_code') || '').trim();
    if (active && !currentCode) {
      const newCode = generateUniqueCode_(usedCodes, APP_CONFIG.ACCESS_CODE_LENGTH);
      setCellValue_(values[i], map, 'access_code', newCode);
      setCellValue_(values[i], map, 'updated_at', new Date());
      changed += 1;
    }
  }

  if (!changed) return;
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  clearExecutionCaches_();
  bumpDataCacheVersion_();
  logAudit_('admin', 'generate_codes', 'Residents', `Generated ${changed} access codes`);
}

function resetResidentCodePrompt() {
  const ui = getUiOrNull_();
  if (!ui) throw new Error('Open the bound spreadsheet to use this prompt.');
  const response = ui.prompt('Reset resident code', 'Enter resident_id exactly as shown in the Residents sheet:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  const residentId = String(response.getResponseText() || '').trim();
  if (!residentId) return;
  const newCode = resetResidentCodeByResidentId(residentId);
  ui.alert(newCode ? `New code for ${residentId}: ${newCode}` : `Resident ${residentId} not found.`);
}

function resetResidentCodeByResidentId(residentId) {
  ensureSchema_();
  const info = getResidentsSheetInfo_();
  const { sheet, values, map } = info;
  const usedCodes = new Set(values.slice(1).map(r => String(getCellValue_(r, map, 'access_code') || '').trim()).filter(Boolean));

  for (let i = 1; i < values.length; i++) {
    if (String(getCellValue_(values[i], map, 'resident_id') || '').trim() !== residentId) continue;
    const oldCode = String(getCellValue_(values[i], map, 'access_code') || '').trim();
    if (oldCode) usedCodes.delete(oldCode);
    const newCode = generateUniqueCode_(usedCodes, APP_CONFIG.ACCESS_CODE_LENGTH);
    setCellValue_(values[i], map, 'access_code', newCode);
    setCellValue_(values[i], map, 'updated_at', new Date());
    sheet.getRange(i + 1, 1, 1, values[i].length).setValues([values[i]]);
    clearExecutionCaches_();
    bumpDataCacheVersion_();
    logAudit_('admin', 'reset_resident_code', residentId, 'Access code reset');
    return newCode;
  }
  return null;
}

function openNextCycleNow() {
  ensureSchema_();
  const settings = getSettingsMap_();
  const activeStartText = normalizeDateCellToYmd_(settings.active_cycle_start);

  const cycleStart = activeStartText
    ? addDays_(parseYmdDate_(activeStartText), APP_CONFIG.CYCLE_DAYS)
    : getNextMonday_(new Date());

  openCycleAt_(cycleStart, { keepTodayVisible: false, actor: 'admin', action: 'open_next_cycle_manual' });
  notifyAdmin_(`Opened cycle ${formatDateYmd_(cycleStart)} – ${formatDateYmd_(addDays_(cycleStart, APP_CONFIG.CYCLE_DAYS - 1))}.`);
}

function simulateNextCycleSwitchNow() {
  ensureSchema_();
  const settings = getSettingsMap_();
  const activeStartText = normalizeDateCellToYmd_(settings.active_cycle_start);
  const activeStart = parseYmdDate_(activeStartText);
  if (!activeStartText || isNaN(activeStart.getTime())) {
    throw new Error('Settings.active_cycle_start is missing. Open the first cycle manually before simulating the next cycle switch.');
  }

  const nextCycleStart = addDays_(activeStart, APP_CONFIG.CYCLE_DAYS);
  openCycleAt_(nextCycleStart, { keepTodayVisible: true, actor: 'admin', action: 'simulate_next_cycle_open' });
  notifyAdmin_(
    `Simulated next cycle opening.\n\n` +
    `New active cycle: ${formatDateYmd_(nextCycleStart)} – ${formatDateYmd_(addDays_(nextCycleStart, APP_CONFIG.CYCLE_DAYS - 1))}.\n` +
    `Carryover display remains visible for today.`
  );
}

function openNextScheduledCycleNow() {
  ensureSchema_();
  const settings = getSettingsMap_();
  const activeStart = parseYmdDate_(settings.active_cycle_start);
  if (!activeStart) {
    throw new Error('Settings.active_cycle_start is missing. Open the first cycle manually before enabling automation.');
  }

  const today = todayInTz_();
  const todayYmd = formatDateYmd_(today);
  const expectedEnd = addDays_(activeStart, APP_CONFIG.CYCLE_DAYS - 1);
  const nextCycleStart = addDays_(activeStart, APP_CONFIG.CYCLE_DAYS);
  const expectedTomorrow = addDays_(today, 1);

  if (formatDateYmd_(expectedEnd) !== todayYmd || formatDateYmd_(nextCycleStart) !== formatDateYmd_(expectedTomorrow)) {
    // Not the rollover Sunday for the current 2-week cycle.
    return false;
  }

  openCycleAt_(nextCycleStart, { keepTodayVisible: true, actor: 'automation', action: 'open_next_cycle_scheduled' });
  return true;
}

function finalizeCarryoverDisplayNow() {
  ensureSchema_();
  const settings = getSettingsMap_();
  const activeCycleStart = normalizeDateCellToYmd_(settings.active_cycle_start);
  if (!activeCycleStart) throw new Error('No active cycle start is set.');

  setSetting_('public_display_start', activeCycleStart, 'Public board now starts at the active cycle start.');
  setSetting_('carryover_finalize_at', '', 'Internal exact datetime used to end carryover display after a cycle opens');

  // Important: at finalize time, archive reservations whose slots have actually ended already.
  archiveOldDataNow(false, {
    forceEndedBeforeIso: nowLocalIso_()
  });

  refreshPublicBoardFull_();
  clearQueuedPublicBoardRefresh_();

  logAudit_('automation', 'finalize_carryover_display', activeCycleStart, 'Carryover Sunday removed from public display');
  notifyAdmin_(`Carryover display closed. PublicBoard now starts at ${activeCycleStart}.`);
}

function rebuildCycleAvailability() {
  ensureSchema_();
  const info = getSlotsSheetInfo_();
  const { sheet, values, map } = info;
  if (values.length < 2) return;
  const settings = getSettingsMap_();
  const bookingOpen = normalizeBoolean_(settings.booking_open);
  const cycleStart = normalizeDateCellToYmd_(settings.active_cycle_start);

  let changed = 0;
  for (let i = 1; i < values.length; i++) {
    const rowCycleStart = normalizeDateCellToYmd_(getCellValue_(values[i], map, 'cycle_start_date'));
    const nextValue = bookingOpen && rowCycleStart === cycleStart;
    if (normalizeBoolean_(getCellValue_(values[i], map, 'is_open')) !== nextValue) {
      setCellValue_(values[i], map, 'is_open', nextValue);
      changed += 1;
    }
  }
  if (!changed) return;
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  clearExecutionCaches_();
  bumpDataCacheVersion_();
}

function refreshPublicBoardNow() {
  ensureSchema_();
  syncReservationMetadata_();
  refreshPublicBoardFull_();
  clearQueuedPublicBoardRefresh_();
  notifyAdmin_('Public board refreshed.');
}

function archiveOldDataNow(showNotice, options) {
  options = options || {};
  ensureSchema_();

  const movedReservations = archivePastReservationsNow(options);
  const disabledResidents = disableInactiveResidentsNow();
  const deletedResidents = deleteOldDisabledResidentsNow();
  const prunedSlots = pruneStaleSlotsNow();
  const trimmedAudit = trimAuditActiveNow();

  if (showNotice !== false) {
    notifyAdmin_(
      `Archive complete.\n` +
      `Archived reservations: ${movedReservations}\n` +
      `Auto-disabled residents: ${disabledResidents}\n` +
      `Deleted disabled residents: ${deletedResidents}\n` +
      `Pruned stale slots: ${prunedSlots}\n` +
      `Archived audit rows: ${trimmedAudit}`
    );
  }

  return {
    movedReservations,
    disabledResidents,
    deletedResidents,
    prunedSlots,
    trimmedAudit
  };
}

function showWebAppHint() {
  const message =
    `Deploy ${APP_CONFIG.APP_NAME} as a web app:\n\n` +
    'Deploy → New deployment → Web app\n' +
    'Execute as: Me\n' +
    'Who has access: Anyone with the link\n\n' +
    'After handing the project over, the next admin should redeploy under their own account and recreate automation triggers.';
  notifyAdmin_(message);
}

// ---------------------------
// Automation
// ---------------------------

function installAutomationTriggers() {
  clearAutomationTriggers_(true);
  ScriptApp.newTrigger('automationDispatcher')
    .timeBased()
    .everyMinutes(APP_CONFIG.AUTOMATION_CHECK_EVERY_MINUTES)
    .create();
  notifyAdmin_('Automation trigger installed. PublicBoard refresh trigger is now created only on demand when a queued rebuild is needed.');
}

function clearAutomationTriggers() {
  clearAutomationTriggers_(false);
}

function clearAutomationTriggers_(silent) {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    const handler = trigger.getHandlerFunction();
    if (handler === 'automationDispatcher' || handler === APP_CONFIG.PUBLIC_BOARD_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  PropertiesService.getScriptProperties().deleteProperty(APP_CONFIG.PUBLIC_BOARD_TRIGGER_INSTALLED_KEY);
  if (!silent) notifyAdmin_('Automation triggers removed.');
}

function installPublicBoardRefreshTrigger_() {
  const props = PropertiesService.getScriptProperties();
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(trigger => trigger.getHandlerFunction() === APP_CONFIG.PUBLIC_BOARD_TRIGGER_HANDLER);

  if (!exists) {
    ScriptApp.newTrigger(APP_CONFIG.PUBLIC_BOARD_TRIGGER_HANDLER)
      .timeBased()
      .everyMinutes(APP_CONFIG.PUBLIC_BOARD_TRIGGER_EVERY_MINUTES)
      .create();
  }

  props.setProperty(APP_CONFIG.PUBLIC_BOARD_TRIGGER_INSTALLED_KEY, 'true');
}

function removePublicBoardRefreshTrigger_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === APP_CONFIG.PUBLIC_BOARD_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  PropertiesService.getScriptProperties().deleteProperty(APP_CONFIG.PUBLIC_BOARD_TRIGGER_INSTALLED_KEY);
}

function automationDispatcher() {
  ensureSchema_();
  ensureCycleOpenedIfDue_();

  const now = new Date();
  const tz = APP_CONFIG.TZ;
  const day = Utilities.formatDate(now, tz, 'u');
  const hhmm = Utilities.formatDate(now, tz, 'HH:mm');
  const ymd = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const nowIso = nowLocalIso_();
  const settings = getSettingsMap_();
  const carryoverFinalizeAt = normalizeDateTimeCellToIso_(settings.carryover_finalize_at);

  if (carryoverFinalizeAt && carryoverFinalizeAt <= nowIso) {
    finalizeCarryoverDisplayNow();
  } else if (!carryoverFinalizeAt && day === '1' && isWithinWindow_(hhmm, APP_CONFIG.AUTOMATION_WINDOWS.FINALIZE_START, APP_CONFIG.AUTOMATION_WINDOWS.FINALIZE_END)) {
    runOncePerDay_('finalize', ymd, () => finalizeCarryoverDisplayNow());
  }

  if (isWithinWindow_(hhmm, APP_CONFIG.AUTOMATION_WINDOWS.ARCHIVE_START, APP_CONFIG.AUTOMATION_WINDOWS.ARCHIVE_END)) {
    runOncePerDay_('archive', ymd, () => archiveOldDataNow(false));
  }
}

function publicBoardDispatcher() {
  ensureSchema_();
  const processed = flushPendingPublicBoardRefresh_();
  if (!getQueuedPublicBoardRefresh_()) {
    removePublicBoardRefreshTrigger_();
  }
  return processed;
}

function runOncePerDay_(jobName, ymd, fn) {
  const props = PropertiesService.getScriptProperties();
  const key = `${APP_CONFIG.AUTOMATION_KEYS_PREFIX}_${jobName}_${ymd}`;
  if (props.getProperty(key)) return false;
  const result = fn();
  props.setProperty(key, new Date().toISOString());
  return result;
}

function runOnceByKey_(jobName, keyPart, fn) {
  const props = PropertiesService.getScriptProperties();
  const key = `${APP_CONFIG.AUTOMATION_KEYS_PREFIX}_${jobName}_${keyPart}`;
  if (props.getProperty(key)) return false;
  const result = fn();
  props.setProperty(key, new Date().toISOString());
  return result;
}

function clearAutomationState() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(key => {
    if (key.startsWith(APP_CONFIG.AUTOMATION_KEYS_PREFIX + '_')) {
      props.deleteProperty(key);
    }
  });
  notifyAdmin_('Automation run-state keys cleared.');
}

function makeRetryableError_(message) {
  return new Error(`${message} ${APP_CONFIG.RETRYABLE_ERROR_TOKEN}`);
}

function throwRetryableError_(message) {
  throw makeRetryableError_(message);
}

function acquireScriptLockOrThrow_(timeoutMs, busyMessage) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(Math.max(0, Number(timeoutMs) || 0))) {
    throwRetryableError_(busyMessage || 'The scheduler is busy. Please retry in a moment.');
  }
  return lock;
}

function isCycleSwitchDueForSettings_(settings, nowIso) {
  settings = settings || {};
  const activeStartText = normalizeDateCellToYmd_(settings.active_cycle_start);
  const activeStart = parseYmdDate_(activeStartText);
  if (!activeStartText || isNaN(activeStart.getTime())) return false;

  const currentIso = normalizeDateTimeCellToIso_(nowIso || nowLocalIso_());
  const simulationOpenAt = getSimulationOpenIso_();
  const openingTime = normalizeHhMm_(settings.opening_time) || '19:00';
  const cycleEndYmd = formatDateYmd_(addDays_(activeStart, APP_CONFIG.CYCLE_DAYS - 1));

  const scheduledOpenIso = cycleEndYmd ? `${cycleEndYmd}T${openingTime}:00` : '';
  return Boolean(
    (simulationOpenAt && simulationOpenAt <= currentIso) ||
    (scheduledOpenIso && scheduledOpenIso <= currentIso)
  );
}

function waitForCycleSwitchToSettle_() {
  const deadline = Date.now() + APP_CONFIG.CYCLE_SWITCH_SETTLE_TOTAL_MS;
  while (Date.now() < deadline) {
    Utilities.sleep(APP_CONFIG.CYCLE_SWITCH_SETTLE_STEP_MS);
    clearExecutionCaches_();
    if (!isCycleSwitchDueForSettings_(getSettingsMap_(), nowLocalIso_())) {
      return true;
    }
  }
  return false;
}

// ---------------------------
// Public API used by Index.html
// ---------------------------

function loginResident(room, accessCode) {
  ensureSchema_();
  ensureCycleOpenedIfDue_();
  room = normalizeRoom_(room);
  accessCode = normalizeAccessCode_(accessCode);
  if (!room || !accessCode) throw new Error('Enter your room number and access code.');
  assertAllowedRoom_(room);

  const resident = findResidentByRoomAndCode_(room, accessCode);
  if (!resident) throw new Error('Invalid room number or access code.');

  try {
    touchResidentLastLogin_(resident.residentId);
    touchResidentActivity_(resident.residentId);
    resident.lastLoginAt = new Date();
    resident.lastActivityAt = new Date();
  } catch (err) {}

  return decorateDashboardAuth_(getDashboardDataCached_(resident), resident);
}

function refreshDashboard(room, accessCode) {
  ensureSchema_();
  ensureCycleOpenedIfDue_();
  const resident = requireResident_(room, accessCode);

  try {
    touchResidentActivity_(resident.residentId);
    resident.lastActivityAt = new Date();
  } catch (err) {}

  return decorateDashboardAuth_(getDashboardDataCached_(resident), resident);
}

function saveRecoveryKey(room, accessCode, recoveryKey) {
  ensureSchema_();
  const resident = requireResident_(room, accessCode);
  recoveryKey = String(recoveryKey || '').trim();
  if (recoveryKey.length < 4) throw new Error('Choose a recovery key with at least 4 characters.');

  const recoveryHash = hashRecoveryKey_(recoveryKey);
  const info = getResidentsSheetInfo_();
  const { sheet, values, map } = info;
  for (let i = 1; i < values.length; i++) {
    if (String(getCellValue_(values[i], map, 'resident_id') || '').trim() !== resident.residentId) continue;
    setCellValue_(values[i], map, 'recovery_hash', recoveryHash);
    setCellValue_(values[i], map, 'updated_at', new Date());
    setCellValue_(values[i], map, 'last_activity_at', new Date());
    sheet.getRange(i + 1, 1, 1, values[i].length).setValues([values[i]]);
    clearExecutionCaches_();
    bumpDataCacheVersion_();
    logAudit_(resident.residentId, 'save_recovery_key', resident.residentId, 'Recovery key updated');
    break;
  }

  resident.recoveryHash = recoveryHash;
  return decorateDashboardAuth_(getDashboardData_(resident), resident);
}

function selfRegisterResident(room, displayName, recoveryKey) {
  ensureSchema_();
  const settings = getSettingsMap_();
  if (!normalizeBoolean_(settings.self_signup_enabled)) {
    throw new Error('Self-service signup is currently disabled.');
  }

  room = normalizeRoom_(room);
  displayName = normalizeDisplayName_(displayName);
  recoveryKey = String(recoveryKey || '').trim();
  validateSignupInputs_(room, displayName, recoveryKey);

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Registration is busy right now. Please retry in a moment.'
  );
  try {
    const existing = findResidentByRoomAndDisplayName_(room, displayName);
    if (existing && existing.active) {
      throw new Error('A resident with this room and name is already registered. Use sign in or reset code instead.');
    }

    const maxPerRoom = Math.max(
      1,
      Number(settings.max_active_residents_per_room || 2) || 2
    );

    const activeResidentsInRoom = countActiveResidentsInRoom_(room);
    if (activeResidentsInRoom >= maxPerRoom) {
      throw new Error(
        `This room already has the maximum number of active residents (${maxPerRoom}). ` +
        `Please contact the dorm representative if an old account should be removed.`
      );
    }

    const usedCodes = new Set(getAllResidents_().map(r => r.accessCode).filter(Boolean));
    const accessCode = generateUniqueCode_(usedCodes, APP_CONFIG.ACCESS_CODE_LENGTH);
    const residentId = generateResidentId_();
    const row = makeResidentRow_({
      residentId,
      room,
      displayName,
      accessCode,
      active: true,
      notes: 'self-service signup',
      recoveryHash: hashRecoveryKey_(recoveryKey),
      registeredAt: new Date(),
      lastLoginAt: new Date(),
      lastActivityAt: new Date(),
      disabledAt: '',
      disabledReason: '',
    });

    const info = getResidentsSheetInfo_();
    info.sheet.getRange(info.sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
    clearExecutionCaches_();
    bumpDataCacheVersion_();

    logAudit_(residentId, 'self_register', residentId, `${displayName} registered from room ${room}`);
    const resident = findResidentByRoomAndCode_(room, accessCode);
    const dashboard = decorateDashboardAuth_(getDashboardData_(resident), resident);
    dashboard.signupResult = { accessCode };
    return dashboard;
  } finally {
    lock.releaseLock();
  }
}

function selfResetAccessCode(room, displayName, recoveryKey) {
  ensureSchema_();
  const settings = getSettingsMap_();
  if (!normalizeBoolean_(settings.self_reset_enabled)) {
    throw new Error('Self-service code reset is currently disabled.');
  }

  room = normalizeRoom_(room);
  displayName = normalizeDisplayName_(displayName);
  recoveryKey = String(recoveryKey || '').trim();
  if (!room || !displayName || !recoveryKey) {
    throw new Error('Enter your room number, name, and recovery key.');
  }
  assertAllowedRoom_(room);

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Code reset is busy right now. Please retry in a moment.'
  );
  try {
    const resident = findResidentByRecoveryKey_(room, displayName, recoveryKey);
    if (!resident) throw new Error('We could not verify those reset details.');

    const newCode = resetResidentCodeByResidentId(resident.residentId);

    touchResidentLastLogin_(resident.residentId);
    touchResidentActivity_(resident.residentId, { force: true });

    const freshResident = findResidentByRoomAndCode_(resident.room, newCode);
    const dashboard = decorateDashboardAuth_(getDashboardData_(freshResident), freshResident);
    dashboard.resetResult = { accessCode: newCode };
    return dashboard;
  } finally {
    lock.releaseLock();
  }
}

function claimSlot(room, accessCode, slotId) {
  ensureSchema_();
  ensureCycleOpenedIfDue_();
  const resident = requireResident_(room, accessCode);

  let slot;
  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Too many residents are changing reservations right now. Please retry in a moment.'
  );
  try {
    // Reload from Sheets only AFTER the script lock is acquired.
    // Another request may have created a reservation while we waited.
    getAllReservations_(true);
    slot = findSlotById_(slotId);
    if (!slot) throw new Error('Selected slot was not found.');

    if (isSlotConcluded_(slot)) {
      throw new Error('This slot has already concluded and can no longer be claimed.');
    }

    if (!slot.isOpen || !normalizeBoolean_(getSettingsMap_().booking_open)) {
      throw new Error('This slot is not open for booking right now.');
    }
    if (!getNextMachineNoForSlot_(slot.slotId)) {
      throw new Error('This slot has no working machine available.');
    }

    const existingForWeek = getConfirmedReservationsForResident_(resident.residentId)
      .filter(r => r.weekKey === slot.weekKey);

    if (existingForWeek.length > 0) {
      throw new Error('You already used or reserved a laundry slot in this week.');
    }

    const created = createReservation_(resident, slot, 'confirmed', 'Claimed from web app');
    logAudit_(resident.residentId, 'claim_slot', slot.slotId, `${resident.displayName} claimed ${slot.dayLabel} ${slot.slotLabel} on machine ${created.machineNo}`);
  } finally {
    lock.releaseLock();
  }

  touchResidentActivity_(resident.residentId, { force: true });
  syncPublicBoardAfterReservationChange_([normalizeDateCellToYmd_(slot.slotDate)]);
  return decorateDashboardAuth_(getDashboardData_(resident), resident);
}

function cancelReservation(room, accessCode, reservationId) {
  ensureSchema_();
  ensureCycleOpenedIfDue_();
  const resident = requireResident_(room, accessCode);

  let slotDate = '';
  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Too many residents are changing reservations right now. Please retry in a moment.'
  );
  try {
    const reservation = findReservationById_(reservationId);
    if (!reservation || reservation.status !== 'confirmed') throw new Error('Reservation was not found.');
    if (reservation.residentId !== resident.residentId) throw new Error('You can only cancel your own reservation.');

    const slot = findSlotById_(reservation.slotId);
    if (slot && isSlotConcluded_(slot)) {
      throw new Error('This reservation has already concluded and can no longer be cancelled.');
    }
    if (slot && isSlotChangeLocked_(slot)) {
      throw new Error('This reservation can no longer be cancelled. The change cutoff has passed.');
    }

    slotDate = slot ? normalizeDateCellToYmd_(slot.slotDate) : '';
    updateReservationStatus_(reservationId, 'cancelled', 'Cancelled by resident');
    logAudit_(resident.residentId, 'cancel_reservation', reservationId, `${resident.displayName} cancelled reservation`);
  } finally {
    lock.releaseLock();
  }

  touchResidentActivity_(resident.residentId, { force: true });
  syncPublicBoardAfterReservationChange_(slotDate ? [slotDate] : []);
  return decorateDashboardAuth_(getDashboardData_(resident), resident);
}

function moveReservation(room, accessCode, reservationId, newSlotId) {
  ensureSchema_();
  ensureCycleOpenedIfDue_();
  const resident = requireResident_(room, accessCode);

  let changedDates = [];
  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Too many residents are changing reservations right now. Please retry in a moment.'
  );
  try {
    // Do not decide on the old reservation or target capacity from a stale cache.
    getAllReservations_(true);
    const reservation = findReservationById_(reservationId);
    if (!reservation || reservation.status !== 'confirmed') throw new Error('Current reservation was not found.');
    if (reservation.residentId !== resident.residentId) throw new Error('You can only move your own reservation.');

    const currentSlot = findSlotById_(reservation.slotId);
    if (currentSlot && isSlotConcluded_(currentSlot)) {
      throw new Error('This reservation has already concluded and can no longer be moved.');
    }
    if (currentSlot && isSlotChangeLocked_(currentSlot)) {
      throw new Error('This reservation can no longer be moved. The change cutoff has passed.');
    }

    const newSlot = findSlotById_(newSlotId);
    if (!newSlot) throw new Error('Target slot was not found.');

    if (isSlotConcluded_(newSlot)) {
      throw new Error('Target slot has already concluded and can no longer be selected.');
    }

    if (!newSlot.isOpen || !normalizeBoolean_(getSettingsMap_().booking_open)) {
      throw new Error('Target slot is not open for booking right now.');
    }

    if (!getNextMachineNoForSlot_(newSlot.slotId)) {
      throw new Error('Target slot has no working machine available.');
    }

    const confirmed = getConfirmedReservationsForResident_(resident.residentId);
    const otherInTargetWeek = confirmed
      .filter(r => r.weekKey === newSlot.weekKey && r.reservationId !== reservationId);

    if (otherInTargetWeek.length > 0) {
      throw new Error('You already used or reserved a laundry slot in the target week.');
    }

    changedDates = [
      currentSlot ? normalizeDateCellToYmd_(currentSlot.slotDate) : '',
      normalizeDateCellToYmd_(newSlot.slotDate),
    ].filter(Boolean);

    updateReservationStatus_(reservationId, 'moved', `Moved to ${newSlot.slotId}`);
    const created = createReservation_(resident, newSlot, 'confirmed', `Move from ${reservation.slotId}`);
    logAudit_(resident.residentId, 'move_reservation', reservationId, `${resident.displayName} moved to ${newSlot.dayLabel} ${newSlot.slotLabel} on machine ${created.machineNo}`);
  } finally {
    lock.releaseLock();
  }

  touchResidentActivity_(resident.residentId, { force: true });
  syncPublicBoardAfterReservationChange_(changedDates);
  return decorateDashboardAuth_(getDashboardData_(resident), resident);
}

function getActiveCycleSummary() {
  ensureSchema_();
  const settings = getSettingsMap_();
  return {
    settings,
    slotCount: getAllSlots_().length,
    reservationCount: getAllReservations_().filter(r => r.status === 'confirmed').length,
    publicBoardRows: Math.max(0, getBackendSpreadsheet_().getSheetByName(APP_CONFIG.SHEETS.PUBLIC_BOARD).getLastRow() - 2),
  };
}

// ---------------------------
// Dashboard builders
// ---------------------------

function buildMachineReportOptions_() {
  const openMachineReportCounts = getOpenMachineReportCountsByMachine_();

  return getMachines_().map(machine => ({
    machineNo: machine.machineNo,
    status: machine.status,
    openReportCount: openMachineReportCounts[machine.machineNo] || 0,
  }));
}

function decorateDashboardAuth_(dashboard, resident) {
  dashboard.auth = { room: resident.room, accessCode: resident.accessCode };

  // Machine report controls must stay current even when the main dashboard
  // payload was served from Apps Script cache.
  dashboard.machineReportOptions = buildMachineReportOptions_();

  return dashboard;
}

function getMachineReportOptions(room, accessCode) {
  ensureSchema_();
  requireResident_(room, accessCode);
  return buildMachineReportOptions_();
}

function getSimulationOpenIso_() {
  const settings = getSettingsMap_();
  const raw = settings.simulation_open_at;

  if (raw === '' || raw === null || raw === undefined) return '';
  if (typeof raw === 'number' && raw === 0) return '';

  const iso = normalizeDateTimeCellToIso_(raw);
  if (!iso) return '';

  const dt = valueToDateTime_(iso);
  if (isNaN(dt.getTime())) return '';
  if (dt.getFullYear() < 2024) return '';

  return iso;
}


function getNextCycleOpenIso_() {
  const simulationIso = getSimulationOpenIso_();
  if (simulationIso) return simulationIso;

  const settings = getSettingsMap_();
  const cycleStartText = normalizeDateCellToYmd_(settings.active_cycle_start);
  if (!cycleStartText) return '';

  const cycleDays = Number(settings.cycle_days || APP_CONFIG.CYCLE_DAYS) || APP_CONFIG.CYCLE_DAYS;
  const openingTime = normalizeHhMm_(settings.opening_time) || '19:00';

  const openDate = addDays_(parseYmdDate_(cycleStartText), cycleDays - 1);
  const [hh, mm] = openingTime.split(':').map(Number);
  openDate.setHours(hh, mm, 0, 0);

  return Utilities.formatDate(openDate, APP_CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss");
}

function getDashboardData_(resident) {
  const settings = getSettingsMap_();

  const bookingOpen = normalizeBoolean_(settings.booking_open);
  const cycleStartText = normalizeDateCellToYmd_(settings.active_cycle_start);
  const cycleDays = Number(settings.cycle_days || APP_CONFIG.CYCLE_DAYS) || APP_CONFIG.CYCLE_DAYS;
  const cycleEndText = cycleStartText ? formatDateYmd_(addDays_(parseYmdDate_(cycleStartText), cycleDays - 1)) : '';
  const carryoverDate = normalizeDateCellToYmd_(settings.public_display_start);

  const allSlots = getAllSlots_();
  const slotById = getSlotByIdMap_();
  const reservationCounts = getConfirmedReservationCountsBySlot_();
  const residentReservations = getConfirmedReservationsForResident_(resident.residentId);
  const openMachineReportCounts = getOpenMachineReportCountsByMachine_();

  const activeSlots = allSlots
    .filter(slot => normalizeDateCellToYmd_(slot.cycleStartDate) === cycleStartText)
    .sort((a, b) => a.startIso.localeCompare(b.startIso));


  // Any confirmed reservation in the week blocks further claims for that week.
  const residentReservationsByWeek = {};
  residentReservations.forEach(r => {
    residentReservationsByWeek[r.weekKey] = r;
  });

  const groupedDays = {};

  activeSlots.forEach(slot => {
    slot.reservedCount = reservationCounts[slot.slotId] || 0;
    slot.operationalCapacity = getWorkingMachineNumbers_().length;
    slot.availableCount = getBookableMachineCountForSlot_(slot.slotId);

    if (!groupedDays[slot.slotDate]) groupedDays[slot.slotDate] = [];

    const ownReservation = residentReservationsByWeek[slot.weekKey] || null;
    const ownReservationSlot = ownReservation ? slotById[ownReservation.slotId] : null;
    const ownReservationCanChange = ownReservationSlot ? canModifyReservationForSlot_(ownReservationSlot) : false;
    const slotConcluded = isSlotConcluded_(slot);
    const ownIsCurrentSlot = ownReservation ? ownReservation.slotId === slot.slotId : false;

    const ownAvailableMachineNos =
      ownIsCurrentSlot && ownReservation
        ? getAvailableMachineNumbersForSlot_(slot.slotId, ownReservation.reservationId)
        : [];

    const ownShowChangeMachine =
      Boolean(ownIsCurrentSlot && ownReservation && canChangeMachineForSlot_(slot));

    const ownCanChangeMachine =
      Boolean(ownShowChangeMachine && ownAvailableMachineNos.length > 0);

    const ownMachineOpenReportCount =
      ownIsCurrentSlot && ownReservation
        ? (openMachineReportCounts[Number(ownReservation.machineNo || 0)] || 0)
        : 0;

    groupedDays[slot.slotDate].push({
      slotId: slot.slotId,
      slotDate: slot.slotDate,
      dayLabel: slot.dayLabel,
      slotLabel: slot.slotLabel,
      startIso: slot.startIso,
      endIso: slot.endIso,
      weekKey: slot.weekKey,
      capacity: slot.capacity,
      operationalCapacity: slot.operationalCapacity,
      reservedCount: slot.reservedCount,
      availableCount: slot.availableCount,
      isOpen: bookingOpen && slot.isOpen,
      isMine: ownIsCurrentSlot,
      ownReservationId: ownReservation ? ownReservation.reservationId : '',
      ownMachineNo: ownIsCurrentSlot && ownReservation ? ownReservation.machineNo : '',
      ownMachineWorking:
        ownIsCurrentSlot && ownReservation
          ? isMachineWorking_(ownReservation.machineNo)
          : true,
      hasOwnReservationThisWeek: Boolean(ownReservation),
      canCancel: Boolean(ownIsCurrentSlot && ownReservationCanChange),
      isConcluded: slotConcluded,
      isChangeLocked: Boolean(ownIsCurrentSlot && !slotConcluded && !ownReservationCanChange),
      availableMachineNos: ownAvailableMachineNos,
      showChangeMachine: ownShowChangeMachine,
      canChangeMachine: ownCanChangeMachine,
      machineOpenReportCount: ownMachineOpenReportCount,
      action: resolveSlotAction_(slot, ownReservation, ownReservationCanChange, bookingOpen),
    });
  });

  Object.keys(groupedDays).forEach(dateKey => {
    groupedDays[dateKey].sort((a, b) => a.startIso.localeCompare(b.startIso));
  });

  const reservationCards = residentReservations
    .map(r => {
      const slot = slotById[r.slotId] || {};
      return { reservation: r, slot };
    })
    .filter(item => item.slot && item.slot.slotId)
    .filter(item => {
      const slotDate = normalizeDateCellToYmd_(item.slot.slotDate);
      return normalizeDateCellToYmd_(item.slot.cycleStartDate) === cycleStartText || (carryoverDate && slotDate === carryoverDate);
    })
    .sort((a, b) => (a.slot.startIso || '').localeCompare(b.slot.startIso || ''))
    .map(item => {
      const r = item.reservation;
      const slot = item.slot;
      const canCancel = canModifyReservationForSlot_(slot);
      const isConcluded = isSlotConcluded_(slot);
      const isChangeLocked = !isConcluded && !canCancel;

      const availableMachineNos = getAvailableMachineNumbersForSlot_(slot.slotId, r.reservationId);
      const showChangeMachine = !isConcluded && canChangeMachineForSlot_(slot);
      const canChangeMachine = showChangeMachine && availableMachineNos.length > 0;
      const machineOpenReportCount =
        openMachineReportCounts[Number(r.machineNo || 0)] || 0;

      return {
        reservationId: r.reservationId,
        residentId: r.residentId,
        room: r.room,
        displayName: r.displayName,
        slotId: r.slotId,
        weekKey: r.weekKey,
        status: r.status,
        note: r.note,
        startIso: slot.startIso || '',
        endIso: slot.endIso || '',
        dayLabel: slot.dayLabel || '',
        slotLabel: slot.slotLabel || '',
        machineNo: r.machineNo,
        machineWorking: isMachineWorking_(r.machineNo),
        residentLabel: r.residentLabel,
        createdAtIso: valueToIsoString_(r.createdAt),
        updatedAtIso: valueToIsoString_(r.updatedAt),
        canCancel,
        isConcluded,
        isChangeLocked,
        availableMachineNos,
        showChangeMachine,
        canChangeMachine,
        machineOpenReportCount,
      };
    });

  return {
    resident: {
      residentId: resident.residentId,
      room: resident.room,
      displayName: resident.displayName,
      accessCode: resident.accessCode,
      hasRecoveryKey: Boolean(resident.recoveryHash),
      isAdmin: canResidentOpenAdminDashboard_(resident),
    },
    bookingOpen,
    cycleStartText,
    cycleEndText,
    nextCycleOpenIso: getNextCycleOpenIso_(),
    serverNowIso: nowLocalIso_(),
    machineReportOptions: buildMachineReportOptions_(),
    reservationCards,
    groupedDays,
  };
}



function getDashboardDataCached_(resident) {
  const cache = CacheService.getScriptCache();
  const key = buildCacheKey_('dashboard', resident && resident.residentId ? resident.residentId : 'anonymous');
  const cached = cache ? cache.get(key) : '';

  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      parsed.serverNowIso = nowLocalIso_();
      return parsed;
    } catch (err) {}
  }

  const dashboard = getDashboardData_(resident);
  if (cache) {
    try {
      cache.put(key, JSON.stringify(dashboard), APP_CONFIG.DASHBOARD_CACHE_TTL_SECONDS);
    } catch (err) {}
  }
  return dashboard;
}

function resolveSlotAction_(slot, ownReservation, ownReservationCanChange, bookingOpen) {
  if (isSlotConcluded_(slot)) return 'past';

  if (ownReservation && ownReservation.slotId === slot.slotId) {
    return ownReservationCanChange ? 'mine' : 'locked';
  }

  if (!bookingOpen || !slot.isOpen) return 'closed';
  if (slot.availableCount <= 0) return 'full';
  if (!ownReservation) return 'claim';

  return ownReservationCanChange ? 'move' : 'used';
}

// ---------------------------
// Cycle and public board
// ---------------------------

function openCycleAt_(cycleStartDate, options) {
  options = options || {};

  const cycleStartText = formatDateYmd_(cycleStartDate);
  const openedAtLocalIso = normalizeDateTimeCellToIso_(options.openedAtLocalIso || nowLocalIso_());
  const displayStartText = options.keepTodayVisible ? openedAtLocalIso.slice(0, 10) : cycleStartText;
  const carryoverFinalizeAt = options.keepTodayVisible
    ? computeCarryoverFinalizeAt_(openedAtLocalIso)
    : '';

  const previousSettings = getSettingsMap_();
  const rollbackEntries = [
    { key: 'booking_open', value: toSheetCellValue_(previousSettings.booking_open), notes: 'Whether residents can claim or move slots' },
    { key: 'active_cycle_start', value: toSheetCellValue_(previousSettings.active_cycle_start), notes: 'The Monday that anchors the current 14-day cycle' },
    { key: 'public_display_start', value: toSheetCellValue_(previousSettings.public_display_start), notes: 'Optional earlier display start used during Sunday carryover' },
    { key: 'carryover_finalize_at', value: toSheetCellValue_(previousSettings.carryover_finalize_at), notes: 'Internal exact datetime used to end carryover display after a cycle opens' },
    { key: 'cycle_days', value: toSheetCellValue_(previousSettings.cycle_days), notes: 'Number of days in one cycle' },
    { key: 'slot_capacity', value: toSheetCellValue_(previousSettings.slot_capacity), notes: 'Reservations per slot' },
    { key: 'timezone', value: toSheetCellValue_(previousSettings.timezone), notes: 'Project timezone' },
  ];

  try {
    buildSlotsForCycle_(cycleStartDate);

    setSettingsBatch_([
      { key: 'booking_open', value: 'TRUE', notes: 'Whether residents can claim or move slots' },
      { key: 'active_cycle_start', value: cycleStartText, notes: 'The Monday that anchors the current 14-day cycle' },
      { key: 'public_display_start', value: displayStartText, notes: 'Optional earlier display start used during Sunday carryover' },
      { key: 'carryover_finalize_at', value: carryoverFinalizeAt, notes: 'Internal exact datetime used to end carryover display after a cycle opens' },
      { key: 'cycle_days', value: String(APP_CONFIG.CYCLE_DAYS), notes: 'Number of days in one cycle' },
      { key: 'slot_capacity', value: String(APP_CONFIG.SLOT_CAPACITY), notes: 'Reservations per slot' },
      { key: 'timezone', value: APP_CONFIG.TZ, notes: 'Project timezone' },
    ]);

    rebuildCycleAvailability();
    archiveOldDataNow(false);
    refreshPublicBoardFull_();
    clearQueuedPublicBoardRefresh_();
    removePublicBoardRefreshTrigger_();

    logAudit_(options.actor || 'system', options.action || 'open_cycle', cycleStartText, `Opened cycle ${cycleStartText}`);
  } catch (err) {
    try {
      setSettingsBatch_(rollbackEntries);
    } catch (rollbackErr) {}
    throw err;
  }
}

function buildSlotsForCycle_(cycleStartDate) {
  const info = getSlotsSheetInfo_();
  const { sheet, values, map } = info;
  const cycleStartText = formatDateYmd_(cycleStartDate);

  // Map existing slot_id -> row index in values
  const existingRowBySlotId = {};
  for (let i = 1; i < values.length; i++) {
    const existingSlotId = String(getCellValue_(values[i], map, 'slot_id') || '').trim();
    if (existingSlotId) existingRowBySlotId[existingSlotId] = i;
  }

  const rowsToAppend = [];
  let changedExisting = false;

  for (let dayOffset = 0; dayOffset < APP_CONFIG.CYCLE_DAYS; dayOffset++) {
    const currentDate = addDays_(cycleStartDate, dayOffset);
    const slotDateText = formatDateYmd_(currentDate);
    const dayLabel = Utilities.formatDate(currentDate, APP_CONFIG.TZ, 'EEE dd MMM');

    APP_CONFIG.SLOT_TEMPLATES.forEach(template => {
      const start = new Date(currentDate);
      start.setHours(template.startHour, 0, 0, 0);
      const end = new Date(start);
      end.setHours(end.getHours() + 4);

      const startText = Utilities.formatDate(start, APP_CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss");
      const endText = Utilities.formatDate(end, APP_CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss");
      const slotId = `SLOT-${startText.replace(/[^0-9]/g, '').slice(0, 12)}`;

      const rowData = [
        slotId,
        weekKeyForDate_(start),
        slotDateText,
        dayLabel,
        template.label,
        startText,
        endText,
        APP_CONFIG.SLOT_CAPACITY,
        true,
        cycleStartText,
        new Date(),
      ];

      const existingRowIndex = existingRowBySlotId[slotId];
      if (existingRowIndex !== undefined) {
        // Reuse overlapping slot rows instead of skipping them.
        // Keep created_at from the existing row.
        const existingRow = values[existingRowIndex];
        setCellValue_(existingRow, map, 'slot_id', rowData[0]);
        setCellValue_(existingRow, map, 'week_key', rowData[1]);
        setCellValue_(existingRow, map, 'slot_date', rowData[2]);
        setCellValue_(existingRow, map, 'day_label', rowData[3]);
        setCellValue_(existingRow, map, 'slot_label', rowData[4]);
        setCellValue_(existingRow, map, 'start_iso', rowData[5]);
        setCellValue_(existingRow, map, 'end_iso', rowData[6]);
        setCellValue_(existingRow, map, 'capacity', rowData[7]);
        setCellValue_(existingRow, map, 'is_open', rowData[8]);
        setCellValue_(existingRow, map, 'cycle_start_date', rowData[9]);
        // created_at intentionally preserved
        changedExisting = true;
      } else {
        rowsToAppend.push(rowData);
      }
    });
  }

  if (changedExisting) {
    sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  }

  if (rowsToAppend.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
  }

  if (changedExisting || rowsToAppend.length) {
    clearExecutionCaches_();
    bumpDataCacheVersion_();
  }
}

function getQueuedPublicBoardRefresh_() {
  const raw = PropertiesService.getScriptProperties().getProperty(APP_CONFIG.PUBLIC_BOARD_DIRTY_KEY);
  if (!raw) return null;

  try {
    const parsed = JSON.parse(raw);
    const changedDateKeys = Array.isArray(parsed.changedDateKeys)
      ? Array.from(new Set(parsed.changedDateKeys.map(normalizeDateCellToYmd_).filter(Boolean)))
      : [];
    return {
      full: Boolean(parsed.full),
      changedDateKeys,
      queuedAtIso: parsed.queuedAtIso || '',
    };
  } catch (err) {
    return { full: true, changedDateKeys: [], queuedAtIso: '' };
  }
}

function clearQueuedPublicBoardRefresh_() {
  PropertiesService.getScriptProperties().deleteProperty(APP_CONFIG.PUBLIC_BOARD_DIRTY_KEY);
}

function queuePublicBoardRefresh_(changedDateKeys, options) {
  options = options || {};
  const props = PropertiesService.getScriptProperties();
  const existing = getQueuedPublicBoardRefresh_() || { full: false, changedDateKeys: [] };
  const mergedDates = new Set(existing.changedDateKeys || []);
  (changedDateKeys || []).forEach(dateKey => {
    const normalized = normalizeDateCellToYmd_(dateKey);
    if (normalized) mergedDates.add(normalized);
  });

  const queued = {
    full: Boolean(options.full || existing.full),
    changedDateKeys: Array.from(mergedDates),
    queuedAtIso: new Date().toISOString(),
  };

  props.setProperty(APP_CONFIG.PUBLIC_BOARD_DIRTY_KEY, JSON.stringify(queued));

  try {
    installPublicBoardRefreshTrigger_();
  } catch (err) {
    if (queued.full) {
      syncReservationMetadata_();
      refreshPublicBoardFull_();
    } else {
      refreshPublicBoardIncremental_(queued.changedDateKeys);
    }
    clearQueuedPublicBoardRefresh_();
    removePublicBoardRefreshTrigger_();
  }
}

function flushPendingPublicBoardRefresh_() {
  const queued = getQueuedPublicBoardRefresh_();
  if (!queued) return false;
  if (!queued.full && !(queued.changedDateKeys || []).length) {
    clearQueuedPublicBoardRefresh_();
    return false;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return false;

  try {
    clearExecutionCaches_();
    const freshQueued = getQueuedPublicBoardRefresh_();
    if (!freshQueued) return false;
    if (!freshQueued.full && !(freshQueued.changedDateKeys || []).length) {
      clearQueuedPublicBoardRefresh_();
      return false;
    }

    if (freshQueued.full) {
      syncReservationMetadata_();
      refreshPublicBoardFull_();
    } else {
      refreshPublicBoardIncremental_(freshQueued.changedDateKeys);
    }

    clearQueuedPublicBoardRefresh_();
    return true;
  } finally {
    lock.releaseLock();
  }
}

function syncPublicBoardAfterReservationChange_(changedDateKeys) {
  const relevantDates = Array.from(new Set((changedDateKeys || []).map(normalizeDateCellToYmd_).filter(Boolean)));
  if (!relevantDates.length) return true;

  try {
    clearExecutionCaches_();
    const updatedInline = refreshPublicBoardIncremental_(relevantDates, { allowFullFallback: false });
    if (!updatedInline) {
      queuePublicBoardRefresh_(relevantDates, { full: true });
      return false;
    }
    return true;
  } catch (err) {
    queuePublicBoardRefresh_(relevantDates, { full: true });
    return false;
  }
}

function refreshPublicBoardFull_() {
  const sheet = getBackendSpreadsheet_().getSheetByName(APP_CONFIG.SHEETS.PUBLIC_BOARD);
  const window = getPublicBoardWindow_();
  const slotLookup = getPublicBoardSlotLookup_(window);
  const reservationsByKey = getPublicBoardReservationLookup_();
  const values = buildFullPublicBoardValues_(window, slotLookup, reservationsByKey);

  const totalRows = Math.max(sheet.getMaxRows(), values.length);
  const totalCols = Math.max(sheet.getMaxColumns(), values[0].length);
  if (totalRows > 0 && totalCols > 0) {
    sheet.getRange(1, 1, totalRows, totalCols).breakApart();
  }
  sheet.clearContents();
  sheet.clearFormats();
  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  applyPublicBoardFormatting_(sheet, values.length, values[0].length, window);
  PropertiesService.getScriptProperties().setProperty(APP_CONFIG.PUBLIC_BOARD_SIGNATURE_KEY, getPublicBoardSignature_(window));
}

function refreshPublicBoardIncremental_(changedDateKeys, options) {
  options = options || {};
  const relevantDates = Array.from(new Set((changedDateKeys || []).map(normalizeDateCellToYmd_).filter(Boolean)));
  if (!relevantDates.length) return false;

  const sheet = getBackendSpreadsheet_().getSheetByName(APP_CONFIG.SHEETS.PUBLIC_BOARD);
  const window = getPublicBoardWindow_();
  const signature = getPublicBoardSignature_(window);
  const storedSignature = PropertiesService.getScriptProperties().getProperty(APP_CONFIG.PUBLIC_BOARD_SIGNATURE_KEY);

  // Fall back to a full rebuild only if the board structure changed.
  if (storedSignature !== signature || sheet.getLastRow() < 2 + window.dates.length * APP_CONFIG.SLOT_CAPACITY) {
    if (options.allowFullFallback === false) return false;
    refreshPublicBoardFull_();
    return true;
  }

  // IMPORTANT: On claim/move/cancel we only change reservation labels inside the timeslot cells.
  // We do NOT touch merges, borders, row heights, headers, or machine/day/date columns here.
  const slotLookup = getPublicBoardSlotLookupForDates_(relevantDates);
  const relevantSlotIds = Object.values(slotLookup).map(slot => slot.slotId);
  const reservationsByKey = getPublicBoardReservationLookupForSlotIds_(relevantSlotIds);

  relevantDates.forEach(dateKey => {
    const dateIndex = window.dates.indexOf(dateKey);
    if (dateIndex === -1) return;

    const rowStart = 3 + dateIndex * APP_CONFIG.SLOT_CAPACITY;
    const contentValues = buildPublicBoardDateBlockContentValues_(dateKey, window, slotLookup, reservationsByKey);

    // Columns D:I only (6 slot columns). Leave A:C formatting completely untouched.
    sheet.getRange(rowStart, 4, contentValues.length, contentValues[0].length).setValues(contentValues);
  });

  return true;
}

function buildFullPublicBoardValues_(window, slotLookup, reservationsByKey) {
  const settings = getSettingsMap_();
  const dormName = String(settings.dorm_name || 'Dorm 14').trim() || 'Dorm 14';
  const cycleRange = window.cycleStart && window.cycleEnd
    ? `${formatPublicLongDateEn_(window.cycleStart)} – ${formatPublicLongDateEn_(window.cycleEnd)}`
    : '—';

  const values = [];
  values.push([
    'Pralni seznam | Laundry Schedule', '', '',
    'Cycle / Obdobje:',
    cycleRange,
    '', '', 'Dom | Dorm:', dormName,
  ]);
  values.push(APP_CONFIG.HEADERS.PUBLIC_BOARD.slice());

  window.dates.forEach(dateKey => {
    const rows = buildPublicBoardDateBlockValues_(dateKey, window, slotLookup, reservationsByKey);
    rows.forEach(row => values.push(row));
  });

  return values;
}

function buildPublicBoardDateBlockValues_(dateKey, window, slotLookup, reservationsByKey) {
  const dateObj = parseYmdDate_(dateKey);
  const dayName = Utilities.formatDate(dateObj, APP_CONFIG.TZ, 'EEEE');
  const dateLabel = `${dateObj.getDate()}. ${dateObj.getMonth() + 1}. ${dateObj.getFullYear()}`;

  slotLookup = slotLookup || getPublicBoardSlotLookup_(window);
  reservationsByKey = reservationsByKey || getPublicBoardReservationLookup_();

  const rows = [];
  for (let machineNo = 1; machineNo <= APP_CONFIG.SLOT_CAPACITY; machineNo++) {
    const row = [
      machineNo === 1 ? dayName : '',
      machineNo === 1 ? dateLabel : '',
      machineNo,
    ];

    APP_CONFIG.SLOT_TEMPLATES.forEach(template => {
      const slot = slotLookup[`${dateKey}::${template.label}`];
      const label = slot ? (reservationsByKey[`${slot.slotId}::${machineNo}`] || '') : '';
      row.push(label);
    });

    rows.push(row);
  }

  return rows;
}


function buildPublicBoardDateBlockContentValues_(dateKey, window, slotLookup, reservationsByKey) {
  slotLookup = slotLookup || getPublicBoardSlotLookup_(window);
  reservationsByKey = reservationsByKey || getPublicBoardReservationLookup_();

  const rows = [];
  for (let machineNo = 1; machineNo <= APP_CONFIG.SLOT_CAPACITY; machineNo++) {
    const row = [];
    APP_CONFIG.SLOT_TEMPLATES.forEach(template => {
      const slot = slotLookup[`${dateKey}::${template.label}`];
      const label = slot ? (reservationsByKey[`${slot.slotId}::${machineNo}`] || '') : '';
      row.push(label);
    });
    rows.push(row);
  }
  return rows;
}

function getPublicBoardWindow_() {
  const settings = getSettingsMap_();
  const cycleStart = normalizeDateCellToYmd_(settings.active_cycle_start);
  const cycleDays = Number(settings.cycle_days || APP_CONFIG.CYCLE_DAYS) || APP_CONFIG.CYCLE_DAYS;
  const cycleEnd = cycleStart ? formatDateYmd_(addDays_(parseYmdDate_(cycleStart), cycleDays - 1)) : '';
  const displayStart = normalizeDateCellToYmd_(settings.public_display_start) || cycleStart;

  const dates = [];
  if (displayStart && cycleEnd) {
    let cursor = parseYmdDate_(displayStart);
    const end = parseYmdDate_(cycleEnd);
    while (cursor.getTime() <= end.getTime()) {
      dates.push(formatDateYmd_(cursor));
      cursor = addDays_(cursor, 1);
    }
  }

  return {
    displayStart,
    cycleStart,
    cycleEnd,
    dates,
  };
}

function getPublicBoardSignature_(window) {
  return `${window.displayStart}|${window.cycleStart}|${window.cycleEnd}|${window.dates.length}`;
}

function getPublicBoardSlotLookup_(window) {
  const lookup = {};
  const allowedDates = new Set(window.dates || []);
  getAllSlots_()
    .filter(slot => allowedDates.has(normalizeDateCellToYmd_(slot.slotDate)))
    .forEach(slot => {
      lookup[`${normalizeDateCellToYmd_(slot.slotDate)}::${slot.slotLabel}`] = slot;
    });
  return lookup;
}

function getPublicBoardReservationLookup_() {
  const lookup = {};
  getAllReservations_()
    .filter(r => r.status === 'confirmed')
    .forEach(r => {
      lookup[`${r.slotId}::${Number(r.machineNo || 0)}`] = r.residentLabel || formatResidentLabel_(r.displayName, r.room);
    });
  return lookup;
}

function applyPublicBoardFormatting_(sheet, totalRows, totalCols, window) {
  sheet.setFrozenRows(2);
  sheet.getRange(1, 1, totalRows, totalCols).setBorder(true, true, true, true, true, true);
  sheet.getRange(1, 1, 1, totalCols).setFontWeight('bold').setBackground('#f1f5fb');
  sheet.getRange(2, 1, 1, totalCols).setFontWeight('bold').setBackground('#dce6f7').setHorizontalAlignment('center');

  if (totalRows > 2) {
    const bodyRows = totalRows - 2;
    sheet.getRange(3, 3, bodyRows, 1).setBackground('#d7e8ff').setHorizontalAlignment('center');
    sheet.getRange(3, 4, bodyRows, 6).setWrap(true);
    sheet.getRange(3, 1, bodyRows, totalCols).setVerticalAlignment('middle');
  }

  for (let i = 0; i < window.dates.length; i++) {
    const rowStart = 3 + i * APP_CONFIG.SLOT_CAPACITY;
    applyPublicBoardBlockFormatting_(sheet, rowStart, APP_CONFIG.SLOT_CAPACITY, totalCols);
  }

  sheet.setColumnWidths(1, 1, 110);
  sheet.setColumnWidths(2, 1, 110);
  sheet.setColumnWidths(3, 1, 90);
  sheet.setColumnWidths(4, 6, 155);
}

function applyPublicBoardBlockFormatting_(sheet, rowStart, rowCount, totalCols) {
  const lastRow = rowStart + rowCount - 1;

  sheet.getRange(rowStart, 1, rowCount, 1).setFontStyle('italic').setHorizontalAlignment('center');
  sheet.getRange(rowStart, 2, rowCount, 1).setFontWeight('bold').setHorizontalAlignment('center');

  // Rebuild merged day/date cells after any incremental rewrite.
  if (rowCount > 1) {
    sheet.getRange(rowStart, 1, rowCount, 1).breakApart();
    sheet.getRange(rowStart, 2, rowCount, 1).breakApart();
    sheet.getRange(rowStart, 1, rowCount, 1).mergeVertically();
    sheet.getRange(rowStart, 2, rowCount, 1).mergeVertically();
  }

  for (let r = rowStart; r < rowStart + rowCount; r++) {
    sheet.setRowHeight(r, 28);
  }

  // Keep normal thin borders everywhere inside the block.
  sheet.getRange(rowStart, 1, rowCount, totalCols)
    .setBorder(null, null, null, null, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);

  // Emphasize only the bottom divider line for each day block, including merged day/date cells.
  sheet.getRange(rowStart, 1, rowCount, 1)
    .setBorder(null, null, true, null, null, null, '#4f4f4f', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange(rowStart, 2, rowCount, 1)
    .setBorder(null, null, true, null, null, null, '#4f4f4f', SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange(lastRow, 3, 1, totalCols - 2)
    .setBorder(null, null, true, null, null, null, '#4f4f4f', SpreadsheetApp.BorderStyle.SOLID_THICK);
}

// ---------------------------
// Sheets and schema
// ---------------------------

function ensureSchema_(force) {
  const props = PropertiesService.getScriptProperties();
  const version = props.getProperty(APP_CONFIG.SCHEMA_VERSION_KEY);
  if (!force && (_SCHEMA_READY_CACHE || version === APP_CONFIG.SCHEMA_VERSION)) {
    _SCHEMA_READY_CACHE = true;
    return;
  }

  const ss = getBackendSpreadsheet_();
  ensureSheet_(ss, APP_CONFIG.SHEETS.SETTINGS, APP_CONFIG.HEADERS.SETTINGS);
  ensureSheet_(ss, APP_CONFIG.SHEETS.RESIDENTS, APP_CONFIG.HEADERS.RESIDENTS);
  ensureSheet_(ss, APP_CONFIG.SHEETS.SLOTS, APP_CONFIG.HEADERS.SLOTS);
  ensureSheet_(ss, APP_CONFIG.SHEETS.RESERVATIONS, APP_CONFIG.HEADERS.RESERVATIONS);
  ensureSheet_(ss, APP_CONFIG.SHEETS.RESERVATIONS_ARCHIVE, APP_CONFIG.HEADERS.RESERVATIONS);
  ensureSheet_(ss, APP_CONFIG.SHEETS.PUBLIC_BOARD, APP_CONFIG.HEADERS.PUBLIC_BOARD);
  ensureSheet_(ss, APP_CONFIG.SHEETS.AUDIT, APP_CONFIG.HEADERS.AUDIT);
  ensureSheet_(ss, APP_CONFIG.SHEETS.AUDIT_ARCHIVE, APP_CONFIG.HEADERS.AUDIT);
  ensureSheet_(ss, APP_CONFIG.SHEETS.MACHINES, APP_CONFIG.HEADERS.MACHINES);
  ensureSheet_(ss, APP_CONFIG.SHEETS.MACHINE_REPORTS, APP_CONFIG.HEADERS.MACHINE_REPORTS);

  writeDefaultSettings_(ss);
  writeDefaultMachines_(ss);
  freezeHeaderRows_(ss);

  props.setProperty(APP_CONFIG.SCHEMA_VERSION_KEY, APP_CONFIG.SCHEMA_VERSION);
  _SCHEMA_READY_CACHE = true;
}

function ensureCycleOpenedIfDue_() {
  const initialSettings = getSettingsMap_();
  if (!isCycleSwitchDueForSettings_(initialSettings, nowLocalIso_())) return false;

  let lock;
  try {
    lock = acquireScriptLockOrThrow_(
      APP_CONFIG.CYCLE_SWITCH_LOCK_TIMEOUT_MS,
      'The new cycle is opening right now. Please retry in a moment.'
    );
  } catch (err) {
    if (waitForCycleSwitchToSettle_()) return false;
    throw err;
  }

  try {
    clearExecutionCaches_();
    const settings = getSettingsMap_();
    const activeStartText = normalizeDateCellToYmd_(settings.active_cycle_start);
    const activeStart = parseYmdDate_(activeStartText);
    if (!activeStartText || isNaN(activeStart.getTime())) return false;

    const freshNowIso = nowLocalIso_();

    const freshSimulationOpenAt = getSimulationOpenIso_();
    if (freshSimulationOpenAt && freshSimulationOpenAt <= freshNowIso) {
      const nextCycleStart = addDays_(activeStart, APP_CONFIG.CYCLE_DAYS);
      openCycleAt_(nextCycleStart, {
        keepTodayVisible: true,
        actor: 'automation',
        action: 'open_cycle_on_request_simulation',
        openedAtLocalIso: freshSimulationOpenAt,
      });
      setSetting_('simulation_open_at', '', 'Optional exact datetime override for test copies, format YYYY-MM-DDTHH:mm:ss');
      return true;
    }

    const freshOpeningTime = normalizeHhMm_(settings.opening_time) || '19:00';
    const freshCycleEndYmd = formatDateYmd_(addDays_(activeStart, APP_CONFIG.CYCLE_DAYS - 1));
    const scheduledOpenIso = freshCycleEndYmd ? `${freshCycleEndYmd}T${freshOpeningTime}:00` : '';
    if (scheduledOpenIso && scheduledOpenIso <= freshNowIso) {
      const nextCycleStart = addDays_(activeStart, APP_CONFIG.CYCLE_DAYS);
      openCycleAt_(nextCycleStart, {
        keepTodayVisible: true,
        actor: 'automation',
        action: 'open_cycle_on_request',
        openedAtLocalIso: freshNowIso,
      });
      return true;
    }

    return false;
  } finally {
    lock.releaseLock();
  }
}

function getBackendSpreadsheet_() {
  if (_SS_CACHE) return _SS_CACHE;
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty(APP_CONFIG.SPREADSHEET_ID_KEY);
  const active = SpreadsheetApp.getActiveSpreadsheet();
  const id = storedId || (active ? active.getId() : '');
  if (!id) throw new Error('Spreadsheet ID is not configured. Run initializePrototype from the bound spreadsheet first.');
  _SS_CACHE = SpreadsheetApp.openById(id);
  if (!storedId) props.setProperty(APP_CONFIG.SPREADSHEET_ID_KEY, id);
  return _SS_CACHE;
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    clearExecutionCaches_({ resetStructure: true });
    return sheet;
  }

  const headerRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), headers.length)).getValues()[0].map(v => String(v || '').trim());
  let changed = false;
  headers.forEach((header, idx) => {
    if (headerRow[idx] === header) return;
    const existingIdx = headerRow.indexOf(header);
    if (existingIdx === -1) {
      sheet.getRange(1, headerRow.length + 1).setValue(header);
      headerRow.push(header);
      changed = true;
    }
  });

  if (changed)   clearExecutionCaches_({ resetStructure: true });

  return sheet;
}

function freezeHeaderRows_(ss) {
  [
    APP_CONFIG.SHEETS.SETTINGS,
    APP_CONFIG.SHEETS.RESIDENTS,
    APP_CONFIG.SHEETS.SLOTS,
    APP_CONFIG.SHEETS.RESERVATIONS,
    APP_CONFIG.SHEETS.RESERVATIONS_ARCHIVE,
    APP_CONFIG.SHEETS.AUDIT,
    APP_CONFIG.SHEETS.AUDIT_ARCHIVE,
    APP_CONFIG.SHEETS.MACHINES,
    APP_CONFIG.SHEETS.MACHINE_REPORTS
  ].forEach(name => {
    const sheet = ss.getSheetByName(name);
    if (sheet) sheet.setFrozenRows(1);
  });
}

function writeDefaultSettings_(ss) {
  const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.SETTINGS);
  const defaults = [
    ['booking_open', 'FALSE', 'Whether residents can claim or move slots'],
    ['active_cycle_start', '', 'Start Monday of the active 14-day cycle'],
    ['public_display_start', '', 'Optional earlier display start used during Sunday carryover'],
    ['cycle_days', String(APP_CONFIG.CYCLE_DAYS), 'Number of days in one cycle'],
    ['slot_capacity', String(APP_CONFIG.SLOT_CAPACITY), 'Reservations per slot'],
    ['timezone', APP_CONFIG.TZ, 'Project timezone'],
    ['self_signup_enabled', 'TRUE', 'Whether residents can create their own code'],
    ['self_reset_enabled', 'TRUE', 'Whether residents can reset their own code with a recovery key'],
    ['max_active_residents_per_room', '2', 'Block signup or room changes once this many non-stale active residents exist in the same room'],
    ['resident_stale_after_days', '14', 'After this many days without activity, an active resident stops counting toward room capacity if they have no relevant reservation'],
    ['resident_disable_after_days', '120', 'After this many days without activity, an active resident may be auto-disabled if they have no relevant reservation'],
    ['disabled_account_delete_after_days', '180', 'After this many days in disabled state, the resident row may be permanently deleted'],
    ['activity_touch_cooldown_hours', '6', 'Throttle passive last_activity_at writes for login and refresh to at most once per resident in this many hours'],
    ['audit_enabled', 'FALSE', 'Whether Audit receives new rows'],
    ['dorm_name', 'Dorm 14', 'Shown on the PublicBoard header'],
    ['archive_keep_past_days', '2', 'For non-active-cycle confirmed reservations, keep them in the live sheet this many days after they end'],
    ['audit_keep_rows', '1000', 'Keep only the newest N audit rows in the live Audit sheet'],
    ['opening_time', '19:00', 'Normal weekly opening threshold time in HH:mm'],
    ['simulation_open_at', '', 'Optional exact datetime override for test copies, format YYYY-MM-DDTHH:mm:ss'],
    ['carryover_finalize_at', '', 'Internal exact datetime used to end carryover display after a cycle opens'],
    ['opening_policy', 'Every 2 weeks on rollover Sunday at 19:00, next cycle opens for the upcoming Monday', 'Notes only'],
  ];

  const existing = sheet.getDataRange().getValues();
  const keys = new Set(existing.slice(1).map(row => String(row[0] || '').trim()));
  const rowsToAdd = defaults.filter(row => !keys.has(row[0]));
  if (rowsToAdd.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    clearExecutionCaches_();
  }
}

function writeDefaultMachines_(ss) {
  const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.MACHINES);
  const values = sheet.getDataRange().getValues();

  const existingMachineNos = new Set(
    values
      .slice(1)
      .map(row => Number(row[0] || 0))
      .filter(machineNo => machineNo > 0)
  );

  const now = new Date();
  const rowsToAdd = [];

  for (
    let machineNo = 1;
    machineNo <= APP_CONFIG.SLOT_CAPACITY;
    machineNo++
  ) {
    if (existingMachineNos.has(machineNo)) continue;

    rowsToAdd.push([
      machineNo,
      'WORKING',
      now,
      ''
    ]);
  }

  if (!rowsToAdd.length) return;

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      rowsToAdd.length,
      APP_CONFIG.HEADERS.MACHINES.length
    )
    .setValues(rowsToAdd);

  clearExecutionCaches_({ resetStructure: true });
}

function getSettingsMap_() {
  if (_SETTINGS_CACHE) return _SETTINGS_CACHE;
  const sheet = getBackendSpreadsheet_().getSheetByName(APP_CONFIG.SHEETS.SETTINGS);
  const values = sheet.getDataRange().getValues();
  const out = {};
  for (let i = 1; i < values.length; i++) {
    out[String(values[i][0] || '').trim()] = values[i][1];
  }
  _SETTINGS_CACHE = out;
  return out;
}

function setSettingsBatch_(entries) {
  const sheet = getBackendSpreadsheet_().getSheetByName(APP_CONFIG.SHEETS.SETTINGS);
  const rowCount = Math.max(sheet.getLastRow(), 1);
  const values = sheet.getRange(1, 1, rowCount, 3).getValues();

  const keyToRow = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) keyToRow[key] = i;
  }

  let changed = false;

  (entries || []).forEach(entry => {
    const key = String((entry && entry.key) || '').trim();
    if (!key) return;

    const value = entry && entry.value !== undefined && entry.value !== null ? entry.value : '';
    const notes = entry && entry.notes !== undefined ? entry.notes : undefined;

    if (keyToRow[key] !== undefined) {
      const rowIndex = keyToRow[key];
      const prevValue = values[rowIndex][1];
      const prevNotes = values[rowIndex][2];

      if (prevValue !== value) {
        values[rowIndex][1] = value;
        changed = true;
      }

      if (notes !== undefined && prevNotes !== notes) {
        values[rowIndex][2] = notes;
        changed = true;
      }
    } else {
      values.push([key, value, notes !== undefined ? notes : '']);
      changed = true;
    }
  });

  if (!changed) return;

  const normalized = values.map(row => [
    row[0] === undefined || row[0] === null ? '' : row[0],
    row[1] === undefined || row[1] === null ? '' : row[1],
    row[2] === undefined || row[2] === null ? '' : row[2],
  ]);

  sheet.getRange(1, 1, normalized.length, 3).setValues(normalized);
  clearExecutionCaches_();
  bumpDataCacheVersion_();
}

function setSetting_(key, value, notes) {
  setSettingsBatch_([{ key, value, notes }]);
}

function clearExecutionCaches_(options) {
  options = options || {};
  const resetStructure = Boolean(options.resetStructure);

  _SETTINGS_CACHE = null;
  _RESIDENTS_CACHE = null;
  _SLOTS_CACHE = null;
  _RESERVATIONS_CACHE = null;
  _SLOT_BY_ID_CACHE = null;
  _RESERVATION_BY_ID_CACHE = null;
  _CONFIRMED_RESERVATION_COUNTS_BY_SLOT_CACHE = null;
  _CONFIRMED_RESERVATIONS_BY_RESIDENT_CACHE = null;
  _DATA_CACHE_VERSION_CACHE = null;
  _CONFIRMED_RESERVATIONS_BY_SLOT_CACHE = null;
  _MACHINES_CACHE = null;
  _MACHINE_REPORTS_CACHE = null;

  if (resetStructure) {
    _HEADER_MAP_CACHE = {};
    _SS_CACHE = null;
  }
}

function getUiOrNull_() {
  try {
    return SpreadsheetApp.getUi();
  } catch (err) {
    return null;
  }
}

function notifyAdmin_(message) {
  const ui = getUiOrNull_();
  if (ui) {
    ui.alert(message);
    return;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.toast(message, APP_CONFIG.APP_NAME, 6);
  } catch (err) {}
  Logger.log(message);
}


function getDataCacheVersion_() {
  if (_DATA_CACHE_VERSION_CACHE !== null) return _DATA_CACHE_VERSION_CACHE;
  const props = PropertiesService.getScriptProperties();
  _DATA_CACHE_VERSION_CACHE = props.getProperty(APP_CONFIG.DATA_CACHE_VERSION_KEY) || '0';
  return _DATA_CACHE_VERSION_CACHE;
}

function bumpDataCacheVersion_() {
  const props = PropertiesService.getScriptProperties();
  const next = String(Number(props.getProperty(APP_CONFIG.DATA_CACHE_VERSION_KEY) || '0') + 1);
  props.setProperty(APP_CONFIG.DATA_CACHE_VERSION_KEY, next);
  _DATA_CACHE_VERSION_CACHE = next;
  return next;
}

function buildCacheKey_(prefix, suffix) {
  return `${prefix}:${getDataCacheVersion_()}:${suffix}`;
}

function getCachedJson_(prefix, suffix, ttlSeconds, builder) {
  const cache = CacheService.getScriptCache();
  const key = buildCacheKey_(prefix, suffix);
  if (cache) {
    const cached = cache.get(key);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (err) {}
    }
  }

  const fresh = builder();
  if (cache) {
    try {
      cache.put(key, JSON.stringify(fresh), ttlSeconds);
    } catch (err) {}
  }
  return fresh;
}

// ---------------------------
// Data access
// ---------------------------

function getHeaderMap_(sheet) {
  const key = sheet.getSheetName();
  if (_HEADER_MAP_CACHE[key]) return _HEADER_MAP_CACHE[key];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(v => String(v || '').trim());
  const map = {};
  headers.forEach((header, idx) => { map[header] = idx; });
  _HEADER_MAP_CACHE[key] = map;
  return map;
}

function getResidentsSheetInfo_() {
  const sheet = getBackendSpreadsheet_().getSheetByName(APP_CONFIG.SHEETS.RESIDENTS);
  return { sheet, values: sheet.getDataRange().getValues(), map: getHeaderMap_(sheet) };
}

function getSlotsSheetInfo_() {
  const sheet = getBackendSpreadsheet_().getSheetByName(APP_CONFIG.SHEETS.SLOTS);
  return { sheet, values: sheet.getDataRange().getValues(), map: getHeaderMap_(sheet) };
}

function getReservationsSheetInfo_() {
  const sheet = getBackendSpreadsheet_().getSheetByName(APP_CONFIG.SHEETS.RESERVATIONS);
  return { sheet, values: sheet.getDataRange().getValues(), map: getHeaderMap_(sheet) };
}

function getAuditSheetInfo_() {
  const sheet = getBackendSpreadsheet_().getSheetByName(APP_CONFIG.SHEETS.AUDIT);
  return { sheet, values: sheet.getDataRange().getValues(), map: getHeaderMap_(sheet) };
}

function getCellValue_(row, map, key) {
  const idx = map[key];
  return idx === undefined ? '' : row[idx];
}

function setCellValue_(row, map, key, value) {
  const idx = map[key];
  if (idx === undefined) return;
  row[idx] = value;
}

function getAllResidents_() {
  if (_RESIDENTS_CACHE) return _RESIDENTS_CACHE;
  _RESIDENTS_CACHE = getCachedJson_('residents', 'all', APP_CONFIG.DATA_CACHE_TTL_SECONDS, () => {
    const { values, map } = getResidentsSheetInfo_();
    if (values.length < 2) return [];
    return values.slice(1).map(row => buildResidentObjectFromRow_(row, map));
  });
  return _RESIDENTS_CACHE;
}

function getAllSlots_() {
  if (_SLOTS_CACHE) return _SLOTS_CACHE;
  _SLOTS_CACHE = getCachedJson_('slots', 'all', APP_CONFIG.DATA_CACHE_TTL_SECONDS, () => {
    const { values, map } = getSlotsSheetInfo_();
    if (values.length < 2) return [];
    return values.slice(1)
      .filter(row => String(getCellValue_(row, map, 'slot_id') || '').trim())
      .map(row => ({
        slotId: String(getCellValue_(row, map, 'slot_id') || '').trim(),
        weekKey: normalizeDateCellToYmd_(getCellValue_(row, map, 'week_key')),
        slotDate: normalizeDateCellToYmd_(getCellValue_(row, map, 'slot_date')),
        dayLabel: String(getCellValue_(row, map, 'day_label') || '').trim(),
        slotLabel: String(getCellValue_(row, map, 'slot_label') || '').trim(),
        startIso: normalizeDateTimeCellToIso_(getCellValue_(row, map, 'start_iso')),
        endIso: normalizeDateTimeCellToIso_(getCellValue_(row, map, 'end_iso')),
        capacity: Number(getCellValue_(row, map, 'capacity') || APP_CONFIG.SLOT_CAPACITY),
        isOpen: normalizeBoolean_(getCellValue_(row, map, 'is_open')),
        cycleStartDate: normalizeDateCellToYmd_(getCellValue_(row, map, 'cycle_start_date')),
        createdAt: getCellValue_(row, map, 'created_at'),
      }));
  });
  return _SLOTS_CACHE;
}

function getMachines_() {
  if (_MACHINES_CACHE) return _MACHINES_CACHE;

  const { sheet, map } = getSheetRef_(APP_CONFIG.SHEETS.MACHINES);
  const values = sheet.getDataRange().getValues();

  _MACHINES_CACHE = values
    .slice(1)
    .map(row => ({
      machineNo: Number(getCellValue_(row, map, 'machine_no') || 0),
      status: String(
        getCellValue_(row, map, 'status') || ''
      ).trim().toUpperCase(),
      statusChangedAt: getCellValue_(row, map, 'status_changed_at'),
      statusNote: String(
        getCellValue_(row, map, 'status_note') || ''
      ).trim(),
    }))
    .filter(machine => machine.machineNo > 0)
    .sort((a, b) => a.machineNo - b.machineNo);

  return _MACHINES_CACHE;
}

function getWorkingMachineNumbers_() {
  return getMachines_()
    .filter(machine => machine.status === 'WORKING')
    .map(machine => machine.machineNo);
}

function getMachineStatus_(machineNo) {
  const machine = getMachines_()
    .find(item => item.machineNo === Number(machineNo));

  return machine ? machine.status : 'OUT_OF_ORDER';
}

function isMachineWorking_(machineNo) {
  return getMachineStatus_(machineNo) === 'WORKING';
}

function getMachineReports_() {
  if (_MACHINE_REPORTS_CACHE) return _MACHINE_REPORTS_CACHE;

  const { sheet, map } = getSheetRef_(APP_CONFIG.SHEETS.MACHINE_REPORTS);
  const values = sheet.getDataRange().getValues();

  _MACHINE_REPORTS_CACHE = values
    .slice(1)
    .filter(row => String(getCellValue_(row, map, 'report_id') || '').trim())
    .map(row => ({
      reportId: String(getCellValue_(row, map, 'report_id') || '').trim(),
      machineNo: Number(getCellValue_(row, map, 'machine_no') || 0),
      residentId: String(getCellValue_(row, map, 'resident_id') || '').trim(),
      room: normalizeRoom_(getCellValue_(row, map, 'room')),
      displayName: normalizeDisplayName_(getCellValue_(row, map, 'display_name')),
      reservationId: String(getCellValue_(row, map, 'reservation_id') || '').trim(),
      slotId: String(getCellValue_(row, map, 'slot_id') || '').trim(),
      reportedAt: getCellValue_(row, map, 'reported_at'),
      description: String(getCellValue_(row, map, 'description') || '').trim(),
      status: String(getCellValue_(row, map, 'status') || '').trim().toUpperCase(),
      resolvedAt: getCellValue_(row, map, 'resolved_at'),
      resolvedByResidentId: String(getCellValue_(row, map, 'resolved_by_resident_id') || '').trim(),
      resolutionNote: String(getCellValue_(row, map, 'resolution_note') || '').trim(),
    }));

  return _MACHINE_REPORTS_CACHE;
}

function findMachineReportById_(reportId) {
  reportId = String(reportId || '').trim();
  return getMachineReports_().find(report => report.reportId === reportId) || null;
}

function getOpenMachineReports_() {
  return getMachineReports_().filter(
    report => report.status === APP_CONFIG.MACHINE_REPORT_STATUSES.OPEN
  );
}

function getOpenMachineReportForReservationMachine_(reservationId, machineNo) {
  reservationId = String(reservationId || '').trim();
  machineNo = Number(machineNo || 0);

  return getOpenMachineReports_().find(report =>
    report.reservationId === reservationId &&
    report.machineNo === machineNo
  ) || null;
}

function getOpenMachineReportCountsByMachine_() {
  const counts = {};

  getOpenMachineReports_().forEach(report => {
    const machineNo = Number(report.machineNo || 0);
    if (!machineNo) return;
    counts[machineNo] = (counts[machineNo] || 0) + 1;
  });

  return counts;
}

function appendMachineReport_(resident, machineNo, description, context) {
  context = context || {};

  const { sheet, map } = getSheetRef_(APP_CONFIG.SHEETS.MACHINE_REPORTS);
  const reportId = `MREP-${Utilities.getUuid().split('-')[0].toUpperCase()}`;
  const row = new Array(APP_CONFIG.HEADERS.MACHINE_REPORTS.length).fill('');

  setCellValue_(row, map, 'report_id', reportId);
  setCellValue_(row, map, 'machine_no', Number(machineNo || 0));
  setCellValue_(row, map, 'resident_id', resident.residentId);
  setCellValue_(row, map, 'room', resident.room);
  setCellValue_(row, map, 'display_name', resident.displayName);
  setCellValue_(row, map, 'reservation_id', String(context.reservationId || '').trim());
  setCellValue_(row, map, 'slot_id', String(context.slotId || '').trim());
  setCellValue_(row, map, 'reported_at', new Date());
  setCellValue_(row, map, 'description', description);
  setCellValue_(row, map, 'status', APP_CONFIG.MACHINE_REPORT_STATUSES.OPEN);

  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  clearExecutionCaches_();
  bumpDataCacheVersion_();

  return reportId;
}

function resolveMachineReport_(reportId, status, resolvedByResidentId, resolutionNote) {
  const allowedStatuses = [
    APP_CONFIG.MACHINE_REPORT_STATUSES.DISMISSED,
    APP_CONFIG.MACHINE_REPORT_STATUSES.MACHINE_OUT_OF_ORDER,
    APP_CONFIG.MACHINE_REPORT_STATUSES.MACHINE_RESTORED,
  ];

  if (!allowedStatuses.includes(status)) {
    throw new Error('Invalid machine report resolution status.');
  }

  const { sheet, map } = getSheetRef_(APP_CONFIG.SHEETS.MACHINE_REPORTS);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowReportId = String(getCellValue_(row, map, 'report_id') || '').trim();
    if (rowReportId !== String(reportId || '').trim()) continue;

    const currentStatus = String(getCellValue_(row, map, 'status') || '').trim().toUpperCase();
    if (currentStatus !== APP_CONFIG.MACHINE_REPORT_STATUSES.OPEN) {
      throw new Error('This machine report has already been resolved.');
    }

    setCellValue_(row, map, 'status', status);
    setCellValue_(row, map, 'resolved_at', new Date());
    setCellValue_(row, map, 'resolved_by_resident_id', String(resolvedByResidentId || '').trim());
    setCellValue_(row, map, 'resolution_note', String(resolutionNote || '').trim());

    sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
    clearExecutionCaches_();
    bumpDataCacheVersion_();
    return true;
  }

  throw new Error('Machine report was not found.');
}

function resolveOpenMachineReportsForMachine_(machineNo, resolvedByResidentId, resolutionNote, resolutionStatus) {
  machineNo = Number(machineNo || 0);
  resolutionStatus = String(
    resolutionStatus || APP_CONFIG.MACHINE_REPORT_STATUSES.MACHINE_OUT_OF_ORDER
  ).trim().toUpperCase();

  const allowedStatuses = [
    APP_CONFIG.MACHINE_REPORT_STATUSES.MACHINE_OUT_OF_ORDER,
    APP_CONFIG.MACHINE_REPORT_STATUSES.MACHINE_RESTORED,
  ];
  if (!allowedStatuses.includes(resolutionStatus)) {
    throw new Error('Invalid bulk machine report resolution status.');
  }

  const { sheet, map } = getSheetRef_(APP_CONFIG.SHEETS.MACHINE_REPORTS);
  const values = sheet.getDataRange().getValues();
  let changed = false;
  const now = new Date();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowMachineNo = Number(getCellValue_(row, map, 'machine_no') || 0);
    const status = String(getCellValue_(row, map, 'status') || '').trim().toUpperCase();

    if (rowMachineNo !== machineNo || status !== APP_CONFIG.MACHINE_REPORT_STATUSES.OPEN) {
      continue;
    }

    setCellValue_(row, map, 'status', resolutionStatus);
    setCellValue_(row, map, 'resolved_at', now);
    setCellValue_(row, map, 'resolved_by_resident_id', String(resolvedByResidentId || '').trim());
    setCellValue_(row, map, 'resolution_note', String(resolutionNote || '').trim());

    sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
    changed = true;
  }

  if (changed) {
    clearExecutionCaches_();
    bumpDataCacheVersion_();
  }

  return changed;
}

function getAllReservations_(forceFresh) {
  if (forceFresh) {
    // Critical reservation mutations must not reuse a snapshot loaded before
    // acquiring the lock (or an older version from CacheService).
    _RESERVATIONS_CACHE = null;
    _RESERVATION_BY_ID_CACHE = null;
    _CONFIRMED_RESERVATION_COUNTS_BY_SLOT_CACHE = null;
    _CONFIRMED_RESERVATIONS_BY_RESIDENT_CACHE = null;
    _CONFIRMED_RESERVATIONS_BY_SLOT_CACHE = null;
  }
  if (_RESERVATIONS_CACHE) return _RESERVATIONS_CACHE;
  const readReservations = () => {
    const { values, map } = getReservationsSheetInfo_();
    if (values.length < 2) return [];
    return values.slice(1)
      .filter(row => String(getCellValue_(row, map, 'reservation_id') || '').trim())
      .map(row => ({
        reservationId: String(getCellValue_(row, map, 'reservation_id') || '').trim(),
        residentId: String(getCellValue_(row, map, 'resident_id') || '').trim(),
        room: normalizeRoom_(getCellValue_(row, map, 'room')),
        displayName: normalizeDisplayName_(getCellValue_(row, map, 'display_name')),
        slotId: String(getCellValue_(row, map, 'slot_id') || '').trim(),
        weekKey: normalizeDateCellToYmd_(getCellValue_(row, map, 'week_key')),
        status: String(getCellValue_(row, map, 'status') || '').trim(),
        createdAt: getCellValue_(row, map, 'created_at'),
        updatedAt: getCellValue_(row, map, 'updated_at'),
        note: String(getCellValue_(row, map, 'note') || ''),
        machineNo: Number(getCellValue_(row, map, 'machine_no') || 0),
        residentLabel: String(getCellValue_(row, map, 'resident_label') || '').trim(),
      }));
  };
  _RESERVATIONS_CACHE = forceFresh
    ? readReservations()
    : getCachedJson_('reservations', 'all', APP_CONFIG.DATA_CACHE_TTL_SECONDS, readReservations);
  return _RESERVATIONS_CACHE;
}

function getSlotByIdMap_() {
  if (_SLOT_BY_ID_CACHE) return _SLOT_BY_ID_CACHE;
  _SLOT_BY_ID_CACHE = Object.fromEntries(getAllSlots_().map(slot => [slot.slotId, slot]));
  return _SLOT_BY_ID_CACHE;
}

function getReservationByIdMap_() {
  if (_RESERVATION_BY_ID_CACHE) return _RESERVATION_BY_ID_CACHE;
  _RESERVATION_BY_ID_CACHE = Object.fromEntries(getAllReservations_().map(reservation => [reservation.reservationId, reservation]));
  return _RESERVATION_BY_ID_CACHE;
}

function getConfirmedReservationCountsBySlot_() {
  if (_CONFIRMED_RESERVATION_COUNTS_BY_SLOT_CACHE) return _CONFIRMED_RESERVATION_COUNTS_BY_SLOT_CACHE;
  const counts = {};
  getAllReservations_().forEach(r => {
    if (r.status !== 'confirmed') return;
    counts[r.slotId] = (counts[r.slotId] || 0) + 1;
  });
  _CONFIRMED_RESERVATION_COUNTS_BY_SLOT_CACHE = counts;
  return _CONFIRMED_RESERVATION_COUNTS_BY_SLOT_CACHE;
}

function getConfirmedReservationsBySlotMap_() {
  if (_CONFIRMED_RESERVATIONS_BY_SLOT_CACHE) return _CONFIRMED_RESERVATIONS_BY_SLOT_CACHE;

  const bySlot = {};
  getAllReservations_().forEach(r => {
    if (r.status !== 'confirmed') return;
    if (!bySlot[r.slotId]) bySlot[r.slotId] = [];
    bySlot[r.slotId].push(r);
  });

  _CONFIRMED_RESERVATIONS_BY_SLOT_CACHE = bySlot;
  return _CONFIRMED_RESERVATIONS_BY_SLOT_CACHE;
}

function getConfirmedReservationsByResidentMap_() {
  if (_CONFIRMED_RESERVATIONS_BY_RESIDENT_CACHE) return _CONFIRMED_RESERVATIONS_BY_RESIDENT_CACHE;
  const byResident = {};
  getAllReservations_().forEach(r => {
    if (r.status !== 'confirmed') return;
    if (!byResident[r.residentId]) byResident[r.residentId] = [];
    byResident[r.residentId].push(r);
  });
  _CONFIRMED_RESERVATIONS_BY_RESIDENT_CACHE = byResident;
  return _CONFIRMED_RESERVATIONS_BY_RESIDENT_CACHE;
}

function getConfirmedReservationsForResident_(residentId) {
  return getConfirmedReservationsByResidentMap_()[residentId] || [];
}

function findReservationById_(reservationId) {
  return getReservationByIdMap_()[reservationId] || null;
}

function findSlotById_(slotId) {
  return getSlotByIdMap_()[slotId] || null;
}

function countConfirmedReservationsForSlot_(slotId) {
  return getConfirmedReservationCountsBySlot_()[slotId] || 0;
}

function getSlotChangeCutoff_(slot) {
  if (!slot || !slot.startIso) return new Date(NaN);
  const cutoff = valueToDateTime_(slot.startIso);
  if (isNaN(cutoff.getTime())) return new Date(NaN);
  cutoff.setHours(cutoff.getHours() + 1);
  return cutoff;
}

function canModifyReservationForSlot_(slot) {
  const cutoff = getSlotChangeCutoff_(slot);
  return !isNaN(cutoff.getTime()) && Date.now() < cutoff.getTime();
}

function isSlotChangeLocked_(slot) {
  const cutoff = getSlotChangeCutoff_(slot);
  return !isNaN(cutoff.getTime()) && Date.now() >= cutoff.getTime();
}

// ---------------------------
// Residents and auth
// ---------------------------

function requireResident_(room, accessCode) {
  const resident = findResidentByRoomAndCode_(normalizeRoom_(room), normalizeAccessCode_(accessCode));
  if (!resident) throw new Error('Session expired or code invalid.');
  return resident;
}

function countActiveResidentsInRoom_(room) {
  room = normalizeRoom_(room);
  return getAllResidents_().filter(r =>
    isResidentBlockingRoomCapacity_(r) &&
    normalizeComparable_(r.room) === normalizeComparable_(room)
  ).length;
}

function findResidentByRoomAndCode_(room, accessCode) {
  room = normalizeRoom_(room);
  accessCode = normalizeAccessCode_(accessCode);
  return getAllResidents_().find(r =>
    r.active &&
    normalizeComparable_(r.room) === normalizeComparable_(room) &&
    normalizeComparable_(r.accessCode) === normalizeComparable_(accessCode)
  ) || null;
}

function findResidentByRoomAndDisplayName_(room, displayName) {
  room = normalizeRoom_(room);
  displayName = normalizeDisplayName_(displayName);
  return getAllResidents_().find(r =>
    normalizeComparable_(r.room) === normalizeComparable_(room) &&
    normalizeComparable_(r.displayName) === normalizeComparable_(displayName)
  ) || null;
}

function findResidentByRecoveryKey_(room, displayName, recoveryKey) {
  const hash = hashRecoveryKey_(recoveryKey);
  return getAllResidents_().find(r =>
    r.active &&
    normalizeComparable_(r.room) === normalizeComparable_(room) &&
    normalizeComparable_(r.displayName) === normalizeComparable_(displayName) &&
    String(r.recoveryHash || '').trim() === hash
  ) || null;
}

function touchResidentLastLogin_(residentId) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `last_login_touch:${residentId}`;
  if (cache && cache.get(cacheKey)) return;

  const { sheet, values, map } = getResidentsSheetInfo_();
  for (let i = 1; i < values.length; i++) {
    if (String(getCellValue_(values[i], map, 'resident_id') || '').trim() !== residentId) continue;
    setCellValue_(values[i], map, 'last_login_at', new Date());
    sheet.getRange(i + 1, 1, 1, values[i].length).setValues([values[i]]);
    clearExecutionCaches_();
    if (cache) {
      try {
        cache.put(cacheKey, '1', APP_CONFIG.LAST_LOGIN_WRITE_THROTTLE_SECONDS);
      } catch (err) {}
    }
    return;
  }
}

function touchResidentActivity_(residentId, options) {
  options = options || {};
  const force = Boolean(options.force);
  const cooldownSeconds = force ? 0 : getActivityTouchCooldownSeconds_();

  const cache = CacheService.getScriptCache();
  const cacheKey = `last_activity_touch:${residentId}`;

  if (!force && cooldownSeconds > 0 && cache && cache.get(cacheKey)) {
    return false;
  }

  const { sheet, values, map } = getResidentsSheetInfo_();
  for (let i = 1; i < values.length; i++) {
    if (String(getCellValue_(values[i], map, 'resident_id') || '').trim() !== residentId) continue;

    setCellValue_(values[i], map, 'last_activity_at', new Date());
    sheet.getRange(i + 1, 1, 1, values[i].length).setValues([values[i]]);
    clearExecutionCaches_();

    if (!force && cooldownSeconds > 0 && cache) {
      try {
        cache.put(cacheKey, '1', cooldownSeconds);
      } catch (err) {}
    }
    return true;
  }

  return false;
}

function makeResidentRow_(data) {
  const now = new Date();
  return [
    data.residentId || generateResidentId_(),
    normalizeRoom_(data.room),
    normalizeDisplayName_(data.displayName),
    normalizeAccessCode_(data.accessCode),
    data.active !== false,
    data.notes || '',
    data.createdAt || now,
    data.updatedAt || now,
    data.recoveryHash || '',
    data.registeredAt || now,
    data.lastLoginAt || '',
    data.lastActivityAt || now,
    data.disabledAt || '',
    data.disabledReason || '',
  ];
}

function buildResidentObjectFromRow_(row, map) {
  return {
    residentId: String(getCellValue_(row, map, 'resident_id') || '').trim(),
    room: normalizeRoom_(getCellValue_(row, map, 'room')),
    displayName: normalizeDisplayName_(getCellValue_(row, map, 'display_name')),
    accessCode: normalizeAccessCode_(getCellValue_(row, map, 'access_code')),
    active: normalizeBoolean_(getCellValue_(row, map, 'active')),
    notes: String(getCellValue_(row, map, 'notes') || ''),
    createdAt: getCellValue_(row, map, 'created_at'),
    updatedAt: getCellValue_(row, map, 'updated_at'),
    recoveryHash: String(getCellValue_(row, map, 'recovery_hash') || '').trim(),
    registeredAt: getCellValue_(row, map, 'registered_at'),
    lastLoginAt: getCellValue_(row, map, 'last_login_at'),
    lastActivityAt: getCellValue_(row, map, 'last_activity_at'),
    disabledAt: getCellValue_(row, map, 'disabled_at'),
    disabledReason: String(getCellValue_(row, map, 'disabled_reason') || '').trim(),
  };
}

function getPositiveNumberSetting_(key, fallback, minValue) {
  const raw = Number(getSettingsMap_()[key]);
  if (!Number.isFinite(raw)) return fallback;
  return Math.max(minValue || 0, raw);
}

function getResidentStaleAfterDays_() {
  return getPositiveNumberSetting_('resident_stale_after_days', 14, 1);
}

function getResidentDisableAfterDays_() {
  return getPositiveNumberSetting_('resident_disable_after_days', 120, 1);
}

function getDisabledAccountDeleteAfterDays_() {
  return getPositiveNumberSetting_('disabled_account_delete_after_days', 180, 1);
}

function getActivityTouchCooldownSeconds_() {
  const hours = getPositiveNumberSetting_('activity_touch_cooldown_hours', 6, 0);
  return Math.round(hours * 3600);
}

function getResidentLastActivityValue_(resident) {
  if (!resident) return '';
  return resident.lastActivityAt ||
    resident.lastLoginAt ||
    resident.registeredAt ||
    resident.updatedAt ||
    resident.createdAt ||
    '';
}

function getResidentLastActivityDate_(resident) {
  const d = valueToDateTime_(getResidentLastActivityValue_(resident));
  return isNaN(d.getTime()) ? new Date(NaN) : d;
}

function getResidentDisabledReferenceValue_(resident) {
  if (!resident) return '';
  return resident.disabledAt || resident.updatedAt || resident.createdAt || '';
}

function getResidentDisabledReferenceDate_(resident) {
  const d = valueToDateTime_(getResidentDisabledReferenceValue_(resident));
  return isNaN(d.getTime()) ? new Date(NaN) : d;
}

function residentHasRelevantReservation_(residentId) {
  return getConfirmedReservationsForResident_(residentId).some(r => {
    const slot = findSlotById_(r.slotId);
    if (!slot || !slot.endIso) return false;
    const end = valueToDateTime_(slot.endIso);
    return !isNaN(end.getTime()) && end.getTime() > Date.now();
  });
}

function isResidentStale_(resident) {
  if (!resident || !resident.active) return false;
  if (residentHasRelevantReservation_(resident.residentId)) return false;

  const lastActivity = getResidentLastActivityDate_(resident);
  if (isNaN(lastActivity.getTime())) return false;

  const cutoffMs = Date.now() - getResidentStaleAfterDays_() * 24 * 60 * 60 * 1000;
  return lastActivity.getTime() <= cutoffMs;
}

function isResidentBlockingRoomCapacity_(resident) {
  return Boolean(resident && resident.active && !isResidentStale_(resident));
}

function shouldAutoDisableResident_(resident) {
  if (!resident || !resident.active) return false;
  if (residentHasRelevantReservation_(resident.residentId)) return false;

  const lastActivity = getResidentLastActivityDate_(resident);
  if (isNaN(lastActivity.getTime())) return false;

  const cutoffMs = Date.now() - getResidentDisableAfterDays_() * 24 * 60 * 60 * 1000;
  return lastActivity.getTime() <= cutoffMs;
}

function shouldDeleteDisabledResident_(resident) {
  if (!resident || resident.active) return false;
  if (residentHasRelevantReservation_(resident.residentId)) return false;

  const disabledRef = getResidentDisabledReferenceDate_(resident);
  if (isNaN(disabledRef.getTime())) return false;

  const cutoffMs = Date.now() - getDisabledAccountDeleteAfterDays_() * 24 * 60 * 60 * 1000;
  return disabledRef.getTime() <= cutoffMs;
}

function appendNote_(existingValue, noteToAppend) {
  const existing = String(existingValue || '').trim();
  const note = String(noteToAppend || '').trim();
  if (!note) return existing;
  return existing ? `${existing} | ${note}` : note;
}

// ---------------------------
// Reservations
// ---------------------------

function createReservation_(resident, slot, status, note) {
  const { sheet, map } = getSheetRef_(APP_CONFIG.SHEETS.RESERVATIONS);
  const now = new Date();
  const reservationId = `RESV-${Utilities.getUuid().split('-')[0].toUpperCase()}`;
  if (status === 'confirmed' && getConfirmedReservationsForResident_(resident.residentId)
    .some(r => r.weekKey === slot.weekKey)) {
    throw new Error('You already used or reserved a laundry slot in this week.');
  }
  const machineNo = status === 'confirmed' ? getNextMachineNoForSlot_(slot.slotId) : 0;
  if (status === 'confirmed' && !machineNo) {
    throw new Error('This slot is already full. Please refresh and choose another slot.');
  }

  const row = new Array(APP_CONFIG.HEADERS.RESERVATIONS.length).fill('');
  setCellValue_(row, map, 'reservation_id', reservationId);
  setCellValue_(row, map, 'resident_id', resident.residentId);
  setCellValue_(row, map, 'room', resident.room);
  setCellValue_(row, map, 'display_name', resident.displayName);
  setCellValue_(row, map, 'slot_id', slot.slotId);
  setCellValue_(row, map, 'week_key', slot.weekKey);
  setCellValue_(row, map, 'status', status);
  setCellValue_(row, map, 'created_at', now);
  setCellValue_(row, map, 'updated_at', now);
  setCellValue_(row, map, 'note', note || '');
  setCellValue_(row, map, 'machine_no', machineNo || '');
  setCellValue_(row, map, 'resident_label', formatResidentLabel_(resident.displayName, resident.room));

  sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);
  // Commit the row before releasing the caller's script lock; otherwise the
  // next locked request may read Sheets before this reservation is visible.
  SpreadsheetApp.flush();
  clearExecutionCaches_();
  bumpDataCacheVersion_();
  return { reservationId, machineNo };
}

function updateReservationStatus_(reservationId, newStatus, note) {
  const info = getReservationsSheetInfo_();
  const { sheet, values, map } = info;
  for (let i = 1; i < values.length; i++) {
    if (String(getCellValue_(values[i], map, 'reservation_id') || '').trim() !== reservationId) continue;

    setCellValue_(values[i], map, 'status', newStatus);
    setCellValue_(values[i], map, 'updated_at', new Date());
    if (note !== undefined && note !== null && note !== '') {
      setCellValue_(values[i], map, 'note', note);
    }

    if (newStatus === 'cancelled' || newStatus === 'moved') {
      appendRowsToSheet_(APP_CONFIG.SHEETS.RESERVATIONS_ARCHIVE, [values[i]], APP_CONFIG.HEADERS.RESERVATIONS);
      sheet.deleteRow(i + 1);
      clearExecutionCaches_();
      bumpDataCacheVersion_();
      return true;
    }

    sheet.getRange(i + 1, 1, 1, values[i].length).setValues([values[i]]);
    clearExecutionCaches_();
    bumpDataCacheVersion_();
    return true;
  }
  return false;
}

function getUsedMachineNumbersForSlot_(slotId) {
  return (getConfirmedReservationsBySlotMap_()[slotId] || [])
    .map(r => Number(r.machineNo || 0))
    .filter(n => n >= 1 && n <= APP_CONFIG.SLOT_CAPACITY)
    .sort((a, b) => a - b);
}

function getNextMachineNoForSlot_(slotId) {
  const confirmedReservations = getConfirmedReservationsBySlotMap_()[slotId] || [];
  const workingMachines = getWorkingMachineNumbers_();

  // A reservation that is still assigned to a newly broken machine still
  // consumes one operational booking place. Otherwise new residents could
  // fill every working machine and leave that resident with nowhere to switch.
  if (confirmedReservations.length >= workingMachines.length) {
    return 0;
  }

  const used = new Set(
    confirmedReservations
      .map(r => Number(r.machineNo || 0))
      .filter(n => n >= 1 && n <= APP_CONFIG.SLOT_CAPACITY)
  );

  for (const machineNo of workingMachines) {
    if (!used.has(machineNo)) {
      return machineNo;
    }
  }

  return 0;
}

function getBookableMachineCountForSlot_(slotId) {
  const confirmedReservations = getConfirmedReservationsBySlotMap_()[slotId] || [];
  const workingMachines = getWorkingMachineNumbers_();

  const remainingOperationalPlaces = Math.max(
    0,
    workingMachines.length - confirmedReservations.length
  );

  if (remainingOperationalPlaces <= 0) return 0;

  const used = new Set(
    confirmedReservations
      .map(r => Number(r.machineNo || 0))
      .filter(n => n >= 1 && n <= APP_CONFIG.SLOT_CAPACITY)
  );

  const freeWorkingMachines = workingMachines
    .filter(machineNo => !used.has(machineNo))
    .length;

  return Math.max(
    0,
    Math.min(freeWorkingMachines, remainingOperationalPlaces)
  );
}

function syncReservationMetadata_() {
  const { sheet, values, map } = getReservationsSheetInfo_();
  if (values.length < 2) return;

  const slotBuckets = {};
  for (let i = 1; i < values.length; i++) {
    const reservationId = String(getCellValue_(values[i], map, 'reservation_id') || '').trim();
    if (!reservationId) continue;
    const slotId = String(getCellValue_(values[i], map, 'slot_id') || '').trim();
    if (!slotBuckets[slotId]) slotBuckets[slotId] = [];
    slotBuckets[slotId].push({
      rowIndex: i,
      status: String(getCellValue_(values[i], map, 'status') || '').trim(),
      createdAt: getCellValue_(values[i], map, 'created_at'),
      machineNo: Number(getCellValue_(values[i], map, 'machine_no') || 0),
      displayName: normalizeDisplayName_(getCellValue_(values[i], map, 'display_name')),
      room: normalizeRoom_(getCellValue_(values[i], map, 'room')),
      residentLabel: String(getCellValue_(values[i], map, 'resident_label') || '').trim(),
    });
  }

  let changed = false;
  Object.keys(slotBuckets).forEach(slotId => {
    const bucket = slotBuckets[slotId]
      .filter(item => item.status === 'confirmed')
      .sort((a, b) => valueToDate_(a.createdAt).getTime() - valueToDate_(b.createdAt).getTime());

    const used = new Set();
    bucket.forEach(item => {
      let machineNo = item.machineNo;
      if (!(machineNo >= 1 && machineNo <= APP_CONFIG.SLOT_CAPACITY) || used.has(machineNo)) {
        machineNo = 0;
        
        const workingMachines = getWorkingMachineNumbers_();

        for (const n of workingMachines) {
          if (!used.has(n)) {
            machineNo = n;
            break;
          }
        }
      }
      if (machineNo && machineNo !== item.machineNo) {
        setCellValue_(values[item.rowIndex], map, 'machine_no', machineNo);
        changed = true;
      }
      if (machineNo) used.add(machineNo);

      const residentLabel = formatResidentLabel_(item.displayName, item.room);
      if (residentLabel !== item.residentLabel) {
        setCellValue_(values[item.rowIndex], map, 'resident_label', residentLabel);
        changed = true;
      }
    });
  });

  if (!changed) return;

  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  clearExecutionCaches_();
  bumpDataCacheVersion_();
}

function getAvailableMachineNumbersForSlot_(slotId, excludeReservationId) {
  const used = new Set(
    (getConfirmedReservationsBySlotMap_()[slotId] || [])
      .filter(r => r.reservationId !== excludeReservationId)
      .map(r => Number(r.machineNo || 0))
      .filter(n => n >= 1 && n <= APP_CONFIG.SLOT_CAPACITY)
  );

  return getWorkingMachineNumbers_()
    .filter(machineNo => !used.has(machineNo));
}

function canChangeMachineForSlot_(slot) {
  if (!slot || !slot.startIso || !slot.endIso) return false;

  const now = Date.now();
  const start = valueToDateTime_(slot.startIso);
  const end = valueToDateTime_(slot.endIso);

  if (isNaN(start.getTime()) || isNaN(end.getTime())) return false;

  return now >= start.getTime() && now < end.getTime();
}

function updateReservationMachine_(reservationId, newMachineNo, note) {
  const info = getReservationsSheetInfo_();
  const { sheet, values, map } = info;

  for (let i = 1; i < values.length; i++) {
    if (String(getCellValue_(values[i], map, 'reservation_id') || '').trim() !== reservationId) continue;

    setCellValue_(values[i], map, 'machine_no', Number(newMachineNo));
    setCellValue_(values[i], map, 'updated_at', new Date());

    if (note) {
      setCellValue_(values[i], map, 'note', note);
    }

    sheet.getRange(i + 1, 1, 1, values[i].length).setValues([values[i]]);
    clearExecutionCaches_();
    bumpDataCacheVersion_();
    return true;
  }

  return false;
}

function reportMachineProblem(room, accessCode, machineNo, description, reservationId) {
  ensureSchema_();
  ensureCycleOpenedIfDue_();
  const resident = requireResident_(room, accessCode);

  machineNo = Number(machineNo || 0);
  description = String(description || '').trim();
  reservationId = String(reservationId || '').trim();

  const machineExists = getMachines_().some(machine => machine.machineNo === machineNo);
  if (!machineExists) {
    throw new Error('Select a valid laundry machine.');
  }
  if (description.length < 3) {
    throw new Error('Please briefly describe the machine problem.');
  }
  if (description.length > APP_CONFIG.MACHINE_REPORT_DESCRIPTION_MAX_LENGTH) {
    throw new Error(
      `Problem description is too long. Use at most ${APP_CONFIG.MACHINE_REPORT_DESCRIPTION_MAX_LENGTH} characters.`
    );
  }

  let reportId = '';
  let context = {};

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Machine reporting is busy right now. Please retry in a moment.'
  );

  try {
    if (reservationId) {
      const reservation = findReservationById_(reservationId);

      if (!reservation || reservation.residentId !== resident.residentId) {
        throw new Error('The reservation attached to this report is not available.');
      }

      if (Number(reservation.machineNo || 0) !== machineNo) {
        throw new Error('The selected machine does not match the attached reservation.');
      }

      context = {
        reservationId: reservation.reservationId,
        slotId: reservation.slotId,
      };
    }

    reportId = appendMachineReport_(resident, machineNo, description, context);

    const contextText = context.reservationId
      ? ` | reservation ${context.reservationId}`
      : '';

    logAudit_(
      resident.residentId,
      'report_machine_problem',
      reportId,
      `${resident.displayName} reported Machine ${machineNo}: ${description}${contextText}`
    );
  } finally {
    lock.releaseLock();
  }

  touchResidentActivity_(resident.residentId, { force: true });

  const dashboard = decorateDashboardAuth_(getDashboardData_(resident), resident);
  dashboard.machineReportResult = { reportId, machineNo };
  return dashboard;
}

function changeReservationMachine(room, accessCode, reservationId, newMachineNo) {
  ensureSchema_();
  ensureCycleOpenedIfDue_();
  const resident = requireResident_(room, accessCode);

  let slotDate = '';

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Too many residents are changing reservations right now. Please retry in a moment.'
  );

  try {
    const reservation = findReservationById_(reservationId);
    if (!reservation || reservation.status !== 'confirmed') {
      throw new Error('Reservation was not found.');
    }
    if (reservation.residentId !== resident.residentId) {
      throw new Error('You can only change the machine for your own reservation.');
    }

    const slot = findSlotById_(reservation.slotId);
    if (!slot) throw new Error('Slot was not found.');
    if (isSlotConcluded_(slot)) {
      throw new Error('This slot has already concluded and the machine can no longer be changed.');
    }

    if (!canChangeMachineForSlot_(slot)) {
      throw new Error('Machine changes are allowed only after the slot starts.');
    }

    newMachineNo = Number(newMachineNo || 0);
    if (!(newMachineNo >= 1 && newMachineNo <= slot.capacity)) {
      throw new Error('Selected machine is invalid.');
    }

    if (!isMachineWorking_(newMachineNo)) {
      throw new Error('That machine is currently out of order.');
    }

    const currentMachineNo = Number(reservation.machineNo || 0);
    if (newMachineNo === currentMachineNo) {
      throw new Error('You already have that machine.');
    }

    const availableMachineNos = getAvailableMachineNumbersForSlot_(slot.slotId, reservationId);
    if (availableMachineNos.indexOf(newMachineNo) === -1) {
      throw new Error('That machine is no longer available.');
    }

    updateReservationMachine_(
      reservationId,
      newMachineNo,
      `Machine changed from ${currentMachineNo} to ${newMachineNo}`
    );

    slotDate = normalizeDateCellToYmd_(slot.slotDate);

    logAudit_(
      resident.residentId,
      'change_machine',
      reservationId,
      `${resident.displayName} changed machine from ${currentMachineNo} to ${newMachineNo}`
    );
  } finally {
    lock.releaseLock();
  }

  touchResidentActivity_(resident.residentId, { force: true });
  syncPublicBoardAfterReservationChange_(slotDate ? [slotDate] : []);
  return decorateDashboardAuth_(getDashboardData_(resident), resident);
}

// ---------------------------
// Archiving and audit
// ---------------------------

function archivePastReservationsNow(options) {
  options = options || {};

  const { sheet, values, map } = getReservationsSheetInfo_();
  if (values.length < 2) return 0;

  const settings = getSettingsMap_();
  const keepPastDays = Math.max(0, Number(settings.archive_keep_past_days || 2) || 2);
  const cutoff = addDays_(todayInTz_(), -keepPastDays);

  const activeCycleStart = normalizeDateCellToYmd_(settings.active_cycle_start);

  const forceEndedBeforeIso = normalizeDateTimeCellToIso_(options.forceEndedBeforeIso);
  const forceEndedBeforeDate = forceEndedBeforeIso ? valueToDateTime_(forceEndedBeforeIso) : null;

  const slotMap = Object.fromEntries(getAllSlots_().map(slot => [slot.slotId, slot]));

  const rowsToArchive = [];
  const rowNumbers = [];

  for (let i = 1; i < values.length; i++) {
    const status = String(getCellValue_(values[i], map, 'status') || '').trim();
    const slotId = String(getCellValue_(values[i], map, 'slot_id') || '').trim();
    const slot = slotMap[slotId];
    const endDate = slot && slot.endIso ? valueToDateTime_(slot.endIso) : null;

    const isActiveCycleReservation =
      status === 'confirmed' &&
      slot &&
      normalizeDateCellToYmd_(slot.cycleStartDate) === activeCycleStart;

    let shouldArchive = false;

    if (status && status !== 'confirmed') {
      shouldArchive = true;
    }

    if (!shouldArchive && status === 'confirmed' && !isActiveCycleReservation && endDate) {
      if (forceEndedBeforeDate && endDate.getTime() < forceEndedBeforeDate.getTime()) {
        shouldArchive = true;
      } else if (endDate.getTime() < cutoff.getTime()) {
        shouldArchive = true;
      }
    }

    if (!shouldArchive) continue;

    rowsToArchive.push(values[i]);
    rowNumbers.push(i + 1);
  }

  if (!rowsToArchive.length) return 0;

  appendRowsToSheet_(APP_CONFIG.SHEETS.RESERVATIONS_ARCHIVE, rowsToArchive, APP_CONFIG.HEADERS.RESERVATIONS);
  deleteRowsDescending_(sheet, rowNumbers);
  clearExecutionCaches_();
  bumpDataCacheVersion_();

  return rowsToArchive.length;
}

function trimAuditActiveNow() {
  const settings = getSettingsMap_();
  const info = getAuditSheetInfo_();
  const { sheet, values } = info;
  if (values.length <= 1) return 0;

  const keepRowsRaw = Number(settings.audit_keep_rows);
  const keepRows = Number.isFinite(keepRowsRaw) ? Math.max(0, Math.floor(keepRowsRaw)) : 1000;
  const totalDataRows = values.length - 1;
  if (totalDataRows <= keepRows) return 0;

  const removeCount = totalDataRows - keepRows;
  const rowsToArchive = values.slice(1, 1 + removeCount);

  appendRowsToSheet_(APP_CONFIG.SHEETS.AUDIT_ARCHIVE, rowsToArchive, APP_CONFIG.HEADERS.AUDIT);
  sheet.deleteRows(2, removeCount);

  clearExecutionCaches_();
  return removeCount;
}

function appendRowsToSheet_(sheetName, rows, headers) {
  if (!rows || !rows.length) return;
  const sheet = ensureSheet_(getBackendSpreadsheet_(), sheetName, headers);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function deleteRowsDescending_(sheet, rowNumbers) {
  if (!rowNumbers || !rowNumbers.length) return;
  rowNumbers.sort((a, b) => b - a).forEach(rowNum => sheet.deleteRow(rowNum));
}

function logAudit_(actor, action, targetId, details) {
  if (!isAuditEnabledFresh_()) return;
  const sheet = getBackendSpreadsheet_().getSheetByName(APP_CONFIG.SHEETS.AUDIT);
  sheet.appendRow([new Date(), actor, action, targetId, details]);
}

function isAuditEnabledFresh_() {
  return normalizeBoolean_(getSettingsMap_().audit_enabled);
}

function pruneStaleSlotsNow() {
  const { sheet, values, map } = getSlotsSheetInfo_();
  if (values.length <= 1) return 0;

  const window = getPublicBoardWindow_();
  const visibleDates = new Set(window.dates || []);
  const referencedSlotIds = new Set(getAllReservations_().map(r => r.slotId).filter(Boolean));
  const rowNumbers = [];

  for (let i = 1; i < values.length; i++) {
    const slotId = String(getCellValue_(values[i], map, 'slot_id') || '').trim();
    if (!slotId) continue;
    const slotDate = normalizeDateCellToYmd_(getCellValue_(values[i], map, 'slot_date'));
    const keep = referencedSlotIds.has(slotId) || visibleDates.has(slotDate);
    if (!keep) rowNumbers.push(i + 1);
  }

  if (!rowNumbers.length) return 0;
  deleteRowsDescending_(sheet, rowNumbers);
  clearExecutionCaches_();
  bumpDataCacheVersion_();

  return rowNumbers.length;
}

function nowLocalIso_() {
  return Utilities.formatDate(new Date(), APP_CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss");
}

function normalizeHhMm_(value) {
  if (value === '' || value === null || value === undefined) return '';

  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, APP_CONFIG.TZ, 'HH:mm');
  }

  const text = String(value).trim();

  const m = text.match(/^(\d{1,2}):(\d{2})$/);
  if (m) {
    return `${String(Number(m[1])).padStart(2, '0')}:${m[2]}`;
  }

  const d = new Date(text);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, APP_CONFIG.TZ, 'HH:mm');
  }

  return '';
}

function addMinutesToHhMm_(hhmm, minutesToAdd) {
  const normalized = normalizeHhMm_(hhmm);
  if (!normalized) return '';
  const parts = normalized.split(':');
  let total = Number(parts[0]) * 60 + Number(parts[1]) + Number(minutesToAdd || 0);
  total = ((total % 1440) + 1440) % 1440;
  const hh = String(Math.floor(total / 60)).padStart(2, '0');
  const mm = String(total % 60).padStart(2, '0');
  return `${hh}:${mm}`;
}

function computeCarryoverFinalizeAt_(openedAtLocalIso) {
  const openedIso = normalizeDateTimeCellToIso_(openedAtLocalIso);
  if (!openedIso) return '';
  const nextDay = formatDateYmd_(addDays_(parseYmdDate_(openedIso.slice(0, 10)), 1));
  return `${nextDay}T${APP_CONFIG.AUTOMATION_WINDOWS.FINALIZE_START}:00`;
}

function isSlotConcluded_(slot) {
  if (!slot || !slot.endIso) return false;
  const end = valueToDateTime_(slot.endIso);
  return !isNaN(end.getTime()) && end.getTime() <= Date.now();
}

function disableInactiveResidentsNow() {
  const { sheet, values, map } = getResidentsSheetInfo_();
  if (values.length < 2) return 0;

  const now = new Date();
  const nowIso = nowLocalIso_();
  let changed = 0;

  for (let i = 1; i < values.length; i++) {
    const resident = buildResidentObjectFromRow_(values[i], map);
    if (!shouldAutoDisableResident_(resident)) continue;

    setCellValue_(values[i], map, 'active', false);
    setCellValue_(values[i], map, 'access_code', '');
    setCellValue_(values[i], map, 'recovery_hash', '');
    setCellValue_(values[i], map, 'updated_at', now);
    setCellValue_(values[i], map, 'disabled_at', now);
    setCellValue_(values[i], map, 'disabled_reason', 'auto_inactive');
    setCellValue_(values[i], map, 'notes', appendNote_(getCellValue_(values[i], map, 'notes'), `Auto-disabled for inactivity on ${nowIso}`));
    changed += 1;
  }

  if (!changed) return 0;

  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  clearExecutionCaches_();
  bumpDataCacheVersion_();
  return changed;
}

function deleteOldDisabledResidentsNow() {
  const { sheet, values, map } = getResidentsSheetInfo_();
  if (values.length < 2) return 0;

  const rowNumbers = [];
  for (let i = 1; i < values.length; i++) {
    const resident = buildResidentObjectFromRow_(values[i], map);
    if (shouldDeleteDisabledResident_(resident)) {
      rowNumbers.push(i + 1);
    }
  }

  if (!rowNumbers.length) return 0;

  deleteRowsDescending_(sheet, rowNumbers);
  clearExecutionCaches_();
  bumpDataCacheVersion_();
  return rowNumbers.length;
}

// ---------------------------
// Utilities
// ---------------------------

function toSheetCellValue_(value) {
  return value === undefined || value === null ? '' : value;
}

function countActiveResidentsInRoomExcluding_(room, excludeResidentId) {
  room = normalizeRoom_(room);
  return getAllResidents_().filter(r =>
    isResidentBlockingRoomCapacity_(r) &&
    r.residentId !== excludeResidentId &&
    normalizeComparable_(r.room) === normalizeComparable_(room)
  ).length;
}
function syncLiveReservationIdentityForResident_(residentId, displayName, room) {
  const info = getReservationsSheetInfo_();
  const { sheet, values, map } = info;
  let changed = false;

  for (let i = 1; i < values.length; i++) {
    if (String(getCellValue_(values[i], map, 'resident_id') || '').trim() !== residentId) continue;
    if (String(getCellValue_(values[i], map, 'status') || '').trim() !== 'confirmed') continue;

    setCellValue_(values[i], map, 'display_name', normalizeDisplayName_(displayName));
    setCellValue_(values[i], map, 'room', normalizeRoom_(room));
    setCellValue_(values[i], map, 'resident_label', formatResidentLabel_(displayName, room));
    setCellValue_(values[i], map, 'updated_at', new Date());
    changed = true;
  }

  if (!changed) return false;

  sheet.getRange(1, 1, values.length, values[0].length).setValues(values);
  clearExecutionCaches_();
  bumpDataCacheVersion_();
  return true;
}

function updateMyProfile(room, accessCode, newDisplayName, newRoom) {
  ensureSchema_();
  ensureCycleOpenedIfDue_();

  const resident = requireResident_(room, accessCode);
  const changedDates = getConfirmedReservationsForResident_(resident.residentId)
  .map(r => {
    const slot = findSlotById_(r.slotId);
    return slot ? normalizeDateCellToYmd_(slot.slotDate) : '';
  })
  .filter(Boolean);

  newDisplayName = normalizeDisplayName_(newDisplayName);
  newRoom = normalizeRoom_(newRoom);

  if (!newDisplayName) throw new Error('Enter your full name.');
  if (newDisplayName.length < 2) throw new Error('Your name is too short.');

  assertAllowedRoom_(newRoom);

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Profile update is busy right now. Please retry in a moment.'
  );

  try {
    const settings = getSettingsMap_();
    const maxPerRoom = Math.max(
      1,
      Number(settings.max_active_residents_per_room || 2) || 2
    );

    const movingToDifferentRoom =
      normalizeComparable_(resident.room) !== normalizeComparable_(newRoom);

    if (movingToDifferentRoom) {
      const activeResidentsInTargetRoom = countActiveResidentsInRoomExcluding_(newRoom, resident.residentId);
      if (activeResidentsInTargetRoom >= maxPerRoom) {
        throw new Error(
          `Room ${newRoom} already has the maximum number of active residents (${maxPerRoom}).`
        );
      }
    }

    const info = getResidentsSheetInfo_();
    const { sheet, values, map } = info;

    let updated = false;
    for (let i = 1; i < values.length; i++) {
      if (String(getCellValue_(values[i], map, 'resident_id') || '').trim() !== resident.residentId) continue;

      setCellValue_(values[i], map, 'display_name', newDisplayName);
      setCellValue_(values[i], map, 'room', newRoom);
      setCellValue_(values[i], map, 'updated_at', new Date());
      setCellValue_(values[i], map, 'last_activity_at', new Date());
      setCellValue_(values[i], map, 'disabled_at', '');
      setCellValue_(values[i], map, 'disabled_reason', '');

      sheet.getRange(i + 1, 1, 1, values[i].length).setValues([values[i]]);
      updated = true;
      break;
    }

    if (!updated) {
      throw new Error('Resident profile was not found.');
    }

    clearExecutionCaches_();
    bumpDataCacheVersion_();

    syncLiveReservationIdentityForResident_(resident.residentId, newDisplayName, newRoom);
    if (changedDates.length) {
      syncPublicBoardAfterReservationChange_(changedDates);
    }

    logAudit_(
      resident.residentId,
      'self_update_profile',
      resident.residentId,
      `Updated profile to ${newDisplayName} / room ${newRoom}`
    );

    const freshResident = findResidentByRoomAndCode_(newRoom, resident.accessCode);
    if (!freshResident) throw new Error('Updated account could not be reloaded.');

    return decorateDashboardAuth_(getDashboardData_(freshResident), freshResident);
  } finally {
    lock.releaseLock();
  }
}

function isSlotRunningNow_(slot) {
  if (!slot || !slot.startIso || !slot.endIso) return false;

  const now = Date.now();
  const start = valueToDateTime_(slot.startIso);
  const end = valueToDateTime_(slot.endIso);

  if (isNaN(start.getTime()) || isNaN(end.getTime())) return false;
  return now >= start.getTime() && now < end.getTime();
}

function deleteMyAccount(room, accessCode) {
  ensureSchema_();
  ensureCycleOpenedIfDue_();

  const resident = requireResident_(room, accessCode);

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Account deletion is busy right now. Please retry in a moment.'
  );

  try {
    const confirmedReservations = getConfirmedReservationsForResident_(resident.residentId);
    const changedDates = [];

    for (let i = 0; i < confirmedReservations.length; i++) {
      const reservation = confirmedReservations[i];
      const slot = findSlotById_(reservation.slotId);
      if (!slot) continue;

      if (isSlotRunningNow_(slot)) {
        throw new Error('You cannot delete your account while one of your reserved slots is currently running.');
      }

      if (!isSlotConcluded_(slot)) {
        changedDates.push(normalizeDateCellToYmd_(slot.slotDate));
        updateReservationStatus_(
          reservation.reservationId,
          'cancelled',
          'Cancelled due to self account deletion'
        );
      }
    }

    const info = getResidentsSheetInfo_();
    const { sheet, values, map } = info;
    let found = false;

    for (let i = 1; i < values.length; i++) {
      if (String(getCellValue_(values[i], map, 'resident_id') || '').trim() !== resident.residentId) continue;

      setCellValue_(values[i], map, 'active', false);
      setCellValue_(values[i], map, 'access_code', '');
      setCellValue_(values[i], map, 'recovery_hash', '');
      setCellValue_(values[i], map, 'updated_at', new Date());
      setCellValue_(values[i], map, 'disabled_at', new Date());
      setCellValue_(values[i], map, 'disabled_reason', 'self_deleted');
      setCellValue_(values[i], map, 'notes', appendNote_(getCellValue_(values[i], map, 'notes'), 'Self-deactivated by resident'));

      sheet.getRange(i + 1, 1, 1, values[i].length).setValues([values[i]]);
      found = true;
      break;
    }

    if (!found) throw new Error('Resident profile was not found.');

    clearExecutionCaches_();
    bumpDataCacheVersion_();

    if (changedDates.length) {
      syncPublicBoardAfterReservationChange_(changedDates);
    }

    logAudit_(
      resident.residentId,
      'self_delete_account',
      resident.residentId,
      'Resident self-deactivated account'
    );

    return { deleted: true };
  } finally {
    lock.releaseLock();
  }
}

function getSheetRef_(sheetName) {
  const sheet = getBackendSpreadsheet_().getSheetByName(sheetName);
  return { sheet, map: getHeaderMap_(sheet) };
}

function getPublicBoardSlotLookupForDates_(dateKeys) {
  const allowedDates = new Set((dateKeys || []).map(normalizeDateCellToYmd_).filter(Boolean));
  const lookup = {};
  getAllSlots_()
    .filter(slot => allowedDates.has(normalizeDateCellToYmd_(slot.slotDate)))
    .forEach(slot => {
      lookup[`${normalizeDateCellToYmd_(slot.slotDate)}::${slot.slotLabel}`] = slot;
    });
  return lookup;
}

function getPublicBoardReservationLookupForSlotIds_(slotIds) {
  const allowedSlotIds = new Set((slotIds || []).filter(Boolean));
  const lookup = {};
  getAllReservations_()
    .filter(r => r.status === 'confirmed' && allowedSlotIds.has(r.slotId))
    .forEach(r => {
      lookup[`${r.slotId}::${Number(r.machineNo || 0)}`] = r.residentLabel || formatResidentLabel_(r.displayName, r.room);
    });
  return lookup;
}

function validateSignupInputs_(room, displayName, recoveryKey) {
  if (!room) throw new Error('Enter your room number.');
  assertAllowedRoom_(room);
  if (!displayName) throw new Error('Enter your full name.');
  if (displayName.length < 2) throw new Error('Your name is too short.');
  if (!recoveryKey || recoveryKey.length < 4) throw new Error('Choose a recovery key with at least 4 characters.');
}

function normalizeRoom_(value) {
  return String(value || '')
    .trim()
    .toUpperCase()
    .replace(/\s+/g, '');
}

function isAllowedRoom_(room) {
  room = normalizeRoom_(room);
  if (!room) return false;

  // ONLY FOR TESTING
  // Testing rooms: T01-T99
  let match = room.match(/^T(\d{2})$/);
  if (match) {
    const n = Number(match[1]);
    return n >= 1 && n <= 99;
  }

  // Ground floor rooms: P01-P14
  match = room.match(/^P(\d{2})$/);
  if (match) {
    const n = Number(match[1]);
    return n >= 1 && n <= 14;
  }

  // Upper floors: 101-114, 201-214, 301-314, 401-414, 501-514
  match = room.match(/^([1-5])(\d{2})$/);
  if (match) {
    const n = Number(match[2]);
    return n >= 1 && n <= 14;
  }

  return false;
}

function assertAllowedRoom_(room) {
  room = normalizeRoom_(room);
  if (!isAllowedRoom_(room)) {
    throw new Error(
      'Enter a valid room number!'
    );
  }
  return room;
}

function normalizeDisplayName_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ');
}

function normalizeAccessCode_(value) {
  return String(value || '').trim().toUpperCase();
}

function normalizeComparable_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toUpperCase();
}

function generateResidentId_() {
  return `RES-${Utilities.getUuid().split('-')[0].toUpperCase()}`;
}

function formatResidentLabel_(displayName, room) {
  return `${normalizeDisplayName_(displayName)} (${normalizeRoom_(room)})`;
}

function hashRecoveryKey_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value || '').trim(), Utilities.Charset.UTF_8);
  return bytes.map(b => {
    const v = b < 0 ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('').toUpperCase();
}

function generateUniqueCode_(usedCodes, length) {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '';
  do {
    code = '';
    for (let i = 0; i < length; i++) {
      code += alphabet[Math.floor(Math.random() * alphabet.length)];
    }
  } while (usedCodes.has(code));
  usedCodes.add(code);
  return code;
}

function getNextMonday_(date) {
  const d = valueToDate_(date);
  d.setHours(0, 0, 0, 0);
  const day = d.getDay(); // 0 Sun, 1 Mon
  const daysUntilNextMonday = day === 0 ? 1 : (8 - day) % 7;
  return addDays_(d, daysUntilNextMonday);
}

function todayInTz_() {
  return parseYmdDate_(Utilities.formatDate(new Date(), APP_CONFIG.TZ, 'yyyy-MM-dd'));
}

function weekKeyForDate_(date) {
  const d = valueToDate_(date);
  d.setHours(0, 0, 0, 0);
  const day = d.getDay();
  const diffToMonday = day === 0 ? -6 : 1 - day;
  const monday = addDays_(d, diffToMonday);
  return formatDateYmd_(monday);
}

function addDays_(date, days) {
  const d = valueToDate_(date);
  d.setDate(d.getDate() + days);
  return d;
}

function formatDateYmd_(date) {
  return Utilities.formatDate(valueToDate_(date), APP_CONFIG.TZ, 'yyyy-MM-dd');
}

function parseYmdDate_(value) {
  const text = String(value || '').trim();
  const m = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return new Date(NaN);
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function normalizeDateCellToYmd_(value) {
  if (!value && value !== 0) return '';
  if (value instanceof Date) return Utilities.formatDate(value, APP_CONFIG.TZ, 'yyyy-MM-dd');
  const text = String(value).trim();
  const m = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[1]}-${m[2]}-${m[3]}`;
  const d = new Date(text);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, APP_CONFIG.TZ, 'yyyy-MM-dd');
}

function normalizeDateTimeCellToIso_(value) {
  if (!value && value !== 0) return '';
  if (value instanceof Date) return Utilities.formatDate(value, APP_CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss");
  const text = String(value).trim();
  const m = text.match(/^(\d{4}-\d{2}-\d{2})(?:[T\s](\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (m) return `${m[1]}T${m[2] || '00'}:${m[3] || '00'}:${m[4] || '00'}`;
  const d = new Date(text);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, APP_CONFIG.TZ, "yyyy-MM-dd'T'HH:mm:ss");
}

function normalizeBoolean_(value) {
  if (value === true) return true;
  const text = String(value || '').trim().toUpperCase();
  return ['TRUE', 'YES', '1'].includes(text);
}

function valueToDate_(value) {
  if (value instanceof Date) return new Date(value.getTime());
  if (typeof value === 'number') return new Date(value);
  const ymd = normalizeDateCellToYmd_(value);
  if (ymd) return parseYmdDate_(ymd);
  const d = new Date(value);
  return isNaN(d.getTime()) ? new Date() : d;
}

function valueToDateTime_(value) {
  if (value instanceof Date) return new Date(value.getTime());
  if (typeof value === 'number') return new Date(value);
  const iso = normalizeDateTimeCellToIso_(value);
  if (iso) {
    const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})$/);
    if (m) {
      return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), Number(m[6]));
    }
  }
  return valueToDate_(value);
}

function valueToIsoString_(value) {
  if (value instanceof Date) return value.toISOString();
  const d = new Date(value);
  return isNaN(d.getTime()) ? '' : d.toISOString();
}

function formatPublicLongDateEn_(ymd) {
  const date = parseYmdDate_(ymd);
  const weekdays = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  return `${weekdays[date.getDay()]}, ${months[date.getMonth()]} ${date.getDate()}${getOrdinalSuffix_(date.getDate())} ${date.getFullYear()}`;
}

function getOrdinalSuffix_(day) {
  if (day >= 11 && day <= 13) return 'th';
  const last = day % 10;
  if (last === 1) return 'st';
  if (last === 2) return 'nd';
  if (last === 3) return 'rd';
  return 'th';
}

function isWithinWindow_(hhmm, start, end) {
  return hhmm >= start && hhmm < end;
}