const ADMIN_APP_CONFIG = {
  HTML_FILE: 'AdminPage',
  SHEETS: {
    RESIDENTS_ARCHIVE: 'ResidentsArchive',
  },
  SETTINGS: {
    ADMIN_ENABLED: 'admin_enabled',
    ADMIN_PIN_HASH: 'admin_pin_hash',
    ADMIN_SESSION_TTL_MINUTES: 'admin_session_ttl_minutes',
    ADMIN_ALLOWED_ROOMS: 'admin_allowed_rooms',
    ADMIN_ALLOWED_RESIDENT_IDS: 'admin_allowed_resident_ids',
  },
  SESSION_CACHE_PREFIX: 'admin_session',
  SESSION_VERSION_KEY: 'ADMIN_SESSION_VERSION',
  MAX_SESSION_TTL_MINUTES: 360,
};

function assertAdminAccess_(room, accessCode) {
  const resident = requireResident_(room, accessCode);
  if (!isResidentAllowedForAdmin_(resident)) {
    throw new Error('Admin access denied.');
  }
  return resident;
}

function ensureAdminSchema_() {
  const ss = getBackendSpreadsheet_();
  ensureSheet_(ss, ADMIN_APP_CONFIG.SHEETS.RESIDENTS_ARCHIVE, APP_CONFIG.HEADERS.RESIDENTS);

  const settingsSheet = ss.getSheetByName(APP_CONFIG.SHEETS.SETTINGS);
  const values = settingsSheet.getDataRange().getValues();
  const existingKeys = new Set(values.slice(1).map(row => String(row[0] || '').trim()));

  const defaults = [
    [ADMIN_APP_CONFIG.SETTINGS.ADMIN_ENABLED, 'TRUE', 'Whether the admin dashboard is enabled'],
    [ADMIN_APP_CONFIG.SETTINGS.ADMIN_PIN_HASH, '', 'SHA-256 hash of the admin PIN. Use setAdminPinPrompt() or setAdminPin("your-pin")'],
    [ADMIN_APP_CONFIG.SETTINGS.ADMIN_SESSION_TTL_MINUTES, '120', 'How long the admin session stays active'],
    [ADMIN_APP_CONFIG.SETTINGS.ADMIN_ALLOWED_ROOMS, '', 'Optional comma-separated room allowlist for admin access'],
    [ADMIN_APP_CONFIG.SETTINGS.ADMIN_ALLOWED_RESIDENT_IDS, '', 'Optional comma-separated resident_id allowlist for admin access'],
  ];

  const rowsToAdd = defaults.filter(row => !existingKeys.has(row[0]));
  if (rowsToAdd.length) {
    settingsSheet.getRange(settingsSheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
    clearExecutionCaches_();
    bumpDataCacheVersion_();
  }

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(ADMIN_APP_CONFIG.SESSION_VERSION_KEY)) {
    props.setProperty(ADMIN_APP_CONFIG.SESSION_VERSION_KEY, '1');
  }
}

function setAdminPinPrompt() {
  ensureSchema_();
  ensureAdminSchema_();

  const ui = getUiOrNull_();
  if (!ui) throw new Error('Open the bound spreadsheet to use this prompt.');

  const response = ui.prompt(
    'Set admin PIN',
    'Enter a new admin PIN. Use at least 6 characters.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const pin = String(response.getResponseText() || '').trim();
  setAdminPin(pin);
  ui.alert('Admin PIN saved successfully.');
}

function setAdminPin(adminPin) {
  ensureSchema_();
  ensureAdminSchema_();

  adminPin = String(adminPin || '').trim();
  if (adminPin.length < 6) {
    throw new Error('Choose an admin PIN with at least 6 characters.');
  }

  setSetting_(
    ADMIN_APP_CONFIG.SETTINGS.ADMIN_PIN_HASH,
    hashRecoveryKey_(adminPin),
    'SHA-256 hash of the admin PIN'
  );

  bumpAdminSessionVersion_();
  return { ok: true };
}

function adminRefreshDashboard(adminToken) {
  ensureSchema_();
  ensureAdminSchema_();
  const session = requireAdminSession_(adminToken);
  return buildAdminEnvelope_(session);
}

function adminSendMachineReportTestEmail(adminToken) {
  ensureSchema_();
  ensureAdminSchema_();
  requireAdminSession_(adminToken);
  // No report or machine status is changed by a diagnostic test.
  try {
    const result = sendConfiguredMachineReportEmail_(
      'Laundry Scheduler: test notification',
      'This is a test of the Dorm 14 laundry scheduler email notification system. No machine report or reservation was created.'
    );
    console.info('Admin test notification accepted by MailApp.');
    return { ok: true, message: `MailApp accepted the test email for ${result.recipientCount} recipient(s). Check your inbox and spam folder.` };
  } catch (err) {
    console.error('Admin test notification failed.', err);
    return { ok: false, message: String(err.message || err) };
  }
}

function adminLogout(adminToken) {
  const token = String(adminToken || '').trim();
  if (!token) return { ok: true };

  try {
    CacheService.getScriptCache().remove(getAdminSessionCacheKey_(token));
  } catch (err) {}

  return { ok: true };
}

function verifyAdminAccess(room, accessCode, adminPin) {
  ensureSchema_();
  ensureAdminSchema_();
  ensureCycleOpenedIfDue_();

  const resident = assertAdminAccess_(room, accessCode);
  const settings = getSettingsMap_();
  if (!normalizeBoolean_(settings[ADMIN_APP_CONFIG.SETTINGS.ADMIN_ENABLED])) {
    throw new Error('The admin dashboard is currently disabled.');
  }

  const storedHash = String(settings[ADMIN_APP_CONFIG.SETTINGS.ADMIN_PIN_HASH] || '').trim();
  if (!storedHash) {
    throw new Error('Admin PIN is not configured yet. Set it first with setAdminPinPrompt() or setAdminPin("your-pin").');
  }

  adminPin = String(adminPin || '').trim();
  if (!adminPin) {
    throw new Error('Enter the admin PIN.');
  }

  const enteredHash = hashRecoveryKey_(adminPin);
  if (enteredHash !== storedHash) {
    throw new Error('Admin PIN is incorrect.');
  }

  const session = createAdminSession_(resident);
  return buildAdminEnvelope_(session, { adminToken: session.token });
}

function buildAdminMachineCards_() {
  const reservationsByMachine = {};
  const reportsByMachine = {};

  getAllReservations_()
    .filter(reservation => reservation.status === 'confirmed')
    .forEach(reservation => {
      const slot = findSlotById_(reservation.slotId);
      if (!slot || isSlotConcluded_(slot)) return;

      const machineNo = Number(reservation.machineNo || 0);
      if (!machineNo) return;

      reservationsByMachine[machineNo] =
        (reservationsByMachine[machineNo] || 0) + 1;
    });

  getOpenMachineReports_().forEach(report => {
    reportsByMachine[report.machineNo] =
      (reportsByMachine[report.machineNo] || 0) + 1;
  });

  return getMachines_().map(machine => ({
    machineNo: machine.machineNo,
    status: machine.status,
    statusChangedAtIso: valueToIsoString_(machine.statusChangedAt),
    statusNote: machine.statusNote,
    openReservationCount: reservationsByMachine[machine.machineNo] || 0,
    openReportCount: reportsByMachine[machine.machineNo] || 0,
  }));
}

function buildAdminMachineReportCards_() {
  return getOpenMachineReports_()
    .map(report => {
      const slot = findSlotById_(report.slotId);

      return {
        reportId: report.reportId,
        machineNo: report.machineNo,
        residentId: report.residentId,
        room: report.room,
        displayName: report.displayName,
        reservationId: report.reservationId,
        slotId: report.slotId,
        reportedAtIso: valueToIsoString_(report.reportedAt),
        description: report.description,
        dayLabel: slot ? slot.dayLabel || '' : '',
        slotDate: slot ? slot.slotDate || '' : '',
        slotLabel: slot ? slot.slotLabel || '' : '',
        startIso: slot ? slot.startIso || '' : '',
        endIso: slot ? slot.endIso || '' : '',
      };
    })
    .sort((a, b) => String(b.reportedAtIso || '').localeCompare(String(a.reportedAtIso || '')));
}

function adminUpdateResidentProfile(adminToken, residentId, newDisplayName, newRoom, notes) {
  ensureSchema_();
  ensureAdminSchema_();
  ensureCycleOpenedIfDue_();

  const session = requireAdminSession_(adminToken);
  const resident = findResidentByIdAdmin_(residentId);
  if (!resident) throw new Error('Resident was not found.');

  newDisplayName = normalizeDisplayName_(newDisplayName);
  newRoom = normalizeRoom_(newRoom);
  notes = String(notes || '').trim();

  if (!newDisplayName) throw new Error('Enter the resident name.');
  if (newDisplayName.length < 2) throw new Error('The resident name is too short.');

  assertAllowedRoom_(newRoom);

  const changedDates = collectChangedDatesForResident_(resident.residentId);

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Admin profile update is busy right now. Please retry in a moment.'
  );

  try {
    if (resident.active && normalizeComparable_(resident.room) !== normalizeComparable_(newRoom)) {
      const settings = getSettingsMap_();
      const maxPerRoom = Math.max(
        1,
        Number(settings.max_active_residents_per_room || 2) || 2
      );

      const activeResidentsInTargetRoom = countActiveResidentsInRoomExcluding_(newRoom, resident.residentId);
      if (activeResidentsInTargetRoom >= maxPerRoom) {
        throw new Error(`Room ${newRoom} already has the maximum number of active residents (${maxPerRoom}).`);
      }
    }

    const rowInfo = getResidentRowInfoById_(resident.residentId);
    if (!rowInfo) throw new Error('Resident row was not found.');

    setCellValue_(rowInfo.row, rowInfo.map, 'display_name', newDisplayName);
    setCellValue_(rowInfo.row, rowInfo.map, 'room', newRoom);
    setCellValue_(rowInfo.row, rowInfo.map, 'notes', notes);
    setCellValue_(rowInfo.row, rowInfo.map, 'updated_at', new Date());

    rowInfo.sheet.getRange(rowInfo.rowNumber, 1, 1, rowInfo.row.length).setValues([rowInfo.row]);
    clearExecutionCaches_();
    bumpDataCacheVersion_();

    syncLiveReservationIdentityForResident_(resident.residentId, newDisplayName, newRoom);
    if (changedDates.length) {
      syncPublicBoardAfterReservationChange_(changedDates);
    }

    logAudit_(
      session.viewerResidentId,
      'admin_update_resident_profile',
      resident.residentId,
      `Updated resident to ${newDisplayName} / room ${newRoom}`
    );

    const freshSession = requireAdminSession_(session.token);
    return buildAdminEnvelope_(freshSession);
  } finally {
    lock.releaseLock();
  }
}

function adminResetResidentCode(adminToken, residentId) {
  ensureSchema_();
  ensureAdminSchema_();

  const session = requireAdminSession_(adminToken);
  const resident = findResidentByIdAdmin_(residentId);
  if (!resident) throw new Error('Resident was not found.');
  if (!resident.active) {
    throw new Error('The resident account is inactive. Reactivate it instead.');
  }

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Admin code reset is busy right now. Please retry in a moment.'
  );

  try {
    const newCode = resetResidentCodeByResidentId(resident.residentId);
    if (!newCode) throw new Error('Resident code could not be reset.');

    logAudit_(
      session.viewerResidentId,
      'admin_reset_resident_code',
      resident.residentId,
      `Reset access code for ${resident.displayName}`
    );

    const freshSession = requireAdminSession_(session.token);
    return buildAdminEnvelope_(freshSession, {
      generatedAccessCode: newCode,
      generatedAccessCodeLabel: `New access code for ${resident.displayName}`,
    });
  } finally {
    lock.releaseLock();
  }
}

function adminSetResidentRecoveryKey(adminToken, residentId, recoveryKey) {
  ensureSchema_();
  ensureAdminSchema_();

  const session = requireAdminSession_(adminToken);
  const resident = findResidentByIdAdmin_(residentId);
  if (!resident) throw new Error('Resident was not found.');

  recoveryKey = String(recoveryKey || '').trim();
  if (recoveryKey.length < 4) {
    throw new Error('Choose a recovery key with at least 4 characters.');
  }

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Admin recovery key update is busy right now. Please retry in a moment.'
  );

  try {
    const rowInfo = getResidentRowInfoById_(resident.residentId);
    if (!rowInfo) throw new Error('Resident row was not found.');

    setCellValue_(rowInfo.row, rowInfo.map, 'recovery_hash', hashRecoveryKey_(recoveryKey));
    setCellValue_(rowInfo.row, rowInfo.map, 'updated_at', new Date());

    rowInfo.sheet.getRange(rowInfo.rowNumber, 1, 1, rowInfo.row.length).setValues([rowInfo.row]);
    clearExecutionCaches_();
    bumpDataCacheVersion_();

    logAudit_(
      session.viewerResidentId,
      'admin_set_resident_recovery_key',
      resident.residentId,
      `Set recovery key for ${resident.displayName}`
    );

    const freshSession = requireAdminSession_(session.token);
    return buildAdminEnvelope_(freshSession);
  } finally {
    lock.releaseLock();
  }
}

function adminClearResidentRecoveryKey(adminToken, residentId) {
  ensureSchema_();
  ensureAdminSchema_();

  const session = requireAdminSession_(adminToken);
  const resident = findResidentByIdAdmin_(residentId);
  if (!resident) throw new Error('Resident was not found.');

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Admin recovery key clear is busy right now. Please retry in a moment.'
  );

  try {
    const rowInfo = getResidentRowInfoById_(resident.residentId);
    if (!rowInfo) throw new Error('Resident row was not found.');

    setCellValue_(rowInfo.row, rowInfo.map, 'recovery_hash', '');
    setCellValue_(rowInfo.row, rowInfo.map, 'updated_at', new Date());

    rowInfo.sheet.getRange(rowInfo.rowNumber, 1, 1, rowInfo.row.length).setValues([rowInfo.row]);
    clearExecutionCaches_();
    bumpDataCacheVersion_();

    logAudit_(
      session.viewerResidentId,
      'admin_clear_resident_recovery_key',
      resident.residentId,
      `Cleared recovery key for ${resident.displayName}`
    );

    const freshSession = requireAdminSession_(session.token);
    return buildAdminEnvelope_(freshSession);
  } finally {
    lock.releaseLock();
  }
}

function adminCancelResidentReservations(adminToken, residentId) {
  ensureSchema_();
  ensureAdminSchema_();
  ensureCycleOpenedIfDue_();

  const session = requireAdminSession_(adminToken);
  const resident = findResidentByIdAdmin_(residentId);
  if (!resident) throw new Error('Resident was not found.');

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Admin reservation cancellation is busy right now. Please retry in a moment.'
  );

  try {
    const result = cancelResidentOpenReservationsForAdmin_(resident.residentId, {
      note: 'Cancelled by admin',
      failIfRunning: true,
    });

    if (result.changedDates.length) {
      syncPublicBoardAfterReservationChange_(result.changedDates);
    }

    logAudit_(
      session.viewerResidentId,
      'admin_cancel_resident_reservations',
      resident.residentId,
      `Cancelled ${result.cancelledCount} reservation(s) for ${resident.displayName}`
    );

    const freshSession = requireAdminSession_(session.token);
    return buildAdminEnvelope_(freshSession);
  } finally {
    lock.releaseLock();
  }
}

function adminSetResidentActive(adminToken, residentId, shouldBeActive) {
  ensureSchema_();
  ensureAdminSchema_();
  ensureCycleOpenedIfDue_();

  const session = requireAdminSession_(adminToken);
  const resident = findResidentByIdAdmin_(residentId);
  if (!resident) throw new Error('Resident was not found.');

  const activate = normalizeBoolean_(shouldBeActive);

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Admin account status update is busy right now. Please retry in a moment.'
  );

  try {
    const rowInfo = getResidentRowInfoById_(resident.residentId);
    if (!rowInfo) throw new Error('Resident row was not found.');

    if (activate) {
      if (!resident.active) {
        const settings = getSettingsMap_();
        const maxPerRoom = Math.max(
          1,
          Number(settings.max_active_residents_per_room || 2) || 2
        );

        const activeResidentsInTargetRoom = countActiveResidentsInRoomExcluding_(resident.room, resident.residentId);
        if (activeResidentsInTargetRoom >= maxPerRoom) {
          throw new Error(`Room ${resident.room} already has the maximum number of active residents (${maxPerRoom}).`);
        }

        const usedCodes = new Set(getAllResidents_().map(r => r.accessCode).filter(Boolean));
        const newCode = generateUniqueCode_(usedCodes, APP_CONFIG.ACCESS_CODE_LENGTH);

        setCellValue_(rowInfo.row, rowInfo.map, 'active', true);
        setCellValue_(rowInfo.row, rowInfo.map, 'access_code', newCode);
        setCellValue_(rowInfo.row, rowInfo.map, 'updated_at', new Date());
        setCellValue_(rowInfo.row, rowInfo.map, 'last_activity_at', new Date());
        setCellValue_(rowInfo.row, rowInfo.map, 'disabled_at', '');
        setCellValue_(rowInfo.row, rowInfo.map, 'disabled_reason', '');

        rowInfo.sheet.getRange(rowInfo.rowNumber, 1, 1, rowInfo.row.length).setValues([rowInfo.row]);
        clearExecutionCaches_();
        bumpDataCacheVersion_();

        logAudit_(
          session.viewerResidentId,
          'admin_reactivate_resident',
          resident.residentId,
          `Reactivated ${resident.displayName}`
        );

        const freshSession = requireAdminSession_(session.token);
        return buildAdminEnvelope_(freshSession, {
          generatedAccessCode: newCode,
          generatedAccessCodeLabel: `New access code for ${resident.displayName}`,
        });
      }

      const freshSession = requireAdminSession_(session.token);
      return buildAdminEnvelope_(freshSession);
    }

    const cancelResult = cancelResidentOpenReservationsForAdmin_(resident.residentId, {
      note: 'Cancelled because resident was deactivated by admin',
      failIfRunning: true,
    });

    setCellValue_(rowInfo.row, rowInfo.map, 'active', false);
    setCellValue_(rowInfo.row, rowInfo.map, 'access_code', '');
    setCellValue_(rowInfo.row, rowInfo.map, 'recovery_hash', '');
    setCellValue_(rowInfo.row, rowInfo.map, 'updated_at', new Date());
    setCellValue_(rowInfo.row, rowInfo.map, 'disabled_at', new Date());
    setCellValue_(rowInfo.row, rowInfo.map, 'disabled_reason', 'admin_deactivated');
    setCellValue_(
      rowInfo.row,
      rowInfo.map,
      'notes',
      appendAdminNote_(getCellValue_(rowInfo.row, rowInfo.map, 'notes'), 'Admin deactivated account')
    );

    rowInfo.sheet.getRange(rowInfo.rowNumber, 1, 1, rowInfo.row.length).setValues([rowInfo.row]);
    clearExecutionCaches_();
    bumpDataCacheVersion_();

    if (cancelResult.changedDates.length) {
      syncPublicBoardAfterReservationChange_(cancelResult.changedDates);
    }

    logAudit_(
      session.viewerResidentId,
      'admin_deactivate_resident',
      resident.residentId,
      `Deactivated ${resident.displayName}`
    );

    const freshSession = requireAdminSession_(session.token);
    return buildAdminEnvelope_(freshSession);
  } finally {
    lock.releaseLock();
  }
}

function adminDeleteResident(adminToken, residentId) {
  ensureSchema_();
  ensureAdminSchema_();
  ensureCycleOpenedIfDue_();

  const session = requireAdminSession_(adminToken);
  const resident = findResidentByIdAdmin_(residentId);
  if (!resident) throw new Error('Resident was not found.');

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Admin resident deletion is busy right now. Please retry in a moment.'
  );

  try {
    const cancelResult = cancelResidentOpenReservationsForAdmin_(resident.residentId, {
      note: 'Cancelled because resident was deleted by admin',
      failIfRunning: true,
    });

    const rowInfo = getResidentRowInfoById_(resident.residentId);
    if (!rowInfo) throw new Error('Resident row was not found.');

    const rowForArchive = rowInfo.row.slice();
    setCellValue_(
      rowForArchive,
      rowInfo.map,
      'notes',
      appendAdminNote_(getCellValue_(rowForArchive, rowInfo.map, 'notes'), `Deleted permanently by admin ${session.viewerDisplayName} on ${nowLocalIso_()}`)
    );
    appendRowsToSheet_(ADMIN_APP_CONFIG.SHEETS.RESIDENTS_ARCHIVE, [rowForArchive], APP_CONFIG.HEADERS.RESIDENTS);

    rowInfo.sheet.deleteRow(rowInfo.rowNumber);
    clearExecutionCaches_();
    bumpDataCacheVersion_();

    if (cancelResult.changedDates.length) {
      syncPublicBoardAfterReservationChange_(cancelResult.changedDates);
    }

    logAudit_(
      session.viewerResidentId,
      'admin_delete_resident',
      resident.residentId,
      `Deleted ${resident.displayName} permanently`
    );

    const freshSession = requireAdminSession_(session.token);
    return buildAdminEnvelope_(freshSession);
  } finally {
    lock.releaseLock();
  }
}

function buildAdminEnvelope_(session, extra) {
  session = session || {};
  const payload = {
    session: {
      viewerResidentId: session.viewerResidentId || '',
      viewerDisplayName: session.viewerDisplayName || '',
      viewerRoom: session.viewerRoom || '',
      expiresAtIso: session.expiresAtIso || '',
    },
    dashboard: getAdminDashboardData_(),
  };

  extra = extra || {};
  Object.keys(extra).forEach(key => {
    payload[key] = extra[key];
  });

  return payload;
}

function getAdminDashboardData_() {
  ensureAdminSchema_();

  const residents = getAllResidents_()
    .map(buildAdminResidentCard_)
    .sort((a, b) => {
      if (a.active !== b.active) return a.active ? -1 : 1;
      const roomCompare = normalizeComparable_(a.room).localeCompare(normalizeComparable_(b.room));
      if (roomCompare !== 0) return roomCompare;
      return normalizeComparable_(a.displayName).localeCompare(normalizeComparable_(b.displayName));
    });

  const stats = {
    totalResidents: residents.length,
    activeResidents: residents.filter(r => r.active).length,
    inactiveResidents: residents.filter(r => !r.active).length,
    staleResidents: residents.filter(r => r.isStale).length,
    residentsWithRecoveryKey: residents.filter(r => r.hasRecoveryKey).length,
    openReservations: residents.reduce((sum, r) => sum + (r.reservationStats.openCount || 0), 0),
    openMachineReports: getOpenMachineReports_().length,
  };

  const reportNotificationSettings = getMachineReportNotificationSettings_();

  return {
    generatedAtIso: new Date().toISOString(),
    stats,
    machines: buildAdminMachineCards_(),
    machineReports: buildAdminMachineReportCards_(),
    machineReportNotifications: {
      enabled: reportNotificationSettings.enabled,
      email: reportNotificationSettings.email,
      configured: Boolean(reportNotificationSettings.email),
    },
    residents,
  };
}

function setMachineStatusLocked_(machineNo, status, note) {
  machineNo = Number(machineNo || 0);
  status = String(status || '').trim().toUpperCase();
  note = String(note || '').trim();

  if (machineNo < 1 || machineNo > APP_CONFIG.SLOT_CAPACITY) {
    throw new Error('Invalid machine number.');
  }
  if (!['WORKING', 'OUT_OF_ORDER'].includes(status)) {
    throw new Error('Invalid machine status.');
  }

  const { sheet, map } = getSheetRef_(APP_CONFIG.SHEETS.MACHINES);
  const values = sheet.getDataRange().getValues();

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowMachineNo = Number(getCellValue_(row, map, 'machine_no') || 0);
    if (rowMachineNo !== machineNo) continue;

    const oldStatus = String(getCellValue_(row, map, 'status') || '').trim().toUpperCase();

    setCellValue_(row, map, 'status', status);
    setCellValue_(row, map, 'status_changed_at', new Date());
    setCellValue_(row, map, 'status_note', status === 'WORKING' ? '' : note);

    sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
    clearExecutionCaches_();
    bumpDataCacheVersion_();

    return { machineNo, oldStatus, status };
  }

  throw new Error(`Machine ${machineNo} was not found.`);
}

function adminSetMachineStatus(adminToken, machineNo, status, note) {
  ensureSchema_();
  ensureAdminSchema_();

  const session = requireAdminSession_(adminToken);
  machineNo = Number(machineNo || 0);
  status = String(status || '').trim().toUpperCase();
  note = String(note || '').trim();

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Machine status is being updated. Please retry in a moment.'
  );

  try {
    _MACHINES_CACHE = null;
    _MACHINE_REPORTS_CACHE = null;
    const result = setMachineStatusLocked_(machineNo, status, note);

    if (status === 'OUT_OF_ORDER') {
      resolveOpenMachineReportsForMachine_(
        machineNo,
        session.viewerResidentId,
        note || `Machine ${machineNo} marked out of order by admin.`,
        APP_CONFIG.MACHINE_REPORT_STATUSES.MACHINE_OUT_OF_ORDER
      );
    } else if (status === 'WORKING') {
      // Marking a repaired machine working closes any reports that arrived
      // during that out-of-order period. They remain in MachineReports as
      // history, while the next breakdown starts with a clean active inbox.
      resolveOpenMachineReportsForMachine_(
        machineNo,
        session.viewerResidentId,
        `Machine ${machineNo} marked working again; previous issue cycle closed.`,
        APP_CONFIG.MACHINE_REPORT_STATUSES.MACHINE_RESTORED
      );
    }

    logAudit_(
      session.viewerResidentId,
      'admin_set_machine_status',
      String(machineNo),
      `Machine ${machineNo}: ${result.oldStatus || 'UNKNOWN'} -> ${status}` +
        (note ? ` | ${note}` : '')
    );

    SpreadsheetApp.flush();
    const freshSession = requireAdminSession_(session.token);
    return buildAdminEnvelope_(freshSession);
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}

function adminDismissMachineReport(adminToken, reportId, resolutionNote) {
  ensureSchema_();
  ensureAdminSchema_();

  const session = requireAdminSession_(adminToken);
  reportId = String(reportId || '').trim();
  resolutionNote = String(resolutionNote || '').trim();
  if (!reportId) throw new Error('Machine report is required.');

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Machine report is being updated. Please retry in a moment.'
  );

  try {
    _MACHINES_CACHE = null;
    _MACHINE_REPORTS_CACHE = null;
    const report = findMachineReportById_(reportId);
    if (!report) throw new Error('Machine report was not found.');
    if (report.status !== APP_CONFIG.MACHINE_REPORT_STATUSES.OPEN) {
      throw new Error('This machine report has already been resolved.');
    }

    resolveMachineReport_(
      reportId,
      APP_CONFIG.MACHINE_REPORT_STATUSES.DISMISSED,
      session.viewerResidentId,
      resolutionNote || 'Dismissed by admin.'
    );

    logAudit_(
      session.viewerResidentId,
      'admin_dismiss_machine_report',
      reportId,
      `Dismissed Machine ${report.machineNo} report from ${report.displayName} (room ${report.room})` +
        (resolutionNote ? ` | ${resolutionNote}` : '')
    );

    SpreadsheetApp.flush();
    const freshSession = requireAdminSession_(session.token);
    return buildAdminEnvelope_(freshSession);
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}

function adminMarkMachineOutOfOrderFromReport(adminToken, reportId, note) {
  ensureSchema_();
  ensureAdminSchema_();

  const session = requireAdminSession_(adminToken);
  reportId = String(reportId || '').trim();
  note = String(note || '').trim();
  if (!reportId) throw new Error('Machine report is required.');

  const lock = acquireScriptLockOrThrow_(
    APP_CONFIG.MUTATION_LOCK_TIMEOUT_MS,
    'Machine report is being processed. Please retry in a moment.'
  );

  try {
    _MACHINES_CACHE = null;
    _MACHINE_REPORTS_CACHE = null;
    const report = findMachineReportById_(reportId);
    if (!report) throw new Error('Machine report was not found.');
    if (report.status !== APP_CONFIG.MACHINE_REPORT_STATUSES.OPEN) {
      throw new Error('This machine report has already been resolved.');
    }

    const machineNote = note || report.description;
    const result = setMachineStatusLocked_(
      report.machineNo,
      'OUT_OF_ORDER',
      machineNote
    );

    resolveOpenMachineReportsForMachine_(
      report.machineNo,
      session.viewerResidentId,
      machineNote || `Machine ${report.machineNo} marked out of order from report ${reportId}.`
    );

    logAudit_(
      session.viewerResidentId,
      'admin_machine_out_of_order_from_report',
      reportId,
      `Machine ${report.machineNo}: ${result.oldStatus || 'UNKNOWN'} -> OUT_OF_ORDER from report ${reportId}` +
        (machineNote ? ` | ${machineNote}` : '')
    );

    SpreadsheetApp.flush();
    const freshSession = requireAdminSession_(session.token);
    return buildAdminEnvelope_(freshSession);
  } finally {
    try { SpreadsheetApp.flush(); } finally { lock.releaseLock(); }
  }
}

function buildAdminResidentCard_(resident) {
  const confirmedReservations = getConfirmedReservationsForResident_(resident.residentId)
    .map(r => {
      const slot = findSlotById_(r.slotId);
      if (!slot) return null;

      return {
        reservationId: r.reservationId,
        slotId: r.slotId,
        weekKey: r.weekKey,
        dayLabel: slot.dayLabel || '',
        slotLabel: slot.slotLabel || '',
        startIso: slot.startIso || '',
        endIso: slot.endIso || '',
        machineNo: r.machineNo || '',
        isConcluded: isSlotConcluded_(slot),
        isRunning: isSlotRunningNow_(slot),
      };
    })
    .filter(Boolean)
    .sort((a, b) => String(a.startIso || '').localeCompare(String(b.startIso || '')));

  const openReservations = confirmedReservations.filter(r => !r.isConcluded);
  const runningReservations = openReservations.filter(r => r.isRunning);
  const futureReservations = openReservations.filter(r => !r.isRunning);
  const isStale = isResidentStale_(resident);

  return {
    residentId: resident.residentId,
    room: resident.room,
    displayName: resident.displayName,
    accessCode: resident.accessCode || '',
    active: resident.active,
    notes: resident.notes || '',
    hasRecoveryKey: Boolean(resident.recoveryHash),
    createdAtIso: valueToIsoString_(resident.createdAt),
    updatedAtIso: valueToIsoString_(resident.updatedAt),
    registeredAtIso: valueToIsoString_(resident.registeredAt),
    lastLoginAtIso: valueToIsoString_(resident.lastLoginAt),
    lastActivityAtIso: valueToIsoString_(resident.lastActivityAt || getResidentLastActivityDate_(resident)),
    disabledAtIso: valueToIsoString_(resident.disabledAt),
    disabledReason: resident.disabledReason || '',
    isStale,
    blocksRoomCapacity: isResidentBlockingRoomCapacity_(resident),
    reservationStats: {
      confirmedCount: confirmedReservations.length,
      openCount: openReservations.length,
      runningCount: runningReservations.length,
      futureCount: futureReservations.length,
    },
    reservations: openReservations,
  };
}

function createAdminSession_(resident) {
  const token = Utilities.getUuid().replace(/-/g, '');
  const ttlSeconds = getAdminSessionTtlSeconds_();
  const expiresAt = new Date(Date.now() + ttlSeconds * 1000);

  const session = {
    token,
    version: getAdminSessionVersion_(),
    viewerResidentId: resident.residentId,
    viewerDisplayName: resident.displayName,
    viewerRoom: resident.room,
    createdAtIso: new Date().toISOString(),
    expiresAtIso: expiresAt.toISOString(),
  };

  CacheService.getScriptCache().put(
    getAdminSessionCacheKey_(token),
    JSON.stringify(session),
    ttlSeconds
  );

  return session;
}

function requireAdminSession_(adminToken) {
  ensureAdminSchema_();

  const token = String(adminToken || '').trim();
  if (!token) throw new Error('Admin session is missing. Sign in again.');

  const cache = CacheService.getScriptCache();
  const raw = cache.get(getAdminSessionCacheKey_(token));
  if (!raw) throw new Error('Admin session expired. Sign in again.');

  let session;
  try {
    session = JSON.parse(raw);
  } catch (err) {
    throw new Error('Admin session is invalid. Sign in again.');
  }

  if (String(session.version || '') !== getAdminSessionVersion_()) {
    cache.remove(getAdminSessionCacheKey_(token));
    throw new Error('Admin session expired. Sign in again.');
  }

  const viewer = findResidentByIdAdmin_(session.viewerResidentId);
  if (!viewer || !viewer.active) {
    cache.remove(getAdminSessionCacheKey_(token));
    throw new Error('Admin session expired. Sign in again.');
  }

  if (!isResidentAllowedForAdmin_(viewer)) {
    cache.remove(getAdminSessionCacheKey_(token));
    throw new Error('This resident account is not allowed to use the admin dashboard.');
  }

  const renewed = Object.assign({}, session, {
    viewerDisplayName: viewer.displayName,
    viewerRoom: viewer.room,
    expiresAtIso: new Date(Date.now() + getAdminSessionTtlSeconds_() * 1000).toISOString(),
  });

  cache.put(
    getAdminSessionCacheKey_(token),
    JSON.stringify(renewed),
    getAdminSessionTtlSeconds_()
  );

  return renewed;
}

function getAdminSessionCacheKey_(token) {
  return `${ADMIN_APP_CONFIG.SESSION_CACHE_PREFIX}:${token}`;
}

function getAdminSessionTtlSeconds_() {
  const settings = getSettingsMap_();
  const ttlMinutes = Math.min(
    ADMIN_APP_CONFIG.MAX_SESSION_TTL_MINUTES,
    Math.max(10, Number(settings[ADMIN_APP_CONFIG.SETTINGS.ADMIN_SESSION_TTL_MINUTES] || 120) || 120)
  );
  return ttlMinutes * 60;
}

function getAdminSessionVersion_() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty(ADMIN_APP_CONFIG.SESSION_VERSION_KEY) || '1';
}

function bumpAdminSessionVersion_() {
  const props = PropertiesService.getScriptProperties();
  const nextVersion = String(Number(props.getProperty(ADMIN_APP_CONFIG.SESSION_VERSION_KEY) || '1') + 1);
  props.setProperty(ADMIN_APP_CONFIG.SESSION_VERSION_KEY, nextVersion);
  return nextVersion;
}

function isResidentAllowedForAdmin_(resident) {
  const settings = getSettingsMap_();
  const allowedRooms = parseAdminCsvList_(settings[ADMIN_APP_CONFIG.SETTINGS.ADMIN_ALLOWED_ROOMS]);
  const allowedResidentIds = parseAdminCsvList_(settings[ADMIN_APP_CONFIG.SETTINGS.ADMIN_ALLOWED_RESIDENT_IDS]);

  if (!allowedRooms.length && !allowedResidentIds.length) {
    return false;
  }

  return allowedResidentIds.indexOf(normalizeComparable_(resident.residentId)) !== -1 ||
    allowedRooms.indexOf(normalizeComparable_(resident.room)) !== -1;
}

function parseAdminCsvList_(value) {
  return String(value || '')
    .split(',')
    .map(item => normalizeComparable_(item))
    .filter(Boolean);
}

function findResidentByIdAdmin_(residentId) {
  const id = String(residentId || '').trim();
  return getAllResidents_().find(r => r.residentId === id) || null;
}

function getResidentRowInfoById_(residentId) {
  const info = getResidentsSheetInfo_();
  const id = String(residentId || '').trim();

  for (let i = 1; i < info.values.length; i++) {
    if (String(getCellValue_(info.values[i], info.map, 'resident_id') || '').trim() !== id) continue;
    return {
      sheet: info.sheet,
      values: info.values,
      map: info.map,
      row: info.values[i],
      rowIndex: i,
      rowNumber: i + 1,
    };
  }

  return null;
}

function collectChangedDatesForResident_(residentId) {
  return Array.from(new Set(
    getConfirmedReservationsForResident_(residentId)
      .map(r => {
        const slot = findSlotById_(r.slotId);
        return slot ? normalizeDateCellToYmd_(slot.slotDate) : '';
      })
      .filter(Boolean)
  ));
}

function cancelResidentOpenReservationsForAdmin_(residentId, options) {
  options = options || {};
  const note = String(options.note || 'Cancelled by admin');
  const failIfRunning = Boolean(options.failIfRunning);

  const reservations = getConfirmedReservationsForResident_(residentId);
  const changedDates = [];
  let cancelledCount = 0;

  reservations.forEach(reservation => {
    const slot = findSlotById_(reservation.slotId);
    if (!slot) return;
    if (isSlotConcluded_(slot)) return;

    if (failIfRunning && isSlotRunningNow_(slot)) {
      throw new Error('This resident has a slot that is currently running. Try again after it finishes.');
    }

    changedDates.push(normalizeDateCellToYmd_(slot.slotDate));
    updateReservationStatus_(reservation.reservationId, 'cancelled', note);
    cancelledCount += 1;
  });

  return {
    cancelledCount,
    changedDates: Array.from(new Set(changedDates.filter(Boolean))),
  };
}

function appendAdminNote_(existingValue, noteToAppend) {
  const existing = String(existingValue || '').trim();
  const note = String(noteToAppend || '').trim();
  if (!note) return existing;
  return existing ? `${existing} | ${note}` : note;
}

function canResidentOpenAdminDashboard_(resident) {
  ensureAdminSchema_();
  const settings = getSettingsMap_();
  if (!normalizeBoolean_(settings[ADMIN_APP_CONFIG.SETTINGS.ADMIN_ENABLED])) {
    return false;
  }
  return isResidentAllowedForAdmin_(resident);
}