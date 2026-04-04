/*******************************
 * GESTION CONGES - WEB APP V2.2 PATCH
 * Backend Apps Script (Google Sheets + HTML)
 * Base V2.1 + correctifs ciblés
 * Ajouts:
 *  - règles de congé par équipe
 *  - dépôt jour unique depuis calendrier (API dédiée)
 *  - lecture détaillée d'une demande depuis calendrier
 *  - validation admin facilitée depuis calendrier (API détail)
 *  - correctif crédit mensuel + soldes crédités (support nombres locaux 2,5)
 *  - modèle mail validation congé (paramétrable)
 *******************************/

const DB_SPREADSHEET_ID = ''; // Laisse vide si script lié au Google Sheet (recommandé)

const APP = {
  TZ: 'Indian/Antananarivo',
  SESSION_HOURS: 8,
  LOGIN_MAX_ATTEMPTS: 5,
  LOGIN_BLOCK_MINUTES: 15,
  DEFAULT_ANNUAL_BALANCE: 0,
  MONTHLY_ACCRUAL_DAYS: 2.5,
  ROLES: { AGENT: 'agent', ADMIN: 'admin' },
  STATUS: {
    PENDING:  'En attente',
    APPROVED: 'Validé',
    REJECTED: 'Refusé',
    CANCELED: 'Annulé',
  },
  LEAVE_TYPES: ['Annuel','Maladie','Exceptionnel','Sans solde'],
  SHEETS: {
    USERS:      'Utilisateurs',
    REQUESTS:   'DemandesConges',
    BALANCES:   'SoldesConges',
    HOLIDAYS:   'JoursFeries',
    HISTORY:    'HistoriqueActions',
    PARAMS:     'Parametres',
    SESSIONS:   'Sessions',
    TEAM_RULES: 'ReglesEquipesConges',
  },
};

const SCHEMA = {
  Utilisateurs: [
    'userId','nomComplet','role','equipe',
    'passwordHash','salt','actif',
    'dateCreation','dernierLogin',
    'tentativesEchouees','compteBloqueJusqua'
  ],
  DemandesConges: [
    'idDemande','userId','nomAgent','equipe',
    'typeConge','dateDebut','dateFin','demiJournee',
    'nbJours','imputationSolde',
    'motif','statut','commentaireAdmin',
    'dateCreation','dateTraitement','adminTraitant'
  ],
  SoldesConges: [
    'userId','annee',
    'soldeInitial','congesPris','soldeRestant',
    'soldeNonConsommeInitial','congesNonConsommesPris','soldeNonConsommeRestant',
    'majLe','dernierMoisCredite'
  ],
  JoursFeries: ['date','libelle'],
  HistoriqueActions: [
    'idLog','idDemande','action',
    'auteurUserId','auteurRole',
    'cibleUserId','cibleEquipe',
    'dateHeure','details'
  ],
  Parametres: ['cle','valeur'],
  Sessions: [
    'token','userId','role',
    'createdAt','expiresAt','isActive','lastSeen'
  ],
  ReglesEquipesConges: [
    'equipe','maxDemandesParJour','delaiMinAvantCongeJours',
    'maxAbsentsParJour','heureLimiteDepot','actif','majLe'
  ]
};

/* ========= CACHE JOURS FERIES ========= */
let _holidayCache = null;
let _teamRulesCache = null;

function getHolidaySet_() {
  if (_holidayCache !== null) return _holidayCache;
  const rows = getRowsAsObjects_(APP.SHEETS.HOLIDAYS);
  _holidayCache = {};
  rows.forEach(r => {
    const iso = toIsoDate_(r.date);
    if (iso) _holidayCache[iso] = true;
  });
  return _holidayCache;
}

function clearHolidayCache_() {
  _holidayCache = null;
}

function clearTeamRulesCache_() {
  _teamRulesCache = null;
}

/* ========= ENTRYPOINT / UI ========= */

function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gestion des congés')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Congés')
    .addItem('1. Générer feuilles + entêtes', 'generateSheetsAndHeaders')
    .addItem('2. Créer premier admin (prompt)', 'createFirstAdminPrompt')
    .addSeparator()
    .addItem('3. Crédit mensuel maintenant (+2.5j)', 'runMonthlyAccrualNow')
    .addItem('4. Installer déclencheur mensuel', 'installMonthlyAccrualTrigger')
    .addSeparator()
    .addItem('Purger sessions expirées', 'purgeExpiredSessions')
    .addToUi();
}

/* ========= SETUP ========= */

function generateSheetsAndHeaders() {
  const ss = getDb_();
  Object.keys(SCHEMA).forEach(sheetName => {
    ensureSheetWithHeaders_(ss, sheetName, SCHEMA[sheetName]);
  });
  setParamIfMissing_('maxAbsentsParEquipe', '3');
  setParamIfMissing_('delaiMinAvantConge', '0');
  setParamIfMissing_('typesCongesAutorises', APP.LEAVE_TYPES.join(','));
  setParamIfMissing_('vueEquipeAgents', 'oui');
  setParamIfMissing_('accrualMensuelJours', normalizeNumberString_(APP.MONTHLY_ACCRUAL_DAYS));
  setParamIfMissing_('modeleMailConge', DEFAULT_MAIL_CONGE_TPL_);
  initializeAccrualCursorForExistingBalances_();
  formatSheets_();
  clearTeamRulesCache_();
  return { ok: true, message: 'Feuilles et entêtes générés.' };
}

function setupLeaveApp() { return generateSheetsAndHeaders(); }

function createFirstAdminPrompt() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Création Admin', 'ID admin (ex: ADM01)', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const userId = (r1.getResponseText() || '').trim().toUpperCase();
  const r2 = ui.prompt('Création Admin', 'Nom complet', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const nom = (r2.getResponseText() || '').trim();
  const r3 = ui.prompt('Création Admin', 'Mot de passe temporaire', ui.ButtonSet.OK_CANCEL);
  if (r3.getSelectedButton() !== ui.Button.OK) return;
  const password = r3.getResponseText();
  if (!userId || !nom || !password) { ui.alert('Champs invalides.'); return; }
  const ss = getDb_();
  ensureSheetWithHeaders_(ss, APP.SHEETS.USERS, SCHEMA[APP.SHEETS.USERS]);
  if (getUserByIdRaw_(userId)) { ui.alert('Cet ID existe déjà.'); return; }
  createOrUpdateUser_({ userId, nomComplet: nom, role: APP.ROLES.ADMIN, equipe: 'ADMIN', actif: 'oui', password });
  ui.alert('Admin créé avec succès : ' + userId);
}

function formatSheets_() {
  const ss = getDb_();
  Object.keys(SCHEMA).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn())
      .setFontWeight('bold').setBackground('#1f1f23').setFontColor('#d4af37');
    sh.autoResizeColumns(1, Math.max(1, sh.getLastColumn()));
  });
}

function ensureSheetWithHeaders_(ss, sheetName, headers) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  const lastCol = sh.getLastColumn();
  const existingHeaders = lastCol > 0
    ? sh.getRange(1, 1, 1, lastCol).getValues()[0].filter(String)
    : [];
  if (existingHeaders.length === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }
  const missing = headers.filter(h => existingHeaders.indexOf(h) === -1);
  if (missing.length) {
    sh.insertColumnsAfter(sh.getLastColumn(), missing.length);
    sh.getRange(1, existingHeaders.length + 1, 1, missing.length).setValues([missing]);
  }
}

/* ========= AUTH / SESSION ========= */

function apiLogin(userId, password) {
  userId   = String(userId   || '').trim().toUpperCase();
  password = String(password || '');
  if (!userId || !password) throw new Error('ID ou mot de passe manquant.');
  const user = getUserByIdRaw_(userId);
  if (!user) throw new Error('Identifiants invalides.');
  if (String(user.actif).toLowerCase() !== 'oui') throw new Error('Compte inactif.');
  const now = new Date();
  const blockedUntil = parseAnyDate_(user.compteBloqueJusqua);
  if (blockedUntil && blockedUntil.getTime() > now.getTime())
    throw new Error('Compte temporairement bloqué. Réessaie plus tard.');
  if (!verifyPassword_(password, user.salt, user.passwordHash)) {
    registerFailedLogin_(user);
    throw new Error('Identifiants invalides.');
  }
  resetFailedLogin_(user);
  updateUserField_(user.userId, 'dernierLogin', now);
  const session = createSession_(user);
  return { ok: true, token: session.token, user: publicUser_(user) };
}

function apiLogout(token) {
  requireSession_(token);
  invalidateSession_(token);
  return { ok: true };
}

function apiGetSession(token) {
  const s = getSessionByToken_(token);
  if (!s) throw new Error('Session invalide ou expirée.');
  const user = getUserByIdRaw_(s.userId);
  if (!user || String(user.actif).toLowerCase() !== 'oui') {
    invalidateSession_(token);
    throw new Error('Compte inactif.');
  }
  touchSession_(token);
  return { ok: true, user: publicUser_(user) };
}

function apiChangePassword(token, currentPassword, newPassword) {
  const session = requireSession_(token);
  const user = getUserByIdRaw_(session.userId);
  if (!user) throw new Error('Utilisateur introuvable.');
  if (!verifyPassword_(String(currentPassword || ''), user.salt, user.passwordHash))
    throw new Error('Mot de passe actuel incorrect.');
  if (!newPassword || String(newPassword).length < 6)
    throw new Error('Le nouveau mot de passe doit contenir au moins 6 caractères.');
  const salt = generateSalt_();
  updateUserFields_(user.userId, { salt, passwordHash: hashPassword_(String(newPassword), salt) });
  return { ok: true };
}

/* ========= API AGENT / ADMIN ========= */

function apiGetInitData(token) {
  const session = requireSession_(token);
  const user    = getUserByIdRaw_(session.userId);
  applyMonthlyAccrualForAll_();
  const year    = currentYear_();
  const balance = computeBalanceWithProvisional_(user.userId, year);
  return {
    ok: true,
    user: publicUser_(user),
    config: {
      leaveTypes:  getAllowedLeaveTypes_(),
      statuses:    APP.STATUS,
      currentYear: year
    },
    myBalance: balance
  };
}

function apiCreateLeave(token, payload) {
  const session = requireSession_(token);
  const user    = getUserByIdRaw_(session.userId);
  if (!user) throw new Error('Utilisateur introuvable.');
  applyMonthlyAccrualForUserYear_(user.userId, currentYear_());
  return createLeaveRequestForUser_(user, user, payload);
}

function apiCreateLeaveSingleDay(token, payload) {
  payload = payload || {};
  const date = toIsoDateString_(payload.date || payload.dateDebut);
  if (!date) throw new Error('Date invalide.');
  return apiCreateLeave(token, {
    typeConge: payload.typeConge,
    demiJournee: payload.demiJournee || 'Aucune',
    dateDebut: date,
    dateFin: date,
    motif: payload.motif || ''
  });
}

function apiAdminCreateLeaveSingleDayForAgent(token, payload) {
  const session = requireSession_(token);
  const admin   = requireAdmin_(session);
  payload = payload || {};
  const date = toIsoDateString_(payload.date || payload.dateDebut);
  const targetId = sanitizeText_(payload.userId).toUpperCase();
  if (!targetId) throw new Error('Agent requis.');
  if (!date) throw new Error('Date invalide.');
  const targetUser = getUserByIdRaw_(targetId);
  if (!targetUser) throw new Error('Agent introuvable.');
  return createLeaveRequestForUser_(targetUser, admin, {
    typeConge: payload.typeConge,
    demiJournee: payload.demiJournee || 'Aucune',
    dateDebut: date,
    dateFin: date,
    motif: payload.motif || '',
    statutInitial: payload.statutInitial || APP.STATUS.PENDING,
    modeValidation: payload.modeValidation || 'normal'
  });
}

/* ========= LISTING DEMANDES ========= */

function apiListLeaves(token, filters) {
  const session = requireSession_(token);
  const user    = getUserByIdRaw_(session.userId);
  const isAdmin = user.role === APP.ROLES.ADMIN;
  filters = filters || {};

  let scope = String(filters.scope || (isAdmin ? 'all' : 'mine')).trim().toLowerCase();
  if (['mine', 'group', 'all'].indexOf(scope) === -1) {
    scope = isAdmin ? 'all' : 'mine';
  }

  const meId     = normUpper_(user.userId || '');
  const meEquipe = normUpper_(user.equipe || '');

  const requests = getRowsAsObjects_(APP.SHEETS.REQUESTS).map(r => {
    const n = normalizeRequestRow_(r);
    n._userIdNorm = normUpper_(n.userId);
    n._equipeNorm = normUpper_(n.equipe);
    return n;
  });

  let out = requests.filter(r => {
    if (!isAdmin) {
      if (scope === 'group') return r._equipeNorm && r._equipeNorm === meEquipe;
      return r._userIdNorm === meId;
    }

    if (scope === 'mine') {
      return r._userIdNorm === meId;
    }

    if (scope === 'group') {
      if (!(r._equipeNorm && r._equipeNorm === meEquipe)) return false;
    }

    if (filters.userId) {
      if (r._userIdNorm !== normUpper_(filters.userId)) return false;
    }
    if (filters.statut && String(r.statut || '') !== String(filters.statut)) return false;
    if (filters.typeConge && String(r.typeConge || '') !== String(filters.typeConge)) return false;
    if (filters.equipe) {
      if (r._equipeNorm !== normUpper_(filters.equipe)) return false;
    }
    return true;
  });

  if (filters.annee) {
    const fYear = Number(filters.annee);
    out = out.filter(r => r.dateDebut && new Date(r.dateDebut + 'T00:00:00').getFullYear() === fYear);
  }

  out.sort((a, b) => {
    const da = new Date(a.dateCreation || a.dateDebut).getTime();
    const db = new Date(b.dateCreation || b.dateDebut).getTime();
    return db - da;
  });

  out = out.map(r => {
    delete r._userIdNorm;
    delete r._equipeNorm;
    return r;
  });

  const total    = out.length;
  const page     = Math.max(1, Number(filters.page || 1));
  const pageSize = Math.max(5, Math.min(500, Number(filters.pageSize || 100)));
  out = out.slice((page - 1) * pageSize, page * pageSize);

  return { ok: true, items: out, total, page, pageSize };
}

function apiGetLeaveDetails(token, idDemande) {
  const session = requireSession_(token);
  const user    = getUserByIdRaw_(session.userId);
  const isAdmin = user.role === APP.ROLES.ADMIN;
  const req = findLeaveById_(idDemande);
  if (!req) throw new Error('Demande introuvable.');

  if (!isAdmin) {
    const allowed = normUpper_(req.userId) === normUpper_(user.userId) ||
      normUpper_(req.equipe) === normUpper_(user.equipe);
    if (!allowed) throw new Error('Accès non autorisé.');
  }

  return {
    ok: true,
    item: req,
    permissions: {
      canValidate: isAdmin && req.statut === APP.STATUS.PENDING,
      canReject: isAdmin && req.statut === APP.STATUS.PENDING,
      canDelete: isAdmin,
      canCancel: !isAdmin && normUpper_(req.userId) === normUpper_(user.userId) && req.statut === APP.STATUS.PENDING
    }
  };
}

function apiProcessLeave(token, payload) {
  const session = requireSession_(token);
  const admin   = requireAdmin_(session);
  payload = payload || {};

  const idDemande      = normTrim_(payload.idDemande);
  const action         = normTrim_(payload.action);
  const commentaire    = sanitizeText_(payload.commentaire || '');
  const modeValRaw     = normTrim_(payload.modeValidation || 'normal').toLowerCase().replace(/[\s_-]/g,'');
  const modeValidation = (modeValRaw === 'nonconsomme') ? 'nonConsomme' : 'normal';

  if (!idDemande) throw new Error('ID demande manquant.');
  if (['validate','reject'].indexOf(action) === -1) throw new Error('Action invalide.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_(APP.SHEETS.REQUESTS);
    const data  = sheet.getDataRange().getValues();
    if (data.length < 2) throw new Error('Aucune demande.');

    const headers = data[0];
    const h = headerIndexMap_(headers);
    if (h.imputationSolde === undefined)
      throw new Error('Colonne imputationSolde manquante. Lance generateSheetsAndHeaders() puis recharge.');

    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (normTrim_(data[i][h.idDemande]) === idDemande) { targetRow = i + 1; break; }
    }
    if (targetRow === -1) throw new Error('Demande introuvable.');

    const rowVals = sheet.getRange(targetRow, 1, 1, headers.length).getValues()[0];
    const reqN    = normalizeRequestRow_(objectFromRow_(headers, rowVals));

    if (normTrim_(reqN.statut) !== normTrim_(APP.STATUS.PENDING))
      throw new Error("Cette demande n'est plus en attente.");

    let mailRendu = null;

    if (action === 'validate') {
      validateNoOverlap_(reqN.userId, reqN.dateDebut, reqN.dateFin, reqN.idDemande);
      validateTeamQuota_(reqN.equipe, reqN.dateDebut, reqN.dateFin, reqN.idDemande);

      const annee = new Date(reqN.dateDebut).getFullYear();
      applyMonthlyAccrualForUserYear_(reqN.userId, annee);

      const isAnnual       = typeConsumesBalance_(reqN.typeConge);
      const useNonConsumed = isAnnual && modeValidation === 'nonConsomme';

      if (isAnnual) {
        const bal = getOrCreateBalance_(reqN.userId, annee);
        if (useNonConsumed) {
          if (num_(bal.soldeNonConsommeRestant) < num_(reqN.nbJours))
            throw new Error('Solde non consommé insuffisant pour validation.');
        } else {
          if (num_(bal.soldeRestant) < num_(reqN.nbJours))
            throw new Error('Solde insuffisant pour validation.');
        }
      }

      rowVals[h.statut]           = APP.STATUS.APPROVED;
      rowVals[h.commentaireAdmin] = commentaire;
      rowVals[h.dateTraitement]   = new Date();
      rowVals[h.adminTraitant]    = admin.userId;
      rowVals[h.imputationSolde]  = isAnnual ? (useNonConsumed ? 'nonConsomme' : 'normal') : '';
      sheet.getRange(targetRow, 1, 1, headers.length).setValues([rowVals]);

      if (isAnnual) {
        recomputeBalance_(reqN.userId, annee);
        recomputeNonConsumedBalance_(reqN.userId, annee);
      }

      logHistory_(reqN.idDemande, 'validation', admin.userId, admin.role,
        `${commentaire || 'Validée'} | Imputation: ${isAnnual ? (useNonConsumed ? 'non consommé' : 'normale') : 'aucune'}`,
        { cibleUserId: reqN.userId, cibleEquipe: reqN.equipe });

      mailRendu = renderMailConge_(reqN, commentaire);

    } else {
      rowVals[h.statut]           = APP.STATUS.REJECTED;
      rowVals[h.commentaireAdmin] = commentaire;
      rowVals[h.dateTraitement]   = new Date();
      rowVals[h.adminTraitant]    = admin.userId;
      rowVals[h.imputationSolde]  = '';
      sheet.getRange(targetRow, 1, 1, headers.length).setValues([rowVals]);

      logHistory_(reqN.idDemande, 'refus', admin.userId, admin.role,
        commentaire || 'Refusée',
        { cibleUserId: reqN.userId, cibleEquipe: reqN.equipe });
    }
    return { ok: true, mailRendu: mailRendu };
  } finally {
    lock.releaseLock();
  }
}

function apiCancelLeave(token, idDemande) {
  const session = requireSession_(token);
  const user    = getUserByIdRaw_(session.userId);
  if (!user) throw new Error('Utilisateur introuvable.');
  idDemande = normTrim_(idDemande);
  if (!idDemande) throw new Error('ID demande manquant.');
  const userIdNorm = normUpper_(user.userId);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_(APP.SHEETS.REQUESTS);
    const data  = sheet.getDataRange().getValues();
    if (data.length < 2) throw new Error('Aucune demande.');
    const headers = data[0];
    const h       = headerIndexMap_(headers);
    let targetRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (normTrim_(data[i][h.idDemande]) === idDemande) { targetRow = i + 1; break; }
    }
    if (targetRow === -1) throw new Error('Demande introuvable.');
    const row = sheet.getRange(targetRow, 1, 1, headers.length).getValues()[0];
    const req = normalizeRequestRow_(objectFromRow_(headers, row));
    if (normUpper_(req.userId) !== userIdNorm) throw new Error('Tu peux annuler uniquement tes demandes.');
    if (normTrim_(req.statut) !== normTrim_(APP.STATUS.PENDING))
      throw new Error('Seules les demandes "En attente" peuvent être annulées.');
    row[h.statut]           = APP.STATUS.CANCELED;
    row[h.dateTraitement]   = new Date();
    row[h.adminTraitant]    = '';
    row[h.commentaireAdmin] = 'Annulée par agent';
    if (h.imputationSolde !== undefined) row[h.imputationSolde] = '';
    sheet.getRange(targetRow, 1, 1, headers.length).setValues([row]);
    logHistory_(req.idDemande, 'annulation_agent', user.userId, user.role,
      'Annulation par agent (En attente)',
      { cibleUserId: user.userId, cibleEquipe: String(user.equipe || '') });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function apiAdminCreateLeaveForAgent(token, payload) {
  const session = requireSession_(token);
  const admin   = requireAdmin_(session);
  payload = payload || {};
  const targetId = sanitizeText_(payload.userId).toUpperCase();
  if (!targetId) throw new Error('Agent requis.');
  const targetUser = getUserByIdRaw_(targetId);
  if (!targetUser) throw new Error('Agent introuvable.');
  return createLeaveRequestForUser_(targetUser, admin, payload);
}

function apiAdminDeleteLeave(token, payload) {
  const session = requireSession_(token);
  const admin   = requireAdmin_(session);
  payload = payload || {};
  const idDemande        = normTrim_(payload.idDemande);
  const motifSuppression = sanitizeText_(payload.motifSuppression || '');
  if (!idDemande) throw new Error('ID demande manquant.');
  if (!motifSuppression) throw new Error('Motif obligatoire pour suppression admin.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh   = getSheet_(APP.SHEETS.REQUESTS);
    const vals = sh.getDataRange().getValues();
    if (vals.length < 2) throw new Error('Aucune demande.');
    const headers = vals[0];
    const h       = headerIndexMap_(headers);
    let rowIdx = -1, req = null;
    for (let i = 1; i < vals.length; i++) {
      if (normTrim_(vals[i][h.idDemande]) === idDemande) {
        rowIdx = i + 1;
        req = normalizeRequestRow_(objectFromRow_(headers, vals[i]));
        break;
      }
    }
    if (rowIdx === -1 || !req) throw new Error('Demande introuvable.');
    sh.deleteRow(rowIdx);
    if (normTrim_(req.statut) === normTrim_(APP.STATUS.APPROVED) && typeConsumesBalance_(req.typeConge)) {
      const annee = new Date(req.dateDebut).getFullYear();
      recomputeBalance_(req.userId, annee);
      recomputeNonConsumedBalance_(req.userId, annee);
    }
    logHistory_(req.idDemande, 'suppression_admin', admin.userId, admin.role,
      `Suppression admin. Motif: ${motifSuppression}. Snapshot: ${req.nomAgent} / ${req.typeConge} / ${req.dateDebut}→${req.dateFin} / ${req.statut}`,
      { cibleUserId: req.userId, cibleEquipe: req.equipe || '' });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function apiGetCalendarEvents(token, filters) {
  const session = requireSession_(token);
  const user    = getUserByIdRaw_(session.userId);
  const isAdmin = user.role === APP.ROLES.ADMIN;

  let list;
  if (isAdmin) {
    list = apiListLeaves(token, Object.assign({
      scope: 'all',
      page: 1,
      pageSize: 500
    }, filters || {}));
  } else {
    list = apiListLeaves(token, {
      scope: 'group',
      page: 1,
      pageSize: 500
    });
  }

  const events = (list.items || []).map(r => {
    const endDate = addDays_(new Date(r.dateFin + 'T00:00:00'), 1);
    const end     = Utilities.formatDate(endDate, APP.TZ, 'yyyy-MM-dd');
    const color   = statusColor_(r.statut);
    return {
      id: r.idDemande,
      title: `${r.nomAgent} • ${r.typeConge} (${r.statut})`,
      start: r.dateDebut,
      end,
      allDay: true,
      backgroundColor: color,
      borderColor: color,
      extendedProps: r
    };
  });

  return { ok: true, events };
}

function apiGetMyBalance(token) {
  const session = requireSession_(token);
  const user    = getUserByIdRaw_(session.userId);
  applyMonthlyAccrualForUserYear_(user.userId, currentYear_());
  const balance = computeBalanceWithProvisional_(user.userId, currentYear_());
  return Object.assign({ ok: true }, balance);
}

function apiGetDashboard(token) {
  const session = requireSession_(token);
  requireAdmin_(session);
  applyMonthlyAccrualForAll_();
  const reqs        = getRowsAsObjects_(APP.SHEETS.REQUESTS).map(r => normalizeRequestRow_(r));
  const currentYear = currentYear_();
  const thisYear    = reqs.filter(r => r.dateDebut && new Date(r.dateDebut).getFullYear() === currentYear);
  const pending     = thisYear.filter(r => r.statut === APP.STATUS.PENDING).length;
  const approved    = thisYear.filter(r => r.statut === APP.STATUS.APPROVED).length;
  const rejected    = thisYear.filter(r => r.statut === APP.STATUS.REJECTED).length;
  let totalHours = 0, countProcessed = 0;
  thisYear.forEach(r => {
    if ((r.statut === APP.STATUS.APPROVED || r.statut === APP.STATUS.REJECTED) && r.dateCreation && r.dateTraitement) {
      const h = (new Date(r.dateTraitement) - new Date(r.dateCreation)) / 3600000;
      if (!isNaN(h)) { totalHours += h; countProcessed++; }
    }
  });
  const byType = {}, byTeamPending = {};
  thisYear.forEach(r => { byType[r.typeConge] = (byType[r.typeConge] || 0) + 1; });
  thisYear.filter(r => r.statut === APP.STATUS.PENDING).forEach(r => {
    byTeamPending[r.equipe || 'Sans équipe'] = (byTeamPending[r.equipe || 'Sans équipe'] || 0) + 1;
  });
  return {
    ok: true,
    cards: {
      pending, approved, rejected,
      avgProcessingHours: countProcessed ? Number((totalHours / countProcessed).toFixed(1)) : 0
    },
    byType: Object.keys(byType).map(k => ({ type: k, count: byType[k] })),
    byTeamPending: Object.keys(byTeamPending).map(k => ({ equipe: k, count: byTeamPending[k] }))
  };
}

function apiListHistory(token, filters) {
  const session = requireSession_(token);
  const user    = getUserByIdRaw_(session.userId);
  const isAdmin = user && user.role === APP.ROLES.ADMIN;
  filters = filters || {};

  const rows = getRowsAsObjects_(APP.SHEETS.HISTORY).map(r => ({
    idLog:        String(r.idLog        || ''),
    idDemande:    String(r.idDemande    || ''),
    action:       String(r.action       || ''),
    auteurUserId: String(r.auteurUserId || ''),
    auteurRole:   String(r.auteurRole   || ''),
    cibleUserId:  String(r.cibleUserId  || ''),
    cibleEquipe:  String(r.cibleEquipe  || ''),
    dateHeure:    toIsoDateTime_(r.dateHeure),
    details:      String(r.details      || '')
  }));

  let out = rows.filter(r => {
    if (isAdmin) {
      if (filters.action && r.action !== filters.action) return false;
      if (filters.userId && r.cibleUserId !== filters.userId && r.auteurUserId !== filters.userId) return false;
      if (filters.equipe && r.cibleEquipe !== filters.equipe) return false;
      return true;
    }
    const sameGroup   = r.cibleEquipe && user.equipe && r.cibleEquipe === user.equipe;
    const selfRelated = r.cibleUserId === user.userId || r.auteurUserId === user.userId;
    return !!sameGroup || !!selfRelated;
  });

  if (filters.annee) {
    const fYear = Number(filters.annee);
    out = out.filter(r => r.dateHeure && new Date(r.dateHeure).getFullYear() === fYear);
  }

  out.sort((a, b) => new Date(b.dateHeure).getTime() - new Date(a.dateHeure).getTime());

  const total    = out.length;
  const page     = Math.max(1, Number(filters.page || 1));
  const pageSize = Math.max(5, Math.min(200, Number(filters.pageSize || 50)));
  out = out.slice((page - 1) * pageSize, page * pageSize);

  return { ok: true, items: out, total, page, pageSize };
}

/* ========= ADMIN USERS / BALANCES / PARAMS / HOLIDAYS / RULES ========= */

function apiAdminListUsers(token) {
  const session = requireSession_(token);
  requireAdmin_(session);
  applyMonthlyAccrualForAll_();
  const users = getRowsAsObjects_(APP.SHEETS.USERS)
    .map(u => ({
      userId:       normTrim_(u.userId),
      nomComplet:   normTrim_(u.nomComplet),
      role:         normTrim_(u.role),
      equipe:       normTrim_(u.equipe),
      actif:        normTrim_(u.actif || 'oui'),
      dateCreation: toIsoDateTime_(u.dateCreation),
      dernierLogin: toIsoDateTime_(u.dernierLogin)
    }))
    .sort((a, b) => a.userId.localeCompare(b.userId, 'fr'));

  const year     = currentYear_();
  const balances = getRowsAsObjects_(APP.SHEETS.BALANCES).filter(b => Number(b.annee) === year);
  const byUser   = {};
  balances.forEach(b => byUser[normUpper_(b.userId)] = b);

  users.forEach(u => {
    const b = byUser[normUpper_(u.userId)] || getOrCreateBalance_(u.userId, year);
    u.balance = {
      annee: year,
      soldeInitial:             num_(b.soldeInitial),
      congesPris:               num_(b.congesPris),
      soldeRestant:             num_(b.soldeRestant),
      soldeNonConsommeInitial:  num_(b.soldeNonConsommeInitial),
      congesNonConsommesPris:   num_(b.congesNonConsommesPris),
      soldeNonConsommeRestant:  num_(b.soldeNonConsommeRestant)
    };
  });
  return { ok: true, users };
}

function apiAdminSaveUser(token, payload) {
  const session = requireSession_(token);
  requireAdmin_(session);
  payload = payload || {};
  const userId     = sanitizeText_(payload.userId).toUpperCase();
  const nomComplet = sanitizeText_(payload.nomComplet);
  const role       = sanitizeText_(payload.role);
  const equipe     = sanitizeText_(payload.equipe || '');
  const actif      = String(payload.actif || 'oui').toLowerCase() === 'non' ? 'non' : 'oui';
  const password   = String(payload.password || '');
  if (!userId) throw new Error('ID requis.');
  if (!nomComplet) throw new Error('Nom requis.');
  if ([APP.ROLES.AGENT, APP.ROLES.ADMIN].indexOf(role) === -1) throw new Error('Rôle invalide.');
  const existing = getUserByIdRaw_(userId);
  const isNew    = !existing;
  if (isNew && !password) throw new Error('Mot de passe requis pour création.');
  if (!isNew && password && password.length < 6) throw new Error('Mot de passe min 6 caractères.');
  if (isNew && password.length < 6) throw new Error('Mot de passe min 6 caractères.');
  createOrUpdateUser_({ userId, nomComplet, role, equipe, actif, password: password || null });
  if (role === APP.ROLES.AGENT) getOrCreateBalance_(userId, currentYear_());
  logHistory_('', isNew ? 'creation_user_admin' : 'maj_user_admin',
    session.userId, 'admin',
    `${isNew ? 'Création' : 'Mise à jour'} utilisateur ${userId} (${role})`,
    { cibleUserId: userId, cibleEquipe: equipe });
  return { ok: true };
}

function apiAdminResetPasswordsBulk(token, payload) {
  const session = requireSession_(token);
  requireAdmin_(session);
  payload = payload || {};
  const userIds     = Array.isArray(payload.userIds) ? payload.userIds : [];
  const newPassword = payload.newPassword != null ? String(payload.newPassword) : '';
  const ids = userIds.map(normUpper_).filter(Boolean);
  if (!ids.length) throw new Error('Aucun utilisateur sélectionné.');
  const useCommon = newPassword.trim().length > 0;
  const commonPwd = newPassword.trim();
  if (useCommon && commonPwd.length < 6) throw new Error('Mot de passe min 6 caractères.');

  const results = [];
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ids.forEach(id => {
      try {
        const u = getUserByIdRaw_(id);
        if (!u) throw new Error('Utilisateur introuvable.');
        const pwd = useCommon ? commonPwd : generateTempPassword_(10);
        setUserPassword_(id, pwd);
        results.push({ userId: id, ok: true, tempPassword: pwd });
      } catch(e) {
        results.push({ userId: id, ok: false, error: (e && e.message) ? e.message : String(e) });
      }
    });
    const okCount = results.filter(r => r.ok).length;
    logHistory_('', 'reset_mdp_bulk', session.userId, 'admin',
      `Reset MDP en masse: ${okCount}/${results.length} utilisateur(s).`,
      { cibleUserId: '', cibleEquipe: '' });
    return { ok: true, items: results, commonPasswordUsed: useCommon };
  } finally {
    lock.releaseLock();
  }
}

function generateTempPassword_(len) {
  len = Number(len || 10);
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnopqrstuvwxyz23456789@#$%';
  let out = '';
  for (let i = 0; i < len; i++) out += chars.charAt(Math.floor(Math.random() * chars.length));
  return out;
}

function setUserPassword_(userId, newPassword) {
  const pwd = String(newPassword || '');
  if (pwd.length < 6) throw new Error('Mot de passe min 6 caractères.');
  const u = getUserByIdRaw_(userId);
  if (!u) throw new Error('Utilisateur introuvable.');
  const salt = generateSalt_();
  updateUserFields_(u.userId, {
    salt, passwordHash: hashPassword_(pwd, salt),
    tentativesEchouees: 0, compteBloqueJusqua: ''
  });
}

function apiAdminSetBalance(token, payload) {
  const session = requireSession_(token);
  requireAdmin_(session);
  payload = payload || {};
  const userId    = normUpper_(payload.userId);
  const annee     = Number(payload.annee || currentYear_());
  const hasNormal = payload.soldeInitial != null && String(payload.soldeInitial) !== '';
  const hasNC     = payload.soldeNonConsommeInitial != null && String(payload.soldeNonConsommeInitial) !== '';
  if (!userId) throw new Error('User ID requis.');
  if (!hasNormal && !hasNC) throw new Error('Aucun solde à mettre à jour.');
  const soldeInitial   = hasNormal ? parseLocaleNumber_(payload.soldeInitial) : null;
  const soldeNCInitial = hasNC ? parseLocaleNumber_(payload.soldeNonConsommeInitial) : null;
  if (hasNormal && (isNaN(soldeInitial) || soldeInitial < 0)) throw new Error('Solde initial normal invalide.');
  if (hasNC && (isNaN(soldeNCInitial) || soldeNCInitial < 0)) throw new Error('Solde non consommé initial invalide.');

  const sheet   = getSheet_(APP.SHEETS.BALANCES);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const h       = headerIndexMap_(headers);
  if (h.soldeNonConsommeInitial === undefined)
    throw new Error('Colonnes soldes non consommés manquantes. Lance generateSheetsAndHeaders() puis recharge.');

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (normUpper_(data[i][h.userId]) === userId && Number(data[i][h.annee]) === annee) {
      rowIndex = i + 1; break;
    }
  }

  if (rowIndex === -1) {
    appendRowObject_(APP.SHEETS.BALANCES, {
      userId, annee,
      soldeInitial: hasNormal ? soldeInitial : APP.DEFAULT_ANNUAL_BALANCE,
      congesPris: 0,
      soldeRestant: hasNormal ? soldeInitial : APP.DEFAULT_ANNUAL_BALANCE,
      soldeNonConsommeInitial: hasNC ? soldeNCInitial : 0,
      congesNonConsommesPris: 0,
      soldeNonConsommeRestant: hasNC ? soldeNCInitial : 0,
      majLe: new Date(),
      dernierMoisCredite: prevMonthKey_()
    });
  } else {
    const row = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    if (hasNormal) row[h.soldeInitial] = soldeInitial;
    if (hasNC) row[h.soldeNonConsommeInitial] = soldeNCInitial;
    row[h.majLe] = new Date();
    if (h.dernierMoisCredite !== undefined && !row[h.dernierMoisCredite])
      row[h.dernierMoisCredite] = prevMonthKey_();
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  }

  recomputeBalance_(userId, annee);
  recomputeNonConsumedBalance_(userId, annee);

  const target = getUserByIdRaw_(userId);
  logHistory_('', 'maj_solde_admin', session.userId, 'admin',
    `Soldes mis à jour ${annee} | normal=${hasNormal ? soldeInitial : '(inchangé)'} | nonConsommé=${hasNC ? soldeNCInitial : '(inchangé)'}`,
    { cibleUserId: userId, cibleEquipe: target ? String(target.equipe || '') : '' });
  return { ok: true };
}

function apiAdminGetParams(token) {
  const session = requireSession_(token);
  requireAdmin_(session);
  return {
    ok: true,
    params: {
      maxAbsentsParEquipe:  getParam_('maxAbsentsParEquipe', '3'),
      delaiMinAvantConge:   getParam_('delaiMinAvantConge', '0'),
      typesCongesAutorises: getParam_('typesCongesAutorises', APP.LEAVE_TYPES.join(',')),
      vueEquipeAgents:      getParam_('vueEquipeAgents', 'oui'),
      accrualMensuelJours:  getParam_('accrualMensuelJours', normalizeNumberString_(APP.MONTHLY_ACCRUAL_DAYS)),
      modeleMailConge:      getParam_('modeleMailConge', DEFAULT_MAIL_CONGE_TPL_)
    }
  };
}

function apiAdminSaveParams(token, payload) {
  const session = requireSession_(token);
  requireAdmin_(session);
  payload = payload || {};
  setParam_('maxAbsentsParEquipe', normalizeNumberString_(payload.maxAbsentsParEquipe || '3'));
  setParam_('delaiMinAvantConge', normalizeNumberString_(payload.delaiMinAvantConge || '0'));
  setParam_('vueEquipeAgents', String(payload.vueEquipeAgents || 'oui'));
  setParam_('accrualMensuelJours', normalizeNumberString_(payload.accrualMensuelJours || APP.MONTHLY_ACCRUAL_DAYS));
  if (payload.typesCongesAutorises)
    setParam_('typesCongesAutorises', String(payload.typesCongesAutorises));
  if (payload.modeleMailConge !== undefined)
    setParam_('modeleMailConge', String(payload.modeleMailConge || ''));
  logHistory_('', 'maj_parametres_admin', session.userId, 'admin', 'Paramètres mis à jour', {});
  return { ok: true };
}

function apiAdminListHolidays(token) {
  const session = requireSession_(token);
  requireAdmin_(session);
  const rows = getRowsAsObjects_(APP.SHEETS.HOLIDAYS)
    .map(r => ({ date: toIsoDate_(r.date), libelle: String(r.libelle || '') }))
    .sort((a, b) => (a.date || '').localeCompare(b.date || ''));
  return { ok: true, items: rows };
}

function apiAdminSaveHoliday(token, payload) {
  const session = requireSession_(token);
  requireAdmin_(session);
  payload = payload || {};
  const date    = toIsoDateString_(payload.date);
  const libelle = sanitizeText_(payload.libelle || '');
  if (!date) throw new Error('Date invalide.');
  if (!libelle) throw new Error('Libellé requis.');

  const sheet   = getSheet_(APP.SHEETS.HOLIDAYS);
  const data    = sheet.getDataRange().getValues();
  const headers = data[0];
  const h       = headerIndexMap_(headers);

  for (let i = 1; i < data.length; i++) {
    if (toIsoDate_(data[i][h.date]) === date) {
      data[i][h.libelle] = libelle;
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([data[i]]);
      clearHolidayCache_();
      logHistory_('', 'maj_ferie_admin', session.userId, 'admin', `Jour férié ${date} = ${libelle}`, {});
      return { ok: true };
    }
  }
  appendRowObject_(APP.SHEETS.HOLIDAYS, { date, libelle });
  clearHolidayCache_();
  logHistory_('', 'creation_ferie_admin', session.userId, 'admin', `Jour férié ${date} = ${libelle}`, {});
  return { ok: true };
}

function apiAdminDeleteHoliday(token, dateIso) {
  const session = requireSession_(token);
  requireAdmin_(session);
  const target = toIsoDateString_(dateIso);
  if (!target) throw new Error('Date invalide.');
  const sh = getSheet_(APP.SHEETS.HOLIDAYS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { ok: true };
  const h = headerIndexMap_(values[0]);
  for (let i = values.length - 1; i >= 1; i--) {
    if (toIsoDate_(values[i][h.date]) === target) sh.deleteRow(i + 1);
  }
  clearHolidayCache_();
  logHistory_('', 'suppression_ferie_admin', session.userId, 'admin', `Suppression jour férié ${target}`, {});
  return { ok: true };
}

function apiAdminListTeamRules(token) {
  const session = requireSession_(token);
  requireAdmin_(session);
  return { ok: true, items: getTeamRules_() };
}

function apiAdminSaveTeamRule(token, payload) {
  const session = requireSession_(token);
  requireAdmin_(session);
  payload = payload || {};

  const equipe = sanitizeText_(payload.equipe || '').toUpperCase();
  if (!equipe) throw new Error('Équipe requise.');

  const maxDemandesParJour = emptyToBlankOrNumber_(payload.maxDemandesParJour);
  const delaiMinAvantCongeJours = emptyToBlankOrNumber_(payload.delaiMinAvantCongeJours);
  const maxAbsentsParJour = emptyToBlankOrNumber_(payload.maxAbsentsParJour);
  const heureLimiteDepot = sanitizeText_(payload.heureLimiteDepot || '');
  const actif = String(payload.actif || 'oui').toLowerCase() === 'non' ? 'non' : 'oui';

  const checks = {
    maxDemandesParJour: maxDemandesParJour,
    delaiMinAvantCongeJours: delaiMinAvantCongeJours,
    maxAbsentsParJour: maxAbsentsParJour
  };
  Object.keys(checks).forEach(function(field) {
    const val = checks[field];
    if (val !== '' && (isNaN(val) || Number(val) < 0)) throw new Error('Valeur invalide pour ' + field + '.');
  });
  if (heureLimiteDepot && !/^\d{2}:\d{2}$/.test(heureLimiteDepot)) throw new Error('Heure limite invalide.');

  const sh = getSheet_(APP.SHEETS.TEAM_RULES);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const h = headerIndexMap_(headers);

  let rowIdx = -1;
  for (let i = 1; i < values.length; i++) {
    if (normUpper_(values[i][h.equipe]) === equipe) { rowIdx = i + 1; break; }
  }

  if (rowIdx === -1) {
    appendRowObject_(APP.SHEETS.TEAM_RULES, {
      equipe,
      maxDemandesParJour: maxDemandesParJour === '' ? '' : maxDemandesParJour,
      delaiMinAvantCongeJours: delaiMinAvantCongeJours === '' ? '' : delaiMinAvantCongeJours,
      maxAbsentsParJour: maxAbsentsParJour === '' ? '' : maxAbsentsParJour,
      heureLimiteDepot,
      actif,
      majLe: new Date()
    });
  } else {
    const row = sh.getRange(rowIdx, 1, 1, headers.length).getValues()[0];
    row[h.maxDemandesParJour] = maxDemandesParJour === '' ? '' : maxDemandesParJour;
    row[h.delaiMinAvantCongeJours] = delaiMinAvantCongeJours === '' ? '' : delaiMinAvantCongeJours;
    row[h.maxAbsentsParJour] = maxAbsentsParJour === '' ? '' : maxAbsentsParJour;
    row[h.heureLimiteDepot] = heureLimiteDepot;
    row[h.actif] = actif;
    row[h.majLe] = new Date();
    sh.getRange(rowIdx, 1, 1, headers.length).setValues([row]);
  }

  clearTeamRulesCache_();
  logHistory_('', 'maj_regle_equipe_admin', session.userId, 'admin',
    `Règle équipe ${equipe} | maxDépôt/j=${maxDemandesParJour || '-'} | délai=${delaiMinAvantCongeJours || '-'} | maxAbs=${maxAbsentsParJour || '-'} | heureLimite=${heureLimiteDepot || '-'}`,
    { cibleUserId: '', cibleEquipe: equipe });
  return { ok: true };
}

function apiAdminDeleteTeamRule(token, equipe) {
  const session = requireSession_(token);
  requireAdmin_(session);
  const team = normUpper_(equipe);
  if (!team) throw new Error('Équipe requise.');
  const sh = getSheet_(APP.SHEETS.TEAM_RULES);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { ok: true };
  const h = headerIndexMap_(values[0]);
  for (let i = values.length - 1; i >= 1; i--) {
    if (normUpper_(values[i][h.equipe]) === team) sh.deleteRow(i + 1);
  }
  clearTeamRulesCache_();
  logHistory_('', 'suppression_regle_equipe_admin', session.userId, 'admin',
    `Suppression règle équipe ${team}`, { cibleUserId: '', cibleEquipe: team });
  return { ok: true };
}

/* ========= CREATION DEMANDE ========= */

function createLeaveRequestForUser_(targetUser, actorUser, payload) {
  if (!targetUser) throw new Error('Agent cible introuvable.');
  if (String(targetUser.actif || '').toLowerCase() !== 'oui') throw new Error('Agent cible inactif.');

  payload = payload || {};
  const typeConge      = sanitizeText_(payload.typeConge);
  const dateDebut      = toIsoDateString_(payload.dateDebut);
  const dateFin        = toIsoDateString_(payload.dateFin);
  const demiJournee    = sanitizeText_(payload.demiJournee || 'Aucune');
  const motif          = sanitizeText_(payload.motif || '');
  const statutInitial  = sanitizeText_(payload.statutInitial || APP.STATUS.PENDING);
  const modeValRaw     = normTrim_(payload.modeValidation || 'normal').toLowerCase().replace(/[\s_-]/g,'');
  const useNonConsumed = (modeValRaw === 'nonconsomme') && typeConsumesBalance_(typeConge);

  const allowedTypes = getAllowedLeaveTypes_();
  if (!typeConge || allowedTypes.indexOf(typeConge) === -1) throw new Error('Type de congé invalide.');
  if (!dateDebut || !dateFin) throw new Error('Dates invalides.');
  if (new Date(dateDebut).getTime() > new Date(dateFin).getTime())
    throw new Error('La date de début doit être <= date de fin.');
  if ([APP.STATUS.PENDING, APP.STATUS.APPROVED].indexOf(statutInitial) === -1)
    throw new Error('Statut initial invalide.');

  const equipe = String(targetUser.equipe || '');
  if (actorUser.role === APP.ROLES.AGENT) {
    validateTeamSubmissionRules_(equipe, dateDebut);
  }

  const nbJours = calculateLeaveDays_(dateDebut, dateFin, demiJournee);
  if (nbJours <= 0) throw new Error('Nombre de jours invalide (week-end / jours fériés).');

  validateNoOverlap_(targetUser.userId, dateDebut, dateFin, null);
  if (statutInitial === APP.STATUS.APPROVED)
    validateTeamQuota_(equipe, dateDebut, dateFin, null);

  const annee = new Date(dateDebut).getFullYear();
  applyMonthlyAccrualForUserYear_(targetUser.userId, annee);

  if (typeConsumesBalance_(typeConge)) {
    const bal = getOrCreateBalance_(targetUser.userId, annee);
    if (statutInitial === APP.STATUS.APPROVED) {
      if (useNonConsumed) {
        if (num_(bal.soldeNonConsommeRestant) < nbJours)
          throw new Error('Solde non consommé insuffisant (' + bal.soldeNonConsommeRestant + ' j restant(s)).');
      } else {
        if (num_(bal.soldeRestant) < nbJours)
          throw new Error('Solde insuffisant (' + bal.soldeRestant + ' jour(s) restant(s)).');
      }
    } else if (actorUser.role !== APP.ROLES.ADMIN) {
      if (num_(bal.soldeRestant) < nbJours)
        throw new Error('Solde insuffisant (' + bal.soldeRestant + ' jour(s) restant(s)).');
    }
  }

  const imputationSoldeInitiale = (statutInitial === APP.STATUS.APPROVED && typeConsumesBalance_(typeConge))
    ? (useNonConsumed ? 'nonConsomme' : 'normal')
    : '';

  const idDemande = 'DC-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  const now = new Date();
  const isAdminApproved = actorUser.role === APP.ROLES.ADMIN && statutInitial === APP.STATUS.APPROVED;

  appendRowObject_(APP.SHEETS.REQUESTS, {
    idDemande,
    userId: targetUser.userId,
    nomAgent: targetUser.nomComplet,
    equipe,
    typeConge,
    dateDebut,
    dateFin,
    demiJournee,
    nbJours,
    imputationSolde: imputationSoldeInitiale,
    motif,
    statut: statutInitial,
    commentaireAdmin: isAdminApproved ? 'Créée et validée par admin' : '',
    dateCreation: now,
    dateTraitement: isAdminApproved ? now : '',
    adminTraitant: isAdminApproved ? actorUser.userId : ''
  });

  if (statutInitial === APP.STATUS.APPROVED && typeConsumesBalance_(typeConge)) {
    recomputeBalance_(targetUser.userId, annee);
    recomputeNonConsumedBalance_(targetUser.userId, annee);
  }

  logHistory_(idDemande,
    actorUser.role === APP.ROLES.ADMIN ? 'création_admin' : 'création',
    actorUser.userId, actorUser.role,
    `${typeConge} ${dateDebut} → ${dateFin} (${nbJours}j)${motif ? ' | ' + motif : ''}`,
    { cibleUserId: targetUser.userId, cibleEquipe: equipe });

  return { ok: true, idDemande };
}

/* ========= CORE LOGIC ========= */

function validateNoOverlap_(userId, dateDebut, dateFin, ignoreIdDemande) {
  const userNorm = normUpper_(userId);
  const ignoreNorm = normTrim_(ignoreIdDemande);
  const reqs = getRowsAsObjects_(APP.SHEETS.REQUESTS).map(normalizeRequestRow_);
  const conflict = reqs.find(r => {
    if (normUpper_(r.userId) !== userNorm) return false;
    if (ignoreNorm && normTrim_(r.idDemande) === ignoreNorm) return false;
    const st = normTrim_(r.statut);
    if (st === normTrim_(APP.STATUS.REJECTED) || st === normTrim_(APP.STATUS.CANCELED)) return false;
    return rangesOverlap_(r.dateDebut, r.dateFin, dateDebut, dateFin);
  });
  if (conflict) throw new Error('Conflit avec une autre demande (' + conflict.idDemande + ').');
}

function validateTeamQuota_(equipe, dateDebut, dateFin, ignoreIdDemande) {
  const equipeNorm = normUpper_(equipe);
  if (!equipeNorm) return;
  const teamRule = getTeamRule_(equipeNorm);
  const maxAbs = teamRule.maxAbsentsParJour !== ''
    ? Number(teamRule.maxAbsentsParJour)
    : Number(getParam_('maxAbsentsParEquipe', '3'));
  if (!maxAbs || maxAbs <= 0) return;

  const ignoreNorm = normTrim_(ignoreIdDemande);
  const reqs = getRowsAsObjects_(APP.SHEETS.REQUESTS)
    .map(normalizeRequestRow_)
    .filter(r =>
      normUpper_(r.equipe) === equipeNorm &&
      normTrim_(r.statut) === normTrim_(APP.STATUS.APPROVED) &&
      (!ignoreNorm || normTrim_(r.idDemande) !== ignoreNorm)
    );
  const days = enumerateDates_(dateDebut, dateFin);
  const hset = getHolidaySet_();
  for (var i = 0; i < days.length; i++) {
    const d = days[i];
    const iso = toIsoDate_(d);
    if (isWeekend_(d) || hset[iso]) continue;
    let count = 0;
    reqs.forEach(r => { if (rangesOverlap_(r.dateDebut, r.dateFin, iso, iso)) count++; });
    if (count >= maxAbs)
      throw new Error('Quota équipe atteint le ' + iso + ' (' + normTrim_(equipe) + ').');
  }
}

function validateTeamSubmissionRules_(equipe, dateDebut) {
  const rule = getTeamRule_(equipe);
  const todayIso = currentDateIso_();
  const ruleDelay = rule.delaiMinAvantCongeJours !== '' ? Number(rule.delaiMinAvantCongeJours) : null;
  const globalDelay = parseLocaleNumber_(getParam_('delaiMinAvantConge', '0'));
  const delayMin = ruleDelay != null && !isNaN(ruleDelay) ? ruleDelay : globalDelay;

  if (delayMin > 0) {
    const today  = stripTime_(new Date(todayIso + 'T00:00:00'));
    const start  = stripTime_(new Date(dateDebut + 'T00:00:00'));
    const diffDays = Math.floor((start - today) / 86400000);
    if (diffDays < delayMin)
      throw new Error('Pour cette équipe, le délai minimum avant congé est de ' + delayMin + ' jour(s).');
  }

  const maxDemandes = rule.maxDemandesParJour !== '' ? Number(rule.maxDemandesParJour) : 0;
  if (maxDemandes > 0) {
    const todayCount = getRowsAsObjects_(APP.SHEETS.REQUESTS)
      .map(normalizeRequestRow_)
      .filter(r =>
        normUpper_(r.equipe) === normUpper_(equipe) &&
        toIsoDateString_(r.dateCreation) === todayIso &&
        r.statut !== APP.STATUS.REJECTED &&
        r.statut !== APP.STATUS.CANCELED
      ).length;
    if (todayCount >= maxDemandes) {
      throw new Error('Pour cette équipe, la limite de dépôt quotidienne est atteinte (' + maxDemandes + ' demande(s) / jour).');
    }
  }

  if (rule.heureLimiteDepot) {
    const nowHm = currentTimeHm_();
    const startIsToday = dateDebut === todayIso;
    if (startIsToday && nowHm > rule.heureLimiteDepot) {
      throw new Error('Pour cette équipe, l\'heure limite de dépôt le jour même est ' + rule.heureLimiteDepot + '.');
    }
  }
}

function calculateLeaveDays_(dateDebutIso, dateFinIso, demiJournee) {
  const hset = getHolidaySet_();
  const days = enumerateDates_(dateDebutIso, dateFinIso);
  let count = 0;
  days.forEach(d => { if (!isWeekend_(d) && !hset[toIsoDate_(d)]) count += 1; });
  const sameDay = dateDebutIso === dateFinIso;
  if (sameDay && demiJournee && demiJournee !== 'Aucune' && count >= 1) count = 0.5;
  return Number(count.toFixed(1));
}

function recomputeBalance_(userId, annee) {
  const userNorm = normUpper_(userId);
  const bal = getOrCreateBalance_(userId, annee);
  const reqs = getRowsAsObjects_(APP.SHEETS.REQUESTS).map(normalizeRequestRow_)
    .filter(r => {
      const impKey = normTrim_(r.imputationSolde).toLowerCase().replace(/[\s_-]/g,'');
      return (
        normUpper_(r.userId) === userNorm &&
        normTrim_(r.statut) === normTrim_(APP.STATUS.APPROVED) &&
        typeConsumesBalance_(r.typeConge) &&
        impKey !== 'nonconsomme' &&
        r.dateDebut && new Date(r.dateDebut).getFullYear() === Number(annee)
      );
    });
  let pris = 0;
  reqs.forEach(r => pris += num_(r.nbJours));
  const soldeRestant = Number((num_(bal.soldeInitial) - pris).toFixed(1));
  updateBalanceRow_(userId, annee, { congesPris: pris, soldeRestant, majLe: new Date() });
  return { congesPris: pris, soldeRestant };
}

function recomputeNonConsumedBalance_(userId, annee) {
  const userNorm = normUpper_(userId);
  const bal = getOrCreateBalance_(userId, annee);
  const reqs = getRowsAsObjects_(APP.SHEETS.REQUESTS).map(normalizeRequestRow_)
    .filter(r => {
      const impKey = normTrim_(r.imputationSolde).toLowerCase().replace(/[\s_-]/g,'');
      return (
        normUpper_(r.userId) === userNorm &&
        normTrim_(r.statut) === normTrim_(APP.STATUS.APPROVED) &&
        typeConsumesBalance_(r.typeConge) &&
        impKey === 'nonconsomme' &&
        r.dateDebut && new Date(r.dateDebut).getFullYear() === Number(annee)
      );
    });
  let prisNC = 0;
  reqs.forEach(r => prisNC += num_(r.nbJours));
  const soldeNCRestant = Number((num_(bal.soldeNonConsommeInitial) - prisNC).toFixed(1));
  updateBalanceRow_(userId, annee, {
    congesNonConsommesPris: prisNC, soldeNonConsommeRestant: soldeNCRestant, majLe: new Date()
  });
  return { congesNonConsommesPris: prisNC, soldeNonConsommeRestant: soldeNCRestant };
}

function typeConsumesBalance_(typeConge) {
  return String(typeConge || '').toLowerCase() === 'annuel';
}

function computeBalanceWithProvisional_(userId, year) {
  const bal = getOrCreateBalance_(userId, year);
  const pendingItems = getRowsAsObjects_(APP.SHEETS.REQUESTS)
    .map(normalizeRequestRow_)
    .filter(r =>
      normUpper_(r.userId) === normUpper_(userId) &&
      r.statut === APP.STATUS.PENDING &&
      typeConsumesBalance_(r.typeConge) &&
      r.dateDebut && new Date(r.dateDebut + 'T00:00:00').getFullYear() === Number(year)
    );
  let joursPendants = 0;
  pendingItems.forEach(r => joursPendants += num_(r.nbJours));
  return {
    annee: year,
    soldeInitial: num_(bal.soldeInitial),
    congesPris: num_(bal.congesPris),
    soldeRestant: num_(bal.soldeRestant),
    joursPendants: Number(joursPendants.toFixed(1)),
    soldeProvisionnel: Number((num_(bal.soldeRestant) - joursPendants).toFixed(1)),
    soldeNonConsommeInitial: num_(bal.soldeNonConsommeInitial),
    congesNonConsommesPris: num_(bal.congesNonConsommesPris),
    soldeNonConsommeRestant: num_(bal.soldeNonConsommeRestant)
  };
}

/* ========= CREDIT MENSUEL ========= */

function currentMonthKey_() {
  return Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM');
}
function prevMonthKey_() {
  const d = new Date(); d.setMonth(d.getMonth() - 1);
  return Utilities.formatDate(d, APP.TZ, 'yyyy-MM');
}
function nextMonthKey_(monthKey) {
  const parts = String(monthKey).split('-');
  const d = new Date(Number(parts[0]), Number(parts[1]) - 1, 1);
  d.setMonth(d.getMonth() + 1);
  return Utilities.formatDate(d, APP.TZ, 'yyyy-MM');
}

function initializeAccrualCursorForExistingBalances_() {
  const sh = getSheet_(APP.SHEETS.BALANCES);
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return;
  const headers = vals[0];
  const h = headerIndexMap_(headers);
  const prevKey = prevMonthKey_();
  for (let i = 1; i < vals.length; i++) {
    if (h.dernierMoisCredite !== undefined && !vals[i][h.dernierMoisCredite]) {
      vals[i][h.dernierMoisCredite] = prevKey;
      sh.getRange(i + 1, 1, 1, headers.length).setValues([vals[i]]);
    }
  }
}

function runMonthlyAccrualNow() {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const result = applyMonthlyAccrualForAll_();
    SpreadsheetApp.getUi().alert('Crédit mensuel terminé : ' + result.updated + ' solde(s) mis à jour.');
    return result;
  } finally { lock.releaseLock(); }
}

function installMonthlyAccrualTrigger() {
  const fn = 'runMonthlyAccrualNow';
  if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === fn))
    ScriptApp.newTrigger(fn).timeBased().onMonthDay(1).atHour(1).create();
  SpreadsheetApp.getUi().alert('Déclencheur mensuel installé (jour 1 à 01h).');
}

function applyMonthlyAccrualForAll_() {
  const users = getRowsAsObjects_(APP.SHEETS.USERS)
    .filter(u => String(u.actif || '').toLowerCase() === 'oui' && String(u.role || '') === APP.ROLES.AGENT);
  let updated = 0;
  const year = currentYear_();
  users.forEach(u => { if (applyMonthlyAccrualForUserYear_(String(u.userId || ''), year)) updated++; });
  return { ok: true, updated };
}

function applyMonthlyAccrualForUserYear_(userId, annee) {
  annee = Number(annee);
  if (!userId || annee !== currentYear_()) return false;
  const bal = getOrCreateBalance_(userId, annee);
  const cursor = String(bal.dernierMoisCredite || '');
  const currentKey = currentMonthKey_();
  if (!cursor) {
    updateBalanceRow_(userId, annee, { dernierMoisCredite: prevMonthKey_(), majLe: new Date() });
    return false;
  }
  let lastKey = cursor, changed = false;
  while (lastKey < currentKey) {
    const nextKey = nextMonthKey_(lastKey);
    creditOneMonth_(userId, annee, nextKey);
    lastKey = nextKey;
    changed = true;
  }
  return changed;
}

function creditOneMonth_(userId, annee, monthKey) {
  const amount = parseLocaleNumber_(getParam_('accrualMensuelJours', normalizeNumberString_(APP.MONTHLY_ACCRUAL_DAYS)));
  if (!amount || amount <= 0) return;
  const bal = getOrCreateBalance_(userId, annee);
  updateBalanceRow_(userId, annee, {
    soldeInitial: Number((num_(bal.soldeInitial) + amount).toFixed(1)),
    soldeRestant: Number((num_(bal.soldeRestant) + amount).toFixed(1)),
    majLe: new Date(),
    dernierMoisCredite: monthKey
  });
  const user = getUserByIdRaw_(userId);
  logHistory_('', 'crédit_mensuel', 'SYSTEM', 'system',
    `Crédit automatique ${amount} jour(s) pour ${monthKey}`,
    { cibleUserId: userId, cibleEquipe: user ? String(user.equipe || '') : '' });
}

/* ========= DATA ACCESS ========= */

function getDb_() {
  if (DB_SPREADSHEET_ID) return SpreadsheetApp.openById(DB_SPREADSHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}
function getSheet_(name) {
  const sh = getDb_().getSheetByName(name);
  if (!sh) throw new Error('Feuille introuvable: ' + name);
  return sh;
}
function getRowsAsObjects_(sheetName) {
  const sh = getSheet_(sheetName);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) rows.push(objectFromRow_(headers, values[i]));
  return rows;
}
function appendRowObject_(sheetName, obj) {
  const sh = getSheet_(sheetName);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = headers.map(h => safeCell_(obj[h] !== undefined ? obj[h] : ''));
  sh.appendRow(row);
}
function objectFromRow_(headers, row) {
  const o = {};
  headers.forEach((h, i) => o[h] = row[i]);
  return o;
}
function headerIndexMap_(headers) {
  const map = {};
  headers.forEach((h, i) => map[h] = i);
  return map;
}

/* ========= USERS / PASSWORDS ========= */

function getUserByIdRaw_(userId) {
  const target = normUpper_(userId);
  return getRowsAsObjects_(APP.SHEETS.USERS).find(u => normUpper_(u.userId) === target) || null;
}

function createOrUpdateUser_(payload) {
  const userId = String(payload.userId).trim().toUpperCase();
  const nomComplet = sanitizeText_(payload.nomComplet || '');
  const role = String(payload.role || APP.ROLES.AGENT);
  const equipe = sanitizeText_(payload.equipe || '');
  const actif = String(payload.actif || 'oui').toLowerCase() === 'non' ? 'non' : 'oui';
  const password = payload.password;

  const sh = getSheet_(APP.SHEETS.USERS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const h = headerIndexMap_(headers);
  let rowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    if (normUpper_(data[i][h.userId]) === normUpper_(userId)) { rowIndex = i + 1; break; }
  }

  let salt = '', hash = '';
  if (password) { salt = generateSalt_(); hash = hashPassword_(String(password), salt); }

  if (rowIndex === -1) {
    appendRowObject_(APP.SHEETS.USERS, {
      userId, nomComplet, role, equipe,
      passwordHash: hash || '', salt: salt || '', actif,
      dateCreation: new Date(), dernierLogin: '',
      tentativesEchouees: 0, compteBloqueJusqua: ''
    });
    return;
  }
  const row = sh.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
  row[h.nomComplet] = nomComplet;
  row[h.role] = role;
  row[h.equipe] = equipe;
  row[h.actif] = actif;
  if (password) {
    row[h.passwordHash] = hash;
    row[h.salt] = salt;
    row[h.tentativesEchouees] = 0;
    row[h.compteBloqueJusqua] = '';
  }
  sh.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
}

function updateUserField_(userId, field, value) { updateUserFields_(userId, { [field]: value }); }

function updateUserFields_(userId, patch) {
  const sh = getSheet_(APP.SHEETS.USERS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const h = headerIndexMap_(headers);
  for (let i = 1; i < data.length; i++) {
    if (normUpper_(data[i][h.userId]) === normUpper_(userId)) {
      Object.keys(patch).forEach(k => { if (h[k] !== undefined) data[i][h[k]] = patch[k]; });
      sh.getRange(i + 1, 1, 1, headers.length).setValues([data[i]]);
      return true;
    }
  }
  return false;
}

function publicUser_(u) {
  return {
    userId: String(u.userId || ''),
    nomComplet: String(u.nomComplet || ''),
    role: String(u.role || ''),
    equipe: String(u.equipe || ''),
    actif: String(u.actif || ''),
    dernierLogin: toIsoDateTime_(u.dernierLogin)
  };
}

function registerFailedLogin_(user) {
  const attempts = num_(user.tentativesEchouees) + 1;
  const patch = { tentativesEchouees: attempts };
  if (attempts >= APP.LOGIN_MAX_ATTEMPTS) {
    patch.compteBloqueJusqua = new Date(Date.now() + APP.LOGIN_BLOCK_MINUTES * 60000);
    patch.tentativesEchouees = 0;
  }
  updateUserFields_(user.userId, patch);
}
function resetFailedLogin_(user) {
  updateUserFields_(user.userId, { tentativesEchouees: 0, compteBloqueJusqua: '' });
}
function generateSalt_() {
  return Utilities.getUuid().replace(/-/g,'') + String(new Date().getTime());
}
function hashPassword_(password, salt) {
  const raw = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, salt + '|' + password, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(raw);
}
function verifyPassword_(password, salt, expectedHash) {
  if (!salt || !expectedHash) return false;
  return hashPassword_(password, salt) === String(expectedHash);
}

/* ========= SESSIONS ========= */

function createSession_(user) {
  purgeExpiredSessions_();
  const token = Utilities.getUuid() + '-' + Utilities.getUuid();
  const now = new Date();
  const expiresAt = new Date(now.getTime() + APP.SESSION_HOURS * 3600000);
  appendRowObject_(APP.SHEETS.SESSIONS, {
    token, userId: user.userId, role: user.role,
    createdAt: now, expiresAt, isActive: 'oui', lastSeen: now
  });
  return { token, expiresAt };
}

function getSessionByToken_(token) {
  if (!token) return null;
  const sessions = getRowsAsObjects_(APP.SHEETS.SESSIONS);
  const now = new Date();
  const s = sessions.find(x =>
    String(x.token || '') === String(token) &&
    String(x.isActive || '').toLowerCase() === 'oui'
  );
  if (!s) return null;
  const exp = parseAnyDate_(s.expiresAt);
  if (!exp || exp.getTime() < now.getTime()) { invalidateSession_(token); return null; }
  return s;
}

function touchSession_(token) {
  const sh = getSheet_(APP.SHEETS.SESSIONS);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;
  const h = headerIndexMap_(data[0]);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][h.token]) === String(token)) {
      data[i][h.lastSeen] = new Date();
      sh.getRange(i + 1, 1, 1, data[0].length).setValues([data[i]]);
      return;
    }
  }
}

function invalidateSession_(token) {
  const sh = getSheet_(APP.SHEETS.SESSIONS);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return;
  const h = headerIndexMap_(data[0]);
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][h.token]) === String(token)) {
      data[i][h.isActive] = 'non';
      data[i][h.lastSeen] = new Date();
      sh.getRange(i + 1, 1, 1, data[0].length).setValues([data[i]]);
      return;
    }
  }
}

function purgeExpiredSessions() { return purgeExpiredSessions_(); }

function purgeExpiredSessions_() {
  const sh = getSheet_(APP.SHEETS.SESSIONS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { ok: true, deleted: 0 };
  const headers = values[0];
  const h = headerIndexMap_(headers);
  const now = new Date();
  let deleted = 0;
  for (let i = values.length - 1; i >= 1; i--) {
    const exp = parseAnyDate_(values[i][h.expiresAt]);
    const active = String(values[i][h.isActive] || '').toLowerCase() === 'oui';
    if (!exp || exp.getTime() < now.getTime() || !active) {
      sh.deleteRow(i + 1); deleted++;
    }
  }
  return { ok: true, deleted };
}

function requireSession_(token) {
  const s = getSessionByToken_(token);
  if (!s) throw new Error('Session expirée. Reconnecte-toi.');
  touchSession_(token);
  return s;
}
function requireAdmin_(session) {
  if (!session || String(session.role) !== APP.ROLES.ADMIN)
    throw new Error('Accès admin requis.');
  return getUserByIdRaw_(session.userId);
}

/* ========= BALANCES ========= */

function getOrCreateBalance_(userId, annee) {
  const userNorm = normUpper_(userId);
  const sh = getSheet_(APP.SHEETS.BALANCES);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const h = headerIndexMap_(headers);

  const required = ['soldeNonConsommeInitial','congesNonConsommesPris','soldeNonConsommeRestant'];
  const missing = required.filter(c => h[c] === undefined);
  if (missing.length)
    throw new Error('Colonnes soldes non consommés manquantes. Lance generateSheetsAndHeaders() puis recharge.');

  for (let i = 1; i < values.length; i++) {
    if (normUpper_(values[i][h.userId]) === userNorm && Number(values[i][h.annee]) === Number(annee)) {
      const obj = objectFromRow_(headers, values[i]);
      const patch = {};
      if (obj.soldeNonConsommeInitial == null || obj.soldeNonConsommeInitial === '') patch.soldeNonConsommeInitial = 0;
      if (obj.congesNonConsommesPris == null || obj.congesNonConsommesPris === '') patch.congesNonConsommesPris = 0;
      if (obj.soldeNonConsommeRestant == null || obj.soldeNonConsommeRestant === '') patch.soldeNonConsommeRestant = 0;
      if (Object.keys(patch).length) {
        patch.majLe = new Date();
        updateBalanceRow_(userId, annee, patch);
        Object.assign(obj, patch);
      }
      return obj;
    }
  }

  appendRowObject_(APP.SHEETS.BALANCES, {
    userId: normTrim_(userId), annee,
    soldeInitial: APP.DEFAULT_ANNUAL_BALANCE, congesPris: 0, soldeRestant: APP.DEFAULT_ANNUAL_BALANCE,
    soldeNonConsommeInitial: 0, congesNonConsommesPris: 0, soldeNonConsommeRestant: 0,
    majLe: new Date(), dernierMoisCredite: prevMonthKey_()
  });
  return {
    userId: normTrim_(userId), annee,
    soldeInitial: APP.DEFAULT_ANNUAL_BALANCE, congesPris: 0, soldeRestant: APP.DEFAULT_ANNUAL_BALANCE,
    soldeNonConsommeInitial: 0, congesNonConsommesPris: 0, soldeNonConsommeRestant: 0,
    majLe: new Date(), dernierMoisCredite: prevMonthKey_()
  };
}

function updateBalanceRow_(userId, annee, patch) {
  const userNorm = normUpper_(userId);
  const sh = getSheet_(APP.SHEETS.BALANCES);
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const h = headerIndexMap_(headers);
  for (let i = 1; i < values.length; i++) {
    if (normUpper_(values[i][h.userId]) === userNorm && Number(values[i][h.annee]) === Number(annee)) {
      Object.keys(patch).forEach(k => { if (h[k] !== undefined) values[i][h[k]] = patch[k]; });
      sh.getRange(i + 1, 1, 1, headers.length).setValues([values[i]]);
      return true;
    }
  }
  return false;
}

/* ========= PARAMS ========= */

function getParam_(key, defaultValue) {
  const rows = getRowsAsObjects_(APP.SHEETS.PARAMS);
  const p = rows.find(r => String(r.cle) === String(key));
  return p ? String(p.valeur) : String(defaultValue || '');
}
function setParam_(key, value) {
  const sh = getSheet_(APP.SHEETS.PARAMS);
  const vals = sh.getDataRange().getValues();
  const headers = vals[0];
  const h = headerIndexMap_(headers);
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][h.cle]) === String(key)) {
      vals[i][h.valeur] = String(value);
      sh.getRange(i + 1, 1, 1, headers.length).setValues([vals[i]]);
      return;
    }
  }
  appendRowObject_(APP.SHEETS.PARAMS, { cle: key, valeur: String(value) });
}
function setParamIfMissing_(key, value) {
  if (!getRowsAsObjects_(APP.SHEETS.PARAMS).find(r => String(r.cle) === String(key)))
    appendRowObject_(APP.SHEETS.PARAMS, { cle: key, valeur: String(value) });
}
function getAllowedLeaveTypes_() {
  const raw = getParam_('typesCongesAutorises', APP.LEAVE_TYPES.join(','));
  const arr = String(raw).split(',').map(s => s.trim()).filter(Boolean);
  return arr.length ? arr : APP.LEAVE_TYPES.slice();
}

/* ========= TEAM RULES ========= */

function getTeamRules_() {
  if (_teamRulesCache !== null) return _teamRulesCache;
  _teamRulesCache = getRowsAsObjects_(APP.SHEETS.TEAM_RULES)
    .map(normalizeTeamRuleRow_)
    .sort((a, b) => a.equipe.localeCompare(b.equipe, 'fr'));
  return _teamRulesCache;
}

function getTeamRule_(equipe) {
  const target = normUpper_(equipe);
  if (!target) return {
    equipe: '',
    maxDemandesParJour: '',
    delaiMinAvantCongeJours: '',
    maxAbsentsParJour: '',
    heureLimiteDepot: '',
    actif: 'oui'
  };
  const found = getTeamRules_().find(r => normUpper_(r.equipe) === target && String(r.actif || 'oui') !== 'non');
  return found || {
    equipe: target,
    maxDemandesParJour: '',
    delaiMinAvantCongeJours: '',
    maxAbsentsParJour: '',
    heureLimiteDepot: '',
    actif: 'oui'
  };
}

/* ========= HISTORY ========= */

function logHistory_(idDemande, action, auteurUserId, auteurRole, details, meta) {
  meta = meta || {};
  appendRowObject_(APP.SHEETS.HISTORY, {
    idLog: 'LOG-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
    idDemande: idDemande || '',
    action: action || '',
    auteurUserId: auteurUserId || '',
    auteurRole: auteurRole || '',
    cibleUserId: meta.cibleUserId || '',
    cibleEquipe: meta.cibleEquipe || '',
    dateHeure: new Date(),
    details: sanitizeText_(details || '')
  });
}

/* ========= HELPERS ========= */

function normalizeRequestRow_(r) {
  return {
    idDemande: String(r.idDemande || '').trim(),
    userId: String(r.userId || '').trim(),
    nomAgent: String(r.nomAgent || '').trim(),
    equipe: String(r.equipe || '').trim(),
    typeConge: String(r.typeConge || '').trim(),
    dateDebut: toIsoDate_(r.dateDebut),
    dateFin: toIsoDate_(r.dateFin),
    demiJournee: String(r.demiJournee || 'Aucune').trim(),
    nbJours: num_(r.nbJours),
    imputationSolde: String(r.imputationSolde || '').trim(),
    motif: String(r.motif || '').trim(),
    statut: String(r.statut || '').trim(),
    commentaireAdmin: String(r.commentaireAdmin || '').trim(),
    dateCreation: toIsoDateTime_(r.dateCreation),
    dateTraitement: toIsoDateTime_(r.dateTraitement),
    adminTraitant: String(r.adminTraitant || '').trim()
  };
}

function normalizeTeamRuleRow_(r) {
  return {
    equipe: String(r.equipe || '').trim(),
    maxDemandesParJour: valueOrBlankNumber_(r.maxDemandesParJour),
    delaiMinAvantCongeJours: valueOrBlankNumber_(r.delaiMinAvantCongeJours),
    maxAbsentsParJour: valueOrBlankNumber_(r.maxAbsentsParJour),
    heureLimiteDepot: String(r.heureLimiteDepot || '').trim(),
    actif: String(r.actif || 'oui').trim().toLowerCase() === 'non' ? 'non' : 'oui',
    majLe: toIsoDateTime_(r.majLe)
  };
}

function sanitizeText_(v) {
  v = String(v == null ? '' : v).trim();
  if (/^[=\-+@]/.test(v)) v = "'" + v;
  return v;
}
function safeCell_(v) {
  if (typeof v === 'string') return sanitizeText_(v);
  return v;
}
function num_(v) { const n = Number(v); return isNaN(n) ? 0 : n; }
function currentYear_() { return Number(Utilities.formatDate(new Date(), APP.TZ, 'yyyy')); }
function currentDateIso_() { return Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd'); }
function currentTimeHm_() { return Utilities.formatDate(new Date(), APP.TZ, 'HH:mm'); }

function parseAnyDate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) return v;
  const d = new Date(v);
  return isNaN(d) ? null : d;
}
function toIsoDate_(v) {
  const d = parseAnyDate_(v);
  if (!d) return '';
  return Utilities.formatDate(d, APP.TZ, 'yyyy-MM-dd');
}
function toIsoDateTime_(v) {
  const d = parseAnyDate_(v);
  if (!d) return '';
  return Utilities.formatDate(d, APP.TZ, 'yyyy-MM-dd HH:mm:ss');
}
function toIsoDateString_(v) {
  if (!v) return '';
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(v)) return v;
  return toIsoDate_(v);
}
function stripTime_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }
function addDays_(d, n) { const x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
function enumerateDates_(dateDebutIso, dateFinIso) {
  const start = new Date(dateDebutIso + 'T00:00:00');
  const end = new Date(dateFinIso + 'T00:00:00');
  const out = [];
  let cur = new Date(start.getTime());
  while (cur.getTime() <= end.getTime()) { out.push(new Date(cur.getTime())); cur = addDays_(cur, 1); }
  return out;
}
function isWeekend_(d) { const day = d.getDay(); return day === 0 || day === 6; }
function isHoliday_(d) { return !!getHolidaySet_()[toIsoDate_(d)]; }

function rangesOverlap_(aStart, aEnd, bStart, bEnd) {
  const a1 = new Date(aStart + 'T00:00:00').getTime();
  const a2 = new Date(aEnd + 'T23:59:59').getTime();
  const b1 = new Date(bStart + 'T00:00:00').getTime();
  const b2 = new Date(bEnd + 'T23:59:59').getTime();
  return a1 <= b2 && b1 <= a2;
}

function statusColor_(statut) {
  if (statut === APP.STATUS.PENDING) return '#f39c12';
  if (statut === APP.STATUS.APPROVED) return '#2ecc71';
  if (statut === APP.STATUS.REJECTED) return '#e74c3c';
  if (statut === APP.STATUS.CANCELED) return '#7f8c8d';
  return '#95a5a6';
}

function normUpper_(v) {
  return String(v == null ? '' : v).replace(/\u00A0/g,' ').trim().toUpperCase();
}
function normTrim_(v) {
  return String(v == null ? '' : v).replace(/\u00A0/g,' ').trim();
}

function parseLocaleNumber_(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  const s = String(v).trim().replace(/\s/g, '').replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? 0 : n;
}

function normalizeNumberString_(v) {
  return String(parseLocaleNumber_(v)).replace(',', '.');
}

function valueOrBlankNumber_(v) {
  if (v == null || v === '') return '';
  const s = String(v).trim().replace(',', '.');
  const n = Number(s);
  return isNaN(n) ? '' : n;
}

function emptyToBlankOrNumber_(v) {
  if (v == null || String(v).trim() === '') return '';
  const n = parseLocaleNumber_(v);
  return isNaN(n) ? '' : n;
}

function findLeaveById_(idDemande) {
  const id = normTrim_(idDemande);
  if (!id) return null;
  const req = getRowsAsObjects_(APP.SHEETS.REQUESTS)
    .map(normalizeRequestRow_)
    .find(r => normTrim_(r.idDemande) === id) || null;
  return req;
}

/* ========= MODÈLE MAIL VALIDATION CONGÉ ========= */

const DEFAULT_MAIL_CONGE_TPL_ =
`Objet : Congé - {{equipe}} - {{userId}} - {{nomAgent}}

Bonjour,

Merci d'enregistrer le congé de l'ambassadeur suivant :
Matricule : {{userId}}
Nom : {{nomAgent}}
Prénom : {{nomAgent}}
Date de départ : {{dateDebut}}
Date de retour : {{dateRetour}}

Cordialement,`;

function renderMailConge_(req, commentaire) {
  const tpl = getParam_('modeleMailConge', DEFAULT_MAIL_CONGE_TPL_);
  if (!tpl || !tpl.trim()) return null;

  // Calcul date de retour = dateFin + 1 jour
  let dateRetour = '';
  if (req.dateFin) {
    const d = new Date(req.dateFin + 'T00:00:00');
    d.setDate(d.getDate() + 1);
    dateRetour = Utilities.formatDate(d, APP.TZ, 'yyyy-MM-dd');
  }

  return tpl
    .replace(/\{\{nomAgent\}\}/g,        String(req.nomAgent        || ''))
    .replace(/\{\{userId\}\}/g,          String(req.userId          || ''))
    .replace(/\{\{typeConge\}\}/g,        String(req.typeConge       || ''))
    .replace(/\{\{dateDebut\}\}/g,        String(req.dateDebut       || ''))
    .replace(/\{\{dateFin\}\}/g,          String(req.dateFin         || ''))
    .replace(/\{\{dateRetour\}\}/g,       dateRetour)
    .replace(/\{\{nbJours\}\}/g,          String(req.nbJours         || ''))
    .replace(/\{\{equipe\}\}/g,           String(req.equipe          || ''))
    .replace(/\{\{commentaireAdmin\}\}/g, String(commentaire         || ''))
    .replace(/\{\{annee\}\}/g,
      req.dateDebut
        ? String(new Date(req.dateDebut + 'T00:00:00').getFullYear())
        : '');
}
