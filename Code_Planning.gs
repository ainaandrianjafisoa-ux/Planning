/*************************************
 * MODULE PLANNING + POINTAGE — V4.2
 * Code_Planning_v4_2.gs
 * Compatible avec Code.gs (module congés)
 *************************************/

const PLANNING_VERSION = 'V4.2.1';

const P_SH = {
  SHIFTS:   'PlanningShifts',
  MODELES:  'ShiftsModeles',
  PLAGES:   'PlagesHoraires',
  POINTAGE: 'Pointage',
  SWAPS:    'PlanningEchanges',
};

const P_SCHEMA = {
  PlanningShifts: [
    'idShift','userId','nomAgent','equipe',
    'date','heureDebut','heureFin',
    'pauseDebut','pauseFin','isOff','dureePrevue',
    'nomModele','couleurModele',
    'statut','notes','dateCreation','creePar'
  ],
  ShiftsModeles: [
    'idModele','nom','heureDebut','heureFin',
    'couleur','description','actif'
  ],
  PlagesHoraires: [
    'idPlage','equipe','nom','heureDebut','heureFin',
    'niveauFlux','joursActifs',
    'effectifMin','couleur','actif'
  ],
  Pointage: [
    'idPointage','userId','nomAgent','equipe','date',
    'heurePrevDebut','heurePrevFin',
    'heureArrivee','heureDepart',
    'dureePrevu','dureeReelle','heuresSup','retardMinutes',
    'statut','notes','valideAdmin','adminValidant','dateCreation'
  ],
  PlanningEchanges: [
    'idSwap','requesterUserId','requesterNom','requesterEquipe',
    'targetUserId','targetNom','targetEquipe','dateShift',
    'requesterShiftId','targetShiftId','requesterComment',
    'adminComment','statut','dateCreation','dateDecision','decidedBy'
  ],
};

/* ========= SETUP ========= */

function generatePlanningSheets() {
  const ss = getDb_();
  Object.keys(P_SCHEMA).forEach(name => ensureSheetWithHeaders_(ss, name, P_SCHEMA[name]));
  _addDefaultModeles_();
  _addDefaultPlages_();
  _formatPlanningSheets_();
  SpreadsheetApp.getActiveSpreadsheet()?.toast('Feuilles planning générées avec succès.');
  return { ok: true, message: 'Feuilles planning générées.' };
}

function _ensurePlanningReady_() {
  const ss = getDb_();
  Object.keys(P_SCHEMA).forEach(name => ensureSheetWithHeaders_(ss, name, P_SCHEMA[name]));
  return true;
}

function apiPlanningVersion() {
  return { ok: true, version: PLANNING_VERSION, stable: true, module: 'planning', generatedAt: toIsoDateTime_(new Date()) };
}

function apiAdminListPlanningAgents(token, filters) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  filters = filters || {};
  const equipe = normUpper_(filters.equipe || '');
  const search = normUpper_(filters.search || '');
  const date = toIsoDateString_(filters.date || '');
  const withStatus = String(filters.withStatus || '').toLowerCase() === 'oui';

  let users = _getActiveUsers_(equipe);
  if (search) {
    users = users.filter(u => normUpper_(u.userId).indexOf(search) !== -1 || normUpper_(u.nomAgent).indexOf(search) !== -1 || normUpper_(u.equipe).indexOf(search) !== -1);
  }

  if (withStatus && date) {
    users = users.map(u => {
      const shift = _getShiftForUserDate_(u.userId, date);
      const leave = _getApprovedLeaveForDate_(u.userId, date);
      return Object.assign({}, u, {
        hasShift: !!shift,
        hasLeave: !!leave,
        shiftLabel: shift ? ((shift.isOff === 'oui') ? 'OFF' : (shift.heureDebut + '–' + shift.heureFin)) : '',
        leaveLabel: leave ? String(leave.typeConge || '') : ''
      });
    });
  }

  return { ok: true, items: users, total: users.length };
}


function _formatPlanningSheets_() {
  const ss = getDb_();
  Object.keys(P_SCHEMA).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    sh.setFrozenRows(1);
    if (sh.getLastColumn() > 0) {
      sh.getRange(1, 1, 1, sh.getLastColumn())
        .setFontWeight('bold')
        .setBackground('#1f1f23')
        .setFontColor('#d4af37');
      sh.autoResizeColumns(1, sh.getLastColumn());
    }
  });
}

function _addDefaultModeles_() {
  try {
    const existing = getRowsAsObjects_(P_SH.MODELES);
    if (existing.length > 0) return;
    const defaults = [
      { nom:'Matin',        heureDebut:'07:00', heureFin:'15:00', couleur:'#2196F3', description:'Shift matin', actif:'oui' },
      { nom:'Après-midi',   heureDebut:'14:00', heureFin:'22:00', couleur:'#FF9800', description:'Shift après-midi', actif:'oui' },
      { nom:'Journée',      heureDebut:'08:00', heureFin:'17:00', couleur:'#4CAF50', description:'Journée complète', actif:'oui' },
      { nom:'Nuit',         heureDebut:'22:00', heureFin:'06:00', couleur:'#9C27B0', description:'Shift nuit', actif:'oui' },
      { nom:'Demi-matin',   heureDebut:'08:00', heureFin:'12:00', couleur:'#00BCD4', description:'Demi-journée matin', actif:'oui' },
      { nom:'Demi-soir',    heureDebut:'13:00', heureFin:'17:00', couleur:'#009688', description:'Demi-journée soir', actif:'oui' },
    ];
    defaults.forEach(m => {
      appendRowObject_(P_SH.MODELES, {
        idModele: 'MOD-' + Utilities.getUuid().slice(0, 6).toUpperCase(),
        nom: m.nom,
        heureDebut: m.heureDebut,
        heureFin: m.heureFin,
        couleur: m.couleur,
        description: m.description,
        actif: m.actif
      });
    });
  } catch (e) {}
}

function _addDefaultPlages_() {
  try {
    const existing = getRowsAsObjects_(P_SH.PLAGES);
    if (existing.length > 0) return;
    const defaults = [
      { equipe:'', nom:'Ouverture',  heureDebut:'08:00', heureFin:'10:00', niveauFlux:'Fort',      joursActifs:'1,2,3,4,5', effectifMin:3, couleur:'#FF5722', actif:'oui' },
      { equipe:'', nom:'Pic midi',   heureDebut:'12:00', heureFin:'14:00', niveauFlux:'Très fort', joursActifs:'1,2,3,4,5', effectifMin:4, couleur:'#F44336', actif:'oui' },
      { equipe:'', nom:'Fermeture',  heureDebut:'17:00', heureFin:'19:00', niveauFlux:'Fort',      joursActifs:'1,2,3,4,5', effectifMin:3, couleur:'#FF5722', actif:'oui' },
      { equipe:'', nom:'Matinée',    heureDebut:'10:00', heureFin:'12:00', niveauFlux:'Normal',    joursActifs:'1,2,3,4,5', effectifMin:2, couleur:'#4CAF50', actif:'oui' },
      { equipe:'', nom:'Après-midi', heureDebut:'14:00', heureFin:'17:00', niveauFlux:'Normal',    joursActifs:'1,2,3,4,5', effectifMin:2, couleur:'#4CAF50', actif:'oui' },
    ];
    defaults.forEach(p => {
      appendRowObject_(P_SH.PLAGES, {
        idPlage: 'PLG-' + Utilities.getUuid().slice(0, 6).toUpperCase(),
        equipe: p.equipe,
        nom: p.nom,
        heureDebut: p.heureDebut,
        heureFin: p.heureFin,
        niveauFlux: p.niveauFlux,
        joursActifs: p.joursActifs,
        effectifMin: p.effectifMin,
        couleur: p.couleur,
        actif: p.actif
      });
    });
  } catch (e) {}
}

/* ========= MODÈLES DE SHIFTS ========= */

function apiAdminListShiftModeles(token) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  const modeles = getRowsAsObjects_(P_SH.MODELES)
    .map(_normalizeModeleRow_)
    .filter(m => m.actif !== 'non')
    .sort((a, b) => a.nom.localeCompare(b.nom, 'fr'));
  return { ok: true, modeles };
}

function apiAdminSaveShiftModele(token, payload) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  payload = payload || {};

  const idModele = normTrim_(payload.idModele) || ('MOD-' + Utilities.getUuid().slice(0, 6).toUpperCase());
  const nom = sanitizeText_(payload.nom || '');
  const heureDebut = _normalizeTimeValue_(payload.heureDebut);
  const heureFin = _normalizeTimeValue_(payload.heureFin);
  const couleur = sanitizeText_(payload.couleur || '#4ea1ff');
  const description = sanitizeText_(payload.description || '');

  if (!nom) throw new Error('Nom du modèle requis.');
  if (!heureDebut || !heureFin) throw new Error('Heures début et fin requises.');

  const sh = getSheet_(P_SH.MODELES);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const h = headerIndexMap_(headers);

  for (let i = 1; i < data.length; i++) {
    if (normTrim_(data[i][h.idModele]) === idModele) {
      data[i][h.nom] = nom;
      data[i][h.heureDebut] = heureDebut;
      data[i][h.heureFin] = heureFin;
      data[i][h.couleur] = couleur;
      data[i][h.description] = description;
      data[i][h.actif] = 'oui';
      sh.getRange(i + 1, 1, 1, headers.length).setValues([data[i]]);
      logHistory_('', 'planning_modele_save', session.userId, 'admin', `Modèle shift mis à jour: ${nom}`, {});
      return { ok: true, idModele };
    }
  }

  appendRowObject_(P_SH.MODELES, {
    idModele, nom, heureDebut, heureFin, couleur, description, actif: 'oui'
  });
  logHistory_('', 'planning_modele_save', session.userId, 'admin', `Modèle shift créé: ${nom}`, {});
  return { ok: true, idModele };
}

function apiAdminDeleteShiftModele(token, idModele) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  const sh = getSheet_(P_SH.MODELES);
  const data = sh.getDataRange().getValues();
  const h = headerIndexMap_(data[0]);
  for (let i = 1; i < data.length; i++) {
    if (normTrim_(data[i][h.idModele]) === normTrim_(idModele)) {
      data[i][h.actif] = 'non';
      sh.getRange(i + 1, 1, 1, data[0].length).setValues([data[i]]);
      logHistory_('', 'planning_modele_delete', session.userId, 'admin', `Modèle shift supprimé: ${data[i][h.nom]}`, {});
      return { ok: true };
    }
  }
  return { ok: true };
}

/* ========= PLAGES HORAIRES ========= */

function apiAdminListPlagesHoraires(token, filters) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  filters = filters || {};
  const equipeFilter = normUpper_(filters.equipe || '');

  let plages = getRowsAsObjects_(P_SH.PLAGES)
    .map(_normalizePlageRow_)
    .filter(p => p.actif !== 'non');

  if (equipeFilter) {
    plages = plages.filter(p => !normUpper_(p.equipe) || normUpper_(p.equipe) === equipeFilter);
  }

  plages.sort((a, b) => {
    const eq = String(a.equipe || '').localeCompare(String(b.equipe || ''), 'fr');
    if (eq !== 0) return eq;
    const hd = a.heureDebut.localeCompare(b.heureDebut);
    if (hd !== 0) return hd;
    return a.nom.localeCompare(b.nom, 'fr');
  });

  return { ok: true, plages };
}

function apiAdminSavePlageHoraire(token, payload) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  payload = payload || {};

  const idPlage = normTrim_(payload.idPlage) || ('PLG-' + Utilities.getUuid().slice(0, 6).toUpperCase());
  const equipe = sanitizeText_(payload.equipe || '');
  const nom = sanitizeText_(payload.nom || '');
  const heureDebut = _normalizeTimeValue_(payload.heureDebut);
  const heureFin = _normalizeTimeValue_(payload.heureFin);
  const niveauFlux = sanitizeText_(payload.niveauFlux || 'Normal');
  const joursActifs = _normalizeDaysList_(payload.joursActifs || '1,2,3,4,5');
  const effectifMin = Math.max(1, Number(payload.effectifMin || 1));
  const couleur = sanitizeText_(payload.couleur || '#888888');

  if (!nom) throw new Error('Nom de la plage requis.');
  if (!heureDebut || !heureFin) throw new Error('Heures début et fin requises.');

  const sh = getSheet_(P_SH.PLAGES);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const h = headerIndexMap_(headers);

  for (let i = 1; i < data.length; i++) {
    if (normTrim_(data[i][h.idPlage]) === idPlage) {
      data[i][h.equipe] = equipe;
      data[i][h.nom] = nom;
      data[i][h.heureDebut] = heureDebut;
      data[i][h.heureFin] = heureFin;
      data[i][h.niveauFlux] = niveauFlux;
      data[i][h.joursActifs] = joursActifs;
      data[i][h.effectifMin] = effectifMin;
      data[i][h.couleur] = couleur;
      data[i][h.actif] = 'oui';
      sh.getRange(i + 1, 1, 1, headers.length).setValues([data[i]]);
      logHistory_('', 'planning_plage_save', session.userId, 'admin', `Plage mise à jour: ${nom}`, {});
      return { ok: true, idPlage };
    }
  }

  appendRowObject_(P_SH.PLAGES, {
    idPlage, equipe, nom, heureDebut, heureFin, niveauFlux, joursActifs, effectifMin, couleur, actif: 'oui'
  });
  logHistory_('', 'planning_plage_save', session.userId, 'admin', `Plage créée: ${nom}`, {});
  return { ok: true, idPlage };
}

function apiAdminDeletePlageHoraire(token, idPlage) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  const sh = getSheet_(P_SH.PLAGES);
  const data = sh.getDataRange().getValues();
  const h = headerIndexMap_(data[0]);
  for (let i = 1; i < data.length; i++) {
    if (normTrim_(data[i][h.idPlage]) === normTrim_(idPlage)) {
      data[i][h.actif] = 'non';
      sh.getRange(i + 1, 1, 1, data[0].length).setValues([data[i]]);
      logHistory_('', 'planning_plage_delete', session.userId, 'admin', `Plage supprimée: ${data[i][h.nom]}`, {});
      return { ok: true };
    }
  }
  return { ok: true };
}

/* ========= PLANNING — LECTURE ========= */

function apiGetMyPlanning(token, filters) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  const user = getUserByIdRaw_(session.userId);
  if (!user) throw new Error('Utilisateur introuvable.');

  filters = filters || {};
  const dateDebut = toIsoDateString_(filters.dateDebut || _mondayOfWeek_(new Date()));
  const dateFin = toIsoDateString_(filters.dateFin || _addDaysToIso_(dateDebut, 6));
  const today = toIsoDate_(new Date());

  const shifts = _getShiftsForUser_(user.userId, dateDebut, dateFin);
  const leaves = _getLeavesForUser_(user.userId, dateDebut, dateFin);
  const todayShift = _getShiftForUserDate_(user.userId, today);
  const todayPointage = _getPointageForUserDate_(user.userId, today);
  const todayLeave = _getUserLeaveForDate_(user.userId, today);
  const todayStatus = _buildDayStatus_(todayShift, todayPointage, todayLeave, new Date());

  return {
    ok: true,
    user: publicUser_(user),
    dateDebut,
    dateFin,
    shifts,
    leaves,
    todayShift,
    todayPointage,
    todayLeave,
    todayStatus
  };
}

function apiGetGroupPlanning(token, filters) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  const user = getUserByIdRaw_(session.userId);
  if (!user) throw new Error('Utilisateur introuvable.');

  filters = filters || {};
  const dateDebut = toIsoDateString_(filters.dateDebut || _mondayOfWeek_(new Date()));
  const dateFin = toIsoDateString_(filters.dateFin || _addDaysToIso_(dateDebut, 6));
  const equipe = normUpper_(user.equipe || '');

  const agents = _getActiveUsers_(equipe);
  const shifts = getRowsAsObjects_(P_SH.SHIFTS).map(normalizeShiftRow_)
    .filter(s => normUpper_(s.equipe) === equipe && s.date >= dateDebut && s.date <= dateFin);
  const leaves = _getLeavesForEquipe_(equipe, dateDebut, dateFin);
  const shiftsByUser = _indexShiftsByUser_(shifts);

  return { ok: true, agents, shifts, shiftsByUser, leaves, dateDebut, dateFin };
}

function apiAdminGetPlanning(token, filters) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  filters = filters || {};

  const dateDebut = toIsoDateString_(filters.dateDebut || _mondayOfWeek_(new Date()));
  const dateFin = toIsoDateString_(filters.dateFin || _addDaysToIso_(dateDebut, 6));
  const equipeFilter = normUpper_(filters.equipe || '');
  const userIdFilter = normUpper_(filters.userId || '');

  let users = _getActiveUsers_(equipeFilter || '');
  if (userIdFilter) users = users.filter(u => normUpper_(u.userId) === userIdFilter);

  let shifts = getRowsAsObjects_(P_SH.SHIFTS).map(normalizeShiftRow_)
    .filter(s => s.date >= dateDebut && s.date <= dateFin);
  if (equipeFilter) shifts = shifts.filter(s => normUpper_(s.equipe) === equipeFilter);
  if (userIdFilter) shifts = shifts.filter(s => normUpper_(s.userId) === userIdFilter);

  const leaves = _getLeavesForPeriod_(dateDebut, dateFin, { equipe: filters.equipe, userId: filters.userId });
  const shiftsByUser = _indexShiftsByUser_(shifts);
  const plages = getRowsAsObjects_(P_SH.PLAGES)
    .map(_normalizePlageRow_)
    .filter(p => p.actif !== 'non')
    .filter(p => !equipeFilter || !normUpper_(p.equipe) || normUpper_(p.equipe) === equipeFilter);

  return {
    ok: true,
    dateDebut,
    dateFin,
    users,
    shifts,
    shiftsByUser,
    leaves,
    plages
  };
}

function apiGetPresenceOverview(token, filters) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  const requester = getUserByIdRaw_(session.userId);
  if (!requester) throw new Error('Utilisateur introuvable.');

  filters = filters || {};
  const date = toIsoDateString_(filters.date || new Date());
  const equipeFilter = normUpper_(filters.equipe || '');
  const isAdmin = requester.role === APP.ROLES.ADMIN;
  const targetEquipe = isAdmin ? equipeFilter : normUpper_(requester.equipe || '');

  let users = _getActiveUsers_(targetEquipe || '');
  if (!isAdmin) users = users.filter(u => normUpper_(u.equipe) === targetEquipe);

  const allShifts = getRowsAsObjects_(P_SH.SHIFTS).map(normalizeShiftRow_).filter(s => s.date === date);
  const allLeaves = getRowsAsObjects_(APP.SHEETS.REQUESTS).map(normalizeRequestRow_)
    .filter(r => r.dateDebut <= date && r.dateFin >= date && (r.statut === APP.STATUS.APPROVED || r.statut === APP.STATUS.PENDING));
  const allPointage = getRowsAsObjects_(P_SH.POINTAGE).map(normalizePointageRow_).filter(p => p.date === date);

  const shiftByUser = {};
  allShifts.forEach(s => shiftByUser[normUpper_(s.userId)] = s);
  const pointageByUser = {};
  allPointage.forEach(p => pointageByUser[normUpper_(p.userId)] = p);
  const leavesByUser = {};
  allLeaves.forEach(l => { leavesByUser[normUpper_(l.userId)] = l; });

  const now = new Date();
  const overview = users.map(u => {
    const uid = normUpper_(u.userId);
    const shift = shiftByUser[uid] || null;
    const pointage = pointageByUser[uid] || null;
    const leave = leavesByUser[uid] || null;
    const status = _buildDayStatus_(shift, pointage, leave, now);
    return {
      userId: u.userId,
      nomAgent: u.nomAgent,
      equipe: u.equipe,
      shift: shift ? {
        idShift: shift.idShift,
        heureDebut: shift.heureDebut,
        heureFin: shift.heureFin,
        pauseDebut: shift.pauseDebut,
        pauseFin: shift.pauseFin,
        isOff: shift.isOff,
        nomModele: shift.nomModele,
        couleur: shift.couleurModele,
        statut: shift.statut
      } : null,
      leave: leave ? { statut: leave.statut, typeConge: leave.typeConge } : null,
      pointage: pointage ? {
        heureArrivee: pointage.heureArrivee,
        heureDepart: pointage.heureDepart,
        heuresSup: pointage.heuresSup,
        retardMinutes: pointage.retardMinutes,
        statut: pointage.statut
      } : null,
      statut: status.label,
      statusKey: status.key,
      details: status.details
    };
  });

  const stats = {
    presents: overview.filter(o => ['present','present_retard','reparti'].indexOf(o.statusKey) !== -1).length,
    absents: overview.filter(o => o.statusKey === 'absent').length,
    enConge: overview.filter(o => o.statusKey === 'en_conge').length,
    congeEnAttente: overview.filter(o => o.statusKey === 'conge_attente').length,
    nonPlanifie: overview.filter(o => o.statusKey === 'non_planifie').length,
    aVenir: overview.filter(o => o.statusKey === 'a_venir').length,
    off: overview.filter(o => o.statusKey === 'off').length,
  };

  return { ok: true, date, overview, stats };
}

/* ========= PLANNING — ÉCRITURE ========= */

function apiAdminSaveShift(token, payload) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  payload = payload || {};

  const userId = normUpper_(payload.userId || '');
  const date = toIsoDateString_(payload.date);
  const user = getUserByIdRaw_(userId);
  if (!userId) throw new Error('Agent requis.');
  if (!date) throw new Error('Date requise.');
  if (!user) throw new Error('Agent introuvable.');

  const isOff = String(payload.isOff || '').toLowerCase() === 'oui' || payload.isOff === true || String(payload.statut || '').toUpperCase() === 'OFF';
  const heureDebut = isOff ? '' : _normalizeTimeValue_(payload.heureDebut);
  const heureFin = isOff ? '' : _normalizeTimeValue_(payload.heureFin);
  const pauseRange = isOff ? { pauseDebut: '', pauseFin: '' } : _normalizePauseRange_(payload.pauseDebut, payload.pauseFin);
  const pauseDebut = pauseRange.pauseDebut;
  const pauseFin = pauseRange.pauseFin;
  const nomModele = sanitizeText_(payload.nomModele || '');
  const couleurModele = sanitizeText_(payload.couleurModele || '#4ea1ff');
  const notes = sanitizeText_(payload.notes || '');
  const statut = isOff ? 'OFF' : 'Planifié';
  const dureePrevue = isOff ? 0 : _calcShiftNetDuration_(heureDebut, heureFin, pauseDebut, pauseFin);

  if (!isOff && (!heureDebut || !heureFin)) throw new Error('Heures requises pour un shift planifié.');
  if (!isOff && pauseDebut && pauseFin && !_isPauseInsideShift_(heureDebut, heureFin, pauseDebut, pauseFin)) {
    throw new Error('La pause doit être comprise dans le shift.');
  }

  const leave = _getUserLeaveForDate_(userId, date);
  const leaveWarning = !!leave;

  const sh = getSheet_(P_SH.SHIFTS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const h = headerIndexMap_(headers);
  const providedId = normTrim_(payload.idShift);

  for (let i = 1; i < data.length; i++) {
    const rowUser = normUpper_(data[i][h.userId]);
    const rowDate = toIsoDate_(data[i][h.date]);
    const rowId = normTrim_(data[i][h.idShift]);
    if ((providedId && rowId === providedId) || (rowUser === userId && rowDate === date)) {
      data[i][h.userId] = user.userId;
      data[i][h.nomAgent] = user.nomComplet;
      data[i][h.equipe] = user.equipe || '';
      data[i][h.date] = date;
      data[i][h.heureDebut] = heureDebut;
      data[i][h.heureFin] = heureFin;
      if (h.pauseDebut !== undefined) data[i][h.pauseDebut] = pauseDebut;
      if (h.pauseFin !== undefined) data[i][h.pauseFin] = pauseFin;
      if (h.isOff !== undefined) data[i][h.isOff] = isOff ? 'oui' : 'non';
      if (h.dureePrevue !== undefined) data[i][h.dureePrevue] = dureePrevue;
      data[i][h.nomModele] = nomModele;
      data[i][h.couleurModele] = couleurModele;
      data[i][h.statut] = statut;
      data[i][h.notes] = notes;
      sh.getRange(i + 1, 1, 1, headers.length).setValues([data[i]]);
      logHistory_('', 'planning_shift_save', session.userId, 'admin', `Shift mis à jour ${user.userId} ${date} ${isOff ? 'OFF' : heureDebut + '-' + heureFin}`, { cibleUserId: user.userId, cibleEquipe: user.equipe || '' });
      return { ok: true, idShift: rowId || providedId, leaveWarning };
    }
  }

  const idShift = providedId || ('SHF-' + Utilities.getUuid().slice(0, 8).toUpperCase());
  appendRowObject_(P_SH.SHIFTS, {
    idShift,
    userId: user.userId,
    nomAgent: user.nomComplet,
    equipe: user.equipe || '',
    date,
    heureDebut,
    heureFin,
    pauseDebut,
    pauseFin,
    isOff: isOff ? 'oui' : 'non',
    dureePrevue,
    nomModele,
    couleurModele,
    statut,
    notes,
    dateCreation: new Date(),
    creePar: session.userId
  });

  logHistory_('', 'planning_shift_save', session.userId, 'admin', `Shift créé ${user.userId} ${date} ${isOff ? 'OFF' : heureDebut + '-' + heureFin}`, { cibleUserId: user.userId, cibleEquipe: user.equipe || '' });
  return { ok: true, idShift, leaveWarning };
}

function apiAdminDeleteShift(token, idShift) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  const sh = getSheet_(P_SH.SHIFTS);
  const data = sh.getDataRange().getValues();
  const h = headerIndexMap_(data[0]);
  for (let i = data.length - 1; i >= 1; i--) {
    if (normTrim_(data[i][h.idShift]) === normTrim_(idShift)) {
      const uid = String(data[i][h.userId] || '');
      const eq = String(data[i][h.equipe] || '');
      sh.deleteRow(i + 1);
      logHistory_('', 'planning_shift_delete', session.userId, 'admin', `Shift supprimé ${idShift}`, { cibleUserId: uid, cibleEquipe: eq });
      return { ok: true };
    }
  }
  return { ok: true };
}

function apiAdminCopyDayShift(token, payload) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  payload = payload || {};

  const sourceUserId = normUpper_(payload.sourceUserId || '');
  const dateSource = toIsoDateString_(payload.dateSource);
  const dateCible = toIsoDateString_(payload.dateCible || payload.dateSource);
  const overwrite = !!payload.overwrite;
  const targetUserIds = Array.isArray(payload.targetUserIds)
    ? payload.targetUserIds.map(normUpper_).filter(Boolean)
    : (payload.targetUserId ? [normUpper_(payload.targetUserId)] : []);

  if (!sourceUserId) throw new Error('Agent source requis.');
  if (!dateSource || !dateCible) throw new Error('Dates source/cible requises.');
  if (!targetUserIds.length) throw new Error('Au moins un agent cible est requis.');

  const srcShift = _getShiftForUserDate_(sourceUserId, dateSource);
  if (!srcShift) throw new Error('Aucun shift source trouvé pour cette journée.');

  let created = 0, updated = 0, skipped = 0;
  const warnings = [];

  targetUserIds.forEach(targetId => {
    const targetUser = getUserByIdRaw_(targetId);
    if (!targetUser || String(targetUser.actif || '').toLowerCase() !== 'oui') {
      skipped++;
      warnings.push(`${targetId}: utilisateur introuvable ou inactif`);
      return;
    }

    const approvedLeave = _getApprovedLeaveForDate_(targetId, dateCible);
    if (approvedLeave) {
      skipped++;
      warnings.push(`${targetId}: congé validé sur ${dateCible}`);
      return;
    }

    const existing = _getShiftForUserDate_(targetId, dateCible);
    if (existing && !overwrite) {
      skipped++;
      warnings.push(`${targetId}: shift déjà existant`);
      return;
    }

    apiAdminSaveShift(token, {
      idShift: existing ? existing.idShift : undefined,
      userId: targetId,
      date: dateCible,
      heureDebut: srcShift.heureDebut,
      heureFin: srcShift.heureFin,
      pauseDebut: srcShift.pauseDebut,
      pauseFin: srcShift.pauseFin,
      isOff: srcShift.isOff === 'oui',
      statut: srcShift.statut,
      nomModele: srcShift.nomModele,
      couleurModele: srcShift.couleurModele,
      notes: srcShift.notes
    });

    if (existing) updated++; else created++;
  });

  logHistory_('', 'planning_copy_day', session.userId, 'admin', `Copie journée ${sourceUserId} ${dateSource} -> ${dateCible} | créés=${created} maj=${updated} ignorés=${skipped}`, {});
  return { ok: true, created, updated, skipped, warnings };
}


function apiAgentRequestShiftSwap(token, payload) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  const requester = getUserByIdRaw_(session.userId);
  if (!requester) throw new Error('Utilisateur introuvable.');
  if (requester.role === APP.ROLES.ADMIN) throw new Error('Action réservée aux agents.');
  payload = payload || {};

  const targetUserId = normUpper_(payload.targetUserId || '');
  const dateShift = toIsoDateString_(payload.dateShift);
  const requesterComment = sanitizeText_(payload.requesterComment || '');
  const todayIso = toIsoDate_(new Date());

  if (!targetUserId) throw new Error('Collègue cible requis.');
  if (!dateShift) throw new Error('Date du shift requise.');
  if (dateShift <= todayIso) throw new Error('Les échanges sont possibles uniquement pour une date future.');
  if (targetUserId === normUpper_(requester.userId)) throw new Error('Choisissez un autre agent.');

  const targetUser = getUserByIdRaw_(targetUserId);
  if (!targetUser || String(targetUser.actif || '').toLowerCase() !== 'oui' || String(targetUser.role || '') === APP.ROLES.ADMIN) {
    throw new Error('Collègue cible introuvable ou inactif.');
  }
  if (normUpper_(targetUser.equipe || '') !== normUpper_(requester.equipe || '')) {
    throw new Error('L’échange doit se faire dans la même équipe.');
  }

  const requesterShift = _getShiftForUserDate_(requester.userId, dateShift);
  const targetShift = _getShiftForUserDate_(targetUserId, dateShift);
  if (!requesterShift || requesterShift.isOff === 'oui') throw new Error('Vous devez avoir un shift planifié ce jour-là.');
  if (!targetShift || targetShift.isOff === 'oui') throw new Error('Le collègue cible doit aussi avoir un shift planifié ce jour-là.');
  if (_getApprovedLeaveForDate_(requester.userId, dateShift) || _getApprovedLeaveForDate_(targetUserId, dateShift)) {
    throw new Error('Échange impossible : un congé validé existe sur cette date.');
  }

  const existingPending = getRowsAsObjects_(P_SH.SWAPS).find(s =>
    normUpper_(s.requesterUserId) === normUpper_(requester.userId) &&
    normUpper_(s.targetUserId) === targetUserId &&
    toIsoDate_(s.dateShift) === dateShift &&
    String(s.statut || '') === 'En attente'
  );
  const existingPendingSameDay = getRowsAsObjects_(P_SH.SWAPS).find(s =>
    toIsoDate_(s.dateShift) === dateShift &&
    String(s.statut || '') === 'En attente' &&
    (normUpper_(s.requesterUserId) === normUpper_(requester.userId) || normUpper_(s.targetUserId) === normUpper_(requester.userId))
  );
  if (existingPending || existingPendingSameDay) throw new Error('Une demande d’échange en attente existe déjà pour cette date.');

  const idSwap = 'SWP-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  appendRowObject_(P_SH.SWAPS, {
    idSwap,
    requesterUserId: requester.userId,
    requesterNom: requester.nomComplet || requester.userId,
    requesterEquipe: requester.equipe || '',
    targetUserId: targetUser.userId,
    targetNom: targetUser.nomComplet || targetUser.userId,
    targetEquipe: targetUser.equipe || '',
    dateShift,
    requesterShiftId: requesterShift.idShift,
    targetShiftId: targetShift.idShift,
    requesterComment,
    adminComment: '',
    statut: 'En attente',
    dateCreation: new Date(),
    dateDecision: '',
    decidedBy: ''
  });

  logHistory_('', 'planning_swap_request', requester.userId, requester.role, `Demande échange ${dateShift} avec ${targetUser.userId}`, { cibleUserId: targetUser.userId, cibleEquipe: targetUser.equipe || '' });
  return { ok: true, idSwap };
}

function apiListShiftSwapRequests(token, filters) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  const user = getUserByIdRaw_(session.userId);
  if (!user) throw new Error('Utilisateur introuvable.');
  filters = filters || {};
  const statut = String(filters.statut || '');
  const isAdmin = user.role === APP.ROLES.ADMIN;

  let rows = getRowsAsObjects_(P_SH.SWAPS).map(_normalizeSwapRow_);
  if (!isAdmin) {
    rows = rows.filter(r => normUpper_(r.requesterUserId) === normUpper_(user.userId) || normUpper_(r.targetUserId) === normUpper_(user.userId));
  } else {
    if (filters.equipe) rows = rows.filter(r => normUpper_(r.requesterEquipe) === normUpper_(filters.equipe) || normUpper_(r.targetEquipe) === normUpper_(filters.equipe));
    if (filters.userId) rows = rows.filter(r => normUpper_(r.requesterUserId) === normUpper_(filters.userId) || normUpper_(r.targetUserId) === normUpper_(filters.userId));
  }
  if (statut) rows = rows.filter(r => r.statut === statut);
  rows.sort((a,b) => (b.dateCreation||'').localeCompare(a.dateCreation||''));

  const total = rows.length;
  const page = Math.max(1, Number(filters.page || 1));
  const pageSize = Math.max(5, Math.min(200, Number(filters.pageSize || 20)));
  rows = rows.slice((page - 1) * pageSize, page * pageSize).map(_hydrateSwapTimeline_);
  return { ok: true, items: rows, total, page, pageSize };
}

function apiCancelShiftSwap(token, idSwap) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  const user = getUserByIdRaw_(session.userId);
  if (!user) throw new Error('Utilisateur introuvable.');
  const sh = getSheet_(P_SH.SWAPS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const h = headerIndexMap_(headers);
  idSwap = normTrim_(idSwap);
  for (let i = 1; i < data.length; i++) {
    if (normTrim_(data[i][h.idSwap]) !== idSwap) continue;
    if (normUpper_(data[i][h.requesterUserId]) !== normUpper_(user.userId)) throw new Error('Vous pouvez annuler uniquement vos propres demandes.');
    if (String(data[i][h.statut] || '') !== 'En attente') throw new Error('Seules les demandes en attente peuvent être annulées.');
    data[i][h.statut] = 'Annulé';
    data[i][h.dateDecision] = new Date();
    data[i][h.decidedBy] = user.userId;
    sh.getRange(i + 1, 1, 1, headers.length).setValues([data[i]]);
    logHistory_('', 'planning_swap_cancel', user.userId, user.role, `Annulation échange ${idSwap}`, { cibleUserId: String(data[i][h.targetUserId] || ''), cibleEquipe: String(data[i][h.targetEquipe] || '') });
    return { ok: true };
  }
  throw new Error('Demande introuvable.');
}

function apiAdminProcessShiftSwap(token, payload) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  const admin = requireAdmin_(session);
  payload = payload || {};
  const idSwap = normTrim_(payload.idSwap);
  const action = String(payload.action || '').toLowerCase();
  const adminComment = sanitizeText_(payload.adminComment || '');
  if (!idSwap) throw new Error('ID échange requis.');
  if (['approve', 'reject'].indexOf(action) === -1) throw new Error('Action invalide.');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = getSheet_(P_SH.SWAPS);
    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const h = headerIndexMap_(headers);
    for (let i = 1; i < data.length; i++) {
      if (normTrim_(data[i][h.idSwap]) !== idSwap) continue;
      if (String(data[i][h.statut] || '') !== 'En attente') throw new Error('Cette demande n’est plus en attente.');
      const swap = _normalizeSwapRow_(objectFromRow_(headers, data[i]));
      if (action === 'approve') {
        if (swap.dateShift <= toIsoDate_(new Date())) throw new Error('Échange impossible : la date n’est plus future.');
        _swapShiftAssignments_(swap);
        data[i][h.statut] = 'Validé';
      } else {
        data[i][h.statut] = 'Refusé';
      }
      data[i][h.adminComment] = adminComment;
      data[i][h.dateDecision] = new Date();
      data[i][h.decidedBy] = admin.userId;
      sh.getRange(i + 1, 1, 1, headers.length).setValues([data[i]]);
      logHistory_('', 'planning_swap_' + (action === 'approve' ? 'approved' : 'rejected'), admin.userId, 'admin', `Échange ${idSwap} ${action === 'approve' ? 'validé' : 'refusé'}${adminComment ? ' | ' + adminComment : ''}`, { cibleUserId: swap.requesterUserId, cibleEquipe: swap.requesterEquipe || '' });
      return { ok: true };
    }
    throw new Error('Demande introuvable.');
  } finally {
    lock.releaseLock();
  }
}

function _swapShiftAssignments_(swap) {
  const sh = getSheet_(P_SH.SHIFTS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const h = headerIndexMap_(headers);
  let rowA = -1, rowB = -1;
  for (let i = 1; i < data.length; i++) {
    if (normTrim_(data[i][h.idShift]) === normTrim_(swap.requesterShiftId)) rowA = i;
    if (normTrim_(data[i][h.idShift]) === normTrim_(swap.targetShiftId)) rowB = i;
  }
  if (rowA === -1 || rowB === -1) throw new Error('Shifts introuvables pour l’échange.');
  const userA = getUserByIdRaw_(swap.requesterUserId);
  const userB = getUserByIdRaw_(swap.targetUserId);
  if (!userA || !userB) throw new Error('Utilisateurs introuvables pour l’échange.');
  if (_getApprovedLeaveForDate_(userA.userId, swap.dateShift) || _getApprovedLeaveForDate_(userB.userId, swap.dateShift)) {
    throw new Error('Échange impossible : congé validé détecté.');
  }
  data[rowA][h.userId] = userB.userId;
  data[rowA][h.nomAgent] = userB.nomComplet || userB.userId;
  data[rowA][h.equipe] = userB.equipe || '';
  data[rowB][h.userId] = userA.userId;
  data[rowB][h.nomAgent] = userA.nomComplet || userA.userId;
  data[rowB][h.equipe] = userA.equipe || '';
  sh.getRange(rowA + 1, 1, 1, headers.length).setValues([data[rowA]]);
  sh.getRange(rowB + 1, 1, 1, headers.length).setValues([data[rowB]]);
}

function apiAdminBulkShifts(token, payload) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  payload = payload || {};
  const operation = String(payload.operation || '');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (operation === 'copyWeek') return _bulkCopyWeek_(session, payload);
    if (operation === 'applyModel') return _bulkApplyModel_(session, payload);
    if (operation === 'clearWeek') return _bulkClearWeek_(session, payload);
    if (operation === 'assignShift') return _bulkAssignShift_(session, token, payload);
    throw new Error('Opération bulk inconnue: ' + operation);
  } finally {
    lock.releaseLock();
  }
}

function _bulkCopyWeek_(session, payload) {
  const srcStart = toIsoDateString_(payload.sourceDateDebut);
  const tgtStart = toIsoDateString_(payload.targetDateDebut);
  const equipe = normUpper_(payload.equipe || '');
  const userIds = Array.isArray(payload.userIds) ? payload.userIds.map(normUpper_) : [];
  if (!srcStart || !tgtStart) throw new Error('Dates source/cible requises.');

  const srcEnd = _addDaysToIso_(srcStart, 6);
  let srcShifts = getRowsAsObjects_(P_SH.SHIFTS).map(normalizeShiftRow_)
    .filter(s => s.date >= srcStart && s.date <= srcEnd);
  if (equipe) srcShifts = srcShifts.filter(s => normUpper_(s.equipe) === equipe);
  if (userIds.length) srcShifts = srcShifts.filter(s => userIds.indexOf(normUpper_(s.userId)) !== -1);

  const diffDays = Math.round((new Date(tgtStart + 'T00:00:00') - new Date(srcStart + 'T00:00:00')) / 86400000);
  const existingMap = _buildShiftMap_();

  let created = 0, skipped = 0;
  srcShifts.forEach(s => {
    const newDate = _addDaysToIso_(s.date, diffDays);
    const key = normUpper_(s.userId) + '_' + newDate;
    if (existingMap[key]) { skipped++; return; }
    const leave = _getApprovedLeaveForDate_(s.userId, newDate);
    if (leave) { skipped++; return; }
    appendRowObject_(P_SH.SHIFTS, {
      idShift: 'SHF-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
      userId: s.userId,
      nomAgent: s.nomAgent,
      equipe: s.equipe,
      date: newDate,
      heureDebut: s.heureDebut,
      heureFin: s.heureFin,
      pauseDebut: s.pauseDebut,
      pauseFin: s.pauseFin,
      isOff: s.isOff,
      dureePrevue: s.dureePrevue,
      nomModele: s.nomModele,
      couleurModele: s.couleurModele,
      statut: s.statut,
      notes: s.notes,
      dateCreation: new Date(),
      creePar: session.userId
    });
    existingMap[key] = true;
    created++;
  });

  logHistory_('', 'planning_bulk_copy', session.userId, 'admin', `Copie semaine ${srcStart}→${tgtStart} | ${created} créés, ${skipped} ignorés`, {});
  return { ok: true, created, skipped };
}

function _bulkApplyModel_(session, payload) {
  const idModele = normTrim_(payload.idModele);
  const dateDebut = toIsoDateString_(payload.dateDebut);
  const dateFin = toIsoDateString_(payload.dateFin);
  const equipe = normUpper_(payload.equipe || '');
  const userIds = Array.isArray(payload.userIds) ? payload.userIds.map(normUpper_) : [];
  const joursSemaine = _normalizeDaysList_(payload.joursSemaine || '1,2,3,4,5').split(',').map(Number);
  const overwrite = !!payload.overwrite;

  const customPauseDebut = normTrim_(payload.pauseDebut || '');
  const customPauseFin = normTrim_(payload.pauseFin || '');

  if (!idModele) throw new Error('Modèle de shift requis.');
  if (!dateDebut || !dateFin) throw new Error('Dates requises.');

  const modele = getRowsAsObjects_(P_SH.MODELES).map(_normalizeModeleRow_).find(m => normTrim_(m.idModele) === idModele && m.actif !== 'non');
  if (!modele) throw new Error('Modèle introuvable.');

  let users = _getActiveUsers_(equipe || '');
  if (userIds.length) users = users.filter(u => userIds.indexOf(normUpper_(u.userId)) !== -1);
  if (!users.length) throw new Error('Aucun agent trouvé avec ces critères.');

  const dates = enumerateDates_(dateDebut, dateFin)
    .filter(d => {
      const dow = d.getDay();
      const normalized = dow === 0 ? 7 : dow;
      return joursSemaine.indexOf(normalized) !== -1;
    })
    .map(d => toIsoDate_(d));

  if (!dates.length) throw new Error('Aucun jour à planifier avec ces critères.');

  const sh = getSheet_(P_SH.SHIFTS);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const h = headerIndexMap_(headers);
  const existing = {};
  for (let i = 1; i < data.length; i++) {
    const uid = normUpper_(data[i][h.userId]);
    const date = toIsoDate_(data[i][h.date]);
    if (uid && date) existing[uid + '_' + date] = i + 1;
  }

  let created = 0, updated = 0, skipped = 0;
  users.forEach(u => {
    dates.forEach(date => {
      const leave = _getApprovedLeaveForDate_(u.userId, date);
      if (leave) { skipped++; return; }
      const key = normUpper_(u.userId) + '_' + date;
      if (existing[key] && !overwrite) { skipped++; return; }

      if (existing[key] && overwrite) {
        const row = sh.getRange(existing[key], 1, 1, headers.length).getValues()[0];
        row[h.heureDebut] = modele.heureDebut;
        row[h.heureFin] = modele.heureFin;
        if (h.pauseDebut !== undefined) row[h.pauseDebut] = customPauseDebut;
        if (h.pauseFin !== undefined) row[h.pauseFin] = customPauseFin;
        if (h.isOff !== undefined) row[h.isOff] = 'non';
        if (h.dureePrevue !== undefined) row[h.dureePrevue] = _calcShiftNetDuration_(modele.heureDebut, modele.heureFin, customPauseDebut, customPauseFin);
        row[h.nomModele] = modele.nom;
        row[h.couleurModele] = modele.couleur || '#4ea1ff';
        row[h.statut] = 'Planifié';
        row[h.notes] = '';
        sh.getRange(existing[key], 1, 1, headers.length).setValues([row]);
        updated++;
      } else {
        appendRowObject_(P_SH.SHIFTS, {
          idShift: 'SHF-' + Utilities.getUuid().slice(0, 8).toUpperCase(),
          userId: u.userId,
          nomAgent: u.nomAgent,
          equipe: u.equipe,
          date,
          heureDebut: modele.heureDebut,
          heureFin: modele.heureFin,
          pauseDebut: customPauseDebut,
          pauseFin: customPauseFin,
          isOff: 'non',
          dureePrevue: _calcShiftNetDuration_(modele.heureDebut, modele.heureFin, customPauseDebut, customPauseFin),
          nomModele: modele.nom,
          couleurModele: modele.couleur || '#4ea1ff',
          statut: 'Planifié',
          notes: '',
          dateCreation: new Date(),
          creePar: session.userId
        });
        existing[key] = true;
        created++;
      }
    });
  });

  logHistory_('', 'planning_bulk_apply', session.userId, 'admin', `Modèle ${modele.nom} appliqué | ${created} créés, ${updated} maj, ${skipped} ignorés`, {});
  return { ok: true, created, updated, skipped };
}

function _bulkClearWeek_(session, payload) {
  const dateDebut = toIsoDateString_(payload.dateDebut);
  const dateFin = toIsoDateString_(payload.dateFin || _addDaysToIso_(dateDebut, 6));
  const equipe = normUpper_(payload.equipe || '');
  const userIds = Array.isArray(payload.userIds) ? payload.userIds.map(normUpper_) : [];
  if (!dateDebut) throw new Error('Date début requise.');

  const sh = getSheet_(P_SH.SHIFTS);
  const data = sh.getDataRange().getValues();
  const h = headerIndexMap_(data[0]);
  let deleted = 0;

  for (let i = data.length - 1; i >= 1; i--) {
    const date = toIsoDate_(data[i][h.date]);
    const uid = normUpper_(data[i][h.userId]);
    const eq = normUpper_(data[i][h.equipe]);
    if (date < dateDebut || date > dateFin) continue;
    if (equipe && eq !== equipe) continue;
    if (userIds.length && userIds.indexOf(uid) === -1) continue;
    sh.deleteRow(i + 1);
    deleted++;
  }

  logHistory_('', 'planning_bulk_clear', session.userId, 'admin', `Vidage planning ${dateDebut}→${dateFin} | ${deleted} shifts supprimés`, {});
  return { ok: true, deleted };
}


function _bulkAssignShift_(session, token, payload) {
  const targetUserIds = Array.isArray(payload.targetUserIds)
    ? payload.targetUserIds.map(normUpper_).filter(Boolean)
    : String(payload.targetUserIds || '').split(',').map(normUpper_).filter(Boolean);
  const dateDebut = toIsoDateString_(payload.dateDebut);
  const dateFin = toIsoDateString_(payload.dateFin || payload.dateDebut);
  const joursSemaine = _normalizeDaysList_(payload.joursSemaine || '1,2,3,4,5').split(',').map(Number);
  const overwrite = !!payload.overwrite;
  if (!targetUserIds.length) throw new Error('Au moins un agent est requis.');
  if (!dateDebut || !dateFin) throw new Error('Dates requises.');

  let base = null;
  const idModele = normTrim_(payload.idModele);
  const pauseRange = _normalizePauseRange_(payload.pauseDebut, payload.pauseFin);
  const pauseDebut = pauseRange.pauseDebut;
  const pauseFin = pauseRange.pauseFin;
  if (idModele) {
    const modele = getRowsAsObjects_(P_SH.MODELES).map(_normalizeModeleRow_).find(m => m.idModele === idModele && m.actif !== 'non');
    if (!modele) throw new Error('Modèle introuvable.');
    if (pauseDebut && pauseFin && !_isPauseInsideShift_(modele.heureDebut, modele.heureFin, pauseDebut, pauseFin)) throw new Error('La pause doit être comprise dans le shift du modèle.');
    base = { heureDebut: modele.heureDebut, heureFin: modele.heureFin, pauseDebut, pauseFin, isOff: false, nomModele: modele.nom, couleurModele: modele.couleur || '#4ea1ff', notes: sanitizeText_(payload.notes || '') };
  } else {
    const heureDebut = _normalizeTimeValue_(payload.heureDebut);
    const heureFin = _normalizeTimeValue_(payload.heureFin);
    if (!heureDebut || !heureFin) throw new Error('Heures requises si aucun modèle n’est choisi.');
    if (pauseDebut && pauseFin && !_isPauseInsideShift_(heureDebut, heureFin, pauseDebut, pauseFin)) throw new Error('La pause doit être comprise dans le shift.');
    base = { heureDebut, heureFin, pauseDebut, pauseFin, isOff: false, nomModele: sanitizeText_(payload.nomModele || ''), couleurModele: sanitizeText_(payload.couleurModele || '#4ea1ff'), notes: sanitizeText_(payload.notes || '') };
  }

  const dates = enumerateDates_(dateDebut, dateFin)
    .filter(d => {
      const dow = d.getDay();
      const normalized = dow === 0 ? 7 : dow;
      return joursSemaine.indexOf(normalized) !== -1;
    })
    .map(d => toIsoDate_(d));
  if (!dates.length) throw new Error('Aucun jour à planifier avec ces critères.');

  let created = 0, updated = 0, skipped = 0;
  const warnings = [];
  targetUserIds.forEach(uid => {
    const user = getUserByIdRaw_(uid);
    if (!user || String(user.actif || '').toLowerCase() !== 'oui' || String(user.role || '') === APP.ROLES.ADMIN) {
      skipped += dates.length;
      warnings.push(uid + ': utilisateur invalide');
      return;
    }
    dates.forEach(date => {
      const existing = _getShiftForUserDate_(uid, date);
      if (existing && !overwrite) { skipped++; return; }
      const leave = _getApprovedLeaveForDate_(uid, date);
      if (leave) { skipped++; warnings.push(uid + ' : congé validé le ' + date); return; }
      apiAdminSaveShift(token, {
        idShift: existing ? existing.idShift : undefined,
        userId: uid,
        date: date,
        heureDebut: base.heureDebut,
        heureFin: base.heureFin,
        pauseDebut: base.pauseDebut,
        pauseFin: base.pauseFin,
        isOff: false,
        nomModele: base.nomModele,
        couleurModele: base.couleurModele,
        notes: base.notes
      });
      if (existing) updated++; else created++;
    });
  });

  logHistory_('', 'planning_bulk_assign', session.userId, 'admin', `Affectation multiple | créés=${created} maj=${updated} ignorés=${skipped}`, {});
  return { ok: true, created, updated, skipped, warnings };
}

/* ========= POINTAGE ========= */

function apiPointageClockIn(token, payload) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  const user = getUserByIdRaw_(session.userId);
  if (!user) throw new Error('Utilisateur introuvable.');
  payload = payload || {};

  const now = new Date();
  const date = toIsoDate_(now);
  const time = _timeStr_(now);
  const leave = _getApprovedLeaveForDate_(user.userId, date);
  if (leave) throw new Error('Pointage impossible pendant un congé validé.');

  const shift = _getShiftForUserDate_(user.userId, date);
  if (shift && shift.isOff === 'oui') throw new Error('Pointage impossible sur un jour OFF.');

  const existing = _getPointageForUserDate_(user.userId, date);
  if (existing && existing.heureArrivee) throw new Error('Vous avez déjà pointé votre arrivée aujourd’hui.');

  const retardMinutes = shift && shift.heureDebut ? _calcRetard_(time, shift.heureDebut) : 0;
  const dureePrevu = shift && shift.heureDebut && shift.heureFin
    ? _calcShiftNetDuration_(shift.heureDebut, shift.heureFin, shift.pauseDebut, shift.pauseFin)
    : 0;

  if (existing) {
    _updatePointageFields_(existing.idPointage, {
      heureArrivee: time,
      retardMinutes,
      statut: retardMinutes > 5 ? 'Retard' : 'En cours'
    });
    logHistory_('', 'pointage_arrivee', user.userId, user.role, `Arrivée ${time}${retardMinutes > 0 ? ' (retard ' + retardMinutes + ' min)' : ''}`, { cibleUserId: user.userId, cibleEquipe: String(user.equipe || '') });
    return { ok: true, heureArrivee: time, retardMinutes };
  }

  const idPointage = 'PTG-' + Utilities.getUuid().slice(0, 8).toUpperCase();
  appendRowObject_(P_SH.POINTAGE, {
    idPointage,
    userId: user.userId,
    nomAgent: user.nomComplet,
    equipe: user.equipe || '',
    date,
    heurePrevDebut: shift ? shift.heureDebut : '',
    heurePrevFin: shift ? shift.heureFin : '',
    heureArrivee: time,
    heureDepart: '',
    dureePrevu,
    dureeReelle: 0,
    heuresSup: 0,
    retardMinutes,
    statut: retardMinutes > 5 ? 'Retard' : 'En cours',
    notes: sanitizeText_(payload.notes || ''),
    valideAdmin: 'non',
    adminValidant: '',
    dateCreation: now
  });

  logHistory_('', 'pointage_arrivee', user.userId, user.role, `Arrivée ${time}${retardMinutes > 0 ? ' (retard ' + retardMinutes + ' min)' : ''}`, { cibleUserId: user.userId, cibleEquipe: String(user.equipe || '') });
  return { ok: true, heureArrivee: time, retardMinutes };
}

function apiPointageClockOut(token, payload) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  const user = getUserByIdRaw_(session.userId);
  if (!user) throw new Error('Utilisateur introuvable.');
  payload = payload || {};

  const now = new Date();
  const date = toIsoDate_(now);
  const time = _timeStr_(now);

  const ptg = _getPointageForUserDate_(user.userId, date);
  if (!ptg) throw new Error('Aucune arrivée pointée aujourd’hui.');
  if (!ptg.heureArrivee) throw new Error('Pointez d’abord votre arrivée.');
  if (ptg.heureDepart) throw new Error('Vous avez déjà pointé votre départ aujourd’hui.');

  const shift = _getShiftForUserDate_(user.userId, date);
  const dureePrevu = _resolveExpectedDurationForPointage_(ptg, shift);
  const dureeReelle = shift
    ? _calcWorkedNetDuration_(ptg.heureArrivee, time, shift.pauseDebut, shift.pauseFin)
    : _calcDuree_(ptg.heureArrivee, time);
  const heuresSup = Math.max(0, Number((dureeReelle - dureePrevu).toFixed(4)));

  _updatePointageFields_(ptg.idPointage, {
    heureDepart: time,
    dureePrevu,
    dureeReelle,
    heuresSup,
    statut: 'Présent',
    notes: payload.notes ? sanitizeText_(payload.notes) : ptg.notes
  });

  logHistory_('', 'pointage_depart', user.userId, user.role, `Départ ${time} | Durée: ${dureeReelle.toFixed(2)}h | H.Sup: ${heuresSup.toFixed(2)}h`, { cibleUserId: user.userId, cibleEquipe: String(user.equipe || '') });
  return { ok: true, heureDepart: time, dureeReelle, heuresSup };
}

function apiGetMyPointage(token, filters) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  const user = getUserByIdRaw_(session.userId);
  if (!user) throw new Error('Utilisateur introuvable.');
  filters = filters || {};

  const dateDebut = toIsoDateString_(filters.dateDebut || _startOfMonth_());
  const dateFin = toIsoDateString_(filters.dateFin || new Date());

  let rows = getRowsAsObjects_(P_SH.POINTAGE).map(normalizePointageRow_)
    .filter(p => normUpper_(p.userId) === normUpper_(user.userId) && p.date >= dateDebut && p.date <= dateFin)
    .map(p => _hydratePointageMetrics_(p));
  rows.sort((a, b) => b.date.localeCompare(a.date));

  const today = toIsoDate_(new Date());
  const todayShift = _getShiftForUserDate_(user.userId, today);
  const todayPointage = _hydratePointageMetrics_(_getPointageForUserDate_(user.userId, today));
  const totalHeuresSup = Number(rows.reduce((acc, p) => acc + num_(p.heuresSup), 0).toFixed(4));
  const totalRetards = rows.filter(p => num_(p.retardMinutes) > 5).length;

  const total = rows.length;
  const page = Math.max(1, Number(filters.page || 1));
  const pageSize = Math.max(5, Number(filters.pageSize || 30));
  rows = rows.slice((page - 1) * pageSize, page * pageSize);

  return { ok: true, items: rows, total, page, pageSize, todayShift, todayPointage, totalHeuresSup, totalRetards };
}

function apiAdminListPointage(token, filters) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  filters = filters || {};

  const dateDebut = toIsoDateString_(filters.dateDebut || _startOfMonth_());
  const dateFin = toIsoDateString_(filters.dateFin || new Date());
  const userId = normUpper_(filters.userId || '');
  const equipe = normUpper_(filters.equipe || '');
  const statut = String(filters.statut || '');

  let rows = getRowsAsObjects_(P_SH.POINTAGE).map(normalizePointageRow_)
    .filter(p => p.date >= dateDebut && p.date <= dateFin)
    .map(p => _hydratePointageMetrics_(p));
  if (userId) rows = rows.filter(p => normUpper_(p.userId) === userId);
  if (equipe) rows = rows.filter(p => normUpper_(p.equipe) === equipe);
  if (statut) rows = rows.filter(p => p.statut === statut);

  rows.sort((a, b) => b.date.localeCompare(a.date) || a.userId.localeCompare(b.userId));

  const total = rows.length;
  const page = Math.max(1, Number(filters.page || 1));
  const pageSize = Math.max(5, Number(filters.pageSize || 50));
  rows = rows.slice((page - 1) * pageSize, page * pageSize);

  return { ok: true, items: rows, total, page, pageSize };
}

function apiAdminAdjustPointage(token, payload) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  payload = payload || {};

  const idPointage = normTrim_(payload.idPointage);
  const heureArrivee = _normalizeTimeValue_(payload.heureArrivee);
  const heureDepart = _normalizeTimeValue_(payload.heureDepart);
  const notes = sanitizeText_(payload.notes || '');
  if (!idPointage) throw new Error('ID pointage requis.');

  const sh = getSheet_(P_SH.POINTAGE);
  const data = sh.getDataRange().getValues();
  const headers = data[0];
  const h = headerIndexMap_(headers);

  for (let i = 1; i < data.length; i++) {
    if (normTrim_(data[i][h.idPointage]) !== idPointage) continue;

    const row = data[i];
    const rowDate = toIsoDate_(row[h.date]);
    const userId = String(row[h.userId] || '');
    const prevDebut = _normalizeTimeValue_(row[h.heurePrevDebut]);
    const prevFin = _normalizeTimeValue_(row[h.heurePrevFin]);

    if (heureArrivee) row[h.heureArrivee] = heureArrivee;
    if (heureDepart) row[h.heureDepart] = heureDepart;
    if (notes) row[h.notes] = notes;

    const arr = _normalizeTimeValue_(row[h.heureArrivee]);
    const dep = _normalizeTimeValue_(row[h.heureDepart]);
    const shift = _getShiftForUserDate_(userId, rowDate);
    const expectedStart = (shift && shift.heureDebut) || prevDebut || '';
    const expectedEnd = (shift && shift.heureFin) || prevFin || '';
    const dureePrevu = _resolveExpectedDurationForPointage_({
      userId: userId,
      date: rowDate,
      dureePrevu: row[h.dureePrevu],
      heurePrevDebut: expectedStart || '',
      heurePrevFin: expectedEnd || ''
    }, shift);

    if (h.heurePrevDebut !== undefined) row[h.heurePrevDebut] = expectedStart || '';
    if (h.heurePrevFin !== undefined) row[h.heurePrevFin] = expectedEnd || '';
    row[h.dureePrevu] = dureePrevu;

    if (arr && dep) {
      const dureeReelle = shift
        ? _calcWorkedNetDuration_(arr, dep, shift.pauseDebut, shift.pauseFin)
        : _calcDuree_(arr, dep);
      row[h.dureeReelle] = dureeReelle;
      row[h.heuresSup] = Math.max(0, Number((dureeReelle - dureePrevu).toFixed(4)));
      row[h.retardMinutes] = expectedStart ? _calcRetard_(arr, expectedStart) : 0;
      row[h.statut] = 'Modifié admin';
    } else {
      row[h.dureeReelle] = 0;
      row[h.heuresSup] = 0;
      row[h.retardMinutes] = arr && expectedStart ? _calcRetard_(arr, expectedStart) : 0;
      row[h.statut] = arr ? (num_(row[h.retardMinutes]) > 5 ? 'Retard' : 'En cours') : 'Absent';
    }

    row[h.valideAdmin] = 'oui';
    row[h.adminValidant] = session.userId;
    sh.getRange(i + 1, 1, 1, headers.length).setValues([row]);

    logHistory_('', 'pointage_ajustement_admin', session.userId, 'admin',
      `Ajustement ${idPointage} | arr:${heureArrivee} dep:${heureDepart}`,
      { cibleUserId: userId, cibleEquipe: String(row[h.equipe] || '') });

    return { ok: true, item: _hydratePointageMetrics_(normalizePointageRow_(objectFromRow_(headers, row))) };
  }
  throw new Error('Pointage introuvable.');
}

function apiAdminRecalculatePointage(token, filters) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  requireAdmin_(session);
  filters = filters || {};

  const dateDebut = toIsoDateString_(filters.dateDebut || _startOfMonth_());
  const dateFin = toIsoDateString_(filters.dateFin || new Date());
  const userIdFilter = normUpper_(filters.userId || '');
  const equipeFilter = normUpper_(filters.equipe || '');

  const sh = getSheet_(P_SH.POINTAGE);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { ok: true, updated: 0 };

  const headers = values[0];
  const h = headerIndexMap_(headers);
  let updated = 0;

  for (let i = 1; i < values.length; i++) {
    const rowDate = toIsoDate_(values[i][h.date]);
    const rowUser = normUpper_(values[i][h.userId]);
    const rowEquipe = normUpper_(values[i][h.equipe]);

    if (!rowDate || rowDate < dateDebut || rowDate > dateFin) continue;
    if (userIdFilter && rowUser !== userIdFilter) continue;
    if (equipeFilter && rowEquipe !== equipeFilter) continue;

    const arr = _normalizeTimeValue_(values[i][h.heureArrivee]);
    const dep = _normalizeTimeValue_(values[i][h.heureDepart]);
    const prevDebut = _normalizeTimeValue_(values[i][h.heurePrevDebut]);
    const prevFin = _normalizeTimeValue_(values[i][h.heurePrevFin]);
    const shift = _getShiftForUserDate_(String(values[i][h.userId] || ''), rowDate);
    const expectedStart = (shift && shift.heureDebut) || prevDebut || '';
    const expectedEnd = (shift && shift.heureFin) || prevFin || '';

    const dureePrevu = _resolveExpectedDurationForPointage_({
      userId: String(values[i][h.userId] || ''),
      date: rowDate,
      dureePrevu: values[i][h.dureePrevu],
      heurePrevDebut: expectedStart || '',
      heurePrevFin: expectedEnd || ''
    }, shift);

    if (h.heurePrevDebut !== undefined) values[i][h.heurePrevDebut] = expectedStart || '';
    if (h.heurePrevFin !== undefined) values[i][h.heurePrevFin] = expectedEnd || '';
    values[i][h.dureePrevu] = dureePrevu;
    values[i][h.retardMinutes] = arr && expectedStart ? _calcRetard_(arr, expectedStart) : 0;

    if (arr && dep) {
      const dureeReelle = shift
        ? _calcWorkedNetDuration_(arr, dep, shift.pauseDebut, shift.pauseFin)
        : _calcDuree_(arr, dep);
      values[i][h.dureeReelle] = dureeReelle;
      values[i][h.heuresSup] = Math.max(0, Number((dureeReelle - dureePrevu).toFixed(4)));
      if (String(values[i][h.statut] || '') !== 'Modifié admin') values[i][h.statut] = 'Présent';
    } else {
      values[i][h.dureeReelle] = 0;
      values[i][h.heuresSup] = 0;
      if (arr) values[i][h.statut] = Number(values[i][h.retardMinutes] || 0) > 5 ? 'Retard' : 'En cours';
    }
    updated++;
  }

  if (updated > 0) sh.getRange(2, 1, values.length - 1, headers.length).setValues(values.slice(1));

  logHistory_('', 'pointage_recalcul_admin', session.userId, 'admin',
    `Recalcul pointage ${dateDebut}→${dateFin} | ${updated} ligne(s) mises à jour`,
    { cibleUserId: userIdFilter || '', cibleEquipe: equipeFilter || '' });

  return { ok: true, updated };
}

function apiGetOvertimeSummary(token, filters) {
  _ensurePlanningReady_();
  const session = requireSession_(token);
  const user = getUserByIdRaw_(session.userId);
  if (!user) throw new Error('Utilisateur introuvable.');
  const isAdmin = user.role === APP.ROLES.ADMIN;
  filters = filters || {};

  if (!isAdmin && filters.userId && normUpper_(filters.userId) !== normUpper_(user.userId)) {
    throw new Error('Accès non autorisé.');
  }

  const targetId = isAdmin ? (filters.userId ? normUpper_(filters.userId) : '') : normUpper_(user.userId);
  const annee = Number(filters.annee || new Date().getFullYear());

  let rows = getRowsAsObjects_(P_SH.POINTAGE).map(normalizePointageRow_)
    .filter(p => p.date && new Date(p.date + 'T00:00:00').getFullYear() === annee);
  if (targetId) rows = rows.filter(p => normUpper_(p.userId) === targetId);

  const byUserMonth = {};
  rows.forEach(p => {
    const key = normUpper_(p.userId) + '_' + p.date.slice(0, 7);
    if (!byUserMonth[key]) {
      byUserMonth[key] = {
        userId: p.userId,
        nomAgent: p.nomAgent,
        equipe: p.equipe,
        mois: p.date.slice(0, 7),
        totalHSup: 0,
        totalRetardMin: 0,
        nbJours: 0,
        nbRetards: 0
      };
    }
    byUserMonth[key].totalHSup += num_(p.heuresSup);
    byUserMonth[key].totalRetardMin += num_(p.retardMinutes);
    byUserMonth[key].nbJours++;
    if (num_(p.retardMinutes) > 5) byUserMonth[key].nbRetards++;
  });

  const summary = Object.keys(byUserMonth).map(k => {
    const e = byUserMonth[k];
    e.totalHSup = Number(e.totalHSup.toFixed(2));
    e.totalRetardMin = Math.round(e.totalRetardMin);
    return e;
  }).sort((a, b) => b.mois.localeCompare(a.mois) || a.userId.localeCompare(b.userId));

  return { ok: true, annee, summary };
}

/* ========= HELPERS PRIVÉS ========= */

function _getActiveUsers_(equipeUpper) {
  return getRowsAsObjects_(APP.SHEETS.USERS)
    .filter(u => String(u.actif || '').toLowerCase() === 'oui' && String(u.role || '') !== APP.ROLES.ADMIN)
    .filter(u => !equipeUpper || normUpper_(u.equipe || '') === equipeUpper)
    .map(u => ({
      userId: normTrim_(u.userId),
      nomAgent: normTrim_(u.nomComplet),
      equipe: normTrim_(u.equipe || '')
    }))
    .sort((a, b) => a.nomAgent.localeCompare(b.nomAgent, 'fr'));
}

function _buildShiftMap_() {
  const out = {};
  getRowsAsObjects_(P_SH.SHIFTS).map(normalizeShiftRow_).forEach(s => {
    out[normUpper_(s.userId) + '_' + s.date] = s;
  });
  return out;
}

function _indexShiftsByUser_(shifts) {
  const out = {};
  (shifts || []).forEach(s => {
    const uid = normUpper_(s.userId);
    if (!out[uid]) out[uid] = {};
    out[uid][s.date] = s;
  });
  return out;
}

function _getShiftsForUser_(userId, dateDebut, dateFin) {
  return getRowsAsObjects_(P_SH.SHIFTS).map(normalizeShiftRow_)
    .filter(s => normUpper_(s.userId) === normUpper_(userId) && s.date >= dateDebut && s.date <= dateFin)
    .sort((a, b) => a.date.localeCompare(b.date));
}

function _getLeavesForUser_(userId, dateDebut, dateFin) {
  return getRowsAsObjects_(APP.SHEETS.REQUESTS).map(normalizeRequestRow_)
    .filter(r => normUpper_(r.userId) === normUpper_(userId) && r.dateFin >= dateDebut && r.dateDebut <= dateFin && (r.statut === APP.STATUS.APPROVED || r.statut === APP.STATUS.PENDING));
}

function _getLeavesForEquipe_(equipe, dateDebut, dateFin) {
  return getRowsAsObjects_(APP.SHEETS.REQUESTS).map(normalizeRequestRow_)
    .filter(r => normUpper_(r.equipe) === normUpper_(equipe) && r.dateFin >= dateDebut && r.dateDebut <= dateFin && (r.statut === APP.STATUS.APPROVED || r.statut === APP.STATUS.PENDING));
}

function _getLeavesForPeriod_(dateDebut, dateFin, filters) {
  filters = filters || {};
  const equipe = normUpper_(filters.equipe || '');
  const userId = normUpper_(filters.userId || '');
  let leaves = getRowsAsObjects_(APP.SHEETS.REQUESTS).map(normalizeRequestRow_)
    .filter(r => r.dateFin >= dateDebut && r.dateDebut <= dateFin && (r.statut === APP.STATUS.APPROVED || r.statut === APP.STATUS.PENDING));
  if (equipe) leaves = leaves.filter(r => normUpper_(r.equipe) === equipe);
  if (userId) leaves = leaves.filter(r => normUpper_(r.userId) === userId);
  return leaves;
}

function _getUserLeaveForDate_(userId, date) {
  return getRowsAsObjects_(APP.SHEETS.REQUESTS).map(normalizeRequestRow_)
    .find(r => normUpper_(r.userId) === normUpper_(userId) && r.dateDebut <= date && r.dateFin >= date && (r.statut === APP.STATUS.APPROVED || r.statut === APP.STATUS.PENDING)) || null;
}

function _getApprovedLeaveForDate_(userId, date) {
  return getRowsAsObjects_(APP.SHEETS.REQUESTS).map(normalizeRequestRow_)
    .find(r => normUpper_(r.userId) === normUpper_(userId) && r.dateDebut <= date && r.dateFin >= date && r.statut === APP.STATUS.APPROVED) || null;
}

function _getShiftForUserDate_(userId, date) {
  const rows = getRowsAsObjects_(P_SH.SHIFTS).map(normalizeShiftRow_)
    .filter(s => normUpper_(s.userId) === normUpper_(userId) && s.date === date)
    .sort((a, b) => (a.isOff === 'oui' ? 1 : 0) - (b.isOff === 'oui' ? 1 : 0));
  return rows.length ? rows[0] : null;
}

function _getPointageForUserDate_(userId, date) {
  const rows = getRowsAsObjects_(P_SH.POINTAGE).map(normalizePointageRow_)
    .filter(p => normUpper_(p.userId) === normUpper_(userId) && p.date === date);
  return rows.length ? rows[0] : null;
}

function _updatePointageFields_(idPointage, patch) {
  const sh = getSheet_(P_SH.POINTAGE);
  const data = sh.getDataRange().getValues();
  const h = headerIndexMap_(data[0]);
  for (let i = 1; i < data.length; i++) {
    if (normTrim_(data[i][h.idPointage]) === normTrim_(idPointage)) {
      Object.keys(patch).forEach(k => { if (h[k] !== undefined) data[i][h[k]] = patch[k]; });
      sh.getRange(i + 1, 1, 1, data[0].length).setValues([data[i]]);
      return true;
    }
  }
  return false;
}

function normalizeShiftRow_(r) {
  const heureDebut = _normalizeTimeValue_(r.heureDebut);
  const heureFin = _normalizeTimeValue_(r.heureFin);
  const pauseDebut = _normalizeTimeValue_(r.pauseDebut);
  const pauseFin = _normalizeTimeValue_(r.pauseFin);
  const isOff = String(r.isOff || '').toLowerCase() === 'oui' || String(r.statut || '').toUpperCase() === 'OFF';
  return {
    idShift: String(r.idShift || '').trim(),
    userId: String(r.userId || '').trim(),
    nomAgent: String(r.nomAgent || '').trim(),
    equipe: String(r.equipe || '').trim(),
    date: toIsoDate_(r.date),
    heureDebut: isOff ? '' : heureDebut,
    heureFin: isOff ? '' : heureFin,
    pauseDebut: isOff ? '' : pauseDebut,
    pauseFin: isOff ? '' : pauseFin,
    isOff: isOff ? 'oui' : 'non',
    dureePrevue: num_(r.dureePrevue) || (isOff ? 0 : _calcShiftNetDuration_(heureDebut, heureFin, pauseDebut, pauseFin)),
    nomModele: String(r.nomModele || '').trim(),
    couleurModele: String(r.couleurModele || '#4ea1ff').trim(),
    statut: isOff ? 'OFF' : String(r.statut || 'Planifié').trim(),
    notes: String(r.notes || '').trim(),
    dateCreation: toIsoDateTime_(r.dateCreation),
    creePar: String(r.creePar || '').trim()
  };
}

function normalizePointageRow_(r) {
  return {
    idPointage: String(r.idPointage || '').trim(),
    userId: String(r.userId || '').trim(),
    nomAgent: String(r.nomAgent || '').trim(),
    equipe: String(r.equipe || '').trim(),
    date: toIsoDate_(r.date),
    heurePrevDebut: _normalizeTimeValue_(r.heurePrevDebut),
    heurePrevFin: _normalizeTimeValue_(r.heurePrevFin),
    heureArrivee: _normalizeTimeValue_(r.heureArrivee),
    heureDepart: _normalizeTimeValue_(r.heureDepart),
    dureePrevu: num_(r.dureePrevu),
    dureeReelle: num_(r.dureeReelle),
    heuresSup: num_(r.heuresSup),
    retardMinutes: num_(r.retardMinutes),
    statut: String(r.statut || '').trim(),
    notes: String(r.notes || '').trim(),
    valideAdmin: String(r.valideAdmin || 'non').trim(),
    adminValidant: String(r.adminValidant || '').trim(),
    dateCreation: toIsoDateTime_(r.dateCreation)
  };
}

function _normalizeModeleRow_(r) {
  return {
    idModele: String(r.idModele || '').trim(),
    nom: String(r.nom || '').trim(),
    heureDebut: _normalizeTimeValue_(r.heureDebut),
    heureFin: _normalizeTimeValue_(r.heureFin),
    couleur: String(r.couleur || '#4ea1ff').trim(),
    description: String(r.description || '').trim(),
    actif: String(r.actif || 'oui').trim().toLowerCase() === 'non' ? 'non' : 'oui'
  };
}

function _normalizePlageRow_(r) {
  return {
    idPlage: String(r.idPlage || '').trim(),
    equipe: String(r.equipe || '').trim(),
    nom: String(r.nom || '').trim(),
    heureDebut: _normalizeTimeValue_(r.heureDebut),
    heureFin: _normalizeTimeValue_(r.heureFin),
    niveauFlux: String(r.niveauFlux || 'Normal').trim(),
    joursActifs: _normalizeDaysList_(r.joursActifs || '1,2,3,4,5'),
    effectifMin: Math.max(1, num_(r.effectifMin)),
    couleur: String(r.couleur || '#888888').trim(),
    actif: String(r.actif || 'oui').trim().toLowerCase() === 'non' ? 'non' : 'oui'
  };
}

function _normalizeTimeValue_(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'string') {
    const s = v.trim();
    if (/^\d{2}:\d{2}$/.test(s)) return s;
    if (/^\d{1}:\d{2}$/.test(s)) return ('0' + s).slice(-5);
    const m = s.match(/(\d{1,2}):(\d{2})/);
    if (m) return ('0' + Number(m[1])).slice(-2) + ':' + m[2];
    const d1 = parseAnyDate_(s);
    if (d1) return Utilities.formatDate(d1, APP.TZ, 'HH:mm');
    return '';
  }
  const d = parseAnyDate_(v);
  if (d) return Utilities.formatDate(d, APP.TZ, 'HH:mm');
  return '';
}

function _normalizePauseRange_(pauseDebut, pauseFin) {
  const start = _normalizeTimeValue_(pauseDebut);
  let end = _normalizeTimeValue_(pauseFin);
  if (start && !end) end = _addMinutesToTime_(start, 60);
  return { pauseDebut: start, pauseFin: end };
}

function _addMinutesToTime_(hm, minsToAdd) {
  const base = _timeToMin_(hm);
  if (base == null || isNaN(base)) return '';
  let total = (base + Number(minsToAdd || 0)) % (24 * 60);
  if (total < 0) total += 24 * 60;
  const h = Math.floor(total / 60);
  const m = total % 60;
  return ('0' + h).slice(-2) + ':' + ('0' + m).slice(-2);
}

function _normalizeDaysList_(v) {
  const arr = String(v || '').split(',').map(s => Number(String(s).trim())).filter(n => n >= 1 && n <= 7);
  const uniq = [];
  arr.forEach(n => { if (uniq.indexOf(n) === -1) uniq.push(n); });
  return uniq.length ? uniq.join(',') : '1,2,3,4,5';
}

function _calcDuree_(heureDebut, heureFin) {
  if (!heureDebut || !heureFin) return 0;
  const a = _timeToMin_(heureDebut);
  const b = _timeToMin_(heureFin);
  let mins = b - a;
  if (mins < 0) mins += 24 * 60;
  return Number((mins / 60).toFixed(4));
}

function _calcShiftNetDuration_(heureDebut, heureFin, pauseDebut, pauseFin) {
  let dur = _calcDuree_(heureDebut, heureFin);
  if (pauseDebut && pauseFin) dur = Math.max(0, dur - _calcDuree_(pauseDebut, pauseFin));
  return Number(dur.toFixed(2));
}

function _calcRetard_(heureReelle, heurePrevue) {
  if (!heureReelle || !heurePrevue) return 0;
  return Math.max(0, _timeToMin_(heureReelle) - _timeToMin_(heurePrevue));
}

function _timeToMin_(hm) {
  const p = String(hm || '00:00').split(':');
  const h = Number(p[0] || 0);
  const m = Number(p[1] || 0);
  return h * 60 + m;
}

function _isPauseInsideShift_(shiftStart, shiftEnd, pauseStart, pauseEnd) {
  if (!pauseStart || !pauseEnd) return true;
  const s1 = _timeToMin_(shiftStart);
  let s2 = _timeToMin_(shiftEnd);
  let p1 = _timeToMin_(pauseStart);
  let p2 = _timeToMin_(pauseEnd);
  if (s2 <= s1) s2 += 1440;
  if (p1 < s1) p1 += 1440;
  if (p2 <= p1) p2 += 1440;
  return p1 >= s1 && p2 <= s2;
}

function _buildDayStatus_(shift, pointage, leave, now) {
  if (leave && leave.statut === APP.STATUS.APPROVED) {
    return { key: 'en_conge', label: 'En congé', details: leave.typeConge || '' };
  }
  if (leave && leave.statut === APP.STATUS.PENDING) {
    return { key: 'conge_attente', label: 'Congé en attente', details: leave.typeConge || '' };
  }
  if (shift && shift.isOff === 'oui') {
    return { key: 'off', label: 'OFF', details: 'Jour de repos' };
  }
  if (!shift) {
    return { key: 'non_planifie', label: 'Non planifié', details: '' };
  }
  if (pointage && pointage.heureDepart) {
    return { key: 'reparti', label: 'Reparti', details: pointage.heureDepart };
  }
  if (pointage && pointage.heureArrivee) {
    if (num_(pointage.retardMinutes) > 5) return { key: 'present_retard', label: 'Présent (retard)', details: pointage.heureArrivee };
    return { key: 'present', label: 'Présent', details: pointage.heureArrivee };
  }

  const nowMin = now.getHours() * 60 + now.getMinutes();
  const startMin = shift.heureDebut ? _timeToMin_(shift.heureDebut) : null;
  const endMin = shift.heureFin ? _timeToMin_(shift.heureFin) : null;
  if (startMin == null || endMin == null) return { key: 'non_planifie', label: 'Non planifié', details: '' };

  if (nowMin < startMin - 15) {
    return { key: 'a_venir', label: 'À venir', details: shift.heureDebut + '–' + shift.heureFin };
  }
  return { key: 'absent', label: 'Absent', details: shift.heureDebut + '–' + shift.heureFin };
}


function _normalizeSwapRow_(r) {
  return {
    idSwap: String(r.idSwap || '').trim(),
    requesterUserId: String(r.requesterUserId || '').trim(),
    requesterNom: String(r.requesterNom || '').trim(),
    requesterEquipe: String(r.requesterEquipe || '').trim(),
    targetUserId: String(r.targetUserId || '').trim(),
    targetNom: String(r.targetNom || '').trim(),
    targetEquipe: String(r.targetEquipe || '').trim(),
    dateShift: toIsoDate_(r.dateShift),
    requesterShiftId: String(r.requesterShiftId || '').trim(),
    targetShiftId: String(r.targetShiftId || '').trim(),
    requesterComment: String(r.requesterComment || '').trim(),
    adminComment: String(r.adminComment || '').trim(),
    statut: String(r.statut || '').trim(),
    dateCreation: toIsoDateTime_(r.dateCreation),
    dateDecision: toIsoDateTime_(r.dateDecision),
    decidedBy: String(r.decidedBy || '').trim()
  };
}

function _getShiftById_(idShift) {
  if (!idShift) return null;
  const rows = getRowsAsObjects_(P_SH.SHIFTS).map(normalizeShiftRow_);
  return rows.find(r => normTrim_(r.idShift) === normTrim_(idShift)) || null;
}

function _hydrateSwapTimeline_(swap) {
  if (!swap) return swap;
  const out = Object.assign({}, swap);
  out.requesterShift = _getShiftById_(swap.requesterShiftId);
  out.targetShift = _getShiftById_(swap.targetShiftId);
  return out;
}

function _calcWorkedNetDuration_(heureDebut, heureFin, pauseDebut, pauseFin) {
  let dur = _calcDuree_(heureDebut, heureFin);
  if (pauseDebut && pauseFin) {
    dur = Math.max(0, dur - _calcOverlapHours_(heureDebut, heureFin, pauseDebut, pauseFin));
  }
  return Number(dur.toFixed(4));
}

function _resolveExpectedDurationForPointage_(pointage, shift) {
  const stored = num_(pointage && pointage.dureePrevu);
  const refShift = shift || (pointage && pointage.userId && pointage.date ? _getShiftForUserDate_(pointage.userId, pointage.date) : null);
  if (refShift && refShift.heureDebut && refShift.heureFin) {
    return _calcShiftNetDuration_(refShift.heureDebut, refShift.heureFin, refShift.pauseDebut, refShift.pauseFin);
  }
  if (stored > 0) return Number(stored.toFixed(4));
  if (pointage && pointage.heurePrevDebut && pointage.heurePrevFin) {
    return _calcDuree_(pointage.heurePrevDebut, pointage.heurePrevFin);
  }
  return 0;
}

function _hydratePointageMetrics_(pointage) {
  if (!pointage) return null;
  const out = Object.assign({}, pointage);
  const shift = (out.userId && out.date) ? _getShiftForUserDate_(out.userId, out.date) : null;

  if ((!out.heurePrevDebut || !out.heurePrevFin) && shift && shift.heureDebut && shift.heureFin) {
    out.heurePrevDebut = shift.heureDebut;
    out.heurePrevFin = shift.heureFin;
  }

  out.dureePrevu = _resolveExpectedDurationForPointage_(out, shift);

  const expectedStart = (shift && shift.heureDebut) || out.heurePrevDebut || '';
  const expectedEnd = (shift && shift.heureFin) || out.heurePrevFin || '';

  if (!out.heurePrevDebut && expectedStart) out.heurePrevDebut = expectedStart;
  if (!out.heurePrevFin && expectedEnd) out.heurePrevFin = expectedEnd;

  if (out.heureArrivee) {
    out.retardMinutes = expectedStart ? _calcRetard_(out.heureArrivee, expectedStart) : num_(out.retardMinutes);
  }

  if (out.heureArrivee && out.heureDepart) {
    out.dureeReelle = shift && shift.heureDebut && shift.heureFin
      ? _calcWorkedNetDuration_(out.heureArrivee, out.heureDepart, shift.pauseDebut, shift.pauseFin)
      : _calcDuree_(out.heureArrivee, out.heureDepart);
    out.heuresSup = Math.max(0, Number((out.dureeReelle - out.dureePrevu).toFixed(4)));
  }
  return out;
}

function _calcOverlapHours_(aStart, aEnd, bStart, bEnd) {
  if (!aStart || !aEnd || !bStart || !bEnd) return 0;
  let a1 = _timeToMin_(aStart), a2 = _timeToMin_(aEnd), b1 = _timeToMin_(bStart), b2 = _timeToMin_(bEnd);
  if (a2 <= a1) a2 += 1440;
  if (b1 < a1) b1 += 1440;
  if (b2 <= b1) b2 += 1440;
  const overlap = Math.max(0, Math.min(a2, b2) - Math.max(a1, b1));
  return Number((overlap / 60).toFixed(4));
}

function _mondayOfWeek_(d) {
  const date = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const day = date.getDay();
  const diff = (day === 0) ? -6 : 1 - day;
  date.setDate(date.getDate() + diff);
  return toIsoDate_(date);
}

function _addDaysToIso_(isoDate, n) {
  const d = new Date(isoDate + 'T00:00:00');
  d.setDate(d.getDate() + n);
  return toIsoDate_(d);
}

function _startOfMonth_() {
  const d = new Date();
  return Utilities.formatDate(new Date(d.getFullYear(), d.getMonth(), 1), APP.TZ, 'yyyy-MM-dd');
}

function _timeStr_(d) {
  return Utilities.formatDate(d, APP.TZ, 'HH:mm');
}
