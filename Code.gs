/**
 * SUIVIS JEFF : backend Google Apps Script
 * ------------------------------------------------------------
 * À coller dans le Sheet « Suivis Jeff » : Extensions > Apps Script.
 * 1. Exécute setup() une fois (autorise l'accès). Ça optimise le Sheet :
 *    onglets Suivis / Listes / Journal / Config, cases à cocher,
 *    listes déroulantes, couleurs, migration des lignes existantes.
 * 2. Menu « 🔐 Suivis > Définir un mot de passe » pour chaque personne.
 * 3. Déployer > Nouveau déploiement > Application Web
 *    Exécuter en tant que : Moi | Accès : Tout le monde.
 *    (L'API refuse tout sans connexion valide : nom + mot de passe.)
 * 4. Copie l'URL /exec dans l'app (Réglages).
 *
 * Sécurité :
 * - Mots de passe jamais stockés en clair : sel aléatoire + SHA-256 itéré 1 500 fois.
 * - Connexion = jeton de session aléatoire (plus de 350 bits), valide 30 jours, révocable.
 * - 5 essais ratés = compte bloqué 15 minutes.
 * - L'auteur de chaque modification vient de la session, jamais du navigateur.
 */

const TABS = { SUIVIS: 'Suivis', LISTES: 'Listes', JOURNAL: 'Journal', CONFIG: 'Config', USERS: 'Utilisateurs' };
const USER_COLS = ['nom', 'courriel', 'role', 'actif', 'sel', 'hash', 'cree_le', 'derniere_connexion'];
const SESSION_DAYS = 30;
const MAX_FAILS = 5;
const LOCK_MINUTES = 15;
const HASH_ROUNDS = 1500;

const COLS = [
  'id', 'fait', 'priorite', 'type', 'dossier', 'action', 'responsable', 'contact',
  'courriel_contact', 'ouvert_le', 'echeance', 'statut', 'notes', 'source',
  'cree_par', 'cree_le', 'maj_le', 'maj_par', 'ferme_le', 'ordre'
];
const DATE_COLS = ['ouvert_le', 'echeance'];
const DATETIME_COLS = ['cree_le', 'maj_le', 'ferme_le'];
const EDITABLE = ['fait', 'priorite', 'type', 'dossier', 'action', 'responsable', 'contact',
  'courriel_contact', 'ouvert_le', 'echeance', 'statut', 'notes', 'source', 'ordre'];

const LISTES = {
  statuts: ['À faire', 'En cours', 'En attente', 'À vérifier', 'Fait', 'Archivé'],
  types: ['On attend après moi', 'Promesse', 'On me doit', 'Calendrier', 'Tâche'],
  priorites: ['Haute', 'Moyenne', 'Basse'],
  personnes: [
    ['Jeff', 'jefft@eckinox.ca'],
    ['Cindy', 'comptabilite@eckinox.ca'],
    ['Marilou', 'marilou@eckinox.ca'],
    ['Rossana', 'rossanag@eckinox.ca'],
    ['Gabrielle', 'gabrielled@eckinox.ca'],
    ['Mathieu', 'mathieum@eckinox.ca'],
    ['Claude', '']
  ]
};

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz_() { return ss_().getSpreadsheetTimeZone() || 'America/Toronto'; }

/* ============================ SETUP ============================ */

function setup() {
  const ss = ss_();
  ss.setSpreadsheetTimeZone('America/Toronto');

  // --- Listes
  const listes = getOrCreate_(TABS.LISTES);
  listes.clear();
  const maxLen = Math.max(LISTES.statuts.length, LISTES.types.length, LISTES.priorites.length, LISTES.personnes.length);
  const rows = [['statuts', 'types', 'priorites', 'personnes', 'courriels', 'dossiers']];
  const dossiers = [];
  for (let i = 0; i < maxLen; i++) {
    rows.push([
      LISTES.statuts[i] || '', LISTES.types[i] || '', LISTES.priorites[i] || '',
      (LISTES.personnes[i] || [''])[0], (LISTES.personnes[i] || ['', ''])[1], ''
    ]);
  }
  listes.getRange(1, 1, rows.length, 6).setValues(rows);
  styleHeader_(listes, 6);

  // --- Config
  const config = getOrCreate_(TABS.CONFIG);
  config.clear();
  const cfg = [
    ['cle', 'valeur', 'description'],
    ['RELANCE_JOURS', 5, 'Jours ouvrables avant qu\'un item soit marqué « dort »'],
    ['NEXT_ID', 1, 'Prochain numéro SUI-xxxx (géré automatiquement)']
  ];
  config.getRange(1, 1, cfg.length, 3).setValues(cfg);
  styleHeader_(config, 3);
  config.setColumnWidth(1, 160); config.setColumnWidth(2, 320); config.setColumnWidth(3, 420);

  // --- Journal
  const journal = getOrCreate_(TABS.JOURNAL);
  if (journal.getLastRow() === 0) {
    journal.getRange(1, 1, 1, 7).setValues([['horodatage', 'id', 'action', 'champ', 'avant', 'apres', 'par']]);
  }
  styleHeader_(journal, 7);

  // --- Suivis (+ migration de l'ancien format)
  let suivis = ss.getSheetByName(TABS.SUIVIS);
  let migrated = [];
  if (!suivis) {
    const old = ss.getSheets().find(s => String(s.getRange(1, 1).getValue()).trim() === 'Fait');
    if (old) migrated = migrateOld_(old);
    suivis = ss.insertSheet(TABS.SUIVIS, 0);
    suivis.getRange(1, 1, 1, COLS.length).setValues([COLS]);
    if (migrated.length) suivis.getRange(2, 1, migrated.length, COLS.length).setValues(migrated);
    if (old) { old.setName('Ancien format (archive)'); old.hideSheet(); }
  }
  migrated.forEach(r => dossiers.indexOf(r[4]) < 0 && r[4] && dossiers.push(r[4]));
  if (dossiers.length) listes.getRange(2, 6, dossiers.length, 1).setValues(dossiers.map(d => [d]));

  formatSuivis_(suivis);

  const maxNum = getAll_().reduce((m, it) => Math.max(m, parseInt(String(it.id).replace(/\D/g, ''), 10) || 0), 0);
  setConfig_('NEXT_ID', maxNum + 1);

  // --- Utilisateurs (sans mot de passe au départ : personne ne peut se connecter)
  const users = getOrCreate_(TABS.USERS);
  if (users.getLastRow() === 0) {
    users.getRange(1, 1, 1, USER_COLS.length).setValues([USER_COLS]);
    const seedUsers = LISTES.personnes.filter(p => p[0] !== 'Claude')
      .map(p => [p[0], p[1], p[0] === 'Jeff' ? 'admin' : 'membre', true, '', '', new Date(), '']);
    users.getRange(2, 1, seedUsers.length, USER_COLS.length).setValues(seedUsers);
  }
  styleHeader_(users, USER_COLS.length);
  users.getRange(2, 4, 200).insertCheckboxes();
  users.getRange(2, 3, 200).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['admin', 'membre'], true).build());
  users.hideColumns(5, 2); // sel + hash : cachés
  [140, 220, 90, 60].forEach((w, i) => users.setColumnWidth(i + 1, w));
  const prot = users.protect().setDescription('Utilisateurs : admin seulement');
  prot.removeEditors(prot.getEditors()); prot.addEditor(Session.getEffectiveUser());
  if (prot.canDomainEdit()) prot.setDomainEdit(false);
  config.protect().setDescription('Config : admin seulement').setWarningOnly(true);

  ss.setActiveSheet(suivis);
  SpreadsheetApp.getUi().alert('Sheet optimisé ✅\n\nProchaine étape : menu « 🔐 Suivis > Définir un mot de passe » pour chaque personne, puis Déployer > Application Web.');
}

/* ============================ MENU ADMIN ============================ */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('🔐 Suivis')
    .addItem('Définir un mot de passe', 'definirMotDePasseUi')
    .addItem('Déconnecter une personne partout', 'deconnecterUi')
    .addItem('Déconnecter tout le monde', 'deconnecterToutLeMonde')
    .addSeparator()
    .addItem('Optimiser / réparer le Sheet (setup)', 'setup')
    .addToUi();
}

function definirMotDePasseUi() {
  const ui = SpreadsheetApp.getUi();
  const r1 = ui.prompt('Définir un mot de passe', 'Nom de la personne (tel qu\'il apparaît dans l\'onglet Utilisateurs) :', ui.ButtonSet.OK_CANCEL);
  if (r1.getSelectedButton() !== ui.Button.OK) return;
  const nom = r1.getResponseText().trim();
  if (!findUser_(nom)) { ui.alert('Introuvable : ' + nom + '. Ajoute la personne dans l\'onglet Utilisateurs.'); return; }
  const r2 = ui.prompt('Mot de passe pour ' + nom, 'Minimum 10 caractères. La personne pourra le changer dans l\'app.', ui.ButtonSet.OK_CANCEL);
  if (r2.getSelectedButton() !== ui.Button.OK) return;
  const pw = r2.getResponseText();
  const err = pwPolicy_(pw);
  if (err) { ui.alert(err); return; }
  setPassword_(nom, pw);
  revokeSessions_(nom);
  ui.alert('✅ Mot de passe défini pour ' + nom + '. Ses anciennes sessions sont fermées.');
}

function deconnecterUi() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Déconnecter', 'Nom de la personne :', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const n = revokeSessions_(r.getResponseText().trim());
  ui.alert(n + ' session(s) fermée(s).');
}

function deconnecterToutLeMonde() {
  const props = PropertiesService.getScriptProperties();
  Object.keys(props.getProperties()).filter(k => k.indexOf('sess_') === 0).forEach(k => props.deleteProperty(k));
  SpreadsheetApp.getUi().alert('Tout le monde est déconnecté.');
}

function migrateOld_(sheet) {
  const v = sheet.getDataRange().getValues();
  const h = v[0].map(x => String(x).trim());
  const idx = name => h.indexOf(name);
  const now = new Date();
  return v.slice(1).filter(r => r.join('').trim()).map((r, i) => {
    const contactRaw = String(r[idx('Contact externe')] || '');
    const mail = (contactRaw.match(/[\w.+-]+@[\w.-]+\.\w+/) || [''])[0];
    const contact = contactRaw.replace(/\(.*?\)/g, '').trim();
    const fait = /^(oui|true|x|1)$/i.test(String(r[idx('Fait')]).trim());
    const num = parseInt(r[idx('#')], 10) || (i + 1);
    const row = {
      id: 'SUI-' + ('000' + num).slice(-4),
      fait: fait,
      priorite: r[idx('Priorité')] || 'Moyenne',
      type: r[idx('Type')] || 'Tâche',
      dossier: r[idx('Dossier')] || '',
      action: r[idx('Livrable / action')] || '',
      responsable: String(r[idx('Responsable')] || ''),
      contact: contact,
      courriel_contact: mail,
      ouvert_le: toDate_(r[idx('Ouvert le')]),
      echeance: toDate_(r[idx('Échéance')]),
      statut: fait ? 'Fait' : (r[idx('Statut')] || 'À faire'),
      notes: r[idx('Notes / prochaine étape')] || '',
      source: 'Registre Claude',
      cree_par: 'Claude', cree_le: now, maj_le: now, maj_par: 'Claude',
      ferme_le: fait ? now : '', ordre: (i + 1) * 10
    };
    return COLS.map(c => row[c] === undefined ? '' : row[c]);
  });
}

function formatSuivis_(sh) {
  const c = name => COLS.indexOf(name) + 1;
  const last = 1000;
  styleHeader_(sh, COLS.length);
  sh.setFrozenRows(1);
  sh.setFrozenColumns(2);

  sh.getRange(2, c('fait'), last).insertCheckboxes();
  const listes = ss_().getSheetByName(TABS.LISTES);
  const dv = col => SpreadsheetApp.newDataValidation()
    .requireValueInRange(listes.getRange(2, col, 100), true).setAllowInvalid(true).build();
  sh.getRange(2, c('statut'), last).setDataValidation(dv(1));
  sh.getRange(2, c('type'), last).setDataValidation(dv(2));
  sh.getRange(2, c('priorite'), last).setDataValidation(dv(3));
  sh.getRange(2, c('responsable'), last).setDataValidation(dv(4));
  sh.getRange(2, c('dossier'), last).setDataValidation(dv(6));

  DATE_COLS.forEach(n => sh.getRange(2, c(n), last).setNumberFormat('yyyy-mm-dd'));
  DATETIME_COLS.forEach(n => sh.getRange(2, c(n), last).setNumberFormat('yyyy-mm-dd hh:mm'));
  sh.getRange(2, c('action'), last).setWrap(true);
  sh.getRange(2, c('notes'), last).setWrap(true);

  const widths = { id: 90, fait: 50, priorite: 85, type: 160, dossier: 190, action: 380, responsable: 110,
    contact: 170, courriel_contact: 200, ouvert_le: 95, echeance: 95, statut: 105, notes: 320, source: 150,
    cree_par: 90, cree_le: 125, maj_le: 125, maj_par: 90, ferme_le: 125, ordre: 60 };
  Object.keys(widths).forEach(k => sh.setColumnWidth(c(k), widths[k]));

  // Couleurs : fermé = gris barré, en retard = rouge, aujourd'hui = orange, priorité haute = pastille
  const all = sh.getRange(2, 1, last, COLS.length);
  const B = 'B2', K = '$K2', L = '$L2';
  const rules = [
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=OR($B2=TRUE,' + L + '="Fait",' + L + '="Archivé")')
      .setFontColor('#9aa3b2').setStrikethrough(true).setBackground('#f3f4f6').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($B2<>TRUE,' + K + '<>"",' + K + '<TODAY())')
      .setBackground('#ffe4e6').setFontColor('#9f1239').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($B2<>TRUE,' + K + '=TODAY())')
      .setBackground('#fff4dd').setRanges([all]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Haute')
      .setBackground('#fecdd3').setFontColor('#9f1239').setBold(true).setRanges([sh.getRange(2, c('priorite'), last)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Moyenne')
      .setBackground('#fef3c7').setFontColor('#92400e').setRanges([sh.getRange(2, c('priorite'), last)]).build(),
    SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Basse')
      .setBackground('#dbeafe').setFontColor('#1e40af').setRanges([sh.getRange(2, c('priorite'), last)]).build()
  ];
  sh.setConditionalFormatRules(rules);
  if (!sh.getFilter()) sh.getRange(1, 1, Math.max(sh.getLastRow(), 2), COLS.length).createFilter();
}

function styleHeader_(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function getOrCreate_(name) { return ss_().getSheetByName(name) || ss_().insertSheet(name); }

/* ============================ DONNÉES ============================ */

function getAll_() {
  const sh = ss_().getSheetByName(TABS.SUIVIS);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues();
  return v.map((r, i) => rowToObj_(r, i + 2)).filter(o => o.id);
}

function rowToObj_(r, rowNum) {
  const o = { _row: rowNum };
  COLS.forEach((c, i) => {
    let val = r[i];
    if (val instanceof Date) val = Utilities.formatDate(val, tz_(), DATE_COLS.indexOf(c) >= 0 ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ss");
    o[c] = val;
  });
  o.fait = o.fait === true || o.fait === 'TRUE';
  return o;
}

function findRow_(id) {
  const sh = ss_().getSheetByName(TABS.SUIVIS);
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function toDate_(v) {
  if (!v) return '';
  if (v instanceof Date) return v;
  const m = String(v).match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : '';
}

function readConfig_() {
  const sh = ss_().getSheetByName(TABS.CONFIG);
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => out[r[0]] = r[1]);
  return out;
}

function setConfig_(key, value) {
  const sh = ss_().getSheetByName(TABS.CONFIG);
  const keys = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < keys.length; i++) if (keys[i][0] === key) { sh.getRange(i + 2, 2).setValue(value); return; }
  sh.appendRow([key, value, '']);
}

function readListes_() {
  const sh = ss_().getSheetByName(TABS.LISTES);
  const v = sh.getDataRange().getValues().slice(1);
  const col = i => v.map(r => r[i]).filter(x => x !== '');
  return {
    statuts: col(0), types: col(1), priorites: col(2),
    personnes: readUsersPublic_(),
    dossiers: col(5)
  };
}

function readUsersPublic_() {
  const sh = ss_().getSheetByName(TABS.USERS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues()
    .filter(r => r[0] && r[3] === true).map(r => ({ nom: r[0], courriel: r[1] }));
}

function log_(id, action, champ, avant, apres, par) {
  ss_().getSheetByName(TABS.JOURNAL).appendRow([new Date(), id, action, champ || '', String(avant ?? ''), String(apres ?? ''), par || '']);
}

/* ============================ API ============================ */
// Tout passe en POST (Content-Type text/plain) : le jeton ne se retrouve jamais dans une URL.

function doGet() {
  return json_({ ok: true, service: 'Suivis API', note: 'Utilise POST avec une session valide.' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse((e.postData && e.postData.contents) || '{}'); }
  catch (err) { return json_({ ok: false, error: 'Requête invalide' }); }

  // Actions publiques
  if (body.action === 'login') return json_(login_(body.nom, body.motdepasse));

  // Tout le reste exige une session
  const sess = getSession_(body.token);
  if (!sess) return json_({ ok: false, code: 'AUTH', error: 'Session expirée, reconnecte-toi.' });
  const by = sess.nom;

  if (body.action === 'logout') { PropertiesService.getScriptProperties().deleteProperty('sess_' + body.token); return json_({ ok: true }); }
  if (body.action === 'me') return json_({ ok: true, nom: sess.nom, role: sess.role });
  if (body.action === 'changePassword') return json_(changePassword_(sess.nom, body.ancien, body.nouveau, body.token));
  if (body.action === 'list') {
    const items = getAll_().filter(o => body.archives || o.statut !== 'Archivé');
    return json_({ ok: true, me: { nom: sess.nom, role: sess.role }, items: items, listes: readListes_(),
      config: { RELANCE_JOURS: readConfig_().RELANCE_JOURS }, now: new Date().toISOString() });
  }
  if (body.action === 'journal') {
    const sh = ss_().getSheetByName(TABS.JOURNAL);
    const n = Math.min(300, Math.max(sh.getLastRow() - 1, 0));
    const rows = n ? sh.getRange(sh.getLastRow() - n + 1, 1, n, 7).getValues().reverse() : [];
    return json_({ ok: true, journal: rows.filter(r => !body.id || r[1] === body.id)
      .map(r => ({ quand: r[0], id: r[1], action: r[2], champ: r[3], avant: r[4], apres: r[5], par: r[6] })) });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    switch (body.action) {
      case 'create': return json_({ ok: true, item: create_(body.fields || {}, by) });
      case 'update': return json_({ ok: true, item: update_(body.id, body.fields || {}, by) });
      case 'complete': return json_({ ok: true, item: update_(body.id, { fait: true, statut: 'Fait' }, by) });
      case 'reopen': return json_({ ok: true, item: update_(body.id, { fait: false, statut: body.statut || 'À faire' }, by) });
      case 'archive': return json_({ ok: true, item: update_(body.id, { statut: 'Archivé' }, by) });
      default: return json_({ ok: false, error: 'Action inconnue' });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  } finally { lock.releaseLock(); }
}

/* ============================ AUTHENTIFICATION ============================ */

function findUser_(nom) {
  const sh = ss_().getSheetByName(TABS.USERS);
  if (!sh || sh.getLastRow() < 2) return null;
  const key = String(nom || '').trim().toLowerCase();
  if (!key) return null;
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, USER_COLS.length).getValues();
  for (let i = 0; i < v.length; i++) {
    const n = String(v[i][0]).trim().toLowerCase(), m = String(v[i][1]).trim().toLowerCase();
    if (key === n || (m && key === m)) {
      const u = { _row: i + 2 }; USER_COLS.forEach((c, j) => u[c] = v[i][j]); return u;
    }
  }
  return null;
}

function hash_(pw, sel) {
  let h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sel + ':' + pw, Utilities.Charset.UTF_8);
  for (let i = 0; i < HASH_ROUNDS; i++) h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h.concat(Utilities.newBlob(sel).getBytes()));
  return Utilities.base64Encode(h);
}

function randomToken_() {
  return (Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
}

function safeEqual_(a, b) {
  a = String(a); b = String(b);
  let diff = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}

function pwPolicy_(pw) {
  pw = String(pw || '');
  if (pw.length < 10) return 'Le mot de passe doit avoir au moins 10 caractères.';
  if (!/[A-Za-zÀ-ÿ]/.test(pw) || !/\d/.test(pw)) return 'Le mot de passe doit contenir des lettres et au moins un chiffre.';
  return '';
}

function setPassword_(nom, pw) {
  const u = findUser_(nom);
  if (!u) throw new Error('Utilisateur introuvable');
  const sel = randomToken_().slice(0, 32);
  const sh = ss_().getSheetByName(TABS.USERS);
  sh.getRange(u._row, USER_COLS.indexOf('sel') + 1, 1, 2).setValues([[sel, hash_(pw, sel)]]);
}

function login_(nom, pw) {
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + String(nom || '').trim().toLowerCase();
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILS) return { ok: false, error: 'Trop d\'essais. Réessaie dans ' + LOCK_MINUTES + ' minutes.' };

  const u = findUser_(nom);
  const ok = u && u.actif === true && u.sel && u.hash && safeEqual_(hash_(String(pw || ''), u.sel), u.hash);
  if (!ok) {
    cache.put(failKey, String(fails + 1), LOCK_MINUTES * 60);
    Utilities.sleep(600 + Math.floor(Math.random() * 400)); // ralentit les essais en rafale
    return { ok: false, error: 'Nom ou mot de passe invalide.' };
  }
  cache.remove(failKey);
  cleanupSessions_();
  const token = randomToken_();
  const exp = Date.now() + SESSION_DAYS * 864e5;
  PropertiesService.getScriptProperties().setProperty('sess_' + token, JSON.stringify({ nom: u.nom, role: u.role, exp: exp }));
  ss_().getSheetByName(TABS.USERS).getRange(u._row, USER_COLS.indexOf('derniere_connexion') + 1).setValue(new Date());
  log_('', 'connexion', '', '', '', u.nom);
  return { ok: true, token: token, nom: u.nom, role: u.role, exp: exp };
}

function getSession_(token) {
  if (!token || !/^[a-f0-9]{96}$/.test(String(token))) return null;
  const raw = PropertiesService.getScriptProperties().getProperty('sess_' + token);
  if (!raw) return null;
  const s = JSON.parse(raw);
  if (s.exp < Date.now()) { PropertiesService.getScriptProperties().deleteProperty('sess_' + token); return null; }
  const u = findUser_(s.nom);
  if (!u || u.actif !== true) return null; // compte désactivé = accès coupé immédiatement
  return s;
}

function changePassword_(nom, ancien, nouveau, keepToken) {
  const u = findUser_(nom);
  if (!u || !safeEqual_(hash_(String(ancien || ''), u.sel), u.hash)) return { ok: false, error: 'Ancien mot de passe invalide.' };
  const err = pwPolicy_(nouveau);
  if (err) return { ok: false, error: err };
  setPassword_(nom, nouveau);
  revokeSessions_(nom, keepToken); // ferme ses autres appareils
  log_('', 'mot de passe changé', '', '', '', nom);
  return { ok: true };
}

function revokeSessions_(nom, exceptToken) {
  const props = PropertiesService.getScriptProperties();
  let n = 0;
  Object.entries(props.getProperties()).forEach(([k, v]) => {
    if (k.indexOf('sess_') !== 0 || k === 'sess_' + exceptToken) return;
    try { if (String(JSON.parse(v).nom).toLowerCase() === String(nom).toLowerCase()) { props.deleteProperty(k); n++; } } catch (e) {}
  });
  return n;
}

function cleanupSessions_() {
  const props = PropertiesService.getScriptProperties();
  const now = Date.now();
  Object.entries(props.getProperties()).forEach(([k, v]) => {
    if (k.indexOf('sess_') !== 0) return;
    try { if (JSON.parse(v).exp < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
  });
}

// Empêche l'injection de formules (=, +, -, @) dans le Sheet
function clean_(v) {
  if (typeof v !== 'string') return v;
  v = v.slice(0, 2000);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}

function create_(f, by) {
  const sh = ss_().getSheetByName(TABS.SUIVIS);
  const n = Number(readConfig_().NEXT_ID) || 1;
  setConfig_('NEXT_ID', n + 1);
  const now = new Date();
  const o = {
    id: 'SUI-' + ('000' + n).slice(-4), fait: false, priorite: 'Moyenne', type: 'Tâche', statut: 'À faire',
    ouvert_le: now, cree_par: by, cree_le: now, maj_le: now, maj_par: by, ordre: n * 10
  };
  EDITABLE.forEach(k => { if (f[k] !== undefined && f[k] !== '') o[k] = DATE_COLS.indexOf(k) >= 0 ? toDate_(f[k]) : clean_(f[k]); });
  sh.appendRow(COLS.map(c => o[c] === undefined ? '' : o[c]));
  const row = sh.getLastRow();
  sh.getRange(row, COLS.indexOf('fait') + 1).insertCheckboxes().setValue(!!o.fait);
  log_(o.id, 'création', '', '', o.action, by);
  return rowToObj_(sh.getRange(row, 1, 1, COLS.length).getValues()[0], row);
}

function update_(id, f, by) {
  const sh = ss_().getSheetByName(TABS.SUIVIS);
  const row = findRow_(id);
  if (row < 0) throw new Error('Introuvable : ' + id);
  const range = sh.getRange(row, 1, 1, COLS.length);
  const cur = range.getValues()[0];
  const next = cur.slice();
  const now = new Date();
  if (f.statut === 'Fait') f.fait = true;
  if (f.fait === true && !f.statut) f.statut = 'Fait';
  if (f.fait === false && cur[COLS.indexOf('statut')] === 'Fait' && !f.statut) f.statut = 'À faire';
  EDITABLE.forEach(k => {
    if (f[k] === undefined) return;
    const i = COLS.indexOf(k);
    const val = DATE_COLS.indexOf(k) >= 0 ? toDate_(f[k]) : clean_(f[k]);
    const before = cur[i] instanceof Date ? Utilities.formatDate(cur[i], tz_(), 'yyyy-MM-dd') : cur[i];
    const after = val instanceof Date ? Utilities.formatDate(val, tz_(), 'yyyy-MM-dd') : val;
    if (String(before) !== String(after)) { next[i] = val; log_(id, 'modif', k, before, after, by); }
  });
  const iFait = COLS.indexOf('fait');
  if (next[iFait] === true && !cur[COLS.indexOf('ferme_le')]) next[COLS.indexOf('ferme_le')] = now;
  if (next[iFait] !== true) next[COLS.indexOf('ferme_le')] = '';
  next[COLS.indexOf('maj_le')] = now;
  next[COLS.indexOf('maj_par')] = by;
  range.setValues([next]);
  return rowToObj_(next, row);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ============ Synchro quand l'équipe coche directement dans le Sheet ============ */

function onEdit(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== TABS.SUIVIS || e.range.getRow() < 2) return;
  const col = e.range.getColumn();
  const row = e.range.getRow();
  const cFait = COLS.indexOf('fait') + 1, cStat = COLS.indexOf('statut') + 1;
  const now = new Date();
  const who = (e.user && e.user.getEmail && e.user.getEmail()) || 'Sheet';
  if (col === cFait) {
    const checked = e.value === 'TRUE';
    sh.getRange(row, cStat).setValue(checked ? 'Fait' : 'À faire');
    sh.getRange(row, COLS.indexOf('ferme_le') + 1).setValue(checked ? now : '');
  } else if (col === cStat) {
    sh.getRange(row, cFait).setValue(e.value === 'Fait');
    if (e.value === 'Fait') sh.getRange(row, COLS.indexOf('ferme_le') + 1).setValue(now);
  }
  if (col !== COLS.indexOf('maj_le') + 1) {
    sh.getRange(row, COLS.indexOf('maj_le') + 1).setValue(now);
    sh.getRange(row, COLS.indexOf('maj_par') + 1).setValue(who);
  }
}
