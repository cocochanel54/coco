/**
 * Journée agent : appli de saisie des techniciens, branchée sur ce Google Sheet.
 *
 * Onglets utilisés (créés par installer()) :
 *   Saisies      : une ligne par intervention, colonnes A à L identiques à « Justificatifs »
 *   BPU          : les prix (modifiables)
 *   Techniciens  : la liste des noms proposés dans l'appli
 *   Projets      : les projets proposés, et s'ils sont des absences
 *   Jours fériés : affichés comme alerte dans l'appli
 */

var ONGLET_SAISIES = 'Saisies';
var ENTETES = ['Date', 'Technicien', 'Projet', 'Prestation (si le projet en a plusieurs)', 'Quantité (vide = 1)',
  'Désignation (auto)', 'Code BPU (auto)', 'Unité (auto)', 'PU (€) (auto)', 'Montant (€) (auto)',
  'N° ticket / référence', 'Commentaire', 'ID (ne pas modifier)', 'Saisi le'];

// Données de départ, reprises du classeur Journee_agent.xlsm
var DEFAUT = /*DATA*/null;

/** Sert la page aux téléphones, ou, avec ?export=tsv, les saisies pour la macro de synchronisation Excel. */
function doGet(e) {
  if (e && e.parameter && e.parameter.export) return exportTsv_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Journée agent')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** À lancer une fois depuis l'éditeur : crée et remplit les onglets. Ne touche pas aux onglets déjà présents. */
function installer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var saisies = ss.getSheetByName(ONGLET_SAISIES);
  if (!saisies) {
    saisies = ss.insertSheet(ONGLET_SAISIES);
    saisies.getRange(1, 1, 1, ENTETES.length).setValues([ENTETES]).setFontWeight('bold');
    saisies.setFrozenRows(1);
    saisies.getRange('A:A').setNumberFormat('dd/mm/yyyy');
    saisies.getRange('I:J').setNumberFormat('#,##0.00 €');
    saisies.getRange('N:N').setNumberFormat('dd/mm/yyyy hh:mm');
  }
  remplir_(ss, 'BPU', ['Projet', 'Code BPU', 'Libellé (choix dans l\'appli)', 'Désignation', 'Unité', 'Prix unitaire (€) (vide = sur devis)'],
    Object.keys(DEFAUT.bpu).reduce(function (rows, p) {
      return rows.concat(DEFAUT.bpu[p].map(function (x) { return [p, x[0], x[1], x[2], x[3], x[4] == null ? '' : x[4]]; }));
    }, []));
  remplir_(ss, 'Techniciens', ['Technicien'], DEFAUT.techs.map(function (t) { return [t]; }));
  remplir_(ss, 'Projets', ['Projet', 'Absence (oui / vide)'], DEFAUT.projets.map(function (p) { return [p.nom, p.absence ? 'oui' : '']; }));
  remplir_(ss, 'Jours fériés', ['Date', 'Jour férié'], Object.keys(DEFAUT.feries).map(function (d) { return [jour_(d), DEFAUT.feries[d]]; }));
  ss.setActiveSheet(saisies);
}

function remplir_(ss, nom, entetes, lignes) {
  if (ss.getSheetByName(nom)) return;
  var sh = ss.insertSheet(nom);
  sh.getRange(1, 1, 1, entetes.length).setValues([entetes]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (lignes.length) sh.getRange(2, 1, lignes.length, entetes.length).setValues(lignes);
  sh.autoResizeColumns(1, entetes.length);
}

/** Listes de l'appli, lues dans les onglets (le responsable les modifie directement dans le Sheet). */
function getConfig() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lire = function (nom) {
    var sh = ss.getSheetByName(nom);
    return sh && sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues() : [];
  };
  var techs = lire('Techniciens').map(function (r) { return String(r[0]).trim(); }).filter(String);
  var projets = lire('Projets').filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) { return { nom: String(r[0]).trim(), absence: /^o/i.test(String(r[1])) }; });
  var bpu = {};
  projets.forEach(function (p) { bpu[p.nom] = []; });
  lire('BPU').forEach(function (r) {
    var p = String(r[0]).trim(), code = String(r[1]).trim();
    if (!p || !code || !bpu[p]) return;
    var prix = r[5] === '' || r[5] === null ? null : Number(r[5]);
    bpu[p].push([code, String(r[2] || r[3]), String(r[3]), String(r[4]), isNaN(prix) ? null : prix]);
  });
  var feries = {};
  lire('Jours fériés').forEach(function (r) { if (r[0] instanceof Date) feries[iso_(r[0])] = String(r[1]); });
  if (!techs.length) return DEFAUT;
  return { techs: techs, projets: projets, bpu: bpu, feries: feries };
}

/** Lignes déjà saisies, pour « Ma journée » et la vue équipe. */
function getLignes() {
  var sh = feuille_();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ENTETES.length).getValues()
    .filter(function (r) { return r[0] instanceof Date && r[12]; })
    .map(function (r) {
      return {
        id: String(r[12]), date: iso_(r[0]), tech: String(r[1]), projet: String(r[2]), prestation: String(r[3]),
        qte: r[4] === '' ? 1 : Number(r[4]), libelle: String(r[3] || r[5]), designation: String(r[5]), code: String(r[6]),
        unite: String(r[7]), pu: Number(r[8]) || 0, montant: Number(r[9]) || 0, ticket: String(r[10]), comment: String(r[11]),
        at: r[13] instanceof Date ? r[13].toISOString() : ''
      };
    });
}

/** Ajoute une ligne. Le même ID n'est jamais ajouté deux fois (renvoi après une coupure réseau). */
function ajouterLigne(l) {
  if (!l || !l.id || !/^\d{4}-\d{2}-\d{2}$/.test(l.date) || !l.tech || !l.projet) throw new Error('Ligne incomplète');
  var verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    var sh = feuille_();
    if (sh.getLastRow() > 1 && sh.getRange(2, 13, sh.getLastRow() - 1, 1).createTextFinder(l.id).matchEntireCell(true).findNext()) return l.id;
    var q = Number(l.qte), pu = Number(l.pu);
    sh.appendRow([jour_(l.date), texte_(l.tech), texte_(l.projet), texte_(l.prestation), q === 1 ? '' : q,
      texte_(l.designation), texte_(l.code), texte_(l.unite), pu, Math.round(q * pu * 100) / 100,
      texte_(l.ticket), texte_(l.comment), l.id, new Date()]);
    return l.id;
  } finally {
    verrou.releaseLock();
  }
}

/** Supprime une ligne (bouton « Supprimer » ou « Annuler » de l'appli). */
function supprimerLigne(id) {
  var verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    var sh = feuille_();
    if (sh.getLastRow() < 2) return;
    var cell = sh.getRange(2, 13, sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
    if (cell) sh.deleteRow(cell.getRow());
  } finally {
    verrou.releaseLock();
  }
}

// Une ligne par saisie, séparateur tabulation, dates en aaaa-mm-jj, nombres avec un point
function exportTsv_() {
  var propre = function (v) { return String(v == null ? '' : v).replace(/[\t\r\n]+/g, ' ').replace(/^'/, ''); };
  var lignes = getLignes().map(function (l) {
    return [l.id, l.date, l.tech, l.projet, l.prestation, l.qte, l.ticket, l.comment, l.pu, l.montant].map(propre).join('\t');
  });
  var texte = ['ID\tDate\tTechnicien\tProjet\tPrestation\tQuantite\tTicket\tCommentaire\tPU\tMontant'].concat(lignes).join('\n');
  return ContentService.createTextOutput(texte).setMimeType(ContentService.MimeType.TEXT);
}

function feuille_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_SAISIES);
  if (!sh) { installer(); sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_SAISIES); }
  return sh;
}
function jour_(s) { var p = s.split('-'); return new Date(+p[0], +p[1] - 1, +p[2]); }
function iso_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
// Empêche qu'un texte saisi soit pris pour une formule
function texte_(v) { v = String(v == null ? '' : v).slice(0, 500); return /^[=+\-@]/.test(v) ? "'" + v : v; }
