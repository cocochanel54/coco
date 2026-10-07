/**
 * Journée agent : appli de saisie des techniciens, branchée sur ce Google Sheet.
 *
 * Onglets utilisés (créés par installer()) :
 *   Saisies      : une ligne par intervention, colonnes A à L identiques à « Justificatifs »
 *   BPU          : les prix (modifiables)
 *   Techniciens  : la liste des noms proposés dans l'appli
 *   Projets      : les projets proposés, et s'ils sont des absences
 *   Jours fériés : affichés comme alerte dans l'appli
 *
 * Les techniciens ne voient jamais les prix : la page ne reçoit que les libellés,
 * et le prix unitaire et le montant sont calculés ici, à partir de l'onglet BPU.
 * (Toute fonction sans « _ » à la fin peut être appelée depuis la page : aucune ne renvoie de prix.)
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

/** Listes envoyées à la page, sans aucun prix. */
function getConfig() {
  var c = lireConfig_();
  Object.keys(c.bpu).forEach(function (p) { c.bpu[p] = c.bpu[p].map(function (x) { return x.slice(0, 4); }); });
  return c;
}

/** Listes lues dans les onglets (le responsable les modifie directement dans le Sheet), avec les prix. */
function lireConfig_() {
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
  if (!techs.length) return JSON.parse(JSON.stringify(DEFAUT));
  return { techs: techs, projets: projets, bpu: bpu, feries: feries };
}

/** Lignes d'un technicien (60 derniers jours) pour « Ma journée », sans prix. */
function getLignes(tech) {
  var depuis = iso_(new Date(Date.now() - 60 * 864e5));
  return lireLignes_().filter(function (l) { return l.tech === tech && l.date >= depuis; }).map(function (l) {
    delete l.pu; delete l.montant; delete l.designation; return l;
  });
}

/** Toutes les lignes de l'onglet Saisies, avec les prix. */
function lireLignes_() {
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

/**
 * Ajoute une ligne. Le prix vient de l'onglet BPU (jamais de la page).
 * Le même ID n'est jamais ajouté deux fois (renvoi après une coupure réseau).
 */
function ajouterLigne(l) {
  if (!l || !l.id || !/^\d{4}-\d{2}-\d{2}$/.test(l.date) || !l.tech || !l.projet) throw new Error('Ligne incomplète');
  var conf = lireConfig_();
  if (conf.techs.indexOf(l.tech) < 0) throw new Error('Technicien inconnu : ' + l.tech);
  var liste = conf.bpu[l.projet];
  if (!liste || !liste.length) throw new Error('Projet inconnu : ' + l.projet);
  var pr = liste.length === 1 ? liste[0] : liste.filter(function (x) { return x[0] === l.code; })[0];
  if (!pr) throw new Error('Prestation inconnue pour ' + l.projet + ' : ' + l.code);
  var q = Number(l.qte);
  if (!isFinite(q) || q < 0) q = 1;
  var pu = pr[4];                                  // null = prix sur devis, à compléter par le responsable
  var verrou = LockService.getScriptLock();
  verrou.waitLock(20000);
  try {
    var sh = feuille_();
    if (sh.getLastRow() > 1 && sh.getRange(2, 13, sh.getLastRow() - 1, 1).createTextFinder(l.id).matchEntireCell(true).findNext()) return l.id;
    sh.appendRow([jour_(l.date), texte_(l.tech), texte_(l.projet), liste.length > 1 ? texte_(pr[1]) : '', q === 1 ? '' : q,
      texte_(pr[2]), texte_(pr[0]), texte_(pr[3]), pu == null ? '' : pu, pu == null ? '' : Math.round(q * pu * 100) / 100,
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
  // PU et Montant restent vides : Excel les calcule avec son propre BPU, et ce lien ne doit pas exposer les prix
  var lignes = lireLignes_().map(function (l) {
    return [l.id, l.date, l.tech, l.projet, l.prestation, l.qte, l.ticket, l.comment, '', ''].map(propre).join('\t');
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
