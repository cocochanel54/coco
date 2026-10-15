/**
 * Journée agent : appli de saisie des techniciens, branchée sur ce Google Sheet.
 *
 * Onglets utilisés (créés par installer()) :
 *   Saisies      : une ligne par intervention, colonnes A à L identiques à « Justificatifs »
 *   BPU          : les prix (modifiables)
 *   Techniciens  : la liste des noms proposés dans l'appli
 *   Projets      : les projets proposés, et s'ils sont des absences
 *   Jours fériés : affichés comme alerte dans l'appli
 *   Règles, Production : la feuille de prod (créées par installerProduction(), voir plus bas)
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
  installerProduction();
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
    majProductionSansErreur_();
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
    if (cell) {
      sh.deleteRow(cell.getRow());
      majProductionSansErreur_();
    }
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

/* ---------------------------------------------------------------------------------------------
 * Feuille de prod : la production des techniciens comptée en journées, pas en euros.
 * Chaque prestation vaut une part de journée selon le nombre qu'un technicien en fait dans une
 * journée normale (onglet Règles) : un technicien n'est plus jugé sur le marché où il est envoyé.
 * Tout se recalcule à chaque saisie de l'appli et à chaque modification du Sheet.
 *   Règles     : combien dans une journée normale, par catégorie de prestation
 *   BPU        : colonne G = catégorie de chaque prestation
 *   Saisies    : colonnes O et P = journées produites et ligne à vérifier (calculées)
 *   Production : par technicien et par marché, entre les dates choisies en B2 et B3
 * Chantier sur plusieurs jours : les jours « Chantier en cours » et la prestation finale portent le
 * même n° de référence ; la valeur de la prestation finale est partagée entre tous ces jours.
 * ------------------------------------------------------------------------------------------- */
var ONGLET_REGLES = 'Règles', ONGLET_PROD = 'Production';
var COL_CATEGORIE = 7;                 // colonne G de l'onglet BPU
var COL_CALCUL = ENTETES.length + 1;   // colonne O de l'onglet Saisies
var JAUNE = '#FFF2CC', BLEU_FONCE = '#1F3864';

/** À lancer une fois : crée les onglets Règles et Production et la colonne Catégorie du BPU. Ne remplace rien de ce qui existe. */
function installerProduction() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  feuille_();
  var conf = lireConfig_();
  var regles = ss.getSheetByName(ONGLET_REGLES);
  if (!regles) {
    var ligneDe = {};
    DEFAUT.prod.regles.forEach(function (r, i) { ligneDe[r[0]] = i + 2; });
    regles = ss.insertSheet(ONGLET_REGLES);
    regles.getRange(1, 1, 1, 6).setValues([['Catégorie', 'Ce qu\'on compte', 'Comment ça compte',
      'Combien dans une journée normale', 'Valeur d\'une journée (€)', 'Remarque']]).setFontWeight('bold');
    regles.getRange(2, 1, DEFAUT.prod.regles.length, 6).setValues(DEFAUT.prod.regles.map(function (r) {
      return [r[0], r[1], r[2], r[3] == null ? '' : r[3], '', r[5]];
    }));
    DEFAUT.prod.regles.forEach(function (r, i) {
      // « au prix » : une journée vaut une journée de tickets du même marché
      if (r[4]) regles.getRange(i + 2, 5).setFormula('=IFERROR(D' + ligneDe[DEFAUT.prod.categories[r[4]]]
        + '*VLOOKUP("' + r[4] + '",BPU!$B:$F,5,FALSE),"")');
      if (r[2] === 'au nombre') regles.getRange(i + 2, 4).setBackground(JAUNE);
    });
    regles.setFrozenRows(1);
    regles.getRange('E:E').setNumberFormat('#,##0 "€"');
    regles.autoResizeColumns(1, 5);
    regles.setColumnWidth(6, 600);
    regles.getRange('F:F').setWrap(true);
  }

  // Lignes « Chantier en cours » du BPU de départ qui manquent dans l'onglet BPU
  var bpu = ss.getSheetByName('BPU');
  var codes = bpu.getLastRow() > 1 ? bpu.getRange(2, 2, bpu.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); }) : [];
  Object.keys(DEFAUT.bpu).forEach(function (p) {
    if (!conf.bpu[p]) return;
    DEFAUT.bpu[p].forEach(function (x) {
      if (/-CHANTIER$/.test(x[0]) && codes.indexOf(x[0]) < 0) bpu.appendRow([p, x[0], x[1], x[2], x[3], x[4] == null ? '' : x[4]]);
    });
  });
  // Colonne G : la catégorie de chaque prestation (les cases déjà remplies ne changent pas)
  bpu.getRange(1, COL_CATEGORIE).setValue('Catégorie (feuille de prod)').setFontWeight('bold');
  var nb = bpu.getLastRow() - 1;
  if (nb > 0) {
    var connues = regles.getRange(2, 1, Math.max(regles.getLastRow() - 1, 1), 1).getValues().map(function (r) { return String(r[0]).trim(); });
    var absences = conf.projets.filter(function (p) { return p.absence; }).map(function (p) { return p.nom; });
    var plage = bpu.getRange(2, 1, nb, COL_CATEGORIE);
    bpu.getRange(2, COL_CATEGORIE, nb, 1).setValues(plage.getValues().map(function (r) {
      return [String(r[6]).trim() || categorieParDefaut_(String(r[0]).trim(), String(r[1]).trim(), absences, connues)];
    })).setBackground(JAUNE).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInRange(regles.getRange('A2:A'), true).setAllowInvalid(true).build());
    bpu.autoResizeColumn(COL_CATEGORIE);
  }

  if (!ss.getSheetByName(ONGLET_PROD)) {
    var prod = ss.insertSheet(ONGLET_PROD, 0);
    var auj = new Date();
    prod.getRange('A1').setValue('PRODUCTION DES TECHNICIENS, EN JOURNÉES (1 = une journée normale, quel que soit le marché)')
      .setFontWeight('bold').setFontSize(13).setFontColor(BLEU_FONCE);
    prod.getRange('A2:B3').setValues([['Du', new Date(auj.getFullYear(), auj.getMonth(), 1)],
      ['Au', new Date(auj.getFullYear(), auj.getMonth() + 1, 0)]]);
    prod.getRange('A2:A3').setFontWeight('bold');
    prod.getRange('B2:B3').setNumberFormat('dd/mm/yyyy').setBackground(JAUNE);
    prod.getRange('C3').setValue('Change les dates en B2 et B3. Tout le reste se met à jour seul, à chaque saisie des techniciens.')
      .setFontStyle('italic').setFontColor('#595959');
    prod.setColumnWidth(1, 170);
    for (var c = 2; c <= 8; c++) prod.setColumnWidth(c, 115);
  }
  majProduction_();
}

/** Catégorie proposée pour une prestation ajoutée au BPU. Vide = au responsable de la choisir. */
function categorieParDefaut_(projet, code, absences, connues) {
  var cat = DEFAUT.prod.categories[code]
    || (/CHANTIER$/i.test(code) ? 'Chantier en cours' : '')
    || (absences.indexOf(projet) >= 0 ? 'Absence' : '')
    || (/-TICKET$/i.test(code) ? projet + ' · Ticket' : '');
  return connues.indexOf(cat) >= 0 ? cat : '';
}

/** Menu « Feuille de prod » du Google Sheet. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Feuille de prod')
    .addItem('Mettre à jour la production', 'majProduction')
    .addItem('Installer la feuille de prod', 'installerProduction')
    .addToUi();
}

/** Recalcule quand le responsable modifie le Sheet (dates, règles, BPU, saisies). */
function onEdit(e) {
  if (!e || !e.range) return;
  var nom = e.range.getSheet().getName();
  if ([ONGLET_SAISIES, 'BPU', ONGLET_REGLES, ONGLET_PROD, 'Techniciens', 'Projets'].indexOf(nom) < 0) return;
  if (nom === ONGLET_PROD && e.range.getRow() > 3) return;
  if (nom === ONGLET_SAISIES && e.range.getColumn() >= COL_CALCUL) return;
  try { majProduction(); } catch (err) { majProduction_(); }
}

function majProduction() {
  var verrou = LockService.getScriptLock();
  if (!verrou.tryLock(20000)) return;
  try { majProduction_(); } finally { verrou.releaseLock(); }
}

// Une erreur de la feuille de prod ne doit jamais empêcher une saisie
function majProductionSansErreur_() {
  try { majProduction_(); } catch (e) { console.error('Feuille de prod : ' + (e && e.stack || e)); }
}

function majProduction_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var prod = ss.getSheetByName(ONGLET_PROD), regles = ss.getSheetByName(ONGLET_REGLES), bpu = ss.getSheetByName('BPU');
  if (!prod || !regles || !bpu) return;   // feuille de prod pas encore installée
  var lire = function (sh, nc) { return sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, nc).getValues() : []; };
  var sh = feuille_(), saisies = lire(sh, ENTETES.length), conf = lireConfig_();
  var res = calculerProduction_(saisies, lire(bpu, COL_CATEGORIE), lire(regles, 5), conf.techs,
    conf.projets.filter(function (p) { return !p.absence; }).map(function (p) { return p.nom; }),
    prod.getRange('B2').getValue(), prod.getRange('B3').getValue());
  sh.getRange(1, COL_CALCUL, 1, 2).setValues([['Journées produites (calculé)', 'À vérifier (calculé)']]).setFontWeight('bold');
  if (saisies.length) {
    sh.getRange(2, COL_CALCUL, saisies.length, 2).setValues(res.lignes.map(function (l) {
      return [l.journees === null ? '' : l.journees, l.verif];
    }));
    sh.getRange(2, COL_CALCUL, saisies.length, 1).setNumberFormat('0.00');
    sh.getRange(2, COL_CALCUL + 1, saisies.length, 1).setFontColor('#C00000');
  }
  ecrireProduction_(prod, res);
}

/**
 * Le calcul, sans Google (testable) : saisies = lignes A à N de Saisies, bpu = colonnes A à G du BPU,
 * regles = colonnes A à E de Règles. Mêmes règles que feuille-prod/Feuille_prod.xlsx.
 */
function calculerProduction_(saisies, bpu, regles, techs, marches, du, au) {
  var R = {}, B = {};
  regles.forEach(function (r) {
    var c = String(r[0]).trim();
    if (c) R[c] = { mode: String(r[2]).trim(), n: nombre_(r[3]), valeur: nombre_(r[4]) };
  });
  bpu.forEach(function (r) {
    var code = String(r[1]).trim();
    if (code && !B[code]) B[code] = { prix: nombre_(r[5]), cat: String(r[6] == null ? '' : r[6]).trim() };
  });
  var debut = du instanceof Date ? jourNum_(du) : -Infinity, fin = au instanceof Date ? jourNum_(au) : Infinity;

  var lignes = saisies.map(function (r, i) {
    var l = { ligne: i + 2, date: r[0], tech: String(r[1]).trim(), projet: String(r[2]).trim(), code: String(r[6]).trim(),
      ref: String(r[10]).trim().toLowerCase(), jour: r[0] instanceof Date ? jourNum_(r[0]) : null,
      mode: '', base: null, journees: null, montant: null, travail: false, verif: '' };
    l.vide = l.jour === null && !String(r[0]).trim() && !l.tech && !l.code;
    if (l.vide) return l;
    var b = B[l.code];
    l.cat = b && b.cat ? b.cat : (/CHANTIER$/i.test(l.code) ? 'Chantier en cours' : '');
    var regle = R[l.cat];
    l.mode = regle ? regle.mode : '';
    var qte = r[4] === '' || r[4] === null ? 1 : Number(r[4]);
    var pu = typeof r[8] === 'number' ? r[8] : b ? b.prix : null;
    if (pu !== null && isFinite(qte)) l.montant = pu * qte;
    if (!isFinite(qte)) l.base = null;
    else if (l.mode === 'au nombre') l.base = regle.n > 0 ? qte / regle.n : null;
    else if (l.mode === 'au prix') l.base = l.montant !== null && regle.valeur > 0 ? l.montant / regle.valeur : null;
    else if (['chantier', 'ne compte pas', 'absence'].indexOf(l.mode) >= 0) l.base = 0;
    l.journees = l.base;
    l.travail = l.jour !== null && !!l.tech && !!l.code && l.mode !== 'absence';
    l.verif = l.jour === null ? 'Date manquante ou non reconnue'
      : !l.tech ? 'Technicien manquant'
      : techs.indexOf(l.tech) < 0 ? 'Technicien absent de l\'onglet Techniciens'
      : !l.code ? 'Code BPU manquant (colonne G)'
      : !b && !l.cat ? 'Code BPU inconnu (onglet BPU)'
      : !l.cat ? 'Catégorie à choisir (onglet BPU, colonne G)'
      : !regle ? 'Catégorie inconnue (onglet Règles)'
      : !isFinite(qte) ? 'Quantité non reconnue'
      : l.mode === 'au prix' && pu === null ? 'Prix sur devis à saisir (colonne I)'
      : l.mode === 'chantier' && !l.ref ? 'Référence manquante (colonne K) : ce jour de chantier ne compte pas'
      : l.base === null ? 'Règle à compléter (onglet Règles)' : '';
    return l;
  });

  // Chantiers : même référence et même projet ; la valeur des prestations finales est partagée
  // en parts égales entre les jours de chantier et le jour de la prestation finale
  var groupes = {};
  lignes.forEach(function (l) {
    if (!l.vide && l.ref) (groupes[l.ref + '\u0001' + l.projet] = groupes[l.ref + '\u0001' + l.projet] || []).push(l);
  });
  Object.keys(groupes).forEach(function (k) {
    var g = groupes[k];
    var nC = g.filter(function (l) { return l.mode === 'chantier'; }).length;
    if (!nC) return;
    var v = g.reduce(function (s, l) { return s + (l.mode !== 'chantier' && l.base !== null ? l.base : 0); }, 0);
    g.forEach(function (l) {
      if (l.mode === 'chantier') l.journees = v / (nC + 1);
      else if (l.base !== null) l.journees = l.base / (nC + 1);
      if (l.mode === 'chantier' && !l.verif && !(l.journees > 0)) l.verif = 'Chantier en attente : prestation finale pas encore saisie avec la même référence';
    });
  });

  var dans = function (l) { return !l.vide && l.jour !== null && l.jour >= debut && l.jour <= fin; };
  var parTech = techs.map(function (t) {
    var jours = {}, j = 0, ca = 0, v = 0;
    lignes.forEach(function (l) {
      if (l.tech !== t || !dans(l)) return;
      if (l.travail) jours[l.jour] = true;
      if (l.journees !== null) j += l.journees;
      if (l.montant !== null) ca += l.montant;
      if (l.verif) v++;
    });
    var nj = Object.keys(jours).length;
    return { tech: t, jours: nj, journees: j, taux: nj ? j / nj : null, ca: ca, verif: v };
  });
  rang_(parTech, 'taux', 'rang', false);
  rang_(parTech, 'ca', 'rangCa', true);
  var total = parTech.reduce(function (s, t) {
    return { jours: s.jours + t.jours, journees: s.journees + t.journees, ca: s.ca + t.ca, verif: s.verif + t.verif };
  }, { jours: 0, journees: 0, ca: 0, verif: 0 });
  total.taux = total.jours ? total.journees / total.jours : null;
  var parMarche = marches.map(function (m) {
    var j = 0, ca = 0;
    lignes.forEach(function (l) {
      if (l.projet !== m || !dans(l)) return;
      if (l.journees !== null) j += l.journees;
      if (l.montant !== null) ca += l.montant;
    });
    return { marche: m, journees: j, ca: ca, caJour: j > 0 ? ca / j : null };
  });
  return { lignes: lignes, techs: parTech, total: total, marches: parMarche,
    verifs: lignes.filter(function (l) { return l.verif; }) };
}

// Classement : 1 = le plus haut ; ex aequo au même rang
function rang_(lignesTab, cle, dest, positifSeulement) {
  lignesTab.forEach(function (r) {
    var v = r[cle];
    r[dest] = v === null || (positifSeulement && !(v > 0)) ? null
      : 1 + lignesTab.filter(function (o) { return o[cle] !== null && o[cle] > v; }).length;
  });
}

function ecrireProduction_(sh, res) {
  var n = res.verifs.length;
  sh.getRange('C2').setValue(n ? n + ' ligne(s) à vérifier : liste en bas, et colonne P de l\'onglet Saisies' : 'Aucune ligne à vérifier')
    .setFontWeight('bold').setFontColor(n ? '#C00000' : '#595959');
  if (sh.getMaxRows() >= 5) sh.getRange(5, 1, sh.getMaxRows() - 4, 8).clear();
  var vide = function (v) { return v === null ? '' : v; };
  var t = res.total;
  var r = tableau_(sh, 5, 'PAR TECHNICIEN',
    ['Technicien', 'Jours travaillés', 'Journées produites', 'Taux de production', 'Classement', 'CA (€)', 'Classement au CA (ancienne méthode)', 'Lignes à vérifier'],
    res.techs.map(function (x) { return [x.tech, x.jours, x.journees, vide(x.taux), vide(x.rang), x.ca, vide(x.rangCa), x.verif]; })
      .concat([['Total équipe', t.jours, t.journees, vide(t.taux), '', t.ca, '', t.verif]]),
    ['@', '0', '0.00', '0%', '0', '#,##0 "€"', '0', '0'],
    ['Taux 100 % = une journée normale par jour travaillé, quel que soit le marché. Les congés et absences ne comptent pas comme jours travaillés.',
     '« Classement au CA » = l\'ancienne façon de compter, en euros, laissée pour comparer.']);
  sh.getRange(5 + 1 + res.techs.length + 1, 1, 1, 8).setFontWeight('bold').setBackground('#F2F2F2');
  r = tableau_(sh, r, 'PAR MARCHÉ, SUR LES MÊMES DATES', ['Marché', 'Journées produites', 'CA (€)', 'CA pour une journée de travail (€)'],
    res.marches.map(function (m) { return [m.marche, m.journees, m.ca, vide(m.caJour)]; }),
    ['@', '0.00', '#,##0 "€"', '#,##0 "€"'],
    ['La dernière colonne montre ce que rapporte une journée de travail sur chaque marché.']);
  var max = 200;
  tableau_(sh, r, 'LIGNES À VÉRIFIER (toutes dates)', ['Ligne de Saisies', 'Date', 'Technicien', 'Projet', 'Code BPU', 'Problème'],
    res.verifs.slice(0, max).map(function (l) { return [l.ligne, l.date, l.tech, l.projet, l.code, l.verif]; }),
    ['0', 'dd/mm/yyyy', '@', '@', '@', '@'],
    n > max ? ['… et ' + (n - max) + ' autres : filtre la colonne P de l\'onglet Saisies.'] : n ? [] : ['Rien à vérifier.']);
}

function tableau_(sh, r, titre, entetes, lignes, formats, notes) {
  sh.getRange(r, 1).setValue(titre).setFontWeight('bold').setFontColor(BLEU_FONCE);
  sh.getRange(r + 1, 1, 1, entetes.length).setValues([entetes]).setFontWeight('bold').setFontColor('#FFFFFF')
    .setBackground(BLEU_FONCE).setWrap(true).setVerticalAlignment('middle');
  if (lignes.length) {
    var plage = sh.getRange(r + 2, 1, lignes.length, entetes.length);
    plage.setNumberFormats(lignes.map(function () { return formats; }));
    plage.setValues(lignes);
  }
  r += 2 + lignes.length;
  (notes || []).forEach(function (txt) { sh.getRange(r++, 1).setValue(txt).setFontStyle('italic').setFontColor('#595959'); });
  return r + 1;
}

function nombre_(v) { return typeof v === 'number' && isFinite(v) ? v : null; }
function jourNum_(d) { return Math.round(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()) / 864e5); }
