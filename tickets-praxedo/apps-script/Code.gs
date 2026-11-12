/**
 * Carte des tickets Praxedo, branchée sur ce Google Sheet.
 *
 * Onglets (créés par installer()) :
 *   Tickets             : le dernier export Praxedo importé (menu Carte tickets → Importer un export Praxedo)
 *   Fins d'intervention : une ligne par mail « [FIN INTERVENTION] » reçu, lu dans Gmail toutes les 5 minutes
 *                         (une règle Outlook transfère ces mails vers ce Gmail)
 *   Créations           : une ligne par mail « [CREATION INTERVENTION] » : le ticket apparaît sur la carte
 *                         (en violet, avec sa GTR) jusqu'à ce qu'un export le contienne
 *   Sortis de l'export  : les tickets présents dans l'import précédent et absents du nouveau (terminés ou annulés)
 *   Adresses            : les coordonnées des adresses (service d'adresses de l'IGN), pour ne pas les chercher deux fois
 *
 * Import automatique : un mail reçu dans Gmail avec l'export Praxedo en pièce jointe (.csv) est importé tout seul
 * dans les 5 minutes (export programmé par Praxedo, ou export que tu t'envoies par mail).
 *
 * Les fins d'intervention arrivent de trois façons : un statut de fin dans l'export (Terminée, Réalisée, Validée…),
 * un ticket qui sort du filtre entre deux imports, ou un mail « [FIN INTERVENTION] » transféré à la main vers Gmail.
 *
 * La carte (Déployer → Application Web) recharge ces données toutes les 2 minutes et relit Gmail au passage.
 * Toute fonction sans « _ » à la fin peut être appelée depuis la page.
 */

var ONGLET_TICKETS = 'Tickets';
var ONGLET_FINS = "Fins d'intervention";
var ONGLET_ADRESSES = 'Adresses';
var ONGLET_SORTIS = "Sortis de l'export";
var ONGLET_CREATIONS = 'Créations';
var ENTETES_CREATIONS = ['Reçu le', 'N° ticket', 'Type', 'Commune (objet)', 'Adresse', 'CP', 'Ville', 'Nom client', 'Description', 'Id du mail'];
var RECHERCHE_CREATIONS = 'subject:"CREATION INTERVENTION" newer_than:30d';
var ENTETES_SORTIS = ['Sorti le', 'N°', 'Statut avant', 'Planifiée', 'Technicien', 'Adresse', 'CP', 'Ville site', 'Site'];
var ENTETES_FINS = ['Reçu le', 'N° ticket', 'Commune', 'Intervention du', 'Rapport', 'Objet du mail', 'Id du mail'];
var ENTETES_ADRESSES = ['Adresse cherchée', 'Latitude', 'Longitude', 'Adresse trouvée', 'Score', 'Type', 'Commune', 'CP'];
// Mails de fin d'intervention : l'objet contient « [FIN INTERVENTION] … N°BB-20261102T0761#0 - COMMUNE »
var RECHERCHE_GMAIL = 'subject:"FIN INTERVENTION" newer_than:30d';
var RE_NUM = /N°\s*([A-Z]{2,4}-\d{8}T\d{4}(?:#\d+)?|\d{5,})/i;
// Exports Praxedo reçus par mail : toute pièce jointe .csv avec une colonne « N° » (les autres CSV sont ignorés)
var RECHERCHE_EXPORTS = 'has:attachment filename:csv newer_than:7d';

/** La carte. */
function doGet() {
  return HtmlService.createTemplateFromFile('Carte').evaluate()
    .setTitle('Carte des tickets')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function inclure_(nom) {
  return HtmlService.createHtmlOutputFromFile(nom).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Carte tickets')
    .addItem('Importer un export Praxedo…', 'ouvrirImport')
    .addItem('Lire les mails maintenant', 'lireMails')
    .addItem('Clé pour la macro Outlook', 'afficherCle')
    .addToUi();
}

/** À lancer une fois depuis l'éditeur : crée les onglets et la lecture des mails toutes les 5 minutes. */
function installer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [[ONGLET_TICKETS, ['Importe un export Praxedo : menu Carte tickets → Importer un export Praxedo']],
   [ONGLET_FINS, ENTETES_FINS], [ONGLET_CREATIONS, ENTETES_CREATIONS], [ONGLET_SORTIS, ENTETES_SORTIS], [ONGLET_ADRESSES, ENTETES_ADRESSES]].forEach(function (o) {
    if (ss.getSheetByName(o[0])) return;
    var f = ss.insertSheet(o[0]);
    f.getRange(1, 1, 1, o[1].length).setValues([o[1]]).setFontWeight('bold');
    f.setFrozenRows(1);
    f.getRange('A:Z').setNumberFormat('@');
  });
  var feuille1 = ss.getSheetByName('Feuille 1') || ss.getSheetByName('Sheet1');
  if (feuille1 && ss.getSheets().length > 1 && feuille1.getLastRow() === 0) ss.deleteSheet(feuille1);
  if (!ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'lireMails'; })) {
    ScriptApp.newTrigger('lireMails').timeBased().everyMinutes(5).create();
  }
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('CLE_MACRO')) props.setProperty('CLE_MACRO', Utilities.getUuid());
  lireMails();
}

function ouvrirImport() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('Import').setWidth(420).setHeight(220), 'Importer un export Praxedo');
}

function afficherCle() {
  SpreadsheetApp.getUi().alert('Clé à coller dans la macro Outlook :\n\n' + PropertiesService.getScriptProperties().getProperty('CLE_MACRO'));
}

/** Remplace l'onglet Tickets par l'export (texte du CSV) et cherche les adresses nouvelles. */
function importerExport(texte) {
  var lignes = lireCsv_(texte);
  if (lignes.length < 2 || lignes[0].indexOf('N°') < 0) throw new Error("Ce fichier n'est pas un export d'interventions Praxedo (pas de colonne « N° »).");
  var entetes = lignes[0];
  while (entetes.length && !entetes[entetes.length - 1]) entetes.pop();
  var n = entetes.length;
  var valeurs = lignes.filter(function (l, i) { return i === 0 || l.some(function (v) { return v; }); })
    .map(function (l) { var r = l.slice(0, n); while (r.length < n) r.push(''); return r.map(texte_); });
  var ss = SpreadsheetApp.getActiveSpreadsheet(), f = ss.getSheetByName(ONGLET_TICKETS);
  noterSortis_(ss, f.getDataRange().getDisplayValues(), valeurs);
  f.clear();
  f.getRange(1, 1, valeurs.length, n).setNumberFormat('@').setValues(valeurs);
  f.getRange(1, 1, 1, n).setFontWeight('bold');
  f.setFrozenRows(1);
  PropertiesService.getScriptProperties().setProperty('IMPORTE_LE', maintenant_());
  var c = {}; entetes.forEach(function (x, i) { c[x] = i; });
  var nouvelles = geocoder_(valeurs.slice(1).map(function (r) {
    return [r[c['Adresse']], r[c['CP']], r[c['Ville site']]].filter(function (x) { return x; }).join(' ');
  }));
  return { tickets: valeurs.length - 1, adresses: nouvelles };
}

/** Tickets de l'ancien import absents du nouveau : ils sont sortis du filtre (terminés ou annulés). Garde 14 jours. */
function noterSortis_(ss, anciens, nouveaux) {
  if (anciens.length < 2 || anciens[0].indexOf('N°') < 0) return;
  var feuille = ss.getSheetByName(ONGLET_SORTIS) || ss.insertSheet(ONGLET_SORTIS);
  if (feuille.getLastRow() === 0) feuille.getRange(1, 1, 1, ENTETES_SORTIS.length).setValues([ENTETES_SORTIS]).setFontWeight('bold');
  var a = {}; anciens[0].forEach(function (c, i) { a[c] = i; });
  var iNum = nouveaux[0].indexOf('N°'), presents = {};
  nouveaux.slice(1).forEach(function (r) { presents[r[iNum]] = 1; });
  var le = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
  var lignes = anciens.slice(1).filter(function (r) { return r[a['N°']] && !presents[r[a['N°']]]; }).map(function (r) {
    var g = function (c) { return a[c] == null ? '' : r[a[c]]; };
    return [le, g('N°'), g('Statut'), g('Planifiée'), [g('Nom technicien'), g('Prénom technicien')].join(' ').trim(),
            g('Adresse'), g('CP'), g('Ville site'), g('Site')].map(texte_);
  });
  var limite = Date.now() - 14 * 86400000, garder = feuille.getDataRange().getDisplayValues().slice(1).filter(function (r) {
    var m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(r[0]);
    return m && new Date(+m[3], m[2] - 1, +m[1]).getTime() > limite;
  });
  var tout = garder.concat(lignes);
  feuille.getRange(2, 1, Math.max(feuille.getLastRow() - 1, 1), ENTETES_SORTIS.length).clearContent();
  if (tout.length) feuille.getRange(2, 1, tout.length, ENTETES_SORTIS.length).setNumberFormat('@').setValues(tout);
}

/** Lit les mails de fin d'intervention arrivés dans Gmail (déclencheur toutes les 5 minutes). */
function lireMails() {
  var verrou = LockService.getScriptLock();
  if (!verrou.tryLock(20000)) return 0;
  try {
    var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_FINS);
    var vus = {};
    if (f.getLastRow() > 1) f.getRange(2, 7, f.getLastRow() - 1, 1).getValues().forEach(function (r) { vus[r[0]] = 1; });
    var ajout = [];
    GmailApp.search(RECHERCHE_GMAIL, 0, 100).forEach(function (fil) {
      fil.getMessages().forEach(function (m) {
        if (vus[m.getId()]) return;
        var l = ligneFin_(m.getSubject(), m.getPlainBody(), m.getDate(), m.getId());
        if (l) { ajout.push(l); vus[m.getId()] = 1; }
      });
    });
    if (ajout.length) f.getRange(f.getLastRow() + 1, 1, ajout.length, ENTETES_FINS.length).setNumberFormat('@').setValues(ajout);
    try { lireCreations_(); } catch (err) { console.warn('Créations : ' + err); }
    try { importerDepuisGmail_(); } catch (err) { console.warn('Import par mail : ' + err); }
    PropertiesService.getScriptProperties().setProperty('LU_LE', maintenant_());
    return ajout.length;
  } finally {
    verrou.releaseLock();
  }
}

/** Mails « [CREATION INTERVENTION] » : une ligne par ticket créé, et son adresse placée sur la carte. */
function lireCreations_() {
  var f = feuille_(ONGLET_CREATIONS, ENTETES_CREATIONS);
  var vus = {};
  if (f.getLastRow() > 1) f.getRange(2, 10, f.getLastRow() - 1, 1).getValues().forEach(function (r) { vus[r[0]] = 1; });
  var ajout = [];
  GmailApp.search(RECHERCHE_CREATIONS, 0, 100).forEach(function (fil) {
    fil.getMessages().forEach(function (m) {
      if (vus[m.getId()]) return;
      var l = ligneCreation_(m.getSubject(), m.getPlainBody(), m.getDate(), m.getId());
      if (l) { ajout.push(l.map(texte_)); vus[m.getId()] = 1; }
    });
  });
  if (!ajout.length) return 0;
  f.getRange(f.getLastRow() + 1, 1, ajout.length, ENTETES_CREATIONS.length).setNumberFormat('@').setValues(ajout);
  geocoder_(ajout.map(function (l) { return [l[4], l[5], l[6]].filter(function (x) { return x; }).join(' '); }));
  return ajout.length;
}

/**
 * Importe le plus récent export Praxedo reçu en pièce jointe, s'il est plus récent que le dernier import.
 * Chaque mail n'est regardé qu'une fois ; un CSV qui n'est pas un export d'interventions est ignoré.
 */
function importerDepuisGmail_() {
  var props = PropertiesService.getScriptProperties();
  var vus = JSON.parse(props.getProperty('EXPORTS_VUS') || '[]');
  var dernier = props.getProperty('IMPORTE_LE') ? new Date(props.getProperty('IMPORTE_LE')) : new Date(0);
  var candidats = [];
  GmailApp.search(RECHERCHE_EXPORTS, 0, 20).forEach(function (fil) {
    fil.getMessages().forEach(function (m) {
      if (vus.indexOf(m.getId()) >= 0) return;
      vus.push(m.getId());
      if (m.getDate() <= dernier) return;
      m.getAttachments().forEach(function (pj) {
        if (/\.csv$/i.test(pj.getName())) candidats.push({ date: m.getDate(), pj: pj });
      });
    });
  });
  props.setProperty('EXPORTS_VUS', JSON.stringify(vus.slice(-200)));
  candidats.sort(function (x, y) { return y.date - x.date; });
  for (var i = 0; i < candidats.length; i++) {
    var texte = candidats[i].pj.getDataAsString('UTF-8');
    if (!/(^|;|")N°("|;)/.test(texte.slice(0, 2000))) continue;  // autre CSV : ignoré
    var r = importerExport(texte);
    props.setProperty('IMPORTE_PAR_MAIL', candidats[i].pj.getName() + ' (mail du ' + Utilities.formatDate(candidats[i].date, 'Europe/Paris', 'dd/MM/yyyy HH:mm') + ')');
    return r;
  }
  return null;
}

/** Pour la macro Outlook, si la DSI bloque le transfert vers Gmail : reçoit l'objet et le texte d'un mail. */
function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  if (!d.cle || d.cle !== PropertiesService.getScriptProperties().getProperty('CLE_MACRO')) {
    return ContentService.createTextOutput('cle refusee');
  }
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_FINS);
  var id = 'outlook:' + d.id;
  var deja = f.getLastRow() > 1 && f.getRange(2, 7, f.getLastRow() - 1, 1).getValues().some(function (r) { return r[0] === id; });
  var l = deja ? null : ligneFin_(d.objet || '', d.corps || '', d.recu ? new Date(d.recu) : new Date(), id);
  if (l) f.appendRow(l.map(texte_));
  return ContentService.createTextOutput(l ? 'ok' : 'ignore');
}

/** Tout ce dont la carte a besoin. Relit Gmail si la dernière lecture date de plus de 2 minutes. */
function donnees() {
  var props = PropertiesService.getScriptProperties();
  var lu = props.getProperty('LU_LE');
  if (!lu || Date.now() - new Date(lu).getTime() > 2 * 60 * 1000) {
    try { lireMails(); } catch (err) { console.warn(err); }
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var t = ss.getSheetByName(ONGLET_TICKETS).getDataRange().getDisplayValues();
  var fins = ss.getSheetByName(ONGLET_FINS).getDataRange().getDisplayValues().slice(1).map(function (r) {
    return { recu: r[0], num: r[1], commune: r[2], jour: r[3], rapport: r[4] };
  });
  var fc = ss.getSheetByName(ONGLET_CREATIONS);
  var creations = fc ? fc.getDataRange().getDisplayValues().slice(1).map(function (r) {
    return { recu: r[0], num: r[1], type: r[2], commune: r[3], adresse: r[4], cp: r[5], ville: r[6], client: r[7], desc: r[8] };
  }) : [];
  var fs = ss.getSheetByName(ONGLET_SORTIS);
  var sortis = fs ? fs.getDataRange().getDisplayValues().slice(1).map(function (r) {
    return { le: r[0], num: r[1], statut: r[2], planif: r[3], tech: r[4], adresse: r[5], cp: r[6], ville: r[7], site: r[8] };
  }) : [];
  var adresses = {};
  ss.getSheetByName(ONGLET_ADRESSES).getDataRange().getValues().slice(1).forEach(function (r) {
    adresses[r[0]] = r[1] === '' ? null : { lat: +r[1], lon: +r[2], label: r[3], score: +r[4], type: r[5], city: r[6], postcode: String(r[7]) };
  });
  return { entetes: t[0], lignes: t.slice(1), fins: fins, creations: creations, sortis: sortis, adresses: adresses,
           importeLe: props.getProperty('IMPORTE_LE') || '', luLe: props.getProperty('LU_LE') || '' };
}

// ---------- outils ----------

/** [reçu le, N°, commune, intervention du, rapport, objet, id] ou null si l'objet n'a pas de n° de ticket. */
function ligneFin_(objet, corps, recu, id) {
  var m = RE_NUM.exec(objet);
  if (!m) return null;
  var apres = objet.slice(m.index + m[0].length).replace(/^\s*-\s*/, '').trim();
  var jour = /faite ce jour\s+(\d{2}\/\d{2}\/\d{4})/i.exec(corps);
  var lien = /https?:\/\/\S*praxedo\S*\/render/i.exec(corps);
  return [Utilities.formatDate(recu, 'Europe/Paris', 'dd/MM/yyyy HH:mm'), m[1].toUpperCase(), apres,
          jour ? jour[1] : '', lien ? lien[0] : '', objet, id];
}

/**
 * [reçu le, N°, type, commune de l'objet, adresse, CP, ville, nom client, description, id] d'un mail de création.
 * Le texte du mail a deux blocs : « info ticket : » (la description Praxedo, avec [04h HO], [Niveau 2]…)
 * et « info Client : » (nom client, Adresse, Ville). Quand l'adresse vaut « 00 voir coordonnées GPS »,
 * on prend celle écrite dans la description (« 3 RUE …, 54000, NANCY » ou « … / 52 AVENUE DE METZ 54320 MAXEVILLE »).
 */
function ligneCreation_(objet, corps, recu, id) {
  var m = RE_NUM.exec(objet);
  if (!m) return null;
  var commune = objet.slice(m.index + m[0].length).replace(/^\s*-\s*/, '').trim();
  var type = (/\]\s*Ticket\s+(.+?)\s+N°/i.exec(objet) || [])[1] || '';
  var texte = String(corps).replace(/\r/g, '').replace(/<[^>\n]*>/g, '');  // le texte brut ajoute les liens en <https://…>
  var bloc = /info ticket\s*:?[ \t]*\n([\s\S]*?)(?:\n\s*info client\s*:|$)/i.exec(texte);
  var desc = bloc ? bloc[1].replace(/\n{3,}/g, '\n\n').trim() : '';
  var champ = function (nom) {
    var r = new RegExp('^\\s*' + nom + '\\s*:\\s*(.*)$', 'im').exec(texte), v = r ? r[1].trim() : '';
    return /voir coordonn|^0+(\s|$)|^NA$/i.test(v) ? '' : v;
  };
  var adresse = champ('Adresse'), ville = champ('Ville'), cp = '';
  var a = /^\s*(\d+[^,\n]*?),\s*(\d{5}),?\s*([^\n,]+?)\s*$/m.exec(desc) || /\/\s*(\d+[^\n\/]*?)\s+(\d{5})\s+([^\n\/]+?)\s*$/m.exec(desc);
  if (a && (!adresse || !ville)) { adresse = a[1].trim(); cp = a[2]; ville = a[3].trim(); }
  else if (a && a[3].trim().toUpperCase() === ville.toUpperCase()) cp = a[2];  // même ville : on garde le CP de la description
  if (!ville && commune && !/^NA$/i.test(commune)) ville = commune;
  var cpm = /\b(\d{5})\b/.exec(adresse + ' ' + ville);
  if (!cp && cpm) cp = cpm[1];
  return [Utilities.formatDate(recu, 'Europe/Paris', 'dd/MM/yyyy HH:mm'), m[1].toUpperCase(), type, commune,
          adresse, cp, ville, champ('nom client'), desc, id];
}

function feuille_(nom, entetes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), f = ss.getSheetByName(nom);
  if (!f) {
    f = ss.insertSheet(nom);
    f.getRange(1, 1, 1, entetes.length).setValues([entetes]).setFontWeight('bold');
    f.setFrozenRows(1);
    f.getRange('A:Z').setNumberFormat('@');
  }
  return f;
}

/** Cherche les adresses inconnues avec le service d'adresses de l'IGN (rue, CP, ville : rien d'autre n'est envoyé). */
function geocoder_(adresses) {
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_ADRESSES);
  var connues = {};
  if (f.getLastRow() > 1) f.getRange(2, 1, f.getLastRow() - 1, 1).getValues().forEach(function (r) { connues[r[0]] = 1; });
  var aChercher = adresses.filter(function (q, i) { return q && !connues[q] && adresses.indexOf(q) === i; });
  var lignes = [];
  for (var i = 0; i < aChercher.length; i += 20) {
    var lot = aChercher.slice(i, i + 20);
    var reponses = UrlFetchApp.fetchAll(lot.map(function (q) {
      var cp = /\b(\d{5})\b/.exec(q);
      return { url: 'https://data.geopf.fr/geocodage/search?limit=1&q=' + encodeURIComponent(q) + (cp ? '&postcode=' + cp[1] : ''),
               muteHttpExceptions: true };
    }));
    reponses.forEach(function (rep, j) {
      var feat = null;
      try { feat = (JSON.parse(rep.getContentText()).features || [])[0]; } catch (err) { /* réponse illisible : adresse non trouvée */ }
      if (!feat && rep.getResponseCode() !== 200) return;  // service indisponible : on réessaiera au prochain import
      lignes.push(feat ? [lot[j], feat.geometry.coordinates[1], feat.geometry.coordinates[0], feat.properties.label,
                          feat.properties.score, feat.properties.type, feat.properties.city, feat.properties.postcode]
                       : [lot[j], '', '', '', '', '', '', '']);
    });
  }
  if (lignes.length) f.getRange(f.getLastRow() + 1, 1, lignes.length, ENTETES_ADRESSES.length).setValues(lignes.map(function (l) { return l.map(texte_); }));
  return lignes.length;
}

/** CSV de Praxedo : séparateur « ; », guillemets, retours à la ligne dans les cellules. */
function lireCsv_(texte) {
  texte = String(texte).replace(/^\uFEFF/, '');
  var lignes = [], ligne = [], champ = '', guil = false;
  for (var i = 0; i < texte.length; i++) {
    var ch = texte[i];
    if (guil) {
      if (ch === '"' && texte[i + 1] === '"') { champ += '"'; i++; }
      else if (ch === '"') guil = false;
      else champ += ch;
    } else if (ch === '"') guil = true;
    else if (ch === ';') { ligne.push(champ); champ = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && texte[i + 1] === '\n') i++;
      ligne.push(champ); lignes.push(ligne); ligne = []; champ = '';
    } else champ += ch;
  }
  if (champ || ligne.length) { ligne.push(champ); lignes.push(ligne); }
  return lignes;
}

function maintenant_() { return new Date().toISOString(); }
function texte_(v) { v = String(v == null ? '' : v).slice(0, 45000); return /^[=+\-@]/.test(v) ? "'" + v : v; }
