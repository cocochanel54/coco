/**
 * Carte des tickets Praxedo, branchée sur ce Google Sheet.
 *
 * Onglets (créés par installer()) :
 *   Tickets             : le dernier export Praxedo importé (menu Carte tickets → Importer un export Praxedo)
 *   Mails Praxedo       : une ligne par mail « [CREATION INTERVENTION] » ou « [FIN INTERVENTION] » reçu,
 *                         lu dans Gmail toutes les 5 minutes (une règle Outlook redirige ces mails vers ce Gmail)
 *   Adresses            : les coordonnées des adresses (service d'adresses de l'IGN), pour ne pas les chercher deux fois
 *
 * La carte (Déployer → Application Web) recharge ces données toutes les 2 minutes et relit Gmail au passage.
 * Toute fonction sans « _ » à la fin peut être appelée depuis la page.
 */

var ONGLET_TICKETS = 'Tickets';
var ONGLET_MAILS = 'Mails Praxedo';
var ONGLET_ADRESSES = 'Adresses';
var ENTETES_MAILS = ['Reçu le', 'Type', 'N° ticket', 'Commune', 'Intervention du', 'Lien Praxedo', 'Objet du mail', 'Id du mail'];
var COL_ID_MAIL = 8;
var ENTETES_ADRESSES = ['Adresse cherchée', 'Latitude', 'Longitude', 'Adresse trouvée', 'Score', 'Type', 'Commune', 'CP'];
// Objets des mails Praxedo :
//   « [CIRCET 54][Maintenance réseau][CREATION INTERVENTION] Ticket Exploitation - Réseau N°BB-20261026T1839#1 - NA »
//   « [CIRCET NANCY (RES) ][FIN INTERVENTION] Ticket Exploitation - Réseau N°BB-20261102T0761#0 - JARVILLE LA MALGRANGE »
// in:anywhere -in:trash : les mails redirigés que Gmail range dans les spams sont lus aussi.
var RECHERCHE_GMAIL = 'subject:("FIN INTERVENTION" OR "CREATION INTERVENTION" OR "CRÉATION INTERVENTION") in:anywhere -in:trash newer_than:30d';
var RE_NUM = /N°\s*([A-Z]{2,4}-\d{8}T\d{4}(?:#\d+)?|\d{5,})/i;
var RE_FIN = /\bFIN\s+INTERVENTION/i;
var RE_CREATION = /CR[EÉ]ATION\s+INTERVENTION/i;
// Une commune est cherchée au centre de la commune, de préférence autour de Nancy (homonymes ailleurs en France).
var PREFIXE_COMMUNE = 'Commune : ';
var CENTRE_RECHERCHE = { lat: 48.69, lon: 6.18 };

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
   [ONGLET_MAILS, ENTETES_MAILS], [ONGLET_ADRESSES, ENTETES_ADRESSES]].forEach(function (o) {
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
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_TICKETS);
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

/** Lit les mails de création et de fin d'intervention arrivés dans Gmail (déclencheur toutes les 5 minutes). */
function lireMails() {
  var verrou = LockService.getScriptLock();
  if (!verrou.tryLock(20000)) return 0;
  try {
    var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_MAILS);
    var vus = {};
    if (f.getLastRow() > 1) f.getRange(2, COL_ID_MAIL, f.getLastRow() - 1, 1).getValues().forEach(function (r) { vus[r[0]] = 1; });
    var ajout = [];
    GmailApp.search(RECHERCHE_GMAIL, 0, 100).forEach(function (fil) {
      fil.getMessages().forEach(function (m) {
        if (vus[m.getId()]) return;
        var l = ligneMail_(m.getSubject(), m.getPlainBody(), m.getDate(), m.getId());
        if (l) { ajout.push(l.map(texte_)); vus[m.getId()] = 1; }
      });
    });
    if (ajout.length) f.getRange(f.getLastRow() + 1, 1, ajout.length, ENTETES_MAILS.length).setNumberFormat('@').setValues(ajout);
    PropertiesService.getScriptProperties().setProperty('LU_LE', maintenant_());
    placerCommunes_(ajout);
    return ajout.length;
  } finally {
    verrou.releaseLock();
  }
}

/** Pour la macro Outlook, si la DSI bloque le transfert vers Gmail : reçoit l'objet et le texte d'un mail. */
function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  if (!d.cle || d.cle !== PropertiesService.getScriptProperties().getProperty('CLE_MACRO')) {
    return ContentService.createTextOutput('cle refusee');
  }
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_MAILS);
  var id = 'outlook:' + d.id;
  var deja = f.getLastRow() > 1 && f.getRange(2, COL_ID_MAIL, f.getLastRow() - 1, 1).getValues().some(function (r) { return r[0] === id; });
  var l = deja ? null : ligneMail_(d.objet || '', d.corps || '', d.recu ? new Date(d.recu) : new Date(), id);
  if (l) { l = l.map(texte_); f.appendRow(l); placerCommunes_([l]); }
  return ContentService.createTextOutput(l ? 'ok' : 'ignore');
}

/** Les tickets créés par mail ne sont pas encore dans l'export : on les place au centre de leur commune. */
function placerCommunes_(lignes) {
  var communes = lignes.filter(function (l) { return l[1] === 'Création' && l[3]; }).map(function (l) { return PREFIXE_COMMUNE + l[3]; });
  if (!communes.length) return;
  try { geocoder_(communes); } catch (err) { console.warn(err); }  // réessayé à la prochaine création dans la même commune
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
  var fins = [], creations = [];
  ss.getSheetByName(ONGLET_MAILS).getDataRange().getDisplayValues().slice(1).forEach(function (r) {
    var m = { recu: r[0], num: r[2], commune: r[3], jour: r[4], lien: r[5] };
    if (r[1] === 'Fin') { m.rapport = r[5]; fins.push(m); }
    else if (r[1] === 'Création') creations.push(m);
  });
  var adresses = {};
  ss.getSheetByName(ONGLET_ADRESSES).getDataRange().getValues().slice(1).forEach(function (r) {
    adresses[r[0]] = r[1] === '' ? null : { lat: +r[1], lon: +r[2], label: r[3], score: +r[4], type: r[5], city: r[6], postcode: String(r[7]) };
  });
  return { entetes: t[0], lignes: t.slice(1), fins: fins, creations: creations, adresses: adresses,
           importeLe: props.getProperty('IMPORTE_LE') || '', luLe: props.getProperty('LU_LE') || '' };
}

// ---------- outils ----------

/**
 * [reçu le, type (Création / Fin), N°, commune, intervention du, lien Praxedo, objet, id]
 * ou null si ce n'est pas un mail de création ou de fin, ou si l'objet n'a pas de n° de ticket.
 */
function ligneMail_(objet, corps, recu, id) {
  var type = RE_FIN.test(objet) ? 'Fin' : RE_CREATION.test(objet) ? 'Création' : '';
  var m = RE_NUM.exec(objet);
  if (!type || !m) return null;
  var commune = objet.slice(m.index + m[0].length).replace(/^\s*-\s*/, '').trim();
  if (/^(N\/?A|NC|-)?$/i.test(commune)) commune = '';  // « - NA » : commune pas renseignée
  var jour = type === 'Fin' ? /faite ce jour\s+(\d{2}\/\d{2}\/\d{4})/i.exec(corps) : null;
  var lien = /https?:\/\/[^\s<>"]*praxedo[^\s<>"]*\/render/i.exec(corps) || /https?:\/\/[^\s<>"]*praxedo[^\s<>"]*/i.exec(corps);
  return [Utilities.formatDate(recu, 'Europe/Paris', 'dd/MM/yyyy HH:mm'), type, m[1].toUpperCase(), commune,
          jour ? jour[1] : '', lien ? lien[0].replace(/[.,;:)\]]+$/, '') : '', objet, id];
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
      if (q.indexOf(PREFIXE_COMMUNE) === 0) {
        return { url: 'https://data.geopf.fr/geocodage/search?limit=1&type=municipality&lat=' + CENTRE_RECHERCHE.lat + '&lon=' + CENTRE_RECHERCHE.lon +
                      '&q=' + encodeURIComponent(q.slice(PREFIXE_COMMUNE.length)), muteHttpExceptions: true };
      }
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
  texte = String(texte).replace(/^﻿/, '');
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
