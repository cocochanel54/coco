/**
 * Carte des tickets Praxedo, branchée sur ce Google Sheet.
 *
 * Onglets (créés par installer()) :
 *   Tickets             : le dernier export Praxedo importé (menu Carte tickets → Importer un export Praxedo)
 *   Fins d'intervention : une ligne par mail « [FIN INTERVENTION] » reçu, lu dans Gmail toutes les 5 minutes
 *                         (une règle Outlook transfère ces mails vers ce Gmail). Le rapport d'intervention (PDF du lien
 *                         du mail) est lu ensuite : horaires réels, défaut corrigé ou non, réparation définitive ou
 *                         non, dommage réseau signalé, compte-rendu du technicien.
 *   Créations           : une ligne par mail « [CREATION INTERVENTION] » : le ticket apparaît sur la carte
 *                         (en violet, avec sa GTR) jusqu'à ce qu'un export le contienne
 *   Sortis de l'export  : les tickets présents dans l'import précédent et absents du nouveau (terminés ou annulés)
 *   Suites à prévoir    : les interventions clôturées en provisoire, non corrigées ou avec un dommage signalé, et leur
 *                         ticket de suite quand il existe
 *   Clôtures            : chaque intervention clôturée, avec ses horaires, sa durée, sa GTR respectée ou non (de la demande
 *                         à la fin de l'intervention), son résultat, et la fin d'impact client pour information
 *   Stats boîtes, Stats villes, Stats semaines : d'après les rapports lus, où on intervient le plus, le temps passé, les provisoires,
 *                         la part de GTR respectées (sur les 90 derniers jours)
 *   Incohérences, Tickets liés, PM chargés : les listes de contrôle de l'Excel, refaites à chaque import,
 *                         à chaque mail de création ou de fin, et au moins une fois par heure
 *   Notes               : tes notes par ticket (N° ticket, Note). Elles restent d'un import à l'autre, alors que
 *                         l'onglet Tickets est remplacé à chaque import. « GTR justifiée car poteau », « poteau à
 *                         remplacer », « attente Orange », « relance Orange 12/11 », « remplacement prévu le 25/11 »…
 *                         passent le ticket en gris avec son suivi.
 *   Accès               : les codes à taper pour ouvrir la carte (techniciens, conduite d'activité)
 *   Adresses            : les coordonnées des adresses (service d'adresses de l'IGN), pour ne pas les chercher deux fois
 *
 * Import automatique : un mail reçu dans Gmail avec l'export Praxedo en pièce jointe (.csv) est importé tout seul
 * dans les 5 minutes (export programmé par Praxedo, ou export que tu t'envoies par mail).
 *
 * Les fins d'intervention arrivent de trois façons : un statut de fin dans l'export (Terminée, Réalisée, Validée…),
 * un ticket qui sort du filtre entre deux imports, ou un mail « [FIN INTERVENTION] » transféré à la main vers Gmail.
 *
 * La carte (Déployer → Application Web) recharge ces données toutes les 2 minutes et relit Gmail au passage.
 * Elle demande un code (onglet Accès) et ne permet de rien modifier. Toute fonction sans « _ » à la fin peut être
 * appelée depuis la page : celles qui modifient le Sheet vérifient que c'est son propriétaire qui les lance.
 */

var ONGLET_TICKETS = 'Tickets';
var ONGLET_FINS = "Fins d'intervention";
var ONGLET_ADRESSES = 'Adresses';
var ONGLET_SORTIS = "Sortis de l'export";
var ONGLET_CREATIONS = 'Créations';
var ONGLET_INCOHERENCES = 'Incohérences', ONGLET_LIES = 'Tickets liés', ONGLET_PM = 'PM chargés', ONGLET_SUITES = 'Suites à prévoir';
var ONGLET_QUALITE = 'Qualité';
var ONGLET_BACKLOG = 'Backlog', ENTETES_BACKLOG = ['Le', 'Créés', 'Clôturés', 'Sortis de l\'export', 'Annulés', 'Regroupés', 'Tickets ouverts', 'dont à planifier',
  'dont GTR KO', 'dont GTR < 48 h', 'dont planifiés', 'dont planifiés GTR KO', 'dont justifiés', 'Ouverts depuis plus de 30 j'];
var ONGLET_CLOTURES = 'Clôtures', ONGLET_STATS_BOITES = 'Stats boîtes', ONGLET_STATS_VILLES = 'Stats villes', ONGLET_STATS_SEMAINES = 'Stats semaines';
var ONGLET_ACCES = 'Accès';
var ONGLET_ANNULES = 'Annulés', ENTETES_ANNULES = ['N° ticket', 'Annulé le', 'Pourquoi'];
var ONGLET_CORRELES = 'Corrélés', ENTETES_CORRELES = ['N° ticket regroupé', 'Sous le ticket', 'Le', 'Par / pourquoi'];
var ONGLET_NOTES = 'Notes', ENTETES_NOTES = ['N° ticket', 'Note', 'Ajoutée'];
var COULEURS_GRAVITE = { '1 - Urgent': '#F8CBAD', '2 - Important': '#FFE699', '3 - À vérifier': '#DDEBF7', '4 - Info': '#EDEDED' };
var ENTETES_CREATIONS = ['Reçu le', 'N° ticket', 'Type', 'Commune (objet)', 'Adresse', 'CP', 'Ville', 'Nom client', 'Description', 'Id du mail'];
var RECHERCHE_CREATIONS = 'subject:"CREATION INTERVENTION"';
var ENTETES_SORTIS = ['Sorti le', 'N°', 'Statut avant', 'Planifiée', 'Technicien', 'Adresse', 'CP', 'Ville site', 'Site'];
var ENTETES_FINS = ['Reçu le', 'N° ticket', 'Commune', 'Intervention du', 'Rapport', 'Objet du mail', 'Id du mail',
  'Rapport lu', 'Début', 'Fin', 'Durée', 'Défaut corrigé', 'Réparation définitive', 'Dommage réseau', 'Compte-rendu du technicien',
  'Boîte(s)', 'PM (site)', 'Nature du défaut', 'Cause du défaut', 'Adresse', 'Demande le', 'Étiquettes GTR', 'Autres éléments du rapport', 'Technicien',
  "Fin d'impact", 'Photos (IA)', 'Photos analysées le', 'Photos vérifiées (OK / NO OK)'];
// colonnes de l'analyse des photos par l'IA (voir analyserPhotos_) ; « Photos vérifiées » : OK ou NO OK après vérification, écrit à la
// main dans le Sheet ou depuis la carte (code « Annulations et corrélations »), avec ou sans analyse par l'IA
var COL_PHOTOS = 'Photos (IA)', COL_PHOTOS_LE = 'Photos analysées le', COL_PHOTOS_AVIS = 'Photos vérifiées (OK / NO OK)';
var JOURS_HISTORIQUE = 90;     // la carte et les statistiques portent sur les clôtures des 90 derniers jours
var RAPPORTS_PAR_PASSAGE = 4;
var LECTURE_RAPPORTS = '2';    // à changer quand la lecture des rapports s'enrichit : ceux des 30 derniers jours sont relus une fois  // rapports PDF lus à chaque passage (les plus récents d'abord), pour garder des passages courts
var ENTETES_ADRESSES = ['Adresse cherchée', 'Latitude', 'Longitude', 'Adresse trouvée', 'Score', 'Type', 'Commune', 'CP'];
// Mails de fin d'intervention : l'objet contient « [FIN INTERVENTION] … N°BB-20261102T0761#0 - COMMUNE »
var RECHERCHE_GMAIL = 'subject:"FIN INTERVENTION"';
var RE_NUM = /N°\s*([A-Z]{2,4}-\d{8}T\d{4}(?:#\d+)?|\d{5,})/i;
// Exports Praxedo reçus par mail : toute pièce jointe .csv avec une colonne « N° » (les autres CSV sont ignorés)
var RECHERCHE_EXPORTS = 'has:attachment filename:csv';
// Google limite un compte gratuit à 20 000 lectures Gmail et 90 minutes de déclencheurs par jour : chaque passage ne
// regarde que les mails arrivés depuis le précédent (fenetre_), et une fois par heure les 2 derniers jours, par sécurité.
var MARGE_MAILS_S = 15 * 60, RELECTURE_LARGE_MS = 3600 * 1000;

/** La carte. */
function doGet() {
  try {  // l'adresse de la carte, pour le lien de l'Excel du jour
    var url = ScriptApp.getService().getUrl(), props = PropertiesService.getScriptProperties();
    if (/\/exec$/.test(url || '') && props.getProperty('LIEN_CARTE') !== url) props.setProperty('LIEN_CARTE', url);
  } catch (err) { /* sans importance pour la carte */ }
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
    .addItem('Lire les mails maintenant', 'lireMailsMaintenant')
    .addItem('Mettre à jour les onglets de contrôle', 'mettreAJourControles')
    .addItem('Fichiers du jour : carte + Excel (pour Teams)…', 'ouvrirCarteDuJour')
    .addItem('Envoi du matin : envoyer maintenant', 'envoyerMaintenant')
    .addItem('Annuler un ticket…', 'annulerTicket')
    .addItem('Regrouper des tickets (corrélation)…', 'correlerTickets')
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Rapports par mail')
      .addItem('Rapport du soir (fins d\'intervention) : envoyer maintenant', 'rapportSoirMaintenant')
      .addItem('Bilan de la semaine : envoyer maintenant', 'rapportSemaineMaintenant')
      .addItem('Bilan du mois dernier : envoyer maintenant', 'rapportMoisMaintenant'))
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Analyse des photos (IA)')
      .addItem('Activer, voir le coût, changer la clé ou arrêter…', 'reglerPhotosIA')
      .addItem('Analyser maintenant', 'analyserPhotosMaintenant'))
    .addItem("Codes d'accès à la carte", 'afficherCodes')
    .addItem('Clé pour la macro Outlook', 'afficherCle')
    .addToUi();
}

/** À lancer une fois depuis l'éditeur : crée les onglets et la lecture des mails toutes les 5 minutes. */
function installer() {
  proprietaire_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  [[ONGLET_TICKETS, ['Importe un export Praxedo : menu Carte tickets → Importer un export Praxedo']],
   [ONGLET_FINS, ENTETES_FINS], [ONGLET_CREATIONS, ENTETES_CREATIONS], [ONGLET_SORTIS, ENTETES_SORTIS], [ONGLET_NOTES, ENTETES_NOTES], [ONGLET_ADRESSES, ENTETES_ADRESSES]].forEach(function (o) {
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
  acces_();
  lireMails_(0);
}

function afficherCodes() {
  var a = acces_();
  SpreadsheetApp.getUi().alert("Codes pour ouvrir la carte :\n\nConduite d'activité : " + a.ca + '\nTechniciens : ' + a.tech +
    '\nAnnulations et corrélations (la personne qui les gère) : ' + a.annul + (a.mailsAnnul.length ? ' · par mail : ' + a.mailsAnnul.join(', ') : ' · son adresse mail : colonne D de l\'onglet Accès') +
    "\n\nPour les changer (et couper l'accès à ceux qui ont l'ancien), modifie-les dans l'onglet « Accès ».");
}

function ouvrirImport() {
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutputFromFile('Import').setWidth(420).setHeight(220), 'Importer un export Praxedo');
}

/**
 * Fichiers du jour, pour les postes où Google est bloqué : la carte complète dans un seul fichier HTML, et un vrai
 * fichier Excel avec toutes les listes et les statistiques. Ce sont des photos du moment : ils ne se mettent pas à jour.
 */
function ouvrirCarteDuJour() {
  var page = '<style>[hidden]{display:none!important}body{font:14px system-ui,sans-serif;margin:12px;color:#1a1a1a}p{margin:0 0 10px}' +
    '.etat{font-weight:600}a,button{font:inherit}button{padding:6px 10px}</style>' +
    '<p>Deux fichiers à déposer dans Teams ou à ouvrir au bureau, sans lien Google ni code : la <b>carte</b> (à ouvrir dans un navigateur) ' +
    'et l\'<b>Excel</b> (à planifier, justifiés, suites, incohérences, clôtures, statistiques). Ce sont des photos de maintenant : ils ne se mettent pas à jour.</p>' +
    '<p class="etat" id="etatC" role="status">Préparation de la carte…</p><p><a id="lienC" hidden></a></p>' +
    '<p class="etat" id="etatX" role="status">Préparation de l\'Excel…</p><p><a id="lienX" hidden></a></p>' +
    '<p><button id="mail" type="button" hidden>M\'envoyer les deux par mail</button></p><p class="etat" id="etatM" role="status"></p>' +
    '<script>function $(i){return document.getElementById(i);}var prets=0,mail=$("mail");' +
    'function pret(lien,etat,b,nom,texte){lien.href=URL.createObjectURL(b);lien.download=nom;lien.textContent="Télécharger "+nom;lien.hidden=false;' +
    'etat.textContent=texte+" ("+Math.round(b.size/1024)+" Ko).";if(++prets===2)mail.hidden=false;}' +
    'function rate(etat){return function(e){etat.textContent=(e&&e.message)||String(e);mail.disabled=false;};}' +
    'google.script.run.withSuccessHandler(function(r){pret($("lienC"),$("etatC"),new Blob([r.html],{type:"text/html"}),r.nom,"Carte prête");}).withFailureHandler(rate($("etatC"))).carteDuJour();' +
    'google.script.run.withSuccessHandler(function(r){var t=atob(r.base64),o=new Uint8Array(t.length);for(var i=0;i<t.length;i++)o[i]=t.charCodeAt(i);' +
    'pret($("lienX"),$("etatX"),new Blob([o],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),r.nom,"Excel prêt");}).withFailureHandler(rate($("etatX"))).excelDuJour();' +
    'mail.onclick=function(){mail.disabled=true;$("etatM").textContent="Envoi…";google.script.run.withSuccessHandler(function(a){$("etatM").textContent="Envoyé à "+a+".";}).withFailureHandler(rate($("etatM"))).envoyerCarteDuJour();};' +
    '</' + 'script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(page).setWidth(480).setHeight(330), 'Fichiers du jour : carte + Excel');
}

/** Pour la fenêtre : le nom et le contenu du fichier de la carte. Réservé au propriétaire. */
function carteDuJour() {
  proprietaire_();
  return { nom: nomCarteDuJour_(), html: construireCarteDuJour_() };
}

/** Pour la fenêtre : le nom et le contenu (en base 64) du fichier Excel. Réservé au propriétaire. */
function excelDuJour() {
  proprietaire_();
  var b = classeurDuJour_();
  return { nom: b.getName(), base64: Utilities.base64Encode(b.getBytes()) };
}

/** Envoie la carte et l'Excel en pièces jointes à l'adresse Gmail du propriétaire (et à personne d'autre). */
function envoyerCarteDuJour() {
  proprietaire_();
  var moi = Session.getEffectiveUser().getEmail();
  GmailApp.sendEmail(moi, 'Carte et Excel du jour des tickets, ' + Utilities.formatDate(new Date(), 'Europe/Paris', "dd/MM/yyyy 'à' HH:mm"),
    'En pièces jointes :\n- la carte du jour : télécharge le fichier .html puis ouvre-le dans un navigateur ;\n' +
    '- l\'Excel du jour : à planifier, justifiés, suites à prévoir, incohérences, clôtures et statistiques.',
    { attachments: [Utilities.newBlob(construireCarteDuJour_(), 'text/html', nomCarteDuJour_()), classeurDuJour_()] });
  return moi;
}

function nomCarteDuJour_() {
  return 'Carte_du_jour_' + Utilities.formatDate(new Date(), 'Europe/Paris', "yyyy-MM-dd_HH'h'mm") + '.html';
}

/**
 * La page de la carte avec les données dedans (rôle conduite d'activité) et Leaflet recopié dans le fichier, pour
 * qu'il s'ouvre sans rien charger d'autre que le fond de carte de l'IGN. Les n° de téléphone sont masqués.
 */
function construireCarteDuJour_() {
  var d = lireDonnees_(), tel = /(?:\+33\s?|0)[1-9](?:[ .-]?\d{2}){4}/g;
  d.role = 'ca';
  d.photoLe = maintenant_();
  d.lignes = d.lignes.map(function (r) { return r.map(function (v) { return String(v).replace(tel, '[tél.]'); }); });
  d.creations.forEach(function (c) { c.desc = String(c.desc || '').replace(tel, '[tél.]'); });
  d.fins.forEach(function (x) { x.resume = String(x.resume || '').replace(tel, '[tél.]'); });
  var page = HtmlService.createTemplateFromFile('Carte').evaluate().getContent();
  try {
    var css = UrlFetchApp.fetch('https://unpkg.com/leaflet@1.9.4/dist/leaflet.css').getContentText();
    var js = UrlFetchApp.fetch('https://unpkg.com/leaflet@1.9.4/dist/leaflet.js').getContentText();
    // remplacements par fonction : le texte inséré peut contenir des « $ »
    page = page.replace('<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css">', function () { return '<style>' + css + '</style>'; })
               .replace('<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></' + 'script>', function () { return '<script>' + js.replace(/<\/script/gi, '<\\/script') + '</' + 'script>'; });
  } catch (err) {
    console.warn('Leaflet non recopié dans le fichier : ' + err);  // le fichier ira alors le chercher sur unpkg.com à l'ouverture
  }
  var donnees = JSON.stringify(d).replace(/</g, '\\u003c');
  return page.replace('<div id="map"></div>', function () { return '<script>window.CARTE_DU_JOUR = ' + donnees + ';</' + 'script><div id="map"></div>'; });
}

var MIME_XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/**
 * L'Excel du jour : un résumé chiffré, les tickets à planifier par priorité, le suivi des justifiés, puis les mêmes
 * onglets que le Sheet (suites, incohérences, tickets liés, PM, clôtures, statistiques). N° de téléphone masqués.
 */
function classeurDuJour_() { return excelDuJour_().blob; }

/** L'Excel du jour (blob) et les lignes de son résumé (pour le texte du mail du matin). */
function excelDuJour_() {
  var a = analyse_(), d = lireDonnees_(), NOW = new Date(), res = a.analyser(d, NOW), c = a.controles(res, NOW);
  var props = PropertiesService.getScriptProperties();
  var l = a.listesDuJour(res, c, NOW, { importeLe: d.importeLe, luLe: d.luLe, lienCarte: props.getProperty('LIEN_CARTE') || '',
    historique: historiqueBacklog_(), depuisSoir: props.getProperty('RAPPORT_SOIR_DEPUIS') || '' }, d), g = l.graphiques;
  var feuilles = [
    { nom: 'Résumé', entetes: l.resume.entetes, lignes: l.resume.lignes, fonds: l.resume.fonds, largeurs: [460, 220], filtre: false },
    // graphiques natifs d'Excel, comme le bilan de la carte ; leurs chiffres sont dans l'onglet « Données graphiques »
    { nom: 'Graphiques', entetes: ['Clôtures des 90 derniers jours (chiffres dans l\'onglet « Données graphiques »)'], lignes: [], largeurs: [780],
      graphiques: [
        { titre: 'Clôtures par semaine (la dernière est en cours)', sens: 'col', donnees: 'Données graphiques', cat: 1, series: [3, 4], n: g.nSemaines, ligne: 1, taille: [760, 330] },
        { titre: 'Les communes où on intervient le plus', sens: 'bar', donnees: 'Données graphiques', cat: 7, series: [8, 9], n: g.nCommunes, ligne: 19, taille: [760, 380] },
        { titre: 'Les boîtes où on intervient le plus', sens: 'bar', donnees: 'Données graphiques', cat: 12, series: [13, 14], n: g.nBoites, ligne: 39, taille: [760, 380] }] },
    { nom: 'À planifier', entetes: l.aPlanifier.entetes, lignes: l.aPlanifier.lignes, fonds: l.aPlanifier.fonds,
      largeurs: [50, 160, 150, 140, 120, 220, 100, 130, 160, 220, 90, 140, 170, 130, 120, 260, 260, 300, 200, 480] },
    { nom: 'Justifiés', entetes: l.justifies.entetes, lignes: l.justifies.lignes, fonds: l.justifies.fonds,
      largeurs: [90, 300, 120, 120, 200, 160, 100, 60, 120, 80, 300, 160, 220, 140, 420] }
  ].concat(ongletsControle_(c).map(function (o) {
    return { nom: o.nom, entetes: o.entetes, lignes: o.lignes, largeurs: o.largeurs, fonds: o.lignes.map(function (x, i) { return o.couleur ? o.couleur(x, i) || null : null; }) };
  })).concat((function () {
    var fb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_BACKLOG);
    if (!fb || fb.getLastRow() < 2) return [];
    var v = fb.getRange(2, 1, fb.getLastRow() - 1, ENTETES_BACKLOG.length).getDisplayValues().slice(-90).reverse();
    return [{ nom: ONGLET_BACKLOG, entetes: ENTETES_BACKLOG, lignes: v.map(function (r) { return r.map(function (x, i) { return i && x !== '' && !isNaN(+x) ? +x : x; }); }),
      largeurs: [130, 70, 80, 110, 80, 90, 110, 110, 100, 110, 100, 140, 100, 160] }];
  })()).concat([{ nom: 'Données graphiques', entetes: g.entetes, lignes: g.lignes, filtre: false,
    largeurs: [120, 70, 70, 120, 90, 70, 20, 200, 120, 90, 70, 20, 190, 120, 90, 70] }]);
  var tel = /(?:\+33\s?|0)[1-9](?:[ .-]?\d{2}){4}/g;
  feuilles.forEach(function (f) { f.lignes = f.lignes.map(function (r) { return r.map(function (v) { return typeof v === 'string' ? v.replace(tel, '[tél.]') : v; }); }); });
  return { resume: l.resume.lignes, blob: Utilities.newBlob(signes_(xlsx_(feuilles)), MIME_XLSX,
    'Tickets_du_jour_' + Utilities.formatDate(NOW, 'Europe/Paris', "yyyy-MM-dd_HH'h'mm") + '.xlsx') };
}

/**
 * Classeur Excel (.xlsx) fabriqué en mémoire, sans passer par le Drive : une feuille par tableau, en-tête bleu figé,
 * filtres, largeurs et couleurs de lignes comme dans le Sheet, liens « rapport » cliquables.
 * feuilles : [{ nom, entetes, lignes, fonds (couleur '#RRGGBB' ou null, une par ligne), largeurs (pixels), filtre (false : sans filtre) }].
 * Les nombres restent des nombres, le reste du texte. Renvoie les octets du fichier (Uint8Array).
 */
function xlsx_(feuilles) {
  function x(v) {
    return String(v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '')
      .replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; });
  }
  function lettre(i) { var s = ''; for (i++; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + (i - 1) % 26) + s; return s; }
  var fonds = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], numFond = {};
  var xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'], numXf = {};
  // style d'une cellule : police (0 normale, 1 en-tête blanc gras, 2 lien) et fond
  function style(police, fond) {
    var cle = police + '|' + fond;
    if (numXf[cle] == null) {
      if (fond && numFond[fond] == null) {
        fonds.push('<fill><patternFill patternType="solid"><fgColor rgb="FF' + fond + '"/><bgColor indexed="64"/></patternFill></fill>');
        numFond[fond] = fonds.length - 1;
      }
      xfs.push('<xf numFmtId="0" fontId="' + police + '" fillId="' + (fond ? numFond[fond] : 0) + '" borderId="1" xfId="0" applyFont="1" applyFill="1" ' +
        'applyBorder="1" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>');
      numXf[cle] = xfs.length - 1;
    }
    return numXf[cle];
  }
  var vus = {}, filtres = [], DEBUT = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  var REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/', RELS = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
  var NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  feuilles.forEach(function (fe, n) {
    // nom d'onglet : 31 caractères au plus, sans : \ / ? * [ ], et unique
    var nom = String(fe.nom).replace(/[:\\\/?*\[\]]/g, ' ').slice(0, 31).trim() || 'Feuille ' + (n + 1);
    while (vus[nom.toLowerCase()]) nom = nom.slice(0, 27) + ' (' + (n + 1) + ')';
    vus[nom.toLowerCase()] = 1;
    fe.nomXl = nom;
    fe.graphes = (fe.graphiques || []).filter(function (g) { return g.n > 0; });
  });
  var docs = feuilles.map(function (fe, n) {
    var nom = fe.nomXl;
    var nc = fe.entetes.length, tout = [fe.entetes].concat(fe.lignes), rangees = [];
    tout.forEach(function (l, i) {
      var fond = i === 0 ? '1F4E78' : String((fe.fonds && fe.fonds[i - 1]) || '').replace('#', '').toUpperCase(), cel = [];
      for (var j = 0; j < nc; j++) {
        var v = l[j], ref = lettre(j) + (i + 1);
        if (i === 0) { if (v !== '' && v != null) cel.push('<c r="' + ref + '" s="' + style(1, fond) + '" t="inlineStr"><is><t>' + x(v) + '</t></is></c>'); continue; }
        if (v == null || v === '') { if (fond) cel.push('<c r="' + ref + '" s="' + style(0, fond) + '"/>'); continue; }
        if (typeof v === 'number' && isFinite(v)) { cel.push('<c r="' + ref + '" s="' + style(0, fond) + '"><v>' + v + '</v></c>'); continue; }
        v = String(v);
        if (/^https:\/\/[^\s"]+$/.test(v) && v.length < 250) {
          var lib = { 'Rapport': 'rapport', 'Google Maps': 'ouvrir', 'Valeur': 'ouvrir' }[fe.entetes[j]] || 'lien';
          cel.push('<c r="' + ref + '" s="' + style(2, fond) + '" t="str"><f>' + x('HYPERLINK("' + v + '","' + lib + '")') + '</f><v>' + lib + '</v></c>');
          continue;
        }
        cel.push('<c r="' + ref + '" s="' + style(0, fond) + '" t="inlineStr"><is><t xml:space="preserve">' + x(v.slice(0, 32000)) + '</t></is></c>');
      }
      rangees.push('<row r="' + (i + 1) + '">' + cel.join('') + '</row>');
    });
    var zone = 'A1:' + lettre(nc - 1) + tout.length;
    var filtre = fe.lignes.length > 0 && fe.filtre !== false;
    if (filtre) filtres.push('<definedName name="_xlnm._FilterDatabase" localSheetId="' + n + '" hidden="1">\'' + x(nom.replace(/'/g, "''")) +
      '\'!$A$1:$' + lettre(nc - 1) + '$' + tout.length + '</definedName>');
    return DEBUT + '<worksheet ' + NS + '><dimension ref="' + zone + '"/>' +
      '<sheetViews><sheetView workbookViewId="0"' + (n === 0 ? ' tabSelected="1"' : '') + '>' +
      '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/><cols>' + fe.entetes.map(function (h, j) {
        var px = (fe.largeurs && fe.largeurs[j]) || 120;
        return '<col min="' + (j + 1) + '" max="' + (j + 1) + '" width="' + Math.max(4, Math.round((px - 5) / 7 * 100) / 100) + '" customWidth="1"/>';
      }).join('') + '</cols><sheetData>' + rangees.join('') + '</sheetData>' +
      (filtre ? '<autoFilter ref="' + zone + '"/>' : '') +
      '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>' +
      '<ignoredErrors><ignoredError sqref="' + zone + '" numberStoredAsText="1"/></ignoredErrors>' +
      (fe.graphes.length ? '<drawing r:id="rId1"/>' : '') + '</worksheet>';
  });
  // graphiques : un dessin par onglet qui en a, un fichier par graphique
  var fichiersGraphiques = [], dessins = 0, numGraphe = 0, typesEnPlus = '';
  feuilles.forEach(function (fe, n) {
    if (!fe.graphes.length) return;
    var d = ++dessins, ancres = [], liensDessin = [];
    fe.graphes.forEach(function (g, k) {
      var num = ++numGraphe;
      fichiersGraphiques.push(['xl/charts/chart' + num + '.xml', graphique(g)]);
      typesEnPlus += '<Override PartName="/xl/charts/chart' + num + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>';
      liensDessin.push('<Relationship Id="rId' + (k + 1) + '" Type="' + REL + 'chart" Target="../charts/chart' + num + '.xml"/>');
      ancres.push('<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + g.ligne + '</xdr:row><xdr:rowOff>95250</xdr:rowOff></xdr:from>' +
        '<xdr:ext cx="' + g.taille[0] * 9525 + '" cy="' + g.taille[1] * 9525 + '"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr>' +
        '<xdr:cNvPr id="' + (k + 2) + '" name="Graphique ' + (k + 1) + '"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>' +
        '<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
        '<c:chart r:id="rId' + (k + 1) + '"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>');
    });
    fichiersGraphiques.push(['xl/drawings/drawing' + d + '.xml', DEBUT + '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="' + REL.replace(/\/$/, '') + '" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">' +
      ancres.join('') + '</xdr:wsDr>']);
    fichiersGraphiques.push(['xl/drawings/_rels/drawing' + d + '.xml.rels', DEBUT + RELS + liensDessin.join('') + '</Relationships>']);
    fichiersGraphiques.push(['xl/worksheets/_rels/sheet' + (n + 1) + '.xml.rels', DEBUT + RELS +
      '<Relationship Id="rId1" Type="' + REL + 'drawing" Target="../drawings/drawing' + d + '.xml"/></Relationships>']);
    typesEnPlus += '<Override PartName="/xl/drawings/drawing' + d + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>';
  });
  var classeur = DEBUT + '<workbook ' + NS + '><bookViews><workbookView activeTab="0"/></bookViews><sheets>' +
    feuilles.map(function (fe, n) { return '<sheet name="' + x(fe.nomXl) + '" sheetId="' + (n + 1) + '" r:id="rId' + (n + 1) + '"/>'; }).join('') + '</sheets>' +
    (filtres.length ? '<definedNames>' + filtres.join('') + '</definedNames>' : '') + '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>';
  var liens = DEBUT + RELS + feuilles.map(function (fe, n) {
    return '<Relationship Id="rId' + (n + 1) + '" Type="' + REL + 'worksheet" Target="worksheets/sheet' + (n + 1) + '.xml"/>';
  }).join('') + '<Relationship Id="rId' + (feuilles.length + 1) + '" Type="' + REL + 'styles" Target="styles.xml"/></Relationships>';
  var police = function (debut, couleur) { return '<font>' + debut + '<sz val="10"/>' + (couleur ? '<color rgb="' + couleur + '"/>' : '') + '<name val="Arial"/><family val="2"/></font>'; };
  var gris = '<color rgb="FFD9D9D9"/>';
  var styles = DEBUT + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="3">' + police('') + police('<b/>', 'FFFFFFFF') + police('<u/>', 'FF0563C1') + '</fonts>' +
    '<fills count="' + fonds.length + '">' + fonds.join('') + '</fills>' +
    '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin">' + gris + '</left><right style="thin">' + gris +
    '</right><top style="thin">' + gris + '</top><bottom style="thin">' + gris + '</bottom><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="' + xfs.length + '">' + xfs.join('') + '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  var types = DEBUT + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
    feuilles.map(function (fe, n) {
      return '<Override PartName="/xl/worksheets/sheet' + (n + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    }).join('') + typesEnPlus + '</Types>';
  var racine = DEBUT + RELS + '<Relationship Id="rId1" Type="' + REL + 'officeDocument" Target="xl/workbook.xml"/></Relationships>';
  return zip_([['[Content_Types].xml', types], ['_rels/.rels', racine], ['xl/workbook.xml', classeur], ['xl/_rels/workbook.xml.rels', liens],
    ['xl/styles.xml', styles]].concat(docs.map(function (doc, n) { return ['xl/worksheets/sheet' + (n + 1) + '.xml', doc]; })).concat(fichiersGraphiques));

  /**
   * Un graphique en barres empilées (sens 'col' : colonnes par semaine ; 'bar' : barres horizontales, la plus grande en haut).
   * Les données viennent d'un autre onglet (g.donnees), ligne 2 et suivantes ; les valeurs sont aussi recopiées dans le
   * graphique pour qu'il s'affiche tout de suite. Couleurs : bleu « sans suite à donner », orange « à reprendre ».
   */
  function graphique(g) {
    var src = feuilles.filter(function (f) { return f.nom === g.donnees; })[0], feuille = "'" + src.nomXl.replace(/'/g, "''") + "'!", barres = g.sens === 'bar';
    var lignes = src.lignes.slice(0, g.n), COUL = ['2A78D6', 'EB6834'];
    function plage(col) { return feuille + '$' + lettre(col) + '$2:$' + lettre(col) + '$' + (g.n + 1); }
    function police(taille, couleur, gras) {
      return '<a:defRPr sz="' + taille + '"' + (gras ? ' b="1"' : '') + '><a:solidFill><a:srgbClr val="' + couleur + '"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr>';
    }
    function texte(taille, couleur) { return '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>' + police(taille, couleur) + '</a:pPr><a:endParaRPr lang="fr-FR"/></a:p></c:txPr>'; }
    var series = g.series.map(function (col, k) {
      return '<c:ser><c:idx val="' + k + '"/><c:order val="' + k + '"/><c:tx><c:strRef><c:f>' + x(feuille + '$' + lettre(col) + '$1') + '</c:f>' +
        '<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + x(src.entetes[col]) + '</c:v></c:pt></c:strCache></c:strRef></c:tx>' +
        '<c:spPr><a:solidFill><a:srgbClr val="' + COUL[k % 2] + '"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>' +
        '<c:invertIfNegative val="0"/><c:cat><c:strRef><c:f>' + x(plage(g.cat)) + '</c:f><c:strCache><c:ptCount val="' + g.n + '"/>' +
        lignes.map(function (l, i) { return '<c:pt idx="' + i + '"><c:v>' + x(l[g.cat]) + '</c:v></c:pt>'; }).join('') + '</c:strCache></c:strRef></c:cat>' +
        '<c:val><c:numRef><c:f>' + x(plage(col)) + '</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="' + g.n + '"/>' +
        lignes.map(function (l, i) { return '<c:pt idx="' + i + '"><c:v>' + (+l[col] || 0) + '</c:v></c:pt>'; }).join('') + '</c:numCache></c:numRef></c:val></c:ser>';
    }).join('');
    return DEBUT + '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart>' +
      '<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>' + police(1200, '0B0B0B', true) + '</a:pPr><a:r><a:rPr lang="fr-FR" sz="1200" b="1">' +
      '<a:solidFill><a:srgbClr val="0B0B0B"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>' + x(g.titre) + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>' +
      '<c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:barChart><c:barDir val="' + (barres ? 'bar' : 'col') + '"/><c:grouping val="stacked"/><c:varyColors val="0"/>' +
      series + '<c:gapWidth val="' + (barres ? 80 : 150) + '"/><c:overlap val="100"/><c:axId val="1001"/><c:axId val="1002"/></c:barChart>' +
      '<c:catAx><c:axId val="1001"/><c:scaling><c:orientation val="' + (barres ? 'maxMin' : 'minMax') + '"/></c:scaling><c:delete val="0"/>' +
      '<c:axPos val="' + (barres ? 'l' : 'b') + '"/><c:numFmt formatCode="General" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/>' +
      '<c:tickLblPos val="nextTo"/><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="BFBFBF"/></a:solidFill></a:ln></c:spPr>' + texte(900, '52514E') +
      '<c:crossAx val="1002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>' +
      '<c:valAx><c:axId val="1002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="' + (barres ? 'b' : 'l') + '"/>' +
      '<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="E6E6E6"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>' +
      '<c:numFmt formatCode="0" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>' +
      '<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>' + texte(900, '52514E') + '<c:crossAx val="1001"/><c:crosses val="' + (barres ? 'max' : 'autoZero') + '"/>' +
      '<c:crossBetween val="between"/></c:valAx><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr></c:plotArea>' +
      '<c:legend><c:legendPos val="t"/><c:overlay val="0"/>' + texte(900, '0B0B0B') + '</c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>' +
      '<c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>' +
      '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:latin typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="fr-FR"/></a:p></c:txPr></c:chartSpace>';
  }
}

/** Archive ZIP sans compression (c'est ce qu'est un .xlsx) : fichiers = [[nom, texte]]. Renvoie un Uint8Array. */
function zip_(fichiers) {
  var table = [];
  for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; table[n] = c >>> 0; }
  var d = new Date(), heure = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
  var jour = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate(), taille = 22;
  var f = fichiers.map(function (x) {
    var nom = utf8_(x[0]), o = utf8_(x[1]), crc = 0xFFFFFFFF;
    for (var i = 0; i < o.length; i++) crc = table[(crc ^ o[i]) & 255] ^ (crc >>> 8);
    taille += 76 + 2 * nom.length + o.length;
    return { nom: nom, o: o, crc: (crc ^ 0xFFFFFFFF) >>> 0 };
  });
  var out = new Uint8Array(taille), v = new DataView(out.buffer), p = 0;
  function ecrire(champs) { champs.forEach(function (c) { if (c[1] === 4) v.setUint32(p, c[0], true); else v.setUint16(p, c[0], true); p += c[1]; }); }
  f.forEach(function (x) {
    x.pos = p;
    ecrire([[0x04034b50, 4], [20, 2], [0x0800, 2], [0, 2], [heure, 2], [jour, 2], [x.crc, 4], [x.o.length, 4], [x.o.length, 4], [x.nom.length, 2], [0, 2]]);
    out.set(x.nom, p); p += x.nom.length;
    out.set(x.o, p); p += x.o.length;
  });
  var central = p;
  f.forEach(function (x) {
    ecrire([[0x02014b50, 4], [20, 2], [20, 2], [0x0800, 2], [0, 2], [heure, 2], [jour, 2], [x.crc, 4], [x.o.length, 4], [x.o.length, 4],
      [x.nom.length, 2], [0, 2], [0, 2], [0, 2], [0, 2], [0, 4], [x.pos, 4]]);
    out.set(x.nom, p); p += x.nom.length;
  });
  ecrire([[0x06054b50, 4], [0, 2], [0, 2], [f.length, 2], [f.length, 2], [p - central, 4], [central, 4], [0, 2]]);
  return out;
}

/** Texte → octets UTF-8 (un demi-caractère isolé devient « � »). */
function utf8_(s) {
  s = String(s);
  var o = new Uint8Array(s.length * 3), n = 0;
  for (var i = 0; i < s.length; i++) {
    var c = s.charCodeAt(i);
    if (c >= 0xD800 && c <= 0xDBFF && i + 1 < s.length && s.charCodeAt(i + 1) >= 0xDC00 && s.charCodeAt(i + 1) <= 0xDFFF) {
      c = 0x10000 + ((c - 0xD800) << 10) + (s.charCodeAt(++i) - 0xDC00);
    } else if (c >= 0xD800 && c <= 0xDFFF) c = 0xFFFD;
    if (c < 0x80) o[n++] = c;
    else if (c < 0x800) { o[n++] = 0xC0 | (c >> 6); o[n++] = 0x80 | (c & 63); }
    else if (c < 0x10000) { o[n++] = 0xE0 | (c >> 12); o[n++] = 0x80 | ((c >> 6) & 63); o[n++] = 0x80 | (c & 63); }
    else { o[n++] = 0xF0 | (c >> 18); o[n++] = 0x80 | ((c >> 12) & 63); o[n++] = 0x80 | ((c >> 6) & 63); o[n++] = 0x80 | (c & 63); }
  }
  return o.subarray(0, n);
}

/** Octets 0..255 → -128..127, comme les attend Utilities.newBlob. */
function signes_(o) { var a = new Array(o.length); for (var i = 0; i < o.length; i++) a[i] = o[i] > 127 ? o[i] - 256 : o[i]; return a; }

function afficherCle() {
  SpreadsheetApp.getUi().alert('Clé à coller dans la macro Outlook :\n\n' + PropertiesService.getScriptProperties().getProperty('CLE_MACRO'));
}

/** Remplace l'onglet Tickets par l'export (texte du CSV) et cherche les adresses nouvelles. Réservé au propriétaire. */
function importerExport(texte) {
  proprietaire_();
  return importerExport_(texte);
}

function importerExport_(texte) {
  var lignes = lireCsv_(texte);
  if (lignes.length < 2 || lignes[0].indexOf('N°') < 0) throw new Error("Ce fichier n'est pas un export d'interventions Praxedo (pas de colonne « N° »).");
  var entetes = lignes[0];
  while (entetes.length && !entetes[entetes.length - 1]) entetes.pop();
  var n = entetes.length;
  var valeurs = lignes.filter(function (l, i) { return i === 0 || l.some(function (v) { return v; }); })
    .map(function (l) { var r = l.slice(0, n); while (r.length < n) r.push(''); return r.map(texte_); });
  var ss = SpreadsheetApp.getActiveSpreadsheet(), f = ss.getSheetByName(ONGLET_TICKETS);
  var anciens = f.getDataRange().getDisplayValues();
  // La date de création n'a besoin d'être dans l'export qu'une fois : elle est gardée ici, par n° de ticket, pour les
  // exports suivants qui n'ont pas la colonne « Créée » ou qui la laissent vide.
  var ia = anciens[0].indexOf('Créée'), na = anciens[0].indexOf('N°'), connues = {};
  if (ia >= 0 && na >= 0) anciens.slice(1).forEach(function (r) { if (r[na] && r[ia]) connues[r[na]] = r[ia]; });
  var iNum = entetes.indexOf('N°'), iCree = entetes.indexOf('Créée');
  if (iCree < 0) { entetes.push('Créée'); iCree = n; n++; valeurs.forEach(function (r, i) { r.push(i ? '' : 'Créée'); }); }
  valeurs.forEach(function (r, i) { if (i && !r[iCree] && connues[r[iNum]]) r[iCree] = connues[r[iNum]]; });
  noterSortis_(ss, anciens, valeurs);
  f.clear();
  grille_(f, valeurs.length, n).getRange(1, 1, valeurs.length, n).setNumberFormat('@').setValues(valeurs);
  f.getRange(1, 1, 1, n).setFontWeight('bold');
  f.setFrozenRows(1);
  PropertiesService.getScriptProperties().setProperty('IMPORTE_LE', maintenant_());
  var c = {}; entetes.forEach(function (x, i) { c[x] = i; });
  var nouvelles = geocoder_(valeurs.slice(1).map(function (r) {
    return [r[c['Adresse']], r[c['CP']], r[c['Ville site']]].filter(function (x) { return x; }).join(' ');
  }));
  var avertissements = [];
  try { avertissements = majControles_().avertissements; } catch (err) { console.warn('Contrôles : ' + err); }
  return { tickets: valeurs.length - 1, adresses: nouvelles, avertissements: avertissements };
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
  if (tout.length) grille_(feuille, tout.length + 1, ENTETES_SORTIS.length).getRange(2, 1, tout.length, ENTETES_SORTIS.length).setNumberFormat('@').setValues(tout);
}

/** Déclencheur toutes les 5 minutes : lit les mails arrivés dans Gmail. Au plus une lecture par minute. */
function lireMails() {
  var n = lireMails_(60);
  try { analyserPhotos_(); } catch (err) { console.warn('Photos (IA) : ' + err); }
  try { envoiDuMatin_(); } catch (err) { console.warn('Envoi du matin : ' + err); }
  try { rapportDuSoir_(); } catch (err) { console.warn('Rapport du soir : ' + err); }
  return n;
}

/** Menu « Lire les mails maintenant ». Réservé au propriétaire. */
function lireMailsMaintenant() {
  proprietaire_();
  return lireMails_(0);
}

/** Fins d'intervention, créations et exports reçus. delaiMin : ne rien faire si la dernière lecture a moins de delaiMin secondes. */
function lireMails_(delaiMin) {
  var verrou = LockService.getScriptLock();
  if (!verrou.tryLock(20000)) return 0;
  try {
    var props = PropertiesService.getScriptProperties(), lu = props.getProperty('LU_LE');
    if (delaiMin && lu && Date.now() - new Date(lu).getTime() < delaiMin * 1000) return 0;
    var debut = Date.now(), fen = fenetre_(props), complet = true;
    var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_FINS);
    var vus = {};
    if (f.getLastRow() > 1) f.getRange(2, 7, f.getLastRow() - 1, 1).getValues().forEach(function (r) { vus[r[0]] = 1; });
    var ajout = [];
    messages_(RECHERCHE_GMAIL + fen.q, 100).forEach(function (m) {
      if (vus[m.getId()]) return;
      if (/annul|corr[ée]l/i.test(m.getSubject())) return;
      var l = ligneFin_(m.getSubject(), m.getPlainBody(), m.getDate(), m.getId());
      if (l) { while (l.length < ENTETES_FINS.length) l.push(''); ajout.push(l); vus[m.getId()] = 1; }
    });
    if (ajout.length) grille_(f, f.getLastRow() + ajout.length, ENTETES_FINS.length).getRange(f.getLastRow() + 1, 1, ajout.length, ENTETES_FINS.length).setNumberFormat('@').setValues(ajout);
    var creees = 0, importe = null, rapports = 0;
    try { rapports = lireRapports_(RAPPORTS_PAR_PASSAGE); } catch (err) { console.warn('Rapports : ' + err); }
    try { creees = lireCreations_(fen.q); } catch (err) { complet = false; console.warn('Créations : ' + err); }
    try { creees += relireCreations_(); } catch (err) { console.warn('Relecture des créations : ' + err); }
    try { creees += lireAnnulations_(fen.q); } catch (err) { console.warn('Annulations par mail : ' + err); }
    try { creees += lireCorrelations_(fen.q); } catch (err) { console.warn('Corrélations par mail : ' + err); }
    try { importe = importerDepuisGmail_(fen.q); } catch (err) { complet = false; console.warn('Import par mail : ' + err); }
    // l'import refait déjà les contrôles ; sinon, après un mail de fin ou de création, et au moins une fois par heure
    var faits = props.getProperty('CONTROLES_LE');
    if (!importe && (ajout.length || creees || rapports || !faits || Date.now() - new Date(faits).getTime() > 3600 * 1000)) {
      try { majControles_(); } catch (err) { console.warn('Contrôles : ' + err); }
    }
    if (complet) {  // en cas d'erreur, le prochain passage reprend la même période
      props.setProperty('MAILS_DEPUIS', String(debut));
      if (fen.large) props.setProperty('MAILS_LARGE', String(debut));
    }
    props.setProperty('LU_LE', maintenant_());
    return ajout.length;
  } finally {
    verrou.releaseLock();
  }
}

/**
 * La période de mails à regarder. En temps normal : depuis la dernière lecture réussie, avec 15 minutes de marge
 * (« after: » en secondes). Une fois par heure, et la première fois : les 2 derniers jours (ou depuis la dernière
 * lecture si le script est resté arrêté plus longtemps, 30 jours au plus), pour ne rien manquer.
 */
function fenetre_(props) {
  var depuis = +props.getProperty('MAILS_DEPUIS') || 0, large = +props.getProperty('MAILS_LARGE') || 0, now = Date.now();
  if (!depuis || now - large > RELECTURE_LARGE_MS) {
    var jours = depuis ? Math.min(30, Math.max(2, Math.ceil((now - depuis) / 86400000) + 1)) : 3;
    return { q: ' newer_than:' + jours + 'd', large: true };
  }
  return { q: ' after:' + (Math.floor(depuis / 1000) - MARGE_MAILS_S), large: false };
}

/** Les messages des fils trouvés, lus en un seul appel à Gmail (au lieu d'un appel par fil). */
function messages_(recherche, max) {
  var fils = GmailApp.search(recherche, 0, max);
  if (!fils.length) return [];
  var out = [];
  GmailApp.getMessagesForThreads(fils).forEach(function (ms) { ms.forEach(function (m) { out.push(m); }); });
  return out;
}

/**
 * Annulations par mail : la personne qui gère les annulations transfère à ce Gmail le mail de création du ticket (ou écrit un
 * mail) avec « ANNULÉ » dans l'objet. Le n° est pris dans l'objet, sinon dans le texte du mail. Seules les adresses de la
 * colonne D de la ligne « Annulations » de l'onglet Accès (et le propriétaire) sont écoutées. Chaque mail est lu une fois.
 */
var RECHERCHE_ANNULATIONS = 'subject:(annulé OR annulée OR annule OR annulation OR annuler)';
function lireAnnulations_(fenetre) {
  var a = acces_(), permis = a.mailsAnnul.slice(), props = PropertiesService.getScriptProperties(), n = 0;
  try { permis.push(String(Session.getEffectiveUser().getEmail() || '').toLowerCase()); } catch (err) { /* sans */ }
  var vus = JSON.parse(props.getProperty('ANNUL_MAILS') || '[]'), deja = {}, f = null, presents = null;
  vus.forEach(function (id) { deja[id] = 1; });
  messages_(RECHERCHE_ANNULATIONS + fenetre, 50).forEach(function (m) {
    var id = m.getId(), objet = m.getSubject() || '';
    if (deja[id] || !/annul/i.test(objet)) return;
    deja[id] = 1; vus.push(id);
    var de = String((/<([^>]+)>/.exec(m.getFrom()) || [null, m.getFrom()])[1] || '').trim().toLowerCase();
    if (permis.indexOf(de) < 0) { console.warn('Annulation par mail ignorée : ' + de + " n'est pas dans l'onglet Accès"); return; }
    var nums = numerosTickets_(objet);
    if (!nums.length) nums = numerosTickets_(m.getPlainBody() || '');
    if (!nums.length) return;
    f = f || ongletAnnules_();
    presents = presents || f.getDataRange().getDisplayValues().slice(1).map(function (r) { return String(r[0] || '').replace(/\s+/g, '').toUpperCase(); });
    nums.forEach(function (num) {
      if (presents.indexOf(num) >= 0) return;
      f.appendRow([num, Utilities.formatDate(m.getDate(), 'Europe/Paris', 'dd/MM/yyyy HH:mm'), 'par mail de ' + de].map(texte_));
      presents.push(num); n++;
    });
  });
  props.setProperty('ANNUL_MAILS', JSON.stringify(vus.slice(-300)));
  return n;
}

/**
 * Corrélations par mail : « CORRÉLATION » (ou « corrélé », « regrouper ») dans l'objet. Le premier n° est le ticket principal,
 * les suivants sont regroupés dessous : ceux de l'objet, puis ceux écrits dans le mail (pas ceux d'un mail transféré en
 * dessous). Mêmes adresses autorisées que pour les annulations ; chaque mail est lu une fois.
 */
var RECHERCHE_CORRELATIONS = 'subject:(corrélation OR correlation OR corrélé OR correle OR corrélés OR regrouper OR regroupement)';
function lireCorrelations_(fenetre) {
  var a = acces_(), permis = a.mailsAnnul.slice(), props = PropertiesService.getScriptProperties(), n = 0;
  try { permis.push(String(Session.getEffectiveUser().getEmail() || '').toLowerCase()); } catch (err) { /* sans */ }
  var vus = JSON.parse(props.getProperty('CORREL_MAILS') || '[]'), deja = {};
  vus.forEach(function (id) { deja[id] = 1; });
  messages_(RECHERCHE_CORRELATIONS + fenetre, 50).forEach(function (m) {
    var id = m.getId(), objet = m.getSubject() || '';
    if (deja[id] || !/corr[ée]l|regroup/i.test(objet)) return;
    deja[id] = 1; vus.push(id);
    var de = String((/<([^>]+)>/.exec(m.getFrom()) || [null, m.getFrom()])[1] || '').trim().toLowerCase();
    if (permis.indexOf(de) < 0) { console.warn('Corrélation par mail ignorée : ' + de + " n'est pas dans l'onglet Accès"); return; }
    // le texte tapé, sans le mail transféré ou cité en dessous
    var corps = String(m.getPlainBody() || '').split(/\n\s*(?:De|From|Envoyé|Sent)\s*:|\n-{2,}\s*(?:Message d'origine|Original Message|Forwarded message)|\n_{8,}|\n>/i)[0];
    var nums = numerosTickets_(objet);
    numerosTickets_(corps).forEach(function (x) { if (nums.indexOf(x) < 0) nums.push(x); });
    if (nums.length < 2) { console.warn('Corrélation par mail ignorée : il faut le ticket principal puis au moins un autre n°'); return; }
    n += ajouterCorrelations_(nums[0], nums.slice(1), 'par mail de ' + de, m.getDate());
  });
  props.setProperty('CORREL_MAILS', JSON.stringify(vus.slice(-300)));
  return n;
}

/** Ajoute à l'onglet Corrélés « nums regroupés sous principal » (sans doublon). Renvoie le nombre de lignes ajoutées. */
function ajouterCorrelations_(principal, nums, pourquoi, quand) {
  var f = ongletCorreles_(), norm = function (x) { return String(x || '').replace(/\s+/g, '').toUpperCase(); }, p = norm(principal), n = 0;
  var presents = f.getDataRange().getDisplayValues().slice(1).map(function (r) { return norm(r[0]) + '|' + norm(r[1]); });
  nums.map(norm).forEach(function (num) {
    if (!/^[A-Z0-9][A-Z0-9#._-]{3,39}$/.test(num) || num === p || num.split('#')[0] === p.split('#')[0] || presents.indexOf(num + '|' + p) >= 0) return;
    f.appendRow([num, p, Utilities.formatDate(quand || new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm'), pourquoi || ''].map(texte_));
    presents.push(num + '|' + p); n++;
  });
  if (n) PropertiesService.getScriptProperties().deleteProperty('CONTROLES_LE');
  return n;
}

/** L'onglet « Corrélés » (créé avec son mode d'emploi s'il n'existe pas). */
function ongletCorreles_() {
  var neuf = !SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_CORRELES), f = feuille_(ONGLET_CORRELES, ENTETES_CORRELES);
  if (neuf) {
    f.getRange(1, 5).setValue('À savoir').setFontWeight('bold');
    f.getRange(2, 5).setValue('Un ticket regroupé par ligne : son n° en colonne A, le ticket principal en colonne B. Les tickets regroupés disparaissent de la ' +
      'carte et des listes ; le principal prend l\'échéance la plus proche du groupe. Pour dégrouper, efface la ligne. La personne qui gère les corrélations ' +
      'peut aussi le faire par mail (« CORRÉLATION » + le n° principal puis les autres) ou sur la carte avec son code.').setWrap(true);
    f.setColumnWidth(1, 200); f.setColumnWidth(2, 200); f.setColumnWidth(3, 130); f.setColumnWidth(4, 240); f.setColumnWidth(5, 480);
  }
  return f;
}

/** Depuis la carte, avec le code « Annulations et corrélations » : regroupe des tickets sous un ticket principal. */
function correlerDepuisCarte(code, principal, nums, raison) {
  if (!droits_(code).annul) throw new Error('Réservé à la personne qui gère les annulations et les corrélations.');
  var liste = (Array.isArray(nums) ? nums : String(nums || '').split(/[\s,;]+/)).filter(function (x) { return x; });
  if (!liste.length) throw new Error('Coche ou écris au moins un ticket à regrouper.');
  var r = String(raison || '').replace(/\s+/g, ' ').trim().slice(0, 200);
  return ajouterCorrelations_(principal, liste, (r ? r + ' · ' : '') + 'depuis la carte');
}

/** Depuis la carte : sort un ticket de son groupe (efface sa ligne de l'onglet Corrélés). */
function degrouperDepuisCarte(code, num) {
  if (!droits_(code).annul) throw new Error('Réservé à la personne qui gère les annulations et les corrélations.');
  num = String(num || '').replace(/\s+/g, '').toUpperCase();
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_CORRELES), n = 0;
  if (!f || !num) return 0;
  var v = f.getDataRange().getDisplayValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0] || '').replace(/\s+/g, '').toUpperCase() === num) { f.getRange(i + 1, 1, 1, 4).clearContent(); n++; }
  }
  if (n) PropertiesService.getScriptProperties().deleteProperty('CONTROLES_LE');
  return n;
}

/** Menu « Regrouper des tickets (corrélation)… » : le n° du ticket principal, puis les n° à regrouper dessous. */
function correlerTickets() {
  proprietaire_();
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Regrouper des tickets (corrélation)', 'Le n° du ticket principal, puis les n° des tickets à regrouper dessous, séparés par des espaces.\n' +
    'Par exemple : BB-20261111T0296#0 BB-20261111T0297#0 BB-20261111T0301#0', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var nums = String(r.getResponseText() || '').toUpperCase().split(/[\s,;]+/).filter(function (x) { return x; });
  if (nums.length < 2) { ui.alert('Il faut le ticket principal puis au moins un autre n°.'); return; }
  var n = ajouterCorrelations_(nums[0], nums.slice(1), 'depuis le Sheet');
  majControles_();
  ui.alert(n + ' ticket' + (n > 1 ? 's' : '') + ' regroupé' + (n > 1 ? 's' : '') + ' sous ' + nums[0] + '. Ils disparaissent de la carte et des listes.\n\n' +
    'Pour dégrouper, efface leur ligne dans l\'onglet « ' + ONGLET_CORRELES + ' ».');
}

/** Les n° de tickets d'un texte : BB-20261111T0296#0, QRB-…, ou « N° 284320 ». */
function numerosTickets_(t) {
  var out = [], m, re = /\b[A-Z]{2,4}-\d{8}T\d{4}(?:#\d+)?/gi, re2 = /N°\s*(\d{5,})/gi;
  while ((m = re.exec(t))) { var k = m[0].toUpperCase(); if (out.indexOf(k) < 0) out.push(k); }
  while ((m = re2.exec(t))) if (out.indexOf(m[1]) < 0) out.push(m[1]);
  return out;
}

/** Mails « [CREATION INTERVENTION] » : une ligne par ticket créé, et son adresse placée sur la carte. */
function lireCreations_(fenetre) {
  var f = feuille_(ONGLET_CREATIONS, ENTETES_CREATIONS);
  var vus = {};
  if (f.getLastRow() > 1) f.getRange(2, 10, f.getLastRow() - 1, 1).getValues().forEach(function (r) { vus[r[0]] = 1; });
  var ajout = [];
  messages_(RECHERCHE_CREATIONS + fenetre, 100).forEach(function (m) {
    if (vus[m.getId()]) return;
    if (/annul|corr[ée]l/i.test(m.getSubject())) return;   // mail de création transféré pour annuler ou regrouper : lu ailleurs
    var l = ligneCreation_(m.getSubject(), m.getPlainBody(), m.getDate(), m.getId());
    if (l) { ajout.push(l.map(texte_)); vus[m.getId()] = 1; }
  });
  if (!ajout.length) return 0;
  grille_(f, f.getLastRow() + ajout.length, ENTETES_CREATIONS.length).getRange(f.getLastRow() + 1, 1, ajout.length, ENTETES_CREATIONS.length).setNumberFormat('@').setValues(ajout);
  geocoder_(ajout.map(function (l) { return [l[4], l[5], l[6]].filter(function (x) { return x; }).join(' '); }));
  return ajout.length;
}

/**
 * Une seule fois après une mise à jour de la lecture des mails de création : relit ceux des 7 derniers jours (les tickets
 * pas encore dans l'export), pour retrouver leur niveau et leur GTR. Au plus 60 mails.
 */
var LECTURE_CREATIONS = '2';
function relireCreations_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('LECTURE_CREATIONS') === LECTURE_CREATIONS) return 0;
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_CREATIONS), n = 0, adresses = [];
  if (f && f.getLastRow() > 1) {
    var v = f.getRange(2, 1, f.getLastRow() - 1, ENTETES_CREATIONS.length).getDisplayValues(), limite = Date.now() - 7 * 86400000;
    for (var i = v.length - 1; i >= 0 && n < 60; i--) {
      var m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(v[i][0]);
      if (!m || new Date(+m[3], m[2] - 1, +m[1]).getTime() < limite || !v[i][9]) continue;
      try {
        var msg = GmailApp.getMessageById(v[i][9]), l = msg && ligneCreation_(msg.getSubject(), msg.getPlainBody(), msg.getDate(), v[i][9]);
        if (l) { f.getRange(i + 2, 1, 1, l.length).setNumberFormat('@').setValues([l.map(texte_)]); adresses.push([l[4], l[5], l[6]].filter(function (x) { return x; }).join(' ')); n++; }
      } catch (err) { console.warn('Relecture de la création ' + v[i][1] + ' : ' + err); }
    }
  }
  if (adresses.length) geocoder_(adresses);
  props.setProperty('LECTURE_CREATIONS', LECTURE_CREATIONS);
  return n;
}

/**
 * Importe le plus récent export Praxedo reçu en pièce jointe, s'il est plus récent que le dernier import.
 * Chaque mail n'est regardé qu'une fois ; un CSV qui n'est pas un export d'interventions est ignoré.
 */
function importerDepuisGmail_(fenetre) {
  var props = PropertiesService.getScriptProperties();
  var vus = JSON.parse(props.getProperty('EXPORTS_VUS') || '[]');
  var dernier = props.getProperty('IMPORTE_LE') ? new Date(props.getProperty('IMPORTE_LE')) : new Date(0);
  var candidats = [];
  messages_(RECHERCHE_EXPORTS + fenetre, 20).forEach(function (m) {
    if (vus.indexOf(m.getId()) >= 0) return;
    vus.push(m.getId());
    if (m.getDate() <= dernier) return;
    m.getAttachments().forEach(function (pj) {
      if (/\.csv$/i.test(pj.getName())) candidats.push({ date: m.getDate(), pj: pj });
    });
  });
  props.setProperty('EXPORTS_VUS', JSON.stringify(vus.slice(-200)));
  candidats.sort(function (x, y) { return y.date - x.date; });
  for (var i = 0; i < candidats.length; i++) {
    var texte = candidats[i].pj.getDataAsString('UTF-8');
    if (!/(^|;|")N°("|;)/.test(texte.slice(0, 2000))) continue;  // autre CSV : ignoré
    var r = importerExport_(texte);
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

/**
 * Tout ce dont la carte a besoin, si le code est bon (onglet Accès) : rôle « ca » (conduite d'activité) ou « tech ».
 * Relit Gmail si la dernière lecture date de plus de 2 minutes.
 */
function donnees(code) {
  var dr = droits_(code), role = dr.role;
  var props = PropertiesService.getScriptProperties();
  var lu = props.getProperty('LU_LE');
  if (!lu || Date.now() - new Date(lu).getTime() > 2 * 60 * 1000) {
    try { lireMails_(120); } catch (err) { console.warn(err); }  // plusieurs personnes en même temps : une seule lecture
  }
  var d = lireDonnees_();
  d.role = role;
  d.peutAnnuler = dr.annul;
  return d;
}

/**
 * Ce que permet un code d'accès : role 'ca' (conduite d'activité) ou 'tech' ; annul = peut retirer un ticket annulé
 * (code « Annulations », qui voit aussi tout comme la conduite d'activité). Un mauvais code attend 1 s puis est refusé.
 */
function droits_(code) {
  var a = acces_(), c = String(code || '').trim();
  if (c && a.annul && c === a.annul) return { role: 'ca', annul: true };
  var role = c && c === a.ca ? 'ca' : c && c === a.tech ? 'tech' : null;
  if (!role) { Utilities.sleep(1000); throw new Error('Code incorrect'); }
  return { role: role, annul: false };
}

/**
 * Depuis la carte, avec le code « Annulations » seulement : retire un ticket annulé dans Praxedo (une ligne de plus dans
 * l'onglet Annulés, avec la date et « depuis la carte »). Les onglets du Sheet sont refaits au passage suivant (5 min).
 */
function annulerDepuisCarte(code, num, raison) {
  if (!droits_(code).annul) throw new Error('Réservé à la personne qui gère les annulations et les corrélations.');
  num = String(num || '').replace(/\s+/g, '').toUpperCase();
  if (!/^[A-Z0-9][A-Z0-9#._-]{3,39}$/.test(num)) throw new Error('N° de ticket invalide.');
  var r = String(raison || '').replace(/\s+/g, ' ').trim().slice(0, 200);
  ongletAnnules_().appendRow([num, Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm'), (r ? r + ' · ' : '') + 'depuis la carte'].map(texte_));
  PropertiesService.getScriptProperties().deleteProperty('CONTROLES_LE');
  return num;
}

/**
 * Depuis la carte, avec le code « Annulations et corrélations » : OK ou NO OK après vérification des photos d'une intervention
 * (colonne « Photos vérifiées » de sa fin d'intervention la plus récente), avec un commentaire. avis '' : efface. → le texte écrit.
 */
function verifierDepuisCarte(code, num, avis, commentaire) {
  if (!droits_(code).annul) throw new Error('Réservé à la personne qui gère les annulations et les corrélations.');
  num = String(num || '').replace(/\s+/g, '').toUpperCase();
  avis = String(avis || '').toUpperCase();
  if (['OK', 'NO OK', ''].indexOf(avis) < 0) throw new Error('Avis inconnu.');
  var verrou = LockService.getScriptLock();   // pas pendant une relecture des rapports, qui réécrit les lignes
  if (!verrou.tryLock(15000)) throw new Error('Le Sheet est occupé : réessaie dans une minute.');
  try {
    var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_FINS), n = f ? f.getLastRow() - 1 : 0, i = n - 1;
    var nums = n > 0 ? f.getRange(2, 2, n, 1).getDisplayValues() : [];
    while (i >= 0 && String(nums[i][0]).replace(/\s+/g, '').toUpperCase() !== num) i--;   // la fin la plus récente de ce ticket
    if (i < 0) throw new Error('Intervention ' + num + ' introuvable dans l\'onglet Fins d\'intervention.');
    var c = String(commentaire || '').replace(/\s+/g, ' ').trim().slice(0, 200);
    var t = avis ? avis + (c ? ' — ' + c : '') + ' (carte, ' + Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm') + ')' : '';
    grille_(f, 1, ENTETES_FINS.length).getRange(i + 2, ENTETES_FINS.indexOf(COL_PHOTOS_AVIS) + 1).setNumberFormat('@').setValue(texte_(t));
    PropertiesService.getScriptProperties().deleteProperty('CONTROLES_LE');
    return t;
  } finally { verrou.releaseLock(); }
}

/** Depuis la carte, avec le code « Annulations et corrélations » : une note sur un ticket (onglet Notes), comme si elle était tapée dans le Sheet. */
function noterDepuisCarte(code, num, note) {
  if (!droits_(code).annul) throw new Error('Réservé à la personne qui gère les annulations et les corrélations.');
  num = String(num || '').replace(/\s+/g, '').toUpperCase();
  if (!/^[A-Z0-9][A-Z0-9#._-]{3,39}$/.test(num)) throw new Error('N° de ticket invalide.');
  var t = String(note || '').replace(/\s+/g, ' ').trim().slice(0, 300);
  if (!t) throw new Error('La note est vide.');
  var f = feuille_(ONGLET_NOTES, ENTETES_NOTES);
  if (!f.getRange(1, 3).getDisplayValue()) f.getRange(1, 3).setValue('Ajoutée').setFontWeight('bold');
  f.appendRow([num, t, 'carte, ' + Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm')].map(texte_));
  PropertiesService.getScriptProperties().deleteProperty('CONTROLES_LE');
  return t;
}

/** Depuis la carte, avec le code « Annulations » : remet un ticket retiré par erreur (efface ses lignes de l'onglet Annulés). */
function retablirDepuisCarte(code, num) {
  if (!droits_(code).annul) throw new Error('Réservé à la personne qui gère les annulations et les corrélations.');
  num = String(num || '').replace(/\s+/g, '').toUpperCase();
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_ANNULES), n = 0;
  if (!f || !num) return 0;
  var v = f.getDataRange().getDisplayValues();
  for (var i = v.length - 1; i >= 1; i--) {
    if (String(v[i][0] || '').replace(/\s+/g, '').toUpperCase() === num) { f.getRange(i + 1, 1, 1, 3).clearContent(); n++; }  // la note de la colonne D reste
  }
  if (n) PropertiesService.getScriptProperties().deleteProperty('CONTROLES_LE');
  return n;
}

function lireDonnees_() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var t = ss.getSheetByName(ONGLET_TICKETS).getDataRange().getDisplayValues();
  // les clôtures et les créations des JOURS_HISTORIQUE derniers jours (les onglets, eux, gardent tout)
  var limite = Date.now() - JOURS_HISTORIQUE * 86400000;
  var recent = function (r) { var m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(r[0]); return m && new Date(+m[3], m[2] - 1, +m[1]).getTime() >= limite; };
  var fins = ss.getSheetByName(ONGLET_FINS).getDataRange().getDisplayValues().slice(1).filter(recent).map(function (r) {
    return { recu: r[0], num: r[1], commune: r[2], jour: r[3], rapport: r[4], lu: r[7] || '', debut: r[8] || '', fin: r[9] || '',
             duree: r[10] || '', corrige: r[11] || '', definitive: r[12] || '', dommage: r[13] || '', resume: r[14] || '',
             boites: r[15] || '', pm: r[16] || '', nature: r[17] || '', cause: r[18] || '', adresse: r[19] || '', demande: r[20] || '', etiquettes: r[21] || '', autres: r[22] || '', tech: r[23] || '',
             finImpact: r[24] || (/fin d'impact client : (\d{2}\/\d{2}\/\d{4} \d{1,2}:\d{2})/.exec(r[22] || '') || [])[1] || '',
             photos: r[25] || '', photosLe: r[26] || '', photosAvis: r[27] || '' };
  });
  var fn = feuille_(ONGLET_NOTES, ENTETES_NOTES);
  var notes = fn.getDataRange().getDisplayValues().slice(1).map(function (r) { return { num: r[0], note: r[1] }; })
    .filter(function (n) { return n.num && n.note; });
  var fc = ss.getSheetByName(ONGLET_CREATIONS);
  var creations = fc ? fc.getDataRange().getDisplayValues().slice(1).filter(recent).map(function (r) {
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
  // tickets annulés dans Praxedo, notés à la main (il n'y a pas de mail d'annulation)
  var fa = ss.getSheetByName(ONGLET_ANNULES);
  var annules = fa ? fa.getDataRange().getDisplayValues().slice(1).map(function (r) { return { num: String(r[0] || '').trim(), le: r[1] || '', pourquoi: r[2] || '' }; })
    .filter(function (a) { return a.num; }) : [];
  // tickets regroupés (corrélés) sous un ticket principal
  var fco = ss.getSheetByName(ONGLET_CORRELES);
  var correles = fco ? fco.getDataRange().getDisplayValues().slice(1).map(function (r) { return { num: String(r[0] || '').trim(), principal: String(r[1] || '').trim(), le: r[2] || '', pourquoi: r[3] || '' }; })
    .filter(function (c) { return c.num && c.principal; }) : [];
  return { entetes: t[0], lignes: t.slice(1), fins: fins, creations: creations, sortis: sortis, notes: notes, adresses: adresses, annules: annules, correles: correles,
           importeLe: props.getProperty('IMPORTE_LE') || '', luLe: props.getProperty('LU_LE') || '' };
}

/** Menu « Mettre à jour les onglets de contrôle ». Réservé au propriétaire. */
function mettreAJourControles() {
  proprietaire_();
  return majControles_();
}

/**
 * Refait les onglets Incohérences, Tickets liés et PM chargés, avec la même analyse que la carte
 * (fichier Analyse : mêmes règles que l'Excel d'analyse_tickets.py).
 */
function majControles_() {
  var a = analyse_(), NOW = new Date(), res = a.analyser(lireDonnees_(), NOW), c = a.controles(res, NOW);
  ongletsControle_(c).forEach(function (o) { ecrireOnglet_(o.nom, o.entetes, o.lignes, o.couleur, o.largeurs); });
  try { graphiquesSheet_(a.listesDuJour(res, c, NOW, {}).graphiques); } catch (err) { console.warn('Graphiques : ' + err); }
  try { adressesEnvoi_(); } catch (err) { console.warn('Onglet ' + ONGLET_ENVOI + ' : ' + err); }
  try { ongletAnnules_(); } catch (err) { console.warn('Onglet ' + ONGLET_ANNULES + ' : ' + err); }
  try { ongletCorreles_(); } catch (err) { console.warn('Onglet ' + ONGLET_CORRELES + ' : ' + err); }
  PropertiesService.getScriptProperties().setProperty('CONTROLES_LE', maintenant_());
  return { incoherences: c.incoherences.length, lies: c.lies.length, pm: c.pm.length, avertissements: res.avertissements };
}

/**
 * Onglet « Graphiques » du Sheet : les trois graphiques de l'Excel du jour (clôtures par semaine, communes et boîtes où on
 * intervient le plus, sur 90 jours), refaits à chaque mise à jour des onglets. Leurs chiffres sont à droite, colonne R.
 */
var ONGLET_GRAPHIQUES = 'Graphiques';
function graphiquesSheet_(g) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), f = ss.getSheetByName(ONGLET_GRAPHIQUES) || ss.insertSheet(ONGLET_GRAPHIQUES);
  f.getCharts().forEach(function (c) { f.removeChart(c); });
  f.clear();
  var C0 = 18, n = g.entetes.length, tout = [g.entetes].concat(g.lignes);
  // noms et dates en texte (« 9 nov. » ne doit pas devenir une date), nombres en nombres
  f.getRange(1, C0, tout.length, n).setNumberFormat('@');
  g.entetes.forEach(function (h, j) { if (g.lignes.length && /^(Sans suite|À reprendre|Clôtures)/.test(h)) f.getRange(2, C0 + j, g.lignes.length, 1).setNumberFormat('0'); });
  f.getRange(1, C0, tout.length, n).setValues(tout);
  f.getRange(1, C0, 1, n).setFontWeight('bold');
  var texte = { color: '#52514e', fontSize: 10 }, axe = { textStyle: texte, gridlines: { color: '#e6e6e6' }, baselineColor: '#bfbfbf' };
  [['col', 1, 3, g.nSemaines, 1, 'Clôtures par semaine (la dernière est en cours)', 330],
   ['bar', 7, 8, g.nCommunes, 19, 'Les communes où on intervient le plus', 380],
   ['bar', 12, 13, g.nBoites, 39, 'Les boîtes où on intervient le plus', 380]].forEach(function (x) {
    if (!x[3]) return;
    var b = x[0] === 'col' ? f.newChart().asColumnChart() : f.newChart().asBarChart();
    f.insertChart(b.addRange(f.getRange(1, C0 + x[1], x[3] + 1, 1)).addRange(f.getRange(1, C0 + x[2], x[3] + 1, 2)).setNumHeaders(1)
      .setPosition(x[4], 1, 0, 0)
      .setOption('title', x[5]).setOption('titleTextStyle', { color: '#0b0b0b', fontSize: 13, bold: true })
      .setOption('isStacked', true).setOption('colors', ['#2a78d6', '#eb6834'])
      .setOption('legend', { position: 'top', textStyle: { color: '#0b0b0b', fontSize: 10 } })
      .setOption('bar', { groupWidth: x[0] === 'col' ? '45%' : '60%' })
      .setOption('hAxis', axe).setOption('vAxis', axe)
      .setOption('width', 760).setOption('height', x[6]).build());
  });
}

/**
 * Envoi du matin : chaque jour ouvré à partir de 7 h (au passage suivant de la lecture des mails, avant midi), l'Excel du
 * jour puis la carte du jour, en deux mails (si la messagerie bloque la carte, l'Excel arrive quand même), aux adresses de
 * l'onglet « Envoi du matin ». Rien n'est envoyé tant que l'onglet est vide. Une seule fois par jour.
 */
var ONGLET_ENVOI = 'Envoi du matin', HEURE_ENVOI = 7;
function envoiDuMatin_() {
  var props = PropertiesService.getScriptProperties(), d = new Date();
  var jour = Utilities.formatDate(d, 'Europe/Paris', 'yyyy-MM-dd'), heure = +Utilities.formatDate(d, 'Europe/Paris', 'H');
  if (props.getProperty('ENVOI_MATIN_LE') === jour || heure < HEURE_ENVOI || heure >= 12) return 0;
  var p = jour.split('-');
  if (!analyse_().ouvre(new Date(+p[0], p[1] - 1, +p[2], 12), '54')) return 0;   // week-end ou jour férié
  var dest = adressesEnvoi_();
  if (!dest.length) return 0;
  var verrou = LockService.getScriptLock();
  if (!verrou.tryLock(5000)) return 0;
  try {
    if (props.getProperty('ENVOI_MATIN_LE') === jour) return 0;
    props.setProperty('ENVOI_MATIN_LE', jour);   // une fois par jour, même si un des deux mails échoue
    return envoyerFichiers_(dest);
  } finally { verrou.releaseLock(); }
}

/** Menu « Envoi du matin : envoyer maintenant » : le même envoi, tout de suite (pour essayer). Réservé au propriétaire. */
function envoyerMaintenant() {
  proprietaire_();
  var dest = adressesEnvoi_(), ui = SpreadsheetApp.getUi();
  if (!dest.length) { ui.alert('Ajoute d\'abord les adresses dans l\'onglet « ' + ONGLET_ENVOI + ' » (une par ligne, colonne A).'); return; }
  envoyerFichiers_(dest);
  ui.alert('L\'Excel et la carte du jour sont partis à ' + dest.length + ' adresse' + (dest.length > 1 ? 's' : '') + ' (deux mails).');
}

/** Les adresses de l'onglet « Envoi du matin » (colonne A), sans doublon. L'onglet est créé s'il n'existe pas. */
function adressesEnvoi_() {
  return lignesEnvoi_().map(function (l) { return l.adresse; });
}

/**
 * Les lignes de l'onglet « Envoi du matin » : [{ adresse, rapports }]. rapports : « oui » en colonne B, cette adresse reçoit
 * aussi les rapports du soir et les bilans (sinon, seulement l'Excel et la carte du matin).
 */
function lignesEnvoi_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), f = ss.getSheetByName(ONGLET_ENVOI), aSavoir = 'Une adresse par ligne, en colonne A. Chaque jour ouvré vers 7 h, ' +
    'l\'Excel du jour puis la carte du jour partent à ces adresses, en deux mails. Écris « oui » en colonne B pour qu\'une adresse reçoive aussi le ' +
    'rapport du soir (vers 17 h) et les bilans de la semaine et du mois. Tant que la colonne A est vide, rien n\'est envoyé.';
  if (!f) {
    f = feuille_(ONGLET_ENVOI, ['Adresse mail', 'Rapports du soir et bilans aussi (oui / non)', 'Remarque', 'À savoir']);
    f.getRange(2, 4).setValue(aSavoir).setWrap(true);
    f.setColumnWidth(1, 260); f.setColumnWidth(2, 170); f.setColumnWidth(3, 200); f.setColumnWidth(4, 520);
  } else if (f.getRange(1, 2).getDisplayValue() === 'Remarque') {   // onglet d'une version d'avant : la colonne B est ajoutée
    f.insertColumnBefore(2);
    f.getRange(1, 2).setValue('Rapports du soir et bilans aussi (oui / non)').setFontWeight('bold');
    f.getRange(2, 4).setValue(aSavoir).setWrap(true);
    f.setColumnWidth(2, 170);
  }
  var vus = {}, out = [];
  if (f.getLastRow() > 1) f.getRange(2, 1, f.getLastRow() - 1, 2).getDisplayValues().forEach(function (r) {
    var a = String(r[0] || '').trim().toLowerCase();
    if (/^[^\s@,;]+@[^\s@,;]+\.[a-z]{2,}$/.test(a) && !vus[a]) { vus[a] = 1; out.push({ adresse: a, rapports: /^\s*oui\b/i.test(r[1]) }); }
  });
  return out;
}

/** Les destinataires des rapports du soir et des bilans : le propriétaire, puis les adresses avec « oui » en colonne B de l'onglet Envoi du matin. */
function destRapports_() {
  var moi = Session.getEffectiveUser().getEmail(), out = [moi];
  try { lignesEnvoi_().forEach(function (l) { if (l.rapports && l.adresse !== String(moi).toLowerCase()) out.push(l.adresse); }); }
  catch (err) { console.warn('Envoi du matin : ' + err); }
  return out;
}

/** Deux mails : l'Excel du jour avec le résumé dans le texte, puis la carte du jour. */
function envoyerFichiers_(dest) {
  var quand = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy'), x = excelDuJour_();
  var resume = x.resume.map(function (l) { return l[1] === '' ? '\n' + l[0] : l[0] + ' : ' + l[1]; }).join('\n').trim();
  GmailApp.sendEmail(dest.join(','), 'Tickets du ' + quand + ' : l\'Excel du jour',
    'Bonjour,\n\nEn pièce jointe : l\'Excel du jour (à planifier, justifiés, suites, incohérences, clôtures, statistiques et graphiques).\n' +
    'La carte du jour arrive dans un second mail.\n\n' + resume + '\n', { attachments: [x.blob], name: 'Carte des tickets' });
  GmailApp.sendEmail(dest.join(','), 'Tickets du ' + quand + ' : la carte du jour',
    'Bonjour,\n\nEn pièce jointe : la carte du jour. Télécharge le fichier puis ouvre-le dans un navigateur (Edge, Chrome…). ' +
    'C\'est une photo de ce matin : elle ne se met pas à jour.\n', { attachments: [Utilities.newBlob(construireCarteDuJour_(), 'text/html', nomCarteDuJour_())], name: 'Carte des tickets' });
  return dest.length;
}

/**
 * Tickets annulés dans Praxedo : il n'y a pas de mail d'annulation, donc on les note dans l'onglet « Annulés » (à la main,
 * ou par le menu « Annuler un ticket… »). Ils disparaissent de la carte, des listes, de l'Excel et des mails du matin.
 * Pour rétablir un ticket : effacer sa ligne.
 */
function ongletAnnules_() {
  var neuf = !SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_ANNULES), f = feuille_(ONGLET_ANNULES, ENTETES_ANNULES);
  if (neuf) {
    f.getRange(1, 4).setValue('À savoir').setFontWeight('bold');
    f.getRange(2, 4).setValue('Un n° de ticket par ligne en colonne A (la date et la raison sont facultatives). Avec « #0 », « #1 »… seule cette fois-là ' +
      'est annulée ; sans « # », toutes les fois du ticket. Le ticket disparaît de la carte, des listes et de l\'Excel. Pour le rétablir, efface sa ligne. ' +
      'La personne qui gère les annulations peut aussi le faire par mail (« ANNULÉ » + n° du ticket) ou sur la carte avec son code (onglet Accès).').setWrap(true);
    f.setColumnWidth(1, 200); f.setColumnWidth(2, 130); f.setColumnWidth(3, 260); f.setColumnWidth(4, 480);
  }
  return f;
}

/** Menu « Annuler un ticket… » : demande le n° (et la raison), l'ajoute à l'onglet Annulés, puis refait les onglets. */
function annulerTicket() {
  proprietaire_();
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('Annuler un ticket', 'N° du ticket, par exemple BB-20261111T0296#0. Tu peux ajouter la raison après un espace.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var m = /^\s*(\S+)\s*([\s\S]*)$/.exec(r.getResponseText() || '');
  if (!m) return;
  var num = m[1].toUpperCase(), f = ongletAnnules_();
  grille_(f, f.getLastRow() + 1, 3).getRange(f.getLastRow() + 1, 1, 1, 3).setNumberFormat('@')
    .setValues([[num, Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm'), m[2].trim()].map(texte_)]);
  // le n° est-il connu (dernier export ou mails de création) ?
  var ss = SpreadsheetApp.getActiveSpreadsheet(), connus = {};
  [[ONGLET_TICKETS, 'N°'], [ONGLET_CREATIONS, 'N° ticket']].forEach(function (o) {
    var fe = ss.getSheetByName(o[0]), v = fe ? fe.getDataRange().getDisplayValues() : [], c = v.length ? v[0].indexOf(o[1]) : -1;
    if (c >= 0) v.slice(1).forEach(function (l) { var n = String(l[c] || '').trim().toUpperCase(); connus[n] = 1; connus[n.split('#')[0]] = 1; });
  });
  majControles_();
  ui.alert('Ticket ' + num + ' annulé : il n\'apparaît plus sur la carte, dans les listes ni dans l\'Excel.' +
    (connus[num] ? '' : '\n\nAttention : ce n° n\'est ni dans le dernier export ni dans les mails de création. Vérifie-le dans l\'onglet « ' + ONGLET_ANNULES + ' ».') +
    '\n\nPour le rétablir, efface sa ligne dans l\'onglet « ' + ONGLET_ANNULES + ' ».');
}

/**
 * Rapport du soir : chaque jour vers 17 h, à ton adresse, les fins d'intervention reçues depuis le rapport d'avant, avec ce
 * qui demande de l'attention puis le détail. Le dimanche, en plus, le bilan de la semaine ; le 1er, celui du mois d'avant.
 * Pas de mail un soir sans rien.
 */
var HEURE_RAPPORT = 17;
function rapportDuSoir_() {
  var props = PropertiesService.getScriptProperties(), d = new Date();
  var jour = Utilities.formatDate(d, 'Europe/Paris', 'yyyy-MM-dd'), heure = +Utilities.formatDate(d, 'Europe/Paris', 'H');
  if (props.getProperty('RAPPORT_SOIR_LE') === jour || heure < HEURE_RAPPORT || heure >= 23) return 0;
  var verrou = LockService.getScriptLock();
  if (!verrou.tryLock(5000)) return 0;
  try {
    if (props.getProperty('RAPPORT_SOIR_LE') === jour) return 0;
    props.setProperty('RAPPORT_SOIR_LE', jour);   // une fois par jour, même si l'envoi échoue
    var n = envoyerRapport_('soir', false), p = jour.split('-');
    if (new Date(+p[0], p[1] - 1, +p[2], 12).getDay() === 0) n += envoyerRapport_('semaine', false);   // dimanche : la semaine
    if (+p[2] === 1) n += envoyerRapport_('mois', false);   // le 1er : le mois d'avant
    return n;
  } finally { verrou.releaseLock(); }
}

/**
 * mode 'soir' : depuis le rapport d'avant (24 h la première fois, 4 jours au plus) ; 'semaine' : depuis lundi ; 'mois' : le mois
 * d'avant, du 1er au dernier jour. force : envoyer même vide, sans déplacer le repère du soir.
 */
function envoyerRapport_(mode, force) {
  var props = PropertiesService.getScriptProperties(), a = analyse_(), d = lireDonnees_(), NOW = new Date();
  var res = a.analyser(d, NOW), c = a.controles(res, NOW), depuis, jusqua = null, hebdo = mode !== 'soir';
  if (mode === 'semaine') { depuis = new Date(NOW.getFullYear(), NOW.getMonth(), NOW.getDate()); depuis.setDate(depuis.getDate() - ((depuis.getDay() + 6) % 7)); }
  else if (mode === 'mois') { depuis = new Date(NOW.getFullYear(), NOW.getMonth() - 1, 1); jusqua = new Date(NOW.getFullYear(), NOW.getMonth(), 1); }
  else {
    var p = props.getProperty('RAPPORT_SOIR_DEPUIS');
    depuis = p ? new Date(p) : new Date(NOW.getTime() - 86400000);
    if (isNaN(depuis) || NOW - depuis > 4 * 86400000) depuis = new Date(NOW.getTime() - 4 * 86400000);
  }
  // l'historique du backlog (une ligne par rapport du soir) : le stock d'avant, et la courbe de la semaine ou du mois
  var historique = historiqueBacklog_();
  var r = a.rapportFins(res, c, NOW, depuis, mode, d, { lienCarte: props.getProperty('LIEN_CARTE') || '', historique: historique,
    backlogAvant: historique.length ? historique[historique.length - 1].ouverts : null, photosIA: etatIA_() }, jusqua);
  if (!hebdo && !force) {
    props.setProperty('RAPPORT_SOIR_DEPUIS', NOW.toISOString());
    try {   // même un soir sans mail : la ligne du jour garde la courbe complète
      var b = r.backlog, st = b.stock;
      feuille_(ONGLET_BACKLOG, ENTETES_BACKLOG).appendRow([Utilities.formatDate(NOW, 'Europe/Paris', 'dd/MM/yyyy HH:mm'), b.crees, b.clotures, b.sortis, b.annules, b.regroupes,
        st.ouverts, st.aPlanifier, st.ko, st.bientot, st.planifies, st.planifiesKo, st.justifies, st.vieux30]);
    } catch (err) { console.warn('Backlog : ' + err); }
  }
  if (!r.aEnvoyer && !force) return 0;
  // à la main (menu, pour voir) : seulement au propriétaire ; tout seul : aussi aux adresses avec « oui » en colonne B de l'onglet Envoi du matin
  GmailApp.sendEmail((force ? [Session.getEffectiveUser().getEmail()] : destRapports_()).join(','), r.sujet, r.texte, { htmlBody: r.html, name: 'Carte des tickets' });
  return 1;
}

/** L'onglet Backlog : [{ le (Date), ouverts }], du plus ancien au plus récent. */
function historiqueBacklog_() {
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_BACKLOG);
  if (!f || f.getLastRow() < 2) return [];
  return f.getRange(2, 1, f.getLastRow() - 1, 7).getDisplayValues().map(function (r) {
    var m = /^(\d{2})\/(\d{2})\/(\d{4}) (\d{1,2}):(\d{2})/.exec(r[0]);
    return { le: m ? new Date(+m[3], m[2] - 1, +m[1], +m[4], +m[5]) : null, ouverts: +r[6] || 0 };
  }).filter(function (x) { return x.le; });
}

/** Menus « Rapports par mail » : le rapport du soir ou le bilan de la semaine, tout de suite (pour voir). Réservés au propriétaire. */
function rapportSoirMaintenant() {
  proprietaire_();
  envoyerRapport_('soir', true);
  SpreadsheetApp.getUi().alert('Rapport des fins d\'intervention envoyé à ' + Session.getEffectiveUser().getEmail() + '.\n\nIl part aussi tout seul chaque soir vers 17 h.' + autresDest_());
}
function rapportSemaineMaintenant() {
  proprietaire_();
  envoyerRapport_('semaine', true);
  SpreadsheetApp.getUi().alert('Bilan de la semaine envoyé à ' + Session.getEffectiveUser().getEmail() + '.\n\nIl part aussi tout seul le dimanche vers 17 h.' + autresDest_());
}
function rapportMoisMaintenant() {
  proprietaire_();
  envoyerRapport_('mois', true);
  SpreadsheetApp.getUi().alert('Bilan du mois dernier envoyé à ' + Session.getEffectiveUser().getEmail() + '.\n\nIl part aussi tout seul le 1er de chaque mois vers 17 h.' + autresDest_());
}

/** Pour les menus des rapports : à qui ils partent aussi quand ils sont envoyés tout seuls. */
function autresDest_() {
  var autres = destRapports_().slice(1);
  return autres.length ? ' À ' + (autres.length > 1 ? 'ces heures-là, il part aussi à : ' : 'ce moment-là, il part aussi à ') + autres.join(', ') + ' (onglet Envoi du matin, colonne B).' : '';
}

/** Les fonctions du fichier Analyse (les mêmes que la carte), pour s'en servir côté serveur. */
function analyse_() {
  var src = HtmlService.createHtmlOutputFromFile('Analyse').getContent().replace(/^\s*<script>/, '').replace(/<\/script>\s*$/, '');
  return new Function(src + '\nreturn { analyser: analyser, controles: controles, listesDuJour: listesDuJour, ouvre: ouvre, rapportFins: rapportFins };')();
}

/**
 * Les onglets de contrôle, décrits une seule fois pour le Sheet et pour l'Excel du jour :
 * nom, en-têtes, lignes, couleur de fond d'une ligne (ou rien) et largeur des colonnes en pixels.
 */
function ongletsControle_(c) {
  // couleurs : on retrouve les colonnes par leur nom, pour qu'un ajout de colonne ne décale rien
  var col = function (t, nom) { return t.entetes.indexOf(nom); };
  var cGtr = col(c.clotures, 'Dans la GTR'), cRes = col(c.clotures, 'Résultat'), bRes = col(c.statsBoites, 'Dernier résultat');
  return [
    { nom: ONGLET_INCOHERENCES, entetes: ['Gravité', 'N°', 'Statut', 'Problème', 'Détail', 'Technicien', 'Ville', 'Adresse'], lignes: c.incoherences,
      couleur: function (l) { return COULEURS_GRAVITE[l[0]]; }, largeurs: [100, 160, 100, 300, 480, 150, 160, 220] },
    { nom: ONGLET_LIES, entetes: ['Groupe', 'Pourquoi liés', 'N°', 'Statut', 'Planifiée', 'Technicien', 'Créée', 'Ville', 'Adresse', 'Immeuble',
      'Boîte(s)', 'Réf. client', 'Description'], lignes: c.lies, couleur: null, largeurs: [60, 220, 160, 100, 120, 150, 120, 160, 220, 130, 170, 130, 420] },
    { nom: ONGLET_PM, entetes: ['PM (Site)', 'Tickets ouverts', 'dont à planifier', 'Communes', 'Créés entre', 'Équipement', 'Tickets'], lignes: c.pm,
      couleur: null, largeurs: [160, 80, 80, 220, 190, 220, 600] },
    { nom: ONGLET_SUITES, entetes: ['Clôturée le', 'N°', 'Commune', 'Technicien', 'Pourquoi une suite', 'Ticket de suite', 'Compte-rendu du technicien', 'Rapport'],
      lignes: c.suites, couleur: function (l) { return l[5] ? null : '#FFE699'; }, largeurs: [120, 160, 170, 150, 220, 200, 480, 260] },
    // en gris : ce qui reste est chez l'opérateur (jarretière, brassage), pas une reprise pour nous
    { nom: ONGLET_CLOTURES, entetes: c.clotures.entetes, lignes: c.clotures.lignes,
      couleur: function (l) { return l[cGtr] === 'non' ? '#F8CBAD' : /^À suivre/.test(l[cRes]) ? '#FFE699' : /côté opérateur/.test(l[cRes]) ? '#EDEDED' : null; },
      largeurs: [120, 160, 170, 150, 120, 120, 80, 120, 120, 200, 120, 90, 200, 190, 150, 180, 480, 260, 480, 320, 360, 140] },
    { nom: ONGLET_STATS_BOITES, entetes: c.statsBoites.entetes, lignes: c.statsBoites.lignes,
      couleur: function (l) { return /^À suivre/.test(l[bRes]) ? '#FFE699' : null; }, largeurs: [190, 150, 170, 90, 100, 90, 90, 90, 90, 100, 100, 220, 240, 420] },
    { nom: ONGLET_STATS_VILLES, entetes: c.statsVilles.entetes, lignes: c.statsVilles.lignes,
      couleur: function (l, i) { return i === 0 ? '#DDEBF7' : null; }, largeurs: [220, 90, 90, 100, 100, 90, 90, 90, 90, 90, 320, 280] },
    { nom: ONGLET_STATS_SEMAINES, entetes: c.statsSemaines.entetes, lignes: c.statsSemaines.lignes, couleur: null,
      largeurs: [130, 100, 100, 90, 110, 100, 110, 110, 150, 150] },
    // peut-être causé par nos interventions, redondances, corrections à faire (voir qualite dans Analyse)
    { nom: ONGLET_QUALITE, entetes: c.qualite.entetes, lignes: c.qualite.lignes,
      couleur: function (l) { return l[1] === 'fort' ? '#F8CBAD' : l[1] === 'moyen' ? '#FFE699' : null; },
      largeurs: [170, 70, 120, 160, 160, 150, 520, 260, 300, 260] }
  ];
}

function ecrireOnglet_(nom, entetes, lignes, couleur, largeurs) {
  var f = feuille_(nom, entetes), n = entetes.length;
  var valeurs = [entetes].concat(lignes).map(function (l) { return l.map(texte_); });
  var filtre = f.getFilter();
  if (filtre) filtre.remove();
  f.clear();
  var tout = grille_(f, valeurs.length, n).getRange(1, 1, valeurs.length, n);
  tout.setNumberFormat('@').setValues(valeurs).setWrap(true).setVerticalAlignment('top');
  f.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#1F4E78').setFontColor('#FFFFFF');
  f.setFrozenRows(1);
  if (couleur && lignes.length) {
    f.getRange(2, 1, lignes.length, n).setBackgrounds(lignes.map(function (l, i) { var c = couleur(l, i) || null; return entetes.map(function () { return c; }); }));
  }
  largeurs.forEach(function (w, i) { f.setColumnWidth(i + 1, w); });
  if (lignes.length) tout.createFilter();
}

/**
 * Lit les rapports d'intervention des mails de fin qui ne l'ont pas encore été, les plus récents d'abord, et remplit
 * les colonnes « Rapport lu » à « Compte-rendu du technicien » de l'onglet Fins d'intervention.
 * « Rapport lu » : oui, illisible (PDF sans texte reconnu), autre ticket, lien refusé, ou erreur 1 à 3 (3 essais).
 */
function lireRapports_(max) {
  var f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_FINS), nc = ENTETES_FINS.length;
  grille_(f, 1, nc);
  if (f.getRange(1, 1, 1, nc).getDisplayValues()[0].join('|') !== ENTETES_FINS.join('|')) f.getRange(1, 1, 1, nc).setValues([ENTETES_FINS]).setFontWeight('bold');
  var n = f.getLastRow() - 1, essais = 0, lus = 0, aPlacer = [], techs = null;
  if (n < 1) return 0;
  // une fois après une mise à jour de la lecture : les rapports des 30 derniers jours sont relus, les plus récents d'abord
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('LECTURE_RAPPORTS') !== LECTURE_RAPPORTS) {
    var etats = f.getRange(2, 1, n, 8).getDisplayValues(), limite = Date.now() - 30 * 86400000, change = false;
    var col = etats.map(function (r) {
      var m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(r[0]);
      if (r[7] === 'oui' && m && new Date(+m[3], m[2] - 1, +m[1]).getTime() >= limite) { change = true; return ['à relire']; }
      return [r[7]];
    });
    if (change) f.getRange(2, 8, n, 1).setNumberFormat('@').setValues(col);
    props.setProperty('LECTURE_RAPPORTS', LECTURE_RAPPORTS);
  }
  // le technicien d'un ticket, tant qu'il est encore dans l'export (ou vient d'en sortir) : le rapport ne le donne pas
  function technicien(num) {
    if (!techs) {
      techs = {};
      var ss = SpreadsheetApp.getActiveSpreadsheet(), fs = ss.getSheetByName(ONGLET_SORTIS), t = ss.getSheetByName(ONGLET_TICKETS).getDataRange().getDisplayValues();
      if (fs) fs.getDataRange().getDisplayValues().slice(1).forEach(function (r) { if (r[1] && r[4]) techs[r[1]] = r[4]; });
      var iN = t[0].indexOf('N°'), iNom = t[0].indexOf('Nom technicien'), iPre = t[0].indexOf('Prénom technicien');
      if (iN >= 0 && iNom >= 0) t.slice(1).forEach(function (r) { var nom = [r[iNom], iPre >= 0 ? r[iPre] : ''].join(' ').trim(); if (r[iN] && nom) techs[r[iN]] = nom; });
    }
    return techs[num] || '';
  }
  var v = f.getRange(2, 1, n, nc).getDisplayValues();
  for (var i = n - 1; i >= 0 && essais < max; i--) {
    var url = v[i][4], etat = v[i][7], champs;
    if (!url || (etat && !/^(erreur [12]|à relire)$/.test(etat))) continue;
    if (!/^https:\/\/([a-z0-9-]+\.)*praxedo\.com\//i.test(url)) champs = { 'Rapport lu': 'lien refusé' };   // on n'ouvre que les liens de Praxedo
    else {
      essais++;
      try {
        var rep = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (rep.getResponseCode() !== 200) throw new Error('réponse ' + rep.getResponseCode());
        var octets = Uint8Array.from(rep.getContent()), c = champsRapport_(texteDuPdf_(octets));
        PDF_LUS_[url] = octets;   // pour l'analyse des photos, juste après
        var base = function (x) { return String(x).split('#')[0].toUpperCase(); };
        if (!c.num && !c.debut && !c.resume) champs = { 'Rapport lu': 'illisible' };
        else if (c.num && base(c.num) !== base(v[i][1])) champs = { 'Rapport lu': 'autre ticket' };
        else {
          champs = { 'Rapport lu': 'oui', 'Début': c.debut, 'Fin': c.fin, 'Durée': c.duree, 'Défaut corrigé': c.corrige, 'Réparation définitive': c.definitive,
            'Dommage réseau': c.dommage, 'Compte-rendu du technicien': c.resume || c.actions, 'Boîte(s)': c.boites, 'PM (site)': c.pm, 'Nature du défaut': c.nature,
            'Cause du défaut': c.cause, 'Adresse': c.adresse, 'Demande le': c.demande, 'Étiquettes GTR': c.etiquettes, 'Autres éléments du rapport': c.autres,
            "Fin d'impact": c.finImpact };
          lus++; if (c.adresse) aPlacer.push(c.adresse);
        }
      } catch (err) {
        console.warn('Rapport ' + v[i][1] + ' : ' + err);
        champs = { 'Rapport lu': 'erreur ' + ((+String(etat).replace(/\D/g, '') || 0) + 1) };
      }
    }
    champs['Technicien'] = v[i][ENTETES_FINS.indexOf('Technicien')] || technicien(v[i][1]);
    [COL_PHOTOS, COL_PHOTOS_LE, COL_PHOTOS_AVIS].forEach(function (h) { champs[h] = v[i][ENTETES_FINS.indexOf(h)] || ''; });   // une relecture ne les efface pas
    // une valeur par colonne, de « Rapport lu » à la dernière, rangée par nom : ajouter une colonne ne décale rien
    f.getRange(i + 2, 8, 1, nc - 7).setNumberFormat('@').setValues([ENTETES_FINS.slice(7).map(function (h) { return texte_(champs[h] || ''); })]);
  }
  if (aPlacer.length) try { geocoder_(aPlacer); } catch (err) { console.warn('Adresses des rapports : ' + err); }   // pour les boîtes sur la carte
  return lus;
}

// ---------- lecture d'un PDF de rapport Praxedo (texte seulement) ----------
// Les rapports sont des PDF simples : texte en flux compressés (Flate), polices en WinAnsi. Pas besoin d'un vrai lecteur PDF.

/** Décompresse un flux zlib/deflate (Uint8Array) : RFC 1950 / 1951. */
function inflate_(src, debut) {
  var pos = (debut || 0) + 2, bitbuf = 0, bitcnt = 0, out = [];   // 2 octets d'en-tête zlib
  function bits(n) {
    while (bitcnt < n) { if (pos >= src.length) throw new Error('flux tronqué'); bitbuf |= src[pos++] << bitcnt; bitcnt += 8; }
    var v = bitbuf & ((1 << n) - 1); bitbuf >>>= n; bitcnt -= n; return v;
  }
  function table(longueurs) {   // Huffman canonique : { compte par longueur, symboles triés }
    var compte = new Uint16Array(16), symboles = new Uint16Array(longueurs.length), decal = new Uint16Array(16), i;
    for (i = 0; i < longueurs.length; i++) compte[longueurs[i]]++;
    compte[0] = 0;
    for (i = 1; i < 16; i++) decal[i] = decal[i - 1] + compte[i - 1];
    for (i = 0; i < longueurs.length; i++) if (longueurs[i]) symboles[decal[longueurs[i]]++] = i;
    return { compte: compte, symboles: symboles };
  }
  function symbole(t) {
    var code = 0, premier = 0, index = 0;
    for (var len = 1; len < 16; len++) {
      code |= bits(1);
      var n = t.compte[len];
      if (code - n < premier) return t.symboles[index + (code - premier)];
      index += n; premier += n; premier <<= 1; code <<= 1;
    }
    throw new Error('code Huffman invalide');
  }
  var LBASE = [3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258],
      LEXTRA = [0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0],
      DBASE = [1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577],
      DEXTRA = [0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13],
      ORDRE = [16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];
  var fixeL = null, fixeD = null, dernier, i;
  do {
    dernier = bits(1);
    var type = bits(2);
    if (type === 0) {                       // bloc non compressé
      bitbuf = 0; bitcnt = 0;
      var n = src[pos] | (src[pos + 1] << 8); pos += 4;
      for (i = 0; i < n; i++) out.push(src[pos++]);
      continue;
    }
    var tl, td;
    if (type === 1) {                       // tables fixes
      if (!fixeL) {
        var l = new Uint8Array(288);
        for (i = 0; i < 144; i++) l[i] = 8; for (; i < 256; i++) l[i] = 9; for (; i < 280; i++) l[i] = 7; for (; i < 288; i++) l[i] = 8;
        fixeL = table(l); var d = new Uint8Array(30); d.fill(5); fixeD = table(d);
      }
      tl = fixeL; td = fixeD;
    } else if (type === 2) {                // tables dynamiques
      var nl = bits(5) + 257, nd = bits(5) + 1, nc = bits(4) + 4, lc = new Uint8Array(19);
      for (i = 0; i < nc; i++) lc[ORDRE[i]] = bits(3);
      var tc = table(lc), longueurs = new Uint8Array(nl + nd), k = 0;
      while (k < nl + nd) {
        var s = symbole(tc), rep, val;
        if (s < 16) { longueurs[k++] = s; continue; }
        if (s === 16) { val = longueurs[k - 1]; rep = 3 + bits(2); }
        else if (s === 17) { val = 0; rep = 3 + bits(3); }
        else { val = 0; rep = 11 + bits(7); }
        while (rep--) longueurs[k++] = val;
      }
      tl = table(longueurs.subarray(0, nl)); td = table(longueurs.subarray(nl));
    } else throw new Error('bloc deflate invalide');
    for (;;) {
      var sym = symbole(tl);
      if (sym < 256) { out.push(sym); continue; }
      if (sym === 256) break;
      sym -= 257;
      var longueur = LBASE[sym] + bits(LEXTRA[sym]), ds = symbole(td), dist = DBASE[ds] + bits(DEXTRA[ds]), from = out.length - dist;
      if (from < 0) throw new Error('distance invalide');
      for (i = 0; i < longueur; i++) out.push(out[from + i]);
    }
  } while (!dernier);
  return out;
}

var CP1252_ = { 128: 8364, 130: 8218, 131: 402, 132: 8222, 133: 8230, 134: 8224, 135: 8225, 136: 710, 137: 8240, 138: 352, 139: 8249, 140: 338, 142: 381,
  145: 8216, 146: 8217, 147: 8220, 148: 8221, 149: 8226, 150: 8211, 151: 8212, 152: 732, 153: 8482, 154: 353, 155: 8250, 156: 339, 158: 382, 159: 376 };

/**
 * Le texte d'un flux de contenu de page : une ligne par positionnement (Tm, Td, T*) ou par bloc BT.
 * images : ajoute une ligne « \u0001Nom » à chaque image dessinée (opérateur Do), là où elle est dans la page.
 */
function texteContenu_(octets, images) {
  var n = octets.length, i = 0, lignes = [], ligne = '', attente = [], nom = '';
  function finLigne() { if (ligne.trim()) lignes.push(ligne.replace(/\s+/g, ' ').trim()); ligne = ''; }
  function car(c) { return String.fromCharCode(CP1252_[c] || c); }
  while (i < n) {
    var c = octets[i];
    if (c === 40) {                                   // ( chaîne )
      var prof = 1, s = ''; i++;
      while (i < n && prof > 0) {
        c = octets[i++];
        if (c === 92) {                               // \
          var e = octets[i++];
          if (e === 110) s += '\n'; else if (e === 114) s += '\r'; else if (e === 116) s += '\t'; else if (e === 98 || e === 102) s += ' ';
          else if (e >= 48 && e <= 55) { var o = e - 48, k = 0; while (k < 2 && octets[i] >= 48 && octets[i] <= 55) { o = o * 8 + octets[i++] - 48; k++; } s += car(o & 255); }
          else if (e === 13) { if (octets[i] === 10) i++; } else if (e === 10) { /* suite de ligne */ }
          else s += car(e);
        } else if (c === 40) { prof++; s += '('; }
        else if (c === 41) { prof--; if (prof > 0) s += ')'; }
        else s += car(c);
      }
      attente.push(s);
    } else if (c === 60 && octets[i + 1] !== 60) {    // <hexa> : police à codes sur 2 octets, non lue
      while (i < n && octets[i] !== 62) i++; i++;
    } else if (c === 37) {                            // % commentaire
      while (i < n && octets[i] !== 10 && octets[i] !== 13) i++;
    } else if ((c >= 65 && c <= 90) || (c >= 97 && c <= 122) || c === 39 || c === 34 || c === 42) {   // opérateur
      var j = i; while (j < n && ((octets[j] >= 65 && octets[j] <= 90) || (octets[j] >= 97 && octets[j] <= 122) || octets[j] === 42 || octets[j] === 39 || octets[j] === 34)) j++;
      var op = String.fromCharCode.apply(null, octets.slice(i, j)); i = j;
      if (op === 'Do' && images) { finLigne(); lignes.push('\u0001' + nom); }
      else if (op === 'TJ' || op === 'Tj') ligne += attente.join('');
      else if (op === "'" || op === '"') { finLigne(); ligne += attente.join(''); }
      else if (op === 'Tm' || op === 'Td' || op === 'TD' || op === 'T*' || op === 'BT' || op === 'ET') finLigne();
      attente = [];
    } else if (c === 47) {                            // /Nom
      var d0 = ++i; while (i < n && octets[i] > 32 && octets[i] !== 47 && octets[i] !== 40 && octets[i] !== 91 && octets[i] !== 93 && octets[i] !== 60 && octets[i] !== 62) i++;
      if (images) nom = String.fromCharCode.apply(null, octets.slice(d0, i));
    } else i++;
  }
  finLigne();
  return lignes.join('\n');
}

/** Tout le texte d'un PDF (Uint8Array), dans l'ordre du fichier. Renvoie '' si rien de lisible. images : voir texteContenu_. */
function texteDuPdf_(d, images) {
  function cherche(motif, depuis) {
    var m0 = motif.charCodeAt(0), L = motif.length;
    boucle: for (var p = depuis; p <= d.length - L; p++) {
      if (d[p] !== m0) continue;
      for (var q = 1; q < L; q++) if (d[p + q] !== motif.charCodeAt(q)) continue boucle;
      return p;
    }
    return -1;
  }
  var pages = [], pos = 0;
  for (;;) {
    var s = cherche('stream', pos);
    if (s < 0) break;
    if (s >= 3 && d[s - 3] === 101 && d[s - 2] === 110 && d[s - 1] === 100) { pos = s + 6; continue; }   // « endstream »
    var debut = s + 6; if (d[debut] === 13) debut++; if (d[debut] === 10) debut++;
    var fin = cherche('endstream', debut); if (fin < 0) break;
    var o = s; while (o > 0 && s - o < 3000 && !(d[o] === 111 && d[o + 1] === 98 && d[o + 2] === 106)) o--;   // « obj » qui précède
    var dict = String.fromCharCode.apply(null, d.subarray(o, s));
    pos = fin + 9;
    if (dict.indexOf('/FlateDecode') < 0 || /\/Length1|\/XRef|\/Image|\/ObjStm|\/FontFile/.test(dict)) continue;
    try {
      var contenu = inflate_(d, debut), t = texteContenu_(contenu, images);
      if (t) pages.push(t);
    } catch (err) { /* flux illisible : ignoré */ }
  }
  return pages.join('\n');
}
var RE_BOITE_ = /\b(?:GNY-(?:BPI|BPE)-[A-Z0-9]{4,6}(?:-\d{2})?|(?:PBO|BPE|BPI|PEC|PEO)-\d{2}-\d{3}-\d{3}-[A-Z0-9]{4}(?:-\d)?)\b/;

/**
 * Les informations utiles d'un rapport d'intervention (texte du PDF, une cellule par ligne) : horaires réels, défaut
 * corrigé ou non, réparation définitive ou non, dommage réseau signalé, compte-rendu du technicien, boîte où il est
 * intervenu, PM, nature et cause du défaut. Champ absent du rapport : ''.
 */
function champsRapport_(t) {
  t = String(t || '');
  function un(re) { var m = re.exec(t); return m ? m[1].replace(/\s+/g, ' ').trim() : ''; }
  // la réponse est sur la ligne qui suit le libellé ; si cette ligne est déjà le libellé suivant, il n'y a pas de réponse
  function reponse(re) { var v = un(re); return /^(Nature|Cause|Localisation|Photo|Description|Détail|Rapport d|Quel|La réparation)/i.test(v) ? '' : v; }
  function boites(src) {
    var out = [], re = new RegExp(RE_BOITE_.source, 'gi'), m;
    while ((m = re.exec(src))) { var b = m[0].toUpperCase(); if (out.indexOf(b) < 0) out.push(b); }
    return out;
  }
  // avant « Compte-rendu d'activités » : le ticket tel qu'il a été créé ; après : ce que le technicien a rempli
  var coupe = t.search(/Compte-rendu d['’]activit/i), terrain = coupe < 0 ? '' : t.slice(coupe), bTerrain = boites(terrain);
  // tableau « Compte-rendu d'activités » : date, début, fin, activité, durée
  var re = /(\d{2}\/\d{2}\/\d{4})\s+(\d{1,2}:\d{2})\s+(\d{1,2}:\d{2})\s+([^\n\d][^\n]*?)\s+(\d{1,3}:\d{2})(?=\s|$)/g, m, acts = [], minutes = 0;
  while ((m = re.exec(t))) { acts.push({ jour: m[1], debut: m[2], fin: m[3] }); minutes += +m[5].split(':')[0] * 60 + +m[5].split(':')[1]; }
  var dernier = acts[acts.length - 1];
  // adresse du site : « Adresse / rue / commune / (code postal) », écrite comme dans l'export : rue, code postal, commune
  var ma = /(?:^|\n)Adresse\n([\s\S]{0,200}?)\n\((\d{5})\)/.exec(t), adresse = '';
  if (ma) { var la = ma[1].split('\n'), commune = la.pop(); adresse = [la.join(' '), ma[2], commune].filter(function (x) { return x; }).join(' ').replace(/\s+/g, ' ').trim(); }
  // ce qui fixe la GTR : les étiquettes du début de la description ([Niveau 3], [04h HO]…) et « OTR 10J OUVRES » s'il est écrit
  var md = /(?:^|\n)Description\n([\s\S]*)/.exec(coupe < 0 ? t : t.slice(0, coupe)), desc = md ? md[1] : '';
  // les autres réponses du rapport, gardées telles quelles dans une colonne (pas encore exploitées par la carte)
  var autres = [];
  [['localisation du défaut', /Localisation du défaut[^\n]*\n([^\n]+)/i],
   ['description du défaut', /Description complémentaire du défaut[^\n]*\n([^\n]+)/i],
   ['position PM', /Position\s*(?:au\s*)?PM\s*:?\s*([^\n]{3,30})/i],
   ['référence client', /Référence du client\s*:\s*([A-Z]{2}\d{3}-[A-Z]{5}-\d{3})/i],
   ['PTO', /\bPTO\s*:\s*((?:TN|GE)-[A-Z0-9]{3,6}-[A-Z0-9]{3,5})/i],
   ['plusieurs intervenants', /Plusieurs intervenants[\s\S]{0,60}?\?\s*(Oui|Non)(?![a-zà-ÿ])/i],
   ['accès à une PTO', /accès à au moins\s+une PTO\s*\?\s*(Oui|Non)(?![a-zà-ÿ])/i],
   ['jarretière présente à l\'arrivée', /jarretière au\s+PM était-elle présente\s*\?\s*(Oui|Non)(?![a-zà-ÿ])/i],
   ['coupleur de la jarretière', /coupleur la jarretière[\s\S]{0,40}?\?\s*([^\n]+)/i],
   ['validée par le NOC ou le Patrimoine', /validée par le\s+NOC ou le Patrimoine\s*\?\s*(Oui|Non)(?![a-zà-ÿ])/i],
   ['matériel utilisé', /Du matériel[^\n?]{0,30}\?\s*(Oui|Non)(?![a-zà-ÿ])/i],
   ['Cerfa, plainte ou constat', /Cerfa\s*\/\s*Plainte\s*\/\s*Constat\n(Oui|Non)(?![a-zà-ÿ])/i],
   ['expertise', /agit-il d['’]une expertise\s*\?\s*(Oui|Non)(?![a-zà-ÿ])/i],
   ['mise à jour NETGEO nécessaire', /mise\s+à jour NETGEO est-elle\s+nécessaire\s*\?\s*(Oui|Non)(?![a-zà-ÿ])/i]].forEach(function (x) {
    var v = reponse(x[1]), h = x[2] ? un(x[2]) : '';
    if (v) autres.push(x[0] + ' : ' + v + (h ? ' ' + h : ''));
  });
  // fin d'impact client (côté opérateur, pour information : la GTR s'arrête à la fin de l'intervention). Date et heure.
  var jourImpact = un(/date de fin d['’]impact[^\n?]*(?:\n[^\n?]*)?\?\s*(\d{2}\/\d{2}\/\d{4})/i);
  var heureImpact = un(/heure de fin d['’]impact[^\n?]*(?:\n[^\n?]*)?\?\s*(\d{1,2}:\d{2})/i);
  var finImpact = jourImpact && heureImpact ? jourImpact + ' ' + heureImpact : '';
  var otr = /OTR\s*\d+\s*J\s*OUVR[A-ZÉ]*/i.exec(desc);
  var etiquettes = (desc.slice(0, 250).match(/\[[^\]\n]{1,30}\]/g) || []).join('') + (otr ? ' GTR: ' + otr[0] : '');
  return {
    num: un(/Intervention\s+N°\s*(\S+)/i),
    demande: un(/Date de la\s+demande\s+(\d{2}\/\d{2}\/\d{4}(?:\s+\d{1,2}:\d{2})?)/i),
    debut: acts.length ? acts[0].jour + ' ' + acts[0].debut : '',
    fin: acts.length ? dernier.jour + ' ' + dernier.fin : '',
    duree: minutes ? Math.floor(minutes / 60) + ' h ' + ('0' + minutes % 60).slice(-2) : '',
    corrige: un(/(?:a été corrigé|est-il (?:rétabli|résolu|corrigé)|a-t-il été (?:rétabli|résolu|corrigé))\s*\?\s*(Oui|Non)(?![a-zà-ÿ])/i),
    definitive: un(/réparation est-elle définitive\s*\?\s*(Oui|Non)(?![a-zà-ÿ])/i),
    dommage: un(/dommage réseau[\s\S]{0,300}?\?\s*(?:\([^)]*\)\s*)?(Oui|Non)(?![a-zà-ÿ])/i),
    actions: un(/Détail des actions effectuées\n([\s\S]*?)\n(?=La réparation|Photo|Rapport d['’]intervention)/i),
    resume: un(/Résumé final de l['’]intervention[^\n]*\n([\s\S]*?)\n(?=Du matériel|Cerfa|S['’]agit-il|Mise à jour Netgeo|Photos? |Rapport d['’]intervention)/i),
    // la boîte où le technicien est intervenu (celle qu'il cite), sinon celle du ticket
    boites: (bTerrain.length ? bTerrain : boites(coupe < 0 ? t : t.slice(0, coupe))).join(', '),
    pm: un(/(?:^|\n)Site\n([^\n]+)/),
    adresse: adresse, etiquettes: etiquettes, autres: autres.join(' · '), finImpact: finImpact,
    nature: reponse(/Nature du défaut[^\n]*\n([^\n]+)/i),
    cause: reponse(/Cause du défaut[^\n]*\n([^\n]+)/i)
  };
}

// ---------- analyse des photos des rapports par l'IA (Claude, d'Anthropic) ----------
// Facultatif : rien ne part tant que la clé API n'est pas enregistrée (menu Analyse des photos (IA)). Les photos de chaque
// rapport lu partent par lots à l'API d'Anthropic (moitié prix ; réponse en quelques minutes, une heure au plus en général),
// avec le strict nécessaire : horaires, commune, boîte et PM, réponses du rapport et compte-rendu sans n° de téléphone.
// Verdict dans la colonne « Photos (IA) » : « OK (9 photos) — … » ou « NO OK — à vérifier : … ».

var IA_MODELE = 'claude-opus-5-5';   // le modèle qui regarde les photos (moins cher : 'claude-sonnet-5-5' ou 'claude-haiku-5-5', voir IA_PRIX_LOT)
var IA_EFFORT = 'medium';            // réflexion : low, medium (conseillé ici), high
var IA_JOURS = 3;                    // seuls les rapports reçus ces 3 derniers jours partent (pas tout l'historique)
var IA_PAR_PASSAGE = 6;              // rapports envoyés à chaque passage (toutes les 5 minutes)
var IA_MAX_PAR_JOUR = 80;            // garde-fou sur le coût
var IA_PHOTOS_MAX = 16;              // photos regardées par rapport, au plus (les premières du rapport)
var IA_PRIX_LOT = { 'claude-opus-5-5': [2, 10], 'claude-sonnet-5-5': [1, 5], 'claude-haiku-5-5': [0.05, 0.25] };   // $ par million de jetons (entrée, sortie), prix des lots
var IA_API = 'https://api.anthropic.com', IA_VERSION = '2026-10-21';
var PDF_LUS_ = {};                   // rapports téléchargés pendant ce passage (lireRapports_), pour ne pas les retélécharger

var IA_CONSIGNES_ = [
  "Tu fais le contrôle qualité des interventions de maintenance sur un réseau de fibre optique FTTH (Grand Nancy), à partir des photos que le technicien a mises dans son rapport d'intervention Praxedo. Ton verdict est lu par la personne qui pilote l'activité : « OK » valide l'intervention, « A_VERIFIER » lui demande de regarder elle-même les photos.",
  "",
  "Ce que montrent les photos d'habitude : boîtiers ouverts (PBO, BPE, BPI) avec leurs cassettes d'épissure, laser rouge pour repérer une fibre, réflectomètre, étiquettes de câble, armoires de PM avec tiroirs optiques et jarretières, câbles, chambres et poteaux. Chaque photo porte en général la date, l'heure, les coordonnées GPS et l'adresse.",
  "",
  "Une photo prise pendant le travail est normale : boîtier ou cassette ouverts, laser allumé, fibres en cours de soudure, outils posés. Ne le compte pas comme un problème.",
  "",
  "Relève seulement ce qu'un chef d'équipe voudrait faire corriger ou vérifier :",
  "- fibres ou surlongueurs qui traînent hors des cassettes ou des guides, lovage absent ou en vrac, fibre pincée, écrasée, cassée ou trop pliée, protection d'épissure hors de son logement ;",
  "- sur une photo de fin de travail : cassette ou boîtier mal refermés, capot absent, câble mal fixé, boîtier abîmé ;",
  "- au PM : jarretières en vrac, qui pendent devant les tiroirs ou gênent l'accès, surlongueurs non rangées dans les guides, connecteur sans bouchon ;",
  "- étiquette absente ou illisible sur la photo d'étiquette ;",
  "- dommage visible sur le réseau (boîtier cassé, câble arraché, poteau ou chambre abîmés) ;",
  "- photo qui ne correspond pas à sa rubrique, ou trop floue ou trop sombre pour juger ce qu'elle devait montrer ;",
  "- photos qui contredisent le compte-rendu du technicien.",
  "",
  "Gravité : « fort » pour ce qui peut provoquer une panne ou une non-conformité nette ; « moyen » pour un travail peu soigné à reprendre ou un doute réel ; « faible » pour une simple remarque (photo floue sans importance, pas de photo du travail terminé, même photo dans deux rubriques). Verdict A_VERIFIER s'il y a au moins un problème fort ou moyen, sinon OK.",
  "",
  "Reste factuel : décris ce que tu vois, sans supposer ce qui est hors de la photo. Si une photo est trop petite ou trop floue pour trancher sur un point important, dis-le en « moyen » plutôt que de deviner. N'invente pas de problème : beaucoup d'interventions sont simplement OK.",
  "",
  "Numérote les photos comme dans le message (photo 1, photo 2…). Pour chacune, recopie la date, l'heure et les coordonnées GPS imprimées dessus, chiffre par chiffre ; laisse vide ce que tu ne lis pas avec certitude. Écris en français, avec les mots du métier (cassette, lovage, jarretière, tiroir, PBO, BPE, PM)."
].join('\n');

var IA_SCHEMA_ = {
  type: 'object', additionalProperties: false, required: ['verdict', 'resume', 'problemes', 'photos'],
  properties: {
    verdict: { type: 'string', enum: ['OK', 'A_VERIFIER'] },
    resume: { type: 'string', description: 'Une phrase courte en français pour la personne qui pilote l\'activité (25 mots au plus)' },
    problemes: { type: 'array', items: { type: 'object', additionalProperties: false, required: ['photo', 'gravite', 'probleme'], properties: {
      photo: { type: 'integer', description: 'N° de la photo concernée, 0 si cela concerne tout le rapport' },
      gravite: { type: 'string', enum: ['fort', 'moyen', 'faible'] },
      probleme: { type: 'string', description: 'Ce qui ne va pas, en quelques mots concrets (15 mots au plus)' } } } },
    photos: { type: 'array', items: { type: 'object', additionalProperties: false, required: ['photo', 'contenu', 'horodatage', 'gps'], properties: {
      photo: { type: 'integer' },
      contenu: { type: 'string', description: 'Ce que montre la photo, en quelques mots' },
      horodatage: { type: 'string', description: 'Date et heure imprimées sur la photo, au format AAAA-MM-JJ HH:MM, ou chaîne vide si rien de lisible' },
      gps: { type: 'string', description: 'Coordonnées imprimées sur la photo en degrés décimaux « latitude longitude » (ex. 48.66455 6.20019), ou chaîne vide' } } } }
  }
};

/**
 * Les photos d'un rapport PDF (Uint8Array) : les images JPEG dans l'ordre du rapport, chacune avec la ou les rubriques
 * où le technicien l'a mise (« Photo du défaut », « Photo de la réparation mise en place »…). Les logos, les petites
 * images et les formats que l'IA ne lit pas (CMYK, JPEG 2000) sont laissés de côté.
 * → [{ octets (Uint8Array), largeur, hauteur, rubriques: [] }]
 */
function photosDuPdf_(d) {
  var s = '', T = 8192;
  for (var k = 0; k < d.length; k += T) s += String.fromCharCode.apply(null, d.subarray(k, k + T));
  // les objets « N 0 obj << … /Subtype /Image … /DCTDecode … >> stream … endstream »
  var images = {}, ordreFichier = [], reObj = /(\d+)\s+\d+\s+obj\b/g, m;
  while ((m = reObj.exec(s))) {
    var dd = reObj.lastIndex, e0 = s.indexOf('endobj', dd), st = s.indexOf('stream', dd);
    if (st < 0 || (e0 >= 0 && e0 < st)) continue;
    var es = s.indexOf('endstream', st + 6);
    if (es < 0) break;
    reObj.lastIndex = es + 9;   // on saute les données (binaires)
    var dict = s.slice(dd, st);
    if (!/\/Subtype\s*\/Image/.test(dict) || !/\/DCTDecode/.test(dict) || /\/JPXDecode|\/DeviceCMYK/.test(dict)) continue;
    var a = st + 6; if (d[a] === 13) a++; if (d[a] === 10) a++;
    var b = es; while (b > a && (d[b - 1] === 10 || d[b - 1] === 13 || d[b - 1] === 32)) b--;
    var w = +((/\/Width\s+(\d+)/.exec(dict) || [])[1] || 0), h = +((/\/Height\s+(\d+)/.exec(dict) || [])[1] || 0);
    if (d[a] !== 0xFF || d[a + 1] !== 0xD8 || w < 120 || h < 120 || w * h < 40000 || Math.max(w, h) > 4 * Math.min(w, h)) continue;
    images[m[1]] = { debut: a, fin: b, largeur: w, hauteur: h };
    ordreFichier.push(m[1]);
  }
  // le nom de chaque image dans les pages (/X5 25 0 R) ; un même nom pour deux images : l'ordre du fichier suffira
  var noms = {}, doublon = false, rx = /\/XObject\s*<<([^>]*)>>/g, mx;
  while ((mx = rx.exec(s))) {
    var rn = /\/([^\s\/<>\[\]()]+)\s+(\d+)\s+\d+\s+R/g, mn;
    while ((mn = rn.exec(mx[1]))) { if (noms[mn[1]] && noms[mn[1]] !== mn[2]) doublon = true; noms[mn[1]] = mn[2]; }
  }
  var photos = [], parNum = {};
  function ajouter(num, rubrique) {
    var im = images[num];
    if (!im) return;
    var p = parNum[num];
    if (!p) { p = parNum[num] = { octets: d.subarray(im.debut, im.fin), largeur: im.largeur, hauteur: im.hauteur, rubriques: [] }; photos.push(p); }
    if (rubrique && p.rubriques.indexOf(rubrique) < 0) p.rubriques.push(rubrique);
  }
  if (!doublon && ordreFichier.length) {
    // dans le texte des pages, la rubrique (« Photo de l'étiquette ») vient juste avant son image
    var rubrique = '', depuis = -9;
    texteDuPdf_(d, true).split('\n').forEach(function (l, i) {
      if (l.charAt(0) === '\u0001') { ajouter(noms[l.slice(1)], i - depuis <= 3 ? rubrique : ''); rubrique = ''; }
      else if (/^Photos?\b/i.test(l)) { rubrique = l.trim(); depuis = i; }
      else if (rubrique && i - depuis <= 2 && l.length <= 40 && /^[a-zà-ÿ(]/.test(l)) rubrique += ' ' + l.trim();   // libellé sur deux lignes
    });
  }
  ordreFichier.forEach(function (num) { ajouter(num, ''); });
  return photos;
}

/** Les paramètres d'une demande d'analyse (Messages API) pour une ligne de l'onglet Fins et ses photos. */
function requetePhotos_(r, photos) {
  var E = ENTETES_FINS, g = function (nom) { return String(r[E.indexOf(nom)] || '').trim(); };
  var tel = /(?:\+33\s?|0)[1-9](?:[ .-]?\d{2}){4}/g, mail = /[^\s@]+@[^\s@]+\.[a-z]{2,}/gi;
  var debut = g('Début'), fin = g('Fin');
  var autres = g('Autres éléments du rapport').split(' · ').filter(function (x) { return x && !/^(référence client|PTO)\b/i.test(x); }).join(' · ');
  var contexte = ["Rapport d'intervention de maintenance fibre (FTTH).",
    'Sur place : ' + (debut ? 'le ' + debut.slice(0, 10) + ' de ' + debut.slice(11) + ' à ' + (fin.slice(0, 10) === debut.slice(0, 10) ? fin.slice(11) : fin) : 'horaires inconnus') +
      (g('Commune') ? ' · commune : ' + g('Commune') : ''),
    'Boîte(s) : ' + (g('Boîte(s)') || 'non précisée') + ' · PM (site) : ' + (g('PM (site)') || 'non précisé'),
    'Défaut corrigé : ' + (g('Défaut corrigé') || '?') + ' · réparation définitive : ' + (g('Réparation définitive') || '?') +
      (g('Nature du défaut') ? ' · nature : ' + g('Nature du défaut') : '') + (g('Cause du défaut') ? ' · cause : ' + g('Cause du défaut') : ''),
    autres ? 'Autres réponses du rapport : ' + autres : '',
    'Compte-rendu du technicien : « ' + (g('Compte-rendu du technicien').replace(tel, '[tél.]').replace(mail, '[mail]').slice(0, 1500) || '(vide)') + ' »',
    'Le rapport contient ' + photos.length + ' photo' + (photos.length > 1 ? 's' : '') + ', dans son ordre. La rubrique où le technicien a mis chaque photo est indiquée juste avant elle.'
  ].filter(function (x) { return x; }).join('\n');
  var blocs = [{ type: 'text', text: contexte }];
  photos.forEach(function (p, i) {
    blocs.push({ type: 'text', text: 'Photo ' + (i + 1) + ' — ' + (p.rubriques.length ? 'rubrique « ' + p.rubriques.join(' » et « ') + ' »' : 'rubrique inconnue') });
    blocs.push({ type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: Utilities.base64Encode(signes_(p.octets)) } });
  });
  blocs.push({ type: 'text', text: 'Donne ton verdict sur ces photos.' });
  return { model: IA_MODELE, max_tokens: 16000, system: IA_CONSIGNES_, messages: [{ role: 'user', content: blocs }],
           output_config: { effort: IA_EFFORT, format: { type: 'json_schema', schema: IA_SCHEMA_ } } };
}

/** Un appel à l'API d'Anthropic. corps : objet envoyé en JSON (POST), sinon GET. → { code, json, texte } */
function appelIA_(chemin, cle, corps) {
  var o = { method: corps ? 'post' : 'get', muteHttpExceptions: true, headers: { 'x-api-key': cle, 'anthropic-version': IA_VERSION } };
  if (corps) { o.contentType = 'application/json'; o.payload = JSON.stringify(corps); }
  var r = UrlFetchApp.fetch(IA_API + chemin, o), texte = r.getContentText(), json = null;
  try { json = JSON.parse(texte); } catch (err) { /* résultats d'un lot : une réponse JSON par ligne */ }
  return { code: r.getResponseCode(), json: json, texte: texte };
}

/** Le message d'erreur d'une réponse de l'API, en clair. */
function erreurIA_(res) {
  var m = res.json && res.json.error ? String(res.json.error.message || res.json.error.type || '') : String(res.texte || '').slice(0, 200);
  if (res.code === 401) return 'clé refusée (vérifiez-la dans la console Anthropic, rubrique API keys)';
  if (/credit balance/i.test(m)) return 'plus de crédit sur le compte Anthropic : ajoutez-en dans la console (rubrique Billing)';
  if (res.code === 404) return 'modèle ' + IA_MODELE + ' introuvable pour ce compte (' + m + ')';
  if (res.code === 429 || res.code === 529) return 'service occupé, nouvel essai au prochain passage (' + m + ')';
  return 'erreur ' + res.code + ' : ' + m;
}

function noterErreurIA_(texte) {
  PropertiesService.getScriptProperties().setProperty('IA_ERREUR', Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm') + ' : ' + texte);
}

/**
 * Le texte de la colonne « Photos (IA) », d'après la réponse de l'IA et deux contrôles faits ici : les photos sont-elles
 * datées du jour de l'intervention, et prises à moins de 5 km de l'adresse ? r : ligne de l'onglet Fins, n : nombre de
 * photos regardées, geo : { lat, lon } de l'adresse du rapport, ou null.
 */
function verdictPhotos_(rep, n, r, geo) {
  var E = ENTETES_FINS, P = [], J = 86400000;
  // les valeurs d'une liste fixe peuvent revenir avec d'autres majuscules : comparées en minuscules
  (rep.problemes || []).forEach(function (p) { var t = String(p.probleme || '').trim(); if (t) P.push({ photo: +p.photo || 0, g: String(p.gravite || '').toLowerCase(), t: t }); });
  function date(v) { var m = /^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{1,2}):(\d{2})/.exec(String(v || '')); return m ? new Date(+m[3], m[2] - 1, +m[1], +m[4], +m[5]) : null; }
  function jour(x) { return Utilities.formatDate(x, 'Europe/Paris', 'dd/MM/yyyy'); }
  var debut = date(r[E.indexOf('Début')]), fin = date(r[E.indexOf('Fin')]) || debut, autreJour = [], loin = [];
  (rep.photos || []).forEach(function (ph) {
    var m = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{1,2}):(\d{2})/.exec(String(ph.horodatage || '').trim());
    if (m && debut) {
      var q = new Date(+m[1], m[2] - 1, +m[3], +m[4], +m[5]), j0 = new Date(debut.getFullYear(), debut.getMonth(), debut.getDate()), j1 = new Date(fin.getFullYear(), fin.getMonth(), fin.getDate() + 1);
      if (q < j0 || q >= j1) autreJour.push([ph.photo, q]);
      else if (q < debut - 2 * 3600000 || q > fin.getTime() + 2 * 3600000) P.push({ photo: ph.photo, g: 'faible', t: 'prise à ' + m[4] + ':' + m[5] + ', loin des horaires déclarés' });
    }
    var c = /(-?\d{1,2}[.,]\d{3,})\D+?(-?\d{1,3}[.,]\d{3,})/.exec(String(ph.gps || ''));
    if (c && geo) {
      var la = +c[1].replace(',', '.'), lo = +c[2].replace(',', '.'), R = Math.PI / 180;
      var x = Math.pow(Math.sin((la - geo.lat) * R / 2), 2) + Math.cos(la * R) * Math.cos(geo.lat * R) * Math.pow(Math.sin((lo - geo.lon) * R / 2), 2);
      var km = 6371 * 2 * Math.asin(Math.sqrt(x));
      if (km > 5) loin.push([ph.photo, km]);
    }
  });
  if (autreJour.length) P.push({ photo: autreJour.length === 1 ? autreJour[0][0] : 0, g: autreJour.length > 1 ? 'fort' : 'moyen',
    t: (autreJour.length === 1 ? 'datée du ' + jour(autreJour[0][1]) : autreJour.length + ' photos datées d\'un autre jour (' +
      autreJour.map(function (a) { return jour(a[1]); }).filter(function (v, i, l) { return l.indexOf(v) === i; }).join(', ') + ')') +
      (debut ? ', intervention du ' + jour(debut) : '') + (autreJour.length > 1 ? ' : photos d\'une autre intervention ?' : '') });
  if (loin.length) P.push({ photo: loin.length === 1 ? loin[0][0] : 0, g: 'moyen', t: (loin.length === 1 ? 'prise à ' : loin.length + ' photos prises à ') +
    Math.round(Math.max.apply(null, loin.map(function (l) { return l[1]; }))) + ' km de l\'adresse du rapport' });
  var lib = function (p) { return (p.photo ? 'photo ' + p.photo + ' : ' : '') + p.t + (p.g === 'fort' ? ' (important)' : ''); };
  var graves = P.filter(function (p) { return p.g === 'fort'; }).concat(P.filter(function (p) { return p.g === 'moyen'; })), faibles = P.filter(function (p) { return p.g !== 'fort' && p.g !== 'moyen'; });
  var resume = String(rep.resume || '').trim(), nb = n + ' photo' + (n > 1 ? 's' : '');
  var t = /^a_v[ée]rifier$/i.test(String(rep.verdict || '').trim()) || graves.length
    ? 'NO OK — à vérifier (' + nb + ') : ' + (graves.length ? graves.map(lib).join(' · ') : resume || faibles.map(lib).join(' · '))
    : 'OK (' + nb + ') — ' + (resume || 'rien à signaler') + (faibles.length ? ' · remarque : ' + faibles.map(lib).join(' ; ') : '');
  return t.length > 900 ? t.slice(0, 897) + '…' : t;
}

/** Le texte de la colonne « Photos (IA) » pour un résultat de lot (voir la doc Message Batches). k : erreurs déjà comptées. */
function resultatPhotos_(res, r, n, geo, k) {
  var essai = function (pourquoi) { return 'erreur ' + (k + 1) + ' : ' + pourquoi; };
  if (!res || res.type === 'canceled' || res.type === 'expired') return essai('analyse non faite (lot ' + (res ? res.type : 'perdu') + ')');
  if (res.type === 'errored') {
    var e = JSON.stringify(res.error || {});
    return /invalid_request/.test(e) ? 'À vérifier à la main — photos refusées par l\'IA (' + ((/"message":"([^"]{0,150})/.exec(e) || [])[1] || 'requête invalide') + ')' : essai('erreur du service');
  }
  var msg = res.message || {};
  if (msg.stop_reason === 'refusal') return 'À vérifier à la main — l\'IA n\'a pas voulu analyser ces photos';
  if (msg.stop_reason === 'max_tokens') return essai('réponse coupée');
  var bloc = (msg.content || []).filter(function (b) { return b.type === 'text'; })[0], rep = null;
  try { rep = JSON.parse(bloc.text); } catch (err) { return essai('réponse illisible'); }
  return verdictPhotos_(rep, n, r, geo);
}

/**
 * À chaque passage (déclencheur), si une clé est enregistrée : récupère les lots terminés et écrit leurs verdicts,
 * puis envoie les photos des derniers rapports lus. → { recus, envoyes, sansPhoto, attente } ou null (pas activé, occupé).
 */
function analyserPhotos_(force) {
  var cle = PropertiesService.getUserProperties().getProperty('CLE_ANTHROPIC');
  if (!cle) return null;
  var verrou = LockService.getDocumentLock();   // pas celui de la lecture des mails : la carte n'attend pas
  if (!verrou || !verrou.tryLock(5000)) return null;
  try {
    var b = { recus: 0, envoyes: 0, sansPhoto: 0, attente: 0 };
    recupererLotsIA_(cle, b);
    envoyerLotIA_(cle, b, force);
    if (b.recus) PropertiesService.getScriptProperties().deleteProperty('CONTROLES_LE');   // onglet Qualité refait au passage suivant
    return b;
  } finally { verrou.releaseLock(); }
}

function envoyerLotIA_(cle, b, force) {
  var props = PropertiesService.getScriptProperties(), f = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_FINS);
  var n = f ? f.getLastRow() - 1 : 0, E = ENTETES_FINS, iP = E.indexOf(COL_PHOTOS), aujourd = Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyy-MM-dd');
  if (n < 1) return;
  grille_(f, 1, E.length);
  // après un refus (clé, crédit, service occupé) : pause, pour ne pas retélécharger les rapports toutes les 5 minutes
  if (!force && Date.now() < +(props.getProperty('IA_PAUSE') || 0)) return;
  var compte = JSON.parse(props.getProperty('IA_JOUR') || '{}');
  if (compte.jour !== aujourd) compte = { jour: aujourd, n: 0 };
  var place = Math.min(IA_PAR_PASSAGE, IA_MAX_PAR_JOUR - compte.n);
  if (place <= 0) return;
  // les 8 premières colonnes et le verdict suffisent pour choisir ; la ligne entière n'est lue que pour les rapports choisis
  var v = f.getRange(2, 1, n, 8).getDisplayValues(), vp = f.getRange(2, iP + 1, n, 1).getDisplayValues();
  var limite = new Date(), requetes = [], suivis = [], debut = Date.now();
  limite.setHours(0, 0, 0, 0); limite.setDate(limite.getDate() - IA_JOURS + 1);   // aujourd'hui et les 2 jours d'avant
  for (var i = n - 1; i >= 0 && requetes.length < place && Date.now() - debut < 60000; i--) {   // les plus récents d'abord
    var r = v[i], etat = vp[i][0], k = +((/^erreur (\d)/.exec(etat) || [])[1] || 0);
    if (r[7] !== 'oui' || !r[6] || (etat && !(k && k < 3))) continue;
    var m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(r[0]);
    if (!m || new Date(+m[3], m[2] - 1, +m[1]) < limite) continue;
    if (!/^https:\/\/([a-z0-9-]+\.)*praxedo\.com\//i.test(r[4])) continue;   // on n'ouvre que les liens de Praxedo
    var ecrire = function (ligne, txt) { f.getRange(ligne + 2, iP + 1, 1, 2).setNumberFormat('@').setValues([[texte_(txt), Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm')]]); };
    try {
      var octets = PDF_LUS_[r[4]];   // déjà téléchargé par lireRapports_ pendant ce passage ?
      if (!octets) {
        var rep = UrlFetchApp.fetch(r[4], { muteHttpExceptions: true });
        if (rep.getResponseCode() !== 200) throw new Error('rapport non téléchargé (réponse ' + rep.getResponseCode() + ')');
        octets = Uint8Array.from(rep.getContent());
      }
      var photos = photosDuPdf_(octets);
      if (!photos.length) { ecrire(i, 'Sans photo — le rapport ne contient aucune photo'); b.sansPhoto++; continue; }
      var id = ('p' + i + '_' + String(r[6]).replace(/[^A-Za-z0-9_-]/g, '')).slice(0, 64);
      requetes.push({ custom_id: id, params: requetePhotos_(f.getRange(i + 2, 1, 1, E.length).getDisplayValues()[0], photos.slice(0, IA_PHOTOS_MAX)) });
      suivis.push({ id: id, mail: r[6], ligne: i, n: Math.min(photos.length, IA_PHOTOS_MAX), k: k });
    } catch (err) {
      console.warn('Photos ' + r[1] + ' : ' + err);
      ecrire(i, 'erreur ' + (k + 1) + ' : ' + String(err.message || err).slice(0, 150));
    }
  }
  if (!requetes.length) return;
  var res = appelIA_('/v1/messages/batches', cle, { requests: requetes });
  if (res.code !== 200 || !res.json || !res.json.id) {   // renvoyés après la pause
    noterErreurIA_(erreurIA_(res));
    props.setProperty('IA_PAUSE', String(Date.now() + (res.code === 429 || res.code >= 500 ? 15 : 60) * 60000));
    return;
  }
  props.deleteProperty('IA_PAUSE');
  var lots = JSON.parse(props.getProperty('IA_LOTS') || '[]');
  lots.push({ id: res.json.id, le: Date.now(), r: suivis.map(function (s) { return [s.id, s.mail, s.n, s.k]; }) });
  props.setProperty('IA_LOTS', JSON.stringify(lots));
  compte.n += requetes.length;
  props.setProperty('IA_JOUR', JSON.stringify(compte));
  props.deleteProperty('IA_ERREUR');
  var quand = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm');
  suivis.forEach(function (s) { f.getRange(s.ligne + 2, iP + 1, 1, 2).setNumberFormat('@').setValues([["en cours d'analyse", quand]]); });
  b.envoyes += requetes.length;
}

function recupererLotsIA_(cle, b) {
  var props = PropertiesService.getScriptProperties(), lots = JSON.parse(props.getProperty('IA_LOTS') || '[]');
  if (!lots.length) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet(), f = grille_(ss.getSheetByName(ONGLET_FINS), 1, ENTETES_FINS.length), E = ENTETES_FINS, iP = E.indexOf(COL_PHOTOS);
  var v = null, ligneDe = {}, geo = null, restants = [], conso = { n: 0, entree: 0, sortie: 0 };
  function lire() {   // l'onglet, relu une fois : les lignes ont pu bouger (tri, ajout)
    if (v) return;
    v = f.getRange(2, 1, Math.max(f.getLastRow() - 1, 1), E.length).getDisplayValues();
    v.forEach(function (r, i) { if (r[6]) ligneDe[r[6]] = i; });
    geo = {};
    var fa = ss.getSheetByName(ONGLET_ADRESSES);
    if (fa) fa.getDataRange().getValues().slice(1).forEach(function (r) { if (r[1] !== '' && r[2] !== '') geo[r[0]] = { lat: +r[1], lon: +r[2] }; });
  }
  function ecrire(s, texte) {
    lire();
    var i = ligneDe[s[1]];
    if (i == null) return;
    f.getRange(i + 2, iP + 1, 1, 2).setNumberFormat('@').setValues([[texte_(texte), Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy HH:mm')]]);
    if (!/^erreur/.test(texte)) b.recus++;   // les erreurs repartent au passage suivant
  }
  lots.forEach(function (lot) {
    var st = appelIA_('/v1/messages/batches/' + lot.id, cle);
    var perdu = st.code === 404 || (st.code === 200 && st.json && st.json.processing_status !== 'ended' && Date.now() - lot.le > 30 * 3600000);
    if (perdu) { lot.r.forEach(function (s) { ecrire(s, resultatPhotos_(null, null, 0, null, s[3])); }); return; }
    if (st.code !== 200 || !st.json || st.json.processing_status !== 'ended') {
      if (st.code !== 200) noterErreurIA_(erreurIA_(st));
      restants.push(lot); b.attente += lot.r.length; return;
    }
    var res = appelIA_('/v1/messages/batches/' + lot.id + '/results', cle);
    if (res.code !== 200) { noterErreurIA_(erreurIA_(res)); restants.push(lot); return; }
    var vus = {};
    res.texte.split('\n').forEach(function (l) {
      var x = null;
      try { x = l.trim() ? JSON.parse(l) : null; } catch (err) { x = null; }
      var s = x && lot.r.filter(function (y) { return y[0] === x.custom_id; })[0];
      if (!s) return;
      vus[s[0]] = 1;
      lire();
      var r = v[ligneDe[s[1]]], u = x.result && x.result.message && x.result.message.usage;
      if (u) { conso.n++; conso.entree += (u.input_tokens || 0) + (u.cache_creation_input_tokens || 0) + (u.cache_read_input_tokens || 0); conso.sortie += u.output_tokens || 0; }
      if (r) ecrire(s, resultatPhotos_(x.result, r, s[2], geo[r[E.indexOf('Adresse')]] || null, s[3]));
    });
    lot.r.forEach(function (s) { if (!vus[s[0]]) ecrire(s, resultatPhotos_(null, null, 0, null, s[3])); });
  });
  props.setProperty('IA_LOTS', JSON.stringify(restants));
  if (conso.n) {
    var mois = Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyy-MM'), c = JSON.parse(props.getProperty('IA_CONSO') || '{}');
    if (c.mois !== mois) c = { mois: mois, rapports: 0, entree: 0, sortie: 0 };
    c.rapports += conso.n; c.entree += conso.entree; c.sortie += conso.sortie;
    props.setProperty('IA_CONSO', JSON.stringify(c));
    props.deleteProperty('IA_ERREUR');
  }
}

/** L'état de l'analyse des photos : { actif, erreur, attente, mois: { rapports, cout ($) } }. */
function etatIA_() {
  var props = PropertiesService.getScriptProperties(), c = JSON.parse(props.getProperty('IA_CONSO') || '{}'), prix = IA_PRIX_LOT[IA_MODELE] || [0, 0];
  var mois = Utilities.formatDate(new Date(), 'Europe/Paris', 'yyyy-MM'), lots = JSON.parse(props.getProperty('IA_LOTS') || '[]');
  var actif = false;
  try { actif = !!PropertiesService.getUserProperties().getProperty('CLE_ANTHROPIC'); } catch (err) { /* autre compte */ }
  return { actif: actif, erreur: actif ? props.getProperty('IA_ERREUR') || '' : '', attente: lots.reduce(function (a, l) { return a + l.r.length; }, 0),
           mois: c.mois === mois ? { rapports: c.rapports, cout: (c.entree * prix[0] + c.sortie * prix[1]) / 1e6 } : { rapports: 0, cout: 0 } };
}

/** Un essai, avec la vraie forme des demandes (sans photo) : la clé, le crédit et le modèle. → { ok, message } */
function testerCleIA_(cle) {
  var res = appelIA_('/v1/messages', cle, { model: IA_MODELE, max_tokens: 4000, system: IA_CONSIGNES_,
    messages: [{ role: 'user', content: [{ type: 'text', text: 'Essai de réglage, sans photo : réponds OK, avec « essai » comme résumé et des listes vides.' }] }],
    output_config: { effort: 'low', format: { type: 'json_schema', schema: IA_SCHEMA_ } } });
  if (res.code !== 200 || !res.json) return { ok: false, message: erreurIA_(res) };
  var bloc = (res.json.content || []).filter(function (x) { return x.type === 'text'; })[0];
  try { JSON.parse(bloc.text); } catch (err) { return { ok: false, message: 'réponse inattendue (' + (res.json.stop_reason || '?') + ')' }; }
  return { ok: true, message: '' };
}

/** Menu « Analyse des photos (IA) → Activer, voir le coût… » : la clé API (gardée dans les propriétés du propriétaire), l'état, le coût. */
function reglerPhotosIA() {
  proprietaire_();
  var ui = SpreadsheetApp.getUi(), up = PropertiesService.getUserProperties(), e = etatIA_();
  var etat = e.actif ? 'Activée (' + IA_MODELE + '). Ce mois-ci : ' + e.mois.rapports + ' rapport' + (e.mois.rapports > 1 ? 's' : '') + ' analysé' + (e.mois.rapports > 1 ? 's' : '') +
      ', coût estimé ' + e.mois.cout.toFixed(2).replace('.', ',') + ' $.' + (e.attente ? ' En cours : ' + e.attente + '.' : '') + (e.erreur ? '\nDernier problème : ' + e.erreur : '')
    : "Pas activée. Les photos de chaque rapport lu partiraient à l'API d'Anthropic (Claude), qui répond OK ou NO OK — à vérifier.";
  var r = ui.prompt('Analyse des photos par l\'IA', etat + '\n\n' + (e.actif ? 'Pour changer de clé, collez la nouvelle. Pour arrêter l\'analyse, tapez STOP.'
    : 'Pour l\'activer, collez votre clé API Anthropic (elle commence par sk-ant-).'), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var t = r.getResponseText().trim();
  if (!t) return;
  if (/^stop$/i.test(t)) { up.deleteProperty('CLE_ANTHROPIC'); ui.alert("Analyse des photos arrêtée. Les verdicts déjà écrits restent dans l'onglet Fins d'intervention."); return; }
  if (!/^sk-ant-[A-Za-z0-9_-]{20,}$/.test(t)) { ui.alert("Ce n'est pas une clé API Anthropic : elle commence par sk-ant-."); return; }
  var test = testerCleIA_(t);
  if (!test.ok) { ui.alert('La clé ne marche pas : ' + test.message); return; }
  up.setProperty('CLE_ANTHROPIC', t);
  PropertiesService.getScriptProperties().deleteProperty('IA_ERREUR');
  PropertiesService.getScriptProperties().deleteProperty('IA_PAUSE');
  ui.alert("C'est activé. Les rapports lus ces " + IA_JOURS + " derniers jours partent au prochain passage (5 minutes), puis chaque nouveau rapport. " +
    "Les verdicts arrivent en général en moins d'une heure dans la colonne « Photos (IA) » de l'onglet Fins d'intervention, dans l'onglet Qualité et dans le rapport du soir.");
}

/** Menu « Analyse des photos (IA) → Analyser maintenant » : un passage tout de suite. Réservé au propriétaire. */
function analyserPhotosMaintenant() {
  proprietaire_();
  var ui = SpreadsheetApp.getUi(), b = analyserPhotos_(true);
  if (!b) { ui.alert(etatIA_().actif ? 'Une analyse est déjà en cours : réessayez dans une minute.' : "L'analyse des photos n'est pas activée (menu Analyse des photos (IA) → Activer…)."); return; }
  var e = etatIA_();
  ui.alert([b.recus ? b.recus + ' verdict' + (b.recus > 1 ? 's' : '') + ' écrit' + (b.recus > 1 ? 's' : '') + ' dans la colonne « Photos (IA) ».' : '',
    b.envoyes ? b.envoyes + ' rapport' + (b.envoyes > 1 ? 's' : '') + ' envoyé' + (b.envoyes > 1 ? 's' : '') + ' à l\'analyse : réponse en général en moins d\'une heure.' : '',
    b.sansPhoto ? b.sansPhoto + ' rapport' + (b.sansPhoto > 1 ? 's' : '') + ' sans photo.' : '',
    b.attente ? b.attente + ' rapport' + (b.attente > 1 ? 's' : '') + ' en cours d\'analyse.' : '',
    e.erreur ? 'Problème : ' + e.erreur : '',
    !b.recus && !b.envoyes && !b.sansPhoto && !b.attente && !e.erreur ? 'Rien de nouveau à analyser.' : ''].filter(function (x) { return x; }).join('\n'));
}

// ---------- outils ----------

/** Codes d'accès à la carte, dans l'onglet Accès (créés au hasard la première fois, modifiables). */
function acces_() {
  var f = feuille_(ONGLET_ACCES, ['Qui', 'Code', 'Ce qu\'ils voient']);
  var code = function () { return String(Math.floor(100000 + Math.random() * 900000)); };
  if (f.getLastRow() < 3) {
    f.getRange(2, 1, 2, 3).setNumberFormat('@').setValues([
      ["Conduite d'activité", code(), 'La carte (tickets, priorités, nouveaux, terminés), les listes de contrôle et le bilan. Rien à modifier.'],
      ['Techniciens', code(), 'La carte (tickets, priorités, nouveaux, terminés) et leurs interventions. Rien à modifier.']]);
    f.setColumnWidth(1, 160); f.setColumnWidth(2, 100); f.setColumnWidth(3, 600);
  }
  // la personne qui gère les annulations : son code (carte) et son adresse mail en colonne D (annulation par mail)
  if (f.getLastRow() < 4) {
    var v0 = f.getRange(2, 2, 2, 1).getDisplayValues(), c = code();
    while (c === v0[0][0].trim() || c === v0[1][0].trim()) c = code();
    f.getRange(4, 1, 1, 3).setNumberFormat('@').setValues([['Annulations et corrélations', c, 'Comme la conduite d\'activité, et peut modifier depuis la carte : retirer ' +
      'un ticket annulé dans Praxedo, regrouper des tickets sous un ticket principal (corrélation), et les remettre ; ajouter une note à un ticket (justifié, ' +
      'relance…) ; noter OK ou NO OK une intervention après avoir vérifié ses photos. En colonne D : son adresse mail, pour annuler ou regrouper par mail ' +
      '(« ANNULÉ » ou « CORRÉLATION » dans l\'objet, avec les n° des tickets).']]);
    f.getRange(1, 4).setValue('Adresse mail (annulations et corrélations par mail)').setFontWeight('bold');
    f.getRange(4, 4).setNumberFormat('@');
    f.setColumnWidth(4, 280);
  }
  var v = f.getRange(2, 2, 3, 1).getDisplayValues(), mails = String(f.getRange(4, 4).getDisplayValue() || '').toLowerCase().split(/[\s,;]+/);
  return { ca: v[0][0].trim(), tech: v[1][0].trim(), annul: v[2][0].trim(),
           mailsAnnul: mails.filter(function (m) { return /^[^\s@]+@[^\s@]+\.[a-z]{2,}$/.test(m); }) };
}

/** Bloque les fonctions qui modifient le Sheet quand elles ne sont pas lancées par son propriétaire (depuis le Sheet ou l'éditeur). */
function proprietaire_() {
  var actif = '', moi = '';
  try { actif = Session.getActiveUser().getEmail(); moi = Session.getEffectiveUser().getEmail(); } catch (err) { /* inconnu : refusé */ }
  if (!actif || actif !== moi) throw new Error('Réservé au propriétaire du Google Sheet.');
}

/** [reçu le, N°, commune, intervention du, rapport, objet, id] ou null si l'objet n'a pas de n° de ticket. */
function ligneFin_(objet, corps, recu, id) {
  var m = RE_NUM.exec(objet);
  if (!m) return null;
  var apres = objet.slice(m.index + m[0].length).replace(/^\s*-\s*/, '').trim();
  var jour = /faite ce jour\s+(\d{2}\/\d{2}\/\d{4})/i.exec(corps);
  var lien = /https:\/\/[A-Za-z0-9.-]*praxedo\.com\/[^\s<>"]*\/render/i.exec(corps);
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
  var texte = String(corps).replace(/\r/g, '').replace(/\u00a0/g, ' ').replace(/<[^>\n]*>/g, '');  // le texte brut ajoute les liens en <https://…>
  // le bloc « info ticket » : ce qui suit le libellé jusqu'à « info Client », que le contenu soit à la ligne ou non
  var i = texte.search(/info\s+ticket/i), desc = '';
  if (i >= 0) {
    var reste = texte.slice(i).replace(/^info\s+ticket[^:\n]{0,5}:?/i, ''), j = reste.search(/info\s+client/i);
    desc = (j >= 0 ? reste.slice(0, j) : reste).replace(/^[\s*_]+|[\s*_]+$/g, '').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n');
  }
  // niveau, GTR en heures et RDV client : s'ils ne sont pas en tête du bloc, on les cherche dans tout le mail (objet compris)
  if (!/niveau\s*\d|\[\s*\d+\s*h\s*ho\s*\]/i.test(desc.slice(0, 250))) {
    // entre crochets dans le mail ([Niveau 3]) ; sans crochets, seulement dans l'objet (« niveau 2 » d'un parking n'est pas une GTR)
    var tout = objet + '\n' + texte, devant = [], mn = /\[\s*niveau\s*(\d)\s*\]/i.exec(tout) || /niveau\s*(\d)/i.exec(objet), mh = /\[\s*(\d+)\s*h\s*ho\s*\]/i.exec(tout);
    if (mh) devant.push('[' + mh[1] + 'h HO]');
    if (mn) devant.push('[Niveau ' + mn[1] + ']');
    if (/\[\s*RDVCLI\s*\]/i.test(tout) && !/\[\s*RDVCLI\s*\]/i.test(desc.slice(0, 250))) devant.push('[RDVCLI]');
    if (devant.length) desc = devant.join('') + (desc ? ' ' + desc : '');
  }
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

/**
 * Agrandit la grille d'un onglet avant d'y écrire jusqu'à la ligne « lignes » et la colonne « colonnes » : un onglet neuf
 * n'a que 1 000 lignes et 26 colonnes (A à Z), et Google refuse d'écrire au-delà.
 */
function grille_(f, lignes, colonnes) {
  var ml = f.getMaxRows(), mc = f.getMaxColumns();
  if (lignes > ml) f.insertRowsAfter(ml, lignes - ml + 200);
  if (colonnes > mc) f.insertColumnsAfter(mc, colonnes - mc);
  return f;
}

function feuille_(nom, entetes) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), f = ss.getSheetByName(nom);
  if (!f) {
    f = ss.insertSheet(nom);
    grille_(f, 1, entetes.length);
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
  if (lignes.length) grille_(f, f.getLastRow() + lignes.length, ENTETES_ADRESSES.length).getRange(f.getLastRow() + 1, 1, lignes.length, ENTETES_ADRESSES.length).setValues(lignes.map(function (l) { return l.map(texte_); }));
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
