/**
 * Suivi des appels clients Isotuiles
 * A coller dans le Google Sheets "Suivi appels clients Isotuiles" :
 * Extensions > Apps Script, remplacer le contenu de Code.gs par ce fichier,
 * puis ajouter le fichier HTML "Formulaire" (voir Formulaire.html).
 *
 * Ce script fait :
 *  - un menu "Isotuiles" dans le Sheets ;
 *  - le remplissage automatique d'une ligne quand vous tapez un numéro ;
 *  - la création/mise à jour d'un rendez-vous "Rappeler ..." dans Google Agenda ;
 *  - un email récapitulatif chaque matin ;
 *  - une petite application web (formulaire de saisie rapide sur téléphone,
 *    réception automatique des appels manqués envoyés par MacroDroid,
 *    et API JSON pour le module de l'ERP Isotéam).
 */

// ---------- Structure de l'onglet Appels (colonnes, base 1) ----------
const ONGLET_APPELS = 'Appels';
const ONGLET_LISTES = 'Listes';
const COL = {
  id: 1, date: 2, tel: 3, nom: 4, email: 5, ville: 6, motif: 7, demande: 8,
  etape: 9, rappel: 10, dateRappel: 11, priorite: 12, tentatives: 13,
  dernierContact: 14, notes: 15, ref: 16, source: 17, agenda: 18
};
const NB_COL = 18; // la colonne S (appels du même numéro) est une formule
const ETAPES_CLOSES = ['Terminé', 'Sans suite'];

// ---------- Paramètres (onglet Listes, colonnes G:H) ----------
function parametres_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(ONGLET_LISTES);
  const lignes = sh.getRange('G2:H20').getValues();
  const p = {};
  lignes.forEach(function (l) { if (l[0]) p[String(l[0]).trim()] = l[1]; });
  return {
    email: p["Email du récap du matin"] || Session.getEffectiveUser().getEmail(),
    heureRecap: Number(p["Heure d'envoi du récap (0 à 23)"] || 7),
    dureeRappel: Number(p["Durée d'un créneau de rappel (minutes)"] || 15),
    delaiRappel: Number(p["Délai de rappel par défaut (heures)"] || 24),
    heureCreneau: Number(p["Heure du premier créneau de rappel (0 à 23)"] || 9),
    cle: String(p["Clé d'accès ERP (à changer)"] || '')
  };
}

// ---------- Menu ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Isotuiles')
    .addItem('Nouvel appel (fiche rapide)', 'ouvrirFormulaire')
    .addItem('Envoyer le récap maintenant', 'envoyerRecap')
    .addItem('Synchroniser l\'agenda', 'synchroniserAgenda')
    .addSeparator()
    .addItem('Installer les automatisations', 'installer')
    .addToUi();
}

function ouvrirFormulaire() {
  const html = HtmlService.createTemplateFromFile('Formulaire');
  html.listes = listes_();
  SpreadsheetApp.getUi().showSidebar(html.evaluate().setTitle('Nouvel appel'));
}

/** A lancer une fois : crée le déclencheur du matin et celui des modifications. */
function installer() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  const p = parametres_();
  ScriptApp.newTrigger('envoyerRecap').timeBased().everyDays(1).atHour(p.heureRecap)
    .inTimezone('Europe/Paris').create();
  ScriptApp.newTrigger('surModification').forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit().create();
  SpreadsheetApp.getActive().toast('Récap programmé chaque jour vers ' + p.heureRecap +
    'h. Le suivi automatique des modifications est actif.', 'Isotuiles', 8);
}

// ---------- Outils ----------
function feuille_() { return SpreadsheetApp.getActive().getSheetByName(ONGLET_APPELS); }

function listes_() {
  const v = SpreadsheetApp.getActive().getSheetByName(ONGLET_LISTES).getRange('A2:E20').getValues();
  const col = function (i) { return v.map(function (l) { return l[i]; }).filter(String); };
  return { motifs: col(0), etapes: col(1), priorites: col(2), sources: col(4) };
}

/** Met un numéro au format 06 12 34 56 78 (gère +33, 0033, points, espaces). */
function normaliserTel_(t) {
  let s = String(t || '').replace(/[^\d+]/g, '');
  if (s.indexOf('+33') === 0) s = '0' + s.slice(3);
  else if (s.indexOf('0033') === 0) s = '0' + s.slice(4);
  else if (s.length === 9 && s[0] !== '0') s = '0' + s;
  if (/^0\d{9}$/.test(s)) return s.replace(/(\d{2})(?=\d)/g, '$1 ').trim();
  return s;
}

function prochainId_(sh) {
  const n = Math.max(sh.getLastRow() - 1, 0);
  let max = 0;
  if (n > 0) sh.getRange(2, COL.id, n, 1).getValues().forEach(function (l) {
    const m = String(l[0]).match(/(\d+)$/);
    if (m) max = Math.max(max, Number(m[1]));
  });
  return 'A-' + ('000' + (max + 1)).slice(-4);
}

/** Prochain créneau de rappel : maintenant + délai, ramené à un jour ouvré à l'heure du premier créneau. */
function dateRappelParDefaut_() {
  const p = parametres_();
  const d = new Date(Date.now() + p.delaiRappel * 3600 * 1000);
  if (d.getHours() < p.heureCreneau || d.getHours() >= 18) {
    if (d.getHours() >= 18) d.setDate(d.getDate() + 1);
    d.setHours(p.heureCreneau, 0, 0, 0);
  }
  while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1);
  return d;
}

/** Cherche le dernier dossier connu pour ce numéro. Renvoie {ligne, valeurs} ou null. */
function chercherParTel_(sh, tel, ligneExclue) {
  const n = sh.getLastRow() - 1;
  if (n < 1 || !tel) return null;
  const v = sh.getRange(2, 1, n, NB_COL).getValues();
  for (let i = v.length - 1; i >= 0; i--) {
    if (i + 2 === ligneExclue) continue;
    if (normaliserTel_(v[i][COL.tel - 1]) === tel) return { ligne: i + 2, valeurs: v[i] };
  }
  return null;
}

function maintenant_() { return new Date(); }

// ---------- Ajout d'un appel (formulaire, MacroDroid, ERP) ----------
/**
 * data : {tel, nom, email, ville, motif, demande, priorite, source, dateRappel (ISO), notes}
 * Si un dossier ouvert existe déjà pour ce numéro, il est mis à jour au lieu d'en créer un autre.
 */
function ajouterAppel(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = feuille_();
    const tel = normaliserTel_(data.tel);
    if (!tel) throw new Error('Numéro de téléphone manquant');
    const existant = chercherParTel_(sh, tel);
    const horodatage = Utilities.formatDate(maintenant_(), 'Europe/Paris', 'dd/MM HH:mm');

    if (existant && ETAPES_CLOSES.indexOf(existant.valeurs[COL.etape - 1]) === -1) {
      // Dossier encore ouvert : on complète au lieu de dupliquer
      const r = existant.ligne;
      const notes = existant.valeurs[COL.notes - 1];
      const ajout = horodatage + ' : ' + (data.source === 'Appel manqué (auto)' ? 'nouvel appel manqué' : 'nouvel appel') +
        (data.demande ? ' (' + data.demande + ')' : '');
      sh.getRange(r, COL.notes).setValue(notes ? notes + '\n' + ajout : ajout);
      sh.getRange(r, COL.rappel).setValue('Oui');
      if (!existant.valeurs[COL.dateRappel - 1] || existant.valeurs[COL.dateRappel - 1] < maintenant_()) {
        sh.getRange(r, COL.dateRappel).setValue(data.dateRappel ? new Date(data.dateRappel) : dateRappelParDefaut_());
      }
      if (existant.valeurs[COL.etape - 1] !== 'Nouveau') sh.getRange(r, COL.etape).setValue('À rappeler');
      ['nom', 'email', 'ville', 'motif'].forEach(function (k) {
        if (data[k] && !existant.valeurs[COL[k] - 1]) sh.getRange(r, COL[k]).setValue(data[k]);
      });
      synchroniserLigne_(sh, r);
      return { ok: true, id: existant.valeurs[COL.id - 1], misAJour: true };
    }

    // Nouveau dossier, en reprenant les coordonnées connues si le client a déjà appelé
    const connu = existant ? existant.valeurs : null;
    const ligne = new Array(NB_COL).fill('');
    ligne[COL.id - 1] = prochainId_(sh);
    ligne[COL.date - 1] = maintenant_();
    ligne[COL.tel - 1] = "'" + tel;
    ligne[COL.nom - 1] = data.nom || (connu ? connu[COL.nom - 1] : '');
    ligne[COL.email - 1] = data.email || (connu ? connu[COL.email - 1] : '');
    ligne[COL.ville - 1] = data.ville || (connu ? connu[COL.ville - 1] : '');
    ligne[COL.motif - 1] = data.motif || '';
    ligne[COL.demande - 1] = data.demande || '';
    ligne[COL.etape - 1] = data.source === 'Appel manqué (auto)' ? 'Nouveau' : 'À rappeler';
    ligne[COL.rappel - 1] = data.rappel === 'Non' ? 'Non' : 'Oui';
    ligne[COL.dateRappel - 1] = ligne[COL.rappel - 1] === 'Oui'
      ? (data.dateRappel ? new Date(data.dateRappel) : dateRappelParDefaut_()) : '';
    ligne[COL.priorite - 1] = data.priorite || (data.motif === 'SAV / réclamation' ? 'Haute' : 'Normale');
    ligne[COL.tentatives - 1] = 0;
    ligne[COL.notes - 1] = data.notes || (connu ? 'Client déjà connu (dossier ' + connu[COL.id - 1] + ')' : '');
    ligne[COL.source - 1] = data.source || 'Saisie manuelle';
    sh.appendRow(ligne);
    const r = sh.getLastRow();
    synchroniserLigne_(sh, r);
    return { ok: true, id: ligne[COL.id - 1], misAJour: false };
  } finally {
    lock.releaseLock();
  }
}

/** Met à jour un dossier depuis l'ERP. changes : {etape, rappel, dateRappel, notes, ...}. action "rappele" ou "reporter". */
function mettreAJour(id, changes, action) {
  const sh = feuille_();
  const n = sh.getLastRow() - 1;
  const ids = sh.getRange(2, COL.id, n, 1).getValues().map(function (l) { return l[0]; });
  const i = ids.indexOf(id);
  if (i === -1) throw new Error('Dossier introuvable : ' + id);
  const r = i + 2;
  changes = changes || {};
  if (action === 'rappele') {
    sh.getRange(r, COL.tentatives).setValue(Number(sh.getRange(r, COL.tentatives).getValue() || 0) + 1);
    sh.getRange(r, COL.dernierContact).setValue(maintenant_());
    if (!changes.etape) changes.etape = 'Rappelé, attente client';
    if (!changes.rappel) changes.rappel = 'Non';
  }
  if (action === 'reporter') {
    const d = new Date(Date.now() + 24 * 3600 * 1000);
    d.setHours(parametres_().heureCreneau, 0, 0, 0);
    while (d.getDay() === 0 || d.getDay() === 6) d.setDate(d.getDate() + 1);
    changes.dateRappel = d.toISOString();
    changes.rappel = 'Oui';
    sh.getRange(r, COL.tentatives).setValue(Number(sh.getRange(r, COL.tentatives).getValue() || 0) + 1);
  }
  Object.keys(changes).forEach(function (k) {
    if (!COL[k] || k === 'id') return;
    let v = changes[k];
    if (k === 'dateRappel') v = v ? new Date(v) : '';
    if (k === 'tel') v = "'" + normaliserTel_(v);
    if (k === 'notes' && action !== 'remplacerNotes') {
      const avant = sh.getRange(r, COL.notes).getValue();
      v = avant ? avant + '\n' + v : v;
    }
    sh.getRange(r, COL[k]).setValue(v);
  });
  if (changes.etape) {
    sh.getRange(r, COL.dernierContact).setValue(maintenant_());
    if (ETAPES_CLOSES.indexOf(changes.etape) !== -1) sh.getRange(r, COL.rappel).setValue('Non');
  }
  synchroniserLigne_(sh, r);
  return { ok: true, id: id };
}

/** Liste des dossiers au format JSON (pour l'ERP). */
function listerDossiers_() {
  const sh = feuille_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  const cles = Object.keys(COL);
  return sh.getRange(2, 1, n, NB_COL + 1).getValues().map(function (l) {
    const o = {};
    cles.forEach(function (k) {
      const v = l[COL[k] - 1];
      o[k] = v instanceof Date ? v.toISOString() : v;
    });
    o.appelsMemeNumero = l[NB_COL];
    return o;
  }).filter(function (o) { return o.id; });
}

// ---------- Saisie directe dans le Sheets ----------
/** Déclencheur installable "à la modification" (voir installer()). */
function surModification(e) {
  const sh = e.range.getSheet();
  if (sh.getName() !== ONGLET_APPELS) return;
  const r = e.range.getRow();
  const c = e.range.getColumn();
  if (r < 2 || e.range.getNumRows() > 1) return;

  // Un numéro tapé sur une ligne vide : on complète la fiche
  if (c === COL.tel && e.value) {
    const tel = normaliserTel_(e.value);
    e.range.setValue("'" + tel);
    if (!sh.getRange(r, COL.id).getValue()) {
      const autre = chercherParTel_(sh, tel, r);
      const connu = autre ? autre.valeurs : null;
      sh.getRange(r, COL.id).setValue(prochainId_(sh));
      sh.getRange(r, COL.date).setValue(maintenant_());
      if (connu && !sh.getRange(r, COL.nom).getValue()) {
        sh.getRange(r, COL.nom).setValue(connu[COL.nom - 1]);
        sh.getRange(r, COL.email).setValue(connu[COL.email - 1]);
        sh.getRange(r, COL.ville).setValue(connu[COL.ville - 1]);
      }
      if (!sh.getRange(r, COL.etape).getValue()) sh.getRange(r, COL.etape).setValue('À rappeler');
      if (!sh.getRange(r, COL.rappel).getValue()) sh.getRange(r, COL.rappel).setValue('Oui');
      if (!sh.getRange(r, COL.dateRappel).getValue()) sh.getRange(r, COL.dateRappel).setValue(dateRappelParDefaut_());
      if (!sh.getRange(r, COL.priorite).getValue()) sh.getRange(r, COL.priorite).setValue('Normale');
      if (sh.getRange(r, COL.tentatives).getValue() === '') sh.getRange(r, COL.tentatives).setValue(0);
      if (!sh.getRange(r, COL.source).getValue()) sh.getRange(r, COL.source).setValue('Saisie manuelle');
    }
    synchroniserLigne_(sh, r);
    return;
  }

  if (c === COL.etape) {
    sh.getRange(r, COL.dernierContact).setValue(maintenant_());
    if (ETAPES_CLOSES.indexOf(e.value) !== -1) sh.getRange(r, COL.rappel).setValue('Non');
    synchroniserLigne_(sh, r);
    return;
  }

  if ([COL.rappel, COL.dateRappel, COL.nom, COL.motif, COL.demande].indexOf(c) !== -1) {
    synchroniserLigne_(sh, r);
  }
}

// ---------- Google Agenda ----------
/**
 * Crée, déplace ou supprime le rendez-vous "Rappeler ..." d'une ligne.
 * L'identifiant de l'événement est gardé dans la colonne Agenda.
 */
function synchroniserLigne_(sh, r) {
  const v = sh.getRange(r, 1, 1, NB_COL).getValues()[0];
  const cal = CalendarApp.getDefaultCalendar();
  const idEvt = v[COL.agenda - 1];
  let evt = null;
  if (idEvt) { try { evt = cal.getEventById(idEvt); } catch (err) { evt = null; } }

  const actif = v[COL.rappel - 1] === 'Oui' && v[COL.dateRappel - 1] instanceof Date &&
    ETAPES_CLOSES.indexOf(v[COL.etape - 1]) === -1;

  if (!actif) {
    if (evt) evt.deleteEvent();
    if (idEvt) sh.getRange(r, COL.agenda).setValue('');
    return;
  }

  const p = parametres_();
  const debut = v[COL.dateRappel - 1];
  const fin = new Date(debut.getTime() + p.dureeRappel * 60000);
  const tel = v[COL.tel - 1];
  const titre = 'Rappeler ' + (v[COL.nom - 1] || tel) + (v[COL.motif - 1] ? ' (' + v[COL.motif - 1] + ')' : '');
  const description = [
    'Téléphone : ' + tel + '  (tel:' + String(tel).replace(/\s/g, '') + ')',
    v[COL.demande - 1] ? 'Demande : ' + v[COL.demande - 1] : '',
    v[COL.ville - 1] ? 'Chantier : ' + v[COL.ville - 1] : '',
    'Étape : ' + v[COL.etape - 1] + ' | Tentatives : ' + (v[COL.tentatives - 1] || 0),
    v[COL.notes - 1] ? 'Notes : ' + v[COL.notes - 1] : '',
    'Dossier ' + v[COL.id - 1] + ' : ' + SpreadsheetApp.getActive().getUrl()
  ].filter(String).join('\n');

  if (evt) {
    evt.setTitle(titre);
    evt.setTime(debut, fin);
    evt.setDescription(description);
  } else {
    evt = cal.createEvent(titre, debut, fin, { description: description });
    evt.addPopupReminder(10);
    sh.getRange(r, COL.agenda).setValue(evt.getId());
  }
}

function synchroniserAgenda() {
  const sh = feuille_();
  for (let r = 2; r <= sh.getLastRow(); r++) {
    if (sh.getRange(r, COL.id).getValue()) synchroniserLigne_(sh, r);
  }
  SpreadsheetApp.getActive().toast('Agenda synchronisé', 'Isotuiles', 5);
}

// ---------- Récap du matin ----------
function envoyerRecap() {
  const p = parametres_();
  const dossiers = listerDossiers_();
  const maintenant = new Date();
  const finJour = new Date(); finJour.setHours(23, 59, 59, 999);
  const hier = new Date(Date.now() - 24 * 3600 * 1000);

  const ouverts = dossiers.filter(function (d) { return ETAPES_CLOSES.indexOf(d.etape) === -1; });
  const aRappeler = ouverts.filter(function (d) { return d.rappel === 'Oui'; })
    .sort(function (a, b) { return new Date(a.dateRappel || 0) - new Date(b.dateRappel || 0); });
  const enRetard = aRappeler.filter(function (d) { return d.dateRappel && new Date(d.dateRappel) < maintenant; });
  const aujourdhui = aRappeler.filter(function (d) {
    return d.dateRappel && new Date(d.dateRappel) >= maintenant && new Date(d.dateRappel) <= finJour;
  });
  const nouveaux = dossiers.filter(function (d) { return d.date && new Date(d.date) >= hier; });
  const devisEnAttente = ouverts.filter(function (d) { return d.etape === 'Devis à faire'; });

  const fmt = function (iso) { return iso ? Utilities.formatDate(new Date(iso), 'Europe/Paris', 'EEE dd/MM HH:mm') : ''; };
  const tableau = function (liste) {
    if (!liste.length) return '<p style="color:#666">Rien.</p>';
    return '<table cellpadding="6" style="border-collapse:collapse;font-family:Arial;font-size:13px">' +
      '<tr style="background:#1F3864;color:#fff"><th align="left">Quand</th><th align="left">Client</th>' +
      '<th align="left">Téléphone</th><th align="left">Motif</th><th align="left">Demande</th><th align="left">Étape</th></tr>' +
      liste.map(function (d) {
        const telBrut = String(d.tel).replace(/\s/g, '');
        return '<tr style="border-bottom:1px solid #ddd"><td>' + fmt(d.dateRappel || d.date) + '</td><td>' +
          (d.nom || '<i>inconnu</i>') + '</td><td><a href="tel:' + telBrut + '">' + d.tel + '</a></td><td>' +
          (d.motif || '') + '</td><td>' + (d.demande || '') + '</td><td>' + d.etape + '</td></tr>';
      }).join('') + '</table>';
  };

  const total = enRetard.length + aujourdhui.length;
  const html = '<div style="font-family:Arial;font-size:14px">' +
    '<p>Bonjour,</p><p>Vous avez <b>' + total + ' rappel' + (total > 1 ? 's' : '') + '</b> à faire aujourd\'hui' +
    (enRetard.length ? ', dont <b style="color:#c0392b">' + enRetard.length + ' en retard</b>' : '') + '.</p>' +
    '<h3 style="color:#c0392b">En retard</h3>' + tableau(enRetard) +
    '<h3>Prévus aujourd\'hui</h3>' + tableau(aujourdhui) +
    '<h3>Nouveaux appels depuis hier</h3>' + tableau(nouveaux) +
    '<h3>Devis à faire</h3>' + tableau(devisEnAttente) +
    '<p><a href="' + SpreadsheetApp.getActive().getUrl() + '">Ouvrir le suivi des appels</a></p></div>';

  MailApp.sendEmail({
    to: p.email,
    subject: 'Rappels clients du ' + Utilities.formatDate(maintenant, 'Europe/Paris', 'dd/MM') + ' : ' +
      total + ' à faire' + (enRetard.length ? ' (' + enRetard.length + ' en retard)' : ''),
    htmlBody: html
  });
}

// ---------- Application web ----------
/**
 * GET  ?page=formulaire&cle=...          formulaire de saisie rapide (à mettre sur l'écran d'accueil du téléphone)
 * GET  ?action=liste&cle=...             liste JSON des dossiers (module ERP)
 * GET  ?action=manque&tel=...&cle=...    appel manqué envoyé par MacroDroid
 * POST {cle, action:"ajouter"|"maj"|"manque", ...}  (corps JSON, Content-Type text/plain)
 */
function doGet(e) {
  const q = e.parameter || {};
  const p = parametres_();
  if (!p.cle || q.cle !== p.cle) return json_({ ok: false, erreur: 'Clé invalide' });

  if (q.page === 'formulaire') {
    const t = HtmlService.createTemplateFromFile('Formulaire');
    t.listes = listes_();
    return t.evaluate().setTitle('Nouvel appel Isotuiles')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  try {
    if (q.action === 'liste') return json_({ ok: true, listes: listes_(), dossiers: listerDossiers_() });
    if (q.action === 'manque') {
      return json_(ajouterAppel({ tel: q.tel, nom: q.nom, source: 'Appel manqué (auto)' }));
    }
  } catch (err) {
    return json_({ ok: false, erreur: String(err.message || err) });
  }
  return json_({ ok: false, erreur: 'Action inconnue' });
}

function doPost(e) {
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) { d = e.parameter || {}; }
  const p = parametres_();
  if (!p.cle || d.cle !== p.cle) return json_({ ok: false, erreur: 'Clé invalide' });
  try {
    if (d.action === 'ajouter') return json_(ajouterAppel(d));
    if (d.action === 'manque') return json_(ajouterAppel({ tel: d.tel, nom: d.nom, source: 'Appel manqué (auto)' }));
    if (d.action === 'maj') return json_(mettreAJour(d.id, d.changes, d.sousAction));
  } catch (err) {
    return json_({ ok: false, erreur: String(err.message || err) });
  }
  return json_({ ok: false, erreur: 'Action inconnue' });
}

/** Appelé par le formulaire dans la barre latérale du Sheets. */
function ajouterDepuisFormulaire(data) { return ajouterAppel(data); }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
