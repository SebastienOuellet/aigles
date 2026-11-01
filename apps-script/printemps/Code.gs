/**
 * Printemps 2027 — Aigles (groupes 2019, 2020, 2021…)
 * Inscription / confirmation de présence / gestion des pratiques / rappels
 *
 * À coller dans Extensions > Apps Script du Google Sheet des inscriptions.
 */

const F_GROUPES = 'Groupes';
const F_PRAT = 'Pratiques';
const F_PARAM = 'Paramètres';
const F_JOURNAL = 'Journal';
const PREF_INSCR = 'Inscriptions ';
const PREF_PRES = 'Présences ';

const PARAM_DEFAUT = [
  ['NOM_CLUB', 'Aigles', 'Nom affiché sur le formulaire'],
  ['SAISON', 'Printemps 2027', 'Nom de la saison'],
  ['INSCRIPTIONS_OUVERTES', 'OUI', 'OUI = nouvelles inscriptions acceptées | NON = seulement les confirmations de présence'],
  ['CODE_GERANTE', '', 'OBLIGATOIRE — code de la page de gestion (ne pas partager)'],
  ['COURRIEL_NOTIF', '', 'Courriel de la gérante — reçoit chaque nouvelle inscription (plusieurs : séparer par des virgules)'],
  ['LIEN_SITE', 'https://sebastienouellet.github.io/aigles/printemps/', 'Adresse du formulaire (la gestion = LIEN_SITE + gerante.html)'],
  ['RAPPEL_ACTIF', 'OUI', 'OUI = courriel de rappel aux parents la veille de chaque pratique'],
  ['RAPPEL_HEURE', '17', 'Heure d’envoi des rappels (0 à 23). Relancer « Configurer » après un changement.']
];

const GROUPES_DEFAUT = [
  ['2019', 'Aigles 2019', 'OUI'],
  ['2020', 'Aigles 2020', 'OUI'],
  ['2021', 'Aigles 2021', 'OUI']
];

const PRAT_CHAMPS = [
  { k: 'groupe', t: 'Groupe', m: h => h === 'groupe' },
  { k: 'date', t: 'Date', m: h => h === 'date' },
  { k: 'heure', t: 'Heure', m: h => h === 'heure' },
  { k: 'endroit', t: 'Endroit', m: h => h === 'endroit' },
  { k: 'cout', t: 'Coût', m: h => h === 'cout' },
  { k: 'notes', t: 'Notes', m: h => h === 'notes' },
  { k: 'rappel', t: 'Rappel envoyé', m: h => h.startsWith('rappel') }
];

const PRES_FIXES = ['ID', 'Prénom', 'Nom', 'Position'];
const PRESENCES = ['Confirmé', 'Présent', 'Absent'];

/** Colonnes d'un onglet Inscriptions, retrouvées par leur titre; celles qui manquent sont ajoutées à droite. */
const CHAMPS = [
  { k: 'nom', t: 'Nom', m: h => h === 'nom' },
  { k: 'prenom', t: 'Prénom', m: h => h === 'prenom' },
  { k: 'parent', t: 'Nom des parents', m: h => h.startsWith('nom des p') || h.startsWith('nom du p') },
  { k: 'courriel', t: 'Adresse courriel', m: h => h.startsWith('adresse c') || h === 'courriel' },
  { k: 'position', t: 'Position', m: h => h === 'position' },
  { k: 'lance', t: 'Lance côté', m: h => h.startsWith('lance') },
  { k: 'dateNaissance', t: 'Date de naissance', m: h => h.startsWith('date de nai') },
  { k: 'nam', t: 'NAM', m: h => h === 'nam' || h.startsWith('nam ') || h.includes('assurance maladie') },
  { k: 'photos', t: 'Photos promo', m: h => h.includes('photo') },
  { k: 'id', t: 'ID', m: h => h === 'id' },
  { k: 'inscritLe', t: 'Inscrit le', m: h => h.startsWith('inscrit') },
  { k: 'source', t: 'Source', m: h => h === 'source' },
  { k: 'statut', t: 'Statut', m: h => h === 'statut' },
  { k: 'confirmeeLe', t: 'Confirmée le', m: h => h.startsWith('confirmee le') },
  { k: 'jeton', t: 'Jeton', m: h => h === 'jeton' }
];

const CHOIX = {
  position: ['Attaquant', 'Défenseur', 'Gardien'],
  lance: ['Gauche', 'Droite'],
  photos: ['Oui', 'Non']
};

/* ============================================================
 * Menu et installation
 * ============================================================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Printemps 2027')
    .addItem('Ouvrir la gestion et les liens', 'ouvrirGestion')
    .addSeparator()
    .addItem('Configurer (première fois / après un changement)', 'setup')
    .addItem('Synchroniser les présences', 'synchroniserMenu')
    .addItem('Importer une ancienne pratique (onglet ouvert)', 'importerPratique')
    .addItem('Envoyer les rappels de demain maintenant', 'rappelsMenu')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());
  const ui = SpreadsheetApp.getUi();
  const notes = [];

  // Paramètres
  let p = ss.getSheetByName(F_PARAM);
  if (!p) {
    p = ss.insertSheet(F_PARAM);
    p.getRange(1, 1, 1, 3).setValues([['Clé', 'Valeur', 'Description']]).setFontWeight('bold');
    p.getRange('B:B').setNumberFormat('@');
    p.setColumnWidth(1, 220); p.setColumnWidth(2, 360); p.setColumnWidth(3, 520);
    p.setFrozenRows(1);
  }
  const cles = p.getDataRange().getValues().map(r => String(r[0]));
  PARAM_DEFAUT.forEach(r => { if (cles.indexOf(r[0]) === -1) p.appendRow([r[0], String(r[1]), r[2]]); });
  if (cles.indexOf('NOM_EQUIPE') > -1) notes.push('La clé NOM_EQUIPE n’est plus utilisée (remplacée par NOM_CLUB et l’onglet Groupes) : tu peux la supprimer.');

  // Groupes
  let g = ss.getSheetByName(F_GROUPES);
  if (!g) {
    g = ss.insertSheet(F_GROUPES);
    g.getRange(1, 1, 1, 3).setValues([['Année', 'Nom', 'Ouvert']]).setFontWeight('bold').setBackground('#f5a800');
    g.getRange('A:A').setNumberFormat('@');
    g.getRange(2, 1, GROUPES_DEFAUT.length, 3).setValues(GROUPES_DEFAUT);
    g.setFrozenRows(1);
    notes.push('Onglet Groupes créé (2019, 2020, 2021). Ajoute une ligne pour un nouveau groupe.');
  }
  const groupes = lireGroupes_();
  if (!groupes.length) throw new Error('Aucun groupe dans l’onglet Groupes.');

  // Migration de l'ancienne version (un seul groupe) → premier groupe
  const premier = groupes[0].annee;
  const vieuxIns = ss.getSheetByName('Inscriptions'), vieuxPres = ss.getSheetByName('Présences');
  let migre = false;
  if (vieuxIns && !ss.getSheetByName(PREF_INSCR + premier)) { vieuxIns.setName(PREF_INSCR + premier); migre = true; }
  if (vieuxPres && !ss.getSheetByName(PREF_PRES + premier)) { vieuxPres.setName(PREF_PRES + premier); migre = true; }
  if (migre) notes.push('Données existantes déplacées dans « ' + PREF_INSCR + premier + ' » et « ' + PREF_PRES + premier + ' ».');

  // Pratiques (+ colonne Groupe : les anciennes pratiques vont au premier groupe)
  const pr = feuillePrat_();
  const CP = colonnesPrat_(pr);
  const nPr = derniereLigne_(pr, CP.date);
  if (nPr > 1) {
    const v = pr.getRange(2, CP.groupe, nPr - 1, 1).getValues();
    let n = 0;
    v.forEach((r, i) => { if (!String(r[0]).trim()) { pr.getRange(i + 2, CP.groupe).setValue(premier); n++; } });
    if (n) notes.push(n + ' pratique(s) existante(s) associée(s) au groupe ' + premier + '.');
  }

  // Inscriptions et Présences de chaque groupe (une erreur dans un groupe n'empêche pas les autres)
  let total = 0;
  const erreurs = [];
  groupes.forEach(gr => { try {
    const sh = feuilleInscr_(gr.annee);
    feuillePres_(gr.annee);
    const C = colonnesInscr_(sh);
    // Joueurs déjà là avant la colonne Statut = déjà acceptés
    if (C._ajouts.indexOf('statut') > -1) {
      const n = derniereLigne_(sh, C.nom, C.prenom);
      if (n > 1) {
        const noms = sh.getRange(2, C.nom, n - 1, 1).getValues();
        sh.getRange(2, C.statut, n - 1, 1).setValues(noms.map(r => [String(r[0]).trim() ? 'Confirmée' : '']));
      }
    }
    total += synchroniser_(gr.annee);
  } catch (e) {
    console.error('Groupe ' + gr.annee, e);
    erreurs.push(gr.annee + ' : ' + e.message + (e.stack ? ' [' + String(e.stack).split('\n')[1].trim() + ']' : ''));
  } });

  try { installerRappels_(); } catch (e) { erreurs.push('Rappels : ' + e.message); }

  const c = lireParam_();
  if (!c.CODE_GERANTE) notes.push('À faire : choisir un CODE_GERANTE dans Paramètres.');
  if (!c.COURRIEL_NOTIF) notes.push('À faire : inscrire le courriel de la gérante (COURRIEL_NOTIF) dans Paramètres.');
  ui.alert('Configuration terminée ✅\n\n' + groupes.map(x => x.nom).join(', ') + ' — ' + total + ' joueur(s).\n' +
    (oui_(c.RAPPEL_ACTIF) ? 'Rappels : chaque jour vers ' + c.RAPPEL_HEURE + ' h pour les pratiques du lendemain.\n' : 'Rappels désactivés.\n') +
    (notes.length ? '\n• ' + notes.join('\n• ') : '') +
    (erreurs.length ? '\n\n⚠ ERREURS — envoie ce message à Sébastien :\n• ' + erreurs.join('\n• ') : ''));
}

function installerRappels_() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'envoyerRappels') ScriptApp.deleteTrigger(t); });
  const c = lireParam_();
  if (!oui_(c.RAPPEL_ACTIF)) return;
  const h = Math.min(23, Math.max(0, parseInt(c.RAPPEL_HEURE, 10) || 17));
  ScriptApp.newTrigger('envoyerRappels').timeBased().everyDays(1).atHour(h).inTimezone(Session.getScriptTimeZone()).create();
}

function synchroniserMenu() {
  const n = lireGroupes_().reduce((s, g) => s + synchroniser_(g.annee), 0);
  SpreadsheetApp.getUi().alert('Présences synchronisées : ' + n + ' joueur(s).');
}

function rappelsMenu() {
  SpreadsheetApp.getUi().alert('Rappels : ' + envoyerRappels(true));
}

function ouvrirGestion() {
  const url = lireParam_().LIEN_SITE;
  const html =
    '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.6">' +
    '<p><b>Gestion des inscriptions et pratiques</b> (code requis)<br>' +
    '<a href="' + url + 'gerante.html" target="_blank" style="font-size:16px;font-weight:bold;color:#7a1426">Ouvrir la gestion →</a></p>' +
    '<p><b>Formulaire pour les parents</b><br><a href="' + url + '" target="_blank">' + esc_(url) + '</a></p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(440).setHeight(190), 'Printemps 2027');
}

/* ============================================================
 * API (appelée par le site) — POST text/plain {action, code?, data?}
 * ============================================================ */

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const d = req.data || {};
    const gerante = ['gestion', 'setPresence', 'ajouterPratique', 'joueur', 'confirmerInscription', 'changerGroupe', 'voirNam'];
    if (gerante.indexOf(req.action) > -1) verifierCode_(req.code);
    switch (req.action) {
      case 'config': out = configPublique_(); if (d.diag) out.diag = diagnostic_(); break;
      case 'inscrire': out = inscrire_(d); break;
      case 'confirmer': out = confirmer_(d); break;
      case 'reponseRappel': out = reponseRappel_(d); break;
      case 'gestion': out = gestion_(d.groupe); break;
      case 'setPresence': out = setPresence_(d); break;
      case 'ajouterPratique': out = ajouterPratique_(d); break;
      case 'joueur': out = ficheJoueur_(d); break;
      case 'confirmerInscription': out = confirmerInscription_(d); break;
      case 'changerGroupe': out = changerGroupe_(d); break;
      case 'voirNam': out = voirNam_(d); break;
      default: throw new Error('Action inconnue.');
    }
    out = { ok: true, data: out };
  } catch (err) {
    console.error(err);
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return HtmlService.createHtmlOutput('<p style="font-family:Arial;text-align:center;padding:40px">' +
    '<a href="' + lireParam_().LIEN_SITE + '" target="_top">Aller au formulaire →</a></p>');
}

/** Diagnostic : à quel Sheet le script parle et ce qu'il voit dans l'onglet Groupes (aucune donnée personnelle). */
function diagnostic_() {
  const ss = classeur_();
  const g = ss.getSheetByName(F_GROUPES);
  return {
    url: ss.getUrl(),
    onglets: ss.getSheets().map(s => s.getName()),
    groupesTrouve: !!g,
    groupesBrut: g && g.getLastRow() > 1 ? g.getRange(2, 1, g.getLastRow() - 1, 3).getValues().map(r => r.map(v => typeof v + ':' + v)) : [],
    groupesLus: lireGroupes_().map(x => x.annee)
  };
}

function configPublique_() {
  const c = lireParam_();
  return {
    club: c.NOM_CLUB, saison: c.SAISON, ouvert: oui_(c.INSCRIPTIONS_OUVERTES),
    feuille: classeur_().getName(),
    groupes: lireGroupes_().map(g => ({ annee: g.annee, nom: g.nom, ouvert: g.ouvert, prochaine: prochainePratique_(g.annee) })),
    choix: CHOIX
  };
}

/* ============================================================
 * Inscription, confirmation, réponse aux rappels
 * ============================================================ */

function inscrire_(d) {
  if (d.website) throw new Error('Soumission refusée.'); // champ piège anti-robot
  const c = lireParam_();
  if (!oui_(c.INSCRIPTIONS_OUVERTES)) throw new Error('Les inscriptions sont fermées. Communique avec la gérante.');

  const t = (v, max) => String(v == null ? '' : v).trim().slice(0, max || 200);
  const o = {
    prenom: t(d.prenom, 60), nom: t(d.nom, 60),
    parent: t(d.parent, 120), courriel: t(d.courriel, 120).toLowerCase(),
    position: t(d.position), lance: t(d.lance),
    dateNaissance: t(d.dateNaissance, 10), nam: normNam_(d.nam), photos: t(d.photos)
  };
  const err = [];
  if (!o.prenom || !o.nom) err.push('Prénom et nom du joueur requis.');
  if (!o.parent) err.push('Nom du parent requis.');
  if (!courrielValide_(o.courriel)) err.push('Courriel invalide.');
  ['position', 'lance', 'photos'].forEach(k => { if (CHOIX[k].indexOf(o[k]) === -1) err.push('Choix invalide : ' + k + '.'); });
  if (!/^\d{4}-\d{2}-\d{2}$/.test(o.dateNaissance)) err.push('Date de naissance invalide.');
  if (!/^[A-Z]{4} \d{4} \d{4}$/.test(o.nam)) err.push('NAM invalide (format : ABCD 1234 5678).');
  if (err.length) throw new Error(err.join(' '));

  const groupe = lireGroupes_().find(g => g.annee === o.dateNaissance.slice(0, 4));
  if (!groupe) throw new Error('Aucun groupe pour les enfants nés en ' + o.dateNaissance.slice(0, 4) + '. Communique avec la gérante.');
  if (!groupe.ouvert) throw new Error('Les inscriptions du groupe ' + groupe.nom + ' sont fermées.');
  o.statut = 'En attente';

  const lock = verrou_();
  let joueur;
  try {
    joueur = enregistrerJoueur_(o, 'Formulaire', groupe.annee);
    if (d.presentProchaine) confirmerPresenceProchaine_(joueur);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }

  const g = lireGroupes_().find(x => x.annee === joueur.annee) || groupe;
  notifier_(o, c, g, !!d.presentProchaine);
  return { prenom: o.prenom, nom: o.nom, groupe: g.nom,
           prochaine: d.presentProchaine ? prochainePratique_(g.annee) : null };
}

/** Joueur qui a déjà pratiqué : on le retrouve par son nom dans tous les groupes et on confirme sa prochaine pratique. */
function confirmer_(d) {
  if (d.website) throw new Error('Soumission refusée.');
  const prenom = String(d.prenom || '').trim().slice(0, 60);
  const nom = String(d.nom || '').trim().slice(0, 60);
  if (!prenom || !nom) throw new Error('Prénom et nom du joueur requis.');

  const lock = verrou_();
  let joueur;
  try {
    joueur = trouverParNom_(prenom, nom);
    if (!joueur) throw new Error('Joueur introuvable : « ' + prenom + ' ' + nom + ' ». Vérifie l’orthographe, ou reviens en arrière et choisis « Non » pour faire l’inscription.');
    if (!prochainePratique_(joueur.annee)) throw new Error('Aucune pratique à venir n’est encore prévue pour ce groupe. Reviens bientôt!');
    confirmerPresenceProchaine_(joueur);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  const g = lireGroupes_().find(x => x.annee === joueur.annee);
  return { prenom: joueur.prenom, nom: joueur.nom, groupe: g ? g.nom : joueur.annee,
           prochaine: prochainePratique_(joueur.annee), ficheIncomplete: !joueur.nam };
}

/** Lien du courriel de rappel : ?c=JETON&r=oui|non */
function reponseRappel_(d) {
  const jeton = String(d.jeton || '').trim();
  const rep = d.reponse === 'non' ? 'non' : 'oui';
  if (!/^[a-f0-9-]{20,}$/i.test(jeton)) throw new Error('Lien invalide.');
  const lock = verrou_();
  try {
    const j = trouverParJeton_(jeton);
    if (!j) throw new Error('Lien invalide ou expiré.');
    const p = prochainePratique_(j.annee);
    if (!p) throw new Error('Aucune pratique à venir pour ce groupe.');
    const sh = feuillePres_(j.annee);
    const r = lignePres_(sh, j);
    const col = colsPratique_(sh, p.date);
    const cell = sh.getRange(r, col.p);
    if (rep === 'non') { cell.setValue('Absent'); sh.getRange(r, col.paye).setValue(false); }
    else if (cell.getValue() !== 'Présent') cell.setValue('Confirmé');
    SpreadsheetApp.flush();
    const g = lireGroupes_().find(x => x.annee === j.annee);
    return { prenom: j.prenom, nom: j.nom, groupe: g ? g.nom : j.annee, prochaine: p, reponse: rep };
  } finally { lock.releaseLock(); }
}

/**
 * Ajoute un joueur dans « Inscriptions AAAA » (ou complète une fiche minimale) et dans « Présences AAAA ».
 * Doublon : même NAM dans n'importe quel groupe = refus.
 */
function enregistrerJoueur_(o, source, annee) {
  const tous = tousJoueurs_();
  if (o.nam && tous.some(j => normNam_(j.nam) === o.nam)) {
    throw new Error('Ce joueur est déjà inscrit (NAM déjà enregistré). Pour confirmer ta présence à une pratique, réponds « Oui » à la première question.');
  }
  // Fiche minimale au même nom (créée par une confirmation) → on la complète, là où elle est
  const cible = o.nam ? tous.find(j => !j.nam && memeNom_(j.prenom, j.nom, o.prenom, o.nom)) : null;
  if (cible) annee = cible.annee;

  const sh = feuilleInscr_(annee);
  const C = colonnesInscr_(sh);
  const id = cible && cible.id ? cible.id : prochainId_(tous.map(j => j.id));
  const r = cible ? cible.ligne : derniereLigne_(sh, C.nom, C.prenom) + 1;

  const set = (k, val, fmt) => { if (C[k] && val !== undefined && val !== '') { const g = sh.getRange(r, C[k]); g.setValue(val); if (fmt) g.setNumberFormat(fmt); } };
  Object.keys(o).forEach(k => { if (C[k] && k !== 'dateNaissance') set(k, sur_(o[k])); });
  if (o.dateNaissance) set('dateNaissance', new Date(o.dateNaissance + 'T12:00:00'), 'yyyy-mm-dd');
  set('id', id);
  if (!(cible && cible.jeton)) set('jeton', Utilities.getUuid());
  if (!cible) { set('inscritLe', new Date(), 'yyyy-mm-dd hh:mm'); set('source', source); }

  const j = { id: id, annee: annee, prenom: o.prenom, nom: o.nom, position: o.position || '', nam: o.nam || '' };
  lignePres_(feuillePres_(annee), j);
  return j;
}

function confirmerPresenceProchaine_(j) {
  const p = prochainePratique_(j.annee);
  if (!p) return;
  const sh = feuillePres_(j.annee);
  const r = lignePres_(sh, j);
  const col = colsPratique_(sh, p.date);
  const cell = sh.getRange(r, col.p);
  if (cell.getValue() !== 'Présent') cell.setValue('Confirmé');
}

/** Courriel à la gérante pour chaque nouvelle inscription (le NAM n'est jamais envoyé par courriel). */
function notifier_(o, c, groupe, presentProchaine) {
  if (!c.COURRIEL_NOTIF) return;
  const p = prochainePratique_(groupe.annee);
  const ligne = (k, v) => '<tr><td style="color:#666;padding:4px 14px 4px 0;vertical-align:top">' + k + '</td><td style="padding:4px 0"><b>' + v + '</b></td></tr>';
  try {
    MailApp.sendEmail({
      to: c.COURRIEL_NOTIF, replyTo: o.courriel, name: c.NOM_CLUB + ' — Inscriptions',
      subject: '[' + groupe.nom + '] Nouvelle inscription — ' + o.prenom + ' ' + o.nom,
      htmlBody:
        '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1a1416">' +
        '<p>Nouvelle inscription <b>en attente</b> pour <b>' + esc_(groupe.nom) + '</b> (' + esc_(c.SAISON) + ').</p>' +
        '<table style="border-collapse:collapse">' +
        ligne('Joueur', esc_(o.prenom + ' ' + o.nom)) + ligne('Date de naissance', esc_(o.dateNaissance)) +
        ligne('Position', esc_(o.position)) + ligne('Lance', esc_(o.lance)) + ligne('Parent', esc_(o.parent)) +
        ligne('Courriel', '<a href="mailto:' + esc_(o.courriel) + '">' + esc_(o.courriel) + '</a>') +
        ligne('Photos promo', esc_(o.photos)) + ligne('NAM', 'reçu ✓ (voir le Google Sheet)') +
        (p ? ligne('Prochaine pratique', esc_(p.date) + (presentProchaine ? ' — présence confirmée' : ' — pas confirmé')) : '') +
        '</table><p style="margin-top:18px"><a href="' + c.LIEN_SITE + 'gerante.html" style="color:#7a1426;font-weight:bold">' +
        'Ouvrir la gestion pour confirmer l’inscription →</a></p></div>'
    });
  } catch (e) { console.error('Notification', e); }
}

/* ============================================================
 * Rappels la veille des pratiques (déclencheur quotidien)
 * ============================================================ */

function envoyerRappels(manuel) {
  const c = lireParam_();
  if (manuel !== true && !oui_(c.RAPPEL_ACTIF)) return 'désactivés';
  const tz = Session.getScriptTimeZone();
  const demain = Utilities.formatDate(new Date(Date.now() + 24 * 3600 * 1000), tz, 'yyyy-MM-dd');
  const pr = feuillePrat_(), CP = colonnesPrat_(pr);
  const groupes = lireGroupes_();
  const resume = [];
  lirePratiques_().filter(p => p.date === demain && !p.rappel).forEach(p => {
    const g = groupes.find(x => x.annee === p.groupe);
    if (!g) return;
    synchroniser_(g.annee); // garantit un jeton à chaque joueur
    const joueurs = joueursGroupe_(g.annee).filter(j => adressesValides_(j.courriel).length);
    if (MailApp.getRemainingDailyQuota() < joueurs.length) {
      journal_('Rappel', g.nom + ' — quota de courriels insuffisant (' + MailApp.getRemainingDailyQuota() + ' restants pour ' + joueurs.length + ')');
    }
    let n = 0;
    joueurs.forEach(j => {
      if (MailApp.getRemainingDailyQuota() < 1) return;
      try { courrielRappel_(j, g, p, c); n++; } catch (e) { console.error('Rappel', j.id, e); }
    });
    pr.getRange(p.ligne, CP.rappel).setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm');
    resume.push(g.nom + ' : ' + n + ' courriel(s)');
    journal_('Rappel', g.nom + ' — pratique du ' + p.date + ' — ' + n + ' courriel(s)');
  });
  return resume.length ? resume.join(' · ') : 'aucune pratique demain (ou rappel déjà envoyé)';
}

function courrielRappel_(j, g, p, c) {
  const lien = r => c.LIEN_SITE + '?c=' + encodeURIComponent(j.jeton) + '&r=' + r;
  const bouton = (txt, url, bg) => '<a href="' + url + '" style="display:inline-block;padding:12px 20px;margin:4px 6px 4px 0;border-radius:10px;' +
    'background:' + bg + ';color:#fff;font-weight:bold;text-decoration:none">' + txt + '</a>';
  const opts = {
    to: adressesValides_(j.courriel).join(','), name: g.nom,
    subject: 'Rappel : pratique ' + g.nom + ' demain' + (p.heure ? ' à ' + p.heure : ''),
    htmlBody:
      '<div style="font-family:Arial,sans-serif;font-size:15px;color:#1a1416;line-height:1.5">' +
      '<p>Bonjour,</p><p>Petit rappel : <b>' + esc_(j.prenom) + '</b> a une pratique avec les <b>' + esc_(g.nom) + '</b> demain.</p>' +
      '<p style="background:#0d0d0f;color:#fff;border-radius:10px;padding:12px 16px;border-left:4px solid #f5a800">' +
      '<b style="color:#f5a800">' + esc_(dateLongue_(p.date)) + (p.heure ? ' · ' + esc_(p.heure) : '') + '</b><br>' +
      esc_(p.endroit || 'Endroit à confirmer') + (p.cout ? '<br>Coût : ' + esc_(p.cout) : '') + (p.notes ? '<br><i>' + esc_(p.notes) + '</i>' : '') + '</p>' +
      '<p>Peux-tu nous dire si ' + esc_(j.prenom) + ' sera là?</p>' +
      bouton('✔ Je serai là', lien('oui'), '#1d7a46') + bouton('✖ Absent', lien('non'), '#7a1426') +
      '<p style="color:#6e6668;font-size:13px;margin-top:18px">' + esc_(c.NOM_CLUB) + ' — ' + esc_(c.SAISON) + '</p></div>'
  };
  const rep = String(c.COURRIEL_NOTIF || '').split(',')[0].trim();
  if (rep) opts.replyTo = rep;
  MailApp.sendEmail(opts);
}

/* ============================================================
 * Gestion (page gérante)
 * ============================================================ */

function gestion_(annee) {
  const groupes = lireGroupes_();
  if (!groupes.length) throw new Error('Aucun groupe configuré.');
  const g = groupes.find(x => x.annee === String(annee)) || groupes[0];
  synchroniser_(g.annee);

  const pratiques = lirePratiques_().filter(p => p.groupe === g.annee);
  const sh = feuillePres_(g.annee);
  const n = derniereLigne_(sh, 1);
  const largeur = sh.getLastColumn();
  const titres = sh.getRange(1, 1, 1, largeur).getValues()[0].map(String);
  const v = n > 1 ? sh.getRange(2, 1, n - 1, largeur).getValues() : [];

  const infos = {};
  joueursGroupe_(g.annee).forEach(j => { infos[j.id] = j; });
  const joueurs = v.filter(r => r[0] && infos[String(r[0])]).map(r => {
    const j = infos[String(r[0])];
    const presences = {};
    pratiques.forEach(p => {
      const ip = titres.indexOf('Présence ' + p.date), ipay = titres.indexOf('Payé ' + p.date);
      presences[p.date] = { p: ip > -1 ? String(r[ip] || '') : '', paye: ipay > -1 ? r[ipay] === true : false };
    });
    return { id: j.id, prenom: j.prenom, nom: j.nom, position: j.position, incomplet: !j.nam,
             statut: j.statut || 'En attente', presences: presences };
  });

  return {
    club: lireParam_().NOM_CLUB,
    groupes: groupes.map(x => ({ annee: x.annee, nom: x.nom,
      enAttente: x.annee === g.annee ? joueurs.filter(j => j.statut !== 'Confirmée').length
                                     : joueursGroupe_(x.annee).filter(j => j.statut !== 'Confirmée').length })),
    groupe: g.annee,
    pratiques: pratiques.map(p => ({ date: p.date, heure: p.heure, endroit: p.endroit, cout: p.cout, notes: p.notes })),
    joueurs: joueurs, prochaine: prochainePratique_(g.annee)
  };
}

function setPresence_(d) {
  const annee = String(d.groupe || ''), date = String(d.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Date de pratique invalide.');
  if (!lirePratiques_().some(p => p.date === date && p.groupe === annee)) throw new Error('Pratique introuvable pour ce groupe.');
  const lock = verrou_();
  try {
    const sh = feuillePres_(annee);
    const r = trouverLignePres_(sh, String(d.id));
    if (!r) throw new Error('Joueur introuvable. Rafraîchis la page.');
    const col = colsPratique_(sh, date);
    if (d.presence !== undefined) {
      const val = String(d.presence || '');
      if (val && PRESENCES.indexOf(val) === -1) throw new Error('Valeur de présence invalide.');
      sh.getRange(r, col.p).setValue(val);
      if (val === 'Absent') sh.getRange(r, col.paye).setValue(false);
    }
    if (d.paye !== undefined) {
      if (d.paye === true && sh.getRange(r, col.p).getValue() === 'Absent') throw new Error('Ce joueur est absent : impossible de le marquer payé.');
      sh.getRange(r, col.paye).setValue(d.paye === true);
    }
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return { ok: true };
}

function ajouterPratique_(d) {
  const annee = String(d.groupe || ''), date = String(d.date || '');
  if (!lireGroupes_().some(g => g.annee === annee)) throw new Error('Groupe invalide.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Date invalide.');
  if (lirePratiques_().some(p => p.date === date && p.groupe === annee)) throw new Error('Ce groupe a déjà une pratique le ' + date + '.');
  const t = (v, m) => sur_(String(v == null ? '' : v).trim().slice(0, m || 200));
  const sh = feuillePrat_(), C = colonnesPrat_(sh);
  const r = derniereLigne_(sh, C.date) + 1;
  const vals = { groupe: annee, date: new Date(date + 'T12:00:00'), heure: t(d.heure, 20), endroit: t(d.endroit), cout: t(d.cout, 20), notes: t(d.notes, 300) };
  Object.keys(vals).forEach(k => sh.getRange(r, C[k]).setValue(vals[k]));
  sh.getRange(r, C.date).setNumberFormat('yyyy-mm-dd');
  colsPratique_(feuillePres_(annee), date);
  return gestion_(annee);
}

/** Fiche d'un joueur (sans NAM). */
function ficheJoueur_(d) {
  const j = trouverParId_(String(d.id));
  if (!j) throw new Error('Joueur introuvable.');
  return {
    id: j.id, groupe: j.annee, prenom: j.prenom, nom: j.nom, parent: j.parent, courriel: j.courriel,
    position: j.position, lance: j.lance, dateNaissance: j.dateNaissance, photos: j.photos,
    statut: j.statut || 'En attente', confirmeeLe: j.confirmeeLe, inscritLe: j.inscritLe, source: j.source,
    namInscrit: !!j.nam
  };
}

function confirmerInscription_(d) {
  const lock = verrou_();
  let j;
  try {
    j = trouverParId_(String(d.id));
    if (!j) throw new Error('Joueur introuvable.');
    const sh = feuilleInscr_(j.annee), C = colonnesInscr_(sh);
    sh.getRange(j.ligne, C.statut).setValue('Confirmée');
    sh.getRange(j.ligne, C.confirmeeLe).setValue(new Date()).setNumberFormat('yyyy-mm-dd');
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }

  const c = lireParam_();
  const g = lireGroupes_().find(x => x.annee === j.annee) || { nom: j.annee, annee: j.annee };
  const dest = adressesValides_(j.courriel);
  let envoye = false;
  if (d.envoyer !== false && dest.length) {
    const p = prochainePratique_(j.annee);
    const opts = {
      to: dest.join(','), name: g.nom,
      subject: 'Inscription confirmée — ' + j.prenom + ' ' + j.nom + ' (' + g.nom + ')',
      htmlBody: '<div style="font-family:Arial,sans-serif;font-size:15px;color:#1a1416;line-height:1.5">' +
        '<p>Bonjour,</p><p>L’inscription de <b>' + esc_(j.prenom + ' ' + j.nom) + '</b> avec les <b>' + esc_(g.nom) + '</b> ' +
        'pour la saison ' + esc_(c.SAISON) + ' est <b>confirmée</b>. Bienvenue!</p>' +
        (p ? '<p style="background:#0d0d0f;color:#fff;border-radius:10px;padding:12px 16px;border-left:4px solid #f5a800">' +
             '<b style="color:#f5a800">Prochaine pratique</b><br>' + esc_(dateLongue_(p.date)) + (p.heure ? ' · ' + esc_(p.heure) : '') +
             '<br>' + esc_(p.endroit || 'Endroit à confirmer') + '</p>' : '') +
        (oui_(c.RAPPEL_ACTIF) ? '<p>Tu recevras un rappel par courriel la veille de chaque pratique.</p>' : '') +
        '<p style="color:#6e6668;font-size:13px">' + esc_(c.NOM_CLUB) + '</p></div>'
    };
    const rep = String(c.COURRIEL_NOTIF || '').split(',')[0].trim();
    if (rep) opts.replyTo = rep;
    MailApp.sendEmail(opts);
    envoye = true;
  }
  journal_('Inscription confirmée', j.id + ' ' + j.prenom + ' ' + j.nom + (envoye ? ' — courriel envoyé' : ''));
  return { statut: 'Confirmée', courriel: envoye ? dest.join(', ') : '' };
}

/** Surclassement : déplace la fiche et la ligne de présences vers un autre groupe. */
function changerGroupe_(d) {
  const nouveau = String(d.nouveau || '');
  if (!lireGroupes_().some(g => g.annee === nouveau)) throw new Error('Groupe invalide.');
  const lock = verrou_();
  try {
    const j = trouverParId_(String(d.id));
    if (!j) throw new Error('Joueur introuvable.');
    if (j.annee === nouveau) return { groupe: nouveau };
    const src = feuilleInscr_(j.annee), CS = colonnesInscr_(src);
    const dst = feuilleInscr_(nouveau), CD = colonnesInscr_(dst);
    const ligne = src.getRange(j.ligne, 1, 1, src.getLastColumn()).getValues()[0];
    const r = derniereLigne_(dst, CD.nom, CD.prenom) + 1;
    CHAMPS.forEach(f => {
      const v = ligne[CS[f.k] - 1];
      if (v !== '' && v != null) { const cell = dst.getRange(r, CD[f.k]); cell.setValue(v); if (v instanceof Date) cell.setNumberFormat('yyyy-mm-dd'); }
    });
    src.deleteRow(j.ligne);
    const ps = feuillePres_(j.annee), rp = trouverLignePres_(ps, j.id);
    if (rp) ps.deleteRow(rp);
    lignePres_(feuillePres_(nouveau), j);
    journal_('Changement de groupe', j.id + ' ' + j.prenom + ' ' + j.nom + ' : ' + j.annee + ' → ' + nouveau);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return { groupe: nouveau };
}

/** NAM pour une urgence — chaque consultation est consignée dans l'onglet Journal. */
function voirNam_(d) {
  const j = trouverParId_(String(d.id));
  if (!j) throw new Error('Joueur introuvable.');
  journal_('Consultation NAM', j.id + ' ' + j.prenom + ' ' + j.nom + (d.motif ? ' — ' + String(d.motif).slice(0, 200) : ''));
  return { nam: j.nam || '' };
}

/* ============================================================
 * Import d'un ancien onglet de pratique (couleurs de la colonne « pmt… »)
 * ============================================================ */

function importerPratique() {
  const ui = SpreadsheetApp.getUi();
  try {
    const src = SpreadsheetApp.getActiveSheet();
    const nomSrc = src.getName();
    if ([F_PRAT, F_PARAM, F_GROUPES, F_JOURNAL].indexOf(nomSrc) > -1 || nomSrc.startsWith(PREF_INSCR) || nomSrc.startsWith(PREF_PRES)) {
      throw new Error('Ouvre d’abord l’onglet de la pratique à importer.');
    }
    const groupes = lireGroupes_();
    const pg = ui.prompt('Importer « ' + nomSrc + ' »', 'Groupe (année) : ' + groupes.map(g => g.annee).join(', ') + '\nLaisse vide pour ' + groupes[0].annee, ui.ButtonSet.OK_CANCEL);
    if (pg.getSelectedButton() !== ui.Button.OK) return;
    const annee = pg.getResponseText().trim() || groupes[0].annee;
    if (!groupes.some(g => g.annee === annee)) throw new Error('Groupe inconnu : ' + annee);

    const m = /(\d{1,2})[-\/. ](\d{1,2})[-\/. ](\d{4})/.exec(nomSrc);
    const devine = m ? m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0') : '';
    const pd = ui.prompt('Importer « ' + nomSrc + ' »', 'Date de la pratique (AAAA-MM-JJ)' + (devine ? '\nLaisse vide pour ' + devine : ''), ui.ButtonSet.OK_CANCEL);
    if (pd.getSelectedButton() !== ui.Button.OK) return;
    const date = pd.getResponseText().trim() || devine;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Date invalide : « ' + date + ' ».');

    const n = src.getLastRow(), largeur = src.getLastColumn();
    if (n < 2) throw new Error('Onglet vide.');
    const titres = src.getRange(1, 1, 1, largeur).getValues()[0].map(norm_);
    const col = test => titres.findIndex(test) + 1;
    const cNom = col(h => h === 'nom'), cPrenom = col(h => h === 'prenom'), cPos = col(h => h === 'position');
    const cCourriel = col(h => h.includes('courriel')), cPmt = col(h => h.startsWith('pmt') || h.startsWith('paiement'));
    if (!cNom || !cPrenom || !cPmt) throw new Error('Colonnes « Nom », « Prénom » ou « pmt… » introuvables dans cet onglet.');
    const valeurs = src.getRange(2, 1, n - 1, largeur).getValues();
    const couleurs = src.getRange(2, cPmt, n - 1, 1).getBackgrounds();

    const lignes = [];
    valeurs.forEach((r, i) => {
      const nom = String(r[cNom - 1] || '').replace(/^\*+/, '').trim();
      const prenom = String(r[cPrenom - 1] || '').replace(/^\*+/, '').trim();
      if (!nom && !prenom) return;
      lignes.push({ nom: nom, prenom: prenom, etat: etatCouleur_(couleurs[i][0]),
        position: cPos ? String(r[cPos - 1] || '').trim() : '', courriel: cCourriel ? String(r[cCourriel - 1] || '').trim() : '' });
    });
    const compte = e => lignes.filter(l => l.etat === e).length;
    const tous = tousJoueurs_();
    const nouveaux = lignes.filter(l => !tous.some(j => memeNom_(j.prenom, j.nom, l.prenom, l.nom)));
    if (ui.alert('Aperçu — ' + annee + ' — pratique du ' + date,
      lignes.length + ' joueurs lus\n\n🟩 Présents et payés : ' + compte('paye') + '\n🟥 Présents, pas payés : ' + compte('present') +
      '\n⬛ Absents : ' + compte('absent') + '\n⬜ Sans couleur (ignorés) : ' + compte('') + '\n\n' +
      (nouveaux.length ? 'Nouveaux joueurs à créer dans ' + annee + ' :\n• ' + nouveaux.map(l => l.prenom + ' ' + l.nom).join('\n• ') + '\n\n' : '') +
      'Importer?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

    const lock = verrou_();
    try {
      if (!lirePratiques_().some(p => p.date === date && p.groupe === annee)) {
        const pr = feuillePrat_(), CP = colonnesPrat_(pr), r = derniereLigne_(pr, CP.date) + 1;
        pr.getRange(r, CP.groupe).setValue(annee);
        pr.getRange(r, CP.date).setValue(new Date(date + 'T12:00:00')).setNumberFormat('yyyy-mm-dd');
        pr.getRange(r, CP.notes).setValue('Importée de « ' + nomSrc + ' »');
        pr.getRange(r, CP.rappel).setValue('—');
      }
      const pres = feuillePres_(annee), cols = colsPratique_(pres, date);
      lignes.forEach(l => {
        let j = trouverParNom_(l.prenom, l.nom);
        if (!j) j = enregistrerJoueur_({ prenom: l.prenom, nom: l.nom, position: l.position, courriel: l.courriel, statut: 'Confirmée' }, 'Import ' + nomSrc, annee);
        else completerFiche_(j, { position: l.position, courriel: l.courriel });
        if (!l.etat || j.annee !== annee) return;
        const r = lignePres_(pres, j);
        pres.getRange(r, cols.p).setValue(l.etat === 'absent' ? 'Absent' : 'Présent');
        pres.getRange(r, cols.paye).setValue(l.etat === 'paye');
      });
      synchroniser_(annee);
      SpreadsheetApp.flush();
    } finally { lock.releaseLock(); }
    ui.alert('Import terminé ✅\nPratique du ' + date + ' (' + annee + ') ajoutée.\nTu peux cacher ou supprimer l’onglet « ' + nomSrc + ' » après vérification.');
  } catch (e) {
    ui.alert('Import impossible : ' + e.message);
  }
}

function etatCouleur_(hex) {
  const h = String(hex || '').replace('#', '');
  if (h.length !== 6) return '';
  const r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
  if (r < 70 && g < 70 && b < 70) return 'absent';
  if (g > 140 && g > r + 50 && g > b + 50) return 'paye';
  if (r > 150 && r > g + 80 && r > b + 80) return 'present';
  return '';
}

/** Remplit seulement les champs vides d'une fiche existante (n'écrase jamais). */
function completerFiche_(j, vals) {
  const sh = feuilleInscr_(j.annee), C = colonnesInscr_(sh);
  Object.keys(vals).forEach(k => {
    if (!vals[k] || !C[k]) return;
    const cell = sh.getRange(j.ligne, C[k]);
    if (!String(cell.getValue()).trim()) cell.setValue(sur_(vals[k]));
  });
}

/* ============================================================
 * Lecture des joueurs (tous les groupes)
 * ============================================================ */

function joueursGroupe_(annee) {
  const sh = classeur_().getSheetByName(PREF_INSCR + annee);
  if (!sh) return [];
  const C = colonnesInscr_(sh);
  const n = derniereLigne_(sh, C.nom, C.prenom);
  if (n < 2) return [];
  const v = sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues();
  const tz = Session.getScriptTimeZone();
  const txt = x => x instanceof Date ? Utilities.formatDate(x, tz, 'yyyy-MM-dd') : String(x == null ? '' : x).trim();
  return v.map((r, i) => {
    const o = { annee: String(annee), ligne: i + 2 };
    CHAMPS.forEach(f => { o[f.k] = txt(r[C[f.k] - 1]); });
    return o;
  }).filter(j => j.prenom || j.nom);
}

function tousJoueurs_() {
  return lireGroupes_().reduce((a, g) => a.concat(joueursGroupe_(g.annee)), []);
}

/** Cherche par nom (accents/majuscules ignorés), dans un groupe ou dans tous; garantit ID et jeton. */
function trouverParNom_(prenom, nom, annee) {
  const chercher = () => (annee ? joueursGroupe_(annee) : tousJoueurs_()).find(x => memeNom_(x.prenom, x.nom, prenom, nom)) || null;
  let j = chercher();
  if (j && (!j.id || !j.jeton)) { synchroniser_(j.annee); j = chercher(); }
  return j;
}

function trouverParId_(id) { return tousJoueurs_().find(j => j.id === id) || null; }
function trouverParJeton_(jeton) { return tousJoueurs_().find(j => j.jeton === jeton) || null; }

/* ============================================================
 * Onglets
 * ============================================================ */

function lireGroupes_() {
  const sh = classeur_().getSheetByName(F_GROUPES);
  if (!sh || sh.getLastRow() < 2) return [];
  // L'année peut être du texte (« 2020 »), un nombre (2020) ou un nombre affiché « 2 020 » selon la langue du Sheet
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .map(r => {
      const annee = (typeof r[0] === 'number' ? String(Math.round(r[0])) : String(r[0])).replace(/\D/g, '');
      return { annee: annee, nom: String(r[1]).trim() || 'Aigles ' + annee, ouvert: !String(r[2]).trim() || oui_(r[2]) };
    })
    .filter(g => /^\d{4}$/.test(g.annee));
}

function feuilleInscr_(annee) {
  const ss = classeur_();
  let sh = ss.getSheetByName(PREF_INSCR + annee);
  if (!sh) {
    sh = ss.insertSheet(PREF_INSCR + annee);
    const t = CHAMPS.map(f => f.t);
    sh.getRange(1, 1, 1, t.length).setValues([t]).setFontWeight('bold').setBackground('#ffff00');
    sh.setFrozenRows(1);
  }
  return sh;
}

function feuillePrat_() {
  const ss = classeur_();
  let pr = ss.getSheetByName(F_PRAT);
  if (!pr) {
    pr = ss.insertSheet(F_PRAT);
    const t = PRAT_CHAMPS.map(f => f.t);
    pr.getRange(1, 1, 1, t.length).setValues([t]).setFontWeight('bold').setBackground('#f5a800');
    pr.setFrozenRows(1);
    pr.setColumnWidth(4, 260); pr.setColumnWidth(6, 300);
  }
  return pr;
}

function colonnesPrat_(sh) { return colonnes_(sh, PRAT_CHAMPS); }
function colonnesInscr_(sh) { return colonnes_(sh, CHAMPS); }

/** Trouve les colonnes par titre; ajoute à droite celles qui manquent (liste dans C._ajouts). */
function colonnes_(sh, champs) {
  const largeur = Math.max(sh.getLastColumn(), 1);
  const titres = sh.getRange(1, 1, 1, largeur).getValues()[0].map(norm_);
  const C = { _ajouts: [] };
  champs.forEach(f => {
    const i = titres.findIndex(h => h && f.m(h));
    if (i > -1) C[f.k] = i + 1;
  });
  let fin = largeur;
  if (!String(sh.getRange(1, largeur).getValue()).trim()) fin = largeur - 1;
  champs.forEach(f => {
    if (!C[f.k]) { fin++; sh.getRange(1, fin).setValue(f.t).setFontWeight('bold'); C[f.k] = fin; C._ajouts.push(f.k); }
  });
  return C;
}

function lirePratiques_() {
  const sh = classeur_().getSheetByName(F_PRAT);
  if (!sh || sh.getLastRow() < 2) return [];
  const C = colonnesPrat_(sh);
  const n = sh.getLastRow() - 1, larg = sh.getLastColumn();
  const v = sh.getRange(2, 1, n, larg).getValues(), aff = sh.getRange(2, 1, n, larg).getDisplayValues();
  const tz = Session.getScriptTimeZone();
  return v.map((r, i) => {
    const brut = r[C.date - 1];
    let date = '';
    if (brut instanceof Date) date = Utilities.formatDate(brut, tz, 'yyyy-MM-dd');
    else if (/^\d{4}-\d{2}-\d{2}$/.test(String(brut).trim())) date = String(brut).trim();
    const a = k => aff[i][C[k] - 1];
    return { ligne: i + 2, groupe: String(a('groupe')).trim(), date: date, heure: a('heure'), endroit: a('endroit'),
             cout: a('cout'), notes: a('notes'), rappel: String(a('rappel')).trim() };
  }).filter(p => p.date && p.groupe).sort((a, b) => a.date.localeCompare(b.date));
}

function prochainePratique_(annee) {
  const auj = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const p = lirePratiques_().filter(x => x.groupe === String(annee) && x.date >= auj)[0];
  return p ? { date: p.date, heure: p.heure, endroit: p.endroit, cout: p.cout, notes: p.notes } : null;
}

function feuillePres_(annee) {
  const ss = classeur_();
  let sh = ss.getSheetByName(PREF_PRES + annee);
  if (!sh) {
    sh = ss.insertSheet(PREF_PRES + annee);
    sh.getRange(1, 1, 1, PRES_FIXES.length).setValues([PRES_FIXES]).setFontWeight('bold').setBackground('#f5a800');
    sh.setFrozenRows(1); sh.setFrozenColumns(3);
    sh.setColumnWidth(1, 60); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 140); sh.setColumnWidth(4, 130);
  }
  return sh;
}

/** Colonnes « Présence AAAA-MM-JJ » et « Payé AAAA-MM-JJ » (créées au besoin, en ordre chronologique). */
function colsPratique_(sh, date) {
  const largeur = Math.max(sh.getLastColumn(), PRES_FIXES.length);
  const titres = sh.getRange(1, 1, 1, largeur).getValues()[0].map(String);
  let p = titres.indexOf('Présence ' + date) + 1, paye = titres.indexOf('Payé ' + date) + 1;
  if (p && paye) return { p: p, paye: paye };

  const apres = titres.findIndex(t => /^Présence \d{4}-\d{2}-\d{2}$/.test(t) && t.slice(9) > date);
  if (apres > -1) { sh.insertColumnsBefore(apres + 1, 2); p = apres + 1; } else p = largeur + 1;
  paye = p + 1;
  sh.getRange(1, p, 1, 2).setValues([['Présence ' + date, 'Payé ' + date]]).setFontWeight('bold').setBackground('#f5a800').setWrap(true);
  sh.setColumnWidth(p, 110); sh.setColumnWidth(paye, 80);
  const max = Math.max(sh.getMaxRows(), 300);
  sh.getRange(2, p, max - 1, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PRESENCES, true).setAllowInvalid(false).build())
    .setHorizontalAlignment('center');
  sh.getRange(2, paye, max - 1, 1).insertCheckboxes();
  appliquerFormats_(sh);
  return { p: p, paye: paye };
}

/**
 * Couleurs de l'onglet Présences : Présent vert · Confirmé jaune · Absent rouge;
 * Payé : payé vert · présent non payé rouge · absent = case masquée.
 * Formules sans séparateur d'arguments : fonctionnent peu importe la langue du Sheet.
 */
function appliquerFormats_(sh) {
  const largeur = sh.getLastColumn();
  if (!largeur) return;
  const titres = sh.getRange(1, 1, 1, largeur).getValues()[0].map(String);
  const max = Math.max(sh.getMaxRows(), 300);
  const R = () => SpreadsheetApp.newConditionalFormatRule();
  const regles = [];
  titres.forEach((t, i) => {
    const m = /^Présence (\d{4}-\d{2}-\d{2})$/.exec(t);
    if (!m) return;
    const cp = i + 1, cpay = titres.indexOf('Payé ' + m[1]) + 1;
    if (!cpay) return;
    const rp = sh.getRange(2, cp, max - 1, 1), rpay = sh.getRange(2, cpay, max - 1, 1);
    const P = colLettre_(cp) + '2', Y = colLettre_(cpay) + '2';
    [['Présent', '#d4edda'], ['Confirmé', '#fff3cd'], ['Absent', '#f8d7da']].forEach(([txt, bg]) => {
      regles.push(R().whenTextEqualTo(txt).setBackground(bg).setRanges([rp]).build());
    });
    regles.push(R().whenFormulaSatisfied('=' + P + '="Absent"').setBackground('#f3f3f3').setFontColor('#f3f3f3').setRanges([rpay]).build());
    regles.push(R().whenFormulaSatisfied('=' + Y + '=TRUE').setBackground('#d4edda').setRanges([rpay]).build());
    regles.push(R().whenFormulaSatisfied('=(' + P + '="Présent")*(' + Y + '<>TRUE)').setBackground('#f8d7da').setRanges([rpay]).build());
  });
  sh.setConditionalFormatRules(regles);
}

/** Modification directe dans un onglet Présences : « Absent » décoche Payé. */
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (!sh.getName().startsWith(PREF_PRES) || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1 || e.value !== 'Absent') return;
    const m = /^Présence (\d{4}-\d{2}-\d{2})$/.exec(String(sh.getRange(1, e.range.getColumn()).getValue()));
    if (!m) return;
    const titres = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const c = titres.indexOf('Payé ' + m[1]) + 1;
    if (c) sh.getRange(e.range.getRow(), c).setValue(false);
  } catch (err) { console.error(err); }
}

function trouverLignePres_(sh, id) {
  const n = derniereLigne_(sh, 1);
  if (n < 2) return 0;
  const v = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]) === String(id)) return i + 2;
  return 0;
}

/** Ligne du joueur dans Présences (créée au besoin; nom/position tenus à jour). */
function lignePres_(sh, j) {
  let r = trouverLignePres_(sh, j.id);
  if (!r) {
    r = derniereLigne_(sh, 1) + 1;
    sh.getRange(r, 1, 1, 4).setValues([[j.id, sur_(j.prenom), sur_(j.nom), sur_(j.position || '')]]);
  } else if (j.position) {
    sh.getRange(r, 2, 1, 3).setValues([[sur_(j.prenom), sur_(j.nom), sur_(j.position)]]);
  }
  return r;
}

/**
 * Pour un groupe : chaque joueur reçoit ID et jeton, et une ligne dans Présences;
 * chaque pratique du groupe a ses colonnes. Retourne le nombre de joueurs.
 */
function synchroniser_(annee) {
  const ins = feuilleInscr_(annee), C = colonnesInscr_(ins);
  const sh = feuillePres_(annee);
  const n = derniereLigne_(ins, C.nom, C.prenom);
  let total = 0;
  if (n > 1) {
    const v = ins.getRange(2, 1, n - 1, ins.getLastColumn()).getValues();
    const ids = tousJoueurs_().map(j => j.id);
    const der = derniereLigne_(sh, 1);
    const existants = {};
    if (der > 1) sh.getRange(2, 1, der - 1, 4).getValues().forEach((r, i) => { existants[String(r[0])] = { r: r, ligne: i + 2 }; });
    const ajouts = [];
    v.forEach((r, i) => {
      const prenom = String(r[C.prenom - 1] || '').trim(), nom = String(r[C.nom - 1] || '').trim();
      if (!prenom && !nom) return;
      total++;
      let id = String(r[C.id - 1] || '');
      if (!id) { id = prochainId_(ids); ids.push(id); ins.getRange(i + 2, C.id).setValue(id); }
      if (!String(r[C.jeton - 1] || '').trim()) ins.getRange(i + 2, C.jeton).setValue(Utilities.getUuid());
      const pos = String(r[C.position - 1] || '');
      const ex = existants[id];
      if (!ex) ajouts.push([id, sur_(prenom), sur_(nom), sur_(pos)]);
      else if (ex.r[1] !== prenom || ex.r[2] !== nom || ex.r[3] !== pos) sh.getRange(ex.ligne, 2, 1, 3).setValues([[sur_(prenom), sur_(nom), sur_(pos)]]);
    });
    if (ajouts.length) sh.getRange(derniereLigne_(sh, 1) + 1, 1, ajouts.length, 4).setValues(ajouts);
  }
  lirePratiques_().filter(p => p.groupe === String(annee)).forEach(p => colsPratique_(sh, p.date));
  appliquerFormats_(sh);
  return total;
}

/* ============================================================
 * Utilitaires
 * ============================================================ */

/**
 * Le Sheet auquel ce script est attaché (fonctionne aussi pour l'API / application Web).
 * On n'utilise plus un identifiant mémorisé : il pouvait pointer vers un autre Sheet (ex. un Sheet de test).
 */
function classeur_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (id) return SpreadsheetApp.openById(id);
  throw new Error('Google Sheet introuvable : ce script doit être ouvert depuis Extensions > Apps Script du Sheet.');
}

function lireParam_() {
  const o = {};
  PARAM_DEFAUT.forEach(r => { o[r[0]] = String(r[1]); });
  const sh = classeur_().getSheetByName(F_PARAM);
  if (sh) sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) o[String(r[0]).trim()] = String(r[1]).trim(); });
  return o;
}

function verrou_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Le système est occupé, réessaie dans quelques secondes.');
  return lock;
}

function journal_(action, details) {
  try {
    const ss = classeur_();
    let sh = ss.getSheetByName(F_JOURNAL);
    if (!sh) {
      sh = ss.insertSheet(F_JOURNAL);
      sh.getRange(1, 1, 1, 3).setValues([['Date', 'Action', 'Détails']]).setFontWeight('bold');
      sh.setFrozenRows(1); sh.setColumnWidth(3, 520);
    }
    sh.appendRow([new Date(), action, sur_(details)]);
  } catch (e) { console.error('Journal', e); }
}

/** Code gérante (bloque 15 min après 10 échecs). */
function verifierCode_(code) {
  const c = lireParam_();
  const cache = CacheService.getScriptCache();
  const echecs = Number(cache.get('echecs_code') || 0);
  if (echecs >= 10) throw new Error('Trop d’essais avec un mauvais code. Réessaie dans 15 minutes.');
  if (!c.CODE_GERANTE || String(code || '').trim() !== String(c.CODE_GERANTE).trim()) {
    cache.put('echecs_code', String(echecs + 1), 900);
    throw new Error('Code gérante invalide.');
  }
}

/**
 * Dernière ligne non vide dans les colonnes données (getLastRow() ne convient pas :
 * les cases à cocher contiennent FALSE et comptent comme du contenu).
 */
function derniereLigne_(sh, ...cols) {
  const n = sh.getLastRow();
  if (n < 2) return 1;
  const vals = cols.filter(Boolean).map(c => sh.getRange(1, c, n, 1).getValues());
  for (let i = n - 1; i >= 1; i--) if (vals.some(v => String(v[i][0]).trim())) return i + 1;
  return 1;
}

function prochainId_(ids) {
  let max = 0;
  ids.forEach(x => { const m = /^J(\d+)$/.exec(String(x || '')); if (m) max = Math.max(max, Number(m[1])); });
  return 'J' + String(max + 1).padStart(3, '0');
}

/** NAM : 4 lettres + 8 chiffres → « ABCD 1234 5678 ». */
function normNam_(s) {
  const x = String(s || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  return /^[A-Z]{4}\d{8}$/.test(x) ? x.slice(0, 4) + ' ' + x.slice(4, 8) + ' ' + x.slice(8) : x;
}

function adressesValides_(s) { return String(s || '').split(/[;,\s]+/).map(x => x.trim().toLowerCase()).filter(courrielValide_); }
function courrielValide_(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(s || '')); }
function memeNom_(p1, n1, p2, n2) { return norm_(p1) === norm_(p2) && norm_(n1) === norm_(n2) && norm_(p1) !== ''; }
function oui_(v) { return /^(oui|o|yes|true|1)$/i.test(String(v).trim()); }

function dateLongue_(iso) {
  const [y, m, d] = iso.split('-').map(Number);
  const jours = ['dimanche', 'lundi', 'mardi', 'mercredi', 'jeudi', 'vendredi', 'samedi'];
  const mois = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
  const dt = new Date(y, m - 1, d, 12);
  const s = jours[dt.getDay()] + ' ' + d + ' ' + mois[m - 1];
  return s.charAt(0).toUpperCase() + s.slice(1);
}

function colLettre_(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function norm_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[‐-―]/g, '-').replace(/\s+/g, ' ').trim();
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
}

/** Empêche l'injection de formules (texte qui commence par = + - @). */
function sur_(v) { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; }
