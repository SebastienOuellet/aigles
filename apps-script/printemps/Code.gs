/**
 * Printemps 2027 — Aigles 2019
 * Formulaire d'inscription / confirmation de présence + gestion des pratiques
 *
 * À coller dans Extensions > Apps Script du Google Sheet
 * « Liste présence camps Aigles 2019 - saison printemps 2027 ».
 */

const F_INSCR = 'Inscriptions';
const F_PRAT = 'Pratiques';
const F_PRES = 'Présences';
const F_PARAM = 'Paramètres';

const PARAM_DEFAUT = [
  ['NOM_EQUIPE', 'Aigles 2019', 'Nom affiché sur le formulaire'],
  ['SAISON', 'Printemps 2027', 'Nom de la saison'],
  ['INSCRIPTIONS_OUVERTES', 'OUI', 'OUI = nouvelles inscriptions acceptées | NON = seulement les confirmations de présence'],
  ['CODE_GERANTE', '', 'OBLIGATOIRE — code de la page de gestion (ne pas partager)'],
  ['COURRIEL_NOTIF', '', 'Courriel de la gérante — reçoit chaque nouvelle inscription (plusieurs adresses : séparer par des virgules)'],
  ['LIEN_SITE', 'https://sebastienouellet.github.io/aigles/printemps/', 'Adresse du formulaire (la gestion = LIEN_SITE + gerante.html)']
];

const PRAT_ENTETES = ['Date', 'Heure', 'Endroit', 'Coût', 'Notes'];
const PRES_FIXES = ['ID', 'Prénom', 'Nom', 'Position'];
const PRESENCES = ['Confirmé', 'Présent', 'Absent'];

/**
 * Colonnes de l'onglet Inscriptions, retrouvées par leur titre (les tiens restent tels quels).
 * occ = n-ième colonne qui correspond (il y a deux colonnes « Grandeur »).
 */
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
  { k: 'source', t: 'Source', m: h => h === 'source' }
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
    .addItem('Configurer (première fois)', 'setup')
    .addItem('Synchroniser les présences', 'synchroniserMenu')
    .addItem('Importer une ancienne pratique (onglet ouvert)', 'importerPratique')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());

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

  feuillePrat_();

  // Inscriptions : colonnes manquantes
  const ins = ss.getSheetByName(F_INSCR);
  if (!ins) throw new Error('Onglet « ' + F_INSCR + ' » introuvable.');
  colonnesInscr_(ins);

  feuillePres_();
  const n = synchroniser_();

  const c = lireParam_();
  const todo = [];
  if (!c.CODE_GERANTE) todo.push('Choisir un CODE_GERANTE dans l’onglet Paramètres.');
  if (!c.COURRIEL_NOTIF) todo.push('Inscrire le courriel de la gérante (COURRIEL_NOTIF) dans Paramètres.');
  if (!lirePratiques_().length) todo.push('Ajouter au moins une pratique (onglet Pratiques ou page de gestion).');
  SpreadsheetApp.getUi().alert('Configuration terminée ✅\n\n' + n + ' joueur(s) dans Présences.' +
    (todo.length ? '\n\nÀ faire :\n• ' + todo.join('\n• ') : ''));
}

/**
 * Importe un ancien onglet de pratique (ex. « Pratique du 26-09-2026 ») d'après les couleurs
 * de la colonne « pmt… » : vert = présent + payé, rouge = présent non payé, noir = absent.
 */
function importerPratique() {
  const ui = SpreadsheetApp.getUi();
  try {
    const src = SpreadsheetApp.getActiveSheet();
    if ([F_INSCR, F_PRAT, F_PRES, F_PARAM].indexOf(src.getName()) > -1) throw new Error('Ouvre d’abord l’onglet de la pratique à importer.');

    // 1. Date
    const m = /(\d{1,2})[-\/. ](\d{1,2})[-\/. ](\d{4})/.exec(src.getName());
    const devine = m ? m[3] + '-' + m[2].padStart(2, '0') + '-' + m[1].padStart(2, '0') : '';
    const pd = ui.prompt('Importer « ' + src.getName() + ' »', 'Date de la pratique (AAAA-MM-JJ)' + (devine ? '\nLaisse vide pour ' + devine : ''), ui.ButtonSet.OK_CANCEL);
    if (pd.getSelectedButton() !== ui.Button.OK) return;
    const date = pd.getResponseText().trim() || devine;
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Date invalide : « ' + date + ' ».');

    // 2. Colonnes de l'onglet source
    const n = src.getLastRow(), largeur = src.getLastColumn();
    if (n < 2) throw new Error('Onglet vide.');
    const titres = src.getRange(1, 1, 1, largeur).getValues()[0].map(norm_);
    const col = test => titres.findIndex(test) + 1;
    const cNom = col(h => h === 'nom'), cPrenom = col(h => h === 'prenom'), cPos = col(h => h === 'position');
    const cCourriel = col(h => h.includes('courriel')), cPmt = col(h => h.startsWith('pmt') || h.startsWith('paiement'));
    if (!cNom || !cPrenom || !cPmt) throw new Error('Colonnes « Nom », « Prénom » ou « pmt… » introuvables dans cet onglet.');

    const valeurs = src.getRange(2, 1, n - 1, largeur).getValues();
    const couleurs = src.getRange(2, cPmt, n - 1, 1).getBackgrounds();

    // 3. Lecture + aperçu
    const lignes = [];
    valeurs.forEach((r, i) => {
      const nom = String(r[cNom - 1] || '').replace(/^\*+/, '').trim();
      const prenom = String(r[cPrenom - 1] || '').replace(/^\*+/, '').trim();
      if (!nom && !prenom) return;
      lignes.push({
        nom: nom, prenom: prenom, etat: etatCouleur_(couleurs[i][0]),
        position: cPos ? String(r[cPos - 1] || '').trim() : '',
        courriel: cCourriel ? String(r[cCourriel - 1] || '').trim() : ''
      });
    });
    const compte = e => lignes.filter(l => l.etat === e).length;
    const nouveaux = lignes.filter(l => !trouverParNomSansCreer_(l.prenom, l.nom));
    const ok = ui.alert('Aperçu — pratique du ' + date,
      lignes.length + ' joueurs lus\n\n' +
      '🟩 Présents et payés : ' + compte('paye') + '\n' +
      '🟥 Présents, pas payés : ' + compte('present') + '\n' +
      '⬛ Absents : ' + compte('absent') + '\n' +
      '⬜ Sans couleur (ignorés) : ' + compte('') + '\n\n' +
      (nouveaux.length ? 'Nouveaux joueurs à créer (fiche incomplète) :\n• ' + nouveaux.map(l => l.prenom + ' ' + l.nom).join('\n• ') + '\n\n' : '') +
      'Importer?', ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;

    // 4. Écriture
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(30000)) throw new Error('Occupé, réessaie.');
    try {
      if (!lirePratiques_().some(p => p.date === date)) {
        const pr = feuillePrat_();
        pr.appendRow([new Date(date + 'T12:00:00'), '', '', '', 'Importée de « ' + src.getName() + ' »']);
        pr.getRange(pr.getLastRow(), 1).setNumberFormat('yyyy-mm-dd');
        pr.getRange(2, 1, pr.getLastRow() - 1, pr.getLastColumn()).sort(1);
      }
      const pres = feuillePres_();
      const cols = colsPratique_(pres, date);
      const ins = classeur_().getSheetByName(F_INSCR);
      const C = colonnesInscr_(ins);
      lignes.forEach(l => {
        let j = trouverParNom_(l.prenom, l.nom);
        if (!j) j = enregistrerJoueur_({ prenom: l.prenom, nom: l.nom, position: l.position, courriel: l.courriel }, 'Import ' + src.getName());
        else completerFiche_(ins, C, j.id, { position: l.position, courriel: l.courriel });
        if (!l.etat) return;
        const r = lignePres_(pres, j);
        pres.getRange(r, cols.p).setValue(l.etat === 'absent' ? 'Absent' : 'Présent');
        pres.getRange(r, cols.paye).setValue(l.etat === 'paye');
      });
      synchroniser_();
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    ui.alert('Import terminé ✅\nPratique du ' + date + ' ajoutée dans Pratiques et Présences.\n' +
      'Tu peux maintenant cacher ou supprimer l’onglet « ' + src.getName() + ' » quand tu as vérifié.');
  } catch (e) {
    ui.alert('Import impossible : ' + e.message);
  }
}

/** Vert → payé, rouge → présent non payé, noir → absent, autre → ''. */
function etatCouleur_(hex) {
  const h = String(hex || '').replace('#', '');
  if (h.length !== 6) return '';
  const r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
  if (r < 70 && g < 70 && b < 70) return 'absent';
  if (g > 140 && g > r + 50 && g > b + 50) return 'paye';
  if (r > 150 && r > g + 80 && r > b + 80) return 'present';
  return '';
}

function trouverParNomSansCreer_(prenom, nom) {
  const sh = classeur_().getSheetByName(F_INSCR);
  const C = colonnesInscr_(sh);
  const n = sh.getLastRow();
  if (n < 2) return false;
  return sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues()
    .some(r => memeNom_(r[C.prenom - 1], r[C.nom - 1], prenom, nom));
}

/** Remplit seulement les champs vides d'une fiche existante (n'écrase jamais). */
function completerFiche_(sh, C, id, vals) {
  const n = sh.getLastRow();
  if (n < 2) return;
  const ids = sh.getRange(2, C.id, n - 1, 1).getValues();
  const i = ids.findIndex(r => String(r[0]) === String(id));
  if (i < 0) return;
  Object.keys(vals).forEach(k => {
    if (!vals[k] || !C[k]) return;
    const cell = sh.getRange(i + 2, C[k]);
    if (!String(cell.getValue()).trim()) cell.setValue(sur_(vals[k]));
  });
}

function synchroniserMenu() {
  const n = synchroniser_();
  SpreadsheetApp.getUi().alert('Présences synchronisées : ' + n + ' joueur(s).');
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
 * API (appelée par le site)
 * POST text/plain {action, code?, data?}
 * ============================================================ */

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const d = req.data || {};
    switch (req.action) {
      case 'config': out = configPublique_(); break;
      case 'inscrire': out = inscrire_(d); break;
      case 'confirmer': out = confirmer_(d); break;
      case 'gestion': verifierCode_(req.code); out = gestion_(); break;
      case 'setPresence': verifierCode_(req.code); out = setPresence_(d); break;
      case 'ajouterPratique': verifierCode_(req.code); out = ajouterPratique_(d); break;
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

function configPublique_() {
  const c = lireParam_();
  return {
    equipe: c.NOM_EQUIPE, saison: c.SAISON,
    ouvert: oui_(c.INSCRIPTIONS_OUVERTES),
    prochaine: prochainePratique_(),
    choix: CHOIX
  };
}

/* ============================================================
 * Inscription et confirmation
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
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(o.courriel)) err.push('Courriel invalide.');
  ['position', 'lance', 'photos'].forEach(k => {
    if (CHOIX[k].indexOf(o[k]) === -1) err.push('Choix invalide : ' + k + '.');
  });
  if (!/^\d{4}-\d{2}-\d{2}$/.test(o.dateNaissance)) err.push('Date de naissance invalide.');
  if (!/^[A-Z]{4} \d{4} \d{4}$/.test(o.nam)) err.push('NAM invalide (format : ABCD 1234 5678).');
  if (err.length) throw new Error(err.join(' '));

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Le système est occupé, réessaie dans quelques secondes.');
  let joueur;
  try {
    joueur = enregistrerJoueur_(o, 'Formulaire');
    if (d.presentProchaine) confirmerPresenceProchaine_(joueur);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  notifier_(o, c, !!d.presentProchaine);
  return { prenom: o.prenom, nom: o.nom, prochaine: prochainePratique_(), confirme: !!d.presentProchaine };
}

/** Joueur qui a déjà pratiqué : on confirme sa présence à la prochaine pratique. */
function confirmer_(d) {
  if (d.website) throw new Error('Soumission refusée.');
  const prenom = String(d.prenom || '').trim().slice(0, 60);
  const nom = String(d.nom || '').trim().slice(0, 60);
  if (!prenom || !nom) throw new Error('Prénom et nom du joueur requis.');
  const prochaine = prochainePratique_();
  if (!prochaine) throw new Error('Aucune pratique à venir n’est encore prévue. Reviens bientôt!');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Le système est occupé, réessaie dans quelques secondes.');
  let joueur, nouveau = false;
  try {
    joueur = trouverParNom_(prenom, nom);
    if (!joueur) {
      // Fiche minimale : la gérante verra « fiche incomplète »
      joueur = enregistrerJoueur_({ prenom: prenom, nom: nom }, 'Confirmation (fiche à compléter)');
      nouveau = true;
    }
    confirmerPresenceProchaine_(joueur);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { prenom: joueur.prenom, nom: joueur.nom, prochaine: prochaine, ficheIncomplete: nouveau || !joueur.nam };
}

/**
 * Ajoute un joueur dans Inscriptions (ou complète une fiche minimale existante) et dans Présences.
 * Doublon : même NAM = refus.
 */
function enregistrerJoueur_(o, source) {
  const sh = classeur_().getSheetByName(F_INSCR);
  const C = colonnesInscr_(sh);
  const n = sh.getLastRow();
  const v = n > 1 ? sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues() : [];

  let cible = 0;
  if (o.nam) {
    for (let i = 0; i < v.length; i++) {
      if (normNam_(v[i][C.nam - 1]) === o.nam) {
        throw new Error('Ce joueur est déjà inscrit (NAM déjà enregistré). Pour confirmer ta présence à une pratique, ' +
          'réponds « Oui » à la première question.');
      }
    }
    // Fiche minimale au même nom (créée par une confirmation) → on la complète
    for (let i = 0; i < v.length; i++) {
      if (!String(v[i][C.nam - 1]).trim() && memeNom_(v[i][C.prenom - 1], v[i][C.nom - 1], o.prenom, o.nom)) { cible = i + 2; break; }
    }
  }

  let id = cible ? String(v[cible - 2][C.id - 1] || '') : '';
  if (!id) id = prochainId_(v.map(r => r[C.id - 1]));
  const r = cible || n + 1;

  const set = (k, val, fmt) => { if (C[k] && val !== undefined && val !== '') { const g = sh.getRange(r, C[k]); g.setValue(val); if (fmt) g.setNumberFormat(fmt); } };
  Object.keys(o).forEach(k => { if (C[k] && k !== 'dateNaissance') set(k, sur_(o[k])); });
  if (o.dateNaissance) set('dateNaissance', new Date(o.dateNaissance + 'T12:00:00'), 'yyyy-mm-dd');
  set('id', id);
  set('inscritLe', new Date(), 'yyyy-mm-dd hh:mm');
  set('source', source);

  const joueur = { id: id, prenom: o.prenom, nom: o.nom, position: o.position || '', nam: o.nam || '' };
  lignePres_(feuillePres_(), joueur);
  return joueur;
}

function trouverParNom_(prenom, nom) {
  const sh = classeur_().getSheetByName(F_INSCR);
  const C = colonnesInscr_(sh);
  const n = sh.getLastRow();
  if (n < 2) return null;
  const v = sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues();
  const ids = v.map(r => r[C.id - 1]);
  for (let i = 0; i < v.length; i++) {
    if (memeNom_(v[i][C.prenom - 1], v[i][C.nom - 1], prenom, nom)) {
      let id = String(v[i][C.id - 1] || '');
      if (!id) { id = prochainId_(ids); sh.getRange(i + 2, C.id).setValue(id); ids.push(id); }
      const j = { id: id, prenom: String(v[i][C.prenom - 1]).trim(), nom: String(v[i][C.nom - 1]).trim(),
                  position: String(v[i][C.position - 1] || ''), nam: String(v[i][C.nam - 1] || '').trim() };
      lignePres_(feuillePres_(), j);
      return j;
    }
  }
  return null;
}

function confirmerPresenceProchaine_(joueur) {
  const p = prochainePratique_();
  if (!p) return;
  const sh = feuillePres_();
  const r = lignePres_(sh, joueur);
  const col = colsPratique_(sh, p.date);
  const cell = sh.getRange(r, col.p);
  if (cell.getValue() !== 'Présent') cell.setValue('Confirmé');
}

/** Courriel à la gérante pour chaque première inscription (le NAM n'est jamais envoyé par courriel). */
function notifier_(o, c, presentProchaine) {
  if (!c.COURRIEL_NOTIF) return;
  const p = prochainePratique_();
  const ligne = (k, v) => '<tr><td style="color:#666;padding:4px 14px 4px 0;vertical-align:top">' + k + '</td><td style="padding:4px 0"><b>' + v + '</b></td></tr>';
  try {
    MailApp.sendEmail({
      to: c.COURRIEL_NOTIF,
      replyTo: o.courriel,
      name: c.NOM_EQUIPE + ' — Inscriptions',
      subject: '[' + c.NOM_EQUIPE + '] Nouvelle inscription — ' + o.prenom + ' ' + o.nom,
      htmlBody:
        '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1a1416">' +
        '<p>Nouvelle inscription reçue pour la saison <b>' + esc_(c.SAISON) + '</b>. ' +
        'Le parent a été avisé que tu confirmeras l’inscription dès que possible.</p>' +
        '<table style="border-collapse:collapse">' +
        ligne('Joueur', esc_(o.prenom + ' ' + o.nom)) +
        ligne('Date de naissance', esc_(o.dateNaissance)) +
        ligne('Position', esc_(o.position)) +
        ligne('Lance', esc_(o.lance)) +
        ligne('Parent', esc_(o.parent)) +
        ligne('Courriel', '<a href="mailto:' + esc_(o.courriel) + '">' + esc_(o.courriel) + '</a>') +
        ligne('Photos promo', esc_(o.photos)) +
        ligne('NAM', 'reçu ✓ (voir le Google Sheet)') +
        (p ? ligne('Prochaine pratique', esc_(p.date) + (presentProchaine ? ' — présence confirmée' : ' — pas confirmé')) : '') +
        '</table>' +
        '<p style="margin-top:18px"><a href="' + c.LIEN_SITE + 'gerante.html" style="color:#7a1426;font-weight:bold">Ouvrir la gestion →</a> ' +
        '&nbsp;·&nbsp;</p></div>'
    });
  } catch (e) { console.error('Notification', e); }
}

/* ============================================================
 * Gestion (page gérante)
 * ============================================================ */

function gestion_() {
  synchroniser_();
  const pratiques = lirePratiques_();
  const sh = feuillePres_();
  const n = derniereLigneId_(sh);
  const largeur = sh.getLastColumn();
  const titres = sh.getRange(1, 1, 1, largeur).getValues()[0].map(String);
  const v = n > 1 ? sh.getRange(2, 1, n - 1, largeur).getValues() : [];

  // Fiches incomplètes (sans NAM) dans Inscriptions
  const ins = classeur_().getSheetByName(F_INSCR);
  const C = colonnesInscr_(ins);
  const incomplets = {};
  if (ins.getLastRow() > 1) {
    ins.getRange(2, 1, ins.getLastRow() - 1, ins.getLastColumn()).getValues().forEach(r => {
      if (r[C.id - 1] && !String(r[C.nam - 1]).trim()) incomplets[String(r[C.id - 1])] = true;
    });
  }

  const joueurs = v.filter(r => r[0] && (r[1] || r[2])).map(r => {
    const presences = {};
    pratiques.forEach(p => {
      const ip = titres.indexOf('Présence ' + p.date), ipay = titres.indexOf('Payé ' + p.date);
      presences[p.date] = { p: ip > -1 ? String(r[ip] || '') : '', paye: ipay > -1 ? r[ipay] === true : false };
    });
    return { id: String(r[0]), prenom: String(r[1]), nom: String(r[2]), position: String(r[3] || ''),
             incomplet: !!incomplets[String(r[0])], presences: presences };
  }).sort((a, b) => (a.nom + a.prenom).localeCompare(b.nom + b.prenom, 'fr'));

  return { equipe: lireParam_().NOM_EQUIPE, pratiques: pratiques, joueurs: joueurs, prochaine: prochainePratique_() };
}

function setPresence_(d) {
  const date = String(d.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Date de pratique invalide.');
  if (!lirePratiques_().some(p => p.date === date)) throw new Error('Pratique introuvable.');
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Occupé, réessaie.');
  try {
    const sh = feuillePres_();
    const r = trouverLignePres_(sh, String(d.id));
    if (!r) throw new Error('Joueur introuvable. Rafraîchis la page.');
    const col = colsPratique_(sh, date);
    if (d.presence !== undefined) {
      const val = String(d.presence || '');
      if (val && PRESENCES.indexOf(val) === -1) throw new Error('Valeur de présence invalide.');
      sh.getRange(r, col.p).setValue(val);
      if (val === 'Absent') sh.getRange(r, col.paye).setValue(false); // absent = pas de paiement
    }
    if (d.paye !== undefined) {
      if (d.paye === true && sh.getRange(r, col.p).getValue() === 'Absent') throw new Error('Ce joueur est absent : impossible de le marquer payé.');
      sh.getRange(r, col.paye).setValue(d.paye === true);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

function ajouterPratique_(d) {
  const date = String(d.date || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('Date invalide.');
  if (lirePratiques_().some(p => p.date === date)) throw new Error('Une pratique existe déjà le ' + date + '.');
  const t = (v, m) => String(v == null ? '' : v).trim().slice(0, m || 200);
  const sh = feuillePrat_();
  sh.appendRow([new Date(date + 'T12:00:00'), t(d.heure, 20), sur_(t(d.endroit)), sur_(t(d.cout, 20)), sur_(t(d.notes, 300))]);
  sh.getRange(sh.getLastRow(), 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(sh.getLastRow(), 2).setNumberFormat('@');
  sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).sort(1);
  colsPratique_(feuillePres_(), date);
  return gestion_();
}

/* ============================================================
 * Onglets Pratiques et Présences
 * ============================================================ */

function lirePratiques_() {
  const sh = classeur_().getSheetByName(F_PRAT);
  if (!sh || sh.getLastRow() < 2) return [];
  const rg = sh.getRange(2, 1, sh.getLastRow() - 1, PRAT_ENTETES.length);
  const v = rg.getValues(), aff = rg.getDisplayValues();
  const tz = Session.getScriptTimeZone();
  return v.map((r, i) => {
    let date = '';
    if (r[0] instanceof Date) date = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
    else if (/^\d{4}-\d{2}-\d{2}$/.test(String(r[0]).trim())) date = String(r[0]).trim();
    return { date: date, heure: aff[i][1], endroit: aff[i][2], cout: aff[i][3], notes: aff[i][4] };
  }).filter(p => p.date).sort((a, b) => a.date.localeCompare(b.date));
}

function prochainePratique_() {
  const auj = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return lirePratiques_().filter(p => p.date >= auj)[0] || null;
}

function feuillePrat_() {
  const ss = classeur_();
  let pr = ss.getSheetByName(F_PRAT);
  if (!pr) {
    pr = ss.insertSheet(F_PRAT);
    pr.getRange(1, 1, 1, PRAT_ENTETES.length).setValues([PRAT_ENTETES]).setFontWeight('bold').setBackground('#f5a800');
    pr.setFrozenRows(1);
    pr.getRange('A2:A').setNumberFormat('yyyy-mm-dd');
    pr.getRange('B2:B').setNumberFormat('@');
    pr.setColumnWidth(3, 260); pr.setColumnWidth(5, 300);
  }
  return pr;
}

function feuillePres_() {
  const ss = classeur_();
  let sh = ss.getSheetByName(F_PRES);
  if (!sh) {
    sh = ss.insertSheet(F_PRES);
    sh.getRange(1, 1, 1, PRES_FIXES.length).setValues([PRES_FIXES]).setFontWeight('bold').setBackground('#f5a800');
    sh.setFrozenRows(1); sh.setFrozenColumns(3);
    sh.setColumnWidth(1, 60); sh.setColumnWidth(2, 120); sh.setColumnWidth(3, 140); sh.setColumnWidth(4, 130);
  }
  return sh;
}

/** Colonnes « Présence AAAA-MM-JJ » et « Payé AAAA-MM-JJ » (créées au besoin, avec listes et couleurs). */
function colsPratique_(sh, date) {
  const largeur = Math.max(sh.getLastColumn(), PRES_FIXES.length);
  const titres = sh.getRange(1, 1, 1, largeur).getValues()[0].map(String);
  let p = titres.indexOf('Présence ' + date) + 1, paye = titres.indexOf('Payé ' + date) + 1;
  if (p && paye) return { p: p, paye: paye };

  // Insère en ordre chronologique : avant la première pratique plus tardive
  const apres = titres.findIndex(t => /^Présence \d{4}-\d{2}-\d{2}$/.test(t) && t.slice(9) > date);
  if (apres > -1) { sh.insertColumnsBefore(apres + 1, 2); p = apres + 1; }
  else p = largeur + 1;
  paye = p + 1;
  sh.getRange(1, p, 1, 2).setValues([['Présence ' + date, 'Payé ' + date]]).setFontWeight('bold').setBackground('#f5a800').setWrap(true);
  sh.setColumnWidth(p, 110); sh.setColumnWidth(paye, 80);
  const max = Math.max(sh.getMaxRows(), 300);
  const rp = sh.getRange(2, p, max - 1, 1), rpay = sh.getRange(2, paye, max - 1, 1);
  rp.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(PRESENCES, true).setAllowInvalid(false).build())
    .setHorizontalAlignment('center');
  rpay.insertCheckboxes();

  appliquerFormats_(sh);
  return { p: p, paye: paye };
}

/**
 * Dernière ligne avec un ID en colonne A. (getLastRow() ne convient pas : les cases à cocher
 * contiennent FALSE et comptent comme du contenu jusqu'au bas de la feuille.)
 */
function derniereLigneId_(sh) {
  const n = sh.getLastRow();
  if (n < 2) return 1;
  const v = sh.getRange(1, 1, n, 1).getValues();
  for (let i = n - 1; i >= 1; i--) if (String(v[i][0]).trim()) return i + 1;
  return 1;
}

/**
 * (Re)construit les couleurs de l'onglet Présences pour toutes les pratiques :
 *  Présence : Présent vert · Confirmé jaune · Absent rouge
 *  Payé     : payé vert · présent non payé rouge · absent = case masquée (texte et fond gris pâle)
 * Formules sans virgule ni point-virgule : fonctionnent peu importe la langue du Sheet.
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

function colLettre_(n) {
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** Modification directe dans le Sheet : « Absent » décoche Payé. */
function onEdit(e) {
  try {
    const sh = e.range.getSheet();
    if (sh.getName() !== F_PRES || e.range.getNumRows() !== 1 || e.range.getNumColumns() !== 1 || e.value !== 'Absent') return;
    const m = /^Présence (\d{4}-\d{2}-\d{2})$/.exec(String(sh.getRange(1, e.range.getColumn()).getValue()));
    if (!m) return;
    const titres = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    const c = titres.indexOf('Payé ' + m[1]) + 1;
    if (c) sh.getRange(e.range.getRow(), c).setValue(false);
  } catch (err) { console.error(err); }
}

function trouverLignePres_(sh, id) {
  const n = derniereLigneId_(sh);
  if (n < 2) return 0;
  const v = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]) === String(id)) return i + 2;
  return 0;
}

/** Ligne du joueur dans Présences (créée au besoin; nom/position tenus à jour). */
function lignePres_(sh, j) {
  let r = trouverLignePres_(sh, j.id);
  if (!r) {
    r = derniereLigneId_(sh) + 1;
    sh.getRange(r, 1, 1, 4).setValues([[j.id, sur_(j.prenom), sur_(j.nom), sur_(j.position || '')]]);
  } else if (j.position) {
    sh.getRange(r, 2, 1, 3).setValues([[sur_(j.prenom), sur_(j.nom), sur_(j.position)]]);
  }
  return r;
}

/**
 * Chaque joueur de Inscriptions (même ajouté à la main) reçoit un ID et une ligne dans Présences,
 * et chaque pratique a ses colonnes. Retourne le nombre de joueurs.
 */
function synchroniser_() {
  const ins = classeur_().getSheetByName(F_INSCR);
  const C = colonnesInscr_(ins);
  const sh = feuillePres_();
  const n = ins.getLastRow();
  let total = 0;
  if (n > 1) {
    const v = ins.getRange(2, 1, n - 1, ins.getLastColumn()).getValues();
    const ids = v.map(r => r[C.id - 1]);
    const existants = {};
    const der = derniereLigneId_(sh);
    if (der > 1) sh.getRange(2, 1, der - 1, 4).getValues().forEach(r => { existants[String(r[0])] = r; });
    const ajouts = [];
    v.forEach((r, i) => {
      const prenom = String(r[C.prenom - 1] || '').trim(), nom = String(r[C.nom - 1] || '').trim();
      if (!prenom && !nom) return;
      total++;
      let id = String(r[C.id - 1] || '');
      if (!id) { id = prochainId_(ids); ids.push(id); ins.getRange(i + 2, C.id).setValue(id); }
      const pos = String(r[C.position - 1] || '');
      const ex = existants[id];
      if (!ex) ajouts.push([id, sur_(prenom), sur_(nom), sur_(pos)]);
      else if (ex[1] !== prenom || ex[2] !== nom || ex[3] !== pos) {
        sh.getRange(trouverLignePres_(sh, id), 2, 1, 3).setValues([[sur_(prenom), sur_(nom), sur_(pos)]]);
      }
    });
    if (ajouts.length) sh.getRange(derniereLigneId_(sh) + 1, 1, ajouts.length, 4).setValues(ajouts);
  }
  lirePratiques_().forEach(p => colsPratique_(sh, p.date));
  appliquerFormats_(sh);
  return total;
}

/** Trouve les colonnes de Inscriptions; ajoute à droite celles qui manquent. */
function colonnesInscr_(sh) {
  const largeur = Math.max(sh.getLastColumn(), 1);
  const titres = sh.getRange(1, 1, 1, largeur).getValues()[0].map(norm_);
  const C = {};
  CHAMPS.forEach(f => {
    let vu = 0;
    for (let i = 0; i < titres.length; i++) {
      if (titres[i] && f.m(titres[i])) { vu++; if (vu === (f.occ || 1)) { C[f.k] = i + 1; break; } }
    }
  });
  let fin = largeur;
  if (!String(sh.getRange(1, largeur).getValue()).trim()) fin = largeur - 1;
  CHAMPS.forEach(f => {
    if (!C[f.k]) {
      fin++;
      sh.getRange(1, fin).setValue(f.t).setFontWeight('bold');
      titres[fin - 1] = norm_(f.t);
      C[f.k] = fin;
    }
  });
  return C;
}

/* ============================================================
 * Utilitaires
 * ============================================================ */

function classeur_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function lireParam_() {
  const o = {};
  PARAM_DEFAUT.forEach(r => { o[r[0]] = String(r[1]); });
  const sh = classeur_().getSheetByName(F_PARAM);
  if (sh) sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) o[String(r[0]).trim()] = String(r[1]).trim(); });
  return o;
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

function memeNom_(p1, n1, p2, n2) {
  return norm_(p1) === norm_(p2) && norm_(n1) === norm_(n2) && norm_(p1) !== '';
}

function oui_(v) { return /^(oui|o|yes|true|1)$/i.test(String(v).trim()); }

function norm_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/[‐-―]/g, '-').replace(/\s+/g, ' ').trim();
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
}

/** Empêche l'injection de formules (texte qui commence par = + - @). */
function sur_(v) { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; }
