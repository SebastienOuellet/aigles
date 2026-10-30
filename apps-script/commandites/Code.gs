/**
 * Commandites LHSS — Aigles AAA de Sherbrooke
 * Formulaire parent -> facture PDF -> courriels -> Google Sheets
 *
 * Fichier : Code.gs (à coller dans Extensions > Apps Script du Google Sheet des commandites)
 */

const FEUILLE_CONFIG = 'Config';
const FEUILLE_FACTURES = 'Factures';
const TAUX_TPS = 0.05;
const TAUX_TVQ = 0.09975;
const LOGO_MAX_OCTETS = 8 * 1024 * 1024; // 8 Mo
const LOGO_TYPES = [
  'image/png', 'image/jpeg', 'image/svg+xml', 'image/webp',
  'application/pdf', 'application/postscript', 'application/illustrator'
];

// Clé, valeur par défaut, description (créées par setup() dans l'onglet Config)
const CONFIG_DEFAUT = [
  ['NOM_ORG', 'AIGLES AAA DE SHERBROOKE', 'Nom affiché sur la facture et les courriels'],
  ['ADRESSE_ORG', '264 rue Hémond, Sherbrooke, QC J1R 0A4', 'Adresse sur la facture'],
  ['COURRIEL_ORG', 'ehae@hotmail.ca', 'Courriel affiché sur la facture'],
  ['SITE_ORG', 'https://aiglessherbrookeaaa.ca/', 'Site web affiché sur la facture'],
  ['NO_TPS', '', 'OBLIGATOIRE si on facture les taxes — ex. 123456789 RT0001'],
  ['NO_TVQ', '', 'OBLIGATOIRE si on facture les taxes — ex. 1234567890 TQ0001'],
  ['PROCHAIN_NO_FACTURE', 202703, 'Ancien compteur commun — sert seulement si une équipe n’a pas son propre PROCHAIN_NO_FACTURE_<équipe>'],
  ['POURCENT_EQUIPE_MIN', 30, 'Pourcentage minimum pour l’équipe'],
  ['SEUIL_100_EQUIPE', 250, 'Montant (taxes incl.) égal ou inférieur = 100 % à l’équipe'],
  ['MODE_TEST', 'OUI', 'OUI = test (pas de vrai n°, rien dans les onglets d’équipe, courriel gérante redirigé) | NON = production'],
  ['LIEN_FORMULAIRE', 'https://sebastienouellet.github.io/aigles/commandites/', 'Lien public du formulaire (site Aigles, GitHub Pages)'],
  ['CODE_GERANTE', '', 'Code secret du mode gérante (lien du formulaire + ?gerante). Vide = mode gérante désactivé'],
  ['COURRIEL_TEST', '', 'En mode test, reçoit le courriel « gérante » (vide = courriel du parent)'],
  ['BASE_RATIO', 'AVANT_TAXES', 'AVANT_TAXES = ratio calculé sur le sous-total | TOTAL = sur le montant taxes incluses'],
  ['DESCRIPTION_DEFAUT', 'Commandite pour LHSS - Saison Hockey 2026-2027', 'Ligne de description sur la facture'],
  ['CHEQUE_ORDRE', 'AIGLES AAA DE SHERBROOKE', 'À l’ordre de (chèques)'],
  ['INTERAC_TEL', '819-679-4208', 'Destinataire du virement Interac'],
  ['INTERAC_QUESTION', 'Quelle Organisation', 'Question de sécurité Interac'],
  ['INTERAC_REPONSE', 'Aigles2018', 'Réponse Interac'],
  ['MESSAGE_MERCI', 'Au nom des jeunes Aigles, un énorme merci pour votre support !!', 'Message au bas de la facture'],
  ['EQUIPES', 'Aigles D1, Aigles D3', 'Noms EXACTS des onglets d’équipe, séparés par des virgules'],
  ['COURRIEL_GERANTE_Aigles D1', '', 'Courriel de la gérante D1 (reçoit chaque facture)'],
  ['COURRIEL_GERANTE_Aigles D3', '', 'Courriel de la gérante D3 (reçoit chaque facture)'],
  ['LOGO_CLUB_ID', '', 'Optionnel : ID Drive du logo du club (PNG/JPG) pour l’entête de la facture'],
  ['DOSSIER_ID', '', 'Rempli automatiquement'],
  ['DOSSIER_LOGOS_ID', '', 'Rempli automatiquement'],
  ['DOSSIER_FACTURES_ID', '', 'Rempli automatiquement']
];

const COLONNES_FACTURES = [
  'Horodatage', 'N° facture', 'Date facture', 'Équipe', 'Joueur',
  'Parent', 'Courriel parent', 'Tél. parent',
  'Commanditaire', 'Personne contact', 'Adresse', 'Ville', 'Province', 'Code postal',
  'Tél. commanditaire', 'Courriel commanditaire', 'Site web',
  'Description', 'Montant total (taxes incl.)', 'Sous-total', 'TPS', 'TVQ',
  '% équipe', '% joueur', '$ équipe', '$ joueur',
  'Logo', 'Facture PDF', 'Notes', 'Paiement reçu', 'Courriels envoyés',
  'Mode de paiement', 'Reçu PDF', 'Reçu émis le', 'Modifiée le'
];

/* ============================================================
 * Menu et installation
 * ============================================================ */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Commandites')
    .addItem('1. Configurer (première fois)', 'setup')
    .addItem('2. Générer une facture TEST (PDF)', 'testerFacture')
    .addItem('Afficher le lien du formulaire', 'afficherLien')
    .addSeparator()
    .addItem('Envoyer la facture par courriel (ligne sélectionnée)', 'envoyerFactureLigne')
    .addItem('Régénérer la facture (ligne sélectionnée)', 'regenererFacture')
    .addItem('Émettre un reçu de paiement (ligne sélectionnée)', 'emettreRecu')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());

  // Onglet Config
  let cfg = ss.getSheetByName(FEUILLE_CONFIG);
  if (!cfg) {
    cfg = ss.insertSheet(FEUILLE_CONFIG);
    cfg.getRange(1, 1, 1, 3).setValues([['Clé', 'Valeur', 'Description']]).setFontWeight('bold');
    cfg.getRange('B:B').setNumberFormat('@');
    cfg.setColumnWidth(1, 260); cfg.setColumnWidth(2, 340); cfg.setColumnWidth(3, 480);
    cfg.setFrozenRows(1);
  }
  const cles = cfg.getDataRange().getValues().map(r => String(r[0]));
  CONFIG_DEFAUT.forEach(r => {
    if (cles.indexOf(r[0]) === -1) cfg.appendRow([r[0], String(r[1]), r[2]]);
  });
  // Une clé de courriel et un compteur de factures par équipe
  const cfgAvant = lireConfig_();
  equipes_(cfgAvant).forEach(eq => {
    const kNo = 'PROCHAIN_NO_FACTURE_' + eq;
    if (cfg.getDataRange().getValues().map(r => String(r[0])).indexOf(kNo) === -1) {
      cfg.appendRow([kNo, String(cfgAvant.PROCHAIN_NO_FACTURE || 202701), 'Prochain n° de facture pour ' + eq + ' (s’incrémente tout seul)']);
    }
    const k = 'COURRIEL_GERANTE_' + eq;
    if (cfg.getDataRange().getValues().map(r => String(r[0])).indexOf(k) === -1) {
      cfg.appendRow([k, '', 'Courriel de la gérante ' + eq]);
    }
  });

  // Dossiers Drive
  let c = lireConfig_();
  const racine = dossier_(c.DOSSIER_ID, () => DriveApp.createFolder('Commandites LHSS 2026-2027'));
  ecrireConfig_('DOSSIER_ID', racine.getId());
  const logos = dossier_(c.DOSSIER_LOGOS_ID, () => racine.createFolder('Logos'));
  ecrireConfig_('DOSSIER_LOGOS_ID', logos.getId());
  const factures = dossier_(c.DOSSIER_FACTURES_ID, () => racine.createFolder('Factures'));
  ecrireConfig_('DOSSIER_FACTURES_ID', factures.getId());

  // Onglet Factures (journal complet)
  let f = ss.getSheetByName(FEUILLE_FACTURES);
  if (!f) {
    f = ss.insertSheet(FEUILLE_FACTURES);
    f.getRange(1, 1, 1, COLONNES_FACTURES.length).setValues([COLONNES_FACTURES])
      .setFontWeight('bold').setBackground('#0b2545').setFontColor('#ffffff');
    f.setFrozenRows(1);
  }

  // Vérifie que les onglets d'équipe existent
  c = lireConfig_();
  const manquants = equipes_(c).filter(eq => !ss.getSheetByName(eq));
  const avertissements = [];
  if (manquants.length) avertissements.push('Onglets introuvables : ' + manquants.join(', ') + ' (vérifie EQUIPES dans Config).');
  if (!c.NO_TPS || !c.NO_TVQ) avertissements.push('Inscris NO_TPS et NO_TVQ dans Config avant d’envoyer de vraies factures.');
  equipes_(c).forEach(eq => { if (!c['COURRIEL_GERANTE_' + eq]) avertissements.push('Courriel de la gérante manquant pour ' + eq + '.'); });

  SpreadsheetApp.getUi().alert(
    'Configuration terminée ✅\n\nDossier Drive : « Commandites LHSS 2026-2027 »\n\n' +
    (avertissements.length ? 'À compléter dans l’onglet Config :\n• ' + avertissements.join('\n• ') : 'Tout est prêt.')
  );
}

function afficherLien() {
  const url = lireConfig_().LIEN_FORMULAIRE;
  const html =
    '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5">' +
    '<b>Lien pour les parents</b><br><a href="' + url + '" target="_blank">' + esc_(url) + '</a><br><br>' +
    '<b>Espace gérante</b> (code requis — ne pas partager)<br><a href="' + url + 'gerante.html" target="_blank">' + esc_(url) + 'gerante.html</a></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(170), 'Liens du formulaire');
}

/** Génère un PDF de démonstration dans le dossier Factures, sans courriel ni numéro consommé. */
function testerFacture() {
  const c = lireConfig_();
  const d = {
    equipe: equipes_(c)[0], joueur: 'Joueur Exemple',
    parentNom: 'Parent Exemple', parentCourriel: 'parent@exemple.com',
    commanditaire: 'Entreprise Exemple inc.', contact: 'Marie Tremblay',
    adresse: '123, rue King Ouest', ville: 'Sherbrooke', province: 'QC', codePostal: 'J1H 1A1',
    telephone: '819-555-1234', courriel: 'info@exemple.com',
    description: c.DESCRIPTION_DEFAUT, montant: 1500, pctEquipe: 30
  };
  const calc = calculer_(d.montant, d.pctEquipe, c);
  const blob = genererPdf_(d, calc, 'TEST', new Date(), c);
  const fichier = DriveApp.getFolderById(c.DOSSIER_FACTURES_ID).createFile(blob.setName('Facture_TEST.pdf'));
  SpreadsheetApp.getUi().alert('PDF test créé :\n' + fichier.getUrl());
}

/* ============================================================
 * Web app
 * ============================================================ */

/** L'ancien lien /exec renvoie vers le site (le formulaire n'est plus servi par Apps Script). */
function doGet() {
  const url = lireConfig_().LIEN_FORMULAIRE;
  return HtmlService.createHtmlOutput(
    '<div style="font-family:Arial,sans-serif;text-align:center;padding:40px 16px">' +
    '<p>Le formulaire des commandites a déménagé.</p>' +
    '<p><a href="' + url + '" target="_top" style="font-size:18px;font-weight:bold;color:#7a1426">Aller au formulaire →</a></p></div>')
    .setTitle('Commandites — Aigles D3')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * API pour le formulaire hébergé hors Google (évite le bug des comptes Google multiples).
 * Requête : POST text/plain {action: 'config' | 'submit', data: {...}}
 */
function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.action === 'config') out = { ok: true, data: getFormConfig() };
    else if (req.action === 'submit') out = { ok: true, data: submitCommandite(req.data) };
    else if (req.action === 'listeFactures') { verifierCode_(req.code); const c0 = lireConfig_(); out = { ok: true, data: { factures: listeFactures_(), pctMin: Number(c0.POURCENT_EQUIPE_MIN) || 30, seuil: seuil_(c0) } }; }
    else if (req.action === 'modifierFacture') { verifierCode_(req.code); out = { ok: true, data: modifierFacture_(req.data || {}) }; }
    else if (req.action === 'emettreRecu') { verifierCode_(req.code); out = { ok: true, data: emettreRecuConfirme(req.data || {}) }; }
    else throw new Error('Action inconnue.');
  } catch (err) {
    console.error(err);
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** Appelé par le formulaire au chargement. */
function getFormConfig() {
  const c = lireConfig_();
  const ss = classeur_();
  return {
    nomOrg: c.NOM_ORG,
    description: c.DESCRIPTION_DEFAUT,
    pctMin: Number(c.POURCENT_EQUIPE_MIN) || 30,
    baseRatio: c.BASE_RATIO === 'TOTAL' ? 'TOTAL' : 'AVANT_TAXES',
    tauxTps: TAUX_TPS, tauxTvq: TAUX_TVQ,
    seuil100: seuil_(c),
    modeTest: modeTest_(c),
    equipes: equipes_(c).map(nom => ({ nom: nom, joueurs: joueurs_(ss.getSheetByName(nom)) }))
  };
}

/** Appelé par le formulaire à la soumission. */
function submitCommandite(d) {
  const c = lireConfig_();
  d = nettoyer_(d, c);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Le système est occupé, réessaie dans quelques secondes.');

  const test = modeTest_(c);
  let no, date, calc, pdfBlob, pdfFile, logoFile, ligne;
  try {
    date = d.dateFacture || new Date();
    no = test
      ? 'TEST-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMdd-HHmmss')
      : (d.noExistant ? reserverNumeroExistant_(d.noExistant, d.equipe) : String(prochainNumero_(d.equipe)));
    calc = calculer_(d.montant, d.pctEquipe, c);

    if (d.logo) logoFile = sauverLogo_(d, no, c);

    pdfBlob = genererPdf_(d, calc, no, date, c).setName(nomFichier_(no, d.commanditaire));
    pdfFile = sousDossier_(c.DOSSIER_FACTURES_ID, d.equipe).createFile(pdfBlob);

    ligne = ecrireFactures_(d, calc, no, date, pdfFile, logoFile);
    const eq = test ? null : ecrireEquipe_(d, calc, no, pdfFile); // en test : seulement l'onglet Factures
    if (d.datePmt) marquerPaye_(ligne, eq, d.datePmt, d.modePmt);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  // Courriels hors du verrou : si ça échoue, la facture est quand même enregistrée
  const resultat = envoyerCourriels_(d, calc, no, date, pdfBlob, pdfFile, logoFile, c, false,
    { parent: d.envoyerParent, gerante: d.envoyerGerante });
  try {
    classeur_().getSheetByName(FEUILLE_FACTURES)
      .getRange(ligne, COLONNES_FACTURES.indexOf('Courriels envoyés') + 1)
      .setValue(resultat.resume);
  } catch (e) { console.error(e); }

  return {
    no: no, test: test,
    total: calc.total,
    sousTotal: calc.sousTotal, tps: calc.tps, tvq: calc.tvq,
    equipe$: calc.equipe, joueur$: calc.joueur,
    nomFichier: pdfBlob.getName(),
    pdfBase64: Utilities.base64Encode(pdfBlob.getBytes()),
    courriels: resultat
  };
}

/* ============================================================
 * Logique métier
 * ============================================================ */

/** Le montant saisi est TAXES INCLUSES (comme la facture Marché Terril : 3000 $ total). */
function calculer_(montant, pctEquipe, c) {
  const total = r2_(montant);
  if (total <= seuil_(c)) pctEquipe = 100; // petits montants : 100 % équipe
  const estime = r2_(total / (1 + TAUX_TPS + TAUX_TVQ));
  const tps = r2_(estime * TAUX_TPS);
  const tvq = r2_(estime * TAUX_TVQ);
  const sousTotal = r2_(total - tps - tvq); // garantit sous-total + taxes = total exact
  const base = c.BASE_RATIO === 'TOTAL' ? total : sousTotal;
  const equipe = r2_(base * pctEquipe / 100);
  return {
    total: total, sousTotal: sousTotal, tps: tps, tvq: tvq,
    base: base, pctEquipe: pctEquipe, pctJoueur: 100 - pctEquipe,
    equipe: equipe, joueur: r2_(base - equipe)
  };
}

function nettoyer_(d, c) {
  if (!d || typeof d !== 'object') throw new Error('Formulaire vide.');
  if (d.siteWebConfirm) throw new Error('Soumission refusée.'); // champ piège anti-robot

  const t = v => String(v == null ? '' : v).trim().slice(0, 300);
  const o = {
    equipe: t(d.equipe), joueur: t(d.joueur),
    parentNom: t(d.parentNom), parentCourriel: t(d.parentCourriel).toLowerCase(), parentTel: t(d.parentTel),
    commanditaire: t(d.commanditaire), contact: t(d.contact),
    adresse: t(d.adresse), ville: t(d.ville), province: t(d.province || 'QC').toUpperCase(),
    codePostal: t(d.codePostal).toUpperCase().replace(/\s|-/g, ''),
    telephone: t(d.telephone), courriel: t(d.courriel).toLowerCase(), siteWeb: t(d.siteWeb),
    description: t(d.description) || c.DESCRIPTION_DEFAUT,
    notes: String(d.notes || '').trim().slice(0, 1000),
    montant: Number(String(d.montant).replace(',', '.')),
    pctEquipe: Math.round(Number(d.pctEquipe)),
    logo: d.logo || null,
    gerante: false, envoyerParent: true, envoyerGerante: true,
    noExistant: '', dateFacture: null, datePmt: null, modePmt: ''
  };

  // Mode gérante : saisie de commandites déjà conclues
  if (d.codeGerante) {
    verifierCode_(d.codeGerante, c);
    o.gerante = true;
    o.envoyerParent = !!d.envoyerParent;
    o.envoyerGerante = !!d.envoyerGerante;
    o.noExistant = String(d.noExistant || '').replace(/\D/g, '');
    o.dateFacture = dateIso_(d.dateFacture);
    o.datePmt = d.dejaPaye ? (dateIso_(d.datePmt) || new Date()) : null;
    o.modePmt = t(d.modePmt).slice(0, 40);
  }

  const min = Number(c.POURCENT_EQUIPE_MIN) || 30;
  if (o.montant > 0 && o.montant <= seuil_(c)) o.pctEquipe = 100;
  const err = [];
  if (equipes_(c).indexOf(o.equipe) === -1) err.push('Équipe invalide.');
  if (!o.joueur) err.push('Nom du joueur requis.');
  // En mode gérante, parent et adresse deviennent optionnels (vieilles commandites)
  if (!o.gerante && !o.parentNom) err.push('Nom du parent requis.');
  if ((!o.gerante || o.envoyerParent || o.parentCourriel) && !courrielValide_(o.parentCourriel)) err.push('Courriel du parent invalide.');
  if (!o.commanditaire) err.push('Nom du commanditaire requis.');
  if (!o.gerante && (!o.adresse || !o.ville)) err.push('Adresse du commanditaire requise.');
  if ((!o.gerante || o.codePostal) && !/^[A-Z]\d[A-Z]\d[A-Z]\d$/.test(o.codePostal)) err.push('Code postal invalide (ex. J1H 1A1).');
  if (o.noExistant && !/^\d{4,10}$/.test(o.noExistant)) err.push('N° de facture existant invalide.');
  if (o.courriel && !courrielValide_(o.courriel)) err.push('Courriel du commanditaire invalide.');
  if (!(o.montant > 0 && o.montant <= 100000)) err.push('Montant invalide.');
  if (!(o.pctEquipe >= min && o.pctEquipe <= 100)) err.push('Le % équipe doit être entre ' + min + ' % et 100 %.');
  if (o.logo) {
    if (LOGO_TYPES.indexOf(o.logo.type) === -1 && !/\.(ai|eps|svg|png|jpe?g|pdf|webp)$/i.test(o.logo.nom || '')) err.push('Format de logo non accepté.');
    if ((String(o.logo.base64 || '').length * 3 / 4) > LOGO_MAX_OCTETS) err.push('Logo trop lourd (max 8 Mo).');
  }
  if (err.length) throw new Error(err.join(' '));

  if (o.codePostal) o.codePostal = o.codePostal.slice(0, 3) + ' ' + o.codePostal.slice(3);
  return o;
}

/** Ligne de Config du compteur de l'équipe (ou du compteur commun si absent). */
function ligneCompteur_(cfg, equipe) {
  const v = cfg.getDataRange().getValues();
  const cles = ['PROCHAIN_NO_FACTURE_' + equipe, 'PROCHAIN_NO_FACTURE'];
  for (const k of cles) for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim() === k) return { r: i + 1, val: parseInt(v[i][1], 10), cle: k };
  throw new Error('Compteur de facture introuvable — lance « Configurer » dans le menu Commandites.');
}

function prochainNumero_(equipe) {
  const cfg = classeur_().getSheetByName(FEUILLE_CONFIG);
  const L = ligneCompteur_(cfg, equipe);
  if (!L.val) throw new Error(L.cle + ' invalide dans Config.');
  cfg.getRange(L.r, 2).setValue(String(L.val + 1));
  return L.val;
}

/** Utilise un n° déjà émis à la main; refuse les doublons (dans l'équipe) et repousse le compteur si nécessaire. */
function reserverNumeroExistant_(no, equipe) {
  const sh = classeur_().getSheetByName(FEUILLE_FACTURES);
  const m = entetesFactures_(sh);
  const cNo = m[norm_('N° facture')], cEq = m[norm_('Équipe')];
  if (cNo && sh.getLastRow() > 1) {
    const v = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
    if (v.some(r => String(r[cNo - 1]).trim() === String(no) && (!cEq || String(r[cEq - 1]).trim() === equipe))) {
      throw new Error('La facture ' + no + ' existe déjà pour ' + equipe + '.');
    }
  }
  const cfg = classeur_().getSheetByName(FEUILLE_CONFIG);
  const L = ligneCompteur_(cfg, equipe);
  if (Number(no) >= L.val) cfg.getRange(L.r, 2).setValue(String(Number(no) + 1));
  return String(no);
}

/** Inscrit un paiement déjà reçu (onglet Factures + onglet d'équipe), sans générer de reçu. */
function marquerPaye_(ligne, eq, datePmt, mode) {
  const sh = classeur_().getSheetByName(FEUILLE_FACTURES);
  const m = entetesFactures_(sh);
  sh.getRange(ligne, colFactures_(sh, m, 'Paiement reçu')).setValue(datePmt).setNumberFormat('yyyy-mm-dd');
  if (mode) sh.getRange(ligne, colFactures_(sh, m, 'Mode de paiement')).setValue(sur_(mode));
  if (!eq) return;
  if (eq.C.pmtRecu) eq.sh.getRange(eq.r, eq.C.pmtRecu).setValue(datePmt).setNumberFormat('yyyy-mm-dd');
  else if (eq.C.note) {
    const n = eq.sh.getRange(eq.r, eq.C.note), actuel = String(n.getValue() || '');
    if (!/payé/i.test(actuel)) n.setValue(sur_((actuel ? actuel + ' — ' : '') + 'payé ' + dateCourte_(datePmt)));
  }
}

function dateIso_(s) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')) ? new Date(s + 'T12:00:00') : null;
}

function sauverLogo_(d, no, c) {
  const ext = (d.logo.nom || '').split('.').pop().toLowerCase() || 'png';
  const nom = [d.equipe.replace(/\s+/g, ''), no, slug_(d.commanditaire), 'logo'].join('_') + '.' + ext;
  const blob = Utilities.newBlob(Utilities.base64Decode(d.logo.base64), d.logo.type || 'application/octet-stream', nom);
  return sousDossier_(c.DOSSIER_LOGOS_ID, d.equipe).createFile(blob);
}

/* ============================================================
 * Google Sheets
 * ============================================================ */

function ecrireFactures_(d, calc, no, date, pdfFile, logoFile) {
  const sh = classeur_().getSheetByName(FEUILLE_FACTURES);
  const ligne = [
    new Date(), /^\d+$/.test(no) ? Number(no) : no, date, d.equipe, d.joueur,
    d.parentNom, d.parentCourriel, d.parentTel,
    d.commanditaire, d.contact, d.adresse, d.ville, d.province, d.codePostal,
    d.telephone, d.courriel, d.siteWeb,
    d.description, calc.total, calc.sousTotal, calc.tps, calc.tvq,
    calc.pctEquipe / 100, calc.pctJoueur / 100, calc.equipe, calc.joueur,
    logoFile ? 'Logo' : '(à recevoir)',
    'Facture ' + no,
    d.notes, '', ''
  ].map(v => typeof v === 'string' ? sur_(v) : v);
  sh.appendRow(ligne);
  const r = sh.getLastRow();
  if (logoFile) lienRiche_(sh.getRange(r, 27), logoFile.getUrl(), 'Logo');
  lienRiche_(sh.getRange(r, 28), pdfFile.getUrl(), 'Facture ' + no);
  sh.getRange(r, 19, 1, 4).setNumberFormat('#,##0.00 $');
  sh.getRange(r, 23, 1, 2).setNumberFormat('0%');
  sh.getRange(r, 25, 1, 2).setNumberFormat('#,##0.00 $');
  sh.getRange(r, 3).setNumberFormat('yyyy-mm-dd');
  return r;
}

/**
 * Écrit une ligne dans l'onglet de l'équipe, sous le bloc du joueur.
 * Les colonnes sont trouvées par leur titre (fonctionne pour le format D1 et D3).
 */
function ecrireEquipe_(d, calc, no, pdfFile) {
  const sh = classeur_().getSheetByName(d.equipe);
  if (!sh) return;
  const C = colonnesEquipe_(sh);

  // Trouve le bloc du joueur
  const derniere = Math.max(sh.getLastRow(), 1);
  const colA = sh.getRange(1, 1, derniere, 1).getValues().map(r => norm_(r[0]));
  const colB = C.commanditaire ? sh.getRange(1, C.commanditaire, derniere, 1).getValues().map(r => String(r[0]).trim()) : [];
  const idx = colA.indexOf(norm_(d.joueur));

  let cible, ecrireNom = false;
  if (idx === -1) {
    cible = derniere + 2; ecrireNom = true;            // nouveau joueur : à la fin
  } else if (!colB[idx]) {
    cible = idx + 1;                                    // ligne du joueur sans commanditaire
  } else {
    let fin = idx;
    while (fin + 1 < derniere && !colA[fin + 1] && colB[fin + 1]) fin++;
    // Commanditaire déjà inscrit à la main pour ce joueur? On complète cette ligne au lieu d'en ajouter une.
    for (let i = idx; i <= fin; i++) {
      if (norm_(colB[i]) === norm_(d.commanditaire)) {
        remplirEquipe_(sh, C, i + 1, d, calc, no, pdfFile, false, null); // garde la note existante
        return { sh: sh, C: C, r: i + 1 };
      }
    }
    sh.insertRowAfter(fin + 1);
    cible = fin + 2;
  }
  const note = d.gerante ? d.notes : ['Facture envoyée', d.notes].filter(Boolean).join(' — ');
  remplirEquipe_(sh, C, cible, d, calc, no, pdfFile, ecrireNom, note);
  return { sh: sh, C: C, r: cible };
}

function colonnesEquipe_(sh) {
  const titres = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(norm_);
  const col = alias => { for (const a of alias) { const i = titres.indexOf(a); if (i > -1) return i + 1; } return 0; };
  return {
    joueur: col(['joueur']), commanditaire: col(['commanditaire']),
    montant: col(['montant', 'montant total']),
    pctEquipe: col(['% equipe']), pctJoueur: col(['% joueur']),
    dEquipe: col(['$ equipe']), dJoueur: col(['$ joueur']),
    note: col(['note', 'notes']), adresse: col(['adresse']),
    tel: col(['telephone', 'tel']), courriel: col(['courriel', 'email']),
    facture: col(['# facture', 'no facture', 'n° facture', 'recu']),
    recuEmis: col(['recu emis']), pmtRecu: col(['pmt recu', 'paiement recu'])
  };
}

function remplirEquipe_(sh, C, cible, d, calc, no, pdfFile, ecrireNom, note) {
  const set = (c, v, fmt) => { if (c) { const g = sh.getRange(cible, c); g.setValue(v); if (fmt) g.setNumberFormat(fmt); } };
  if (ecrireNom) set(C.joueur, sur_(d.joueur));
  set(C.commanditaire, sur_(d.commanditaire));
  set(C.montant, calc.total, '#,##0.00 $');
  if (C.dEquipe || C.dJoueur) {
    // Format D3 : % numériques + colonnes $
    set(C.pctEquipe, calc.pctEquipe); set(C.pctJoueur, calc.pctJoueur || '');
    set(C.dEquipe, calc.equipe, '#,##0.00 $'); set(C.dJoueur, calc.joueur || '', '#,##0.00 $');
  } else {
    // Format D1 : les colonnes « % » contiennent des montants en $
    set(C.pctEquipe, calc.equipe, '#,##0.00 $'); set(C.pctJoueur, calc.joueur || '-', '#,##0.00 $');
  }
  set(C.adresse, sur_([d.adresse, d.ville, d.province, d.codePostal].join(', ')));
  set(C.tel, sur_(d.telephone));
  set(C.courriel, sur_(d.courriel));
  if (note != null) set(C.note, sur_(note));
  if (C.facture && pdfFile) lienRiche_(sh.getRange(cible, C.facture), pdfFile.getUrl(), no);
}

/** Retrouve la ligne d'une facture dans l'onglet d'équipe (par le n° dans la colonne # facture). */
function trouverLigneEquipe_(equipe, no) {
  const sh = classeur_().getSheetByName(equipe);
  if (!sh || sh.getLastRow() < 2) return null;
  const C = colonnesEquipe_(sh);
  if (!C.facture) return null;
  const v = sh.getRange(2, C.facture, sh.getLastRow() - 1, 1).getDisplayValues();
  for (let i = 0; i < v.length; i++) if (String(v[i][0]).trim() === String(no)) return { sh: sh, C: C, r: i + 2 };
  return null;
}

/* ============================================================
 * PDF
 * ============================================================ */

function genererPdf_(d, calc, no, date, c) {
  const html = construireHtmlFacture_(d, calc, no, date, c, logoClub_(c));
  return Utilities.newBlob(html, 'text/html', 'facture.html').getAs('application/pdf');
}

function construireHtmlFacture_(d, calc, no, date, c, logoDataUri) {
  const e = esc_;
  const adresse = [d.adresse, [d.ville, d.province, d.codePostal].filter(Boolean).join(' ')].filter(Boolean).join(', ');
  const taxes = (c.NO_TPS || c.NO_TVQ)
    ? '<div class="taxnos">TPS : ' + e(c.NO_TPS || '—') + ' &nbsp;·&nbsp; TVQ : ' + e(c.NO_TVQ || '—') + '</div>' : '';
  const ligneInfo = (lbl, val) => val ? '<tr><td class="lbl">' + lbl + '</td><td>' + e(val) + '</td></tr>' : '';

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' +
    cssDoc_() +
    '</style></head><body>' +
    (modeTest_(c) ? '<div class="test">FACTURE TEST — NE PAS PAYER</div>' : '') +

    '<table><tr><td style="vertical-align:top">' +
      (logoDataUri ? '<img src="' + logoDataUri + '" style="height:70px;margin-bottom:6px"><br>' : '') +
      '<div class="org">' + e(c.NOM_ORG) + '</div>' +
      '<div class="muted">' + e(c.ADRESSE_ORG) + '<br>' + e(c.COURRIEL_ORG) + ' &nbsp;|&nbsp; ' + e(c.SITE_ORG) + '</div>' + taxes +
    '</td><td style="vertical-align:top;width:230px">' +
      '<div class="titre">FACTURE</div>' +
      '<table class="box" style="margin-top:8px"><tr><td class="lbl">N° de facture</td><td class="r"><b>' + e(no) + '</b></td></tr>' +
      '<tr><td class="lbl">Date</td><td class="r">' + dateFr_(date) + '</td></tr>' +
      '<tr><td class="lbl">Équipe</td><td class="r">' + e(d.equipe) + '</td></tr></table>' +
    '</td></tr></table>' +

    '<div class="bar"></div>' +

    '<div class="h">Facturer à</div>' +
    '<table class="box">' +
      '<tr><td class="lbl">Entreprise</td><td><b>' + e(d.commanditaire) + '</b></td></tr>' +
      ligneInfo('Contact', d.contact) + ligneInfo('Adresse', adresse) +
      ligneInfo('Téléphone', d.telephone) + ligneInfo('Courriel', d.courriel) +
    '</table>' +

    '<table class="items" style="margin-top:22px"><tr>' +
      '<th style="width:90px">Date</th><th>Description</th><th class="r" style="width:40px">Qté</th>' +
      '<th class="r" style="width:95px">Prix unitaire</th><th class="r" style="width:95px">Total</th></tr>' +
    '<tr><td>' + dateCourte_(date) + '</td><td>' + e(d.description) +
      '<br><span class="muted">Joueur : ' + e(d.joueur) + '</span></td>' +
      '<td class="r">1</td><td class="r">' + argent_(calc.sousTotal) + '</td><td class="r">' + argent_(calc.sousTotal) + '</td></tr>' +
    '</table>' +

    '<table class="tot" style="width:300px;margin:14px 0 0 auto">' +
      '<tr><td>Sous-total</td><td class="r">' + argent_(calc.sousTotal) + '</td></tr>' +
      '<tr><td>TPS (5 %)</td><td class="r">' + argent_(calc.tps) + '</td></tr>' +
      '<tr><td>TVQ (9,975 %)</td><td class="r">' + argent_(calc.tvq) + '</td></tr>' +
      '<tr><td>Acompte reçu</td><td class="r">' + argent_(0) + '</td></tr>' +
      '<tr class="grand"><td>TOTAL À PAYER</td><td class="r">' + argent_(calc.total) + '</td></tr>' +
    '</table>' +

    '<div class="pay"><b>Modalités de paiement</b><br>' +
      'Chèque à l’ordre de <b>' + e(c.CHEQUE_ORDRE) + '</b><br>' +
      'Virement Interac : <b>' + e(c.INTERAC_TEL) + '</b><br>' +
      'Question de sécurité : ' + e(c.INTERAC_QUESTION) + ' &nbsp;—&nbsp; Réponse : <b>' + e(c.INTERAC_REPONSE) + '</b><br>' +
      'Merci d’indiquer le <b>n° de facture ' + e(no) + '</b> dans le message du virement.' +
    '</div>' +

    '<div class="merci">' + e(c.MESSAGE_MERCI) + '</div>' +
    '</body></html>';
}

function cssDoc_() {
  return 'body{font-family:Arial,Helvetica,sans-serif;font-size:10.5pt;color:#1a1a1a;margin:28px 34px;}' +
    'table{border-collapse:collapse;width:100%;}' +
    '.org{font-size:18pt;font-weight:bold;color:#0b2545;letter-spacing:.5px;}' +
    '.muted{color:#555;font-size:9.5pt;line-height:1.45;}' +
    '.taxnos{color:#555;font-size:9pt;margin-top:2px;}' +
    '.titre{font-size:22pt;font-weight:bold;color:#0b2545;text-align:right;}' +
    '.bar{height:4px;background:#c9a227;margin:14px 0 18px;}' +
    '.lbl{color:#555;width:95px;vertical-align:top;padding:2px 8px 2px 0;}' +
    '.box td{padding:2px 0;}' +
    '.h{font-size:9pt;text-transform:uppercase;color:#0b2545;font-weight:bold;letter-spacing:.6px;margin-bottom:4px;}' +
    '.items th{background:#0b2545;color:#fff;font-size:9pt;text-align:left;padding:7px 8px;}' +
    '.items td{padding:9px 8px;border-bottom:1px solid #ddd;vertical-align:top;}' +
    '.r{text-align:right;}' +
    '.tot td{padding:4px 8px;}' +
    '.grand td{font-size:13pt;font-weight:bold;color:#0b2545;border-top:2px solid #0b2545;padding-top:8px;}' +
    '.pay{background:#f4f6fa;border-left:4px solid #c9a227;padding:12px 14px;margin-top:24px;line-height:1.6;}' +
    '.merci{text-align:center;margin-top:26px;font-weight:bold;color:#0b2545;font-size:11.5pt;}' +
    '.test{background:#b42318;color:#fff;text-align:center;font-weight:bold;padding:8px;margin-bottom:16px;letter-spacing:1px;}';
}

function logoClub_(c) {
  if (!c.LOGO_CLUB_ID) return '';
  try {
    const b = DriveApp.getFileById(c.LOGO_CLUB_ID).getBlob();
    return 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
  } catch (e) { console.warn('Logo club introuvable', e); return ''; }
}

/* ============================================================
 * Courriels
 * ============================================================ */

function envoyerCourriels_(d, calc, no, date, pdfBlob, pdfFile, logoFile, c, corrigee, opts) {
  opts = opts || { parent: true, gerante: true };
  const test = modeTest_(c);
  // En mode test, la vraie gérante ne reçoit RIEN
  const gerante = test ? (c.COURRIEL_TEST || d.parentCourriel) : c['COURRIEL_GERANTE_' + d.equipe];
  const pre = (test ? '[TEST] ' : '') + (corrigee ? '[CORRIGÉE] ' : '');
  const res = { parent: false, gerante: false, resume: test ? 'MODE TEST — ' : '' };
  const style = 'font-family:Arial,sans-serif;font-size:14px;color:#1a1a1a;line-height:1.5';

  if (opts.parent && d.parentCourriel) try {
    MailApp.sendEmail({
      to: d.parentCourriel,
      replyTo: gerante || c.COURRIEL_ORG,
      name: c.NOM_ORG,
      subject: pre + 'Facture ' + no + ' — Commandite ' + d.commanditaire,
      htmlBody:
        '<div style="' + style + '">Bonjour ' + esc_(d.parentNom) + ',<br><br>' +
        'Merci d’avoir trouvé une commandite pour <b>' + esc_(d.joueur) + '</b>! 🦅<br><br>' +
        'Voici la facture <b>' + no + '</b> à transmettre à <b>' + esc_(d.commanditaire) + '</b> ' +
        '(montant : <b>' + argent_(calc.total) + '</b>).<br><br>' +
        'Les modalités de paiement (chèque ou virement Interac) sont indiquées sur la facture.<br><br>' +
        'Répartition : équipe ' + calc.pctEquipe + ' %' +
        (calc.pctJoueur ? ' · joueur ' + calc.pctJoueur + ' % (' + argent_(calc.joueur) + ')' : '') + '.<br><br>' +
        'Pour toute question, réponds simplement à ce courriel.<br><br>' + esc_(c.NOM_ORG) + '</div>',
      attachments: [pdfBlob]
    });
    res.parent = true;
  } catch (e) { console.error('Courriel parent', e); }

  if (gerante && opts.gerante) {
    try {
      const r = (k, v) => '<tr><td style="color:#555;padding:3px 12px 3px 0">' + k + '</td><td>' + v + '</td></tr>';
      const pj = [pdfBlob];
      if (logoFile && logoFile.getSize() < 5 * 1024 * 1024) pj.push(logoFile.getBlob());
      MailApp.sendEmail({
        to: gerante,
        replyTo: d.parentCourriel || c.COURRIEL_ORG,
        name: 'Commandites ' + d.equipe,
        subject: pre + '[Commandite] ' + d.equipe + ' — ' + d.joueur + ' — ' + d.commanditaire + ' — ' + argent_(calc.total),
        htmlBody:
          '<div style="' + style + '"><b>Nouvelle commandite — facture ' + no + '</b><br><br><table>' +
          r('Joueur', esc_(d.joueur)) +
          r('Parent', esc_(d.parentNom) + ' · ' + esc_(d.parentCourriel) + (d.parentTel ? ' · ' + esc_(d.parentTel) : '')) +
          r('Commanditaire', esc_(d.commanditaire) + (d.contact ? ' (' + esc_(d.contact) + ')' : '')) +
          r('Adresse', esc_([d.adresse, d.ville, d.province, d.codePostal].join(', '))) +
          r('Contact', esc_([d.telephone, d.courriel].filter(Boolean).join(' · ') || '—')) +
          r('Montant', '<b>' + argent_(calc.total) + '</b> (sous-total ' + argent_(calc.sousTotal) + ' + TPS ' + argent_(calc.tps) + ' + TVQ ' + argent_(calc.tvq) + ')') +
          r('Équipe', argent_(calc.equipe) + ' (' + calc.pctEquipe + ' %)') +
          r('Joueur', argent_(calc.joueur) + ' (' + calc.pctJoueur + ' %)') +
          r('Logo', logoFile ? '<a href="' + logoFile.getUrl() + '">Voir dans Drive</a>' : '<i>à recevoir</i>') +
          r('Facture', '<a href="' + pdfFile.getUrl() + '">Voir dans Drive</a>') +
          (d.notes ? r('Notes', esc_(d.notes)) : '') +
          '</table></div>',
        attachments: pj
      });
      res.gerante = true;
    } catch (e) { console.error('Courriel gérante', e); }
  }

  const etat = (voulu, ok) => !voulu ? 'non demandé' : (ok ? 'oui' : 'ÉCHEC');
  res.resume += 'Parent : ' + etat(opts.parent && !!d.parentCourriel, res.parent) +
    ' · Gérante : ' + (gerante ? etat(opts.gerante, res.gerante) : 'non configurée');
  return res;
}

/* ============================================================
 * Gérante : régénérer une facture / émettre un reçu
 * (exécuté avec le compte de la personne qui clique dans le menu)
 * ============================================================ */

function regenererFacture() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sel = ligneSelectionnee_();
    const c = lireConfig_();
    const L = lireLigne_(sel.sh, sel.r);
    validerLigne_(L, c);
    const calc = calculer_(L.d.montant, L.d.pctEquipe, c);
    const rep = ui.alert('Régénérer la facture ' + L.no,
      L.d.commanditaire + ' — ' + argent_(calc.total) + ' — équipe ' + calc.pctEquipe + ' %\n\n' +
      'OUI : régénérer ET renvoyer par courriel au parent et à la gérante\n' +
      'NON : régénérer seulement (aucun courriel)', ui.ButtonSet.YES_NO_CANCEL);
    if (rep !== ui.Button.YES && rep !== ui.Button.NO) return;
    const res = regenererLigne_(sel.sh, sel.r, c, rep === ui.Button.YES ? { parent: true, gerante: true } : null);
    ui.alert(res.message);
  } catch (e) {
    ui.alert('Impossible de régénérer : ' + e.message);
  }
}

/**
 * Refait le PDF d'une ligne de Factures (même n°), recalcule les colonnes, met à jour l'onglet d'équipe
 * et, si demandé, renvoie la facture marquée [CORRIGÉE].
 */
function regenererLigne_(sh, r, c, envoi) {
  const L = lireLigne_(sh, r);
  const d = L.d;
  validerLigne_(L, c);
  const calc = calculer_(d.montant, d.pctEquipe, c);

  const blob = genererPdf_(d, calc, L.no, L.date, c).setName(nomFichier_(L.no, d.commanditaire));
  const cPdf = colFactures_(sh, L.m, 'Facture PDF');
  remplacerFichier_(sh.getRange(r, cPdf));
  const file = sousDossier_(c.DOSSIER_FACTURES_ID, d.equipe).createFile(blob);
  lienRiche_(sh.getRange(r, cPdf), file.getUrl(), 'Facture ' + L.no);

  const set = (nom, v, fmt) => { const g = sh.getRange(r, colFactures_(sh, L.m, nom)); g.setValue(v); if (fmt) g.setNumberFormat(fmt); };
  set('Sous-total', calc.sousTotal, '#,##0.00 $'); set('TPS', calc.tps, '#,##0.00 $'); set('TVQ', calc.tvq, '#,##0.00 $');
  set('% équipe', calc.pctEquipe / 100, '0%'); set('% joueur', calc.pctJoueur / 100, '0%');
  set('$ équipe', calc.equipe, '#,##0.00 $'); set('$ joueur', calc.joueur, '#,##0.00 $');
  set('Modifiée le', new Date(), 'yyyy-mm-dd hh:mm');

  const eq = /^TEST/i.test(L.no) ? null : trouverLigneEquipe_(d.equipe, L.no);
  if (eq) remplirEquipe_(eq.sh, eq.C, eq.r, d, calc, L.no, file, false, null);

  let msg = 'Facture ' + L.no + ' régénérée ✅' + (eq ? '\nOnglet ' + d.equipe + ' mis à jour.' : '');
  let courriels = 'aucun';
  if (envoi && (envoi.parent || envoi.gerante)) {
    const res = envoyerCourriels_(d, calc, L.no, L.date, blob, file, null, c, true, envoi);
    set('Courriels envoyés', res.resume + ' (corrigée ' + dateCourte_(new Date()) + ')');
    courriels = res.resume;
    msg += '\nCourriels : ' + res.resume;
  }
  return { message: msg, no: L.no, courriels: courriels, nomFichier: blob.getName(), pdfBase64: Utilities.base64Encode(blob.getBytes()) };
}

/** API : modifie les champs d'une facture depuis la page de gestion, puis la régénère. */
function modifierFacture_(p) {
  const c = lireConfig_();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Le système est occupé, réessaie dans quelques secondes.');
  try {
    const sh = classeur_().getSheetByName(FEUILLE_FACTURES);
    const r = Number(p.r);
    const L = lireLigne_(sh, r);
    if (!r || L.no !== String(p.no)) throw new Error('La facture a changé de ligne. Rafraîchis la liste et recommence.');

    const v = p.champs || {};
    const t = x => String(x == null ? '' : x).trim().slice(0, 300);
    const montant = Number(String(v.montant).replace(/\s|\$/g, '').replace(',', '.'));
    let pct = Math.round(Number(v.pctEquipe));
    if (montant > 0 && montant <= seuil_(c)) pct = 100;
    let cp = t(v.codePostal).toUpperCase().replace(/\s|-/g, '');

    const err = [];
    if (!t(v.commanditaire)) err.push('Commanditaire requis.');
    if (!(montant > 0 && montant <= 100000)) err.push('Montant invalide.');
    const min = Number(c.POURCENT_EQUIPE_MIN) || 30;
    if (!(pct >= min && pct <= 100)) err.push('% équipe entre ' + min + ' et 100.');
    if (cp && !/^[A-Z]\d[A-Z]\d[A-Z]\d$/.test(cp)) err.push('Code postal invalide.');
    if (t(v.courriel) && !courrielValide_(t(v.courriel))) err.push('Courriel du commanditaire invalide.');
    if (t(v.parentCourriel) && !courrielValide_(t(v.parentCourriel))) err.push('Courriel du parent invalide.');
    if (p.envoyerParent && !courrielValide_(t(v.parentCourriel))) err.push('Courriel du parent requis pour lui envoyer la facture.');
    if (err.length) throw new Error(err.join(' '));
    if (cp) cp = cp.slice(0, 3) + ' ' + cp.slice(3);

    const ecrire = (nom, val, fmt) => { const g = sh.getRange(r, colFactures_(sh, L.m, nom)); g.setValue(val); if (fmt) g.setNumberFormat(fmt); };
    [['Commanditaire', 'commanditaire'], ['Personne contact', 'contact'], ['Adresse', 'adresse'], ['Ville', 'ville'],
     ['Tél. commanditaire', 'telephone'], ['Courriel commanditaire', 'courriel'], ['Description', 'description'],
     ['Parent', 'parentNom'], ['Courriel parent', 'parentCourriel']]
      .forEach(([col, k]) => ecrire(col, sur_(k === 'courriel' || k === 'parentCourriel' ? t(v[k]).toLowerCase() : t(v[k]))));
    ecrire('Province', sur_(t(v.province || 'QC').toUpperCase()));
    ecrire('Code postal', cp);
    ecrire('Montant total (taxes incl.)', montant, '#,##0.00 $');
    ecrire('% équipe', pct / 100, '0%');
    SpreadsheetApp.flush();

    return regenererLigne_(sh, r, c, { parent: !!p.envoyerParent, gerante: !!p.envoyerGerante });
  } finally {
    lock.releaseLock();
  }
}

/** Envoie plus tard une facture saisie sans courriel (mode gérante). */
function envoyerFactureLigne() {
  const ui = SpreadsheetApp.getUi();
  try {
    const sel = ligneSelectionnee_();
    const c = lireConfig_();
    const L = lireLigne_(sel.sh, sel.r);
    validerLigne_(L, c);
    const d = L.d;
    const cPdf = colFactures_(sel.sh, L.m, 'Facture PDF');
    const rt = sel.sh.getRange(sel.r, cPdf).getRichTextValue();
    const url = rt && rt.getLinkUrl();
    const id = url && (url.match(/\/d\/([\w-]+)/) || url.match(/[?&]id=([\w-]+)/) || [])[1];
    if (!id) throw new Error('Aucun PDF lié à cette ligne. Utilise « Régénérer la facture ».');
    const file = DriveApp.getFileById(id);
    if (!courrielValide_(d.parentCourriel)) throw new Error('Inscris d’abord un courriel valide dans la colonne « Courriel parent ».');

    const rep = ui.alert('Envoyer la facture ' + L.no,
      'Destinataire : ' + d.parentCourriel + '\n\nOUI : parent + copie à la gérante\nNON : parent seulement', ui.ButtonSet.YES_NO_CANCEL);
    if (rep !== ui.Button.YES && rep !== ui.Button.NO) return;

    const calc = calculer_(d.montant, d.pctEquipe, c);
    const res = envoyerCourriels_(d, calc, L.no, L.date, file.getBlob(), file, null, c, false,
      { parent: true, gerante: rep === ui.Button.YES });
    sel.sh.getRange(sel.r, colFactures_(sel.sh, L.m, 'Courriels envoyés')).setValue(res.resume + ' (' + dateCourte_(new Date()) + ')');
    ui.alert('Facture ' + L.no + ' envoyée ✅\n' + res.resume);
  } catch (e) {
    ui.alert('Impossible d’envoyer : ' + e.message);
  }
}

/**
 * Reçu de paiement via des boîtes natives (ui.prompt / ui.alert).
 * Pas de google.script.run côté client : immunisé contre le bug PERMISSION_DENIED
 * des navigateurs connectés à plusieurs comptes Google.
 */
function emettreRecu() {
  const ui = SpreadsheetApp.getUi();
  const OK = ui.Button.OK, titre = t => 'Reçu — ' + t;
  try {
    const sel = ligneSelectionnee_();
    const L = lireLigne_(sel.sh, sel.r);
    validerLigne_(L, lireConfig_());
    const d = L.d;
    const entete = d.commanditaire + ' — ' + argent_(d.montant) + '\n\n';

    const deja = L.get('Reçu émis le');
    if (deja && ui.alert(titre(L.no), entete + '⚠ Un reçu a déjà été émis le ' +
        (deja instanceof Date ? dateCourte_(deja) : deja) + '.\nLe remplacer ?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

    // 1. Date
    let date = null;
    for (let essai = 0; !date; essai++) {
      const p = ui.prompt(titre(L.no + ' (1/4)'), entete + 'Date du paiement (AAAA-MM-JJ)\nLaisse vide pour aujourd’hui (' + dateCourte_(new Date()) + ').', ui.ButtonSet.OK_CANCEL);
      if (p.getSelectedButton() !== OK) return;
      const v = p.getResponseText().trim();
      if (!v) date = dateCourte_(new Date());
      else if (/^\d{4}-\d{2}-\d{2}$/.test(v) && !isNaN(new Date(v + 'T12:00:00'))) date = v;
      else if (essai >= 2) throw new Error('Date invalide. Format attendu : 2026-10-29');
      else ui.alert('Date invalide', 'Format attendu : AAAA-MM-JJ (ex. 2026-10-29).', ui.ButtonSet.OK);
    }

    // 2. Mode
    const MODES = ['Virement Interac', 'Chèque', 'Argent comptant', 'Autre'];
    const pm = ui.prompt(titre(L.no + ' (2/4)'), 'Mode de paiement — tape le chiffre :\n1 = Virement Interac\n2 = Chèque\n3 = Argent comptant\n4 = Autre\n\n(Vide = 1)', ui.ButtonSet.OK_CANCEL);
    if (pm.getSelectedButton() !== OK) return;
    const mode = MODES[(parseInt(pm.getResponseText(), 10) || 1) - 1] || MODES[0];

    // 3. Référence
    const pr = ui.prompt(titre(L.no + ' (3/4)'), (mode === 'Chèque' ? 'N° du chèque' : 'Référence') + ' (optionnel — laisse vide si aucun)', ui.ButtonSet.OK_CANCEL);
    if (pr.getSelectedButton() !== OK) return;
    const ref = pr.getResponseText().trim();

    // 4. Destinataire
    const defaut = d.courriel || d.parentCourriel;
    const pd = ui.prompt(titre(L.no + ' (4/4)'),
      'Envoyer le reçu par courriel à :\n' + (defaut || '(aucun courriel connu)') + '\n\n' +
      'Laisse vide pour garder cette adresse, tape une autre adresse pour la changer,\nou tape NON pour ne pas envoyer de courriel.', ui.ButtonSet.OK_CANCEL);
    if (pd.getSelectedButton() !== OK) return;
    const rd = pd.getResponseText().trim();
    const envoyer = !/^non$/i.test(rd);
    const dest = rd && envoyer ? rd : defaut;
    if (envoyer && !courrielValide_(String(dest || ''))) throw new Error('Courriel invalide ou manquant : « ' + (dest || '') + ' ».');

    let copieParent = false;
    if (envoyer && d.parentCourriel && d.parentCourriel !== dest) {
      copieParent = ui.alert(titre(L.no), 'Envoyer une copie au parent (' + d.parentCourriel + ') ?', ui.ButtonSet.YES_NO) === ui.Button.YES;
    }

    // Confirmation
    const ok = ui.alert(titre(L.no) + ' — confirmation', entete +
      'Payé le : ' + date + '\nMode : ' + mode + (ref ? ' #' + ref : '') + '\n' +
      'Courriel : ' + (envoyer ? dest + (copieParent ? ' (+ copie parent)' : '') : 'aucun') + '\n\nÉmettre le reçu ?', ui.ButtonSet.YES_NO);
    if (ok !== ui.Button.YES) return;

    const res = emettreRecuConfirme({ r: sel.r, no: L.no, date: date, mode: mode, ref: ref, dest: dest, copieParent: copieParent, envoyer: envoyer });
    ui.alert(res.message);
  } catch (e) {
    ui.alert('Impossible d’émettre le reçu : ' + e.message);
  }
}

/** Vérifie le code gérante (bloque 15 min après 10 échecs). */
function verifierCode_(code, c) {
  c = c || lireConfig_();
  const cache = CacheService.getScriptCache();
  const echecs = Number(cache.get('echecs_code') || 0);
  if (echecs >= 10) throw new Error('Trop d’essais avec un mauvais code. Réessaie dans 15 minutes.');
  if (!c.CODE_GERANTE || String(code || '').trim() !== String(c.CODE_GERANTE).trim()) {
    cache.put('echecs_code', String(echecs + 1), 900);
    throw new Error('Code gérante invalide.');
  }
}

/** Liste des factures pour la page de gestion des reçus (plus récentes en premier). */
function listeFactures_() {
  const sh = classeur_().getSheetByName(FEUILLE_FACTURES);
  const n = sh.getLastRow();
  if (n < 2) return [];
  const m = entetesFactures_(sh);
  const v = sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues();
  const g = (row, nom) => { const col = m[norm_(nom)]; return col ? row[col - 1] : ''; };
  const txt = x => x instanceof Date ? dateCourte_(x) : String(x == null ? '' : x).trim();
  return v.map((row, i) => ({
    r: i + 2,
    no: txt(g(row, 'N° facture')),
    date: txt(g(row, 'Date facture')),
    equipe: txt(g(row, 'Équipe')),
    joueur: txt(g(row, 'Joueur')),
    commanditaire: txt(g(row, 'Commanditaire')),
    montant: Number(g(row, 'Montant total (taxes incl.)')) || 0,
    courriel: txt(g(row, 'Courriel commanditaire')),
    parentCourriel: txt(g(row, 'Courriel parent')),
    parentNom: txt(g(row, 'Parent')),
    contact: txt(g(row, 'Personne contact')),
    adresse: txt(g(row, 'Adresse')), ville: txt(g(row, 'Ville')),
    province: txt(g(row, 'Province')) || 'QC', codePostal: txt(g(row, 'Code postal')),
    telephone: txt(g(row, 'Tél. commanditaire')),
    description: txt(g(row, 'Description')),
    pctEquipe: (n => n > 0 && n <= 1 ? Math.round(n * 100) : Math.round(n || 100))(Number(g(row, '% équipe'))),
    modifiee: txt(g(row, 'Modifiée le')),
    paye: txt(g(row, 'Paiement reçu')),
    mode: txt(g(row, 'Mode de paiement')),
    recu: txt(g(row, 'Reçu émis le'))
  })).filter(x => x.no && x.commanditaire).reverse();
}

/** Appelé par la page de reçus (API) et par le menu du Sheet. */
function emettreRecuConfirme(p) {
  const sh = classeur_().getSheetByName(FEUILLE_FACTURES);
  const L = lireLigne_(sh, p.r);
  if (L.no !== String(p.no)) throw new Error('La ligne a changé pendant que la fenêtre était ouverte. Recommence.');
  const c = lireConfig_();
  validerLigne_(L, c);
  if (p.envoyer && !courrielValide_(String(p.dest || '').trim())) throw new Error('Courriel du destinataire invalide.');

  const d = L.d;
  const calc = calculer_(d.montant, d.pctEquipe, c);
  const datePmt = /^\d{4}-\d{2}-\d{2}$/.test(p.date) ? new Date(p.date + 'T12:00:00') : new Date();
  const mode = String(p.mode || '').slice(0, 40), ref = String(p.ref || '').trim().slice(0, 60);

  const html = construireHtmlRecu_(d, calc, L.no, datePmt, mode, ref, c, logoClub_(c));
  const blob = Utilities.newBlob(html, 'text/html', 'recu.html').getAs('application/pdf')
    .setName('Recu_' + L.no + '_' + slug_(d.commanditaire) + '.pdf');
  const cRecu = colFactures_(sh, L.m, 'Reçu PDF');
  remplacerFichier_(sh.getRange(p.r, cRecu));
  const file = sousDossier_(sousDossier_(c.DOSSIER_ID, 'Reçus').getId(), d.equipe).createFile(blob);

  // Onglet Factures
  lienRiche_(sh.getRange(p.r, cRecu), file.getUrl(), 'Reçu ' + L.no);
  sh.getRange(p.r, colFactures_(sh, L.m, 'Paiement reçu')).setValue(datePmt).setNumberFormat('yyyy-mm-dd');
  sh.getRange(p.r, colFactures_(sh, L.m, 'Mode de paiement')).setValue(sur_(mode + (ref ? ' #' + ref : '')));
  sh.getRange(p.r, colFactures_(sh, L.m, 'Reçu émis le')).setValue(new Date()).setNumberFormat('yyyy-mm-dd');

  // Onglet d'équipe
  const eq = trouverLigneEquipe_(d.equipe, L.no);
  if (eq) {
    if (eq.C.pmtRecu) eq.sh.getRange(eq.r, eq.C.pmtRecu).setValue(datePmt).setNumberFormat('yyyy-mm-dd');
    if (eq.C.recuEmis) lienRiche_(eq.sh.getRange(eq.r, eq.C.recuEmis), file.getUrl(), dateCourte_(new Date()));
    if (!eq.C.pmtRecu && eq.C.note) {
      const n = eq.sh.getRange(eq.r, eq.C.note);
      const actuel = String(n.getValue() || '');
      if (!/payé/i.test(actuel)) n.setValue(sur_((actuel ? actuel + ' — ' : '') + 'payé ' + dateCourte_(datePmt)));
    }
  }

  let envoi = 'non envoyé';
  if (p.envoyer) {
    const test = modeTest_(c);
    const to = test ? (c.COURRIEL_TEST || d.parentCourriel) : String(p.dest).trim();
    const cc = (!test && p.copieParent && d.parentCourriel && d.parentCourriel !== to) ? d.parentCourriel : '';
    MailApp.sendEmail({
      to: to, cc: cc,
      replyTo: c['COURRIEL_GERANTE_' + d.equipe] || c.COURRIEL_ORG,
      name: c.NOM_ORG,
      subject: (test ? '[TEST] ' : '') + 'Reçu de paiement — facture ' + L.no + ' — ' + d.commanditaire,
      htmlBody: '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.5">Bonjour,<br><br>' +
        'Nous confirmons la réception de votre paiement de <b>' + argent_(calc.total) + '</b> pour la facture <b>' + L.no + '</b>. ' +
        'Vous trouverez le reçu en pièce jointe.<br><br>' + esc_(c.MESSAGE_MERCI) + '<br><br>' + esc_(c.NOM_ORG) + '</div>',
      attachments: [blob]
    });
    envoi = 'envoyé à ' + to + (cc ? ' (copie ' + cc + ')' : '');
  }
  return {
    message: 'Reçu ' + L.no + ' émis ✅\nCourriel : ' + envoi + (eq ? '\nOnglet ' + d.equipe + ' mis à jour.' : ''),
    envoi: envoi, no: L.no,
    nomFichier: blob.getName(),
    pdfBase64: Utilities.base64Encode(blob.getBytes())
  };
}

function construireHtmlRecu_(d, calc, no, datePmt, mode, ref, c, logoDataUri) {
  const e = esc_;
  const adresse = [d.adresse, [d.ville, d.province, d.codePostal].filter(Boolean).join(' ')].filter(Boolean).join(', ');
  const taxes = (c.NO_TPS || c.NO_TVQ)
    ? '<div class="taxnos">TPS : ' + e(c.NO_TPS || '—') + ' &nbsp;·&nbsp; TVQ : ' + e(c.NO_TVQ || '—') + '</div>' : '';
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><style>' + cssDoc_() +
    '.paye{display:inline-block;border:3px solid #127a43;color:#127a43;font-size:20pt;font-weight:bold;padding:4px 18px;letter-spacing:3px;margin-top:6px;}' +
    '</style></head><body>' +
    (modeTest_(c) ? '<div class="test">REÇU TEST — NON VALIDE</div>' : '') +
    '<table><tr><td style="vertical-align:top">' +
      (logoDataUri ? '<img src="' + logoDataUri + '" style="height:70px;margin-bottom:6px"><br>' : '') +
      '<div class="org">' + e(c.NOM_ORG) + '</div>' +
      '<div class="muted">' + e(c.ADRESSE_ORG) + '<br>' + e(c.COURRIEL_ORG) + ' &nbsp;|&nbsp; ' + e(c.SITE_ORG) + '</div>' + taxes +
    '</td><td style="vertical-align:top;width:240px">' +
      '<div class="titre">REÇU</div>' +
      '<table class="box" style="margin-top:8px"><tr><td class="lbl">Facture n°</td><td class="r"><b>' + e(no) + '</b></td></tr>' +
      '<tr><td class="lbl">Payé le</td><td class="r">' + dateFr_(datePmt) + '</td></tr>' +
      '<tr><td class="lbl">Équipe</td><td class="r">' + e(d.equipe) + '</td></tr></table>' +
    '</td></tr></table>' +
    '<div class="bar"></div>' +
    '<div class="h">Reçu de</div>' +
    '<table class="box"><tr><td class="lbl">Entreprise</td><td><b>' + e(d.commanditaire) + '</b></td></tr>' +
      (adresse ? '<tr><td class="lbl">Adresse</td><td>' + e(adresse) + '</td></tr>' : '') + '</table>' +
    '<table class="items" style="margin-top:22px"><tr><th>Description</th><th class="r" style="width:110px">Montant</th></tr>' +
      '<tr><td>' + e(d.description) + '<br><span class="muted">Joueur : ' + e(d.joueur) + '</span></td><td class="r">' + argent_(calc.sousTotal) + '</td></tr></table>' +
    '<table class="tot" style="width:300px;margin:14px 0 0 auto">' +
      '<tr><td>Sous-total</td><td class="r">' + argent_(calc.sousTotal) + '</td></tr>' +
      '<tr><td>TPS (5 %)</td><td class="r">' + argent_(calc.tps) + '</td></tr>' +
      '<tr><td>TVQ (9,975 %)</td><td class="r">' + argent_(calc.tvq) + '</td></tr>' +
      '<tr class="grand"><td>TOTAL PAYÉ</td><td class="r">' + argent_(calc.total) + '</td></tr></table>' +
    '<table style="margin-top:22px"><tr><td style="vertical-align:middle">' +
      '<div class="pay" style="margin-top:0"><b>Paiement reçu</b><br>Mode : ' + e(mode) + (ref ? ' — Réf. ' + e(ref) : '') +
      '<br>Date : ' + dateFr_(datePmt) + '</div></td>' +
      '<td style="width:170px;text-align:center;vertical-align:middle"><div class="paye">PAYÉ</div></td></tr></table>' +
    '<div class="muted" style="margin-top:18px;font-size:8.5pt">Ce reçu confirme le paiement de la facture n° ' + e(no) +
      '. Il ne constitue pas un reçu officiel de don aux fins de l’impôt.</div>' +
    '<div class="merci">' + e(c.MESSAGE_MERCI) + '</div></body></html>';
}

/* --- helpers gérante --- */

function ligneSelectionnee_() {
  const sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== FEUILLE_FACTURES) throw new Error('Va dans l’onglet « Factures » et clique sur la ligne de la facture.');
  const r = sh.getActiveRange().getRow();
  if (r < 2) throw new Error('Clique sur une ligne de facture (pas sur la ligne de titres).');
  return { sh: sh, r: r };
}

/** Index des colonnes de l'onglet Factures par titre (normalisé). */
function entetesFactures_(sh) {
  const m = {};
  sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].forEach((v, i) => { if (v) m[norm_(v)] = i + 1; });
  return m;
}

/** Trouve une colonne de Factures, ou la crée à droite si elle n'existe pas encore. */
function colFactures_(sh, m, nom) {
  const k = norm_(nom);
  if (!m[k]) {
    const col = sh.getLastColumn() + 1;
    sh.getRange(1, col).setValue(nom).setFontWeight('bold').setBackground('#0b2545').setFontColor('#ffffff');
    m[k] = col;
  }
  return m[k];
}

function lireLigne_(sh, r) {
  const m = entetesFactures_(sh);
  const row = sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0];
  const get = nom => { const col = m[norm_(nom)]; return col ? row[col - 1] : ''; };
  const s = nom => String(get(nom) == null ? '' : get(nom)).trim();
  let pct = Number(get('% équipe'));
  if (pct > 0 && pct <= 1) pct = pct * 100;
  const df = get('Date facture');
  return {
    m: m, get: get, no: s('N° facture'),
    date: df instanceof Date ? df : new Date(),
    d: {
      equipe: s('Équipe'), joueur: s('Joueur'),
      parentNom: s('Parent'), parentCourriel: s('Courriel parent').toLowerCase(), parentTel: s('Tél. parent'),
      commanditaire: s('Commanditaire'), contact: s('Personne contact'),
      adresse: s('Adresse'), ville: s('Ville'), province: s('Province') || 'QC', codePostal: s('Code postal').toUpperCase(),
      telephone: s('Tél. commanditaire'), courriel: s('Courriel commanditaire').toLowerCase(), siteWeb: s('Site web'),
      description: s('Description'), notes: s('Notes'),
      montant: Number(String(get('Montant total (taxes incl.)')).replace(/[^\d.,-]/g, '').replace(',', '.')),
      pctEquipe: Math.round(pct)
    }
  };
}

function validerLigne_(L, c) {
  const d = L.d, err = [];
  if (!L.no) err.push('N° de facture manquant.');
  if (equipes_(c).indexOf(d.equipe) === -1) err.push('Équipe invalide : « ' + d.equipe + ' ».');
  if (!d.commanditaire) err.push('Commanditaire manquant.');
  if (!(d.montant > 0)) err.push('Montant invalide.');
  const min = Number(c.POURCENT_EQUIPE_MIN) || 30;
  if (d.montant > seuil_(c) && !(d.pctEquipe >= min && d.pctEquipe <= 100)) err.push('% équipe doit être entre ' + min + ' % et 100 %.');
  if (d.parentCourriel && !courrielValide_(d.parentCourriel)) err.push('Courriel du parent invalide.');
  if (d.courriel && !courrielValide_(d.courriel)) err.push('Courriel du commanditaire invalide.');
  if (err.length) throw new Error(err.join(' '));
}

/** Met l'ancien PDF lié dans la cellule à la corbeille (ou le renomme si on n'en est pas propriétaire). */
function remplacerFichier_(cellule) {
  try {
    const rt = cellule.getRichTextValue();
    const url = rt && rt.getLinkUrl();
    const id = url && (url.match(/\/d\/([\w-]+)/) || url.match(/[?&]id=([\w-]+)/) || [])[1];
    if (!id) return;
    const f = DriveApp.getFileById(id);
    try { f.setTrashed(true); } catch (e) { f.setName('REMPLACÉ_' + f.getName()); }
  } catch (e) { console.warn('Ancien fichier non remplacé', e); }
}

/* ============================================================
 * Utilitaires
 * ============================================================ */

function classeur_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function lireConfig_() {
  const sh = classeur_().getSheetByName(FEUILLE_CONFIG);
  if (!sh) throw new Error('Onglet Config introuvable — lance « Commandites > Configurer » dans le Google Sheet.');
  const o = {};
  CONFIG_DEFAUT.forEach(r => { o[r[0]] = String(r[1]); }); // valeurs par défaut si une clé manque
  sh.getDataRange().getValues().slice(1).forEach(r => { if (r[0]) o[String(r[0]).trim()] = String(r[1]).trim(); });
  return o;
}

function ecrireConfig_(cle, valeur) {
  const sh = classeur_().getSheetByName(FEUILLE_CONFIG);
  const v = sh.getDataRange().getValues();
  for (let i = 0; i < v.length; i++) if (v[i][0] === cle) { sh.getRange(i + 1, 2).setValue(valeur); return; }
  sh.appendRow([cle, valeur, '']);
}

function seuil_(c) { const n = Number(String(c.SEUIL_100_EQUIPE).replace(',', '.')); return isNaN(n) ? 250 : n; }
function modeTest_(c) { return /^(oui|o|yes|true|1)$/i.test(String(c.MODE_TEST).trim()); }

function equipes_(c) { return String(c.EQUIPES || '').split(',').map(s => s.trim()).filter(Boolean); }

function joueurs_(sh) {
  if (!sh || sh.getLastRow() < 2) return [];
  const vus = {};
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues()
    .map(r => String(r[0]).trim())
    .filter(n => n && !vus[n] && (vus[n] = true))
    .sort((a, b) => a.localeCompare(b, 'fr'));
}

function dossier_(id, creer) {
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* recréé */ } }
  return creer();
}

function sousDossier_(parentId, nom) {
  const p = DriveApp.getFolderById(parentId);
  const it = p.getFoldersByName(nom);
  return it.hasNext() ? it.next() : p.createFolder(nom);
}

function r2_(n) { return Math.round((Number(n) + Number.EPSILON) * 100) / 100; }

function argent_(n) {
  const [ent, dec] = r2_(n).toFixed(2).split('.');
  return ent.replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ',' + dec + ' $';
}

const MOIS_ = ['janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet', 'août', 'septembre', 'octobre', 'novembre', 'décembre'];
function dateFr_(d) { return d.getDate() + ' ' + MOIS_[d.getMonth()] + ' ' + d.getFullYear(); }
function dateCourte_(d) { const p = n => String(n).padStart(2, '0'); return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate()); }

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[ch]));
}

/** Empêche l'injection de formules dans le Sheet (texte qui commence par = + - @). */
function sur_(v) { v = String(v == null ? '' : v); return /^[=+\-@]/.test(v) ? "'" + v : v; }

/** Lien cliquable sans formule (évite #ERROR! selon la langue du Sheet). */
function lienRiche_(range, url, texte) {
  range.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(String(texte)).setLinkUrl(url).build());
}

function norm_(s) {
  return String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().replace(/\s+/g, ' ').trim();
}

function slug_(s) { return norm_(s).replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40) || 'commanditaire'; }

function nomFichier_(no, commanditaire) { return 'Facture_' + no + '_' + slug_(commanditaire) + '.pdf'; }

function courrielValide_(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(s); }
