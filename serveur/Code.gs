/**
 * ═══════════════════════════════════════════════════════════
 *   ICONIC COLLECTION — Gestion des abonnements
 *   Fichier : Code.gs
 * ═══════════════════════════════════════════════════════════
 *   1. Colle ce code dans le fichier « Code.gs »
 *   2. Ajoute un fichier HTML nommé « Index » et colle le 2e code
 *   3. Exécute la fonction « installer » (une seule fois)
 *   4. Déploie en « Application Web » (Exécuter en tant que : Moi,
 *      Accès : Tout le monde), puis ouvre le lien du déploiement
 * ═══════════════════════════════════════════════════════════
 */

const VERSION = '2.3.4';
const MARQUE = 'ICONIC Collection';
const TZ = 'Africa/Ouagadougou';

const S = {
  PARAM: 'Paramètres', TYPES: 'Types', SESSIONS: 'Sessions', CARTES: 'Cartes',
  COMPTES: 'Comptes', PLACES: 'Places', CLIENTS: 'Clients', ABOS: 'Abonnements',
  VENTES: 'Ventes', PAIEMENTS: 'Paiements comptes', HIST: 'Historique', JOURNAL: 'Journal'
};

const H = {};
H[S.PARAM]     = ['Clé', 'Valeur', 'Description'];
H[S.TYPES]     = ['Type', 'Places par défaut', 'Coût mensuel par défaut'];
H[S.SESSIONS]  = ['Code', 'Type', 'Libellé', 'Durée (mois)', 'Prix'];
H[S.CARTES]    = ['ID', 'Repère', 'Titulaire', 'Statut', 'Remplacée par', 'Date ajout', 'Notes', 'Supprimée le', 'Statut avant'];
H[S.COMPTES]   = ['ID', 'Type', 'Mail', 'Mot de passe', 'Nom', 'Téléphone', 'Carte', 'Coût mensuel', 'Places', 'Expiration', 'Statut', 'Créé le', 'Bloqué le', 'Lien', 'Adresse'];
H[S.PLACES]    = ['ID', 'Compte', 'N°', 'Profil', 'Code', 'État', 'Client', 'Abonnement', 'Depuis'];
H[S.CLIENTS]   = ['ID', 'Nom', 'Numéros', 'Créé le', 'Notes', 'WhatsApp', 'Statut', 'Supprimé le', 'Relancé le'];
H[S.ABOS]      = ['ID', 'Client', 'Type', 'Session', 'Début', 'Expiration', 'Statut', 'Place', 'Compte', 'Mis à jour', 'Import'];
H[S.VENTES]    = ['ID', 'Date', 'Trans ID', 'Numéro', 'Nom SMS', 'Montant reçu', 'Sessions possibles', 'Session', 'Prix', 'Client', 'Abonnement', 'Statut', 'Genre', 'SMS', 'Validée le', 'Ancienne expiration', 'Ancien statut', 'Ancienne session', 'Auto'];
H[S.PAIEMENTS] = ['ID', 'Date', 'Compte', 'Carte', 'Montant', 'Nouvelle expiration', 'Ancienne expiration', 'Statut'];
H[S.HIST]      = ['Date', 'Objet', 'ID', 'Action', 'Détail'];
H[S.JOURNAL]   = ['Date', 'Niveau', 'Message'];

// Colonnes à garder en texte (zéros au début, numéros…)
const TEXTE = {};
TEXTE[S.COMPTES] = ['Mail', 'Mot de passe', 'Téléphone', 'Lien', 'Adresse'];
TEXTE[S.PLACES]  = ['Profil', 'Code'];
TEXTE[S.CLIENTS] = ['Numéros', 'WhatsApp'];
TEXTE[S.VENTES]  = ['Numéro', 'Trans ID'];
TEXTE[S.CARTES]  = ['Repère'];

const EN_COURS = ['Actif', 'Grâce', 'À retirer', 'En attente', 'Dette'];
const MOIS = ['janvier','février','mars','avril','mai','juin','juillet','août','septembre','octobre','novembre','décembre'];

/* ═════════════════════ INSTALLATION ═════════════════════ */

function installer() {
  const props = props_();
  let ss;
  if (props.getProperty('SS_ID')) {
    ss = SpreadsheetApp.openById(props.getProperty('SS_ID'));
  } else {
    ss = SpreadsheetApp.create('ICONIC Collection – Gestion');
    props.setProperty('SS_ID', ss.getId());
  }
  ss.setSpreadsheetTimeZone(TZ);
  SS_CACHE = ss;

  Object.keys(H).forEach(function (nom) {
    let sh = ss.getSheetByName(nom);
    if (!sh) sh = ss.insertSheet(nom);
    const h = H[nom];
    sh.getRange(1, 1, 1, h.length).setValues([h])
      .setFontWeight('bold').setBackground('#0b0b0b').setFontColor('#d4af37');
    sh.setFrozenRows(1);
  });
  ['Feuille 1', 'Sheet1', 'Feuille1'].forEach(function (n) {
    const f = ss.getSheetByName(n);
    if (f && ss.getSheets().length > 1) ss.deleteSheet(f);
  });

  // Paramètres (seulement ceux qui manquent)
  const existants = rows_(S.PARAM).map(function (r) { return r['Clé']; });
  [
    ['EMAIL', Session.getEffectiveUser().getEmail(), 'Adresse qui reçoit les notifications'],
    ['GRACE_JOURS', 3, 'Jours de grâce après expiration avant de retirer un client'],
    ['RAPPEL_CLIENT_JOURS', 3, 'Prévenir X jours avant l\'expiration d\'un client'],
    ['RAPPEL_COMPTE_JOURS', 5, 'Prévenir X jours avant l\'expiration d\'un compte à payer'],
    ['MARGE', 100, 'Montant accepté en plus du prix (frais ajoutés par le client)'],
    ['MAX_COMPTES_PAR_CARTE', 3, 'Nombre maximum de comptes reliés à une même carte'],
    ['INDICATIF', '226', 'Indicatif pays pour WhatsApp'],
    ['PREPARER_RAPPEL_JOURS', 1, 'Rappel si une place reste « À préparer » depuis X jours'],
    ['NUMERO_DEPOT', '', 'Ton numéro Orange Money de dépôt (affiché dans les rappels clients)'],
    ['NOTIF_EMAIL', 'oui', 'Recevoir aussi les alertes par e-mail (oui / non)'],
    ['VALIDATION_AUTO', 'oui', 'Valider seul les renouvellements sans ambiguïté (oui / non)'],
    ['USSD_PAIEMENT', '*144*2*1*{numero}*{montant}#', 'Code Orange Money pour te payer ({numero} et {montant} sont remplacés)'],
    ['ONESIGNAL_APP_ID', '', 'OneSignal : App ID (notifications de l\'app)'],
    ['ONESIGNAL_API_KEY', '', 'OneSignal : clé API REST (à garder secrète)']
  ].forEach(function (p) {
    if (existants.indexOf(p[0]) < 0) insert_(S.PARAM, { 'Clé': p[0], 'Valeur': p[1], 'Description': p[2] });
  });

  const nd = rows_(S.PARAM).find(function (r) { return r['Clé'] === 'NUMERO_DEPOT'; });
  if (nd && !String(nd.Valeur).trim()) { nd.Valeur = "'66290308"; update_(S.PARAM, nd); }
  if (rows_(S.TYPES).length === 0) {
    [['Films et séries', 5, 6900], ['Musique et podcasts', 5, 4500], ['PRM', 6, 6000]].forEach(function (t) {
      insert_(S.TYPES, { 'Type': t[0], 'Places par défaut': t[1], 'Coût mensuel par défaut': t[2] });
    });
  }
  if (rows_(S.SESSIONS).length === 0) {
    [
      ['FS-TEL-1M', 'Films et séries', 'Films et séries · Téléphone 1 mois', 1, 2500],
      ['FS-TEL-3M', 'Films et séries', 'Films et séries · Téléphone 3 mois', 3, 7000],
      ['FS-TEL-12M', 'Films et séries', 'Films et séries · Téléphone 1 an', 12, 25000],
      ['FS-TV-1M', 'Films et séries', 'Films et séries · Télévision 1 mois', 1, 3500],
      ['FS-TV-3M', 'Films et séries', 'Films et séries · Télévision 3 mois', 3, 9000],
      ['MU-1M-A', 'Musique et podcasts', 'Musique et podcasts · 1 mois', 1, 1600],
      ['MU-1M-B', 'Musique et podcasts', 'Musique et podcasts · 1 mois', 1, 2000],
      ['MU-3M-A', 'Musique et podcasts', 'Musique et podcasts · 3 mois', 3, 4000],
      ['MU-3M-B', 'Musique et podcasts', 'Musique et podcasts · 3 mois', 3, 5000],
      ['PRM-1M', 'PRM', 'PRM · 1 mois', 1, 3500],
      ['PRM-3M', 'PRM', 'PRM · 3 mois', 3, 9000]
    ].forEach(function (s) {
      insert_(S.SESSIONS, { 'Code': s[0], 'Type': s[1], 'Libellé': s[2], 'Durée (mois)': s[3], 'Prix': s[4] });
    });
  }

  if (!props.getProperty('APP_KEY')) props.setProperty('APP_KEY', cle_(32));
  if (!props.getProperty('SMS_TOKEN')) props.setProperty('SMS_TOKEN', cle_(24));

  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'rappelQuotidien') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('rappelQuotidien').timeBased().everyDays(1).atHour(7).inTimezone(TZ).create();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sauvegardeMensuelle') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sauvegardeMensuelle').timeBased().onMonthDay(1).atHour(6).inTimezone(TZ).create();

  Logger.log('✅ Installation terminée. Tableau : ' + ss.getUrl());
  Logger.log('➡️ Étape suivante : Déployer > Nouveau déploiement > Application Web.');
}

/** À lancer depuis l'éditeur si tu as perdu tes liens : ils te sont renvoyés par e-mail. */
function envoyerMesLiens() {
  const l = liens_();
  MailApp.sendEmail({
    to: params_().email, name: MARQUE, subject: '🔐 Tes liens ' + MARQUE,
    htmlBody: emailHtml_('Tes liens personnels', [{ titre: 'À garder secrets', lignes: [
      '<b>Application :</b><br><a style="color:#f3dfa2" href="' + l.app + '">' + l.app + '</a>',
      '<b>Adresse pour l\'application SMS (Android) :</b><br>' + l.sms
    ] }], false)
  });
  Logger.log('Liens envoyés à ' + params_().email);
}

/** Test : simule la réception d'un SMS Orange Money. */
function testSms() {
  const r = traiterSms_('Vous avez recu 2525.0 FCFA, Frais:  FCFA, Taxe:  FCFA du 70000000,CLIENT TEST. Le solde de votre compte est de 10000.00 FCFA Trans ID: TEST' + Date.now() + '. Flashez le QR CODE');
  Logger.log(JSON.stringify(r));
}

/* ═════════════════════ RÉCEPTION DES SMS ═════════════════════ */

function doPost(e) {
  const token = (e && e.parameter && e.parameter.token) || '';
  if (!token) return appelApp_(e);
  if (token !== prop_('SMS_TOKEN')) return json_({ ok: false, erreur: 'accès refusé' });
  const brut = (e.postData && e.postData.contents) || '';
  let texte = '';
  try { texte = chercherTexte_(JSON.parse(brut)); } catch (err) { /* pas du JSON */ }
  if (!texte) texte = (e.parameter && (e.parameter.text || e.parameter.message || e.parameter.msg)) || brut;
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return json_(traiterSms_(texte));
  } catch (err) {
    journal_('ERREUR', 'SMS : ' + err.message);
    return json_({ ok: false, erreur: err.message });
  } finally {
    lock.releaseLock();
  }
}

/** Appels de l'application ICONIC installée sur l'écran d'accueil (version GitHub). */
function appelApp_(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); } catch (err) { /* vide */ }
  try {
    return json_({ ok: true, data: api(req.key, req.action, req.data) });
  } catch (err) {
    return json_({ ok: false, erreur: err.message || String(err) });
  }
}

function chercherTexte_(o) {
  if (typeof o === 'string') return /re[çc]u/i.test(o) ? o : '';
  if (o && typeof o === 'object') {
    for (const k in o) { const t = chercherTexte_(o[k]); if (t) return t; }
  }
  return '';
}

function parseSms_(texte) {
  const t = String(texte || '').replace(/\s+/g, ' ');
  const m = t.match(/Vous avez re[çc]u\s*([\d\s.,]+?)\s*FCFA/i);
  if (!m) return null;
  const montant = parseFloat(m[1].replace(/\s/g, '').replace(',', '.'));
  const f = t.match(/\bdu\s*(\+?\d[\d\s]{6,}?)\s*,\s*([^.]*)\./i);
  const id = t.match(/Trans\s*ID\s*:\s*(\S+)/i);
  const transId = id ? id[1].replace(/\.+$/, '') : '';
  return {
    date: dateTrans_(transId),
    montant: montant,
    numero: f ? normNum_(f[1]) : '',
    nom: f ? f[2].trim() : '',
    transId: id ? id[1].replace(/\.+$/, '') : ''
  };
}

function traiterSms_(texte) {
  const p = parseSms_(texte);
  if (!p || !p.montant) {
    journal_('INFO', 'SMS ignoré : ' + String(texte).slice(0, 150));
    return { ok: false, raison: 'SMS non reconnu' };
  }
  if (p.transId && rows_(S.VENTES).some(function (v) { return String(v['Trans ID']) === p.transId; })) {
    return { ok: true, raison: 'déjà enregistré' };
  }
  const P = params_();
  const toutesS = sessions_();
  let cands = candidats_(p.montant, toutesS, P.marge);
  const combos = combinaisons_(p.montant, toutesS, P.marge);
  const codesCands = cands.map(function (s) { return s.code; }).concat(combos);
  const cl = clientParNumero_(rows_(S.CLIENTS), p.numero);
  const v = {
    'ID': nextId_('V'), 'Date': p.date || now_(), 'Trans ID': p.transId, 'Numéro': p.numero, 'Nom SMS': p.nom,
    'Montant reçu': p.montant, 'Sessions possibles': codesCands.join(','),
    'Session': '', 'Prix': '', 'Client': cl ? cl.ID : '', 'Abonnement': '', 'Statut': codesCands.length ? 'À valider' : 'Hors tarif',
    'Genre': '', 'SMS': String(texte).slice(0, 500)
  };
  insert_(S.VENTES, v);
  // Montant qui ne correspond à aucun tarif : gardé sans notification, Moïse le saisira lui-même en vente manuelle
  if (!codesCands.length) { journal_('INFO', 'Montant hors tarif (non signalé) : ' + p.montant + ' F de ' + (cl ? cl.Nom : p.numero)); return { ok: true, id: v.ID, horsTarif: true }; }
  // Validation automatique : client reconnu, une seule session possible, un seul abonnement de ce type à prolonger
  if (P.validationAuto && cl && cands.length === 1 && !combos.length) {
    const enc = rows_(S.ABOS).filter(function (a) { return a.Client === cl.ID && a.Type === cands[0].type && EN_COURS.indexOf(a.Statut) >= 0 && a.Statut !== 'En attente'; });
    if (enc.length === 1 && enc[0].Place) {
      try {
        const r = validerVente_({ venteId: v.ID, session: cands[0].code, clientId: cl.ID, aboId: enc[0].ID, auto: true });
        push_('✅ ' + cl.Nom + ' renouvelé(e)', cands[0].libelle + ' · jusqu\'au ' + fmt_(r.expiration) + ' · touche pour voir ou annuler', 'accueil', P);
        if (P.notifEmail) { try { MailApp.sendEmail({ to: P.email, name: MARQUE, subject: '✅ ' + cl.Nom + ' renouvelé(e) automatiquement', htmlBody: emailHtml_('Renouvellement automatique', [{ titre: esc_(cl.Nom), lignes: [esc_(cands[0].libelle) + ' · ' + F_(p.montant), 'Valable jusqu\'au ' + fmt_(r.expiration), 'Tu peux l\'annuler dans l\'app pendant 24 h.'] }], true, 'accueil') }); } catch (e) {} }
        return { ok: true, id: v.ID, auto: true };
      } catch (err) { journal_('ERREUR', 'Validation auto : ' + err.message); }
    }
  }
  if (P.notifEmail) { try { notifierVente_(v, cands, cl, P); } catch (err) { journal_('ERREUR', 'E-mail : ' + err.message); } }
  push_('💰 ' + F_(p.montant) + ' reçus', (cl ? cl.Nom : (p.nom || p.numero)) + ' · touche pour valider', 'valider', P);
  return { ok: true, id: v.ID };
}

function notifierVente_(v, cands, cl, P) {
  const lignes = [
    '<b>Montant :</b> ' + F_(v['Montant reçu']),
    '<b>Numéro :</b> ' + esc_(v['Numéro']) + ' · <b>Nom Orange Money :</b> ' + esc_(v['Nom SMS']),
    cl ? '✅ <b>Client reconnu :</b> ' + esc_(cl.Nom) : '🆕 <b>Numéro inconnu</b> : nouveau client ou client existant ?',
    cands.length === 1 ? '<b>Session :</b> ' + esc_(cands[0].libelle)
      : cands.length > 1 ? '<b>Session à choisir :</b> ' + cands.map(function (s) { return esc_(s.libelle); }).join(' ou ')
      : '⚠️ <b>Montant qui ne correspond à aucune session</b>'
  ];
  MailApp.sendEmail({
    to: P.email, name: MARQUE,
    subject: '💰 ' + F_(v['Montant reçu']) + ' reçus de ' + (cl ? cl.Nom : v['Nom SMS']) + ' – à valider',
    htmlBody: emailHtml_('Nouveau paiement à valider', [{ titre: 'Paiement Orange Money', lignes: lignes }], true, 'valider')
  });
}

/* ═════════════════════ APPLICATION WEB ═════════════════════ */

function doGet(e) {
  const props = props_();
  const k = (e && e.parameter && e.parameter.k) || '';
  if (!props.getProperty('SS_ID')) {
    return page_('Installation incomplète', '<p>Exécute d\'abord la fonction <b>installer</b> dans l\'éditeur.</p>');
  }
  if (!props.getProperty('SETUP_DONE')) return pageInstallation_();
  if (k !== props.getProperty('APP_KEY')) return page_('Accès refusé', '<p>Ce lien n\'est pas valide.</p>');
  const t = HtmlService.createTemplateFromFile('Index');
  t.KEY = k;
  t.TAB = (e.parameter.tab || '').replace(/[^a-z]/g, '');
  return t.evaluate()
    .setTitle(MARQUE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function pageInstallation_() {
  const props = props_();
  props.setProperty('APP_URL', ScriptApp.getService().getUrl());
  const l = liens_();
  props.setProperty('SETUP_DONE', '1');
  try { envoyerMesLiens(); } catch (err) { journal_('ERREUR', 'Envoi des liens : ' + err.message); }
  return page_('Installation terminée ✨',
    '<p>Voici tes <b>deux liens personnels</b>. Ils t\'ont aussi été envoyés par e-mail. <b>Ne les partage avec personne.</b></p>' +
    '<p><b>1. Ton application</b> (à ouvrir sur ton iPhone) :</p><textarea readonly onclick="this.select()">' + l.app + '</textarea>' +
    '<p><a class="b" href="' + l.app + '" target="_top">Ouvrir l\'application</a></p>' +
    '<p><b>2. L\'adresse pour l\'application SMS</b> (à mettre sur l\'Android) :</p><textarea readonly onclick="this.select()">' + l.sms + '</textarea>' +
    '<p class="m">Cette page ne s\'affichera plus. En cas de perte, exécute la fonction « envoyerMesLiens » dans l\'éditeur.</p>');
}

function page_(titre, corps) {
  return HtmlService.createHtmlOutput(
    '<meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<style>body{background:#070707;color:#eee;font-family:Arial,sans-serif;padding:24px;line-height:1.5}' +
    'h1{font-family:Georgia,serif;color:#f3dfa2;letter-spacing:2px;font-size:22px}' +
    'textarea{width:100%;height:90px;background:#111;color:#f3dfa2;border:1px solid #6b5a22;border-radius:8px;padding:8px;font-size:13px}' +
    '.b{display:inline-block;background:#d4af37;color:#111;padding:12px 18px;border-radius:8px;text-decoration:none;font-weight:bold}.m{color:#999;font-size:13px}</style>' +
    '<div style="text-align:center;font-family:Georgia,serif;color:#f3dfa2;letter-spacing:5px;font-size:24px">◆ ICONIC<div style="font-size:10px;letter-spacing:7px;color:#d4af37">COLLECTION</div></div>' +
    '<h1>' + titre + '</h1>' + corps
  ).setTitle(MARQUE).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Point d'entrée unique appelé par l'application. */
function api(key, action, d) {
  if (!key || key !== prop_('APP_KEY')) throw new Error('Accès refusé');
  d = d || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    RCACHE = {};
    let r;
    switch (action) {
      case 'load': r = charger_(); break;
      case 'valider': r = validerVente_(d); break;
      case 'ignorer': r = ignorerVente_(d); break;
      case 'manuelle': r = venteManuelle_(d); break;
      case 'liberer': r = liberer_(d); break;
      case 'preparer': r = preparerPlace_(d); break;
      case 'ajouterCompte': r = ajouterCompte_(d); break;
      case 'modifierCompte': r = modifierCompte_(d); break;
      case 'fermerCompte': r = fermerCompte_(d); break;
      case 'payerCompte': r = payerCompte_(d); break;
      case 'ajouterCarte': r = ajouterCarte_(d); break;
      case 'modifierCarte': r = modifierCarte_(d); break;
      case 'remplacerCarte': r = remplacerCarte_(d); break;
      case 'modifierClient': r = modifierClient_(d); break;
      case 'annulerVente': r = annulerVente_(d); break;
      case 'reprendreVente': r = reprendreVente_(d); break;
      case 'annulerLiberation': r = annulerLiberation_(d); break;
      case 'annulerPaiement': r = annulerPaiementCompte_(d); break;
      case 'annulerAttente': r = annulerAttente_(d); break;
      case 'supprimerClient': r = supprimerClient_(d); break;
      case 'restaurerClient': r = restaurerClient_(d); break;
      case 'supprimerCarte': r = supprimerCarte_(d); break;
      case 'restaurerCarte': r = restaurerCarte_(d); break;
      case 'relance': r = marquerRelance_(d); break;
      case 'bloquer': r = bloquerCompte_(d); break;
      case 'sauvegarde': r = sauvegardeMensuelle(); break;
      case 'app': r = enregistrerApp_(d); break;
      case 'inventaire': r = inventaire_(d); break;
      case 'annulerInventaire': r = annulerInventaire_(d); break;
      case 'inventaireAuto': r = inventaireAuto_(); break;
      case 'completerCompte': r = completerCompte_(d); break;
      case 'reglages': r = lireReglages_(); break;
      case 'sauverReglages': r = sauverReglages_(d); break;
      case 'testNotif': r = testNotif_(); break;
      case 'whatsappLot': r = whatsappLot_(d); break;
      case 'modifierMontant': r = modifierMontant_(d); break;
      case 'version': r = infoVersion_(); break;
      case 'miseAJour': r = miseAJour_(true); break;
      default: throw new Error('Action inconnue : ' + action);
    }
    if (d._load && action !== 'load') r = { r: r === undefined ? null : r, load: charger_() };
    return JSON.parse(JSON.stringify(r === undefined ? null : r));
  } finally {
    lock.releaseLock();
  }
}

/* ═════════════════════ CHARGEMENT ═════════════════════ */

function charger_() {
  const P = params_();
  const dg = digest_(P);
  const sess = sessions_();
  const clients = rows_(S.CLIENTS), abos = rows_(S.ABOS);
  const aValider = rows_(S.VENTES).filter(function (v) { return v.Statut === 'À valider'; }).map(function (v) {
    const cands = String(v['Sessions possibles'] || '').split(',').filter(String);
    const cl = v.Client ? clients.find(function (c) { return c.ID === v.Client; }) : clientParNumero_(clients, v['Numéro']);
    let pre = cands.length === 1 ? cands[0] : '';
    if (cl && cands.length > 1) {
      const types = abos.filter(function (a) { return a.Client === cl.ID && EN_COURS.indexOf(a.Statut) >= 0; })
        .map(function (a) { return a.Type; });
      const m = cands.filter(function (c) { const s = sess.find(function (x) { return x.code === c; }); return s && types.indexOf(s.type) >= 0; });
      if (m.length === 1) pre = m[0];
    }
    const sugg = cl ? [] : clients.map(function (c) { return { id: c.ID, s: score_(c.Nom, v['Nom SMS']) }; })
      .filter(function (x) { return x.s > 0; }).sort(function (a, b) { return b.s - a.s; }).slice(0, 5)
      .map(function (x) { return x.id; });
    return { id: v.ID, date: v.Date, numero: String(v['Numéro']), nom: v['Nom SMS'], montant: v['Montant reçu'],
      candidats: cands, preselection: pre, clientId: cl ? cl.ID : '', suggestions: sugg };
  });
  const hist = rows_(S.HIST).filter(function (h) { return h.Objet === 'Carte' || h.Action === 'Carte changée'; }).slice(-50).reverse();
  return {
    aujourdhui: today_(),
    version: VERSION, maj: infoVersion_(),
    ssUrl: ss_().getUrl(),
    params: { grace: P.grace, maxCarte: P.maxCarte, indicatif: P.indicatif, numeroDepot: P.numeroDepot },
    types: rows_(S.TYPES).map(function (t) { return { type: t.Type, places: Number(t['Places par défaut']), cout: Number(t['Coût mensuel par défaut']) }; }),
    sessions: sess,
    clients: clients.map(mapClient_),
    comptes: rows_(S.COMPTES).map(mapCompte_),
    places: rows_(S.PLACES).filter(function (p) { return p['État'] !== 'Supprimée'; }).map(mapPlace_),
    cartes: rows_(S.CARTES).map(mapCarte_),
    abos: rows_(S.ABOS).map(mapAbo_),
    aValider: aValider,
    horsTarif: rows_(S.VENTES).filter(function (v) { return v.Statut === 'Hors tarif' && joursEntre_(day_(v.Date), today_()) <= 45; }).reverse().map(function (v) {
      const cl = v.Client ? clients.find(function (c) { return c.ID === v.Client; }) : clientParNumero_(clients, v['Numéro']);
      return { id: v.ID, date: v.Date, numero: String(v['Numéro']), nom: v['Nom SMS'], montant: Number(v['Montant reçu'] || 0), clientId: cl ? cl.ID : '' };
    }),
    digest: dg,
    gains: gains_(6),
    historique: hist.map(function (h) { return { date: h.Date, objet: h.Objet, id: h.ID, action: h.Action, detail: h['Détail'] }; }),
    ventes: rows_(S.VENTES).filter(function (v) { return v.Statut !== 'À valider'; }).slice(-800).map(function (v) {
      return { id: v.ID, date: v.Date, client: v.Client, session: v.Session, prix: Number(v.Prix || 0), montant: Number(v['Montant reçu'] || 0),
        genre: v.Genre, statut: v.Statut, valideeLe: v['Validée le'], abo: v.Abonnement, auto: v.Auto === 'oui' };
    }),
    business: business_(P, sess, clients, abos),
    notif: { appId: P.osAppId, email: P.notifEmail }
  };
}

function mapClient_(c) { return { id: c.ID, nom: c.Nom, numeros: numsOf_(c), cree: c['Créé le'], notes: c.Notes, whatsapp: String(c.WhatsApp || ''), statut: c.Statut || 'Actif', supprimeLe: c['Supprimé le'], relanceLe: c['Relancé le'] }; }
function mapCompte_(c) { return { id: c.ID, type: c.Type, lien: String(c.Lien || ''), adresse: String(c.Adresse || ''), mail: String(c.Mail), mdp: String(c['Mot de passe']), nom: c.Nom, tel: String(c['Téléphone']), carte: c.Carte, cout: Number(c['Coût mensuel']), places: Number(c.Places), expiration: c.Expiration, statut: c.Statut }; }
function mapPlace_(p) { return { id: p.ID, compte: p.Compte, n: Number(p['N°']), profil: String(p.Profil), code: String(p.Code), etat: p['État'], client: p.Client, abo: p.Abonnement, depuis: p.Depuis }; }
function mapCarte_(c) { return { id: c.ID, repere: String(c['Repère']), titulaire: c.Titulaire, statut: c.Statut, remplacee: c['Remplacée par'], ajout: c['Date ajout'], notes: c.Notes, supprimeeLe: c['Supprimée le'] }; }
function mapAbo_(a) { return { id: a.ID, client: a.Client, type: a.Type, session: a.Session, debut: a['Début'], expiration: a.Expiration, statut: a.Statut, place: a.Place, compte: a.Compte, imp: a.Import ? infoImport_(a) : null }; }

/* ═════════════════════ VENTES & ABONNEMENTS ═════════════════════ */

function validerVente_(d) {
  const P = params_();
  const v = rows_(S.VENTES).find(function (x) { return x.ID === d.venteId; });
  if (!v) throw new Error('Vente introuvable.');
  if (v.Statut !== 'À valider') throw new Error('Cette vente a déjà été traitée.');
  const toutes = sessions_();
  // Une ou plusieurs sessions (ex. "FS-TEL-1M+MU-1M-B")
  const codes = (d.sessions && d.sessions.length ? d.sessions : String(d.session || '').split('+')).filter(String);
  const ss = codes.map(function (c) { const x = toutes.find(function (z) { return z.code === String(c); }); if (!x) throw new Error('Session inconnue : ' + c); return x; });
  if (!ss.length) throw new Error('Choisis une session.');

  const clients = rows_(S.CLIENTS);
  const num = normNum_(v['Numéro']);
  let client;
  if (d.clientId) {
    client = clients.find(function (c) { return c.ID === d.clientId; });
    if (!client) throw new Error('Client introuvable.');
    if (num && numsOf_(client).indexOf(num) < 0) {
      client['Numéros'] = numsOf_(client).concat(num).join(', ');
      update_(S.CLIENTS, client);
      hist_('Client', client.ID, 'Numéro ajouté', num);
    }
  } else {
    const nom = String(d.nom || v['Nom SMS'] || '').trim() || ('Client ' + num);
    client = { 'ID': nextId_('CL'), 'Nom': nom, 'Numéros': num, 'Créé le': now_(), 'Notes': '' };
    insert_(S.CLIENTS, client);
    hist_('Client', client.ID, 'Création', nom + ' ' + num);
  }
  if (d.whatsapp !== undefined && String(d.whatsapp).trim()) {
    client.WhatsApp = normNum_(d.whatsapp);
    update_(S.CLIENTS, client);
  }

  // Regroupement : même session 2 fois + « prolonger » → une seule prolongation de 2 × la durée
  let plan = [];
  ss.forEach(function (s) {
    const deja = plan.find(function (p) { return p.s.code === s.code; });
    if (deja && d.mode !== 'place') deja.duree += s.duree;
    else plan.push({ s: s, duree: s.duree, nouvelle: !!deja });
  });
  const date = day_(v.Date), res = [], prev = [];
  plan.forEach(function (p, i) {
    const r = appliquerAbonnement_(client, p.s, date, P, i === 0 ? d.placeId : '', {
      duree: p.duree, aboId: i === 0 ? d.aboId : (d.aboIds && d.aboIds[i]) || '', nouvellePlace: p.nouvelle || (i === 0 && d.mode === 'place' && !d.aboId && d.nouvellePlace)
    });
    res.push(r);
    prev.push({ abo: r.abo.ID, genre: r.genre, exp: r.prev ? r.prev.expiration : '', statut: r.prev ? r.prev.statut : '', session: r.prev ? r.prev.session : '' });
  });
  const tarif = ss.reduce(function (t, s) { return t + s.prix; }, 0);
  const prix = (d.montant !== undefined && d.montant !== '' && !isNaN(Number(d.montant))) ? Math.round(Number(d.montant)) : tarif;
  if (prix < tarif) hist_('Vente', v.ID, 'Remise', client.Nom + ' · ' + F_(prix) + ' reçus au lieu de ' + F_(tarif));
  v.Session = ss.map(function (s) { return s.code; }).join('+'); v.Prix = prix; v.Client = client.ID;
  v.Abonnement = res.map(function (r) { return r.abo.ID; }).join(',');
  v.Statut = 'Validée'; v.Genre = res[0].genre; v['Validée le'] = now_();
  v['Ancienne expiration'] = res[0].prev ? res[0].prev.expiration : '';
  v['Ancien statut'] = res[0].prev ? res[0].prev.statut : '';
  v['Ancienne session'] = JSON.stringify(prev);
  v.Auto = d.auto ? 'oui' : '';
  update_(S.VENTES, v);
  const out = res.map(function (r, i) { return resultat_(r, client, plan[i].s, P); });
  const o = out[0];
  o.autres = out.slice(1);
  if (plan.length > 1) o.session = plan.map(function (p) { return p.s.libelle; }).join(' + ');
  return o;
}

function appliquerAbonnement_(client, s, datePaiement, P, placeId, opt) {
  opt = opt || {};
  const duree = opt.duree || s.duree;
  const abos = rows_(S.ABOS).filter(function (a) { return a.Client === client.ID && a.Type === s.type; });
  let enCours = null;
  if (opt.aboId) {
    enCours = abos.find(function (a) { return a.ID === opt.aboId; });
    if (!enCours) throw new Error('Abonnement à prolonger introuvable.');
  } else if (!opt.nouvellePlace) {
    enCours = abos.find(function (a) { return EN_COURS.indexOf(a.Statut) >= 0; });
  }
  let abo, genre, prev = null;
  if (enCours) {
    // Renouvellement : on garde la place et on prolonge
    genre = 'Renouvellement';
    abo = enCours;
    prev = { expiration: abo.Expiration, statut: abo.Statut, session: abo.Session };
    const exp = day_(abo.Expiration);
    const dette = abo.Statut === 'Dette';
    const base = (exp && (dette || exp.getTime() >= addDays_(datePaiement, -P.grace).getTime())) ? exp : datePaiement;
    abo.Expiration = addMonths_(base, duree);
    abo.Session = s.code;
    abo['Mis à jour'] = now_();
    if (abo.Place) {
      abo.Statut = statutPour_(abo, P);
      if (dette && day_(abo.Expiration).getTime() < today_().getTime()) abo.Statut = 'Dette';
    } else placer_(abo, client, P, placeId);
    update_(S.ABOS, abo);
  } else {
    genre = abos.some(function (a) { return a.Statut !== 'Inventaire'; }) ? (opt.nouvellePlace ? 'Nouvelle place' : 'Retour') : 'Nouveau';
    abo = { 'ID': nextId_('AB'), 'Client': client.ID, 'Type': s.type, 'Session': s.code, 'Début': datePaiement,
      'Expiration': addMonths_(datePaiement, duree), 'Statut': 'En attente', 'Place': '', 'Compte': '', 'Mis à jour': now_() };
    placer_(abo, client, P, placeId, !!opt.nouvellePlace);
    insert_(S.ABOS, abo);
  }
  hist_('Abonnement', abo.ID, genre, client.Nom + ' · ' + s.libelle + ' · jusqu\'au ' + fmt_(abo.Expiration));
  return { abo: abo, genre: genre, prev: prev };
}

/** Attribue une place libre (priorité : ancienne place, puis ancien compte). */
function placer_(abo, client, P, placeId, deuxieme) {
  const comptes = rows_(S.COMPTES).filter(function (c) { return c.Type === abo.Type && c.Statut === 'Actif'; });
  const ids = comptes.map(function (c) { return c.ID; });
  const libres = rows_(S.PLACES).filter(function (p) { return ids.indexOf(p.Compte) >= 0 && p['État'] === 'Libre'; });
  let choix = null;
  if (placeId) {
    choix = libres.find(function (p) { return p.ID === placeId; });
    if (!choix) throw new Error('Cette place n\'est pas libre ou ne correspond pas au type ' + abo.Type + '.');
  }
  if (!choix && libres.length) {
    const anciens = deuxieme ? [] : rows_(S.ABOS).filter(function (a) { return a.Client === client.ID && a.Type === abo.Type && a.Place && a.ID !== abo.ID && EN_COURS.indexOf(a.Statut) < 0; })
      .sort(function (a, b) { return new Date(b['Mis à jour']) - new Date(a['Mis à jour']); });
    for (let i = 0; i < anciens.length && !choix; i++) choix = libres.find(function (p) { return p.ID === anciens[i].Place; });
    for (let i = 0; i < anciens.length && !choix; i++) choix = libres.find(function (p) { return p.Compte === anciens[i].Compte; });
    if (!choix) {
      libres.sort(function (a, b) { return (ids.indexOf(a.Compte) - ids.indexOf(b.Compte)) || (a['N°'] - b['N°']); });
      choix = libres[0];
    }
  }
  if (!choix) { abo.Statut = 'En attente'; return null; }
  choix['État'] = 'Occupée'; choix.Client = client.ID; choix.Abonnement = abo.ID; choix.Depuis = now_();
  update_(S.PLACES, choix);
  abo.Place = choix.ID; abo.Compte = choix.Compte; abo.Statut = statutPour_(abo, P);
  return choix;
}

function resultat_(r, client, s, P) {
  const abo = r.abo;
  const out = { genre: r.genre, client: client.Nom, session: s.libelle, expiration: abo.Expiration, statut: abo.Statut, acces: null, whatsapp: '', info: '' };
  const tel = waNum_(client);
  if (abo.Place) {
    const pl = rows_(S.PLACES).find(function (p) { return p.ID === abo.Place; });
    const cp = rows_(S.COMPTES).find(function (c) { return c.ID === abo.Compte; });
    out.acces = { mail: String(cp.Mail), mdp: String(cp['Mot de passe']), profil: String(pl.Profil), code: String(pl.Code), place: pl['N°'] };
    let msg;
    if (r.genre === 'Renouvellement') {
      msg = 'Bonjour ' + client.Nom + ' 👋\nMerci ! Votre abonnement ' + s.libelle + ' chez ' + MARQUE +
        ' est renouvelé jusqu\'au ' + fmt_(abo.Expiration) + '.\nVos accès restent les mêmes. Bon divertissement ✨';
      out.info = 'Place conservée : ' + cp.Mail + ', place ' + pl['N°'] + '.';
    } else {
      msg = messageAcces_(client.Nom, s.libelle, out.acces, abo.Expiration);
      out.info = 'Ajoute ' + client.Nom + ' sur ' + cp.Mail + ', place ' + pl['N°'] + '.';
    }
    out.whatsapp = wa_(tel, msg, P);
  } else {
    out.info = 'Aucune place libre en « ' + s.type + ' ». Ajoute un compte ou prépare une place : ' + client.Nom + ' y sera placé(e) automatiquement.';
  }
  return out;
}

function messageAcces_(nom, libelle, a, exp) {
  return 'Bonjour ' + nom + ' 👋\nMerci pour votre abonnement ' + libelle + ' chez ' + MARQUE + ' ✨\n\nVos accès :\n' +
    '📧 Mail : ' + a.mail + '\n🔑 Mot de passe : ' + a.mdp + '\n👤 Profil : ' + a.profil + '\n🔢 Code : ' + a.code +
    '\n\n📅 Valable jusqu\'au ' + fmt_(exp) + '.\nMerci de ne modifier aucune information du compte 🙏';
}

function ignorerVente_(d) {
  const v = rows_(S.VENTES).find(function (x) { return x.ID === d.venteId; });
  if (!v) throw new Error('Vente introuvable.');
  v.Statut = 'Ignorée';
  update_(S.VENTES, v);
  return { ok: true };
}

function venteManuelle_(d) {
  const s = sessions_().find(function (x) { return x.code === String(d.session); });
  if (!s) throw new Error('Choisis une session.');
  let num = normNum_(d.numero);
  const montant = d.montant !== undefined && d.montant !== '' ? Math.round(Number(d.montant)) : s.prix;
  if (isNaN(montant) || montant < 0) throw new Error('Montant invalide.');
  if (d.venteId) {
    const h = rows_(S.VENTES).find(function (x) { return x.ID === d.venteId; });
    if (!h || h.Statut !== 'Hors tarif') throw new Error('Ce paiement a déjà été utilisé.');
    h.Statut = 'À valider'; h['Sessions possibles'] = s.code; update_(S.VENTES, h);
    num = normNum_(h['Numéro']);
    let cid = d.clientId || (h.Client || '');
    if (!cid && num) { const c = clientParNumero_(rows_(S.CLIENTS), num); if (c) cid = c.ID; }
    return validerVente_({ venteId: h.ID, session: s.code, clientId: cid, nom: d.nom || h['Nom SMS'], placeId: d.placeId, montant: montant });
  }
  if (!num && !d.clientId) throw new Error('Indique un numéro ou choisis un client.');
  const v = {
    'ID': nextId_('V'), 'Date': d.date ? day_(new Date(d.date)) : now_(), 'Trans ID': 'MANUEL', 'Numéro': num,
    'Nom SMS': d.nom || '', 'Montant reçu': montant, 'Sessions possibles': s.code,
    'Session': '', 'Prix': '', 'Client': '', 'Abonnement': '', 'Statut': 'À valider', 'Genre': '', 'SMS': 'Saisie manuelle'
  };
  insert_(S.VENTES, v);
  let clientId = d.clientId;
  if (!clientId && num) { const c = clientParNumero_(rows_(S.CLIENTS), num); if (c) clientId = c.ID; }
  return validerVente_({ venteId: v.ID, session: s.code, clientId: clientId, nom: d.nom, placeId: d.placeId, montant: montant });
}

/** Corrige après coup la somme réellement reçue pour une vente validée. */
function modifierMontant_(d) {
  const v = rows_(S.VENTES).find(function (x) { return x.ID === d.venteId; });
  if (!v || v.Statut !== 'Validée') throw new Error('Vente introuvable.');
  const m = Math.round(Number(d.montant));
  if (isNaN(m) || m < 0) throw new Error('Montant invalide.');
  const avant = Number(v.Prix || 0);
  v.Prix = m;
  if (String(v['Trans ID']) === 'MANUEL') v['Montant reçu'] = m;
  update_(S.VENTES, v);
  hist_('Vente', v.ID, 'Montant corrigé', F_(avant) + ' → ' + F_(m));
  return { ok: true, avant: avant };
}

/** Le client est retiré de l'application : la place passe « À préparer ». */
function liberer_(d) {
  const abo = rows_(S.ABOS).find(function (a) { return a.ID === d.aboId; });
  if (!abo) throw new Error('Abonnement introuvable.');
  const client = rows_(S.CLIENTS).find(function (c) { return c.ID === abo.Client; }) || { Nom: '?' };
  let placeId = '';
  if (abo.Place) {
    const pl = rows_(S.PLACES).find(function (p) { return p.ID === abo.Place; });
    if (pl && pl.Abonnement === abo.ID) {
      pl['État'] = sansProfil_(pl.Compte) ? 'Libre' : 'À préparer'; pl.Client = ''; pl.Abonnement = ''; pl.Depuis = now_();
      update_(S.PLACES, pl);
      placeId = pl.ID;
      hist_('Place', pl.ID, 'Libérée', client.Nom + ' retiré(e) · ancien profil : ' + pl.Profil);
    }
  }
  abo.Statut = 'Terminé'; abo['Mis à jour'] = now_();
  update_(S.ABOS, abo);
  let rembourse = null;
  if (d.retirerArgent) rembourse = retirerArgent_(abo, client);
  return { placeId: placeId, rembourse: rembourse };
}

/** Retire de tes gains le dernier paiement de cet abonnement (paiement enregistré par erreur). */
function retirerArgent_(abo, client) {
  const sess = {}; sessions_().forEach(function (x) { sess[x.code] = x; });
  const v = rows_(S.VENTES).filter(function (x) { return x.Statut === 'Validée' && String(x.Abonnement || '').split(',').indexOf(abo.ID) >= 0; })
    .sort(function (a, b) { return new Date(b.Date) - new Date(a.Date); })[0];
  if (!v) return null;
  const abos = String(v.Abonnement).split(','), codes = String(v.Session || '').split('+');
  let montant = Number(v.Prix || 0);
  if (abos.length > 1) {   // paiement groupé : on ne retire que la part de cet abonnement
    const tarif = codes.reduce(function (t, c) { return t + (sess[c] ? sess[c].prix : 0); }, 0);
    const part = sess[codes[abos.indexOf(abo.ID)]] || sess[abo.Session];
    montant = tarif && part ? Math.round(Number(v.Prix || 0) * part.prix / tarif) : 0;
  }
  v.Prix = Number(v.Prix || 0) - montant;
  if (v.Prix <= 0) { v.Prix = 0; v.Statut = 'Annulée'; }
  update_(S.VENTES, v);
  hist_('Vente', v.ID, 'Argent retiré', client.Nom + ' · ' + F_(montant) + ' retirés des gains (retrait)');
  return { venteId: v.ID, montant: montant };
}

/** Saisie du nouveau profil + code. Place libre → attribuée aux clients en attente. */
function preparerPlace_(d) {
  const P = params_();
  const pl = rows_(S.PLACES).find(function (p) { return p.ID === d.placeId; });
  if (!pl) throw new Error('Place introuvable.');
  const profil = String(d.profil || '').trim(), code = String(d.code || '').trim();
  if (!profil || !code) throw new Error('Saisis le profil et le code.');
  hist_('Place', pl.ID, 'Profil/code modifiés', 'Ancien : ' + pl.Profil + ' / ' + pl.Code);
  pl.Profil = profil; pl.Code = code;
  if (pl['État'] !== 'Occupée') { pl['État'] = 'Libre'; pl.Depuis = now_(); }
  update_(S.PLACES, pl);
  if (pl['État'] !== 'Libre') return { placement: null };

  const compte = rows_(S.COMPTES).find(function (c) { return c.ID === pl.Compte; });
  const attente = rows_(S.ABOS).filter(function (a) { return a.Statut === 'En attente' && a.Type === compte.Type; })
    .sort(function (a, b) { return new Date(a['Mis à jour']) - new Date(b['Mis à jour']); })[0];
  if (!attente) return { placement: null };
  const client = rows_(S.CLIENTS).find(function (c) { return c.ID === attente.Client; });
  placer_(attente, client, P, pl.ID);
  update_(S.ABOS, attente);
  const s = sessions_().find(function (x) { return x.code === String(attente.Session); }) || { libelle: attente.Type, type: attente.Type };
  return { placement: resultat_({ abo: attente, genre: 'Nouveau' }, client, s, P) };
}

/* ═════════════════════ COMPTES ═════════════════════ */

function ajouterCompte_(d) {
  const P = params_();
  const t = rows_(S.TYPES).find(function (x) { return x.Type === d.type; });
  if (!t) throw new Error('Choisis un type de compte.');
  if (!String(d.mail || '').trim()) throw new Error('Le mail est obligatoire.');
  verifierCarte_(d.carte, null, P, d.type);
  const nb = Number(d.places || t['Places par défaut']);
  const c = {
    'ID': nextId_('CPT'), 'Type': d.type, 'Mail': String(d.mail).trim(), 'Mot de passe': d.mdp || '', 'Nom': d.nom || '',
    'Téléphone': d.tel || '', 'Carte': d.carte || '', 'Coût mensuel': Number(d.cout || t['Coût mensuel par défaut']),
    'Places': nb, 'Expiration': d.expiration ? day_(new Date(d.expiration)) : '', 'Statut': 'Actif', 'Créé le': now_(),
    'Lien': String(d.lien || '').trim(), 'Adresse': String(d.adresse || '').trim()
  };
  insert_(S.COMPTES, c);
  const profils = d.profils || [];
  for (let i = 1; i <= nb; i++) ajouterPlace_(c.ID, i, profils[i - 1]);
  hist_('Compte', c.ID, 'Création', c.Mail + (c.Carte ? ' · carte ' + c.Carte : ''));
  if (c.Carte) hist_('Carte', c.Carte, 'Compte relié', c.Mail);
  return { id: c.ID };
}

/** Comptes « Musique » : on rejoint avec un lien et une adresse, sans profil ni code. */
function sansProfil_(compteId) {
  const c = rows_(S.COMPTES).find(function (x) { return x.ID === compteId; });
  return !!c && /musique/i.test(String(c.Type));
}

function ajouterPlace_(compteId, n, pc) {
  const ok = (pc && String(pc.profil || '').trim() && String(pc.code || '').trim()) || sansProfil_(compteId);
  insert_(S.PLACES, {
    'ID': nextId_('PL'), 'Compte': compteId, 'N°': n, 'Profil': ok && pc ? String(pc.profil || '').trim() : '',
    'Code': ok && pc ? String(pc.code || '').trim() : '', 'État': ok ? 'Libre' : 'À préparer', 'Client': '', 'Abonnement': '', 'Depuis': now_()
  });
}

function modifierCompte_(d) {
  const P = params_();
  const c = rows_(S.COMPTES).find(function (x) { return x.ID === d.id; });
  if (!c) throw new Error('Compte introuvable.');
  if (d.carte !== undefined && String(d.carte) !== String(c.Carte)) {
    verifierCarte_(d.carte, c.ID, P, c.Type);
    hist_('Compte', c.ID, 'Carte changée', c.Mail + ' : ' + (c.Carte || 'aucune') + ' → ' + (d.carte || 'aucune'));
    if (c.Carte) hist_('Carte', c.Carte, 'Compte retiré', c.Mail);
    if (d.carte) hist_('Carte', d.carte, 'Compte relié', c.Mail);
    c.Carte = d.carte || '';
  }
  ['mail:Mail', 'mdp:Mot de passe', 'nom:Nom', 'tel:Téléphone', 'lien:Lien', 'adresse:Adresse'].forEach(function (m) {
    const k = m.split(':');
    if (d[k[0]] !== undefined) c[k[1]] = d[k[0]];
  });
  if (d.cout !== undefined && d.cout !== '') c['Coût mensuel'] = Number(d.cout);
  if (d.expiration) c.Expiration = day_(new Date(d.expiration));

  if (d.places !== undefined && Number(d.places) !== Number(c.Places)) {
    const voulu = Number(d.places);
    const pls = rows_(S.PLACES).filter(function (p) { return p.Compte === c.ID && p['État'] !== 'Supprimée'; });
    if (voulu > pls.length) {
      let max = pls.reduce(function (m, p) { return Math.max(m, Number(p['N°'])); }, 0);
      for (let i = pls.length; i < voulu; i++) ajouterPlace_(c.ID, ++max, null);
    } else if (voulu < pls.length) {
      const retirables = pls.filter(function (p) { return p['État'] !== 'Occupée'; }).sort(function (a, b) { return b['N°'] - a['N°']; });
      const aRetirer = pls.length - voulu;
      if (retirables.length < aRetirer) throw new Error('Impossible : trop de places sont occupées. Retire d\'abord des clients.');
      retirables.slice(0, aRetirer).forEach(function (p) { p['État'] = 'Supprimée'; update_(S.PLACES, p); });
    }
    c.Places = voulu;
  }
  update_(S.COMPTES, c);
  return { ok: true };
}

function fermerCompte_(d) {
  const c = rows_(S.COMPTES).find(function (x) { return x.ID === d.id; });
  if (!c) throw new Error('Compte introuvable.');
  const pls = rows_(S.PLACES).filter(function (p) { return p.Compte === c.ID && p['État'] !== 'Supprimée'; });
  if (pls.some(function (p) { return p['État'] === 'Occupée'; })) throw new Error('Des clients sont encore sur ce compte. Retire-les d\'abord.');
  pls.forEach(function (p) { p['État'] = 'Supprimée'; update_(S.PLACES, p); });
  c.Statut = 'Fermé';
  update_(S.COMPTES, c);
  hist_('Compte', c.ID, 'Fermeture', c.Mail);
  if (c.Carte) hist_('Carte', c.Carte, 'Compte retiré', c.Mail + ' (fermé)');
  return { ok: true };
}

function payerCompte_(d) {
  const c = rows_(S.COMPTES).find(function (x) { return x.ID === d.id; });
  if (!c) throw new Error('Compte introuvable.');
  const t = today_();
  const exp = day_(c.Expiration);
  const base = (exp && exp.getTime() >= addDays_(t, -10).getTime()) ? exp : t;
  const nouvelle = addMonths_(base, Number(d.mois || 1));
  const montant = Number(d.montant || c['Coût mensuel'] || 0);
  const pid = nextId_('PC');
  insert_(S.PAIEMENTS, { 'ID': pid, 'Date': now_(), 'Compte': c.ID, 'Carte': c.Carte, 'Montant': montant, 'Nouvelle expiration': nouvelle,
    'Ancienne expiration': c.Expiration || '', 'Statut': 'Payé' });
  c.Expiration = nouvelle;
  update_(S.COMPTES, c);
  hist_('Compte', c.ID, 'Paiement', F_(montant) + ' · jusqu\'au ' + fmt_(nouvelle));
  return { expiration: nouvelle, paiementId: pid };
}

/* ═════════════════════ CARTES ═════════════════════ */

/** Limite : MAX_COMPTES_PAR_CARTE comptes de chaque type par carte (ex. 3 Films + 3 Musique + 3 PRM). */
function verifierCarte_(carteId, compteExclu, P, type) {
  if (!carteId) return;
  const ca = rows_(S.CARTES).find(function (x) { return x.ID === carteId; });
  if (!ca) throw new Error('Carte introuvable.');
  if (ca.Statut !== 'Active') throw new Error('Cette carte n\'est plus active.');
  const n = rows_(S.COMPTES).filter(function (c) { return c.Statut === 'Actif' && c.Carte === carteId && c.ID !== compteExclu && (!type || c.Type === type); }).length;
  if (n >= P.maxCarte) throw new Error('La carte « ' + ca['Repère'] + ' » a déjà ' + n + ' comptes ' + (type || '') + ' (maximum ' + P.maxCarte + ' par type).');
}

function verifierRepere_(r) {
  r = String(r || '').trim();
  if (!r) throw new Error('Donne un repère à la carte (ex. « Visa Coris …4521 »).');
  if ((r.match(/\d/g) || []).length >= 9) throw new Error('Ne saisis pas le numéro complet de la carte : seulement les 4 derniers chiffres.');
  return r;
}

function ajouterCarte_(d) {
  const ca = { 'ID': nextId_('CA'), 'Repère': verifierRepere_(d.repere), 'Titulaire': d.titulaire || '', 'Statut': 'Active',
    'Remplacée par': '', 'Date ajout': now_(), 'Notes': d.notes || '' };
  insert_(S.CARTES, ca);
  hist_('Carte', ca.ID, 'Ajout', ca['Repère']);
  return { id: ca.ID };
}

function modifierCarte_(d) {
  const ca = rows_(S.CARTES).find(function (x) { return x.ID === d.id; });
  if (!ca) throw new Error('Carte introuvable.');
  if (d.repere !== undefined) ca['Repère'] = verifierRepere_(d.repere);
  if (d.titulaire !== undefined) ca.Titulaire = d.titulaire;
  if (d.notes !== undefined) ca.Notes = d.notes;
  if (d.statut && d.statut !== ca.Statut) {
    if (d.statut !== 'Active' && rows_(S.COMPTES).some(function (c) { return c.Statut === 'Actif' && c.Carte === ca.ID; })) {
      throw new Error('Des comptes sont encore reliés à cette carte. Utilise « Remplacer ».');
    }
    hist_('Carte', ca.ID, 'Statut', ca.Statut + ' → ' + d.statut);
    ca.Statut = d.statut;
  }
  update_(S.CARTES, ca);
  return { ok: true };
}

/** Remplace une carte par une autre (existante ou nouvelle) sur tous ses comptes. */
function remplacerCarte_(d) {
  const P = params_();
  const ancienne = rows_(S.CARTES).find(function (x) { return x.ID === d.id; });
  if (!ancienne) throw new Error('Carte introuvable.');
  let nouvelleId = d.nouvelleId;
  if (!nouvelleId) nouvelleId = ajouterCarte_({ repere: d.repere, titulaire: d.titulaire }).id;
  if (nouvelleId === ancienne.ID) throw new Error('Choisis une autre carte.');
  const nouvelle = rows_(S.CARTES).find(function (x) { return x.ID === nouvelleId; });
  if (!nouvelle || nouvelle.Statut !== 'Active') throw new Error('La nouvelle carte doit être active.');
  const comptes = rows_(S.COMPTES).filter(function (c) { return c.Statut === 'Actif' && c.Carte === ancienne.ID; });
  const types = {};
  comptes.forEach(function (c) { types[c.Type] = (types[c.Type] || 0) + 1; });
  Object.keys(types).forEach(function (ty) {
    const deja = rows_(S.COMPTES).filter(function (c) { return c.Statut === 'Actif' && c.Carte === nouvelleId && c.Type === ty; }).length;
    if (deja + types[ty] > P.maxCarte) throw new Error('La carte « ' + nouvelle['Repère'] + ' » aurait ' + (deja + types[ty]) + ' comptes ' + ty + ' (maximum ' + P.maxCarte + ' par type).');
  });
  comptes.forEach(function (c) {
    c.Carte = nouvelleId;
    update_(S.COMPTES, c);
    hist_('Compte', c.ID, 'Carte changée', c.Mail + ' : ' + ancienne['Repère'] + ' → ' + nouvelle['Repère']);
    hist_('Carte', nouvelleId, 'Compte relié', c.Mail);
  });
  ancienne.Statut = 'Remplacée'; ancienne['Remplacée par'] = nouvelleId;
  update_(S.CARTES, ancienne);
  hist_('Carte', ancienne.ID, 'Remplacée', 'par ' + nouvelle['Repère'] + ' (' + comptes.length + ' compte(s) transféré(s))');
  return { ok: true };
}

/* ═════════════════════ CLIENTS ═════════════════════ */

function modifierClient_(d) {
  const c = rows_(S.CLIENTS).find(function (x) { return x.ID === d.id; });
  if (!c) throw new Error('Client introuvable.');
  if (d.nom !== undefined) c.Nom = String(d.nom).trim() || c.Nom;
  if (d.numeros !== undefined) {
    const nums = String(d.numeros).split(/[,;\n]+/).map(normNum_).filter(String);
    const autres = rows_(S.CLIENTS).filter(function (x) { return x.ID !== c.ID; });
    nums.forEach(function (n) {
      const o = autres.find(function (x) { return numsOf_(x).indexOf(n) >= 0; });
      if (o) throw new Error('Le numéro ' + n + ' appartient déjà à ' + o.Nom + '.');
    });
    c['Numéros'] = nums.join(', ');
  }
  if (d.notes !== undefined) c.Notes = d.notes;
  if (d.whatsapp !== undefined) c.WhatsApp = normNum_(d.whatsapp);
  update_(S.CLIENTS, c);
  return { ok: true };
}

/** Enregistre en une fois les numéros WhatsApp retrouvés dans les contacts. */
function whatsappLot_(d) {
  const liste = (d && d.liste) || [];
  const ix = {}; rows_(S.CLIENTS).forEach(function (c) { ix[c.ID] = c; });
  let n = 0;
  liste.forEach(function (x) {
    const c = ix[x.id], w = normNum_(x.whatsapp);
    if (c && d.vider) { if (c.WhatsApp) { c.WhatsApp = ''; update_(S.CLIENTS, c); n++; } return; }
    if (!c || !w || normNum_(c.WhatsApp) === w) return;
    c.WhatsApp = w; update_(S.CLIENTS, c); n++;
  });
  if (n) hist_('Client', '', 'WhatsApp importés', n + ' numéro(s) depuis les contacts');
  return { ok: true, n: n };
}

/* ═════════════════════ STATUTS, ALERTES, GAINS ═════════════════════ */

function statutPour_(a, P) {
  const t = today_().getTime(), exp = day_(a.Expiration);
  if (!exp) return 'Actif';
  if (t <= exp.getTime()) return 'Actif';
  if (t <= addDays_(exp, P.grace).getTime()) return 'Grâce';
  return 'À retirer';
}

function majStatuts_(P) {
  const t = today_().getTime();
  rows_(S.ABOS).forEach(function (a) {
    if (EN_COURS.indexOf(a.Statut) < 0 || a.Statut === 'Dette') return;
    let ns;
    if (!a.Place) {
      const exp = day_(a.Expiration);
      if (exp && t > addDays_(exp, P.grace).getTime()) ns = 'Terminé'; else return;
    } else ns = statutPour_(a, P);
    if (ns !== a.Statut) { a.Statut = ns; update_(S.ABOS, a); }
  });
}

function digest_(P) {
  majStatuts_(P);
  const t = today_();
  const clients = index_(rows_(S.CLIENTS)), comptesL = rows_(S.COMPTES), comptes = index_(comptesL);
  const placesL = rows_(S.PLACES).filter(function (p) { return p['État'] !== 'Supprimée'; }), places = index_(placesL);
  const cartesL = rows_(S.CARTES), cartes = index_(cartesL), abosL = rows_(S.ABOS);
  const d = { expirent: [], grace: [], aRetirer: [], enAttente: [], dettes: [], comptes: [], aPreparer: [], ventes: 0, anomalies: [], libres: {}, inventaire: 0 };

  d.inventaire = abosL.filter(function (a) { return a.Statut === 'Inventaire'; }).length;
  abosL.forEach(function (a) {
    if (EN_COURS.indexOf(a.Statut) < 0) return;
    const c = clients[a.Client] || { Nom: '?', 'Numéros': '' };
    const exp = day_(a.Expiration), j = joursEntre_(t, exp);
    const pl = places[a.Place], cp = comptes[a.Compte];
    const b = { id: a.ID, client: c.Nom, clientId: a.Client, type: a.Type, expiration: exp, jours: j,
      compte: cp ? String(cp.Mail) : '', place: pl ? pl['N°'] : '', profil: pl ? String(pl.Profil) : '' };
    if (a.Statut === 'Actif' && j >= 0 && j <= P.rappelClient) {
      b.wa = wa_(waNum_(c), 'Bonjour ' + c.Nom + ' 👋\nVotre abonnement ' + a.Type + ' chez ' + MARQUE + ' expire le ' + fmt_(exp) +
        '.\nPour continuer à en profiter, pensez à renouveler' + (P.numeroDepot ? ' par Orange Money au ' + P.numeroDepot : '') + ' 🙏' + lignePaiement_(a.Session, P), P);
      d.expirent.push(b);
    } else if (a.Statut === 'Grâce') {
      b.limite = addDays_(exp, P.grace);
      b.wa = wa_(waNum_(c), 'Bonjour ' + c.Nom + ' 👋\nVotre abonnement ' + a.Type + ' chez ' + MARQUE + ' a expiré le ' + fmt_(exp) +
        '.\nVous pouvez encore renouveler jusqu\'au ' + fmt_(b.limite) + ' pour garder votre place' +
        (P.numeroDepot ? ' (Orange Money : ' + P.numeroDepot + ')' : '') + ' 🙏' + lignePaiement_(a.Session, P), P);
      d.grace.push(b);
    } else if (a.Statut === 'À retirer') {
      b.wa = wa_(waNum_(c), 'Bonjour ' + c.Nom + ' 👋\nVotre abonnement ' + a.Type + ' chez ' + MARQUE + ' a expiré le ' + fmt_(exp) +
        '.\nPour le garder, merci de renouveler dès que possible' + (P.numeroDepot ? ' par Orange Money au ' + P.numeroDepot : '') + ', sinon votre place sera libérée 🙏' + lignePaiement_(a.Session, P), P);
      d.aRetirer.push(b);
    }
    else if (a.Statut === 'En attente') d.enAttente.push(b);
    else if (a.Statut === 'Dette') {
      b.mois = Math.max(1, Math.ceil(-j / 30));
      b.wa = wa_(waNum_(c), 'Bonjour ' + c.Nom + ' 👋\nPetit rappel : votre abonnement ' + a.Type + ' chez ' + MARQUE + ' n\'est plus réglé depuis le ' + fmt_(exp) +
        ' (' + b.mois + ' mois).\nMerci de régulariser' + (P.numeroDepot ? ' par Orange Money au ' + P.numeroDepot : '') + ' pour garder votre place 🙏' + lignePaiement_(a.Session, P), P);
      d.dettes.push(b);
    }
  });
  d.expirent.sort(function (a, b) { return a.jours - b.jours; });

  comptesL.filter(function (c) { return c.Statut === 'Actif'; }).forEach(function (c) {
    const exp = day_(c.Expiration);
    if (!exp) return;
    const j = joursEntre_(t, exp);
    if (j <= P.rappelCompte) {
      const ca = cartes[c.Carte];
      d.comptes.push({ id: c.ID, mail: String(c.Mail), type: c.Type, expiration: exp, jours: j, cout: Number(c['Coût mensuel']),
        carte: ca ? String(ca['Repère']) : '', clients: placesL.filter(function (p) { return p.Compte === c.ID && p['État'] === 'Occupée'; }).length });
    }
  });
  d.comptes.sort(function (a, b) { return a.jours - b.jours; });

  placesL.filter(function (p) { return p['État'] === 'À préparer'; }).forEach(function (p) {
    const cp = comptes[p.Compte];
    if (!cp || cp.Statut !== 'Actif') return;
    d.aPreparer.push({ id: p.ID, compte: String(cp.Mail), type: cp.Type, n: p['N°'], jours: joursEntre_(day_(p.Depuis), t) });
  });

  rows_(S.TYPES).forEach(function (ty) {
    const ids = comptesL.filter(function (c) { return c.Statut === 'Actif' && c.Type === ty.Type; }).map(function (c) { return c.ID; });
    d.libres[ty.Type] = { libres: placesL.filter(function (p) { return ids.indexOf(p.Compte) >= 0 && p['État'] === 'Libre'; }).length,
      total: placesL.filter(function (p) { return ids.indexOf(p.Compte) >= 0; }).length };
  });
  d.ventes = rows_(S.VENTES).filter(function (v) { return v.Statut === 'À valider'; }).length;
  d.anomalies = anomalies_(P, t, clients, comptesL, cartesL, cartes, placesL, abosL, d.libres);
  return d;
}

function anomalies_(P, t, clients, comptesL, cartesL, cartes, placesL, abosL, libres) {
  const out = [];
  const actifs = comptesL.filter(function (c) { return c.Statut === 'Actif'; });
  const abos = index_(abosL);
  cartesL.forEach(function (ca) {
    const parType = {};
    actifs.filter(function (c) { return c.Carte === ca.ID; }).forEach(function (c) { parType[c.Type] = (parType[c.Type] || 0) + 1; });
    Object.keys(parType).forEach(function (ty) {
      if (parType[ty] > P.maxCarte) out.push('La carte « ' + ca['Repère'] + ' » a ' + parType[ty] + ' comptes ' + ty + ' (maximum ' + P.maxCarte + ' par type).');
    });
  });
  actifs.forEach(function (c) {
    if (!c.Carte) out.push('Le compte ' + c.Mail + ' n\'a pas de carte. Ajoutes-en une.');
    else {
      const ca = cartes[c.Carte];
      if (!ca) out.push('Le compte ' + c.Mail + ' est relié à une carte inconnue.');
      else if (ca.Statut !== 'Active') out.push('Le compte ' + c.Mail + ' est encore relié à la carte « ' + ca['Repère'] + ' » (' + String(ca.Statut).toLowerCase() + '). Remplace-la.');
    }
    const exp = day_(c.Expiration);
    if (!exp) out.push('Le compte ' + c.Mail + ' n\'a pas de date d\'expiration.');
    else if (exp.getTime() < t.getTime()) out.push('Le compte ' + c.Mail + ' a expiré le ' + fmt_(exp) + ' sans paiement noté.');
    const occ = placesL.filter(function (p) { return p.Compte === c.ID && p['État'] === 'Occupée'; }).length;
    if (occ > Number(c.Places)) out.push('Le compte ' + c.Mail + ' a ' + occ + ' clients pour ' + c.Places + ' places.');
  });
  placesL.filter(function (p) { return p['État'] === 'Occupée'; }).forEach(function (p) {
    const a = abos[p.Abonnement];
    if (!a || EN_COURS.indexOf(a.Statut) < 0) {
      const cp = comptesL.find(function (c) { return c.ID === p.Compte; });
      out.push('Place ' + p['N°'] + ' de ' + (cp ? cp.Mail : '?') + ' : occupée mais sans abonnement en cours.');
    }
  });
  const vus = {};
  abosL.filter(function (a) { return a.Place && EN_COURS.indexOf(a.Statut) >= 0; }).forEach(function (a) {
    const k = a.Client + '|' + a.Type;
    vus[k] = (vus[k] || 0) + 1;
    // plusieurs places du même type pour un client : autorisé (2e place)
  });
  Object.keys(libres).forEach(function (ty) {
    if (libres[ty].total > 0 && libres[ty].libres === 0) out.push('Plus aucune place libre en « ' + ty + ' ».');
  });
  const w = surveillanceSms_();
  if (w) out.unshift(w);
  return out;
}

function gains_(nb) {
  const sess = {};
  sessions_().forEach(function (s) { sess[s.code] = s; });
  const ventes = rows_(S.VENTES).filter(function (v) { return v.Statut === 'Validée'; });
  const pays = rows_(S.PAIEMENTS).filter(function (p) { return p.Statut !== 'Annulé'; });
  const departs = rows_(S.ABOS).filter(function (a) { return a.Statut === 'Terminé'; });
  const t = today_();
  const out = [];
  for (let i = 0; i < nb; i++) {
    const y = new Date(Date.UTC(t.getUTCFullYear(), t.getUTCMonth() - i, 1));
    const memeMois = function (d) { d = new Date(d); return d.getUTCFullYear() === y.getUTCFullYear() && d.getUTCMonth() === y.getUTCMonth(); };
    const vs = ventes.filter(function (v) { return memeMois(v.Date); });
    const ps = pays.filter(function (p) { return memeMois(p.Date); });
    const parType = {};
    vs.forEach(function (v) {
      const codes = String(v.Session || '').split('+');
      const tarif = codes.reduce(function (t, c) { return t + (sess[c] ? sess[c].prix : 0); }, 0);
      codes.forEach(function (code) {
        const se = sess[code], ty = se ? se.type : 'Autre';
        parType[ty] = parType[ty] || { ventes: 0, encaisse: 0 };
        parType[ty].ventes++;
        parType[ty].encaisse += se && tarif ? Math.round(Number(v.Prix || 0) * se.prix / tarif) : Number(v.Prix || 0);
      });
    });
    const enc = vs.reduce(function (s, v) { return s + Number(v.Prix || 0); }, 0);
    const nRen = vs.filter(function (v) { return v.Genre === 'Renouvellement'; }).length;
    const nDep = departs.filter(function (a) { return memeMois(a['Mis à jour']); }).length;
    const dep = ps.reduce(function (s, p) { return s + Number(p.Montant || 0); }, 0);
    out.push({
      mois: MOIS[y.getUTCMonth()] + ' ' + y.getUTCFullYear(), ventes: vs.length,
      nouveaux: vs.filter(function (v) { return v.Genre === 'Nouveau'; }).length,
      renouvellements: vs.filter(function (v) { return v.Genre === 'Renouvellement'; }).length,
      retours: vs.filter(function (v) { return v.Genre === 'Retour'; }).length,
      encaisse: enc, recu: vs.reduce(function (s, v) { return s + Number(v['Montant reçu'] || 0); }, 0),
      depenses: dep, benefice: enc - dep, parType: parType,
      departs: nDep, base: nRen + nDep, taux: (nRen + nDep) ? Math.round(100 * nRen / (nRen + nDep)) : null
    });
  }
  return out;
}

/* ═════════════════════ RAPPEL QUOTIDIEN (7 h) ═════════════════════ */

function rappelQuotidien() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const P = params_();
    const d = digest_(P);
    const blocs = [];
    const lien = function (url, txt) { return ' <a style="color:#25d366" href="' + url + '">' + txt + '</a>'; };

    if (d.aRetirer.length) blocs.push({ titre: '🔴 À retirer (délai de grâce terminé)', lignes: d.aRetirer.map(function (b) {
      return '<b>' + esc_(b.client) + '</b> · ' + esc_(b.compte) + ', place ' + b.place + ' (profil ' + esc_(b.profil) + ')'; }) });
    if (d.comptes.length) blocs.push({ titre: '💳 Comptes à payer', lignes: d.comptes.map(function (c) {
      return '<b>' + esc_(c.mail) + '</b> (' + esc_(c.type) + ') · ' + (c.jours < 0 ? '⚠️ expiré depuis le ' : 'expire le ') + fmt_(c.expiration) +
        ' · ' + F_(c.cout) + (c.carte ? ' · carte ' + esc_(c.carte) : '') + ' · ' + c.clients + ' client(s)'; }) });
    if (d.expirent.length) blocs.push({ titre: '⏳ Clients qui expirent bientôt', lignes: d.expirent.map(function (b) {
      return '<b>' + esc_(b.client) + '</b> · ' + esc_(b.type) + ' · ' + (b.jours === 0 ? 'aujourd\'hui' : 'dans ' + b.jours + ' j (' + fmt_(b.expiration) + ')') + lien(b.wa, 'Rappel WhatsApp'); }) });
    if (d.grace.length) blocs.push({ titre: '🟠 En délai de grâce', lignes: d.grace.map(function (b) {
      return '<b>' + esc_(b.client) + '</b> · ' + esc_(b.type) + ' · à retirer après le ' + fmt_(b.limite) + lien(b.wa, 'Rappel WhatsApp'); }) });
    const prep = d.aPreparer.filter(function (p) { return p.jours >= P.preparer; });
    if (prep.length) blocs.push({ titre: '🟠 Places à préparer (nouveau profil + code)', lignes: prep.map(function (p) {
      return esc_(p.compte) + ', place ' + p.n + ' · depuis ' + p.jours + ' j'; }) });
    if (d.dettes.length) blocs.push({ titre: '💸 Clients qui te doivent', lignes: d.dettes.map(function (b) {
      return '<b>' + esc_(b.client) + '</b> · ' + esc_(b.type) + ' · ' + b.mois + ' mois depuis le ' + fmt_(b.expiration) + lien(b.wa, 'Rappel WhatsApp'); }) });
    if (d.inventaire) blocs.push({ titre: '📋 Inventaire', lignes: [d.inventaire + ' client(s) importé(s) restent à vérifier dans l\'app.'] });
    if (d.enAttente.length) blocs.push({ titre: '⌛ Clients en attente de place', lignes: d.enAttente.map(function (b) {
      return '<b>' + esc_(b.client) + '</b> · ' + esc_(b.type); }) });
    if (d.ventes) blocs.push({ titre: '💰 Paiements à valider', lignes: [d.ventes + ' paiement(s) en attente de validation.'] });
    if (d.anomalies.length) blocs.push({ titre: '⚠️ Anomalies détectées', lignes: d.anomalies.map(esc_) });

    const t = today_();
    const dernierJour = addDays_(t, 1).getUTCDate() === 1;
    if (dernierJour) {
      const g = gains_(1)[0];
      const lignes = [
        '<b>Ventes :</b> ' + g.ventes + ' (' + g.nouveaux + ' nouveaux, ' + g.renouvellements + ' renouvellements, ' + g.retours + ' retours)',
        '<b>Encaissé :</b> ' + F_(g.encaisse), '<b>Dépenses comptes :</b> ' + F_(g.depenses),
        '<b style="color:#f3dfa2">Bénéfice : ' + F_(g.benefice) + '</b>'
      ];
      Object.keys(g.parType).forEach(function (ty) { lignes.push(esc_(ty) + ' : ' + g.parType[ty].ventes + ' vente(s) · ' + F_(g.parType[ty].encaisse)); });
      blocs.unshift({ titre: '📊 Bilan de ' + g.mois, lignes: lignes });
    }
    const veille = surveillanceSms_();
    if (veille) push_('⚠️ Android déconnecté ?', veille.replace('⚠️ ', ''), 'accueil', P);
    if (!blocs.length) return;
    const nb = d.aRetirer.length + d.comptes.length + d.expirent.length + d.aPreparer.length + d.ventes;
    push_('🔔 ' + (dernierJour ? 'Bilan du mois et rappels' : 'Tes rappels du jour'),
      nb + ' élément(s) à traiter aujourd\'hui' + (d.aRetirer.length ? ' · ' + d.aRetirer.length + ' client(s) à retirer' : ''), 'accueil', P);
    if (!P.notifEmail && !dernierJour) return;
    MailApp.sendEmail({ to: P.email, name: MARQUE,
      subject: (dernierJour ? '📊 Bilan du mois & ' : '') + '🔔 Rappels du ' + fmt_(t) + ' – ' + MARQUE,
      htmlBody: emailHtml_('Tes rappels du jour', blocs, true, 'accueil') });
  } finally {
    lock.releaseLock();
  }
}

/* ═════════════════════ VERSION 2 ═════════════════════ */

/** À exécuter une fois après la mise à jour v2 (ajoute les colonnes, la sauvegarde mensuelle…). */
function activerV2() {
  installer();
  DriveApp.getFileById(prop_('SS_ID')).getName(); // autorisation Drive (sauvegardes)
  Logger.log('✅ Version 2 activée.');
}

/** Date exacte lue dans le Trans ID Orange Money (ex. PP261118.2133… = 18/11/2026 21:33). */
function dateTrans_(id) {
  const m = String(id || '').match(/^[A-Z]{1,4}(\d{2})(\d{2})(\d{2})\.(\d{2})(\d{2})/);
  if (!m) return null;
  const y = 2000 + Number(m[1]), mo = Number(m[2]), j = Number(m[3]), h = Number(m[4]), mi = Number(m[5]);
  if (mo < 1 || mo > 12 || j < 1 || j > 31 || h > 23 || mi > 59) return null;
  const d = new Date(Date.UTC(y, mo - 1, j, h, mi));
  const ecart = (d.getTime() - now_().getTime()) / 86400000;
  return (ecart > 1 || ecart < -60) ? null : d;
}

/** Lien « payer en un geste » : ouvre le clavier avec le code Orange Money déjà rempli. */
function lienPaiement_(montant, P) {
  P = P || params_();
  const num = normNum_(P.numeroDepot);
  if (!num || !montant) return '';
  const code = P.ussd.replace('{numero}', num).replace('{montant}', String(Math.round(montant)));
  const base = prop_('PWA_URL');
  return base ? base + 'payer.html?c=' + encodeURIComponent(code) + '&m=' + Math.round(montant) : code;
}
function prixSession_(code) { const s = sessions_().find(function (x) { return x.code === String(code); }); return s ? s.prix : 0; }
function lignePaiement_(code, P) {
  const l = lienPaiement_(prixSession_(code), P);
  return l ? '\n👉 Payer ' + F_(prixSession_(code)) + ' en un geste : ' + l : '';
}

function waNum_(c) { return normNum_(c.WhatsApp) || numsOf_(c)[0] || ''; }

function enregistrerApp_(d) {
  if (d.url && /^https:\/\//.test(d.url)) props_().setProperty('PWA_URL', String(d.url).split('#')[0]);
  return { ok: true };
}

/** Notification sur le téléphone via OneSignal (si configuré). */
function push_(titre, texte, onglet, P) {
  try {
    P = P || params_();
    if (!P.osAppId || !P.osKey) return;
    const url = prop_('PWA_URL');
    const corps = { app_id: P.osAppId, target_channel: 'push', included_segments: ['Total Subscriptions'],
      headings: { en: titre, fr: titre }, contents: { en: texte, fr: texte } };
    if (url) corps.url = url + '#' + (onglet || 'accueil');
    const r = UrlFetchApp.fetch('https://api.onesignal.com/notifications?c=push', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Key ' + P.osKey }, payload: JSON.stringify(corps)
    });
    if (r.getResponseCode() >= 300) journal_('ERREUR', 'Notification : ' + r.getContentText().slice(0, 200));
  } catch (err) { journal_('ERREUR', 'Notification : ' + err.message); }
}

/* ---------- Annulations ---------- */

function annulerVente_(d) {
  const P = params_();
  const v = rows_(S.VENTES).find(function (x) { return x.ID === d.venteId; });
  if (!v || v.Statut !== 'Validée') throw new Error('Cette vente ne peut pas être annulée.');
  const quand = new Date(v['Validée le'] || v.Date).getTime();
  if (now_().getTime() - quand > 86400000) throw new Error('Une vente ne peut être annulée que dans les 24 h qui suivent sa validation.');
  let multi = null;
  try { multi = JSON.parse(v['Ancienne session']); } catch (e) { multi = null; }
  if (Array.isArray(multi)) {
    multi.forEach(function (m) { annulerAbo_(m.abo, m.genre, m.exp, m.statut, m.session, P); });
    v.Statut = 'Annulée';
    update_(S.VENTES, v);
    hist_('Vente', v.ID, 'Annulation', F_(v.Prix) + ' · ' + v.Session);
    return { ok: true };
  }
  const abo = rows_(S.ABOS).find(function (a) { return a.ID === v.Abonnement; });
  if (abo) {
    if (v.Genre === 'Renouvellement' && v['Ancienne expiration']) {
      abo.Expiration = v['Ancienne expiration'];
      abo.Session = v['Ancienne session'] || abo.Session;
      abo.Statut = abo.Place ? statutPour_(abo, P) : (v['Ancien statut'] || 'En attente');
    } else {
      if (abo.Place) {
        const pl = rows_(S.PLACES).find(function (p) { return p.ID === abo.Place; });
        if (pl && pl.Abonnement === abo.ID) {
          pl['État'] = 'Libre'; pl.Client = ''; pl.Abonnement = ''; pl.Depuis = now_();
          update_(S.PLACES, pl);
        }
      }
      abo.Statut = 'Annulé';
    }
    abo['Mis à jour'] = now_();
    update_(S.ABOS, abo);
  }
  v.Statut = 'Annulée';
  update_(S.VENTES, v);
  hist_('Vente', v.ID, 'Annulation', F_(v.Prix) + ' · ' + v.Session);
  return { ok: true };
}

function annulerAbo_(aboId, genre, exp, statut, session, P) {
  const abo = rows_(S.ABOS).find(function (a) { return a.ID === aboId; });
  if (!abo) return;
  if (genre === 'Renouvellement' && exp) {
    abo.Expiration = exp; abo.Session = session || abo.Session;
    abo.Statut = statut === 'Dette' ? 'Dette' : (abo.Place ? statutPour_(abo, P) : (statut || 'En attente'));
  } else {
    if (abo.Place) {
      const pl = rows_(S.PLACES).find(function (p) { return p.ID === abo.Place; });
      if (pl && pl.Abonnement === abo.ID) { pl['État'] = 'Libre'; pl.Client = ''; pl.Abonnement = ''; pl.Depuis = now_(); update_(S.PLACES, pl); }
    }
    abo.Statut = 'Annulé';
  }
  abo['Mis à jour'] = now_();
  update_(S.ABOS, abo);
}

function reprendreVente_(d) {
  const v = rows_(S.VENTES).find(function (x) { return x.ID === d.venteId; });
  if (!v || v.Statut !== 'Ignorée') throw new Error('Rien à reprendre.');
  v.Statut = 'À valider';
  update_(S.VENTES, v);
  return { ok: true };
}

function annulerLiberation_(d) {
  const P = params_();
  const abo = rows_(S.ABOS).find(function (a) { return a.ID === d.aboId; });
  if (!abo || abo.Statut !== 'Terminé' || !abo.Place) throw new Error('Impossible d\'annuler.');
  const pl = rows_(S.PLACES).find(function (p) { return p.ID === abo.Place; });
  if (!pl || pl['État'] === 'Occupée') throw new Error('La place a déjà été réattribuée.');
  pl['État'] = 'Occupée'; pl.Client = abo.Client; pl.Abonnement = abo.ID; pl.Depuis = now_();
  update_(S.PLACES, pl);
  abo.Statut = statutPour_(abo, P); abo['Mis à jour'] = now_();
  update_(S.ABOS, abo);
  if (d.venteId && d.montant) {
    const v = rows_(S.VENTES).find(function (x) { return x.ID === d.venteId; });
    if (v) { v.Prix = Number(v.Prix || 0) + Number(d.montant); v.Statut = 'Validée'; update_(S.VENTES, v); }
  }
  return { ok: true };
}

function annulerPaiementCompte_(d) {
  const p = rows_(S.PAIEMENTS).find(function (x) { return x.ID === d.paiementId; });
  if (!p || p.Statut === 'Annulé') throw new Error('Paiement introuvable.');
  const c = rows_(S.COMPTES).find(function (x) { return x.ID === p.Compte; });
  if (c) { c.Expiration = p['Ancienne expiration'] || ''; update_(S.COMPTES, c); }
  p.Statut = 'Annulé';
  update_(S.PAIEMENTS, p);
  return { ok: true };
}

/** Client en attente de place : on annule (et on peut retirer la vente des gains). */
function annulerAttente_(d) {
  const abo = rows_(S.ABOS).find(function (a) { return a.ID === d.aboId; });
  if (!abo || abo.Statut !== 'En attente') throw new Error('Ce client n\'est pas en attente.');
  abo.Statut = 'Annulé'; abo['Mis à jour'] = now_();
  update_(S.ABOS, abo);
  if (d.rembourser) {
    const v = rows_(S.VENTES).filter(function (x) { return x.Abonnement === abo.ID && x.Statut === 'Validée'; }).pop();
    if (v) { v.Statut = 'Remboursée'; update_(S.VENTES, v); }
  }
  return { ok: true };
}

/* ---------- Corbeille ---------- */

function supprimerClient_(d) {
  const c = rows_(S.CLIENTS).find(function (x) { return x.ID === d.id; });
  if (!c) throw new Error('Client introuvable.');
  const places = [];
  rows_(S.ABOS).filter(function (a) { return a.Client === c.ID && EN_COURS.indexOf(a.Statut) >= 0; }).forEach(function (a) {
    if (a.Place) { const r = liberer_({ aboId: a.ID }); if (r.placeId) places.push(r.placeId); }
    else { a.Statut = 'Annulé'; a['Mis à jour'] = now_(); update_(S.ABOS, a); }
  });
  c.Statut = 'Supprimé'; c['Supprimé le'] = now_();
  update_(S.CLIENTS, c);
  hist_('Client', c.ID, 'Corbeille', c.Nom);
  return { places: places };
}

function restaurerClient_(d) {
  const c = rows_(S.CLIENTS).find(function (x) { return x.ID === d.id; });
  if (!c) throw new Error('Client introuvable.');
  c.Statut = ''; c['Supprimé le'] = '';
  update_(S.CLIENTS, c);
  hist_('Client', c.ID, 'Restauré', c.Nom);
  return { ok: true };
}

function supprimerCarte_(d) {
  const ca = rows_(S.CARTES).find(function (x) { return x.ID === d.id; });
  if (!ca) throw new Error('Carte introuvable.');
  if (rows_(S.COMPTES).some(function (c) { return c.Statut === 'Actif' && c.Carte === ca.ID; })) {
    throw new Error('Des comptes utilisent encore cette carte. Remplace-la d\'abord.');
  }
  ca['Statut avant'] = ca.Statut; ca.Statut = 'Supprimée'; ca['Supprimée le'] = now_();
  update_(S.CARTES, ca);
  hist_('Carte', ca.ID, 'Corbeille', ca['Repère']);
  return { ok: true };
}

function restaurerCarte_(d) {
  const ca = rows_(S.CARTES).find(function (x) { return x.ID === d.id; });
  if (!ca || ca.Statut !== 'Supprimée') throw new Error('Carte introuvable.');
  ca.Statut = ca['Statut avant'] || 'Inactive'; ca['Supprimée le'] = '';
  update_(S.CARTES, ca);
  hist_('Carte', ca.ID, 'Restaurée', ca['Repère']);
  return { ok: true };
}

function marquerRelance_(d) {
  const c = rows_(S.CLIENTS).find(function (x) { return x.ID === d.id; });
  if (!c) throw new Error('Client introuvable.');
  c['Relancé le'] = now_();
  update_(S.CLIENTS, c);
  return { ok: true };
}

/* ---------- Compte bloqué ---------- */

function bloquerCompte_(d) {
  const P = params_();
  const compte = rows_(S.COMPTES).find(function (c) { return c.ID === d.id; });
  if (!compte || compte.Statut !== 'Actif') throw new Error('Compte introuvable ou déjà fermé.');
  const autres = rows_(S.COMPTES).filter(function (c) { return c.Statut === 'Actif' && c.Type === compte.Type && c.ID !== compte.ID; })
    .map(function (c) { return c.ID; });
  const places = rows_(S.PLACES);
  const libres = places.filter(function (p) { return autres.indexOf(p.Compte) >= 0 && p['État'] === 'Libre'; });
  const occupees = places.filter(function (p) { return p.Compte === compte.ID && p['État'] === 'Occupée'; });
  const clients = index_(rows_(S.CLIENTS)), abos = index_(rows_(S.ABOS)), comptes = index_(rows_(S.COMPTES));
  const plan = occupees.map(function (p, i) {
    const dest = libres[i] || null;
    return { placeId: p.ID, aboId: p.Abonnement, client: (clients[p.Client] || {}).Nom || '?',
      vers: dest ? { id: dest.ID, compte: String(comptes[dest.Compte].Mail), n: dest['N°'] } : null };
  });
  if (!d.confirmer) return { plan: plan, type: compte.Type, mail: String(compte.Mail), secours: Math.max(0, libres.length - occupees.length) };

  const resultats = [];
  plan.forEach(function (x) {
    const abo = rows_(S.ABOS).find(function (a) { return a.ID === x.aboId; });
    const ancienne = rows_(S.PLACES).find(function (p) { return p.ID === x.placeId; });
    if (ancienne) { ancienne['État'] = 'Supprimée'; ancienne.Client = ''; ancienne.Abonnement = ''; update_(S.PLACES, ancienne); }
    if (!abo) return;
    const client = rows_(S.CLIENTS).find(function (c) { return c.ID === abo.Client; });
    abo.Place = ''; abo.Compte = '';
    if (x.vers) placer_(abo, client, P, x.vers.id); else abo.Statut = 'En attente';
    abo['Mis à jour'] = now_();
    update_(S.ABOS, abo);
    const s = sessions_().find(function (z) { return z.code === String(abo.Session); }) || { libelle: abo.Type, type: abo.Type };
    resultats.push(resultat_({ abo: abo, genre: 'Nouveau' }, client, s, P));
  });
  rows_(S.PLACES).filter(function (p) { return p.Compte === compte.ID && p['État'] !== 'Supprimée'; })
    .forEach(function (p) { p['État'] = 'Supprimée'; update_(S.PLACES, p); });
  compte.Statut = 'Bloqué'; compte['Bloqué le'] = now_();
  update_(S.COMPTES, compte);
  hist_('Compte', compte.ID, 'Bloqué', compte.Mail + ' · ' + plan.length + ' client(s) déplacé(s)');
  if (compte.Carte) hist_('Carte', compte.Carte, 'Compte retiré', compte.Mail + ' (bloqué)');
  return { resultats: resultats };
}

/* ---------- Indicateurs business ---------- */

function business_(P, sess, clients, abos) {
  const t = today_();
  const S1 = {}; sess.forEach(function (s) { S1[s.code] = s; });
  const mensuel = function (code) { const s = S1[code]; return s ? s.prix / Math.max(1, s.duree) : 0; };
  const prixType = {};
  rows_(S.TYPES).forEach(function (ty) {
    const l = sess.filter(function (s) { return s.type === ty.Type && s.duree === 1; });
    const src = l.length ? l : sess.filter(function (s) { return s.type === ty.Type; });
    prixType[ty.Type] = src.length ? src.reduce(function (a, s) { return a + s.prix / Math.max(1, s.duree); }, 0) / src.length : 0;
  });
  const aboIx = index_(abos);
  const places = rows_(S.PLACES).filter(function (p) { return p['État'] !== 'Supprimée'; });

  // Rentabilité par compte
  const rentabilite = rows_(S.COMPTES).filter(function (c) { return c.Statut === 'Actif'; }).map(function (c) {
    const pls = places.filter(function (p) { return p.Compte === c.ID; });
    // Seules les places des clients à jour (actifs ou dans les 3 jours de grâce) rapportent.
    // Une place occupée par un client expiré depuis plus de 3 jours ou qui te doit compte comme une perte.
    let rev = 0, occ = 0, impayes = 0, perte = 0;
    pls.forEach(function (p) {
      const a = aboIx[p.Abonnement];
      if (p['État'] !== 'Occupée' || !a || EN_COURS.indexOf(a.Statut) < 0) return;
      if (a.Statut === 'Actif' || a.Statut === 'Grâce') { occ++; rev += mensuel(a.Session); }
      else { impayes++; perte += mensuel(a.Session) || (prixType[c.Type] || 0); }
    });
    const vides = pls.length - occ - impayes, cout = Number(c['Coût mensuel'] || 0);
    return { id: c.ID, mail: String(c.Mail), type: c.Type, places: pls.length, occupees: occ, payees: occ, impayes: impayes, vides: vides,
      revenu: Math.round(rev), cout: cout, profit: Math.round(rev - cout), perte: Math.round(perte),
      manque: Math.round(vides * (prixType[c.Type] || 0) + perte) };
  }).sort(function (a, b) { return b.profit - a.profit; });

  // Taux de renouvellement (mois en cours, sinon mois précédent)
  const g = gains_(6);
  // Taux fiable seulement avec assez de données (au moins 5 renouvellements ou départs)
  const taux = g[0].base >= 5 ? g[0].taux : (g[1].base >= 5 ? g[1].taux : null);

  // Prévision sur 30 jours
  const fin = addDays_(t, 30).getTime();
  const aRenouveler = abos.filter(function (a) {
    const e = day_(a.Expiration);
    return (a.Statut === 'Actif' || a.Statut === 'Grâce') && e && e.getTime() <= fin;
  });
  const brut = aRenouveler.reduce(function (s, a) { return s + (S1[a.Session] ? S1[a.Session].prix : 0); }, 0);
  const prevision = { nombre: aRenouveler.length, brut: brut, attendu: Math.round(brut * (taux != null ? taux / 100 : 0.9)) };

  // Anciens clients à relancer (partis depuis moins de 90 jours)
  const ventes = rows_(S.VENTES).filter(function (v) { return v.Statut === 'Validée'; });
  const anciens = [];
  let relances = 0, revenus = 0;
  clients.forEach(function (c) {
    if (c.Statut === 'Supprimé') return;
    const sesAbos = abos.filter(function (a) { return a.Client === c.ID; });
    const relance = c['Relancé le'] ? new Date(c['Relancé le']).getTime() : 0;
    if (relance) {
      relances++;
      if (ventes.some(function (v) { return v.Client === c.ID && new Date(v.Date).getTime() > relance; })) revenus++;
    }
    if (sesAbos.some(function (a) { return EN_COURS.indexOf(a.Statut) >= 0; })) return;
    const fins = sesAbos.filter(function (a) { return a.Statut === 'Terminé'; })
      .sort(function (a, b) { return new Date(b['Mis à jour']) - new Date(a['Mis à jour']); });
    if (!fins.length) return;
    const jours = joursEntre_(day_(fins[0]['Mis à jour']), t);
    if (jours > 90) return;
    anciens.push({ id: c.ID, nom: c.Nom, type: fins[0].Type, jours: jours,
      renouvellements: ventes.filter(function (v) { return v.Client === c.ID && v.Genre === 'Renouvellement'; }).length,
      relanceLe: c['Relancé le'] || '', whatsapp: waNum_(c), session: fins[0].Session, paiement: lienPaiement_(prixSession_(fins[0].Session), P) });
  });
  anciens.sort(function (a, b) { return a.jours - b.jours; });
  const potentiel = Math.round(anciens.reduce(function (s, a) { return s + mensuel(a.session); }, 0));

  // Alertes d'achat
  const achats = [];
  rows_(S.TYPES).forEach(function (ty) {
    const ids = rows_(S.COMPTES).filter(function (c) { return c.Statut === 'Actif' && c.Type === ty.Type; }).map(function (c) { return c.ID; });
    const libres = places.filter(function (p) { return ids.indexOf(p.Compte) >= 0 && p['État'] === 'Libre'; }).length;
    const attente = abos.filter(function (a) { return a.Statut === 'En attente' && a.Type === ty.Type; }).length;
    const cout = Number(ty['Coût mensuel par défaut'] || 0), pm = prixType[ty.Type] || 0;
    if (attente > 0 || (ids.length && libres === 0)) {
      achats.push({ type: ty.Type, attente: attente, libres: libres, cout: cout, rentableDes: pm ? Math.ceil(cout / pm) : null });
    }
  });

  return { rentabilite: rentabilite, taux: taux, prevision: prevision, anciens: anciens, potentiel: potentiel,
    relances: relances, revenus: revenus, achats: achats, sante: sante_(clients, abos, places, mensuel, taux, g) };
}

/** Santé du business : durée de vie des clients, départs, fidélité, valeur d'un client. */
function sante_(clients, abos, places, mensuel, taux, g) {
  const t = today_(), J = 864e5, VIVANT = ['Actif', 'Grâce', 'En attente', 'Dette'];
  const vivants = abos.filter(function (a) { return a.Statut !== 'Annulé' && a.Statut !== 'Inventaire' && a['Début']; });
  const parClient = {};
  vivants.forEach(function (a) { (parClient[a.Client] = parClient[a.Client] || []).push(a); });
  const fiches = [];
  clients.forEach(function (c) {
    const l = parClient[c.ID]; if (!l || c.Statut === 'Supprimé') return;
    const debut = Math.min.apply(null, l.map(function (a) { return day_(a['Début']).getTime(); }));
    const actif = l.some(function (a) { return VIVANT.indexOf(a.Statut) >= 0; });
    const fin = actif ? t.getTime() : Math.max.apply(null, l.map(function (a) { const e = day_(a.Expiration); return e ? e.getTime() : debut; }));
    const finAbo = l.filter(function (a) { return a.Statut === 'Terminé'; }).map(function (a) { return new Date(a['Mis à jour']).getTime(); })
      .concat(l.filter(function (a) { return a.Statut === 'À retirer'; }).map(function (a) { return day_(a.Expiration).getTime(); }));
    fiches.push({ id: c.ID, nom: c.Nom, actif: actif, debut: debut, mois: Math.max(1, Math.round((fin - debut) / (30.4 * J))),
      mensuel: l.filter(function (a) { return a.Statut === 'Actif' || a.Statut === 'Grâce' || a.Statut === 'En attente'; }).reduce(function (s, a) { return s + mensuel(a.Session); }, 0),
      parti: actif ? 0 : (finAbo.length ? Math.max.apply(null, finAbo) : fin) });
  });
  const moy = function (l) { return l.length ? Math.round(10 * l.reduce(function (s, x) { return s + x.mois; }, 0) / l.length) / 10 : null; };
  const actifs = fiches.filter(function (f) { return f.actif; }), partis = fiches.filter(function (f) { return !f.actif; });
  const depuis30 = t.getTime() - 30 * J;
  const departs30 = partis.filter(function (f) { return f.parti >= depuis30; }).length;
  const nouveaux30 = fiches.filter(function (f) { return f.debut >= depuis30; }).length;
  const revenuMensuel = actifs.reduce(function (s, f) { return s + f.mensuel; }, 0);
  const panier = actifs.length ? Math.round(revenuMensuel / actifs.length) : 0;
  const vie = moy(fiches);
  const churn = (actifs.length + departs30) ? Math.round(100 * departs30 / (actifs.length + departs30)) : null;
  const unMois = partis.length ? Math.round(100 * partis.filter(function (f) { return f.mois <= 1; }).length / partis.length) : null;
  const aboIx = index_(abos), payee = function (p) { const a = aboIx[p.Abonnement]; return p['État'] === 'Occupée' && a && (a.Statut === 'Actif' || a.Statut === 'Grâce'); };
  const tot = places.length, occ = places.filter(payee).length;
  const impayees = places.filter(function (p) { return p['État'] === 'Occupée' && !payee(p); }).length;
  const occupation = tot ? Math.round(100 * occ / tot) : 0;
  const fideles = actifs.filter(function (f) { return f.mois >= 6; }).sort(function (a, b) { return b.mois - a.mois; });
  const tr = taux != null ? taux : (churn != null ? 100 - churn : 80);
  const solde = nouveaux30 - departs30;
  const score = Math.round(0.5 * tr + 0.3 * occupation + 0.2 * (solde > 0 ? 100 : solde === 0 ? 70 : 40));
  return {
    score: score, niveau: score >= 75 ? 'Excellente' : score >= 55 ? 'Bonne' : 'À surveiller',
    actifs: actifs.length, vie: vie, vieActifs: moy(actifs), viePartis: moy(partis),
    panier: panier, valeur: vie != null ? Math.round(panier * vie) : 0, revenuMensuel: Math.round(revenuMensuel),
    departs30: departs30, nouveaux30: nouveaux30, solde: solde, churn: churn, partisApres1: unMois,
    occupation: occupation, impayees: impayees, placesLibres: places.filter(function (p) { return p['État'] === 'Libre'; }).length,
    aRetirer: abos.filter(function (a) { return a.Statut === 'À retirer'; }).length,
    total: fiches.length, importes: fiches.filter(function (f) { return parClient[f.id].every(function (a) { return a.Import; }); }).length,
    fideles: fideles.length, top: fideles.slice(0, 5).map(function (f) { return { id: f.id, nom: f.nom, mois: f.mois }; }),
    evolution: g.map(function (x, i) {
      const a = Date.UTC(t.getUTCFullYear(), t.getUTCMonth() - i, 1), b = Date.UTC(t.getUTCFullYear(), t.getUTCMonth() - i + 1, 1);
      return { mois: x.mois, nouveaux: fiches.filter(function (f) { return f.debut >= a && f.debut < b; }).length,
        departs: partis.filter(function (f) { return f.parti >= a && f.parti < b; }).length };
    })
  };
}

/* ---------- Sauvegarde mensuelle ---------- */

function sauvegardeMensuelle() {
  const P = params_();
  const id = prop_('SS_ID');
  const nom = 'ICONIC-sauvegarde-' + Utilities.formatDate(now_(), TZ, 'yyyy-MM-dd') + '.xlsx';
  const rep = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true
  });
  if (rep.getResponseCode() !== 200) throw new Error('Sauvegarde impossible (' + rep.getResponseCode() + ').');
  MailApp.sendEmail({ to: P.email, name: MARQUE, subject: '💾 Sauvegarde ' + MARQUE + ' – ' + fmt_(now_()),
    htmlBody: emailHtml_('Ta sauvegarde', [{ titre: 'Fichier joint', lignes: ['Ton tableau complet au format Excel. Garde cet e-mail : il te permet de tout restaurer.'] }], false),
    attachments: [rep.getBlob().setName(nom)] });
  return { ok: true, fichier: nom };
}

/* ═════════════════════ VERSION 2.2 ═════════════════════ */

/** Surveillance de l'Android : alerte si aucun SMS depuis 48 h alors qu'il en arrive d'habitude. */
function surveillanceSms_() {
  const sms = rows_(S.VENTES).filter(function (v) { return v['Trans ID'] && v['Trans ID'] !== 'MANUEL'; });
  if (sms.length < 5) return '';
  const t = now_().getTime();
  const recents = sms.filter(function (v) { return t - new Date(v.Date).getTime() < 30 * 86400000; }).length;
  const dernier = Math.max.apply(null, sms.map(function (v) { return new Date(v.Date).getTime(); }));
  const h = Math.floor((t - dernier) / 3600000);
  if (recents >= 5 && h >= 48) return '⚠️ Aucun SMS Orange Money reçu depuis ' + Math.floor(h / 24) + ' jours : vérifie que l\'Android est allumé, connecté à Internet et que l\'app de transfert fonctionne.';
  return '';
}

/* ---------- Inventaire (clients importés à vérifier) ---------- */

function infoImport_(a) { try { return JSON.parse(a.Import || '{}'); } catch (e) { return {}; } }

function inventaire_(d) {
  const P = params_();
  const abo = rows_(S.ABOS).find(function (a) { return a.ID === d.aboId; });
  if (!abo || abo.Statut !== 'Inventaire') throw new Error('Ce client a déjà été vérifié.');
  const info = infoImport_(abo);
  const client = rows_(S.CLIENTS).find(function (c) { return c.ID === abo.Client; });
  if (d.choix === 'parti') {
    abo.Statut = 'Terminé'; abo['Mis à jour'] = now_();
    update_(S.ABOS, abo);
    return { ok: true };
  }
  const duree = Number(d.duree || info.duree || 1);
  const date = d.date ? day_(new Date(d.date)) : day_(abo['Début']);
  abo['Début'] = date;
  abo.Expiration = addMonths_(date, duree);
  const ses = sessions_().filter(function (x) { return x.type === abo.Type; });
  const s = ses.find(function (x) { return x.duree === duree; }) || ses[0];
  if (s) abo.Session = s.code;
  const pl = placerInventaire_(abo, info, d.placeId);
  abo.Place = pl.ID; abo.Compte = pl.Compte;
  abo.Statut = d.choix === 'dette' ? 'Dette' : statutPour_(abo, P);
  abo['Mis à jour'] = now_();
  info.placeCreee = pl.creee ? pl.ID : '';
  abo.Import = JSON.stringify(info);
  update_(S.ABOS, abo);
  hist_('Abonnement', abo.ID, 'Inventaire', (client ? client.Nom : '') + ' · ' + d.choix + ' · jusqu\'au ' + fmt_(abo.Expiration));
  return { ok: true, statut: abo.Statut, expiration: abo.Expiration, place: pl['N°'] };
}

/** Place un client importé sur son compte, en reprenant son profil/code si on le connaît. */
function placerInventaire_(abo, info, placeId) {
  const compteId = info.compte || abo.Compte;
  let places = rows_(S.PLACES).filter(function (p) { return p.Compte === compteId && p['État'] !== 'Supprimée'; });
  let choix = null;
  if (placeId) {
    choix = places.find(function (p) { return p.ID === placeId && p['État'] !== 'Occupée'; });
    if (!choix) throw new Error('Cette place n\'est pas disponible.');
  }
  const prof = String(info.profil || '').trim().toLowerCase(), code = String(info.code || '').trim();
  let creee = false;
  if (!choix && prof) {
    const meme = places.filter(function (p) { return String(p.Profil).trim().toLowerCase() === prof && (!code || String(p.Code) === code); });
    choix = meme.find(function (p) { return p['État'] !== 'Occupée'; });
    if (!choix && meme.length) choix = nouvellePlace_(compteId, info.profil, code, places);
  }
  if (!choix) choix = places.find(function (p) { return p['État'] === 'À préparer'; });
  if (!choix) choix = places.find(function (p) { return p['État'] === 'Libre' && !p.Profil; });
  if (!choix) choix = places.find(function (p) { return p['État'] === 'Libre'; });
  if (!choix) { choix = nouvellePlace_(compteId, info.profil || '', code, places); }
  if (choix._nouvelle) creee = true;
  if (!choix.Profil && info.profil) choix.Profil = info.profil;
  if (!choix.Code && code) choix.Code = code;
  choix['État'] = 'Occupée'; choix.Client = abo.Client; choix.Abonnement = abo.ID; choix.Depuis = now_();
  update_(S.PLACES, choix);
  choix.creee = creee;
  return choix;
}

function nouvellePlace_(compteId, profil, code, places) {
  const n = places.reduce(function (m, p) { return Math.max(m, Number(p['N°'])); }, 0) + 1;
  const p = { 'ID': nextId_('PL'), 'Compte': compteId, 'N°': n, 'Profil': profil || '', 'Code': code || '', 'État': sansProfil_(compteId) ? 'Libre' : 'À préparer', 'Client': '', 'Abonnement': '', 'Depuis': now_() };
  insert_(S.PLACES, p);
  const c = rows_(S.COMPTES).find(function (x) { return x.ID === compteId; });
  if (c) { c.Places = Number(c.Places || 0) + 1; update_(S.COMPTES, c); }
  p._nouvelle = true;
  return p;
}

/** Tout accepter selon les notes : ❌ → « Me doit », sinon « À jour » avec la date notée. */
function inventaireAuto_() {
  const l = rows_(S.ABOS).filter(function (a) { return a.Statut === 'Inventaire'; });
  let ok = 0, sansDate = 0;
  l.forEach(function (a) {
    const info = infoImport_(a);
    if (!a['Début']) { sansDate++; return; }
    inventaire_({ aboId: a.ID, choix: Number(info.dus) > 0 ? 'dette' : 'ajour', date: day_(a['Début']).toISOString().slice(0, 10), duree: info.duree || 1 });
    ok++;
  });
  return { faits: ok, sansDate: sansDate };
}

function annulerInventaire_(d) {
  const abo = rows_(S.ABOS).find(function (a) { return a.ID === d.aboId; });
  if (!abo) throw new Error('Introuvable.');
  const info = infoImport_(abo);
  if (abo.Place) {
    const pl = rows_(S.PLACES).find(function (p) { return p.ID === abo.Place; });
    if (pl && pl.Abonnement === abo.ID) {
      pl.Client = ''; pl.Abonnement = ''; pl['État'] = ((pl.Profil && pl.Code) || sansProfil_(pl.Compte)) ? 'Libre' : 'À préparer';
      if (info.placeCreee === pl.ID) {
        pl['État'] = 'Supprimée';
        const c = rows_(S.COMPTES).find(function (x) { return x.ID === pl.Compte; });
        if (c) { c.Places = Math.max(0, Number(c.Places || 0) - 1); update_(S.COMPTES, c); }
      }
      update_(S.PLACES, pl);
    }
  }
  abo.Statut = 'Inventaire'; abo.Place = ''; abo.Compte = info.compte || abo.Compte; info.placeCreee = '';
  abo.Import = JSON.stringify(info);
  update_(S.ABOS, abo);
  return { ok: true };
}

/* ---------- Réglages depuis l'app ---------- */

const REGLAGES = ['ONESIGNAL_APP_ID', 'ONESIGNAL_API_KEY', 'NOTIF_EMAIL', 'VALIDATION_AUTO', 'GRACE_JOURS', 'RAPPEL_CLIENT_JOURS', 'RAPPEL_COMPTE_JOURS', 'MARGE', 'NUMERO_DEPOT', 'USSD_PAIEMENT'];

function lireReglages_() {
  const o = {};
  rows_(S.PARAM).forEach(function (r) { if (REGLAGES.indexOf(r['Clé']) >= 0) o[r['Clé']] = String(r.Valeur); });
  if (o.ONESIGNAL_API_KEY) o.ONESIGNAL_API_KEY = '•••• ' + o.ONESIGNAL_API_KEY.slice(-4);
  return o;
}

function sauverReglages_(d) {
  const rows = rows_(S.PARAM);
  Object.keys(d || {}).forEach(function (k) {
    if (REGLAGES.indexOf(k) < 0) return;
    let v = String(d[k] == null ? '' : d[k]).trim();
    if (k === 'ONESIGNAL_API_KEY' && (v === '' || v.indexOf('••••') === 0)) return; // clé inchangée
    if (['GRACE_JOURS', 'RAPPEL_CLIENT_JOURS', 'RAPPEL_COMPTE_JOURS', 'MARGE'].indexOf(k) >= 0 && (v === '' || isNaN(Number(v)))) throw new Error('Valeur invalide pour ' + k);
    const r = rows.find(function (x) { return x['Clé'] === k; });
    if (r) { r.Valeur = v; update_(S.PARAM, r); }
    else insert_(S.PARAM, { 'Clé': k, 'Valeur': v, 'Description': '' });
  });
  return lireReglages_();
}

/** Assistant « Compléter mes comptes » : date, carte et profils/codes d'un compte en une fois. */
function completerCompte_(d) {
  const c = rows_(S.COMPTES).find(function (x) { return x.ID === d.id; });
  if (!c) throw new Error('Compte introuvable.');
  const m = {};
  if (d.expiration) m.expiration = d.expiration;
  if (d.carte !== undefined && String(d.carte) !== String(c.Carte || '')) m.carte = d.carte;
  if (d.mdp !== undefined && String(d.mdp) !== String(c['Mot de passe'] || '')) m.mdp = d.mdp;
  if (Object.keys(m).length) { m.id = c.ID; modifierCompte_(m); }
  let prets = 0;
  (d.places || []).forEach(function (p) {
    if (String(p.profil || '').trim() && String(p.code || '').trim()) { preparerPlace_({ placeId: p.placeId, profil: p.profil, code: p.code }); prets++; }
  });
  return { ok: true, places: prets };
}

/** Test d'une notification depuis l'app. */
function testNotif_() {
  const P = params_();
  if (!P.osAppId || !P.osKey) throw new Error('Renseigne d\'abord l\'App ID et la clé API OneSignal.');
  push_('🔔 Test ICONIC', 'Les notifications fonctionnent !', 'accueil', P);
  return { ok: true };
}

/* ═════════════════════ OUTILS ═════════════════════ */

let SS_CACHE = null;
function props_() { return PropertiesService.getScriptProperties(); }
function prop_(k) { return props_().getProperty(k); }
function ss_() {
  if (SS_CACHE) return SS_CACHE;
  const id = prop_('SS_ID');
  if (!id) throw new Error('Exécute d\'abord la fonction « installer ».');
  SS_CACHE = SpreadsheetApp.openById(id);
  return SS_CACHE;
}
function sh_(n) {
  const s = ss_().getSheetByName(n);
  if (!s) throw new Error('Onglet manquant : ' + n + '. Relance « installer ».');
  return s;
}
// Mémoire de lecture : chaque onglet n'est lu qu'une fois par appel (et relu après une écriture).
let RCACHE = {};
function rows_(n) {
  if (!RCACHE[n]) RCACHE[n] = lire_(n);
  return RCACHE[n].map(function (o) { return Object.assign({}, o); });
}
function lire_(n) {
  const h = H[n], sh = sh_(n), last = sh.getLastRow();
  if (last < 2) return [];
  const v = sh.getRange(2, 1, last - 1, h.length).getValues();
  const out = [];
  v.forEach(function (r, i) {
    if (r[0] === '' || r[0] === null) return;
    const o = { _row: i + 2 };
    h.forEach(function (k, j) { o[k] = r[j]; });
    out.push(o);
  });
  return out;
}
function ligne_(n, o) {
  const txt = TEXTE[n] || [];
  return H[n].map(function (k) {
    let v = o[k];
    if (v === undefined || v === null) v = '';
    if (txt.indexOf(k) >= 0 && v !== '') v = "'" + String(v);
    return v;
  });
}
function insert_(n, o) {
  delete RCACHE[n];
  const sh = sh_(n);
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, H[n].length).setValues([ligne_(n, o)]);
  o._row = r;
}
function update_(n, o) {
  if (!o._row) throw new Error('Ligne inconnue (' + n + ')');
  if (RCACHE[n]) {
    const i = RCACHE[n].findIndex(function (x) { return x._row === o._row; });
    if (i >= 0) { const c = Object.assign({}, o); RCACHE[n][i] = c; } else delete RCACHE[n];
  }
  sh_(n).getRange(o._row, 1, 1, H[n].length).setValues([ligne_(n, o)]);
}
function index_(l) { const o = {}; l.forEach(function (x) { o[x.ID] = x; }); return o; }
function hist_(objet, id, action, detail) { insert_(S.HIST, { 'Date': now_(), 'Objet': objet, 'ID': id, 'Action': action, 'Détail': detail }); }
function journal_(niv, msg) { try { insert_(S.JOURNAL, { 'Date': now_(), 'Niveau': niv, 'Message': msg }); } catch (e) { Logger.log(msg); } }
function nextId_(prefix) {
  const p = props_(), k = 'SEQ_' + prefix;
  const n = Number(p.getProperty(k) || 0) + 1;
  p.setProperty(k, String(n));
  return prefix + '-' + ('0000' + n).slice(-Math.max(4, String(n).length));
}
function cle_(n) {
  const c = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789';
  let s = '';
  for (let i = 0; i < n; i++) s += c.charAt(Math.floor(Math.random() * c.length));
  return s;
}
function liens_() {
  const p = props_();
  const url = p.getProperty('APP_URL') || ScriptApp.getService().getUrl();
  return { app: url + '?k=' + p.getProperty('APP_KEY'), sms: url + '?token=' + p.getProperty('SMS_TOKEN') };
}
function params_() {
  const o = {};
  rows_(S.PARAM).forEach(function (r) { o[r['Clé']] = r.Valeur; });
  const n = function (v, d) { return (v === '' || v === undefined || isNaN(Number(v))) ? d : Number(v); };
  return {
    email: String(o.EMAIL || Session.getEffectiveUser().getEmail()), grace: n(o.GRACE_JOURS, 3),
    rappelClient: n(o.RAPPEL_CLIENT_JOURS, 3), rappelCompte: n(o.RAPPEL_COMPTE_JOURS, 5), marge: n(o.MARGE, 100),
    maxCarte: n(o.MAX_COMPTES_PAR_CARTE, 3), indicatif: String(o.INDICATIF || '226'),
    preparer: n(o.PREPARER_RAPPEL_JOURS, 1), numeroDepot: String(o.NUMERO_DEPOT || ''),
    notifEmail: String(o.NOTIF_EMAIL === undefined ? 'oui' : o.NOTIF_EMAIL).trim().toLowerCase() !== 'non',
    osAppId: String(o.ONESIGNAL_APP_ID || '').trim(), osKey: String(o.ONESIGNAL_API_KEY || '').trim(),
    validationAuto: String(o.VALIDATION_AUTO === undefined ? 'oui' : o.VALIDATION_AUTO).trim().toLowerCase() !== 'non',
    ussd: String(o.USSD_PAIEMENT || '*144*2*1*{numero}*{montant}#').trim()
  };
}
function sessions_() {
  return rows_(S.SESSIONS).map(function (s) {
    return { code: String(s.Code), type: s.Type, libelle: s['Libellé'], duree: Number(s['Durée (mois)']), prix: Number(s.Prix) };
  });
}
function candidats_(montant, sess, marge) {
  return sess.filter(function (s) { return montant >= s.prix && montant <= s.prix + marge; });
}
/** Combinaisons de 2 sessions dont la somme correspond au montant (ex. Films + Musique, ou 2 × Films). */
function combinaisons_(montant, sess, marge) {
  const out = [];
  for (let i = 0; i < sess.length; i++) for (let j = i; j < sess.length; j++) {
    const t = sess[i].prix + sess[j].prix;
    if (montant >= t && montant <= t + marge * 2) out.push(sess[i].code + '+' + sess[j].code);
  }
  return out;
}
function normNum_(n) {
  let x = String(n || '').replace(/\D/g, '');
  if (x.indexOf('00') === 0) x = x.slice(2);
  if (x.length > 8 && x.indexOf('226') === 0) x = x.slice(3);
  return x;
}
function numsOf_(c) { return String(c['Numéros'] || '').split(/[,;\s]+/).map(normNum_).filter(String); }
function clientParNumero_(clients, num) {
  num = normNum_(num);
  if (!num) return null;
  return clients.find(function (c) { return c.Statut !== 'Supprimé' && numsOf_(c).indexOf(num) >= 0; }) || null;
}
function sansAccent_(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z ]/g, ' ').trim(); }
function score_(a, b) {
  a = sansAccent_(a); b = sansAccent_(b);
  if (!a || !b) return 0;
  if (a === b) return 100;
  let s = 0;
  a.split(/\s+/).forEach(function (x) {
    b.split(/\s+/).forEach(function (y) {
      if (x === y) s += 50;
      else if (x.length > 2 && y.length > 2 && (x.indexOf(y) === 0 || y.indexOf(x) === 0)) s += 30;
    });
  });
  return s;
}
function wa_(tel, msg, P) {
  const n = normNum_(tel);
  return 'https://wa.me/' + (n ? P.indicatif + n : '') + '?text=' + encodeURIComponent(msg);
}
function now_() { return new Date(); }
function day_(d) {
  if (!d) return null;
  d = new Date(d);
  if (isNaN(d.getTime())) return null;
  return new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
}
function today_() { return day_(now_()); }
function addDays_(d, n) { return new Date(day_(d).getTime() + n * 86400000); }
function addMonths_(d, n) {
  const x = day_(d), jour = x.getUTCDate();
  const r = new Date(Date.UTC(x.getUTCFullYear(), x.getUTCMonth() + n, 1));
  const dernier = new Date(Date.UTC(r.getUTCFullYear(), r.getUTCMonth() + 1, 0)).getUTCDate();
  r.setUTCDate(Math.min(jour, dernier));
  return r;
}
function joursEntre_(a, b) { if (!a || !b) return 0; return Math.round((day_(b).getTime() - day_(a).getTime()) / 86400000); }
function fmt_(d) { d = day_(d); return d ? Utilities.formatDate(d, 'UTC', 'dd/MM/yyyy') : ''; }
function F_(n) { return String(Math.round(Number(n || 0))).replace(/\B(?=(\d{3})+(?!\d))/g, ' ') + ' F'; }
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function emailHtml_(titre, blocs, bouton, onglet) {
  let h = '<div style="background:#0a0a0a;padding:20px;font-family:Arial,sans-serif;color:#eeeeee;font-size:14px">' +
    '<div style="text-align:center;font-family:Georgia,serif;letter-spacing:5px;color:#f3dfa2;font-size:22px">◆ ICONIC' +
    '<div style="font-size:10px;letter-spacing:7px;color:#d4af37">COLLECTION</div></div>' +
    '<h2 style="color:#f3dfa2;font-family:Georgia,serif;font-size:18px;margin-top:22px">' + titre + '</h2>';
  blocs.forEach(function (b) {
    h += '<div style="border:1px solid #3a3115;border-radius:10px;padding:12px;margin:12px 0;background:#141414">' +
      '<div style="color:#d4af37;font-weight:bold;margin-bottom:6px">' + b.titre + '</div>' +
      b.lignes.map(function (l) { return '<div style="padding:7px 0;border-top:1px solid #222">' + l + '</div>'; }).join('') + '</div>';
  });
  if (bouton && prop_('APP_URL')) {
    h += '<p style="text-align:center;margin-top:20px"><a href="' + liens_().app + '&tab=' + (onglet || 'accueil') +
      '" style="background:#d4af37;color:#111;padding:12px 22px;border-radius:8px;text-decoration:none;font-weight:bold">Ouvrir l\'application</a></p>';
  }
  return h + '</div>';
}


/* ═════════════════════ MISES À JOUR AUTOMATIQUES (2.3) ═════════════════════
 * Le nouveau code est publié sur GitHub (dossier « serveur »). Toutes les heures, le script
 * vérifie s'il existe une version plus récente, la colle à la place de Code.gs et publie
 * lui-même une « Nouvelle version » du déploiement. Rien d'autre n'est modifié (Import, Index…).
 * Activation (une seule fois) : exécuter activerMisesAJour.
 */

const DEPOT_MAJ = 'https://raw.githubusercontent.com/moisekarfo80-design/iconic/main/serveur/';

function activerMisesAJour() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'verifierMiseAJour') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('verifierMiseAJour').timeBased().everyHours(1).create();
  // Vérifie que l'API Apps Script est activée et que les autorisations sont accordées
  const rep = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId() + '/deployments',
    { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (rep.getResponseCode() !== 200) {
    throw new Error('Active d\'abord « API Google Apps Script » sur script.google.com/home/usersettings, attends 2 minutes et réessaie. (' + rep.getResponseCode() + ')');
  }
  Logger.log('✅ Mises à jour automatiques activées (version ' + VERSION + ').');
  const r = miseAJour_(true);
  Logger.log(r.installee ? '⬆️ Version ' + r.version + ' installée.' : 'Tu as déjà la dernière version.');
}

function verifierMiseAJour() {
  try { miseAJour_(false); } catch (e) { journal_('ERREUR', 'Mise à jour : ' + e.message); }
}

function versionPlusRecente_(a, b) {
  const x = String(a).split('.').map(Number), y = String(b).split('.').map(Number);
  for (let i = 0; i < 3; i++) { if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) > (y[i] || 0); }
  return false;
}

function infoVersion_() {
  return { version: VERSION, derniere: prop_('MAJ_DERNIERE') || '', le: prop_('MAJ_LE') || '', erreur: prop_('MAJ_ERREUR') || '' };
}

function miseAJour_(manuel) {
  // Avec le chargeur (Code.gs court), la mise à jour se fait depuis GitHub sans l'API Apps Script.
  if (typeof chargeurMaj_ === 'function') return chargeurMaj_(false);
  const props = props_();
  const lire = function (f) {
    const r = UrlFetchApp.fetch(DEPOT_MAJ + f + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('GitHub injoignable (' + f + ' : ' + r.getResponseCode() + ')');
    return r.getContentText();
  };
  const info = JSON.parse(lire('version.json'));
  if (!versionPlusRecente_(info.version, VERSION)) return { installee: false, version: VERSION };
  const source = lire('Code.gs');
  if (source.indexOf("const VERSION = '" + info.version + "'") < 0) throw new Error('Fichier de mise à jour incomplet, réessaie plus tard.');

  const api = 'https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId();
  const appel = function (methode, chemin, corps) {
    const o = { method: methode, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true, contentType: 'application/json' };
    if (corps) o.payload = JSON.stringify(corps);
    const r = UrlFetchApp.fetch(api + chemin, o), c = r.getResponseCode();
    if (c !== 200) throw new Error('API Apps Script (' + c + ') : ' + r.getContentText().slice(0, 300));
    return JSON.parse(r.getContentText() || '{}');
  };
  try {
    // 1. Remplacer uniquement le fichier du serveur
    const contenu = appel('get', '/content');
    const fic = contenu.files.find(function (f) { return f.type === 'SERVER_JS' && String(f.name).toLowerCase() === 'code'; })
      || contenu.files.find(function (f) { return f.type === 'SERVER_JS' && String(f.source).indexOf('function doPost') >= 0; });
    if (!fic) throw new Error('Fichier Code introuvable dans le projet.');
    const ancienne = fic.source;
    fic.source = source;
    appel('put', '/content', { files: contenu.files });
    // 2. Créer une version et y faire pointer le déploiement de l'app
    const v = appel('post', '/versions', { description: 'ICONIC ' + info.version + ' (automatique)' });
    const deps = (appel('get', '/deployments').deployments || []).filter(function (d) {
      return d.deploymentConfig && d.deploymentConfig.versionNumber && (d.entryPoints || []).some(function (e) { return e.entryPointType === 'WEB_APP'; });
    });
    deps.forEach(function (d) {
      appel('put', '/deployments/' + d.deploymentId, { deploymentConfig: {
        scriptId: ScriptApp.getScriptId(), versionNumber: v.versionNumber, manifestFileName: 'appsscript', description: 'ICONIC ' + info.version } });
    });
    props.setProperty('MAJ_PRECEDENTE', JSON.stringify({ version: VERSION, numero: deps.length ? deps[0].deploymentConfig.versionNumber : 0 }));
    try { DriveApp.createFile('ICONIC-Code-' + VERSION + '.gs', ancienne, 'text/plain'); } catch (e) {}
    props.setProperty('MAJ_DERNIERE', info.version); props.setProperty('MAJ_LE', now_().toISOString()); props.deleteProperty('MAJ_ERREUR');
    // 3. Migrations éventuelles : exécutées par le nouveau code dans 1 minute
    ScriptApp.newTrigger('apresMiseAJour').timeBased().after(60 * 1000).create();
    hist_('Système', '', 'Mise à jour', VERSION + ' → ' + info.version + (info.notes ? ' · ' + info.notes : ''));
    try { push_('⬆️ ICONIC mis à jour', 'Version ' + info.version + (info.notes ? ' : ' + info.notes : ''), 'accueil', params_()); } catch (e) {}
    return { installee: true, version: info.version, notes: info.notes || '' };
  } catch (e) {
    props.setProperty('MAJ_ERREUR', now_().toISOString() + ' · ' + e.message);
    throw e;
  }
}

/** Lancée une minute après une mise à jour, avec le nouveau code : ajoute les colonnes, déclencheurs… */
function apresMiseAJour() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'apresMiseAJour') ScriptApp.deleteTrigger(t); });
  installer();
  migrerMusique_();
  journal_('INFO', 'Migration ' + VERSION + ' terminée.');
  hist_('Système', '', 'Mise à jour', 'Version ' + VERSION + ' installée');
  try { push_('⬆️ ICONIC mis à jour', 'Version ' + VERSION + ' installée', 'accueil', params_()); } catch (e) {}
}

/** 2.3.4 : les places Musique « à préparer » deviennent libres (pas de profil ni de code). */
function migrerMusique_() {
  const mus = {}; rows_(S.COMPTES).forEach(function (c) { if (/musique/i.test(String(c.Type))) mus[c.ID] = 1; });
  let n = 0;
  rows_(S.PLACES).forEach(function (p) { if (mus[p.Compte] && p['État'] === 'À préparer') { p['État'] = 'Libre'; update_(S.PLACES, p); n++; } });
  if (n) journal_('INFO', n + ' place(s) Musique rendues libres.');
  return n;
}

/** En cas de problème : remet la version précédente de l'app (à exécuter dans l'éditeur). */
function annulerMiseAJour() {
  const p = JSON.parse(prop_('MAJ_PRECEDENTE') || 'null');
  if (!p || !p.numero) throw new Error('Aucune version précédente connue.');
  const api = 'https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId();
  const h = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  const deps = JSON.parse(UrlFetchApp.fetch(api + '/deployments', { headers: h }).getContentText()).deployments || [];
  deps.filter(function (d) { return d.deploymentConfig && d.deploymentConfig.versionNumber && (d.entryPoints || []).some(function (e) { return e.entryPointType === 'WEB_APP'; }); })
    .forEach(function (d) {
      UrlFetchApp.fetch(api + '/deployments/' + d.deploymentId, { method: 'put', headers: h, contentType: 'application/json',
        payload: JSON.stringify({ deploymentConfig: { scriptId: ScriptApp.getScriptId(), versionNumber: p.numero, manifestFileName: 'appsscript', description: 'Retour à ' + p.version } }) });
    });
  Logger.log('↩️ L\'app utilise de nouveau la version ' + p.version + '. Le code précédent est aussi dans ton Google Drive (ICONIC-Code-' + p.version + '.gs).');
}
