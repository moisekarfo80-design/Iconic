/**
 * ICONIC Collection – chargeur (à coller une seule fois à la place de Code.gs).
 *
 * Ce petit fichier ne change plus jamais. Le vrai code de l'app est sur ton GitHub
 * (dossier « serveur »). Le chargeur le télécharge, le vérifie, le garde en mémoire,
 * et va chercher toutes les heures s'il y a une nouvelle version : les mises à jour
 * s'installent seules, sans copier-coller ni « Nouvelle version » à publier.
 *
 * Si GitHub est injoignable, la dernière version qui a marché est utilisée.
 */

const CHARGEUR = '1.0';
const DEPOT_SOURCE = 'https://raw.githubusercontent.com/moisekarfo80-design/iconic/main/serveur/';
const NOMS_EXPORTES = ['doGet', 'doPost', 'api', 'installer', 'activerV2', 'envoyerMesLiens', 'testSms',
  'rappelQuotidien', 'sauvegardeMensuelle', 'apresMiseAJour', 'infoVersion_'];
let ICONIC__ = null;

/* ---------- Points d'entrée (appelés par Google, l'app et les déclencheurs) ---------- */
function doGet(e) { return iconic_().doGet(e); }
function doPost(e) { return iconic_().doPost(e); }
function api(key, action, d) { return iconic_().api(key, action, d); }
function installer() { return iconic_().installer(); }
function activerV2() { return iconic_().activerV2(); }
function envoyerMesLiens() { return iconic_().envoyerMesLiens(); }
function testSms() { return iconic_().testSms(); }
function rappelQuotidien() { return iconic_().rappelQuotidien(); }
function sauvegardeMensuelle() { return iconic_().sauvegardeMensuelle(); }
function apresMiseAJour() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'apresMiseAJour') ScriptApp.deleteTrigger(t); });
  return iconic_().apresMiseAJour();
}

/** À exécuter une fois après avoir collé ce fichier. */
function activerMisesAJour() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'verifierMiseAJour') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('verifierMiseAJour').timeBased().everyHours(1).create();
  const r = chargeurMaj_(true);
  iconic_().installer();
  Logger.log('✅ Mises à jour automatiques activées. Version installée : ' + r.version);
}

function verifierMiseAJour() {
  try { chargeurMaj_(false); } catch (e) { PropertiesService.getScriptProperties().setProperty('MAJ_ERREUR', new Date().toISOString() + ' · ' + e.message); }
}

/** Revenir à la version précédente (si une mise à jour pose problème). */
function annulerMiseAJour() {
  const P = PropertiesService.getScriptProperties();
  const prec = lireMorceaux_(P, 'SRC_PREC');
  if (!prec) throw new Error('Aucune version précédente gardée.');
  const v = P.getProperty('SRC_PREC_VERSION'), refusee = P.getProperty('SRC_VERSION');
  enregistrer_(prec, v);
  P.setProperty('MAJ_BLOQUEE', refusee || '');   // ne pas la réinstaller automatiquement
  Logger.log('↩️ Retour à la version ' + v + '. La version ' + refusee + ' ne sera pas réinstallée ; la suivante, oui.');
}

/* ---------- Mécanique ---------- */

function iconic_() {
  if (ICONIC__) return ICONIC__;
  const src = sourceIconic_();
  ICONIC__ = compiler_(src);
  return ICONIC__;
}

function compiler_(src) {
  return new Function(src + '\n;return {' + NOMS_EXPORTES.map(function (n) {
    return n + ': (typeof ' + n + " === 'function' ? " + n + ' : undefined)';
  }).join(', ') + '};')();
}

/** Source : mémoire rapide (6 h), sinon copie durable, sinon GitHub. */
function sourceIconic_() {
  const C = CacheService.getScriptCache(), P = PropertiesService.getScriptProperties();
  const n = Number(C.get('SRC_N') || 0);
  if (n) {
    const parts = C.getAll(Array.from({ length: n }, function (_, i) { return 'SRC_' + i; }));
    const s = Array.from({ length: n }, function (_, i) { return parts['SRC_' + i]; });
    if (s.every(function (x) { return x != null; })) return s.join('');
  }
  let src = lireMorceaux_(P, 'SRC');
  if (!src) { chargeurMaj_(true); src = lireMorceaux_(P, 'SRC'); }
  if (!src) throw new Error('Code ICONIC introuvable : vérifie ta connexion et exécute activerMisesAJour.');
  memoriser_(src);
  return src;
}

/** Va voir sur GitHub s'il y a plus récent ; installe si oui. */
function chargeurMaj_(forcer) {
  const P = PropertiesService.getScriptProperties();
  const lire = function (f) {
    const r = UrlFetchApp.fetch(DEPOT_SOURCE + f + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('GitHub injoignable (' + f + ' : ' + r.getResponseCode() + ')');
    return r.getContentText();
  };
  const info = JSON.parse(lire('version.json'));
  const actuelle = P.getProperty('SRC_VERSION') || '';
  if (!forcer && actuelle && !plusRecente_(info.version, actuelle)) return { installee: false, version: actuelle };
  if (actuelle && info.version === P.getProperty('MAJ_BLOQUEE')) return { installee: false, version: actuelle, bloquee: true };
  if (actuelle === info.version && lireMorceaux_(P, 'SRC')) return { installee: false, version: actuelle };
  const src = lire('Code.gs');
  if (src.indexOf("const VERSION = '" + info.version + "'") < 0) throw new Error('Fichier de mise à jour incomplet, réessaie plus tard.');
  const mod = compiler_(src);            // refuse un code qui ne se lance pas
  if (typeof mod.doPost !== 'function' || typeof mod.api !== 'function') throw new Error('Mise à jour invalide.');
  const ancien = lireMorceaux_(P, 'SRC');
  if (ancien) { ecrireMorceaux_(P, 'SRC_PREC', ancien); P.setProperty('SRC_PREC_VERSION', actuelle); }
  enregistrer_(src, info.version, false);
  P.setProperty('MAJ_DERNIERE', info.version); P.setProperty('MAJ_LE', new Date().toISOString()); P.deleteProperty('MAJ_ERREUR');
  if (actuelle) ScriptApp.newTrigger('apresMiseAJour').timeBased().after(60 * 1000).create();
  return { installee: true, version: info.version, notes: info.notes || '', avant: actuelle };
}

function enregistrer_(src, version) {
  const P = PropertiesService.getScriptProperties();
  ecrireMorceaux_(P, 'SRC', src);
  if (version) P.setProperty('SRC_VERSION', version);
  memoriser_(src);
  ICONIC__ = null;
}

function memoriser_(src) {
  const C = CacheService.getScriptCache(), T = 30000, o = {};   // < 100 Ko même avec accents et emojis
  const n = Math.ceil(src.length / T);
  for (let i = 0; i < n; i++) o['SRC_' + i] = src.slice(i * T, (i + 1) * T);
  o.SRC_N = String(n);
  C.putAll(o, 21600);
}

/* Les propriétés du script acceptent ~9 Ko par valeur : le code est découpé en morceaux. */
function ecrireMorceaux_(P, cle, txt) {
  const T = 2200, n = Math.ceil(txt.length / T), o = {};   // < 9 Ko même avec accents et emojis
  for (let i = 0; i < n; i++) o[cle + '_' + i] = txt.slice(i * T, (i + 1) * T);
  o[cle + '_N'] = String(n);
  const vieux = Number(P.getProperty(cle + '_N') || 0);
  P.setProperties(o, false);
  for (let i = n; i < vieux; i++) P.deleteProperty(cle + '_' + i);
}

function lireMorceaux_(P, cle) {
  const n = Number(P.getProperty(cle + '_N') || 0);
  if (!n) return '';
  const all = P.getProperties(), s = [];
  for (let i = 0; i < n; i++) { if (all[cle + '_' + i] == null) return ''; s.push(all[cle + '_' + i]); }
  return s.join('');
}

function plusRecente_(a, b) {
  const x = String(a).split('.').map(Number), y = String(b).split('.').map(Number);
  for (let i = 0; i < 3; i++) { if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) > (y[i] || 0); }
  return false;
}
