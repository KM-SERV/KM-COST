/**
 * KM Services — Coût de revient
 * Sauvegarde des paramètres dans un Google Sheet du Drive.
 *
 * À coller dans Extensions > Apps Script d'un classeur Google Sheets,
 * puis à déployer en application web (voir la marche à suivre en bas).
 */

// ⚠️ Changez cette clé et reportez la même dans l'application web.
var CLE_PARTAGEE = 'km-services-2026';

/* ============================================================
   POINTS D'ENTRÉE
   ============================================================ */

function doGet(e)  { return traiter(e); }
function doPost(e) { return traiter(e); }

function traiter(e) {
  try {
    var p = {};
    if (e && e.postData && e.postData.contents) {
      p = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      p = e.parameter;
    }

    if (String(p.cle || '') !== CLE_PARTAGEE) {
      return reponse({ ok: false, erreur: 'Clé partagée incorrecte' });
    }

    if (p.action === 'lire') {
      var d = lire();
      return reponse({ ok: true, data: d.data, date: d.date });
    }

    if (p.action === 'ecrire') {
      var date = ecrire(p.data, p.resume);
      return reponse({ ok: true, date: date });
    }

    return reponse({ ok: false, erreur: 'Action inconnue' });

  } catch (err) {
    return reponse({ ok: false, erreur: String(err && err.message ? err.message : err) });
  }
}

function reponse(objet) {
  return ContentService
    .createTextOutput(JSON.stringify(objet))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
   LECTURE / ÉCRITURE
   ============================================================ */

function lire() {
  var f = feuille('_donnees');
  var brut = f.getRange('A2').getValue();
  if (!brut) return { data: null, date: '' };
  return {
    data: JSON.parse(brut),
    date: formatDate(f.getRange('B2').getValue())
  };
}

function ecrire(data, resume) {
  var d = (typeof data === 'string') ? JSON.parse(data) : data;
  var r = (typeof resume === 'string') ? JSON.parse(resume) : (resume || {});
  var maintenant = new Date();

  // 1. Sauvegarde technique, relue par l'application
  var f = feuille('_donnees');
  f.getRange('A1').setValue('Sauvegarde automatique — ne pas modifier à la main');
  f.getRange('A2').setValue(JSON.stringify(d));
  f.getRange('B2').setValue(maintenant);
  f.getRange('B2').setNumberFormat('dd/MM/yyyy HH:mm');
  f.hideSheet();

  // 2. Version lisible des paramètres
  ecrireLisible(d, maintenant);

  // 3. Une ligne d'historique par enregistrement
  historiser(r, maintenant);

  return formatDate(maintenant);
}

/* ============================================================
   FEUILLE « Paramètres » — lisible par un humain
   ============================================================ */

function ecrireLisible(d, maintenant) {
  var f = feuille('Paramètres');
  f.clear();

  var JOURS = ['Lundi','Mardi','Mercredi','Jeudi','Vendredi','Samedi','Dimanche'];
  var lignes = [];

  lignes.push(['KM Services — Coût de revient', '']);
  lignes.push(['Dernière mise à jour', formatDate(maintenant)]);
  lignes.push(['', '']);

  lignes.push(['TEMPS DE TRAVAIL', '']);
  for (var i = 0; i < 7; i++) {
    lignes.push([JOURS[i], Number(d.jours[i]) || 0]);
  }
  lignes.push(['Semaines travaillées par an', Number(d.semaines) || 0]);
  lignes.push(['Jours fériés déduits', Number(d.feries) || 0]);
  lignes.push(['', '']);

  lignes.push(['RÉMUNÉRATION (mensuel)', '']);
  lignes.push(['Collaborateur', 'Loyer', 'Rémunération', 'Cotisations', 'Autres', 'Total']);
  (d.employes || []).forEach(function (e) {
    var t = n(e.loyer) + n(e.remuneration) + n(e.cotisations) + n(e.autres);
    lignes.push([e.nom, n(e.loyer), n(e.remuneration), n(e.cotisations), n(e.autres), t]);
  });
  lignes.push(['', '']);

  lignes.push(['CHARGES FIXES', '']);
  lignes.push(['Poste', 'Par mois', 'Par an']);
  var total = 0;
  (d.charges || []).forEach(function (c) {
    var an = n(c.montant);
    total += an;
    lignes.push([c.label, arrondi(an / 12, 2), an]);
  });
  lignes.push(['Total', arrondi(total / 12, 2), total]);
  lignes.push(['Réparties sur', Number(d.repart) || 1, 'personne(s)']);

  var largeur = 6;
  var grille = lignes.map(function (l) {
    var copie = l.slice();
    while (copie.length < largeur) copie.push('');
    return copie;
  });

  f.getRange(1, 1, grille.length, largeur).setValues(grille);
  f.getRange(1, 1, 1, largeur).setFontWeight('bold').setFontSize(12);
  f.setColumnWidth(1, 240);
  f.autoResizeColumns(2, largeur - 1);
}

/* ============================================================
   FEUILLE « Historique » — un enregistrement par ligne
   ============================================================ */

function historiser(r, maintenant) {
  var f = feuille('Historique');
  if (f.getLastRow() === 0) {
    f.appendRow([
      'Date', 'Collaborateur', 'Mode',
      'Coût horaire', 'Coût journalier', 'Coût hebdomadaire',
      'Coût mensuel', 'Coût annuel',
      'Heures / semaine', 'Heures facturables / an'
    ]);
    f.getRange(1, 1, 1, 10).setFontWeight('bold');
    f.setFrozenRows(1);
  }
  f.appendRow([
    maintenant, r.employe || '', r.mode || '',
    arrondi(r.horaire, 2), arrondi(r.jour, 2), arrondi(r.semaine, 2),
    arrondi(r.mois, 2), arrondi(r.an, 2),
    arrondi(r.hSem, 1), arrondi(r.hAn, 1)
  ]);
  f.getRange(f.getLastRow(), 1).setNumberFormat('dd/MM/yyyy HH:mm');
  f.getRange(f.getLastRow(), 4, 1, 5).setNumberFormat('# ##0.00 €');
}

/* ============================================================
   UTILITAIRES
   ============================================================ */

function feuille(nom) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var f = ss.getSheetByName(nom);
  if (!f) f = ss.insertSheet(nom);
  return f;
}

function n(v) {
  var x = parseFloat(v);
  return isNaN(x) ? 0 : x;
}

function arrondi(v, d) {
  var x = parseFloat(v);
  if (isNaN(x) || !isFinite(x)) return '';
  var p = Math.pow(10, d);
  return Math.round(x * p) / p;
}

function formatDate(v) {
  if (!v) return '';
  return Utilities.formatDate(new Date(v), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
}

/* ============================================================
   MARCHE À SUIVRE

   1. Créer un Google Sheets dans le Drive, le nommer par exemple
      « KM Services — Coût de revient ».
   2. Menu Extensions > Apps Script. Effacer le contenu, coller ce
      fichier, puis Enregistrer.
   3. Changer la valeur de CLE_PARTAGEE en haut du fichier.
   4. Bouton Déployer > Nouveau déploiement.
        Type          : Application web
        Exécuter en tant que : Moi
        Accès         : Tout le monde
   5. Autoriser l'accès quand Google le demande (passer par
      « Paramètres avancés » puis « Accéder au projet » si un
      avertissement s'affiche).
   6. Copier l'adresse qui se termine par /exec, la coller dans
      l'application, page Paramètres > Sauvegarde, avec la clé.

   Après chaque modification du script, refaire Déployer > Gérer les
   déploiements > modifier > Nouvelle version, sinon l'ancienne version
   reste en ligne.
   ============================================================ */
