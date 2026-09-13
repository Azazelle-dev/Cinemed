/**
 * Common.gs — Logique métier : lecture des mouvements confirmés dans ARRIVEES,
 * calcul de l'heure de pick-up, écriture non destructive dans les onglets journaliers.
 * Aucun déclencheur ici, uniquement des fonctions appelées par Main.gs.
 */

/**
 * Formate une date en clé de comparaison stable ("dd/MM/yyyy").
 * @param {Date} date
 * @return {string}
 */
function formatDateCle(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
}

/**
 * Fusionne une date et une heure stockées dans des cellules séparées en un seul objet Date complet.
 * @param {Date} date
 * @param {Date} heure
 * @return {Date}
 */
function combinerDateEtHeure(date, heure) {
  const resultat = new Date(date);
  resultat.setHours(heure.getHours());
  resultat.setMinutes(heure.getMinutes());
  resultat.setSeconds(0);
  resultat.setMilliseconds(0);
  return resultat;
}

/**
 * Calcule l'heure de pick-up pour un départ (soustrait le délai du trajet).
 * @param {string} trajet - clé valide de Global.DELAIS_TRAJET
 * @param {Date} heureDepart - date + heure complètes du vol/train
 * @return {Date}
 */
function calculerHeurePickupDepart(trajet, heureDepart) {
  const delaiMs = Global.DELAIS_TRAJET[trajet];
  return new Date(heureDepart.getTime() - delaiMs);
}

/**
 * Convertit un objet Date en texte "HH:mm" pur (ex: "15:05"). On écrit ce texte
 * directement dans la cellule, jamais l'objet Date complet, pour éviter tout
 * risque que Sheets affiche la date en plus de l'heure, quel que soit le format
 * appliqué à la cellule. Comme c'est du 24h avec zéros, un tri alphabétique
 * donne le même ordre qu'un tri chronologique (chaque onglet ne couvre qu'un jour).
 * @param {Date} date
 * @return {string}
 */
function formatHeureAffichage(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
}

/**
 * Lit la liste des dates disponibles depuis Paramètres (colonne D, à partir de la ligne 2) —
 * seule source de vérité pour savoir quelles dates sont valides.
 * @return {Date[]} dates triées chronologiquement
 */
function obtenirDatesDisponibles() {
  const feuille = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_PARAMETRES);
  const derniereLigne = feuille.getLastRow();
  if (derniereLigne < 2) return [];

  const valeurs = feuille
    .getRange(2, Global.PARAMETRES_COLONNE_DATES, derniereLigne - 1, 1)
    .getValues();

  return valeurs
    .map(ligne => ligne[0])
    .filter(valeur => valeur instanceof Date)
    .sort((a, b) => a - b);
}

/**
 * Convertit une date en nom d'onglet journalier attendu, ex: "VENDREDI 17".
 * @param {Date} date
 * @return {string}
 */
function nomOngletPourDate(date) {
  const jour = Global.JOURS_FR[date.getDay()];
  return jour + " " + date.getDate();
}

/**
 * Parcourt toute la feuille ARRIVEES et retourne les mouvements confirmés
 * (arrivée et/ou départ) dont la date fait partie de datesAutorisees.
 * @param {Set<string>} datesAutorisees - clés formatDateCle des dates valides
 * @return {Array<Object>} mouvements { date, type, nom, prenom, telephone, trajet, heurePickup }
 */
function collecterMouvementsConfirmes(datesAutorisees) {
  const feuille = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_ARRIVEES);
  const donnees = feuille.getDataRange().getValues();
  const enTetes = donnees[0];

  const idx = {};
  Object.keys(Global.COLONNES_MASTER).forEach(cle => {
    idx[cle] = enTetes.indexOf(Global.COLONNES_MASTER[cle]);
  });

  const lignes = donnees.slice(1);
  const mouvements = [];

  lignes.forEach(ligne => {
    if (!ligne || !ligne[idx.NOM]) return; // ligne vide, on ignore

    const nom = ligne[idx.NOM];
    const prenom = ligne[idx.PRENOM];
    const telephone = ligne[idx.TELEPHONE];

    // --- Côté arrivée ---
    const modeArrivee = ligne[idx.MODE_ARRIVEE];
    const dateArrivee = ligne[idx.DATE_ARRIVEE];
    if (Global.DELAIS_TRAJET.hasOwnProperty(modeArrivee) && dateArrivee instanceof Date) {
      const cle = formatDateCle(dateArrivee);
      if (datesAutorisees.has(cle)) {
        mouvements.push({
          date: dateArrivee,
          type: "Arrivée",
          nom: nom,
          prenom: prenom,
          telephone: telephone,
          trajet: modeArrivee,
          // Pas de délai à soustraire pour une arrivée : le chauffeur récupère au moment de l'arrivée
          // On stocke directement le texte "HH:mm", jamais l'objet Date complet
          heurePickup: formatHeureAffichage(combinerDateEtHeure(dateArrivee, ligne[idx.HEURE_ARRIVEE]))
        });
      }
    }

    // --- Côté départ ---
    const modeDepart = ligne[idx.MODE_DEPART];
    const dateDepart = ligne[idx.DATE_DEPART];
    if (Global.DELAIS_TRAJET.hasOwnProperty(modeDepart) && dateDepart instanceof Date) {
      const cle = formatDateCle(dateDepart);
      if (datesAutorisees.has(cle)) {
        const heureDepartComplete = combinerDateEtHeure(dateDepart, ligne[idx.HEURE_DEPART]);
        mouvements.push({
          date: dateDepart,
          type: "Départ",
          nom: nom,
          prenom: prenom,
          telephone: telephone,
          trajet: modeDepart,
          heurePickup: formatHeureAffichage(calculerHeurePickupDepart(modeDepart, heureDepartComplete))
        });
      }
    }
  });

  return mouvements;
}

/**
 * Crée l'onglet avec juste les en-têtes s'il n'existe pas encore, et lui applique
 * tout de suite la mise en forme (voir formaterOnglet). Ne fait rien s'il existe
 * déjà (jamais d'écrasement).
 * @param {string} nomOnglet
 * @return {Sheet}
 */
function assurerOngletExiste(nomOnglet) {
  let feuille = SpreadsheetApp.getActive().getSheetByName(nomOnglet);
  if (!feuille) {
    feuille = SpreadsheetApp.getActive().insertSheet(nomOnglet);
    feuille.getRange(1, 1, 1, Global.ENTETES_PLANNING_JOUR.length).setValues([Global.ENTETES_PLANNING_JOUR]);
    formaterOnglet(feuille);
  }
  return feuille;
}

/**
 * Mise en forme "tableau propre" d'un onglet journalier : en-tête en gras avec
 * fond coloré, lignes de données centrées avec bordures fines, bandes de couleur
 * alternées, colonnes ajustées à la largeur du contenu, première ligne figée.
 * Rejouable à tout moment sans dégrader le rendu.
 * @param {Sheet} feuille
 */
function formaterOnglet(feuille) {
  const derniereLigne = feuille.getLastRow();
  const derniereColonne = feuille.getLastColumn();
  if (derniereLigne < 1 || derniereColonne < 1) return;

  feuille.getRange(1, 1, 1, derniereColonne)
    .setFontWeight("bold")
    .setBackground("#1c4587")
    .setFontColor("#ffffff")
    .setHorizontalAlignment("center");

  feuille.setFrozenRows(1);

  if (derniereLigne > 1) {
    const donnees = feuille.getRange(2, 1, derniereLigne - 1, derniereColonne);
    donnees.setHorizontalAlignment("center");
    donnees.setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);

    // Bandes alternées : on retire les anciennes avant de réappliquer sur la plage à jour
    feuille.getBandings().forEach(bande => bande.remove());
    feuille.getRange(1, 1, derniereLigne, derniereColonne)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  }

  feuille.autoResizeColumns(1, derniereColonne);
}

/**
 * Écrit une liste de mouvements dans l'onglet journalier correspondant.
 * Utilise assurerOngletExiste pour ne jamais planter si l'onglet n'a pas encore
 * été créé. Ajoute les mouvements absents, met à jour Heure de pick up / Trajet
 * pour ceux déjà présents, ne touche jamais à Chauffeur (assignation manuelle),
 * puis trie par ordre chronologique.
 * @param {string} nomOnglet
 * @param {Array<Object>} mouvements
 */
function ecrireMouvementsDansOnglet(nomOnglet, mouvements) {
  const feuille = assurerOngletExiste(nomOnglet);

  const enTetes = feuille.getRange(1, 1, 1, feuille.getLastColumn()).getValues()[0];
  const idxNom          = enTetes.indexOf("Nom") + 1;
  const idxPrenom       = enTetes.indexOf("Prénom") + 1;
  const idxType         = enTetes.indexOf("Arrivée ou départ") + 1;
  const idxHeurePickup  = enTetes.indexOf("Heure de pick up") + 1;
  const idxTrajet       = enTetes.indexOf("Trajet") + 1;

  // Sans ça, Sheets réinterprète une chaîne comme "15:05" écrite via setValue comme
  // une heure/date (même comportement qu'une saisie manuelle sur une cellule au format
  // "Automatique"). Le format "@" (texte brut) empêche cette reconversion.
  if (idxHeurePickup > 0) {
    feuille.getRange(1, idxHeurePickup, feuille.getMaxRows(), 1).setNumberFormat("@");
  }

  // Repérer les lignes déjà présentes (clé Nom|Prénom|Type → numéro de ligne)
  const derniereLigneAvant = feuille.getLastRow();
  const dejaPresents = new Map();
  if (derniereLigneAvant > 1) {
    const existants = feuille.getRange(2, 1, derniereLigneAvant - 1, feuille.getLastColumn()).getValues();
    existants.forEach((ligne, i) => {
      const cle = ligne[idxNom - 1] + "|" + ligne[idxPrenom - 1] + "|" + ligne[idxType - 1];
      dejaPresents.set(cle, i + 2); // +2 : décalage en-tête + index 0-based
    });
  }

  const nouvellesLignes = [];
  let quelqueChoseAChange = false;

  mouvements.forEach(mvt => {
    const cle = mvt.nom + "|" + mvt.prenom + "|" + mvt.type;

    if (dejaPresents.has(cle)) {
      // Déjà listé : on rafraîchit juste l'heure et le trajet si besoin, sans toucher au reste
      const ligneExistante = dejaPresents.get(cle);
      const heureActuelle = feuille.getRange(ligneExistante, idxHeurePickup).getValue();
      const trajetActuel  = feuille.getRange(ligneExistante, idxTrajet).getValue();

      if (heureActuelle !== mvt.heurePickup || trajetActuel !== mvt.trajet) {
        feuille.getRange(ligneExistante, idxHeurePickup).setValue(mvt.heurePickup);
        feuille.getRange(ligneExistante, idxTrajet).setValue(mvt.trajet);
        quelqueChoseAChange = true;
      }
    } else {
      nouvellesLignes.push([
        "",              // Chauffeur — assignation manuelle
        mvt.nom,
        mvt.prenom,
        mvt.heurePickup,
        mvt.type,
        mvt.trajet,
        mvt.telephone
      ]);
    }
  });

  if (nouvellesLignes.length > 0) {
    feuille.getRange(feuille.getLastRow() + 1, 1, nouvellesLignes.length, nouvellesLignes[0].length)
      .setValues(nouvellesLignes);
    quelqueChoseAChange = true;
  }

  // On ne trie/reformate que s'il y a eu un changement réel : évite des appels API inutiles
  // quand une génération est relancée sans qu'aucune donnée n'ait bougé.
  if (!quelqueChoseAChange) return;

  const derniereLigneApres = feuille.getLastRow();
  if (derniereLigneApres > 1) {
    feuille.getRange(2, 1, derniereLigneApres - 1, feuille.getLastColumn())
      .sort({ column: idxHeurePickup, ascending: true });
  }

  formaterOnglet(feuille); // réapplique la mise en forme (bordures, bandes, largeur des colonnes) sur la plage à jour
}

/**
 * Orchestration, en deux temps :
 * 1. Crée un onglet pour chaque date demandée, même sans aucun mouvement confirmé
 *    (répond à "un onglet pour chaque date de Paramètres, même vide").
 * 2. Rassemble les mouvements confirmés et les écrit dans les onglets concernés.
 * @param {Date[]} [datesCiblees]
 */
function genererPlannings(datesCiblees) {
  const dates = (datesCiblees && datesCiblees.length) ? datesCiblees : obtenirDatesDisponibles();

  dates.forEach(date => assurerOngletExiste(nomOngletPourDate(date)));

  const clesAutorisees = new Set(dates.map(formatDateCle));
  const mouvements = collecterMouvementsConfirmes(clesAutorisees);

  const parOnglet = {};
  mouvements.forEach(mvt => {
    const nomOnglet = nomOngletPourDate(mvt.date);
    if (!parOnglet[nomOnglet]) parOnglet[nomOnglet] = [];
    parOnglet[nomOnglet].push(mvt);
  });

  Object.keys(parOnglet).forEach(nomOnglet => ecrireMouvementsDansOnglet(nomOnglet, parOnglet[nomOnglet]));
}
