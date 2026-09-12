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
 * Parcourt ARRIVEES et retourne les mouvements confirmés (arrivée et/ou départ)
 * dont la date fait partie de datesAutorisees.
 * @param {Set<string>} datesAutorisees - clés formatDateCle des dates valides
 * @param {number} [ligneCible] - si fourni, ne traite que cette ligne (1-based, utile pour onEdit)
 * @return {Array<Object>} mouvements { date, type, nom, prenom, telephone, trajet, heurePickup }
 */
function collecterMouvementsConfirmes(datesAutorisees, ligneCible) {
  const feuille = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_ARRIVEES);
  const donnees = feuille.getDataRange().getValues();
  const enTetes = donnees[0];

  const idx = {};
  Object.keys(Global.COLONNES_MASTER).forEach(cle => {
    idx[cle] = enTetes.indexOf(Global.COLONNES_MASTER[cle]);
  });

  const lignes = ligneCible ? [donnees[ligneCible - 1]] : donnees.slice(1);
  const mouvements = [];

  lignes.forEach(ligne => {
    if (!ligne || !ligne[idx.NOM]) return; // ligne vide ou hors bornes, on ignore

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
          heurePickup: combinerDateEtHeure(dateArrivee, ligne[idx.HEURE_ARRIVEE])
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
          heurePickup: calculerHeurePickupDepart(modeDepart, heureDepartComplete)
        });
      }
    }
  });

  return mouvements;
}

/**
 * Écrit une liste de mouvements dans l'onglet journalier correspondant.
 * Crée l'onglet s'il n'existe pas. Ajoute les mouvements absents, met à jour
 * Heure de pick up / Trajet pour ceux déjà présents, ne touche jamais à Chauffeur
 * (assignation manuelle), puis trie par ordre chronologique.
 * @param {string} nomOnglet
 * @param {Array<Object>} mouvements
 */
function ecrireMouvementsDansOnglet(nomOnglet, mouvements) {
  let feuille = SpreadsheetApp.getActive().getSheetByName(nomOnglet);
  if (!feuille) {
    feuille = SpreadsheetApp.getActive().insertSheet(nomOnglet);
    feuille.getRange(1, 1, 1, Global.ENTETES_PLANNING_JOUR.length).setValues([Global.ENTETES_PLANNING_JOUR]);
  }

  const enTetes = feuille.getRange(1, 1, 1, feuille.getLastColumn()).getValues()[0];
  const idxNom          = enTetes.indexOf("Nom") + 1;
  const idxPrenom       = enTetes.indexOf("Prénom") + 1;
  const idxType         = enTetes.indexOf("Arrivée ou départ") + 1;
  const idxHeurePickup  = enTetes.indexOf("Heure de pick up") + 1;
  const idxTrajet       = enTetes.indexOf("Trajet") + 1;

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

  mouvements.forEach(mvt => {
    const cle = mvt.nom + "|" + mvt.prenom + "|" + mvt.type;

    if (dejaPresents.has(cle)) {
      // Déjà listé : on rafraîchit juste l'heure et le trajet, sans toucher au reste
      const ligneExistante = dejaPresents.get(cle);
      feuille.getRange(ligneExistante, idxHeurePickup).setValue(mvt.heurePickup);
      feuille.getRange(ligneExistante, idxTrajet).setValue(mvt.trajet);
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
  }

  const derniereLigneApres = feuille.getLastRow();
  if (derniereLigneApres > 1) {
    feuille.getRange(2, 1, derniereLigneApres - 1, feuille.getLastColumn())
      .sort({ column: idxHeurePickup, ascending: true });
  }
}

/**
 * Orchestration : rassemble les mouvements confirmés (pour les dates données, ou
 * toutes celles de Paramètres si non précisé), les regroupe par onglet cible, puis écrit.
 * @param {Date[]} [datesCiblees]
 */
function genererPlannings(datesCiblees) {
  const dates = (datesCiblees && datesCiblees.length) ? datesCiblees : obtenirDatesDisponibles();
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
