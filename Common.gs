/**
 * Common.gs — Fonctions utilitaires réutilisables.
 * Aucun déclencheur ici : uniquement des fonctions pures ou d'aide,
 * appelées par Main.gs.
 */

/**
 * Calcule l'heure de pick-up à partir du trajet et de l'heure de départ.
 * @param {string} trajet - valeur exacte du menu déroulant (ex: "Corum>MRS")
 * @param {Date} heureDepart - date + heure complètes du vol/train
 * @return {Date} date + heure du pick-up
 */
function calculerHeurePickup(trajet, heureDepart) {
  if (!trajet) {
    throw new Error("Aucun trajet sélectionné pour cette ligne.");
  }

  const delaiMs = Global.DELAIS_TRAJET[trajet];
  if (delaiMs === undefined) {
    throw new Error("Délai non défini pour le trajet : " + trajet);
  }

  return new Date(heureDepart.getTime() - delaiMs);
}

/**
 * Fusionne une date et une heure (souvent stockées séparément dans le sheet)
 * en un seul objet Date complet.
 * @param {Date} date - date pure (ex: DateDépart)
 * @param {Date} heure - heure pure, généralement avec une date par défaut Apps Script
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
 * Lit les en-têtes de la ligne 1 d'une feuille et retourne les index de colonnes
 * (1-based) nécessaires au calcul. Retourne 0 pour une colonne introuvable.
 * @param {Sheet} feuille
 * @return {{trajet:number, dateDepart:number, heureDepart:number, heurePickup:number}}
 */
function obtenirIndexColonnes(feuille) {
  const enTetes = feuille.getRange(1, 1, 1, feuille.getLastColumn()).getValues()[0];
  return {
    trajet:       enTetes.indexOf(Global.COLONNES.TRAJET) + 1,
    dateDepart:   enTetes.indexOf(Global.COLONNES.DATE_DEPART) + 1,
    heureDepart:  enTetes.indexOf(Global.COLONNES.HEURE_DEPART) + 1,
    heurePickup:  enTetes.indexOf(Global.COLONNES.HEURE_PICKUP) + 1
  };
}

/**
 * Recalcule et écrit l'heure de pick-up pour une seule ligne d'un onglet journalier.
 * Ne fait rien (silencieusement) si les données sont incomplètes.
 * @param {Sheet} feuille
 * @param {number} ligne - numéro de ligne (1-based)
 * @param {{trajet:number, dateDepart:number, heureDepart:number, heurePickup:number}} idx
 */
function recalculerLignePickup(feuille, ligne, idx) {
  const trajet       = feuille.getRange(ligne, idx.trajet).getValue();
  const dateDepart    = feuille.getRange(ligne, idx.dateDepart).getValue();
  const heureDepart   = feuille.getRange(ligne, idx.heureDepart).getValue();

  if (!trajet || !dateDepart || !heureDepart) {
    return; // données incomplètes, on n'écrit rien
  }

  try {
    const heureDepartComplete = combinerDateEtHeure(dateDepart, heureDepart);
    const heurePickup = calculerHeurePickup(trajet, heureDepartComplete);
    feuille.getRange(ligne, idx.heurePickup).setValue(heurePickup);
  } catch (erreur) {
    console.warn("Onglet " + feuille.getName() + ", ligne " + ligne + " ignorée : " + erreur.message);
  }
}

/**
 * Lit la liste des dates disponibles depuis Paramètres!D2:D (source de vérité unique).
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
  const jour = JOURS_FR[date.getDay()];
  const numero = date.getDate();
  return jour + " " + numero;
}

/**
 * Met à jour en place la colonne Heure Pick up de toutes les lignes valides
 * d'un onglet journalier. N'écrase rien d'autre. Ignore silencieusement
 * (avec log) si l'onglet ou les colonnes attendues sont introuvables.
 * @param {string} nomOnglet
 */
function mettreAJourOnglet(nomOnglet) {
  const feuille = SpreadsheetApp.getActive().getSheetByName(nomOnglet);
  if (!feuille) {
    // Onglet pas encore créé pour cette date. Comportement actuel : on ignore.
    // Évolution future prévue : créer l'onglet automatiquement à partir d'un modèle
    // plutôt que de l'ignorer (cf. section 6 de la spec, "Évolution prévue").
    console.warn("Onglet introuvable, ignoré : " + nomOnglet);
    return;
  }

  const idx = obtenirIndexColonnes(feuille);
  if (!idx.trajet || !idx.dateDepart || !idx.heureDepart || !idx.heurePickup) {
    console.warn("Colonnes manquantes sur l'onglet : " + nomOnglet);
    return;
  }

  const derniereLigne = feuille.getLastRow();
  for (let ligne = 2; ligne <= derniereLigne; ligne++) {
    recalculerLignePickup(feuille, ligne, idx);
  }
}
