/**
 * Main.gs — Point d'entrée du script.
 * Menu (onOpen), actions du menu, et onEdit qui réagit aux modifications
 * de ModeArrivée/ModeDépart dans ARRIVEES.
 */

/**
 * Ajoute le menu personnalisé "Planning Chauffeurs" à l'ouverture du Sheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Planning Chauffeurs")
    .addItem("Générer tous les plannings", "menuGenererTout")
    .addItem("Générer des dates spécifiques…", "menuOuvrirSelectionDates")
    .addToUi();
}

/**
 * Action de menu : génère/met à jour les plannings pour toutes les dates de Paramètres.
 */
function menuGenererTout() {
  genererPlannings(); // pas d'argument = toutes les dates de Paramètres
  SpreadsheetApp.getUi().alert("Tous les plannings ont été générés/mis à jour.");
}

/**
 * Action de menu : ouvre la fenêtre de sélection de dates.
 */
function menuOuvrirSelectionDates() {
  const html = HtmlService.createHtmlOutputFromFile("SelectionDates").setWidth(300).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Choisir les dates à générer");
}

/**
 * Appelée depuis SelectionDates.html pour récupérer la liste des dates à afficher.
 * @return {string[]} dates formatées "dd/MM/yyyy"
 */
function getDatesPourAffichage() {
  return obtenirDatesDisponibles().map(formatDateCle);
}

/**
 * Appelée depuis SelectionDates.html quand l'utilisateur valide sa sélection.
 * Ignore toute valeur qui ne fait pas partie des dates officielles de Paramètres.
 * @param {string[]} datesTexte - dates au format "dd/MM/yyyy"
 */
function genererOngletsSelectionnes(datesTexte) {
  const toutesLesDates = obtenirDatesDisponibles();
  const datesValides = toutesLesDates.filter(d => datesTexte.includes(formatDateCle(d)));
  genererPlannings(datesValides);
}

/**
 * Déclencheur simple, exécuté automatiquement à chaque modification de cellule.
 * Ne réagit que sur l'onglet ARRIVEES, colonnes ModeArrivée/ModeDépart : dès qu'un
 * trajet y est choisi ou changé, la ligne concernée est immédiatement (re)traitée
 * et placée dans le bon onglet journalier.
 * @param {Object} e - objet événement fourni par Apps Script
 */
function onEdit(e) {
  const feuille = e.range.getSheet();
  if (feuille.getName() !== Global.ONGLET_ARRIVEES) return;

  const ligne = e.range.getRow();
  if (ligne === 1) return; // en-tête

  const enTetes = feuille.getRange(1, 1, 1, feuille.getLastColumn()).getValues()[0];
  const idxModeArrivee = enTetes.indexOf(Global.COLONNES_MASTER.MODE_ARRIVEE) + 1;
  const idxModeDepart  = enTetes.indexOf(Global.COLONNES_MASTER.MODE_DEPART) + 1;
  const colonne = e.range.getColumn();

  if (colonne !== idxModeArrivee && colonne !== idxModeDepart) return; // pas la bonne colonne, rien à faire

  try {
    const datesAutorisees = new Set(obtenirDatesDisponibles().map(formatDateCle));
    const mouvements = collecterMouvementsConfirmes(datesAutorisees, ligne);

    const parOnglet = {};
    mouvements.forEach(mvt => {
      const nomOnglet = nomOngletPourDate(mvt.date);
      if (!parOnglet[nomOnglet]) parOnglet[nomOnglet] = [];
      parOnglet[nomOnglet].push(mvt);
    });

    Object.keys(parOnglet).forEach(nomOnglet => ecrireMouvementsDansOnglet(nomOnglet, parOnglet[nomOnglet]));

  } catch (erreur) {
    console.warn("onEdit ARRIVEES, ligne " + ligne + " ignorée : " + erreur.message);
  }
}
