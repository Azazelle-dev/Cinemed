/**
 * Main.gs — Point d'entrée du script.
 * Menu (onOpen) et actions du menu — aucun déclencheur automatique : la
 * génération se fait uniquement à la demande, via le menu.
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
 * Restaure l'onglet actif de départ une fois la génération terminée.
 */
function menuGenererTout() {
  const ongletDeDepart = SpreadsheetApp.getActive().getActiveSheet();

  genererPlannings(); // pas d'argument = toutes les dates de Paramètres

  SpreadsheetApp.getActive().setActiveSheet(ongletDeDepart);
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
 * Restaure l'onglet actif de départ une fois la génération terminée.
 * @param {string[]} datesTexte - dates au format "dd/MM/yyyy"
 */
function genererOngletsSelectionnes(datesTexte) {
  const ongletDeDepart = SpreadsheetApp.getActive().getActiveSheet();

  const toutesLesDates = obtenirDatesDisponibles();
  const datesValides = toutesLesDates.filter(d => datesTexte.includes(formatDateCle(d)));
  genererPlannings(datesValides);

  SpreadsheetApp.getActive().setActiveSheet(ongletDeDepart);
}
