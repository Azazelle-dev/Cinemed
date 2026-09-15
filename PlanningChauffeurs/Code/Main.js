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
    .addItem("Supprimer tous les plannings générés", "menuSupprimerPlannings")
    .addToUi();
}

/**
 * Action de menu : génère/met à jour les plannings pour toutes les dates déduites
 * d'ARRIVEES/DEPARTS. Restaure l'onglet actif de départ une fois la génération terminée.
 */
function menuGenererTout() {
  const ongletDeDepart = SpreadsheetApp.getActive().getActiveSheet();

  genererPlannings(); // pas d'argument = toutes les dates déduites

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
 * Ignore toute valeur qui ne fait pas partie des dates déduites d'ARRIVEES/DEPARTS.
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

/**
 * Action de menu : supprime tous les onglets de planning générés, après confirmation.
 * Action irréversible — ne touche jamais à ARRIVEES, DEPARTS, Paramètres ou Planning Source.
 */
function menuSupprimerPlannings() {
  const ui = SpreadsheetApp.getUi();
  const reponse = ui.alert(
    "Supprimer tous les plannings",
    "Ça va supprimer définitivement tous les onglets de planning générés (nommés \"JOUR NUMÉRO\"). Cette action est irréversible. Continuer ?",
    ui.ButtonSet.YES_NO
  );

  if (reponse !== ui.Button.YES) return;

  supprimerPlanningsGeneres();
  ui.alert("Tous les plannings générés ont été supprimés.");
}
