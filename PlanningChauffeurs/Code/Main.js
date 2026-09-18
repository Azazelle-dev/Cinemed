/**
 * Main.gs — Point d'entrée du script.
 * Menu (onOpen) et actions du menu. La génération se fait à la demande via le
 * menu, ou automatiquement si l'utilisateur a activé la mise à jour auto
 * (déclencheur installable onEdit + déclencheur ponctuel à retardement, voir
 * plus bas).
 */

/**
 * Ajoute le menu personnalisé "Planning Chauffeurs" à l'ouverture du Sheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Planning Chauffeurs")
    .addItem("Générer tous les plannings", "menuGenererTout")
    .addItem("Générer des dates spécifiques…", "menuOuvrirSelectionDates")
    .addSeparator()
    .addItem("Activer la mise à jour automatique", "menuActiverMiseAJourAuto")
    .addItem("Désactiver la mise à jour automatique", "menuDesactiverMiseAJourAuto")
    .addSeparator()
    .addItem("Supprimer tous les plannings générés", "menuSupprimerPlannings")
    .addToUi();
}

/**
 * Action de menu : génère/met à jour les plannings pour toutes les dates déduites
 * d'ARRIVEES/DEPARTS. L'onglet actif de départ est déjà restauré par genererPlannings.
 */
function menuGenererTout() {
  genererPlannings(); // pas d'argument = toutes les dates déduites
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
 * L'onglet actif de départ est déjà restauré par genererPlannings.
 * @param {string[]} datesTexte - dates au format "dd/MM/yyyy"
 */
function genererOngletsSelectionnes(datesTexte) {
  const toutesLesDates = obtenirDatesDisponibles();
  const datesValides = toutesLesDates.filter(d => datesTexte.includes(formatDateCle(d)));
  genererPlannings(datesValides);
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

/**
 * Déclencheur installable sur modification (posé par menuActiverMiseAJourAuto) :
 * reprogramme une génération automatique Global.DELAI_INACTIVITE_MS après la
 * dernière modification sur ARRIVEES ou DEPARTS. Ignore les modifications sur
 * les autres onglets (plannings générés, Paramètres, Planning Source) pour ne
 * pas se redéclencher sur un ajustement manuel de planning (assignation d'un
 * chauffeur, etc.).
 * @param {Object} e - événement d'édition transmis par Apps Script
 */
function onEditInstallable(e) {
  const nomOnglet = e && e.range ? e.range.getSheet().getName() : "";
  if (nomOnglet !== Global.ONGLET_ARRIVEES && nomOnglet !== Global.ONGLET_DEPARTS) return;

  annulerGenerationAutoProgrammee();

  const trigger = ScriptApp.newTrigger(Global.NOM_FONCTION_AUTO)
    .timeBased()
    .after(Global.DELAI_INACTIVITE_MS)
    .create();

  PropertiesService.getScriptProperties().setProperty(Global.PROPRIETE_TRIGGER_AUTO, trigger.getUniqueId());
}

/**
 * Supprime le déclencheur ponctuel de génération automatique programmé par
 * l'édition précédente, s'il n'a pas encore eu le temps de s'exécuter — c'est
 * ce qui fait repartir le compte des 15 secondes à chaque nouvelle modification.
 */
function annulerGenerationAutoProgrammee() {
  const proprietes = PropertiesService.getScriptProperties();
  const idPrecedent = proprietes.getProperty(Global.PROPRIETE_TRIGGER_AUTO);
  if (!idPrecedent) return;

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getUniqueId() === idPrecedent) ScriptApp.deleteTrigger(trigger);
  });
  proprietes.deleteProperty(Global.PROPRIETE_TRIGGER_AUTO);
}

/**
 * Exécutée par le déclencheur ponctuel une fois le délai d'inactivité écoulé :
 * génère les plannings puis nettoie la référence au déclencheur (déjà
 * supprimé de lui-même par Apps Script après exécution, puisque ponctuel).
 */
function genererPlanningsAuto() {
  PropertiesService.getScriptProperties().deleteProperty(Global.PROPRIETE_TRIGGER_AUTO);
  genererPlannings();
}

/**
 * Action de menu : active la mise à jour automatique. Pose le déclencheur
 * installable onEdit qui programmera lui-même les générations à retardement.
 */
function menuActiverMiseAJourAuto() {
  desactiverMiseAJourAuto(); // évite les doublons si déjà actif

  ScriptApp.newTrigger(Global.NOM_FONCTION_ON_EDIT)
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert(
    "Mise à jour automatique activée : les plannings se régénèrent 15 secondes " +
    "après la dernière modification sur ARRIVEES ou DEPARTS."
  );
}

/**
 * Action de menu : désactive la mise à jour automatique (déclencheur onEdit
 * et toute génération ponctuelle en attente).
 */
function menuDesactiverMiseAJourAuto() {
  desactiverMiseAJourAuto();
  SpreadsheetApp.getUi().alert("Mise à jour automatique désactivée.");
}

/**
 * Supprime le déclencheur installable onEdit et toute génération ponctuelle
 * déjà programmée. Sans effet si la mise à jour automatique n'était pas active.
 */
function desactiverMiseAJourAuto() {
  annulerGenerationAutoProgrammee();
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === Global.NOM_FONCTION_ON_EDIT) ScriptApp.deleteTrigger(trigger);
  });
}
