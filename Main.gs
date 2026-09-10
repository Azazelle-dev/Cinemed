/**
 * Main.gs — Point d'entrée du script.
 * Déclencheurs automatiques (onEdit, onOpen) et actions du menu personnalisé.
 */

/**
 * Déclencheur simple, exécuté automatiquement par Google à chaque modification
 * de cellule. Ne recalcule que les lignes des onglets journaliers dont une
 * colonne surveillée (Trajet, DateDépart, HeureDépart) a été modifiée.
 * Gère aussi bien l'édition d'une seule cellule que le collage d'une plage
 * de plusieurs lignes/colonnes d'un coup.
 * @param {Object} e - objet événement fourni par Apps Script
 */
function onEdit(e) {
  const feuille = e.range.getSheet();

  // On ignore tout onglet qui n'est pas un onglet journalier (ex: "Paramètres")
  if (!Global.MOTIF_NOM_ONGLET_JOURNALIER.test(feuille.getName())) {
    return;
  }

  const idx = obtenirIndexColonnes(feuille);
  if (!idx.trajet || !idx.dateDepart || !idx.heureDepart || !idx.heurePickup) {
    return; // structure de colonnes inattendue sur cet onglet
  }

  const colonnesSurveillees = [idx.trajet, idx.dateDepart, idx.heureDepart];

  const colonneDebut = e.range.getColumn();
  const colonneFin = colonneDebut + e.range.getNumColumns() - 1;
  const uneColonneSurveilleeToucheé = colonnesSurveillees.some(
    col => col >= colonneDebut && col <= colonneFin
  );
  if (!uneColonneSurveilleeToucheé) {
    return; // rien à faire
  }

  const ligneDebut = e.range.getRow();
  const nombreLignes = e.range.getNumRows();

  for (let i = 0; i < nombreLignes; i++) {
    const ligne = ligneDebut + i;
    if (ligne === 1) continue; // on ignore la ligne d'en-têtes
    recalculerLignePickup(feuille, ligne, idx);
  }
}

/**
 * Ajoute le menu personnalisé "Planning Chauffeurs" à l'ouverture du Sheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Planning Chauffeurs")
    .addItem("Mettre à jour tous les plannings", "menuMettreAJourTout")
    .addItem("Mettre à jour des dates spécifiques…", "menuOuvrirSelectionDates")
    .addToUi();
}

/**
 * Action de menu : recalcule Heure Pick up sur tous les onglets journaliers
 * listés dans Paramètres.
 */
function menuMettreAJourTout() {
  const dates = obtenirDatesDisponibles();
  dates.forEach(date => mettreAJourOnglet(nomOngletPourDate(date)));

  SpreadsheetApp.getUi().alert("Tous les plannings ont été mis à jour.");
}

/**
 * Action de menu : ouvre la fenêtre de sélection de dates.
 */
function menuOuvrirSelectionDates() {
  const html = HtmlService.createHtmlOutputFromFile("SelectionDates")
    .setWidth(300)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Choisir les dates à mettre à jour");
}

/**
 * Appelée depuis SelectionDates.html pour récupérer la liste des dates à afficher.
 * @return {string[]} dates formatées "dd/MM/yyyy"
 */
function getDatesPourAffichage() {
  return obtenirDatesDisponibles().map(date =>
    Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy")
  );
}

/**
 * Appelée depuis SelectionDates.html quand l'utilisateur valide sa sélection.
 * Ignore toute valeur qui ne fait pas partie des dates officielles de Paramètres.
 * @param {string[]} datesSelectionneesTexte - dates au format "dd/MM/yyyy"
 */
function mettreAJourOngletsSelectionnes(datesSelectionneesTexte) {
  const fuseau = Session.getScriptTimeZone();
  const datesDisponibles = obtenirDatesDisponibles();
  const datesAutorisees = datesDisponibles.map(date =>
    Utilities.formatDate(date, fuseau, "dd/MM/yyyy")
  );

  datesSelectionneesTexte
    .filter(texte => datesAutorisees.includes(texte))
    .forEach(texte => {
      const position = datesAutorisees.indexOf(texte);
      const date = datesDisponibles[position];
      mettreAJourOnglet(nomOngletPourDate(date));
    });
}
