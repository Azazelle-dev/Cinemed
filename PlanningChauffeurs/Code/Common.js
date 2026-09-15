/**
 * Common.gs — Logique métier : tables de délais, dates, collecte des mouvements,
 * écriture non destructive dans les onglets journaliers, mise en forme (plannings
 * et sources), suppression. Aucun déclencheur ici, uniquement des fonctions
 * appelées par Main.gs.
 */

/**
 * Inverse un identifiant de trajet "A>B" en "B>A", pour retrouver l'entrée
 * correspondante dans l'autre table de délais sans maintenir une correspondance séparée.
 * @param {string} trajet
 * @return {string|null} null si le format n'est pas celui attendu
 */
function nomTrajetInverse(trajet) {
  const parties = trajet.split(">");
  if (parties.length !== 2) return null;
  return parties[1] + ">" + parties[0];
}

/**
 * Extrait le nom de la gare/aéroport d'un trajet, quel que soit son sens.
 * ex: "SDF>Corum" → "SDF", "Corum>SDF" → "SDF"
 * @param {string} trajet
 * @return {string}
 */
function extraireLieu(trajet) {
  if (!trajet) return "";
  const parties = trajet.split(">");
  if (parties.length !== 2) return trajet;
  return parties[0] === "Corum" ? parties[1] : parties[0];
}

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
  return resultat;
}

/**
 * Convertit un objet Date en texte "HH:mm" pur (ex: "15:05").
 * @param {Date} date
 * @return {string}
 */
function formatHeureAffichage(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
}

/**
 * Convertit une cellule de durée (ex: 01:00:00, lue comme objet Date par Apps Script)
 * en millisecondes. Utilise les getters UTC volontairement : les cellules de durée
 * pure n'ont pas de fuseau horaire, donc passer par getHours() (heure locale du
 * script) risquerait de réintroduire un décalage.
 * @param {Date|number} valeur
 * @return {number}
 */
function dureeEnMillisecondes(valeur) {
  if (valeur instanceof Date) {
    return valeur.getUTCHours() * 3600000 + valeur.getUTCMinutes() * 60000 + valeur.getUTCSeconds() * 1000;
  }
  if (typeof valeur === "number") {
    return Math.round(valeur * 24 * 60 * 60 * 1000);
  }
  return 0;
}

/**
 * Lit la table de délais "Arrivée" (colonnes A/B de Paramètres), utilisée pour ARRIVEES.
 * @return {Object<string, number>} trajet → délai en ms
 */
function obtenirTableDelaisArrivee() {
  const feuille = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_PARAMETRES);
  const derniereLigne = feuille.getLastRow();
  if (derniereLigne < 2) return {};
  const valeurs = feuille.getRange(2, 1, derniereLigne - 1, 2).getValues();

  const table = {};
  valeurs.forEach(ligne => {
    const trajet = ligne[0];
    if (trajet) table[trajet.toString().trim()] = dureeEnMillisecondes(ligne[1]);
  });
  return table;
}

/**
 * Lit la table de délais "Départ" (colonnes D/E de Paramètres), utilisée pour DEPARTS.
 * @return {Object<string, number>} trajet → délai en ms
 */
function obtenirTableDelaisDepart() {
  const feuille = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_PARAMETRES);
  const derniereLigne = feuille.getLastRow();
  if (derniereLigne < 2) return {};
  const valeurs = feuille.getRange(2, 4, derniereLigne - 1, 2).getValues();

  const table = {};
  valeurs.forEach(ligne => {
    const trajet = ligne[0];
    if (trajet) table[trajet.toString().trim()] = dureeEnMillisecondes(ligne[1]);
  });
  return table;
}

/**
 * Déduit la liste des dates disponibles en scannant DateArrivée dans ARRIVEES et
 * DateDépart dans DEPARTS — il n'y a pas de liste de dates dédiée dans Paramètres.
 * @return {Date[]} dates triées chronologiquement
 */
function obtenirDatesDisponibles() {
  const cles = new Set();

  ajouterDatesDepuisFeuille(Global.ONGLET_ARRIVEES, Global.COLONNES_ARRIVEES.DATE_ARRIVEE, cles);
  ajouterDatesDepuisFeuille(Global.ONGLET_DEPARTS, Global.COLONNES_DEPARTS.DATE_DEPART, cles);

  return Array.from(cles)
    .map(cle => Utilities.parseDate(cle, Session.getScriptTimeZone(), "dd/MM/yyyy"))
    .sort((a, b) => a - b);
}

/**
 * Ajoute au Set `cles` la clé formatDateCle de chaque date valide trouvée dans la
 * colonne nomColonneDate de l'onglet nomOnglet.
 * @param {string} nomOnglet
 * @param {string} nomColonneDate
 * @param {Set<string>} cles
 */
function ajouterDatesDepuisFeuille(nomOnglet, nomColonneDate, cles) {
  const feuille = SpreadsheetApp.getActive().getSheetByName(nomOnglet);
  const donnees = feuille.getDataRange().getValues();
  const enTetes = donnees[0];
  const idxDate = enTetes.indexOf(nomColonneDate);

  donnees.slice(1).forEach(ligne => {
    const date = ligne[idxDate];
    if (date instanceof Date) cles.add(formatDateCle(date));
  });
}

/**
 * Convertit une date en nom d'onglet journalier attendu, ex: "VENDREDI 16".
 * @param {Date} date
 * @return {string}
 */
function nomOngletPourDate(date) {
  const jour = Global.JOURS_FR[date.getDay()];
  return jour + " " + date.getDate();
}

/**
 * Colore les lignes de ARRIVEES/DEPARTS par blocs consécutifs partageant la même
 * date, en alternant deux couleurs à chaque changement — pour repérer un jour du
 * suivant d'un coup d'œil. Suppose que les lignes sont déjà groupées par date
 * (lignes consécutives) ; si ce n'est plus le cas un jour, le bloc de couleur se
 * casserait à cet endroit, sans que ce soit un bug.
 * @param {string} nomOnglet
 * @param {string} nomColonneDate
 */
function colorerBlocsParDate(nomOnglet, nomColonneDate) {
  const feuille = SpreadsheetApp.getActive().getSheetByName(nomOnglet);
  const derniereLigne = feuille.getLastRow();
  const derniereColonne = feuille.getLastColumn();
  if (derniereLigne < 2) return;

  const donnees = feuille.getRange(1, 1, derniereLigne, derniereColonne).getValues();
  const enTetes = donnees[0];
  const idxDate = enTetes.indexOf(nomColonneDate);
  if (idxDate === -1) return;

  const couleurs = ["#ffffff", "#f3f3f3"]; // blanc / gris clair — à ajuster au goût
  let indexCouleur = 0;
  let cleDatePrecedente = null;

  for (let i = 1; i < donnees.length; i++) {
    const valeurDate = donnees[i][idxDate];
    const cleDate = (valeurDate instanceof Date) ? formatDateCle(valeurDate) : String(valeurDate);

    if (cleDate !== cleDatePrecedente) {
      indexCouleur = 1 - indexCouleur; // bascule à chaque changement de date
      cleDatePrecedente = cleDate;
    }

    feuille.getRange(i + 1, 1, 1, derniereColonne).setBackground(couleurs[indexCouleur]);
  }
}

/**
 * Applique colorerBlocsParDate aux deux onglets sources, à chaque génération.
 */
function colorerBlocsDeDatesDesSources() {
  colorerBlocsParDate(Global.ONGLET_ARRIVEES, Global.COLONNES_ARRIVEES.DATE_ARRIVEE);
  colorerBlocsParDate(Global.ONGLET_DEPARTS, Global.COLONNES_DEPARTS.DATE_DEPART);
}

/**
 * Parcourt ARRIVEES et retourne toute personne dont DateArrivée tombe dans
 * datesAutorisees — même si ModeArrivée est encore vide ou non reconnu. Dans ce
 * cas heurePickup/lieuPickup restent vides en attendant que le trajet soit renseigné,
 * pour que la personne apparaisse quand même dans le planning du jour.
 * @param {Object<string, number>} tableDelaisArrivee
 * @param {Set<string>} datesAutorisees - clés formatDateCle des dates valides
 * @return {Array<Object>} mouvements
 */
function collecterArrivees(tableDelaisArrivee, datesAutorisees) {
  const feuille = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_ARRIVEES);
  const donnees = feuille.getDataRange().getValues();
  const enTetes = donnees[0];

  const idx = {};
  Object.keys(Global.COLONNES_ARRIVEES).forEach(cle => {
    idx[cle] = enTetes.indexOf(Global.COLONNES_ARRIVEES[cle]);
  });

  const mouvements = [];

  donnees.slice(1).forEach(ligne => {
    if (!ligne[idx.NOM]) return;

    const date = ligne[idx.DATE_ARRIVEE];
    if (!(date instanceof Date)) return;

    const cleDate = formatDateCle(date);
    if (!datesAutorisees.has(cleDate)) return;

    const heureEvenementBrute = formatHeureAffichage(combinerDateEtHeure(date, ligne[idx.HEURE_ARRIVEE]));
    const mode = ligne[idx.MODE_ARRIVEE] ? ligne[idx.MODE_ARRIVEE].toString().trim() : "";

    let heurePickup = "";
    let lieuPickup = "";

    if (tableDelaisArrivee.hasOwnProperty(mode)) {
      const heureEvenement = combinerDateEtHeure(date, ligne[idx.HEURE_ARRIVEE]);
      heurePickup = formatHeureAffichage(new Date(heureEvenement.getTime() - tableDelaisArrivee[mode]));
      lieuPickup = extraireLieu(mode);
    }

    mouvements.push({
      date: date,
      nom: ligne[idx.NOM],
      prenom: ligne[idx.PRENOM],
      fonction: ligne[idx.FONCTION],
      telephone: ligne[idx.TELEPHONE],
      heurePickup: heurePickup,
      lieuPickup: lieuPickup,
      lieuDepose: ligne[idx.HOTEL],
      heureEvenement: heureEvenementBrute
    });
  });

  return mouvements;
}

/**
 * Parcourt DEPARTS et retourne toute personne dont DateDépart tombe dans
 * datesAutorisees — même si ModeDépart est encore vide ou non reconnu (mêmes
 * règles que collecterArrivees). Calcule aussi heureRetourCorum via la table
 * "Arrivée" (trajet inversé) quand le trajet est reconnu — usage interne réservé
 * à une future optimisation, jamais écrit dans un onglet.
 * @param {Object<string, number>} tableDelaisDepart
 * @param {Object<string, number>} tableDelaisArrivee
 * @param {Set<string>} datesAutorisees
 * @return {Array<Object>} mouvements
 */
function collecterDeparts(tableDelaisDepart, tableDelaisArrivee, datesAutorisees) {
  const feuille = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_DEPARTS);
  const donnees = feuille.getDataRange().getValues();
  const enTetes = donnees[0];

  const idx = {};
  Object.keys(Global.COLONNES_DEPARTS).forEach(cle => {
    idx[cle] = enTetes.indexOf(Global.COLONNES_DEPARTS[cle]);
  });

  const mouvements = [];

  donnees.slice(1).forEach(ligne => {
    if (!ligne[idx.NOM]) return;

    const date = ligne[idx.DATE_DEPART];
    if (!(date instanceof Date)) return;

    const cleDate = formatDateCle(date);
    if (!datesAutorisees.has(cleDate)) return;

    const heureEvenementBrute = formatHeureAffichage(combinerDateEtHeure(date, ligne[idx.HEURE_DEPART]));
    const mode = ligne[idx.MODE_DEPART] ? ligne[idx.MODE_DEPART].toString().trim() : "";

    let heurePickup = "";
    let lieuDepose = "";
    let heureRetourCorum = "";

    if (tableDelaisDepart.hasOwnProperty(mode)) {
      const heureEvenement = combinerDateEtHeure(date, ligne[idx.HEURE_DEPART]);
      heurePickup = formatHeureAffichage(new Date(heureEvenement.getTime() - tableDelaisDepart[mode]));
      lieuDepose = extraireLieu(mode);

      const trajetInverse = nomTrajetInverse(mode);
      if (trajetInverse && tableDelaisArrivee.hasOwnProperty(trajetInverse)) {
        const retour = new Date(heureEvenement.getTime() + tableDelaisArrivee[trajetInverse]);
        heureRetourCorum = formatHeureAffichage(retour);
      }
    }

    mouvements.push({
      date: date,
      nom: ligne[idx.NOM],
      prenom: ligne[idx.PRENOM],
      fonction: ligne[idx.FONCTION],
      telephone: ligne[idx.TELEPHONE],
      heurePickup: heurePickup,
      lieuPickup: ligne[idx.HOTEL],
      lieuDepose: lieuDepose,
      heureEvenement: heureEvenementBrute,
      heureRetourCorum: heureRetourCorum // interne uniquement, jamais écrite dans un onglet
    });
  });

  return mouvements;
}

/**
 * Crée l'onglet journalier en copiant le modèle Planning Source s'il n'existe
 * pas encore. Ne fait rien s'il existe déjà (jamais d'écrasement).
 * @param {string} nomOnglet
 * @return {Sheet}
 */
function assurerOngletExiste(nomOnglet) {
  let feuille = SpreadsheetApp.getActive().getSheetByName(nomOnglet);
  if (!feuille) {
    const modele = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_PLANNING_SOURCE);
    if (!modele) {
      throw new Error("L'onglet modèle \"" + Global.ONGLET_PLANNING_SOURCE + "\" est introuvable.");
    }
    feuille = modele.copyTo(SpreadsheetApp.getActive());
    feuille.setName(nomOnglet);
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

    feuille.getBandings().forEach(bande => bande.remove());
    feuille.getRange(1, 1, derniereLigne, derniereColonne)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  }

  feuille.autoResizeColumns(1, derniereColonne);
}

/**
 * Écrit une liste de mouvements dans l'onglet journalier correspondant. La ligne
 * à insérer est construite par position d'en-tête retrouvée dynamiquement (pas
 * par ordre fixe), pour rester correcte même si l'ordre des colonnes du modèle
 * Planning Source change. Ajoute les mouvements absents, met à jour Heure Pick up /
 * Lieu Pick up / Lieu de dépose pour ceux déjà présents, ne touche jamais aux
 * colonnes manuelles (Chauffeur, Nb, Film/Projet, Statut, Pays, Langue), puis trie
 * par ordre chronologique. La clé de dédoublonnage est Nom|Prénom|Heure départ
 * (l'heure brute de l'événement), qui distingue naturellement une arrivée d'un
 * départ pour une même personne.
 * @param {string} nomOnglet
 * @param {Array<Object>} mouvements
 */
function ecrireMouvementsDansOnglet(nomOnglet, mouvements) {
  const feuille = assurerOngletExiste(nomOnglet);

  const derniereColonne = feuille.getLastColumn();
  const enTetes = feuille.getRange(1, 1, 1, derniereColonne).getValues()[0]
    .map(valeur => (valeur || "").toString().trim());

  const idxNom         = enTetes.indexOf("Nom") + 1;
  const idxPrenom      = enTetes.indexOf("Prénom") + 1;
  const idxFonction    = enTetes.indexOf("Fonction") + 1;
  const idxTelephone   = enTetes.indexOf("Téléphone") + 1;
  const idxHeurePickup = enTetes.indexOf("Heure Pick up") + 1;
  const idxLieuPickup  = enTetes.indexOf("Lieu Pick up") + 1;
  const idxLieuDepose  = enTetes.indexOf("Lieu de dépose") + 1;
  const idxHeureDepart = enTetes.indexOf("Heure départ") + 1;

  const derniereLigneAvant = feuille.getLastRow();
  const dejaPresents = new Map();
  if (derniereLigneAvant > 1) {
    const existants = feuille.getRange(2, 1, derniereLigneAvant - 1, derniereColonne).getValues();
    existants.forEach((ligne, i) => {
      const cle = ligne[idxNom - 1] + "|" + ligne[idxPrenom - 1] + "|" + ligne[idxHeureDepart - 1];
      dejaPresents.set(cle, i + 2);
    });
  }

  const nouvellesLignes = [];
  let quelqueChoseAChange = false;

  mouvements.forEach(mvt => {
    const cle = mvt.nom + "|" + mvt.prenom + "|" + mvt.heureEvenement;

    if (dejaPresents.has(cle)) {
      const ligneExistante = dejaPresents.get(cle);
      const heureActuelle     = feuille.getRange(ligneExistante, idxHeurePickup).getValue();
      const lieuPickupActuel  = feuille.getRange(ligneExistante, idxLieuPickup).getValue();
      const lieuDeposeActuel  = feuille.getRange(ligneExistante, idxLieuDepose).getValue();

      if (heureActuelle !== mvt.heurePickup || lieuPickupActuel !== mvt.lieuPickup || lieuDeposeActuel !== mvt.lieuDepose) {
        feuille.getRange(ligneExistante, idxHeurePickup).setValue(mvt.heurePickup);
        feuille.getRange(ligneExistante, idxLieuPickup).setValue(mvt.lieuPickup);
        feuille.getRange(ligneExistante, idxLieuDepose).setValue(mvt.lieuDepose);
        quelqueChoseAChange = true;
      }
    } else {
      const ligne = new Array(derniereColonne).fill("");
      ligne[idxNom - 1]         = mvt.nom;
      ligne[idxPrenom - 1]      = mvt.prenom;
      ligne[idxFonction - 1]    = mvt.fonction;
      ligne[idxTelephone - 1]   = mvt.telephone;
      ligne[idxHeurePickup - 1] = mvt.heurePickup;
      ligne[idxLieuPickup - 1]  = mvt.lieuPickup;
      ligne[idxLieuDepose - 1]  = mvt.lieuDepose;
      ligne[idxHeureDepart - 1] = mvt.heureEvenement;
      // Chauffeur, Nb, Film/Projet, Statut, Pays, Langue restent vides (pas de source / manuel)
      nouvellesLignes.push(ligne);
    }
  });

  if (nouvellesLignes.length > 0) {
    feuille.getRange(feuille.getLastRow() + 1, 1, nouvellesLignes.length, derniereColonne)
      .setValues(nouvellesLignes);
    quelqueChoseAChange = true;
  }

  if (!quelqueChoseAChange) return;

  const derniereLigneApres = feuille.getLastRow();
  if (derniereLigneApres > 1) {
    feuille.getRange(2, 1, derniereLigneApres - 1, derniereColonne)
      .sort({ column: idxHeurePickup, ascending: true });
  }

  formaterOnglet(feuille);
}

/**
 * Orchestration : colore les sources par blocs de dates, lit les deux tables de
 * délais une seule fois chacune, collecte arrivées et départs avec leur table
 * respective, puis écrit dans les onglets concernés. Un onglet n'est créé (copie
 * de Planning Source) que pour une date ayant au moins une personne.
 * @param {Date[]} [datesCiblees]
 */
function genererPlannings(datesCiblees) {
  colorerBlocsDeDatesDesSources(); // mise en forme des sources à chaque génération

  const dates = (datesCiblees && datesCiblees.length) ? datesCiblees : obtenirDatesDisponibles();

  const tableDelaisArrivee = obtenirTableDelaisArrivee();
  const tableDelaisDepart = obtenirTableDelaisDepart();
  const clesAutorisees = new Set(dates.map(formatDateCle));

  const mouvements = collecterArrivees(tableDelaisArrivee, clesAutorisees)
    .concat(collecterDeparts(tableDelaisDepart, tableDelaisArrivee, clesAutorisees));

  const parOnglet = {};
  mouvements.forEach(mvt => {
    const nomOnglet = nomOngletPourDate(mvt.date);
    if (!parOnglet[nomOnglet]) parOnglet[nomOnglet] = [];
    parOnglet[nomOnglet].push(mvt);
  });

  Object.keys(parOnglet).forEach(nomOnglet => ecrireMouvementsDansOnglet(nomOnglet, parOnglet[nomOnglet]));
}

/**
 * Supprime tous les onglets de planning générés (nom au format "JOUR NUMÉRO",
 * ex: "VENDREDI 16"). Ne touche jamais à ARRIVEES, DEPARTS, Paramètres ou
 * Planning Source, puisque ces noms ne correspondent jamais à ce motif.
 */
function supprimerPlanningsGeneres() {
  const classeur = SpreadsheetApp.getActive();
  const motifNomJour = new RegExp("^(" + Global.JOURS_FR.join("|") + ") \\d{1,2}$");

  classeur.getSheets().forEach(feuille => {
    if (motifNomJour.test(feuille.getName())) {
      classeur.deleteSheet(feuille);
    }
  });
}
