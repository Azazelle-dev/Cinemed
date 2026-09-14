/**
 * Common.gs — Logique métier : lecture des tables de délais, collecte des
 * mouvements confirmés (arrivées + départs séparément), écriture non
 * destructive dans les onglets journaliers, mise en forme.
 * Aucun déclencheur ici, uniquement des fonctions appelées par Main.gs.
 */

/**
 * Inverse un identifiant de trajet "A>B" en "B>A", pour retrouver l'entrée
 * correspondante dans l'autre table de délais sans maintenir une correspondance séparée.
 * @param {string} trajet
 * @return {string|null} null si le format n'est pas celui attendu
 */
function nomTrajetInverse(trajet) {
  const parties = trajet.split(">");
  if (parties.length !== 2) return null; // format inattendu, on ne devine pas
  return parties[1] + ">" + parties[0];
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
 * Convertit un objet Date en texte "HH:mm" pur (ex: "15:05"). On écrit ce texte
 * directement dans la cellule, jamais l'objet Date complet, pour éviter tout
 * risque que Sheets affiche la date en plus de l'heure.
 * @param {Date} date
 * @return {string}
 */
function formatHeureAffichage(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "HH:mm");
}

/**
 * Préfixe une chaîne d'une apostrophe pour forcer Sheets à la garder en texte
 * littéral à l'écriture (setValue/setValues). Même une cellule au format "@"
 * peut voir une valeur qui ressemble à une heure réinterprétée en date/heure :
 * c'est le même mécanisme de détection qu'une saisie manuelle, appliqué aussi
 * via l'API. L'apostrophe n'apparaît jamais à l'affichage ni dans un getValue()
 * ultérieur — à utiliser uniquement au moment d'écrire, jamais pour les comparaisons.
 * @param {string} texte
 * @return {string}
 */
function forcerTexteLitteral(texte) {
  return "'" + texte;
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
    return Math.round(valeur * 24 * 60 * 60 * 1000); // au cas où la cellule est une fraction de jour
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
  const valeurs = feuille.getRange(2, 1, derniereLigne - 1, 2).getValues(); // colonnes A et B

  const table = {};
  valeurs.forEach(ligne => {
    const trajet = ligne[0];
    if (trajet) table[trajet] = dureeEnMillisecondes(ligne[1]);
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
  const valeurs = feuille.getRange(2, 4, derniereLigne - 1, 2).getValues(); // colonnes D et E

  const table = {};
  valeurs.forEach(ligne => {
    const trajet = ligne[0];
    if (trajet) table[trajet] = dureeEnMillisecondes(ligne[1]);
  });
  return table;
}

/**
 * Lit la liste des dates disponibles depuis Paramètres (colonne G, à partir de la ligne 2) —
 * seule source de vérité pour savoir quelles dates sont valides.
 * @return {Date[]} dates triées chronologiquement
 */
function obtenirDatesDisponibles() {
  const feuille = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_PARAMETRES);
  const derniereLigne = feuille.getLastRow();
  if (derniereLigne < 2) return [];
  const valeurs = feuille.getRange(2, 7, derniereLigne - 1, 1).getValues(); // colonne G

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
 * Parcourt ARRIVEES et retourne les mouvements confirmés dont ModeArrivée est
 * une clé connue de tableDelaisArrivee et dont DateArrivée tombe dans datesAutorisees.
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

    const mode = ligne[idx.MODE_ARRIVEE];
    const date = ligne[idx.DATE_ARRIVEE];
    if (!tableDelaisArrivee.hasOwnProperty(mode) || !(date instanceof Date)) return;

    const cleDate = formatDateCle(date);
    if (!datesAutorisees.has(cleDate)) return;

    const heureEvenement = combinerDateEtHeure(date, ligne[idx.HEURE_ARRIVEE]);
    const heurePickup = new Date(heureEvenement.getTime() - tableDelaisArrivee[mode]);

    mouvements.push({
      date: date,
      type: "Arrivée",
      nom: ligne[idx.NOM],
      prenom: ligne[idx.PRENOM],
      telephone: ligne[idx.TELEPHONE],
      hotel: ligne[idx.HOTEL],
      trajet: mode,
      heurePickup: formatHeureAffichage(heurePickup)
    });
  });

  return mouvements;
}

/**
 * Parcourt DEPARTS et retourne les mouvements confirmés dont ModeDépart est une
 * clé connue de tableDelaisDepart et dont DateDépart tombe dans datesAutorisees.
 * Calcule aussi heureRetourCorum via la table "Arrivée" (trajet inversé) — valeur
 * gardée en mémoire pour un usage futur, jamais écrite dans un onglet.
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

    const mode = ligne[idx.MODE_DEPART];
    const date = ligne[idx.DATE_DEPART];
    if (!tableDelaisDepart.hasOwnProperty(mode) || !(date instanceof Date)) return;

    const cleDate = formatDateCle(date);
    if (!datesAutorisees.has(cleDate)) return;

    const heureEvenement = combinerDateEtHeure(date, ligne[idx.HEURE_DEPART]);
    const heurePickup = new Date(heureEvenement.getTime() - tableDelaisDepart[mode]);

    // Heure de retour au Corum : le chauffeur dépose la personne à heureEvenement,
    // puis revient au Corum en utilisant le délai du trajet inversé (table A/B)
    let heureRetourCorum = "";
    const trajetInverse = nomTrajetInverse(mode);
    if (trajetInverse && tableDelaisArrivee.hasOwnProperty(trajetInverse)) {
      const retour = new Date(heureEvenement.getTime() + tableDelaisArrivee[trajetInverse]);
      heureRetourCorum = formatHeureAffichage(retour);
    }

    mouvements.push({
      date: date,
      type: "Départ",
      nom: ligne[idx.NOM],
      prenom: ligne[idx.PRENOM],
      telephone: ligne[idx.TELEPHONE],
      hotel: ligne[idx.HOTEL],
      trajet: mode,
      heurePickup: formatHeureAffichage(heurePickup),
      heureRetourCorum: heureRetourCorum
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
 * Planning Source change. Ajoute les mouvements absents, met à jour Heure de
 * pick up / Trajet pour ceux déjà présents, ne touche jamais à Chauffeurs
 * (assignation manuelle), puis trie par ordre chronologique.
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
  const idxHeurePickup = enTetes.indexOf("Heure de pick up") + 1;
  const idxType        = enTetes.indexOf("Arrivée/Départ") + 1;
  const idxTrajet      = enTetes.indexOf("Trajet") + 1;
  const idxTelephone   = enTetes.indexOf("Téléphone") + 1;
  const idxHotel       = enTetes.indexOf("Hôtel") + 1;
  // Pas d'index pour "Heure retour Corum" : cette valeur n'est jamais écrite dans le sheet

  // Sans ça, Sheets réinterprète une chaîne comme "15:05" écrite via setValue comme
  // une heure/date (même comportement qu'une saisie manuelle). Le format "@" (texte
  // brut) réduit ce risque, combiné à forcerTexteLitteral() à l'écriture.
  if (idxHeurePickup > 0) {
    feuille.getRange(1, idxHeurePickup, feuille.getMaxRows(), 1).setNumberFormat("@");
  }

  // Lignes déjà présentes (clé Nom|Prénom|Type → numéro de ligne)
  const derniereLigneAvant = feuille.getLastRow();
  const dejaPresents = new Map();
  if (derniereLigneAvant > 1) {
    const existants = feuille.getRange(2, 1, derniereLigneAvant - 1, derniereColonne).getValues();
    existants.forEach((ligne, i) => {
      const cle = ligne[idxNom - 1] + "|" + ligne[idxPrenom - 1] + "|" + ligne[idxType - 1];
      dejaPresents.set(cle, i + 2);
    });
  }

  const nouvellesLignes = [];
  let quelqueChoseAChange = false;

  mouvements.forEach(mvt => {
    const cle = mvt.nom + "|" + mvt.prenom + "|" + mvt.type;

    if (dejaPresents.has(cle)) {
      const ligneExistante = dejaPresents.get(cle);
      const heureActuelle = feuille.getRange(ligneExistante, idxHeurePickup).getValue();
      const trajetActuel  = feuille.getRange(ligneExistante, idxTrajet).getValue();

      if (heureActuelle !== mvt.heurePickup || trajetActuel !== mvt.trajet) {
        feuille.getRange(ligneExistante, idxHeurePickup).setValue(forcerTexteLitteral(mvt.heurePickup));
        feuille.getRange(ligneExistante, idxTrajet).setValue(mvt.trajet);
        quelqueChoseAChange = true;
      }
    } else {
      const ligne = new Array(derniereColonne).fill("");
      ligne[idxNom - 1]         = mvt.nom;
      ligne[idxPrenom - 1]      = mvt.prenom;
      ligne[idxHeurePickup - 1] = forcerTexteLitteral(mvt.heurePickup);
      ligne[idxType - 1]        = mvt.type;
      ligne[idxTrajet - 1]      = mvt.trajet;
      ligne[idxTelephone - 1]   = mvt.telephone;
      ligne[idxHotel - 1]       = mvt.hotel;
      // "heureRetourCorum" existe sur mvt mais n'est volontairement jamais écrit dans le sheet
      // Chauffeurs (idxChauffeurs) reste vide : assignation manuelle
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
 * Orchestration : crée un onglet pour chaque date ciblée (même sans mouvement
 * confirmé), lit les deux tables de délais une seule fois chacune, puis traite
 * arrivées et départs avec leur table respective.
 * @param {Date[]} [datesCiblees]
 */
function genererPlannings(datesCiblees) {
  const dates = (datesCiblees && datesCiblees.length) ? datesCiblees : obtenirDatesDisponibles();

  // 1. Un onglet pour chaque date, même sans aucun mouvement confirmé
  dates.forEach(date => assurerOngletExiste(nomOngletPourDate(date)));

  // 2. Collecte des mouvements confirmés, chaque source avec sa propre table de délais
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
