/**
 * Common.gs — Logique métier : lecture de la table Paramètres (unifiée),
 * détection de lieu depuis un texte libre, dates, collecte des mouvements,
 * écriture (ajout uniquement), mise en forme (plannings et sources),
 * suppression. Aucun déclencheur ici, uniquement des fonctions appelées par
 * Main.gs.
 */

/**
 * Normalise un nom de lieu pour comparaison : retire les espaces, met en
 * majuscules. Permet de faire matcher "St Roch" et "StRoch".
 * @param {string} texte
 * @return {string}
 */
function normaliserStation(texte) {
  if (!texte) return "";
  return texte.toString().replace(/\s+/g, "").toUpperCase();
}

/**
 * Extrait l'abréviation de lieu d'un texte libre de ModeArrivée/ModeDépart :
 * "StRoch>Corum", "Corum>SDF", "MRS>Corum ", "PPM", "MRS OS399",
 * "SDF TGV 6047"... Le format "A>B" prend le côté qui n'est pas "Corum" ;
 * sinon on cherche le plus long code connu en préfixe, pour ignorer un
 * numéro de vol/train accolé.
 * @param {string} texteLibre
 * @param {Set<string>} codesConnus
 * @return {string} abréviation normalisée, ou le texte normalisé en repli
 */
function extraireStationDepuisModeLibre(texteLibre, codesConnus) {
  if (!texteLibre) return "";
  const texte = texteLibre.toString().trim();

  const parties = texte.split(">");
  if (parties.length === 2) {
    const gauche = normaliserStation(parties[0]);
    const droite = normaliserStation(parties[1]);
    return gauche === "CORUM" ? droite : gauche;
  }

  const texteNormalise = normaliserStation(texte);

  if (codesConnus) {
    const codesTries = Array.from(codesConnus).sort((a, b) => b.length - a.length);
    for (const code of codesTries) {
      if (code && texteNormalise.startsWith(code)) return code;
    }
  }

  return texteNormalise; // repli : aucun code connu ne correspond, ne matchera probablement aucune table
}

const FONCTIONS_EXCEPTION = ["avant-première", "jury antigone d'or", "jury bourse d'aide"];

/**
 * Indique si une fonction fait partie des exceptions à la règle d'exclusion
 * Saint-Roch/PPM (comparaison insensible à la casse/aux espaces superflus).
 * @param {string} fonction
 * @return {boolean}
 */
function estFonctionException(fonction) {
  if (!fonction) return false;
  return FONCTIONS_EXCEPTION.includes(fonction.toString().trim().toLowerCase());
}

/**
 * @param {Date} date
 * @return {number} minutes écoulées depuis minuit
 */
function minutesDepuisMinuit(date) {
  return date.getHours() * 60 + date.getMinutes();
}

/**
 * Vrai pour Saint-Roch ("SR") ou pour un code "moyen propre" (PPM) : les deux
 * suivent les mêmes exceptions.
 * @param {string} station - abréviation déjà extraite/normalisée
 * @return {boolean}
 */
function estStationSoumiseAExceptions(station) {
  if (station === "SR") return true;
  return Global.CODES_MOYEN_PROPRE.some(code => normaliserStation(code) === station);
}

/**
 * Règle Saint-Roch/PPM — arrivée : exclue sauf fonction exception ou arrivée après 21h.
 * @param {string} station
 * @param {Date} heureEvenement
 * @param {string} fonction
 * @return {boolean}
 */
function estExcluArrivee(station, heureEvenement, fonction) {
  if (!estStationSoumiseAExceptions(station)) return false;
  if (estFonctionException(fonction)) return false;
  return minutesDepuisMinuit(heureEvenement) <= 21 * 60; // pas "après 21h" → exclue
}

/**
 * Règle Saint-Roch/PPM — départ : exclu sauf fonction exception ou départ avant 8h.
 * @param {string} station
 * @param {Date} heureEvenement
 * @param {string} fonction
 * @return {boolean}
 */
function estExcluDepart(station, heureEvenement, fonction) {
  if (!estStationSoumiseAExceptions(station)) return false;
  if (estFonctionException(fonction)) return false;
  return minutesDepuisMinuit(heureEvenement) >= 8 * 60; // pas "avant 8h" → exclu
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
 * Normalise une valeur de cellule "heure" lue via getValue() en texte "HH:mm",
 * qu'elle soit restée une chaîne ou que Sheets l'ait réinterprétée en heure/date
 * à l'écriture (même mécanisme qu'une saisie manuelle). Nécessaire pour
 * dédoublonner de façon fiable quel que soit le type effectivement stocké.
 * @param {Date|string} valeur
 * @return {string}
 */
function normaliserValeurHeure(valeur) {
  if (valeur instanceof Date) return formatHeureAffichage(valeur);
  return (valeur || "").toString();
}

/**
 * Préfixe une chaîne d'une apostrophe pour forcer Sheets à la garder en texte
 * littéral à l'écriture (setValue/setValues) : sans ça, une valeur qui
 * ressemble à une heure (ex. "15:05") peut être réinterprétée en heure/date,
 * comme lors d'une saisie manuelle — même sur une cellule au format "@".
 * L'apostrophe n'apparaît jamais à l'affichage ni dans un getValue()
 * ultérieur — à utiliser uniquement à l'écriture, jamais pour les
 * comparaisons (normaliserValeurHeure s'en charge côté lecture).
 * @param {string} texte
 * @return {string}
 */
function forcerTexteLitteral(texte) {
  return texte ? "'" + texte : texte;
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
 * Lit la table `Paramètres` (colonnes A à E), indexée sur la colonne B
 * (Abbréviations 4D) normalisée. La colonne A (nom complet) est ignorée par
 * le script, purement informative pour les humains. Une ligne dont les
 * colonnes C et D sont toutes les deux vides n'est pas ajoutée (cas de SPE
 * et PPM, qui n'ont pas de délai chronométrable) : ces codes resteront avec
 * Heure Pick up vide.
 * @return {Object<string, {delaiDepart:number, delaiArrivee:number, dureeOccupation:number}>}
 */
function obtenirTableLieux() {
  const feuille = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_PARAMETRES);
  const derniereLigne = feuille.getLastRow();
  if (derniereLigne < 2) return {};

  const valeurs = feuille.getRange(2, 1, derniereLigne - 1, 5).getValues(); // colonnes A à E

  const table = {};
  valeurs.forEach(ligne => {
    const abbreviation = normaliserStation(ligne[1]); // colonne B
    if (!abbreviation) return;

    const celluleDepart = ligne[2];  // colonne C
    const celluleArrivee = ligne[3]; // colonne D
    if (!celluleDepart && !celluleArrivee) return; // ex. SPE, PPM : pas de délai chronométrable

    table[abbreviation] = {
      delaiDepart: dureeEnMillisecondes(celluleDepart),
      delaiArrivee: dureeEnMillisecondes(celluleArrivee),
      dureeOccupation: dureeEnMillisecondes(ligne[4]) // colonne E, jamais affichée
    };
  });
  return table;
}

/**
 * Rassemble les codes utilisables pour la détection par préfixe : les lieux
 * réels de la table Paramètres (avec délai) + les codes "moyen propre" (PPM),
 * même sans délai associé.
 * @param {Object} tableLieux
 * @return {Set<string>}
 */
function obtenirCodesStationConnus(tableLieux) {
  const codes = new Set(Object.keys(tableLieux));
  Global.CODES_MOYEN_PROPRE.forEach(code => codes.add(normaliserStation(code)));
  return codes;
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
 * cas heurePickup/lieuPickup restent vides en attendant que le lieu soit
 * reconnu, pour que la personne apparaisse quand même dans le planning du
 * jour. Une personne dont le lieu détecté est Saint-Roch ("SR"), ou dont le
 * mode est un code "moyen propre" (PPM), est retirée du résultat si la règle
 * d'exclusion s'applique (cf. estExcluArrivee) — évaluable seulement si
 * l'heure est connue, sinon la personne reste visible.
 * @param {Object} tableLieux
 * @param {Set<string>} datesAutorisees - clés formatDateCle des dates valides
 * @param {Set<string>} codesConnus - codes connus, pour extraireStationDepuisModeLibre
 * @return {Array<Object>} mouvements
 */
function collecterArrivees(tableLieux, datesAutorisees, codesConnus) {
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

    const heureArrivee = ligne[idx.HEURE_ARRIVEE];
    const heureRenseignee = heureArrivee instanceof Date;
    const heureEvenementDate = heureRenseignee ? combinerDateEtHeure(date, heureArrivee) : null;
    const heureEvenementBrute = heureEvenementDate ? formatHeureAffichage(heureEvenementDate) : "";
    const modeLibre = ligne[idx.MODE_ARRIVEE] ? ligne[idx.MODE_ARRIVEE].toString().trim() : "";
    const station = extraireStationDepuisModeLibre(modeLibre, codesConnus);

    // Règle Saint-Roch/PPM : évaluable seulement si l'heure est connue ; sinon
    // la personne reste visible, comme pour tout lieu non encore reconnu.
    if (heureEvenementDate && estExcluArrivee(station, heureEvenementDate, ligne[idx.FONCTION])) return;

    let heurePickup = "";
    let lieuPickup = "";
    let dureeOccupationChauffeur = 0;

    if (heureRenseignee && station && tableLieux.hasOwnProperty(station)) {
      const infosLieu = tableLieux[station];
      heurePickup = formatHeureAffichage(new Date(heureEvenementDate.getTime() - infosLieu.delaiArrivee));
      lieuPickup = station;
      dureeOccupationChauffeur = infosLieu.dureeOccupation; // interne, jamais affichée
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
      heureEvenement: heureEvenementBrute,
      dureeOccupationChauffeur: dureeOccupationChauffeur // interne uniquement, jamais écrite dans un onglet
    });
  });

  return mouvements;
}

/**
 * Parcourt DEPARTS et retourne toute personne dont DateDépart tombe dans
 * datesAutorisees — même si ModeDépart est encore vide ou non reconnu (mêmes
 * règles que collecterArrivees, y compris la règle Saint-Roch/PPM via
 * estExcluDepart).
 * @param {Object} tableLieux
 * @param {Set<string>} datesAutorisees
 * @param {Set<string>} codesConnus
 * @return {Array<Object>} mouvements
 */
function collecterDeparts(tableLieux, datesAutorisees, codesConnus) {
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

    const heureDepart = ligne[idx.HEURE_DEPART];
    const heureRenseignee = heureDepart instanceof Date;
    const heureEvenementDate = heureRenseignee ? combinerDateEtHeure(date, heureDepart) : null;
    const heureEvenementBrute = heureEvenementDate ? formatHeureAffichage(heureEvenementDate) : "";
    const modeLibre = ligne[idx.MODE_DEPART] ? ligne[idx.MODE_DEPART].toString().trim() : "";
    const station = extraireStationDepuisModeLibre(modeLibre, codesConnus);

    // Règle Saint-Roch/PPM : évaluable seulement si l'heure est connue ; sinon
    // la personne reste visible, comme pour tout lieu non encore reconnu.
    if (heureEvenementDate && estExcluDepart(station, heureEvenementDate, ligne[idx.FONCTION])) return;

    let heurePickup = "";
    let lieuDepose = "";
    let dureeOccupationChauffeur = 0;

    if (heureRenseignee && station && tableLieux.hasOwnProperty(station)) {
      const infosLieu = tableLieux[station];
      heurePickup = formatHeureAffichage(new Date(heureEvenementDate.getTime() - infosLieu.delaiDepart));
      lieuDepose = station;
      dureeOccupationChauffeur = infosLieu.dureeOccupation; // interne, jamais affichée
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
      dureeOccupationChauffeur: dureeOccupationChauffeur // interne uniquement, jamais écrite dans un onglet
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
 * Écrit une liste de mouvements dans l'onglet journalier correspondant, en
 * AJOUT UNIQUEMENT : une ligne déjà présente (clé Nom|Prénom|Heure de
 * l'événement) est ignorée, elle ne sera plus jamais modifiée par une
 * génération ultérieure — seules les nouvelles personnes sont ajoutées à
 * chaque génération. Une correction après coup se fait à la main directement
 * dans l'onglet de planning. La ligne à insérer est construite par position
 * d'en-tête retrouvée dynamiquement (pas par ordre fixe), pour rester
 * correcte même si l'ordre des colonnes du modèle Planning Source change. Ne
 * touche jamais aux colonnes manuelles (Chauffeur, Nb, Film/Projet, Statut,
 * Pays, Langue).
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

  // Empêche Sheets de réinterpréter une écriture future en heure/date (même
  // mécanisme que pour une saisie manuelle) — combiné à forcerTexteLitteral()
  // à l'écriture ci-dessous. Sans ça, la clé de dédoublonnage (basée sur
  // Heure départ) ne matcherait plus jamais et la même personne serait
  // ajoutée en double à chaque génération.
  if (idxHeurePickup > 0) {
    feuille.getRange(1, idxHeurePickup, feuille.getMaxRows(), 1).setNumberFormat("@");
  }
  if (idxHeureDepart > 0) {
    feuille.getRange(1, idxHeureDepart, feuille.getMaxRows(), 1).setNumberFormat("@");
  }

  const derniereLigneAvant = feuille.getLastRow();
  const dejaPresents = new Set();
  if (derniereLigneAvant > 1) {
    const existants = feuille.getRange(2, 1, derniereLigneAvant - 1, derniereColonne).getValues();
    existants.forEach(ligne => {
      const cle = ligne[idxNom - 1] + "|" + ligne[idxPrenom - 1] + "|" + normaliserValeurHeure(ligne[idxHeureDepart - 1]);
      dejaPresents.add(cle);
    });
  }

  const nouvellesLignes = [];

  mouvements.forEach(mvt => {
    const cle = mvt.nom + "|" + mvt.prenom + "|" + mvt.heureEvenement;
    if (dejaPresents.has(cle)) return; // déjà présente, on n'y touche plus jamais

    const ligne = new Array(derniereColonne).fill("");
    ligne[idxNom - 1]         = mvt.nom;
    ligne[idxPrenom - 1]      = mvt.prenom;
    ligne[idxFonction - 1]    = mvt.fonction;
    ligne[idxTelephone - 1]   = mvt.telephone;
    ligne[idxHeurePickup - 1] = forcerTexteLitteral(mvt.heurePickup);
    ligne[idxLieuPickup - 1]  = mvt.lieuPickup;
    ligne[idxLieuDepose - 1]  = mvt.lieuDepose;
    ligne[idxHeureDepart - 1] = forcerTexteLitteral(mvt.heureEvenement);
    // Chauffeur, Nb, Film/Projet, Statut, Pays, Langue restent vides (pas de source / manuel)
    nouvellesLignes.push(ligne);
  });

  if (nouvellesLignes.length === 0) return; // rien de nouveau, on ne touche à rien

  feuille.getRange(feuille.getLastRow() + 1, 1, nouvellesLignes.length, derniereColonne)
    .setValues(nouvellesLignes);

  const derniereLigneApres = feuille.getLastRow();
  feuille.getRange(2, 1, derniereLigneApres - 1, derniereColonne)
    .sort({ column: idxHeurePickup, ascending: true });

  formaterOnglet(feuille);
}

/**
 * Orchestration : colore les sources par blocs de dates, lit la table
 * Paramètres unifiée une seule fois, en déduit les codes connus, collecte
 * arrivées et départs, puis écrit (ajout uniquement) dans les onglets
 * concernés. Un onglet n'est créé (copie de Planning Source) que pour une
 * date ayant au moins une personne.
 * @param {Date[]} [datesCiblees]
 */
function genererPlannings(datesCiblees) {
  colorerBlocsDeDatesDesSources(); // mise en forme des sources à chaque génération

  const dates = (datesCiblees && datesCiblees.length) ? datesCiblees : obtenirDatesDisponibles();

  const tableLieux = obtenirTableLieux();
  const codesConnus = obtenirCodesStationConnus(tableLieux);
  const clesAutorisees = new Set(dates.map(formatDateCle));

  const mouvements = collecterArrivees(tableLieux, clesAutorisees, codesConnus)
    .concat(collecterDeparts(tableLieux, clesAutorisees, codesConnus));

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
