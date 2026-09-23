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
// Référence Sheets (30/12/1899, minuit) construite de la même façon que les
// valeurs de durée lues depuis la feuille, pour que le décalage historique de
// fuseau horaire (Paris Mean Time ≈ UTC+0:09:21 avant 1911) s'annule dans la
// soustraction plutôt que de fausser le résultat (getUTCHours() sur une date
// de 1899 restait décalé d'environ 9 minutes par rapport à l'heure locale
// attendue).
const EPOQUE_SHEETS = new Date(1899, 11, 30, 0, 0, 0, 0);

function dureeEnMillisecondes(valeur) {
  if (valeur instanceof Date) {
    return valeur.getTime() - EPOQUE_SHEETS.getTime();
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
 * Lit toutes les abréviations réellement présentes en colonne B de
 * Paramètres, y compris celles sans délai chronométrable (ex. SPE) —
 * contrairement à obtenirTableLieux() qui ne garde que les lieux avec délai.
 * Sert à distinguer une abréviation simplement pas encore chronométrable
 * (reconnue, mais Heure Pick up restera vide) d'une abréviation carrément
 * absente de Paramètres (faute de frappe, ancien format non ressaisi...).
 * @return {Set<string>} abréviations normalisées présentes dans Paramètres
 */
function obtenirAbreviationsParametres() {
  const feuille = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_PARAMETRES);
  const derniereLigne = feuille.getLastRow();
  if (derniereLigne < 2) return new Set();

  const valeurs = feuille.getRange(2, 2, derniereLigne - 1, 1).getValues(); // colonne B seule
  const codes = new Set();
  valeurs.forEach(ligne => {
    const abbreviation = normaliserStation(ligne[0]);
    if (abbreviation) codes.add(abbreviation);
  });
  return codes;
}

/**
 * Rassemble les codes utilisables pour la détection par préfixe : toutes les
 * abréviations présentes dans Paramètres (avec ou sans délai) + les codes
 * "moyen propre" (PPM), même si PPM n'y est pas encore saisi.
 * @param {Object} tableLieux
 * @return {Set<string>}
 */
function obtenirCodesStationConnus(tableLieux) {
  const codes = obtenirAbreviationsParametres();
  Object.keys(tableLieux).forEach(code => codes.add(code));
  Global.CODES_MOYEN_PROPRE.forEach(code => codes.add(normaliserStation(code)));
  return codes;
}

/**
 * Déduit la liste des dates disponibles en scannant DateArrivée et DateDépart
 * dans ARRIVEES (les deux jeux de colonnes vivent sur la même feuille) — il
 * n'y a pas de liste de dates dédiée dans Paramètres.
 * @return {Date[]} dates triées chronologiquement
 */
function obtenirDatesDisponibles() {
  const cles = new Set();

  ajouterDatesDepuisFeuille(Global.ONGLET_ARRIVEES, Global.COLONNES_ARRIVEES.DATE_ARRIVEE.position, cles);
  ajouterDatesDepuisFeuille(Global.ONGLET_ARRIVEES, Global.COLONNES_DEPARTS.DATE_DEPART.position, cles);

  return Array.from(cles)
    .map(cle => Utilities.parseDate(cle, Session.getScriptTimeZone(), "dd/MM/yyyy"))
    .sort((a, b) => a - b);
}

/**
 * Ajoute au Set `cles` la clé formatDateCle de chaque date valide trouvée à la
 * position positionColonneDate (fixe) de l'onglet nomOnglet.
 * @param {string} nomOnglet
 * @param {number} positionColonneDate
 * @param {Set<string>} cles
 */
function ajouterDatesDepuisFeuille(nomOnglet, positionColonneDate, cles) {
  const feuille = SpreadsheetApp.getActive().getSheetByName(nomOnglet);
  const donnees = feuille.getDataRange().getValues();

  donnees.slice(1).forEach(ligne => {
    const date = ligne[positionColonneDate];
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
 * Convertit une position de colonne 0-based en lettre de colonne façon Sheets
 * ("A", "B", ... "Z", "AA", ...), pour des messages d'erreur lisibles.
 * @param {number} position
 * @return {string}
 */
function lettreColonne(position) {
  let n = position + 1;
  let lettre = "";
  while (n > 0) {
    const reste = (n - 1) % 26;
    lettre = String.fromCharCode(65 + reste) + lettre;
    n = Math.floor((n - 1) / 26);
  }
  return lettre;
}

/**
 * Vérifie que les en-têtes réels de nomOnglet correspondent, position par
 * position, aux colonnes attendues (Object<cle, {nom, position}>). ARRIVEES
 * est lue par position fixe (jamais par recherche dynamique du nom) pour que
 * le script détecte immédiatement une colonne déplacée plutôt que d'écrire
 * silencieusement au mauvais endroit.
 * @param {string} nomOnglet
 * @param {Object<string, {nom:string, position:number}>} colonnesAttendues
 * @return {Array<{onglet:string, colonne:string, attendu:string, trouve:string}>} écarts trouvés (vide si tout correspond)
 */
function verifierEnTetesOnglet(nomOnglet, colonnesAttendues) {
  const feuille = SpreadsheetApp.getActive().getSheetByName(nomOnglet);
  if (!feuille) {
    return [{ onglet: nomOnglet, colonne: "-", attendu: "(onglet existant)", trouve: "onglet introuvable" }];
  }

  const derniereColonne = feuille.getLastColumn();
  const enTetes = derniereColonne > 0 ? feuille.getRange(1, 1, 1, derniereColonne).getValues()[0] : [];

  const erreurs = [];
  Object.keys(colonnesAttendues).forEach(cle => {
    const attendu = colonnesAttendues[cle];
    const entete = (enTetes[attendu.position] || "").toString().trim();
    if (entete !== attendu.nom) {
      erreurs.push({
        onglet: nomOnglet,
        colonne: lettreColonne(attendu.position),
        attendu: attendu.nom,
        trouve: entete || "(vide)"
      });
    }
  });
  return erreurs;
}

/**
 * Vérifie les en-têtes d'ARRIVEES (colonnes arrivée ET départ, toutes deux sur
 * cette même feuille) avant toute génération. En cas d'écart, affiche une
 * alerte claire listant chaque colonne concernée et retourne false — appelée
 * en tout premier dans genererPlannings() pour que la génération s'arrête net
 * plutôt que d'écrire une donnée au mauvais endroit.
 * @return {boolean}
 */
function verifierEnTetesSources() {
  const erreurs = verifierEnTetesOnglet(Global.ONGLET_ARRIVEES, Global.COLONNES_ARRIVEES)
    .concat(verifierEnTetesOnglet(Global.ONGLET_ARRIVEES, Global.COLONNES_DEPARTS));

  if (erreurs.length === 0) return true;

  const details = erreurs.map(e =>
    "• " + e.onglet + ", colonne " + e.colonne + " : attendu \"" + e.attendu + "\", trouvé \"" + e.trouve + "\""
  ).join("\n");

  SpreadsheetApp.getUi().alert(
    "Génération interrompue — en-têtes de colonnes incorrects",
    "L'ordre des colonnes ne correspond plus à ce qu'attend le script. Corrige les en-têtes " +
    "ci-dessous (ou remets-les à leur position d'origine) avant de relancer la génération :\n\n" + details,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return false;
}

/**
 * Colore les lignes de nomOnglet par blocs consécutifs partageant la même
 * date, en alternant deux couleurs à chaque changement — pour repérer un jour du
 * suivant d'un coup d'œil. Suppose que les lignes sont déjà groupées par date
 * (lignes consécutives) ; si ce n'est plus le cas un jour, le bloc de couleur se
 * casserait à cet endroit, sans que ce soit un bug.
 * @param {string} nomOnglet
 * @param {number} positionColonneDate
 */
function colorerBlocsParDate(nomOnglet, positionColonneDate) {
  const feuille = SpreadsheetApp.getActive().getSheetByName(nomOnglet);
  const derniereLigne = feuille.getLastRow();
  const derniereColonne = feuille.getLastColumn();
  if (derniereLigne < 2) return;

  const donnees = feuille.getRange(1, 1, derniereLigne, derniereColonne).getValues();

  const couleurs = ["#ffffff", "#f3f3f3"]; // blanc / gris clair — à ajuster au goût
  let indexCouleur = 0;
  let cleDatePrecedente = null;

  for (let i = 1; i < donnees.length; i++) {
    const valeurDate = donnees[i][positionColonneDate];
    const cleDate = (valeurDate instanceof Date) ? formatDateCle(valeurDate) : String(valeurDate);

    if (cleDate !== cleDatePrecedente) {
      indexCouleur = 1 - indexCouleur; // bascule à chaque changement de date
      cleDatePrecedente = cleDate;
    }

    feuille.getRange(i + 1, 1, 1, derniereColonne).setBackground(couleurs[indexCouleur]);
  }
}

/**
 * Colore ARRIVEES par blocs de DateArrivée, à chaque génération. Un seul
 * onglet source désormais (les colonnes départ vivent sur les mêmes lignes).
 */
function colorerBlocsDeDatesDesSources() {
  colorerBlocsParDate(Global.ONGLET_ARRIVEES, Global.COLONNES_ARRIVEES.DATE_ARRIVEE.position);
}

/**
 * Parcourt ARRIVEES une seule fois et retourne, pour chaque ligne, jusqu'à
 * deux mouvements distincts : un pour l'arrivée (colonnes DateArrivée/
 * HeureArrivée/ModeArrivée) si DateArrivée tombe dans datesAutorisees, un
 * pour le départ (colonnes DateDépart/HeureDépart/ModeDépart) si DateDépart
 * y tombe — une ligne peut donc produire 0, 1 ou 2 mouvements. Fusionner en
 * une seule fonction (plutôt que deux fonctions relisant chacune toute la
 * feuille) évite toute ambiguïté sur l'origine réelle de chaque mouvement
 * (source d'un bug où "Arrivée/Départ" affichait toujours "Arrivée").
 *
 * Mêmes règles des deux côtés : un mode vide ou non reconnu laisse la
 * personne visible (heurePickup/lieuPickup ou lieuDepose restent vides) ; un
 * mode renseigné mais non reconnu est signalé comme erreur ; la règle
 * Saint-Roch/PPM ne s'évalue que si l'heure est connue ; lieuPickup/
 * lieuDepose (copie brute du mode saisi) ne sont remplis que si
 * l'abréviation est reconnue — une abréviation non reconnue est une erreur,
 * elle ne doit pas être recopiée.
 * @param {Object} tableLieux
 * @param {Set<string>} datesAutorisees - clés formatDateCle des dates valides
 * @param {Set<string>} codesConnus - codes connus, pour extraireStationDepuisModeLibre
 * @param {Array<Object>} erreurs - complété (mutation) avec une entrée par
 *   mode renseigné (arrivée ou départ) mais dont l'abréviation ne correspond
 *   à aucune entrée de Paramètres (faute de frappe, ancien format non
 *   ressaisi...)
 * @param {Array<Object>} lignesVerifiees - complété (mutation) avec {onglet, ligne}
 *   pour chaque ligne effectivement passée en revue cette génération (côté
 *   arrivée et/ou départ) — sert à ne remettre en noir (marquerLignesEnErreur)
 *   que les lignes réellement revérifiées, pas toute la feuille, en cas de
 *   génération partielle (dates spécifiques).
 * @return {Array<Object>} mouvements
 */
function collecterMouvements(tableLieux, datesAutorisees, codesConnus, erreurs, lignesVerifiees) {
  const feuille = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_ARRIVEES);
  const donnees = feuille.getDataRange().getValues();

  // Positions fixes (pas de recherche dynamique) : verifierEnTetesSources() a
  // déjà garanti, avant l'appel à cette fonction, que ces positions sont bonnes.
  const idxA = {};
  Object.keys(Global.COLONNES_ARRIVEES).forEach(cle => {
    idxA[cle] = Global.COLONNES_ARRIVEES[cle].position;
  });
  const idxD = {};
  Object.keys(Global.COLONNES_DEPARTS).forEach(cle => {
    idxD[cle] = Global.COLONNES_DEPARTS[cle].position;
  });

  const mouvements = [];

  donnees.slice(1).forEach((ligne, i) => {
    if (!ligne[idxA.NOM]) return;

    const numeroLigne = i + 2; // ligne 1 = en-tête, i=0 -> ligne 2

    // ---- ARRIVÉE : uniquement DateArrivée/HeureArrivée/ModeArrivée ----
    const dateArrivee = ligne[idxA.DATE_ARRIVEE];
    if (dateArrivee instanceof Date && datesAutorisees.has(formatDateCle(dateArrivee))) {
      lignesVerifiees.push({ onglet: Global.ONGLET_ARRIVEES, ligne: numeroLigne });

      const heureArrivee = ligne[idxA.HEURE_ARRIVEE];
      const heureRenseignee = heureArrivee instanceof Date;
      const heureEvenementDate = heureRenseignee ? combinerDateEtHeure(dateArrivee, heureArrivee) : null;
      const modeLibre = ligne[idxA.MODE_ARRIVEE] ? ligne[idxA.MODE_ARRIVEE].toString().trim() : "";
      const station = extraireStationDepuisModeLibre(modeLibre, codesConnus);

      // Mode renseigné mais abréviation absente de Paramètres — vide n'est
      // jamais une erreur (trajet pas encore saisi), une valeur non reconnue si.
      if (modeLibre && !codesConnus.has(station)) {
        erreurs.push({
          onglet: Global.ONGLET_ARRIVEES,
          ligne: numeroLigne,
          nom: ligne[idxA.NOM],
          prenom: ligne[idxA.PRENOM],
          modeBrut: modeLibre
        });
      }

      // Règle Saint-Roch/PPM : évaluable seulement si l'heure est connue ;
      // sinon la personne reste visible, comme pour tout lieu non reconnu.
      const exclue = heureEvenementDate && estExcluArrivee(station, heureEvenementDate, ligne[idxA.FONCTION]);
      if (!exclue) {
        let heurePickup = "";
        let lieuPickup = "";
        let dureeOccupationChauffeur = 0;

        if (heureRenseignee && station && tableLieux.hasOwnProperty(station)) {
          const infosLieu = tableLieux[station];
          heurePickup = formatHeureAffichage(new Date(heureEvenementDate.getTime() - infosLieu.delaiArrivee));
          lieuPickup = modeLibre;
          dureeOccupationChauffeur = infosLieu.dureeOccupation; // interne, jamais affichée
        }

        mouvements.push({
          date: dateArrivee,
          nom: ligne[idxA.NOM],
          prenom: ligne[idxA.PRENOM],
          fonction: ligne[idxA.FONCTION],
          telephone: ligne[idxA.TELEPHONE],
          heurePickup: heurePickup,
          lieuPickup: lieuPickup,
          lieuDepose: ligne[idxA.HOTEL],
          heureEvenement: heureEvenementDate ? formatHeureAffichage(heureEvenementDate) : "",
          dureeOccupationChauffeur: dureeOccupationChauffeur, // interne uniquement, jamais écrite dans un onglet
          origine: "Arrivée"
        });
      }
    }

    // ---- DÉPART : uniquement DateDépart/HeureDépart/ModeDépart ----
    const dateDepart = ligne[idxD.DATE_DEPART];
    if (dateDepart instanceof Date && datesAutorisees.has(formatDateCle(dateDepart))) {
      lignesVerifiees.push({ onglet: Global.ONGLET_ARRIVEES, ligne: numeroLigne });

      const heureDepart = ligne[idxD.HEURE_DEPART];
      const heureRenseignee = heureDepart instanceof Date;
      const heureEvenementDate = heureRenseignee ? combinerDateEtHeure(dateDepart, heureDepart) : null;
      const modeLibre = ligne[idxD.MODE_DEPART] ? ligne[idxD.MODE_DEPART].toString().trim() : "";
      const station = extraireStationDepuisModeLibre(modeLibre, codesConnus);

      if (modeLibre && !codesConnus.has(station)) {
        erreurs.push({
          onglet: Global.ONGLET_ARRIVEES,
          ligne: numeroLigne,
          nom: ligne[idxD.NOM],
          prenom: ligne[idxD.PRENOM],
          modeBrut: modeLibre
        });
      }

      const exclue = heureEvenementDate && estExcluDepart(station, heureEvenementDate, ligne[idxD.FONCTION]);
      if (!exclue) {
        let heurePickup = "";
        let lieuDepose = "";
        let dureeOccupationChauffeur = 0;

        if (heureRenseignee && station && tableLieux.hasOwnProperty(station)) {
          const infosLieu = tableLieux[station];
          heurePickup = formatHeureAffichage(new Date(heureEvenementDate.getTime() - infosLieu.delaiDepart));
          lieuDepose = modeLibre;
          dureeOccupationChauffeur = infosLieu.dureeOccupation; // interne, jamais affichée
        }

        mouvements.push({
          date: dateDepart,
          nom: ligne[idxD.NOM],
          prenom: ligne[idxD.PRENOM],
          fonction: ligne[idxD.FONCTION],
          telephone: ligne[idxD.TELEPHONE],
          heurePickup: heurePickup,
          lieuPickup: ligne[idxD.HOTEL],
          lieuDepose: lieuDepose,
          heureEvenement: heureEvenementDate ? formatHeureAffichage(heureEvenementDate) : "",
          dureeOccupationChauffeur: dureeOccupationChauffeur, // interne uniquement, jamais écrite dans un onglet
          origine: "Départ"
        });
      }
    }
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
 * Mise en forme "tableau propre" d'un onglet journalier : en-tête TOUJOURS
 * bleu avec texte blanc, lignes de données centrées avec bordures fines,
 * colorées manuellement par type (colorerLignesParType), colonnes ajustées à
 * la largeur du contenu, première ligne figée. Aucun thème de bandes
 * automatique (applyRowBanding) : il écraserait l'en-tête bleu/blanc juste
 * défini et empêcherait la coloration par type. Rejouable à tout moment sans
 * dégrader le rendu.
 * @param {Sheet} feuille
 */
function formaterOnglet(feuille) {
  const derniereLigne = feuille.getLastRow();
  const derniereColonne = feuille.getLastColumn();
  if (derniereLigne < 1 || derniereColonne < 1) return;

  // En-tête : toujours bleu avec texte blanc, quoi qu'il arrive ensuite.
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

    // Retire toute bande automatique existante : elle écraserait l'en-tête
    // bleu/blanc et empêcherait la coloration manuelle par type ci-dessous.
    feuille.getBandings().forEach(bande => bande.remove());

    colorerLignesParType(feuille);
  }

  // Force le format Heure (HH:mm) sur les colonnes calculées, au cas où la
  // colonne aurait hérité d'un format Durée ([h]:mm:ss) depuis le modèle.
  const enTetes = feuille.getRange(1, 1, 1, derniereColonne).getValues()[0]
    .map(valeur => (valeur || "").toString().trim());
  const idxHeurePickup = enTetes.indexOf(Global.COLONNES_PLANNING.HEURE_PICKUP) + 1;
  const idxHeureDepart = enTetes.indexOf(Global.COLONNES_PLANNING.HEURE_DEPART) + 1;

  if (derniereLigne > 1) {
    if (idxHeurePickup > 0) {
      feuille.getRange(2, idxHeurePickup, derniereLigne - 1, 1).setNumberFormat("HH:mm");
    }
    if (idxHeureDepart > 0) {
      feuille.getRange(2, idxHeureDepart, derniereLigne - 1, 1).setNumberFormat("HH:mm");
    }
  }

  feuille.autoResizeColumns(1, derniereColonne);
}

/**
 * Colore chaque ligne de données selon la colonne "Arrivée/Départ" : gris
 * clair pour un Départ, blanc pour une Arrivée — pour les distinguer d'un
 * coup d'œil maintenant que le thème de bandes automatique est retiré
 * (voir formaterOnglet). Ne colore rien si la colonne est introuvable
 * (faute de frappe dans l'en-tête de Planning Source) plutôt que de planter.
 * @param {Sheet} feuille
 */
function colorerLignesParType(feuille) {
  const derniereLigne = feuille.getLastRow();
  const derniereColonne = feuille.getLastColumn();
  if (derniereLigne < 2) return;

  const enTetes = feuille.getRange(1, 1, 1, derniereColonne).getValues()[0]
    .map(valeur => (valeur || "").toString().trim());
  const idxOrigine = enTetes.indexOf(Global.COLONNES_PLANNING.ORIGINE) + 1;
  if (idxOrigine === 0) return;

  const COULEUR_DEPART = "#f3f3f3";
  const COULEUR_ARRIVEE = "#ffffff";

  const valeurs = feuille.getRange(2, idxOrigine, derniereLigne - 1, 1).getValues();

  valeurs.forEach((ligne, i) => {
    const origine = (ligne[0] || "").toString().trim();
    const couleur = origine === "Départ" ? COULEUR_DEPART : COULEUR_ARRIVEE;
    feuille.getRange(2 + i, 1, 1, derniereColonne).setBackground(couleur);
  });
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
 * Pays, Langue). Écrit aussi "Arrivée/Départ" (si la colonne existe dans le
 * modèle) selon que le mouvement vient des colonnes arrivée ou départ.
 * @param {string} nomOnglet
 * @param {Array<Object>} mouvements
 */
function ecrireMouvementsDansOnglet(nomOnglet, mouvements) {
  const feuille = assurerOngletExiste(nomOnglet);

  const derniereColonne = feuille.getLastColumn();
  const enTetes = feuille.getRange(1, 1, 1, derniereColonne).getValues()[0]
    .map(valeur => (valeur || "").toString().trim());

  const idxNom         = enTetes.indexOf(Global.COLONNES_PLANNING.NOM) + 1;
  const idxPrenom      = enTetes.indexOf(Global.COLONNES_PLANNING.PRENOM) + 1;
  const idxFonction    = enTetes.indexOf(Global.COLONNES_PLANNING.FONCTION) + 1;
  const idxTelephone   = enTetes.indexOf(Global.COLONNES_PLANNING.TELEPHONE) + 1;
  const idxHeurePickup = enTetes.indexOf(Global.COLONNES_PLANNING.HEURE_PICKUP) + 1;
  const idxLieuPickup  = enTetes.indexOf(Global.COLONNES_PLANNING.LIEU_PICKUP) + 1;
  const idxLieuDepose  = enTetes.indexOf(Global.COLONNES_PLANNING.LIEU_DEPOSE) + 1;
  const idxHeureDepart = enTetes.indexOf(Global.COLONNES_PLANNING.HEURE_DEPART) + 1;
  const idxOrigine     = enTetes.indexOf(Global.COLONNES_PLANNING.ORIGINE) + 1;

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
    ligne[idxOrigine - 1]     = mvt.origine;
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
 * Orchestration : vérifie d'abord les en-têtes d'ARRIVEES (arrête tout avec
 * une alerte claire en cas d'écart), colore ARRIVEES par blocs de dates, lit
 * la table Paramètres unifiée une seule fois, en déduit les codes connus,
 * collecte les mouvements (arrivées et départs, les deux depuis ARRIVEES, en
 * un seul passage via collecterMouvements), puis écrit (ajout uniquement)
 * dans les onglets concernés. Un onglet n'est créé (copie de
 * Planning Source) que pour une date ayant au moins une personne.
 * @param {Date[]} [datesCiblees]
 */
function genererPlannings(datesCiblees) {
  if (!verifierEnTetesSources()) return; // écart détecté, alerte déjà affichée

  colorerBlocsDeDatesDesSources(); // mise en forme des sources à chaque génération

  const dates = (datesCiblees && datesCiblees.length) ? datesCiblees : obtenirDatesDisponibles();

  const tableLieux = obtenirTableLieux();
  const codesConnus = obtenirCodesStationConnus(tableLieux);
  const clesAutorisees = new Set(dates.map(formatDateCle));

  const erreurs = [];
  const lignesVerifiees = [];
  const mouvements = collecterMouvements(tableLieux, clesAutorisees, codesConnus, erreurs, lignesVerifiees);

  marquerLignesEnErreur(lignesVerifiees, erreurs); // rouge sur les lignes en erreur, remis en noir sinon

  const parOnglet = {};
  mouvements.forEach(mvt => {
    const nomOnglet = nomOngletPourDate(mvt.date);
    if (!parOnglet[nomOnglet]) parOnglet[nomOnglet] = [];
    parOnglet[nomOnglet].push(mvt);
  });

  Object.keys(parOnglet).forEach(nomOnglet => ecrireMouvementsDansOnglet(nomOnglet, parOnglet[nomOnglet]));

  if (erreurs.length > 0) afficherErreursDetection(erreurs);
}

/**
 * Remet en noir toutes les lignes effectivement revérifiées cette génération
 * (lignesVerifiees), puis repasse en rouge celles listées dans erreurs
 * (abréviation non reconnue) — pour les repérer d'un coup d'œil, et qu'elles
 * redeviennent noires dès que corrigées et régénérées. Ne touche pas aux
 * lignes hors du périmètre de cette génération (dates non ciblées lors d'une
 * génération partielle), pour ne jamais effacer à tort le signalement d'une
 * ligne pas encore revérifiée. Un seul onglet désormais (ARRIVEES) : une
 * ligne peut être revérifiée côté arrivée et/ou côté départ, tout finit sur
 * la même ligne physique.
 * @param {Array<{onglet:string, ligne:number}>} lignesVerifiees
 * @param {Array<{onglet:string, ligne:number}>} erreurs
 */
function marquerLignesEnErreur(lignesVerifiees, erreurs) {
  const feuille = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_ARRIVEES);
  const derniereColonne = feuille.getLastColumn();
  if (derniereColonne < 1) return;

  lignesVerifiees.forEach(l => feuille.getRange(l.ligne, 1, 1, derniereColonne).setFontColor("#000000"));
  erreurs.forEach(e => feuille.getRange(e.ligne, 1, 1, derniereColonne).setFontColor("#ff0000"));
}

/**
 * Affiche une alerte listant toutes les personnes dont l'abréviation de lieu
 * saisie (ModeArrivée/ModeDépart) ne correspond à aucune entrée de
 * Paramètres. Les lignes concernées sont déjà mises en rouge dans ARRIVEES
 * par marquerLignesEnErreur() au moment de l'appel.
 * @param {Array<{onglet:string, nom:string, prenom:string, modeBrut:string}>} erreurs
 */
function afficherErreursDetection(erreurs) {
  const details = erreurs.map(e =>
    "• " + e.onglet + " : " + e.nom + " " + e.prenom + " — abréviation \"" + e.modeBrut + "\" non reconnue dans Paramètres"
  ).join("\n");

  SpreadsheetApp.getUi().alert(
    erreurs.length + " abréviation(s) non reconnue(s)",
    "Ces personnes ont été incluses avec Heure Pick up vide, faute d'abréviation reconnue dans " +
    "Paramètres. Corrige la saisie ou ajoute le code dans Paramètres puis régénère — les lignes " +
    "concernées sont surlignées en rouge dans ARRIVEES :\n\n" + details,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
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

/**
 * Outil de diagnostic manuel (menu → "Diagnostic détection des lieux") : logue
 * la table Paramètres telle que lue par le script, alerte si elle est vide, et
 * pour les 10 premières lignes d'ARRIVEES logue le mode brut saisi côté
 * arrivée ET départ, le code détecté, et si ce code matche bien une entrée de
 * Paramètres. À utiliser quand Heure Pick up/Lieu Pick up restent vides pour
 * des lignes qui semblent pourtant avoir un code valide.
 */
function diagnostiquerDetectionLieux() {
  Logger.log("=== Vérification des en-têtes ARRIVEES ===");
  const erreursEnTetes = verifierEnTetesOnglet(Global.ONGLET_ARRIVEES, Global.COLONNES_ARRIVEES)
    .concat(verifierEnTetesOnglet(Global.ONGLET_ARRIVEES, Global.COLONNES_DEPARTS));
  if (erreursEnTetes.length === 0) {
    Logger.log("OK, toutes les colonnes attendues sont à la bonne position.");
  } else {
    erreursEnTetes.forEach(e => Logger.log("⚠️ " + e.onglet + ", colonne " + e.colonne +
      " : attendu \"" + e.attendu + "\", trouvé \"" + e.trouve + "\""));
  }

  const tableLieux = obtenirTableLieux();
  Logger.log("=== Table Paramètres lue ===");
  Logger.log(JSON.stringify(tableLieux, null, 2));
  if (Object.keys(tableLieux).length === 0) {
    Logger.log("⚠️ ALERTE : la table est vide. Vérifie que l'onglet 'Paramètres' existe, " +
      "que les colonnes C/D contiennent des HEURES (pas du texte), et qu'au moins une ligne " +
      "a une valeur non vide en C ou D.");
  }

  const codesConnus = obtenirCodesStationConnus(tableLieux);

  const feuille = SpreadsheetApp.getActive().getSheetByName(Global.ONGLET_ARRIVEES);
  const donnees = feuille.getDataRange().getValues();
  const idxNom = Global.COLONNES_SOURCE.NOM.position;
  const idxModeArrivee = Global.COLONNES_ARRIVEES.MODE_ARRIVEE.position;
  const idxModeDepart = Global.COLONNES_DEPARTS.MODE_DEPART.position;

  Logger.log("=== " + Global.ONGLET_ARRIVEES + " ===");
  donnees.slice(1, 11).forEach(ligne => { // 10 premières lignes à titre d'échantillon
    if (!ligne[idxNom]) return;

    const modeArrivee = ligne[idxModeArrivee] ? ligne[idxModeArrivee].toString().trim() : "";
    const stationArrivee = extraireStationDepuisModeLibre(modeArrivee, codesConnus);
    const matcheArrivee = stationArrivee && tableLieux.hasOwnProperty(stationArrivee);
    Logger.log(ligne[idxNom] + " | ARRIVÉE mode brut: \"" + modeArrivee + "\" -> détecté: \"" + stationArrivee +
      "\" -> " + (matcheArrivee ? "OK, trouvé dans Paramètres" : "PAS DE MATCH"));

    const modeDepart = ligne[idxModeDepart] ? ligne[idxModeDepart].toString().trim() : "";
    const stationDepart = extraireStationDepuisModeLibre(modeDepart, codesConnus);
    const matcheDepart = stationDepart && tableLieux.hasOwnProperty(stationDepart);
    Logger.log(ligne[idxNom] + " | DÉPART mode brut: \"" + modeDepart + "\" -> détecté: \"" + stationDepart +
      "\" -> " + (matcheDepart ? "OK, trouvé dans Paramètres" : "PAS DE MATCH"));
  });
}
