/**
 * Global.gs — Constantes globales du projet.
 * Aucune logique ici, uniquement des valeurs de configuration.
 */

// En-têtes réels ET position fixe (0-based, "A" = 0) de la ligne 1
// d'ARRIVEES — source de vérité unique. ARRIVEES contient à la fois les
// colonnes d'arrivée (E-G) et de départ (H-J) pour une même personne sur une
// même ligne ; il n'y a plus d'onglet DEPARTS séparé. Le code lit ARRIVEES
// par position fixe (jamais par recherche dynamique du nom), donc
// verifierEnTetesSources() DOIT être appelée avant toute lecture pour
// garantir que la position réelle correspond encore au nom attendu (voir
// Common.gs).
const COLONNES_SOURCE = {
  NOM:              { nom: "Nom",             position: 0 },
  PRENOM:           { nom: "Prénom",          position: 1 },
  FONCTION:         { nom: "Fonction",        position: 2 },
  TELEPHONE:        { nom: "Téléphone",       position: 3 },
  DATE_ARRIVEE:     { nom: "DateArrivée",     position: 4 },
  HEURE_ARRIVEE:    { nom: "HeureArrivée",    position: 5 },
  MODE_ARRIVEE:     { nom: "ModeArrivée",     position: 6 },
  DATE_DEPART:      { nom: "DateDépart",      position: 7 },
  HEURE_DEPART:     { nom: "HeureDépart",     position: 8 },
  MODE_DEPART:      { nom: "ModeDépart",      position: 9 },
  REMARQUES:        { nom: "Remarques",       position: 10 },
  HOTEL:            { nom: "NomHotel",        position: 11 },
  CHECK_IN:         { nom: "CheckIn",         position: 12 },
  CHECK_OUT:        { nom: "CheckOut",        position: 13 },
  NUITS_EFFECTIVES: { nom: "NuitsEffectives", position: 14 },
  NUITS_PEN_CHARGE: { nom: "NuitsPenCharge",  position: 15 },
  CHAMBRE:          { nom: "Chambre",         position: 16 },
  NOTE2:            { nom: "Note2",           position: 17 },
  COLONNE_1:        { nom: "Colonne 1",       position: 18 },
  COLONNE_2:        { nom: "Colonne 2",       position: 19 },
  COLONNE_3:        { nom: "Colonne 3",       position: 20 },
  COLONNE_4:        { nom: "Colonne 4",       position: 21 }
};

// En-têtes de l'onglet modèle Planning Source. Contrairement à ARRIVEES/
// DEPARTS, ces colonnes restent recherchées par nom (enTetes.indexOf) et non
// par position fixe : Planning Source peut être réorganisé librement, seul
// l'intitulé de chaque en-tête compte (voir ecrireMouvementsDansOnglet).
const COLONNES_PLANNING = {
  NOM: "Nom",
  PRENOM: "Prénom",
  FONCTION: "Fonction",
  TELEPHONE: "Téléphone",
  HEURE_PICKUP: "Heure Pick up",
  LIEU_PICKUP: "Lieu Pick up",
  LIEU_DEPOSE: "Lieu de dépose",
  HEURE_DEPART: "Heure départ",
  ORIGINE: "Arrivée/Départ"
};

const Global = {

  ONGLET_PARAMETRES: "Paramètres",
  ONGLET_ARRIVEES: "ARRIVEES",
  ONGLET_PLANNING_SOURCE: "Planning Source",

  COLONNES_SOURCE: COLONNES_SOURCE,
  COLONNES_PLANNING: COLONNES_PLANNING,

  // Colonnes utiles de la liste des arrivées (chacune avec son nom ET sa position fixe)
  COLONNES_ARRIVEES: {
    NOM: COLONNES_SOURCE.NOM,
    PRENOM: COLONNES_SOURCE.PRENOM,
    FONCTION: COLONNES_SOURCE.FONCTION,
    TELEPHONE: COLONNES_SOURCE.TELEPHONE,
    DATE_ARRIVEE: COLONNES_SOURCE.DATE_ARRIVEE,
    HEURE_ARRIVEE: COLONNES_SOURCE.HEURE_ARRIVEE,
    MODE_ARRIVEE: COLONNES_SOURCE.MODE_ARRIVEE,
    HOTEL: COLONNES_SOURCE.HOTEL
  },

  // Colonnes utiles côté départ, dans ARRIVEES (même feuille, mêmes lignes que COLONNES_ARRIVEES)
  COLONNES_DEPARTS: {
    NOM: COLONNES_SOURCE.NOM,
    PRENOM: COLONNES_SOURCE.PRENOM,
    FONCTION: COLONNES_SOURCE.FONCTION,
    TELEPHONE: COLONNES_SOURCE.TELEPHONE,
    DATE_DEPART: COLONNES_SOURCE.DATE_DEPART,
    HEURE_DEPART: COLONNES_SOURCE.HEURE_DEPART,
    MODE_DEPART: COLONNES_SOURCE.MODE_DEPART,
    HOTEL: COLONNES_SOURCE.HOTEL
  },

  CODES_MOYEN_PROPRE: ["PPM"], // valeurs qui excluent totalement la personne du planning

  JOURS_FR: ["DIMANCHE", "LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI"]
};
