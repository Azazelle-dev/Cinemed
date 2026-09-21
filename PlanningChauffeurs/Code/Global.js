/**
 * Global.gs — Constantes globales du projet.
 * Aucune logique ici, uniquement des valeurs de configuration.
 */

// En-têtes réels de la ligne 1 d'ARRIVEES/DEPARTS (mêmes colonnes des deux
// côtés) — source de vérité unique des noms d'en-tête. ARRIVEES/DEPARTS et
// Common.gs piochent seulement les colonnes dont ils ont besoin ci-dessous,
// jamais de chaîne d'en-tête écrite en dur ailleurs dans le code.
const COLONNES_SOURCE = {
  NOM: "Nom",
  PRENOM: "Prénom",
  FONCTION: "Fonction",
  TELEPHONE: "Téléphone",
  DATE_ARRIVEE: "DateArrivée",
  HEURE_ARRIVEE: "HeureArrivée",
  MODE_ARRIVEE: "ModeArrivée",
  DATE_DEPART: "DateDépart",
  HEURE_DEPART: "HeureDépart",
  MODE_DEPART: "ModeDépart",
  REMARQUES: "Remarques",
  HOTEL: "NomHotel",
  CHECK_IN: "CheckIn",
  CHECK_OUT: "CheckOut",
  NUITS_EFFECTIVES: "NuitsEffectives",
  NUITS_PEN_CHARGE: "NuitsPenCharge",
  CHAMBRE: "Chambre",
  NOTE2: "Note2",
  COLONNE_1: "Colonne 1",
  COLONNE_2: "Colonne 2",
  COLONNE_3: "Colonne 3",
  COLONNE_4: "Colonne 4"
};

// En-têtes de l'onglet modèle Planning Source.
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
  ONGLET_DEPARTS: "DEPARTS",
  ONGLET_PLANNING_SOURCE: "Planning Source",

  COLONNES_SOURCE: COLONNES_SOURCE,
  COLONNES_PLANNING: COLONNES_PLANNING,

  // Colonnes utiles de la liste des arrivées : A à G + L
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

  // Colonnes utiles de la liste des départs : A, B, C, D, H, I, J, L
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
