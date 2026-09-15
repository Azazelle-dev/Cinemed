/**
 * Global.gs — Constantes globales du projet.
 * Aucune logique ici, uniquement des valeurs de configuration.
 */

const Global = {

  ONGLET_PARAMETRES: "Paramètres",
  ONGLET_ARRIVEES: "ARRIVEES",
  ONGLET_DEPARTS: "DEPARTS",
  ONGLET_PLANNING_SOURCE: "Planning Source",

  // Colonnes utiles de la liste des arrivées : A à G + L
  COLONNES_ARRIVEES: {
    NOM: "Nom",
    PRENOM: "Prénom",
    FONCTION: "Fonction",
    TELEPHONE: "Téléphone",
    DATE_ARRIVEE: "DateArrivée",
    HEURE_ARRIVEE: "HeureArrivée",
    MODE_ARRIVEE: "ModeArrivée",
    HOTEL: "NomHotel"
  },

  // Colonnes utiles de la liste des départs : A, B, C, D, H, I, J, L
  COLONNES_DEPARTS: {
    NOM: "Nom",
    PRENOM: "Prénom",
    FONCTION: "Fonction",
    TELEPHONE: "Téléphone",
    DATE_DEPART: "DateDépart",
    HEURE_DEPART: "HeureDépart",
    MODE_DEPART: "ModeDépart",
    HOTEL: "NomHotel"
  },

  JOURS_FR: ["DIMANCHE", "LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI"]
};
