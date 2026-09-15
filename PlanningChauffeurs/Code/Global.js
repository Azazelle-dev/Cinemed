/**
 * Global.gs — Constantes globales du projet.
 * Aucune logique ici, uniquement des valeurs de configuration.
 */

const Global = {

  ONGLET_PARAMETRES: "Paramètres",
  ONGLET_ARRIVEES: "ARRIVEES",
  ONGLET_DEPARTS: "DEPARTS",
  ONGLET_PLANNING_SOURCE: "Planning Source",

  // Colonnes utiles de la liste des arrivées (A, B, D, E, F, G + L — on ignore Fonction en C)
  COLONNES_ARRIVEES: {
    NOM: "Nom",
    PRENOM: "Prénom",
    TELEPHONE: "Téléphone",
    DATE_ARRIVEE: "DateArrivée",
    HEURE_ARRIVEE: "HeureArrivée",
    MODE_ARRIVEE: "ModeArrivée",
    HOTEL: "NomHotel"
  },

  // Colonnes utiles de la liste des départs (A, B, D + H,I,J + L — on ignore Fonction en C)
  COLONNES_DEPARTS: {
    NOM: "Nom",
    PRENOM: "Prénom",
    TELEPHONE: "Téléphone",
    DATE_DEPART: "DateDépart",
    HEURE_DEPART: "HeureDépart",
    MODE_DEPART: "ModeDépart",
    HOTEL: "NomHotel"
  },

  JOURS_FR: ["DIMANCHE", "LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI"]
};
