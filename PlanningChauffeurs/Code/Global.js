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

  JOURS_FR: ["DIMANCHE", "LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI"],

  // Génération automatique : régénère les plannings ce délai après la dernière
  // modification sur ARRIVEES/DEPARTS (aucune activité entre-temps).
  DELAI_INACTIVITE_MS: 15 * 1000,
  NOM_FONCTION_AUTO: "genererPlanningsAuto",
  NOM_FONCTION_ON_EDIT: "onEditInstallable",
  PROPRIETE_TRIGGER_AUTO: "triggerAutoGenerationId"
};
