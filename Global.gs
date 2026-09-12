/**
 * Global.gs — Constantes globales du projet.
 * Aucune logique ici, uniquement des valeurs de configuration.
 */

const Global = {

  // Table de correspondance trajet → délai (en ms). Sert aussi de liste blanche :
  // une valeur de ModeArrivée/ModeDépart n'est "confirmée" que si elle est une clé de cet objet.
  DELAIS_TRAJET: {
    "Corum>SDF":     1 * 60 * 60 * 1000, // 1h
    "Corum>St Roch": 1 * 60 * 60 * 1000, // 1h
    "Corum>MPL":     2 * 60 * 60 * 1000, // 2h
    "Corum>MRS":     4 * 60 * 60 * 1000  // 4h
  },

  ONGLET_PARAMETRES: "Paramètres",
  ONGLET_ARRIVEES: "ARRIVEES",

  // Colonne (index 1-based) de l'onglet Paramètres contenant les dates disponibles (colonne D)
  PARAMETRES_COLONNE_DATES: 4,

  // Noms des colonnes de la liste maîtresse ARRIVEES
  COLONNES_MASTER: {
    NOM:           "Nom",
    PRENOM:        "Prénom",
    TELEPHONE:     "Téléphone",
    DATE_ARRIVEE:  "DateArrivée",
    HEURE_ARRIVEE: "HeureArrivée",
    MODE_ARRIVEE:  "ModeArrivée",
    DATE_DEPART:   "DateDépart",
    HEURE_DEPART:  "HeureDépart",
    MODE_DEPART:   "ModeDépart"
  },

  // En-têtes du planning journalier généré, dans l'ordre exact demandé
  ENTETES_PLANNING_JOUR: ["Chauffeur", "Nom", "Prénom", "Heure de pick up", "Arrivée ou départ", "Trajet", "Téléphone"],

  JOURS_FR: ["DIMANCHE", "LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI"]
};
