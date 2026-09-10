/**
 * Global.gs — Constantes globales du projet.
 * Aucune logique ici, uniquement des valeurs de configuration.
 */

const Global = {

  // Table de correspondance trajet → délai (en millisecondes)
  DELAIS_TRAJET: {
    "Corum>SDF":     1 * 60 * 60 * 1000, // 1h
    "Corum>St Roch": 1 * 60 * 60 * 1000, // 1h
    "Corum>MPL":     2 * 60 * 60 * 1000, // 2h
    "Corum>MRS":     4 * 60 * 60 * 1000  // 4h
  },

  // Noms des colonnes attendues dans chaque onglet journalier
  COLONNES: {
    TRAJET:       "Trajet",
    DATE_DEPART:  "DateDépart",
    HEURE_DEPART: "HeureDépart",
    HEURE_PICKUP: "Heure Pick up"
  },

  // Nom de l'onglet contenant la table des délais et les dates disponibles
  ONGLET_PARAMETRES: "Paramètres",

  // Colonne (index 1-based) de l'onglet Paramètres contenant les dates disponibles (colonne D)
  PARAMETRES_COLONNE_DATES: 4,

  // Motif attendu pour le nom des onglets journaliers, ex: "VENDREDI 17"
  MOTIF_NOM_ONGLET_JOURNALIER: /^[A-ZÀ-Ý]+ \d{1,2}$/
};

// Jours de la semaine en français, indexés comme Date.getDay() (0 = dimanche)
const JOURS_FR = ["DIMANCHE", "LUNDI", "MARDI", "MERCREDI", "JEUDI", "VENDREDI", "SAMEDI"];
