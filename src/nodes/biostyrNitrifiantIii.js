// SQUELETTE — port de F1_Biostyr_N_III.cls à réaliser (source VBA : vba-source/F1_Biostyr_N_III.cls).
// Les paramètres ci-dessous sont les cellules de valeurs forcées (CelVF_*) du bloc
// correspondant de l'onglet Calculs ; le calcul laisse l'eau inchangée tant que
// compute() n'est pas porté.
import { passthrough } from './_stub.js'

export default passthrough({
  id: "biostyr-nitrifiant-iii",
  label: "Biostyr nitrifiant (tertiaire)",
  short: "Biostyr N III",
  family: "tertiaire",
  vba: "F1_Biostyr_N_III.cls",
  description: "Biofiltre BIOSTYR nitrifiant en étage tertiaire.",
  params: [
    { key: "hauteur_media", label: "Hauteur de média", unit: "m", group: 'Valeurs forcées', default: undefined },
    { key: "surface_filtration_min", label: "Surface de filtration minimale", unit: "m²", group: 'Valeurs forcées', default: undefined },
    { key: "nb_cellules", label: "Nombre TOTAL de cellules de filtration", unit: "-", group: 'Valeurs forcées', default: undefined },
    { key: "surface_filtration", label: "Surface réelle de filtration (calculée avec formats standards)", unit: "m²", group: 'Valeurs forcées', default: undefined },
    { key: "Cv_DCO", label: "Charge volumique DCO appliquée", unit: "kg/(m³.j)", group: 'Valeurs forcées', default: undefined },
    { key: "Cv_nit", label: "Charge volumique Nnit appliquée", unit: "kg/(m³.j)", group: 'Valeurs forcées', default: undefined },
    { key: "sortie_DCO", label: "Concentration DCO sortie", unit: "mg/L", group: 'Valeurs forcées', default: undefined },
    { key: "sortie_NH4", label: "Concentration N-NH4+ sortie", unit: "mg/L", group: 'Valeurs forcées', default: undefined },
    { key: "diffuseur_encrassement", label: "   Durée depuis l'installation / le dernier nettoyage des diffuseurs", unit: "an(s)", group: 'Valeurs forcées', default: undefined },
    { key: "air_P_refoulement", label: "Pression de refoulement des surpresseurs", unit: "mCE", group: 'Valeurs forcées', default: undefined },
    { key: "O2_besoin", label: "Besoin total en oxygène", unit: "kg O2/j", group: 'Valeurs forcées', default: undefined },
    { key: "O2_rdt_transfert", label: "Rendement de transfert de l'oxygène", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "air_Q_Nm3j", label: "Débit d'air process refoulé", unit: "Nm³/j", group: 'Valeurs forcées', default: undefined },
    { key: "surpresseur_conso_spec", label: "Consommation électrique spécifique des surpresseurs", unit: "Wh/(Nm³.mCE)", group: 'Valeurs forcées', default: undefined },
    { key: "ratio_ES_volume_media", label: "Ratio volume Eaux Sales / volume de média", unit: "-", group: 'Valeurs forcées', default: undefined },
    { key: "ES_nb_pompe", label: "Nombre de pompes", unit: "u", group: 'Valeurs forcées', default: undefined },
    { key: "ES_P_refoulement", label: "Pression de refoulement", unit: "mCE", group: 'Valeurs forcées', default: undefined },
    { key: "ES_tps_fonctionnement", label: "Durée de fonctionnement", unit: "h/j", group: 'Valeurs forcées', default: undefined },
    { key: "ES_Q_unitaire", label: "Débit de pompage unitaire", unit: "m³/h", group: 'Valeurs forcées', default: undefined },
    { key: "ES_pompe_rdt", label: "Rendement global pompes d'extraction des eaux sales (machine+moteur)", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "lavage_air_Q", label: "Débit d'air de lavage", unit: "Nm³/h", group: 'Valeurs forcées', default: undefined },
    { key: "lavage_air_P_refoulement", label: "Pression de refoulement", unit: "mCE", group: 'Valeurs forcées', default: undefined },
    { key: "lavage_air_tps_fonctionnement", label: "Durée de fonctionnement par cellule", unit: "h/j", group: 'Valeurs forcées', default: undefined },
    { key: "ES_concentration", label: "Concentration des Eaux Sales", unit: "mg/L", group: 'Valeurs forcées', default: undefined },
    { key: "ES_MV_MES", label: "Pourcentage MV/MES", unit: "%", group: 'Valeurs forcées', default: undefined },
  ],
})
