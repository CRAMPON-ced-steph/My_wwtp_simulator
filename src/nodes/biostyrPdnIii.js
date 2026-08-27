// SQUELETTE — port de F2_Biostyr_PDN_III.cls à réaliser (source VBA : vba-source/F2_Biostyr_PDN_III.cls).
// Les paramètres ci-dessous sont les cellules de valeurs forcées (CelVF_*) du bloc
// correspondant de l'onglet Calculs ; le calcul laisse l'eau inchangée tant que
// compute() n'est pas porté.
import { passthrough } from './_stub.js'

export default passthrough({
  id: "biostyr-pdn-iii",
  label: "Biostyr PDN (tertiaire)",
  short: "Biostyr PDN III",
  family: "tertiaire",
  vba: "F2_Biostyr_PDN_III.cls",
  description: "Biofiltre BIOSTYR post-dénitrification (méthanol) en étage tertiaire.",
  params: [
    { key: "NO3_garantie", label: "Concentration en NO3 garantie en sortie PDN (eau traitée sur l'ouvrage)", unit: "mg/L", group: 'Valeurs forcées', default: undefined },
    { key: "NO3_visee", label: "Concentration en NO3 en sortie (eau traitée + eau by-passée)", unit: "mg/L", group: 'Valeurs forcées', default: undefined },
    { key: "inlet_ratio_admis", label: "Fraction du débit traité en PDN", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "hauteur_media", label: "Hauteur de média", unit: "m", group: 'Valeurs forcées', default: undefined },
    { key: "surface_filtration_min", label: "Surface de filtration minimale", unit: "m²", group: 'Valeurs forcées', default: undefined },
    { key: "nb_cellules", label: "Nombre TOTAL de cellules de filtration", unit: "-", group: 'Valeurs forcées', default: undefined },
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
    { key: "methanol_flux", label: "Consommation de méthanol", unit: "kg/j", group: 'Valeurs forcées', default: undefined },
  ],
})
