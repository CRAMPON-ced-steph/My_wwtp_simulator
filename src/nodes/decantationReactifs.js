// SQUELETTE — port de D2_Decantation_reactif.cls à réaliser (source VBA : vba-source/D2_Decantation_reactif.cls).
// Les paramètres ci-dessous sont les cellules de valeurs forcées (CelVF_*) du bloc
// correspondant de l'onglet Calculs ; le calcul laisse l'eau inchangée tant que
// compute() n'est pas porté.
import { passthrough } from './_stub.js'

export default passthrough({
  id: "decantation-reactifs",
  label: "Décantation avec réactifs",
  short: "Décanteur I réactifs",
  family: "primaire",
  vba: "D2_Decantation_reactif.cls",
  description: "Décantation primaire physico-chimique (MULTIFLO DUO/TRIO, ACTIFLO) : coagulation FeCl3, floculation polymère, option microsable et recirculation.",
  params: [
    { key: "bypass_calcul", label: "Eau by-passée au traitement primaire", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "dosage_FeCl3", label: "Dosage réactifs (FeCl3 pur)", unit: "mg/L", group: 'Valeurs forcées', default: undefined },
    { key: "dosage_polymere", label: "Dosage polymère", unit: "mg/L", group: 'Valeurs forcées', default: undefined },
    { key: "NombreOuvrages", label: "Nombre d'ouvrages", unit: "-", group: 'Valeurs forcées', default: undefined },
    { key: "Vmirroir_limite", label: "Vitesse au miroir limite (temps de pluie)", unit: "m/h", group: 'Valeurs forcées', default: undefined },
    { key: "Smirroir", label: "Surface au miroir", unit: "m²", group: 'Valeurs forcées', default: undefined },
    { key: "hauteur", label: "Hauteur du décanteur", unit: "m", group: 'Valeurs forcées', default: undefined },
    { key: "rdt_MES", label: "Rendement d'élimination MES", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "rdt_DCO", label: "Rendement d'élimination DCO", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "rdt_DBO", label: "Rendement d'élimination DBO5", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "rdt_NK", label: "Rendement d'élimination NK", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "rdt_P", label: "Rendement d'élimination P", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "MV_MES", label: "Pourcentage MV / MES", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "recirculation_taux", label: "Taux de recirculation", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "recirculation_nb_pompe", label: "Nombre de pompes", unit: "u", group: 'Valeurs forcées', default: undefined },
    { key: "recirculation_P_refoulement", label: "Pression de refoulement", unit: "mCE", group: 'Valeurs forcées', default: undefined },
    { key: "recirculation_tps_fonctionnement", label: "Durée de fonctionnement", unit: "h/j", group: 'Valeurs forcées', default: undefined },
    { key: "recirculation_Qunitaire", label: "Débit de pompage unitaire", unit: "m³/h", group: 'Valeurs forcées', default: undefined },
    { key: "recirculation_pompe_rdt", label: "Rendement global des pompes de recirculation (machine+moteur)", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "boues_concentration", label: "Concentration en MES des boues extraites", unit: "g/L", group: 'Valeurs forcées', default: undefined },
    { key: "boues_MES", label: "Boues à extraire", unit: "kg MES/j", group: 'Valeurs forcées', default: undefined },
    { key: "boues_Q", label: "m³/j", unit: "#REF!", group: 'Valeurs forcées', default: undefined },
    { key: "extraction_nb_pompe", label: "Nombre de pompes", unit: "u", group: 'Valeurs forcées', default: undefined },
    { key: "extraction_Prefoulement", label: "Pression de refoulement", unit: "mCE", group: 'Valeurs forcées', default: undefined },
    { key: "extraction_tps_fonctionnement", label: "Durée de fonctionnement", unit: "h/j", group: 'Valeurs forcées', default: undefined },
    { key: "extraction_Qunitaire", label: "Débit de pompage unitaire", unit: "m³/h", group: 'Valeurs forcées', default: undefined },
    { key: "extraction_pompe_rdt", label: "Rendement global des pompes d'extraction (machine+moteur)", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "electricite", label: "Consommation électrique du décanteur physico-chimique", unit: "kWh/j", group: 'Valeurs forcées', default: undefined },
  ],
})
