// SQUELETTE — port de E2_BA_moyenne_charge.cls à réaliser (source VBA : vba-source/E2_BA_moyenne_charge.cls).
// Les paramètres ci-dessous sont les cellules de valeurs forcées (CelVF_*) du bloc
// correspondant de l'onglet Calculs ; le calcul laisse l'eau inchangée tant que
// compute() n'est pas porté.
import { passthrough } from './_stub.js'

export default passthrough({
  id: "ba-moyenne-charge",
  label: "Boue activée moyenne charge",
  short: "BA moyenne charge",
  family: "secondaire",
  vba: "E2_BA_moyenne_charge.cls",
  description: "Boue activée moyenne charge : carbone + nitrification partielle, clarificateur.",
  params: [
    { key: "nominal_age_boue", label: "Age de boue", unit: "j", group: 'Valeurs forcées', default: undefined },
    { key: "nominal_MES_concentration_bassin", label: "Concentration en MES dans les bassins", unit: "g/L", group: 'Valeurs forcées', default: undefined },
    { key: "nominal_MV_MES", label: "Pourcentage MV/MES", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "volume_bassins", label: "Volume total des bassins", unit: "m³", group: 'Valeurs forcées', default: undefined },
    { key: "reel_age_boue", label: "Age de boue", unit: "j", group: 'Valeurs forcées', default: undefined },
    { key: "reel_MV_MES", label: "Pourcentage MV/MES", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "reel_MES_concentration_bassin", label: "Concentration en MES dans les bassins", unit: "g/L", group: 'Valeurs forcées', default: undefined },
    { key: "O2_dissous", label: "Concentration moyenne en O2 dissous dans les bassins en phase aérée", unit: "mg/L", group: 'Valeurs forcées', default: undefined },
    { key: "sortie_DBO", label: "Concentration en DBO5 en sortie", unit: "mg/L", group: 'Valeurs forcées', default: undefined },
    { key: "hauteur_bassin", label: "Hauteur d'eau du bassin", unit: "m", group: 'Valeurs forcées', default: undefined },
    { key: "O2_besoin", label: "Besoin total en oxygène", unit: "kg O2/j", group: 'Valeurs forcées', default: undefined },
    { key: "O2_facteur_alpha", label: "Facteur alpha (transfert O2 en eaux usées / transfert O2 en eau claire)", unit: "-", group: 'Valeurs forcées', default: undefined },
    { key: "O2_rdt_transfert", label: "Rendement de transfert de l'oxygène", unit: "%/m", group: 'Valeurs forcées', default: undefined },
    { key: "air_Q_Nm3j", label: "Débit d'air process refoulé", unit: "Nm³/j", group: 'Valeurs forcées', default: undefined },
    { key: "diffuseur_encrassement", label: "Durée depuis l'installation / le dernier nettoyage des diffuseurs", unit: "an(s)", group: 'Valeurs forcées', default: undefined },
    { key: "air_P_refoulement", label: "Pression de refoulement des surpresseurs", unit: "mCE", group: 'Valeurs forcées', default: undefined },
    { key: "surpresseur_conso_spec", label: "Consommation électrique spécifique des surpresseurs", unit: "Wh/(Nm³.mCE)", group: 'Valeurs forcées', default: undefined },
    { key: "ASB_eau_claire", label: "ASB des aérateurs de surface en eau claire (kg 02/kW)", unit: "kgO2/kWh", group: 'Valeurs forcées', default: undefined },
    { key: "recirculation_taux", label: "Taux de recirculation des boues du clarificateur", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "recirculation_P_refoulement", label: "Pression de refoulement de la recirculation des boues du clarificateur", unit: "mCE", group: 'Valeurs forcées', default: undefined },
    { key: "recirculation_pompe_rdt", label: "Rendement global des pompes de recirculation (machine+moteur)", unit: "%", group: 'Valeurs forcées', default: undefined },
    { key: "NombreClarificateurs", label: "Nombre d'ouvrages", unit: "-", group: 'Valeurs forcées', default: undefined },
    { key: "boues_indice_Mohlman", label: "Indice de boues", unit: "mL/g", group: 'Valeurs forcées', default: undefined },
    { key: "sortie_MES", label: "Concentration en MES de l'effluent", unit: "mg/L", group: 'Valeurs forcées', default: undefined },
    { key: "clarif_hauteur", label: "Hauteur du clarificateur", unit: "m", group: 'Valeurs forcées', default: undefined },
    { key: "clarif_vitesse_max", label: "Vitesse hydraulique maximale", unit: "m/h", group: 'Valeurs forcées', default: undefined },
    { key: "clarif_surface", label: "Surface de décantation", unit: "m²", group: 'Valeurs forcées', default: undefined },
    { key: "boues_concentration", label: "Concentration en MES des boues extraites", unit: "g/L", group: 'Valeurs forcées', default: undefined },
    { key: "boues_MES", label: "Boues à extraire", unit: "kg MES/j", group: 'Valeurs forcées', default: undefined },
    { key: "extraction_P_refoulement", label: "Pression de refoulement de l'extraction", unit: "mCE", group: 'Valeurs forcées', default: undefined },
    { key: "extraction_pompe_rdt", label: "Rendement global des pompes d'extraction (machine+moteur)", unit: "%", group: 'Valeurs forcées', default: undefined },
  ],
})
