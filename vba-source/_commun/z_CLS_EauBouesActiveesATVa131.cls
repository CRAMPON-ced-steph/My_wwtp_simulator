Option Explicit

'******************************************************************************************************************************************************
'Boues activées avec traitement de l'azote voire traitement du phosphore selon le modèle ATV A131
'   Ce procédé comporte:
'       - une entrée eau qui correspond aux usées à traiter
'       - une sortie eau qui correspond à l'eau traitée et clarifiée
'       - une sortie boues au niveau du clarificateur
'   Ce module calcule la consommation électrique journalière de la boues activées. Cela comprend les consommateurs suivants:
'       - l'agitation des bassins de boues activées
'       - le système d'aération
'       - le clarificateur (pont-râcleur)
'       - les pompes de recirculation de boues entre le clarificateur et les bassins de boues activées
'       - le cas échéant, les pompes de recirculation de la liqueur mixte pour la dénitrification
'   Ce module réalise les bilans matières sur le procédé, calcule les différents flux en sortie ainsi que les besoins en air.
'   Il effectue un pré-Privateensionnement des bassins de boues activées (volume) ainsi que du clarificateur (surface).
'   Il calcule les besoins en réactifs
'       - coagulant (chlorure ferrique)
'       - source externe de carbone (méthanol)
'******************************************************************************************************************************************************

'Critère de convergence de l'âge de boues
Private Const md_CRITERE_CONVERGENCE_AGE_BOUES As Double = 0.0001

'constantes pour le choix concernant la déphosphatation
Private Const mi_DEPHOSPHATATION_NOMBRE_CHOIX = 4  'nombre de choix possibles
Private Const mi_DEPHOSPHATATION_NON = 1
Private Const mi_DEPHOSPHATATION_BIOLOGIQUE = 2
Private Const mi_DEPHOSPHATATION_PRECIPITATION = 3
Private Const mi_DEPHOSPHATATION_BIOLOGIQUE_ET_PRECIPITATION = 4

'constantes pour le choix du type de bassin
Private Const mi_BASSINS_NOMBRE_CONFIGURATION = 2  'nombre de type de configurations de bassins
Private Const mi_BASSINS_SEPARES = 1
Private Const mi_BASSINS_CHENAUX = 2

'constantes pour le choix du type d'aérateur
Private Const mi_AERATEUR_NOMBRE_TYPE = 5  'nombre de type d'aérateurs
Private Const mi_AERATEUR_FINES_BULLES = 1
Private Const mi_AERATEUR_MOYENNES_BULLES = 2
Private Const mi_AERATEUR_BROSSES = 3
Private Const mi_AERATEUR_TURBINES_LENTES = 4
Private Const mi_AERATEUR_TURBINES_RAPIDES = 5

'constantes pour le choix du type de régulation d'air
Private Const mi_AERATION_REGULATION_NOMBRE_TYPE = 4  'nombre de types de régulation
Private Const mi_AERATION_REGULATION_HORLOGE = 1
Private Const mi_AERATION_REGULATION_CLASSIQUE_SANS_VARIATION_FREQUENCE = 2
Private Const mi_AERATION_REGULATION_CLASSIQUE_AVEC_VARIATION_FREQUENCE = 3
Private Const mi_AERATION_REGULATION_AVANCEE = 4

'constantes pour le choix du type de surpresseur
Private Const mi_SURPRESSEUR_NOMBRE_TECHNOLOGIE = 3
Private Const mi_SURPRESSEUR_LOBES = 1
Private Const mi_SURPRESSEUR_VIS = 2
Private Const mi_SURPRESSEUR_CENTRIFUGE = 3
    Private Const nb_surpresseur_gamme_Q = 1
    Private Const surpresseur_gamme_Q_0_infini = 1
    Private Const nb_surpresseur_gamme_P = 3
    Private Const surpresseur_gamme_P_0_Px1 = 1
    Private Const Px1 = 6    'mCE
    Private Const surpresseur_gamme_P_Px1_Px2 = 2
    Private Const Px2 = 8    'mCE
    Private Const surpresseur_gamme_P_Px2_infini = 3

'constantes pour le choix du type de râcleur
Private Const mi_RACLEUR_NOMBRE_TECHNOLOGIE = 3  'nombre de type de râcleurs
Private Const mi_RACLEUR_RACLE = 1
Private Const mi_RACLEUR_RACLE_SUCE = 2
Private Const mi_RACLEUR_KRUGER = 3

'variables de choix internes
Private mi_ChoixDephosphatation As Integer
Private mi_ChoixConfigurationBassins As Integer
Private mi_ChoixTypeAeration As Integer
Private mi_ChoixRegulationAeration As Integer
Private mi_ChoixTypeSurpresseur As Integer
Private mi_ChoixTypeRacleur As Integer

'variables échangées avec Excel
Private md_NominalAgeBoues_j As Double
Private md_NominalConcentrationBassinsMES_gL As Double
Public pd_NominalRatioMVMES As Double
Private md_VolumeBassinsTotal_m3 As Double
Private md_VolumeBassinAnaerobie_m3 As Double
Private md_VolumeBassinAnoxie_m3 As Double
Private md_VolumeBassinAerobie_m3 As Double
Private md_VolumeChenal_m3 As Double
Private md_PuissanceAgitationAnoxie_Wm3 As Double
Private md_PuissanceAgitationAerobie_Wm3 As Double
Private md_PuissanceAgitationAnaerobie_Wm3 As Double
Private md_PuissanceAgitationChenal_Wm3 As Double
Private md_ReelAgeBoues_j As Double
Public pd_ReelRatioMVMES As Double
Private md_ReelConcentrationBassinsMES_gL As Double
Private md_ConcentrationO2dissous_mgL As Double
Private md_ConcentrationSortieDBO_mgL As Double
Private md_ConcentrationSortieNH4_mgL As Double
Private md_ConcentrationSortieNO3_mgL As Double
Private md_RendementDephosphatationBiologique As Double
Private md_ConcentrationSortiePt_mgL As Double
Public pd_ConsommationMethanolPur_kgj As Double
Public pd_ConsommationFeCl3Pur_kgj As Double
Private md_HauteurBassinAeration_m As Double
Private md_BesoinsOxygene_kgO2j As Double
Private md_RendementTransfertOxygeneMembraneEauClaire_1m As Double
Private md_FacteurAlpha As Double
Private md_DebitAir_Nm3j As Double
Private md_DernierNettoyageOuRemplacamentMembranes As Double
Private md_PressionAirSurpresse_mCE As Double
Private md_ConsommationSpecifiqueSurpresseur_Wh_Nm3mCE As Double
Private md_ApportSpecifiqueBrutEauClaire_kgO2_kWh As Double
Private md_TauxRecirculationBouesClarificateur As Double
Private md_PompesRecirculationBouesPression_mCE As Double
Private md_PompesRecirculationBouesRendement As Double
Private md_TauxRecirculationMLSS As Double
Private md_PompesRecirculationMLSSPression_mCE As Double
Private md_PompesRecirculationMLSSRendement As Double
Private md_BouesISV_mL_g As Double
Private md_EauClarifieeMES_mgL As Double
Private md_HauteurClarificateur_m As Double
Private md_ClarificationVitesseMaximale_mh As Double
Private md_SurfaceTotaleClarification_m2 As Double
Private md_ClarificationVitesseMaximaleRecalculee_mh As Double
Private md_ClarificateurChargeRadierNominal As Double
Private md_BouesExtraitesConcentrationMES_gL As Double
Public pd_BouesExtraitesFluxMES_kgj As Double
Public pd_BouesExtraitesDebit_m3j As Double
Private md_PompesBouesExcesPression As Double
Private md_PompesBouesExcesRendement As Double
Public pd_ConsommationElectrique_kWhj As Double
Private md_ConsommationElectriqueAeration_kWhj As Double

Public P_precipite As Double
Public pd_ConsommationElectrique_kWhj_fixe As Double
Private besoins_O2_respiration As Double

'variables montrant si les valeurs ont été forcées
Private mn_NominalAgeBoues_j_force As Boolean
Private mn_NominalConcentrationBassinsMES_gL_force As Boolean
Private mn_NominalRatioMVMES_force As Boolean
Private mn_VolumeBassinsTotal_m3_force As Boolean
Private mn_VolumeBassinAnaerobie_m3_force As Boolean
Private mn_VolumeBassinAnoxie_m3_force As Boolean
Private mn_VolumeBassinAerobie_m3_force As Boolean
Private mn_VolumeChenal_m3_force As Boolean
Private mn_PuissanceAgitationAnoxie_Wm3_force As Boolean
Private mn_PuissanceAgitationAerobie_Wm3_force As Boolean
Private mn_PuissanceAgitationAnaerobie_Wm3_force As Boolean
Private mn_PuissanceAgitationChenal_Wm3_force As Boolean
Private mn_ReelAgeBoues_j_force As Boolean
Private mn_ReelRatioMVMES_force As Boolean
Private mn_ReelConcentrationBassinsMES_gL_force As Boolean
Private mn_ConcentrationO2dissous_mgL_force As Boolean
Private mn_ConcentrationSortieDBO_mgL_force As Boolean
Private mn_ConcentrationSortieNH4_mgL_force As Boolean
Private mn_ConcentrationSortieNO3_mgL_force As Boolean
Private mn_RendementDephosphatationBiologique_force As Boolean
Private mn_ConcentrationSortiePt_mgL_force As Boolean
Private mn_ConsommationMethanolPur_kgj_force As Boolean
Private mn_ConsommationFeCl3Pur_kgj_force As Boolean
Private mn_HauteurBassinAeration_m_force As Boolean
Private mn_BesoinsOxygene_kgO2j_force As Boolean
Private mn_RendementTransfertOxygeneMembraneEauClaire_1m_force As Boolean
Private mn_FacteurAlpha_force As Boolean
Private mn_DebitAir_Nm3j_force As Boolean
Private mn_DernierNettoyageOuRemplacamentMembranes_force As Boolean
Private mn_PressionAirSurpresse_mCE_force As Boolean
Private mn_ConsommationSpecifiqueSurpresseur_Wh_Nm3mCE_force As Boolean
Private mn_ApportSpecifiqueBrutEauClaire_kgO2_kWh_force As Boolean
Private mn_TauxRecirculationBouesClarificateur_force As Boolean
Private mn_PompesRecirculationBouesPression_mCE_force As Boolean
Private mn_PompesRecirculationBouesRendement_force As Boolean
Private mn_TauxRecirculationMLSS_force As Boolean
Private mn_PompesRecirculationMLSSPression_mCE_force As Boolean
Private mn_PompesRecirculationMLSSRendement_force As Boolean
Private mn_BouesISV_mL_g_force As Boolean
Private mn_EauClarifieeMES_mgL_force As Boolean
Private mn_HauteurClarificateur_m_force As Boolean
Private mn_ClarificationVitesseMaximale_mh_force As Boolean
Private mn_SurfaceTotaleClarification_m2_force As Boolean
Private mn_BouesExtraitesConcentrationMES_gL_force As Boolean
Private mn_BouesExtraitesFluxMES_kgj_force As Boolean
Private mn_BouesExtraitesDebit_m3j_force As Boolean
Private mn_PompesBouesExcesPression_force As Boolean
Private md_PompesBouesExcesRendement_force As Boolean

'valeurs forcées par l'utilisateur
Private md_NominalAgeBoues_j_f As Double
Private md_NominalConcentrationBassinsMES_gL_f   As Double
Private pd_NominalRatioMVMES_f As Double
Private md_VolumeBassinsTotal_m3_f As Double
Private md_VolumeBassinAnaerobie_m3_f As Double
Private md_VolumeBassinAnoxie_m3_f As Double
Private md_VolumeBassinAerobie_m3_f As Double
Private md_VolumeChenal_m3_f As Double
Private md_PuissanceAgitationAnoxie_Wm3_f As Double
Private md_PuissanceAgitationAerobie_Wm3_f As Double
Private md_PuissanceAgitationAnaerobie_Wm3_f As Double
Private md_PuissanceAgitationChenal_Wm3_f As Double
Private md_ReelAgeBoues_j_f As Double
Private md_ReelRatioMVMES_f As Double
Private md_ReelConcentrationBassinsMES_gL_f As Double
Private md_ConcentrationO2dissous_mgL_f As Double
Private md_ConcentrationSortieDBO_mgL_f As Double
Private md_ConcentrationSortieNH4_mgL_f As Double
Private md_ConcentrationSortieNO3_mgL_f As Double
Private md_RendementDephosphatationBiologique_f As Double
Private md_ConcentrationSortiePt_mgL_f As Double
Private md_ConsommationMethanolPur_kgj_f As Double
Private md_ConsommationFeCl3Pur_kgj_f As Double
Private md_HauteurBassinAeration_m_f As Double
Private md_BesoinsOxygene_kgO2j_f As Double
Private md_RendementTransfertOxygeneMembraneEauClaire_1m_f As Double
Private md_FacteurAlpha_f As Double
Private md_DebitAir_Nm3j_f As Double
Private md_DernierNettoyageOuRemplacamentMembranes_f As Double
Private md_PressionAirSurpresse_mCE_f As Double
Private md_ConsommationSpecifiqueSurpresseur_Wh_Nm3mCE_f As Double
Private md_ApportSpecifiqueBrutEauClaire_kgO2_kWh_f As Double
Private md_TauxRecirculationBouesClarificateur_f As Double
Private md_PompesRecirculationBouesPression_mCE_f As Double
Private md_PompesRecirculationBouesRendement_f As Double
Private md_TauxRecirculationMLSS_f As Double
Private md_PompesRecirculationMLSSPression_mCE_f As Double
Private md_PompesRecirculationMLSSRendement_f As Double
Private md_BouesISV_mL_g_f As Double
Private md_EauClarifieeMES_mgL_f As Double
Private md_HauteurClarificateur_m_f As Double
Private md_ClarificationVitesseMaximale_mh_f As Double
Private md_SurfaceTotaleClarification_m2_f As Double
Private md_BouesExtraitesConcentrationMES_gL_f As Double
Private pd_BouesExtraitesFluxMES_kgj_f As Double
Private pd_BouesExtraitesDebit_m3j_f As Double
Private md_PompesBouesExcesPression_f As Double
Private md_PompesBouesExcesRendement_f As Double

'Paramètres du modèle ATV A131
Private md_CoefficientAConsommationSpecifiqueO2RespirationCarbone_kgO2_kgDBO As Double
Private md_CoefficientBConsommationSpecifiqueO2RespirationCarbone_kgO2_kgDBOj As Double
Private md_CoefficientCConsommationSpecifiqueO2RespirationCarbone As Double
Private md_CoefficientDConsommationSpecifiqueO2RespirationCarbone As Double
Private md_CoefficientBFacteurImpactTemperature_°C As Double
Private md_CoefficientAFacteurImpactTemperature As Double
Private md_CoefficientAProductionSpecifiqueBouesCarbone_kgMES_kgDBO As Double
Private md_CoefficientBProductionSpecifiqueBouesCarbone As Double
Private md_CoefficientCProductionSpecifiqueBouesCarbone As Double
Private md_CoefficientDProductionSpecifiqueBouesCarbone As Double
Private md_CoefficientEProductionSpecifiqueBouesCarbone_kgMES_kgDBO As Double
Private md_CoefficientFProductionSpecifiqueBouesCarbone As Double
Private md_CoefficientGProductionSpecifiqueBouesCarbone As Double
Private md_CoefficientARatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie As Double

'Autres paramètres
Private md_TempsPassageDebitEntreeAnaerobie_j As Double
Private md_ConsommationSpecifiqueO2Nitrification_kgO2_kgNNitrifie As Double
Private md_ConsommationSpecifiqueDCODenitrification_kgDCO_kgN As Double
Private md_AssimilationN_kgN_kgDBOeliminee As Double
Private md_ConsommationDCOexterneDenitrification_kgDCO_kgNDenitrifie As Double
Private md_DCOdansMethanol_kgDCO_kg As Double
Private md_RatioDBOsurDCOcarboneExterne_kgDBO_kgDCO As Double
Private md_CoefficientPolynomeCalculRendementDBO(0 To 3) As Double
Private md_TemperatureReferenceAgeBouesEquivalent_°C As Double
Private md_CoefficientCalculAgeBouesEquivalent As Double
Private md_AgeBouesCarboneSeul12°C_j As Double
Private md_AgeBouesStabilisationAerobie12°C_j As Double

'Variables à véhiculer
Private mn_RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobieFixe As Boolean


'variables des hypothèses
Private G_reference As Double
Private MV_MES_reference As Double
Private ratio_DCO_NK_10 As Double
Private a1_Cm As Double
Private a0_Cm As Double
Private ratio_P_synthese As Double
Private a_Pbio_coef_K As Double
Private b_Pbio_coef_K As Double
Private a_Kp_bio_1 As Double
Private b_Kp_bio_1 As Double
Private lim_Kp_bio_1 As Double
Private rdt_Psol_max As Double
Private ratio_Psol_P As Double
Private ratio_FeCl3_Pt As Double
Private a_dosis_MMS As Double
Private b_dosis_MMS As Double
Private rdt_DCO As Double
Private ratio_DCO_sol_EB As Double
Private rdt_DCO_sol_primaire As Double
Private ratio_C_N_limite_nominal As Double
Private ratio_boues_methanol As Double
Private charge_radier_max(mi_RACLEUR_NOMBRE_TECHNOLOGIE) As Double
Private facteur_boues_C_lim As Double
Private O2_alpha0(mi_AERATEUR_NOMBRE_TYPE) As Double
Private a0_alpha As Double
Private a_alpha_MVS As Double
Private a_alpha_G As Double
Private a0_correctif_alfa_MES As Double
Private a1_correctif_alfa_MES As Double
Private MES_reference_alfa As Double
Private O2_facteur_beta As Double
Private P0 As Double
Private a0_Patm As Double
Private a1_Patm As Double
Private a2_Patm As Double
Private a0_delta_P As Double
Private a1_delta_P As Double
Private hauteur_diffuseur As Double
Private a_O2_saturation(4) As Double
Private md_CoefficientCalculAgeBouesEquivalent_K As Double
Private T_ref_K As Double
Private ratio_kgO2_Nm3air As Double
Private O2_rdt_dissolution(mi_AERATEUR_NOMBRE_TYPE) As Double
Private ASB(mi_AERATEUR_NOMBRE_TYPE) As Double
Private impact_regulation(mi_AERATION_REGULATION_NOMBRE_TYPE) As Double
Private rate_puissance_fixe As Double
Private rate_puissance_debit As Double
Private agitation_fct As Double
Private ratio_elec_recirculation As Double
Private ratio_elec_recirculation_MLSS As Double
Private ratio_elec_extraction As Double
Private surpresseur_conso_Wh_Nm3mCE(mi_SURPRESSEUR_NOMBRE_TECHNOLOGIE, nb_surpresseur_gamme_Q, nb_surpresseur_gamme_P) As Double
Private Pt_residuel As Double

'stockage de valeurs réelles
Private stockage_MVS_bassins As Double
Private stockage_Q As Double
Private boues_chimiques As Double
Private boues_methanol As Double

Private remarque(5) As String

Private choix_fct_SHUNT As Integer

Sub dimensionnement_vieux(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef Sh, ByRef DCO_apportee_methanol, ByRef DCO_conso_bio)
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'variables internes
    Dim G_equivalent As Double
    Dim Cm_equivalente As Double
    Dim Cm As Double
    Dim boues_bassins_MVS As Double
    Dim rdt_DBO As Double
    Dim boues_produites As Double
    Dim I As Integer
    Dim facteur_T_O2 As Double
    Dim ratio_O2_DBO_respiration As Double
    Dim nitrification As Double
    Dim denitrification As Double
    Dim P_synthese As Double
    Dim P_bio_coeff_K As Double
    Dim volume_optimal As Double
    Dim Kp_bio_1 As Double
    Dim Kp_bio_2 As Double
    Dim ratio_DBO_P_in As Double
    Dim dose_MMS As Double
    Dim P_precipitation As Double
    Dim ratio_molaire_Ptbr As Double
    Dim DCO_soluble As Double
    Dim ratio_C_N As Double
    Dim DCO_necessaire As Double
    Dim boues_bassins_MES As Double
    Dim Ptbr As Double

    'DIMENSIONNEMENT DES BASSINS

    'Age de boues equivalent à md_TemperatureReferenceAgeBouesEquivalent_°C
    G_equivalent = md_NominalAgeBoues_j * md_CoefficientCalculAgeBouesEquivalent ^ (pd_TemperatureEauDimensionnement_°C - md_TemperatureReferenceAgeBouesEquivalent_°C)
    'Charge massique équivalente
    Cm_equivalente = 1 / (a1_Cm * G_equivalent + a0_Cm)
    'Charge massique à T_design
    Cm = Cm_equivalente * md_CoefficientCalculAgeBouesEquivalent ^ (pd_TemperatureEauDimensionnement_°C - md_TemperatureReferenceAgeBouesEquivalent_°C)
    'Masse de boues dans les bassins kgMVS
    boues_bassins_MVS = DBO / Cm
    'on recalcule la part de MVS dans les MES
    boues_bassins_MES = boues_bassins_MVS / MV_MES_reference + (boues_chimiques + boues_methanol) * md_NominalAgeBoues_j
    pd_NominalRatioMVMES = (boues_bassins_MVS + boues_methanol * md_NominalAgeBoues_j) / boues_bassins_MES
    If mn_NominalRatioMVMES_force = True Then
        pd_NominalRatioMVMES = pd_NominalRatioMVMES_f
    End If
    'Volume de bassins m3
    md_VolumeBassinsTotal_m3 = boues_bassins_MVS / (pd_NominalRatioMVMES * md_NominalConcentrationBassinsMES_gL)
    If mn_VolumeBassinsTotal_m3_force = True Then
        md_VolumeBassinsTotal_m3 = md_VolumeBassinsTotal_m3_f
        'calcul de la Cm_equivalente avec ce volume de bassins
        boues_bassins_MVS = md_VolumeBassinsTotal_m3 * pd_NominalRatioMVMES * md_NominalConcentrationBassinsMES_gL
        Cm = DBO / boues_bassins_MVS
        Cm_equivalente = Cm / md_CoefficientCalculAgeBouesEquivalent ^ (pd_TemperatureEauDimensionnement_°C - md_TemperatureReferenceAgeBouesEquivalent_°C)
    End If

    'volume anaérobie
    If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE Or mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE_ET_PRECIPITATION Then
        md_VolumeBassinAnaerobie_m3 = Q / md_TempsPassageDebitEntreeAnaerobie_j * md_CoefficientCalculAgeBouesEquivalent ^ (md_TemperatureReferenceAgeBouesEquivalent_°C - pd_TemperatureEauDimensionnement_°C)
        If mn_VolumeBassinAnaerobie_m3_force = True Then
            md_VolumeBassinAnaerobie_m3 = md_VolumeBassinAnaerobie_m3_f
        End If
    Else
        md_VolumeBassinAnaerobie_m3 = 0
    End If

    'autres volumes
    If mi_ChoixConfigurationBassins = mi_BASSINS_SEPARES Then
        If DCO / NK < ratio_DCO_NK_10 Then
            md_VolumeBassinAnoxie_m3 = 0.4 * (md_VolumeBassinsTotal_m3 - md_VolumeBassinAnaerobie_m3)
        Else
            md_VolumeBassinAnoxie_m3 = 0.2 * (md_VolumeBassinsTotal_m3 - md_VolumeBassinAnaerobie_m3)
        End If
        If mn_VolumeBassinAnoxie_m3_force = True Then
            md_VolumeBassinAnoxie_m3 = md_VolumeBassinAnoxie_m3_f
        End If
        md_VolumeBassinAerobie_m3 = md_VolumeBassinsTotal_m3 - (md_VolumeBassinAnoxie_m3 + md_VolumeBassinAnaerobie_m3)
        If mn_VolumeBassinAerobie_m3_force = True Then
            md_VolumeBassinAerobie_m3 = md_VolumeBassinAerobie_m3_f
        End If
        If md_VolumeBassinsTotal_m3 <> (md_VolumeBassinAnaerobie_m3 + md_VolumeBassinAnoxie_m3 + md_VolumeBassinAerobie_m3) Then
            remarque(1) = Feuil13.Range("BA_faible_REPEREerreur_volume").Cells(1, 1)
        Else
            remarque(1) = ""
        End If
    Else
        md_VolumeChenal_m3 = md_VolumeBassinsTotal_m3 - md_VolumeBassinAnaerobie_m3
        If mn_VolumeChenal_m3_force = True Then
            md_VolumeChenal_m3 = md_VolumeChenal_m3_f
        End If
    End If

    'surface de radier du clarificateur (m2)
    md_SurfaceTotaleClarification_m2 = (Q_nominal * pd_CoefficientPointeHydrauliqueTempsPluie + (Q - Q_nominal)) / (pd_NOMBRE_HEURE_PAR_JOUR * md_ClarificationVitesseMaximale_mh)
    If mn_SurfaceTotaleClarification_m2_force = True Then
        md_SurfaceTotaleClarification_m2 = md_SurfaceTotaleClarification_m2_f
    End If


    'FONCTIONNEMENT NOMINAL

    'rendement d'élimination de la DBO
    rdt_DBO = 0
    For I = 0 To 3
        rdt_DBO = rdt_DBO + md_CoefficientPolynomeCalculRendementDBO(I) * (Cm_equivalente / pd_NominalRatioMVMES) ^ I
    Next I

    'traitement de l'azote
    nitrification = NH4 - md_ConcentrationSortieNH4_mgL * Q / 1000
    denitrification = NO3 + nitrification - md_ConcentrationSortieNO3_mgL * Q / 1000
    'calcul de l'ajout de méthanol pour la dénitrification
    DCO_soluble = DCO_nominal * ratio_DCO_sol_EB * (1 - rdt_DCO_sol_primaire)
    ratio_C_N = DCO_soluble / denitrification
    If ratio_C_N < ratio_C_N_limite_nominal Then
        DCO_necessaire = ratio_C_N_limite_nominal * denitrification - DCO_soluble
        pd_ConsommationMethanolPur_kgj = DCO_necessaire / md_DCOdansMethanol_kgDCO_kg
    End If
    DCO_necessaire = pd_ConsommationMethanolPur_kgj * md_DCOdansMethanol_kgDCO_kg
    DCO_apportee_methanol = DCO_necessaire
    boues_methanol = ratio_boues_methanol * DCO_necessaire
    DCO = DCO + DCO_necessaire * (1 - ratio_boues_methanol)

    'traitement du phosphore
    P_synthese = ratio_P_synthese * DBO * rdt_DBO
    If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE Then
        If mn_RendementDephosphatationBiologique_force = True Then
            md_RendementDephosphatationBiologique = md_RendementDephosphatationBiologique_f
        Else
            P_bio_coeff_K = a_Pbio_coef_K + b_Pbio_coef_K / G_equivalent
            ratio_DBO_P_in = DBO / Pt
            volume_optimal = Q / md_TempsPassageDebitEntreeAnaerobie_j * md_CoefficientCalculAgeBouesEquivalent ^ (md_TemperatureReferenceAgeBouesEquivalent_°C - pd_TemperatureEauConditionsReelles_°C)
            Kp_bio_1 = md_VolumeBassinAnaerobie_m3 / volume_optimal
            If Kp_bio_1 <= 1 Then
                Kp_bio_1 = Kp_bio_1 ^ 2
            ElseIf Kp_bio_1 <= 1.5 Then
                Kp_bio_1 = a_Kp_bio_1 + b_Kp_bio_1 * Kp_bio_1
            Else
                Kp_bio_1 = lim_Kp_bio_1
            End If
            Kp_bio_2 = 1
            md_RendementDephosphatationBiologique = P_bio_coeff_K * Kp_bio_1 * Kp_bio_2 * ratio_DBO_P_in
            If md_RendementDephosphatationBiologique > rdt_Psol_max Then
                md_RendementDephosphatationBiologique = rdt_Psol_max
            End If
        End If
        Pt = Pt - md_RendementDephosphatationBiologique * ratio_Psol_P * Pt
        md_ConcentrationSortiePt_mgL = Pt / (Q - pd_BouesExtraitesDebit_m3j) * 1000
        boues_chimiques = 0
        P_precipite = 0
    ElseIf mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE_ET_PRECIPITATION Then
        If mn_RendementDephosphatationBiologique_force = True Then
            md_RendementDephosphatationBiologique = md_RendementDephosphatationBiologique_f
        Else
            P_bio_coeff_K = a_Pbio_coef_K + b_Pbio_coef_K / G_equivalent
            ratio_DBO_P_in = DBO / Pt
            volume_optimal = Q / md_TempsPassageDebitEntreeAnaerobie_j * md_CoefficientCalculAgeBouesEquivalent ^ (md_TemperatureReferenceAgeBouesEquivalent_°C - pd_TemperatureEauConditionsReelles_°C)
            Kp_bio_1 = md_VolumeBassinAnaerobie_m3 / volume_optimal
            If Kp_bio_1 <= 1 Then
                Kp_bio_1 = Kp_bio_1 ^ 2
            ElseIf Kp_bio_1 <= 1.5 Then
                Kp_bio_1 = a_Kp_bio_1 + b_Kp_bio_1 * Kp_bio_1
            Else
                Kp_bio_1 = lim_Kp_bio_1
            End If
            Kp_bio_2 = 1
            md_RendementDephosphatationBiologique = P_bio_coeff_K * Kp_bio_1 * Kp_bio_2 * ratio_DBO_P_in
            If md_RendementDephosphatationBiologique > rdt_Psol_max Then
                md_RendementDephosphatationBiologique = rdt_Psol_max
            End If
        End If
        Pt = Pt - md_RendementDephosphatationBiologique * ratio_Psol_P * Pt
        Pt = Pt - P_synthese
        P_precipitation = Pt - md_ConcentrationSortiePt_mgL * (Q - pd_BouesExtraitesDebit_m3j) / 1000
        If P_precipitation > 0 Then
            Ptbr = P_precipitation / Q * 1000
            ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + md_ConcentrationSortiePt_mgL) ^ (-0.5) - ((md_ConcentrationSortiePt_mgL) ^ (-0.5) - 0.1)) / ((Ptbr + md_ConcentrationSortiePt_mgL) - (md_ConcentrationSortiePt_mgL - 0.1))
            pd_ConsommationFeCl3Pur_kgj = ratio_molaire_Ptbr * ratio_FeCl3_Pt * P_precipitation
        Else
            P_precipitation = 0
            pd_ConsommationFeCl3Pur_kgj = 0
        End If
        dose_MMS = (ratio_molaire_Ptbr - 1) * a_dosis_MMS + b_dosis_MMS
        boues_chimiques = dose_MMS * P_precipitation
        Pt = Pt - P_precipitation
        P_precipite = P_precipitation
    ElseIf mi_ChoixDephosphatation = mi_DEPHOSPHATATION_PRECIPITATION Then
        Pt = Pt - P_synthese
        P_precipitation = Pt - md_ConcentrationSortiePt_mgL * (Q - pd_BouesExtraitesDebit_m3j) / 1000
        If P_precipitation > 0 Then
            Ptbr = P_precipitation / Q * 1000
            ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + md_ConcentrationSortiePt_mgL) ^ (-0.5) - ((md_ConcentrationSortiePt_mgL) ^ (-0.5) - 0.1)) / ((Ptbr + md_ConcentrationSortiePt_mgL) - (md_ConcentrationSortiePt_mgL - 0.1))
            pd_ConsommationFeCl3Pur_kgj = ratio_molaire_Ptbr * ratio_FeCl3_Pt * P_precipitation
        Else
            P_precipitation = 0
            pd_ConsommationFeCl3Pur_kgj = 0
        End If
        dose_MMS = (ratio_molaire_Ptbr - 1) * a_dosis_MMS + b_dosis_MMS
        boues_chimiques = dose_MMS * P_precipitation
        Pt = Pt - P_precipitation
        P_precipite = P_precipitation
    Else
        Pt = Pt - P_synthese
        boues_chimiques = 0
        P_precipite = 0
    End If


    'ratio O2 respiration et élimination DBO5
    facteur_T_O2 = md_CoefficientAFacteurImpactTemperature ^ (pd_TemperatureEauConditionsReelles_°C - md_CoefficientBFacteurImpactTemperature_°C)
    ratio_O2_DBO_respiration = md_CoefficientAConsommationSpecifiqueO2RespirationCarbone_kgO2_kgDBO + (md_CoefficientBConsommationSpecifiqueO2RespirationCarbone_kgO2_kgDBOj * md_ReelAgeBoues_j * facteur_T_O2) / (md_CoefficientCConsommationSpecifiqueO2RespirationCarbone + md_CoefficientDConsommationSpecifiqueO2RespirationCarbone * md_ReelAgeBoues_j * facteur_T_O2)

    'calcul des besoins en O2
    md_BesoinsOxygene_kgO2j = ratio_O2_DBO_respiration * rdt_DBO * DBO
    md_BesoinsOxygene_kgO2j = md_BesoinsOxygene_kgO2j + md_ConsommationSpecifiqueO2Nitrification_kgO2_kgNNitrifie * nitrification
    md_BesoinsOxygene_kgO2j = md_BesoinsOxygene_kgO2j + md_ConsommationSpecifiqueDCODenitrification_kgDCO_kgN * denitrification
    md_BesoinsOxygene_kgO2j = md_BesoinsOxygene_kgO2j + pd_ConsommationMethanolPur_kgj * md_DCOdansMethanol_kgDCO_kg  'on ne consomme pas le méthanol directement mais il représente de la DCO non consommée par la dénit
    'DCO_conso_bio = ratio_O2_DBO_respiration * (DBO - md_ConcentrationSortieDBO_mgL * stockage_Q / 1000) + DCO_apportee_methanol * (1 - ratio_boues_methanol) * rdt_DCO - md_ConsommationSpecifiqueDCODenitrification_kgDCO_kgN * denitrification

    'Production de boues
    boues_produites = boues_bassins_MVS / (pd_NominalRatioMVMES * md_NominalAgeBoues_j)

    'vitesse hydraulique max (recalculée) clarificateur
    md_ClarificationVitesseMaximaleRecalculee_mh = (Q_nominal * pd_CoefficientPointeHydrauliqueTempsPluie + (Q - Q_nominal)) / (pd_NOMBRE_HEURE_PAR_JOUR * md_SurfaceTotaleClarification_m2)
    'charge au radier clarificateur
    md_ClarificateurChargeRadierNominal = md_NominalConcentrationBassinsMES_gL * Q * (1 + md_TauxRecirculationBouesClarificateur) / (pd_NOMBRE_HEURE_PAR_JOUR * md_SurfaceTotaleClarification_m2)
    If md_ClarificateurChargeRadierNominal > charge_radier_max(mi_ChoixTypeRacleur) Then
        remarque(2) = Feuil13.Range("BA_faible_REPEREerreur_charge_radier").Cells(1, 1)
    Else
        remarque(2) = ""
    End If


    'CALCUL DES SORTIES

    'Eau traitée
    Q = Q - pd_BouesExtraitesDebit_m3j
    MES = md_EauClarifieeMES_mgL * Q / 1000
    DCO = DCO * (1 - rdt_DCO)
    DBO = DBO * (1 - rdt_DBO)
    NH4 = NH4 - nitrification
    NK = NH4

    'Boues
    pd_BouesExtraitesFluxMES_kgj = boues_produites - MES
    md_BouesExtraitesConcentrationMES_gL = md_NominalConcentrationBassinsMES_gL * (1 + md_TauxRecirculationBouesClarificateur) / md_TauxRecirculationBouesClarificateur
    pd_BouesExtraitesDebit_m3j = pd_BouesExtraitesFluxMES_kgj / md_BouesExtraitesConcentrationMES_gL

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "13.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Private Function F_age_boues_conseille_j(ByVal EntreeDBO_kgj As Double, ByVal TemperatureEauxUsees_°C As Double, ByVal NDenitrifie_kgj As Double, ByRef RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie As Double) As Double


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette fonction permet de calculer l'âge de boues recommandé par la méthode ATV A131 par itération sur l'âge de boues.

'VARIABLES D'ENTREE:
'   EntreeDBO_kgj: flux de DBO5 en entrée de la boues activées en kg/j
'   TemperatureEauxUsees_°C: température des eaux usées en °C
'   NDenitrifie_kgj: quantité d'azote déntirifié
'   F_consommation_specifique_O2_respiration_carbone_kgO2_kgDBO: consommation spécifique d'oxygène pour abattre le carbone et pour la respiration
'   md_ConsommationSpecifiqueDCODenitrification_kgDCO_kgN: consommation spécifique de la DCO pour la dénitrification
'   md_CRITERE_CONVERGENCE_AGE_BOUES: critère de convergence pour le calcul de l'âge de boues

'VARIABLES DE SORTIE:
'   F_age_boues_conseille_j: âge de boues conseillé par l'ATV A131
'   RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie: part du volume utilisé pour la dénitrification (volume anoxique)
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim AgeBouesDefini_j As Double
Dim FacteurSecurite As Double
Dim AgeBouesAerobieMinimum_j As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "13.9"
AgeBouesDefini_j = 0
FacteurSecurite = 0
AgeBouesAerobieMinimum_j = 0
F_age_boues_conseille_j = 0

'Calcul du facteur de sécurité
If EntreeDBO_kgj <= 1200 Then
    FacteurSecurite = 1.8
ElseIf EntreeDBO_kgj >= 6000 Then
    FacteurSecurite = 1.45
Else
    FacteurSecurite = 1.8 + (EntreeDBO_kgj - 1200) / (6000 - 1200) * (1.45 - 1.8)
End If

'Calcul de l'âge de boues aérobie nécessaire
AgeBouesAerobieMinimum_j = 3.4 * FacteurSecurite * 1.103 ^ (15 - TemperatureEauxUsees_°C)
If AgeBouesAerobieMinimum_j < 5 Then
    AgeBouesAerobieMinimum_j = 5
End If

'Calcul par itération de l'âge de boues conseillé
F_age_boues_conseille_j = 20
Do
    AgeBouesDefini_j = F_age_boues_conseille_j
    'Calcul du ratio volume anoxique / volume anoxique et aérobie
    RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie = F_ratio_volume_anoxique_sur_volume_anoxique_et_aerobie(TemperatureEauxUsees_°C, NDenitrifie_kgj, AgeBouesDefini_j, EntreeDBO_kgj)
    'Calcul de l'âge de boues conseillé
    F_age_boues_conseille_j = AgeBouesAerobieMinimum_j / (1 - RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie)
Loop While ((F_age_boues_conseille_j - AgeBouesDefini_j) ^ 2 / AgeBouesDefini_j) > md_CRITERE_CONVERGENCE_AGE_BOUES


Exit Function
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Function


Private Function F_consommation_specifique_O2_respiration_carbone_kgO2_kgDBO(ByVal TemperatureEauxUsees_°C As Double, ByVal AgeBoues_j As Double) As Double


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette fonction permet de calculer la consommation spécifique en O2 pour la dégradation du carbone et la respiration selon la méthode ATV A131.

'VARIABLES D'ENTREE:
'   TemperatureEauxUsees_°C: température des eaux usées en °C
'   AgeBoues_j: âge de boues
'   F_facteur_impact_temperature: facteur de correction en fonction de la température
'   Coefficients de la formule de calcul de F_consommation_specifique_O2_respiration_carbone_kgO2_kgDBO:
'       md_CoefficientAConsommationSpecifiqueO2RespirationCarbone_kgO2_kgDBO
'       md_CoefficientBConsommationSpecifiqueO2RespirationCarbone_kgO2_kgDBOj
'       md_CoefficientCConsommationSpecifiqueO2RespirationCarbone
'       md_CoefficientDConsommationSpecifiqueO2RespirationCarbone
    
'VARIABLES DE SORTIE:
'   F_consommation_specifique_O2_respiration_carbone_kgO2_kgDBO: consommation spécifique en O2 pour la dégradation du carbone et la respiration
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "13.9"


'Calcul de la consommation spécifique en O2 pour la dégradation du carbone et la respiration
F_consommation_specifique_O2_respiration_carbone_kgO2_kgDBO = md_CoefficientAConsommationSpecifiqueO2RespirationCarbone_kgO2_kgDBO + (md_CoefficientBConsommationSpecifiqueO2RespirationCarbone_kgO2_kgDBOj * AgeBoues_j * F_facteur_impact_temperature(TemperatureEauxUsees_°C)) / (md_CoefficientCConsommationSpecifiqueO2RespirationCarbone + md_CoefficientDConsommationSpecifiqueO2RespirationCarbone * AgeBoues_j * F_facteur_impact_temperature(TemperatureEauxUsees_°C))


Exit Function
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Function

Private Function F_facteur_impact_temperature(ByVal TemperatureEauxUsees_°C As Double) As Double


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette fonction permet de calculer l'impact de la température sur la biologie (production de boues et consommation d'oxygène).

'VARIABLES D'ENTREE:
'   TemperatureEauxUsees_°C: température des eaux usées en °C
'   Coefficients de la formule de calcul de F_facteur_impact_temperature:
'       md_CoefficientAFacteurImpactTemperature
'       md_CoefficientBFacteurImpactTemperature_°C

'VARIABLES DE SORTIE:
'   F_facteur_impact_temperature: facteur de correction en fonction de la température
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "13.9"


'Calcul de la consommation spécifique en O2 pour la dégradation du carbone et la respiration
F_facteur_impact_temperature = md_CoefficientAFacteurImpactTemperature ^ (TemperatureEauxUsees_°C - md_CoefficientBFacteurImpactTemperature_°C)


Exit Function
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Function

Private Function F_rendement_dephosphatation_biologique(ByVal EntreeDBO_kgj As Double, ByVal TemperatureEauxUsees_°C As Double, ByVal AgeBouesDefini_j As Double, ByVal EntreePt_kgj As Double, ByVal Debit_m3j As Double) As Double


'******************************************************************************************************************************************************
'OBJECTIF:

'VARIABLES D'ENTREE:

'VARIABLES DE SORTIE:
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim AgeBouesEquivalent12°C_j As Double
Dim BioPCoefficientK As Double
Dim EntreeRatioDBOsurPt As Double
Dim VolumeAnaerobieOptimal_m3 As Double
Dim CoefficientKpBio1 As Double
Dim CoefficientKpBio2 As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "13.9"
AgeBouesEquivalent12°C_j = 0
BioPCoefficientK = 0
EntreeRatioDBOsurPt = 0
VolumeAnaerobieOptimal_m3 = 0
CoefficientKpBio1 = 0
CoefficientKpBio2 = 0

'on calcule l'âge de boues équivalent
AgeBouesEquivalent12°C_j = AgeBouesDefini_j * md_CoefficientCalculAgeBouesEquivalent ^ (TemperatureEauxUsees_°C - md_TemperatureReferenceAgeBouesEquivalent_°C)






BioPCoefficientK = a_Pbio_coef_K + b_Pbio_coef_K / AgeBouesEquivalent12°C_j
EntreeRatioDBOsurPt = EntreeDBO_kgj / EntreePt_kgj
VolumeAnaerobieOptimal_m3 = Debit_m3j / md_TempsPassageDebitEntreeAnaerobie_j * md_CoefficientCalculAgeBouesEquivalent ^ (md_TemperatureReferenceAgeBouesEquivalent_°C - pd_TemperatureEauConditionsReelles_°C)
CoefficientKpBio1 = md_VolumeBassinAnaerobie_m3 / VolumeAnaerobieOptimal_m3
If CoefficientKpBio1 <= 1 Then
    CoefficientKpBio1 = CoefficientKpBio1 ^ 2
ElseIf CoefficientKpBio1 <= 1.5 Then
    CoefficientKpBio1 = a_Kp_bio_1 + b_Kp_bio_1 * CoefficientKpBio1
Else
    CoefficientKpBio1 = lim_Kp_bio_1
End If
CoefficientKpBio2 = 1
F_rendement_dephosphatation_biologique = BioPCoefficientK * CoefficientKpBio1 * CoefficientKpBio2 * EntreeRatioDBOsurPt
If F_rendement_dephosphatation_biologique > rdt_Psol_max Then
    F_rendement_dephosphatation_biologique = rdt_Psol_max
End If


Exit Function
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Function

Private Sub calcul_traitement_phosphore(ByVal Debit_m3j As Double, ByVal EntreeDBO_kgj As Double, ByRef FluxPt_kgj As Double, ByVal TemperatureEauxUsees_°C As Double, ByVal DBOeliminee_kgj As Double, ByVal AgeBouesDefini_j As Double)


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette fonction permet de calculer l'impact de la température sur la biologie (production de boues et consommation d'oxygène).

'VARIABLES D'ENTREE:
'   TemperatureEauxUsees_°C: température des eaux usées en °C
'   Coefficients de la formule de calcul de F_facteur_impact_temperature:
'       md_CoefficientAFacteurImpactTemperature
'       md_CoefficientBFacteurImpactTemperature_°C

'VARIABLES DE SORTIE:
'   F_facteur_impact_temperature: facteur de correction en fonction de la température
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim EntreePt_kgj As Double
Dim ConsommationPtSynthese_kgj As Double
Dim PtPrecipite_kgj As Double
Dim PtPrecipite_mgL As Double
Dim RatioMolairePtbr As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "13.9"
EntreePt_kgj = 0
ConsommationPtSynthese_kgj = 0
PtPrecipite_kgj = 0
PtPrecipite_mgL = 0
RatioMolairePtbr = 0


EntreePt_kgj = FluxPt_kgj
ConsommationPtSynthese_kgj = ratio_P_synthese * DBOeliminee_kgj
If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE Then
    If mn_RendementDephosphatationBiologique_force = True Then
        md_RendementDephosphatationBiologique = md_RendementDephosphatationBiologique_f
    Else
        md_RendementDephosphatationBiologique = F_rendement_dephosphatation_biologique(EntreeDBO_kgj, TemperatureEauxUsees_°C, AgeBouesDefini_j, EntreePt_kgj, Debit_m3j)
    End If
    FluxPt_kgj = EntreePt_kgj * (1 - md_RendementDephosphatationBiologique * ratio_Psol_P)
    md_ConcentrationSortiePt_mgL = FluxPt_kgj / (Debit_m3j - pd_BouesExtraitesDebit_m3j) * 1000
    boues_chimiques = 0
    P_precipite = 0
ElseIf mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE_ET_PRECIPITATION Then
    If mn_RendementDephosphatationBiologique_force = True Then
        md_RendementDephosphatationBiologique = md_RendementDephosphatationBiologique_f
    Else
        md_RendementDephosphatationBiologique = F_rendement_dephosphatation_biologique(EntreeDBO_kgj, TemperatureEauxUsees_°C, AgeBouesDefini_j, EntreePt_kgj, Debit_m3j)
    End If
    FluxPt_kgj = EntreePt_kgj * (1 - md_RendementDephosphatationBiologique * ratio_Psol_P)
    FluxPt_kgj = FluxPt_kgj - ConsommationPtSynthese_kgj
    PtPrecipite_kgj = FluxPt_kgj - md_ConcentrationSortiePt_mgL * (Debit_m3j - pd_BouesExtraitesDebit_m3j) / 1000
    If PtPrecipite_kgj > 0 Then
        PtPrecipite_mgL = PtPrecipite_kgj / Debit_m3j * 1000
        RatioMolairePtbr = 1.1 - 6 * ((PtPrecipite_mgL + md_ConcentrationSortiePt_mgL) ^ (-0.5) - ((md_ConcentrationSortiePt_mgL) ^ (-0.5) - 0.1)) / ((PtPrecipite_mgL + md_ConcentrationSortiePt_mgL) - (md_ConcentrationSortiePt_mgL - 0.1))
        pd_ConsommationFeCl3Pur_kgj = RatioMolairePtbr * ratio_FeCl3_Pt * PtPrecipite_kgj
    Else
        PtPrecipite_kgj = 0
        pd_ConsommationFeCl3Pur_kgj = 0
    End If
    dose_MMS = (RatioMolairePtbr - 1) * a_dosis_MMS + b_dosis_MMS
    boues_chimiques = dose_MMS * PtPrecipite_kgj
    FluxPt_kgj = Pt - PtPrecipite_kgj
    P_precipite = PtPrecipite_kgj
ElseIf mi_ChoixDephosphatation = mi_DEPHOSPHATATION_PRECIPITATION Then
    FluxPt_kgj = EntreePt_kgj - ConsommationPtSynthese_kgj
    PtPrecipite_kgj = FluxPt_kgj - md_ConcentrationSortiePt_mgL * (Debit_m3j - pd_BouesExtraitesDebit_m3j) / 1000
    If PtPrecipite_kgj > 0 Then
        PtPrecipite_mgL = PtPrecipite_kgj / Q * 1000
        RatioMolairePtbr = 1.1 - 6 * ((PtPrecipite_mgL + md_ConcentrationSortiePt_mgL) ^ (-0.5) - ((md_ConcentrationSortiePt_mgL) ^ (-0.5) - 0.1)) / ((PtPrecipite_mgL + md_ConcentrationSortiePt_mgL) - (md_ConcentrationSortiePt_mgL - 0.1))
        pd_ConsommationFeCl3Pur_kgj = RatioMolairePtbr * ratio_FeCl3_Pt * PtPrecipite_kgj
    Else
        PtPrecipite_kgj = 0
        pd_ConsommationFeCl3Pur_kgj = 0
    End If
    dose_MMS = (RatioMolairePtbr - 1) * a_dosis_MMS + b_dosis_MMS
    boues_chimiques = dose_MMS * PtPrecipite_kgj
    FluxPt_kgj = FluxPt_kgj - PtPrecipite_kgj
    P_precipite = PtPrecipite_kgj
Else
    FluxPt_kgj = EntreePt_kgj - ConsommationPtSynthese_kgj
    boues_chimiques = 0
    P_precipite = 0
End If



Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Private Function F_ratio_volume_anoxique_sur_volume_anoxique_et_aerobie(ByVal TemperatureEauxUsees_°C As Double, ByVal NDenitrifie_kgj As Double, ByVal AgeBouesDefini_j As Double, ByVal EntreeDBO_kgj As Double) As Double


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette fonction permet de calculer l'impact de la température sur la biologie (production de boues et consommation d'oxygène).

'VARIABLES D'ENTREE:
'   TemperatureEauxUsees_°C: température des eaux usées en °C
'   Coefficients de la formule de calcul de F_facteur_impact_temperature:
'       md_CoefficientAFacteurImpactTemperature
'       md_CoefficientBFacteurImpactTemperature_°C

'VARIABLES DE SORTIE:
'   F_ratio_volume_anoxique_sur_volume_anoxique_et_aerobie: part du volume utilisé pour la dénitrification (volume anoxique)
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "13.9"


'Calcul du ratio volume anoxique / volume anoxique et aérobie
If mn_RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobieFixe = True Then
    F_ratio_volume_anoxique_sur_volume_anoxique_et_aerobie = md_VolumeBassinAnoxie_m3 / (md_VolumeBassinAnoxie_m3 + md_VolumeBassinAerobie_m3)
Else
    F_ratio_volume_anoxique_sur_volume_anoxique_et_aerobie = (md_ConsommationSpecifiqueDCODenitrification_kgDCO_kgN * NDenitrifie_kgj) / (md_CoefficientARatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie * F_consommation_specifique_O2_respiration_carbone_kgO2_kgDBO(TemperatureEauxUsees_°C, AgeBouesDefini_j) * EntreeDBO_kgj)
    If F_ratio_volume_anoxique_sur_volume_anoxique_et_aerobie < 0.2 Then
        F_ratio_volume_anoxique_sur_volume_anoxique_et_aerobie = 0.2
    ElseIf F_ratio_volume_anoxique_sur_volume_anoxique_et_aerobie > 0.5 Then
        F_ratio_volume_anoxique_sur_volume_anoxique_et_aerobie = 0.5
    End If
End If


Exit Function
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Function

Private Function F_production_specifique_boues_carbone_kgMES_kgDBO(ByVal TemperatureEauxUsees_°C As Double, ByVal AgeBoues_j As Double, ByVal EntreeMES_kgj As Double, ByVal EntreeDBO_kgj As Double) As Double


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette fonction permet de calculer la production spécifique de boues due à l'abattement du carbone selon la méthode ATV A131.

'VARIABLES D'ENTREE:
'   TemperatureEauxUsees_°C: température des eaux usées en °C
'   AgeBoues_j: âge de boues en j
'   F_facteur_impact_temperature: facteur de correction en fonction de la température
'   EntreeMES_kgj: flux de MES en entrée de la boues activées
'   EntreeDBO_kgj: flux de DBO5 en entrée de la boues activées
'   Coefficients de la formule de calcul de F_production_specifique_boues_carbone_kgMES_kgDBO:
'       md_CoefficientAProductionSpecifiqueBouesCarbone_kgMES_kgDBO
'       md_CoefficientBProductionSpecifiqueBouesCarbone
'       md_CoefficientCProductionSpecifiqueBouesCarbone
'       md_CoefficientDProductionSpecifiqueBouesCarbone
'       md_CoefficientEProductionSpecifiqueBouesCarbone_kgMES_kgDBO
'       md_CoefficientFProductionSpecifiqueBouesCarbone
'       md_CoefficientGProductionSpecifiqueBouesCarbone
  
'VARIABLES DE SORTIE:
'   F_production_specifique_boues_carbone_kgMES_kgDBO: production spécifique de boues due à l'abattement du carbone en kgMES/kgDBO5
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "13.9"


'Calcul de la production spécifique de boues
F_production_specifique_boues_carbone_kgMES_kgDBO = md_CoefficientAProductionSpecifiqueBouesCarbone_kgMES_kgDBO
F_production_specifique_boues_carbone_kgMES_kgDBO = F_production_specifique_boues_carbone_kgMES_kgDBO + md_CoefficientBProductionSpecifiqueBouesCarbone * (EntreeMES_kgj / EntreeDBO_kgj)
F_production_specifique_boues_carbone_kgMES_kgDBO = F_production_specifique_boues_carbone_kgMES_kgDBO - (md_CoefficientCProductionSpecifiqueBouesCarbone * md_CoefficientDProductionSpecifiqueBouesCarbone * md_CoefficientEProductionSpecifiqueBouesCarbone_kgMES_kgDBO * AgeBoues_j * F_facteur_impact_temperature(TemperatureEauxUsees_°C)) / (md_CoefficientFProductionSpecifiqueBouesCarbone + md_CoefficientGProductionSpecifiqueBouesCarbone * AgeBoues_j * F_facteur_impact_temperature(TemperatureEauxUsees_°C))


Exit Function
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Function

Private Function F_RendementEliminationDBO(ByVal TemperatureEauxUsees_°C As Double, ByVal EntreeDBO_kgj As Double, ByVal BassinsConcentrationMES_gL As Double) As Double


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette fonction permet de calculer l'impact de la température sur la biologie (production de boues et consommation d'oxygène).

'VARIABLES D'ENTREE:
'   TemperatureEauxUsees_°C: température des eaux usées en °C
'   Coefficients de la formule de calcul de F_facteur_impact_temperature:
'       md_CoefficientAFacteurImpactTemperature
'       md_CoefficientBFacteurImpactTemperature_°C

'VARIABLES DE SORTIE:
'   F_RendementEliminationDBO: rendement d'élimination de la DBO5
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim ChargeMassique_kgDBO_kgMES As Double
Dim ChargeMassiqueEquivalente_kgDBO_kgMES As Double
'Initialisation des variables locales
NumeroErreurNonGeree = "13.9"
I = 0
ChargeMassique_kgDBO_kgMES = 0
ChargeMassiqueEquivalente_kgDBO_kgMES = 0


'Calcul de la charge massique équivalente à 12°C
ChargeMassique_kgDBO_kgMES = EntreeDBO_kgj / ((md_VolumeBassinsTotal_m3 - md_VolumeBassinAnaerobie_m3) * BassinsConcentrationMES_gL)
ChargeMassiqueEquivalente_kgDBO_kgMES = ChargeMassique_kgDBO_kgMES / (md_CoefficientCalculAgeBouesEquivalent ^ (TemperatureEauxUsees_°C - md_TemperatureReferenceAgeBouesEquivalent_°C))

'Calcul du rendement sur la DBO
F_RendementEliminationDBO = 0
For I = 0 To 3
    F_RendementEliminationDBO = F_RendementEliminationDBO + md_CoefficientPolynomeCalculRendementDBO(I) * ChargeMassiqueEquivalente_kgDBO_kgMES ^ I
Next I


Exit Function
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Function

Private Function F_consommation_carbone_externe_kgDCOj(ByVal TemperatureEauxUsees_°C As Double, ByVal AgeBouesDefini_j As Double, ByVal RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie As Double, ByVal EntreeDBO_kgj As Double, ByVal NDenitrifie_kgj As Double) As Double


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette fonction permet de calculer la consommation de carbone externe (actuellement, seul le méthanol est utilisable) pour la dénitrification.

'VARIABLES D'ENTREE:
'   TemperatureEauxUsees_°C: température des eaux usées en °C
'   AgeBouesDefini_j: âge de boues en j
'   RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie: part du volume utilisé pour la dénitrification (volume anoxique)
'   EntreeDBO_kgj: flux de DBO5 en entrée de la boues activées en kg/j
'   NDenitrifie_kgj: quantité d'azote déntirifié

'VARIABLES DE SORTIE:
'   F_consommation_carbone_externe_kgDCOj: consommation externe de carbone
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim NDenitrifieCarboneExterne_kgj As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "13.9"
NDenitrifieCarboneExterne_kgj = 0

'NNO3 à dénitrifier par une source externe de carbone
NDenitrifieCarboneExterne_kgj = (md_CoefficientARatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie * F_consommation_specifique_O2_respiration_carbone_kgO2_kgDBO(TemperatureEauxUsees_°C, AgeBouesDefini_j)) / md_ConsommationSpecifiqueDCODenitrification_kgDCO_kgN * RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie * EntreeDBO_kgj
NDenitrifieCarboneExterne_kgj = NDenitrifieCarboneExterne_kgj - NDenitrifie_kgj

'On teste si la source externe est nécessaire ou non
If NDenitrifieCarboneExterne_kgj > 0 Then
    F_consommation_carbone_externe_kgDCOj = md_ConsommationDCOexterneDenitrification_kgDCO_kgNDenitrifie * NDenitrifieCarboneExterne_kgj
Else
    F_consommation_carbone_externe_kgDCOj = 0
End If


Exit Function
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Function

Public Sub dimensionnement(ByRef Debit_m3j As Double, ByRef FluxDCO_kgj As Double, ByRef FluxDBO_kgj As Double, ByRef FluxMES_kgj As Double, ByRef FluxNK_kgj As Double, ByRef FluxNH4_kgj As Double, ByRef FluxNO3_kgj As Double, ByRef FluxPt_kgj As Double, ByRef FluxSH_kgj As Double, ByRef MethanolApportDCO_kgj As Double)


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure permet de:
'       - dimensionner les bassins de boues activées
'       - calculer la surface de clarification nécessaire à la séparation des boues

'VARIABLES D'ENTREE:

'VARIABLES DE SORTIE:
'   F_facteur_impact_temperature: facteur de correction en fonction de la température
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim DBOeliminee_kgj As Double
Dim Nnitrifie_kgj As Double
Dim NDenitrifie_kgj As Double
Dim RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie As Double
Dim ProductionBoues_kgMESj As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "13.5"
DBOeliminee_kgj = 0
Nnitrifie_kgj = 0
NDenitrifie_kgj = 0
RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie = 0
ProductionBoues_kgMESj = 0


'On effectue d'abord les bilans matières car nécessaire pour les modèles ATV
DBOeliminee_kgj = FluxDBO_kgj - md_ConcentrationSortieDBO_mgL * Debit_m3j / 1000
If DBOeliminee_kgj < 0 Then
    DBOeliminee_kgj = 0
'    md_ConcentrationSortieDBO_mgL = fluxdbokgj / Debit_m3j * 1000
End If
Nnitrifie_kgj = FluxNK_kgj - md_AssimilationN_kgN_kgDBOeliminee * DBOeliminee_kgj - md_ConcentrationSortieNH4_mgL * Debit_m3j / 1000
If Nnitrifie_kgj < 0 Then
    Nnitrifie_kgj = 0
'    If md_ConcentrationSortieNH4_mgL > (FluxNK_kgj / Debit_m3j * 1000) Then
'        md_ConcentrationSortieNH4_mgL = FluxNK_kgj / Debit_m3j * 1000
'    End If
End If
NDenitrifie_kgj = Nnitrifie_kgj - md_ConcentrationSortieNO3_mgL * Debit_m3j / 1000
If NDenitrifie_kgj < 0 Then
    NDenitrifie_kgj = 0
'    md_ConcentrationSortieNO3_mgL = Nnitrifie_kgj / Debit_m3j * 1000
End If

'On calcule l'âge de boues
If mn_NominalAgeBoues_j_force = True Then
    md_NominalAgeBoues_j = md_NominalAgeBoues_j_f
    RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie = F_ratio_volume_anoxique_sur_volume_anoxique_et_aerobie(pd_TemperatureEauDimensionnement_°C, NDenitrifie_kgj, md_NominalAgeBoues_j, FluxDBO_kgj)
Else
    If pi_ChoixQualiteRejet = pi_QUALITE_REJET_TRAITEMENT_CARBONE Then
        md_NominalAgeBoues_j = md_AgeBouesCarboneSeul12°C_j
    ElseIf pi_ChoixQualiteRejet = pi_QUALITE_REJET_ZONE_SENSIBLE Then
        If pi_ChoixDigestion = True Then
            md_NominalAgeBoues_j = F_age_boues_conseille_j(FluxDBO_kgj, pd_TemperatureEauDimensionnement_°C, NDenitrifie_kgj, RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie)
        Else        'pas de digestion --> stabilisation aérobie des boues
            md_NominalAgeBoues_j = md_AgeBouesStabilisationAerobie12°C_j * md_CoefficientCalculAgeBouesEquivalent ^ (md_TemperatureReferenceAgeBouesEquivalent_°C - pd_TemperatureEauDimensionnement_°C)
            RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie = F_ratio_volume_anoxique_sur_volume_anoxique_et_aerobie(pd_TemperatureEauDimensionnement_°C, NDenitrifie_kgj, md_NominalAgeBoues_j, FluxDBO_kgj)
        End If
    ElseIf pi_ChoixQualiteRejet = pi_QUALITE_REJET_ZONE_TRES_SENSIBLE Then
        md_NominalAgeBoues_j = md_AgeBouesStabilisationAerobie12°C_j * md_CoefficientCalculAgeBouesEquivalent ^ (md_TemperatureReferenceAgeBouesEquivalent_°C - pd_TemperatureEauDimensionnement_°C)
        RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie = F_ratio_volume_anoxique_sur_volume_anoxique_et_aerobie(pd_TemperatureEauDimensionnement_°C, NDenitrifie_kgj, md_NominalAgeBoues_j, FluxDBO_kgj)
    End If
End If

'On calcule l'ajout éventuel de méthanol et on incrémente la DBO et la DCO
pd_ConsommationMethanolPur_kgj = F_consommation_carbone_externe_kgDCOj(pd_TemperatureEauDimensionnement_°C, md_NominalAgeBoues_j, RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie, FluxDBO_kgj, NDenitrifie_kgj) / md_DCOdansMethanol_kgDCO_kg
FluxDBO_kgj = FluxDBO_kgj + pd_ConsommationMethanolPur_kgj * md_DCOdansMethanol_kgDCO_kg * md_RatioDBOsurDCOcarboneExterne_kgDBO_kgDCO
    DBOeliminee_kgj = DBOeliminee_kgj + pd_ConsommationMethanolPur_kgj * md_DCOdansMethanol_kgDCO_kg * md_RatioDBOsurDCOcarboneExterne_kgDBO_kgDCO
FluxDCO_kgj = FluxDCO_kgj + pd_ConsommationMethanolPur_kgj * md_DCOdansMethanol_kgDCO_kg
MethanolApportDCO_kgj = pd_ConsommationMethanolPur_kgj * md_DCOdansMethanol_kgDCO_kg

'Phosphore
'on calcule le volume anaérobie
If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE Or mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE_ET_PRECIPITATION Then
    If mn_VolumeBassinAnaerobie_m3_force = True Then
        md_VolumeBassinAnaerobie_m3 = md_VolumeBassinAnaerobie_m3_f
    Else
        md_VolumeBassinAnaerobie_m3 = Debit_m3j * md_TempsPassageDebitEntreeAnaerobie_j
    End If
Else
    md_VolumeBassinAnaerobie_m3 = 0
End If
'On simule ce qu'il se passe sur le phosphore


'On calcule la production de boues
ProductionBoues_kgMESj = F_production_specifique_boues_carbone_kgMES_kgDBO(pd_TemperatureEauDimensionnement_°C, md_NominalAgeBoues_j, FluxMES_kgj, FluxDBO_kgj) * FluxDBO_kgj

'On dimensionne les autres bassins
If mi_ChoixConfigurationBassins = mi_BASSINS_SEPARES Then   '#on a des cuves séparées pour la nit et la dénit
    md_VolumeChenal_m3 = 0
    If mn_VolumeBassinsTotal_m3_force = True Then               '## volume total forcé
        md_VolumeBassinsTotal_m3 = md_VolumeBassinsTotal_m3_f
        'On recalcule la concentration dans les bassins au nominal si elle n'a pas été forcée
        If mn_NominalConcentrationBassinsMES_gL_force = False Then
            md_NominalConcentrationBassinsMES_gL = (md_NominalAgeBoues_j * ProductionBoues_kgMESj) / (md_VolumeBassinsTotal_m3 - md_VolumeBassinAnaerobie_m3)
        End If
        If mn_VolumeBassinAnoxie_m3_force = True Then               '### volumes total et anoxique forcés
            md_VolumeBassinAnoxie_m3 = md_VolumeBassinAnoxie_m3_f
            If mn_VolumeBassinAerobie_m3_force = True Then              '#### volumes total, anoxique et aérobie forcés
                md_VolumeBassinAerobie_m3 = md_VolumeBassinAerobie_m3_f
                mn_RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobieFixe = True
                'On vérifie la cohérence des volumes
                If md_VolumeBassinsTotal_m3 <> (md_VolumeBassinAnaerobie_m3 + md_VolumeBassinAnoxie_m3 + md_VolumeBassinAerobie_m3) Then
                    remarque(1) = Feuil13.Range("BA_faible_REPEREerreur_volume").Cells(1, 1)
                Else
                    remarque(1) = ""
                End If
            Else                                                        '#### volumes total et anoxique forcés
                md_VolumeBassinAerobie_m3 = md_VolumeBassinsTotal_m3 - md_VolumeBassinAnoxie_m3 - md_VolumeBassinAnaerobie_m3
                mn_RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobieFixe = True
            End If
        ElseIf mn_VolumeBassinAerobie_m3_force = True Then          '### volumes total et aérobie forcés
            md_VolumeBassinAerobie_m3 = md_VolumeBassinAerobie_m3_f
            md_VolumeBassinAnoxie_m3 = md_VolumeBassinsTotal_m3 - md_VolumeBassinAerobie_m3 - md_VolumeBassinAnaerobie_m3
            mn_RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobieFixe = True
        Else                                                        '### volume total forcé uniquement
            md_VolumeBassinAnoxie_m3 = (md_VolumeBassinsTotal_m3 - md_VolumeBassinAnaerobie_m3) * RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie
            md_VolumeBassinAerobie_m3 = md_VolumeBassinsTotal_m3 - md_VolumeBassinAnoxie_m3 - md_VolumeBassinAnaerobie_m3
        End If
    ElseIf mn_VolumeBassinAnoxie_m3_force = True Then           '## volume anoxique forcé
        md_VolumeBassinAnoxie_m3 = md_VolumeBassinAnoxie_m3_f
        If mn_VolumeBassinAerobie_m3_force = True Then              '### volumes anoxique et aérobie forcés
            md_VolumeBassinAerobie_m3 = md_VolumeBassinAerobie_m3_f
            md_VolumeBassinsTotal_m3 = md_VolumeBassinAnoxie_m3 + md_VolumeBassinAerobie_m3 + md_VolumeBassinAnaerobie_m3
            mn_RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobieFixe = True
            'On recalcule la concentration dans les bassins au nominal si elle n'a pas été forcée
            If mn_NominalConcentrationBassinsMES_gL_force = False Then
                md_NominalConcentrationBassinsMES_gL = (md_NominalAgeBoues_j * ProductionBoues_kgMESj) / (md_VolumeBassinsTotal_m3 - md_VolumeBassinAnaerobie_m3)
            End If
        Else                                                        '### volume anoxique forcé uniquement
            md_VolumeBassinAerobie_m3 = md_VolumeBassinAnoxie_m3 * (1 / RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie - 1)
            md_VolumeBassinsTotal_m3 = md_VolumeBassinAnoxie_m3 + md_VolumeBassinAerobie_m3 + md_VolumeBassinAnaerobie_m3
            'On recalcule la concentration dans les bassins au nominal si elle n'a pas été forcée
            If mn_NominalConcentrationBassinsMES_gL_force = False Then
                md_NominalConcentrationBassinsMES_gL = (md_NominalAgeBoues_j * ProductionBoues_kgMESj) / (md_VolumeBassinsTotal_m3 - md_VolumeBassinAnaerobie_m3)
            End If
        End If
    ElseIf mn_VolumeBassinAerobie_m3_force = True Then          '## volume aérobie forcé uniquement
        md_VolumeBassinAerobie_m3 = md_VolumeBassinAerobie_m3_f
        md_VolumeBassinAnoxie_m3 = RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie / (1 - RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie) * md_VolumeBassinAerobie_m3
        md_VolumeBassinsTotal_m3 = md_VolumeBassinAnoxie_m3 + md_VolumeBassinAerobie_m3 + md_VolumeBassinAnaerobie_m3
        'On recalcule la concentration dans les bassins au nominal si elle n'a pas été forcée
        If mn_NominalConcentrationBassinsMES_gL_force = False Then
            md_NominalConcentrationBassinsMES_gL = (md_NominalAgeBoues_j * ProductionBoues_kgMESj) / (md_VolumeBassinsTotal_m3 - md_VolumeBassinAnaerobie_m3)
        End If
    Else                                                        '## aucun volume forcé
        'On commence par les volumes aérobies et anoxiques
        md_VolumeBassinsTotal_m3 = (md_NominalAgeBoues_j * ProductionBoues_kgMESj) / md_NominalConcentrationBassinsMES_gL
        md_VolumeBassinAnoxie_m3 = RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie * md_VolumeBassinsTotal_m3
        md_VolumeBassinAerobie_m3 = (1 - RatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie) * md_VolumeBassinsTotal_m3
        'On ajoute le volume anaérobie et la zone de contact
        md_VolumeBassinsTotal_m3 = md_VolumeBassinsTotal_m3 + md_VolumeBassinAnaerobie_m3
    End If
Else                                                         '#on a un chenal d'aération
    md_VolumeBassinAnoxie_m3 = 0
    md_VolumeBassinAerobie_m3 = 0
    If mn_VolumeBassinsTotal_m3_force = True Then               '## volume total forcé
        md_VolumeBassinsTotal_m3 = md_VolumeBassinsTotal_m3_f
        'On recalcule la concentration dans les bassins au nominal si elle n'a pas été forcée
        If mn_NominalConcentrationBassinsMES_gL_force = False Then
            md_NominalConcentrationBassinsMES_gL = (md_NominalAgeBoues_j * ProductionBoues_kgMESj) / (md_VolumeBassinsTotal_m3 - md_VolumeBassinAnaerobie_m3)
        End If
        If mn_VolumeChenal_m3_force = True Then                     '### volumes total et chenal forcés
            md_VolumeChenal_m3 = md_VolumeChenal_m3_f
            'On vérifie la cohérence des volumes
            If md_VolumeBassinsTotal_m3 <> (md_VolumeBassinAnaerobie_m3 + md_VolumeChenal_m3) Then
                remarque(1) = Feuil13.Range("BA_faible_REPEREerreur_volume").Cells(1, 1)
            Else
                remarque(1) = ""
            End If
        Else
            md_VolumeChenal_m3 = (md_VolumeBassinsTotal_m3 - md_VolumeBassinAnaerobie_m3)
        End If
    ElseIf mn_VolumeChenal_m3_force = True Then                  '## volume chenal forcé
        md_VolumeChenal_m3 = md_VolumeChenal_m3_f
        md_VolumeBassinsTotal_m3 = md_VolumeChenal_m3 + md_VolumeBassinAnaerobie_m3
        'On recalcule la concentration dans les bassins au nominal si elle n'a pas été forcée
        If mn_NominalConcentrationBassinsMES_gL_force = False Then
            md_NominalConcentrationBassinsMES_gL = (md_NominalAgeBoues_j * ProductionBoues_kgMESj) / (md_VolumeBassinsTotal_m3 - md_VolumeBassinAnaerobie_m3)
        End If
    Else
        md_VolumeChenal_m3 = (md_NominalAgeBoues_j * ProductionBoues_kgMESj) / md_NominalConcentrationBassinsMES_gL
        md_VolumeBassinsTotal_m3 = md_VolumeChenal_m3 + md_VolumeBassinAnaerobie_m3
    End If
End If

'On calcule les besoins en oxygène
md_BesoinsOxygene_kgO2j = F_consommation_specifique_O2_respiration_carbone_kgO2_kgDBO(pd_TemperatureEauDimensionnement_°C, md_NominalAgeBoues_j) * DBOeliminee_kgj
md_BesoinsOxygene_kgO2j = md_BesoinsOxygene_kgO2j + md_ConsommationSpecifiqueO2Nitrification_kgO2_kgNNitrifie * Nnitrifie_kgj
md_BesoinsOxygene_kgO2j = md_BesoinsOxygene_kgO2j - md_ConsommationSpecifiqueDCODenitrification_kgDCO_kgN * NDenitrifie_kgj



'On estime les concentrations en sortie
If mn_ConcentrationSortieDBO_mgL_force = True Then
    md_ConcentrationSortieDBO_mgL = md_ConcentrationSortieDBO_mgL_f
Else
    md_ConcentrationSortieDBO_mgL = FluxDBO_kgj * (1 - F_RendementEliminationDBO(pd_TemperatureEauDimensionnement_°C, FluxDBO_kgj, md_NominalConcentrationBassinsMES_gL)) / Debit_m3j * 1000
End If
If md_NominalAgeBoues_j >= md_AgeBouesCarboneSeul12°C_j Then
    If mn_ConcentrationSortieNH4_mgL_force = True Then
        md_ConcentrationSortieNH4_mgL = md_ConcentrationSortieNH4_mgL_f
    Else
        md_ConcentrationSortieNH4_mgL = md_HypotheseConcentrationSortieNH4_mgL
    End If
Else
    If mn_ConcentrationSortieNH4_mgL_force = True Then
        md_ConcentrationSortieNH4_mgL = md_ConcentrationSortieNH4_mgL_f
    Else
        md_ConcentrationSortieNH4_mgL = md_HypotheseConcentrationSortieNH4_mgL
    End If
    
End If


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Sub fonctionnement_reel(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef Sh, ByRef DCO_apportee_methanol, ByRef DCO_conso_bio)
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'variables internes
    Dim G_equivalent As Double
    Dim Cm_equivalente As Double
    Dim Cm As Double
    Dim boues_bassins_MVS As Double
    Dim boues_bassins_MES As Double
    Dim rdt_DBO As Double
    Dim boues_produites As Double
    Dim P_synthese As Double
    Dim P_bio_coeff_K As Double
    Dim volume_optimal As Double
    Dim Kp_bio_1 As Double
    Dim Kp_bio_2 As Double
    Dim ratio_DBO_P_in As Double
    Dim dose_MMS As Double
    Dim P_precipitation As Double
    Dim ratio_molaire_Ptbr As Double
    Dim I As Integer
    Dim nitrification As Double
    Dim denitrification As Double
    Dim rdt_denitrification As Double
    Dim DCO_soluble As Double
    Dim ratio_C_N As Double
    Dim DCO_necessaire As Double
    Dim facteur_T_O2 As Double
    Dim ratio_O2_DBO_respiration As Double
    Dim delta_O2 As Double
    Dim rdt_denit As Double
    Dim Ptbr As Double
    Dim Psol_inlet As Double

    stockage_Q = Q


    'FONCTIONNEMENT REEL

    'Age de boues equivalent à md_TemperatureReferenceAgeBouesEquivalent_°C
    G_equivalent = md_ReelAgeBoues_j * md_CoefficientCalculAgeBouesEquivalent ^ (pd_TemperatureEauConditionsReelles_°C - md_TemperatureReferenceAgeBouesEquivalent_°C)
    'Charge massique équivalente
    Cm_equivalente = 1 / (a1_Cm * G_equivalent + a0_Cm)
    'Charge massique à T_design
    Cm = Cm_equivalente * md_CoefficientCalculAgeBouesEquivalent ^ (pd_TemperatureEauConditionsReelles_°C - md_TemperatureReferenceAgeBouesEquivalent_°C)
    'Masse de boues dans les bassins kgMVS
    boues_bassins_MVS = DBO / Cm
    'on recalcule la part de MVS dans les MES
    boues_bassins_MES = boues_bassins_MVS / MV_MES_reference + (boues_chimiques + boues_methanol) * md_ReelAgeBoues_j
    pd_ReelRatioMVMES = (boues_bassins_MVS + md_ReelAgeBoues_j * boues_methanol) / boues_bassins_MES
    If mn_ReelRatioMVMES_force = True Then
        pd_ReelRatioMVMES = md_ReelRatioMVMES_f
    End If
    'Calcul de la concentration en MES dans les bassins
    md_ReelConcentrationBassinsMES_gL = boues_bassins_MVS / (pd_ReelRatioMVMES * md_VolumeBassinsTotal_m3)
    If mn_ReelConcentrationBassinsMES_gL_force = True Then
        md_ReelConcentrationBassinsMES_gL = md_ReelConcentrationBassinsMES_gL_f
        'calcul de la Cm équivalente avec cette concentration
        boues_bassins_MVS = md_VolumeBassinsTotal_m3 * pd_ReelRatioMVMES * md_ReelConcentrationBassinsMES_gL
        Cm = DBO / boues_bassins_MVS
        Cm_equivalente = Cm / md_CoefficientCalculAgeBouesEquivalent ^ (pd_TemperatureEauConditionsReelles_°C - md_TemperatureReferenceAgeBouesEquivalent_°C)
    End If

    stockage_MVS_bassins = boues_bassins_MVS

    'TRAITEMENT REEL

    'rendement d'élimination de la DBO
    rdt_DBO = 0
    For I = 0 To 3
        rdt_DBO = rdt_DBO + md_CoefficientPolynomeCalculRendementDBO(I) * (Cm_equivalente / pd_ReelRatioMVMES) ^ I
    Next I
    md_ConcentrationSortieDBO_mgL = DBO * (1 - rdt_DBO) / Q * 1000
    If mn_ConcentrationSortieDBO_mgL_force = True Then
        md_ConcentrationSortieDBO_mgL = md_ConcentrationSortieDBO_mgL_f
    End If

    'traitement de l'azote
    nitrification = NH4 - md_ConcentrationSortieNH4_mgL * Q / 1000
    denitrification = NO3 + nitrification - md_ConcentrationSortieNO3_mgL * Q / 1000
    'calcul de la recirculation de liqueur mixte approppriée à ce traitement si cuves séparées
    If mi_ChoixConfigurationBassins = mi_BASSINS_SEPARES Then
        'VU avec Kim Sorensen le 13/05/2011
        'rdt_denit = ((NK + NO3) - (md_ConcentrationSortieNH4_mgL + md_ConcentrationSortieNO3_mgL + 1) * Q / 1000) / (NK + NO3)  '1 correspond à l'azote réduit autre que NH4
        'md_TauxRecirculationMLSS = rdt_denit / (1 - rdt_denit)
        md_TauxRecirculationMLSS = nitrification * 1000 / (Q * md_ConcentrationSortieNO3_mgL) - (1 + md_TauxRecirculationBouesClarificateur)
        If mn_TauxRecirculationMLSS_force = True Then
            If md_TauxRecirculationMLSS > md_TauxRecirculationMLSS_f Then
                remarque(4) = Feuil13.Range("BA_faible_REPEREerreur_taux_recirc_MLSS").Cells(1, 1)
            Else
                remarque(4) = ""
            End If
            md_TauxRecirculationMLSS = md_TauxRecirculationMLSS_f
        End If
    Else
        md_TauxRecirculationMLSS = 0
    End If
    'calcul de l'ajout de méthanol pour la dénitrification
    DCO_soluble = DCO_nominal * pd_PourcentageChargeReelleDCO * ratio_DCO_sol_EB * (1 - rdt_DCO_sol_primaire)
    ratio_C_N = DCO_soluble / denitrification
    If ratio_C_N < md_ConsommationDCOexterneDenitrification_kgDCO_kgNDenitrifie Then
        DCO_necessaire = md_ConsommationDCOexterneDenitrification_kgDCO_kgNDenitrifie * denitrification - DCO_soluble
        pd_ConsommationMethanolPur_kgj = DCO_necessaire / md_DCOdansMethanol_kgDCO_kg
    End If
    If mn_ConsommationMethanolPur_kgj_force = True Then
        pd_ConsommationMethanolPur_kgj = md_ConsommationMethanolPur_kgj_f
    End If
    DCO_necessaire = pd_ConsommationMethanolPur_kgj * md_DCOdansMethanol_kgDCO_kg
    DCO_apportee_methanol = DCO_necessaire
    boues_methanol = ratio_boues_methanol * DCO_necessaire
    DCO = DCO + DCO_necessaire * (1 - ratio_boues_methanol)

    'traitement du phosphore
    P_synthese = ratio_P_synthese * DBO * rdt_DBO
    If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE Then
        If mn_RendementDephosphatationBiologique_force = True Then
            md_RendementDephosphatationBiologique = md_RendementDephosphatationBiologique_f
        Else
            P_bio_coeff_K = a_Pbio_coef_K + b_Pbio_coef_K / G_equivalent
            ratio_DBO_P_in = DBO / Pt
            volume_optimal = Q / md_TempsPassageDebitEntreeAnaerobie_j * md_CoefficientCalculAgeBouesEquivalent ^ (md_TemperatureReferenceAgeBouesEquivalent_°C - pd_TemperatureEauConditionsReelles_°C)
            Kp_bio_1 = md_VolumeBassinAnaerobie_m3 / volume_optimal
            If Kp_bio_1 <= 1 Then
                Kp_bio_1 = Kp_bio_1 ^ 2
            ElseIf Kp_bio_1 <= 1.5 Then
                Kp_bio_1 = a_Kp_bio_1 + b_Kp_bio_1 * Kp_bio_1
            Else
                Kp_bio_1 = lim_Kp_bio_1
            End If
            Kp_bio_2 = 1
            md_RendementDephosphatationBiologique = P_bio_coeff_K * Kp_bio_1 * Kp_bio_2 * ratio_DBO_P_in
            If md_RendementDephosphatationBiologique > rdt_Psol_max Then
                md_RendementDephosphatationBiologique = rdt_Psol_max
            End If
        End If
        Pt = Pt - md_RendementDephosphatationBiologique * ratio_Psol_P * Pt
        md_ConcentrationSortiePt_mgL = Pt / (Q - pd_BouesExtraitesDebit_m3j) * 1000
        boues_chimiques = 0
        P_precipite = 0
    ElseIf mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE_ET_PRECIPITATION Then
        If mn_RendementDephosphatationBiologique_force = True Then
            md_RendementDephosphatationBiologique = md_RendementDephosphatationBiologique_f
        Else
            P_bio_coeff_K = a_Pbio_coef_K + b_Pbio_coef_K / G_equivalent
            ratio_DBO_P_in = DBO / Pt
            volume_optimal = Q / md_TempsPassageDebitEntreeAnaerobie_j * md_CoefficientCalculAgeBouesEquivalent ^ (md_TemperatureReferenceAgeBouesEquivalent_°C - pd_TemperatureEauConditionsReelles_°C)
            Kp_bio_1 = md_VolumeBassinAnaerobie_m3 / volume_optimal
            If Kp_bio_1 <= 1 Then
                Kp_bio_1 = Kp_bio_1 ^ 2
            ElseIf Kp_bio_1 <= 1.5 Then
                Kp_bio_1 = a_Kp_bio_1 + b_Kp_bio_1 * Kp_bio_1
            Else
                Kp_bio_1 = lim_Kp_bio_1
            End If
            Kp_bio_2 = 1
            md_RendementDephosphatationBiologique = P_bio_coeff_K * Kp_bio_1 * Kp_bio_2 * ratio_DBO_P_in
            If md_RendementDephosphatationBiologique > rdt_Psol_max Then
                md_RendementDephosphatationBiologique = rdt_Psol_max
            End If
        End If
        Pt = Pt - md_RendementDephosphatationBiologique * ratio_Psol_P * Pt
        Pt = Pt - P_synthese
        P_precipitation = Pt - md_ConcentrationSortiePt_mgL * (Q - pd_BouesExtraitesDebit_m3j) / 1000
        If P_precipitation > 0 Then
            Ptbr = P_precipitation / Q * 1000
            ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + md_ConcentrationSortiePt_mgL) ^ (-0.5) - ((md_ConcentrationSortiePt_mgL) ^ (-0.5) - 0.1)) / ((Ptbr + md_ConcentrationSortiePt_mgL) - (md_ConcentrationSortiePt_mgL - 0.1))
            pd_ConsommationFeCl3Pur_kgj = ratio_molaire_Ptbr * ratio_FeCl3_Pt * P_precipitation
        Else
            P_precipitation = 0
            pd_ConsommationFeCl3Pur_kgj = 0
        End If
        dose_MMS = (ratio_molaire_Ptbr - 1) * a_dosis_MMS + b_dosis_MMS
        boues_chimiques = dose_MMS * P_precipitation
        Pt = Pt - P_precipitation
        P_precipite = P_precipitation
    ElseIf mi_ChoixDephosphatation = mi_DEPHOSPHATATION_PRECIPITATION Then
        Pt = Pt - P_synthese
        P_precipitation = Pt - md_ConcentrationSortiePt_mgL * (Q - pd_BouesExtraitesDebit_m3j) / 1000
        If P_precipitation > 0 Then
            Ptbr = P_precipitation / Q * 1000
            ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + md_ConcentrationSortiePt_mgL) ^ (-0.5) - ((md_ConcentrationSortiePt_mgL) ^ (-0.5) - 0.1)) / ((Ptbr + md_ConcentrationSortiePt_mgL) - (md_ConcentrationSortiePt_mgL - 0.1))
            pd_ConsommationFeCl3Pur_kgj = ratio_molaire_Ptbr * ratio_FeCl3_Pt * P_precipitation
        Else
            P_precipitation = 0
            pd_ConsommationFeCl3Pur_kgj = 0
        End If
        dose_MMS = (ratio_molaire_Ptbr - 1) * a_dosis_MMS + b_dosis_MMS
        boues_chimiques = dose_MMS * P_precipitation
        Pt = Pt - P_precipitation
        P_precipite = P_precipitation
    Else
        Pt = Pt - P_synthese
        boues_chimiques = 0
        P_precipite = 0
    End If

    'ratio O2 respiration et élimination DBO5
    facteur_T_O2 = md_CoefficientAFacteurImpactTemperature ^ (pd_TemperatureEauConditionsReelles_°C - md_CoefficientBFacteurImpactTemperature_°C)
    ratio_O2_DBO_respiration = md_CoefficientAConsommationSpecifiqueO2RespirationCarbone_kgO2_kgDBO + (md_CoefficientBConsommationSpecifiqueO2RespirationCarbone_kgO2_kgDBOj * md_ReelAgeBoues_j * facteur_T_O2) / (1 + md_CoefficientCConsommationSpecifiqueO2RespirationCarbone * md_ReelAgeBoues_j * facteur_T_O2)

    'calcul des besoins en O2
    besoins_O2_respiration = ((md_CoefficientBConsommationSpecifiqueO2RespirationCarbone_kgO2_kgDBOj * md_ReelAgeBoues_j * facteur_T_O2) / (1 + md_CoefficientCConsommationSpecifiqueO2RespirationCarbone * md_ReelAgeBoues_j * facteur_T_O2)) * (DBO - md_ConcentrationSortieDBO_mgL * stockage_Q / 1000)
    md_BesoinsOxygene_kgO2j = ratio_O2_DBO_respiration * (DBO - md_ConcentrationSortieDBO_mgL * stockage_Q / 1000)
    md_BesoinsOxygene_kgO2j = md_BesoinsOxygene_kgO2j + md_ConsommationSpecifiqueO2Nitrification_kgO2_kgNNitrifie * nitrification
    md_BesoinsOxygene_kgO2j = md_BesoinsOxygene_kgO2j + md_ConsommationSpecifiqueDCODenitrification_kgDCO_kgN * denitrification
    md_BesoinsOxygene_kgO2j = md_BesoinsOxygene_kgO2j + pd_ConsommationMethanolPur_kgj * md_DCOdansMethanol_kgDCO_kg  'on ne consomme pas le méthanol directement mais il représente de la DCO non consommée par la dénit
    'DCO_conso_bio = ratio_O2_DBO_respiration * (DBO - md_ConcentrationSortieDBO_mgL * stockage_Q / 1000) + DCO_apportee_methanol * (1 - ratio_boues_methanol) * rdt_DCO - md_ConsommationSpecifiqueDCODenitrification_kgDCO_kgN * denitrification
    If mn_BesoinsOxygene_kgO2j_force = True Then
        delta_O2 = md_BesoinsOxygene_kgO2j_f - md_BesoinsOxygene_kgO2j
        md_BesoinsOxygene_kgO2j = md_BesoinsOxygene_kgO2j_f
        'DCO_conso_bio = DCO_conso_bio + delta_O2
    End If



    'Production de boues
    boues_produites = md_ReelConcentrationBassinsMES_gL * md_VolumeBassinsTotal_m3 / md_ReelAgeBoues_j


    'CALCUL DES SORTIES

    'Eau traitée
    Q = Q - pd_BouesExtraitesDebit_m3j
    MES = md_EauClarifieeMES_mgL * Q / 1000
    DCO = DCO * (1 - rdt_DCO)
    DBO = md_ConcentrationSortieDBO_mgL * Q / 1000
    NH4 = md_ConcentrationSortieNH4_mgL * Q / 1000
    NK = NH4
    NO3 = md_ConcentrationSortieNO3_mgL * Q / 1000

    'Boues
    pd_BouesExtraitesFluxMES_kgj = boues_produites - MES
    If mn_BouesExtraitesFluxMES_kgj_force = True Then
        pd_BouesExtraitesFluxMES_kgj = pd_BouesExtraitesFluxMES_kgj_f
    End If
    md_BouesExtraitesConcentrationMES_gL = md_ReelConcentrationBassinsMES_gL * (1 + md_TauxRecirculationBouesClarificateur) / md_TauxRecirculationBouesClarificateur
    If mn_BouesExtraitesConcentrationMES_gL_force = True Then
        md_BouesExtraitesConcentrationMES_gL = md_BouesExtraitesConcentrationMES_gL_f
    End If
    pd_BouesExtraitesDebit_m3j = pd_BouesExtraitesFluxMES_kgj / md_BouesExtraitesConcentrationMES_gL

    If md_BouesExtraitesConcentrationMES_gL > facteur_boues_C_lim / md_BouesISV_mL_g Then
        remarque(3) = Feuil13.Range("BA_faible_REPEREerreur_taux_recirc").Cells(1, 1)
        If mn_TauxRecirculationBouesClarificateur_force = False Then
            md_TauxRecirculationBouesClarificateur = md_TauxRecirculationBouesClarificateur + 0.1
        End If
    Else
        remarque(3) = ""
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "13.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Function coagulation_floculation_Ptbr(ByVal Q, ByVal Pt, ByVal dosage_FeCl3)

    Const precipitation_nb_iteration = 30

    Dim x_min As Double
    Dim x_max As Double
    Dim Psol_in As Double
    Dim Psol_out As Double
    Dim dose_FeCl3_dichotomie As Double
    Dim Ptbr As Double
    Dim ratio_molaire_Ptbr As Double

    dosage_FeCl3 = dosage_FeCl3 / Q * 1000

    If dosage_FeCl3 > 0 Then
        Psol_in = Pt / Q * 1000    '* ratio_Psol_P
        x_min = 0.1
        x_max = Psol_in
        Ptbr = (x_min + x_max) / 2
        Do While (x_max - x_min) / Ptbr > 0.00001
            Psol_out = Psol_in - Ptbr
            ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + Psol_out) ^ (-0.5) - (Psol_out ^ (-0.5) - 0.1)) / ((Ptbr + Psol_out) - (Psol_out - 0.1))
            dose_FeCl3_dichotomie = Ptbr * ratio_molaire_Ptbr * ratio_FeCl3_Pt
            If dose_FeCl3_dichotomie < dosage_FeCl3 Then
                x_min = Ptbr
            Else
                x_max = Ptbr
            End If
            Ptbr = (x_min + x_max) / 2
        Loop
        coagulation_floculation_Ptbr = Ptbr * Q / 1000
    Else
        coagulation_floculation_Ptbr = 0
    End If


End Function


Sub calcul_consommation_electrique(ByRef pd_ConsommationElectrique_kWhj_II_aeration)
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer

    Dim clarif_diametre_limite As Double

    'variables pour l'aération
    Dim gamme_P As Integer
    Dim gamme_Q As Integer
    Dim Patm As Double
    Dim delta_P As Double
    Dim O2_saturation_20_P As Double
    Dim O2_saturation_T_P As Double
    Dim O2_facteur_K As Double
    Dim surpresseur_puissance As Double
    Dim pd_ConsommationElectrique_kWhj_racleur As Double
    Dim pd_ConsommationElectrique_kWhj_agitation As Double
    Dim pd_ConsommationElectrique_kWhj_recirculation_MLSS As Double
    Dim pd_ConsommationElectrique_kWhj_recirculation As Double
    Dim pd_ConsommationElectrique_kWhj_extraction As Double


    'AERATION

    'calcul facteur alpha
    If mi_ChoixTypeAeration = mi_AERATEUR_FINES_BULLES Then
        'source: Oxygen transfer in AS (Henkel,Cornel and Wagner)
        md_FacteurAlpha = a0_alpha + a_alpha_MVS * (stockage_MVS_bassins / md_VolumeBassinsTotal_m3) + a_alpha_G * (md_ReelAgeBoues_j * md_CoefficientCalculAgeBouesEquivalent ^ (pd_TemperatureEauConditionsReelles_°C - md_TemperatureReferenceAgeBouesEquivalent_°C))
    Else
        md_FacteurAlpha = O2_alpha0(mi_ChoixTypeAeration)
        If md_ReelConcentrationBassinsMES_gL > MES_reference_alfa Then
            md_FacteurAlpha = md_FacteurAlpha * (a0_correctif_alfa_MES + a1_correctif_alfa_MES * md_ReelConcentrationBassinsMES_gL) / (a0_correctif_alfa_MES + a1_correctif_alfa_MES * MES_reference_alfa)
        End If
    End If
    If mn_FacteurAlpha_force = True Then
        md_FacteurAlpha = md_FacteurAlpha_f
    End If

    'calcul O2 saturation à (T,P=Patm+delta_P) et (20,P0)
    Patm = P0 * ((a0_Patm + a1_Patm * altitude) ^ a2_Patm) / 100    'en hPa
    If mi_ChoixTypeAeration = mi_AERATEUR_FINES_BULLES Or mi_ChoixTypeAeration = mi_AERATEUR_MOYENNES_BULLES Then
        delta_P = a0_delta_P / a1_delta_P * P0 / 100 * (md_HauteurBassinAeration_m - hauteur_diffuseur)
    Else
        delta_P = 0
    End If
    O2_saturation_20_P = 0
    O2_saturation_T_P = 0
    For I = 0 To 4
        O2_saturation_20_P = O2_saturation_20_P + a_O2_saturation(I) * 20 ^ I
        O2_saturation_T_P = O2_saturation_T_P + a_O2_saturation(I) * pd_TemperatureEauConditionsReelles_°C ^ I
    Next I
    O2_saturation_T_P = O2_saturation_T_P * (Patm + delta_P) / (P0 / 100)
    O2_saturation_20_P = O2_saturation_20_P * (Patm + delta_P) / (P0 / 100)

    'calcul facteur K
    O2_facteur_K = md_FacteurAlpha * (O2_facteur_beta * O2_saturation_T_P - md_ConcentrationO2dissous_mgL) / O2_saturation_20_P
    O2_facteur_K = O2_facteur_K * md_CoefficientCalculAgeBouesEquivalent_K ^ (pd_TemperatureEauConditionsReelles_°C - T_ref_K)

    'calcul consommations sur l'aération
    If mi_ChoixTypeAeration = mi_AERATEUR_FINES_BULLES Or mi_ChoixTypeAeration = mi_AERATEUR_MOYENNES_BULLES Then
        'débit d'air
        md_RendementTransfertOxygeneMembraneEauClaire_1m = O2_rdt_dissolution(mi_ChoixTypeAeration)
        If mn_RendementTransfertOxygeneMembraneEauClaire_1m_force = True Then
            md_RendementTransfertOxygeneMembraneEauClaire_1m = md_RendementTransfertOxygeneMembraneEauClaire_1m_f
        End If
        md_DebitAir_Nm3j = md_BesoinsOxygene_kgO2j / (O2_facteur_K * md_RendementTransfertOxygeneMembraneEauClaire_1m / 100 * (md_HauteurBassinAeration_m - hauteur_diffuseur) * ratio_kgO2_Nm3air)
        If mn_DebitAir_Nm3j_force = True Then
            md_DebitAir_Nm3j = md_DebitAir_Nm3j_f
        End If
        If mn_ConsommationSpecifiqueSurpresseur_Wh_Nm3mCE_force = True Then
            md_ConsommationSpecifiqueSurpresseur_Wh_Nm3mCE = md_ConsommationSpecifiqueSurpresseur_Wh_Nm3mCE_f
            'gamme de pression
            If md_PressionAirSurpresse_mCE <= Px1 Then
                gamme_P = surpresseur_gamme_P_0_Px1
            ElseIf md_PressionAirSurpresse_mCE <= Px2 Then
                gamme_P = surpresseur_gamme_P_Px1_Px2
            Else
                gamme_P = surpresseur_gamme_P_Px2_infini
                If mi_ChoixTypeSurpresseur = mi_SURPRESSEUR_LOBES Then
                    remarque(5) = Feuil13.Range("BA_faible_REPEREerreur_mi_ChoixTypeSurpresseur").Cells(1, 1)
                End If
            End If
        Else
            'gamme de pression
            If md_PressionAirSurpresse_mCE <= Px1 Then
                gamme_P = surpresseur_gamme_P_0_Px1
            ElseIf md_PressionAirSurpresse_mCE <= Px2 Then
                gamme_P = surpresseur_gamme_P_Px1_Px2
            Else
                gamme_P = surpresseur_gamme_P_Px2_infini
                If mi_ChoixTypeSurpresseur = mi_SURPRESSEUR_LOBES Then
                    remarque(5) = Feuil13.Range("BA_faible_REPEREerreur_mi_ChoixTypeSurpresseur").Cells(1, 1)
                End If
            End If
            'gamme de débit   (pour le nominal)
            gamme_Q = surpresseur_gamme_Q_0_infini
            md_ConsommationSpecifiqueSurpresseur_Wh_Nm3mCE = surpresseur_conso_Wh_Nm3mCE(mi_ChoixTypeSurpresseur, gamme_Q, gamme_P) * (1 + impact_regulation(mi_ChoixRegulationAeration))
        End If
        'calcul de la consommation électrique de l'aération
        md_ConsommationElectriqueAeration_kWhj = md_DebitAir_Nm3j * md_PressionAirSurpresse_mCE * md_ConsommationSpecifiqueSurpresseur_Wh_Nm3mCE / 1000
    Else
        'calcul de la consommation électrique de l'aération
        md_ConsommationElectriqueAeration_kWhj = md_BesoinsOxygene_kgO2j / (md_ApportSpecifiqueBrutEauClaire_kgO2_kWh * O2_facteur_K)
    End If
    pd_ConsommationElectrique_kWhj_II_aeration = pd_ConsommationElectrique_kWhj_II_aeration + md_ConsommationElectriqueAeration_kWhj


    'CLARIFICATEUR
    'pd_ConsommationElectrique_kWhj_racleur = rate_puissance_fixe * pd_NOMBRE_HEURE_PAR_JOUR + rate_puissance_debit * stockage_Q / 1000
    'DCA le 15/07/11 données Pascal JOUAFFRE  de 12 à 32m diamètre --> 0.55 kW installés, de 33 à 52m --> 0.75 kW
    clarif_diametre_limite = 32
    If md_SurfaceTotaleClarification_m2 < pd_NOMBRE_PI * (clarif_diametre_limite / 2) ^ 2 Then
        pd_ConsommationElectrique_kWhj_racleur = 0.55 * pd_NOMBRE_HEURE_PAR_JOUR
    Else
        pd_ConsommationElectrique_kWhj_racleur = 0.75 * pd_NOMBRE_HEURE_PAR_JOUR
    End If

    'AGITATION
    If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE Or mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE_ET_PRECIPITATION Then
        pd_ConsommationElectrique_kWhj_agitation = md_PuissanceAgitationAnaerobie_Wm3 * md_VolumeBassinAnaerobie_m3 * agitation_fct / 1000
        If mi_ChoixConfigurationBassins = mi_BASSINS_SEPARES Then
            pd_ConsommationElectrique_kWhj_agitation = pd_ConsommationElectrique_kWhj_agitation + md_PuissanceAgitationAnoxie_Wm3 * md_VolumeBassinAnoxie_m3 * agitation_fct / 1000
            pd_ConsommationElectrique_kWhj_agitation = pd_ConsommationElectrique_kWhj_agitation + md_PuissanceAgitationAerobie_Wm3 * md_VolumeBassinAerobie_m3 * agitation_fct / 1000
        Else
            pd_ConsommationElectrique_kWhj_agitation = pd_ConsommationElectrique_kWhj_agitation + md_PuissanceAgitationChenal_Wm3 * md_VolumeChenal_m3 * agitation_fct / 1000
            'pd_ConsommationElectrique_kWhj_agitation = pd_ConsommationElectrique_kWhj_agitation + md_PuissanceAgitationAnoxie_Wm3 * (md_VolumeBassinsTotal_m3 - md_VolumeBassinAnaerobie_m3 - md_VolumeChenal_m3) * agitation_fct / 1000
        End If
    Else
        If mi_ChoixConfigurationBassins = mi_BASSINS_SEPARES Then
            pd_ConsommationElectrique_kWhj_agitation = md_PuissanceAgitationAnoxie_Wm3 * md_VolumeBassinAnoxie_m3 * agitation_fct / 1000
            pd_ConsommationElectrique_kWhj_agitation = pd_ConsommationElectrique_kWhj_agitation + md_PuissanceAgitationAerobie_Wm3 * md_VolumeBassinAerobie_m3 * agitation_fct / 1000
        Else
            pd_ConsommationElectrique_kWhj_agitation = md_PuissanceAgitationChenal_Wm3 * md_VolumeChenal_m3 * agitation_fct / 1000
            'pd_ConsommationElectrique_kWhj_agitation = pd_ConsommationElectrique_kWhj_agitation + md_PuissanceAgitationAnoxie_Wm3 * (md_VolumeBassinsTotal_m3 - md_VolumeChenal_m3) * agitation_fct / 1000
        End If
    End If


    'POMPES DE RECIRCULATION ET EXTRACTION
    pd_ConsommationElectrique_kWhj_recirculation_MLSS = ratio_elec_recirculation_MLSS * md_TauxRecirculationMLSS * stockage_Q * md_PompesRecirculationMLSSPression_mCE
    pd_ConsommationElectrique_kWhj_recirculation = ratio_elec_recirculation * md_TauxRecirculationBouesClarificateur * stockage_Q * md_PompesRecirculationBouesPression_mCE
    pd_ConsommationElectrique_kWhj_extraction = ratio_elec_extraction * pd_BouesExtraitesDebit_m3j * md_PompesBouesExcesPression


    'CONSOMMATION GLOBALE D'pd_ConsommationElectrique_kWhj
    pd_ConsommationElectrique_kWhj = md_ConsommationElectriqueAeration_kWhj + pd_ConsommationElectrique_kWhj_racleur + pd_ConsommationElectrique_kWhj_agitation + pd_ConsommationElectrique_kWhj_recirculation_MLSS + pd_ConsommationElectrique_kWhj_recirculation + pd_ConsommationElectrique_kWhj_extraction
    pd_ConsommationElectrique_kWhj_fixe = besoins_O2_respiration / md_BesoinsOxygene_kgO2j * md_ConsommationElectriqueAeration_kWhj + pd_ConsommationElectrique_kWhj_racleur + pd_ConsommationElectrique_kWhj_agitation

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "13.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub lecture_choix()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    choix_fct_SHUNT = Feuil57.Range("choix_fct_SHUNT").Cells(1, pi_FiliereConsideree)

    'lecture des choix internes
    mi_ChoixDephosphatation = Feuil13.Range("BA_faible_REPEREchoix_dephosphatation").Cells(1, pi_FiliereConsideree)
    mi_ChoixConfigurationBassins = Feuil13.Range("BA_faible_REPEREchoix_configuration").Cells(1, pi_FiliereConsideree)
    mi_ChoixTypeAeration = Feuil13.Range("BA_faible_REPEREchoix_aerateur").Cells(1, pi_FiliereConsideree)
    mi_ChoixRegulationAeration = Feuil13.Range("BA_faible_REPEREchoix_regulation").Cells(1, pi_FiliereConsideree)
    mi_ChoixTypeSurpresseur = Feuil13.Range("BA_faible_REPEREchoix_surpresseur").Cells(1, pi_FiliereConsideree)
    mi_ChoixTypeRacleur = Feuil13.Range("BA_faible_REPEREmi_ChoixTypeRacleur").Cells(1, pi_FiliereConsideree)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "13.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub lecture_valeurs_forcees()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    'lecture du choix de forcer des valeurs
    mn_NominalAgeBoues_j_force = Feuil13.Range("BA_faible_REPEREnominal_age_boue").Cells(1, pi_FiliereConsideree)
    mn_NominalConcentrationBassinsMES_gL_force = Feuil13.Range("BA_faible_REPEREmd_NominalConcentrationBassinsMES_gL").Cells(1, pi_FiliereConsideree)
    mn_NominalRatioMVMES_force = Feuil13.Range("BA_faible_REPEREnominal_MV_MES").Cells(1, pi_FiliereConsideree)
    mn_VolumeBassinsTotal_m3_force = Feuil13.Range("BA_faible_REPEREvolume_bassins").Cells(1, pi_FiliereConsideree)
    If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE Or mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE_ET_PRECIPITATION Then
        mn_VolumeBassinAnaerobie_m3_force = Feuil13.Range("BA_faible_REPEREvolume_anaerobie").Cells(1, pi_FiliereConsideree)
        mn_PuissanceAgitationAnaerobie_Wm3_force = Feuil13.Range("BA_faible_REPEREratio_elec_anaerobie").Cells(1, pi_FiliereConsideree)
        mn_RendementDephosphatationBiologique_force = Feuil13.Range("BA_faible_REPERErendement_bioP").Cells(1, pi_FiliereConsideree)
    End If
    If mi_ChoixConfigurationBassins = mi_BASSINS_SEPARES Then
        mn_VolumeBassinAnoxie_m3_force = Feuil13.Range("BA_faible_REPEREvolume_anoxie").Cells(1, pi_FiliereConsideree)
        mn_PuissanceAgitationAnoxie_Wm3_force = Feuil13.Range("BA_faible_REPEREratio_elec_anoxie").Cells(1, pi_FiliereConsideree)
        mn_VolumeBassinAerobie_m3_force = Feuil13.Range("BA_faible_REPEREvolume_aerobie").Cells(1, pi_FiliereConsideree)
        mn_PuissanceAgitationAerobie_Wm3_force = Feuil13.Range("BA_faible_REPEREratio_elec_aerobie").Cells(1, pi_FiliereConsideree)
        mn_TauxRecirculationMLSS_force = Feuil13.Range("BA_faible_REPERErecirculation_MLSS_taux").Cells(1, pi_FiliereConsideree)
        mn_PompesRecirculationMLSSPression_mCE_force = Feuil13.Range("BA_faible_REPERErecirculation_MLSS_P_refoulement").Cells(1, pi_FiliereConsideree)
        mn_PompesRecirculationMLSSRendement_force = Feuil13.Range("BA_faible_REPERErecirculation_MLSS_pompe_rdt").Cells(1, pi_FiliereConsideree)
    Else
        mn_VolumeChenal_m3_force = Feuil13.Range("BA_faible_REPEREvolume_chenal").Cells(1, pi_FiliereConsideree)
        mn_PuissanceAgitationChenal_Wm3_force = Feuil13.Range("BA_faible_REPEREratio_elec_chenal").Cells(1, pi_FiliereConsideree)
    End If
    mn_ReelAgeBoues_j_force = Feuil13.Range("BA_faible_REPEREreel_age_boue").Cells(1, pi_FiliereConsideree)
    mn_ReelRatioMVMES_force = Feuil13.Range("BA_faible_REPEREreel_MV_MES").Cells(1, pi_FiliereConsideree)
    mn_ReelConcentrationBassinsMES_gL_force = Feuil13.Range("BA_faible_REPEREreel_MES_concentration_bassin").Cells(1, pi_FiliereConsideree)
    mn_ConcentrationO2dissous_mgL_force = Feuil13.Range("BA_faible_REPEREO2_dissous").Cells(1, pi_FiliereConsideree)
    mn_ConcentrationSortieDBO_mgL_force = Feuil13.Range("BA_faible_REPEREsortie_DBO").Cells(1, pi_FiliereConsideree)
    mn_ConcentrationSortieNH4_mgL_force = Feuil13.Range("BA_faible_REPEREsortie_NH4").Cells(1, pi_FiliereConsideree)
    mn_ConcentrationSortieNO3_mgL_force = Feuil13.Range("BA_faible_REPEREsortie_NO3").Cells(1, pi_FiliereConsideree)
    mn_ConsommationMethanolPur_kgj_force = Feuil13.Range("BA_faible_REPEREmethanol").Cells(1, pi_FiliereConsideree)
    If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_PRECIPITATION Or mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE_ET_PRECIPITATION Then
        mn_ConcentrationSortiePt_mgL_force = Feuil13.Range("BA_faible_REPEREsortie_Pt").Cells(1, pi_FiliereConsideree)
        mn_ConsommationFeCl3Pur_kgj_force = Feuil13.Range("BA_faible_REPEREFeCl3").Cells(1, pi_FiliereConsideree)
    End If
    mn_HauteurBassinAeration_m_force = Feuil13.Range("BA_faible_REPEREhauteur_bassin").Cells(1, pi_FiliereConsideree)
    mn_BesoinsOxygene_kgO2j_force = Feuil13.Range("BA_faible_REPEREO2_besoin").Cells(1, pi_FiliereConsideree)
    mn_FacteurAlpha_force = Feuil13.Range("BA_faible_REPEREO2_facteur_alpha").Cells(1, pi_FiliereConsideree)
    If mi_ChoixTypeAeration = mi_AERATEUR_FINES_BULLES Or mi_ChoixTypeAeration = mi_AERATEUR_MOYENNES_BULLES Then
        mn_RendementTransfertOxygeneMembraneEauClaire_1m_force = Feuil13.Range("BA_faible_REPEREO2_rdt_transfert").Cells(1, pi_FiliereConsideree)
        mn_DebitAir_Nm3j_force = Feuil13.Range("BA_faible_REPEREair_Q_Nm3j").Cells(1, pi_FiliereConsideree)
        mn_DernierNettoyageOuRemplacamentMembranes_force = Feuil13.Range("BA_faible_REPEREdiffuseur_encrassement").Cells(1, pi_FiliereConsideree)
        mn_PressionAirSurpresse_mCE_force = Feuil13.Range("BA_faible_REPEREair_P_refoulement").Cells(1, pi_FiliereConsideree)
        mn_ConsommationSpecifiqueSurpresseur_Wh_Nm3mCE_force = Feuil13.Range("BA_faible_REPEREsurpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
    Else
        mn_ApportSpecifiqueBrutEauClaire_kgO2_kWh_force = Feuil13.Range("BA_faible_REPEREASB_eau_claire").Cells(1, pi_FiliereConsideree)
    End If
    mn_TauxRecirculationBouesClarificateur_force = Feuil13.Range("BA_faible_REPERErecirculation_taux").Cells(1, pi_FiliereConsideree)
    mn_PompesRecirculationBouesPression_mCE_force = Feuil13.Range("BA_faible_REPERErecirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
    mn_PompesRecirculationBouesRendement_force = Feuil13.Range("BA_faible_REPERErecirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
    mn_BouesISV_mL_g_force = Feuil13.Range("BA_faible_REPEREmd_BouesISV_mL_g").Cells(1, pi_FiliereConsideree)
    mn_EauClarifieeMES_mgL_force = Feuil13.Range("BA_faible_REPEREmd_EauClarifieeMES_mgL").Cells(1, pi_FiliereConsideree)
    mn_HauteurClarificateur_m_force = Feuil13.Range("BA_faible_REPEREmd_HauteurClarificateur_m").Cells(1, pi_FiliereConsideree)
    mn_ClarificationVitesseMaximale_mh_force = Feuil13.Range("BA_faible_REPEREmd_ClarificationVitesseMaximale_mh").Cells(1, pi_FiliereConsideree)
    mn_SurfaceTotaleClarification_m2_force = Feuil13.Range("BA_faible_REPEREmd_SurfaceTotaleClarification_m2").Cells(1, pi_FiliereConsideree)
    mn_BouesExtraitesConcentrationMES_gL_force = Feuil13.Range("BA_faible_REPEREmd_BouesExtraitesConcentrationMES_gL").Cells(1, pi_FiliereConsideree)
    mn_BouesExtraitesFluxMES_kgj_force = Feuil13.Range("BA_faible_REPEREpd_BouesExtraitesFluxMES_kgj").Cells(1, pi_FiliereConsideree)
    mn_BouesExtraitesDebit_m3j_force = False
    mn_PompesBouesExcesPression_force = Feuil13.Range("BA_faible_REPEREmd_PompesBouesExcesPression").Cells(1, pi_FiliereConsideree)
    md_PompesBouesExcesRendement_force = Feuil13.Range("BA_faible_REPEREmd_PompesBouesExcesRendement").Cells(1, pi_FiliereConsideree)

    'lecture des valeurs forcées
    If mn_NominalAgeBoues_j_force = True Then
        md_NominalAgeBoues_j_f = Feuil13.Range("BA_faible_REPEREnominal_age_boue").Cells(1, pi_FiliereConsideree)
    End If
    If mn_NominalConcentrationBassinsMES_gL_force = True Then
        md_NominalConcentrationBassinsMES_gL_f = Feuil13.Range("BA_faible_REPEREmd_NominalConcentrationBassinsMES_gL").Cells(1, pi_FiliereConsideree)
    End If
    If mn_NominalRatioMVMES_force = True Then
        pd_NominalRatioMVMES_f = Feuil13.Range("BA_faible_REPEREnominal_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If mn_VolumeBassinsTotal_m3_force = True Then
        md_VolumeBassinsTotal_m3_f = Feuil13.Range("BA_faible_REPEREvolume_bassins").Cells(1, pi_FiliereConsideree)
    End If
    If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE Or mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE_ET_PRECIPITATION Then
        If mn_VolumeBassinAnaerobie_m3_force = True Then
            md_VolumeBassinAnaerobie_m3_f = Feuil13.Range("BA_faible_REPEREvolume_anaerobie").Cells(1, pi_FiliereConsideree)
        End If
        If mn_PuissanceAgitationAnaerobie_Wm3_force = True Then
            md_PuissanceAgitationAnaerobie_Wm3_f = Feuil13.Range("BA_faible_REPEREratio_elec_anaerobie").Cells(1, pi_FiliereConsideree)
        End If
        If mn_RendementDephosphatationBiologique_force = True Then
            md_RendementDephosphatationBiologique_f = Feuil13.Range("BA_faible_REPERErendement_bioP").Cells(1, pi_FiliereConsideree)
        End If
    End If
    If mi_ChoixConfigurationBassins = mi_BASSINS_SEPARES Then
        If mn_VolumeBassinAnoxie_m3_force = True Then
            md_VolumeBassinAnoxie_m3_f = Feuil13.Range("BA_faible_REPEREvolume_anoxie").Cells(1, pi_FiliereConsideree)
        End If
        If mn_PuissanceAgitationAnoxie_Wm3_force = True Then
            md_PuissanceAgitationAnoxie_Wm3_f = Feuil13.Range("BA_faible_REPEREratio_elec_anoxie").Cells(1, pi_FiliereConsideree)
        End If
        If mn_VolumeBassinAerobie_m3_force = True Then
            md_VolumeBassinAerobie_m3_f = Feuil13.Range("BA_faible_REPEREratio_elec_aerobie").Cells(1, pi_FiliereConsideree)
        End If
        If mn_PuissanceAgitationAerobie_Wm3_force = True Then
            md_PuissanceAgitationAerobie_Wm3_f = Feuil13.Range("BA_faible_REPEREratio_elec_aerobie").Cells(1, pi_FiliereConsideree)
        End If
        If mn_TauxRecirculationMLSS_force = True Then
            md_TauxRecirculationMLSS_f = Feuil13.Range("BA_faible_REPERErecirculation_MLSS_taux").Cells(1, pi_FiliereConsideree)
        End If
        If mn_PompesRecirculationMLSSPression_mCE_force = True Then
            md_PompesRecirculationMLSSPression_mCE_f = Feuil13.Range("BA_faible_REPERErecirculation_MLSS_P_refoulement").Cells(1, pi_FiliereConsideree)
        End If
        If mn_PompesRecirculationMLSSRendement_force = True Then
            md_PompesRecirculationMLSSRendement_f = Feuil13.Range("BA_faible_REPERErecirculation_MLSS_pompe_rdt").Cells(1, pi_FiliereConsideree)
        End If
    Else
        If mn_VolumeChenal_m3_force = True Then
            md_VolumeChenal_m3_f = Feuil13.Range("BA_faible_REPEREvolume_chenal").Cells(1, pi_FiliereConsideree)
        End If
        If mn_PuissanceAgitationChenal_Wm3_force = True Then
            md_PuissanceAgitationChenal_Wm3_f = Feuil13.Range("BA_faible_REPEREratio_elec_chenal").Cells(1, pi_FiliereConsideree)
        End If
    End If
    If mn_ReelAgeBoues_j_force = True Then
        md_ReelAgeBoues_j_f = Feuil13.Range("BA_faible_REPEREreel_age_boue").Cells(1, pi_FiliereConsideree)
    End If
    If mn_ReelRatioMVMES_force = True Then
        md_ReelRatioMVMES_f = Feuil13.Range("BA_faible_REPEREreel_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If mn_ReelConcentrationBassinsMES_gL_force = True Then
        md_ReelConcentrationBassinsMES_gL_f = Feuil13.Range("BA_faible_REPEREreel_MES_concentration_bassin").Cells(1, pi_FiliereConsideree)
    End If
    If mn_ConcentrationO2dissous_mgL_force = True Then
        md_ConcentrationO2dissous_mgL_f = Feuil13.Range("BA_faible_REPEREO2_dissous").Cells(1, pi_FiliereConsideree)
    End If
    If mn_ConcentrationSortieDBO_mgL_force = True Then
        md_ConcentrationSortieDBO_mgL_f = Feuil13.Range("BA_faible_REPEREsortie_DBO").Cells(1, pi_FiliereConsideree)
    End If
    If mn_ConcentrationSortieNH4_mgL_force = True Then
        md_ConcentrationSortieNH4_mgL_f = Feuil13.Range("BA_faible_REPEREsortie_NH4").Cells(1, pi_FiliereConsideree)
    End If
    If mn_ConcentrationSortieNO3_mgL_force = True Then
        md_ConcentrationSortieNO3_mgL_f = Feuil13.Range("BA_faible_REPEREsortie_NO3").Cells(1, pi_FiliereConsideree)
    End If
    If mn_ConsommationMethanolPur_kgj_force = True Then
        md_ConsommationMethanolPur_kgj_f = Feuil13.Range("BA_faible_REPEREmethanol").Cells(1, pi_FiliereConsideree)
    End If
    If mn_ConcentrationSortiePt_mgL_force = True Then
        md_ConcentrationSortiePt_mgL_f = Feuil13.Range("BA_faible_REPEREsortie_Pt").Cells(1, pi_FiliereConsideree)
    End If
    If mn_ConsommationFeCl3Pur_kgj_force = True Then
        md_ConsommationFeCl3Pur_kgj_f = Feuil13.Range("BA_faible_REPEREFeCl3").Cells(1, pi_FiliereConsideree)
    End If
    If mn_HauteurBassinAeration_m_force = True Then
        md_HauteurBassinAeration_m_f = Feuil13.Range("BA_faible_REPEREhauteur_bassin").Cells(1, pi_FiliereConsideree)
    End If
    If mn_BesoinsOxygene_kgO2j_force = True Then
        md_BesoinsOxygene_kgO2j_f = Feuil13.Range("BA_faible_REPEREO2_besoin").Cells(1, pi_FiliereConsideree)
    End If
    If mn_FacteurAlpha_force = True Then
        md_FacteurAlpha_f = Feuil13.Range("BA_faible_REPEREO2_facteur_alpha").Cells(1, pi_FiliereConsideree)
    End If
    If mi_ChoixTypeAeration = mi_AERATEUR_FINES_BULLES Or mi_ChoixTypeAeration = mi_AERATEUR_MOYENNES_BULLES Then
        If mn_RendementTransfertOxygeneMembraneEauClaire_1m_force = True Then
            md_RendementTransfertOxygeneMembraneEauClaire_1m_f = Feuil13.Range("BA_faible_REPEREO2_rdt_transfert").Cells(1, pi_FiliereConsideree)
        End If
        If mn_DebitAir_Nm3j_force = True Then
            md_DebitAir_Nm3j_f = Feuil13.Range("BA_faible_REPEREair_Q_Nm3j").Cells(1, pi_FiliereConsideree)
        End If
        If mn_DernierNettoyageOuRemplacamentMembranes_force = True Then
            md_DernierNettoyageOuRemplacamentMembranes_f = Feuil13.Range("BA_faible_REPEREdiffuseur_encrassement").Cells(1, pi_FiliereConsideree)
        End If
        If mn_PressionAirSurpresse_mCE_force = True Then
            md_PressionAirSurpresse_mCE_f = Feuil13.Range("BA_faible_REPEREair_P_refoulement").Cells(1, pi_FiliereConsideree)
        End If
        If mn_ConsommationSpecifiqueSurpresseur_Wh_Nm3mCE_force = True Then
            md_ConsommationSpecifiqueSurpresseur_Wh_Nm3mCE_f = Feuil13.Range("BA_faible_REPEREsurpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
        End If
    Else
        If mn_ApportSpecifiqueBrutEauClaire_kgO2_kWh_force = True Then
            md_ApportSpecifiqueBrutEauClaire_kgO2_kWh_f = Feuil13.Range("BA_faible_REPEREASB_eau_claire").Cells(1, pi_FiliereConsideree)
        End If
    End If
    If mn_TauxRecirculationBouesClarificateur_force = True Then
        md_TauxRecirculationBouesClarificateur_f = Feuil13.Range("BA_faible_REPERErecirculation_taux").Cells(1, pi_FiliereConsideree)
    End If
    If mn_PompesRecirculationBouesPression_mCE_force = True Then
        md_PompesRecirculationBouesPression_mCE_f = Feuil13.Range("BA_faible_REPERErecirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If mn_PompesRecirculationBouesRendement_force = True Then
        md_PompesRecirculationBouesRendement_f = Feuil13.Range("BA_faible_REPERErecirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    If mn_BouesISV_mL_g_force = True Then
        md_BouesISV_mL_g_f = Feuil13.Range("BA_faible_REPEREmd_BouesISV_mL_g").Cells(1, pi_FiliereConsideree)
    End If
    If mn_EauClarifieeMES_mgL_force = True Then
        md_EauClarifieeMES_mgL_f = Feuil13.Range("BA_faible_REPEREmd_EauClarifieeMES_mgL").Cells(1, pi_FiliereConsideree)
    End If
    If mn_HauteurClarificateur_m_force = True Then
        md_HauteurClarificateur_m_f = Feuil13.Range("BA_faible_REPEREmd_HauteurClarificateur_m").Cells(1, pi_FiliereConsideree)
    End If
    If mn_ClarificationVitesseMaximale_mh_force = True Then
        md_ClarificationVitesseMaximale_mh_f = Feuil13.Range("BA_faible_REPEREmd_ClarificationVitesseMaximale_mh").Cells(1, pi_FiliereConsideree)
    End If
    If mn_SurfaceTotaleClarification_m2_force = True Then
        md_SurfaceTotaleClarification_m2_f = Feuil13.Range("BA_faible_REPEREmd_SurfaceTotaleClarification_m2").Cells(1, pi_FiliereConsideree)
    End If
    If mn_BouesExtraitesConcentrationMES_gL_force = True Then
        md_BouesExtraitesConcentrationMES_gL_f = Feuil13.Range("BA_faible_REPEREmd_BouesExtraitesConcentrationMES_gL").Cells(1, pi_FiliereConsideree)
    End If
    If mn_BouesExtraitesFluxMES_kgj_force = True Then
        pd_BouesExtraitesFluxMES_kgj_f = Feuil13.Range("BA_faible_REPEREpd_BouesExtraitesFluxMES_kgj").Cells(1, pi_FiliereConsideree)
    End If
    If mn_BouesExtraitesDebit_m3j_force = True Then
        pd_BouesExtraitesFluxMES_kgj_f = Feuil13.Range("BA_faible_REPEREpd_BouesExtraitesDebit_m3j").Cells(1, pi_FiliereConsideree)
    End If
    If mn_PompesBouesExcesPression_force = True Then
        md_PompesBouesExcesPression_f = Feuil13.Range("BA_faible_REPEREmd_PompesBouesExcesPression").Cells(1, pi_FiliereConsideree)
    End If
    If md_PompesBouesExcesRendement_force = True Then
        md_PompesBouesExcesRendement_f = Feuil13.Range("BA_faible_REPEREmd_PompesBouesExcesRendement").Cells(1, pi_FiliereConsideree)
    End If
    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "13.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub hypotheses()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'Voici l'ensemble des hypothèses sur le process

    Dim I As Integer
    Dim j As Integer

    Dim G_ratio_augmentation_P_bio As Double


    'Age de boues conseillés
    md_AgeBouesCarboneSeul12°C_j = 4
    md_AgeBouesStabilisationAerobie12°C_j = 25

    'Assimilation de l'azote
    md_AssimilationN_kgN_kgDBOeliminee = Feuil48.Range("hyp_II_assimilation_N").Cells(1, 1)     '0.05

    'Température de référence pour la cinétique biologique (°C)
    md_TemperatureReferenceAgeBouesEquivalent_°C = Feuil48.Range("hyp_II_BA_bio_equiv_Tref").Cells(1, 1)    ' 12

    'Age de boues à md_TemperatureReferenceAgeBouesEquivalent_°C (j)
    If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE Or mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE_ET_PRECIPITATION Then
        G_ratio_augmentation_P_bio = 0.08
    Else
        G_ratio_augmentation_P_bio = 0
    End If
    G_reference = 14 * (1 + G_ratio_augmentation_P_bio)

    'ratio MV/MES référence
    If choix_primaire = True Then
        MV_MES_reference = 0.77
    Else
        MV_MES_reference = 0.67
    End If

    'Facteur correctif température pour cinétique
    md_CoefficientCalculAgeBouesEquivalent = Feuil48.Range("hyp_II_BA_bio_equiv_correctionT").Cells(1, 1)    '1.072

    'coefficient permettant le calcul du volume anaérobie (temps de séjour)
    md_TempsPassageDebitEntreeAnaerobie_j = 2 / pd_NOMBRE_HEURE_PAR_JOUR

    'ratio DCO/N limite pour calcul de zone anoxie
    ratio_DCO_NK_10 = 10

    'résiduel de Pt au dessous duquel on ne pourra pas descendre (un peu aléatoire)
    Pt_residuel = 0.1   'mg/L

    'coefficients relation charge massique (kgDBO/kgMVS/j)  1/Cm=a1*G_eq+a0
    a1_Cm = Feuil48.Range("hyp_II_BA_charge_massique_a1").Cells(1, 1)    '0.4959
    a0_Cm = Feuil48.Range("hyp_II_BA_charge_massique_a0").Cells(1, 1)    '1.3908

    'coefficients relation rendement DBO fonction de Cm_equivalente rdt_DBO=a2*Cm^2+a1*Cm+a0
    If choix_primaire = True Then
        For I = 1 To 4
            md_CoefficientPolynomeCalculRendementDBO(I - 1) = Feuil48.Range("hyp_II_BA_rdtDBO_coef_ED").Cells(I, 1)
        Next I
    Else
        For I = 1 To 4
            md_CoefficientPolynomeCalculRendementDBO(I - 1) = Feuil48.Range("hyp_II_BA_rdtDBO_coef_EB").Cells(I, 1)
        Next I
    End If

    'P consommé dans synthèse
    ratio_P_synthese = Feuil48.Range("hyp_II_assimilation_P").Cells(1, 1)  ' 0.01

    'coefficient K
    a_Pbio_coef_K = 0.0276
    b_Pbio_coef_K = 0.079

    'Kp_bio_1 (correction volume cuve anaérobie
    a_Kp_bio_1 = 0.7
    b_Kp_bio_1 = 0.3
    lim_Kp_bio_1 = 1.15
    If choix_primaire = True Then
        ratio_Psol_P = 0.85
    Else
        ratio_Psol_P = 0.7
    End If
    rdt_Psol_max = 0.75

    'precipitation de P
    ratio_FeCl3_Pt = 5.242    '  162.5 / 31
    a_dosis_MMS = 3.447    ' (55.85 + 3 * (16 + 1)) / 31
    b_dosis_MMS = 4.866     ' (55.85 + 31 + 4 * 16) / 31

    'calcul du méthanol
    ratio_DCO_sol_EB = 0.33
    If choix_decanteur_reactif = True Then
        rdt_DCO_sol_primaire = 0.08
    Else
        rdt_DCO_sol_primaire = 0
    End If
    
    'Paramètres pour la consommation externe de carbone
    md_ConsommationDCOexterneDenitrification_kgDCO_kgNDenitrifie = 5
    md_DCOdansMethanol_kgDCO_kg = 1.5
    md_RatioDBOsurDCOcarboneExterne_kgDBO_kgDCO = 0.5
    
    
    ratio_boues_methanol = 0.25

    'hypothèse de rendement sur la DCO
    rdt_DCO = 0.9

    'Charge radier maximale kg/(m2.h)
    For I = mi_RACLEUR_RACLE To mi_RACLEUR_KRUGER
        charge_radier_max(I) = Feuil48.Range("hyp_II_BA_charge_radier_max").Cells(I, 1)
    Next I

    'facteur pour calcul de concentration limite des boues extraites et recirculées pour bonne décantation
    facteur_boues_C_lim = 1200

    'hypothèses pour le calcul des besoins en O2
    md_CoefficientAConsommationSpecifiqueO2RespirationCarbone_kgO2_kgDBO = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_a0").Cells(1, 1)   ' 0.56
    md_CoefficientBConsommationSpecifiqueO2RespirationCarbone_kgO2_kgDBOj = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_a1").Cells(1, 1)  ' 0.15
    md_CoefficientCConsommationSpecifiqueO2RespirationCarbone = 1
    md_CoefficientDConsommationSpecifiqueO2RespirationCarbone = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_a2").Cells(1, 1)   '0.17
    md_CoefficientAFacteurImpactTemperature = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_correctionT").Cells(1, 1)   '1.072
    md_CoefficientBFacteurImpactTemperature_°C = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_Tref").Cells(1, 1)   '15

    'coefficient production de boues
     md_CoefficientAProductionSpecifiqueBouesCarbone_kgMES_kgDBO = 0.75
     md_CoefficientBProductionSpecifiqueBouesCarbone = 0.6
     md_CoefficientCProductionSpecifiqueBouesCarbone = (1 - 0.2)
     md_CoefficientDProductionSpecifiqueBouesCarbone = 0.17
     md_CoefficientEProductionSpecifiqueBouesCarbone_kgMES_kgDBO = 0.75
     md_CoefficientFProductionSpecifiqueBouesCarbone = 1
     md_CoefficientGProductionSpecifiqueBouesCarbone = 0.17

    'Coefficient calcul de Vd/Vbb (part de volume pour la dénitrification)
    md_CoefficientARatioVolumeAnoxiqueSurVolumeAnoxiqueEtAerobie = 0.75
    

    'spécificité SHUNT
    If choix_fct_SHUNT = True Then
        md_ConsommationSpecifiqueO2Nitrification_kgO2_kgNNitrifie = 4.57 * 0.75
        md_ConsommationSpecifiqueDCODenitrification_kgDCO_kgN = 2.86 * 0.6
        ratio_C_N_limite_nominal = ratio_C_N_limite_nominal * 0.6
        md_ConsommationDCOexterneDenitrification_kgDCO_kgNDenitrifie = md_ConsommationDCOexterneDenitrification_kgDCO_kgNDenitrifie * 0.6
    Else
        md_ConsommationSpecifiqueO2Nitrification_kgO2_kgNNitrifie = 4.57
        md_ConsommationSpecifiqueDCODenitrification_kgDCO_kgN = 2.86
    End If

    'facteurs alpha0 (celui de fines bulles est calculé)
    O2_alpha0(mi_AERATEUR_MOYENNES_BULLES) = 0.8
    O2_alpha0(mi_AERATEUR_BROSSES) = 0.8
    O2_alpha0(mi_AERATEUR_TURBINES_LENTES) = 0.8
    O2_alpha0(mi_AERATEUR_TURBINES_RAPIDES) = 0.8

    'correction alpha en fonction des MES  alfa=alfa*(a0+a1*MES_reel)/(a0+a1*MES_lim)
    a0_correctif_alfa_MES = Feuil48.Range("hyp_II_BA_alpha_correction_MES_a0").Cells(1, 1)    ' 1.0007
    a1_correctif_alfa_MES = Feuil48.Range("hyp_II_BA_alpha_correction_MES_a1").Cells(1, 1)    ' -0.035
    MES_reference_alfa = Feuil48.Range("hyp_II_BA_alpha_correction_MES_Cref").Cells(1, 1)    ' 5

    'relation calcul alpha fines bulles    ao+a_MVS*MVS+a_G*G_eq
    a0_alpha = Feuil48.Range("hyp_II_BA_alpha_finebulle_a0").Cells(1, 1)    ' 0.51
    a_alpha_MVS = Feuil48.Range("hyp_II_BA_alpha_finebulle_aMVS").Cells(1, 1)    '-0.062
    a_alpha_G = Feuil48.Range("hyp_II_BA_alpha_finebulle_aGeq").Cells(1, 1)    ' 0.019

    'facteur beta
    O2_facteur_beta = Feuil48.Range("hyp_calcul_facteurK_facteur_Beta").Cells(1, 1)    ' 0.95

    'calcul de la pression atmosphérique  Patm=P0*((a0-a1*altitude)^a2)/100
    P0 = Feuil48.Range("hyp_calcul_Patm_coef").Cells(1, 1)   ' 101325
    a0_Patm = Feuil48.Range("hyp_calcul_Patm_coef").Cells(2, 1)    '1
    a1_Patm = Feuil48.Range("hyp_calcul_Patm_coef").Cells(3, 1)    '-0.0000225577
    a2_Patm = Feuil48.Range("hyp_calcul_Patm_coef").Cells(4, 1)    '5.25588

    'calcul de la surpression en diffusion de bulles  delta_P=a0/a1*P0/100*(md_HauteurBassinAeration_m-hauteur_diffuseur)
    a0_delta_P = Feuil48.Range("hyp_insufflation_deltaP_a0").Cells(1, 1)  '0.35
    a1_delta_P = Feuil48.Range("hyp_insufflation_deltaP_a1").Cells(1, 1)   '10.33
    hauteur_diffuseur = Feuil48.Range("hyp_insufflation_bulles_hauteur_diffuseur").Cells(1, 1)   '0.25 'm

    'calcul O2 saturation  O2_saturation(T,P0)= somme(ai*Teau^i)
    For I = 1 To 5
        a_O2_saturation(I - 1) = Feuil48.Range("hyp_calcul_O2sat_T_P0_coef_a").Cells(I, 1)
    Next I

    'correction température pour facteur K
    md_CoefficientCalculAgeBouesEquivalent_K = Feuil48.Range("hyp_transfertO2_facteurK_correction_T").Cells(1, 1)    ' 1.024
    T_ref_K = Feuil48.Range("hyp_transfertO2_facteurK_Tref").Cells(1, 1)  '20

    'O2 présent dans l'air (kgO2/Nm3air)
    ratio_kgO2_Nm3air = Feuil48.Range("hyp_ratio_kgO2_Nm3air").Cells(1, 1)  ' 0.3

    'rendement de dissolution en eau claire (%/m)
    If mn_HauteurBassinAeration_m_force = True And md_HauteurBassinAeration_m_f < Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(4, 1) Then
        O2_rdt_dissolution(mi_AERATEUR_FINES_BULLES) = Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(5, 1) * md_HauteurBassinAeration_m_f + Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(6, 1)
    Else
        O2_rdt_dissolution(mi_AERATEUR_FINES_BULLES) = Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(1, 1)
    End If
    O2_rdt_dissolution(mi_AERATEUR_MOYENNES_BULLES) = Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(2, 1)

    'Apport spécifique brut (eau claire)  (kWh/kO2_transféré)   source:fndae 31
    ASB(mi_AERATEUR_BROSSES) = Feuil48.Range("hyp_II_BA_aerateur_surface_ASB").Cells(1, 1)   '1.7
    ASB(mi_AERATEUR_TURBINES_LENTES) = Feuil48.Range("hyp_II_BA_aerateur_surface_ASB").Cells(2, 1)   '1.65
    ASB(mi_AERATEUR_TURBINES_RAPIDES) = Feuil48.Range("hyp_II_BA_aerateur_surface_ASB").Cells(3, 1)   '1.2

    'consommations spécifiques par type de surpresseurs      en Wh/(Nm3.mCE)
    For I = 1 To nb_surpresseur_gamme_Q
        For j = 1 To nb_surpresseur_gamme_P
            surpresseur_conso_Wh_Nm3mCE(mi_SURPRESSEUR_LOBES, I, j) = 5
            surpresseur_conso_Wh_Nm3mCE(mi_SURPRESSEUR_VIS, I, j) = 3.8
            surpresseur_conso_Wh_Nm3mCE(mi_SURPRESSEUR_CENTRIFUGE, I, j) = 3.5
        Next j
    Next I

    'impact du type de régulation
    impact_regulation(mi_AERATION_REGULATION_HORLOGE) = 0.4
    impact_regulation(mi_AERATION_REGULATION_CLASSIQUE_SANS_VARIATION_FREQUENCE) = 0.25
    impact_regulation(mi_AERATION_REGULATION_CLASSIQUE_AVEC_VARIATION_FREQUENCE) = 0.1
    impact_regulation(mi_AERATION_REGULATION_AVANCEE) = 0

    'clarificateur
    rate_puissance_fixe = 0.1    'Wh/m3 par rapport au nominal
    rate_puissance_debit = 1    'Wh/m3 par rapport au débit réel

    agitation_fct = 24

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "13.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Private Sub attribution_valeur_par_defaut()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On retrouve ici l'ensemble des valeurs pour lesquelles on attribue une valeur guide par défaut
    Dim volume_corrige As Double
    Dim G_equivalent_md_NominalAgeBoues_j As Double
    Dim rdt_pompe As Double
    Dim rdt_moteur_pompe As Double


    'Age de boues design (j)
    If mn_NominalAgeBoues_j_force = True Then
        md_NominalAgeBoues_j = md_NominalAgeBoues_j_f
    Else
        'formule standard G (j) = 14 * 1,07 * (12 - T ) , si Pbio = 14,8 * 1,07 * (12 - T )
        md_NominalAgeBoues_j = G_reference * md_CoefficientCalculAgeBouesEquivalent ^ (md_TemperatureReferenceAgeBouesEquivalent_°C - pd_TemperatureEauDimensionnement_°C)
    End If

    'Concentration MES design (dans les bassins) g/L
    If mn_NominalConcentrationBassinsMES_gL_force = True Then
        md_NominalConcentrationBassinsMES_gL = md_NominalConcentrationBassinsMES_gL_f
    Else
        If choix_primaire = True Then
            md_NominalConcentrationBassinsMES_gL = 3.5
        Else
            md_NominalConcentrationBassinsMES_gL = 5
        End If
    End If

    'Age de boues réel (j)
    If mn_ReelAgeBoues_j_force = True Then
        md_ReelAgeBoues_j = md_ReelAgeBoues_j_f
    Else
        G_equivalent_md_NominalAgeBoues_j = md_NominalAgeBoues_j / md_CoefficientCalculAgeBouesEquivalent ^ (md_TemperatureReferenceAgeBouesEquivalent_°C - pd_TemperatureEauDimensionnement_°C)
        md_ReelAgeBoues_j = G_equivalent_md_NominalAgeBoues_j * md_CoefficientCalculAgeBouesEquivalent ^ (md_TemperatureReferenceAgeBouesEquivalent_°C - pd_TemperatureEauConditionsReelles_°C)
    End If

    'agitation des cuves (W/m3)
    If mn_PuissanceAgitationAnoxie_Wm3_force = True Then
        md_PuissanceAgitationAnoxie_Wm3 = md_PuissanceAgitationAnoxie_Wm3_f
    Else
        md_PuissanceAgitationAnoxie_Wm3 = 3
    End If
    If mn_PuissanceAgitationAerobie_Wm3_force = True Then
        md_PuissanceAgitationAerobie_Wm3 = md_PuissanceAgitationAerobie_Wm3_f
    Else
        md_PuissanceAgitationAerobie_Wm3 = 0
    End If
    If mn_PuissanceAgitationChenal_Wm3_force = True Then
        md_PuissanceAgitationChenal_Wm3 = md_PuissanceAgitationChenal_Wm3_f
    Else
        md_PuissanceAgitationChenal_Wm3 = 3
    End If
    If mn_PuissanceAgitationAnaerobie_Wm3_force = True Then
        md_PuissanceAgitationAnaerobie_Wm3 = md_PuissanceAgitationAnaerobie_Wm3_f
    Else
        md_PuissanceAgitationAnaerobie_Wm3 = 10
    End If

    'Concentration moyenne en O2 dissous (mg/L)
    If mn_ConcentrationO2dissous_mgL_force = True Then
        md_ConcentrationO2dissous_mgL = md_ConcentrationO2dissous_mgL_f
    Else
        If choix_fct_SHUNT = True Then
            md_ConcentrationO2dissous_mgL = 0.5   'spécificité SHUNT / O2 dissous très faible nécessaire
        Else
            If mi_ChoixRegulationAeration = mi_AERATION_REGULATION_AVANCEE Then
                If pd_TemperatureEauConditionsReelles_°C < 14 Then
                    md_ConcentrationO2dissous_mgL = 1.5
                Else
                    md_ConcentrationO2dissous_mgL = 0.8
                End If
            Else
                If pd_TemperatureEauConditionsReelles_°C < 14 Then
                    md_ConcentrationO2dissous_mgL = 2
                Else
                    md_ConcentrationO2dissous_mgL = 1
                End If
            End If
        End If
    End If

    'Concentration de sortie en NH4  (mg/L)
    If mn_ConcentrationSortieNH4_mgL_force = True Then
        md_ConcentrationSortieNH4_mgL = md_ConcentrationSortieNH4_mgL_f
    Else
        md_ConcentrationSortieNH4_mgL = 3
    End If

    'Concentration de sortie en NO3  (mg/L)
    If mn_ConcentrationSortieNO3_mgL_force = True Then
        md_ConcentrationSortieNO3_mgL = md_ConcentrationSortieNO3_mgL_f
    Else
        md_ConcentrationSortieNO3_mgL = 5
    End If

    'Concentration de sortie en Pt (mg/L)
    If mn_ConcentrationSortiePt_mgL_force = True Then
        md_ConcentrationSortiePt_mgL = md_ConcentrationSortiePt_mgL_f
    Else
        md_ConcentrationSortiePt_mgL = Pt_garantie
    End If

    'Hauteur moyenne des bassins  (m)
    If mn_HauteurBassinAeration_m_force = True Then
        md_HauteurBassinAeration_m = md_HauteurBassinAeration_m_f
    Else
        If mi_ChoixTypeAeration = mi_AERATEUR_FINES_BULLES Or mi_ChoixTypeAeration = mi_AERATEUR_MOYENNES_BULLES Then
            md_HauteurBassinAeration_m = 6
        Else
            md_HauteurBassinAeration_m = 3
        End If
    End If

    'Durée depuis le dernier nettoyage des diffuseurs  (ans)
    If mn_DernierNettoyageOuRemplacamentMembranes_force = True Then
        md_DernierNettoyageOuRemplacamentMembranes = md_DernierNettoyageOuRemplacamentMembranes_f
    Else
        md_DernierNettoyageOuRemplacamentMembranes = 0
    End If

    'Pression de refoulement des surpresseurs (mCE)
    If mn_PressionAirSurpresse_mCE_force = True Then
        md_PressionAirSurpresse_mCE = md_PressionAirSurpresse_mCE_f
    Else
        md_PressionAirSurpresse_mCE = md_HauteurBassinAeration_m + 2 + 0.25 * md_DernierNettoyageOuRemplacamentMembranes
    End If

    'ASB
    If mn_ApportSpecifiqueBrutEauClaire_kgO2_kWh_force = True Then
        md_ApportSpecifiqueBrutEauClaire_kgO2_kWh = md_ApportSpecifiqueBrutEauClaire_kgO2_kWh_f
    Else
        md_ApportSpecifiqueBrutEauClaire_kgO2_kWh = ASB(mi_ChoixTypeAeration)
    End If

    'Taux de recirculation de liqueur mixte
    'If mn_TauxRecirculationMLSS_force = True Then
    'md_TauxRecirculationMLSS = md_TauxRecirculationMLSS_f
    'Else
    'md_TauxRecirculationMLSS = 1.5
    'End If

    'Pression des pompes de recirculation de liqueur mixte (mCE)
    If mn_PompesRecirculationMLSSPression_mCE_force = True Then
        md_PompesRecirculationMLSSPression_mCE = md_PompesRecirculationMLSSPression_mCE_f
    Else
        md_PompesRecirculationMLSSPression_mCE = 2
    End If

    'recirculation de liqueur mixte
    If mn_PompesRecirculationMLSSRendement_force = True Then
        md_PompesRecirculationMLSSRendement = md_PompesRecirculationMLSSRendement_f
    Else
        rdt_pompe = 0.7
        rdt_moteur_pompe = 0.88
        md_PompesRecirculationMLSSRendement = rdt_pompe * rdt_moteur_pompe
    End If
    ratio_elec_recirculation_MLSS = pd_ACCELERATION_PESANTEUR_m_s2 / (pd_NOMBRE_SECONDE_PAR_HEURE * md_PompesRecirculationMLSSRendement)    ' kWh/(m3*mCE)


    'Taux de recirculation des boues du clarif
    If mn_TauxRecirculationBouesClarificateur_force = True Then
        md_TauxRecirculationBouesClarificateur = md_TauxRecirculationBouesClarificateur_f
    Else
        md_TauxRecirculationBouesClarificateur = 1
    End If

    'Pression des pompes de recirculation des boues du clarif (mCE)
    If mn_PompesRecirculationBouesPression_mCE_force = True Then
        md_PompesRecirculationBouesPression_mCE = md_PompesRecirculationBouesPression_mCE_f
    Else
        md_PompesRecirculationBouesPression_mCE = 5
    End If

    'recirculation des boues du clarif
    If mn_PompesRecirculationBouesRendement_force = True Then
        md_PompesRecirculationBouesRendement = md_PompesRecirculationBouesRendement_f
    Else
        rdt_pompe = 0.7
        rdt_moteur_pompe = 0.88
        md_PompesRecirculationBouesRendement = rdt_pompe * rdt_moteur_pompe
    End If
    ratio_elec_recirculation = pd_ACCELERATION_PESANTEUR_m_s2 / (pd_NOMBRE_SECONDE_PAR_HEURE * md_PompesRecirculationBouesRendement)    ' kWh/(m3*mCE)

    'Indice de Mohlman (mL/g)
    If mn_BouesISV_mL_g_force = True Then
        md_BouesISV_mL_g = md_BouesISV_mL_g_f
    Else
        If choix_primaire = True Then
            If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_NON Then
                md_BouesISV_mL_g = 165
            Else
                md_BouesISV_mL_g = 185
            End If
        Else
            If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_NON Then
                md_BouesISV_mL_g = 125
            Else
                md_BouesISV_mL_g = 135
            End If
        End If
    End If

    'Sortie MES eau traitée (mg/L)
    If mn_EauClarifieeMES_mgL_force = True Then
        md_EauClarifieeMES_mgL = md_EauClarifieeMES_mgL_f
    Else
        md_EauClarifieeMES_mgL = 20
    End If

    'Hauteur du clarificateur (m)
    If mn_HauteurClarificateur_m_force = True Then
        md_HauteurClarificateur_m = md_HauteurClarificateur_m_f
    Else
        md_HauteurClarificateur_m = 4
    End If

    'Vitesse hydraulique maximum (m/h)
    If mn_ClarificationVitesseMaximale_mh_force = True Then
        md_ClarificationVitesseMaximale_mh = md_ClarificationVitesseMaximale_mh_f
    Else
        volume_corrige = md_BouesISV_mL_g * md_NominalConcentrationBassinsMES_gL
        md_ClarificationVitesseMaximale_mh = (100 * md_HauteurClarificateur_m * (md_EauClarifieeMES_mgL / 3.15) ^ 0.5) / ((1 + md_TauxRecirculationBouesClarificateur) * volume_corrige)
    End If

    'Pression de refoulement de l'extraction (mCE)
    If mn_PompesBouesExcesPression_force = True Then
        md_PompesBouesExcesPression = md_PompesBouesExcesPression_f
    Else
        md_PompesBouesExcesPression = 5
    End If

    'extraction des boues du clarif
    If md_PompesBouesExcesRendement_force = True Then
        md_PompesBouesExcesRendement = md_PompesBouesExcesRendement_f
    Else
        rdt_pompe = 0.7
        rdt_moteur_pompe = 0.88
        md_PompesBouesExcesRendement = rdt_pompe * rdt_moteur_pompe
    End If
    ratio_elec_extraction = pd_ACCELERATION_PESANTEUR_m_s2 / (pd_NOMBRE_SECONDE_PAR_HEURE * md_PompesBouesExcesRendement)    ' kWh/(m3*mCE)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "13.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub ecriture_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil13.Range("BA_faible_REPEREnominal_age_boue").Cells(1, Colonne) = md_NominalAgeBoues_j
    Feuil13.Range("BA_faible_REPEREmd_NominalConcentrationBassinsMES_gL").Cells(1, Colonne) = md_NominalConcentrationBassinsMES_gL
    Feuil13.Range("BA_faible_REPEREnominal_MV_MES").Cells(1, Colonne) = pd_NominalRatioMVMES
    Feuil13.Range("BA_faible_REPEREvolume_bassins").Cells(1, Colonne) = md_VolumeBassinsTotal_m3
    If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE Or mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE_ET_PRECIPITATION Then
        Feuil13.Range("BA_faible_REPEREvolume_anaerobie").Cells(1, Colonne) = md_VolumeBassinAnaerobie_m3
        Feuil13.Range("BA_faible_REPEREratio_elec_anaerobie").Cells(1, Colonne) = md_PuissanceAgitationAnaerobie_Wm3
        Feuil13.Range("BA_faible_REPERErendement_bioP").Cells(1, Colonne) = md_RendementDephosphatationBiologique
    End If
    If mi_ChoixConfigurationBassins = mi_BASSINS_SEPARES Then
        Feuil13.Range("BA_faible_REPEREvolume_anoxie").Cells(1, Colonne) = md_VolumeBassinAnoxie_m3
        Feuil13.Range("BA_faible_REPEREratio_elec_anoxie").Cells(1, Colonne) = md_PuissanceAgitationAnoxie_Wm3
        Feuil13.Range("BA_faible_REPEREvolume_aerobie").Cells(1, Colonne) = md_VolumeBassinAerobie_m3
        Feuil13.Range("BA_faible_REPEREratio_elec_aerobie").Cells(1, Colonne) = md_PuissanceAgitationAerobie_Wm3
        Feuil13.Range("BA_faible_REPERErecirculation_MLSS_taux").Cells(1, Colonne) = md_TauxRecirculationMLSS
        Feuil13.Range("BA_faible_REPERErecirculation_MLSS_P_refoulement").Cells(1, Colonne) = md_PompesRecirculationMLSSPression_mCE
        Feuil13.Range("BA_faible_REPERErecirculation_MLSS_pompe_rdt").Cells(1, Colonne) = md_PompesRecirculationMLSSRendement
    Else
        Feuil13.Range("BA_faible_REPEREvolume_chenal").Cells(1, Colonne) = md_VolumeChenal_m3
        Feuil13.Range("BA_faible_REPEREratio_elec_chenal").Cells(1, Colonne) = md_PuissanceAgitationChenal_Wm3
    End If
    Feuil13.Range("BA_faible_REPEREreel_age_boue").Cells(1, Colonne) = md_ReelAgeBoues_j
    Feuil13.Range("BA_faible_REPEREreel_MV_MES").Cells(1, Colonne) = pd_ReelRatioMVMES
    Feuil13.Range("BA_faible_REPEREreel_MES_concentration_bassin").Cells(1, Colonne) = md_ReelConcentrationBassinsMES_gL
    Feuil13.Range("BA_faible_REPEREO2_dissous").Cells(1, Colonne) = md_ConcentrationO2dissous_mgL
    Feuil13.Range("BA_faible_REPEREsortie_DBO").Cells(1, Colonne) = md_ConcentrationSortieDBO_mgL
    Feuil13.Range("BA_faible_REPEREsortie_NH4").Cells(1, Colonne) = md_ConcentrationSortieNH4_mgL
    Feuil13.Range("BA_faible_REPEREsortie_NO3").Cells(1, Colonne) = md_ConcentrationSortieNO3_mgL
    Feuil13.Range("BA_faible_REPEREmethanol").Cells(1, Colonne) = pd_ConsommationMethanolPur_kgj
    If mi_ChoixDephosphatation <> mi_DEPHOSPHATATION_NON Then
        Feuil13.Range("BA_faible_REPEREsortie_Pt").Cells(1, Colonne) = md_ConcentrationSortiePt_mgL
    End If
    If mi_ChoixDephosphatation = mi_DEPHOSPHATATION_PRECIPITATION Or mi_ChoixDephosphatation = mi_DEPHOSPHATATION_BIOLOGIQUE_ET_PRECIPITATION Then
        Feuil13.Range("BA_faible_REPEREFeCl3").Cells(1, Colonne) = pd_ConsommationFeCl3Pur_kgj
    End If
    Feuil13.Range("BA_faible_REPEREhauteur_bassin").Cells(1, Colonne) = md_HauteurBassinAeration_m
    Feuil13.Range("BA_faible_REPEREO2_besoin").Cells(1, Colonne) = md_BesoinsOxygene_kgO2j
    Feuil13.Range("BA_faible_REPEREO2_facteur_alpha").Cells(1, Colonne) = md_FacteurAlpha
    If mi_ChoixTypeAeration = mi_AERATEUR_FINES_BULLES Or mi_ChoixTypeAeration = mi_AERATEUR_MOYENNES_BULLES Then
        Feuil13.Range("BA_faible_REPEREO2_rdt_transfert").Cells(1, Colonne) = md_RendementTransfertOxygeneMembraneEauClaire_1m
        Feuil13.Range("BA_faible_REPEREair_Q_Nm3j").Cells(1, Colonne) = md_DebitAir_Nm3j
        Feuil13.Range("BA_faible_REPEREdiffuseur_encrassement").Cells(1, Colonne) = md_DernierNettoyageOuRemplacamentMembranes
        Feuil13.Range("BA_faible_REPEREair_P_refoulement").Cells(1, Colonne) = md_PressionAirSurpresse_mCE
        Feuil13.Range("BA_faible_REPEREsurpresseur_conso_spec").Cells(1, Colonne) = md_ConsommationSpecifiqueSurpresseur_Wh_Nm3mCE
    Else
        Feuil13.Range("BA_faible_REPEREASB_eau_claire").Cells(1, Colonne) = md_ApportSpecifiqueBrutEauClaire_kgO2_kWh
    End If
    Feuil13.Range("BA_faible_REPERErecirculation_taux").Cells(1, Colonne) = md_TauxRecirculationBouesClarificateur
    Feuil13.Range("BA_faible_REPERErecirculation_P_refoulement").Cells(1, Colonne) = md_PompesRecirculationBouesPression_mCE
    Feuil13.Range("BA_faible_REPERErecirculation_pompe_rdt").Cells(1, Colonne) = md_PompesRecirculationBouesRendement
    Feuil13.Range("BA_faible_REPEREmd_BouesISV_mL_g").Cells(1, Colonne) = md_BouesISV_mL_g
    Feuil13.Range("BA_faible_REPEREmd_EauClarifieeMES_mgL").Cells(1, Colonne) = md_EauClarifieeMES_mgL
    Feuil13.Range("BA_faible_REPEREmd_HauteurClarificateur_m").Cells(1, Colonne) = md_HauteurClarificateur_m
    Feuil13.Range("BA_faible_REPEREmd_ClarificationVitesseMaximale_mh").Cells(1, Colonne) = md_ClarificationVitesseMaximale_mh
    Feuil13.Range("BA_faible_REPEREmd_SurfaceTotaleClarification_m2").Cells(1, Colonne) = md_SurfaceTotaleClarification_m2
    Feuil13.Range("BA_faible_REPEREmd_ClarificationVitesseMaximaleRecalculee_mhulee").Cells(1, Colonne) = md_ClarificationVitesseMaximaleRecalculee_mh
    Feuil13.Range("BA_faible_REPEREmd_ClarificateurChargeRadierNominal").Cells(1, Colonne) = md_ClarificateurChargeRadierNominal
    Feuil13.Range("BA_faible_REPEREmd_BouesExtraitesConcentrationMES_gL").Cells(1, Colonne) = md_BouesExtraitesConcentrationMES_gL
    Feuil13.Range("BA_faible_REPEREpd_BouesExtraitesFluxMES_kgj").Cells(1, Colonne) = pd_BouesExtraitesFluxMES_kgj
    Feuil13.Range("BA_faible_REPEREpd_BouesExtraitesDebit_m3j").Cells(1, Colonne) = pd_BouesExtraitesDebit_m3j
    Feuil13.Range("BA_faible_REPEREmd_PompesBouesExcesPression").Cells(1, Colonne) = md_PompesBouesExcesPression
    Feuil13.Range("BA_faible_REPEREmd_PompesBouesExcesRendement").Cells(1, Colonne) = md_PompesBouesExcesRendement
    Feuil13.Range("BA_faible_REPEREpd_ConsommationElectrique_kWhj").Cells(1, Colonne) = pd_ConsommationElectrique_kWhj
    Feuil13.Range("BA_faible_REPEREmd_ConsommationElectriqueAeration_kWhj").Cells(1, Colonne) = md_ConsommationElectriqueAeration_kWhj

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "13.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub effacer_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On efface les valeurs sur Excel
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


    Feuil13.Range("BA_faible_REPEREnominal_age_boue").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREmd_NominalConcentrationBassinsMES_gL").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREnominal_MV_MES").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREvolume_bassins").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREvolume_anaerobie").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREvolume_anoxie").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREvolume_aerobie").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREvolume_chenal").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREratio_elec_anoxie").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREratio_elec_aerobie").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREratio_elec_anaerobie").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREratio_elec_chenal").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPERErecirculation_MLSS_taux").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPERErecirculation_MLSS_P_refoulement").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPERErecirculation_MLSS_pompe_rdt").Cells(1, Colonne) = ""      'md_PompesRecirculationMLSSRendement
    Feuil13.Range("BA_faible_REPEREreel_age_boue").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREreel_MV_MES").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREreel_MES_concentration_bassin").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREO2_dissous").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREsortie_DBO").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREsortie_NH4").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREsortie_NO3").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPERErendement_bioP").Cells(1, Colonne) = ""    ' md_RendementDephosphatationBiologique
    Feuil13.Range("BA_faible_REPEREsortie_Pt").Cells(1, Colonne) = ""      'md_ConcentrationSortiePt_mgL
    Feuil13.Range("BA_faible_REPEREmethanol").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREFeCl3").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREhauteur_bassin").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREO2_besoin").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREO2_facteur_alpha").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREO2_rdt_transfert").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREair_Q_Nm3j").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREdiffuseur_encrassement").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREair_P_refoulement").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREsurpresseur_conso_spec").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREASB_eau_claire").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPERErecirculation_taux").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPERErecirculation_P_refoulement").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPERErecirculation_pompe_rdt").Cells(1, Colonne) = ""     'md_PompesRecirculationBouesRendement
    Feuil13.Range("BA_faible_REPEREmd_BouesISV_mL_g").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREmd_EauClarifieeMES_mgL").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREmd_HauteurClarificateur_m").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREmd_ClarificationVitesseMaximale_mh").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREmd_SurfaceTotaleClarification_m2").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREmd_ClarificationVitesseMaximaleRecalculee_mhulee").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREmd_ClarificateurChargeRadierNominal").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREmd_BouesExtraitesConcentrationMES_gL").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREpd_BouesExtraitesFluxMES_kgj").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREpd_BouesExtraitesDebit_m3j").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREmd_PompesBouesExcesPression").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREmd_PompesBouesExcesRendement").Cells(1, Colonne) = ""     'md_PompesBouesExcesRendement
    Feuil13.Range("BA_faible_REPEREpd_ConsommationElectrique_kWhj").Cells(1, Colonne) = ""
    Feuil13.Range("BA_faible_REPEREmd_ConsommationElectriqueAeration_kWhj").Cells(1, Colonne) = ""            'md_ConsommationElectriqueAeration_kWhj

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "13.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub affich_remarque()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 5  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil13.Range("BA_faible_REPEREerreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf
    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "13.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Private Sub Class_Initialize()

    If pn_EffacementResultatsEnCours = False Then
        Call lecture_choix
        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
            Call lecture_valeurs_forcees
        End If
        Call hypotheses
        Call attribution_valeur_par_defaut
    End If



End Sub

Private Sub Class_Terminate()

    If pn_EffacementResultatsEnCours = False Then
        'Call calcul_consommation_electrique
        Call ecriture_resultats
        Call affich_remarque
    End If

End Sub

