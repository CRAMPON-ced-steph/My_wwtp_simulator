Option Explicit

'variable concernant le fonctionnement en SHUNT ou en normal
Dim choix_fct_SHUNT As Boolean

'constantes pour le choix concernant la déphosphatation
Private Const nb_choix_dephosphatation = 4  'nombre de choix possibles
Private Const dephosphatation_non = 1
Private Const dephosphatation_bio = 2
Private Const dephosphatation_precipitation = 3
Private Const dephosphatation_co_precipitation = 4

'constantes pour le choix du type de bassin
Private Const nb_configuration_bassin = 2  'nombre de type de configurations de bassins
Private Const config_bassin_plugflow = 1
Private Const config_bassin_chenal = 2

'constantes pour le choix du type d'aérateur
Private Const nb_type_aerateur = 5  'nombre de type d'aérateurs
Private Const aerateur_fine_bulle = 1
Private Const aerateur_moyenne_bulle = 2
Private Const aerateur_brosse = 3
Private Const aerateur_turbine_lente = 4
Private Const aerateur_turbine_rapide = 5

'constantes pour le choix du type de régulation d'air
Private Const nb_type_regulation = 4  'nombre de type d'aérateurs
Private Const regulation_horloge = 1
Private Const regulation_sans_variation = 2
Private Const regulation_avec_variation = 3
Private Const regulation_avance = 4

'constantes pour le choix du type de surpresseur
Private Const nb_type_surpresseur = 3
Private Const surpresseur_roots = 1
Private Const surpresseur_vis = 2
Private Const surpresseur_turbo = 3
Private Const nb_surpresseur_gamme_Q = 1
Private Const surpresseur_gamme_Q_0_infini = 1
Private Const nb_surpresseur_gamme_P = 3
Private Const surpresseur_gamme_P_0_Px1 = 1
Private Const Px1 = 6    'mCE
Private Const surpresseur_gamme_P_Px1_Px2 = 2
Private Const Px2 = 8    'mCE
Private Const surpresseur_gamme_P_Px2_infini = 3

'constantes pour le choix du type de râcleur
Private Const nb_type_racleur = 3  'nombre de type de râcleurs
Private Const racleur_racle = 1
Private Const racleur_racle_suce = 2
Private Const racleur_Kruger = 3

Private Const critere_incoherence = 0.1

'variables de choix internes
Dim choix_dephosphatation As Integer
Dim choix_config_bassin As Integer
Dim choix_aerateur As Integer
Dim choix_regulation As Integer
Dim choix_surpresseur As Integer
Dim choix_racleur As Integer

'variables échangées avec Excel
Dim nominal_G As Double
Dim nominal_MES_concentration_bassin As Double
Public nominal_MV_MES As Double
Dim volume_bassins As Double
Dim volume_anaerobie As Double
Dim volume_anoxie As Double
Dim volume_aerobie As Double
Dim volume_chenal As Double
Dim ratio_elec_anoxie As Double
Dim ratio_elec_aerobie As Double
Dim ratio_elec_anaerobie As Double
Dim ratio_elec_chenal As Double
Dim reel_G As Double
Public reel_MV_MES As Double
Dim reel_MES_concentration_bassin As Double
Dim O2_dissous As Double
Dim sortie_DBO As Double
Dim sortie_NH4 As Double
Dim sortie_NO3 As Double
Dim rendement_bioP As Double
Dim sortie_Pt As Double
Public methanol_pur As Double
Public FeCl3_pur As Double
Dim hauteur_bassin As Double
Dim O2_besoin As Double
Dim O2_rdt_transfert As Double
Dim O2_facteur_alfa As Double
Dim air_Q_Nm3j As Double
Dim diffuseur_encrassement As Double
Dim air_P_refoulement As Double
Dim surpresseur_conso_spec As Double
Dim ASB_eau_claire As Double
Dim recirculation_taux As Double
Dim recirculation_P_refoulement As Double
Dim recirculation_pompe_rdt As Double
Dim recirculation_MLSS_taux As Double
Dim recirculation_MLSS_P_refoulement As Double
Dim recirculation_MLSS_pompe_rdt As Double
Dim mi_NombreClarificateurs As Integer
Dim boues_indice_Mohlman As Double
Dim sortie_MES As Double
Dim clarif_hauteur As Double
Dim clarif_vitesse_max As Double
Dim clarif_surface As Double
Dim clarif_vitesse_max_recalc As Double
Dim clarif_charge_radier_nominal As Double
Dim boues_concentration As Double
Public boues_MES As Double
Public boues_Q As Double
Dim extraction_P_refoulement As Double
Dim extraction_pompe_rdt As Double
Public electricite As Double
Dim electricite_aeration As Double

Public P_precipite As Double
Public electricite_fixe As Double
Dim besoins_O2_respiration As Double

'variables montrant si les valeurs ont été forcées
Dim nominal_G_force As Boolean
Dim nominal_MES_concentration_bassin_force As Boolean
Dim nominal_MV_MES_force As Boolean
Dim volume_bassins_force As Boolean
Dim volume_anaerobie_force As Boolean
Dim volume_anoxie_force As Boolean
Dim volume_aerobie_force As Boolean
Dim volume_chenal_force As Boolean
Dim ratio_elec_anoxie_force As Boolean
Dim ratio_elec_aerobie_force As Boolean
Dim ratio_elec_anaerobie_force As Boolean
Dim ratio_elec_chenal_force As Boolean
Dim reel_G_force As Boolean
Dim reel_MV_MES_force As Boolean
Dim reel_MES_concentration_bassin_force As Boolean
Dim O2_dissous_force As Boolean
Dim sortie_DBO_force As Boolean
Dim sortie_NH4_force As Boolean
Dim sortie_NO3_force As Boolean
Dim rendement_bioP_force As Boolean
Dim sortie_Pt_force As Boolean
Dim methanol_pur_force As Boolean
Dim FeCl3_pur_force As Boolean
Dim hauteur_bassin_force As Boolean
Dim O2_besoin_force As Boolean
Dim O2_rdt_transfert_force As Boolean
Dim O2_facteur_alfa_force As Boolean
Dim air_Q_Nm3j_force As Boolean
Dim diffuseur_encrassement_force As Boolean
Dim air_P_refoulement_force As Boolean
Dim surpresseur_conso_spec_force As Boolean
Dim ASB_eau_claire_force As Boolean
Dim recirculation_taux_force As Boolean
Dim recirculation_P_refoulement_force As Boolean
Dim recirculation_pompe_rdt_force As Boolean
Dim recirculation_MLSS_taux_force As Boolean
Dim recirculation_MLSS_P_refoulement_force As Boolean
Dim recirculation_MLSS_pompe_rdt_force As Boolean
Dim mn_NombreClarificateurs_force As Boolean
Dim boues_indice_Mohlman_force As Boolean
Dim sortie_MES_force As Boolean
Dim clarif_hauteur_force As Boolean
Dim clarif_vitesse_max_force As Boolean
Dim clarif_surface_force As Boolean
Dim boues_concentration_force As Boolean
Dim boues_MES_force As Boolean
Dim boues_Q_force As Boolean
Dim extraction_P_refoulement_force As Boolean
Dim extraction_pompe_rdt_force As Boolean

'valeurs forcées par l'utilisateur
Dim nominal_G_f As Double
Dim nominal_MES_concentration_bassin_f As Double
Dim nominal_MV_MES_f As Double
Dim volume_bassins_f As Double
Dim volume_anaerobie_f As Double
Dim volume_anoxie_f As Double
Dim volume_aerobie_f As Double
Dim volume_chenal_f As Double
Dim ratio_elec_anoxie_f As Double
Dim ratio_elec_aerobie_f As Double
Dim ratio_elec_anaerobie_f As Double
Dim ratio_elec_chenal_f As Double
Dim reel_G_f As Double
Dim reel_MV_MES_f As Double
Dim reel_MES_concentration_bassin_f As Double
Dim O2_dissous_f As Double
Dim sortie_DBO_f As Double
Dim sortie_NH4_f As Double
Dim sortie_NO3_f As Double
Dim rendement_bioP_f As Double
Dim sortie_Pt_f As Double
Dim methanol_pur_f As Double
Dim FeCl3_pur_f As Double
Dim hauteur_bassin_f As Double
Dim O2_besoin_f As Double
Dim O2_rdt_transfert_f As Double
Dim O2_facteur_alfa_f As Double
Dim air_Q_Nm3j_f As Double
Dim diffuseur_encrassement_f As Double
Dim air_P_refoulement_f As Double
Dim surpresseur_conso_spec_f As Double
Dim ASB_eau_claire_f As Double
Dim recirculation_taux_f As Double
Dim recirculation_P_refoulement_f As Double
Dim recirculation_pompe_rdt_f As Double
Dim recirculation_MLSS_taux_f As Double
Dim recirculation_MLSS_P_refoulement_f As Double
Dim recirculation_MLSS_pompe_rdt_f As Double
Dim mi_NombreClarificateurs_f As Integer
Dim boues_indice_Mohlman_f As Double
Dim sortie_MES_f As Double
Dim clarif_hauteur_f As Double
Dim clarif_vitesse_max_f As Double
Dim clarif_surface_f As Double
Dim boues_concentration_f As Double
Dim boues_MES_f As Double
Dim boues_Q_f As Double
Dim extraction_P_refoulement_f As Double
Dim extraction_pompe_rdt_f As Double

'variables des hypothèses
Dim T_reference As Double
Dim G_reference As Double
Dim MV_MES_reference As Double
Dim correctif_T As Double
Dim anaerobie_tps_passage_ref As Double
Dim ratio_DCO_NK_10 As Double
Dim a1_Cm As Double
Dim a0_Cm As Double
Dim a_rdt_DBO(3) As Double
Dim ratio_P_synthese As Double
Dim a_Pbio_coef_K As Double
Dim b_Pbio_coef_K As Double
Dim a_Kp_bio_1 As Double
Dim b_Kp_bio_1 As Double
Dim lim_Kp_bio_1 As Double
Dim rdt_Psol_max As Double
Dim ratio_Psol_P As Double
Dim ratio_FeCl3_Pt As Double
Dim a_dosis_MMS As Double
Dim b_dosis_MMS As Double
Dim rdt_DCO As Double
Dim ratio_DCO_sol_EB As Double
Dim rdt_DCO_sol_primaire As Double
Dim ratio_C_N_limite_nominal As Double
Dim ratio_C_N_limite_reel As Double
Dim DCO_methanol As Double
Dim ratio_boues_methanol As Double
Dim charge_radier_max(nb_type_racleur) As Double
Dim facteur_boues_C_lim As Double
Dim a0_O2_DBO_respiration As Double
Dim a1_O2_DBO_respiration As Double
Dim a2_O2_DBO_respiration As Double
Dim ratio_O2_nit As Double
Dim ratio_O2_denit As Double
Dim T_ref_ATV As Double
Dim correctif_T_ATV As Double
Dim O2_alpha0(nb_type_aerateur) As Double
Dim a0_alpha As Double
Dim a_alpha_MVS As Double
Dim a_alpha_G As Double
Dim O2_facteur_alfa_maxi As Double
Dim O2_facteur_alfa_mini As Double
Dim a0_correctif_alfa_MES As Double
Dim a1_correctif_alfa_MES As Double
Dim MES_reference_alfa As Double
Dim O2_facteur_beta As Double
Dim P0 As Double
Dim a0_Patm As Double
Dim a1_Patm As Double
Dim a2_Patm As Double
Dim a0_delta_P As Double
Dim a1_delta_P As Double
Dim hauteur_diffuseur As Double
Dim a_O2_saturation(4) As Double
Dim correctif_T_K As Double
Dim T_ref_K As Double
Dim ratio_kgO2_Nm3air As Double
Dim O2_rdt_dissolution(nb_type_aerateur) As Double
Dim ASB(nb_type_aerateur) As Double
'Dim impact_regulation(nb_type_regulation) As Double
Dim md_ConcentrationOxygeneDissous_mgL(1 To nb_type_regulation) As Double
Dim rate_puissance_fixe As Double
Dim rate_puissance_debit As Double
Dim agitation_fct As Double
Dim ratio_elec_recirculation As Double
Dim ratio_elec_recirculation_MLSS As Double
Dim ratio_elec_extraction As Double
Dim surpresseur_conso_Wh_Nm3mCE(nb_type_surpresseur, nb_surpresseur_gamme_Q, nb_surpresseur_gamme_P) As Double
Dim Pt_residuel As Double

'stockage de valeurs réelles
Dim stockage_MVS_bassins As Double
Dim stockage_Q As Double
Dim boues_chimiques As Double
Dim boues_methanol As Double

Dim remarque(5) As String


Sub dimensionnement(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef Sh, ByRef DCO_apportee_methanol, ByRef DCO_conso_bio)
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

    'Age de boues equivalent à T_reference
    G_equivalent = nominal_G * correctif_T ^ (pd_TemperatureEauDimensionnement_°C - T_reference)
    'Charge massique équivalente
    Cm_equivalente = 1 / (a1_Cm * G_equivalent + a0_Cm)
    'Charge massique à T_design
    Cm = Cm_equivalente * correctif_T ^ (pd_TemperatureEauDimensionnement_°C - T_reference)
    'Masse de boues dans les bassins kgMVS
    boues_bassins_MVS = DBO / Cm
    'on recalcule la part de MVS dans les MES
    boues_bassins_MES = boues_bassins_MVS / MV_MES_reference + (boues_chimiques + boues_methanol) * nominal_G
    nominal_MV_MES = (boues_bassins_MVS + boues_methanol * nominal_G) / boues_bassins_MES
    If nominal_MV_MES_force = True Then
        nominal_MV_MES = nominal_MV_MES_f
    End If
    'Volume de bassins m3
    volume_bassins = boues_bassins_MVS / (nominal_MV_MES * nominal_MES_concentration_bassin)
    If volume_bassins_force = True Then
        volume_bassins = volume_bassins_f
        'calcul de la Cm_equivalente avec ce volume de bassins
        boues_bassins_MVS = volume_bassins * nominal_MV_MES * nominal_MES_concentration_bassin
        Cm = DBO / boues_bassins_MVS
        Cm_equivalente = Cm / correctif_T ^ (pd_TemperatureEauDimensionnement_°C - T_reference)
    End If

    'volume anaérobie
    If choix_dephosphatation = dephosphatation_bio Or choix_dephosphatation = dephosphatation_co_precipitation Then
        volume_anaerobie = Q / anaerobie_tps_passage_ref * correctif_T ^ (T_reference - pd_TemperatureEauDimensionnement_°C)
        If volume_anaerobie_force = True Then
            volume_anaerobie = volume_anaerobie_f
        End If
    Else
        volume_anaerobie = 0
    End If

    'autres volumes
    If choix_config_bassin = config_bassin_plugflow Then
        If DCO / NK < ratio_DCO_NK_10 Then
            volume_anoxie = 0.4 * (volume_bassins - volume_anaerobie)
        Else
            volume_anoxie = 0.2 * (volume_bassins - volume_anaerobie)
        End If
        If volume_anoxie_force = True Then
            volume_anoxie = volume_anoxie_f
        End If
        volume_aerobie = volume_bassins - (volume_anoxie + volume_anaerobie)
        If volume_aerobie_force = True Then
            volume_aerobie = volume_aerobie_f
        End If
        volume_chenal = 0
    Else
        volume_chenal = volume_bassins - volume_anaerobie
        If volume_chenal_force = True Then
            volume_chenal = volume_chenal_f
        End If
        volume_anoxie = 0
        volume_aerobie = 0
    End If
    If volume_bassins <> (volume_anaerobie + volume_anoxie + volume_aerobie + volume_chenal) Then
        remarque(1) = Feuil3.Range("BA_prolongee_erreur_volume").Cells(1, 1)
    Else
        remarque(1) = ""
    End If

    'surface de radier du clarificateur (m2)
    clarif_surface = (Q_nominal * pd_CoefficientPointeHydrauliqueTempsPluie + (Q - Q_nominal)) / (pd_NOMBRE_HEURE_PAR_JOUR * clarif_vitesse_max)
    If clarif_surface_force = True Then
        clarif_surface = clarif_surface_f
    End If


    'FONCTIONNEMENT NOMINAL

    'rendement d'élimination de la DBO
    rdt_DBO = 0
    For I = 0 To 3
        rdt_DBO = rdt_DBO + a_rdt_DBO(I) * (Cm_equivalente / nominal_MV_MES) ^ I
    Next I

    'traitement de l'azote
    nitrification = NH4 - sortie_NH4 * Q / 1000
    denitrification = NO3 + nitrification - sortie_NO3 * Q / 1000
    'calcul de l'ajout de méthanol pour la dénitrification
    DCO_soluble = DCO_nominal * ratio_DCO_sol_EB * (1 - rdt_DCO_sol_primaire)
    ratio_C_N = DCO_soluble / denitrification
    If ratio_C_N < ratio_C_N_limite_nominal Then
        DCO_necessaire = ratio_C_N_limite_nominal * denitrification - DCO_soluble
        methanol_pur = DCO_necessaire / DCO_methanol
    End If
    DCO_necessaire = methanol_pur * DCO_methanol
    DCO_apportee_methanol = DCO_necessaire
    boues_methanol = ratio_boues_methanol * DCO_necessaire
    DCO = DCO + DCO_necessaire * (1 - ratio_boues_methanol)

    'traitement du phosphore
    P_synthese = ratio_P_synthese * DBO * rdt_DBO
    If choix_dephosphatation = dephosphatation_bio Then
        If rendement_bioP_force = True Then
            rendement_bioP = rendement_bioP_f
        Else
            P_bio_coeff_K = a_Pbio_coef_K + b_Pbio_coef_K / G_equivalent
            ratio_DBO_P_in = DBO / Pt
            volume_optimal = Q / anaerobie_tps_passage_ref * correctif_T ^ (T_reference - pd_TemperatureEauConditionsReelles_°C)
            Kp_bio_1 = volume_anaerobie / volume_optimal
            If Kp_bio_1 <= 1 Then
                Kp_bio_1 = Kp_bio_1 ^ 2
            ElseIf Kp_bio_1 <= 1.5 Then
                Kp_bio_1 = a_Kp_bio_1 + b_Kp_bio_1 * Kp_bio_1
            Else
                Kp_bio_1 = lim_Kp_bio_1
            End If
            Kp_bio_2 = 1
            rendement_bioP = P_bio_coeff_K * Kp_bio_1 * Kp_bio_2 * ratio_DBO_P_in
            If rendement_bioP > rdt_Psol_max Then
                rendement_bioP = rdt_Psol_max
            End If
        End If
        Pt = Pt - rendement_bioP * ratio_Psol_P * Pt
        sortie_Pt = Pt / (Q - boues_Q) * 1000
        FeCl3_pur = 0
        boues_chimiques = 0
        P_precipite = 0
    ElseIf choix_dephosphatation = dephosphatation_co_precipitation Then
        If rendement_bioP_force = True Then
            rendement_bioP = rendement_bioP_f
        Else
            P_bio_coeff_K = a_Pbio_coef_K + b_Pbio_coef_K / G_equivalent
            ratio_DBO_P_in = DBO / Pt
            volume_optimal = Q / anaerobie_tps_passage_ref * correctif_T ^ (T_reference - pd_TemperatureEauConditionsReelles_°C)
            Kp_bio_1 = volume_anaerobie / volume_optimal
            If Kp_bio_1 <= 1 Then
                Kp_bio_1 = Kp_bio_1 ^ 2
            ElseIf Kp_bio_1 <= 1.5 Then
                Kp_bio_1 = a_Kp_bio_1 + b_Kp_bio_1 * Kp_bio_1
            Else
                Kp_bio_1 = lim_Kp_bio_1
            End If
            Kp_bio_2 = 1
            rendement_bioP = P_bio_coeff_K * Kp_bio_1 * Kp_bio_2 * ratio_DBO_P_in
            If rendement_bioP > rdt_Psol_max Then
                rendement_bioP = rdt_Psol_max
            End If
        End If
        Pt = Pt - rendement_bioP * ratio_Psol_P * Pt
        Pt = Pt - P_synthese
        P_precipitation = Pt - sortie_Pt * (Q - boues_Q) / 1000
        If P_precipitation > 0 Then
            Ptbr = P_precipitation / Q * 1000
            ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + sortie_Pt) ^ (-0.5) - ((sortie_Pt) ^ (-0.5) - 0.1)) / ((Ptbr + sortie_Pt) - (sortie_Pt - 0.1))
            FeCl3_pur = ratio_molaire_Ptbr * ratio_FeCl3_Pt * P_precipitation
        Else
            P_precipitation = 0
            FeCl3_pur = 0
        End If
        dose_MMS = (ratio_molaire_Ptbr - 1) * a_dosis_MMS + b_dosis_MMS
        boues_chimiques = dose_MMS * P_precipitation
        Pt = Pt - P_precipitation
        P_precipite = P_precipitation
    ElseIf choix_dephosphatation = dephosphatation_precipitation Then
        Pt = Pt - P_synthese
        P_precipitation = Pt - sortie_Pt * (Q - boues_Q) / 1000
        If P_precipitation > 0 Then
            Ptbr = P_precipitation / Q * 1000
            ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + sortie_Pt) ^ (-0.5) - ((sortie_Pt) ^ (-0.5) - 0.1)) / ((Ptbr + sortie_Pt) - (sortie_Pt - 0.1))
            FeCl3_pur = ratio_molaire_Ptbr * ratio_FeCl3_Pt * P_precipitation
        Else
            P_precipitation = 0
            FeCl3_pur = 0
        End If
        dose_MMS = (ratio_molaire_Ptbr - 1) * a_dosis_MMS + b_dosis_MMS
        boues_chimiques = dose_MMS * P_precipitation
        Pt = Pt - P_precipitation
        P_precipite = P_precipitation
    Else
        Pt = Pt - P_synthese
        FeCl3_pur = 0
        boues_chimiques = 0
        P_precipite = 0
    End If


    'ratio O2 respiration et élimination DBO5
    facteur_T_O2 = correctif_T_ATV ^ (pd_TemperatureEauConditionsReelles_°C - T_ref_ATV)
    ratio_O2_DBO_respiration = a0_O2_DBO_respiration + (a1_O2_DBO_respiration * reel_G * facteur_T_O2) / (1 + a2_O2_DBO_respiration * reel_G * facteur_T_O2)

    'calcul des besoins en O2
    O2_besoin = ratio_O2_DBO_respiration * rdt_DBO * DBO
    O2_besoin = O2_besoin + ratio_O2_nit * nitrification
    O2_besoin = O2_besoin + ratio_O2_denit * denitrification
    O2_besoin = O2_besoin + methanol_pur * DCO_methanol  'on ne consomme pas le méthanol directement mais il représente de la DCO non consommée par la dénit
    'DCO_conso_bio = ratio_O2_DBO_respiration * (DBO - sortie_DBO * stockage_Q / 1000) + DCO_apportee_methanol * (1 - ratio_boues_methanol) * rdt_DCO - ratio_O2_denit * denitrification

    'Production de boues
    boues_produites = boues_bassins_MVS / (nominal_MV_MES * nominal_G)

    'vitesse hydraulique max (recalculée) clarificateur
    clarif_vitesse_max_recalc = (Q_nominal * pd_CoefficientPointeHydrauliqueTempsPluie + (Q - Q_nominal)) / (pd_NOMBRE_HEURE_PAR_JOUR * clarif_surface)
    'charge au radier clarificateur
    clarif_charge_radier_nominal = nominal_MES_concentration_bassin * Q * (1 + recirculation_taux) / (pd_NOMBRE_HEURE_PAR_JOUR * clarif_surface)
    If clarif_charge_radier_nominal > charge_radier_max(choix_racleur) Then
        remarque(2) = Feuil3.Range("BA_prolongee_erreur_charge_radier").Cells(1, 1)
    Else
        remarque(2) = ""
    End If


    'CALCUL DES SORTIES

    'Eau traitée
    Q = Q - boues_Q
    MES = sortie_MES * Q / 1000
    DCO = DCO * (1 - rdt_DCO)
    DBO = DBO * (1 - rdt_DBO)
    NH4 = NH4 - nitrification
    NK = NH4

    'Boues
    boues_MES = boues_produites - MES
    boues_concentration = nominal_MES_concentration_bassin * (1 + recirculation_taux) / recirculation_taux
    boues_Q = boues_MES / boues_concentration

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "3.5"

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

    stockage_Q = Q


    'FONCTIONNEMENT REEL

    'Age de boues equivalent à T_reference
    G_equivalent = reel_G * correctif_T ^ (pd_TemperatureEauConditionsReelles_°C - T_reference)
    'Charge massique équivalente
    Cm_equivalente = 1 / (a1_Cm * G_equivalent + a0_Cm)
    'Charge massique à T_design
    Cm = Cm_equivalente * correctif_T ^ (pd_TemperatureEauConditionsReelles_°C - T_reference)
    'Masse de boues dans les bassins kgMVS
    boues_bassins_MVS = DBO / Cm
    'on recalcule la part de MVS dans les MES
    boues_bassins_MES = boues_bassins_MVS / MV_MES_reference + (boues_chimiques + boues_methanol) * reel_G
    reel_MV_MES = (boues_bassins_MVS + reel_G * boues_methanol) / boues_bassins_MES
    If reel_MV_MES_force = True Then
        reel_MV_MES = reel_MV_MES_f
    End If
    'Calcul de la concentration en MES dans les bassins
    reel_MES_concentration_bassin = boues_bassins_MVS / (reel_MV_MES * volume_bassins)
    If reel_MES_concentration_bassin_force = True Then
        reel_MES_concentration_bassin = reel_MES_concentration_bassin_f
        'calcul de la Cm équivalente avec cette concentration
        boues_bassins_MVS = volume_bassins * reel_MV_MES * reel_MES_concentration_bassin
        Cm = DBO / boues_bassins_MVS
        Cm_equivalente = Cm / correctif_T ^ (pd_TemperatureEauConditionsReelles_°C - T_reference)
    End If

    stockage_MVS_bassins = boues_bassins_MVS

    'TRAITEMENT REEL

    'rendement d'élimination de la DBO
    rdt_DBO = 0
    For I = 0 To 3
        rdt_DBO = rdt_DBO + a_rdt_DBO(I) * (Cm_equivalente / reel_MV_MES) ^ I
    Next I
    sortie_DBO = DBO * (1 - rdt_DBO) / Q * 1000
    If sortie_DBO_force = True Then
        sortie_DBO = sortie_DBO_f
    End If

    'traitement de l'azote
    nitrification = NH4 - sortie_NH4 * Q / 1000
    denitrification = NO3 + nitrification - sortie_NO3 * Q / 1000
    'calcul de la recirculation de liqueur mixte approppriée à ce traitement si cuves séparées
    If choix_config_bassin = config_bassin_plugflow Then
        'VU avec Kim Sorensen le 13/05/2011
        'rdt_denit = ((NK + NO3) - (sortie_NH4 + sortie_NO3 + 1) * Q / 1000) / (NK + NO3)    '1 correspond à l'azote réduit autre que NH4
        'recirculation_MLSS_taux = rdt_denit / (1 - rdt_denit)
        recirculation_MLSS_taux = nitrification * 1000 / (Q * sortie_NO3) - (1 + recirculation_taux)
        If recirculation_MLSS_taux_force = True Then
            If recirculation_MLSS_taux > recirculation_MLSS_taux_f Then
                remarque(4) = Feuil3.Range("BA_prolongee_erreur_taux_recirc_MLSS").Cells(1, 1)
            Else
                remarque(4) = ""
            End If
            recirculation_MLSS_taux = recirculation_MLSS_taux_f
        End If
    Else
        recirculation_MLSS_taux = 0
    End If
    'calcul de l'ajout de méthanol pour la dénitrification
    DCO_soluble = DCO_nominal * pd_PourcentageChargeReelleDCO * ratio_DCO_sol_EB * (1 - rdt_DCO_sol_primaire)
    ratio_C_N = DCO_soluble / denitrification
    If ratio_C_N < ratio_C_N_limite_reel Then
        DCO_necessaire = ratio_C_N_limite_reel * denitrification - DCO_soluble
        methanol_pur = DCO_necessaire / DCO_methanol
    End If
    If methanol_pur_force = True Then
        methanol_pur = methanol_pur_f
    End If
    DCO_necessaire = methanol_pur * DCO_methanol
    DCO_apportee_methanol = DCO_necessaire
    boues_methanol = ratio_boues_methanol * DCO_necessaire
    DCO = DCO + DCO_necessaire * (1 - ratio_boues_methanol)

    'traitement du phosphore
    P_synthese = ratio_P_synthese * DBO * rdt_DBO
    If choix_dephosphatation = dephosphatation_bio Then
        If rendement_bioP_force = True Then
            rendement_bioP = rendement_bioP_f
        Else
            P_bio_coeff_K = a_Pbio_coef_K + b_Pbio_coef_K / G_equivalent
            ratio_DBO_P_in = DBO / Pt
            volume_optimal = Q / anaerobie_tps_passage_ref * correctif_T ^ (T_reference - pd_TemperatureEauConditionsReelles_°C)
            Kp_bio_1 = volume_anaerobie / volume_optimal
            If Kp_bio_1 <= 1 Then
                Kp_bio_1 = Kp_bio_1 ^ 2
            ElseIf Kp_bio_1 <= 1.5 Then
                Kp_bio_1 = a_Kp_bio_1 + b_Kp_bio_1 * Kp_bio_1
            Else
                Kp_bio_1 = lim_Kp_bio_1
            End If
            Kp_bio_2 = 1
            rendement_bioP = P_bio_coeff_K * Kp_bio_1 * Kp_bio_2 * ratio_DBO_P_in
            If rendement_bioP > rdt_Psol_max Then
                rendement_bioP = rdt_Psol_max
            End If
        End If
        Pt = Pt - rendement_bioP * ratio_Psol_P * Pt
        sortie_Pt = Pt / (Q - boues_Q) * 1000
        FeCl3_pur = 0
        boues_chimiques = 0
        P_precipite = 0
    ElseIf choix_dephosphatation = dephosphatation_co_precipitation Then
        If rendement_bioP_force = True Then
            rendement_bioP = rendement_bioP_f
        Else
            P_bio_coeff_K = a_Pbio_coef_K + b_Pbio_coef_K / G_equivalent
            ratio_DBO_P_in = DBO / Pt
            volume_optimal = Q / anaerobie_tps_passage_ref * correctif_T ^ (T_reference - pd_TemperatureEauConditionsReelles_°C)
            Kp_bio_1 = volume_anaerobie / volume_optimal
            If Kp_bio_1 <= 1 Then
                Kp_bio_1 = Kp_bio_1 ^ 2
            ElseIf Kp_bio_1 <= 1.5 Then
                Kp_bio_1 = a_Kp_bio_1 + b_Kp_bio_1 * Kp_bio_1
            Else
                Kp_bio_1 = lim_Kp_bio_1
            End If
            Kp_bio_2 = 1
            rendement_bioP = P_bio_coeff_K * Kp_bio_1 * Kp_bio_2 * ratio_DBO_P_in
            If rendement_bioP > rdt_Psol_max Then
                rendement_bioP = rdt_Psol_max
            End If
        End If
        Pt = Pt - rendement_bioP * ratio_Psol_P * Pt
        Pt = Pt - P_synthese
        P_precipitation = Pt - sortie_Pt * (Q - boues_Q) / 1000
        If P_precipitation > 0 Then
            Ptbr = P_precipitation / Q * 1000
            ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + sortie_Pt) ^ (-0.5) - ((sortie_Pt) ^ (-0.5) - 0.1)) / ((Ptbr + sortie_Pt) - (sortie_Pt - 0.1))
            FeCl3_pur = ratio_molaire_Ptbr * ratio_FeCl3_Pt * P_precipitation
        Else
            P_precipitation = 0
            FeCl3_pur = 0
        End If
        If FeCl3_pur_force = True Then
            FeCl3_pur = FeCl3_pur_f
        End If
        dose_MMS = (ratio_molaire_Ptbr - 1) * a_dosis_MMS + b_dosis_MMS
        boues_chimiques = dose_MMS * P_precipitation
        Pt = Pt - P_precipitation
        P_precipite = P_precipitation
    ElseIf choix_dephosphatation = dephosphatation_precipitation Then
        Pt = Pt - P_synthese
        P_precipitation = Pt - sortie_Pt * (Q - boues_Q) / 1000
        If P_precipitation > 0 Then
            Ptbr = P_precipitation / Q * 1000
            ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + sortie_Pt) ^ (-0.5) - ((sortie_Pt) ^ (-0.5) - 0.1)) / ((Ptbr + sortie_Pt) - (sortie_Pt - 0.1))
            FeCl3_pur = ratio_molaire_Ptbr * ratio_FeCl3_Pt * P_precipitation
        Else
            P_precipitation = 0
            FeCl3_pur = 0
        End If
        If FeCl3_pur_force = True Then
            FeCl3_pur = FeCl3_pur_f
        End If
        dose_MMS = (ratio_molaire_Ptbr - 1) * a_dosis_MMS + b_dosis_MMS
        boues_chimiques = dose_MMS * P_precipitation
        Pt = Pt - P_precipitation
        P_precipite = P_precipitation
    Else
        Pt = Pt - P_synthese
        FeCl3_pur = 0
        boues_chimiques = 0
        P_precipite = 0
    End If

    'ratio O2 respiration et élimination DBO5
    facteur_T_O2 = correctif_T_ATV ^ (pd_TemperatureEauConditionsReelles_°C - T_ref_ATV)
    ratio_O2_DBO_respiration = a0_O2_DBO_respiration + (a1_O2_DBO_respiration * reel_G * facteur_T_O2) / (1 + a2_O2_DBO_respiration * reel_G * facteur_T_O2)

    'calcul des besoins en O2
    besoins_O2_respiration = ((a1_O2_DBO_respiration * reel_G * facteur_T_O2) / (1 + a2_O2_DBO_respiration * reel_G * facteur_T_O2)) * (DBO - sortie_DBO * stockage_Q / 1000)
    O2_besoin = ratio_O2_DBO_respiration * (DBO - sortie_DBO * stockage_Q / 1000)
    O2_besoin = O2_besoin + ratio_O2_nit * nitrification
    O2_besoin = O2_besoin + ratio_O2_denit * denitrification
    O2_besoin = O2_besoin + methanol_pur * DCO_methanol  'on ne consomme pas le méthanol directement mais il représente de la DCO non consommée par la dénit
    'DCO_conso_bio = ratio_O2_DBO_respiration * (DBO - sortie_DBO * stockage_Q / 1000) + DCO_apportee_methanol * (1 - ratio_boues_methanol) * rdt_DCO - ratio_O2_denit * denitrification
    If O2_besoin_force = True Then
        delta_O2 = O2_besoin_f - O2_besoin
        O2_besoin = O2_besoin_f
        'DCO_conso_bio = DCO_conso_bio + delta_O2
    End If



    'Production de boues
    boues_produites = reel_MES_concentration_bassin * volume_bassins / reel_G


    'CALCUL DES SORTIES

    'Eau traitée
    Q = Q - boues_Q
    MES = sortie_MES * Q / 1000
    DCO = DCO * (1 - rdt_DCO)
    DBO = sortie_DBO * Q / 1000
    NH4 = sortie_NH4 * Q / 1000
    NK = NH4
    NO3 = sortie_NO3 * Q / 1000

    'Boues
    boues_MES = boues_produites - MES
    If boues_MES_force = True Then
        boues_MES = boues_MES_f
    End If
    boues_concentration = reel_MES_concentration_bassin * (1 + recirculation_taux) / recirculation_taux
    If boues_concentration_force = True Then
        boues_concentration = boues_concentration_f
    End If
    boues_Q = boues_MES / boues_concentration

    If boues_concentration > facteur_boues_C_lim / boues_indice_Mohlman Then
        remarque(3) = Feuil3.Range("BA_prolongee_erreur_taux_recirc").Cells(1, 1)
        If recirculation_taux_force = False Then
            recirculation_taux = recirculation_taux + 0.1
        End If
    Else
        remarque(3) = ""
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "3.6"

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

Sub calcul_consommation_electrique(ByRef electricite_II_aeration)
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
    Dim electricite_racleur As Double
    Dim electricite_agitation As Double
    Dim electricite_recirculation_MLSS As Double
    Dim electricite_recirculation As Double
    Dim electricite_extraction As Double


    'AERATION

    'calcul facteur alpha
    If choix_aerateur = aerateur_fine_bulle Then
        'source: Oxygen transfer in AS (Henkel,Cornel and Wagner)
        O2_facteur_alfa = a0_alpha + a_alpha_MVS * (stockage_MVS_bassins / volume_bassins) + a_alpha_G * (reel_G * correctif_T ^ (pd_TemperatureEauConditionsReelles_°C - T_reference))
        If O2_facteur_alfa > O2_facteur_alfa_maxi Then
            O2_facteur_alfa = O2_facteur_alfa_maxi
        ElseIf O2_facteur_alfa < O2_facteur_alfa_mini Then
            O2_facteur_alfa = O2_facteur_alfa_mini
        End If
    Else
        O2_facteur_alfa = O2_alpha0(choix_aerateur)
        If reel_MES_concentration_bassin > MES_reference_alfa Then
            O2_facteur_alfa = O2_facteur_alfa * (a0_correctif_alfa_MES + a1_correctif_alfa_MES * reel_MES_concentration_bassin) / (a0_correctif_alfa_MES + a1_correctif_alfa_MES * MES_reference_alfa)
        End If
    End If
    If O2_facteur_alfa_force = True Then
        O2_facteur_alfa = O2_facteur_alfa_f
    End If

    'calcul O2 saturation à (T,P=Patm+delta_P) et (20,P0)
    Patm = P0 * ((a0_Patm + a1_Patm * altitude) ^ a2_Patm) / 100    'en hPa
    If choix_aerateur = aerateur_fine_bulle Or choix_aerateur = aerateur_moyenne_bulle Then
        delta_P = a0_delta_P / a1_delta_P * P0 / 100 * (hauteur_bassin - hauteur_diffuseur)
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
    O2_facteur_K = O2_facteur_alfa * (O2_facteur_beta * O2_saturation_T_P - O2_dissous) / O2_saturation_20_P
    O2_facteur_K = O2_facteur_K * correctif_T_K ^ (pd_TemperatureEauConditionsReelles_°C - T_ref_K)

    'calcul consommations sur l'aération
    If choix_aerateur = aerateur_fine_bulle Or choix_aerateur = aerateur_moyenne_bulle Then
        'débit d'air
        O2_rdt_transfert = O2_rdt_dissolution(choix_aerateur)
        If O2_rdt_transfert_force = True Then
            O2_rdt_transfert = O2_rdt_transfert_f
        End If
        air_Q_Nm3j = O2_besoin / (O2_facteur_K * O2_rdt_transfert / 100 * (hauteur_bassin - hauteur_diffuseur) * ratio_kgO2_Nm3air)
        If air_Q_Nm3j_force = True Then
            air_Q_Nm3j = air_Q_Nm3j_f
        End If
        If surpresseur_conso_spec_force = True Then
            surpresseur_conso_spec = surpresseur_conso_spec_f
            'gamme de pression
            If air_P_refoulement <= Px1 Then
                gamme_P = surpresseur_gamme_P_0_Px1
            ElseIf air_P_refoulement <= Px2 Then
                gamme_P = surpresseur_gamme_P_Px1_Px2
            Else
                gamme_P = surpresseur_gamme_P_Px2_infini
                If choix_surpresseur = surpresseur_roots Then
                    remarque(5) = Feuil3.Range("BA_prolongee_erreur_choix_surpresseur").Cells(1, 1)
                End If
            End If
        Else
            'gamme de pression
            If air_P_refoulement <= Px1 Then
                gamme_P = surpresseur_gamme_P_0_Px1
            ElseIf air_P_refoulement <= Px2 Then
                gamme_P = surpresseur_gamme_P_Px1_Px2
            Else
                gamme_P = surpresseur_gamme_P_Px2_infini
                If choix_surpresseur = surpresseur_roots Then
                    remarque(5) = Feuil3.Range("BA_prolongee_erreur_choix_surpresseur").Cells(1, 1)
                End If
            End If
            'gamme de débit   (pour le nominal)
            gamme_Q = surpresseur_gamme_Q_0_infini
'            surpresseur_conso_spec = surpresseur_conso_Wh_Nm3mCE(choix_surpresseur, gamme_Q, gamme_P) * (1 + impact_regulation(choix_regulation))
            surpresseur_conso_spec = surpresseur_conso_Wh_Nm3mCE(choix_surpresseur, gamme_Q, gamme_P)
        End If
        'calcul de la consommation électrique de l'aération
        electricite_aeration = air_Q_Nm3j * air_P_refoulement * surpresseur_conso_spec / 1000
    Else
        'calcul de la consommation électrique de l'aération
        electricite_aeration = O2_besoin / (ASB_eau_claire * O2_facteur_K)
    End If
    electricite_II_aeration = electricite_II_aeration + electricite_aeration


    'CLARIFICATEUR
    'electricite_racleur = rate_puissance_fixe * pd_NOMBRE_HEURE_PAR_JOUR + rate_puissance_debit * stockage_Q / 1000
    'DCA le 15/07/11 données Pascal JOUAFFRE  de 12 à 32m diamètre --> 0.55 kW installés, de 33 à 52m --> 0.75 kW
    clarif_diametre_limite = 32
    If clarif_surface / mi_NombreClarificateurs < pd_NOMBRE_PI * (clarif_diametre_limite / 2) ^ 2 Then
        electricite_racleur = mi_NombreClarificateurs * 0.55 * pd_NOMBRE_HEURE_PAR_JOUR
    Else
        electricite_racleur = mi_NombreClarificateurs * 0.75 * pd_NOMBRE_HEURE_PAR_JOUR
    End If


    'AGITATION
    If choix_dephosphatation = dephosphatation_bio Or choix_dephosphatation = dephosphatation_co_precipitation Then
        electricite_agitation = ratio_elec_anaerobie * volume_anaerobie * agitation_fct / 1000
        If choix_config_bassin = config_bassin_plugflow Then
            electricite_agitation = electricite_agitation + ratio_elec_anoxie * volume_anoxie * agitation_fct / 1000
            electricite_agitation = electricite_agitation + ratio_elec_aerobie * volume_aerobie * agitation_fct / 1000
        Else
            electricite_agitation = electricite_agitation + ratio_elec_chenal * volume_chenal * agitation_fct / 1000
            'electricite_agitation = electricite_agitation + ratio_elec_anoxie * (volume_bassins - volume_anaerobie - volume_chenal) * agitation_fct / 1000
        End If
    Else
        If choix_config_bassin = config_bassin_plugflow Then
            electricite_agitation = ratio_elec_anoxie * volume_anoxie * agitation_fct / 1000
            electricite_agitation = electricite_agitation + ratio_elec_aerobie * volume_aerobie * agitation_fct / 1000
        Else
            electricite_agitation = ratio_elec_chenal * volume_chenal * agitation_fct / 1000
            'electricite_agitation = electricite_agitation + ratio_elec_anoxie * (volume_bassins - volume_chenal) * agitation_fct / 1000
        End If
    End If


    'POMPES DE RECIRCULATION ET EXTRACTION
    electricite_recirculation_MLSS = ratio_elec_recirculation_MLSS * recirculation_MLSS_taux * stockage_Q * recirculation_MLSS_P_refoulement
    electricite_recirculation = ratio_elec_recirculation * recirculation_taux * stockage_Q * recirculation_P_refoulement
    electricite_extraction = ratio_elec_extraction * boues_Q * extraction_P_refoulement


    'CONSOMMATION GLOBALE D'ELECTRICITE
    electricite = electricite_aeration + electricite_racleur + electricite_agitation + electricite_recirculation_MLSS + electricite_recirculation + electricite_extraction
    electricite_fixe = besoins_O2_respiration / O2_besoin * electricite_aeration + electricite_racleur + electricite_agitation

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "3.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub lecture_choix()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    choix_fct_SHUNT = Feuil57.Range("choix_fct_SHUNT").Cells(1, pi_FiliereConsideree)

    'lecture des choix internes
    choix_dephosphatation = Feuil3.Range("BA_prolongee_choix_dephosphatation").Cells(1, pi_FiliereConsideree)
    choix_config_bassin = Feuil3.Range("BA_prolongee_choix_configuration").Cells(1, pi_FiliereConsideree)
    choix_aerateur = Feuil3.Range("BA_prolongee_choix_aerateur").Cells(1, pi_FiliereConsideree)
    choix_regulation = Feuil3.Range("BA_prolongee_choix_regulation").Cells(1, pi_FiliereConsideree)
    choix_surpresseur = Feuil3.Range("BA_prolongee_choix_surpresseur").Cells(1, pi_FiliereConsideree)
    choix_racleur = Feuil3.Range("BA_prolongee_choix_racleur").Cells(1, pi_FiliereConsideree)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "3.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub lecture_valeurs_forcees()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'lecture du choix de forcer des valeurs
    nominal_G_force = Feuil3.Range("BA_prolongee_nominal_age_boue").Cells(1, pi_FiliereConsideree)
    nominal_MES_concentration_bassin_force = Feuil3.Range("BA_prolongee_nominal_MES_concentration_bassin").Cells(1, pi_FiliereConsideree)
    nominal_MV_MES_force = Feuil3.Range("BA_prolongee_nominal_MV_MES").Cells(1, pi_FiliereConsideree)
    volume_bassins_force = Feuil3.Range("BA_prolongee_volume_bassins").Cells(1, pi_FiliereConsideree)
    If choix_dephosphatation = dephosphatation_bio Or choix_dephosphatation = dephosphatation_co_precipitation Then
        volume_anaerobie_force = Feuil3.Range("BA_prolongee_volume_anaerobie").Cells(1, pi_FiliereConsideree)
        ratio_elec_anaerobie_force = Feuil3.Range("BA_prolongee_ratio_elec_anaerobie").Cells(1, pi_FiliereConsideree)
        rendement_bioP_force = Feuil3.Range("BA_prolongee_rendement_bioP").Cells(1, pi_FiliereConsideree)
    End If
    If choix_config_bassin = config_bassin_plugflow Then
        volume_anoxie_force = Feuil3.Range("BA_prolongee_volume_anoxie").Cells(1, pi_FiliereConsideree)
        ratio_elec_anoxie_force = Feuil3.Range("BA_prolongee_ratio_elec_anoxie").Cells(1, pi_FiliereConsideree)
        volume_aerobie_force = Feuil3.Range("BA_prolongee_volume_aerobie").Cells(1, pi_FiliereConsideree)
        ratio_elec_aerobie_force = Feuil3.Range("BA_prolongee_ratio_elec_aerobie").Cells(1, pi_FiliereConsideree)
        recirculation_MLSS_taux_force = Feuil3.Range("BA_prolongee_recirculation_MLSS_taux").Cells(1, pi_FiliereConsideree)
        recirculation_MLSS_P_refoulement_force = Feuil3.Range("BA_prolongee_recirculation_MLSS_P_refoulement").Cells(1, pi_FiliereConsideree)
        recirculation_MLSS_pompe_rdt_force = Feuil3.Range("BA_prolongee_recirculation_MLSS_pompe_rdt").Cells(1, pi_FiliereConsideree)
    Else
        volume_chenal_force = Feuil3.Range("BA_prolongee_volume_chenal").Cells(1, pi_FiliereConsideree)
        ratio_elec_chenal_force = Feuil3.Range("BA_prolongee_ratio_elec_chenal").Cells(1, pi_FiliereConsideree)
    End If
    reel_G_force = Feuil3.Range("BA_prolongee_reel_age_boue").Cells(1, pi_FiliereConsideree)
    reel_MV_MES_force = Feuil3.Range("BA_prolongee_reel_MV_MES").Cells(1, pi_FiliereConsideree)
    reel_MES_concentration_bassin_force = Feuil3.Range("BA_prolongee_reel_MES_concentration_bassin").Cells(1, pi_FiliereConsideree)
    O2_dissous_force = Feuil3.Range("BA_prolongee_O2_dissous").Cells(1, pi_FiliereConsideree)
    sortie_DBO_force = Feuil3.Range("BA_prolongee_sortie_DBO").Cells(1, pi_FiliereConsideree)
    sortie_NH4_force = Feuil3.Range("BA_prolongee_sortie_NH4").Cells(1, pi_FiliereConsideree)
    sortie_NO3_force = Feuil3.Range("BA_prolongee_sortie_NO3").Cells(1, pi_FiliereConsideree)
    methanol_pur_force = Feuil3.Range("BA_prolongee_methanol").Cells(1, pi_FiliereConsideree)
    If choix_dephosphatation = dephosphatation_precipitation Or choix_dephosphatation = dephosphatation_co_precipitation Then
        sortie_Pt_force = Feuil3.Range("BA_prolongee_sortie_Pt").Cells(1, pi_FiliereConsideree)
        FeCl3_pur_force = Feuil3.Range("BA_prolongee_FeCl3").Cells(1, pi_FiliereConsideree)
    End If
    FeCl3_pur_force = Feuil3.Range("BA_prolongee_FeCl3").Cells(1, pi_FiliereConsideree)
    hauteur_bassin_force = Feuil3.Range("BA_prolongee_hauteur_bassin").Cells(1, pi_FiliereConsideree)
    O2_besoin_force = Feuil3.Range("BA_prolongee_O2_besoin").Cells(1, pi_FiliereConsideree)
    O2_facteur_alfa_force = Feuil3.Range("BA_prolongee_O2_facteur_alpha").Cells(1, pi_FiliereConsideree)
    If choix_aerateur = aerateur_fine_bulle Or choix_aerateur = aerateur_moyenne_bulle Then
        O2_rdt_transfert_force = Feuil3.Range("BA_prolongee_O2_rdt_transfert").Cells(1, pi_FiliereConsideree)
        air_Q_Nm3j_force = Feuil3.Range("BA_prolongee_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
        diffuseur_encrassement_force = Feuil3.Range("BA_prolongee_diffuseur_encrassement").Cells(1, pi_FiliereConsideree)
        air_P_refoulement_force = Feuil3.Range("BA_prolongee_air_P_refoulement").Cells(1, pi_FiliereConsideree)
        surpresseur_conso_spec_force = Feuil3.Range("BA_prolongee_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
    Else
        ASB_eau_claire_force = Feuil3.Range("BA_prolongee_ASB_eau_claire").Cells(1, pi_FiliereConsideree)
    End If
    recirculation_taux_force = Feuil3.Range("BA_prolongee_recirculation_taux").Cells(1, pi_FiliereConsideree)
    recirculation_P_refoulement_force = Feuil3.Range("BA_prolongee_recirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
    recirculation_pompe_rdt_force = Feuil3.Range("BA_prolongee_recirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
    mn_NombreClarificateurs_force = Feuil3.Range("BA_prolongee_NombreClarificateurs").Cells(1, pi_FiliereConsideree)
    boues_indice_Mohlman_force = Feuil3.Range("BA_prolongee_boues_indice_Mohlman").Cells(1, pi_FiliereConsideree)
    sortie_MES_force = Feuil3.Range("BA_prolongee_sortie_MES").Cells(1, pi_FiliereConsideree)
    clarif_hauteur_force = Feuil3.Range("BA_prolongee_clarif_hauteur").Cells(1, pi_FiliereConsideree)
    clarif_vitesse_max_force = Feuil3.Range("BA_prolongee_clarif_vitesse_max").Cells(1, pi_FiliereConsideree)
    clarif_surface_force = Feuil3.Range("BA_prolongee_clarif_surface").Cells(1, pi_FiliereConsideree)
    boues_concentration_force = Feuil3.Range("BA_prolongee_boues_concentration").Cells(1, pi_FiliereConsideree)
    boues_MES_force = Feuil3.Range("BA_prolongee_boues_MES").Cells(1, pi_FiliereConsideree)
    boues_Q_force = False
    extraction_P_refoulement_force = Feuil3.Range("BA_prolongee_extraction_P_refoulement").Cells(1, pi_FiliereConsideree)
    extraction_pompe_rdt_force = Feuil3.Range("BA_prolongee_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)

    'lecture des valeurs forcées
    If nominal_G_force = True Then
        nominal_G_f = Feuil3.Range("BA_prolongee_nominal_age_boue").Cells(1, pi_FiliereConsideree)
    End If
    If nominal_MES_concentration_bassin_force = True Then
        nominal_MES_concentration_bassin_f = Feuil3.Range("BA_prolongee_nominal_MES_concentration_bassin").Cells(1, pi_FiliereConsideree)
    End If
    If nominal_MV_MES_force = True Then
        nominal_MV_MES_f = Feuil3.Range("BA_prolongee_nominal_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If volume_bassins_force = True Then
        volume_bassins_f = Feuil3.Range("BA_prolongee_volume_bassins").Cells(1, pi_FiliereConsideree)
    End If
    If choix_dephosphatation = dephosphatation_bio Or choix_dephosphatation = dephosphatation_co_precipitation Then
        If volume_anaerobie_force = True Then
            volume_anaerobie_f = Feuil3.Range("BA_prolongee_volume_anaerobie").Cells(1, pi_FiliereConsideree)
        End If
        If ratio_elec_anaerobie_force = True Then
            ratio_elec_anaerobie_f = Feuil3.Range("BA_prolongee_ratio_elec_anaerobie").Cells(1, pi_FiliereConsideree)
        End If
    End If
    If choix_config_bassin = config_bassin_plugflow Then
        If volume_anoxie_force = True Then
            volume_anoxie_f = Feuil3.Range("BA_prolongee_volume_anoxie").Cells(1, pi_FiliereConsideree)
        End If
        If ratio_elec_anoxie_force = True Then
            ratio_elec_anoxie_f = Feuil3.Range("BA_prolongee_ratio_elec_anoxie").Cells(1, pi_FiliereConsideree)
        End If
        If volume_aerobie_force = True Then
            volume_aerobie_f = Feuil3.Range("BA_prolongee_volume_aerobie").Cells(1, pi_FiliereConsideree)
        End If
        If ratio_elec_aerobie_force = True Then
            ratio_elec_aerobie_f = Feuil3.Range("BA_prolongee_ratio_elec_aerobie").Cells(1, pi_FiliereConsideree)
        End If
        If recirculation_MLSS_taux_force = True Then
            recirculation_MLSS_taux_f = Feuil3.Range("BA_prolongee_recirculation_MLSS_taux").Cells(1, pi_FiliereConsideree)
        End If
        If recirculation_MLSS_P_refoulement_force = True Then
            recirculation_MLSS_P_refoulement_f = Feuil3.Range("BA_prolongee_recirculation_MLSS_P_refoulement").Cells(1, pi_FiliereConsideree)
        End If
        If recirculation_MLSS_pompe_rdt_force = True Then
            recirculation_MLSS_pompe_rdt_f = Feuil3.Range("BA_prolongee_recirculation_MLSS_pompe_rdt").Cells(1, pi_FiliereConsideree)
        End If
    Else
        If volume_chenal_force = True Then
            volume_chenal_f = Feuil3.Range("BA_prolongee_volume_chenal").Cells(1, pi_FiliereConsideree)
        End If
        If ratio_elec_chenal_force = True Then
            ratio_elec_chenal_f = Feuil3.Range("BA_prolongee_ratio_elec_chenal").Cells(1, pi_FiliereConsideree)
        End If
    End If
    If reel_G_force = True Then
        reel_G_f = Feuil3.Range("BA_prolongee_reel_age_boue").Cells(1, pi_FiliereConsideree)
    End If
    If reel_MV_MES_force = True Then
        reel_MV_MES_f = Feuil3.Range("BA_prolongee_reel_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If reel_MES_concentration_bassin_force = True Then
        reel_MES_concentration_bassin_f = Feuil3.Range("BA_prolongee_reel_MES_concentration_bassin").Cells(1, pi_FiliereConsideree)
    End If
    If O2_dissous_force = True Then
        O2_dissous_f = Feuil3.Range("BA_prolongee_O2_dissous").Cells(1, pi_FiliereConsideree)
    End If
    If sortie_DBO_force = True Then
        sortie_DBO_f = Feuil3.Range("BA_prolongee_sortie_DBO").Cells(1, pi_FiliereConsideree)
    End If
    If sortie_NH4_force = True Then
        sortie_NH4_f = Feuil3.Range("BA_prolongee_sortie_NH4").Cells(1, pi_FiliereConsideree)
    End If
    If sortie_NO3_force = True Then
        sortie_NO3_f = Feuil3.Range("BA_prolongee_sortie_NO3").Cells(1, pi_FiliereConsideree)
    End If
    If rendement_bioP_force = True Then
        rendement_bioP_f = Feuil3.Range("BA_prolongee_rendement_bioP").Cells(1, pi_FiliereConsideree)
    End If
    If sortie_Pt_force = True Then
        sortie_Pt_f = Feuil3.Range("BA_prolongee_sortie_Pt").Cells(1, pi_FiliereConsideree)
    End If
    If methanol_pur_force = True Then
        methanol_pur_f = Feuil3.Range("BA_prolongee_methanol").Cells(1, pi_FiliereConsideree)
    End If
    If FeCl3_pur_force = True Then
        FeCl3_pur_f = Feuil3.Range("BA_prolongee_FeCl3").Cells(1, pi_FiliereConsideree)
    End If
    If hauteur_bassin_force = True Then
        hauteur_bassin_f = Feuil3.Range("BA_prolongee_hauteur_bassin").Cells(1, pi_FiliereConsideree)
    End If
    If O2_besoin_force = True Then
        O2_besoin_f = Feuil3.Range("BA_prolongee_O2_besoin").Cells(1, pi_FiliereConsideree)
    End If
    If O2_facteur_alfa_force = True Then
        O2_facteur_alfa_f = Feuil3.Range("BA_prolongee_O2_facteur_alpha").Cells(1, pi_FiliereConsideree)
    End If
    If choix_aerateur = aerateur_fine_bulle Or choix_aerateur = aerateur_moyenne_bulle Then
        If O2_rdt_transfert_force = True Then
            O2_rdt_transfert_f = Feuil3.Range("BA_prolongee_O2_rdt_transfert").Cells(1, pi_FiliereConsideree)
        End If
        If air_Q_Nm3j_force = True Then
            air_Q_Nm3j_f = Feuil3.Range("BA_prolongee_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
        End If
        If diffuseur_encrassement_force = True Then
            diffuseur_encrassement_f = Feuil3.Range("BA_prolongee_diffuseur_encrassement").Cells(1, pi_FiliereConsideree)
        End If
        If air_P_refoulement_force = True Then
            air_P_refoulement_f = Feuil3.Range("BA_prolongee_air_P_refoulement").Cells(1, pi_FiliereConsideree)
        End If
        If surpresseur_conso_spec_force = True Then
            surpresseur_conso_spec_f = Feuil3.Range("BA_prolongee_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
        End If
    Else
        If ASB_eau_claire_force = True Then
            ASB_eau_claire_f = Feuil3.Range("BA_prolongee_ASB_eau_claire").Cells(1, pi_FiliereConsideree)
        End If
    End If
    If recirculation_taux_force = True Then
        recirculation_taux_f = Feuil3.Range("BA_prolongee_recirculation_taux").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_P_refoulement_force = True Then
        recirculation_P_refoulement_f = Feuil3.Range("BA_prolongee_recirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_pompe_rdt_force = True Then
        recirculation_pompe_rdt_f = Feuil3.Range("BA_prolongee_recirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    If mn_NombreClarificateurs_force = True Then
        mi_NombreClarificateurs_f = Feuil3.Range("BA_prolongee_NombreClarificateurs").Cells(1, pi_FiliereConsideree)
    End If
    If boues_indice_Mohlman_force = True Then
        boues_indice_Mohlman_f = Feuil3.Range("BA_prolongee_boues_indice_Mohlman").Cells(1, pi_FiliereConsideree)
    End If
    If sortie_MES_force = True Then
        sortie_MES_f = Feuil3.Range("BA_prolongee_sortie_MES").Cells(1, pi_FiliereConsideree)
    End If
    If clarif_hauteur_force = True Then
        clarif_hauteur_f = Feuil3.Range("BA_prolongee_clarif_hauteur").Cells(1, pi_FiliereConsideree)
    End If
    If clarif_vitesse_max_force = True Then
        clarif_vitesse_max_f = Feuil3.Range("BA_prolongee_clarif_vitesse_max").Cells(1, pi_FiliereConsideree)
    End If
    If clarif_surface_force = True Then
        clarif_surface_f = Feuil3.Range("BA_prolongee_clarif_surface").Cells(1, pi_FiliereConsideree)
    End If
    If boues_concentration_force = True Then
        boues_concentration_f = Feuil3.Range("BA_prolongee_boues_concentration").Cells(1, pi_FiliereConsideree)
    End If
    If boues_MES_force = True Then
        boues_MES_f = Feuil3.Range("BA_prolongee_boues_MES").Cells(1, pi_FiliereConsideree)
    End If
    If boues_Q_force = True Then
        boues_MES_f = Feuil3.Range("BA_prolongee_boues_Q").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement_f = Feuil3.Range("BA_prolongee_extraction_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt_f = Feuil3.Range("BA_prolongee_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "3.2"

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


    'Température de référence pour la cinétique biologique (°C)
    T_reference = Feuil48.Range("hyp_II_BA_bio_equiv_Tref").Cells(1, 1)    ' 12

    'Age de boues à T_reference (j)
    If choix_dephosphatation = dephosphatation_bio Or choix_dephosphatation = dephosphatation_co_precipitation Then
        G_ratio_augmentation_P_bio = 0.08
    Else
        G_ratio_augmentation_P_bio = 0
    End If
    G_reference = 20 * (1 + G_ratio_augmentation_P_bio)

    'ratio MV/MES référence
    If choix_primaire = True Then
        MV_MES_reference = 0.77
    Else
        MV_MES_reference = 0.67
    End If

    'Facteur correctif température pour cinétique
    correctif_T = Feuil48.Range("hyp_II_BA_bio_equiv_correctionT").Cells(1, 1)    '1.07

    'coefficient permettant le calcul du volume anaérobie (temps de séjour)
    anaerobie_tps_passage_ref = 10

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
            a_rdt_DBO(I - 1) = Feuil48.Range("hyp_II_BA_rdtDBO_coef_ED").Cells(I, 1)
        Next I
    Else
        For I = 1 To 4
            a_rdt_DBO(I - 1) = Feuil48.Range("hyp_II_BA_rdtDBO_coef_EB").Cells(I, 1)
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
    If pd_TemperatureEauDimensionnement_°C >= 15 Then
        ratio_C_N_limite_nominal = 4
    ElseIf pd_TemperatureEauDimensionnement_°C <= 5 Then
        ratio_C_N_limite_nominal = 5
    Else
        ratio_C_N_limite_nominal = 4 + (5 - 4) * (pd_TemperatureEauDimensionnement_°C - 5) / (15 - 5)
    End If
    If pd_TemperatureEauConditionsReelles_°C >= 15 Then
        ratio_C_N_limite_reel = 4
    ElseIf pd_TemperatureEauDimensionnement_°C <= 5 Then
        ratio_C_N_limite_reel = 5
    Else
        ratio_C_N_limite_reel = 4 + (5 - 4) * (pd_TemperatureEauConditionsReelles_°C - 5) / (15 - 5)
    End If
    DCO_methanol = 1.5
    ratio_boues_methanol = 0.25

    'hypothèse de rendement sur la DCO
    rdt_DCO = 0.9

    'Charge radier maximale kg/(m2.h)
    For I = racleur_racle To racleur_Kruger
        charge_radier_max(I) = Feuil48.Range("hyp_II_BA_charge_radier_max").Cells(I, 1)
    Next I

    'facteur pour calcul de concentration limite des boues extraites et recirculées pour bonne décantation
    facteur_boues_C_lim = 1200

    'hypothèses pour le calcul des besoins en O2
    a0_O2_DBO_respiration = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_a0").Cells(1, 1)   ' 0.56
    a1_O2_DBO_respiration = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_a1").Cells(1, 1)  ' 0.15
    a2_O2_DBO_respiration = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_a2").Cells(1, 1)   '0.17
    correctif_T_ATV = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_correctionT").Cells(1, 1)   '1.072
    T_ref_ATV = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_Tref").Cells(1, 1)   '15

    'spécificité SHUNT
    If choix_fct_SHUNT = True Then
        ratio_O2_nit = 4.57 * 0.75
        ratio_O2_denit = -2.86 * 0.6
        ratio_C_N_limite_nominal = ratio_C_N_limite_nominal * 0.6
        ratio_C_N_limite_reel = ratio_C_N_limite_reel * 0.6
    Else
        ratio_O2_nit = 4.57
        ratio_O2_denit = -2.86
    End If

    'facteurs alpha0 (celui de fines bulles est calculé)
    O2_alpha0(aerateur_moyenne_bulle) = 0.8
    O2_alpha0(aerateur_brosse) = 0.8
    O2_alpha0(aerateur_turbine_lente) = 0.8
    O2_alpha0(aerateur_turbine_rapide) = 0.8

    'correction alpha en fonction des MES  alfa=alfa*(a0+a1*MES_reel)/(a0+a1*MES_lim)
    a0_correctif_alfa_MES = Feuil48.Range("hyp_II_BA_alpha_correction_MES_a0").Cells(1, 1)    ' 1.0007
    a1_correctif_alfa_MES = Feuil48.Range("hyp_II_BA_alpha_correction_MES_a1").Cells(1, 1)    ' -0.035
    MES_reference_alfa = Feuil48.Range("hyp_II_BA_alpha_correction_MES_Cref").Cells(1, 1)    ' 5

    'relation calcul alpha fines bulles    ao+a_MVS*MVS+a_G*G_eq
    a0_alpha = Feuil48.Range("hyp_II_BA_alpha_finebulle_a0").Cells(1, 1)    ' 0.51
    a_alpha_MVS = Feuil48.Range("hyp_II_BA_alpha_finebulle_aMVS").Cells(1, 1)    '-0.062
    a_alpha_G = Feuil48.Range("hyp_II_BA_alpha_finebulle_aGeq").Cells(1, 1)    ' 0.019
    O2_facteur_alfa_maxi = 1
    O2_facteur_alfa_mini = 0.05
    
    'facteur beta
    O2_facteur_beta = Feuil48.Range("hyp_calcul_facteurK_facteur_Beta").Cells(1, 1)    ' 0.95

    'calcul de la pression atmosphérique  Patm=P0*((a0-a1*altitude)^a2)/100
    P0 = Feuil48.Range("hyp_calcul_Patm_coef").Cells(1, 1)   ' 101325
    a0_Patm = Feuil48.Range("hyp_calcul_Patm_coef").Cells(2, 1)    '1
    a1_Patm = Feuil48.Range("hyp_calcul_Patm_coef").Cells(3, 1)    '-0.0000225577
    a2_Patm = Feuil48.Range("hyp_calcul_Patm_coef").Cells(4, 1)    '5.25588

    'calcul de la surpression en diffusion de bulles  delta_P=a0/a1*P0/100*(hauteur_bassin-hauteur_diffuseur)
    a0_delta_P = Feuil48.Range("hyp_insufflation_deltaP_a0").Cells(1, 1)  '0.35
    a1_delta_P = Feuil48.Range("hyp_insufflation_deltaP_a1").Cells(1, 1)   '10.33
    hauteur_diffuseur = Feuil48.Range("hyp_insufflation_bulles_hauteur_diffuseur").Cells(1, 1)   '0.25 'm

    'calcul O2 saturation  O2_saturation(T,P0)= somme(ai*Teau^i)
    For I = 1 To 5
        a_O2_saturation(I - 1) = Feuil48.Range("hyp_calcul_O2sat_T_P0_coef_a").Cells(I, 1)
    Next I

    'correction température pour facteur K
    correctif_T_K = Feuil48.Range("hyp_transfertO2_facteurK_correction_T").Cells(1, 1)    ' 1.024
    T_ref_K = Feuil48.Range("hyp_transfertO2_facteurK_Tref").Cells(1, 1)  '20

    'O2 présent dans l'air (kgO2/Nm3air)
    ratio_kgO2_Nm3air = Feuil48.Range("hyp_ratio_kgO2_Nm3air").Cells(1, 1)  ' 0.3

    'rendement de dissolution en eau claire (%/m)
    If hauteur_bassin_force = True And hauteur_bassin_f < Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(4, 1) Then
        O2_rdt_dissolution(aerateur_fine_bulle) = Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(5, 1) * hauteur_bassin_f + Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(6, 1)
    Else
        O2_rdt_dissolution(aerateur_fine_bulle) = Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(1, 1)
    End If
    O2_rdt_dissolution(aerateur_moyenne_bulle) = Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(2, 1)

    'Apport spécifique brut (eau claire)  (kWh/kO2_transféré)   source:fndae 31
    ASB(aerateur_brosse) = Feuil48.Range("hyp_II_BA_aerateur_surface_ASB").Cells(1, 1)   '1.7
    ASB(aerateur_turbine_lente) = Feuil48.Range("hyp_II_BA_aerateur_surface_ASB").Cells(2, 1)   '1.65
    ASB(aerateur_turbine_rapide) = Feuil48.Range("hyp_II_BA_aerateur_surface_ASB").Cells(3, 1)   '1.2

    'consommations spécifiques par type de surpresseurs      en Wh/(Nm3.mCE)
    For I = 1 To nb_surpresseur_gamme_Q
        For j = 1 To nb_surpresseur_gamme_P
            surpresseur_conso_Wh_Nm3mCE(surpresseur_roots, I, j) = 4.5
            surpresseur_conso_Wh_Nm3mCE(surpresseur_vis, I, j) = 3.8
            surpresseur_conso_Wh_Nm3mCE(surpresseur_turbo, I, j) = 3.5
        Next j
    Next I

    'impact du type de régulation
'    impact_regulation(regulation_horloge) = 0.4
'    impact_regulation(regulation_sans_variation) = 0.25
'    impact_regulation(regulation_avec_variation) = 0.1
'    impact_regulation(regulation_avance) = 0
    If pd_TemperatureEauConditionsReelles_°C < 14 Then
        md_ConcentrationOxygeneDissous_mgL(regulation_horloge) = 4
        md_ConcentrationOxygeneDissous_mgL(regulation_sans_variation) = 2.5
        md_ConcentrationOxygeneDissous_mgL(regulation_avec_variation) = 2
        md_ConcentrationOxygeneDissous_mgL(regulation_avance) = 1.5
    Else
        md_ConcentrationOxygeneDissous_mgL(regulation_horloge) = 4
        md_ConcentrationOxygeneDissous_mgL(regulation_sans_variation) = 2
        md_ConcentrationOxygeneDissous_mgL(regulation_avec_variation) = 1.5
        md_ConcentrationOxygeneDissous_mgL(regulation_avance) = 0.8
    End If

    'clarificateur
    rate_puissance_fixe = 0.1    'Wh/m3 par rapport au nominal
    rate_puissance_debit = 1    'Wh/m3 par rapport au débit réel

    agitation_fct = 24

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "3.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub attribution_valeur_par_defaut()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On retrouve ici l'ensemble des valeurs pour lesquelles on attribue une valeur guide par défaut
    Dim volume_corrige As Double
    Dim G_equivalent_nominal_G As Double
    Dim rdt_pompe As Double
    Dim rdt_moteur_pompe As Double


    'Age de boues design (j)
    If nominal_G_force = True Then
        nominal_G = nominal_G_f
    Else
        nominal_G = G_reference * correctif_T ^ (T_reference - pd_TemperatureEauDimensionnement_°C)
    End If

    'Concentration MES design (dans les bassins) g/L
    If nominal_MES_concentration_bassin_force = True Then
        nominal_MES_concentration_bassin = nominal_MES_concentration_bassin_f
    Else
        If choix_primaire = True Then
            nominal_MES_concentration_bassin = 3.5
        Else
            nominal_MES_concentration_bassin = 5
        End If
    End If

    'agitation des cuves (W/m3)
    If ratio_elec_anoxie_force = True Then
        ratio_elec_anoxie = ratio_elec_anoxie_f
    Else
        ratio_elec_anoxie = 3
    End If
    If ratio_elec_aerobie_force = True Then
        ratio_elec_aerobie = ratio_elec_aerobie_f
    Else
        ratio_elec_aerobie = 0
    End If
    If ratio_elec_chenal_force = True Then
        ratio_elec_chenal = ratio_elec_chenal_f
    Else
        ratio_elec_chenal = 3
    End If
    If ratio_elec_anaerobie_force = True Then
        ratio_elec_anaerobie = ratio_elec_anaerobie_f
    Else
        ratio_elec_anaerobie = 10
    End If

    'Age de boues réel (j)
    If reel_G_force = True Then
        reel_G = reel_G_f
    Else
        G_equivalent_nominal_G = nominal_G / correctif_T ^ (T_reference - pd_TemperatureEauDimensionnement_°C)
        reel_G = G_equivalent_nominal_G * correctif_T ^ (T_reference - pd_TemperatureEauConditionsReelles_°C)
    End If

    'Concentration moyenne en O2 dissous (mg/L)
    If O2_dissous_force = True Then
        O2_dissous = O2_dissous_f
    Else
        If choix_fct_SHUNT = True Then
            O2_dissous = 0.5   'spécificité SHUNT / O2 dissous très faible nécessaire
        Else
            O2_dissous = md_ConcentrationOxygeneDissous_mgL(choix_regulation)
        End If
    End If

    'Concentration de sortie en NH4  (mg/L)
    If sortie_NH4_force = True Then
        sortie_NH4 = sortie_NH4_f
    Else
        sortie_NH4 = 1
    End If

    'Concentration de sortie en NO3  (mg/L)
    If sortie_NO3_force = True Then
        sortie_NO3 = sortie_NO3_f
    Else
        sortie_NO3 = 6
    End If

    'Concentration de sortie en Pt (mg/L)
    If sortie_Pt_force = True Then
        sortie_Pt = sortie_Pt_f
    Else
        sortie_Pt = Pt_garantie
    End If

    'Hauteur moyenne des bassins  (m)
    If hauteur_bassin_force = True Then
        hauteur_bassin = hauteur_bassin_f
    Else
        If choix_aerateur = aerateur_fine_bulle Or choix_aerateur = aerateur_moyenne_bulle Then
            hauteur_bassin = 6
        Else
            hauteur_bassin = 3
        End If
    End If

    'Durée depuis le dernier nettoyage des diffuseurs  (ans)
    If diffuseur_encrassement_force = True Then
        diffuseur_encrassement = diffuseur_encrassement_f
    Else
        diffuseur_encrassement = 0
    End If

    'Pression de refoulement des surpresseurs (mCE)
    If air_P_refoulement_force = True Then
        air_P_refoulement = air_P_refoulement_f
    Else
        air_P_refoulement = hauteur_bassin + 2 + 0.25 * diffuseur_encrassement
    End If

    'ASB
    If ASB_eau_claire_force = True Then
        ASB_eau_claire = ASB_eau_claire_f
    Else
        ASB_eau_claire = ASB(choix_aerateur)
    End If

    'Taux de recirculation de liqueur mixte
    'If recirculation_MLSS_taux_force = True Then
    'recirculation_MLSS_taux = recirculation_MLSS_taux_f
    'Else
    'recirculation_MLSS_taux = 1.5
    'End If

    'Pression des pompes de recirculation de liqueur mixte (mCE)
    If recirculation_MLSS_P_refoulement_force = True Then
        recirculation_MLSS_P_refoulement = recirculation_MLSS_P_refoulement_f
    Else
        recirculation_MLSS_P_refoulement = 2
    End If

    'recirculation de liqueur mixte
    If recirculation_MLSS_pompe_rdt_force = True Then
        recirculation_MLSS_pompe_rdt = recirculation_MLSS_pompe_rdt_f
    Else
        rdt_pompe = 0.7
        rdt_moteur_pompe = 0.88
        recirculation_MLSS_pompe_rdt = rdt_pompe * rdt_moteur_pompe
    End If
    ratio_elec_recirculation_MLSS = pd_ACCELERATION_PESANTEUR_m_s2 / (pd_NOMBRE_SECONDE_PAR_HEURE * recirculation_MLSS_pompe_rdt)    ' kWh/(m3*mCE)


    'Taux de recirculation des boues du clarif
    If recirculation_taux_force = True Then
        recirculation_taux = recirculation_taux_f
    Else
        recirculation_taux = 1
    End If

    'Pression des pompes de recirculation des boues du clarif (mCE)
    If recirculation_P_refoulement_force = True Then
        recirculation_P_refoulement = recirculation_P_refoulement_f
    Else
        recirculation_P_refoulement = 5
    End If

    'recirculation des boues du clarif
    If recirculation_pompe_rdt_force = True Then
        recirculation_pompe_rdt = recirculation_pompe_rdt_f
    Else
        rdt_pompe = 0.7
        rdt_moteur_pompe = 0.88
        recirculation_pompe_rdt = rdt_pompe * rdt_moteur_pompe
    End If
    ratio_elec_recirculation = pd_ACCELERATION_PESANTEUR_m_s2 / (pd_NOMBRE_SECONDE_PAR_HEURE * recirculation_pompe_rdt)    ' kWh/(m3*mCE)

    'Nombre de clarificateurs
    If mn_NombreClarificateurs_force = True Then
        mi_NombreClarificateurs = mi_NombreClarificateurs_f
    Else
        mi_NombreClarificateurs = 1
    End If
    
    'Indice de Mohlman (mL/g)
    If boues_indice_Mohlman_force = True Then
        boues_indice_Mohlman = boues_indice_Mohlman_f
    Else
        If choix_primaire = True Then
            If choix_dephosphatation = dephosphatation_non Then
                boues_indice_Mohlman = 165
            Else
                boues_indice_Mohlman = 185
            End If
        Else
            If choix_dephosphatation = dephosphatation_non Then
                boues_indice_Mohlman = 125
            Else
                boues_indice_Mohlman = 135
            End If
        End If
    End If

    'Sortie MES eau traitée (mg/L)
    If sortie_MES_force = True Then
        sortie_MES = sortie_MES_f
    Else
        sortie_MES = 20
    End If

    'Hauteur du clarificateur (m)
    If clarif_hauteur_force = True Then
        clarif_hauteur = clarif_hauteur_f
    Else
        clarif_hauteur = 4
    End If

    'Vitesse hydraulique maximum (m/h)
    If clarif_vitesse_max_force = True Then
        clarif_vitesse_max = clarif_vitesse_max_f
    Else
        volume_corrige = boues_indice_Mohlman * nominal_MES_concentration_bassin
        clarif_vitesse_max = (100 * clarif_hauteur * (sortie_MES / 3.15) ^ 0.5) / ((1 + recirculation_taux) * volume_corrige)
    End If

    'Pression de refoulement de l'extraction (mCE)
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement = extraction_P_refoulement_f
    Else
        extraction_P_refoulement = 5
    End If

    'extraction des boues du clarif
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt = extraction_pompe_rdt_f
    Else
        rdt_pompe = 0.7
        rdt_moteur_pompe = 0.88
        extraction_pompe_rdt = rdt_pompe * rdt_moteur_pompe
    End If
    ratio_elec_extraction = pd_ACCELERATION_PESANTEUR_m_s2 / (pd_NOMBRE_SECONDE_PAR_HEURE * extraction_pompe_rdt)    ' kWh/(m3*mCE)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "3.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub ecriture_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil3.Range("BA_prolongee_nominal_age_boue").Cells(1, Colonne) = nominal_G
    Feuil3.Range("BA_prolongee_nominal_MES_concentration_bassin").Cells(1, Colonne) = nominal_MES_concentration_bassin
    Feuil3.Range("BA_prolongee_nominal_MV_MES").Cells(1, Colonne) = nominal_MV_MES
    Feuil3.Range("BA_prolongee_volume_bassins").Cells(1, Colonne) = volume_bassins
    If choix_dephosphatation = dephosphatation_bio Or choix_dephosphatation = dephosphatation_co_precipitation Then
        Feuil3.Range("BA_prolongee_volume_anaerobie").Cells(1, Colonne) = volume_anaerobie
        Feuil3.Range("BA_prolongee_ratio_elec_anaerobie").Cells(1, Colonne) = ratio_elec_anaerobie
        Feuil3.Range("BA_prolongee_rendement_bioP").Cells(1, Colonne) = rendement_bioP
    End If
    If choix_config_bassin = config_bassin_plugflow Then
        Feuil3.Range("BA_prolongee_volume_anoxie").Cells(1, Colonne) = volume_anoxie
        Feuil3.Range("BA_prolongee_ratio_elec_anoxie").Cells(1, Colonne) = ratio_elec_anoxie
        Feuil3.Range("BA_prolongee_volume_aerobie").Cells(1, Colonne) = volume_aerobie
        Feuil3.Range("BA_prolongee_ratio_elec_aerobie").Cells(1, Colonne) = ratio_elec_aerobie
        Feuil3.Range("BA_prolongee_recirculation_MLSS_taux").Cells(1, Colonne) = recirculation_MLSS_taux
        Feuil3.Range("BA_prolongee_recirculation_MLSS_P_refoulement").Cells(1, Colonne) = recirculation_MLSS_P_refoulement
        Feuil3.Range("BA_prolongee_recirculation_MLSS_pompe_rdt").Cells(1, Colonne) = recirculation_MLSS_pompe_rdt
    Else
        Feuil3.Range("BA_prolongee_volume_chenal").Cells(1, Colonne) = volume_chenal
        Feuil3.Range("BA_prolongee_ratio_elec_chenal").Cells(1, Colonne) = ratio_elec_chenal
    End If
    Feuil3.Range("BA_prolongee_reel_age_boue").Cells(1, Colonne) = reel_G
    Feuil3.Range("BA_prolongee_reel_MV_MES").Cells(1, Colonne) = reel_MV_MES
    Feuil3.Range("BA_prolongee_reel_MES_concentration_bassin").Cells(1, Colonne) = reel_MES_concentration_bassin
    Feuil3.Range("BA_prolongee_O2_dissous").Cells(1, Colonne) = O2_dissous
    Feuil3.Range("BA_prolongee_sortie_DBO").Cells(1, Colonne) = sortie_DBO
    Feuil3.Range("BA_prolongee_sortie_NH4").Cells(1, Colonne) = sortie_NH4
    Feuil3.Range("BA_prolongee_sortie_NO3").Cells(1, Colonne) = sortie_NO3
    Feuil3.Range("BA_prolongee_methanol").Cells(1, Colonne) = methanol_pur
    If choix_dephosphatation <> dephosphatation_non Then
        Feuil3.Range("BA_prolongee_sortie_Pt").Cells(1, Colonne) = sortie_Pt
    End If
    If choix_dephosphatation = dephosphatation_precipitation Or choix_dephosphatation = dephosphatation_co_precipitation Then
        Feuil3.Range("BA_prolongee_FeCl3").Cells(1, Colonne) = FeCl3_pur
    End If
    Feuil3.Range("BA_prolongee_hauteur_bassin").Cells(1, Colonne) = hauteur_bassin
    Feuil3.Range("BA_prolongee_O2_besoin").Cells(1, Colonne) = O2_besoin
    Feuil3.Range("BA_prolongee_O2_facteur_alpha").Cells(1, Colonne) = O2_facteur_alfa
    If choix_aerateur = aerateur_fine_bulle Or choix_aerateur = aerateur_moyenne_bulle Then
        Feuil3.Range("BA_prolongee_O2_rdt_transfert").Cells(1, Colonne) = O2_rdt_transfert
        Feuil3.Range("BA_prolongee_air_Q_Nm3j").Cells(1, Colonne) = air_Q_Nm3j
        Feuil3.Range("BA_prolongee_diffuseur_encrassement").Cells(1, Colonne) = diffuseur_encrassement
        Feuil3.Range("BA_prolongee_air_P_refoulement").Cells(1, Colonne) = air_P_refoulement
        Feuil3.Range("BA_prolongee_surpresseur_conso_spec").Cells(1, Colonne) = surpresseur_conso_spec
    Else
        Feuil3.Range("BA_prolongee_ASB_eau_claire").Cells(1, Colonne) = ASB_eau_claire
    End If
    Feuil3.Range("BA_prolongee_recirculation_taux").Cells(1, Colonne) = recirculation_taux
    Feuil3.Range("BA_prolongee_recirculation_P_refoulement").Cells(1, Colonne) = recirculation_P_refoulement
    Feuil3.Range("BA_prolongee_recirculation_pompe_rdt").Cells(1, Colonne) = recirculation_pompe_rdt
    Feuil3.Range("BA_prolongee_NombreClarificateurs").Cells(1, Colonne) = mi_NombreClarificateurs
    Feuil3.Range("BA_prolongee_boues_indice_Mohlman").Cells(1, Colonne) = boues_indice_Mohlman
    Feuil3.Range("BA_prolongee_sortie_MES").Cells(1, Colonne) = sortie_MES
    Feuil3.Range("BA_prolongee_clarif_hauteur").Cells(1, Colonne) = clarif_hauteur
    Feuil3.Range("BA_prolongee_clarif_vitesse_max").Cells(1, Colonne) = clarif_vitesse_max
    Feuil3.Range("BA_prolongee_clarif_surface").Cells(1, Colonne) = clarif_surface
    Feuil3.Range("BA_prolongee_clarif_vitesse_max_recalculee").Cells(1, Colonne) = clarif_vitesse_max_recalc
    Feuil3.Range("BA_prolongee_clarif_charge_radier_nominal").Cells(1, Colonne) = clarif_charge_radier_nominal
    Feuil3.Range("BA_prolongee_boues_concentration").Cells(1, Colonne) = boues_concentration
    Feuil3.Range("BA_prolongee_boues_MES").Cells(1, Colonne) = boues_MES
    Feuil3.Range("BA_prolongee_boues_Q").Cells(1, Colonne) = boues_Q
    Feuil3.Range("BA_prolongee_extraction_P_refoulement").Cells(1, Colonne) = extraction_P_refoulement
    Feuil3.Range("BA_prolongee_extraction_pompe_rdt").Cells(1, Colonne) = extraction_pompe_rdt
    Feuil3.Range("BA_prolongee_electricite").Cells(1, Colonne) = electricite
    Feuil3.Range("BA_prolongee_electricite_aeration").Cells(1, Colonne) = electricite_aeration

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "3.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub effacer_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On efface les valeurs sur Excel
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil3.Range("BA_prolongee_nominal_age_boue").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_nominal_MES_concentration_bassin").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_nominal_MV_MES").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_volume_bassins").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_volume_anaerobie").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_volume_anoxie").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_volume_aerobie").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_volume_chenal").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_ratio_elec_anaerobie").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_ratio_elec_anoxie").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_ratio_elec_aerobie").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_ratio_elec_chenal").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_recirculation_MLSS_taux").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_recirculation_MLSS_P_refoulement").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_recirculation_MLSS_pompe_rdt").Cells(1, Colonne) = ""     'recirculation_MLSS_pompe_rdt
    Feuil3.Range("BA_prolongee_reel_age_boue").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_reel_MV_MES").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_reel_MES_concentration_bassin").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_O2_dissous").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_sortie_DBO").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_sortie_NH4").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_sortie_NO3").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_rendement_bioP").Cells(1, Colonne) = ""      'rendement_bioP
    Feuil3.Range("BA_prolongee_sortie_Pt").Cells(1, Colonne) = ""      'sortie_Pt
    Feuil3.Range("BA_prolongee_methanol").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_FeCl3").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_hauteur_bassin").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_O2_besoin").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_O2_facteur_alpha").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_O2_rdt_transfert").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_air_Q_Nm3j").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_diffuseur_encrassement").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_air_P_refoulement").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_surpresseur_conso_spec").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_ASB_eau_claire").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_recirculation_taux").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_recirculation_P_refoulement").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_recirculation_pompe_rdt").Cells(1, Colonne) = ""     'recirculation_pompe_rdt
    Feuil3.Range("BA_prolongee_NombreClarificateurs").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_boues_indice_Mohlman").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_sortie_MES").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_clarif_hauteur").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_clarif_vitesse_max").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_clarif_surface").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_clarif_vitesse_max_recalculee").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_clarif_charge_radier_nominal").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_boues_concentration").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_boues_MES").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_boues_Q").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_extraction_P_refoulement").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_extraction_pompe_rdt").Cells(1, Colonne) = ""     'extraction_pompe_rdt
    Feuil3.Range("BA_prolongee_electricite").Cells(1, Colonne) = ""
    Feuil3.Range("BA_prolongee_electricite_aeration").Cells(1, Colonne) = ""           'electricite_aeration

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "3.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 5  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil3.Range("BA_prolongee_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf
    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "3.8"

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











