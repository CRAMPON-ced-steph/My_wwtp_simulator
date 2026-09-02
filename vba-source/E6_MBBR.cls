Option Explicit

'Fonctionnement en mode Mox
Dim choix_fct_Mox As Boolean

Private Const critere_convergence = 0.0000000001

Private Const critere_volume_nul = 0.00001

'constantes pour les cuves
Private Const nb_cuves_max = 10
Private Const cuve_predenit1 = 1
Private Const cuve_predenit2 = 2
Private Const cuve_C1 = 3
Private Const cuve_C2 = 4
Private Const cuve_N1 = 5
Private Const cuve_N2 = 6
Private Const cuve_deox = 7
Private Const cuve_postdenit1 = 8
Private Const cuve_postdenit2 = 9
Private Const cuve_reox = 10

'constantes pour les erreurs (les 10 premières correspondent aux HRT des cuves)
Private Const nb_erreurs = 21
'Private Const cuve_predenit1 = 1
'Private Const cuve_predenit2 = 2
'Private Const cuve_C1 = 3
'Private Const cuve_C2 = 5
'Private Const cuve_N1 = 5
'Private Const cuve_N2 = 6
'Private Const cuve_deox = 7
'Private Const cuve_postdenit1 = 8
'Private Const cuve_postdenit2 = 9
'Private Const cuve_reox = 10
Private Const erreur_predenit_2cuves = 11
Dim erreur_Qair_augmente_agitation(nb_cuves_max) As Integer     'de 12 à 21

Private Const cuve_oui = 1
Private Const cuve_non = 0

'constantes pour le type de source de carbone
Private Const methanol = 1
Private Const ethanol = 2

'constantes pour le type de média
Private Const nb_type_media = 6
Private Const media_K1 = 1
Private Const media_K1_heavy = 2
Private Const media_K3 = 3
Private Const media_K5 = 4
Private Const media_chip_P = 5
Private Const media_chip_M = 6

'constantes pour les qualités du flux entre chaque étape
Private Const nb_flux = 13
Private Const flux_inlet = 0
Private Const flux_inlet_recirculation = 1
Private Const flux_predenit1_predenit2 = 2
Private Const flux_predenit2_C1 = 3
Private Const flux_C1_C2 = 4
Private Const flux_C2_N1 = 5
Private Const flux_N1_N2 = 6
Private Const flux_N2_deox = 7
Private Const flux_deox_separationMLSS = 8
Private Const flux_recirculation_MLSS = 9
Private Const flux_separationMLSS_postdenit1 = 10
Private Const flux_postdenit1_postdenit2 = 11
Private Const flux_postdenit2_reox = 12
Private Const flux_outlet = 13
Private Const nb_parametres_flux = 9
Private Const repere_Q = 1
Private Const repere_DCO = 2
Private Const repere_DBO = 3
Private Const repere_MES = 4
Private Const repere_NK = 5
Private Const repere_NH4 = 6
Private Const repere_NO3 = 7
Private Const repere_Pt = 8
Private Const repere_O2 = 9

'types d'aérateurs
Private Const aerateur_fine_bulle = 1
Private Const aerateur_moyenne_bulle = 2

'Agitation des cuves aérées
Private Const cuve_aeree_agitation_air = 1
Private Const cuve_aeree_agitation_mecanique = 2

'constantes pour le choix du type de régulation d'air
Private Const nb_type_regulation = 4  'nombre de types de régulation
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

'constantes sur le type de séparation
Private Const clarif = 1
Private Const clarif_polymere = 2
Private Const clarif_polymere_coag = 3

'variables pour les choix de l'utilisateur
Dim choix_type_separation As Integer
Dim choix_cuve(nb_cuves_max) As Integer
Dim choix_media(nb_cuves_max) As Integer
Dim choix_cuve_aerateur(nb_cuves_max) As Integer
Dim choix_agitation_cuve_aeree(nb_cuves_max) As Integer
Dim choix_source_C As Integer
Dim choix_surpresseur As Integer
Dim choix_regulation As Integer

'variables échangées avec l'utilisateur
Dim nominal_outlet_DBO_mgL As Double
Dim nominal_outlet_NH4_mgL As Double
Dim nominal_outlet_NO3_mgL As Double
Dim Norga_dure_soluble As Double
Dim nominal_predenit_NO3_mgL As Double
Dim nominal_last_aerated_tank_NO3_mgL As Double
Dim nominal_recirculation_MLSS_taux As Double
Dim ratio_elimination_predenit1_predenit12 As Double
Dim ratio_elimination_C1_C12 As Double
Dim ratio_elimination_N1_N12 As Double
Dim ratio_elimination_postdenit1_postdenit12 As Double
Dim cuve_O2_dissous(nb_cuves_max) As Double
Dim cuve_filling(nb_cuves_max) As Double
Dim cuve_volume(nb_cuves_max) As Double
Dim cuve_ratio_agitation_W_m3(nb_cuves_max) As Double
Dim cuve_O2_facteur_alfa(nb_cuves_max) As Double
Dim cuve_O2_rdt_transfert(nb_cuves_max) As Double
Dim cuve_hauteur(nb_cuves_max) As Double
Dim cuve_aeree_Vair_agitation_Nmh(nb_cuves_max) As Double
Dim cuve_aeree_brassage_air_hj(nb_cuves_max) As Double
Public postdenit_carbone_apporte_flux As Double
Dim volume_total_bassins As Double
Dim reel_recirculation_MLSS_taux As Double
Dim reel_outlet_DBO_mgL As Double
Dim reel_outlet_NH4_mgL As Double
Dim reel_outlet_NO3_mgL As Double
Dim O2_besoin_total As Double
Dim air_Q_Nm3j As Double
Dim suraeration_brassage_Nm3j As Double
Dim diffuseur_encrassement As Double
Dim air_P_refoulement_moyenne As Double
Dim surpresseur_conso_spec As Double
Dim recirculation_MLSS_P_refoulement As Double
Dim recirculation_MLSS_pompe_rdt As Double
'Dim boues_indice_Mohlman As Double
Dim sortie_MES As Double   '***************** pour l'instant, n'est pas calculée ni affiché donc à voir *************************
Dim boues_concentration As Double
Public boues_MES As Double
Public boues_MV_MES As Double
Public boues_Q As Double
Dim extraction_P_refoulement As Double
Dim extraction_pompe_rdt As Double
Public electricite As Double
Dim electricite_aeration As Double

Public electricite_fixe As Double

'variables indiquant les paramètres forcés
Dim nominal_outlet_DBO_mgL_force As Boolean
Dim nominal_outlet_NH4_mgL_force As Boolean
Dim nominal_outlet_NO3_mgL_force As Boolean
Dim Norga_dure_soluble_force As Boolean
Dim nominal_predenit_NO3_mgL_force As Boolean
Dim nominal_last_aerated_tank_NO3_mgL_force As Boolean
Dim nominal_recirculation_MLSS_taux_force As Boolean
Dim ratio_elimination_predenit1_predenit12_force As Boolean
Dim ratio_elimination_C1_C12_force As Boolean
Dim ratio_elimination_N1_N12_force As Boolean
Dim ratio_elimination_postdenit1_postdenit12_force As Boolean
Dim cuve_O2_dissous_force(nb_cuves_max) As Boolean
Dim cuve_filling_force(nb_cuves_max) As Boolean
Dim cuve_volume_force(nb_cuves_max) As Boolean
Dim cuve_ratio_agitation_W_m3_force(nb_cuves_max) As Boolean
Dim cuve_O2_facteur_alfa_force(nb_cuves_max) As Boolean
Dim cuve_O2_rdt_transfert_force(nb_cuves_max) As Boolean
Dim cuve_hauteur_force(nb_cuves_max) As Boolean
Dim cuve_aeree_Vair_agitation_Nmh_force(nb_cuves_max) As Boolean
Dim cuve_aeree_brassage_air_hj_force(nb_cuves_max) As Boolean
Dim postdenit_carbone_apporte_flux_force As Boolean
Dim reel_recirculation_MLSS_taux_force As Boolean
Dim reel_outlet_DBO_mgL_force As Boolean
Dim reel_outlet_NH4_mgL_force As Boolean
Dim reel_outlet_NO3_mgL_force As Boolean
Dim O2_besoin_total_force As Boolean
Dim air_Q_Nm3j_force As Boolean
Dim diffuseur_encrassement_force As Boolean
Dim air_P_refoulement_moyenne_force As Boolean
Dim surpresseur_conso_spec_force As Boolean
Dim recirculation_MLSS_P_refoulement_force As Boolean
Dim recirculation_MLSS_pompe_rdt_force As Boolean
'Dim boues_indice_Mohlman_force As Boolean
Dim sortie_MES_force As Boolean
Dim boues_concentration_force As Boolean
Public boues_MES_force As Boolean
Public boues_MV_MES_force As Boolean
Dim extraction_P_refoulement_force As Boolean
Dim extraction_pompe_rdt_force As Boolean

'variables contenant les valeurs des paramètres forcés
Dim nominal_outlet_DBO_mgL_f As Double
Dim nominal_outlet_NH4_mgL_f As Double
Dim nominal_outlet_NO3_mgL_f As Double
Dim Norga_dure_soluble_f As Double
Dim nominal_predenit_NO3_mgL_f As Double
Dim nominal_last_aerated_tank_NO3_mgL_f As Double
Dim nominal_recirculation_MLSS_taux_f As Double
Dim ratio_elimination_predenit1_predenit12_f As Double
Dim ratio_elimination_C1_C12_f As Double
Dim ratio_elimination_N1_N12_f As Double
Dim ratio_elimination_postdenit1_postdenit12_f As Double
Dim cuve_O2_dissous_f(nb_cuves_max) As Double
Dim cuve_filling_f(nb_cuves_max) As Double
Dim cuve_volume_f(nb_cuves_max) As Double
Dim cuve_ratio_agitation_W_m3_f(nb_cuves_max) As Double
Dim cuve_O2_facteur_alfa_f(nb_cuves_max) As Double
Dim cuve_O2_rdt_transfert_f(nb_cuves_max) As Double
Dim cuve_hauteur_f(nb_cuves_max) As Double
Dim cuve_aeree_Vair_agitation_Nmh_f(nb_cuves_max) As Double
Dim cuve_aeree_brassage_air_hj_f(nb_cuves_max) As Double
Dim postdenit_carbone_apporte_flux_f As Double
Dim reel_recirculation_MLSS_taux_f As Double
Dim reel_outlet_DBO_mgL_f As Double
Dim reel_outlet_NH4_mgL_f As Double
Dim reel_outlet_NO3_mgL_f As Double
Dim O2_besoin_total_f As Double
Dim air_Q_Nm3j_f As Double
Dim diffuseur_encrassement_f As Double
Dim air_P_refoulement_moyenne_f As Double
Dim surpresseur_conso_spec_f As Double
Dim recirculation_MLSS_P_refoulement_f As Double
Dim recirculation_MLSS_pompe_rdt_f As Double
'Dim boues_indice_Mohlman_f As Double
Dim sortie_MES_f As Double
Dim boues_concentration_f As Double
Public boues_MES_f As Double
Public boues_MV_MES_f As Double
Dim extraction_P_refoulement_f As Double
Dim extraction_pompe_rdt_f As Double

'hypothèses
Dim media_filling_max(nb_cuves_max, nb_type_media) As Double
Dim media_surface_m2m3(nb_type_media) As Double
Dim a_ratio_EB_MV_MES(2) As Double
Dim ratio_EB_MV_MES_min As Double
Dim ratio_EB_MV_MES_max As Double
Dim ratio_ED_MV_MES As Double
Dim ratio_N_assimile As Double
Dim ratio_P_assimile As Double
Dim ratio_N_assimile_Capporte As Double
Dim ratio_O2_nit As Double
Dim rate_DBO_apportee_denit_nominal As Double
Dim rate_DBO_apportee_denit_reel As Double
Dim rate_DBO_methanol As Double
Dim rate_DBO_ethanol As Double
Dim ratio_boues_DBO_apportee As Double
Dim ratio_O2_denit As Double
Dim ratio_NO3eq_O2 As Double
Dim ratio_O2_DBO As Double
Dim EB_rate_DBOsol_DBO As Double
Dim EB_rate_Norgasol_Norga As Double
Dim rate_Norgapart_hydrolyse As Double
Dim rate_DBOpart_hydrolyse As Double
Dim k_nit As Double   'constante de réaction pour la nit qui dépend du traitement primaire et de la présence d'une prédénit
Dim hyp_rate_DBOin_appliquee_10°C_CN As Double
Dim hyp_cuve_C1_DBOsol_mgL As Double
Dim hyp_cuve_N_deox_DBOsol_mgL As Double
Dim surface_media_totale As Double
Dim DBO_apportee_fuite_mgL As Double
Dim ratio_DCO_DBO_source_C As Double
Dim rate_DBO_source_C As Double
Dim cuve_HRTmini_h(nb_cuves_max) As Double
Dim hyp_cuve_filling(nb_cuves_max, nb_type_media) As Double
Dim rate_N_30(nb_type_media) As Double
Dim predenit_NO3_limite As Double
Dim postdenit_NO3_limite As Double
Dim delta_NO3_post_DN As Double
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
Dim a_Pvap(6) As Double
Dim surpresseur_conso_Wh_Nm3mCE(nb_type_surpresseur, nb_surpresseur_gamme_Q, nb_surpresseur_gamme_P) As Double
'Dim impact_regulation(nb_type_regulation) As Double
Dim md_ConcentrationOxygeneDissousExcesRegulation_mgL(1 To nb_type_regulation) As Double
Dim media_Vmin_aeration_Nmh(nb_type_media) As Double
Dim ratio_elec_recirculation_MLSS As Double
Dim ratio_elec_extraction As Double
Dim ratio_elec_agitation As Double
Dim agitation_fct As Double
Dim ratio_DCO_dure_total As Double
Dim DCO_dure_mgL As Double
Dim hyp_outlet_DBO_mgL_mini As Double

'Variable pour le Mox
Dim rate_N_T_design As Double
Dim rate_N_T_design_force As Boolean
Dim rate_N_T_design_f As Double
Dim rate_N_T_reel As Double
Dim rate_N_T_reel_force As Boolean
Dim rate_N_T_reel_f As Double
Dim ratio_NO3f_nit As Double
Dim ratio_O2_nit_Mox As Double
Dim cuve_NH4_outlet_Mox_mgL As Double

'valeurs véhiculées
Dim reel_flux_kgj(nb_flux, nb_parametres_flux) As Double
Dim nominal_flux_kgj(nb_flux, nb_parametres_flux) As Double
Dim reox_DBOremoved As Double
Dim boues_postdenit As Double
Dim cuve_ratio_elimination(nb_cuves_max) As Double
Dim cuve_debit_air_agitation(nb_cuves_max) As Double
Dim nb_etapes_traitement As Integer
Dim deox_nit As Double
Dim cuve_suraeration_brassage_Nm3j(nb_cuves_max) As Double

Dim remarque(nb_erreurs) As String
Private Function aeration_rdt_transfert(ByVal cuve)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim OTR_cw_NoMedia As Double
    Dim OTR_cw_Media As Double
    Dim MF As Double   'média factor

    Dim a1_OTRcw As Double
    Dim a2_OTRcw As Double
    Dim a3_OTRcw As Double
    Dim a4_OTRcw As Double
    Dim filling_OTRcw_min As Double
    Dim filling_OTRcw_max As Double
    Dim filling_OTRcw_retenu As Double


    'On attribue les valeurs des coefficients
    If choix_cuve_aerateur(cuve) = aerateur_fine_bulle Then
        OTR_cw_NoMedia = 17    'gO2/(Nm3.m)
        If choix_media(cuve) = media_chip_M Or choix_media(cuve) = media_chip_P Then
            a1_OTRcw = -12
            a2_OTRcw = 20
            a3_OTRcw = 2
            a4_OTRcw = 70
            filling_OTRcw_min = 10
            filling_OTRcw_max = 55
        Else
            a1_OTRcw = -2
            a2_OTRcw = 7
            a3_OTRcw = 3
            a4_OTRcw = 10
            filling_OTRcw_min = 10
            If choix_media(cuve) = media_K5 Then
                filling_OTRcw_max = 60
            Else
                filling_OTRcw_max = 65
            End If
        End If
    Else  'moyenne bulle
        OTR_cw_NoMedia = 9    'gO2/(Nm3.m)
        If choix_media(cuve) = media_chip_M Or choix_media(cuve) = media_chip_P Then
            a1_OTRcw = -4
            a2_OTRcw = 20
            a3_OTRcw = 7
            a4_OTRcw = 70
            filling_OTRcw_min = 10
            filling_OTRcw_max = 55
        ElseIf choix_media(cuve) = media_K5 Then
            a1_OTRcw = 1
            a2_OTRcw = 7
            a3_OTRcw = 3
            a4_OTRcw = 30
            filling_OTRcw_min = 10
            filling_OTRcw_max = 60
        ElseIf choix_media(cuve) = media_K3 Then
            a1_OTRcw = 5.5
            a2_OTRcw = 7
            a3_OTRcw = 3
            a4_OTRcw = 30
            filling_OTRcw_min = 10
            filling_OTRcw_max = 65
        Else
            a1_OTRcw = 5
            a2_OTRcw = 7
            a3_OTRcw = 3
            a4_OTRcw = 30
            filling_OTRcw_min = 10
            filling_OTRcw_max = 65
        End If
    End If

    'on regarde le domaine de validité de la formule
    If cuve_filling(cuve) * 100 < filling_OTRcw_min Then
        filling_OTRcw_retenu = filling_OTRcw_min
    ElseIf cuve_filling(cuve) * 100 > filling_OTRcw_max Then
        filling_OTRcw_retenu = filling_OTRcw_max
    Else
        filling_OTRcw_retenu = cuve_filling(cuve) * 100
    End If

    'on calcule le facteur OTR_cw_Media
    OTR_cw_Media = OTR_cw_NoMedia + (a1_OTRcw * (filling_OTRcw_retenu + a2_OTRcw) ^ a3_OTRcw) / ((filling_OTRcw_retenu + a2_OTRcw) ^ a3_OTRcw + a4_OTRcw ^ a3_OTRcw)

    'calcul du media factor
    filling_OTRcw_retenu = cuve_filling(cuve) * 100
    MF = 1 + (0.2 * (filling_OTRcw_retenu + 10) ^ 5) / ((filling_OTRcw_retenu + 10) ^ 5 + 20 ^ 5)

    aeration_rdt_transfert = OTR_cw_Media * MF    '* cuve_O2_facteur_alfa(cuve)    'en gO2/(Nm3.m)
    aeration_rdt_transfert = aeration_rdt_transfert / (ratio_kgO2_Nm3air * 1000) * 100    'en %/m


    Exit Function
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.27"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Function

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = nb_erreurs  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil16.Range("MBBR_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub calcul_consommation_electrique(ByRef electricite_II_aeration)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer

    'variables pour l'aération
    Dim gamme_P As Integer
    Dim gamme_Q As Integer
    Dim surpresseur_puissance As Double
    Dim electricite_agitation As Double
    Dim electricite_recirculation_MLSS As Double
    Dim electricite_recirculation As Double
    Dim electricite_extraction As Double
    Dim debit_air_agitation As Double

    'AERATION
    'calcul consommations sur l'aération
    If surpresseur_conso_spec_force = True Then
        surpresseur_conso_spec = surpresseur_conso_spec_f
        'gamme de pression
        If air_P_refoulement_moyenne <= Px1 Then
            gamme_P = surpresseur_gamme_P_0_Px1
        ElseIf air_P_refoulement_moyenne <= Px2 Then
            gamme_P = surpresseur_gamme_P_Px1_Px2
        Else
            gamme_P = surpresseur_gamme_P_Px2_infini
            If choix_surpresseur = surpresseur_roots Then
                'remarque(erreur_choix_surpresseur) = Feuil13.Range("HybAS_erreur_choix_surpresseur").Cells(1, 1)
            End If
        End If
    Else
        'gamme de pression
        If air_P_refoulement_moyenne <= Px1 Then
            gamme_P = surpresseur_gamme_P_0_Px1
        ElseIf air_P_refoulement_moyenne <= Px2 Then
            gamme_P = surpresseur_gamme_P_Px1_Px2
        Else
            gamme_P = surpresseur_gamme_P_Px2_infini
            If choix_surpresseur = surpresseur_roots Then
                'remarque(erreur_choix_surpresseur) = Feuil13.Range("BA_faible_erreur_choix_surpresseur").Cells(1, 1)
            End If
        End If
        'gamme de débit   (pour le nominal)
        gamme_Q = surpresseur_gamme_Q_0_infini
'        surpresseur_conso_spec = surpresseur_conso_Wh_Nm3mCE(choix_surpresseur, gamme_Q, gamme_P) * (1 + impact_regulation(choix_regulation))
        surpresseur_conso_spec = surpresseur_conso_Wh_Nm3mCE(choix_surpresseur, gamme_Q, gamme_P)
    End If
    'calcul de la consommation électrique de l'aération
    electricite_aeration = air_Q_Nm3j * air_P_refoulement_moyenne * surpresseur_conso_spec / 1000
    electricite_II_aeration = electricite_II_aeration + electricite_aeration

    'POMPES DE RECIRCULATION ET EXTRACTION
    electricite_recirculation_MLSS = ratio_elec_recirculation_MLSS * reel_flux_kgj(flux_recirculation_MLSS, repere_Q) * recirculation_MLSS_P_refoulement
    electricite_extraction = ratio_elec_extraction * boues_Q * extraction_P_refoulement

    'AGITATION
    electricite_agitation = 0
    debit_air_agitation = 0
    suraeration_brassage_Nm3j = 0
    For I = 1 To nb_cuves_max
        If I < cuve_C1 Or (I > cuve_N2 And I <> cuve_reox) Then
            If choix_cuve(I) = cuve_oui Then
                electricite_agitation = electricite_agitation + cuve_ratio_agitation_W_m3(I) * cuve_volume(I) * agitation_fct / 1000
            End If
        Else
            If choix_cuve(I) = cuve_oui And choix_agitation_cuve_aeree(I) = cuve_aeree_agitation_mecanique Then
                electricite_agitation = electricite_agitation + cuve_ratio_agitation_W_m3(I) * cuve_volume(I) * agitation_fct / 1000
            End If
        End If
        suraeration_brassage_Nm3j = suraeration_brassage_Nm3j + cuve_suraeration_brassage_Nm3j(I)
        debit_air_agitation = debit_air_agitation + cuve_debit_air_agitation(I)
    Next I

    'CONSOMMATION GLOBALE D'ELECTRICITE
    electricite = electricite_aeration + electricite_agitation + electricite_recirculation_MLSS + electricite_extraction
    electricite_fixe = electricite_agitation + debit_air_agitation / air_Q_Nm3j * electricite_aeration


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Function fct_boues_pdtes(ByVal DBO_eliminee, ByVal Nnitrifie)

    'calcul de la production de boues dans les cuves aérobies

    fct_boues_pdtes = 0.2 * DBO_eliminee + 0.42 * Nnitrifie



End Function

Private Function MV_MES_influent(ByVal MESin, ByVal DCOin, ByVal DBOin)

    'calcul des MV/MES de l'eau brute

    If choix_decanteur_simple = True Then
        MV_MES_influent = 0.782
    ElseIf choix_decanteur_reactif = True Then
        MV_MES_influent = 0.716
    Else
        MV_MES_influent = 1.213 - MESin / (DCOin * 3) - 0.1429 * DCOin / DBOin
        If MV_MES_influent < 0.4 Then
            MV_MES_influent = 0.4
        ElseIf MV_MES_influent > 0.8 Then
            MV_MES_influent = 0.8
        End If
    End If


End Function

Sub simule_C(ByVal cuve, ByVal repere_flux_in, ByRef DBO_removed)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim S_media As Double
    Dim outlet_DBO_soluble_mgL As Double
    Dim rate_DBOin_appliquee_10°C As Double
    Dim rate_DBOin_appliquee_T°C As Double

    'calcul de la surface
    If cuve = cuve_C1 Then
        S_media = cuve_volume(cuve) * cuve_filling(cuve) * media_surface_m2m3(choix_media(cuve))
    Else
        S_media = cuve_volume(cuve_C1) * cuve_filling(cuve_C1) * media_surface_m2m3(choix_media(cuve_C1)) + cuve_volume(cuve_C2) * cuve_filling(cuve_C2) * media_surface_m2m3(choix_media(cuve_C2))
    End If

    'calcul du rate à 10°C
    rate_DBOin_appliquee_T°C = reel_flux_kgj(flux_inlet, repere_DBO) - (reel_flux_kgj(flux_inlet_recirculation, repere_DBO) - reel_flux_kgj(flux_predenit2_C1, repere_DBO))
    rate_DBOin_appliquee_T°C = 1000 * rate_DBOin_appliquee_T°C / S_media
    'rate_DBOin_appliquee_T°C = 1000 * reel_flux_kgj(repere_flux_in, repere_DBO) / S_media
    If pd_TemperatureEauConditionsReelles_°C <= 10 Then
        rate_DBOin_appliquee_10°C = rate_DBOin_appliquee_T°C / (1.07 ^ (pd_TemperatureEauConditionsReelles_°C - 10))
    ElseIf pd_TemperatureEauConditionsReelles_°C <= 25 Then
        rate_DBOin_appliquee_10°C = rate_DBOin_appliquee_T°C / (1.06 ^ (pd_TemperatureEauConditionsReelles_°C - 10))
    Else
        rate_DBOin_appliquee_10°C = rate_DBOin_appliquee_T°C / (1.06 ^ (25 - 10))
    End If

    'calcul de la DBO soluble en sortie de cuve C
    If choix_type_separation = clarif Then
        outlet_DBO_soluble_mgL = (rate_DBOin_appliquee_10°C + 0.273) / 0.8346
    ElseIf choix_type_separation = clarif_polymere Then
        outlet_DBO_soluble_mgL = (rate_DBOin_appliquee_10°C - 0.056) / 0.7688
    Else
        If rate_DBOin_appliquee_10°C <= 3.9 Then
            outlet_DBO_soluble_mgL = (5 / 3.9) * rate_DBOin_appliquee_10°C
        Else
            outlet_DBO_soluble_mgL = (rate_DBOin_appliquee_10°C + 2.5705) / 1.281
        End If
    End If

    If outlet_DBO_soluble_mgL < 0 Then
        outlet_DBO_soluble_mgL = 0
    End If

    'calcul de la charge éliminée
    DBO_removed = reel_flux_kgj(repere_flux_in, repere_DBO) - outlet_DBO_soluble_mgL * reel_flux_kgj(repere_flux_in, repere_Q) / 1000


    'MsgBox (outlet_DBO_soluble_mgL)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.24"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub simule_deox(ByRef interm_O2_dissous_mgL)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim correction_T As Double
    Dim S_media As Double

    Dim O2_mgL As Double
    Dim epsilone_O2 As Double
    Dim O2_fonction As Double
    Dim O2_fonction_derivee As Double
    Dim inlet_O2_mgL As Double
    'Dim NH4_limitant As Boolean

    'calcul de la correction en T
    If pd_TemperatureEauConditionsReelles_°C <= 25 Then
        correction_T = 1.07 ^ (pd_TemperatureEauConditionsReelles_°C - 10)
    Else
        correction_T = 1.07 ^ (25 - 10)
    End If

    'surface de média
    S_media = cuve_volume(cuve_deox) * cuve_filling(cuve_deox) * media_surface_m2m3(choix_media(cuve_deox))

    inlet_O2_mgL = reel_flux_kgj(flux_N2_deox, repere_O2)
    'on retire la respiration
    If pd_TemperatureEauConditionsReelles_°C >= 24 Then
        inlet_O2_mgL = inlet_O2_mgL - 5 * 10 ^ (-3) * 0.012 * 24 * S_media
    Else
        inlet_O2_mgL = inlet_O2_mgL - 2 * 10 ^ (-3) * 1.07 ^ (pd_TemperatureEauConditionsReelles_°C - 10) * 0.012 * 24 * S_media
    End If
    inlet_O2_mgL = inlet_O2_mgL / reel_flux_kgj(flux_N2_deox, repere_Q) * 1000

    O2_mgL = 1

    epsilone_O2 = 1

    Do While Abs(epsilone_O2) > critere_convergence

        'on calcule f et f' (équation sur le traitement de NO3 avec le rate et la surface de média)
        O2_fonction = ((inlet_O2_mgL - O2_mgL) * reel_flux_kgj(flux_N2_deox, repere_Q) / 1000) / ratio_O2_nit - S_media / 1000 * k_nit * ((O2_mgL - 0.5) / 3.2) ^ 0.7 * correction_T
        O2_fonction_derivee = -1 * (reel_flux_kgj(flux_N2_deox, repere_Q) / 1000) / ratio_O2_nit - S_media / 1000 * k_nit * correction_T / (3.2 ^ 0.7) * 0.7 * (O2_mgL - 0.5) ^ (0.7 - 1)
        'on recalcule la charge dénitrifiée
        O2_mgL = O2_mgL - O2_fonction / O2_fonction_derivee

        epsilone_O2 = O2_fonction

    Loop

    interm_O2_dissous_mgL = O2_mgL


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.22"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub simule_reox(ByRef DBO_removed)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim rate_DBOeliminee_10°C As Double
    Dim rate_DBOeliminee_T°C As Double
    Dim S_media As Double


    'calcul des vitesses de nitrification en gNH4/(m2.j)
    rate_DBOeliminee_10°C = 4
    If pd_TemperatureEauConditionsReelles_°C <= 10 Then
        rate_DBOeliminee_T°C = rate_DBOeliminee_10°C * 1.07 ^ (pd_TemperatureEauConditionsReelles_°C - 10)
    ElseIf pd_TemperatureEauConditionsReelles_°C <= 25 Then
        rate_DBOeliminee_T°C = rate_DBOeliminee_10°C * 1.06 ^ (pd_TemperatureEauConditionsReelles_°C - 10)
    Else
        rate_DBOeliminee_T°C = rate_DBOeliminee_10°C * 1.06 ^ (25 - 10)
    End If

    'surface de média
    S_media = cuve_volume(cuve_reox) * cuve_filling(cuve_reox) * media_surface_m2m3(choix_media(cuve_reox))


    DBO_removed = S_media / 1000 * rate_DBOeliminee_T°C


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.25"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub aeration_cuve(ByVal cuve, ByVal repere_flux_in, ByVal vecteur_flux, ByVal temperature, ByVal DBO_removed, ByVal N_removed, ByRef cuve_O2_besoin, ByRef air_Q_Nm3j_cuve, ByRef Sh)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'Dim cuve_O2_besoin As Double
    Dim S_media As Double

    Dim I As Integer

    Dim Patm As Double
    Dim delta_P As Double
    Dim O2_saturation_20_P As Double
    Dim O2_saturation_T_P As Double
    Dim O2_facteur_K As Double
    Dim facteur_alfa As Double

    Dim section_cuve As Double
'    Dim cuve_aeree_Vair_agitation_Nmh As Double  'vitesse minimale pour agitation


    'surface de média
    S_media = cuve_volume(cuve) * cuve_filling(cuve) * media_surface_m2m3(choix_media(cuve))

    'on calcule les besoins en O2 pour éliminer la pollution
    If choix_fct_Mox = True And cuve = cuve_N1 Then
        cuve_O2_besoin = ratio_O2_DBO * DBO_removed + ratio_O2_nit_Mox * N_removed
    Else
        cuve_O2_besoin = ratio_O2_DBO * DBO_removed + ratio_O2_nit * N_removed
    End If
    
    'on traite les sulfures si l'on n'a pas de pré-dénit
    If choix_cuve(cuve_predenit1) = cuve_non Then
        cuve_O2_besoin = cuve_O2_besoin + besoins_O2_HS(Sh)
    End If
    'on calcule la respiration de la biomasse
    If temperature >= 24 Then
        cuve_O2_besoin = cuve_O2_besoin + 5 * 10 ^ (-3) * 0.012 * 24 * S_media
    Else
        cuve_O2_besoin = cuve_O2_besoin + 2 * 10 ^ (-3) * 1.07 ^ (temperature - 10) * 0.012 * 24 * S_media
    End If
    'on n'oublie pas l'excès ou la surconsommation d'O2
    cuve_O2_besoin = cuve_O2_besoin + (vecteur_flux(repere_flux_in + 1, repere_O2) - vecteur_flux(repere_flux_in, repere_O2))


    'attribution du facteur alfa
    facteur_alfa = cuve_O2_facteur_alfa(cuve)

    'Calcul du coefficient K
    Patm = P0 * ((a0_Patm + a1_Patm * altitude) ^ a2_Patm) / 100    'en hPa
    delta_P = a0_delta_P / a1_delta_P * P0 / 100 * (cuve_hauteur(cuve) - hauteur_diffuseur)
    O2_saturation_20_P = 0
    O2_saturation_T_P = 0
    For I = 0 To 4
        O2_saturation_20_P = O2_saturation_20_P + a_O2_saturation(I) * 20 ^ I
        O2_saturation_T_P = O2_saturation_T_P + a_O2_saturation(I) * temperature ^ I
    Next I
    O2_saturation_T_P = O2_saturation_T_P * (Patm + delta_P) / (P0 / 100)
    O2_saturation_20_P = O2_saturation_20_P * (Patm + delta_P) / (P0 / 100)

    O2_facteur_K = facteur_alfa * (O2_facteur_beta * O2_saturation_T_P - cuve_O2_dissous(cuve)) / O2_saturation_20_P
    O2_facteur_K = O2_facteur_K * correctif_T_K ^ (temperature - T_ref_K)

    If cuve_O2_rdt_transfert_force(cuve) = True Then
        cuve_O2_rdt_transfert(cuve) = cuve_O2_rdt_transfert_f(cuve)
    Else
        cuve_O2_rdt_transfert(cuve) = aeration_rdt_transfert(cuve)
    End If

    'calcul consommations sur l'aération
    air_Q_Nm3j_cuve = cuve_O2_besoin / (O2_facteur_K * cuve_O2_rdt_transfert(cuve) / 100 * (cuve_hauteur(cuve) - hauteur_diffuseur) * ratio_kgO2_Nm3air)


    'vérification de la capacité d'agitation
'    If cuve_hauteur(cuve) >= 5 Then
'        cuve_aeree_Vair_agitation_Nmh = 7   'Nm3/(m2.h)
'    Else
'        cuve_aeree_Vair_agitation_Nmh = media_Vmin_aeration_Nmh(choix_media(cuve))   'Nm3/(m2.h)
'    End If

    If choix_agitation_cuve_aeree(cuve) = cuve_aeree_agitation_air Then
        section_cuve = cuve_volume(cuve) / cuve_hauteur(cuve)
        If air_Q_Nm3j_cuve / pd_NOMBRE_HEURE_PAR_JOUR < cuve_aeree_Vair_agitation_Nmh(cuve) * section_cuve Then
            cuve_suraeration_brassage_Nm3j(cuve) = (cuve_aeree_Vair_agitation_Nmh(cuve) * section_cuve - air_Q_Nm3j_cuve / pd_NOMBRE_HEURE_PAR_JOUR) * cuve_aeree_brassage_air_hj(cuve)
            air_Q_Nm3j_cuve = air_Q_Nm3j_cuve + cuve_suraeration_brassage_Nm3j(cuve)
            remarque(erreur_Qair_augmente_agitation(cuve)) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
            remarque(erreur_Qair_augmente_agitation(cuve)) = remarque(erreur_Qair_augmente_agitation(cuve)) + "  " + Feuil16.Range("MBBR_erreur_Qair_augmente_agitation").Cells(1, 1)
        Else
            cuve_suraeration_brassage_Nm3j(cuve) = 0
        End If
        cuve_debit_air_agitation(cuve) = cuve_aeree_Vair_agitation_Nmh(cuve) * section_cuve * cuve_aeree_brassage_air_hj(cuve)
    Else
        cuve_debit_air_agitation(cuve) = 0
        cuve_suraeration_brassage_Nm3j(cuve) = 0
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.26"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Function nitrification_cuve_C(ByVal cuve, ByVal DBO_out_mgL, ByVal temperature)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim correction_T As Double
    Dim S_media As Double

    Dim correction_nit_C As Double

    'calcul de la correction en T
    If temperature <= 25 Then
        correction_T = 1.07 ^ (temperature - 10)
    Else
        correction_T = 1.07 ^ (25 - 10)
    End If

    'calcul de la correction sur la sortie en DBO5
    correction_nit_C = (1 - 0.125 * DBO_out_mgL) * (1 - 0)
    If correction_nit_C < 0 Then
        correction_nit_C = 0
    End If

    'surface de média
    S_media = cuve_volume(cuve) * cuve_filling(cuve) * media_surface_m2m3(choix_media(cuve))

    'on calcule la nit (cas où NH4 non limitant)
    nitrification_cuve_C = S_media / 1000 * correction_nit_C * k_nit * ((cuve_O2_dissous(cuve) - 0.5) / 3.2) ^ 0.7 * correction_T



    Exit Function
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.28"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Function

Sub simule_N(ByVal cuve, ByVal repere_flux_in, ByRef interm_NH4_out_mgL)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim correction_T As Double
    Dim S_media As Double

    Dim nit As Double
    Dim epsilone_nit As Double
    Dim nit_fonction As Double
    Dim nit_fonction_derivee As Double
    Dim NH4_limitant As Boolean

    'surface de média
    S_media = cuve_volume(cuve) * cuve_filling(cuve) * media_surface_m2m3(choix_media(cuve))


    If choix_fct_Mox = True And cuve = cuve_N1 Then
        If rate_N_T_reel_force = True Then
            rate_N_T_reel = rate_N_T_reel_f
        Else
            rate_N_T_reel = rate_N_30(choix_media(cuve)) * 1.07 ^ (pd_TemperatureEauConditionsReelles_°C - 30)
        End If
        nit = S_media / 1000 * rate_N_T_reel
        interm_NH4_out_mgL = (reel_flux_kgj(repere_flux_in, repere_NH4) - nit) / reel_flux_kgj(repere_flux_in, repere_Q) * 1000
    Else
        'calcul de la correction en T
        If pd_TemperatureEauConditionsReelles_°C <= 25 Then
            correction_T = 1.07 ^ (pd_TemperatureEauConditionsReelles_°C - 10)
        Else
            correction_T = 1.07 ^ (25 - 10)
        End If
        'on calcule la nit si NH4 non limitant
        nit = S_media / 1000 * k_nit * ((cuve_O2_dissous(cuve) - 0.5) / 3.2) ^ 0.7 * correction_T
        interm_NH4_out_mgL = (reel_flux_kgj(repere_flux_in, repere_NH4) - nit) / reel_flux_kgj(repere_flux_in, repere_Q) * 1000
        If interm_NH4_out_mgL < (cuve_O2_dissous(cuve) - 0.5) / 3.2 Then
            NH4_limitant = True
        Else
            NH4_limitant = False
        End If
        'On utilise Newton-Raphson sur la charge nitrifiée si NH4 limitant
        If NH4_limitant = True Then
            epsilone_nit = 1
            nit = reel_flux_kgj(repere_flux_in, repere_NH4) - nominal_outlet_NH4_mgL * reel_flux_kgj(repere_flux_in, repere_Q) / 1000
            Do While Abs(epsilone_nit) > critere_convergence
                'on calcule f et f' (équation sur le traitement de NO3 avec le rate et la surface de média)
                nit_fonction = nit - S_media / 1000 * k_nit * ((reel_flux_kgj(repere_flux_in, repere_NH4) - nit) / reel_flux_kgj(repere_flux_in, repere_Q) * 1000) ^ 0.7 * correction_T
                nit_fonction_derivee = 1 - S_media * correction_T / 1000 * k_nit * ((1000 / reel_flux_kgj(repere_flux_in, repere_Q)) ^ 0.7) * (-1) * (reel_flux_kgj(repere_flux_in, repere_NH4) - nit) ^ (-0.3)
                'on recalcule la charge dénitrifiée
                nit = nit - nit_fonction / nit_fonction_derivee
                'on calcule la concentration de NH4 en sortie (on aurait aussi pu créer une variable pour le facteur qui change et le calculer directenment ici)
                interm_NH4_out_mgL = (reel_flux_kgj(repere_flux_in, repere_NH4) - nit) / reel_flux_kgj(repere_flux_in, repere_Q) * 1000
                epsilone_nit = nit_fonction
            Loop
        End If
    End If
    
    If interm_NH4_out_mgL < 0 Then
        interm_NH4_out_mgL = 0
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.21"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub simule_predenit(ByVal cuve, ByVal repere_flux_in, ByVal DBOsol_EB, ByRef interm_NO3_out_mgL)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim correction_NO3 As Double
    Dim correction_T As Double
    Dim S_media As Double

    Dim NO3in As Double
    Dim epsilone_NO3out As Double
    Dim NO3out As Double
    Dim NO3out_fonction As Double
    Dim NO3out_fonction_derivee As Double


    'correction température
    If pd_TemperatureEauConditionsReelles_°C <= 10 Then
        correction_T = 1.07 ^ (pd_TemperatureEauConditionsReelles_°C - 10)
    ElseIf pd_TemperatureEauConditionsReelles_°C <= 25 Then
        correction_T = 1.06 ^ (pd_TemperatureEauConditionsReelles_°C - 10)
    Else
        correction_T = 1.06 ^ (25 - 10)
    End If

    'surface de média
    S_media = cuve_volume(cuve) * cuve_filling(cuve) * media_surface_m2m3(choix_media(cuve))


    NO3in = (reel_flux_kgj(repere_flux_in, repere_NO3) + ratio_NO3eq_O2 * reel_flux_kgj(repere_flux_in, repere_O2)) / reel_flux_kgj(repere_flux_in, repere_Q) * 1000

    'On utilise Newton-Raphson sur NO3out
    epsilone_NO3out = 1
    NO3out = 1
    correction_NO3 = 1

    Do While Abs(epsilone_NO3out) > critere_convergence

        'epsilone_NO3out = NO3out

        'on calcule la correction NO3
        'correction_NO3 = (0.879 * NO3out ^ 0.4381 + 0.2) / (0.879 * 3 ^ 0.4381 + 0.2)
        correction_NO3 = NO3out / (NO3out + 0.4)

        'on calcule f et f' (équation sur le traitement de NO3 avce le rate et la surface de média)
        NO3out_fonction = (NO3in - NO3out) * reel_flux_kgj(repere_flux_in, repere_Q) / 1000 - (correction_NO3 * Exp(Log(0.3) * (-ratio_O2_denit * ((NO3in - NO3out) * reel_flux_kgj(repere_flux_in, repere_Q) / 1000) / DBOsol_EB) ^ 2) * correction_T * S_media / 1000)
        NO3out_fonction_derivee = -reel_flux_kgj(repere_flux_in, repere_Q) / 1000 - (0.4 / ((NO3out + 0.4) ^ 2) * Exp(Log(0.3) * (-ratio_O2_denit * ((NO3in - NO3out) * reel_flux_kgj(repere_flux_in, repere_Q) / 1000) / DBOsol_EB) ^ 2) * correction_T * S_media / 1000) - correction_NO3 * correction_T * S_media / 1000 * (Log(0.3) * (-ratio_O2_denit * reel_flux_kgj(repere_flux_in, repere_Q) / (1000 * DBOsol_EB)) ^ 2 * (-2) * (NO3in - NO3out)) * Exp(Log(0.3) * (-ratio_O2_denit * ((NO3in - NO3out) * reel_flux_kgj(repere_flux_in, repere_Q) / 1000) / DBOsol_EB) ^ 2)
        'on recalcule la charge dénitrifiée
        NO3out = NO3out - NO3out_fonction / NO3out_fonction_derivee


        'epsilone_denit = (epsilone_denit - denit) ^ 2
        epsilone_NO3out = NO3out_fonction

    Loop

    'on calcule la concentration de NO3 en sortie
    If NO3out >= 0 Then
        interm_NO3_out_mgL = NO3out       '(reel_flux_kgj(repere_flux_in, repere_NO3) + ratio_NO3eq_O2 * reel_flux_kgj(repere_flux_in, repere_O2) - denit) / reel_flux_kgj(repere_flux_in, repere_Q) * 1000
    Else
        interm_NO3_out_mgL = 0
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.20"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub simule_postdenit(ByVal cuve, ByVal repere_flux_in, ByRef interm_NO3_out_mgL)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim correction_T As Double
    Dim S_media As Double
    Dim NO3_in_kgj As Double

    Dim epsilone_NO3 As Double
    Dim NO3_mgL As Double
    Dim NO3_fonction As Double
    Dim NO3_fonction_derivee As Double


    'correction température
    If pd_TemperatureEauConditionsReelles_°C <= 10 Then
        correction_T = 1.07 ^ (pd_TemperatureEauConditionsReelles_°C - 10)
    ElseIf pd_TemperatureEauConditionsReelles_°C <= 25 Then
        correction_T = 1.06 ^ (pd_TemperatureEauConditionsReelles_°C - 10)
    Else
        correction_T = 1.06 ^ (25 - 10)
    End If

    'surface de média
    S_media = cuve_volume(cuve) * cuve_filling(cuve) * media_surface_m2m3(choix_media(cuve))

    NO3_in_kgj = reel_flux_kgj(repere_flux_in, repere_NO3) + ratio_NO3eq_O2 * reel_flux_kgj(repere_flux_in, repere_O2)

    'On utilise Newton-Raphson sur la charge dénitrifiée
    epsilone_NO3 = 1
    NO3_mgL = nominal_outlet_NO3_mgL

    Do While Abs(epsilone_NO3) > critere_convergence

        'epsilone_denit = denit

        'on calcule f et f' (équation sur le traitement de NO3 avce le rate et la surface de média)
        NO3_fonction = (NO3_in_kgj - NO3_mgL * reel_flux_kgj(repere_flux_in, repere_Q) / 1000) - S_media * correction_T / 1000 * (2.29 * NO3_mgL / (NO3_mgL + 2.9))
        NO3_fonction_derivee = -reel_flux_kgj(repere_flux_in, repere_Q) / 1000 - S_media * correction_T / 1000 * 2.29 * 2.9 / ((NO3_mgL + 2.9) ^ 2)
        'on recalcule la charge dénitrifiée
        NO3_mgL = NO3_mgL - NO3_fonction / NO3_fonction_derivee

        'epsilone_denit = (epsilone_denit - denit) ^ 2
        epsilone_NO3 = NO3_fonction

    Loop


    If NO3_mgL >= 0 Then
        interm_NO3_out_mgL = NO3_mgL
    Else
        interm_NO3_out_mgL = 0
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.23"

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


Private Sub dimensionne_C_configCN(ByVal cuve, ByVal DBO_eliminee, ByVal rate_DBOin_appliquee_10°C, ByVal Qin)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'Ici on dimensionne les cuves C

    Dim rate_DBOin_appliquee_T°C As Double
    Dim pointe_m3h As Double
    Dim volume_media_necessaire As Double
    Dim filling_HRTmini As Double  'correspond au filling maxi pour respecter le HRT mini

    'calcul du rate à T°C
    If pd_TemperatureEauDimensionnement_°C <= 10 Then
        rate_DBOin_appliquee_T°C = rate_DBOin_appliquee_10°C * 1.07 ^ (pd_TemperatureEauDimensionnement_°C - 10)
    ElseIf pd_TemperatureEauDimensionnement_°C <= 25 Then
        rate_DBOin_appliquee_T°C = rate_DBOin_appliquee_10°C * 1.06 ^ (pd_TemperatureEauDimensionnement_°C - 10)
    Else
        rate_DBOin_appliquee_T°C = rate_DBOin_appliquee_10°C * 1.06 ^ (25 - 10)
    End If

    'calcul du média nécessaire pour obtenir le traitement souhaité
    volume_media_necessaire = 1000 * nominal_flux_kgj(flux_predenit2_C1, repere_DBO) / (rate_DBOin_appliquee_T°C * media_surface_m2m3(choix_media(cuve)))
    pointe_m3h = Qin * pd_CoefficientPointeHydrauliqueTempsPluie / pd_NOMBRE_HEURE_PAR_JOUR
    filling_HRTmini = volume_media_necessaire / (pointe_m3h * cuve_HRTmini_h(cuve))

    'calcul du volume et/ou du taux de remplissage
    If cuve_volume_force(cuve) = True Then
        cuve_volume(cuve) = cuve_volume_f(cuve)
        If cuve_filling_force(cuve) = True Then
            cuve_filling(cuve) = cuve_filling_f(cuve)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 2 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 2
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 3
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
            Else
                remarque(cuve) = ""
            End If
        Else
            cuve_filling(cuve) = volume_media_necessaire / cuve_volume(cuve)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                'on attribue la valeur maximale possible de taux de remplissage
                cuve_filling(cuve) = media_filling_max(cuve, choix_media(cuve))
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve) = ""
                    'remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then   'on ne peut pas avoir l'erreur 3 si on ne modifie pas le filling
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
            Else
                remarque(cuve) = ""
            End If
        End If
    Else
        If cuve_filling_force(cuve) = True Then
            cuve_filling(cuve) = cuve_filling_f(cuve)
            cuve_volume(cuve) = volume_media_necessaire / cuve_filling(cuve)
            'erreurs possibles: filling trop élevé et HRT insuffisant
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                Else
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            Else
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            End If
        Else
            If filling_HRTmini < hyp_cuve_filling(cuve, choix_media(cuve)) Then
                cuve_filling(cuve) = filling_HRTmini
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_diminue").Cells(1, 1)
            Else
                cuve_filling(cuve) = hyp_cuve_filling(cuve, choix_media(cuve))
            End If
            'pas d'erreur possible
        End If
        cuve_volume(cuve) = volume_media_necessaire / cuve_filling(cuve)
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.12"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Sub dimensionne_C_C(ByRef DBO_eliminee, ByRef volume_media_total)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'permet de calculer la surface nécessaire
    Dim outlet_DBO_soluble_mgL As Double
    Dim rate_DBOin_appliquee_10°C As Double
    Dim rate_DBOin_appliquee_T°C As Double


    If choix_type_separation = clarif Then
        outlet_DBO_soluble_mgL = 0.5 * nominal_outlet_DBO_mgL - 5
        'message d'erreur à rajouter
        If nominal_outlet_DBO_mgL <= 21.7 Then
            rate_DBOin_appliquee_10°C = 3.9
        ElseIf nominal_outlet_DBO_mgL >= 37.5 Then
            rate_DBOin_appliquee_10°C = 11.5
        Else
            rate_DBOin_appliquee_10°C = 0.8346 * outlet_DBO_soluble_mgL - 0.273
        End If
    ElseIf choix_type_separation = clarif_polymere Then
        outlet_DBO_soluble_mgL = 0.5275 * nominal_outlet_DBO_mgL - 3
        'message d'erreur à rajouter
        If nominal_outlet_DBO_mgL <= 11.33 Then
            rate_DBOin_appliquee_10°C = 3.9
        ElseIf nominal_outlet_DBO_mgL >= 34.029 Then
            rate_DBOin_appliquee_10°C = 11.5
        Else
            rate_DBOin_appliquee_10°C = 0.7688 * outlet_DBO_soluble_mgL + 0.056
        End If
    Else
        outlet_DBO_soluble_mgL = 0.3936 * nominal_outlet_DBO_mgL - 1.1553
        'message d'erreur à rajouter
        If nominal_outlet_DBO_mgL <= 10 Then
            rate_DBOin_appliquee_10°C = 3.9
        ElseIf nominal_outlet_DBO_mgL >= 24.895 Then
            rate_DBOin_appliquee_10°C = 11.5
        Else
            rate_DBOin_appliquee_10°C = 1.281 * outlet_DBO_soluble_mgL - 2.505
        End If
    End If

    'calcul du rate à T°C
    If pd_TemperatureEauDimensionnement_°C <= 10 Then
        rate_DBOin_appliquee_T°C = rate_DBOin_appliquee_10°C * 1.07 ^ (pd_TemperatureEauDimensionnement_°C - 10)
    ElseIf pd_TemperatureEauDimensionnement_°C <= 25 Then
        rate_DBOin_appliquee_T°C = rate_DBOin_appliquee_10°C * 1.06 ^ (pd_TemperatureEauDimensionnement_°C - 10)
    Else
        rate_DBOin_appliquee_T°C = rate_DBOin_appliquee_10°C * 1.06 ^ (25 - 10)
    End If

    DBO_eliminee = nominal_flux_kgj(flux_predenit2_C1, repere_DBO) - outlet_DBO_soluble_mgL * nominal_flux_kgj(flux_predenit2_C1, repere_Q) / 1000

    volume_media_total = (nominal_flux_kgj(flux_inlet, repere_DBO) - (nominal_flux_kgj(flux_inlet_recirculation, repere_DBO) - nominal_flux_kgj(flux_predenit2_C1, repere_DBO))) * 1000 / rate_DBOin_appliquee_T°C



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.11"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Sub dimensionne_C_configCseul(ByVal cuve, ByVal DBO_eliminee, ByVal surface_media, ByVal Qin)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'on calcule le volume de la cuve C à partir de la surface de média

    Dim pointe_m3h As Double
    Dim volume_media_necessaire As Double
    Dim filling_HRTmini As Double  'correspond au filling maxi pour respecter le HRT mini


    'calcul du média nécessaire pour obtenir le traitement souhaité
    volume_media_necessaire = surface_media / media_surface_m2m3(choix_media(cuve))    'ATTENTION ici c'est différent car on connaît déjà la surface de média dans la cuve
    pointe_m3h = Qin * pd_CoefficientPointeHydrauliqueTempsPluie / pd_NOMBRE_HEURE_PAR_JOUR
    filling_HRTmini = volume_media_necessaire / (pointe_m3h * cuve_HRTmini_h(cuve))

    'calcul du volume et/ou du taux de remplissage
    If cuve_volume_force(cuve) = True Then
        cuve_volume(cuve) = cuve_volume_f(cuve)
        If cuve_filling_force(cuve) = True Then
            cuve_filling(cuve) = cuve_filling_f(cuve)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 2 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 2
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 3
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
            Else
                remarque(cuve) = ""
            End If
        Else
            cuve_filling(cuve) = volume_media_necessaire / cuve_volume(cuve)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                'on attribue la valeur maximale possible de taux de remplissage
                cuve_filling(cuve) = media_filling_max(cuve, choix_media(cuve))
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve) = ""
                    'remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then   'on ne peut pas avoir l'erreur 3 si on ne modifie pas le filling
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
            Else
                remarque(cuve) = ""
            End If
        End If
    Else
        If cuve_filling_force(cuve) = True Then
            cuve_filling(cuve) = cuve_filling_f(cuve)
            cuve_volume(cuve) = volume_media_necessaire / cuve_filling(cuve)
            'erreurs possibles: filling trop élevé et HRT insuffisant
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                Else
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            Else
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            End If
        Else
            If filling_HRTmini < hyp_cuve_filling(cuve, choix_media(cuve)) Then
                cuve_filling(cuve) = filling_HRTmini
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_diminue").Cells(1, 1)
            Else
                cuve_filling(cuve) = hyp_cuve_filling(cuve, choix_media(cuve))
            End If
            'pas d'erreur possible
        End If
        cuve_volume(cuve) = volume_media_necessaire / cuve_filling(cuve)
    End If




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.13"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Private Sub dimensionne_deox(ByVal deox_nit, ByVal Qin)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'ici on dimensionne la DéOx
    Dim rate_N_10°C As Double
    Dim rate_N_T_design°C As Double
    Dim pointe_m3h As Double
    Dim volume_media_necessaire As Double
    Dim filling_HRTmini As Double  'correspond au filling maxi pour respecter le HRT mini

    'prise en compte de la respiration (après remarques de Dorothée, voir si meilleurs résultats)
    Dim O2_conso_max As Double
    Dim interm_conso_O2 As Double
    Dim S_media_min As Double
    Dim S_media_max As Double
    Dim S_media As Double


    'calcul des vitesses de nitrification en gNH4/(m2.j)
    rate_N_10°C = k_nit * ((cuve_O2_dissous(cuve_deox) - 0.5) / 3.2) ^ 0.7
    If pd_TemperatureEauDimensionnement_°C <= 25 Then
        rate_N_T_design°C = rate_N_10°C * 1.07 ^ (pd_TemperatureEauDimensionnement_°C - 10)
    Else
        rate_N_T_design°C = rate_N_10°C * 1.07 ^ (25 - 10)
    End If


    O2_conso_max = deox_nit * ratio_O2_nit

    S_media_min = 0
    S_media_max = 1000 * deox_nit / rate_N_T_design°C

    S_media = (S_media_min + S_media_max) / 2

    Do While (S_media_max - S_media_min) / S_media > critere_convergence
        interm_conso_O2 = ratio_O2_nit * rate_N_T_design°C * S_media / 1000    'nitrification
        If pd_TemperatureEauDimensionnement_°C >= 24 Then
            interm_conso_O2 = interm_conso_O2 + 5 * 10 ^ (-3) * 0.012 * 24 * S_media
        Else
            interm_conso_O2 = interm_conso_O2 + 2 * 10 ^ (-3) * 1.07 ^ (pd_TemperatureEauDimensionnement_°C - 10) * 0.012 * 24 * S_media
        End If
        If interm_conso_O2 < O2_conso_max Then
            S_media_min = S_media
        ElseIf interm_conso_O2 > O2_conso_max Then
            S_media_max = S_media
        Else
            S_media_max = S_media
            S_media_min = S_media
        End If
        S_media = (S_media_min + S_media_max) / 2
        deox_nit = rate_N_T_design°C * S_media / 1000    'nitrification
    Loop

    volume_media_necessaire = S_media / media_surface_m2m3(choix_media(cuve_deox))

    'calcul du média nécessaire pour obtenir le traitement souhaité
    'volume_media_necessaire = 1000 * deox_nit / (rate_N_T_design°C * media_surface_m2m3(choix_media(cuve_deox)))
    pointe_m3h = Qin * pd_CoefficientPointeHydrauliqueTempsPluie / pd_NOMBRE_HEURE_PAR_JOUR
    filling_HRTmini = volume_media_necessaire / (pointe_m3h * cuve_HRTmini_h(cuve_deox))

    'calcul du volume et/ou du taux de remplissage
    If cuve_volume_force(cuve_deox) = True Then
        cuve_volume(cuve_deox) = cuve_volume_f(cuve_deox)
        If cuve_filling_force(cuve_deox) = True Then
            cuve_filling(cuve_deox) = cuve_filling_f(cuve_deox)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve_deox) > media_filling_max(cuve_deox, choix_media(cuve_deox)) Then
                If cuve_volume(cuve_deox) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve_deox) * cuve_filling(cuve_deox) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                        remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                        remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve_deox) * cuve_filling(cuve_deox) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                    remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                    remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve_deox) < volume_media_necessaire / filling_HRTmini Then
                If cuve_volume(cuve_deox) * cuve_filling(cuve_deox) < volume_media_necessaire Then   'erreurs 2 3
                    remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                    remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 2
                    remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                    remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve_deox) * cuve_filling(cuve_deox) < volume_media_necessaire Then   'erreurs 3
                remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
            Else
                remarque(cuve_deox) = ""
            End If
        Else
            cuve_filling(cuve_deox) = volume_media_necessaire / cuve_volume(cuve_deox)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve_deox) > media_filling_max(cuve_deox, choix_media(cuve_deox)) Then
                'on attribue la valeur maximale possible de taux de remplissage
                cuve_filling(cuve_deox) = media_filling_max(cuve_deox, choix_media(cuve_deox))
                If cuve_volume(cuve_deox) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve_deox) * cuve_filling(cuve_deox) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                        'remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                        'remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve_deox) * cuve_filling(cuve_deox) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                    'remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve_deox) = ""
                    'remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                    'remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve_deox) < volume_media_necessaire / filling_HRTmini Then   'on ne peut pas avoir l'erreur 3 si on ne modifie pas le filling
                remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
            Else
                remarque(cuve_deox) = ""
            End If
        End If
    Else
        If cuve_filling_force(cuve_deox) = True Then
            cuve_filling(cuve_deox) = cuve_filling_f(cuve_deox)
            cuve_volume(cuve_deox) = volume_media_necessaire / cuve_filling(cuve_deox)
            'erreurs possibles: filling trop élevé et HRT insuffisant
            If cuve_filling(cuve_deox) > media_filling_max(cuve_deox, choix_media(cuve_deox)) Then
                If cuve_volume(cuve_deox) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                    remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                Else
                    remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                    remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            Else
                If cuve_volume(cuve_deox) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                    remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            End If
        Else
            If filling_HRTmini < hyp_cuve_filling(cuve_deox, choix_media(cuve_deox)) Then
                cuve_filling(cuve_deox) = filling_HRTmini
                remarque(cuve_deox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_deox, 1)
                remarque(cuve_deox) = remarque(cuve_deox) + " " + Feuil16.Range("MBBR_erreur_filling_diminue").Cells(1, 1)
            Else
                cuve_filling(cuve_deox) = hyp_cuve_filling(cuve_deox, choix_media(cuve_deox))
            End If
            'pas d'erreur possible
        End If
        cuve_volume(cuve_deox) = volume_media_necessaire / cuve_filling(cuve_deox)
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.14"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Private Sub dimensionne_reox(ByVal reox_DBO_eliminee, ByVal Qin)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'ici on dimensionne la DéOx
    Dim rate_DBOeliminee_10°C As Double
    Dim rate_DBOeliminee_T°C As Double
    Dim pointe_m3h As Double
    Dim volume_media_necessaire As Double
    Dim filling_HRTmini As Double  'correspond au filling maxi pour respecter le HRT mini

    'calcul des vitesses de nitrification en gNH4/(m2.j)
    rate_DBOeliminee_10°C = 4
    If pd_TemperatureEauDimensionnement_°C <= 10 Then
        rate_DBOeliminee_T°C = rate_DBOeliminee_10°C * 1.07 ^ (pd_TemperatureEauDimensionnement_°C - 10)
    ElseIf pd_TemperatureEauDimensionnement_°C <= 25 Then
        rate_DBOeliminee_T°C = rate_DBOeliminee_10°C * 1.06 ^ (pd_TemperatureEauDimensionnement_°C - 10)
    Else
        rate_DBOeliminee_T°C = rate_DBOeliminee_10°C * 1.06 ^ (25 - 10)
    End If


    'calcul du média nécessaire pour obtenir le traitement souhaité
    volume_media_necessaire = 1000 * reox_DBO_eliminee / (rate_DBOeliminee_T°C * media_surface_m2m3(choix_media(cuve_reox)))
    pointe_m3h = Qin * pd_CoefficientPointeHydrauliqueTempsPluie / pd_NOMBRE_HEURE_PAR_JOUR
    filling_HRTmini = volume_media_necessaire / (pointe_m3h * cuve_HRTmini_h(cuve_reox))

    'calcul du volume et/ou du taux de remplissage
    If cuve_volume_force(cuve_reox) = True Then
        cuve_volume(cuve_reox) = cuve_volume_f(cuve_reox)
        If cuve_filling_force(cuve_reox) = True Then
            cuve_filling(cuve_reox) = cuve_filling_f(cuve_reox)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve_reox) > media_filling_max(cuve_reox, choix_media(cuve_reox)) Then
                If cuve_volume(cuve_reox) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve_reox) * cuve_filling(cuve_reox) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                        remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                        remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve_reox) * cuve_filling(cuve_reox) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                    remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                    remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve_reox) < volume_media_necessaire / filling_HRTmini Then
                If cuve_volume(cuve_reox) * cuve_filling(cuve_reox) < volume_media_necessaire Then   'erreurs 2 3
                    remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                    remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 2
                    remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                    remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve_reox) * cuve_filling(cuve_reox) < volume_media_necessaire Then   'erreurs 3
                remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
            Else
                remarque(cuve_reox) = ""
            End If
        Else
            cuve_filling(cuve_reox) = volume_media_necessaire / cuve_volume(cuve_reox)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve_reox) > media_filling_max(cuve_reox, choix_media(cuve_reox)) Then
                'on attribue la valeur maximale possible de taux de remplissage
                cuve_filling(cuve_reox) = media_filling_max(cuve_reox, choix_media(cuve_reox))
                If cuve_volume(cuve_reox) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve_reox) * cuve_filling(cuve_reox) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                        'remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                        'remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve_reox) * cuve_filling(cuve_reox) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                    'remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve_reox) = ""
                    'remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                    'remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve_reox) < volume_media_necessaire / filling_HRTmini Then   'on ne peut pas avoir l'erreur 3 si on ne modifie pas le filling
                remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
            Else
                remarque(cuve_reox) = ""
            End If
        End If
    Else
        If cuve_filling_force(cuve_reox) = True Then
            cuve_filling(cuve_reox) = cuve_filling_f(cuve_reox)
            cuve_volume(cuve_reox) = volume_media_necessaire / cuve_filling(cuve_reox)
            'erreurs possibles: filling trop élevé et HRT insuffisant
            If cuve_filling(cuve_reox) > media_filling_max(cuve_reox, choix_media(cuve_reox)) Then
                If cuve_volume(cuve_reox) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                    remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                Else
                    remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                    remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            Else
                If cuve_volume(cuve_reox) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                    remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            End If
        Else
            If filling_HRTmini < hyp_cuve_filling(cuve_reox, choix_media(cuve_reox)) Then
                cuve_filling(cuve_reox) = filling_HRTmini
                remarque(cuve_reox) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve_reox, 1)
                remarque(cuve_reox) = remarque(cuve_reox) + " " + Feuil16.Range("MBBR_erreur_filling_diminue").Cells(1, 1)
            Else
                cuve_filling(cuve_reox) = hyp_cuve_filling(cuve_reox, choix_media(cuve_reox))
            End If
            'pas d'erreur possible
        End If
        cuve_volume(cuve_reox) = volume_media_necessaire / cuve_filling(cuve_reox)
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.19"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub



Private Sub dimensionne_N(ByVal cuve, ByVal cuve_nit, ByVal NH4_out_mgL, ByVal Qin)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'permet de dimensionner les cuves N

    Dim rate_N_10°C As Double
    Dim rate_N_T_design°C As Double
    Dim pointe_m3h As Double
    Dim volume_media_necessaire As Double
    Dim filling_HRTmini As Double  'correspond au filling maxi pour respecter le HRT mini

    If choix_fct_Mox = True And cuve = cuve_N1 Then
        'calcul du Mox
        If rate_N_T_design_force = True Then
            rate_N_T_design = rate_N_T_design_f
        Else
            rate_N_T_design = rate_N_30(choix_media(cuve)) * 1.07 ^ (pd_TemperatureEauDimensionnement_°C - 30)
        End If
        'calcul du volume
        volume_media_necessaire = 1000 * cuve_nit / (rate_N_T_design * media_surface_m2m3(choix_media(cuve)))
        pointe_m3h = Qin * pd_CoefficientPointeHydrauliqueTempsPluie / pd_NOMBRE_HEURE_PAR_JOUR
        filling_HRTmini = volume_media_necessaire / (pointe_m3h * 0.2)
    Else
        'calcul des vitesses de nitrification en gNH4/(m2.j)
        If (cuve_O2_dissous(cuve) - 0.5) / 3.2 <= NH4_out_mgL Then
            rate_N_10°C = k_nit * ((cuve_O2_dissous(cuve) - 0.5) / 3.2) ^ 0.7
        Else
            rate_N_10°C = k_nit * NH4_out_mgL ^ 0.7
        End If
        If pd_TemperatureEauDimensionnement_°C <= 25 Then
            rate_N_T_design°C = rate_N_10°C * 1.07 ^ (pd_TemperatureEauDimensionnement_°C - 10)
        Else
            rate_N_T_design°C = rate_N_10°C * 1.07 ^ (25 - 10)
        End If
        'calcul du média nécessaire pour obtenir le traitement souhaité
        volume_media_necessaire = 1000 * cuve_nit / (rate_N_T_design°C * media_surface_m2m3(choix_media(cuve)))
        pointe_m3h = Qin * pd_CoefficientPointeHydrauliqueTempsPluie / pd_NOMBRE_HEURE_PAR_JOUR
        filling_HRTmini = volume_media_necessaire / (pointe_m3h * cuve_HRTmini_h(cuve))
    End If

    'calcul du volume et/ou du taux de remplissage
    If cuve_volume_force(cuve) = True Then
        cuve_volume(cuve) = cuve_volume_f(cuve)
        If cuve_filling_force(cuve) = True Then
            cuve_filling(cuve) = cuve_filling_f(cuve)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 2 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 2
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 3
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
            Else
                remarque(cuve) = ""
            End If
        Else
            cuve_filling(cuve) = volume_media_necessaire / cuve_volume(cuve)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                'on attribue la valeur maximale possible de taux de remplissage
                cuve_filling(cuve) = media_filling_max(cuve, choix_media(cuve))
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve) = ""
                    'remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then   'on ne peut pas avoir l'erreur 3 si on ne modifie pas le filling
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
            Else
                remarque(cuve) = ""
            End If
        End If
    Else
        If cuve_filling_force(cuve) = True Then
            cuve_filling(cuve) = cuve_filling_f(cuve)
            cuve_volume(cuve) = volume_media_necessaire / cuve_filling(cuve)
            'erreurs possibles: filling trop élevé et HRT insuffisant
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                Else
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            Else
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            End If
        Else
            If filling_HRTmini < hyp_cuve_filling(cuve, choix_media(cuve)) Then
                cuve_filling(cuve) = filling_HRTmini
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_diminue").Cells(1, 1)
            Else
                cuve_filling(cuve) = hyp_cuve_filling(cuve, choix_media(cuve))
            End If
            'pas d'erreur possible
        End If
        cuve_volume(cuve) = volume_media_necessaire / cuve_filling(cuve)
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.15"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Sub dimensionne_postdenit(ByVal cuve, ByVal denit, ByVal cuve_NO3_mgL, ByVal Qin)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'Ici on dimensionne les cuves de Post DN
    Dim rate_DN_10°C As Double
    Dim rate_DN_T°C As Double
    Dim pointe_m3h As Double
    Dim volume_media_necessaire As Double
    Dim filling_HRTmini As Double  'correspond au filling maxi pour respecter le HRT mini

    'calcul du rate de dénit à 10°C et à T°C
    rate_DN_10°C = 2.29 * cuve_NO3_mgL / (cuve_NO3_mgL + 2.9)
    If pd_TemperatureEauDimensionnement_°C <= 10 Then
        rate_DN_T°C = rate_DN_10°C * 1.07 ^ (pd_TemperatureEauDimensionnement_°C - 10)
    ElseIf pd_TemperatureEauDimensionnement_°C <= 25 Then
        rate_DN_T°C = rate_DN_10°C * 1.06 ^ (pd_TemperatureEauDimensionnement_°C - 10)
    Else
        rate_DN_T°C = rate_DN_10°C * 1.06 ^ (25 - 10)
    End If


    'calcul du média nécessaire pour obtenir le traitement souhaité
    volume_media_necessaire = 1000 * denit / (rate_DN_T°C * media_surface_m2m3(choix_media(cuve)))
    pointe_m3h = Qin * pd_CoefficientPointeHydrauliqueTempsPluie / pd_NOMBRE_HEURE_PAR_JOUR
    filling_HRTmini = volume_media_necessaire / (pointe_m3h * cuve_HRTmini_h(cuve))

    'calcul du volume et/ou du taux de remplissage
    If cuve_volume_force(cuve) = True Then
        cuve_volume(cuve) = cuve_volume_f(cuve)
        If cuve_filling_force(cuve) = True Then
            cuve_filling(cuve) = cuve_filling_f(cuve)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 2 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 2
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 3
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
            Else
                remarque(cuve) = ""
            End If
        Else
            cuve_filling(cuve) = volume_media_necessaire / cuve_volume(cuve)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                'on attribue la valeur maximale possible de taux de remplissage
                cuve_filling(cuve) = media_filling_max(cuve, choix_media(cuve))
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve) = ""
                    'remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then   'on ne peut pas avoir l'erreur 3 si on ne modifie pas le filling
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
            Else
                remarque(cuve) = ""
            End If
        End If
    Else
        If cuve_filling_force(cuve) = True Then
            cuve_filling(cuve) = cuve_filling_f(cuve)
            cuve_volume(cuve) = volume_media_necessaire / cuve_filling(cuve)
            'erreurs possibles: filling trop élevé et HRT insuffisant
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                Else
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            Else
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            End If
        Else
            If filling_HRTmini < hyp_cuve_filling(cuve, choix_media(cuve)) Then
                cuve_filling(cuve) = filling_HRTmini
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_diminue").Cells(1, 1)
            Else
                cuve_filling(cuve) = hyp_cuve_filling(cuve, choix_media(cuve))
            End If
            'pas d'erreur possible
        End If
        cuve_volume(cuve) = volume_media_necessaire / cuve_filling(cuve)
    End If




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.16"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Private Sub dimensionne_predenit(ByVal cuve, ByVal NO3_out_mgL, ByVal DBOsol_EB, ByVal denit, ByVal Qin)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'Ici on dimensionne une cuve de pré-dénit

    Dim DBO_conso As Double
    Dim correction_NO3 As Double
    Dim rate_DN_10°C As Double
    Dim rate_DN_T°C As Double
    Dim pointe_m3h As Double
    Dim volume_media_necessaire As Double
    Dim filling_HRTmini As Double  'correspond au filling maxi pour respecter le HRT mini

    DBO_conso = -ratio_O2_denit * denit     'en kg/j

    'correction sur la valeur de sortie souhaitée en NO3
    'correction_NO3 = (0.879 * NO3_out_mgL ^ 0.4381 + 0.2) / (0.879 * 3 ^ 0.4381 + 0.2)
    'If correction_NO3 > 1 Then
    '    correction_NO3 = 1       'on prend le min entre formule au dessus et 1
    'End If
    correction_NO3 = NO3_out_mgL / (NO3_out_mgL + 0.4)


    'calcul rate à 10°C et rate à T    en gNO3/(m2.j)
    rate_DN_10°C = correction_NO3 * Exp(Log(0.3) * (DBO_conso / DBOsol_EB) ^ 2)
    If pd_TemperatureEauDimensionnement_°C <= 10 Then
        rate_DN_T°C = rate_DN_10°C * 1.07 ^ (pd_TemperatureEauDimensionnement_°C - 10)
    ElseIf pd_TemperatureEauDimensionnement_°C <= 25 Then
        rate_DN_T°C = rate_DN_10°C * 1.06 ^ (pd_TemperatureEauDimensionnement_°C - 10)
    Else
        rate_DN_T°C = rate_DN_10°C * 1.06 ^ (25 - 10)
    End If

    'calcul du média nécessaire pour obtenir le traitement souhaité
    volume_media_necessaire = 1000 * denit / (rate_DN_T°C * media_surface_m2m3(choix_media(cuve)))
    pointe_m3h = Qin * pd_CoefficientPointeHydrauliqueTempsPluie / pd_NOMBRE_HEURE_PAR_JOUR
    filling_HRTmini = volume_media_necessaire / (pointe_m3h * cuve_HRTmini_h(cuve))

    'calcul du volume et/ou du taux de remplissage
    If cuve_volume_force(cuve) = True Then
        cuve_volume(cuve) = cuve_volume_f(cuve)
        If cuve_filling_force(cuve) = True Then
            cuve_filling(cuve) = cuve_filling_f(cuve)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 2 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 2
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 3
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
            Else
                remarque(cuve) = ""
            End If
        Else
            cuve_filling(cuve) = volume_media_necessaire / cuve_volume(cuve)
            'test pour le taux max de média et le HRT min
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                'on attribue la valeur maximale possible de taux de remplissage
                cuve_filling(cuve) = media_filling_max(cuve, choix_media(cuve))
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    If cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 2 3
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                    Else   'erreurs 1 2
                        remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                        'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                        remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                    End If
                ElseIf cuve_volume(cuve) * cuve_filling(cuve) < volume_media_necessaire Then   'erreurs 1 3
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_media_insuffisant").Cells(1, 1)
                Else   'erreurs 1
                    remarque(cuve) = ""
                    'remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    'remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            ElseIf cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then   'on ne peut pas avoir l'erreur 3 si on ne modifie pas le filling
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
            Else
                remarque(cuve) = ""
            End If
        End If
    Else
        If cuve_filling_force(cuve) = True Then
            cuve_filling(cuve) = cuve_filling_f(cuve)
            cuve_volume(cuve) = volume_media_necessaire / cuve_filling(cuve)
            'erreurs possibles: filling trop élevé et HRT insuffisant
            If cuve_filling(cuve) > media_filling_max(cuve, choix_media(cuve)) Then
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                Else
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_trop").Cells(1, 1)
                End If
            Else
                If cuve_volume(cuve) < volume_media_necessaire / filling_HRTmini Then
                    remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                    remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_HRT_insuffisant").Cells(1, 1)
                End If
            End If
        Else
            If filling_HRTmini < hyp_cuve_filling(cuve, choix_media(cuve)) Then
                cuve_filling(cuve) = filling_HRTmini
                remarque(cuve) = Feuil16.Range("MBBR_erreur_nom_cuve").Cells(cuve, 1)
                remarque(cuve) = remarque(cuve) + " " + Feuil16.Range("MBBR_erreur_filling_diminue").Cells(1, 1)
            Else
                cuve_filling(cuve) = hyp_cuve_filling(cuve, choix_media(cuve))
            End If
            'pas d'erreur possible
        End If
        cuve_volume(cuve) = volume_media_necessaire / cuve_filling(cuve)
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.17"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub dimensionnement(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef Sh, ByRef DCO_apportee_methanol, ByRef DCO_conso_bio)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer

    Dim N_nit_total As Double   'charge à nitrifier au total
    Dim predenit_denit As Double
    Dim DBOsol_EB As Double
    Dim repere_flux_in As Integer
    Dim interm_NO3_out_mgL As Double
    Dim interm_O2_mgL As Double
    Dim epsilone_recirculation_mlss As Double
    Dim epsilone_DBO As Double
    Dim epsilone_NH4 As Double
    Dim epsilone_NO3 As Double
    Dim cuve_nit As Double
    Dim interm_DBO_eliminee As Double
    Dim rate_DBOin_appliquee_10°C As Double
    Dim DBO_eliminee_C_C As Double
    Dim cuve_C_C_media_total_m2 As Double
    Dim interm_surface_media As Double
    Dim postdenit_Nassimile As Double
    Dim reox_Nassimile As Double
    Dim cuve_NH4_removed As Double
    Dim cuve_NH4_outlet_mgL As Double
    Dim postdenit_denit As Double
    Dim interm_cuve_NO3_mgL As Double
    Dim DBO_apportee_postdenit_kgj As Double
    Dim outlet_DBO_soluble_mgL As Double
    Dim cuve_O2_besoin As Double
    Dim cuve_air_Q_Nm3j As Double
    Dim nit_cuve_C As Double

    Dim interm_DCO_conso_bio As Double

    Dim boues_pdtes(nb_cuves_max) As Double
    Dim boues_pdtes_aerobie As Double


    'on calcule les MV/MES de l'influent
    boues_MV_MES = MV_MES_influent(MES, DCO, DBO) * MES


    'on attribue les valeurs des flux en entrée
    nominal_flux_kgj(flux_inlet, repere_Q) = Q
    nominal_flux_kgj(flux_inlet, repere_DCO) = DCO
    nominal_flux_kgj(flux_inlet, repere_DBO) = DBO
    nominal_flux_kgj(flux_inlet, repere_MES) = MES
    nominal_flux_kgj(flux_inlet, repere_NK) = NK
    nominal_flux_kgj(flux_inlet, repere_NH4) = NH4
    nominal_flux_kgj(flux_inlet, repere_NO3) = NO3
    nominal_flux_kgj(flux_inlet, repere_Pt) = Pt
    nominal_flux_kgj(flux_inlet, repere_O2) = 0

    interm_DCO_conso_bio = DBO


    'calcul de la charge à nitrifier
    N_nit_total = NH4 - (nominal_outlet_NH4_mgL + Norga_dure_soluble) * (Q - boues_Q) / 1000 + EB_rate_Norgasol_Norga * (NK - NH4) + (1 - EB_rate_Norgasol_Norga) * rate_Norgapart_hydrolyse * (NK - NH4)
    N_nit_total = N_nit_total - ratio_N_assimile * (DBO - nominal_flux_kgj(flux_separationMLSS_postdenit1, repere_DBO)) - ratio_N_assimile_Capporte * boues_postdenit - ratio_N_assimile * reox_DBOremoved


    'ON RENTRE DANS LE DUR!!!!
    epsilone_recirculation_mlss = 1 + critere_convergence

    'initialisation du flux recirculé
    If choix_cuve(cuve_predenit1) = cuve_oui Then
        nominal_flux_kgj(flux_recirculation_MLSS, repere_Q) = (N_nit_total - Q / 1000 * (nominal_last_aerated_tank_NO3_mgL - nominal_predenit_NO3_mgL)) / (Q / 1000 * (nominal_last_aerated_tank_NO3_mgL - nominal_predenit_NO3_mgL)) * Q
        nominal_flux_kgj(flux_recirculation_MLSS, repere_DBO) = nominal_outlet_DBO_mgL * nominal_flux_kgj(flux_recirculation_MLSS, repere_Q) / 1000
        nominal_flux_kgj(flux_recirculation_MLSS, repere_DCO) = nominal_flux_kgj(flux_recirculation_MLSS, repere_DBO)
        'nominal_flux_kgj(flux_recirculation_MLSS, repere_MES) = MES
        nominal_flux_kgj(flux_recirculation_MLSS, repere_NK) = (nominal_outlet_NH4_mgL + Norga_dure_soluble) * nominal_flux_kgj(flux_recirculation_MLSS, repere_Q) / 1000
        nominal_flux_kgj(flux_recirculation_MLSS, repere_NH4) = nominal_outlet_NH4_mgL * nominal_flux_kgj(flux_recirculation_MLSS, repere_Q) / 1000
        nominal_flux_kgj(flux_recirculation_MLSS, repere_NO3) = nominal_last_aerated_tank_NO3_mgL * nominal_flux_kgj(flux_recirculation_MLSS, repere_Q) / 1000
        nominal_flux_kgj(flux_recirculation_MLSS, repere_Pt) = Pt / Q * nominal_flux_kgj(flux_recirculation_MLSS, repere_Q)
        nominal_flux_kgj(flux_recirculation_MLSS, repere_O2) = 0   'ça ne vaut pas forcément la peine de s'embêter à chercher quelle cuve va servir à recirculer (nit 1, nit 2, déox)
    End If

    Do While epsilone_recirculation_mlss > critere_convergence

        surface_media_totale = 0
        O2_besoin_total = 0
        air_Q_Nm3j = 0

        epsilone_DBO = nominal_flux_kgj(flux_recirculation_MLSS, repere_DBO)
        epsilone_NH4 = nominal_flux_kgj(flux_recirculation_MLSS, repere_NH4)
        epsilone_NO3 = nominal_flux_kgj(flux_recirculation_MLSS, repere_NO3)

        'Calcul du flux après recirculation
        For j = 1 To nb_parametres_flux
            nominal_flux_kgj(flux_inlet_recirculation, j) = nominal_flux_kgj(flux_inlet, j) + nominal_flux_kgj(flux_recirculation_MLSS, j)
        Next j


        'PRE-DENIT
        If choix_cuve(cuve_predenit1) = cuve_oui Then
            If nominal_recirculation_MLSS_taux_force = True Then
                nominal_recirculation_MLSS_taux = nominal_recirculation_MLSS_taux_f
            Else
                'calcul du taux de recirculation de MLSS  (UNIQUEMENT SI PRE-DENIT CAR SINON PAS D'INTERET)
                nominal_recirculation_MLSS_taux = (N_nit_total - Q / 1000 * (nominal_last_aerated_tank_NO3_mgL - nominal_predenit_NO3_mgL)) / (Q / 1000 * (nominal_last_aerated_tank_NO3_mgL - nominal_predenit_NO3_mgL))
            End If
            predenit_denit = nominal_flux_kgj(flux_inlet_recirculation, repere_NO3) + ratio_NO3eq_O2 * nominal_flux_kgj(flux_inlet_recirculation, repere_O2) - nominal_predenit_NO3_mgL * Q / 1000 * (1 + nominal_recirculation_MLSS_taux)
            interm_DBO_eliminee = 0
            For I = cuve_predenit1 To cuve_predenit2
                repere_flux_in = flux_inlet_recirculation + (I - cuve_predenit1)  'repère le flux en entrée
                DBOsol_EB = nominal_flux_kgj(flux_inlet, repere_DBO) * (EB_rate_DBOsol_DBO + (1 - EB_rate_DBOsol_DBO) * rate_DBOpart_hydrolyse) + nominal_flux_kgj(flux_recirculation_MLSS, repere_DBO) - interm_DBO_eliminee
                If choix_cuve(I) = cuve_oui Then
                    interm_NO3_out_mgL = nominal_flux_kgj(repere_flux_in, repere_NO3) + ratio_NO3eq_O2 * nominal_flux_kgj(repere_flux_in, repere_O2) - cuve_ratio_elimination(I) * predenit_denit
                    interm_NO3_out_mgL = interm_NO3_out_mgL * 1000 / ((1 + nominal_recirculation_MLSS_taux) * Q)
                    Call dimensionne_predenit(I, interm_NO3_out_mgL, DBOsol_EB, predenit_denit * cuve_ratio_elimination(I), nominal_flux_kgj(repere_flux_in, repere_Q))
                    'on attribue les valeurs des flux
                    nominal_flux_kgj(repere_flux_in + 1, repere_Q) = nominal_flux_kgj(repere_flux_in, repere_Q)
                    nominal_flux_kgj(repere_flux_in + 1, repere_NO3) = interm_NO3_out_mgL * nominal_flux_kgj(repere_flux_in + 1, repere_Q) / 1000
                    interm_DBO_eliminee = -ratio_O2_denit * predenit_denit * cuve_ratio_elimination(I)
                    nominal_flux_kgj(repere_flux_in + 1, repere_DBO) = nominal_flux_kgj(repere_flux_in, repere_DBO) - interm_DBO_eliminee
                    'faut-il calculer la production de boues et véhiculer les MES au fur et à mesure ou est-ce suffisant de le faire à la fin
                    nominal_flux_kgj(repere_flux_in + 1, repere_DCO) = nominal_flux_kgj(repere_flux_in, repere_DCO) - interm_DBO_eliminee
                    nominal_flux_kgj(repere_flux_in + 1, repere_NK) = nominal_flux_kgj(repere_flux_in, repere_NK) - ratio_N_assimile * interm_DBO_eliminee
                    nominal_flux_kgj(repere_flux_in + 1, repere_NH4) = nominal_flux_kgj(repere_flux_in, repere_NH4) - ratio_N_assimile * interm_DBO_eliminee
                    nominal_flux_kgj(repere_flux_in + 1, repere_Pt) = nominal_flux_kgj(repere_flux_in, repere_Pt) - ratio_P_assimile * interm_DBO_eliminee
                    nominal_flux_kgj(repere_flux_in + 1, repere_O2) = 0
                    'calcul de la surface totale de média
                    surface_media_totale = surface_media_totale + cuve_volume(I) * cuve_filling(I) * media_surface_m2m3(choix_media(I))
                Else
                    For j = 1 To nb_parametres_flux
                        nominal_flux_kgj(repere_flux_in + 1, j) = nominal_flux_kgj(repere_flux_in, j)
                    Next j
                End If
            Next I
        Else
            'on attribue les valeurs aux flux intermédiaires jusqu'à l'entrée de la cuve C1 et on dit qu'il n' y a pas besoin de recirculation
            nominal_recirculation_MLSS_taux = 0
            For I = flux_predenit1_predenit2 To flux_predenit2_C1
                For j = 1 To nb_parametres_flux
                    nominal_flux_kgj(I, j) = nominal_flux_kgj(flux_inlet_recirculation, j)
                Next j
            Next I
        End If


        'Traitement C
        If choix_cuve(cuve_C1) = cuve_oui And nb_etapes_traitement = 1 Then
            Call dimensionne_C_C(DBO_eliminee_C_C, cuve_C_C_media_total_m2)
            nit_cuve_C = 0
            For I = cuve_C1 To cuve_C2
                'définir le rate sur chaque cuve (afin d'obtenir la même surface développée dans les deux cuves
                repere_flux_in = flux_predenit2_C1 + (I - cuve_C1)  'repère le flux en entrée
                interm_DBO_eliminee = cuve_ratio_elimination(I) * DBO_eliminee_C_C
                interm_surface_media = cuve_C_C_media_total_m2 * 0.5
                Call dimensionne_C_configCseul(I, interm_DBO_eliminee, interm_surface_media, nominal_flux_kgj(repere_flux_in, repere_Q))
                'on attribue les valeurs des flux
                nominal_flux_kgj(repere_flux_in + 1, repere_Q) = nominal_flux_kgj(repere_flux_in, repere_Q)
                nominal_flux_kgj(repere_flux_in + 1, repere_DBO) = nominal_flux_kgj(repere_flux_in, repere_DBO) - interm_DBO_eliminee
                'faut-il calculer la production de boues et véhiculer les MES au fur et à mesure ou est-ce suffisant de le faire à la fin
                nominal_flux_kgj(repere_flux_in + 1, repere_DCO) = nominal_flux_kgj(repere_flux_in, repere_DCO) - interm_DBO_eliminee
                'nit dans cuve C
                cuve_nit = nitrification_cuve_C(I, nominal_flux_kgj(repere_flux_in + 1, repere_DBO) / nominal_flux_kgj(repere_flux_in + 1, repere_Q) * 1000, pd_TemperatureEauDimensionnement_°C)
                nit_cuve_C = nit_cuve_C + cuve_nit
                nominal_flux_kgj(repere_flux_in + 1, repere_NK) = nominal_flux_kgj(repere_flux_in, repere_NK) - ratio_N_assimile * interm_DBO_eliminee - cuve_nit
                nominal_flux_kgj(repere_flux_in + 1, repere_NH4) = nominal_flux_kgj(repere_flux_in, repere_NH4) - ratio_N_assimile * interm_DBO_eliminee - cuve_nit
                nominal_flux_kgj(repere_flux_in + 1, repere_NO3) = nominal_flux_kgj(repere_flux_in, repere_NO3) + cuve_nit
                nominal_flux_kgj(repere_flux_in + 1, repere_Pt) = nominal_flux_kgj(repere_flux_in, repere_Pt) - ratio_P_assimile * interm_DBO_eliminee
                nominal_flux_kgj(repere_flux_in + 1, repere_O2) = cuve_O2_dissous(I) * nominal_flux_kgj(repere_flux_in + 1, repere_Q) / 1000
                surface_media_totale = surface_media_totale + cuve_volume(I) * cuve_filling(I) * media_surface_m2m3(choix_media(I))
                boues_pdtes(I) = fct_boues_pdtes(interm_DBO_eliminee, cuve_nit)
                'on calcule l'aération
                Call aeration_cuve(I, repere_flux_in, nominal_flux_kgj, pd_TemperatureEauDimensionnement_°C, interm_DBO_eliminee, cuve_nit, cuve_O2_besoin, cuve_air_Q_Nm3j, Sh)
                O2_besoin_total = O2_besoin_total + cuve_O2_besoin
                air_Q_Nm3j = air_Q_Nm3j + cuve_air_Q_Nm3j
            Next I
        ElseIf choix_cuve(cuve_C1) = cuve_oui Then   'forcément une seule cuve
            repere_flux_in = flux_predenit2_C1
            'définir la concentration en sortie de C1 et la DBO éliminée
            interm_DBO_eliminee = nominal_flux_kgj(repere_flux_in, repere_DBO) - hyp_cuve_C1_DBOsol_mgL * nominal_flux_kgj(repere_flux_in, repere_Q) / 1000
            'définir le rate
            rate_DBOin_appliquee_10°C = hyp_rate_DBOin_appliquee_10°C_CN
            Call dimensionne_C_configCN(cuve_C1, interm_DBO_eliminee, rate_DBOin_appliquee_10°C, nominal_flux_kgj(repere_flux_in, repere_Q))
            'nitrification cuve_C
            nit_cuve_C = nitrification_cuve_C(cuve_C1, hyp_cuve_C1_DBOsol_mgL, pd_TemperatureEauDimensionnement_°C)
            For I = flux_C1_C2 To flux_C2_N1
                'on attribue les valeurs des flux
                nominal_flux_kgj(I, repere_Q) = nominal_flux_kgj(repere_flux_in, repere_Q)
                nominal_flux_kgj(I, repere_DBO) = nominal_flux_kgj(repere_flux_in, repere_DBO) - interm_DBO_eliminee
                'faut-il calculer la production de boues et véhiculer les MES au fur et à mesure ou est-ce suffisant de le faire à la fin
                nominal_flux_kgj(I, repere_DCO) = nominal_flux_kgj(repere_flux_in, repere_DCO) - interm_DBO_eliminee
                nominal_flux_kgj(I, repere_NK) = nominal_flux_kgj(repere_flux_in, repere_NK) - ratio_N_assimile * interm_DBO_eliminee - nit_cuve_C
                nominal_flux_kgj(I, repere_NH4) = nominal_flux_kgj(repere_flux_in, repere_NH4) - ratio_N_assimile * interm_DBO_eliminee - nit_cuve_C
                nominal_flux_kgj(I, repere_NO3) = nominal_flux_kgj(repere_flux_in, repere_NO3) + nit_cuve_C
                nominal_flux_kgj(I, repere_Pt) = nominal_flux_kgj(repere_flux_in, repere_Pt) - ratio_P_assimile * interm_DBO_eliminee
                nominal_flux_kgj(I, repere_O2) = cuve_O2_dissous(I) * nominal_flux_kgj(repere_flux_in + 1, repere_Q) / 1000
            Next I
            boues_pdtes(cuve_C1) = fct_boues_pdtes(interm_DBO_eliminee, nit_cuve_C)
            surface_media_totale = surface_media_totale + cuve_volume(cuve_C1) * cuve_filling(cuve_C1) * media_surface_m2m3(choix_media(cuve_C1))
            'on calcule l'aération
            Call aeration_cuve(cuve_C1, repere_flux_in, nominal_flux_kgj, pd_TemperatureEauDimensionnement_°C, interm_DBO_eliminee, nit_cuve_C, cuve_O2_besoin, cuve_air_Q_Nm3j, Sh)
            O2_besoin_total = O2_besoin_total + cuve_O2_besoin
            air_Q_Nm3j = air_Q_Nm3j + cuve_air_Q_Nm3j
        Else
            For I = flux_C1_C2 To flux_C2_N1
                For j = 1 To nb_parametres_flux
                    nominal_flux_kgj(I, j) = nominal_flux_kgj(flux_predenit2_C1, j)
                Next j
            Next I
        End If


        'Traitement N
        If choix_cuve(cuve_N1) = cuve_oui Then
            For I = cuve_N1 To cuve_N2
                repere_flux_in = flux_C2_N1 + (I - cuve_N1)  'repère le flux en entrée
                If choix_cuve(I) = cuve_oui Then
                    'A voir la validité de ce qui suit: on a considéré au design que l'on traite encore un peu de DBO soluble donc on va assimiler du N que l'on va retrancher à la charge en entrée pour faciliter les calculs
                    interm_DBO_eliminee = nominal_flux_kgj(repere_flux_in, repere_DBO) - hyp_cuve_N_deox_DBOsol_mgL * nominal_flux_kgj(repere_flux_in, repere_Q) / 1000
                    'nominal_flux_kgj(repere_flux_in, repere_DBO) = nominal_flux_kgj(repere_flux_in, repere_DBO) - interm_DBO_eliminee
                    nominal_flux_kgj(repere_flux_in, repere_NH4) = nominal_flux_kgj(repere_flux_in, repere_NH4) - ratio_N_assimile * interm_DBO_eliminee
                    nominal_flux_kgj(repere_flux_in, repere_NK) = nominal_flux_kgj(repere_flux_in, repere_NK) - ratio_N_assimile * interm_DBO_eliminee
                    If choix_fct_Mox = True And I = cuve_N1 And ratio_elimination_N1_N12_force = False Then
                        cuve_NH4_removed = nominal_flux_kgj(repere_flux_in, repere_NH4) - cuve_NH4_outlet_Mox_mgL / 1000 * nominal_flux_kgj(repere_flux_in, repere_Q)
                        cuve_ratio_elimination(cuve_N1) = 1 / nominal_flux_kgj(flux_C2_N1, repere_NH4) * (cuve_NH4_removed + (1 + nominal_recirculation_MLSS_taux) * (Q / 1000 * nominal_outlet_NH4_mgL + postdenit_Nassimile + reox_Nassimile))
                        ratio_elimination_N1_N12 = cuve_ratio_elimination(cuve_N1)
                        cuve_ratio_elimination(cuve_N2) = 1 - cuve_ratio_elimination(cuve_N1)
                        cuve_nit = cuve_ratio_elimination(I) * (N_nit_total - deox_nit - nit_cuve_C)
                    Else
                        'on reprend le cours
                        cuve_nit = cuve_ratio_elimination(I) * (N_nit_total - deox_nit - nit_cuve_C)
                        cuve_NH4_removed = cuve_ratio_elimination(I) * (nominal_flux_kgj(flux_C2_N1, repere_NH4) - (1 + nominal_recirculation_MLSS_taux) * (Q / 1000 * nominal_outlet_NH4_mgL + postdenit_Nassimile + reox_Nassimile))
                        cuve_NH4_outlet_mgL = (nominal_flux_kgj(repere_flux_in, repere_NH4) - cuve_NH4_removed) * 1000 / nominal_flux_kgj(repere_flux_in, repere_Q)
                    End If
                    Call dimensionne_N(I, cuve_nit, cuve_NH4_outlet_mgL, nominal_flux_kgj(repere_flux_in, repere_Q))
                    'on attribue les valeurs des flux
                    nominal_flux_kgj(repere_flux_in + 1, repere_Q) = nominal_flux_kgj(repere_flux_in, repere_Q)
                    If choix_fct_Mox = True Then
                        If I = cuve_N1 Then 'cuve pour le Mox
                            nominal_flux_kgj(repere_flux_in + 1, repere_NO3) = nominal_flux_kgj(repere_flux_in, repere_NO3) + cuve_nit * ratio_NO3f_nit
                        ElseIf I = cuve_N2 Then  'cuve normale
                            nominal_flux_kgj(repere_flux_in + 1, repere_NO3) = nominal_flux_kgj(repere_flux_in, repere_NO3) + cuve_nit
                        End If
                    Else
                        nominal_flux_kgj(repere_flux_in + 1, repere_NO3) = nominal_flux_kgj(repere_flux_in, repere_NO3) + cuve_nit
                    End If
                    nominal_flux_kgj(repere_flux_in + 1, repere_DBO) = nominal_flux_kgj(repere_flux_in, repere_DBO) - interm_DBO_eliminee
                    'faut-il calculer la production de boues et véhiculer les MES au fur et à mesure ou est-ce suffisant de le faire à la fin
                    nominal_flux_kgj(repere_flux_in + 1, repere_DCO) = nominal_flux_kgj(repere_flux_in, repere_DCO) - interm_DBO_eliminee
                    nominal_flux_kgj(repere_flux_in + 1, repere_NK) = nominal_flux_kgj(repere_flux_in, repere_NK) - cuve_nit
                    nominal_flux_kgj(repere_flux_in + 1, repere_NH4) = nominal_flux_kgj(repere_flux_in, repere_NH4) - cuve_NH4_removed
                    nominal_flux_kgj(repere_flux_in + 1, repere_Pt) = nominal_flux_kgj(repere_flux_in, repere_Pt)
                    nominal_flux_kgj(repere_flux_in + 1, repere_O2) = cuve_O2_dissous(I) * nominal_flux_kgj(repere_flux_in + 1, repere_Q) / 1000
                    boues_pdtes(I) = fct_boues_pdtes(interm_DBO_eliminee, cuve_nit)
                    surface_media_totale = surface_media_totale + cuve_volume(I) * cuve_filling(I) * media_surface_m2m3(choix_media(I))
                    'on calcule l'aération
                    Call aeration_cuve(I, repere_flux_in, nominal_flux_kgj, pd_TemperatureEauDimensionnement_°C, interm_DBO_eliminee, cuve_nit, cuve_O2_besoin, cuve_air_Q_Nm3j, Sh)
                    O2_besoin_total = O2_besoin_total + cuve_O2_besoin
                    air_Q_Nm3j = air_Q_Nm3j + cuve_air_Q_Nm3j
                Else
                    For j = 1 To nb_parametres_flux
                        nominal_flux_kgj(repere_flux_in + 1, j) = nominal_flux_kgj(repere_flux_in, j)
                    Next j
                End If
            Next I
        Else
            'on attribue les valeurs aux flux intermédiaires jusqu'à l'entrée de la cuve DéOx
            For I = flux_N1_N2 To flux_N2_deox
                For j = 1 To nb_parametres_flux
                    nominal_flux_kgj(I, j) = nominal_flux_kgj(flux_C2_N1, j)
                Next j
            Next I
        End If


        'DéOx
        If choix_cuve(cuve_deox) = cuve_oui Then
            interm_O2_mgL = nominal_flux_kgj(flux_N2_deox, repere_O2) / nominal_flux_kgj(flux_N2_deox, repere_Q) * 1000
            'on calcule la nitrification maximale que l'on pourra avoir
            deox_nit = ((interm_O2_mgL - cuve_O2_dissous(cuve_deox)) * Q * (1 + nominal_recirculation_MLSS_taux) / 1000) / ratio_O2_nit
            repere_flux_in = flux_N2_deox  'repère le flux en entrée
            'A voir la validité de ce qui suit: on a considéré au design que l'on traite encore un peu de DBO soluble donc on va assimiler du N que l'on va retrancher à la charge en entrée pour faciliter les calculs
            interm_DBO_eliminee = nominal_flux_kgj(flux_N2_deox, repere_DBO) - hyp_cuve_N_deox_DBOsol_mgL * nominal_flux_kgj(flux_N2_deox, repere_Q) / 1000
            'nominal_flux_kgj(flux_N2_deox, repere_DBO) = nominal_flux_kgj(flux_N2_deox, repere_DBO) - interm_DBO_eliminee
            nominal_flux_kgj(flux_N2_deox, repere_NH4) = nominal_flux_kgj(flux_N2_deox, repere_NH4) - ratio_N_assimile * interm_DBO_eliminee
            nominal_flux_kgj(flux_N2_deox, repere_NK) = nominal_flux_kgj(flux_N2_deox, repere_NK) - ratio_N_assimile * interm_DBO_eliminee
            'on continue les calculs
            Call dimensionne_deox(deox_nit, nominal_flux_kgj(repere_flux_in, repere_Q))
            'on attribue les valeurs des flux
            nominal_flux_kgj(repere_flux_in + 1, repere_Q) = nominal_flux_kgj(repere_flux_in, repere_Q)
            nominal_flux_kgj(repere_flux_in + 1, repere_NO3) = nominal_flux_kgj(repere_flux_in, repere_NO3) + deox_nit
            nominal_flux_kgj(repere_flux_in + 1, repere_DBO) = nominal_flux_kgj(repere_flux_in, repere_DBO) - interm_DBO_eliminee
            'faut-il calculer la production de boues et véhiculer les MES au fur et à mesure ou est-ce suffisant de le faire à la fin
            nominal_flux_kgj(repere_flux_in + 1, repere_DCO) = nominal_flux_kgj(repere_flux_in, repere_DCO) - interm_DBO_eliminee
            nominal_flux_kgj(repere_flux_in + 1, repere_NK) = nominal_flux_kgj(repere_flux_in, repere_NK) - deox_nit
            nominal_flux_kgj(repere_flux_in + 1, repere_NH4) = nominal_flux_kgj(repere_flux_in, repere_NH4) - deox_nit
            nominal_flux_kgj(repere_flux_in + 1, repere_Pt) = nominal_flux_kgj(repere_flux_in, repere_Pt)
            nominal_flux_kgj(repere_flux_in + 1, repere_O2) = cuve_O2_dissous(I) * nominal_flux_kgj(repere_flux_in + 1, repere_Q) / 1000
            boues_pdtes(cuve_deox) = fct_boues_pdtes(interm_DBO_eliminee, deox_nit)
            surface_media_totale = surface_media_totale + cuve_volume(cuve_deox) * cuve_filling(cuve_deox) * media_surface_m2m3(choix_media(cuve_deox))
        Else
            For j = 1 To nb_parametres_flux
                nominal_flux_kgj(flux_deox_separationMLSS, j) = nominal_flux_kgj(flux_N2_deox, j)
            Next j
        End If


        'Calcul du flux de recirculation
        For j = 1 To nb_parametres_flux
            nominal_flux_kgj(flux_recirculation_MLSS, j) = nominal_flux_kgj(flux_deox_separationMLSS, j) * nominal_recirculation_MLSS_taux / (1 + nominal_recirculation_MLSS_taux)
        Next j

        epsilone_DBO = epsilone_DBO - nominal_flux_kgj(flux_recirculation_MLSS, repere_DBO)
        epsilone_NH4 = epsilone_NH4 - nominal_flux_kgj(flux_recirculation_MLSS, repere_NH4)
        epsilone_NO3 = epsilone_NO3 - nominal_flux_kgj(flux_recirculation_MLSS, repere_NO3)

        epsilone_DBO = epsilone_DBO ^ 2
        epsilone_NH4 = epsilone_NH4 ^ 2
        epsilone_NO3 = epsilone_NO3 ^ 2

        epsilone_recirculation_mlss = epsilone_DBO + epsilone_NH4 + epsilone_NO3

    Loop


    'Calcul du flux post séparation de la recirculation
    For j = 1 To nb_parametres_flux
        nominal_flux_kgj(flux_separationMLSS_postdenit1, j) = nominal_flux_kgj(flux_deox_separationMLSS, j) - nominal_flux_kgj(flux_recirculation_MLSS, j)
    Next j

    outlet_DBO_soluble_mgL = nominal_flux_kgj(flux_separationMLSS_postdenit1, repere_DBO) / nominal_flux_kgj(flux_separationMLSS_postdenit1, repere_Q) * 1000

    'Post-Dénitrification
    If choix_cuve(cuve_postdenit1) = cuve_oui Then   '************* à reprendre car si cas où pas assez de NH4 pour assimilation, pas nécessaire d'ajouter autant de C car NO3 assimilé, vérifier raisonnement avec Xavier LE TALLEC **************
        postdenit_denit = nominal_flux_kgj(flux_separationMLSS_postdenit1, repere_NO3) + ratio_NO3eq_O2 * nominal_flux_kgj(flux_separationMLSS_postdenit1, repere_O2)
        postdenit_denit = postdenit_denit - nominal_outlet_NO3_mgL * nominal_flux_kgj(flux_separationMLSS_postdenit1, repere_Q) / 1000
        DBO_apportee_postdenit_kgj = rate_DBO_apportee_denit_nominal * postdenit_denit
        postdenit_carbone_apporte_flux = DBO_apportee_postdenit_kgj / rate_DBO_source_C
        boues_postdenit = ratio_boues_DBO_apportee * DBO_apportee_postdenit_kgj
        For I = cuve_postdenit1 To cuve_postdenit2
            repere_flux_in = flux_separationMLSS_postdenit1 + (I - cuve_postdenit1)  'repère le flux en entrée
            If choix_cuve(I) = cuve_oui Then
                interm_cuve_NO3_mgL = nominal_flux_kgj(repere_flux_in, repere_NO3) + ratio_NO3eq_O2 * nominal_flux_kgj(repere_flux_in, repere_O2)
                interm_cuve_NO3_mgL = (interm_cuve_NO3_mgL - postdenit_denit * cuve_ratio_elimination(I)) * 1000 / nominal_flux_kgj(flux_separationMLSS_postdenit1, repere_Q)
                Call dimensionne_postdenit(I, postdenit_denit * cuve_ratio_elimination(I), interm_cuve_NO3_mgL, nominal_flux_kgj(repere_flux_in, repere_Q))
                'on attribue les valeurs des flux
                nominal_flux_kgj(repere_flux_in + 1, repere_Q) = nominal_flux_kgj(repere_flux_in, repere_Q)
                nominal_flux_kgj(repere_flux_in + 1, repere_DBO) = nominal_flux_kgj(repere_flux_in, repere_DBO)
                'faut-il calculer la production de boues et véhiculer les MES au fur et à mesure ou est-ce suffisant de le faire à la fin
                nominal_flux_kgj(repere_flux_in + 1, repere_DCO) = nominal_flux_kgj(repere_flux_in, repere_DCO)
                If nominal_flux_kgj(repere_flux_in, repere_NH4) < cuve_ratio_elimination(I) * boues_postdenit * ratio_N_assimile_Capporte Then
                    nominal_flux_kgj(repere_flux_in + 1, repere_NO3) = nominal_flux_kgj(repere_flux_in, repere_NO3) - (cuve_ratio_elimination(I) * boues_postdenit * ratio_N_assimile_Capporte - nominal_flux_kgj(repere_flux_in, repere_NH4))
                    nominal_flux_kgj(repere_flux_in + 1, repere_NO3) = nominal_flux_kgj(repere_flux_in + 1, repere_NO3) - cuve_ratio_elimination(I) * postdenit_denit
                    nominal_flux_kgj(repere_flux_in + 1, repere_NK) = nominal_flux_kgj(repere_flux_in, repere_NK) - nominal_flux_kgj(repere_flux_in, repere_NH4)
                    nominal_flux_kgj(repere_flux_in + 1, repere_NH4) = 0
                Else
                    nominal_flux_kgj(repere_flux_in + 1, repere_NH4) = nominal_flux_kgj(repere_flux_in, repere_NH4) - cuve_ratio_elimination(I) * boues_postdenit * ratio_N_assimile_Capporte
                    nominal_flux_kgj(repere_flux_in + 1, repere_NO3) = nominal_flux_kgj(repere_flux_in, repere_NO3) - cuve_ratio_elimination(I) * postdenit_denit
                    nominal_flux_kgj(repere_flux_in + 1, repere_NK) = nominal_flux_kgj(repere_flux_in, repere_NK) - cuve_ratio_elimination(I) * boues_postdenit * ratio_N_assimile_Capporte
                End If
                nominal_flux_kgj(repere_flux_in + 1, repere_Pt) = nominal_flux_kgj(repere_flux_in, repere_Pt)
                nominal_flux_kgj(repere_flux_in + 1, repere_O2) = 0
                surface_media_totale = surface_media_totale + cuve_volume(I) * cuve_filling(I) * media_surface_m2m3(choix_media(I))
            Else
                For j = 1 To nb_parametres_flux
                    nominal_flux_kgj(repere_flux_in + 1, j) = nominal_flux_kgj(repere_flux_in, j)
                Next j
            End If
        Next I
        nominal_flux_kgj(flux_postdenit2_reox, repere_DBO) = nominal_flux_kgj(flux_separationMLSS_postdenit1, repere_DBO) + DBO_apportee_fuite_mgL * nominal_flux_kgj(flux_postdenit2_reox, repere_Q) / 1000
        nominal_flux_kgj(flux_postdenit2_reox, repere_DCO) = nominal_flux_kgj(flux_separationMLSS_postdenit1, repere_DCO) + ratio_DCO_DBO_source_C * DBO_apportee_fuite_mgL * nominal_flux_kgj(flux_postdenit2_reox, repere_Q) / 1000
    Else
        'on attribue les valeurs aux flux intermédiaires jusqu'à l'entrée de la cuve DéOx
        For I = flux_postdenit1_postdenit2 To flux_postdenit2_reox
            For j = 1 To nb_parametres_flux
                nominal_flux_kgj(I, j) = nominal_flux_kgj(flux_separationMLSS_postdenit1, j)
            Next j
        Next I
    End If


    'Ré-Oxygénation
    If choix_cuve(cuve_reox) = cuve_oui Then
        repere_flux_in = flux_postdenit2_reox  'repère le flux en entrée
        interm_DBO_eliminee = nominal_flux_kgj(repere_flux_in, repere_DBO) - outlet_DBO_soluble_mgL * nominal_flux_kgj(repere_flux_in, repere_Q) / 1000
        Call dimensionne_reox(interm_DBO_eliminee, nominal_flux_kgj(repere_flux_in, repere_Q))
        'on attribue les valeurs des flux
        nominal_flux_kgj(repere_flux_in + 1, repere_Q) = nominal_flux_kgj(repere_flux_in, repere_Q)
        nominal_flux_kgj(repere_flux_in + 1, repere_NO3) = nominal_flux_kgj(repere_flux_in, repere_NO3)
        nominal_flux_kgj(repere_flux_in + 1, repere_DBO) = nominal_flux_kgj(repere_flux_in, repere_DBO) - interm_DBO_eliminee
        'faut-il calculer la production de boues et véhiculer les MES au fur et à mesure ou est-ce suffisant de le faire à la fin
        nominal_flux_kgj(repere_flux_in + 1, repere_DCO) = nominal_flux_kgj(repere_flux_in, repere_DCO)
        nominal_flux_kgj(repere_flux_in + 1, repere_NK) = nominal_flux_kgj(repere_flux_in, repere_NK)
        nominal_flux_kgj(repere_flux_in + 1, repere_NH4) = nominal_flux_kgj(repere_flux_in, repere_NH4)
        nominal_flux_kgj(repere_flux_in + 1, repere_Pt) = nominal_flux_kgj(repere_flux_in, repere_Pt)
        nominal_flux_kgj(repere_flux_in + 1, repere_O2) = cuve_O2_dissous(cuve_reox) * nominal_flux_kgj(repere_flux_in + 1, repere_Q) / 1000
        boues_pdtes(cuve_reox) = fct_boues_pdtes(interm_DBO_eliminee, 0)
        surface_media_totale = surface_media_totale + cuve_volume(cuve_reox) * cuve_filling(cuve_reox) * media_surface_m2m3(choix_media(cuve_reox))
        'on calcule l'aération
        Call aeration_cuve(cuve_reox, repere_flux_in, nominal_flux_kgj, pd_TemperatureEauDimensionnement_°C, interm_DBO_eliminee, 0, cuve_O2_besoin, cuve_air_Q_Nm3j, Sh)
        O2_besoin_total = O2_besoin_total + cuve_O2_besoin
        air_Q_Nm3j = air_Q_Nm3j + cuve_air_Q_Nm3j
    Else
        'on attribue les valeurs aux flux intermédiaires jusqu'à l'entrée de la cuve DéOx
        For j = 1 To nb_parametres_flux
            nominal_flux_kgj(flux_outlet, j) = nominal_flux_kgj(flux_postdenit2_reox, j)
        Next j
    End If

    volume_total_bassins = 0
    For I = 1 To nb_cuves_max
        volume_total_bassins = volume_total_bassins + cuve_volume(I)
    Next I

    boues_pdtes_aerobie = 0
    For I = 1 To nb_cuves_max
        boues_pdtes_aerobie = boues_pdtes_aerobie + boues_pdtes(I)
    Next I
    boues_MV_MES = boues_MV_MES * (1 - rate_DBOpart_hydrolyse) + boues_pdtes_aerobie + boues_postdenit
    boues_pdtes_aerobie = boues_pdtes_aerobie + nominal_flux_kgj(flux_inlet, repere_MES) * (1 - rate_DBOpart_hydrolyse)

    'Séparation liquide/solide
    Call separation_boues(nominal_flux_kgj(flux_inlet, repere_MES), nominal_flux_kgj(flux_inlet, repere_DBO), nominal_flux_kgj(flux_inlet, repere_Q), pd_TemperatureEauDimensionnement_°C, nominal_outlet_DBO_mgL, nominal_flux_kgj(flux_outlet, repere_DBO) / nominal_flux_kgj(flux_outlet, repere_Q) * 1000, boues_pdtes_aerobie, False)

    If boues_MV_MES_force = True Then
        boues_MV_MES = boues_MV_MES_f
    End If


    Q = nominal_flux_kgj(flux_outlet, repere_Q)
    DBO = nominal_flux_kgj(flux_outlet, repere_DBO)
    interm_DCO_conso_bio = interm_DCO_conso_bio - DBO
    'DCO = nominal_flux_kgj(flux_outlet, repere_DCO) - interm_DCO_conso_bio
    DCO_dure_mgL = DCO_nominal * ratio_DCO_dure_total / Q_nominal * 1000
    DCO = DBO + (DCO_dure_mgL + 1.45 * sortie_MES * boues_MV_MES) * Q / 1000
    MES = sortie_MES * nominal_flux_kgj(flux_outlet, repere_Q) / 1000
    NK = nominal_flux_kgj(flux_outlet, repere_NK)
    NH4 = nominal_flux_kgj(flux_outlet, repere_NH4)
    NO3 = nominal_flux_kgj(flux_outlet, repere_NO3)
    Pt = nominal_flux_kgj(flux_outlet, repere_Pt)

    'MsgBox (boues_pdtes_aerobie - MES)

    'DCO_conso_bio = DCO_conso_bio + interm_DCO_conso_bio


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Ligne = 1

    Feuil16.Range("MBBR_nominal_outlet_DBO_mgL").Cells(Ligne, Colonne) = nominal_outlet_DBO_mgL
    Feuil16.Range("MBBR_nominal_outlet_NH4_mgL").Cells(Ligne, Colonne) = nominal_outlet_NH4_mgL
    Feuil16.Range("MBBR_nominal_outlet_NO3_mgL").Cells(Ligne, Colonne) = nominal_outlet_NO3_mgL
    Feuil16.Range("MBBR_Norga_dure_soluble").Cells(Ligne, Colonne) = Norga_dure_soluble
    Feuil16.Range("MBBR_nominal_predenit_NO3_mgL").Cells(Ligne, Colonne) = nominal_predenit_NO3_mgL
    If nominal_recirculation_MLSS_taux <> 0 Then
        Feuil16.Range("MBBR_nominal_recirculation_MLSS_taux").Cells(Ligne, Colonne) = nominal_recirculation_MLSS_taux
    Else
        Feuil16.Range("MBBR_nominal_recirculation_MLSS_taux").Cells(Ligne, Colonne) = "'-"    'nominal_recirculation_MLSS_taux
    End If
    Feuil16.Range("MBBR_ratio_elimination_predenit1_predenit12").Cells(Ligne, Colonne) = ratio_elimination_predenit1_predenit12
    Feuil16.Range("MBBR_ratio_elimination_C1_C12").Cells(Ligne, Colonne) = ratio_elimination_C1_C12
    Feuil16.Range("MBBR_ratio_elimination_N1_N12").Cells(Ligne, Colonne) = ratio_elimination_N1_N12
    Feuil16.Range("MBBR_ratio_elimination_postdenit1_postdenit12").Cells(Ligne, Colonne) = ratio_elimination_postdenit1_postdenit12
    If choix_cuve(cuve_postdenit1) = cuve_oui Then
        Feuil16.Range("MBBR_postdenit_carbone_apporte_flux").Cells(Ligne, Colonne) = postdenit_carbone_apporte_flux
        Feuil16.Range("MBBR_nominal_last_aerated_tank_NO3_mgL").Cells(Ligne, Colonne) = nominal_last_aerated_tank_NO3_mgL
    Else
        Feuil16.Range("MBBR_postdenit_carbone_apporte_flux").Cells(Ligne, Colonne) = "'-"    'postdenit_carbone_apporte_flux
        Feuil16.Range("MBBR_nominal_last_aerated_tank_NO3_mgL").Cells(Ligne, Colonne) = "'-"    'nominal_last_aerated_tank_NO3_mgL
    End If
    Feuil16.Range("MBBR_volume_total_bassins").Cells(Ligne, Colonne) = volume_total_bassins
    Feuil16.Range("MBBR_reel_outlet_DBO_mgL").Cells(Ligne, Colonne) = reel_outlet_DBO_mgL
    Feuil16.Range("MBBR_reel_outlet_NH4_mgL").Cells(Ligne, Colonne) = reel_outlet_NH4_mgL
    Feuil16.Range("MBBR_reel_outlet_NO3_mgL").Cells(Ligne, Colonne) = reel_outlet_NO3_mgL
    Feuil16.Range("MBBR_O2_besoin_total").Cells(Ligne, Colonne) = O2_besoin_total
    Feuil16.Range("MBBR_air_Q_Nm3j").Cells(Ligne, Colonne) = air_Q_Nm3j
    Feuil16.Range("MBBR_suraeration_brassage_Nm3j").Cells(Ligne, Colonne) = suraeration_brassage_Nm3j
    Feuil16.Range("MBBR_diffuseur_encrassement").Cells(Ligne, Colonne) = diffuseur_encrassement
    Feuil16.Range("MBBR_air_P_refoulement_moyenne").Cells(Ligne, Colonne) = air_P_refoulement_moyenne
    Feuil16.Range("MBBR_surpresseur_conso_spec").Cells(Ligne, Colonne) = surpresseur_conso_spec
    If reel_recirculation_MLSS_taux <> 0 Then
        Feuil16.Range("MBBR_reel_recirculation_MLSS_taux").Cells(Ligne, Colonne) = reel_recirculation_MLSS_taux
        Feuil16.Range("MBBR_recirculation_MLSS_P_refoulement").Cells(Ligne, Colonne) = recirculation_MLSS_P_refoulement
        Feuil16.Range("MBBR_recirculation_MLSS_pompe_rdt").Cells(Ligne, Colonne) = recirculation_MLSS_pompe_rdt
    Else
        Feuil16.Range("MBBR_reel_recirculation_MLSS_taux").Cells(Ligne, Colonne) = "'-"    'reel_recirculation_MLSS_taux
        Feuil16.Range("MBBR_recirculation_MLSS_P_refoulement").Cells(Ligne, Colonne) = "'-"    'recirculation_MLSS_P_refoulement
        Feuil16.Range("MBBR_recirculation_MLSS_pompe_rdt").Cells(Ligne, Colonne) = "'-"    'recirculation_MLSS_pompe_rdt
    End If
    Feuil16.Range("MBBR_outlet_MES_concentration").Cells(Ligne, Colonne) = sortie_MES
    Feuil16.Range("MBBR_boues_concentration").Cells(Ligne, Colonne) = boues_concentration
    Feuil16.Range("MBBR_boues_MES").Cells(Ligne, Colonne) = boues_MES
    Feuil16.Range("MBBR_boues_MV_MES").Cells(Ligne, Colonne) = boues_MV_MES
    Feuil16.Range("MBBR_boues_Q").Cells(Ligne, Colonne) = boues_Q
    Feuil16.Range("MBBR_extraction_P_refoulement").Cells(Ligne, Colonne) = extraction_P_refoulement
    Feuil16.Range("MBBR_extraction_pompe_rdt").Cells(Ligne, Colonne) = extraction_pompe_rdt
    Feuil16.Range("MBBR_electricite").Cells(Ligne, Colonne) = electricite
    Feuil16.Range("MBBR_electricite_aeration").Cells(Ligne, Colonne) = electricite_aeration


    For Ligne = 1 To nb_cuves_max
        If choix_cuve(Ligne) = cuve_oui Then
            Feuil16.Range("MBBR_cuve_O2_dissous").Cells(Ligne, Colonne) = cuve_O2_dissous(Ligne)
            Feuil16.Range("MBBR_cuve_filling").Cells(Ligne, Colonne) = cuve_filling(Ligne)
            Feuil16.Range("MBBR_cuve_volume").Cells(Ligne, Colonne) = cuve_volume(Ligne)
            Feuil16.Range("MBBR_cuve_ratio_agitation_W_m3").Cells(Ligne, Colonne) = cuve_ratio_agitation_W_m3(Ligne)
            Feuil16.Range("MBBR_cuve_O2_facteur_alfa").Cells(Ligne, Colonne) = cuve_O2_facteur_alfa(Ligne)
            Feuil16.Range("MBBR_cuve_O2_rdt_transfert").Cells(Ligne, Colonne) = cuve_O2_rdt_transfert(Ligne)
            Feuil16.Range("MBBR_cuve_hauteur").Cells(Ligne, Colonne) = cuve_hauteur(Ligne)
            Feuil16.Range("MBBR_cuve_aeree_Vair_brassage_Nmh").Cells(Ligne, Colonne) = cuve_aeree_Vair_agitation_Nmh(Ligne)
            Feuil16.Range("MBBR_cuve_aeree_brassage_air_hj").Cells(Ligne, Colonne) = cuve_aeree_brassage_air_hj(Ligne)
            Feuil16.Range("MBBR_cuve_aeree_suraeration_brassage_Nm3j").Cells(Ligne, Colonne) = cuve_suraeration_brassage_Nm3j(Ligne)
        Else
            Feuil16.Range("MBBR_cuve_O2_dissous").Cells(Ligne, Colonne) = "'-"    'cuve_O2_dissous(ligne)
            Feuil16.Range("MBBR_cuve_filling").Cells(Ligne, Colonne) = "'-"    'cuve_filling(ligne)
            Feuil16.Range("MBBR_cuve_volume").Cells(Ligne, Colonne) = "'-"    'cuve_volume(ligne)
            Feuil16.Range("MBBR_cuve_ratio_agitation_W_m3").Cells(Ligne, Colonne) = "'-"    'cuve_ratio_agitation_W_m3(ligne)
            Feuil16.Range("MBBR_cuve_O2_facteur_alfa").Cells(Ligne, Colonne) = "'-"    'cuve_O2_facteur_alfa(ligne)
            Feuil16.Range("MBBR_cuve_O2_rdt_transfert").Cells(Ligne, Colonne) = "'-"    'cuve_O2_rdt_transfert(ligne)
            Feuil16.Range("MBBR_cuve_hauteur").Cells(Ligne, Colonne) = "'-"    'cuve_hauteur(ligne)
            Feuil16.Range("MBBR_cuve_aeree_Vair_brassage_Nmh").Cells(Ligne, Colonne) = "'-"    'cuve_aeree_Vair_agitation_Nmh(ligne)
            Feuil16.Range("MBBR_cuve_aeree_brassage_air_hj").Cells(Ligne, Colonne) = "'-"    'cuve_aeree_brassage_air_hj(ligne)
            Feuil16.Range("MBBR_cuve_aeree_suraeration_brassage_Nm3j").Cells(Ligne, Colonne) = ""    'cuve_suraeration_brassage_Nm3j(ligne)
        End If
    Next Ligne

    'Spécifiques au Mox
    If choix_fct_Mox = True Then
        Feuil57.Range("rate_N_T_design").Cells(1, 3 + pi_FiliereConsideree) = rate_N_T_design
        Feuil57.Range("rate_N_T_reel").Cells(1, 3 + pi_FiliereConsideree) = rate_N_T_reel
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim Ligne As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Ligne = 1

    Feuil16.Range("MBBR_nominal_outlet_DBO_mgL").Cells(Ligne, Colonne) = ""    'nominal_outlet_DBO_mgL
    Feuil16.Range("MBBR_nominal_outlet_NH4_mgL").Cells(Ligne, Colonne) = ""    'nominal_outlet_NH4_mgL
    Feuil16.Range("MBBR_nominal_outlet_NO3_mgL").Cells(Ligne, Colonne) = ""    'nominal_outlet_NO3_mgL
    Feuil16.Range("MBBR_Norga_dure_soluble").Cells(Ligne, Colonne) = ""    'Norga_dure_soluble
    Feuil16.Range("MBBR_nominal_predenit_NO3_mgL").Cells(Ligne, Colonne) = ""    'nominal_predenit_NO3_mgL
    Feuil16.Range("MBBR_nominal_last_aerated_tank_NO3_mgL").Cells(Ligne, Colonne) = ""    'nominal_last_aerated_tank_NO3_mgL
    Feuil16.Range("MBBR_nominal_recirculation_MLSS_taux").Cells(Ligne, Colonne) = ""    'nominal_recirculation_MLSS_taux
    Feuil16.Range("MBBR_ratio_elimination_predenit1_predenit12").Cells(Ligne, Colonne) = ""    'ratio_elimination_predenit1_predenit12
    Feuil16.Range("MBBR_ratio_elimination_C1_C12").Cells(Ligne, Colonne) = ""    'ratio_elimination_C1_C12
    Feuil16.Range("MBBR_ratio_elimination_N1_N12").Cells(Ligne, Colonne) = ""    'ratio_elimination_N1_N12
    Feuil16.Range("MBBR_ratio_elimination_postdenit1_postdenit12").Cells(Ligne, Colonne) = ""    'ratio_elimination_postdenit1_postdenit12
    Feuil16.Range("MBBR_postdenit_carbone_apporte_flux").Cells(Ligne, Colonne) = ""    'postdenit_carbone_apporte_flux
    Feuil16.Range("MBBR_volume_total_bassins").Cells(Ligne, Colonne) = ""    'volume_total_bassins
    Feuil16.Range("MBBR_reel_recirculation_MLSS_taux").Cells(Ligne, Colonne) = ""    'reel_recirculation_MLSS_taux
    Feuil16.Range("MBBR_reel_outlet_DBO_mgL").Cells(Ligne, Colonne) = ""    'reel_outlet_DBO_mgL
    Feuil16.Range("MBBR_reel_outlet_NH4_mgL").Cells(Ligne, Colonne) = ""    'reel_outlet_NH4_mgL
    Feuil16.Range("MBBR_reel_outlet_NO3_mgL").Cells(Ligne, Colonne) = ""    'reel_outlet_NO3_mgL
    Feuil16.Range("MBBR_O2_besoin_total").Cells(Ligne, Colonne) = ""    'O2_besoin_total
    Feuil16.Range("MBBR_air_Q_Nm3j").Cells(Ligne, Colonne) = ""    'air_Q_Nm3j
    Feuil16.Range("MBBR_suraeration_brassage_Nm3j").Cells(Ligne, Colonne) = ""    'suraeration_brassage_Nm3j
    Feuil16.Range("MBBR_diffuseur_encrassement").Cells(Ligne, Colonne) = ""    'diffuseur_encrassement
    Feuil16.Range("MBBR_air_P_refoulement_moyenne").Cells(Ligne, Colonne) = ""    'air_P_refoulement_moyenne
    Feuil16.Range("MBBR_surpresseur_conso_spec").Cells(Ligne, Colonne) = ""    'surpresseur_conso_spec
    Feuil16.Range("MBBR_recirculation_MLSS_P_refoulement").Cells(Ligne, Colonne) = ""    'recirculation_MLSS_P_refoulement
    Feuil16.Range("MBBR_recirculation_MLSS_pompe_rdt").Cells(Ligne, Colonne) = ""    'recirculation_MLSS_pompe_rdt
    Feuil16.Range("MBBR_outlet_MES_concentration").Cells(Ligne, Colonne) = ""    'sortie_MES
    Feuil16.Range("MBBR_boues_concentration").Cells(Ligne, Colonne) = ""    'boues_concentration
    Feuil16.Range("MBBR_boues_MES").Cells(Ligne, Colonne) = ""    'boues_MES
    Feuil16.Range("MBBR_boues_MV_MES").Cells(Ligne, Colonne) = ""    'boues_MV_MES
    Feuil16.Range("MBBR_boues_Q").Cells(Ligne, Colonne) = ""    'boues_Q
    Feuil16.Range("MBBR_extraction_P_refoulement").Cells(Ligne, Colonne) = ""    'extraction_P_refoulement
    Feuil16.Range("MBBR_extraction_pompe_rdt").Cells(Ligne, Colonne) = ""    'extraction_pompe_rdt
    Feuil16.Range("MBBR_electricite").Cells(Ligne, Colonne) = ""    'electricite
    Feuil16.Range("MBBR_electricite_aeration").Cells(Ligne, Colonne) = ""    'electricite_aeration

    For Ligne = 1 To nb_cuves_max
        Feuil16.Range("MBBR_cuve_O2_dissous").Cells(Ligne, Colonne) = ""    'cuve_O2_dissous(ligne)
        Feuil16.Range("MBBR_cuve_filling").Cells(Ligne, Colonne) = ""    'cuve_filling(ligne)
        Feuil16.Range("MBBR_cuve_volume").Cells(Ligne, Colonne) = ""    'cuve_volume(ligne)
        Feuil16.Range("MBBR_cuve_ratio_agitation_W_m3").Cells(Ligne, Colonne) = ""    'cuve_ratio_agitation_W_m3(ligne)
        Feuil16.Range("MBBR_cuve_O2_facteur_alfa").Cells(Ligne, Colonne) = ""    'cuve_O2_facteur_alfa(ligne)
        Feuil16.Range("MBBR_cuve_O2_rdt_transfert").Cells(Ligne, Colonne) = ""    'cuve_O2_rdt_transfert(ligne)
        Feuil16.Range("MBBR_cuve_hauteur").Cells(Ligne, Colonne) = ""    'cuve_hauteur(ligne)
        Feuil16.Range("MBBR_cuve_aeree_Vair_brassage_Nmh").Cells(Ligne, Colonne) = ""    'cuve_aeree_Vair_agitation_Nmh(ligne)
        Feuil16.Range("MBBR_cuve_aeree_brassage_air_hj").Cells(Ligne, Colonne) = ""    'cuve_aeree_brassage_air_hj(ligne)
        Feuil16.Range("MBBR_cuve_aeree_suraeration_brassage_Nm3j").Cells(Ligne, Colonne) = ""    'cuve_suraeration_brassage_Nm3j(ligne)
    Next Ligne

    'Spécifiques au Mox
    Feuil57.Range("rate_N_T_design").Cells(1, 3 + pi_FiliereConsideree) = ""
    Feuil57.Range("rate_N_T_reel").Cells(1, 3 + pi_FiliereConsideree) = ""


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub fonctionnement_reel(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef Sh, ByRef DCO_apportee_methanol, ByRef DCO_conso_bio)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer

    Dim N_nit_total As Double   'charge à nitrifier au total
    Dim predenit_denit As Double
    Dim DBOsol_EB As Double
    Dim repere_flux_in As Integer
    Dim interm_NO3_out_mgL As Double
    Dim interm_O2_mgL As Double
    Dim deox_nit As Double
    Dim epsilone_recirculation_mlss As Double
    Dim epsilone_DBO As Double
    Dim epsilone_NH4 As Double
    Dim epsilone_NO3 As Double
    Dim cuve_nit As Double
    Dim interm_DBO_eliminee As Double
    Dim rate_DBOin_appliquee_10°C As Double
    Dim DBO_eliminee_C_C As Double
    Dim cuve_C_C_media_total_m2 As Double
    Dim interm_surface_media As Double
    Dim postdenit_Nassimile As Double
    Dim reox_Nassimile As Double
    Dim cuve_NH4_removed As Double
    Dim cuve_NH4_outlet_mgL As Double
    Dim postdenit_denit As Double
    Dim interm_cuve_NO3_mgL As Double
    Dim DBO_apportee_postdenit_kgj As Double
    Dim outlet_DBO_soluble_mgL As Double
    Dim cuve_O2_besoin As Double
    Dim cuve_air_Q_Nm3j As Double
    Dim interm_DCO_conso_bio As Double

    Dim boues_pdtes(nb_cuves_max) As Double
    Dim boues_pdtes_aerobie As Double


    Dim delta_O2_besoins As Double

    'on calcule les MV/MES de l'influent
    boues_MV_MES = MV_MES_influent(MES, DCO, DBO) * MES

    'on attribue les valeurs des flux en entrée
    reel_flux_kgj(flux_inlet, repere_Q) = Q
    reel_flux_kgj(flux_inlet, repere_DCO) = DCO
    reel_flux_kgj(flux_inlet, repere_DBO) = DBO
    reel_flux_kgj(flux_inlet, repere_MES) = MES
    reel_flux_kgj(flux_inlet, repere_NK) = NK
    reel_flux_kgj(flux_inlet, repere_NH4) = NH4 + EB_rate_Norgasol_Norga * (NK - NH4) + (1 - EB_rate_Norgasol_Norga) * rate_Norgapart_hydrolyse * (NK - NH4) - Norga_dure_soluble * Q / 1000
    reel_flux_kgj(flux_inlet, repere_NO3) = NO3
    reel_flux_kgj(flux_inlet, repere_Pt) = Pt
    reel_flux_kgj(flux_inlet, repere_O2) = 0

    interm_DCO_conso_bio = DBO


    'calcul de la recirculation de MLSS
    If reel_recirculation_MLSS_taux_force = True And choix_cuve(cuve_predenit1) = cuve_oui Then
        reel_recirculation_MLSS_taux = reel_recirculation_MLSS_taux_f
    Else
        reel_recirculation_MLSS_taux = nominal_recirculation_MLSS_taux
    End If


    'ON RENTRE DANS LE DUR!!!!
    epsilone_recirculation_mlss = 1 + critere_convergence

    'initialisation du flux recirculé
    If choix_cuve(cuve_predenit1) = cuve_oui Then
        reel_flux_kgj(flux_recirculation_MLSS, repere_Q) = reel_recirculation_MLSS_taux * Q
        reel_flux_kgj(flux_recirculation_MLSS, repere_DBO) = nominal_outlet_DBO_mgL * reel_flux_kgj(flux_recirculation_MLSS, repere_Q) / 1000
        reel_flux_kgj(flux_recirculation_MLSS, repere_DCO) = reel_flux_kgj(flux_recirculation_MLSS, repere_DBO)
        'reel_flux_kgj(flux_recirculation_MLSS, repere_MES) = MES
        reel_flux_kgj(flux_recirculation_MLSS, repere_NK) = (nominal_outlet_NH4_mgL + Norga_dure_soluble) * reel_flux_kgj(flux_recirculation_MLSS, repere_Q) / 1000
        reel_flux_kgj(flux_recirculation_MLSS, repere_NH4) = nominal_outlet_NH4_mgL * reel_flux_kgj(flux_recirculation_MLSS, repere_Q) / 1000
        reel_flux_kgj(flux_recirculation_MLSS, repere_NO3) = nominal_last_aerated_tank_NO3_mgL * reel_flux_kgj(flux_recirculation_MLSS, repere_Q) / 1000
        reel_flux_kgj(flux_recirculation_MLSS, repere_Pt) = Pt / Q * reel_flux_kgj(flux_recirculation_MLSS, repere_Q)
        reel_flux_kgj(flux_recirculation_MLSS, repere_O2) = 0
    End If

    Do While epsilone_recirculation_mlss > critere_convergence

        epsilone_DBO = reel_flux_kgj(flux_recirculation_MLSS, repere_DBO)
        epsilone_NH4 = reel_flux_kgj(flux_recirculation_MLSS, repere_NH4)
        epsilone_NO3 = reel_flux_kgj(flux_recirculation_MLSS, repere_NO3)

        O2_besoin_total = 0
        air_Q_Nm3j = 0

        'Calcul du flux après recirculation
        For j = 1 To nb_parametres_flux
            reel_flux_kgj(flux_inlet_recirculation, j) = reel_flux_kgj(flux_inlet, j) + reel_flux_kgj(flux_recirculation_MLSS, j)
        Next j


        'Pré-DN
        interm_DBO_eliminee = 0
        For I = cuve_predenit1 To cuve_predenit2
            repere_flux_in = flux_inlet_recirculation + (I - cuve_predenit1)  'repère le flux en entrée
            If choix_cuve(I) = cuve_oui Then
                DBOsol_EB = reel_flux_kgj(flux_inlet, repere_DBO) * (EB_rate_DBOsol_DBO + (1 - EB_rate_DBOsol_DBO) * rate_DBOpart_hydrolyse) + reel_flux_kgj(flux_recirculation_MLSS, repere_DBO) - interm_DBO_eliminee
                Call simule_predenit(I, repere_flux_in, DBOsol_EB, interm_NO3_out_mgL)
                predenit_denit = reel_flux_kgj(repere_flux_in, repere_NO3) + ratio_NO3eq_O2 * reel_flux_kgj(repere_flux_in, repere_O2) - interm_NO3_out_mgL * reel_flux_kgj(repere_flux_in, repere_Q) / 1000
                interm_DBO_eliminee = -ratio_O2_denit * predenit_denit
                'charges en sortie
                reel_flux_kgj(repere_flux_in + 1, repere_Q) = reel_flux_kgj(repere_flux_in, repere_Q)
                reel_flux_kgj(repere_flux_in + 1, repere_NO3) = reel_flux_kgj(repere_flux_in, repere_NO3) + ratio_NO3eq_O2 * reel_flux_kgj(repere_flux_in, repere_O2) - predenit_denit
                reel_flux_kgj(repere_flux_in + 1, repere_DCO) = reel_flux_kgj(repere_flux_in, repere_DCO) - interm_DBO_eliminee
                reel_flux_kgj(repere_flux_in + 1, repere_DBO) = reel_flux_kgj(repere_flux_in, repere_DBO) - interm_DBO_eliminee
                reel_flux_kgj(repere_flux_in + 1, repere_NK) = reel_flux_kgj(repere_flux_in, repere_NK) - ratio_N_assimile * interm_DBO_eliminee
                reel_flux_kgj(repere_flux_in + 1, repere_NH4) = reel_flux_kgj(repere_flux_in, repere_NH4) - ratio_N_assimile * interm_DBO_eliminee
                reel_flux_kgj(repere_flux_in + 1, repere_Pt) = reel_flux_kgj(repere_flux_in, repere_Pt)
                reel_flux_kgj(repere_flux_in + 1, repere_O2) = 0
            Else
                For j = 1 To nb_parametres_flux
                    reel_flux_kgj(repere_flux_in + 1, j) = reel_flux_kgj(repere_flux_in, j)
                Next j
            End If
        Next I

        'Carbone
        If choix_cuve(cuve_C1) = cuve_oui And nb_etapes_traitement = 1 Then
            I = cuve_C2  'on va l'utiliser pour savoir qu'on ne traite que le C.
            repere_flux_in = flux_predenit2_C1    '+ (i - cuve_C1)  'repère le flux en entrée
            Call simule_C(I, repere_flux_in, interm_DBO_eliminee)
            'charges en sortie de C1
            I = cuve_C1
            reel_flux_kgj(repere_flux_in + 1, repere_Q) = reel_flux_kgj(repere_flux_in, repere_Q)
            reel_flux_kgj(repere_flux_in + 1, repere_DCO) = reel_flux_kgj(repere_flux_in, repere_DCO) - cuve_ratio_elimination(I) * interm_DBO_eliminee
            reel_flux_kgj(repere_flux_in + 1, repere_DBO) = reel_flux_kgj(repere_flux_in, repere_DBO) - cuve_ratio_elimination(I) * interm_DBO_eliminee
            'nitrification cuve C
            cuve_nit = nitrification_cuve_C(I, reel_flux_kgj(repere_flux_in + 1, repere_DBO) / reel_flux_kgj(repere_flux_in + 1, repere_Q) * 1000, pd_TemperatureEauConditionsReelles_°C)
            reel_flux_kgj(repere_flux_in + 1, repere_NK) = reel_flux_kgj(repere_flux_in, repere_NK) - ratio_N_assimile * cuve_ratio_elimination(I) * interm_DBO_eliminee - cuve_nit
            reel_flux_kgj(repere_flux_in + 1, repere_NH4) = reel_flux_kgj(repere_flux_in, repere_NH4) - ratio_N_assimile * cuve_ratio_elimination(I) * interm_DBO_eliminee - cuve_nit
            reel_flux_kgj(repere_flux_in + 1, repere_NO3) = reel_flux_kgj(repere_flux_in, repere_NO3) + cuve_nit
            reel_flux_kgj(repere_flux_in + 1, repere_Pt) = reel_flux_kgj(repere_flux_in, repere_Pt) - ratio_P_assimile * cuve_ratio_elimination(I) * interm_DBO_eliminee
            reel_flux_kgj(repere_flux_in + 1, repere_O2) = cuve_O2_dissous(I) * reel_flux_kgj(repere_flux_in + 1, repere_Q) / 1000
            boues_pdtes(I) = fct_boues_pdtes(interm_DBO_eliminee, cuve_nit)
            'on calcule l'aération
            Call aeration_cuve(I, repere_flux_in, reel_flux_kgj, pd_TemperatureEauConditionsReelles_°C, cuve_ratio_elimination(I) * interm_DBO_eliminee, cuve_nit, cuve_O2_besoin, cuve_air_Q_Nm3j, Sh)
            O2_besoin_total = O2_besoin_total + cuve_O2_besoin
            air_Q_Nm3j = air_Q_Nm3j + cuve_air_Q_Nm3j
            'charges en sortie de C2
            I = cuve_C2
            repere_flux_in = flux_predenit2_C1 + (I - cuve_C1)  'repère le flux en entrée
            reel_flux_kgj(repere_flux_in + 1, repere_Q) = reel_flux_kgj(repere_flux_in, repere_Q)
            reel_flux_kgj(repere_flux_in + 1, repere_DCO) = reel_flux_kgj(repere_flux_in, repere_DCO) - cuve_ratio_elimination(I) * interm_DBO_eliminee
            reel_flux_kgj(repere_flux_in + 1, repere_DBO) = reel_flux_kgj(repere_flux_in, repere_DBO) - cuve_ratio_elimination(I) * interm_DBO_eliminee
            'nitrification cuve C
            cuve_nit = nitrification_cuve_C(I, reel_flux_kgj(repere_flux_in + 1, repere_DBO) / reel_flux_kgj(repere_flux_in + 1, repere_Q) * 1000, pd_TemperatureEauConditionsReelles_°C)
            reel_flux_kgj(repere_flux_in + 1, repere_NK) = reel_flux_kgj(repere_flux_in, repere_NK) - ratio_N_assimile * cuve_ratio_elimination(I) * interm_DBO_eliminee - cuve_nit
            reel_flux_kgj(repere_flux_in + 1, repere_NH4) = reel_flux_kgj(repere_flux_in, repere_NH4) - ratio_N_assimile * cuve_ratio_elimination(I) * interm_DBO_eliminee - cuve_nit
            reel_flux_kgj(repere_flux_in + 1, repere_NO3) = reel_flux_kgj(repere_flux_in, repere_NO3) + cuve_nit
            reel_flux_kgj(repere_flux_in + 1, repere_Pt) = reel_flux_kgj(repere_flux_in, repere_Pt) - ratio_P_assimile * cuve_ratio_elimination(I) * interm_DBO_eliminee
            reel_flux_kgj(repere_flux_in + 1, repere_O2) = cuve_O2_dissous(I) * reel_flux_kgj(repere_flux_in + 1, repere_Q) / 1000
            boues_pdtes(I) = fct_boues_pdtes(interm_DBO_eliminee, cuve_nit)
            'on calcule l'aération
            Call aeration_cuve(I, repere_flux_in, reel_flux_kgj, pd_TemperatureEauConditionsReelles_°C, cuve_ratio_elimination(I) * interm_DBO_eliminee, cuve_nit, cuve_O2_besoin, cuve_air_Q_Nm3j, Sh)
            O2_besoin_total = O2_besoin_total + cuve_O2_besoin
            air_Q_Nm3j = air_Q_Nm3j + cuve_air_Q_Nm3j
        ElseIf choix_cuve(cuve_C1) = cuve_oui Then
            I = cuve_C1
            repere_flux_in = flux_predenit2_C1 + (I - cuve_C1)  'repère le flux en entrée
            Call simule_C(I, repere_flux_in, interm_DBO_eliminee)
            'charges en sortie de C1
            reel_flux_kgj(repere_flux_in + 1, repere_Q) = reel_flux_kgj(repere_flux_in, repere_Q)
            reel_flux_kgj(repere_flux_in + 1, repere_DCO) = reel_flux_kgj(repere_flux_in, repere_DCO) - interm_DBO_eliminee
            reel_flux_kgj(repere_flux_in + 1, repere_DBO) = reel_flux_kgj(repere_flux_in, repere_DBO) - interm_DBO_eliminee
            'nitrification cuve C
            cuve_nit = nitrification_cuve_C(I, reel_flux_kgj(repere_flux_in + 1, repere_DBO) / reel_flux_kgj(repere_flux_in + 1, repere_Q) * 1000, pd_TemperatureEauConditionsReelles_°C)
            If cuve_nit > (reel_flux_kgj(repere_flux_in, repere_NH4) - ratio_N_assimile * interm_DBO_eliminee) Then
                cuve_nit = (reel_flux_kgj(repere_flux_in, repere_NH4) - ratio_N_assimile * interm_DBO_eliminee)
            End If
            reel_flux_kgj(repere_flux_in + 1, repere_NK) = reel_flux_kgj(repere_flux_in, repere_NK) - ratio_N_assimile * interm_DBO_eliminee - cuve_nit
            reel_flux_kgj(repere_flux_in + 1, repere_NH4) = reel_flux_kgj(repere_flux_in, repere_NH4) - ratio_N_assimile * interm_DBO_eliminee - cuve_nit
            reel_flux_kgj(repere_flux_in + 1, repere_NO3) = reel_flux_kgj(repere_flux_in, repere_NO3) + cuve_nit
            reel_flux_kgj(repere_flux_in + 1, repere_Pt) = reel_flux_kgj(repere_flux_in, repere_Pt) - ratio_P_assimile * interm_DBO_eliminee
            reel_flux_kgj(repere_flux_in + 1, repere_O2) = cuve_O2_dissous(I) * reel_flux_kgj(repere_flux_in + 1, repere_Q) / 1000
            boues_pdtes(I) = fct_boues_pdtes(interm_DBO_eliminee, cuve_nit)
            'on calcule l'aération
            Call aeration_cuve(I, repere_flux_in, reel_flux_kgj, pd_TemperatureEauConditionsReelles_°C, interm_DBO_eliminee, cuve_nit, cuve_O2_besoin, cuve_air_Q_Nm3j, Sh)
            O2_besoin_total = O2_besoin_total + cuve_O2_besoin
            air_Q_Nm3j = air_Q_Nm3j + cuve_air_Q_Nm3j
            'charges au point suivant qui correspondrait à un éventuel 2ème réacteur C
            For j = 1 To nb_parametres_flux
                reel_flux_kgj(repere_flux_in + 2, j) = reel_flux_kgj(repere_flux_in + 1, j)
            Next j
        Else
            repere_flux_in = flux_predenit2_C1
            For j = 1 To nb_parametres_flux
                reel_flux_kgj(repere_flux_in + 1, j) = reel_flux_kgj(repere_flux_in, j)
                reel_flux_kgj(repere_flux_in + 2, j) = reel_flux_kgj(repere_flux_in, j)
            Next j
        End If


        'Nitrification
        For I = cuve_N1 To cuve_N2
            repere_flux_in = flux_C2_N1 + (I - cuve_N1)  'repère le flux en entrée
            If choix_cuve(I) = cuve_oui Then
                'A voir la validité de ce qui suit: on a considéré au design que l'on traite encore un peu de DBO soluble donc on va assimiler du N que l'on va retrancher à la charge en entrée pour faciliter les calculs
                interm_DBO_eliminee = reel_flux_kgj(repere_flux_in, repere_DBO) - hyp_cuve_N_deox_DBOsol_mgL * reel_flux_kgj(repere_flux_in, repere_Q) / 1000
                'reel_flux_kgj(repere_flux_in, repere_DBO) = reel_flux_kgj(repere_flux_in, repere_DBO) - interm_DBO_eliminee
                reel_flux_kgj(repere_flux_in, repere_NH4) = reel_flux_kgj(repere_flux_in, repere_NH4) - ratio_N_assimile * interm_DBO_eliminee
                reel_flux_kgj(repere_flux_in, repere_NK) = reel_flux_kgj(repere_flux_in, repere_NK) - ratio_N_assimile * interm_DBO_eliminee
                'on continue les calculs
                Call simule_N(I, repere_flux_in, cuve_NH4_outlet_mgL)
                cuve_nit = reel_flux_kgj(repere_flux_in, repere_NH4) - cuve_NH4_outlet_mgL * reel_flux_kgj(repere_flux_in, repere_Q) / 1000
                'charges en sortie
                reel_flux_kgj(repere_flux_in + 1, repere_Q) = reel_flux_kgj(repere_flux_in, repere_Q)
                If choix_fct_Mox = True Then
                    If I = cuve_N1 Then
                        reel_flux_kgj(repere_flux_in + 1, repere_NO3) = reel_flux_kgj(repere_flux_in, repere_NO3) + cuve_nit * ratio_NO3f_nit
                    ElseIf I = cuve_N2 Then
                        reel_flux_kgj(repere_flux_in + 1, repere_NO3) = reel_flux_kgj(repere_flux_in, repere_NO3) + cuve_nit
                    End If
                Else
                    reel_flux_kgj(repere_flux_in + 1, repere_NO3) = reel_flux_kgj(repere_flux_in, repere_NO3) + cuve_nit
                End If
                reel_flux_kgj(repere_flux_in + 1, repere_DCO) = reel_flux_kgj(repere_flux_in, repere_DCO) - interm_DBO_eliminee
                reel_flux_kgj(repere_flux_in + 1, repere_DBO) = reel_flux_kgj(repere_flux_in, repere_DBO) - interm_DBO_eliminee
                reel_flux_kgj(repere_flux_in + 1, repere_NK) = reel_flux_kgj(repere_flux_in, repere_NK) - cuve_nit
                reel_flux_kgj(repere_flux_in + 1, repere_NH4) = reel_flux_kgj(repere_flux_in, repere_NH4) - cuve_nit
                reel_flux_kgj(repere_flux_in + 1, repere_Pt) = reel_flux_kgj(repere_flux_in, repere_Pt)
                reel_flux_kgj(repere_flux_in + 1, repere_O2) = cuve_O2_dissous(I) * reel_flux_kgj(repere_flux_in + 1, repere_Q) / 1000
                boues_pdtes(I) = fct_boues_pdtes(interm_DBO_eliminee, cuve_nit)
                'on calcule l'aération
                Call aeration_cuve(I, repere_flux_in, reel_flux_kgj, pd_TemperatureEauConditionsReelles_°C, interm_DBO_eliminee, cuve_nit, cuve_O2_besoin, cuve_air_Q_Nm3j, Sh)
                O2_besoin_total = O2_besoin_total + cuve_O2_besoin
                air_Q_Nm3j = air_Q_Nm3j + cuve_air_Q_Nm3j
            Else
                For j = 1 To nb_parametres_flux
                    reel_flux_kgj(repere_flux_in + 1, j) = reel_flux_kgj(repere_flux_in, j)
                Next j
            End If
        Next I



        'DéOx
        If choix_cuve(cuve_deox) = cuve_oui Then
            'A voir la validité de ce qui suit: on a considéré au design que l'on traite encore un peu de DBO soluble donc on va assimiler du N que l'on va retrancher à la charge en entrée pour faciliter les calculs
            interm_DBO_eliminee = reel_flux_kgj(flux_N2_deox, repere_DBO) - hyp_cuve_N_deox_DBOsol_mgL * reel_flux_kgj(flux_N2_deox, repere_Q) / 1000
            'reel_flux_kgj(flux_N2_deox, repere_DBO) = reel_flux_kgj(flux_N2_deox, repere_DBO) - interm_DBO_eliminee
            reel_flux_kgj(flux_N2_deox, repere_NH4) = reel_flux_kgj(flux_N2_deox, repere_NH4) - ratio_N_assimile * interm_DBO_eliminee
            reel_flux_kgj(flux_N2_deox, repere_NK) = reel_flux_kgj(flux_N2_deox, repere_NK) - ratio_N_assimile * interm_DBO_eliminee
            'on continue les calculs
            Call simule_deox(interm_O2_mgL)
            cuve_nit = (reel_flux_kgj(flux_N2_deox, repere_O2) - interm_O2_mgL * reel_flux_kgj(flux_N2_deox, repere_Q) / 1000) / ratio_O2_nit
            If cuve_nit > reel_flux_kgj(flux_N2_deox, repere_NH4) Then
                cuve_nit = reel_flux_kgj(flux_N2_deox, repere_NH4)
            End If
            'charges en sortie
            reel_flux_kgj(flux_deox_separationMLSS, repere_Q) = reel_flux_kgj(flux_N2_deox, repere_Q)
            reel_flux_kgj(flux_deox_separationMLSS, repere_NO3) = reel_flux_kgj(flux_N2_deox, repere_NO3) + cuve_nit
            reel_flux_kgj(flux_deox_separationMLSS, repere_DCO) = reel_flux_kgj(flux_N2_deox, repere_DCO) - interm_DBO_eliminee
            reel_flux_kgj(flux_deox_separationMLSS, repere_DBO) = reel_flux_kgj(flux_N2_deox, repere_DBO) - interm_DBO_eliminee
            reel_flux_kgj(flux_deox_separationMLSS, repere_NK) = reel_flux_kgj(flux_N2_deox, repere_NK) - cuve_nit
            reel_flux_kgj(flux_deox_separationMLSS, repere_NH4) = reel_flux_kgj(flux_N2_deox, repere_NH4) - cuve_nit
            reel_flux_kgj(flux_deox_separationMLSS, repere_Pt) = reel_flux_kgj(flux_N2_deox, repere_Pt)
            reel_flux_kgj(flux_deox_separationMLSS, repere_O2) = interm_O2_mgL * reel_flux_kgj(flux_deox_separationMLSS, repere_Q) / 1000
            boues_pdtes(cuve_deox) = fct_boues_pdtes(interm_DBO_eliminee, deox_nit)
        Else
            For j = 1 To nb_parametres_flux
                reel_flux_kgj(flux_deox_separationMLSS, j) = reel_flux_kgj(flux_N2_deox, j)
            Next j
        End If


        'Calcul du flux de recirculation
        For j = 1 To nb_parametres_flux
            reel_flux_kgj(flux_recirculation_MLSS, j) = reel_flux_kgj(flux_deox_separationMLSS, j) * reel_recirculation_MLSS_taux / (1 + reel_recirculation_MLSS_taux)
        Next j


        epsilone_DBO = epsilone_DBO - reel_flux_kgj(flux_recirculation_MLSS, repere_DBO)
        epsilone_NH4 = epsilone_NH4 - reel_flux_kgj(flux_recirculation_MLSS, repere_NH4)
        epsilone_NO3 = epsilone_NO3 - reel_flux_kgj(flux_recirculation_MLSS, repere_NO3)

        epsilone_DBO = epsilone_DBO ^ 2
        epsilone_NH4 = epsilone_NH4 ^ 2
        epsilone_NO3 = epsilone_NO3 ^ 2

        epsilone_recirculation_mlss = epsilone_DBO + epsilone_NH4 + epsilone_NO3

    Loop


    'Calcul du flux post séparation de la recirculation
    For j = 1 To nb_parametres_flux
        reel_flux_kgj(flux_separationMLSS_postdenit1, j) = reel_flux_kgj(flux_deox_separationMLSS, j) - reel_flux_kgj(flux_recirculation_MLSS, j)
    Next j

    outlet_DBO_soluble_mgL = reel_flux_kgj(flux_separationMLSS_postdenit1, repere_DBO) / reel_flux_kgj(flux_separationMLSS_postdenit1, repere_Q) * 1000

    'Post-DN
    DBO_apportee_postdenit_kgj = 0
    For I = cuve_postdenit1 To cuve_postdenit2
        repere_flux_in = flux_separationMLSS_postdenit1 + (I - cuve_postdenit1)  'repère le flux en entrée
        If choix_cuve(I) = cuve_oui Then
            Call simule_postdenit(I, repere_flux_in, interm_NO3_out_mgL)
            postdenit_denit = reel_flux_kgj(repere_flux_in, repere_NO3) - interm_NO3_out_mgL * reel_flux_kgj(repere_flux_in, repere_Q) / 1000
            DBO_apportee_postdenit_kgj = DBO_apportee_postdenit_kgj + rate_DBO_apportee_denit_reel * postdenit_denit
            'charges en sortie
            reel_flux_kgj(repere_flux_in + 1, repere_Q) = reel_flux_kgj(repere_flux_in, repere_Q)
            reel_flux_kgj(repere_flux_in + 1, repere_NO3) = reel_flux_kgj(repere_flux_in, repere_NO3) - postdenit_denit
            If I = cuve_postdenit2 Or choix_cuve(cuve_postdenit2) = cuve_non Then
                reel_flux_kgj(repere_flux_in + 1, repere_DCO) = reel_flux_kgj(repere_flux_in, repere_DCO) + ratio_DCO_DBO_source_C * DBO_apportee_fuite_mgL * reel_flux_kgj(repere_flux_in + 1, repere_Q) / 1000
                reel_flux_kgj(repere_flux_in + 1, repere_DBO) = reel_flux_kgj(repere_flux_in, repere_DBO) + DBO_apportee_fuite_mgL * reel_flux_kgj(repere_flux_in + 1, repere_Q) / 1000
            Else
                reel_flux_kgj(repere_flux_in + 1, repere_DCO) = reel_flux_kgj(repere_flux_in, repere_DCO)
                reel_flux_kgj(repere_flux_in + 1, repere_DBO) = reel_flux_kgj(repere_flux_in, repere_DBO)
            End If
            reel_flux_kgj(repere_flux_in + 1, repere_NK) = reel_flux_kgj(repere_flux_in, repere_NK) - ratio_N_assimile_Capporte * ratio_boues_DBO_apportee * rate_DBO_apportee_denit_reel * postdenit_denit
            If reel_flux_kgj(repere_flux_in + 1, repere_NK) < 0 Then
                reel_flux_kgj(repere_flux_in + 1, repere_NK) = 0
            End If
            reel_flux_kgj(repere_flux_in + 1, repere_NH4) = reel_flux_kgj(repere_flux_in, repere_NH4) - ratio_N_assimile_Capporte * ratio_boues_DBO_apportee * rate_DBO_apportee_denit_reel * postdenit_denit
            If reel_flux_kgj(repere_flux_in + 1, repere_NH4) < 0 Then
                reel_flux_kgj(repere_flux_in + 1, repere_NH4) = 0
            End If
            reel_flux_kgj(repere_flux_in + 1, repere_Pt) = reel_flux_kgj(repere_flux_in, repere_Pt)
            reel_flux_kgj(repere_flux_in + 1, repere_O2) = 0
        Else
            For j = 1 To nb_parametres_flux
                reel_flux_kgj(repere_flux_in + 1, j) = reel_flux_kgj(repere_flux_in, j)
            Next j
        End If
    Next I
    postdenit_carbone_apporte_flux = DBO_apportee_postdenit_kgj / rate_DBO_source_C
    boues_postdenit = ratio_boues_DBO_apportee * DBO_apportee_postdenit_kgj
    '********************** rajouter les fuites de DBO pour calculer la RéOx *****************************

    'RéOx
    If choix_cuve(cuve_reox) = cuve_oui Then
        repere_flux_in = flux_postdenit2_reox  'repère le flux en entrée
        'interm_DBO_eliminee
        Call simule_reox(interm_DBO_eliminee)
        If interm_DBO_eliminee > reel_flux_kgj(repere_flux_in, repere_DBO) Then
            interm_DBO_eliminee = reel_flux_kgj(repere_flux_in, repere_DBO)
        End If
        'charges en sortie
        reel_flux_kgj(repere_flux_in + 1, repere_Q) = reel_flux_kgj(repere_flux_in, repere_Q)
        reel_flux_kgj(repere_flux_in + 1, repere_DCO) = reel_flux_kgj(repere_flux_in, repere_DCO) - interm_DBO_eliminee
        reel_flux_kgj(repere_flux_in + 1, repere_DBO) = reel_flux_kgj(repere_flux_in, repere_DBO) - interm_DBO_eliminee
        reel_flux_kgj(repere_flux_in + 1, repere_NK) = reel_flux_kgj(repere_flux_in, repere_NK)    '- ratio_N_assimile * interm_DBO_eliminee
        reel_flux_kgj(repere_flux_in + 1, repere_NH4) = reel_flux_kgj(repere_flux_in, repere_NH4)    '- ratio_N_assimile * interm_DBO_eliminee
        reel_flux_kgj(repere_flux_in + 1, repere_NO3) = reel_flux_kgj(repere_flux_in, repere_NO3)
        reel_flux_kgj(repere_flux_in + 1, repere_Pt) = reel_flux_kgj(repere_flux_in, repere_Pt) - ratio_P_assimile * interm_DBO_eliminee
        reel_flux_kgj(repere_flux_in + 1, repere_O2) = cuve_O2_dissous(I) * reel_flux_kgj(repere_flux_in + 1, repere_Q) / 1000
        'production de boues
        boues_pdtes(cuve_reox) = fct_boues_pdtes(interm_DBO_eliminee, 0)
        'on calcule l'aération
        Call aeration_cuve(cuve_reox, repere_flux_in, reel_flux_kgj, pd_TemperatureEauConditionsReelles_°C, interm_DBO_eliminee, 0, cuve_O2_besoin, cuve_air_Q_Nm3j, Sh)
        O2_besoin_total = O2_besoin_total + cuve_O2_besoin
        air_Q_Nm3j = air_Q_Nm3j + cuve_air_Q_Nm3j
    Else
        For j = 1 To nb_parametres_flux
            reel_flux_kgj(flux_outlet, j) = reel_flux_kgj(flux_postdenit2_reox, j)
        Next j
    End If

    'ici on calcule le soluble
    outlet_DBO_soluble_mgL = reel_flux_kgj(flux_outlet, repere_DBO) / reel_flux_kgj(flux_outlet, repere_Q) * 1000
    ' on considère que le soluble sera le soluble en sortie des cuves aérobies (hyp: non modifié avec réOx)
    If nb_etapes_traitement = 1 Then
        'on en déduit la DBO totale (démarche à vérifier)
        If choix_type_separation = clarif Then
            reel_outlet_DBO_mgL = (outlet_DBO_soluble_mgL + 5) / 0.5
        ElseIf choix_type_separation = clarif_polymere Then
            reel_outlet_DBO_mgL = (outlet_DBO_soluble_mgL + 3) / 0.5275
        Else
            reel_outlet_DBO_mgL = (outlet_DBO_soluble_mgL + 1.1553) / 0.3936
        End If
    Else
        If reel_outlet_DBO_mgL_force = True Then
            reel_outlet_DBO_mgL = reel_outlet_DBO_mgL_f
        Else
            reel_outlet_DBO_mgL = nominal_outlet_DBO_mgL * (outlet_DBO_soluble_mgL / (nominal_flux_kgj(flux_outlet, repere_DBO) / nominal_flux_kgj(flux_outlet, repere_Q) * 1000))
            If reel_outlet_DBO_mgL <= hyp_outlet_DBO_mgL_mini Then
                reel_outlet_DBO_mgL = hyp_outlet_DBO_mgL_mini
            End If
        End If
    End If

    boues_pdtes_aerobie = 0
    For I = 1 To nb_cuves_max
        boues_pdtes_aerobie = boues_pdtes_aerobie + boues_pdtes(I)
    Next I
    boues_MV_MES = boues_MV_MES * (1 - rate_DBOpart_hydrolyse) + boues_pdtes_aerobie + boues_postdenit
    boues_pdtes_aerobie = boues_pdtes_aerobie + reel_flux_kgj(flux_inlet, repere_MES) * (1 - rate_DBOpart_hydrolyse)

    'Séparation liquide/solide
    Call separation_boues(reel_flux_kgj(flux_inlet, repere_MES), reel_flux_kgj(flux_inlet, repere_DBO), reel_flux_kgj(flux_inlet, repere_Q), pd_TemperatureEauConditionsReelles_°C, reel_outlet_DBO_mgL, reel_flux_kgj(flux_outlet, repere_DBO) / reel_flux_kgj(flux_outlet, repere_Q) * 1000, boues_pdtes_aerobie, True)

    If boues_MV_MES_force = True Then
        boues_MV_MES = boues_MV_MES_f
    End If

    'ne pas supprimer car utile par la suite pour être cohérents avec NK/NH4 et DCO/DBO si valeurs forcées
    Q = reel_flux_kgj(flux_outlet, repere_Q)
    DBO = reel_flux_kgj(flux_outlet, repere_DBO)
    DCO = reel_flux_kgj(flux_outlet, repere_DCO)
    MES = sortie_MES * reel_flux_kgj(flux_outlet, repere_Q) / 1000
    NK = reel_flux_kgj(flux_outlet, repere_NK)
    NH4 = reel_flux_kgj(flux_outlet, repere_NH4)
    NO3 = reel_flux_kgj(flux_outlet, repere_NO3)
    Pt = reel_flux_kgj(flux_outlet, repere_Pt)


    reel_outlet_NH4_mgL = reel_flux_kgj(flux_outlet, repere_NH4) / reel_flux_kgj(flux_outlet, repere_Q) * 1000
    reel_outlet_NO3_mgL = reel_flux_kgj(flux_outlet, repere_NO3) / reel_flux_kgj(flux_outlet, repere_Q) * 1000
    'reel_outlet_DBO_mgL = reel_flux_kgj(flux_outlet, repere_DBO) / reel_flux_kgj(flux_outlet, repere_Q) * 1000

    'Comment gérer le forçage des besoins en O2, du débit d'air et surtout des concentrations de sortie? Bonne question. On va faire ça un peu comme on peut parce que c'est quand même plus compliqué qu'une BA
    If reel_outlet_NH4_mgL_force = True Then
        delta_O2_besoins = ratio_O2_nit * (reel_outlet_NH4_mgL - reel_outlet_NH4_mgL_f) * Q / 1000
        If O2_besoin_total_force = True Then
            If air_Q_Nm3j_force = True Then
                O2_besoin_total = O2_besoin_total_f
                air_Q_Nm3j = air_Q_Nm3j_f
            Else
                air_Q_Nm3j = O2_besoin_total_f / O2_besoin_total * air_Q_Nm3j
                O2_besoin_total = O2_besoin_total_f
            End If
        Else
            If air_Q_Nm3j_force = True Then
                air_Q_Nm3j = air_Q_Nm3j_f
                O2_besoin_total = O2_besoin_total + delta_O2_besoins
            Else
                air_Q_Nm3j = (O2_besoin_total + delta_O2_besoins) / O2_besoin_total * air_Q_Nm3j
                O2_besoin_total = O2_besoin_total + delta_O2_besoins
            End If
        End If
        reel_outlet_NH4_mgL = reel_outlet_NH4_mgL_f
    ElseIf O2_besoin_total_force = True Then
        If air_Q_Nm3j_force = True Then
            air_Q_Nm3j = air_Q_Nm3j_f
        Else
            air_Q_Nm3j = O2_besoin_total_f / O2_besoin_total * air_Q_Nm3j
        End If
        O2_besoin_total = O2_besoin_total_f
    ElseIf air_Q_Nm3j_force = True Then
        air_Q_Nm3j = air_Q_Nm3j_f
    End If
    If reel_outlet_NO3_mgL_force = True Then
        reel_outlet_NO3_mgL = reel_outlet_NO3_mgL_f
    End If


    NK = NK - NH4
    DCO = DCO - DBO

    'on attribue les charges de sortie
    Q = reel_flux_kgj(flux_outlet, repere_Q)
    DBO = reel_outlet_DBO_mgL * Q / 1000
    interm_DCO_conso_bio = interm_DCO_conso_bio - DBO
    DCO_dure_mgL = DCO_nominal * pd_PourcentageChargeReelleDCO * ratio_DCO_dure_total / (Q_nominal * pd_PourcentageChargeReelleDebitVolumique) * 1000
    DCO = DBO + (DCO_dure_mgL + 1.45 * sortie_MES * boues_MV_MES) * Q / 1000
    'formule HybAS (outlet_reel_DBO + DCO_dure_mgL + 1.45 * sortie_MES * reel_MV_MES) * Q / 1000    'A VOIR SI COHERENT   + 1.45 * sortie_MES * reel_MV_MES
    MES = sortie_MES * reel_flux_kgj(flux_outlet, repere_Q) / 1000
    NH4 = reel_outlet_NH4_mgL * Q / 1000
    NK = NK + NH4   'pour être sûr d'être cohérents
    NO3 = reel_outlet_NO3_mgL * Q / 1000
    Pt = reel_flux_kgj(flux_outlet, repere_Pt)


    'DCO_conso_bio = DCO_conso_bio + interm_DCO_conso_bio

    '**************** POUR VERIFIER LA COHERENCE DES FLUX A CHAQUE ETAPE ***************************
    'If iteration = pi_NOMBRE_ITERATION_RETOURS Then
    '    For i = flux_inlet To flux_outlet
    '        For j = repere_Q To repere_O2
    '            MsgBox (i & " - " & j & " : " & nominal_flux_kgj(i, j) - reel_flux_kgj(i, j))
    '        Next j
    '    Next i
    'End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub



Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'on attribue les valeurs guides par défaut
    Dim I As Integer
    Dim hauteur_bassin_max As Double
    Dim rdt_pompe As Double
    Dim rdt_moteur_pompe As Double


    'concentration en NH4 en sortie
    If nominal_outlet_NH4_mgL_force = True Then
        nominal_outlet_NH4_mgL = nominal_outlet_NH4_mgL_f
    Else
        nominal_outlet_NH4_mgL = 3
    End If

    'concentration en NO3 en sortie
    If nominal_outlet_NO3_mgL_force = True Then
        nominal_outlet_NO3_mgL = nominal_outlet_NO3_mgL_f
    Else
        If choix_cuve(cuve_postdenit1) = cuve_oui Then
            nominal_outlet_NO3_mgL = 5
        Else
            nominal_outlet_NO3_mgL = 15
        End If
    End If

    'concentration en DBO en sortie
    If nominal_outlet_DBO_mgL_force = True Then
        nominal_outlet_DBO_mgL = nominal_outlet_DBO_mgL_f
    Else
        nominal_outlet_DBO_mgL = 20   'à voir si on le relie aux garanties de sortie pi_FiliereConsideree
    End If

    'N orga dur et N orga soluble (talon de NK non éliminé hors NH4)
    If Norga_dure_soluble_force = True Then
        Norga_dure_soluble = Norga_dure_soluble_f
    Else
        Norga_dure_soluble = 1.5
    End If

    'concentration en NO3 en amont de la préDN
    If choix_cuve(cuve_predenit1) = cuve_oui Then
        If choix_cuve(cuve_postdenit1) = cuve_oui Then
            If nominal_last_aerated_tank_NO3_mgL_force = True Then
                nominal_last_aerated_tank_NO3_mgL = nominal_last_aerated_tank_NO3_mgL_f
            Else
                nominal_last_aerated_tank_NO3_mgL = nominal_outlet_NO3_mgL + delta_NO3_post_DN
            End If
        Else
            nominal_last_aerated_tank_NO3_mgL = nominal_outlet_NO3_mgL
        End If
    End If

    'concentration en NO3 après la Pré-DN
    If choix_cuve(cuve_predenit1) = cuve_oui Then
        If nominal_predenit_NO3_mgL_force = True Then
            nominal_predenit_NO3_mgL = nominal_predenit_NO3_mgL_f
        Else
            nominal_predenit_NO3_mgL = 3
        End If
    End If


    '******************************************** ON REGARDE LE NOMBRE DE CUVES QUE L'ON A BESOIN POUR CHAQUE APPLICATION SELON LES VALEURS ATTRIBUEES CI-DESSUS ********************************
    'PréDN
    If choix_cuve(cuve_predenit1) = cuve_oui Then
        If ratio_elimination_predenit1_predenit12_force = True Then
            ratio_elimination_predenit1_predenit12 = ratio_elimination_predenit1_predenit12_f
            'si ça vaut 100%, une seule cuve, sinon 2
            If ratio_elimination_predenit1_predenit12 = 1 Then
                choix_cuve(cuve_predenit2) = cuve_non
                'on vérifie qu'il n'y a pas de conflit avec une éventuelle valeur de volume renseignée
                If cuve_volume_force(cuve_predenit2) = True And cuve_volume_f(cuve_predenit2) > critere_volume_nul Then
                    '************************ message pour incohérence des valeurs forcées **********************************
                End If
            Else
                choix_cuve(cuve_predenit2) = cuve_oui
                'on vérifie qu'il n'y a pas de conflit avec une éventuelle valeur de volume renseignée
                If cuve_volume_force(cuve_predenit2) = True And cuve_volume_f(cuve_predenit2) <= critere_volume_nul Then
                    '************************ message pour incohérence des valeurs forcées **********************************
                End If
            End If
        Else
            If cuve_volume_force(cuve_predenit2) = True Then
                If cuve_volume_f(cuve_predenit2) <= critere_volume_nul Then   'le volume est forcé à 0 pour la cuve 2 --> a priori si l'utilisateur veut mettre une seule cuve si nous on lui en impose 2, il va pas fixer la 1 à 0 sinon c'est qu'il est sacrément tordu!
                    choix_cuve(cuve_predenit2) = cuve_non
                Else
                    choix_cuve(cuve_predenit2) = cuve_oui
                End If
            Else
                If nominal_predenit_NO3_mgL < predenit_NO3_limite Then
                    choix_cuve(cuve_predenit2) = cuve_oui
                Else
                    choix_cuve(cuve_predenit2) = cuve_non
                End If
            End If
            'on effectue la répartition des charges à traiter
            If choix_cuve(cuve_predenit2) = cuve_oui Then
                ratio_elimination_predenit1_predenit12 = 0.5
            Else
                ratio_elimination_predenit1_predenit12 = 1
            End If
        End If
        cuve_ratio_elimination(cuve_predenit1) = ratio_elimination_predenit1_predenit12
        cuve_ratio_elimination(cuve_predenit2) = 1 - cuve_ratio_elimination(cuve_predenit1)
    End If
    'On vérifie qu'il mette le nombre de cuves conseillé pour l'alerter si ce n'est pas le cas.
    If choix_cuve(cuve_predenit2) = cuve_oui And nominal_predenit_NO3_mgL > predenit_NO3_limite Then  'en théorie on devrait mettre >= mais bon c'est un cas intermédiaire où l'utilisateur peut préférer en mettre 2 pour assurer
        'une cuve devrait suffire
    ElseIf choix_cuve(cuve_predenit2) = cuve_non And nominal_predenit_NO3_mgL < predenit_NO3_limite Then
        'deux cuves conseillées
    End If


    'Etape C
    If choix_cuve(cuve_C1) = cuve_oui And nb_etapes_traitement = 1 Then
        If ratio_elimination_C1_C12_force = True Then
            ratio_elimination_C1_C12 = ratio_elimination_C1_C12_f
            'si ça vaut 100%, une seule cuve, sinon 2
            If ratio_elimination_C1_C12 = 1 Then
                choix_cuve(cuve_C2) = cuve_non
                'on vérifie qu'il n'y a pas de conflit avec une éventuelle valeur de volume renseignée
                If cuve_volume_force(cuve_C2) = True And cuve_volume_f(cuve_C2) > critere_volume_nul Then
                    '************************ message pour incohérence des valeurs forcées **********************************
                End If
            Else
                choix_cuve(cuve_C2) = cuve_oui
                'on vérifie qu'il n'y a pas de conflit avec une éventuelle valeur de volume renseignée
                If cuve_volume_force(cuve_C2) = True And cuve_volume_f(cuve_C2) <= critere_volume_nul Then
                    '************************ message pour incohérence des valeurs forcées **********************************
                End If
            End If
        Else
            If cuve_volume_force(cuve_C2) = True Then
                If cuve_volume_f(cuve_C2) <= critere_volume_nul Then   'le volume est forcé à 0 pour la cuve 2 --> a priori si l'utilisateur veut mettre une seule cuve si nous on lui en impose 2, il va pas fixer la 1 à 0 sinon c'est qu'il est sacrément tordu!
                    choix_cuve(cuve_C2) = cuve_non
                Else
                    choix_cuve(cuve_C2) = cuve_oui
                End If
            Else
                choix_cuve(cuve_C2) = cuve_oui
            End If
            'on effectue la répartition des charges à traiter
            If choix_cuve(cuve_C2) = cuve_oui Then
                ratio_elimination_C1_C12 = 0.7
            Else
                ratio_elimination_C1_C12 = 1
            End If
        End If
        cuve_ratio_elimination(cuve_C1) = ratio_elimination_C1_C12
        cuve_ratio_elimination(cuve_C2) = 1 - cuve_ratio_elimination(cuve_C1)
    Else
        choix_cuve(cuve_C2) = cuve_non
    End If
    'On vérifie qu'il mette le nombre de cuves conseillé pour l'alerter si ce n'est pas le cas.
    If choix_cuve(cuve_C2) = cuve_non And nb_etapes_traitement = 1 Then
        'deux cuves conseillées
    End If

    'Etape N
    If choix_cuve(cuve_N1) = cuve_oui Then
        If ratio_elimination_N1_N12_force = True Then
            ratio_elimination_N1_N12 = ratio_elimination_N1_N12_f
            'si ça vaut 100%, une seule cuve, sinon 2
            If ratio_elimination_N1_N12 = 1 Then
                choix_cuve(cuve_N2) = cuve_non
                'on vérifie qu'il n'y a pas de conflit avec une éventuelle valeur de volume renseignée
                If cuve_volume_force(cuve_N2) = True And cuve_volume_f(cuve_N2) > critere_volume_nul Then
                    '************************ message pour incohérence des valeurs forcées **********************************
                End If
            Else
                choix_cuve(cuve_N2) = cuve_oui
                'on vérifie qu'il n'y a pas de conflit avec une éventuelle valeur de volume renseignée
                If cuve_volume_force(cuve_N2) = True And cuve_volume_f(cuve_N2) <= critere_volume_nul Then
                    '************************ message pour incohérence des valeurs forcées **********************************
                End If
            End If
        Else
            If cuve_volume_force(cuve_N2) = True Then
                If cuve_volume_f(cuve_N2) <= critere_volume_nul Then   'le volume est forcé à 0 pour la cuve 2 --> a priori si l'utilisateur veut mettre une seule cuve si nous on lui en impose 2, il va pas fixer la 1 à 0 sinon c'est qu'il est sacrément tordu!
                    choix_cuve(cuve_N2) = cuve_non
                Else
                    choix_cuve(cuve_N2) = cuve_oui
                End If
            Else
                choix_cuve(cuve_N2) = cuve_oui
            End If
            'on effectue la répartition des charges à traiter
            If choix_cuve(cuve_N2) = cuve_oui Then
                ratio_elimination_N1_N12 = 0.5
            Else
                ratio_elimination_N1_N12 = 1
            End If
        End If
        cuve_ratio_elimination(cuve_N1) = ratio_elimination_N1_N12
        cuve_ratio_elimination(cuve_N2) = 1 - cuve_ratio_elimination(cuve_N1)
    End If

    'Post-DN
    If choix_cuve(cuve_postdenit1) = cuve_oui Then
        If ratio_elimination_postdenit1_postdenit12_force = True Then
            ratio_elimination_postdenit1_postdenit12 = ratio_elimination_postdenit1_postdenit12_f
            'si ça vaut 100%, une seule cuve, sinon 2
            If ratio_elimination_postdenit1_postdenit12 = 1 Then
                choix_cuve(cuve_postdenit2) = cuve_non
                'on vérifie qu'il n'y a pas de conflit avec une éventuelle valeur de volume renseignée
                If cuve_volume_force(cuve_postdenit2) = True And cuve_volume_f(cuve_postdenit2) > critere_volume_nul Then
                    '************************ message pour incohérence des valeurs forcées **********************************
                End If
            Else
                choix_cuve(cuve_postdenit2) = cuve_oui
                'on vérifie qu'il n'y a pas de conflit avec une éventuelle valeur de volume renseignée
                If cuve_volume_force(cuve_postdenit2) = True And cuve_volume_f(cuve_postdenit2) <= critere_volume_nul Then
                    '************************ message pour incohérence des valeurs forcées **********************************
                End If
            End If
        Else
            If cuve_volume_force(cuve_postdenit2) = True Then
                If cuve_volume_f(cuve_postdenit2) <= critere_volume_nul Then   'le volume est forcé à 0 pour la cuve 2 --> a priori si l'utilisateur veut mettre une seule cuve si nous on lui en impose 2, il va pas fixer la 1 à 0 sinon c'est qu'il est sacrément tordu!
                    choix_cuve(cuve_postdenit2) = cuve_non
                Else
                    choix_cuve(cuve_postdenit2) = cuve_oui
                End If
            Else
                If nominal_outlet_NO3_mgL < postdenit_NO3_limite Then
                    choix_cuve(cuve_postdenit2) = cuve_oui
                Else
                    choix_cuve(cuve_postdenit2) = cuve_non
                End If
            End If
            'on effectue la répartition des charges à traiter
            If choix_cuve(cuve_postdenit2) = cuve_oui Then
                ratio_elimination_postdenit1_postdenit12 = 0.5
            Else
                ratio_elimination_postdenit1_postdenit12 = 1
            End If
        End If
        cuve_ratio_elimination(cuve_postdenit1) = ratio_elimination_postdenit1_postdenit12
        cuve_ratio_elimination(cuve_postdenit2) = 1 - cuve_ratio_elimination(cuve_postdenit1)
    End If
    'On vérifie qu'il mette le nombre de cuves conseillé pour l'alerter si ce n'est pas le cas.
    If choix_cuve(cuve_postdenit2) = cuve_oui And nominal_outlet_NO3_mgL > postdenit_NO3_limite Then  'en théorie on devrait mettre >= mais bon c'est un cas intermédiaire où l'utilisateur peut préférer en mettre 2 pour assurer
        'une cuve devrait suffire
    ElseIf choix_cuve(cuve_postdenit2) = cuve_non And nominal_outlet_NO3_mgL < postdenit_NO3_limite Then
        'deux cuves conseillées
    End If

    '*************************************************************** FIN DE LA DETERMINATION DES CUVES ********************************************************************


    'concentration des boues
    If boues_concentration_force = True Then    'g/L     voir si on calcule en fonction du type de séparation
        boues_concentration = boues_concentration_f
    Else
        boues_concentration = 20
    End If


    'DO des cuves aérobies
    hauteur_bassin_max = 0
    'cuves C
    For I = cuve_C1 To cuve_C2
        If choix_cuve(I) = cuve_oui Then
            If cuve_O2_dissous_force(I) = True Then
                cuve_O2_dissous(I) = cuve_O2_dissous_f(I)
            Else
                If pd_TemperatureEauDimensionnement_°C <= 11 Then
                    cuve_O2_dissous(I) = 3 + md_ConcentrationOxygeneDissousExcesRegulation_mgL(choix_regulation)
                ElseIf pd_TemperatureEauDimensionnement_°C < 16 Then
                    cuve_O2_dissous(I) = 2.5 + md_ConcentrationOxygeneDissousExcesRegulation_mgL(choix_regulation)
                Else
                    cuve_O2_dissous(I) = 2 + md_ConcentrationOxygeneDissousExcesRegulation_mgL(choix_regulation)
                End If
            End If
        End If
        '**************************** penser aux valeurs sur l'aération
        If cuve_O2_facteur_alfa_force(I) = True Then
            cuve_O2_facteur_alfa(I) = cuve_O2_facteur_alfa_f(I)
        Else
            If choix_cuve_aerateur(I) = aerateur_fine_bulle Then
                cuve_O2_facteur_alfa(I) = 0.65
            Else
                cuve_O2_facteur_alfa(I) = 0.8
            End If
        End If
        If cuve_hauteur_force(I) = True Then
            cuve_hauteur(I) = cuve_hauteur_f(I)
        Else
            cuve_hauteur(I) = 6
        End If
        If cuve_hauteur(I) > hauteur_bassin_max Then
            hauteur_bassin_max = cuve_hauteur(I)
        End If
        If cuve_O2_rdt_transfert_force(I) = True Then
            cuve_O2_rdt_transfert(I) = cuve_O2_rdt_transfert_f(I)
        Else
            cuve_O2_rdt_transfert(I) = 4
        End If
    Next I
    'cuves N
    For I = cuve_N1 To cuve_N2
        If choix_cuve(I) = cuve_oui Then
            If cuve_O2_dissous_force(I) = True Then
                cuve_O2_dissous(I) = cuve_O2_dissous_f(I)
            Else
                If choix_fct_Mox = True And I = cuve_N1 Then
                    cuve_O2_dissous(I) = 1.5
                Else
                    If pd_TemperatureEauDimensionnement_°C <= 20 Then
                        cuve_O2_dissous(I) = 5 + md_ConcentrationOxygeneDissousExcesRegulation_mgL(choix_regulation)
                    Else
                        cuve_O2_dissous(I) = 4 + md_ConcentrationOxygeneDissousExcesRegulation_mgL(choix_regulation)
                    End If
                End If
            End If
        End If
        '**************************** penser aux valeurs sur l'aération
        If cuve_O2_facteur_alfa_force(I) = True Then
            cuve_O2_facteur_alfa(I) = cuve_O2_facteur_alfa_f(I)
        Else
            If choix_cuve_aerateur(I) = aerateur_fine_bulle Then
                cuve_O2_facteur_alfa(I) = 0.75
            Else
                cuve_O2_facteur_alfa(I) = 0.8
            End If
        End If
        If cuve_hauteur_force(I) = True Then
            cuve_hauteur(I) = cuve_hauteur_f(I)
        Else
            If choix_fct_Mox = True Then
                cuve_hauteur(I) = 7
            Else
                cuve_hauteur(I) = 6
            End If
        End If
        If cuve_hauteur(I) > hauteur_bassin_max Then
            hauteur_bassin_max = cuve_hauteur(I)
        End If
        If cuve_O2_rdt_transfert_force(I) = True Then
            cuve_O2_rdt_transfert(I) = cuve_O2_rdt_transfert_f(I)
        Else
            cuve_O2_rdt_transfert(I) = 4
        End If
    Next I
    'cuve DéOx
    If choix_cuve(cuve_deox) = cuve_oui Then
        If cuve_O2_dissous_force(cuve_deox) = True Then
            cuve_O2_dissous(cuve_deox) = cuve_O2_dissous_f(cuve_deox)
        Else
            cuve_O2_dissous(cuve_deox) = 2
        End If
    End If
    'cuve RéOx
    If choix_cuve(cuve_reox) = cuve_oui Then
        If cuve_O2_dissous_force(cuve_reox) = True Then
            cuve_O2_dissous(cuve_reox) = cuve_O2_dissous_f(cuve_reox)
        Else
            cuve_O2_dissous(cuve_reox) = 1.5 + md_ConcentrationOxygeneDissousExcesRegulation_mgL(choix_regulation)
        End If
        '**************************** penser aux valeurs sur l'aération
        If cuve_O2_facteur_alfa_force(cuve_reox) = True Then
            cuve_O2_facteur_alfa(cuve_reox) = cuve_O2_facteur_alfa_f(cuve_reox)
        Else
            If choix_cuve_aerateur(cuve_reox) = aerateur_fine_bulle Then
                cuve_O2_facteur_alfa(cuve_reox) = 0.65
            Else
                cuve_O2_facteur_alfa(cuve_reox) = 0.8
            End If
        End If
        If cuve_hauteur_force(cuve_reox) = True Then
            cuve_hauteur(cuve_reox) = cuve_hauteur_f(cuve_reox)
        Else
            cuve_hauteur(cuve_reox) = 6
        End If
        If cuve_hauteur(cuve_reox) > hauteur_bassin_max Then
            hauteur_bassin_max = cuve_hauteur(cuve_reox)
        End If
        If cuve_O2_rdt_transfert_force(cuve_reox) = True Then
            cuve_O2_rdt_transfert(cuve_reox) = cuve_O2_rdt_transfert_f(cuve_reox)
        Else
            cuve_O2_rdt_transfert(cuve_reox) = 4
        End If
    End If

    'débit d'air pour agitation (Nm3/(m2.h))
    For I = 1 To nb_cuves_max
        If (I >= cuve_C1 And I <= cuve_N2) Or I = cuve_reox Then
            If choix_agitation_cuve_aeree(I) = cuve_aeree_agitation_air Then
                If cuve_aeree_Vair_agitation_Nmh_force(I) = True Then
                    cuve_aeree_Vair_agitation_Nmh(I) = cuve_aeree_Vair_agitation_Nmh_f(I)
                Else
                    If cuve_hauteur(I) >= 5 Then
                        cuve_aeree_Vair_agitation_Nmh(I) = 7   'Nm3/(m2.h)
                    Else
                        cuve_aeree_Vair_agitation_Nmh(I) = media_Vmin_aeration_Nmh(choix_media(I))   'Nm3/(m2.h)
                    End If
                End If
                If cuve_aeree_brassage_air_hj_force(I) = True Then
                    cuve_aeree_brassage_air_hj(I) = cuve_aeree_brassage_air_hj_f(I)
                Else
                    cuve_aeree_brassage_air_hj(I) = 24
                End If
            End If
        End If
    Next I
    
    'ratio d'agitation
    For I = 1 To nb_cuves_max
        If I < cuve_C1 Or (I > cuve_N2 And I <> cuve_reox) Then
            If choix_cuve(I) = cuve_oui Then
                If cuve_ratio_agitation_W_m3_force(I) = True Then
                    cuve_ratio_agitation_W_m3(I) = cuve_ratio_agitation_W_m3_f(I)
                Else
                    cuve_ratio_agitation_W_m3(I) = ratio_elec_agitation
                End If
            Else
                cuve_ratio_agitation_W_m3(I) = 0
            End If
        Else
            If choix_agitation_cuve_aeree(I) = cuve_aeree_agitation_mecanique Then
                If cuve_ratio_agitation_W_m3_force(I) = True Then
                    cuve_ratio_agitation_W_m3(I) = cuve_ratio_agitation_W_m3_f(I)
                Else
                    cuve_ratio_agitation_W_m3(I) = 3
                End If
            Else
                cuve_ratio_agitation_W_m3(I) = 0
            End If
        End If
    Next I

    'Durée depuis le dernier nettoyage des diffuseurs  (ans)
    If diffuseur_encrassement_force = True Then
        diffuseur_encrassement = diffuseur_encrassement_f
    Else
        diffuseur_encrassement = 0
    End If

    'Pression de refoulement des surpresseurs (mCE)
    If air_P_refoulement_moyenne_force = True Then
        air_P_refoulement_moyenne = air_P_refoulement_moyenne_f
    Else
        air_P_refoulement_moyenne = hauteur_bassin_max + 2 + 0.25 * diffuseur_encrassement
    End If

    'Pression des pompes de recirculation de liqueur mixte (mCE)
    If recirculation_MLSS_P_refoulement_force = True Then
        recirculation_MLSS_P_refoulement = recirculation_MLSS_P_refoulement_f
    Else
        recirculation_MLSS_P_refoulement = 5
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
    NumeroErreurNonGeree = "16.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer


    'Hypothèses Mox sur le mainstream
    If choix_fct_Mox = True Then
        rate_N_30(media_K3) = 2.8
        rate_N_30(media_K5) = 2.3
        rate_N_30(media_chip_M) = 2
        'kg N-NO3 formés par kg de N nitrifiés
        ratio_NO3f_nit = 0.11
        'Consommation d'O2 pour la nit en mox
        ratio_O2_nit_Mox = 1.94
        'concentration en sortie de la cuve Mox
        cuve_NH4_outlet_Mox_mgL = 10
        'concentration en sortie de C1 et N1 (si C+N)
        hyp_cuve_C1_DBOsol_mgL = 10
        hyp_cuve_N_deox_DBOsol_mgL = 10
    Else
        rate_N_30(media_K3) = 0
        rate_N_30(media_K5) = 0
        rate_N_30(media_chip_M) = 0
        'kg N-NO3 formés par kg de N nitrifiés
        ratio_NO3f_nit = 1
        'concentration en sortie de C1 et N1 (si C+N)
        hyp_cuve_C1_DBOsol_mgL = 5
        hyp_cuve_N_deox_DBOsol_mgL = 2
    End If

    'ratio de conso d'O2 de la nit
    ratio_O2_nit = 4.57

    'valeurs pour le message d'erreur
    For I = 1 To nb_cuves_max
        erreur_Qair_augmente_agitation(I) = erreur_predenit_2cuves + I
    Next I


    'concentration à rajouter à celle de la Post-denit pour avoir celle du dernier tank aéré   mg/L
    delta_NO3_post_DN = 10

    'calcul du ratio MV/MES eau brute   a0+a1*MES/DCO+a2*DCO/DBO
    a_ratio_EB_MV_MES(0) = 1.213
    a_ratio_EB_MV_MES(1) = -1 / 3
    a_ratio_EB_MV_MES(2) = -0.1429
    ratio_EB_MV_MES_min = 0.4
    ratio_EB_MV_MES_max = 0.8
    'ratio MV/MES eau décantée
    ratio_ED_MV_MES = 0.75

    'ratio de N assimilé
    ratio_N_assimile = 0.05

    'ratio de N assimilé par la DCO du méthanol qui sert à la production de boues
    ratio_N_assimile_Capporte = 0.08

    'ratio de P assimilé par le traitement de la DBO
    ratio_P_assimile = 0.01

    'ratio de DCO dure par rapport à la DCO totale de l'eau brute
    ratio_DCO_dure_total = 0.33 * 0.13

    'ratio de conso d'O2 (en l'ocurrence DBO) lors de la dénit
    ratio_O2_denit = -2.86

    'NO3 équivalents dus à l'O2 dissous
    ratio_NO3eq_O2 = 0.35

    'compartimentation de l'eau brute +hydrolyse de l'eau brute
    If choix_primaire = True Then
        If choix_decanteur_simple = True Then
            EB_rate_DBOsol_DBO = 0.482
            EB_rate_Norgasol_Norga = 0.489
            rate_Norgapart_hydrolyse = 0.95
            rate_DBOpart_hydrolyse = 0.3
        Else
            EB_rate_DBOsol_DBO = 0.699
            EB_rate_Norgasol_Norga = 0.71
            rate_Norgapart_hydrolyse = 0.95
            rate_DBOpart_hydrolyse = 0.3
        End If
    Else
        EB_rate_DBOsol_DBO = 0.406
        EB_rate_Norgasol_Norga = 0.386
        rate_Norgapart_hydrolyse = 0.95
        rate_DBOpart_hydrolyse = 0.3
    End If


    'ratio de conso de l'O2 pour la DBO5
    ratio_O2_DBO = 0.75

    'ratio de DCO à apporter pour la dénit   kgO2/kgNNO3
    If pd_TemperatureEauDimensionnement_°C < 5 Then
        rate_DBO_apportee_denit_nominal = 5
    ElseIf pd_TemperatureEauDimensionnement_°C > 15 Then
        rate_DBO_apportee_denit_nominal = 4
    Else
        rate_DBO_apportee_denit_nominal = 5 - 0.1 * (pd_TemperatureEauDimensionnement_°C - 5)
    End If
    If pd_TemperatureEauConditionsReelles_°C < 5 Then
        rate_DBO_apportee_denit_reel = 5
    ElseIf pd_TemperatureEauConditionsReelles_°C > 15 Then
        rate_DBO_apportee_denit_reel = 4
    Else
        rate_DBO_apportee_denit_reel = 5 - 0.1 * (pd_TemperatureEauConditionsReelles_°C - 5)
    End If

    'DCO du méthanol   kgO2/kg méthanol
    rate_DBO_methanol = 0.96
    rate_DBO_ethanol = 1.35
    rate_DBO_source_C = 0.96
    ratio_boues_DBO_apportee = 0.25  'kg MES/kg O2

    'taux de remplissage maximum
    For I = 1 To nb_cuves_max
        For j = 1 To nb_type_media
            media_filling_max(I, j) = 0
        Next j
    Next I
    For I = media_K1 To media_K5
        media_filling_max(cuve_predenit1, I) = 0.5
        media_filling_max(cuve_predenit2, I) = 0.5
        media_filling_max(cuve_deox, I) = 0.5
        media_filling_max(cuve_postdenit1, I) = 0.5
        media_filling_max(cuve_postdenit2, I) = 0.5
    Next I
    For j = media_K3 To media_K5
        For I = cuve_C1 To cuve_N2
            media_filling_max(I, j) = 0.65
        Next I
        media_filling_max(cuve_reox, j) = 0.65
    Next j
    For I = media_chip_P To media_chip_M
        media_filling_max(cuve_reox, I) = 0.55
        media_filling_max(cuve_N1, I) = 0.55
        media_filling_max(cuve_N2, I) = 0.55
    Next I
    For I = cuve_C1 To cuve_C2
        media_filling_max(I, media_chip_P) = 0.55
        media_filling_max(I, media_chip_P) = 0.55
    Next I

    'valeurs guides de taux de remplissage              par défaut égal au taux max mais permet de modifier si on préfère fixer des valeurs plus faibles
    For I = 1 To nb_cuves_max
        For j = 1 To nb_type_media
            hyp_cuve_filling(I, j) = media_filling_max(I, j)
        Next j
    Next I

    'surface spécifique de chaque type de média (m2/m3)
    media_surface_m2m3(media_K1) = 500
    media_surface_m2m3(media_K1_heavy) = 500
    media_surface_m2m3(media_K3) = 500
    media_surface_m2m3(media_K5) = 800
    media_surface_m2m3(media_chip_P) = 900
    media_surface_m2m3(media_chip_M) = 1200

    'vitesse mini de l'air pour agitation (Nm/h)
    media_Vmin_aeration_Nmh(media_K1) = 10
    media_Vmin_aeration_Nmh(media_K1_heavy) = 10
    media_Vmin_aeration_Nmh(media_K3) = 10
    media_Vmin_aeration_Nmh(media_K5) = 10
    media_Vmin_aeration_Nmh(media_chip_P) = 15
    media_Vmin_aeration_Nmh(media_chip_M) = 10

    'constante de réaction pour les cuves N et la DéOx
    If choix_cuve(cuve_predenit1) = cuve_oui Then
        If choix_primaire = True Then
            If choix_decanteur_simple = True Then
                k_nit = 0.53
            Else
                k_nit = 0.58
            End If
        Else
            k_nit = 0.47
        End If
    Else
        If choix_primaire = True Then
            If choix_decanteur_simple = True Then
                k_nit = 0.47
            Else
                k_nit = 0.58
            End If
        Else
            k_nit = 0.4
        End If
    End If

    'vitesse de traitement de la DBO cuve C (C+N)à 10°C
    hyp_rate_DBOin_appliquee_10°C_CN = 3.9

    'DBO apportée non consommée
    DBO_apportee_fuite_mgL = 5

    'ratio DCO/DBO source C   '**************************** à modifier lorsque choix possible ********************************
    ratio_DCO_DBO_source_C = 1.5 / 0.96    ' pour le MetOH
    'ratio_DCO_DBO_source_C = 2.09 / 1.35    ' pour l'EtOH

    'HRT minimum de chaque cuve
    For I = 1 To nb_cuves_max
        cuve_HRTmini_h(I) = 15 / pd_NOMBRE_MINUTE_PAR_HEURE
    Next I

    'concentrations limites pour^passer de 1 à 2 cuves   en mg/L les concentrations en sortie d'étape de traitement
    predenit_NO3_limite = 3
    postdenit_NO3_limite = 2

    'facteur beta
    O2_facteur_beta = 0.95

    'calcul de la pression atmosphérique  Patm=P0*((a0-a1*altitude)^a2)/100
    P0 = 101325
    a0_Patm = 1
    a1_Patm = -0.0000225577
    a2_Patm = 5.25588

    'calcul de la surpression en diffusion de bulles  delta_P=a0/a1*P0/100*(hauteur_bassin-hauteur_diffuseur)
    a0_delta_P = 0.35
    a1_delta_P = 10.33
    hauteur_diffuseur = 0.25    'm

    'calcul O2 saturation  O2_saturation(T,P0)= somme(ai*Teau^i)
    a_O2_saturation(0) = 14.628
    a_O2_saturation(1) = -0.4223
    a_O2_saturation(2) = 0.0102
    a_O2_saturation(3) = -0.0002
    a_O2_saturation(4) = 0.000001

    'calcul de la pression de vapeur saturante     Pvap(T)=somme(ai*Tair^i)
    a_Pvap(0) = 6.1089
    a_Pvap(1) = 0.4456
    a_Pvap(2) = 0.0139
    a_Pvap(3) = 0.0003
    a_Pvap(4) = 0.000002
    a_Pvap(5) = 0.00000003
    a_Pvap(6) = -0.00000000002

    'correction température pour facteur K
    correctif_T_K = 1.024
    T_ref_K = 20

    'O2 présent dans l'air (kgO2/Nm3air)
    ratio_kgO2_Nm3air = 0.3

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
    md_ConcentrationOxygeneDissousExcesRegulation_mgL(regulation_horloge) = 2
    md_ConcentrationOxygeneDissousExcesRegulation_mgL(regulation_sans_variation) = 1
    md_ConcentrationOxygeneDissousExcesRegulation_mgL(regulation_avec_variation) = 0.5
    md_ConcentrationOxygeneDissousExcesRegulation_mgL(regulation_avance) = 0

    'agitation des cuves anoxiques (W/m3)
    ratio_elec_agitation = 7.5
    agitation_fct = 24

    'ratio de DCO dure par rapport à la DCO totale de l'eau brute
    ratio_DCO_dure_total = 0.33 * 0.13

    'concentration minimale en DBO en sortie
    hyp_outlet_DBO_mgL_mini = 10


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub






Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    
    choix_fct_Mox = Feuil57.Range("choix_fct_Mox").Cells(1, pi_FiliereConsideree)
    
    'on regarde l'ensemble des choix
    choix_type_separation = Feuil16.Range("MBBR_choix_type_separation").Cells(1, pi_FiliereConsideree)
    nb_etapes_traitement = 0
    For I = 1 To nb_cuves_max
        choix_cuve(I) = Feuil16.Range("MBBR_choix_cuve").Cells(I, pi_FiliereConsideree)
        If choix_cuve(I) = cuve_oui Then
            nb_etapes_traitement = nb_etapes_traitement + 1
        End If
        choix_media(I) = Feuil16.Range("MBBR_choix_media").Cells(I, pi_FiliereConsideree)
        choix_cuve_aerateur(I) = Feuil16.Range("MBBR_choix_cuve_aerateur").Cells(I, pi_FiliereConsideree)
        choix_agitation_cuve_aeree(I) = Feuil16.Range("MBBR_choix_agitation_cuve_aeree").Cells(I, pi_FiliereConsideree)
    Next I
    choix_source_C = methanol
    choix_surpresseur = Feuil16.Range("MBBR_choix_surpresseur").Cells(1, pi_FiliereConsideree)
    choix_regulation = Feuil16.Range("MBBR_choix_regulation").Cells(1, pi_FiliereConsideree)


    If choix_cuve(cuve_predenit1) = cuve_oui And choix_cuve(cuve_N1) = cuve_non Then
        MsgBox Feuil16.Range("MBBR_erreur_nom") & ": " & Feuil16.Range("MBBR_erreur_predenit_sans_N").Cells(1, 1), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
    ElseIf choix_cuve(cuve_postdenit1) = cuve_oui And choix_cuve(cuve_N1) = cuve_non Then
        MsgBox Feuil16.Range("MBBR_erreur_nom") & ": " & Feuil16.Range("MBBR_erreur_postdenit_sans_N").Cells(1, 1), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
    ElseIf choix_cuve(cuve_reox) = cuve_oui And choix_cuve(cuve_postdenit1) = cuve_non Then
        MsgBox Feuil16.Range("MBBR_erreur_nom") & ": " & Feuil16.Range("MBBR_erreur_reox_sans_postdenit").Cells(1, 1), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
    ElseIf choix_fct_Mox = True And (choix_cuve(cuve_predenit1) = cuve_oui Or choix_cuve(cuve_C1 = cuve_oui)) Then
        MsgBox ("La configuration MBBR Mox ne peut pas être utilisée avec de la prédénitrification et/ou de l'abattement du carbone")
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
    ElseIf choix_fct_Mox = True And (choix_media(cuve_N1) = media_chip_P Or choix_media(cuve_N1) = media_K1 Or choix_media(cuve_N1) = media_K1_heavy) Then
        MsgBox ("media non utilisable en configuration Mox --> choix possibles: K3, K5, Chip M")
        pn_EffacementResultatsEnCours = True
        Exit Sub
    End If

    

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'on lit les valeurs forcées

    Dim I As Integer


    nominal_outlet_DBO_mgL_force = Feuil16.Range("MBBR_nominal_outlet_DBO_mgL").Cells(1, pi_FiliereConsideree)
    nominal_outlet_NH4_mgL_force = Feuil16.Range("MBBR_nominal_outlet_NH4_mgL").Cells(1, pi_FiliereConsideree)
    nominal_outlet_NO3_mgL_force = Feuil16.Range("MBBR_nominal_outlet_NO3_mgL").Cells(1, pi_FiliereConsideree)
    Norga_dure_soluble_force = Feuil16.Range("MBBR_Norga_dure_soluble").Cells(1, pi_FiliereConsideree)
    nominal_predenit_NO3_mgL_force = Feuil16.Range("MBBR_nominal_predenit_NO3_mgL").Cells(1, pi_FiliereConsideree)
    nominal_last_aerated_tank_NO3_mgL_force = Feuil16.Range("MBBR_nominal_last_aerated_tank_NO3_mgL").Cells(1, pi_FiliereConsideree)
    nominal_recirculation_MLSS_taux_force = Feuil16.Range("MBBR_nominal_recirculation_MLSS_taux").Cells(1, pi_FiliereConsideree)
    ratio_elimination_predenit1_predenit12_force = Feuil16.Range("MBBR_ratio_elimination_predenit1_predenit12").Cells(1, pi_FiliereConsideree)
    ratio_elimination_C1_C12_force = Feuil16.Range("MBBR_ratio_elimination_C1_C12").Cells(1, pi_FiliereConsideree)
    ratio_elimination_N1_N12_force = Feuil16.Range("MBBR_ratio_elimination_N1_N12").Cells(1, pi_FiliereConsideree)
    ratio_elimination_postdenit1_postdenit12_force = Feuil16.Range("MBBR_ratio_elimination_postdenit1_postdenit12").Cells(1, pi_FiliereConsideree)
    For I = 1 To nb_cuves_max
        cuve_O2_dissous_force(I) = Feuil16.Range("MBBR_cuve_O2_dissous").Cells(I, pi_FiliereConsideree)
        If cuve_O2_dissous_force(I) = True Then
            cuve_O2_dissous_f(I) = Feuil16.Range("MBBR_cuve_O2_dissous").Cells(I, pi_FiliereConsideree)
        End If
        cuve_filling_force(I) = Feuil16.Range("MBBR_cuve_filling").Cells(I, pi_FiliereConsideree)
        If cuve_filling_force(I) = True Then
            cuve_filling_f(I) = Feuil16.Range("MBBR_cuve_filling").Cells(I, pi_FiliereConsideree)
        End If
        cuve_volume_force(I) = Feuil16.Range("MBBR_cuve_volume").Cells(I, pi_FiliereConsideree)
        If cuve_volume_force(I) = True Then
            cuve_volume_f(I) = Feuil16.Range("MBBR_cuve_volume").Cells(I, pi_FiliereConsideree)
        End If
        cuve_ratio_agitation_W_m3_force(I) = Feuil16.Range("MBBR_cuve_ratio_agitation_W_m3").Cells(I, pi_FiliereConsideree)
        If cuve_ratio_agitation_W_m3_force(I) = True Then
            cuve_ratio_agitation_W_m3_f(I) = Feuil16.Range("MBBR_cuve_ratio_agitation_W_m3").Cells(I, pi_FiliereConsideree)
        End If
        cuve_O2_facteur_alfa_force(I) = Feuil16.Range("MBBR_cuve_O2_facteur_alfa").Cells(I, pi_FiliereConsideree)
        If cuve_O2_facteur_alfa_force(I) = True Then
            cuve_O2_facteur_alfa_f(I) = Feuil16.Range("MBBR_cuve_O2_facteur_alfa").Cells(I, pi_FiliereConsideree)
        End If
        cuve_O2_rdt_transfert_force(I) = Feuil16.Range("MBBR_cuve_O2_rdt_transfert").Cells(I, pi_FiliereConsideree)
        If cuve_O2_rdt_transfert_force(I) = True Then
            cuve_O2_rdt_transfert_f(I) = Feuil16.Range("MBBR_cuve_O2_rdt_transfert").Cells(I, pi_FiliereConsideree)
        End If
        cuve_hauteur_force(I) = Feuil16.Range("MBBR_cuve_hauteur").Cells(I, pi_FiliereConsideree)
        If cuve_hauteur_force(I) = True Then
            cuve_hauteur_f(I) = Feuil16.Range("MBBR_cuve_hauteur").Cells(I, pi_FiliereConsideree)
        End If
        cuve_aeree_Vair_agitation_Nmh_force(I) = Feuil16.Range("MBBR_cuve_aeree_Vair_brassage_Nmh").Cells(I, pi_FiliereConsideree)
        If cuve_aeree_Vair_agitation_Nmh_force(I) = True Then
            cuve_aeree_Vair_agitation_Nmh_f(I) = Feuil16.Range("MBBR_cuve_aeree_Vair_brassage_Nmh").Cells(I, pi_FiliereConsideree)
        End If
        cuve_aeree_brassage_air_hj_force(I) = Feuil16.Range("MBBR_cuve_aeree_brassage_air_hj").Cells(I, pi_FiliereConsideree)
        If cuve_aeree_brassage_air_hj_force(I) = True Then
            cuve_aeree_brassage_air_hj_f(I) = Feuil16.Range("MBBR_cuve_aeree_brassage_air_hj").Cells(I, pi_FiliereConsideree)
        End If
    Next I
    postdenit_carbone_apporte_flux_force = Feuil16.Range("MBBR_postdenit_carbone_apporte_flux").Cells(1, pi_FiliereConsideree)
    reel_recirculation_MLSS_taux_force = Feuil16.Range("MBBR_reel_recirculation_MLSS_taux").Cells(1, pi_FiliereConsideree)
    reel_outlet_DBO_mgL_force = Feuil16.Range("MBBR_reel_outlet_DBO_mgL").Cells(1, pi_FiliereConsideree)
    reel_outlet_NH4_mgL_force = Feuil16.Range("MBBR_reel_outlet_NH4_mgL").Cells(1, pi_FiliereConsideree)
    reel_outlet_NO3_mgL_force = Feuil16.Range("MBBR_reel_outlet_NO3_mgL").Cells(1, pi_FiliereConsideree)
    O2_besoin_total_force = Feuil16.Range("MBBR_O2_besoin_total").Cells(1, pi_FiliereConsideree)
    air_Q_Nm3j_force = Feuil16.Range("MBBR_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
    diffuseur_encrassement_force = Feuil16.Range("MBBR_diffuseur_encrassement").Cells(1, pi_FiliereConsideree)
    air_P_refoulement_moyenne_force = Feuil16.Range("MBBR_air_P_refoulement_moyenne").Cells(1, pi_FiliereConsideree)
    surpresseur_conso_spec_force = Feuil16.Range("MBBR_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
    recirculation_MLSS_P_refoulement_force = Feuil16.Range("MBBR_recirculation_MLSS_P_refoulement").Cells(1, pi_FiliereConsideree)
    recirculation_MLSS_pompe_rdt_force = Feuil16.Range("MBBR_recirculation_MLSS_pompe_rdt").Cells(1, pi_FiliereConsideree)
    'boues_indice_Mohlman_force
    sortie_MES_force = Feuil16.Range("MBBR_outlet_MES_concentration").Cells(1, pi_FiliereConsideree)
    boues_concentration_force = Feuil16.Range("MBBR_boues_concentration").Cells(1, pi_FiliereConsideree)
    boues_MES_force = Feuil16.Range("MBBR_boues_MES").Cells(1, pi_FiliereConsideree)
    boues_MV_MES_force = Feuil16.Range("MBBR_boues_MV_MES").Cells(1, pi_FiliereConsideree)
    extraction_P_refoulement_force = Feuil16.Range("MBBR_extraction_P_refoulement").Cells(1, pi_FiliereConsideree)
    extraction_pompe_rdt_force = Feuil16.Range("MBBR_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)




    If nominal_outlet_DBO_mgL_force = True Then
        nominal_outlet_DBO_mgL_f = Feuil16.Range("MBBR_nominal_outlet_DBO_mgL").Cells(1, pi_FiliereConsideree)
    End If
    If nominal_outlet_NH4_mgL_force = True Then
        nominal_outlet_NH4_mgL_f = Feuil16.Range("MBBR_nominal_outlet_NH4_mgL").Cells(1, pi_FiliereConsideree)
    End If
    If nominal_outlet_NO3_mgL_force = True Then
        nominal_outlet_NO3_mgL_f = Feuil16.Range("MBBR_nominal_outlet_NO3_mgL").Cells(1, pi_FiliereConsideree)
    End If
    If Norga_dure_soluble_force = True Then
        Norga_dure_soluble_f = Feuil16.Range("MBBR_Norga_dure_soluble").Cells(1, pi_FiliereConsideree)
    End If
    If nominal_predenit_NO3_mgL_force = True Then
        nominal_predenit_NO3_mgL_f = Feuil16.Range("MBBR_nominal_predenit_NO3_mgL").Cells(1, pi_FiliereConsideree)
    End If
    If nominal_last_aerated_tank_NO3_mgL_force = True Then
        nominal_last_aerated_tank_NO3_mgL_f = Feuil16.Range("MBBR_nominal_last_aerated_tank_NO3_mgL").Cells(1, pi_FiliereConsideree)
    End If
    If nominal_recirculation_MLSS_taux_force = True Then
        nominal_recirculation_MLSS_taux_f = Feuil16.Range("MBBR_nominal_recirculation_MLSS_taux").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_elimination_predenit1_predenit12_force = True Then
        ratio_elimination_predenit1_predenit12_f = Feuil16.Range("MBBR_ratio_elimination_predenit1_predenit12").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_elimination_C1_C12_force = True Then
        ratio_elimination_C1_C12_f = Feuil16.Range("MBBR_ratio_elimination_C1_C12").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_elimination_N1_N12_force = True Then
        ratio_elimination_N1_N12_f = Feuil16.Range("MBBR_ratio_elimination_N1_N12").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_elimination_postdenit1_postdenit12_force = True Then
        ratio_elimination_postdenit1_postdenit12_f = Feuil16.Range("MBBR_ratio_elimination_postdenit1_postdenit12").Cells(1, pi_FiliereConsideree)
    End If
    If postdenit_carbone_apporte_flux_force = True Then
        postdenit_carbone_apporte_flux_f = Feuil16.Range("MBBR_postdenit_carbone_apporte_flux").Cells(1, pi_FiliereConsideree)
    End If
    If reel_recirculation_MLSS_taux_force = True Then
        reel_recirculation_MLSS_taux_f = Feuil16.Range("MBBR_reel_recirculation_MLSS_taux").Cells(1, pi_FiliereConsideree)
    End If
    If reel_outlet_DBO_mgL_force = True Then
        reel_outlet_DBO_mgL_f = Feuil16.Range("MBBR_reel_outlet_DBO_mgL").Cells(1, pi_FiliereConsideree)
    End If
    If reel_outlet_NH4_mgL_force = True Then
        reel_outlet_NH4_mgL_f = Feuil16.Range("MBBR_reel_outlet_NH4_mgL").Cells(1, pi_FiliereConsideree)
    End If
    If reel_outlet_NO3_mgL_force = True Then
        reel_outlet_NO3_mgL_f = Feuil16.Range("MBBR_reel_outlet_NO3_mgL").Cells(1, pi_FiliereConsideree)
    End If
    If O2_besoin_total_force = True Then
        O2_besoin_total_f = Feuil16.Range("MBBR_O2_besoin_total").Cells(1, pi_FiliereConsideree)
    End If
    If air_Q_Nm3j_force = True Then
        air_Q_Nm3j_f = Feuil16.Range("MBBR_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
    End If
    If diffuseur_encrassement_force = True Then
        diffuseur_encrassement_f = Feuil16.Range("MBBR_diffuseur_encrassement").Cells(1, pi_FiliereConsideree)
    End If
    If air_P_refoulement_moyenne_force = True Then
        air_P_refoulement_moyenne_f = Feuil16.Range("MBBR_air_P_refoulement_moyenne").Cells(1, pi_FiliereConsideree)
    End If
    If surpresseur_conso_spec_force = True Then
        surpresseur_conso_spec_f = Feuil16.Range("MBBR_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_MLSS_P_refoulement_force = True Then
        recirculation_MLSS_P_refoulement_f = Feuil16.Range("MBBR_recirculation_MLSS_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_MLSS_pompe_rdt_force = True Then
        recirculation_MLSS_pompe_rdt_f = Feuil16.Range("MBBR_recirculation_MLSS_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    'boues_indice_Mohlman_f
    If sortie_MES_force = True Then
        sortie_MES_f = Feuil16.Range("MBBR_outlet_MES_concentration").Cells(1, pi_FiliereConsideree)
    End If
    If boues_concentration_force = True Then
        boues_concentration_f = Feuil16.Range("MBBR_boues_concentration").Cells(1, pi_FiliereConsideree)
    End If
    If boues_MES_force = True Then
        boues_MES_f = Feuil16.Range("MBBR_boues_MES").Cells(1, pi_FiliereConsideree)
    End If
    If boues_MV_MES_force = True Then
        boues_MV_MES_f = Feuil16.Range("MBBR_boues_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement_f = Feuil16.Range("MBBR_extraction_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt_f = Feuil16.Range("MBBR_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If

    'Spécifiques au Mox
    If choix_fct_Mox = True Then
        rate_N_T_design_force = Feuil57.Range("rate_N_T_design").Cells(1, pi_FiliereConsideree)
        If rate_N_T_design_force = True Then
            rate_N_T_design_f = Feuil57.Range("rate_N_T_design").Cells(1, pi_FiliereConsideree)
        End If
        rate_N_T_reel_force = Feuil57.Range("rate_N_T_reel").Cells(1, pi_FiliereConsideree)
        If rate_N_T_reel_force = True Then
            rate_N_T_reel_f = Feuil57.Range("rate_N_T_reel").Cells(1, pi_FiliereConsideree)
        End If
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub separation_boues(ByVal MESin, ByVal DBOin, ByVal Qin, ByVal temperature, ByVal outlet_DBO_mgL, ByVal outlet_DBOsol_mgL, ByVal boues_pdtes_aerobie, ByVal CalculConditionsRelles As Boolean)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'Ici on calcule les productions de boues et on simule la séparation

    Dim boues_pdtes_bio As Double
    Dim boues_pdtes_chimiques As Double
    Dim ratio_boues_DBOeliminee As Double

    If sortie_MES_force = True Then
        sortie_MES = sortie_MES_f
    Else
        If nb_etapes_traitement = 1 And choix_cuve(cuve_C1) = cuve_oui Then
            If choix_type_separation = clarif Then
                sortie_MES = 0.5356 * outlet_DBO_mgL + 16.668
            ElseIf choix_type_separation = clarif_polymere Then
                sortie_MES = 0.5431 * outlet_DBO_mgL + 10.23
            Else
                sortie_MES = 0.8268 * outlet_DBO_mgL + 0.7594
            End If
        Else
            sortie_MES = (outlet_DBO_mgL - outlet_DBOsol_mgL) / 0.5555
        End If
    End If

    'ratio_boues_DBOeliminee = 0.6 * (MESin / DBOin + 1) - 0.72 * 1.072 ^ (temperature - 15) / (1000 * DBOin / surface_media_totale + 1.33 * 1.072 ^ (temperature - 15))
    'boues_pdtes_bio = ratio_boues_DBOeliminee * (DBOin - outlet_DBO_mgL * Qin / 1000)  'à vérifier comment on calcule la concentration en DBO en sous-charge
    boues_pdtes_bio = boues_pdtes_aerobie
    boues_pdtes_bio = boues_pdtes_bio + boues_postdenit

    boues_MV_MES = boues_MV_MES / (boues_pdtes_bio + boues_pdtes_chimiques)
    
    If CalculConditionsRelles = True And boues_MES_force = True Then
        boues_MES = boues_MES_f
    Else
        boues_MES = (boues_pdtes_bio + boues_pdtes_chimiques) - sortie_MES * Qin / 1000
    End If
    boues_Q = boues_MES / boues_concentration




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "16.18"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub



