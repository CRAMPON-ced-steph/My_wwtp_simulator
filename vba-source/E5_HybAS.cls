Option Explicit

'Fonctionnement en mode Mox
Dim choix_fct_Mox As Boolean

'constantes pour tracer les commentaires
Private Const erreur_deox_volume_insuffisant = 1
Private Const erreur_charge_radier = 2
Private Const erreur_taux_recirc = 3
Private Const erreur_taux_recirc_MLSS = 4
Private Const erreur_choix_surpresseur = 5
Private Const erreur_aeration_augmentee_agitation = 6
Private Const erreur_postDN_media = 7
Private Const erreur_post_DN_volume_insuffisant = 8
Private Const erreur_traitement_DBO_insuffisant = 9
Private Const erreur_media_filling = 10

'constantes pour les calculs
Private Const denit_nb_iteration = 10
Private Const critere_convergence = 0.000001

'constantes pour le choix des cuves
Private Const cuve_non = 0
Private Const cuve_oui_sans_media = 1
Private Const cuve_oui_avec_media = 2

'constante pour la configuration des cuves aérées
Private Const nb_configurations_cuves_aerees = 6
Private Const config_CN_H = 1
Private Const config_CN_H_CN = 2
Private Const config_CN_H_H = 3
Private Const config_H_H = 4
Private Const config_H_CN = 5
Private Const config_H = 6

'constantes pour le type de cuve aérée
Private Const cuve_aeree_CN = 1
Private Const cuve_aeree_H = 2

'types d'aérateurs
Private Const aerateur_fine_bulle = 1
Private Const aerateur_moyenne_bulle = 2

'constantes pour le type de média
Private Const nb_medias = 6
Private Const no_media = 0
Private Const media_K1 = 1
Private Const media_K1_heavy = 2
Private Const media_K3 = 3
Private Const media_K5 = 4
Private Const media_chip_P = 5
Private Const media_chip_M = 6

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

'constantes pour le choix du type de râcleur
Private Const nb_type_racleur = 3  'nombre de type de râcleurs
Private Const racleur_racle = 1
Private Const racleur_racle_suce = 2
Private Const racleur_Kruger = 3

'choix de l'utilisateur
Dim choix_pre_anoxie As Integer
Dim choix_anaerobie As Integer
Dim choix_pre_denit As Integer
Dim choix_deox As Integer
Dim choix_reox As Integer
Dim choix_aerateur_reox As Integer
Dim choix_post_denit As Integer
Dim choix_media_post_denit As Integer
Dim choix_configuration_cuves_aerees As Integer
Dim choix_nb_cuves_aerees As Integer
Dim choix_type_cuve_aeree() As Integer  'on redimensionne avec le nombre de cuves
Dim choix_media_cuve_aeree() As Integer   'on redimensionne avec le nombre de cuves
Dim choix_aerateur_cuve_aeree() As Integer   'on redimensionne avec le nombre de cuves
'Dim choix_aerateur As Integer    A VOIR SI BESOIN OU NON
Dim choix_regulation As Integer
Dim choix_surpresseur As Integer
Dim choix_racleur As Integer

'variables échangées avec l'utilisateur
Dim outlet_NH4 As Double
Dim outlet_NO3 As Double
Dim pre_denit_NO3 As Double
Dim NO3_last_aerated_tank As Double
Dim Norga_dure_soluble As Double
Dim outlet_DBO_soluble As Double
Dim recirculation_taux As Double   'boues du clarif
Dim nominal_MES_concentration As Double
Public nominal_MV_MES As Double
Dim nominal_recirculation_MLSS_taux As Double
Dim reel_MES_concentration As Double
Public reel_MV_MES As Double
Dim reel_recirculation_MLSS_taux As Double
Dim pre_anoxie_volume As Double
Dim anaerobie_volume As Double
Public anaerobie_FeCl3_flux As Double
Dim pre_denit_volume As Double
Dim deox_O2_dissous As Double
Dim deox_volume As Double
Dim reox_O2_dissous As Double
Dim reox_volume As Double
Dim reox_O2_facteur_alfa As Double
Dim reox_O2_rdt_transfert As Double
Dim reox_hauteur_bassin As Double
Public post_denit_methanol_flux As Double
Dim post_denit_media_filling As Double
Dim post_denit_volume As Double
Dim ratio_Nnit_media_Nnit_total As Double
Dim ratio_Nnit_mediaH1_Nnit_media As Double
Dim ratio_Nnit_CN1_Nnit_CN12 As Double
Dim cuve_aeree_filling() As Double
Dim cuve_aeree_O2_dissous() As Double
Dim cuve_aeree_volume() As Double
Dim cuve_aeree_O2_facteur_alfa() As Double
Dim cuve_aeree_O2_rdt_transfert() As Double
Dim cuve_aeree_hauteur_bassin() As Double
Dim volume_total_bassins As Double
Dim outlet_reel_DBO As Double
Dim outlet_reel_NH4 As Double
Dim outlet_reel_NO3 As Double
Dim ratio_elec_anaerobie As Double
Dim ratio_elec_pre_anoxie As Double
Dim ratio_elec_pre_denit As Double
Dim ratio_elec_deox As Double
Dim ratio_elec_post_denit As Double
Dim O2_besoin_total As Double
Dim air_Q_Nm3j As Double
Dim diffuseur_encrassement As Double
Dim air_P_refoulement_moyenne As Double
Dim surpresseur_conso_spec As Double
Dim recirculation_P_refoulement As Double
Dim recirculation_pompe_rdt As Double
Dim recirculation_MLSS_P_refoulement As Double
Dim recirculation_MLSS_pompe_rdt As Double
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

Public deltaP_precipitation As Double
Public electricite_fixe As Double

'variables forcées par l'utilisateur (pour repérer si forcée ou non)
Dim outlet_NH4_force As Boolean
Dim outlet_NO3_force As Boolean
Dim pre_denit_NO3_force As Boolean
Dim NO3_last_aerated_tank_force As Boolean
Dim Norga_dure_soluble_force As Boolean
Dim outlet_DBO_soluble_force As Boolean
Dim recirculation_taux_force As Boolean   'boues du clarif
Dim nominal_MES_concentration_force As Boolean
Dim nominal_MV_MES_force As Boolean
Dim nominal_recirculation_MLSS_taux_force As Boolean
Dim reel_MES_concentration_force As Boolean
Dim reel_MV_MES_force As Boolean
Dim reel_recirculation_MLSS_taux_force As Boolean
Dim pre_anoxie_volume_force As Boolean
Dim anaerobie_volume_force As Boolean
Dim anaerobie_FeCl3_flux_force As Boolean
Dim pre_denit_volume_force As Boolean
Dim deox_O2_dissous_force As Boolean
Dim deox_volume_force As Boolean
Dim reox_O2_dissous_force As Boolean
Dim reox_volume_force As Boolean
Dim reox_O2_facteur_alfa_force As Boolean
Dim reox_O2_rdt_transfert_force As Boolean
Dim reox_hauteur_bassin_force As Boolean
Dim post_denit_methanol_flux_force As Boolean
Dim post_denit_media_filling_force As Boolean
Dim post_denit_volume_force As Boolean
Dim ratio_Nnit_media_Nnit_total_force As Boolean
Dim ratio_Nnit_mediaH1_Nnit_media_force As Boolean
Dim ratio_Nnit_CN1_Nnit_CN12_force As Boolean
Dim cuve_aeree_filling_force() As Boolean
Dim cuve_aeree_O2_dissous_force() As Boolean
Dim cuve_aeree_volume_force() As Boolean
Dim cuve_aeree_O2_facteur_alfa_force() As Boolean
Dim cuve_aeree_O2_rdt_transfert_force() As Boolean
Dim cuve_aeree_hauteur_bassin_force() As Boolean
Dim outlet_reel_DBO_force As Boolean
Dim outlet_reel_NH4_force As Boolean
Dim outlet_reel_NO3_force As Boolean
Dim ratio_elec_anaerobie_force As Boolean
Dim ratio_elec_pre_anoxie_force As Boolean
Dim ratio_elec_pre_denit_force As Boolean
Dim ratio_elec_deox_force As Boolean
Dim ratio_elec_post_denit_force As Boolean
Dim O2_besoin_total_force As Boolean
Dim air_Q_Nm3j_force As Boolean
Dim diffuseur_encrassement_force As Boolean
Dim air_P_refoulement_moyenne_force As Boolean
Dim surpresseur_conso_spec_force As Boolean
Dim recirculation_P_refoulement_force As Boolean
Dim recirculation_pompe_rdt_force As Boolean
Dim recirculation_MLSS_P_refoulement_force As Boolean
Dim recirculation_MLSS_pompe_rdt_force As Boolean
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

'variables forcées par l'utilisateur
Dim outlet_NH4_f As Double
Dim outlet_NO3_f As Double
Dim pre_denit_NO3_f As Double
Dim NO3_last_aerated_tank_f As Double
Dim Norga_dure_soluble_f As Double
Dim outlet_DBO_soluble_f As Double
Dim recirculation_taux_f As Double   'boues du clarif
Dim nominal_MES_concentration_f As Double
Dim nominal_MV_MES_f As Double
Dim nominal_recirculation_MLSS_taux_f As Double
Dim reel_MES_concentration_f As Double
Dim reel_MV_MES_f As Double
Dim reel_recirculation_MLSS_taux_f As Double
Dim pre_anoxie_volume_f As Double
Dim anaerobie_volume_f As Double
Dim anaerobie_FeCl3_flux_f As Double
Dim pre_denit_volume_f As Double
Dim deox_O2_dissous_f As Double
Dim deox_volume_f As Double
Dim reox_O2_dissous_f As Double
Dim reox_volume_f As Double
Dim reox_O2_facteur_alfa_f As Double
Dim reox_O2_rdt_transfert_f As Double
Dim reox_hauteur_bassin_f As Double
Dim post_denit_methanol_flux_f As Double
Dim post_denit_media_filling_f As Double
Dim post_denit_volume_f As Double
Dim ratio_Nnit_media_Nnit_total_f As Double
Dim ratio_Nnit_mediaH1_Nnit_media_f As Double
Dim ratio_Nnit_CN1_Nnit_CN12_f As Double
Dim cuve_aeree_filling_f() As Double
Dim cuve_aeree_O2_dissous_f() As Double
Dim cuve_aeree_volume_f() As Double
Dim cuve_aeree_O2_facteur_alfa_f() As Double
Dim cuve_aeree_O2_rdt_transfert_f() As Double
Dim cuve_aeree_hauteur_bassin_f() As Double
Dim outlet_reel_DBO_f As Double
Dim outlet_reel_NH4_f As Double
Dim outlet_reel_NO3_f As Double
Dim ratio_elec_anaerobie_f As Double
Dim ratio_elec_pre_anoxie_f As Double
Dim ratio_elec_pre_denit_f As Double
Dim ratio_elec_deox_f As Double
Dim ratio_elec_post_denit_f As Double
Dim O2_besoin_total_f As Double
Dim air_Q_Nm3j_f As Double
Dim diffuseur_encrassement_f As Double
Dim air_P_refoulement_moyenne_f As Double
Dim surpresseur_conso_spec_f As Double
Dim recirculation_P_refoulement_f As Double
Dim recirculation_pompe_rdt_f As Double
Dim recirculation_MLSS_P_refoulement_f As Double
Dim recirculation_MLSS_pompe_rdt_f As Double
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

'repérage du nombre de cuves de chaque type (configuration centrale)
Dim nb_cuves_CN As Integer
Dim nb_cuves_H As Integer

'hypothèses
Dim delta_NO3_post_DN As Double
Dim ratio_elec_recirculation_MLSS As Double
Dim ratio_elec_recirculation As Double
Dim ratio_elec_extraction As Double
Dim a_ratio_EB_MV_MES(2) As Double
Dim ratio_EB_MV_MES_min As Double
Dim ratio_EB_MV_MES_max As Double
Dim ratio_ED_MV_MES As Double
Dim ratio_N_assimile As Double
Dim ratio_N_assimile_methanol As Double
Dim ratio_boues_MESin As Double
Dim ratio_boues_nit As Double
Dim ratio_boues_DBOin As Double
Dim recirculation_O2 As Double
Dim ratio_NO3eq_O2 As Double
Dim pre_anoxie_NO3 As Double
Dim ratio_DBO_denit_max As Double
Dim ratio_O2_denit As Double
Dim anaerobie_HRT_P_high As Double
Dim anaerobie_HRT_P_low As Double
Dim anaerobie_P_limite As Double
Dim ratio_P_assimile As Double
Dim anaerobie_DBO_Pt As Double
Dim MM_FeO3H3 As Double
Dim MM_FePO4 As Double
Dim MM_P As Double
Dim MM_FeCl3 As Double
Dim deox_HRT_mini As Double
Dim ratio_O2_nit As Double
Dim reox_HRT As Double
Dim rate_DCO_apportee_denit_nominal As Double
Dim rate_DCO_apportee_denit_reel As Double
Dim rate_DCO_methanol As Double
Dim a_cinetique_postDN_nominal As Double
Dim a_cinetique_postDN_reel As Double
Dim post_DN_SRT_mini As Double
Dim hyp_cinetique_post_denit_15 As Double
Dim cinetique_post_denit_15_mini As Double
Dim a_cinetique_postDN_media_10°C(2) As Double
Dim a_cinetique_postDN_media_nominal As Double
Dim a_cinetique_postDN_media_reel As Double
Dim media_surface_m2_m3(nb_medias) As Double
Dim media_filling_max(nb_medias) As Double
Dim media_filling_max_PDN(nb_medias) As Double
Dim media_deplacement(nb_medias) As Double
Dim rate_N_30(nb_medias) As Double
Dim ratio_boues_DCO_methanol As Double
Dim ratio_DCO_dure_total As Double
Dim charge_radier_max(nb_type_racleur) As Double
Dim facteur_boues_C_lim As Double
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
Dim agitation_fct As Double
Dim ratio_O2_DBO As Double
Dim ratio_Norgaduresoluble_DCOtot As Double
Dim cuve_H_V_Nmh_mini As Double

'Variable pour le Mox
Dim rate_N_T_design As Double
Dim rate_N_T_design_force As Boolean
Dim rate_N_T_design_f As Double
Dim rate_N_T_reel As Double
Dim rate_N_T_reel_force As Boolean
Dim rate_N_T_reel_f As Double
Dim ratio_NO3f_nit As Double

'variables pour conserver valeur d'une itératio à l'autre
Dim reox_nit As Double
Dim DBO_sortie As Double
Dim boues_methanol As Double
Dim stockage_boues_pdtes_nominal As Double
Dim stockage_Q_reel As Double
Dim rate_bacteries_nit As Double
Dim stockage_cinetique_post_DN_15 As Double

Dim remarque(12) As String


Private Function air_Q_Nm3j_cuve(ByVal DOi As Double, ByVal O2_besoin_cuve As Double, ByVal temperature, ByVal volume_cuve, ByVal cuve_media As Boolean, ByVal hauteur_bassin, ByVal facteur_alfa, ByVal rdt_transfert)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    '************* calcul du débit d'air pour chaque cuve *****************
    Dim I As Integer

    Dim Patm As Double
    Dim delta_P As Double
    Dim O2_saturation_20_P As Double
    Dim O2_saturation_T_P As Double
    Dim O2_facteur_K As Double

    Dim surface_cuve As Double
    Dim cuve_H_V_Nmh_mini As Double    'vitesse minimale pour agitation

    'Calcul du coefficient K
    Patm = P0 * ((a0_Patm + a1_Patm * altitude) ^ a2_Patm) / 100    'en hPa
    delta_P = a0_delta_P / a1_delta_P * P0 / 100 * (hauteur_bassin - hauteur_diffuseur)
    O2_saturation_20_P = 0
    O2_saturation_T_P = 0
    For I = 0 To 4
        O2_saturation_20_P = O2_saturation_20_P + a_O2_saturation(I) * 20 ^ I
        O2_saturation_T_P = O2_saturation_T_P + a_O2_saturation(I) * temperature ^ I
    Next I
    O2_saturation_T_P = O2_saturation_T_P * (Patm + delta_P) / (P0 / 100)
    O2_saturation_20_P = O2_saturation_20_P * (Patm + delta_P) / (P0 / 100)

    O2_facteur_K = facteur_alfa * (O2_facteur_beta * O2_saturation_T_P - DOi) / O2_saturation_20_P
    O2_facteur_K = O2_facteur_K * correctif_T_K ^ (temperature - T_ref_K)


    'calcul consommations sur l'aération
    air_Q_Nm3j_cuve = O2_besoin_cuve / (O2_facteur_K * rdt_transfert / 100 * (hauteur_bassin - hauteur_diffuseur) * ratio_kgO2_Nm3air)

    If cuve_media = True Then
        surface_cuve = volume_cuve / hauteur_bassin
        If air_Q_Nm3j_cuve / pd_NOMBRE_HEURE_PAR_JOUR < cuve_H_V_Nmh_mini * surface_cuve Then
            air_Q_Nm3j_cuve = cuve_H_V_Nmh_mini * surface_cuve * pd_NOMBRE_HEURE_PAR_JOUR
            remarque(erreur_aeration_augmentee_agitation) = Feuil15.Range("HybAS_aeration_augmentee_agitation")
        End If
    End If


    Exit Function
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.11"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Function




Private Function anaerobie_Kp(ByVal anaerobie_V_Voptimal, ByVal anaerobie_NO3in_concentration, ByVal interm_Pt, ByVal DBO) As Double

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Kp_bio_1 As Double
    Dim Kp_bio_2 As Double
    Dim ratio_DBOsol_Pt As Double

    'calcul de Kp_bio_1
    If anaerobie_V_Voptimal <= 1 Then
        Kp_bio_1 = anaerobie_V_Voptimal ^ 2
    ElseIf anaerobie_V_Voptimal <= 1.5 Then
        Kp_bio_1 = 0.7 + 0.3 * anaerobie_V_Voptimal
    Else
        Kp_bio_1 = 1.15
    End If

    'calcul de Kp_bio_2
    If anaerobie_NO3in_concentration >= 0.3 Then
        Kp_bio_2 = (0.3 / anaerobie_NO3in_concentration) ^ 2
    Else
        Kp_bio_2 = 1.1 - anaerobie_NO3in_concentration / 3
    End If

    'ratio DBOsoluble/Pt   hyp: DBOsol=DBOtotale vu avec Kim Sorensen
    ratio_DBOsol_Pt = DBO / interm_Pt

    'calcul de Kp
    anaerobie_Kp = 0.028 * Kp_bio_1 * Kp_bio_2 * ratio_DBOsol_Pt



    Exit Function
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.12"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Function

Private Function cinetique_denit(ByVal ratio_DBO_denit, ByVal temperature) As Double

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim ratio_DBO_denit_max As Double
    Dim ratio_DBO_denit_limite As Double   'changement de pente

    ratio_DBO_denit_max = 15    ' au dessus de ce ratio, le rate est constant dans chacun des cas étudiés

    If ratio_DBO_denit > ratio_DBO_denit_max Then
        ratio_DBO_denit = ratio_DBO_denit_max
    End If


    If choix_decanteur_simple = True Then
        ratio_DBO_denit_limite = 5
        If ratio_DBO_denit < ratio_DBO_denit_limite Then
            cinetique_denit = (0.8333 * ratio_DBO_denit - 1.6667) / 0.77
        Else
            cinetique_denit = (0.8333 * ratio_DBO_denit_limite - 1.6667 + 0.05 * (ratio_DBO_denit - ratio_DBO_denit_limite)) / 0.77
        End If
    ElseIf choix_decanteur_reactif = True Then
        ratio_DBO_denit_limite = 4
        If ratio_DBO_denit < ratio_DBO_denit_limite Then
            cinetique_denit = (1.25 * ratio_DBO_denit - 2.5) / 0.77
        Else
            cinetique_denit = (1.25 * ratio_DBO_denit_limite - 2.5 + 0.0455 * (ratio_DBO_denit - ratio_DBO_denit_limite)) / 0.77
        End If
    Else
        ratio_DBO_denit_limite = 6.5
        If ratio_DBO_denit < ratio_DBO_denit_limite Then
            cinetique_denit = (0.5556 * ratio_DBO_denit - 1.1111) / 0.77
        Else
            cinetique_denit = (0.5556 * ratio_DBO_denit_limite - 1.1111 + 0.0588 * (ratio_DBO_denit - ratio_DBO_denit_limite)) / 0.77
        End If
    End If

    'correction en température
    cinetique_denit = cinetique_denit * 1.08 ^ (temperature - 15)



    Exit Function
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.13"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Function


Private Function cinetique_nit_MLSS(ByVal DOi, ByVal boues_pdtes, ByVal N_nit_aerees, ByVal temperature) As Double

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    'DOi est la concentration en O2 dissous dans la cuve

    If DOi <= 2.7 Then
        cinetique_nit_MLSS = 63 * (DOi / (0.7 + DOi)) * rate_bacteries_nit * 1.103 ^ (temperature - 15)
    Else
        cinetique_nit_MLSS = 50 * rate_bacteries_nit * 1.103 ^ (temperature - 15)
    End If



    Exit Function
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.14"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Function

Sub dimensionne_config_3(ByVal boues_pdtes, ByVal N_nit_aerees, ByVal temperature, ByVal cuves_centrales_nit, ByVal cuves_centrales_DBOin, ByRef cuve_aeree_O2_besoins)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'DIMENSIONNE    C/N - H - H

    Dim I As Integer
    Dim j As Integer
    Dim epsilone_global As Double

    Dim rate_N_MLSS(3) As Double
    Dim Nnit_media(3) As Double
    Dim Nnit_media_total As Double
    Dim k_factor As Double
    Dim DBO_appliquee As Double
    Dim V_total As Double
    Dim nit_MLSS_restant As Double
    Dim epsilone_interm As Double

    Nnit_media_total = N_nit_aerees * ratio_Nnit_media_Nnit_total

    Nnit_media(2) = ratio_Nnit_mediaH1_Nnit_media * Nnit_media_total
    Nnit_media(3) = (1 - ratio_Nnit_mediaH1_Nnit_media) * Nnit_media_total


    'on calcule la cinétique de nit MLSS
    For I = 1 To choix_nb_cuves_aerees
        rate_N_MLSS(I) = cinetique_nit_MLSS(cuve_aeree_O2_dissous(I), boues_pdtes, N_nit_aerees, temperature)
    Next I

    'on commence les boucles
    DBO_appliquee = 0.5
    epsilone_global = critere_convergence + 1
    Do While epsilone_global > critere_convergence
        V_total = 0
        epsilone_global = 0
        nit_MLSS_restant = cuves_centrales_nit - Nnit_media_total
        For I = 2 To 3
            epsilone_interm = cuve_aeree_volume(I)
            cuve_aeree_volume(I) = volume_cuve_H(I, Nnit_media(I), DBO_appliquee, temperature)
            epsilone_interm = epsilone_interm - cuve_aeree_volume(I)
            epsilone_interm = epsilone_interm ^ 2 / cuve_aeree_volume(I)
            epsilone_global = epsilone_global + epsilone_interm
            V_total = V_total + cuve_aeree_volume(I)
            nit_MLSS_restant = nit_MLSS_restant - rate_N_MLSS(I) * cuve_aeree_volume(I) * (1 - cuve_aeree_filling(I) * media_deplacement(choix_media_cuve_aeree(I))) * pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration / 1000
        Next I
        If (nit_MLSS_restant < 0) Then
            If ratio_Nnit_media_Nnit_total_force = True Then
                MsgBox Feuil15.Range("HybAS_nit_media_trop"), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub
            Else
                ratio_Nnit_media_Nnit_total = ratio_Nnit_media_Nnit_total * 0.9
                Nnit_media_total = N_nit_aerees * ratio_Nnit_media_Nnit_total
                Nnit_media(2) = ratio_Nnit_mediaH1_Nnit_media * Nnit_media_total
                Nnit_media(3) = (1 - ratio_Nnit_mediaH1_Nnit_media) * Nnit_media_total
                epsilone_global = critere_convergence + 1
            End If
        End If
        epsilone_interm = cuve_aeree_volume(1)
        cuve_aeree_volume(1) = 1000 * nit_MLSS_restant / (pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration * rate_N_MLSS(1))
        epsilone_interm = epsilone_interm - cuve_aeree_volume(1)
        epsilone_interm = epsilone_interm ^ 2 / cuve_aeree_volume(1)
        epsilone_global = epsilone_global + epsilone_interm
        V_total = V_total + cuve_aeree_volume(1)
        DBO_appliquee = cuves_centrales_DBOin / (V_total * nominal_MES_concentration)  ' * nominal_MV_MES)
    Loop

    For I = 1 To choix_nb_cuves_aerees
        If choix_type_cuve_aeree(I) = cuve_aeree_H Then
            cuve_aeree_O2_besoins(I) = ratio_O2_nit * (Nnit_media(I) + rate_N_MLSS(I) * cuve_aeree_volume(I) * (1 - cuve_aeree_filling(I) * media_deplacement(choix_media_cuve_aeree(I))) * pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration / 1000)
        Else
            cuve_aeree_O2_besoins(I) = ratio_O2_nit * (rate_N_MLSS(I) * cuve_aeree_volume(I) * pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration / 1000)
        End If
    Next I



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.17"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub dimensionne_config_2(ByVal boues_pdtes, ByVal N_nit_aerees, ByVal temperature, ByVal cuves_centrales_nit, ByVal cuves_centrales_DBOin, ByRef cuve_aeree_O2_besoins)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'DIMENSIONNE    C/N - H - C/N

    Const nb_cuves_utilisees = 3

    Dim I As Integer
    Dim j As Integer
    Dim epsilone_global As Double

    Dim rate_N_MLSS(nb_cuves_utilisees) As Double
    Dim Nnit_media(nb_cuves_utilisees) As Double
    Dim Nnit_media_total As Double
    Dim k_factor As Double
    Dim DBO_appliquee As Double
    Dim V_total As Double
    Dim nit_MLSS_restant As Double
    Dim epsilone_interm As Double
    Dim ratio_nit_MLSS_CNi_CNtotal(nb_cuves_utilisees) As Double


    Nnit_media_total = N_nit_aerees * ratio_Nnit_media_Nnit_total

    ratio_nit_MLSS_CNi_CNtotal(1) = ratio_Nnit_CN1_Nnit_CN12
    ratio_nit_MLSS_CNi_CNtotal(3) = 1 - ratio_Nnit_CN1_Nnit_CN12

    'on calcule la cinétique de nit MLSS
    For I = 1 To choix_nb_cuves_aerees
        rate_N_MLSS(I) = cinetique_nit_MLSS(cuve_aeree_O2_dissous(I), boues_pdtes, N_nit_aerees, temperature)
    Next I

    'on commence les boucles
    DBO_appliquee = 0.5
    epsilone_global = critere_convergence + 1
    Do While epsilone_global > critere_convergence
        'initialisation des variables
        V_total = 0
        epsilone_global = 0
        nit_MLSS_restant = cuves_centrales_nit - Nnit_media_total
        'calcul de la cuve H
        epsilone_interm = cuve_aeree_volume(2)
        cuve_aeree_volume(2) = volume_cuve_H(2, Nnit_media_total, DBO_appliquee, temperature)
        epsilone_interm = epsilone_interm - cuve_aeree_volume(2)
        epsilone_interm = epsilone_interm ^ 2 / cuve_aeree_volume(2)
        epsilone_global = epsilone_global + epsilone_interm
        V_total = V_total + cuve_aeree_volume(2)
        nit_MLSS_restant = nit_MLSS_restant - rate_N_MLSS(2) * cuve_aeree_volume(2) * (1 - cuve_aeree_filling(2) * media_deplacement(choix_media_cuve_aeree(2))) * pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration / 1000
        If (nit_MLSS_restant < 0) Then
            If ratio_Nnit_media_Nnit_total_force = True Then
                MsgBox Feuil15.Range("HybAS_nit_media_trop"), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub
            Else
                ratio_Nnit_media_Nnit_total = ratio_Nnit_media_Nnit_total * 0.9
                Nnit_media_total = N_nit_aerees * ratio_Nnit_media_Nnit_total
                epsilone_global = critere_convergence + 1
            End If
        End If
        'calcul des cuves C/N
        For I = 1 To nb_cuves_utilisees
            If choix_type_cuve_aeree(I) = cuve_aeree_CN Then
                epsilone_interm = cuve_aeree_volume(I)
                cuve_aeree_volume(I) = 1000 * nit_MLSS_restant * ratio_nit_MLSS_CNi_CNtotal(I) / (pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration * rate_N_MLSS(I))
                epsilone_interm = epsilone_interm - cuve_aeree_volume(I)
                epsilone_interm = epsilone_interm ^ 2 / cuve_aeree_volume(I)
                epsilone_global = epsilone_global + epsilone_interm
                V_total = V_total + cuve_aeree_volume(I)
            End If
        Next I
        DBO_appliquee = cuves_centrales_DBOin / (V_total * nominal_MES_concentration)  ' * nominal_MV_MES)
    Loop

    'besoins en O2
    cuve_aeree_O2_besoins(1) = ratio_O2_nit * (rate_N_MLSS(1) * cuve_aeree_volume(1) * pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration / 1000)
    cuve_aeree_O2_besoins(2) = ratio_O2_nit * (Nnit_media_total + rate_N_MLSS(2) * cuve_aeree_volume(2) * (1 - cuve_aeree_filling(2) * media_deplacement(choix_media_cuve_aeree(2))) * pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration / 1000)
    cuve_aeree_O2_besoins(3) = ratio_O2_nit * (rate_N_MLSS(3) * cuve_aeree_volume(3) * pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration / 1000)




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.16"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub dimensionne_config_1_5(ByVal boues_pdtes, ByVal N_nit_aerees, ByVal temperature, ByVal cuves_centrales_nit, ByVal cuves_centrales_DBOin, ByRef cuve_aeree_O2_besoins)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'DIMENSIONNE    C/N - H ou H - C/N

    Const nb_cuves_utilisees = 2

    Dim I As Integer
    Dim j As Integer
    Dim epsilone_global As Double

    Dim rate_N_MLSS(nb_cuves_utilisees) As Double
    Dim Nnit_media(nb_cuves_utilisees) As Double
    Dim Nnit_media_total As Double
    Dim k_factor As Double
    Dim DBO_appliquee As Double
    Dim V_total As Double
    Dim nit_MLSS_restant As Double
    Dim epsilone_interm As Double
    Dim ratio_nit_MLSS_CNi_CNtotal(nb_cuves_utilisees) As Double

xlt01:
    Nnit_media_total = N_nit_aerees * ratio_Nnit_media_Nnit_total

    'on calcule la cinétique de nit MLSS
    For I = 1 To nb_cuves_utilisees
        rate_N_MLSS(I) = cinetique_nit_MLSS(cuve_aeree_O2_dissous(I), boues_pdtes, N_nit_aerees, temperature)
    Next I

    'on commence les boucles
    DBO_appliquee = 0.5
    epsilone_global = critere_convergence + 1
    Do While epsilone_global > critere_convergence
        'initialisation des variables
        V_total = 0
        epsilone_global = 0
        nit_MLSS_restant = cuves_centrales_nit - Nnit_media_total  'on le retire ici
        'calcul de la cuve H
        For I = 1 To nb_cuves_utilisees
            If choix_type_cuve_aeree(I) = cuve_aeree_H Then
                epsilone_interm = cuve_aeree_volume(I)
                cuve_aeree_volume(I) = volume_cuve_H(I, Nnit_media_total, DBO_appliquee, temperature)
                epsilone_interm = epsilone_interm - cuve_aeree_volume(I)
                epsilone_interm = epsilone_interm ^ 2 / cuve_aeree_volume(I)
                epsilone_global = epsilone_global + epsilone_interm
                V_total = V_total + cuve_aeree_volume(I)
                nit_MLSS_restant = nit_MLSS_restant - rate_N_MLSS(I) * cuve_aeree_volume(I) * (1 - cuve_aeree_filling(I) * media_deplacement(choix_media_cuve_aeree(I))) * pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration / 1000
                If (nit_MLSS_restant < 0) Then
                    If ratio_Nnit_media_Nnit_total_force = True Then
                        MsgBox Feuil15.Range("HybAS_nit_media_trop"), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                        pn_EffacementResultatsEnCours = True
                        pn_ErreurNonGeree = True
                        Exit Sub
                    Else
                        ratio_Nnit_media_Nnit_total = ratio_Nnit_media_Nnit_total * 0.9
                        Nnit_media_total = N_nit_aerees * ratio_Nnit_media_Nnit_total
                        epsilone_global = critere_convergence + 1
                    End If
                    ' GoTo xlt01
                End If
            End If
        Next I
        'calcul des cuves C/N
        For I = 1 To nb_cuves_utilisees
            If choix_type_cuve_aeree(I) = cuve_aeree_CN Then
                epsilone_interm = cuve_aeree_volume(I)
                cuve_aeree_volume(I) = 1000 * nit_MLSS_restant / (pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration * rate_N_MLSS(I))
                epsilone_interm = epsilone_interm - cuve_aeree_volume(I)
                epsilone_interm = epsilone_interm ^ 2 / cuve_aeree_volume(I)
                epsilone_global = epsilone_global + epsilone_interm
                V_total = V_total + cuve_aeree_volume(I)
            End If
        Next I
        DBO_appliquee = cuves_centrales_DBOin / (V_total * nominal_MES_concentration)   '* nominal_MV_MES)
    Loop

    For I = 1 To choix_nb_cuves_aerees
        If choix_type_cuve_aeree(I) = cuve_aeree_H Then
            cuve_aeree_O2_besoins(I) = ratio_O2_nit * (Nnit_media_total + rate_N_MLSS(I) * cuve_aeree_volume(I) * (1 - cuve_aeree_filling(I) * media_deplacement(choix_media_cuve_aeree(I))) * pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration / 1000)
        Else
            cuve_aeree_O2_besoins(I) = ratio_O2_nit * (rate_N_MLSS(I) * cuve_aeree_volume(I) * pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration / 1000)
        End If
    Next I




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.15"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub dimensionne_config_4_6(ByVal boues_pdtes, ByVal N_nit_aerees, ByVal temperature, ByVal cuves_centrales_nit, ByVal cuves_centrales_DBOin, ByRef cuve_aeree_O2_besoins)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'DIMENSIONNE    H  et  H - H

    Dim I As Integer
    Dim j As Integer

    Dim nit_cuve As Double

    Dim V_total_min As Double
    Dim V_total_max As Double
    Dim rate_N_MLSS() As Double
    Dim ratio_volume_somme As Double
    Dim ratio_volume_somme_autres As Double
    Dim ratio_nit_media() As Double
    Dim ratio_volume_i_total() As Double
    Dim rate_N_MLSS_moyen As Double
    Dim V_total As Double
    Dim volume_media_total As Double
    Dim DBO_appliquee As Double
    Dim k_factor As Double
    Dim nit_total As Double
    Dim surface_rate As Double
    Dim interm_O2_besoin_cuve() As Double
    ReDim interm_O2_besoin_cuve(choix_nb_cuves_aerees)
    ReDim rate_N_MLSS(choix_nb_cuves_aerees)
    ReDim ratio_cinetique_media(choix_nb_cuves_aerees)
    ReDim ratio_nit_media(2)  'pour le moment, on ne peut avoir que 2 cuves au maximum
    ReDim ratio_volume_i_total(choix_nb_cuves_aerees)

    ratio_nit_media(1) = ratio_Nnit_mediaH1_Nnit_media
    ratio_nit_media(2) = 1 - ratio_Nnit_mediaH1_Nnit_media

    ratio_volume_somme = 0
    For I = 1 To choix_nb_cuves_aerees
        rate_N_MLSS(I) = cinetique_nit_MLSS(cuve_aeree_O2_dissous(I), boues_pdtes, N_nit_aerees, temperature)
        'calculé à partir des caractéristiques influençant la cinétique sur le média
        If choix_fct_Mox = False Then
            ratio_cinetique_media(I) = ratio_nit_media(I) * media_surface_m2_m3(choix_media_cuve_aeree(I)) * cuve_aeree_filling(I) * (cuve_aeree_O2_dissous(I) - 1) ^ 0.7
            ratio_volume_somme = ratio_volume_somme + ratio_cinetique_media(I)
        End If
    Next I
    If choix_nb_cuves_aerees > 1 Then
        For I = 1 To choix_nb_cuves_aerees
            ratio_volume_somme_autres = 0
            For j = 1 To choix_nb_cuves_aerees   'on fait une relation générale si jamais un jour les configurations sont modifiées
                If j <> I Then
                    ratio_volume_somme_autres = ratio_volume_somme_autres + ratio_cinetique_media(j)
                End If
            Next j
            ratio_volume_i_total(I) = ratio_volume_somme_autres / ratio_volume_somme
        Next I
    Else
        ratio_volume_i_total(1) = 1
    End If


    V_total_min = 1
    'on calcule le volume que l'on aurait besoin sans média et il nous servira de maximum
    rate_N_MLSS_moyen = 0
    For I = 1 To choix_nb_cuves_aerees
        rate_N_MLSS_moyen = rate_N_MLSS_moyen + ratio_volume_i_total(I) * rate_N_MLSS(I)
    Next I
    V_total_max = 1000 * cuves_centrales_nit / pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration * rate_N_MLSS_moyen

    'on commence à faire la dichotomie
    V_total = (V_total_min + V_total_max) / 2
    Do While (V_total_max - V_total_min) / V_total > critere_convergence
        For I = 1 To choix_nb_cuves_aerees
            cuve_aeree_volume(I) = ratio_volume_i_total(I) * V_total
        Next I
        volume_media_total = 0
        For I = 1 To choix_nb_cuves_aerees
            volume_media_total = volume_media_total + cuve_aeree_filling(I) * cuve_aeree_volume(I) * media_deplacement(choix_media_cuve_aeree(I))
        Next I
        DBO_appliquee = cuves_centrales_DBOin / (nominal_MES_concentration * (V_total - volume_media_total + deox_volume + reox_volume))  '*nominal_MV_MES)
        'on calcule la nit pour l'ensemble des cuves
        nit_total = 0
        ratio_Nnit_media_Nnit_total = 0
        For I = 1 To choix_nb_cuves_aerees
            nit_cuve = nitrification_H_kgj(DBO_appliquee, temperature, I, rate_N_MLSS(I), nominal_MES_concentration, True)
            nit_total = nit_total + nit_cuve
            'besoins en O2
            interm_O2_besoin_cuve(I) = ratio_O2_nit * nit_cuve
        Next I
        If nit_total < cuves_centrales_nit Then
            V_total_min = V_total
        Else
            V_total_max = V_total
        End If
        V_total = (V_total_min + V_total_max) / 2
    Loop

    For I = 1 To choix_nb_cuves_aerees
        cuve_aeree_O2_besoins(I) = interm_O2_besoin_cuve(I)
    Next I




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.18"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Function volume_cuve_H(ByVal position_cuve, ByVal Nnit_media, ByVal DBO_appliquee, ByVal temperature)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim DBO_appliquee_10°C As Double
    Dim k_factor As Double
    Dim surface_rate As Double

    If choix_fct_Mox = True Then
        If rate_N_T_design_force = True Then
            rate_N_T_design = rate_N_T_design_f
        Else
            rate_N_T_design = rate_N_30(choix_media_cuve_aeree(position_cuve)) * 1.07 ^ (temperature - 30)
        End If
        'calcul du volume de média
        volume_cuve_H = 1000 * Nnit_media / (rate_N_T_design * media_surface_m2_m3(choix_media_cuve_aeree(position_cuve)))
        'calcul du volume total
        volume_cuve_H = volume_cuve_H / cuve_aeree_filling(position_cuve)
    Else
        DBO_appliquee_10°C = DBO_appliquee * 1.06 ^ (10 - temperature)    'on calcule l'équivalent à 10°C
        If DBO_appliquee_10°C < 0.13 Then
            k_factor = 0.55
        ElseIf DBO_appliquee_10°C < 0.48 Then
            k_factor = -0.11 / 0.35 * DBO_appliquee_10°C + 0.55 + 0.0143 / 0.35
        Else
            k_factor = 0.44 * Exp(-1000 * (DBO_appliquee_10°C - 0.48) ^ 2)
        End If
        surface_rate = k_factor * ((cuve_aeree_O2_dissous(position_cuve) - 1) / 3.2) ^ 0.7 * 1.06 ^ (temperature - 10)
        'calcul du volume de média
        volume_cuve_H = 1000 * Nnit_media / (surface_rate * media_surface_m2_m3(choix_media_cuve_aeree(position_cuve)))
        'calcul du volume total
        volume_cuve_H = volume_cuve_H / cuve_aeree_filling(position_cuve)
    End If


    Exit Function
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.22"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Function

Private Function nitrification_H_kgj(ByVal DBO_appliquee, ByVal temperature, ByVal position_cuve, ByVal rate_N_MLSS, ByVal MES_concentration, ByVal design As Boolean)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim DBO_appliquee_10°C As Double
    Dim k_factor As Double
    Dim surface_rate As Double


    If choix_fct_Mox = True Then
        If design = False Then
            If rate_N_T_reel_force = True Then
                rate_N_T_reel = rate_N_T_reel_f
            Else
                rate_N_T_reel = rate_N_30(choix_media_cuve_aeree(position_cuve)) * 1.07 ^ (temperature - 30)
            End If
            nitrification_H_kgj = rate_N_T_reel / 1000 * media_surface_m2_m3(choix_media_cuve_aeree(position_cuve)) * cuve_aeree_volume(position_cuve) * cuve_aeree_filling(position_cuve)
        Else
            If rate_N_T_design_force = True Then
                rate_N_T_design = rate_N_T_design_f
            Else
                rate_N_T_design = rate_N_30(choix_media_cuve_aeree(position_cuve)) * 1.07 ^ (temperature - 30)
            End If
            nitrification_H_kgj = rate_N_T_design / 1000 * media_surface_m2_m3(choix_media_cuve_aeree(position_cuve)) * cuve_aeree_volume(position_cuve) * cuve_aeree_filling(position_cuve)
            ratio_Nnit_media_Nnit_total = ratio_Nnit_media_Nnit_total + rate_N_T_design / 1000 * media_surface_m2_m3(choix_media_cuve_aeree(position_cuve)) * cuve_aeree_volume(position_cuve) * cuve_aeree_filling(position_cuve)
        End If
    Else
        'on calcule la charge appliquée équivalente pour obtenir la nit sur média
        DBO_appliquee_10°C = DBO_appliquee * 1.06 ^ (10 - temperature)    'on calcule l'équivalent à 10°C
        If DBO_appliquee_10°C < 0.13 Then
            k_factor = 0.55
        ElseIf DBO_appliquee_10°C < 0.48 Then
            k_factor = -0.11 / 0.35 * DBO_appliquee_10°C + 0.55 + 0.0143 / 0.35
        Else
            k_factor = 0.44 * Exp(-1000 * (DBO_appliquee_10°C - 0.48) ^ 2)
        End If
        surface_rate = k_factor * ((cuve_aeree_O2_dissous(position_cuve) - 1) / 3.2) ^ 0.7 * 1.06 ^ (temperature - 10)
        nitrification_H_kgj = surface_rate * media_surface_m2_m3(choix_media_cuve_aeree(position_cuve)) * cuve_aeree_volume(position_cuve) * cuve_aeree_filling(position_cuve) / 1000
        If nb_cuves_CN = 0 And design = True Then
            ratio_Nnit_media_Nnit_total = ratio_Nnit_media_Nnit_total + surface_rate * media_surface_m2_m3(choix_media_cuve_aeree(position_cuve)) * cuve_aeree_volume(position_cuve) * cuve_aeree_filling(position_cuve) / 1000
        End If
    End If
    

    'on calcule la nitrification MLSS
    nitrification_H_kgj = nitrification_H_kgj + rate_N_MLSS * cuve_aeree_volume(position_cuve) * (1 - cuve_aeree_filling(position_cuve) * media_deplacement(choix_media_cuve_aeree(position_cuve))) * pd_NOMBRE_HEURE_PAR_JOUR * MES_concentration / 1000




    Exit Function
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.19"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Function


Private Function rdt_DBO(ByVal DBO_appliquee_kgDBO_kgMES_j)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'Calcul du rendement d'élimination de la DBO sur les cuves aérées

    If choix_decanteur_simple = True Then
        rdt_DBO = 25 + 71 * Exp(-0.6 * (DBO_appliquee_kgDBO_kgMES_j + 0.1) ^ 2)
    ElseIf choix_decanteur_reactif = True Then
        rdt_DBO = 14 + 80 * Exp(-0.6 * (DBO_appliquee_kgDBO_kgMES_j + 0.1) ^ 2)
    Else
        rdt_DBO = 44.57 + 53 * Exp(-0.75 * (DBO_appliquee_kgDBO_kgMES_j + 0.1) ^ 2)
    End If

    rdt_DBO = rdt_DBO / 100



    Exit Function
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.21"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Function

Private Function respiration_gL(ByVal temperature, ByVal MLSS) As Double

    respiration_gL = 0.002 * MLSS * pd_NOMBRE_HEURE_PAR_JOUR * 1.07 ^ (temperature - 10)


End Function

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil15.Range("HybAS_outlet_NH4").Cells(1, Colonne) = outlet_NH4
    Feuil15.Range("HybAS_outlet_NO3").Cells(1, Colonne) = outlet_NO3
    Feuil15.Range("HybAS_Norga_dure_soluble").Cells(1, Colonne) = Norga_dure_soluble
    Feuil15.Range("HybAS_outlet_DBO_soluble").Cells(1, Colonne) = outlet_DBO_soluble
    Feuil15.Range("HybAS_recirculation_taux").Cells(1, Colonne) = recirculation_taux
    Feuil15.Range("HybAS_nominal_MES_concentration").Cells(1, Colonne) = nominal_MES_concentration
    Feuil15.Range("HybAS_nominal_MV_MES").Cells(1, Colonne) = nominal_MV_MES
    Feuil15.Range("HybAS_nominal_recirculation_MLSS_taux").Cells(1, Colonne) = nominal_recirculation_MLSS_taux
    Feuil15.Range("HybAS_reel_MES_concentration").Cells(1, Colonne) = reel_MES_concentration
    Feuil15.Range("HybAS_reel_MV_MES").Cells(1, Colonne) = reel_MV_MES
    Feuil15.Range("HybAS_reel_recirculation_MLSS_taux").Cells(1, Colonne) = reel_recirculation_MLSS_taux
    If choix_pre_anoxie <> cuve_non Then
        Feuil15.Range("HybAS_pre_anoxie_volume").Cells(1, Colonne) = pre_anoxie_volume
        Feuil15.Range("HybAS_ratio_elec_pre_anoxie").Cells(1, Colonne) = ratio_elec_pre_anoxie
    End If
    If choix_anaerobie <> cuve_non Then
        Feuil15.Range("HybAS_anaerobie_volume").Cells(1, Colonne) = anaerobie_volume
        Feuil15.Range("HybAS_anaerobie_FeCl3_flux").Cells(1, Colonne) = anaerobie_FeCl3_flux
        Feuil15.Range("HybAS_ratio_elec_anaerobie").Cells(1, Colonne) = ratio_elec_anaerobie
    End If
    If choix_pre_denit <> cuve_non Then
        Feuil15.Range("HybAS_pre_denit_volume").Cells(1, Colonne) = pre_denit_volume
        Feuil15.Range("HybAS_pre_denit_NO3").Cells(1, Colonne) = pre_denit_NO3
        Feuil15.Range("HybAS_ratio_elec_pre_denit").Cells(1, Colonne) = ratio_elec_pre_denit
    End If
    If choix_deox <> cuve_non Then
        Feuil15.Range("HybAS_deox_O2_dissous").Cells(1, Colonne) = deox_O2_dissous
        Feuil15.Range("HybAS_deox_volume").Cells(1, Colonne) = deox_volume
        Feuil15.Range("HybAS_ratio_elec_deox").Cells(1, Colonne) = ratio_elec_deox
    End If
    If choix_reox <> cuve_non Then
        Feuil15.Range("HybAS_reox_O2_dissous").Cells(1, Colonne) = reox_O2_dissous
        Feuil15.Range("HybAS_reox_volume").Cells(1, Colonne) = reox_volume
        Feuil15.Range("HybAS_reox_O2_alpha").Cells(1, Colonne) = reox_O2_facteur_alfa
        Feuil15.Range("HybAS_reox_O2_rdt_transfert").Cells(1, Colonne) = reox_O2_rdt_transfert
        Feuil15.Range("HybAS_reox_hauteur_bassin").Cells(1, Colonne) = reox_hauteur_bassin
    End If
    If choix_post_denit <> cuve_non Then
        Feuil15.Range("HybAS_NO3_last_aerated_tank").Cells(1, Colonne) = NO3_last_aerated_tank
        Feuil15.Range("HybAS_post_denit_methanol_flux").Cells(1, Colonne) = post_denit_methanol_flux
        Feuil15.Range("HybAS_post_denit_media_filling").Cells(1, Colonne) = post_denit_media_filling
        Feuil15.Range("HybAS_post_denit_volume").Cells(1, Colonne) = post_denit_volume
        Feuil15.Range("HybAS_ratio_elec_post_denit").Cells(1, Colonne) = ratio_elec_post_denit
    End If
    Feuil15.Range("HybAS_ratio_Nnit_media_Nnit_total").Cells(1, Colonne) = ratio_Nnit_media_Nnit_total
    If nb_cuves_H = 2 Then
        Feuil15.Range("HybAS_ratio_Nnit_mediaH1_Nnit_media").Cells(1, Colonne) = ratio_Nnit_mediaH1_Nnit_media
    Else
        Feuil15.Range("HybAS_ratio_Nnit_mediaH1_Nnit_media").Cells(1, Colonne) = ""
    End If
    If nb_cuves_CN = 2 Then
        Feuil15.Range("HybAS_ratio_Nnit_CN1_Nnit_CN12").Cells(1, Colonne) = ratio_Nnit_CN1_Nnit_CN12
    Else
        Feuil15.Range("HybAS_ratio_Nnit_CN1_Nnit_CN12").Cells(1, Colonne) = ""
    End If
    Feuil15.Range("HybAS_outlet_reel_DBO").Cells(1, Colonne) = outlet_reel_DBO
    Feuil15.Range("HybAS_outlet_reel_NH4").Cells(1, Colonne) = outlet_reel_NH4
    Feuil15.Range("HybAS_outlet_reel_NO3").Cells(1, Colonne) = outlet_reel_NO3
    Feuil15.Range("HybAS_volume_total_bassins").Cells(1, Colonne) = volume_total_bassins
    Feuil15.Range("HybAS_O2_besoin_total").Cells(1, Colonne) = O2_besoin_total
    Feuil15.Range("HybAS_air_Q_Nm3j").Cells(1, Colonne) = air_Q_Nm3j
    Feuil15.Range("HybAS_diffuseur_encrassement").Cells(1, Colonne) = diffuseur_encrassement
    Feuil15.Range("HybAS_air_P_refoulement_moyenne").Cells(1, Colonne) = air_P_refoulement_moyenne
    Feuil15.Range("HybAS_surpresseur_conso_spec").Cells(1, Colonne) = surpresseur_conso_spec
    Feuil15.Range("HybAS_recirculation_P_refoulement").Cells(1, Colonne) = recirculation_P_refoulement
    Feuil15.Range("HybAS_recirculation_pompe_rdt").Cells(1, Colonne) = recirculation_pompe_rdt
    Feuil15.Range("HybAS_recirculation_MLSS_P_refoulement").Cells(1, Colonne) = recirculation_MLSS_P_refoulement
    Feuil15.Range("HybAS_recirculation_MLSS_pompe_rdt").Cells(1, Colonne) = recirculation_MLSS_pompe_rdt
    Feuil15.Range("HybAS_boues_indice_Mohlman").Cells(1, Colonne) = boues_indice_Mohlman
    Feuil15.Range("HybAS_sortie_MES").Cells(1, Colonne) = sortie_MES
    Feuil15.Range("HybAS_clarif_hauteur").Cells(1, Colonne) = clarif_hauteur
    Feuil15.Range("HybAS_clarif_vitesse_max").Cells(1, Colonne) = clarif_vitesse_max
    Feuil15.Range("HybAS_clarif_surface").Cells(1, Colonne) = clarif_surface
    Feuil15.Range("HybAS_clarif_vitesse_max_recalc").Cells(1, Colonne) = clarif_vitesse_max_recalc
    Feuil15.Range("HybAS_clarif_charge_radier_nominal").Cells(1, Colonne) = clarif_charge_radier_nominal
    Feuil15.Range("HybAS_boues_concentration").Cells(1, Colonne) = boues_concentration
    Feuil15.Range("HybAS_boues_MES").Cells(1, Colonne) = boues_MES
    Feuil15.Range("HybAS_boues_Q").Cells(1, Colonne) = boues_Q
    Feuil15.Range("HybAS_extraction_P_refoulement").Cells(1, Colonne) = extraction_P_refoulement
    Feuil15.Range("HybAS_extraction_pompe_rdt").Cells(1, Colonne) = extraction_pompe_rdt
    Feuil15.Range("HybAS_electricite").Cells(1, Colonne) = electricite
    Feuil15.Range("HybAS_electricite_aeration").Cells(1, Colonne) = electricite_aeration


    For I = 1 To choix_nb_cuves_aerees
        If choix_type_cuve_aeree(I) = cuve_aeree_H Then
            Feuil15.Range("HybAS_cuve_aeree_filling").Cells(I, Colonne) = cuve_aeree_filling(I)
        Else
            Feuil15.Range("HybAS_cuve_aeree_filling").Cells(I, Colonne) = ""
        End If
        Feuil15.Range("HybAS_cuve_aeree_O2_dissous").Cells(I, Colonne) = cuve_aeree_O2_dissous(I)
        Feuil15.Range("HybAS_cuve_aeree_volume").Cells(I, Colonne) = cuve_aeree_volume(I)
        Feuil15.Range("HybAS_cuve_aeree_O2_alpha").Cells(I, Colonne) = cuve_aeree_O2_facteur_alfa(I)
        Feuil15.Range("HybAS_cuve_aeree_O2_rdt_transfert").Cells(I, Colonne) = cuve_aeree_O2_rdt_transfert(I)
        Feuil15.Range("HybAS_cuve_aeree_hauteur_bassin").Cells(I, Colonne) = cuve_aeree_hauteur_bassin(I)
    Next I

    'Spécifiques au Mox
    If choix_fct_Mox = True Then
        Feuil57.Range("rate_N_T_design").Cells(1, 3 + pi_FiliereConsideree) = rate_N_T_design
        Feuil57.Range("rate_N_T_reel").Cells(1, 3 + pi_FiliereConsideree) = rate_N_T_reel
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil15.Range("HybAS_outlet_NH4").Cells(1, Colonne) = ""    'outlet_NH4
    Feuil15.Range("HybAS_outlet_NO3").Cells(1, Colonne) = ""    'outlet_NO3
    Feuil15.Range("HybAS_pre_denit_NO3").Cells(1, Colonne) = ""    'pre_denit_NO3
    Feuil15.Range("HybAS_NO3_last_aerated_tank").Cells(1, Colonne) = ""    'NO3_last_aerated_tank
    Feuil15.Range("HybAS_Norga_dure_soluble").Cells(1, Colonne) = ""    'Norga_dure_soluble
    Feuil15.Range("HybAS_outlet_DBO_soluble").Cells(1, Colonne) = ""    'outlet_DBO_soluble
    Feuil15.Range("HybAS_recirculation_taux").Cells(1, Colonne) = ""    'recirculation_taux
    Feuil15.Range("HybAS_nominal_MES_concentration").Cells(1, Colonne) = ""    'nominal_MES_concentration
    Feuil15.Range("HybAS_nominal_MV_MES").Cells(1, Colonne) = ""    'nominal_MV_MES
    Feuil15.Range("HybAS_nominal_recirculation_MLSS_taux").Cells(1, Colonne) = ""    'nominal_recirculation_MLSS_taux
    Feuil15.Range("HybAS_reel_MES_concentration").Cells(1, Colonne) = ""    'reel_MES_concentration
    Feuil15.Range("HybAS_reel_MV_MES").Cells(1, Colonne) = ""    'reel_MV_MES
    Feuil15.Range("HybAS_reel_recirculation_MLSS_taux").Cells(1, Colonne) = ""    'reel_recirculation_MLSS_taux
    Feuil15.Range("HybAS_pre_anoxie_volume").Cells(1, Colonne) = ""    'pre_anoxie_volume
    Feuil15.Range("HybAS_anaerobie_volume").Cells(1, Colonne) = ""    'anaerobie_volume
    Feuil15.Range("HybAS_anaerobie_FeCl3_flux").Cells(1, Colonne) = ""    'anaerobie_FeCl3_flux
    Feuil15.Range("HybAS_pre_denit_volume").Cells(1, Colonne) = ""    'pre_denit_volume
    Feuil15.Range("HybAS_deox_O2_dissous").Cells(1, Colonne) = ""    'deox_O2_dissous
    Feuil15.Range("HybAS_deox_volume").Cells(1, Colonne) = ""    'deox_volume
    Feuil15.Range("HybAS_reox_O2_dissous").Cells(1, Colonne) = ""    'reox_O2_dissous
    Feuil15.Range("HybAS_reox_volume").Cells(1, Colonne) = ""    'reox_volume
    Feuil15.Range("HybAS_reox_O2_alpha").Cells(1, Colonne) = ""    'reox_O2_facteur_alfa
    Feuil15.Range("HybAS_reox_O2_rdt_transfert").Cells(1, Colonne) = ""    'reox_O2_rdt_transfert
    Feuil15.Range("HybAS_reox_hauteur_bassin").Cells(1, Colonne) = ""    'reox_hauteur_bassin
    Feuil15.Range("HybAS_post_denit_methanol_flux").Cells(1, Colonne) = ""    'post_denit_methanol_flux
    Feuil15.Range("HybAS_post_denit_media_filling").Cells(1, Colonne) = ""    'post_denit_media_filling
    Feuil15.Range("HybAS_post_denit_volume").Cells(1, Colonne) = ""    'post_denit_volume
    Feuil15.Range("HybAS_ratio_Nnit_media_Nnit_total").Cells(1, Colonne) = ""    'ratio_Nnit_media_Nnit_total
    Feuil15.Range("HybAS_ratio_Nnit_mediaH1_Nnit_media").Cells(1, Colonne) = ""    'ratio_Nnit_mediaH1_Nnit_media
    Feuil15.Range("HybAS_ratio_Nnit_CN1_Nnit_CN12").Cells(1, Colonne) = ""    'ratio_Nnit_CN1_Nnit_CN12
    Feuil15.Range("HybAS_volume_total_bassins").Cells(1, Colonne) = ""    'volume_total_bassins
    Feuil15.Range("HybAS_O2_besoin_total").Cells(1, Colonne) = ""    'O2_besoin_total
    Feuil15.Range("HybAS_air_Q_Nm3j").Cells(1, Colonne) = ""    'air_Q_Nm3j
    Feuil15.Range("HybAS_diffuseur_encrassement").Cells(1, Colonne) = ""    'diffuseur_encrassement
    Feuil15.Range("HybAS_air_P_refoulement_moyenne").Cells(1, Colonne) = ""    'air_P_refoulement_moyenne
    Feuil15.Range("HybAS_surpresseur_conso_spec").Cells(1, Colonne) = ""    'surpresseur_conso_spec
    Feuil15.Range("HybAS_recirculation_P_refoulement").Cells(1, Colonne) = ""    'recirculation_P_refoulement
    Feuil15.Range("HybAS_recirculation_pompe_rdt").Cells(1, Colonne) = ""    'recirculation_pompe_rdt
    Feuil15.Range("HybAS_recirculation_MLSS_P_refoulement").Cells(1, Colonne) = ""    'recirculation_MLSS_P_refoulement
    Feuil15.Range("HybAS_recirculation_MLSS_pompe_rdt").Cells(1, Colonne) = ""    'recirculation_MLSS_pompe_rdt
    Feuil15.Range("HybAS_boues_indice_Mohlman").Cells(1, Colonne) = ""    'boues_indice_Mohlman
    Feuil15.Range("HybAS_sortie_MES").Cells(1, Colonne) = ""    'sortie_MES
    Feuil15.Range("HybAS_clarif_hauteur").Cells(1, Colonne) = ""    'clarif_hauteur
    Feuil15.Range("HybAS_clarif_vitesse_max").Cells(1, Colonne) = ""    'clarif_vitesse_max
    Feuil15.Range("HybAS_clarif_surface").Cells(1, Colonne) = ""    'clarif_surface
    Feuil15.Range("HybAS_clarif_vitesse_max_recalc").Cells(1, Colonne) = ""    'clarif_vitesse_max_recalc
    Feuil15.Range("HybAS_clarif_charge_radier_nominal").Cells(1, Colonne) = ""    'clarif_charge_radier_nominal
    Feuil15.Range("HybAS_boues_concentration").Cells(1, Colonne) = ""    'boues_concentration
    Feuil15.Range("HybAS_boues_MES").Cells(1, Colonne) = ""    'boues_MES
    Feuil15.Range("HybAS_boues_Q").Cells(1, Colonne) = ""    'boues_Q
    Feuil15.Range("HybAS_extraction_P_refoulement").Cells(1, Colonne) = ""    'extraction_P_refoulement
    Feuil15.Range("HybAS_extraction_pompe_rdt").Cells(1, Colonne) = ""    'extraction_pompe_rdt
    Feuil15.Range("HybAS_electricite").Cells(1, Colonne) = ""    'electricite
    Feuil15.Range("HybAS_electricite_aeration").Cells(1, Colonne) = ""     'electricite_aeration

    Feuil15.Range("HybAS_outlet_reel_DBO").Cells(1, Colonne) = ""    'outlet_reel_DBO
    Feuil15.Range("HybAS_outlet_reel_NH4").Cells(1, Colonne) = ""    'outlet_reel_NH4
    Feuil15.Range("HybAS_outlet_reel_NO3").Cells(1, Colonne) = ""    'outlet_reel_NO3

    Feuil15.Range("HybAS_ratio_elec_anaerobie").Cells(1, Colonne) = ""
    Feuil15.Range("HybAS_ratio_elec_pre_anoxie").Cells(1, Colonne) = ""
    Feuil15.Range("HybAS_ratio_elec_pre_denit").Cells(1, Colonne) = ""
    Feuil15.Range("HybAS_ratio_elec_deox").Cells(1, Colonne) = ""
    Feuil15.Range("HybAS_ratio_elec_post_denit").Cells(1, Colonne) = ""


    For I = 1 To 3
        Feuil15.Range("HybAS_cuve_aeree_filling").Cells(I, Colonne) = ""    'cuve_aeree_filling(i)
        Feuil15.Range("HybAS_cuve_aeree_O2_dissous").Cells(I, Colonne) = ""    'cuve_aeree_O2_dissous(i)
        Feuil15.Range("HybAS_cuve_aeree_volume").Cells(I, Colonne) = ""    'cuve_aeree_volume(i)
        Feuil15.Range("HybAS_cuve_aeree_O2_alpha").Cells(I, Colonne) = ""    'cuve_aeree_O2_facteur_alfa(i)
        Feuil15.Range("HybAS_cuve_aeree_O2_rdt_transfert").Cells(I, Colonne) = ""    'cuve_aeree_O2_rdt_transfert(i)
        Feuil15.Range("HybAS_cuve_aeree_hauteur_bassin").Cells(I, Colonne) = ""    'cuve_aeree_hauteur_bassin(i)
    Next I

    'Spécifiques au Mox
    Feuil57.Range("rate_N_T_design").Cells(1, 3 + pi_FiliereConsideree) = ""
    Feuil57.Range("rate_N_T_reel").Cells(1, 3 + pi_FiliereConsideree) = ""


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim nb_cuves As Integer
    
    nb_cuves = 0
    
    choix_fct_Mox = Feuil57.Range("choix_fct_Mox").Cells(1, pi_FiliereConsideree)
    
    choix_pre_anoxie = Feuil15.Range("HybAS_choix_pre_anoxie").Cells(1, pi_FiliereConsideree)
    If choix_pre_anoxie <> cuve_non Then
        nb_cuves = nb_cuves + 1
    End If
    choix_anaerobie = Feuil15.Range("HybAS_choix_anaerobie").Cells(1, pi_FiliereConsideree)
    If choix_anaerobie <> cuve_non Then
        nb_cuves = nb_cuves + 1
    End If
    If choix_anaerobie <> cuve_non And choix_pre_anoxie = cuve_non Then
        MsgBox Feuil15.Range("HybAS_erreur_anaerobie_sans_preanoxie"), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If
    choix_pre_denit = Feuil15.Range("HybAS_choix_pre_denit").Cells(1, pi_FiliereConsideree)
    If choix_pre_denit <> cuve_non Then
        nb_cuves = nb_cuves + 1
    End If
    choix_deox = Feuil15.Range("HybAS_choix_deox").Cells(1, pi_FiliereConsideree)
    If choix_deox <> cuve_non Then
        nb_cuves = nb_cuves + 1
    End If
    choix_reox = Feuil15.Range("HybAS_choix_reox").Cells(1, pi_FiliereConsideree)
    If choix_reox <> cuve_non Then
        nb_cuves = nb_cuves + 1
    End If
    choix_aerateur_reox = Feuil15.Range("HybAS_choix_aerateur_reox").Cells(1, pi_FiliereConsideree)
    choix_post_denit = Feuil15.Range("HybAS_choix_post_denit").Cells(1, pi_FiliereConsideree)
    If choix_post_denit <> cuve_non Then
        nb_cuves = nb_cuves + 1
    End If
    choix_media_post_denit = Feuil15.Range("HybAS_choix_media_post_denit").Cells(1, pi_FiliereConsideree)
    choix_regulation = Feuil15.Range("HybAS_choix_regulation").Cells(1, pi_FiliereConsideree)
    choix_surpresseur = Feuil15.Range("HybAS_choix_surpresseur").Cells(1, pi_FiliereConsideree)
    choix_racleur = Feuil15.Range("HybAS_choix_racleur").Cells(1, pi_FiliereConsideree)


    choix_configuration_cuves_aerees = Feuil15.Range("HybAS_choix_configuration_cuves_aerees").Cells(1, pi_FiliereConsideree)

    'détail sur les cuves
    If choix_configuration_cuves_aerees = config_CN_H Then
        choix_nb_cuves_aerees = 2
        ReDim choix_type_cuve_aeree(choix_nb_cuves_aerees)
        ReDim choix_media_cuve_aeree(choix_nb_cuves_aerees)
        'caractéristiques des cuves
        choix_type_cuve_aeree(1) = cuve_aeree_CN
        choix_type_cuve_aeree(2) = cuve_aeree_H
    ElseIf choix_configuration_cuves_aerees = config_CN_H_CN Then
        choix_nb_cuves_aerees = 3
        ReDim choix_type_cuve_aeree(choix_nb_cuves_aerees)
        ReDim choix_media_cuve_aeree(choix_nb_cuves_aerees)
        'caractéristiques des cuves
        choix_type_cuve_aeree(1) = cuve_aeree_CN
        choix_type_cuve_aeree(2) = cuve_aeree_H
        choix_type_cuve_aeree(3) = cuve_aeree_CN
    ElseIf choix_configuration_cuves_aerees = config_CN_H_H Then
        choix_nb_cuves_aerees = 3
        ReDim choix_type_cuve_aeree(choix_nb_cuves_aerees)
        ReDim choix_media_cuve_aeree(choix_nb_cuves_aerees)
        'caractéristiques des cuves
        choix_type_cuve_aeree(1) = cuve_aeree_CN
        choix_type_cuve_aeree(2) = cuve_aeree_H
        choix_type_cuve_aeree(3) = cuve_aeree_H
    ElseIf choix_configuration_cuves_aerees = config_H_H Then
        choix_nb_cuves_aerees = 2
        ReDim choix_type_cuve_aeree(choix_nb_cuves_aerees)
        ReDim choix_media_cuve_aeree(choix_nb_cuves_aerees)
        'caractéristiques des cuves
        choix_type_cuve_aeree(1) = cuve_aeree_H
        choix_type_cuve_aeree(2) = cuve_aeree_H
    ElseIf choix_configuration_cuves_aerees = config_H_CN Then
        choix_nb_cuves_aerees = 2
        ReDim choix_type_cuve_aeree(choix_nb_cuves_aerees)
        ReDim choix_media_cuve_aeree(choix_nb_cuves_aerees)
        'caractéristiques des cuves
        choix_type_cuve_aeree(1) = cuve_aeree_H
        choix_type_cuve_aeree(2) = cuve_aeree_CN
    ElseIf choix_configuration_cuves_aerees = config_H Then
        choix_nb_cuves_aerees = 1
        ReDim choix_type_cuve_aeree(choix_nb_cuves_aerees)
        ReDim choix_media_cuve_aeree(choix_nb_cuves_aerees)
        'caractéristiques des cuves
        choix_type_cuve_aeree(1) = cuve_aeree_H
    End If
    nb_cuves = nb_cuves + choix_nb_cuves_aerees
    If choix_fct_Mox = True Then
        If nb_cuves > 1 Then
            MsgBox ("Mox utilisable uniquement en configuration H sans cuve en amont ni en aval.")
            pn_EffacementResultatsEnCours = True
            Exit Sub
        End If
        If choix_media_cuve_aeree(1) = media_chip_P Or choix_media_cuve_aeree(1) = media_K1 Or choix_media_cuve_aeree(1) = media_K1_heavy Then
            MsgBox ("media non utilisable en configuration Mox --> choix possibles: K3, K5, Chip M")
            pn_EffacementResultatsEnCours = True
            Exit Sub
        End If
    End If
    
    
    'on redimensionne tous les vecteurs le nécessitant
    ReDim choix_aerateur_cuve_aeree(choix_nb_cuves_aerees)
    ReDim cuve_aeree_volume(choix_nb_cuves_aerees)
    ReDim cuve_aeree_filling(choix_nb_cuves_aerees)
    ReDim cuve_aeree_O2_dissous(choix_nb_cuves_aerees)
    ReDim cuve_aeree_O2_facteur_alfa(choix_nb_cuves_aerees)
    ReDim cuve_aeree_O2_rdt_transfert(choix_nb_cuves_aerees)
    ReDim cuve_aeree_hauteur_bassin(choix_nb_cuves_aerees)
    ReDim cuve_aeree_volume_f(choix_nb_cuves_aerees)
    ReDim cuve_aeree_filling_f(choix_nb_cuves_aerees)
    ReDim cuve_aeree_O2_dissous_f(choix_nb_cuves_aerees)
    ReDim cuve_aeree_O2_facteur_alfa_f(choix_nb_cuves_aerees)
    ReDim cuve_aeree_O2_rdt_transfert_f(choix_nb_cuves_aerees)
    ReDim cuve_aeree_hauteur_bassin_f(choix_nb_cuves_aerees)
    ReDim cuve_aeree_volume_force(choix_nb_cuves_aerees)
    ReDim cuve_aeree_filling_force(choix_nb_cuves_aerees)
    ReDim cuve_aeree_O2_dissous_force(choix_nb_cuves_aerees)
    ReDim cuve_aeree_O2_facteur_alfa_force(choix_nb_cuves_aerees)
    ReDim cuve_aeree_O2_rdt_transfert_force(choix_nb_cuves_aerees)
    ReDim cuve_aeree_hauteur_bassin_force(choix_nb_cuves_aerees)


    'choix du média
    nb_cuves_CN = 0
    nb_cuves_H = 0
    For I = 1 To choix_nb_cuves_aerees
        If choix_type_cuve_aeree(I) = cuve_aeree_H Then
            choix_media_cuve_aeree(I) = Feuil15.Range("HybAS_choix_media_cuves_aerees").Cells(I, pi_FiliereConsideree)
            nb_cuves_H = nb_cuves_H + 1
        Else
            choix_media_cuve_aeree(I) = no_media
            nb_cuves_CN = nb_cuves_CN + 1
        End If
        choix_aerateur_cuve_aeree(I) = Feuil15.Range("HybAS_choix_aerateur_cuves_aerees").Cells(I, pi_FiliereConsideree)
    Next I


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim I As Integer


    outlet_NH4_force = Feuil15.Range("HybAS_outlet_NH4").Cells(1, pi_FiliereConsideree)
    outlet_NO3_force = Feuil15.Range("HybAS_outlet_NO3").Cells(1, pi_FiliereConsideree)
    Norga_dure_soluble_force = Feuil15.Range("HybAS_Norga_dure_soluble").Cells(1, pi_FiliereConsideree)
    outlet_DBO_soluble_force = Feuil15.Range("HybAS_outlet_DBO_soluble").Cells(1, pi_FiliereConsideree)
    recirculation_taux_force = Feuil15.Range("HybAS_recirculation_taux").Cells(1, pi_FiliereConsideree)
    nominal_MES_concentration_force = Feuil15.Range("HybAS_nominal_MES_concentration").Cells(1, pi_FiliereConsideree)
    nominal_MV_MES_force = Feuil15.Range("HybAS_nominal_MV_MES").Cells(1, pi_FiliereConsideree)
    nominal_recirculation_MLSS_taux_force = Feuil15.Range("HybAS_nominal_recirculation_MLSS_taux").Cells(1, pi_FiliereConsideree)
    reel_MES_concentration_force = Feuil15.Range("HybAS_reel_MES_concentration").Cells(1, pi_FiliereConsideree)
    reel_MV_MES_force = Feuil15.Range("HybAS_reel_MV_MES").Cells(1, pi_FiliereConsideree)
    reel_recirculation_MLSS_taux_force = Feuil15.Range("HybAS_reel_recirculation_MLSS_taux").Cells(1, pi_FiliereConsideree)
    If choix_pre_anoxie <> cuve_non Then
        pre_anoxie_volume_force = Feuil15.Range("HybAS_pre_anoxie_volume").Cells(1, pi_FiliereConsideree)
        ratio_elec_pre_anoxie_force = Feuil15.Range("HybAS_ratio_elec_pre_anoxie").Cells(1, pi_FiliereConsideree)
    End If
    If choix_anaerobie <> cuve_non Then
        anaerobie_volume_force = Feuil15.Range("HybAS_anaerobie_volume").Cells(1, pi_FiliereConsideree)
        anaerobie_FeCl3_flux_force = Feuil15.Range("HybAS_anaerobie_FeCl3_flux").Cells(1, pi_FiliereConsideree)
        ratio_elec_anaerobie_force = Feuil15.Range("HybAS_ratio_elec_anaerobie").Cells(1, pi_FiliereConsideree)
    End If
    If choix_pre_denit <> cuve_non Then
        pre_denit_volume_force = Feuil15.Range("HybAS_pre_denit_volume").Cells(1, pi_FiliereConsideree)
        pre_denit_NO3_force = Feuil15.Range("HybAS_pre_denit_NO3").Cells(1, pi_FiliereConsideree)
        ratio_elec_pre_denit_force = Feuil15.Range("HybAS_ratio_elec_pre_denit").Cells(1, pi_FiliereConsideree)
    End If
    If choix_deox <> cuve_non Then
        deox_O2_dissous_force = Feuil15.Range("HybAS_deox_O2_dissous").Cells(1, pi_FiliereConsideree)
        deox_volume_force = Feuil15.Range("HybAS_deox_volume").Cells(1, pi_FiliereConsideree)
        ratio_elec_deox_force = Feuil15.Range("HybAS_ratio_elec_deox").Cells(1, pi_FiliereConsideree)
    End If
    If choix_reox <> cuve_non Then
        reox_O2_dissous_force = Feuil15.Range("HybAS_reox_O2_dissous").Cells(1, pi_FiliereConsideree)
        reox_volume_force = Feuil15.Range("HybAS_reox_volume").Cells(1, pi_FiliereConsideree)
        reox_O2_facteur_alfa_force = Feuil15.Range("HybAS_reox_O2_alpha").Cells(1, pi_FiliereConsideree)
        reox_O2_rdt_transfert_force = Feuil15.Range("HybAS_reox_O2_rdt_transfert").Cells(1, pi_FiliereConsideree)
        reox_hauteur_bassin_force = Feuil15.Range("HybAS_reox_hauteur_bassin").Cells(1, pi_FiliereConsideree)
    End If
    If choix_post_denit <> cuve_non Then
        NO3_last_aerated_tank_force = Feuil15.Range("HybAS_NO3_last_aerated_tank").Cells(1, pi_FiliereConsideree)
        post_denit_methanol_flux_force = Feuil15.Range("HybAS_post_denit_methanol_flux").Cells(1, pi_FiliereConsideree)
        post_denit_media_filling_force = Feuil15.Range("HybAS_post_denit_media_filling").Cells(1, pi_FiliereConsideree)
        post_denit_volume_force = Feuil15.Range("HybAS_post_denit_volume").Cells(1, pi_FiliereConsideree)
        ratio_elec_post_denit_force = Feuil15.Range("HybAS_ratio_elec_post_denit").Cells(1, pi_FiliereConsideree)
    End If
    If nb_cuves_CN > 0 Then  'dans ce cas-là, il est juste calculé
        ratio_Nnit_media_Nnit_total_force = Feuil15.Range("HybAS_ratio_Nnit_media_Nnit_total").Cells(1, pi_FiliereConsideree)
    Else
        ratio_Nnit_media_Nnit_total_force = False
    End If
    If nb_cuves_H = 2 Then
        ratio_Nnit_mediaH1_Nnit_media_force = Feuil15.Range("HybAS_ratio_Nnit_mediaH1_Nnit_media").Cells(1, pi_FiliereConsideree)
    Else
        ratio_Nnit_mediaH1_Nnit_media_force = False
    End If
    If nb_cuves_CN = 2 Then
        ratio_Nnit_CN1_Nnit_CN12_force = Feuil15.Range("HybAS_ratio_Nnit_CN1_Nnit_CN12").Cells(1, pi_FiliereConsideree)
    Else
        ratio_Nnit_CN1_Nnit_CN12_force = False
    End If
    outlet_reel_DBO_force = Feuil15.Range("HybAS_outlet_reel_DBO").Cells(1, pi_FiliereConsideree)
    outlet_reel_NH4_force = Feuil15.Range("HybAS_outlet_reel_NH4").Cells(1, pi_FiliereConsideree)
    outlet_reel_NO3_force = Feuil15.Range("HybAS_outlet_reel_NO3").Cells(1, pi_FiliereConsideree)
    O2_besoin_total_force = Feuil15.Range("HybAS_O2_besoin_total").Cells(1, pi_FiliereConsideree)
    air_Q_Nm3j_force = Feuil15.Range("HybAS_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
    diffuseur_encrassement_force = Feuil15.Range("HybAS_diffuseur_encrassement").Cells(1, pi_FiliereConsideree)
    air_P_refoulement_moyenne_force = Feuil15.Range("HybAS_air_P_refoulement_moyenne").Cells(1, pi_FiliereConsideree)
    surpresseur_conso_spec_force = Feuil15.Range("HybAS_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
    recirculation_P_refoulement_force = Feuil15.Range("HybAS_recirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
    recirculation_pompe_rdt_force = Feuil15.Range("HybAS_recirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
    recirculation_MLSS_P_refoulement_force = Feuil15.Range("HybAS_recirculation_MLSS_P_refoulement").Cells(1, pi_FiliereConsideree)
    recirculation_MLSS_pompe_rdt_force = Feuil15.Range("HybAS_recirculation_MLSS_pompe_rdt").Cells(1, pi_FiliereConsideree)
    boues_indice_Mohlman_force = Feuil15.Range("HybAS_boues_indice_Mohlman").Cells(1, pi_FiliereConsideree)
    sortie_MES_force = Feuil15.Range("HybAS_sortie_MES").Cells(1, pi_FiliereConsideree)
    clarif_hauteur_force = Feuil15.Range("HybAS_clarif_hauteur").Cells(1, pi_FiliereConsideree)
    clarif_vitesse_max_force = Feuil15.Range("HybAS_clarif_vitesse_max").Cells(1, pi_FiliereConsideree)
    clarif_surface_force = Feuil15.Range("HybAS_clarif_surface").Cells(1, pi_FiliereConsideree)
    boues_concentration_force = Feuil15.Range("HybAS_boues_concentration").Cells(1, pi_FiliereConsideree)
    boues_MES_force = Feuil15.Range("HybAS_boues_MES").Cells(1, pi_FiliereConsideree)
    boues_Q_force = Feuil15.Range("HybAS_boues_Q").Cells(1, pi_FiliereConsideree)
    extraction_P_refoulement_force = Feuil15.Range("HybAS_extraction_P_refoulement").Cells(1, pi_FiliereConsideree)
    extraction_pompe_rdt_force = Feuil15.Range("HybAS_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)





    If outlet_NH4_force = True Then
        outlet_NH4_f = Feuil15.Range("HybAS_outlet_NH4").Cells(1, pi_FiliereConsideree)
    End If
    If outlet_NO3_force = True Then
        outlet_NO3_f = Feuil15.Range("HybAS_outlet_NO3").Cells(1, pi_FiliereConsideree)
    End If
    If pre_denit_NO3_force = True Then
        pre_denit_NO3_f = Feuil15.Range("HybAS_pre_denit_NO3").Cells(1, pi_FiliereConsideree)
    End If
    If NO3_last_aerated_tank_force = True Then
        NO3_last_aerated_tank_f = Feuil15.Range("HybAS_NO3_last_aerated_tank").Cells(1, pi_FiliereConsideree)
    End If
    If Norga_dure_soluble_force = True Then
        Norga_dure_soluble_f = Feuil15.Range("HybAS_Norga_dure_soluble").Cells(1, pi_FiliereConsideree)
    End If
    If outlet_DBO_soluble_force = True Then
        outlet_DBO_soluble_f = Feuil15.Range("HybAS_outlet_DBO_soluble").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_taux_force = True Then
        recirculation_taux_f = Feuil15.Range("HybAS_recirculation_taux").Cells(1, pi_FiliereConsideree)
    End If
    If nominal_MES_concentration_force = True Then
        nominal_MES_concentration_f = Feuil15.Range("HybAS_nominal_MES_concentration").Cells(1, pi_FiliereConsideree)
    End If
    If nominal_MV_MES_force = True Then
        nominal_MV_MES_f = Feuil15.Range("HybAS_nominal_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If nominal_recirculation_MLSS_taux_force = True Then
        nominal_recirculation_MLSS_taux_f = Feuil15.Range("HybAS_nominal_recirculation_MLSS_taux").Cells(1, pi_FiliereConsideree)
    End If
    If reel_MES_concentration_force = True Then
        reel_MES_concentration_f = Feuil15.Range("HybAS_reel_MES_concentration").Cells(1, pi_FiliereConsideree)
    End If
    If reel_MV_MES_force = True Then
        reel_MV_MES_f = Feuil15.Range("HybAS_reel_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If reel_recirculation_MLSS_taux_force = True Then
        reel_recirculation_MLSS_taux_f = Feuil15.Range("HybAS_reel_recirculation_MLSS_taux").Cells(1, pi_FiliereConsideree)
    End If
    If pre_anoxie_volume_force = True Then
        pre_anoxie_volume_f = Feuil15.Range("HybAS_pre_anoxie_volume").Cells(1, pi_FiliereConsideree)
    End If
    If anaerobie_volume_force = True Then
        anaerobie_volume_f = Feuil15.Range("HybAS_anaerobie_volume").Cells(1, pi_FiliereConsideree)
    End If
    If anaerobie_FeCl3_flux_force = True Then
        anaerobie_FeCl3_flux_f = Feuil15.Range("HybAS_anaerobie_FeCl3_flux").Cells(1, pi_FiliereConsideree)
    End If
    If pre_denit_volume_force = True Then
        pre_denit_volume_f = Feuil15.Range("HybAS_pre_denit_volume").Cells(1, pi_FiliereConsideree)
    End If
    If deox_O2_dissous_force = True Then
        deox_O2_dissous_f = Feuil15.Range("HybAS_deox_O2_dissous").Cells(1, pi_FiliereConsideree)
    End If
    If deox_volume_force = True Then
        deox_volume_f = Feuil15.Range("HybAS_deox_volume").Cells(1, pi_FiliereConsideree)
    End If
    If reox_O2_dissous_force = True Then
        reox_O2_dissous_f = Feuil15.Range("HybAS_reox_O2_dissous").Cells(1, pi_FiliereConsideree)
    End If
    If reox_volume_force = True Then
        reox_volume_f = Feuil15.Range("HybAS_reox_volume").Cells(1, pi_FiliereConsideree)
    End If
    If reox_O2_facteur_alfa_force = True Then
        reox_O2_facteur_alfa_f = Feuil15.Range("HybAS_reox_O2_alpha").Cells(1, pi_FiliereConsideree)
    End If
    If reox_O2_rdt_transfert_force = True Then
        reox_O2_rdt_transfert_f = Feuil15.Range("HybAS_reox_O2_rdt_transfert").Cells(1, pi_FiliereConsideree)
    End If
    If reox_hauteur_bassin_force = True Then
        reox_hauteur_bassin_f = Feuil15.Range("HybAS_reox_hauteur_bassin").Cells(1, pi_FiliereConsideree)
    End If
    If post_denit_methanol_flux_force = True Then
        post_denit_methanol_flux_f = Feuil15.Range("HybAS_post_denit_methanol_flux").Cells(1, pi_FiliereConsideree)
    End If
    If post_denit_media_filling_force = True Then
        post_denit_media_filling_f = Feuil15.Range("HybAS_post_denit_media_filling").Cells(1, pi_FiliereConsideree)
    End If
    If post_denit_volume_force = True Then
        post_denit_volume_f = Feuil15.Range("HybAS_post_denit_volume").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_Nnit_media_Nnit_total_force = True Then
        ratio_Nnit_media_Nnit_total_f = Feuil15.Range("HybAS_ratio_Nnit_media_Nnit_total").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_Nnit_mediaH1_Nnit_media_force = True Then
        ratio_Nnit_mediaH1_Nnit_media_f = Feuil15.Range("HybAS_ratio_Nnit_mediaH1_Nnit_media").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_Nnit_CN1_Nnit_CN12_force = True Then
        ratio_Nnit_CN1_Nnit_CN12_f = Feuil15.Range("HybAS_ratio_Nnit_CN1_Nnit_CN12").Cells(1, pi_FiliereConsideree)
    End If
    If outlet_reel_DBO_force = True Then
        outlet_reel_DBO_f = Feuil15.Range("HybAS_outlet_reel_DBO").Cells(1, pi_FiliereConsideree)
    End If
    If outlet_reel_NH4_force = True Then
        outlet_reel_NH4_f = Feuil15.Range("HybAS_outlet_reel_NH4").Cells(1, pi_FiliereConsideree)
    End If
    If outlet_reel_NO3_force = True Then
        outlet_reel_NO3_f = Feuil15.Range("HybAS_outlet_reel_NO3").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_elec_pre_anoxie_force = True Then
        ratio_elec_pre_anoxie_f = Feuil15.Range("HybAS_ratio_elec_pre_anoxie").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_elec_anaerobie_force = True Then
        ratio_elec_anaerobie_f = Feuil15.Range("HybAS_ratio_elec_anaerobie").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_elec_pre_denit_force = True Then
        ratio_elec_pre_denit_f = Feuil15.Range("HybAS_ratio_elec_pre_denit").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_elec_deox_force = True Then
        ratio_elec_deox_f = Feuil15.Range("HybAS_ratio_elec_deox").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_elec_post_denit_force = True Then
        ratio_elec_post_denit_f = Feuil15.Range("HybAS_ratio_elec_post_denit").Cells(1, pi_FiliereConsideree)
    End If
    If O2_besoin_total_force = True Then
        O2_besoin_total_f = Feuil15.Range("HybAS_O2_besoin_total").Cells(1, pi_FiliereConsideree)
    End If
    If air_Q_Nm3j_force = True Then
        air_Q_Nm3j_f = Feuil15.Range("HybAS_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
    End If
    If diffuseur_encrassement_force = True Then
        diffuseur_encrassement_f = Feuil15.Range("HybAS_diffuseur_encrassement").Cells(1, pi_FiliereConsideree)
    End If
    If air_P_refoulement_moyenne_force = True Then
        air_P_refoulement_moyenne_f = Feuil15.Range("HybAS_air_P_refoulement_moyenne").Cells(1, pi_FiliereConsideree)
    End If
    If surpresseur_conso_spec_force = True Then
        surpresseur_conso_spec_f = Feuil15.Range("HybAS_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_P_refoulement_force = True Then
        recirculation_P_refoulement_f = Feuil15.Range("HybAS_recirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_pompe_rdt_force = True Then
        recirculation_pompe_rdt_f = Feuil15.Range("HybAS_recirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_MLSS_P_refoulement_force = True Then
        recirculation_MLSS_P_refoulement_f = Feuil15.Range("HybAS_recirculation_MLSS_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_MLSS_pompe_rdt_force = True Then
        recirculation_MLSS_pompe_rdt_f = Feuil15.Range("HybAS_recirculation_MLSS_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    If boues_indice_Mohlman_force = True Then
        boues_indice_Mohlman_f = Feuil15.Range("HybAS_boues_indice_Mohlman").Cells(1, pi_FiliereConsideree)
    End If
    If sortie_MES_force = True Then
        sortie_MES_f = Feuil15.Range("HybAS_sortie_MES").Cells(1, pi_FiliereConsideree)
    End If
    If clarif_hauteur_force = True Then
        clarif_hauteur_f = Feuil15.Range("HybAS_clarif_hauteur").Cells(1, pi_FiliereConsideree)
    End If
    If clarif_vitesse_max_force = True Then
        clarif_vitesse_max_f = Feuil15.Range("HybAS_clarif_vitesse_max").Cells(1, pi_FiliereConsideree)
    End If
    If clarif_surface_force = True Then
        clarif_surface_f = Feuil15.Range("HybAS_clarif_surface").Cells(1, pi_FiliereConsideree)
    End If
    If boues_concentration_force = True Then
        boues_concentration_f = Feuil15.Range("HybAS_boues_concentration").Cells(1, pi_FiliereConsideree)
    End If
    If boues_MES_force = True Then
        boues_MES_f = Feuil15.Range("HybAS_boues_MES").Cells(1, pi_FiliereConsideree)
    End If
    If boues_Q_force = True Then
        boues_Q_f = Feuil15.Range("HybAS_boues_Q").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement_f = Feuil15.Range("HybAS_extraction_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt_f = Feuil15.Range("HybAS_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If


    'CAS PARTICULIER DES CUVES CENTRALES
    For I = 1 To choix_nb_cuves_aerees
        If choix_type_cuve_aeree(I) = cuve_aeree_H Then
            cuve_aeree_filling_force(I) = Feuil15.Range("HybAS_cuve_aeree_filling").Cells(I, pi_FiliereConsideree)
            If cuve_aeree_filling_force(I) = True Then
                cuve_aeree_filling_f(I) = Feuil15.Range("HybAS_cuve_aeree_filling").Cells(I, pi_FiliereConsideree)
            End If
        End If
        cuve_aeree_O2_dissous_force(I) = Feuil15.Range("HybAS_cuve_aeree_O2_dissous").Cells(I, pi_FiliereConsideree)
        If cuve_aeree_O2_dissous_force(I) = True Then
            cuve_aeree_O2_dissous_f(I) = Feuil15.Range("HybAS_cuve_aeree_O2_dissous").Cells(I, pi_FiliereConsideree)
        End If
        cuve_aeree_volume_force(I) = Feuil15.Range("HybAS_cuve_aeree_volume").Cells(I, pi_FiliereConsideree)
        If cuve_aeree_volume_force(I) = True Then
            cuve_aeree_volume_f(I) = Feuil15.Range("HybAS_cuve_aeree_volume").Cells(I, pi_FiliereConsideree)
        End If
        cuve_aeree_O2_facteur_alfa_force(I) = Feuil15.Range("HybAS_cuve_aeree_O2_alpha").Cells(I, pi_FiliereConsideree)
        If cuve_aeree_O2_facteur_alfa_force(I) = True Then
            cuve_aeree_O2_facteur_alfa_f(I) = Feuil15.Range("HybAS_cuve_aeree_O2_alpha").Cells(I, pi_FiliereConsideree)
        End If
        cuve_aeree_O2_rdt_transfert_force(I) = Feuil15.Range("HybAS_cuve_aeree_O2_rdt_transfert").Cells(I, pi_FiliereConsideree)
        If cuve_aeree_O2_rdt_transfert_force(I) = True Then
            cuve_aeree_O2_rdt_transfert_f(I) = Feuil15.Range("HybAS_cuve_aeree_O2_rdt_transfert").Cells(I, pi_FiliereConsideree)
        End If
        cuve_aeree_hauteur_bassin_force(I) = Feuil15.Range("HybAS_cuve_aeree_hauteur_bassin").Cells(I, pi_FiliereConsideree)
        If cuve_aeree_hauteur_bassin_force(I) = True Then
            cuve_aeree_hauteur_bassin_f(I) = Feuil15.Range("HybAS_cuve_aeree_hauteur_bassin").Cells(I, pi_FiliereConsideree)
        End If
    Next I

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
    NumeroErreurNonGeree = "15.2"

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
        rate_N_30(media_K3) = 2.8 * 2
        rate_N_30(media_K5) = 2.3 * 2
        rate_N_30(media_chip_M) = 2 * 2
        'ratio de production de boues par rapport aux MES inlet
        ratio_boues_MESin = 0.35
        'kg N-NO3 formés par kg de N nitrifiés
        ratio_NO3f_nit = 0.11
        'ratio de conso d'O2 de la nit
        ratio_O2_nit = 1.94
    Else
        rate_N_30(media_K3) = 0
        rate_N_30(media_K5) = 0
        rate_N_30(media_chip_M) = 0
        'ratio de production de boues par rapport aux MES inlet
        If choix_primaire = True Then
            ratio_boues_MESin = 0.35
        Else
            ratio_boues_MESin = 0.5
        End If
        'kg N-NO3 formés par kg de N nitrifiés
        ratio_NO3f_nit = 1
        'ratio de conso d'O2 de la nit
        ratio_O2_nit = 4.57
    End If

    'concentration à rajouter à celle de la Post-denit pour avoir celle du dernier tank aéré   mg/L
    delta_NO3_post_DN = 6

    'calcul du ratio MV/MES eau brute   a0+a1*MES/DCO+a2*DCO/DBO
    a_ratio_EB_MV_MES(0) = 1.213
    a_ratio_EB_MV_MES(1) = -1 / 3
    a_ratio_EB_MV_MES(2) = -0.1429
    ratio_EB_MV_MES_min = 0.4
    ratio_EB_MV_MES_max = 0.8
    'ratio MV/MES eau décantée
    ratio_ED_MV_MES = 0.75

    'ratio de N assimilé
    ratio_N_assimile = 0.06

    'ratio de N assimilé par la DCO du méthanol qui sert à la production de boues
    ratio_N_assimile_methanol = 0.08      '***************** juste pour la diffusion car erreur introuvable en post DN

    'ratio de DCO dure par rapport à la DCO totale de l'eau brute
    ratio_DCO_dure_total = 0.33 * 0.13

    'ratio de production de boues avec la nit
    ratio_boues_nit = 0.125

    'ratio de production de boues avec la DBO
    ratio_boues_DBOin = 0.6

    'O2 dissous dans les boues recirculées  mg/L
    recirculation_O2 = 1

    'charge à dénitrifier en plus à cause de l'O2
    ratio_NO3eq_O2 = 0.35

    'concentration attendue en NO3 sortie pré-anoxie   mg/L
    pre_anoxie_NO3 = 0.3

    'ratio C/N limite max pour la dénit
    ratio_DBO_denit_max = 15

    'ratio de conso d'O2 (en l'ocurrence DBO) lors de la dénit
    ratio_O2_denit = -2.86

    'concentration limite pour changer de HRT en anaérobie
    anaerobie_P_limite = 10    'mg/L
    anaerobie_HRT_P_low = 1.3    'h de temps de séjour si inférieur à P limite
    anaerobie_HRT_P_high = 2  'h de temps de séjour si supérieur à P limite

    'assimilation du P lorsque DBO traitée
    ratio_P_assimile = 0.01

    'ratio DBO soluble éliminée/P traité par biologie  (facteur Beta pour Xavier LE TALLEC)
    anaerobie_DBO_Pt = 5

    'masses molaires utiles pour anaérobie  en g/mol
    MM_P = 31
    'MM_O = 16
    'MM_H = 1
    'MM_Fe = 55.85
    MM_FeO3H3 = 55.85 + 3 * (16 + 1)    'FE(OH)3
    MM_FePO4 = 55.85 + 31 + 4 * 16
    MM_FeCl3 = 162.5

    'temps de séjour minimal cuve Dé-Ox  en h
    deox_HRT_mini = 5 / 60   '5 min

    'ratio de conso de l'O2 pour la DBO5
    ratio_O2_DBO = 0.75

    'temps de séjour cuve Ré-Ox   en j
    If pd_CoefficientPointeHydrauliqueTempsPluie <= 1 Then
        reox_HRT = 20 / (60 * 24)
    ElseIf pd_CoefficientPointeHydrauliqueTempsPluie >= 3 Then
        reox_HRT = (25 - 5 * pd_CoefficientPointeHydrauliqueTempsPluie) / (60 * 24)
    Else
        reox_HRT = 10 / (60 * 24)
    End If

    'ratio de DCO à apporter pour la dénit   kgO2/kgNNO3
    If pd_TemperatureEauDimensionnement_°C < 5 Then
        rate_DCO_apportee_denit_nominal = 5
    ElseIf pd_TemperatureEauDimensionnement_°C > 15 Then
        rate_DCO_apportee_denit_nominal = 4
    Else
        rate_DCO_apportee_denit_nominal = 5 - 0.1 * (pd_TemperatureEauDimensionnement_°C - 5)
    End If
    If pd_TemperatureEauConditionsReelles_°C < 5 Then
        rate_DCO_apportee_denit_reel = 5
    ElseIf pd_TemperatureEauConditionsReelles_°C > 15 Then
        rate_DCO_apportee_denit_reel = 4
    Else
        rate_DCO_apportee_denit_reel = 5 - 0.1 * (pd_TemperatureEauConditionsReelles_°C - 5)
    End If

    'DCO du méthanol   kgO2/kg méthanol
    rate_DCO_methanol = 1.5
    ratio_boues_DCO_methanol = 0.25  'kg MES/kg O2

    'hypothèse de cinétique en post-denit  (sans média)  gNNO3/(kgVSS.h)
    hyp_cinetique_post_denit_15 = 6     'à 15°C 'methanol
    a_cinetique_postDN_nominal = 1.08 ^ (pd_TemperatureEauDimensionnement_°C - 15)
    a_cinetique_postDN_reel = 1.08 ^ (pd_TemperatureEauConditionsReelles_°C - 15)
    cinetique_post_denit_15_mini = 3   'cinétique limite à 15°C pour demander à mettre du média
    'avec media
    a_cinetique_postDN_media_10°C(1) = 2.29    ' = a(1)*NO3_out/(NO3_out+a(2))
    a_cinetique_postDN_media_10°C(2) = 2.9
    a_cinetique_postDN_media_nominal = 1.06 ^ (pd_TemperatureEauDimensionnement_°C - 10)
    a_cinetique_postDN_media_reel = 1.06 ^ (pd_TemperatureEauConditionsReelles_°C - 10)

    'temps de séjour mini
    post_DN_SRT_mini = 1.35 * 1.07 ^ (14 - pd_TemperatureEauDimensionnement_°C)

    'données sur les médias
    'surface spécifique m2/m3
    media_surface_m2_m3(media_K1) = 500
    media_surface_m2_m3(media_K1_heavy) = 500
    media_surface_m2_m3(media_K3) = 500
    media_surface_m2_m3(media_K5) = 800
    media_surface_m2_m3(media_chip_P) = 900
    media_surface_m2_m3(media_chip_M) = 1200
    'taux de remplissage maximum cuves H
    media_filling_max(media_K3) = 0.65
    media_filling_max(media_K5) = 0.6
    media_filling_max(media_chip_P) = 0.55
    media_filling_max(media_chip_M) = 0.55
    'taux de remplissage maximum cuve Post DN
    media_filling_max_PDN(media_K1) = 0.5
    media_filling_max_PDN(media_K1_heavy) = 0.5
    media_filling_max_PDN(media_K3) = 0.5
    media_filling_max_PDN(media_K5) = 0.5
    'déplacement
    media_deplacement(media_K1) = 0.14
    media_deplacement(media_K1_heavy) = 0.14
    media_deplacement(media_K3) = 0.11
    media_deplacement(media_K5) = 0.123
    media_deplacement(media_chip_P) = 0.173
    media_deplacement(media_chip_M) = 0.23

    'Charge radier maximale kg/(m2.h)
    charge_radier_max(racleur_racle) = 5
    charge_radier_max(racleur_racle_suce) = 7.5
    charge_radier_max(racleur_Kruger) = 9

    'facteur pour calcul de concentration limite des boues extraites et recirculées pour bonne décantation
    facteur_boues_C_lim = 1200

    'facteur beta
    O2_facteur_beta = 0.99

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

    agitation_fct = 24

    'ratio de N organique dur soluble par rapport à DCO en entrée
    ratio_Norgaduresoluble_DCOtot = 0.25 / 100    '0.25% de la DCO totale

    'vitesse mini agitation cuves H
    If choix_fct_Mox = True Then
        cuve_H_V_Nmh_mini = 0   'pour le Mox, on installe des agitateurs donc on annule la condition
    Else
        cuve_H_V_Nmh_mini = 7   'Nm3/(m2.h)   vu avec Kim SORENSEN le 28/11/11
    End If
    

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim rdt_pompe As Double
    Dim rdt_moteur_pompe As Double
    Dim volume_corrige As Double
    Dim hauteur_bassin_max As Double

    'concentration de sortie en NNH4   mg/L
    If outlet_NH4_force = True Then
        outlet_NH4 = outlet_NH4_f
    Else
        outlet_NH4 = 3
    End If

    'concentration de sortie en NNO3   mg/L
    If outlet_NO3_force = True Then
        outlet_NO3 = outlet_NO3_f
    Else
        If choix_post_denit <> cuve_non Then
            outlet_NO3 = 5
        Else
            outlet_NO3 = 10
        End If
    End If

    'concentration en NNO3 en sortie de préDN
    If choix_pre_denit <> cuve_non Then
        If pre_denit_NO3_force = True Then
            pre_denit_NO3 = pre_denit_NO3_f
        Else
            pre_denit_NO3 = 0.1
        End If
    End If

    'concentration en NNO3 du dernier tank aéré (juste avant recirculation MLSS et la post_denit)  en mg/L
    If NO3_last_aerated_tank_force = True Then
        NO3_last_aerated_tank = NO3_last_aerated_tank_f
    Else
        If choix_post_denit <> cuve_non Then
            NO3_last_aerated_tank = outlet_NO3 + delta_NO3_post_DN
        Else
            NO3_last_aerated_tank = outlet_NO3
        End If
    End If

    'N organique dure et soluble qui sortira par la file eau
    If Norga_dure_soluble_force = True Then
        Norga_dure_soluble = Norga_dure_soluble_f
    Else
        Norga_dure_soluble = 1.5
    End If

    'DBO soluble en sortie   mg/L
    If outlet_DBO_soluble_force = True Then
        outlet_DBO_soluble = outlet_DBO_soluble_f
    Else
        outlet_DBO_soluble = 2
    End If

    'Taux de recirculation des boues du clarif  %
    If recirculation_taux_force = True Then
        recirculation_taux = recirculation_taux_f
    Else
        recirculation_taux = 1
    End If

    'concentration en MES dans les bassins au nominal
    If nominal_MES_concentration_force = True Then
        nominal_MES_concentration = nominal_MES_concentration_f
    Else
        If choix_fct_Mox = True Then
            nominal_MES_concentration = 1
        Else
            nominal_MES_concentration = 4
        End If
    End If

    'O2 dissous dans la cuve de DéOx  mg/L
    If deox_O2_dissous_force = True Then
        deox_O2_dissous = deox_O2_dissous_f
    Else
        deox_O2_dissous = 1
    End If


    'O2 dissous dans la cuve de RéOx  mg/L
    If reox_O2_dissous_force = True Then
        reox_O2_dissous = reox_O2_dissous_f
    Else
        reox_O2_dissous = 3 + md_ConcentrationOxygeneDissousExcesRegulation_mgL(choix_regulation)
    End If
    If reox_O2_facteur_alfa_force = True Then
        reox_O2_facteur_alfa = reox_O2_facteur_alfa_f
    Else
        If choix_aerateur_reox = aerateur_fine_bulle Then
            reox_O2_facteur_alfa = 0.65
        Else
            reox_O2_facteur_alfa = 0.8
        End If
    End If
    If reox_O2_rdt_transfert_force = True Then
        reox_O2_rdt_transfert = reox_O2_rdt_transfert_f
    Else
        If choix_aerateur_reox = aerateur_fine_bulle Then
            reox_O2_rdt_transfert = 16 / 300 * 100
        Else
            reox_O2_rdt_transfert = 12 / 300 * 100
        End If
    End If
    If reox_hauteur_bassin_force = True Then
        reox_hauteur_bassin = reox_hauteur_bassin_f
    Else
        reox_hauteur_bassin = 8
    End If
    hauteur_bassin_max = reox_hauteur_bassin

    'taux de remplissage cuve post-DN  %
    If post_denit_media_filling_force = True Then
        post_denit_media_filling = post_denit_media_filling_f
        If post_denit_media_filling > media_filling_max_PDN(choix_media_post_denit) Then
            'If remarque(erreur_media_filling) <> "" Then
            '    remarque(erreur_media_filling) = remarque(erreur_media_filling) + vbCrLf
            'End If
            remarque(erreur_media_filling) = Feuil15.Range("HybAS_erreur_filling") + "Post DN" + Feuil15.Range("HybAS_erreur_filling_2")
        End If
    Else
        post_denit_media_filling = 0.4
    End If

    'ratio nit média / nit total  %
    If ratio_Nnit_media_Nnit_total_force = True Then
        ratio_Nnit_media_Nnit_total = ratio_Nnit_media_Nnit_total_f
    Else
        If nb_cuves_CN <> 0 Then
            ratio_Nnit_media_Nnit_total = 0.6  'valeur modifiable car recalculée lorsque 2 cuves H
        End If
    End If

    'ratio nit media H1 / nit media  %
    If nb_cuves_H = 2 Then
        If ratio_Nnit_mediaH1_Nnit_media_force = True Then
            ratio_Nnit_mediaH1_Nnit_media = ratio_Nnit_mediaH1_Nnit_media_f
        Else
            ratio_Nnit_mediaH1_Nnit_media = 0.5
        End If
    ElseIf nb_cuves_CN = 0 Then
        ratio_Nnit_mediaH1_Nnit_media = 1
    End If

    'ratio nit CN1 / nit CN1&2
    If nb_cuves_CN = 2 Then
        If ratio_Nnit_CN1_Nnit_CN12_force = True Then
            ratio_Nnit_CN1_Nnit_CN12 = ratio_Nnit_CN1_Nnit_CN12_f
        Else
            ratio_Nnit_CN1_Nnit_CN12 = 0.6
        End If
    End If

    'taux de remplissage des cuves centrales hybrides  %   et oxygène dissous  mg/L
    For I = 1 To choix_nb_cuves_aerees
        If choix_type_cuve_aeree(I) = cuve_aeree_H Then
            If cuve_aeree_filling_force(I) = True Then
                cuve_aeree_filling(I) = cuve_aeree_filling_f(I)
                If cuve_aeree_filling(I) > media_filling_max(choix_media_cuve_aeree(I)) Then
                    If remarque(erreur_media_filling) <> "" Then
                        remarque(erreur_media_filling) = remarque(erreur_media_filling) + vbCrLf
                    End If
                    remarque(erreur_media_filling) = remarque(erreur_media_filling) + Feuil15.Range("HybAS_erreur_filling") + CStr(I) + Feuil15.Range("HybAS_erreur_filling_2")
                End If
            Else
                cuve_aeree_filling(I) = 0.5
            End If
        End If
        If cuve_aeree_O2_dissous_force(I) = True Then
            cuve_aeree_O2_dissous(I) = cuve_aeree_O2_dissous_f(I)
        Else
            If I = 1 And choix_type_cuve_aeree(I) = cuve_aeree_CN Then
                cuve_aeree_O2_dissous(I) = 2 + md_ConcentrationOxygeneDissousExcesRegulation_mgL(choix_regulation)
            ElseIf choix_fct_Mox = True Then
                cuve_aeree_O2_dissous(I) = 0.5
            Else
                cuve_aeree_O2_dissous(I) = 3 + md_ConcentrationOxygeneDissousExcesRegulation_mgL(choix_regulation)
            End If
        End If
        If cuve_aeree_hauteur_bassin_force(I) = True Then
            cuve_aeree_hauteur_bassin(I) = cuve_aeree_hauteur_bassin_f(I)
        Else
            If choix_fct_Mox = True Then
                cuve_aeree_hauteur_bassin(I) = 7
            Else
                cuve_aeree_hauteur_bassin(I) = 8
            End If
        End If
        If hauteur_bassin_max < cuve_aeree_hauteur_bassin(I) Then
            hauteur_bassin_max = cuve_aeree_hauteur_bassin(I)
        End If
        If cuve_aeree_O2_facteur_alfa_force(I) = True Then
            cuve_aeree_O2_facteur_alfa(I) = cuve_aeree_O2_facteur_alfa_f(I)
        Else
            If choix_aerateur_cuve_aeree(I) = aerateur_fine_bulle Then
                If choix_type_cuve_aeree(I) = cuve_aeree_CN Then
                    cuve_aeree_O2_facteur_alfa(I) = 0.65
                Else
                    cuve_aeree_O2_facteur_alfa(I) = 0.75
                End If
            Else
                If choix_type_cuve_aeree(I) = cuve_aeree_CN Then
                    cuve_aeree_O2_facteur_alfa(I) = 0.8
                Else
                    cuve_aeree_O2_facteur_alfa(I) = 0.8
                End If
            End If
        End If
        If cuve_aeree_O2_rdt_transfert_force(I) = True Then
            cuve_aeree_O2_rdt_transfert(I) = cuve_aeree_O2_rdt_transfert_f(I)
        Else
            If choix_aerateur_cuve_aeree(I) = aerateur_fine_bulle Then
                If choix_type_cuve_aeree(I) = cuve_aeree_CN Then
                    cuve_aeree_O2_rdt_transfert(I) = 16 / 300 * 100    '16 gO2/(Nm3 air.m) donc on divise par 300 gO2/Nm3 air
                Else
                    cuve_aeree_O2_rdt_transfert(I) = 16 / 300 * 100
                End If
            Else
                If choix_type_cuve_aeree(I) = cuve_aeree_CN Then
                    cuve_aeree_O2_rdt_transfert(I) = 12 / 300 * 100
                Else
                    cuve_aeree_O2_rdt_transfert(I) = 12 / 300 * 100
                End If
            End If
        End If
    Next I


    '******* vérification O2 dissous entre DéOx et cuves centrales
    If choix_deox <> cuve_non And deox_O2_dissous > cuve_aeree_O2_dissous(choix_nb_cuves_aerees) Then
        MsgBox Feuil15.Range("HybAS_erreur_deox_O2_dissous"), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If

    'agitation des cuves (W/m3)
    If ratio_elec_pre_anoxie_force = True Then
        ratio_elec_pre_anoxie = ratio_elec_pre_anoxie_f
    Else
        ratio_elec_pre_anoxie = 3
    End If
    If ratio_elec_anaerobie_force = True Then
        ratio_elec_anaerobie = ratio_elec_anaerobie_f
    Else
        ratio_elec_anaerobie = 10
    End If
    If ratio_elec_pre_denit_force = True Then
        ratio_elec_pre_denit = ratio_elec_pre_denit_f
    Else
        ratio_elec_pre_denit = 3
    End If
    If ratio_elec_deox_force = True Then
        ratio_elec_deox = ratio_elec_deox_f
    Else
        ratio_elec_deox = 3
    End If
    If ratio_elec_post_denit_force = True Then
        ratio_elec_post_denit = ratio_elec_post_denit_f
    Else
        If choix_post_denit = cuve_oui_sans_media Then
            ratio_elec_post_denit = 3
        ElseIf choix_post_denit = cuve_oui_avec_media Then
            ratio_elec_post_denit = 7.5
        End If
    End If

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

    'Indice de Mohlman (mL/g)
    If boues_indice_Mohlman_force = True Then
        boues_indice_Mohlman = boues_indice_Mohlman_f
    Else
        boues_indice_Mohlman = 100
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
        volume_corrige = boues_indice_Mohlman * nominal_MES_concentration
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
    NumeroErreurNonGeree = "15.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub dimensionnement(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef Sh, ByRef DCO_apportee_methanol, ByRef DCO_conso_bio)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer

    'variables intermédiaires sur les pollutions
    Dim interm_Q As Double
    Dim interm_DCO As Double
    Dim interm_DBO As Double
    Dim interm_MES As Double
    Dim interm_NK As Double
    Dim interm_NH4 As Double
    Dim interm_NO3 As Double
    Dim interm_Pt As Double

    Dim dicho_denit As Double
    Dim dicho_denit_min As Double
    Dim dicho_denit_max As Double

    Dim DCO_dure_mgL As Double

    Dim N_nit_total As Double    'N à nitrifier au total
    Dim N_nit_aerees As Double    'N à nitrifier cuves aérées centrales + DéOx
    Dim boues_pdtes As Double
    Dim ratio_DBO_denit As Double
    Dim pre_anoxie_denit As Double
    Dim O2_dissous_flux As Double
    Dim O2_dissous_MLSS_flux As Double
    Dim cuve_aeree_O2_besoins() As Double
    ReDim cuve_aeree_O2_besoins(choix_nb_cuves_aerees)
    Dim temperature As Double
    Dim anaerobie_V_optimal As Double
    Dim anaerobie_V_Voptimal As Double
    Dim anaerobie_NO3in_concentration As Double
    Dim deltaP_bio As Double
    Dim deltaP_synthese As Double
    Dim boues_minerales_pdtes As Double
    Dim pre_denit_denit As Double
    Dim deox_nit As Double
    Dim deox_V_mini As Double
    Dim deox_O2_conso As Double
    Dim post_denit_denit As Double
    Dim post_denit_DCO_apportee As Double
    Dim cinetique_post_denit_recalc As Double
    Dim cuves_centrales_nit As Double
    Dim cuves_centrales_DBOin As Double
    Dim cuves_centrales_DBOout As Double
    Dim post_denit_NO3_out As Double
    Dim cinetique_postDN_media_T As Double
    Dim cuve_aeree_DBO_eliminee() As Double
    ReDim cuve_aeree_DBO_eliminee(choix_nb_cuves_aerees)
    Dim DBO_appliquee_kgDBO_kgMES_j As Double

    Dim O2_dissous_flux_inlet_cuves_aerees As Double
    Dim O2_dissous_flux_inlet_reox As Double
    Dim interm_O2_exces As Double

    Dim O2_besoin_cuve As Double
    Dim cuve_media As Boolean


    If Norga_dure_soluble_force = True Then
        Norga_dure_soluble = Norga_dure_soluble_f
    Else
        Norga_dure_soluble = ratio_Norgaduresoluble_DCOtot * (DCO_nominal / Q_nominal * 1000)
    End If

    air_Q_Nm3j = 0
    O2_besoin_total = 0

    'ATTRIBUTION DES VALEURS DES FLUX DE POLLUTUION A DES VARIABLES INTERMEDIAIRES
    'eau arrivant sur le procédé
    interm_Q = Q
    interm_DCO = DCO
    DCO_dure_mgL = DCO_nominal * ratio_DCO_dure_total / Q_nominal * 1000
    interm_DBO = DBO
    interm_MES = MES
    interm_NK = NK
    interm_NH4 = NK - Norga_dure_soluble * Q / 1000
    interm_NO3 = NO3
    interm_Pt = Pt
    temperature = pd_TemperatureEauDimensionnement_°C

    'on incrémente avec les valeurs de la recirculation
    interm_Q = interm_Q + recirculation_taux * Q
    interm_DCO = interm_DCO
    interm_DBO = interm_DBO + recirculation_taux * Q / 1000 * outlet_DBO_soluble
    interm_MES = interm_MES
    interm_NK = interm_NK
    interm_NH4 = interm_NH4 + outlet_NH4 * recirculation_taux * Q / 1000
    interm_NO3 = interm_NO3 + outlet_NO3 * recirculation_taux * Q / 1000
    interm_Pt = interm_Pt + Pt_garantie * recirculation_taux * Q / 1000
    O2_dissous_flux = recirculation_O2 * recirculation_taux * Q / 1000

    'CALCULS PRELIMINAIRES

    'calcul du ratio MV/MES
    If nominal_MV_MES_force = True Then
        nominal_MV_MES = nominal_MV_MES_f
    Else
        If choix_primaire = True Then
            nominal_MV_MES = ratio_ED_MV_MES
        Else
            nominal_MV_MES = a_ratio_EB_MV_MES(0) + a_ratio_EB_MV_MES(1) * MES / DCO + a_ratio_EB_MV_MES(2) * DCO / DBO
            If nominal_MV_MES < ratio_EB_MV_MES_min Then
                nominal_MV_MES = ratio_EB_MV_MES_min
            ElseIf nominal_MV_MES > ratio_EB_MV_MES_max Then
                nominal_MV_MES = ratio_EB_MV_MES_max
            End If
        End If
    End If

    'calcul du N à nitrifier au total
    N_nit_total = NK - Norga_dure_soluble * Q / 1000 - outlet_NH4 * Q / 1000 - ratio_N_assimile * (DBO - outlet_DBO_soluble * Q / 1000) - ratio_N_assimile_methanol * boues_methanol
    'calcul du N à nitrifier dans les cuves aérées
    N_nit_aerees = N_nit_total - reox_nit

    'calcul du taux de recirculation
    If nominal_recirculation_MLSS_taux_force = True Then
        nominal_recirculation_MLSS_taux = nominal_recirculation_MLSS_taux_f
        If NO3_last_aerated_tank_force = False Then
            NO3_last_aerated_tank = pre_denit_NO3 + ratio_NO3f_nit * N_nit_aerees / ((1 + recirculation_taux + nominal_recirculation_MLSS_taux) * Q / 1000)
        Else
            If (NO3_last_aerated_tank - (pre_denit_NO3 + N_nit_aerees * ratio_NO3f_nit / ((1 + recirculation_taux + nominal_recirculation_MLSS_taux) * Q / 1000))) ^ 2 > critere_convergence Then
                remarque(erreur_taux_recirc_MLSS) = Feuil15.Range("HybAS_erreur_taux_recirc_MLSS")
            Else
                remarque(erreur_taux_recirc_MLSS) = ""
            End If
        End If
    Else
        nominal_recirculation_MLSS_taux = (N_nit_aerees * ratio_NO3f_nit - (1 + recirculation_taux) * Q / 1000 * (NO3_last_aerated_tank - pre_denit_NO3)) / ((NO3_last_aerated_tank - pre_denit_NO3) * Q / 1000)
        If nominal_recirculation_MLSS_taux < 0 Then
            If choix_post_denit <> cuve_non Then
                If NO3_last_aerated_tank_force = False Then
                    NO3_last_aerated_tank = 0.9 * NO3_last_aerated_tank
                End If
            Else
                If outlet_NO3_force = False Then
                    outlet_NO3 = 0.9 * outlet_NO3
                    NO3_last_aerated_tank = outlet_NO3
                End If
            End If
        End If
    End If

    'production de boues
    boues_pdtes = ratio_boues_MESin * MES + ratio_boues_DBOin * DBO + ratio_boues_nit * N_nit_total + boues_methanol
    stockage_boues_pdtes_nominal = boues_pdtes

    rate_bacteries_nit = ratio_boues_nit * N_nit_total / boues_pdtes

    'PRE-ANOXIE
    If choix_pre_anoxie <> cuve_non Then
        If pre_anoxie_volume_force = True Then
            pre_anoxie_volume = pre_anoxie_volume_f
            'on regarde comment on traite les NO3
            pre_anoxie_denit = interm_NO3 + ratio_NO3eq_O2 * O2_dissous_flux    '- pre_anoxie_NO3 * interm_Q / 1000
            dicho_denit_min = 0
            dicho_denit_max = pre_anoxie_denit
            dicho_denit = pre_anoxie_denit / 2
            Do While (dicho_denit_max - dicho_denit_min) / dicho_denit > critere_convergence
                ratio_DBO_denit = DBO / dicho_denit
                pre_anoxie_denit = pre_anoxie_volume * (pd_NOMBRE_HEURE_PAR_JOUR * cinetique_denit(ratio_DBO_denit, temperature) * nominal_MES_concentration * nominal_MV_MES) / 1000
                If pre_anoxie_denit < dicho_denit Then
                    dicho_denit_max = dicho_denit
                ElseIf pre_anoxie_denit > dicho_denit Then
                    dicho_denit_min = dicho_denit
                End If
                dicho_denit = (dicho_denit_max + dicho_denit_min) / 2
            Loop
            pre_anoxie_denit = dicho_denit
        Else
            pre_anoxie_denit = interm_NO3 + ratio_NO3eq_O2 * O2_dissous_flux - pre_anoxie_NO3 * interm_Q / 1000
            ratio_DBO_denit = DBO / pre_anoxie_denit
            If ratio_DBO_denit > ratio_DBO_denit_max Then
                ratio_DBO_denit = ratio_DBO_denit_max
            End If
            pre_anoxie_volume = 1000 * pre_anoxie_denit / (pd_NOMBRE_HEURE_PAR_JOUR * cinetique_denit(ratio_DBO_denit, temperature) * nominal_MES_concentration * nominal_MV_MES)
        End If
        interm_NO3 = interm_NO3 + ratio_NO3eq_O2 * O2_dissous_flux - pre_anoxie_denit
        interm_DBO = interm_DBO + ratio_O2_denit * pre_anoxie_denit
        O2_dissous_flux = 0
        'DCO_conso_bio = DCO_conso_bio - ratio_O2_denit * pre_anoxie_denit   'attention ratio_O2_denit est négatif --> -
    Else
        pre_anoxie_volume = 0
        pre_anoxie_denit = 0
    End If


    'ANAEROBIE
    If choix_anaerobie <> cuve_non Then
        If anaerobie_volume_force = True Then
            anaerobie_volume = anaerobie_volume_f
        Else
            If Pt / Q * 1000 < anaerobie_P_limite Then   'eau brute
                anaerobie_volume = anaerobie_HRT_P_low * Q / pd_NOMBRE_HEURE_PAR_JOUR
            Else
                anaerobie_volume = anaerobie_HRT_P_high * Q / pd_NOMBRE_HEURE_PAR_JOUR
            End If
        End If
        'volume anaérobie optimal
        If Pt / Q * 1000 < anaerobie_P_limite Then   'eau brute
            anaerobie_V_optimal = anaerobie_HRT_P_low * Q / pd_NOMBRE_HEURE_PAR_JOUR
        Else
            anaerobie_V_optimal = anaerobie_HRT_P_high * Q / pd_NOMBRE_HEURE_PAR_JOUR
        End If
        anaerobie_V_Voptimal = anaerobie_volume / anaerobie_V_optimal
        anaerobie_NO3in_concentration = interm_NO3 / interm_Q * 1000
        'P traité par la biologie
        deltaP_bio = interm_Pt * anaerobie_Kp(anaerobie_V_Voptimal, anaerobie_NO3in_concentration, interm_Pt, DBO)
        interm_DBO = interm_DBO - anaerobie_DBO_Pt * deltaP_bio
        'DCO_conso_bio = DCO_conso_bio + anaerobie_DBO_Pt * deltaP_bio
        'P éliminé par synthèse
        deltaP_synthese = ratio_P_assimile * (DBO - DBO_sortie)
        'P à traiter par la physico-chimie
        deltaP_precipitation = interm_Pt - (deltaP_bio + deltaP_synthese + Pt_garantie * Q * (1 + recirculation_taux) / 1000)
        If deltaP_precipitation > 0 Then
            'FeCl3 consommé
            anaerobie_FeCl3_flux = ratio_molaire_precipitation(deltaP_precipitation, Q) * MM_FeCl3 / MM_P * deltaP_precipitation
            'production de boues minerales
            boues_minerales_pdtes = (ratio_molaire_precipitation(deltaP_precipitation, Q) - 1) * MM_FeO3H3 / MM_P + MM_FePO4 / MM_P
            boues_minerales_pdtes = boues_minerales_pdtes * deltaP_precipitation
        Else
            deltaP_precipitation = 0
            anaerobie_FeCl3_flux = 0
            boues_minerales_pdtes = 0
        End If
        'on calcule le P sortant
        interm_Pt = interm_Pt - (deltaP_synthese + deltaP_precipitation + deltaP_bio)
        O2_dissous_flux = 0
    Else
        anaerobie_volume = 0
        'P éliminé par synthèse
        deltaP_synthese = ratio_P_assimile * (DBO - DBO_sortie)
        deltaP_precipitation = 0
        interm_Pt = interm_Pt - deltaP_synthese
    End If



    'RETOUR DE LA LIQUEUR MIXTE
    'O2 dissous
    If choix_deox <> cuve_non Then
        O2_dissous_MLSS_flux = deox_O2_dissous * nominal_recirculation_MLSS_taux * Q / 1000
    Else
        O2_dissous_MLSS_flux = cuve_aeree_O2_dissous(choix_nb_cuves_aerees) * nominal_recirculation_MLSS_taux * Q / 1000
    End If
    O2_dissous_flux = O2_dissous_flux + O2_dissous_MLSS_flux
    'NNO3
    interm_NO3 = interm_NO3 + NO3_last_aerated_tank * nominal_recirculation_MLSS_taux * Q / 1000
    interm_DBO = interm_DBO + nominal_recirculation_MLSS_taux * Q / 1000 * outlet_DBO_soluble



    'CUVE DE PRE-DENIT
    If choix_pre_denit <> cuve_non Then
        If pre_denit_volume_force = True Then
            pre_denit_volume = pre_denit_volume_f
            'on regarde comment on traite les NO3
            pre_denit_denit = interm_NO3 + ratio_NO3eq_O2 * O2_dissous_flux    '- pre_denit_NO3 * interm_Q / 1000
            dicho_denit_min = 0
            dicho_denit_max = pre_denit_denit
            dicho_denit = pre_denit_denit / 2
            Do While (dicho_denit_max - dicho_denit_min) / dicho_denit > critere_convergence
                ratio_DBO_denit = DBO / dicho_denit
                pre_denit_denit = pre_denit_volume * (pd_NOMBRE_HEURE_PAR_JOUR * cinetique_denit(ratio_DBO_denit, temperature) * nominal_MES_concentration * nominal_MV_MES) / 1000
                If pre_denit_denit < dicho_denit Then
                    dicho_denit_max = dicho_denit
                ElseIf pre_denit_denit > dicho_denit Then
                    dicho_denit_min = dicho_denit
                End If
                dicho_denit = (dicho_denit_max + dicho_denit_min) / 2
            Loop
            pre_denit_denit = dicho_denit
        Else
            pre_denit_denit = interm_NO3 + ratio_NO3eq_O2 * O2_dissous_flux - pre_denit_NO3 * (Q * (1 + recirculation_taux + nominal_recirculation_MLSS_taux)) / 1000
            ratio_DBO_denit = DBO / pre_denit_denit
            If ratio_DBO_denit > ratio_DBO_denit_max Then
                ratio_DBO_denit = ratio_DBO_denit_max
            End If
            pre_denit_volume = 1000 * pre_denit_denit / (pd_NOMBRE_HEURE_PAR_JOUR * cinetique_denit(ratio_DBO_denit, temperature) * nominal_MES_concentration * nominal_MV_MES)
        End If
        interm_NO3 = interm_NO3 + ratio_NO3eq_O2 * O2_dissous_flux - pre_denit_denit
        interm_DBO = interm_DBO + ratio_O2_denit * pre_denit_denit
        O2_dissous_flux = 0
        'DCO_conso_bio = DCO_conso_bio - ratio_O2_denit * pre_denit_denit
    Else
        pre_denit_volume = 0
        pre_denit_denit = 0
    End If

    'on récupère cette valeur pour les dimensionnements des cuves centrales
    cuves_centrales_DBOin = interm_DBO
    O2_dissous_flux_inlet_cuves_aerees = O2_dissous_flux


    ''''ATTENTION: L'ORDRE NE CORRESPOND PLUS A LA REALITE'''


    'CUVE DE DE-OXYGENATION
    If choix_deox <> cuve_non Then
        If deox_volume_force = True Then
            deox_volume = deox_volume_f
            deox_O2_conso = (1 + recirculation_taux + nominal_recirculation_MLSS_taux) * Q / 1000 * (cuve_aeree_O2_dissous(choix_nb_cuves_aerees) - deox_O2_dissous)
            deox_nit = (deox_O2_conso - respiration_gL(temperature, nominal_MES_concentration) * deox_volume) / ratio_O2_nit
        Else
            deox_O2_conso = (1 + recirculation_taux + nominal_recirculation_MLSS_taux) * Q / 1000 * (cuve_aeree_O2_dissous(choix_nb_cuves_aerees) - deox_O2_dissous)
            deox_volume = deox_O2_conso / (ratio_O2_nit * nominal_MES_concentration * pd_NOMBRE_HEURE_PAR_JOUR / 1000 * cinetique_nit_MLSS(deox_O2_dissous, boues_pdtes, N_nit_aerees, temperature) + respiration_gL(temperature, nominal_MES_concentration))
            deox_nit = (deox_O2_conso - respiration_gL(temperature, nominal_MES_concentration) * deox_volume) / ratio_O2_nit
        End If
        If deox_nit < 0 Then
            deox_nit = 0
        End If
        'message pour volume minimum conseillé
        deox_V_mini = (1 + recirculation_taux + nominal_recirculation_MLSS_taux) * Q * deox_HRT_mini / pd_NOMBRE_HEURE_PAR_JOUR
        If deox_volume < deox_V_mini Then
            remarque(erreur_deox_volume_insuffisant) = Feuil15.Range("HybAS_erreur_deox_volume_insuffisant") & Int(deox_V_mini)
        Else
            remarque(erreur_deox_volume_insuffisant) = ""
        End If
    Else
        deox_volume = 0
        deox_nit = 0
    End If


    'CUVE DE RE-OXYGENATION
    If choix_reox <> cuve_non Then
        If choix_post_denit <> cuve_non Then
            O2_dissous_flux_inlet_reox = 0
        ElseIf choix_deox <> cuve_non Then
            O2_dissous_flux_inlet_reox = deox_O2_dissous
        Else
            O2_dissous_flux_inlet_reox = cuve_aeree_O2_dissous(choix_nb_cuves_aerees)
        End If
        interm_O2_exces = (reox_O2_dissous - O2_dissous_flux_inlet_reox) * Q / 1000 * (1 + recirculation_taux)
        If reox_volume_force = True Then
            reox_volume = reox_volume_f
        Else
            reox_volume = reox_HRT * Q
        End If
        reox_nit = cinetique_nit_MLSS(reox_O2_dissous, boues_pdtes, N_nit_aerees, temperature) * reox_volume * pd_NOMBRE_HEURE_PAR_JOUR * nominal_MES_concentration / 1000
        If reox_nit > ((1 + recirculation_taux) * Q / 1000 * (reox_O2_dissous - O2_dissous_flux_inlet_reox)) / ratio_O2_nit Then    'N nitrifiable maximum
            reox_nit = ((1 + recirculation_taux) * Q / 1000 * (reox_O2_dissous - O2_dissous_flux_inlet_reox)) / ratio_O2_nit
        End If
        O2_besoin_cuve = ratio_O2_nit * reox_nit + interm_O2_exces
        O2_besoin_total = O2_besoin_total + O2_besoin_cuve
        air_Q_Nm3j = air_Q_Nm3j + air_Q_Nm3j_cuve(reox_O2_dissous, O2_besoin_cuve, temperature, reox_volume, False, reox_hauteur_bassin, reox_O2_facteur_alfa, reox_O2_rdt_transfert)
    Else
        reox_volume = 0
        reox_nit = 0
    End If


    'CUVE POST DENIT
    If choix_post_denit <> cuve_non Then
        O2_dissous_MLSS_flux = O2_dissous_MLSS_flux * (1 + recirculation_taux) / nominal_recirculation_MLSS_taux
        If post_denit_volume_force = True Then
            post_denit_volume = post_denit_volume_f
            'calcul de la charge à dénitrifier et de la DCO à apporter  ********** ici on ne recalcule pas la capacité de traitement
            post_denit_denit = (NO3_last_aerated_tank - outlet_NO3) * (1 + recirculation_taux) * Q / 1000 + reox_nit + ratio_NO3eq_O2 * O2_dissous_MLSS_flux
            post_denit_DCO_apportee = rate_DCO_apportee_denit_nominal * post_denit_denit
            post_denit_methanol_flux = post_denit_DCO_apportee / rate_DCO_methanol
            If choix_post_denit = cuve_oui_sans_media Then
                post_denit_media_filling = 0
                'on regarde si l'âge de boues est suffisant
                If (post_denit_volume * nominal_MES_concentration / boues_pdtes) < post_DN_SRT_mini Then
                    remarque(erreur_post_DN_volume_insuffisant) = Feuil15.Range("HybAS_erreur_post_DN_volume_insuffisant") & Int(post_DN_SRT_mini * boues_pdtes / nominal_MES_concentration)
                Else
                    remarque(erreur_post_DN_volume_insuffisant) = ""
                End If
                'on vérifie si il ne faut pas conseiller de mettre du média
                cinetique_post_denit_recalc = (post_denit_denit * 1000) / (pd_NOMBRE_HEURE_PAR_JOUR * post_denit_volume * nominal_MES_concentration * nominal_MV_MES)
                cinetique_post_denit_recalc = cinetique_post_denit_recalc / a_cinetique_postDN_nominal    'à 15°C
                If cinetique_post_denit_recalc < cinetique_post_denit_15_mini Then
                    remarque(erreur_postDN_media) = Feuil15.Range("HybAS_erreur_postDN_media")
                Else
                    remarque(erreur_postDN_media) = ""
                End If
                stockage_cinetique_post_DN_15 = cinetique_post_denit_recalc
            ElseIf choix_post_denit = cuve_oui_avec_media Then
                If post_denit_media_filling_force = False Then
                    'on calcule le taux de remplissage du média
                    post_denit_NO3_out = outlet_NO3 - reox_nit / (Q / 1000 * (1 + recirculation_taux))
                    cinetique_postDN_media_T = a_cinetique_postDN_media_10°C(1) * post_denit_NO3_out / (post_denit_NO3_out + a_cinetique_postDN_media_10°C(2))
                    cinetique_postDN_media_T = cinetique_postDN_media_T * a_cinetique_postDN_media_nominal
                    post_denit_media_filling = (1000 * post_denit_denit) / (post_denit_volume * media_surface_m2_m3(choix_media_post_denit) * cinetique_postDN_media_T)
                    If post_denit_media_filling > media_filling_max_PDN(choix_media_post_denit) Then
                        post_denit_media_filling = media_filling_max_PDN(choix_media_post_denit)
                    End If
                End If
            End If
        Else
            'calcul de la charge à dénitrifier et de la DCO à apporter
            post_denit_denit = (NO3_last_aerated_tank - outlet_NO3) * (1 + recirculation_taux) * Q / 1000 + reox_nit + ratio_NO3eq_O2 * O2_dissous_MLSS_flux
            post_denit_DCO_apportee = rate_DCO_apportee_denit_nominal * post_denit_denit
            post_denit_methanol_flux = post_denit_DCO_apportee / rate_DCO_methanol
            'on regarde s'il y a du média
            If choix_post_denit = cuve_oui_sans_media Then
                post_denit_media_filling = 0
                cinetique_post_denit_recalc = hyp_cinetique_post_denit_15 * a_cinetique_postDN_nominal
                post_denit_volume = post_denit_denit * 1000 / (pd_NOMBRE_HEURE_PAR_JOUR * cinetique_post_denit_recalc * nominal_MES_concentration * nominal_MV_MES)
                'on regarde si l'âge de boues est suffisant
                If (post_denit_volume * nominal_MES_concentration / boues_pdtes) < post_DN_SRT_mini Then
                    post_denit_volume = post_DN_SRT_mini * boues_pdtes / nominal_MES_concentration
                    cinetique_post_denit_recalc = (post_denit_denit * 1000) / (pd_NOMBRE_HEURE_PAR_JOUR * post_denit_volume * nominal_MES_concentration * nominal_MV_MES)
                    cinetique_post_denit_recalc = cinetique_post_denit_recalc / a_cinetique_postDN_nominal    'à 15°C
                    If cinetique_post_denit_recalc < cinetique_post_denit_15_mini Then
                        remarque(erreur_postDN_media) = Feuil15.Range("HybAS_erreur_postDN_media")
                    Else
                        remarque(erreur_postDN_media) = ""
                    End If
                End If
                stockage_cinetique_post_DN_15 = cinetique_post_denit_recalc
            ElseIf choix_post_denit = cuve_oui_avec_media Then
                post_denit_NO3_out = outlet_NO3 - reox_nit / (Q / 1000 * (1 + recirculation_taux))
                cinetique_postDN_media_T = a_cinetique_postDN_media_10°C(1) * post_denit_NO3_out / (post_denit_NO3_out + a_cinetique_postDN_media_10°C(2))
                cinetique_postDN_media_T = cinetique_postDN_media_T * a_cinetique_postDN_media_nominal
                post_denit_volume = 1000 * post_denit_denit / (post_denit_media_filling * media_surface_m2_m3(choix_media_post_denit) * cinetique_postDN_media_T)
            End If
        End If
        boues_methanol = ratio_boues_DCO_methanol * post_denit_DCO_apportee
        'DCO_conso_bio = DCO_conso_bio + post_denit_DCO_apportee * (1 - ratio_boues_DCO_methanol)
    Else
        post_denit_volume = 0
        post_denit_media_filling = 0
        post_denit_methanol_flux = 0
    End If


    'CUVES CENTRALES
    cuves_centrales_nit = N_nit_aerees - deox_nit
    'la valeur de DBO en entrée est récupérée plus haut après la pré-dénit
    cuves_centrales_DBOout = outlet_DBO_soluble * Q / 1000 * (1 + recirculation_taux + nominal_recirculation_MLSS_taux)
    'DCO_conso_bio = DCO_conso_bio + cuves_centrales_DBOin - cuves_centrales_DBOout
    If nb_cuves_CN = 0 Then   'configurations 4,6
        Call dimensionne_config_4_6(boues_pdtes, N_nit_aerees, temperature, cuves_centrales_nit, cuves_centrales_DBOin, cuve_aeree_O2_besoins)
        ratio_Nnit_media_Nnit_total = ratio_Nnit_media_Nnit_total / N_nit_aerees
    ElseIf nb_cuves_H = 2 Then    'configuration 3   (4 déjà faite ci-dessus)
        Call dimensionne_config_3(boues_pdtes, N_nit_aerees, temperature, cuves_centrales_nit, cuves_centrales_DBOin, cuve_aeree_O2_besoins)
    ElseIf nb_cuves_CN = 2 Then    'configuration 2
        Call dimensionne_config_2(boues_pdtes, N_nit_aerees, temperature, cuves_centrales_nit, cuves_centrales_DBOin, cuve_aeree_O2_besoins)
    Else    'configurations 1,5
        Call dimensionne_config_1_5(boues_pdtes, N_nit_aerees, temperature, cuves_centrales_nit, cuves_centrales_DBOin, cuve_aeree_O2_besoins)
    End If
    If pn_EffacementResultatsEnCours = True Then
        Exit Sub
    End If
    'on regarde si des valeurs de volume de cuves ont été forcées
    For I = 1 To choix_nb_cuves_aerees
        If cuve_aeree_volume_force(I) = True Then
            cuve_aeree_volume(I) = cuve_aeree_volume_f(I)
        End If
    Next I
    'on regarde si on élimine bien la DBO et on termine le calcul sur les besoins en O2
    For I = 1 To choix_nb_cuves_aerees
        interm_O2_exces = -O2_dissous_flux_inlet_cuves_aerees
        interm_DBO = cuves_centrales_DBOin
        cuve_aeree_DBO_eliminee(I) = interm_DBO
        If choix_type_cuve_aeree(I) = cuve_aeree_CN Then
            DBO_appliquee_kgDBO_kgMES_j = interm_DBO / (cuve_aeree_volume(I) * nominal_MES_concentration)
        Else
            DBO_appliquee_kgDBO_kgMES_j = interm_DBO / ((cuve_aeree_volume(I) * (1 - cuve_aeree_filling(I) * media_deplacement(choix_media_cuve_aeree(I)))) * nominal_MES_concentration)
        End If
        interm_DBO = interm_DBO * (1 - rdt_DBO(DBO_appliquee_kgDBO_kgMES_j))
        If interm_DBO < cuves_centrales_DBOout Then
            interm_DBO = cuves_centrales_DBOout
        End If
        cuve_aeree_DBO_eliminee(I) = cuve_aeree_DBO_eliminee(I) - interm_DBO
        O2_dissous_flux_inlet_cuves_aerees = cuve_aeree_O2_dissous(I) * Q / 1000 * (1 + recirculation_taux + nominal_recirculation_MLSS_taux)
        interm_O2_exces = interm_O2_exces + O2_dissous_flux_inlet_cuves_aerees
        cuve_aeree_O2_besoins(I) = cuve_aeree_O2_besoins(I) + interm_O2_exces
        cuve_aeree_O2_besoins(I) = cuve_aeree_O2_besoins(I) + ratio_O2_DBO * cuve_aeree_DBO_eliminee(I) + respiration_gL(temperature, nominal_MES_concentration) * cuve_aeree_volume(I)
        If choix_pre_anoxie = False And choix_pre_denit = False And choix_anaerobie = False Then
            cuve_aeree_O2_besoins(I) = cuve_aeree_O2_besoins(I) + besoins_O2_HS(Sh)
        End If
        cuves_centrales_DBOin = interm_DBO
    Next I
    For I = 1 To choix_nb_cuves_aerees
        O2_besoin_total = O2_besoin_total + cuve_aeree_O2_besoins(I)
        If choix_type_cuve_aeree(I) = cuve_aeree_H Then
            cuve_media = True
        Else
            cuve_media = False
        End If
        air_Q_Nm3j = air_Q_Nm3j + air_Q_Nm3j_cuve(cuve_aeree_O2_dissous(I), cuve_aeree_O2_besoins(I), temperature, cuve_aeree_volume(I), cuve_media, cuve_aeree_hauteur_bassin(I), cuve_aeree_O2_facteur_alfa(I), cuve_aeree_O2_rdt_transfert(I))
    Next I



    If interm_DBO > cuves_centrales_DBOout Then
        remarque(erreur_traitement_DBO_insuffisant) = Feuil15.Range("HybAS_erreur_traitement_DBO_insuffisant").Cells(1, 1)
    Else
        remarque(erreur_traitement_DBO_insuffisant) = ""
    End If


    'CALCUL DU VOLUME TOTAL DE BASSINS
    volume_total_bassins = pre_anoxie_volume + anaerobie_volume + pre_denit_volume + deox_volume + post_denit_volume + reox_volume
    For I = 1 To choix_nb_cuves_aerees
        volume_total_bassins = volume_total_bassins + cuve_aeree_volume(I)
    Next I

    'DIMENSIONNEMENT DU CLARIFICATEUR
    'surface de radier du clarificateur (m2)
    clarif_surface = (Q_nominal * pd_CoefficientPointeHydrauliqueTempsPluie + (Q - Q_nominal)) / (pd_NOMBRE_HEURE_PAR_JOUR * clarif_vitesse_max)
    If clarif_surface_force = True Then
        clarif_surface = clarif_surface_f
    End If
    'vitesse hydraulique max (recalculée) clarificateur
    clarif_vitesse_max_recalc = (Q_nominal * pd_CoefficientPointeHydrauliqueTempsPluie + (Q - Q_nominal)) / (pd_NOMBRE_HEURE_PAR_JOUR * clarif_surface)
    'charge au radier clarificateur
    clarif_charge_radier_nominal = nominal_MES_concentration * Q * (1 + recirculation_taux) / (pd_NOMBRE_HEURE_PAR_JOUR * clarif_surface)
    If clarif_charge_radier_nominal > charge_radier_max(choix_racleur) Then
        remarque(erreur_charge_radier) = Feuil15.Range("HybAS_erreur_charge_radier").Cells(1, 1)
    Else
        remarque(erreur_charge_radier) = ""
    End If
    'vérification de la capacité à décanter
    If boues_concentration > facteur_boues_C_lim / boues_indice_Mohlman Then
        remarque(erreur_taux_recirc) = Feuil15.Range("HybAS_erreur_taux_recirc").Cells(1, 1)
        If recirculation_taux_force = False Then
            recirculation_taux = recirculation_taux + 0.1
        End If
    Else
        remarque(erreur_taux_recirc) = ""
    End If

    DBO_sortie = interm_DBO    'stockée pour anaérobie

    'SORTIE EAU ET AUTRES DONNEES
    interm_Q = Q - boues_Q
    interm_MES = sortie_MES * interm_Q / 1000
    interm_DBO = (outlet_DBO_soluble + 0.6 * sortie_MES * nominal_MV_MES) * interm_Q / 1000
    interm_DCO = (DCO_dure_mgL + outlet_DBO_soluble + 1.45 * sortie_MES * nominal_MV_MES) * interm_Q / 1000
    interm_NH4 = outlet_NH4 * interm_Q / 1000
    interm_NK = (outlet_NH4 + Norga_dure_soluble) * interm_Q / 1000
    interm_NO3 = outlet_NO3 * interm_Q / 1000
    interm_Pt = interm_Pt

    boues_MES = (boues_pdtes + boues_minerales_pdtes) - interm_MES
    boues_concentration = ((1 + recirculation_taux) * nominal_MES_concentration - sortie_MES / 1000) / recirculation_taux
    boues_Q = boues_MES / boues_concentration

    DCO_apportee_methanol = DCO_apportee_methanol + post_denit_DCO_apportee

    'DCO_conso_bio = O2_besoin_total - ratio_O2_nit * N_nit_total


    'ON AFFECTE DE NOUVEAU LES VALEURS DES CHARGES
    Q = interm_Q
    DCO = interm_DCO
    DBO = interm_DBO
    MES = interm_MES
    NK = interm_NK
    NH4 = interm_NH4
    NO3 = interm_NO3
    Pt = interm_Pt


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub fonctionnement_reel(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef Sh, ByRef DCO_apportee_methanol, ByRef DCO_conso_bio)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    Dim cuve As Integer

    Dim iter_sludge As Integer
    Dim iter_MLSS As Integer

    Dim epsilone_boues As Double
    Dim epsilone_NH4 As Double
    Dim epsilone_NO3 As Double
    Dim epsilone_DBO As Double
    Dim epsilone_MLSS As Double
    Dim epsilone_MLSS_NH4 As Double
    Dim epsilone_MLSS_NO3 As Double
    Dim epsilone_postDN_NO3 As Double

    Dim dicho_denit As Double
    Dim dicho_denit_min As Double
    Dim dicho_denit_max As Double

    'variables intermédiaires sur les pollutions
    Dim interm_Q As Double
    Dim interm_DCO As Double
    Dim interm_DBO As Double
    Dim interm_MES As Double
    Dim interm_NK As Double
    Dim interm_NH4 As Double
    Dim interm_NO3 As Double
    Dim interm_Pt As Double

    Dim stockage_Q As Double
    Dim stockage_DCO As Double
    Dim stockage_DBO As Double
    Dim stockage_MES As Double
    Dim stockage_NK As Double
    Dim stockage_NH4 As Double
    Dim stockage_NO3 As Double
    Dim stockage_Pt As Double
    Dim stockage_O2_dissous_flux As Double

    Dim DCO_dure_mgL As Double

    Dim NO3_last_aerated_tank_reel As Double
    Dim NO3_pre_denit_reel As Double

    Dim interm_nit_aerees As Double

    Dim N_nit_total As Double    'N à nitrifier au total
    Dim N_nit_aerees As Double    'N à nitrifier cuves aérées centrales + DéOx
    Dim boues_pdtes As Double
    Dim ratio_DBO_denit As Double
    Dim pre_anoxie_denit As Double
    Dim O2_dissous_flux As Double
    Dim O2_dissous_MLSS_flux As Double
    Dim temperature As Double
    Dim anaerobie_V_optimal As Double
    Dim anaerobie_V_Voptimal As Double
    Dim anaerobie_NO3in_concentration As Double
    Dim deltaP_bio As Double
    Dim deltaP_synthese As Double
    Dim boues_minerales_pdtes As Double
    Dim pre_denit_denit As Double
    Dim deox_nit As Double
    Dim deox_O2_conso As Double
    Dim post_denit_denit As Double
    Dim post_denit_DCO_apportee As Double
    Dim cinetique_post_denit_recalc As Double
    Dim cuves_centrales_nit As Double
    Dim volume_central As Double
    Dim cuves_centrales_DBOout As Double
    Dim post_denit_NO3_out As Double
    Dim cinetique_postDN_media_T As Double
    Dim DBO_appliquee As Double
    Dim rate_N_MLSS As Double
    Dim cuves_centrales_DBOin As Double
    Dim interm_O2_besoin_total As Double
    Dim interm_air_Q_Nm3j As Double
    Dim interm_nit As Double
    Dim cuve_aeree_DBO_eliminee() As Double
    ReDim cuve_aeree_DBO_eliminee(choix_nb_cuves_aerees)
    Dim DBO_appliquee_kgDBO_kgMES_j As Double
    Dim cuve_media As Boolean
    Dim cuve_aeree_O2_besoins() As Double
    ReDim cuve_aeree_O2_besoins(choix_nb_cuves_aerees)
    Dim interm_O2_exces As Double
    Dim ratio_Q_MLSS_total As Double
    Dim deox_O2_dissous_reel As Double
    Dim dicho_deox_O2_min As Double
    Dim dicho_deox_O2_max As Double
    Dim dicho_deox_O2 As Double
    Dim deox_O2_nitrifiable As Double
    Dim outlet_reel_NK As Double
    Dim outlet_reel_Pt As Double
    Dim Norga_dure_soluble_reel As Double
    Dim delta_O2_besoins As Double
    Dim NH4_last_aerated_tank As Double

    Dim interm_DCO_conso_bio As Double
    Dim stockage_DCO_conso_bio As Double

    stockage_Q_reel = Q

    '********** non forçable au réel *******************
    Norga_dure_soluble_reel = ratio_Norgaduresoluble_DCOtot * (DCO_nominal * pd_PourcentageChargeReelleDCO / (Q_nominal * pd_PourcentageChargeReelleDebitVolumique) * 1000)


    DCO_dure_mgL = DCO_nominal * pd_PourcentageChargeReelleDCO * ratio_DCO_dure_total / (Q_nominal * pd_PourcentageChargeReelleDebitVolumique) * 1000
    NO3_last_aerated_tank_reel = NO3_last_aerated_tank
    NO3_pre_denit_reel = pre_denit_NO3
    temperature = pd_TemperatureEauConditionsReelles_°C

    'calcul du ratio MV/MES
    If reel_MV_MES_force = True Then
        reel_MV_MES = reel_MV_MES_f
    Else
        If choix_primaire = True Then
            reel_MV_MES = ratio_ED_MV_MES
        Else
            reel_MV_MES = a_ratio_EB_MV_MES(0) + a_ratio_EB_MV_MES(1) * MES / DCO + a_ratio_EB_MV_MES(2) * DCO / DBO
            If reel_MV_MES < ratio_EB_MV_MES_min Then
                reel_MV_MES = ratio_EB_MV_MES_min
            ElseIf reel_MV_MES > ratio_EB_MV_MES_max Then
                reel_MV_MES = ratio_EB_MV_MES_max
            End If
        End If
    End If

    'initialisation de Nit totale et Nit aérée ainsi que de la production de boues
    N_nit_total = NK - Norga_dure_soluble_reel * Q / 1000 - outlet_NH4 * Q / 1000 - ratio_N_assimile * (DBO - outlet_DBO_soluble * Q / 1000) - ratio_N_assimile_methanol * boues_methanol
    N_nit_aerees = N_nit_total - reox_nit
    boues_pdtes = ratio_boues_MESin * MES + ratio_boues_DBOin * DBO + ratio_boues_nit * N_nit_total + boues_methanol
    outlet_reel_NH4 = outlet_NH4
    outlet_reel_NK = outlet_reel_NH4 + Norga_dure_soluble_reel
    outlet_reel_NO3 = outlet_NO3
    outlet_reel_DBO = outlet_DBO_soluble
    deox_O2_dissous_reel = deox_O2_dissous

    epsilone_boues = critere_convergence + 1

    interm_Q = recirculation_taux * Q
    interm_DCO = recirculation_taux * Q / 1000 * (outlet_reel_DBO + DCO_dure_mgL)
    interm_DBO = recirculation_taux * Q / 1000 * outlet_reel_DBO
    interm_NK = outlet_reel_NK * recirculation_taux * Q / 1000
    interm_NH4 = outlet_reel_NH4 * recirculation_taux * Q / 1000
    interm_NO3 = outlet_reel_NO3 * recirculation_taux * Q / 1000
    interm_Pt = Pt_garantie * recirculation_taux * Q / 1000


    iter_sludge = 0

    Do While epsilone_boues > critere_convergence

        iter_sludge = iter_sludge + 1

        air_Q_Nm3j = 0
        O2_besoin_total = 0
        interm_DCO_conso_bio = 0

        'on récupère les valeurs pour le critère de convergence
        epsilone_NH4 = outlet_reel_NH4
        epsilone_NO3 = outlet_reel_NO3
        epsilone_DBO = outlet_reel_DBO
        
        If reel_MES_concentration_force = True Then
            reel_MES_concentration = reel_MES_concentration_f
        Else
            reel_MES_concentration = boues_pdtes / stockage_boues_pdtes_nominal * nominal_MES_concentration  'on estime la concentration dans les bassins par un ratio
        End If
        
        'ATTRIBUTION DES VALEURS DES FLUX DE POLLUTUION A DES VARIABLES INTERMEDIAIRES
        'eau arrivant sur le procédé
        interm_Q = Q + recirculation_taux * Q
        interm_DCO = DCO + recirculation_taux / (1 + recirculation_taux) * interm_DCO
        interm_DBO = DBO + recirculation_taux / (1 + recirculation_taux) * interm_DBO
        'interm_MES = MES
        interm_NK = NK + recirculation_taux / (1 + recirculation_taux) * interm_NK
        interm_NH4 = (NK - Norga_dure_soluble_reel * Q / 1000) + recirculation_taux / (1 + recirculation_taux) * interm_NH4
        interm_NO3 = NO3 + recirculation_taux / (1 + recirculation_taux) * interm_NO3
        interm_Pt = Pt + recirculation_taux / (1 + recirculation_taux) * interm_Pt

        'on incrémente avec les valeurs de la recirculation
        'interm_Q = interm_Q + recirculation_taux * Q
        'interm_DCO = interm_DCO + recirculation_taux * Q / 1000 * (outlet_reel_DBO + DCO_dure_mgL)
        'interm_DBO = interm_DBO + recirculation_taux * Q / 1000 * outlet_reel_DBO
        'interm_MES=
        'interm_NK = interm_NK + outlet_NK * recirculation_taux * Q / 1000
        'interm_NH4 = interm_NH4 + outlet_reel_NH4 * recirculation_taux * Q / 1000
        'interm_NO3 = interm_NO3 + outlet_reel_NO3 * recirculation_taux * Q / 1000
        'interm_Pt = interm_Pt + Pt_garantie * recirculation_taux * Q / 1000
        O2_dissous_flux = recirculation_O2 * recirculation_taux * Q / 1000



        'PRE-ANOXIE
        If choix_pre_anoxie <> cuve_non Then
            pre_anoxie_denit = interm_NO3 + ratio_NO3eq_O2 * O2_dissous_flux    '- pre_anoxie_NO3 * interm_Q / 1000
            dicho_denit_min = 0
            dicho_denit_max = pre_anoxie_denit
            dicho_denit = pre_anoxie_denit / 2
            Do While (dicho_denit_max - dicho_denit_min) / dicho_denit > critere_convergence
                ratio_DBO_denit = DBO / dicho_denit
                pre_anoxie_denit = pre_anoxie_volume * (pd_NOMBRE_HEURE_PAR_JOUR * cinetique_denit(ratio_DBO_denit, temperature) * reel_MES_concentration * reel_MV_MES) / 1000
                If pre_anoxie_denit < dicho_denit Then
                    dicho_denit_max = dicho_denit
                ElseIf pre_anoxie_denit > dicho_denit Then
                    dicho_denit_min = dicho_denit
                End If
                dicho_denit = (dicho_denit_max + dicho_denit_min) / 2
            Loop
            pre_anoxie_denit = dicho_denit
            interm_NO3 = interm_NO3 + ratio_NO3eq_O2 * O2_dissous_flux
            If interm_NO3 < pre_anoxie_denit Then
                pre_anoxie_denit = interm_NO3
                interm_NO3 = 0
            Else
                interm_NO3 = interm_NO3 - pre_anoxie_denit
            End If
            interm_DBO = interm_DBO + ratio_O2_denit * pre_anoxie_denit    'ratio_O2_denit négatif d'où le +
            interm_DCO = interm_DCO + ratio_O2_denit * pre_anoxie_denit   'ratio_O2_denit négatif d'où le +
            interm_Pt = interm_Pt
            interm_NK = interm_NK + ratio_N_assimile * ratio_O2_denit * pre_anoxie_denit  'ratio_O2_denit négatif d'où le +
            interm_NH4 = interm_NH4 + ratio_N_assimile * ratio_O2_denit * pre_anoxie_denit  'ratio_O2_denit négatif d'où le +
            O2_dissous_flux = 0
            interm_DCO_conso_bio = interm_DCO_conso_bio - ratio_O2_denit * pre_anoxie_denit   'attention ratio_O2_denit est négatif --> -
        Else
            pre_anoxie_denit = 0
        End If


        'ANAEROBIE
        If choix_anaerobie <> cuve_non Then
            'volume anaérobie optimal
            If Pt / Q * 1000 < anaerobie_P_limite Then
                anaerobie_V_optimal = anaerobie_HRT_P_low * Q / pd_NOMBRE_HEURE_PAR_JOUR
            Else
                anaerobie_V_optimal = anaerobie_HRT_P_high * Q / pd_NOMBRE_HEURE_PAR_JOUR
            End If
            anaerobie_V_Voptimal = anaerobie_volume / anaerobie_V_optimal
            anaerobie_NO3in_concentration = interm_NO3 / interm_Q * 1000
            'P traité par la biologie
            deltaP_bio = interm_Pt * anaerobie_Kp(anaerobie_V_Voptimal, anaerobie_NO3in_concentration, interm_Pt, DBO)
            interm_DBO = interm_DBO - anaerobie_DBO_Pt * deltaP_bio
            interm_DCO = interm_DCO - anaerobie_DBO_Pt * deltaP_bio
            interm_DCO_conso_bio = interm_DCO_conso_bio + anaerobie_DBO_Pt * deltaP_bio
            interm_NH4 = interm_NH4 - ratio_N_assimile * anaerobie_DBO_Pt * deltaP_bio
            interm_NK = interm_NK - ratio_N_assimile * anaerobie_DBO_Pt * deltaP_bio
            'P éliminé par synthèse
            deltaP_synthese = ratio_P_assimile * (DBO - DBO_sortie)
            'P à traiter par la physico-chimie
            deltaP_precipitation = interm_Pt - (deltaP_bio + deltaP_synthese + Pt_garantie * interm_Q / 1000)
            If deltaP_precipitation > 0 Then
                'FeCl3 consommé
                anaerobie_FeCl3_flux = ratio_molaire_precipitation(deltaP_precipitation, Q) * MM_FeCl3 / MM_P * deltaP_precipitation
                'production de boues minerales
                boues_minerales_pdtes = ((ratio_molaire_precipitation(deltaP_precipitation, Q) - 1) * MM_FeO3H3 / MM_P + MM_FePO4 / MM_P) * deltaP_precipitation
            Else
                anaerobie_FeCl3_flux = 0
                boues_minerales_pdtes = 0
                deltaP_precipitation = 0
            End If
            If anaerobie_FeCl3_flux_force = True Then
                anaerobie_FeCl3_flux = anaerobie_FeCl3_flux_f
            End If
            'on calcule le P sortant
            interm_Pt = interm_Pt - (deltaP_synthese + deltaP_precipitation + deltaP_bio)
            O2_dissous_flux = 0
        Else
            'P éliminé par synthèse
            deltaP_synthese = ratio_P_assimile * (DBO - DBO_sortie)
            deltaP_precipitation = 0
            interm_Pt = interm_Pt - deltaP_synthese
        End If

        'stockage des données de flux
        stockage_Q = interm_Q
        stockage_DCO = interm_DCO
        stockage_DBO = interm_DBO
        stockage_NK = interm_NK
        stockage_NH4 = interm_NH4
        stockage_NO3 = interm_NO3
        stockage_Pt = interm_Pt
        stockage_O2_dissous_flux = O2_dissous_flux
        stockage_DCO_conso_bio = interm_DCO_conso_bio

        epsilone_MLSS = critere_convergence + 1

        iter_MLSS = 0

        Do While epsilone_MLSS > critere_convergence

            iter_MLSS = iter_MLSS + 1

            interm_air_Q_Nm3j = air_Q_Nm3j
            interm_O2_besoin_total = O2_besoin_total

            epsilone_MLSS_NO3 = NO3_last_aerated_tank_reel
            epsilone_MLSS_NH4 = NH4_last_aerated_tank


            interm_nit_aerees = 0

            'RECIRCULATION DE MLSS        *************il faudra peut-être rajouter ici une deuxième boucle, en théorie, il vaudrait mieux mais on regarde d'abord le comportement**********
            'on calcule le nouveau taux de recirculation de MLSS
            If reel_recirculation_MLSS_taux_force = True Then
                reel_recirculation_MLSS_taux = reel_recirculation_MLSS_taux_f
                'NO3_last_aerated_tank_reel = NO3_pre_denit_reel + N_nit_aerees / ((1 + recirculation_taux + reel_recirculation_MLSS_taux) * Q / 1000)
            Else
                'reel_recirculation_MLSS_taux = (N_nit_aerees - (1 + recirculation_taux) * Q / 1000 * (NO3_last_aerated_tank_reel - NO3_pre_denit_reel)) / ((NO3_last_aerated_tank_reel - NO3_pre_denit_reel) * Q / 1000)
                reel_recirculation_MLSS_taux = nominal_recirculation_MLSS_taux
            End If
            If iter_MLSS > 1 Then
                'on incrémente avec les valeurs de la recirculation
                ratio_Q_MLSS_total = reel_recirculation_MLSS_taux / (1 + recirculation_taux + reel_recirculation_MLSS_taux)
                interm_Q = Q * (1 + recirculation_taux + reel_recirculation_MLSS_taux)                           '+ reel_recirculation_MLSS_taux * Q
                interm_DCO = stockage_DCO + ratio_Q_MLSS_total * interm_DCO                                 ' + reel_recirculation_MLSS_taux * Q / 1000 * outlet_reel_DBO
                interm_DBO = stockage_DBO + ratio_Q_MLSS_total * interm_DBO                                 ' + reel_recirculation_MLSS_taux * Q / 1000 * outlet_reel_DBO    'à vérifier
                interm_NK = stockage_NK + ratio_Q_MLSS_total * interm_NK                                 ' + (outlet_NK * reel_recirculation_MLSS_taux * Q / 1000 + reox_nit)
                interm_NH4 = stockage_NH4 + ratio_Q_MLSS_total * interm_NH4                                 ' + (outlet_reel_NH4 * reel_recirculation_MLSS_taux * Q / 1000 + reox_nit)
                interm_NO3 = stockage_NO3 + ratio_Q_MLSS_total * interm_NO3                                 ' + NO3_last_aerated_tank_reel * reel_recirculation_MLSS_taux * Q / 1000
                interm_Pt = stockage_Pt + ratio_Q_MLSS_total * interm_Pt                                 ' + Pt_garantie * reel_recirculation_MLSS_taux * Q / 1000
                interm_DCO_conso_bio = stockage_DCO_conso_bio
            Else
                interm_Q = Q * (1 + recirculation_taux + reel_recirculation_MLSS_taux)
                interm_DCO = stockage_DCO + reel_recirculation_MLSS_taux * Q / 1000 * (outlet_reel_DBO + DCO_dure_mgL)
                interm_DBO = stockage_DBO + reel_recirculation_MLSS_taux * Q / 1000 * outlet_reel_DBO     'à vérifier
                interm_NK = stockage_NK + (outlet_reel_NK * reel_recirculation_MLSS_taux * Q / 1000)
                interm_NH4 = stockage_NH4 + (outlet_reel_NH4 * reel_recirculation_MLSS_taux * Q / 1000)
                interm_NO3 = stockage_NO3 + NO3_last_aerated_tank_reel * reel_recirculation_MLSS_taux * Q / 1000
                interm_Pt = stockage_Pt + Pt_garantie * reel_recirculation_MLSS_taux * Q / 1000
                interm_DCO_conso_bio = stockage_DCO_conso_bio
            End If
            'O2 dissous
            If choix_deox <> cuve_non Then
                O2_dissous_MLSS_flux = deox_O2_dissous_reel * reel_recirculation_MLSS_taux * Q / 1000
            Else
                O2_dissous_MLSS_flux = cuve_aeree_O2_dissous(choix_nb_cuves_aerees) * reel_recirculation_MLSS_taux * Q / 1000
            End If
            O2_dissous_flux = stockage_O2_dissous_flux + O2_dissous_MLSS_flux


            'CUVE DE PRE-DENIT
            If choix_pre_denit <> cuve_non Then
                'on regarde comment on traite les NO3
                pre_denit_denit = interm_NO3 + ratio_NO3eq_O2 * O2_dissous_flux    '- pre_denit_NO3 * interm_Q / 1000
                dicho_denit_min = 0
                dicho_denit_max = pre_denit_denit
                dicho_denit = pre_denit_denit / 2
                Do While (dicho_denit_max - dicho_denit_min) / dicho_denit > critere_convergence
                    ratio_DBO_denit = DBO / dicho_denit
                    pre_denit_denit = pre_denit_volume * (pd_NOMBRE_HEURE_PAR_JOUR * cinetique_denit(ratio_DBO_denit, temperature) * reel_MES_concentration * reel_MV_MES) / 1000
                    If pre_denit_denit < dicho_denit Then
                        dicho_denit_max = dicho_denit
                    ElseIf pre_denit_denit > dicho_denit Then
                        dicho_denit_min = dicho_denit
                    End If
                    dicho_denit = (dicho_denit_max + dicho_denit_min) / 2
                Loop
                pre_denit_denit = dicho_denit
                interm_NO3 = interm_NO3 + ratio_NO3eq_O2 * O2_dissous_flux
                If interm_NO3 < pre_denit_denit Then
                    pre_denit_denit = interm_NO3
                    interm_NO3 = 0
                Else
                    interm_NO3 = interm_NO3 - pre_denit_denit
                End If
                interm_DBO = interm_DBO + ratio_O2_denit * pre_denit_denit
                interm_DCO = interm_DCO + ratio_O2_denit * pre_denit_denit
                interm_NK = interm_NK + ratio_N_assimile * ratio_O2_denit * pre_denit_denit  'ratio_O2_denit négatif d'où le +
                interm_NH4 = interm_NH4 + ratio_N_assimile * ratio_O2_denit * pre_denit_denit  'ratio_O2_denit négatif d'où le +
                O2_dissous_flux = 0
                interm_DCO_conso_bio = interm_DCO_conso_bio - ratio_O2_denit * pre_denit_denit  'attention ratio_O2_denit est négatif --> -
            Else
                pre_denit_denit = 0
            End If

            'calcul de la concentration en PréDN
            NO3_pre_denit_reel = interm_NO3 / interm_Q * 1000


            'CUVES CENTRALES
            volume_central = 0
            For cuve = 1 To choix_nb_cuves_aerees
                volume_central = volume_central + cuve_aeree_volume(cuve)
            Next cuve
            DBO_appliquee = interm_DBO / (volume_central * reel_MES_concentration)    ' * reel_MV_MES)  'pour la nit
            cuves_centrales_DBOin = interm_DBO
            interm_DCO_conso_bio = interm_DCO_conso_bio + cuves_centrales_DBOin
            cuves_centrales_DBOout = outlet_DBO_soluble * Q / 1000 * (1 + recirculation_taux + reel_recirculation_MLSS_taux)   'le minimum en dessous duquel on ne peut pas aller
            cuves_centrales_nit = 0
            For cuve = 1 To choix_nb_cuves_aerees
                interm_O2_exces = -O2_dissous_flux  'négatif car ensuite on fera sortie-entrée pour calculer l'excès d'O2 transféré
                If choix_type_cuve_aeree(cuve) = cuve_aeree_H Then
                    cuve_media = True
                    rate_N_MLSS = cinetique_nit_MLSS(cuve_aeree_O2_dissous(cuve), boues_pdtes, N_nit_aerees, temperature)
                    interm_nit = nitrification_H_kgj(DBO_appliquee, temperature, cuve, rate_N_MLSS, reel_MES_concentration, False)
                Else    'cuve de type C/N
                    cuve_media = False
                    rate_N_MLSS = cinetique_nit_MLSS(cuve_aeree_O2_dissous(cuve), boues_pdtes, N_nit_aerees, temperature)
                    interm_nit = pd_NOMBRE_HEURE_PAR_JOUR / 1000 * rate_N_MLSS * reel_MES_concentration * cuve_aeree_volume(cuve)
                End If
                'on regarde la DBO
                interm_DBO = cuves_centrales_DBOin
                cuve_aeree_DBO_eliminee(cuve) = interm_DBO
                If choix_type_cuve_aeree(cuve) = cuve_aeree_CN Then
                    DBO_appliquee_kgDBO_kgMES_j = interm_DBO / (cuve_aeree_volume(cuve) * reel_MES_concentration)
                Else
                    DBO_appliquee_kgDBO_kgMES_j = interm_DBO / ((cuve_aeree_volume(cuve) * (1 - cuve_aeree_filling(cuve) * media_deplacement(choix_media_cuve_aeree(cuve)))) * reel_MES_concentration)
                End If
                interm_DBO = interm_DBO * (1 - rdt_DBO(DBO_appliquee_kgDBO_kgMES_j))
                If interm_DBO < cuves_centrales_DBOout Then
                    interm_DBO = cuves_centrales_DBOout
                End If
                cuve_aeree_DBO_eliminee(cuve) = cuve_aeree_DBO_eliminee(cuve) - interm_DBO
                cuves_centrales_DBOin = interm_DBO
                interm_DCO = interm_DCO - cuve_aeree_DBO_eliminee(cuve)
                interm_NH4 = interm_NH4 - ratio_N_assimile * cuve_aeree_DBO_eliminee(cuve)
                interm_NK = interm_NK - ratio_N_assimile * cuve_aeree_DBO_eliminee(cuve)
                If interm_NH4 < interm_nit Then
                    interm_nit = interm_NH4
                    interm_NH4 = 0
                Else
                    interm_NH4 = interm_NH4 - interm_nit
                End If
                interm_NK = interm_NK - interm_nit
                interm_NO3 = interm_NO3 + interm_nit * ratio_NO3f_nit
                cuves_centrales_nit = cuves_centrales_nit + interm_nit
                'besoins en O2
                cuve_aeree_O2_besoins(cuve) = interm_nit * ratio_O2_nit + cuve_aeree_DBO_eliminee(cuve) * ratio_O2_DBO + cuve_aeree_volume(cuve) * respiration_gL(temperature, reel_MES_concentration)
                O2_dissous_flux = cuve_aeree_O2_dissous(cuve) * Q / 1000 * (1 + recirculation_taux + reel_recirculation_MLSS_taux)
                interm_O2_exces = interm_O2_exces + O2_dissous_flux
                cuve_aeree_O2_besoins(cuve) = cuve_aeree_O2_besoins(cuve) + interm_O2_exces
                If choix_pre_anoxie = False And choix_pre_denit = False And choix_anaerobie = False Then
                    cuve_aeree_O2_besoins(I) = cuve_aeree_O2_besoins(I) + besoins_O2_HS(Sh)
                End If
                interm_O2_besoin_total = interm_O2_besoin_total + cuve_aeree_O2_besoins(cuve)
                interm_air_Q_Nm3j = interm_air_Q_Nm3j + air_Q_Nm3j_cuve(cuve_aeree_O2_dissous(cuve), cuve_aeree_O2_besoins(cuve), temperature, cuve_aeree_volume(cuve), cuve_media, cuve_aeree_hauteur_bassin(cuve), cuve_aeree_O2_facteur_alfa(cuve), cuve_aeree_O2_rdt_transfert(cuve))
            Next cuve
            O2_dissous_flux = cuve_aeree_O2_dissous(choix_nb_cuves_aerees) / 1000 * Q * (1 + recirculation_taux + reel_recirculation_MLSS_taux)
            If interm_DBO > cuves_centrales_DBOout Then
                cuves_centrales_DBOout = interm_DBO
            End If
            interm_DCO_conso_bio = interm_DCO_conso_bio - cuves_centrales_DBOout
            'calcul de la nit aérée
            interm_nit_aerees = cuves_centrales_nit


            'DEOX
            If choix_deox <> cuve_non Then
                dicho_deox_O2_min = 0
                deox_O2_nitrifiable = cuve_aeree_O2_dissous(choix_nb_cuves_aerees) * interm_Q / 1000 - respiration_gL(temperature, reel_MES_concentration) * deox_volume   'on enlève la respiration
                dicho_deox_O2_max = deox_O2_nitrifiable
                If dicho_deox_O2_max > 0 Then
                    dicho_deox_O2 = (dicho_deox_O2_min + dicho_deox_O2_max) / 2
                    deox_O2_dissous_reel = cuve_aeree_O2_dissous(choix_nb_cuves_aerees) - (respiration_gL(temperature, reel_MES_concentration) * deox_volume + dicho_deox_O2) / (interm_Q / 1000)
                    Do While (dicho_deox_O2_max - dicho_deox_O2_min) / dicho_deox_O2 > critere_convergence
                        deox_nit = pd_NOMBRE_HEURE_PAR_JOUR / 1000 * reel_MES_concentration * cinetique_nit_MLSS(deox_O2_dissous_reel, boues_pdtes, N_nit_aerees, temperature) * deox_volume
                        deox_O2_conso = deox_nit * ratio_O2_nit
                        'deox_O2_dissous_reel = cuve_aeree_O2_dissous(choix_nb_cuves_aerees) - deox_O2_conso / interm_Q * 1000
                        If deox_O2_conso < dicho_deox_O2 Then
                            dicho_deox_O2_max = dicho_deox_O2
                        Else
                            dicho_deox_O2_min = dicho_deox_O2
                        End If
                        dicho_deox_O2 = (dicho_deox_O2_min + dicho_deox_O2_max) / 2
                        deox_O2_dissous_reel = cuve_aeree_O2_dissous(choix_nb_cuves_aerees) - (respiration_gL(temperature, reel_MES_concentration) * deox_volume + dicho_deox_O2) / (interm_Q / 1000)
                    Loop
                End If
                interm_NK = interm_NK - deox_nit
                interm_NH4 = interm_NH4 - deox_nit
                interm_NO3 = interm_NO3 + deox_nit
                interm_nit_aerees = interm_nit_aerees + deox_nit
                O2_dissous_flux = deox_O2_dissous_reel / 1000 * Q * (1 + recirculation_taux + reel_recirculation_MLSS_taux)
            End If
            NO3_last_aerated_tank_reel = interm_NO3 / (Q * (1 + recirculation_taux + reel_recirculation_MLSS_taux)) * 1000
            NH4_last_aerated_tank = interm_NH4 / (Q * (1 + recirculation_taux + reel_recirculation_MLSS_taux)) * 1000

            N_nit_aerees = interm_nit_aerees
            N_nit_total = N_nit_aerees + reox_nit
            boues_pdtes = ratio_boues_MESin * MES + ratio_boues_DBOin * DBO + ratio_boues_nit * N_nit_total + boues_methanol

            '********************** test de convergence ***********************
            'fait sur les NNO3
            epsilone_MLSS_NO3 = epsilone_MLSS_NO3 - NO3_last_aerated_tank_reel
            epsilone_MLSS_NO3 = epsilone_MLSS_NO3 ^ 2
            epsilone_MLSS_NH4 = epsilone_MLSS_NH4 - NH4_last_aerated_tank
            epsilone_MLSS_NH4 = epsilone_MLSS_NH4 ^ 2
            epsilone_MLSS = epsilone_MLSS_NO3 + epsilone_MLSS_NH4

        Loop


        air_Q_Nm3j = air_Q_Nm3j + interm_air_Q_Nm3j
        O2_besoin_total = O2_besoin_total + interm_O2_besoin_total

        'SEPARATION DES FLUX
        interm_Q = interm_Q * (1 + recirculation_taux) / (1 + recirculation_taux + reel_recirculation_MLSS_taux)
        interm_DBO = interm_DBO * (1 + recirculation_taux) / (1 + recirculation_taux + reel_recirculation_MLSS_taux)
        interm_DCO = interm_DCO * (1 + recirculation_taux) / (1 + recirculation_taux + reel_recirculation_MLSS_taux)
        interm_NK = interm_NK * (1 + recirculation_taux) / (1 + recirculation_taux + reel_recirculation_MLSS_taux)
        interm_NH4 = interm_NH4 * (1 + recirculation_taux) / (1 + recirculation_taux + reel_recirculation_MLSS_taux)
        interm_NO3 = interm_NO3 * (1 + recirculation_taux) / (1 + recirculation_taux + reel_recirculation_MLSS_taux)
        interm_Pt = interm_Pt * (1 + recirculation_taux) / (1 + recirculation_taux + reel_recirculation_MLSS_taux)
        O2_dissous_flux = O2_dissous_flux * (1 + recirculation_taux) / (1 + recirculation_taux + reel_recirculation_MLSS_taux)

        'POST-DENITRIFICATION
        If choix_post_denit = cuve_oui_sans_media Then
            interm_NO3 = interm_NO3 + ratio_NO3eq_O2 * O2_dissous_flux
            cinetique_post_denit_recalc = stockage_cinetique_post_DN_15 * a_cinetique_postDN_reel
            post_denit_denit = (post_denit_volume * pd_NOMBRE_HEURE_PAR_JOUR * cinetique_post_denit_recalc * reel_MES_concentration * reel_MV_MES) / 1000
            If post_denit_methanol_flux_force = True Then
                post_denit_methanol_flux = post_denit_methanol_flux_f
            Else
                post_denit_DCO_apportee = rate_DCO_apportee_denit_reel * post_denit_denit
                post_denit_methanol_flux = post_denit_DCO_apportee / rate_DCO_methanol
            End If
            boues_methanol = ratio_boues_DCO_methanol * post_denit_DCO_apportee
            interm_DCO_conso_bio = interm_DCO_conso_bio + post_denit_DCO_apportee * (1 - ratio_boues_DCO_methanol)
            interm_NO3 = interm_NO3 - post_denit_denit
            If ratio_N_assimile_methanol * boues_methanol > interm_NH4 Then
                interm_NO3 = interm_NO3 - (ratio_N_assimile_methanol * boues_methanol - interm_NH4)
                interm_NK = interm_NK - interm_NH4
                interm_NH4 = 0
            Else
                interm_NH4 = interm_NH4 - ratio_N_assimile_methanol * boues_methanol
                interm_NK = interm_NK - ratio_N_assimile_methanol * boues_methanol
            End If
            O2_dissous_flux = 0#
        ElseIf choix_post_denit = cuve_oui_avec_media Then
            interm_NO3 = interm_NO3 + ratio_NO3eq_O2 * O2_dissous_flux
            'itération pour calculer les NO3 en sortie post-dénit
            post_denit_NO3_out = outlet_NO3
            epsilone_postDN_NO3 = 1
            Do While epsilone_postDN_NO3 > critere_convergence
                epsilone_postDN_NO3 = post_denit_NO3_out
                cinetique_postDN_media_T = a_cinetique_postDN_media_10°C(1) * post_denit_NO3_out / (post_denit_NO3_out + a_cinetique_postDN_media_10°C(2))
                cinetique_postDN_media_T = cinetique_postDN_media_T * a_cinetique_postDN_media_reel
                post_denit_NO3_out = interm_NO3 - (post_denit_volume * post_denit_media_filling * media_surface_m2_m3(choix_media_post_denit) * cinetique_postDN_media_T) / 1000
                post_denit_NO3_out = post_denit_NO3_out / interm_Q * 1000
                epsilone_postDN_NO3 = epsilone_postDN_NO3 - post_denit_NO3_out
                epsilone_postDN_NO3 = epsilone_postDN_NO3 ^ 2
            Loop
            post_denit_denit = interm_NO3 - post_denit_NO3_out * interm_Q / 1000
            If post_denit_methanol_flux_force = True Then
                post_denit_methanol_flux = post_denit_methanol_flux_f
            Else
                post_denit_DCO_apportee = rate_DCO_apportee_denit_reel * post_denit_denit
                post_denit_methanol_flux = post_denit_DCO_apportee / rate_DCO_methanol
            End If
            boues_methanol = ratio_boues_DCO_methanol * post_denit_DCO_apportee
            interm_DCO_conso_bio = interm_DCO_conso_bio + post_denit_DCO_apportee * (1 - ratio_boues_DCO_methanol)
            interm_NO3 = interm_NO3 - post_denit_denit
            If ratio_N_assimile_methanol * boues_methanol > interm_NH4 Then
                interm_NO3 = interm_NO3 - (ratio_N_assimile_methanol * boues_methanol - interm_NH4)
                interm_NK = interm_NK - interm_NH4
                interm_NH4 = 0
            Else
                interm_NH4 = interm_NH4 - ratio_N_assimile_methanol * boues_methanol
                interm_NK = interm_NK - ratio_N_assimile_methanol * boues_methanol
            End If
            O2_dissous_flux = 0#
        End If


        'RE-OXYGENATION
        If choix_reox <> cuve_non Then
            interm_O2_exces = -O2_dissous_flux
            reox_nit = cinetique_nit_MLSS(reox_O2_dissous, boues_pdtes, N_nit_aerees, temperature) * reox_volume * pd_NOMBRE_HEURE_PAR_JOUR * reel_MES_concentration / 1000
            If reox_nit > ((1 + recirculation_taux) * Q / 1000 * reox_O2_dissous - O2_dissous_flux) / ratio_O2_nit Then    'N nitrifiable maximum
                reox_nit = ((1 + recirculation_taux) * Q / 1000 * reox_O2_dissous - O2_dissous_flux) / ratio_O2_nit
            End If
            interm_NO3 = interm_NO3 + reox_nit
            interm_NH4 = interm_NH4 - reox_nit
            interm_NK = interm_NK - reox_nit
            'Besoins en O2
            O2_dissous_flux = reox_O2_dissous * Q / 1000 * (1 + recirculation_taux)
            interm_O2_exces = interm_O2_exces + O2_dissous_flux
            O2_besoin_total = O2_besoin_total + ratio_O2_nit * reox_nit + interm_O2_exces
            air_Q_Nm3j = air_Q_Nm3j + air_Q_Nm3j_cuve(reox_O2_dissous, ratio_O2_nit * reox_nit + interm_O2_exces, temperature, reox_volume, False, reox_hauteur_bassin, reox_O2_facteur_alfa, reox_O2_rdt_transfert)
        End If


        '**************** test de convergence ******************
        'calcul des concentrations de sortie
        outlet_reel_NH4 = interm_NH4 / (Q * (1 + recirculation_taux)) * 1000
        outlet_reel_DBO = interm_DBO / (Q * (1 + recirculation_taux)) * 1000
        outlet_reel_NK = interm_NK / (Q * (1 + recirculation_taux)) * 1000
        outlet_reel_NO3 = interm_NO3 / (Q * (1 + recirculation_taux)) * 1000
        outlet_reel_Pt = interm_Pt / (Q * (1 + recirculation_taux)) * 1000
        'outlet_reel_dco = outlet_DBO + Q / 1000 * (1 + recirculation_taux) * DCO_dure_mgL
        recirculation_O2 = 1   'O2_dissous_flux / (Q * (1 + recirculation_taux)) * 1000

        'calcul de Nit totale et Nit aérée ainsi que de la production de boues
        'N_nit_total = NK - Norga_dure_soluble_reel * Q / 1000 - outlet_reel_NH4 * Q / 1000 - ratio_N_assimile * (DBO - outlet_reel_DBO * Q / 1000)
        'N_nit_aerees = N_nit_total - reox_nit - ratio_N_assimile_methanol * boues_methanol
        N_nit_total = N_nit_aerees + reox_nit
        boues_pdtes = ratio_boues_MESin * MES + ratio_boues_DBOin * DBO + ratio_boues_nit * N_nit_total + boues_methanol


        'à faire sur Nit totale et sur la Nit aérée
        epsilone_NH4 = epsilone_NH4 - outlet_reel_NH4
        epsilone_NO3 = epsilone_NO3 - outlet_reel_NO3
        epsilone_DBO = epsilone_DBO - outlet_reel_DBO
        epsilone_NH4 = epsilone_NH4 ^ 2
        epsilone_NO3 = epsilone_NO3 ^ 2
        epsilone_DBO = epsilone_DBO ^ 2
        epsilone_boues = epsilone_NH4 + epsilone_NO3 + epsilone_DBO


    Loop


    interm_Q = interm_Q * 1 / (1 + recirculation_taux)
    interm_DCO = interm_DCO * 1 / (1 + recirculation_taux)
    interm_DBO = interm_DBO * 1 / (1 + recirculation_taux)
    interm_NK = interm_NK * 1 / (1 + recirculation_taux)
    interm_NH4 = interm_NH4 * 1 / (1 + recirculation_taux)
    interm_NO3 = interm_NO3 * 1 / (1 + recirculation_taux)
    interm_Pt = interm_Pt * 1 / (1 + recirculation_taux)



    'vérification de la capacité à décanter
    If boues_concentration > facteur_boues_C_lim / boues_indice_Mohlman Then
        remarque(erreur_taux_recirc) = Feuil15.Range("HybAS_erreur_taux_recirc").Cells(1, 1)
    Else
        remarque(erreur_taux_recirc) = ""
    End If



    'SORTIE EAU ET AUTRES DONNEES
    interm_Q = Q - boues_Q

    If boues_MES_force = True Then
        boues_MES = boues_MES_f
    Else
        boues_MES = (boues_pdtes + boues_minerales_pdtes) - sortie_MES * Q / 1000
    End If
    If boues_concentration_force = True Then
        boues_concentration = boues_concentration_f
    Else
        boues_concentration = ((1 + recirculation_taux) * reel_MES_concentration - sortie_MES / 1000) / recirculation_taux
    End If

    boues_Q = boues_MES / boues_concentration

    DBO_sortie = outlet_reel_DBO    'stockée pour anaérobie
    DCO_apportee_methanol = DCO_apportee_methanol + post_denit_DCO_apportee

    interm_DCO_conso_bio = interm_DCO_conso_bio + post_denit_DCO_apportee * (1 - ratio_boues_DCO_methanol) + (DBO - interm_DBO - anaerobie_DBO_Pt * deltaP_bio)
    'DCO_conso_bio = DCO_conso_bio + interm_DCO_conso_bio
    'DCO_conso_bio = O2_besoin_total - ratio_O2_nit * N_nit_total


    'Comment gérer le forçage des besoins en O2, du débit d'air et surtout des concentrations de sortie? Bonne question. On va faire ça un peu comme on peut parce que c'est quand même plus compliqué qu'une BA
    If outlet_reel_NH4_force = True Then
        delta_O2_besoins = ratio_O2_nit * (outlet_reel_NH4 - outlet_reel_NH4_f) * Q / 1000
        outlet_reel_NH4 = outlet_reel_NH4_f
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
    If outlet_reel_NO3_force = True Then
        outlet_reel_NO3 = outlet_reel_NO3_f
    End If
    

    'ON AFFECTE DE NOUVEAU LES VALEURS DES CHARGES
    Q = Q - boues_Q
    MES = sortie_MES * Q / 1000
    DCO = (outlet_reel_DBO + DCO_dure_mgL + 1.45 * sortie_MES * reel_MV_MES) * Q / 1000    'A VOIR SI COHERENT   + 1.45 * sortie_MES * reel_MV_MES
    If outlet_reel_DBO_force = True Then
        outlet_reel_DBO = outlet_reel_DBO_f
        DBO = outlet_reel_DBO * Q / 1000
    Else
        DBO = (outlet_reel_DBO + 0.6 * sortie_MES * reel_MV_MES) * Q / 1000  '+ 0.6 * sortie_MES * reel_MV_MES
        outlet_reel_DBO = DBO / Q * 1000
    End If
    NK = outlet_reel_NK * Q / 1000
    NH4 = outlet_reel_NH4 * Q / 1000
    NO3 = outlet_reel_NO3 * Q / 1000
    Pt = outlet_reel_Pt * Q / 1000


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub calcul_consommation_electrique(ByRef electricite_II_aeration)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim I As Integer

    Dim clarif_diametre_limite As Double

    'variables pour l'aération
    Dim gamme_P As Integer
    Dim gamme_Q As Integer
    Dim surpresseur_puissance As Double
    Dim electricite_racleur As Double
    Dim electricite_agitation As Double
    Dim electricite_recirculation_MLSS As Double
    Dim electricite_recirculation As Double
    Dim electricite_extraction As Double


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
                remarque(erreur_choix_surpresseur) = Feuil15.Range("HybAS_erreur_choix_surpresseur").Cells(1, 1)
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
                remarque(erreur_choix_surpresseur) = Feuil15.Range("HybAS_erreur_choix_surpresseur").Cells(1, 1)
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


    'CLARIFICATEUR
    'electricite_racleur = rate_puissance_fixe * pd_NOMBRE_HEURE_PAR_JOUR + rate_puissance_debit * stockage_Q / 1000
    'DCA le 15/07/11 données Pascal JOUAFFRE  de 12 à 32m diamètre --> 0.55 kW installés, de 33 à 52m --> 0.75 kW
    clarif_diametre_limite = 32
    If clarif_surface < pd_NOMBRE_PI * (clarif_diametre_limite / 2) ^ 2 Then
        electricite_racleur = 0.55 * pd_NOMBRE_HEURE_PAR_JOUR
    Else
        electricite_racleur = 0.75 * pd_NOMBRE_HEURE_PAR_JOUR
    End If

    'AGITATION
    electricite_agitation = ratio_elec_anaerobie * anaerobie_volume * agitation_fct / 1000
    electricite_agitation = electricite_agitation + ratio_elec_pre_anoxie * pre_anoxie_volume * agitation_fct / 1000
    electricite_agitation = electricite_agitation + ratio_elec_pre_denit * pre_denit_volume * agitation_fct / 1000
    electricite_agitation = electricite_agitation + ratio_elec_deox * deox_volume * agitation_fct / 1000
    electricite_agitation = electricite_agitation + ratio_elec_post_denit * post_denit_volume * agitation_fct / 1000
    If choix_fct_Mox = True Then
        electricite_agitation = electricite_agitation + 3 * cuve_aeree_volume(1) * agitation_fct / 1000
    End If

    'POMPES DE RECIRCULATION ET EXTRACTION
    electricite_recirculation_MLSS = ratio_elec_recirculation_MLSS * reel_recirculation_MLSS_taux * stockage_Q_reel * recirculation_MLSS_P_refoulement
    electricite_recirculation = ratio_elec_recirculation * recirculation_taux * stockage_Q_reel * recirculation_P_refoulement
    electricite_extraction = ratio_elec_extraction * boues_Q * extraction_P_refoulement


    'CONSOMMATION GLOBALE D'ELECTRICITE
    electricite = electricite_aeration + electricite_racleur + electricite_agitation + electricite_recirculation_MLSS + electricite_recirculation + electricite_extraction
    electricite_fixe = electricite_agitation + electricite_racleur + respiration_gL(pd_TemperatureEauConditionsReelles_°C, reel_MES_concentration) * volume_total_bassins / O2_besoin_total * electricite_aeration

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 12  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil15.Range("HybAS_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Function ratio_molaire_precipitation(ByVal deltaP_precipitation, ByVal Q) As Double

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim P_sortie As Double
    Dim Ptbr As Double

    P_sortie = Pt_garantie
    Ptbr = deltaP_precipitation / Q * 1000

    ratio_molaire_precipitation = (Ptbr + P_sortie) ^ (-0.5) - (P_sortie ^ (-0.5) - 0.1)
    ratio_molaire_precipitation = ratio_molaire_precipitation / ((Ptbr + P_sortie) - (P_sortie - 0.1))
    ratio_molaire_precipitation = 1.1 - 6 * ratio_molaire_precipitation



    Exit Function
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "15.20"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Function

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







