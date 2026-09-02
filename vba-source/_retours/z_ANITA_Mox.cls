Attribute VB_Name = "z_ANITA_Mox"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Dim mcls_RetoursAdmis As CLS_RetoursAdmisTraitement

'constantes pour le type de média
Private Const nb_type_media = 3
Private Const media_K3 = 1
Private Const media_K5 = 2
Private Const media_chip_M = 3

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

'constantes pour le choix du type de régulation d'air
Private Const nb_type_regulation = 1  'nombre de types de régulation
Private Const regulation_Mox = 1

'choix de l'utilisateur
Dim choix_media_cuve_C As Integer
Dim choix_media_Mox As Integer
Dim choix_surpresseur As Integer
Dim choix_regulation As Integer

'variables à échanger avec la synthèse
Dim temperature As Double
'début modif MSM 20/11/12
Dim md_NominalQ_m3j As Double
Dim md_NominalNNH4_mgL As Double
Dim md_NominalDCOSolubleDegradable_mgL As Double
Dim md_NominalMES_mgL As Double
'fin modif MSM 20/11/12
Dim cuve_C_filling As Double
Dim cuve_C_volume As Double
Dim cuve_Mox_filling As Double
Dim cuve_Mox_volume As Double
Dim cuve_Mox_tps_sejour As Double
Dim O2_besoin As Double
Dim O2_rdt_transfert As Double
Dim O2_facteur_alfa As Double
Dim hauteur_bassin As Double
Dim O2_dissous As Double
Dim air_Q_Nm3j As Double
Dim air_P_refoulement As Double
Dim surpresseur_conso_spec As Double
Public electricite As Double

'variables à échanger avec la synthèse
Dim temperature_force As Boolean
'début modif MSM 20/11/12
Dim mn_NominalQ_m3j_force As Boolean
Dim mn_NominalNNH4_mgL_force As Boolean
Dim mn_NominalDCOSolubleDegradable_mgL_force As Boolean
Dim mn_NominalMES_mgL_force As Boolean
'fin modif MSM 20/11/12
Dim cuve_C_filling_force As Boolean
Dim cuve_C_volume_force As Boolean
Dim cuve_Mox_filling_force As Boolean
Dim cuve_Mox_volume_force As Boolean
Dim O2_besoin_force As Boolean
Dim O2_rdt_transfert_force As Boolean
Dim O2_facteur_alfa_force As Boolean
Dim hauteur_bassin_force As Boolean
Dim O2_dissous_force As Boolean
Dim air_Q_Nm3j_force As Boolean
Dim air_P_refoulement_force As Boolean
Dim surpresseur_conso_spec_force As Boolean

'variables à échanger avec la synthèse
Dim temperature_f As Double
'début modif MSM 20/11/12
Dim md_NominalQ_m3j_f As Double
Dim md_NominalNNH4_mgL_f As Double
Dim md_NominalDCOSolubleDegradable_mgL_f As Double
Dim md_NominalMES_mgL_f As Double
'fin modif MSM 20/11/12
Dim cuve_C_filling_f As Double
Dim cuve_C_volume_f As Double
Dim cuve_Mox_filling_f As Double
Dim cuve_Mox_volume_f As Double
Dim O2_besoin_f As Double
Dim O2_rdt_transfert_f As Double
Dim O2_facteur_alfa_f As Double
Dim hauteur_bassin_f As Double
Dim O2_dissous_f As Double
Dim air_Q_Nm3j_f As Double
Dim air_P_refoulement_f As Double
Dim surpresseur_conso_spec_f As Double

'hypothèses
Dim T_rdt_max_DCOsolDeg_NNH4 As Double
Dim cuve_Mox_tps_sejour_min As Double
Dim rdt_nit As Double
Dim rdt_nit_max As Double
Dim rate_NO3f_Nnit As Double
Dim assimilation_N_DCOsolubleDegradee As Double
Dim rdt_DCOsolubleDegradable_lim As Double
Dim rate_DCOsolubleDegradable_30_max As Double
Dim DCOsolubleDegradable_correction_T As Double
Dim DCOsolubleDegradable_T_ref As Double
Dim surface_spec(nb_type_media) As Double
Dim filling_max(nb_type_media) As Double
Dim rate_N_30(nb_type_media) As Double
Dim N_correction_T As Double
Dim N_T_ref As Double
Dim rate_NorgParticulaire_ammonifie As Double
'Dim a_K1(1) As double
Dim ratio_O2_DCOsoluble_degradable As Double
Dim ratio_O2_nit As Double
Dim ratio_O2_MES As Double
Dim impact_regulation(nb_type_regulation) As Double
Dim air_Qmin_agitation As Double
Dim air_tps_fonctionnement As Double
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
Dim surpresseur_conso_Wh_Nm3mCE(nb_type_surpresseur, nb_surpresseur_gamme_Q, nb_surpresseur_gamme_P) As Double
Dim respO2_correctif_T As Double
Dim respO2_T_ref As Double
Dim DCOsolubleDegradable_DCOsoluble As Double
Dim ratio_DCOsolubleDegradable_NNH4_maxi As Double
Dim ratio_DCOsolubleDegradable_NNH4_sortie_cuveC As Double

Dim MES_concentration As Double
Dim ratio_O2_Mox As Double   'part des besoins en O2 dans la cuve Mox

Dim remarque(5) As String

Sub affich_remarque()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 5  'penser à modifier la capacité du vecteur "message_erreur"


    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil30.Range("ANITA_Mox_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "30.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub calcul_consommation_electrique()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer

    Dim Patm As Double
    Dim delta_P As Double
    Dim O2_saturation_20_P As Double
    Dim O2_saturation_T_P As Double
    Dim O2_facteur_K As Double
    Dim air_Q_Nm3j_DBO As Double
    Dim air_Q_Nm3j_Mox As Double
    Dim gamme_P As Integer
    Dim gamme_Q As Integer

    Dim electricite_aeration As Double


    'AERATION
    If air_Q_Nm3j_force = True Then
        air_Q_Nm3j = air_Q_Nm3j_f
    Else
        'calcul O2 saturation à (T,P=Patm+delta_P) et (20,P0)
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

        'calcul facteur K
        O2_facteur_K = O2_facteur_alfa * (O2_facteur_beta * O2_saturation_T_P - O2_dissous) / O2_saturation_20_P
        O2_facteur_K = O2_facteur_K * correctif_T_K ^ (temperature - T_ref_K)
        'cuve DBO
        If cuve_C_volume <> 0 Then
            air_Q_Nm3j_DBO = (1 - ratio_O2_Mox) * O2_besoin / (O2_facteur_K * O2_rdt_transfert / 100 * (hauteur_bassin - hauteur_diffuseur) * ratio_kgO2_Nm3air)
            If air_Q_Nm3j_DBO / (air_tps_fonctionnement * cuve_C_volume / hauteur_bassin) < air_Qmin_agitation Then
                air_Q_Nm3j_DBO = air_Qmin_agitation * air_tps_fonctionnement * cuve_C_volume / hauteur_bassin
            End If
        End If
        'cuve Mox
        air_Q_Nm3j_Mox = ratio_O2_Mox * O2_besoin / (O2_facteur_K * O2_rdt_transfert / 100 * (hauteur_bassin - hauteur_diffuseur) * ratio_kgO2_Nm3air)
        If air_Q_Nm3j_Mox / (air_tps_fonctionnement * cuve_Mox_volume / hauteur_bassin) < air_Qmin_agitation Then
            air_Q_Nm3j_Mox = air_Qmin_agitation * air_tps_fonctionnement * cuve_Mox_volume / hauteur_bassin
        End If
        'global
        air_Q_Nm3j = air_Q_Nm3j_DBO + air_Q_Nm3j_Mox
    End If
    'calcul de la consommation électrique de l'aération
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
                remarque(5) = Feuil30.Range("ANITA_Mox_erreur_choix_surpresseur").Cells(1, 1)
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
                remarque(5) = Feuil30.Range("ANITA_Mox_erreur_choix_surpresseur").Cells(1, 1)
            End If
        End If
        'gamme de débit   (pour le nominal)
        gamme_Q = surpresseur_gamme_Q_0_infini
        surpresseur_conso_spec = surpresseur_conso_Wh_Nm3mCE(choix_surpresseur, gamme_Q, gamme_P) * (1 + impact_regulation(choix_regulation))
    End If
    electricite_aeration = air_Q_Nm3j * air_P_refoulement * surpresseur_conso_spec / 1000

    electricite = electricite_aeration

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "30.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub dimensionnement(ByRef retour_flux() As Double, ByRef retour_flux_soluble() As Double, ByRef retour_digestion() As Double, ByRef retour_digestion_soluble() As Double, ByRef retour_athos() As Double, ByRef retour_athos_soluble() As Double)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    Const DBO_iteration = 10

    Dim Q As Double
    Dim DCO As Double
    Dim DCOsoluble_degradable As Double
    Dim DBO As Double
    Dim MES As Double
    Dim NK As Double
    Dim NH4 As Double
    Dim NO3 As Double
    Dim Pt As Double

    Dim FacteurDebit As Double

    Dim DCOsolubleDegradee As Double
    Dim rate_DCOsolubleDegradable_T As Double
    Dim N_assimile As Double
    Dim volume_media As Double
    Dim NNH4_abattu As Double
    Dim NNH4_degradable_max As Double
    Dim rate_NNH4abattu_T As Double

    Dim ratio_DCOsolubleDegradable_NNH4 As Double

    Dim retour_flux_non_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_soluble_non_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_soluble_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_admis(retour_caracteristique_nb) As Double

    'On regarde les retours admis
    Call mcls_RetoursAdmis.dimensionnement(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble, retour_flux_non_admis, retour_flux_soluble_non_admis, retour_flux_admis, retour_flux_soluble_admis)

    'on récupère les valeurs des retours

    'MSM 21/11/12 bloc à modifier
    'ces valeurs sont des charges en kg/j
    '        Q = mcls_RetoursAdmis.pd_Q_m3j
    '        DCO = mcls_RetoursAdmis.pd_DCO_mgL * Q / 1000
    '        DBO = mcls_RetoursAdmis.pd_DBO_mgL * Q / 1000
    '        MES = mcls_RetoursAdmis.pd_MES_mgL * Q / 1000
    '        NK = mcls_RetoursAdmis.pd_NK_mgL * Q / 1000
    '        NH4 = mcls_RetoursAdmis.pd_NNH4_mgL * Q / 1000
    '        NO3 = mcls_RetoursAdmis.pd_NNO3_mgL * Q / 1000
    '        Pt = mcls_RetoursAdmis.pd_Pt_mgL * Q / 1000
    '
    '        DCOsoluble_degradable = DCOsolubleDegradable_DCOsoluble * retour_flux_soluble_admis(repere_ret_DCO)
    '        If DCOsoluble_degradable > DCO Then
    '            DCOsoluble_degradable = DCO
    '        End If
    'MSM 21/11/12 bloc à modifier

    'début modif MSM 22/11/12
    '    md_NominalQ_m3j
    '    md_NominalNNH4_mgL
    '    md_NominalDCOSolubleDegradable_mgL
    '    md_NominalMES_mgL

    If mn_NominalQ_m3j_force = True Then
        md_NominalQ_m3j = md_NominalQ_m3j_f
    Else
        md_NominalQ_m3j = mcls_RetoursAdmis.pd_Q_m3j
    End If
    Q = md_NominalQ_m3j
    FacteurDebit = mcls_RetoursAdmis.pd_Q_m3j / md_NominalQ_m3j
    DCO = mcls_RetoursAdmis.pd_DCO_mgL * FacteurDebit * Q / 1000
    DBO = mcls_RetoursAdmis.pd_DBO_mgL * FacteurDebit * Q / 1000

    If mn_NominalMES_mgL_force = True Then
        md_NominalMES_mgL = md_NominalMES_mgL_f
        MES = md_NominalMES_mgL * Q / 1000
    Else
        MES = mcls_RetoursAdmis.pd_MES_mgL * FacteurDebit * Q / 1000
        md_NominalMES_mgL = MES / Q * 1000
    End If

    If mn_NominalNNH4_mgL_force = True Then
        md_NominalNNH4_mgL = md_NominalNNH4_mgL_f
        NH4 = md_NominalNNH4_mgL * Q / 1000
    Else
        NH4 = mcls_RetoursAdmis.pd_NNH4_mgL * FacteurDebit * Q / 1000
        md_NominalNNH4_mgL = NH4 / Q * 1000
    End If
    NK = (mcls_RetoursAdmis.pd_NK_mgL - mcls_RetoursAdmis.pd_NNH4_mgL) * FacteurDebit * Q / 1000 + NH4
    NO3 = mcls_RetoursAdmis.pd_NNO3_mgL * Q / 1000
    Pt = mcls_RetoursAdmis.pd_Pt_mgL * Q / 1000

    If mn_NominalDCOSolubleDegradable_mgL_force = True Then
        md_NominalDCOSolubleDegradable_mgL = md_NominalDCOSolubleDegradable_mgL_f
        DCOsoluble_degradable = md_NominalDCOSolubleDegradable_mgL * Q / 1000
    Else
        DCOsoluble_degradable = DCOsolubleDegradable_DCOsoluble * retour_flux_soluble_admis(repere_ret_DCO)
        md_NominalDCOSolubleDegradable_mgL = DCOsoluble_degradable / Q * 1000
    End If

    If DCOsoluble_degradable > DCO Then
        DCO = DCOsoluble_degradable
    End If
    'fin modif MSM 22/11/12

    ratio_DCOsolubleDegradable_NNH4 = DCOsoluble_degradable / NH4

    'dimensionnement cuve DBO
    If ratio_DCOsolubleDegradable_NNH4 > ratio_DCOsolubleDegradable_NNH4_maxi Then
        If cuve_C_volume_force = True Then
            cuve_C_volume = cuve_C_volume_f
            If cuve_C_filling_force = False Then
                DCOsolubleDegradee = DCOsoluble_degradable - ratio_DCOsolubleDegradable_NNH4_sortie_cuveC * NH4
                If DCOsolubleDegradee > 0 Then
                    If temperature > T_rdt_max_DCOsolDeg_NNH4 Then
                        rate_DCOsolubleDegradable_T = rate_DCOsolubleDegradable_30_max
                    Else
                        rate_DCOsolubleDegradable_T = rate_DCOsolubleDegradable_30_max * DCOsolubleDegradable_correction_T ^ (temperature - DCOsolubleDegradable_T_ref)
                    End If
                    volume_media = DCOsolubleDegradee * 1000 / (rate_DCOsolubleDegradable_T * surface_spec(choix_media_cuve_C))
                    cuve_C_filling = volume_media / cuve_C_volume
                Else
                    DCOsolubleDegradee = 0
                    cuve_C_filling = 0
                End If
            Else
                If temperature > T_rdt_max_DCOsolDeg_NNH4 Then
                    rate_DCOsolubleDegradable_T = rate_DCOsolubleDegradable_30_max
                Else
                    rate_DCOsolubleDegradable_T = rate_DCOsolubleDegradable_30_max * DCOsolubleDegradable_correction_T ^ (temperature - DCOsolubleDegradable_T_ref)
                End If
                DCOsolubleDegradee = rate_DCOsolubleDegradable_T * surface_spec(choix_media_cuve_C) * volume_media / 1000
                If DCOsolubleDegradee / DCOsoluble_degradable > rdt_DCOsolubleDegradable_lim Then
                    DCOsolubleDegradee = DCOsoluble_degradable * rdt_DCOsolubleDegradable_lim
                End If
            End If
        Else
            DCOsolubleDegradee = DCOsoluble_degradable - ratio_DCOsolubleDegradable_NNH4_sortie_cuveC * NH4
            If DCOsolubleDegradee > 0 Then
                If temperature > T_rdt_max_DCOsolDeg_NNH4 Then
                    rate_DCOsolubleDegradable_T = rate_DCOsolubleDegradable_30_max
                Else
                    rate_DCOsolubleDegradable_T = rate_DCOsolubleDegradable_30_max * DCOsolubleDegradable_correction_T ^ (temperature - DCOsolubleDegradable_T_ref)
                End If
                volume_media = DCOsolubleDegradee * 1000 / (rate_DCOsolubleDegradable_T * surface_spec(choix_media_cuve_C))
                cuve_C_volume = volume_media / cuve_C_filling
            Else
                DCOsolubleDegradee = 0
                cuve_C_volume = 0
            End If
        End If
    Else
        cuve_C_volume = 0
        cuve_C_filling = 0
        DCOsolubleDegradee = 0
    End If
    N_assimile = assimilation_N_DCOsolubleDegradee * DCOsolubleDegradee

    'PENSER A RAJOUTER PRODUCTION DE BOUES SI CUVE DBO


    'dimensionnement cuve Amox
    If cuve_Mox_volume_force = True Then
        cuve_Mox_volume = cuve_Mox_volume_f
        If cuve_Mox_filling_force = False Then
            NNH4_abattu = rdt_nit * (NH4 + rate_NorgParticulaire_ammonifie * (NK - NH4) - N_assimile)
            If temperature > T_rdt_max_DCOsolDeg_NNH4 Then
                rate_NNH4abattu_T = rate_N_30(choix_media_Mox)
            Else
                rate_NNH4abattu_T = rate_N_30(choix_media_Mox) * N_correction_T ^ (temperature - N_T_ref)
            End If
            volume_media = NNH4_abattu * 1000 / (rate_NNH4abattu_T * surface_spec(choix_media_Mox))
            cuve_Mox_filling = volume_media / cuve_Mox_volume
            If cuve_Mox_filling > filling_max(choix_media_Mox) Then
                cuve_Mox_filling = filling_max(choix_media_Mox)
                NNH4_abattu = (rate_NNH4abattu_T * cuve_Mox_volume * cuve_Mox_filling * surface_spec(choix_media_Mox)) / 1000
            End If
        Else
            If temperature > T_rdt_max_DCOsolDeg_NNH4 Then
                rate_NNH4abattu_T = rate_N_30(choix_media_Mox)
            Else
                rate_NNH4abattu_T = rate_N_30(choix_media_Mox) * N_correction_T ^ (temperature - N_T_ref)
            End If
            NNH4_abattu = (rate_NNH4abattu_T * cuve_Mox_volume * cuve_Mox_filling * surface_spec(choix_media_Mox)) / 1000
            NNH4_degradable_max = rdt_nit_max * (NH4 + rate_NorgParticulaire_ammonifie * (NK - NH4) - N_assimile)
            If NNH4_abattu > NNH4_degradable_max Then
                NNH4_abattu = NNH4_degradable_max
            End If
        End If
        If cuve_Mox_volume / (Q / pd_NOMBRE_HEURE_PAR_JOUR) < cuve_Mox_tps_sejour_min Then
            remarque(2) = Feuil30.Range("ANITA_Mox_erreur_tps_sejour").Cells(1, 1)
        End If
    Else
        NNH4_abattu = rdt_nit * (NH4 + rate_NorgParticulaire_ammonifie * (NK - NH4) - N_assimile)
        If temperature > T_rdt_max_DCOsolDeg_NNH4 Then
            rate_NNH4abattu_T = rate_N_30(choix_media_Mox)
        Else
            rate_NNH4abattu_T = rate_N_30(choix_media_Mox) * N_correction_T ^ (temperature - N_T_ref)
        End If
        volume_media = NNH4_abattu * 1000 / (rate_NNH4abattu_T * surface_spec(choix_media_Mox))
        cuve_Mox_volume = volume_media / cuve_Mox_filling
        If cuve_Mox_volume / (Q / pd_NOMBRE_HEURE_PAR_JOUR) < cuve_Mox_tps_sejour_min Then
            cuve_Mox_volume = Q * cuve_Mox_tps_sejour_min / pd_NOMBRE_HEURE_PAR_JOUR
            If cuve_Mox_filling_force = False Then
                cuve_Mox_filling = volume_media / cuve_Mox_volume
            Else
                NNH4_abattu = (rate_NNH4abattu_T * cuve_Mox_volume * cuve_Mox_filling * surface_spec(choix_media_Mox)) / 1000
                NNH4_degradable_max = rdt_nit_max * (NH4 + rate_NorgParticulaire_ammonifie * (NK - NH4) - N_assimile)
                If NNH4_abattu > NNH4_degradable_max Then
                    NNH4_abattu = NNH4_degradable_max
                End If
            End If
        End If
    End If

    'calcul des besoins en O2
    MES_concentration = MES / Q
    O2_besoin = ratio_O2_DCOsoluble_degradable * DCOsolubleDegradee + ratio_O2_nit * NNH4_abattu + ratio_O2_MES * respO2_correctif_T ^ (temperature - respO2_T_ref) * MES_concentration * (cuve_C_volume + cuve_Mox_volume)


    'calcul des charges en sortie
    Q = Q
    DCO = DCO - DCOsolubleDegradee
    If DBO > DCOsolubleDegradee Then
        DBO = DBO - DCOsolubleDegradee
    Else
        DBO = 0
    End If
    MES = MES
    NH4 = NH4 + rate_NorgParticulaire_ammonifie * (NK - NH4) - NNH4_abattu - N_assimile
    NK = NK - NNH4_abattu - N_assimile
    NO3 = rate_NO3f_Nnit * NNH4_abattu
    Pt = Pt


    'on récupère les valeurs des retours admis
    retour_flux_admis(repere_ret_Q) = Q
    retour_flux_admis(repere_ret_DCO) = DCO
    retour_flux_admis(repere_ret_DBO) = DBO
    retour_flux_admis(repere_ret_MES) = MES
    retour_flux_admis(repere_ret_NK) = NK
    retour_flux_admis(repere_ret_NH4) = NH4
    retour_flux_admis(repere_ret_NO3) = NO3
    retour_flux_admis(repere_ret_Pt) = Pt
    retour_flux_soluble_admis(repere_ret_NK) = retour_flux_soluble_admis(repere_ret_NK) - retour_flux_soluble_admis(repere_ret_NH4) + NH4
    retour_flux_soluble_admis(repere_ret_NH4) = NH4
    retour_flux_soluble_admis(repere_ret_DCO) = retour_flux_soluble_admis(repere_ret_DCO) - DCOsolubleDegradee
    If retour_flux_soluble_admis(repere_ret_DBO) > DCOsolubleDegradee Then
        retour_flux_soluble_admis(repere_ret_DBO) = retour_flux_soluble_admis(repere_ret_DBO) - DCOsolubleDegradee
    Else
        retour_flux_soluble_admis(repere_ret_DBO) = 0
    End If
    retour_flux_soluble_admis(repere_ret_NO3) = NO3
    Call mcls_RetoursAdmis.flux_sortie_process(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble, retour_flux_non_admis, retour_flux_soluble_non_admis, retour_flux_admis, retour_flux_soluble_admis)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "30.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub ecriture_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil30.Range("ANITA_Mox_temperature").Cells(1, Colonne) = temperature
    'début modif MSM 20/11/12
    Feuil30.Range("ANITA_Mox_NominalQ_m3j").Cells(1, Colonne) = md_NominalQ_m3j
    Feuil30.Range("ANITA_Mox_NominalNNH4_mgL").Cells(1, Colonne) = md_NominalNNH4_mgL
    Feuil30.Range("ANITA_Mox_NominalDCOSolubleDegradable_mgL").Cells(1, Colonne) = md_NominalDCOSolubleDegradable_mgL
    Feuil30.Range("ANITA_Mox_NominalMES_mgL").Cells(1, Colonne) = md_NominalMES_mgL
    'fin modif MSM 20/11/12
    Feuil30.Range("ANITA_Mox_cuve_DBO_filling").Cells(1, Colonne) = cuve_C_filling
    Feuil30.Range("ANITA_Mox_cuve_DBO_volume").Cells(1, Colonne) = cuve_C_volume
    Feuil30.Range("ANITA_Mox_cuve_Mox_filling").Cells(1, Colonne) = cuve_Mox_filling
    Feuil30.Range("ANITA_Mox_cuve_Mox_volume").Cells(1, Colonne) = cuve_Mox_volume
    Feuil30.Range("ANITA_Mox_cuve_Mox_tps_sejour").Cells(1, Colonne) = cuve_Mox_tps_sejour
    Feuil30.Range("ANITA_Mox_O2_besoin").Cells(1, Colonne) = O2_besoin
    Feuil30.Range("ANITA_Mox_O2_facteur_alpha").Cells(1, Colonne) = O2_facteur_alfa
    Feuil30.Range("ANITA_Mox_O2_rdt_transfert").Cells(1, Colonne) = O2_rdt_transfert
    Feuil30.Range("ANITA_Mox_hauteur_bassin").Cells(1, Colonne) = hauteur_bassin
    Feuil30.Range("ANITA_Mox_O2_dissous").Cells(1, Colonne) = O2_dissous
    Feuil30.Range("ANITA_Mox_air_Q_Nm3j").Cells(1, Colonne) = air_Q_Nm3j
    Feuil30.Range("ANITA_Mox_air_P_refoulement").Cells(1, Colonne) = air_P_refoulement
    Feuil30.Range("ANITA_Mox_surpresseur_conso_spec").Cells(1, Colonne) = surpresseur_conso_spec
    Feuil30.Range("ANITA_Mox_electricite").Cells(1, Colonne) = electricite

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "30.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub effacer_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


    Feuil30.Range("ANITA_Mox_temperature").Cells(1, Colonne) = ""     'temperature
    'début modif MSM 20/11/12
    Feuil30.Range("ANITA_Mox_NominalQ_m3j").Cells(1, Colonne).ClearContents
    Feuil30.Range("ANITA_Mox_NominalNNH4_mgL").Cells(1, Colonne).ClearContents
    Feuil30.Range("ANITA_Mox_NominalDCOSolubleDegradable_mgL").Cells(1, Colonne).ClearContents
    Feuil30.Range("ANITA_Mox_NominalMES_mgL").Cells(1, Colonne).ClearContents
    'fin modif MSM 20/11/12
    Feuil30.Range("ANITA_Mox_cuve_DBO_filling").Cells(1, Colonne) = ""     'cuve_C_filling
    Feuil30.Range("ANITA_Mox_cuve_DBO_volume").Cells(1, Colonne) = ""     'cuve_C_volume
    Feuil30.Range("ANITA_Mox_cuve_Mox_filling").Cells(1, Colonne) = ""     'cuve_Mox_filling
    Feuil30.Range("ANITA_Mox_cuve_Mox_volume").Cells(1, Colonne) = ""     'cuve_Mox_volume
    Feuil30.Range("ANITA_Mox_cuve_Mox_tps_sejour").Cells(1, Colonne) = ""     'cuve_Mox_tps_sejour
    Feuil30.Range("ANITA_Mox_O2_besoin").Cells(1, Colonne) = ""     'O2_besoin
    Feuil30.Range("ANITA_Mox_O2_rdt_transfert").Cells(1, Colonne) = ""     'O2_rdt_transfert
    Feuil30.Range("ANITA_Mox_O2_facteur_alpha").Cells(1, Colonne) = ""     'O2_facteur_alfa
    Feuil30.Range("ANITA_Mox_hauteur_bassin").Cells(1, Colonne) = ""     'hauteur_bassin
    Feuil30.Range("ANITA_Mox_O2_dissous").Cells(1, Colonne) = ""     'O2_dissous
    Feuil30.Range("ANITA_Mox_air_Q_Nm3j").Cells(1, Colonne) = ""     'air_Q_Nm3j
    Feuil30.Range("ANITA_Mox_air_P_refoulement").Cells(1, Colonne) = ""     'air_P_refoulement
    Feuil30.Range("ANITA_Mox_surpresseur_conso_spec").Cells(1, Colonne) = ""     'surpresseur_conso_spec
    Feuil30.Range("ANITA_Mox_electricite").Cells(1, Colonne) = ""     'electricite



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "30.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub fonctionnement_reel(ByRef retour_flux() As Double, ByRef retour_flux_soluble() As Double, ByRef retour_digestion() As Double, ByRef retour_digestion_soluble() As Double, ByRef retour_athos() As Double, ByRef retour_athos_soluble() As Double)
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim I As Integer
    Const DBO_iteration = 10

    Dim Q As Double
    Dim DCO As Double
    Dim DCOsoluble_degradable As Double
    Dim DBO As Double
    Dim MES As Double
    Dim NK As Double
    Dim NH4 As Double
    Dim NO3 As Double
    Dim Pt As Double

    Dim DCOsolubleDegradee As Double
    Dim rate_DCOsolubleDegradable_T As Double
    Dim N_assimile As Double
    Dim volume_media As Double
    Dim NNH4_abattu As Double
    Dim NNH4_degradable_max As Double
    Dim rate_NNH4abattu_T As Double

    Dim ratio_DCOsolubleDegradable_NNH4 As Double

    Dim retour_flux_non_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_soluble_non_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_soluble_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_admis(retour_caracteristique_nb) As Double

    'On regarde les retours admis
    Call mcls_RetoursAdmis.fonctionnement_reel(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble, retour_flux_non_admis, retour_flux_soluble_non_admis, retour_flux_admis, retour_flux_soluble_admis)

    'on récupère les valeurs des retours
    Q = mcls_RetoursAdmis.pd_Q_m3j
    DCO = mcls_RetoursAdmis.pd_DCO_mgL * Q / 1000
    DBO = mcls_RetoursAdmis.pd_DBO_mgL * Q / 1000
    MES = mcls_RetoursAdmis.pd_MES_mgL * Q / 1000
    NK = mcls_RetoursAdmis.pd_NK_mgL * Q / 1000
    NH4 = mcls_RetoursAdmis.pd_NNH4_mgL * Q / 1000
    NO3 = mcls_RetoursAdmis.pd_NNO3_mgL * Q / 1000
    Pt = mcls_RetoursAdmis.pd_Pt_mgL * Q / 1000

    DCOsoluble_degradable = DCOsolubleDegradable_DCOsoluble * retour_flux_soluble_admis(repere_ret_DCO)
    If DCOsoluble_degradable > DCO Then
        DCOsoluble_degradable = DCO
    End If
    ratio_DCOsolubleDegradable_NNH4 = DCOsoluble_degradable / NH4

    'charges traitées dans la cuve DBO
    If cuve_C_volume <> 0 Then
        If temperature > T_rdt_max_DCOsolDeg_NNH4 Then
            rate_DCOsolubleDegradable_T = rate_DCOsolubleDegradable_30_max
        Else
            rate_DCOsolubleDegradable_T = rate_DCOsolubleDegradable_30_max * DCOsolubleDegradable_correction_T ^ (temperature - DCOsolubleDegradable_T_ref)
        End If
        DCOsolubleDegradee = rate_DCOsolubleDegradable_T * surface_spec(choix_media_cuve_C) * volume_media / 1000
        If DCOsolubleDegradee / DCOsoluble_degradable > rdt_DCOsolubleDegradable_lim Then
            DCOsolubleDegradee = DCOsoluble_degradable * rdt_DCOsolubleDegradable_lim
        End If
    Else
        DCOsolubleDegradee = 0
    End If
    N_assimile = assimilation_N_DCOsolubleDegradee * DCOsolubleDegradee

    'PENSER A RAJOUTER PRODUCTION DE BOUES SI CUVE DBO

    'temps de séjour dans la cuve Mox
    cuve_Mox_tps_sejour = cuve_Mox_volume / (Q / pd_NOMBRE_HEURE_PAR_JOUR)
    If cuve_Mox_tps_sejour < cuve_Mox_tps_sejour_min Then
        remarque(2) = Feuil30.Range("ANITA_Mox_erreur_tps_sejour").Cells(1, 1)
    End If

    'charges traitées dans la cuve Mox
    If temperature > T_rdt_max_DCOsolDeg_NNH4 Then
        rate_NNH4abattu_T = rate_N_30(choix_media_Mox)
    Else
        rate_NNH4abattu_T = rate_N_30(choix_media_Mox) * N_correction_T ^ (temperature - N_T_ref)
    End If
    NNH4_abattu = (rate_NNH4abattu_T * cuve_Mox_volume * cuve_Mox_filling * surface_spec(choix_media_Mox)) / 1000
    NNH4_degradable_max = rdt_nit_max * (NH4 + rate_NorgParticulaire_ammonifie * (NK - NH4) - N_assimile)
    If NNH4_abattu > NNH4_degradable_max Then
        NNH4_abattu = NNH4_degradable_max
    End If


    'calcul des besoins en O2
    MES_concentration = MES / Q
    O2_besoin = ratio_O2_DCOsoluble_degradable * DCOsolubleDegradee + ratio_O2_nit * NNH4_abattu + ratio_O2_MES * respO2_correctif_T ^ (temperature - respO2_T_ref) * MES_concentration * (cuve_C_volume + cuve_Mox_volume)
    ratio_O2_Mox = (ratio_O2_nit * NNH4_abattu + ratio_O2_MES * respO2_correctif_T ^ (temperature - respO2_T_ref) * MES_concentration * cuve_Mox_volume) / O2_besoin
    If O2_besoin_force = True Then
        O2_besoin = O2_besoin_f
    End If

    'calcul des charges en sortie
    Q = Q
    DCO = DCO - DCOsolubleDegradee
    If DBO > DCOsolubleDegradee Then
        DBO = DBO - DCOsolubleDegradee
    Else
        DBO = 0
    End If
    MES = MES
    NH4 = NH4 + rate_NorgParticulaire_ammonifie * (NK - NH4) - NNH4_abattu - N_assimile
    NK = NK - NNH4_abattu - N_assimile
    NO3 = rate_NO3f_Nnit * NNH4_abattu
    Pt = Pt



    'on récupère les valeurs des retours admis
    retour_flux_admis(repere_ret_Q) = Q
    retour_flux_admis(repere_ret_DCO) = DCO
    retour_flux_admis(repere_ret_DBO) = DBO
    retour_flux_admis(repere_ret_MES) = MES
    retour_flux_admis(repere_ret_NK) = NK
    retour_flux_admis(repere_ret_NH4) = NH4
    retour_flux_admis(repere_ret_NO3) = NO3
    retour_flux_admis(repere_ret_Pt) = Pt
    retour_flux_soluble_admis(repere_ret_NK) = (1 - rate_NorgParticulaire_ammonifie) * (retour_flux_soluble_admis(repere_ret_NK) - retour_flux_soluble_admis(repere_ret_NH4)) + NH4
    retour_flux_soluble_admis(repere_ret_NH4) = NH4
    retour_flux_soluble_admis(repere_ret_DCO) = retour_flux_soluble_admis(repere_ret_DCO) - DCOsolubleDegradee
    If retour_flux_soluble_admis(repere_ret_DBO) > DCOsolubleDegradee Then
        retour_flux_soluble_admis(repere_ret_DBO) = retour_flux_soluble_admis(repere_ret_DBO) - DCOsolubleDegradee
    Else
        retour_flux_soluble_admis(repere_ret_DBO) = 0
    End If
    retour_flux_soluble_admis(repere_ret_NO3) = NO3
    Call mcls_RetoursAdmis.flux_sortie_process(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble, retour_flux_non_admis, retour_flux_soluble_non_admis, retour_flux_admis, retour_flux_soluble_admis)



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "30.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub attribution_valeur_par_defaut()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    'température traitement
    If temperature_force = True Then
        temperature = temperature_f
        If temperature < 20 Then
            remarque(1) = Feuil30.Range("ANITA_Mox_erreur_temperature").Cells(1, 1)
        End If
    Else
        temperature = 25    '°C
    End If

    'Taux de remplissage média cuve DBO
    If cuve_C_filling_force = True Then
        cuve_C_filling = cuve_C_filling_f
        If cuve_Mox_filling > filling_max(choix_media_Mox) Then
            remarque(3) = Feuil30.Range("ANITA_Mox_erreur_media_DBO").Cells(1, 1)
        End If
    Else
        cuve_C_filling = 0.5    'quelque soit le type de média
    End If

    'Taux de remplissage média cuve Mox
    If cuve_Mox_filling_force = True Then
        cuve_Mox_filling = cuve_Mox_filling_f
        If cuve_Mox_filling > filling_max(choix_media_Mox) Then
            remarque(4) = Feuil30.Range("ANITA_Mox_erreur_media_Mox").Cells(1, 1)
        End If
    Else
        If choix_media_Mox = media_chip_M Then
            cuve_Mox_filling = 0.4
        Else
            cuve_Mox_filling = 0.5
        End If
    End If

    'rendement de dissolution de l'O2 en eau claire  (%/m)
    If O2_rdt_transfert_force = True Then
        O2_rdt_transfert = O2_rdt_transfert_f
    Else
        O2_rdt_transfert = 4
    End If

    'facteur alpha
    If O2_facteur_alfa_force = True Then
        O2_facteur_alfa = O2_facteur_alfa_f
    Else
        O2_facteur_alfa = 0.7
    End If

    'Hauteur moyenne des bassins  (m)
    If hauteur_bassin_force = True Then
        hauteur_bassin = hauteur_bassin_f
    Else
        hauteur_bassin = 6
    End If

    'O2 dissous  (mg/L)
    If O2_dissous_force = True Then
        O2_dissous = O2_dissous_f
    Else
        O2_dissous = 2
    End If

    'Pression de refoulement des surpresseurs (mCE)
    If air_P_refoulement_force = True Then
        air_P_refoulement = air_P_refoulement_f
    Else
        air_P_refoulement = 2 + hauteur_bassin
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "30.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub


Sub hypotheses()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer



    'température minimale pour utiliser le Mox   °C
    T_rdt_max_DCOsolDeg_NNH4 = 30

    'ratio pour savoir si l'on traite le C ou non
    ratio_DCOsolubleDegradable_NNH4_maxi = 0.4
    ratio_DCOsolubleDegradable_NNH4_sortie_cuveC = 0.2

    'temps de séjour minimum dans la cuve Mox   en heures
    cuve_Mox_tps_sejour_min = 12   'h    'Vérifié avec Romain LEMAIRE le 09/11/2012

    'rendement de nitrification
    rdt_nit = 0.9    'Vérifié avec Romain LEMAIRE le 09/11/2012
    rdt_nit_max = 0.9    'Vérifié avec Romain LEMAIRE le 09/11/2012

    'pourcentage de N nitrifié qui donne du NO3
    rate_NO3f_Nnit = 0.11    'Vérifié avec Romain LEMAIRE le 09/11/2012

    'ratio de N assimilé
    assimilation_N_DCOsolubleDegradee = 0.04    'Vérifié avec Romain LEMAIRE le 09/11/2012

    'rendement maxi sur la DCO soluble dégradable
    rdt_DCOsolubleDegradable_lim = 0.9    '%    'Vérifié avec Romain LEMAIRE le 09/11/2012

    'Part de la DCO soluble dégradable     'Vérifié avec Romain LEMAIRE le 09/11/2012
    If choix_dig_simple = True Then
        DCOsolubleDegradable_DCOsoluble = 0.5
    Else
        DCOsolubleDegradable_DCOsoluble = 0.3
    End If

    'rate DBO 10°C calculé  =rate_max*exp(rate0+(rdtDBO-rate1)^rate2)
    rate_DCOsolubleDegradable_30_max = 20   '20 si T>30 sinon    'Vérifié avec Romain LEMAIRE le 09/11/2012
    'correction température
    DCOsolubleDegradable_correction_T = 1.06    'Vérifié avec Romain LEMAIRE le 09/11/2012
    DCOsolubleDegradable_T_ref = 30    '10°C    'Vérifié avec Romain LEMAIRE le 09/11/2012

    'caractéristiques média    'Vérifié avec Romain LEMAIRE le 09/11/2012
    'surface spécifique (m2/m3)
    surface_spec(media_K3) = 500
    surface_spec(media_K5) = 800
    surface_spec(media_chip_M) = 1200
    'remplissage maximal (%)
    filling_max(media_K3) = 0.65
    filling_max(media_K5) = 0.55
    filling_max(media_chip_M) = 0.55
    'rate N à 30°C   (g/(m2.j)    'Vérifié avec Romain LEMAIRE le 09/11/2012
    If choix_dig_simple = True Then
        rate_N_30(media_K3) = 2.6
        rate_N_30(media_K5) = 2.6
        rate_N_30(media_chip_M) = 2.25
    Else
        rate_N_30(media_K3) = 1.35
        rate_N_30(media_K5) = 1.35
        rate_N_30(media_chip_M) = 1.12
    End If

    'correction cinétique de la nit    'Vérifié avec Romain LEMAIRE le 09/11/2012
    N_correction_T = 1.07
    N_T_ref = 30  '30°C

    'ratio de N organique ammonifié    'Vérifié avec Romain LEMAIRE le 09/11/2012
    rate_NorgParticulaire_ammonifie = 0.5

    'calcul des productions de boues
    'coefficient K1    a0+a1*(MESin/DBOin)
    'a_K1(0) = 1
    'a_K1(1) = 0.6
    'coefficient K2

    'consommation spécifique en O2
    'DBO (kg O2/kg DBO éliminée)
    ratio_O2_DCOsoluble_degradable = 1.05   'Vérifié avec Romain LEMAIRE le 09/11/2012  + 5% pour l'endogène calculée à la louche
    'nitrification (kg O2/kg N nitrifié)
    ratio_O2_nit = 1.94   'Vérifié avec Romain LEMAIRE le 09/11/2012  à voir pour modifier à 1.85
    'respiration (kg O2/kg MVS)
    ratio_O2_MES = 0   ' 24 * 0.002    A revoir
    'correctifs pour la conso O2 pour respiration
    respO2_correctif_T = 1.07
    respO2_T_ref = 10

    'impact du type de régulation
    impact_regulation(regulation_Mox) = 0

    'débit d'air minimal pour agitation  (Nm3/(m2.h))
    air_Qmin_agitation = 4

    'temps de fonctionnement de l'aération
    air_tps_fonctionnement = 24    'h/j

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

    'correction température pour facteur K
    correctif_T_K = 1.024
    T_ref_K = 20

    'O2 présent dans l'air (kgO2/Nm3air)
    ratio_kgO2_Nm3air = 0.3

    'facteur beta
    O2_facteur_beta = 0.95

    'consommations spécifiques par type de surpresseurs      en Wh/(Nm3.mCE)
    For I = 1 To nb_surpresseur_gamme_Q
        For j = 1 To nb_surpresseur_gamme_P
            surpresseur_conso_Wh_Nm3mCE(surpresseur_roots, I, j) = 4.5
            surpresseur_conso_Wh_Nm3mCE(surpresseur_vis, I, j) = 3.8
            surpresseur_conso_Wh_Nm3mCE(surpresseur_turbo, I, j) = 3.5
        Next j
    Next I


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "30.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub lecture_choix()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    'lecture des choix internes
    choix_media_cuve_C = Feuil30.Range("ANITA_Mox_choix_media_DBO").Cells(1, pi_FiliereConsideree)
    choix_media_Mox = Feuil30.Range("ANITA_Mox_choix_media_Mox").Cells(1, pi_FiliereConsideree)
    choix_surpresseur = Feuil30.Range("ANITA_Mox_choix_surpresseur").Cells(1, pi_FiliereConsideree)
    choix_regulation = Feuil30.Range("ANITA_Mox_choix_regulation").Cells(1, pi_FiliereConsideree)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "30.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub


Sub lecture_valeurs_forcees()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    temperature_force = Feuil30.Range("ANITA_Mox_temperature").Cells(1, pi_FiliereConsideree)
    'début modif MSM 20/11/12
    mn_NominalQ_m3j_force = Feuil30.Range("ANITA_Mox_NominalQ_m3j").Cells(1, pi_FiliereConsideree)
    mn_NominalNNH4_mgL_force = Feuil30.Range("ANITA_Mox_NominalNNH4_mgL").Cells(1, pi_FiliereConsideree)
    mn_NominalDCOSolubleDegradable_mgL_force = Feuil30.Range("ANITA_Mox_NominalDCOSolubleDegradable_mgL").Cells(1, pi_FiliereConsideree)
    mn_NominalMES_mgL_force = Feuil30.Range("ANITA_Mox_NominalMES_mgL").Cells(1, pi_FiliereConsideree)
    'fin modif MSM 20/11/12
    cuve_C_filling_force = Feuil30.Range("ANITA_Mox_cuve_DBO_filling").Cells(1, pi_FiliereConsideree)
    cuve_C_volume_force = Feuil30.Range("ANITA_Mox_cuve_DBO_volume").Cells(1, pi_FiliereConsideree)
    cuve_Mox_filling_force = Feuil30.Range("ANITA_Mox_cuve_Mox_filling").Cells(1, pi_FiliereConsideree)
    cuve_Mox_volume_force = Feuil30.Range("ANITA_Mox_cuve_Mox_volume").Cells(1, pi_FiliereConsideree)
    O2_besoin_force = Feuil30.Range("ANITA_Mox_O2_besoin").Cells(1, pi_FiliereConsideree)
    O2_facteur_alfa_force = Feuil30.Range("ANITA_Mox_O2_facteur_alpha").Cells(1, pi_FiliereConsideree)
    O2_rdt_transfert_force = Feuil30.Range("ANITA_Mox_O2_rdt_transfert").Cells(1, pi_FiliereConsideree)
    hauteur_bassin_force = Feuil30.Range("ANITA_Mox_hauteur_bassin").Cells(1, pi_FiliereConsideree)
    O2_dissous_force = Feuil30.Range("ANITA_Mox_O2_dissous").Cells(1, pi_FiliereConsideree)
    air_Q_Nm3j_force = Feuil30.Range("ANITA_Mox_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
    air_P_refoulement_force = Feuil30.Range("ANITA_Mox_air_P_refoulement").Cells(1, pi_FiliereConsideree)
    surpresseur_conso_spec_force = Feuil30.Range("ANITA_Mox_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)



    If temperature_force = True Then
        temperature_f = Feuil30.Range("ANITA_Mox_temperature").Cells(1, pi_FiliereConsideree)
    End If
    'début modif MSM 20/11/12
    If mn_NominalQ_m3j_force = True Then
        md_NominalQ_m3j_f = Feuil30.Range("ANITA_Mox_NominalQ_m3j").Cells(1, pi_FiliereConsideree)
    End If
    If mn_NominalNNH4_mgL_force = True Then
        md_NominalNNH4_mgL_f = Feuil30.Range("ANITA_Mox_NominalNNH4_mgL").Cells(1, pi_FiliereConsideree)
    End If
    If mn_NominalDCOSolubleDegradable_mgL_force = True Then
        md_NominalDCOSolubleDegradable_mgL_f = Feuil30.Range("ANITA_Mox_NominalDCOSolubleDegradable_mgL").Cells(1, pi_FiliereConsideree)
    End If
    If mn_NominalMES_mgL_force = True Then
        md_NominalMES_mgL_f = Feuil30.Range("ANITA_Mox_NominalMES_mgL").Cells(1, pi_FiliereConsideree)
    End If
    'fin modif MSM 20/11/12
    If cuve_C_filling_force = True Then
        cuve_C_filling_f = Feuil30.Range("ANITA_Mox_cuve_DBO_filling").Cells(1, pi_FiliereConsideree)
    End If
    If cuve_C_volume_force = True Then
        cuve_C_volume_f = Feuil30.Range("ANITA_Mox_cuve_DBO_volume").Cells(1, pi_FiliereConsideree)
    End If
    If cuve_Mox_filling_force = True Then
        cuve_Mox_filling_f = Feuil30.Range("ANITA_Mox_cuve_Mox_filling").Cells(1, pi_FiliereConsideree)
    End If
    If cuve_Mox_volume_force = True Then
        cuve_Mox_volume_f = Feuil30.Range("ANITA_Mox_cuve_Mox_volume").Cells(1, pi_FiliereConsideree)
    End If
    If O2_besoin_force = True Then
        O2_besoin_f = Feuil30.Range("ANITA_Mox_O2_besoin").Cells(1, pi_FiliereConsideree)
    End If
    If O2_facteur_alfa_force = True Then
        O2_facteur_alfa_f = Feuil30.Range("ANITA_Mox_O2_facteur_alpha").Cells(1, pi_FiliereConsideree)
    End If
    If O2_rdt_transfert_force = True Then
        O2_rdt_transfert_f = Feuil30.Range("ANITA_Mox_O2_rdt_transfert").Cells(1, pi_FiliereConsideree)
    End If
    If hauteur_bassin_force = True Then
        hauteur_bassin_f = Feuil30.Range("ANITA_Mox_hauteur_bassin").Cells(1, pi_FiliereConsideree)
    End If
    If O2_dissous_force = True Then
        O2_dissous_f = Feuil30.Range("ANITA_Mox_O2_dissous").Cells(1, pi_FiliereConsideree)
    End If
    If air_Q_Nm3j_force = True Then
        air_Q_Nm3j_f = Feuil30.Range("ANITA_Mox_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
    End If
    If air_P_refoulement_force = True Then
        air_P_refoulement_f = Feuil30.Range("ANITA_Mox_air_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If surpresseur_conso_spec_force = True Then
        surpresseur_conso_spec_f = Feuil30.Range("ANITA_Mox_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "30.2"

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

    Set mcls_RetoursAdmis = New CLS_RetoursAdmisTraitement
    If choix_MAP_retours = True Then
        mcls_RetoursAdmis.pn_EcritureResultats = False
    Else
        mcls_RetoursAdmis.pn_EcritureResultats = True
    End If

End Sub

Private Sub Class_Terminate()

    If pn_EffacementResultatsEnCours = False Then
        'Call calcul_consommation_electrique
        Call ecriture_resultats
        Call affich_remarque
    Else
        Call mcls_RetoursAdmis.effacer_resultats
    End If

    Set mcls_RetoursAdmis = Nothing

End Sub

