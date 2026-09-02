Attribute VB_Name = "z_ANITA_Shunt"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Dim mcls_RetoursAdmis As CLS_RetoursAdmisTraitement

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
Private Const regulation_Shunt = 1


Private Const critere_incoherence = 0.1



'choix de l'utilisateur
Dim choix_surpresseur As Integer
Dim choix_regulation As Integer

'variables à échanger avec la synthèse
Dim alim_jour_semaine As Double
Dim alim_heure_jour As Double
Dim volume_tampon As Double
Dim alim_nb_pompe As Double
Dim alim_P_refoulement As Double
Dim alim_tps_fct_pompe As Double
Dim alim_Qunitaire As Double
Dim alim_pompe_rdt As Double
Dim volume_SBR As Double
Dim rdt_NH4 As Double
Dim outlet_MES As Double
Public methanol_flux As Double
Dim boues_concentration As Double
Dim boues_MES As Double
Dim boues_Q As Double
Dim boues_MV_MES As Double
Dim extraction_nb_pompe As Double
Dim extraction_P_refoulement As Double
Dim extraction_tps_fct_pompe As Double
Dim extraction_Qunitaire As Double
Dim extraction_pompe_rdt As Double
Dim temperature As Double
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
Dim alim_jour_semaine_force As Boolean
Dim alim_heure_jour_force As Boolean
Dim volume_tampon_force As Boolean
Dim alim_nb_pompe_force As Boolean
Dim alim_P_refoulement_force As Boolean
Dim alim_tps_fct_pompe_force As Boolean
Dim alim_Qunitaire_force As Boolean
Dim alim_pompe_rdt_force As Boolean
Dim volume_SBR_force As Boolean
Dim rdt_NH4_force As Boolean
Dim outlet_MES_force As Boolean
Dim methanol_flux_force As Boolean
Dim boues_concentration_force As Boolean
Dim boues_MES_force As Boolean
'Dim boues_Q_force As Boolean
Dim boues_MV_MES_force As Boolean
Dim extraction_nb_pompe_force As Boolean
Dim extraction_P_refoulement_force As Boolean
Dim extraction_tps_fct_pompe_force As Boolean
Dim extraction_Qunitaire_force As Boolean
Dim extraction_pompe_rdt_force As Boolean
Dim temperature_force As Boolean
Dim O2_besoin_force As Boolean
Dim O2_rdt_transfert_force As Boolean
Dim O2_facteur_alfa_force As Boolean
Dim hauteur_bassin_force As Boolean
Dim O2_dissous_force As Boolean
Dim air_Q_Nm3j_force As Boolean
Dim air_P_refoulement_force As Boolean
Dim surpresseur_conso_spec_force As Boolean

'variables à échanger avec la synthèse
Dim alim_jour_semaine_f As Double
Dim alim_heure_jour_f As Double
Dim volume_tampon_f As Double
Dim alim_nb_pompe_f As Double
Dim alim_P_refoulement_f As Double
Dim alim_tps_fct_pompe_f As Double
Dim alim_Qunitaire_f As Double
Dim alim_pompe_rdt_f As Double
Dim volume_SBR_f As Double
Dim rdt_NH4_f As Double
Dim outlet_MES_f As Double
Dim methanol_flux_f As Double
Dim boues_concentration_f As Double
Dim boues_MES_f As Double
'Dim boues_Q_f As double
Dim boues_MV_MES_f As Double
Dim extraction_nb_pompe_f As Double
Dim extraction_P_refoulement_f As Double
Dim extraction_tps_fct_pompe_f As Double
Dim extraction_Qunitaire_f As Double
Dim extraction_pompe_rdt_f As Double
Dim temperature_f As Double
Dim O2_besoin_f As Double
Dim O2_rdt_transfert_f As Double
Dim O2_facteur_alfa_f As Double
Dim hauteur_bassin_f As Double
Dim O2_dissous_f As Double
Dim air_Q_Nm3j_f As Double
Dim air_P_refoulement_f As Double
Dim surpresseur_conso_spec_f As Double

'hypothèses
Dim tampon_tps_mini_retention As Double
Dim tampon_securite As Double
Dim G_SBR As Double
Dim inlet_MES_max As Double
Dim rdt_nit_vise As Double
Dim rdt_DBO_sol As Double
Dim rdt_NO2 As Double
Dim rate_nit_MVS As Double
Dim tps_anoxie As Double
Dim agitation_tampon_Pw As Double
Dim agitation_SBR_Pw As Double
Dim SBR_concentration_MES As Double
Dim ratio_O2_nit As Double
Dim ratio_O2_MES As Double
Dim facteur_DCO_apportee As Double
Dim DCO_methanol As Double
Dim impact_regulation(nb_type_regulation) As Double
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
Dim a0_correctif_alfa_MES As Double
Dim a1_correctif_alfa_MES As Double
Dim MES_reference_alfa As Double
Dim surpresseur_conso_Wh_Nm3mCE(nb_type_surpresseur, nb_surpresseur_gamme_Q, nb_surpresseur_gamme_P) As Double
Dim respO2_correctif_T As Double
Dim respO2_T_ref As Double

Dim inlet_Q As Double

Dim remarque(5) As String


Sub affich_remarque()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 5  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil40.Range("ANITA_Shunt_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf
    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "40.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub



Sub dimensionnement(ByRef retour_flux() As Double, ByRef retour_digestion() As Double, ByRef retour_flux_soluble() As Double, ByRef retour_digestion_soluble() As Double, ByRef retour_athos() As Double, ByRef retour_athos_soluble() As Double, ByRef boues_Shunt() As Double)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer


    Dim Q As Double
    Dim DCO As Double
    Dim DCO_soluble As Double
    Dim inlet_DCO As Double
    Dim DBO As Double
    Dim DBO_soluble As Double
    Dim inlet_DBO As Double
    Dim MES As Double
    Dim inlet_MES As Double
    Dim NK As Double
    Dim inlet_NK As Double
    Dim NK_soluble As Double
    Dim NH4 As Double
    Dim NO3 As Double
    Dim Pt As Double
    Dim Pt_soluble As Double
    Dim inlet_Pt As Double

    Dim tampon_tps_retention As Double  'tampon_tps_mini_retention
    Dim Q_m3h As Double
    Dim NH4_nit As Double
    Dim NH4_nit_max As Double
    Dim DBO_traitee As Double
    Dim SBR_MVS As Double


    'on ré-initialise les flux de boues issues du Shunt
    For I = 1 To repere_graisse_Pt_MES
        boues_Shunt(I) = 0
    Next I


    Dim retour_flux_non_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_soluble_non_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_soluble_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_admis(retour_caracteristique_nb) As Double

    'On regarde les retours admis
    Call mcls_RetoursAdmis.dimensionnement(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble, retour_flux_non_admis, retour_flux_soluble_non_admis, retour_flux_admis, retour_flux_soluble_admis)

    'on récupère les valeurs des retours
    Q = mcls_RetoursAdmis.pd_Q_m3j
    DCO = mcls_RetoursAdmis.pd_DCO_mgL * Q / 1000
        DCO_soluble = retour_flux_soluble_admis(repere_ret_DCO)
    DBO = mcls_RetoursAdmis.pd_DBO_mgL * Q / 1000
        DBO_soluble = retour_flux_soluble_admis(repere_ret_DBO)
    MES = mcls_RetoursAdmis.pd_MES_mgL * Q / 1000
    NK = mcls_RetoursAdmis.pd_NK_mgL * Q / 1000
        NK_soluble = retour_flux_soluble_admis(repere_ret_NK)
    NH4 = mcls_RetoursAdmis.pd_NNH4_mgL * Q / 1000
    NO3 = mcls_RetoursAdmis.pd_NNO3_mgL * Q / 1000
    Pt = mcls_RetoursAdmis.pd_Pt_mgL * Q / 1000
        Pt_soluble = retour_flux_soluble_admis(repere_ret_Pt)

    inlet_MES = MES
    inlet_DCO = DCO
    inlet_DBO = DBO
    inlet_NK = NK
    inlet_Pt = Pt
    inlet_Q = Q

    'Dimensionnement de la bâche tampon
    If volume_tampon_force = True Then
        volume_tampon = volume_tampon_f
    Else
        'débit horaire réel
        Q_m3h = Q / alim_jour_semaine * pd_NOMBRE_JOUR_PAR_SEMAINE
        Q_m3h = Q_m3h / alim_heure_jour
        If (pd_NOMBRE_HEURE_PAR_JOUR - alim_heure_jour) <= tampon_tps_mini_retention Then
            tampon_tps_retention = tampon_tps_mini_retention
            volume_tampon = tampon_tps_retention * Q_m3h
        Else
            tampon_tps_retention = pd_NOMBRE_HEURE_PAR_JOUR - alim_heure_jour
            volume_tampon = tampon_tps_retention * Q_m3h
        End If
        volume_tampon = volume_tampon * tampon_securite
    End If

    'Nécessité d'un décanteur
    If MES / Q > inlet_MES_max Then
        remarque(1) = Feuil40.Range("ANITA_Shunt_erreur_inlet_MES").Cells(1, 1)
    End If

    'Dimensionnement volume de SBR
    NH4_nit_max = NH4
    If volume_SBR_force = True Then
        volume_SBR = volume_SBR_f
        SBR_MVS = (SBR_concentration_MES * boues_MV_MES) * volume_SBR
        NH4_nit = rate_nit_MVS * SBR_MVS
        If NH4_nit > NH4_nit_max Then
            NH4_nit = NH4_nit_max
        End If
    Else
        NH4_nit = rdt_nit_vise * NH4
        SBR_MVS = NH4_nit / rate_nit_MVS
        volume_SBR = SBR_MVS / (SBR_concentration_MES * boues_MV_MES)
    End If

    'Ajout de méthanol et élimination de la DBO
    DBO_traitee = DBO_soluble * rdt_DBO_sol
    If facteur_DCO_apportee * (rdt_NO2 * NH4_nit) > DBO_traitee Then
        methanol_flux = (facteur_DCO_apportee * (rdt_NO2 * NH4_nit) - DBO_traitee) / DCO_methanol
    Else
        methanol_flux = 0
    End If
    
    'Besoins en O2
    O2_besoin = ratio_O2_nit * NH4_nit + ratio_O2_MES * respO2_correctif_T ^ (temperature - respO2_T_ref) * (SBR_MVS / boues_MV_MES)

    'sortie boues
    boues_MES = (volume_SBR * SBR_concentration_MES) / G_SBR - outlet_MES * (Q - boues_Q) / 1000   'boues_Q=0 d'abord puis garde la valeur par la suite de l'itération antérieure
    boues_Q = boues_MES / boues_concentration


    'sortie eau
    Q = Q - boues_Q
    MES = outlet_MES * Q / 1000
    DCO = (DCO_soluble - DBO_traitee) + MES * (inlet_DCO - DCO_soluble) / inlet_MES
    DBO = (DBO_soluble - DBO_traitee) + MES * (inlet_DBO - DBO_soluble) / inlet_MES
    NK = (NK_soluble - NH4_nit) + MES * (inlet_NK - NK_soluble) / inlet_MES
    NH4 = NH4 - NH4_nit
    NO3 = NO3 * Q / (Q + boues_Q)


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
    retour_flux_soluble_admis(repere_ret_DCO) = (DCO_soluble - DBO_traitee)
    retour_flux_soluble_admis(repere_ret_DBO) = (DBO_soluble - DBO_traitee)
    retour_flux_soluble_admis(repere_ret_NO3) = NO3
    Call mcls_RetoursAdmis.flux_sortie_process(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble, retour_flux_non_admis, retour_flux_soluble_non_admis, retour_flux_admis, retour_flux_soluble_admis)


    'on récupère les valeurs pour les boues
    boues_Shunt(repere_graisse_MS) = boues_MES
    boues_Shunt(repere_graisse_Q) = boues_Q
    boues_Shunt(repere_graisse_MV_MES) = boues_MV_MES
    boues_Shunt(repere_graisse_DCO_MES) = (inlet_DCO - (DCO + DBO_traitee)) / boues_MES
    boues_Shunt(repere_graisse_DBO_MES) = (inlet_DBO - (DBO + DBO_traitee)) / boues_MES
    boues_Shunt(repere_graisse_NK_MES) = (inlet_NK - (NK + NH4_nit)) / boues_MES
    boues_Shunt(repere_graisse_Pt_MES) = (inlet_Pt - Pt) / boues_MES


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "40.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub



Sub fonctionnement_reel(ByRef retour_flux() As Double, ByRef retour_digestion() As Double, ByRef retour_flux_soluble() As Double, ByRef retour_digestion_soluble() As Double, ByRef retour_athos() As Double, ByRef retour_athos_soluble() As Double, ByRef boues_Shunt() As Double, ByRef DCO_methanol_Shunt As Double)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String




    Dim I As Integer

    Dim Q As Double
    Dim DCO As Double
    Dim DCO_soluble As Double
    Dim inlet_DCO As Double
    Dim DBO As Double
    Dim DBO_soluble As Double
    Dim inlet_DBO As Double
    Dim MES As Double
    Dim inlet_MES As Double
    Dim NK As Double
    Dim NK_soluble As Double
    Dim inlet_NK As Double
    Dim NH4 As Double
    Dim NO3 As Double
    Dim Pt As Double
    Dim Pt_soluble As Double
    Dim inlet_Pt As Double

    Dim NH4_nit As Double
    Dim NH4_nit_max As Double
    Dim DBO_traitee As Double
    Dim SBR_MVS As Double

    'on ré-initialise les flux de boues issues du Shunt
    For I = 1 To repere_graisse_Pt_MES
        boues_Shunt(I) = 0
    Next I

    Dim retour_flux_non_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_soluble_non_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_soluble_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_admis(retour_caracteristique_nb) As Double

    'On regarde les retours admis
    Call mcls_RetoursAdmis.fonctionnement_reel(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble, retour_flux_non_admis, retour_flux_soluble_non_admis, retour_flux_admis, retour_flux_soluble_admis)

    'on récupère les valeurs des retours
    Q = mcls_RetoursAdmis.pd_Q_m3j
    DCO = mcls_RetoursAdmis.pd_DCO_mgL * Q / 1000
        DCO_soluble = retour_flux_soluble_admis(repere_ret_DCO)
    DBO = mcls_RetoursAdmis.pd_DBO_mgL * Q / 1000
        DBO_soluble = retour_flux_soluble_admis(repere_ret_DBO)
    MES = mcls_RetoursAdmis.pd_MES_mgL * Q / 1000
    NK = mcls_RetoursAdmis.pd_NK_mgL * Q / 1000
        NK_soluble = retour_flux_soluble_admis(repere_ret_NK)
    NH4 = mcls_RetoursAdmis.pd_NNH4_mgL * Q / 1000
    NO3 = mcls_RetoursAdmis.pd_NNO3_mgL * Q / 1000
    Pt = mcls_RetoursAdmis.pd_Pt_mgL * Q / 1000
        Pt_soluble = retour_flux_soluble_admis(repere_ret_Pt)

    inlet_MES = MES
    inlet_DCO = DCO
    inlet_DBO = DBO
    inlet_NK = NK
    inlet_Pt = Pt
    inlet_Q = Q

    'Performances du traitement
    If rdt_NH4_force = True Then
        rdt_NH4 = rdt_NH4_f
        NH4_nit = rdt_NH4 * NH4
    Else
        NH4_nit_max = NH4
        SBR_MVS = boues_MV_MES * volume_SBR * SBR_concentration_MES
        NH4_nit = rate_nit_MVS * SBR_MVS
        If NH4_nit > NH4_nit_max Then
            NH4_nit = NH4_nit_max
        End If
        rdt_NH4 = NH4_nit / NH4
    End If

    'Ajout de méthanol et élimination de la DBO
    DBO_traitee = DBO_soluble * rdt_DBO_sol
    If methanol_flux_force = True Then
        methanol_flux = methanol_flux_f
    Else
        If facteur_DCO_apportee * (rdt_NO2 * NH4_nit) > DBO_traitee Then
            methanol_flux = (facteur_DCO_apportee * (rdt_NO2 * NH4_nit) - DBO_traitee) / DCO_methanol
        Else
            methanol_flux = 0
        End If
    End If
    DCO_methanol_Shunt = methanol_flux * DCO_methanol

    'Besoins en O2
    If O2_besoin_force = True Then
        O2_besoin = O2_besoin_f
    Else
        O2_besoin = ratio_O2_nit * NH4_nit + ratio_O2_MES * respO2_correctif_T ^ (temperature - respO2_T_ref) * (SBR_MVS / boues_MV_MES)
    End If
    
    'sortie boues
    If boues_MES_force = True Then
        boues_MES = boues_MES_f
    Else
        boues_MES = (volume_SBR * SBR_concentration_MES) / G_SBR - outlet_MES * (Q - boues_Q) / 1000   'boues_Q=0 d'abord puis garde la valeur par la suite de l'itération antérieure
    End If
    boues_Q = boues_MES / boues_concentration


    'sortie eau
    Q = Q - boues_Q
    MES = outlet_MES * Q / 1000
    DCO = (DCO_soluble - DBO_traitee) + MES * (inlet_DCO - DCO_soluble) / inlet_MES
    DBO = (DBO_soluble - DBO_traitee) + MES * (inlet_DBO - DBO_soluble) / inlet_MES
    NK = (NK_soluble - NH4_nit) + MES * (inlet_NK - NK_soluble) / inlet_MES
    NH4 = NH4 - NH4_nit
    NO3 = NO3 * Q / (Q + boues_Q)
    Pt = Pt_soluble + MES * (inlet_Pt - Pt_soluble) / inlet_MES


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
    retour_flux_soluble_admis(repere_ret_DCO) = (DCO_soluble - DBO_traitee)
    retour_flux_soluble_admis(repere_ret_DBO) = (DBO_soluble - DBO_traitee)
    retour_flux_soluble_admis(repere_ret_NO3) = NO3
    Call mcls_RetoursAdmis.flux_sortie_process(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble, retour_flux_non_admis, retour_flux_soluble_non_admis, retour_flux_admis, retour_flux_soluble_admis)


    'on récupère les valeurs pour les boues
    boues_Shunt(repere_graisse_MS) = boues_MES
    boues_Shunt(repere_graisse_Q) = boues_Q
    boues_Shunt(repere_graisse_MV_MES) = boues_MV_MES
    boues_Shunt(repere_graisse_DCO_MES) = (inlet_DCO - (DCO + DBO_traitee)) / boues_MES
    boues_Shunt(repere_graisse_DBO_MES) = (inlet_DBO - (DBO + DBO_traitee)) / boues_MES
    boues_Shunt(repere_graisse_NK_MES) = (inlet_NK - (NK + NH4_nit)) / boues_MES
    boues_Shunt(repere_graisse_Pt_MES) = (inlet_Pt - Pt) / boues_MES


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "40.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String




    'fonctionnement centrifs
    If alim_jour_semaine_force = True Then
        alim_jour_semaine = alim_jour_semaine_f
    Else
        alim_jour_semaine = 7    'jours/semaine
    End If
    If alim_heure_jour_force = True Then
        alim_heure_jour = alim_heure_jour_f
    Else
        alim_heure_jour = 24   'h/j
    End If

    'alimentation de l'eau vers les SBR
    If alim_nb_pompe_force = True Then
        alim_nb_pompe = alim_nb_pompe_f
    Else
        alim_nb_pompe = 2
    End If
    If alim_P_refoulement_force = True Then
        alim_P_refoulement = alim_P_refoulement_f
    Else
        alim_P_refoulement = 5
    End If
    If alim_tps_fct_pompe_force = True Then
        alim_tps_fct_pompe = alim_tps_fct_pompe_f
    Else
        alim_tps_fct_pompe = pd_NOMBRE_HEURE_PAR_JOUR
    End If

    'température traitement
    If temperature_force = True Then
        temperature = temperature_f
    Else
        temperature = 25    '°C
    End If

    'sortie MES effluent   (mg/L)
    If outlet_MES_force = True Then
        outlet_MES = outlet_MES_f
    Else
        outlet_MES = 100
    End If

    'rendement de dissolution de l'O2 en eau claire  (%/m)
    If O2_rdt_transfert_force = True Then
        O2_rdt_transfert = O2_rdt_transfert_f
    Else
        O2_rdt_transfert = 5.5
    End If

    'facteur alpha
    If O2_facteur_alfa_force = True Then
        O2_facteur_alfa = O2_facteur_alfa_f
    Else
        O2_facteur_alfa = 0.69
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
        O2_dissous = 1
    End If

    'Pression de refoulement des surpresseurs (mCE)
    If air_P_refoulement_force = True Then
        air_P_refoulement = air_P_refoulement_f
    Else
        air_P_refoulement = 2 + hauteur_bassin
    End If

    'concentration des boues extraites (g/L)
    If boues_concentration_force = True Then
        boues_concentration = boues_concentration_f
    Else
        boues_concentration = 1000 / 150    '1000/indice de boues    cf Kim Sorensen   (avec IB=150mL/g; valeur 1000 au lieu de 1200 car a priori sécurité)
    End If

    'ratio de MV/MES des boues extraites
    If boues_MV_MES_force = True Then
        boues_MV_MES = boues_MV_MES_f
    Else
        boues_MV_MES = 0.8   'valeur Julien Chauzy
    End If

    'extraction des boues
    If extraction_nb_pompe_force = True Then
        extraction_nb_pompe = extraction_nb_pompe_f
    Else
        extraction_nb_pompe = 2
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement = extraction_P_refoulement_f
    Else
        extraction_P_refoulement = 15
    End If
    If extraction_tps_fct_pompe_force = True Then
        extraction_tps_fct_pompe = extraction_tps_fct_pompe_f
    Else
        extraction_tps_fct_pompe = pd_NOMBRE_HEURE_PAR_JOUR
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "40.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub lecture_choix()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    'lecture des choix internes
    choix_surpresseur = Feuil40.Range("ANITA_Shunt_choix_surpresseur").Cells(1, pi_FiliereConsideree)
    choix_regulation = Feuil40.Range("ANITA_Shunt_choix_regulation").Cells(1, pi_FiliereConsideree)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "40.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub



Sub lecture_valeurs_forcees()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String




    temperature_force = Feuil40.Range("ANITA_Shunt_temperature").Cells(1, pi_FiliereConsideree)
    alim_jour_semaine_force = Feuil40.Range("ANITA_Shunt_alim_jour_semaine").Cells(1, pi_FiliereConsideree)
    alim_heure_jour_force = Feuil40.Range("ANITA_Shunt_alim_heure_jour").Cells(1, pi_FiliereConsideree)
    volume_tampon_force = Feuil40.Range("ANITA_Shunt_volume_tampon").Cells(1, pi_FiliereConsideree)
    alim_nb_pompe_force = Feuil40.Range("ANITA_Shunt_alim_nb_pompe").Cells(1, pi_FiliereConsideree)
    alim_P_refoulement_force = Feuil40.Range("ANITA_Shunt_alim_Prefoulement").Cells(1, pi_FiliereConsideree)
    alim_tps_fct_pompe_force = Feuil40.Range("ANITA_Shunt_alim_tps_fct_pompe").Cells(1, pi_FiliereConsideree)
    alim_Qunitaire_force = Feuil40.Range("ANITA_Shunt_alim_Qunitaire").Cells(1, pi_FiliereConsideree)
    alim_pompe_rdt_force = Feuil40.Range("ANITA_Shunt_alim_pompe_rdt").Cells(1, pi_FiliereConsideree)
    volume_SBR_force = Feuil40.Range("ANITA_Shunt_volume_SBR").Cells(1, pi_FiliereConsideree)
    rdt_NH4_force = Feuil40.Range("ANITA_Shunt_rdt_NH4").Cells(1, pi_FiliereConsideree)
    outlet_MES_force = Feuil40.Range("ANITA_Shunt_outlet_MES").Cells(1, pi_FiliereConsideree)
    methanol_flux_force = Feuil40.Range("ANITA_Shunt_methanol_flux").Cells(1, pi_FiliereConsideree)
    O2_besoin_force = Feuil40.Range("ANITA_Shunt_O2_besoin").Cells(1, pi_FiliereConsideree)
    O2_facteur_alfa_force = Feuil40.Range("ANITA_Shunt_O2_facteur_alpha").Cells(1, pi_FiliereConsideree)
    O2_rdt_transfert_force = Feuil40.Range("ANITA_Shunt_O2_rdt_transfert").Cells(1, pi_FiliereConsideree)
    hauteur_bassin_force = Feuil40.Range("ANITA_Shunt_hauteur_bassin").Cells(1, pi_FiliereConsideree)
    O2_dissous_force = Feuil40.Range("ANITA_Shunt_O2_dissous").Cells(1, pi_FiliereConsideree)
    air_Q_Nm3j_force = Feuil40.Range("ANITA_Shunt_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
    air_P_refoulement_force = Feuil40.Range("ANITA_Shunt_air_P_refoulement").Cells(1, pi_FiliereConsideree)
    surpresseur_conso_spec_force = Feuil40.Range("ANITA_Shunt_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
    boues_concentration_force = Feuil40.Range("ANITA_Shunt_boues_concentration").Cells(1, pi_FiliereConsideree)
    boues_MES_force = Feuil40.Range("ANITA_Shunt_boues_MES").Cells(1, pi_FiliereConsideree)
    boues_MV_MES_force = Feuil40.Range("ANITA_Shunt_MV_MES").Cells(1, pi_FiliereConsideree)
    extraction_nb_pompe_force = Feuil40.Range("ANITA_Shunt_extraction_nb_pompe").Cells(1, pi_FiliereConsideree)
    extraction_P_refoulement_force = Feuil40.Range("ANITA_Shunt_extraction_Prefoulement").Cells(1, pi_FiliereConsideree)
    extraction_tps_fct_pompe_force = Feuil40.Range("ANITA_Shunt_extraction_tps_fct_pompe").Cells(1, pi_FiliereConsideree)
    extraction_Qunitaire_force = Feuil40.Range("ANITA_Shunt_extraction_Qunitaire").Cells(1, pi_FiliereConsideree)
    extraction_pompe_rdt_force = Feuil40.Range("ANITA_Shunt_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)


    If temperature_force = True Then
        temperature_f = Feuil40.Range("ANITA_Shunt_temperature").Cells(1, pi_FiliereConsideree)
    End If
    If alim_jour_semaine_force = True Then
        alim_jour_semaine_f = Feuil40.Range("ANITA_Shunt_alim_jour_semaine").Cells(1, pi_FiliereConsideree)
    End If
    If alim_heure_jour_force = True Then
        alim_heure_jour_f = Feuil40.Range("ANITA_Shunt_alim_heure_jour").Cells(1, pi_FiliereConsideree)
    End If
    If volume_tampon_force = True Then
        volume_tampon_f = Feuil40.Range("ANITA_Shunt_volume_tampon").Cells(1, pi_FiliereConsideree)
    End If
    If alim_nb_pompe_force = True Then
        alim_nb_pompe_f = Feuil40.Range("ANITA_Shunt_alim_nb_pompe").Cells(1, pi_FiliereConsideree)
    End If
    If alim_P_refoulement_force = True Then
        alim_P_refoulement_f = Feuil40.Range("ANITA_Shunt_alim_Prefoulement").Cells(1, pi_FiliereConsideree)
    End If
    If alim_tps_fct_pompe_force = True Then
        alim_tps_fct_pompe_f = Feuil40.Range("ANITA_Shunt_alim_tps_fct_pompe").Cells(1, pi_FiliereConsideree)
    End If
    If alim_Qunitaire_force = True Then
        alim_Qunitaire_f = Feuil40.Range("ANITA_Shunt_alim_Qunitaire").Cells(1, pi_FiliereConsideree)
    End If
    If alim_pompe_rdt_force = True Then
        alim_pompe_rdt_f = Feuil40.Range("ANITA_Shunt_alim_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    If volume_SBR_force = True Then
        volume_SBR_f = Feuil40.Range("ANITA_Shunt_volume_SBR").Cells(1, pi_FiliereConsideree)
    End If
    If rdt_NH4_force = True Then
        rdt_NH4_f = Feuil40.Range("ANITA_Shunt_rdt_NH4").Cells(1, pi_FiliereConsideree)
    End If
    If outlet_MES_force = True Then
        outlet_MES_f = Feuil40.Range("ANITA_Shunt_outlet_MES").Cells(1, pi_FiliereConsideree)
    End If
    If methanol_flux_force = True Then
        methanol_flux_f = Feuil40.Range("ANITA_Shunt_methanol_flux").Cells(1, pi_FiliereConsideree)
    End If
    If O2_besoin_force = True Then
        O2_besoin_f = Feuil40.Range("ANITA_Shunt_O2_besoin").Cells(1, pi_FiliereConsideree)
    End If
    If O2_facteur_alfa_force = True Then
        O2_facteur_alfa_f = Feuil40.Range("ANITA_Shunt_O2_facteur_alpha").Cells(1, pi_FiliereConsideree)
    End If
    If O2_rdt_transfert_force = True Then
        O2_rdt_transfert_f = Feuil40.Range("ANITA_Shunt_O2_rdt_transfert").Cells(1, pi_FiliereConsideree)
    End If
    If hauteur_bassin_force = True Then
        hauteur_bassin_f = Feuil40.Range("ANITA_Shunt_hauteur_bassin").Cells(1, pi_FiliereConsideree)
    End If
    If O2_dissous_force = True Then
        O2_dissous_f = Feuil40.Range("ANITA_Shunt_O2_dissous").Cells(1, pi_FiliereConsideree)
    End If
    If air_Q_Nm3j_force = True Then
        air_Q_Nm3j_f = Feuil40.Range("ANITA_Shunt_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
    End If
    If air_P_refoulement_force = True Then
        air_P_refoulement_f = Feuil40.Range("ANITA_Shunt_air_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If surpresseur_conso_spec_force = True Then
        surpresseur_conso_spec_f = Feuil40.Range("ANITA_Shunt_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
    End If
    If boues_concentration_force = True Then
        boues_concentration_f = Feuil40.Range("ANITA_Shunt_boues_concentration").Cells(1, pi_FiliereConsideree)
    End If
    If boues_MES_force = True Then
        boues_MES_f = Feuil40.Range("ANITA_Shunt_boues_MES").Cells(1, pi_FiliereConsideree)
    End If
    If boues_MV_MES_force = True Then
        boues_MV_MES_f = Feuil40.Range("ANITA_Shunt_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_nb_pompe_force = True Then
        extraction_nb_pompe_f = Feuil40.Range("ANITA_Shunt_extraction_nb_pompe").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement_f = Feuil40.Range("ANITA_Shunt_extraction_Prefoulement").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_tps_fct_pompe_force = True Then
        extraction_tps_fct_pompe_f = Feuil40.Range("ANITA_Shunt_extraction_tps_fct_pompe").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_Qunitaire_force = True Then
        extraction_Qunitaire_f = Feuil40.Range("ANITA_Shunt_extraction_Qunitaire").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt_f = Feuil40.Range("ANITA_Shunt_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "40.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub ecriture_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil40.Range("ANITA_Shunt_alim_jour_semaine").Cells(1, Colonne) = alim_jour_semaine
    Feuil40.Range("ANITA_Shunt_alim_heure_jour").Cells(1, Colonne) = alim_heure_jour
    Feuil40.Range("ANITA_Shunt_volume_tampon").Cells(1, Colonne) = volume_tampon
    Feuil40.Range("ANITA_Shunt_alim_nb_pompe").Cells(1, Colonne) = alim_nb_pompe
    Feuil40.Range("ANITA_Shunt_alim_Prefoulement").Cells(1, Colonne) = alim_P_refoulement
    Feuil40.Range("ANITA_Shunt_alim_tps_fct_pompe").Cells(1, Colonne) = alim_tps_fct_pompe
    Feuil40.Range("ANITA_Shunt_alim_Qunitaire").Cells(1, Colonne) = alim_Qunitaire
    Feuil40.Range("ANITA_Shunt_alim_pompe_rdt").Cells(1, Colonne) = alim_pompe_rdt

    Feuil40.Range("ANITA_Shunt_rdt_NH4").Cells(1, Colonne) = rdt_NH4
    Feuil40.Range("ANITA_Shunt_outlet_MES").Cells(1, Colonne) = outlet_MES
    Feuil40.Range("ANITA_Shunt_methanol_flux").Cells(1, Colonne) = methanol_flux
    Feuil40.Range("ANITA_Shunt_boues_concentration").Cells(1, Colonne) = boues_concentration
    Feuil40.Range("ANITA_Shunt_boues_MES").Cells(1, Colonne) = boues_MES
    Feuil40.Range("ANITA_Shunt_boues_Q").Cells(1, Colonne) = boues_Q
    Feuil40.Range("ANITA_Shunt_MV_MES").Cells(1, Colonne) = boues_MV_MES
    Feuil40.Range("ANITA_Shunt_extraction_nb_pompe").Cells(1, Colonne) = extraction_nb_pompe
    Feuil40.Range("ANITA_Shunt_extraction_Prefoulement").Cells(1, Colonne) = extraction_P_refoulement
    Feuil40.Range("ANITA_Shunt_extraction_tps_fct_pompe").Cells(1, Colonne) = extraction_tps_fct_pompe
    Feuil40.Range("ANITA_Shunt_extraction_Qunitaire").Cells(1, Colonne) = extraction_Qunitaire
    Feuil40.Range("ANITA_Shunt_extraction_pompe_rdt").Cells(1, Colonne) = extraction_pompe_rdt

    Feuil40.Range("ANITA_Shunt_temperature").Cells(1, Colonne) = temperature
    Feuil40.Range("ANITA_Shunt_volume_SBR").Cells(1, Colonne) = volume_SBR
    Feuil40.Range("ANITA_Shunt_O2_besoin").Cells(1, Colonne) = O2_besoin
    Feuil40.Range("ANITA_Shunt_O2_facteur_alpha").Cells(1, Colonne) = O2_facteur_alfa
    Feuil40.Range("ANITA_Shunt_O2_rdt_transfert").Cells(1, Colonne) = O2_rdt_transfert
    Feuil40.Range("ANITA_Shunt_hauteur_bassin").Cells(1, Colonne) = hauteur_bassin
    Feuil40.Range("ANITA_Shunt_O2_dissous").Cells(1, Colonne) = O2_dissous
    Feuil40.Range("ANITA_Shunt_air_Q_Nm3j").Cells(1, Colonne) = air_Q_Nm3j
    Feuil40.Range("ANITA_Shunt_air_P_refoulement").Cells(1, Colonne) = air_P_refoulement
    Feuil40.Range("ANITA_Shunt_surpresseur_conso_spec").Cells(1, Colonne) = surpresseur_conso_spec
    Feuil40.Range("ANITA_Shunt_electricite").Cells(1, Colonne) = electricite

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "40.9"

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
    Dim electricite_alim As Double
    Dim electricite_extraction As Double
    Dim electricite_agitation As Double

    Dim incoherence As Double
    Dim rendement_pompe As Double
    Dim puissance_pompe As Double
    Dim rendement_moteur As Double



    'POMPAGES
    'ALIMENTATION EAU
    'alim_Qunitaire = inlet_Q / (alim_nb_pompe * alim_tps_fct_pompe)
    If alim_Qunitaire_force = True Then
        alim_Qunitaire = alim_Qunitaire_f
        If alim_tps_fct_pompe_force = True Then
            'on ne fait rien
            If alim_nb_pompe_force = False Then
                If Int((inlet_Q) / (alim_tps_fct_pompe * alim_Qunitaire)) < ((inlet_Q) / (alim_tps_fct_pompe * alim_Qunitaire)) Then
                    alim_nb_pompe = Int((inlet_Q) / (alim_tps_fct_pompe * alim_Qunitaire)) + 1
                Else
                    alim_nb_pompe = Int((inlet_Q) / (alim_tps_fct_pompe * alim_Qunitaire))
                End If
            End If
        Else
            alim_tps_fct_pompe = (inlet_Q) / (alim_nb_pompe * alim_Qunitaire)
            If alim_tps_fct_pompe > pd_NOMBRE_HEURE_PAR_JOUR Then
                alim_tps_fct_pompe = pd_NOMBRE_HEURE_PAR_JOUR
                If alim_nb_pompe_force = False Then
                    If Int((inlet_Q) / (alim_tps_fct_pompe * alim_Qunitaire)) < ((inlet_Q) / (alim_tps_fct_pompe * alim_Qunitaire)) Then
                        alim_nb_pompe = Int((inlet_Q) / (alim_tps_fct_pompe * alim_Qunitaire)) + 1
                    Else
                        alim_nb_pompe = Int((inlet_Q) / (alim_tps_fct_pompe * alim_Qunitaire))
                    End If
                    alim_tps_fct_pompe = (inlet_Q) / (alim_nb_pompe * alim_Qunitaire)
                End If
            End If
        End If
    Else
        alim_Qunitaire = (inlet_Q) / (alim_nb_pompe * alim_tps_fct_pompe)
    End If
    incoherence = (alim_Qunitaire * alim_nb_pompe * alim_tps_fct_pompe - (inlet_Q)) / (inlet_Q)
    If Abs(incoherence) > critere_incoherence Then
        remarque(2) = Feuil40.Range("ANITA_Shunt_erreur_pompe_alim").Cells(1, 1)
    End If
    If alim_pompe_rdt_force = True Then
        alim_pompe_rdt = alim_pompe_rdt_f
    Else
        rendement_pompe = 0.6    'hypothèse
        puissance_pompe = alim_Qunitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * alim_P_refoulement / rendement_pompe
        If puissance_pompe < 60 Then
            rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
        Else
            rendement_moteur = 0.961
        End If
        alim_pompe_rdt = rendement_pompe * rendement_moteur
    End If
    electricite_alim = alim_Qunitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * alim_P_refoulement / alim_pompe_rdt * alim_nb_pompe * alim_tps_fct_pompe
    electricite = electricite + electricite_alim
    'EXTRACTION DES BOUES
    'extraction_Qunitaire = boues_Q / (extraction_nb_pompe * extraction_tps_fct_pompe)
    If extraction_Qunitaire_force = True Then
        extraction_Qunitaire = extraction_Qunitaire_f
        If extraction_tps_fct_pompe_force = True Then
            'on ne fait rien
            If extraction_nb_pompe_force = False Then
                If Int(boues_Q / (extraction_tps_fct_pompe * extraction_Qunitaire)) < (boues_Q / (extraction_tps_fct_pompe * extraction_Qunitaire)) Then
                    extraction_nb_pompe = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Qunitaire)) + 1
                Else
                    extraction_nb_pompe = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Qunitaire))
                End If
            End If
        Else
            extraction_tps_fct_pompe = boues_Q / (extraction_nb_pompe * extraction_Qunitaire)
            If extraction_tps_fct_pompe > pd_NOMBRE_HEURE_PAR_JOUR Then
                extraction_tps_fct_pompe = pd_NOMBRE_HEURE_PAR_JOUR
                If extraction_nb_pompe_force = False Then
                    If Int(boues_Q / (extraction_tps_fct_pompe * extraction_Qunitaire)) < (boues_Q / (extraction_tps_fct_pompe * extraction_Qunitaire)) Then
                        extraction_nb_pompe = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Qunitaire)) + 1
                    Else
                        extraction_nb_pompe = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Qunitaire))
                    End If
                    extraction_tps_fct_pompe = boues_Q / (extraction_nb_pompe * extraction_Qunitaire)
                End If
            End If
        End If
    Else
        extraction_Qunitaire = boues_Q / (extraction_nb_pompe * extraction_tps_fct_pompe)
    End If
    incoherence = (extraction_Qunitaire * extraction_nb_pompe * extraction_tps_fct_pompe - boues_Q) / boues_Q
    If Abs(incoherence) > critere_incoherence Then
        remarque(3) = Feuil40.Range("ANITA_Shunt_erreur_pompe_extraction").Cells(1, 1)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt = extraction_pompe_rdt_f
    Else
        rendement_pompe = 0.3    'hypothèse
        puissance_pompe = extraction_Qunitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * extraction_P_refoulement / rendement_pompe
        If puissance_pompe < 60 Then
            rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
        Else
            rendement_moteur = 0.961
        End If
        extraction_pompe_rdt = rendement_pompe * rendement_moteur
    End If
    electricite_extraction = extraction_Qunitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * extraction_P_refoulement / extraction_pompe_rdt * extraction_nb_pompe * extraction_tps_fct_pompe
    electricite = electricite + electricite_extraction






    'AGITATION
    electricite_agitation = (agitation_tampon_Pw * volume_tampon * pd_NOMBRE_HEURE_PAR_JOUR + agitation_SBR_Pw * volume_SBR * tps_anoxie) * alim_jour_semaine / pd_NOMBRE_JOUR_PAR_SEMAINE / 1000
    electricite = electricite + electricite_agitation


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
        If SBR_concentration_MES > MES_reference_alfa And O2_facteur_alfa_force = False Then
            O2_facteur_alfa = O2_facteur_alfa * (a0_correctif_alfa_MES + a1_correctif_alfa_MES * SBR_concentration_MES) / (a0_correctif_alfa_MES + a1_correctif_alfa_MES * MES_reference_alfa)
        End If
        O2_facteur_K = O2_facteur_alfa * (O2_facteur_beta * O2_saturation_T_P - O2_dissous) / O2_saturation_20_P
        O2_facteur_K = O2_facteur_K * correctif_T_K ^ (temperature - T_ref_K)
        air_Q_Nm3j = O2_besoin / (O2_facteur_K * O2_rdt_transfert / 100 * (hauteur_bassin - hauteur_diffuseur) * ratio_kgO2_Nm3air)
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

    electricite = electricite + electricite_aeration



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "40.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub



Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim I As Integer
    Dim j As Integer


    'temps retention bassin tampon  (h)
    tampon_tps_mini_retention = 4

    'sécurité sur le dimensionnement tampon 10%
    tampon_securite = 1.1

    'âge de boues par défaut   (en jours)
    G_SBR = 15

    'concentration max en entrée en MES
    inlet_MES_max = 1.2    'g/L

    'rendement de nit visé
    rdt_nit_vise = 0.95

    'rendement de DBO_soluble
    rdt_DBO_sol = 0.95

    'rendement denit   (tous les nitrites sont éliminés)
    rdt_NO2 = 1

    'vitesse de nitrification en phase aérée (kgNNH4/(kgMVS.j))
    rate_nit_MVS = 6 / 1000 * 13.5  ' 6kgNNH4/(kgMVS.h aération) * tps_aération ( 4.5h/cycle de 8h --> 13.5h/jour )

    'temps d'anoxie   h/j
    tps_anoxie = 1.5 / 8 * 24   '(1.5h/cycle de 8h)

    'agitation
    agitation_tampon_Pw = 3    'W/m3
    agitation_SBR_Pw = 3    'W/m3

    'concentration en MES en fin de cycle    g/L
    SBR_concentration_MES = 5

    'consommation spécifique en O2
    'nitrification (kg O2/kg N nitrifié)
    ratio_O2_nit = 4.57 * 0.75
    'respiration (kg O2/kg MES)
    ratio_O2_MES = 24 * 0.002    'cf tableur HybAS
    'correctifs pour la conso O2 pour respiration
    respO2_correctif_T = 1.07
    respO2_T_ref = 10

    'consommation spécifique de méthanol
    facteur_DCO_apportee = 4 * 0.6
    DCO_methanol = 1.5

    'impact du type de régulation
    impact_regulation(regulation_Shunt) = 0

    'temps de fonctionnement de l'aération
    air_tps_fonctionnement = 13.5    'h/j

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
    O2_facteur_beta = 0.99

    'correction alpha en fonction des MES  alfa=alfa*(a0+a1*MES_reel)/(a0+a1*MES_lim)
    a0_correctif_alfa_MES = 1.0007
    a1_correctif_alfa_MES = -0.035
    MES_reference_alfa = 5

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
    NumeroErreurNonGeree = "40.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub effacer_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String




    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil40.Range("ANITA_Shunt_alim_jour_semaine").Cells(1, Colonne) = ""     ' alim_jour_semaine
    Feuil40.Range("ANITA_Shunt_alim_heure_jour").Cells(1, Colonne) = ""     ' alim_heure_jour
    Feuil40.Range("ANITA_Shunt_volume_tampon").Cells(1, Colonne) = ""     ' volume_tampon
    Feuil40.Range("ANITA_Shunt_alim_nb_pompe").Cells(1, Colonne) = ""     ' alim_nb_pompe
    Feuil40.Range("ANITA_Shunt_alim_Prefoulement").Cells(1, Colonne) = ""     ' alim_P_refoulement
    Feuil40.Range("ANITA_Shunt_alim_tps_fct_pompe").Cells(1, Colonne) = ""     ' alim_tps_fct_pompe
    Feuil40.Range("ANITA_Shunt_alim_Qunitaire").Cells(1, Colonne) = ""     ' alim_Qunitaire
    Feuil40.Range("ANITA_Shunt_alim_pompe_rdt").Cells(1, Colonne) = ""     ' alim_pompe_rdt

    Feuil40.Range("ANITA_Shunt_rdt_NH4").Cells(1, Colonne) = ""     ' rdt_NH4
    Feuil40.Range("ANITA_Shunt_outlet_MES").Cells(1, Colonne) = ""     ' outlet_MES
    Feuil40.Range("ANITA_Shunt_methanol_flux").Cells(1, Colonne) = ""     ' methanol_flux
    Feuil40.Range("ANITA_Shunt_boues_concentration").Cells(1, Colonne) = ""     ' boues_concentration
    Feuil40.Range("ANITA_Shunt_boues_MES").Cells(1, Colonne) = ""     ' boues_MES
    Feuil40.Range("ANITA_Shunt_boues_Q").Cells(1, Colonne) = ""     ' boues_Q
    Feuil40.Range("ANITA_Shunt_MV_MES").Cells(1, Colonne) = ""     ' MV_MES
    Feuil40.Range("ANITA_Shunt_extraction_nb_pompe").Cells(1, Colonne) = ""     ' extraction_nb_pompe
    Feuil40.Range("ANITA_Shunt_extraction_Prefoulement").Cells(1, Colonne) = ""     ' extraction_P_refoulement
    Feuil40.Range("ANITA_Shunt_extraction_tps_fct_pompe").Cells(1, Colonne) = ""     ' extraction_tps_fct_pompe
    Feuil40.Range("ANITA_Shunt_extraction_Qunitaire").Cells(1, Colonne) = ""     ' extraction_Qunitaire
    Feuil40.Range("ANITA_Shunt_extraction_pompe_rdt").Cells(1, Colonne) = ""     ' extraction_pompe_rdt

    Feuil40.Range("ANITA_Shunt_temperature").Cells(1, Colonne) = ""     'temperature
    Feuil40.Range("ANITA_Shunt_volume_SBR").Cells(1, Colonne) = ""     'volume_SBR
    Feuil40.Range("ANITA_Shunt_O2_besoin").Cells(1, Colonne) = ""     'O2_besoin
    Feuil40.Range("ANITA_Shunt_O2_rdt_transfert").Cells(1, Colonne) = ""     'O2_rdt_transfert
    Feuil40.Range("ANITA_Shunt_O2_facteur_alpha").Cells(1, Colonne) = ""     'O2_facteur_alfa
    Feuil40.Range("ANITA_Shunt_hauteur_bassin").Cells(1, Colonne) = ""     'hauteur_bassin
    Feuil40.Range("ANITA_Shunt_O2_dissous").Cells(1, Colonne) = ""     'O2_dissous
    Feuil40.Range("ANITA_Shunt_air_Q_Nm3j").Cells(1, Colonne) = ""     'air_Q_Nm3j
    Feuil40.Range("ANITA_Shunt_air_P_refoulement").Cells(1, Colonne) = ""     'air_P_refoulement
    Feuil40.Range("ANITA_Shunt_surpresseur_conso_spec").Cells(1, Colonne) = ""     'surpresseur_conso_spec
    Feuil40.Range("ANITA_Shunt_electricite").Cells(1, Colonne) = ""     'electricite


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "40.10"

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


