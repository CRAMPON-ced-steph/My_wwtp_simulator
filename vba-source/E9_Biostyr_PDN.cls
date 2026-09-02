Option Explicit

Private Const critere_incoherence = 0.1


'choix
Dim choix_devenir_ES As Integer
Dim choix_surface_unitaire As Double   'cas particulier: donne directement la valeur choisie
Dim choix_diametre_media As Double   'cas particulier: donne directement la valeur choisie

'gammes de concentration
Private Const nb_gamme_inlet_MES = 4
Private Const MES_x1 = 15
Private Const MES_x2 = 20
Private Const MES_x3 = 25
Private Const gamme_MES_0_x1 = 1
Private Const gamme_MES_x1_x2 = 2
Private Const gamme_MES_x2_x3 = 3
Private Const gamme_MES_x3_infini = 4

'variables échangées avec excel
Dim NO3_garantie As Double
Dim NO3_visee As Double
Dim inlet_ratio_admis As Double
Dim hauteur_media As Double
Dim surface_filtration_min As Double
Dim nb_cellules As Integer
Dim surface_filtration As Double
Dim ratio_ES_volume_media As Double
Dim ES_nb_pompe As Integer
Dim ES_P_refoulement As Double
Dim ES_tps_fonctionnement As Double
Dim ES_Q_unitaire As Double
Dim ES_pompe_rdt As Double
Dim lavage_air_Q As Double
Dim lavage_air_P_refoulement As Double
Dim lavage_air_tps_fonctionnement As Double
Dim ES_concentration As Double
Dim ES_MV_MES As Double
Dim ES_MES As Double
Public methanol_flux As Double
Public electricite As Double

'variables forcées
Dim NO3_garantie_force As Boolean
Dim NO3_visee_force As Boolean
Dim inlet_ratio_admis_force As Boolean
Dim hauteur_media_force As Boolean
Dim surface_filtration_min_force As Boolean
Dim nb_cellules_force As Boolean
Dim ratio_ES_volume_media_force As Boolean
Dim ES_nb_pompe_force As Boolean
Dim ES_P_refoulement_force As Boolean
Dim ES_tps_fonctionnement_force As Boolean
Dim ES_Q_unitaire_force As Boolean
Dim ES_pompe_rdt_force As Boolean
Dim lavage_air_Q_force As Boolean
Dim lavage_air_P_refoulement_force As Boolean
Dim lavage_air_tps_fonctionnement_force As Boolean
Dim ES_concentration_force As Boolean
Dim ES_MV_MES_force As Boolean
Dim ES_MES_force As Boolean
Dim methanol_flux_force As Boolean

'variables échangées avec excel
Dim NO3_garantie_f As Double
Dim NO3_visee_f As Double
Dim inlet_ratio_admis_f As Double
Dim hauteur_media_f As Double
Dim surface_filtration_min_f As Double
Dim nb_cellules_f As Integer
Dim ratio_ES_volume_media_f As Double
Dim ES_nb_pompe_f As Integer
Dim ES_P_refoulement_f As Double
Dim ES_tps_fonctionnement_f As Double
Dim ES_Q_unitaire_f As Double
Dim ES_pompe_rdt_f As Double
Dim lavage_air_Q_f As Double
Dim lavage_air_P_refoulement_f As Double
Dim lavage_air_tps_fonctionnement_f As Double
Dim ES_concentration_f As Double
Dim ES_MV_MES_f As Double
Dim ES_MES_f As Double
Dim methanol_flux_f As Double

'hypothèses
Dim inlet_O2_mgL As Double
Dim Cv_DCO_max As Double
Dim Cv_denit_max As Double
Dim Cv_Q_max As Double
Dim correctif_T As Double
Dim T_reference As Double
Dim ratio_denit_O2dissous As Double
Dim facteur_DCO_apportee_nominal As Double
Dim facteur_DCO_apportee_reel As Double
Dim DCO_methanol As Double
Dim rdt_MES(nb_gamme_inlet_MES) As Double
Dim ES_tps_retour As Double
Dim ratio_boues_MES_elimine As Double
Dim ratio_boues_DBO_elimine As Double
Dim DBO_methanol As Double
Dim ES_Q As Double

Dim inlet_Q As Double

Dim remarque(4) As String

Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'lecture des choix internes
    choix_devenir_ES = Feuil18.Range("biostyr_PDN_choix_devenir_ES").Cells(1, pi_FiliereConsideree)
    choix_surface_unitaire = Feuil18.Range("biostyr_PDN_choix_surface_unitaire").Cells(1, pi_FiliereConsideree)
    choix_diametre_media = Feuil18.Range("biostyr_PDN_choix_diametre_media").Cells(1, pi_FiliereConsideree)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "18.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'lecture du choix de forcer des valeurs
    NO3_garantie_force = Feuil18.Range("biostyr_PDN_NO3_garantie").Cells(1, pi_FiliereConsideree)
    NO3_visee_force = Feuil18.Range("biostyr_PDN_NO3_visee").Cells(1, pi_FiliereConsideree)
    inlet_ratio_admis_force = Feuil18.Range("biostyr_PDN_inlet_ratio_admis").Cells(1, pi_FiliereConsideree)
    hauteur_media_force = Feuil18.Range("biostyr_PDN_hauteur_media").Cells(1, pi_FiliereConsideree)
    surface_filtration_min_force = Feuil18.Range("biostyr_PDN_surface_filtration_min").Cells(1, pi_FiliereConsideree)
    nb_cellules_force = Feuil18.Range("biostyr_PDN_nb_cellules").Cells(1, pi_FiliereConsideree)
    ratio_ES_volume_media_force = Feuil18.Range("biostyr_PDN_ratio_ES_volume_media").Cells(1, pi_FiliereConsideree)
    ES_nb_pompe_force = Feuil18.Range("biostyr_PDN_ES_nb_pompe").Cells(1, pi_FiliereConsideree)
    ES_P_refoulement_force = Feuil18.Range("biostyr_PDN_ES_P_refoulement").Cells(1, pi_FiliereConsideree)
    ES_tps_fonctionnement_force = Feuil18.Range("biostyr_PDN_ES_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    ES_Q_unitaire_force = Feuil18.Range("biostyr_PDN_ES_Q_unitaire").Cells(1, pi_FiliereConsideree)
    ES_pompe_rdt_force = Feuil18.Range("biostyr_PDN_ES_pompe_rdt").Cells(1, pi_FiliereConsideree)
    lavage_air_Q_force = Feuil18.Range("biostyr_PDN_lavage_air_Q").Cells(1, pi_FiliereConsideree)
    lavage_air_P_refoulement_force = Feuil18.Range("biostyr_PDN_lavage_air_P_refoulement").Cells(1, pi_FiliereConsideree)
    lavage_air_tps_fonctionnement_force = Feuil18.Range("biostyr_PDN_lavage_air_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    ES_concentration_force = Feuil18.Range("biostyr_PDN_ES_concentration").Cells(1, pi_FiliereConsideree)
    ES_MV_MES_force = Feuil18.Range("biostyr_PDN_ES_MV_MES").Cells(1, pi_FiliereConsideree)
    ES_MES_force = Feuil18.Range("biostyr_PDN_ES_MES").Cells(1, pi_FiliereConsideree)
    methanol_flux_force = Feuil18.Range("biostyr_PDN_methanol_flux").Cells(1, pi_FiliereConsideree)

    If NO3_garantie_force = True Then
        NO3_garantie_f = Feuil18.Range("biostyr_PDN_NO3_garantie").Cells(1, pi_FiliereConsideree)
    End If
    If NO3_visee_force = True Then
        NO3_visee_f = Feuil18.Range("biostyr_PDN_NO3_visee").Cells(1, pi_FiliereConsideree)
    End If
    If inlet_ratio_admis_force = True Then
        inlet_ratio_admis_f = Feuil18.Range("biostyr_PDN_inlet_ratio_admis").Cells(1, pi_FiliereConsideree)
    End If
    If hauteur_media_force = True Then
        hauteur_media_f = Feuil18.Range("biostyr_PDN_hauteur_media").Cells(1, pi_FiliereConsideree)
    End If
    If surface_filtration_min_force = True Then
        surface_filtration_min_f = Feuil18.Range("biostyr_PDN_surface_filtration_min").Cells(1, pi_FiliereConsideree)
    End If
    If nb_cellules_force = True Then
        nb_cellules_f = Feuil18.Range("biostyr_PDN_nb_cellules").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_ES_volume_media_force = True Then
        ratio_ES_volume_media_f = Feuil18.Range("biostyr_PDN_ratio_ES_volume_media").Cells(1, pi_FiliereConsideree)
    End If
    If ES_nb_pompe_force = True Then
        ES_nb_pompe_f = Feuil18.Range("biostyr_PDN_ES_nb_pompe").Cells(1, pi_FiliereConsideree)
    End If
    If ES_P_refoulement_force = True Then
        ES_P_refoulement_f = Feuil18.Range("biostyr_PDN_ES_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If ES_tps_fonctionnement_force = True Then
        ES_tps_fonctionnement_f = Feuil18.Range("biostyr_PDN_ES_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    End If
    If ES_Q_unitaire_force = True Then
        ES_Q_unitaire_f = Feuil18.Range("biostyr_PDN_ES_Q_unitaire").Cells(1, pi_FiliereConsideree)
    End If
    If ES_pompe_rdt_force = True Then
        ES_pompe_rdt_f = Feuil18.Range("biostyr_PDN_ES_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    If lavage_air_Q_force = True Then
        lavage_air_Q_f = Feuil18.Range("biostyr_PDN_lavage_air_Q").Cells(1, pi_FiliereConsideree)
    End If
    If lavage_air_P_refoulement_force = True Then
        lavage_air_P_refoulement_f = Feuil18.Range("biostyr_PDN_lavage_air_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If lavage_air_tps_fonctionnement_force = True Then
        lavage_air_tps_fonctionnement_f = Feuil18.Range("biostyr_PDN_lavage_air_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    End If
    If ES_concentration_force = True Then
        ES_concentration_f = Feuil18.Range("biostyr_PDN_ES_concentration").Cells(1, pi_FiliereConsideree)
    End If
    If ES_MV_MES_force = True Then
        ES_MV_MES_f = Feuil18.Range("biostyr_PDN_ES_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If ES_MES_force = True Then
        ES_MES_f = Feuil18.Range("biostyr_PDN_ES_MES").Cells(1, pi_FiliereConsideree)
    End If
    If methanol_flux_force = True Then
        methanol_flux_f = Feuil18.Range("biostyr_PDN_methanol_flux").Cells(1, pi_FiliereConsideree)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "18.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil18.Range("biostyr_PDN_NO3_garantie").Cells(1, Colonne) = NO3_garantie
    Feuil18.Range("biostyr_PDN_NO3_visee").Cells(1, Colonne) = NO3_visee
    Feuil18.Range("biostyr_PDN_inlet_ratio_admis").Cells(1, Colonne) = inlet_ratio_admis
    Feuil18.Range("biostyr_PDN_hauteur_media").Cells(1, Colonne) = hauteur_media
    Feuil18.Range("biostyr_PDN_surface_filtration_min").Cells(1, Colonne) = surface_filtration_min
    Feuil18.Range("biostyr_PDN_nb_cellules").Cells(1, Colonne) = nb_cellules
    Feuil18.Range("biostyr_PDN_surface_filtration").Cells(1, Colonne) = surface_filtration
    Feuil18.Range("biostyr_PDN_ratio_ES_volume_media").Cells(1, Colonne) = ratio_ES_volume_media
    Feuil18.Range("biostyr_PDN_ES_nb_pompe").Cells(1, Colonne) = ES_nb_pompe
    Feuil18.Range("biostyr_PDN_ES_P_refoulement").Cells(1, Colonne) = ES_P_refoulement
    Feuil18.Range("biostyr_PDN_ES_tps_fonctionnement").Cells(1, Colonne) = ES_tps_fonctionnement
    Feuil18.Range("biostyr_PDN_ES_Q_unitaire").Cells(1, Colonne) = ES_Q_unitaire
    Feuil18.Range("biostyr_PDN_ES_pompe_rdt").Cells(1, Colonne) = ES_pompe_rdt
    Feuil18.Range("biostyr_PDN_lavage_air_Q").Cells(1, Colonne) = lavage_air_Q
    Feuil18.Range("biostyr_PDN_lavage_air_P_refoulement").Cells(1, Colonne) = lavage_air_P_refoulement
    Feuil18.Range("biostyr_PDN_lavage_air_tps_fonctionnement").Cells(1, Colonne) = lavage_air_tps_fonctionnement
    Feuil18.Range("biostyr_PDN_ES_concentration").Cells(1, Colonne) = ES_concentration
    Feuil18.Range("biostyr_PDN_ES_MV_MES").Cells(1, Colonne) = ES_MV_MES
    Feuil18.Range("biostyr_PDN_ES_MES").Cells(1, Colonne) = ES_MES
    Feuil18.Range("biostyr_PDN_methanol_flux").Cells(1, Colonne) = methanol_flux
    Feuil18.Range("biostyr_PDN_electricite").Cells(1, Colonne) = electricite



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "18.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil18.Range("biostyr_PDN_NO3_garantie").Cells(1, Colonne) = ""       'NO3_garantie
    Feuil18.Range("biostyr_PDN_NO3_visee").Cells(1, Colonne) = ""       'NO3_visee
    Feuil18.Range("biostyr_PDN_inlet_ratio_admis").Cells(1, Colonne) = ""       'inlet_ratio_admis
    Feuil18.Range("biostyr_PDN_hauteur_media").Cells(1, Colonne) = ""       'hauteur_media
    Feuil18.Range("biostyr_PDN_surface_filtration_min").Cells(1, Colonne) = ""       'surface_filtration_min
    Feuil18.Range("biostyr_PDN_nb_cellules").Cells(1, Colonne) = ""       'nb_cellules
    Feuil18.Range("biostyr_PDN_surface_filtration").Cells(1, Colonne) = ""       'surface_filtration
    Feuil18.Range("biostyr_PDN_ratio_ES_volume_media").Cells(1, Colonne) = ""       'ratio_ES_volume_media
    Feuil18.Range("biostyr_PDN_ES_nb_pompe").Cells(1, Colonne) = ""       'ES_nb_pompe
    Feuil18.Range("biostyr_PDN_ES_P_refoulement").Cells(1, Colonne) = ""       'ES_P_refoulement
    Feuil18.Range("biostyr_PDN_ES_tps_fonctionnement").Cells(1, Colonne) = ""       'ES_tps_fonctionnement
    Feuil18.Range("biostyr_PDN_ES_Q_unitaire").Cells(1, Colonne) = ""       'ES_Q_unitaire
    Feuil18.Range("biostyr_PDN_ES_pompe_rdt").Cells(1, Colonne) = ""       'ES_pompe_rdt
    Feuil18.Range("biostyr_PDN_lavage_air_Q").Cells(1, Colonne) = ""       'lavage_air_Q
    Feuil18.Range("biostyr_PDN_lavage_air_P_refoulement").Cells(1, Colonne) = ""       'lavage_air_P_refoulement
    Feuil18.Range("biostyr_PDN_lavage_air_tps_fonctionnement").Cells(1, Colonne) = ""       'lavage_air_tps_fonctionnement
    Feuil18.Range("biostyr_PDN_ES_concentration").Cells(1, Colonne) = ""       'ES_concentration
    Feuil18.Range("biostyr_PDN_ES_MV_MES").Cells(1, Colonne) = ""       'ES_MV_MES
    Feuil18.Range("biostyr_PDN_ES_MES").Cells(1, Colonne) = ""       'ES_MES
    Feuil18.Range("biostyr_PDN_methanol_flux").Cells(1, Colonne) = ""       'methanol_flux
    Feuil18.Range("biostyr_PDN_electricite").Cells(1, Colonne) = ""       'electricite


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "18.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'NO3 garantie en sortie du biostyr PDN     en mg/L  A VOIR
    If NO3_garantie_force = True Then
        NO3_garantie = NO3_garantie_f
    Else
        NO3_garantie = 2
    End If

    'hauteur de média    en m
    If hauteur_media_force = True Then
        hauteur_media = hauteur_media_f
    Else
        hauteur_media = 2.5
    End If

    'eaux sales
    If ratio_ES_volume_media_force = True Then
        ratio_ES_volume_media = ratio_ES_volume_media_f
    Else
        ratio_ES_volume_media = 2.5   '2.5*volume_media
    End If
    If ES_nb_pompe_force = True Then
        ES_nb_pompe = ES_nb_pompe_f
    Else
        ES_nb_pompe = 2
    End If
    If ES_P_refoulement_force = True Then
        ES_P_refoulement = ES_P_refoulement_f
    Else
        ES_P_refoulement = 10   'mCE
    End If
    If ES_tps_fonctionnement_force = True Then
        ES_tps_fonctionnement = ES_tps_fonctionnement_f
    Else
        ES_tps_fonctionnement = 18   'h/j
    End If

    'lavage des filtres (injection d'air)
    If lavage_air_P_refoulement_force = True Then
        lavage_air_P_refoulement = lavage_air_P_refoulement_f
    Else
        lavage_air_P_refoulement = 7    'mCE
    End If
    If lavage_air_tps_fonctionnement_force = True Then
        lavage_air_tps_fonctionnement = lavage_air_tps_fonctionnement_f
    Else
        lavage_air_tps_fonctionnement = 0.1    'h/j  6 min par lavage donc 0.1h/j sur l'hypothèse d'un lavage
    End If
    If lavage_air_Q_force = True Then
        lavage_air_Q = lavage_air_Q_f
    Else
        lavage_air_Q = 12 * choix_surface_unitaire
    End If

    'qualité des eaux sales
    'If ES_concentration_force = True Then   'ATTENTION, concentration des eaux sales en mg/L
    '    ES_concentration = ES_concentration_f
    'Else
    '    ES_concentration = 400   'mg/L
    'End If
    If ES_MV_MES_force = True And choix_devenir_ES <> devenir_ES_primaire Then
        ES_MV_MES = ES_MV_MES_f
    Else
        ES_MV_MES = 0.9  ' 90% MV/MES    'DCA modifié après tests experts 07/11
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "18.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    'O2 dissous en entrée de PDN    en mg/L
    inlet_O2_mgL = 6
    ratio_denit_O2dissous = 0.35    'idem à l'HybAS
    If pd_TemperatureEauDimensionnement_°C < 5 Then
        facteur_DCO_apportee_nominal = 5
    ElseIf pd_TemperatureEauDimensionnement_°C < 15 Then
        facteur_DCO_apportee_nominal = 4 + (pd_TemperatureEauDimensionnement_°C - 5) * 0.1
    Else
        facteur_DCO_apportee_nominal = 4
    End If
    If pd_TemperatureEauConditionsReelles_°C < 5 Then
        facteur_DCO_apportee_reel = 5
    ElseIf pd_TemperatureEauConditionsReelles_°C < 15 Then
        facteur_DCO_apportee_reel = 4 + (pd_TemperatureEauDimensionnement_°C - 5) * 0.1
    Else
        facteur_DCO_apportee_reel = 4
    End If
    DCO_methanol = 1.5        ' g DCO / g methanol
    DBO_methanol = 0.94    ' g DBO5 / g methanol

    'calcul des charges appliquées limites (en kg/(m3.j) pour DCO et denit; en m/h pour Q)
    If choix_diametre_media = 4.5 Then
        Cv_DCO_max = 10
        Cv_denit_max = 1.7
        Cv_Q_max = 20
    ElseIf choix_diametre_media = 5 Then
        Cv_DCO_max = 12
        Cv_denit_max = 1.5
        Cv_Q_max = 25
    End If
    'coefficients pour calcul correction en température des charges volumiques limites   correctif_T^(T-T_reference)
    'Facteur correctif température pour cinétique
    correctif_T = 1.07
    'Température de référence pour la cinétique biologique (°C)
    T_reference = 12


    'performances sur la filtration des MES (concentration fixée par granulométrie de média)   A VOIR
    'on donne des rendements qui dépendent de la concentration en entrée (et donc du média sur le biostyr NDN) et du choix de média dans le PDN
    If choix_diametre_media = 4.5 Then
        rdt_MES(gamme_MES_0_x1) = 0
        rdt_MES(gamme_MES_x1_x2) = 0.1
        rdt_MES(gamme_MES_x2_x3) = 0.2
        rdt_MES(gamme_MES_x3_infini) = 0.25
    Else
        rdt_MES(gamme_MES_0_x1) = 0
        rdt_MES(gamme_MES_x1_x2) = 0
        rdt_MES(gamme_MES_x2_x3) = 0.1
        rdt_MES(gamme_MES_x3_infini) = 0.13
    End If

    'calcul de la production de boues
    ratio_boues_MES_elimine = 1
    ratio_boues_DBO_elimine = 0.4

    'élimination de P
    'ratio_P_synthese = 0.01

    'temps de retour des eaux sales
    ES_tps_retour = 18    'h/j




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "18.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub dimensionnement(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef DCO_conso_bio, ByRef eaux_sales, ByRef choix_eaux_sales, ByVal iteration_ES, ByRef DCO_apportee_methanol)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim inlet_NO3_mgL As Double
    Dim facteur_correction_T As Double
    Dim surface_necessaire As Double
    Dim surface_retenue As Double
    Dim Cv_DCO_max_T As Double
    Dim Cv_denit_max_T As Double
    Dim Q_considere As Double
    Dim denitrification As Double
    Dim facteur_repartition As Double
    Dim inlet_MES_mgL As Double
    Dim sortie_MES_mgL As Double
    Dim gamme_concentration_MES_inlet As Double
    Dim correction_denit_O2dissous As Double
    Dim MES_elimine As Double

    'CALCUL DU FLUX TRAITE
    'concentration maximale en NO3 pour respecter la garantie en NGL
    correction_denit_O2dissous = ratio_denit_O2dissous * inlet_O2_mgL
    'NO3_visee = NGL_garantie - (1000 * NK / Q - correction_denit_O2dissous)
    NO3_visee = NGL_garantie - (1000 * NK / Q)  '- correction_denit_O2dissous)
    'calcul du % du flux à admettre pour atteindre cette concentration en NO3
    inlet_NO3_mgL = 1000 * NO3 / Q + correction_denit_O2dissous
    inlet_ratio_admis = (NO3_visee * (Q - ES_Q) - inlet_NO3_mgL * Q + NO3_garantie * ES_Q) / (Q * (NO3_garantie - inlet_NO3_mgL))
    If inlet_ratio_admis > 1 Then
        inlet_ratio_admis = 1
        NO3_visee = NO3_garantie
    ElseIf inlet_ratio_admis < 0 Then
        inlet_ratio_admis = 0.01
        NO3_visee = inlet_NO3_mgL
    End If
    inlet_Q = inlet_ratio_admis * Q

    'CALCUL DE LA SURFACE DE FILTRATION
    If surface_filtration_min_force = True Then
        surface_filtration_min = surface_filtration_min_f
    Else
        'correction par l'effet de la température des charges appliquées maximales sur les biofiltres (pollutions)
        facteur_correction_T = correctif_T ^ (pd_TemperatureEauDimensionnement_°C - T_reference)
        Cv_DCO_max_T = Cv_DCO_max * facteur_correction_T
        Cv_denit_max_T = Cv_denit_max * facteur_correction_T
        'calcul de la surface nécessaire
        'DCO
        surface_necessaire = (inlet_ratio_admis * DCO) / (Cv_DCO_max_T * hauteur_media)
        surface_retenue = surface_necessaire
        'NO3
        surface_necessaire = (NO3 * inlet_ratio_admis) / (Cv_denit_max_T * hauteur_media)
        If surface_necessaire > surface_retenue Then
            surface_retenue = surface_necessaire
        End If
        'hydraulique   (pour le PDN, on se rapporte à la totalité des cellules et non uniquement à celles en marche (i.e. hors cellules en lavage)
        Q_considere = inlet_ratio_admis * (Q_nominal * pd_CoefficientPointeHydrauliqueTempsPluie / pd_NOMBRE_HEURE_PAR_JOUR + (Q - Q_nominal) / ES_tps_retour)    'Q_ptp+Q_retour  (Q_retour lissé sur 24h)
        surface_necessaire = Q_considere / Cv_Q_max
        If surface_necessaire > surface_retenue Then
            surface_retenue = surface_necessaire
        End If
        surface_filtration_min = surface_retenue
    End If
    If nb_cellules_force = True Then
        nb_cellules = nb_cellules_f
    Else
        nb_cellules = surface_filtration_min / choix_surface_unitaire
        If nb_cellules < surface_filtration_min / choix_surface_unitaire Then
            nb_cellules = nb_cellules + 1
        End If
    End If
    surface_filtration = nb_cellules * choix_surface_unitaire


    'CALCUL VOLUME EAUX SALES ET DE LA CONCENTRATION EN MES
    'ré-initialisation des eaux sales
    For I = 1 To nb_parametres_eaux_sales
        eaux_sales(repere_ES_biostyr_PDN, I) = 0
    Next I
    choix_eaux_sales(repere_ES_biostyr_PDN) = choix_devenir_ES
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_origine) = II_biostyr_PDN
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) = ratio_ES_volume_media * (surface_filtration * hauteur_media)
    ES_Q = eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q)
    'eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) / 1000
    'ES_MES = eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES)
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_MV_MES) = ES_MV_MES
    'NO3
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_NO3) = NO3_garantie * eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) / 1000
    'autres pollutions: on considère une répartition homogène en concentration
    facteur_repartition = eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) / Q
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_DCO) = DCO * facteur_repartition
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_DBO) = DBO * facteur_repartition
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_NK) = NK * facteur_repartition
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_NH4) = NH4 * facteur_repartition
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_Pt) = Pt * facteur_repartition

    'CALCUL DU METHANOL AJOUTE
    'part NO3
    denitrification = inlet_ratio_admis * NO3 - eaux_sales(repere_ES_biostyr_PDN, repere_ES_NO3) - NO3_garantie * (inlet_Q - eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q)) / 1000
    'part O2 dissous
    denitrification = denitrification + ratio_denit_O2dissous * inlet_O2_mgL * inlet_Q / 1000
    'NH4
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_NH4) = eaux_sales(repere_ES_biostyr_PDN, repere_ES_NH4)    '- facteur_repartition * ratio_denit_O2dissous * inlet_O2_mgL * inlet_Q / 1000
    'NH4 = NH4 - ratio_denit_O2dissous * inlet_O2_mgL * inlet_Q / 1000
    'NK
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_NK) = eaux_sales(repere_ES_biostyr_PDN, repere_ES_NK)    '- facteur_repartition * ratio_denit_O2dissous * inlet_O2_mgL * inlet_Q / 1000
    'NK = NK - ratio_denit_O2dissous * inlet_O2_mgL * inlet_Q / 1000
    methanol_flux = facteur_DCO_apportee_nominal * denitrification / DCO_methanol
    DCO_apportee_methanol = DCO_methanol * methanol_flux
    DCO_conso_bio = DCO_conso_bio + DCO_apportee_methanol

    'CALCUL DU RENDEMENT DE FILTRATION  (rendement exprimé par rapport aux concentrations en MES en entrée et en sortie du PDN)
    inlet_MES_mgL = MES / Q * 1000   'concentration en entrée
    If inlet_MES_mgL < MES_x1 Then
        gamme_concentration_MES_inlet = gamme_MES_0_x1
    ElseIf inlet_MES_mgL < MES_x2 Then
        gamme_concentration_MES_inlet = gamme_MES_x1_x2
    ElseIf inlet_MES_mgL < MES_x3 Then
        gamme_concentration_MES_inlet = gamme_MES_x2_x3
    Else
        gamme_concentration_MES_inlet = gamme_MES_x3_infini
    End If
    sortie_MES_mgL = inlet_MES_mgL * (1 - rdt_MES(gamme_concentration_MES_inlet))

    'calcul de la concentration des boues
    If type_eau_nominal = eau_diluee Then
        ratio_boues_DBO_elimine = 0.2
    ElseIf type_eau_nominal = eau_standard Then
        ratio_boues_DBO_elimine = 0.4
    Else
        ratio_boues_DBO_elimine = 0.4
    End If
    MES_elimine = inlet_ratio_admis * (inlet_MES_mgL - sortie_MES_mgL) * Q / 1000
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES) = ratio_boues_MES_elimine * MES_elimine + ratio_boues_DBO_elimine * DBO_methanol / DCO_methanol * DCO_apportee_methanol
    ES_concentration = eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES) / eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) * 1000
    GoTo sauter_boucle
    '************ A VOIR SI AU DESSUS CA ESTIME BIEN POUR ENLEVER CE QUI SUIT *************
    If ES_concentration_force = True Then
        ES_concentration = ES_concentration_f
        eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) / 1000
    Else
        If type_eau_nominal = eau_diluee Then
            ES_concentration = 400
        ElseIf type_eau_nominal = eau_standard Then
            ES_concentration = 600
        Else
            ES_concentration = 800
        End If
        eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) / 1000
    End If
sauter_boucle:

    ES_MES = eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES)


    'CHARGES EN SORTIE
    'attention, certaines valeurs seront remodifiées plus loin car dépendent encore des hypothèses de composition des boues comme les NK et le Pt
    Q = Q - eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q)
    DCO = DCO - eaux_sales(repere_ES_biostyr_PDN, repere_ES_DCO)
    DBO = DBO - eaux_sales(repere_ES_biostyr_PDN, repere_ES_DBO)
    MES = (1 - inlet_ratio_admis) * MES + sortie_MES_mgL * (inlet_Q - eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q)) / 1000
    NH4 = NH4 - eaux_sales(repere_ES_biostyr_PDN, repere_ES_NH4)
    NK = NK - eaux_sales(repere_ES_biostyr_PDN, repere_ES_NK)
    NO3 = NO3_visee * Q / 1000
    Pt = Pt - eaux_sales(repere_ES_biostyr_PDN, repere_ES_Pt)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "18.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub



Sub fonctionnement_reel(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef DCO_conso_bio, ByRef eaux_sales, ByRef choix_eaux_sales, ByVal iteration_ES, ByRef DCO_apportee_methanol)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim inlet_NO3_mgL As Double
    Dim facteur_correction_T As Double
    Dim surface_necessaire As Double
    Dim surface_retenue As Double
    Dim Cv_DCO_max_T As Double
    Dim Cv_denit_max_T As Double
    Dim Q_considere As Double
    Dim denitrification As Double
    Dim facteur_repartition As Double
    Dim inlet_MES_mgL As Double
    Dim sortie_MES_mgL As Double
    Dim gamme_concentration_MES_inlet As Double
    Dim correction_denit_O2dissous As Double
    Dim MES_elimine As Double


    'CALCUL DU FLUX TRAITE
    'concentration maximale en NO3 pour respecter la garantie en NGL
    If NO3_visee_force = True Then
        NO3_visee = NO3_visee_f
    Else
        correction_denit_O2dissous = ratio_denit_O2dissous * inlet_O2_mgL
        'NO3_visee = NGL_garantie - (1000 * NK / Q - correction_denit_O2dissous)
        NO3_visee = NGL_garantie - (1000 * NK / Q)  '- correction_denit_O2dissous)
    End If
    If NO3_visee < NO3_garantie Then
        remarque(1) = Feuil18.Range("biostyr_PDN_erreur_NO3").Cells(1, 1)
    End If
    'calcul du % du flux à admettre pour atteindre cette concentration en NO3
    'inlet_NO3_mgL = 1000 * NO3 / Q
    inlet_NO3_mgL = 1000 * NO3 / Q + correction_denit_O2dissous
    If inlet_ratio_admis_force = True Then
        inlet_ratio_admis = inlet_ratio_admis_f
        If inlet_ratio_admis > 1 Or inlet_ratio_admis < 0 Then
            remarque(2) = Feuil18.Range("biostyr_PDN_erreur_bypass").Cells(1, 1)
        End If
    Else
        'inlet_ratio_admis = (NO3_visee - inlet_NO3_mgL) / (NO3_garantie - inlet_NO3_mgL)
        inlet_ratio_admis = (NO3_visee * (Q - ES_Q) - inlet_NO3_mgL * Q + NO3_garantie * ES_Q) / (Q * (NO3_garantie - inlet_NO3_mgL))
        If inlet_ratio_admis > 1 Then
            inlet_ratio_admis = 1
            NO3_visee = NO3_garantie
        ElseIf inlet_ratio_admis < 0 Then
            inlet_ratio_admis = 0.01
            NO3_visee = inlet_NO3_mgL
        End If
    End If
    inlet_Q = inlet_ratio_admis * Q


    'CALCUL VOLUME EAUX SALES ET DE LA CONCENTRATION EN MES
    'ré-initialisation des eaux sales
    For I = 1 To nb_parametres_eaux_sales
        eaux_sales(repere_ES_biostyr_PDN, I) = 0
    Next I
    choix_eaux_sales(repere_ES_biostyr_PDN) = choix_devenir_ES
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_origine) = II_biostyr_PDN
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) = ratio_ES_volume_media * (surface_filtration * hauteur_media)
    ES_Q = eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q)
    'eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) / 1000
    'ES_MES = eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES)
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_MV_MES) = ES_MV_MES
    'NO3
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_NO3) = NO3_garantie * eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) / 1000
    'autres pollutions: on considère une répartition homogène en concentration (hors NH4 et NO3 qui dépendent de l'O2 dissous)
    facteur_repartition = eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) / Q
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_DCO) = DCO * facteur_repartition
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_DBO) = DBO * facteur_repartition
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_NK) = NK * facteur_repartition
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_NH4) = NH4 * facteur_repartition
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_Pt) = Pt * facteur_repartition

    'CALCUL DU METHANOL AJOUTE
    'part NO3
    denitrification = inlet_ratio_admis * NO3 - eaux_sales(repere_ES_biostyr_PDN, repere_ES_NO3) - NO3_garantie * (inlet_Q - eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q)) / 1000
    'part O2 dissous
    denitrification = denitrification + ratio_denit_O2dissous * inlet_O2_mgL * inlet_Q / 1000
    'NH4
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_NH4) = eaux_sales(repere_ES_biostyr_PDN, repere_ES_NH4)    '- facteur_repartition * ratio_denit_O2dissous * inlet_O2_mgL * inlet_Q / 1000
    'NH4 = NH4 - ratio_denit_O2dissous * inlet_O2_mgL * inlet_Q / 1000
    'NK
    eaux_sales(repere_ES_biostyr_PDN, repere_ES_NK) = eaux_sales(repere_ES_biostyr_PDN, repere_ES_NK)    '- facteur_repartition * ratio_denit_O2dissous * inlet_O2_mgL * inlet_Q / 1000
    'NK = NK - ratio_denit_O2dissous * inlet_O2_mgL * inlet_Q / 1000
    If methanol_flux_force = True Then
        methanol_flux = methanol_flux_f
    Else
        methanol_flux = facteur_DCO_apportee_nominal * denitrification / DCO_methanol
    End If
    DCO_apportee_methanol = DCO_methanol * methanol_flux
    DCO_conso_bio = DCO_conso_bio + DCO_apportee_methanol


    'CALCUL DU RENDEMENT DE FILTRATION  (rendement exprimé par rapport aux concentrations en MES en entrée et en sortie du PDN)
    inlet_MES_mgL = MES / Q * 1000   'concentration en entrée
    If inlet_MES_mgL < MES_x1 Then
        gamme_concentration_MES_inlet = gamme_MES_0_x1
    ElseIf inlet_MES_mgL < MES_x2 Then
        gamme_concentration_MES_inlet = gamme_MES_x1_x2
    ElseIf inlet_MES_mgL < MES_x3 Then
        gamme_concentration_MES_inlet = gamme_MES_x2_x3
    Else
        gamme_concentration_MES_inlet = gamme_MES_x3_infini
    End If
    sortie_MES_mgL = inlet_MES_mgL * (1 - rdt_MES(gamme_concentration_MES_inlet))

    'calcul de la concentration des boues
    If type_eau_reel = eau_diluee Then
        ratio_boues_DBO_elimine = 0.2
    ElseIf type_eau_reel = eau_standard Then
        ratio_boues_DBO_elimine = 0.4
    Else
        ratio_boues_DBO_elimine = 0.4
    End If
    If ES_concentration_force = True Then
        ES_concentration = ES_concentration_f
        eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) / 1000
    Else
        MES_elimine = inlet_ratio_admis * (inlet_MES_mgL - sortie_MES_mgL) * Q / 1000
        eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES) = ratio_boues_MES_elimine * MES_elimine + ratio_boues_DBO_elimine * DBO_methanol / DCO_methanol * DCO_apportee_methanol
        ES_concentration = eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES) / eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) * 1000
    End If
    GoTo sauter_boucle
    '************ A VOIR SI AU DESSUS CA ESTIME BIEN POUR ENLEVER CE QUI SUIT *************
    If ES_concentration_force = True Then
        ES_concentration = ES_concentration_f
        eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) / 1000
    Else
        If type_eau_reel = eau_diluee Then
            ES_concentration = 400
        ElseIf type_eau_reel = eau_standard Then
            ES_concentration = 600
        Else
            ES_concentration = 800
        End If
        eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q) / 1000
    End If
sauter_boucle:

    ES_MES = eaux_sales(repere_ES_biostyr_PDN, repere_ES_MES)


    'CHARGES EN SORTIE
    'attention, certaines valeurs seront remodifiées plus loin car dépendent encore des hypothèses de composition des boues comme les NK et le Pt
    Q = Q - eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q)
    DCO = DCO - eaux_sales(repere_ES_biostyr_PDN, repere_ES_DCO)
    DBO = DBO - eaux_sales(repere_ES_biostyr_PDN, repere_ES_DBO)
    MES = (1 - inlet_ratio_admis) * MES + sortie_MES_mgL * (inlet_Q - eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q)) / 1000
    NH4 = NH4 - eaux_sales(repere_ES_biostyr_PDN, repere_ES_NH4)
    NK = NK - eaux_sales(repere_ES_biostyr_PDN, repere_ES_NK)
    NO3 = NO3_visee * Q / 1000
    Pt = Pt - eaux_sales(repere_ES_biostyr_PDN, repere_ES_Pt)

    'stockage volume eaux sales
    ES_Q = eaux_sales(repere_ES_biostyr_PDN, repere_ES_Q)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "18.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub

Sub calcul_consommation_electrique()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'variables pour l'aération
    Dim surpresseur_conso_Wh_Nm3mCE As Double
    Dim electricite_lavage_air As Double

    'recirculation
    Dim rendement_pompe As Double
    Dim puissance_pompe As Double
    Dim rendement_moteur As Double
    Dim electricite_ES As Double

    'méthanol
    Dim ratio_elec_methanol_pompe As Double
    Dim electricite_methanol As Double

    Dim facteur_conso_additionnelle As Double

    'erreurs
    Dim incoherence As Double


    electricite = 0





    'AERATION LAVAGE
    surpresseur_conso_Wh_Nm3mCE = 4.5     'Wh/(Nm3.mCE)  valeur médiane
    electricite_lavage_air = lavage_air_Q * lavage_air_tps_fonctionnement * nb_cellules * lavage_air_P_refoulement * surpresseur_conso_Wh_Nm3mCE / 1000
    electricite = electricite + electricite_lavage_air


    'POMPAGE DES EAUX SALES DU BIOSTYR VERS LE TRAITEMENT SPECIFIQUE
    'ES_Q_unitaire = ES_Q / (ES_nb_pompe * ES_tps_fonctionnement)
    If ES_Q_unitaire_force = True Then
        ES_Q_unitaire = ES_Q_unitaire_f
        If ES_tps_fonctionnement_force = True Then
            'on ne fait rien
            If ES_nb_pompe_force = False Then
                If Int(ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire)) < (ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire)) Then
                    ES_nb_pompe = Int(ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire)) + 1
                Else
                    ES_nb_pompe = Int(ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire))
                End If
            End If
        Else
            ES_tps_fonctionnement = ES_Q / (ES_nb_pompe * ES_Q_unitaire)
            If ES_tps_fonctionnement > ES_tps_retour Then
                ES_tps_fonctionnement = ES_tps_retour
                If ES_nb_pompe_force = False Then
                    If Int(ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire)) < (ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire)) Then
                        ES_nb_pompe = Int(ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire)) + 1
                    Else
                        ES_nb_pompe = Int(ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire))
                    End If
                    ES_tps_fonctionnement = ES_Q / (ES_nb_pompe * ES_Q_unitaire)
                End If
            End If
        End If
    Else
        ES_Q_unitaire = ES_Q / (ES_nb_pompe * ES_tps_fonctionnement)
    End If
    incoherence = (ES_Q_unitaire * ES_nb_pompe * ES_tps_fonctionnement - ES_Q) / (ES_Q)
    If Abs(incoherence) > critere_incoherence Then
        remarque(3) = Feuil18.Range("biostyr_PDN_erreur_pompe_ES").Cells(1, 1)
    End If
    If ES_pompe_rdt_force = True Then
        ES_pompe_rdt = ES_pompe_rdt_f
    Else
        rendement_pompe = 0.7    'hypothèse
        puissance_pompe = ES_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * ES_P_refoulement / rendement_pompe
        If puissance_pompe < 60 Then
            rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
        Else
            rendement_moteur = 0.961
        End If
        ES_pompe_rdt = rendement_pompe * rendement_moteur
    End If
    electricite_ES = ES_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * ES_P_refoulement / ES_pompe_rdt * ES_nb_pompe * ES_tps_fonctionnement
    electricite = electricite + electricite_ES

    'POMPES METHANOL     A REVOIR
    ratio_elec_methanol_pompe = 1.16   'Wh/m3 d'eau entrante  (0.73 pour Angers, 1.62 pour Nantes et 1.12 pour Douarnenez)
    electricite_methanol = ratio_elec_methanol_pompe * inlet_Q / 1000
    electricite = electricite + electricite_methanol

    'CONSOMMATION ADDITIONNELLE
    facteur_conso_additionnelle = 0.02  'vannes ...
    electricite = electricite * (1 + facteur_conso_additionnelle)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "18.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub



Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 4  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil18.Range("biostyr_PDN_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "18.8"

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





