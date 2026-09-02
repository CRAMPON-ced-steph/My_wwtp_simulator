Attribute VB_Name = "z_CLS_RetoursMAP"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'Gestion des retours
Dim mcls_RetoursAdmis As CLS_RetoursAdmisTraitement

'Type de process
Private Const mi_OSTARA As Integer = 1
Private Const mi_TURBOFIX_MAP As Integer = 2

'Nombre d'erreurs
Private Const mi_NbErreurs As Integer = 9
Private Const mi_ErrInletpH As Integer = 1
Private Const mi_ErrInletPPO4Faible As Integer = 2
Private Const mi_ErrInletMESTrop As Integer = 3
Private Const mi_ErrOutletPPO4Faible = 4
Private Const mi_ErrRatioMolaireNH4PO4Faible As Integer = 5
Private Const mi_ErrRatioMolaireCaPO4Trop As Integer = 6
Private Const mi_ErrRatioMolaireMgPFaible As Integer = 7
Private Const mi_ErrRatioMolaireMgPTrop As Integer = 8
Private Const mi_ErrOstaraProductionStruviteFaible As Integer = 9

'Choix de l'utilisateur
Dim mi_ChoixProcess As Integer
'voir pour le choix des retours d'Athos ou non

'Valeurs renvoyées à l'utilisateur
Dim md_InletMg_mgL As Double
Dim md_InletCa_mgL As Double
Dim md_NominalQ_m3j As Double
Dim md_NominalPPO4_mgL As Double
Dim md_NominalNNH4_mgL As Double
Dim md_NominalStruviteProduite_kgj As Double
Dim ms_ReacteursInstalles As String
Dim md_InletpH As Double
Public pd_NaOHpur_kgj As Double
Dim md_InletTemperature_°C As Double
Dim md_RatioMolaireMgP As Double
Dim md_RendmentPO4 As Double
Dim md_OutletPPO4_mgL As Double
Dim md_OutletNNH4_mgL As Double    'MSM 15/11/12
Public pd_MgCl2pur_kgj As Double
Public pd_StruviteProduite_kgj As Double
Public pd_electricite_kWhj As Double

'Repère si valeur forcée par l'utilisateur
Dim mn_InletMg_mgL_force As Boolean
Dim mn_InletCa_mgL_force As Boolean
Dim mn_NominalQ_m3j_force As Boolean
Dim mn_NominalPPO4_mgL_force As Boolean
Dim mn_NominalNNH4_mgL_force As Boolean
Dim mn_InletpH_force As Boolean
Dim mn_NaOHpur_kgj_force As Boolean
Dim mn_InletTemperature_°C_force As Boolean
Dim mn_RatioMolaireMgP_force As Boolean
Dim mn_RendmentPO4_force As Boolean
'Dim mn_OutletPPO4_mgL_force As Boolean 'MSM 15/11/12

'Valeurs forcées par l'utilisateur
Dim md_InletMg_mgL_f As Double
Dim md_InletCa_mgL_f As Double
Dim md_NominalQ_m3j_f As Double
Dim md_NominalPPO4_mgL_f As Double
Dim md_NominalNNH4_mgL_f As Double
Dim md_InletpH_f As Double
Dim md_NaOHpur_kgj_f As Double
Dim md_InletTemperature_°C_f As Double
Dim md_RatioMolaireMgP_f As Double
Dim md_RendmentPO4_f As Double
'Dim md_OutletPPO4_mgL_f As Double 'MSM 15/11/12

'Hypothèses
Dim md_MM_MgCl2_kg_mol As Double
Dim md_MM_MgNH4PO4H12O6_kg_mol As Double
Dim md_HypotheseRendmentPO4 As Double
Dim md_OutletPPO4Mini_mgL As Double
Dim md_ProductionMinimaleStruvite_kgj As Double
Dim md_InletPPO4Mini_mgL As Double
Dim md_InletMESMaxi_mgL As Double
Dim md_RatioMolaireMiniNNH4PPO4 As Double
Dim md_RatioMolaireMaxiCaPPO4 As Double
Dim md_ProductionMaximalePEARL500_kgj As Double
Dim md_ProductionMaximalePEARL2000_kgj As Double
Dim md_HypRatioMolaireMgP As Double
Dim md_HypElectricitePEARL500_kWh_kg As Double
Dim md_HypElectricitePEARL2000_kWh_kg As Double

'messages d'erreurs
Dim ms_Remarque(mi_NbErreurs) As String




Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'lecture des choix internes
    mi_ChoixProcess = Feuil70.Range("MAP_retours_choix_procede").Cells(1, pi_FiliereConsideree)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "70.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'masses molaires   en g/mol
    md_MM_MgCl2_kg_mol = pd_MASSE_MOLAIRE_Mg_kg_mol + 2 * pd_MASSE_MOLAIRE_Cl_kg_mol   'masse molaire du réactif
    md_MM_MgNH4PO4H12O6_kg_mol = pd_MASSE_MOLAIRE_Mg_kg_mol + pd_MASSE_MOLAIRE_N_kg_mol + 4 * pd_MASSE_MOLAIRE_H_kg_mol + pd_MASSE_MOLAIRE_P_kg_mol + 4 * pd_MASSE_MOLAIRE_O_kg_mol + 12 * pd_MASSE_MOLAIRE_H_kg_mol + 6 * pd_MASSE_MOLAIRE_O_kg_mol    'masse molaire de la struvite 6 fois hydratée

    'Rendement théorique sur le PO4 et concentration minimale atteignable      source: Christophe MELE
    md_HypotheseRendmentPO4 = 0.8
    md_OutletPPO4Mini_mgL = 10

    If mi_ChoixProcess = mi_OSTARA Then
        'Production maximale pour installer le PEARL 500
        md_ProductionMaximalePEARL500_kgj = 500
        'Production maximale pour installer le PEARL 500
        md_ProductionMaximalePEARL2000_kgj = 2000
        'Production minimale pour installer le réacteur PEARL 500 (le plus petit)
        md_ProductionMinimaleStruvite_kgj = 0.6 * md_ProductionMaximalePEARL500_kgj
    End If

    'Concentration minimale en entrée en PPO4 (mg/L)      source: Christophe MELE
    md_InletPPO4Mini_mgL = 70

    'Concentration maximale en entrée en MES (mg/L)      source: Christophe MELE
    md_InletMESMaxi_mgL = 2000

    'Ratio molaire NNH4/PPO4 minimal pour faire la struvite      source: Christophe MELE
    md_RatioMolaireMiniNNH4PPO4 = 2

    'Ratio molaire maxi en Ca/PPO4      source: Christophe MELE
    md_RatioMolaireMaxiCaPPO4 = 1

    'Ratio molaire conseillé en Mg/P     entre 1 et 1,5  source: Christophe MELE
    md_HypRatioMolaireMgP = 1.2

    'Consommation électrique spécifique (kWh/kg de struvite) source Christophe MELE
    md_HypElectricitePEARL500_kWh_kg = 0.7      'Chartres (1*500, production 441 kg/j): 0.68 kWh/kg de struvite produite / Perpignan (1*500, production 208 kg/j): 0.77 kWh/kg de struvite produite
    md_HypElectricitePEARL2000_kWh_kg = 0.35     'Bruxelles (1*2000, production 1851 kg/j): 0.35 kWh/kg de struvite produite


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "70.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub calcul_cristallisation(ByRef Q_m3j As Double, ByRef PPO4_kgj As Double, ByRef Pt_kgj As Double, ByRef NNH4_kgj As Double, ByRef NK_kgj As Double, ByRef StruviteProduite_kgj As Double)

    Dim MgCl2_molj As Double
    Dim PPO4precipite_kgj As Double
    Dim PO4precipite_molj As Double
    Dim NNH4precipite_kgj As Double
    Dim PPO4Inlet_mgL As Double
    Dim PPO4Outlet_mgL As Double

    PPO4Inlet_mgL = PPO4_kgj / Q_m3j * 1000

    'Vérification des ratios molaires
    If ((NNH4_kgj / pd_MASSE_MOLAIRE_N_kg_mol) / (PPO4_kgj / pd_MASSE_MOLAIRE_P_kg_mol)) < md_RatioMolaireMiniNNH4PPO4 Then
        '********************************* message: Attention, carence en NH4 **********************************
        ms_Remarque(mi_ErrRatioMolaireNH4PO4Faible) = Feuil70.Range("MAP_retours_ErrRatioMolaireNH4PO4Faible").Cells(1, 1)
    Else
        ms_Remarque(mi_ErrRatioMolaireNH4PO4Faible) = ""     'Feuil70.Range("MAP_retours_ErrRatioMolaireNH4PO4Faible").Cells(1, 1)
    End If
    If ((md_InletCa_mgL / pd_MASSE_MOLAIRE_Ca_kg_mol) / (PPO4Inlet_mgL / pd_MASSE_MOLAIRE_P_kg_mol)) > md_RatioMolaireMaxiCaPPO4 Then
        '********************************* message: Attention, risque de précipitation de phosphates de calcium **********************************
        ms_Remarque(mi_ErrRatioMolaireCaPO4Trop) = Feuil70.Range("MAP_retours_ErrRatioMolaireCaPO4Trop").Cells(1, 1)
    Else
        ms_Remarque(mi_ErrRatioMolaireCaPO4Trop) = ""     'Feuil70.Range("MAP_retours_ErrRatioMolaireCaPO4Trop").Cells(1, 1)
    End If

    'Calcul du rendement
    If mn_RendmentPO4_force = True Then
        md_RendmentPO4 = md_RendmentPO4_f
        PPO4Outlet_mgL = PPO4Inlet_mgL * (1 - md_RendmentPO4)
    Else
        md_RendmentPO4 = md_HypotheseRendmentPO4
        PPO4Outlet_mgL = PPO4Inlet_mgL * (1 - md_RendmentPO4)
        If PPO4Outlet_mgL < md_OutletPPO4Mini_mgL Then
            PPO4Outlet_mgL = md_OutletPPO4Mini_mgL
            md_RendmentPO4 = (PPO4Inlet_mgL - PPO4Outlet_mgL) / PPO4Inlet_mgL
        End If
    End If

    'Calcul du PPO4 précipité
    PPO4precipite_kgj = PPO4_kgj * md_RendmentPO4
    PO4precipite_molj = PPO4precipite_kgj / pd_MASSE_MOLAIRE_P_kg_mol   'car exprinmé en kg de P

    'Calcul de la quantité de réactif injecté
    If mn_RatioMolaireMgP_force = True Then
        md_RatioMolaireMgP = md_RatioMolaireMgP_f
    Else
        md_RatioMolaireMgP = md_HypRatioMolaireMgP
    End If
    MgCl2_molj = md_RatioMolaireMgP * (PO4precipite_molj / md_RendmentPO4)   'ce que l'on a besoin pour produire la struvite
    MgCl2_molj = MgCl2_molj - (md_InletMg_mgL * Q_m3j / 1000) / pd_MASSE_MOLAIRE_Mg_kg_mol     'on retire le Mg déjà présent dans l'eau
    If MgCl2_molj < 0 Then
        If mn_RatioMolaireMgP_force = True Then

        Else
            md_RatioMolaireMgP = ((md_InletMg_mgL * Q_m3j / 1000) / pd_MASSE_MOLAIRE_Mg_kg_mol) / (PO4precipite_molj / md_RendmentPO4)
            MgCl2_molj = 0
        End If
    End If
    pd_MgCl2pur_kgj = MgCl2_molj * md_MM_MgCl2_kg_mol

    'Calcul de la struvite formée
    StruviteProduite_kgj = PO4precipite_molj * md_MM_MgNH4PO4H12O6_kg_mol

    'Calcul du NH4 précipité
    NNH4precipite_kgj = PO4precipite_molj * pd_MASSE_MOLAIRE_N_kg_mol

    'On calcule le phosphore et l'azote en sortie de réacteur
    Pt_kgj = Pt_kgj - PPO4precipite_kgj
    PPO4_kgj = PPO4_kgj - PPO4precipite_kgj
    NNH4_kgj = NNH4_kgj - NNH4precipite_kgj
    NK_kgj = NK_kgj - NNH4precipite_kgj



End Sub

Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'Mg dissous dans les retours en entrée
    If mn_InletMg_mgL_force = True Then
        md_InletMg_mgL = md_InletMg_mgL_f
    Else
        md_InletMg_mgL = 0     'Christophe MELE, pessimiste car sur-estime la consommation de réactif
    End If

    'Mg dissous dans les retours en entrée
    If mn_InletCa_mgL_force = True Then
        md_InletCa_mgL = md_InletCa_mgL_f
    Else
        md_InletCa_mgL = 0
    End If

    'pH des retours en entrée
    If mn_InletpH_force = True Then
        md_InletpH = md_InletpH_f
    Else
        md_InletpH = 7.5
    End If

    If md_InletpH < 7.5 Then
        ms_Remarque(mi_ErrInletpH) = Feuil70.Range("MAP_retours_ErrInletpH").Cells(1, 1)
        '************************************ message: pH trop bas, chute du rendement probable, ajout soude envisagé ***************************************
    Else
        ms_Remarque(mi_ErrInletpH) = ""
    End If

    'Température
    If mn_InletTemperature_°C_force = True Then
        md_InletTemperature_°C = md_InletTemperature_°C_f
    Else
        md_InletTemperature_°C = 25
    End If

    'Ratio molaire de Mg/PPO4 en entrée, source: Christophe MELE à confirmer si sur entrée ou précipité
    If mn_RatioMolaireMgP_force = True Then
        md_RatioMolaireMgP = md_RatioMolaireMgP_f
    Else
        md_RatioMolaireMgP = md_HypRatioMolaireMgP
    End If
    If md_RatioMolaireMgP < 1 Then
        '********** message: ratio molaire Mg/P trop faible pour précipitation struvite **************************
        ms_Remarque(mi_ErrRatioMolaireMgPFaible) = Feuil70.Range("MAP_retours_ErrRatioMolaireMgPFaible").Cells(1, 1)
        ms_Remarque(mi_ErrRatioMolaireMgPTrop) = ""    'Feuil70.Range("MAP_retours_ErrRatioMolaireMgPTrop").Cells(1, 1)
    ElseIf md_RatioMolaireMgP > 1.5 Then
        '********** message: consommation de MgCl2 excessive **************************
        ms_Remarque(mi_ErrRatioMolaireMgPFaible) = ""    'Feuil70.Range("MAP_retours_ErrRatioMolaireMgPFaible").Cells(1, 1)
        ms_Remarque(mi_ErrRatioMolaireMgPTrop) = Feuil70.Range("MAP_retours_ErrRatioMolaireMgPTrop").Cells(1, 1)
    Else
        ms_Remarque(mi_ErrRatioMolaireMgPFaible) = ""    'Feuil70.Range("MAP_retours_ErrRatioMolaireMgPFaible").Cells(1, 1)
        ms_Remarque(mi_ErrRatioMolaireMgPTrop) = ""    'Feuil70.Range("MAP_retours_ErrRatioMolaireMgPTrop").Cells(1, 1)
    End If

    If mn_NaOHpur_kgj_force = True Then
        pd_NaOHpur_kgj = md_NaOHpur_kgj_f
    Else
        pd_NaOHpur_kgj = 0
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "70.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub


Sub dimensionnement(ByRef retour_flux() As Double, ByRef retour_flux_soluble() As Double, ByRef retour_digestion() As Double, ByRef retour_digestion_soluble() As Double, ByRef retour_athos() As Double, ByRef retour_athos_soluble() As Double)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim Q As Double
    Dim DCO As Double
    Dim DBO As Double
    Dim MES As Double
    Dim NK As Double
    Dim NNH4 As Double
    Dim NO3 As Double
    Dim PPO4 As Double
    Dim Pt As Double

    Dim FacteurDebit As Double

    Dim NbReacteurs As Integer

    Dim retour_flux_non_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_soluble_non_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_soluble_admis(retour_caracteristique_nb) As Double
    Dim retour_flux_admis(retour_caracteristique_nb) As Double

    'On regarde les retours admis
    Call mcls_RetoursAdmis.dimensionnement(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble, retour_flux_non_admis, retour_flux_soluble_non_admis, retour_flux_admis, retour_flux_soluble_admis)

    'on récupère les valeurs des retours
    If mn_NominalQ_m3j_force = True Then
        md_NominalQ_m3j = md_NominalQ_m3j_f
    Else
        md_NominalQ_m3j = mcls_RetoursAdmis.pd_Q_m3j
    End If
    Q = md_NominalQ_m3j
    FacteurDebit = mcls_RetoursAdmis.pd_Q_m3j / md_NominalQ_m3j
    DCO = mcls_RetoursAdmis.pd_DCO_mgL * FacteurDebit * Q / 1000
    DBO = mcls_RetoursAdmis.pd_DBO_mgL * FacteurDebit * Q / 1000
    MES = mcls_RetoursAdmis.pd_MES_mgL * FacteurDebit * Q / 1000
    If mn_NominalNNH4_mgL_force = True Then
        NNH4 = md_NominalNNH4_mgL_f * Q / 1000
        md_NominalNNH4_mgL = md_NominalNNH4_mgL_f
    Else
        NNH4 = mcls_RetoursAdmis.pd_NNH4_mgL * FacteurDebit * Q / 1000
        md_NominalNNH4_mgL = NNH4 / Q * 1000
    End If
    NK = (mcls_RetoursAdmis.pd_NK_mgL - mcls_RetoursAdmis.pd_NNH4_mgL) * FacteurDebit * Q / 1000 + NNH4
    NO3 = mcls_RetoursAdmis.pd_NNO3_mgL * FacteurDebit * Q / 1000
    If mn_NominalPPO4_mgL_force = True Then
        PPO4 = md_NominalPPO4_mgL_f * Q / 1000
        md_NominalPPO4_mgL = md_NominalPPO4_mgL_f
    Else
        PPO4 = mcls_RetoursAdmis.pd_PPO4_mgL * FacteurDebit * Q / 1000
        md_NominalPPO4_mgL = PPO4 / Q * 1000
    End If
    Pt = (mcls_RetoursAdmis.pd_Pt_mgL - mcls_RetoursAdmis.pd_PPO4_mgL) * FacteurDebit * Q / 1000 + PPO4

    'On calcule la cristallisation
    Call calcul_cristallisation(Q, PPO4, Pt, NNH4, NK, md_NominalStruviteProduite_kgj)

    'On dimensionne suivant le type de procédé
    If mi_ChoixProcess = mi_OSTARA Then
        'On vérifie la production de struvite
        If md_NominalStruviteProduite_kgj < md_ProductionMinimaleStruvite_kgj Then
            '******************** message: production de struvite trop faible pour OSTARA *******************************
            ms_Remarque(mi_ErrOstaraProductionStruviteFaible) = Feuil70.Range("MAP_retours_ErrOstaraProductionStruviteFaible").Cells(1, 1)
            ms_ReacteursInstalles = "'-"
        ElseIf md_NominalStruviteProduite_kgj < md_ProductionMaximalePEARL500_kgj Then
            ms_ReacteursInstalles = "1 PEARL 500"
            ms_Remarque(mi_ErrOstaraProductionStruviteFaible) = ""    'Feuil70.Range("MAP_retours_ErrOstaraProductionStruviteFaible").Cells(1, 1)
        ElseIf md_NominalStruviteProduite_kgj < 2 * md_ProductionMaximalePEARL500_kgj Then
            ms_ReacteursInstalles = "2 PEARL 500"
            ms_Remarque(mi_ErrOstaraProductionStruviteFaible) = ""    'Feuil70.Range("MAP_retours_ErrOstaraProductionStruviteFaible").Cells(1, 1)
        Else
            NbReacteurs = 0
            Do
                NbReacteurs = NbReacteurs + 1
            Loop While NbReacteurs < (md_NominalStruviteProduite_kgj / md_ProductionMaximalePEARL2000_kgj)
            ms_ReacteursInstalles = NbReacteurs & " PEARL 2000"
            ms_Remarque(mi_ErrOstaraProductionStruviteFaible) = ""    'Feuil70.Range("MAP_retours_ErrOstaraProductionStruviteFaible").Cells(1, 1)
        End If
    Else
        ms_ReacteursInstalles = "'-"
    End If

    'calcul des charges en sortie
    Q = Q
    DCO = DCO
    DBO = DBO
    MES = MES
    NNH4 = NNH4
    NK = NK
    NO3 = NO3
    Pt = Pt


    'on récupère les valeurs des retours admis
    retour_flux_admis(repere_ret_Q) = Q
    retour_flux_admis(repere_ret_DCO) = DCO
    retour_flux_admis(repere_ret_DBO) = DBO
    retour_flux_admis(repere_ret_MES) = MES
    retour_flux_admis(repere_ret_NK) = NK
    retour_flux_admis(repere_ret_NH4) = NNH4
    retour_flux_admis(repere_ret_NO3) = NO3
    retour_flux_admis(repere_ret_Pt) = Pt
    retour_flux_soluble_admis(repere_ret_NK) = retour_flux_soluble_admis(repere_ret_NK) - retour_flux_soluble_admis(repere_ret_NH4) + NNH4
    retour_flux_soluble_admis(repere_ret_NH4) = NNH4
    retour_flux_soluble_admis(repere_ret_Pt) = PPO4
    Call mcls_RetoursAdmis.flux_sortie_process(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble, retour_flux_non_admis, retour_flux_soluble_non_admis, retour_flux_admis, retour_flux_soluble_admis)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "70.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim Colonne As Integer


    Colonne = pi_FiliereConsideree

    mn_InletMg_mgL_force = Feuil70.Range("MAP_retours_InletMg_mgL").Cells(1, Colonne)
    mn_InletCa_mgL_force = Feuil70.Range("MAP_retours_InletCa_mgL").Cells(1, Colonne)
    mn_NominalQ_m3j_force = Feuil70.Range("MAP_retours_NominalQ_m3j").Cells(1, Colonne)
    mn_NominalPPO4_mgL_force = Feuil70.Range("MAP_retours_NominalPPO4_mgL").Cells(1, Colonne)
    mn_NominalNNH4_mgL_force = Feuil70.Range("MAP_retours_NominalNNH4_mgL").Cells(1, Colonne)
    mn_InletpH_force = Feuil70.Range("MAP_retours_InletpH").Cells(1, Colonne)
    mn_NaOHpur_kgj_force = Feuil70.Range("MAP_retours_NaOHpur_kgj").Cells(1, Colonne)
    mn_InletTemperature_°C_force = Feuil70.Range("MAP_retours_InletTemperature__C").Cells(1, Colonne)
    mn_RatioMolaireMgP_force = Feuil70.Range("MAP_retours_RatioMolaireMgP").Cells(1, Colonne)
    mn_RendmentPO4_force = Feuil70.Range("MAP_retours_RendmentPO4").Cells(1, Colonne)

    If mn_InletMg_mgL_force = True Then
        md_InletMg_mgL_f = Feuil70.Range("MAP_retours_InletMg_mgL").Cells(1, Colonne)
    End If
    If mn_InletCa_mgL_force = True Then
        md_InletCa_mgL_f = Feuil70.Range("MAP_retours_InletCa_mgL").Cells(1, Colonne)
    End If
    If mn_NominalQ_m3j_force = True Then
        md_NominalQ_m3j_f = Feuil70.Range("MAP_retours_NominalQ_m3j").Cells(1, Colonne)
    End If
    If mn_NominalPPO4_mgL_force = True Then
        md_NominalPPO4_mgL_f = Feuil70.Range("MAP_retours_NominalPPO4_mgL").Cells(1, Colonne)
    End If
    If mn_NominalNNH4_mgL_force = True Then
        md_NominalNNH4_mgL_f = Feuil70.Range("MAP_retours_NominalNNH4_mgL").Cells(1, Colonne)
    End If
    If mn_InletpH_force = True Then
        md_InletpH_f = Feuil70.Range("MAP_retours_InletpH").Cells(1, Colonne)
    End If
    If mn_NaOHpur_kgj_force = True Then
        md_NaOHpur_kgj_f = Feuil70.Range("MAP_retours_NaOHpur_kgj").Cells(1, Colonne)
    End If
    If mn_InletTemperature_°C_force = True Then
        md_InletTemperature_°C_f = Feuil70.Range("MAP_retours_InletTemperature__C").Cells(1, Colonne)
    End If
    If mn_RatioMolaireMgP_force = True Then
        md_RatioMolaireMgP_f = Feuil70.Range("MAP_retours_RatioMolaireMgP").Cells(1, Colonne)
    End If
    If mn_RendmentPO4_force = True Then
        md_RendmentPO4_f = Feuil70.Range("MAP_retours_RendmentPO4").Cells(1, Colonne)
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "70.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub affich_remarque()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = mi_NbErreurs  'penser à modifier la capacité du vecteur "message_erreur"


    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil70.Range("MAP_retours_erreur_nom")

    For I = 1 To nb_remarque
        If ms_Remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + ms_Remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "70.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub fonctionnement_reel(ByRef retour_flux() As Double, ByRef retour_flux_soluble() As Double, ByRef retour_digestion() As Double, ByRef retour_digestion_soluble() As Double, ByRef retour_athos() As Double, ByRef retour_athos_soluble() As Double)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim Q As Double
    Dim DCO As Double
    Dim DBO As Double
    Dim MES As Double
    Dim NK As Double
    Dim NNH4 As Double
    Dim NO3 As Double
    Dim PPO4 As Double
    Dim Pt As Double

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
    NNH4 = mcls_RetoursAdmis.pd_NNH4_mgL * Q / 1000
    NO3 = mcls_RetoursAdmis.pd_NNO3_mgL * Q / 1000
    PPO4 = mcls_RetoursAdmis.pd_PPO4_mgL * Q / 1000
    Pt = mcls_RetoursAdmis.pd_Pt_mgL * Q / 1000


    'On vérifie la concentration en entrée en PO4
    If (PPO4 / Q) * 1000 < md_InletPPO4Mini_mgL Then
        '******************** message: concentration trop faible en entrée *******************************
        ms_Remarque(mi_ErrInletPPO4Faible) = Feuil70.Range("MAP_retours_ErrInletPPO4Faible").Cells(1, 1)
    Else
        ms_Remarque(mi_ErrInletPPO4Faible) = ""     'Feuil70.Range("MAP_retours_ErrInletPPO4Faible").Cells(1, 1)
    End If
    'On vérifie la concentration en entrée en MES
    If (MES / Q) * 1000 > md_InletMESMaxi_mgL Then
        '******************** message: concentration élevée en MES en entrée, risque de dégradation de la qualité *******************************
        ms_Remarque(mi_ErrInletMESTrop) = Feuil70.Range("MAP_retours_ErrInletMESTrop").Cells(1, 1)
    Else
        ms_Remarque(mi_ErrInletMESTrop) = ""     'Feuil70.Range("MAP_retours_ErrInletMESTrop").Cells(1, 1)
    End If

    'On calcule la cristallisation
    Call calcul_cristallisation(Q, PPO4, Pt, NNH4, NK, pd_StruviteProduite_kgj)

    'Concentration en PPO4 en sortie
    md_OutletPPO4_mgL = PPO4 / Q * 1000
    If md_OutletPPO4_mgL < md_OutletPPO4Mini_mgL Then
        '******************** message: concentration élevée en MES en entrée, risque de dégradation de la qualité *******************************
        ms_Remarque(mi_ErrOutletPPO4Faible) = Feuil70.Range("MAP_retours_ErrOutletPPO4Faible").Cells(1, 1)
    Else
        ms_Remarque(mi_ErrOutletPPO4Faible) = ""     'Feuil70.Range("MAP_retours_ErrOutletPPO4Faible").Cells(1, 1)
    End If

    'début modif MSM 15/11/12
    '-->idem mais pour NNH4 creer sortir afficher variable
    'Concentration en NNH4 en sortie
    md_OutletNNH4_mgL = NNH4 / Q * 1000

    'calcul des charges en sortie
    Q = Q
    DCO = DCO
    DBO = DBO
    MES = MES
    NNH4 = NNH4
    NK = NK
    NO3 = NO3
    PPO4 = PPO4
    Pt = Pt

    'on récupère les valeurs des retours admis
    retour_flux_admis(repere_ret_Q) = Q
    retour_flux_admis(repere_ret_DCO) = DCO
    retour_flux_admis(repere_ret_DBO) = DBO
    retour_flux_admis(repere_ret_MES) = MES
    retour_flux_admis(repere_ret_NK) = NK
    retour_flux_admis(repere_ret_NH4) = NNH4
    retour_flux_admis(repere_ret_NO3) = NO3
    retour_flux_admis(repere_ret_Pt) = Pt
    retour_flux_soluble_admis(repere_ret_NK) = retour_flux_soluble_admis(repere_ret_NK) - retour_flux_soluble_admis(repere_ret_NH4) + NNH4
    retour_flux_soluble_admis(repere_ret_NH4) = NNH4
    retour_flux_soluble_admis(repere_ret_Pt) = PPO4
    Call mcls_RetoursAdmis.flux_sortie_process(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble, retour_flux_non_admis, retour_flux_soluble_non_admis, retour_flux_admis, retour_flux_soluble_admis)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "70.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub ecriture_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil70.Range("MAP_retours_InletMg_mgL").Cells(1, Colonne) = md_InletMg_mgL
    Feuil70.Range("MAP_retours_InletCa_mgL").Cells(1, Colonne) = md_InletCa_mgL
    Feuil70.Range("MAP_retours_NominalQ_m3j").Cells(1, Colonne) = md_NominalQ_m3j
    Feuil70.Range("MAP_retours_NominalPPO4_mgL").Cells(1, Colonne) = md_NominalPPO4_mgL
    Feuil70.Range("MAP_retours_NominalNNH4_mgL").Cells(1, Colonne) = md_NominalNNH4_mgL
    Feuil70.Range("MAP_retours_NominalStruviteProduite_kgj").Cells(1, Colonne) = md_NominalStruviteProduite_kgj
    Feuil70.Range("MAP_retours_ReacteursInstalles").Cells(1, Colonne) = ms_ReacteursInstalles
    Feuil70.Range("MAP_retours_InletpH").Cells(1, Colonne) = md_InletpH
    Feuil70.Range("MAP_retours_NaOHpur_kgj").Cells(1, Colonne) = pd_NaOHpur_kgj
    Feuil70.Range("MAP_retours_InletTemperature__C").Cells(1, Colonne) = md_InletTemperature_°C
    Feuil70.Range("MAP_retours_RatioMolaireMgP").Cells(1, Colonne) = md_RatioMolaireMgP
    Feuil70.Range("MAP_retours_RendmentPO4").Cells(1, Colonne) = md_RendmentPO4
    Feuil70.Range("MAP_retours_OutletPPO4_mgL").Cells(1, Colonne) = md_OutletPPO4_mgL
    Feuil70.Range("MAP_retours_OutletNNH4_mgL").Cells(1, Colonne) = md_OutletNNH4_mgL    ' MSM 15/11/12
    Feuil70.Range("MAP_retours_MgCl2pur_kgj").Cells(1, Colonne) = pd_MgCl2pur_kgj
    Feuil70.Range("MAP_retours_StruviteProduite_kgj").Cells(1, Colonne) = pd_StruviteProduite_kgj
    Feuil70.Range("MAP_retours_electricite_kWhj").Cells(1, Colonne) = pd_electricite_kWhj

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "70.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub effacer_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil70.Range("MAP_retours_InletMg_mgL").Cells(1, Colonne) = ""    'md_InletMg_mgL
    Feuil70.Range("MAP_retours_InletCa_mgL").Cells(1, Colonne) = ""    'md_InletCa_mgL
    Feuil70.Range("MAP_retours_NominalQ_m3j").Cells(1, Colonne) = ""    'md_NominalQ_m3j
    Feuil70.Range("MAP_retours_NominalPPO4_mgL").Cells(1, Colonne) = ""    'md_NominalPPO4_mgL
    Feuil70.Range("MAP_retours_NominalNNH4_mgL").Cells(1, Colonne) = ""    'md_NominalNNH4_mgL
    Feuil70.Range("MAP_retours_NominalStruviteProduite_kgj").Cells(1, Colonne) = ""    'md_NominalStruviteProduite_kgj
    Feuil70.Range("MAP_retours_ReacteursInstalles").Cells(1, Colonne) = ""    'ms_ReacteursInstalles
    Feuil70.Range("MAP_retours_InletpH").Cells(1, Colonne) = ""    'md_InletpH
    Feuil70.Range("MAP_retours_NaOHpur_kgj").Cells(1, Colonne) = ""    'pd_NaOHpur_kgj
    Feuil70.Range("MAP_retours_InletTemperature__C").Cells(1, Colonne) = ""    'md_InletTemperature_°C
    Feuil70.Range("MAP_retours_RatioMolaireMgP").Cells(1, Colonne) = ""    'md_RatioMolaireMgP
    Feuil70.Range("MAP_retours_RendmentPO4").Cells(1, Colonne) = ""    'md_RendmentPO4
    Feuil70.Range("MAP_retours_OutletPPO4_mgL").Cells(1, Colonne) = ""    'md_OutletPPO4_mgL
    Feuil70.Range("MAP_retours_OutletNNH4_mgL").Cells(1, Colonne) = ""    'md_OutletNNH4_mgL        MSM 15/11/12
    Feuil70.Range("MAP_retours_MgCl2pur_kgj").Cells(1, Colonne) = ""    'pd_MgCl2pur_kgj
    Feuil70.Range("MAP_retours_StruviteProduite_kgj").Cells(1, Colonne) = ""    'pd_StruviteProduite_kgj
    Feuil70.Range("MAP_retours_electricite_kWhj").Cells(1, Colonne) = ""    'pd_electricite_kWhj

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "70.10"

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
    mcls_RetoursAdmis.pn_EcritureResultats = True

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


Sub calcul_consommation_electrique()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    If mi_ChoixProcess = mi_OSTARA Then
        If md_NominalStruviteProduite_kgj < 2 * md_ProductionMaximalePEARL500_kgj Then
            pd_electricite_kWhj = md_HypElectricitePEARL500_kWh_kg * pd_StruviteProduite_kgj
        Else
            pd_electricite_kWhj = md_HypElectricitePEARL2000_kWh_kg * pd_StruviteProduite_kgj
        End If
    Else
        pd_electricite_kWhj = 0
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "70.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


