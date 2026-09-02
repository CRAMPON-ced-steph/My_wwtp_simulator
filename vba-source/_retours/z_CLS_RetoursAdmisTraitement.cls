Attribute VB_Name = "z_CLS_RetoursAdmisTraitement"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'Constante sur le choix des retours admis
Private Const mi_JUS_DIGESTION As Integer = 1
Private Const mi_JUS_ATHOS As Integer = 2
Private Const mi_JUS_DIGESTION_ATHOS As Integer = 3

'Choix des retours admis sur le traitement séparé
Dim mi_ChoixRetoursAdmisTraitement As Integer
Public pn_EcritureResultats As Boolean

'Valeurs retournées à l'utilisateur
Public pd_Q_m3j As Double
Public pd_DCO_mgL As Double
Public pd_DBO_mgL As Double
Public pd_MES_mgL As Double
Public pd_NK_mgL As Double
Public pd_NNH4_mgL As Double
Public pd_NNO3_mgL As Double
Public pd_Pt_mgL As Double
Public pd_PPO4_mgL As Double
Dim md_DCO_kgj As Double
Dim md_DBO_kgj As Double
Dim md_MES_kgj As Double
Dim md_NK_kgj As Double
Dim md_NNH4_kgj As Double
Dim md_NNO3_kgj As Double
Dim md_Pt_kgj As Double
Dim md_PPO4_kgj As Double

'Repère le forçage des valeurs par l'utilisateur
Dim mn_Q_m3j_force As Boolean
Dim mn_DCO_kgj_force As Boolean
Dim mn_DBO_kgj_force As Boolean
Dim mn_MES_kgj_force As Boolean
Dim mn_NK_kgj_force As Boolean
Dim mn_NNH4_kgj_force As Boolean
Dim mn_NNO3_kgj_force As Boolean
Dim mn_Pt_kgj_force As Boolean
Dim mn_PPO4_kgj_force As Boolean

'Stocke la valeur forcée
Dim md_Q_m3j_f As Double
Dim md_DCO_kgj_f As Double
Dim md_DBO_kgj_f As Double
Dim md_MES_kgj_f As Double
Dim md_NK_kgj_f As Double
Dim md_NNH4_kgj_f As Double
Dim md_NNO3_kgj_f As Double
Dim md_Pt_kgj_f As Double
Dim md_PPO4_kgj_f As Double


Private Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer
    Dim message_erreur As String

    Ligne = 1
    Colonne = pi_FiliereConsideree

    'On lit le choix sur les retours admis dans le traitement
    mi_ChoixRetoursAdmisTraitement = Feuil5.Range("choix_retours_admis_traitement").Cells(Ligne, Colonne)

    If mi_ChoixRetoursAdmisTraitement = mi_JUS_DIGESTION Then
        If pi_ChoixDigestion = False Then
            message_erreur = Feuil5.Range("erreur_choix_retours_admis").Cells(1, 1)
            MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
            pn_EffacementResultatsEnCours = True
            pn_ErreurNonGeree = True
            Exit Sub    'met fin aux calculs
        End If
    ElseIf mi_ChoixRetoursAdmisTraitement = mi_JUS_ATHOS Then
        If choix_athos = False Then
            message_erreur = Feuil5.Range("erreur_choix_retours_admis").Cells(1, 1)
            MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
            pn_EffacementResultatsEnCours = True
            pn_ErreurNonGeree = True
            Exit Sub    'met fin aux calculs
        End If
    Else
        If pi_ChoixDigestion = False And choix_athos = False Then
            message_erreur = Feuil5.Range("erreur_choix_retours_admis").Cells(1, 1)
            MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
            pn_EffacementResultatsEnCours = True
            pn_ErreurNonGeree = True
            Exit Sub    'met fin aux calculs
        End If
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "5.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree
    Ligne = 1

    'On écrit les résultats sur la composition des retours
    If pn_EcritureResultats = True Then
        Feuil6.Range("retours_admis_Q_m3j").Cells(Ligne, Colonne) = pd_Q_m3j
        Feuil6.Range("retours_admis_DCO_mgL").Cells(Ligne, Colonne) = pd_DCO_mgL
        Feuil6.Range("retours_admis_DBO_mgL").Cells(Ligne, Colonne) = pd_DBO_mgL
        Feuil6.Range("retours_admis_MES_mgL").Cells(Ligne, Colonne) = pd_MES_mgL
        Feuil6.Range("retours_admis_NK_mgL").Cells(Ligne, Colonne) = pd_NK_mgL
        Feuil6.Range("retours_admis_NNH4_mgL").Cells(Ligne, Colonne) = pd_NNH4_mgL
        Feuil6.Range("retours_admis_NNO3_mgL").Cells(Ligne, Colonne) = pd_NNO3_mgL
        Feuil6.Range("retours_admis_Pt_mgL").Cells(Ligne, Colonne) = pd_Pt_mgL
        Feuil6.Range("retours_admis_PPO4_mgL").Cells(Ligne, Colonne) = pd_PPO4_mgL

        Feuil6.Range("retours_admis_DCO_kgj").Cells(Ligne, Colonne) = md_DCO_kgj
        Feuil6.Range("retours_admis_DBO_kgj").Cells(Ligne, Colonne) = md_DBO_kgj
        Feuil6.Range("retours_admis_MES_kgj").Cells(Ligne, Colonne) = md_MES_kgj
        Feuil6.Range("retours_admis_NK_kgj").Cells(Ligne, Colonne) = md_NK_kgj
        Feuil6.Range("retours_admis_NNH4_kgj").Cells(Ligne, Colonne) = md_NNH4_kgj
        Feuil6.Range("retours_admis_NNO3_kgj").Cells(Ligne, Colonne) = md_NNO3_kgj
        Feuil6.Range("retours_admis_Pt_kgj").Cells(Ligne, Colonne) = md_Pt_kgj
        Feuil6.Range("retours_admis_PPO4_kgj").Cells(Ligne, Colonne) = md_PPO4_kgj
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "5.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer

    Colonne = pi_FiliereConsideree
    Ligne = 1

    'On écrit les résultats sur la composition des retours
    mn_Q_m3j_force = Feuil6.Range("retours_admis_Q_m3j").Cells(Ligne, Colonne)
    mn_DCO_kgj_force = Feuil6.Range("retours_admis_DCO_kgj").Cells(Ligne, Colonne)
    mn_DBO_kgj_force = Feuil6.Range("retours_admis_DBO_kgj").Cells(Ligne, Colonne)
    mn_MES_kgj_force = Feuil6.Range("retours_admis_MES_kgj").Cells(Ligne, Colonne)
    mn_NK_kgj_force = Feuil6.Range("retours_admis_NK_kgj").Cells(Ligne, Colonne)
    mn_NNH4_kgj_force = Feuil6.Range("retours_admis_NNH4_kgj").Cells(Ligne, Colonne)
    mn_NNO3_kgj_force = Feuil6.Range("retours_admis_NNO3_kgj").Cells(Ligne, Colonne)
    mn_Pt_kgj_force = Feuil6.Range("retours_admis_Pt_kgj").Cells(Ligne, Colonne)
    mn_PPO4_kgj_force = Feuil6.Range("retours_admis_PPO4_kgj").Cells(Ligne, Colonne)

    If mn_Q_m3j_force = True Then
        md_Q_m3j_f = Feuil6.Range("retours_admis_Q_m3j").Cells(Ligne, Colonne)
    End If
    If mn_DCO_kgj_force = True Then
        md_DCO_kgj_f = Feuil6.Range("retours_admis_DCO_kgj").Cells(Ligne, Colonne)
    End If
    If mn_DBO_kgj_force = True Then
        md_DBO_kgj_f = Feuil6.Range("retours_admis_DBO_kgj").Cells(Ligne, Colonne)
    End If
    If mn_MES_kgj_force = True Then
        md_MES_kgj_f = Feuil6.Range("retours_admis_MES_kgj").Cells(Ligne, Colonne)
    End If
    If mn_NK_kgj_force = True Then
        md_NK_kgj_f = Feuil6.Range("retours_admis_NK_kgj").Cells(Ligne, Colonne)
    End If
    If mn_NNH4_kgj_force = True Then
        md_NNH4_kgj_f = Feuil6.Range("retours_admis_NNH4_kgj").Cells(Ligne, Colonne)
    End If
    If mn_NNO3_kgj_force = True Then
        md_NNO3_kgj_f = Feuil6.Range("retours_admis_NNO3_kgj").Cells(Ligne, Colonne)
    End If
    If mn_Pt_kgj_force = True Then
        md_Pt_kgj_f = Feuil6.Range("retours_admis_Pt_kgj").Cells(Ligne, Colonne)
    End If
    If mn_PPO4_kgj_force = True Then
        md_PPO4_kgj_f = Feuil6.Range("retours_admis_PPO4_kgj").Cells(Ligne, Colonne)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "5.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "5.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "5.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Public Sub dimensionnement(ByRef RetourFlux() As Double, ByRef RetourFluxSoluble() As Double, ByRef RetourFluxDigestion() As Double, ByRef RetourFluxDigestionSoluble() As Double, ByRef RetourFluxAthos() As Double, ByRef RetourFluxAthosSoluble() As Double, ByRef RetourFluxNonAdmisTraitement() As Double, ByRef RetourFluxNonAdmisTraitementSoluble() As Double, ByRef RetourFluxAdmisTraitement() As Double, ByRef RetourFluxAdmisTraitementSoluble() As Double)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer

    'On calcule les flux qui rentrent sur le traitement séparé et ce qui le by-passe
    If mi_ChoixRetoursAdmisTraitement = mi_JUS_DIGESTION Then
        For I = 1 To retour_caracteristique_nb
            RetourFluxNonAdmisTraitement(I) = RetourFlux(I) - RetourFluxDigestion(I)
            RetourFluxNonAdmisTraitementSoluble(I) = RetourFluxSoluble(I) - RetourFluxDigestionSoluble(I)
'            RetourFluxAdmisTraitement(i) = RetourFluxDigestion(i)
            RetourFluxAdmisTraitementSoluble(I) = RetourFluxDigestionSoluble(I)
        Next I
        'On attribue les valeurs
        pd_Q_m3j = RetourFluxDigestion(repere_ret_Q)
        md_DCO_kgj = RetourFluxDigestion(repere_ret_DCO)
        md_DBO_kgj = RetourFluxDigestion(repere_ret_DBO)
        md_MES_kgj = RetourFluxDigestion(repere_ret_MES)
        md_NK_kgj = RetourFluxDigestion(repere_ret_NK)
        md_NNH4_kgj = RetourFluxDigestion(repere_ret_NH4)
        md_NNO3_kgj = RetourFluxDigestion(repere_ret_NO3)
        md_Pt_kgj = RetourFluxDigestion(repere_ret_Pt)
        md_PPO4_kgj = RetourFluxDigestionSoluble(repere_ret_Pt)
    ElseIf mi_ChoixRetoursAdmisTraitement = mi_JUS_ATHOS Then
        For I = 1 To retour_caracteristique_nb
            RetourFluxNonAdmisTraitement(I) = RetourFlux(I) - RetourFluxAthos(I)
            RetourFluxNonAdmisTraitementSoluble(I) = RetourFluxSoluble(I) - RetourFluxAthosSoluble(I)
'            RetourFluxAdmisTraitement(i) = RetourFluxAthos(i)
            RetourFluxAdmisTraitementSoluble(I) = RetourFluxAthosSoluble(I)
        Next I
        pd_Q_m3j = RetourFluxAthos(repere_ret_Q)
        md_DCO_kgj = RetourFluxAthos(repere_ret_DCO)
        md_DBO_kgj = RetourFluxAthos(repere_ret_DBO)
        md_MES_kgj = RetourFluxAthos(repere_ret_MES)
        md_NK_kgj = RetourFluxAthos(repere_ret_NK)
        md_NNH4_kgj = RetourFluxAthos(repere_ret_NH4)
        md_NNO3_kgj = RetourFluxAthos(repere_ret_NO3)
        md_Pt_kgj = RetourFluxAthos(repere_ret_Pt)
        md_PPO4_kgj = RetourFluxAthosSoluble(repere_ret_Pt)
    Else
        For I = 1 To retour_caracteristique_nb
            RetourFluxNonAdmisTraitement(I) = RetourFlux(I) - RetourFluxDigestion(I) - RetourFluxAthos(I)
            RetourFluxNonAdmisTraitementSoluble(I) = RetourFluxSoluble(I) - RetourFluxDigestionSoluble(I) - RetourFluxAthosSoluble(I)
'            RetourFluxAdmisTraitement(i) = RetourFluxAthos(i) + RetourFluxDigestion(i)
            RetourFluxAdmisTraitementSoluble(I) = RetourFluxAthosSoluble(I) + RetourFluxDigestionSoluble(I)
        Next I
        pd_Q_m3j = RetourFluxDigestion(repere_ret_Q) + RetourFluxAthos(repere_ret_Q)
        md_DCO_kgj = (RetourFluxDigestion(repere_ret_DCO) + RetourFluxAthos(repere_ret_DCO))
        md_DBO_kgj = (RetourFluxDigestion(repere_ret_DBO) + RetourFluxAthos(repere_ret_DBO))
        md_MES_kgj = (RetourFluxDigestion(repere_ret_MES) + RetourFluxAthos(repere_ret_MES))
        md_NK_kgj = (RetourFluxDigestion(repere_ret_NK) + RetourFluxAthos(repere_ret_NK))
        md_NNH4_kgj = (RetourFluxDigestion(repere_ret_NH4) + RetourFluxAthos(repere_ret_NH4))
        md_NNO3_kgj = (RetourFluxDigestion(repere_ret_NO3) + RetourFluxAthos(repere_ret_NO3))
        md_Pt_kgj = (RetourFluxDigestion(repere_ret_Pt) + RetourFluxAthos(repere_ret_Pt))
        md_PPO4_kgj = (RetourFluxDigestionSoluble(repere_ret_Pt) + RetourFluxAthosSoluble(repere_ret_Pt))
    End If

    'On calcule la concentration
    If pd_Q_m3j > 0 Then
        pd_DCO_mgL = md_DCO_kgj / pd_Q_m3j * 1000
        pd_DBO_mgL = md_DBO_kgj / pd_Q_m3j * 1000
        pd_MES_mgL = md_MES_kgj / pd_Q_m3j * 1000
        pd_NK_mgL = md_NK_kgj / pd_Q_m3j * 1000
        pd_NNH4_mgL = md_NNH4_kgj / pd_Q_m3j * 1000
        pd_NNO3_mgL = md_NNO3_kgj / pd_Q_m3j * 1000
        pd_Pt_mgL = md_Pt_kgj / pd_Q_m3j * 1000
        pd_PPO4_mgL = md_PPO4_kgj / pd_Q_m3j * 1000
    Else
        pd_DCO_mgL = 0
        pd_DBO_mgL = 0
        pd_MES_mgL = 0
        pd_NK_mgL = 0
        pd_NNH4_mgL = 0
        pd_NNO3_mgL = 0
        pd_Pt_mgL = 0
        pd_PPO4_mgL = 0
        'Erreur et arrêt des calculs
        MsgBox Feuil6.Range("erreur_aucun_retours_admis_traitement_separe"), vbCritical, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "5.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Public Sub fonctionnement_reel(ByRef RetourFlux() As Double, ByRef RetourFluxSoluble() As Double, ByRef RetourFluxDigestion() As Double, ByRef RetourFluxDigestionSoluble() As Double, ByRef RetourFluxAthos() As Double, ByRef RetourFluxAthosSoluble() As Double, ByRef RetourFluxNonAdmisTraitement() As Double, ByRef RetourFluxNonAdmisTraitementSoluble() As Double, ByRef RetourFluxAdmisTraitement() As Double, ByRef RetourFluxAdmisTraitementSoluble() As Double)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer

    'On calcule les flux qui rentrent sur le traitement séparé et ce qui le by-passe
    If mi_ChoixRetoursAdmisTraitement = mi_JUS_DIGESTION Then
        For I = 1 To retour_caracteristique_nb
            RetourFluxNonAdmisTraitement(I) = RetourFlux(I) - RetourFluxDigestion(I)
            RetourFluxNonAdmisTraitementSoluble(I) = RetourFluxSoluble(I) - RetourFluxDigestionSoluble(I)
'            RetourFluxAdmisTraitement(i) = RetourFluxDigestion(i)
            RetourFluxAdmisTraitementSoluble(I) = RetourFluxDigestionSoluble(I)
        Next I
        'On attribue les valeurs
        pd_Q_m3j = RetourFluxDigestion(repere_ret_Q)
        md_DCO_kgj = RetourFluxDigestion(repere_ret_DCO)
        md_DBO_kgj = RetourFluxDigestion(repere_ret_DBO)
        md_MES_kgj = RetourFluxDigestion(repere_ret_MES)
        md_NK_kgj = RetourFluxDigestion(repere_ret_NK)
        md_NNH4_kgj = RetourFluxDigestion(repere_ret_NH4)
        md_NNO3_kgj = RetourFluxDigestion(repere_ret_NO3)
        md_Pt_kgj = RetourFluxDigestion(repere_ret_Pt)
        md_PPO4_kgj = RetourFluxDigestionSoluble(repere_ret_Pt)
    ElseIf mi_ChoixRetoursAdmisTraitement = mi_JUS_ATHOS Then
        For I = 1 To retour_caracteristique_nb
            RetourFluxNonAdmisTraitement(I) = RetourFlux(I) - RetourFluxAthos(I)
            RetourFluxNonAdmisTraitementSoluble(I) = RetourFluxSoluble(I) - RetourFluxAthosSoluble(I)
'            RetourFluxAdmisTraitement(i) = RetourFluxAthos(i)
            RetourFluxAdmisTraitementSoluble(I) = RetourFluxAthosSoluble(I)
        Next I
        pd_Q_m3j = RetourFluxAthos(repere_ret_Q)
        md_DCO_kgj = RetourFluxAthos(repere_ret_DCO)
        md_DBO_kgj = RetourFluxAthos(repere_ret_DBO)
        md_MES_kgj = RetourFluxAthos(repere_ret_MES)
        md_NK_kgj = RetourFluxAthos(repere_ret_NK)
        md_NNH4_kgj = RetourFluxAthos(repere_ret_NH4)
        md_NNO3_kgj = RetourFluxAthos(repere_ret_NO3)
        md_Pt_kgj = RetourFluxAthos(repere_ret_Pt)
        md_PPO4_kgj = RetourFluxAthosSoluble(repere_ret_Pt)
    Else
        For I = 1 To retour_caracteristique_nb
            RetourFluxNonAdmisTraitement(I) = RetourFlux(I) - RetourFluxDigestion(I) - RetourFluxAthos(I)
            RetourFluxNonAdmisTraitementSoluble(I) = RetourFluxSoluble(I) - RetourFluxDigestionSoluble(I) - RetourFluxAthosSoluble(I)
'            RetourFluxAdmisTraitement(i) = RetourFluxAthos(i) + RetourFluxDigestion(i)
            RetourFluxAdmisTraitementSoluble(I) = RetourFluxAthosSoluble(I) + RetourFluxDigestionSoluble(I)
        Next I
        pd_Q_m3j = RetourFluxDigestion(repere_ret_Q) + RetourFluxAthos(repere_ret_Q)
        md_DCO_kgj = (RetourFluxDigestion(repere_ret_DCO) + RetourFluxAthos(repere_ret_DCO))
        md_DBO_kgj = (RetourFluxDigestion(repere_ret_DBO) + RetourFluxAthos(repere_ret_DBO))
        md_MES_kgj = (RetourFluxDigestion(repere_ret_MES) + RetourFluxAthos(repere_ret_MES))
        md_NK_kgj = (RetourFluxDigestion(repere_ret_NK) + RetourFluxAthos(repere_ret_NK))
        md_NNH4_kgj = (RetourFluxDigestion(repere_ret_NH4) + RetourFluxAthos(repere_ret_NH4))
        md_NNO3_kgj = (RetourFluxDigestion(repere_ret_NO3) + RetourFluxAthos(repere_ret_NO3))
        md_Pt_kgj = (RetourFluxDigestion(repere_ret_Pt) + RetourFluxAthos(repere_ret_Pt))
        md_PPO4_kgj = (RetourFluxDigestionSoluble(repere_ret_Pt) + RetourFluxAthosSoluble(repere_ret_Pt))
    End If

    'On regarde les valeurs forcées
    If mn_Q_m3j_force = True Then
        pd_Q_m3j = md_Q_m3j_f
    End If
    If mn_DCO_kgj_force = True Then
        RetourFluxAdmisTraitementSoluble(repere_ret_DCO) = RetourFluxAdmisTraitementSoluble(repere_ret_DCO) * md_DCO_kgj_f / md_DCO_kgj
        md_DCO_kgj = md_DCO_kgj_f
    End If
    If mn_DBO_kgj_force = True Then
        RetourFluxAdmisTraitementSoluble(repere_ret_DBO) = RetourFluxAdmisTraitementSoluble(repere_ret_DBO) * md_DBO_kgj_f / md_DBO_kgj
        md_DBO_kgj = md_DBO_kgj_f
    End If
    If mn_MES_kgj_force = True Then
        md_MES_kgj = md_MES_kgj_f
    End If
    If mn_NNH4_kgj_force = True Then
        md_NNH4_kgj = md_NNH4_kgj_f
    End If
    If mn_NK_kgj_force = True Then
        RetourFluxAdmisTraitementSoluble(repere_ret_NK) = RetourFluxAdmisTraitementSoluble(repere_ret_NK) * md_NK_kgj_f / md_NK_kgj
        If RetourFluxAdmisTraitementSoluble(repere_ret_NK) < md_NNH4_kgj Then
            RetourFluxAdmisTraitementSoluble(repere_ret_NK) = md_NNH4_kgj
        End If
        md_NK_kgj = md_NK_kgj_f
    End If
    If mn_NNO3_kgj_force = True Then
        md_NNO3_kgj = md_NNO3_kgj_f
    End If
    If mn_Pt_kgj_force = True Then
        md_Pt_kgj = md_Pt_kgj_f
    End If
    If mn_PPO4_kgj_force = True Then
        md_PPO4_kgj = md_PPO4_kgj_f
        RetourFluxAdmisTraitementSoluble(repere_ret_Pt) = md_PPO4_kgj
    End If

    'On calcule la concentration
    pd_DCO_mgL = md_DCO_kgj / pd_Q_m3j * 1000
    pd_DBO_mgL = md_DBO_kgj / pd_Q_m3j * 1000
    pd_MES_mgL = md_MES_kgj / pd_Q_m3j * 1000
    pd_NK_mgL = md_NK_kgj / pd_Q_m3j * 1000
    pd_NNH4_mgL = md_NNH4_kgj / pd_Q_m3j * 1000
    pd_NNO3_mgL = md_NNO3_kgj / pd_Q_m3j * 1000
    pd_Pt_mgL = md_Pt_kgj / pd_Q_m3j * 1000
    pd_PPO4_mgL = md_PPO4_kgj / pd_Q_m3j * 1000


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "5.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub
Public Sub flux_sortie_process(ByRef RetourFlux() As Double, ByRef RetourFluxSoluble() As Double, ByRef RetourFluxDigestion() As Double, ByRef RetourFluxDigestionSoluble() As Double, ByRef RetourFluxAthos() As Double, ByRef RetourFluxAthosSoluble() As Double, ByRef RetourFluxNonAdmisTraitement() As Double, ByRef RetourFluxNonAdmisTraitementSoluble() As Double, ByRef RetourFluxAdmisTraitement() As Double, ByRef RetourFluxAdmisTraitementSoluble() As Double)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer


    'retours globaux
    For I = repere_ret_Q To repere_ret_Pt
        RetourFlux(I) = RetourFluxNonAdmisTraitement(I) + RetourFluxAdmisTraitement(I)
    Next I

    'retours solubles globaux
    RetourFluxSoluble(repere_ret_Q) = RetourFluxNonAdmisTraitementSoluble(repere_ret_Q)
    For I = repere_ret_DCO To repere_ret_Pt
        RetourFluxSoluble(I) = RetourFluxNonAdmisTraitementSoluble(I) + RetourFluxAdmisTraitementSoluble(I)
    Next I

    'On calcule les flux qui rentrent sur le traitement séparé et ce qui le by-passe
    If mi_ChoixRetoursAdmisTraitement = mi_JUS_DIGESTION Then
        For I = 1 To retour_caracteristique_nb
            RetourFluxDigestionSoluble(I) = RetourFluxAdmisTraitementSoluble(I)
            RetourFluxDigestion(I) = RetourFluxAdmisTraitement(I)
        Next I
    ElseIf mi_ChoixRetoursAdmisTraitement = mi_JUS_ATHOS Then
        For I = 1 To retour_caracteristique_nb
            RetourFluxAthosSoluble(I) = RetourFluxAdmisTraitementSoluble(I)
            RetourFluxAthos(I) = RetourFluxAdmisTraitement(I)
        Next I
    Else
        'On répartit par proportion entrante
        For I = 1 To retour_caracteristique_nb
            If (RetourFluxDigestionSoluble(I) + RetourFluxAthosSoluble(I)) > 0 Then
                RetourFluxDigestionSoluble(I) = RetourFluxDigestionSoluble(I) / (RetourFluxDigestionSoluble(I) + RetourFluxAthosSoluble(I)) * RetourFluxAdmisTraitementSoluble(I)
                RetourFluxAthosSoluble(I) = RetourFluxAdmisTraitementSoluble(I) - RetourFluxDigestionSoluble(I)
            Else
                RetourFluxDigestionSoluble(I) = 0
                RetourFluxAthosSoluble(I) = 0
            End If
            If (RetourFluxDigestion(I) + RetourFluxAthos(I)) > 0 Then
                RetourFluxDigestion(I) = RetourFluxDigestion(I) / (RetourFluxDigestion(I) + RetourFluxAthos(I)) * RetourFluxAdmisTraitement(I)
                RetourFluxAthos(I) = RetourFluxAdmisTraitement(I) - RetourFluxDigestion(I)
            Else
                RetourFluxDigestion(I) = 0
                RetourFluxAthos(I) = 0
            End If
        Next I
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "5.13"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Public Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree
    Ligne = 1

    'On écrit les résultats sur la composition des retours
    Feuil6.Range("retours_admis_Q_m3j").Cells(Ligne, Colonne) = ""      'pd_Q_m3j
    Feuil6.Range("retours_admis_DCO_mgL").Cells(Ligne, Colonne) = ""      'pd_DCO_mgL
    Feuil6.Range("retours_admis_DBO_mgL").Cells(Ligne, Colonne) = ""      'pd_DBO_mgL
    Feuil6.Range("retours_admis_MES_mgL").Cells(Ligne, Colonne) = ""      'pd_MES_mgL
    Feuil6.Range("retours_admis_NK_mgL").Cells(Ligne, Colonne) = ""      'pd_NK_mgL
    Feuil6.Range("retours_admis_NNH4_mgL").Cells(Ligne, Colonne) = ""      'pd_NNH4_mgL
    Feuil6.Range("retours_admis_NNO3_mgL").Cells(Ligne, Colonne) = ""      'pd_NNO3_mgL
    Feuil6.Range("retours_admis_Pt_mgL").Cells(Ligne, Colonne) = ""      'pd_Pt_mgL
    Feuil6.Range("retours_admis_PPO4_mgL").Cells(Ligne, Colonne) = ""      'pd_PPO4_mgL


    Feuil6.Range("retours_admis_DCO_kgj").Cells(Ligne, Colonne) = ""      'md_DCO_kgj
    Feuil6.Range("retours_admis_DBO_kgj").Cells(Ligne, Colonne) = ""      'md_DBO_kgj
    Feuil6.Range("retours_admis_MES_kgj").Cells(Ligne, Colonne) = ""      'md_MES_kgj
    Feuil6.Range("retours_admis_NK_kgj").Cells(Ligne, Colonne) = ""      'md_NK_kgj
    Feuil6.Range("retours_admis_NNH4_kgj").Cells(Ligne, Colonne) = ""      'md_NNH4_kgj
    Feuil6.Range("retours_admis_NNO3_kgj").Cells(Ligne, Colonne) = ""      'md_NNO3_kgj
    Feuil6.Range("retours_admis_Pt_kgj").Cells(Ligne, Colonne) = ""      'md_Pt_kgj
    Feuil6.Range("retours_admis_PPO4_kgj").Cells(Ligne, Colonne) = ""      'md_PPO4_kgj


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "5.10"

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
        Call ecriture_resultats
        Call affich_remarque
    End If

End Sub


Private Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "5.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

