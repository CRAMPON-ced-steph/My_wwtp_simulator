Attribute VB_Name = "Utilites"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Const electricite_autre_non = 0

Dim utilites_debit_eau_service As Double
Dim utilites_debit_retours As Double
Dim utilites_pression_refoulement_eau_service As Double
Dim utilites_pression_refoulement_retours As Double

Dim utilites_rendement_pompes_eau_service As Double
Dim utilites_rendement_pompes_retours As Double

Dim utilites_consommation_electrique_eau_service As Double
Dim utilites_consommation_electrique_retours As Double

Dim utilites_consommation_electrique_eclairage As Double
Dim utilites_eclairage_pourcentage As Double

Dim utilites_pertes_enligne As Double
Dim utilites_pertes_pourcentage As Double

Dim utilites_debit_eau_service_force As Boolean
Dim utilites_pression_refoulement_eau_service_force As Boolean
Dim utilites_pression_refoulement_retours_force As Boolean
Dim utilites_rendement_pompes_eau_service_force As Boolean
Dim utilites_rendement_pompes_retours_force As Boolean
Dim utilites_eclairage_pourcentage_force As Boolean
Dim utilites_pertes_pourcentage_force As Boolean

Dim utilites_debit_eau_service_f As Double
Dim utilites_pression_refoulement_eau_service_f As Double
Dim utilites_pression_refoulement_retours_f As Double
Dim utilites_rendement_pompes_eau_service_f As Double
Dim utilites_rendement_pompes_retours_f As Double
Dim utilites_eclairage_pourcentage_f As Double
Dim utilites_pertes_pourcentage_f As Double

Dim debit As Double
Dim pourcentage_charge_hydraulic As Double
Dim calcul_consommation_electrique_total As Double

Sub lecture_electricite_autre(ByRef electricite_autre)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim choix_electricite_autre As Integer

    choix_electricite_autre = Feuil6.Range("choix_electricite_autre").Cells(1, pi_FiliereConsideree)
    electricite_autre = 0
    If choix_electricite_autre <> electricite_autre_non Then
        For I = 1 To choix_electricite_autre
            electricite_autre = electricite_autre + Feuil6.Range("electricite_autre").Cells(I, pi_FiliereConsideree)
        Next I
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "39.14"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub
Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    utilites_pression_refoulement_eau_service_force = Feuil39.Range("utilites_pression_refoulement_eau_service").Cells(1, pi_FiliereConsideree)
    utilites_rendement_pompes_eau_service_force = Feuil39.Range("utilites_rendement_pompes_eau_service").Cells(1, pi_FiliereConsideree)
    utilites_debit_eau_service_force = Feuil39.Range("utilites_debit_eau_service").Cells(1, pi_FiliereConsideree)
    utilites_pression_refoulement_retours_force = Feuil39.Range("utilites_pression_refoulement_retours").Cells(1, pi_FiliereConsideree)
    utilites_rendement_pompes_retours_force = Feuil39.Range("utilites_rendement_pompes_retours").Cells(1, pi_FiliereConsideree)
    utilites_eclairage_pourcentage_force = Feuil39.Range("utilites_eclairage_pourcentage").Cells(1, pi_FiliereConsideree)
    utilites_pertes_pourcentage_force = Feuil39.Range("utilites_pertes_pourcentage").Cells(1, pi_FiliereConsideree)

    If utilites_pression_refoulement_eau_service_force = True Then
        utilites_pression_refoulement_eau_service_f = Feuil39.Range("utilites_pression_refoulement_eau_service").Cells(1, pi_FiliereConsideree)
    End If

    If utilites_rendement_pompes_eau_service_force = True Then
        utilites_rendement_pompes_eau_service_f = Feuil39.Range("utilites_rendement_pompes_eau_service").Cells(1, pi_FiliereConsideree)
    End If

    If utilites_debit_eau_service_force = True Then
        utilites_debit_eau_service_f = Feuil39.Range("utilites_debit_eau_service").Cells(1, pi_FiliereConsideree)
    End If

    If utilites_pression_refoulement_retours_force = True Then
        utilites_pression_refoulement_retours_f = Feuil39.Range("utilites_pression_refoulement_retours").Cells(1, pi_FiliereConsideree)
    End If

    If utilites_rendement_pompes_retours_force = True Then
        utilites_rendement_pompes_retours_f = Feuil39.Range("utilites_rendement_pompes_retours").Cells(1, pi_FiliereConsideree)
    End If

    If utilites_eclairage_pourcentage_force = True Then
        utilites_eclairage_pourcentage_f = Feuil39.Range("utilites_eclairage_pourcentage").Cells(1, pi_FiliereConsideree)
    End If

    If utilites_pertes_pourcentage_force = True Then
        utilites_pertes_pourcentage_f = Feuil39.Range("utilites_pertes_pourcentage").Cells(1, pi_FiliereConsideree)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "39.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub calcul_eau_service(ByRef electricite_consommee, ByRef boues_flux, ByRef retour_flux, ByRef electricite_utilites)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'guide_defaut
    If utilites_pression_refoulement_eau_service_force = True Then
        utilites_pression_refoulement_eau_service = utilites_pression_refoulement_eau_service_f
    Else
        utilites_pression_refoulement_eau_service = 60
    End If

    If utilites_rendement_pompes_eau_service_force = True Then
        utilites_rendement_pompes_eau_service = utilites_rendement_pompes_eau_service_f
    Else
        utilites_rendement_pompes_eau_service = 0.7
    End If

    'debit_eau_service
    debit = Q_nominal
    pourcentage_charge_hydraulic = pd_PourcentageChargeReelleDebitVolumique

    If utilites_debit_eau_service_force = True Then
        utilites_debit_eau_service = utilites_debit_eau_service_f
    Else
        utilites_debit_eau_service = 0.01 * debit * pourcentage_charge_hydraulic
    End If

    'utilites_debit_eau_service
    utilites_consommation_electrique_eau_service = utilites_debit_eau_service / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * utilites_pression_refoulement_eau_service / utilites_rendement_pompes_eau_service
    electricite_consommee = electricite_consommee + utilites_consommation_electrique_eau_service
    electricite_utilites = electricite_utilites + utilites_consommation_electrique_eau_service

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "39.12a"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub
Sub calcul_retour(ByRef electricite_consommee, ByRef boues_flux, ByRef retour_flux, ByRef electricite_utilites)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Q_retour As Double

    'guide_defaut
    If utilites_pression_refoulement_retours_force = True Then
        utilites_pression_refoulement_retours = utilites_pression_refoulement_retours_f
    Else
        utilites_pression_refoulement_retours = 5
    End If

    If utilites_rendement_pompes_retours_force = True Then
        utilites_rendement_pompes_retours = utilites_rendement_pompes_retours_f
    Else
        utilites_rendement_pompes_retours = 0.7
    End If

    'utilites_debit_retours
    utilites_debit_retours = retour_flux(repere_ret_Q)
    utilites_consommation_electrique_retours = utilites_debit_retours / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * utilites_pression_refoulement_retours / utilites_rendement_pompes_retours
    electricite_consommee = electricite_consommee + utilites_consommation_electrique_retours
    electricite_utilites = electricite_utilites + utilites_consommation_electrique_retours

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "39.12d"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub
Sub calcul_eclairage(ByRef electricite_consommee, ByRef electricite_utilites)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'guide_defaut
    If utilites_eclairage_pourcentage_force = True Then
        utilites_eclairage_pourcentage = utilites_eclairage_pourcentage_f
    Else
        utilites_eclairage_pourcentage = 0.01
    End If

    'utilites_pertes_enligne
    utilites_consommation_electrique_eclairage = utilites_eclairage_pourcentage * electricite_consommee
    electricite_consommee = electricite_consommee + utilites_consommation_electrique_eclairage
    electricite_utilites = electricite_utilites + utilites_consommation_electrique_eclairage

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "39.12b"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub calcul_pertes(ByRef electricite_consommee, ByRef electricite_utilites)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim electricite_autre As Double
    Call lecture_electricite_autre(electricite_autre)

    'guide_defaut
    If utilites_pertes_pourcentage_force = True Then
        utilites_pertes_pourcentage = utilites_pertes_pourcentage_f
    Else
        utilites_pertes_pourcentage = 0.05
    End If

    'utilites_consommation_electrique_eclairage
    calcul_consommation_electrique_total = electricite_consommee + electricite_autre
    utilites_pertes_enligne = utilites_pertes_pourcentage * calcul_consommation_electrique_total
    electricite_consommee = electricite_consommee + utilites_pertes_enligne
    electricite_utilites = electricite_utilites + utilites_pertes_enligne

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "39.12c"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub ecriture_resultats_eau_service()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil39.Range("utilites_debit_eau_service").Cells(1, Colonne) = utilites_debit_eau_service
    Feuil39.Range("utilites_pression_refoulement_eau_service").Cells(1, Colonne) = utilites_pression_refoulement_eau_service
    Feuil39.Range("utilites_rendement_pompes_eau_service").Cells(1, Colonne) = utilites_rendement_pompes_eau_service
    Feuil39.Range("utilites_consommation_electrique_eau_service").Cells(1, Colonne) = utilites_consommation_electrique_eau_service

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "39.9a"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub
Sub ecriture_resultats_retour()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil39.Range("utilites_pression_refoulement_retours").Cells(1, Colonne) = utilites_pression_refoulement_retours
    Feuil39.Range("utilites_debit_retours").Cells(1, Colonne) = utilites_debit_retours
    Feuil39.Range("utilites_rendement_pompes_retours").Cells(1, Colonne) = utilites_rendement_pompes_retours
    Feuil39.Range("utilites_consommation_electrique_retours").Cells(1, Colonne) = utilites_consommation_electrique_retours

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "39.9d"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub
Sub ecriture_resultats_eclairage()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree
    Feuil39.Range("utilites_eclairage_pourcentage").Cells(1, Colonne) = utilites_eclairage_pourcentage
    Feuil39.Range("utilites_consommation_electrique_eclairage").Cells(1, Colonne) = utilites_consommation_electrique_eclairage

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "39.9b"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub ecriture_resultats_pertes()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree
    Feuil39.Range("utilites_pertes_pourcentage").Cells(1, Colonne) = utilites_pertes_pourcentage
    Feuil39.Range("utilites_pertes_enligne").Cells(1, Colonne) = utilites_pertes_enligne

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "39.9c"

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
    Feuil39.Range("utilites_debit_eau_service").Cells(Ligne, Colonne) = ""
    Feuil39.Range("utilites_debit_retours").Cells(Ligne, Colonne) = ""
    Feuil39.Range("utilites_pression_refoulement_eau_service").Cells(Ligne, Colonne) = ""
    Feuil39.Range("utilites_pression_refoulement_retours").Cells(Ligne, Colonne) = ""
    Feuil39.Range("utilites_rendement_pompes_eau_service").Cells(Ligne, Colonne) = ""
    Feuil39.Range("utilites_rendement_pompes_retours").Cells(Ligne, Colonne) = ""
    Feuil39.Range("utilites_consommation_electrique_eau_service").Cells(Ligne, Colonne) = ""
    Feuil39.Range("utilites_consommation_electrique_retours").Cells(Ligne, Colonne) = ""
    Feuil39.Range("utilites_eclairage_pourcentage").Cells(Ligne, Colonne) = ""
    Feuil39.Range("utilites_consommation_electrique_eclairage").Cells(Ligne, Colonne) = ""
    Feuil39.Range("utilites_pertes_pourcentage").Cells(Ligne, Colonne) = ""
    Feuil39.Range("utilites_pertes_enligne").Cells(Ligne, Colonne) = ""

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "39.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 0  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil39.Range("utilites_erreur_nom")

    For I = 1 To nb_remarque
        'If remarque(i) <> "" Then
        ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf    '+ remarque(i)
        'End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "39.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Private Sub Class_Terminate()

    If pn_EffacementResultatsEnCours = False Then
        Call affich_remarque
    End If

End Sub

