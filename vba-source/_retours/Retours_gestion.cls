Attribute VB_Name = "Retours_gestion"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Dim retours_flux_force(retour_caracteristique_nb) As Boolean
Dim retours_flux_f(retour_caracteristique_nb) As Double

Sub forcer_flux(ByRef retours_flux)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer

    For I = 1 To retour_caracteristique_nb
        If retours_flux_force(I) = True Then
            retours_flux(I) = retours_flux_f(I)
        End If
    Next I


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "6.15"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim Colonne As Integer

    Colonne = pi_FiliereConsideree

    For I = 1 To retour_caracteristique_nb
        retours_flux_force(I) = Feuil6.Range("retours_flux").Cells(I, Colonne)
        If retours_flux_force(I) = True Then
            retours_flux_f(I) = Feuil6.Range("retours_flux").Cells(I, Colonne)
        End If
    Next I

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "6.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Sub Class_Initialize()

    Call lecture_valeurs_forcees

End Sub

