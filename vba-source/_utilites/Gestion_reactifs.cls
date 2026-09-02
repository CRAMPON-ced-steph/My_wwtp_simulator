Attribute VB_Name = "Gestion_reactifs"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'constantes pour repérer les type de data
Private Const data_nb = 3
Private Const produit_pur = 1
Private Const purete_pdt_commercial = 2
Private Const produit_commercial = 3

'variables à échanger avec l'utilisateur
Dim DATA(reactifs_nb, data_nb) As Double
Dim data_force(reactifs_nb, data_nb) As Boolean
Dim data_f(reactifs_nb, data_nb) As Double

'hypothèses
Dim hyp_purete_reactif_commercial(reactifs_nb) As Double   'chaque valeur dans l'unité la plus adéquate

'autres variables
Dim purete_retenue(reactifs_nb) As Double

'Pureté Nutriox --> Utilisation pour le calcul de la consommation
Public pd_PureteNitrateDeCalcium As Double

Public Sub initialisation_procede()

    If pn_EffacementResultatsEnCours = False Then
        Call lecture_choix
        Call hypothese
        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
            Call lecture_valeurs_forcees
        End If
        Call attribution_valeur_par_defaut
    End If

End Sub

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer
    Dim Ligne As Integer
    Dim Colonne As Integer


    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    'on écrit les valeurs
    For I = 1 To reactifs_nb
        For j = 1 To data_nb
            Ligne = data_nb * (I - 1) + j
            If DATA(I, produit_pur) <> 0 Then
                Feuil54.Range("reactifs_data").Cells(Ligne, Colonne) = DATA(I, j)
            Else
                Feuil54.Range("reactifs_data").Cells(Ligne, Colonne) = "'-"
            End If
        Next j
    Next I

    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
        Call ecriture_resultats_OPEX
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "54.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub ecriture_resultats_OPEX()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim Colonne As Integer


    Colonne = 3 * pi_CALCUL_VALEUR_GUIDE + pi_FiliereConsideree

    For I = 1 To reactifs_nb
        If I = reactifs_repere_H2SO4 Or I = reactifs_repere_NaOH Or I = reactifs_repere_NaOCl_desodo Or I = reactifs_repere_NaHSO3 Then
            If DATA(I, produit_commercial) <> 0 Then
                Feuil55.Range("OPEX_reactifs_fixe").Cells(I, Colonne) = DATA(I, produit_commercial)
            Else
                Feuil55.Range("OPEX_reactifs_fixe").Cells(I, Colonne) = "'-"
            End If
            Feuil55.Range("OPEX_reactifs_variable").Cells(I, Colonne) = "'-"
        Else
            If DATA(I, produit_commercial) <> 0 Then
                Feuil55.Range("OPEX_reactifs_variable").Cells(I, Colonne) = DATA(I, produit_commercial)
            Else
                Feuil55.Range("OPEX_reactifs_variable").Cells(I, Colonne) = "'-"
            End If
            Feuil55.Range("OPEX_reactifs_fixe").Cells(I, Colonne) = "'-"
        End If
    Next I




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "54.11"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats_OPEX()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim Colonne As Integer


    Colonne = 3 * pi_CALCUL_VALEUR_GUIDE + pi_FiliereConsideree

    For I = 1 To reactifs_nb
        Feuil55.Range("OPEX_reactifs_fixe").Cells(I, Colonne) = ""
    Next I
    For I = 1 To reactifs_nb
        Feuil55.Range("OPEX_reactifs_variable").Cells(I, Colonne) = ""
    Next I




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "54.12"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer
    Dim Ligne As Integer
    Dim Colonne As Integer


    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    'on écrit les valeurs
    For I = 1 To reactifs_nb
        For j = 1 To data_nb
            Ligne = data_nb * (I - 1) + j
            Feuil54.Range("reactifs_data").Cells(Ligne, Colonne) = ""     'data(i, j)
        Next j
    Next I

    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
        Call effacer_resultats_OPEX
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "54.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'pas de choix en théorie



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "54.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer


    'pureté du réactif commercial retenu
    For I = 1 To reactifs_nb
        If data_force(I, purete_pdt_commercial) = True Then
            DATA(I, purete_pdt_commercial) = data_f(I, purete_pdt_commercial)
        Else
            DATA(I, purete_pdt_commercial) = hyp_purete_reactif_commercial(I)
        End If
    Next I
    pd_PureteNitrateDeCalcium = DATA(reactifs_repere_Ca_2NO3, purete_pdt_commercial)
    
    'on convertit les valeurs pour qu'elles soient directement utilisable sous la forme commercial=pur/pureté
    For I = 1 To reactifs_nb
        purete_retenue(I) = DATA(I, purete_pdt_commercial)
    Next I



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "54.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub

Sub hypothese()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'hypothèses de pureté du réactif commercial    (puretés vues avec Jean D'ARIES)
    hyp_purete_reactif_commercial(reactifs_repere_polymere_eau_poudre_anionique) = 0.98   'pur
    hyp_purete_reactif_commercial(reactifs_repere_polymere_eau_poudre_cationique) = 0.98   'pur
    hyp_purete_reactif_commercial(reactifs_repere_polymere_eau_liquide_anionique) = 0.5  'pur
    hyp_purete_reactif_commercial(reactifs_repere_polymere_eau_liquide_cationique) = 0.5   'pur
    hyp_purete_reactif_commercial(reactifs_repere_polymere_boues_poudre_anionique) = 0.98   'pur
    hyp_purete_reactif_commercial(reactifs_repere_polymere_boues_poudre_cationique) = 0.98   'pur
    hyp_purete_reactif_commercial(reactifs_repere_polymere_boues_liquide_anionique) = 0.5   'pur
    hyp_purete_reactif_commercial(reactifs_repere_polymere_boues_liquide_cationique) = 0.5   'pur
    hyp_purete_reactif_commercial(reactifs_repere_FeCl3_eau) = 0.41      '% en masse
    hyp_purete_reactif_commercial(reactifs_repere_FeCl3_boues) = 0.41      '% en masse
    hyp_purete_reactif_commercial(reactifs_repere_methanol) = 1    'pur
    hyp_purete_reactif_commercial(reactifs_repere_chaux_eteinte) = 0.9
    hyp_purete_reactif_commercial(reactifs_repere_chaux_vive) = 0.95
    hyp_purete_reactif_commercial(reactifs_repere_H2SO4) = 0.98      '% en masse   en France 98%, à l'étranger 96%
    hyp_purete_reactif_commercial(reactifs_repere_NaOH) = 0.3      '% en masse
    hyp_purete_reactif_commercial(reactifs_repere_NaOCl_eau) = 48    'en °Beaumé
    hyp_purete_reactif_commercial(reactifs_repere_NaOCl_desodo) = 48    'en °Beaumé
    'pour "reactifs_repere_NaOCl", calcul de la pureté massique dans la sub calcul  (donnée Eric GUIBELIN)
    hyp_purete_reactif_commercial(reactifs_repere_NaHSO3) = 0.35      '% en masse
    hyp_purete_reactif_commercial(reactifs_repere_NaHCO3) = 1
    hyp_purete_reactif_commercial(reactifs_repere_Ca_2NO3) = 0.45    'recherche internet "Caractéristiques Nutriox" --> page trouvée sur le site de Yara
    hyp_purete_reactif_commercial(reactifs_repere_O2_liquide) = 1    'pur
    hyp_purete_reactif_commercial(reactifs_repere_H2O2) = 0.3    'Eric GUIBELIN
    hyp_purete_reactif_commercial(reactifs_repere_charbon_actif) = 1  'en poudre
    hyp_purete_reactif_commercial(reactifs_repere_CuSO4) = 0.98    'A DEMANDER
    hyp_purete_reactif_commercial(reactifs_repere_MgCl2) = 0.3    'Christophe MELE
    hyp_purete_reactif_commercial(reactifs_repere_Cl2_gazeux) = 1    'pur
    hyp_purete_reactif_commercial(reactifs_repere_Ammoniaque) = 0.25    'solution
    hyp_purete_reactif_commercial(reactifs_repere_Uree) = 0.45    'solution




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "54.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub



Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    Dim j As Integer
    Dim Ligne As Integer
    Dim Colonne As Integer



    Colonne = pi_FiliereConsideree

    'on lit les valeurs forcées
    For I = 1 To reactifs_nb
        For j = 1 To data_nb
            If j = purete_pdt_commercial Then
                Ligne = data_nb * (I - 1) + j
                data_force(I, j) = Feuil54.Range("reactifs_data").Cells(Ligne, Colonne)
                If data_force(I, j) = True Then
                    data_f(I, j) = Feuil54.Range("reactifs_data").Cells(Ligne, Colonne)
                End If
            Else
                data_force(I, j) = False
            End If
        Next j
    Next I





    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "54.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub calcul(ByVal polymere_eau_flux, ByVal polymere_boues_flux, ByVal chaux_eteinte_flux, ByVal chaux_vive_flux, ByVal FeCl3_eau_flux, ByVal FeCl3_boues_flux, ByVal methanol_flux, ByVal H2SO4pur_flux, ByVal NaOHpur_flux, ByVal NaOClpur_eau_flux, ByVal NaOClpur_desodo_flux, ByVal NaHSO3pur_flux, ByVal NaHCO3_flux, ByVal O2_liquide_flux, ByVal charbon_actif_flux, ByVal Cl2_gazeux_flux_kgj, ByVal Ca_2NO3_pur_flux, ByVal H2O2pur_flux, ByVal CuSO4_flux, ByVal MgCl2_flux, ByVal ConsommationAmmoniaquePur_kgj As Double, ByVal ConsommationUreePur_kgj As Double)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer


    DATA(reactifs_repere_polymere_eau_poudre_anionique, produit_pur) = polymere_eau_flux(polymere_poudre_anion)
    DATA(reactifs_repere_polymere_eau_poudre_cationique, produit_pur) = polymere_eau_flux(polymere_poudre_cation)
    DATA(reactifs_repere_polymere_eau_liquide_anionique, produit_pur) = polymere_eau_flux(polymere_liquide_anion)
    DATA(reactifs_repere_polymere_eau_liquide_cationique, produit_pur) = polymere_eau_flux(polymere_liquide_cation)

    DATA(reactifs_repere_polymere_boues_poudre_anionique, produit_pur) = polymere_boues_flux(polymere_poudre_anion)
    DATA(reactifs_repere_polymere_boues_poudre_cationique, produit_pur) = polymere_boues_flux(polymere_poudre_cation)
    DATA(reactifs_repere_polymere_boues_liquide_anionique, produit_pur) = polymere_boues_flux(polymere_liquide_anion)
    DATA(reactifs_repere_polymere_boues_liquide_cationique, produit_pur) = polymere_boues_flux(polymere_liquide_cation)

    DATA(reactifs_repere_FeCl3_eau, produit_pur) = FeCl3_eau_flux
    DATA(reactifs_repere_FeCl3_boues, produit_pur) = FeCl3_boues_flux
    DATA(reactifs_repere_methanol, produit_pur) = methanol_flux
    DATA(reactifs_repere_chaux_eteinte, produit_pur) = chaux_eteinte_flux
    DATA(reactifs_repere_chaux_vive, produit_pur) = chaux_vive_flux
    DATA(reactifs_repere_H2SO4, produit_pur) = H2SO4pur_flux
    DATA(reactifs_repere_NaOH, produit_pur) = NaOHpur_flux

    DATA(reactifs_repere_NaOCl_eau, produit_pur) = NaOClpur_eau_flux
    DATA(reactifs_repere_NaOCl_desodo, produit_pur) = NaOClpur_desodo_flux

    DATA(reactifs_repere_NaHSO3, produit_pur) = NaHSO3pur_flux
    DATA(reactifs_repere_NaHCO3, produit_pur) = NaHCO3_flux
    DATA(reactifs_repere_Ca_2NO3, produit_pur) = Ca_2NO3_pur_flux
    DATA(reactifs_repere_O2_liquide, produit_pur) = O2_liquide_flux
    DATA(reactifs_repere_H2O2, produit_pur) = H2O2pur_flux
    DATA(reactifs_repere_CuSO4, produit_pur) = CuSO4_flux
    DATA(reactifs_repere_MgCl2, produit_pur) = MgCl2_flux
    DATA(reactifs_repere_charbon_actif, produit_pur) = charbon_actif_flux
    DATA(reactifs_repere_Cl2_gazeux, produit_pur) = Cl2_gazeux_flux_kgj
    DATA(reactifs_repere_Ammoniaque, produit_pur) = ConsommationAmmoniaquePur_kgj
    DATA(reactifs_repere_Uree, produit_pur) = ConsommationUreePur_kgj




    For I = 1 To reactifs_nb
        If I = reactifs_repere_NaOCl_eau Or I = reactifs_repere_NaOCl_desodo Then
            DATA(I, produit_commercial) = DATA(I, produit_pur) / (0.152 * purete_retenue(I) / 48)
        Else
            DATA(I, produit_commercial) = DATA(I, produit_pur) / purete_retenue(I)
        End If
    Next I



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "54.12"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Private Sub Class_Terminate()


    If pn_EffacementResultatsEnCours = False Then
        Call ecriture_resultats
    End If


End Sub







