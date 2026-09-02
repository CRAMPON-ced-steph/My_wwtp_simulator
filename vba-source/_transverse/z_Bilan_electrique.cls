Attribute VB_Name = "z_Bilan_electrique"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Const electricite_autre_non = 0
Sub ecrire_resultats(ByVal DCO, ByVal DBO, ByRef electricite_consommee, ByVal electricite_fixe, ByVal electricite_variable, ByVal electricite_II_aeration, ByVal electricite_III_aeration, ByVal electricite_verte, ByVal electricite_verte_biogaz, ByVal electricite_produite_solaire_photovoltaique, ByVal electricite_produite_turbine_hydraulique, ByVal electricite_produite_autre)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer

    Dim electricite_autre As Double
    Dim electricite_autosuffisance As Double
    Dim electricite_autosuffisance_biogaz As Double
    Dim electricite_file_eau As Double
    Dim electricite_file_boues As Double
    Dim ratio_Q As Double
    Dim ratio_Q_usuel As String
    Dim ratio_DCOremoved As Double
    Dim ratio_DCOremoved_usuel As String
    Dim ratio_EH As Double
    Dim ratio_EH_usuel As String
    Dim ratio_DBOremoved As Double
    Dim ratio_DBOremoved_usuel As String
    Dim Colonne As Integer

    Dim ratio_kWh_kgDBO_min As Double
    Dim ratio_kWh_kgDBO_max As Double
    Dim ratio_min As Integer
    Dim ratio_max As Integer



    Call lecture_electricite_autre(electricite_autre)
    electricite_fixe(electricite_postes_repere_autres) = electricite_fixe(electricite_postes_repere_autres)
    electricite_variable(electricite_postes_repere_autres) = electricite_variable(electricite_postes_repere_autres) + electricite_autre
    electricite_consommee = electricite_consommee + electricite_autre

    'Calculs
    electricite_file_eau = 0
    For I = 1 To electricite_postes_eau_nb
        electricite_file_eau = electricite_file_eau + electricite_fixe(I) + electricite_variable(I)
    Next I
    electricite_file_boues = 0
    For I = electricite_postes_eau_nb + 1 To electricite_postes_boues_nb
        electricite_file_boues = electricite_file_boues + electricite_fixe(I) + electricite_variable(I)
    Next I

    '    If choix_electricite_autre_production = True Then
    '        colonne =
    '        electricite_produite_autre = Feuil6.Range("electricite_autre_production").Cells(1, 3 * pi_FiliereConsideree)
    '    End If

    If electricite_consommee > 0 Then
        'modifié par DCA 7/11/2012 pour prendre en compte dans l'empreinte CO2 et les OPEX, electricite_verte comprend déjà toutes les productions
        'electricite_verte = electricite_verte + electricite_produite_solaire_photovoltaique + electricite_produite_turbine_hydraulique + electricite_produite_autre ' electricite_verte_biogaz a déjà été ajouté au bilan dans gestion_energie_thermique (à déplacer par la suite)
        electricite_autosuffisance = electricite_verte / electricite_consommee
        electricite_autosuffisance_biogaz = electricite_verte_biogaz / electricite_consommee
    Else
        electricite_autosuffisance = 0
        electricite_autosuffisance_biogaz = 0
    End If

    If choix_primaire = False And choix_secondaire = False And choix_tertiaire = False Then
        ratio_Q = 0
        ratio_DCOremoved = 0
        ratio_EH = electricite_consommee * pd_NOMBRE_JOUR_PAR_AN / (pd_CapaciteSTEP_EH * pd_PourcentageChargeReelleDBO)   'ratio basé sur  1 EH <--> 60 gDBO/j
        ratio_DBOremoved = 0
    Else
        ratio_Q = electricite_consommee / (Q_nominal * pd_PourcentageChargeReelleDebitVolumique)
        ratio_DCOremoved = electricite_consommee / (DCO_nominal * pd_PourcentageChargeReelleDCO - DCO)
        ratio_EH = electricite_consommee * pd_NOMBRE_JOUR_PAR_AN / (pd_CapaciteSTEP_EH * pd_PourcentageChargeReelleDBO)   'ratio basé sur  1 EH <--> 60 gDBO/j
        ratio_DBOremoved = electricite_consommee / (DBO_nominal * pd_PourcentageChargeReelleDBO - DBO)
    End If


    'Report des valeurs
    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree
    Feuil6.Range("electricite_total").Cells(1, Colonne) = electricite_consommee
    Feuil6.Range("electricite_verte").Cells(1, Colonne) = electricite_verte
    Feuil6.Range("electricite_autosuffisance").Cells(1, Colonne) = electricite_autosuffisance
    Feuil6.Range("electricite_autosuffisance_biogaz").Cells(1, Colonne) = electricite_autosuffisance_biogaz

    'Modif MSM 25/10/12
    Feuil6.Range("electricite_verte_biogaz").Cells(1, Colonne) = electricite_verte_biogaz
    Feuil6.Range("electricite_produite_solaire_photovoltaique").Cells(1, Colonne) = electricite_produite_solaire_photovoltaique
    Feuil6.Range("electricite_produite_turbine_hydraulique").Cells(1, Colonne) = electricite_produite_turbine_hydraulique
    Feuil6.Range("electricite_produite_autre").Cells(1, Colonne) = electricite_produite_autre
    'Fin modif MSM 25/10/12

    Feuil6.Range("electricite_file_eau").Cells(1, Colonne) = electricite_file_eau
    Feuil6.Range("electricite_file_boues").Cells(1, Colonne) = electricite_file_boues
    Feuil6.Range("electricite_secondaire_aeration_only").Cells(1, Colonne) = electricite_II_aeration
    Feuil6.Range("electricite_tertiaire_aeration_only").Cells(1, Colonne) = electricite_III_aeration

    Feuil6.Range("electricite_relevement").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_relevement) + electricite_variable(electricite_postes_repere_relevement)
    Feuil6.Range("electricite_pretraitement").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_pretraitement) + electricite_variable(electricite_postes_repere_pretraitement)
    Feuil6.Range("electricite_primaire").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_primaire) + electricite_variable(electricite_postes_repere_primaire)
    Feuil6.Range("electricite_secondaire").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_secondaire) + electricite_variable(electricite_postes_repere_secondaire)
    Feuil6.Range("electricite_tertiaire").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_tertiaire) + electricite_variable(electricite_postes_repere_tertiaire)
    Feuil6.Range("electricite_desinfection").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_desinfection) + electricite_variable(electricite_postes_repere_desinfection)
    Feuil6.Range("electricite_eaux_sales").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_eaux_sales) + electricite_variable(electricite_postes_repere_eaux_sales)
    Feuil6.Range("electricite_biolix").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_biolix) + electricite_variable(electricite_postes_repere_biolix)
    Feuil6.Range("electricite_epaississement").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_epaississement) + electricite_variable(electricite_postes_repere_epaississement)
    Feuil6.Range("electricite_digestion").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_digestion) + electricite_variable(electricite_postes_repere_digestion)
    Feuil6.Range("electricite_athos").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_athos) + electricite_variable(electricite_postes_repere_athos)
    Feuil6.Range("electricite_deshydratation").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_deshydratation) + electricite_variable(electricite_postes_repere_deshydratation)
    Feuil6.Range("electricite_chaulage").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_chaulage) + electricite_variable(electricite_postes_repere_chaulage)
    Feuil6.Range("electricite_sechage").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_sechage) + electricite_variable(electricite_postes_repere_sechage)
    Feuil6.Range("electricite_incineration").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_incineration) + electricite_variable(electricite_postes_repere_incineration)
    Feuil6.Range("electricite_traitement_retours").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_trait_retours) + electricite_variable(electricite_postes_repere_trait_retours)
    Feuil6.Range("electricite_file_odeur").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_desodorisation) + electricite_variable(electricite_postes_repere_desodorisation)
    Feuil6.Range("electricite_utilites").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_utilites) + electricite_variable(electricite_postes_repere_utilites)
    Feuil6.Range("electricite_autre_total").Cells(1, Colonne) = electricite_fixe(electricite_postes_repere_autres) + electricite_variable(electricite_postes_repere_autres)


    If choix_primaire = False And choix_secondaire = False And choix_tertiaire = False Then
        Feuil6.Range("electricite_ratio_Q").Cells(1, Colonne) = "-"
        Feuil6.Range("electricite_ratio_DCOremoved").Cells(1, Colonne) = "-"
        Feuil6.Range("electricite_ratio_EH").Cells(1, Colonne) = ratio_EH
        Feuil6.Range("electricite_ratio_DBOremoved").Cells(1, Colonne) = "-"
    Else
        Feuil6.Range("electricite_ratio_Q").Cells(1, Colonne) = ratio_Q
        Feuil6.Range("electricite_ratio_DCOremoved").Cells(1, Colonne) = ratio_DCOremoved
        Feuil6.Range("electricite_ratio_EH").Cells(1, Colonne) = ratio_EH
        Feuil6.Range("electricite_ratio_DBOremoved").Cells(1, Colonne) = ratio_DBOremoved
    End If

    'ATTENTION DANS LA FEUILLE ON ECRIT A LA PLACE DES VALEURS FORCEES
    If pi_ChoixQualiteRejet = pi_QUALITE_REJET_TRAITEMENT_CARBONE Then
        'DBO
        ratio_kWh_kgDBO_min = 0.8
        ratio_kWh_kgDBO_max = 1.2
        ratio_DBOremoved_usuel = "'" & ratio_kWh_kgDBO_min & "  -  " & ratio_kWh_kgDBO_max
        Feuil6.Range("electricite_ratio_DBOremoved_usuel").Cells(1, 3 + pi_FiliereConsideree) = ratio_kWh_kgDBO_max
        Feuil6.Range("electricite_ratio_DBOremoved_usuel").Cells(1, 6 + pi_FiliereConsideree) = ratio_kWh_kgDBO_min
        'EH
        ratio_min = 10 * (pd_REJET_DBO_PAR_HABITANT_ET_PAR_JOUR * pd_NOMBRE_JOUR_PAR_AN * ratio_kWh_kgDBO_min)
        ratio_max = 10 * (pd_REJET_DBO_PAR_HABITANT_ET_PAR_JOUR * pd_NOMBRE_JOUR_PAR_AN * ratio_kWh_kgDBO_max)
        ratio_EH_usuel = "'" & ratio_min / 10 & "  -  " & ratio_max / 10
        Feuil6.Range("electricite_ratio_EH_usuel").Cells(1, 3 + pi_FiliereConsideree) = ratio_max / 10
        Feuil6.Range("electricite_ratio_EH_usuel").Cells(1, 6 + pi_FiliereConsideree) = ratio_min / 10
        'DCO
        ratio_min = 100 * (pd_PourcentageChargeReelleDBO * DBO_nominal / (pd_PourcentageChargeReelleDCO * DCO_nominal) * ratio_kWh_kgDBO_min)
        ratio_max = 100 * (pd_PourcentageChargeReelleDBO * DBO_nominal / (pd_PourcentageChargeReelleDCO * DCO_nominal) * ratio_kWh_kgDBO_max)
        ratio_DCOremoved_usuel = "'" & ratio_min / 100 & "  -  " & ratio_max / 100
        Feuil6.Range("electricite_ratio_DCOremoved_usuel").Cells(1, 3 + pi_FiliereConsideree) = ratio_max / 100
        Feuil6.Range("electricite_ratio_DCOremoved_usuel").Cells(1, 6 + pi_FiliereConsideree) = ratio_min / 100
        'Q
        ratio_min = 100 * (pd_PourcentageChargeReelleDBO * DBO_nominal / (pd_PourcentageChargeReelleDebitVolumique * Q_nominal) * ratio_kWh_kgDBO_min)
        ratio_max = 100 * (pd_PourcentageChargeReelleDBO * DBO_nominal / (pd_PourcentageChargeReelleDebitVolumique * Q_nominal) * ratio_kWh_kgDBO_max)
        ratio_Q_usuel = "'" & ratio_min / 100 & "  -  " & ratio_max / 100
        Feuil6.Range("electricite_ratio_Q_usuel").Cells(1, 3 + pi_FiliereConsideree) = ratio_max / 100
        Feuil6.Range("electricite_ratio_Q_usuel").Cells(1, 6 + pi_FiliereConsideree) = ratio_min / 100
    ElseIf pi_ChoixQualiteRejet = pi_QUALITE_REJET_ZONE_SENSIBLE Then
        'DBO
        ratio_kWh_kgDBO_min = 1
        ratio_kWh_kgDBO_max = 1.7
        ratio_DBOremoved_usuel = "'" & ratio_kWh_kgDBO_min & "  -  " & ratio_kWh_kgDBO_max
        Feuil6.Range("electricite_ratio_DBOremoved_usuel").Cells(1, 3 + pi_FiliereConsideree) = ratio_kWh_kgDBO_max
        Feuil6.Range("electricite_ratio_DBOremoved_usuel").Cells(1, 6 + pi_FiliereConsideree) = ratio_kWh_kgDBO_min
        'EH
        ratio_min = 10 * (pd_REJET_DBO_PAR_HABITANT_ET_PAR_JOUR * pd_NOMBRE_JOUR_PAR_AN * ratio_kWh_kgDBO_min)
        ratio_max = 10 * (pd_REJET_DBO_PAR_HABITANT_ET_PAR_JOUR * pd_NOMBRE_JOUR_PAR_AN * ratio_kWh_kgDBO_max)
        ratio_EH_usuel = "'" & ratio_min / 10 & "  -  " & ratio_max / 10
        Feuil6.Range("electricite_ratio_EH_usuel").Cells(1, 3 + pi_FiliereConsideree) = ratio_max / 10
        Feuil6.Range("electricite_ratio_EH_usuel").Cells(1, 6 + pi_FiliereConsideree) = ratio_min / 10
        'DCO
        ratio_min = 100 * (pd_PourcentageChargeReelleDBO * DBO_nominal / (pd_PourcentageChargeReelleDCO * DCO_nominal) * ratio_kWh_kgDBO_min)
        ratio_max = 100 * (pd_PourcentageChargeReelleDBO * DBO_nominal / (pd_PourcentageChargeReelleDCO * DCO_nominal) * ratio_kWh_kgDBO_max)
        ratio_DCOremoved_usuel = "'" & ratio_min / 100 & "  -  " & ratio_max / 100
        Feuil6.Range("electricite_ratio_DCOremoved_usuel").Cells(1, 3 + pi_FiliereConsideree) = ratio_max / 100
        Feuil6.Range("electricite_ratio_DCOremoved_usuel").Cells(1, 6 + pi_FiliereConsideree) = ratio_min / 100
        'Q
        ratio_min = 100 * (pd_PourcentageChargeReelleDBO * DBO_nominal / (pd_PourcentageChargeReelleDebitVolumique * Q_nominal) * ratio_kWh_kgDBO_min)
        ratio_max = 100 * (pd_PourcentageChargeReelleDBO * DBO_nominal / (pd_PourcentageChargeReelleDebitVolumique * Q_nominal) * ratio_kWh_kgDBO_max)
        ratio_Q_usuel = "'" & ratio_min / 100 & "  -  " & ratio_max / 100
        Feuil6.Range("electricite_ratio_Q_usuel").Cells(1, 3 + pi_FiliereConsideree) = ratio_max / 100
        Feuil6.Range("electricite_ratio_Q_usuel").Cells(1, 6 + pi_FiliereConsideree) = ratio_min / 100
    Else
        'DBO
        ratio_kWh_kgDBO_min = 1.4
        ratio_kWh_kgDBO_max = 2.2
        ratio_DBOremoved_usuel = "'" & ratio_kWh_kgDBO_min & "  -  " & ratio_kWh_kgDBO_max
        Feuil6.Range("electricite_ratio_DBOremoved_usuel").Cells(1, 3 + pi_FiliereConsideree) = ratio_kWh_kgDBO_max
        Feuil6.Range("electricite_ratio_DBOremoved_usuel").Cells(1, 6 + pi_FiliereConsideree) = ratio_kWh_kgDBO_min
        'EH
        ratio_min = 10 * (pd_REJET_DBO_PAR_HABITANT_ET_PAR_JOUR * pd_NOMBRE_JOUR_PAR_AN * ratio_kWh_kgDBO_min)
        ratio_max = 10 * (pd_REJET_DBO_PAR_HABITANT_ET_PAR_JOUR * pd_NOMBRE_JOUR_PAR_AN * ratio_kWh_kgDBO_max)
        ratio_EH_usuel = "'" & ratio_min / 10 & "  -  " & ratio_max / 10
        Feuil6.Range("electricite_ratio_EH_usuel").Cells(1, 3 + pi_FiliereConsideree) = ratio_max / 10
        Feuil6.Range("electricite_ratio_EH_usuel").Cells(1, 6 + pi_FiliereConsideree) = ratio_min / 10
        'DCO
        ratio_min = 100 * (pd_PourcentageChargeReelleDBO * DBO_nominal / (pd_PourcentageChargeReelleDCO * DCO_nominal) * ratio_kWh_kgDBO_min)
        ratio_max = 100 * (pd_PourcentageChargeReelleDBO * DBO_nominal / (pd_PourcentageChargeReelleDCO * DCO_nominal) * ratio_kWh_kgDBO_max)
        ratio_DCOremoved_usuel = "'" & ratio_min / 100 & "  -  " & ratio_max / 100
        Feuil6.Range("electricite_ratio_DCOremoved_usuel").Cells(1, 3 + pi_FiliereConsideree) = ratio_max / 100
        Feuil6.Range("electricite_ratio_DCOremoved_usuel").Cells(1, 6 + pi_FiliereConsideree) = ratio_min / 100
        'Q
        ratio_min = 100 * (pd_PourcentageChargeReelleDBO * DBO_nominal / (pd_PourcentageChargeReelleDebitVolumique * Q_nominal) * ratio_kWh_kgDBO_min)
        ratio_max = 100 * (pd_PourcentageChargeReelleDBO * DBO_nominal / (pd_PourcentageChargeReelleDebitVolumique * Q_nominal) * ratio_kWh_kgDBO_max)
        ratio_Q_usuel = "'" & ratio_min / 100 & "  -  " & ratio_max / 100
        Feuil6.Range("electricite_ratio_Q_usuel").Cells(1, 3 + pi_FiliereConsideree) = ratio_max / 100
        Feuil6.Range("electricite_ratio_Q_usuel").Cells(1, 6 + pi_FiliereConsideree) = ratio_min / 100
    End If

    Feuil6.Range("electricite_ratio_Q_usuel").Cells(1, pi_FiliereConsideree) = ratio_Q_usuel
    Feuil6.Range("electricite_ratio_DCOremoved_usuel").Cells(1, pi_FiliereConsideree) = ratio_DCOremoved_usuel
    Feuil6.Range("electricite_ratio_EH_usuel").Cells(1, pi_FiliereConsideree) = ratio_EH_usuel
    Feuil6.Range("electricite_ratio_DBOremoved_usuel").Cells(1, pi_FiliereConsideree) = ratio_DBOremoved_usuel


    'Report des valeurs pour les graphes
    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
        'comparatif
        Feuil37.Range("graph_electricite_global").Cells(1, pi_FiliereConsideree) = electricite_consommee
        Feuil37.Range("graph_electricite_global").Cells(2, pi_FiliereConsideree) = electricite_verte
        'répartition
        Colonne = 1
        If electricite_consommee > 0 Then
            For I = 1 To electricite_postes_nb
                Feuil37.Range("graph_electricite_step_" & pi_FiliereConsideree).Cells(I, Colonne) = (electricite_fixe(I) + electricite_variable(I)) / electricite_consommee
            Next I
        Else
            For I = 1 To electricite_postes_nb
                Feuil37.Range("graph_electricite_step_" & pi_FiliereConsideree).Cells(I, Colonne) = 0
            Next I
        End If
        'autosuffisance
        Colonne = 1
        Feuil37.Range("graph_autosuffisance_step_" & pi_FiliereConsideree).Cells(1, Colonne) = electricite_autosuffisance
        'prédiction pour calcul des OPEX
        Call ecriture_resultats_OPEX(electricite_fixe, electricite_variable)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "6.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub ecriture_resultats_OPEX(ByVal electricite_fixe, ByVal electricite_variable)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    Dim Colonne As Integer

    Dim interm_somme As Double

    Colonne = 3 * pi_CALCUL_VALEUR_GUIDE + pi_FiliereConsideree   'ATTENTION, ici on définit les valeurs guides des OPEX à partir des valeurs réelles précédentes

    interm_somme = 0
    For I = 1 To electricite_postes_nb
        If electricite_fixe(I) <> 0 Then
            Feuil55.Range("OPEX_electricite_fixe_poste").Cells(I, Colonne) = electricite_fixe(I)
        Else
            Feuil55.Range("OPEX_electricite_fixe_poste").Cells(I, Colonne) = "'-"
        End If
        interm_somme = interm_somme + electricite_fixe(I)
    Next I
    Feuil55.Range("OPEX_electricite_fixe_totale").Cells(1, Colonne) = interm_somme

    interm_somme = 0
    For I = 1 To electricite_postes_nb
        If electricite_variable(I) <> 0 Then
            Feuil55.Range("OPEX_electricite_variable_poste").Cells(I, Colonne) = electricite_variable(I)
        Else
            Feuil55.Range("OPEX_electricite_variable_poste").Cells(I, Colonne) = "'-"
        End If
        interm_somme = interm_somme + electricite_variable(I)
    Next I
    Feuil55.Range("OPEX_electricite_variable_totale").Cells(1, Colonne) = interm_somme



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "6.11"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats_OPEX()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_CALCUL_VALEUR_GUIDE + pi_FiliereConsideree   'ATTENTION, ici on définit les valeurs guides des OPEX à partir des valeurs réelles précédentes

    For I = 1 To electricite_postes_nb
        Feuil55.Range("OPEX_electricite_fixe_poste").Cells(I, Colonne) = ""
    Next I
    For I = 1 To electricite_postes_nb
        Feuil55.Range("OPEX_electricite_variable_poste").Cells(I, Colonne) = ""
    Next I
    Feuil55.Range("OPEX_electricite_fixe_totale").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_electricite_variable_totale").Cells(1, Colonne) = ""



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "6.12"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer
    'Dim i As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree
    Feuil6.Range("electricite_total").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_verte").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_autosuffisance").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_autosuffisance_biogaz").Cells(1, Colonne) = ""

    'modif MSM 25/10/12
    Feuil6.Range("electricite_verte_biogaz").Cells(1, Colonne).ClearContents
    Feuil6.Range("electricite_produite_solaire_photovoltaique").Cells(1, Colonne).ClearContents
    Feuil6.Range("electricite_produite_turbine_hydraulique").Cells(1, Colonne).ClearContents
    Feuil6.Range("electricite_produite_autre").Cells(1, Colonne).ClearContents
    'fin modif MSM 25/10/12

    Feuil6.Range("electricite_file_eau").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_relevement").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_pretraitement").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_primaire").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_secondaire").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_secondaire_aeration_only").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_tertiaire").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_tertiaire_aeration_only").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_desinfection").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_eaux_sales").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_file_boues").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_biolix").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_epaississement").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_digestion").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_athos").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_deshydratation").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_chaulage").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_sechage").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_incineration").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_traitement_retours").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_file_odeur").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_utilites").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_autre_total").Cells(1, Colonne) = ""

    Feuil6.Range("electricite_ratio_Q").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_ratio_DCOremoved").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_ratio_EH").Cells(1, Colonne) = ""
    Feuil6.Range("electricite_ratio_DBOremoved").Cells(1, Colonne) = ""

    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_GUIDE Then
        Feuil6.Range("electricite_ratio_Q_usuel").Cells(1, pi_FiliereConsideree) = ""
        Feuil6.Range("electricite_ratio_DCOremoved_usuel").Cells(1, pi_FiliereConsideree) = ""
        Feuil6.Range("electricite_ratio_EH_usuel").Cells(1, pi_FiliereConsideree) = ""
        Feuil6.Range("electricite_ratio_DBOremoved_usuel").Cells(1, pi_FiliereConsideree) = ""
    End If

    'For i = 1 To 5
    '    Feuil6.Range("electricite_autre").Cells(i, colonne) = ""
    'Next i

    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
        Call effacer_resultats_OPEX
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "6.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub


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
    NumeroErreurNonGeree = "6.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

