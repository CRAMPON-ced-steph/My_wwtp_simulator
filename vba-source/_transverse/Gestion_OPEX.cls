Attribute VB_Name = "Gestion_OPEX"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'repère si le calcul a été effectué
Public calcul_OPEX_OK As Boolean

'nombre de paramètres de coûts
Private Const nb_parametres_couts = 11
Private Const cout_fixe_total = 1
Private Const cout_variable_total = 2
Private Const cout_electricite_fixe = 3
Private Const cout_electricite_variable = 4
Private Const cout_reactifs_fixe = 5
Private Const cout_reactifs_variable = 6
Private Const cout_combustibles_fixe = 7
Private Const cout_combustibles_variable = 8
Private Const cout_transport_boues_variable = 9
Private Const cout_autre_fixe = 10
Private Const cout_autre_variable = 11

Dim nb_postes_conso As Integer    'doit intégrer réactifs, combustibles, électricité, transport
Dim gaz_naturel As Integer
Dim fioul As Integer
Dim electricite As Integer
Dim transport_boues As Integer
Dim OPEX_additionnel As Integer

'variables potentiellement récupérables des calculs précédents
Dim electricite_fixe_kWhj(electricite_postes_nb) As Double
Dim electricite_variable_kWhj(electricite_postes_nb) As Double
Dim reactifs_fixe_conso(reactifs_nb) As Double
Dim reactifs_variable_conso(reactifs_nb) As Double
Dim combustibles_fixe_conso_kWh_PCI(combustibles_nb) As Double
Dim combustibles_variable_conso_kWh_PCI(combustibles_nb) As Double
Dim boues_evacuees_Tj As Double
Dim md_QuantiteRefusDegrillageEvacuee_Tj As Double
Dim graisses_evacuees_Tj As Double
Dim cendres_evacuees_Tj As Double
Dim REFIB_evacuees_Tj As Double
Dim struvite_vendue_kgj As Double
Dim electricite_pdte_vendue_kWhj As Double
Dim electricite_pdte_conso_site_kWhj As Double
Dim biogaz_vendu_kWhj As Double

'variables pour le prix
Dim prix_electricite_achat_kWh As Double
Dim prix_electricite_vente_kWh As Double
Dim prix_biogaz_vendu_kWh As Double
Dim prix_reactifs_T(reactifs_nb) As Double
Dim prix_combustibles_kWh_PCI(combustibles_nb) As Double
Dim prix_evacuation_boues_T As Double
Dim md_PrixEvacuationRefusDegrillage_T As Double
Dim prix_evacuation_graisses_T As Double
Dim prix_evacuation_cendres_T As Double
Dim prix_evacuation_REFIB_T As Double
Dim prix_struvite_T As Double
Dim depense_autre_electricite_fixe_j As Double
Dim depense_autre_electricite_variable_j As Double
Dim benefice_autre_j As Double

'variables potentiellement récupérables des calculs précédents
Dim electricite_fixe_kWhj_force(electricite_postes_nb) As Boolean
Dim electricite_variable_kWhj_force(electricite_postes_nb) As Boolean
Dim reactifs_fixe_conso_force(reactifs_nb) As Boolean
Dim reactifs_variable_conso_force(reactifs_nb) As Boolean
Dim combustibles_fixe_conso_kWh_PCI_force(combustibles_nb) As Boolean
Dim combustibles_variable_conso_kWh_PCI_force(combustibles_nb) As Boolean
Dim boues_evacuees_Tj_force As Boolean
Dim mn_QuantiteRefusDegrillageEvacuee_Tj_force As Boolean
Dim graisses_evacuees_Tj_force As Boolean
Dim cendres_evacuees_Tj_force As Boolean
Dim REFIB_evacuees_Tj_force As Boolean
Dim struvite_vendue_kgj_force As Boolean
Dim electricite_pdte_vendue_kWhj_force As Boolean
Dim electricite_pdte_conso_site_kWhj_force As Boolean
Dim biogaz_vendu_kWhj_force As Boolean

'variables pour savoir si valeur forcée
Dim prix_electricite_achat_kWh_force As Boolean
Dim prix_electricite_vente_kWh_force As Boolean
Dim prix_biogaz_vendu_kWh_force As Boolean
Dim prix_reactifs_T_force(reactifs_nb) As Boolean
Dim prix_combustibles_kWh_PCI_force(combustibles_nb) As Boolean
Dim prix_evacuation_boues_T_force As Boolean
Dim mn_PrixEvacuationRefusDegrillage_T_force As Boolean
Dim prix_evacuation_graisses_T_force As Boolean
Dim prix_evacuation_cendres_T_force As Boolean
Dim prix_evacuation_REFIB_T_force As Boolean
Dim prix_struvite_T_force As Boolean
Dim depense_autre_electricite_fixe_j_force As Boolean
Dim depense_autre_electricite_variable_j_force As Boolean
Dim benefice_autre_j_force As Boolean

'variables de résultats calculés
Dim devise1(nb_parametres_couts) As Double
Dim devise1_reactifs_detail(reactifs_nb) As Double
Dim devise1_combustibles_detail(combustibles_nb) As Double
Dim devise1_benefices_totaux As Double
Dim devise1_benefice_electricite As Double
Dim devise1_benefice_biogaz_vendu As Double
Dim devise1_benefice_struvite As Double

Dim devise1_couts_electricite_file_eau As Double
Dim devise1_couts_electricite_file_boues As Double
Dim devise1_couts_electricite_file_odeurs As Double
Dim devise1_couts_electricite_autres As Double

Dim devise1_couts_reactifs_file_eau As Double
Dim devise1_couts_reactifs_file_boues As Double
Dim devise1_couts_reactifs_file_odeurs As Double
Dim devise1_couts_reactifs_autres As Double

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    Dim Colonne As Integer
    Dim interm_somme As Double

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
        'On écrit les valeurs de consommations réellement prises en compte
        'électricité consommée
        interm_somme = 0
        For I = 1 To electricite_postes_nb
            If electricite_fixe_kWhj(I) <> 0 Then
                Feuil55.Range("OPEX_electricite_fixe_poste").Cells(I, Colonne) = electricite_fixe_kWhj(I)
            Else
                Feuil55.Range("OPEX_electricite_fixe_poste").Cells(I, Colonne) = "'-"
            End If
            interm_somme = interm_somme + electricite_fixe_kWhj(I)
        Next I
        Feuil55.Range("OPEX_electricite_fixe_totale").Cells(1, Colonne) = interm_somme
        interm_somme = 0
        For I = 1 To electricite_postes_nb
            If electricite_variable_kWhj(I) <> 0 Then
                Feuil55.Range("OPEX_electricite_variable_poste").Cells(I, Colonne) = electricite_variable_kWhj(I)
            Else
                Feuil55.Range("OPEX_electricite_variable_poste").Cells(I, Colonne) = "'-"
            End If
            interm_somme = interm_somme + electricite_variable_kWhj(I)
        Next I
        Feuil55.Range("OPEX_electricite_variable_totale").Cells(1, Colonne) = interm_somme
        'réactifs
        For I = 1 To reactifs_nb
            If reactifs_fixe_conso(I) <> 0 Then
                Feuil55.Range("OPEX_reactifs_fixe").Cells(I, Colonne) = reactifs_fixe_conso(I)
            Else
                Feuil55.Range("OPEX_reactifs_fixe").Cells(I, Colonne) = "'-"
            End If
        Next I
        For I = 1 To reactifs_nb
            If reactifs_variable_conso(I) <> 0 Then
                Feuil55.Range("OPEX_reactifs_variable").Cells(I, Colonne) = reactifs_variable_conso(I)
            Else
                Feuil55.Range("OPEX_reactifs_variable").Cells(I, Colonne) = "'-"
            End If
        Next I
        'combustibles
        For I = 1 To combustibles_nb
            Feuil55.Range("OPEX_combustibles_fixe_kWhj").Cells(I, Colonne) = combustibles_fixe_conso_kWh_PCI(I)
        Next I
        For I = 1 To combustibles_nb
            Feuil55.Range("OPEX_combustibles_variable_kWhj").Cells(I, Colonne) = combustibles_variable_conso_kWh_PCI(I)
        Next I
        'boues et électricité pdte
        Feuil55.Range("OPEX_boues_evacuees_Tj").Cells(1, Colonne) = boues_evacuees_Tj
        Feuil55.Range("OPEX_EvacuationRefusDegrillage_Tj").Cells(1, Colonne) = md_QuantiteRefusDegrillageEvacuee_Tj
        Feuil55.Range("OPEX_graisses_Tj").Cells(1, Colonne) = graisses_evacuees_Tj
        Feuil55.Range("OPEX_cendres_evacuees_Tj").Cells(1, Colonne) = cendres_evacuees_Tj
        Feuil55.Range("OPEX_REFIB_evacuees_Tj").Cells(1, Colonne) = REFIB_evacuees_Tj
        Feuil55.Range("OPEX_electricite_pdte_vendue").Cells(1, Colonne) = electricite_pdte_vendue_kWhj
        Feuil55.Range("OPEX_electricite_pdte_conso_site").Cells(1, Colonne) = electricite_pdte_conso_site_kWhj
        Feuil55.Range("OPEX_biogaz_revendu_kWhj").Cells(1, Colonne) = biogaz_vendu_kWhj
        Feuil55.Range("OPEX_struvite_vendue_kgj").Cells(1, Colonne) = struvite_vendue_kgj
    End If

    Feuil55.Range("OPEX_devise1_couts_totaux").Cells(1, Colonne) = devise1(cout_fixe_total) + devise1(cout_variable_total)
    Feuil55.Range("OPEX_devise1_couts_fixes_totaux").Cells(1, Colonne) = devise1(cout_fixe_total)
    Feuil55.Range("OPEX_devise1_couts_variables_totaux").Cells(1, Colonne) = devise1(cout_variable_total)

    Feuil55.Range("OPEX_devise1_couts_electricite_file_eau").Cells(1, Colonne) = devise1_couts_electricite_file_eau
    Feuil55.Range("OPEX_devise1_couts_electricite_file_boues").Cells(1, Colonne) = devise1_couts_electricite_file_boues
    Feuil55.Range("OPEX_devise1_couts_electricite_file_odeurs").Cells(1, Colonne) = devise1_couts_electricite_file_odeurs
    Feuil55.Range("OPEX_devise1_couts_electricite_autres").Cells(1, Colonne) = devise1_couts_electricite_autres
    Feuil55.Range("OPEX_devise1_couts_reactifs_file_eau").Cells(1, Colonne) = devise1_couts_reactifs_file_eau
    Feuil55.Range("OPEX_devise1_couts_reactifs_file_boues").Cells(1, Colonne) = devise1_couts_reactifs_file_boues
    Feuil55.Range("OPEX_devise1_couts_reactifs_file_odeurs").Cells(1, Colonne) = devise1_couts_reactifs_file_odeurs
    Feuil55.Range("OPEX_devise1_couts_reactifs_autres").Cells(1, Colonne) = devise1_couts_reactifs_autres
    Feuil55.Range("OPEX_devise1_couts_totaux_file_eau").Cells(1, Colonne) = devise1_couts_reactifs_file_eau + devise1_couts_electricite_file_eau
    Feuil55.Range("OPEX_devise1_couts_totaux_file_boues").Cells(1, Colonne) = devise1_couts_reactifs_file_boues + devise1_couts_electricite_file_boues
    Feuil55.Range("OPEX_devise1_couts_totaux_file_odeurs").Cells(1, Colonne) = devise1_couts_reactifs_file_odeurs + devise1_couts_electricite_file_odeurs
    Feuil55.Range("OPEX_devise1_couts_totaux_autres").Cells(1, Colonne) = devise1(cout_fixe_total) + devise1(cout_variable_total) - (devise1_couts_reactifs_file_eau + devise1_couts_electricite_file_eau + devise1_couts_reactifs_file_boues + devise1_couts_electricite_file_boues + devise1_couts_reactifs_file_odeurs + devise1_couts_electricite_file_odeurs)

    Feuil55.Range("OPEX_devise1_benefices_totaux").Cells(1, Colonne) = devise1_benefices_totaux

    Feuil55.Range("OPEX_devise1_couts_electricite_totaux").Cells(1, Colonne) = devise1(cout_electricite_fixe) + devise1(cout_electricite_variable)
    Feuil55.Range("OPEX_devise1_couts_electricite_fixe").Cells(1, Colonne) = devise1(cout_electricite_fixe)
    Feuil55.Range("OPEX_devise1_couts_electricite_variable").Cells(1, Colonne) = devise1(cout_electricite_variable)

    Feuil55.Range("OPEX_devise1_couts_reactifs_totaux").Cells(1, Colonne) = devise1(cout_reactifs_fixe) + devise1(cout_reactifs_variable)
    Feuil55.Range("OPEX_devise1_couts_reactifs_fixe").Cells(1, Colonne) = devise1(cout_reactifs_fixe)
    Feuil55.Range("OPEX_devise1_couts_reactifs_variable").Cells(1, Colonne) = devise1(cout_reactifs_variable)

    Feuil55.Range("OPEX_devise1_couts_combustibles_totaux").Cells(1, Colonne) = devise1(cout_combustibles_fixe) + devise1(cout_combustibles_variable)
    Feuil55.Range("OPEX_devise1_couts_combustibles_fixe").Cells(1, Colonne) = devise1(cout_combustibles_fixe)
    Feuil55.Range("OPEX_devise1_couts_combustibles_variable").Cells(1, Colonne) = devise1(cout_combustibles_variable)


    For I = 1 To reactifs_nb
        Feuil55.Range("OPEX_devise1_couts_reactifs_type").Cells(I, Colonne) = devise1_reactifs_detail(I)
    Next I

    For I = 1 To combustibles_nb
        Feuil55.Range("OPEX_devise1_couts_combustibles_type").Cells(I, Colonne) = devise1_combustibles_detail(I)
    Next I

    Feuil55.Range("OPEX_devise1_couts_boues_evacuation").Cells(1, Colonne) = devise1(cout_transport_boues_variable)

    Feuil55.Range("OPEX_devise1_depenses_autres_totales").Cells(1, Colonne) = devise1(cout_autre_fixe) + devise1(cout_autre_variable)
    Feuil55.Range("OPEX_devise1_depenses_autres_fixes").Cells(1, Colonne) = devise1(cout_autre_fixe)
    Feuil55.Range("OPEX_devise1_depenses_autres_variables").Cells(1, Colonne) = devise1(cout_autre_variable)

    Feuil55.Range("OPEX_devise1_benefices_electricite").Cells(1, Colonne) = devise1_benefice_electricite
    Feuil55.Range("OPEX_devise1_benefices_biogaz_revendu").Cells(1, Colonne) = devise1_benefice_biogaz_vendu
    Feuil55.Range("OPEX_devise1_benefices_struvite").Cells(1, Colonne) = devise1_benefice_struvite
    Feuil55.Range("OPEX_devise1_benefices_autres").Cells(1, Colonne) = benefice_autre_j



    '******** GRAPHIQUES ******************
    'comparatif
    Feuil37.Range("graph_OPEX_global").Cells(1, pi_FiliereConsideree) = (devise1(cout_fixe_total) + devise1(cout_variable_total)) * pd_NOMBRE_JOUR_PAR_AN
    Feuil37.Range("graph_OPEX_global").Cells(2, pi_FiliereConsideree) = devise1_benefices_totaux * pd_NOMBRE_JOUR_PAR_AN
    If calcul_OPEX_OK = True Then
        'répartition
        Colonne = 1
        For I = 1 To reactifs_nb
            Feuil37.Range("graph_OPEX_step_" & pi_FiliereConsideree).Cells(I, Colonne) = devise1_reactifs_detail(I) / (devise1(cout_fixe_total) + devise1(cout_variable_total))
        Next I
        For I = 1 To combustibles_nb
            Feuil37.Range("graph_OPEX_step_" & pi_FiliereConsideree).Cells(reactifs_nb + I, Colonne) = devise1_combustibles_detail(I) / (devise1(cout_fixe_total) + devise1(cout_variable_total))
        Next I
        Feuil37.Range("graph_OPEX_step_" & pi_FiliereConsideree).Cells(electricite, Colonne) = (devise1(cout_electricite_fixe) + devise1(cout_electricite_variable)) / (devise1(cout_fixe_total) + devise1(cout_variable_total))
        Feuil37.Range("graph_OPEX_step_" & pi_FiliereConsideree).Cells(transport_boues, Colonne) = devise1(cout_transport_boues_variable) / (devise1(cout_fixe_total) + devise1(cout_variable_total))
        Feuil37.Range("graph_OPEX_step_" & pi_FiliereConsideree).Cells(OPEX_additionnel, Colonne) = (devise1(cout_autre_fixe) + devise1(cout_autre_variable)) / (devise1(cout_fixe_total) + devise1(cout_variable_total))
        'graphique fixe/variable
        Colonne = 1
        Feuil37.Range("graph_OPEX_fixe_variable_" & pi_FiliereConsideree).Cells(1, Colonne) = devise1(cout_fixe_total) / (devise1(cout_fixe_total) + devise1(cout_variable_total))
        Feuil37.Range("graph_OPEX_fixe_variable_" & pi_FiliereConsideree).Cells(2, Colonne) = devise1(cout_variable_total) / (devise1(cout_fixe_total) + devise1(cout_variable_total))
        Feuil37.Range("graph_OPEX_fixe_variable_" & pi_FiliereConsideree).Cells(3, Colonne) = 1
    Else
        'répartition
        Colonne = 1
        For I = 1 To reactifs_nb
            Feuil37.Range("graph_OPEX_step_" & pi_FiliereConsideree).Cells(I, Colonne) = ""
        Next I
        Feuil37.Range("graph_OPEX_step_" & pi_FiliereConsideree).Cells(gaz_naturel, Colonne) = ""
        Feuil37.Range("graph_OPEX_step_" & pi_FiliereConsideree).Cells(fioul, Colonne) = ""
        Feuil37.Range("graph_OPEX_step_" & pi_FiliereConsideree).Cells(electricite, Colonne) = ""
        Feuil37.Range("graph_OPEX_step_" & pi_FiliereConsideree).Cells(transport_boues, Colonne) = ""
        Feuil37.Range("graph_OPEX_step_" & pi_FiliereConsideree).Cells(OPEX_additionnel, Colonne) = ""
        'autosuffisance
        Colonne = 1
        Feuil37.Range("graph_OPEX_fixe_variable_" & pi_FiliereConsideree).Cells(1, Colonne) = ""
        Feuil37.Range("graph_OPEX_fixe_variable_" & pi_FiliereConsideree).Cells(2, Colonne) = ""
        Feuil37.Range("graph_OPEX_fixe_variable_" & pi_FiliereConsideree).Cells(3, Colonne) = ""
    End If

    '************** RETOURS SUR INVESTISSEMENT *****************
    Colonne = 3 * pi_CALCUL_VALEUR_GUIDE + pi_FiliereConsideree
    Feuil55.Range("Investissement_OPEX_devise_an").Cells(1, Colonne) = (devise1(cout_fixe_total) + devise1(cout_variable_total) - devise1_benefices_totaux) * pd_NOMBRE_JOUR_PAR_AN


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "55.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
        'On écrit les valeurs de consommations réellement prises en compte
        'électricité consommée
        For I = 1 To electricite_postes_nb
            Feuil55.Range("OPEX_electricite_fixe_poste").Cells(I, Colonne) = ""
        Next I
        Feuil55.Range("OPEX_electricite_fixe_totale").Cells(1, Colonne) = ""
        For I = 1 To electricite_postes_nb
            Feuil55.Range("OPEX_electricite_variable_poste").Cells(I, Colonne) = ""
        Next I
        Feuil55.Range("OPEX_electricite_variable_totale").Cells(1, Colonne) = ""
        'réactifs
        For I = 1 To reactifs_nb
            Feuil55.Range("OPEX_reactifs_fixe").Cells(I, Colonne) = ""
        Next I
        For I = 1 To reactifs_nb
            Feuil55.Range("OPEX_reactifs_variable").Cells(I, Colonne) = ""
        Next I
        'combustibles
        For I = 1 To combustibles_nb
            Feuil55.Range("OPEX_combustibles_fixe_kWhj").Cells(I, Colonne) = ""
        Next I
        For I = 1 To combustibles_nb
            Feuil55.Range("OPEX_combustibles_variable_kWhj").Cells(I, Colonne) = ""
        Next I
        'boues et électricité pdte
        Feuil55.Range("OPEX_boues_evacuees_Tj").Cells(1, Colonne) = ""
        Feuil55.Range("OPEX_EvacuationRefusDegrillage_Tj").Cells(1, Colonne) = ""
        Feuil55.Range("OPEX_graisses_Tj").Cells(1, Colonne) = ""
        Feuil55.Range("OPEX_cendres_evacuees_Tj").Cells(1, Colonne) = ""
        Feuil55.Range("OPEX_REFIB_evacuees_Tj").Cells(1, Colonne) = ""
        Feuil55.Range("OPEX_electricite_pdte_vendue").Cells(1, Colonne) = ""
        Feuil55.Range("OPEX_electricite_pdte_conso_site").Cells(1, Colonne) = ""
        Feuil55.Range("OPEX_biogaz_revendu_kWhj").Cells(1, Colonne) = ""
        Feuil55.Range("OPEX_struvite_vendue_kgj").Cells(1, Colonne) = ""
    End If


    Feuil55.Range("OPEX_devise1_couts_totaux").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_fixes_totaux").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_variables_totaux").Cells(1, Colonne) = ""

    Feuil55.Range("OPEX_devise1_couts_totaux_file_eau").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_totaux_file_boues").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_totaux_file_odeurs").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_totaux_autres").Cells(1, Colonne) = ""

    Feuil55.Range("OPEX_devise1_couts_electricite_file_eau").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_electricite_file_boues").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_electricite_file_odeurs").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_electricite_autres").Cells(1, Colonne) = ""

    Feuil55.Range("OPEX_devise1_couts_reactifs_file_eau").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_reactifs_file_boues").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_reactifs_file_odeurs").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_reactifs_autres").Cells(1, Colonne) = ""

    Feuil55.Range("OPEX_devise1_benefices_totaux").Cells(1, Colonne) = ""

    Feuil55.Range("OPEX_devise1_couts_electricite_totaux").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_electricite_fixe").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_electricite_variable").Cells(1, Colonne) = ""

    Feuil55.Range("OPEX_devise1_couts_reactifs_totaux").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_reactifs_fixe").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_reactifs_variable").Cells(1, Colonne) = ""

    Feuil55.Range("OPEX_devise1_couts_combustibles_totaux").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_combustibles_fixe").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_couts_combustibles_variable").Cells(1, Colonne) = ""

    For I = 1 To reactifs_nb
        Feuil55.Range("OPEX_devise1_couts_reactifs_type").Cells(I, Colonne) = ""
    Next I

    For I = 1 To combustibles_nb
        Feuil55.Range("OPEX_devise1_couts_combustibles_type").Cells(I, Colonne) = ""
    Next I

    Feuil55.Range("OPEX_devise1_couts_boues_evacuation").Cells(1, Colonne) = ""

    Feuil55.Range("OPEX_devise1_depenses_autres_totales").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_depenses_autres_fixes").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_depenses_autres_variables").Cells(1, Colonne) = ""

    Feuil55.Range("OPEX_devise1_benefices_electricite").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_benefices_biogaz_revendu").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_benefices_struvite").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_devise1_benefices_autres").Cells(1, Colonne) = ""


    '************** RETOURS SUR INVESTISSEMENT *****************
    Colonne = 3 * pi_CALCUL_VALEUR_GUIDE + pi_FiliereConsideree
    Feuil55.Range("Investissement_OPEX_devise_an").Cells(1, Colonne) = ""

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "55.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    Dim Colonne As Integer


    Colonne = pi_FiliereConsideree

    'Toutes ces données sont lues ici car elles sont utilisées dans le cas guide et le cas réel

    prix_electricite_achat_kWh_force = Feuil55.Range("OPEX_electricite_achat_prix_kWh").Cells(1, Colonne)
    prix_electricite_vente_kWh_force = Feuil55.Range("OPEX_electricite_vente_prix_kWh").Cells(1, Colonne)
    prix_biogaz_vendu_kWh_force = Feuil55.Range("OPEX_biogaz_revendu_prix_kWh").Cells(1, Colonne)
    prix_evacuation_boues_T_force = Feuil55.Range("OPEX_boues_evacuees_prix_T").Cells(1, Colonne)
    mn_PrixEvacuationRefusDegrillage_T_force = Feuil55.Range("OPEX_EvacuationRefusDegrillage_prix_T").Cells(1, Colonne)
    prix_evacuation_graisses_T_force = Feuil55.Range("OPEX_graisses_evacuees_prix_T").Cells(1, Colonne)
    prix_evacuation_cendres_T_force = Feuil55.Range("OPEX_cendres_evacuees_prix_T").Cells(1, Colonne)
    prix_evacuation_REFIB_T_force = Feuil55.Range("OPEX_REFIB_evacuees_prix_T").Cells(1, Colonne)
    prix_struvite_T_force = Feuil55.Range("OPEX_struvite_vendue_prix_T").Cells(1, Colonne)
    depense_autre_electricite_fixe_j_force = Feuil55.Range("OPEX_depenses_fixes_autres_j").Cells(1, Colonne)
    depense_autre_electricite_variable_j_force = Feuil55.Range("OPEX_depenses_variables_autres_j").Cells(1, Colonne)
    benefice_autre_j_force = Feuil55.Range("OPEX_benefices_autres_j").Cells(1, Colonne)


    For I = 1 To reactifs_nb
        prix_reactifs_T_force(I) = Feuil55.Range("OPEX_reactifs_prix_T").Cells(I, Colonne)
        If prix_reactifs_T_force(I) = True Then
            prix_reactifs_T(I) = Feuil55.Range("OPEX_reactifs_prix_T").Cells(I, Colonne)
        Else
            prix_reactifs_T(I) = 0
        End If
    Next I

    For I = 1 To combustibles_nb
        prix_combustibles_kWh_PCI_force(I) = Feuil55.Range("OPEX_combustibles_prix_kWh_PCI").Cells(I, Colonne)
        If prix_combustibles_kWh_PCI_force(I) = True Then
            prix_combustibles_kWh_PCI(I) = Feuil55.Range("OPEX_combustibles_prix_kWh_PCI").Cells(I, Colonne)
        Else
            prix_combustibles_kWh_PCI(I) = 0
        End If
    Next I

    If prix_electricite_achat_kWh_force = True Then
        prix_electricite_achat_kWh = Feuil55.Range("OPEX_electricite_achat_prix_kWh").Cells(1, Colonne)
    Else
        prix_electricite_achat_kWh = 0
    End If
    If prix_electricite_vente_kWh_force = True Then
        prix_electricite_vente_kWh = Feuil55.Range("OPEX_electricite_vente_prix_kWh").Cells(1, Colonne)
    Else
        prix_electricite_vente_kWh = 0
    End If
    If prix_biogaz_vendu_kWh_force = True Then
        prix_biogaz_vendu_kWh = Feuil55.Range("OPEX_biogaz_revendu_prix_kWh").Cells(1, Colonne)
    Else
        prix_biogaz_vendu_kWh = 0
    End If
    If prix_evacuation_boues_T_force = True Then
        prix_evacuation_boues_T = Feuil55.Range("OPEX_boues_evacuees_prix_T").Cells(1, Colonne)
    Else
        prix_evacuation_boues_T = 0
    End If
    If mn_PrixEvacuationRefusDegrillage_T_force = True Then
        md_PrixEvacuationRefusDegrillage_T = Feuil55.Range("OPEX_EvacuationRefusDegrillage_prix_T").Cells(1, Colonne)
    Else
        md_PrixEvacuationRefusDegrillage_T = 0
    End If
    If prix_evacuation_graisses_T_force = True Then
        prix_evacuation_graisses_T = Feuil55.Range("OPEX_graisses_evacuees_prix_T").Cells(1, Colonne)
    Else
        prix_evacuation_graisses_T = 0
    End If
    If prix_evacuation_cendres_T_force = True Then
        prix_evacuation_cendres_T = Feuil55.Range("OPEX_cendres_evacuees_prix_T").Cells(1, Colonne)
    Else
        prix_evacuation_cendres_T = 0
    End If
    If prix_evacuation_REFIB_T_force = True Then
        prix_evacuation_REFIB_T = Feuil55.Range("OPEX_REFIB_evacuees_prix_T").Cells(1, Colonne)
    Else
        prix_evacuation_REFIB_T = 0
    End If
    If prix_struvite_T_force = True Then
        prix_struvite_T = Feuil55.Range("OPEX_struvite_vendue_prix_T").Cells(1, Colonne)
    Else
        prix_struvite_T = 0
    End If
    If depense_autre_electricite_fixe_j_force = True Then
        depense_autre_electricite_fixe_j = Feuil55.Range("OPEX_depenses_fixes_autres_j").Cells(1, Colonne)
    Else
        depense_autre_electricite_fixe_j = 0
    End If
    If depense_autre_electricite_variable_j_force = True Then
        depense_autre_electricite_variable_j = Feuil55.Range("OPEX_depenses_variables_autres_j").Cells(1, Colonne)
    Else
        depense_autre_electricite_variable_j = 0
    End If
    If benefice_autre_j_force = True Then
        benefice_autre_j = Feuil55.Range("OPEX_benefices_autres_j").Cells(1, Colonne)
    Else
        benefice_autre_j = 0
    End If




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "55.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim I As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_CALCUL_VALEUR_GUIDE + pi_FiliereConsideree   'on utilise les valeurs calculées auparavant par OCEAN

    If boues_evacuees_Tj_force = True Then
        boues_evacuees_Tj = Feuil55.Range("OPEX_boues_evacuees_Tj").Cells(1, pi_FiliereConsideree)
    Else
        boues_evacuees_Tj = Feuil55.Range("OPEX_boues_evacuees_Tj").Cells(1, Colonne)
    End If
    If mn_QuantiteRefusDegrillageEvacuee_Tj_force = True Then
        md_QuantiteRefusDegrillageEvacuee_Tj = Feuil55.Range("OPEX_EvacuationRefusDegrillage_Tj").Cells(1, pi_FiliereConsideree)
    Else
        md_QuantiteRefusDegrillageEvacuee_Tj = Feuil55.Range("OPEX_EvacuationRefusDegrillage_Tj").Cells(1, Colonne)
    End If
    If graisses_evacuees_Tj_force = True Then
        graisses_evacuees_Tj = Feuil55.Range("OPEX_graisses_Tj").Cells(1, pi_FiliereConsideree)
    Else
        graisses_evacuees_Tj = Feuil55.Range("OPEX_graisses_Tj").Cells(1, Colonne)
    End If
    If cendres_evacuees_Tj_force = True Then
        cendres_evacuees_Tj = Feuil55.Range("OPEX_cendres_evacuees_Tj").Cells(1, pi_FiliereConsideree)
    Else
        cendres_evacuees_Tj = Feuil55.Range("OPEX_cendres_evacuees_Tj").Cells(1, Colonne)
    End If
    If struvite_vendue_kgj_force = True Then
        struvite_vendue_kgj = Feuil55.Range("OPEX_struvite_vendue_kgj").Cells(1, pi_FiliereConsideree)
    Else
        struvite_vendue_kgj = Feuil55.Range("OPEX_struvite_vendue_kgj").Cells(1, Colonne)
    End If
    If REFIB_evacuees_Tj_force = True Then
        REFIB_evacuees_Tj = Feuil55.Range("OPEX_REFIB_evacuees_Tj").Cells(1, pi_FiliereConsideree)
    Else
        REFIB_evacuees_Tj = Feuil55.Range("OPEX_REFIB_evacuees_Tj").Cells(1, Colonne)
    End If
    If electricite_pdte_vendue_kWhj_force = True Then
        electricite_pdte_vendue_kWhj = Feuil55.Range("OPEX_electricite_pdte_vendue").Cells(1, pi_FiliereConsideree)
    Else
        electricite_pdte_vendue_kWhj = Feuil55.Range("OPEX_electricite_pdte_vendue").Cells(1, Colonne)
    End If
    If electricite_pdte_conso_site_kWhj_force = True Then
        electricite_pdte_conso_site_kWhj = Feuil55.Range("OPEX_electricite_pdte_conso_site").Cells(1, pi_FiliereConsideree)
    Else
        electricite_pdte_conso_site_kWhj = Feuil55.Range("OPEX_electricite_pdte_conso_site").Cells(1, Colonne)
    End If
    If biogaz_vendu_kWhj_force = True Then
        biogaz_vendu_kWhj = Feuil55.Range("OPEX_biogaz_revendu_kWhj").Cells(1, pi_FiliereConsideree)
    Else
        biogaz_vendu_kWhj = Feuil55.Range("OPEX_biogaz_revendu_kWhj").Cells(1, Colonne)
    End If

    For I = 1 To electricite_postes_nb
        If electricite_fixe_kWhj_force(I) = True Then
            electricite_fixe_kWhj(I) = Feuil55.Range("OPEX_electricite_fixe_poste").Cells(I, pi_FiliereConsideree)
        Else
            If Feuil55.Range("OPEX_electricite_fixe_poste").Cells(I, Colonne) <> "-" Then
                electricite_fixe_kWhj(I) = Feuil55.Range("OPEX_electricite_fixe_poste").Cells(I, Colonne)
            Else
                electricite_fixe_kWhj(I) = 0
            End If
        End If
        If electricite_variable_kWhj_force(I) = True Then
            electricite_variable_kWhj(I) = Feuil55.Range("OPEX_electricite_variable_poste").Cells(I, pi_FiliereConsideree)
        Else
            If Feuil55.Range("OPEX_electricite_variable_poste").Cells(I, Colonne) <> "-" Then
                electricite_variable_kWhj(I) = Feuil55.Range("OPEX_electricite_variable_poste").Cells(I, Colonne)
            Else
                electricite_variable_kWhj(I) = 0
            End If
        End If
    Next I

    For I = 1 To reactifs_nb
        If reactifs_fixe_conso_force(I) = True Then
            reactifs_fixe_conso(I) = Feuil55.Range("OPEX_reactifs_fixe").Cells(I, pi_FiliereConsideree)
        Else
            If Feuil55.Range("OPEX_reactifs_fixe").Cells(I, Colonne) <> "-" Then
                reactifs_fixe_conso(I) = Feuil55.Range("OPEX_reactifs_fixe").Cells(I, Colonne)
            Else
                reactifs_fixe_conso(I) = 0
            End If
        End If
        If reactifs_variable_conso_force(I) = True Then
            reactifs_variable_conso(I) = Feuil55.Range("OPEX_reactifs_variable").Cells(I, pi_FiliereConsideree)
        Else
            If Feuil55.Range("OPEX_reactifs_variable").Cells(I, Colonne) <> "-" Then
                reactifs_variable_conso(I) = Feuil55.Range("OPEX_reactifs_variable").Cells(I, Colonne)
            Else
                reactifs_variable_conso(I) = 0
            End If
        End If
    Next I

    For I = 1 To combustibles_nb
        If combustibles_fixe_conso_kWh_PCI_force(I) = True Then
            combustibles_fixe_conso_kWh_PCI(I) = Feuil55.Range("OPEX_combustibles_fixe_kWhj").Cells(I, pi_FiliereConsideree)
        Else
            combustibles_fixe_conso_kWh_PCI(I) = Feuil55.Range("OPEX_combustibles_fixe_kWhj").Cells(I, Colonne)
        End If
        If combustibles_variable_conso_kWh_PCI_force(I) = True Then
            combustibles_variable_conso_kWh_PCI(I) = Feuil55.Range("OPEX_combustibles_variable_kWhj").Cells(I, pi_FiliereConsideree)
        Else
            combustibles_variable_conso_kWh_PCI(I) = Feuil55.Range("OPEX_combustibles_variable_kWhj").Cells(I, Colonne)
        End If
    Next I






    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "55.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub

Sub hypothese()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String






    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "55.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub



Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim Colonne As Integer

    Colonne = pi_FiliereConsideree

    boues_evacuees_Tj_force = Feuil55.Range("OPEX_boues_evacuees_Tj").Cells(1, Colonne)
    mn_QuantiteRefusDegrillageEvacuee_Tj_force = Feuil55.Range("OPEX_EvacuationRefusDegrillage_Tj").Cells(1, Colonne)
    graisses_evacuees_Tj_force = Feuil55.Range("OPEX_graisses_Tj").Cells(1, Colonne)
    cendres_evacuees_Tj_force = Feuil55.Range("OPEX_cendres_evacuees_Tj").Cells(1, Colonne)
    REFIB_evacuees_Tj_force = Feuil55.Range("OPEX_REFIB_evacuees_Tj").Cells(1, Colonne)
    electricite_pdte_vendue_kWhj_force = Feuil55.Range("OPEX_electricite_pdte_vendue").Cells(1, Colonne)
    electricite_pdte_conso_site_kWhj_force = Feuil55.Range("OPEX_electricite_pdte_conso_site").Cells(1, Colonne)
    biogaz_vendu_kWhj_force = Feuil55.Range("OPEX_biogaz_revendu_kWhj").Cells(1, Colonne)
    struvite_vendue_kgj_force = Feuil55.Range("OPEX_struvite_vendue_kgj").Cells(1, Colonne)


    For I = 1 To electricite_postes_nb
        electricite_fixe_kWhj_force(I) = Feuil55.Range("OPEX_electricite_fixe_poste").Cells(I, Colonne)
        electricite_variable_kWhj_force(I) = Feuil55.Range("OPEX_electricite_variable_poste").Cells(I, Colonne)
    Next I

    For I = 1 To reactifs_nb
        reactifs_fixe_conso_force(I) = Feuil55.Range("OPEX_reactifs_fixe").Cells(I, Colonne)
        reactifs_variable_conso_force(I) = Feuil55.Range("OPEX_reactifs_variable").Cells(I, Colonne)
    Next I

    For I = 1 To combustibles_nb
        combustibles_fixe_conso_kWh_PCI_force(I) = Feuil55.Range("OPEX_combustibles_fixe_kWhj").Cells(I, Colonne)
        combustibles_variable_conso_kWh_PCI_force(I) = Feuil55.Range("OPEX_combustibles_variable_kWhj").Cells(I, Colonne)
    Next I



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "55.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Sub Class_Initialize()


    'on définit les constantes à utiliser
    nb_postes_conso = reactifs_nb + combustibles_nb + OPEX_autres
    gaz_naturel = reactifs_nb + combustible_repere_gaz_naturel
    fioul = reactifs_nb + combustible_repere_fioul
    electricite = reactifs_nb + combustibles_nb + OPEX_autres_electricite
    transport_boues = reactifs_nb + combustibles_nb + OPEX_autres_transport_boues
    OPEX_additionnel = reactifs_nb + combustibles_nb + OPEX_autres_OPEX_additionnel


    If pn_EffacementResultatsEnCours = False Then
        Call lecture_choix
        Call hypothese
        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
            Call lecture_valeurs_forcees
        End If
        Call attribution_valeur_par_defaut
    End If

End Sub

Sub calcul()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    Dim somme As Double

    'on vérifie que les prix ont été renseignés pour mettre des commentaires
    For I = 1 To reactifs_nb
        If prix_reactifs_T_force(I) = False Then
            If (reactifs_fixe_conso(I) + reactifs_variable_conso(I)) > 0.001 Then
                ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil55.Range("OPEX_reactifs_erreur_prix").Cells(I, 1) + vbCrLf
            End If
        End If
    Next I
    For I = 1 To combustibles_nb
        If prix_combustibles_kWh_PCI_force(I) = False Then
            If (combustibles_fixe_conso_kWh_PCI(I) + combustibles_variable_conso_kWh_PCI(I)) > 0.001 Then
                ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil55.Range("OPEX_combustibles_erreur_prix").Cells(I, 1) + vbCrLf
            End If
        End If
    Next I
    If prix_electricite_achat_kWh_force = False Then
        somme = 0
        For I = 1 To electricite_postes_nb
            somme = somme + electricite_fixe_kWhj(I) + electricite_variable_kWhj(I)
        Next I
        If somme > 0.001 Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil55.Range("OPEX_electricite_erreur_prix_achat").Cells(1, 1) + vbCrLf
        End If
    End If
    If prix_electricite_vente_kWh_force = False Then
        If electricite_pdte_vendue_kWhj > 0.001 Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil55.Range("OPEX_electricite_erreur_prix_revente").Cells(1, 1) + vbCrLf
        End If
    End If
    If prix_biogaz_vendu_kWh_force = False Then
        If biogaz_vendu_kWhj > 0.001 Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil55.Range("OPEX_biogaz_revendu_erreur_prix").Cells(1, 1) + vbCrLf
        End If
    End If
    If prix_evacuation_boues_T_force = False Then
        If boues_evacuees_Tj > 0.001 Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil55.Range("OPEX_boues_erreur_prix_evacuation").Cells(1, 1) + vbCrLf
        End If
    End If
    If mn_PrixEvacuationRefusDegrillage_T_force = False Then
        If md_QuantiteRefusDegrillageEvacuee_Tj > 0.001 Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil55.Range("OPEX_RefusDegrillage_erreur_prix_evacuation").Cells(1, 1) + vbCrLf
        End If
    End If
    If prix_evacuation_graisses_T_force = False Then
        If graisses_evacuees_Tj > 0.001 Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil55.Range("OPEX_graisses_erreur_prix_evacuation").Cells(1, 1) + vbCrLf
        End If
    End If
    If prix_evacuation_cendres_T_force = False Then
        If cendres_evacuees_Tj > 0.001 Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil55.Range("OPEX_cendres_erreur_prix_evacuation").Cells(1, 1) + vbCrLf
        End If
    End If
    If prix_evacuation_REFIB_T_force = False Then
        If REFIB_evacuees_Tj > 0.001 Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil55.Range("OPEX_REFIB_erreur_prix_evacuation").Cells(1, 1) + vbCrLf
        End If
    End If
    If prix_struvite_T_force = False Then
        If struvite_vendue_kgj > 0.001 Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil55.Range("OPEX_struvite_erreur_prix_revente").Cells(1, 1) + vbCrLf
        End If
    End If


    'On calcule l'électricité
    devise1(cout_electricite_fixe) = 0
    devise1(cout_electricite_variable) = 0
    devise1_couts_electricite_file_eau = 0
    devise1_couts_electricite_file_boues = 0
    devise1_couts_electricite_file_odeurs = 0
    devise1_couts_electricite_autres = 0
    For I = 1 To electricite_postes_nb
        devise1(cout_electricite_fixe) = devise1(cout_electricite_fixe) + electricite_fixe_kWhj(I) * prix_electricite_achat_kWh
        devise1(cout_electricite_variable) = devise1(cout_electricite_variable) + electricite_variable_kWhj(I) * prix_electricite_achat_kWh
        If I <= electricite_postes_eau_nb Then
            devise1_couts_electricite_file_eau = devise1_couts_electricite_file_eau + (electricite_fixe_kWhj(I) + electricite_variable_kWhj(I)) * prix_electricite_achat_kWh
        ElseIf I <= electricite_postes_boues_nb Then
            devise1_couts_electricite_file_boues = devise1_couts_electricite_file_boues + (electricite_fixe_kWhj(I) + electricite_variable_kWhj(I)) * prix_electricite_achat_kWh
        ElseIf I = electricite_postes_repere_desodorisation Then
            devise1_couts_electricite_file_odeurs = devise1_couts_electricite_file_odeurs + (electricite_fixe_kWhj(I) + electricite_variable_kWhj(I)) * prix_electricite_achat_kWh
        Else
            devise1_couts_electricite_autres = devise1_couts_electricite_autres + (electricite_fixe_kWhj(I) + electricite_variable_kWhj(I)) * prix_electricite_achat_kWh
        End If
    Next I
    devise1(cout_electricite_variable) = devise1(cout_electricite_variable) - electricite_pdte_conso_site_kWhj * prix_electricite_achat_kWh

    'On calcule les réactifs
    devise1(cout_reactifs_fixe) = 0
    devise1(cout_reactifs_variable) = 0
    For I = 1 To reactifs_nb
        devise1(cout_reactifs_fixe) = devise1(cout_reactifs_fixe) + reactifs_fixe_conso(I) * prix_reactifs_T(I) / 1000
        devise1(cout_reactifs_variable) = devise1(cout_reactifs_variable) + reactifs_variable_conso(I) * prix_reactifs_T(I) / 1000
        devise1_reactifs_detail(I) = devise1_reactifs_detail(I) + (reactifs_fixe_conso(I) + reactifs_variable_conso(I)) * prix_reactifs_T(I) / 1000
    Next I
    devise1_couts_reactifs_file_eau = devise1_reactifs_detail(reactifs_repere_polymere_eau_poudre_anionique)
    devise1_couts_reactifs_file_eau = devise1_couts_reactifs_file_eau + devise1_reactifs_detail(reactifs_repere_polymere_eau_poudre_cationique)
    devise1_couts_reactifs_file_eau = devise1_couts_reactifs_file_eau + devise1_reactifs_detail(reactifs_repere_polymere_eau_liquide_anionique)
    devise1_couts_reactifs_file_eau = devise1_couts_reactifs_file_eau + devise1_reactifs_detail(reactifs_repere_polymere_eau_liquide_cationique)
    devise1_couts_reactifs_file_eau = devise1_couts_reactifs_file_eau + devise1_reactifs_detail(reactifs_repere_FeCl3_eau)
    devise1_couts_reactifs_file_eau = devise1_couts_reactifs_file_eau + devise1_reactifs_detail(reactifs_repere_methanol)
    devise1_couts_reactifs_file_eau = devise1_couts_reactifs_file_eau + devise1_reactifs_detail(reactifs_repere_NaOCl_eau)
    devise1_couts_reactifs_file_eau = devise1_couts_reactifs_file_eau + devise1_reactifs_detail(reactifs_repere_Ca_2NO3)
    devise1_couts_reactifs_file_eau = devise1_couts_reactifs_file_eau + devise1_reactifs_detail(reactifs_repere_H2O2)
    devise1_couts_reactifs_file_eau = devise1_couts_reactifs_file_eau + devise1_reactifs_detail(reactifs_repere_Cl2_gazeux)
    devise1_couts_reactifs_file_boues = devise1_reactifs_detail(reactifs_repere_polymere_boues_poudre_anionique)
    devise1_couts_reactifs_file_boues = devise1_couts_reactifs_file_boues + devise1_reactifs_detail(reactifs_repere_polymere_boues_poudre_cationique)
    devise1_couts_reactifs_file_boues = devise1_couts_reactifs_file_boues + devise1_reactifs_detail(reactifs_repere_polymere_boues_liquide_anionique)
    devise1_couts_reactifs_file_boues = devise1_couts_reactifs_file_boues + devise1_reactifs_detail(reactifs_repere_polymere_boues_liquide_cationique)
    devise1_couts_reactifs_file_boues = devise1_couts_reactifs_file_boues + devise1_reactifs_detail(reactifs_repere_FeCl3_boues)
    devise1_couts_reactifs_file_boues = devise1_couts_reactifs_file_boues + devise1_reactifs_detail(reactifs_repere_chaux_eteinte)
    devise1_couts_reactifs_file_boues = devise1_couts_reactifs_file_boues + devise1_reactifs_detail(reactifs_repere_chaux_vive)
    devise1_couts_reactifs_file_boues = devise1_couts_reactifs_file_boues + devise1_reactifs_detail(reactifs_repere_NaHCO3)
    devise1_couts_reactifs_file_boues = devise1_couts_reactifs_file_boues + devise1_reactifs_detail(reactifs_repere_O2_liquide)
    devise1_couts_reactifs_file_boues = devise1_couts_reactifs_file_boues + devise1_reactifs_detail(reactifs_repere_charbon_actif)
    devise1_couts_reactifs_file_boues = devise1_couts_reactifs_file_boues + devise1_reactifs_detail(reactifs_repere_CuSO4)
    devise1_couts_reactifs_file_boues = devise1_couts_reactifs_file_boues + devise1_reactifs_detail(reactifs_repere_MgCl2)
    devise1_couts_reactifs_file_odeurs = devise1_reactifs_detail(reactifs_repere_H2SO4)
    devise1_couts_reactifs_file_odeurs = devise1_couts_reactifs_file_odeurs + devise1_reactifs_detail(reactifs_repere_NaOH)
    devise1_couts_reactifs_file_odeurs = devise1_couts_reactifs_file_odeurs + devise1_reactifs_detail(reactifs_repere_NaOCl_desodo)
    devise1_couts_reactifs_file_odeurs = devise1_couts_reactifs_file_odeurs + devise1_reactifs_detail(reactifs_repere_NaHSO3)
    devise1_couts_reactifs_autres = 0


    'On calcule les combustibles
    devise1(cout_combustibles_fixe) = 0
    devise1(cout_combustibles_variable) = 0
    For I = 1 To combustibles_nb
        devise1(cout_combustibles_fixe) = devise1(cout_combustibles_fixe) + combustibles_fixe_conso_kWh_PCI(I) * prix_combustibles_kWh_PCI(I)
        devise1(cout_combustibles_variable) = devise1(cout_combustibles_variable) + combustibles_variable_conso_kWh_PCI(I) * prix_combustibles_kWh_PCI(I)
        devise1_combustibles_detail(I) = devise1_combustibles_detail(I) + (combustibles_fixe_conso_kWh_PCI(I) + combustibles_variable_conso_kWh_PCI(I)) * prix_combustibles_kWh_PCI(I)
    Next I

    'transport boues
    devise1(cout_transport_boues_variable) = boues_evacuees_Tj * prix_evacuation_boues_T + graisses_evacuees_Tj * prix_evacuation_graisses_T + cendres_evacuees_Tj * prix_evacuation_cendres_T + REFIB_evacuees_Tj * prix_evacuation_REFIB_T + md_QuantiteRefusDegrillageEvacuee_Tj * md_PrixEvacuationRefusDegrillage_T

    'autres dépenses
    devise1(cout_autre_fixe) = depense_autre_electricite_fixe_j
    devise1(cout_autre_variable) = depense_autre_electricite_variable_j

    'total des dépenses fixes et variables
    devise1(cout_fixe_total) = devise1(cout_electricite_fixe) + devise1(cout_reactifs_fixe) + devise1(cout_combustibles_fixe) + devise1(cout_autre_fixe)
    devise1(cout_variable_total) = devise1(cout_electricite_variable) + devise1(cout_reactifs_variable) + devise1(cout_combustibles_variable) + devise1(cout_transport_boues_variable) + devise1(cout_autre_variable)

    'Bénéfices
    devise1_benefice_electricite = electricite_pdte_vendue_kWhj * prix_electricite_vente_kWh
    devise1_benefice_biogaz_vendu = biogaz_vendu_kWhj * prix_biogaz_vendu_kWh
    devise1_benefice_struvite = prix_struvite_T * struvite_vendue_kgj / 1000
    devise1_benefices_totaux = benefice_autre_j + devise1_benefice_electricite + devise1_benefice_biogaz_vendu + devise1_benefice_struvite

    If (devise1(cout_fixe_total) + devise1(cout_variable_total)) > 0 Then
        calcul_OPEX_OK = True
    Else
        calcul_OPEX_OK = False
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "55.12"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Private Sub Class_Terminate()


    If pn_EffacementResultatsEnCours = False Then
        Call ecriture_resultats
    End If


End Sub









