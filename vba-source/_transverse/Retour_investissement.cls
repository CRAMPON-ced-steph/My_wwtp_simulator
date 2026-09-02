Attribute VB_Name = "Retour_investissement"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'repérage pour l'écriture détaillée du temps de retour sur investissement
Private Const nb_repere_temps = 3
Private Const repere_annees = 1
Private Const repere_mois = 2
Private Const repere_jours = 3

'calcul du temps de retour effectué
Dim retour_invest_effectue(3) As Boolean

'choix de la pi_FiliereConsideree de référence
Dim choix_step_reference As Integer

'données d'OPEX
Dim OPEX_reference_an As Double
Dim OPEX_tous_an(3) As Double

Dim cout_investissement(3) As Double
Dim duree_contrat_an(3) As Double
Dim taux_actualisation(3) As Double
Dim annees_max As Integer

'résultats
Dim CAPEX_OPEX_x_annees(3) As Double
Dim VAN_devise(3) As Double
'Dim tps_retour_investissement(3, 3) As Long
Dim tps_retour_investissement_an(3) As Double
Dim tps_retour_investissement_actualise_an(3) As Double


Dim graph_VAN_series() As Double
Sub calcul()

    'on calcule le temps de retour sur investissement
    Call calcul_tps_retour_investissement

    'on calcule le VAN et le coût CAPEX+xOPEX
    Call calcul_CAPEX_xOPEX
    Call calcul_VAN

    Call gestion_graphique

End Sub

Sub calcul_CAPEX_xOPEX()

    Dim I As Integer

    For I = 1 To nb_step_comparaison
        CAPEX_OPEX_x_annees(I) = cout_investissement(I) + duree_contrat_an(I) * OPEX_tous_an(I)
    Next I


End Sub


Sub calcul_tps_retour_investissement()

    Dim I As Integer

    'dim tps_retour_annees as long
    'Dim tps_retour_mois As Long
    'Dim tps_retour_jours As Long

    For I = 1 To nb_step_comparaison
        If I <> choix_step_reference Then
            'ca ne marche pas si OPEX_ref<OPEX_post_investissement
            If OPEX_reference_an <= OPEX_tous_an(I) Then
                'message pour dire que le calcul n'a pas de sens
                ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil55.Range("Investissement_filiere").Cells(1, 1) + " " + Str(I) + ": " + Feuil55.Range("Investissement_erreur_rentabilite_jamais").Cells(1, 1) + vbCrLf
                retour_invest_effectue(I) = False
            Else
                tps_retour_investissement_an(I) = (cout_investissement(I) - cout_investissement(choix_step_reference)) / (OPEX_reference_an - OPEX_tous_an(I))
                'on calcule le temps en années
                'tps_retour_an = partie_entiere(cout_investissement(i) / (OPEX_reference_an - OPEX_tous_an(i)))
                'on adapte le calcul en années/mois/jours
                'tps_retour_annees = partie_entiere(tps_retour_an)
                '    tps_retour_jours = (tps_retour_an - tps_retour_annees) * pd_NOMBRE_JOUR_PAR_AN
                'tps_retour_mois = partie_entiere(tps_retour_jours / (pd_NOMBRE_JOUR_PAR_AN / pd_NOMBRE_MOIS_PAR_AN))
                '    tps_retour_jours = tps_retour_jours - tps_retour_mois * (pd_NOMBRE_JOUR_PAR_AN / pd_NOMBRE_MOIS_PAR_AN)
                'tps_retour_investissement(i, repere_annees) = tps_retour_annees
                'tps_retour_investissement(i, repere_mois) = tps_retour_mois
                'tps_retour_investissement(i, repere_jours) = tps_retour_jours
                retour_invest_effectue(I) = True
            End If
        End If
    Next I


End Sub


Sub calcul_VAN()

    Dim I As Integer
    Dim j As Integer

    annees_max = 0
    For I = 1 To nb_step_comparaison
        If annees_max < duree_contrat_an(I) Then
            annees_max = duree_contrat_an(I)
        End If
    Next I
    ReDim graph_VAN_series(nb_step_comparaison, annees_max)

    For I = 1 To nb_step_comparaison
        If I <> choix_step_reference Then
            graph_VAN_series(I, 0) = cout_investissement(choix_step_reference) - cout_investissement(I)
            For j = 1 To duree_contrat_an(I)
                graph_VAN_series(I, j) = graph_VAN_series(I, j - 1) + (OPEX_tous_an(choix_step_reference) - OPEX_tous_an(I)) / ((1 + taux_actualisation(I)) ^ j)
                If graph_VAN_series(I, j) >= 0 And graph_VAN_series(I, j - 1) < 0 Then
                    tps_retour_investissement_actualise_an(I) = j - 1 + Abs(graph_VAN_series(I, j - 1)) / (graph_VAN_series(I, j) + Abs(graph_VAN_series(I, j - 1)))
                End If
            Next j
            VAN_devise(I) = graph_VAN_series(I, duree_contrat_an(I))
        End If
    Next I


End Sub

Sub ecrire_resultats()

    Dim I As Integer
    Dim j As Integer
    Dim Colonne As Integer

    For I = 1 To nb_step_comparaison  'filière
        Colonne = 3 * pi_EtapeCalculConsideree + I
        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
            Feuil55.Range("Investissement_OPEX_devise_an").Cells(1, Colonne) = OPEX_tous_an(I)
        End If
        Feuil55.Range("Investissement_CAPEX_OPEX_x_annees").Cells(1, Colonne) = CAPEX_OPEX_x_annees(I)
        Feuil55.Range("Investissement_VAN_devise").Cells(1, Colonne) = VAN_devise(I)
        If I <> choix_step_reference Then
            If retour_invest_effectue(I) = True Then
                Feuil55.Range("Investissement_tps_retour_annees").Cells(1, Colonne) = tps_retour_investissement_an(I)
                If VAN_devise(I) >= 0 Then
                    Feuil55.Range("Investissement_tps_retour_annees_actualise").Cells(1, Colonne) = tps_retour_investissement_actualise_an(I)
                Else
                    Feuil55.Range("Investissement_tps_retour_annees_actualise").Cells(1, Colonne) = Feuil55.Range("Investissement_TRI_trop_long").Cells(1, 1)
                End If
            Else
                Feuil55.Range("Investissement_tps_retour_annees").Cells(1, Colonne) = "'-"
                Feuil55.Range("Investissement_tps_retour_annees_actualise").Cells(1, Colonne) = "'-"
            End If
        Else
            Feuil55.Range("Investissement_tps_retour_annees").Cells(1, Colonne) = "'-"
        End If
        'For j = 1 To nb_repere_temps  'années/mois/jours
        '    If retour_invest_effectue(i) = True Then
        '        Feuil1.Range("RetourInvest_tps_retour_investissement").Cells(j, i) = tps_retour_investissement(i, j)
        '    Else
        '        Feuil1.Range("RetourInvest_tps_retour_investissement").Cells(j, i) = "'-"
        '    End If
        'Next j
    Next I


End Sub

Sub effacer_resultats()

    Dim I As Integer
    Dim j As Integer

    Dim Colonne As Integer

    For I = 1 To 3  'filière
        Colonne = 3 * pi_EtapeCalculConsideree + I
        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
            Feuil55.Range("Investissement_OPEX_devise_an").Cells(1, Colonne) = ""
        End If
        Feuil55.Range("Investissement_CAPEX_OPEX_x_annees").Cells(1, Colonne) = ""
        Feuil55.Range("Investissement_VAN_devise").Cells(1, Colonne) = ""
        Feuil55.Range("Investissement_tps_retour_annees").Cells(1, Colonne) = ""
        Feuil55.Range("Investissement_tps_retour_annees_actualise").Cells(1, Colonne) = ""
        'For j = 1 To nb_repere_temps  'années/mois/jours
        '    If retour_invest_effectue(i) = True Then
        '        Feuil1.Range("RetourInvest_tps_retour_investissement").Cells(j, i) = tps_retour_investissement(i, j)
        '    Else
        '        Feuil1.Range("RetourInvest_tps_retour_investissement").Cells(j, i) = "'-"
        '    End If
        'Next j
    Next I

    Feuil59.ChartObjects("Retour_investissement_graph_VAN").Visible = False

End Sub

Sub gestion_graphique()

    Dim I As Integer
    Dim j As Integer
    Dim matrice_serie() As Double

    Dim feuille_active As Worksheet
    Set feuille_active = ActiveSheet
    Dim Plage As Range
    Dim zone_graph As ChartObject
    Dim graph As Chart

    Dim nb_positif As Integer
    Dim nb_negatif As Integer
    Dim somme As Integer
    Dim index_couleur(1 To 3) As Integer
    'On regarde pour les couleurs
    nb_positif = 0
    nb_negatif = 0
    For I = 1 To nb_step_comparaison
        somme = 0
        For j = 1 To nb_step_comparaison
            If j <> I Then
                If VAN_devise(I) > VAN_devise(j) Then
                    somme = somme + 1
                End If
            End If
        Next j
        index_couleur(I) = somme
        If somme > nb_positif Then
            nb_positif = somme
        End If
    Next I
    somme = index_couleur(choix_step_reference)
    nb_positif = nb_positif - somme
    nb_negatif = nb_negatif - somme
    For I = 1 To nb_step_comparaison
        index_couleur(I) = index_couleur(I) - somme
    Next I

    Set zone_graph = Feuil59.ChartObjects("Retour_investissement_graph_VAN")
    Set graph = zone_graph.Chart
    pause_systeme

    With graph
        pause_systeme
        .ChartType = xlXYScatterLines
        pause_systeme
        .HasTitle = True
        pause_systeme
        .ChartTitle.Caption = Feuil55.Range("Investissement_titre_graph").Cells(1, 1)
        pause_systeme
        .Axes(xlCategory).AxisTitle.Caption = Feuil55.Range("Investissement_titre_legende_abscisse").Cells(1, 1)
        pause_systeme
        .Axes(xlCategory).MinimumScale = 0
        pause_systeme
        .Axes(xlCategory).MaximumScale = annees_max
        pause_systeme
        .Axes(xlValue).AxisTitle.Caption = Feuil55.Range("Investissement_titre_legende_ordonnee").Cells(1, 1)
        For j = 1 To graph.SeriesCollection.Count
            pause_systeme
            .SeriesCollection(1).Delete
        Next j
        For I = 1 To nb_step_comparaison
            ReDim matrice_serie(0 To duree_contrat_an(I), 1 To 2)
            Application.EnableEvents = False
            For j = 0 To duree_contrat_an(I)
                matrice_serie(j, 1) = j
                Feuil59.Range("graph_RI_step_" & I).Cells(j + 2, 1) = matrice_serie(j, 1)
                DoneExInvalidateCommand Range("graph_RI_step_" & I).Cells(j + 2, 1)
                matrice_serie(j, 2) = graph_VAN_series(I, j)
                Feuil59.Range("graph_RI_step_" & I).Cells(j + 2, 2) = matrice_serie(j, 2)
                DoneExInvalidateCommand Range("graph_RI_step_" & I).Cells(j + 2, 2)
            Next j
            Application.EnableEvents = True
            Set Plage = Feuil59.Range(Feuil59.Range("graph_RI_step_" & I).Cells(2, 2), Feuil59.Range("graph_RI_step_" & I).Cells(duree_contrat_an(I) + 2, 2))
            .SeriesCollection.Add source:=Plage
            Set Plage = Feuil59.Range(Feuil59.Range("graph_RI_step_" & I).Cells(2, 1), Feuil59.Range("graph_RI_step_" & I).Cells(duree_contrat_an(I) + 2, 1))
            .SeriesCollection(I).XValues = Plage
            .SeriesCollection(I).Name = Feuil2.Range("synthese_nom_step" & I).Cells(1, 1)
            '.SeriesCollection(i).Legend = "Filière " & i
            If I = choix_step_reference Then
                '.PlotArea.Interior.Color = RGB(255, 0, 0)
                If Application.Version = version_Excel_2003 Then
                    Feuil59.ChartObjects("Retour_investissement_graph_VAN").Activate
                    ActiveChart.SeriesCollection(I).Select
                Else
                    pause_systeme
                    graph.SeriesCollection(I).Select
                End If
                'marqueurs
                Selection.MarkerBackgroundColor = RGB(50, 50, 50)
                Selection.MarkerForegroundColor = RGB(50, 50, 50)
                Selection.MarkerStyle = xlDiamond
                Selection.Smooth = False
                Selection.MarkerSize = 8
                Selection.Shadow = False
                'sélection des traits
                Selection.Border.color = RGB(50, 50, 50)
                Selection.Border.Weight = xlThick
                Selection.Border.LineStyle = xlContinuous
            Else
                If I = 3 Then
                    '.PlotArea.Interior.Color = RGB(255, 0, 0)
                    If Application.Version = version_Excel_2003 Then
                        Feuil59.ChartObjects("Retour_investissement_graph_VAN").Activate
                        ActiveChart.SeriesCollection(I).Select
                    Else
                        pause_systeme
                        graph.SeriesCollection(I).Select
                    End If
                    'marqueurs
                    Selection.MarkerBackgroundColor = RGB(50, 189, 50)
                    Selection.MarkerForegroundColor = RGB(50, 189, 50)
                    Selection.MarkerStyle = xlDiamond
                    Selection.Smooth = False
                    Selection.MarkerSize = 8
                    Selection.Shadow = False
                    'sélection des traits
                    Selection.Border.color = RGB(50, 189, 50)
                    Selection.Border.Weight = xlThick
                    Selection.Border.LineStyle = xlContinuous
                    'On attribue les couleurs
                    If nb_negatif = -1 And index_couleur(I) = nb_negatif Then
                        index_couleur(I) = -2
                    ElseIf nb_positif = 1 And index_couleur(I) = nb_positif Then
                        index_couleur(I) = 2
                    End If
                    Select Case index_couleur(I)
                    Case -2
                        Selection.MarkerBackgroundColor = RGB(255, 0, 0)
                        Selection.MarkerForegroundColor = RGB(255, 0, 0)
                        Selection.Border.color = RGB(255, 0, 0)
                    Case -1
                        Selection.MarkerBackgroundColor = RGB(255, 153, 0)
                        Selection.MarkerForegroundColor = RGB(255, 153, 0)
                        Selection.Border.color = RGB(255, 153, 0)
                    Case 1
                        Selection.MarkerBackgroundColor = RGB(153, 255, 51)
                        Selection.MarkerForegroundColor = RGB(153, 255, 51)
                        Selection.Border.color = RGB(153, 255, 51)
                    Case 2
                        Selection.MarkerBackgroundColor = RGB(0, 204, 102)
                        Selection.MarkerForegroundColor = RGB(0, 204, 102)
                        Selection.Border.color = RGB(0, 204, 102)
                    End Select
                Else
                    '.PlotArea.Interior.Color = RGB(255, 0, 0)
                    If Application.Version = version_Excel_2003 Then
                        Feuil59.ChartObjects("Retour_investissement_graph_VAN").Activate
                        ActiveChart.SeriesCollection(I).Select
                    Else
                        graph.SeriesCollection(I).Select
                    End If
                    'marqueurs
                    Selection.MarkerBackgroundColor = RGB(0, 0, 255)
                    Selection.MarkerForegroundColor = RGB(0, 0, 255)
                    Selection.MarkerStyle = xlDiamond
                    Selection.Smooth = False
                    Selection.MarkerSize = 8
                    Selection.Shadow = False
                    'sélection des traits
                    Selection.Border.color = RGB(0, 0, 255)
                    Selection.Border.Weight = xlThick
                    Selection.Border.LineStyle = xlContinuous
                    'On attribue les couleurs
                    If nb_negatif = -1 And index_couleur(I) = nb_negatif Then
                        index_couleur(I) = -2
                    ElseIf nb_positif = 1 And index_couleur(I) = nb_positif Then
                        index_couleur(I) = 2
                    End If
                    Select Case index_couleur(I)
                    Case -2
                        Selection.MarkerBackgroundColor = RGB(255, 0, 0)
                        Selection.MarkerForegroundColor = RGB(255, 0, 0)
                        Selection.Border.color = RGB(255, 0, 0)
                    Case -1
                        Selection.MarkerBackgroundColor = RGB(255, 153, 0)
                        Selection.MarkerForegroundColor = RGB(255, 153, 0)
                        Selection.Border.color = RGB(255, 153, 0)
                    Case 1
                        Selection.MarkerBackgroundColor = RGB(153, 255, 51)
                        Selection.MarkerForegroundColor = RGB(153, 255, 51)
                        Selection.Border.color = RGB(153, 255, 51)
                    Case 2
                        Selection.MarkerBackgroundColor = RGB(0, 204, 102)
                        Selection.MarkerForegroundColor = RGB(0, 204, 102)
                        Selection.Border.color = RGB(0, 204, 102)
                    End Select
                End If
            End If
        Next I
    End With

    feuille_active.Activate
    pause_systeme
    zone_graph.Visible = True

End Sub

Sub lecture_choix()

    Dim I As Integer
    Dim Colonne As Integer

    Dim invest_min As Double
    Dim invest_ref As Double

    'lecture du choix de référence
    choix_step_reference = Feuil55.Range("Investissement_choix_filiere_ref").Cells(1, 1)
    If choix_step_reference > nb_step_comparaison Then
        pn_EffacementResultatsEnCours = True
        MsgBox Feuil55.Range("Investissement_erreur_choix_step_ref").Cells(1, 1), Buttons:=vbCritical, Title:=Feuil55.Range("Investissement_erreur_titre").Cells(1, 1)
        Exit Sub
    End If

    'lecture des OPEX calculées (plus tard seront transmises dans le calcul)
    For I = 1 To nb_step_comparaison
        Colonne = 3 * pi_CALCUL_VALEUR_GUIDE + I
        If Feuil55.Range("Investissement_OPEX_devise_an").Cells(1, I) <> 0 Then
            OPEX_tous_an(I) = Feuil55.Range("Investissement_OPEX_devise_an").Cells(1, I)
        Else
            OPEX_tous_an(I) = Feuil55.Range("Investissement_OPEX_devise_an").Cells(1, Colonne)
        End If
    Next I
    OPEX_reference_an = OPEX_tous_an(choix_step_reference)

    'lecture des investissements
    For I = 1 To nb_step_comparaison
        Colonne = 3 * pi_EtapeCalculConsideree + I
        cout_investissement(I) = Feuil55.Range("Investissement_CAPEX_devise").Cells(1, I)
    Next I

    invest_ref = cout_investissement(choix_step_reference)
    invest_min = invest_ref
    For I = 1 To nb_step_comparaison
        If I <> choix_step_reference Then
            If cout_investissement(I) < invest_ref Then
                invest_min = cout_investissement(I)
            End If
        End If
    Next I
    If invest_min < invest_ref Then
        pn_EffacementResultatsEnCours = True
        MsgBox Feuil55.Range("Investissement_erreur_choix_step_ref").Cells(1, 1), Buttons:=vbCritical, Title:=Feuil55.Range("Investissement_erreur_titre").Cells(1, 1)
        Exit Sub
    End If

    'lecture de la durée du contrat et du taux d'actualisation
    For I = 1 To nb_step_comparaison
        Colonne = I
        duree_contrat_an(I) = Feuil55.Range("Investissement_contrat_annees").Cells(1, Colonne)
    Next I
    For I = 1 To nb_step_comparaison
        Colonne = 3 * pi_EtapeCalculConsideree + I
        taux_actualisation(I) = Feuil55.Range("Investissement_taux_actualisation").Cells(1, Colonne)
    Next I

End Sub


Private Sub Class_Initialize()

    Dim I As Integer

    For I = 1 To 3
        retour_invest_effectue(I) = False
    Next I

    If pn_EffacementResultatsEnCours = False Then
        Call lecture_choix
    End If


End Sub

Private Sub Class_Terminate()

    If pn_EffacementResultatsEnCours = False Then
        Call ecrire_resultats
    End If

End Sub


