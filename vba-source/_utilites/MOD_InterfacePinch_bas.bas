Attribute VB_Name = "MOD_InterfacePinch"
'#### DECLARATION DES VARIABLES POUR L'INTERFACE

'**** à effacer avant intégration OCEAN
'Public nb_step_comparaison As Integer
'****

'### Variables communes
Public numStationSelectionnee As Integer
Public colonneStation As Integer
Public Const INDICE_PREMIERE_LIGNE = 4

'### Variables de la section "Gestion des procédés"
Public modeModif_proc As Boolean
Public indiceLVProcedeLigneASupprimer As Integer
Public indiceLVProcedeLigneAModifier As Integer

'### Variables de la section "Interdiction d'un échange"
Public modeModif_inter As Boolean
Public fluxChaudInterdiction As String
Public fluxFroidInterdiction As String
Public indiceLVInterdictionLigneASupprimer As Integer
Public indiceLVInterdictionLigneAModifier As Integer

'### Variables de la section "Imposition d'un échange"
Public modeModif_impos As Boolean
Public fluxChaudImpose As String
Public fluxFroidImpose As String
Public indiceLVImpositionLigneASupprimer As Integer
Public indiceLVImpositionLigneAModifier As Integer

'Bornes des températures et MCp des fluxs
Public TFCMin As Double
Public TFCMax As Double
Public TFFMin As Double
Public TFFMax As Double
Public MCpc As Double
Public MCpf As Double

'Variables pour les tests des bornes et du croisement
Public bornesValidesValeursSaisies As Boolean
Public croisementValideValeursSaisies As Boolean
Public bornesValidesValeursCalculees As Boolean
Public croisementValideValeursCalculees As Boolean

Public erreursValeursBornesSaisies As String    'Message d'erreur bornes saisies
Public erreursValeursCroisementSaisies As String    'Message d'erreur croisement saisie
Public erreursValeursBornesCalculees As String    'Message d'erreur bornes calculées
Public erreursValeursCroisementCalculees As String    'Message d'erreur croisement calculées

Public erreurTemperature() As String
Public tabErreurFluxProc(1 To 3) As Boolean
Public erreurFluxProc As String
Public echangeursValides As Boolean
Public donneesFluxProcValides As Boolean
Public erreurFluxOuvert As Boolean
Public donneesPincementSeuilValides As Boolean
Public erreurDonneesPincementSeuil As String

'Paramètres saisis et calculés à travers l'interface
Public param1 As Double
Public param2 As Double
Public param3 As Double
Public tempEFC As Double
Public tempSFC As Double
Public tempEFF As Double
Public tempSFF As Double
Public puissance As Double
Public nomsCombin(1 To 3) As String
Public combin As Integer

'### Variables de la section "Gestion des utilités"
Public modeModif_util As Boolean
Public indiceLVUtiliteLigneASupprimer As Integer
Public indiceLVUtiliteLigneAModifier As Integer

Public Sub ecriture_nom_flux_interdiction_echange(ByVal NomFlux, ByVal NomFeuilleEcriture As String, ByVal NomPlageEcriture As String, ByVal LigneEcriture As Integer, ByVal ColonneEcriture As Integer, ByVal NomFeuilleEtiquette As String, ByVal NomPlageEtiquette As String)
'Ajout DCA 25/05/2014 pour changement de langue

Dim LigneEtiquette As Integer
Dim ColonneEtiquette As Integer
Dim EtiquetteTrouvee As Boolean
Dim NomEtiquette As String

LigneEtiquette = 1
ColonneEtiquette = 1
EtiquetteTrouvee = False
Do Until Sheets(NomFeuilleEtiquette).Range(NomPlageEtiquette).Cells(LigneEtiquette, ColonneEtiquette) = "" Or EtiquetteTrouvee = True
    NomEtiquette = Sheets(NomFeuilleEtiquette).Range(NomPlageEtiquette).Cells(LigneEtiquette, ColonneEtiquette)
    If Range(NomEtiquette).Cells(1, 1) = NomFlux Then
        Sheets(NomFeuilleEcriture).Range(NomPlageEcriture).Cells(LigneEcriture, ColonneEcriture).Formula = "= " & NomEtiquette
        EtiquetteTrouvee = True
    End If
    LigneEtiquette = LigneEtiquette + 1
Loop
If EtiquetteTrouvee = False Then
    Sheets(NomFeuilleEcriture).Range(NomPlageEcriture).Cells(LigneEcriture, ColonneEcriture) = NomFlux
End If

End Sub

Sub interfacePinch()

    uf_choixPincement.Show

End Sub

Sub miseAJourLV()
    '--> Met à jour l'ensemble des listviews

    'Alimentation de la LV procedes
    alimentationLVProcedes

    'Alimentation de la LV utilites
    alimentationLVUtilites

    'Alimentation états des tris
    alimentationEtatTriAuto

    'Alimentation de la LV Interdiction
    alimentationLVFluxInterdiction

    'Alimentation de la LV Imposition
    alimentationLVFluxImposition

    'Alimentation de la TB Valeur pincement
    alimentationTBValeurPincement

    'Alimentation de la TB Seuil puissance
    alimentationTBSeuilPuissance

End Sub

Sub alimentationEtatTriAuto()

    Dim LigneTab As Integer

    colonneStation = 1 + (numStationSelectionnee - 1)
    LigneTab = INDICE_PREMIERE_LIGNE

    'Tri automatique flux process
    uf_choixPincement.cb_choixTriProcedesAuto.Value = Feuil63.Range("DonneesPincementTriAutoProcedes").Cells(LigneTab, colonneStation).Value

    'Tri automatique flux utilites
    uf_choixPincement.cb_choixTriUtilitesAuto.Value = Feuil63.Range("DonneesPincementTriAutoUtilites").Cells(LigneTab, colonneStation).Value

End Sub

Sub alimentationLVFluxImposition()
    '#### remplissage de la list view à partir des données du tableau

    '### DECLARATION DES VARIABLES
    Dim LigneTab As Integer
    Dim ligneLV As Integer
    Dim colonneLV As Integer
    Dim COL_NOM As Boolean
    Dim TinFChaud As Double
    Dim ToutFChaud As Double
    Dim TinFFroid As Double
    Dim ToutFFroid As Double
    Dim Puiss As Double
    Dim nomFluxChaud As String
    Dim nomFluxFroid As String
    Dim WIDTH_LV As Double

    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * 9
    LigneTab = INDICE_PREMIERE_LIGNE
    ligneLV = 1

    '### PROCEDURE
    With uf_choixPincement

        With .lv_choixFluxImposition
            '## Mise en forme des colonnes
            'Réinitialisation du ListView
            .ListItems.Clear
            .ColumnHeaders.Clear

            'Paramètres d'affichages du ListView
            .View = 3    'lvwReport
            .Gridlines = True
            .CheckBoxes = True
            .FullRowSelect = True
            .LabelEdit = lvwManual

            'Création des colonnes
            WIDTH_LV = .Width    'largeur du LV
            With .ColumnHeaders
                .Add , , Feuil5.Range("l_fluxChaud").Value, (WIDTH_LV / 9) * 2
                .Add , , Feuil5.Range("l_TempEntreeFluxChaud").Value & " " & Feuil5.Range("l_uniteDegCelcius").Value, WIDTH_LV / 9, lvwColumnCenter
                .Add , , Feuil5.Range("l_TempSortieFluxChaud").Value & " " & Feuil5.Range("l_uniteDegCelcius").Value, WIDTH_LV / 9, lvwColumnCenter
                .Add , , Feuil5.Range("l_fluxFroid").Value, (WIDTH_LV / 9) * 2, lvwColumnCenter
                .Add , , Feuil5.Range("l_TempEntreeFluxFroid").Value & " " & Feuil5.Range("l_uniteDegCelcius").Value, WIDTH_LV / 9, lvwColumnCenter
                .Add , , Feuil5.Range("l_TempSortieFluxFroid").Value & " " & Feuil5.Range("l_uniteDegCelcius").Value, WIDTH_LV / 9, lvwColumnCenter
                .Add , , Feuil5.Range("l_PuissanceEchange").Value & " " & Feuil5.Range("l_uniteKiloWatt").Value, WIDTH_LV / 9, lvwColumnCenter
            End With

            'remplissage ligne par ligne
            While Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation) <> ""

                nomFluxChaud = Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation + 1)
                nomFluxFroid = Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation + 2)
                TinFChaud = Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation + 3)
                ToutFChaud = Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation + 4)
                TinFFroid = Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation + 5)
                ToutFFroid = Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation + 6)
                Puiss = Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation + 7)

                'l'ordre des paramètres est changé pour simplifier la compréhension
                .ListItems.Add , , nomFluxChaud
                .ListItems(ligneLV).ListSubItems.Add , , Format(TinFChaud, "#0.00")
                .ListItems(ligneLV).ListSubItems.Add , , Format(ToutFChaud, "#0.00")

                .ListItems(ligneLV).ListSubItems.Add , , nomFluxFroid
                .ListItems(ligneLV).ListSubItems.Add , , Format(TinFFroid, "#0.00")
                .ListItems(ligneLV).ListSubItems.Add , , Format(ToutFFroid, "#0.00")
                .ListItems(ligneLV).ListSubItems.Add , , Format(Puiss, "#0.00")

                ' coloriage des lignes
                .ListItems(ligneLV).ForeColor = RGB(255, 0, 0)     'rouge nomFluxChaud
                .ListItems(ligneLV).ListSubItems(1).ForeColor = RGB(255, 0, 0)   'rouge TinFChaud
                .ListItems(ligneLV).ListSubItems(2).ForeColor = RGB(255, 0, 0)   'rouge ToutFChaud

                .ListItems(ligneLV).ListSubItems(3).ForeColor = RGB(0, 0, 255)   'bleu nomFluxFroid
                .ListItems(ligneLV).ListSubItems(4).ForeColor = RGB(0, 0, 255)   'bleu TinFFroid
                .ListItems(ligneLV).ListSubItems(5).ForeColor = RGB(0, 0, 255)   'bleu ToutFFroid

                'cochage de la case
                'Si la cellule est à l'état TRUE on coche la ligne correspondante dans le LV
                If Feuil63.Range("DonneesPincementImpositionEchange").Cells(ligneLV + (INDICE_PREMIERE_LIGNE - 1), colonneStation).Value = True Then
                    .ListItems(ligneLV).Checked = True
                End If

                LigneTab = LigneTab + 1
                ligneLV = ligneLV + 1
            Wend
        End With
        'Astuce permettant de repositionner le LV
        .top = .top + 1
        .top = .top - 1
    End With

End Sub

Sub alimentationLVFluxInterdiction()
    'remplissage de la list view à partir des données du tableau

    '### DECLARATION DES VARIABLES
    Dim LigneTab As Integer
    Dim ligneLV As Integer
    Dim colonneLV As Integer
    Dim COL_NOM As Boolean
    Dim nomFluxChaud As String
    Dim nomFluxFroid As String
    Dim WIDTH_LV As Double

    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * 3
    LigneTab = INDICE_PREMIERE_LIGNE
    ligneLV = 1

    '### PROCEDURE
    'Mise en forme des colonnes
    With uf_choixPincement

        With .lv_choixFluxInterdiction

            'raz LV
            .ListItems.Clear
            .ColumnHeaders.Clear

            'affichage LV
            .View = 3    'lvwReport
            .Gridlines = True
            .CheckBoxes = True
            .FullRowSelect = True
            .LabelEdit = lvwManual

            'Création des colonnes
            WIDTH_LV = .Width
            With .ColumnHeaders
                .Add , , Feuil5.Range("l_fluxChaud").Value, WIDTH_LV / 2
                .Add , , Feuil5.Range("l_fluxFroid").Value, WIDTH_LV / 2, lvwColumnCenter
            End With

            'remplissage ligne par ligne

            While Feuil63.Range("DonneesPincementInterdictionEchange").Cells(LigneTab, colonneStation) <> ""

                nomFluxChaud = Feuil63.Range("DonneesPincementInterdictionEchange").Cells(LigneTab, colonneStation + 1)
                nomFluxFroid = Feuil63.Range("DonneesPincementInterdictionEchange").Cells(LigneTab, colonneStation + 2)

                .ListItems.Add , , nomFluxChaud
                .ListItems(ligneLV).ListSubItems.Add , , nomFluxFroid

                ' coloriage des lignes
                .ListItems(ligneLV).ForeColor = RGB(255, 0, 0)     'rouge
                .ListItems(ligneLV).ListSubItems(1).ForeColor = RGB(0, 0, 255)    'rouge

                'cochage de la case
                'Si la cellule est à l'état TRUE on coche la ligne correspondante dans le LV
                If Feuil63.Range("DonneesPincementInterdictionEchange").Cells(ligneLV + (INDICE_PREMIERE_LIGNE - 1), colonneStation).Value = True Then
                    .ListItems(ligneLV).Checked = True
                End If

                LigneTab = LigneTab + 1
                ligneLV = ligneLV + 1
            Wend

        End With
        'Astuce permettant de repositionner le LV
        .top = .top + 1
        .top = .top - 1
    End With

End Sub

Sub alimentationLVProcedes()

    'remplissage de la list view à partir des données du tableau

    '### DECLARATION DES VARIABLES
    Dim LigneTab As Integer
    Dim ligneLV As Integer
    Dim colonneLV As Integer
    Dim COL_NOM As Boolean
    Dim nom As String
    Dim Tin As Double, Tout As Double, MCp As Double, DH As Double
    Dim WIDTH_LV As Double

    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * flux_thermique_process_nb_caracteristiques
    LigneTab = INDICE_PREMIERE_LIGNE
    ligneLV = 1

    '### PROCEDURE
    'création des colonnes
    With uf_choixPincement

        With .lv_choixFluxProcedes
            'raz LV
            .ListItems.Clear
            .ColumnHeaders.Clear

            'affichage LV
            .View = 3    'lvwReport
            .Gridlines = True
            .CheckBoxes = True
            .FullRowSelect = True
            .LabelEdit = lvwManual

            'Création des colonnes
            WIDTH_LV = .Width
            With .ColumnHeaders
                ' .Clear
                .Add , , Feuil5.Range("l_nomFluxProcede").Value, (WIDTH_LV / flux_thermique_process_nb_caracteristiques) * 2
                .Add , , Feuil5.Range("l_temperatureEntree2").Value & " " & Feuil5.Range("l_uniteDegCelcius").Value, WIDTH_LV / flux_thermique_process_nb_caracteristiques, lvwColumnCenter
                .Add , , Feuil5.Range("l_temperatureSortie2").Value & " " & Feuil5.Range("l_uniteDegCelcius").Value, WIDTH_LV / flux_thermique_process_nb_caracteristiques, lvwColumnCenter
                .Add , , Feuil5.Range("l_MCp").Value & " " & Feuil5.Range("l_uniteMCp").Value, WIDTH_LV / flux_thermique_process_nb_caracteristiques, lvwColumnCenter
                .Add , , Feuil5.Range("l_deltaH").Value & " " & Feuil5.Range("l_uniteKiloWatt").Value, WIDTH_LV / flux_thermique_process_nb_caracteristiques, lvwColumnCenter
            End With

            'remplissage ligne par ligne

            While Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation) <> ""

                nom = Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation + 1)
                Tin = Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation + 2)
                Tout = Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation + 3)
                MCp = Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation + 4)
                DH = MCp * (Tin - Tout)

                .ListItems.Add , , nom
                .ListItems(ligneLV).ListSubItems.Add , , Format(Tin, "#0.00")
                .ListItems(ligneLV).ListSubItems.Add , , Format(Tout, "#0.00")
                .ListItems(ligneLV).ListSubItems.Add , , Format(MCp, "#0.00")
                .ListItems(ligneLV).ListSubItems.Add , , Format(DH, "#0.00")

                COL_NOM = False

                ' coloriage des lignes
                For colonneLV = 1 To 4
                    If DH > 0 Then    'flux chaud
                        If COL_NOM = False Then
                            .ListItems(ligneLV).ForeColor = RGB(255, 0, 0)    'rouge
                            COL_NOM = True
                        End If
                        .ListItems(ligneLV).ListSubItems(colonneLV).ForeColor = RGB(255, 0, 0)    'rouge
                    Else    'flux froid
                        If COL_NOM = False Then
                            .ListItems(ligneLV).ForeColor = RGB(0, 0, 255)    'bleu
                            COL_NOM = True
                        End If
                        .ListItems(ligneLV).ListSubItems(colonneLV).ForeColor = RGB(0, 0, 255)    'bleu
                    End If
                Next

                ' Actualisation de l'état des cases à cocher
                If Feuil63.Range("DonneesPincementFluxProcedes").Cells(ligneLV + (INDICE_PREMIERE_LIGNE - 1), colonneStation).Value = True Then
                    .ListItems(ligneLV).Checked = True
                End If

                LigneTab = LigneTab + 1
                ligneLV = ligneLV + 1
            Wend

        End With
        'Astuce permettant de repositionner le LV
        .top = .top + 1
        .top = .top - 1
    End With

End Sub

Sub alimentationLVUtilites()
    'remplissage de la list view à partir des données du tableau

    '### DECLARATION DES VARIABLES
    Dim LigneTab As Integer
    Dim ligneLV As Integer
    Dim colonneLV As Integer
    Dim COL_NOM As Boolean
    Dim nom, tipe As String
    Dim Tin As Double, MCp As Double, Ptot As Double


    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * flux_thermique_utilites_nb_caracteristiques
    LigneTab = INDICE_PREMIERE_LIGNE
    ligneLV = 1

    '### PROCEDURE
    'création des colonnes
    With uf_choixPincement

        With .lv_choixFluxUtilites
            'raz LV
            .ListItems.Clear
            .ColumnHeaders.Clear

            'affichage LV
            .View = 3    'lvwReport
            .Gridlines = True
            .CheckBoxes = True
            .FullRowSelect = True
            .LabelEdit = lvwManual


            'Création des colonnes
            WIDTH_LV = .Width
            With .ColumnHeaders
                ' .Clear
                .Add , , Feuil5.Range("l_nomFluxProcede").Value, (WIDTH_LV / 5) * 2
                .Add , , Feuil5.Range("l_temperatureEntree2").Value & " " & Feuil5.Range("l_uniteDegCelcius").Value, WIDTH_LV / 5, lvwColumnCenter
                .Add , , Feuil5.Range("l_MCp").Value & " " & Feuil5.Range("l_uniteMCp").Value, WIDTH_LV / 5, lvwColumnCenter
                .Add , , Feuil5.Range("l_puissanceDisponible").Value & " " & Feuil5.Range("l_uniteKiloWatt").Value, WIDTH_LV / 5, lvwColumnCenter
            End With

            'remplissage ligne par ligne

            While Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation) <> ""

                nom = Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation + 1)
                Tin = Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation + 3)
                MCp = Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation + 4)
                Ptot = Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation + 5)

                .ListItems.Add , , nom
                .ListItems(ligneLV).ListSubItems.Add , , Format(Tin, "#0.00")
                .ListItems(ligneLV).ListSubItems.Add , , Format(MCp, "#0.00")
                .ListItems(ligneLV).ListSubItems.Add , , Format(Ptot, "#0.00")

                COL_NOM = False

                ' coloriage des lignes
                tipe = Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation + 2)
                For colonneLV = 1 To 3
                    If tipe = "Ch" Then  'flux chaud
                        If COL_NOM = False Then
                            .ListItems(ligneLV).ForeColor = RGB(255, 0, 0)    'rouge
                            COL_NOM = True
                        End If
                        .ListItems(ligneLV).ListSubItems(colonneLV).ForeColor = RGB(255, 0, 0)    'rouge
                    ElseIf tipe = "Fr" Then   'flux froid
                        If COL_NOM = False Then
                            .ListItems(ligneLV).ForeColor = RGB(0, 0, 255)    'bleu
                            COL_NOM = True
                        End If
                        .ListItems(ligneLV).ListSubItems(colonneLV).ForeColor = RGB(0, 0, 255)    'bleu
                    End If
                Next

                'cochage de la case
                'Si la cellule est à l'état TRUE on coche la ligne correspondante dans le LV
                If Feuil63.Range("DonneesPincementFluxUtilites").Cells(ligneLV + (INDICE_PREMIERE_LIGNE - 1), colonneStation).Value = True Then
                    .ListItems(ligneLV).Checked = True
                End If

                LigneTab = LigneTab + 1
                ligneLV = ligneLV + 1
            Wend

        End With
        'Astuce permettant de repositionner le LV
        .top = .top + 1
        .top = .top - 1
    End With

End Sub

Sub nomsCombinaisonParam(combinaison As Integer)

    '### DECLARATION DES VARIABLES
    Dim noms(1 To 5) As String
    Dim combinaisons As Variant
    Dim indice As Integer
    Dim indiceDebut As Integer
    Dim I As Integer

    '### INITIALISATION DES VARIABLES
    indiceDebut = 1

    '### PROCEDURE
    'Chargement des noms de paramètre
    '    For indice = indiceDebut To indiceDebut + 4
    '        noms(indice) = (Feuil5.Range("B" & indice + 1))
    '    Next indice
    noms(1) = Feuil5.Range("l_Te_ch").Value
    noms(2) = Feuil5.Range("l_Ts_ch").Value
    noms(3) = Feuil5.Range("l_Te_fr").Value
    noms(4) = Feuil5.Range("l_Ts_fr").Value
    noms(5) = Feuil5.Range("l_P").Value

    'Chargement des combinaisons de paramètres (1:TCin; 2:TCout; 3:TFin; 4:TFout; 5:P)
    combinaisons = Array(123, 124, 134, 234, 135, 145, 235, 245)

    ' Affectation des noms en fonction de la combinaison
    ReDim nomsCombinaisonParam(1 To 3)
    For I = 1 To 3
        nomsCombin(I) = noms(Mid(combinaisons(combinaison), I, 1))
    Next I

End Sub

Sub alimentationLBCombinaisons()

    '### DECLARATION DES VARIABLES
    Dim noms(1 To 5) As String
    Dim combinaisons As Variant
    Dim indice, indiceDebut, I As Integer

    '### INITIALISATION DES VARIABLES
    indiceDebut = 1

    '### PROCEDURE
    'Chargement des noms de paramètre depuis la feuille "gestion choix"
    '    For indice = indiceDebut To indiceDebut + 4
    '        noms(indice) = (Feuil5.Range("B" & indice + 1))
    '    Next indice
    noms(1) = Feuil5.Range("l_Te_ch").Value
    noms(2) = Feuil5.Range("l_Ts_ch").Value
    noms(3) = Feuil5.Range("l_Te_fr").Value
    noms(4) = Feuil5.Range("l_Ts_fr").Value
    noms(5) = Feuil5.Range("l_P").Value

    'Chargement des combinaisons de paramètres (1:TCin; 2:TCout; 3:TFin; 4:TFout; 5:P)
    combinaisons = Array(123, 124, 134, 234, 135, 145, 235, 245)

    ' remplissage de la list box
    Dim affichageCombinaison As String
    indiceDebut = 0
    For indice = indiceDebut To indiceDebut + 7
        affichageCombinaison = noms(Mid(combinaisons(indice), 1, 1)) & "  /  " & noms(Mid(combinaisons(indice), 2, 1)) & "  /  " & noms(Mid(combinaisons(indice), 3, 1))
        uf_choixFluxImposition.lb_choixCombinaisonParam.AddItem affichageCombinaison
    Next indice

End Sub

Sub informationsFC()

    '### DECLARATION DES VARIABLES
    Dim LigneTab As Integer

    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * flux_thermique_process_nb_caracteristiques
    LigneTab = INDICE_PREMIERE_LIGNE

    '### PROCEDURE
    While Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation + 1) <> uf_choixFluxImposition.cb_choixFluxChaud.Value
        LigneTab = LigneTab + 1    'incrémentation de l'indice de la ligne
    Wend

    TFCMax = Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation + 2)
    TFCMin = Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation + 3)
    MCpc = Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation + 4)

    'Affichage des informations sur le flux sélectionné
    affichageInfosFlux

End Sub

Sub informationsFF()

    '### DECLARATION DES VARIABLES
    Dim LigneTab As Integer

    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * flux_thermique_process_nb_caracteristiques
    LigneTab = INDICE_PREMIERE_LIGNE

    '### PROCEDURE
    While Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation + 1) <> uf_choixFluxImposition.cb_choixFluxFroid.Value
        LigneTab = LigneTab + 1    'incrémentation de l'indice de la ligne
    Wend

    TFFMin = Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation + 2)
    TFFMax = Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation + 3)
    MCpf = Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation + 4)

    'Affichage des informations sur le flux sélectionné
    affichageInfosFlux

End Sub

Sub affichageInfosFlux()

    uf_choixFluxImposition.l_bornesFluxChaud.Caption = Format(TFCMax, "#0.00") & Feuil5.Range("l_uniteDegCelcius").Value & ">" & Feuil5.Range("l_T").Value & ">" & Format(TFCMin, "#0.00") & Feuil5.Range("l_uniteDegCelcius").Value & "  " & Feuil5.Range("l_MCp").Value & " = " & Format(MCpc, "#0.00") & " " & Feuil5.Range("l_uniteMCp").Value
    uf_choixFluxImposition.l_bornesFluxFroid.Caption = Format(TFFMin, "#0.00") & Feuil5.Range("l_uniteDegCelcius").Value & "<" & Feuil5.Range("l_T").Value & "<" & Format(TFFMax, "#0.00") & Feuil5.Range("l_uniteDegCelcius").Value & "  " & Feuil5.Range("l_MCp").Value & " = " & Format(MCpf, "#0.00") & " " & Feuil5.Range("l_uniteMCp").Value

End Sub

Function PuissanceFluxChaud(ByVal temperature_entree As Double, ByVal temperature_sortie As Double, ByVal MCp_chaud As Double)
    'retourne la puissance : P = MCp * deltaT

    PuissanceFluxChaud = (temperature_entree - temperature_sortie) * MCp_chaud

End Function

Function PuissanceFluxFroid(ByVal temperature_entree As Double, ByVal temperature_sortie As Double, ByVal MCp_froid As Double)
    'retourne la puissance : MCp * deltaT

    PuissanceFluxFroid = (temperature_sortie - temperature_entree) * MCp_froid

End Function

Function TemperatureEntreeFluxChaud(ByVal temperature_sortie_flux_chaud As Double, ByVal puissance As Double, ByVal MCp_chaud As Double)
    'retourne la température d'entrée (flux chaud)

    TemperatureEntreeFluxChaud = temperature_sortie_flux_chaud + (puissance / MCp_chaud)

End Function

Function TemperatureEntreeFluxFroid(ByVal temperature_sortie_flux_froid As Double, ByVal puissance As Double, ByVal MCp_froid As Double)
    'retourne la température d'entrée (flux froid)

    TemperatureEntreeFluxFroid = temperature_sortie_flux_froid - (puissance / MCp_froid)

End Function

Function TemperatureSortieFluxChaud(ByVal temperature_entree_flux_chaud As Double, ByVal puissance As Double, ByVal MCp_chaud As Double)
    'retourne la température de sortie (flux chaud)

    TemperatureSortieFluxChaud = temperature_entree_flux_chaud - (puissance / MCp_chaud)

End Function

Function TemperatureSortieFluxFroid(ByVal temperature_entree_flux_froid As Double, ByVal puissance As Double, ByVal MCp_froid As Double)
    'retourne la température de sortie (flux froid)

    TemperatureSortieFluxFroid = temperature_entree_flux_froid + (puissance / MCp_froid)

End Function

Sub affichageParametresInformations()

    With uf_choixFluxImposition
        ' Couleurs flux chaud:
        .l_nomFChaud.ForeColor = RGB(255, 0, 0)     'rouge
        .l_Tcin.ForeColor = RGB(255, 0, 0)     'rouge
        .l_Tcout.ForeColor = RGB(255, 0, 0)     'rouge
        .l_flecheFC.ForeColor = RGB(255, 0, 0)     'rouge

        ' Couleurs flux froid:
        .l_nomFFroid.ForeColor = RGB(0, 0, 255)     'bleu
        .l_Tfin.ForeColor = RGB(0, 0, 255)     'bleu
        .l_Tfout.ForeColor = RGB(0, 0, 255)     'bleu
        .l_flecheFF.ForeColor = RGB(0, 0, 255)     'bleu

        'Labels schéma échangeur
        .l_Tcin = Feuil5.Range("msgE_tempEFC") & " : " & Format(tempEFC, "#0.00")
        .l_Tcout = Feuil5.Range("msgE_tempSFC") & " : " & Format(tempSFC, "#0.00")
        .l_Tfin = Feuil5.Range("msgE_tempEFF") & " : " & Format(tempEFF, "#0.00")
        .l_Tfout = Feuil5.Range("msgE_tempSFF") & " : " & Format(tempSFF, "#0.00")
        .l_puiss = Feuil5.Range("l_puissance") & " : " & Format(puissance, "#0.00")

        
    End With

End Sub

Sub calculCaracteristiquesEchangeur()

    If version_Developpement = False Then On Error GoTo suite

    '### PROCEDURE
    Select Case combin

    Case 123
        'Récupération des informations saisies
        tempEFC = uf_choixFluxImposition.tb_param1.Value
        tempSFC = uf_choixFluxImposition.tb_param2.Value
        tempEFF = uf_choixFluxImposition.tb_param3.Value

        'Vérification de la validité des informations saisies
        testerBornesValeursSaisies     ' Vérification des bornes
        testerCroisementValeursSaisies     ' Vérification du croisement

        'Calcul des paramètres restants
        If bornesValidesValeursSaisies = True And croisementValideValeursSaisies = True Then
            puissance = PuissanceFluxChaud(tempEFC, tempSFC, MCpc)
            tempSFF = TemperatureSortieFluxFroid(tempEFF, puissance, MCpf)

            'Vérification de la validité des informations calculées
            testerBornesValeursCalculees    ' Vérification des bornes
            testerCroisementValeursCalculees     ' Vérification du croisement
        End If

    Case 124
        'Récupération des informations saisies
        tempEFC = uf_choixFluxImposition.tb_param1.Value
        tempSFC = uf_choixFluxImposition.tb_param2.Value
        tempSFF = uf_choixFluxImposition.tb_param3.Value

        'Vérification de la validité des informations saisies
        testerBornesValeursSaisies     ' Vérification des bornes
        testerCroisementValeursSaisies     ' Vérification du croisement

        'Calcul des paramètres restants
        If bornesValidesValeursSaisies = True And croisementValideValeursSaisies = True Then
            puissance = PuissanceFluxChaud(tempEFC, tempSFC, MCpc)
            tempEFF = TemperatureEntreeFluxFroid(tempSFF, puissance, MCpf)

            'Vérification de la validité des informations calculées
            testerBornesValeursCalculees    ' Vérification des bornes
            testerCroisementValeursCalculees     ' Vérification du croisement
        End If

    Case 134
        'Récupération des informations saisies
        tempEFC = uf_choixFluxImposition.tb_param1.Value
        tempEFF = uf_choixFluxImposition.tb_param2.Value
        tempSFF = uf_choixFluxImposition.tb_param3.Value

        'Vérification de la validité des informations saisies
        testerBornesValeursSaisies     ' Vérification des bornes
        testerCroisementValeursSaisies     ' Vérification du croisement

        'Calcul des paramètres restants
        If bornesValidesValeursSaisies = True And croisementValideValeursSaisies = True Then
            puissance = PuissanceFluxFroid(tempEFF, tempSFF, MCpf)
            tempSFC = TemperatureSortieFluxChaud(tempEFC, puissance, MCpc)

            'Vérification de la validité des informations calculées
            testerBornesValeursCalculees    ' Vérification des bornes
            testerCroisementValeursCalculees     ' Vérification du croisement
        End If

    Case 234
        'Récupération des informations saisies
        tempSFC = uf_choixFluxImposition.tb_param1.Value
        tempEFF = uf_choixFluxImposition.tb_param2.Value
        tempSFF = uf_choixFluxImposition.tb_param3.Value

        'Vérification de la validité des informations saisies
        testerBornesValeursSaisies     ' Vérification des bornes
        testerCroisementValeursSaisies     ' Vérification du croisement

        'Calcul des paramètres restants
        If bornesValidesValeursSaisies = True And croisementValideValeursSaisies = True Then
            puissance = PuissanceFluxFroid(tempEFF, tempSFF, MCpf)
            tempEFC = TemperatureEntreeFluxChaud(tempSFC, puissance, MCpc)

            'Vérification de la validité des informations calculées
            testerBornesValeursCalculees    ' Vérification des bornes
            testerCroisementValeursCalculees     ' Vérification du croisement
        End If

    Case 135
        'Récupération des informations saisies
        tempEFC = uf_choixFluxImposition.tb_param1.Value
        tempEFF = uf_choixFluxImposition.tb_param2.Value
        puissance = uf_choixFluxImposition.tb_param3.Value

        'Vérification de la validité des informations saisies
        testerBornesValeursSaisies     ' Vérification des bornes
        testerCroisementValeursSaisies     ' Vérification du croisement

        'Calcul des paramètres restants
        If bornesValidesValeursSaisies = True And croisementValideValeursSaisies = True Then
            tempSFC = TemperatureSortieFluxChaud(tempEFC, puissance, MCpc)
            tempSFF = TemperatureSortieFluxFroid(tempEFF, puissance, MCpf)

            'Vérification de la validité des informations calculées
            testerBornesValeursCalculees    ' Vérification des bornes
            testerCroisementValeursCalculees     ' Vérification du croisement
        End If

    Case 145
        'Récupération des informations saisies
        tempEFC = uf_choixFluxImposition.tb_param1.Value
        tempSFF = uf_choixFluxImposition.tb_param2.Value
        puissance = uf_choixFluxImposition.tb_param3.Value

        'Vérification de la validité des informations saisies
        testerBornesValeursSaisies     ' Vérification des bornes
        testerCroisementValeursSaisies     ' Vérification du croisement

        'Calcul des paramètres restants
        If bornesValidesValeursSaisies = True And croisementValideValeursSaisies = True Then
            tempSFC = TemperatureSortieFluxChaud(tempEFC, puissance, MCpc)
            tempEFF = TemperatureEntreeFluxFroid(tempSFF, puissance, MCpf)

            'Vérification de la validité des informations calculées
            testerBornesValeursCalculees    ' Vérification des bornes
            testerCroisementValeursCalculees     ' Vérification du croisement
        End If

    Case 235
        'Récupération des informations saisies
        tempSFC = uf_choixFluxImposition.tb_param1.Value
        tempEFF = uf_choixFluxImposition.tb_param2.Value
        puissance = uf_choixFluxImposition.tb_param3.Value

        'Vérification de la validité des informations saisies
        testerBornesValeursSaisies     ' Vérification des bornes
        testerCroisementValeursSaisies     ' Vérification du croisement

        'Calcul des paramètres restants
        If bornesValidesValeursSaisies = True And croisementValideValeursSaisies = True Then
            tempEFC = TemperatureEntreeFluxChaud(tempSFC, puissance, MCpc)
            tempSFF = TemperatureSortieFluxFroid(tempEFF, puissance, MCpf)

            'Vérification de la validité des informations calculées
            testerBornesValeursCalculees    ' Vérification des bornes
            testerCroisementValeursCalculees     ' Vérification du croisement
        End If

    Case 245
        'Récupération des informations saisies
        tempSFC = uf_choixFluxImposition.tb_param1.Value
        tempSFF = uf_choixFluxImposition.tb_param2.Value
        puissance = uf_choixFluxImposition.tb_param3.Value

        'Vérification de la validité des informations saisies
        testerBornesValeursSaisies     ' Vérification des bornes
        testerCroisementValeursSaisies     ' Vérification du croisement

        'Calcul des paramètres restants
        If bornesValidesValeursSaisies = True And croisementValideValeursSaisies = True Then
            tempEFC = TemperatureEntreeFluxChaud(tempSFC, puissance, MCpc)
            tempEFF = TemperatureEntreeFluxFroid(tempSFF, puissance, MCpf)

            'Vérification de la validité des informations calculées
            testerBornesValeursCalculees    ' Vérification des bornes
            testerCroisementValeursCalculees     ' Vérification du croisement
        End If

    End Select

    'Mise à jour du schéma
    affichageParametresInformations

    Exit Sub
suite:

End Sub

Sub testerBornesValeursSaisies()

    '### DECLARATION DES VARIABLES
    erreursValeursBornesSaisies = ""
    bornesValidesValeursSaisies = True

    '### PROCEDURE
    Select Case combin

    Case 123
        '## Récupération des informations saisies
        tempEFC = uf_choixFluxImposition.tb_param1.Value
        tempSFC = uf_choixFluxImposition.tb_param2.Value
        tempEFF = uf_choixFluxImposition.tb_param3.Value

        '## Vérification de la validité des informations saisies
        If tempEFC > TFCMax Or tempEFC < TFCMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempEFC")    ' "EFC"
            bornesValidesValeursSaisies = False
        End If

        If tempSFC > TFCMax Or tempSFC < TFCMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempSFC")    ' "SFC"
            bornesValidesValeursSaisies = False
        End If

        If tempEFF > TFFMax Or tempEFF < TFFMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempEFF")    ' "EFF"
            bornesValidesValeursSaisies = False
        End If

    Case 124
        '## Récupération des informations saisies
        tempEFC = uf_choixFluxImposition.tb_param1.Value
        tempSFC = uf_choixFluxImposition.tb_param2.Value
        tempSFF = uf_choixFluxImposition.tb_param3.Value

        '## Vérification de la validité des informations saisies
        If tempEFC > TFCMax Or tempEFC < TFCMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempEFC")    ' "EFC"
            bornesValidesValeursSaisies = False
        End If

        If tempSFC > TFCMax Or tempSFC < TFCMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempSFC")    ' "SFC"
            bornesValidesValeursSaisies = False
        End If

        If tempSFF > TFFMax Or tempSFF < TFFMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempSFF")    ' "SFF"
            bornesValidesValeursSaisies = False
        End If

    Case 134
        '## Récupération des informations saisies
        tempEFC = uf_choixFluxImposition.tb_param1.Value
        tempEFF = uf_choixFluxImposition.tb_param2.Value
        tempSFF = uf_choixFluxImposition.tb_param3.Value

        '## Vérification de la validité des informations saisies
        If tempEFC > TFCMax Or tempEFC < TFCMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempEFC")    ' "EFC"
            bornesValidesValeursSaisies = False
        End If

        If tempEFF > TFFMax Or tempEFF < TFFMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempEFF")    ' "EFF"
            bornesValidesValeursSaisies = False
        End If

        If tempSFF > TFFMax Or tempSFF < TFFMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempSFF")    ' "SFF"
            bornesValidesValeursSaisies = False
        End If

    Case 234
        '## Récupération des informations saisies
        tempSFC = uf_choixFluxImposition.tb_param1.Value
        tempEFF = uf_choixFluxImposition.tb_param2.Value
        tempSFF = uf_choixFluxImposition.tb_param3.Value

        '## Vérification de la validité des informations saisies
        If tempSFC > TFCMax Or tempSFC < TFCMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempSFC")    ' "SFC"
            bornesValidesValeursSaisies = False
        End If

        If tempEFF > TFFMax Or tempEFF < TFFMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempEFF")    ' "EFF"
            bornesValidesValeursSaisies = False
        End If

        If tempSFF > TFFMax Or tempSFF < TFFMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempSFF")    ' "SFF"
            bornesValidesValeursSaisies = False
        End If

    Case 135
        '## Récupération des informations saisies
        tempEFC = uf_choixFluxImposition.tb_param1.Value
        tempEFF = uf_choixFluxImposition.tb_param2.Value
        puissance = uf_choixFluxImposition.tb_param3.Value

        '## Vérification de la validité des informations saisies
        If tempEFC > TFCMax Or tempEFC < TFCMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempEFC")    ' "EFC"
            bornesValidesValeursSaisies = False
        End If

        If tempEFF > TFFMax Or tempEFF < TFFMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempEFF")    ' "EFF"
            bornesValidesValeursSaisies = False
        End If

    Case 145
        '## Récupération des informations saisies
        tempEFC = uf_choixFluxImposition.tb_param1.Value
        tempSFF = uf_choixFluxImposition.tb_param2.Value
        puissance = uf_choixFluxImposition.tb_param3.Value

        '## Vérification de la validité des informations saisies
        If tempEFC > TFCMax Or tempEFC < TFCMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempEFC")    ' "EFC"
            bornesValidesValeursSaisies = False
        End If

        If tempSFF > TFFMax Or tempSFF < TFFMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempSFF")    ' "SFF"
            bornesValidesValeursSaisies = False
        End If

    Case 235
        '## Récupération des informations saisies
        tempSFC = uf_choixFluxImposition.tb_param1.Value
        tempEFF = uf_choixFluxImposition.tb_param2.Value
        puissance = uf_choixFluxImposition.tb_param3.Value

        '## Vérification de la validité des informations saisies
        If tempSFC > TFCMax Or tempSFC < TFCMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempSFC")    ' "SFC"
            bornesValidesValeursSaisies = False
        End If

        If tempEFF > TFFMax Or tempEFF < TFFMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempEFF")    ' "EFF"
            bornesValidesValeursSaisies = False
        End If

    Case 245
        '## Récupération des informations saisies
        tempSFC = uf_choixFluxImposition.tb_param1.Value
        tempSFF = uf_choixFluxImposition.tb_param2.Value
        puissance = uf_choixFluxImposition.tb_param3.Value

        '## Vérification de la validité des informations saisies
        If tempSFC > TFCMax Or tempSFC < TFCMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempSFC")    ' "SFC"
            bornesValidesValeursSaisies = False
        End If

        If tempSFF > TFFMax Or tempSFF < TFFMin Then
            erreursValeursBornesSaisies = erreursValeursBornesSaisies & Feuil5.Range("msgE_tempSFF")    ' "SFF"
            bornesValidesValeursSaisies = False
        End If

    End Select

    If bornesValidesValeursSaisies = False Then
        uf_choixFluxImposition.l_erreurEchangeur1 = Feuil5.Range("msgE_valeurSaisieHorsBornes") & " : " & erreursValeursBornesSaisies
        miseEnFormeRemarqueSchema (1)
    Else
        uf_choixFluxImposition.l_erreurEchangeur1 = ""
        miseEnFormeRemarqueSchema (3)
    End If

End Sub

Sub testerCroisementValeursSaisies()

    '### INITIALISATION DES VARIABLES
    erreursValeursCroisementSaisies = ""
    croisementValideValeursSaisies = True

    '### PROCEDURE
    Select Case combin

    Case 123
        '## Récupération des informations saisies
        tempEFC = uf_choixFluxImposition.tb_param1.Value
        tempSFC = uf_choixFluxImposition.tb_param2.Value

        '## Vérification de la validité des informations saisies
        If tempEFC = tempSFC Then
            erreursValeursCroisementSaisies = erreursValeursCroisementSaisies + Feuil5.Range("msgE_tempEFC.ega.tempSFC")    '"EFC=SFC"
            croisementValideValeursSaisies = False
        End If

        If tempEFC < tempSFC Then
            erreursValeursCroisementSaisies = erreursValeursCroisementSaisies + Feuil5.Range("msgE_tempEFC.inf.tempSFC")    ' "EFC<SFC"
            croisementValideValeursSaisies = False
        End If

    Case 124
        '## Récupération des informations saisies
        tempEFC = uf_choixFluxImposition.tb_param1.Value
        tempSFC = uf_choixFluxImposition.tb_param2.Value

        '## Vérification de la validité des informations saisies
        If tempEFC = tempSFC Then
            erreursValeursCroisementSaisies = erreursValeursCroisementSaisies + Feuil5.Range("msgE_tempEFC.ega.tempSFC")    ' "EFC=SFC"
            croisementValideValeursSaisies = False
        End If

        If tempEFC < tempSFC Then
            erreursValeursCroisementSaisies = erreursValeursCroisementSaisies + Feuil5.Range("msgE_tempEFC.inf.tempSFC")    ' "EFC<SFC"
            croisementValideValeursSaisies = False
        End If

    Case 134
        '## Récupération des informations saisies
        tempEFF = uf_choixFluxImposition.tb_param2.Value
        tempSFF = uf_choixFluxImposition.tb_param3.Value

        '## Vérification de la validité des informations saisies
        If tempEFF = tempSFF Then
            erreursValeursCroisementSaisies = erreursValeursCroisementSaisies + Feuil5.Range("msgE_tempEFF.ega.tempSFF")    ' "EFF=SFF"
            croisementValideValeursSaisies = False
        End If

        If tempEFF > tempSFF Then
            erreursValeursCroisementSaisies = erreursValeursCroisementSaisies + Feuil5.Range("msgE_tempEFF.sup.tempSFF")    ' "EFF>SFF"
            croisementValideValeursSaisies = False
        End If

    Case 234
        '## Récupération des informations saisies
        tempEFF = uf_choixFluxImposition.tb_param2.Value
        tempSFF = uf_choixFluxImposition.tb_param3.Value

        '## Vérification de la validité des informations saisies
        If tempEFF = tempSFF Then
            erreursValeursCroisementSaisies = erreursValeursCroisementSaisies + Feuil5.Range("msgE_tempEFF.ega.tempSFF")    ' "EFF=SFF"
            croisementValideValeursSaisies = False
        End If

        If tempEFF > tempSFF Then
            erreursValeursCroisementSaisies = erreursValeursCroisementSaisies + Feuil5.Range("msgE_tempEFF.sup.tempSFF")    ' "EFF>SFF"
            croisementValideValeursSaisies = False
        End If

    Case 135

    Case 145

    Case 235

    Case 245

    End Select

    If croisementValideValeursSaisies = False Then
        uf_choixFluxImposition.l_erreurEchangeur2 = Feuil5.Range("msgE_valeurSaisieCroisement") & " : " & erreursValeursCroisementSaisies
        miseEnFormeRemarqueSchema (2)
    Else
        uf_choixFluxImposition.l_erreurEchangeur2 = ""
        miseEnFormeRemarqueSchema (4)
    End If

End Sub

Sub testerBornesValeursCalculees()

    '### DECLARATION DES VARIABLES
    erreursValeursBornesCalculees = ""
    bornesValidesValeursCalculees = True

    '### PROCEDURE
    Select Case combin

    Case 123    ' tempSFF et Puissance sont calculés
        '## Vérification de la validité des informations calculées
        If tempSFF > TFFMax Or tempSFF < TFFMin Then
            erreursValeursBornesCalculees = erreursValeursBornesCalculees & Feuil5.Range("msgE_tempSFF")    ' "SFF"
            bornesValidesValeursCalculees = False
        End If

    Case 124    ' tempEFF et Puissance sont calculés
        '## Vérification de la validité des informations calculées
        If tempEFF > TFFMax Or tempEFF < TFFMin Then
            erreursValeursBornesCalculees = erreursValeursBornesCalculees & Feuil5.Range("msgE_tempEFF")    ' "EFF"
            bornesValidesValeursCalculees = False
        End If

    Case 134    ' tempSFC et Puissance sont calculés
        '## Vérification de la validité des informations calculées
        If tempSFC > TFCMax Or tempSFC < TFCMin Then
            erreursValeursBornesCalculees = erreursValeursBornesCalculees & Feuil5.Range("msgE_tempSFC")    ' "SFC"
            bornesValidesValeursCalculees = False
        End If

    Case 234    ' tempEFC et Puissance sont calculés
        '## Vérification de la validité des informations calculées
        If tempEFC > TFCMax Or tempEFC < TFCMin Then
            erreursValeursBornesCalculees = erreursValeursBornesCalculees & Feuil5.Range("msgE_tempEFC")    ' "EFC"
            bornesValidesValeursCalculees = False
        End If

    Case 135    ' tempSFC et tempSFF sont calculés
        '## Vérification de la validité des informations calculées
        If tempSFC > TFCMax Or tempSFC < TFCMin Then
            erreursValeursBornesCalculees = erreursValeursBornesCalculees & Feuil5.Range("msgE_tempSFC")    ' "SFC"
            bornesValidesValeursCalculees = False
        End If

        If tempSFF > TFFMax Or tempSFF < TFFMin Then
            erreursValeursBornesCalculees = erreursValeursBornesCalculees & Feuil5.Range("msgE_tempSFF")    ' "SFF"
            bornesValidesValeursCalculees = False
        End If

    Case 145    ' tempSFC et tempEFF sont calculés
        '## Vérification de la validité des informations calculées
        If tempSFC > TFCMax Or tempSFC < TFCMin Then
            erreursValeursBornesCalculees = erreursValeursBornesCalculees & Feuil5.Range("msgE_tempSFC")    ' "SFC"
            bornesValidesValeursCalculees = False
        End If

        If tempEFF > TFFMax Or tempEFF < TFFMin Then
            erreursValeursBornesCalculees = erreursValeursBornesCalculees & Feuil5.Range("msgE_tempEFF")    ' "EFF"
            bornesValidesValeursCalculees = False
        End If

    Case 235    ' tempEFC et tempSFF sont calculés
        '## Vérification de la validité des informations calculées
        If tempEFC > TFCMax Or tempEFC < TFCMin Then
            erreursValeursBornesCalculees = erreursValeursBornesCalculees & Feuil5.Range("msgE_tempEFC")    ' "EFC"
            bornesValidesValeursCalculees = False
        End If

        If tempSFF > TFFMax Or tempSFF < TFFMin Then
            erreursValeursBornesCalculees = erreursValeursBornesCalculees & Feuil5.Range("msgE_tempSFF")    ' "SFF"
            bornesValidesValeursCalculees = False
        End If

    Case 245    ' tempEFC et tempEFF sont calculés
        '## Vérification de la validité des informations calculées
        If tempEFC > TFCMax Or tempEFC < TFCMin Then
            erreursValeursBornesCalculees = erreursValeursBornesCalculees & Feuil5.Range("msgE_tempEFC")    ' "EFC"
            bornesValidesValeursCalculees = False
        End If

        If tempEFF > TFFMax Or tempEFF < TFFMin Then
            erreursValeursBornesCalculees = erreursValeursBornesCalculees & Feuil5.Range("msgE_tempEFF")    ' "EFF"
            bornesValidesValeursCalculees = False
        End If

    End Select

    If bornesValidesValeursCalculees = False Then
        uf_choixFluxImposition.l_erreurEchangeur1 = Feuil5.Range("msgE_valeurCalculeeHorsBornes") & " : " & erreursValeursBornesCalculees
        miseEnFormeRemarqueSchema (1)
    Else
        uf_choixFluxImposition.l_erreurEchangeur1 = ""
        miseEnFormeRemarqueSchema (3)
    End If

End Sub

Sub testerCroisementValeursCalculees()

    '### DECLARATION DES VARIABLES
    erreursValeursCroisementCalculees = ""
    croisementValideValeursCalculees = True

    '### PROCEDURE
    Select Case combin

    Case 123    ' tempSFF et Puissance sont calculés
        '## Vérification de la validité des informations calculées
        If tempEFF = tempSFF Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFF.ega.tempSFF")    ' "EFF=SFF"
            croisementValideValeursCalculees = False
        End If

        If tempEFF > tempSFF Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFF.sup.tempSFF")    ' "EFF>SFF"
            croisementValideValeursCalculees = False
        End If

    Case 124    ' tempEFF et Puissance sont calculés
        '## Vérification de la validité des informations calculées
        If tempEFF = tempSFF Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFF.ega.tempSFF")    ' "EFF=SFF"
            croisementValideValeursCalculees = False
        End If

        If tempEFF > tempSFF Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFF.sup.tempSFF")    ' "EFF>SFF"
            croisementValideValeursCalculees = False
        End If

    Case 134    ' tempSFC et Puissance sont calculés
        '## Vérification de la validité des informations calculées
        If tempEFC = tempSFC Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFC.ega.tempSFC")    ' "EFC=SFC"
            croisementValideValeursCalculees = False
        End If

        If tempEFC < tempSFC Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFC.inf.tempSFC")    ' "EFC<SFC"
            croisementValideValeursCalculees = False
        End If

    Case 234    ' tempEFC et Puissance sont calculés
        '## Vérification de la validité des informations calculées
        If tempEFC = tempSFC Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFC.ega.tempSFC")    ' "EFC=SFC"
            croisementValideValeursCalculees = False
        End If

        If tempEFC < tempSFC Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFC.inf.tempSFC")    ' "EFC<SFC"
            croisementValideValeursCalculees = False
        End If

    Case 135    ' tempSFC et tempSFF sont calculés
        '## Vérification de la validité des informations calculées
        If tempEFC = tempSFC Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFC.ega.tempSFC")    ' "EFC=SFC"
            croisementValideValeursCalculees = False
        End If

        If tempEFC < tempSFC Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFC.inf.tempSFC")    ' "EFC<SFC"
            croisementValideValeursCalculees = False
        End If

        If tempEFF = tempSFF Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFF.ega.tempSFF")    ' "EFF=SFF"
            croisementValideValeursCalculees = False
        End If

        If tempEFF > tempSFF Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFF.sup.tempSFF")    ' "EFF>SFF"
            croisementValideValeursCalculees = False
        End If

    Case 145    ' tempSFC et tempEFF sont calculés
        '## Vérification de la validité des informations calculées
        If tempEFC = tempSFC Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFC.ega.tempSFC")    ' "EFC=SFC"
            croisementValideValeursCalculees = False
        End If

        If tempEFC < tempSFC Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFC.inf.tempSFC")    ' "EFC<SFC"
            croisementValideValeursCalculees = False
        End If

        If tempEFF = tempSFF Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFF.ega.tempSFF")    ' "EFF=SFF"
            croisementValideValeursCalculees = False
        End If

        If tempEFF > tempSFF Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFF.sup.tempSFF")    ' "EFF>SFF"
            croisementValideValeursCalculees = False
        End If

    Case 235    ' tempEFC et tempSFF sont calculés
        '## Vérification de la validité des informations calculées
        If tempEFC = tempSFC Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFC.ega.tempSFC")    ' "EFC=SFC"
            croisementValideValeursCalculees = False
        End If

        If tempEFC < tempSFC Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFC.inf.tempSFC")    '"EFC<SFC"
            croisementValideValeursCalculees = False
        End If

        If tempEFF = tempSFF Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFF.ega.tempSFF")    '"EFF=SFF"
            croisementValideValeursCalculees = False
        End If

        If tempEFF > tempSFF Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFF.sup.tempSFF")    '"EFF>SFF"
            croisementValideValeursCalculees = False
        End If

    Case 245    ' tempEFC et tempEFF sont calculés
        '## Vérification de la validité des informations calculées
        If tempEFC = tempSFC Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFC.ega.tempSFC")    ' "EFC=SFC"
            croisementValideValeursCalculees = False
        End If

        If tempEFC < tempSFC Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFC.inf.tempSFC")    ' "EFC<SFC"
            croisementValideValeursCalculees = False
        End If

        If tempEFF = tempSFF Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFF.ega.tempSFF")    '"EFF=SFF"
            croisementValideValeursCalculees = False
        End If

        If tempEFF > tempSFF Then
            erreursValeursCroisementCalculees = erreursValeursCroisementCalculees + Feuil5.Range("msgE_tempEFF.sup.tempSFF")    '"EFF>SFF"
            croisementValideValeursCalculees = False
        End If

    End Select

    If croisementValideValeursCalculees = False Then
        uf_choixFluxImposition.l_erreurEchangeur2 = Feuil5.Range("msgE_valeurCalculeeCroisement") & " : " & erreursValeursCroisementCalculees
        miseEnFormeRemarqueSchema (2)
    Else
        uf_choixFluxImposition.l_erreurEchangeur2 = ""
        miseEnFormeRemarqueSchema (4)
    End If

End Sub

Sub miseEnFormeRemarqueSchema(action As Integer)

    Select Case action

    Case 1    'erreur A Signaler
        With uf_choixFluxImposition
            .l_erreurEchangeur1.Font.Bold = True
            .l_erreurEchangeur1.ForeColor = RGB(255, 255, 255)
            .l_erreurEchangeur1.BackStyle = fmBackStyleOpaque
            .l_erreurEchangeur1.BackColor = RGB(255, 0, 0)
        End With

    Case 2    'erreur A Signaler
        With uf_choixFluxImposition
            .l_erreurEchangeur2.Font.Bold = True
            .l_erreurEchangeur2.ForeColor = RGB(255, 255, 255)
            .l_erreurEchangeur2.BackStyle = fmBackStyleOpaque
            .l_erreurEchangeur2.BackColor = RGB(255, 0, 0)
        End With

    Case 3    'pas D Erreur
        With uf_choixFluxImposition
            .l_erreurEchangeur1.Font.Bold = False
            .l_erreurEchangeur1.ForeColor = RGB(0, 0, 0)
            .l_erreurEchangeur1.BackStyle = fmBackStyleTransparent
        End With

    Case 4    'pas D Erreur
        With uf_choixFluxImposition
            .l_erreurEchangeur2.Font.Bold = False
            .l_erreurEchangeur2.ForeColor = RGB(0, 0, 0)
            .l_erreurEchangeur2.BackStyle = fmBackStyleTransparent
        End With

    End Select

End Sub

Sub ajouterFluxInterdiction()

    '### DECLARATION DES VARIABLES
    Dim nbLigneMax As Integer
    Dim LigneTab As Integer
    Dim indiceTab As Integer
    
    '### INITIALISATION DES VARIABLES
    nbLigneMax = 0
    colonneStation = 1 + (numStationSelectionnee - 1) * 3
    LigneTab = INDICE_PREMIERE_LIGNE

    '### PROCEDURE
    '## Chargement des caractéristiques du flux utilité sélectionné
    fluxChaudInterdiction = uf_choixFluxInterdiction.cb_choixFluxChaud.Value
    fluxFroidInterdiction = uf_choixFluxInterdiction.cb_choixFluxFroid.Value

    '## Détermination du nombre de ligne
    While Feuil63.Range("DonneesPincementInterdictionEchange").Cells(LigneTab, colonneStation) <> ""
        LigneTab = LigneTab + 1
        nbLigneMax = nbLigneMax + 1
    Wend

    '## Création de la nouvelle ligne
    Feuil63.Range("DonneesPincementInterdictionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation) = "TRUE"
'    Feuil63.Range("DonneesPincementInterdictionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 1) = fluxChaudInterdiction    ' nom du flux chaud
    Call ecriture_nom_flux_interdiction_echange(fluxChaudInterdiction, Feuil63.Name, "DonneesPincementInterdictionEchange", INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 1, Feuil63.Name, "PINCH_NomsPlagesFluxProcede")
'    Feuil63.Range("DonneesPincementInterdictionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 2) = fluxFroidInterdiction     ' nom du flux froid
    Call ecriture_nom_flux_interdiction_echange(fluxFroidInterdiction, Feuil63.Name, "DonneesPincementInterdictionEchange", INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 2, Feuil63.Name, "PINCH_NomsPlagesFluxProcede")

End Sub

Sub chargerFluxInterdiction(indiceLVLigneAModifier As Integer)

    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * 3

    '### PROCEDURE
    '## Chargement des caractéristiques du flux utilité sélectionné
    fluxChaudInterdiction = Feuil63.Range("DonneesPincementInterdictionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 1).Value    ' nom du flux chaud
    fluxFroidInterdiction = Feuil63.Range("DonneesPincementInterdictionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 2).Value   ' nom du flux froid

    '## Affectation des données
    uf_choixFluxInterdiction.cb_choixFluxChaud = fluxChaudInterdiction  ' nom du flux chaud
    uf_choixFluxInterdiction.cb_choixFluxFroid = fluxFroidInterdiction  ' nom du flux froid

End Sub

Sub modifierFluxInterdiction(indiceLVLigneAModifier As Integer)

    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * 3

    '### PROCEDURE
    '## Chargement des caractéristiques du flux utilité sélectionné
    fluxChaudInterdiction = uf_choixFluxInterdiction.cb_choixFluxChaud.Value
    fluxFroidInterdiction = uf_choixFluxInterdiction.cb_choixFluxFroid.Value

    '## Modification de la ligne
'    Feuil63.Range("DonneesPincementInterdictionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 1) = fluxChaudInterdiction    ' nom du flux chaud
    Call ecriture_nom_flux_interdiction_echange(fluxChaudInterdiction, Feuil63.Name, "DonneesPincementInterdictionEchange", (INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 1, Feuil63.Name, "PINCH_NomsPlagesFluxProcede")
'    Feuil63.Range("DonneesPincementInterdictionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 2) = fluxFroidInterdiction     ' nom du flux froid
    Call ecriture_nom_flux_interdiction_echange(fluxFroidInterdiction, Feuil63.Name, "DonneesPincementInterdictionEchange", (INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 2, Feuil63.Name, "PINCH_NomsPlagesFluxProcede")
    
    modeModif_inter = False

End Sub

Sub supprimerFluxInterdiction(indiceLVLigneASupprimer As Integer)

    '### DECLARATION DES VARIABLES
    Dim nbLigneMax
    Dim LigneTab As Integer
    Dim indiceTab As Integer

    '### INITIALISATION DES VARIABLES
    nbLigneMax = 0
    colonneStation = 1 + (numStationSelectionnee - 1) * 3
    LigneTab = INDICE_PREMIERE_LIGNE

    '### PROCEDURE
    'Détermination du nombre de ligne
    While Feuil63.Range("DonneesPincementInterdictionEchange").Cells(LigneTab, colonneStation) <> ""
        LigneTab = LigneTab + 1
        nbLigneMax = nbLigneMax + 1
    Wend

    If nbLigneMax <> 0 Then

        'suppression de la ligne et réorganisation du tableau
        Dim etat As Boolean

        If indiceLVLigneASupprimer = nbLigneMax Then    ' la ligne a supprimer est la dernière de la liste
            Feuil63.Range("DonneesPincementInterdictionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation).ClearContents
            Feuil63.Range("DonneesPincementInterdictionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 1).ClearContents
            Feuil63.Range("DonneesPincementInterdictionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 2).ClearContents

        Else    ' la ligne à supprimer n'est pas la première de la liste
            For indiceTab = (INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer To 2 + nbLigneMax
                'stockage de la ligne à remonter
                etat = Feuil63.Range("DonneesPincementInterdictionEchange").Cells(indiceTab + 1, colonneStation)
                fluxChaudImpose = Feuil63.Range("DonneesPincementInterdictionEchange").Cells(indiceTab + 1, colonneStation + 1).Formula
                fluxFroidImpose = Feuil63.Range("DonneesPincementInterdictionEchange").Cells(indiceTab + 1, colonneStation + 2).Formula

                'réécriture de la ligne stockée sur la ligne du dessus
                Feuil63.Range("DonneesPincementInterdictionEchange").Cells(indiceTab, colonneStation) = etat
                Feuil63.Range("DonneesPincementInterdictionEchange").Cells(indiceTab, colonneStation + 1).Formula = fluxChaudImpose
                Feuil63.Range("DonneesPincementInterdictionEchange").Cells(indiceTab, colonneStation + 2).Formula = fluxFroidImpose

            Next
            'suppression de la dernière ligne
            Feuil63.Range("DonneesPincementInterdictionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation).ClearContents
            Feuil63.Range("DonneesPincementInterdictionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 1).ClearContents
            Feuil63.Range("DonneesPincementInterdictionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 2).ClearContents
        End If

    End If

End Sub

Sub ajouterFluxProcede()
    '#### Ajoute une ligne

    '### DECLARATION DES VARIABLES
    Dim nbLigneMax
    Dim LigneTab As Integer
    Dim indiceTab As Integer

    '### INITIALISATION DES VARIABLES
    nbLigneMax = 0
    colonneStation = 1 + (numStationSelectionnee - 1) * flux_thermique_process_nb_caracteristiques
    LigneTab = INDICE_PREMIERE_LIGNE

    '### PROCEDURE
    '## Détermination du nombre de lignes
    While Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation) <> ""
        LigneTab = LigneTab + 1
        nbLigneMax = nbLigneMax + 1
    Wend

    '## Création de la nouvelle ligne
    Feuil63.Range("DonneesPincementFluxProcedes").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation) = "TRUE"    ' état
    Feuil63.Range("DonneesPincementFluxProcedes").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 1) = uf_choixFluxProcede.tb_nomProcede.Value     ' nom
    Feuil63.Range("DonneesPincementFluxProcedes").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 2) = CDbl(uf_choixFluxProcede.tb_tempEntree.Value)    'température d'entrée
    Feuil63.Range("DonneesPincementFluxProcedes").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 3) = CDbl(uf_choixFluxProcede.tb_tempSortie.Value)   'température de sortie
    Feuil63.Range("DonneesPincementFluxProcedes").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 4) = CDbl(uf_choixFluxProcede.tb_MCp.Value)    'MCp
    Feuil63.Range("DonneesPincementFluxProcedes").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 5) = "FALSE"    'Original

End Sub

Sub chargerFluxProcede(indiceLVLigneAModifier As Integer)

    '### DECLARATION DES VARIABLES
    Dim nomProc As String
    Dim tempEntree As Double
    Dim tempSortie As Double
    Dim MCp As Double

    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * flux_thermique_process_nb_caracteristiques

    '### PROCEDURE
    '## Chargement des caractéristiques du flux utilité sélectionné
    nomProc = Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 1).Value    ' nom
    tempEntree = Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 2).Value   'température d'entrée
    tempSortie = Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 3).Value    'température de sortie
    MCp = Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 4).Value   'MCp

    '## Affectation des données
    uf_choixFluxProcede.tb_nomProcede = nomProc    'nom du procédé
    uf_choixFluxProcede.tb_tempEntree = tempEntree   'température d'entrée
    uf_choixFluxProcede.tb_tempSortie = tempSortie   'température de sortie
    uf_choixFluxProcede.tb_MCp = MCp   'MCp

End Sub

Sub modifierFluxProcede(indiceLVLigneAModifier As Integer)

    '### DECLARATION DES VARIABLES
    Dim LigneTab As Integer

    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * flux_thermique_process_nb_caracteristiques
    LigneTab = INDICE_PREMIERE_LIGNE

    '### PROCEDURE
    '## Modification de la ligne
    Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 1) = uf_choixFluxProcede.tb_nomProcede.Value   ' nom
    Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 2) = CDbl(uf_choixFluxProcede.tb_tempEntree.Value)    'température d'entrée
    Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 3) = CDbl(uf_choixFluxProcede.tb_tempSortie.Value)    'température de sortie
    Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 4) = CDbl(uf_choixFluxProcede.tb_MCp.Value)    'MCp

    modeModif_proc = False

End Sub

Sub supprimerFluxProcede(indiceLVLigneASupprimer As Integer)

    '### DECLARATION DES VARIABLES
    Dim nbLigneMax As Integer
    Dim LigneTab As Integer
    Dim indiceTab As Integer
    Dim etat As Boolean
    Dim nomProc As String
    Dim tempEntree As Double
    Dim tempSortie As Double
    Dim MCp As Double
    Dim original As Boolean

    '### INITIALISATION DES VARIABLES
    nbLigneMax = 0
    colonneStation = 1 + (numStationSelectionnee - 1) * flux_thermique_process_nb_caracteristiques
    LigneTab = INDICE_PREMIERE_LIGNE

    '### PROCEDURE
    '## Détermination du nombre de ligne
    While Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation) <> ""
        LigneTab = LigneTab + 1
        nbLigneMax = nbLigneMax + 1
    Wend

    If nbLigneMax <> 0 Then

        '## Suppression de la ligne et réorganisation du tableau

        If indiceLVLigneASupprimer = nbLigneMax Then    ' la ligne a supprimer est la dernière de la liste
            Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation).ClearContents    'état
            Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 1).ClearContents    ' nom
            Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 2).ClearContents    'température d'entrée
            Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 3).ClearContents    'température de sortie
            Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 4).ClearContents    ' MCp
            Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 5).ClearContents    'Original

        Else    ' la ligne à supprimer n'est pas la première de la liste
            For indiceTab = (INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer To 2 + nbLigneMax
                'stockage de la ligne à remonter
                etat = Feuil63.Range("DonneesPincementFluxProcedes").Cells(indiceTab + 1, colonneStation)    'état
                nomProc = Feuil63.Range("DonneesPincementFluxProcedes").Cells(indiceTab + 1, colonneStation + 1)    ' nom
                tempEntree = Feuil63.Range("DonneesPincementFluxProcedes").Cells(indiceTab + 1, colonneStation + 2)    'température d'entrée
                tempSortie = Feuil63.Range("DonneesPincementFluxProcedes").Cells(indiceTab + 1, colonneStation + 3)    'température de sortie
                MCp = Feuil63.Range("DonneesPincementFluxProcedes").Cells(indiceTab + 1, colonneStation + 4)    ' MCp
                original = Feuil63.Range("DonneesPincementFluxProcedes").Cells(indiceTab + 1, colonneStation + 5)    'Original

                'réécriture de la ligne stockée sur la ligne du dessus
                Feuil63.Range("DonneesPincementFluxProcedes").Cells(indiceTab, colonneStation) = etat    'état
                Feuil63.Range("DonneesPincementFluxProcedes").Cells(indiceTab, colonneStation + 1) = nomProc    ' nom
                Feuil63.Range("DonneesPincementFluxProcedes").Cells(indiceTab, colonneStation + 2) = tempEntree    'température d'entrée
                Feuil63.Range("DonneesPincementFluxProcedes").Cells(indiceTab, colonneStation + 3) = tempSortie    'température de sortie
                Feuil63.Range("DonneesPincementFluxProcedes").Cells(indiceTab, colonneStation + 4) = MCp    ' MCp
                Feuil63.Range("DonneesPincementFluxProcedes").Cells(indiceTab, colonneStation + 5) = original    'Original

            Next
            'suppression de la dernière ligne
            Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation).ClearContents    'état
            Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 1).ClearContents    ' nom
            Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 2).ClearContents    'température d'entrée
            Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 3).ClearContents  'température de sortie
            Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 4).ClearContents    ' MCp
            Feuil63.Range("DonneesPincementFluxProcedes").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 5).ClearContents    'Original

        End If
    End If
End Sub

Sub ajouterFluxUtilite()
    '#### Ajoute une ligne

    '### DECLARATION DES VARIABLES
    Dim nbLigneMax
    Dim LigneTab As Integer
    Dim indiceTab As Integer

    '### INITIALISATION DES VARIABLES
    nbLigneMax = 0
    colonneStation = 1 + (numStationSelectionnee - 1) * flux_thermique_utilites_nb_caracteristiques
    LigneTab = INDICE_PREMIERE_LIGNE

    '### PROCEDURE
    '## Détermination du nombre de lignes
    While Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation) <> ""
        LigneTab = LigneTab + 1
        nbLigneMax = nbLigneMax + 1
    Wend

    '## Création de la nouvelle ligne
    Feuil63.Range("DonneesPincementFluxUtilites").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation) = "TRUE"    ' état
    Feuil63.Range("DonneesPincementFluxUtilites").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 1) = uf_choixFluxUtilite.tb_nomUtilite    ' nom
    If uf_choixFluxUtilite.cmbo_typeFlux = Feuil5.Range("l_fluxChaud").Value Then   ' flux chaud
        Feuil63.Range("DonneesPincementFluxUtilites").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 2) = "Ch"
    ElseIf uf_choixFluxUtilite.cmbo_typeFlux = Feuil5.Range("l_fluxFroid").Value Then   ' flux froid
        Feuil63.Range("DonneesPincementFluxUtilites").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 2) = "Fr"
    End If
    Feuil63.Range("DonneesPincementFluxUtilites").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 3) = CDbl(uf_choixFluxUtilite.tb_tempEntree.Value)     'température d'entrée
    Feuil63.Range("DonneesPincementFluxUtilites").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 4) = CDbl(uf_choixFluxUtilite.tb_MCp.Value)     'MCp
    Feuil63.Range("DonneesPincementFluxUtilites").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 5) = CDbl(uf_choixFluxUtilite.tb_Ptot.Value)     'Ptot
    Feuil63.Range("DonneesPincementFluxUtilites").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 6) = "FALSE"    'Original

End Sub

Sub chargerFluxUtilite(indiceLVLigneAModifier As Integer)

    '### DECLARATION DES VARIABLES
    Dim nomUtil As String
    Dim typeFlux As String
    Dim tempEntree As Double
    Dim MCp As Double
    Dim Ptot As Double

    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * flux_thermique_utilites_nb_caracteristiques

    '### PROCEDURE
    '## Chargement des caractéristiques du flux utilité sélectionné
    nomUtil = Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 1).Value     ' nom
    typeFlux = Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 2).Value     ' Type de flux
    tempEntree = Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 3).Value     'température d'entrée
    MCp = Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 4).Value     'MCp
    Ptot = Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 5).Value     'Ptot

    '## Affectation des données
    uf_choixFluxUtilite.tb_nomUtilite = nomUtil    'nom de l'utilité
    If typeFlux = "Ch" Then    ' Type de flux
        uf_choixFluxUtilite.cmbo_typeFlux.Value = Feuil5.Range("l_fluxChaud").Value
    End If
    If typeFlux = "Fr" Then
        uf_choixFluxUtilite.cmbo_typeFlux.Value = Feuil5.Range("l_fluxFroid").Value
    End If
    uf_choixFluxUtilite.tb_tempEntree = tempEntree   'température d'entrée
    uf_choixFluxUtilite.tb_MCp = MCp   'MCp
    uf_choixFluxUtilite.tb_Ptot = Ptot   'Ptot

End Sub

Sub modifierFluxUtilite(indiceLVLigneAModifier As Integer)

    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * flux_thermique_utilites_nb_caracteristiques

    '### PROCEDURE
    '## Modification de la ligne
    Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 1) = uf_choixFluxUtilite.tb_nomUtilite   ' nom
    If uf_choixFluxUtilite.cmbo_typeFlux.Value = Feuil5.Range("l_fluxChaud").Value Then    ' Type de flux
        Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 2) = "Ch"
    End If
    If uf_choixFluxUtilite.cmbo_typeFlux.Value = Feuil5.Range("l_fluxFroid").Value Then
        Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 2) = "Fr"
    End If
    Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 3) = CDbl(uf_choixFluxUtilite.tb_tempEntree.Value)     'température d'entrée
    Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 4) = CDbl(uf_choixFluxUtilite.tb_MCp.Value)     'MCp
    Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 5) = CDbl(uf_choixFluxUtilite.tb_Ptot.Value)      'Ptot

    modeModif_util = False

End Sub

Sub supprimerFluxUtilite(indiceLVLigneASupprimer As Integer)

    '### DECLARATION DES VARIABLES
    Dim nbLigneMax As Integer
    Dim LigneTab As Integer
    Dim indiceTab As Integer
    Dim etat As Boolean
    Dim nomUtil As String
    Dim typeFlux As String
    Dim tempEntree As Double
    Dim MCp As Double
    Dim Ptot As Double
    Dim original As Boolean

    '### INITIALISATION DES VARIABLES
    nbLigneMax = 0
    colonneStation = 1 + (numStationSelectionnee - 1) * flux_thermique_utilites_nb_caracteristiques
    LigneTab = INDICE_PREMIERE_LIGNE

    '### PROCEDURE
    '## Détermination du nombre de ligne
    While Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation) <> ""
        LigneTab = LigneTab + 1
        nbLigneMax = nbLigneMax + 1
    Wend

    If nbLigneMax <> 0 Then

        '## Suppression de la ligne et réorganisation du tableau

        If indiceLVLigneASupprimer = nbLigneMax Then    ' la ligne a supprimer est la dernière de la liste
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation).ClearContents
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 1).ClearContents
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 2).ClearContents
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 3).ClearContents
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 4).ClearContents
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 5).ClearContents
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 6).ClearContents

        Else    ' la ligne à supprimer n'est pas la première de la liste
            For indiceTab = (INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer To 2 + nbLigneMax
                'stockage de la ligne à remonter
                etat = Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab + 1, colonneStation)
                nomUtil = Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab + 1, colonneStation + 1)
                typeFlux = Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab + 1, colonneStation + 2)
                tempEntree = Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab + 1, colonneStation + 3)
                MCp = Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab + 1, colonneStation + 4)
                Ptot = Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab + 1, colonneStation + 5)
                original = Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab + 1, colonneStation + 6)

                'réécriture de la ligne stockée sur la ligne du dessus
                Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab, colonneStation) = etat
                Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab, colonneStation + 1) = nomUtil
                Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab, colonneStation + 2) = typeFlux
                Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab, colonneStation + 3) = tempEntree
                Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab, colonneStation + 4) = MCp
                Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab, colonneStation + 5) = Ptot
                Feuil63.Range("DonneesPincementFluxUtilites").Cells(indiceTab, colonneStation + 6) = original

            Next
            'suppression de la dernière ligne
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation).ClearContents
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 1).ClearContents
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 2).ClearContents
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 3).ClearContents
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 4).ClearContents
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 5).ClearContents
            Feuil63.Range("DonneesPincementFluxUtilites").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 6).ClearContents

        End If

    End If

End Sub

Public Sub ajouterFluxImposition()
    '#### Ajoute une ligne

    '### DECLARATION DES VARIABLES
    Dim nbLigneMax
    Dim LigneTab As Integer
    Dim indiceTab As Integer

    '### INITIALISATION DES VARIABLES
    nbLigneMax = 0
    colonneStation = 1 + (numStationSelectionnee - 1) * 9
    LigneTab = INDICE_PREMIERE_LIGNE

    '### PROCEDURE
    '## Détermination du nombre de lignes
    While Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation) <> ""
        LigneTab = LigneTab + 1
        nbLigneMax = nbLigneMax + 1
    Wend

    '## Création de la nouvelle ligne
    ' état et noms des flux:
    Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation) = "TRUE"
'    Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 1) = fluxChaudImpose
    Call ecriture_nom_flux_interdiction_echange(fluxChaudImpose, Feuil63.Name, "DonneesPincementImpositionEchange", INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 1, Feuil63.Name, "PINCH_NomsPlagesFluxProcede")
'    Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 2) = fluxFroidImpose
    Call ecriture_nom_flux_interdiction_echange(fluxFroidImpose, Feuil63.Name, "DonneesPincementImpositionEchange", INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 2, Feuil63.Name, "PINCH_NomsPlagesFluxProcede")

    'paramètres saisis:
    Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + Mid(combin, 1, 1) + 2) = param1
    Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + Mid(combin, 2, 1) + 2) = param2
    Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + Mid(combin, 3, 1) + 2) = param3

    'paramètres calculés:
    Select Case combin
    Case 123
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 6) = tempSFF
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 7) = puissance
    Case 124
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 5) = tempEFF
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 7) = puissance
    Case 134
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 4) = tempSFC
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 7) = puissance
    Case 234
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 3) = tempEFC
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 7) = puissance
    Case 135
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 4) = tempSFC
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 6) = tempSFF
    Case 145
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 4) = tempSFC
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 5) = tempEFF
    Case 235
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 3) = tempEFC
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 6) = tempSFF
    Case 245
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 3) = tempEFC
        Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 5) = tempEFF
    End Select

    'combinaison:
    Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + nbLigneMax, colonneStation + 8) = combin

End Sub

Public Sub chargerFluxImposition(indiceLVLigneAModifier As Integer)

    '### DECLARATION DES VARIABLES
    Dim combinaison As String
    Dim nomFluxChaud As String
    Dim nomFluxFroid As String
    Dim combinaisons As Variant

    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * 9
    combinaison = Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 8)

    '### PROCEDURE
    With uf_choixFluxImposition
        'Chargement des noms des flux

        nomFluxChaud = Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 1)
        nomFluxFroid = Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 2)

        .cb_choixFluxChaud.Value = nomFluxChaud
        .cb_choixFluxFroid.Value = nomFluxFroid

        'Chargement des données paramètres
        colonneStation = 1 + (numStationSelectionnee - 1) * 9
        .tb_param1 = Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + Mid(combinaison, 1, 1) + 2)
        .tb_param2 = Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + Mid(combinaison, 2, 1) + 2)
        .tb_param3 = Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + Mid(combinaison, 3, 1) + 2)

    End With

    Dim indice, indiceCible As Integer
    'Chargement des combinaisons de paramètres (1:TCin; 2:TCout; 3:TFin; 4:TFout; 5:P)
    combinaisons = Array(123, 124, 134, 234, 135, 145, 235, 245)

    'Recherche de la ligne correspondant à la combinaison
    For indice = 0 To 7
        If combinaisons(indice) = combinaison Then
            indiceCible = indice
        End If
    Next indice

    'Sélection de la ligne correspondant à la combinaison
    uf_choixFluxImposition.lb_choixCombinaisonParam.ListIndex = indiceCible

End Sub

Public Sub modifierFluxImposition(indiceLVLigneAModifier As Integer)

    '### INITIALISATION DES VARIABLES
    colonneStation = 1 + (numStationSelectionnee - 1) * 9

    '### PROCEDURE
    'Modification de la ligne
    Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation) = "TRUE"
'    Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 1) = fluxChaudImpose
    Call ecriture_nom_flux_interdiction_echange(fluxChaudImpose, Feuil63.Name, "DonneesPincementImpositionEchange", (INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 1, Feuil63.Name, "PINCH_NomsPlagesFluxProcede")
'    Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 2) = fluxFroidImpose
    Call ecriture_nom_flux_interdiction_echange(fluxFroidImpose, Feuil63.Name, "DonneesPincementImpositionEchange", (INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 2, Feuil63.Name, "PINCH_NomsPlagesFluxProcede")

    Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + Mid(combin, 1, 1) + 2) = param1
    Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + Mid(combin, 2, 1) + 2) = param2
    Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + Mid(combin, 3, 1) + 2) = param3

    Select Case combin
    Case 123
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 6) = tempSFF
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 7) = puissance
    Case 124
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 5) = tempEFF
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 7) = puissance
    Case 134
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 4) = tempSFC
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 7) = puissance
    Case 234
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 3) = tempEFC
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 7) = puissance
    Case 135
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 4) = tempSFC
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 6) = tempSFF
    Case 145
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 4) = tempSFC
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 5) = tempEFF
    Case 235
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 3) = tempEFC
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 6) = tempSFF
    Case 245
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 3) = tempEFC
        Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 5) = tempEFF
    End Select

    Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneAModifier, colonneStation + 8) = combin

    modeModif_impos = False

End Sub

Public Sub supprimerFluxImposition(indiceLVLigneASupprimer As Integer)

    '### DECLARATION DES VARIABLES
    Dim nbLigneMax As Integer
    Dim LigneTab As Integer
    Dim indiceTab As Integer

    '### INITIALISATION DES VARIABLES
    nbLigneMax = 0
    colonneStation = 1 + (numStationSelectionnee - 1) * 9
    LigneTab = INDICE_PREMIERE_LIGNE

    '### PROCEDURE
    'Détermination du nombre de ligne
    While Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation) <> ""
        LigneTab = LigneTab + 1
        nbLigneMax = nbLigneMax + 1
    Wend

    If nbLigneMax <> 0 Then

        'suppression de la ligne et réorganisation du tableau
        Dim etat As Boolean
        Dim nomFluxChaud, nomFluxFroid As String
        Dim TCin, TCout, TFin, TFout, Puiss, combinaison As Double

        If indiceLVLigneASupprimer = nbLigneMax Then    ' la ligne a supprimer est la dernière de la liste
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 1).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 2).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 3).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 4).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 5).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 6).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 7).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer, colonneStation + 8).ClearContents

        Else    ' la ligne à supprimer n'est pas la première de la liste
            For indiceTab = (INDICE_PREMIERE_LIGNE - 1) + indiceLVLigneASupprimer To 2 + nbLigneMax
                'stockage de la ligne à remonter
                etat = Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab + 1, colonneStation)
                nomFluxChaud = Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab + 1, colonneStation + 1).Formula
                nomFluxFroid = Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab + 1, colonneStation + 2).Formula
                TCin = Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab + 1, colonneStation + 3)
                TCout = Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab + 1, colonneStation + 4)
                TFin = Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab + 1, colonneStation + 5)
                TFout = Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab + 1, colonneStation + 6)
                Puiss = Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab + 1, colonneStation + 7)
                combinaison = Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab + 1, colonneStation + 8)

                'réécriture de la ligne stockée sur la ligne du dessus
                Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab, colonneStation) = etat
                Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab, colonneStation + 1).Formula = nomFluxChaud
                Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab, colonneStation + 2).Formula = nomFluxFroid
                Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab, colonneStation + 3) = TCin
                Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab, colonneStation + 4) = TCout
                Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab, colonneStation + 5) = TFin
                Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab, colonneStation + 6) = TFout
                Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab, colonneStation + 7) = Puiss
                Feuil63.Range("DonneesPincementImpositionEchange").Cells(indiceTab, colonneStation + 8) = combinaison

            Next
            'suppression de la dernière ligne
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 1).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 2).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 3).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 4).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 5).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 6).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 7).ClearContents
            Feuil63.Range("DonneesPincementImpositionEchange").Cells((INDICE_PREMIERE_LIGNE - 1) + nbLigneMax, colonneStation + 8).ClearContents

        End If
    End If

End Sub

Sub alimentationTBValeurPincement()

    Dim LigneTab As Integer

    colonneStation = 1 + (numStationSelectionnee - 1)
    LigneTab = INDICE_PREMIERE_LIGNE

    uf_choixPincement.tb_valeurPincement.Value = Feuil63.Range("DonneesPincementPinch").Cells(LigneTab, colonneStation).Value

End Sub

Sub modifierTBValeurPincement()

    Dim LigneTab As Integer

    colonneStation = 1 + (numStationSelectionnee - 1)
    LigneTab = INDICE_PREMIERE_LIGNE

    Feuil63.Range("DonneesPincementPinch").Cells(LigneTab, colonneStation).Value = CDbl(uf_choixPincement.tb_valeurPincement.Value)

End Sub

Sub alimentationTBSeuilPuissance()

    Dim LigneTab As Integer

    colonneStation = 1 + (numStationSelectionnee - 1)
    LigneTab = INDICE_PREMIERE_LIGNE

    uf_choixPincement.tb_seuilPuissance.Value = Feuil63.Range("DonneesPincementSeuilPuissance").Cells(LigneTab, colonneStation).Value

End Sub

Sub modifierTBSeuilPuissance()

    Dim LigneTab As Integer

    colonneStation = 1 + (numStationSelectionnee - 1)
    LigneTab = INDICE_PREMIERE_LIGNE

    Feuil63.Range("DonneesPincementSeuilPuissance").Cells(LigneTab, colonneStation).Value = CDbl(uf_choixPincement.tb_seuilPuissance.Value)

End Sub

Sub verificationFlux()
    '--> vérifie la validité de certains paramètres
    '--> affecte la valeur du boléen echangeursValides a true si les données sont valides, false sinon

    '### DECLARATION DES VARIABLES

    Dim nombreEchangeurs As Integer
    Dim numEchangeur As Integer
    Dim numEchangeurAVerifier As Integer

    Dim numStep As Integer
    Dim nbflux As Integer
    Dim flux As Integer
    Dim LigneTab As Integer
    Dim nbLigneErreur As Integer

    Dim etat As Boolean
    Dim nomFluxChaudOrigin As String
    Dim nomFluxFroidOrigin As String
    Dim TinFChaudOrigin As Double
    Dim ToutFChaudOrigin As Double
    Dim TinFFroidOrigin As Double
    Dim ToutFFroidOrigin As Double

    Dim nomFluxChaud As String
    Dim nomFluxFroid As String
    Dim TinFChaud As Double
    Dim ToutFChaud As Double
    Dim TinFFroid As Double
    Dim ToutFFroid As Double

    '### INITIALISATION DES VARIABLES
    Erase erreurTemperature
    nbLigneErreur = 1
    echangeursValides = True    ' initialisation du booleen
    donneesFluxProcValides = True
    tabErreurFluxProc(1) = False
    tabErreurFluxProc(2) = False
    tabErreurFluxProc(3) = False
    erreurFluxProc = ""
    donneesPincementSeuilValides = True
    erreurDonneesPincementSeuil = ""

    '### PROCEDURE


    '## Vérification de la validité des valeurs Pincement et Seuil

    ' Pincement
    If uf_choixPincement.tb_valeurPincement = "" Or IsNumeric(uf_choixPincement.tb_valeurPincement) = False Then
        donneesPincementSeuilValides = False
        erreurDonneesPincementSeuil = erreurDonneesPincementSeuil & Feuil5.Range("l_valeurPincement").Value & vbCrLf
        '    ElseIf uf_choixPincement.tb_valeurPincement = 0 Then
        '        MsgBox Feuil5.Range("l_valeurPincement").Value, vbCritical, Feuil5.Range("msgE_erreurSaisie").Value
        '        Exit Sub
    End If

    ' Seuil
    If uf_choixPincement.tb_seuilPuissance = "" Or IsNumeric(uf_choixPincement.tb_seuilPuissance) = False Then
        donneesPincementSeuilValides = False
        erreurDonneesPincementSeuil = erreurDonneesPincementSeuil & Feuil5.Range("l_seuilPuissance").Value & vbCrLf
        '    ElseIf uf_choixPincement.tb_seuilPuissance = 0 Then
        '        MsgBox Feuil5.Range("l_seuilPuissance").Value, vbCritical, Feuil5.Range("msgE_erreurSaisie").Value
        '        Exit Sub
    End If



    '## Vérification de la présence de flux procédés

    'MODIF 15/10/12 MSM
    '    'Pour chaque pi_FiliereConsideree
    '    For numStep = 1 To nb_step_comparaison
    numStep = numStationSelectionnee
    'FIN MODIF 15/10/12 MSM

    nbflux = 0

    'initialisation des coordonnées
    LigneTab = INDICE_PREMIERE_LIGNE
    colonneStation = 1 + (numStep - 1) * flux_thermique_process_nb_caracteristiques

    'détermination du nombre de flux
    Do While Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation) <> ""
        nbflux = nbflux + 1
        LigneTab = LigneTab + 1
    Loop

    'interprétation
    If nbflux = 0 Then
        tabErreurFluxProc(numStep) = True    ' il y a une erreur avec la filière numStep : pas de flux
    Else
        tabErreurFluxProc(numStep) = True    'on suppose qu'il n'y a pas de flux sélectionnés
        'on vérifie sil y a au moins 1 flux sélectionné
        For flux = INDICE_PREMIERE_LIGNE To INDICE_PREMIERE_LIGNE + nbflux
            If Feuil63.Range("DonneesPincementFluxProcedes").Cells(flux, colonneStation) = True Then
                tabErreurFluxProc(numStep) = False
                Exit For
            End If
        Next flux
    End If

    If tabErreurFluxProc(numStep) = True Then
        donneesFluxProcValides = False    'on affecte le booleen s'il y a au moins 1 erreur
        erreurFluxProc = erreurFluxProc & Feuil5.Range("msgE_STEP").Value & " " & numStep & " "
    End If

    'MODIF 15/10/12 MSM
    '   Next numStep
    'FIN MODIF 15/10/12 MSM



    '## Vérification du chevauchement des températures pour les échangeurs contenant les mêmes flux

    'MODIF 15/10/12 MSM
    '    'Pour chaque pi_FiliereConsideree
    '    For numStep = 1 To nb_step_comparaison
    numStep = numStationSelectionnee
    'FIN MODIF 15/10/12 MSM
    colonneStation = 1 + (numStep - 1) * 9

    ' Determination du nombre d'échangeurs
    LigneTab = INDICE_PREMIERE_LIGNE
    nombreEchangeurs = 0
    While Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation) <> ""
        nombreEchangeurs = nombreEchangeurs + 1
        LigneTab = LigneTab + 1
    Wend

    'Si des échangeurs personnalisés sont créés
    If nombreEchangeurs <> 0 Then

        'Pour chaque échangeur coché
        For numEchangeur = 1 To nombreEchangeurs

            etat = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeur - 1, colonneStation)
            If etat = True Then

                'Chargement des paramètres
                nomFluxChaudOrigin = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeur - 1, colonneStation + 1)
                nomFluxFroidOrigin = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeur - 1, colonneStation + 2)
                TinFChaudOrigin = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeur - 1, colonneStation + 3)
                ToutFChaudOrigin = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeur - 1, colonneStation + 4)
                TinFFroidOrigin = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeur - 1, colonneStation + 5)
                ToutFFroidOrigin = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeur - 1, colonneStation + 6)

                'Vérification du chevauchement des températures avec les autres échangeurs cochés
                For numEchangeurAVerifier = numEchangeur + 1 To nombreEchangeurs

                    etat = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeurAVerifier - 1, colonneStation)
                    If etat = True Then

                        'Chargement des paramètres
                        nomFluxChaud = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeurAVerifier - 1, colonneStation + 1)
                        nomFluxFroid = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeurAVerifier - 1, colonneStation + 2)
                        TinFChaud = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeurAVerifier - 1, colonneStation + 3)
                        ToutFChaud = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeurAVerifier - 1, colonneStation + 4)
                        TinFFroid = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeurAVerifier - 1, colonneStation + 5)
                        ToutFFroid = Feuil63.Range("DonneesPincementImpositionEchange").Cells(INDICE_PREMIERE_LIGNE + numEchangeurAVerifier - 1, colonneStation + 6)

                        'Flux chaud
                        If nomFluxChaud = nomFluxChaudOrigin Then
                            'vérif temp
                            If TinFChaud >= ToutFChaudOrigin And TinFChaud <= TinFChaudOrigin Then

                                ReDim Preserve erreurTemperature(1 To 5, 1 To nbLigneErreur)
                                erreurTemperature(1, nbLigneErreur) = CStr(numStep)
                                erreurTemperature(2, nbLigneErreur) = CStr(numEchangeur)
                                erreurTemperature(3, nbLigneErreur) = CStr(numEchangeurAVerifier)
                                erreurTemperature(4, nbLigneErreur) = nomFluxChaud
                                erreurTemperature(5, nbLigneErreur) = Feuil5.Range("msgE_TinFC")
                                nbLigneErreur = nbLigneErreur + 1

                                'affectation booleen
                                echangeursValides = False
                            End If

                            'vérif temp
                            If ToutFChaud >= ToutFChaudOrigin And ToutFChaud <= TinFChaudOrigin Then

                                ReDim Preserve erreurTemperature(1 To 5, 1 To nbLigneErreur)
                                erreurTemperature(1, nbLigneErreur) = CStr(numStep)
                                erreurTemperature(2, nbLigneErreur) = CStr(numEchangeur)
                                erreurTemperature(3, nbLigneErreur) = CStr(numEchangeurAVerifier)
                                erreurTemperature(4, nbLigneErreur) = nomFluxChaud
                                erreurTemperature(5, nbLigneErreur) = Feuil5.Range("msgE_ToutFC")
                                nbLigneErreur = nbLigneErreur + 1

                                'affectation booleen
                                echangeursValides = False
                            End If
                        End If

                        'Flux froid
                        If nomFluxFroid = nomFluxFroidOrigin Then
                            'vérif temp
                            If TinFFroid >= TinFFroidOrigin And TinFFroid <= ToutFFroidOrigin Then

                                ReDim Preserve erreurTemperature(1 To 5, 1 To nbLigneErreur)
                                erreurTemperature(1, nbLigneErreur) = CStr(numStep)
                                erreurTemperature(2, nbLigneErreur) = CStr(numEchangeur)
                                erreurTemperature(3, nbLigneErreur) = CStr(numEchangeurAVerifier)
                                erreurTemperature(4, nbLigneErreur) = nomFluxFroid
                                erreurTemperature(5, nbLigneErreur) = Feuil5.Range("msgE_TinFF")
                                nbLigneErreur = nbLigneErreur + 1

                                'affectation booleen
                                echangeursValides = False
                            End If

                            'vérif temp
                            If ToutFFroid >= TinFFroidOrigin And ToutFFroid <= ToutFFroidOrigin Then

                                ReDim Preserve erreurTemperature(1 To 5, 1 To nbLigneErreur)
                                erreurTemperature(1, nbLigneErreur) = CStr(numStep)
                                erreurTemperature(2, nbLigneErreur) = CStr(numEchangeur)
                                erreurTemperature(3, nbLigneErreur) = CStr(numEchangeurAVerifier)
                                erreurTemperature(4, nbLigneErreur) = nomFluxFroid
                                erreurTemperature(5, nbLigneErreur) = Feuil5.Range("msgE_ToutFF")
                                nbLigneErreur = nbLigneErreur + 1

                                'affectation booleen
                                echangeursValides = False
                            End If
                        End If

                    End If
                Next numEchangeurAVerifier

            End If
        Next numEchangeur

    End If
    'MODIF 15/10/12 MSM
    '   Next numStep
    'FIN MODIF 15/10/12 MSM

End Sub

Public Sub lancer_calcul()


    Dim biogaz_caract(nb_biogaz_caracteristiques) As Double
    Dim energie_elec As Double
    Dim type_valorisation As Integer
    Dim valorisation_chaudiere As Integer
    Dim ratio_biogaz_valorise As Double


    Application.Calculation = xlCalculationManual

    pn_EffacementResultatsEnCours = False

    PINCH_DANS_OCEAN = True
    If PINCH_DANS_OCEAN = True Then
        pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE
        Call lecture_flux_process
        Call lecture_flux_utilites
        Dim gestion_energie As New Gestion_energie_thermique_PINCH
        Call gestion_energie.lecture_donnees_modification_pinch(biogaz_caract, energie_elec, type_valorisation, valorisation_chaudiere, ratio_biogaz_valorise)
        Call gestion_energie.fonctionnement_reel(biogaz_caract, energie_elec, type_valorisation, valorisation_chaudiere, ratio_biogaz_valorise)
        Set gestion_energie = Nothing
    Else
        Dim pincement As New PINCH
        Call pincement.VeriPinch(1)    'Calculs
        Call pincement.VeriPinch(0)    ' HEN
        Set pincement = Nothing
    End If


    'End If

    Application.Calculation = xlCalculationAutomatic

End Sub

Sub razFiliere(Filiere As Integer)
    '--> Réinitialise la filière sélectionnée

    Dim LigneTab As Integer
    Dim Colonne As Integer

    'Suppression des flux procédés ajoutés
    numStationSelectionnee = Filiere
    colonneStation = 1 + (Filiere - 1) * flux_thermique_process_nb_caracteristiques
    LigneTab = INDICE_PREMIERE_LIGNE
    While Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation) <> ""
        If Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation + 5).Value = False Then
            supprimerFluxProcede (LigneTab - (INDICE_PREMIERE_LIGNE - 1))
        End If
        LigneTab = LigneTab + 1
    Wend

    'RAZ des états des cases à cocher  pour les flux procédés (VRAI)
    LigneTab = INDICE_PREMIERE_LIGNE
    While Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation) <> ""
        If Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation).Value = False Then
            Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation).Value = True
        End If
        LigneTab = LigneTab + 1
    Wend

    'Suppression des flux utilités ajoutés
    colonneStation = 1 + (Filiere - 1) * flux_thermique_utilites_nb_caracteristiques
    LigneTab = INDICE_PREMIERE_LIGNE
    While Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation) <> ""
        If Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation + 6).Value = False Then
            supprimerFluxUtilite (LigneTab - (INDICE_PREMIERE_LIGNE - 1))
        End If
        LigneTab = LigneTab + 1
    Wend

    'RAZ des états des cases à cocher  pour les utilités (VRAI)
    LigneTab = INDICE_PREMIERE_LIGNE
    While Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation) <> ""
        If Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation).Value = False Then
            Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation).Value = True
        End If
        LigneTab = LigneTab + 1
    Wend

    'RAZ des échanges interdits
    colonneStation = 1 + (Filiere - 1) * 3
    LigneTab = INDICE_PREMIERE_LIGNE
    While Feuil63.Range("DonneesPincementInterdictionEchange").Cells(LigneTab, colonneStation) <> ""
        For Colonne = 0 To 2
            Feuil63.Range("DonneesPincementInterdictionEchange").Cells(LigneTab, colonneStation + Colonne).ClearContents
        Next Colonne
        LigneTab = LigneTab + 1
    Wend

    'RAZ des échanges imposés
    colonneStation = 1 + (Filiere - 1) * 9
    LigneTab = INDICE_PREMIERE_LIGNE
    While Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation) <> ""
        For Colonne = 0 To 8
            Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation + Colonne).ClearContents
        Next Colonne
        LigneTab = LigneTab + 1
    Wend

    'RAZ de la valeur du pincement
    colonneStation = Filiere
    LigneTab = INDICE_PREMIERE_LIGNE
    Feuil63.Range("DonneesPincementPinch").Cells(LigneTab, colonneStation).Value = 10    ' valeur par défaut du pincement = 10°C

    'RAZ de l'état du tri
    colonneStation = Filiere
    LigneTab = INDICE_PREMIERE_LIGNE
    Feuil63.Range("DonneesPincementTriAutoProcedes").Cells(LigneTab, colonneStation).Value = True
    Feuil63.Range("DonneesPincementTriAutoUtilites").Cells(LigneTab, colonneStation).Value = True

    If Filiere = 1 Then

        pi_FiliereConsideree = Filiere
        PINCH_ORDONNER_FLUX_PROCESS = Feuil63.Range("DonneesPincementTriAutoProcedes").Cells(INDICE_PREMIERE_LIGNE, pi_FiliereConsideree).Value
        PINCH_ORDONNER_FLUX_UTILITES = Feuil63.Range("DonneesPincementTriAutoUtilites").Cells(INDICE_PREMIERE_LIGNE, pi_FiliereConsideree).Value
        lancer_calcul

    ElseIf Filiere = 2 Then

        pi_FiliereConsideree = Filiere
        PINCH_ORDONNER_FLUX_PROCESS = Feuil63.Range("DonneesPincementTriAutoProcedes").Cells(INDICE_PREMIERE_LIGNE, pi_FiliereConsideree).Value
        PINCH_ORDONNER_FLUX_UTILITES = Feuil63.Range("DonneesPincementTriAutoUtilites").Cells(INDICE_PREMIERE_LIGNE, pi_FiliereConsideree).Value
        lancer_calcul

    ElseIf Filiere = 3 Then

        pi_FiliereConsideree = Filiere
        PINCH_ORDONNER_FLUX_PROCESS = Feuil63.Range("DonneesPincementTriAutoProcedes").Cells(INDICE_PREMIERE_LIGNE, pi_FiliereConsideree).Value
        PINCH_ORDONNER_FLUX_UTILITES = Feuil63.Range("DonneesPincementTriAutoUtilites").Cells(INDICE_PREMIERE_LIGNE, pi_FiliereConsideree).Value
        lancer_calcul

    End If

    'mise à jour des LV de l'UF
    miseAJourLV

End Sub


