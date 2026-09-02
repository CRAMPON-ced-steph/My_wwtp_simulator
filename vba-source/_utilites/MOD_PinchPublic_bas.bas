Attribute VB_Name = "MOD_PinchPublic"
Option Explicit

'Constante pour gérer les différentes interfaces du PINCH (indépendant ou dans OCEAN)
Public PINCH_DANS_OCEAN As Boolean   'TRUE si dans OCEAN, FALSE si indépendant
Public PINCH_ORDONNER_FLUX_PROCESS As Boolean     'ordonne les flux process avant le calcul
Public PINCH_ORDONNER_FLUX_UTILITES As Boolean     'ordonne les flux utilités avant le calcul
Public Const REPORT_FLUX_ORDONNES = True     'réécrit les flux ordonnés sur Excel à la place du tableau initial

'Choix analyse pincement thermique
Public choix_analyse_pinch As Boolean
Public flux_thermique_process() As Variant
Public Const flux_thermique_process_nb_caracteristiques = 6
Public Const repere_flux_thermique_process_etat = 1
Public Const repere_flux_thermique_process_nom = 2
Public Const repere_flux_thermique_process_T_in_°C = 3
Public Const repere_flux_thermique_process_T_out_°C = 4
Public Const repere_flux_thermique_process_MCp_kW_°C = 5
Public Const repere_flux_thermique_process_original = 6
Public utilites_flux_energie() As Variant
Public nb_flux_utilites As Integer
Public Const flux_thermique_utilites_nb_caracteristiques = 7
Public Const repere_flux_thermique_utilites_etat = 1
Public Const repere_flux_thermique_utilites_nom = 2
Public Const repere_flux_thermique_utilites_type = 3
Public Const repere_flux_thermique_utilites_T_in_°C = 4
Public Const repere_flux_thermique_utilites_MCp_kW_°C = 5
Public Const repere_flux_thermique_utilites_Pw_kW = 6
Public Const repere_flux_thermique_utilites_original = 7

Public Sub lecture_flux_process()

    Const LigneTab = 3

    Dim nb_flux As Integer
    Dim colonneStation As Integer

    Dim I As Integer
    Dim j As Integer

    If pi_FiliereConsideree > 0 And pi_FiliereConsideree <= 3 Then
        'lecture des flux de procédés
        colonneStation = (pi_FiliereConsideree - 1) * flux_thermique_process_nb_caracteristiques
        I = 1
        ReDim flux_thermique_process(1 To flux_thermique_process_nb_caracteristiques, 1 To I)
        Do While Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab + I, colonneStation + repere_flux_thermique_process_etat) <> ""
            ReDim Preserve flux_thermique_process(1 To flux_thermique_process_nb_caracteristiques, 1 To I)
            For j = repere_flux_thermique_process_etat To flux_thermique_process_nb_caracteristiques
                flux_thermique_process(j, I) = Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab + I, colonneStation + j)
            Next j
            I = I + 1
        Loop
    Else
        MsgBox ("Sub lecture_flux_process: ATTENTION à définir le nombre de filières")
    End If

End Sub

Public Sub lecture_flux_utilites()

    Const LigneTab = 3

    Dim nb_flux As Integer
    Dim colonneStation As Integer

    Dim I As Integer
    Dim j As Integer

    If pi_FiliereConsideree > 0 And pi_FiliereConsideree <= 3 Then
        'lecture des flux de procédés
        colonneStation = (pi_FiliereConsideree - 1) * flux_thermique_utilites_nb_caracteristiques
        I = 1
        ReDim utilites_flux_energie(1 To flux_thermique_utilites_nb_caracteristiques, 1 To I)
        Do While Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab + I, colonneStation + repere_flux_thermique_utilites_etat) <> ""
            ReDim Preserve utilites_flux_energie(1 To flux_thermique_utilites_nb_caracteristiques, 1 To I)
            For j = repere_flux_thermique_utilites_etat To flux_thermique_utilites_nb_caracteristiques
                utilites_flux_energie(j, I) = Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab + I, colonneStation + j)
            Next j
            I = I + 1
        Loop
        nb_flux_utilites = I - 1
    Else
        MsgBox ("Sub lecture_flux_process: ATTENTION à définir le nombre de filières")
    End If


End Sub

'Mod CRE - voir Sub test()
'MODIF CRE 05/09/2012   Pas besoin de ça ici. A supprimer après vérification que c'est vraiment jamais utilisé
'Public Sub ExportHEN (pi_FiliereConsideree As Integer, t() As String, size As Integer)
'    'Macro d'export du tableau de flux avec réseau d'échangeurs en image (enregistrée dans mes Documents par défaut)
'
'    Dim i As Integer
'    Dim j As Integer
'    Dim k As Integer
'    Dim m As Integer
'
'    Dim x As Integer
'    Dim y As Integer
'    Dim Lligne As Integer
'    Dim Hligne As Integer
'    Dim sepligne As Integer
'    Dim LargCel As Integer
'    Dim posX As Integer
'    Dim posX2 As Integer
'    Dim Ltotal As Single
'    Dim Htotal As Single
'    Dim hactuel As Single
'    Dim lactuel As Single
'
'
'    'Initialisation des variables globales
'    Sheets("Tampon" & pi_FiliereConsideree).Activate
'    x = 50
'    y = 50
'    Lligne = (size - 6) * 25
'    Hligne = 20
'    sepligne = Hligne + 10
'    LargCel = 40
'
'    'Création des dessins
'    For i = 1 To UBound(t())
'
'        posX = x
'        If t(i, 1) <> "" Then
'            'une ligne pour chaque flux
'            ActiveSheet.Shapes.AddLine(x, y, x + Lligne, y).Select
'
'            'une flèche selon le type de flux et une boite avec son nom
'            If t(i, 1) = "Ch" Or t(i, 1) = "UCh" Then
'                Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
'                Selection.ShapeRange.Line.ForeColor.SchemeColor = 10
'                Selection.ShapeRange.Line.Weight = 1.25
'                ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y - Hligne, LargCel, Hligne).Select
'            Else
'                Selection.ShapeRange.Line.BeginArrowheadStyle = msoArrowheadTriangle
'                Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
'                Selection.ShapeRange.Line.Weight = 1.25
'                ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, x + Lligne - LargCel, y - Hligne, LargCel, Hligne).Select
'            End If
'            Selection.Characters.Text = t(i, 2)
'            Selection.HorizontalAlignment = xlCenter
'
'            For j = 3 To size - 6
'
'                posX = posX + 25
'                If left(t(i, j), 2) = "P_" Then
'
'                    'puissance échangée dans un ovale
'                    ActiveSheet.Shapes.AddShape(msoShapeOval, posX, y - Hligne / 2, LargCel, Hligne).Select
'                    Selection.Characters.Text = Round(CDbl(Right(t(i, j), Len(t(i, j)) - 2)), 0)
'                    Selection.HorizontalAlignment = xlCenter
'
'                    'liaison entre les ovales
'                    For k = 1 To UBound(t())
'                        posX2 = x
'                        For m = 3 To size - 7
'                            If t(i, j) = t(k, m) And i < k And t(0, j) = t(0, m) Then
'                                ActiveSheet.Shapes.AddLine(posX + LargCel / 2, y + Hligne / 2, posX2 + (x + LargCel) / 2, y - Hligne / 2 + (k - i) * sepligne).Select
'                                k = UBound(t())
'                            End If
'                            posX2 = posX2 + 25
'                        Next
'                    Next
'
'                    'tempérture d'entrée du flux
'                    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, posX - LargCel, y - Hligne, LargCel, Hligne).Select
'                    Selection.Characters.Text = Round(t(i, j - 1), 1)
'                    Selection.ShapeRange.Line.Visible = msoFalse
'                    Selection.ShapeRange.fill.Visible = msoFalse
'                    Selection.HorizontalAlignment = xlRight
'                    Selection.VerticalAlignment = xlBottom
'
'                    'température de sortie du flux
'                    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, posX + LargCel, y - Hligne, LargCel, Hligne).Select
'                    Selection.Characters.Text = Round(t(i, j + 1), 1)
'                    Selection.ShapeRange.Line.Visible = msoFalse
'                    Selection.ShapeRange.fill.Visible = msoFalse
'                    Selection.HorizontalAlignment = xlLeft
'                    Selection.VerticalAlignment = xlBottom
'
'                End If
'
'                'ovale représentant un besoin chaud non couvert
'                If left(t(i, j), 4) = "Hot_" Then
'                    ActiveSheet.Shapes.AddShape(msoShapeOval, posX, y - Hligne / 2, LargCel, Hligne).Select
'                    Selection.Characters.Text = Round(CDbl(Right(t(i, j), Len(t(i, j)) - 4)), 0)
'                    Selection.HorizontalAlignment = xlCenter
'                    Selection.ShapeRange.Line.ForeColor.SchemeColor = 10
'                    Selection.ShapeRange.Line.Weight = 3
'                End If
'
'                'ovale représentant un besoin froid non couvert
'                If left(t(i, j), 5) = "Cold_" Then
'                    ActiveSheet.Shapes.AddShape(msoShapeOval, posX, y - Hligne / 2, LargCel, Hligne).Select
'                    Selection.Characters.Text = Round(CDbl(Right(t(i, j), Len(t(i, j)) - 5)), 0)
'                    Selection.HorizontalAlignment = xlCenter
'                    Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
'                    Selection.ShapeRange.Line.Weight = 3
'                End If
'            Next
'
'        End If
'
'        y = y + sepligne
'    Next
'
'    'ajout d'un repère pour le Pinch
'    posX = x
'    For j = 3 To size - 7
'        posX = posX + 25
'        If t(0, j) = "t_pinch_gch" Then
'            ActiveSheet.Shapes.AddLine(posX + 25, 10, posX + 25, 10 + (UBound(t()) + 1) * (sepligne)).Select
'            Selection.ShapeRange.Line.ForeColor.SchemeColor = 10
'            Selection.ShapeRange.Line.Weight = 3
'            j = size - 7
'        End If
'    Next
'
'    'sélection des dessins
'    ActiveSheet.Shapes.SelectAll
'    Selection.ShapeRange.Group.Name = "Group99"
'    Ltotal = ActiveSheet.Shapes("Group99").Width
'    Htotal = ActiveSheet.Shapes("Group99").Height
'
'    'ajout d'un graphique
'    Sheets("Results" & pi_FiliereConsideree).Activate
'    Charts.Add
'    If ActiveChart.SeriesCollection.Count > 0 Then  'Suppression de la série parasite si présente
'        ActiveChart.SeriesCollection(1).Delete
'    End If
'    ActiveChart.Location Where:=xlLocationAsObject, Name:="Results" & pi_FiliereConsideree
'    ActiveChart.Parent.Name = "bj"
'    hactuel = ActiveSheet.Shapes("bj").Height
'    lactuel = ActiveSheet.Shapes("bj").Width
'    ActiveSheet.Shapes("bj").ScaleWidth Ltotal / lactuel, msoFalse, msoScaleFromTopLeft
'    ActiveSheet.Shapes("bj").ScaleHeight Htotal / hactuel, msoFalse, msoScaleFromTopLeft
'    ActiveSheet.Shapes("bj").top = 10
'    ActiveSheet.Shapes("bj").left = 10
'
'    'collage des dessins et export du graphique
'    Sheets("Tampon" & pi_FiliereConsideree).Activate
'    ActiveSheet.Shapes("Group99").Select
'    Selection.Copy
'    Sheets("Results" & pi_FiliereConsideree).Activate
'    ActiveSheet.Shapes("bj").Select
'    ActiveChart.Paste
'    ActiveChart.Export FileName:="image.jpg", FilterName:="JPEG"
'
'    ActiveSheet.Range("A1").Select
'    Sheets("Tampon" & pi_FiliereConsideree).Activate
'    ActiveSheet.Shapes("Group99").Select
'    Selection.Delete
'    Sheets("Results" & pi_FiliereConsideree).Activate
'
'End Sub

Sub effacerFluxProcedes(numStation As Integer)

    '### DECLARATION DES VARIABLES
    Dim nbLigneMax As Integer
    Dim LigneTab As Integer
    Dim indiceLigne As Integer

    '### INITIALISATION DES VARIABLES
    nbLigneMax = 0
    colonneStation = 1 + (numStation - 1) * flux_thermique_process_nb_caracteristiques
    LigneTab = INDICE_PREMIERE_LIGNE

    '## Détermination du nombre de ligne
    While Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab, colonneStation) <> ""
        LigneTab = LigneTab + 1
        nbLigneMax = nbLigneMax + 1
    Wend

    '## Suppression des lignes
    If nbLigneMax <> 0 Then
        For indiceLigne = 1 To nbLigneMax
            supprimerFluxProcede (1)
        Next indiceLigne
    End If

End Sub

Sub effacerFluxUtilites(numStation As Integer)

    '### DECLARATION DES VARIABLES
    Dim nbLigneMax As Integer
    Dim LigneTab As Integer
    Dim indiceLigne As Integer

    '### INITIALISATION DES VARIABLES
    nbLigneMax = 0
    colonneStation = 1 + (numStation - 1) * flux_thermique_utilites_nb_caracteristiques
    LigneTab = INDICE_PREMIERE_LIGNE

    '## Détermination du nombre de ligne
    While Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab, colonneStation) <> ""
        LigneTab = LigneTab + 1
        nbLigneMax = nbLigneMax + 1
    Wend

    '## Suppression des lignes
    If nbLigneMax <> 0 Then
        For indiceLigne = 1 To nbLigneMax
            supprimerFluxUtilite (1)
        Next indiceLigne
    End If

End Sub

Sub effacerFluxImposition(numStation As Integer)

    '### DECLARATION DES VARIABLES
    Dim nbLigneMax As Integer
    Dim LigneTab As Integer
    Dim indiceLigne As Integer

    '### INITIALISATION DES VARIABLES
    nbLigneMax = 0
    colonneStation = 1 + (numStation - 1) * 9
    LigneTab = INDICE_PREMIERE_LIGNE

    '## Détermination du nombre de ligne
    While Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab, colonneStation) <> ""
        LigneTab = LigneTab + 1
        nbLigneMax = nbLigneMax + 1
    Wend

    '## Suppression des lignes
    If nbLigneMax <> 0 Then
        For indiceLigne = 1 To nbLigneMax
            supprimerFluxImposition (1)
        Next indiceLigne
    End If

End Sub

Sub effacerFluxInterdiction(numStation As Integer)

    '### DECLARATION DES VARIABLES
    Dim nbLigneMax As Integer
    Dim LigneTab As Integer
    Dim indiceLigne As Integer

    '### INITIALISATION DES VARIABLES
    nbLigneMax = 0
    colonneStation = 1 + (numStation - 1) * 3
    LigneTab = INDICE_PREMIERE_LIGNE

    '## Détermination du nombre de ligne
    While Feuil63.Range("DonneesPincementInterdictionEchange").Cells(LigneTab, colonneStation) <> ""
        LigneTab = LigneTab + 1
        nbLigneMax = nbLigneMax + 1
    Wend

    '## Suppression des lignes
    If nbLigneMax <> 0 Then
        For indiceLigne = 1 To nbLigneMax
            supprimerFluxInterdiction (1)
        Next indiceLigne
    End If

End Sub

'Sub TDC()
'
'    '### DECLARATION DES VARIABLES
'    Dim nbLigneTotProc As Integer
'    Dim nbLigneTotImp As Integer
'    Dim nbLigneTotInt As Integer
'    Dim ligneTab As Integer
'    Dim indiceLigneImp As Integer
'    Dim indiceLigneInt As Integer
'    Dim indiceLigneProc As Integer
'    Dim imposASupp() As Integer
'    Dim interASupp() As Integer
'    Dim existanceFlux As Boolean
'    Dim nomRef As String
'    Dim nomFluxChaud As String
'    Dim nomFluxFroid As String
'
'    '### INITIALISATION DES VARIABLES
'    ReDim imposASupp(1 To 1) As Integer
'    ReDim interASupp(1 To 1) As Integer
'
'    '## Détermination des nombres de ligne
'
'    '# Procédés
'    colonneStation = 1 + (numStationSelectionnee - 1) * flux_thermique_process_nb_caracteristiques
'    While Feuil63.Range("DonneesPincementFluxProcedes").Cells(ligneTab, colonneStation) <> ""
'        ligneTab = ligneTab + 1
'        nbLigneTotProc = nbLigneTotProc + 1
'    Wend
'
'    '# Imposition
'    colonneStation = 1 + (numStationSelectionnee - 1) * 9
'    While Feuil63.Range("DonneesPincementFluxImposition").Cells(ligneTab, colonneStation) <> ""
'        ligneTab = ligneTab + 1
'        nbLigneTotImp = nbLigneTotImp + 1
'    Wend
'
'    '## Interdiction
'    colonneStation = 1 + (numStationSelectionnee - 1) * 3
'    While Feuil63.Range("DonneesPincementFluxInterdiction").Cells(ligneTab, colonneStation) <> ""
'        ligneTab = ligneTab + 1
'        nbLigneTotInt = nbLigneTotInt + 1
'    Wend
'
'    '## Comparaison des noms
'
'    'Imposition: parcourt de toutes les lignes
'    If nbLigneTotImp <> 0 Then
'
'        'Boucle sur les échangeurs imposés
'        For indiceLigneImp = 0 To nbLigneTotImp
'
'            existanceFlux = False
'
'            'Boucle sur les procédés
'            For indiceLigneProc = 0 To nbLigneTotProc
'
'                'Si un des flux de l'échangeur n'est
'                If Feuil63.Range("DonneesPincementFluxProcedes").Cells(INDICE_PREMIERE_LIGNE + indiceLigneProc, colonneStation + 1) = Feuil63.Range("DonneesPincementFluxImposition").Cells(INDICE_PREMIERE_LIGNE + indiceLigne, colonneStation + 1).Value _
                 '                   Or Feuil63.Range("DonneesPincementFluxProcedes").Cells(INDICE_PREMIERE_LIGNE + indiceLigneProc, colonneStation + 1) = Feuil63.Range("DonneesPincementFluxImposition").Cells(INDICE_PREMIERE_LIGNE + indiceLigne, colonneStation + 2).Value Then
'
'                    existanceFlux = True
'
'                End If
'                If existanceFlux = False Then
'                    imposASupp(UBound(imposASupp)) = indiceLigneProc
'                    ReDim Preserve imposASupp(1 To UBound(imposASupp) + 1)
'                End If
'
'            Next indiceLigneProc
'
'            supprimerFluxInterdiction (INDICE_PREMIERE_LIGNE + indiceLigne)
'
'        Next indiceLigne
'
'    End If
'
'End Sub

Public Sub effacerLaFiliere(numStation As Integer)


    'flux procédés
    effacerFluxProcedes (numStation)

    'flux utilités
    effacerFluxUtilites (numStation)

    'flux imposition
    effacerFluxImposition (numStation)

    'flux interdiction
    effacerFluxInterdiction (numStation)

    'Valeur pincement


    'Tri auto Procédés


    'Tri auto Utilités


    'Seuil puissance


End Sub

' procédure pour remplir les informations necessaires au pincement thermique
Public Sub copie_dans_flux_thermique_process(ByVal nom_flux As String, ByVal T_in As Double, ByVal T_out As Double, ByVal MCp_kWh_j°C As Double)

    Const zero_force = 0.00000000001

    'ajout d'un nouveau flux "i" dans le tableau
    Dim I As Integer    ' indice de la ligne
    Dim Pw_kWhj As Double

    Pw_kWhj = Abs(MCp_kWh_j°C * (T_in - T_out))

    If Pw_kWhj > zero_force Then
        If flux_thermique_process(repere_flux_thermique_process_etat, 1) = "" Then      'si le flux est le premier a être entré dans le tableau -> on écrit dans la première ligne de celui-ci
            I = 1
        Else        'sinon -> ajout d'une ligne au tableau avec conservation des données
            I = UBound(flux_thermique_process, 2) + 1
        End If

        ReDim Preserve flux_thermique_process(1 To flux_thermique_process_nb_caracteristiques, 1 To I)

        If MCp_kWh_j°C <= 0 Then
            MsgBox nom_flux   'Feuil5.Range("erreur_flux_thermique") & ": " & nom_flux, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
            pn_EffacementResultatsEnCours = True
            pn_ErreurNonGeree = True
            Exit Sub
        End If

        'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
        flux_thermique_process(repere_flux_thermique_process_etat, I) = True
        flux_thermique_process(repere_flux_thermique_process_nom, I) = nom_flux
        flux_thermique_process(repere_flux_thermique_process_T_in_°C, I) = T_in
        flux_thermique_process(repere_flux_thermique_process_T_out_°C, I) = T_out

        flux_thermique_process(repere_flux_thermique_process_MCp_kW_°C, I) = MCp_kWh_j°C / pd_NOMBRE_HEURE_PAR_JOUR
        flux_thermique_process(repere_flux_thermique_process_original, I) = True
    End If

End Sub

