Option Explicit

'Public Type typeFlux
'    MES As Integer
'    DCO As String
'    DBO(1 To 3) As Double
'End Type
'
'Public Flux1 As typeFlux
'Public flux(1 To 10) As typeFlux
'
'Sub toto()
'
' Flux1.DBO (1)
'
' With Flux1
'    .DCO = "tt2"
' End With
'
'End Sub

    
Public Function besoins_O2_HS(ByRef HS)

    'Ici on va calculer pour la biologie la consommation en O2 due à HS-

    besoins_O2_HS = 2 * HS    'on a besoin de 2gO2/gS
    HS = 0   'hypothèse: on consomme tout

End Function


Public Function humidite_air_gH2O_gAS(ByVal MM_air_sec, ByVal P_absolue_bar, ByVal T_°C, ByVal humidite_relative)


    Const MM_H2O = 18   'g/mol

    Dim Pvap_bar As Double


    'calcul de la pression de vapeur saturante à T
    Pvap_bar = humidite_relative * Pv_sat_H2O_bar(T_°C)


    'calcul de l'humidité absolue
    humidite_air_gH2O_gAS = (MM_H2O / MM_air_sec) * (Pvap_bar / (P_absolue_bar - Pvap_bar))

    If humidite_air_gH2O_gAS < 0 Then
        humidite_air_gH2O_gAS = 0
    End If

End Function


Public Function Pv_sat_H2O_bar(ByVal T_°C)

    Dim t As Double

    'valable entre 0 et 373 °C    fichier "Données thermodynamiques ATHOS"
    If T_°C < 0 Then
        t = 0
    ElseIf T_°C > 373 Then
        t = 373
    Else
        t = T_°C
    End If

    'calcul de Pv sat
    Pv_sat_H2O_bar = 9.481E-14 * t ^ 6 - 0.00000000009203 * t ^ 5 + 0.00000004801 * t ^ 4 - 0.00000731 * t ^ 3 + 0.0005988 * t ^ 2 - 0.01745 * t + 0.136


End Function


Sub attribuerNomVF()

    Dim pb As Integer
    Dim Ligne As Integer
    Dim numLigne As Integer
    Dim nomParam As String
    Dim x As String

    Const FEUIL_SOURCE = "Feuil1"
    Const FEUIL_CIBLE = "AA"
    Const COL_NUM = "C"
    Const COL_NOM = "D"

    pb = 0

    For Ligne = 1 To 144
        numLigne = Worksheets(FEUIL_SOURCE).Range(COL_NUM & Ligne).Value
        nomParam = CStr(Worksheets(FEUIL_SOURCE).Range(COL_NOM & Ligne).Value)

        If numLigne <> 0 Then
            'changement de couleur
            Worksheets(FEUIL_CIBLE).Range("T" & numLigne).Select
            With Selection.Interior
                .ColorIndex = 4   '4 : vert; 2: blanc
                .Pattern = xlSolid
            End With

            'nom Cel
            x = CStr("=" & FEUIL_CIBLE & "!R" & numLigne & "C9")
            ActiveWorkbook.Names.Add Name:=nomParam, RefersToR1C1:=x
        Else
            pb = pb + 1
        End If

    Next Ligne

    Debug.Print pb

End Sub

Sub attribuerNomVF2()

    Dim Ligne As Integer
    Dim numLigne As Integer
    Dim nomParam As String
    Dim x As String

    Const FEUIL_SOURCE = "Feuil1"
    Const FEUIL_CIBLE = "PINCH"
    Const COL_NOM = "F"


    For Ligne = 1 To 87
        numLigne = 16
        nomParam = CStr(Worksheets(FEUIL_SOURCE).Range(COL_NOM & Ligne).Value)
        'Worksheets("Calculs").Range("T" & numLigne).Value = nomParam
        'Worksheets("Calculs").Range("T" & numLigne).ClearContents

            'nom Cel
            x = CStr("=" & FEUIL_CIBLE & "!R" & numLigne & "C" & Ligne)
            ActiveWorkbook.Names.Add Name:=nomParam, RefersToR1C1:=x

            Worksheets(FEUIL_CIBLE).Cells(16, Ligne).Value = nomParam

    Next Ligne

End Sub

'####################################################################  AUDIT DE FORMULE
Sub Audit_dep()
    ListeDependents (Worksheets("Calculs").Range("C530"))
End Sub

Sub ListeDependents(Cellule As Range)
    Dim Ws As Worksheet
    Dim Plage As Range, Cell As Range, DirectDep As Range
    Dim I As Integer, x As Integer
    Dim Cible As String, strDepenDent As String, strRefer As String
    'La liste des dépendants  va être stockée dans une collection
    Dim Un As New Collection


    'Active la feuille contenant la cellule à contrôler
    Cellule.Parent.Activate

    strDepenDent = Cellule.Parent.Name & "!" & Cellule.Address(0, 0)


    'Vérifie s'il y a des dépendants directs dans la feuille:
    If version_Developpement = False Then On Error Resume Next
    'Définit la plage de cellules dépendantes, dans la feuille active
    Set Plage = Cellule.DirectDependents.Cells
    If version_Developpement = False Then On Error GoTo 0

    If Not Plage Is Nothing Then
        'Boucle sur les dépendants contenus dans la feuille active
        For Each DirectDep In Cellule.DirectDependents.Cells
            Un.Add Cellule.Parent.Name & "!" & DirectDep.Address, _
                   Cellule.Parent.Name & "!" & DirectDep.Address
        Next DirectDep
    End If

    Set Plage = Nothing



    'Boucle sur les autres feuilles du classeur:
    For Each Ws In ThisWorkbook.Worksheets
        'Si la feuille est différente de la feuille active
        If Ws.Name <> Cellule.Parent.Name Then

            If version_Developpement = False Then On Error Resume Next
            'Définit la plage de cellules contenant des formules
            Set Plage = Ws.UsedRange.SpecialCells(xlCellTypeFormulas)
            If version_Developpement = False Then On Error GoTo 0


            'Vérifie si la feuille contient des formules
            If Not Plage Is Nothing Then

                'Boucle sur les cellules contenant des formules
                For Each Cell In Plage

                    'Gestion des références relatives et absolues
                    Cible = Replace(Cell.Formula, "$", "")


                    'Vérifie si le nom de la feuille apparait dans la formule.
                    If InStr(1, Cible, Cellule.Parent.Name) > 0 Then

                        'Vérifie si la formule contient une référence correspondant à la
                        'cellule à contrôler
                        I = 0
                        I = InStr(1, Cible, strDepenDent)

                        'Si la référence est trouvée on l'intègre dans la collection
                        If I > 0 And Not IsNumeric(Mid(Cible, I + Len(strDepenDent), 1)) Then

                            Un.Add Ws.Name & "!" & Cell.Address, Ws.Name & "!" & Cell.Address

                        Else

                            'Recherche des références dans les plages de cellules
                            For x = 1 To Len(Cible)
                                I = 0
                                I = InStr(1, Cible, ":")

                                If I > 0 Then
                                    strRefer = ExtractionReferences(Cible)

                                    'Si la cellule à contrôler se trouve dans la plage,
                                    'on l'intègre dans la collection.
                                    If VerifIntersect(Cellule, Range(strRefer)) And _
                                       InStr(1, Cible, Cellule.Parent.Name & "!" & strRefer) > 0 Then

                                        If version_Developpement = False Then On Error Resume Next
                                        Un.Add Ws.Name & "!" & Cell.Address, Ws.Name & "!" & Cell.Address
                                        If version_Developpement = False Then On Error GoTo 0

                                        Exit For
                                    End If

                                    Cible = Mid(Cible, I + 1)
                                Else
                                    Exit For
                                End If
                            Next x
                            '--------------
                        End If
                    End If
                Next Cell

            End If
        End If

        Set Plage = Nothing
    Next Ws


    'Boucle sur la collection qui contient la liste des dépendants
    For I = 1 To Un.Count
        'Affiche le résultat dans la fenêtre d'exécution (Ctrl+G)
        Debug.Print Un.item(I)
    Next I
End Sub

'Extrait les références spécifiées dans les formules
Function ExtractionReferences(Chaine As String) As String
    Dim I As Integer, j As Integer
    Dim strPlage As String, Caract As String

    I = InStr(1, Chaine, ":")


    'Renvoie la référence avant les deux points ":"
    For j = I - 1 To 1 Step -1
        Caract = Mid(Chaine, j, 1)

        Select Case Asc(Caract)
        Case 48 To 57, 65 To 90, 97 To 122
            strPlage = Caract & strPlage
        Case Else: Exit For
        End Select
    Next j

    strPlage = strPlage & ":"

    'Renvoie la référence après les deux points ":"
    For j = I + 1 To Len(Chaine)
        Caract = Mid(Chaine, j, 1)

        Select Case Asc(Caract)
        Case 48 To 57, 65 To 90, 97 To 122
            strPlage = strPlage & Caract
        Case Else: Exit For
        End Select
    Next j

    ExtractionReferences = strPlage
End Function


'Vérifie si la référence extraite dans la formule a une intersection
'avec la cellule dont on contrôle les dépendances.
Function VerifIntersect(objDepend As Range, objReference As Range) As Boolean
    Dim objRange As Range

    Set objRange = Intersect(objDepend, objReference)

    If objRange Is Nothing Then
        VerifIntersect = False
    Else
        VerifIntersect = True
    End If
End Function
'####################################################################




















