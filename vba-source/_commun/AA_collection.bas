
'Private Sub Worksheet_Change(ByVal Target As Range)
' If Target.Address = "$B$2" Then
'    Select Case Target.Value
'        Case Is = 2
'            Macro1
'        Case Is = 3
'            macro2
'        Case Is = 4
'            macro3
'    End Select
' End If
'
'End Sub



Sub ListeDictionnaire()
  Set d = CreateObject("Scripting.Dictionary")
  d.item("Dupont") = 35      ' ou If Not d.Exists("aa") Then d.Add "Dupont", 35
  d.item("Durand") = 40
  d.item("Martin") = 27       ' ou d("Durand")=40
  d.item("Espinasse") = 32
  '---- élément pour une clé
  clé = "Durand"
  MsgBox clé & ":" & d.item(clé)   ' ou MsgBox d(clé)
  '--- toutes les clés et valeurs associées
  For Each c In d.keys
     MsgBox c & ":" & d.item(c) ' ou MsgBox c & ":" & d(c)
  Next c
  '---- 3eme élément
  a = d.keys     ' dans un tableau a(0 To d.Count-1)
  b = d.items    ' dans un tableau b(0 To d.Count-1)
  MsgBox a(2) & ":" & b(2)
  '--- Rang d'une clé
  clé = "Durand"
  p = Application.Match(clé, d.keys, 0)
  MsgBox "position de " & clé & ":" & p
  '--- Stats
  MsgBox "Total:" & Application.Sum(d.items)
  MsgBox "Mini:" & Application.Min(d.items)
  MsgBox "Moyenne:" & Application.Average(d.items)
End Sub




'Public MaCollectionBoue As Collection

Public Sub AjoutCollection()
     Set MaCollection = New Collection

     MaCollectionBoue.Add (NewObject1)
End Sub


Public ListeDictionnaire2 As Object

Public codig_graisses

'Public Sub ListeDictionnaire2()
'  Set codig_graisses = CreateObject("Scripting.Dictionary")
'
'  'sampleVisualBasicColl.Add(item1, "firstkey")
'
'  codig_graisses.item("case"NK_MV"") = 0.002      ' ou If Not d.Exists("aa") Then d.Add "Dupont", 35
'  codig_graisses.item("case"Pt_MES"") = 0.002
'  codig_graisses.item("case"DCO_MV"") = 2.8     ' ou d("Durand")=40
'  codig_graisses.item(" case"DBO_MV"") = 2.5
'
'
' a = codig_graisses.item("case"NK_MV"")
'
'
'End Sub



Function les_notes(prenom)
Set Notes = CreateObject("Scripting.Dictionary")
Notes.Add "Arthur", 12
Notes.Add "Sophie", 17
Notes.Add "Nicolas", 8
Notes.Add "Emilie", "absente"
Notes.Add "Marie", 18
les_notes = Notes(prenom)
Set Notes = Nothing
End Function
'Si on écrit =les_notes("Marie") dans une page Excel, on obtient 18.


Sub tt()
 b = les_notes("Marie")

End Sub




Public Function ratio(type_boue, type_ratio)

Select Case type_boue

Case "I_simple", "I_reactif"  'primaire reactif'primaire simple
        Select Case type_ratio
            Case "NK_MV": ratio = 0.059
            Case "Pt_MES": ratio = 0.009 'primaire simple
            Case "DCO_MV": ratio = 1.71 'primaire simple
            Case "DBO_MV": ratio = 0.7 'primaire simple
        End Select
Case "II_forte" 'forte charge
        Select Case type_ratio
            Case "NK_MV": ratio = 0.065
            Case "Pt_MES": ratio = 0.014 'primaire simple
            Case "DCO_MV": ratio = 1.71 'primaire simple
            Case "DBO_MV": ratio = 0.9 'primaire simple
        End Select
Case "II_moyenne" 'moyenne charge
        Select Case type_ratio
            Case "NK_MV": ratio = 0.08
            Case "Pt_MES": ratio = 0.02 'primaire simple
            Case "DCO_MV": ratio = 1.6 'primaire simple
            Case "DBO_MV": ratio = 0.75 'primaire simple
        End Select
Case "II_faible_EB", "II_HybAS" 'faible charge eau brute
        Select Case type_ratio
            Case "NK_MV": ratio = 0.095
            Case "Pt_MES": ratio = 0.02 'primaire simple
            Case "DCO_MV": ratio = 1.5 'primaire simple
            Case "DBO_MV": ratio = 0.5 'primaire simple
        End Select
Case "II_faible_ED", "II_prolongee_EB", "II_prolongee_ED" 'faible charge eau décantée 'aération prolongée eau brute'aération prolongée eau brute
        Select Case type_ratio
            Case "NK_MV": ratio = 0.095
            Case "Pt_MES": ratio = 0.02 'primaire simple
            Case "DCO_MV": ratio = 1.45 'primaire simple
            Case "DBO_MV": ratio = 0.5 'primaire simple
        End Select
Case "II_MBR", "II_MBBR"  'MBR    A REVOIR
        Select Case type_ratio
            Case "NK_MV": ratio = 0.095
            Case "Pt_MES": ratio = 0.02 'primaire simple
            Case "DCO_MV": ratio = 1.5 'primaire simple
            Case "DBO_MV": ratio = 0.5 'primaire simple
        End Select
Case "II_biostyr_C" 'biostyr_C
        Select Case type_ratio
            Case "NK_MV": ratio = 0.01
            Case "Pt_MES": ratio = 0.025
            Case "DCO_MV": ratio = 1.6 'biostyr_C
            Case "DBO_MV": ratio = 0.9 'biostyr_C
        End Select
Case "II_biostyr_N" 'biostyr_N
        Select Case type_ratio
            Case "NK_MV": ratio = 0.01
            Case "Pt_MES": ratio = 0.025
            Case "DCO_MV": ratio = 1.55 'biostyr_C
            Case "DBO_MV": ratio = 0.9 'biostyr_C
        End Select
Case "II_biostyr_NDN" 'biostyr_NDN
        Select Case type_ratio
            Case "NK_MV": ratio = 0.01
            Case "Pt_MES": ratio = 0.025
            Case "DCO_MV": ratio = 1.5 'biostyr_C
            Case "DBO_MV": ratio = 0.9 'biostyr_C
        End Select
Case "II_biostyr_PDN" 'biostyr_PDN
        Select Case type_ratio
            Case "NK_MV": ratio = 0.01
            Case "Pt_MES": ratio = 0.015
            Case "DCO_MV": ratio = 1.45 'biostyr_C
            Case "DBO_MV": ratio = 0.9 'biostyr_C
        End Select
Case "III_decantation", "III_biostyr_N", "III_biostyr_PDN"  'décantation III  'biostyr_N III 'biostyr_PDN III
            Select Case type_ratio
                Case "NK_MV": ratio = 0.041
                Case "Pt_MES": ratio = 0.04
                Case "DCO_MV": ratio = 1.45 'biostyr_C
                Case "DBO_MV": ratio = 0.45 'biostyr_C
            End Select
Case "codigestion_graisses" 'graisses
        Select Case type_ratio
            Case "NK_MV": ratio = 0.002
            Case "Pt_MES": ratio = 0.002
            Case "DCO_MV": ratio = 2.8
            Case "DBO_MV": ratio = 2.8
        End Select
End Select
End Function

