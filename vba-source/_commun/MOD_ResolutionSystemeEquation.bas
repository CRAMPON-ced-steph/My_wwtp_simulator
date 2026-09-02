Option Explicit

Private Function F_determinant_matrice(ByVal DimensionMatrice As Integer, ByRef Matrice() As Double) As Double


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure permet de calculer le déterminant d'une matrice carrée de dimension "DimensionMatrice"

'VARIABLES D'ENTREE:
'   DimensionMatrice: contient la taille de la matrice carrée
'   Matrice(): contient la matrice dont le déterminant doit être calculé

'VARIABLES DE SORTIE:
'   F_determinant_matrice: valeur de la fonction
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim Ligne As Integer
Dim Colonne As Integer
Dim ValeurIntermediaire As Double
Dim facteur As Double
Dim pivot As Double
Dim LigneEchangee As Integer
Dim NombreColonnes As Integer
Dim MatriceNonInversible As Boolean

'Initialisation des variables locales
NumeroErreurNonGeree = "MOD 16.2"
I = 0
Ligne = 0
Colonne = 0
ValeurIntermediaire = 0
facteur = 0
pivot = 0
LigneEchangee = 0
NombreColonnes = DimensionMatrice
MatriceNonInversible = False

'Diagonalisation de la matrice
For I = 1 To DimensionMatrice
    pivot = Matrice(I, I)
    If pivot = 0 Then
        LigneEchangee = I
        Do While Matrice(LigneEchangee, I) = 0 And MatriceNonInversible = False
            LigneEchangee = LigneEchangee + 1
            If LigneEchangee > DimensionMatrice Then
                MatriceNonInversible = True
            End If
        Loop
        If MatriceNonInversible = True Then
            GoTo determinant_nul
        End If
        For Colonne = I To NombreColonnes
            ValeurIntermediaire = Matrice(LigneEchangee, Colonne)
            Matrice(LigneEchangee, Colonne) = Matrice(I, Colonne)
            Matrice(I, Colonne) = ValeurIntermediaire
        Next Colonne
        pivot = Matrice(I, I)
    End If
    For Ligne = 1 To DimensionMatrice
        If Ligne <> I Then
            If Matrice(Ligne, I) <> 0 Then
                facteur = Matrice(Ligne, I) / pivot
                For Colonne = I To NombreColonnes
                    Matrice(Ligne, Colonne) = Matrice(Ligne, Colonne) - facteur * Matrice(I, Colonne)
                Next Colonne
            End If
        End If
    Next Ligne
Next I

'Calcul du déterminant
F_determinant_matrice = 1
For I = 1 To DimensionMatrice
    F_determinant_matrice = F_determinant_matrice * Matrice(I, I)
Next I


Exit Function
'*********************************************************************
determinant_nul:
F_determinant_matrice = 0

Exit Function
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Function

Public Sub solve_systeme_lineaire_Gauss(ByVal DimensionMatrice As Integer, ByRef Matrice() As Double)


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure permet de résoudre des systèmes d'équations linéaires par la méthode du pivot de Gauss
'   Pour cela, elle utilise la fonction suivante:
'       - "F_determinant_matrice" qui calcule le déterminant du système

'VARIABLES D'ENTREE:
'   DimensionMatrice: contient la taille du système d'équations à résoudre
'   Matrice(): contient le système linéaire à résoudre

'VARIABLES DE SORTIE:
'   Matrice(): on récupère la matrice après résolution par le pivot de Gauss
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim Ligne As Integer
Dim Colonne As Integer
Dim MatriceIntermediaire() As Double
Dim ValeurIntermediaire As Double
Dim facteur As Double
Dim pivot As Double
Dim LigneEchangee As Integer
Dim NombreColonnes As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "MOD 16.1"
I = 0
Ligne = 0
Colonne = 0
ReDim MatriceIntermediaire(1 To DimensionMatrice, 1 To DimensionMatrice)
    For Ligne = 1 To DimensionMatrice
        For Colonne = 1 To DimensionMatrice
            MatriceIntermediaire(Ligne, Colonne) = Matrice(Ligne, Colonne)
        Next Colonne
    Next Ligne
ValeurIntermediaire = 0
facteur = 0
pivot = 0
LigneEchangee = 0
NombreColonnes = 0

If F_determinant_matrice(DimensionMatrice, MatriceIntermediaire) <> 0 Then
    NombreColonnes = DimensionMatrice + 1
    For I = 1 To DimensionMatrice
        pivot = Matrice(I, I)
        If pivot = 0 Then
            LigneEchangee = I
            Do While Matrice(LigneEchangee, I) = 0
                LigneEchangee = LigneEchangee + 1
            Loop
            For Colonne = I To NombreColonnes
                ValeurIntermediaire = Matrice(LigneEchangee, Colonne)
                Matrice(LigneEchangee, Colonne) = Matrice(I, Colonne)
                Matrice(I, Colonne) = ValeurIntermediaire
            Next Colonne
            pivot = Matrice(I, I)
        End If
        For Ligne = 1 To DimensionMatrice
            If Ligne <> I Then
                If Matrice(Ligne, I) <> 0 Then
                    facteur = Matrice(Ligne, I) / pivot
                    For Colonne = I To NombreColonnes
                        Matrice(Ligne, Colonne) = Matrice(Ligne, Colonne) - facteur * Matrice(I, Colonne)
                    Next Colonne
                End If
            Else
                facteur = 1 / pivot
                For Colonne = I To NombreColonnes
                    Matrice(Ligne, Colonne) = Matrice(Ligne, Colonne) * facteur
                Next Colonne
                pivot = Matrice(Ligne, I)
            End If
        Next Ligne
    Next I
Else
    MsgBox ("système sans solution")
End If

Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub



