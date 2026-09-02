Option Explicit

'erreurs
Private Const nb_erreurs = 2

'constantes sur le choix du réactif
Private Const Cl2_gazeux = 1
Private Const HClO = 2

'variables de choix
Dim choix_reactif As Integer

'variables échangées avec l'utilisateur
Dim Q_ratio_admis As Double
'Dim reactif_purete As Double
Dim C_dosage_mgL As Double
Dim CN_dosage_mg_mgNH4 As Double
Dim CN_NO2_mgL As Double
Dim CN_dosage_mg_mgNO2 As Double
Dim CN_residuel_mgL As Double
Dim CN_dosage_resultant_mgL As Double

'Public conso_commercial_kgj As Double
Dim conso_Cl2_pur_kgj As Double
Public Cl2_gazeux_pur_kgj As Double
Public Cl2_HClO_pur_kgj As Double

'variables repérant les valeurs forcées
Dim Q_ratio_admis_force As Boolean
'Dim reactif_purete_force As Boolean
Dim C_dosage_mgL_force As Boolean
Dim CN_dosage_mg_mgNH4_force As Boolean
Dim CN_NO2_mgL_force As Boolean
Dim CN_dosage_mg_mgNO2_force As Boolean
Dim CN_residuel_mgL_force As Boolean

'variables stockant les valeurs forcées
Dim Q_ratio_admis_f As Double
'Dim reactif_purete_f As Double
Dim C_dosage_mgL_f As Double
Dim CN_dosage_mg_mgNH4_f As Double
Dim CN_NO2_mgL_f As Double
Dim CN_dosage_mg_mgNO2_f As Double
Dim CN_residuel_mgL_f As Double


Dim remarque(nb_erreurs) As String

Sub affich_remarque()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = nb_erreurs  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil52.Range("desinfection_Cl_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "52.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub calcul_consommation_electrique()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'electricite = 0



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "52.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub dimensionnement(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    If pi_ChoixQualiteRejet = pi_QUALITE_REJET_TRAITEMENT_CARBONE Then
        conso_Cl2_pur_kgj = C_dosage_mgL * Q_ratio_admis * Q / 1000
    Else
        conso_Cl2_pur_kgj = CN_dosage_mg_mgNH4 * Q_ratio_admis * NH4
        conso_Cl2_pur_kgj = conso_Cl2_pur_kgj + CN_dosage_mg_mgNO2 * CN_NO2_mgL * Q_ratio_admis * Q / 1000
        conso_Cl2_pur_kgj = conso_Cl2_pur_kgj + CN_residuel_mgL * Q_ratio_admis * Q / 1000
        CN_dosage_resultant_mgL = conso_Cl2_pur_kgj * 1000 / (Q_ratio_admis * Q)
    End If

    If choix_reactif = Cl2_gazeux Then
        Cl2_gazeux_pur_kgj = conso_Cl2_pur_kgj
    Else
        Cl2_HClO_pur_kgj = conso_Cl2_pur_kgj
    End If

    'conso_commercial_kgj = conso_Cl2_pur_kgj / reactif_purete


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "52.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub


Sub ecriture_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil52.Range("desinfection_Cl_Q_ratio_admis").Cells(1, Colonne) = Q_ratio_admis
    'Feuil52.Range("desinfection_Cl_reactif_purete").Cells(1, colonne) = reactif_purete
    If pi_ChoixQualiteRejet = pi_QUALITE_REJET_TRAITEMENT_CARBONE Then
        Feuil52.Range("desinfection_Cl_C_dosage_mgL").Cells(1, Colonne) = C_dosage_mgL
    Else
        Feuil52.Range("desinfection_Cl_CN_dosage_mg_mgNH4").Cells(1, Colonne) = CN_dosage_mg_mgNH4
        Feuil52.Range("desinfection_Cl_CN_NO2_mgL").Cells(1, Colonne) = CN_NO2_mgL
        Feuil52.Range("desinfection_Cl_CN_dosage_mg_mgNO2").Cells(1, Colonne) = CN_dosage_mg_mgNO2
        Feuil52.Range("desinfection_Cl_CN_residuel_mgL").Cells(1, Colonne) = CN_residuel_mgL
        Feuil52.Range("desinfection_Cl_CN_dosage_resultant_mgL").Cells(1, Colonne) = CN_dosage_resultant_mgL
    End If
    'Feuil52.Range("desinfection_Cl_conso_commercial_kgj").Cells(1, colonne) = conso_commercial_kgj


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "52.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil52.Range("desinfection_Cl_Q_ratio_admis").Cells(1, Colonne) = ""    'Q_ratio_admis
    Feuil52.Range("desinfection_Cl_reactif_purete").Cells(1, Colonne) = ""    'reactif_purete
    Feuil52.Range("desinfection_Cl_C_dosage_mgL").Cells(1, Colonne) = ""    'C_dosage_mgL
    Feuil52.Range("desinfection_Cl_CN_dosage_mg_mgNH4").Cells(1, Colonne) = ""    'CN_dosage_mg_mgNH4
    Feuil52.Range("desinfection_Cl_CN_NO2_mgL").Cells(1, Colonne) = ""    'CN_NO2_mgL
    Feuil52.Range("desinfection_Cl_CN_dosage_mg_mgNO2").Cells(1, Colonne) = ""    'CN_dosage_mg_mgNO2
    Feuil52.Range("desinfection_Cl_CN_residuel_mgL").Cells(1, Colonne) = ""    'CN_residuel_mgL
    Feuil52.Range("desinfection_Cl_CN_dosage_resultant_mgL").Cells(1, Colonne) = ""    'CN_dosage_resultant_mgL
    Feuil52.Range("desinfection_Cl_conso_commercial_kgj").Cells(1, Colonne) = ""    'conso_commercial_kgj




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "52.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub fonctionnement_reel(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    If pi_ChoixQualiteRejet = pi_QUALITE_REJET_TRAITEMENT_CARBONE Then
        conso_Cl2_pur_kgj = C_dosage_mgL * Q_ratio_admis * Q / 1000
    Else
        conso_Cl2_pur_kgj = CN_dosage_mg_mgNH4 * Q_ratio_admis * NH4
        conso_Cl2_pur_kgj = conso_Cl2_pur_kgj + CN_dosage_mg_mgNO2 * CN_NO2_mgL * Q_ratio_admis * Q / 1000
        conso_Cl2_pur_kgj = conso_Cl2_pur_kgj + CN_residuel_mgL * Q_ratio_admis * Q / 1000
        CN_dosage_resultant_mgL = conso_Cl2_pur_kgj * 1000 / (Q_ratio_admis * Q)
    End If

    If choix_reactif = Cl2_gazeux Then
        Cl2_gazeux_pur_kgj = conso_Cl2_pur_kgj
    Else
        Cl2_HClO_pur_kgj = conso_Cl2_pur_kgj
    End If

    'conso_commercial_kgj = conso_Cl2_pur_kgj / reactif_purete


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "52.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub



Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'part du débit admis sur l'ouvrage
    If Q_ratio_admis_force = True Then
        Q_ratio_admis = Q_ratio_admis_f
    Else
        Q_ratio_admis = 1
    End If

    'pureté du réactif
    'If reactif_purete_force = True Then
    '    reactif_purete = reactif_purete_f
    'Else
    '    If choix_reactif = Cl2_gazeux Then
    '        reactif_purete = 1
    '    Else
    '        reactif_purete = 0.15
    '    End If
    'End If

    'paramètre sur les doses de chlore
    If pi_ChoixQualiteRejet = pi_QUALITE_REJET_TRAITEMENT_CARBONE Then
        If C_dosage_mgL_force = True Then
            C_dosage_mgL = C_dosage_mgL_f
        Else
            C_dosage_mgL = 10  'mg/L
        End If
    Else
        If CN_dosage_mg_mgNH4_force = True Then
            CN_dosage_mg_mgNH4 = CN_dosage_mg_mgNH4_f
        Else
            CN_dosage_mg_mgNH4 = 5    'mgCl2/mgNNH4
        End If
        If CN_NO2_mgL_force = True Then
            CN_NO2_mgL = CN_NO2_mgL_f
        Else
            CN_NO2_mgL = 0
        End If
        If CN_dosage_mg_mgNO2_force = True Then
            CN_dosage_mg_mgNO2 = CN_dosage_mg_mgNO2_f
        Else
            CN_dosage_mg_mgNO2 = 5   'mgCl2/mgNNO2
        End If
        If CN_residuel_mgL_force = True Then
            CN_residuel_mgL = CN_residuel_mgL_f
        Else
            CN_residuel_mgL = 1
        End If
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "52.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "52.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    choix_reactif = Feuil52.Range("desinfection_Cl_choix_reactif").Cells(1, pi_FiliereConsideree)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "52.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim Colonne As Integer

    Colonne = pi_FiliereConsideree

    Q_ratio_admis_force = Feuil52.Range("desinfection_Cl_Q_ratio_admis").Cells(1, Colonne)
    'reactif_purete_force = Feuil52.Range("desinfection_Cl_reactif_purete").Cells(1, colonne)
    If pi_ChoixQualiteRejet = pi_QUALITE_REJET_TRAITEMENT_CARBONE Then
        C_dosage_mgL_force = Feuil52.Range("desinfection_Cl_C_dosage_mgL").Cells(1, Colonne)
    Else
        CN_dosage_mg_mgNH4_force = Feuil52.Range("desinfection_Cl_CN_dosage_mg_mgNH4").Cells(1, Colonne)
        CN_NO2_mgL_force = Feuil52.Range("desinfection_Cl_CN_NO2_mgL").Cells(1, Colonne)
        CN_dosage_mg_mgNO2_force = Feuil52.Range("desinfection_Cl_CN_dosage_mg_mgNO2").Cells(1, Colonne)
        CN_residuel_mgL_force = Feuil52.Range("desinfection_Cl_CN_residuel_mgL").Cells(1, Colonne)
    End If

    If Q_ratio_admis_force = True Then
        Q_ratio_admis_f = Feuil52.Range("desinfection_Cl_Q_ratio_admis").Cells(1, Colonne)
    End If
    'If reactif_purete_force = True Then
    '    reactif_purete_f = Feuil52.Range("desinfection_Cl_reactif_purete").Cells(1, colonne)
    'End If
    If C_dosage_mgL_force = True Then
        C_dosage_mgL_f = Feuil52.Range("desinfection_Cl_C_dosage_mgL").Cells(1, Colonne)
    End If
    If CN_dosage_mg_mgNH4_force = True Then
        CN_dosage_mg_mgNH4_f = Feuil52.Range("desinfection_Cl_CN_dosage_mg_mgNH4").Cells(1, Colonne)
    End If
    If CN_NO2_mgL_force = True Then
        CN_NO2_mgL_f = Feuil52.Range("desinfection_Cl_CN_NO2_mgL").Cells(1, Colonne)
    End If
    If CN_dosage_mg_mgNO2_force = True Then
        CN_dosage_mg_mgNO2_f = Feuil52.Range("desinfection_Cl_CN_dosage_mg_mgNO2").Cells(1, Colonne)
    End If
    If CN_residuel_mgL_force = True Then
        CN_residuel_mgL_f = Feuil52.Range("desinfection_Cl_CN_residuel_mgL").Cells(1, Colonne)
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "52.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Private Sub Class_Initialize()

    If pn_EffacementResultatsEnCours = False Then
        Call lecture_choix
        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
            Call lecture_valeurs_forcees
        End If
        Call hypotheses
        Call attribution_valeur_par_defaut
    End If


End Sub


Private Sub Class_Terminate()

    If pn_EffacementResultatsEnCours = False Then
        'Call calcul_consommation_electrique
        Call ecriture_resultats
        Call affich_remarque
    End If

End Sub




