Option Explicit

'constantes pour les choix possibilité d'adapter la puissance des lampes
Private Const manuel = 1
Private Const automatique = 2
Dim choix_possibilite_lampes As Integer

'constantes pour les choix
Private Const basse_pression = 1
Private Const moyenne_pression = 2
Dim choix_technologie As Integer

'entrée
Dim Q_retour As Double
Dim debit As Double
Dim debit_retours As Double
Dim pourcentage_charge_hydraulic As Double
Dim fact_hydraulic_pluie As Double

'guide defaut

'hypothese
Dim nombre_bancs_prevu As Double
Dim nombre_bancs_prevu_ent As Integer
Dim dose_elec As Double
Dim temps_operation_jour As Double
Dim factor_reel As Double
Dim electricite_specifique As Double

'calcul
Dim debit_nominal As Double
Dim debit_reel_jour As Double
Dim debit_reel_heure As Double
Dim puissance_nominal As Double
Dim puissance_par_banc As Double
Dim puissance_elec_reel As Double
Dim puissance_ratio As Double
Dim nombre_bancs_reel As Double
Dim nombre_bancs_ent As Integer
Public electricite As Double

'verifie forcée
Dim dose_elec_force As Boolean
Dim debit_nominal_force As Boolean
Dim debit_reel_jour_force As Boolean
Dim nombre_bancs_prevu_force As Boolean
Dim nombre_bancs_reel_force As Boolean
'si necessaire
Dim temps_operation_jour_force As Boolean

'valeur forcée
Dim dose_elec_f As Double
Dim debit_nominal_f As Double
Dim debit_reel_jour_f As Double
Dim nombre_bancs_prevu_f As Double
Dim nombre_bancs_reel_f As Double
'si necessaire
Dim temps_operation_jour_f As Double

Dim remarque(2) As String



Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'lecture des choix internes
    choix_possibilite_lampes = Feuil24.Range("UV_choix_possibilite_lampes").Cells(1, pi_FiliereConsideree)
    choix_technologie = Feuil24.Range("UV_choix_puissance_tech").Cells(1, pi_FiliereConsideree)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "24.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    factor_reel = 0.6  'il est possible de réguler l'intensité des bancs entre 60% et 100% --> intérêt si 1 seul banc
    temps_operation_jour = 24    'h

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "24.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    
    If choix_technologie = basse_pression Then
        dose_elec = 30  'Wh/m3  dépend de la qualité d'eau, valeur estimée avec J. PANNEJON
        puissance_par_banc = 70       'kW  J. PANNEJON: un banc avec 10 modules de 9 paires de lampes (au total 180 lampes de 400 W), est modifiée si l'utilisateur rentre le nombre de bancs au nominal
    Else
        dose_elec = 70  'Wh/m3  dépend de la qualité d'eau, valeur estimée avec J. PANNEJON / plutôt eau potable, est modifiée si l'utilisateur rentre le nombre de bancs au nominal
        puissance_par_banc = 40       'kW  J. PANNEJON: un banc avec 6 lampes de 6 kW
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "24.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim column As Integer

    column = pi_FiliereConsideree

    dose_elec_force = Feuil24.Range("uv_dose_elec").Cells(1, column)
    debit_nominal_force = Feuil24.Range("uv_debit_nominal").Cells(1, column)
    debit_reel_jour_force = Feuil24.Range("uv_debit_reel_jour").Cells(1, column)
    nombre_bancs_prevu_force = Feuil24.Range("uv_nombre_bancs_prevu").Cells(1, column)
    nombre_bancs_reel_force = Feuil24.Range("uv_nombre_bancs_reel").Cells(1, column)

    If dose_elec_force = True Then
        dose_elec_f = Feuil24.Range("uv_dose_elec").Cells(1, column)
    End If

    If debit_nominal_force = True Then
        debit_nominal_f = Feuil24.Range("uv_debit_nominal").Cells(1, column)
    End If

    If debit_reel_jour_force = True Then
        debit_reel_jour_f = Feuil24.Range("uv_debit_reel_jour").Cells(1, column)
    End If

    If nombre_bancs_prevu_force = True Then
        nombre_bancs_prevu_f = Feuil24.Range("uv_nombre_bancs_prevu").Cells(1, column)
    End If

    If nombre_bancs_reel_force = True Then
        nombre_bancs_reel_f = Feuil24.Range("uv_nombre_bancs_reel").Cells(1, column)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "24.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub dimensionnement(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    fact_hydraulic_pluie = pd_CoefficientPointeHydrauliqueTempsPluie
    debit_retours = Q - Q_nominal
    debit = Q_nominal
    pourcentage_charge_hydraulic = pd_PourcentageChargeReelleDebitVolumique

    If dose_elec_force = True Then
        dose_elec = dose_elec_f
    Else
        dose_elec = dose_elec
    End If

    If debit_nominal_force = True Then
        debit_nominal = debit_nominal_f
    Else
        debit_nominal = (debit * fact_hydraulic_pluie + debit_retours) / temps_operation_jour
    End If

    puissance_nominal = dose_elec / 1000 * debit_nominal

    If nombre_bancs_prevu_force = True Then
        nombre_bancs_prevu = nombre_bancs_prevu_f
        puissance_par_banc = puissance_nominal / nombre_bancs_prevu
    Else
        nombre_bancs_prevu = puissance_nominal / puissance_par_banc
    End If

    nombre_bancs_prevu_ent = nombre_bancs_prevu

    If nombre_bancs_prevu_ent = nombre_bancs_prevu Then
        nombre_bancs_prevu = nombre_bancs_prevu_ent
    Else
        nombre_bancs_prevu = nombre_bancs_prevu_ent + 1
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "24.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub fonctionnement_reel(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    If debit_reel_jour_force = True Then
        debit_reel_jour = debit_reel_jour_f
        debit_reel_heure = debit_reel_jour / temps_operation_jour
    Else
        debit_reel_heure = (debit * pourcentage_charge_hydraulic + debit_retours) / temps_operation_jour
        debit_reel_jour = (debit * pourcentage_charge_hydraulic + debit_retours)
    End If

    If nombre_bancs_reel_force = True Then
        nombre_bancs_reel = nombre_bancs_reel_f
        puissance_ratio = nombre_bancs_reel
    Else
        puissance_elec_reel = dose_elec / 1000 * debit_reel_heure
        puissance_ratio = puissance_elec_reel / puissance_par_banc
        nombre_bancs_ent = puissance_ratio
        If nombre_bancs_ent = puissance_ratio Then
            nombre_bancs_reel = nombre_bancs_ent
        Else
            nombre_bancs_reel = nombre_bancs_ent + 1
        End If
    End If


    'VERIFICATION DE LA NOMBRE DES BANCS
    If nombre_bancs_reel > nombre_bancs_prevu Then
        remarque(1) = Feuil24.Range("uv_erreur").Cells(1, 1)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "24.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub calcul_consommation_electrique()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim electricite_specifique_1 As Double  'calculé avec no. banc réel et factor réel (= 0.6)
    Dim electricite_specifique_2 As Double  'calculé avec ratio

    If choix_possibilite_lampes = manuel Then
        electricite_specifique = nombre_bancs_reel * temps_operation_jour * puissance_par_banc
    Else
        'DEBUT MODIF DCA 13/03/2014
        If nombre_bancs_reel = 1 Then 'on peut réguler de 60% à 100%
            If puissance_ratio < factor_reel Then
                electricite_specifique = factor_reel * nombre_bancs_reel * temps_operation_jour * puissance_par_banc
            Else
                electricite_specifique = dose_elec / 1000 * debit_reel_heure
            End If
        ElseIf nombre_bancs_reel = 1 Then  'on peut réguler de 60% à 100% avec 1 banc puis de 120% à 200% avec 2 bancs
            If puissance_ratio < nombre_bancs_reel * factor_reel And puissance_ratio > (nombre_bancs_reel - 1) Then
                electricite_specifique = factor_reel * nombre_bancs_reel * temps_operation_jour * puissance_par_banc
            Else
                electricite_specifique = dose_elec / 1000 * debit_reel_heure
            End If
        Else
            electricite_specifique = dose_elec / 1000 * debit_reel_heure
        End If
'        electricite_specifique_1 = factor_reel * nombre_bancs_reel * temps_operation_jour * puissance_par_banc
'        electricite_specifique_2 = puissance_ratio * temps_operation_jour * puissance_par_banc
'        If electricite_specifique_1 > electricite_specifique_2 Then
'            electricite_specifique = electricite_specifique_1
'        Else
'            electricite_specifique = electricite_specifique_2
'        End If
        'FIN MODIF DCA 13/03/2014
    End If

    electricite = electricite_specifique

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "24.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


    Feuil24.Range("uv_debit_nominal").Cells(1, Colonne) = debit_nominal
    Feuil24.Range("uv_nombre_bancs_prevu").Cells(1, Colonne) = nombre_bancs_prevu
    Feuil24.Range("uv_dose_elec").Cells(1, Colonne) = dose_elec
    Feuil24.Range("uv_debit_reel_jour").Cells(1, Colonne) = debit_reel_jour
    Feuil24.Range("uv_nombre_bancs_reel").Cells(1, Colonne) = nombre_bancs_reel
    Feuil24.Range("uv_electricite_specifique").Cells(1, Colonne) = electricite_specifique


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "24.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil24.Range("uv_debit_nominal").Cells(1, Colonne) = ""
    Feuil24.Range("uv_nombre_bancs_prevu").Cells(1, Colonne) = ""
    Feuil24.Range("uv_dose_elec").Cells(1, Colonne) = ""
    Feuil24.Range("uv_debit_reel_jour").Cells(1, Colonne) = ""
    Feuil24.Range("uv_nombre_bancs_reel").Cells(1, Colonne) = ""
    Feuil24.Range("uv_electricite_specifique").Cells(1, Colonne) = ""

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "24.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 1  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil24.Range("uv_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "24.8"

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
        Call ecriture_resultats
        Call affich_remarque
    End If



End Sub







