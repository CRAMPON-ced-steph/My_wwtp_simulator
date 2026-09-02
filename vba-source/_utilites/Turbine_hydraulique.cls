Attribute VB_Name = "Turbine_hydraulique"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'GESTION DES CHOIX
Dim choix_debit As Integer
Const DEBIT_TOTAL = 1
Const DEBIT_PARTIEL = 2

'PARAMETRES
Dim rhoEau_kg_m3 As Double    'Masse volumique de l'eau traitée ou non-traitée (kg/m3)
Dim debitEauDim_m3_j As Double    'Debit dimensionnement initial (m3/j)
Dim pourcentageDebit As Double    'Pourcentage du débit total considéré (%)
Dim debitEau_m3_h As Double    'Débit considéré (m3/h)
Dim hauteurTotale_m As Double    'Hauteur totale de chute (m)
Dim perteDeCharge_m As Double    'Perte de charge hydraulique (mce)
Dim rendementTurbine As Double    'Rendement de la turbine
Dim rendementGenerateur As Double    'Rendement du générateur
Dim rendementGlobal_70 As Double    'Rendement global à 70% de charge hydraulique
Dim rendementGlobalMoyen As Double    'Rendement global moyen
Dim dureeFonctionnement As Double    'Durée de fonctionnement (h/an)

'ETAT DES VALEURS A FORCER
Dim debitEauDim_m3_j_forcee As Boolean
Dim hauteurTotale_m_forcee As Boolean
Dim perteDeCharge_m_forcee As Boolean
Dim puissanceNette_forcee As Boolean
Dim rendementGlobalMoyen_forcee As Boolean
Dim dureeFonctionnement_forcee As Boolean

'VALEURS FORCEES
Dim debitEauDim_m3_j_f As Double
Dim hauteurTotale_m_f As Double
Dim perteDeCharge_m_f As Double
Dim puissanceNette_f As Double
Dim rendementGlobalMoyen_f As Double
Dim dureeFonctionnement_f As Double

'ELEMENTS CALCULES
Dim puissanceBrute As Double
Dim puissanceNette As Double
Dim puissanceElectriqueMoyenne As Double
Dim productionAnnuelle As Double
Public electricite_produite_kwh_j As Double

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

Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'Débit
    choix_debit = Feuil65.Range("turbine_hydrau_choix_debit").Cells(1, pi_FiliereConsideree)

    'Pourcentage du débit total considéré (%)
    If choix_debit = DEBIT_PARTIEL Then
        pourcentageDebit = Feuil65.Range("Turbine_hydrau_pourcentageDebit").Cells(1, pi_FiliereConsideree)
    ElseIf choix_debit = DEBIT_TOTAL Then
        pourcentageDebit = 1
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("65.1")
    pn_EffacementResultatsEnCours = True

End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree


    'LECTURE DES CHOIX

    debitEauDim_m3_j_forcee = Feuil65.Range("Turbine_hydrau_debitEauDim_m3_j").Cells(1, pi_FiliereConsideree)    'Debit dimensionnement initial (m3/j)
    hauteurTotale_m_forcee = Feuil65.Range("Turbine_hydrau_hauteurTotale_m").Cells(1, pi_FiliereConsideree)    'Hauteur totale de chute (m)
    perteDeCharge_m_forcee = Feuil65.Range("Turbine_hydrau_perteDeCharge_m").Cells(1, pi_FiliereConsideree)    'Perte de charge hydraulique (mce)
    puissanceNette_forcee = Feuil65.Range("Turbine_hydrau_puissanceNette").Cells(1, pi_FiliereConsideree)    ' puissance nette
    rendementGlobalMoyen_forcee = Feuil65.Range("Turbine_hydrau_rendementGlobalMoyen").Cells(1, pi_FiliereConsideree)    'Rendement global moyen
    dureeFonctionnement_forcee = Feuil65.Range("Turbine_hydrau_dureeFonctionnement").Cells(1, pi_FiliereConsideree)    'Durée de fonctionnement (h/an)


    'AFFECTATION DES VALEURS

    'Debit dimensionnement initial (m3/j)
    If debitEauDim_m3_j_forcee = True Then
        debitEauDim_m3_j_f = Feuil65.Range("Turbine_hydrau_debitEauDim_m3_j").Cells(1, pi_FiliereConsideree) * pd_NOMBRE_HEURE_PAR_JOUR
    End If

    'Hauteur totale de chute (m)
    If hauteurTotale_m_forcee = True Then
        hauteurTotale_m_f = Feuil65.Range("Turbine_hydrau_hauteurTotale_m").Cells(1, pi_FiliereConsideree)
    End If

    'Perte de charge hydraulique (mce)
    If perteDeCharge_m_forcee = True Then
        perteDeCharge_m_f = Feuil65.Range("Turbine_hydrau_perteDeCharge_m").Cells(1, pi_FiliereConsideree)
    End If

    'Puissance nette (kW)
    If puissanceNette_forcee = True Then
        puissanceNette_f = Feuil65.Range("Turbine_hydrau_puissanceNette").Cells(1, pi_FiliereConsideree)
    End If

    'Rendement global moyen
    If rendementGlobalMoyen_forcee = True Then
        rendementGlobalMoyen_f = Feuil65.Range("Turbine_hydrau_rendementGlobalMoyen").Cells(1, pi_FiliereConsideree)
    End If

    'Durée de fonctionnement (h/an)
    If dureeFonctionnement_forcee = True Then
        dureeFonctionnement_f = Feuil65.Range("Turbine_hydrau_dureeFonctionnement").Cells(1, pi_FiliereConsideree)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("65.2")
    pn_EffacementResultatsEnCours = True

End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'Masse volumique de l'eau traitée ou non-traitée (kg/m3)
    rhoEau_kg_m3 = 1000

    'Rendement de la turbine
    rendementTurbine = 0.93

    'Rendement du générateur
    rendementGenerateur = 0.97

    'Rendement global à 70% de charge hydraulique
    rendementGlobal_70 = 0.67

    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("65.3")
    pn_EffacementResultatsEnCours = True

End Sub

Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'Debit dimensionnement initial (m3/j)
    If debitEauDim_m3_j_forcee = True Then
        debitEauDim_m3_j = debitEauDim_m3_j_f
    Else
        debitEauDim_m3_j = Q_nominal * pourcentageDebit
    End If

    'Hauteur totale de chute (m)
    If hauteurTotale_m_forcee = True Then
        hauteurTotale_m = hauteurTotale_m_f
    Else
        hauteurTotale_m = 5
    End If

    'Perte de charge hydraulique (mce)
    If perteDeCharge_m_forcee Then
        perteDeCharge_m = perteDeCharge_m_f
    Else
        perteDeCharge_m = 1
    End If

    'Rendement global moyen
    If rendementGlobalMoyen_forcee = True Then
        rendementGlobalMoyen = rendementGlobalMoyen_f
    Else
        rendementGlobalMoyen = (rendementGlobal_70 + rendementGenerateur * rendementTurbine) / 2
    End If

    'Durée de fonctionnement (h/an)
    If dureeFonctionnement_forcee = True Then
        If dureeFonctionnement_f > 8760 Then    'valeur max : 365 jours * 24h
            dureeFonctionnement = 8760
        Else
            dureeFonctionnement = dureeFonctionnement_f
        End If
    Else
        dureeFonctionnement = 8500
    End If

    Exit Sub

    '    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("65.4")
    pn_EffacementResultatsEnCours = True

End Sub

Sub calcul()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'Calcul du débit concerné
    debitEau_m3_h = Q_nominal * pd_PourcentageChargeReelleDebitVolumique * pourcentageDebit / pd_NOMBRE_HEURE_PAR_JOUR

    'Calcul de la puissance brute de la chute d'eau (kW)
    puissanceBrute = rhoEau_kg_m3 * pd_ACCELERATION_PESANTEUR_m_s2 * (debitEau_m3_h / pd_NOMBRE_SECONDE_PAR_HEURE) * hauteurTotale_m / 1000

    'Calcul de la puissance nette (la hauteur nette correspond à la hauteur totale diminuée de la perte de charge hydraulique (mce) (kW)
    If puissanceNette_forcee = True Then
        puissanceNette = puissanceNette_f
    Else
        puissanceNette = rhoEau_kg_m3 * pd_ACCELERATION_PESANTEUR_m_s2 * (debitEau_m3_h / pd_NOMBRE_SECONDE_PAR_HEURE) * (hauteurTotale_m - perteDeCharge_m) / 1000
    End If

    'Calcul de la puissance electrique moyenne (kW)
    puissanceElectriqueMoyenne = puissanceNette * rendementGlobalMoyen

    'Calcul de la production d'électricité
    productionAnnuelle = dureeFonctionnement * puissanceElectriqueMoyenne
    electricite_produite_kwh_j = productionAnnuelle / pd_NOMBRE_JOUR_PAR_AN

    Exit Sub
    '    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("65.5")
    pn_EffacementResultatsEnCours = True

End Sub

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil65.Range("Turbine_hydrau_debitEauDim_m3_j").Cells(1, Colonne) = debitEauDim_m3_j / pd_NOMBRE_HEURE_PAR_JOUR
    Feuil65.Range("Turbine_hydrau_debitEau_m3_h").Cells(1, Colonne) = debitEau_m3_h * pd_NOMBRE_HEURE_PAR_JOUR
    Feuil65.Range("Turbine_hydrau_hauteurTotale_m").Cells(1, Colonne) = hauteurTotale_m
    Feuil65.Range("Turbine_hydrau_perteDeCharge_m").Cells(1, Colonne) = perteDeCharge_m
    Feuil65.Range("Turbine_hydrau_rendementGlobalMoyen").Cells(1, Colonne) = rendementGlobalMoyen
    Feuil65.Range("Turbine_hydrau_dureeFonctionnement").Cells(1, Colonne) = dureeFonctionnement
    Feuil65.Range("Turbine_hydrau_puissanceBrute").Cells(1, Colonne) = puissanceBrute
    Feuil65.Range("Turbine_hydrau_puissanceNette").Cells(1, Colonne) = puissanceNette
    Feuil65.Range("Turbine_hydrau_puissanceElectriqueMoyenne").Cells(1, Colonne) = puissanceElectriqueMoyenne
    Feuil65.Range("Turbine_hydrau_productionAnnuelle").Cells(1, Colonne) = productionAnnuelle
    Feuil65.Range("Turbine_hydrau_electricite_produite_kwh_j").Cells(1, Colonne) = electricite_produite_kwh_j

    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("65.6")
    pn_EffacementResultatsEnCours = True

End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil65.Range("Turbine_hydrau_debitEauDim_m3_j").Cells(1, Colonne).ClearContents
    Feuil65.Range("Turbine_hydrau_debitEau_m3_h").Cells(1, Colonne).ClearContents
    Feuil65.Range("Turbine_hydrau_hauteurTotale_m").Cells(1, Colonne).ClearContents
    Feuil65.Range("Turbine_hydrau_perteDeCharge_m").Cells(1, Colonne).ClearContents
    Feuil65.Range("Turbine_hydrau_rendementGlobalMoyen").Cells(1, Colonne).ClearContents
    Feuil65.Range("Turbine_hydrau_dureeFonctionnement").Cells(1, Colonne).ClearContents
    Feuil65.Range("Turbine_hydrau_puissanceBrute").Cells(1, Colonne).ClearContents
    Feuil65.Range("Turbine_hydrau_puissanceNette").Cells(1, Colonne).ClearContents
    Feuil65.Range("Turbine_hydrau_puissanceElectriqueMoyenne").Cells(1, Colonne).ClearContents
    Feuil65.Range("Turbine_hydrau_productionAnnuelle").Cells(1, Colonne).ClearContents
    Feuil65.Range("Turbine_hydrau_electricite_produite_kwh_j").Cells(1, Colonne).ClearContents

    Exit Sub
    '    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("65.7")
    pn_EffacementResultatsEnCours = True

End Sub

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    '
    '
    '    Dim nb_remarque As Integer
    '    Dim i As Integer
    '
    '    nb_remarque = 1  'penser à modifier la capacité du vecteur "message_erreur"
    '
    '    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil9.Range("dec_simple_erreur_nom")
    '
    '    For i = 1 To nb_remarque
    '        If remarque(i) <> "" Then
    '            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(i)
    '        End If
    '    Next i
    '
    '    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf
    '
    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("65.8")
    pn_EffacementResultatsEnCours = True

End Sub

Private Sub Class_Terminate()

    If pn_EffacementResultatsEnCours = False Then
        Call ecriture_resultats
        Call affich_remarque
    End If

End Sub





