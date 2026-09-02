Attribute VB_Name = "z_Desodorisation_biologique"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Const nb_type_couverture_step = 2
Private Const couverture_partielle = 1
Private Const couverture_totale = 2

Private Const soufflage_oui = 1

'variables de choix
Dim choix_couverture_step As Integer
Dim choix_soufflage_air As Integer

'variables apparaissant à l'utilisateur
Public air_vicie_Q_Nm3h As Double
Dim air_vicie_T As Double
Dim ventilation_tps_fonctionnement As Double
Dim ventilation_P_reseau As Double
Dim ventilation_rdt_global As Double
Dim pompe_lavage_rate_Leau_m3air As Double
Dim pompe_lavage_HMT As Double
Dim pompe_lavage_rdt As Double
Dim soufflage_P_reseau As Double
Dim soufflage_Tair As Double
Dim soufflage_rdt_global As Double
Public electricite As Double

'variables indiquant le choix de forcer les valeurs
Dim air_vicie_Q_Nm3h_force As Boolean
Dim air_vicie_T_force As Boolean
Dim ventilation_tps_fonctionnement_force As Boolean
Dim ventilation_P_reseau_force As Boolean
Dim ventilation_rdt_global_force As Boolean
Dim pompe_lavage_rate_Leau_m3air_force As Boolean
Dim pompe_lavage_HMT_force As Boolean
Dim pompe_lavage_rdt_force As Boolean
Dim soufflage_P_reseau_force As Boolean
Dim soufflage_Tair_force As Boolean
Dim soufflage_rdt_global_force As Boolean

'valeurs des variables forcées
Dim air_vicie_Q_Nm3h_f As Double
Dim air_vicie_T_f As Double
Dim ventilation_tps_fonctionnement_f As Double
Dim ventilation_P_reseau_f As Double
Dim ventilation_rdt_global_f As Double
Dim pompe_lavage_rate_Leau_m3air_f As Double
Dim pompe_lavage_HMT_f As Double
Dim pompe_lavage_rdt_f As Double
Dim soufflage_P_reseau_f As Double
Dim soufflage_rdt_global_f As Double
Dim soufflage_Tair_f As Double

'hypothèses
Dim rate_air_vicie_Qnominal As Double
Dim ventilateur_rdt_machine As Double
Dim ventilateur_rdt_moteur As Double

Dim remarque(3) As String

Sub calcul()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String











    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "43.12"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub calcul_consommation_electrique()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim electricite_ventilation As Double
    Dim electricite_soufflage As Double
    Dim electricite_lavage As Double

    Dim Q_lavage_unitaire As Double
    Dim rendement_pompe As Double
    Dim rendement_moteur As Double
    Dim puissance_pompe As Double


    'CONSO ELECTRIQUE
    electricite = 0

    'Ventilation
    electricite_ventilation = (air_vicie_Q_Nm3h * ((air_vicie_T + pd_CONVERSION_K_°C) / pd_CONVERSION_K_°C) / pd_NOMBRE_SECONDE_PAR_HEURE) * ventilation_P_reseau / ventilation_rdt_global / 1000
    electricite_ventilation = ventilation_tps_fonctionnement * electricite_ventilation
    electricite = electricite + electricite_ventilation

    'Soufflage
    If choix_soufflage_air = soufflage_oui Then
        electricite_soufflage = (air_vicie_Q_Nm3h * ((soufflage_Tair + pd_CONVERSION_K_°C) / pd_CONVERSION_K_°C) / pd_NOMBRE_SECONDE_PAR_HEURE) * soufflage_P_reseau / soufflage_rdt_global / 1000
        electricite_soufflage = ventilation_tps_fonctionnement * electricite_soufflage
    Else
        electricite_soufflage = 0
    End If
    electricite = electricite + electricite_soufflage

    'Eaux de lavage
    Q_lavage_unitaire = pompe_lavage_rate_Leau_m3air / 1000 * air_vicie_Q_Nm3h
    'incoherence = Q_lavage_unitaire
    'If Q_lavage_unitaire_force = True Then
    '    Q_lavage_unitaire = Q_lavage_unitaire_f
    'End If
    'incoherence = (incoherence - Q_lavage_unitaire) / Q_lavage_unitaire
    'If incoherence > critere_incoherence Then
    '    remarque(2) = Feuil25.Range("epaississement_erreur_pompage_alimentation").Cells(1, 1)
    'End If
    If pompe_lavage_rdt_force = True Then
        pompe_lavage_rdt = pompe_lavage_rdt_f
    Else
        rendement_pompe = 0.8    'hypothèse
        puissance_pompe = Q_lavage_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * pompe_lavage_HMT / rendement_pompe
        If puissance_pompe < 60 Then
            rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
        Else
            rendement_moteur = 0.961
        End If
        pompe_lavage_rdt = rendement_pompe * rendement_moteur
    End If
    electricite_lavage = Q_lavage_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * pompe_lavage_HMT / pompe_lavage_rdt * ventilation_tps_fonctionnement
    electricite = electricite + electricite_lavage




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "43.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub


Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil43.Range("desodo_bio_air_vicie_Q_Nm3h").Cells(1, Colonne) = air_vicie_Q_Nm3h
    Feuil43.Range("desodo_bio_air_vicie_T").Cells(1, Colonne) = air_vicie_T
    Feuil43.Range("desodo_bio_ventilation_tps_fonctionnement").Cells(1, Colonne) = ventilation_tps_fonctionnement
    Feuil43.Range("desodo_bio_ventilation_P_reseau").Cells(1, Colonne) = ventilation_P_reseau
    Feuil43.Range("desodo_bio_ventilation_rdt_global").Cells(1, Colonne) = ventilation_rdt_global
    Feuil43.Range("desodo_bio_pompe_lavage_rate_Leau_m3air").Cells(1, Colonne) = pompe_lavage_rate_Leau_m3air
    Feuil43.Range("desodo_bio_pompe_lavage_HMT").Cells(1, Colonne) = pompe_lavage_HMT
    Feuil43.Range("desodo_bio_pompe_lavage_rdt").Cells(1, Colonne) = pompe_lavage_rdt
    Feuil43.Range("desodo_bio_soufflage_P_reseau").Cells(1, Colonne) = soufflage_P_reseau
    Feuil43.Range("desodo_bio_soufflage_Tair").Cells(1, Colonne) = soufflage_Tair
    Feuil43.Range("desodo_bio_soufflage_rdt_global").Cells(1, Colonne) = soufflage_rdt_global
    Feuil43.Range("desodo_bio_electricite").Cells(1, Colonne) = electricite



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "43.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil43.Range("desodo_bio_air_vicie_Q_Nm3h").Cells(1, Colonne) = ""    ' air_vicie_Q_Nm3h
    Feuil43.Range("desodo_bio_air_vicie_T").Cells(1, Colonne) = ""    ' air_vicie_T
    Feuil43.Range("desodo_bio_ventilation_tps_fonctionnement").Cells(1, Colonne) = ""    ' ventilation_tps_fonctionnement
    Feuil43.Range("desodo_bio_ventilation_P_reseau").Cells(1, Colonne) = ""    ' ventilation_P_reseau
    Feuil43.Range("desodo_bio_ventilation_rdt_global").Cells(1, Colonne) = ""   ' ventilation_rdt_global
    Feuil43.Range("desodo_bio_pompe_lavage_rate_Leau_m3air").Cells(1, Colonne) = ""   ' pompe_lavage_rate_Leau_m3air
    Feuil43.Range("desodo_bio_pompe_lavage_HMT").Cells(1, Colonne) = ""    ' pompe_lavage_HMT
    Feuil43.Range("desodo_bio_pompe_lavage_rdt").Cells(1, Colonne) = ""    ' pompe_lavage_rdt
    Feuil43.Range("desodo_bio_soufflage_P_reseau").Cells(1, Colonne) = ""    ' soufflage_P_reseau
    Feuil43.Range("desodo_bio_soufflage_Tair").Cells(1, Colonne) = ""    ' soufflage_Tair
    Feuil43.Range("desodo_bio_soufflage_rdt_global").Cells(1, Colonne) = ""    ' soufflage_rdt_global
    Feuil43.Range("desodo_bio_electricite").Cells(1, Colonne) = ""    ' electricite



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "43.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 3  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil43.Range("desodo_bio_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "43.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    'débit d'air à traiter   (en m3 air / m3 eau traitée au nominal)
    If choix_couverture_step = couverture_partielle Then
        rate_air_vicie_Qnominal = 13
    Else
        rate_air_vicie_Qnominal = 42
    End If

    'rendements du ventilateur
    ventilateur_rdt_machine = 0.7
    ventilateur_rdt_moteur = 0.8


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "43.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer


    'temps de fonctionnement de la désodo   en h/j
    If ventilation_tps_fonctionnement_force = True Then
        ventilation_tps_fonctionnement = ventilation_tps_fonctionnement_f
    Else
        ventilation_tps_fonctionnement = pd_NOMBRE_HEURE_PAR_JOUR
    End If

    'débit d'air vicié    Nm3/h
    If air_vicie_Q_Nm3h_force = True Then
        air_vicie_Q_Nm3h = air_vicie_Q_Nm3h_f
    Else
        air_vicie_Q_Nm3h = rate_air_vicie_Qnominal * Q_nominal / ventilation_tps_fonctionnement
    End If

    'Température de l'air à traiter
    If air_vicie_T_force = True Then
        air_vicie_T = air_vicie_T_f
    Else
        air_vicie_T = T_air_aspire
    End If

    'Pression du réseau   en Pa
    If ventilation_P_reseau_force = True Then
        ventilation_P_reseau = ventilation_P_reseau_f
    Else
        ventilation_P_reseau = 2500
    End If

    'Rendement global du ventilateur
    If ventilation_rdt_global_force = True Then
        ventilation_rdt_global = ventilation_rdt_global_f
    Else
        ventilation_rdt_global = ventilateur_rdt_machine * ventilateur_rdt_moteur     'mécanique * moteur
    End If

    'ratio eau de lavage / air  (L/m3 et par tour)
    If pompe_lavage_rate_Leau_m3air_force = True Then
        pompe_lavage_rate_Leau_m3air = pompe_lavage_rate_Leau_m3air_f
    Else
        pompe_lavage_rate_Leau_m3air = 0.1
    End If

    'HMT pompes de lavage    en mCE
    If pompe_lavage_HMT_force = True Then
        pompe_lavage_HMT = pompe_lavage_HMT_f
    Else
        pompe_lavage_HMT = 20
    End If

    'rendement des pompes de lavage    %
    If pompe_lavage_rdt_force = True Then
        pompe_lavage_rdt = pompe_lavage_rdt_f
    Else
        pompe_lavage_rdt = 0.8    'uniquement rendement machine
    End If

    'pression réseau de soufflage   en Pa
    If soufflage_P_reseau_force = True Then
        soufflage_P_reseau = soufflage_P_reseau_f
    Else
        soufflage_P_reseau = 300
    End If

    'Température de l'air avant soufflage (°C)
    If soufflage_Tair_force = True Then
        soufflage_Tair = soufflage_Tair_f
    Else
        soufflage_Tair = air_vicie_T
    End If

    'rendement ventilateur de soufflage
    If soufflage_rdt_global_force = True Then
        soufflage_rdt_global = soufflage_rdt_global_f
    Else
        soufflage_rdt_global = ventilateur_rdt_machine * ventilateur_rdt_moteur     'mécanique * moteur
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "43.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub



Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    choix_couverture_step = Feuil43.Range("desodo_bio_choix_couvertureSTEP").Cells(1, pi_FiliereConsideree)
    choix_soufflage_air = Feuil43.Range("desodo_bio_choix_soufflage").Cells(1, pi_FiliereConsideree)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "43.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    air_vicie_Q_Nm3h_force = Feuil43.Range("desodo_bio_air_vicie_Q_Nm3h").Cells(1, pi_FiliereConsideree)
    air_vicie_T_force = Feuil43.Range("desodo_bio_air_vicie_T").Cells(1, pi_FiliereConsideree)
    ventilation_tps_fonctionnement_force = Feuil43.Range("desodo_bio_ventilation_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    ventilation_P_reseau_force = Feuil43.Range("desodo_bio_ventilation_P_reseau").Cells(1, pi_FiliereConsideree)
    ventilation_rdt_global_force = Feuil43.Range("desodo_bio_ventilation_rdt_global").Cells(1, pi_FiliereConsideree)
    pompe_lavage_rate_Leau_m3air_force = Feuil43.Range("desodo_bio_pompe_lavage_rate_Leau_m3air").Cells(1, pi_FiliereConsideree)
    pompe_lavage_HMT_force = Feuil43.Range("desodo_bio_pompe_lavage_HMT").Cells(1, pi_FiliereConsideree)
    pompe_lavage_rdt_force = Feuil43.Range("desodo_bio_pompe_lavage_rdt").Cells(1, pi_FiliereConsideree)
    soufflage_P_reseau_force = Feuil43.Range("desodo_bio_soufflage_P_reseau").Cells(1, pi_FiliereConsideree)
    soufflage_Tair_force = Feuil43.Range("desodo_bio_soufflage_Tair").Cells(1, pi_FiliereConsideree)
    soufflage_rdt_global_force = Feuil43.Range("desodo_bio_soufflage_rdt_global").Cells(1, pi_FiliereConsideree)


    If air_vicie_Q_Nm3h_force = True Then
        air_vicie_Q_Nm3h_f = Feuil43.Range("desodo_bio_air_vicie_Q_Nm3h").Cells(1, pi_FiliereConsideree)
    End If
    If air_vicie_T_force = True Then
        air_vicie_T_f = Feuil43.Range("desodo_bio_air_vicie_T").Cells(1, pi_FiliereConsideree)
    End If
    If ventilation_tps_fonctionnement_force = True Then
        ventilation_tps_fonctionnement_f = Feuil43.Range("desodo_bio_ventilation_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    End If
    If ventilation_P_reseau_force = True Then
        ventilation_P_reseau_f = Feuil43.Range("desodo_bio_ventilation_P_reseau").Cells(1, pi_FiliereConsideree)
    End If
    If ventilation_rdt_global_force = True Then
        ventilation_rdt_global_f = Feuil43.Range("desodo_bio_ventilation_rdt_global").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_lavage_rate_Leau_m3air_force = True Then
        pompe_lavage_rate_Leau_m3air_f = Feuil43.Range("desodo_bio_pompe_lavage_rate_Leau_m3air").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_lavage_HMT_force = True Then
        pompe_lavage_HMT_f = Feuil43.Range("desodo_bio_pompe_lavage_HMT").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_lavage_rdt_force = True Then
        pompe_lavage_rdt_f = Feuil43.Range("desodo_bio_pompe_lavage_rdt").Cells(1, pi_FiliereConsideree)
    End If
    If soufflage_P_reseau_force = True Then
        soufflage_P_reseau_f = Feuil43.Range("desodo_bio_soufflage_P_reseau").Cells(1, pi_FiliereConsideree)
    End If
    If soufflage_Tair_force = True Then
        soufflage_Tair_f = Feuil43.Range("desodo_bio_soufflage_Tair").Cells(1, pi_FiliereConsideree)
    End If
    If soufflage_rdt_global_force = True Then
        soufflage_rdt_global_f = Feuil43.Range("desodo_bio_soufflage_rdt_global").Cells(1, pi_FiliereConsideree)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "43.2"

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





