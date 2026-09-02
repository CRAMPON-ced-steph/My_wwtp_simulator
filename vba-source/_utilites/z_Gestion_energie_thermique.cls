Attribute VB_Name = "z_Gestion_energie_thermique"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'constantes technologies
Private Const nb_technologies = 3
Private Const chaudiere = 1
Private Const micro_turbine = 2
Private Const moteur = 3

'critère d'égalité
Private Const critere_egalite = 1
Private Const zero_force As Double = 0.00001

Private Const energie_vapeur_global = 3

'constantes objectifs
Private Const objectif_autosuffisance_thermique = 1
Private Const objectif_autosuffisance_electrique = 2

'constantes cogénération
Private Const cogeneration_oui = 1
Private Const cogeneration_non = 2


'constantes combustible
Private Const combustible_utilisation_chaudiere = 1

'choix de l'utilisateur
Dim choix_objectif As Integer
Dim choix_cogeneration As Integer
Dim choix_combustible As Integer
Dim choix_combustible_utilisation As Integer

'variables à échanger avec l'utilisateur
Dim besoins_thermiques(nb_niveaux_energie) As Double
Dim pression_vapeur(nb_vapeur_max) As Double
Dim besoins_vapeur_kgj(nb_vapeur_max) As Double
Dim disponibilites_thermiques(nb_niveaux_energie) As Double
Dim biogaz_caracteristiques(nb_biogaz_caracteristiques) As Double
Dim biogaz_chaudiere_rdt As Double
Dim biogaz_chaudiere_Tch As Double
Dim biogaz_chaudiere_Tfr As Double
Dim biogaz_CHP_type As Integer
Dim biogaz_CHP_rdt_elec As Double
Dim biogaz_CHP_rdt_eau_BT As Double
Dim biogaz_CHP_BT_Tch As Double
Dim biogaz_CHP_BT_Tfr As Double
Dim biogaz_CHP_rdt_eau_HT As Double
Dim biogaz_CHP_HT_Tch As Double
Dim biogaz_CHP_HT_Tfr As Double
Dim biogaz_CHP_rdt_vapeur As Double
Dim biogaz_CHP_fumees_Tch As Double
Dim biogaz_CHP_fumees_Tfr As Double
Dim biogaz_CHP_Pw As Double
Dim biogaz_CHP_Pw_reel As Double
Dim biogaz_revendu_rate As Double
Dim biogaz_chaudiere_rate As Double
Dim biogaz_CHP_rate As Double
Dim biogaz_sechage_rate As Double
Dim biogaz_torchere_rate As Double
Dim combustible_chaudiere_rdt As Double
Dim combustible_chaudiere_Tch As Double
Dim combustible_chaudiere_Tfr As Double
Dim combustible_flux As Double
Dim combustible_sechage As Double
Dim energie_fatale_kWhj(energie_vapeur_global) As Double
Dim vapeur_fatale_kgj As Double
Dim autosuffisance_thermique As Double
Public electricite_biogaz As Double

'valeurs forcées dans le tableur
Dim biogaz_chaudiere_rdt_force As Boolean
Dim biogaz_chaudiere_Tch_force As Boolean
Dim biogaz_chaudiere_Tfr_force As Boolean
Dim biogaz_CHP_rdt_elec_force As Boolean
Dim biogaz_CHP_rdt_eau_BT_force As Boolean
Dim biogaz_CHP_BT_Tch_force As Boolean
Dim biogaz_CHP_BT_Tfr_force As Boolean
Dim biogaz_CHP_rdt_eau_HT_force As Boolean
Dim biogaz_CHP_HT_Tch_force As Boolean
Dim biogaz_CHP_HT_Tfr_force As Boolean
Dim biogaz_CHP_rdt_vapeur_force As Boolean
Dim biogaz_CHP_fumees_Tch_force As Boolean
Dim biogaz_CHP_fumees_Tfr_force As Boolean
Dim biogaz_CHP_Pw_force As Boolean
Dim biogaz_revendu_rate_force As Boolean
Dim biogaz_chaudiere_rate_force As Boolean
Dim biogaz_CHP_rate_force As Boolean
Dim biogaz_torchere_rate_force As Boolean
Dim biogaz_sechage_rate_force As Boolean
Dim combustible_chaudiere_rdt_force As Boolean
Dim combustible_chaudiere_Tch_force As Boolean
Dim combustible_chaudiere_Tfr_force As Boolean

'valeurs forcées dans le tableur
Dim biogaz_chaudiere_rdt_f As Double
Dim biogaz_chaudiere_Tch_f As Double
Dim biogaz_chaudiere_Tfr_f As Double
Dim biogaz_CHP_rdt_elec_f As Double
Dim biogaz_CHP_rdt_eau_BT_f As Double
Dim biogaz_CHP_BT_Tch_f As Double
Dim biogaz_CHP_BT_Tfr_f As Double
Dim biogaz_CHP_rdt_eau_HT_f As Double
Dim biogaz_CHP_HT_Tch_f As Double
Dim biogaz_CHP_HT_Tfr_f As Double
Dim biogaz_CHP_rdt_vapeur_f As Double
Dim biogaz_CHP_fumees_Tch_f As Double
Dim biogaz_CHP_fumees_Tfr_f As Double
Dim biogaz_CHP_Pw_f As Double
Dim biogaz_revendu_rate_f As Double
Dim biogaz_chaudiere_rate_f As Double
Dim biogaz_CHP_rate_f As Double
Dim biogaz_torchere_rate_f As Double
Dim biogaz_sechage_rate_f As Double
Dim combustible_chaudiere_rdt_f As Double
Dim combustible_chaudiere_Tch_f As Double
Dim combustible_chaudiere_Tfr_f As Double

'hypothèses
Dim PCI_CH4 As Double
Dim PCI(combustibles_nb) As Double
Dim a_Tvap_sat(5) As Double
Dim a_Hvap_sat(1) As Double
Dim rdt_thermique(nb_technologies, energie_vapeur_global) As Double
Dim rdt_electrique(nb_technologies) As Double
Dim limit_min_moteur As Double
Dim biogaz_CHP_rate_maxi As Double
Dim biogaz_chaudiere_rate_mini As Double
Dim biogaz_torchere_rate_mini As Double


Dim remarque(6) As String
Public gaz_naturel_kWhPCIj As Double
Public fioul_kWhPCIj As Double
Public pd_BesoinsCombustibleExterneIncineration_kWhj As Double

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 6    'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil41.Range("thermique_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub calcul_biogaz_repartition(ByRef CHP_maxi, ByRef besoins_th, ByRef dispo, ByRef energie_primaire_biogaz)

    If version_Developpement = False Then On Error GoTo erreur_non_geree 'a
    Dim NumeroErreurNonGeree As String


    'On souhaite trouver l'optimum de biogaz à envoyer en CHP  (autosuffisance thermique maxi avec cogénération)

    'pour la précision des calculs
    Const precision = 0.00001
    Const nb_points = 2

    'constantes pour le bilan
    Const bilan_positif = 1
    Const bilan_negatif = -1

    Dim iter_max As Integer
    Dim iter As Integer
    Dim I As Integer
    Dim j As Integer
    Dim energie_primaire_CHP As Double

    Dim x_CHP(nb_points) As Double
    Dim x_chaudiere As Double
    Dim intermediaire_besoins(energie_vapeur_global) As Double
    Dim intermediaire_dispo(energie_vapeur_global) As Double
    Dim bilan(nb_points) As Integer
    Dim moteur_possible As Boolean
    Dim CHP_mini As Double
    Dim rdt_th_coge(energie_vapeur_global) As Double

    'test limite pour savoir si on peut mettre un moteur ou non
    'utile car si l'énergie primaire limite pour le moteur est dans l'intervalle --> courbe non continue et non strictement monotone sur l'intervalle --> dichotomie impossible
    biogaz_CHP_rate = pd_NOMBRE_HEURE_PAR_JOUR * limit_min_moteur / energie_primaire_biogaz
    If biogaz_CHP_rate < CHP_maxi Then
        'on initialise les besoins
        For j = 1 To energie_vapeur_global
            intermediaire_besoins(j) = besoins_th(j)
            intermediaire_dispo(j) = 0
        Next j
        'on calcule la répartition
        x_chaudiere = (1 - biogaz_sechage_rate) - biogaz_torchere_rate - biogaz_CHP_rate
        'on calcule l'énergie de la chaudière
        intermediaire_dispo(energie_vapeur_global) = x_chaudiere * energie_primaire_biogaz * biogaz_chaudiere_rdt
        'on définit le type de cogé
        energie_primaire_CHP = energie_primaire_biogaz * biogaz_CHP_rate
        Call definition_type_CHP(energie_primaire_CHP)
        rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
        rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
        rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
        'on calcule les besoins après
        Call calcul_dispo_besoins_CHP(intermediaire_besoins, intermediaire_dispo, energie_primaire_CHP, rdt_th_coge)
        'on va regarder si on couvre les besoins
        moteur_possible = True
        For j = 1 To energie_vapeur_global
            If intermediaire_besoins(j) > 0 Then
                moteur_possible = False
            End If
        Next j
        If moteur_possible = True Then
            CHP_mini = biogaz_CHP_rate
            CHP_maxi = CHP_maxi
        Else
            CHP_mini = 0
            CHP_maxi = biogaz_CHP_rate
        End If
    Else
        CHP_mini = 0
        CHP_maxi = CHP_maxi
    End If

    'On initialise les valeurs de x_CHP
    x_CHP(0) = CHP_mini
    x_CHP(2) = CHP_maxi

    'calculer le nombre d'itération pour atteindre la précision souhaitée
    If x_CHP(2) - x_CHP(0) > 0 Then
        iter_max = Log((x_CHP(2) - x_CHP(0)) / precision) / Log(2)
    Else
        iter_max = 1
    End If

    'on fait les itérations
    For iter = 1 To iter_max
        'on définit le point intermédiaire
        x_CHP(1) = (x_CHP(0) + x_CHP(2)) / 2
        'pour chaque point de fonctionnement
        For I = 0 To nb_points
            'on calcule la répartition
            x_chaudiere = (1 - biogaz_sechage_rate) - biogaz_torchere_rate - x_CHP(I)
            'on initialise les besoins
            For j = 1 To energie_vapeur_global
                intermediaire_besoins(j) = besoins_th(j)
                intermediaire_dispo(j) = 0
            Next j
            bilan(I) = bilan_positif
            'on calcule l'énergie de la chaudière
            intermediaire_dispo(energie_vapeur_global) = x_chaudiere * energie_primaire_biogaz * biogaz_chaudiere_rdt
            'on définit le type de cogé
            energie_primaire_CHP = energie_primaire_biogaz * x_CHP(I)
            Call definition_type_CHP(energie_primaire_CHP)
            rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
            rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
            rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
            'on calcule les besoins après
            Call calcul_dispo_besoins_CHP(intermediaire_besoins, intermediaire_dispo, energie_primaire_CHP, rdt_th_coge)
            'on va regarder si on couvre les besoins
            For j = 1 To energie_vapeur_global
                If intermediaire_besoins(j) > 0 Then
                    bilan(I) = bilan_negatif
                End If
            Next j
        Next I
        'on va redéfinir les bornes de l'intervalle
        If bilan(0) < 0 Then
            x_CHP(0) = CHP_mini
            x_CHP(2) = CHP_mini
        ElseIf bilan(0) * bilan(1) < 0 Then
            x_CHP(0) = x_CHP(0)
            x_CHP(2) = x_CHP(1)
        ElseIf bilan(1) * bilan(2) < 0 Then
            x_CHP(0) = x_CHP(1)
            x_CHP(2) = x_CHP(2)
        ElseIf bilan(2) > 0 Then
            x_CHP(0) = CHP_maxi
            x_CHP(2) = CHP_maxi
        End If
    Next iter

    'on calcule tout bien avec la valeur conservée
    biogaz_CHP_rate = (x_CHP(0) + x_CHP(2)) / 2
    biogaz_chaudiere_rate = (1 - biogaz_sechage_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
    dispo(energie_vapeur_global) = dispo(energie_vapeur_global) + biogaz_chaudiere_rate * energie_primaire_biogaz * biogaz_chaudiere_rdt
    'on définit le type de cogé
    energie_primaire_CHP = energie_primaire_biogaz * biogaz_CHP_rate
    Call definition_type_CHP(energie_primaire_CHP)
    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
    'on calcule les besoins après
    Call calcul_dispo_besoins_CHP(besoins_th, dispo, energie_primaire_CHP, rdt_th_coge)



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.12a"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub calcul_combustible(ByRef besoins_restants, ByRef potentiel_energie, ByRef besoins, ByRef dispo)

    If version_Developpement = False Then On Error GoTo erreur_non_geree   'b
    Dim NumeroErreurNonGeree As String


    'on va calculer les combustibles à ajouter si nécessaire
    Dim I As Integer

    If potentiel_energie >= besoins_restants Then
        'les besoins sont couverts
        dispo(energie_vapeur_global) = dispo(energie_vapeur_global) + potentiel_energie - besoins_restants
        besoins_restants = 0
        For I = 1 To energie_vapeur_global
            besoins(I) = 0
        Next I
        'pas besoin de combustible d'appoint
        combustible_flux = 0
        potentiel_energie = 0    'réutilisée pour l'énergie apportée par le combustible --> calcul autosuffisance thermique
    Else
        'on calcul l'énergie manquante
        besoins_restants = besoins_restants - potentiel_energie
        'on calcule l'énergie primaire à apporter
        potentiel_energie = besoins_restants / combustible_chaudiere_rdt    'réutilisée pour l'énergie apportée par le combustible --> calcul autosuffisance thermique
        combustible_flux = potentiel_energie / PCI(choix_combustible)
        'on n'a plus de besoins
        besoins_restants = 0
        For I = 1 To energie_vapeur_global
            besoins(I) = 0
        Next I
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.12b"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub calcul_dispo_besoins_CHP(ByRef besoins, ByRef dispo, ByVal energie_primaire_CHP, ByVal rdt_th_coge)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'on calcule l'énergie disponible grâce à la CHP et on couvre les besoins
    Dim I As Integer
    Dim j As Integer
    Dim repere_energie As Integer

    For I = 1 To energie_vapeur_global
        dispo(I) = dispo(I) + energie_primaire_CHP * rdt_th_coge(I)
        If dispo(I) >= besoins(I) Then    'suffisamment d'énergie pour couvrir le niveau
            dispo(I) = dispo(I) - besoins(I)
            besoins(I) = 0
            If I > 1 Then
                For j = 1 To I - 1    'on regarde dans les niveaux inférieurs si on peut subvenir aux besoins restants
                    repere_energie = I - j
                    If besoins(repere_energie) > 0 Then   ' il reste des besoins à couvrir
                        If besoins(repere_energie) <= dispo(I) Then
                            dispo(I) = dispo(I) - besoins(repere_energie)
                            besoins(repere_energie) = 0
                        Else
                            besoins(repere_energie) = besoins(repere_energie) - dispo(I)
                            dispo(I) = 0
                        End If
                    End If
                Next j
            End If
        Else   'trop de besoins par rapport à l'énergie disponible
            besoins(I) = besoins(I) - dispo(I)
            dispo(I) = 0
        End If
    Next I


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.12c"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub definition_type_CHP(ByRef energie_primaire_CHP)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String
    Dim puissance_primaire_moteur As Double

    'Détermination du type de CHP --> micro-turbine ou moteur
    puissance_primaire_moteur = energie_primaire_CHP / pd_NOMBRE_HEURE_PAR_JOUR    'kW
    If puissance_primaire_moteur < limit_min_moteur Then    'kW
        biogaz_CHP_type = micro_turbine
    Else
        biogaz_CHP_type = moteur
    End If

    'Attribution des rendements en fonction du type de CHP et du niveau d'énergie (à partir des hypothèses)
    If biogaz_CHP_rdt_eau_BT_force = True Then
        biogaz_CHP_rdt_eau_BT = biogaz_CHP_rdt_eau_BT_f
    Else
        biogaz_CHP_rdt_eau_BT = rdt_thermique(biogaz_CHP_type, energie_eau_BT)
    End If
    If biogaz_CHP_rdt_eau_HT_force = True Then
        biogaz_CHP_rdt_eau_HT = biogaz_CHP_rdt_eau_HT_f
    Else
        biogaz_CHP_rdt_eau_HT = rdt_thermique(biogaz_CHP_type, energie_eau_HT)
    End If
    If biogaz_CHP_rdt_vapeur_force = True Then
        biogaz_CHP_rdt_vapeur = biogaz_CHP_rdt_vapeur_f
    Else
        biogaz_CHP_rdt_vapeur = rdt_thermique(biogaz_CHP_type, energie_vapeur_global)
    End If

    If biogaz_CHP_rdt_elec_force = True Then
        biogaz_CHP_rdt_elec = biogaz_CHP_rdt_elec_f
    Else
        biogaz_CHP_rdt_elec = rdt_electrique(biogaz_CHP_type)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.13"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub dimensionnement(ByVal besoins_th, ByVal P_vapeur, ByVal biogaz_caract, ByVal disponibilites_th, ByRef energie_elec, ByRef type_valorisation_biogaz, ByRef valorisation_chaudiere, ByRef ratio_biogaz_valorise)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim I As Integer
    Dim j As Integer
    Dim repere_energie As Integer
    'variables intermédiaires
    Dim besoins(energie_vapeur_global) As Double
    Dim dispo(energie_vapeur_global) As Double
    Dim besoins_restants As Double
    Dim dispo_energie As Double
    Dim besoins_totaux As Double
    Dim energie_primaire_biogaz_after_sechage As Double
    Dim energie_primaire_biogaz As Double
    Dim energie_combustible_sechage As Double
    Dim energie_primaire_CHP As Double
    Dim energie_primaire_chau As Double
    Dim potentiel_energie As Double
    Dim puissance_primaire_moteur As Double
    Dim besoins_th_coge(energie_vapeur_global) As Double
    Dim dispo_th_coge(energie_vapeur_global) As Double
    Dim rdt_th_coge(energie_vapeur_global) As Double
    Dim besoins_th_coge_restants As Double
    Dim CHP_maxi As Double
    Dim biogaz_sechage_rate_max As Double
    Dim biogaz_sechage_non_couvert As Double


    'RECUPERATION DES DONNEES
    'besoins thermiques
    For I = 1 To nb_niveaux_energie
        besoins_thermiques(I) = besoins_th(I)
    Next I
    'données sur la vapeur
    For I = 1 To nb_vapeur_max
        pression_vapeur(I) = P_vapeur(I)
    Next I
    'disponibilités
    For I = 1 To nb_niveaux_energie
        disponibilites_thermiques(I) = disponibilites_th(I)
    Next I
    'caractéristiques biogaz
    For I = 1 To nb_biogaz_caracteristiques
        biogaz_caracteristiques(I) = biogaz_caract(I)
    Next I

    'calcul de l'énergie primaire du biogaz
    energie_primaire_biogaz = PCI_CH4 * biogaz_caracteristiques(Q_Nm3j) * biogaz_caracteristiques(teneur_CH4)

    'on regroupe les énergies de vapeur
    besoins_totaux = 0
    For I = 1 To energie_eau_HT
        besoins(I) = besoins_thermiques(I)
        besoins_totaux = besoins_totaux + besoins(I)
        dispo(I) = disponibilites_thermiques(I)
    Next I
    besoins(energie_vapeur_global) = besoins_thermiques(energie_vapeur_sat1) + besoins_thermiques(energie_vapeur_sat2)
    dispo(energie_vapeur_global) = disponibilites_thermiques(energie_vapeur_sat1) + disponibilites_thermiques(energie_vapeur_sat2)
    besoins_totaux = besoins_totaux + besoins(energie_vapeur_global)
    besoins_totaux = besoins_totaux + besoins_thermiques(energie_biogaz)
    besoins_totaux = besoins_totaux + besoins_thermiques(energie_combustible)


    'CALCULS
    If besoins_totaux > 0 Then
        'On utilise en priorité les énergies eau_BT et eau_HT
        For I = 1 To energie_vapeur_global
            If dispo(I) >= besoins(I) Then    'suffisamment d'énergie pour couvrir le niveau
                dispo(I) = dispo(I) - besoins(I)
                besoins(I) = 0
                If I > 1 Then
                    For j = 1 To I - 1    'on regarde dans les niveaux inférieurs si on peut subvenir aux besoins restants
                        repere_energie = I - j
                        If besoins(repere_energie) > 0 Then   ' il reste des besoins à couvrir
                            If besoins(repere_energie) <= dispo(I) Then
                                dispo(I) = dispo(I) - besoins(repere_energie)
                                besoins(repere_energie) = 0
                            Else
                                besoins(repere_energie) = besoins(repere_energie) - dispo(I)
                                dispo(I) = 0
                            End If
                        End If
                    Next j
                End If
            Else   'trop de besoins par rapport à l'énergie disponible
                besoins(I) = besoins(I) - dispo(I)
                dispo(I) = 0
            End If
        Next I
        If energie_primaire_biogaz > 0 Then
            'on calcule l'énergie en cas de sécheur
            If besoins_thermiques(energie_biogaz) > 0 Then
                If biogaz_sechage_rate_force = True Then
                    biogaz_sechage_rate = biogaz_sechage_rate_f
                    energie_primaire_biogaz_after_sechage = energie_primaire_biogaz * (1 - biogaz_sechage_rate - biogaz_revendu_rate)
                    biogaz_sechage_non_couvert = besoins_thermiques(energie_biogaz) - energie_primaire_biogaz * biogaz_sechage_rate
                Else
                    If biogaz_torchere_rate_force = True Then
                        biogaz_torchere_rate = biogaz_torchere_rate_f
                        biogaz_sechage_rate_max = 1 - (biogaz_torchere_rate_f + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f + biogaz_revendu_rate)
                    Else
                        biogaz_sechage_rate_max = 1 - (biogaz_torchere_rate_mini + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f + biogaz_revendu_rate)
                        biogaz_torchere_rate = biogaz_torchere_rate_mini
                    End If
                    'besoins thermiques
                    If besoins_thermiques(energie_biogaz) > energie_primaire_biogaz * biogaz_sechage_rate_max Then    'le biogaz est insuffisant
                        biogaz_sechage_rate = biogaz_sechage_rate_max
                        energie_primaire_biogaz_after_sechage = energie_primaire_biogaz * (1 - biogaz_sechage_rate - biogaz_revendu_rate)
                        'on ajoute ce qui manque aux besoins en combustible
                        biogaz_sechage_non_couvert = besoins_thermiques(energie_biogaz) - energie_primaire_biogaz * biogaz_sechage_rate
                    Else    'il y a suffisamment de biogaz pour le séchage
                        If biogaz_chaudiere_rate_force = True And biogaz_CHP_rate_force = True And biogaz_torchere_rate_force = True Then
                            biogaz_sechage_rate = 1 - (biogaz_torchere_rate_f + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f + biogaz_revendu_rate)
                            'comme la valeur est fixée on l'autoforce
                            biogaz_sechage_rate_force = True
                            biogaz_sechage_rate_f = 1 - (biogaz_torchere_rate_f + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f + biogaz_revendu_rate)
                        Else
                            biogaz_sechage_rate = besoins_thermiques(energie_biogaz) / energie_primaire_biogaz
                        End If
                        energie_primaire_biogaz_after_sechage = energie_primaire_biogaz * (1 - biogaz_sechage_rate - biogaz_revendu_rate)
                    End If
                End If
            Else
                energie_primaire_biogaz_after_sechage = energie_primaire_biogaz
            End If
            If biogaz_sechage_non_couvert < 0 Then
                biogaz_sechage_non_couvert = 0
                remarque(5) = Feuil41.Range("thermique_erreur_biogaz_sechage_trop").Cells(1, 1)
            ElseIf biogaz_sechage_non_couvert > 0 Then
                remarque(6) = Feuil41.Range("thermique_erreur_biogaz_sechage_insuffisant").Cells(1, 1)
            End If
            'on calcule l'énergie primaire de combustible à apporter pour le séchage
            If besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert > 0 Then
                combustible_sechage = (besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / PCI(choix_combustible)
            Else
                combustible_sechage = 0
            End If
            If (biogaz_torchere_rate_mini + biogaz_chaudiere_rate_mini) > (1 - biogaz_sechage_rate - biogaz_revendu_rate) Then
                biogaz_torchere_rate_mini = biogaz_torchere_rate_mini / (biogaz_torchere_rate_mini + biogaz_chaudiere_rate_mini) * (1 - biogaz_sechage_rate - biogaz_revendu_rate)
                biogaz_chaudiere_rate_mini = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini
            End If
            'On rentre dans le vif du sujet avec chaudières et cogénérations
            If choix_objectif = objectif_autosuffisance_thermique Then    'AUTOSUFFISANCE THERMIQUE MAXIMALE SANS COGENERATION
                If choix_cogeneration = cogeneration_non Then   'seulement chaudières                             'fait le 6-7/06/2011 DCA
                    'on veut connaître la répartition du biogaz
                    'on calcule l'ensemble des besoins totaux (indépendemment du niveau)
                    besoins_restants = 0
                    For I = 1 To energie_vapeur_global
                        besoins_restants = besoins_restants + besoins(I)
                    Next I
                    If biogaz_chaudiere_rate_force = True Then
                        biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                        If biogaz_torchere_rate_force = True Then
                            biogaz_torchere_rate = biogaz_torchere_rate_f
                        Else
                            biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate
                        End If
                    ElseIf biogaz_torchere_rate_force = True Then
                        biogaz_torchere_rate = biogaz_torchere_rate_f
                        biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate
                    Else
                        'on calcule le potentiel d'énergie sur la chaudière
                        potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rdt * ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini)
                        If besoins_restants <= potentiel_energie Then   'énergie suffisante pour couvrir l'ensemble des besoins liés au biogaz
                            'on calcule la répartition du biogaz
                            biogaz_chaudiere_rate = (besoins_restants / biogaz_chaudiere_rdt) / energie_primaire_biogaz
                            biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate
                        Else
                            biogaz_torchere_rate = biogaz_torchere_rate_mini
                            biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate
                        End If
                    End If
                    'on va calculer les réactifs à ajouter si nécessaire
                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                    autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                ElseIf choix_cogeneration = cogeneration_oui Then    'AUTOSUFFISANCE THERMIQUE MAXIMALE AVEC COGENERATION
                    If biogaz_CHP_rate_force = True Then                                                         'fait par KME puis revu le 15/06/2011 par DCA
                        biogaz_CHP_rate = biogaz_CHP_rate_f
                        'on calcule d'abord les besoins restants car utile pour calculer la répartition chaudière-torchère lorsque nécessaire
                        'type de cogénération installable
                        energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                        Call definition_type_CHP(energie_primaire_CHP)
                        rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
                        rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
                        rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
                        'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
                        Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
                        'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
                        besoins_restants = 0
                        For I = 1 To energie_vapeur_global
                            besoins_restants = besoins_restants + besoins(I)
                        Next I
                        'on regarde les autres valeurs pour voir si elles sont forcées
                        If biogaz_chaudiere_rate_force = True Then
                            biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                            If biogaz_torchere_rate_force = True Then
                                biogaz_torchere_rate = biogaz_torchere_rate_f
                            Else
                                biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
                            End If
                        Else
                            If biogaz_torchere_rate_force = True Then
                                biogaz_torchere_rate = biogaz_torchere_rate_f
                                biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
                            Else
                                If biogaz_CHP_rate > biogaz_CHP_rate_maxi Then
                                    biogaz_torchere_rate = ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_CHP_rate) / 2  'on partage en 2 la chaudière et la torchère
                                    biogaz_chaudiere_rate = ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_CHP_rate) / 2  'on partage en 2 la chaudière et la torchère
                                Else   'on répartit au mieux entre torchère et chaudière
                                    'énergie maximale que l'on peut atteindre avec la chaudière
                                    potentiel_energie = ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_CHP_rate - biogaz_torchere_rate_mini) * energie_primaire_biogaz * biogaz_chaudiere_rdt
                                    If potentiel_energie < besoins_restants Then   'on fixe la torchère au min et la chaudière prend ce qu'il reste
                                        biogaz_torchere_rate = biogaz_torchere_rate_mini
                                        biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
                                    Else   'on calcule la répartition permettant de couvrir exactement les besoins
                                        biogaz_chaudiere_rate = (besoins_restants / biogaz_chaudiere_rdt) / energie_primaire_biogaz
                                        biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
                                    End If
                                End If
                            End If
                        End If
                        'on va calculer les réactifs à ajouter si nécessaire
                        potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                        Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                        autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                        'calcul de l'énergie électrique produite
                        energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                    ElseIf biogaz_chaudiere_rate_force = True Then
                        biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                        'on regarde si la torchère est fixée
                        If biogaz_torchere_rate_force = True Then
                            biogaz_torchere_rate = biogaz_torchere_rate_f
                            biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
                        Else
                            If biogaz_chaudiere_rate > ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini) Then
                                biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate
                                biogaz_CHP_rate = 0
                            ElseIf biogaz_chaudiere_rate < ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini - biogaz_CHP_rate_maxi) Then  'pour ne pas dépasser le maximum utilisable en CHP
                                biogaz_CHP_rate = biogaz_CHP_rate_maxi
                                biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_CHP_rate)
                            Else
                                biogaz_torchere_rate = biogaz_torchere_rate_mini
                                biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
                            End If
                        End If
                        'type de cogénération installable
                        energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                        Call definition_type_CHP(energie_primaire_CHP)
                        rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
                        rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
                        rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
                        'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
                        Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
                        'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
                        besoins_restants = 0
                        For I = 1 To energie_vapeur_global
                            besoins_restants = besoins_restants + besoins(I)
                        Next I
                        'on va calculer les réactifs à ajouter si nécessaire
                        potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                        Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                        autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                        'calcul de l'énergie électrique produite
                        energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                    Else
                        If biogaz_torchere_rate_force = True Then
                            biogaz_torchere_rate = biogaz_torchere_rate_f
                        Else
                            biogaz_torchere_rate = biogaz_torchere_rate_mini
                        End If
                        CHP_maxi = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate_mini)
                        If CHP_maxi < 0 Then
                            CHP_maxi = 0
                        End If
                        If CHP_maxi > biogaz_CHP_rate_maxi Then    'si CHP_maxi>0.9
                            CHP_maxi = biogaz_CHP_rate_maxi
                        End If
                        Call calcul_biogaz_repartition(CHP_maxi, besoins, dispo, energie_primaire_biogaz)
                        'on calcule les besoins restants
                        besoins_restants = 0
                        For I = 1 To energie_vapeur_global
                            besoins_restants = besoins_restants + besoins(I)
                        Next I
                        'on va calculer les réactifs à ajouter si nécessaire
                        potentiel_energie = 0   'énergie chaudière déjà utilisée
                        Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                        autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                        'calcul de l'énergie électrique produite
                        energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                        energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                    End If
                End If
            ElseIf choix_objectif = objectif_autosuffisance_electrique Then    'AUTOSUFFISANCE ELECTRIQUE MAXIMALE
                If biogaz_CHP_rate_force = True Then                                                'fait par KME puis revu le 14/06/2011 par DCA
                    biogaz_CHP_rate = biogaz_CHP_rate_f
                    'on calcule d'abord les besoins restants car utile pour calculer la répartition chaudière-torchère lorsque nécessaire
                    'type de cogénération installable
                    energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                    Call definition_type_CHP(energie_primaire_CHP)
                    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
                    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
                    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
                    'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
                    Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
                    'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
                    besoins_restants = 0
                    For I = 1 To energie_vapeur_global
                        besoins_restants = besoins_restants + besoins(I)
                    Next I
                    'on regarde les autres valeurs pour voir si elles sont forcées
                    If biogaz_chaudiere_rate_force = True Then
                        biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                        If biogaz_torchere_rate_force = True Then
                            biogaz_torchere_rate = biogaz_torchere_rate_f
                        Else
                            biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
                        End If
                    Else
                        If biogaz_torchere_rate_force = True Then
                            biogaz_torchere_rate = biogaz_torchere_rate_f
                            biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
                        Else
                            If biogaz_CHP_rate > biogaz_CHP_rate_maxi Then
                                biogaz_torchere_rate = ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_CHP_rate) / 2  'on partage en 2 la chaudière et la torchère
                                biogaz_chaudiere_rate = ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_CHP_rate) / 2  'on partage en 2 la chaudière et la torchère
                            Else   'on répartit au mieux entre torchère et chaudière
                                'énergie maximale que l'on peut atteindre avec la chaudière
                                potentiel_energie = ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_CHP_rate - biogaz_torchere_rate_mini) * energie_primaire_biogaz * biogaz_chaudiere_rdt
                                If potentiel_energie < besoins_restants Then   'on fixe la torchère au min et la chaudière prend ce qu'il reste
                                    biogaz_torchere_rate = biogaz_torchere_rate_mini
                                    biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
                                Else   'on calcule la répartition permettant de couvrir exactement les besoins
                                    biogaz_chaudiere_rate = (besoins_restants / biogaz_chaudiere_rdt) / energie_primaire_biogaz
                                    biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
                                End If
                            End If
                        End If
                    End If
                    'on va calculer les réactifs à ajouter si nécessaire
                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                    autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                    'calcul de l'énergie électrique produite
                    energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                ElseIf biogaz_chaudiere_rate_force = True Then                                      'fait par KME puis revu le 15/06/2011 par DCA
                    biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                    'on regarde si la torchère est fixée
                    If biogaz_torchere_rate_force = True Then
                        biogaz_torchere_rate = biogaz_torchere_rate_f
                        biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_torchere_rate)
                    Else
                        If biogaz_chaudiere_rate > ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini) Then  'il ne reste pas assez pour le minimum en torchère
                            biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate
                            biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_torchere_rate)
                        ElseIf biogaz_chaudiere_rate < ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini - biogaz_CHP_rate_maxi) Then  'pour ne pas dépasser le maximum utilisable en CHP
                            biogaz_CHP_rate = biogaz_CHP_rate_maxi
                            biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_CHP_rate)
                        Else
                            biogaz_torchere_rate = biogaz_torchere_rate_mini
                            biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_torchere_rate)
                        End If
                    End If
                    'type de cogénération installable
                    energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                    Call definition_type_CHP(energie_primaire_CHP)
                    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
                    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
                    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
                    'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
                    Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
                    'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
                    besoins_restants = 0
                    For I = 1 To energie_vapeur_global
                        besoins_restants = besoins_restants + besoins(I)
                    Next I
                    'on va calculer les réactifs à ajouter si nécessaire
                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                    autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                    'calcul de l'énergie électrique produite
                    energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                ElseIf biogaz_torchere_rate_force = True Then                                      'fait par KME puis revu le 15/06/2011 par DCA
                    biogaz_torchere_rate = biogaz_torchere_rate_f
                    If biogaz_torchere_rate > ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate_mini) Then
                        biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate
                        biogaz_CHP_rate = 0
                    ElseIf biogaz_torchere_rate < ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate_mini - biogaz_CHP_rate_maxi) Then
                        biogaz_CHP_rate = biogaz_CHP_rate_maxi
                        biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_CHP_rate)
                    Else
                        biogaz_chaudiere_rate = biogaz_chaudiere_rate_mini
                        biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
                    End If
                    'type de cogénération installable
                    energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                    Call definition_type_CHP(energie_primaire_CHP)
                    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
                    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
                    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
                    'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
                    Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
                    'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
                    besoins_restants = 0
                    For I = 1 To energie_vapeur_global
                        besoins_restants = besoins_restants + besoins(I)
                    Next I
                    'on va calculer les réactifs à ajouter si nécessaire
                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                    autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                    'calcul de l'énergie électrique produite
                    energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                Else                                                                                  'fait par KME puis revu le 15/06/2011 par DCA
                    If biogaz_sechage_rate < (1 - biogaz_torchere_rate_mini - biogaz_chaudiere_rate_mini) Then
                        biogaz_chaudiere_rate = biogaz_chaudiere_rate_mini
                        biogaz_torchere_rate = biogaz_torchere_rate_mini
                        biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
                    Else
                        biogaz_torchere_rate = biogaz_torchere_rate_mini
                        biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_torchere_rate)
                        biogaz_CHP_rate = 0
                    End If
                    'type de cogénération installable
                    energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                    Call definition_type_CHP(energie_primaire_CHP)
                    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
                    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
                    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
                    'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
                    Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
                    'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
                    besoins_restants = 0
                    For I = 1 To energie_vapeur_global
                        besoins_restants = besoins_restants + besoins(I)
                    Next I
                    'on va calculer les réactifs à ajouter si nécessaire
                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                    autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                    'calcul de l'énergie électrique produite
                    energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                End If
            End If
            'VERIFICATION DE LA REPARTITION
            If Abs(biogaz_CHP_rate) + Abs(biogaz_chaudiere_rate) + Abs(biogaz_torchere_rate) + Abs(biogaz_sechage_rate) + Abs(biogaz_revendu_rate) <> 1 Then
                MsgBox Feuil2.Range("synthese_thermique").Cells(1, 1) & ": " & Feuil41.Range("thermique_erreur_repartition_globale").Cells(1, 1), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub
            End If
            'Vérification des flux minimaux et maximaux
            If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                If biogaz_chaudiere_rate < biogaz_chaudiere_rate_mini Then
                    remarque(1) = Feuil41.Range("thermique_erreur_repartition_mini_chaudiere").Cells(1, 1)
                End If
                If biogaz_torchere_rate < biogaz_torchere_rate_mini Then
                    remarque(2) = Feuil41.Range("thermique_erreur_repartition_mini_torchere").Cells(1, 1)
                End If
                If biogaz_CHP_rate > biogaz_CHP_rate_maxi Then
                    remarque(3) = Feuil41.Range("thermique_erreur_repartition_maxi_CHP").Cells(1, 1)
                End If
            End If
        Else   'PAS DE BIOGAZ A REPARTIR
            biogaz_torchere_rate = 0
            biogaz_chaudiere_rate = 0
            biogaz_CHP_rate = 0
            biogaz_sechage_rate = 0
            'on calcule les besoins restants pour savoir ce qu'il manque
            besoins_restants = 0
            For I = 1 To energie_vapeur_global
                besoins_restants = besoins_restants + besoins(I)
            Next I
            potentiel_energie = 0       'ici pas d'énergie valorisable en chaudière
            Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
            biogaz_sechage_non_couvert = besoins_thermiques(energie_biogaz)
            If biogaz_sechage_non_couvert < 0 Then
                biogaz_sechage_non_couvert = 0
                remarque(5) = Feuil41.Range("thermique_erreur_biogaz_sechage_trop").Cells(1, 1)
            Else
                remarque(6) = Feuil41.Range("thermique_erreur_biogaz_sechage_insuffisant").Cells(1, 1)
            End If
            autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
            'on calcule l'énergie primaire de combustible à apporter pour le séchage
            If besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert > 0 Then
                combustible_sechage = (besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / PCI(choix_combustible)
            Else
                combustible_sechage = 0
            End If
        End If
    Else
        autosuffisance_thermique = 1
    End If

    'CALCUL DE L'ENERGIE NON UTILISEE
    For I = 1 To energie_vapeur_global
        energie_fatale_kWhj(I) = dispo(I) - besoins(I)
    Next I


    'GESTION DU NOMBRE D'EQUIPEMENTS DE COGENERATION
    'calcul de la puissance électrique de la cogénération
    biogaz_CHP_Pw = energie_elec / pd_NOMBRE_HEURE_PAR_JOUR   'production électrique moyenne journalière rapportée en puissance instantanée moyenne sur la journée
    electricite_biogaz = energie_elec

    'on calcule le PCI correspondant au combustible utilisé (pour empreinte CO2)
    If choix_combustible = combustible_repere_gaz_naturel Then
        gaz_naturel_kWhPCIj = (combustible_flux + combustible_sechage) * PCI(choix_combustible)
    ElseIf choix_combustible = combustible_repere_fioul Then
        fioul_kWhPCIj = (combustible_flux + combustible_sechage) * PCI(choix_combustible)
    End If

    'ATTRIBUTION DES PARAMETRES SUR VALORISATION POUR CONSOS DE LA LIGNE DE BIOGAZ (inclus dans digestion)
    If biogaz_CHP_rate > 0 Then
        type_valorisation_biogaz = biogaz_CHP_type
    Else
        type_valorisation_biogaz = chaudiere
    End If
    valorisation_chaudiere = chaudiere  'constante pour savoir si on a utilisé une chaudière
    ratio_biogaz_valorise = 1 - biogaz_torchere_rate

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub fonctionnement_reel(ByVal besoins_th, ByVal P_vapeur, ByVal biogaz_caract, ByVal disponibilites_th, ByRef energie_elec, ByRef type_valorisation_biogaz, ByRef valorisation_chaudiere, ByRef ratio_biogaz_valorise)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim I As Integer
    Dim j As Integer
    Dim repere_energie As Integer
    'variables intermédiaires
    Dim besoins(energie_vapeur_global) As Double
    Dim dispo(energie_vapeur_global) As Double
    Dim besoins_restants As Double
    Dim dispo_energie As Double
    Dim besoins_totaux As Double
    Dim energie_primaire_biogaz_after_sechage As Double
    Dim energie_primaire_biogaz As Double
    Dim energie_combustible_sechage As Double
    Dim energie_primaire_CHP As Double
    Dim energie_primaire_chau As Double
    Dim potentiel_energie As Double
    Dim puissance_primaire_moteur As Double
    Dim besoins_th_coge(energie_vapeur_global) As Double
    Dim dispo_th_coge(energie_vapeur_global) As Double
    Dim rdt_th_coge(energie_vapeur_global) As Double
    Dim besoins_th_coge_restants As Double
    Dim CHP_maxi As Double
    Dim biogaz_sechage_rate_max As Double
    Dim biogaz_sechage_non_couvert As Double


    'RECUPERATION DES DONNEES
    'besoins thermiques
    For I = 1 To nb_niveaux_energie
        besoins_thermiques(I) = besoins_th(I)
    Next I
    'données sur la vapeur
    For I = 1 To nb_vapeur_max
        pression_vapeur(I) = P_vapeur(I)
    Next I
    'disponibilités
    For I = 1 To nb_niveaux_energie
        disponibilites_thermiques(I) = disponibilites_th(I)
    Next I
    'caractéristiques biogaz
    For I = 1 To nb_biogaz_caracteristiques
        biogaz_caracteristiques(I) = biogaz_caract(I)
    Next I

    'calcul de l'énergie primaire du biogaz
    energie_primaire_biogaz = PCI_CH4 * biogaz_caracteristiques(Q_Nm3j) * biogaz_caracteristiques(teneur_CH4)

    'on regroupe les énergies de vapeur
    besoins_totaux = 0
    For I = 1 To energie_eau_HT
        besoins(I) = besoins_thermiques(I)
        besoins_totaux = besoins_totaux + besoins(I)
        dispo(I) = disponibilites_thermiques(I)
    Next I
    besoins(energie_vapeur_global) = besoins_thermiques(energie_vapeur_sat1) + besoins_thermiques(energie_vapeur_sat2)
    dispo(energie_vapeur_global) = disponibilites_thermiques(energie_vapeur_sat1) + disponibilites_thermiques(energie_vapeur_sat2)
    besoins_totaux = besoins_totaux + besoins(energie_vapeur_global)
    besoins_totaux = besoins_totaux + besoins_thermiques(energie_biogaz)
    besoins_totaux = besoins_totaux + besoins_thermiques(energie_combustible)


    'CALCULS
    If besoins_totaux > 0 Then
        'On utilise en priorité les énergies eau_BT et eau_HT
        For I = 1 To energie_vapeur_global
            If dispo(I) >= besoins(I) Then    'suffisamment d'énergie pour couvrir le niveau
                dispo(I) = dispo(I) - besoins(I)
                besoins(I) = 0
                If I > 1 Then
                    For j = 1 To I - 1    'on regarde dans les niveaux inférieurs si on peut subvenir aux besoins restants
                        repere_energie = I - j
                        If besoins(repere_energie) > 0 Then   ' il reste des besoins à couvrir
                            If besoins(repere_energie) <= dispo(I) Then
                                dispo(I) = dispo(I) - besoins(repere_energie)
                                besoins(repere_energie) = 0
                            Else
                                besoins(repere_energie) = besoins(repere_energie) - dispo(I)
                                dispo(I) = 0
                            End If
                        End If
                    Next j
                End If
            Else   'trop de besoins par rapport à l'énergie disponible
                besoins(I) = besoins(I) - dispo(I)
                dispo(I) = 0
            End If
        Next I
        If energie_primaire_biogaz > 0 Then
            'on calcule l'énergie en cas de sécheur
            If besoins_thermiques(energie_biogaz) > 0 Then
                If biogaz_sechage_rate_force = True Then
                    biogaz_sechage_rate = biogaz_sechage_rate_f
                    energie_primaire_biogaz_after_sechage = energie_primaire_biogaz * (1 - biogaz_sechage_rate)
                    biogaz_sechage_non_couvert = besoins_thermiques(energie_biogaz) - energie_primaire_biogaz * biogaz_sechage_rate
                Else
                    If biogaz_torchere_rate_force = True Then
                        biogaz_torchere_rate = biogaz_torchere_rate_f
                        biogaz_sechage_rate_max = 1 - (biogaz_torchere_rate_f + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f + biogaz_revendu_rate)
                    Else
                        biogaz_sechage_rate_max = 1 - (biogaz_torchere_rate_mini + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f + biogaz_revendu_rate)
                        biogaz_torchere_rate = biogaz_torchere_rate_mini
                    End If
                    'besoins thermiques
                    If besoins_thermiques(energie_biogaz) > energie_primaire_biogaz * biogaz_sechage_rate_max Then    'le biogaz est insuffisant
                        biogaz_sechage_rate = biogaz_sechage_rate_max
                        energie_primaire_biogaz_after_sechage = energie_primaire_biogaz * (1 - biogaz_sechage_rate)
                        'on ajoute ce qui manque aux besoins en combustible
                        biogaz_sechage_non_couvert = besoins_thermiques(energie_biogaz) - energie_primaire_biogaz * biogaz_sechage_rate
                    Else    'il y a suffisamment de biogaz pour le séchage
                        If biogaz_chaudiere_rate_force = True And biogaz_CHP_rate_force = True And biogaz_torchere_rate_force = True Then
                            biogaz_sechage_rate = 1 - (biogaz_torchere_rate_f + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f + biogaz_revendu_rate)
                            'comme la valeur est fixée on l'autoforce
                            biogaz_sechage_rate_force = True
                            biogaz_sechage_rate_f = 1 - (biogaz_torchere_rate_f + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f + biogaz_revendu_rate)
                        Else
                            biogaz_sechage_rate = besoins_thermiques(energie_biogaz) / energie_primaire_biogaz
                        End If
                        energie_primaire_biogaz_after_sechage = energie_primaire_biogaz * (1 - biogaz_sechage_rate)
                    End If
                End If
            Else
                energie_primaire_biogaz_after_sechage = energie_primaire_biogaz
            End If
            If biogaz_sechage_non_couvert < 0 Then
                biogaz_sechage_non_couvert = 0
                remarque(5) = Feuil41.Range("thermique_erreur_biogaz_sechage_trop").Cells(1, 1)
            ElseIf biogaz_sechage_non_couvert > 0 Then
                remarque(6) = Feuil41.Range("thermique_erreur_biogaz_sechage_insuffisant").Cells(1, 1)
            End If
            'on calcule l'énergie primaire de combustible à apporter pour le séchage
            If besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert > 0 Then
                pd_BesoinsCombustibleExterneIncineration_kWhj = (besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert)
                combustible_sechage = (besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / PCI(choix_combustible)
            Else
                combustible_sechage = 0
            End If
            If (biogaz_torchere_rate_mini + biogaz_chaudiere_rate_mini) > (1 - biogaz_sechage_rate - biogaz_revendu_rate) Then
                biogaz_torchere_rate_mini = biogaz_torchere_rate_mini / (biogaz_torchere_rate_mini + biogaz_chaudiere_rate_mini) * (1 - biogaz_sechage_rate - biogaz_revendu_rate)
                biogaz_chaudiere_rate_mini = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini
            End If
            'On rentre dans le vif du sujet avec chaudières et cogénérations
            If choix_objectif = objectif_autosuffisance_thermique Then    'AUTOSUFFISANCE THERMIQUE MAXIMALE SANS COGENERATION
                If choix_cogeneration = cogeneration_non Then   'seulement chaudières                             'fait le 6-7/06/2011 DCA
                    'on veut connaître la répartition du biogaz
                    'on calcule l'ensemble des besoins totaux (indépendemment du niveau)
                    besoins_restants = 0
                    For I = 1 To energie_vapeur_global
                        besoins_restants = besoins_restants + besoins(I)
                    Next I
                    If biogaz_chaudiere_rate_force = True Then
                        biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                        If biogaz_torchere_rate_force = True Then
                            biogaz_torchere_rate = biogaz_torchere_rate_f
                        Else
                            biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate
                        End If
                    ElseIf biogaz_torchere_rate_force = True Then
                        biogaz_torchere_rate = biogaz_torchere_rate_f
                        biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate
                    Else
                        'on calcule le potentiel d'énergie sur la chaudière
                        potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rdt * ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini)
                        If besoins_restants <= potentiel_energie Then   'énergie suffisante pour couvrir l'ensemble des besoins liés au biogaz
                            'on calcule la répartition du biogaz
                            biogaz_chaudiere_rate = (besoins_restants / biogaz_chaudiere_rdt) / energie_primaire_biogaz
                            biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate
                        Else
                            biogaz_torchere_rate = biogaz_torchere_rate_mini
                            biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate
                        End If
                    End If
                    'on va calculer les réactifs à ajouter si nécessaire
                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                    autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                ElseIf choix_cogeneration = cogeneration_oui Then    'AUTOSUFFISANCE THERMIQUE MAXIMALE AVEC COGENERATION
                    If biogaz_CHP_rate_force = True Then                                                         'fait par KME puis revu le 15/06/2011 par DCA
                        biogaz_CHP_rate = biogaz_CHP_rate_f
                        'on calcule d'abord les besoins restants car utile pour calculer la répartition chaudière-torchère lorsque nécessaire
                        'type de cogénération installable
                        energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                        Call definition_type_CHP(energie_primaire_CHP)
                        rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
                        rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
                        rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
                        'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
                        Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
                        'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
                        besoins_restants = 0
                        For I = 1 To energie_vapeur_global
                            besoins_restants = besoins_restants + besoins(I)
                        Next I
                        'on regarde les autres valeurs pour voir si elles sont forcées
                        If biogaz_chaudiere_rate_force = True Then
                            biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                            If biogaz_torchere_rate_force = True Then
                                biogaz_torchere_rate = biogaz_torchere_rate_f
                            Else
                                biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
                            End If
                        Else
                            If biogaz_torchere_rate_force = True Then
                                biogaz_torchere_rate = biogaz_torchere_rate_f
                                biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
                            Else
                                If biogaz_CHP_rate > biogaz_CHP_rate_maxi Then
                                    biogaz_torchere_rate = ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_CHP_rate) / 2  'on partage en 2 la chaudière et la torchère
                                    biogaz_chaudiere_rate = ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_CHP_rate) / 2  'on partage en 2 la chaudière et la torchère
                                Else   'on répartit au mieux entre torchère et chaudière
                                    'énergie maximale que l'on peut atteindre avec la chaudière
                                    potentiel_energie = ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_CHP_rate - biogaz_torchere_rate_mini) * energie_primaire_biogaz * biogaz_chaudiere_rdt
                                    If potentiel_energie < besoins_restants Then   'on fixe la torchère au min et la chaudière prend ce qu'il reste
                                        biogaz_torchere_rate = biogaz_torchere_rate_mini
                                        biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
                                    Else   'on calcule la répartition permettant de couvrir exactement les besoins
                                        biogaz_chaudiere_rate = (besoins_restants / biogaz_chaudiere_rdt) / energie_primaire_biogaz
                                        biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
                                    End If
                                End If
                            End If
                        End If
                        'on va calculer les réactifs à ajouter si nécessaire
                        potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                        Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                        autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                        'calcul de l'énergie électrique produite
                        energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                    ElseIf biogaz_chaudiere_rate_force = True Then
                        biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                        'on regarde si la torchère est fixée
                        If biogaz_torchere_rate_force = True Then
                            biogaz_torchere_rate = biogaz_torchere_rate_f
                            biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
                        Else
                            If biogaz_chaudiere_rate > ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini) Then
                                biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate
                                biogaz_CHP_rate = 0
                            ElseIf biogaz_chaudiere_rate < ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini - biogaz_CHP_rate_maxi) Then  'pour ne pas dépasser le maximum utilisable en CHP
                                biogaz_CHP_rate = biogaz_CHP_rate_maxi
                                biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_CHP_rate)
                            Else
                                biogaz_torchere_rate = biogaz_torchere_rate_mini
                                biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
                            End If
                        End If
                        'type de cogénération installable
                        energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                        Call definition_type_CHP(energie_primaire_CHP)
                        rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
                        rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
                        rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
                        'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
                        Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
                        'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
                        besoins_restants = 0
                        For I = 1 To energie_vapeur_global
                            besoins_restants = besoins_restants + besoins(I)
                        Next I
                        'on va calculer les réactifs à ajouter si nécessaire
                        potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                        Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                        autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                        'calcul de l'énergie électrique produite
                        energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                    Else
                        If biogaz_torchere_rate_force = True Then
                            biogaz_torchere_rate = biogaz_torchere_rate_f
                        Else
                            biogaz_torchere_rate = biogaz_torchere_rate_mini
                        End If
                        CHP_maxi = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate_mini)
                        If CHP_maxi < 0 Then
                            CHP_maxi = 0
                        End If
                        If CHP_maxi > biogaz_CHP_rate_maxi Then    'si CHP_maxi>0.9
                            CHP_maxi = biogaz_CHP_rate_maxi
                        End If
                        Call calcul_biogaz_repartition(CHP_maxi, besoins, dispo, energie_primaire_biogaz)
                        'on calcule les besoins restants
                        besoins_restants = 0
                        For I = 1 To energie_vapeur_global
                            besoins_restants = besoins_restants + besoins(I)
                        Next I
                        'on va calculer les réactifs à ajouter si nécessaire
                        potentiel_energie = 0   'énergie chaudière déjà utilisée
                        Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                        autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                        'calcul de l'énergie électrique produite
                        energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                        energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                    End If
                End If
            ElseIf choix_objectif = objectif_autosuffisance_electrique Then    'AUTOSUFFISANCE ELECTRIQUE MAXIMALE
                If biogaz_CHP_rate_force = True Then                                                'fait par KME puis revu le 14/06/2011 par DCA
                    biogaz_CHP_rate = biogaz_CHP_rate_f
                    'on calcule d'abord les besoins restants car utile pour calculer la répartition chaudière-torchère lorsque nécessaire
                    'type de cogénération installable
                    energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                    Call definition_type_CHP(energie_primaire_CHP)
                    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
                    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
                    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
                    'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
                    Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
                    'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
                    besoins_restants = 0
                    For I = 1 To energie_vapeur_global
                        besoins_restants = besoins_restants + besoins(I)
                    Next I
                    'on regarde les autres valeurs pour voir si elles sont forcées
                    If biogaz_chaudiere_rate_force = True Then
                        biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                        If biogaz_torchere_rate_force = True Then
                            biogaz_torchere_rate = biogaz_torchere_rate_f
                        Else
                            biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
                        End If
                    Else
                        If biogaz_torchere_rate_force = True Then
                            biogaz_torchere_rate = biogaz_torchere_rate_f
                            biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
                        Else
                            If biogaz_CHP_rate > biogaz_CHP_rate_maxi Then
                                biogaz_torchere_rate = ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_CHP_rate) / 2  'on partage en 2 la chaudière et la torchère
                                biogaz_chaudiere_rate = ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_CHP_rate) / 2  'on partage en 2 la chaudière et la torchère
                            Else   'on répartit au mieux entre torchère et chaudière
                                'énergie maximale que l'on peut atteindre avec la chaudière
                                potentiel_energie = ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_CHP_rate - biogaz_torchere_rate_mini) * energie_primaire_biogaz * biogaz_chaudiere_rdt
                                If potentiel_energie < besoins_restants Then   'on fixe la torchère au min et la chaudière prend ce qu'il reste
                                    biogaz_torchere_rate = biogaz_torchere_rate_mini
                                    biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
                                Else   'on calcule la répartition permettant de couvrir exactement les besoins
                                    biogaz_chaudiere_rate = (besoins_restants / biogaz_chaudiere_rdt) / energie_primaire_biogaz
                                    biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
                                End If
                            End If
                        End If
                    End If
                    'on va calculer les réactifs à ajouter si nécessaire
                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                    autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                    'calcul de l'énergie électrique produite
                    energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                ElseIf biogaz_chaudiere_rate_force = True Then                                      'fait par KME puis revu le 15/06/2011 par DCA
                    biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                    'on regarde si la torchère est fixée
                    If biogaz_torchere_rate_force = True Then
                        biogaz_torchere_rate = biogaz_torchere_rate_f
                        biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_torchere_rate)
                    Else
                        If biogaz_chaudiere_rate > ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini) Then  'il ne reste pas assez pour le minimum en torchère
                            biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate
                            biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_torchere_rate)
                        ElseIf biogaz_chaudiere_rate < ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini - biogaz_CHP_rate_maxi) Then  'pour ne pas dépasser le maximum utilisable en CHP
                            biogaz_CHP_rate = biogaz_CHP_rate_maxi
                            biogaz_torchere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_CHP_rate)
                        Else
                            biogaz_torchere_rate = biogaz_torchere_rate_mini
                            biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_torchere_rate)
                        End If
                    End If
                    'type de cogénération installable
                    energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                    Call definition_type_CHP(energie_primaire_CHP)
                    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
                    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
                    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
                    'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
                    Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
                    'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
                    besoins_restants = 0
                    For I = 1 To energie_vapeur_global
                        besoins_restants = besoins_restants + besoins(I)
                    Next I
                    'on va calculer les réactifs à ajouter si nécessaire
                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                    autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                    'calcul de l'énergie électrique produite
                    energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                ElseIf biogaz_torchere_rate_force = True Then                                      'fait par KME puis revu le 15/06/2011 par DCA
                    biogaz_torchere_rate = biogaz_torchere_rate_f
                    If biogaz_torchere_rate > ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate_mini) Then
                        biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_torchere_rate
                        biogaz_CHP_rate = 0
                    ElseIf biogaz_torchere_rate < ((1 - biogaz_sechage_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate_mini - biogaz_CHP_rate_maxi) Then
                        biogaz_CHP_rate = biogaz_CHP_rate_maxi
                        biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_CHP_rate)
                    Else
                        biogaz_chaudiere_rate = biogaz_chaudiere_rate_mini
                        biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
                    End If
                    'type de cogénération installable
                    energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                    Call definition_type_CHP(energie_primaire_CHP)
                    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
                    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
                    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
                    'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
                    Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
                    'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
                    besoins_restants = 0
                    For I = 1 To energie_vapeur_global
                        besoins_restants = besoins_restants + besoins(I)
                    Next I
                    'on va calculer les réactifs à ajouter si nécessaire
                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                    autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                    'calcul de l'énergie électrique produite
                    energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                Else                                                                                  'fait par KME puis revu le 15/06/2011 par DCA
                    If biogaz_sechage_rate < (1 - biogaz_torchere_rate_mini - biogaz_chaudiere_rate_mini) Then
                        biogaz_chaudiere_rate = biogaz_chaudiere_rate_mini
                        biogaz_torchere_rate = biogaz_torchere_rate_mini
                        biogaz_CHP_rate = (1 - biogaz_sechage_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
                    Else
                        biogaz_torchere_rate = biogaz_torchere_rate_mini
                        biogaz_chaudiere_rate = (1 - biogaz_sechage_rate - biogaz_torchere_rate)
                        biogaz_CHP_rate = 0
                    End If
                    'type de cogénération installable
                    energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                    Call definition_type_CHP(energie_primaire_CHP)
                    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
                    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
                    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
                    'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
                    Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
                    'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
                    besoins_restants = 0
                    For I = 1 To energie_vapeur_global
                        besoins_restants = besoins_restants + besoins(I)
                    Next I
                    'on va calculer les réactifs à ajouter si nécessaire
                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
                    autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
                    'calcul de l'énergie électrique produite
                    energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
                End If
            End If
            'VERIFICATION DE LA REPARTITION
            If Abs(biogaz_CHP_rate) + Abs(biogaz_chaudiere_rate) + Abs(biogaz_torchere_rate) + Abs(biogaz_sechage_rate) + Abs(biogaz_revendu_rate) <> 1 Then
                MsgBox Feuil2.Range("synthese_thermique").Cells(1, 1) & ": " & Feuil41.Range("thermique_erreur_repartition_globale").Cells(1, 1), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub
            End If
            'Vérification des flux minimaux et maximaux
            If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                If biogaz_chaudiere_rate < biogaz_chaudiere_rate_mini Then
                    remarque(1) = Feuil41.Range("thermique_erreur_repartition_mini_chaudiere").Cells(1, 1)
                End If
                If biogaz_torchere_rate < biogaz_torchere_rate_mini Then
                    remarque(2) = Feuil41.Range("thermique_erreur_repartition_mini_torchere").Cells(1, 1)
                End If
                If biogaz_CHP_rate > biogaz_CHP_rate_maxi Then
                    remarque(3) = Feuil41.Range("thermique_erreur_repartition_maxi_CHP").Cells(1, 1)
                End If
            End If
        Else   'PAS DE BIOGAZ A REPARTIR
            biogaz_torchere_rate = 0
            biogaz_chaudiere_rate = 0
            biogaz_CHP_rate = 0
            biogaz_sechage_rate = 0
            'on calcule les besoins restants pour savoir ce qu'il manque
            besoins_restants = 0
            For I = 1 To energie_vapeur_global
                besoins_restants = besoins_restants + besoins(I)
            Next I
            potentiel_energie = 0       'ici pas d'énergie valorisable en chaudière
            Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
            biogaz_sechage_non_couvert = besoins_thermiques(energie_biogaz)
            If biogaz_sechage_non_couvert < 0 Then
                biogaz_sechage_non_couvert = 0
                remarque(5) = Feuil41.Range("thermique_erreur_biogaz_sechage_trop").Cells(1, 1)
            Else
                remarque(6) = Feuil41.Range("thermique_erreur_biogaz_sechage_insuffisant").Cells(1, 1)
            End If
            autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
            'on calcule l'énergie primaire de combustible à apporter pour le séchage
            If besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert > 0 Then
                combustible_sechage = (besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / PCI(choix_combustible)
            Else
                combustible_sechage = 0
            End If
        End If
    Else
        autosuffisance_thermique = 1
    End If

    'CALCUL DE L'ENERGIE NON UTILISEE
    For I = 1 To energie_vapeur_global
        energie_fatale_kWhj(I) = dispo(I) - besoins(I)
    Next I


    'GESTION DU NOMBRE D'EQUIPEMENTS DE COGENERATION
    'calcul de la puissance électrique de la cogénération
    biogaz_CHP_Pw_reel = energie_elec / pd_NOMBRE_HEURE_PAR_JOUR
    electricite_biogaz = energie_elec


    'on calcule le PCI correspondant au combustible utilisé (pour empreinte CO2)
    If choix_combustible = combustible_repere_gaz_naturel Then
        gaz_naturel_kWhPCIj = (combustible_flux + combustible_sechage) * PCI(choix_combustible)
    ElseIf choix_combustible = combustible_repere_fioul Then
        fioul_kWhPCIj = (combustible_flux + combustible_sechage) * PCI(choix_combustible)
    End If

    'ATTRIBUTION DES PARAMETRES SUR VALORISATION POUR CONSOS DE LA LIGNE DE BIOGAZ (inclus dans digestion)
    If biogaz_CHP_rate > 0 Then
        type_valorisation_biogaz = biogaz_CHP_type
    Else
        type_valorisation_biogaz = chaudiere
    End If
    valorisation_chaudiere = chaudiere  'constante pour savoir si on a utilisé une chaudière
    ratio_biogaz_valorise = 1 - biogaz_torchere_rate

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    For I = 1 To nb_niveaux_energie
        Feuil41.Range("thermique_besoins_thermiques_kWhj").Cells(I, Colonne) = besoins_thermiques(I)
    Next I
    For I = 1 To nb_vapeur_max
        Feuil41.Range("thermique_pression_vapeur").Cells(I, Colonne) = pression_vapeur(I)
    Next I
    For I = 1 To nb_vapeur_max
        Feuil41.Range("thermique_besoins_vapeur_kgj").Cells(I, Colonne) = besoins_vapeur_kgj(I)
    Next I
    For I = 1 To energie_vapeur_global
        Feuil41.Range("thermique_disponibilites_thermiques_kWhj").Cells(I, Colonne) = disponibilites_thermiques(I)
    Next I
    For I = 1 To nb_biogaz_caracteristiques
        Feuil41.Range("thermique_biogaz_caracteristiques").Cells(I, Colonne) = biogaz_caracteristiques(I)
    Next I
    If biogaz_chaudiere_rate > zero_force Then
        Feuil41.Range("thermique_biogaz_chaudiere_rdt").Cells(1, Colonne) = biogaz_chaudiere_rdt
        Feuil41.Range("thermique_biogaz_chaudiere_Tch").Cells(1, Colonne) = biogaz_chaudiere_Tch
        Feuil41.Range("thermique_biogaz_chaudiere_Tfr").Cells(1, Colonne) = biogaz_chaudiere_Tfr
    Else
        Feuil41.Range("thermique_biogaz_chaudiere_rdt").Cells(1, Colonne) = "'-"
        Feuil41.Range("thermique_biogaz_chaudiere_Tch").Cells(1, Colonne) = "'-"
        Feuil41.Range("thermique_biogaz_chaudiere_Tfr").Cells(1, Colonne) = "'-"
    End If
    If choix_cogeneration = cogeneration_oui And biogaz_CHP_rate > 0 Then
        If biogaz_CHP_type = micro_turbine Then
            Feuil41.Range("thermique_biogaz_CHP_type").Cells(1, Colonne) = Feuil41.Range("thermique_microturbine").Cells(1, 1)
        ElseIf biogaz_CHP_type = moteur Then
            Feuil41.Range("thermique_biogaz_CHP_type").Cells(1, Colonne) = Feuil41.Range("thermique_moteur").Cells(1, 1)
        Else
            Feuil41.Range("thermique_biogaz_CHP_type").Cells(1, Colonne) = ""
        End If
        Feuil41.Range("thermique_biogaz_CHP_rdt_elec").Cells(1, Colonne) = biogaz_CHP_rdt_elec
        Feuil41.Range("thermique_biogaz_CHP_rdt_eau_BT").Cells(1, Colonne) = biogaz_CHP_rdt_eau_BT
        If biogaz_CHP_rdt_eau_BT > zero_force Then
            Feuil41.Range("thermique_biogaz_CHP_BT_Tch").Cells(1, Colonne) = biogaz_CHP_BT_Tch
            Feuil41.Range("thermique_biogaz_CHP_BT_Tfr").Cells(1, Colonne) = biogaz_CHP_BT_Tfr
        Else
            Feuil41.Range("thermique_biogaz_CHP_BT_Tch").Cells(1, Colonne) = "'-"     'biogaz_CHP_BT_Tch
            Feuil41.Range("thermique_biogaz_CHP_BT_Tfr").Cells(1, Colonne) = "'-"     'biogaz_CHP_BT_Tfr
        End If
        Feuil41.Range("thermique_biogaz_CHP_rdt_eau_HT").Cells(1, Colonne) = biogaz_CHP_rdt_eau_HT
        If biogaz_CHP_rdt_eau_HT > zero_force Then
            Feuil41.Range("thermique_biogaz_CHP_HT_Tch").Cells(1, Colonne) = biogaz_CHP_HT_Tch
            Feuil41.Range("thermique_biogaz_CHP_HT_Tfr").Cells(1, Colonne) = biogaz_CHP_HT_Tfr
        Else
            Feuil41.Range("thermique_biogaz_CHP_HT_Tch").Cells(1, Colonne) = "'-"     'biogaz_CHP_HT_Tch
            Feuil41.Range("thermique_biogaz_CHP_HT_Tfr").Cells(1, Colonne) = "'-"     'biogaz_CHP_HT_Tfr
        End If
        Feuil41.Range("thermique_biogaz_CHP_rdt_vapeur").Cells(1, Colonne) = biogaz_CHP_rdt_vapeur
        If biogaz_CHP_rdt_vapeur > zero_force Then
            Feuil41.Range("thermique_biogaz_CHP_fumees_Tch").Cells(1, Colonne) = biogaz_CHP_fumees_Tch
            Feuil41.Range("thermique_biogaz_CHP_fumees_Tfr").Cells(1, Colonne) = biogaz_CHP_fumees_Tfr
        Else
            Feuil41.Range("thermique_biogaz_CHP_fumees_Tch").Cells(1, Colonne) = "'-"     'biogaz_CHP_fumees_Tch
            Feuil41.Range("thermique_biogaz_CHP_fumees_Tfr").Cells(1, Colonne) = "'-"     'biogaz_CHP_fumees_Tfr
        End If
        Feuil41.Range("thermique_biogaz_CHP_Pw").Cells(1, Colonne) = biogaz_CHP_Pw
        Feuil41.Range("thermique_biogaz_CHP_taux_charge").Cells(1, Colonne) = biogaz_CHP_Pw_reel
        Feuil41.Range("thermique_biogaz_CHP_rate").Cells(1, Colonne) = biogaz_CHP_rate
        Feuil41.Range("thermique_electricite_biogaz").Cells(1, Colonne) = electricite_biogaz
    Else
        Feuil41.Range("thermique_biogaz_CHP_type").Cells(1, Colonne) = Feuil41.Range("thermique_no_CHP").Cells(1, 1)    'Feuil41.Range("thermique_microturbine").Cells(1, 1)
        Feuil41.Range("thermique_biogaz_CHP_rdt_elec").Cells(1, Colonne) = "'-"    'biogaz_CHP_rdt_elec
        Feuil41.Range("thermique_biogaz_CHP_rdt_eau_BT").Cells(1, Colonne) = "'-"    'biogaz_CHP_rdt_eau_BT
        Feuil41.Range("thermique_biogaz_CHP_BT_Tch").Cells(1, Colonne) = "'-"     'biogaz_CHP_BT_Tch
        Feuil41.Range("thermique_biogaz_CHP_BT_Tfr").Cells(1, Colonne) = "'-"     'biogaz_CHP_BT_Tfr
        Feuil41.Range("thermique_biogaz_CHP_rdt_eau_HT").Cells(1, Colonne) = "'-"    'biogaz_CHP_rdt_eau_HT
        Feuil41.Range("thermique_biogaz_CHP_HT_Tch").Cells(1, Colonne) = "'-"     'biogaz_CHP_HT_Tch
        Feuil41.Range("thermique_biogaz_CHP_HT_Tfr").Cells(1, Colonne) = "'-"     'biogaz_CHP_HT_Tfr
        Feuil41.Range("thermique_biogaz_CHP_rdt_vapeur").Cells(1, Colonne) = "'-"    'biogaz_CHP_rdt_vapeur
        Feuil41.Range("thermique_biogaz_CHP_fumees_Tch").Cells(1, Colonne) = "'-"     'biogaz_CHP_fumees_Tch
        Feuil41.Range("thermique_biogaz_CHP_fumees_Tfr").Cells(1, Colonne) = "'-"     'biogaz_CHP_fumees_Tfr
        Feuil41.Range("thermique_biogaz_CHP_Pw").Cells(1, Colonne) = "'-"    'biogaz_CHP_Pw
        Feuil41.Range("thermique_biogaz_CHP_taux_charge").Cells(1, Colonne) = "'-"    'biogaz_CHP_Pw_reel
        Feuil41.Range("thermique_biogaz_CHP_rate").Cells(1, Colonne) = biogaz_CHP_rate
        Feuil41.Range("thermique_electricite_biogaz").Cells(1, Colonne) = "'-"    'electricite_biogaz
    End If
    Feuil41.Range("thermique_biogaz_chaudiere_rate").Cells(1, Colonne) = biogaz_chaudiere_rate
    Feuil41.Range("thermique_biogaz_torchere_rate").Cells(1, Colonne) = biogaz_torchere_rate
    Feuil41.Range("thermique_biogaz_sechage_rate").Cells(1, Colonne) = biogaz_sechage_rate
    Feuil41.Range("thermique_combustible_chaudiere_rdt").Cells(1, Colonne) = combustible_chaudiere_rdt
    Feuil41.Range("thermique_combustible_flux").Cells(1, Colonne) = combustible_flux
    Feuil41.Range("thermique_combustible_sechage").Cells(1, Colonne) = combustible_sechage
    For I = 1 To energie_vapeur_global
        Feuil41.Range("thermique_energie_fatale_kWhj").Cells(I, Colonne) = energie_fatale_kWhj(I)
    Next I
    Feuil41.Range("thermique_vapeur_fatale_kgj").Cells(1, Colonne) = vapeur_fatale_kgj
    Feuil41.Range("thermique_autosuffisance").Cells(1, Colonne) = autosuffisance_thermique


    If biogaz_revendu_rate <> 0 Then
        Feuil41.Range("thermique_biogaz_revendu_rate").Cells(1, Colonne) = biogaz_revendu_rate
    Else
        Feuil41.Range("thermique_biogaz_revendu_rate").Cells(1, Colonne) = "'-"       'biogaz_revendu_rate
    End If

    If biogaz_sechage_rate = 0 Then
        Feuil41.Range("thermique_biogaz_sechage_rate").Cells(1, Colonne) = "'-"    'biogaz_sechage_rate
    End If

    If biogaz_caracteristiques(Q_Nm3j) = 0 Then
        Feuil41.Range("thermique_biogaz_chaudiere_rdt").Cells(1, Colonne) = "'-"    'biogaz_chaudiere_rdt
        Feuil41.Range("thermique_biogaz_CHP_type").Cells(1, Colonne) = "'-"    'Feuil41.Range("thermique_no_CHP").Cells(1, 1)
        Feuil41.Range("thermique_biogaz_CHP_rate").Cells(1, Colonne) = "'-"    'biogaz_CHP_rate
        Feuil41.Range("thermique_biogaz_chaudiere_rate").Cells(1, Colonne) = "'-"    'biogaz_chaudiere_rate
        Feuil41.Range("thermique_biogaz_torchere_rate").Cells(1, Colonne) = "'-"    'biogaz_torchere_rate
        Feuil41.Range("thermique_biogaz_sechage_rate").Cells(1, Colonne) = "'-"    'biogaz_sechage_rate
        Feuil41.Range("thermique_biogaz_caracteristiques").Cells(teneur_CH4, Colonne) = "'-"
    End If
    If combustible_flux > 0 Then
        Feuil41.Range("thermique_combustible_chaudiere_rdt").Cells(1, Colonne) = combustible_chaudiere_rdt
        Feuil41.Range("thermique_combustible_chaudiere_Tch").Cells(1, Colonne) = combustible_chaudiere_Tch
        Feuil41.Range("thermique_combustible_chaudiere_Tfr").Cells(1, Colonne) = combustible_chaudiere_Tfr
        Feuil41.Range("thermique_combustible_flux").Cells(1, Colonne) = combustible_flux
    Else
        Feuil41.Range("thermique_combustible_chaudiere_rdt").Cells(1, Colonne) = "'-"    'combustible_chaudiere_rdt
        Feuil41.Range("thermique_combustible_chaudiere_Tch").Cells(1, Colonne) = "'-"     'combustible_chaudiere_Tch
        Feuil41.Range("thermique_combustible_chaudiere_Tfr").Cells(1, Colonne) = "'-"     'combustible_chaudiere_Tfr
        Feuil41.Range("thermique_combustible_flux").Cells(1, Colonne) = "'-"    'combustible_flux
    End If
    If combustible_sechage = 0 Then
        Feuil41.Range("thermique_combustible_sechage").Cells(1, Colonne) = "'-"    'combustible_sechage
    End If

    Call ecriture_resultats_OPEX


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub ecriture_resultats_OPEX()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer
    Dim I As Integer

    Colonne = 3 * pi_CALCUL_VALEUR_GUIDE + pi_FiliereConsideree

    For I = 1 To combustibles_nb
        If I = choix_combustible Then
            Feuil55.Range("OPEX_combustibles_fixe_kWhj").Cells(I, Colonne) = 0
            Feuil55.Range("OPEX_combustibles_variable_kWhj").Cells(I, Colonne) = (combustible_flux + combustible_sechage) * PCI(choix_combustible)
        Else
            Feuil55.Range("OPEX_combustibles_fixe_kWhj").Cells(I, Colonne) = 0
            Feuil55.Range("OPEX_combustibles_variable_kWhj").Cells(I, Colonne) = 0
        End If
    Next I
    Feuil55.Range("OPEX_biogaz_revendu_kWhj").Cells(1, Colonne) = biogaz_revendu_rate * PCI_CH4 * biogaz_caracteristiques(Q_Nm3j) * biogaz_caracteristiques(teneur_CH4)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.11"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats_OPEX()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim Colonne As Integer


    Colonne = 3 * pi_CALCUL_VALEUR_GUIDE + pi_FiliereConsideree

    For I = 1 To combustibles_nb
        If I = choix_combustible Then
            Feuil55.Range("OPEX_combustibles_fixe_kWhj").Cells(I, Colonne) = ""
            Feuil55.Range("OPEX_combustibles_variable_kWhj").Cells(I, Colonne) = ""
        Else
            Feuil55.Range("OPEX_combustibles_fixe_kWhj").Cells(I, Colonne) = ""
            Feuil55.Range("OPEX_combustibles_variable_kWhj").Cells(I, Colonne) = ""
        End If
    Next I
    Feuil55.Range("OPEX_biogaz_revendu_kWhj").Cells(1, Colonne) = ""      'biogaz_revendu_rate * PCI_CH4 * biogaz_caracteristiques(Q_Nm3j) * biogaz_caracteristiques(teneur_CH4)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.12"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    For I = 1 To nb_niveaux_energie
        Feuil41.Range("thermique_besoins_thermiques_kWhj").Cells(I, Colonne) = ""    'besoins_thermiques(i)
    Next I
    For I = 1 To nb_vapeur_max
        Feuil41.Range("thermique_pression_vapeur").Cells(I, Colonne) = ""    'pression_vapeur(i)
    Next I
    For I = 1 To nb_vapeur_max
        Feuil41.Range("thermique_besoins_vapeur_kgj").Cells(I, Colonne) = ""    'besoins_vapeur_kgj(i)
    Next I
    For I = 1 To energie_vapeur_global
        Feuil41.Range("thermique_disponibilites_thermiques_kWhj").Cells(I, Colonne) = ""    'disponibilites_thermiques(i)
    Next I
    For I = 1 To nb_biogaz_caracteristiques
        Feuil41.Range("thermique_biogaz_caracteristiques").Cells(I, Colonne) = ""    'biogaz_caracteristiques(i)
    Next I
    Feuil41.Range("thermique_biogaz_chaudiere_rdt").Cells(1, Colonne) = ""    'biogaz_chaudiere_rdt
    Feuil41.Range("thermique_biogaz_chaudiere_Tch").Cells(1, Colonne) = ""
    Feuil41.Range("thermique_biogaz_chaudiere_Tfr").Cells(1, Colonne) = ""
    Feuil41.Range("thermique_biogaz_CHP_type").Cells(1, Colonne) = ""    'Feuil41.Range("thermique_microturbine").Cells(1, 1)
    Feuil41.Range("thermique_biogaz_CHP_rdt_elec").Cells(1, Colonne) = ""    'biogaz_CHP_rdt_elec
    Feuil41.Range("thermique_biogaz_CHP_rdt_eau_BT").Cells(1, Colonne) = ""    'biogaz_CHP_rdt_eau_BT
    Feuil41.Range("thermique_biogaz_CHP_BT_Tch").Cells(1, Colonne) = ""     'biogaz_CHP_BT_Tch
    Feuil41.Range("thermique_biogaz_CHP_BT_Tfr").Cells(1, Colonne) = ""     'biogaz_CHP_BT_Tfr
    Feuil41.Range("thermique_biogaz_CHP_rdt_eau_HT").Cells(1, Colonne) = ""    'biogaz_CHP_rdt_eau_HT
    Feuil41.Range("thermique_biogaz_CHP_HT_Tch").Cells(1, Colonne) = ""     'biogaz_CHP_HT_Tch
    Feuil41.Range("thermique_biogaz_CHP_HT_Tfr").Cells(1, Colonne) = ""     'biogaz_CHP_HT_Tfr
    Feuil41.Range("thermique_biogaz_CHP_rdt_vapeur").Cells(1, Colonne) = ""    'biogaz_CHP_rdt_vapeur
    Feuil41.Range("thermique_biogaz_CHP_fumees_Tch").Cells(1, Colonne) = ""     'biogaz_CHP_fumees_Tch
    Feuil41.Range("thermique_biogaz_CHP_fumees_Tfr").Cells(1, Colonne) = ""     'biogaz_CHP_fumees_Tfr
    Feuil41.Range("thermique_biogaz_CHP_Pw").Cells(1, Colonne) = ""    'biogaz_CHP_Pw
    Feuil41.Range("thermique_biogaz_CHP_taux_charge").Cells(1, Colonne) = ""    'biogaz_CHP_Pw
    Feuil41.Range("thermique_biogaz_revendu_rate").Cells(1, Colonne) = ""       'biogaz_revendu_rate
    Feuil41.Range("thermique_biogaz_chaudiere_rate").Cells(1, Colonne) = ""    'biogaz_chaudiere_rate
    Feuil41.Range("thermique_biogaz_CHP_rate").Cells(1, Colonne) = ""    'biogaz_CHP_rate
    Feuil41.Range("thermique_biogaz_torchere_rate").Cells(1, Colonne) = ""    'biogaz_torchere_rate
    Feuil41.Range("thermique_biogaz_sechage_rate").Cells(1, Colonne) = ""    'thermique_biogaz_sechage_rate
    Feuil41.Range("thermique_combustible_chaudiere_rdt").Cells(1, Colonne) = ""    'combustible_chaudiere_rdt
    Feuil41.Range("thermique_combustible_chaudiere_Tch").Cells(1, Colonne) = ""     'combustible_chaudiere_Tch
    Feuil41.Range("thermique_combustible_chaudiere_Tfr").Cells(1, Colonne) = ""     'combustible_chaudiere_Tfr
    Feuil41.Range("thermique_combustible_flux").Cells(1, Colonne) = ""    'combustible_flux
    Feuil41.Range("thermique_combustible_sechage").Cells(1, Colonne) = ""    'combustible_sechage
    For I = 1 To energie_vapeur_global
        Feuil41.Range("thermique_energie_fatale_kWhj").Cells(I, Colonne) = ""    'energie_fatale_kWhj(i)
    Next I
    Feuil41.Range("thermique_vapeur_fatale_kgj").Cells(1, Colonne) = ""    'vapeur_fatale_kgj
    Feuil41.Range("thermique_autosuffisance").Cells(1, Colonne) = ""    'autosuffisance_thermique

    'Feuil41.Range("thermique_electricite_alternatif").Cells(1, colonne) = ""     'electricite_alternatif
    Feuil41.Range("thermique_electricite_biogaz").Cells(1, Colonne) = ""    'electricite_biogaz

    Call effacer_resultats_OPEX

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'biogaz vendu dans le réseau
    If biogaz_revendu_rate_force = True Then
        biogaz_revendu_rate = biogaz_revendu_rate_f
    Else
        biogaz_revendu_rate = 0
    End If

    'chaudière biogaz
    If biogaz_chaudiere_rdt_force = True Then
        biogaz_chaudiere_rdt = biogaz_chaudiere_rdt_f
    Else
        biogaz_chaudiere_rdt = rdt_thermique(chaudiere, energie_vapeur_global)
    End If
    If biogaz_chaudiere_Tch_force = True Then
        biogaz_chaudiere_Tch = biogaz_chaudiere_Tch_f
    Else
        biogaz_chaudiere_Tch = 800
    End If
    If biogaz_chaudiere_Tfr_force = True Then
        biogaz_chaudiere_Tfr = biogaz_chaudiere_Tfr_f
    Else
        biogaz_chaudiere_Tfr = 200
    End If
    
    'boucle BT de la CHP
    If biogaz_CHP_BT_Tch_force = True Then
        biogaz_CHP_BT_Tch = biogaz_CHP_BT_Tch_f
    Else
        biogaz_CHP_BT_Tch = 60
    End If
    If biogaz_CHP_BT_Tfr_force = True Then
        biogaz_CHP_BT_Tfr = biogaz_CHP_BT_Tfr_f
    Else
        biogaz_CHP_BT_Tfr = 40
    End If
    
    'boucle HT de la CHP
    If biogaz_CHP_HT_Tch_force = True Then
        biogaz_CHP_HT_Tch = biogaz_CHP_HT_Tch_f
    Else
        biogaz_CHP_HT_Tch = 90
    End If
    If biogaz_CHP_HT_Tfr_force = True Then
        biogaz_CHP_HT_Tfr = biogaz_CHP_HT_Tfr_f
    Else
        biogaz_CHP_HT_Tfr = 70
    End If
    
    'fumées de la CHP
    If biogaz_CHP_fumees_Tch_force = True Then
        biogaz_CHP_fumees_Tch = biogaz_CHP_fumees_Tch_f
    Else
        biogaz_CHP_fumees_Tch = 400
    End If
    If biogaz_CHP_fumees_Tfr_force = True Then
        biogaz_CHP_fumees_Tfr = biogaz_CHP_fumees_Tfr_f
    Else
        biogaz_CHP_fumees_Tfr = 200
    End If

    'chaudière combustible d'appoint
    If combustible_chaudiere_rdt_force = True Then
        combustible_chaudiere_rdt = combustible_chaudiere_rdt_f
    Else
        combustible_chaudiere_rdt = rdt_thermique(chaudiere, energie_vapeur_global)
    End If
    If combustible_chaudiere_Tch_force = True Then
        combustible_chaudiere_Tch = combustible_chaudiere_Tch_f
    Else
        combustible_chaudiere_Tch = 800
    End If
    If combustible_chaudiere_Tfr_force = True Then
        combustible_chaudiere_Tfr = combustible_chaudiere_Tfr_f
    Else
        combustible_chaudiere_Tfr = 200
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub



Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim I As Integer
    Dim j As Integer


    'REFERENCE DES ENTHALPIES: EAU LIQUIDE A 0°C ET 1atm


    'PCI du CH4  (kWh/Nm3)   pour calcul PCI du biogaz
    PCI_CH4 = pd_PCI_CH4_kWh_Nm3

    'PCI des autres combustibles
    PCI(combustible_repere_gaz_naturel) = pd_PCI_GAZ_NATUREL_kWh_Nm3
    PCI(combustible_repere_fioul) = pd_PCI_FIOUL_kWh_L

    'coefficients nécessaires au calcul de la température de vapeur saturée à P   (en °C)
    '    Tvap_sat(P) = a1/(a2 - ln(a3*P)) + a4 - a5   avec P = pression absolue en bar
    a_Tvap_sat(1) = 3816.44
    a_Tvap_sat(2) = 18.3
    a_Tvap_sat(3) = 750
    a_Tvap_sat(4) = 46.13
    a_Tvap_sat(5) = 273.15

    'coefficients pour calcul de l'enthalpie totale de la vapeur saturée à T   (en kcal/kg)  Hvap_sat(T) = a0 + a1*T
    a_Hvap_sat(0) = 613
    a_Hvap_sat(1) = 0.2806

    'rendements de base pour élec et thermique pour chaudière, microturbine et moteur   ATTENTION: ici les deux niveaux de vapeur sont traités ensemble
    For I = 1 To energie_vapeur_global
        For j = 1 To nb_technologies
            rdt_thermique(j, I) = 0
        Next j
    Next I
    'chaudière
    rdt_thermique(chaudiere, energie_vapeur_global) = 0.85
    'microturbine
    rdt_thermique(micro_turbine, energie_eau_BT) = 0
    rdt_thermique(micro_turbine, energie_eau_HT) = 0.4
    rdt_thermique(micro_turbine, energie_vapeur_global) = 0
    'moteur
    rdt_thermique(moteur, energie_eau_BT) = 0
    rdt_thermique(moteur, energie_eau_HT) = 0.2
    rdt_thermique(moteur, energie_vapeur_global) = 0.25

    'rendements électriques
    rdt_electrique(micro_turbine) = 0.25
    rdt_electrique(moteur) = 0.38
    'limit de microturbine ou moteur 'kW
    limit_min_moteur = 300

    'répartitions limites du biogaz
    biogaz_CHP_rate_maxi = 0.95
    biogaz_chaudiere_rate_mini = 0.01
    biogaz_torchere_rate_mini = 0.04

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub


Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String




    choix_objectif = Feuil41.Range("thermique_choix_objectif").Cells(1, pi_FiliereConsideree)
    If choix_objectif = objectif_autosuffisance_thermique Then
        choix_cogeneration = Feuil41.Range("thermique_choix_cogeneration").Cells(1, pi_FiliereConsideree)
    Else
        choix_cogeneration = cogeneration_oui
    End If
    choix_combustible = Feuil41.Range("thermique_choix_combustible").Cells(1, pi_FiliereConsideree)
    choix_combustible_utilisation = Feuil41.Range("thermique_choix_combustible_utilisation").Cells(1, pi_FiliereConsideree)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer

    Colonne = pi_FiliereConsideree

    biogaz_chaudiere_rdt_force = Feuil41.Range("thermique_biogaz_chaudiere_rdt").Cells(1, Colonne)
    biogaz_chaudiere_Tch_force = Feuil41.Range("thermique_biogaz_chaudiere_Tch").Cells(1, Colonne)
    biogaz_chaudiere_Tfr_force = Feuil41.Range("thermique_biogaz_chaudiere_Tfr").Cells(1, Colonne)
    If choix_cogeneration = cogeneration_oui Then
        biogaz_CHP_rdt_elec_force = Feuil41.Range("thermique_biogaz_CHP_rdt_elec").Cells(1, Colonne)
        biogaz_CHP_rdt_eau_BT_force = Feuil41.Range("thermique_biogaz_CHP_rdt_eau_BT").Cells(1, Colonne)
        biogaz_CHP_BT_Tch_force = Feuil41.Range("thermique_biogaz_CHP_BT_Tch").Cells(1, Colonne)
        biogaz_CHP_BT_Tfr_force = Feuil41.Range("thermique_biogaz_CHP_BT_Tfr").Cells(1, Colonne)
        biogaz_CHP_rdt_eau_HT_force = Feuil41.Range("thermique_biogaz_CHP_rdt_eau_HT").Cells(1, Colonne)
        biogaz_CHP_HT_Tch_force = Feuil41.Range("thermique_biogaz_CHP_HT_Tch").Cells(1, Colonne)
        biogaz_CHP_HT_Tfr_force = Feuil41.Range("thermique_biogaz_CHP_HT_Tfr").Cells(1, Colonne)
        biogaz_CHP_rdt_vapeur_force = Feuil41.Range("thermique_biogaz_CHP_rdt_vapeur").Cells(1, Colonne)
        biogaz_CHP_fumees_Tch_force = Feuil41.Range("thermique_biogaz_CHP_fumees_Tch").Cells(1, Colonne)
        biogaz_CHP_fumees_Tfr_force = Feuil41.Range("thermique_biogaz_CHP_fumees_Tfr").Cells(1, Colonne)
        biogaz_CHP_Pw_force = Feuil41.Range("thermique_biogaz_CHP_Pw").Cells(1, Colonne)
        biogaz_CHP_rate_force = Feuil41.Range("thermique_biogaz_CHP_rate").Cells(1, Colonne)
    Else
        biogaz_CHP_rdt_elec_force = False
        biogaz_CHP_rdt_eau_BT_force = False
        biogaz_CHP_BT_Tch_force = False
        biogaz_CHP_BT_Tfr_force = False
        biogaz_CHP_rdt_eau_HT_force = False
        biogaz_CHP_HT_Tch_force = False
        biogaz_CHP_HT_Tfr_force = False
        biogaz_CHP_rdt_vapeur_force = False
        biogaz_CHP_fumees_Tch_force = False
        biogaz_CHP_fumees_Tfr_force = False
        biogaz_CHP_Pw_force = False
        biogaz_CHP_rate_force = False
    End If
    biogaz_revendu_rate_force = Feuil41.Range("thermique_biogaz_revendu_rate").Cells(1, Colonne)
    biogaz_chaudiere_rate_force = Feuil41.Range("thermique_biogaz_chaudiere_rate").Cells(1, Colonne)
    biogaz_torchere_rate_force = Feuil41.Range("thermique_biogaz_torchere_rate").Cells(1, Colonne)
    biogaz_sechage_rate_force = Feuil41.Range("thermique_biogaz_sechage_rate").Cells(1, Colonne)
    combustible_chaudiere_rdt_force = Feuil41.Range("thermique_combustible_chaudiere_rdt").Cells(1, Colonne)
    combustible_chaudiere_Tch_force = Feuil41.Range("thermique_combustible_chaudiere_Tch").Cells(1, Colonne)
    combustible_chaudiere_Tfr_force = Feuil41.Range("thermique_combustible_chaudiere_Tfr").Cells(1, Colonne)


    If biogaz_chaudiere_rdt_force = True Then
        biogaz_chaudiere_rdt_f = Feuil41.Range("thermique_biogaz_chaudiere_rdt").Cells(1, Colonne)
    End If
    If biogaz_chaudiere_Tch_force = True Then
        biogaz_chaudiere_Tch_f = Feuil41.Range("thermique_biogaz_chaudiere_Tch").Cells(1, Colonne)
    End If
    If biogaz_chaudiere_Tfr_force = True Then
        biogaz_chaudiere_Tfr_f = Feuil41.Range("thermique_biogaz_chaudiere_Tfr").Cells(1, Colonne)
    End If
    If biogaz_CHP_rdt_elec_force = True Then
        biogaz_CHP_rdt_elec_f = Feuil41.Range("thermique_biogaz_CHP_rdt_elec").Cells(1, Colonne)
    End If
    If biogaz_CHP_rdt_eau_BT_force = True Then
        biogaz_CHP_rdt_eau_BT_f = Feuil41.Range("thermique_biogaz_CHP_rdt_eau_BT").Cells(1, Colonne)
    End If
    If biogaz_CHP_BT_Tch_force = True Then
        biogaz_CHP_BT_Tch_f = Feuil41.Range("thermique_biogaz_CHP_BT_Tch").Cells(1, Colonne)
    End If
    If biogaz_CHP_BT_Tfr_force = True Then
        biogaz_CHP_BT_Tfr_f = Feuil41.Range("thermique_biogaz_CHP_BT_Tfr").Cells(1, Colonne)
    End If
    If biogaz_CHP_rdt_eau_HT_force = True Then
        biogaz_CHP_rdt_eau_HT_f = Feuil41.Range("thermique_biogaz_CHP_rdt_eau_HT").Cells(1, Colonne)
    End If
    If biogaz_CHP_HT_Tch_force = True Then
        biogaz_CHP_HT_Tch_f = Feuil41.Range("thermique_biogaz_CHP_HT_Tch").Cells(1, Colonne)
    End If
    If biogaz_CHP_HT_Tfr_force = True Then
        biogaz_CHP_HT_Tfr_f = Feuil41.Range("thermique_biogaz_CHP_HT_Tfr").Cells(1, Colonne)
    End If
    If biogaz_CHP_rdt_vapeur_force = True Then
        biogaz_CHP_rdt_vapeur_f = Feuil41.Range("thermique_biogaz_CHP_rdt_vapeur").Cells(1, Colonne)
    End If
    If biogaz_CHP_fumees_Tch_force = True Then
        biogaz_CHP_fumees_Tch_f = Feuil41.Range("thermique_biogaz_CHP_fumees_Tch").Cells(1, Colonne)
    End If
    If biogaz_CHP_fumees_Tfr_force = True Then
        biogaz_CHP_fumees_Tfr_f = Feuil41.Range("thermique_biogaz_CHP_fumees_Tfr").Cells(1, Colonne)
    End If
    If biogaz_CHP_Pw_force = True Then
        biogaz_CHP_Pw_f = Feuil41.Range("thermique_biogaz_CHP_Pw").Cells(1, Colonne)
    End If
    If biogaz_revendu_rate_force = True Then
        biogaz_revendu_rate_f = Feuil41.Range("thermique_biogaz_revendu_rate").Cells(1, Colonne)
    End If
    If biogaz_chaudiere_rate_force = True Then
        biogaz_chaudiere_rate_f = Feuil41.Range("thermique_biogaz_chaudiere_rate").Cells(1, Colonne)
    End If
    If biogaz_CHP_rate_force = True Then
        biogaz_CHP_rate_f = Feuil41.Range("thermique_biogaz_CHP_rate").Cells(1, Colonne)
    End If
    If biogaz_torchere_rate_force = True Then
        biogaz_torchere_rate_f = Feuil41.Range("thermique_biogaz_torchere_rate").Cells(1, Colonne)
    End If
    If combustible_chaudiere_rdt_force = True Then
        combustible_chaudiere_rdt_f = Feuil41.Range("thermique_combustible_chaudiere_rdt").Cells(1, Colonne)
    End If
    If combustible_chaudiere_Tch_force = True Then
        combustible_chaudiere_Tch_f = Feuil41.Range("thermique_combustible_chaudiere_Tch").Cells(1, Colonne)
    End If
    If combustible_chaudiere_Tfr_force = True Then
        combustible_chaudiere_Tfr_f = Feuil41.Range("thermique_combustible_chaudiere_Tfr").Cells(1, Colonne)
    End If
    If biogaz_sechage_rate_force = True Then
        biogaz_sechage_rate_f = Feuil41.Range("thermique_biogaz_sechage_rate").Cells(1, Colonne)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub











'Sub PINCH()
'
'
'Dim energie_primaire_biogaz As Double
'Dim energie_primaire_biogaz_after_sechage As Double
'
'
''calcul de l'énergie primaire du biogaz
'energie_primaire_biogaz = PCI_CH4 * biogaz_caracteristiques(Q_Nm3j) * biogaz_caracteristiques(teneur_CH4)
'
'    If energie_primaire_biogaz > 0 Then
'
'        '!!! cela intègre peut le pyrofluid
'        'on calcule l'énergie en cas de sécheur
'        If besoins_thermiques(energie_biogaz) > 0 Then    'Si il y a un besoin thermique sous forme de biogaz pour le séchage : biogaz directement utilisé dans le sécheur
'
'            If biogaz_sechage_rate_force = True Then ' si la valeur du taux de biogaz utilisée dans le sécheur a été forcée
'                biogaz_sechage_rate = biogaz_sechage_rate_f 'récupération de la valeur forcée
'                energie_primaire_biogaz_after_sechage = energie_primaire_biogaz * (1 - biogaz_sechage_rate) ' calcul du nouveau taux
'                biogaz_sechage_non_couvert = besoins_thermiques(energie_biogaz) - energie_primaire_biogaz * biogaz_sechage_rate 'calcul de la différence entre les besoins et l'énergie utilisée
'
'            Else ' sinon la valeur du taux de biogaz utilisée dans le sécheur n'a pas été forcée
'
'                'on regarde la répatition des taux qui ont été forcés
'                If biogaz_torchere_rate_force = True Then ' si le taux de répartition vers la torchère a été forcée
'                    biogaz_torchere_rate = biogaz_torchere_rate_f
'                    biogaz_sechage_rate_max = 1 - (biogaz_torchere_rate_f + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f) ' maximum utiisable en sécheur après avoir forcé les valeurs
'                Else
'                    biogaz_torchere_rate = biogaz_torchere_rate_mini
'                    biogaz_sechage_rate_max = 1 - (biogaz_torchere_rate + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f)
'                End If
'
'                'attribution de la part de biogaz vers le sécheur --> calcul de biogaz_sechage_rate, energie_primaire_biogaz_after_sechage, biogaz_sechage_non_couvert
'                If besoins_thermiques(energie_biogaz) > energie_primaire_biogaz * biogaz_sechage_rate_max Then 'le biogaz est insuffisant
'                    biogaz_sechage_rate = biogaz_sechage_rate_max
'                    energie_primaire_biogaz_after_sechage = energie_primaire_biogaz * (1 - biogaz_sechage_rate) ' biogaz total moins la part consacrée au séchage
'                    'on ajoute ce qui manque aux besoins en combustible
'                    biogaz_sechage_non_couvert = besoins_thermiques(energie_biogaz) - energie_primaire_biogaz * biogaz_sechage_rate ' part des besoins en séchage non couverts avec le biogaz consacré
'                Else    'il y a suffisamment de biogaz pour le séchage
'                    If biogaz_chaudiere_rate_force = True And biogaz_CHP_rate_force = True And biogaz_torchere_rate_force = True Then ' si tout est forcé sauf le séchage
'                        'comme la valeur est fixée on l'autoforce
'                        biogaz_sechage_rate_force = True
'                        biogaz_sechage_rate_f = 1 - (biogaz_torchere_rate_f + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f)
'                        biogaz_sechage_rate = biogaz_sechage_rate_f
'                    Else 'si d'autres valeurs que le séchages ne sont pas forcée (au moins une)
'                        biogaz_sechage_rate = besoins_thermiques(energie_biogaz) / energie_primaire_biogaz
'                    End If
'                    energie_primaire_biogaz_after_sechage = energie_primaire_biogaz * (1 - biogaz_sechage_rate) ' calcul de la qté de biogaz restant (cf biogaz_sechage_rate)
'                End If
'            End If
'        Else ' pas de besoins thermiques sous forme de biogaz pour le séchage
'            energie_primaire_biogaz_after_sechage = energie_primaire_biogaz ' on a la qté d'énergie produite à partir de biogaz initiale
'        End If
'
'        'interprétation de la valeur de biogaz_sechage_non_couvert
'        If biogaz_sechage_non_couvert < 0 Then
'            biogaz_sechage_non_couvert = 0 'remise à zéro
'            remarque(5) = Feuil41.Range("thermique_erreur_biogaz_sechage_trop").Cells(1, 1) ' taux de répartition vers le séchage trop important (2 cas: taux séchage forcé ou tous les autres taux sauf séchage forcées)
'        ElseIf biogaz_sechage_non_couvert > 0 Then
'            remarque(6) = Feuil41.Range("thermique_erreur_biogaz_sechage_insuffisant").Cells(1, 1) 'taux de répartition vers le séchage trop faible (3 cas:taux séchage forcé; tous les autres taux sauf séchage forcées; 1 ou 2 valeurs autres que le sécheur on été forcées et l'énergie restante est trop faible pour couvrir les besoins du sécheur(la différence [100% - leurs somme] est trop faible)
'        End If
'
'        'on calcule l'énergie primaire de combustible à apporter pour le séchage
'        If besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert > 0 Then
'            combustible_sechage = (besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / PCI(choix_combustible)
'        Else
'            combustible_sechage = 0
'        End If
'
'
'        'On rentre dans le vif du sujet avec chaudières et cogénérations
'        If choix_objectif = objectif_autosuffisance_thermique Then    'AUTOSUFFISANCE THERMIQUE MAXIMALE SANS COGENERATION
'
'            If choix_cogeneration = cogeneration_non Then   'seulement chaudières                             'fait le 6-7/06/2011 DCA
'                'on veut connaître la répartition du biogaz
'                    'on calcule l'ensemble des besoins totaux (indépendemment du niveau)
'                    besoins_restants = 0
'                    For i = 1 To energie_vapeur_global
'                        besoins_restants = besoins_restants + besoins(i)
'                    Next i
'                If biogaz_chaudiere_rate_force = True Then
'                    biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
'                    If biogaz_torchere_rate_force = True Then
'                        biogaz_torchere_rate = biogaz_torchere_rate_f
'                    Else
'                        biogaz_torchere_rate = (1 - biogaz_sechage_rate) - biogaz_chaudiere_rate
'                    End If
'                ElseIf biogaz_torchere_rate_force = True Then
'                    biogaz_torchere_rate = biogaz_torchere_rate_f
'                    biogaz_chaudiere_rate = (1 - biogaz_sechage_rate) - biogaz_torchere_rate
'                Else
'                    'on calcule le potentiel d'énergie sur la chaudière
'                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rdt * ((1 - biogaz_sechage_rate) - biogaz_torchere_rate_mini)
'                    If besoins_restants <= potentiel_energie Then   'énergie suffisante pour couvrir l'ensemble des besoins liés au biogaz
'                        'on calcule la répartition du biogaz
'                        biogaz_chaudiere_rate = (besoins_restants / biogaz_chaudiere_rdt) / energie_primaire_biogaz
'                        biogaz_torchere_rate = (1 - biogaz_sechage_rate) - biogaz_chaudiere_rate
'                    Else
'                        biogaz_torchere_rate = biogaz_torchere_rate_mini
'                        biogaz_chaudiere_rate = (1 - biogaz_sechage_rate) - biogaz_torchere_rate
'                    End If
'                End If
'                'on va calculer les réactifs à ajouter si nécessaire
'                potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
'                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
'                autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
'
'            ElseIf choix_cogeneration = cogeneration_oui Then    'AUTOSUFFISANCE THERMIQUE MAXIMALE AVEC COGENERATION
'                If biogaz_CHP_rate_force = True Then                                                         'fait par KME puis revu le 15/06/2011 par DCA
'                    biogaz_CHP_rate = biogaz_CHP_rate_f
'                    'on calcule d'abord les besoins restants car utile pour calculer la répartition chaudière-torchère lorsque nécessaire
'                    'type de cogénération installable
'                    energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
'                    Call definition_type_CHP(energie_primaire_CHP)
'                        rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
'                        rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
'                        rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
'                    'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
'                        Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
'                    'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
'                    besoins_restants = 0 'raz
'                    For i = 1 To energie_vapeur_global 'somme des besoins après répartition de l'énergie en CHP
'                        besoins_restants = besoins_restants + besoins(i)
'                    Next i
'
'                    'on regarde les autres valeurs pour voir si elles sont forcées (2²=4 cas -> chaudière et torchère, forcées ou non)
'                    If biogaz_chaudiere_rate_force = True Then ' taux chaudière forcé
'                        biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
'                        If biogaz_torchere_rate_force = True Then 'taux torchère forcé
'                            biogaz_torchere_rate = biogaz_torchere_rate_f
'                        Else ' taux torchère non forcé, on prend ce qu'il reste
'                            biogaz_torchere_rate = (1 - biogaz_sechage_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
'                        End If
'                    Else 'taux chaudière non forcé
'                        If biogaz_torchere_rate_force = True Then
'                            biogaz_torchere_rate = biogaz_torchere_rate_f
'                            biogaz_chaudiere_rate = (1 - biogaz_sechage_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
'                        Else
'                            If biogaz_CHP_rate > biogaz_CHP_rate_maxi Then ' le taux défini vers CHP est > au taux maximal (90 % dans les hypothèses) donc on répartit le reste entre chaudière entre les taux de chuadière et de torchère comme étant équivalent
'                                biogaz_torchere_rate = ((1 - biogaz_sechage_rate) - biogaz_CHP_rate) / 2 'on partage en 2 la chaudière et la torchère
'                                biogaz_chaudiere_rate = ((1 - biogaz_sechage_rate) - biogaz_CHP_rate) / 2 'on partage en 2 la chaudière et la torchère
'                            Else   'on répartit au mieux entre torchère et chaudière
'                                'énergie maximale que l'on peut atteindre avec la chaudière
'                                potentiel_energie = ((1 - biogaz_sechage_rate) - biogaz_CHP_rate - biogaz_torchere_rate_mini) * energie_primaire_biogaz * biogaz_chaudiere_rdt
'                                If potentiel_energie < besoins_restants Then   'on fixe la torchère au min et la chaudière prend ce qu'il reste
'                                    biogaz_torchere_rate = biogaz_torchere_rate_mini
'                                    biogaz_chaudiere_rate = (1 - biogaz_sechage_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
'                                Else   'on calcule la répartition permettant de couvrir exactement les besoins
'                                    biogaz_chaudiere_rate = (besoins_restants / biogaz_chaudiere_rdt) / energie_primaire_biogaz
'                                    biogaz_torchere_rate = (1 - biogaz_sechage_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
'                                End If
'                            End If
'                        End If
'                    End If
'                    'on va calculer les combustibles à ajouter si nécessaire
'                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
'                        Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
'                    'calcul du taux d'autosuffisance thermique--> [besoins couverts sans apport ext] / [besoins totaux]
'                    autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
'                    'calcul de l'énergie électrique produite
'                    energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
'
'                ElseIf biogaz_chaudiere_rate_force = True Then
'                    biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
'                    'on regarde si la torchère est fixée
'                    If biogaz_torchere_rate_force = True Then
'                        biogaz_torchere_rate = biogaz_torchere_rate_f
'                        biogaz_CHP_rate = (1 - biogaz_sechage_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
'                    Else
'                        If biogaz_chaudiere_rate > ((1 - biogaz_sechage_rate) - biogaz_torchere_rate_mini) Then
'                            biogaz_torchere_rate = (1 - biogaz_sechage_rate) - biogaz_chaudiere_rate
'                            biogaz_CHP_rate = 0
'                        ElseIf biogaz_chaudiere_rate < ((1 - biogaz_sechage_rate) - biogaz_torchere_rate_mini - biogaz_CHP_rate_maxi) Then 'pour ne pas dépasser le maximum utilisable en CHP
'                            biogaz_CHP_rate = biogaz_CHP_rate_maxi
'                            biogaz_torchere_rate = (1 - biogaz_sechage_rate) - (biogaz_chaudiere_rate + biogaz_CHP_rate)
'                        Else
'                            biogaz_torchere_rate = biogaz_torchere_rate_mini
'                            biogaz_CHP_rate = (1 - biogaz_sechage_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
'                        End If
'                    End If
'                    'type de cogénération installable
'                    energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
'                    Call definition_type_CHP(energie_primaire_CHP)
'                        rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
'                        rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
'                        rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
'                    'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
'                        Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
'                    'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
'                    besoins_restants = 0
'                    For i = 1 To energie_vapeur_global
'                        besoins_restants = besoins_restants + besoins(i)
'                    Next i
'                    'on va calculer les réactifs à ajouter si nécessaire
'                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
'                        Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
'                    autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
'                    'calcul de l'énergie électrique produite
'                    energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
'
'                Else
'                    If biogaz_torchere_rate_force = True Then
'                        biogaz_torchere_rate = biogaz_torchere_rate_f
'                    Else
'                        biogaz_torchere_rate = biogaz_torchere_rate_mini
'                    End If
'                    CHP_maxi = (1 - biogaz_sechage_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate_mini)
'                    If CHP_maxi < 0 Then
'                        CHP_maxi = 0
'                    End If
'                    If CHP_maxi > biogaz_CHP_rate_maxi Then    'si CHP_maxi>0.9
'                        CHP_maxi = biogaz_CHP_rate_maxi
'                    End If
'                    Call calcul_biogaz_repartition(CHP_maxi, besoins, dispo, energie_primaire_biogaz)
'                    'on calcule les besoins restants
'                    besoins_restants = 0
'                    For i = 1 To energie_vapeur_global
'                        besoins_restants = besoins_restants + besoins(i)
'                    Next i
'                    'on va calculer les réactifs à ajouter si nécessaire
'                    potentiel_energie = 0   'énergie chaudière déjà utilisée
'                        Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
'                    autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
'                    'calcul de l'énergie électrique produite
'                    energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
'                    energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
'                End If
'            End If
'
'
'        ElseIf choix_objectif = objectif_autosuffisance_electrique Then    'AUTOSUFFISANCE ELECTRIQUE MAXIMALE
'            If biogaz_CHP_rate_force = True Then                                                'fait par KME puis revu le 14/06/2011 par DCA
'                biogaz_CHP_rate = biogaz_CHP_rate_f
'                'on calcule d'abord les besoins restants car utile pour calculer la répartition chaudière-torchère lorsque nécessaire
'                'type de cogénération installable
'                energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
'                Call definition_type_CHP(energie_primaire_CHP)
'                    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
'                    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
'                    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
'                'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
'                    Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
'                'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
'                besoins_restants = 0
'                For i = 1 To energie_vapeur_global
'                    besoins_restants = besoins_restants + besoins(i)
'                Next i
'                'on regarde les autres valeurs pour voir si elles sont forcées
'                If biogaz_chaudiere_rate_force = True Then
'                    biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
'                    If biogaz_torchere_rate_force = True Then
'                        biogaz_torchere_rate = biogaz_torchere_rate_f
'                    Else
'                        biogaz_torchere_rate = (1 - biogaz_sechage_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
'                    End If
'                Else
'                    If biogaz_torchere_rate_force = True Then
'                        biogaz_torchere_rate = biogaz_torchere_rate_f
'                        biogaz_chaudiere_rate = (1 - biogaz_sechage_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
'                    Else
'                        If biogaz_CHP_rate > biogaz_CHP_rate_maxi Then
'                            biogaz_torchere_rate = ((1 - biogaz_sechage_rate) - biogaz_CHP_rate) / 2 'on partage en 2 la chaudière et la torchère
'                            biogaz_chaudiere_rate = ((1 - biogaz_sechage_rate) - biogaz_CHP_rate) / 2 'on partage en 2 la chaudière et la torchère
'                        Else   'on répartit au mieux entre torchère et chaudière
'                            'énergie maximale que l'on peut atteindre avec la chaudière
'                            potentiel_energie = ((1 - biogaz_sechage_rate) - biogaz_CHP_rate - biogaz_torchere_rate_mini) * energie_primaire_biogaz * biogaz_chaudiere_rdt
'                            If potentiel_energie < besoins_restants Then   'on fixe la torchère au min et la chaudière prend ce qu'il reste
'                                biogaz_torchere_rate = biogaz_torchere_rate_mini
'                                biogaz_chaudiere_rate = (1 - biogaz_sechage_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
'                            Else   'on calcule la répartition permettant de couvrir exactement les besoins
'                                biogaz_chaudiere_rate = (besoins_restants / biogaz_chaudiere_rdt) / energie_primaire_biogaz
'                                biogaz_torchere_rate = (1 - biogaz_sechage_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
'                            End If
'                        End If
'                    End If
'                End If
'                'on va calculer les réactifs à ajouter si nécessaire
'                potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
'                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
'                autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
'                'calcul de l'énergie électrique produite
'                energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
'            ElseIf biogaz_chaudiere_rate_force = True Then                                      'fait par KME puis revu le 15/06/2011 par DCA
'                biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
'                'on regarde si la torchère est fixée
'                If biogaz_torchere_rate_force = True Then
'                    biogaz_torchere_rate = biogaz_torchere_rate_f
'                    biogaz_CHP_rate = (1 - biogaz_sechage_rate) - (biogaz_chaudiere_rate + biogaz_torchere_rate)
'                Else
'                    If biogaz_chaudiere_rate > ((1 - biogaz_sechage_rate) - biogaz_torchere_rate_mini) Then 'il ne reste pas assez pour le minimum en torchère
'                        biogaz_torchere_rate = (1 - biogaz_sechage_rate) - biogaz_chaudiere_rate
'                        biogaz_CHP_rate = (1 - biogaz_sechage_rate) - (biogaz_chaudiere_rate + biogaz_torchere_rate)
'                    ElseIf biogaz_chaudiere_rate < ((1 - biogaz_sechage_rate) - biogaz_torchere_rate_mini - biogaz_CHP_rate_maxi) Then 'pour ne pas dépasser le maximum utilisable en CHP
'                        biogaz_CHP_rate = biogaz_CHP_rate_maxi
'                        biogaz_torchere_rate = (1 - biogaz_sechage_rate) - (biogaz_chaudiere_rate + biogaz_CHP_rate)
'                    Else
'                        biogaz_torchere_rate = biogaz_torchere_rate_mini
'                        biogaz_CHP_rate = (1 - biogaz_sechage_rate) - (biogaz_chaudiere_rate + biogaz_torchere_rate)
'                    End If
'                End If
'                'type de cogénération installable
'                energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
'                Call definition_type_CHP(energie_primaire_CHP)
'                    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
'                    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
'                    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
'                'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
'                    Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
'                'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
'                besoins_restants = 0
'                For i = 1 To energie_vapeur_global
'                    besoins_restants = besoins_restants + besoins(i)
'                Next i
'                'on va calculer les réactifs à ajouter si nécessaire
'                potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
'                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
'                autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
'                'calcul de l'énergie électrique produite
'                energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
'            ElseIf biogaz_torchere_rate_force = True Then                                      'fait par KME puis revu le 15/06/2011 par DCA
'                biogaz_torchere_rate = biogaz_torchere_rate_f
'                If biogaz_torchere_rate > ((1 - biogaz_sechage_rate) - biogaz_chaudiere_rate_mini) Then
'                    biogaz_chaudiere_rate = (1 - biogaz_sechage_rate) - biogaz_torchere_rate
'                    biogaz_CHP_rate = 0
'                ElseIf biogaz_torchere_rate < ((1 - biogaz_sechage_rate) - biogaz_chaudiere_rate_mini - biogaz_CHP_rate_maxi) Then
'                    biogaz_CHP_rate = biogaz_CHP_rate_maxi
'                    biogaz_chaudiere_rate = (1 - biogaz_sechage_rate) - (biogaz_torchere_rate + biogaz_CHP_rate)
'                Else
'                    biogaz_chaudiere_rate = biogaz_chaudiere_rate_mini
'                    biogaz_CHP_rate = (1 - biogaz_sechage_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
'                End If
'                'type de cogénération installable
'                energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
'                Call definition_type_CHP(energie_primaire_CHP)
'                    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
'                    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
'                    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
'                'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
'                    Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
'                'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
'                besoins_restants = 0
'                For i = 1 To energie_vapeur_global
'                    besoins_restants = besoins_restants + besoins(i)
'                Next i
'                'on va calculer les réactifs à ajouter si nécessaire
'                potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
'                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
'                autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
'                'calcul de l'énergie électrique produite
'                energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
'            Else                                                                                  'fait par KME puis revu le 15/06/2011 par DCA
'                biogaz_chaudiere_rate = biogaz_chaudiere_rate_mini
'                biogaz_torchere_rate = biogaz_torchere_rate_mini
'                biogaz_CHP_rate = (1 - biogaz_sechage_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
'                'type de cogénération installable
'                energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
'                Call definition_type_CHP(energie_primaire_CHP)
'                    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
'                    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
'                    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
'                'on calcule les disponibilités thermiques liées au moteur et on couvre les besoins que l'on peut
'                    Call calcul_dispo_besoins_CHP(besoins, dispo, energie_primaire_CHP, rdt_th_coge)
'                'on calcule les besoins restants pour ajuster le biogaz vers la chaudière lorsque nécessaire
'                besoins_restants = 0
'                For i = 1 To energie_vapeur_global
'                    besoins_restants = besoins_restants + besoins(i)
'                Next i
'                'on va calculer les réactifs à ajouter si nécessaire
'                potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rate * biogaz_chaudiere_rdt
'                    Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
'                autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
'                'calcul de l'énergie électrique produite
'                energie_elec = energie_primaire_CHP * biogaz_CHP_rdt_elec
'            End If
'        End If
'        'VERIFICATION DE LA REPARTITION
'        If Abs(biogaz_CHP_rate) + Abs(biogaz_chaudiere_rate) + Abs(biogaz_torchere_rate) + Abs(biogaz_sechage_rate) <> 1 Then
'            MsgBox Feuil2.Range("synthese_thermique").Cells(1, 1) & ": " & Feuil41.Range("thermique_erreur_repartition_globale").Cells(1, 1), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
'            pn_EffacementResultatsEnCours = True
'            pn_ErreurNonGeree = True
'            Exit Sub
'        End If
'        'Vérification des flux minimaux et maximaux
'        If iteration = pi_NOMBRE_ITERATION_RETOURS Then
'            If biogaz_chaudiere_rate < biogaz_chaudiere_rate_mini Then
'                remarque(1) = Feuil41.Range("thermique_erreur_repartition_mini_chaudiere").Cells(1, 1)
'            End If
'            If biogaz_torchere_rate < biogaz_torchere_rate_mini Then
'                remarque(2) = Feuil41.Range("thermique_erreur_repartition_mini_torchere").Cells(1, 1)
'            End If
'            If biogaz_CHP_rate > biogaz_CHP_rate_maxi Then
'                remarque(3) = Feuil41.Range("thermique_erreur_repartition_maxi_CHP").Cells(1, 1)
'            End If
'        End If
'    Else   'PAS DE BIOGAZ A REPARTIR
'        biogaz_torchere_rate = 0
'        biogaz_chaudiere_rate = 0
'        biogaz_CHP_rate = 0
'        biogaz_sechage_rate = 0
'        'on calcule les besoins restants pour savoir ce qu'il manque
'        besoins_restants = 0
'        For i = 1 To energie_vapeur_global
'            besoins_restants = besoins_restants + besoins(i)
'        Next i
'        potentiel_energie = 0       'ici pas d'énergie valorisable en chaudière
'        Call calcul_combustible(besoins_restants, potentiel_energie, besoins, dispo)
'        biogaz_sechage_non_couvert = besoins_thermiques(energie_biogaz)
'        If biogaz_sechage_non_couvert < 0 Then
'            biogaz_sechage_non_couvert = 0
'            remarque(5) = Feuil41.Range("thermique_erreur_biogaz_sechage_trop").Cells(1, 1)
'        Else
'            remarque(6) = Feuil41.Range("thermique_erreur_biogaz_sechage_insuffisant").Cells(1, 1)
'        End If
'        autosuffisance_thermique = 1 - (potentiel_energie * combustible_chaudiere_rdt + besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / besoins_totaux
'        'on calcule l'énergie primaire de combustible à apporter pour le séchage
'        If besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert > 0 Then
'            combustible_sechage = (besoins_thermiques(energie_combustible) + biogaz_sechage_non_couvert) / PCI(choix_combustible)
'        Else
'            combustible_sechage = 0
'        End If
'    End If
'Else
'    autosuffisance_thermique = 1
'End If
'
'End Sub

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








