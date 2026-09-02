Attribute VB_Name = "z_Gestion_energie_therm_PINCH"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'modif MSM 25/10/12 : sortir la production d'élec alternative de ce module

Dim pincement As PINCH

'constante pour PINCH pour lecture / ecriture
Private Const LigneTab = 3
Private Const critere_convergence = 0.001
Private Const zero_force As Double = 0.00001

'constantes technologies
Private Const nb_technologies = 3
Private Const chaudiere = 1
Private Const micro_turbine = 2
Private Const moteur = 3

'critère d'égalité
Private Const critere_egalite = 1
Private Const energie_vapeur_global = 3

'constantes objectifs
Private Const objectif_autosuffisance_thermique = 1
Private Const objectif_autosuffisance_electrique = 2

'constantes cogénération
Private Const cogeneration_oui = 1
Private Const cogeneration_non = 2

'Private Const electricite_alternatif_oui = 1 'modif MSM 25/10/12

'constantes combustible
Private Const combustible_utilisation_chaudiere = 1

'choix de l'utilisateur
Dim choix_objectif As Integer
Dim choix_cogeneration As Integer
Dim choix_combustible As Integer
Dim choix_combustible_utilisation As Integer
'Dim choix_electricite_alternatif As Integer'modif MSM 25/10/12

'variables à échanger avec l'utilisateur
Public besoins_thermiques_biogaz_incineration As Double
Public besoins_thermiques_combustible_incineration As Double
Public pd_BesoinsCombustibleExterneIncineration_kWhj As Double
Public pd_BesoinsCombustibleExterneHorsIncineration_kWhj As Double
Dim pression_vapeur(nb_vapeur_max) As Double
Dim besoins_vapeur_kgj(nb_vapeur_max) As Double
Dim disponibilites_thermiques(nb_niveaux_energie - nb_vapeur_max) As Double
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
Dim biogaz_incinerateur_rate As Double
Dim biogaz_torchere_rate As Double
Dim combustible_chaudiere_rdt As Double
Dim combustible_chaudiere_Tch As Double
Dim combustible_chaudiere_Tfr As Double
Dim combustible_flux As Double
Dim combustible_incinerateur As Double
Dim energie_fatale_kWhj(energie_vapeur_global) As Double
Dim vapeur_fatale_kgj As Double
Dim autosuffisance_thermique As Double
'Dim electricite_alternatif As Double 'modif MSM 25/10/12
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
Dim biogaz_incinerateur_rate_force As Boolean
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
Dim biogaz_incinerateur_rate_f As Double
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


'Flux d'énergie des utilités pour le Pinch

'Flux d'utilités renseignées par l'utilisateur
Dim nb_utilites_ajoutees As Integer
Dim utilites_ajoutees() As Variant


Dim remarque(6) As String
Public gaz_naturel_kWhPCIj As Double
Public fioul_kWhPCIj As Double


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


Sub dimensionnement(ByVal biogaz_caract, ByRef energie_elec, ByRef type_valorisation_biogaz, ByRef valorisation_chaudiere, ByRef ratio_biogaz_valorise)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer
    Dim Colonne As Integer

    'variables intermédiaires
    Dim besoins_totaux_HOT As Double
    Dim besoins_totaux_COLD As Double

    Dim energie_primaire_biogaz As Double
    Dim potentiel_energie As Double

    Dim energie_primaire_CHP As Double
    Dim CHP_maxi As Double

    Dim besoins_HOT As Double
    Dim besoins_COLD As Double

    Dim delta_H As Double

    Dim biogaz_incinerateur_rate_max As Double
    Dim biogaz_incinerateur_non_couvert As Double


    'On récupère ce qui nous est utile
    'caractéristiques biogaz
    For I = 1 To nb_biogaz_caracteristiques
        biogaz_caracteristiques(I) = biogaz_caract(I)
    Next I
    'Besoins totaux chaux
    besoins_totaux_HOT = 0
    besoins_totaux_COLD = 0
    For I = 1 To UBound(flux_thermique_process, 2)
        delta_H = CDbl(flux_thermique_process(repere_flux_thermique_process_MCp_kW_°C, I)) * (CDbl(flux_thermique_process(repere_flux_thermique_process_T_in_°C, I)) - CDbl(flux_thermique_process(repere_flux_thermique_process_T_out_°C, I)))
        If delta_H < 0 Then
            besoins_totaux_HOT = besoins_totaux_HOT + Abs(delta_H)
        Else
            besoins_totaux_COLD = besoins_totaux_COLD + Abs(delta_H)
        End If
    Next I
    besoins_totaux_HOT = besoins_totaux_HOT + (besoins_thermiques_biogaz_incineration + besoins_thermiques_combustible_incineration) / pd_NOMBRE_HEURE_PAR_JOUR
    besoins_totaux_COLD = besoins_totaux_COLD / pd_NOMBRE_HEURE_PAR_JOUR

    'on réinitialise les flux d'utilités
    'colonne = flux_thermique_utilites_nb_caracteristiques *  (pi_FiliereConsideree - 1)
    'i = 1
    'Do While Feuil63.Range("DonneesPincementFluxUtilites").Cells(ligneTab + i, colonne + repere_flux_thermique_utilites_etat) <> ""
    '    Feuil63.Range(Feuil63.Range("DonneesPincementFluxUtilites").Cells(ligneTab + i, colonne + repere_flux_thermique_utilites_etat), Feuil63.Range("DonneesPincementFluxUtilites").Cells(ligneTab + i, colonne + flux_thermique_utilites_nb_caracteristiques)).ClearContents
    '    i = i + 1
    'Loop

    'calcul de l'énergie primaire du biogaz
    energie_primaire_biogaz = PCI_CH4 * biogaz_caracteristiques(Q_Nm3j) * biogaz_caracteristiques(teneur_CH4)

    If (Abs(besoins_totaux_HOT) + Abs(besoins_totaux_COLD)) > 0 Then
        'ON COMMENCE PAR TRACER LES COURBES COMPOSITES
        '************ CREATION DU MODULE DE CLASSE "PINCH" ******************
        '************ APPEL DE LA SUB POUR TRACER LES COURBES COMPOSITES (récupérer les besoins chauds et froids restants)
        Set pincement = New PINCH
        Call pincement.VeriPinch(1)   '1 donc fait tout (CC, HEN & utilités)
        besoins_HOT = pincement.besoins_HOT
        If energie_primaire_biogaz > 0 Then
            If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                besoins_HOT = 0
            End If
            besoins_COLD = pincement.besoins_COLD
            If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                besoins_COLD = 0
            End If
        End If
        '************ APPEL DE LA SUB POUR FAIRE LE RESEAU D'ECHANGEURS DE CHALEURS
        'call pincement.HEN (besoins_HOT, besoins_COLD)   SANS les utilités
        '    --> Remarque: si les besoins chauds et froids renvoyés sont différents dans les deux cas alors le HEN n'est pas optimal
        '************ GESTION DES BESOINS DU SECHAGE ***********************
        '     --> On traitera dans un premier temps les besoins de séchage avec la chaudière, i.e. biogaz_chaudiere_rate correspond au biogaz vers la chaudière et le sécheur réunis
        '         --> On supprimerait alors les choix sur les combustibles utilisés dans les sécheurs (PAS DE SUITE, QUAND ON SERA FIXE A 100%)
        '         --> cela implique que l'utilisateur devra être en mesure de forcer des utilités avec des flux process ou que l'on trouve un moyen de contourner ce manque dans le code du PINCH
        'On lance la répartition du biogaz dans le cas où l'on en a
        If energie_primaire_biogaz > 0 Then
            'On enlève la part de biogaz utilisé dans l'incinérateur
            If besoins_thermiques_biogaz_incineration > 0 Then
                If biogaz_incinerateur_rate_force = True Then
                    biogaz_incinerateur_rate = biogaz_incinerateur_rate_f
                    '                energie_primaire_biogaz_after_sechage = energie_primaire_biogaz * (1 - biogaz_sechage_rate - biogaz_revendu_rate)
                    biogaz_incinerateur_non_couvert = besoins_thermiques_biogaz_incineration - energie_primaire_biogaz * biogaz_incinerateur_rate
                Else
                    If biogaz_torchere_rate_force = True Then
                        biogaz_torchere_rate = biogaz_torchere_rate_f
                        biogaz_incinerateur_rate_max = 1 - (biogaz_torchere_rate_f + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f + biogaz_revendu_rate)
                    Else
                        biogaz_incinerateur_rate_max = 1 - (biogaz_torchere_rate_mini + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f + biogaz_revendu_rate)
                        biogaz_torchere_rate = biogaz_torchere_rate_mini
                    End If
                    'besoins thermiques
                    If besoins_thermiques_biogaz_incineration > energie_primaire_biogaz * biogaz_incinerateur_rate_max Then    'le biogaz est insuffisant
                        biogaz_incinerateur_rate = biogaz_incinerateur_rate_max
                        '                    energie_primaire_biogaz_after_sechage = energie_primaire_biogaz * (1 - biogaz_sechage_rate - biogaz_revendu_rate)
                        'on ajoute ce qui manque aux besoins en combustible
                        biogaz_incinerateur_non_couvert = besoins_thermiques_biogaz_incineration - energie_primaire_biogaz * biogaz_incinerateur_rate
                    Else    'il y a suffisamment de biogaz pour le séchage
                        If biogaz_chaudiere_rate_force = True And biogaz_CHP_rate_force = True And biogaz_torchere_rate_force = True Then
                            biogaz_incinerateur_rate = 1 - (biogaz_torchere_rate_f + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f + biogaz_revendu_rate)
                            'comme la valeur est fixée on l'autoforce
                            biogaz_incinerateur_rate_force = True
                            biogaz_incinerateur_rate_f = 1 - (biogaz_torchere_rate_f + biogaz_chaudiere_rate_f + biogaz_CHP_rate_f + biogaz_revendu_rate)
                        Else
                            biogaz_incinerateur_rate = besoins_thermiques_biogaz_incineration / energie_primaire_biogaz
                        End If
                        '                    energie_primaire_biogaz_after_sechage = energie_primaire_biogaz * (1 - biogaz_sechage_rate - biogaz_revendu_rate)
                    End If
                End If
            End If
            If biogaz_incinerateur_non_couvert < 0 Then
                biogaz_incinerateur_non_couvert = 0
                remarque(5) = Feuil41.Range("thermique_erreur_biogaz_sechage_trop").Cells(1, 1)
            ElseIf biogaz_incinerateur_non_couvert > 0 Then
                remarque(6) = Feuil41.Range("thermique_erreur_biogaz_sechage_insuffisant").Cells(1, 1)
            End If
            'on calcule l'énergie primaire de combustible à apporter pour l'incinération
            If besoins_thermiques_combustible_incineration + biogaz_incinerateur_non_couvert > 0 Then
                pd_BesoinsCombustibleExterneIncineration_kWhj = (besoins_thermiques_combustible_incineration + biogaz_incinerateur_non_couvert)
                combustible_incinerateur = (besoins_thermiques_combustible_incineration + biogaz_incinerateur_non_couvert) / PCI(choix_combustible)
            Else
                combustible_incinerateur = 0
            End If
            If (biogaz_torchere_rate_mini + biogaz_chaudiere_rate_mini) > (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) Then
                biogaz_torchere_rate_mini = biogaz_torchere_rate_mini / (biogaz_torchere_rate_mini + biogaz_chaudiere_rate_mini) * (1 - biogaz_incinerateur_rate - biogaz_revendu_rate)
                biogaz_chaudiere_rate_mini = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini
            End If
            'On passe à la suite
            biogaz_torchere_rate = 0
            biogaz_CHP_rate = 0
            biogaz_chaudiere_rate = 0
            If (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) > 0 Then
                '**************************************** AJOUT INITIALISATION TABLEAU UTILITES *************************
                If choix_objectif = objectif_autosuffisance_thermique Then    'AUTOSUFFISANCE THERMIQUE MAXIMALE SANS COGENERATION
                    If choix_cogeneration = cogeneration_non Then   'seulement chaudière
                        If biogaz_chaudiere_rate_force = True Then   'chaudière forcée
                            biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                            If biogaz_torchere_rate_force = True Then   'chaudière et torchère forcées
                                biogaz_torchere_rate = biogaz_torchere_rate_f
                            Else                                         'chaudière forcée mais torchère non forcée
                                biogaz_torchere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate
                            End If
                        ElseIf biogaz_torchere_rate_force = True Then   'chaudière non forcée mais torchère forcée
                            biogaz_torchere_rate = biogaz_torchere_rate_f
                            biogaz_chaudiere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_torchere_rate
                        Else
                            'on calcule le potentiel d'énergie sur la chaudière
                            potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rdt * ((1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini)
                            '******************** PRENDRE QUE LES BESOINS CHAUDS "besoins_HOT" ************************************
                            If (besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR) <= potentiel_energie Then   'énergie suffisante pour couvrir l'ensemble des besoins restants
                                'on calcule la répartition du biogaz
                                biogaz_chaudiere_rate = ((besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR) / biogaz_chaudiere_rdt) / energie_primaire_biogaz
                                biogaz_torchere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate
                            Else
                                biogaz_torchere_rate = biogaz_torchere_rate_mini
                                biogaz_chaudiere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_torchere_rate
                            End If
                        End If
                        '******************* ON RENSEIGNE UN TABLEAU UTILITES *************************************************
                        Call utilites_biogaz(energie_primaire_biogaz)
                        '******************* ON LANCE LE RESEAU D'ECHANGEURS DE CHALEUR AVEC LES UTILITES **************************
                        Call pincement.VeriPinch(3)    '3 donc fait uniquement l'ajout d'utilités
                        besoins_HOT = pincement.besoins_HOT
                        If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                            besoins_HOT = 0
                        End If
                        besoins_COLD = pincement.besoins_COLD
                        If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                            besoins_COLD = 0
                        End If
                    ElseIf choix_cogeneration = cogeneration_oui Then    'AUTOSUFFISANCE THERMIQUE MAXIMALE AVEC COGENERATION
                        If biogaz_CHP_rate_force = True Then     'CHP forcée
                            biogaz_CHP_rate = biogaz_CHP_rate_f
                            'type de cogénération installable
                            energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                            Call definition_type_CHP(energie_primaire_CHP)
                            'On continue pour les autres utilités
                            If biogaz_chaudiere_rate_force = True Then    'CHP et chaudière forcées
                                biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                                If biogaz_torchere_rate_force = True Then
                                    biogaz_torchere_rate = biogaz_torchere_rate_f
                                Else
                                    biogaz_torchere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
                                End If
                            Else    'CHP forcée mais chaudière non forcée
                                If biogaz_torchere_rate_force = True Then
                                    biogaz_torchere_rate = biogaz_torchere_rate_f
                                    biogaz_chaudiere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
                                Else
                                    'on doit répartir entre chaudière et torchère
                                    'on calcule le potentiel d'énergie sur la chaudière
                                    potentiel_energie = energie_primaire_biogaz * biogaz_chaudiere_rdt * ((1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini - biogaz_CHP_rate)
                                    '******************* ON RENSEIGNE UN TABLEAU UTILITES *************************************************
                                    Call utilites_biogaz(energie_primaire_biogaz)
                                    '******************* ON LANCE LE RESEAU D'ECHANGEURS DE CHALEUR AVEC LA CHP **************************
                                    Call pincement.VeriPinch(3)    '3 donc fait uniquement l'ajout d'utilités
                                    besoins_HOT = pincement.besoins_HOT
                                    If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                                        besoins_HOT = 0
                                    End If
                                    besoins_COLD = pincement.besoins_COLD
                                    If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                                        besoins_COLD = 0
                                    End If
                                    '******************** PRENDRE QUE LES BESOINS CHAUDS "besoins_HOT" ************************************
                                    If (besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR) <= potentiel_energie Then   'énergie suffisante pour couvrir l'ensemble des besoins restants
                                        'on calcule la répartition du biogaz
                                        biogaz_chaudiere_rate = ((besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR) / biogaz_chaudiere_rdt) / energie_primaire_biogaz
                                        biogaz_torchere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
                                    Else
                                        biogaz_torchere_rate = biogaz_torchere_rate_mini
                                        biogaz_chaudiere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
                                    End If
                                End If
                            End If
                            '******************* ON RENSEIGNE UN TABLEAU UTILITES *************************************************
                            Call utilites_biogaz(energie_primaire_biogaz)
                            '******************* ON LANCE LE RESEAU D'ECHANGEURS DE CHALEUR AVEC LES UTILITES **************************
                            Call pincement.VeriPinch(3)    '3 donc fait uniquement l'ajout d'utilités
                            besoins_HOT = pincement.besoins_HOT
                            If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                                besoins_HOT = 0
                            End If
                            besoins_COLD = pincement.besoins_COLD
                            If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                                besoins_COLD = 0
                            End If
                        ElseIf biogaz_chaudiere_rate_force = True Then    'CHP non forcée mais chaudière forcée
                            biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                            If biogaz_torchere_rate_force = True Then    'CHP non forcée mais chaudière et torchère forcées
                                biogaz_torchere_rate = biogaz_torchere_rate_f
                                biogaz_CHP_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
                            Else       'CHP et torchère non forcées mais chaudière forcée
                                'on doit répartir entre la CHP et la torchère
                                If biogaz_chaudiere_rate > ((1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini) Then
                                    biogaz_torchere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate
                                    biogaz_CHP_rate = 0
                                ElseIf biogaz_chaudiere_rate < ((1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini - biogaz_CHP_rate_maxi) Then    'pour ne pas dépasser le maximum utilisable en CHP
                                    biogaz_CHP_rate = biogaz_CHP_rate_maxi
                                    biogaz_torchere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_CHP_rate)
                                Else
                                    biogaz_torchere_rate = biogaz_torchere_rate_mini
                                    biogaz_CHP_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
                                End If
                            End If
                            'type de cogénération installable
                            energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                            Call definition_type_CHP(energie_primaire_CHP)
                            '******************* ON RENSEIGNE UN TABLEAU UTILITES *************************************************
                            Call utilites_biogaz(energie_primaire_biogaz)
                            '******************* ON LANCE LE RESEAU D'ECHANGEURS DE CHALEUR AVEC LES UTILITES **************************
                            Call pincement.VeriPinch(3)    '3 donc fait uniquement l'ajout d'utilités
                            besoins_HOT = pincement.besoins_HOT
                            If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                                besoins_HOT = 0
                            End If
                            besoins_COLD = pincement.besoins_COLD
                            If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                                besoins_COLD = 0
                            End If
                        Else   'CHP et chaudière non forcés
                            If biogaz_torchere_rate_force = True Then
                                biogaz_torchere_rate = biogaz_torchere_rate_f
                            Else
                                biogaz_torchere_rate = biogaz_torchere_rate_mini
                            End If
                            CHP_maxi = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate_mini)
                            If CHP_maxi < 0 Then
                                CHP_maxi = 0
                            ElseIf CHP_maxi > biogaz_CHP_rate_maxi Then    'si CHP_maxi>0.9
                                CHP_maxi = biogaz_CHP_rate_maxi
                            End If
                            'On fait des itérations dans la subroutine calcul_biogaz_repartition (elle est à adapter)
                            Call calcul_biogaz_repartition(CHP_maxi, energie_primaire_biogaz)
                            '******************* ON RENSEIGNE UN TABLEAU UTILITES *************************************************
                            Call utilites_biogaz(energie_primaire_biogaz)
                            '******************* ON LANCE LE RESEAU D'ECHANGEURS DE CHALEUR AVEC LES UTILITES **************************
                            Call pincement.VeriPinch(3)    '3 donc fait uniquement l'ajout d'utilités
                            besoins_HOT = pincement.besoins_HOT
                            If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                                besoins_HOT = 0
                            End If
                            besoins_COLD = pincement.besoins_COLD
                            If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                                besoins_COLD = 0
                            End If
                        End If
                    End If
                ElseIf choix_objectif = objectif_autosuffisance_electrique Then    'AUTOSUFFISANCE ELECTRIQUE MAXIMALE
                    If biogaz_CHP_rate_force = True Then
                        biogaz_CHP_rate = biogaz_CHP_rate_f
                        'type de cogénération installable
                        energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                        Call definition_type_CHP(energie_primaire_CHP)
                        'on regarde les autres valeurs pour voir si elles sont forcées
                        If biogaz_chaudiere_rate_force = True Then     'CHP et chaudière forcés
                            biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                            If biogaz_torchere_rate_force = True Then
                                biogaz_torchere_rate = biogaz_torchere_rate_f
                            Else
                                biogaz_torchere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
                            End If
                        Else
                            If biogaz_torchere_rate_force = True Then
                                biogaz_torchere_rate = biogaz_torchere_rate_f
                                biogaz_chaudiere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
                            Else
                                If biogaz_CHP_rate > biogaz_CHP_rate_maxi Then
                                    biogaz_torchere_rate = ((1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_CHP_rate) / 2  'on partage en 2 la chaudière et la torchère
                                    biogaz_chaudiere_rate = ((1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_CHP_rate) / 2  'on partage en 2 la chaudière et la torchère
                                Else   'on répartit au mieux entre torchère et chaudière
                                    '******************* ON RENSEIGNE UN TABLEAU UTILITES *************************************************
                                    Call utilites_biogaz(energie_primaire_biogaz)
                                    '******************* ON LANCE LE RESEAU D'ECHANGEURS DE CHALEUR AVEC LES UTILITES **************************
                                    Call pincement.VeriPinch(3)    '3 donc fait uniquement l'ajout d'utilités
                                    besoins_HOT = pincement.besoins_HOT
                                    If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                                        besoins_HOT = 0
                                    End If
                                    besoins_COLD = pincement.besoins_COLD
                                    If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                                        besoins_COLD = 0
                                    End If
                                    'énergie maximale que l'on peut atteindre avec la chaudière
                                    potentiel_energie = ((1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_CHP_rate - biogaz_torchere_rate_mini) * energie_primaire_biogaz * biogaz_chaudiere_rdt
                                    If potentiel_energie < (besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR) Then   'on fixe la torchère au min et la chaudière prend ce qu'il reste
                                        biogaz_torchere_rate = biogaz_torchere_rate_mini
                                        biogaz_chaudiere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
                                    Else   'on calcule la répartition permettant de couvrir exactement les besoins
                                        biogaz_chaudiere_rate = ((besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR) / biogaz_chaudiere_rdt) / energie_primaire_biogaz
                                        biogaz_torchere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_chaudiere_rate)
                                    End If
                                End If
                            End If
                            '******************* ON RENSEIGNE UN TABLEAU UTILITES *************************************************
                            Call utilites_biogaz(energie_primaire_biogaz)
                            '******************* ON LANCE LE RESEAU D'ECHANGEURS DE CHALEUR AVEC LES UTILITES **************************
                            Call pincement.VeriPinch(3)    '3 donc fait uniquement l'ajout d'utilités
                            besoins_HOT = pincement.besoins_HOT
                            If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                                besoins_HOT = 0
                            End If
                            besoins_COLD = pincement.besoins_COLD
                            If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                                besoins_COLD = 0
                            End If
                        End If
                    ElseIf biogaz_chaudiere_rate_force = True Then
                        biogaz_chaudiere_rate = biogaz_chaudiere_rate_f
                        'on regarde si la torchère est fixée
                        If biogaz_torchere_rate_force = True Then
                            biogaz_torchere_rate = biogaz_torchere_rate_f
                            biogaz_CHP_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_torchere_rate)
                        Else
                            If biogaz_chaudiere_rate > ((1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini) Then    'il ne reste pas assez pour le minimum en torchère
                                biogaz_torchere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate
                                biogaz_CHP_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_torchere_rate)
                            ElseIf biogaz_chaudiere_rate < ((1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_torchere_rate_mini - biogaz_CHP_rate_maxi) Then    'pour ne pas dépasser le maximum utilisable en CHP
                                biogaz_CHP_rate = biogaz_CHP_rate_maxi
                                biogaz_torchere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_CHP_rate)
                            Else
                                biogaz_torchere_rate = biogaz_torchere_rate_mini
                                biogaz_CHP_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_chaudiere_rate + biogaz_torchere_rate)
                            End If
                        End If
                        'type de cogénération installable
                        energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                        Call definition_type_CHP(energie_primaire_CHP)
                        '******************* ON RENSEIGNE UN TABLEAU UTILITES *************************************************
                        Call utilites_biogaz(energie_primaire_biogaz)
                        '******************* ON LANCE LE RESEAU D'ECHANGEURS DE CHALEUR AVEC LES UTILITES **************************
                        Call pincement.VeriPinch(3)    '3 donc fait uniquement l'ajout d'utilités
                        besoins_HOT = pincement.besoins_HOT
                        If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                            besoins_HOT = 0
                        End If
                        besoins_COLD = pincement.besoins_COLD
                        If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                            besoins_COLD = 0
                        End If
                    ElseIf biogaz_torchere_rate_force = True Then
                        biogaz_torchere_rate = biogaz_torchere_rate_f
                        If biogaz_torchere_rate > ((1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate_mini) Then
                            biogaz_chaudiere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_torchere_rate
                            biogaz_CHP_rate = 0
                        ElseIf biogaz_torchere_rate < ((1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_chaudiere_rate_mini - biogaz_CHP_rate_maxi) Then
                            biogaz_CHP_rate = biogaz_CHP_rate_maxi
                            biogaz_chaudiere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_CHP_rate)
                        Else
                            biogaz_chaudiere_rate = biogaz_chaudiere_rate_mini
                            biogaz_CHP_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
                        End If
                        'type de cogénération installable
                        energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                        Call definition_type_CHP(energie_primaire_CHP)
                        '******************* ON RENSEIGNE UN TABLEAU UTILITES *************************************************
                        Call utilites_biogaz(energie_primaire_biogaz)
                        '******************* ON LANCE LE RESEAU D'ECHANGEURS DE CHALEUR AVEC LES UTILITES **************************
                        Call pincement.VeriPinch(3)    '3 donc fait uniquement l'ajout d'utilités
                        besoins_HOT = pincement.besoins_HOT
                        If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                            besoins_HOT = 0
                        End If
                        besoins_COLD = pincement.besoins_COLD
                        If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                            besoins_COLD = 0
                        End If
                    Else
                        biogaz_chaudiere_rate = biogaz_chaudiere_rate_mini
                        biogaz_torchere_rate = biogaz_torchere_rate_mini
                        biogaz_CHP_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_torchere_rate + biogaz_chaudiere_rate)
                        'type de cogénération installable
                        energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
                        Call definition_type_CHP(energie_primaire_CHP)
                        '******************* ON RENSEIGNE UN TABLEAU UTILITES *************************************************
                        Call utilites_biogaz(energie_primaire_biogaz)
                        '******************* ON LANCE LE RESEAU D'ECHANGEURS DE CHALEUR AVEC LES UTILITES **************************
                        Call pincement.VeriPinch(3)    '3 donc fait uniquement l'ajout d'utilités
                        besoins_HOT = pincement.besoins_HOT
                        If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                            besoins_HOT = 0
                        End If
                        besoins_COLD = pincement.besoins_COLD
                        If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                            besoins_COLD = 0
                        End If
                    End If
                End If
            End If
            'calcul de l'énergie électrique produite
            energie_elec = biogaz_CHP_rate * energie_primaire_biogaz * biogaz_CHP_rdt_elec
            'VERIFICATION DE LA REPARTITION
            If Abs(biogaz_CHP_rate) + Abs(biogaz_chaudiere_rate) + Abs(biogaz_torchere_rate) + Abs(biogaz_incinerateur_rate) + Abs(biogaz_revendu_rate) <> 1 Then
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
        Else
            biogaz_torchere_rate = 0
            biogaz_chaudiere_rate = 0
            biogaz_CHP_rate = 0
        End If
        'GoTo bypass_combustible
        'ICI ON AJOUTE LE COMBUSTIBLE EXTERNE
        If besoins_HOT > 0 Then
            combustible_flux = (besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / combustible_chaudiere_rdt) / PCI(choix_combustible)
            pd_BesoinsCombustibleExterneHorsIncineration_kWhj = (besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / combustible_chaudiere_rdt)
            nb_flux_utilites = nb_flux_utilites + 1
            'on redimensionne le tableau
            ReDim Preserve utilites_flux_energie(1 To flux_thermique_utilites_nb_caracteristiques, 1 To nb_flux_utilites)
            'on remplit les données
            utilites_flux_energie(repere_flux_thermique_utilites_etat, nb_flux_utilites) = True
            utilites_flux_energie(repere_flux_thermique_utilites_nom, nb_flux_utilites) = Feuil41.Range("thermique_combustible_chaudiere_nom_flux").Cells(1, 1)
            utilites_flux_energie(repere_flux_thermique_utilites_type, nb_flux_utilites) = "Ch"
            utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, nb_flux_utilites) = combustible_chaudiere_Tch
            utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, nb_flux_utilites) = combustible_flux * combustible_chaudiere_rdt * PCI(choix_combustible) / pd_NOMBRE_HEURE_PAR_JOUR
            utilites_flux_energie(repere_flux_thermique_utilites_MCp_kW_°C, nb_flux_utilites) = utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, nb_flux_utilites) / (utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, nb_flux_utilites) - combustible_chaudiere_Tfr)
            utilites_flux_energie(repere_flux_thermique_utilites_original, nb_flux_utilites) = True
            For j = 1 To flux_thermique_utilites_nb_caracteristiques
                Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab + nb_flux_utilites, Colonne + j) = utilites_flux_energie(j, nb_flux_utilites)
            Next j
            Call pincement.ecriture_formule_pour_changement_langue_noms_utilites(pi_FiliereConsideree)
        Else
            combustible_flux = 0
        End If
bypass_combustible:
        '******************* ON LANCE LE RESEAU D'ECHANGEURS DE CHALEUR AVEC LES UTILITES **************************
        Call pincement.VeriPinch(3)    '3 donc fait uniquement l'ajout d'utilités
        besoins_HOT = pincement.besoins_HOT
        If energie_primaire_biogaz > 0 Then
            If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                besoins_HOT = 0
            End If
            besoins_COLD = pincement.besoins_COLD
            If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                besoins_COLD = 0
            End If
        End If
        'on représente le réseau d'échangeurs une fois que l'on arrive à la dernière itération
        If besoins_totaux_HOT > 0 Then
            autosuffisance_thermique = 1 - ((combustible_flux * combustible_chaudiere_rdt + combustible_incinerateur) * PCI(choix_combustible)) / (besoins_totaux_HOT * pd_NOMBRE_HEURE_PAR_JOUR)
        Else
            autosuffisance_thermique = 1
        End If

        '******* EXPORT DU HEN AU FORMAT GRAPHIQUE *************** MODIF CRE 18/09/2012
        Call pincement.VeriPinch(0)

        '**************************** SUPRESSION DU MODULE DE CLASSE PINCH *********************************
        Set pincement = Nothing
    Else
        autosuffisance_thermique = 1
        '******************* VOIR SI L'ON PRODUIT QUAND MEME DU BIOGAZ (UASB) ******************************
        If energie_primaire_biogaz > 0 Then
            biogaz_chaudiere_rate = 0
            biogaz_torchere_rate = biogaz_torchere_rate_mini
            biogaz_CHP_rate = 1 - (biogaz_chaudiere_rate + biogaz_torchere_rate)
            energie_primaire_CHP = biogaz_CHP_rate * energie_primaire_biogaz
            Call definition_type_CHP(energie_primaire_CHP)
            'On n'a pas besoin de récupérer de l'énergie donc on fixe les rendements à 0
            biogaz_CHP_rdt_eau_BT = 0
            biogaz_CHP_rdt_eau_HT = 0
            biogaz_CHP_rdt_vapeur = 0
            energie_elec = biogaz_CHP_rate * energie_primaire_biogaz * biogaz_CHP_rdt_elec
        End If
    End If

    'A cause du changement de feuille dans le PINCH
    Application.Calculation = xlCalculationManual

    'GESTION DU NOMBRE D'EQUIPEMENTS DE COGENERATION
    'calcul de la puissance électrique de la cogénération
    biogaz_CHP_Pw_reel = energie_elec / pd_NOMBRE_HEURE_PAR_JOUR   'production électrique moyenne journalière rapportée en puissance instantanée moyenne sur la journée
    electricite_biogaz = energie_elec

    'modif MSM 25/10/12
    '    If choix_electricite_alternatif = electricite_alternatif_oui Then
    '        energie_elec = energie_elec + electricite_alternatif
    '    End If
    'fin modif MSM 25/10/12

    'on calcule le PCI correspondant au combustible utilisé (pour empreinte CO2)
    If choix_combustible = combustible_repere_gaz_naturel Then
        gaz_naturel_kWhPCIj = (combustible_flux + combustible_incinerateur) * PCI(choix_combustible)
    ElseIf choix_combustible = combustible_repere_fioul Then
        fioul_kWhPCIj = (combustible_flux + combustible_incinerateur) * PCI(choix_combustible)
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



Sub calcul_biogaz_repartition(ByRef CHP_maxi, ByVal energie_primaire_biogaz)

    If version_Developpement = False Then On Error GoTo erreur_non_geree 'a
    Dim NumeroErreurNonGeree As String


    'On souhaite trouver l'optimum de biogaz à envoyer en CHP  (autosuffisance thermique maxi avec cogénération)

    'pour la précision des calculs
    Const precision = 0.00000001
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

    Dim besoins_HOT As Double
    Dim besoins_COLD As Double

    '    Dim besoins_HOT_CRE() As Double
    '    ReDim besoins_HOT_CRE(0 To 2)

    'test limite pour savoir si on peut mettre un moteur ou non
    'utile car si l'énergie primaire limite pour le moteur est dans l'intervalle --> courbe non continue et non strictement monotone sur l'intervalle --> dichotomie impossible
    biogaz_CHP_rate = pd_NOMBRE_HEURE_PAR_JOUR * limit_min_moteur / energie_primaire_biogaz
    If biogaz_CHP_rate < CHP_maxi Then
        'on calcule la répartition
        x_chaudiere = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_torchere_rate - biogaz_CHP_rate
        biogaz_chaudiere_rate = x_chaudiere
        'on définit le type de cogé
        energie_primaire_CHP = energie_primaire_biogaz * biogaz_CHP_rate
        Call definition_type_CHP(energie_primaire_CHP)
        rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
        rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
        rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
        '******************* ON RENSEIGNE UN TABLEAU UTILITES *************************************************
        Call utilites_biogaz(energie_primaire_biogaz)
        '******************* ON LANCE LE RESEAU D'ECHANGEURS DE CHALEUR AVEC LES UTILITES **************************
        Call pincement.VeriPinch(3)    '3 donc fait uniquement l'ajout d'utilités
        besoins_HOT = pincement.besoins_HOT
        If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
            besoins_HOT = 0
        End If
        besoins_COLD = pincement.besoins_COLD
        If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
            besoins_COLD = 0
        End If
        'on va regarder si on couvre les besoins
        If besoins_HOT > 0 Then
            moteur_possible = False
        Else
            moteur_possible = True
        End If
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
            x_chaudiere = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - biogaz_torchere_rate - x_CHP(I)
            biogaz_CHP_rate = x_CHP(I)
            biogaz_chaudiere_rate = x_chaudiere
            bilan(I) = bilan_positif
            'on définit le type de cogé
            energie_primaire_CHP = energie_primaire_biogaz * biogaz_CHP_rate
            Call definition_type_CHP(energie_primaire_CHP)
            rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
            rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
            rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur
            '******************* ON RENSEIGNE UN TABLEAU UTILITES *************************************************
            Call utilites_biogaz(energie_primaire_biogaz)
            '******************* ON LANCE LE RESEAU D'ECHANGEURS DE CHALEUR AVEC LES UTILITES **************************
            Call pincement.VeriPinch(3)    '3 donc fait uniquement l'ajout d'utilités
            besoins_HOT = pincement.besoins_HOT
            'MODIF CRE 20/09/2012   ajout de besoins_HOT_CRE(i) pour relever la valeur à chaque point. Utile à la fin des itérations, pour garder le taux de répartition qui donne des besoins HOT nuls
            '            besoins_HOT_CRE(i) = pincement.besoins_HOT
            If Abs(besoins_HOT * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                besoins_HOT = 0
            End If
            '            If Abs(besoins_HOT_CRE(i) * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
            '                besoins_HOT_CRE(i) = 0
            '            End If
            besoins_COLD = pincement.besoins_COLD
            If Abs(besoins_COLD * pd_NOMBRE_HEURE_PAR_JOUR / energie_primaire_biogaz) <= critere_convergence Then
                besoins_COLD = 0
            End If
            If besoins_HOT > 0 Then
                bilan(I) = bilan_negatif
            End If
            '            If besoins_HOT_CRE(i) > 0 Then
            '                bilan(i) = bilan_negatif
            '            End If
        Next I
        'on va redéfinir les bornes de l'intervalle
        If bilan(0) < 0 Then
            x_CHP(0) = CHP_mini
            x_CHP(2) = CHP_mini
            GoTo suite_calcul
        ElseIf bilan(0) * bilan(1) < 0 Then
            x_CHP(0) = x_CHP(0)
            x_CHP(2) = x_CHP(1)
        ElseIf bilan(1) * bilan(2) < 0 Then
            x_CHP(0) = x_CHP(1)
            x_CHP(2) = x_CHP(2)
        ElseIf bilan(2) > 0 Then
            x_CHP(0) = CHP_maxi
            x_CHP(2) = CHP_maxi
            GoTo suite_calcul
        End If
    Next iter

suite_calcul:

    'on calcule tout bien avec la valeur conservée
    'MOD CRE 20/09/2012 on garde la valeur du taux de répartition qui donne des besoins nuls. A priori, les 3 valeurs sont très proches, donc on prends la première qui donne besoins_HOT = 0
    '    For i = 0 To 2
    '        If besoins_HOT_CRE(i) = 0 Then
    '        biogaz_CHP_rate = x_CHP(i)
    '        End If
    '    Next
    biogaz_CHP_rate = x_CHP(0)   '(x_CHP(0) + x_CHP(2)) / 2
    biogaz_chaudiere_rate = (1 - biogaz_incinerateur_rate - biogaz_revendu_rate) - (biogaz_CHP_rate + biogaz_torchere_rate)
    'on définit le type de cogé
    energie_primaire_CHP = energie_primaire_biogaz * biogaz_CHP_rate
    Call definition_type_CHP(energie_primaire_CHP)
    rdt_th_coge(energie_eau_BT) = biogaz_CHP_rdt_eau_BT
    rdt_th_coge(energie_eau_HT) = biogaz_CHP_rdt_eau_HT
    rdt_th_coge(energie_vapeur_global) = biogaz_CHP_rdt_vapeur




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.12a"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub fonctionnement_reel(ByVal biogaz_caract, ByRef energie_elec, ByRef type_valorisation_biogaz, ByRef valorisation_chaudiere, ByRef ratio_biogaz_valorise)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Call dimensionnement(biogaz_caract, energie_elec, type_valorisation_biogaz, valorisation_chaudiere, ratio_biogaz_valorise)




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

    '    For i = 1 To nb_niveaux_energie
    '        Feuil41.Range("thermique_besoins_thermiques_kWhj").Cells(i, colonne) = besoins_thermiques(i)
    '    Next i
    '    For i = 1 To nb_vapeur_max
    '        Feuil41.Range("thermique_pression_vapeur").Cells(i, colonne) = pression_vapeur(i)
    '    Next i
    '    For i = 1 To nb_vapeur_max
    '        Feuil41.Range("thermique_besoins_vapeur_kgj").Cells(i, colonne) = besoins_vapeur_kgj(i)
    '    Next i
    '    For i = 1 To energie_vapeur_global
    '        Feuil41.Range("thermique_disponibilites_thermiques_kWhj").Cells(i, colonne) = disponibilites_thermiques(i)
    '    Next i
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
    Feuil41.Range("thermique_biogaz_sechage_rate").Cells(1, Colonne) = biogaz_incinerateur_rate
    '    For i = 1 To energie_vapeur_global
    '        Feuil41.Range("thermique_energie_fatale_kWhj").Cells(i, colonne) = energie_fatale_kWhj(i)
    '    Next i
    Feuil41.Range("thermique_vapeur_fatale_kgj").Cells(1, Colonne) = vapeur_fatale_kgj
    Feuil41.Range("thermique_autosuffisance").Cells(1, Colonne) = autosuffisance_thermique


    If biogaz_revendu_rate > 0 Then
        Feuil41.Range("thermique_biogaz_revendu_rate").Cells(1, Colonne) = biogaz_revendu_rate
    Else
        Feuil41.Range("thermique_biogaz_revendu_rate").Cells(1, Colonne) = "'-"       'biogaz_revendu_rate
    End If

    If biogaz_incinerateur_rate = 0 Then
        Feuil41.Range("thermique_biogaz_sechage_rate").Cells(1, Colonne) = "'-"    'biogaz_incinerateur_rate
    End If

    If biogaz_caracteristiques(Q_Nm3j) = 0 Then
        Feuil41.Range("thermique_biogaz_chaudiere_rdt").Cells(1, Colonne) = "'-"    'biogaz_chaudiere_rdt
        Feuil41.Range("thermique_biogaz_CHP_type").Cells(1, Colonne) = "'-"    'Feuil41.Range("thermique_no_CHP").Cells(1, 1)
        Feuil41.Range("thermique_biogaz_CHP_rate").Cells(1, Colonne) = "'-"    'biogaz_CHP_rate
        Feuil41.Range("thermique_biogaz_chaudiere_rate").Cells(1, Colonne) = "'-"    'biogaz_chaudiere_rate
        Feuil41.Range("thermique_biogaz_torchere_rate").Cells(1, Colonne) = "'-"    'biogaz_torchere_rate
        Feuil41.Range("thermique_biogaz_sechage_rate").Cells(1, Colonne) = "'-"    'biogaz_incinerateur_rate
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
    If combustible_incinerateur > 0 Then
        Feuil41.Range("thermique_combustible_sechage").Cells(1, Colonne) = combustible_incinerateur
    Else
        Feuil41.Range("thermique_combustible_sechage").Cells(1, Colonne) = "'-"    'combustible_incinerateur
    End If

    Call ecriture_resultats_OPEX

    'modif MSM 25/10/12
    '    If choix_electricite_alternatif = electricite_alternatif_oui Then
    '        Feuil41.Range("thermique_electricite_alternatif").Cells(1, colonne) = electricite_alternatif
    '    Else
    '        Feuil41.Range("thermique_electricite_alternatif").Cells(1, colonne) = "'-"     'electricite_alternatif
    '    End If
    'fin modif MSM 25/10/12

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
            Feuil55.Range("OPEX_combustibles_variable_kWhj").Cells(I, Colonne) = (combustible_flux + combustible_incinerateur) * PCI(choix_combustible)
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
    Feuil41.Range("thermique_combustible_sechage").Cells(1, Colonne) = ""    'combustible_incinerateur
    For I = 1 To energie_vapeur_global
        Feuil41.Range("thermique_energie_fatale_kWhj").Cells(I, Colonne) = ""    'energie_fatale_kWhj(i)
    Next I
    Feuil41.Range("thermique_vapeur_fatale_kgj").Cells(1, Colonne) = ""    'vapeur_fatale_kgj
    Feuil41.Range("thermique_autosuffisance").Cells(1, Colonne) = ""    'autosuffisance_thermique

    '    Feuil41.Range("thermique_electricite_alternatif").Cells(1, colonne) = ""     'electricite_alternatif 'modif MSM 25/10/12
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
    PCI(combustible_repere_gaz_naturel) = pd_PCI_GAZ_NATUREL_kWh_Nm3   'kWh/Nm3
    PCI(combustible_repere_fioul) = pd_PCI_FIOUL_kWh_L    'kWh/L   source: Delphine NAWAWI

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

    Dim I As Integer
    Dim j As Integer
    Dim Colonne As Integer
    Dim flux_ajoutes As Boolean

    choix_objectif = Feuil41.Range("thermique_choix_objectif").Cells(1, pi_FiliereConsideree)
    If choix_objectif = objectif_autosuffisance_thermique Then
        choix_cogeneration = Feuil41.Range("thermique_choix_cogeneration").Cells(1, pi_FiliereConsideree)
    Else
        choix_cogeneration = cogeneration_oui
    End If
    choix_combustible = Feuil41.Range("thermique_choix_combustible").Cells(1, pi_FiliereConsideree)
    choix_combustible_utilisation = Feuil41.Range("thermique_choix_combustible_utilisation").Cells(1, pi_FiliereConsideree)

    'modif MSM 25/10/12
    '    choix_electricite_alternatif = Feuil41.Range("thermique_choix_elec_alternatif").Cells(1, pi_FiliereConsideree)
    '    If choix_electricite_alternatif = electricite_alternatif_oui Then
    '        electricite_alternatif = Feuil41.Range("thermique_electricite_alternatif").Cells(1, pi_FiliereConsideree)
    '    Else
    '        electricite_alternatif = 0
    '    End If
    'modif MSM 25/10/12

    'On récupère les utilités rentrées par l'utilisateur
    ReDim utilites_ajoutees(1 To flux_thermique_utilites_nb_caracteristiques, 1 To 1)
    nb_utilites_ajoutees = 0
    I = INDICE_PREMIERE_LIGNE
    Colonne = (pi_FiliereConsideree - 1) * flux_thermique_utilites_nb_caracteristiques + repere_flux_thermique_utilites_original
    Do While Feuil63.Range("DonneesPincementFluxUtilites").Cells(I, Colonne) <> ""
        If Feuil63.Range("DonneesPincementFluxUtilites").Cells(I, Colonne) = False Then
            nb_utilites_ajoutees = nb_utilites_ajoutees + 1
            ReDim Preserve utilites_ajoutees(1 To flux_thermique_utilites_nb_caracteristiques, 1 To nb_utilites_ajoutees)
            For j = 1 To flux_thermique_utilites_nb_caracteristiques
                utilites_ajoutees(j, nb_utilites_ajoutees) = Feuil63.Range("DonneesPincementFluxUtilites").Cells(I, Colonne - repere_flux_thermique_utilites_original + j)
            Next j
        End If
        I = I + 1
    Loop

    'On regarde les flux process ajoutés par l'utilisateur
    I = INDICE_PREMIERE_LIGNE
    Colonne = (pi_FiliereConsideree - 1) * flux_thermique_process_nb_caracteristiques + repere_flux_thermique_process_original
    flux_ajoutes = False
    Do While Feuil63.Range("DonneesPincementFluxProcedes").Cells(I, Colonne) <> ""
        If Feuil63.Range("DonneesPincementFluxProcedes").Cells(I, Colonne) = False Then
            For j = 1 To UBound(flux_thermique_process, 2)
                If Feuil63.Range("DonneesPincementFluxProcedes").Cells(I, Colonne - repere_flux_thermique_process_original + repere_flux_thermique_process_nom) = flux_thermique_process(repere_flux_thermique_process_nom, j) Then
                    GoTo appel_par_Martin
                End If
            Next j
            If flux_thermique_process(repere_flux_thermique_process_etat, 1) = "" Then      'si le flux est le premier a être entré dans le tableau -> on écrit dans la première ligne de celui-ci
                ReDim flux_thermique_process(1 To flux_thermique_process_nb_caracteristiques, 1 To 1)
            Else        'sinon -> ajout d'une ligne au tableau avec conservation des données
                ReDim Preserve flux_thermique_process(1 To flux_thermique_process_nb_caracteristiques, 1 To UBound(flux_thermique_process, 2) + 1)
            End If
            flux_thermique_process(repere_flux_thermique_process_etat, UBound(flux_thermique_process, 2)) = Feuil63.Range("DonneesPincementFluxProcedes").Cells(I, Colonne - repere_flux_thermique_process_original + repere_flux_thermique_process_etat)
            flux_thermique_process(repere_flux_thermique_process_nom, UBound(flux_thermique_process, 2)) = Feuil63.Range("DonneesPincementFluxProcedes").Cells(I, Colonne - repere_flux_thermique_process_original + repere_flux_thermique_process_nom)
            flux_thermique_process(repere_flux_thermique_process_T_in_°C, UBound(flux_thermique_process, 2)) = Feuil63.Range("DonneesPincementFluxProcedes").Cells(I, Colonne - repere_flux_thermique_process_original + repere_flux_thermique_process_T_in_°C)
            flux_thermique_process(repere_flux_thermique_process_T_out_°C, UBound(flux_thermique_process, 2)) = Feuil63.Range("DonneesPincementFluxProcedes").Cells(I, Colonne - repere_flux_thermique_process_original + repere_flux_thermique_process_T_out_°C)
            flux_thermique_process(repere_flux_thermique_process_MCp_kW_°C, UBound(flux_thermique_process, 2)) = Feuil63.Range("DonneesPincementFluxProcedes").Cells(I, Colonne - repere_flux_thermique_process_original + repere_flux_thermique_process_MCp_kW_°C)
            flux_thermique_process(repere_flux_thermique_process_original, UBound(flux_thermique_process, 2)) = Feuil63.Range("DonneesPincementFluxProcedes").Cells(I, Colonne - repere_flux_thermique_process_original + repere_flux_thermique_process_original)
            flux_ajoutes = True
        End If
        I = I + 1
    Loop
appel_par_Martin:
    If Feuil63.Range("DonneesPincementTriAutoProcedes").Cells(INDICE_PREMIERE_LIGNE, pi_FiliereConsideree) = False And pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
        Call reordonnement_flux_process
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub lecture_donnees_modification_pinch(ByRef biogaz_caract() As Double, ByRef energie_elec, ByRef type_valorisation_biogaz, ByRef valorisation_chaudiere, ByRef ratio_biogaz_valorise)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim Colonne As Integer
    Dim I As Integer

    Colonne = 3 * pi_CALCUL_VALEUR_EFFECTIVE + pi_FiliereConsideree

    For I = 1 To nb_biogaz_caracteristiques
        If Not IsNumeric(Feuil41.Range("thermique_biogaz_caracteristiques").Cells(I, Colonne)) Then
            biogaz_caract(I) = 0
        Else
            biogaz_caract(I) = Feuil41.Range("thermique_biogaz_caracteristiques").Cells(I, Colonne)
        End If
    Next I

    energie_elec = Feuil6.Range("electricite_total").Cells(1, Colonne)

    type_valorisation_biogaz = 0
    valorisation_chaudiere = 0
    ratio_biogaz_valorise = 0


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.15"

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
    biogaz_incinerateur_rate_force = Feuil41.Range("thermique_biogaz_sechage_rate").Cells(1, Colonne)
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
    If biogaz_incinerateur_rate_force = True Then
        biogaz_incinerateur_rate_f = Feuil41.Range("thermique_biogaz_sechage_rate").Cells(1, Colonne)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub




Private Sub reordonnement_flux_process()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'On reclasse les flux procédés dans l'ordre antérieur

    Dim I As Integer
    Dim j As Integer

    Dim flux_retries() As Variant
    Dim Ligne As Integer
    Dim Colonne As Integer
    Dim nb_flux_retries As Integer
    Dim nb_flux_anterieurs As Integer
    Dim nb_flux_actuels As Integer

    ReDim flux_retries(1 To flux_thermique_process_nb_caracteristiques, 1 To 1)

    Colonne = (pi_FiliereConsideree - 1) * flux_thermique_process_nb_caracteristiques + repere_flux_thermique_process_nom

    Ligne = INDICE_PREMIERE_LIGNE
    nb_flux_retries = 0
    nb_flux_actuels = UBound(flux_thermique_process, 2)

    Do While Feuil63.Range("DonneesPincementFluxProcedes").Cells(Ligne, Colonne) <> ""
        'On regarde si le flux est maintenant
        For I = 1 To nb_flux_actuels
            If Feuil63.Range("DonneesPincementFluxProcedes").Cells(Ligne, Colonne) = flux_thermique_process(repere_flux_thermique_process_nom, I) Then
                nb_flux_retries = nb_flux_retries + 1
                ReDim Preserve flux_retries(1 To flux_thermique_process_nb_caracteristiques, 1 To nb_flux_retries)
                For j = 1 To flux_thermique_process_nb_caracteristiques
                    flux_retries(j, nb_flux_retries) = flux_thermique_process(j, I)
                Next j
            End If
        Next I
        Ligne = Ligne + 1
    Loop

    nb_flux_anterieurs = Ligne - INDICE_PREMIERE_LIGNE

    PINCH_ORDONNER_FLUX_PROCESS = True
    Feuil63.Range("DonneesPincementTriAutoProcedes").Cells(INDICE_PREMIERE_LIGNE, pi_FiliereConsideree) = True

    If nb_flux_actuels <= nb_flux_anterieurs Then
        If nb_flux_retries = nb_flux_actuels Then
            For I = 1 To nb_flux_actuels
                For j = 1 To flux_thermique_process_nb_caracteristiques
                    flux_thermique_process(j, I) = flux_retries(j, I)
                Next j
            Next I
            PINCH_ORDONNER_FLUX_PROCESS = False
            Feuil63.Range("DonneesPincementTriAutoProcedes").Cells(INDICE_PREMIERE_LIGNE, pi_FiliereConsideree) = False
        End If
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "41.16"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub utilites_biogaz(ByVal energie_primaire_biogaz As Double)

    'On remplit le tableau comprenant les différents flux d'utilités sur le biogaz
    '***********************************************************************************************************
    'ATTENTION IL FAUDRA ALLER LIRE LES NOMS DES FLUX DANS EXCEL POUR LES AVOIR EN FONCTION DE LA LANGUE CHOISIE
    '***********************************************************************************************************

    Dim I As Integer
    Dim j As Integer
    Dim Colonne As Integer


    'on réinitialise les flux d'utilités
    I = 1
    Colonne = flux_thermique_utilites_nb_caracteristiques * (pi_FiliereConsideree - 1)
    Do While Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab + I, Colonne + repere_flux_thermique_utilites_etat) <> ""
        Feuil63.Range(Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab + I, Colonne + repere_flux_thermique_utilites_etat), Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab + I, Colonne + flux_thermique_utilites_nb_caracteristiques)).ClearContents
        I = I + 1
    Loop
    ReDim utilites_flux_energie(1 To flux_thermique_utilites_nb_caracteristiques, 1 To 1)

    'On remplit le tableau
    nb_flux_utilites = 0
    PINCH_ORDONNER_FLUX_UTILITES = False
    ' Ajouter les utilités renseignées à la main
    If nb_utilites_ajoutees > 0 Then
        nb_flux_utilites = nb_utilites_ajoutees
        ReDim Preserve utilites_flux_energie(1 To flux_thermique_utilites_nb_caracteristiques, 1 To nb_flux_utilites)
        For I = 1 To nb_utilites_ajoutees
            For j = 1 To flux_thermique_utilites_nb_caracteristiques
                utilites_flux_energie(j, I) = utilites_ajoutees(j, I)
            Next j
        Next I
    End If

    If biogaz_CHP_rate > 0 Then
        If biogaz_CHP_rdt_eau_BT > zero_force Then
            nb_flux_utilites = nb_flux_utilites + 1
            'on redimensionne le tableau
            ReDim Preserve utilites_flux_energie(1 To flux_thermique_utilites_nb_caracteristiques, 1 To nb_flux_utilites)
            'on remplit les données
            utilites_flux_energie(repere_flux_thermique_utilites_etat, nb_flux_utilites) = True
            utilites_flux_energie(repere_flux_thermique_utilites_nom, nb_flux_utilites) = Feuil41.Range("thermique_biogaz_CHP_BT_nom_flux").Cells(1, 1)
            utilites_flux_energie(repere_flux_thermique_utilites_type, nb_flux_utilites) = "Ch"
            utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, nb_flux_utilites) = biogaz_CHP_BT_Tch
            utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, nb_flux_utilites) = biogaz_CHP_rate * biogaz_CHP_rdt_eau_BT * energie_primaire_biogaz / pd_NOMBRE_HEURE_PAR_JOUR
            utilites_flux_energie(repere_flux_thermique_utilites_MCp_kW_°C, nb_flux_utilites) = utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, nb_flux_utilites) / (utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, nb_flux_utilites) - biogaz_CHP_BT_Tfr)
            utilites_flux_energie(repere_flux_thermique_utilites_original, nb_flux_utilites) = True
        End If
        If biogaz_CHP_rdt_eau_HT > zero_force Then
            nb_flux_utilites = nb_flux_utilites + 1
            'on redimensionne le tableau
            ReDim Preserve utilites_flux_energie(1 To flux_thermique_utilites_nb_caracteristiques, 1 To nb_flux_utilites)
            'on remplit les données
            utilites_flux_energie(repere_flux_thermique_utilites_etat, nb_flux_utilites) = True
            utilites_flux_energie(repere_flux_thermique_utilites_nom, nb_flux_utilites) = Feuil41.Range("thermique_biogaz_CHP_HT_nom_flux").Cells(1, 1)
            utilites_flux_energie(repere_flux_thermique_utilites_type, nb_flux_utilites) = "Ch"
            utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, nb_flux_utilites) = biogaz_CHP_HT_Tch
            utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, nb_flux_utilites) = biogaz_CHP_rate * biogaz_CHP_rdt_eau_HT * energie_primaire_biogaz / pd_NOMBRE_HEURE_PAR_JOUR
            utilites_flux_energie(repere_flux_thermique_utilites_MCp_kW_°C, nb_flux_utilites) = utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, nb_flux_utilites) / (utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, nb_flux_utilites) - biogaz_CHP_HT_Tfr)
            utilites_flux_energie(repere_flux_thermique_utilites_original, nb_flux_utilites) = True
        End If
        If biogaz_CHP_rdt_vapeur > zero_force Then
            nb_flux_utilites = nb_flux_utilites + 1
            'on redimensionne le tableau
            ReDim Preserve utilites_flux_energie(1 To flux_thermique_utilites_nb_caracteristiques, 1 To nb_flux_utilites)
            'on remplit les données
            utilites_flux_energie(repere_flux_thermique_utilites_etat, nb_flux_utilites) = True
            utilites_flux_energie(repere_flux_thermique_utilites_nom, nb_flux_utilites) = Feuil41.Range("thermique_biogaz_CHP_fumees_nom_flux").Cells(1, 1)
            utilites_flux_energie(repere_flux_thermique_utilites_type, nb_flux_utilites) = "Ch"
            utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, nb_flux_utilites) = biogaz_CHP_fumees_Tch
            utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, nb_flux_utilites) = biogaz_CHP_rate * biogaz_CHP_rdt_vapeur * energie_primaire_biogaz / pd_NOMBRE_HEURE_PAR_JOUR
            utilites_flux_energie(repere_flux_thermique_utilites_MCp_kW_°C, nb_flux_utilites) = utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, nb_flux_utilites) / (utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, nb_flux_utilites) - biogaz_CHP_fumees_Tfr)
            utilites_flux_energie(repere_flux_thermique_utilites_original, nb_flux_utilites) = True
        End If
    End If
    If biogaz_chaudiere_rate > zero_force And biogaz_chaudiere_rdt > zero_force Then
        nb_flux_utilites = nb_flux_utilites + 1
        'on redimensionne le tableau
        ReDim Preserve utilites_flux_energie(1 To flux_thermique_utilites_nb_caracteristiques, 1 To nb_flux_utilites)
        'on remplit les données
        utilites_flux_energie(repere_flux_thermique_utilites_etat, nb_flux_utilites) = True
        utilites_flux_energie(repere_flux_thermique_utilites_nom, nb_flux_utilites) = Feuil41.Range("thermique_biogaz_chaudiere_nom_flux").Cells(1, 1)
        utilites_flux_energie(repere_flux_thermique_utilites_type, nb_flux_utilites) = "Ch"
        utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, nb_flux_utilites) = biogaz_chaudiere_Tch
        utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, nb_flux_utilites) = biogaz_chaudiere_rate * biogaz_chaudiere_rdt * energie_primaire_biogaz / pd_NOMBRE_HEURE_PAR_JOUR
        utilites_flux_energie(repere_flux_thermique_utilites_MCp_kW_°C, nb_flux_utilites) = utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, nb_flux_utilites) / (utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, nb_flux_utilites) - biogaz_chaudiere_Tfr)
        utilites_flux_energie(repere_flux_thermique_utilites_original, nb_flux_utilites) = True
    End If


End Sub


Private Sub Class_Initialize()

Dim FeuilleActive As Worksheet

    'MSM 05/12/12
    Set FeuilleActive = ActiveSheet
    Sheets("Results0").Visible = xlSheetVisible
    Dim feuille As Worksheet
    For Each feuille In ThisWorkbook.Worksheets
        If feuille.CodeName = "Feuil_Results11" Or feuille.CodeName = "Feuil_Results12" Or feuille.CodeName = "Feuil_Results13" Then
            feuille.Visible = xlSheetVisible
        End If
    Next feuille
    FeuilleActive.Activate

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

    'MSM 05/12/12
    Sheets("Results0").Visible = xlSheetVeryHidden
    Dim feuille As Worksheet
    For Each feuille In ThisWorkbook.Worksheets
        If feuille.CodeName = "Feuil_Results11" Or feuille.CodeName = "Feuil_Results12" Or feuille.CodeName = "Feuil_Results13" Then
            feuille.Visible = xlSheetVeryHidden
        End If
    Next feuille

End Sub












