Attribute VB_Name = "HVAC_bat_admin"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'GESTION DES CHOIX
Dim identifiant_station_meteo As Integer
Dim ligne_station_meteo As Integer
'choix du type de chauffage
Dim choix_typeChauffage As Integer
Private Const TYPE_CHAUFF_NON = 1
Private Const TYPE_CHAUFF_ELEC = 2
Private Const TYPE_CHAUFF_FIOUL = 3
Private Const TYPE_CHAUFF_GAZ = 4
Private Const TYPE_CHAUFF_PAC = 5
Private Const TYPE_CHAUFF_RECUP_PROCESS = 6
'choix du type d'appoint
Dim choix_typeAppoint As Integer
Private Const TYPE_APPOINT_ELEC = 1
Private Const TYPE_APPOINT_FIOUL = 2
Private Const TYPE_APPOINT_GAZ = 3
'choix du type de climatisation
Dim choix_typeClimatisation As Integer
Private Const TYPE_CLIM_NON = 1
Private Const TYPE_CLIM_GF = 2
Private Const TYPE_CLIM_PAC = 3
'choix du coefficient de déperdition thermique volumique
Dim choix_coefDeperdition As Double
Dim coefDeperdition As Double    '[W/°C/m3]
Dim VECT_COEF_DEPERD_TH(1 To 6) As Double
Private Const DEPERD_TH_TYPE_RT2005 = 1
Private Const DEPERD_TH_TYPE_RT2000 = 2
Private Const DEPERD_TH_TYPE_ap1980 = 3
Private Const DEPERD_TH_TYPE_MOY = 4
Private Const DEPERD_TH_TYPE_NON = 5
Private Const DEPERD_TH_TYPE_AUTRE = 6
'choix de la surface totale
Dim surfaceBat As Double    ' [m2]

'PARAMETRES
Dim PINCH As Boolean
Dim PatmAmbiante As Double
Dim tempDesignHiver As Double    'chauffage
Dim baseDJChauff As Double
Dim mois As Integer
Dim tempDesignEte As Double    'clim
Dim baseDJfroid As Double
Dim T_air_mois(1 To 12) As Double  'moyenne mensuelle de la température journalière de l'air en °C
Dim HR_mois(1 To 12) As Double  'moyenne mensuelle de l'humidité relative %
Dim HR_mois_g_kg(1 To 12) As Double  'moyenne mensuelle de l'humidité relative [g eau/kg air]
Dim Patm_mois(1 To 12) As Double    'moyenne mensuelle de la pression atmosphérique [kPa]
Dim chaleurVolAir As Double     'chaleur volumique de l'air [Wh/m3/°C]
Dim chaleurVapEau As Double     'chaleur de vaporisation de l'eau [Wh/g]

Dim QairNeufPers As Double    'débit d'air neuf par personne [m3/h]
Dim surfaceBureauPers As Double    ' surface moyenne de travail par personne [m2]

Dim HRAmbianteHiver As Double
Dim HRAmbianteEte As Double
'rendements
Dim rendementChaudiere As Double
Dim rendementChauffElec As Double
Dim rendementDistributionEmission As Double
Dim rendementRegulationChauffChaud As Double
Dim rendementRegulationChauffElec As Double
'temps de fonctionnement
Dim tpsFonctionnementRechauf As Double
Dim tpsFonctionnementDeshum As Double
Dim tpsFonctionnementHumid As Double
Dim tpsFonctionnementRefroid As Double
Dim tpsFonctionnementVentil As Double
'correction T ambiante
Dim correctionApportsEte As Double
Dim correctionApportsHiver As Double
'réseau ventil
Dim ventilation_P_reseau As Double
Dim ventilateur_rdt_machine As Double
Dim ventilateur_rdt_moteur As Double
Dim ventilation_rdt_global As Double
'conso elec ventil
Dim electricite_ventilation_hiver As Double
Dim electricite_ventilation_ete As Double
Dim electricite_ventilation_reste As Double
Dim electricite_ventilation_tot As Double

'ELEMENTS CALCULES
Dim DJChauffage As Double
Dim nbJoursChauff As Double
Dim deperditionsBat As Double

Dim DJRechauffement As Double
Dim DHRechauffement As Double
Dim nbJoursRechauf As Double
Dim rechauffementAir As Double

Dim GJHumidification As Double
Dim GHHumidification As Double
Dim nbJoursHumid As Double
Dim humidificationAir As Double

Dim DJRefroidissement As Double
Dim DHRefroidissement As Double
Dim nbJoursRefroid As Double
Dim refroidissementAir As Double

Dim GJDeshumidification As Double
Dim GHDeshumidification As Double
Dim nbJoursDeshum As Double
Dim deshumidificationAir As Double

Dim besoinAnnuelChauffage As Double
Dim besoinAnnuelClimatisation As Double

Dim Qair As Double
Dim tempEauChauffage As Double

'PARAMETRES RESTITUES
Dim hauteurPlafond As Double
Dim volumeBat As Double
Dim tauxRenouvellementAir As Double

Dim tempAmbianteHiver As Double
Dim besoinAnnuelChauffage_m2 As Double
Dim ratioDimensionnementPAC As Double
Dim rendementGlobal As Double
Dim rendementGlobal_appoint As Double
Dim deltaT As Double
Dim puissanceEU As Double
Dim debitEUPAC As Double
Dim COP_chaud As Double
Dim consoElecAnnuelleChauffage As Double
Dim consoFioulAnnuelleChauffage As Double
Dim consoGazAnnuelleChauffage As Double

Dim tempAmbianteEte As Double
Dim besoinAnnuelClimatisation_m2 As Double
Dim deltaT_clim As Double
Dim puissanceEU_clim As Double
Dim debitEUPAC_clim As Double
Dim COP_froid As Double
Dim consoElecAnnuelleClimatisation As Double

Public consoElecTotale As Double

'BOOLEEN PARAMETRES FORCES
'chauffage
Dim hauteurPlafond_force As Boolean
Dim tauxRenouvellementAir_force As Boolean
Dim tempAmbianteHiver_force As Boolean
Dim besoinAnnuelChauffage_m2_force As Boolean
Dim ratioDimensionnementPAC_force As Boolean
Dim deltaT_force As Boolean
Dim debitEUPAC_force As Boolean
Dim COP_chaud_force As Boolean
Dim rendementGlobal_force As Boolean
Dim rendementGlobal_appoint_force As Boolean
'clim
Dim tempAmbianteEte_force As Boolean
Dim besoinAnnuelClimatisation_m2_force As Boolean
Dim deltaT_clim_force As Boolean
Dim debitEUPAC_clim_force As Boolean
Dim COP_froid_force As Boolean

'PARAMETRES FORCES
'chauffage
Dim hauteurPlafond_f As Double
Dim tauxRenouvellementAir_f As Double
Dim tempAmbianteHiver_f As Double
Dim besoinAnnuelChauffage_m2_f As Double
Dim ratioDimensionnementPAC_f As Double
Dim deltaT_f As Double
Dim debitEUPAC_f As Double
Dim COP_chaud_f As Double
Dim rendementGlobal_f As Double
Dim rendementGlobal_appoint_f As Double
'clim
Dim tempAmbianteEte_f As Double
Dim besoinAnnuelClimatisation_m2_f As Double
Dim deltaT_clim_f As Double
Dim debitEUPAC_clim_f As Double
Dim COP_froid_f As Double


Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'choix station météo
    identifiant_station_meteo = Feuil5.Range("choixIdStationBDClimat").Cells(1, pi_FiliereConsideree)
    ligne_station_meteo = identifiant_station_meteo + 1    ' +1 pour correspondre à la ligne

    'choix_typeChauffage
    choix_typeChauffage = Feuil71.Range("HVAC_bat_admin_choix_typeChauffage").Cells(1, pi_FiliereConsideree)

    'choix_typeAppoint
    If choix_typeChauffage = TYPE_CHAUFF_PAC Then
        choix_typeAppoint = Feuil71.Range("HVAC_bat_admin_choix_typeAppoint").Cells(1, pi_FiliereConsideree)
    End If

    'choix_typeClimatisation
    choix_typeClimatisation = Feuil71.Range("HVAC_bat_admin_choix_typeClimatisation").Cells(1, pi_FiliereConsideree)

    'surface bâtiment
    surfaceBat = Feuil71.Range("HVAC_bat_admin_surfaceBat").Cells(1, pi_FiliereConsideree)

    'coefDeperdition
    choix_coefDeperdition = Feuil71.Range("HVAC_bat_admin_choix_coefDeperdition").Cells(1, pi_FiliereConsideree)

    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("71.1")
    pn_EffacementResultatsEnCours = True

End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    '#chauffage
    hauteurPlafond_force = Feuil71.Range("HVAC_bat_admin_hauteurPlafond").Cells(1, pi_FiliereConsideree)
    tauxRenouvellementAir_force = Feuil71.Range("HVAC_bat_admin_tauxRenouvellementAir").Cells(1, pi_FiliereConsideree)
    tempAmbianteHiver_force = Feuil71.Range("HVAC_bat_admin_tempAmbianteHiver").Cells(1, pi_FiliereConsideree)
    besoinAnnuelChauffage_m2_force = Feuil71.Range("HVAC_bat_admin_besoinAnnuelChauffage_m2").Cells(1, pi_FiliereConsideree)
    ratioDimensionnementPAC_force = Feuil71.Range("HVAC_bat_admin_ratioDimensionnementPAC").Cells(1, pi_FiliereConsideree)
    deltaT_force = Feuil71.Range("HVAC_bat_admin_deltaT").Cells(1, pi_FiliereConsideree)
    COP_chaud_force = Feuil71.Range("HVAC_bat_admin_COP_chaud").Cells(1, pi_FiliereConsideree)
    rendementGlobal_force = Feuil71.Range("HVAC_bat_admin_rendementGlobal").Cells(1, pi_FiliereConsideree)
    rendementGlobal_appoint_force = Feuil71.Range("HVAC_bat_admin_rendementGlobal_appoint").Cells(1, pi_FiliereConsideree)
    '#clim
    tempAmbianteEte_force = Feuil71.Range("HVAC_bat_admin_tempAmbianteEte").Cells(1, pi_FiliereConsideree)
    besoinAnnuelClimatisation_m2_force = Feuil71.Range("HVAC_bat_admin_besoinAnnuelClimatisation_m2").Cells(1, pi_FiliereConsideree)
    deltaT_clim_force = Feuil71.Range("HVAC_bat_admin_deltaT_clim").Cells(1, pi_FiliereConsideree)
    COP_froid_force = Feuil71.Range("HVAC_bat_admin_COP_froid").Cells(1, pi_FiliereConsideree)


    '#chauffage
    'hauteurPlafond
    If hauteurPlafond_force = True Then
        hauteurPlafond_f = Feuil71.Range("HVAC_bat_admin_hauteurPlafond").Cells(1, pi_FiliereConsideree)
    End If
    'tauxRenouvellementAir
    If tauxRenouvellementAir_force = True Then
        tauxRenouvellementAir_f = Feuil71.Range("HVAC_bat_admin_tauxRenouvellementAir").Cells(1, pi_FiliereConsideree)
    End If
    'tempAmbianteHiver
    If tempAmbianteHiver_force = True Then
        tempAmbianteHiver_f = Feuil71.Range("HVAC_bat_admin_tempAmbianteHiver").Cells(1, pi_FiliereConsideree)
    End If
    'besoinAnnuelChauffage_m2
    If besoinAnnuelChauffage_m2_force = True Then
        besoinAnnuelChauffage_m2_f = Feuil71.Range("HVAC_bat_admin_besoinAnnuelChauffage_m2").Cells(1, pi_FiliereConsideree)
    End If
    'ratioDimensionnementPAC
    If ratioDimensionnementPAC_force = True Then
        ratioDimensionnementPAC_f = Feuil71.Range("HVAC_bat_admin_ratioDimensionnementPAC").Cells(1, pi_FiliereConsideree)
    End If
    'deltaT
    If deltaT_force = True Then
        deltaT_f = Feuil71.Range("HVAC_bat_admin_deltaT").Cells(1, pi_FiliereConsideree)
    End If
    'COP_chaud
    If COP_chaud_force = True Then
        COP_chaud_f = Feuil71.Range("HVAC_bat_admin_COP_chaud").Cells(1, pi_FiliereConsideree)
    End If
    'rendementGlobal
    If rendementGlobal_force = True Then
        rendementGlobal_f = Feuil71.Range("HVAC_bat_admin_rendementGlobal").Cells(1, pi_FiliereConsideree)
    End If
    'rendementGlobal_appoint
    If rendementGlobal_appoint_force = True Then
        rendementGlobal_appoint_f = Feuil71.Range("HVAC_bat_admin_rendementGlobal_appoint").Cells(1, pi_FiliereConsideree)
    End If

    '#clim
    'tempAmbianteEte
    If tempAmbianteEte_force = True Then
        tempAmbianteEte_f = Feuil71.Range("HVAC_bat_admin_tempAmbianteEte").Cells(1, pi_FiliereConsideree)
    End If
    'besoinAnnuelClimatisation_m2
    If besoinAnnuelClimatisation_m2_force = True Then
        besoinAnnuelClimatisation_m2_f = Feuil71.Range("HVAC_bat_admin_besoinAnnuelClimatisation_m2").Cells(1, pi_FiliereConsideree)
    End If
    'deltaT_clim
    If deltaT_clim_force = True Then
        deltaT_clim_f = Feuil71.Range("HVAC_bat_admin_deltaT_clim").Cells(1, pi_FiliereConsideree)
    End If
    'COP_froid
    If COP_froid_force = True Then
        COP_froid_f = Feuil71.Range("HVAC_bat_admin_COP_froid").Cells(1, pi_FiliereConsideree)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("71.2")
    pn_EffacementResultatsEnCours = True

End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'Valeurs du coefficient de déperdition thermique [W/°C/m3]
    VECT_COEF_DEPERD_TH(DEPERD_TH_TYPE_RT2005) = 0.65
    VECT_COEF_DEPERD_TH(DEPERD_TH_TYPE_RT2000) = 0.75
    VECT_COEF_DEPERD_TH(DEPERD_TH_TYPE_ap1980) = 0.9
    VECT_COEF_DEPERD_TH(DEPERD_TH_TYPE_MOY) = 1.2
    VECT_COEF_DEPERD_TH(DEPERD_TH_TYPE_NON) = 1.8

    For mois = 1 To 12
        T_air_mois(mois) = Feuil68.Range("AirT_Jan").Cells(ligne_station_meteo, mois).Value    'température moyenne de l'air extérieur (base de donnée) [°C]
        Patm_mois(mois) = Feuil68.Range("Press_Jan").Cells(ligne_station_meteo, mois).Value    'P atm moyenne(base de donnée) [kPa]
        HR_mois(mois) = Feuil68.Range("Hum_Jan").Cells(ligne_station_meteo, mois).Value    'humidité relative moyenne de l'air extérieur (base de donnée) [%]
    Next
    PatmAmbiante = (101325 * (288.15 / (288.15 - 0.0065 * altitude)) ^ -5.25588) / 1000

    'HR
    HRAmbianteHiver = 50
    HRAmbianteEte = 50

    'chaleur volumique de l'air [Wh/m3/°C]
    chaleurVolAir = 0.34

    'tempEauChauffage
    tempEauChauffage = 60    '°C

    'chaleur de vaporisation de l'eau [Wh/g]
    chaleurVapEau = 0.694

    'débit d'air neuf par personne [m3/h]
    QairNeufPers = 25    'Valeur valable pour un bureau, sans travail physique
    'Code du travail Article R232-5-3, Décret 2008-244 du 7/4/2008 art 9(V)

    'surface moyenne de travail par personne [m2]
    surfaceBureauPers = 9
    'Norme NF X35-102 : Dimension des espaces de travail en bureau (9m2/1p; 12M2/2p; 21m2/3p)

    'rendement de la chaudière
    rendementChaudiere = 0.85

    'rendement de la distribution du chauffage
    rendementDistributionEmission = 0.95

    'rendement régulation du chauffage chaudière
    rendementRegulationChauffChaud = 0.95

    'rendement du chauffage électrique
    rendementChauffElec = 0.98

    'rendement régulation du chauffage électrique
    rendementRegulationChauffElec = 0.98

    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("71.3")
    pn_EffacementResultatsEnCours = True

End Sub

Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'Isolation
    If choix_coefDeperdition <> DEPERD_TH_TYPE_AUTRE Then
        coefDeperdition = VECT_COEF_DEPERD_TH(choix_coefDeperdition)
    ElseIf choix_coefDeperdition = DEPERD_TH_TYPE_AUTRE Then
        coefDeperdition = Feuil71.Range("HVAC_bat_admin_coefDeperdition").Cells(1, pi_FiliereConsideree)
    End If

    '    'Température de design pour le chauffage
    '    If tempDesignHiver_force = True Then
    '        tempDesignHiver = tempDesignHiver_f
    '    Else
    '        tempDesignHiver = Feuil68.Range("Hdesign").Cells(ligne_station_meteo, 1).Value
    '    End If
    '
    '    'Température de design pour la clim
    '    If tempDesignEte_force = True Then
    '        tempDesignEte = tempDesignHiver_f
    '    Else
    '        tempDesignEte = Feuil68.Range("Cdesign").Cells(ligne_station_meteo, 1).Value
    '    End If

    'Température de confort en été (pour la clim)
    If tempAmbianteEte_force = True Then
        tempAmbianteEte = tempAmbianteEte_f
    Else
        tempAmbianteEte = 25
    End If

    'Température de confort en hiver (pour le chauffage)
    If tempAmbianteHiver_force = True Then
        tempAmbianteHiver = tempAmbianteHiver_f
    Else
        tempAmbianteHiver = 19
    End If

    'hauteur du plafond pour calculer le volume
    If hauteurPlafond_force = True Then
        hauteurPlafond = hauteurPlafond_f
    Else
        hauteurPlafond = 2.5
    End If

    'rendement global
    If rendementGlobal_force = True Then
        rendementGlobal = rendementGlobal_f
    Else
        If choix_typeChauffage <> TYPE_CHAUFF_ELEC Then
            rendementGlobal = rendementChaudiere * rendementDistributionEmission * rendementRegulationChauffChaud
        Else
            rendementGlobal = rendementChauffElec * rendementRegulationChauffElec
        End If
    End If

    'rendement global appoint
    If rendementGlobal_appoint_force = True Then
        rendementGlobal_appoint = rendementGlobal_appoint_f
    Else
        If choix_typeAppoint <> TYPE_APPOINT_ELEC Then
            rendementGlobal_appoint = rendementChaudiere * rendementDistributionEmission * rendementRegulationChauffChaud
        Else
            rendementGlobal_appoint = rendementChauffElec * rendementRegulationChauffElec
        End If
    End If

    'COP froid
    If COP_chaud_force = True Then
        COP_chaud = COP_chaud_f
    Else
        COP_chaud = -0.09 * tempEauChauffage + 8.7
        'COP = -0.1 * tempEauChauffage + 9
    End If

    'COP froid
    If COP_froid_force = True Then
        COP_froid = COP_froid_f
    Else
        COP_froid = 3.5
    End If

    'ratioDimensionnementPAC
    If ratioDimensionnementPAC_force = True Then
        ratioDimensionnementPAC = ratioDimensionnementPAC_f
    Else
        ratioDimensionnementPAC = 1
    End If

    'deltaT
    If deltaT_force = True Then
        deltaT = deltaT_f
    Else
        deltaT = 2    '°C
    End If

    'deltaT_clim
    If deltaT_clim_force = True Then
        deltaT_clim = deltaT_clim_f
    Else
        deltaT_clim = 2    '°C
    End If

    Exit Sub

    '    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("71.4")
    pn_EffacementResultatsEnCours = True

End Sub

Sub calcul()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    '## DECLARATION DES VARIABLES
    Dim mois As Integer
    Dim jour As Integer
    Dim T_air_jour As Variant
    Dim DJ As Variant
    Dim GJ As Variant
    Dim T_moy_max_min As Variant
    Dim T_moy_max As Double
    Dim T_moy_min As Double
    Dim tempAmbianteHiverCorr As Double
    Dim tempAmbianteEteCorr As Double

    '## INITIALISATION DES VARIABLES

    PINCH = False

    'Calcul du volume du bâtiment
    volumeBat = surfaceBat * hauteurPlafond    '[m3]

    'Taux renouvellement d'air
    If tauxRenouvellementAir_force = True Then
        tauxRenouvellementAir = tauxRenouvellementAir_f
        Qair = tauxRenouvellementAir * volumeBat
    Else
        'Calcul du débit d'air neuf en m3/h en fonction du débit d'air neuf nécessaire par personne et du nombre de personnes occupant la surface du bâtiment (1 personne pour 9 m2)
        Qair = QairNeufPers * surfaceBat / surfaceBureauPers    'débit d'air neuf total [m3/h]
        tauxRenouvellementAir = Qair / volumeBat    'm3 air neuf /(m3 bâtiment . heure) --> taux/heure
    End If

    tpsFonctionnementRechauf = 24
    tpsFonctionnementDeshum = 24
    tpsFonctionnementHumid = 24
    tpsFonctionnementRefroid = 24
    tpsFonctionnementVentil = 24

    T_moy_max_min = moy_6max_6min(T_air_mois)
    T_moy_max = T_moy_max_min(1)
    T_moy_min = T_moy_max_min(2)

    correctionApportsHiver = (3 / 5) * T_moy_min
    tempAmbianteHiverCorr = tempAmbianteHiver - Abs(correctionApportsHiver)
    correctionApportsEte = (3 / 5) * T_moy_max
    tempAmbianteEteCorr = tempAmbianteEte - Abs(correctionApportsEte)

    ventilation_P_reseau = 2000
    ventilateur_rdt_machine = 0.7
    ventilateur_rdt_moteur = 0.8
    ventilation_rdt_global = ventilateur_rdt_machine * ventilateur_rdt_moteur



    '********** CHAUFFAGE ***********
    '# Déperditions du bâtiment  '[kWh u]
    DJ = degresJours("Chaud", tempAmbianteHiverCorr, T_air_mois, "Jour")        'Degrés-jours pour le chauffage
    DJChauffage = DJ(1)
    nbJoursChauff = DJ(2)
    deperditionsBat = coefDeperdition * volumeBat * pd_NOMBRE_HEURE_PAR_JOUR * DJChauffage / 1000    '[kWh u]

    '# Déperditions ventilation  '[kWh u]
    DJ = degresJours("Chaud", tempAmbianteHiverCorr, T_air_mois, "Jour")       'Degrés-jours pour le réchauffement de l'air apporté par la ventilation
    DJRechauffement = DJ(1)
    nbJoursRechauf = DJ(2)
    DHRechauffement = DJRechauffement * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementRechauf / pd_NOMBRE_HEURE_PAR_JOUR)
    rechauffementAir = Qair * chaleurVolAir * DHRechauffement / 1000     '[kWh u]

    '# humidification de l'air   '[kWh u]
    GJ = grammesJours("Hiver", tempAmbianteHiverCorr, PatmAmbiante, HRAmbianteHiver, T_air_mois, Patm_mois, HR_mois, "Mois")
    GJHumidification = GJ(1)
    nbJoursHumid = GJ(2)
    GHHumidification = GJHumidification * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementHumid / pd_NOMBRE_HEURE_PAR_JOUR)
    humidificationAir = Qair * chaleurVapEau * GHHumidification / 1000     '[kWh u]

    '# Ventilation '[kWh e]
    ventilation_rdt_global = ventilateur_rdt_machine * ventilateur_rdt_moteur
    electricite_ventilation_hiver = (Qair * ((tempAmbianteHiver + pd_CONVERSION_K_°C) / pd_CONVERSION_K_°C) / pd_NOMBRE_SECONDE_PAR_HEURE) * ventilation_P_reseau / ventilation_rdt_global / 1000
    electricite_ventilation_hiver = electricite_ventilation_hiver * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementVentil / pd_NOMBRE_HEURE_PAR_JOUR) * nbJoursRechauf    '[kWh e]

    '********** CLIMATISATION ***********
    '# Apports ventil  '[kWh u]
    DJ = degresJours("Froid", tempAmbianteEteCorr, T_air_mois, "Jour")      'Degrés-jours pour le refroidissement de l'air apporté par la ventilation
    DJRefroidissement = DJ(1)
    nbJoursRefroid = DJ(2)
    DHRefroidissement = DJRefroidissement * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementRefroid / pd_NOMBRE_HEURE_PAR_JOUR)
    refroidissementAir = Qair * chaleurVolAir * DHRefroidissement / 1000    '[kWh u]

    '# déhumidification de l'air  '[kWh u]
    GJ = grammesJours("Ete", tempAmbianteEteCorr, PatmAmbiante, HRAmbianteEte, T_air_mois, Patm_mois, HR_mois, "Mois")
    GJDeshumidification = GJ(1)
    nbJoursDeshum = GJ(2)
    GHDeshumidification = GJDeshumidification * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementDeshum / pd_NOMBRE_HEURE_PAR_JOUR)
    deshumidificationAir = Qair * chaleurVapEau * GHDeshumidification / 1000     '[kWh u]

    '# Ventilation  '[kWh e]
    electricite_ventilation_ete = (Qair * ((tempAmbianteEte + pd_CONVERSION_K_°C) / pd_CONVERSION_K_°C) / pd_NOMBRE_SECONDE_PAR_HEURE) * ventilation_P_reseau / ventilation_rdt_global / 1000
    electricite_ventilation_ete = electricite_ventilation_ete * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementVentil / pd_NOMBRE_HEURE_PAR_JOUR) * nbJoursRefroid    '[kWh e]

    '********** HORS SAISON ***********
    '# Reste ventilation
    If (pd_NOMBRE_JOUR_PAR_AN - (nbJoursRefroid + nbJoursRechauf)) > 0 Then
        electricite_ventilation_reste = (Qair * (((tempAmbianteEte + tempAmbianteHiver) / 2 + pd_CONVERSION_K_°C) / pd_CONVERSION_K_°C) / pd_NOMBRE_SECONDE_PAR_HEURE) * ventilation_P_reseau / ventilation_rdt_global / 1000
        electricite_ventilation_reste = electricite_ventilation_reste * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementVentil / pd_NOMBRE_HEURE_PAR_JOUR) * (pd_NOMBRE_JOUR_PAR_AN - (nbJoursRefroid + nbJoursRechauf))
    Else
        electricite_ventilation_reste = 0
    End If


    '#Détermination du besoin annuelle de chauffage
    besoinAnnuelChauffage = deperditionsBat + rechauffementAir + humidificationAir   '[kWh th /an]
    '# Détermination du besoin annuel de climatisation
    besoinAnnuelClimatisation = refroidissementAir + deshumidificationAir  '[kWh th /an]


    '*********** CALCUL DES CONSOMMATIONS *************
    If choix_typeChauffage <> TYPE_CHAUFF_NON And besoinAnnuelChauffage <> 0 Then

        If besoinAnnuelChauffage_m2_force = False Then
            besoinAnnuelChauffage_m2 = besoinAnnuelChauffage / surfaceBat    'Calcul des besoins thermiques annuels  [kWh th/m²/an]
        Else
            besoinAnnuelChauffage_m2 = besoinAnnuelChauffage_m2_f
            besoinAnnuelChauffage = besoinAnnuelChauffage_m2 * surfaceBat
        End If

        'Calcul des consommations du chauffage
        If choix_typeChauffage = TYPE_CHAUFF_ELEC Then
            consoElecAnnuelleChauffage = besoinAnnuelChauffage_m2 * surfaceBat / rendementGlobal  '[kWh elec]
        ElseIf choix_typeChauffage = TYPE_CHAUFF_FIOUL Then
            consoFioulAnnuelleChauffage = besoinAnnuelChauffage_m2 * surfaceBat / (pd_PCI_FIOUL_kWh_L * rendementGlobal)    '[l]
        ElseIf choix_typeChauffage = TYPE_CHAUFF_GAZ Then
            consoGazAnnuelleChauffage = besoinAnnuelChauffage_m2 * surfaceBat / (pd_PCI_GAZ_NATUREL_kWh_Nm3 * rendementGlobal)    '[m3]
        ElseIf choix_typeChauffage = TYPE_CHAUFF_PAC Then

            '## Calcul de la consommation électrique de la PAC et du débit d'EU nécessaire

            'Calcul de la conso d'élec de la PAC
            consoElecAnnuelleChauffage = besoinAnnuelChauffage_m2 * surfaceBat * ratioDimensionnementPAC / COP_chaud    '[kWh elec]

            'Calcul de la puissance extraite des EU
            puissanceEU = besoinAnnuelChauffage_m2 * surfaceBat * ratioDimensionnementPAC * (COP_chaud - 1) / COP_chaud  '[kWh]
            puissanceEU = puissanceEU / 24

            'Calcul du débit d'EU necessaire
            debitEUPAC = puissanceEU / (1.16 * deltaT)    '[m3/h]
            '<-> P(W) = m * CPeau * delta_T
            '<-> P(W) = (Q * rho_eau)/3600 * 4180 * delta_T ;
            '<-> P(W) = (Q * 10^3) * (4180/3600) * delta_T
            '<-> P (kW) = Q * 1.16 * delta_T

            '## Calcul des consommations du chauffage d'appoint
            If choix_typeAppoint = TYPE_APPOINT_ELEC Then
                consoElecAnnuelleChauffage = consoElecAnnuelleChauffage + besoinAnnuelChauffage_m2 * surfaceBat * (1 - ratioDimensionnementPAC) / rendementGlobal_appoint     '[kWh elec]
            ElseIf choix_typeAppoint = TYPE_APPOINT_FIOUL Then
                consoFioulAnnuelleChauffage = besoinAnnuelChauffage_m2 * surfaceBat * (1 - ratioDimensionnementPAC) / (pd_PCI_FIOUL_kWh_L * rendementGlobal_appoint)     '[l]
            ElseIf choix_typeAppoint = TYPE_APPOINT_GAZ Then
                consoGazAnnuelleChauffage = besoinAnnuelChauffage_m2 * surfaceBat * (1 - ratioDimensionnementPAC) / (pd_PCI_GAZ_NATUREL_kWh_Nm3 * rendementGlobal_appoint)      '[m3]
            End If

        ElseIf choix_typeChauffage = TYPE_CHAUFF_RECUP_PROCESS Then
            PINCH = True
        End If

    Else
        besoinAnnuelChauffage_m2 = 0
    End If



    If choix_typeClimatisation <> TYPE_CLIM_NON And besoinAnnuelClimatisation <> 0 Then

        If besoinAnnuelClimatisation_m2_force = False Then
            besoinAnnuelClimatisation_m2 = besoinAnnuelClimatisation / surfaceBat    'Calcul des besoins thermiques annuels  [kWh th/m²/an]
        Else
            besoinAnnuelClimatisation_m2 = besoinAnnuelClimatisation_m2_f
        End If

        If choix_typeClimatisation = TYPE_CLIM_GF Then
            'Calcul de la consommation électrique de la clim
            consoElecAnnuelleClimatisation = (besoinAnnuelClimatisation_m2 * surfaceBat / COP_froid) + electricite_ventilation_ete

            '            'Calcul de la puissance extraite des EU
            '            puissanceEU_clim = besoinAnnuelClimatisation_m2 * surfaceBat * (COP_froid - 1) / COP_froid   '[kWh]
            '            puissanceEU_clim = puissanceEU_clim / (nbJoursRefroid * 24)

        ElseIf choix_typeClimatisation = TYPE_CLIM_PAC Then
            'Calcul de la consommation électrique de la clim
            consoElecAnnuelleClimatisation = (besoinAnnuelClimatisation_m2 * surfaceBat / COP_froid) + electricite_ventilation_ete

            'Calcul de la puissance extraite des EU
            puissanceEU_clim = besoinAnnuelClimatisation_m2 * surfaceBat * ratioDimensionnementPAC * (COP_froid - 1) / COP_froid  '[kWh]
            puissanceEU_clim = puissanceEU_clim / 24

            'Calcul du débit d'EU necessaire
            debitEUPAC_clim = puissanceEU_clim / 1.16 / deltaT_clim    '[m3/h]

        End If

    Else
        besoinAnnuelClimatisation_m2 = 0
    End If

    'Calcul de la consommation électrique journalière totale
    consoElecTotale = (consoElecAnnuelleChauffage + consoElecAnnuelleClimatisation + electricite_ventilation_reste) / pd_NOMBRE_JOUR_PAR_AN



    Exit Sub
    '    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("71.5")
    pn_EffacementResultatsEnCours = True

End Sub
Sub transfert_flux_energie_vers_pinch()

If PINCH = True Then
    Dim nom_flux As String
    Dim T_in As Double
    Dim T_out As Double
    Dim MCp As Double
    Dim m As Double
    Dim Cp As Double

    'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
    ' CHAUFFAGE - Bâtiments d'exploitation
    nom_flux = Feuil5.Range("nom_flux_chauffageClimBatAdmin1").Cells(1, 1).Value
    T_in = 35
    T_out = tempEauChauffage
    MCp = besoinAnnuelChauffage / (rendementGlobal * pd_NOMBRE_JOUR_PAR_AN * Abs(T_out - T_in))    'kWh/j/°C
    'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
    Call copie_dans_flux_thermique_process(nom_flux, T_in, T_out, MCp)
End If

End Sub
Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil71.Range("HVAC_bat_admin_coefDeperdition").Cells(1, Colonne) = coefDeperdition
    Feuil71.Range("HVAC_bat_admin_hauteurPlafond").Cells(1, Colonne) = hauteurPlafond
    Feuil71.Range("HVAC_bat_admin_volumeBat").Cells(1, Colonne) = volumeBat
    Feuil71.Range("HVAC_bat_admin_tauxRenouvellementAir").Cells(1, Colonne) = tauxRenouvellementAir

    Feuil71.Range("HVAC_bat_admin_tempAmbianteHiver").Cells(1, Colonne) = tempAmbianteHiver
    Feuil71.Range("HVAC_bat_admin_besoinAnnuelChauffage_m2").Cells(1, Colonne) = besoinAnnuelChauffage_m2
    Feuil71.Range("HVAC_bat_admin_ratioDimensionnementPAC").Cells(1, Colonne) = ratioDimensionnementPAC
    Feuil71.Range("HVAC_bat_admin_rendementGlobal").Cells(1, Colonne) = rendementGlobal
    Feuil71.Range("HVAC_bat_admin_rendementGlobal_appoint").Cells(1, Colonne) = rendementGlobal_appoint
    Feuil71.Range("HVAC_bat_admin_deltaT").Cells(1, Colonne) = deltaT
    Feuil71.Range("HVAC_bat_admin_debitEUPAC").Cells(1, Colonne) = debitEUPAC
    Feuil71.Range("HVAC_bat_admin_COP_chaud").Cells(1, Colonne) = COP_chaud
    Feuil71.Range("HVAC_bat_admin_consoElecAnnuelleChauffage").Cells(1, Colonne) = consoElecAnnuelleChauffage
    Feuil71.Range("HVAC_bat_admin_consoFioulAnnuelleChauffage").Cells(1, Colonne) = consoFioulAnnuelleChauffage
    Feuil71.Range("HVAC_bat_admin_consoGazAnnuelleChauffage").Cells(1, Colonne) = consoGazAnnuelleChauffage

    Feuil71.Range("HVAC_bat_admin_tempAmbianteEte").Cells(1, Colonne) = tempAmbianteEte
    Feuil71.Range("HVAC_bat_admin_besoinAnnuelClimatisation_m2").Cells(1, Colonne) = besoinAnnuelClimatisation_m2
    Feuil71.Range("HVAC_bat_admin_deltaT_clim").Cells(1, Colonne) = deltaT_clim
    Feuil71.Range("HVAC_bat_admin_debitEUPAC_clim").Cells(1, Colonne) = debitEUPAC_clim
    Feuil71.Range("HVAC_bat_admin_COP_froid").Cells(1, Colonne) = COP_froid
    Feuil71.Range("HVAC_bat_admin_consoElecAnnuelleClimatisation").Cells(1, Colonne) = consoElecAnnuelleClimatisation

    Feuil71.Range("HVAC_bat_admin_consoElecTotale").Cells(1, Colonne) = consoElecTotale

    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("71.6")
    pn_EffacementResultatsEnCours = True

End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil71.Range("HVAC_bat_admin_coefDeperdition").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_hauteurPlafond").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_volumeBat").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_tauxRenouvellementAir").Cells(1, Colonne).ClearContents

    Feuil71.Range("HVAC_bat_admin_tempAmbianteHiver").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_besoinAnnuelChauffage_m2").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_ratioDimensionnementPAC").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_rendementGlobal").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_rendementGlobal_appoint").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_deltaT").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_debitEUPAC").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_COP_chaud").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_consoElecAnnuelleChauffage").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_consoFioulAnnuelleChauffage").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_consoGazAnnuelleChauffage").Cells(1, Colonne).ClearContents

    Feuil71.Range("HVAC_bat_admin_tempAmbianteEte").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_besoinAnnuelClimatisation_m2").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_deltaT_clim").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_debitEUPAC_clim").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_COP_froid").Cells(1, Colonne).ClearContents
    Feuil71.Range("HVAC_bat_admin_consoElecAnnuelleClimatisation").Cells(1, Colonne).ClearContents

    Feuil71.Range("HVAC_bat_admin_consoElecTotale").Cells(1, Colonne).ClearContents

    Exit Sub
    '    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("71.7")
    pn_EffacementResultatsEnCours = True

End Sub

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

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

    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("71.8")
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

