Attribute VB_Name = "HVAC_bat_elec"
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
'choix refroidissement par ventilation
Dim choix_ventil As Integer
Private Const REFROID_VENTIL_OUI = 1
Private Const REFROID_VENTIL_NON = 2
'choix du type de climatisation
Dim choix_clim As Integer
Private Const TYPE_CLIM_NON = 1
Private Const TYPE_CLIM_GF = 2
Private Const TYPE_CLIM_PAC = 3
'choix des surfaces totales
Dim surfaceBat_ventil As Double    ' [m2]
Dim surfaceBat_clim As Double    ' [m2]
'puissances équipements
Dim puissanceEquipement_ventil As Double
Dim puissanceEquipement_clim As Double

'PARAMETRES
Dim PatmAmbiante As Double
Dim mois As Integer
Dim tempDesignEte As Double    'clim
Dim T_air_mois(1 To 12) As Double  'moyenne mensuelle de la température journalière de l'air en °C
Dim HR_mois(1 To 12) As Double  'moyenne mensuelle de l'humidité relative %
Dim HR_mois_g_kg(1 To 12) As Double  'moyenne mensuelle de l'humidité relative [g eau/kg air]
Dim Patm_mois(1 To 12) As Double    'moyenne mensuelle de la pression atmosphérique [kPa]
Dim chaleurVolAir As Double     'chaleur volumique de l'air [Wh/m3/°C]
Dim chaleurVapEau As Double    ' chaleur de vaporisation de l'eau [Wh/g]

Dim HRAmbianteHiver As Double
Dim HRAmbianteEte As Double
'temps de fonctionnement
Dim tpsFonctionnementDeshum As Double
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
Dim electricite_ventilation As Double
Dim electricite_ventilation_clim As Double

Dim DJRefroidissement As Double
Dim DHRefroidissement As Double
Dim nbJoursRefroid As Double
Dim refroidissementAir As Double

Dim GJDeshumidification As Double
Dim GHDeshumidification As Double
Dim nbJoursDeshum As Double
Dim deshumidificationAir As Double

Dim Qair_ventil As Double
Dim Qair_clim As Double
Dim PuissDissEquipement_ventil As Double
Dim PuissDissEquipement_clim As Double
Dim puissanceEU_clim As Double

'PARAMETRES RESTITUES
Dim hauteurPlafond_ventil As Double
Dim volumeBat_ventil As Double
Dim tauxRenouvellementAir_ventil As Double
Dim tempAmbiante_ventil As Double
Dim rendementGlobalEquipement_ventil As Double
Dim consoElecAnnuelleRefroidVentil As Double

Dim hauteurPlafond_clim As Double
Dim volumeBat_clim As Double
Dim rendementGlobalEquipement_clim As Double
Dim tauxRenouvellementAir_clim As Double
Dim tempAmbiante_clim As Double
Dim besoinAnnuelClimatisation_m2 As Double
Dim debitEUPAC_clim As Double
Dim deltaT_clim As Double
Dim COP_froid As Double
Dim consoElecAnnuelleClimatisation As Double

Public consoElecTotale As Double


'BOOLEEN PARAMETRES FORCES
Dim hauteurPlafond_ventil_force As Boolean
Dim tauxRenouvellementAir_ventil_force As Boolean
Dim tempAmbiante_ventil_force As Boolean
Dim rendementGlobalEquipement_ventil_force As Boolean

Dim hauteurPlafond_clim_force As Boolean
Dim tauxRenouvellementAir_clim_force As Boolean
Dim tempAmbiante_clim_force As Boolean
Dim rendementGlobalEquipement_clim_force As Boolean
Dim besoinAnnuelClimatisation_m2_force As Boolean

Dim deltaT_clim_force As Boolean
Dim COP_froid_force As Boolean

'PARAMETRES FORCES
Dim hauteurPlafond_ventil_f As Double
Dim tauxRenouvellementAir_ventil_f As Double
Dim tempAmbiante_ventil_f As Double
Dim rendementGlobalEquipement_ventil_f As Double

Dim hauteurPlafond_clim_f As Double
Dim tauxRenouvellementAir_clim_f As Double
Dim tempAmbiante_clim_f As Double
Dim rendementGlobalEquipement_clim_f As Double
Dim besoinAnnuelClimatisation_m2_f As Double

Dim deltaT_clim_f As Double
Dim COP_froid_f As Double





Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'choix station météo
    identifiant_station_meteo = Feuil5.Range("choixIdStationBDClimat").Cells(1, pi_FiliereConsideree)
    ligne_station_meteo = identifiant_station_meteo + 1    ' +1 pour correspondre à la ligne

    'choix_ventil
    choix_ventil = Feuil73.Range("HVAC_bat_elec_choix_ventil").Cells(1, pi_FiliereConsideree)

    'choix_clim
    choix_clim = Feuil73.Range("HVAC_bat_elec_choix_clim").Cells(1, pi_FiliereConsideree)

    'surfaces locaux
    surfaceBat_ventil = Feuil73.Range("HVAC_bat_elec_surfaceBat_ventil").Cells(1, pi_FiliereConsideree)
    surfaceBat_clim = Feuil73.Range("HVAC_bat_elec_surfaceBat_clim").Cells(1, pi_FiliereConsideree)

    'puissances équipements
    puissanceEquipement_ventil = Feuil73.Range("HVAC_bat_elec_puissanceEquipement_ventil").Cells(1, pi_FiliereConsideree)
    puissanceEquipement_clim = Feuil73.Range("HVAC_bat_elec_puissanceEquipement_clim").Cells(1, pi_FiliereConsideree)

    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("73.1")
    pn_EffacementResultatsEnCours = True

End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    '#ventil
    hauteurPlafond_ventil_force = Feuil73.Range("HVAC_bat_elec_hauteurPlafond_ventil").Cells(1, pi_FiliereConsideree)
    tauxRenouvellementAir_ventil_force = Feuil73.Range("HVAC_bat_elec_tauxRenouvellementAir_ventil").Cells(1, pi_FiliereConsideree)
    tempAmbiante_ventil_force = Feuil73.Range("HVAC_bat_elec_tempAmbiante_ventil").Cells(1, pi_FiliereConsideree)
    rendementGlobalEquipement_ventil_force = Feuil73.Range("HVAC_bat_elec_rendementGlobalEquipement_ventil").Cells(1, pi_FiliereConsideree)
    '#clim
    hauteurPlafond_clim_force = Feuil73.Range("HVAC_bat_elec_hauteurPlafond_clim").Cells(1, pi_FiliereConsideree)
    tauxRenouvellementAir_clim_force = Feuil73.Range("HVAC_bat_elec_tauxRenouvellementAir_clim").Cells(1, pi_FiliereConsideree)
    tempAmbiante_clim_force = Feuil73.Range("HVAC_bat_elec_tempAmbiante_clim").Cells(1, pi_FiliereConsideree)
    rendementGlobalEquipement_clim_force = Feuil73.Range("HVAC_bat_elec_rendementGlobalEquipement_clim").Cells(1, pi_FiliereConsideree)
    besoinAnnuelClimatisation_m2_force = Feuil73.Range("HVAC_bat_elec_besoinAnnuelClimatisation_m2").Cells(1, pi_FiliereConsideree)
    deltaT_clim_force = Feuil73.Range("HVAC_bat_elec_deltaT_clim").Cells(1, pi_FiliereConsideree)
    COP_froid_force = Feuil73.Range("HVAC_bat_elec_COP_froid").Cells(1, pi_FiliereConsideree)


    '#ventil
    'hauteurPlafond_ventil
    If hauteurPlafond_ventil_force = True Then
        hauteurPlafond_ventil_f = Feuil73.Range("HVAC_bat_elec_hauteurPlafond_ventil").Cells(1, pi_FiliereConsideree)
    End If
    'tauxRenouvellementAir_ventil
    If tauxRenouvellementAir_ventil_force = True Then
        tauxRenouvellementAir_ventil_f = Feuil73.Range("HVAC_bat_elec_tauxRenouvellementAir_ventil").Cells(1, pi_FiliereConsideree)
    End If
    'tempAmbiante_ventil
    If tempAmbiante_ventil_force = True Then
        tempAmbiante_ventil_f = Feuil73.Range("HVAC_bat_elec_tempAmbiante_ventil").Cells(1, pi_FiliereConsideree)
    End If
    'rendementGlobalEquipement_ventil
    If rendementGlobalEquipement_ventil_force = True Then
        rendementGlobalEquipement_ventil_f = Feuil73.Range("HVAC_bat_elec_rendementGlobalEquipement_ventil").Cells(1, pi_FiliereConsideree)
    End If

    '#clim
    'hauteurPlafond_clim
    If hauteurPlafond_clim_force = True Then
        hauteurPlafond_clim_f = Feuil73.Range("HVAC_bat_elec_hauteurPlafond_clim").Cells(1, pi_FiliereConsideree)
    End If
    'tauxRenouvellementAir_clim
    If tauxRenouvellementAir_clim_force = True Then
        tauxRenouvellementAir_clim_f = Feuil73.Range("HVAC_bat_elec_tauxRenouvellementAir_clim").Cells(1, pi_FiliereConsideree)
    End If
    'tempAmbiante_clim
    If tempAmbiante_clim_force = True Then
        tempAmbiante_clim_f = Feuil73.Range("HVAC_bat_elec_tempAmbiante_clim").Cells(1, pi_FiliereConsideree)
    End If
    'rendementGlobalEquipement_clim
    If rendementGlobalEquipement_clim_force = True Then
        rendementGlobalEquipement_clim_f = Feuil73.Range("HVAC_bat_elec_rendementGlobalEquipement_clim").Cells(1, pi_FiliereConsideree)
    End If
    'besoinAnnuelClimatisation_m2
    If besoinAnnuelClimatisation_m2_force = True Then
        besoinAnnuelClimatisation_m2_f = Feuil73.Range("HVAC_bat_elec_besoinAnnuelClimatisation_m2").Cells(1, pi_FiliereConsideree)
    End If
    'deltaT_clim
    If deltaT_clim_force = True Then
        deltaT_clim_f = Feuil73.Range("HVAC_bat_elec_deltaT_clim").Cells(1, pi_FiliereConsideree)
    End If
    'COP_froid
    If COP_froid_force = True Then
        COP_froid_f = Feuil73.Range("HVAC_bat_elec_COP_froid").Cells(1, pi_FiliereConsideree)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("73.2")
    pn_EffacementResultatsEnCours = True

End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

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

    'chaleur de vaporisation de l'eau [Wh/g]
    chaleurVapEau = 0.694

    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("73.3")
    pn_EffacementResultatsEnCours = True

End Sub

Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'tempAmbiante_ventil
    If tempAmbiante_ventil_force = True Then
        tempAmbiante_ventil = tempAmbiante_ventil_f
    Else
        tempAmbiante_ventil = 35
    End If

    'tempAmbiante_clim
    If tempAmbiante_clim_force = True Then
        tempAmbiante_clim = tempAmbiante_clim_f
    Else
        tempAmbiante_clim = 25
    End If

    'hauteur du plafond local ventilé
    If hauteurPlafond_ventil_force = True Then
        hauteurPlafond_ventil = hauteurPlafond_ventil_f
    Else
        hauteurPlafond_ventil = 2.5
    End If

    'hauteur du plafond local climatisé
    If hauteurPlafond_clim_force = True Then
        hauteurPlafond_clim = hauteurPlafond_clim_f
    Else
        hauteurPlafond_clim = 2.5
    End If

    'COP_froid
    If COP_froid_force = True Then
        COP_froid = COP_froid_f
    Else
        COP_froid = 3.5
    End If

    'deltaT_clim
    If deltaT_clim_force = True Then
        deltaT_clim = deltaT_clim_f
    Else
        deltaT_clim = 4    '°C
    End If

    'rendements équipements
    If rendementGlobalEquipement_ventil_force = True Then
        rendementGlobalEquipement_ventil = rendementGlobalEquipement_ventil_f
    Else
        rendementGlobalEquipement_ventil = 0.9
    End If
    If rendementGlobalEquipement_clim_force = True Then
        rendementGlobalEquipement_clim = rendementGlobalEquipement_clim_f
    Else
        rendementGlobalEquipement_clim = 0.9
    End If

    Exit Sub

    '    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("73.4")
    pn_EffacementResultatsEnCours = True

End Sub

Sub calcul()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    '## DECLARATION DES VARIABLES
    Dim mois As Integer
    Dim jour As Integer
    Dim T_air_jour As Variant
    Dim T_air_moy_an As Double
    Dim DJ As Variant
    Dim GJ As Variant
    Dim besoinAnnuelChauffage As Double
    Dim besoinAnnuelClimatisation As Double

    '## INITIALISATION DES VARIABLES

    tpsFonctionnementDeshum = 24
    tpsFonctionnementRefroid = 24
    tpsFonctionnementVentil = 24

    ventilation_P_reseau = 2000
    ventilateur_rdt_machine = 0.7
    ventilateur_rdt_moteur = 0.8
    ventilation_rdt_global = ventilateur_rdt_machine * ventilateur_rdt_moteur

    T_air_moy_an = T_air_aspire

    If choix_ventil = REFROID_VENTIL_OUI Then

        '# Volume du local
        volumeBat_ventil = surfaceBat_ventil * hauteurPlafond_ventil       '[m3]

        '# Puissances dissipées
        PuissDissEquipement_ventil = puissanceEquipement_ventil * (1 - rendementGlobalEquipement_ventil)

        '# Taux renouvellement d'air
        If tauxRenouvellementAir_ventil_force = True Then
            tauxRenouvellementAir_ventil = tauxRenouvellementAir_ventil_f
            Qair_ventil = tauxRenouvellementAir_ventil * volumeBat_ventil
        Else
            Qair_ventil = PuissDissEquipement_ventil / (chaleurVolAir * Abs(tempAmbiante_ventil - T_air_moy_an))    'débit d'air neuf total [m3/h]
            tauxRenouvellementAir_ventil = Qair_ventil / volumeBat_ventil
        End If

        '# Ventilation  '[kWh e]
        consoElecAnnuelleRefroidVentil = (Qair_ventil * ((tempAmbiante_ventil + pd_CONVERSION_K_°C) / pd_CONVERSION_K_°C) / pd_NOMBRE_SECONDE_PAR_HEURE) * ventilation_P_reseau / ventilation_rdt_global / 1000
        consoElecAnnuelleRefroidVentil = consoElecAnnuelleRefroidVentil * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementVentil / pd_NOMBRE_HEURE_PAR_JOUR) * pd_NOMBRE_JOUR_PAR_AN     '[kWh e]

    End If



    If choix_clim = TYPE_CLIM_GF Then

        '# Volume du local
        volumeBat_clim = surfaceBat_clim * hauteurPlafond_clim       '[m3]

        '# Puissances dissipées
        PuissDissEquipement_clim = puissanceEquipement_clim * (1 - rendementGlobalEquipement_clim)

        '# Taux renouvellement d'air
        If tauxRenouvellementAir_clim_force = True Then
            tauxRenouvellementAir_clim = tauxRenouvellementAir_clim_f
            Qair_clim = tauxRenouvellementAir_clim * volumeBat_clim
        Else
            Qair_clim = PuissDissEquipement_clim / (chaleurVolAir * Abs(tempAmbiante_clim - T_air_moy_an))    'débit d'air neuf total [m3/h]
            tauxRenouvellementAir_clim = Qair_clim / volumeBat_clim
        End If

        '# Apports ventil  '[kWh u]
        DJ = degresJours("Froid", tempAmbiante_clim, T_air_mois, "Jour")      'Degrés-jours pour le refroidissement de l'air apporté par la ventilation
        DJRefroidissement = DJ(1)
        nbJoursRefroid = DJ(2)
        DHRefroidissement = DJRefroidissement * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementRefroid / pd_NOMBRE_HEURE_PAR_JOUR)
        refroidissementAir = Qair_clim * chaleurVolAir * DHRefroidissement / 1000    '[kWh u]

        '# Déhumidification de l'air  '[kWh u]
        GJ = grammesJours("Ete", tempAmbiante_clim, PatmAmbiante, HRAmbianteEte, T_air_mois, Patm_mois, HR_mois, "Mois")
        GJDeshumidification = GJ(1)
        nbJoursDeshum = GJ(2)
        GHDeshumidification = GJDeshumidification * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementDeshum / pd_NOMBRE_HEURE_PAR_JOUR)
        deshumidificationAir = Qair_clim * chaleurVapEau * GHDeshumidification / 1000     '[kWh u]

        '# Besoin total
        besoinAnnuelClimatisation = PuissDissEquipement_clim + refroidissementAir + deshumidificationAir

        '# Besoin par unité de surface
        If besoinAnnuelClimatisation_m2_force = False Then
            besoinAnnuelClimatisation_m2 = besoinAnnuelClimatisation / surfaceBat_clim
        Else
            besoinAnnuelClimatisation_m2 = besoinAnnuelClimatisation_m2_f
        End If

        '# Ventilation  '[kWh e]
        electricite_ventilation_clim = (Qair_clim * ((tempAmbiante_clim + pd_CONVERSION_K_°C) / pd_CONVERSION_K_°C) / pd_NOMBRE_SECONDE_PAR_HEURE) * ventilation_P_reseau / ventilation_rdt_global / 1000
        electricite_ventilation_clim = electricite_ventilation_clim * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementVentil / pd_NOMBRE_HEURE_PAR_JOUR) * pd_NOMBRE_JOUR_PAR_AN     '[kWh e]

        'Calcul de la consommation électrique de la clim
        consoElecAnnuelleClimatisation = (besoinAnnuelClimatisation_m2 * surfaceBat_clim / COP_froid) + electricite_ventilation_clim

    ElseIf choix_clim = TYPE_CLIM_PAC Then

        '# Volume du local
        volumeBat_clim = surfaceBat_clim * hauteurPlafond_clim       '[m3]

        '# Puissances dissipées
        PuissDissEquipement_clim = puissanceEquipement_clim * (1 - rendementGlobalEquipement_clim)

        '# Taux renouvellement d'air
        If tauxRenouvellementAir_clim_force = True Then
            tauxRenouvellementAir_clim = tauxRenouvellementAir_clim_f
            Qair_clim = tauxRenouvellementAir_clim * volumeBat_clim
        Else
            Qair_clim = PuissDissEquipement_clim / (chaleurVolAir * Abs(tempAmbiante_clim - T_air_moy_an))    'débit d'air neuf total [m3/h]
            tauxRenouvellementAir_clim = Qair_clim / volumeBat_clim
        End If

        '# Apports ventil  '[kWh u]
        DJ = degresJours("Froid", tempAmbiante_clim, T_air_mois, "Jour")      'Degrés-jours pour le refroidissement de l'air apporté par la ventilation
        DJRefroidissement = DJ(1)
        nbJoursRefroid = DJ(2)
        DHRefroidissement = DJRefroidissement * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementRefroid / pd_NOMBRE_HEURE_PAR_JOUR)
        refroidissementAir = Qair_clim * chaleurVolAir * DHRefroidissement / 1000    '[kWh u]

        '# Déhumidification de l'air  '[kWh u]
        GJ = grammesJours("Ete", tempAmbiante_clim, PatmAmbiante, HRAmbianteEte, T_air_mois, Patm_mois, HR_mois, "Mois")
        GJDeshumidification = GJ(1)
        nbJoursDeshum = GJ(2)
        GHDeshumidification = GJDeshumidification * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementDeshum / pd_NOMBRE_HEURE_PAR_JOUR)
        deshumidificationAir = Qair_clim * chaleurVapEau * GHDeshumidification / 1000     '[kWh u]

        '# Besoin total
        besoinAnnuelClimatisation = PuissDissEquipement_clim + refroidissementAir + deshumidificationAir

        '# Besoin par unité de surface
        If besoinAnnuelClimatisation_m2_force = False Then
            besoinAnnuelClimatisation_m2 = besoinAnnuelClimatisation / surfaceBat_clim
        Else
            besoinAnnuelClimatisation_m2 = besoinAnnuelClimatisation_m2_f
        End If

        '# Ventilation  '[kWh e]
        electricite_ventilation_clim = (Qair_clim * ((tempAmbiante_clim + pd_CONVERSION_K_°C) / pd_CONVERSION_K_°C) / pd_NOMBRE_SECONDE_PAR_HEURE) * ventilation_P_reseau / ventilation_rdt_global / 1000
        electricite_ventilation_clim = electricite_ventilation_clim * pd_NOMBRE_HEURE_PAR_JOUR * (tpsFonctionnementVentil / pd_NOMBRE_HEURE_PAR_JOUR) * pd_NOMBRE_JOUR_PAR_AN     '[kWh e]

        'Calcul de la consommation électrique de la clim
        consoElecAnnuelleClimatisation = (besoinAnnuelClimatisation_m2 * surfaceBat_clim / COP_froid) + electricite_ventilation_clim

        'Calcul de la puissance extraite des EU
        puissanceEU_clim = besoinAnnuelClimatisation_m2 * surfaceBat_clim * (COP_froid - 1) / COP_froid   '[kWh]
        puissanceEU_clim = puissanceEU_clim / 24

        'Calcul du débit d'EU necessaire
        debitEUPAC_clim = puissanceEU_clim / 1.16 / deltaT_clim    '[m3/h]

    Else
        besoinAnnuelClimatisation_m2 = 0
    End If

    consoElecTotale = (consoElecAnnuelleClimatisation + consoElecAnnuelleRefroidVentil) / pd_NOMBRE_JOUR_PAR_AN

    Exit Sub
    '    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("73.5")
    pn_EffacementResultatsEnCours = True

End Sub

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil73.Range("HVAC_bat_elec_hauteurPlafond_ventil").Cells(1, Colonne) = hauteurPlafond_ventil
    Feuil73.Range("HVAC_bat_elec_volumeBat_ventil").Cells(1, Colonne) = volumeBat_ventil
    Feuil73.Range("HVAC_bat_elec_tauxRenouvellementAir_ventil").Cells(1, Colonne) = tauxRenouvellementAir_ventil
    Feuil73.Range("HVAC_bat_elec_tempAmbiante_ventil").Cells(1, Colonne) = tempAmbiante_ventil
    Feuil73.Range("HVAC_bat_elec_rendementGlobalEquipement_ventil").Cells(1, Colonne) = rendementGlobalEquipement_ventil
    Feuil73.Range("HVAC_bat_elec_consoElecAnnuelleRefroidVentil").Cells(1, Colonne) = consoElecAnnuelleRefroidVentil
    Feuil73.Range("HVAC_bat_elec_hauteurPlafond_clim").Cells(1, Colonne) = hauteurPlafond_clim
    Feuil73.Range("HVAC_bat_elec_volumeBat_clim").Cells(1, Colonne) = volumeBat_clim
    Feuil73.Range("HVAC_bat_elec_tauxRenouvellementAir_clim").Cells(1, Colonne) = tauxRenouvellementAir_clim
    Feuil73.Range("HVAC_bat_elec_tempAmbiante_clim").Cells(1, Colonne) = tempAmbiante_clim
    Feuil73.Range("HVAC_bat_elec_rendementGlobalEquipement_clim").Cells(1, Colonne) = rendementGlobalEquipement_clim
    Feuil73.Range("HVAC_bat_elec_besoinAnnuelClimatisation_m2").Cells(1, Colonne) = besoinAnnuelClimatisation_m2
    Feuil73.Range("HVAC_bat_elec_deltaT_clim").Cells(1, Colonne) = deltaT_clim
    Feuil73.Range("HVAC_bat_elec_debitEUPAC_clim").Cells(1, Colonne) = debitEUPAC_clim
    Feuil73.Range("HVAC_bat_elec_COP_froid").Cells(1, Colonne) = COP_froid
    Feuil73.Range("HVAC_bat_elec_consoElecAnnuelleClimatisation").Cells(1, Colonne) = consoElecAnnuelleClimatisation
    Feuil73.Range("HVAC_bat_elec_consoElecTotale").Cells(1, Colonne) = consoElecTotale

    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("73.6")
    pn_EffacementResultatsEnCours = True

End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil73.Range("HVAC_bat_elec_hauteurPlafond_ventil").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_volumeBat_ventil").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_tauxRenouvellementAir_ventil").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_tempAmbiante_ventil").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_rendementGlobalEquipement_ventil").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_consoElecAnnuelleRefroidVentil").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_hauteurPlafond_clim").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_volumeBat_clim").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_tauxRenouvellementAir_clim").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_tempAmbiante_clim").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_rendementGlobalEquipement_clim").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_besoinAnnuelClimatisation_m2").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_deltaT_clim").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_debitEUPAC_clim").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_COP_froid").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_consoElecAnnuelleClimatisation").Cells(1, Colonne).ClearContents
    Feuil73.Range("HVAC_bat_elec_consoElecTotale").Cells(1, Colonne).ClearContents

    Exit Sub
    '    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("73.7")
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

    Call gestion_erreur_non_geree("73.8")
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






