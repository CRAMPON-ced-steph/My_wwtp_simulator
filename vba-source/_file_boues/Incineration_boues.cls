Attribute VB_Name = "Incineration_boues"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'constante pour le choix des entrées
Private Const nb_inlet = 2

Private Const critere_convergence = 0.00001

'Constates pour le choix de traitement SNCR (Selective Non Catalytic Reaction) dans le four (contre les NOx)
Private Const mi_PAS_DE_TRAITEMENT_SNCR As Integer = 0
Private Const mi_TRAITEMENT_SNCR_AVEC_AMMONIAQUE As Integer = 1
Private Const mi_TRAITEMENT_SNCR_AVEC_UREE As Integer = 2

'constantes pour les choix de combustible
Private Const combustible_biogaz = 0
'combustibles externes --> constantes publiques définies dans programme_principal (combustible_repere_gaz_naturel/combustible_repere_fioul)

'constante pour le choix de la filière de traitement des fumées
Private Const fumees_filiere_humide = 0
Private Const fumees_filiere_seche = 1

'constantes pour les choix de boues en entree
Private Const inlet_aucun = 0

'choix de la synthèse
Dim choix_inlet(nb_inlet) As Integer
Dim inlet_etape_anterieure(nb_inlet) As Integer
Dim inlet_type_boues(nb_inlet) As Double
Dim inlet_ratio_admis(nb_inlet) As Double
Dim choix_combustible As Integer
Private mi_ChoixTraitementNOxFour As Integer
Dim choix_nb_jours_travailles As Integer
Dim choix_traitement_fumees As Integer

'variables échangées avec l'interface
Dim inlet_MES As Double
Public pd_EntreeBouesNonSechees_kgMSj As Double
Dim inlet_siccite As Double
Dim inlet_Q As Double
Dim inlet_BH As Double
Dim inlet_MV_MES As Double
Dim inlet_boues_temperature As Double
Dim md_BouesEntreesRatioSoufreSurMV As Double
Dim md_BouesEntreesRatioChloreSurMV As Double
Dim PCI_MV_kWh_kg As Double
Dim PCI_BH_kWh_kg As Double
Dim combustible_Qvol As Double
Dim energie_refroidissement_incinerateur_kWhj As Double
Dim debit_air_kgj As Double
Dim debit_air_Nm3j As Double
Dim T_air_inlet_°C As Double
Dim T_air_prechauffe_°C As Double   'à ajouter
Dim debit_fumees_kgj As Double
Dim debit_fumees_Nm3j As Double
Private md_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3 As Double
Private md_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3 As Double
Public md_TraitementSNCRConsommationNH3Pur_kgj As Double
Public md_TraitementSNCRConsommationUreePure_kgj As Double
Dim T_four_°C As Double
Dim T_fumees_inlet_electrofiltre_°C As Double
Dim rendement_electrofiltre As Double
Dim T_fumees_outlet_electrofiltre_°C As Double
Dim T_fumees_inlet_filtre_manches_°C As Double
Dim air_dilution_Nm3j As Double
Public siccite_cendres As Double
Public cendres_kgj As Double
Private md_CoefficientAutoneutralisationSoufreElectrofiltre As Double   'correspond au souffre piégé dans l'électrofiltre
Dim fumees_sortie_electrofiltre_mgSO2_Nm3 As Double
Dim fumees_sortie_electrofiltre_mgHCl_Nm3 As Double
Public siccite_REFIB As Double
Public REFIB_kgj As Double
Public NaOHpur_kgj As Double
Public NaHCO3pur_kgj As Double
Public charbon_actif_kgj As Double
Dim four_diametre As Double
Dim besoins_thermiques_air_kWhj As Double
Dim energie_recuperee_fumees_kWhj As Double
Dim energie_thermique_maintien_kWhj As Double
Dim energie_thermique_demarrage_kWhj As Double
Private md_ConsommationElectriqueSpecifique_kWh_tMS As Double
Public electricite As Double

'Boolean pour valeurs forcées
Dim mn_BouesEntreesRatioSoufreSurMV_force As Boolean
Dim mn_BouesEntreesRatioChloreSurMV_force As Boolean
Dim inlet_boues_temperature_force As Boolean
Dim PCI_MV_kWh_kg_force As Boolean
Dim PCI_BH_kWh_kg_force As Boolean
Dim combustible_Qvol_force As Boolean
Dim energie_refroidissement_incinerateur_kWhj_force As Boolean
Dim debit_air_Nm3j_force As Boolean
Dim T_air_inlet_°C_force As Boolean
Dim T_air_prechauffe_°C_force As Boolean   'à ajouter
Dim T_four_°C_force As Boolean
Dim debit_fumees_Nm3j_force As Boolean
Private mn_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3_force As Boolean
Private mn_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3_force As Boolean
Private mn_TraitementSNCRConsommationNH3Pur_kgj_force As Boolean
Private mn_TraitementSNCRConsommationUreePure_kgj_force As Boolean
Dim T_fumees_inlet_electrofiltre_°C_force As Boolean
Dim rendement_electrofiltre_force As Boolean
Dim siccite_cendres_force As Boolean
Private mn_CoefficientAutoneutralisationSoufreElectrofiltre_force As Boolean
Dim fumees_sortie_electrofiltre_mgSO2_Nm3_force As Boolean
Dim fumees_sortie_electrofiltre_mgHCl_Nm3_force As Boolean
Dim cendres_kgj_force As Boolean
Dim siccite_REFIB_force As Boolean
Dim REFIB_kgj_force As Boolean
Dim NaOHpur_kgj_force As Boolean
Dim NaHCO3pur_kgj_force As Boolean
Dim charbon_actif_kgj_force As Boolean
Dim T_fumees_outlet_electrofiltre_°C_force As Boolean
Dim T_fumees_inlet_filtre_manches_°C_force As Boolean
Dim air_dilution_Nm3j_force As Boolean
Dim four_diametre_force As Boolean
Dim besoins_thermiques_air_kWhj_force As Boolean
Dim energie_recuperee_fumees_kWhj_force As Boolean
Dim energie_thermique_maintien_kWhj_force As Boolean
Dim energie_thermique_demarrage_kWhj_force As Boolean
Private mn_ConsommationElectriqueSpecifique_kWh_tMS_force As Boolean

'valeurs forcées
Dim md_BouesEntreesRatioSoufreSurMV_f As Double
Dim md_BouesEntreesRatioChloreSurMV_f As Double
Dim inlet_boues_temperature_f As Double
Dim PCI_MV_kWh_kg_f As Double
Dim PCI_BH_kWh_kg_f As Double
Dim combustible_Qvol_f As Double
Dim energie_refroidissement_incinerateur_kWhj_f As Double
Dim debit_air_Nm3j_f As Double
Dim T_air_inlet_°C_f As Double
Dim T_air_prechauffe_°C_f As Double   'à ajouter
Dim T_four_°C_f As Double
Dim debit_fumees_Nm3j_f As Double
Private md_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3_f As Double
Private md_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3_f As Double
Private md_TraitementSNCRConsommationNH3Pur_kgj_f As Double
Private md_TraitementSNCRConsommationUreePure_kgj_f As Double
Dim T_fumees_inlet_electrofiltre_°C_f As Double
Dim rendement_electrofiltre_f As Double
Dim siccite_cendres_f As Double
Dim cendres_kgj_f As Double
Private md_CoefficientAutoneutralisationSoufreElectrofiltre_f As Double
Dim fumees_sortie_electrofiltre_mgSO2_Nm3_f As Double
Dim fumees_sortie_electrofiltre_mgHCl_Nm3_f As Double
Dim siccite_REFIB_f As Double
Dim REFIB_kgj_f As Double
Dim NaOHpur_kgj_f As Double
Dim NaHCO3pur_kgj_f As Double
Dim charbon_actif_kgj_f As Double
Dim T_fumees_outlet_electrofiltre_°C_f As Double
Dim T_fumees_inlet_filtre_manches_°C_f As Double
Dim air_dilution_Nm3j_f As Double
Dim four_diametre_f As Double
Dim besoins_thermiques_air_kWhj_f As Double
Dim energie_recuperee_fumees_kWhj_f As Double
Dim energie_thermique_maintien_kWhj_f As Double
Dim energie_thermique_demarrage_kWhj_f As Double
Private md_ConsommationElectriqueSpecifique_kWh_tMS_f As Double

'hypothèses
Dim PCI_BH_autocombustibilite_kWh_kg As Double
Dim PCI_fioul As Double
Dim PCI_CH4 As Double
Dim PCI_gaz_nat As Double
Dim R_gaz_parfait As Double
Dim T_normale_K As Double
Dim Patm_Pa As Double
Dim MM_CH4_kg_mol As Double
Dim MM_CO2_kg_mol As Double
Dim air_inlet_kgair_kgMV_boues_non_digerees As Double
Dim air_inlet_kgair_kgMV_boues_digerees As Double
Dim air_inlet_kgair_kgfioul As Double
Dim air_inlet_kgair_kgGazNat As Double
Dim air_inlet_kgair_kgbiogaz As Double
Dim deltaH_vap_kWh_kg As Double
Dim masse_volumique_boues As Double
Dim densite_fioul As Double
Dim densite_gaz_nat As Double
Dim densite_biogaz As Double
Dim Cp_air_kWh_kg_K As Double
Dim Cp_eau_kWh_kg_K As Double
Dim Cp_MS_kWh_kg_K As Double
Dim Cp_boues_kWh_kg_K As Double
Dim Cp_fumees_kWh_kg_K As Double
Dim vitesse_fumees_ms As Double
Dim MM_fumees_kg_mol As Double
Dim MM_air_kg_mol As Double
Dim efficacite_prechauffeur As Double
Dim T_max_filtre_manches As Double
'Dim T_min_fumees_corrosion As Double
Dim T_air_dilution_°C As Double
Dim sechage_siccite_maxi As Double
Dim MM_O2_kg_mol As Double
Dim MM_SO2_kg_mol As Double
Dim NaHCO3_dose_g_gSO2 As Double
Dim NaHCO3_dose_g_gHCl As Double
Dim charbon_actif_dose_kg_m3fumees As Double
Dim NaOH_dose_g_gSO2 As Double
Dim NaOH_dose_g_gHCl As Double
Private md_HypotheseRatioSoufreSurMVBoues(1 To nb_type_boues) As Double
Private md_HypotheseRatioChloreSurMVBoues(1 To nb_type_boues) As Double
Dim NaHCO3_exces As Double
Dim NaOH_exces As Double
Dim MM_Na2SO4_kg_mol As Double
Dim MM_HCl_kg_mol As Double
Dim MM_NaCl_kg_mol As Double
Dim MM_NaHCO3_kg_mol As Double
Dim ratio_pertes_thermiques As Double
Private md_MasseMolaireNO2_kg_mol As Double
Private md_MasseMolaireUreeNH2CONH2_kg_mol As Double
Private md_MasseMolaireAmmoniaqueNH3_kg_mol As Double
Private md_ReactionSNCRStoechiometrieUree_molNH2CONH2_molNO2 As Double
Private md_ReactionSNCRStoechiometrieAmmoniaque_molNH3_molNO2 As Double
Private md_ReactionSNCRExcesReactif As Double


Dim remarque(2) As String




Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'Vitesse des fumées pour design du four (m/s à T,P réels)
    vitesse_fumees_ms = 0.85

    'Cp air, fumées, boues
    Cp_air_kWh_kg_K = 1.063 / pd_NOMBRE_SECONDE_PAR_HEURE
    Cp_fumees_kWh_kg_K = 1.39 / pd_NOMBRE_SECONDE_PAR_HEURE
    Cp_eau_kWh_kg_K = 4.18 / pd_NOMBRE_SECONDE_PAR_HEURE
    Cp_MS_kWh_kg_K = 0.285 * 4.18 / pd_NOMBRE_SECONDE_PAR_HEURE    'R. LE GUILLY

    'calcul de conditions normales pour le biogaz et les fumées
    Patm_Pa = 101325
    T_normale_K = 273.15
    R_gaz_parfait = 8.314

    'hypothèse masse volumique boues
    masse_volumique_boues = 1000

    'densité des combustibles externes
    densite_fioul = 0.84
    densite_gaz_nat = 0.82    'R. LE GUILLY

    'PCI des combustibles
    PCI_fioul = pd_PCI_FIOUL_kWh_L    'PCI du fioul kWh/ L
    PCI_gaz_nat = pd_PCI_GAZ_NATUREL_kWh_Nm3    'PCI du gaz naturel KWh/ Nm3
    PCI_CH4 = pd_PCI_CH4_kWh_Nm3  'pour le biogaz

    'Ratio de pertes thermiques
    ratio_pertes_thermiques = 0.05   'vu avec R. LE GUILLY

    'Enthalipie de vaporisation de l'eau
    deltaH_vap_kWh_kg = 598 * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE

    'débit d'air pour le combustible
    air_inlet_kgair_kgGazNat = 19.2
    air_inlet_kgair_kgfioul = 21.2
    air_inlet_kgair_kgbiogaz = 8.5      'avant 21.7 mais ne correspond pas à 5% O2 en sortie    21.7   'pour 5% O2 sur sec en sortie avec biogaz à 62% CH4

    'Débit d'air pour combustion des boues (ratio en kg/kgMV)
    air_inlet_kgair_kgMV_boues_digerees = 10.9    'Debit d’air boues mixtes ou biologiques kg air/ kg MV
    air_inlet_kgair_kgMV_boues_non_digerees = 11.4    'Debit d’air boues digerees kg air/ kg MV

    'masse molaire des fumées en attendant un calcul plus détaillé
    MM_fumees_kg_mol = 0.025  'kg/mol

    'efficacité du préchauffeur d'air
    efficacite_prechauffeur = 0.95    'Efficacite prechauffeur d’air

    'températures limites pour les fumées
    T_max_filtre_manches = 220  'température maximale des fumées pour admission sur des filtres
    'T_min_fumees_corrosion = 200    'température minimale des fumées (au dessous, risque de corrosion)

    'température de l'air de dilution
    T_air_dilution_°C = 15

    'siccité maximale des boues séchées (pour ajuster la siccité en amont
    sechage_siccite_maxi = 0.9

    'masses molaires
    MM_CH4_kg_mol = pd_MASSE_MOLAIRE_C_kg_mol + 4 * pd_MASSE_MOLAIRE_H_kg_mol
    MM_CO2_kg_mol = pd_MASSE_MOLAIRE_C_kg_mol + 2 * pd_MASSE_MOLAIRE_O_kg_mol
    MM_air_kg_mol = 0.21 * (2 * pd_MASSE_MOLAIRE_O_kg_mol) + 0.79 * (2 * pd_MASSE_MOLAIRE_N_kg_mol)
    MM_O2_kg_mol = 2 * pd_MASSE_MOLAIRE_O_kg_mol
    MM_SO2_kg_mol = pd_MASSE_MOLAIRE_S_kg_mol + 2 * pd_MASSE_MOLAIRE_O_kg_mol
    MM_Na2SO4_kg_mol = 2 * pd_MASSE_MOLAIRE_Na_kg_mol + pd_MASSE_MOLAIRE_S_kg_mol + 4 * pd_MASSE_MOLAIRE_O_kg_mol
    MM_HCl_kg_mol = pd_MASSE_MOLAIRE_H_kg_mol + pd_MASSE_MOLAIRE_Cl_kg_mol
    MM_NaCl_kg_mol = pd_MASSE_MOLAIRE_Na_kg_mol + pd_MASSE_MOLAIRE_Cl_kg_mol
    MM_NaHCO3_kg_mol = pd_MASSE_MOLAIRE_Na_kg_mol + pd_MASSE_MOLAIRE_H_kg_mol + pd_MASSE_MOLAIRE_C_kg_mol + 3 * pd_MASSE_MOLAIRE_O_kg_mol
    
    'Réactifs
    NaHCO3_exces = 0.25    '25% excès   Romuald LE GUILLY 09/2012
    NaOH_exces = 0.1     '10% excès   Romuald LE GUILLY 09/2012
    If choix_traitement_fumees = fumees_filiere_seche Then
        NaHCO3_dose_g_gSO2 = 2.625    '2.625 g/g stoechiométrique + 25% excès   Romuald LE GUILLY 09/2012
        NaHCO3_dose_g_gHCl = 2.301       '2.301 g/g stoechiométrique + 25% excès   Romuald LE GUILLY 09/2012
        charbon_actif_dose_kg_m3fumees = 0.0001    '100 mg/m3   Romuald LE GUILLY 09/2012
        NaOH_dose_g_gSO2 = 0
        NaOH_dose_g_gHCl = 0
    Else
        NaHCO3_dose_g_gSO2 = 0
        NaHCO3_dose_g_gHCl = 0
        charbon_actif_dose_kg_m3fumees = 0
        NaOH_dose_g_gSO2 = 1.25   '1.25 g/g stoechiométrique + 10% excès   Romuald LE GUILLY 09/2012
        NaOH_dose_g_gHCl = 1.1   '1.1 g/g stoechiométrique + 10% excès   Romuald LE GUILLY 09/2012
    End If
    
    'Réactifs traitement SNCR four contre les NOx
    md_MasseMolaireNO2_kg_mol = pd_MASSE_MOLAIRE_N_kg_mol + 2 * pd_MASSE_MOLAIRE_O_kg_mol
    md_MasseMolaireUreeNH2CONH2_kg_mol = 2 * pd_MASSE_MOLAIRE_N_kg_mol + 4 * pd_MASSE_MOLAIRE_H_kg_mol + pd_MASSE_MOLAIRE_C_kg_mol + pd_MASSE_MOLAIRE_O_kg_mol
    md_MasseMolaireAmmoniaqueNH3_kg_mol = pd_MASSE_MOLAIRE_N_kg_mol + 3 * pd_MASSE_MOLAIRE_H_kg_mol
    md_ReactionSNCRStoechiometrieUree_molNH2CONH2_molNO2 = 1 / 2     'Réaction: 2 NO2 + CO(NH2)2 --> 2 N2 + CO2 + 2 H2O + 0.5 O2
    md_ReactionSNCRStoechiometrieAmmoniaque_molNH3_molNO2 = 1     'Réaction: NO2 + NH3 --> N2 + 1.5 H2O + 0.25 O2
    md_ReactionSNCRExcesReactif = 2 'On injecte 2 fois la quantité stoechiométrique en excès --> au final on met 3 fois (1+2) la quantité de réactifs stoechiométrique  source R. LE GUILLY et J. CHAUZY, cf tableur de calcul R. LE GUILLY
    
    'Hypothèses de teneur en soufre des boues à incinérer par type de boues / source: R. LE GUILLY et J. CHAUZY 10/2013
    'on suppose que s'il y a une digestion, toutes les boues seront digérées
    If pi_ChoixDigestion = True Then
        md_HypotheseRatioSoufreSurMVBoues(boues_I) = 0.022
        md_HypotheseRatioSoufreSurMVBoues(boues_II) = 0.022
        md_HypotheseRatioSoufreSurMVBoues(boues_III) = 0.022
        md_HypotheseRatioSoufreSurMVBoues(boues_externes_1) = 0.022
        md_HypotheseRatioSoufreSurMVBoues(boues_externes_2) = 0.022
        md_HypotheseRatioSoufreSurMVBoues(graisses) = 0.022
    Else
        md_HypotheseRatioSoufreSurMVBoues(boues_I) = 0.015
        md_HypotheseRatioSoufreSurMVBoues(boues_II) = 0.013
        md_HypotheseRatioSoufreSurMVBoues(boues_III) = 0.013
        md_HypotheseRatioSoufreSurMVBoues(boues_externes_1) = 0.013
        md_HypotheseRatioSoufreSurMVBoues(boues_externes_2) = 0.013
        md_HypotheseRatioSoufreSurMVBoues(graisses) = 0.015
    End If

    'Hypothèses de teneur en chlore des boues à incinérer par type de boues / source: R. LE GUILLY et J. CHAUZY 10/2013
    'on suppose que s'il y a une digestion, toutes les boues seront digérées
    If pi_ChoixDigestion = True Then
        md_HypotheseRatioChloreSurMVBoues(boues_I) = 0.02
        md_HypotheseRatioChloreSurMVBoues(boues_II) = 0.02
        md_HypotheseRatioChloreSurMVBoues(boues_III) = 0.02
        md_HypotheseRatioChloreSurMVBoues(boues_externes_1) = 0.02
        md_HypotheseRatioChloreSurMVBoues(boues_externes_2) = 0.02
        md_HypotheseRatioChloreSurMVBoues(graisses) = 0
    Else
        md_HypotheseRatioChloreSurMVBoues(boues_I) = 0.01
        md_HypotheseRatioChloreSurMVBoues(boues_II) = 0.01
        md_HypotheseRatioChloreSurMVBoues(boues_III) = 0.01
        md_HypotheseRatioChloreSurMVBoues(boues_externes_1) = 0.01
        md_HypotheseRatioChloreSurMVBoues(boues_externes_2) = 0.01
        md_HypotheseRatioChloreSurMVBoues(graisses) = 0
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "46.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'température des boues en entrée
    If inlet_boues_temperature_force = True Then
        inlet_boues_temperature = inlet_boues_temperature_f
    Else
        inlet_boues_temperature = pd_TemperatureEauConditionsReelles_°C   'à adapter suivant traitement amont
    End If

    'température de l'air en entrée après soufflante
    If T_air_inlet_°C_force = True Then
        T_air_inlet_°C = T_air_inlet_°C_f
    Else
        T_air_inlet_°C = 60
    End If

    'température de l'air préchauffé (entrée boîte à vent)
    If T_air_prechauffe_°C_force = True Then
        T_air_prechauffe_°C = T_air_prechauffe_°C_f
    Else
        T_air_prechauffe_°C = 550
    End If

    'température du four
    If T_four_°C_force = True Then
        T_four_°C = T_four_°C_f
    Else
        T_four_°C = 890
    End If

    'PCI de la combustion des MV
    If PCI_MV_kWh_kg_force = True Then
        PCI_MV_kWh_kg = PCI_MV_kWh_kg_f
    Else
        PCI_MV_kWh_kg = 5500 * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE   '5412 / 861    'valeur Julien CHAUZY (voir s'il faut faire varier suivant le type de boues)
    End If

    'Concentration en NOx des fumées si pas de traitement SNCR (Selective Non Catalytic Reaction) dans le four
    If mn_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3_force = True Then
        md_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3 = md_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3_f
    Else
        md_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3 = 400    '400 mgNO2eq/Nm3fumées source: R. LE GUILLY et J. CHAUZY  valeur maximale autorisée en moyenne 15 min
    End If
    
    'Concentration en NOx des fumées avec traitement SNCR (Selective Non Catalytic Reaction) dans le four
    If mn_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3_force = True Then
        md_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3 = md_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3_f
    Else
        md_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3 = 200    '200 mgNO2eq/Nm3fumées source: R. LE GUILLY et J. CHAUZY  valeur maximale autorisée en moyenne journalière
    End If
    
    'température en entrée de l'électrofiltre
    If T_fumees_inlet_electrofiltre_°C_force = True Then
        T_fumees_inlet_electrofiltre_°C = T_fumees_inlet_electrofiltre_°C_f
    Else
        T_fumees_inlet_electrofiltre_°C = 275    'entre 250°C et 300°C  source: Romuald LE GUILLY
    End If

    If choix_traitement_fumees = fumees_filiere_seche Then
        'température en sortie de l'électrofiltre
        If T_fumees_outlet_electrofiltre_°C_force = True Then
            T_fumees_outlet_electrofiltre_°C = T_fumees_outlet_electrofiltre_°C_f
        Else
            T_fumees_outlet_electrofiltre_°C = T_fumees_inlet_electrofiltre_°C - 25  'on perd entre 20°C et 30°C dans l'électrofiltre   source Romuald LE GUILLY
        End If
        'température en entrée du filtre à manches
        If T_fumees_inlet_filtre_manches_°C_force = True Then
            T_fumees_inlet_filtre_manches_°C = T_fumees_inlet_filtre_manches_°C_f
        Else
            If T_fumees_outlet_electrofiltre_°C < T_max_filtre_manches Then
                T_fumees_inlet_filtre_manches_°C = T_fumees_outlet_electrofiltre_°C
            Else
                T_fumees_inlet_filtre_manches_°C = T_max_filtre_manches
            End If
        End If
    End If

    'rendement de l'électrofiltre
    If rendement_electrofiltre_force = True Then
        rendement_electrofiltre = rendement_electrofiltre_f
    Else
        rendement_electrofiltre = 0.97    'source Romuald LE GUILLY
    End If

    'Coefficient d'autoneutralité du soufre (rendement d'élimination du soufre dans l'électrofiltre)
    If mn_CoefficientAutoneutralisationSoufreElectrofiltre_force = True Then
        md_CoefficientAutoneutralisationSoufreElectrofiltre = md_CoefficientAutoneutralisationSoufreElectrofiltre_f
    Else
        md_CoefficientAutoneutralisationSoufreElectrofiltre = 0.2  '20% d'après R. LE GUILLY et J. CHAUZY
    End If
    
    'siccité des cendres
    If siccite_cendres_force = True Then
        siccite_cendres = siccite_cendres_f
    Else
        siccite_cendres = 0.995    'source Romuald LE GUILLY: 99,5%, elles peuvent être humidifiées à 20-30% suivant la destination
    End If

    'siccité des REFIB
    If siccite_REFIB_force = True Then
        siccite_REFIB = siccite_REFIB_f
    Else
        siccite_REFIB = 0.98    'source Romuald LE GUILLY
    End If

    'Consommation électrique spécifique en kWh/tMS
    If mn_ConsommationElectriqueSpecifique_kWh_tMS_force = True Then
        md_ConsommationElectriqueSpecifique_kWh_tMS = md_ConsommationElectriqueSpecifique_kWh_tMS_f
    Else
        md_ConsommationElectriqueSpecifique_kWh_tMS = 350        'vu avec R. LE GUILLY et J. CHAUZY en attendant d'avoir des données détaillées 10/2013
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "46.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub




Sub dimensionnement(ByRef boues_flux, ByRef boues_pollution_soluble, ByRef besoins_th, ByRef disponibilites_thermiques, ByVal biogaz_compo_CH4 As Double, ByRef DCO_oxydation)


    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Q_retour As Double

    Dim inlet_MES_detail(2) As Double
    Dim inlet_Q_detail(2) As Double

    Dim intermediaire_MES As Double
    Dim intermediaire_Q As Double
    Dim intermediaire_NK As Double
    Dim intermediaire_Pt As Double
    Dim intermediaire_DCO As Double
    Dim intermediaire_DBO As Double
    Dim inlet_DCO As Double
    Dim outlet_DCO As Double
    Dim inlet_DBO As Double
    Dim outlet_DBO As Double
    Dim inlet_NK As Double
    Dim outlet_NK As Double
    Dim inlet_Pt As Double
    Dim outlet_Pt As Double
    Dim ratio_DCO As Double
    Dim intermediaire_MV As Double
    Dim intermediaire_reduction_MV As Double
    Dim inlet_MV As Double
    Dim inlet_MV_pci As Double
    Dim outlet_MV As Double
    Dim outlet_concentration As Double
    Dim soluble_kg(nb_repere_mgL) As Double
    Dim condensat_direct_eau_inlet_kgj As Double
    Dim EntreeBouesSoufre_kgj As Double
    Dim EntreeBouesChlore_kgj As Double
    
    Dim ratio_energie_fumeesETpertes_fumees As Double

    Dim energie_appoint_kWhj As Double   'à ajouter
    Dim combustible_kgj As Double
    Dim densite_biogaz As Double

    Dim siccite_autocombustibilite As Double

    Dim NaHCO3pur_kgj_traitSO2 As Double
    Dim NaHCO3pur_kgj_traitHCl As Double
    Dim NaHCO3pur_kgj_exces As Double

    Dim NaOHpur_kgj_traitSO2 As Double
    Dim NaOHpur_kgj_traitHCl As Double
    Dim NaOHpur_kgj_exces As Double

    'gestion des ratios de MV reduits par type de boues
    Dim MV_reduits(nb_type_boues) As Double
    Dim delta_MV_reduits As Double

    inlet_MES = 0
    inlet_MV = 0
    inlet_Q = 0
    inlet_DCO = 0
    inlet_DBO = 0
    inlet_NK = 0
    inlet_Pt = 0
    EntreeBouesSoufre_kgj = 0
    EntreeBouesChlore_kgj = 0
    'inlet_MES_a = 0
    'inlet_Q_a = 0
    'inlet_MES_b = 0
    'inlet_Q_b = 0

    For I = 1 To 2
        inlet_MES_detail(I) = 0
        inlet_Q_detail(I) = 0
    Next I

    For I = 1 To nb_inlet
        If choix_inlet(I) <> inlet_aucun Then
            If inlet_etape_anterieure(I) = boues_inlet Then
                'variables d'entree
                intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MES)
                If intermediaire_MES <> 0 Then
                    intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_Q)
                    intermediaire_MV = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MV_MES) * intermediaire_MES
                    intermediaire_DCO = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DCO_MES) * intermediaire_MES
                    intermediaire_DBO = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DBO_MES) * intermediaire_MES
                    intermediaire_NK = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_NK_MES) * intermediaire_MES
                    intermediaire_Pt = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_Pt_MES) * intermediaire_MES

                    inlet_MES_detail(I) = intermediaire_MES
                    inlet_Q_detail(I) = intermediaire_Q
                    
                    'ENTREE SOLUBLE
                    'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                    For k = 1 To nb_repere_mgL
                        soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), k) * intermediaire_Q / 1000
                        boues_pollution_soluble(boues_incinerees, inlet_type_boues(I), k) = 0
                    Next k

                    inlet_MES = inlet_MES + intermediaire_MES
                    inlet_Q = inlet_Q + intermediaire_Q
                    inlet_DCO = inlet_DCO + intermediaire_DCO + soluble_kg(repere_mgL_DCO)
                    inlet_DBO = inlet_DBO + intermediaire_DBO
                    inlet_NK = inlet_NK + intermediaire_NK + soluble_kg(repere_mgL_NK)
                    inlet_Pt = inlet_Pt + intermediaire_Pt + soluble_kg(repere_mgL_Pt)
                    inlet_MV = inlet_MV + intermediaire_MV
                    EntreeBouesSoufre_kgj = EntreeBouesSoufre_kgj + md_HypotheseRatioSoufreSurMVBoues(inlet_type_boues(I)) * intermediaire_MV
                    EntreeBouesChlore_kgj = EntreeBouesChlore_kgj + md_HypotheseRatioChloreSurMVBoues(inlet_type_boues(I)) * intermediaire_MV

                    boues_flux(boues_incinerees, inlet_type_boues(I), repere_flux_in) = boues_flux(boues_incinerees, inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_incinerees, inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_incinerees, inlet_type_boues(I), repere_flux_in)

                    'je retranche la part de flux d'entree utilisee (objectif: verif des bilans matières)
                    boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                End If
            Else
                For j = 1 To nb_type_boues
                    'variables d'entree
                    intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_MES)
                    If intermediaire_MES <> 0 Then
                        intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_Q)
                        intermediaire_MV = boues_flux(inlet_etape_anterieure(I), j, repere_MV_MES) * intermediaire_MES
                        intermediaire_DCO = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DCO_MES) * intermediaire_MES
                        intermediaire_DBO = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DBO_MES) * intermediaire_MES
                        intermediaire_NK = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_NK_MES) * intermediaire_MES
                        intermediaire_Pt = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_Pt_MES) * intermediaire_MES

                        inlet_MES_detail(I) = inlet_MES_detail(I) + intermediaire_MES
                        inlet_Q_detail(I) = inlet_Q_detail(I) + intermediaire_Q

                        'ENTREE SOLUBLE
                        'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                        For k = 1 To nb_repere_mgL
                            soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), j, k) * intermediaire_Q / 1000
                            boues_pollution_soluble(boues_incinerees, j, k) = 0
                        Next k

                        inlet_MES = inlet_MES + intermediaire_MES
                        inlet_Q = inlet_Q + intermediaire_Q
                        inlet_DCO = inlet_DCO + intermediaire_DCO + soluble_kg(repere_mgL_DCO)
                        inlet_DBO = inlet_DBO + intermediaire_DBO
                        inlet_NK = inlet_NK + intermediaire_NK + soluble_kg(repere_mgL_NK)
                        inlet_Pt = inlet_Pt + intermediaire_Pt + soluble_kg(repere_mgL_Pt)
                        inlet_MV = inlet_MV + intermediaire_MV
                        EntreeBouesSoufre_kgj = EntreeBouesSoufre_kgj + md_HypotheseRatioSoufreSurMVBoues(j) * intermediaire_MV
                        EntreeBouesChlore_kgj = EntreeBouesChlore_kgj + md_HypotheseRatioChloreSurMVBoues(j) * intermediaire_MV

                        boues_flux(boues_incinerees, j, repere_flux_in) = boues_flux(boues_incinerees, j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                        boues_flux(boues_incinerees, j, repere_verif_flux) = boues_flux(boues_incinerees, j, repere_flux_in)

                        'je retranche la part de flux d'entree utilisee (objectif: verif des bilans matières)
                        boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                    End If
                Next j
            End If
        End If
    Next I

    If inlet_MES = 0 Or inlet_Q = 0 Then
        MsgBox Feuil5.Range("erreur_choix_incineration_inlet") & " ", Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If

    If mn_BouesEntreesRatioSoufreSurMV_force = True Then
        md_BouesEntreesRatioSoufreSurMV = md_BouesEntreesRatioSoufreSurMV_f
    Else
        md_BouesEntreesRatioSoufreSurMV = EntreeBouesSoufre_kgj / inlet_MV
    End If
    If mn_BouesEntreesRatioChloreSurMV_force = True Then
        md_BouesEntreesRatioChloreSurMV = md_BouesEntreesRatioChloreSurMV_f
    Else
        md_BouesEntreesRatioChloreSurMV = EntreeBouesChlore_kgj / inlet_MV
    End If
    
    DCO_oxydation = DCO_oxydation + inlet_DCO

    '*********************** BILAN MASSE ***********************************************************************************

    'On récupère les données sur les boues d'entrée
    inlet_siccite = inlet_MES / (inlet_Q * masse_volumique_boues)
    inlet_BH = inlet_MES / inlet_siccite
    inlet_MV_MES = inlet_MV / inlet_MES

    'Calcul du PCI de la boue
    PCI_BH_kWh_kg = inlet_siccite * inlet_MV_MES * PCI_MV_kWh_kg - deltaH_vap_kWh_kg * (1 - inlet_siccite)
    Cp_boues_kWh_kg_K = inlet_siccite * Cp_MS_kWh_kg_K + (1 - inlet_siccite) * Cp_eau_kWh_kg_K

    'Calcul du PCI d'autocombustibilité (avec la siccité actuelle) et de la siccité d'autocombustibilité (MS actuels reçus)
    ratio_energie_fumeesETpertes_fumees = 1 / (1 - ratio_pertes_thermiques)
    If pi_ChoixDigestion = True Then
        PCI_BH_autocombustibilite_kWh_kg = (ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C - Cp_boues_kWh_kg_K * inlet_boues_temperature) + air_inlet_kgair_kgMV_boues_digerees * inlet_MV_MES * inlet_siccite * (ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C - Cp_air_kWh_kg_K * T_air_prechauffe_°C)
        'PCI_BH_autocombustibilite_kWh_kg = pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE * (-0.46 * inlet_boues_temperature + 858.5)
        siccite_autocombustibilite = ((ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C - Cp_boues_kWh_kg_K * inlet_boues_temperature) + deltaH_vap_kWh_kg) / (PCI_MV_kWh_kg * inlet_MV_MES + deltaH_vap_kWh_kg + air_inlet_kgair_kgMV_boues_digerees * inlet_MV_MES * (Cp_air_kWh_kg_K * T_air_prechauffe_°C - ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C))
    Else
        PCI_BH_autocombustibilite_kWh_kg = (ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C - Cp_boues_kWh_kg_K * inlet_boues_temperature) + air_inlet_kgair_kgMV_boues_non_digerees * inlet_MV_MES * inlet_siccite * (ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C - Cp_air_kWh_kg_K * T_air_prechauffe_°C)
        'PCI_BH_autocombustibilite_kWh_kg = pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE * (-0.493 * inlet_boues_temperature + 893)
        siccite_autocombustibilite = ((ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C - Cp_boues_kWh_kg_K * inlet_boues_temperature) + deltaH_vap_kWh_kg) / (PCI_MV_kWh_kg * inlet_MV_MES + deltaH_vap_kWh_kg + air_inlet_kgair_kgMV_boues_non_digerees * inlet_MV_MES * (Cp_air_kWh_kg_K * T_air_prechauffe_°C - ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C))
    End If

    'On gère l'attribution de la siccité d'autocombustibilité
    If incineration_siccite_ajustable(1) = True Then
        If incineration_siccite_ajustable(2) = True Then
            For I = 1 To 2
                incineration_siccite_autocombustibilite(I) = siccite_autocombustibilite
            Next I
        Else
            If inlet_MES_detail(2) > 0 Then
                incineration_siccite_autocombustibilite(1) = inlet_MES_detail(1) / ((inlet_MES / (masse_volumique_boues * siccite_autocombustibilite) - inlet_Q_detail(2)) * masse_volumique_boues)
            Else
                incineration_siccite_autocombustibilite(1) = siccite_autocombustibilite
            End If
        End If
    ElseIf incineration_siccite_ajustable(2) = True Then
        If inlet_MES_detail(1) > 0 Then
            incineration_siccite_autocombustibilite(2) = inlet_MES_detail(2) / ((inlet_MES / (masse_volumique_boues * siccite_autocombustibilite) - inlet_Q_detail(1)) * masse_volumique_boues)
        Else
            incineration_siccite_autocombustibilite(2) = siccite_autocombustibilite
        End If
    End If
    For I = 1 To 2
        If incineration_siccite_autocombustibilite(I) > sechage_siccite_maxi Then
            incineration_siccite_autocombustibilite(I) = sechage_siccite_maxi
        End If
    Next I

    'On calcule l'ajout de combustible d'appoint ou le refroidissement
    If Abs(PCI_BH_kWh_kg - PCI_BH_autocombustibilite_kWh_kg) / Abs(PCI_BH_kWh_kg) > critere_convergence Then
        If PCI_BH_kWh_kg < PCI_BH_autocombustibilite_kWh_kg Then
            'énergie manquante
            energie_appoint_kWhj = (PCI_BH_autocombustibilite_kWh_kg - PCI_BH_kWh_kg) * inlet_BH
            energie_refroidissement_incinerateur_kWhj = 0
            'calcul du débit de combustible, on diminue le PCI par l'énergie qu'il faudra pour chauffer l'air de combustion dû au combustible
            If choix_combustible = combustible_repere_gaz_naturel Then
                combustible_Qvol = energie_appoint_kWhj / (PCI_gaz_nat - air_inlet_kgair_kgGazNat * densite_gaz_nat * Cp_air_kWh_kg_K * T_air_prechauffe_°C)
                combustible_kgj = densite_gaz_nat * combustible_Qvol
                debit_air_kgj = air_inlet_kgair_kgGazNat * combustible_kgj
                besoins_th(energie_combustible) = besoins_th(energie_combustible) + combustible_Qvol * PCI_gaz_nat
            ElseIf choix_combustible = combustible_repere_fioul Then
                combustible_Qvol = energie_appoint_kWhj / (PCI_fioul - air_inlet_kgair_kgfioul * densite_fioul * Cp_air_kWh_kg_K * T_air_prechauffe_°C)
                combustible_kgj = densite_fioul * combustible_Qvol
                debit_air_kgj = air_inlet_kgair_kgfioul * combustible_kgj
                besoins_th(energie_combustible) = besoins_th(energie_combustible) + combustible_Qvol * PCI_fioul
            Else   'biogaz
                densite_biogaz = Patm_Pa / (R_gaz_parfait * T_normale_K) * (biogaz_compo_CH4 * MM_CH4_kg_mol + (1 - biogaz_compo_CH4) * MM_CO2_kg_mol)
                combustible_Qvol = energie_appoint_kWhj / (PCI_CH4 * biogaz_compo_CH4 - air_inlet_kgair_kgbiogaz * densite_biogaz * Cp_air_kWh_kg_K * T_air_prechauffe_°C)
                combustible_kgj = densite_biogaz * combustible_Qvol
                debit_air_kgj = air_inlet_kgair_kgbiogaz * combustible_kgj
                besoins_th(energie_biogaz) = besoins_th(energie_biogaz) + combustible_Qvol * PCI_CH4 * biogaz_compo_CH4
            End If
        Else
            'surplus d'énergie
            energie_appoint_kWhj = 0
            energie_refroidissement_incinerateur_kWhj = (PCI_BH_kWh_kg - PCI_BH_autocombustibilite_kWh_kg) * inlet_BH
            combustible_Qvol = 0
            combustible_kgj = 0
            debit_air_kgj = 0
        End If
    Else
        energie_appoint_kWhj = 0
        energie_refroidissement_incinerateur_kWhj = 0
        combustible_Qvol = 0
        combustible_kgj = 0
        debit_air_kgj = 0
    End If

    'On calcule l'air ajouté pour la combustion des boues
    If pi_ChoixDigestion = True Then  'probabilité faible d'avoir une digestion et une incinération avec des boues différentes
        debit_air_kgj = debit_air_kgj + (air_inlet_kgair_kgMV_boues_non_digerees * inlet_MES * inlet_MV_MES)
    Else
        debit_air_kgj = debit_air_kgj + (air_inlet_kgair_kgMV_boues_digerees * inlet_MES * inlet_MV_MES)
    End If
    debit_air_Nm3j = debit_air_kgj * R_gaz_parfait * T_normale_K / (MM_air_kg_mol * Patm_Pa)


    'On calcule les fumées
    debit_fumees_kgj = debit_air_kgj + combustible_kgj + inlet_BH
    debit_fumees_Nm3j = debit_fumees_kgj * R_gaz_parfait * T_normale_K / (MM_fumees_kg_mol * Patm_Pa)
    
    'Calcul du traitement SNCR
    If mi_ChoixTraitementNOxFour <> mi_PAS_DE_TRAITEMENT_SNCR Then
        Call calcul_traitement_SNCR
    End If

    'dimensionnement du four
    If four_diametre_force = True Then
        four_diametre = four_diametre_f
    Else
        four_diametre = pd_NOMBRE_JOUR_PAR_SEMAINE / choix_nb_jours_travailles * (((debit_fumees_Nm3j / pd_NOMBRE_HEURE_PAR_JOUR / pd_NOMBRE_SECONDE_PAR_HEURE * (T_normale_K + T_four_°C) / T_normale_K) / vitesse_fumees_ms) * (4 / pd_NOMBRE_PI)) ^ (1 / 2)
    End If

    'On calcule les cendres à évacuer (récupérées sur l'électrofiltre) et les résidus REFIB (récupérés sur les filtres à manches pour la filière sèche ou sur la filière humide en quantité moindre)
    cendres_kgj = rendement_electrofiltre * inlet_MES * (1 - inlet_MV_MES) / siccite_cendres

    'On calcule la compo en soufre en en HCl des fumées
    fumees_sortie_electrofiltre_mgSO2_Nm3 = (md_BouesEntreesRatioSoufreSurMV * (MM_SO2_kg_mol / pd_MASSE_MOLAIRE_S_kg_mol) * inlet_MV * (1 - md_CoefficientAutoneutralisationSoufreElectrofiltre)) / debit_fumees_Nm3j * 1000000
    fumees_sortie_electrofiltre_mgHCl_Nm3 = (MM_HCl_kg_mol / pd_MASSE_MOLAIRE_Cl_kg_mol) * (md_BouesEntreesRatioChloreSurMV * inlet_MV) / debit_fumees_Nm3j * 1000000

    'On calcule le bicarbonate
    NaHCO3pur_kgj_traitSO2 = NaHCO3_dose_g_gSO2 * fumees_sortie_electrofiltre_mgSO2_Nm3 / 1000000 * debit_fumees_Nm3j
    NaHCO3pur_kgj_traitHCl = NaHCO3_dose_g_gHCl * fumees_sortie_electrofiltre_mgHCl_Nm3 / 1000000 * debit_fumees_Nm3j
    NaHCO3pur_kgj_exces = (NaHCO3pur_kgj_traitSO2 + NaHCO3pur_kgj_traitHCl) * (1 + NaHCO3_exces)
    NaHCO3pur_kgj = NaHCO3pur_kgj_traitSO2 + NaHCO3pur_kgj_traitHCl + NaHCO3pur_kgj_exces

    'On calcule la soude
    NaOHpur_kgj_traitSO2 = NaOH_dose_g_gSO2 * fumees_sortie_electrofiltre_mgSO2_Nm3 / 1000000 * debit_fumees_Nm3j
    NaOHpur_kgj_traitHCl = NaOH_dose_g_gHCl * fumees_sortie_electrofiltre_mgHCl_Nm3 / 1000000 * debit_fumees_Nm3j
    NaOHpur_kgj_exces = (NaOHpur_kgj_traitSO2 + NaOHpur_kgj_traitHCl) * (1 + NaOH_exces)
    NaOHpur_kgj = NaOHpur_kgj_traitSO2 + NaOHpur_kgj_traitHCl + NaOHpur_kgj_exces

    'On calcule le charbon actif
    charbon_actif_kgj = charbon_actif_dose_kg_m3fumees * debit_fumees_Nm3j * (T_normale_K + T_fumees_outlet_electrofiltre_°C) / T_normale_K

    'On calcule les REFIB
    REFIB_kgj = (1 - rendement_electrofiltre) * inlet_MES * (1 - inlet_MV_MES)   'résidus non captés avant l'électrofiltre
    REFIB_kgj = REFIB_kgj + NaHCO3pur_kgj_traitSO2 * MM_Na2SO4_kg_mol / (2 * MM_NaHCO3_kg_mol)  'Na2SO4 formé avec le SO2 et le bicarbonate ou la soude
    REFIB_kgj = REFIB_kgj + NaHCO3pur_kgj_traitHCl * MM_NaCl_kg_mol / MM_NaHCO3_kg_mol    'NaCl formé avec le HCl et le bicarbonate ou la soude
    REFIB_kgj = REFIB_kgj + NaHCO3pur_kgj_exces    'excès de bicarbonate
    REFIB_kgj = REFIB_kgj + charbon_actif_kgj
    REFIB_kgj = REFIB_kgj / siccite_REFIB

    '********************* BILAN THERMIQUE ***********************************************************
    Call bilan_thermique
    disponibilites_thermiques(energie_eau_HT + 1) = disponibilites_thermiques(energie_eau_HT + 1) + (energie_refroidissement_incinerateur_kWhj + debit_fumees_kgj * Cp_fumees_kWh_kg_K * (T_four_°C - T_fumees_inlet_electrofiltre_°C))
    'besoins_th(energie_eau_HT + 1) = besoins_th(energie_eau_HT + 1) + debit_air_kgj * Cp_air_kWh_kg_K * (T_air_prechauffe_°C - T_air_inlet_°C)
    If choix_combustible = combustible_repere_gaz_naturel Or choix_combustible = combustible_repere_fioul Then
        besoins_th(energie_combustible) = besoins_th(energie_combustible) + energie_thermique_maintien_kWhj + energie_thermique_demarrage_kWhj
    Else
        besoins_th(energie_biogaz) = besoins_th(energie_biogaz) + energie_thermique_maintien_kWhj + energie_thermique_demarrage_kWhj
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "46.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub

Sub calcul_consommation_electrique()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    electricite = md_ConsommationElectriqueSpecifique_kWh_tMS * inlet_MES / 1000



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "46.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub
Private Sub calcul_traitement_SNCR()


'******************************************************************************************************************************************************
'OBJECTIF
'   Cette procédure permet de calculer les consommations de réactifs pur (ammoniaque ou urée) utilisés dans le four lors du traitement SNCR
'   pour réduire la concentration en NOx des fumées afin de respecter les garanties de rejet.

'VARIABLES D'ENTREE:

'VARIABLES DE SORTIE:
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim NOxAEliminer_kgNO2eq_j As Double
Dim NOxAEliminer_molNO2eq_j As Double
Dim ConsommationReactifPur_molj As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "46.20"
NOxAEliminer_kgNO2eq_j = 0
NOxAEliminer_molNO2eq_j = 0
ConsommationReactifPur_molj = 0

'On calcule la quantité de NOx à éliminer
NOxAEliminer_kgNO2eq_j = debit_fumees_Nm3j * (md_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3 - md_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3) / 1000000
NOxAEliminer_molNO2eq_j = NOxAEliminer_kgNO2eq_j / md_MasseMolaireNO2_kg_mol

'On calcule les quantités de réactifs
If mi_ChoixTraitementNOxFour = mi_TRAITEMENT_SNCR_AVEC_AMMONIAQUE Then
    ConsommationReactifPur_molj = md_ReactionSNCRStoechiometrieAmmoniaque_molNH3_molNO2 * NOxAEliminer_molNO2eq_j * (1 + md_ReactionSNCRExcesReactif)
    md_TraitementSNCRConsommationNH3Pur_kgj = ConsommationReactifPur_molj * md_MasseMolaireAmmoniaqueNH3_kg_mol
ElseIf mi_ChoixTraitementNOxFour = mi_TRAITEMENT_SNCR_AVEC_UREE Then
    ConsommationReactifPur_molj = md_ReactionSNCRStoechiometrieUree_molNH2CONH2_molNO2 * NOxAEliminer_molNO2eq_j * (1 + md_ReactionSNCRExcesReactif)
    md_TraitementSNCRConsommationUreePure_kgj = ConsommationReactifPur_molj * md_MasseMolaireUreeNH2CONH2_kg_mol
End If


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True


End Sub

Sub bilan_thermique()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim T_fumees_out_prechauffeur_air_°C As Double
    Dim air_dilution_kgj As Double
    Dim capacite_incineration_kWhj As Double
    Dim conso_spec_maintien_pourcent As Double
    Dim conso_spec_demarrage_pourcent As Double


    'on calcule la température des fumées après récupération de l'énergie avec l'air
    If besoins_thermiques_air_kWhj_force = True Then
        besoins_thermiques_air_kWhj = besoins_thermiques_air_kWhj_f
    Else
        besoins_thermiques_air_kWhj = debit_air_kgj * Cp_air_kWh_kg_K * (T_air_prechauffe_°C - T_air_inlet_°C)
    End If
    T_fumees_out_prechauffeur_air_°C = T_four_°C - (besoins_thermiques_air_kWhj) / (debit_fumees_kgj * Cp_fumees_kWh_kg_K * efficacite_prechauffeur)

    'on calcule l'énergie valorisée (ou valorisable) sur les fumées
    If energie_recuperee_fumees_kWhj_force = True Then
        energie_recuperee_fumees_kWhj = energie_recuperee_fumees_kWhj_f
    Else
        energie_recuperee_fumees_kWhj = efficacite_prechauffeur * debit_fumees_kgj * Cp_fumees_kWh_kg_K * (T_four_°C - T_fumees_inlet_electrofiltre_°C)
    End If

    'on calcul la dilution à l'air avant de rentrer dans les filtres à manches
    If choix_traitement_fumees = fumees_filiere_seche Then
        'on calcule le débit d'air de dilution lorsque filière sèche et nécessaire
        If air_dilution_Nm3j_force = True Then
            air_dilution_Nm3j = air_dilution_Nm3j_f
        Else
            air_dilution_kgj = debit_fumees_kgj * (Cp_fumees_kWh_kg_K / Cp_air_kWh_kg_K) * ((T_fumees_inlet_filtre_manches_°C - T_fumees_outlet_electrofiltre_°C) / (T_air_dilution_°C - T_fumees_inlet_filtre_manches_°C))
            air_dilution_Nm3j = air_dilution_kgj * R_gaz_parfait * T_normale_K / (MM_air_kg_mol * Patm_Pa)
        End If
    End If


    'on calcule la consommation de maintien en température et au démarrage
    capacite_incineration_kWhj = inlet_MES * inlet_MV_MES * PCI_MV_kWh_kg
    'maintien en température
    If energie_thermique_maintien_kWhj_force = True Then
        energie_thermique_maintien_kWhj = energie_thermique_maintien_kWhj_f
    Else
        conso_spec_maintien_pourcent = 2.4127 * (capacite_incineration_kWhj * pd_NOMBRE_JOUR_PAR_AN) ^ (-0.0468)  'Consommation spécifique maintien %
        If conso_spec_maintien_pourcent < 1.01 Then
            energie_thermique_maintien_kWhj = 1248408    'kWh/an pour un fonctionnement 5j/7j
        ElseIf conso_spec_maintien_pourcent < 1.017 Then
            energie_thermique_maintien_kWhj = 1248408    'kWh/an pour un fonctionnement 5j/7j
        ElseIf conso_spec_maintien_pourcent < 1.025 Then
            energie_thermique_maintien_kWhj = 1081038    'kWh/an pour un fonctionnement 5j/7j
        ElseIf conso_spec_maintien_pourcent < 1.035 Then
            energie_thermique_maintien_kWhj = 925710    'kWh/an pour un fonctionnement 5j/7j
        ElseIf conso_spec_maintien_pourcent < 1.046 Then
            energie_thermique_maintien_kWhj = 755210    'kWh/an pour un fonctionnement 5j/7j
        ElseIf conso_spec_maintien_pourcent < 1.062 Then
            energie_thermique_maintien_kWhj = 602048    'kWh/an pour un fonctionnement 5j/7j
        ElseIf conso_spec_maintien_pourcent < 1.081 Then
            energie_thermique_maintien_kWhj = 445275    'kWh/an pour un fonctionnement 5j/7j
        ElseIf conso_spec_maintien_pourcent < 1.108 Then
            energie_thermique_maintien_kWhj = 312102    'kWh/an pour un fonctionnement 5j/7j
        Else
            energie_thermique_maintien_kWhj = 188802    'kWh/an pour un fonctionnement 5j/7j
        End If
        energie_thermique_maintien_kWhj = energie_thermique_maintien_kWhj * (pd_NOMBRE_JOUR_PAR_SEMAINE - choix_nb_jours_travailles) / 2    'car hypothèse EOLIA 2jours de maintien / semaine
        energie_thermique_maintien_kWhj = energie_thermique_maintien_kWhj / pd_NOMBRE_JOUR_PAR_AN   '(on ramène la consommation comme si les jours de maintien étaient répartis équitablement sur toute la semaine)
    End If
    'démarrage du four (on suppose un arrêt annuel pour maintenance entre 10 et 15 jours)
    If energie_thermique_demarrage_kWhj_force = True Then
        energie_thermique_demarrage_kWhj = energie_thermique_demarrage_kWhj_f
    Else
        conso_spec_demarrage_pourcent = 0.3753 * (capacite_incineration_kWhj * pd_NOMBRE_JOUR_PAR_AN) ^ (-0.0468)    'Consommation spécifique démarrage %
        If conso_spec_demarrage_pourcent < 0.157 Then
            energie_thermique_demarrage_kWhj = 194197    'kWh/an sur la base d'un arrêt par an
        ElseIf conso_spec_demarrage_pourcent < 0.158 Then
            energie_thermique_demarrage_kWhj = 194197    'kWh/an sur la base d'un arrêt par an
        ElseIf conso_spec_demarrage_pourcent < 0.159 Then
            energie_thermique_demarrage_kWhj = 168161    'kWh/an sur la base d'un arrêt par an
        ElseIf conso_spec_demarrage_pourcent < 0.161 Then
            energie_thermique_demarrage_kWhj = 143999    'kWh/an sur la base d'un arrêt par an
        ElseIf conso_spec_demarrage_pourcent < 0.163 Then
            energie_thermique_demarrage_kWhj = 117477    'kWh/an sur la base d'un arrêt par an
        ElseIf conso_spec_demarrage_pourcent < 0.165 Then
            energie_thermique_demarrage_kWhj = 93652    'kWh/an sur la base d'un arrêt par an
        ElseIf conso_spec_demarrage_pourcent < 0.168 Then
            energie_thermique_demarrage_kWhj = 69265    'kWh/an sur la base d'un arrêt par an
        ElseIf conso_spec_demarrage_pourcent < 0.172 Then
            energie_thermique_demarrage_kWhj = 48549    'kWh/an sur la base d'un arrêt par an
        Else
            energie_thermique_demarrage_kWhj = 29369    'kWh/an sur la base d'un arrêt par an
        End If
        energie_thermique_demarrage_kWhj = energie_thermique_demarrage_kWhj / pd_NOMBRE_JOUR_PAR_AN   'on répartit la consommation de démarrage sur l'année
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "46.11"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer

    'Boues en entrée de séchage
    choix_inlet(1) = Feuil46.Range("incineration_choix_inlet1").Cells(1, pi_FiliereConsideree)
    inlet_ratio_admis(1) = Feuil46.Range("incineration_inlet1_ratio").Cells(1, pi_FiliereConsideree)

    choix_inlet(2) = Feuil46.Range("incineration_choix_inlet2").Cells(1, pi_FiliereConsideree)
    inlet_ratio_admis(2) = Feuil46.Range("incineration_inlet2_ratio").Cells(1, pi_FiliereConsideree)

    'on retranscrit le choix réalisé dans la synthèse pour l'adapter à notre matrice
    For I = 1 To nb_inlet
        If choix_inlet(I) <> inlet_aucun Then
            If choix_inlet(I) <= graisses Then
                inlet_etape_anterieure(I) = boues_inlet
                inlet_type_boues(I) = choix_inlet(I)
            Else
                inlet_etape_anterieure(I) = choix_inlet(I) - graisses + 1
                'ici on prendra directement les valeurs du flux antérieur
            End If
        End If
    Next I


    For I = 1 To 2
        incineration_repere_boues_inlet(I) = inlet_etape_anterieure(I)
        If incineration_repere_boues_inlet(I) >= boues_sechees_inos(1) And incineration_repere_boues_inlet(I) <= boues_sechees(2) Then
            incineration_siccite_ajustable(I) = True
        Else
            incineration_siccite_ajustable(I) = False
        End If
    Next I

    'lecture des choix internes
    choix_combustible = Feuil46.Range("incineration_choix_combustible").Cells(1, pi_FiliereConsideree)
    If choix_combustible = combustible_biogaz Then
        choix_combustible = combustible_biogaz   'ligne inutile
    Else   'combustible externe --> on lit dans la gestion de l'énergie thermique
        choix_combustible = Feuil41.Range("thermique_choix_combustible").Cells(1, pi_FiliereConsideree)
    End If

    mi_ChoixTraitementNOxFour = Feuil46.Range("incineration_ChoixTraitementNOxFour").Cells(1, pi_FiliereConsideree)

    choix_nb_jours_travailles = Feuil46.Range("incineration_choix_sel_maintien").Cells(1, pi_FiliereConsideree)
    choix_traitement_fumees = Feuil46.Range("incineration_choix_traitement_fum").Cells(1, pi_FiliereConsideree)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "46.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim column As Integer

    column = pi_FiliereConsideree

    mn_BouesEntreesRatioSoufreSurMV_force = Feuil46.Range("incineration_BouesEntreeRatioSoufreSurMV").Cells(1, column)
    mn_BouesEntreesRatioChloreSurMV_force = Feuil46.Range("incineration_BouesEntreeRatioChloreSurMV").Cells(1, column)
    inlet_boues_temperature_force = Feuil46.Range("incineration_inlet_temp_boues").Cells(1, column)
    PCI_MV_kWh_kg_force = Feuil46.Range("incineration_PCI_MV_kWh_kg").Cells(1, column)
    PCI_BH_kWh_kg_force = Feuil46.Range("incineration_PCI_BH_kWh_kg").Cells(1, column)
    combustible_Qvol_force = Feuil46.Range("incineration_combustible_Qvol").Cells(1, column)
    energie_refroidissement_incinerateur_kWhj_force = Feuil46.Range("incineration_energie_refroidissement_incinerateur_kWhj").Cells(1, column)
    debit_air_Nm3j_force = Feuil46.Range("incineration_debit_air_Nm3j").Cells(1, column)
    T_air_inlet_°C_force = Feuil46.Range("incineration_T_air_inlet_C").Cells(1, column)
    T_air_prechauffe_°C_force = Feuil46.Range("incineration_T_apres_prechauff").Cells(1, column)
    T_four_°C_force = Feuil46.Range("incineration_T_four__C").Cells(1, column)
    debit_fumees_Nm3j_force = Feuil46.Range("incineration_debit_fumees_Nm3j").Cells(1, column)
    If mi_ChoixTraitementNOxFour <> mi_PAS_DE_TRAITEMENT_SNCR Then
        mn_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3_force = Feuil46.Range("incineration_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3").Cells(1, column)
        mn_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3_force = Feuil46.Range("incineration_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3").Cells(1, column)
        mn_TraitementSNCRConsommationNH3Pur_kgj_force = Feuil46.Range("incineration_TraitementSNCRConsommationNH3Pur_kgj").Cells(1, column)
        mn_TraitementSNCRConsommationUreePure_kgj_force = Feuil46.Range("incineration_TraitementSNCRConsommationUreePure_kgj").Cells(1, column)
    End If
    T_fumees_inlet_electrofiltre_°C_force = Feuil46.Range("incineration_T_fumees_inlet_electrofiltre__C").Cells(1, column)
    rendement_electrofiltre_force = Feuil46.Range("incineration_rendement_electrofiltre").Cells(1, column)
    mn_CoefficientAutoneutralisationSoufreElectrofiltre_force = Feuil46.Range("incineration_CoefficientAutoneutralisationSoufreElectrofiltre").Cells(1, column)
    fumees_sortie_electrofiltre_mgSO2_Nm3_force = Feuil46.Range("incineration_fumees_sortie_electrofiltre_mgSO2_Nm3").Cells(1, column)
    fumees_sortie_electrofiltre_mgHCl_Nm3_force = Feuil46.Range("incineration_fumees_sortie_electrofiltre_mgHCl_Nm3").Cells(1, column)
    siccite_cendres_force = Feuil46.Range("incineration_siccite_cendres").Cells(1, column)
    cendres_kgj_force = Feuil46.Range("incineration_cendres_kgj").Cells(1, column)
    siccite_REFIB_force = Feuil46.Range("incineration_siccite_REFIB").Cells(1, column)
    REFIB_kgj_force = Feuil46.Range("incineration_REFIB_kgj").Cells(1, column)
    If choix_traitement_fumees = fumees_filiere_seche Then
        T_fumees_outlet_electrofiltre_°C_force = Feuil46.Range("incineration_T_fumees_outlet_electrofiltre__C").Cells(1, column)
        T_fumees_inlet_filtre_manches_°C_force = Feuil46.Range("incineration_T_fumees_inlet_filtre_manches__C").Cells(1, column)
        air_dilution_Nm3j_force = Feuil46.Range("incineration_air_dilution_Nm3j").Cells(1, column)
    End If
    NaOHpur_kgj_force = Feuil46.Range("incineration_NaOHpur_kgj").Cells(1, column)
    NaHCO3pur_kgj_force = Feuil46.Range("incineration_HCO3pur_kgj").Cells(1, column)
    charbon_actif_kgj_force = Feuil46.Range("incineration_charbon_actif_kgj").Cells(1, column)
    four_diametre_force = Feuil46.Range("incineration_four_diametre").Cells(1, column)
    besoins_thermiques_air_kWhj_force = Feuil46.Range("incineration_besoins_thermiques_air_kWhj").Cells(1, column)
    energie_recuperee_fumees_kWhj_force = Feuil46.Range("incineration_energie_recuperee_fumees_kWhj").Cells(1, column)
    energie_thermique_maintien_kWhj_force = Feuil46.Range("incineration_energie_thermique_maintien_kWhj").Cells(1, column)
    energie_thermique_demarrage_kWhj_force = Feuil46.Range("incineration_energie_thermique_demarrage_kWhj").Cells(1, column)
    mn_ConsommationElectriqueSpecifique_kWh_tMS_force = Feuil46.Range("incineration_ConsommationElectriqueSpecifique_kWh_tMS").Cells(1, column)

    If mn_BouesEntreesRatioSoufreSurMV_force = True Then
        md_BouesEntreesRatioSoufreSurMV_f = Feuil46.Range("incineration_BouesEntreeRatioSoufreSurMV").Cells(1, column)
    End If
    If mn_BouesEntreesRatioChloreSurMV_force = True Then
        md_BouesEntreesRatioChloreSurMV_f = Feuil46.Range("incineration_BouesEntreeRatioChloreSurMV").Cells(1, column)
    End If
    If inlet_boues_temperature_force = True Then
        inlet_boues_temperature_f = Feuil46.Range("incineration_inlet_temp_boues").Cells(1, column)
    End If
    If PCI_MV_kWh_kg_force = True Then
        PCI_MV_kWh_kg_f = PCI_kWh_kg(Feuil46.Range("incineration_PCI_MV_kWh_kg").Cells(1, column))
    End If
    If PCI_BH_kWh_kg_force = True Then
        PCI_BH_kWh_kg_f = PCI_kWh_kg(Feuil46.Range("incineration_PCI_BH_kWh_kg").Cells(1, column))
    End If
    If combustible_Qvol_force = True Then
        combustible_Qvol_f = Feuil46.Range("incineration_combustible_Qvol").Cells(1, column)
    End If
    If energie_refroidissement_incinerateur_kWhj_force = True Then
        energie_refroidissement_incinerateur_kWhj_f = Feuil46.Range("incineration_energie_refroidissement_incinerateur_kWhj").Cells(1, column)
    End If
    If debit_air_Nm3j_force = True Then
        debit_air_Nm3j_f = Feuil46.Range("incineration_debit_air_Nm3j").Cells(1, column)
    End If
    If T_air_inlet_°C_force = True Then
        T_air_inlet_°C_f = Feuil46.Range("incineration_T_air_inlet_C").Cells(1, column)
    End If
    If T_air_prechauffe_°C_force = True Then
        T_air_prechauffe_°C_f = Feuil46.Range("incineration_T_apres_prechauff").Cells(1, column)
    End If
    If T_four_°C_force = True Then
        T_four_°C_f = Feuil46.Range("incineration_T_four__C").Cells(1, column)
    End If
    If debit_fumees_Nm3j_force = True Then
        debit_fumees_Nm3j_f = Feuil46.Range("incineration_debit_fumees_Nm3j").Cells(1, column)
    End If
    If mn_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3_force = True Then
        md_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3_f = Feuil46.Range("incineration_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3").Cells(1, column)
    End If
    If mn_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3_force = True Then
        md_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3_f = Feuil46.Range("incineration_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3").Cells(1, column)
    End If
    If mn_TraitementSNCRConsommationNH3Pur_kgj_force = True Then
        md_TraitementSNCRConsommationNH3Pur_kgj_f = Feuil46.Range("incineration_TraitementSNCRConsommationNH3Pur_kgj").Cells(1, column)
    End If
    If mn_TraitementSNCRConsommationUreePure_kgj_force = True Then
        md_TraitementSNCRConsommationUreePure_kgj_f = Feuil46.Range("incineration_TraitementSNCRConsommationUreePure_kgj").Cells(1, column)
    End If
    If T_fumees_inlet_electrofiltre_°C_force = True Then
        T_fumees_inlet_electrofiltre_°C_f = Feuil46.Range("incineration_T_fumees_inlet_electrofiltre__C").Cells(1, column)
    End If
    If T_fumees_outlet_electrofiltre_°C_force = True Then
        T_fumees_outlet_electrofiltre_°C_f = Feuil46.Range("incineration_T_fumees_outlet_electrofiltre__C").Cells(1, column)
    End If
    If T_fumees_inlet_filtre_manches_°C_force = True Then
        T_fumees_inlet_filtre_manches_°C_f = Feuil46.Range("incineration_T_fumees_inlet_filtre_manches__C").Cells(1, column)
    End If
    If four_diametre_force = True Then
        four_diametre_f = Feuil46.Range("incineration_four_diametre").Cells(1, column)
    End If
    If energie_thermique_maintien_kWhj_force = True Then
        energie_thermique_maintien_kWhj_f = Feuil46.Range("incineration_energie_thermique_maintien_kWhj").Cells(1, column)
    End If
    If energie_thermique_demarrage_kWhj_force = True Then
        energie_thermique_demarrage_kWhj_f = Feuil46.Range("incineration_energie_thermique_demarrage_kWhj").Cells(1, column)
    End If
    If rendement_electrofiltre_force = True Then
        rendement_electrofiltre_f = Feuil46.Range("incineration_rendement_electrofiltre").Cells(1, column)
    End If
    If mn_CoefficientAutoneutralisationSoufreElectrofiltre_force = True Then
        md_CoefficientAutoneutralisationSoufreElectrofiltre_f = Feuil46.Range("incineration_CoefficientAutoneutralisationSoufreElectrofiltre").Cells(1, column)
    End If
    If siccite_cendres_force = True Then
        siccite_cendres_f = Feuil46.Range("incineration_siccite_cendres").Cells(1, column)
    End If
    If cendres_kgj_force = True Then
        cendres_kgj_f = Feuil46.Range("incineration_cendres_kgj").Cells(1, column)
    End If
    If siccite_REFIB_force = True Then
        siccite_REFIB_f = Feuil46.Range("incineration_siccite_REFIB").Cells(1, column)
    End If
    If REFIB_kgj_force = True Then
        REFIB_kgj_f = Feuil46.Range("incineration_REFIB_kgj").Cells(1, column)
    End If
    If fumees_sortie_electrofiltre_mgSO2_Nm3_force = True Then
        fumees_sortie_electrofiltre_mgSO2_Nm3_f = Feuil46.Range("incineration_fumees_sortie_electrofiltre_mgSO2_Nm3").Cells(1, column)
    End If
    If fumees_sortie_electrofiltre_mgHCl_Nm3_force = True Then
        fumees_sortie_electrofiltre_mgHCl_Nm3_f = Feuil46.Range("incineration_fumees_sortie_electrofiltre_mgHCl_Nm3").Cells(1, column)
    End If
    If air_dilution_Nm3j_force = True Then
        air_dilution_Nm3j_f = Feuil46.Range("incineration_air_dilution_Nm3j").Cells(1, column)
    End If
    If NaOHpur_kgj_force = True Then
        NaOHpur_kgj_f = Feuil46.Range("incineration_NaOHpur_kgj").Cells(1, column)
    End If
    If NaHCO3pur_kgj_force = True Then
        NaHCO3pur_kgj_f = Feuil46.Range("incineration_HCO3pur_kgj").Cells(1, column)
    End If
    If charbon_actif_kgj_force = True Then
        charbon_actif_kgj_f = Feuil46.Range("incineration_charbon_actif_kgj").Cells(1, column)
    End If
    If besoins_thermiques_air_kWhj_force = True Then
        besoins_thermiques_air_kWhj_f = Feuil46.Range("incineration_besoins_thermiques_air_kWhj").Cells(1, column)
    End If
    If energie_recuperee_fumees_kWhj_force = True Then
        energie_recuperee_fumees_kWhj_f = Feuil46.Range("incineration_energie_recuperee_fumees_kWhj").Cells(1, column)
    End If
    If mn_ConsommationElectriqueSpecifique_kWh_tMS_force = True Then
        md_ConsommationElectriqueSpecifique_kWh_tMS_f = Feuil46.Range("incineration_ConsommationElectriqueSpecifique_kWh_tMS").Cells(1, column)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "46.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub fonctionnement_reel(ByRef boues_flux, ByRef boues_pollution_soluble, ByRef besoins_th, ByRef disponibilites_thermiques, ByVal biogaz_compo_CH4 As Double, ByRef DCO_oxydation)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Q_retour As Double

    Dim inlet_MES_detail(2) As Double
    Dim inlet_Q_detail(2) As Double

    Dim intermediaire_MES As Double
    Dim intermediaire_Q As Double
    Dim intermediaire_NK As Double
    Dim intermediaire_Pt As Double
    Dim intermediaire_DCO As Double
    Dim intermediaire_DBO As Double
    Dim inlet_DCO As Double
    Dim outlet_DCO As Double
    Dim inlet_DBO As Double
    Dim outlet_DBO As Double
    Dim inlet_NK As Double
    Dim outlet_NK As Double
    Dim inlet_Pt As Double
    Dim outlet_Pt As Double
    Dim ratio_DCO As Double
    Dim intermediaire_MV As Double
    Dim intermediaire_reduction_MV As Double
    Dim inlet_MV As Double
    Dim inlet_MV_pci As Double
    Dim outlet_MV As Double
    Dim outlet_concentration As Double
    Dim soluble_kg(nb_repere_mgL) As Double
    Dim condensat_direct_eau_inlet_kgj As Double
    Dim EntreeBouesSoufre_kgj As Double
    Dim EntreeBouesChlore_kgj As Double

    Dim ratio_energie_fumeesETpertes_fumees As Double

    Dim energie_appoint_kWhj As Double   'à ajouter
    Dim combustible_kgj As Double
    Dim densite_biogaz As Double

    Dim siccite_autocombustibilite As Double

    Dim NaHCO3pur_kgj_traitSO2 As Double
    Dim NaHCO3pur_kgj_traitHCl As Double
    Dim NaHCO3pur_kgj_exces As Double

    Dim NaOHpur_kgj_traitSO2 As Double
    Dim NaOHpur_kgj_traitHCl As Double
    Dim NaOHpur_kgj_exces As Double

    'gestion des ratios de MV reduits par type de boues
    Dim MV_reduits(nb_type_boues) As Double
    Dim delta_MV_reduits As Double

    inlet_MES = 0
    inlet_MV = 0
    inlet_Q = 0
    inlet_DCO = 0
    inlet_DBO = 0
    inlet_NK = 0
    inlet_Pt = 0
    EntreeBouesSoufre_kgj = 0
    EntreeBouesChlore_kgj = 0
    'inlet_MES_a = 0
    'inlet_Q_a = 0
    'inlet_MES_b = 0
    'inlet_Q_b = 0

    For I = 1 To 2
        inlet_MES_detail(I) = 0
        inlet_Q_detail(I) = 0
    Next I

    pd_EntreeBouesNonSechees_kgMSj = 0
    
    For I = 1 To nb_inlet
        If choix_inlet(I) <> inlet_aucun Then
            If inlet_etape_anterieure(I) = boues_inlet Then
                'variables d'entree
                intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MES)
                If intermediaire_MES <> 0 Then
                    intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_Q)
                    intermediaire_MV = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MV_MES) * intermediaire_MES
                    intermediaire_DCO = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DCO_MES) * intermediaire_MES
                    intermediaire_DBO = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DBO_MES) * intermediaire_MES
                    intermediaire_NK = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_NK_MES) * intermediaire_MES
                    intermediaire_Pt = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_Pt_MES) * intermediaire_MES

                    inlet_MES_detail(I) = intermediaire_MES
                    inlet_Q_detail(I) = intermediaire_Q

                    'ENTREE SOLUBLE
                    'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                    For k = 1 To nb_repere_mgL
                        soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), k) * intermediaire_Q / 1000
                        boues_pollution_soluble(boues_incinerees, inlet_type_boues(I), k) = 0
                    Next k

                    inlet_MES = inlet_MES + intermediaire_MES
                        pd_EntreeBouesNonSechees_kgMSj = pd_EntreeBouesNonSechees_kgMSj + intermediaire_MES
                    inlet_Q = inlet_Q + intermediaire_Q
                    inlet_DCO = inlet_DCO + intermediaire_DCO + soluble_kg(repere_mgL_DCO)
                    inlet_DBO = inlet_DBO + intermediaire_DBO
                    inlet_NK = inlet_NK + intermediaire_NK + soluble_kg(repere_mgL_NK)
                    inlet_Pt = inlet_Pt + intermediaire_Pt + soluble_kg(repere_mgL_Pt)
                    inlet_MV = inlet_MV + intermediaire_MV
                    EntreeBouesSoufre_kgj = EntreeBouesSoufre_kgj + md_HypotheseRatioSoufreSurMVBoues(inlet_type_boues(I)) * intermediaire_MV
                    EntreeBouesChlore_kgj = EntreeBouesChlore_kgj + md_HypotheseRatioChloreSurMVBoues(inlet_type_boues(I)) * intermediaire_MV

                    boues_flux(boues_incinerees, inlet_type_boues(I), repere_flux_in) = boues_flux(boues_incinerees, inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_incinerees, inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_incinerees, inlet_type_boues(I), repere_flux_in)

                    'je retranche la part de flux d'entree utilisee (objectif: verif des bilans matières)
                    boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                End If
            Else
                For j = 1 To nb_type_boues
                    'variables d'entree
                    intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_MES)
                    If intermediaire_MES <> 0 Then
                        intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_Q)
                        intermediaire_MV = boues_flux(inlet_etape_anterieure(I), j, repere_MV_MES) * intermediaire_MES
                        intermediaire_DCO = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DCO_MES) * intermediaire_MES
                        intermediaire_DBO = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DBO_MES) * intermediaire_MES
                        intermediaire_NK = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_NK_MES) * intermediaire_MES
                        intermediaire_Pt = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_Pt_MES) * intermediaire_MES

                        inlet_MES_detail(I) = inlet_MES_detail(I) + intermediaire_MES
                        inlet_Q_detail(I) = inlet_Q_detail(I) + intermediaire_Q

                        'ENTREE SOLUBLE
                        'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                        For k = 1 To nb_repere_mgL
                            soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), j, k) * intermediaire_Q / 1000
                            boues_pollution_soluble(boues_incinerees, j, k) = 0
                        Next k

                        inlet_MES = inlet_MES + intermediaire_MES
                            If inlet_etape_anterieure(I) <> boues_sechees_inos(1) And inlet_etape_anterieure(I) <> boues_sechees_inos(2) And inlet_etape_anterieure(I) <> boues_sechees_bioco(1) And inlet_etape_anterieure(I) <> boues_sechees_bioco(2) And inlet_etape_anterieure(I) <> boues_sechees(1) And inlet_etape_anterieure(I) <> boues_sechees(2) Then
                                pd_EntreeBouesNonSechees_kgMSj = pd_EntreeBouesNonSechees_kgMSj + intermediaire_MES
                            End If
                        inlet_Q = inlet_Q + intermediaire_Q
                        inlet_DCO = inlet_DCO + intermediaire_DCO + soluble_kg(repere_mgL_DCO)
                        inlet_DBO = inlet_DBO + intermediaire_DBO
                        inlet_NK = inlet_NK + intermediaire_NK + soluble_kg(repere_mgL_NK)
                        inlet_Pt = inlet_Pt + intermediaire_Pt + soluble_kg(repere_mgL_Pt)
                        inlet_MV = inlet_MV + intermediaire_MV
                        EntreeBouesSoufre_kgj = EntreeBouesSoufre_kgj + md_HypotheseRatioSoufreSurMVBoues(j) * intermediaire_MV
                        EntreeBouesChlore_kgj = EntreeBouesChlore_kgj + md_HypotheseRatioChloreSurMVBoues(j) * intermediaire_MV

                        boues_flux(boues_incinerees, j, repere_flux_in) = boues_flux(boues_incinerees, j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                        boues_flux(boues_incinerees, j, repere_verif_flux) = boues_flux(boues_incinerees, j, repere_flux_in)

                        'je retranche la part de flux d'entree utilisee (objectif: verif des bilans matières)
                        boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                    End If
                Next j
            End If
        End If
    Next I

    If inlet_MES = 0 Or inlet_Q = 0 Then
        MsgBox Feuil5.Range("erreur_choix_incineration_inlet") & " ", Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If

    If mn_BouesEntreesRatioSoufreSurMV_force = True Then
        md_BouesEntreesRatioSoufreSurMV = md_BouesEntreesRatioSoufreSurMV_f
    Else
        md_BouesEntreesRatioSoufreSurMV = EntreeBouesSoufre_kgj / inlet_MV
    End If
    If mn_BouesEntreesRatioChloreSurMV_force = True Then
        md_BouesEntreesRatioChloreSurMV = md_BouesEntreesRatioChloreSurMV_f
    Else
        md_BouesEntreesRatioChloreSurMV = EntreeBouesChlore_kgj / inlet_MV
    End If

    DCO_oxydation = DCO_oxydation + inlet_DCO

    '*********************** BILAN MASSE ***********************************************************************************

    'On récupère les données sur les boues d'entrée
    inlet_siccite = inlet_MES / (inlet_Q * masse_volumique_boues)
    inlet_BH = inlet_MES / inlet_siccite
    inlet_MV_MES = inlet_MV / inlet_MES

    'Calcul du PCI de la boue
    If PCI_BH_kWh_kg_force = True Then
        PCI_BH_kWh_kg = PCI_BH_kWh_kg_f
    Else
        PCI_BH_kWh_kg = inlet_siccite * inlet_MV_MES * PCI_MV_kWh_kg - deltaH_vap_kWh_kg * (1 - inlet_siccite)
    End If
    Cp_boues_kWh_kg_K = inlet_siccite * Cp_MS_kWh_kg_K + (1 - inlet_siccite) * Cp_eau_kWh_kg_K

    'Calcul du PCI d'autocombustibilité (avec la siccité actuelle) et de la siccité d'autocombustibilité (MS actuels reçus)
    ratio_energie_fumeesETpertes_fumees = 1 / (1 - ratio_pertes_thermiques)
    If pi_ChoixDigestion = True Then
        PCI_BH_autocombustibilite_kWh_kg = (ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C - Cp_boues_kWh_kg_K * inlet_boues_temperature) + air_inlet_kgair_kgMV_boues_digerees * inlet_MV_MES * inlet_siccite * (ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C - Cp_air_kWh_kg_K * T_air_prechauffe_°C)
        'PCI_BH_autocombustibilite_kWh_kg = pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE * (-0.46 * inlet_boues_temperature + 858.5)
        siccite_autocombustibilite = ((ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C - Cp_boues_kWh_kg_K * inlet_boues_temperature) + deltaH_vap_kWh_kg) / (PCI_MV_kWh_kg * inlet_MV_MES + deltaH_vap_kWh_kg + air_inlet_kgair_kgMV_boues_digerees * inlet_MV_MES * (Cp_air_kWh_kg_K * T_air_prechauffe_°C - ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C))
    Else
        PCI_BH_autocombustibilite_kWh_kg = (ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C - Cp_boues_kWh_kg_K * inlet_boues_temperature) + air_inlet_kgair_kgMV_boues_non_digerees * inlet_MV_MES * inlet_siccite * (ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C - Cp_air_kWh_kg_K * T_air_prechauffe_°C)
        'PCI_BH_autocombustibilite_kWh_kg = pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE * (-0.493 * inlet_boues_temperature + 893)
        siccite_autocombustibilite = ((ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C - Cp_boues_kWh_kg_K * inlet_boues_temperature) + deltaH_vap_kWh_kg) / (PCI_MV_kWh_kg * inlet_MV_MES + deltaH_vap_kWh_kg + air_inlet_kgair_kgMV_boues_non_digerees * inlet_MV_MES * (Cp_air_kWh_kg_K * T_air_prechauffe_°C - ratio_energie_fumeesETpertes_fumees * Cp_fumees_kWh_kg_K * T_four_°C))
    End If

    'On gère l'attribution de la siccité d'autocombustibilité
    If incineration_siccite_ajustable(1) = True Then
        If incineration_siccite_ajustable(2) = True Then
            For I = 1 To nb_inlet
                incineration_siccite_autocombustibilite(I) = siccite_autocombustibilite
            Next I
        Else
            If inlet_MES_detail(2) > 0 Then
                incineration_siccite_autocombustibilite(1) = inlet_MES_detail(1) / ((inlet_MES / (masse_volumique_boues * siccite_autocombustibilite) - inlet_Q_detail(2)) * masse_volumique_boues)
            Else
                incineration_siccite_autocombustibilite(1) = siccite_autocombustibilite
            End If
        End If
    ElseIf incineration_siccite_ajustable(2) = True Then
        If inlet_MES_detail(1) > 0 Then
            incineration_siccite_autocombustibilite(2) = inlet_MES_detail(2) / ((inlet_MES / (masse_volumique_boues * siccite_autocombustibilite) - inlet_Q_detail(1)) * masse_volumique_boues)
        Else
            incineration_siccite_autocombustibilite(2) = siccite_autocombustibilite
        End If
    End If
    For I = 1 To nb_inlet
        If incineration_siccite_autocombustibilite(I) > sechage_siccite_maxi Then
            incineration_siccite_autocombustibilite(I) = sechage_siccite_maxi
        End If
    Next I

    'On calcule l'ajout de combustible d'appoint ou le refroidissement
    If Abs(PCI_BH_kWh_kg - PCI_BH_autocombustibilite_kWh_kg) / Abs(PCI_BH_kWh_kg) > critere_convergence Then
        If PCI_BH_kWh_kg < PCI_BH_autocombustibilite_kWh_kg Then
            'énergie en trop à évacuer
            If energie_refroidissement_incinerateur_kWhj_force = True Then
                energie_refroidissement_incinerateur_kWhj = energie_refroidissement_incinerateur_kWhj_f
            Else
                energie_refroidissement_incinerateur_kWhj = 0
            End If
            'énergie manquante
            energie_appoint_kWhj = (PCI_BH_autocombustibilite_kWh_kg - PCI_BH_kWh_kg) * inlet_BH
            'calcul du débit de combustible, on diminue le PCI par l'énergie qu'il faudra pour chauffer l'air de combustion dû au combustible
            If choix_combustible = combustible_repere_gaz_naturel Then
                combustible_Qvol = energie_appoint_kWhj / (PCI_gaz_nat - air_inlet_kgair_kgGazNat * densite_gaz_nat * Cp_air_kWh_kg_K * T_air_prechauffe_°C)
                If combustible_Qvol_force = True Then
                    combustible_Qvol = combustible_Qvol_f
                End If
                combustible_kgj = densite_gaz_nat * combustible_Qvol
                debit_air_kgj = air_inlet_kgair_kgGazNat * combustible_kgj
                besoins_th(energie_combustible) = besoins_th(energie_combustible) + combustible_Qvol * PCI_gaz_nat
            ElseIf choix_combustible = combustible_repere_fioul Then
                combustible_Qvol = energie_appoint_kWhj / (PCI_fioul - air_inlet_kgair_kgfioul * densite_fioul * Cp_air_kWh_kg_K * T_air_prechauffe_°C)
                If combustible_Qvol_force = True Then
                    combustible_Qvol = combustible_Qvol_f
                End If
                combustible_kgj = densite_fioul * combustible_Qvol
                debit_air_kgj = air_inlet_kgair_kgfioul * combustible_kgj
                besoins_th(energie_combustible) = besoins_th(energie_combustible) + combustible_Qvol * PCI_fioul
            Else   'biogaz
                densite_biogaz = Patm_Pa / (R_gaz_parfait * T_normale_K) * (biogaz_compo_CH4 * MM_CH4_kg_mol + (1 - biogaz_compo_CH4) * MM_CO2_kg_mol)
                combustible_Qvol = energie_appoint_kWhj / (PCI_CH4 * biogaz_compo_CH4 - air_inlet_kgair_kgbiogaz * densite_biogaz * Cp_air_kWh_kg_K * T_air_prechauffe_°C)
                If combustible_Qvol_force = True Then
                    combustible_Qvol = combustible_Qvol_f
                End If
                combustible_kgj = densite_biogaz * combustible_Qvol
                debit_air_kgj = air_inlet_kgair_kgbiogaz * combustible_kgj
                besoins_th(energie_biogaz) = besoins_th(energie_biogaz) + combustible_Qvol * PCI_CH4 * biogaz_compo_CH4
            End If
        Else
            'surplus d'énergie
            If energie_refroidissement_incinerateur_kWhj_force = True Then
                energie_refroidissement_incinerateur_kWhj = energie_refroidissement_incinerateur_kWhj_f
            Else
                energie_refroidissement_incinerateur_kWhj = (PCI_BH_kWh_kg - PCI_BH_autocombustibilite_kWh_kg) * inlet_BH
            End If
            energie_appoint_kWhj = 0
            If combustible_Qvol_force = True Then
                combustible_Qvol = combustible_Qvol_f
                If choix_combustible = combustible_repere_gaz_naturel Then
                    combustible_kgj = densite_gaz_nat * combustible_Qvol
                    debit_air_kgj = air_inlet_kgair_kgGazNat * combustible_kgj
                ElseIf choix_combustible = combustible_repere_fioul Then
                    combustible_kgj = densite_fioul * combustible_Qvol
                    debit_air_kgj = air_inlet_kgair_kgfioul * combustible_kgj
                Else
                    densite_biogaz = Patm_Pa / (R_gaz_parfait * T_normale_K) * (biogaz_compo_CH4 * MM_CH4_kg_mol + (1 - biogaz_compo_CH4) * MM_CO2_kg_mol)
                    combustible_kgj = densite_biogaz * combustible_Qvol
                    debit_air_kgj = air_inlet_kgair_kgbiogaz * combustible_kgj
                End If
            Else
                combustible_Qvol = 0
                combustible_kgj = 0
                debit_air_kgj = 0
            End If
        End If
    Else
        If energie_refroidissement_incinerateur_kWhj_force = True Then
            energie_refroidissement_incinerateur_kWhj = energie_refroidissement_incinerateur_kWhj_f
        Else
            energie_refroidissement_incinerateur_kWhj = 0
        End If
        energie_appoint_kWhj = 0
        If combustible_Qvol_force = True Then
            combustible_Qvol = combustible_Qvol_f
            If choix_combustible = combustible_repere_gaz_naturel Then
                combustible_kgj = densite_gaz_nat * combustible_Qvol
                debit_air_kgj = air_inlet_kgair_kgGazNat * combustible_kgj
            ElseIf choix_combustible = combustible_repere_fioul Then
                combustible_kgj = densite_fioul * combustible_Qvol
                debit_air_kgj = air_inlet_kgair_kgfioul * combustible_kgj
            Else
                densite_biogaz = Patm_Pa / (R_gaz_parfait * T_normale_K) * (biogaz_compo_CH4 * MM_CH4_kg_mol + (1 - biogaz_compo_CH4) * MM_CO2_kg_mol)
                combustible_kgj = densite_biogaz * combustible_Qvol
                debit_air_kgj = air_inlet_kgair_kgbiogaz * combustible_kgj
            End If
        Else
            combustible_Qvol = 0
            combustible_kgj = 0
            debit_air_kgj = 0
            energie_appoint_kWhj = 0
        End If
    End If


    'On calcule l'air ajouté pour la combustion des boues
    If pi_ChoixDigestion = True Then  'probabilité faible d'avoir une digestion et une incinération avec des boues différentes
        debit_air_kgj = debit_air_kgj + (air_inlet_kgair_kgMV_boues_non_digerees * inlet_MES * inlet_MV_MES)
    Else
        debit_air_kgj = debit_air_kgj + (air_inlet_kgair_kgMV_boues_digerees * inlet_MES * inlet_MV_MES)
    End If
    debit_air_Nm3j = debit_air_kgj * (R_gaz_parfait * T_normale_K) / (MM_air_kg_mol * Patm_Pa)
    If debit_air_Nm3j_force = True Then
        debit_air_Nm3j = debit_air_Nm3j_f
        debit_air_kgj = debit_air_Nm3j * (MM_air_kg_mol * Patm_Pa) / (R_gaz_parfait * T_normale_K)
    End If

    'On calcule les fumées
    debit_fumees_kgj = debit_air_kgj + combustible_kgj + inlet_BH
    debit_fumees_Nm3j = debit_fumees_kgj * (R_gaz_parfait * T_normale_K) / (MM_fumees_kg_mol * Patm_Pa)
    If debit_fumees_Nm3j_force = True Then
        debit_fumees_Nm3j = debit_fumees_Nm3j_f
        debit_fumees_kgj = debit_fumees_Nm3j * (MM_fumees_kg_mol * Patm_Pa) / (R_gaz_parfait * T_normale_K)
    End If
    
    'Calcul du traitement SNCR
    If mi_ChoixTraitementNOxFour <> mi_PAS_DE_TRAITEMENT_SNCR Then
        Call calcul_traitement_SNCR
        'On regarde si les consommations de réactifs ont été forcées
        If mn_TraitementSNCRConsommationNH3Pur_kgj_force = True Then
            md_TraitementSNCRConsommationNH3Pur_kgj = md_TraitementSNCRConsommationNH3Pur_kgj_f
        End If
        If mn_TraitementSNCRConsommationUreePure_kgj_force = True Then
            md_TraitementSNCRConsommationUreePure_kgj = md_TraitementSNCRConsommationUreePure_kgj_f
        End If
    End If
    

    'On calcule les cendres à évacuer (récupérées sur l'électrofiltre) et les résidus REFIB (récupérés sur les filtres à manches pour la filière sèche ou sur la filière humide en quantité moindre)
    If cendres_kgj_force = True Then
        cendres_kgj = cendres_kgj_f
    Else
        cendres_kgj = rendement_electrofiltre * inlet_MES * (1 - inlet_MV_MES) / siccite_cendres
    End If

    'On calcule la compo en soufre et en HCl des fumées
    If fumees_sortie_electrofiltre_mgSO2_Nm3_force = True Then
        fumees_sortie_electrofiltre_mgSO2_Nm3 = fumees_sortie_electrofiltre_mgSO2_Nm3_f
    Else
        fumees_sortie_electrofiltre_mgSO2_Nm3 = (MM_SO2_kg_mol / pd_MASSE_MOLAIRE_S_kg_mol) * (md_BouesEntreesRatioSoufreSurMV * inlet_MV * (1 - md_CoefficientAutoneutralisationSoufreElectrofiltre)) / debit_fumees_Nm3j * 1000000
    End If
    If fumees_sortie_electrofiltre_mgHCl_Nm3_force = True Then
        fumees_sortie_electrofiltre_mgHCl_Nm3 = fumees_sortie_electrofiltre_mgHCl_Nm3_f
    Else
        fumees_sortie_electrofiltre_mgHCl_Nm3 = (MM_HCl_kg_mol / pd_MASSE_MOLAIRE_Cl_kg_mol) * (md_BouesEntreesRatioChloreSurMV * inlet_MV) / debit_fumees_Nm3j * 1000000
    End If

    'On calcule le bicarbonate
    NaHCO3pur_kgj_traitSO2 = NaHCO3_dose_g_gSO2 * fumees_sortie_electrofiltre_mgSO2_Nm3 / 1000000 * debit_fumees_Nm3j
    NaHCO3pur_kgj_traitHCl = NaHCO3_dose_g_gHCl * fumees_sortie_electrofiltre_mgHCl_Nm3 / 1000000 * debit_fumees_Nm3j
    NaHCO3pur_kgj_exces = (NaHCO3pur_kgj_traitSO2 + NaHCO3pur_kgj_traitHCl) * (1 + NaHCO3_exces)
    NaHCO3pur_kgj = NaHCO3pur_kgj_traitSO2 + NaHCO3pur_kgj_traitHCl + NaHCO3pur_kgj_exces
    If NaHCO3pur_kgj_force = True Then
        NaHCO3pur_kgj = NaHCO3pur_kgj_f
        If NaHCO3pur_kgj < (NaHCO3pur_kgj_traitSO2 + NaHCO3pur_kgj_traitHCl) Then
            NaHCO3pur_kgj_traitSO2 = NaHCO3pur_kgj_traitSO2 / (NaHCO3pur_kgj_traitSO2 + NaHCO3pur_kgj_traitHCl) * NaHCO3pur_kgj
            NaHCO3pur_kgj_traitHCl = NaHCO3pur_kgj - NaHCO3pur_kgj_traitSO2
            NaHCO3pur_kgj_exces = 0
        Else
            NaHCO3pur_kgj_exces = NaHCO3pur_kgj - (NaHCO3pur_kgj_traitSO2 + NaHCO3pur_kgj_traitHCl)
        End If
    End If

    'On calcule la soude
    NaOHpur_kgj_traitSO2 = NaOH_dose_g_gSO2 * fumees_sortie_electrofiltre_mgSO2_Nm3 / 1000000 * debit_fumees_Nm3j
    NaOHpur_kgj_traitHCl = NaOH_dose_g_gHCl * fumees_sortie_electrofiltre_mgHCl_Nm3 / 1000000 * debit_fumees_Nm3j
    NaOHpur_kgj_exces = (NaOHpur_kgj_traitSO2 + NaOHpur_kgj_traitHCl) * (1 + NaOH_exces)
    NaOHpur_kgj = NaOHpur_kgj_traitSO2 + NaOHpur_kgj_traitHCl + NaOHpur_kgj_exces
    If NaOHpur_kgj_force = True Then
        NaOHpur_kgj = NaOHpur_kgj_f
        If NaOHpur_kgj < (NaOHpur_kgj_traitSO2 + NaOHpur_kgj_traitHCl) Then
            NaOHpur_kgj_traitSO2 = NaOHpur_kgj_traitSO2 / (NaOHpur_kgj_traitSO2 + NaOHpur_kgj_traitHCl) * NaOHpur_kgj
            NaOHpur_kgj_traitHCl = NaOHpur_kgj - NaOHpur_kgj_traitSO2
            NaOHpur_kgj_exces = 0
        Else
            NaOHpur_kgj_exces = NaOHpur_kgj - (NaOHpur_kgj_traitSO2 + NaOHpur_kgj_traitHCl)
        End If
    End If

    'On calcule le charbon actif
    If charbon_actif_kgj_force = True Then
        charbon_actif_kgj = charbon_actif_kgj_f
    Else
        charbon_actif_kgj = charbon_actif_dose_kg_m3fumees * debit_fumees_Nm3j * (T_normale_K + T_fumees_outlet_electrofiltre_°C) / T_normale_K
    End If

    'On calcule les REFIB
    If REFIB_kgj_force = True Then
        REFIB_kgj = REFIB_kgj_f
    Else
        REFIB_kgj = (1 - rendement_electrofiltre) * inlet_MES * (1 - inlet_MV_MES)   'résidus non captés avant l'électrofiltre
        REFIB_kgj = REFIB_kgj + NaHCO3pur_kgj_traitSO2 * MM_Na2SO4_kg_mol / (2 * MM_NaHCO3_kg_mol)  'Na2SO4 formé avec le SO2 et le bicarbonate ou la soude
        REFIB_kgj = REFIB_kgj + NaHCO3pur_kgj_traitHCl * MM_NaCl_kg_mol / MM_NaHCO3_kg_mol    'NaCl formé avec le HCl et le bicarbonate ou la soude
        REFIB_kgj = REFIB_kgj + NaHCO3pur_kgj_exces    'excès de bicarbonate
        REFIB_kgj = REFIB_kgj + charbon_actif_kgj
        REFIB_kgj = REFIB_kgj / siccite_REFIB
    End If

    '********************* BILAN THERMIQUE ***********************************************************
    Call bilan_thermique
    '    Call bilan_thermique
    disponibilites_thermiques(energie_eau_HT + 1) = disponibilites_thermiques(energie_eau_HT + 1) + (energie_refroidissement_incinerateur_kWhj + debit_fumees_kgj * Cp_fumees_kWh_kg_K * (T_four_°C - T_fumees_inlet_electrofiltre_°C))
    'besoins_th(energie_eau_HT + 1) = besoins_th(energie_eau_HT + 1) + debit_air_kgj * Cp_air_kWh_kg_K * (T_air_prechauffe_°C - T_air_inlet_°C)
    If choix_combustible = combustible_repere_gaz_naturel Or choix_combustible = combustible_repere_fioul Then
        besoins_th(energie_combustible) = besoins_th(energie_combustible) + energie_thermique_maintien_kWhj + energie_thermique_demarrage_kWhj
    Else
        besoins_th(energie_biogaz) = besoins_th(energie_biogaz) + energie_thermique_maintien_kWhj + energie_thermique_demarrage_kWhj
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "46.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub
Private Function PCI_J_g(ByVal PCI_kWh_kg As Double) As Double

    PCI_J_g = PCI_kWh_kg * pd_NOMBRE_SECONDE_PAR_HEURE


End Function

Private Function PCI_kWh_kg(ByVal PCI_J_g As Double) As Double

    PCI_kWh_kg = PCI_J_g / pd_NOMBRE_SECONDE_PAR_HEURE


End Function

Sub transfert_flux_energie_vers_pinch()

    Dim nom_flux As String
    Dim T_in As Double
    Dim T_out As Double
    Dim MCp As Double
    Dim m As Double
    Dim Cp As Double

    'PYROFLUID - PRECHAUFFAGE AIR

    'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
    '    nom_flux = "PYROFLUID - PRECHAUFFAGE AIR"
    nom_flux = Feuil5.Range("nom_flux_pyrofluid1").Cells(1, 1).Value
    T_in = T_air_inlet_°C
    T_out = T_air_prechauffe_°C
'    m = debit_air_kgj  ' débit en kg/j
'    Cp = Cp_air_kWh_kg_K  'kWh/(kg.°C) - chaleur spécifique de l'air
    MCp = besoins_thermiques_air_kWhj / (T_out - T_in) 'kWh/j/°C

    'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
    Call copie_dans_flux_thermique_process(nom_flux, T_in, T_out, MCp)



    'PYROFLUID - REFROIDISSEMENT FUMEES

    'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
    '    nom_flux = "PYROFLUID - REFROIDISSEMENT FUMEES"
    nom_flux = Feuil5.Range("nom_flux_pyrofluid2").Cells(1, 1).Value
    T_in = T_four_°C
    T_out = T_fumees_inlet_electrofiltre_°C
    MCp = (energie_recuperee_fumees_kWhj + energie_refroidissement_incinerateur_kWhj) / (T_in - T_out)

    'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
    Call copie_dans_flux_thermique_process(nom_flux, T_in, T_out, MCp)



    'PYROFLUID - REFROIDISSEMENT FOUR   On injecte de l'eau dans le four qui ensuite se retrouve dans les fumées donc l'énergie a été ajoutée aux fumées
    '    If energie_refroidissement_incinerateur_kWhj > 0 Then
    '        'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
    '        nom_flux = "PYROFLUID - REFROIDISSEMENT FOUR"
    '        T_in = T_four_°C + 1
    '        T_out = T_four_°C
    '        MCp = energie_refroidissement_incinerateur_kWhj / (T_in - T_out)
    '
    '        'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
    '        Call copie_dans_flux_thermique_process(nom_flux, T_in, T_out, MCp)
    '    End If


End Sub


Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 0    'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil46.Range("incineration_erreur_nom") & " "

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "46.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer
    Ligne = 1
    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil46.Range("incineration_inlet1_ratio").Cells(Ligne, Colonne) = inlet_ratio_admis(1)
    Feuil46.Range("incineration_inlet2_ratio").Cells(Ligne, Colonne) = inlet_ratio_admis(2)
    Feuil46.Range("incineration_inlet_siccite").Cells(Ligne, Colonne) = inlet_siccite
    Feuil46.Range("incineration_inlet_MES").Cells(Ligne, Colonne) = inlet_MES
    Feuil46.Range("incineration_inlet_MV_MES").Cells(Ligne, Colonne) = inlet_MV_MES
    Feuil46.Range("incineration_inlet_BH").Cells(Ligne, Colonne) = inlet_BH
    Feuil46.Range("incineration_BouesEntreeRatioSoufreSurMV").Cells(Ligne, Colonne) = md_BouesEntreesRatioSoufreSurMV
    Feuil46.Range("incineration_BouesEntreeRatioChloreSurMV").Cells(Ligne, Colonne) = md_BouesEntreesRatioChloreSurMV
    Feuil46.Range("incineration_inlet_temp_boues").Cells(Ligne, Colonne) = inlet_boues_temperature
    Feuil46.Range("incineration_PCI_MV_kWh_kg").Cells(Ligne, Colonne) = PCI_J_g(PCI_MV_kWh_kg)
    Feuil46.Range("incineration_PCI_BH_kWh_kg").Cells(Ligne, Colonne) = PCI_J_g(PCI_BH_kWh_kg)
    Feuil46.Range("incineration_debit_air_Nm3j").Cells(Ligne, Colonne) = debit_air_Nm3j
    Feuil46.Range("incineration_combustible_Qvol").Cells(Ligne, Colonne) = combustible_Qvol
    Feuil46.Range("incineration_energie_refroidissement_incinerateur_kWhj").Cells(Ligne, Colonne) = energie_refroidissement_incinerateur_kWhj
    Feuil46.Range("incineration_T_air_inlet_C").Cells(Ligne, Colonne) = T_air_inlet_°C
    Feuil46.Range("incineration_T_apres_prechauff").Cells(Ligne, Colonne) = T_air_prechauffe_°C
    Feuil46.Range("incineration_debit_fumees_Nm3j").Cells(Ligne, Colonne) = debit_fumees_Nm3j
    If mi_ChoixTraitementNOxFour <> mi_PAS_DE_TRAITEMENT_SNCR Then
        Feuil46.Range("incineration_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3").Cells(Ligne, Colonne) = md_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3
        Feuil46.Range("incineration_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3").Cells(Ligne, Colonne) = md_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3
        Feuil46.Range("incineration_TraitementSNCRConsommationNH3Pur_kgj").Cells(Ligne, Colonne) = md_TraitementSNCRConsommationNH3Pur_kgj
        Feuil46.Range("incineration_TraitementSNCRConsommationUreePure_kgj").Cells(Ligne, Colonne) = md_TraitementSNCRConsommationUreePure_kgj
    End If
    Feuil46.Range("incineration_T_four__C").Cells(Ligne, Colonne) = T_four_°C
    Feuil46.Range("incineration_T_fumees_inlet_electrofiltre__C").Cells(Ligne, Colonne) = T_fumees_inlet_electrofiltre_°C
    Feuil46.Range("incineration_rendement_electrofiltre").Cells(Ligne, Colonne) = rendement_electrofiltre
    Feuil46.Range("incineration_CoefficientAutoneutralisationSoufreElectrofiltre").Cells(Ligne, Colonne) = md_CoefficientAutoneutralisationSoufreElectrofiltre
    Feuil46.Range("incineration_fumees_sortie_electrofiltre_mgSO2_Nm3").Cells(Ligne, Colonne) = fumees_sortie_electrofiltre_mgSO2_Nm3
    Feuil46.Range("incineration_fumees_sortie_electrofiltre_mgHCl_Nm3").Cells(Ligne, Colonne) = fumees_sortie_electrofiltre_mgHCl_Nm3
    Feuil46.Range("incineration_siccite_cendres").Cells(Ligne, Colonne) = siccite_cendres
    Feuil46.Range("incineration_siccite_REFIB").Cells(Ligne, Colonne) = siccite_REFIB
    If choix_traitement_fumees = fumees_filiere_seche Then
        Feuil46.Range("incineration_T_fumees_outlet_electrofiltre__C").Cells(Ligne, Colonne) = T_fumees_outlet_electrofiltre_°C
        Feuil46.Range("incineration_T_fumees_inlet_filtre_manches__C").Cells(Ligne, Colonne) = T_fumees_inlet_filtre_manches_°C
        Feuil46.Range("incineration_air_dilution_Nm3j").Cells(Ligne, Colonne) = air_dilution_Nm3j
    End If
    Feuil46.Range("incineration_cendres_kgj").Cells(Ligne, Colonne) = cendres_kgj
    Feuil46.Range("incineration_REFIB_kgj").Cells(Ligne, Colonne) = REFIB_kgj
    Feuil46.Range("incineration_NaOHpur_kgj").Cells(Ligne, Colonne) = NaOHpur_kgj
    Feuil46.Range("incineration_HCO3pur_kgj").Cells(Ligne, Colonne) = NaHCO3pur_kgj
    Feuil46.Range("incineration_charbon_actif_kgj").Cells(Ligne, Colonne) = charbon_actif_kgj
    Feuil46.Range("incineration_four_diametre").Cells(Ligne, Colonne) = four_diametre
    Feuil46.Range("incineration_besoins_thermiques_air_kWhj").Cells(Ligne, Colonne) = besoins_thermiques_air_kWhj
    Feuil46.Range("incineration_energie_recuperee_fumees_kWhj").Cells(Ligne, Colonne) = energie_recuperee_fumees_kWhj
    Feuil46.Range("incineration_energie_thermique_maintien_kWhj").Cells(Ligne, Colonne) = energie_thermique_maintien_kWhj
    Feuil46.Range("incineration_energie_thermique_demarrage_kWhj").Cells(Ligne, Colonne) = energie_thermique_demarrage_kWhj
    Feuil46.Range("incineration_ConsommationElectriqueSpecifique_kWh_tMS").Cells(Ligne, Colonne) = md_ConsommationElectriqueSpecifique_kWh_tMS
    Feuil46.Range("incineration_electricite").Cells(Ligne, Colonne) = electricite



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "46.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer
    Ligne = 1
    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil46.Range("incineration_inlet1_ratio").Cells(Ligne, Colonne) = ""      'inlet_ratio_admis(1)
    Feuil46.Range("incineration_inlet2_ratio").Cells(Ligne, Colonne) = ""      'inlet_ratio_admis(2)
    Feuil46.Range("incineration_inlet_siccite").Cells(Ligne, Colonne) = ""      'inlet_siccite
    Feuil46.Range("incineration_inlet_MES").Cells(Ligne, Colonne) = ""      'inlet_MES
    Feuil46.Range("incineration_inlet_MV_MES").Cells(Ligne, Colonne) = ""      'inlet_MV_MES
    Feuil46.Range("incineration_inlet_BH").Cells(Ligne, Colonne) = ""      'inlet_BH
    Feuil46.Range("incineration_BouesEntreeRatioSoufreSurMV").Cells(Ligne, Colonne) = ""      'md_BouesEntreesRatioSoufreSurMV
    Feuil46.Range("incineration_BouesEntreeRatioChloreSurMV").Cells(Ligne, Colonne) = ""      'md_BouesEntreesRatioChloreSurMV
    Feuil46.Range("incineration_inlet_temp_boues").Cells(Ligne, Colonne) = ""      'inlet_boues_temperature
    Feuil46.Range("incineration_PCI_MV_kWh_kg").Cells(Ligne, Colonne) = ""      'PCI_MV_kWh_kg
    Feuil46.Range("incineration_PCI_BH_kWh_kg").Cells(Ligne, Colonne) = ""      'PCI_BH_kWh_kg
    Feuil46.Range("incineration_combustible_Qvol").Cells(Ligne, Colonne) = ""       'combustible_Qvol
    Feuil46.Range("incineration_energie_refroidissement_incinerateur_kWhj").Cells(Ligne, Colonne) = ""      'energie_refroidissement_incinerateur_kWhj
    Feuil46.Range("incineration_debit_air_Nm3j").Cells(Ligne, Colonne) = ""      'debit_air_Nm3j
    Feuil46.Range("incineration_T_air_inlet_C").Cells(Ligne, Colonne) = ""      'T_air_inlet_°C
    Feuil46.Range("incineration_T_apres_prechauff").Cells(Ligne, Colonne) = ""      'T_air_prechauffe_°C
    Feuil46.Range("incineration_debit_fumees_Nm3j").Cells(Ligne, Colonne) = ""      'debit_fumees_Nm3j
    Feuil46.Range("incineration_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3").Cells(Ligne, Colonne) = ""      'md_ConcentrationNOxFumeesSansTraitementSNCR_mgNO2eq_Nm3
    Feuil46.Range("incineration_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3").Cells(Ligne, Colonne) = ""      'md_ConcentrationNOxFumeesAvecTraitementSNCR_mgNO2eq_Nm3
    Feuil46.Range("incineration_TraitementSNCRConsommationNH3Pur_kgj").Cells(Ligne, Colonne) = ""      'md_TraitementSNCRConsommationNH3Pur_kgj
    Feuil46.Range("incineration_TraitementSNCRConsommationUreePure_kgj").Cells(Ligne, Colonne) = ""      'md_TraitementSNCRConsommationUreePure_kgj
    Feuil46.Range("incineration_T_four__C").Cells(Ligne, Colonne) = ""      'T_four_°C
    Feuil46.Range("incineration_T_fumees_inlet_electrofiltre__C").Cells(Ligne, Colonne) = ""      'T_fumees_inlet_electrofiltre_°C
    Feuil46.Range("incineration_rendement_electrofiltre").Cells(Ligne, Colonne) = ""       'rendement_electrofiltre
    Feuil46.Range("incineration_CoefficientAutoneutralisationSoufreElectrofiltre").Cells(Ligne, Colonne) = ""       'md_CoefficientAutoneutralisationSoufreElectrofiltre
    Feuil46.Range("incineration_fumees_sortie_electrofiltre_mgSO2_Nm3").Cells(Ligne, Colonne) = ""     'fumees_sortie_electrofiltre_mgSO2_Nm3
    Feuil46.Range("incineration_fumees_sortie_electrofiltre_mgHCl_Nm3").Cells(Ligne, Colonne) = ""     'fumees_sortie_electrofiltre_mgHCl_Nm3
    Feuil46.Range("incineration_siccite_cendres").Cells(Ligne, Colonne) = ""     'siccite_cendres
    Feuil46.Range("incineration_T_fumees_outlet_electrofiltre__C").Cells(Ligne, Colonne) = ""      'T_fumees_outlet_electrofiltre_°C
    Feuil46.Range("incineration_T_fumees_inlet_filtre_manches__C").Cells(Ligne, Colonne) = ""     'T_fumees_inlet_filtre_manches_°C
    Feuil46.Range("incineration_air_dilution_Nm3j").Cells(Ligne, Colonne) = ""      'air_dilution_Nm3j
    Feuil46.Range("incineration_cendres_kgj").Cells(Ligne, Colonne) = ""      'cendres_kgj
    Feuil46.Range("incineration_siccite_REFIB").Cells(Ligne, Colonne) = ""    'siccite_REFIB
    Feuil46.Range("incineration_REFIB_kgj").Cells(Ligne, Colonne) = ""      'REFIB_kgj
    Feuil46.Range("incineration_NaOHpur_kgj").Cells(Ligne, Colonne) = ""      'NaOHpur_kgj
    Feuil46.Range("incineration_HCO3pur_kgj").Cells(Ligne, Colonne) = ""      'NaHCO3pur_kgj
    Feuil46.Range("incineration_charbon_actif_kgj").Cells(Ligne, Colonne) = ""      'charbon_actif_kgj
    Feuil46.Range("incineration_four_diametre").Cells(Ligne, Colonne) = ""      'four_diametre
    Feuil46.Range("incineration_besoins_thermiques_air_kWhj").Cells(Ligne, Colonne) = ""      'besoins_thermiques_air_kWhj
    Feuil46.Range("incineration_energie_recuperee_fumees_kWhj").Cells(Ligne, Colonne) = ""      'energie_recuperee_fumees_kWhj
    Feuil46.Range("incineration_energie_thermique_maintien_kWhj").Cells(Ligne, Colonne) = ""     'energie_thermique_maintien_kWhj
    Feuil46.Range("incineration_energie_thermique_demarrage_kWhj").Cells(Ligne, Colonne) = ""     'energie_thermique_demarrage_kWhj
    Feuil46.Range("incineration_ConsommationElectriqueSpecifique_kWh_tMS").Cells(Ligne, Colonne) = ""     'md_ConsommationElectriqueSpecifique_kWh_tMS
    Feuil46.Range("incineration_electricite").Cells(Ligne, Colonne) = ""      'electricite


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "46.10"

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
        Call affich_remarque
    End If

End Sub




