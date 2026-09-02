Attribute VB_Name = "z_Desodorisation_chimique"
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

Private Const choix_tour_oui = 1
Private Const choix_tour_non = 2

Private Const soufflage_oui = 1

'variables de choix
Dim choix_couverture_step As Integer
Dim choix_soufflage_air As Integer
Dim choix_tour_acide As Integer
Dim choix_tour_oxydobase_pH9 As Integer
Dim choix_tour_oxydobase_pH11 As Integer
Dim choix_tour_neutre As Integer

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
Dim air_vicie_NH3 As Double
Dim air_vicie_Norga As Double
Dim air_vicie_H2S As Double
Dim air_vicie_RSH As Double
'Dim H2SO4_titre As Double
Public acide_H2SO4_pur As Double
'Dim acide_H2SO4_commercial As Double
'Dim NaOH_titre As Double
Public oxydobase_pH9_NaOH_pur As Double
'Dim oxydobase_pH9_NaOH_commercial As Double
Public oxydobase_pH11_NaOH_pur As Double
'Dim oxydobase_pH11_NaOH_commercial As Double
'Dim NaOCl_titre As Double
Public oxydobase_Cl2_pur As Double
'Dim oxydobase_NaOCl_commercial As Double
'Dim NaHSO3_titre As Double
Public neutre_NaHSO3_pur As Double
'Dim neutre_NaHSO3_commercial As Double
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
Dim air_vicie_NH3_force As Boolean
Dim air_vicie_Norga_force As Boolean
Dim air_vicie_H2S_force As Boolean
Dim air_vicie_RSH_force As Boolean
'Dim H2SO4_titre_force As Boolean
Dim acide_H2SO4_pur_force As Boolean
'Dim acide_H2SO4_commercial_force As Boolean
'Dim NaOH_titre_force As Boolean
Dim oxydobase_pH9_NaOH_pur_force As Boolean
'Dim oxydobase_pH9_NaOH_commercial_force As Boolean
Dim oxydobase_pH11_NaOH_pur_force As Boolean
'Dim oxydobase_pH11_NaOH_commercial_force As Boolean
'Dim NaOCl_titre_force As Boolean
Dim oxydobase_Cl2_pur_force As Boolean
'Dim oxydobase_NaOCl_commercial_force As Boolean
'Dim NaHSO3_titre_force As Boolean
Dim neutre_NaHSO3_pur_force As Boolean
'Dim neutre_NaHSO3_commercial_force As Boolean

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
Dim air_vicie_NH3_f As Double
Dim air_vicie_Norga_f As Double
Dim air_vicie_H2S_f As Double
Dim air_vicie_RSH_f As Double
'Dim H2SO4_titre_f As Double
Dim acide_H2SO4_pur_f As Double
'Dim acide_H2SO4_commercial_f As Double
'Dim NaOH_titre_f As Double
Dim oxydobase_pH9_NaOH_pur_f As Double
'Dim oxydobase_pH9_NaOH_commercial_f As Double
Dim oxydobase_pH11_NaOH_pur_f As Double
'Dim oxydobase_pH11_NaOH_commercial_f As Double
'Dim NaOCl_titre_f As Double
Dim oxydobase_Cl2_pur_f As Double
'Dim oxydobase_NaOCl_commercial_f As Double
'Dim NaHSO3_titre_f As Double
Dim neutre_NaHSO3_pur_f As Double
'Dim neutre_NaHSO3_commercial_f As Double

'hypothèses
Dim rate_air_vicie_Qnominal As Double
Dim ventilateur_rdt_machine As Double
Dim ventilateur_rdt_moteur As Double
Dim a_densite_H2SO4(6) As Double
Dim masse_volumique_H2SO4 As Double
Dim acide_rate_H2SO4_NH3 As Double
Dim acide_rate_H2SO4_Norga As Double
Dim oxydobase_pH9_rate_NaOH_H2S As Double
Dim oxydobase_pH9_rate_NaOH_RSH As Double
Dim oxydobase_pH9_rate_NaOH_CO2 As Double
Dim oxydobase_pH11_rate_NaOH_H2S As Double
Dim oxydobase_pH11_rate_NaOH_RSH As Double
Dim oxydobase_pH11_rate_NaOH_CO2 As Double
Dim NaOH_concentration As Double
Dim concentration_Cl2_javel As Double
Dim oxydobase_rate_Cl2_H2S As Double
Dim oxydobase_rate_Cl2_RSH As Double
Dim masse_volumique_NaHSO3 As Double
Dim neutre_rate_NaHSO3_Cl2 As Double
Dim residuel_Cl2_mg_Nm3 As Double

Dim nb_tours_lavage As Integer
Dim message_erreur As Integer

Dim remarque(3) As String

Sub calcul(ByVal HS_strippe)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim NH3_flux_gh As Double   'en g/h
    Dim Norga_flux_gh As Double   'en g/h
    Dim H2S_flux_gh As Double   'en g/h
    Dim RSH_flux_gh As Double   'en g/h



    'QUALITE DE L'AIR A TRAITER
    NH3_flux_gh = air_vicie_Q_Nm3h * air_vicie_NH3 / 1000
    Norga_flux_gh = air_vicie_Q_Nm3h * air_vicie_Norga / 1000
    RSH_flux_gh = air_vicie_Q_Nm3h * air_vicie_RSH / 1000

    If air_vicie_H2S_force = False Then
        air_vicie_H2S = (HS_strippe / pd_NOMBRE_HEURE_PAR_JOUR) / air_vicie_Q_Nm3h * 1000000
    End If
    H2S_flux_gh = air_vicie_Q_Nm3h * air_vicie_H2S / 1000


    'REACTIFS

    'tour acide
    If choix_tour_acide = choix_tour_oui Then
        If acide_H2SO4_pur_force = True Then
            acide_H2SO4_pur = acide_H2SO4_pur_f
        Else
            acide_H2SO4_pur = acide_rate_H2SO4_NH3 * NH3_flux_gh / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
            acide_H2SO4_pur = acide_H2SO4_pur + acide_rate_H2SO4_Norga * Norga_flux_gh / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
        End If
        '        If acide_H2SO4_commercial_force = True Then
        '            acide_H2SO4_commercial = acide_H2SO4_commercial_f
        '        Else
        '            acide_H2SO4_commercial = acide_H2SO4_pur / (H2SO4_titre * masse_volumique_H2SO4 / 1000)
        '        End If
        'rendements de l'ordre de 99% voire plus donc --> 100%
        NH3_flux_gh = 0
        Norga_flux_gh = 0
    End If

    'tour oxydo-basique pH=9-9,5
    If choix_tour_oxydobase_pH9 = choix_tour_oui Then
        'NaOH
        If oxydobase_pH9_NaOH_pur_force = True Then
            oxydobase_pH9_NaOH_pur = oxydobase_pH9_NaOH_pur_f
        Else
            oxydobase_pH9_NaOH_pur = oxydobase_pH9_rate_NaOH_H2S * H2S_flux_gh / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
            oxydobase_pH9_NaOH_pur = oxydobase_pH9_NaOH_pur + oxydobase_pH9_rate_NaOH_RSH * RSH_flux_gh / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
            'consommation due au CO2
            oxydobase_pH9_NaOH_pur = oxydobase_pH9_NaOH_pur + oxydobase_pH9_rate_NaOH_CO2 * (air_vicie_Q_Nm3h / 1000) / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
        End If
        '        If oxydobase_pH9_NaOH_commercial_force = True Then
        '            oxydobase_pH9_NaOH_commercial = oxydobase_pH9_NaOH_commercial_f
        '        Else
        '            oxydobase_pH9_NaOH_commercial = oxydobase_pH9_NaOH_pur / (NaOH_concentration / 1000)
        '        End If
        'NaOCl
        If oxydobase_Cl2_pur_force = True Then
            oxydobase_Cl2_pur = oxydobase_Cl2_pur_f
        Else
            oxydobase_Cl2_pur = oxydobase_rate_Cl2_H2S * H2S_flux_gh / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
            oxydobase_Cl2_pur = oxydobase_Cl2_pur + oxydobase_rate_Cl2_RSH * RSH_flux_gh / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
        End If
        '        If oxydobase_NaOCl_commercial_force = True Then
        '            oxydobase_NaOCl_commercial = oxydobase_NaOCl_commercial_f
        '        Else
        '            oxydobase_NaOCl_commercial = oxydobase_Cl2_pur / (concentration_Cl2_javel / 1000)
        '        End If
        'rendements de l'ordre de 99% voire plus donc --> 100%
        H2S_flux_gh = 0
        RSH_flux_gh = 0
    End If

    'tour oxydo-basique pH=10,5-11
    If choix_tour_oxydobase_pH11 = choix_tour_oui Then
        If oxydobase_pH11_NaOH_pur_force = True Then
            oxydobase_pH11_NaOH_pur = oxydobase_pH11_NaOH_pur_f
        Else
            oxydobase_pH11_NaOH_pur = oxydobase_pH11_rate_NaOH_H2S * H2S_flux_gh / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
            oxydobase_pH11_NaOH_pur = oxydobase_pH11_NaOH_pur + oxydobase_pH11_rate_NaOH_RSH * RSH_flux_gh / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
            'consommation due au CO2
            oxydobase_pH11_NaOH_pur = oxydobase_pH11_NaOH_pur + oxydobase_pH11_rate_NaOH_CO2 * (air_vicie_Q_Nm3h / 1000) / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
        End If
        '        If oxydobase_pH11_NaOH_commercial_force = True Then
        '            oxydobase_pH11_NaOH_commercial = oxydobase_pH11_NaOH_commercial_f
        '        Else
        '            oxydobase_pH11_NaOH_commercial = oxydobase_pH11_NaOH_pur / (NaOH_concentration / 1000)
        '        End If
        'NaOCl
        If oxydobase_Cl2_pur_force = True Then
            oxydobase_Cl2_pur = oxydobase_Cl2_pur_f
        Else
            oxydobase_Cl2_pur = oxydobase_Cl2_pur + oxydobase_rate_Cl2_H2S * H2S_flux_gh / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
            oxydobase_Cl2_pur = oxydobase_Cl2_pur + oxydobase_rate_Cl2_RSH * RSH_flux_gh / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
        End If
        '        If oxydobase_NaOCl_commercial_force = True Then
        '            oxydobase_NaOCl_commercial = oxydobase_NaOCl_commercial_f
        '        Else
        '            oxydobase_NaOCl_commercial = oxydobase_NaOCl_commercial + oxydobase_Cl2_pur / (concentration_Cl2_javel / 1000)
        '        End If
        'rendements de l'ordre de 99% voire plus donc --> 100%
        H2S_flux_gh = 0
        RSH_flux_gh = 0
    End If

    'tour neutre
    If choix_tour_neutre = choix_tour_oui Then
        If neutre_NaHSO3_pur_force = True Then
            neutre_NaHSO3_pur = neutre_NaHSO3_pur_f
        Else
            neutre_NaHSO3_pur = neutre_rate_NaHSO3_Cl2 * residuel_Cl2_mg_Nm3 * (air_vicie_Q_Nm3h / 1000) / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
        End If
        '        If neutre_NaHSO3_commercial_force = True Then
        '            neutre_NaHSO3_commercial = neutre_NaHSO3_commercial_f
        '        Else
        '            neutre_NaHSO3_commercial = neutre_NaHSO3_pur / (NaHSO3_titre * masse_volumique_NaHSO3 / 1000)
        '        End If
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "29.12"

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
    Q_lavage_unitaire = pompe_lavage_rate_Leau_m3air / 1000 * air_vicie_Q_Nm3h * nb_tours_lavage
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
    NumeroErreurNonGeree = "29.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil29.Range("desodo_chim_air_vicie_Q_Nm3h").Cells(1, Colonne) = air_vicie_Q_Nm3h
    Feuil29.Range("desodo_chim_air_vicie_T").Cells(1, Colonne) = air_vicie_T
    Feuil29.Range("desodo_chim_ventilation_tps_fonctionnement").Cells(1, Colonne) = ventilation_tps_fonctionnement
    Feuil29.Range("desodo_chim_ventilation_P_reseau").Cells(1, Colonne) = ventilation_P_reseau
    Feuil29.Range("desodo_chim_ventilation_rdt_global").Cells(1, Colonne) = ventilation_rdt_global
    Feuil29.Range("desodo_chim_pompe_lavage_rate_Leau_m3air").Cells(1, Colonne) = pompe_lavage_rate_Leau_m3air
    Feuil29.Range("desodo_chim_pompe_lavage_HMT").Cells(1, Colonne) = pompe_lavage_HMT
    Feuil29.Range("desodo_chim_pompe_lavage_rdt").Cells(1, Colonne) = pompe_lavage_rdt
    Feuil29.Range("desodo_chim_soufflage_P_reseau").Cells(1, Colonne) = soufflage_P_reseau
    Feuil29.Range("desodo_chim_soufflage_Tair").Cells(1, Colonne) = soufflage_Tair
    Feuil29.Range("desodo_chim_soufflage_rdt_global").Cells(1, Colonne) = soufflage_rdt_global
    Feuil29.Range("desodo_chim_air_vicie_NH3").Cells(1, Colonne) = air_vicie_NH3
    Feuil29.Range("desodo_chim_air_vicie_Norga").Cells(1, Colonne) = air_vicie_Norga
    Feuil29.Range("desodo_chim_air_vicie_H2S").Cells(1, Colonne) = air_vicie_H2S
    Feuil29.Range("desodo_chim_air_vicie_RSH").Cells(1, Colonne) = air_vicie_RSH
    '    Feuil29.Range("desodo_chim_H2SO4_titre").Cells(1, colonne) = H2SO4_titre
    Feuil29.Range("desodo_chim_acide_H2SO4_pur").Cells(1, Colonne) = acide_H2SO4_pur
    '    Feuil29.Range("desodo_chim_acide_H2SO4_commercial").Cells(1, colonne) = acide_H2SO4_commercial
    '    Feuil29.Range("desodo_chim_NaOH_titre").Cells(1, colonne) = NaOH_titre
    Feuil29.Range("desodo_chim_oxydobase_pH9_NaOH_pur").Cells(1, Colonne) = oxydobase_pH9_NaOH_pur
    '    Feuil29.Range("desodo_chim_oxydobase_pH9_NaOH_commercial").Cells(1, colonne) = oxydobase_pH9_NaOH_commercial
    Feuil29.Range("desodo_chim_oxydobase_pH11_NaOH_pur").Cells(1, Colonne) = oxydobase_pH11_NaOH_pur
    '    Feuil29.Range("desodo_chim_oxydobase_pH11_NaOH_commercial").Cells(1, colonne) = oxydobase_pH11_NaOH_commercial
    '    Feuil29.Range("desodo_chim_NaOCl_titre").Cells(1, colonne) = NaOCl_titre
    Feuil29.Range("desodo_chim_oxydobase_Cl2_pur").Cells(1, Colonne) = oxydobase_Cl2_pur
    '    Feuil29.Range("desodo_chim_oxydobase_NaOCl_commercial").Cells(1, colonne) = oxydobase_NaOCl_commercial
    '    Feuil29.Range("desodo_chim_NaHSO3_titre").Cells(1, colonne) = NaHSO3_titre
    Feuil29.Range("desodo_chim_neutre_NaHSO3_pur").Cells(1, Colonne) = neutre_NaHSO3_pur
    '    Feuil29.Range("desodo_chim_neutre_NaHSO3_commercial").Cells(1, colonne) = neutre_NaHSO3_commercial
    Feuil29.Range("desodo_chim_electricite").Cells(1, Colonne) = electricite



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "29.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil29.Range("desodo_chim_air_vicie_Q_Nm3h").Cells(1, Colonne) = ""    ' air_vicie_Q_Nm3h
    Feuil29.Range("desodo_chim_air_vicie_T").Cells(1, Colonne) = ""    ' air_vicie_T
    Feuil29.Range("desodo_chim_ventilation_tps_fonctionnement").Cells(1, Colonne) = ""    ' ventilation_tps_fonctionnement
    Feuil29.Range("desodo_chim_ventilation_P_reseau").Cells(1, Colonne) = ""    ' ventilation_P_reseau
    Feuil29.Range("desodo_chim_ventilation_rdt_global").Cells(1, Colonne) = ""   ' ventilation_rdt_global
    Feuil29.Range("desodo_chim_pompe_lavage_rate_Leau_m3air").Cells(1, Colonne) = ""   ' pompe_lavage_rate_Leau_m3air
    Feuil29.Range("desodo_chim_pompe_lavage_HMT").Cells(1, Colonne) = ""    ' pompe_lavage_HMT
    Feuil29.Range("desodo_chim_pompe_lavage_rdt").Cells(1, Colonne) = ""    ' pompe_lavage_rdt
    Feuil29.Range("desodo_chim_soufflage_P_reseau").Cells(1, Colonne) = ""    ' soufflage_P_reseau
    Feuil29.Range("desodo_chim_soufflage_Tair").Cells(1, Colonne) = ""    ' soufflage_Tair
    Feuil29.Range("desodo_chim_soufflage_rdt_global").Cells(1, Colonne) = ""    ' soufflage_rdt_global
    Feuil29.Range("desodo_chim_air_vicie_NH3").Cells(1, Colonne) = ""    ' air_vicie_NH3
    Feuil29.Range("desodo_chim_air_vicie_Norga").Cells(1, Colonne) = ""    ' air_vicie_Norga
    Feuil29.Range("desodo_chim_air_vicie_H2S").Cells(1, Colonne) = ""    ' air_vicie_H2S
    Feuil29.Range("desodo_chim_air_vicie_RSH").Cells(1, Colonne) = ""    ' air_vicie_RSH
    Feuil29.Range("desodo_chim_H2SO4_titre").Cells(1, Colonne) = ""    ' H2SO4_titre
    Feuil29.Range("desodo_chim_acide_H2SO4_pur").Cells(1, Colonne) = ""    ' acide_H2SO4_pur
    Feuil29.Range("desodo_chim_acide_H2SO4_commercial").Cells(1, Colonne) = ""    ' acide_H2SO4_commercial
    Feuil29.Range("desodo_chim_NaOH_titre").Cells(1, Colonne) = ""    ' NaOH_titre
    Feuil29.Range("desodo_chim_oxydobase_pH9_NaOH_pur").Cells(1, Colonne) = ""    ' oxydobase_pH9_NaOH_pur
    Feuil29.Range("desodo_chim_oxydobase_pH9_NaOH_commercial").Cells(1, Colonne) = ""    ' oxydobase_pH9_NaOH_commercial
    Feuil29.Range("desodo_chim_oxydobase_pH11_NaOH_pur").Cells(1, Colonne) = ""    ' oxydobase_pH11_NaOH_pur
    Feuil29.Range("desodo_chim_oxydobase_pH11_NaOH_commercial").Cells(1, Colonne) = ""    ' oxydobase_pH11_NaOH_commercial
    Feuil29.Range("desodo_chim_NaOCl_titre").Cells(1, Colonne) = ""    ' NaOCl_titre
    Feuil29.Range("desodo_chim_oxydobase_Cl2_pur").Cells(1, Colonne) = ""    ' oxydobase_Cl2_pur
    Feuil29.Range("desodo_chim_oxydobase_NaOCl_commercial").Cells(1, Colonne) = ""    ' oxydobase_NaOCl_commercial
    Feuil29.Range("desodo_chim_NaHSO3_titre").Cells(1, Colonne) = ""    ' NaHSO3_titre
    Feuil29.Range("desodo_chim_neutre_NaHSO3_pur").Cells(1, Colonne) = ""    ' neutre_NaHSO3_pur
    Feuil29.Range("desodo_chim_neutre_NaHSO3_commercial").Cells(1, Colonne) = ""    ' neutre_NaHSO3_commercial
    Feuil29.Range("desodo_chim_electricite").Cells(1, Colonne) = ""    ' electricite



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "29.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 3  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil29.Range("desodo_chim_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "29.8"

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

    'calcul de la densité de H2SO4 en fonction du titre massique   =somme(a(i)*x^i)
    a_densite_H2SO4(0) = 0.9998
    a_densite_H2SO4(1) = 0.6021
    a_densite_H2SO4(2) = 0.6771
    a_densite_H2SO4(3) = -1.072
    a_densite_H2SO4(4) = 0.1613
    a_densite_H2SO4(5) = 2.645
    a_densite_H2SO4(6) = -2.185

    'taux de réactifs par composé éliminé   en g/g
    acide_rate_H2SO4_NH3 = 2.9
    acide_rate_H2SO4_Norga = 3.5
    oxydobase_pH9_rate_NaOH_H2S = 2.5
    oxydobase_pH9_rate_NaOH_RSH = 4.1
    oxydobase_pH11_rate_NaOH_H2S = 3
    oxydobase_pH11_rate_NaOH_RSH = 3

    'consommation de réactif dûe au CO2  (mg/m3)
    oxydobase_pH9_rate_NaOH_CO2 = 30
    oxydobase_pH11_rate_NaOH_CO2 = 100

    'concentration_Cl2_javel -->  se reporter à valeur guide par défaut

    'consommation de javel    (g/g)
    oxydobase_rate_Cl2_H2S = 9.5
    oxydobase_rate_Cl2_RSH = 11

    'densite_NaHSO3 -->  se reporter à valeur guide par défaut

    'consommation de bisulfite de sodium   en g/g
    neutre_rate_NaHSO3_Cl2 = 1.85

    'résiduel Cl2 après passage dans tours oxydo-basique
    residuel_Cl2_mg_Nm3 = 1


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "29.3"

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
        ventilation_P_reseau = 800 * nb_tours_lavage
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
        pompe_lavage_rate_Leau_m3air = 3
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
        soufflage_P_reseau = 300    'Pa = 30 mmCE
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

    'composition de l'air vicié à traiter      teneurs en mg/Nm3
    If choix_couverture_step = couverture_partielle Then
        'hypothèse: air moyennement pollué
        If air_vicie_NH3_force = True Then
            air_vicie_NH3 = air_vicie_NH3_f
        Else
            air_vicie_NH3 = 8
        End If
        If air_vicie_Norga_force = True Then
            air_vicie_Norga = air_vicie_Norga_f
        Else
            air_vicie_Norga = 1
        End If
        If air_vicie_H2S_force = True Then
            air_vicie_H2S = air_vicie_H2S_f
        Else
            air_vicie_H2S = 5
        End If
        If air_vicie_RSH_force = True Then
            air_vicie_RSH = air_vicie_RSH_f
        Else
            air_vicie_RSH = 1.5
        End If
    Else
        'hypothèse: air peu pollué
        If air_vicie_NH3_force = True Then
            air_vicie_NH3 = air_vicie_NH3_f
        Else
            air_vicie_NH3 = 5
        End If
        If air_vicie_Norga_force = True Then
            air_vicie_Norga = air_vicie_Norga_f
        Else
            air_vicie_Norga = 0.5
        End If
        If air_vicie_H2S_force = True Then
            air_vicie_H2S = air_vicie_H2S_f
        Else
            air_vicie_H2S = 2
        End If
        If air_vicie_RSH_force = True Then
            air_vicie_RSH = air_vicie_RSH_f
        Else
            air_vicie_RSH = 0.5
        End If
    End If

    'titre en H2SO4  % massique  à vérifier
    '    If H2SO4_titre_force = True Then
    '        H2SO4_titre = H2SO4_titre_f
    '        masse_volumique_H2SO4 = 0
    '        For i = 0 To 6
    '            masse_volumique_H2SO4 = masse_volumique_H2SO4 + a_densite_H2SO4(i) * H2SO4_titre ^ i
    '        Next i
    '        masse_volumique_H2SO4 = masse_volumique_H2SO4 * 1000
    '    Else
    '        H2SO4_titre = 0.96
    '        masse_volumique_H2SO4 = 0
    '        For i = 0 To 6
    '            masse_volumique_H2SO4 = masse_volumique_H2SO4 + a_densite_H2SO4(i) * H2SO4_titre ^ i
    '        Next i
    '        masse_volumique_H2SO4 = masse_volumique_H2SO4 * 1000
    '    End If

    'titre en NaOH
    '    If NaOH_titre_force = True Then
    '        NaOH_titre = NaOH_titre_f
    '        NaOH_concentration = 19.71 * 100 * NaOH_titre - 222.86
    '    Else
    '        NaOH_titre = 0.3
    '        NaOH_concentration = 19.71 * 100 * NaOH_titre - 222.86
    '    End If

    'titre en NaOCl
    '    If NaOCl_titre_force = True Then
    '        NaOCl_titre = NaOCl_titre_f
    '        concentration_Cl2_javel = 3.17 * NaOCl_titre
    '    Else
    '        NaOCl_titre = 48
    '        concentration_Cl2_javel = 3.17 * NaOCl_titre
    '    End If

    'titre en NaHSO3
    '    If NaHSO3_titre_force = True Then
    '        NaHSO3_titre = NaHSO3_titre_f
    '        masse_volumique_NaHSO3 = (1 + 1.32 * NaHSO3_titre) * 1000
    '    Else
    '        NaHSO3_titre = 0.25
    '        masse_volumique_NaHSO3 = (1 + 1.32 * NaHSO3_titre) * 1000
    '    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "29.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub



Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String
    Dim message_erreur As String

    choix_couverture_step = Feuil29.Range("desodo_chim_choix_couvertureSTEP").Cells(1, pi_FiliereConsideree)
    choix_soufflage_air = Feuil29.Range("desodo_chim_choix_soufflage").Cells(1, pi_FiliereConsideree)

    nb_tours_lavage = 0
    choix_tour_acide = Feuil29.Range("desodo_chim_choix_tour_acide").Cells(1, pi_FiliereConsideree)
    If choix_tour_acide = choix_tour_oui Then
        nb_tours_lavage = nb_tours_lavage + 1
    End If
    choix_tour_oxydobase_pH9 = Feuil29.Range("desodo_chim_choix_tour_oxydobase_pH9").Cells(1, pi_FiliereConsideree)
    If choix_tour_oxydobase_pH9 = choix_tour_oui Then
        nb_tours_lavage = nb_tours_lavage + 1
    End If
    choix_tour_oxydobase_pH11 = Feuil29.Range("desodo_chim_choix_tour_oxydobase_pH11").Cells(1, pi_FiliereConsideree)
    If choix_tour_oxydobase_pH11 = choix_tour_oui Then
        nb_tours_lavage = nb_tours_lavage + 1
    End If
    choix_tour_neutre = Feuil29.Range("desodo_chim_choix_tour_neutre").Cells(1, pi_FiliereConsideree)
    If choix_tour_neutre = choix_tour_oui Then
        nb_tours_lavage = nb_tours_lavage + 1
    End If
    
    'JPO 22/05/2014
    If nb_tours_lavage = 0 Then
        message_erreur = Feuil29.Range("erreur_desodorisation_chimique").Cells(1, 1) & " " & Feuil29.Range("erreur_selection_tours").Cells(1, 1)
        MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub 'met fin aux calculs
    End If
  
    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "29.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    air_vicie_Q_Nm3h_force = Feuil29.Range("desodo_chim_air_vicie_Q_Nm3h").Cells(1, pi_FiliereConsideree)
    air_vicie_T_force = Feuil29.Range("desodo_chim_air_vicie_T").Cells(1, pi_FiliereConsideree)
    ventilation_tps_fonctionnement_force = Feuil29.Range("desodo_chim_ventilation_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    ventilation_P_reseau_force = Feuil29.Range("desodo_chim_ventilation_P_reseau").Cells(1, pi_FiliereConsideree)
    ventilation_rdt_global_force = Feuil29.Range("desodo_chim_ventilation_rdt_global").Cells(1, pi_FiliereConsideree)
    pompe_lavage_rate_Leau_m3air_force = Feuil29.Range("desodo_chim_pompe_lavage_rate_Leau_m3air").Cells(1, pi_FiliereConsideree)
    pompe_lavage_HMT_force = Feuil29.Range("desodo_chim_pompe_lavage_HMT").Cells(1, pi_FiliereConsideree)
    pompe_lavage_rdt_force = Feuil29.Range("desodo_chim_pompe_lavage_rdt").Cells(1, pi_FiliereConsideree)
    soufflage_P_reseau_force = Feuil29.Range("desodo_chim_soufflage_P_reseau").Cells(1, pi_FiliereConsideree)
    soufflage_Tair_force = Feuil29.Range("desodo_chim_soufflage_Tair").Cells(1, pi_FiliereConsideree)
    soufflage_rdt_global_force = Feuil29.Range("desodo_chim_soufflage_rdt_global").Cells(1, pi_FiliereConsideree)
    air_vicie_NH3_force = Feuil29.Range("desodo_chim_air_vicie_NH3").Cells(1, pi_FiliereConsideree)
    air_vicie_Norga_force = Feuil29.Range("desodo_chim_air_vicie_Norga").Cells(1, pi_FiliereConsideree)
    air_vicie_H2S_force = Feuil29.Range("desodo_chim_air_vicie_H2S").Cells(1, pi_FiliereConsideree)
    air_vicie_RSH_force = Feuil29.Range("desodo_chim_air_vicie_RSH").Cells(1, pi_FiliereConsideree)
    '    H2SO4_titre_force = Feuil29.Range("desodo_chim_H2SO4_titre").Cells(1, pi_FiliereConsideree)
    acide_H2SO4_pur_force = Feuil29.Range("desodo_chim_acide_H2SO4_pur").Cells(1, pi_FiliereConsideree)
    '    acide_H2SO4_commercial_force = Feuil29.Range("desodo_chim_acide_H2SO4_commercial").Cells(1, pi_FiliereConsideree)
    '    NaOH_titre_force = Feuil29.Range("desodo_chim_NaOH_titre").Cells(1, pi_FiliereConsideree)
    oxydobase_pH9_NaOH_pur_force = Feuil29.Range("desodo_chim_oxydobase_pH9_NaOH_pur").Cells(1, pi_FiliereConsideree)
    '    oxydobase_pH9_NaOH_commercial_force = Feuil29.Range("desodo_chim_oxydobase_pH9_NaOH_commercial").Cells(1, pi_FiliereConsideree)
    oxydobase_pH11_NaOH_pur_force = Feuil29.Range("desodo_chim_oxydobase_pH11_NaOH_pur").Cells(1, pi_FiliereConsideree)
    '    oxydobase_pH11_NaOH_commercial_force = Feuil29.Range("desodo_chim_oxydobase_pH11_NaOH_commercial").Cells(1, pi_FiliereConsideree)
    '    NaOCl_titre_force = Feuil29.Range("desodo_chim_NaOCl_titre").Cells(1, pi_FiliereConsideree)
    oxydobase_Cl2_pur_force = Feuil29.Range("desodo_chim_oxydobase_Cl2_pur").Cells(1, pi_FiliereConsideree)
    '    oxydobase_NaOCl_commercial_force = Feuil29.Range("desodo_chim_oxydobase_NaOCl_commercial").Cells(1, pi_FiliereConsideree)
    '    NaHSO3_titre_force = Feuil29.Range("desodo_chim_NaHSO3_titre").Cells(1, pi_FiliereConsideree)
    neutre_NaHSO3_pur_force = Feuil29.Range("desodo_chim_neutre_NaHSO3_pur").Cells(1, pi_FiliereConsideree)
    '    neutre_NaHSO3_commercial_force = Feuil29.Range("desodo_chim_neutre_NaHSO3_commercial").Cells(1, pi_FiliereConsideree)


    If air_vicie_Q_Nm3h_force = True Then
        air_vicie_Q_Nm3h_f = Feuil29.Range("desodo_chim_air_vicie_Q_Nm3h").Cells(1, pi_FiliereConsideree)
    End If
    If air_vicie_T_force = True Then
        air_vicie_T_f = Feuil29.Range("desodo_chim_air_vicie_T").Cells(1, pi_FiliereConsideree)
    End If
    If ventilation_tps_fonctionnement_force = True Then
        ventilation_tps_fonctionnement_f = Feuil29.Range("desodo_chim_ventilation_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    End If
    If ventilation_P_reseau_force = True Then
        ventilation_P_reseau_f = Feuil29.Range("desodo_chim_ventilation_P_reseau").Cells(1, pi_FiliereConsideree)
    End If
    If ventilation_rdt_global_force = True Then
        ventilation_rdt_global_f = Feuil29.Range("desodo_chim_ventilation_rdt_global").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_lavage_rate_Leau_m3air_force = True Then
        pompe_lavage_rate_Leau_m3air_f = Feuil29.Range("desodo_chim_pompe_lavage_rate_Leau_m3air").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_lavage_HMT_force = True Then
        pompe_lavage_HMT_f = Feuil29.Range("desodo_chim_pompe_lavage_HMT").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_lavage_rdt_force = True Then
        pompe_lavage_rdt_f = Feuil29.Range("desodo_chim_pompe_lavage_rdt").Cells(1, pi_FiliereConsideree)
    End If
    If soufflage_P_reseau_force = True Then
        soufflage_P_reseau_f = Feuil29.Range("desodo_chim_soufflage_P_reseau").Cells(1, pi_FiliereConsideree)
    End If
    If soufflage_Tair_force = True Then
        soufflage_Tair_f = Feuil29.Range("desodo_chim_soufflage_Tair").Cells(1, pi_FiliereConsideree)
    End If
    If soufflage_rdt_global_force = True Then
        soufflage_rdt_global_f = Feuil29.Range("desodo_chim_soufflage_rdt_global").Cells(1, pi_FiliereConsideree)
    End If
    If air_vicie_NH3_force = True Then
        air_vicie_NH3_f = Feuil29.Range("desodo_chim_air_vicie_NH3").Cells(1, pi_FiliereConsideree)
    End If
    If air_vicie_Norga_force = True Then
        air_vicie_Norga_f = Feuil29.Range("desodo_chim_air_vicie_Norga").Cells(1, pi_FiliereConsideree)
    End If
    If air_vicie_H2S_force = True Then
        air_vicie_H2S_f = Feuil29.Range("desodo_chim_air_vicie_H2S").Cells(1, pi_FiliereConsideree)
    End If
    If air_vicie_RSH_force = True Then
        air_vicie_RSH_f = Feuil29.Range("desodo_chim_air_vicie_RSH").Cells(1, pi_FiliereConsideree)
    End If
    '    If H2SO4_titre_force = True Then
    '        H2SO4_titre_f = Feuil29.Range("desodo_chim_H2SO4_titre").Cells(1, pi_FiliereConsideree)
    '    End If
    If acide_H2SO4_pur_force = True Then
        acide_H2SO4_pur_f = Feuil29.Range("desodo_chim_acide_H2SO4_pur").Cells(1, pi_FiliereConsideree)
    End If
    '    If acide_H2SO4_commercial_force = True Then
    '        acide_H2SO4_commercial_f = Feuil29.Range("desodo_chim_acide_H2SO4_commercial").Cells(1, pi_FiliereConsideree)
    '    End If
    '    If NaOH_titre_force = True Then
    '        NaOH_titre_f = Feuil29.Range("desodo_chim_NaOH_titre").Cells(1, pi_FiliereConsideree)
    '    End If
    If oxydobase_pH9_NaOH_pur_force = True Then
        oxydobase_pH9_NaOH_pur_f = Feuil29.Range("desodo_chim_oxydobase_pH9_NaOH_pur").Cells(1, pi_FiliereConsideree)
    End If
    '    If oxydobase_pH9_NaOH_commercial_force = True Then
    '        oxydobase_pH9_NaOH_commercial_f = Feuil29.Range("desodo_chim_oxydobase_pH9_NaOH_commercial").Cells(1, pi_FiliereConsideree)
    '    End If
    If oxydobase_pH11_NaOH_pur_force = True Then
        oxydobase_pH11_NaOH_pur_f = Feuil29.Range("desodo_chim_oxydobase_pH11_NaOH_pur").Cells(1, pi_FiliereConsideree)
    End If
    '    If oxydobase_pH11_NaOH_commercial_force = True Then
    '        oxydobase_pH11_NaOH_commercial_f = Feuil29.Range("desodo_chim_oxydobase_pH11_NaOH_commercial").Cells(1, pi_FiliereConsideree)
    '    End If
    '    If NaOCl_titre_force = True Then
    '        NaOCl_titre_f = Feuil29.Range("desodo_chim_NaOCl_titre").Cells(1, pi_FiliereConsideree)
    '    End If
    If oxydobase_Cl2_pur_force = True Then
        oxydobase_Cl2_pur_f = Feuil29.Range("desodo_chim_oxydobase_Cl2_pur").Cells(1, pi_FiliereConsideree)
    End If
    '    If oxydobase_NaOCl_commercial_force = True Then
    '        oxydobase_NaOCl_commercial_f = Feuil29.Range("desodo_chim_oxydobase_NaOCl_commercial").Cells(1, pi_FiliereConsideree)
    '    End If
    '    If NaHSO3_titre_force = True Then
    '        NaHSO3_titre_f = Feuil29.Range("desodo_chim_NaHSO3_titre").Cells(1, pi_FiliereConsideree)
    '    End If
    If neutre_NaHSO3_pur_force = True Then
        neutre_NaHSO3_pur_f = Feuil29.Range("desodo_chim_neutre_NaHSO3_pur").Cells(1, pi_FiliereConsideree)
    End If
    '    If neutre_NaHSO3_commercial_force = True Then
    '        neutre_NaHSO3_commercial_f = Feuil29.Range("desodo_chim_neutre_NaHSO3_commercial").Cells(1, pi_FiliereConsideree)
    '    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "29.2"

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






