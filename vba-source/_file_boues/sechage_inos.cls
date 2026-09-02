Attribute VB_Name = "sechage_inos"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'constantes pour le choix d'une bâche des boues à sécher
Private Const bache_amont_non = 1
Private Const bache_amont_1 = 2
Private Const bache_amont_2 = 3

Private Const nb_inlet = 2

Private Const critere_incoherence = 0.1

'constantes pour les choix de Siccité intermédiaire après déshydratation
Private Const conditionnement_polymere = 1
Private Const conditionnement_polymere_FeCl3 = 2
Private Const conditionnement_chaux_FeCl3 = 3
Dim choix_conditionnement As Integer

'mode de couverture des besoins thermiques
Private Const source_eau_HT = 1   'si basse température uniquement
Private Const source_vapeur = 2
Private Const source_biogaz = 3
Private Const source_combustible = 4

'technologie de condensation de l'eau évaporée
Private Const condensation_indirecte = 1
Private Const condensation_directe = 2
Private Const condensation_indirecte_directe = 3   'd'abord indirecte puis directe

'constantes pour les choix de boues en entrée
Private Const inlet_aucun = 0
Private Const inlet_boues_I = 1
Private Const inlet_boues_II = 2
Private Const inlet_boues_III = 3
Private Const inlet_boues_externes_1 = 4
Private Const inlet_boues_externes_2 = 5
Private Const inlet_graisses = 6
Private Const inlet_epaississeur_1 = 7
Private Const inlet_epaississeur_2 = 8
Private Const inlet_epaississeur_3 = 9
Private Const inlet_epaississeur_4 = 10
Private Const inlet_epaississeur_5 = 11
Private Const inlet_digestion = 12
Dim ratio_boues_digerees As Double

'variables de choix
'Dim choix_performance As Integer
Dim choix_inlet(nb_inlet) As Integer
Dim choix_bache_amont As Integer
Dim intermediaire_origine As Integer
Dim choix_source_energie As Integer
Dim choix_condensation As Integer

Public choix_polymere_type As Integer

'entrée
Dim inlet_siccite As Double
Dim inlet_concentration As Double
Public polymere_flux As Double
Public chaux_eteinte_flux As Double
Public FeCl3_flux As Double
Dim desh_siccite As Double
Dim chaux As Double
Dim polymere_kg_tMS As Double
Dim FeCl3_kg_tMS As Double
Public inlet_MES As Double
Dim inlet_Q As Double
Dim capture_MS As Double
Dim inlet_MV_MES As Double
'guide defaut
Dim outlet_siccite As Double
'hypothese
Dim ratio_elec As Double
Dim ratio_chal As Double
Dim masse_volumique_boues As Double
'calcul
Dim inlet_BH As Double
Dim outlet_MES As Double
Dim outlet_MV_MES As Double
Dim outlet_BH As Double
Dim eau_evaporee As Double
Dim besoins_thermiques As Double
Dim conso_spec As Double
Dim intermediaire_concentration As Double
Dim ratio_boues_bio As Double

'variables relatives à l'énergie thermique
Dim condensat_indirect_ratio_vapeur_condensee As Double
Dim condensat_indirect_recup_eau_HT As Double     'énergie récupérée de type eau à haute température
Dim condensat_indirect_ratio_vapeur_condensee_force As Boolean
Dim condensat_indirect_recup_eau_HT_force As Boolean     'énergie récupérée de type eau à haute température
Dim condensat_indirect_ratio_vapeur_condensee_f As Double
Dim condensat_indirect_recup_eau_HT_f As Double     'énergie récupérée de type eau à haute température

'variables à échanger avec l'interface
Dim inlet_ratio_admis(nb_inlet) As Double
Dim alimentation_pompe_nb As Integer
Dim alimentation_P_refoulement As Double
Dim alimentation_tps_fonctionnement As Double
Dim alimentation_Q_unitaire As Double
Dim alimentation_rendement_pompe As Double
Dim alimentation_rendement_global As Double
Dim alimentation2_pompe_nb As Integer
Dim alimentation2_P_refoulement As Double
Dim alimentation2_tps_fonctionnement As Double
Dim alimentation2_Q_unitaire As Double
Dim alimentation2_rendement_pompe As Double
Dim alimentation2_rendement_global As Double
Public electricite As Double

'déclaration des variables relatives aux valeurs renseignées (cases bleues)
'entrée

Dim inlet_siccite_force As Boolean
Dim inlet_MES_force As Boolean
Dim inlet_MV_MES_force As Boolean
Dim outlet_siccite_force As Boolean
Dim capture_MS_force As Boolean

'calcul
Dim inlet_BH_force As Boolean
Dim outlet_MES_force As Boolean
Dim outlet_MV_MES_force As Boolean
Dim outlet_BH_force As Boolean
Dim eau_evaporee_force As Boolean
Dim electricite_specifique_force As Boolean
Dim besoins_thermiques_force As Boolean
'variables à échanger avec l'interface
Dim alimentation_pompe_nb_force As Boolean
Dim alimentation_P_refoulement_force As Boolean
Dim alimentation_tps_fonctionnement_force As Boolean
Dim alimentation_Q_unitaire_force As Boolean
Dim alimentation_rendement_global_force As Boolean
Dim alimentation2_pompe_nb_force As Boolean
Dim alimentation2_P_refoulement_force As Boolean
Dim alimentation2_tps_fonctionnement_force As Boolean
Dim alimentation2_Q_unitaire_force As Boolean
Dim alimentation2_rendement_global_force As Boolean
Dim intermediaire_concentration_force As Boolean

Dim chaux_force As Boolean
Dim polymere_kg_tMS_force As Boolean
Dim FeCl3_kg_tMS_force As Boolean

'guide defaut

Dim inlet_siccite_f As Double
Dim inlet_MES_f As Double
Dim inlet_MV_MES_f As Double
Dim capture_MS_f As Double
Dim outlet_siccite_f As Double
Dim outlet_siccite_c As Double
Dim desh_siccite_f As Double

'hypothese
Dim ratio_elec_f As Double
Dim ratio_chal_f As Double
Dim conso_specifique_type As Double
'calcul
Dim inlet_BH_f As Double
Dim outlet_MES_f As Double
Dim outlet_MV_MES_f As Double
Dim outlet_BH_f As Double
Dim eau_evaporee_f As Double
Dim electricite_specifique_f As Double
Dim besoins_thermiques_f As Double
Dim intermediaire_concentration_f As Double

Dim chaux_f As Double
Dim polymere_kg_tMS_f As Double
Dim FeCl3_kg_tMS_f As Double

'variables à échanger avec l'interface
Dim alimentation_pompe_nb_f As Integer
Dim alimentation_P_refoulement_f As Double
Dim alimentation_tps_fonctionnement_f As Double
Dim alimentation_Q_unitaire_f As Double
Dim alimentation_rendement_global_f As Double
Dim alimentation2_pompe_nb_f As Integer
Dim alimentation2_P_refoulement_f As Double
Dim alimentation2_tps_fonctionnement_f As Double
Dim alimentation2_Q_unitaire_f As Double
Dim alimentation2_rendement_global_f As Double

'variables intermédiaires
Dim inlet_etape_anterieure(nb_inlet) As Integer
Dim inlet_type_boues(nb_inlet) As Double

'hypothèses
Dim intermediaire_concentration_reference(nb_qualites_boues) As Double
Dim polymere_Q As Double
Dim polymere_concentration As Double

Dim deltaH_vap_kWh As Double
Dim condensat_direct_Teau_inlet As Double
Dim condensat_direct_Teau_outlet As Double
Dim Cp_eau_kWh As Double
Dim ratio_DCOret_MVin As Double
Dim ratio_NKret_NKin As Double
Dim ratio_Pret_Pin As Double
Dim condensat_direct_Tair_inlet As Double
Dim ratio_retour_DBO_DCO As Double
Dim hyp_tps_fct_pompe_alim As Double

Dim secheur_pas_necessaire As Boolean

'erreurs
Dim stockage_numero_sechage As Integer
Dim remarque(2) As String
Dim correction_concentration_boues_digerees As Double

'pour le PINCH
Dim chauffage_T_out As Double
Dim chauffage_T_out_force As Boolean
Dim chauffage_T_out_f As Double
Dim chauffage_T_in As Double
Dim chauffage_T_in_force As Boolean
Dim chauffage_T_in_f As Double
Dim condensation_indirecte_Teau_out As Double
Dim condensation_indirecte_Teau_out_force As Boolean
Dim condensation_indirecte_Teau_out_f As Double
Dim condensation_indirecte_Teau_in As Double
Dim condensation_indirecte_Teau_in_force As Boolean
Dim condensation_indirecte_Teau_in_f As Double


Sub lecture_choix(ByVal numero_sechage_inos)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer

    stockage_numero_sechage = numero_sechage_inos

    'Boues en entrée de séchage


    choix_inlet(1) = Feuil38.Range("sechage_inos_choix_inlet1").Cells(numero_sechage_inos, pi_FiliereConsideree)
    inlet_ratio_admis(1) = Feuil38.Range("sechage_inos_inlet1_ratio").Cells(numero_sechage_inos, pi_FiliereConsideree)

    choix_inlet(2) = Feuil38.Range("sechage_inos_choix_inlet2").Cells(numero_sechage_inos, pi_FiliereConsideree)
    inlet_ratio_admis(2) = Feuil38.Range("sechage_inos_inlet2_ratio").Cells(numero_sechage_inos, pi_FiliereConsideree)

    'on retranscrit le choix réalisé dans la synthèse pour l'adapter à notre matrice
    For I = 1 To nb_inlet
        If choix_inlet(I) <> inlet_aucun Then
            If choix_inlet(I) <= inlet_graisses Then
                inlet_etape_anterieure(I) = boues_inlet
                inlet_type_boues(I) = choix_inlet(I)
            Else
                'ici on prendra directement les valeurs du flux antérieur
                inlet_etape_anterieure(I) = choix_inlet(I) - inlet_graisses + 1
            End If
        End If
    Next I

    'lecture des choix internes
    choix_conditionnement = Feuil38.Range("sechage_inos_choix_type").Cells(numero_sechage_inos, pi_FiliereConsideree)
    choix_polymere_type = Feuil38.Range("sechage_inos_choix_polymere_type").Cells(numero_sechage_inos, pi_FiliereConsideree)
    choix_source_energie = Feuil38.Range("sechage_inos_choix_source_energie").Cells(numero_sechage_inos, pi_FiliereConsideree)
    choix_condensation = Feuil38.Range("sechage_inos_choix_condensation").Cells(numero_sechage_inos, pi_FiliereConsideree)
    choix_bache_amont = Feuil38.Range("sechage_inos_choix_bache_amont").Cells(numero_sechage_inos, pi_FiliereConsideree)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "38.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub hypotheses(ByVal numero_sechage_inos)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'ratios de consommation électrique et thermique

    ratio_chal = 1100       'kWh/ TEE
    outlet_siccite_c = 0.9      '%

    'concentration espérée en sortie   en g/L
    'primaire
    intermediaire_concentration_reference(I_simple) = 420
    intermediaire_concentration_reference(I_reactif) = 400
    'secondaire
    intermediaire_concentration_reference(II_forte) = 360
    intermediaire_concentration_reference(II_moyenne) = 290
    intermediaire_concentration_reference(II_faible_EB) = 260
    intermediaire_concentration_reference(II_faible_ED) = 260
    intermediaire_concentration_reference(II_prolongee_EB) = 260
    intermediaire_concentration_reference(II_prolongee_ED) = 260
    intermediaire_concentration_reference(II_MBR) = 240
    intermediaire_concentration_reference(II_MBBR) = 250
    intermediaire_concentration_reference(II_HybAS) = 250
    intermediaire_concentration_reference(II_biostyr_C) = 320
    intermediaire_concentration_reference(II_biostyr_N) = 300
    intermediaire_concentration_reference(II_biostyr_NDN) = 270
    intermediaire_concentration_reference(II_biostyr_PDN) = 240
    'tertiaire
    intermediaire_concentration_reference(III_decantation) = 240
    intermediaire_concentration_reference(III_biostyr_N) = 240
    intermediaire_concentration_reference(III_biostyr_PDN) = 240
    'graisses
    intermediaire_concentration_reference(codigestion_graisses) = 200

    correction_concentration_boues_digerees = 20


    'masse volumique des boues entrantes ~ eau
    masse_volumique_boues = 1000   'kg/m3

    'enthalpie de vaporisation de l'eau (hyporthèse T=105°C)
    deltaH_vap_kWh = 530 * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE    'kWh/kg   530 kcal/kg pour T~110-115°C

    'températures des différents courants au niveau de la condensation directe
    condensat_direct_Teau_inlet = 20    '°C
    condensat_direct_Teau_outlet = 40    '°C
    Cp_eau_kWh = 4.18 / pd_NOMBRE_SECONDE_PAR_HEURE    'kWh/(kg.°C)

    'part des pollutions que l'on retrouve dans les retours
    ratio_DCOret_MVin = 0.015
    ratio_NKret_NKin = 0.01   'partie des NK hors NH4 qui se retrouve dans les retours sous forme de NH4 --> le soluble + cette part qui titre en NH4 et aussi en NK
    ratio_Pret_Pin = 0
    ratio_retour_DBO_DCO = 0.3

    'température de l'air après séchage et avant condensation
    condensat_direct_Tair_inlet = 110   '°C


    hyp_tps_fct_pompe_alim = 5



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "38.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub
Public Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'ratio en %
    If outlet_siccite_force = True Then
        outlet_siccite = outlet_siccite_f
    Else
        outlet_siccite = outlet_siccite_c
    End If


    If capture_MS_force = True Then
        capture_MS = capture_MS_f
    Else
        capture_MS = 0.98
    End If

    'alimentation
    If choix_bache_amont <> bache_amont_non Then
        If alimentation_pompe_nb_force = True Then
            alimentation_pompe_nb = alimentation_pompe_nb_f
        Else
            alimentation_pompe_nb = 1
        End If
        If alimentation_P_refoulement_force = True Then
            alimentation_P_refoulement = alimentation_P_refoulement_f
        Else
            alimentation_P_refoulement = 20
        End If
        If alimentation_tps_fonctionnement_force = True Then
            alimentation_tps_fonctionnement = alimentation_tps_fonctionnement_f
        Else
            alimentation_tps_fonctionnement = hyp_tps_fct_pompe_alim
        End If

        alimentation_rendement_pompe = 0.4    'hypothèse

        If alimentation2_pompe_nb_force = True Then
            alimentation2_pompe_nb = alimentation2_pompe_nb_f
        Else
            alimentation2_pompe_nb = 1
        End If
        If alimentation2_P_refoulement_force = True Then
            alimentation2_P_refoulement = alimentation2_P_refoulement_f
        Else
            alimentation2_P_refoulement = 20
        End If
        If alimentation2_tps_fonctionnement_force = True Then
            alimentation2_tps_fonctionnement = alimentation2_tps_fonctionnement_f
        Else
            alimentation2_tps_fonctionnement = hyp_tps_fct_pompe_alim
        End If

        alimentation2_rendement_pompe = 0.4    'hypothèse

    End If

    'température du fluide de séchage au niveau de l'échangeur
    If chauffage_T_out_force = True Then
        chauffage_T_out = chauffage_T_out_f
    Else
        chauffage_T_out = 90  '°C
    End If
    If chauffage_T_in_force = True Then
        chauffage_T_in = chauffage_T_in_f
    Else
        chauffage_T_in = 60  '°C
    End If

    'condensation de l'eau évaporée
    If condensat_indirect_ratio_vapeur_condensee_force = True Then
        condensat_indirect_ratio_vapeur_condensee = condensat_indirect_ratio_vapeur_condensee_f
    Else
        condensat_indirect_ratio_vapeur_condensee = 0.5   '50% de l'eau évaporée condensée dans la partie indirecte
    End If
    If condensation_indirecte_Teau_in_force = True Then
        condensation_indirecte_Teau_in = condensation_indirecte_Teau_in_f
    Else
        condensation_indirecte_Teau_in = 60
    End If
    If condensation_indirecte_Teau_out_force = True Then
        condensation_indirecte_Teau_out = condensation_indirecte_Teau_out_f
    Else
        condensation_indirecte_Teau_out = 40
    End If


    If choix_conditionnement = conditionnement_polymere Then    'polymere
        chaux = 0                                  'kg/t MV
        If polymere_kg_tMS_force = True Then
            polymere_kg_tMS = polymere_kg_tMS_f
        Else
            polymere_kg_tMS = 2.86                            'kg/t MV
        End If
        FeCl3_kg_tMS = 0                                  'kg/t MV
    ElseIf choix_conditionnement = conditionnement_polymere_FeCl3 Then    'polymere_fecl3
        chaux = 0                                  'kg/t MV
        If polymere_kg_tMS_force = True Then
            polymere_kg_tMS = polymere_kg_tMS_f
        Else
            polymere_kg_tMS = 4.29                            'kg/t MV
        End If
        If FeCl3_kg_tMS_force = True Then
            FeCl3_kg_tMS = FeCl3_kg_tMS_f
        Else
            FeCl3_kg_tMS = 71.4                               'kg/t MV
        End If
    Else                   'chaux_fecl3    ' chaux éteinte (eric guibelin)
        If chaux_force = True Then
            chaux = chaux_f
        Else
            chaux = 286                                'kg/t MV
        End If
        polymere_kg_tMS = 0                               'kg/t MV
        If FeCl3_kg_tMS_force = True Then
            FeCl3_kg_tMS = FeCl3_kg_tMS_f
        Else
            FeCl3_kg_tMS = 57.1                               'kg/t MV
        End If
    End If




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "38.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub


Sub lecture_valeurs_forcees(ByVal numero_sechage_inos)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim column As Integer

    column = pi_FiliereConsideree

    inlet_siccite_force = Feuil38.Range("sechage_inos_inlet_siccite").Cells(numero_sechage_inos, column)
    inlet_MES_force = Feuil38.Range("sechage_inos_inlet_MES").Cells(numero_sechage_inos, column)
    inlet_MV_MES_force = Feuil38.Range("sechage_inos_inlet_MV_MES").Cells(numero_sechage_inos, column)
    inlet_BH_force = Feuil38.Range("sechage_inos_inlet_BH").Cells(numero_sechage_inos, column)

    outlet_MES_force = Feuil38.Range("sechage_inos_outlet_MES").Cells(numero_sechage_inos, column)
    outlet_MV_MES_force = Feuil38.Range("sechage_inos_outlet_MV_MES").Cells(numero_sechage_inos, column)
    outlet_BH_force = Feuil38.Range("sechage_inos_outlet_BH").Cells(numero_sechage_inos, column)
    outlet_siccite_force = Feuil38.Range("sechage_inos_outlet_siccite").Cells(numero_sechage_inos, column)
    capture_MS_force = Feuil38.Range("sechage_inos_capture_MS").Cells(numero_sechage_inos, column)

    chaux_force = Feuil38.Range("sechage_inos_chaux").Cells(numero_sechage_inos, column)
    polymere_kg_tMS_force = Feuil38.Range("sechage_inos_polymere").Cells(numero_sechage_inos, column)
    FeCl3_kg_tMS_force = Feuil38.Range("sechage_inos_fecl3").Cells(numero_sechage_inos, column)

    'electricite_specifique_force = Feuil38.Range("sechage_inos_electricite_specifique").Cells(numero_sechage_inos, column)
    besoins_thermiques_force = Feuil38.Range("sechage_inos_chaleur").Cells(numero_sechage_inos, column)
    intermediaire_concentration_force = Feuil38.Range("sechage_inos_outlet_concentration").Cells(numero_sechage_inos, column)

    alimentation_pompe_nb_force = Feuil38.Range("sechage_inos_alimentation_nb_pompe").Cells(numero_sechage_inos, column)
    alimentation_P_refoulement_force = Feuil38.Range("sechage_inos_alimentation_P_refoulement").Cells(numero_sechage_inos, column)
    alimentation_tps_fonctionnement_force = Feuil38.Range("sechage_inos_alimentation_tps_fonctionnement").Cells(numero_sechage_inos, column)
    alimentation_Q_unitaire_force = Feuil38.Range("sechage_inos_alimentation_Q_unitaire").Cells(numero_sechage_inos, column)
    alimentation_rendement_global_force = Feuil38.Range("sechage_inos_alimentation_rendement_pompe").Cells(numero_sechage_inos, column)

    alimentation2_pompe_nb_force = Feuil38.Range("sechage_inos_alimentation2_nb_pompe").Cells(numero_sechage_inos, column)
    alimentation2_P_refoulement_force = Feuil38.Range("sechage_inos_alimentation2_P_refoulement").Cells(numero_sechage_inos, column)
    alimentation2_tps_fonctionnement_force = Feuil38.Range("sechage_inos_alimentation2_tps_fonctionnement").Cells(numero_sechage_inos, column)
    alimentation2_Q_unitaire_force = Feuil38.Range("sechage_inos_alimentation2_Q_unitaire").Cells(numero_sechage_inos, column)
    alimentation2_rendement_global_force = Feuil38.Range("sechage_inos_alimentation2_rendement_pompe").Cells(numero_sechage_inos, column)

    chauffage_T_out_force = Feuil38.Range("sechage_inos_chauffage_T_out").Cells(numero_sechage_inos, column)
    chauffage_T_in_force = Feuil38.Range("sechage_inos_chauffage_T_in").Cells(numero_sechage_inos, column)


    If choix_condensation <> condensation_directe Then
        condensat_indirect_ratio_vapeur_condensee_force = Feuil38.Range("sechage_inos_condensat_indirect_ratio_vapeur_condensee").Cells(numero_sechage_inos, column)
        condensat_indirect_recup_eau_HT_force = Feuil38.Range("sechage_inos_condensat_indirect_recup_eau_HT").Cells(numero_sechage_inos, column)
        condensation_indirecte_Teau_in_force = Feuil38.Range("sechage_inos_condensation_indirecte_Teau_in").Cells(numero_sechage_inos, column)
        condensation_indirecte_Teau_out_force = Feuil38.Range("sechage_inos_condensation_indirecte_Teau_out").Cells(numero_sechage_inos, column)
    End If

    If inlet_MES_force = True Then
        inlet_MES_f = Feuil38.Range("sechage_inos_inlet_MES").Cells(numero_sechage_inos, column)
    End If

    If inlet_siccite_force = True Then
        inlet_siccite_f = Feuil38.Range("sechage_inos_inlet_siccite").Cells(numero_sechage_inos, column)
    End If

    If inlet_MV_MES_force = True Then
        inlet_MV_MES_f = Feuil38.Range("sechage_inos_inlet_MV_MES").Cells(numero_sechage_inos, column)
    End If

    If inlet_BH_force = True Then
        inlet_BH_f = Feuil38.Range("sechage_inos_inlet_BH").Cells(numero_sechage_inos, column)
    End If

    If outlet_MES_force = True Then
        outlet_MES_f = Feuil38.Range("sechage_inos_outlet_MES").Cells(numero_sechage_inos, column)
    End If

    If outlet_MV_MES_force = True Then
        outlet_MV_MES_f = Feuil38.Range("sechage_inos_outlet_MV_MES").Cells(numero_sechage_inos, column)
    End If

    If outlet_BH_force = True Then
        outlet_BH_f = Feuil38.Range("sechage_inos_outlet_BH").Cells(numero_sechage_inos, column)
    End If

    If intermediaire_concentration_force = True Then
        intermediaire_concentration_f = Feuil38.Range("sechage_inos_outlet_concentration").Cells(numero_sechage_inos, column)
    End If

    If outlet_siccite_force = True Then
        outlet_siccite_f = Feuil38.Range("sechage_inos_outlet_siccite").Cells(numero_sechage_inos, column)
    End If

    'If electricite_specifique_force = True Then
    '    electricite_specifique_f = Feuil38.Range("sechage_inos_electricite_specifique").Cells(numero_sechage_inos, column)
    'End If
    If chaux_force = True Then
        chaux_f = Feuil38.Range("sechage_inos_chaux").Cells(numero_sechage_inos, column)
    End If
    If polymere_kg_tMS_force = True Then
        polymere_kg_tMS_f = Feuil38.Range("sechage_inos_polymere").Cells(numero_sechage_inos, column)
    End If
    If FeCl3_kg_tMS_force = True Then
        FeCl3_kg_tMS_f = Feuil38.Range("sechage_inos_fecl3").Cells(numero_sechage_inos, column)
    End If
    If chauffage_T_out_force = True Then
        chauffage_T_out_f = Feuil38.Range("sechage_inos_chauffage_T_out").Cells(numero_sechage_inos, column)
    End If
    If chauffage_T_in_force = True Then
        chauffage_T_in_f = Feuil38.Range("sechage_inos_chauffage_T_in").Cells(numero_sechage_inos, column)
    End If

    If besoins_thermiques_force = True Then
        besoins_thermiques_f = Feuil38.Range("sechage_inos_chaleur").Cells(numero_sechage_inos, column)
    End If
    If capture_MS_force = True Then
        capture_MS_f = Feuil38.Range("sechage_inos_capture_MS").Cells(numero_sechage_inos, column)
    End If
    If alimentation_pompe_nb_force = True Then
        alimentation_pompe_nb_f = Feuil38.Range("sechage_inos_alimentation_nb_pompe").Cells(numero_sechage_inos, column)
    End If
    If alimentation_P_refoulement_force = True Then
        alimentation_P_refoulement_f = Feuil38.Range("sechage_inos_alimentation_P_refoulement").Cells(numero_sechage_inos, column)
    End If
    If alimentation_tps_fonctionnement_force = True Then
        alimentation_tps_fonctionnement_f = Feuil38.Range("sechage_inos_alimentation_tps_fonctionnement").Cells(numero_sechage_inos, column)
    End If
    If alimentation_Q_unitaire_force = True Then
        alimentation_Q_unitaire_f = Feuil38.Range("sechage_inos_alimentation_Q_unitaire").Cells(numero_sechage_inos, column)
    End If
    If alimentation_rendement_global_force = True Then
        alimentation_rendement_global_f = Feuil38.Range("sechage_inos_alimentation_rendement_pompe").Cells(numero_sechage_inos, column)
    End If
    If alimentation2_pompe_nb_force = True Then
        alimentation2_pompe_nb_f = Feuil38.Range("sechage_inos_alimentation2_nb_pompe").Cells(numero_sechage_inos, column)
    End If
    If alimentation2_P_refoulement_force = True Then
        alimentation2_P_refoulement_f = Feuil38.Range("sechage_inos_alimentation2_P_refoulement").Cells(numero_sechage_inos, column)
    End If
    If alimentation2_tps_fonctionnement_force = True Then
        alimentation2_tps_fonctionnement_f = Feuil38.Range("sechage_inos_alimentation2_tps_fonctionnement").Cells(numero_sechage_inos, column)
    End If
    If alimentation2_Q_unitaire_force = True Then
        alimentation2_Q_unitaire_f = Feuil38.Range("sechage_inos_alimentation2_Q_unitaire").Cells(numero_sechage_inos, column)
    End If
    If alimentation2_rendement_global_force = True Then
        alimentation2_rendement_global_f = Feuil38.Range("sechage_inos_alimentation2_rendement_pompe").Cells(numero_sechage_inos, column)
    End If

    'énergie thermique
    If condensat_indirect_ratio_vapeur_condensee_force = True Then
        condensat_indirect_ratio_vapeur_condensee_f = Feuil38.Range("sechage_inos_condensat_indirect_ratio_vapeur_condensee").Cells(numero_sechage_inos, column)
    End If
    If condensat_indirect_recup_eau_HT_force = True Then
        condensat_indirect_recup_eau_HT_f = Feuil38.Range("sechage_inos_condensat_indirect_recup_eau_HT").Cells(numero_sechage_inos, column)
    End If
    If condensation_indirecte_Teau_in_force = True Then
        condensation_indirecte_Teau_in_f = Feuil38.Range("sechage_inos_condensation_indirecte_Teau_in").Cells(numero_sechage_inos, column)
    End If
    If condensation_indirecte_Teau_out_force = True Then
        condensation_indirecte_Teau_out_f = Feuil38.Range("sechage_inos_condensation_indirecte_Teau_out").Cells(numero_sechage_inos, column)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "38.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub transfert_flux_energie_vers_pinch(ByVal numero_sechage_inos)

    Dim nom_flux As String
    Dim T_in As Double
    Dim T_out As Double
    Dim MCp As Double

    'INOS - Fluide de séchage

    'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
    '    nom_flux = "INOS " & numero_sechage_inos & " - Fluide de séchage"
    nom_flux = Feuil5.Range("nom_flux_inos1").Cells(1, 1).Value & " #" & numero_sechage_inos
    T_in = chauffage_T_in
    T_out = chauffage_T_out
    MCp = besoins_thermiques / (T_out - T_in)    'kWh/j/°C

    'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
    Call copie_dans_flux_thermique_process(nom_flux, T_in, T_out, MCp)

    'INOS - Condensation indirecte

    'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
    '    nom_flux = "INOS " & numero_sechage_inos & " - Condensation indirecte"
    nom_flux = Feuil5.Range("nom_flux_inos2").Cells(1, 1).Value & " #" & numero_sechage_inos
    T_in = condensation_indirecte_Teau_in
    T_out = condensation_indirecte_Teau_out
    MCp = condensat_indirect_recup_eau_HT / (T_in - T_out)    'kWh/j/°C

    'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
    Call copie_dans_flux_thermique_process(nom_flux, T_in, T_out, MCp)


End Sub

Sub dimensionnement(ByRef boues_flux, ByVal numero_sechage_inos, ByRef retour_flux, ByVal boues_I_detail, ByVal boues_II_detail, ByVal boues_III_detail, ByRef boues_pollution_soluble, ByRef besoins_th, ByRef disponibilites_thermiques)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Q_retour As Double


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
    Dim outlet_MV As Double
    Dim soluble_kg(nb_repere_mgL) As Double
    Dim ratio_boues_bio As Double
    'gestion des ratios de MV réduits par type de boues
    Dim MV_reduits(nb_type_boues) As Double
    Dim delta_MV_reduits As Double
    Dim condensat_direct_eau_inlet_kgj As Double

    inlet_MES = 0
    inlet_MV = 0
    inlet_Q = 0
    inlet_DCO = 0
    inlet_NK = 0
    inlet_Pt = 0

    outlet_MES = 0
    outlet_MV = 0
    outlet_DCO = 0
    outlet_NK = 0
    outlet_Pt = 0

    For I = 1 To nb_inlet

        If choix_inlet(I) <> inlet_aucun Then

            If inlet_etape_anterieure(I) = boues_inlet Then

                'variables d'entrée
                intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MES)
                If intermediaire_MES <> 0 Then
                    intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_Q)
                    intermediaire_MV = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MV_MES) * intermediaire_MES
                    intermediaire_DCO = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DCO_MES) * intermediaire_MES
                    intermediaire_DBO = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DBO_MES) * intermediaire_MES
                    intermediaire_NK = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_NK_MES) * intermediaire_MES
                    intermediaire_Pt = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_Pt_MES) * intermediaire_MES

                    'ENTREE SOLUBLE
                    'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                    For k = 1 To nb_repere_mgL
                        soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), k) * intermediaire_Q / 1000
                    Next k

                    inlet_MES = inlet_MES + intermediaire_MES
                    inlet_Q = inlet_Q + intermediaire_Q
                    inlet_DCO = inlet_DCO + intermediaire_DCO    '+ soluble_kg(repere_mgL_DCO)
                    inlet_DBO = inlet_DBO + intermediaire_DBO
                    inlet_NK = inlet_NK + intermediaire_NK    '+ soluble_kg(repere_mgL_NK)
                    inlet_Pt = inlet_Pt + intermediaire_Pt    '+ soluble_kg(repere_mgL_Pt)
                    inlet_MV = inlet_MV + intermediaire_MV

                    'SEPARATION
                    intermediaire_MES = intermediaire_MES * capture_MS
                    intermediaire_MV = capture_MS * intermediaire_MV
                    intermediaire_DCO = intermediaire_DCO - ratio_DCOret_MVin * intermediaire_MV    '* (1 - capture_MS)
                    If intermediaire_DCO < 0 Then
                        intermediaire_DCO = 0
                    End If
                    intermediaire_DBO = intermediaire_DBO - ratio_retour_DBO_DCO * ratio_DCOret_MVin * intermediaire_MV    '* (1 - capture_MS)
                    If intermediaire_DBO < 0 Then
                        intermediaire_DBO = 0
                    End If
                    intermediaire_NK = intermediaire_NK - ratio_NKret_NKin * intermediaire_NK    '- soluble_kg(repere_mgL_NK)
                    If intermediaire_NK < 0 Then
                        intermediaire_NK = 0
                    End If
                    intermediaire_Pt = (1 - ratio_Pret_Pin) * intermediaire_Pt
                    If intermediaire_Pt < 0 Then
                        intermediaire_Pt = 0
                    End If
                    'intermediaire_Q = intermediaire_MES / outlet_concentration

                    'variables de sortie
                    outlet_MES = outlet_MES + intermediaire_MES
                    outlet_DCO = outlet_DCO + intermediaire_DCO
                    outlet_DBO = outlet_DBO + intermediaire_DBO
                    outlet_NK = outlet_NK + intermediaire_NK
                    outlet_Pt = outlet_Pt + intermediaire_Pt
                    'outlet_Q = outlet_Q + intermediaire_Q

                    'je calcule les paramètres des boues séchees
                    intermediaire_MV = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MV_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES) + intermediaire_MV
                    intermediaire_DCO = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_DCO_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES) + intermediaire_DCO
                    intermediaire_NK = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_NK_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES) + intermediaire_NK
                    intermediaire_Pt = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_Pt_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES) + intermediaire_Pt
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES) = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES) + intermediaire_MES
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_Q) = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_Q) + intermediaire_MES / outlet_siccite / 1000
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MV_MES) = intermediaire_MV / boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_DBO_MES) = intermediaire_DBO / boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES)

                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_flux_in) = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_flux_in)

                    'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                    boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                End If

            Else

                For j = 1 To nb_type_boues
                    'variables d'entrée
                    intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_MES)
                    If intermediaire_MES <> 0 Then
                        intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_Q)
                        intermediaire_MV = boues_flux(inlet_etape_anterieure(I), j, repere_MV_MES) * intermediaire_MES
                        intermediaire_DCO = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DCO_MES) * intermediaire_MES
                        intermediaire_DBO = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DBO_MES) * intermediaire_MES
                        intermediaire_NK = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_NK_MES) * intermediaire_MES
                        intermediaire_Pt = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_Pt_MES) * intermediaire_MES

                        'ENTREE SOLUBLE
                        'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                        For k = 1 To nb_repere_mgL
                            soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), j, k) * intermediaire_Q / 1000
                        Next k

                        inlet_MES = inlet_MES + intermediaire_MES
                        inlet_Q = inlet_Q + intermediaire_Q
                        inlet_DCO = inlet_DCO + intermediaire_DCO + soluble_kg(repere_mgL_DCO)
                        inlet_DBO = inlet_DBO + intermediaire_DBO
                        inlet_NK = inlet_NK + intermediaire_NK + soluble_kg(repere_mgL_NK)
                        inlet_Pt = inlet_Pt + intermediaire_Pt + soluble_kg(repere_mgL_Pt)
                        inlet_MV = inlet_MV + intermediaire_MV

                        'SEPARATION
                        intermediaire_MES = intermediaire_MES * capture_MS
                        intermediaire_MV = capture_MS * intermediaire_MV
                        intermediaire_DCO = intermediaire_DCO - ratio_DCOret_MVin * intermediaire_MV    '* (1 - capture_MS)
                        If intermediaire_DCO < 0 Then
                            intermediaire_DCO = 0
                        End If
                        intermediaire_DBO = intermediaire_DBO - ratio_retour_DBO_DCO * ratio_DCOret_MVin * intermediaire_MV    '* (1 - capture_MS)
                        If intermediaire_DBO < 0 Then
                            intermediaire_DBO = 0
                        End If
                        intermediaire_NK = intermediaire_NK - ratio_NKret_NKin * intermediaire_NK    '- soluble_kg(repere_mgL_NK)
                        If intermediaire_NK < 0 Then
                            intermediaire_NK = 0
                        End If
                        intermediaire_Pt = (1 - ratio_Pret_Pin) * intermediaire_Pt
                        If intermediaire_Pt < 0 Then
                            intermediaire_Pt = 0
                        End If
                        'intermediaire_Q = intermediaire_MES / outlet_concentration

                        'variables de sortie
                        outlet_MES = outlet_MES + intermediaire_MES
                        outlet_DCO = outlet_DCO + intermediaire_DCO
                        outlet_DBO = outlet_DBO + intermediaire_DBO
                        outlet_NK = outlet_NK + intermediaire_NK
                        outlet_Pt = outlet_Pt + intermediaire_Pt
                        'outlet_Q = outlet_Q + intermediaire_Q

                        'je calcule les paramètres des boues séchees
                        intermediaire_MV = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MV_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES) + intermediaire_MV
                        intermediaire_DCO = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_DCO_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES) + intermediaire_DCO
                        intermediaire_NK = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_NK_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES) + intermediaire_NK
                        intermediaire_Pt = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_Pt_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES) + intermediaire_Pt
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES) = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES) + intermediaire_MES
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_Q) = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_Q) + intermediaire_MES / outlet_siccite / 1000
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MV_MES) = intermediaire_MV / boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES)
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES)
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_DBO_MES) = intermediaire_DBO / boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES)
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES)
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES)

                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_flux_in) = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_verif_flux) = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_flux_in)

                        'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                        boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                    End If

                Next j

            End If

        End If

    Next I

    If inlet_MES = 0 Or inlet_Q = 0 Then
        MsgBox Feuil5.Range("erreur_choix_sechage_inos_inlet") & " " & numero_sechage_inos, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If

    inlet_concentration = inlet_MES / inlet_Q

    'CALCUL DE LA CONCENTRATION EN SORTIE AVEC UNE TELLE REPARTITION DU TYPE DE BOUES

    ratio_boues_digerees = ratio_boues_digerees / inlet_MES
    correction_concentration_boues_digerees = correction_concentration_boues_digerees * ratio_boues_digerees
    intermediaire_concentration = 0
    For I = 1 To nb_type_boues
        If I = boues_I Then
            ratio_boues_bio = 0
            For j = 1 To nb_eaux_sales_max
                ratio_boues_bio = ratio_boues_bio + boues_I_detail(j, repere_ratio_MES)
                intermediaire_origine = boues_I_detail(j, repere_origine)
                intermediaire_MES = boues_I_detail(j, repere_ratio_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_MES)
                intermediaire_concentration = intermediaire_concentration + intermediaire_MES / outlet_MES * intermediaire_concentration_reference(intermediaire_origine)
            Next j
            intermediaire_MES = boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_MES) * (1 - ratio_boues_bio)
            intermediaire_origine = boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_origine)
            intermediaire_concentration = intermediaire_concentration + intermediaire_MES / outlet_MES * intermediaire_concentration_reference(intermediaire_origine)
        ElseIf I = boues_II And choix_biostyr = True Then
            For j = 1 To nb_eaux_sales_max
                ratio_boues_bio = boues_II_detail(j, repere_ratio_MES)
                intermediaire_origine = boues_II_detail(j, repere_origine)
                intermediaire_MES = boues_II_detail(j, repere_ratio_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_MES)
                intermediaire_concentration = intermediaire_concentration + intermediaire_MES / outlet_MES * intermediaire_concentration_reference(intermediaire_origine)
            Next j
        ElseIf I = boues_III And choix_biostyr_N_III = True Then
            For j = 1 To nb_eaux_sales_max
                ratio_boues_bio = boues_III_detail(j, repere_ratio_MES)
                intermediaire_origine = boues_III_detail(j, repere_origine)
                intermediaire_MES = boues_III_detail(j, repere_ratio_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_MES)
                intermediaire_concentration = intermediaire_concentration + intermediaire_MES / outlet_MES * intermediaire_concentration_reference(intermediaire_origine)
            Next j
        Else
            intermediaire_MES = boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_MES)
            intermediaire_origine = boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_origine)
            intermediaire_concentration = intermediaire_concentration + intermediaire_MES / outlet_MES * intermediaire_concentration_reference(intermediaire_origine)
        End If
    Next I
    intermediaire_concentration = intermediaire_concentration + correction_concentration_boues_digerees


    'BOUES EN ENTREE
    inlet_siccite = inlet_MES / (inlet_Q * masse_volumique_boues)
    inlet_BH = inlet_MES / inlet_siccite
    inlet_MV_MES = inlet_MV / inlet_MES

    'BOUES EN SORTIE
    For I = 1 To 2
        If incineration_repere_boues_inlet(I) = boues_sechees_inos(numero_sechage_inos) Then
            If incineration_siccite_ajustable(I) = True Then
                If outlet_siccite_force = True Then
                    outlet_siccite = outlet_siccite_f
                    incineration_siccite_ajustable(I) = False
                Else
                    'If choix_performance = sechage_partiel Then
                    If incineration_siccite_autocombustibilite(I) < intermediaire_concentration / masse_volumique_boues Then
                        outlet_siccite = intermediaire_concentration / masse_volumique_boues
                        secheur_pas_necessaire = True
                        incineration_siccite_ajustable(I) = False
                    Else
                        outlet_siccite = incineration_siccite_autocombustibilite(I)
                    End If
                    'Else
                    '    incineration_siccite_ajustable(i) = False
                End If
                'End If
            ElseIf secheur_pas_necessaire = True Then
                outlet_siccite = intermediaire_concentration / masse_volumique_boues
            End If
        End If
    Next I

    outlet_BH = outlet_MES / outlet_siccite

    outlet_MV_MES = inlet_MV_MES

    'EAU EVAPOREE
    If outlet_siccite > (intermediaire_concentration / 1000) Then
        eau_evaporee = ((outlet_MES * (1 - (intermediaire_concentration / 1000)) / (intermediaire_concentration / 1000)) - (outlet_MES * (1 - outlet_siccite) / outlet_siccite))
    Else
        eau_evaporee = 0
    End If

    chaux_eteinte_flux = chaux / 1000 * inlet_MES
    FeCl3_flux = FeCl3_kg_tMS / 1000 * inlet_MES
    polymere_flux = polymere_kg_tMS / 1000 * inlet_MES

    'énergie thermique nécessaire
    besoins_thermiques = eau_evaporee / 1000 * ratio_chal

    'RECUPERATION D'ENERGIE ET RETOURS
    If choix_condensation = condensation_directe Then
        'hypothèse: toute la vapeur restante est évaporée sur la condensation directe
        condensat_direct_eau_inlet_kgj = eau_evaporee * (deltaH_vap_kWh + Cp_eau_kWh * (condensat_direct_Tair_inlet - condensat_direct_Teau_outlet)) / (Cp_eau_kWh * (condensat_direct_Teau_outlet - condensat_direct_Teau_inlet))
        Q_retour = eau_evaporee + condensat_direct_eau_inlet_kgj
        condensat_indirect_recup_eau_HT = 0
    ElseIf choix_condensation = condensation_indirecte Then
        Q_retour = condensat_indirect_ratio_vapeur_condensee * eau_evaporee
        condensat_indirect_recup_eau_HT = Q_retour * deltaH_vap_kWh
    Else   'indirecte puis directe
        Q_retour = condensat_indirect_ratio_vapeur_condensee * eau_evaporee
        condensat_indirect_recup_eau_HT = Q_retour * deltaH_vap_kWh
        'hypothèse: toute la vapeur restante est évaporée sur la condensation directe
        condensat_direct_eau_inlet_kgj = (eau_evaporee - Q_retour) * (deltaH_vap_kWh + Cp_eau_kWh * (condensat_direct_Tair_inlet - condensat_direct_Teau_outlet)) / (Cp_eau_kWh * (condensat_direct_Teau_outlet - condensat_direct_Teau_inlet))
        Q_retour = eau_evaporee + condensat_direct_eau_inlet_kgj
    End If
    Q_retour = Q_retour / 1000       'avant, on le calcule en kg/j et là on le passe en m3/j


    'Besoins thermiques et disponibilites
    If choix_analyse_pinch = False Then
        If choix_source_energie = source_eau_HT Then
            besoins_th(energie_eau_HT) = besoins_th(energie_eau_HT) + besoins_thermiques
        ElseIf choix_source_energie = source_vapeur Then
            besoins_th(energie_vapeur_sat2) = besoins_th(energie_vapeur_sat2) + besoins_thermiques    'attention si plusieurs sécheurs
        ElseIf choix_source_energie = source_biogaz Then
            besoins_th(energie_biogaz) = besoins_th(energie_biogaz) + besoins_thermiques
        Else
            besoins_th(energie_combustible) = besoins_th(energie_combustible) + besoins_thermiques
        End If
        disponibilites_thermiques(energie_eau_HT) = disponibilites_thermiques(energie_eau_HT) + condensat_indirect_recup_eau_HT
    End If

    retour_flux(repere_ret_Q) = retour_flux(repere_ret_Q) + Q_retour
    retour_flux(repere_ret_DCO) = retour_flux(repere_ret_DCO) + inlet_DCO - outlet_DCO
    retour_flux(repere_ret_DBO) = retour_flux(repere_ret_DBO) + inlet_DBO - outlet_DBO
    retour_flux(repere_ret_MES) = retour_flux(repere_ret_MES) + inlet_MES - outlet_MES
    retour_flux(repere_ret_NK) = retour_flux(repere_ret_NK) + inlet_NK - outlet_NK
    retour_flux(repere_ret_NH4) = retour_flux(repere_ret_NH4) + inlet_NK - outlet_NK
    retour_flux(repere_ret_Pt) = retour_flux(repere_ret_Pt) + inlet_Pt - outlet_Pt

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "38.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub fonctionnement_reel(ByRef boues_flux, ByVal numero_sechage_inos, ByVal boues_I_detail, ByVal boues_II_detail, ByVal boues_III_detail, ByRef retour_flux, ByRef boues_pollution_soluble, ByRef besoins_th, ByRef disponibilites_thermiques)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Q_retour As Double


    Dim intermediaire_MES As Double
    Dim intermediaire_Q As Double
    Dim intermediaire_NK As Double
    Dim intermediaire_Pt As Double
    Dim intermediaire_DCO As Double
    Dim intermediaire_DBO As Double
    Dim outlet_Q As Double
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
    Dim outlet_MV As Double
    Dim soluble_kg(nb_repere_mgL) As Double
    Dim ratio_boues_bio As Double
    'gestion des ratios de MV réduits par type de boues
    Dim MV_reduits(nb_type_boues) As Double
    Dim delta_MV_reduits As Double
    Dim condensat_direct_eau_inlet_kgj As Double

    inlet_MES = 0
    inlet_MV = 0
    inlet_Q = 0
    inlet_DCO = 0
    inlet_NK = 0
    inlet_Pt = 0

    outlet_MES = 0
    outlet_MV = 0
    outlet_DCO = 0
    outlet_NK = 0
    outlet_Pt = 0

    For I = 1 To nb_inlet

        If choix_inlet(I) <> inlet_aucun Then
            If inlet_etape_anterieure(I) = boues_inlet Then

                'variables d'entrée
                intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MES)
                If intermediaire_MES <> 0 Then
                    intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_Q)
                    intermediaire_MV = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MV_MES) * intermediaire_MES
                    intermediaire_DCO = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DCO_MES) * intermediaire_MES
                    intermediaire_DBO = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DBO_MES) * intermediaire_MES
                    intermediaire_NK = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_NK_MES) * intermediaire_MES
                    intermediaire_Pt = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_Pt_MES) * intermediaire_MES

                    'ENTREE SOLUBLE
                    'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                    For k = 1 To nb_repere_mgL
                        soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), k) * intermediaire_Q / 1000
                    Next k

                    inlet_MES = inlet_MES + intermediaire_MES
                    inlet_Q = inlet_Q + intermediaire_Q
                    inlet_DCO = inlet_DCO + intermediaire_DCO + soluble_kg(repere_mgL_DCO)
                    inlet_DBO = inlet_DBO + intermediaire_DBO
                    inlet_NK = inlet_NK + intermediaire_NK + soluble_kg(repere_mgL_NK)
                    inlet_Pt = inlet_Pt + intermediaire_Pt + soluble_kg(repere_mgL_Pt)
                    inlet_MV = inlet_MV + intermediaire_MV

                    'SEPARATION
                    intermediaire_MES = intermediaire_MES * capture_MS
                    intermediaire_MV = capture_MS * intermediaire_MV
                    intermediaire_DCO = intermediaire_DCO - ratio_DCOret_MVin * intermediaire_MV    '* (1 - capture_MS)
                    If intermediaire_DCO < 0 Then
                        intermediaire_DCO = 0
                    End If
                    intermediaire_DBO = intermediaire_DBO - ratio_retour_DBO_DCO * ratio_DCOret_MVin * intermediaire_MV    '* (1 - capture_MS)
                    If intermediaire_DBO < 0 Then
                        intermediaire_DBO = 0
                    End If
                    intermediaire_NK = intermediaire_NK - ratio_NKret_NKin * intermediaire_NK    '- soluble_kg(repere_mgL_NK)
                    intermediaire_Pt = (1 - ratio_Pret_Pin) * intermediaire_Pt
                    'intermediaire_Q = intermediaire_MES / outlet_concentration

                    'variables de sortie
                    outlet_MES = outlet_MES + intermediaire_MES
                    outlet_DCO = outlet_DCO + intermediaire_DCO
                    outlet_DBO = outlet_DBO + intermediaire_DBO
                    outlet_NK = outlet_NK + intermediaire_NK
                    outlet_Pt = outlet_Pt + intermediaire_Pt
                    'outlet_Q = outlet_Q + intermediaire_Q

                    'je calcule les paramètres des boues séchees
                    intermediaire_MV = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MV_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES) + intermediaire_MV
                    intermediaire_DCO = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_DCO_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES) + intermediaire_DCO
                    intermediaire_NK = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_NK_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES) + intermediaire_NK
                    intermediaire_Pt = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_Pt_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES) + intermediaire_Pt
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES) = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES) + intermediaire_MES
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_Q) = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_Q) + intermediaire_MES / outlet_siccite / 1000
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MV_MES) = intermediaire_MV / boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_DBO_MES) = intermediaire_DBO / boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_MES)

                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_flux_in) = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_sechees_inos(numero_sechage_inos), inlet_type_boues(I), repere_flux_in)

                    'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                    boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                End If

            Else

                For j = 1 To nb_type_boues

                    'variables d'entrée
                    intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_MES)
                    If intermediaire_MES <> 0 Then
                        intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_Q)
                        intermediaire_MV = boues_flux(inlet_etape_anterieure(I), j, repere_MV_MES) * intermediaire_MES
                        intermediaire_DCO = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DCO_MES) * intermediaire_MES
                        intermediaire_DBO = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DBO_MES) * intermediaire_MES
                        intermediaire_NK = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_NK_MES) * intermediaire_MES
                        intermediaire_Pt = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_Pt_MES) * intermediaire_MES

                        'ENTREE SOLUBLE
                        'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                        For k = 1 To nb_repere_mgL
                            soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), j, k) * intermediaire_Q / 1000
                        Next k

                        inlet_MES = inlet_MES + intermediaire_MES
                        inlet_Q = inlet_Q + intermediaire_Q
                        inlet_DCO = inlet_DCO + intermediaire_DCO    '+ soluble_kg(repere_mgL_DCO)
                        inlet_DBO = inlet_DBO + intermediaire_DBO
                        inlet_NK = inlet_NK + intermediaire_NK    '+ soluble_kg(repere_mgL_NK)
                        inlet_Pt = inlet_Pt + intermediaire_Pt    '+ soluble_kg(repere_mgL_Pt)
                        inlet_MV = inlet_MV + intermediaire_MV

                        'SEPARATION
                        intermediaire_MES = intermediaire_MES * capture_MS
                        intermediaire_MV = capture_MS * intermediaire_MV
                        intermediaire_DCO = intermediaire_DCO - ratio_DCOret_MVin * intermediaire_MV    '* (1 - capture_MS)
                        If intermediaire_DCO < 0 Then
                            intermediaire_DCO = 0
                        End If
                        intermediaire_DBO = intermediaire_DBO - ratio_retour_DBO_DCO * ratio_DCOret_MVin * intermediaire_MV    '* (1 - capture_MS)
                        If intermediaire_DBO < 0 Then
                            intermediaire_DBO = 0
                        End If
                        intermediaire_NK = intermediaire_NK - ratio_NKret_NKin * intermediaire_NK    '- soluble_kg(repere_mgL_NK)
                        If intermediaire_NK < 0 Then
                            intermediaire_NK = 0
                        End If
                        intermediaire_Pt = (1 - ratio_Pret_Pin) * intermediaire_Pt
                        If intermediaire_Pt < 0 Then
                            intermediaire_Pt = 0
                        End If
                        'intermediaire_Q = intermediaire_MES / outlet_concentration

                        'variables de sortie
                        outlet_MES = outlet_MES + intermediaire_MES
                        outlet_DCO = outlet_DCO + intermediaire_DCO
                        outlet_DBO = outlet_DBO + intermediaire_DBO
                        outlet_NK = outlet_NK + intermediaire_NK
                        outlet_Pt = outlet_Pt + intermediaire_Pt
                        'outlet_Q = outlet_Q + intermediaire_Q

                        'je calcule les paramètres des boues séchees
                        intermediaire_MV = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MV_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES) + intermediaire_MV
                        intermediaire_DCO = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_DCO_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES) + intermediaire_DCO
                        intermediaire_NK = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_NK_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES) + intermediaire_NK
                        intermediaire_Pt = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_Pt_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES) + intermediaire_Pt
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES) = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES) + intermediaire_MES
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_Q) = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_Q) + intermediaire_MES / outlet_siccite / 1000
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MV_MES) = intermediaire_MV / boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES)
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES)
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_DBO_MES) = intermediaire_DBO / boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES)
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES)
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_MES)

                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_flux_in) = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                        boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_verif_flux) = boues_flux(boues_sechees_inos(numero_sechage_inos), j, repere_flux_in)

                        'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                        boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                    End If

                Next j
            End If
        End If
    Next I




    inlet_concentration = inlet_MES / inlet_Q

    'CALCUL DE LA CONCENTRATION EN SORTIE AVEC UNE TELLE REPARTITION DU TYPE DE BOUES

    If intermediaire_concentration_force = True Then
        intermediaire_concentration = intermediaire_concentration_f
    Else
        ratio_boues_digerees = ratio_boues_digerees / inlet_MES
        correction_concentration_boues_digerees = correction_concentration_boues_digerees * ratio_boues_digerees
        intermediaire_concentration = 0
        For I = 1 To nb_type_boues
            If I = boues_I Then
                ratio_boues_bio = 0
                For j = 1 To nb_eaux_sales_max
                    ratio_boues_bio = ratio_boues_bio + boues_I_detail(j, repere_ratio_MES)
                    intermediaire_origine = boues_I_detail(j, repere_origine)
                    intermediaire_MES = boues_I_detail(j, repere_ratio_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_MES)
                    intermediaire_concentration = intermediaire_concentration + intermediaire_MES / outlet_MES * intermediaire_concentration_reference(intermediaire_origine)
                Next j
                intermediaire_MES = boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_MES) * (1 - ratio_boues_bio)
                intermediaire_origine = boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_origine)
                intermediaire_concentration = intermediaire_concentration + intermediaire_MES / outlet_MES * intermediaire_concentration_reference(intermediaire_origine)
            ElseIf I = boues_II And choix_biostyr = True Then
                For j = 1 To nb_eaux_sales_max
                    ratio_boues_bio = boues_II_detail(j, repere_ratio_MES)
                    intermediaire_origine = boues_II_detail(j, repere_origine)
                    intermediaire_MES = boues_II_detail(j, repere_ratio_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_MES)
                    intermediaire_concentration = intermediaire_concentration + intermediaire_MES / outlet_MES * intermediaire_concentration_reference(intermediaire_origine)
                Next j
            ElseIf I = boues_III And choix_biostyr_N_III = True Then
                For j = 1 To nb_eaux_sales_max
                    ratio_boues_bio = boues_III_detail(j, repere_ratio_MES)
                    intermediaire_origine = boues_III_detail(j, repere_origine)
                    intermediaire_MES = boues_III_detail(j, repere_ratio_MES) * boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_MES)
                    intermediaire_concentration = intermediaire_concentration + intermediaire_MES / outlet_MES * intermediaire_concentration_reference(intermediaire_origine)
                Next j
            Else
                intermediaire_MES = boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_MES)
                intermediaire_origine = boues_flux(boues_sechees_inos(numero_sechage_inos), I, repere_origine)
                intermediaire_concentration = intermediaire_concentration + intermediaire_MES / outlet_MES * intermediaire_concentration_reference(intermediaire_origine)
            End If
        Next I
        intermediaire_concentration = intermediaire_concentration + correction_concentration_boues_digerees
    End If


    'BOUES EN ENTREE
    inlet_siccite = inlet_MES / (inlet_Q * masse_volumique_boues)
    inlet_BH = inlet_MES / inlet_siccite
    inlet_MV_MES = inlet_MV / inlet_MES

    'BOUES EN SORTIE
    For I = 1 To 2
        If incineration_repere_boues_inlet(I) = boues_sechees_inos(numero_sechage_inos) Then
            If incineration_siccite_ajustable(I) = True Then
                If outlet_siccite_force = True Then
                    outlet_siccite = outlet_siccite_f
                    incineration_siccite_ajustable(I) = False
                Else
                    'If choix_performance = sechage_partiel Then
                    If incineration_siccite_autocombustibilite(I) < intermediaire_concentration / masse_volumique_boues Then
                        outlet_siccite = intermediaire_concentration / masse_volumique_boues
                        secheur_pas_necessaire = True
                        incineration_siccite_ajustable(I) = False
                    Else
                        outlet_siccite = incineration_siccite_autocombustibilite(I)
                    End If
                    'Else
                    '    incineration_siccite_ajustable(i) = False
                End If
                'End If
            ElseIf secheur_pas_necessaire = True Then
                outlet_siccite = intermediaire_concentration / masse_volumique_boues
            End If
        End If
    Next I

    If outlet_BH_force = True Then
        outlet_BH = outlet_BH_f
    Else
        outlet_BH = outlet_MES / outlet_siccite
    End If

    If outlet_MV_MES_force = True Then
        outlet_MV_MES = outlet_MV_MES_f
    Else
        outlet_MV_MES = inlet_MV_MES
    End If

    'EAU EVAPOREE
    If outlet_siccite > (intermediaire_concentration / 1000) Then
        eau_evaporee = ((outlet_MES * (1 - (intermediaire_concentration / 1000)) / (intermediaire_concentration / 1000)) - (outlet_MES * (1 - outlet_siccite) / outlet_siccite))
    Else
        eau_evaporee = 0
    End If

    chaux_eteinte_flux = chaux / 1000 * inlet_MES
    FeCl3_flux = FeCl3_kg_tMS / 1000 * inlet_MES
    polymere_flux = polymere_kg_tMS / 1000 * inlet_MES

    If besoins_thermiques_force = True Then
        besoins_thermiques = besoins_thermiques_f
    Else
        besoins_thermiques = eau_evaporee / 1000 * ratio_chal
    End If

    'RECUPERATION D'ENERGIE ET RETOURS
    If choix_condensation = condensation_directe Then
        'hypothèse: toute la vapeur restante est évaporée sur la condensation directe
        condensat_direct_eau_inlet_kgj = eau_evaporee * (deltaH_vap_kWh + Cp_eau_kWh * (condensat_direct_Tair_inlet - condensat_direct_Teau_outlet)) / (Cp_eau_kWh * (condensat_direct_Teau_outlet - condensat_direct_Teau_inlet))
        Q_retour = eau_evaporee + condensat_direct_eau_inlet_kgj
    ElseIf choix_condensation = condensation_indirecte Then
        Q_retour = condensat_indirect_ratio_vapeur_condensee * eau_evaporee
        If condensat_indirect_recup_eau_HT_force = True Then
            condensat_indirect_recup_eau_HT = condensat_indirect_recup_eau_HT_f
        Else
            condensat_indirect_recup_eau_HT = Q_retour * deltaH_vap_kWh
        End If
    Else   'indirecte puis directe
        Q_retour = condensat_indirect_ratio_vapeur_condensee * eau_evaporee
        If condensat_indirect_recup_eau_HT_force = True Then
            condensat_indirect_recup_eau_HT = condensat_indirect_recup_eau_HT_f
        Else
            condensat_indirect_recup_eau_HT = Q_retour * deltaH_vap_kWh
        End If
        'hypothèse: toute la vapeur restante est évaporée sur la condensation directe
        condensat_direct_eau_inlet_kgj = (eau_evaporee - Q_retour) * (deltaH_vap_kWh + Cp_eau_kWh * (condensat_direct_Tair_inlet - condensat_direct_Teau_outlet)) / (Cp_eau_kWh * (condensat_direct_Teau_outlet - condensat_direct_Teau_inlet))
        Q_retour = eau_evaporee + condensat_direct_eau_inlet_kgj
    End If
    Q_retour = Q_retour / 1000       'avant, on le calcule en kg/j et là on le passe en m3/j


    'Besoins thermiques et disponibilites
    If choix_analyse_pinch = False Then
        If choix_source_energie = source_eau_HT Then
            besoins_th(energie_eau_HT) = besoins_th(energie_eau_HT) + besoins_thermiques
        ElseIf choix_source_energie = source_vapeur Then
            besoins_th(energie_vapeur_sat2) = besoins_th(energie_vapeur_sat2) + besoins_thermiques    'attention si plusieurs sécheurs
        ElseIf choix_source_energie = source_biogaz Then
            besoins_th(energie_biogaz) = besoins_th(energie_biogaz) + besoins_thermiques
        Else
            besoins_th(energie_combustible) = besoins_th(energie_combustible) + besoins_thermiques
        End If
        disponibilites_thermiques(energie_eau_HT) = disponibilites_thermiques(energie_eau_HT) + condensat_indirect_recup_eau_HT
    End If

    retour_flux(repere_ret_Q) = retour_flux(repere_ret_Q) + Q_retour
    retour_flux(repere_ret_DCO) = retour_flux(repere_ret_DCO) + inlet_DCO - outlet_DCO
    retour_flux(repere_ret_DBO) = retour_flux(repere_ret_DBO) + inlet_DBO - outlet_DBO
    retour_flux(repere_ret_MES) = retour_flux(repere_ret_MES) + inlet_MES - outlet_MES
    retour_flux(repere_ret_NK) = retour_flux(repere_ret_NK) + inlet_NK - outlet_NK
    retour_flux(repere_ret_NH4) = retour_flux(repere_ret_NH4) + inlet_NK - outlet_NK
    retour_flux(repere_ret_Pt) = retour_flux(repere_ret_Pt) + inlet_Pt - outlet_Pt

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "38.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub
Sub calcul_consommation_electrique()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim puissance_pompe_a As Double   'puissance pompe alimentation
    Dim puissance_pompe_b As Double   'puissance pompe alimentation
    Dim puissance_pompe As Double    'puissance pompe extraction
    Dim rendement_moteur_a As Double
    Dim rendement_moteur As Double
    Dim rendement_moteur_b As Double
    Dim incoherence_a As Double
    Dim incoherence_b As Double
    Dim incoherence As Double
    Dim electricite_specifique As Double
    Dim electricite_alimentation As Double
    Dim electricite_alimentation1 As Double
    Dim electricite_alimentation2 As Double
    Dim ratio_boues_digerees As Double

    ratio_elec = 0
    If ratio_boues_digerees > 0 Then
        If choix_conditionnement = 1 Then
            ratio_elec = ratio_elec + ratio_boues_digerees * (143.24 - 0.0135 * inlet_MES / 1000)
            ratio_elec = ratio_elec + (1 - ratio_boues_digerees) * (143.24 - 0.0135 * inlet_MES / 1000)
        ElseIf choix_conditionnement = 2 Then
            ratio_elec = ratio_elec + ratio_boues_digerees * (143.24 - 0.0135 * inlet_MES / 1000)
            ratio_elec = ratio_elec + (1 - ratio_boues_digerees) * (143.24 - 0.0135 * inlet_MES / 1000)
        Else
            ratio_elec = ratio_elec + ratio_boues_digerees * (143.24 - 0.0135 * inlet_MES / 1000)
            ratio_elec = ratio_elec + (1 - ratio_boues_digerees) * (143.24 - 0.0135 * inlet_MES / 1000)

        End If
    Else
        If choix_conditionnement = 1 Then
            ratio_elec = ratio_elec + ratio_boues_digerees * (249.84 - 0.0044 * ratio_boues_digerees * inlet_MES / 1000)
            ratio_elec = ratio_elec + (1 - ratio_boues_digerees) * (249.84 - 0.0044 * (1 - ratio_boues_digerees) * inlet_MES / 1000)

        ElseIf choix_conditionnement = 2 Then
            ratio_elec = ratio_elec + ratio_boues_digerees * (180.84 - 0.0067 * ratio_boues_digerees * inlet_MES / 1000)
            ratio_elec = ratio_elec + (1 - ratio_boues_digerees) * (180.84 - 0.0067 * (1 - ratio_boues_digerees) * inlet_MES / 1000)
        Else
            ratio_elec = ratio_elec + ratio_boues_digerees * (178.72 - 0.0141 * ratio_boues_digerees * inlet_MES / 1000)
            ratio_elec = ratio_elec + (1 - ratio_boues_digerees) * (178.72 - 0.0141 * (1 - ratio_boues_digerees) * inlet_MES / 1000)
        End If
    End If

    electricite_specifique = inlet_MES / 1000 * ratio_elec


    If besoins_thermiques_force = True Then
        besoins_thermiques = besoins_thermiques_f
    Else
        besoins_thermiques = eau_evaporee / 1000 * ratio_chal
    End If


    'ALIMENTATION DES BOUES
    If choix_bache_amont = bache_amont_non Then   'si bâche amont --> pompes d'alimentation
        electricite_alimentation1 = 0
        electricite_alimentation2 = 0
    ElseIf choix_bache_amont = bache_amont_1 Then   'si bâche amont --> pompes d'alimentation
        'alimentation_Q_unitaire = inlet_Q / (alimentation_pompe_nb * alimentation_tps_fonctionnement)
        If alimentation_Q_unitaire_force = True Then
            alimentation_Q_unitaire = alimentation_Q_unitaire_f
            If alimentation_tps_fonctionnement_force = True Then
                'on ne fait rien
                If alimentation_pompe_nb_force = False Then
                    If Int(inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire)) < (inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire)) Then
                        alimentation_pompe_nb = Int(inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire)) + 1
                    Else
                        alimentation_pompe_nb = Int(inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire))
                    End If
                End If
            Else
                alimentation_tps_fonctionnement = inlet_Q / (alimentation_pompe_nb * alimentation_Q_unitaire)
                If alimentation_tps_fonctionnement > hyp_tps_fct_pompe_alim Then
                    alimentation_tps_fonctionnement = hyp_tps_fct_pompe_alim
                    If alimentation_pompe_nb_force = False Then
                        If Int(inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire)) < (inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire)) Then
                            alimentation_pompe_nb = Int(inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire)) + 1
                        Else
                            alimentation_pompe_nb = Int(inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire))
                        End If
                        alimentation_tps_fonctionnement = inlet_Q / (alimentation_pompe_nb * alimentation_Q_unitaire)
                    End If
                End If
            End If
        Else
            alimentation_Q_unitaire = inlet_Q / (alimentation_pompe_nb * alimentation_tps_fonctionnement)
        End If
        incoherence_a = (alimentation_Q_unitaire * alimentation_pompe_nb * alimentation_tps_fonctionnement - inlet_Q) / inlet_Q
        If Abs(incoherence_a) > critere_incoherence Then
            remarque(1) = Feuil38.Range("sechage_inos_erreur_pompage1").Cells(1, 1)
        End If
        puissance_pompe_a = alimentation_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * alimentation_P_refoulement / alimentation_rendement_pompe
        If puissance_pompe_a < 60 Then
            rendement_moteur_a = (5.7195 * Log(puissance_pompe_a) + 72.682) / 100
        Else
            rendement_moteur_a = 0.961
        End If
        If alimentation_rendement_global_force = True Then
            alimentation_rendement_global = alimentation_rendement_global_f
            electricite_alimentation1 = ((puissance_pompe_a * alimentation_rendement_pompe) / alimentation_rendement_global) * alimentation_pompe_nb * alimentation_tps_fonctionnement
        Else
            alimentation_rendement_global = rendement_moteur_a * alimentation_rendement_pompe
            electricite_alimentation1 = puissance_pompe_a / rendement_moteur_a * alimentation_pompe_nb * alimentation_tps_fonctionnement
        End If
        electricite_alimentation2 = 0
    Else
        'alimentation_Q_unitaire = inlet_Q / (alimentation_pompe_nb * alimentation_tps_fonctionnement)
        If alimentation_Q_unitaire_force = True Then
            alimentation_Q_unitaire = alimentation_Q_unitaire_f
            If alimentation_tps_fonctionnement_force = True Then
                'on ne fait rien
                If alimentation_pompe_nb_force = False Then
                    If Int(inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire)) < (inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire)) Then
                        alimentation_pompe_nb = Int(inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire)) + 1
                    Else
                        alimentation_pompe_nb = Int(inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire))
                    End If
                End If
            Else
                alimentation_tps_fonctionnement = inlet_Q / (alimentation_pompe_nb * alimentation_Q_unitaire)
                If alimentation_tps_fonctionnement > hyp_tps_fct_pompe_alim Then
                    alimentation_tps_fonctionnement = hyp_tps_fct_pompe_alim
                    If alimentation_pompe_nb_force = False Then
                        If Int(inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire)) < (inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire)) Then
                            alimentation_pompe_nb = Int(inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire)) + 1
                        Else
                            alimentation_pompe_nb = Int(inlet_Q / (alimentation_tps_fonctionnement * alimentation_Q_unitaire))
                        End If
                        alimentation_tps_fonctionnement = inlet_Q / (alimentation_pompe_nb * alimentation_Q_unitaire)
                    End If
                End If
            End If
        Else
            alimentation_Q_unitaire = inlet_Q / (alimentation_pompe_nb * alimentation_tps_fonctionnement)
        End If
        incoherence_a = (alimentation_Q_unitaire * alimentation_pompe_nb * alimentation_tps_fonctionnement - inlet_Q) / inlet_Q
        If Abs(incoherence_a) > critere_incoherence Then
            remarque(1) = Feuil38.Range("sechage_inos_erreur_pompage1").Cells(1, 1)
        End If
        puissance_pompe_a = alimentation_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * alimentation_P_refoulement / alimentation_rendement_pompe
        If puissance_pompe_a < 60 Then
            rendement_moteur_a = (5.7195 * Log(puissance_pompe_a) + 72.682) / 100
        Else
            rendement_moteur_a = 0.961
        End If
        If alimentation_rendement_global_force = True Then
            alimentation_rendement_global = alimentation_rendement_global_f
            electricite_alimentation1 = ((puissance_pompe_a * alimentation_rendement_pompe) / alimentation_rendement_global) * alimentation_pompe_nb * alimentation_tps_fonctionnement
        Else
            alimentation_rendement_global = rendement_moteur_a * alimentation_rendement_pompe
            electricite_alimentation1 = puissance_pompe_a / rendement_moteur_a * alimentation_pompe_nb * alimentation_tps_fonctionnement
        End If
        'alimentation2_Q_unitaire = inlet_Q / (alimentation2_pompe_nb * alimentation2_tps_fonctionnement)
        If alimentation2_Q_unitaire_force = True Then
            alimentation2_Q_unitaire = alimentation2_Q_unitaire_f
            If alimentation2_tps_fonctionnement_force = True Then
                'on ne fait rien
                If alimentation2_pompe_nb_force = False Then
                    If Int(inlet_Q / (alimentation2_tps_fonctionnement * alimentation2_Q_unitaire)) < (inlet_Q / (alimentation2_tps_fonctionnement * alimentation2_Q_unitaire)) Then
                        alimentation2_pompe_nb = Int(inlet_Q / (alimentation2_tps_fonctionnement * alimentation2_Q_unitaire)) + 1
                    Else
                        alimentation2_pompe_nb = Int(inlet_Q / (alimentation2_tps_fonctionnement * alimentation2_Q_unitaire))
                    End If
                End If
            Else
                alimentation2_tps_fonctionnement = inlet_Q / (alimentation2_pompe_nb * alimentation2_Q_unitaire)
                If alimentation2_tps_fonctionnement > hyp_tps_fct_pompe_alim Then
                    alimentation2_tps_fonctionnement = hyp_tps_fct_pompe_alim
                    If alimentation2_pompe_nb_force = False Then
                        If Int(inlet_Q / (alimentation2_tps_fonctionnement * alimentation2_Q_unitaire)) < (inlet_Q / (alimentation2_tps_fonctionnement * alimentation2_Q_unitaire)) Then
                            alimentation2_pompe_nb = Int(inlet_Q / (alimentation2_tps_fonctionnement * alimentation2_Q_unitaire)) + 1
                        Else
                            alimentation2_pompe_nb = Int(inlet_Q / (alimentation2_tps_fonctionnement * alimentation2_Q_unitaire))
                        End If
                        alimentation2_tps_fonctionnement = inlet_Q / (alimentation2_pompe_nb * alimentation2_Q_unitaire)
                    End If
                End If
            End If
        Else
            alimentation2_Q_unitaire = inlet_Q / (alimentation2_pompe_nb * alimentation2_tps_fonctionnement)
        End If
        incoherence_b = (alimentation2_Q_unitaire * alimentation2_pompe_nb * alimentation2_tps_fonctionnement - inlet_Q) / inlet_Q
        If Abs(incoherence_b) > critere_incoherence Then
            remarque(2) = Feuil38.Range("sechage_inos_erreur_pompage2").Cells(1, 1)
        End If
        puissance_pompe_b = alimentation2_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * alimentation2_P_refoulement / alimentation2_rendement_pompe
        If puissance_pompe_b < 60 Then
            rendement_moteur_b = (5.7195 * Log(puissance_pompe_b) + 72.682) / 100
        Else
            rendement_moteur_b = 0.961
        End If
        If alimentation2_rendement_global_force = True Then
            alimentation2_rendement_global = alimentation2_rendement_global_f
            electricite_alimentation2 = ((puissance_pompe_b * alimentation2_rendement_pompe) / alimentation2_rendement_global) * alimentation2_pompe_nb * alimentation2_tps_fonctionnement
        Else
            alimentation2_rendement_global = rendement_moteur_b * alimentation2_rendement_pompe
            electricite_alimentation2 = puissance_pompe_b / rendement_moteur_b * alimentation2_pompe_nb * alimentation2_tps_fonctionnement
        End If
    End If
    electricite_alimentation = electricite_alimentation1 + electricite_alimentation2


    electricite = electricite_specifique + electricite_alimentation

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "38.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub ecriture_resultats(ByVal numero_sechage_inos)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Ligne = numero_sechage_inos

    Feuil38.Range("sechage_inos_inlet1_ratio").Cells(Ligne, Colonne) = inlet_ratio_admis(1)
    Feuil38.Range("sechage_inos_inlet2_ratio").Cells(Ligne, Colonne) = inlet_ratio_admis(2)
    Feuil38.Range("sechage_inos_inlet_siccite").Cells(Ligne, Colonne) = inlet_siccite
    Feuil38.Range("sechage_inos_inlet_MES").Cells(Ligne, Colonne) = inlet_MES
    Feuil38.Range("sechage_inos_inlet_MV_MES").Cells(Ligne, Colonne) = inlet_MV_MES
    Feuil38.Range("sechage_inos_outlet_siccite").Cells(Ligne, Colonne) = outlet_siccite
    Feuil38.Range("sechage_inos_capture_MS").Cells(Ligne, Colonne) = capture_MS
    Feuil38.Range("sechage_inos_inlet_BH").Cells(Ligne, Colonne) = inlet_BH
    Feuil38.Range("sechage_inos_outlet_MES").Cells(Ligne, Colonne) = outlet_MES
    Feuil38.Range("sechage_inos_outlet_MV_MES").Cells(Ligne, Colonne) = outlet_MV_MES
    Feuil38.Range("sechage_inos_outlet_BH").Cells(Ligne, Colonne) = outlet_BH
    Feuil38.Range("sechage_inos_eau_evaporee").Cells(Ligne, Colonne) = eau_evaporee
    Feuil38.Range("sechage_inos_electricite").Cells(Ligne, Colonne) = electricite
    Feuil38.Range("sechage_inos_chaleur").Cells(Ligne, Colonne) = besoins_thermiques
    Feuil38.Range("sechage_inos_chauffage_T_out").Cells(Ligne, Colonne) = chauffage_T_out
    Feuil38.Range("sechage_inos_chauffage_T_in").Cells(Ligne, Colonne) = chauffage_T_in
    Feuil38.Range("sechage_inos_condensation_indirecte_Teau_in").Cells(Ligne, Colonne) = condensation_indirecte_Teau_in
    Feuil38.Range("sechage_inos_condensation_indirecte_Teau_out").Cells(Ligne, Colonne) = condensation_indirecte_Teau_out
    Feuil38.Range("sechage_inos_inlet_concentration").Cells(Ligne, Colonne) = inlet_concentration
    Feuil38.Range("sechage_inos_outlet_concentration").Cells(Ligne, Colonne) = intermediaire_concentration
    Feuil38.Range("sechage_inos_chaux").Cells(Ligne, Colonne) = chaux
    Feuil38.Range("sechage_inos_polymere").Cells(Ligne, Colonne) = polymere_kg_tMS
    Feuil38.Range("sechage_inos_fecl3").Cells(Ligne, Colonne) = FeCl3_kg_tMS

    Feuil38.Range("sechage_inos_alimentation_nb_pompe").Cells(Ligne, Colonne) = alimentation_pompe_nb
    Feuil38.Range("sechage_inos_alimentation_P_refoulement").Cells(Ligne, Colonne) = alimentation_P_refoulement
    Feuil38.Range("sechage_inos_alimentation_tps_fonctionnement").Cells(Ligne, Colonne) = alimentation_tps_fonctionnement
    Feuil38.Range("sechage_inos_alimentation_Q_unitaire").Cells(Ligne, Colonne) = alimentation_Q_unitaire
    Feuil38.Range("sechage_inos_alimentation_rendement_pompe").Cells(Ligne, Colonne) = alimentation_rendement_global
    Feuil38.Range("sechage_inos_alimentation2_nb_pompe").Cells(Ligne, Colonne) = alimentation2_pompe_nb
    Feuil38.Range("sechage_inos_alimentation2_P_refoulement").Cells(Ligne, Colonne) = alimentation2_P_refoulement
    Feuil38.Range("sechage_inos_alimentation2_tps_fonctionnement").Cells(Ligne, Colonne) = alimentation2_tps_fonctionnement
    Feuil38.Range("sechage_inos_alimentation2_Q_unitaire").Cells(Ligne, Colonne) = alimentation2_Q_unitaire
    Feuil38.Range("sechage_inos_alimentation2_rendement_pompe").Cells(Ligne, Colonne) = alimentation2_rendement_global

    If choix_condensation <> condensation_directe Then
        Feuil38.Range("sechage_inos_condensat_indirect_ratio_vapeur_condensee").Cells(Ligne, Colonne) = condensat_indirect_ratio_vapeur_condensee
        Feuil38.Range("sechage_inos_condensat_indirect_recup_eau_HT").Cells(Ligne, Colonne) = condensat_indirect_recup_eau_HT
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "38.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub effacer_resultats(ByVal numero_sechage_inos)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Ligne = numero_sechage_inos

    Feuil38.Range("sechage_inos_inlet1_ratio").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_inlet1_ratio").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_inlet2_ratio").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_inlet_siccite").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_inlet_MES").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_inlet_MV_MES").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_outlet_siccite").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_capture_MS").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_inlet_BH").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_outlet_MES").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_outlet_MV_MES").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_outlet_BH").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_eau_evaporee").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_electricite").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_chaleur").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_chauffage_T_out").Cells(Ligne, Colonne) = ""     'chauffage_T_out
    Feuil38.Range("sechage_inos_chauffage_T_in").Cells(Ligne, Colonne) = ""      'chauffage_T_in
    Feuil38.Range("sechage_inos_condensation_indirecte_Teau_in").Cells(Ligne, Colonne) = ""    ' condensation_indirecte_Teau_in
    Feuil38.Range("sechage_inos_condensation_indirecte_Teau_out").Cells(Ligne, Colonne) = ""    ' condensation_indirecte_Teau_out
    Feuil38.Range("sechage_inos_inlet_concentration").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_outlet_concentration").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_chaux").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_polymere").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_fecl3").Cells(Ligne, Colonne) = ""

    Feuil38.Range("sechage_inos_alimentation_nb_pompe").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_alimentation_P_refoulement").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_alimentation_tps_fonctionnement").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_alimentation_Q_unitaire").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_alimentation_rendement_pompe").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_alimentation2_nb_pompe").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_alimentation2_P_refoulement").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_alimentation2_tps_fonctionnement").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_alimentation2_Q_unitaire").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_alimentation2_rendement_pompe").Cells(Ligne, Colonne) = ""
    Feuil38.Range("sechage_inos_condensat_indirect_ratio_vapeur_condensee").Cells(Ligne, Colonne) = ""   'condensat_indirect_ratio_vapeur_condensee
    Feuil38.Range("sechage_inos_condensat_indirect_recup_eau_HT").Cells(Ligne, Colonne) = ""     'condensat_indirect_recup_eau_HT


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "38.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub
Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 2  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil38.Range("sechage_inos_erreur_nom") & " " & stockage_numero_sechage

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "38.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub Class_Terminate()



    If pn_EffacementResultatsEnCours = False Then
        Call affich_remarque
    End If



End Sub












