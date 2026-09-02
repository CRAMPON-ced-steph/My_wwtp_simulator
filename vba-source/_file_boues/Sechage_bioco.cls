Attribute VB_Name = "Sechage_bioco"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'constantes pour le choix du performance
Private Const nb_type_performance = 2  'nombre de types
Private Const sechage_partiel = 1
Private Const sechage_total = 2

'constantes pour le choix d'une bâche des boues à sécher
Private Const bache_amont_non = 1
Private Const bache_amont_1 = 2
Private Const bache_amont_2 = 3

Private Const nb_inlet = 2

Private Const critere_incoherence = 0.1

'mode de couverture des besoins thermiques
Private Const source_vapeur = 1   'si basse température uniquement
Private Const source_biogaz = 2
Private Const source_combustible = 3

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
Private Const inlet_deshydratation_1 = 13
Private Const inlet_deshydratation_2 = 14
Private Const inlet_deshydratation_3 = 15
Private Const inlet_chaulage_1 = 16
Private Const inlet_chaulage_2 = 17
Private Const inlet_chaulage_3 = 18

'variables de choix
'Dim choix_performance As Integer
Dim choix_inlet(nb_inlet) As Integer
Dim choix_bache_amont As Integer
Dim choix_source_energie As Integer
Dim choix_condensation As Integer
Dim choix_performance As Integer

'entrée
Dim inlet_siccite As Double
Public inlet_MES As Double
Dim inlet_Q As Double
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

'variables relatives à l'énergie thermique
Dim condensat_indirect_ratio_vapeur_condensee As Double
Dim condensat_indirect_recup_eau_HT As Double     'énergie récupérée de type eau à haute température
Dim condensat_indirect_ratio_vapeur_condensee_force As Boolean
Dim condensat_indirect_recup_eau_HT_force As Boolean     'énergie récupérée de type eau à haute température
Dim condensat_indirect_ratio_vapeur_condensee_f As Double
Dim condensat_indirect_recup_eau_HT_f As Double     'énergie récupérée de type eau à haute température

'températures
Dim FluideSechageTch As Double
Dim FluideSechageTch_force As Boolean
Dim FluideSechageTch_f As Double
Dim FluideSechageTfr As Double
Dim FluideSechageTfr_force As Boolean
Dim FluideSechageTfr_f As Double
Dim FluideCondensationTch As Double
Dim FluideCondensationTch_force As Boolean
Dim FluideCondensationTch_f As Double
Dim FluideCondensationTfr As Double
Dim FluideCondensationTfr_force As Boolean
Dim FluideCondensationTfr_f As Double

'variables à échanger avec l'interface
Dim inlet_ratio_admis(nb_inlet) As Double
Dim alimentation_pompe_nb As Integer
Dim alimentation_P_refoulement As Double
Dim alimentation_tps_fonctionnement As Double
Dim alimentation_Q_unitaire As Double
Dim alimentation_rendement_pompe As Double
Dim alimentation2_pompe_nb As Integer
Dim alimentation2_P_refoulement As Double
Dim alimentation2_tps_fonctionnement As Double
Dim alimentation2_Q_unitaire As Double
Dim alimentation2_rendement_pompe As Double
Dim alimentation_rendement_global As Double
Dim alimentation2_rendement_global As Double
Public electricite As Double

'déclaration des variables relatives aux valeurs renseignées (cases bleues)
'entrée
Dim sechage_bioco_inlet1_ratio As Double
Dim sechage_bioco_inlet2_ratio As Double
Dim sechage_bioco_inlet1_ratio_force As Boolean
Dim sechage_bioco_inlet2_ratio_force As Boolean
Dim sechage_bioco_inlet1_ratio_f As Integer
Dim sechage_bioco_inlet2_ratio_f As Integer
Dim inlet_siccite_force As Boolean
Dim inlet_MES_force As Boolean
Dim inlet_MV_MES_force As Boolean
Dim outlet_siccite_force As Boolean
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

'guide defaut

Dim inlet_siccite_f As Double
Dim inlet_MES_f As Double
Dim inlet_MV_MES_f As Double
Dim outlet_siccite_f As Double
Dim outlet_siccite_c As Double
'hypothese
Dim ratio_elec_f As Double
Dim ratio_chal_f As Double
'calcul
Dim inlet_BH_f As Double
Dim outlet_MES_f As Double
Dim outlet_MV_MES_f As Double
Dim outlet_BH_f As Double
Dim eau_evaporee_f As Double
Dim electricite_specifique_f As Double
Dim besoins_thermiques_f As Double

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

Dim secheur_pas_necessaire As Boolean    'si la siccité en entrée permet l'autocombustibilité lorsqu'un incinérateur est placé derrière

'erreurs
Dim stockage_numero_sechage As Integer
Dim remarque(2) As String

Sub lecture_choix(ByVal numero_sechage_bioco)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer

    stockage_numero_sechage = numero_sechage_bioco

    'Boues en entrée de séchage


    choix_inlet(1) = Feuil35.Range("sechage_bioco_choix_inlet1").Cells(numero_sechage_bioco, pi_FiliereConsideree)
    inlet_ratio_admis(1) = Feuil35.Range("sechage_bioco_inlet1_ratio").Cells(numero_sechage_bioco, pi_FiliereConsideree)

    choix_inlet(2) = Feuil35.Range("sechage_bioco_choix_inlet2").Cells(numero_sechage_bioco, pi_FiliereConsideree)
    inlet_ratio_admis(2) = Feuil35.Range("sechage_bioco_inlet2_ratio").Cells(numero_sechage_bioco, pi_FiliereConsideree)

    'on retranscrit le choix réalisé dans la synthèse pour l'adapter à notre matrice
    For I = 1 To nb_inlet
        If choix_inlet(I) <> inlet_aucun Then
            If choix_inlet(I) <= inlet_graisses Then
                inlet_etape_anterieure(I) = boues_inlet
                inlet_type_boues(I) = choix_inlet(I)
            Else
                inlet_etape_anterieure(I) = choix_inlet(I) - inlet_graisses + 1
                'ici on prendra directement les valeurs du flux antérieur
            End If
        End If
    Next I

    'lecture des choix internes
    choix_performance = Feuil35.Range("sechage_bioco_choix_performance").Cells(numero_sechage_bioco, pi_FiliereConsideree)
    choix_bache_amont = Feuil35.Range("sechage_bioco_choix_bache_amont").Cells(numero_sechage_bioco, pi_FiliereConsideree)
    choix_source_energie = Feuil35.Range("sechage_bioco_choix_source_energie").Cells(numero_sechage_bioco, pi_FiliereConsideree)
    choix_condensation = Feuil35.Range("sechage_bioco_choix_condensation").Cells(numero_sechage_bioco, pi_FiliereConsideree)
    If choix_performance = sechage_total Then
        If choix_incineration = True Then
            MsgBox Feuil35.Range("sechage_bioco_erreur_nom") & " " & stockage_numero_sechage & ": " & Feuil35.Range("sechage_bioco_erreur_incineration").Cells(1, 1), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
            pn_EffacementResultatsEnCours = True
            pn_ErreurNonGeree = True
            Exit Sub
        End If
    End If
    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "35.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'ratios de consommation électrique et thermique
    ratio_elec = 43     'kWh/ TEE
    ratio_chal = 880        'kWh/ TEE

    'siccité des boues en sortie (en fonction de la performance du sécheur)
    If choix_performance = sechage_partiel Then
        outlet_siccite_c = 0.65
    Else
        outlet_siccite_c = 0.9      '%
    End If

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
    NumeroErreurNonGeree = "35.3"

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

    'Température du fluide de séchage
    If FluideSechageTch_force = True Then
        FluideSechageTch = FluideSechageTch_f
    Else
        FluideSechageTch = 170
    End If
    If FluideSechageTfr_force = True Then
        FluideSechageTfr = FluideSechageTfr_f
    Else
        FluideSechageTfr = 50
    End If

    'condensation de l'eau évaporée
    If condensat_indirect_ratio_vapeur_condensee_force = True Then
        condensat_indirect_ratio_vapeur_condensee = condensat_indirect_ratio_vapeur_condensee_f
    Else
        condensat_indirect_ratio_vapeur_condensee = 0.5   '50% de l'eau évaporée condensée dans la partie indirecte
    End If
    If choix_condensation <> condensation_directe Then
        If FluideCondensationTch_force = True Then
            FluideCondensationTch = FluideCondensationTch_f
        Else
            FluideCondensationTch = 90
        End If
        If FluideCondensationTfr_force = True Then
            FluideCondensationTfr = FluideCondensationTfr_f
        Else
            FluideCondensationTfr = 70
        End If
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "35.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub lecture_valeurs_forcees(ByVal numero_sechage_bioco)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer
    Dim Ligne As Integer
    
    Ligne = numero_sechage_bioco
    Colonne = pi_FiliereConsideree

    inlet_siccite_force = Feuil35.Range("sechage_bioco_inlet_siccite").Cells(Ligne, Colonne)
    inlet_MES_force = Feuil35.Range("sechage_bioco_inlet_MES").Cells(Ligne, Colonne)
    inlet_MV_MES_force = Feuil35.Range("sechage_bioco_inlet_MV_MES").Cells(Ligne, Colonne)
    outlet_siccite_force = Feuil35.Range("sechage_bioco_outlet_siccite").Cells(Ligne, Colonne)
    'calcul
    inlet_BH_force = Feuil35.Range("sechage_bioco_inlet_BH").Cells(Ligne, Colonne)
    outlet_MES_force = Feuil35.Range("sechage_bioco_outlet_MES").Cells(Ligne, Colonne)
    outlet_MV_MES_force = Feuil35.Range("sechage_bioco_outlet_MV_MES").Cells(Ligne, Colonne)
    outlet_BH_force = Feuil35.Range("sechage_bioco_outlet_BH").Cells(Ligne, Colonne)
    besoins_thermiques_force = Feuil35.Range("sechage_bioco_chaleur").Cells(Ligne, Colonne)
        FluideSechageTch_force = Feuil35.Range("sechage_bioco_FluideSechageTch").Cells(Ligne, Colonne)
        FluideSechageTfr_force = Feuil35.Range("sechage_bioco_FluideSechageTfr").Cells(Ligne, Colonne)

    alimentation_pompe_nb_force = Feuil35.Range("sechage_bioco_alimentation_nb_pompe").Cells(Ligne, Colonne)
    alimentation_P_refoulement_force = Feuil35.Range("sechage_bioco_alimentation_P_refoulement").Cells(Ligne, Colonne)
    alimentation_tps_fonctionnement_force = Feuil35.Range("sechage_bioco_alimentation_tps_fonctionnement").Cells(Ligne, Colonne)
    alimentation_Q_unitaire_force = Feuil35.Range("sechage_bioco_alimentation_Q_unitaire").Cells(Ligne, Colonne)
    alimentation_rendement_global_force = Feuil35.Range("sechage_bioco_alimentation_rendement_pompe").Cells(Ligne, Colonne)

    alimentation2_pompe_nb_force = Feuil35.Range("sechage_bioco_alimentation2_nb_pompe").Cells(Ligne, Colonne)
    alimentation2_P_refoulement_force = Feuil35.Range("sechage_bioco_alimentation2_P_refoulement").Cells(Ligne, Colonne)
    alimentation2_tps_fonctionnement_force = Feuil35.Range("sechage_bioco_alimentation2_tps_fonctionnement").Cells(Ligne, Colonne)
    alimentation2_Q_unitaire_force = Feuil35.Range("sechage_bioco_alimentation2_Q_unitaire").Cells(Ligne, Colonne)
    alimentation2_rendement_global_force = Feuil35.Range("sechage_bioco_alimentation2_rendement_pompe").Cells(Ligne, Colonne)

    If choix_condensation <> condensation_directe Then
        condensat_indirect_ratio_vapeur_condensee_force = Feuil35.Range("sechage_bioco_condensat_indirect_ratio_vapeur_condensee").Cells(Ligne, Colonne)
        condensat_indirect_recup_eau_HT_force = Feuil35.Range("sechage_bioco_condensat_indirect_recup_eau_HT").Cells(Ligne, Colonne)
        FluideCondensationTch_force = Feuil35.Range("sechage_bioco_FluideCondensationTch").Cells(Ligne, Colonne)
        FluideCondensationTfr_force = Feuil35.Range("sechage_bioco_FluideCondensationTfr").Cells(Ligne, Colonne)
    End If

    If alimentation_pompe_nb_force = True Then
        alimentation_pompe_nb_f = Feuil35.Range("sechage_bioco_alimentation_nb_pompe").Cells(Ligne, Colonne)
    End If
    If alimentation_P_refoulement_force = True Then
        alimentation_P_refoulement_f = Feuil35.Range("sechage_bioco_alimentation_P_refoulement").Cells(Ligne, Colonne)
    End If
    If alimentation_tps_fonctionnement_force = True Then
        alimentation_tps_fonctionnement_f = Feuil35.Range("sechage_bioco_alimentation_tps_fonctionnement").Cells(Ligne, Colonne)
    End If
    If alimentation_Q_unitaire_force = True Then
        alimentation_Q_unitaire_f = Feuil35.Range("sechage_bioco_alimentation_Q_unitaire").Cells(Ligne, Colonne)
    End If
    If alimentation_rendement_global_force = True Then
        alimentation_rendement_global_f = Feuil35.Range("sechage_bioco_alimentation_rendement_pompe").Cells(Ligne, Colonne)
    End If

    If alimentation2_pompe_nb_force = True Then
        alimentation2_pompe_nb_f = Feuil35.Range("sechage_bioco_alimentation2_nb_pompe").Cells(Ligne, Colonne)
    End If
    If alimentation2_P_refoulement_force = True Then
        alimentation2_P_refoulement_f = Feuil35.Range("sechage_bioco_alimentation2_P_refoulement").Cells(Ligne, Colonne)
    End If
    If alimentation2_tps_fonctionnement_force = True Then
        alimentation2_tps_fonctionnement_f = Feuil35.Range("sechage_bioco_alimentation2_tps_fonctionnement").Cells(Ligne, Colonne)
    End If
    If alimentation2_Q_unitaire_force = True Then
        alimentation2_Q_unitaire_f = Feuil35.Range("sechage_bioco_alimentation2_Q_unitaire").Cells(Ligne, Colonne)
    End If
    If alimentation2_rendement_global_force = True Then
        alimentation2_rendement_global_f = Feuil35.Range("sechage_bioco_alimentation2_rendement_pompe").Cells(Ligne, Colonne)
    End If

    If inlet_MES_force = True Then
        inlet_MES_f = Feuil35.Range("sechage_bioco_inlet_MES").Cells(Ligne, Colonne)
    End If

    If inlet_siccite_force = True Then
        inlet_siccite_f = Feuil35.Range("sechage_bioco_inlet_siccite").Cells(Ligne, Colonne)
    End If

    If inlet_MV_MES_force = True Then
        inlet_MV_MES_f = Feuil35.Range("sechage_bioco_inlet_MV_MES").Cells(Ligne, Colonne)
    End If

    If outlet_siccite_force = True Then
        outlet_siccite_f = Feuil35.Range("sechage_bioco_outlet_siccite").Cells(Ligne, Colonne)
    End If

    If inlet_BH_force = True Then
        inlet_BH_f = Feuil35.Range("sechage_bioco_inlet_BH").Cells(Ligne, Colonne)
    End If

    If outlet_MES_force = True Then
        outlet_MES_f = Feuil35.Range("sechage_bioco_outlet_MES").Cells(Ligne, Colonne)
    End If

    If outlet_MV_MES_force = True Then
        outlet_MV_MES_f = Feuil35.Range("sechage_bioco_outlet_MV_MES").Cells(Ligne, Colonne)
    End If

    If outlet_BH_force = True Then
        outlet_BH_f = Feuil35.Range("sechage_bioco_outlet_BH").Cells(Ligne, Colonne)
    End If

    'If electricite_specifique_force = True Then
    '    electricite_specifique_f = Feuil35.Range("sechage_bioco_electricite_specifique").Cells(ligne, colonne)
    'End If

    If besoins_thermiques_force = True Then
        besoins_thermiques_f = Feuil35.Range("sechage_bioco_chaleur").Cells(Ligne, Colonne)
    End If
    If FluideSechageTch_force = True Then
        FluideSechageTch_f = Feuil35.Range("sechage_bioco_FluideSechageTch").Cells(Ligne, Colonne)
    End If
    If FluideSechageTfr_force = True Then
        FluideSechageTfr_f = Feuil35.Range("sechage_bioco_FluideSechageTfr").Cells(Ligne, Colonne)
    End If
    
    'énergie thermique
    If condensat_indirect_ratio_vapeur_condensee_force = True Then
        condensat_indirect_ratio_vapeur_condensee_f = Feuil35.Range("sechage_bioco_condensat_indirect_ratio_vapeur_condensee").Cells(Ligne, Colonne)
    End If
    If condensat_indirect_recup_eau_HT_force = True Then
        condensat_indirect_recup_eau_HT_f = Feuil35.Range("sechage_bioco_condensat_indirect_recup_eau_HT").Cells(Ligne, Colonne)
    End If
    If FluideCondensationTch_force = True Then
        FluideCondensationTch_f = Feuil35.Range("sechage_bioco_FluideCondensationTch").Cells(Ligne, Colonne)
    End If
    If FluideCondensationTfr_force = True Then
        FluideCondensationTfr_f = Feuil35.Range("sechage_bioco_FluideCondensationTfr").Cells(Ligne, Colonne)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "35.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub dimensionnement(ByRef boues_flux, ByVal numero_sechage_bioco, ByRef retour_flux, ByRef boues_pollution_soluble, ByRef besoins_th, ByRef disponibilites_thermiques)

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
    Dim outlet_concentration As Double
    Dim soluble_kg(nb_repere_mgL) As Double
    Dim condensat_direct_eau_inlet_kgj As Double

    'gestion des ratios de MV réduits par type de boues
    Dim MV_reduits(nb_type_boues) As Double
    Dim delta_MV_reduits As Double

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
                    inlet_NK = inlet_NK + intermediaire_NK + soluble_kg(repere_mgL_NK)
                    inlet_Pt = inlet_Pt + intermediaire_Pt + soluble_kg(repere_mgL_Pt)
                    inlet_MV = inlet_MV + intermediaire_MV

                    'SEPARATION
                    intermediaire_MES = intermediaire_MES
                    intermediaire_MV = intermediaire_MV
                    intermediaire_DCO = intermediaire_DCO
                    intermediaire_NK = intermediaire_NK
                    intermediaire_Pt = intermediaire_Pt
                    'intermediaire_Q = intermediaire_MES / outlet_concentration

                    'variables de sortie
                    outlet_MES = outlet_MES + intermediaire_MES
                    outlet_DCO = outlet_DCO + intermediaire_DCO
                    outlet_NK = outlet_NK + intermediaire_NK
                    outlet_Pt = outlet_Pt + intermediaire_Pt
                    'outlet_Q = outlet_Q + intermediaire_Q

                    'je calcule les paramètres des boues séchees
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MES) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MES) + intermediaire_MES
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_Q) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_Q) + intermediaire_MES / outlet_siccite / 1000
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MV_MES) = intermediaire_MV / boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MES)

                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_flux_in) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_flux_in)

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
                        inlet_NK = inlet_NK + intermediaire_NK + soluble_kg(repere_mgL_NK)
                        inlet_Pt = inlet_Pt + intermediaire_Pt + soluble_kg(repere_mgL_Pt)
                        inlet_MV = inlet_MV + intermediaire_MV

                        'SEPARATION
                        intermediaire_MES = intermediaire_MES
                        intermediaire_MV = intermediaire_MV
                        intermediaire_DCO = intermediaire_DCO
                        intermediaire_NK = intermediaire_NK
                        intermediaire_Pt = intermediaire_Pt
                        'intermediaire_Q = intermediaire_MES / outlet_concentration

                        'variables de sortie
                        outlet_MES = outlet_MES + intermediaire_MES
                        outlet_DCO = outlet_DCO + intermediaire_DCO
                        outlet_NK = outlet_NK + intermediaire_NK
                        outlet_Pt = outlet_Pt + intermediaire_Pt
                        'outlet_Q = outlet_Q + intermediaire_Q

                        'je calcule les paramètres des boues séchees
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MES) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MES) + intermediaire_MES
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_Q) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_Q) + intermediaire_MES / outlet_siccite / 1000
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MV_MES) = intermediaire_MV / boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MES)
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MES)
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MES)
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MES)

                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_flux_in) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_verif_flux) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_flux_in)

                        'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                        boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                    End If
                Next j

            End If

        End If

    Next I

    If inlet_MES = 0 Or inlet_Q = 0 Then
        MsgBox Feuil5.Range("erreur_choix_sechage_bioco_inlet") & " " & numero_sechage_bioco, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If

    'BOUES EN ENTREE
    inlet_siccite = inlet_MES / (inlet_Q * masse_volumique_boues)
    inlet_BH = inlet_MES / inlet_siccite
    inlet_MV_MES = inlet_MV / inlet_MES


    eau_evaporee = ((inlet_MES * (1 - inlet_siccite) / inlet_siccite) - (outlet_MES * (1 - outlet_siccite) / outlet_siccite))
    If besoins_thermiques_force = True Then
        besoins_thermiques = besoins_thermiques_f
    Else
        besoins_thermiques = eau_evaporee / 1000 * ratio_chal
    End If


    'BOUES EN SORTIE
    For I = 1 To 2
        If incineration_repere_boues_inlet(I) = boues_sechees_bioco(numero_sechage_bioco) Then
            If incineration_siccite_ajustable(I) = True Then
                If outlet_siccite_force = True Then
                    outlet_siccite = outlet_siccite_f
                    incineration_siccite_ajustable(I) = False
                Else
                    If choix_performance = sechage_partiel Then
                        If incineration_siccite_autocombustibilite(I) < inlet_siccite Then
                            outlet_siccite = inlet_siccite
                            secheur_pas_necessaire = True
                            incineration_siccite_ajustable(I) = False
                        Else
                            outlet_siccite = incineration_siccite_autocombustibilite(I)
                        End If
                    Else
                        incineration_siccite_ajustable(I) = False
                    End If
                End If
            ElseIf secheur_pas_necessaire = True Then
                outlet_siccite = inlet_siccite
            End If
        End If
    Next I


    outlet_BH = outlet_MES / outlet_siccite

    outlet_MV_MES = inlet_MV_MES

    'EAU EVAPOREE
    eau_evaporee = ((inlet_MES * (1 - inlet_siccite) / inlet_siccite) - (outlet_MES * (1 - outlet_siccite) / outlet_siccite))


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
        If choix_source_energie = source_vapeur Then
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
    NumeroErreurNonGeree = "35.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub
Sub transfert_flux_energie_vers_pinch(ByVal numero_secheur_bioco As Integer)

    Dim nom_flux As String
    Dim T_in As Double
    Dim T_out As Double
    Dim MCp As Double

    'SECHAGE BIOCO - CHAUFFAGE AIR
    'If choix_source_energie = source_vapeur Then

    'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
    nom_flux = Feuil5.Range("nom_flux_bioco1").Cells(1, 1).Value & " #" & numero_secheur_bioco
    T_in = FluideSechageTfr    '°C - hyp
    T_out = FluideSechageTch    '°C - hyp
    MCp = besoins_thermiques / (T_out - T_in)  'kWh/j/°C

    'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
    Call copie_dans_flux_thermique_process(nom_flux, T_in, T_out, MCp)

    '    Else
    '        '************* PENSER A RAJOUTER LES BESOINS AVANT LE PINCH DU BIOGAZ ***************
    '    End If

    'SECHAGE BIOCO - CONDENSATION INDIRECTE
    If choix_condensation <> condensation_directe Then

        'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
        '        nom_flux = "SECHAGE BIOCO " & numero_secheur_bioco & " - CONDENSATION INDIRECTE"
        nom_flux = Feuil5.Range("nom_flux_bioco2").Cells(1, 1).Value & " #" & numero_secheur_bioco
        T_in = FluideCondensationTch    '°C - hyp
        T_out = FluideCondensationTfr    '°C - hyp
        MCp = condensat_indirect_recup_eau_HT / (T_in - T_out)  'kWh/j/°C

        'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
        Call copie_dans_flux_thermique_process(nom_flux, T_in, T_out, MCp)
    End If

End Sub

Sub bilan_thermique(ByRef besoins_th, ByRef disponibilites_thermiques)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'énergie thermique nécessaire
    besoins_thermiques = eau_evaporee / 1000 * ratio_chal

    'Besoins thermiques et disponibilites
    If choix_source_energie = source_vapeur Then
        besoins_th(energie_vapeur_sat2) = besoins_th(energie_vapeur_sat2) + besoins_thermiques    'attention si plusieurs sécheurs
    ElseIf choix_source_energie = source_biogaz Then
        besoins_th(energie_biogaz) = besoins_th(energie_biogaz) + besoins_thermiques
    Else
        besoins_th(energie_combustible) = besoins_th(energie_combustible) + besoins_thermiques
    End If
    disponibilites_thermiques(energie_eau_HT) = disponibilites_thermiques(energie_eau_HT) + condensat_indirect_recup_eau_HT


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "26.11"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub
Sub fonctionnement_reel(ByRef boues_flux, ByVal numero_sechage_bioco, ByRef retour_flux, ByRef boues_pollution_soluble, ByRef besoins_th, ByRef disponibilites_thermiques)

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
    Dim outlet_concentration As Double
    Dim soluble_kg(nb_repere_mgL) As Double
    Dim condensat_direct_eau_inlet_kgj As Double

    'gestion des ratios de MV réduits par type de boues
    Dim MV_reduits(nb_type_boues) As Double
    Dim delta_MV_reduits As Double

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
                    inlet_NK = inlet_NK + intermediaire_NK + soluble_kg(repere_mgL_NK)
                    inlet_Pt = inlet_Pt + intermediaire_Pt + soluble_kg(repere_mgL_Pt)
                    inlet_MV = inlet_MV + intermediaire_MV

                    'SEPARATION
                    intermediaire_MES = intermediaire_MES
                    intermediaire_MV = intermediaire_MV
                    intermediaire_DCO = intermediaire_DCO
                    intermediaire_NK = intermediaire_NK
                    intermediaire_Pt = intermediaire_Pt
                    'intermediaire_Q = intermediaire_MES / outlet_concentration

                    'variables de sortie
                    outlet_MES = outlet_MES + intermediaire_MES
                    outlet_DCO = outlet_DCO + intermediaire_DCO
                    outlet_NK = outlet_NK + intermediaire_NK
                    outlet_Pt = outlet_Pt + intermediaire_Pt
                    'outlet_Q = outlet_Q + intermediaire_Q

                    'je calcule les paramètres des boues séchees
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MES) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MES) + intermediaire_MES
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_Q) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_Q) + intermediaire_MES / outlet_siccite / 1000
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MV_MES) = intermediaire_MV / boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_MES)

                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_flux_in) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), inlet_type_boues(I), repere_flux_in)

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
                        inlet_NK = inlet_NK + intermediaire_NK + soluble_kg(repere_mgL_NK)
                        inlet_Pt = inlet_Pt + intermediaire_Pt + soluble_kg(repere_mgL_Pt)
                        inlet_MV = inlet_MV + intermediaire_MV

                        'SEPARATION
                        intermediaire_MES = intermediaire_MES
                        intermediaire_MV = intermediaire_MV
                        intermediaire_DCO = intermediaire_DCO
                        intermediaire_NK = intermediaire_NK
                        intermediaire_Pt = intermediaire_Pt
                        'intermediaire_Q = intermediaire_MES / outlet_concentration

                        'variables de sortie
                        outlet_MES = outlet_MES + intermediaire_MES
                        outlet_DCO = outlet_DCO + intermediaire_DCO
                        outlet_NK = outlet_NK + intermediaire_NK
                        outlet_Pt = outlet_Pt + intermediaire_Pt
                        'outlet_Q = outlet_Q + intermediaire_Q

                        'je calcule les paramètres des boues séchees
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MES) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MES) + intermediaire_MES
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_Q) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_Q) + intermediaire_MES / outlet_siccite / 1000
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MV_MES) = intermediaire_MV / boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MES)
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MES)
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MES)
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_MES)

                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_flux_in) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                        boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_verif_flux) = boues_flux(boues_sechees_bioco(numero_sechage_bioco), j, repere_flux_in)

                        'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                        boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                    End If
                Next j

            End If

        End If
    Next I

    'BOUES EN ENTREE
    inlet_siccite = inlet_MES / (inlet_Q * masse_volumique_boues)
    inlet_BH = inlet_MES / inlet_siccite
    inlet_MV_MES = inlet_MV / inlet_MES

    'BOUES EN SORTIE
    For I = 1 To 2
        If incineration_repere_boues_inlet(I) = boues_sechees_bioco(numero_sechage_bioco) Then
            If incineration_siccite_ajustable(I) = True Then
                If outlet_siccite_force = True Then
                    outlet_siccite = outlet_siccite_f
                    incineration_siccite_ajustable(I) = False
                Else
                    If choix_performance = sechage_partiel Then
                        If incineration_siccite_autocombustibilite(I) < inlet_siccite Then
                            outlet_siccite = inlet_siccite
                            secheur_pas_necessaire = True
                            incineration_siccite_ajustable(I) = False
                        Else
                            outlet_siccite = incineration_siccite_autocombustibilite(I)
                        End If
                    Else
                        incineration_siccite_ajustable(I) = False
                    End If
                End If
            ElseIf secheur_pas_necessaire = True Then
                outlet_siccite = inlet_siccite
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

    eau_evaporee = ((inlet_MES * (1 - inlet_siccite) / inlet_siccite) - (outlet_MES * (1 - outlet_siccite) / outlet_siccite))

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
        If choix_source_energie = source_vapeur Then
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
    NumeroErreurNonGeree = "35.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub calcul_consommation_electrique()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim rendement_pompe As Double
    Dim puissance_pompe_a As Double   'puissance pompe alimentation
    Dim puissance_pompe As Double    'puissance pompe extraction
    Dim puissance_pompe_b As Double   'puissance pompe alimentation
    Dim rendement_moteur_a As Double
    Dim rendement_moteur As Double
    Dim rendement_moteur_b As Double
    Dim incoherence_a As Double
    Dim incoherence As Double
    Dim incoherence_b As Double
    Dim electricite_alimentation1 As Double
    Dim electricite_alimentation2 As Double
    Dim electricite_specifique As Double
    Dim electricite_alimentation As Double



    electricite_specifique = eau_evaporee / 1000 * ratio_elec


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
            remarque(1) = Feuil35.Range("sechage_bioco_erreur_pompage1").Cells(1, 1)
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
            remarque(1) = Feuil35.Range("sechage_bioco_erreur_pompage1").Cells(1, 1)
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
            remarque(2) = Feuil35.Range("sechage_bioco_erreur_pompage2").Cells(1, 1)
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
    NumeroErreurNonGeree = "35.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True






End Sub

Sub ecriture_resultats(ByVal numero_sechage_bioco)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Ligne = numero_sechage_bioco

    Feuil35.Range("sechage_bioco_inlet1_ratio").Cells(Ligne, Colonne) = inlet_ratio_admis(1)
    Feuil35.Range("sechage_bioco_inlet2_ratio").Cells(Ligne, Colonne) = inlet_ratio_admis(2)
    Feuil35.Range("sechage_bioco_inlet_siccite").Cells(Ligne, Colonne) = inlet_siccite
    Feuil35.Range("sechage_bioco_inlet_MES").Cells(Ligne, Colonne) = inlet_MES
    Feuil35.Range("sechage_bioco_inlet_MV_MES").Cells(Ligne, Colonne) = inlet_MV_MES
    Feuil35.Range("sechage_bioco_outlet_siccite").Cells(Ligne, Colonne) = outlet_siccite
    Feuil35.Range("sechage_bioco_inlet_BH").Cells(Ligne, Colonne) = inlet_BH
    Feuil35.Range("sechage_bioco_outlet_MES").Cells(Ligne, Colonne) = outlet_MES
    Feuil35.Range("sechage_bioco_outlet_MV_MES").Cells(Ligne, Colonne) = outlet_MV_MES
    Feuil35.Range("sechage_bioco_outlet_BH").Cells(Ligne, Colonne) = outlet_BH
    Feuil35.Range("sechage_bioco_eau_evaporee").Cells(Ligne, Colonne) = eau_evaporee
    Feuil35.Range("sechage_bioco_electricite").Cells(Ligne, Colonne) = electricite
    Feuil35.Range("sechage_bioco_chaleur").Cells(Ligne, Colonne) = besoins_thermiques
        Feuil35.Range("sechage_bioco_FluideSechageTch").Cells(Ligne, Colonne) = FluideSechageTch
        Feuil35.Range("sechage_bioco_FluideSechageTfr").Cells(Ligne, Colonne) = FluideSechageTfr

    Feuil35.Range("sechage_bioco_alimentation_nb_pompe").Cells(Ligne, Colonne) = alimentation_pompe_nb
    Feuil35.Range("sechage_bioco_alimentation_P_refoulement").Cells(Ligne, Colonne) = alimentation_P_refoulement
    Feuil35.Range("sechage_bioco_alimentation_tps_fonctionnement").Cells(Ligne, Colonne) = alimentation_tps_fonctionnement
    Feuil35.Range("sechage_bioco_alimentation_Q_unitaire").Cells(Ligne, Colonne) = alimentation_Q_unitaire
    Feuil35.Range("sechage_bioco_alimentation_rendement_pompe").Cells(Ligne, Colonne) = alimentation_rendement_global

    Feuil35.Range("sechage_bioco_alimentation2_nb_pompe").Cells(Ligne, Colonne) = alimentation2_pompe_nb
    Feuil35.Range("sechage_bioco_alimentation2_P_refoulement").Cells(Ligne, Colonne) = alimentation2_P_refoulement
    Feuil35.Range("sechage_bioco_alimentation2_tps_fonctionnement").Cells(Ligne, Colonne) = alimentation2_tps_fonctionnement
    Feuil35.Range("sechage_bioco_alimentation2_Q_unitaire").Cells(Ligne, Colonne) = alimentation2_Q_unitaire
    Feuil35.Range("sechage_bioco_alimentation2_rendement_pompe").Cells(Ligne, Colonne) = alimentation2_rendement_global

    If choix_condensation <> condensation_directe Then
        Feuil35.Range("sechage_bioco_condensat_indirect_ratio_vapeur_condensee").Cells(Ligne, Colonne) = condensat_indirect_ratio_vapeur_condensee
        Feuil35.Range("sechage_bioco_condensat_indirect_recup_eau_HT").Cells(Ligne, Colonne) = condensat_indirect_recup_eau_HT
        Feuil35.Range("sechage_bioco_FluideCondensationTch").Cells(Ligne, Colonne) = FluideCondensationTch
        Feuil35.Range("sechage_bioco_FluideCondensationTfr").Cells(Ligne, Colonne) = FluideCondensationTfr
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "35.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats(ByVal numero_sechage_bioco)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Ligne = numero_sechage_bioco

    Feuil35.Range("sechage_bioco_inlet1_ratio").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_inlet2_ratio").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_inlet_siccite").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_inlet_MES").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_inlet_MV_MES").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_outlet_siccite").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_inlet_BH").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_outlet_MES").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_outlet_MV_MES").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_outlet_BH").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_eau_evaporee").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_electricite").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_chaleur").Cells(Ligne, Colonne) = ""
        Feuil35.Range("sechage_bioco_FluideSechageTch").Cells(Ligne, Colonne) = ""
        Feuil35.Range("sechage_bioco_FluideSechageTfr").Cells(Ligne, Colonne) = ""

    Feuil35.Range("sechage_bioco_alimentation_nb_pompe").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_alimentation_P_refoulement").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_alimentation_tps_fonctionnement").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_alimentation_Q_unitaire").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_alimentation_rendement_pompe").Cells(Ligne, Colonne) = ""

    Feuil35.Range("sechage_bioco_alimentation2_nb_pompe").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_alimentation2_P_refoulement").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_alimentation2_tps_fonctionnement").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_alimentation2_Q_unitaire").Cells(Ligne, Colonne) = ""
    Feuil35.Range("sechage_bioco_alimentation2_rendement_pompe").Cells(Ligne, Colonne) = ""

    Feuil35.Range("sechage_bioco_condensat_indirect_ratio_vapeur_condensee").Cells(Ligne, Colonne) = ""   'condensat_indirect_ratio_vapeur_condensee
    Feuil35.Range("sechage_bioco_condensat_indirect_recup_eau_HT").Cells(Ligne, Colonne) = ""     'condensat_indirect_recup_eau_HT
        Feuil35.Range("sechage_bioco_FluideCondensationTch").Cells(Ligne, Colonne) = ""
        Feuil35.Range("sechage_bioco_FluideCondensationTfr").Cells(Ligne, Colonne) = ""
    
    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "35.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub
Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 2  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil35.Range("sechage_bioco_erreur_nom") & " " & stockage_numero_sechage

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "35.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub Class_Terminate()



    If pn_EffacementResultatsEnCours = False Then
        Call affich_remarque
    End If



End Sub







