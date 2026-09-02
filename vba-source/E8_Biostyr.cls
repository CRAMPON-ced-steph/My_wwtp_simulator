Option Explicit

Private Const critere_incoherence = 0.1

'constantes pour le choix du type de biostyr
Private Const nb_type_biostyr = 4  'nombre de types de biostyr
Private Const biostyr_C = 1  'carbone
Private Const biostyr_N = 2  'nitrifiant
Private Const biostyr_NDNc = 3  'nit-dénit classique
Private Const biostyr_NDNs = 4  'nit-dénit simultané

'constantes pour le choix du type de régulation d'air
Private Const nb_type_regulation = 4  'nombre de type de régulations
Private Const regulation_horloge = 1
Private Const regulation_sans_variation = 2
Private Const regulation_avec_variation = 3
Private Const regulation_avance = 4

'constantes pour le choix du type de surpresseur
Private Const nb_type_surpresseur = 3
Private Const surpresseur_roots = 1
Private Const surpresseur_vis = 2
Private Const surpresseur_turbo = 3
Private Const nb_surpresseur_gamme_Q = 1
Private Const surpresseur_gamme_Q_0_infini = 1
Private Const nb_surpresseur_gamme_P = 3
Private Const surpresseur_gamme_P_0_Px1 = 1
Private Const Px1 = 6    'mCE
Private Const surpresseur_gamme_P_Px1_Px2 = 2
Private Const Px2 = 8    'mCE
Private Const surpresseur_gamme_P_Px2_infini = 3

'choix
Dim choix_type As Integer
Dim choix_devenir_ES As Integer
Dim choix_surface_unitaire As Double   'cas particulier: donne directement la valeur choisie
Dim choix_diametre_media As Double   'cas particulier: donne directement la valeur choisie
Dim choix_regulation As Integer
Dim choix_surpresseur As Integer


'variables échangées avec excel
Dim hauteur_media As Double
Dim surface_filtration_min As Double
Dim nb_cellules As Integer
Dim surface_filtration As Double
Dim Cv_DCO As Double
Dim Cv_nit As Double
Dim sortie_DCO As Double    ' MSM 15/11/12
Dim sortie_NH4 As Double
Dim sortie_NO3 As Double
Dim surpresseur_conso_spec As Double
Dim O2_besoin As Double
Dim O2_rdt_transfert As Double
Dim air_Q_Nm3j As Double
Dim diffuseur_encrassement As Double
Dim air_P_refoulement As Double
Dim recirculation_taux As Double
Dim recirculation_nb_pompe As Integer
Dim recirculation_P_refoulement As Double
Dim recirculation_tps_fonctionnement As Double
Dim recirculation_Q_unitaire As Double
Dim recirculation_pompe_rdt As Double
Dim ratio_ES_volume_media As Double
Dim ES_nb_pompe As Integer
Dim ES_P_refoulement As Double
Dim ES_tps_fonctionnement As Double
Dim ES_Q_unitaire As Double
Dim ES_pompe_rdt As Double
Dim lavage_air_Q As Double
Dim lavage_air_P_refoulement As Double
Dim lavage_air_tps_fonctionnement As Double
Dim ES_concentration As Double
Dim ES_MV_MES As Double
Dim ES_MES As Double
Public electricite As Double
Dim electricite_aeration As Double

'variables forcées
Dim hauteur_media_force As Boolean
Dim surface_filtration_min_force As Boolean
Dim nb_cellules_force As Boolean
Dim Cv_DCO_force As Boolean
Dim Cv_nit_force As Boolean
Dim sortie_DCO_force As Boolean    'MSM 15/11/12
Dim sortie_NH4_force As Boolean
Dim sortie_NO3_force As Boolean
Dim surpresseur_conso_spec_force As Boolean
Dim O2_besoin_force As Boolean
Dim O2_rdt_transfert_force As Boolean
Dim air_Q_Nm3j_force As Boolean
Dim diffuseur_encrassement_force As Boolean
Dim air_P_refoulement_force As Boolean
Dim recirculation_taux_force As Boolean
Dim recirculation_nb_pompe_force As Boolean
Dim recirculation_P_refoulement_force As Boolean
Dim recirculation_tps_fonctionnement_force As Boolean
Dim recirculation_Q_unitaire_force As Boolean
Dim recirculation_pompe_rdt_force As Boolean
Dim ratio_ES_volume_media_force As Boolean
Dim ES_nb_pompe_force As Boolean
Dim ES_P_refoulement_force As Boolean
Dim ES_tps_fonctionnement_force As Boolean
Dim ES_Q_unitaire_force As Boolean
Dim ES_pompe_rdt_force As Boolean
Dim lavage_air_Q_force As Boolean
Dim lavage_air_P_refoulement_force As Boolean
Dim lavage_air_tps_fonctionnement_force As Boolean
Dim ES_concentration_force As Boolean
Dim ES_MV_MES_force As Boolean
Dim ES_MES_force As Boolean

'variables échangées avec excel
Dim hauteur_media_f As Double
Dim surface_filtration_min_f As Double
Dim nb_cellules_f As Integer
Dim Cv_DCO_f As Double
Dim Cv_nit_f As Double
Dim sortie_DCO_f As Double  'MSM 15/11/12
Dim sortie_NH4_f As Double
Dim sortie_NO3_f As Double
Dim surpresseur_conso_spec_f As Double
Dim O2_besoin_f As Double
Dim O2_rdt_transfert_f As Double
Dim air_Q_Nm3j_f As Double
Dim diffuseur_encrassement_f As Double
Dim air_P_refoulement_f As Double
Dim recirculation_taux_f As Double
Dim recirculation_nb_pompe_f As Integer
Dim recirculation_P_refoulement_f As Double
Dim recirculation_tps_fonctionnement_f As Double
Dim recirculation_Q_unitaire_f As Double
Dim recirculation_pompe_rdt_f As Double
Dim ratio_ES_volume_media_f As Double
Dim ES_nb_pompe_f As Integer
Dim ES_P_refoulement_f As Double
Dim ES_tps_fonctionnement_f As Double
Dim ES_Q_unitaire_f As Double
Dim ES_pompe_rdt_f As Double
Dim lavage_air_Q_f As Double
Dim lavage_air_P_refoulement_f As Double
Dim lavage_air_tps_fonctionnement_f As Double
Dim ES_concentration_f As Double
Dim ES_MV_MES_f As Double
Dim ES_MES_f As Double

'hypothèses
Dim a_rdt_DCO As Double
Dim b_rdt_DCO As Double
Dim a_rdt_DBO As Double
Dim a_rdt_nit As Double
Dim b_rdt_nit As Double
Dim rdt_nit_max_biostyr_NDNs As Double
Dim rdt_nit As Double
Dim sortie_NH4_min_biostyr_NDNs As Double
Dim rdt_denit As Double
Dim conversion_DCO_design As Double
Dim conversion_DCO_exploit As Double
Dim ratio_O2_nit As Double
Dim ratio_O2_denit As Double
Dim a_hauteur_eau_plafond(2) As Double
Dim correctif_T As Double
Dim T_reference As Double
Dim Cv_DCO_max As Double
Dim Cv_nit_max As Double
Dim Cv_MES_max As Double
Dim Cv_Q_max As Double
Dim nb_cellules_file_max As Integer
Dim nb_cellules_file_min As Integer
Dim nb_cellules_lavage_file As Integer
Dim pollution_ratio_Cv_Cvmax_limite As Double
Dim Cv_Q_correction_eau_diluee As Double
Dim sortie_MES As Double
Dim ratio_P_synthese As Double
Dim surpresseur_conso_Wh_Nm3mCE(nb_type_surpresseur, nb_surpresseur_gamme_Q, nb_surpresseur_gamme_P) As Double
Dim cellule_tps_filtration As Double
Dim hauteur_media_aere As Double
Dim v_eau_limite As Double
Dim air_Q_Nm3j_nominal As Double
Dim ratio_kgO2_Nm3air As Double
Dim md_ImpactRegulationAirSurRendementTransfertOxygene(nb_type_regulation) As Double
Dim facteur_correction_media As Double
Dim ES_tps_retour As Double
Dim ratio_boues_MES_elimine As Double
Dim ratio_boues_DBO_elimine As Double
Dim ratio_boues_N_nitrifie As Double
Dim ratio_DBO_denit As Double
Dim rdt_denit_max_NDN As Double
Dim ratio_Norgaduresoluble_DCOtot As Double

Dim porosite_lit As Double
Dim hauteur_sous_filtre As Double
Dim ratio_eau_non_traite_lit As Double

Dim remarque(4) As String


Dim Q_traite_reel As Double
Dim ES_Q As Double


Sub calcul_consommation_electrique(ByRef electricite_II_aeration)
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'variables pour l'aération
    Dim gamme_P As Integer
    Dim gamme_Q As Integer
    Dim electricite_lavage_air As Double

    'recirculation
    Dim rendement_pompe As Double
    Dim puissance_pompe As Double
    Dim rendement_moteur As Double
    Dim electricite_recirculation As Double
    Dim electricite_ES As Double

    Dim facteur_conso_additionnelle As Double

    'erreurs
    Dim incoherence As Double


    electricite = 0




    'AERATION PROCESS  (A REPRENDRE ET REGARDER LES DONNEES DANS LA SYNTHESE)
    If surpresseur_conso_spec_force = True Then
        surpresseur_conso_spec = surpresseur_conso_spec_f
        'gamme de pression
        If air_P_refoulement < Px1 Then
            gamme_P = surpresseur_gamme_P_0_Px1
        ElseIf air_P_refoulement < Px2 Then
            gamme_P = surpresseur_gamme_P_Px1_Px2
        Else
            gamme_P = surpresseur_gamme_P_Px2_infini
            If choix_surpresseur = surpresseur_roots Then
                remarque(1) = Feuil14.Range("biostyr_erreur_choix_surpresseur").Cells(1, 1)
            End If
        End If
    Else
        'gamme de pression
        If air_P_refoulement < Px1 Then
            gamme_P = surpresseur_gamme_P_0_Px1
        ElseIf air_P_refoulement < Px2 Then
            gamme_P = surpresseur_gamme_P_Px1_Px2
        Else
            gamme_P = surpresseur_gamme_P_Px2_infini
            If choix_surpresseur = surpresseur_roots Then
                remarque(1) = Feuil14.Range("biostyr_erreur_choix_surpresseur").Cells(1, 1)
            End If
        End If
        'gamme de débit   (pour le nominal)
        gamme_Q = surpresseur_gamme_Q_0_infini
'        surpresseur_conso_spec = surpresseur_conso_Wh_Nm3mCE(choix_surpresseur, gamme_Q, gamme_P) * (1 + md_ImpactRegulationAirSurRendementTransfertOxygene(choix_regulation))
        surpresseur_conso_spec = surpresseur_conso_Wh_Nm3mCE(choix_surpresseur, gamme_Q, gamme_P)
    End If
    'calcul de la conso moyenne journalière
    electricite_aeration = air_Q_Nm3j * air_P_refoulement * surpresseur_conso_spec / 1000
    electricite = electricite + electricite_aeration
    electricite_II_aeration = electricite_II_aeration + electricite_aeration

    'AERATION LAVAGE
    electricite_lavage_air = lavage_air_Q * lavage_air_tps_fonctionnement * nb_cellules * lavage_air_P_refoulement * surpresseur_conso_spec / 1000
    electricite = electricite + electricite_lavage_air

    'RECIRCULATION (NDN uniquement)
    If choix_type = biostyr_NDNc Or choix_type = biostyr_NDNs Then
        'recirculation_Q_unitaire = recirculation_taux * Q_traite_reel / (recirculation_nb_pompe * recirculation_tps_fonctionnement)
        If recirculation_Q_unitaire_force = True Then
            recirculation_Q_unitaire = recirculation_Q_unitaire_f
            If recirculation_tps_fonctionnement_force = True Then
                'on ne fait rien
                If recirculation_nb_pompe_force = False Then
                    If Int(recirculation_taux * Q_traite_reel / (recirculation_tps_fonctionnement * recirculation_Q_unitaire)) < (recirculation_taux * Q_traite_reel / (recirculation_tps_fonctionnement * recirculation_Q_unitaire)) Then
                        recirculation_nb_pompe = Int(recirculation_taux * Q_traite_reel / (recirculation_tps_fonctionnement * recirculation_Q_unitaire)) + 1
                    Else
                        recirculation_nb_pompe = Int(recirculation_taux * Q_traite_reel / (recirculation_tps_fonctionnement * recirculation_Q_unitaire))
                    End If
                End If
            Else
                recirculation_tps_fonctionnement = recirculation_taux * Q_traite_reel / (recirculation_nb_pompe * recirculation_Q_unitaire)
                If recirculation_tps_fonctionnement > pd_NOMBRE_HEURE_PAR_JOUR Then
                    recirculation_tps_fonctionnement = pd_NOMBRE_HEURE_PAR_JOUR
                    If recirculation_nb_pompe_force = False Then
                        If Int(recirculation_taux * Q_traite_reel / (recirculation_tps_fonctionnement * recirculation_Q_unitaire)) < (recirculation_taux * Q_traite_reel / (recirculation_tps_fonctionnement * recirculation_Q_unitaire)) Then
                            recirculation_nb_pompe = Int(recirculation_taux * Q_traite_reel / (recirculation_tps_fonctionnement * recirculation_Q_unitaire)) + 1
                        Else
                            recirculation_nb_pompe = Int(recirculation_taux * Q_traite_reel / (recirculation_tps_fonctionnement * recirculation_Q_unitaire))
                        End If
                        recirculation_tps_fonctionnement = recirculation_taux * Q_traite_reel / (recirculation_nb_pompe * recirculation_Q_unitaire)
                    End If
                End If
            End If
        Else
            recirculation_Q_unitaire = recirculation_taux * Q_traite_reel / (recirculation_nb_pompe * recirculation_tps_fonctionnement)
        End If
        incoherence = (recirculation_Q_unitaire * recirculation_nb_pompe * recirculation_tps_fonctionnement - recirculation_taux * Q_traite_reel) / (recirculation_taux * Q_traite_reel)
        If Abs(incoherence) > critere_incoherence Then
            remarque(2) = Feuil14.Range("biostyr_erreur_pompe_recirculation").Cells(1, 1)
        End If
        If recirculation_pompe_rdt_force = True Then
            recirculation_pompe_rdt = recirculation_pompe_rdt_f
        Else
            rendement_pompe = 0.7    'hypothèse
            puissance_pompe = recirculation_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * recirculation_P_refoulement / rendement_pompe
            If puissance_pompe < 60 Then
                rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
            Else
                rendement_moteur = 0.961
            End If
            recirculation_pompe_rdt = rendement_pompe * rendement_moteur
        End If
        electricite_recirculation = recirculation_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * recirculation_P_refoulement / recirculation_pompe_rdt * recirculation_nb_pompe * recirculation_tps_fonctionnement
        electricite = electricite + electricite_recirculation
    End If


    'POMPAGE DES EAUX SALES DU BIOSTYR VERS LE TRAITEMENT SPECIFIQUE
    'ES_Q_unitaire = ES_Q / (ES_nb_pompe * ES_tps_fonctionnement)
    If ES_Q_unitaire_force = True Then
        ES_Q_unitaire = ES_Q_unitaire_f
        If ES_tps_fonctionnement_force = True Then
            'on ne fait rien
            If ES_nb_pompe_force = False Then
                If Int(ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire)) < (ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire)) Then
                    ES_nb_pompe = Int(ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire)) + 1
                Else
                    ES_nb_pompe = Int(ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire))
                End If
            End If
        Else
            ES_tps_fonctionnement = ES_Q / (ES_nb_pompe * ES_Q_unitaire)
            If ES_tps_fonctionnement > ES_tps_retour Then
                ES_tps_fonctionnement = ES_tps_retour
                If ES_nb_pompe_force = False Then
                    If Int(ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire)) < (ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire)) Then
                        ES_nb_pompe = Int(ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire)) + 1
                    Else
                        ES_nb_pompe = Int(ES_Q / (ES_tps_fonctionnement * ES_Q_unitaire))
                    End If
                    ES_tps_fonctionnement = ES_Q / (ES_nb_pompe * ES_Q_unitaire)
                End If
            End If
        End If
    Else
        ES_Q_unitaire = ES_Q / (ES_nb_pompe * ES_tps_fonctionnement)
    End If
    incoherence = (ES_Q_unitaire * ES_nb_pompe * ES_tps_fonctionnement - ES_Q) / (ES_Q)
    If Abs(incoherence) > critere_incoherence Then
        remarque(3) = Feuil14.Range("biostyr_erreur_pompe_ES").Cells(1, 1)
    End If
    If ES_pompe_rdt_force = True Then
        ES_pompe_rdt = ES_pompe_rdt_f
    Else
        rendement_pompe = 0.7    'hypothèse
        puissance_pompe = ES_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * ES_P_refoulement / rendement_pompe
        If puissance_pompe < 60 Then
            rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
        Else
            rendement_moteur = 0.961
        End If
        ES_pompe_rdt = rendement_pompe * rendement_moteur
    End If
    electricite_ES = ES_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * ES_P_refoulement / ES_pompe_rdt * ES_nb_pompe * ES_tps_fonctionnement
    electricite = electricite + electricite_ES

    'CONSOMMATION ADDITIONNELLE
    facteur_conso_additionnelle = 0.02  'vannes ...
    electricite = electricite * (1 + facteur_conso_additionnelle)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "14.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub




Sub dimensionnement(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef Sh, ByRef DCO_conso_bio, ByRef eaux_sales, ByRef choix_eaux_sales, ByVal iteration_ES, ByVal boues_ratio_DBO_MV)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim flag_Q_limitant As Boolean
    Dim I As Integer

    Dim facteur_correction_T As Double
    Dim surface_necessaire As Double
    Dim surface_retenue As Double
    Dim Cv_DCO_max_T As Double
    Dim Cv_nit_max_T As Double
    Dim Cv_MES_max_T As Double
    Dim nb_cellules_lavage As Integer
    Dim nb_file As Integer
    Dim correction_Cv_Q_max As Double
    Dim Cv_MES As Double
    Dim Cv_DCO_eq As Double
    Dim Cv_nit_eq As Double
    Dim Q_considere As Double
    Dim pollution_ratio_Cv_Cvmax As Double
    Dim rdt_DCO As Double
    Dim rdt_DBO As Double
    Dim nitrification As Double
    Dim denitrification As Double
    Dim hauteur_eau_plafond As Double
    Dim v_eau As Double
    Dim DBO_elimine As Double
    Dim MES_elimine As Double
    Dim MESin_mgL As Double
    Dim Norga_dure_soluble As Double


    Norga_dure_soluble = ratio_Norgaduresoluble_DCOtot * (DCO_nominal / (Q_nominal) * 1000)


    MESin_mgL = MES / Q * 1000


    If type_eau_nominal = eau_diluee Then
        If choix_type = biostyr_NDNc Then
            ratio_DBO_denit = 7
        Else
            If choix_regulation = regulation_avance Then
                ratio_DBO_denit = 11
            Else
                ratio_DBO_denit = 11
            End If
        End If
    ElseIf type_eau_nominal = eau_standard Then
        If choix_type = biostyr_NDNc Then
            ratio_DBO_denit = 5
        Else
            If choix_regulation = regulation_avance Then
                ratio_DBO_denit = 8
            Else
                ratio_DBO_denit = 8
            End If
        End If
    Else
        If choix_type = biostyr_NDNc Then
            ratio_DBO_denit = 4
        Else
            If choix_regulation = regulation_avance Then
                ratio_DBO_denit = 5.5
            Else
                ratio_DBO_denit = 5.5
            End If
        End If
    End If


    'correction par l'effet de la température des charges appliquées maximales sur les biofiltres (pollutions)
    facteur_correction_T = correctif_T ^ (pd_TemperatureEauDimensionnement_°C - T_reference)

    'CALCUL DE LA SURFACE DE FILTRATION
    If surface_filtration_min_force = True Then
        surface_filtration_min = surface_filtration_min_f
    Else
        Cv_DCO_max_T = Cv_DCO_max * facteur_correction_T
        Cv_nit_max_T = Cv_nit_max * facteur_correction_T
        Cv_MES_max_T = Cv_MES_max * facteur_correction_T
        'calcul de la surface nécessaire
        correction_Cv_Q_max = 1
        'initialisation nb de cellules
        If iteration = 1 And iteration_ES = 1 Then
            nb_cellules = 10
        End If
        For I = 1 To 2
            'DCO
            surface_necessaire = DCO / (Cv_DCO_max_T * hauteur_media)
            surface_retenue = surface_necessaire
            'MES
            surface_necessaire = MES / (Cv_MES_max_T * hauteur_media)
            If surface_necessaire > surface_retenue Then
                surface_retenue = surface_necessaire
            End If
            'NH4
            If choix_type <> biostyr_C Then
                surface_necessaire = NH4 / (Cv_nit_max_T * hauteur_media)
                If surface_necessaire > surface_retenue Then
                    surface_retenue = surface_necessaire
                End If
            End If
            'hydraulique
            flag_Q_limitant = False
            'nombre de cellules en lavage
            nb_file = nb_cellules / nb_cellules_file_max  'ici on obtient la partie entière
            If nb_file < (nb_cellules / nb_cellules_file_max) Then
                nb_file = nb_file + 1    'si la partie entière est différente du calcul exact, on arrondit au supérieur car on veut le "plafond"
            End If
            nb_cellules_lavage = nb_cellules_lavage_file * nb_file
            'calcul temps de pluie
            Q_considere = Q_nominal * pd_CoefficientPointeHydrauliqueTempsPluie / pd_NOMBRE_HEURE_PAR_JOUR + (Q - Q_nominal) / ES_tps_retour   'Q_ptp+Q_retour  (Q_retour lissé sur 24h)
            surface_necessaire = Q_considere / (Cv_Q_max * correction_Cv_Q_max)
            surface_necessaire = surface_necessaire * nb_cellules / (nb_cellules - nb_cellules_lavage)
            If surface_necessaire > surface_retenue Then
                surface_retenue = surface_necessaire
                flag_Q_limitant = True
            End If
            If choix_type = biostyr_NDNc Or choix_type = biostyr_NDNs Then
                'calcul temps sec (uniquement en NDN)
                Q_considere = (Q_nominal * pd_CoefficientPointeHydrauliqueTempsSec / pd_NOMBRE_HEURE_PAR_JOUR + (Q - Q_nominal) / ES_tps_retour) * (1 + recirculation_taux)    'Q_ptp+Q_retour  (Q_retour lissé sur 24h)
                surface_necessaire = Q_considere / (Cv_Q_max * correction_Cv_Q_max)  'correspond à la surface de filtration réelle (cellules utilisées)
                surface_necessaire = surface_necessaire * nb_cellules / (nb_cellules - nb_cellules_lavage)
                If surface_necessaire > surface_retenue Then
                    surface_retenue = surface_necessaire
                    flag_Q_limitant = True
                End If
            End If
            If flag_Q_limitant = True And I = 1 Then
                Cv_DCO = DCO / (surface_retenue * hauteur_media)
                Cv_MES = MES / (surface_retenue * hauteur_media)
                pollution_ratio_Cv_Cvmax = (Cv_DCO / Cv_DCO_max_T) * (Cv_MES / Cv_MES_max_T)
                If pollution_ratio_Cv_Cvmax < pollution_ratio_Cv_Cvmax_limite Then
                    correction_Cv_Q_max = Cv_Q_correction_eau_diluee
                End If
            End If
        Next I
        surface_filtration_min = surface_retenue
    End If
    If nb_cellules_force = True Then
        nb_cellules = nb_cellules_f
    Else
        nb_cellules = surface_filtration_min / choix_surface_unitaire
        If nb_cellules < surface_filtration_min / choix_surface_unitaire Then
            nb_cellules = nb_cellules + 1
        End If
        If nb_cellules < nb_cellules_file_min Then
            nb_cellules = nb_cellules_file_min
        End If
    End If
    surface_filtration = nb_cellules * choix_surface_unitaire


    'CALCUL VOLUME EAUX SALES       CONCENTRATION EN MES PLUS BAS
    'ré-initialisation des eaux sales
    For I = 1 To nb_parametres_eaux_sales
        eaux_sales(repere_ES_biostyr, I) = 0
    Next I
    choix_eaux_sales(repere_ES_biostyr) = choix_devenir_ES
    If choix_type = biostyr_C Then
        eaux_sales(repere_ES_biostyr, repere_ES_origine) = II_biostyr_C
    ElseIf choix_type = biostyr_N Then
        eaux_sales(repere_ES_biostyr, repere_ES_origine) = II_biostyr_N
    Else
        eaux_sales(repere_ES_biostyr, repere_ES_origine) = II_biostyr_NDN
    End If
    eaux_sales(repere_ES_biostyr, repere_ES_Q) = ratio_ES_volume_media * (surface_filtration * hauteur_media)
    'eaux_sales(repere_ES_biostyr, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr, repere_ES_Q) / 1000
    eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) = ES_MV_MES

    'FONCTIONNEMENT AU NOMINAL    (NDN-->la recirculation est considérée comme étant interne au process et n'intervient donc pas dans les calculs de rendement)
    'pollution carbonée  (tous types de biofiltres)  --> conversion = particulaire dans boues , le reste --> demande en O2, (1-rdt) --> sortie uniquement
    Cv_DCO = DCO / (surface_filtration * hauteur_media)
    'charge en DCO équivalente
    Cv_DCO_eq = Cv_DCO / (facteur_correction_media * facteur_correction_T)
    rdt_DCO = a_rdt_DCO * Exp(b_rdt_DCO * Cv_DCO_eq)   'rendement par rapport aux concentrations entrée/sortiecomme dans simulo
    rdt_DCO = 1 - (Q - eaux_sales(repere_ES_biostyr, repere_ES_Q)) / Q * (1 - rdt_DCO)      'correction pour obtenir le rendement en fonction des charges entrée/sortie (car débits volumiques différents)
    DCO_conso_bio = DCO_conso_bio + (1 - conversion_DCO_design) * (rdt_DCO * DCO)
    rdt_DBO = rdt_DCO + a_rdt_DBO
    DBO_elimine = DBO * rdt_DBO
    'DBO = sortie_MES * (Q - eaux_sales(repere_ES_biostyr, repere_ES_Q)) / 1000 * ES_MV_MES * boues_ratio_DBO_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine))
    'DBO_elimine = DBO_elimine - DBO
    'calcul des charges dans les eaux sales
    'DCO: on considère que ce n'est que le particulaire représenté par le facteur de conversion)
    eaux_sales(repere_ES_biostyr, repere_ES_DCO) = conversion_DCO_design * (rdt_DCO * DCO)
    'pollution azotée (N-NH4 et N-NO3)
    If choix_type <> biostyr_C Then
        Cv_nit = NH4 / (surface_filtration * hauteur_media)
        'charge équivalente en NH4
        Cv_nit_eq = Cv_nit / (facteur_correction_media * facteur_correction_T)
        rdt_nit = a_rdt_nit * Exp(b_rdt_nit * Cv_nit_eq)
        'pour le biostyr NDNs, on ne dépasse pas 90% de rendement
        If choix_type = biostyr_NDNs Then
            If rdt_nit > rdt_nit_max_biostyr_NDNs Then
                rdt_nit = rdt_nit_max_biostyr_NDNs
            End If
        End If
    End If
    'NH4 éliminé ( on considère le débit entrant --> concentration sortie = concentration eaux sales )
    nitrification = NH4 * rdt_nit
    sortie_NH4 = (NH4 - nitrification) / Q * 1000
    If choix_type = biostyr_NDNs Then
        If sortie_NH4 < sortie_NH4_min_biostyr_NDNs Then
            sortie_NH4 = sortie_NH4_min_biostyr_NDNs
            nitrification = NH4 - sortie_NH4 * Q / 1000
        End If
    End If
    'calcul des NH4 dans les eaux sales (hyp: même concentration en NH4 dans les eaux sales que dans l'eau traitée)
    eaux_sales(repere_ES_biostyr, repere_ES_NH4) = eaux_sales(repere_ES_biostyr, repere_ES_Q) * sortie_NH4 / 1000
    'on calcule les NK dans les eaux sales (on ajoutera par la suite les NK particulaire liés aux MES en 1ère approche)
    eaux_sales(repere_ES_biostyr, repere_ES_NK) = eaux_sales(repere_ES_biostyr, repere_ES_NH4)
    'NO3 éliminé (NDN)
    If choix_type = biostyr_NDNc Or choix_type = biostyr_NDNs Then
        denitrification = DBO * rdt_DBO / ratio_DBO_denit
        If denitrification > rdt_denit_max_NDN * (NO3 + nitrification) Then
            denitrification = rdt_denit_max_NDN * (NO3 + nitrification)
        End If
    Else
        denitrification = (NO3 + nitrification) * rdt_denit
    End If
    sortie_NO3 = (NO3 + nitrification - denitrification) / Q * 1000
    'calcul de la dénitrification réelle
    eaux_sales(repere_ES_biostyr, repere_ES_NO3) = eaux_sales(repere_ES_biostyr, repere_ES_Q) * sortie_NO3 / 1000


    'calcul de la concentration des boues
    MES_elimine = MES - sortie_MES * (Q - eaux_sales(repere_ES_biostyr, repere_ES_Q)) / 1000
    'boues produites   2ème ligne DCA à partir d'infos de Xavier LE TALLEC le 22/07/11
    If type_eau_nominal = eau_diluee Then
        ratio_boues_DBO_elimine = 0.2
    ElseIf type_eau_nominal = eau_standard Then
        ratio_boues_DBO_elimine = 0.4
    Else
        ratio_boues_DBO_elimine = 0.4
    End If
    eaux_sales(repere_ES_biostyr, repere_ES_MES) = ratio_boues_MES_elimine * MES + ratio_boues_DBO_elimine * DBO_elimine + ratio_boues_N_nitrifie * nitrification
    eaux_sales(repere_ES_biostyr, repere_ES_MES) = eaux_sales(repere_ES_biostyr, repere_ES_MES) + MESin_mgL * (hauteur_sous_filtre * surface_filtration + ratio_eau_non_traite_lit * porosite_lit * hauteur_media * surface_filtration) / 1000
    'eaux_sales(repere_ES_biostyr, repere_ES_MES) = eaux_sales(repere_ES_biostyr, repere_ES_MES) * 1.25
    ES_concentration = eaux_sales(repere_ES_biostyr, repere_ES_MES) / eaux_sales(repere_ES_biostyr, repere_ES_Q) * 1000
    '************ A VOIR SI AU DESSUS CA ESTIME BIEN POUR ENLEVER CE QUI SUIT *************
    GoTo sauter_boucle
    If ES_concentration_force = True Then
        ES_concentration = ES_concentration_f
        eaux_sales(repere_ES_biostyr, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr, repere_ES_Q) / 1000
    Else
        If type_eau_nominal = eau_diluee Then
            If choix_type = biostyr_C Then
                ES_concentration = 200
            Else
                ES_concentration = 300
            End If
        ElseIf type_eau_nominal = eau_standard Then
            If choix_type = biostyr_C Then
                ES_concentration = 250
            Else
                ES_concentration = 400
            End If
        Else
            If choix_type = biostyr_C Then
                ES_concentration = 300
            Else
                ES_concentration = 500
            End If
        End If
        eaux_sales(repere_ES_biostyr, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr, repere_ES_Q) / 1000
    End If
sauter_boucle:


    'AERATION
    'calcul des besoins en O2
    O2_besoin = (1 - conversion_DCO_design) * (rdt_DCO * DCO) + ratio_O2_nit * nitrification + ratio_O2_denit * denitrification
    O2_besoin = O2_besoin + besoins_O2_HS(Sh)
    'calcul du rendement de transfert de l'O2
    If O2_rdt_transfert_force = True Then
        O2_rdt_transfert = O2_rdt_transfert_f
    Else
        'vitesse en eau
        v_eau = Q / (choix_surface_unitaire * (nb_cellules - nb_cellules_lavage) * cellule_tps_filtration)
        'calcul de la hauteur d'eau sur le plafond
        hauteur_eau_plafond = 0
        For I = 0 To 1
            hauteur_eau_plafond = hauteur_eau_plafond + a_hauteur_eau_plafond(I) * (hauteur_media / nb_cellules) ^ I
        Next I
        If choix_type = biostyr_NDNc Then
            O2_rdt_transfert = (6.2 + 0.375 * v_eau) * (hauteur_media_aere - 0.4) * (10.7 + hauteur_eau_plafond + hauteur_media_aere / 2) / 12.7
        ElseIf v_eau < v_eau_limite Then
            O2_rdt_transfert = (6.2 + 0.375 * v_eau) * hauteur_media_aere * (10.7 + hauteur_eau_plafond + hauteur_media_aere / 2) / 13.4
        Else
            O2_rdt_transfert = 8.45 * hauteur_media_aere * (10.7 + hauteur_eau_plafond + hauteur_media_aere / 2) / 13.4
        End If
        O2_rdt_transfert = O2_rdt_transfert / 100    'pour l'avoir en %
        'impact du type de régulation
        O2_rdt_transfert = O2_rdt_transfert * (1 - md_ImpactRegulationAirSurRendementTransfertOxygene(choix_regulation))
    End If
    air_Q_Nm3j_nominal = O2_besoin / (O2_rdt_transfert * ratio_kgO2_Nm3air)
    If air_Q_Nm3j_force = True And air_Q_Nm3j_f > air_Q_Nm3j_nominal Then
        air_Q_Nm3j_nominal = air_Q_Nm3j_f
    End If




    'CHARGES EN SORTIE
    'attention, certaines valeurs seront remodifiées plus loin car dépendent encore des hypothèses de composition des boues comme les NK et le Pt
    Q = Q - eaux_sales(repere_ES_biostyr, repere_ES_Q)
    DCO = DCO * (1 - rdt_DCO)
    DBO = DBO * (1 - rdt_DBO)
    MES = sortie_MES * Q / 1000
    NH4 = sortie_NH4 * Q / 1000
    NK = NH4 + Norga_dure_soluble * Q / 1000
    NO3 = sortie_NO3 * Q / 1000
    Pt = Pt - eaux_sales(repere_ES_biostyr, repere_ES_Pt)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "14.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub fonctionnement_reel(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef Sh, ByRef DCO_conso_bio, ByRef eaux_sales, ByRef choix_eaux_sales, ByVal iteration_ES, ByVal boues_ratio_DBO_MV)
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer

    Dim facteur_correction_T As Double
    Dim nb_cellules_lavage As Integer
    Dim nb_file As Integer
    Dim Cv_MES As Double
    Dim Cv_DCO_eq As Double
    Dim Cv_nit_eq As Double
    Dim Q_considere As Double
    Dim rdt_DCO As Double
    Dim rdt_DBO As Double
    Dim nitrification As Double
    Dim denitrification As Double
    Dim hauteur_eau_plafond As Double
    Dim v_eau As Double
    Dim DBO_elimine As Double
    Dim MES_elimine As Double
    Dim MESin_mgL As Double
    Dim Norga_dure_soluble As Double


    Norga_dure_soluble = ratio_Norgaduresoluble_DCOtot * (DCO_nominal * pd_PourcentageChargeReelleDCO / (Q_nominal * pd_PourcentageChargeReelleDebitVolumique) * 1000)


    MESin_mgL = MES / Q * 1000


    If type_eau_reel = eau_diluee Then
        If choix_type = biostyr_NDNc Then
            ratio_DBO_denit = 7
        Else
            If choix_regulation = regulation_avance Then
                ratio_DBO_denit = 11
            Else
                ratio_DBO_denit = 11
            End If
        End If
    ElseIf type_eau_reel = eau_standard Then
        If choix_type = biostyr_NDNc Then
            ratio_DBO_denit = 5
        Else
            If choix_regulation = regulation_avance Then
                ratio_DBO_denit = 8
            Else
                ratio_DBO_denit = 8
            End If
        End If
    Else
        If choix_type = biostyr_NDNc Then
            ratio_DBO_denit = 4
        Else
            If choix_regulation = regulation_avance Then
                ratio_DBO_denit = 5.5
            Else
                ratio_DBO_denit = 5.5
            End If
        End If
    End If




    Q_traite_reel = Q

    facteur_correction_T = correctif_T ^ (pd_TemperatureEauConditionsReelles_°C - T_reference)


    'CALCUL VOLUME EAUX SALES ET DE LA CONCENTRATION EN MES
    'ré-initialisation des eaux sales
    For I = 1 To nb_parametres_eaux_sales
        eaux_sales(repere_ES_biostyr, I) = 0
    Next I
    choix_eaux_sales(repere_ES_biostyr) = choix_devenir_ES
    If choix_type = biostyr_C Then
        eaux_sales(repere_ES_biostyr, repere_ES_origine) = II_biostyr_C
    ElseIf choix_type = biostyr_N Then
        eaux_sales(repere_ES_biostyr, repere_ES_origine) = II_biostyr_N
    Else
        eaux_sales(repere_ES_biostyr, repere_ES_origine) = II_biostyr_NDN
    End If
    eaux_sales(repere_ES_biostyr, repere_ES_Q) = ratio_ES_volume_media * (surface_filtration * hauteur_media)
    'eaux_sales(repere_ES_biostyr, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr, repere_ES_Q) / 1000
    'ES_MES = eaux_sales(repere_ES_biostyr, repere_ES_MES)
    eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) = ES_MV_MES

    'FONCTIONNEMENT REEL   (NDN-->la recirculation est considérée comme étant interne au process et n'intervient donc pas dans les calculs de rendement)
    'pollution carbonée  (tous types de biofiltres)  --> conversion = particulaire dans boues , le reste --> demande en O2, (1-rdt) --> sortie uniquement
    If Cv_DCO_force = True Then
        Cv_DCO = Cv_DCO_f
    Else
        Cv_DCO = DCO / (surface_filtration * hauteur_media)
    End If
    'charge en DCO équivalente
    Cv_DCO_eq = Cv_DCO / (facteur_correction_media * facteur_correction_T)

    'début modif MSM 15/11/12
    '    rdt_DCO = a_rdt_DCO * Exp(b_rdt_DCO * Cv_DCO_eq)   'rendement par rapport aux concentrations entrée/sortie comme dans simulo
    '    rdt_DCO = 1 - (Q - eaux_sales(repere_ES_biostyr, repere_ES_Q)) / Q * (1 - rdt_DCO)      'correction pour obtenir le rendement en fonction des charges entrée/sortie (car débits volumiques différents)
    '    DCO_conso_bio = DCO_conso_bio + (1 - conversion_DCO_design) * (rdt_DCO * DCO)
    '--> recalculer rendement si concentration en sortie est forcée
    If sortie_DCO_force = True Then
        sortie_DCO = sortie_DCO_f
        rdt_DCO = 1 - (sortie_DCO * (Q - eaux_sales(repere_ES_biostyr, repere_ES_Q)) / 1000 / DCO)    'calculé à partir de l'équation ci dessous (*)
    Else
        rdt_DCO = a_rdt_DCO * Exp(b_rdt_DCO * Cv_DCO_eq)   'rendement par rapport aux concentrations entrée/sortiecomme dans simulo
        rdt_DCO = 1 - (Q - eaux_sales(repere_ES_biostyr, repere_ES_Q)) / Q * (1 - rdt_DCO)      'correction pour obtenir le rendement en fonction des charges entrée/sortie (car débits volumiques différents)
        sortie_DCO = DCO * (1 - rdt_DCO) / (Q - eaux_sales(repere_ES_biostyr, repere_ES_Q)) * 1000    '(*)
    End If
    DCO_conso_bio = DCO_conso_bio + (1 - conversion_DCO_design) * (rdt_DCO * DCO)
    'fin modif MSM 15/11/12

    rdt_DBO = rdt_DCO + a_rdt_DBO
    'calcul des charges dans les eaux sales
    'DCO: on considère que ce n'est que le particulaire représenté par le facteur de conversion)
    eaux_sales(repere_ES_biostyr, repere_ES_DCO) = conversion_DCO_design * (rdt_DCO * DCO)
    'DBO: on considère ce que l'on a éliminé et on enlève la partie de la DCO dégradée par les bactéries
    DBO_elimine = DBO * rdt_DBO
    'DBO = sortie_MES * (Q - eaux_sales(repere_ES_biostyr, repere_ES_Q)) / 1000 * ES_MV_MES * boues_ratio_DBO_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine))
    'DBO_elimine = DBO_elimine - DBO
    'eaux_sales(repere_ES_biostyr, repere_ES_DBO) = rdt_DBO * DBO - DCO_conso_bio
    'If eaux_sales(repere_ES_biostyr, repere_ES_DBO) < 0 Then
    '    eaux_sales(repere_ES_biostyr, repere_ES_DBO) = 0
    '    rdt_DBO = 1
    'End If
    'élimination de Pt
    'Pt dans les eaux sales: on suppose l'assimilation lors de l'élimination de la DBO + Pt particulaire (avec ratio Pt/MES) (ce sera faux mais mieux que rien)
    'eaux_sales(repere_ES_biostyr, repere_ES_Pt) = ratio_P_synthese * DCO_conso_bio
    'pollution azotée (N-NH4 et N-NO3)
    If choix_type <> biostyr_C Then
        If Cv_nit_force = True Then
            Cv_nit = Cv_nit_f
        Else
            Cv_nit = NH4 / (surface_filtration * hauteur_media)
        End If
        Cv_nit_eq = Cv_nit / (facteur_correction_media * facteur_correction_T)
        rdt_nit = a_rdt_nit * Exp(b_rdt_nit * Cv_nit_eq)
        'pour le biostyr NDNs, on ne dépasse pas 90% de rendement
        If choix_type = biostyr_NDNs Then
            If rdt_nit > rdt_nit_max_biostyr_NDNs Then
                rdt_nit = rdt_nit_max_biostyr_NDNs
            End If
        End If
    End If
    'NH4 éliminé ( on considère le débit entrant --> concentration sortie = concentration eaux sales )
    If sortie_NH4_force = True Then
        sortie_NH4 = sortie_NH4_f
        nitrification = NH4 - sortie_NH4 * Q / 1000
    Else
        nitrification = NH4 * rdt_nit
        sortie_NH4 = (NH4 - nitrification) / Q * 1000
        If choix_type = biostyr_NDNs Then
            If sortie_NH4 < sortie_NH4_min_biostyr_NDNs Then
                sortie_NH4 = sortie_NH4_min_biostyr_NDNs
                nitrification = NH4 - sortie_NH4 * Q / 1000
            End If
        End If
    End If
    'calcul des NH4 dans les eaux sales (hyp: même concentration en NH4 dans les eaux sales que dans l'eau traitée)
    eaux_sales(repere_ES_biostyr, repere_ES_NH4) = eaux_sales(repere_ES_biostyr, repere_ES_Q) * sortie_NH4 / 1000
    'on calcule les NK dans les eaux sales (on ajoutera par la suite les NK particulaire liés aux MES en 1ère approche)
    eaux_sales(repere_ES_biostyr, repere_ES_NK) = eaux_sales(repere_ES_biostyr, repere_ES_NH4)
    'NO3 éliminé (NDN)
    If sortie_NO3_force = True Then
        sortie_NO3 = sortie_NO3_f
        denitrification = (NO3 + nitrification) - sortie_NO3 * Q / 1000
    Else
        If choix_type = biostyr_NDNc Or choix_type = biostyr_NDNs Then
            denitrification = DBO * rdt_DBO / ratio_DBO_denit
            If denitrification > rdt_denit_max_NDN * (NO3 + nitrification) Then
                denitrification = rdt_denit_max_NDN * (NO3 + nitrification)
            End If
        Else
            denitrification = (NO3 + nitrification) * rdt_denit
        End If
        sortie_NO3 = (NO3 + nitrification - denitrification) / Q * 1000
    End If
    'calcul de la dénitrification réelle
    eaux_sales(repere_ES_biostyr, repere_ES_NO3) = eaux_sales(repere_ES_biostyr, repere_ES_Q) * sortie_NO3 / 1000


    'calcul de la concentration des boues
    If type_eau_reel = eau_diluee Then
        ratio_boues_DBO_elimine = 0.2
    ElseIf type_eau_reel = eau_standard Then
        ratio_boues_DBO_elimine = 0.4
    Else
        ratio_boues_DBO_elimine = 0.4
    End If
    If ES_concentration_force = True Then
        ES_concentration = ES_concentration_f
        eaux_sales(repere_ES_biostyr, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr, repere_ES_Q) / 1000
    Else
        MES_elimine = MES - sortie_MES * (Q - eaux_sales(repere_ES_biostyr, repere_ES_Q)) / 1000
        'boues produites   2ème ligne DCA à partir d'infos de Xavier LE TALLEC le 22/07/11
        eaux_sales(repere_ES_biostyr, repere_ES_MES) = ratio_boues_MES_elimine * MES + ratio_boues_DBO_elimine * DBO_elimine + ratio_boues_N_nitrifie * nitrification
        eaux_sales(repere_ES_biostyr, repere_ES_MES) = eaux_sales(repere_ES_biostyr, repere_ES_MES) + MESin_mgL * (hauteur_sous_filtre * surface_filtration + ratio_eau_non_traite_lit * porosite_lit * hauteur_media * surface_filtration) / 1000
        'eaux_sales(repere_ES_biostyr, repere_ES_MES) = eaux_sales(repere_ES_biostyr, repere_ES_MES) * 1.25
        ES_concentration = eaux_sales(repere_ES_biostyr, repere_ES_MES) / eaux_sales(repere_ES_biostyr, repere_ES_Q) * 1000
    End If
    '************ A VOIR SI AU DESSUS CA ESTIME BIEN POUR ENLEVER CE QUI SUIT *************
    GoTo sauter_boucle
    If ES_concentration_force = True Then
        ES_concentration = ES_concentration_f
        eaux_sales(repere_ES_biostyr, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr, repere_ES_Q) / 1000
    Else
        If type_eau_reel = eau_diluee Then
            If choix_type = biostyr_C Then
                ES_concentration = 200
            Else
                ES_concentration = 300
            End If
        ElseIf type_eau_reel = eau_standard Then
            If choix_type = biostyr_C Then
                ES_concentration = 250
            Else
                ES_concentration = 400
            End If
        Else
            If choix_type = biostyr_C Then
                ES_concentration = 300
            Else
                ES_concentration = 500
            End If
        End If
        eaux_sales(repere_ES_biostyr, repere_ES_MES) = ES_concentration * eaux_sales(repere_ES_biostyr, repere_ES_Q) / 1000
    End If
sauter_boucle:

    ES_MES = eaux_sales(repere_ES_biostyr, repere_ES_MES)


    'AERATION
    'calcul des besoins en O2
    If O2_besoin_force = True Then
        O2_besoin = O2_besoin_f
    Else
        O2_besoin = (1 - conversion_DCO_design) * (rdt_DCO * DCO) + ratio_O2_nit * nitrification + ratio_O2_denit * denitrification
        O2_besoin = O2_besoin + besoins_O2_HS(Sh)
    End If
    'calcul du rendement de transfert de l'O2
    If O2_rdt_transfert_force = True Then
        O2_rdt_transfert = O2_rdt_transfert_f
    Else
        'vitesse en eau
        v_eau = Q / (choix_surface_unitaire * (nb_cellules - nb_cellules_lavage) * cellule_tps_filtration)
        'calcul de la hauteur d'eau sur le plafond
        hauteur_eau_plafond = 0
        For I = 0 To 1
            hauteur_eau_plafond = hauteur_eau_plafond + a_hauteur_eau_plafond(I) * (hauteur_media / nb_cellules) ^ I
        Next I
        If choix_type = biostyr_NDNc Then
            O2_rdt_transfert = (6.2 + 0.375 * v_eau) * (hauteur_media_aere - 0.4) * (10.7 + hauteur_eau_plafond + hauteur_media_aere / 2) / 12.7
        ElseIf v_eau < v_eau_limite Then
            O2_rdt_transfert = (6.2 + 0.375 * v_eau) * hauteur_media_aere * (10.7 + hauteur_eau_plafond + hauteur_media_aere / 2) / 13.4
        Else
            O2_rdt_transfert = 8.45 * hauteur_media_aere * (10.7 + hauteur_eau_plafond + hauteur_media_aere / 2) / 13.4
        End If
        O2_rdt_transfert = O2_rdt_transfert / 100    'pour l'avoir en %
        'impact du type de régulation
        O2_rdt_transfert = O2_rdt_transfert * (1 - md_ImpactRegulationAirSurRendementTransfertOxygene(choix_regulation))
    End If
    'calcul du débit nécessaire
    If air_Q_Nm3j_force = True Then
        air_Q_Nm3j = air_Q_Nm3j_f
    Else
        air_Q_Nm3j = O2_besoin / (O2_rdt_transfert * ratio_kgO2_Nm3air)
    End If

    'CHARGES EN SORTIE
    'attention, certaines valeurs seront remodifiées plus loin car dépendent encore des hypothèses de composition des boues comme les NK et le Pt
    Q = Q - eaux_sales(repere_ES_biostyr, repere_ES_Q)
    DCO = sortie_DCO * Q / 1000    'MSM 15/11/12
    DBO = DBO * (1 - rdt_DBO)
    MES = sortie_MES * Q / 1000
    NH4 = sortie_NH4 * Q / 1000
    NK = NH4 + Norga_dure_soluble * Q / 1000
    NO3 = sortie_NO3 * Q / 1000
    Pt = Pt    '- eaux_sales(repere_ES_biostyr, repere_ES_Pt)

    'stockage eaux sales
    ES_Q = eaux_sales(repere_ES_biostyr, repere_ES_Q)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "14.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub attribution_valeur_par_defaut()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'hauteur de média
    If hauteur_media_force = True Then
        hauteur_media = hauteur_media_f
        If hauteur_media < hauteur_media_aere Then
            hauteur_media_aere = hauteur_media
        End If
    Else
        hauteur_media = 3.5
    End If

    'valeurs relatives à l'aération
    If air_P_refoulement_force = True Then
        air_P_refoulement = air_P_refoulement_f
    Else
        air_P_refoulement = 10   'mCE
    End If

    'recirculation nitrates
    If recirculation_taux_force = True Then
        recirculation_taux = recirculation_taux_f
    Else
        If choix_type = biostyr_NDNc Then
            recirculation_taux = 1.5    '150% pour le classique
        ElseIf choix_type = biostyr_NDNs Then
            recirculation_taux = 0.8    '80% pour le simultané
        Else
            recirculation_taux = 0    'pas de recirculation pour les autres car pas dobjectif de dénitrifier
        End If
    End If
    If recirculation_nb_pompe_force = True Then
        recirculation_nb_pompe = recirculation_nb_pompe_f
    Else
        recirculation_nb_pompe = 2
    End If
    If recirculation_P_refoulement_force = True Then
        recirculation_P_refoulement = recirculation_P_refoulement_f
    Else
        recirculation_P_refoulement = 4   'mCE
    End If
    If recirculation_tps_fonctionnement_force = True Then
        recirculation_tps_fonctionnement = recirculation_tps_fonctionnement_f
    Else
        recirculation_tps_fonctionnement = 24   'h/j
    End If

    'eaux sales
    If ratio_ES_volume_media_force = True Then
        ratio_ES_volume_media = ratio_ES_volume_media_f
    Else
        If choix_type = biostyr_C Then
            ratio_ES_volume_media = 3.5  ' 3.5*volume_media
        Else
            ratio_ES_volume_media = 2.5   '2.5*volume_media
        End If
    End If
    If ES_nb_pompe_force = True Then
        ES_nb_pompe = ES_nb_pompe_f
    Else
        ES_nb_pompe = 2
    End If
    If ES_P_refoulement_force = True Then
        ES_P_refoulement = ES_P_refoulement_f
    Else
        ES_P_refoulement = 10   'mCE
    End If
    If ES_tps_fonctionnement_force = True Then
        ES_tps_fonctionnement = ES_tps_fonctionnement_f
    Else
        ES_tps_fonctionnement = 18   'h/j
    End If

    'lavage des filtres (injection d'air)
    If lavage_air_P_refoulement_force = True Then
        lavage_air_P_refoulement = lavage_air_P_refoulement_f
    Else
        lavage_air_P_refoulement = 8    'mCE
    End If
    If lavage_air_tps_fonctionnement_force = True Then
        lavage_air_tps_fonctionnement = lavage_air_tps_fonctionnement_f
    Else
        lavage_air_tps_fonctionnement = 0.1    'h/j  6 min par lavage donc 0.1h/j sur l'hypothèse d'un lavage
    End If
    If lavage_air_Q_force = True Then
        lavage_air_Q = lavage_air_Q_f
    Else
        If choix_type = biostyr_NDNc Then
            lavage_air_Q = 15 * choix_surface_unitaire
        Else
            lavage_air_Q = 12 * choix_surface_unitaire
        End If
    End If

    'qualité des eaux sales
    'If ES_concentration_force = True Then   'ATTENTION, concentration des eaux sales en mg/L
    '    ES_concentration = ES_concentration_f
    'Else
    '    If choix_type = biostyr_C Then
    '        ES_concentration = 500   'mg/L
    '    Else
    '        ES_concentration = 400   'mg/L
    '    End If
    'End If

    If ES_MV_MES_force = True And choix_devenir_ES <> devenir_ES_primaire Then
        ES_MV_MES = ES_MV_MES_f
    Else
        ES_MV_MES = 0.85  ' 85% MV/MES
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "14.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub hypotheses()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer



    'élimination de la DCO     a*exp(b*Cv_DCO)
    a_rdt_DCO = 91.027 / 100
    b_rdt_DCO = -0.0431

    'élimination de la DBO     rdt_DCO+a
    a_rdt_DBO = 0.1     '0.05

    'nitrification (sauf biostyr_C)    a*exp(b*Cv_nit)
    a_rdt_nit = 99.648 / 100
    b_rdt_nit = -0.1762
    rdt_nit_max_biostyr_NDNs = 0.9   'rendement maximum atteignable en nit-dénit simultanée
    'nitrification (biostyr C)
    rdt_nit = 0.2    'rendement de 20% par défaut pour le biostyr C
    'concentration en NH4 limite pour le biostyr NDNs au dessous de laquelle on est incapable de descendre
    If choix_regulation = regulation_avance Then
        sortie_NH4_min_biostyr_NDNs = 5   'mg/L
    Else
        sortie_NH4_min_biostyr_NDNs = 4   'mg/L
    End If

    'dénitrification
    If choix_type = biostyr_C Then
        rdt_denit = 0   'très peu de nit et pas de dénit
    ElseIf choix_type = biostyr_N Then
        rdt_denit = 0   'pas de dénit
    ElseIf choix_type = biostyr_NDNc Then
        'rdt_denit = 0.6   'dénit dans la zone inférieure du filtre
        ratio_DBO_denit = 6
    ElseIf choix_type = biostyr_NDNs Then
        If choix_regulation = regulation_avance Then
            'rdt_denit = 0.4   'dénit dans les couches inférieures de la biomasse fixée (car pas d'O2)
            ratio_DBO_denit = 7
        Else
            'rdt_denit = 0.3   'dénit dans les couches inférieures de la biomasse fixée (car pas d'O2)
            ratio_DBO_denit = 7.5
        End If
    End If
    rdt_denit_max_NDN = 0.75

    'besoins en O2
    'coefficient de conversion de la DCO  --> représente la part de DCO éliminée sous forme particulaire
    'conditions de design
    If pd_TemperatureEauDimensionnement_°C <= 16 Then
        conversion_DCO_design = 0.46
    ElseIf pd_TemperatureEauDimensionnement_°C <= 20 Then
        conversion_DCO_design = 0.32
    Else
        conversion_DCO_design = 0.26
    End If
    'conditions d'exploitation
    If pd_TemperatureEauConditionsReelles_°C <= 16 Then
        conversion_DCO_exploit = 0.46
    ElseIf pd_TemperatureEauConditionsReelles_°C <= 20 Then
        conversion_DCO_exploit = 0.32
    Else
        conversion_DCO_exploit = 0.26
    End If
    'nitrification
    ratio_O2_nit = 4.57
    'dénitrification
    ratio_O2_denit = -2.86

    'calcul du rendement de transfert d'O2  (les coefficients des relations
    'temps de fonctionnement moyen d'une cellule  (h/j)
    cellule_tps_filtration = 23.5  'toute la journée moins le temps d'un lavage ~ 30 min
    'calcul de la hauteur d'eau sur plafond   a(0)+a(1)*(volume_media/surface_totale*nb_cellules)
    a_hauteur_eau_plafond(0) = 0.5
    If choix_type = biostyr_C Then
        a_hauteur_eau_plafond(1) = 3.5
    Else
        a_hauteur_eau_plafond(1) = 2.5
    End If
    If choix_type = biostyr_NDNc Then
        hauteur_media_aere = 2.5
    Else
        hauteur_media_aere = 3.5
    End If
    v_eau_limite = 6  'm/h
    'O2 présent dans l'air (kgO2/Nm3air)
    ratio_kgO2_Nm3air = 0.3
    'impact du type de régulation
    md_ImpactRegulationAirSurRendementTransfertOxygene(regulation_horloge) = 0.4
    md_ImpactRegulationAirSurRendementTransfertOxygene(regulation_sans_variation) = 0.25
    md_ImpactRegulationAirSurRendementTransfertOxygene(regulation_avec_variation) = 0.1
    md_ImpactRegulationAirSurRendementTransfertOxygene(regulation_avance) = 0


    'calcul des charges appliquées limites (en kg/(m3.j) pour DCO, nit et MES; en m/h pour Q)
    If choix_type = biostyr_C Then
        Cv_Q_correction_eau_diluee = 1.25  'lorsque eau diluée --> application de la correction  modif CAM 13/12/12
        If choix_diametre_media = 4.5 Then
            Cv_DCO_max = 8
            Cv_MES_max = 2.8
            Cv_Q_max = 6
        ElseIf choix_diametre_media = 5 Then
            Cv_DCO_max = 7
            Cv_MES_max = 3.5
            Cv_Q_max = 8
        End If
    ElseIf choix_type = biostyr_N Then
        Cv_Q_correction_eau_diluee = 1.75  'lorsque eau diluée --> application de la correction  modif CAM 13/12/12 (1,25-->1,75)
        If choix_diametre_media = 3.6 Then
            Cv_DCO_max = 6
            Cv_nit_max = 0.7
            Cv_MES_max = 1.8
            Cv_Q_max = 4
        ElseIf choix_diametre_media = 4 Then
            Cv_DCO_max = 5.4
            Cv_nit_max = 0.6
            Cv_MES_max = 2.3
            Cv_Q_max = 6
        ElseIf choix_diametre_media = 4.5 Then
            Cv_DCO_max = 4.8
            Cv_nit_max = 0.5
            Cv_MES_max = 2.8
            Cv_Q_max = 8
        End If
    Else   'pour les deux nit-dénit
        Cv_Q_correction_eau_diluee = 1.25  'lorsque eau diluée --> application de la correction  modif CAM 13/12/12
        If choix_diametre_media = 3.6 Then
            Cv_DCO_max = 6
            Cv_nit_max = 0.6
            Cv_MES_max = 1.8
            Cv_Q_max = 6
        ElseIf choix_diametre_media = 4 Then
            Cv_DCO_max = 5.4
            Cv_nit_max = 0.55
            Cv_MES_max = 2.3
            Cv_Q_max = 8
        ElseIf choix_diametre_media = 4.5 Then
            Cv_DCO_max = 4.8
            Cv_nit_max = 0.5
            Cv_MES_max = 2.8
            Cv_Q_max = 10
        End If
    End If
    'correction par rapport au type de média   (pour la partie simulation des performances)
    facteur_correction_media = 3.6 / choix_diametre_media

    'calcul du nombre de cellules en lavage simultanément
    nb_cellules_file_max = 20    '20 cellules au maximum par file
    nb_cellules_file_min = 4     '4 cellules au minimum par file
    nb_cellules_lavage_file = 1  '1 seule cellule par file en lavage
    'coefficients pour calcul correction en température des charges volumiques limites   correctif_T^(T-T_reference)
    'Facteur correctif température pour cinétique
    correctif_T = 1.07
    'Température de référence pour la cinétique biologique (°C)
    T_reference = 12
    'coefficients de correction si l'eau à traiter est très diluée
    pollution_ratio_Cv_Cvmax_limite = 0.75   'ratio limite (Cv_MES*Cv_DCO)/(Cvmax_DCO*Cvmax_MES) au dessous duquel on peut augmenter la charge hydraulique


    'performances sur la filtration des MES (concentration fixée par granulométrie de média)
    If choix_diametre_media = 3.6 Then
        sortie_MES = 13
    ElseIf choix_diametre_media = 4 Then
        sortie_MES = 18
    ElseIf choix_diametre_media = 4.5 Then
        sortie_MES = 23
    ElseIf choix_diametre_media = 5 Then
        sortie_MES = 28
    End If

    'calcul de la production de boues
    If choix_type = biostyr_C Then
        ratio_boues_MES_elimine = 1
        ratio_boues_DBO_elimine = 0.4
        ratio_boues_N_nitrifie = 0
    Else
        ratio_boues_MES_elimine = 0.86
        ratio_boues_DBO_elimine = 0.4
        ratio_boues_N_nitrifie = 0.2
    End If


    'élimination de P
    ratio_P_synthese = 0.01

    'consommations spécifiques par type de surpresseurs      en Wh/(Nm3.mCE)
    For I = 1 To nb_surpresseur_gamme_Q
        For j = 1 To nb_surpresseur_gamme_P
            surpresseur_conso_Wh_Nm3mCE(surpresseur_roots, I, j) = 4.5
            surpresseur_conso_Wh_Nm3mCE(surpresseur_vis, I, j) = 3.8
            surpresseur_conso_Wh_Nm3mCE(surpresseur_turbo, I, j) = 3.5
        Next j
    Next I

    'temps de retour des eaux sales
    ES_tps_retour = 18    'h/j

    'hypothèse pour calcul de boues produites
    porosite_lit = 0.3   '     Xavier LE TALLEC
    hauteur_sous_filtre = 0.55   'm     Xavier LE TALLEC
    ratio_eau_non_traite_lit = 0.5    '     Xavier LE TALLEC

    'ratio de N orga dur soluble par rapport à la DCO totale
    ratio_Norgaduresoluble_DCOtot = 0.25 / 100


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "14.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub



Sub lecture_choix()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'lecture des choix internes
    choix_type = Feuil14.Range("Biostyr_choix_type").Cells(1, pi_FiliereConsideree)
    choix_devenir_ES = Feuil14.Range("Biostyr_choix_devenir_ES").Cells(1, pi_FiliereConsideree)
    choix_surface_unitaire = Feuil14.Range("Biostyr_choix_surface_unitaire").Cells(1, pi_FiliereConsideree)
    choix_diametre_media = Feuil14.Range("Biostyr_choix_diametre_media").Cells(1, pi_FiliereConsideree)
    choix_regulation = Feuil14.Range("Biostyr_choix_regulation").Cells(1, pi_FiliereConsideree)
    choix_surpresseur = Feuil14.Range("Biostyr_choix_surpresseur").Cells(1, pi_FiliereConsideree)


    If choix_type = biostyr_C Then
        If choix_diametre_media = 3.6 Or choix_diametre_media = 4 Then
            GoTo erreur_choix_media
        End If
    Else
        If choix_diametre_media = 5 Then
            GoTo erreur_choix_media
        End If
    End If


    Exit Sub
    '*********************************************************************
erreur_choix_media:

    MsgBox Feuil14.Range("biostyr_erreur_nom") & ": " & Feuil14.Range("biostyr_erreur_diametre_media").Cells(1, 1), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
    pn_EffacementResultatsEnCours = True
    pn_ErreurNonGeree = True

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "14.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub lecture_valeurs_forcees()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'lecture du choix de forcer des valeurs
    hauteur_media_force = Feuil14.Range("Biostyr_hauteur_media").Cells(1, pi_FiliereConsideree)
    surface_filtration_min_force = Feuil14.Range("Biostyr_surface_filtration_min").Cells(1, pi_FiliereConsideree)
    nb_cellules_force = Feuil14.Range("Biostyr_nb_cellules").Cells(1, pi_FiliereConsideree)
    Cv_DCO_force = Feuil14.Range("Biostyr_Cv_DCO").Cells(1, pi_FiliereConsideree)
    Cv_nit_force = Feuil14.Range("Biostyr_Cv_nit").Cells(1, pi_FiliereConsideree)
    sortie_DCO_force = Feuil14.Range("Biostyr_sortie_DCO").Cells(1, pi_FiliereConsideree)    ' MSM 15/11/12
    sortie_NH4_force = Feuil14.Range("Biostyr_sortie_NH4").Cells(1, pi_FiliereConsideree)
    sortie_NO3_force = Feuil14.Range("Biostyr_sortie_NO3").Cells(1, pi_FiliereConsideree)
    surpresseur_conso_spec_force = Feuil14.Range("Biostyr_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
    O2_besoin_force = Feuil14.Range("Biostyr_O2_besoin").Cells(1, pi_FiliereConsideree)
    O2_rdt_transfert_force = Feuil14.Range("Biostyr_O2_rdt_transfert").Cells(1, pi_FiliereConsideree)
    air_Q_Nm3j_force = Feuil14.Range("Biostyr_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
    diffuseur_encrassement_force = Feuil14.Range("Biostyr_diffuseur_encrassement").Cells(1, pi_FiliereConsideree)
    air_P_refoulement_force = Feuil14.Range("Biostyr_air_P_refoulement").Cells(1, pi_FiliereConsideree)
    recirculation_taux_force = Feuil14.Range("Biostyr_recirculation_taux").Cells(1, pi_FiliereConsideree)
    recirculation_nb_pompe_force = Feuil14.Range("Biostyr_recirculation_nb_pompe").Cells(1, pi_FiliereConsideree)
    recirculation_P_refoulement_force = Feuil14.Range("Biostyr_recirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
    recirculation_tps_fonctionnement_force = Feuil14.Range("Biostyr_recirculation_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    recirculation_Q_unitaire_force = Feuil14.Range("Biostyr_recirculation_Q_unitaire").Cells(1, pi_FiliereConsideree)
    recirculation_pompe_rdt_force = Feuil14.Range("Biostyr_recirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
    ratio_ES_volume_media_force = Feuil14.Range("Biostyr_ratio_ES_volume_media").Cells(1, pi_FiliereConsideree)
    ES_nb_pompe_force = Feuil14.Range("Biostyr_ES_nb_pompe").Cells(1, pi_FiliereConsideree)
    ES_P_refoulement_force = Feuil14.Range("Biostyr_ES_P_refoulement").Cells(1, pi_FiliereConsideree)
    ES_tps_fonctionnement_force = Feuil14.Range("Biostyr_ES_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    ES_Q_unitaire_force = Feuil14.Range("Biostyr_ES_Q_unitaire").Cells(1, pi_FiliereConsideree)
    ES_pompe_rdt_force = Feuil14.Range("Biostyr_ES_pompe_rdt").Cells(1, pi_FiliereConsideree)
    lavage_air_Q_force = Feuil14.Range("Biostyr_lavage_air_Q").Cells(1, pi_FiliereConsideree)
    lavage_air_P_refoulement_force = Feuil14.Range("Biostyr_lavage_air_P_refoulement").Cells(1, pi_FiliereConsideree)
    lavage_air_tps_fonctionnement_force = Feuil14.Range("Biostyr_lavage_air_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    ES_concentration_force = Feuil14.Range("Biostyr_ES_concentration").Cells(1, pi_FiliereConsideree)
    ES_MV_MES_force = Feuil14.Range("Biostyr_ES_MV_MES").Cells(1, pi_FiliereConsideree)
    ES_MES_force = Feuil14.Range("Biostyr_ES_MES").Cells(1, pi_FiliereConsideree)



    If hauteur_media_force = True Then
        hauteur_media_f = Feuil14.Range("Biostyr_hauteur_media").Cells(1, pi_FiliereConsideree)
    End If
    If surface_filtration_min_force = True Then
        surface_filtration_min_f = Feuil14.Range("Biostyr_surface_filtration_min").Cells(1, pi_FiliereConsideree)
    End If
    If nb_cellules_force = True Then
        nb_cellules_f = Feuil14.Range("Biostyr_nb_cellules").Cells(1, pi_FiliereConsideree)
    End If
    If Cv_DCO_force = True Then
        Cv_DCO_f = Feuil14.Range("Biostyr_Cv_DCO").Cells(1, pi_FiliereConsideree)
    End If
    If Cv_nit_force = True Then
        Cv_nit_f = Feuil14.Range("Biostyr_Cv_nit").Cells(1, pi_FiliereConsideree)
    End If
    'début modif MSM 15/11/12
    If sortie_DCO_force = True Then
        sortie_DCO_f = Feuil14.Range("Biostyr_sortie_DCO").Cells(1, pi_FiliereConsideree)
    End If
    'fin modif MSM 15/11/12
    If sortie_NH4_force = True Then
        sortie_NH4_f = Feuil14.Range("Biostyr_sortie_NH4").Cells(1, pi_FiliereConsideree)
    End If
    If sortie_NO3_force = True Then
        sortie_NO3_f = Feuil14.Range("Biostyr_sortie_NO3").Cells(1, pi_FiliereConsideree)
    End If
    If surpresseur_conso_spec_force = True Then
        surpresseur_conso_spec_f = Feuil14.Range("Biostyr_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
    End If
    If O2_besoin_force = True Then
        O2_besoin_f = Feuil14.Range("Biostyr_O2_besoin").Cells(1, pi_FiliereConsideree)
    End If
    If O2_rdt_transfert_force = True Then
        O2_rdt_transfert_f = Feuil14.Range("Biostyr_O2_rdt_transfert").Cells(1, pi_FiliereConsideree)
    End If
    If air_Q_Nm3j_force = True Then
        air_Q_Nm3j_f = Feuil14.Range("Biostyr_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
    End If
    If diffuseur_encrassement_force = True Then
        diffuseur_encrassement_f = Feuil14.Range("Biostyr_diffuseur_encrassement").Cells(1, pi_FiliereConsideree)
    End If
    If air_P_refoulement_force = True Then
        air_P_refoulement_f = Feuil14.Range("Biostyr_air_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_taux_force = True Then
        recirculation_taux_f = Feuil14.Range("Biostyr_recirculation_taux").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_nb_pompe_force = True Then
        recirculation_nb_pompe_f = Feuil14.Range("Biostyr_recirculation_nb_pompe").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_P_refoulement_force = True Then
        recirculation_P_refoulement_f = Feuil14.Range("Biostyr_recirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_tps_fonctionnement_force = True Then
        recirculation_tps_fonctionnement_f = Feuil14.Range("Biostyr_recirculation_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_Q_unitaire_force = True Then
        recirculation_Q_unitaire_f = Feuil14.Range("Biostyr_recirculation_Q_unitaire").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_pompe_rdt_force = True Then
        recirculation_pompe_rdt_f = Feuil14.Range("Biostyr_recirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    If ratio_ES_volume_media_force = True Then
        ratio_ES_volume_media_f = Feuil14.Range("Biostyr_ratio_ES_volume_media").Cells(1, pi_FiliereConsideree)
    End If
    If ES_nb_pompe_force = True Then
        ES_nb_pompe_f = Feuil14.Range("Biostyr_ES_nb_pompe").Cells(1, pi_FiliereConsideree)
    End If
    If ES_P_refoulement_force = True Then
        ES_P_refoulement_f = Feuil14.Range("Biostyr_ES_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If ES_tps_fonctionnement_force = True Then
        ES_tps_fonctionnement_f = Feuil14.Range("Biostyr_ES_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    End If
    If ES_Q_unitaire_force = True Then
        ES_Q_unitaire_f = Feuil14.Range("Biostyr_ES_Q_unitaire").Cells(1, pi_FiliereConsideree)
    End If
    If ES_pompe_rdt_force = True Then
        ES_pompe_rdt_f = Feuil14.Range("Biostyr_ES_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    If lavage_air_Q_force = True Then
        lavage_air_Q_f = Feuil14.Range("Biostyr_lavage_air_Q").Cells(1, pi_FiliereConsideree)
    End If
    If lavage_air_P_refoulement_force = True Then
        lavage_air_P_refoulement_f = Feuil14.Range("Biostyr_lavage_air_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If lavage_air_tps_fonctionnement_force = True Then
        lavage_air_tps_fonctionnement_f = Feuil14.Range("Biostyr_lavage_air_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    End If
    If ES_concentration_force = True Then
        ES_concentration_f = Feuil14.Range("Biostyr_ES_concentration").Cells(1, pi_FiliereConsideree)
    End If
    If ES_MV_MES_force = True Then
        ES_MV_MES_f = Feuil14.Range("Biostyr_ES_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If ES_MES_force = True Then
        ES_MES_f = Feuil14.Range("Biostyr_ES_MES").Cells(1, pi_FiliereConsideree)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "14.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub ecriture_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil14.Range("biostyr_hauteur_media").Cells(1, Colonne) = hauteur_media
    Feuil14.Range("Biostyr_surface_filtration_min").Cells(1, Colonne) = surface_filtration_min
    Feuil14.Range("Biostyr_nb_cellules").Cells(1, Colonne) = nb_cellules
    Feuil14.Range("Biostyr_surface_filtration").Cells(1, Colonne) = surface_filtration
    Feuil14.Range("Biostyr_Cv_DCO").Cells(1, Colonne) = Cv_DCO
    Feuil14.Range("Biostyr_Cv_nit").Cells(1, Colonne) = Cv_nit
    Feuil14.Range("Biostyr_sortie_DCO").Cells(1, Colonne) = sortie_DCO    ' MSM 15/11/12
    Feuil14.Range("Biostyr_sortie_NH4").Cells(1, Colonne) = sortie_NH4
    Feuil14.Range("Biostyr_sortie_NO3").Cells(1, Colonne) = sortie_NO3
    Feuil14.Range("Biostyr_surpresseur_conso_spec").Cells(1, Colonne) = surpresseur_conso_spec
    Feuil14.Range("Biostyr_O2_besoin").Cells(1, Colonne) = O2_besoin
    Feuil14.Range("Biostyr_O2_rdt_transfert").Cells(1, Colonne) = O2_rdt_transfert
    Feuil14.Range("Biostyr_air_Q_Nm3j").Cells(1, Colonne) = air_Q_Nm3j
    Feuil14.Range("Biostyr_diffuseur_encrassement").Cells(1, Colonne) = diffuseur_encrassement
    Feuil14.Range("Biostyr_air_P_refoulement").Cells(1, Colonne) = air_P_refoulement
    Feuil14.Range("Biostyr_recirculation_taux").Cells(1, Colonne) = recirculation_taux
    Feuil14.Range("Biostyr_recirculation_nb_pompe").Cells(1, Colonne) = recirculation_nb_pompe
    Feuil14.Range("Biostyr_recirculation_P_refoulement").Cells(1, Colonne) = recirculation_P_refoulement
    Feuil14.Range("Biostyr_recirculation_tps_fonctionnement").Cells(1, Colonne) = recirculation_tps_fonctionnement
    Feuil14.Range("Biostyr_recirculation_Q_unitaire").Cells(1, Colonne) = recirculation_Q_unitaire
    Feuil14.Range("Biostyr_recirculation_pompe_rdt").Cells(1, Colonne) = recirculation_pompe_rdt
    Feuil14.Range("Biostyr_ratio_ES_volume_media").Cells(1, Colonne) = ratio_ES_volume_media
    Feuil14.Range("Biostyr_ES_nb_pompe").Cells(1, Colonne) = ES_nb_pompe
    Feuil14.Range("Biostyr_ES_P_refoulement").Cells(1, Colonne) = ES_P_refoulement
    Feuil14.Range("Biostyr_ES_tps_fonctionnement").Cells(1, Colonne) = ES_tps_fonctionnement
    Feuil14.Range("Biostyr_ES_Q_unitaire").Cells(1, Colonne) = ES_Q_unitaire
    Feuil14.Range("Biostyr_ES_pompe_rdt").Cells(1, Colonne) = ES_pompe_rdt
    Feuil14.Range("Biostyr_lavage_air_Q").Cells(1, Colonne) = lavage_air_Q
    Feuil14.Range("Biostyr_lavage_air_P_refoulement").Cells(1, Colonne) = lavage_air_P_refoulement
    Feuil14.Range("Biostyr_lavage_air_tps_fonctionnement").Cells(1, Colonne) = lavage_air_tps_fonctionnement
    Feuil14.Range("Biostyr_ES_concentration").Cells(1, Colonne) = ES_concentration
    Feuil14.Range("Biostyr_ES_MV_MES").Cells(1, Colonne) = ES_MV_MES
    Feuil14.Range("Biostyr_ES_MES").Cells(1, Colonne) = ES_MES
    Feuil14.Range("Biostyr_electricite").Cells(1, Colonne) = electricite
    Feuil14.Range("biostyr_electricite_aeration").Cells(1, Colonne) = electricite_aeration

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "14.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub effacer_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil14.Range("Biostyr_hauteur_media").Cells(1, Colonne) = ""       'hauteur_media
    Feuil14.Range("Biostyr_surface_filtration_min").Cells(1, Colonne) = ""       'surface_filtration_min
    Feuil14.Range("Biostyr_nb_cellules").Cells(1, Colonne) = ""       'nb_cellules
    Feuil14.Range("Biostyr_surface_filtration").Cells(1, Colonne) = ""       'surface_filtration
    Feuil14.Range("Biostyr_Cv_DCO").Cells(1, Colonne) = ""       'Cv_DCO
    Feuil14.Range("Biostyr_Cv_nit").Cells(1, Colonne) = ""       'Cv_nit
    Feuil14.Range("Biostyr_sortie_DCO").Cells(1, Colonne).ClearContents    ' sortie_DCO           MSM 15/11/12
    Feuil14.Range("Biostyr_sortie_NH4").Cells(1, Colonne) = ""       'sortie_NH4
    Feuil14.Range("Biostyr_sortie_NO3").Cells(1, Colonne) = ""       'sortie_NO3
    Feuil14.Range("Biostyr_surpresseur_conso_spec").Cells(1, Colonne) = ""       'surpresseur_conso_spec
    Feuil14.Range("Biostyr_O2_besoin").Cells(1, Colonne) = ""       'O2_besoin
    Feuil14.Range("Biostyr_O2_rdt_transfert").Cells(1, Colonne) = ""       'O2_rdt_transfert
    Feuil14.Range("Biostyr_air_Q_Nm3j").Cells(1, Colonne) = ""       'air_Q_Nm3j
    Feuil14.Range("Biostyr_diffuseur_encrassement").Cells(1, Colonne) = ""       'diffuseur_encrassement
    Feuil14.Range("Biostyr_air_P_refoulement").Cells(1, Colonne) = ""       'air_P_refoulement
    Feuil14.Range("Biostyr_recirculation_taux").Cells(1, Colonne) = ""       'recirculation_taux
    Feuil14.Range("Biostyr_recirculation_nb_pompe").Cells(1, Colonne) = ""       'recirculation_nb_pompe
    Feuil14.Range("Biostyr_recirculation_P_refoulement").Cells(1, Colonne) = ""       'recirculation_P_refoulement
    Feuil14.Range("Biostyr_recirculation_tps_fonctionnement").Cells(1, Colonne) = ""       'recirculation_tps_fonctionnement
    Feuil14.Range("Biostyr_recirculation_Q_unitaire").Cells(1, Colonne) = ""       'recirculation_Q_unitaire
    Feuil14.Range("Biostyr_recirculation_pompe_rdt").Cells(1, Colonne) = ""       'recirculation_pompe_rdt
    Feuil14.Range("Biostyr_ratio_ES_volume_media").Cells(1, Colonne) = ""       'ratio_ES_volume_media
    Feuil14.Range("Biostyr_ES_nb_pompe").Cells(1, Colonne) = ""       'ES_nb_pompe
    Feuil14.Range("Biostyr_ES_P_refoulement").Cells(1, Colonne) = ""       'ES_P_refoulement
    Feuil14.Range("Biostyr_ES_tps_fonctionnement").Cells(1, Colonne) = ""       'ES_tps_fonctionnement
    Feuil14.Range("Biostyr_ES_Q_unitaire").Cells(1, Colonne) = ""       'ES_Q_unitaire
    Feuil14.Range("Biostyr_ES_pompe_rdt").Cells(1, Colonne) = ""       'ES_pompe_rdt
    Feuil14.Range("Biostyr_lavage_air_Q").Cells(1, Colonne) = ""       'lavage_air_Q
    Feuil14.Range("Biostyr_lavage_air_P_refoulement").Cells(1, Colonne) = ""       'lavage_air_P_refoulement
    Feuil14.Range("Biostyr_lavage_air_tps_fonctionnement").Cells(1, Colonne) = ""       'lavage_air_tps_fonctionnement
    Feuil14.Range("Biostyr_ES_concentration").Cells(1, Colonne) = ""       'ES_concentration
    Feuil14.Range("Biostyr_ES_MV_MES").Cells(1, Colonne) = ""       'ES_MV_MES
    Feuil14.Range("Biostyr_ES_MES").Cells(1, Colonne) = ""       'ES_MES
    Feuil14.Range("Biostyr_electricite").Cells(1, Colonne) = ""       'electricite
    Feuil14.Range("biostyr_electricite_aeration").Cells(1, Colonne) = ""        'electricite_aeration

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "14.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub affich_remarque()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 4  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil14.Range("biostyr_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf
    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "14.8"

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









