Option Explicit

Private Const critere_incoherence = 0.1


'erreurs
Private Const nb_erreurs = 2
Private Const erreur_pompe_ES = 1
Private Const erreur_choix_filtrasable = 2

'types de matériaux
Private Const nb_types_materiaux = 7
Private Const mat_ponce = 1
Private Const mat_schiste_TEinf_1m = 2
Private Const mat_schiste_TEinf_2m = 3
Private Const mat_sable_fin_TE1 = 4
Private Const mat_sable_TEsup = 5
Private Const mat_bicouche_sable_ponce_1m = 6
Private Const mat_bicouche_sable_ponce_2m = 7

'retour des eaux sales
Private Const ES_retour_pompage = 1
Private Const ES_retour_gravitaire = 0

'choix de l'utilisateur
Dim choix_materiau As Integer
Dim choix_retour_ES As Integer

'variables échangées aec l'utilisateur
Dim ratio_Q_admis As Double
Dim Q_design As Double
Dim Q_admis As Double
Dim hauteur_materiau As Double
Dim surface_necessaire As Double
Dim surface_unitaire As Double
Dim nb_filtres As Integer
Dim surface_reelle As Double
Dim outlet_MES_mgL As Double
Dim capacite_retention As Double
Dim nb_lavage_jour_filtre As Double  'ce n'est pas un entier car on peut avoir 2 lavages tous les 3 jours par ex
Dim conso_eau_lavage_filtre_lavage As Double
Dim ES_Q As Double
Dim ES_MES As Double
Dim pompe_lavage_Prefoulement As Double
Dim pompe_lavage_nb As Integer
Dim pompe_lavage_rdt_global As Double
Dim air_lavage_vitesse As Double
Dim air_lavage_Prefoulement As Double
Dim air_lavage_tps_fct_lavage As Double
Dim pompe_ES_nb As Integer
Dim pompe_ES_Prefoulement As Double
Dim pompe_ES_tps_fct As Double
Dim pompe_ES_Qunitaire As Double
Dim pompe_ES_rdt_global As Double
Public electricite As Double

'variables échangées aec l'utilisateur
Dim ratio_Q_admis_force As Boolean
Dim Q_design_force As Boolean
Dim hauteur_materiau_force As Boolean
Dim surface_necessaire_force As Boolean
Dim surface_unitaire_force As Boolean
Dim nb_filtres_force As Boolean
Dim outlet_MES_mgL_force As Boolean
Dim capacite_retention_force As Boolean
Dim nb_lavage_jour_filtre_force As Boolean  'ce n'est pas un entier car on peut avoir 2 lavages tous les 3 jours par ex
Dim conso_eau_lavage_filtre_lavage_force As Boolean
Dim pompe_lavage_Prefoulement_force As Boolean
Dim pompe_lavage_nb_force As Boolean
Dim pompe_lavage_rdt_global_force As Boolean
Dim air_lavage_vitesse_force As Boolean
Dim air_lavage_Prefoulement_force As Boolean
Dim air_lavage_tps_fct_lavage_force As Boolean
Dim pompe_ES_nb_force As Boolean
Dim pompe_ES_Prefoulement_force As Boolean
Dim pompe_ES_tps_fct_force As Boolean
Dim pompe_ES_Qunitaire_force As Boolean
Dim pompe_ES_rdt_global_force As Boolean

'variables échangées aec l'utilisateur
Dim ratio_Q_admis_f As Double
Dim Q_design_f As Double
Dim hauteur_materiau_f As Double
Dim surface_necessaire_f As Double
Dim surface_unitaire_f As Double
Dim nb_filtres_f As Integer
Dim outlet_MES_mgL_f As Double
Dim capacite_retention_f As Double
Dim nb_lavage_jour_filtre_f As Double  'ce n'est pas un entier car on peut avoir 2 lavages tous les 3 jours par ex
Dim conso_eau_lavage_filtre_lavage_f As Double
Dim pompe_lavage_Prefoulement_f As Double
Dim pompe_lavage_nb_f As Integer
Dim pompe_lavage_rdt_global_f As Double
Dim air_lavage_vitesse_f As Double
Dim air_lavage_Prefoulement_f As Double
Dim air_lavage_tps_fct_lavage_f As Double
Dim pompe_ES_nb_f As Integer
Dim pompe_ES_Prefoulement_f As Double
Dim pompe_ES_tps_fct_f As Double
Dim pompe_ES_Qunitaire_f As Double
Dim pompe_ES_rdt_global_f As Double

'hypothèses
Dim hyp_P_eau_lavage(nb_types_materiaux) As Double
Dim hyp_P_air_lavage(nb_types_materiaux) As Double
Dim hyp_conso_eau_lavage_m3_m2_lavage(nb_types_materiaux) As Double
Dim hyp_capacite_retention(nb_types_materiaux) As Double
Dim hyp_vitesse_filtration(nb_types_materiaux) As Double
Dim hyp_hauteur_materiau(nb_types_materiaux) As Double
Dim hyp_vitesse_eau_lavage(nb_types_materiaux) As Double
Dim hyp_vitesse_air_lavage(nb_types_materiaux) As Double
Dim surface_unitaire_maximale As Double
Dim nb_filtres_mini As Integer
Dim ES_tps_retour As Double
Dim surpresseur_conso_spec As Double


Dim remarque(nb_erreurs) As String

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer

    'hypothèses dépendantes du type de matériau
    For I = 1 To nb_types_materiaux
        hyp_P_eau_lavage(I) = Feuil48.Range("hyp_III_filtrasable_P_eau_lavage").Cells(I, 1)
        hyp_P_air_lavage(I) = Feuil48.Range("hyp_III_filtrasable_P_air_lavage").Cells(I, 1)
        hyp_conso_eau_lavage_m3_m2_lavage(I) = Feuil48.Range("hyp_III_filtrasable_eau_lavage_m3_m2_lavage").Cells(I, 1)
        hyp_capacite_retention(I) = Feuil48.Range("hyp_III_filtrasable_capacite_retention").Cells(I, 1)
        hyp_vitesse_filtration(I) = Feuil48.Range("hyp_III_filtrasable_vitesse_filtration").Cells(I, 1)
        hyp_hauteur_materiau(I) = Feuil48.Range("hyp_III_filtrasable_hauteur_materiau").Cells(I, 1)
        hyp_vitesse_eau_lavage(I) = Feuil48.Range("hyp_III_filtrasable_eau_lavage_vitesse").Cells(I, 1)
        hyp_vitesse_air_lavage(I) = Feuil48.Range("hyp_III_filtrasable_air_lavage_vitesse").Cells(I, 1)
    Next I

    'surface maximale d'une cellule de filtration
    surface_unitaire_maximale = Feuil48.Range("hyp_III_filtrasable_surface_unitaire_max").Cells(1, 1)   '125     'm2

    'nombre de filtres minimum
    nb_filtres_mini = Feuil48.Range("hyp_III_filtrasable_nb_filtres_mini").Cells(1, 1)  '4

    ES_tps_retour = Feuil48.Range("hyp_III_filtrasable_pompe_ES_tps_fct").Cells(1, 1)    'h/j

    surpresseur_conso_spec = 4.5    'Wh/(m3.mCE)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "50.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub attribution_valeur_par_defaut()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On retrouve ici l'ensemble des valeurs pour lesquelles on attribue une valeur guide par défaut

    'part de débit traité
    If ratio_Q_admis_force = True Then
        ratio_Q_admis = ratio_Q_admis_f
    Else
        ratio_Q_admis = 1
    End If

    'hauteur matériau
    If hauteur_materiau_force = True Then
        hauteur_materiau = hauteur_materiau_f
    Else
        hauteur_materiau = hyp_hauteur_materiau(choix_materiau)
    End If

    'sortie en MES eau traitée
    If outlet_MES_mgL_force = True Then
        outlet_MES_mgL = outlet_MES_mgL_f
    Else
        outlet_MES_mgL = 2  'mg/L
    End If

    'capacité de rétention
    If capacite_retention_force = True Then
        capacite_retention = capacite_retention_f
    Else
        capacite_retention = hyp_capacite_retention(choix_materiau)
    End If

    'eau de lavage utilisée
    If conso_eau_lavage_filtre_lavage_force = True Then
        conso_eau_lavage_filtre_lavage = conso_eau_lavage_filtre_lavage_f
    Else
        conso_eau_lavage_filtre_lavage = hyp_conso_eau_lavage_m3_m2_lavage(choix_materiau)
    End If

    'pression pompe eau de lavage
    If pompe_lavage_Prefoulement_force = True Then
        pompe_lavage_Prefoulement = pompe_lavage_Prefoulement_f
    Else
        pompe_lavage_Prefoulement = hyp_P_eau_lavage(choix_materiau)
    End If

    'nombre pompes de lavage
    If pompe_lavage_nb_force = True Then
        pompe_lavage_nb = pompe_lavage_nb_f
    Else
        pompe_lavage_nb = 1
    End If

    'vitesse air de lavage
    If air_lavage_vitesse_force = True Then
        air_lavage_vitesse = air_lavage_vitesse_f
    Else
        air_lavage_vitesse = hyp_vitesse_air_lavage(choix_materiau)
    End If

    'P refoulement air de lavage
    If air_lavage_Prefoulement_force = True Then
        air_lavage_Prefoulement = air_lavage_Prefoulement_f
    Else
        air_lavage_Prefoulement = hyp_P_air_lavage(choix_materiau)
    End If

    'temps de fonctionnement air de lavage par lavage
    If air_lavage_tps_fct_lavage_force = True Then
        air_lavage_tps_fct_lavage = air_lavage_tps_fct_lavage_f
    Else
        air_lavage_tps_fct_lavage = Feuil48.Range("hyp_III_filtrasable_air_lavage_tps_fct_min_lavage").Cells(1, 1)   '9min
    End If

    'pompage des eaux sales retournées en tête
    If choix_retour_ES = ES_retour_pompage Then
        If pompe_ES_nb_force = True Then
            pompe_ES_nb = pompe_ES_nb_f
        Else
            pompe_ES_nb = 2
        End If
        If pompe_ES_Prefoulement_force = True Then
            pompe_ES_Prefoulement = pompe_ES_Prefoulement_f
        Else
            pompe_ES_Prefoulement = 10   'mCE
        End If
        If pompe_ES_tps_fct_force = True Then
            pompe_ES_tps_fct = pompe_ES_tps_fct_f
        Else
            pompe_ES_tps_fct = ES_tps_retour    '18h/j
        End If
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "50.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub dimensionnement(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef eaux_sales_filtrasable)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Q_traite As Double
    Dim DCO_traite As Double
    Dim DBO_traite As Double
    Dim MES_traite As Double
    Dim NK_traite As Double
    Dim NH4_traite As Double
    Dim NO3_traite As Double
    Dim Pt_traite As Double
    Dim interm_nb_filtres As Double

    'Débit et polluants réellement admis
    Q_traite = Q * ratio_Q_admis
    Q_admis = Q_traite
    DCO_traite = DCO * ratio_Q_admis
    DBO_traite = DBO * ratio_Q_admis
    MES_traite = MES * ratio_Q_admis
    NK_traite = NK * ratio_Q_admis
    NH4_traite = NH4 * ratio_Q_admis
    NO3_traite = NO3 * ratio_Q_admis
    Pt_traite = Pt * ratio_Q_admis


    'Débit de design
    If Q_design_force = True Then
        Q_design = Q_design_f
    Else
        Q_design = Q_traite * pd_CoefficientPointeHydrauliqueTempsPluie / pd_NOMBRE_HEURE_PAR_JOUR
    End If

    'calcul de la surface nécessaire de filtration
    If surface_necessaire_force = True Then
        surface_necessaire = surface_necessaire_f
    Else
        surface_necessaire = Q_design / hyp_vitesse_filtration(choix_materiau)
    End If

    'Calcul du nb de filtres et de la surface unitaire
    If nb_filtres_force = True Then
        nb_filtres = nb_filtres_f
        If surface_unitaire_force = True Then
            surface_unitaire = surface_unitaire_f
        Else
            surface_unitaire = surface_necessaire / nb_filtres
        End If
    Else
        If surface_unitaire_force = True Then
            surface_unitaire = surface_unitaire_f
            'calcul du nb de filtres
            interm_nb_filtres = surface_necessaire / surface_unitaire
            nb_filtres = Int(interm_nb_filtres)
            If nb_filtres < interm_nb_filtres Then
                nb_filtres = nb_filtres + 1
            End If
        Else
            If surface_unitaire_maximale * nb_filtres_mini <= surface_necessaire Then
                'on fixe la surface maximale unitaire possible et on calcule le nombre de filtres
                surface_unitaire = surface_unitaire_maximale
                interm_nb_filtres = surface_necessaire / surface_unitaire
                nb_filtres = Int(interm_nb_filtres)
                If nb_filtres < interm_nb_filtres Then
                    nb_filtres = nb_filtres + 1
                End If
            Else
                'on fixe un minimum de 4 cellules et on recalcule une surface unitaire
                nb_filtres = nb_filtres_mini
                surface_unitaire = surface_necessaire / nb_filtres
            End If
        End If
    End If

    'calcul de la surface réelle de filtration
    surface_reelle = nb_filtres * surface_unitaire

    'calcul du nombre de lavages journaliers
    If MES_traite / Q_traite * 1000 < outlet_MES_mgL Then
        outlet_MES_mgL = MES_traite / Q_traite * 1000
        ES_MES = 0
    Else
        ES_MES = MES_traite - outlet_MES_mgL * (Q_traite - ES_Q) / 1000
    End If
    nb_lavage_jour_filtre = ES_MES / (capacite_retention * (surface_reelle * hauteur_materiau))

    ES_Q = nb_lavage_jour_filtre * surface_reelle * conso_eau_lavage_filtre_lavage

    'caractéristiques des eaux sales
    eaux_sales_filtrasable(repere_ES_Q) = ES_Q
    eaux_sales_filtrasable(repere_ES_MES) = ES_MES
    eaux_sales_filtrasable(repere_ES_NH4) = NH4 * ES_Q / Q
    eaux_sales_filtrasable(repere_ES_NO3) = NO3 * ES_Q / Q
    eaux_sales_filtrasable(repere_ES_DCO) = DCO * ES_Q / Q
    eaux_sales_filtrasable(repere_ES_DBO) = DBO * ES_Q / Q
    eaux_sales_filtrasable(repere_ES_NK) = NK * ES_Q / Q
    eaux_sales_filtrasable(repere_ES_Pt) = Pt * ES_Q / Q

    'eau en sortie
    Q = (1 - ratio_Q_admis) * Q + (Q_traite - eaux_sales_filtrasable(repere_ES_Q))
    MES = (1 - ratio_Q_admis) * MES + (MES_traite - eaux_sales_filtrasable(repere_ES_MES))
    DCO = (1 - ratio_Q_admis) * DCO + (DCO_traite - eaux_sales_filtrasable(repere_ES_DCO))
    DBO = (1 - ratio_Q_admis) * DBO + (DBO_traite - eaux_sales_filtrasable(repere_ES_DBO))
    NK = (1 - ratio_Q_admis) * NK + (NK_traite - eaux_sales_filtrasable(repere_ES_NK))
    NH4 = (1 - ratio_Q_admis) * NH4 + (NH4_traite - eaux_sales_filtrasable(repere_ES_NH4))
    NO3 = (1 - ratio_Q_admis) * NO3 + (NO3_traite - eaux_sales_filtrasable(repere_ES_NO3))
    Pt = (1 - ratio_Q_admis) * Pt + (Pt_traite - eaux_sales_filtrasable(repere_ES_Pt))



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "50.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub fonctionnement_reel(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef eaux_sales_filtrasable)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Q_traite As Double
    Dim DCO_traite As Double
    Dim DBO_traite As Double
    Dim MES_traite As Double
    Dim NK_traite As Double
    Dim NH4_traite As Double
    Dim NO3_traite As Double
    Dim Pt_traite As Double
    Dim interm_nb_filtres As Double

    'Débit et polluants réellement admis
    Q_traite = Q * ratio_Q_admis
    Q_admis = Q_traite
    DCO_traite = DCO * ratio_Q_admis
    DBO_traite = DBO * ratio_Q_admis
    MES_traite = MES * ratio_Q_admis
    NK_traite = NK * ratio_Q_admis
    NH4_traite = NH4 * ratio_Q_admis
    NO3_traite = NO3 * ratio_Q_admis
    Pt_traite = Pt * ratio_Q_admis




    'calcul du nombre de lavages journaliers
    If MES_traite / Q_traite * 1000 < outlet_MES_mgL Then
        outlet_MES_mgL = MES_traite / Q_traite * 1000
        ES_MES = 0
        remarque(erreur_choix_filtrasable) = Feuil50.Range("filtrasable_erreur_choix_filtrasable").Cells(1, 1)
    Else
        ES_MES = MES_traite - outlet_MES_mgL * (Q_traite - ES_Q) / 1000
    End If
    nb_lavage_jour_filtre = ES_MES / (capacite_retention * (surface_reelle * hauteur_materiau))

    ES_Q = nb_lavage_jour_filtre * surface_reelle * conso_eau_lavage_filtre_lavage

    'caractéristiques des eaux sales
    eaux_sales_filtrasable(repere_ES_Q) = ES_Q
    eaux_sales_filtrasable(repere_ES_MES) = ES_MES
    eaux_sales_filtrasable(repere_ES_NH4) = NH4 * ES_Q / Q
    eaux_sales_filtrasable(repere_ES_NO3) = NO3 * ES_Q / Q
    eaux_sales_filtrasable(repere_ES_DCO) = DCO * ES_Q / Q
    eaux_sales_filtrasable(repere_ES_DBO) = DBO * ES_Q / Q
    eaux_sales_filtrasable(repere_ES_NK) = NK * ES_Q / Q
    eaux_sales_filtrasable(repere_ES_Pt) = Pt * ES_Q / Q


    'eau en sortie
    Q = (1 - ratio_Q_admis) * Q + (Q_traite - eaux_sales_filtrasable(repere_ES_Q))
    MES = (1 - ratio_Q_admis) * MES + (MES_traite - eaux_sales_filtrasable(repere_ES_MES))
    DCO = (1 - ratio_Q_admis) * DCO + (DCO_traite - eaux_sales_filtrasable(repere_ES_DCO))
    DBO = (1 - ratio_Q_admis) * DBO + (DBO_traite - eaux_sales_filtrasable(repere_ES_DBO))
    NK = (1 - ratio_Q_admis) * NK + (NK_traite - eaux_sales_filtrasable(repere_ES_NK))
    NH4 = (1 - ratio_Q_admis) * NH4 + (NH4_traite - eaux_sales_filtrasable(repere_ES_NH4))
    NO3 = (1 - ratio_Q_admis) * NO3 + (NO3_traite - eaux_sales_filtrasable(repere_ES_NO3))
    Pt = (1 - ratio_Q_admis) * Pt + (Pt_traite - eaux_sales_filtrasable(repere_ES_Pt))




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "50.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub calcul_consommation_electrique()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim EL_Q_unitaire As Double
    Dim rendement_pompe As Double
    Dim puissance_pompe As Double
    Dim rendement_moteur As Double

    Dim electricite_pompe_EL As Double   'eaux de lavage
    Dim electricite_pompe_ES As Double    'eaux sales
    Dim electricite_air_lavage As Double

    'erreurs
    Dim incoherence As Double

    electricite = 0

    'Pompe(s) de lavage
    If nb_lavage_jour_filtre > 0 Then
        If pompe_lavage_rdt_global_force = True Then
            pompe_lavage_rdt_global = pompe_lavage_rdt_global_f
        Else
            EL_Q_unitaire = hyp_vitesse_eau_lavage(choix_materiau) * surface_reelle
            rendement_pompe = 0.7
            puissance_pompe = EL_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * pompe_lavage_Prefoulement / rendement_pompe
            If puissance_pompe < 60 Then
                rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
            Else
                rendement_moteur = 0.961
            End If
            pompe_lavage_rdt_global = rendement_pompe * rendement_moteur
        End If
        electricite_pompe_EL = ES_Q / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * pompe_lavage_Prefoulement / pompe_lavage_rdt_global
    Else
        electricite_pompe_EL = 0
    End If
    electricite = electricite + electricite_pompe_EL

    'Pompes eaux sales
    If nb_lavage_jour_filtre > 0 Then
        If choix_retour_ES = ES_retour_pompage Then
            If pompe_ES_Qunitaire_force = True Then
                pompe_ES_Qunitaire = pompe_ES_Qunitaire_f
                If pompe_ES_tps_fct_force = True Then
                    'on ne fait rien
                    If pompe_ES_nb_force = False Then
                        If Int(ES_Q / (pompe_ES_tps_fct * pompe_ES_Qunitaire)) < (ES_Q / (pompe_ES_tps_fct * pompe_ES_Qunitaire)) Then
                            pompe_ES_nb = Int(ES_Q / (pompe_ES_tps_fct * pompe_ES_Qunitaire)) + 1
                        Else
                            pompe_ES_nb = Int(ES_Q / (pompe_ES_tps_fct * pompe_ES_Qunitaire))
                        End If
                    End If
                Else
                    pompe_ES_tps_fct = ES_Q / (pompe_ES_nb * pompe_ES_Qunitaire)
                    If pompe_ES_tps_fct > ES_tps_retour Then
                        pompe_ES_tps_fct = ES_tps_retour
                        If pompe_ES_nb_force = False Then
                            If Int(ES_Q / (pompe_ES_tps_fct * pompe_ES_Qunitaire)) < (ES_Q / (pompe_ES_tps_fct * pompe_ES_Qunitaire)) Then
                                pompe_ES_nb = Int(ES_Q / (pompe_ES_tps_fct * pompe_ES_Qunitaire)) + 1
                            Else
                                pompe_ES_nb = Int(ES_Q / (pompe_ES_tps_fct * pompe_ES_Qunitaire))
                            End If
                            pompe_ES_tps_fct = ES_Q / (pompe_ES_nb * pompe_ES_Qunitaire)
                        End If
                    End If
                End If
            Else
                pompe_ES_Qunitaire = ES_Q / (pompe_ES_nb * pompe_ES_tps_fct)
            End If
            incoherence = (pompe_ES_Qunitaire * pompe_ES_nb * pompe_ES_tps_fct - ES_Q) / ES_Q
            If Abs(incoherence) > critere_incoherence Then
                remarque(erreur_pompe_ES) = Feuil50.Range("filtrasable_erreur_pompe_ES").Cells(1, 1)
            End If
            If pompe_ES_rdt_global_force = True Then
                pompe_ES_rdt_global = pompe_ES_rdt_global_f
            Else
                rendement_pompe = 0.7    'hypothèse
                puissance_pompe = pompe_ES_Qunitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * pompe_ES_Prefoulement / rendement_pompe
                If puissance_pompe < 60 Then
                    rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
                Else
                    rendement_moteur = 0.961
                End If
                pompe_ES_rdt_global = rendement_pompe * rendement_moteur
            End If
            electricite_pompe_ES = pompe_ES_Qunitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * pompe_ES_Prefoulement / pompe_ES_rdt_global * pompe_ES_nb * pompe_ES_tps_fct
            electricite = electricite + electricite_pompe_ES
        End If
    End If

    'air de lavage
    electricite_air_lavage = air_lavage_vitesse * surface_reelle * (air_lavage_tps_fct_lavage / pd_NOMBRE_MINUTE_PAR_HEURE) * nb_lavage_jour_filtre * air_lavage_Prefoulement * surpresseur_conso_spec / 1000
    electricite = electricite + electricite_air_lavage


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "50.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub affich_remarque()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = nb_erreurs  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil50.Range("filtrasable_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "50.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub lecture_choix()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'lecture des choix internes
    choix_materiau = Feuil50.Range("filtrasable_choix_materiau").Cells(1, pi_FiliereConsideree)
    choix_retour_ES = Feuil50.Range("filtrasable_choix_retour_ES").Cells(1, pi_FiliereConsideree)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "50.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub lecture_valeurs_forcees()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    ratio_Q_admis_force = Feuil50.Range("filtrasable_ratio_Q_admis").Cells(1, pi_FiliereConsideree)
    Q_design_force = Feuil50.Range("filtrasable_Q_design").Cells(1, pi_FiliereConsideree)
    hauteur_materiau_force = Feuil50.Range("filtrasable_hauteur_materiau").Cells(1, pi_FiliereConsideree)
    surface_necessaire_force = Feuil50.Range("filtrasable_surface_necessaire").Cells(1, pi_FiliereConsideree)
    surface_unitaire_force = Feuil50.Range("filtrasable_surface_unitaire").Cells(1, pi_FiliereConsideree)
    nb_filtres_force = Feuil50.Range("filtrasable_nb_filtres").Cells(1, pi_FiliereConsideree)
    outlet_MES_mgL_force = Feuil50.Range("filtrasable_outlet_MES_mgL").Cells(1, pi_FiliereConsideree)
    capacite_retention_force = Feuil50.Range("filtrasable_capacite_retention").Cells(1, pi_FiliereConsideree)
    nb_lavage_jour_filtre_force = Feuil50.Range("filtrasable_nb_lavage_jour_filtre").Cells(1, pi_FiliereConsideree)
    conso_eau_lavage_filtre_lavage_force = Feuil50.Range("filtrasable_conso_eau_lavage_filtre_lavage").Cells(1, pi_FiliereConsideree)
    pompe_lavage_Prefoulement_force = Feuil50.Range("filtrasable_pompe_lavage_Prefoulement").Cells(1, pi_FiliereConsideree)
    pompe_lavage_nb_force = Feuil50.Range("filtrasable_pompe_lavage_nb").Cells(1, pi_FiliereConsideree)
    pompe_lavage_rdt_global_force = Feuil50.Range("filtrasable_pompe_lavage_rdt_global").Cells(1, pi_FiliereConsideree)
    air_lavage_vitesse_force = Feuil50.Range("filtrasable_air_lavage_vitesse").Cells(1, pi_FiliereConsideree)
    air_lavage_Prefoulement_force = Feuil50.Range("filtrasable_air_lavage_Prefoulement").Cells(1, pi_FiliereConsideree)
    air_lavage_tps_fct_lavage_force = Feuil50.Range("filtrasable_air_lavage_tps_fct_lavage").Cells(1, pi_FiliereConsideree)
    If choix_retour_ES = ES_retour_pompage Then
        pompe_ES_nb_force = Feuil50.Range("filtrasable_pompe_ES_nb").Cells(1, pi_FiliereConsideree)
        pompe_ES_Prefoulement_force = Feuil50.Range("filtrasable_pompe_ES_Prefoulement").Cells(1, pi_FiliereConsideree)
        pompe_ES_tps_fct_force = Feuil50.Range("filtrasable_pompe_ES_tps_fct").Cells(1, pi_FiliereConsideree)
        pompe_ES_Qunitaire_force = Feuil50.Range("filtrasable_pompe_ES_Qunitaire").Cells(1, pi_FiliereConsideree)
        pompe_ES_rdt_global_force = Feuil50.Range("filtrasable_pompe_ES_rdt_global").Cells(1, pi_FiliereConsideree)
    End If


    If ratio_Q_admis_force = True Then
        ratio_Q_admis_f = Feuil50.Range("filtrasable_ratio_Q_admis").Cells(1, pi_FiliereConsideree)
    End If
    If Q_design_force = True Then
        Q_design_f = Feuil50.Range("filtrasable_Q_design").Cells(1, pi_FiliereConsideree)
    End If
    If hauteur_materiau_force = True Then
        hauteur_materiau_f = Feuil50.Range("filtrasable_hauteur_materiau").Cells(1, pi_FiliereConsideree)
    End If
    If surface_necessaire_force = True Then
        surface_necessaire_f = Feuil50.Range("filtrasable_surface_necessaire").Cells(1, pi_FiliereConsideree)
    End If
    If surface_unitaire_force = True Then
        surface_unitaire_f = Feuil50.Range("filtrasable_surface_unitaire").Cells(1, pi_FiliereConsideree)
    End If
    If nb_filtres_force = True Then
        nb_filtres_f = Feuil50.Range("filtrasable_nb_filtres").Cells(1, pi_FiliereConsideree)
    End If
    If outlet_MES_mgL_force = True Then
        outlet_MES_mgL_f = Feuil50.Range("filtrasable_outlet_MES_mgL").Cells(1, pi_FiliereConsideree)
    End If
    If capacite_retention_force = True Then
        capacite_retention_f = Feuil50.Range("filtrasable_capacite_retention").Cells(1, pi_FiliereConsideree)
    End If
    If nb_lavage_jour_filtre_force = True Then
        nb_lavage_jour_filtre_f = Feuil50.Range("filtrasable_nb_lavage_jour_filtre").Cells(1, pi_FiliereConsideree)
    End If
    If conso_eau_lavage_filtre_lavage_force = True Then
        conso_eau_lavage_filtre_lavage_f = Feuil50.Range("filtrasable_conso_eau_lavage_filtre_lavage").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_lavage_Prefoulement_force = True Then
        pompe_lavage_Prefoulement_f = Feuil50.Range("filtrasable_pompe_lavage_Prefoulement").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_lavage_nb_force = True Then
        pompe_lavage_nb_f = Feuil50.Range("filtrasable_pompe_lavage_nb").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_lavage_rdt_global_force = True Then
        pompe_lavage_rdt_global_f = Feuil50.Range("filtrasable_pompe_lavage_rdt_global").Cells(1, pi_FiliereConsideree)
    End If
    If air_lavage_vitesse_force = True Then
        air_lavage_vitesse_f = Feuil50.Range("filtrasable_air_lavage_vitesse").Cells(1, pi_FiliereConsideree)
    End If
    If air_lavage_Prefoulement_force = True Then
        air_lavage_Prefoulement_f = Feuil50.Range("filtrasable_air_lavage_Prefoulement").Cells(1, pi_FiliereConsideree)
    End If
    If air_lavage_tps_fct_lavage_force = True Then
        air_lavage_tps_fct_lavage_f = Feuil50.Range("filtrasable_air_lavage_tps_fct_lavage").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_ES_nb_force = True Then
        pompe_ES_nb_f = Feuil50.Range("filtrasable_pompe_ES_nb").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_ES_Prefoulement_force = True Then
        pompe_ES_Prefoulement_f = Feuil50.Range("filtrasable_pompe_ES_Prefoulement").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_ES_tps_fct_force = True Then
        pompe_ES_tps_fct_f = Feuil50.Range("filtrasable_pompe_ES_tps_fct").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_ES_Qunitaire_force = True Then
        pompe_ES_Qunitaire_f = Feuil50.Range("filtrasable_pompe_ES_Qunitaire").Cells(1, pi_FiliereConsideree)
    End If
    If pompe_ES_rdt_global_force = True Then
        pompe_ES_rdt_global_f = Feuil50.Range("filtrasable_pompe_ES_rdt_global").Cells(1, pi_FiliereConsideree)
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "50.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub ecriture_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil50.Range("filtrasable_ratio_Q_admis").Cells(1, Colonne) = ratio_Q_admis
    Feuil50.Range("filtrasable_Q_design").Cells(1, Colonne) = Q_design
    Feuil50.Range("filtrasable_Q_admis").Cells(1, Colonne) = Q_admis
    Feuil50.Range("filtrasable_hauteur_materiau").Cells(1, Colonne) = hauteur_materiau
    Feuil50.Range("filtrasable_surface_necessaire").Cells(1, Colonne) = surface_necessaire
    Feuil50.Range("filtrasable_surface_unitaire").Cells(1, Colonne) = surface_unitaire
    Feuil50.Range("filtrasable_nb_filtres").Cells(1, Colonne) = nb_filtres
    Feuil50.Range("filtrasable_surface_reelle").Cells(1, Colonne) = surface_reelle
    Feuil50.Range("filtrasable_outlet_MES_mgL").Cells(1, Colonne) = outlet_MES_mgL
    Feuil50.Range("filtrasable_capacite_retention").Cells(1, Colonne) = capacite_retention
    Feuil50.Range("filtrasable_nb_lavage_jour_filtre").Cells(1, Colonne) = nb_lavage_jour_filtre
    Feuil50.Range("filtrasable_conso_eau_lavage_filtre_lavage").Cells(1, Colonne) = conso_eau_lavage_filtre_lavage
    Feuil50.Range("filtrasable_ES_Q").Cells(1, Colonne) = ES_Q
    Feuil50.Range("filtrasable_ES_MES").Cells(1, Colonne) = ES_MES
    Feuil50.Range("filtrasable_pompe_lavage_Prefoulement").Cells(1, Colonne) = pompe_lavage_Prefoulement
    Feuil50.Range("filtrasable_pompe_lavage_nb").Cells(1, Colonne) = pompe_lavage_nb
    Feuil50.Range("filtrasable_pompe_lavage_rdt_global").Cells(1, Colonne) = pompe_lavage_rdt_global
    Feuil50.Range("filtrasable_air_lavage_vitesse").Cells(1, Colonne) = air_lavage_vitesse
    Feuil50.Range("filtrasable_air_lavage_Prefoulement").Cells(1, Colonne) = air_lavage_Prefoulement
    Feuil50.Range("filtrasable_air_lavage_tps_fct_lavage").Cells(1, Colonne) = air_lavage_tps_fct_lavage
    If choix_retour_ES = ES_retour_pompage Then
        Feuil50.Range("filtrasable_pompe_ES_nb").Cells(1, Colonne) = pompe_ES_nb
        Feuil50.Range("filtrasable_pompe_ES_Prefoulement").Cells(1, Colonne) = pompe_ES_Prefoulement
        Feuil50.Range("filtrasable_pompe_ES_tps_fct").Cells(1, Colonne) = pompe_ES_tps_fct
        Feuil50.Range("filtrasable_pompe_ES_Qunitaire").Cells(1, Colonne) = pompe_ES_Qunitaire
        Feuil50.Range("filtrasable_pompe_ES_rdt_global").Cells(1, Colonne) = pompe_ES_rdt_global
    End If
    Feuil50.Range("filtrasable_electricite").Cells(1, Colonne) = electricite



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "50.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub effacer_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On efface les valeurs sur Excel
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil50.Range("filtrasable_ratio_Q_admis").Cells(1, Colonne) = ""             'ratio_Q_admis
    Feuil50.Range("filtrasable_Q_design").Cells(1, Colonne) = ""             'Q_design
    Feuil50.Range("filtrasable_Q_admis").Cells(1, Colonne) = ""             'Q_admis
    Feuil50.Range("filtrasable_hauteur_materiau").Cells(1, Colonne) = ""             'hauteur_materiau
    Feuil50.Range("filtrasable_surface_necessaire").Cells(1, Colonne) = ""             'surface_necessaire
    Feuil50.Range("filtrasable_surface_unitaire").Cells(1, Colonne) = ""             'surface_unitaire
    Feuil50.Range("filtrasable_nb_filtres").Cells(1, Colonne) = ""             'nb_filtres
    Feuil50.Range("filtrasable_surface_reelle").Cells(1, Colonne) = ""             'surface_reelle
    Feuil50.Range("filtrasable_outlet_MES_mgL").Cells(1, Colonne) = ""             'outlet_MES_mgL
    Feuil50.Range("filtrasable_capacite_retention").Cells(1, Colonne) = ""             'capacite_retention
    Feuil50.Range("filtrasable_nb_lavage_jour_filtre").Cells(1, Colonne) = ""             'nb_lavage_jour_filtre
    Feuil50.Range("filtrasable_conso_eau_lavage_filtre_lavage").Cells(1, Colonne) = ""             'conso_eau_lavage_filtre_lavage
    Feuil50.Range("filtrasable_ES_Q").Cells(1, Colonne) = ""             'ES_Q
    Feuil50.Range("filtrasable_ES_MES").Cells(1, Colonne) = ""             'ES_MES
    Feuil50.Range("filtrasable_pompe_lavage_Prefoulement").Cells(1, Colonne) = ""             'pompe_lavage_Prefoulement
    Feuil50.Range("filtrasable_pompe_lavage_nb").Cells(1, Colonne) = ""             'pompe_lavage_nb
    Feuil50.Range("filtrasable_pompe_lavage_rdt_global").Cells(1, Colonne) = ""             'pompe_lavage_rdt_global
    Feuil50.Range("filtrasable_air_lavage_vitesse").Cells(1, Colonne) = ""             'air_lavage_vitesse
    Feuil50.Range("filtrasable_air_lavage_Prefoulement").Cells(1, Colonne) = ""             'air_lavage_Prefoulement
    Feuil50.Range("filtrasable_air_lavage_tps_fct_lavage").Cells(1, Colonne) = ""             'air_lavage_tps_fct_lavage
    Feuil50.Range("filtrasable_pompe_ES_nb").Cells(1, Colonne) = ""             'pompe_ES_nb
    Feuil50.Range("filtrasable_pompe_ES_Prefoulement").Cells(1, Colonne) = ""             'pompe_ES_Prefoulement
    Feuil50.Range("filtrasable_pompe_ES_tps_fct").Cells(1, Colonne) = ""             'pompe_ES_tps_fct
    Feuil50.Range("filtrasable_pompe_ES_Qunitaire").Cells(1, Colonne) = ""             'pompe_ES_Qunitaire
    Feuil50.Range("filtrasable_pompe_ES_rdt_global").Cells(1, Colonne) = ""             'pompe_ES_rdt_global
    Feuil50.Range("filtrasable_electricite").Cells(1, Colonne) = ""             'electricite


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "50.10"

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


