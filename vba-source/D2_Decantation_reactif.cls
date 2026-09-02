Option Explicit


'constantes pour le choix de dimensionnement
Private Const dimensionne_100 = 1
Private Const dimensionne_bypass = 2

'constantes pour le choix des lamelles
Private Const lamelle_oui = 1
Private Const lamelle_non = 2

'constantes pour le choix de l'injection de microsable
Private Const microsable_oui = 1
Private Const microsable_non = 2

'constantes pour le choix de la recirculation (ACTIFLO)
Private Const recirculation_oui = 1
Private Const recirculation_non = 2

Private Const critere_incoherence = 0.1

'variables de choix internes
Dim choix_dimension_bypass As Integer
Dim choix_lamellaire As Integer
Public choix_polymere_type As Integer
Dim choix_microsable As Integer
Dim choix_recirculation As Integer

'variables échangées avec Excel
Dim bypass As Double
Public dosage_FeCl3 As Double
Public dosage_polymere As Double
Dim mi_NombreOuvrages As Integer
Dim V_mirroir_limite As Double
Dim S_mirroir As Double
Dim V_mirroir_limite_recalculee As Double
Dim hauteur As Double
Dim V_mirroir_moyenne As Double
Dim temps_retention As Double
Dim rdt_MES As Double
Dim rdt_DCO As Double
Dim rdt_DBO As Double
Dim rdt_NK As Double
Dim rdt_P As Double
Public MV_MES As Double
Dim recirculation_taux As Double
Dim recirculation_nb_pompe As Integer
Dim recirculation_P_refoulement As Double
Dim recirculation_tps_fct_pompe As Double
Dim recirculation_Q_unitaire As Double
Dim recirculation_pompe_rdt As Double
Dim boues_concentration As Double
Public boues_MES As Double
Public boues_Q As Double
Dim extraction_nb_pompe As Integer
Dim extraction_P_refoulement As Double
Dim extraction_tps_fct_pompe As Double
Dim extraction_Q_unitaire As Double
Dim extraction_pompe_rdt As Double
Public electricite As Double

Public P_precipite As Double
Public electricite_fixe As Double

'variables montrant si les valeurs ont été forcées
Dim bypass_force As Boolean
Dim dosage_FeCl3_force As Boolean
Dim dosage_polymere_force As Boolean
Dim mn_NombreOuvrages_force As Boolean
Dim V_mirroir_limite_force As Boolean
Dim S_mirroir_force As Boolean
Dim hauteur_force As Boolean
Dim rdt_MES_force As Boolean
Dim rdt_DCO_force As Boolean
Dim rdt_DBO_force As Boolean
Dim rdt_NK_force As Boolean
Dim rdt_P_force As Boolean
Dim MV_MES_force As Boolean
Dim recirculation_taux_force As Boolean
Dim recirculation_nb_pompe_force As Boolean
Dim recirculation_P_refoulement_force As Boolean
Dim recirculation_tps_fct_pompe_force As Boolean
Dim recirculation_Q_unitaire_force As Boolean
Dim recirculation_pompe_rdt_force As Boolean
Dim boues_concentration_force As Boolean
Dim boues_MES_force As Boolean
Dim boues_Q_force As Boolean
Dim extraction_nb_pompe_force As Boolean
Dim extraction_P_refoulement_force As Boolean
Dim extraction_tps_fct_pompe_force As Boolean
Dim extraction_Q_unitaire_force As Boolean
Dim extraction_pompe_rdt_force As Boolean

'valeurs forcées par l'utilisateur
Dim bypass_f As Double
Dim dosage_FeCl3_f As Double
Dim dosage_polymere_f As Double
Dim mi_NombreOuvrages_f As Integer
Dim V_mirroir_limite_f As Double
Dim S_mirroir_f As Double
Dim hauteur_f As Double
Dim rdt_MES_f As Double
Dim rdt_DCO_f As Double
Dim rdt_DBO_f As Double
Dim rdt_NK_f As Double
Dim rdt_P_f As Double
Dim MV_MES_f As Double
Dim recirculation_taux_f As Double
Dim recirculation_nb_pompe_f As Integer
Dim recirculation_P_refoulement_f As Double
Dim recirculation_tps_fct_pompe_f As Double
Dim recirculation_Q_unitaire_f As Double
Dim recirculation_pompe_rdt_f As Double
Dim boues_concentration_f As Double
Dim boues_MES_f As Double
Dim boues_Q_f As Double
Dim extraction_nb_pompe_f As Integer
Dim extraction_P_refoulement_f As Double
Dim extraction_tps_fct_pompe_f As Double
Dim extraction_Q_unitaire_f As Double
Dim extraction_pompe_rdt_f As Double

Dim Q_traite_nominal As Double
Dim Q_traite_reel As Double
Dim MV_MES_reference As Double
Dim MES_formes As Double
Dim hyp_rdt_P_particulaire As Double
Dim ratio_Psol_P As Double
Dim ratio_FeCl3_Pt As Double
Dim ratio_PO4_Psol As Double
Dim MM_P As Double
Dim MM_FeO3H3 As Double
Dim MM_FePO4 As Double
Dim MM_FeCl3 As Double

Dim remarque(4) As String


Sub dimensionnement(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByVal Q_retour, ByVal eaux_sales, ByVal choix_eaux_sales, ByRef boues_I_detail, ByVal ES_discfilter)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'variables internes
    Dim I As Integer

    Dim DCO_traite As Double
    Dim DBO_traite As Double
    Dim MES_traite As Double
    Dim NK_traite As Double
    Dim NH4_traite As Double
    Dim NO3_traite As Double
    Dim Pt_traite As Double
    Dim Q_retour_traite As Double
    Dim Q_eaux_sales_traite As Double
    Dim FeCl3_flux_massique As Double
    Dim rdt_P_soluble As Double
    Dim P_outlet As Double
    Dim P_soluble_removed As Double
    Dim ratio_molaire_Ptbr As Double

    Dim volume_decanteur As Double

    Dim Q_eaux_sales As Double
    Const repere_MV_ES1 = 1
    Const repere_MV_ES2 = 2
    Const repere_MV_primaire = 3
    Dim intermediaire_MV(repere_MV_primaire) As Double
    Dim intermediaire_MES(repere_MV_primaire) As Double
    Dim MES_eaux_sales As Double
    Dim inlet_MV As Double

    'détermination du débit d'eaux sales au global
    Q_eaux_sales = 0
    MES_eaux_sales = 0
    inlet_MV = 0
    For I = 1 To nb_eaux_sales_max
        If choix_eaux_sales(I) = devenir_ES_primaire Then
            'débit d'eaux sales
            Q_eaux_sales = Q_eaux_sales + eaux_sales(I, repere_ES_Q)
            'MES des eaux sales (pour calcul du ratio MV/MES
            MES_eaux_sales = MES_eaux_sales + eaux_sales(I, repere_ES_MES)
            'report de l'origine des boues
            boues_I_detail(I, repere_origine) = eaux_sales(I, repere_ES_origine)
            'calcul des MV de chaque type d'eaux sales
            intermediaire_MV(I) = eaux_sales(I, repere_ES_MV_MES) * eaux_sales(I, repere_ES_MES)
            'calcul des MES de chaque type d'eaux sales
            intermediaire_MES(I) = eaux_sales(I, repere_ES_MES)
            inlet_MV = inlet_MV + intermediaire_MV(I)
        End If
    Next I
    'on ajoute les MV de l'eau brute (que l'on obtiendrait en théorie sans les eaux sales
    intermediaire_MV(repere_MV_primaire) = MV_MES_reference * (MES - MES_eaux_sales)
    inlet_MV = inlet_MV + intermediaire_MV(repere_MV_primaire)
    'on calcule la provenance des MV par type de boues (ratio par rapport aux MV totaux)
    For I = repere_MV_ES1 To repere_MV_ES2
        boues_I_detail(I, repere_ratio_MV) = intermediaire_MV(I) / inlet_MV
        boues_I_detail(I, repere_ratio_MES) = intermediaire_MES(I) / MES
    Next I
    'on calcule le ratio de MV/MES   (ici Q_traite_nominal/Q correspond au flux admis)
    If MV_MES_force = True Then
        MV_MES = MV_MES_f
    Else
        If iteration = 1 Then
            Q_traite_nominal = Q
        End If
        MV_MES = (inlet_MV * Q_traite_nominal / Q) / (MES * Q_traite_nominal / Q + MES_formes)
    End If




    'charges à considérer pour dimensionnement
    If choix_dimension_bypass = dimensionne_100 Then
        Q_traite_nominal = Q
        DCO_traite = DCO
        DBO_traite = DBO
        MES_traite = MES
        NK_traite = NK
        NH4_traite = NH4
        NO3_traite = NO3
        Pt_traite = Pt
        Q_retour_traite = Q_retour
        Q_eaux_sales_traite = Q_eaux_sales
    Else
        Q_traite_nominal = (1 - bypass) * Q
        DCO_traite = (1 - bypass) * DCO
        DBO_traite = (1 - bypass) * DBO
        MES_traite = (1 - bypass) * MES
        NK_traite = (1 - bypass) * NK
        NH4_traite = (1 - bypass) * NH4
        NO3_traite = (1 - bypass) * NO3
        Pt_traite = (1 - bypass) * Pt
        Q_retour_traite = (1 - bypass) * Q_retour
        Q_eaux_sales_traite = (1 - bypass) * Q_eaux_sales
    End If

    'calcul du rendement en Pt
    If Pt_traite / Q_traite_nominal * 1000 < 10 Then
        hyp_rdt_P_particulaire = 0.8
    Else
        hyp_rdt_P_particulaire = 0.85
    End If
    If rdt_P_force = True Then
        rdt_P = rdt_P_f
        If dosage_FeCl3_force = True Then
            P_soluble_removed = rdt_P * Pt_traite - ES_discfilter(repere_ES_Pt) - (1 - ratio_Psol_P) * (Pt_traite - ES_discfilter(repere_ES_Pt)) * hyp_rdt_P_particulaire
            If P_soluble_removed <= 0 Then
                P_soluble_removed = 0
                ratio_molaire_Ptbr = 1
            Else
                ratio_molaire_Ptbr = dosage_FeCl3 / (ratio_FeCl3_Pt * (P_soluble_removed / Q_traite_nominal * 1000))
            End If
        Else
            P_outlet = Pt_traite * (1 - rdt_P)
            P_soluble_removed = rdt_P * Pt_traite - ES_discfilter(repere_ES_Pt) - (1 - ratio_Psol_P) * (Pt_traite - ES_discfilter(repere_ES_Pt)) * hyp_rdt_P_particulaire
            If P_soluble_removed <= 0 Then
                P_soluble_removed = 0
                ratio_molaire_Ptbr = 1
            Else
                ratio_molaire_Ptbr = ratio_molaire_precipitation(P_soluble_removed, Q_traite_nominal, P_outlet)
            End If
            dosage_FeCl3 = ratio_molaire_Ptbr * ratio_FeCl3_Pt * (P_soluble_removed / Q_traite_nominal * 1000)
        End If
    Else
        rdt_P_soluble = coagulation_floculation_rdt_P_soluble(Q_traite_nominal, Pt_traite - ES_discfilter(repere_ES_Pt))
        rdt_P = ratio_Psol_P * ratio_PO4_Psol * rdt_P_soluble + (1 - ratio_Psol_P) * hyp_rdt_P_particulaire
        rdt_P = (rdt_P * (Pt_traite - ES_discfilter(repere_ES_Pt)) + (ES_discfilter(repere_ES_Pt))) / Pt_traite   '* Pt_traite / Pt
        P_soluble_removed = rdt_P_soluble * ratio_Psol_P * ratio_PO4_Psol * (Pt_traite - ES_discfilter(repere_ES_Pt))
        ratio_molaire_Ptbr = dosage_FeCl3 / (ratio_FeCl3_Pt * (P_soluble_removed / Q_traite_nominal * 1000))
    End If

    P_precipite = P_soluble_removed


    'poids du FeCl3 dans les MES
    FeCl3_flux_massique = dosage_FeCl3 * Q_traite_nominal / 1000
    'MES_formes = FeCl3_flux_massique * rate_MES_FeCl3
    MES_formes = (ratio_molaire_Ptbr - 1) * MM_FeO3H3 / MM_P + MM_FePO4 / MM_P
    MES_formes = MES_formes * P_soluble_removed

    'calcul de la surface au mirroir
    S_mirroir = (pd_CoefficientPointeHydrauliqueTempsPluie * (Q_traite_nominal - Q_retour_traite - Q_eaux_sales_traite) + Q_retour_traite + Q_eaux_sales_traite) / (pd_NOMBRE_HEURE_PAR_JOUR * V_mirroir_limite)
    If S_mirroir_force = True Then
        S_mirroir = S_mirroir_f
    End If

    'calcul de la vitesse au mirroir limite recalculée
    V_mirroir_limite_recalculee = (pd_CoefficientPointeHydrauliqueTempsPluie * (Q_traite_nominal - Q_retour_traite - Q_eaux_sales_traite) + Q_retour_traite + Q_eaux_sales_traite) / (pd_NOMBRE_HEURE_PAR_JOUR * S_mirroir)


    'calcul des boues à extraire
    boues_MES = MES_traite * rdt_MES + MES_formes
    boues_Q = boues_MES / boues_concentration


    'calcul des débit et charges en sortie eau
    Q = Q - boues_Q
    DCO = DCO - DCO_traite * rdt_DCO
    DBO = DBO - DBO_traite * rdt_DBO
    MES = MES - MES_traite * rdt_MES
    NK = NK - NK_traite * rdt_NK
    NH4 = NH4 * Q / (Q + boues_Q)
    NO3 = NO3 * Q / (Q + boues_Q)
    Pt = Pt - Pt_traite * rdt_P


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'attention, il faudra calculer la quantité de boues primaires et la quantité de boues bio si biostyr'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "10.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Function ratio_molaire_precipitation(ByVal deltaP_precipitation, ByVal Q, ByVal P_outlet) As Double

    Dim Ptbr As Double
    Dim P_sortie As Double

    P_sortie = P_outlet / Q * 1000
    Ptbr = deltaP_precipitation / Q * 1000

    ratio_molaire_precipitation = (Ptbr + P_sortie) ^ (-0.5) - (P_sortie ^ (-0.5) - 0.1)
    ratio_molaire_precipitation = ratio_molaire_precipitation / ((Ptbr + P_sortie) - (P_sortie - 0.1))
    ratio_molaire_precipitation = 1.1 - 6 * ratio_molaire_precipitation


End Function

Private Function coagulation_floculation_rdt_P_soluble(ByVal Q, ByVal Pt)

    Const precipitation_nb_iteration = 30

    Dim x_min As Double
    Dim x_max As Double
    Dim Psol_in As Double
    Dim Psol_out As Double
    Dim dose_FeCl3_dichotomie As Double
    Dim Ptbr As Double
    Dim ratio_molaire_Ptbr As Double

    If dosage_FeCl3 > 0 Then
        Psol_in = Pt / Q * 1000 * ratio_Psol_P * ratio_PO4_Psol
        x_min = 0.1
        x_max = Psol_in
        Ptbr = (x_min + x_max) / 2
        Do While (x_max - x_min) / Ptbr > 0.00001
            Psol_out = Psol_in - Ptbr
            ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + Psol_out) ^ (-0.5) - (Psol_out ^ (-0.5) - 0.1)) / ((Ptbr + Psol_out) - (Psol_out - 0.1))
            dose_FeCl3_dichotomie = Ptbr * ratio_molaire_Ptbr * ratio_FeCl3_Pt
            If dose_FeCl3_dichotomie < dosage_FeCl3 Then
                x_min = Ptbr
            Else
                x_max = Ptbr
            End If
            Ptbr = (x_min + x_max) / 2
        Loop
        coagulation_floculation_rdt_P_soluble = Ptbr / Psol_in
    Else
        coagulation_floculation_rdt_P_soluble = 0
    End If

End Function


Sub fonctionnement_reel(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByVal Q_retour, ByVal eaux_sales, ByVal choix_eaux_sales, ByRef boues_I_detail, ByVal ES_discfilter)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'variables internes
    Dim I As Integer
    Dim DCO_traite As Double
    Dim DBO_traite As Double
    Dim MES_traite As Double
    Dim NK_traite As Double
    Dim NH4_traite As Double
    Dim NO3_traite As Double
    Dim Pt_traite As Double
    Dim volume_decanteur As Double
    Dim FeCl3_flux_massique As Double
    Dim polymere_flux_massique As Double
    Dim rdt_P_soluble As Double
    Dim P_outlet As Double
    Dim P_soluble_removed As Double
    Dim ratio_molaire_Ptbr As Double

    'erreurs
    Dim incoherence As Double


    Dim Q_eaux_sales As Double
    Const repere_MV_ES1 = 1
    Const repere_MV_ES2 = 2
    Const repere_MV_primaire = 3
    Dim intermediaire_MV(repere_MV_primaire) As Double
    Dim intermediaire_MES(repere_MV_primaire) As Double
    Dim MES_eaux_sales As Double
    Dim inlet_MV As Double


    'détermination du débit d'eaux sales au global
    Q_eaux_sales = 0
    MES_eaux_sales = 0
    inlet_MV = 0
    For I = 1 To nb_eaux_sales_max
        If choix_eaux_sales(I) = devenir_ES_primaire Then
            'débit d'eaux sales
            Q_eaux_sales = Q_eaux_sales + eaux_sales(I, repere_ES_Q)
            'MES des eaux sales (pour calcul du ratio MV/MES
            MES_eaux_sales = MES_eaux_sales + eaux_sales(I, repere_ES_MES)
            'report de l'origine des boues
            boues_I_detail(I, repere_origine) = eaux_sales(I, repere_ES_origine)
            'calcul des MV de chaque type d'eaux sales
            intermediaire_MV(I) = eaux_sales(I, repere_ES_MV_MES) * eaux_sales(I, repere_ES_MES)
            'calcul des MES de chaque type d'eaux sales
            intermediaire_MES(I) = eaux_sales(I, repere_ES_MES)
            inlet_MV = inlet_MV + intermediaire_MV(I)
        End If
    Next I
    'on ajoute les MV de l'eau brute (que l'on obtiendrait en théorie sans les eaux sales
    intermediaire_MV(repere_MV_primaire) = MV_MES_reference * (MES - MES_eaux_sales)
    inlet_MV = inlet_MV + intermediaire_MV(repere_MV_primaire)
    'on calcule la provenance des MV par type de boues (ratio par rapport aux MV totaux)
    For I = repere_MV_ES1 To repere_MV_ES2
        boues_I_detail(I, repere_ratio_MV) = intermediaire_MV(I) / inlet_MV
        boues_I_detail(I, repere_ratio_MES) = intermediaire_MES(I) / MES
    Next I
    'on calcule le ratio de MV/MES   (ici Q_traite_nominal/Q correspond au flux admis)
    If MV_MES_force = True Then
        MV_MES = MV_MES_f
    Else
        If iteration = 1 Then
            Q_traite_reel = Q
        End If
        MV_MES = (inlet_MV * Q_traite_reel / Q) / (MES * Q_traite_reel / Q + MES_formes)
    End If




    'charges traitées
    Q_traite_reel = (1 - bypass) * Q
    DCO_traite = (1 - bypass) * DCO
    DBO_traite = (1 - bypass) * DBO
    MES_traite = (1 - bypass) * MES
    NK_traite = (1 - bypass) * NK
    NH4_traite = (1 - bypass) * NH4
    NO3_traite = (1 - bypass) * NO3
    Pt_traite = (1 - bypass) * Pt

    'calcul du rendement en Pt
    If Pt_traite / Q_traite_reel * 1000 < 10 Then
        hyp_rdt_P_particulaire = 0.8
    Else
        hyp_rdt_P_particulaire = 0.85
    End If
    If rdt_P_force = True Then
        rdt_P = rdt_P_f
        If dosage_FeCl3_force = True Then
            P_soluble_removed = rdt_P * Pt_traite - ES_discfilter(repere_ES_Pt) - (1 - ratio_Psol_P) * (Pt_traite - ES_discfilter(repere_ES_Pt)) * hyp_rdt_P_particulaire
            If P_soluble_removed <= 0 Then
                P_soluble_removed = 0
                ratio_molaire_Ptbr = 1
            Else
                ratio_molaire_Ptbr = dosage_FeCl3 / (ratio_FeCl3_Pt * (P_soluble_removed / Q_traite_reel * 1000))
            End If
        Else
            P_outlet = Pt_traite * (1 - rdt_P)
            P_soluble_removed = rdt_P * Pt_traite - ES_discfilter(repere_ES_Pt) - (1 - ratio_Psol_P) * (Pt_traite - ES_discfilter(repere_ES_Pt)) * hyp_rdt_P_particulaire
            If P_soluble_removed <= 0 Then
                P_soluble_removed = 0
                ratio_molaire_Ptbr = 1
            Else
                ratio_molaire_Ptbr = ratio_molaire_precipitation(P_soluble_removed, Q_traite_reel, P_outlet)
            End If
            dosage_FeCl3 = ratio_molaire_Ptbr * ratio_FeCl3_Pt * (P_soluble_removed / Q_traite_reel * 1000)
        End If
    Else
        rdt_P_soluble = coagulation_floculation_rdt_P_soluble(Q_traite_reel, Pt_traite - ES_discfilter(repere_ES_Pt))
        rdt_P = ratio_Psol_P * ratio_PO4_Psol * rdt_P_soluble + (1 - ratio_Psol_P) * hyp_rdt_P_particulaire
        rdt_P = (rdt_P * (Pt_traite - ES_discfilter(repere_ES_Pt)) + (ES_discfilter(repere_ES_Pt))) / Pt_traite   '* Pt_traite / Pt
        P_soluble_removed = rdt_P_soluble * ratio_Psol_P * ratio_PO4_Psol * (Pt_traite - ES_discfilter(repere_ES_Pt))
        ratio_molaire_Ptbr = dosage_FeCl3 / (ratio_FeCl3_Pt * (P_soluble_removed / Q_traite_reel * 1000))
    End If

    P_precipite = P_soluble_removed

    remarque(4) = Feuil10.Range("dec_reactif_ratio_molaire") & CStr(Int(100 * ratio_molaire_Ptbr) / 100)

    'poids du FeCl3 dans les MES
    FeCl3_flux_massique = dosage_FeCl3 * Q_traite_reel / 1000
    'MES_formes = FeCl3_flux_massique * rate_MES_FeCl3
    MES_formes = (ratio_molaire_Ptbr - 1) * MM_FeO3H3 / MM_P + MM_FePO4 / MM_P
    MES_formes = MES_formes * P_soluble_removed

    'polymère ajouté
    polymere_flux_massique = dosage_polymere * Q_traite_reel / 1000

    'calcul des paramètres de fonctionnement
    V_mirroir_moyenne = Q_traite_reel / (pd_NOMBRE_HEURE_PAR_JOUR * S_mirroir)
    volume_decanteur = S_mirroir * hauteur
    temps_retention = volume_decanteur / (Q_traite_reel / pd_NOMBRE_HEURE_PAR_JOUR)


    'calcul des boues à extraire
    boues_MES = MES_traite * rdt_MES + MES_formes
    If boues_MES_force = True Then
        boues_MES = boues_MES_f
    End If
    boues_Q = boues_MES / boues_concentration
    If boues_Q_force = True Then
        boues_Q = boues_Q_f
    End If

    'calcul des débit et charges en sortie eau
    Q = Q - boues_Q
    DCO = DCO - DCO_traite * rdt_DCO
    DBO = DBO - DBO_traite * rdt_DBO
    MES = MES - MES_traite * rdt_MES
    NK = NK - NK_traite * rdt_NK
    NH4 = NH4 * Q / (Q + boues_Q)
    NO3 = NO3 * Q / (Q + boues_Q)
    Pt = Pt - Pt_traite * rdt_P

    'incoherence = Abs(boues_Q * boues_concentration - boues_MES) / boues_MES
    'If incoherence > critere_incoherence Then
    'ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil9.Range("dec_reactif_erreur_boues_extraites").Cells(1, 1) + vbCrLf
    'End If


    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'attention, il faudra calculer la quantité de boues primaires et la quantité de boues bio si biostyr'
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "10.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub calcul_consommation_electrique()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim rate_elec_agitation_coagulation As Double
    Dim rate_elec_agitation_floculation As Double
    Dim rate_elec_polymere_preparation As Double
    Dim rate_elec_polymere_pompage As Double
    Dim rate_elec_FeCl3 As Double
    Dim rate_elec_microsable As Double
    Dim rate_puissance_fixe As Double
    Dim rate_puissance_debit As Double
    Dim rate_elec_autre As Double

    Dim rendement_pompe As Double
    Dim puissance_pompe As Double
    Dim rendement_moteur As Double
    Dim dec_diametre_limite As Double

    Dim electricite_agitation As Double
    Dim electricite_racleur As Double
    Dim electricite_reactifs As Double
    Dim electricite_recirculation As Double
    Dim electricite_extraction As Double
    Dim electricite_autre As Double

    'erreurs
    Dim incoherence As Double


    'AGITATION CUVES COAGULATION-FLOCULATION
    rate_elec_agitation_coagulation = Feuil48.Range("hyp_dec_I_conso_elec_coagulation_agitation").Cells(1, 1)    'Wh/m3  moyenne(1.08;1.77)
    rate_elec_agitation_floculation = Feuil48.Range("hyp_dec_I_conso_elec_floculation_agitation").Cells(1, 1)    'Wh/m3  vérifier sources
    electricite_agitation = (rate_elec_agitation_coagulation + rate_elec_agitation_floculation) * Q_traite_nominal / 1000

    'RACLEUR
    'rate_puissance_fixe = 0.1 'Wh/m3 par rapport au nominal
    'rate_puissance_debit = 1 'Wh/m3 par rapport au débit réel
    'Nantes: racleur/herse: 0.38 Wh/m3
    'Angers racleur/herse: 0.88 Wh/m3
    'electricite_racleur = 1 / 1000 * (rate_puissance_fixe * Q_traite_nominal + rate_puissance_debit * Q_traite_reel)
    'DCA le 15/07/11 données Pascal JOUAFFRE  de 12 à 32m diamètre --> 0.55 kW installés, de 33 à 52m --> 0.75 kW
    dec_diametre_limite = Feuil48.Range("hyp_dec_I_diametre_limite").Cells(1, 1)
    If S_mirroir / mi_NombreOuvrages < pd_NOMBRE_PI * (dec_diametre_limite / 2) ^ 2 Then
        electricite_racleur = mi_NombreOuvrages * Feuil48.Range("hyp_dec_I_Pw_racleur_inf_Dlim").Cells(1, 1) * pd_NOMBRE_HEURE_PAR_JOUR
    Else
        electricite_racleur = mi_NombreOuvrages * Feuil48.Range("hyp_dec_I_Pw_racleur_sup_Dlim").Cells(1, 1) * pd_NOMBRE_HEURE_PAR_JOUR
    End If


    'INJECTION DE REACTIF, POLYMERE ET MICROSABLE
    rate_elec_polymere_preparation = Feuil48.Range("hyp_dec_I_elec_prepa_polymere").Cells(1, 1)    'Wh/m3 revoir les sources
    rate_elec_polymere_pompage = Feuil48.Range("hyp_dec_I_elec_pompage_polymere").Cells(1, 1)    'Wh/m3 revoir les sources
    rate_elec_FeCl3 = Feuil48.Range("hyp_dec_I_elec_injection_FeCl3").Cells(1, 1)
    If choix_microsable = microsable_oui Then
        rate_elec_microsable = Feuil48.Range("hyp_dec_I_elec_microsable").Cells(1, 1)
    Else
        rate_elec_microsable = 0
    End If
    electricite_reactifs = (rate_elec_polymere_preparation + rate_elec_polymere_pompage + rate_elec_FeCl3 + rate_elec_microsable) * Q_traite_reel / 1000

    'RECIRCULATION DES BOUES    ici, le rendement_moteur correspond au rendement machine (mécanique)
    If choix_recirculation = recirculation_oui Then
        If recirculation_Q_unitaire_force = True Then
            recirculation_Q_unitaire = recirculation_Q_unitaire_f
            If recirculation_tps_fct_pompe_force = True Then
                'on ne fait rien
                If recirculation_nb_pompe_force = False Then
                    If Int((recirculation_taux * Q_traite_reel) / (recirculation_tps_fct_pompe * recirculation_Q_unitaire)) < ((recirculation_taux * Q_traite_reel) / (recirculation_tps_fct_pompe * recirculation_Q_unitaire)) Then
                        recirculation_nb_pompe = Int((recirculation_taux * Q_traite_reel) / (recirculation_tps_fct_pompe * recirculation_Q_unitaire)) + 1
                    Else
                        recirculation_nb_pompe = Int((recirculation_taux * Q_traite_reel) / (recirculation_tps_fct_pompe * recirculation_Q_unitaire))
                    End If
                End If
            Else
                recirculation_tps_fct_pompe = (recirculation_taux * Q_traite_reel) / (recirculation_nb_pompe * recirculation_Q_unitaire)
                If recirculation_tps_fct_pompe > pd_NOMBRE_HEURE_PAR_JOUR Then
                    recirculation_tps_fct_pompe = pd_NOMBRE_HEURE_PAR_JOUR
                    If recirculation_nb_pompe_force = False Then
                        If Int((recirculation_taux * Q_traite_reel) / (recirculation_tps_fct_pompe * recirculation_Q_unitaire)) < ((recirculation_taux * Q_traite_reel) / (recirculation_tps_fct_pompe * recirculation_Q_unitaire)) Then
                            recirculation_nb_pompe = Int((recirculation_taux * Q_traite_reel) / (recirculation_tps_fct_pompe * recirculation_Q_unitaire)) + 1
                        Else
                            recirculation_nb_pompe = Int((recirculation_taux * Q_traite_reel) / (recirculation_tps_fct_pompe * recirculation_Q_unitaire))
                        End If
                        recirculation_tps_fct_pompe = (recirculation_taux * Q_traite_reel) / (recirculation_nb_pompe * recirculation_Q_unitaire)
                    End If
                End If
            End If
        Else
            recirculation_Q_unitaire = (recirculation_taux * Q_traite_reel) / (recirculation_nb_pompe * recirculation_tps_fct_pompe)
        End If
        incoherence = (recirculation_Q_unitaire * recirculation_nb_pompe * recirculation_tps_fct_pompe - (recirculation_taux * Q_traite_reel)) / (recirculation_taux * Q_traite_reel)
        If Abs(incoherence) > critere_incoherence Then
            remarque(1) = Feuil10.Range("dec_reactif_erreur_pompe_recirculation").Cells(1, 1)
        End If
        If recirculation_pompe_rdt_force = True Then
            recirculation_pompe_rdt = recirculation_pompe_rdt_f
        Else
            rendement_pompe = 0.53 * 0.92    'hypothèse
            puissance_pompe = recirculation_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * recirculation_P_refoulement / rendement_pompe
            If puissance_pompe < 60 Then
                rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
            Else
                rendement_moteur = 0.961
            End If
            recirculation_pompe_rdt = rendement_pompe * rendement_moteur
        End If
        electricite_recirculation = recirculation_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * recirculation_P_refoulement / recirculation_pompe_rdt * recirculation_nb_pompe * recirculation_tps_fct_pompe
    End If

    'EXTRACTION DES BOUES    ici, le rendement_moteur correspond au rendement machine (mécanique)
    If extraction_Q_unitaire_force = True Then
        extraction_Q_unitaire = extraction_Q_unitaire_f
        If extraction_tps_fct_pompe_force = True Then
            'on ne fait rien
            If extraction_nb_pompe_force = False Then
                If Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) < (boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) Then
                    extraction_nb_pompe = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) + 1
                Else
                    extraction_nb_pompe = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire))
                End If
            End If
        Else
            extraction_tps_fct_pompe = boues_Q / (extraction_nb_pompe * extraction_Q_unitaire)
            If extraction_tps_fct_pompe > pd_NOMBRE_HEURE_PAR_JOUR Then
                extraction_tps_fct_pompe = pd_NOMBRE_HEURE_PAR_JOUR
                If extraction_nb_pompe_force = False Then
                    If Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) < (boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) Then
                        extraction_nb_pompe = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) + 1
                    Else
                        extraction_nb_pompe = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire))
                    End If
                    extraction_tps_fct_pompe = boues_Q / (extraction_nb_pompe * extraction_Q_unitaire)
                End If
            End If
        End If
    Else
        extraction_Q_unitaire = boues_Q / (extraction_nb_pompe * extraction_tps_fct_pompe)
    End If
    incoherence = (extraction_Q_unitaire * extraction_nb_pompe * extraction_tps_fct_pompe - boues_Q) / boues_Q
    If Abs(incoherence) > critere_incoherence Then
        remarque(2) = Feuil10.Range("dec_reactif_erreur_pompe_extraction").Cells(1, 1)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt = extraction_pompe_rdt_f
    Else
        rendement_pompe = 0.7    'hypothèse
        puissance_pompe = extraction_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * extraction_P_refoulement / rendement_pompe
        If puissance_pompe < 60 Then
            rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
        Else
            rendement_moteur = 0.961
        End If
        extraction_pompe_rdt = rendement_pompe * rendement_moteur
    End If
    electricite_extraction = extraction_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * extraction_P_refoulement / extraction_pompe_rdt * extraction_nb_pompe * extraction_tps_fct_pompe

    'POMPAGE AUTRES (reprise vidanges partielles, horizontale à écumes)
    rate_elec_autre = Feuil48.Range("hyp_dec_I_elec_divers").Cells(1, 1)
    electricite_autre = rate_elec_autre * Q_traite_reel / 1000

    'CONSOMMATION GLOBALE D'ELECTRICITE
    electricite = electricite_racleur + electricite_recirculation + electricite_extraction + electricite_agitation + electricite_reactifs + electricite_autre
    electricite_fixe = electricite_racleur + electricite_agitation


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "10.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'lecture des choix internes

    choix_dimension_bypass = Feuil10.Range("dec_reactif_choix_bypass").Cells(1, pi_FiliereConsideree)
    If choix_dimension_bypass = dimensionne_bypass Then
        bypass_force = True
        bypass_f = Feuil10.Range("dec_reactif_bypass").Cells(1, pi_FiliereConsideree)
    End If

    choix_lamellaire = Feuil10.Range("dec_reactif_choix_lamelle").Cells(1, pi_FiliereConsideree)
    choix_polymere_type = Feuil10.Range("dec_reactif_choix_polymere_type").Cells(1, pi_FiliereConsideree)
    choix_microsable = Feuil10.Range("dec_reactif_choix_microsable").Cells(1, pi_FiliereConsideree)
    choix_recirculation = Feuil10.Range("dec_reactif_choix_recirculation").Cells(1, pi_FiliereConsideree)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "10.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'lecture du choix de forcer des valeurs
    bypass_force = Feuil10.Range("dec_reactif_bypass").Cells(1, pi_FiliereConsideree)
    dosage_FeCl3_force = Feuil10.Range("dec_reactif_dosage_FeCl3").Cells(1, pi_FiliereConsideree)
    dosage_polymere_force = Feuil10.Range("dec_reactif_dosage_polymere").Cells(1, pi_FiliereConsideree)
    mn_NombreOuvrages_force = Feuil10.Range("dec_reactif_NombreOuvrages").Cells(1, pi_FiliereConsideree)
    V_mirroir_limite_force = Feuil10.Range("dec_reactif_Vmirroir_limite").Cells(1, pi_FiliereConsideree)
    S_mirroir_force = Feuil10.Range("dec_reactif_Smirroir").Cells(1, pi_FiliereConsideree)
    hauteur_force = Feuil10.Range("dec_reactif_hauteur").Cells(1, pi_FiliereConsideree)
    rdt_MES_force = Feuil10.Range("dec_reactif_rdt_MES").Cells(1, pi_FiliereConsideree)
    rdt_DCO_force = Feuil10.Range("dec_reactif_rdt_DCO").Cells(1, pi_FiliereConsideree)
    rdt_DBO_force = Feuil10.Range("dec_reactif_rdt_DBO").Cells(1, pi_FiliereConsideree)
    rdt_NK_force = Feuil10.Range("dec_reactif_rdt_NK").Cells(1, pi_FiliereConsideree)
    rdt_P_force = Feuil10.Range("dec_reactif_rdt_P").Cells(1, pi_FiliereConsideree)
    MV_MES_force = Feuil10.Range("dec_reactif_MV_MES").Cells(1, pi_FiliereConsideree)
    If choix_recirculation = recirculation_oui Then
        recirculation_taux_force = Feuil10.Range("dec_reactif_recirculation_taux").Cells(1, pi_FiliereConsideree)
        recirculation_nb_pompe_force = Feuil10.Range("dec_reactif_recirculation_nb_pompe").Cells(1, pi_FiliereConsideree)
        recirculation_P_refoulement_force = Feuil10.Range("dec_reactif_recirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
        recirculation_tps_fct_pompe_force = Feuil10.Range("dec_reactif_recirculation_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
        recirculation_Q_unitaire_force = Feuil10.Range("dec_reactif_recirculation_Qunitaire").Cells(1, pi_FiliereConsideree)
        recirculation_pompe_rdt_force = Feuil10.Range("dec_reactif_recirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    boues_concentration_force = Feuil10.Range("dec_reactif_boues_concentration").Cells(1, pi_FiliereConsideree)
    boues_MES_force = Feuil10.Range("dec_reactif_boues_MES").Cells(1, pi_FiliereConsideree)
    boues_Q_force = False
    extraction_nb_pompe_force = Feuil10.Range("dec_reactif_extraction_nb_pompe").Cells(1, pi_FiliereConsideree)
    extraction_P_refoulement_force = Feuil10.Range("dec_reactif_extraction_Prefoulement").Cells(1, pi_FiliereConsideree)
    extraction_tps_fct_pompe_force = Feuil10.Range("dec_reactif_extraction_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    extraction_Q_unitaire_force = Feuil10.Range("dec_reactif_extraction_Qunitaire").Cells(1, pi_FiliereConsideree)
    extraction_pompe_rdt_force = Feuil10.Range("dec_reactif_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)

    'lecture des valeurs forcées
    If bypass_force = True Then
        bypass_f = Feuil10.Range("dec_reactif_bypass").Cells(1, pi_FiliereConsideree)
    End If
    If dosage_FeCl3_force = True Then
        dosage_FeCl3_f = Feuil10.Range("dec_reactif_dosage_FeCl3").Cells(1, pi_FiliereConsideree)
    End If
    If dosage_polymere_force = True Then
        dosage_polymere_f = Feuil10.Range("dec_reactif_dosage_polymere").Cells(1, pi_FiliereConsideree)
    End If
    If mn_NombreOuvrages_force = True Then
        mi_NombreOuvrages_f = Feuil10.Range("dec_reactif_NombreOuvrages").Cells(1, pi_FiliereConsideree)
    End If
    If V_mirroir_limite_force = True Then
        V_mirroir_limite_f = Feuil10.Range("dec_reactif_Vmirroir_limite").Cells(1, pi_FiliereConsideree)
    End If
    If S_mirroir_force = True Then
        S_mirroir_f = Feuil10.Range("dec_reactif_Smirroir").Cells(1, pi_FiliereConsideree)
    End If
    If hauteur_force = True Then
        hauteur_f = Feuil10.Range("dec_reactif_hauteur").Cells(1, pi_FiliereConsideree)
    End If
    If rdt_MES_force = True Then
        rdt_MES_f = Feuil10.Range("dec_reactif_rdt_MES").Cells(1, pi_FiliereConsideree)
    End If
    If rdt_DCO_force = True Then
        rdt_DCO_f = Feuil10.Range("dec_reactif_rdt_DCO").Cells(1, pi_FiliereConsideree)
    End If
    If rdt_DBO_force = True Then
        rdt_DBO_f = Feuil10.Range("dec_reactif_rdt_DBO").Cells(1, pi_FiliereConsideree)
    End If
    If rdt_NK_force = True Then
        rdt_NK_f = Feuil10.Range("dec_reactif_rdt_NK").Cells(1, pi_FiliereConsideree)
    End If
    If rdt_P_force = True Then
        rdt_P_f = Feuil10.Range("dec_reactif_rdt_P").Cells(1, pi_FiliereConsideree)
    End If
    If MV_MES_force = True Then
        MV_MES_f = Feuil10.Range("dec_reactif_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If choix_recirculation = recirculation_oui Then
        If recirculation_taux_force = True Then
            recirculation_taux_f = Feuil10.Range("dec_reactif_recirculation_taux").Cells(1, pi_FiliereConsideree)
        End If
        If recirculation_nb_pompe_force = True Then
            recirculation_nb_pompe_f = Feuil10.Range("dec_reactif_recirculation_nb_pompe").Cells(1, pi_FiliereConsideree)
        End If
        If recirculation_P_refoulement_force = True Then
            recirculation_P_refoulement_f = Feuil10.Range("dec_reactif_recirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
        End If
        If recirculation_tps_fct_pompe_force = True Then
            recirculation_tps_fct_pompe_f = Feuil10.Range("dec_reactif_recirculation_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
        End If
        If recirculation_Q_unitaire_force = True Then
            recirculation_Q_unitaire_f = Feuil10.Range("dec_reactif_recirculation_Qunitaire").Cells(1, pi_FiliereConsideree)
        End If
        If recirculation_pompe_rdt_force = True Then
            recirculation_pompe_rdt_f = Feuil10.Range("dec_reactif_recirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
        End If
    End If
    If boues_concentration_force = True Then
        boues_concentration_f = Feuil10.Range("dec_reactif_boues_concentration").Cells(1, pi_FiliereConsideree)
    End If
    If boues_MES_force = True Then
        boues_MES_f = Feuil10.Range("dec_reactif_boues_MES").Cells(1, pi_FiliereConsideree)
    End If
    If boues_Q_force = True Then
        boues_Q_f = Feuil10.Range("dec_reactif_boues_Q").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_nb_pompe_force = True Then
        extraction_nb_pompe_f = Feuil10.Range("dec_reactif_extraction_nb_pompe").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement_f = Feuil10.Range("dec_reactif_extraction_Prefoulement").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_tps_fct_pompe_force = True Then
        extraction_tps_fct_pompe_f = Feuil10.Range("dec_reactif_extraction_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_Q_unitaire_force = True Then
        extraction_Q_unitaire_f = Feuil10.Range("dec_reactif_extraction_Qunitaire").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt_f = Feuil10.Range("dec_reactif_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "10.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'Voici l'ensemble des hypothèses sur le process
    ratio_Psol_P = Feuil48.Range("hyp_dec_I_ratio_Psol_P").Cells(1, 1)
    ratio_PO4_Psol = Feuil48.Range("hyp_dec_I_ratio_PO4_Psol").Cells(1, 1)
    ratio_FeCl3_Pt = 162.5 / 31

    'rendement P particulaire
    hyp_rdt_P_particulaire = 0.85

    'composition en MV des boues 1aire (sans ajout de réactifs et sans eaux sales biostyr)
    MV_MES_reference = Feuil48.Range("hyp_dec_I_boues_MV_MES").Cells(1, 1)

    'masses molaires utiles pour anaérobie  en g/mol
    MM_P = 31
    'MM_O = 16
    'MM_H = 1
    'MM_Fe = 55.85
    MM_FeO3H3 = 55.85 + 3 * (16 + 1)    'FE(OH)3
    MM_FePO4 = 55.85 + 31 + 4 * 16
    MM_FeCl3 = 162.5

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "10.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On retrouve ici l'ensemble des valeurs pour lesquelles on attribue une valeur guide par défaut

    'by-pass
    If bypass_force = True Then
        bypass = bypass_f
    Else
        bypass = 0
    End If

    'dosage de chlorrure ferrique
    If dosage_FeCl3_force = True Then
        dosage_FeCl3 = dosage_FeCl3_f
    Else
        dosage_FeCl3 = 30  'mg/L de produit pur
    End If

    'dosage de chlorrure ferrique
    If dosage_polymere_force = True Then
        dosage_polymere = dosage_polymere_f
    Else
        dosage_polymere = 0.7  ' g/L de produit pur
    End If

    'Nombre d'ouvrages
    If mn_NombreOuvrages_force = True Then
        mi_NombreOuvrages = mi_NombreOuvrages_f
    Else
        mi_NombreOuvrages = 1
    End If
    
    'vitesse au mirroir limite
    If V_mirroir_limite_force = True Then
        V_mirroir_limite = V_mirroir_limite_f
    Else
        If choix_biostyr = True Then
            If choix_microsable = microsable_oui Then
                V_mirroir_limite = 80  'm/h Aquademy et CAM
            Else
                V_mirroir_limite = 30  'm/h  Aquademy et CAM
            End If
        Else
            If choix_microsable = microsable_oui Then
                V_mirroir_limite = 120  'm/h Aquademy et CAM
            Else
                V_mirroir_limite = 40   'm/h Aquademy et CAM
            End If
        End If
    End If

    'hauteur du décanteur
    If hauteur_force = True Then
        hauteur = hauteur_f
    Else
        hauteur = 4
    End If

    'rendements d'élimination et concentration boues en sortie
    If choix_microsable = microsable_oui Then
        If rdt_MES_force = True Then
            rdt_MES = rdt_MES_f
        Else
            rdt_MES = 0.87
        End If
        If rdt_DCO_force = True Then
            rdt_DCO = rdt_DCO_f
        Else
            rdt_DCO = 0.6
        End If
        If rdt_DBO_force = True Then
            rdt_DBO = rdt_DBO_f
        Else
            rdt_DBO = 0.57
        End If
        If boues_concentration_force = True Then
            boues_concentration = boues_concentration_f
        Else
            boues_concentration = 30
        End If
    Else
        If rdt_MES_force = True Then
            rdt_MES = rdt_MES_f
        Else
            rdt_MES = 0.83
        End If
        If rdt_DCO_force = True Then
            rdt_DCO = rdt_DCO_f
        Else
            rdt_DCO = 0.55
        End If
        If rdt_DBO_force = True Then
            rdt_DBO = rdt_DBO_f
        Else
            rdt_DBO = 0.53
        End If
        If boues_concentration_force = True Then
            boues_concentration = boues_concentration_f
        Else
            boues_concentration = 30
        End If
    End If
    If rdt_NK_force = True Then
        rdt_NK = rdt_NK_f
    Else
        rdt_NK = 0.2
    End If
    If rdt_P_force = True Then
        rdt_P = rdt_P_f
    Else
        rdt_P = 0.4
    End If

    If MV_MES_force = True Then
        MV_MES = MV_MES_f
    Else
        MV_MES = MV_MES_reference
    End If

    'recirculation des boues
    If choix_recirculation = recirculation_oui Then
        If recirculation_taux_force = True Then
            recirculation_taux = recirculation_taux_f
        Else
            recirculation_taux = 0.06
        End If
        If recirculation_nb_pompe_force = True Then
            recirculation_nb_pompe = recirculation_nb_pompe_f
        Else
            recirculation_nb_pompe = 2
        End If
        If recirculation_P_refoulement_force = True Then
            recirculation_P_refoulement = recirculation_P_refoulement_f
        Else
            recirculation_P_refoulement = 25
        End If
        If recirculation_tps_fct_pompe_force = True Then
            recirculation_tps_fct_pompe = recirculation_tps_fct_pompe_f
        Else
            recirculation_tps_fct_pompe = pd_NOMBRE_HEURE_PAR_JOUR
        End If
    End If

    'extraction des boues
    If extraction_nb_pompe_force = True Then
        extraction_nb_pompe = extraction_nb_pompe_f
    Else
        extraction_nb_pompe = 2
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement = extraction_P_refoulement_f
    Else
        extraction_P_refoulement = 15
    End If
    If extraction_tps_fct_pompe_force = True Then
        extraction_tps_fct_pompe = extraction_tps_fct_pompe_f
    Else
        extraction_tps_fct_pompe = pd_NOMBRE_HEURE_PAR_JOUR
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "10.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil10.Range("dec_reactif_bypass").Cells(1, Colonne) = bypass
    Feuil10.Range("dec_reactif_dosage_FeCl3").Cells(1, Colonne) = dosage_FeCl3
    Feuil10.Range("dec_reactif_dosage_polymere").Cells(1, Colonne) = dosage_polymere
    Feuil10.Range("dec_reactif_NombreOuvrages").Cells(1, Colonne) = mi_NombreOuvrages
    Feuil10.Range("dec_reactif_Vmirroir_limite").Cells(1, Colonne) = V_mirroir_limite
    Feuil10.Range("dec_reactif_Smirroir").Cells(1, Colonne) = S_mirroir
    Feuil10.Range("dec_reactif_Vmirroir_limite_recalculee").Cells(1, Colonne) = V_mirroir_limite_recalculee
    Feuil10.Range("dec_reactif_hauteur").Cells(1, Colonne) = hauteur
    Feuil10.Range("dec_reactif_Vmirroir_moyenne").Cells(1, Colonne) = V_mirroir_moyenne
    Feuil10.Range("dec_reactif_tps_retention").Cells(1, Colonne) = temps_retention
    Feuil10.Range("dec_reactif_rdt_MES").Cells(1, Colonne) = rdt_MES
    Feuil10.Range("dec_reactif_rdt_DCO").Cells(1, Colonne) = rdt_DCO
    Feuil10.Range("dec_reactif_rdt_DBO").Cells(1, Colonne) = rdt_DBO
    Feuil10.Range("dec_reactif_rdt_NK").Cells(1, Colonne) = rdt_NK
    Feuil10.Range("dec_reactif_rdt_P").Cells(1, Colonne) = rdt_P
    Feuil10.Range("dec_reactif_MV_MES").Cells(1, Colonne) = MV_MES
    If choix_recirculation = recirculation_oui Then
        Feuil10.Range("dec_reactif_recirculation_taux").Cells(1, Colonne) = recirculation_taux
        Feuil10.Range("dec_reactif_recirculation_nb_pompe").Cells(1, Colonne) = recirculation_nb_pompe
        Feuil10.Range("dec_reactif_recirculation_P_refoulement").Cells(1, Colonne) = recirculation_P_refoulement
        Feuil10.Range("dec_reactif_recirculation_tps_fonctionnement").Cells(1, Colonne) = recirculation_tps_fct_pompe
        Feuil10.Range("dec_reactif_recirculation_Qunitaire").Cells(1, Colonne) = recirculation_Q_unitaire
        Feuil10.Range("dec_reactif_recirculation_pompe_rdt").Cells(1, Colonne) = recirculation_pompe_rdt
    End If
    Feuil10.Range("dec_reactif_boues_concentration").Cells(1, Colonne) = boues_concentration
    Feuil10.Range("dec_reactif_boues_MES").Cells(1, Colonne) = boues_MES
    Feuil10.Range("dec_reactif_boues_Q").Cells(1, Colonne) = boues_Q
    Feuil10.Range("dec_reactif_extraction_nb_pompe").Cells(1, Colonne) = extraction_nb_pompe
    Feuil10.Range("dec_reactif_extraction_Prefoulement").Cells(1, Colonne) = extraction_P_refoulement
    Feuil10.Range("dec_reactif_extraction_tps_fonctionnement").Cells(1, Colonne) = extraction_tps_fct_pompe
    Feuil10.Range("dec_reactif_extraction_Qunitaire").Cells(1, Colonne) = extraction_Q_unitaire
    Feuil10.Range("dec_reactif_extraction_pompe_rdt").Cells(1, Colonne) = extraction_pompe_rdt
    Feuil10.Range("dec_reactif_electricite").Cells(1, Colonne) = electricite

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "10.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On efface les valeurs sur Excel
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil10.Range("dec_reactif_bypass").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_dosage_FeCl3").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_dosage_polymere").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_NombreOuvrages").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_Vmirroir_limite").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_Smirroir").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_Vmirroir_limite_recalculee").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_hauteur").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_Vmirroir_moyenne").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_tps_retention").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_rdt_MES").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_rdt_DCO").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_rdt_DBO").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_rdt_NK").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_rdt_P").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_MV_MES").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_recirculation_taux").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_recirculation_nb_pompe").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_recirculation_P_refoulement").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_recirculation_tps_fonctionnement").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_recirculation_Qunitaire").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_recirculation_pompe_rdt").Cells(1, Colonne) = ""      'recirculation_pompe_rdt
    Feuil10.Range("dec_reactif_boues_concentration").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_boues_MES").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_boues_Q").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_extraction_nb_pompe").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_extraction_Prefoulement").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_extraction_tps_fonctionnement").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_extraction_Qunitaire").Cells(1, Colonne) = ""
    Feuil10.Range("dec_reactif_extraction_pompe_rdt").Cells(1, Colonne) = ""     'extraction_pompe_rdt
    Feuil10.Range("dec_reactif_electricite").Cells(1, Colonne) = ""


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "10.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 4  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil10.Range("dec_reactif_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "10.8"

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









