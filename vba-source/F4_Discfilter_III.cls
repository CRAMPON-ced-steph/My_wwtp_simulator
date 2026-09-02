Option Explicit

'constantes pour les choix
Private Const dix_micrometer = 1
Private Const dixhuit_micrometer = 2
Dim choix_filtre As Integer

'choix type polymère
Public choix_polymere_type As Integer

'guide defaut

'hypothese
Dim fonction_limit As Double
Dim p_installee As Double
Dim Q_moyen As Double
Dim rdt_limit_MES_elim As Double
Dim a_rdt_MES(4) As Double
Dim a_ES_concentration_MES(2) As Double

'calcul
'pompes lavage
Dim ratio_ES_lavage As Double
Dim Q_lavage As Double
Dim Q_pompe_lavage As Double
Dim ratio_eau_lavage As Double
Dim nombre_pompe_lavage As Double
Dim duree_func As Double
Dim pression_refoulement_lavage As Double
Dim rdt_machine_pompe_lavage As Double
Dim puissance_pompe_lavage As Double
Dim rdt_moteur_lavage As Double
Dim puissance_moteur_lavage As Double
Dim elec_lavage As Double
Dim rdt_pompe_lavage As Double

'disques
Dim duree_disques As Double
Dim ratio_disques_Pw_m3 As Double
Dim puissance_installee_disques As Double
Dim rdt_moteur_disques As Double
Dim puissance_moteur_disques As Double
Dim elec_disques As Double
'pompage retour
Dim volume_eaux_sales As Double
Dim pression_refoulement_retour As Double
Dim nombre_pompe_retour As Double
Dim Q_pompe_retour As Double
Dim rdt_machine_pompe_retour As Double
Dim puissance_pompe_retour As Double
Dim rdt_moteur_retour As Double
Dim puissance_moteur_retour As Double
Dim elec_retour As Double
Dim rdt_pompe_retour As Double
'eaux sales
Dim Q_eaux_sales As Double
Dim MES_eaux_sales As Double
Dim ES_concentration_MES As Double
Dim conc_MES_entree As Double
Dim rdt_MES As Double
Dim rdt_DCO As Double
Dim rdt_DBO As Double
Dim rdt_NK As Double
Dim rdt_P As Double
Dim dosage_polymere As Double
Dim dosage_FeCl3 As Double
Public polymere_flux As Double
Public FeCl3_flux As Double

Dim electricite_discfilter As Double
Public electricite As Double

'verifie forcée
Dim rdt_MES_force As Boolean
Dim rdt_DCO_force As Boolean
Dim rdt_DBO_force As Boolean
Dim rdt_NK_force As Boolean
Dim rdt_P_force As Boolean
Dim ES_concentration_MES_force As Boolean
Dim pression_refoulement_lavage_force As Boolean
Dim rdt_pompe_lavage_force As Boolean
Dim pression_refoulement_retour_force As Boolean
Dim rdt_pompe_retour_force As Boolean
Dim dosage_polymere_force As Boolean
Dim dosage_FeCl3_force As Boolean

'si necessaire
Dim duree_func_force As Boolean

'valeur forcée
Dim dosage_polymere_f As Double
Dim dosage_FeCl3_f As Double
Dim rdt_MES_f As Double
Dim rdt_DCO_f As Double
Dim rdt_DBO_f As Double
Dim rdt_NK_f As Double
Dim rdt_P_f As Double
Dim ES_concentration_MES_f As Double
Dim pression_refoulement_lavage_f As Double
Dim rdt_pompe_lavage_f As Double
Dim pression_refoulement_retour_f As Double
Dim rdt_pompe_retour_f As Double
'si necessaire
Dim duree_func_f As Double

Dim remarque(2) As String

'hypothèses
Dim rate_MES_FeCl3 As Double
Dim hyp_rdt_P_particulaire As Double
Dim ratio_Psol_P As Double
Dim ratio_FeCl3_Pt As Double

Dim stockage_Q As Double


Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'lecture des choix internes
    choix_filtre = Feuil23.Range("discfil_choix_filtre").Cells(1, pi_FiliereConsideree)
    choix_polymere_type = Feuil23.Range("discfil_choix_polymere_type").Cells(1, pi_FiliereConsideree)
    If choix_BA_forte = True Then
        remarque(1) = Feuil23.Range("discfil_erreur").Cells(1, 1)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "23.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    ratio_ES_lavage = 0.5                  '%
    nombre_pompe_lavage = 2              'unites
    nombre_pompe_retour = 2              'unites
    ratio_eau_lavage = 0.5               '%
    fonction_limit = 60                  'kW
    p_installee = 1.1                    'kW 'sur Rambouillet moteur de 1,1 kW (P installée) pour débit moyen de 7000 m3/h
    Q_moyen = 7000                       'm3/h
    rdt_moteur_disques = 0.8             '%
    duree_func = 24                      'h
    rdt_limit_MES_elim = 20              '%
    rdt_DCO = 0.1                         '%
    rdt_DBO = 0.15                          '%
    rdt_NK = 0.11                         '%

    a_rdt_MES(3) = 0.00001
    a_rdt_MES(2) = -0.0012
    a_rdt_MES(1) = 0.0428
    a_rdt_MES(0) = 0.23
    a_ES_concentration_MES(1) = 0.0455
    a_ES_concentration_MES(0) = 0.075
    rdt_machine_pompe_lavage = 0.7               '%
    rdt_machine_pompe_retour = 0.7              '%
    dosage_polymere = 1.5     'mg/l


    rate_MES_FeCl3 = 2.8 * 55.845 / 162.04    'kg MES/kg FeCl3   pour calcul MES néo-formés
    ratio_Psol_P = 0.95  'modifié le 13/07/2011 par DCA et CAM après comparaison avec SIMULO
    ratio_FeCl3_Pt = 162.5 / 31

    'rendement P particulaire
    hyp_rdt_P_particulaire = 0   'modifié le 13/07/2011 par DCA et CAM après comparaison avec SIMULO



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "23.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    If pression_refoulement_lavage_force = True Then
        pression_refoulement_lavage = pression_refoulement_lavage_f
    Else
        pression_refoulement_lavage = 80     'mCE
    End If
    If pression_refoulement_retour_force = True Then
        pression_refoulement_retour = pression_refoulement_retour_f
    Else
        pression_refoulement_retour = 10     'mCE
    End If
    If rdt_P_force = True Then
        rdt_P = rdt_P_f
    End If
    If rdt_NK_force = True Then
        rdt_NK = rdt_NK_f
    End If
    If rdt_DBO_force = True Then
        rdt_DBO = rdt_DBO_f
    End If
    If rdt_DCO_force = True Then
        rdt_DCO = rdt_DCO_f
    End If
    If dosage_polymere_force = True Then
        dosage_polymere = dosage_polymere_f
    End If
    If dosage_FeCl3_force = True Then
        dosage_FeCl3 = dosage_FeCl3_f
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "23.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim column As Integer

    column = pi_FiliereConsideree

    rdt_MES_force = Feuil23.Range("discfil_rdt_MES").Cells(1, column)
    rdt_DCO_force = Feuil23.Range("discfil_rdt_DCO").Cells(1, column)
    rdt_DBO_force = Feuil23.Range("discfil_rdt_DBO").Cells(1, column)
    rdt_NK_force = Feuil23.Range("discfil_rdt_NK").Cells(1, column)
    rdt_P_force = Feuil23.Range("discfil_rdt_Pt").Cells(1, column)
    ES_concentration_MES_force = Feuil23.Range("discfil_conc_MES").Cells(1, column)
    pression_refoulement_lavage_force = Feuil23.Range("discfil_pression_refoulement_lavage").Cells(1, column)
    rdt_pompe_lavage_force = Feuil23.Range("discfil_rdt_pompe_lavage").Cells(1, column)
    pression_refoulement_retour_force = Feuil23.Range("discfil_pression_refoulement_retour").Cells(1, column)
    rdt_pompe_retour_force = Feuil23.Range("discfil_rdt_pompe_retour").Cells(1, column)
    dosage_FeCl3_force = Feuil23.Range("discfil_dosage_FeCl3").Cells(1, column)
    dosage_polymere_force = Feuil23.Range("discfil_dosage_polymere").Cells(1, column)

    If rdt_MES_force = True Then
        rdt_MES_f = Feuil23.Range("discfil_rdt_MES").Cells(1, column)
    End If
    If rdt_DCO_force = True Then
        rdt_DCO_f = Feuil23.Range("discfil_rdt_DCO").Cells(1, column)
    End If
    If rdt_DBO_force = True Then
        rdt_DBO_f = Feuil23.Range("discfil_rdt_DBO").Cells(1, column)
    End If
    If rdt_NK_force = True Then
        rdt_NK_f = Feuil23.Range("discfil_rdt_NK").Cells(1, column)
    End If
    If rdt_P_force = True Then
        rdt_P_f = Feuil23.Range("discfil_rdt_Pt").Cells(1, column)
    End If

    If ES_concentration_MES_force = True Then
        ES_concentration_MES_f = Feuil23.Range("discfil_conc_MES").Cells(1, column)
    End If

    If pression_refoulement_lavage_force = True Then
        pression_refoulement_lavage_f = Feuil23.Range("discfil_pression_refoulement_lavage").Cells(1, column)
    End If

    If pression_refoulement_retour_force = True Then
        pression_refoulement_retour_f = Feuil23.Range("discfil_pression_refoulement_retour").Cells(1, column)
    End If

    If rdt_pompe_lavage_force = True Then
        rdt_pompe_lavage_f = Feuil23.Range("discfil_rdt_pompe_lavage").Cells(1, column)
    End If

    If rdt_pompe_retour_force = True Then
        rdt_pompe_retour_f = Feuil23.Range("discfil_rdt_pompe_retour").Cells(1, column)
    End If

    If dosage_polymere_force = True Then
        dosage_polymere_f = Feuil23.Range("discfil_dosage_polymere").Cells(1, column)
    End If

    If dosage_FeCl3_force = True Then
        dosage_FeCl3_f = Feuil23.Range("discfil_dosage_FeCl3").Cells(1, column)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "23.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub dimensionnement(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef ES_discfilter)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim MES_formes As Double
    Dim P_eliminer As Double
    Dim Ptbr As Double
    Dim rdt_P_soluble As Double
    Dim ratio_molaire_Ptbr As Double

    conc_MES_entree = MES / Q * 1000   'mg/L

    If choix_filtre = dix_micrometer Then
        'calcul du rendement MES
        If rdt_MES_force = True Then
            rdt_MES = rdt_MES_f
        Else
            rdt_MES = 0
            For I = 0 To 3
                rdt_MES = rdt_MES + a_rdt_MES(I) * conc_MES_entree ^ I
            Next I
        End If
        'calcul de la concentration en eaux sales
        If ES_concentration_MES_force = True Then
            ES_concentration_MES = ES_concentration_MES_f / 1000
        Else
            ES_concentration_MES = 0
            For I = 0 To 1
                ES_concentration_MES = ES_concentration_MES + a_ES_concentration_MES(I) * conc_MES_entree ^ I
            Next I
        End If
    Else
        If conc_MES_entree < 30 Then
            If rdt_MES_force = True Then
                rdt_MES = rdt_MES_f
            Else
                rdt_MES = 0.55
            End If
            If ES_concentration_MES_force = True Then
                ES_concentration_MES = ES_concentration_MES_f / 1000
            Else
                ES_concentration_MES = 0.8
            End If
        Else
            If rdt_MES_force = True Then
                rdt_MES = rdt_MES_f
            Else
                rdt_MES = 0.65
            End If
            If ES_concentration_MES_force = True Then
                ES_concentration_MES = ES_concentration_MES_f / 1000
            Else
                ES_concentration_MES = 1.25
            End If
        End If
    End If

    'calcul du rendement en Pt et du dosage de FeCl3
    If rdt_P_force = True Then
        rdt_P = rdt_P_f
        If dosage_FeCl3_force = True Then
            '**** à voir si on met un message pour vérifier la cohérence *****
        Else
            rdt_P_soluble = (rdt_P - (1 - ratio_Psol_P) * hyp_rdt_P_particulaire) / ratio_Psol_P
            If rdt_P_soluble < 0 Then
                dosage_FeCl3 = 0
            Else
                If rdt_P_soluble > 1 Then
                    rdt_P_soluble = 1
                End If
                P_eliminer = rdt_P_soluble * ratio_Psol_P * Pt
                Ptbr = P_eliminer / Q * 1000
                ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + Pt_garantie) ^ (-0.5) - ((Pt_garantie) ^ (-0.5) - 0.1)) / ((Ptbr + Pt_garantie) - (Pt_garantie - 0.1))
                dosage_FeCl3 = ratio_molaire_Ptbr * ratio_FeCl3_Pt * Ptbr
                If dosage_FeCl3 < 0 Then
                    dosage_FeCl3 = 0
                End If
            End If
        End If
    Else
        If dosage_FeCl3_force = True Then
            rdt_P_soluble = coagulation_floculation_rdt_P_soluble(Q, Pt)
            If rdt_P_soluble > 1 Then
                rdt_P_soluble = 1
            End If
            rdt_P = ratio_Psol_P * rdt_P_soluble + (1 - ratio_Psol_P) * hyp_rdt_P_particulaire
        Else
            P_eliminer = Pt - Pt_garantie * Q / 1000
            If P_eliminer < 0 Then
                rdt_P = (1 - ratio_Psol_P) * hyp_rdt_P_particulaire    'pas d'ajout de réactifs
                dosage_FeCl3 = 0
            Else
                rdt_P = P_eliminer / Pt
                rdt_P_soluble = (rdt_P - (1 - ratio_Psol_P) * hyp_rdt_P_particulaire) / ratio_Psol_P
                If rdt_P_soluble < 0 Then
                    dosage_FeCl3 = 0
                Else
                    If rdt_P_soluble > 1 Then
                        rdt_P_soluble = 1
                    End If
                    P_eliminer = rdt_P_soluble * ratio_Psol_P * Pt
                    Ptbr = P_eliminer / Q * 1000
                    ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + Pt_garantie) ^ (-0.5) - ((Pt_garantie) ^ (-0.5) - 0.1)) / ((Ptbr + Pt_garantie) - (Pt_garantie - 0.1))
                    dosage_FeCl3 = ratio_molaire_Ptbr * ratio_FeCl3_Pt * Ptbr
                    If dosage_FeCl3 < 0 Then
                        dosage_FeCl3 = 0
                    End If
                End If
            End If
        End If
    End If

    'poids du FeCl3 dans les MES
    FeCl3_flux = dosage_FeCl3 * Q / 1000
    MES_formes = FeCl3_flux * rate_MES_FeCl3
    polymere_flux = Q / 1000 * dosage_polymere


    MES_eaux_sales = MES * rdt_MES + MES_formes
    Q_eaux_sales = MES_eaux_sales / ES_concentration_MES
    Q_lavage = Q_eaux_sales / ratio_ES_lavage


    'calcul des eaux sales de discfilter
    ES_discfilter(repere_ES_Q) = Q_eaux_sales
    ES_discfilter(repere_ES_DCO) = rdt_DCO * DCO
    ES_discfilter(repere_ES_DBO) = rdt_DBO * DBO
    ES_discfilter(repere_ES_MES) = MES_eaux_sales
    ES_discfilter(repere_ES_NK) = rdt_NK * NK
    ES_discfilter(repere_ES_NH4) = NH4 * Q_eaux_sales / Q
    ES_discfilter(repere_ES_NO3) = NO3 * Q_eaux_sales / Q
    ES_discfilter(repere_ES_Pt) = rdt_P * Pt


    'calcul des débit et charges en sortie eau
    Q = Q + (Q_lavage - Q_eaux_sales)
    DCO = DCO - rdt_DCO * DCO
    DBO = DBO - rdt_DBO * DBO
    MES = MES * (1 - rdt_MES)
    NK = NK - rdt_NK * NK
    NH4 = NH4 * (1 - Q_eaux_sales / (Q - (Q_lavage - Q_eaux_sales)))  'on a modifié le Q avant donc attention
    NO3 = NO3 * (1 - Q_eaux_sales / (Q - (Q_lavage - Q_eaux_sales)))    'on a modifié le Q avant donc attention
    Pt = Pt - rdt_P * Pt


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "23.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


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
        Psol_in = Pt / Q * 1000 * ratio_Psol_P
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

Sub fonctionnement_reel(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef ES_discfilter)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim MES_formes As Double
    Dim P_eliminer As Double
    Dim Ptbr As Double
    Dim rdt_P_soluble As Double
    Dim ratio_molaire_Ptbr As Double


    stockage_Q = Q

    conc_MES_entree = MES / Q * 1000   'mg/L

    If choix_filtre = dix_micrometer Then
        'calcul du rendement MES
        If rdt_MES_force = True Then
            rdt_MES = rdt_MES_f
        Else
            rdt_MES = 0
            For I = 0 To 3
                rdt_MES = rdt_MES + a_rdt_MES(I) * conc_MES_entree ^ I
            Next I
        End If
        'calcul de la concentration en eaux sales
        If ES_concentration_MES_force = True Then
            ES_concentration_MES = ES_concentration_MES_f / 1000
        Else
            ES_concentration_MES = 0
            For I = 0 To 1
                ES_concentration_MES = ES_concentration_MES + a_ES_concentration_MES(I) * conc_MES_entree ^ I
            Next I
        End If
    Else
        If conc_MES_entree < 30 Then
            If rdt_MES_force = True Then
                rdt_MES = rdt_MES_f
            Else
                rdt_MES = 0.55
            End If
            If ES_concentration_MES_force = True Then
                ES_concentration_MES = ES_concentration_MES_f / 1000
            Else
                ES_concentration_MES = 0.8
            End If
        Else
            If rdt_MES_force = True Then
                rdt_MES = rdt_MES_f
            Else
                rdt_MES = 0.65
            End If
            If ES_concentration_MES_force = True Then
                ES_concentration_MES = ES_concentration_MES_f / 1000
            Else
                ES_concentration_MES = 1.25
            End If
        End If
    End If

    'calcul du rendement en Pt et du dosage de FeCl3
    If rdt_P_force = True Then
        rdt_P = rdt_P_f
        If dosage_FeCl3_force = True Then
            '**** à voir si on met un message pour vérifier la cohérence *****
        Else
            rdt_P_soluble = (rdt_P - (1 - ratio_Psol_P) * hyp_rdt_P_particulaire) / ratio_Psol_P
            If rdt_P_soluble < 0 Then
                dosage_FeCl3 = 0
            Else
                If rdt_P_soluble > 1 Then
                    rdt_P_soluble = 1
                End If
                P_eliminer = rdt_P_soluble * ratio_Psol_P * Pt
                Ptbr = P_eliminer / Q * 1000
                ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + Pt_garantie) ^ (-0.5) - ((Pt_garantie) ^ (-0.5) - 0.1)) / ((Ptbr + Pt_garantie) - (Pt_garantie - 0.1))
                dosage_FeCl3 = ratio_molaire_Ptbr * ratio_FeCl3_Pt * Ptbr
                If dosage_FeCl3 < 0 Then
                    dosage_FeCl3 = 0
                End If
            End If
        End If
    Else
        If dosage_FeCl3_force = True Then
            rdt_P_soluble = coagulation_floculation_rdt_P_soluble(Q, Pt)
            If rdt_P_soluble > 1 Then
                rdt_P_soluble = 1
            End If
            rdt_P = ratio_Psol_P * rdt_P_soluble + (1 - ratio_Psol_P) * hyp_rdt_P_particulaire
        Else
            P_eliminer = Pt - Pt_garantie * Q / 1000
            If P_eliminer < 0 Then
                rdt_P = (1 - ratio_Psol_P) * hyp_rdt_P_particulaire    'pas d'ajout de réactifs
                dosage_FeCl3 = 0
            Else
                rdt_P = P_eliminer / Pt
                rdt_P_soluble = (rdt_P - (1 - ratio_Psol_P) * hyp_rdt_P_particulaire) / ratio_Psol_P
                If rdt_P_soluble < 0 Then
                    dosage_FeCl3 = 0
                Else
                    If rdt_P_soluble > 1 Then
                        rdt_P_soluble = 1
                    End If
                    P_eliminer = rdt_P_soluble * ratio_Psol_P * Pt
                    Ptbr = P_eliminer / Q * 1000
                    ratio_molaire_Ptbr = 1.1 - 6 * ((Ptbr + Pt_garantie) ^ (-0.5) - ((Pt_garantie) ^ (-0.5) - 0.1)) / ((Ptbr + Pt_garantie) - (Pt_garantie - 0.1))
                    dosage_FeCl3 = ratio_molaire_Ptbr * ratio_FeCl3_Pt * Ptbr
                    If dosage_FeCl3 < 0 Then
                        dosage_FeCl3 = 0
                    End If
                End If
            End If
        End If
    End If

    'poids du FeCl3 dans les MES
    FeCl3_flux = dosage_FeCl3 * Q / 1000
    MES_formes = FeCl3_flux * rate_MES_FeCl3
    polymere_flux = Q / 1000 * dosage_polymere


    MES_eaux_sales = MES * rdt_MES + MES_formes
    Q_eaux_sales = MES_eaux_sales / ES_concentration_MES
    Q_lavage = Q_eaux_sales / ratio_ES_lavage


    'calcul des eaux sales de discfilter
    ES_discfilter(repere_ES_Q) = Q_eaux_sales
    ES_discfilter(repere_ES_DCO) = rdt_DCO * DCO
    ES_discfilter(repere_ES_DBO) = rdt_DBO * DBO
    ES_discfilter(repere_ES_MES) = MES_eaux_sales
    ES_discfilter(repere_ES_NK) = rdt_NK * NK
    ES_discfilter(repere_ES_NH4) = NH4 * Q_eaux_sales / Q
    ES_discfilter(repere_ES_NO3) = NO3 * Q_eaux_sales / Q
    ES_discfilter(repere_ES_Pt) = rdt_P * Pt


    'calcul des débit et charges en sortie eau
    Q = Q + (Q_lavage - Q_eaux_sales)
    DCO = DCO - rdt_DCO * DCO
    DBO = DBO - rdt_DBO * DBO
    MES = MES * (1 - rdt_MES)
    NK = NK - rdt_NK * NK
    NH4 = NH4 * (1 - Q_eaux_sales / (Q - (Q_lavage - Q_eaux_sales)))  'on a modifié le Q avant donc attention
    NO3 = NO3 * (1 - Q_eaux_sales / (Q - (Q_lavage - Q_eaux_sales)))    'on a modifié le Q avant donc attention
    Pt = Pt - rdt_P * Pt

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "23.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub calcul_consommation_electrique()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    If rdt_pompe_lavage_force = True Then
        rdt_pompe_lavage = rdt_pompe_lavage_f
        elec_lavage = (Q_lavage / pd_NOMBRE_SECONDE_PAR_HEURE) * pd_ACCELERATION_PESANTEUR_m_s2 * (pression_refoulement_lavage / rdt_pompe_lavage)
    Else
        'lavage
        Q_pompe_lavage = Q_lavage / (nombre_pompe_lavage * duree_func)
        puissance_pompe_lavage = (Q_pompe_lavage / pd_NOMBRE_SECONDE_PAR_HEURE) * pd_ACCELERATION_PESANTEUR_m_s2 * (pression_refoulement_lavage / rdt_machine_pompe_lavage)
        If puissance_pompe_lavage < fonction_limit Then
            rdt_moteur_lavage = (5.7195 * Log(puissance_pompe_lavage) + 72.682) / 100
        Else
            rdt_moteur_lavage = 0.961
        End If
        puissance_moteur_lavage = puissance_pompe_lavage / rdt_moteur_lavage
        rdt_pompe_lavage = rdt_moteur_lavage * rdt_machine_pompe_lavage
        elec_lavage = puissance_moteur_lavage * nombre_pompe_lavage * duree_func
    End If

    'disques
    ratio_disques_Pw_m3 = p_installee / Q_moyen     'kW installé /m3
    puissance_installee_disques = stockage_Q * ratio_disques_Pw_m3   'kW
    puissance_moteur_disques = puissance_installee_disques / rdt_moteur_disques   'kW
    elec_disques = duree_func * puissance_moteur_disques   'kWh/j

    'retour
    If rdt_pompe_retour_force = True Then
        rdt_pompe_retour = rdt_pompe_retour_f
        elec_retour = (Q_eaux_sales / pd_NOMBRE_SECONDE_PAR_HEURE) * pd_ACCELERATION_PESANTEUR_m_s2 * (pression_refoulement_retour / rdt_pompe_retour)
    Else
        Q_pompe_retour = Q_eaux_sales / (nombre_pompe_retour * duree_func)
        puissance_pompe_retour = (Q_pompe_retour / pd_NOMBRE_SECONDE_PAR_HEURE) * pd_ACCELERATION_PESANTEUR_m_s2 * (pression_refoulement_retour / rdt_machine_pompe_retour)
        If puissance_pompe_retour < fonction_limit Then
            rdt_moteur_retour = (5.7195 * Log(puissance_pompe_retour) + 72.682) / 100
        Else
            rdt_moteur_retour = 0.961
        End If
        puissance_moteur_retour = puissance_pompe_retour / rdt_moteur_retour
        rdt_pompe_retour = rdt_moteur_retour * rdt_machine_pompe_retour
        elec_retour = puissance_moteur_retour * nombre_pompe_retour * duree_func
    End If


    electricite_discfilter = elec_lavage + elec_disques + elec_retour

    electricite = electricite_discfilter

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "23.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil23.Range("discfil_rdt_MES").Cells(1, Colonne) = rdt_MES
    Feuil23.Range("discfil_rdt_DCO").Cells(1, Colonne) = rdt_DCO
    Feuil23.Range("discfil_rdt_DBO").Cells(1, Colonne) = rdt_DBO
    Feuil23.Range("discfil_rdt_Nk").Cells(1, Colonne) = rdt_NK
    Feuil23.Range("discfil_rdt_Pt").Cells(1, Colonne) = rdt_P

    Feuil23.Range("discfil_MES_eaux_sales").Cells(1, Colonne) = MES_eaux_sales
    Feuil23.Range("discfil_Q_eaux_sales").Cells(1, Colonne) = Q_eaux_sales
    Feuil23.Range("discfil_conc_MES").Cells(1, Colonne) = ES_concentration_MES * 1000

    Feuil23.Range("discfil_pression_refoulement_lavage").Cells(1, Colonne) = pression_refoulement_lavage
    Feuil23.Range("discfil_rdt_pompe_lavage").Cells(1, Colonne) = rdt_pompe_lavage
    Feuil23.Range("discfil_pression_refoulement_retour").Cells(1, Colonne) = pression_refoulement_retour
    Feuil23.Range("discfil_rdt_pompe_retour").Cells(1, Colonne) = rdt_pompe_retour

    Feuil23.Range("discfil_electricite_discfilter").Cells(1, Colonne) = electricite_discfilter

    Feuil23.Range("discfil_dosage_FeCl3").Cells(1, Colonne) = dosage_FeCl3
    Feuil23.Range("discfil_dosage_polymere").Cells(1, Colonne) = dosage_polymere


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "23.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil23.Range("discfil_rdt_MES").Cells(1, Colonne) = ""
    Feuil23.Range("discfil_rdt_DCO").Cells(1, Colonne) = ""
    Feuil23.Range("discfil_rdt_DBO").Cells(1, Colonne) = ""
    Feuil23.Range("discfil_rdt_Nk").Cells(1, Colonne) = ""
    Feuil23.Range("discfil_rdt_Pt").Cells(1, Colonne) = ""

    Feuil23.Range("discfil_MES_eaux_sales").Cells(1, Colonne) = ""
    Feuil23.Range("discfil_Q_eaux_sales").Cells(1, Colonne) = ""
    Feuil23.Range("discfil_conc_MES").Cells(1, Colonne) = ""

    Feuil23.Range("discfil_pression_refoulement_lavage").Cells(1, Colonne) = ""
    Feuil23.Range("discfil_rdt_pompe_lavage").Cells(1, Colonne) = ""
    Feuil23.Range("discfil_pression_refoulement_retour").Cells(1, Colonne) = ""
    Feuil23.Range("discfil_rdt_pompe_retour").Cells(1, Colonne) = ""

    Feuil23.Range("discfil_electricite_discfilter").Cells(1, Colonne) = ""

    Feuil23.Range("discfil_dosage_FeCl3").Cells(1, Colonne) = ""
    Feuil23.Range("discfil_dosage_polymere").Cells(1, Colonne) = ""

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "23.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 1  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil23.Range("discfil_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "23.8"

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
        Call ecriture_resultats
        Call affich_remarque
    End If



End Sub





