Option Explicit

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
Dim choix_lamellaire As Integer
Public choix_polymere_type As Integer
Dim choix_microsable As Integer
Dim choix_recirculation As Integer

'variables échangées avec Excel
Public dosage_FeCl3 As Double
Public dosage_polymere As Double
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

Public electricite_fixe As Double

'variables montrant si les valeurs ont été forcées
Dim dosage_FeCl3_force As Boolean
Dim dosage_polymere_force As Boolean
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
Dim dosage_FeCl3_f As Double
Dim dosage_polymere_f As Double
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

Dim rate_MES_FeCl3 As Double
Dim Q_traite_nominal As Double
Dim Q_traite_reel As Double

Dim remarque(3) As String


Sub dimensionnement(ByRef eaux_sales, ByRef choix_eaux_sales, ByRef boues_x_detail, ByRef C_elimine, ByRef DBO_elimine, ByRef N_elimine, ByRef P_elimine)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'variables internes
    Dim I As Integer

    Dim Q As Double
    Dim DCO As Double
    Dim DBO As Double
    Dim MES As Double
    Dim NK As Double
    Dim NH4 As Double
    Dim NO3 As Double
    Dim Pt As Double
    Dim FeCl3_flux_massique As Double
    Dim MES_formes As Double

    Dim intermediaire_MV(nb_eaux_sales_max) As Double
    Dim intermediaire_MES(nb_eaux_sales_max) As Double
    Dim MES_eaux_sales As Double
    Dim inlet_MV As Double
    Dim ratio_ES As Double

    'calcul du flux d'eaux sales arrivant sur l'ouvrage
    Q = 0
    DCO = 0
    DBO = 0
    MES = 0
    NK = 0
    NH4 = 0
    NO3 = 0
    Pt = 0
    For I = 1 To nb_eaux_sales_max
        If choix_eaux_sales(I) <> devenir_ES_primaire Then
            'flux entrant
            Q = Q + eaux_sales(I, repere_ES_Q)
            DCO = DCO + eaux_sales(I, repere_ES_DCO)
            DBO = DBO + eaux_sales(I, repere_ES_DBO)
            MES = MES + eaux_sales(I, repere_ES_MES)
            NK = NK + eaux_sales(I, repere_ES_NK)
            NH4 = NH4 + eaux_sales(I, repere_ES_NH4)
            NO3 = NO3 + eaux_sales(I, repere_ES_NO3)
            Pt = Pt + eaux_sales(I, repere_ES_Pt)
            'report de l'origine des boues
            boues_x_detail(I, repere_origine) = eaux_sales(I, repere_ES_origine)
            'calcul des MV de chaque type d'eaux sales
            intermediaire_MV(I) = eaux_sales(I, repere_ES_MV_MES) * eaux_sales(I, repere_ES_MES)
            'calcul des MES de chaque type d'eaux sales
            intermediaire_MES(I) = eaux_sales(I, repere_ES_MES)
            inlet_MV = inlet_MV + intermediaire_MV(I)
        End If
    Next I
    'on calcule la provenance des MV par type de boues (ratio par rapport aux MV totaux)
    For I = 1 To nb_eaux_sales_max
        boues_x_detail(I, repere_ratio_MV) = intermediaire_MV(I) / inlet_MV
        boues_x_detail(I, repere_ratio_MES) = intermediaire_MES(I) / MES
    Next I

    'variables pour la composition des boues extraites
    C_elimine = DCO
    DBO_elimine = DBO
    N_elimine = NK
    P_elimine = Pt
    Q_traite_nominal = Q

    'poids du FeCl3 dans les MES
    FeCl3_flux_massique = dosage_FeCl3 * Q / 1000
    MES_formes = FeCl3_flux_massique * rate_MES_FeCl3


    'on calcule le ratio de MV/MES
    If MV_MES_force = True Then
        MV_MES = MV_MES_f
    Else
        MV_MES = inlet_MV / (MES + MES_formes)
    End If


    'calcul des boues à extraire
    boues_MES = (MES + MES_formes) * rdt_MES
    boues_Q = boues_MES / boues_concentration


    'calcul des débit et charges en sortie eau
    Q = Q - boues_Q
    DCO = DCO * (1 - rdt_DCO)
    DBO = DBO * (1 - rdt_DBO)
    MES = (MES + MES_formes) * (1 - rdt_MES)
    NK = NK * (1 - rdt_NK)
    NH4 = NH4 * Q / (Q + boues_Q)
    NO3 = NO3 * Q / (Q + boues_Q)
    Pt = Pt * (1 - rdt_P)

    C_elimine = C_elimine - DCO
    DBO_elimine = DBO_elimine - DBO
    N_elimine = N_elimine - NK
    P_elimine = P_elimine - Pt


    'report des valeurs dans le vecteur des eaux sales (de façon aléatoire entre les flux 1 et 2 en considérant les proportions de MES)
    For I = 1 To nb_eaux_sales_max
        ratio_ES = boues_x_detail(I, repere_ratio_MES)
        If ratio_ES <> 0 Then
            eaux_sales(I, repere_ES_Q) = ratio_ES * Q
            eaux_sales(I, repere_ES_DCO) = ratio_ES * DCO
            eaux_sales(I, repere_ES_DBO) = ratio_ES * DBO
            eaux_sales(I, repere_ES_MES) = ratio_ES * MES
            eaux_sales(I, repere_ES_NK) = ratio_ES * NK
            eaux_sales(I, repere_ES_NH4) = ratio_ES * NH4
            eaux_sales(I, repere_ES_NO3) = ratio_ES * NO3
            eaux_sales(I, repere_ES_Pt) = ratio_ES * Pt
        End If
    Next I



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "19.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub fonctionnement_reel(ByRef eaux_sales, ByRef choix_eaux_sales, ByRef boues_x_detail, ByRef C_elimine, ByRef DBO_elimine, ByRef N_elimine, ByRef P_elimine)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'variables internes
    Dim I As Integer

    Dim Q As Double
    Dim DCO As Double
    Dim DBO As Double
    Dim MES As Double
    Dim NK As Double
    Dim NH4 As Double
    Dim NO3 As Double
    Dim Pt As Double
    Dim FeCl3_flux_massique As Double
    Dim MES_formes As Double
    Dim polymere_flux_massique As Double
    Dim ratio_ES As Double

    Dim intermediaire_MV(nb_eaux_sales_max) As Double
    Dim intermediaire_MES(nb_eaux_sales_max) As Double
    Dim MES_eaux_sales As Double
    Dim inlet_MV As Double


    'calcul du flux d'eaux sales arrivant sur l'ouvrage
    Q = 0
    DCO = 0
    DBO = 0
    MES = 0
    NK = 0
    NH4 = 0
    NO3 = 0
    Pt = 0
    For I = 1 To nb_eaux_sales_max
        If choix_eaux_sales(I) <> devenir_ES_primaire Then
            'flux entrant
            Q = Q + eaux_sales(I, repere_ES_Q)
            DCO = DCO + eaux_sales(I, repere_ES_DCO)
            DBO = DBO + eaux_sales(I, repere_ES_DBO)
            MES = MES + eaux_sales(I, repere_ES_MES)
            NK = NK + eaux_sales(I, repere_ES_NK)
            NH4 = NH4 + eaux_sales(I, repere_ES_NH4)
            NO3 = NO3 + eaux_sales(I, repere_ES_NO3)
            Pt = Pt + eaux_sales(I, repere_ES_Pt)
            'report de l'origine des boues
            boues_x_detail(I, repere_origine) = eaux_sales(I, repere_ES_origine)
            'calcul des MV de chaque type d'eaux sales
            intermediaire_MV(I) = eaux_sales(I, repere_ES_MV_MES) * eaux_sales(I, repere_ES_MES)
            'calcul des MES de chaque type d'eaux sales
            intermediaire_MES(I) = eaux_sales(I, repere_ES_MES)
            inlet_MV = inlet_MV + intermediaire_MV(I)
        End If
    Next I
    'on calcule la provenance des MV par type de boues (ratio par rapport aux MV totaux)
    For I = 1 To nb_eaux_sales_max
        boues_x_detail(I, repere_ratio_MV) = intermediaire_MV(I) / inlet_MV
        boues_x_detail(I, repere_ratio_MES) = intermediaire_MES(I) / MES
    Next I

    'variables pour la composition des boues extraites
    C_elimine = DCO
    DBO_elimine = DBO
    N_elimine = NK
    P_elimine = Pt
    Q_traite_reel = Q

    'poids du FeCl3 dans les MES
    FeCl3_flux_massique = dosage_FeCl3 * Q / 1000
    MES_formes = FeCl3_flux_massique * rate_MES_FeCl3

    'polymère ajouté
    polymere_flux_massique = dosage_polymere * Q / 1000


    'on calcule le ratio de MV/MES
    If MV_MES_force = True Then
        MV_MES = MV_MES_f
    Else
        MV_MES = inlet_MV / (MES + MES_formes)
    End If


    'calcul des boues à extraire
    boues_MES = (MES + MES_formes) * rdt_MES
    If boues_MES_force = True Then
        boues_MES = boues_MES_f
    End If
    boues_Q = boues_MES / boues_concentration
    If boues_Q_force = True Then
        boues_Q = boues_Q_f
    End If


    'calcul des débit et charges en sortie eau
    Q = Q - boues_Q
    DCO = DCO * (1 - rdt_DCO)
    DBO = DBO * (1 - rdt_DBO)
    MES = (MES + MES_formes) * (1 - rdt_MES)
    NK = NK * (1 - rdt_NK)
    NH4 = NH4 * Q / (Q + boues_Q)
    NO3 = NO3 * Q / (Q + boues_Q)
    Pt = Pt * (1 - rdt_P)

    'MSM 15/11/12 retour biostyr
    'sortie_Q_m3_j =Q
    'sortie_DCO_kg_j =DCO
    '...

    C_elimine = C_elimine - DCO
    DBO_elimine = DBO_elimine - DBO
    N_elimine = N_elimine - NK
    P_elimine = P_elimine - Pt


    'report des valeurs dans le vecteur des eaux sales (de façon aléatoire entre les flux 1 et 2 en considérant les proportions de MES)
    For I = 1 To nb_eaux_sales_max
        ratio_ES = boues_x_detail(I, repere_ratio_MES)
        If ratio_ES <> 0 Then
            eaux_sales(I, repere_ES_Q) = ratio_ES * Q
            eaux_sales(I, repere_ES_DCO) = ratio_ES * DCO
            eaux_sales(I, repere_ES_DBO) = ratio_ES * DBO
            eaux_sales(I, repere_ES_MES) = ratio_ES * MES
            eaux_sales(I, repere_ES_NK) = ratio_ES * NK
            eaux_sales(I, repere_ES_NH4) = ratio_ES * NH4
            eaux_sales(I, repere_ES_NO3) = ratio_ES * NO3
            eaux_sales(I, repere_ES_Pt) = ratio_ES * Pt
        End If
    Next I



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "19.6"

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

    Dim electricite_agitation As Double
    Dim electricite_racleur As Double
    Dim electricite_reactifs As Double
    Dim electricite_recirculation As Double
    Dim electricite_extraction As Double
    Dim electricite_autre As Double

    'erreurs
    Dim incoherence As Double


    'AGITATION CUVES COAGULATION-FLOCULATION
    rate_elec_agitation_coagulation = 1.425    'Wh/m3  moyenne(1.08;1.77)
    rate_elec_agitation_floculation = 1.5    'Wh/m3  vérifier sources
    electricite_agitation = (rate_elec_agitation_coagulation + rate_elec_agitation_floculation) * Q_traite_nominal / 1000

    'RACLEUR
    rate_puissance_fixe = 1.5    'Wh/m3 par rapport au nominal    'modifié par DCA le 15/07/11 valeur Pascal JOUAFFRE
    rate_puissance_debit = 0    'Wh/m3 par rapport au débit réel
    'Nantes: racleur/herse: 0.38 Wh/m3
    'Angers racleur/herse: 0.88 Wh/m3
    electricite_racleur = 1 / 1000 * (rate_puissance_fixe * Q_traite_nominal + rate_puissance_debit * Q_traite_reel)

    'INJECTION DE REACTIF, POLYMERE ET MICROSABLE
    rate_elec_polymere_preparation = 0.88    'Wh/m3 revoir les sources
    rate_elec_polymere_pompage = 0.705    'Wh/m3 revoir les sources
    rate_elec_FeCl3 = 0.35
    If choix_microsable = microsable_oui Then
        rate_elec_microsable = 0.3
    Else
        rate_elec_microsable = 0
    End If
    electricite_reactifs = (rate_elec_polymere_preparation + rate_elec_polymere_pompage + rate_elec_FeCl3 + rate_elec_microsable) * Q_traite_reel / 1000

    'RECIRCULATION DES BOUES
    If choix_recirculation = recirculation_oui Then
        'recirculation_Q_unitaire = recirculation_taux * Q_traite_reel / (recirculation_nb_pompe * recirculation_tps_fct_pompe)
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
            remarque(1) = Feuil19.Range("dec_ES_erreur_pompe_recirculation").Cells(1, 1)
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

    'EXTRACTION DES BOUES
    'extraction_Q_unitaire = boues_Q / (extraction_nb_pompe * extraction_tps_fct_pompe)
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
        remarque(2) = Feuil19.Range("dec_ES_erreur_pompe_extraction").Cells(1, 1)
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
    rate_elec_autre = (4.56 + 31.19) / 32.9
    electricite_autre = rate_elec_autre * Q_traite_reel / 1000

    'CONSOMMATION GLOBALE D'ELECTRICITE
    electricite = electricite_racleur + electricite_recirculation + electricite_extraction + electricite_agitation + electricite_reactifs + electricite_autre
    electricite_fixe = electricite_racleur + electricite_agitation

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "19.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'lecture des choix internes


    choix_lamellaire = Feuil19.Range("dec_ES_choix_lamelle").Cells(1, pi_FiliereConsideree)
    choix_polymere_type = Feuil19.Range("dec_ES_choix_polymere_type").Cells(1, pi_FiliereConsideree)
    choix_microsable = Feuil19.Range("dec_ES_choix_microsable").Cells(1, pi_FiliereConsideree)
    choix_recirculation = Feuil19.Range("dec_ES_choix_recirculation").Cells(1, pi_FiliereConsideree)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "19.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'lecture du choix de forcer des valeurs
    dosage_FeCl3_force = Feuil19.Range("dec_ES_dosage_FeCl3").Cells(1, pi_FiliereConsideree)
    dosage_polymere_force = Feuil19.Range("dec_ES_dosage_polymere").Cells(1, pi_FiliereConsideree)
    rdt_MES_force = Feuil19.Range("dec_ES_rdt_MES").Cells(1, pi_FiliereConsideree)
    rdt_DCO_force = Feuil19.Range("dec_ES_rdt_DCO").Cells(1, pi_FiliereConsideree)
    rdt_DBO_force = Feuil19.Range("dec_ES_rdt_DBO").Cells(1, pi_FiliereConsideree)
    rdt_NK_force = Feuil19.Range("dec_ES_rdt_NK").Cells(1, pi_FiliereConsideree)
    rdt_P_force = Feuil19.Range("dec_ES_rdt_P").Cells(1, pi_FiliereConsideree)
    MV_MES_force = Feuil19.Range("dec_ES_MV_MES").Cells(1, pi_FiliereConsideree)
    If choix_recirculation = recirculation_oui Then
        recirculation_taux_force = Feuil19.Range("dec_ES_recirculation_taux").Cells(1, pi_FiliereConsideree)
        recirculation_nb_pompe_force = Feuil19.Range("dec_ES_recirculation_nb_pompe").Cells(1, pi_FiliereConsideree)
        recirculation_P_refoulement_force = Feuil19.Range("dec_ES_recirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
        recirculation_tps_fct_pompe_force = Feuil19.Range("dec_ES_recirculation_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
        recirculation_Q_unitaire_force = Feuil19.Range("dec_ES_recirculation_Qunitaire").Cells(1, pi_FiliereConsideree)
        recirculation_pompe_rdt_force = Feuil19.Range("dec_ES_recirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    boues_concentration_force = Feuil19.Range("dec_ES_boues_concentration").Cells(1, pi_FiliereConsideree)
    boues_MES_force = Feuil19.Range("dec_ES_boues_MES").Cells(1, pi_FiliereConsideree)
    boues_Q_force = False
    extraction_nb_pompe_force = Feuil19.Range("dec_ES_extraction_nb_pompe").Cells(1, pi_FiliereConsideree)
    extraction_P_refoulement_force = Feuil19.Range("dec_ES_extraction_Prefoulement").Cells(1, pi_FiliereConsideree)
    extraction_tps_fct_pompe_force = Feuil19.Range("dec_ES_extraction_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    extraction_Q_unitaire_force = Feuil19.Range("dec_ES_extraction_Qunitaire").Cells(1, pi_FiliereConsideree)
    extraction_pompe_rdt_force = Feuil19.Range("dec_ES_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)

    'lecture des valeurs forcées
    If dosage_FeCl3_force = True Then
        dosage_FeCl3_f = Feuil19.Range("dec_ES_dosage_FeCl3").Cells(1, pi_FiliereConsideree)
    End If
    If dosage_polymere_force = True Then
        dosage_polymere_f = Feuil19.Range("dec_ES_dosage_polymere").Cells(1, pi_FiliereConsideree)
    End If
    If rdt_MES_force = True Then
        rdt_MES_f = Feuil19.Range("dec_ES_rdt_MES").Cells(1, pi_FiliereConsideree)
    End If
    If rdt_DCO_force = True Then
        rdt_DCO_f = Feuil19.Range("dec_ES_rdt_DCO").Cells(1, pi_FiliereConsideree)
    End If
    If rdt_DBO_force = True Then
        rdt_DBO_f = Feuil19.Range("dec_ES_rdt_DBO").Cells(1, pi_FiliereConsideree)
    End If
    If rdt_NK_force = True Then
        rdt_NK_f = Feuil19.Range("dec_ES_rdt_NK").Cells(1, pi_FiliereConsideree)
    End If
    If rdt_P_force = True Then
        rdt_P_f = Feuil19.Range("dec_ES_rdt_P").Cells(1, pi_FiliereConsideree)
    End If
    If MV_MES_force = True Then
        MV_MES_f = Feuil19.Range("dec_ES_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If choix_recirculation = recirculation_oui Then
        If recirculation_taux_force = True Then
            recirculation_taux_f = Feuil19.Range("dec_ES_recirculation_taux").Cells(1, pi_FiliereConsideree)
        End If
        If recirculation_nb_pompe_force = True Then
            recirculation_nb_pompe_f = Feuil19.Range("dec_ES_recirculation_nb_pompe").Cells(1, pi_FiliereConsideree)
        End If
        If recirculation_P_refoulement_force = True Then
            recirculation_P_refoulement_f = Feuil19.Range("dec_ES_recirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
        End If
        If recirculation_tps_fct_pompe_force = True Then
            recirculation_tps_fct_pompe_f = Feuil19.Range("dec_ES_recirculation_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
        End If
        If recirculation_Q_unitaire_force = True Then
            recirculation_Q_unitaire_f = Feuil19.Range("dec_ES_recirculation_Qunitaire").Cells(1, pi_FiliereConsideree)
        End If
        If recirculation_pompe_rdt_force = True Then
            recirculation_pompe_rdt_f = Feuil19.Range("dec_ES_recirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
        End If
    End If
    If boues_concentration_force = True Then
        boues_concentration_f = Feuil19.Range("dec_ES_boues_concentration").Cells(1, pi_FiliereConsideree)
    End If
    If boues_MES_force = True Then
        boues_MES_f = Feuil19.Range("dec_ES_boues_MES").Cells(1, pi_FiliereConsideree)
    End If
    If boues_Q_force = True Then
        boues_Q_f = Feuil19.Range("dec_ES_boues_Q").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_nb_pompe_force = True Then
        extraction_nb_pompe_f = Feuil19.Range("dec_ES_extraction_nb_pompe").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement_f = Feuil19.Range("dec_ES_extraction_Prefoulement").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_tps_fct_pompe_force = True Then
        extraction_tps_fct_pompe_f = Feuil19.Range("dec_ES_extraction_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_Q_unitaire_force = True Then
        extraction_Q_unitaire_f = Feuil19.Range("dec_ES_extraction_Qunitaire").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt_f = Feuil19.Range("dec_ES_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "19.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'Voici l'ensemble des hypothèses sur le process
    rate_MES_FeCl3 = 2.8 * 55.845 / 162.04    'kg MES/kg FeCl3   pour calcul MES néo-formés


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "19.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On retrouve ici l'ensemble des valeurs pour lesquelles on attribue une valeur guide par défaut

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

    'rendements d'élimination et concentration boues en sortie
    If choix_microsable = microsable_oui Then
        If rdt_MES_force = True Then
            rdt_MES = rdt_MES_f
        Else
            rdt_MES = 0.85
        End If
        If rdt_DCO_force = True Then
            rdt_DCO = rdt_DCO_f
        Else
            rdt_DCO = 0.85
        End If
        If rdt_DBO_force = True Then
            rdt_DBO = rdt_DBO_f
        Else
            rdt_DBO = 0.88
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
            rdt_MES = 0.85
        End If
        If rdt_DCO_force = True Then
            rdt_DCO = rdt_DCO_f
        Else
            rdt_DCO = 0.85
        End If
        If rdt_DBO_force = True Then
            rdt_DBO = rdt_DBO_f
        Else
            rdt_DBO = 0.88
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
        rdt_NK = 0.83
    End If
    If rdt_P_force = True Then
        rdt_P = rdt_P_f
    Else
        rdt_P = 0.65
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
    NumeroErreurNonGeree = "19.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil19.Range("dec_ES_dosage_FeCl3").Cells(1, Colonne) = dosage_FeCl3
    Feuil19.Range("dec_ES_dosage_polymere").Cells(1, Colonne) = dosage_polymere
    Feuil19.Range("dec_ES_rdt_MES").Cells(1, Colonne) = rdt_MES
    Feuil19.Range("dec_ES_rdt_DCO").Cells(1, Colonne) = rdt_DCO
    Feuil19.Range("dec_ES_rdt_DBO").Cells(1, Colonne) = rdt_DBO
    Feuil19.Range("dec_ES_rdt_NK").Cells(1, Colonne) = rdt_NK
    Feuil19.Range("dec_ES_rdt_P").Cells(1, Colonne) = rdt_P
    Feuil19.Range("dec_ES_MV_MES").Cells(1, Colonne) = MV_MES
    If choix_recirculation = recirculation_oui Then
        Feuil19.Range("dec_ES_recirculation_taux").Cells(1, Colonne) = recirculation_taux
        Feuil19.Range("dec_ES_recirculation_nb_pompe").Cells(1, Colonne) = recirculation_nb_pompe
        Feuil19.Range("dec_ES_recirculation_P_refoulement").Cells(1, Colonne) = recirculation_P_refoulement
        Feuil19.Range("dec_ES_recirculation_tps_fonctionnement").Cells(1, Colonne) = recirculation_tps_fct_pompe
        Feuil19.Range("dec_ES_recirculation_Qunitaire").Cells(1, Colonne) = recirculation_Q_unitaire
        Feuil19.Range("dec_ES_recirculation_pompe_rdt").Cells(1, Colonne) = recirculation_pompe_rdt
    End If
    Feuil19.Range("dec_ES_boues_concentration").Cells(1, Colonne) = boues_concentration
    Feuil19.Range("dec_ES_boues_MES").Cells(1, Colonne) = boues_MES
    Feuil19.Range("dec_ES_boues_Q").Cells(1, Colonne) = boues_Q
    Feuil19.Range("dec_ES_extraction_nb_pompe").Cells(1, Colonne) = extraction_nb_pompe
    Feuil19.Range("dec_ES_extraction_Prefoulement").Cells(1, Colonne) = extraction_P_refoulement
    Feuil19.Range("dec_ES_extraction_tps_fonctionnement").Cells(1, Colonne) = extraction_tps_fct_pompe
    Feuil19.Range("dec_ES_extraction_Qunitaire").Cells(1, Colonne) = extraction_Q_unitaire
    Feuil19.Range("dec_ES_extraction_pompe_rdt").Cells(1, Colonne) = extraction_pompe_rdt
    Feuil19.Range("dec_ES_electricite").Cells(1, Colonne) = electricite


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "19.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On efface les valeurs sur Excel
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil19.Range("dec_ES_dosage_FeCl3").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_dosage_polymere").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_rdt_MES").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_rdt_DCO").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_rdt_DBO").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_rdt_NK").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_rdt_P").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_MV_MES").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_recirculation_taux").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_recirculation_nb_pompe").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_recirculation_P_refoulement").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_recirculation_tps_fonctionnement").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_recirculation_Qunitaire").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_recirculation_pompe_rdt").Cells(1, Colonne) = ""    'recirculation_pompe_rdt
    Feuil19.Range("dec_ES_boues_concentration").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_boues_MES").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_boues_Q").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_extraction_nb_pompe").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_extraction_Prefoulement").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_extraction_tps_fonctionnement").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_extraction_Qunitaire").Cells(1, Colonne) = ""
    Feuil19.Range("dec_ES_extraction_pompe_rdt").Cells(1, Colonne) = ""    'extraction_pompe_rdt
    Feuil19.Range("dec_ES_electricite").Cells(1, Colonne) = ""


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "19.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 2  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil19.Range("dec_ES_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "19.8"

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







