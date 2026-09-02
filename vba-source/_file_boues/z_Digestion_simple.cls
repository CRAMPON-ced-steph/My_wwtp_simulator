Attribute VB_Name = "z_Digestion_simple"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'constantes pour le choix de type de digestion
Private Const mesophile = 1
Private Const thermophile = 2

'constantes pour le choix du chauffage des boues
Private Const CHAUFFAGE_BOUES_AMONT = 1
Private Const CHAUFFAGE_BOUES_RECIRCULATION = 2

'constantes pour le choix d'agitation
Private Const agitation_mecanique = 1
Private Const agitation_recirculation = 2
Private Const agitation_biogaz = 3

'constantes pour l'isolation
Private Const isolation_oui = 1
Private Const isolation_non = 2

'constantes pour le stockage de biogaz
Private Const biogaz_stock_oui = 1
Private Const biogaz_stock_non = 2


'constantes nécessaires pour l'entrée
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

Private Const nb_inlet = 3

Private Const critere_incoherence = 0.1

Dim inlet_etape_anterieure(nb_inlet) As Integer
Dim inlet_type_boues(nb_inlet) As Integer

'variables de choix
Dim choix_inlet(nb_inlet) As Integer
Dim choix_type As Integer
Dim choix_chauffage_boues As Integer
Dim choix_agitation As Integer
Dim choix_biogaz_stock As Integer
Dim choix_isolation As Integer

'variables à échanger avec l'utilisateur
Private inlet_ratio_admis(nb_inlet) As Double
Public inlet_MES As Double
Dim inlet_Q As Double
Dim inlet_concentration As Double
Public inlet_MV_MES As Double
Dim T_digesteur As Double
Dim T_inlet As Double
Dim temperature_entree_ech As Double
Dim temperature_sortie_ech As Double
Dim SRT_design As Double
Dim volume_total As Double
Dim SRT_reel As Double
Dim nb_digesteur As Double
Dim hauteur_digesteur As Double
Dim tx_recirculation_agitation As Double
Dim outlet_concentration As Double
Dim outlet_MES As Double
Dim outlet_Q As Double
Dim outlet_MV_MES As Double
Dim reduction_MV As Double
Dim extraction_pompe_nb As Integer
Dim extraction_P_refoulement As Double
Dim extraction_tps_fonctionnement As Double
Dim extraction_Q_unitaire As Double
Dim extraction_pompe_rdt As Double
Dim biogaz_teneur_CH4 As Double
Dim ratio_biogaz_MV As Double
Public biogaz_Q_Nm3j As Double
Dim energie_maintien_dig_kWhj As Double
'Dim energie_eau90 As Double
Public electricite As Double

Public electricite_fixe As Double

Dim outlet_MS As Double
Dim inlet_MS As Double

'valeurs de production de biogaz pour les co-ferments
Dim coferments_BMP_Nm3CH4_tMVapplique As Double
Dim coferments_BMP_facteur_securite As Double

'choix de forcer les valeurs
Dim T_digesteur_force As Boolean
Dim T_inlet_force As Boolean
Dim temperature_entree_ech_force As Boolean
Dim temperature_sortie_ech_force As Boolean
Dim SRT_design_force As Boolean
Dim volume_total_force As Boolean
Dim nb_digesteur_force As Boolean
Dim hauteur_digesteur_force As Boolean
Dim tx_recirculation_agitation_force As Boolean
Dim reduction_MV_force As Boolean
Dim outlet_concentration_force As Boolean
Dim outlet_MES_force As Boolean
Dim extraction_pompe_nb_force As Boolean
Dim extraction_P_refoulement_force As Boolean
Dim extraction_tps_fonctionnement_force As Boolean
Dim extraction_Q_unitaire_force As Boolean
Dim extraction_pompe_rdt_force As Boolean
Dim biogaz_teneur_CH4_force As Boolean
Dim biogaz_Q_Nm3j_force As Boolean
Dim energie_maintien_dig_kWhj_force As Boolean
'Dim energie_eau90_force As Boolean

'valeurs forcées
Dim T_digesteur_f As Double
Dim T_inlet_f As Double
Dim temperature_entree_ech_f As Double
Dim temperature_sortie_ech_f As Double
Dim SRT_design_f As Double
Dim volume_total_f As Double
Dim nb_digesteur_f As Double
Dim hauteur_digesteur_f As Double
Dim tx_recirculation_agitation_f As Double
Dim reduction_MV_f As Double
Dim outlet_concentration_f As Double
Dim outlet_MES_f As Double
Dim extraction_pompe_nb_f As Integer
Dim extraction_P_refoulement_f As Double
Dim extraction_tps_fonctionnement_f As Double
Dim extraction_Q_unitaire_f As Double
Dim extraction_pompe_rdt_f As Double
Dim biogaz_teneur_CH4_f As Double
Dim biogaz_Q_Nm3j_f As Double
Dim energie_maintien_dig_kWhj_f As Double
'Dim energie_eau90_f As Double

'hypothèses
Dim reduction_MV_reference(nb_qualites_boues) As Double
Dim T_reference As Double
Dim SRT_reference As Double
Dim ratio_SRT_lim1 As Double
Dim a_ratio_SRT_lim1 As Double
Dim ratio_SRT_lim2 As Double
Dim a_ratio_SRT_lim2 As Double
Dim b_ratio_SRT_lim2 As Double
Dim ratio_SRT_correction_max As Double
Dim correction_T As Double
Dim a_densite_biogaz(1) As Double
Dim ratio_CH4_DCO As Double
Dim ratio_NK_MVdeg As Double
Dim ratio_P_Pinlet As Double
Dim ratio_DCOsol_outlet As Double
Dim digesteur_ratio_pertes_thermiques(climat_type_nb, isolation_non) As Double
Dim efficacite_echangeur As Double
Dim Cp_boues As Double
Dim a_correction_SRT_gain As Double
Dim a_correction_SRT_n As Double
Dim hyp_tps_fct_pompe_extraction As Double
Dim masse_volumique_boues As Double
Dim delta_T_boues As Double
Dim T_limite_BT_HT As Double

'stockage valeurs
Dim biogaz_Q_Nm3j_nominal As Double

'erreurs
Dim remarque(3) As String

Sub transfert_flux_energie_vers_pinch()

    Dim nom_flux As String
    Dim T_in As Double
    Dim T_out As Double
    Dim MCp As Double

    Dim message_erreur As String

    'DIGESTION SIMPLE - CHAUFFAGE BOUES

    'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
    '    nom_flux = "DIGESTION SIMPLE - CHAUFFAGE BOUES"
    nom_flux = Feuil5.Range("nom_flux_digestion1").Cells(1, 1).Value

    'T_in : température des boues à l'entrée de l'échangeur
    If choix_chauffage_boues = CHAUFFAGE_BOUES_AMONT Then
        T_in = T_inlet
        temperature_entree_ech = T_in    ' report de la valeur pour l'écriture
    ElseIf choix_chauffage_boues = CHAUFFAGE_BOUES_RECIRCULATION Then
        If temperature_entree_ech_force = True Then
            T_in = temperature_entree_ech
        Else
            T_in = T_digesteur
            temperature_entree_ech = T_in    ' report de la valeur pour l'écriture
        End If
    End If

    'T_out : température des boues à la sortie de l'échangeur
    If choix_chauffage_boues = CHAUFFAGE_BOUES_AMONT Then
        If temperature_sortie_ech_force = True Then
            T_out = temperature_sortie_ech
        Else
            T_out = T_digesteur + (digesteur_ratio_pertes_thermiques(choix_climat, choix_isolation) * volume_total / efficacite_echangeur) / (energie_maintien_dig_kWhj - digesteur_ratio_pertes_thermiques(choix_climat, choix_isolation) * volume_total / efficacite_echangeur) * (T_digesteur - T_in)
            temperature_sortie_ech = T_out    ' report de la valeur pour l'écriture
        End If
    ElseIf choix_chauffage_boues = CHAUFFAGE_BOUES_RECIRCULATION Then
        If temperature_sortie_ech_force = True Then
            T_out = temperature_sortie_ech
        Else
            T_out = T_digesteur + delta_T_boues
            temperature_sortie_ech = T_out    ' report de la valeur pour l'écriture
        End If
    End If


    'Vérification des températures (Tout>Tin)    Ajout DCA 28/02/2013
    If T_out <= T_in Then
        If pn_EffacementResultatsEnCours = False Then
            message_erreur = Feuil26.Range("erreur_digestion_simple_temperatures").Cells(1, 1)
            MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
            pn_EffacementResultatsEnCours = True
            pn_ErreurNonGeree = True
            Exit Sub    'met fin aux calculs
        End If
    End If

    'MCp
    MCp = (energie_maintien_dig_kWhj) / (T_out - T_in)

    'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
    Call copie_dans_flux_thermique_process(nom_flux, T_in, T_out, MCp)

End Sub

Sub bilan_thermique(ByRef besoins_thermiques, ByRef boucle_reel)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'CALCUL DES BESOINS THERMIQUES DU DIGESTEUR
    If energie_maintien_dig_kWhj_force = True And boucle_reel = True Then
        energie_maintien_dig_kWhj = energie_maintien_dig_kWhj_f
    Else
        'pertes themiques
        energie_maintien_dig_kWhj = digesteur_ratio_pertes_thermiques(choix_climat, choix_isolation) * volume_total
        'chauffage des boues
        energie_maintien_dig_kWhj = energie_maintien_dig_kWhj + inlet_Q * Cp_boues * (T_digesteur - T_inlet) / efficacite_echangeur
    End If

    If T_digesteur < T_limite_BT_HT Then
        besoins_thermiques(energie_eau_BT) = besoins_thermiques(energie_eau_BT) + energie_maintien_dig_kWhj
    Else
        besoins_thermiques(energie_eau_HT) = besoins_thermiques(energie_eau_HT) + energie_maintien_dig_kWhj
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "26.11"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub calcul_consommation_electrique(ByRef type_valorisation_biogaz, ByRef valorisation_chaudiere, ByRef ratio_biogaz_valorise)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    Dim ratio_Pw_volume As Double
    Dim tps_fonctionnement As Double
    Dim brassage_HMT As Double
    Dim HMT_pertes As Double
    Dim flux_brassage As Double
    Dim brassage_Q_Nm3j As Double
    Dim ratio_conso_surpresseur As Double
    Dim electricite_agitation As Double

    Dim recirculation_Q As Double
    Dim puissance_pompe As Double
    Dim rendement_pompe As Double
    Dim rendement_moteur As Double
    Dim electricite_recirculation As Double
    Dim electricite_extraction As Double

    Dim SRT_gazometre As Double
    Dim volume_gazometre As Double
    Dim soufflage_Q_Nm3h As Double
    Dim electricite_stockage_biogaz As Double

    Dim electricite_surpresseur As Double

    Dim incoherence As Double



    electricite = 0


    'AGITATION DU DIGESTEUR
    If choix_agitation = agitation_mecanique Then
        ratio_Pw_volume = 7    'W/m3   E. GUIBELIN    modifié le 23/04/2013 par DCA (2W/m3 avant)
        tps_fonctionnement = 24    'h/j
        electricite_agitation = ratio_Pw_volume * tps_fonctionnement * volume_total / 1000
    ElseIf choix_agitation = agitation_recirculation Then
        tps_fonctionnement = 24    'h/j
        HMT_pertes = 10    'mCE
        If tx_recirculation_agitation_force = False Then
            tx_recirculation_agitation = tx_recirculation_agitation
        End If
        recirculation_Q = tx_recirculation_agitation * volume_total / tps_fonctionnement
        rendement_pompe = 0.4    'hypothèse
        puissance_pompe = recirculation_Q / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * HMT_pertes / rendement_pompe
        If puissance_pompe < 60 Then
            rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
        Else
            rendement_moteur = 0.961
        End If
        electricite_agitation = puissance_pompe / rendement_moteur * tps_fonctionnement
    Else  'if choix_agitatio=agitation_biogaz
        flux_brassage = 1.25    'Nm/h   P.JOUAFFRE--> 1 ; Prisma--> entre 1.5 et 2
        tps_fonctionnement = 24    'h/j
        HMT_pertes = 2    'mCE
        brassage_Q_Nm3j = flux_brassage * tps_fonctionnement * volume_total / hauteur_digesteur
        brassage_HMT = hauteur_digesteur + HMT_pertes
        'EN ATTENTE POUR METTRE LES RENDEMENTS DES SURPRESSEURS CHOISIS
        ratio_conso_surpresseur = 4    'Wh/(Nm3.mCE)
        electricite_agitation = ratio_conso_surpresseur * brassage_Q_Nm3j * brassage_HMT / 1000
    End If
    electricite = electricite + electricite_agitation


    'POMPAGE DES BOUES VERS ECHANGEUR
    If choix_chauffage_boues = CHAUFFAGE_BOUES_RECIRCULATION Then
        tps_fonctionnement = 24    'h/j
        HMT_pertes = 10    'mCE
        recirculation_Q = (energie_maintien_dig_kWhj) / (Cp_boues * delta_T_boues) / tps_fonctionnement
        rendement_pompe = 0.4    'hypothèse
        puissance_pompe = recirculation_Q / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * HMT_pertes / rendement_pompe
        If puissance_pompe < 60 Then
            rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
        Else
            rendement_moteur = 0.961
        End If
        electricite_recirculation = puissance_pompe / rendement_moteur * tps_fonctionnement
        electricite = electricite + electricite_recirculation
    End If

    'EXTRACTION DES BOUES
    extraction_Q_unitaire = outlet_Q / (extraction_pompe_nb * extraction_tps_fonctionnement)
    If extraction_Q_unitaire_force = True Then
        extraction_Q_unitaire = extraction_Q_unitaire_f
        If extraction_tps_fonctionnement_force = True Then
            'on ne fait rien
            If extraction_pompe_nb_force = False Then
                If Int(outlet_Q / (extraction_tps_fonctionnement * extraction_Q_unitaire)) < (outlet_Q / (extraction_tps_fonctionnement * extraction_Q_unitaire)) Then
                    extraction_pompe_nb = Int(outlet_Q / (extraction_tps_fonctionnement * extraction_Q_unitaire)) + 1
                Else
                    extraction_pompe_nb = Int(outlet_Q / (extraction_tps_fonctionnement * extraction_Q_unitaire))
                End If
            End If
        Else
            extraction_tps_fonctionnement = outlet_Q / (extraction_pompe_nb * extraction_Q_unitaire)
            If extraction_tps_fonctionnement > hyp_tps_fct_pompe_extraction Then
                extraction_tps_fonctionnement = hyp_tps_fct_pompe_extraction
                If extraction_pompe_nb_force = False Then
                    If Int(outlet_Q / (extraction_tps_fonctionnement * extraction_Q_unitaire)) < (outlet_Q / (extraction_tps_fonctionnement * extraction_Q_unitaire)) Then
                        extraction_pompe_nb = Int(outlet_Q / (extraction_tps_fonctionnement * extraction_Q_unitaire)) + 1
                    Else
                        extraction_pompe_nb = Int(outlet_Q / (extraction_tps_fonctionnement * extraction_Q_unitaire))
                    End If
                    extraction_tps_fonctionnement = outlet_Q / (extraction_pompe_nb * extraction_Q_unitaire)
                End If
            End If
        End If
    Else
        extraction_Q_unitaire = outlet_Q / (extraction_pompe_nb * extraction_tps_fonctionnement)
    End If
    incoherence = (extraction_Q_unitaire * extraction_pompe_nb * extraction_tps_fonctionnement - outlet_Q) / outlet_Q
    If Abs(incoherence) > critere_incoherence Then
        remarque(3) = Feuil26.Range("dig_simple_erreur_pompage").Cells(1, 1)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt = extraction_pompe_rdt_f
    Else
        rendement_pompe = 0.4    'hypothèse
        puissance_pompe = extraction_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * extraction_P_refoulement / rendement_pompe
        If puissance_pompe < 60 Then
            rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
        Else
            rendement_moteur = 0.961
        End If
        extraction_pompe_rdt = rendement_pompe * rendement_moteur
    End If
    electricite_extraction = extraction_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * extraction_P_refoulement / extraction_pompe_rdt * extraction_pompe_nb * extraction_tps_fonctionnement
    electricite = electricite + electricite_extraction


    'STOCKAGE BIOGAZ
    If choix_biogaz_stock = biogaz_stock_oui Then
        If type_valorisation_biogaz = valorisation_chaudiere Then
            SRT_gazometre = 3    '3h de temps de séjour pour chaudière
        Else
            SRT_gazometre = 8    '8h pour autre type de valorisation
        End If
        volume_gazometre = biogaz_Q_Nm3j_nominal / pd_NOMBRE_HEURE_PAR_JOUR * SRT_gazometre
        If volume_gazometre < 1225 Then  '1225 m3
            soufflage_Q_Nm3h = 300  'Nm3/h
        Else
            soufflage_Q_Nm3h = 500   'Nm3/h
        End If
        HMT_pertes = 0.25    'mCE
        ratio_conso_surpresseur = 4    'Wh/(Nm3.mCE)
        electricite_stockage_biogaz = ratio_conso_surpresseur * soufflage_Q_Nm3h * HMT_pertes / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
        electricite = electricite + electricite_stockage_biogaz
    End If


    'ALIMENTATION CHAUDIERE,MICROTURBINE,COGENERATION
    soufflage_Q_Nm3h = biogaz_Q_Nm3j * ratio_biogaz_valorise / pd_NOMBRE_HEURE_PAR_JOUR
    If type_valorisation_biogaz = valorisation_chaudiere Then
        HMT_pertes = 1    '1 mCE si chaudière
    Else
        HMT_pertes = 3    '3 mCE sinon
    End If
    ratio_conso_surpresseur = 4    'Wh/(Nm3.mCE)
    electricite_surpresseur = ratio_conso_surpresseur * soufflage_Q_Nm3h * HMT_pertes / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
    electricite = electricite + electricite_surpresseur

    electricite_fixe = electricite_agitation + electricite_recirculation

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "26.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 3  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil26.Range("dig_simple_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "26.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub dimensionnement(ByRef boues_flux, ByVal choix_eaux_sales_1aire, ByVal boues_I_detail, ByVal boues_II_detail, ByVal boues_III_detail, ByVal graisses_particulaire_ratio_interne, ByRef DCO_biogaz, ByRef boues_pollution_soluble, ByRef besoins_thermiques, ByRef biogaz)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim boucle_reel As Boolean

    Dim intermediaire_MES As Double
    Dim intermediaire_MS As Double
    Dim intermediaire_Q As Double
    Dim intermediaire_MV As Double
    Dim intermediaire_reduction_MV As Double
    Dim inlet_MV As Double
    Dim outlet_MV As Double
    Dim correction_SRT As Double
    Dim ratio_SRT As Double

    Dim inlet_MS As Double

    'gestion des boues de biostyr hors traitement séparé
    Dim coeff_reduction_MV As Double
    Dim ratio_boues_bio As Double

    'gestion des ratios de MV réduits par type de boues
    Dim MV_reduits(nb_type_boues) As Double
    Dim delta_MV_reduits As Double

    'gestion de la compartimentation de la pollution
    Dim dissous As Double
    Dim intermediaire_DCO As Double
    Dim intermediaire_NK As Double
    Dim intermediaire_Pt As Double

    'densité du biogaz
    Dim biogaz_densite As Double

    'co-fermentation
    Dim intermediaire_reduction_MV_soluble As Double
    Dim intermediaire_rdt_MV_coferments As Double

    'soluble passant à travers
    Dim NK_soluble_in_mgL As Double
    Dim Pt_soluble_in_mgL As Double

    Dim incoherence As Double



    'correction par rapport au temps de séjour de référence
    ratio_SRT = SRT_design / SRT_reference
    'If ratio_SRT < ratio_SRT_lim1 Then
    '    correction_SRT = a_ratio_SRT_lim1 * ratio_SRT
    'ElseIf ratio_SRT < ratio_SRT_lim2 Then
    '    correction_SRT = a_ratio_SRT_lim2 * Log(ratio_SRT) + b_ratio_SRT_lim2
    'Else
    '    correction_SRT = ratio_SRT_correction_max
    'End If
    correction_SRT = (a_correction_SRT_gain * ratio_SRT ^ a_correction_SRT_n) / (ratio_SRT ^ a_correction_SRT_n + a_correction_SRT_gain - 1)

    'densité du biogaz
    biogaz_densite = 0
    For I = 0 To 1
        biogaz_densite = biogaz_densite + a_densite_biogaz(I) * biogaz_teneur_CH4 ^ I
    Next I
    ratio_biogaz_MV = 1 / biogaz_densite


    inlet_MES = 0
    inlet_MV = 0
    inlet_MS = 0
    inlet_Q = 0
    outlet_MES = 0
    outlet_MS = 0
    outlet_Q = 0
    outlet_MV = 0
    For I = 1 To nb_type_boues
        MV_reduits(I) = 0
    Next I


    For I = 1 To nb_inlet

        If choix_inlet(I) <> inlet_aucun Then

            If inlet_etape_anterieure(I) = boues_inlet Then

                'variable d'entrée
                intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MES)
                If intermediaire_MES <> 0 Then
                    intermediaire_MV = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MV_MES)
                    'on récupère la matière particulaire entrante pour chaque type de boues (en kg/j), on calculera le ratio en %/MES dans une boucle en aval
                    intermediaire_DCO = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DCO_MES)
                    intermediaire_NK = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_NK_MES)
                    intermediaire_Pt = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_Pt_MES)
                    intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_Q)

                    'pollution soluble
                    NK_soluble_in_mgL = NK_soluble_in_mgL + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_NK) * intermediaire_Q / 1000
                    Pt_soluble_in_mgL = Pt_soluble_in_mgL + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_Pt) * intermediaire_Q / 1000

                    inlet_MES = inlet_MES + intermediaire_MES
                    intermediaire_MS = intermediaire_MES
                    inlet_MS = inlet_MS + intermediaire_MS
                    inlet_MV = inlet_MV + intermediaire_MV
                    inlet_Q = inlet_Q + intermediaire_Q

                    intermediaire_reduction_MV_soluble = 0

                    'REDUCTION DES MATIERES VOLATILES
                    If inlet_type_boues(I) = boues_I And choix_eaux_sales_1aire = True Then
                        ratio_boues_bio = 0
                        coeff_reduction_MV = 0
                        For k = 1 To nb_eaux_sales_max
                            ratio_boues_bio = ratio_boues_bio + boues_I_detail(k, repere_ratio_MV)
                            coeff_reduction_MV = coeff_reduction_MV + boues_I_detail(k, repere_ratio_MV) * reduction_MV_reference(boues_I_detail(k, repere_origine))
                        Next k
                        coeff_reduction_MV = coeff_reduction_MV + (1 - ratio_boues_bio) * reduction_MV_reference(boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_origine))
                        coeff_reduction_MV = coeff_reduction_MV * correction_T * correction_SRT
                        intermediaire_reduction_MV = intermediaire_MV * coeff_reduction_MV
                    ElseIf inlet_type_boues(I) = boues_II And choix_biostyr = True Then
                        coeff_reduction_MV = 0
                        For k = 1 To nb_eaux_sales_max
                            coeff_reduction_MV = coeff_reduction_MV + boues_II_detail(k, repere_ratio_MV) * reduction_MV_reference(boues_II_detail(k, repere_origine))
                        Next k
                        coeff_reduction_MV = coeff_reduction_MV * correction_T * correction_SRT
                        intermediaire_reduction_MV = intermediaire_MV * coeff_reduction_MV
                    ElseIf inlet_type_boues(I) = boues_III And choix_biostyr_N_III = True Then
                        coeff_reduction_MV = 0
                        For k = 1 To nb_eaux_sales_max
                            coeff_reduction_MV = coeff_reduction_MV + boues_III_detail(k, repere_ratio_MV) * reduction_MV_reference(boues_III_detail(k, repere_origine))
                        Next k
                        coeff_reduction_MV = coeff_reduction_MV * correction_T * correction_SRT
                        intermediaire_reduction_MV = intermediaire_MV * coeff_reduction_MV
                    ElseIf inlet_type_boues(I) = graisses And choix_coferments = True Then
                        'graisse interne
                        intermediaire_reduction_MV = reduction_MV_reference(boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_origine)) * correction_T * correction_SRT * intermediaire_MV * graisses_particulaire_ratio_interne(repere_graisse_particulaire_ratioMV)
                        'coferments
                        intermediaire_rdt_MV_coferments = coferments_BMP_facteur_securite * (coferments_BMP_Nm3CH4_tMVapplique / 1000) / (ratio_biogaz_MV * biogaz_teneur_CH4)
                        If intermediaire_rdt_MV_coferments > 1 Then
                            intermediaire_rdt_MV_coferments = 1
                        End If
                        intermediaire_reduction_MV_soluble = intermediaire_rdt_MV_coferments * boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_MV_soluble) * intermediaire_Q / 1000
                        intermediaire_reduction_MV = intermediaire_reduction_MV + intermediaire_reduction_MV_soluble + intermediaire_rdt_MV_coferments * intermediaire_MV * (1 - graisses_particulaire_ratio_interne(repere_graisse_particulaire_ratioMV))
                        'on rajoute la partie soluble pour les MS, les MV, la DCO, le N et le P
                        intermediaire_MS = intermediaire_MS + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_MS_soluble) * intermediaire_Q / 1000
                        inlet_MS = inlet_MS + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_MS_soluble) * intermediaire_Q / 1000
                        intermediaire_MV = intermediaire_MV + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_MV_soluble) * intermediaire_Q / 1000
                        inlet_MV = inlet_MV + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_MV_soluble) * intermediaire_Q / 1000
                        intermediaire_DCO = intermediaire_DCO + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_DCO) * intermediaire_Q / 1000
                        intermediaire_NK = intermediaire_NK + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_NK) * intermediaire_Q / 1000
                        intermediaire_Pt = intermediaire_Pt + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_Pt) * intermediaire_Q / 1000
                    Else
                        intermediaire_reduction_MV = reduction_MV_reference(boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_origine)) * correction_T * correction_SRT * intermediaire_MV
                    End If
                    MV_reduits(inlet_type_boues(I)) = MV_reduits(inlet_type_boues(I)) + intermediaire_reduction_MV
                    intermediaire_MV = intermediaire_MV - intermediaire_reduction_MV
                    intermediaire_MES = intermediaire_MES - (intermediaire_reduction_MV - intermediaire_reduction_MV_soluble)
                    intermediaire_MS = intermediaire_MS - intermediaire_reduction_MV
                    intermediaire_Q = intermediaire_Q - intermediaire_reduction_MV / 1000

                    'variables de sortie
                    outlet_MES = outlet_MES + intermediaire_MES
                    outlet_MS = outlet_MS + intermediaire_MS
                    outlet_MV = outlet_MV + intermediaire_MV
                    outlet_Q = outlet_Q + intermediaire_Q

                    'je calcule les paramètres des boues digérées
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_MV_MES) = (intermediaire_MV + boues_flux(boues_digerees, inlet_type_boues(I), repere_MV_MES) * boues_flux(boues_digerees, inlet_type_boues(I), repere_MES)) / (intermediaire_MES + boues_flux(boues_digerees, inlet_type_boues(I), repere_MES))
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_MES) = boues_flux(boues_digerees, inlet_type_boues(I), repere_MES) + intermediaire_MES
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_ratio_DCO_MES) = boues_flux(boues_digerees, inlet_type_boues(I), repere_ratio_DCO_MES) + intermediaire_DCO
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_ratio_NK_MES) = boues_flux(boues_digerees, inlet_type_boues(I), repere_ratio_NK_MES) + intermediaire_NK
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_ratio_Pt_MES) = boues_flux(boues_digerees, inlet_type_boues(I), repere_ratio_Pt_MES) + intermediaire_Pt
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_Q) = boues_flux(boues_digerees, inlet_type_boues(I), repere_Q) + intermediaire_Q

                    'pour la vérification des flux
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_flux_in) = boues_flux(boues_digerees, inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_digerees, inlet_type_boues(I), repere_flux_in)

                    'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                    boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                End If

            Else

                For j = 1 To nb_type_boues

                    'variables d'entrée
                    intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_MES)
                    If intermediaire_MES <> 0 Then

                        intermediaire_MV = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), j, repere_MV_MES)
                        'on récupère la matière particulaire entrante pour chaque type de boues (en kg/j), on calculera le ratio en %/MES dans une boucle en aval
                        intermediaire_DCO = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DCO_MES)
                        intermediaire_NK = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), j, repere_ratio_NK_MES)
                        intermediaire_Pt = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), j, repere_ratio_Pt_MES)
                        intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_Q)


                        'pollution soluble
                        NK_soluble_in_mgL = NK_soluble_in_mgL + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_NK) * intermediaire_Q / 1000
                        Pt_soluble_in_mgL = Pt_soluble_in_mgL + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_Pt) * intermediaire_Q / 1000


                        inlet_MES = inlet_MES + intermediaire_MES
                        intermediaire_MS = intermediaire_MES
                        inlet_MS = inlet_MS + intermediaire_MS
                        inlet_MV = inlet_MV + intermediaire_MV
                        inlet_Q = inlet_Q + intermediaire_Q

                        intermediaire_reduction_MV_soluble = 0

                        'REDUCTION DES MATIERES VOLATILES
                        If j = boues_I And choix_eaux_sales_1aire = True Then
                            ratio_boues_bio = 0
                            coeff_reduction_MV = 0
                            For k = 1 To nb_eaux_sales_max
                                ratio_boues_bio = ratio_boues_bio + boues_I_detail(k, repere_ratio_MV)
                                coeff_reduction_MV = coeff_reduction_MV + boues_I_detail(k, repere_ratio_MV) * reduction_MV_reference(boues_I_detail(k, repere_origine))
                            Next k
                            coeff_reduction_MV = coeff_reduction_MV + (1 - ratio_boues_bio) * reduction_MV_reference(boues_flux(inlet_etape_anterieure(I), j, repere_origine))
                            coeff_reduction_MV = coeff_reduction_MV * correction_T * correction_SRT
                            intermediaire_reduction_MV = intermediaire_MV * coeff_reduction_MV
                        ElseIf j = boues_II And choix_biostyr = True Then
                            coeff_reduction_MV = 0
                            For k = 1 To nb_eaux_sales_max
                                coeff_reduction_MV = coeff_reduction_MV + boues_II_detail(k, repere_ratio_MV) * reduction_MV_reference(boues_II_detail(k, repere_origine))
                            Next k
                            coeff_reduction_MV = coeff_reduction_MV * correction_T * correction_SRT
                            intermediaire_reduction_MV = intermediaire_MV * coeff_reduction_MV
                        ElseIf j = boues_III And choix_biostyr_N_III = True Then
                            coeff_reduction_MV = 0
                            For k = 1 To nb_eaux_sales_max
                                coeff_reduction_MV = coeff_reduction_MV + boues_III_detail(k, repere_ratio_MV) * reduction_MV_reference(boues_III_detail(k, repere_origine))
                            Next k
                            coeff_reduction_MV = coeff_reduction_MV * correction_T * correction_SRT
                            intermediaire_reduction_MV = intermediaire_MV * coeff_reduction_MV
                        ElseIf j = graisses And choix_coferments = True Then
                            'graisse interne
                            intermediaire_reduction_MV = reduction_MV_reference(boues_flux(inlet_etape_anterieure(I), j, repere_origine)) * correction_T * correction_SRT * intermediaire_MV * graisses_particulaire_ratio_interne(repere_graisse_particulaire_ratioMV)
                            'coferments
                            intermediaire_rdt_MV_coferments = coferments_BMP_facteur_securite * (coferments_BMP_Nm3CH4_tMVapplique / 1000) / (ratio_biogaz_MV * biogaz_teneur_CH4)
                            If intermediaire_rdt_MV_coferments > 1 Then
                                intermediaire_rdt_MV_coferments = 1
                            End If
                            intermediaire_reduction_MV_soluble = intermediaire_rdt_MV_coferments * boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MV_soluble) * intermediaire_Q / 1000
                            intermediaire_reduction_MV = intermediaire_reduction_MV + intermediaire_reduction_MV_soluble + intermediaire_rdt_MV_coferments * intermediaire_MV * (1 - graisses_particulaire_ratio_interne(repere_graisse_particulaire_ratioMV))
                            'on rajoute la partie soluble pour les MV, la DCO, le N et le P
                            intermediaire_MS = intermediaire_MS + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MS_soluble) * intermediaire_Q / 1000
                            inlet_MS = inlet_MS + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MS_soluble) * intermediaire_Q / 1000
                            intermediaire_MV = intermediaire_MV + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MV_soluble) * intermediaire_Q / 1000
                            inlet_MV = inlet_MV + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MV_soluble) * intermediaire_Q / 1000
                            intermediaire_DCO = intermediaire_DCO + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_DCO) * intermediaire_Q / 1000
                            intermediaire_NK = intermediaire_NK + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_NK) * intermediaire_Q / 1000
                            intermediaire_Pt = intermediaire_Pt + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_Pt) * intermediaire_Q / 1000
                        Else
                            intermediaire_reduction_MV = reduction_MV_reference(boues_flux(inlet_etape_anterieure(I), j, repere_origine)) * correction_T * correction_SRT * intermediaire_MV
                        End If
                        MV_reduits(j) = MV_reduits(j) + intermediaire_reduction_MV
                        intermediaire_MV = intermediaire_MV - intermediaire_reduction_MV
                        intermediaire_MES = intermediaire_MES - (intermediaire_reduction_MV - intermediaire_reduction_MV_soluble)
                        intermediaire_MS = intermediaire_MS - intermediaire_reduction_MV
                        intermediaire_Q = intermediaire_Q - intermediaire_reduction_MV / 1000

                        'variables de sortie
                        outlet_MES = outlet_MES + intermediaire_MES
                        outlet_MS = outlet_MS + intermediaire_MS
                        outlet_MV = outlet_MV + intermediaire_MV
                        outlet_Q = outlet_Q + intermediaire_Q

                        'je calcule les paramètres des boues digérées
                        boues_flux(boues_digerees, j, repere_MV_MES) = (intermediaire_MV + boues_flux(boues_digerees, j, repere_MV_MES) * boues_flux(boues_digerees, j, repere_MES)) / (intermediaire_MES + boues_flux(boues_digerees, j, repere_MES))
                        boues_flux(boues_digerees, j, repere_MES) = boues_flux(boues_digerees, j, repere_MES) + intermediaire_MES
                        boues_flux(boues_digerees, j, repere_ratio_DCO_MES) = boues_flux(boues_digerees, j, repere_ratio_DCO_MES) + intermediaire_DCO
                        boues_flux(boues_digerees, j, repere_ratio_NK_MES) = boues_flux(boues_digerees, j, repere_ratio_NK_MES) + intermediaire_NK
                        boues_flux(boues_digerees, j, repere_ratio_Pt_MES) = boues_flux(boues_digerees, j, repere_ratio_Pt_MES) + intermediaire_Pt
                        boues_flux(boues_digerees, j, repere_Q) = boues_flux(boues_digerees, j, repere_Q) + intermediaire_Q


                        boues_flux(boues_digerees, j, repere_flux_in) = boues_flux(boues_digerees, j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                        boues_flux(boues_digerees, j, repere_verif_flux) = boues_flux(boues_digerees, j, repere_flux_in)

                        'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                        boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)

                    End If
                Next j

            End If

        End If

    Next I

    If inlet_MES = 0 Or inlet_Q = 0 Then
        MsgBox Feuil5.Range("erreur_pi_ChoixDigestion_simple_inlet") & " ", Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If

    'calcul de la concentration en entrée
    inlet_concentration = inlet_MS / inlet_Q
    inlet_MV_MES = inlet_MV / inlet_MS

    'calcul du volume du digesteur
    volume_total = inlet_Q * SRT_design
    If volume_total_force = True Then
        incoherence = (volume_total - volume_total_f) / volume_total_f
        If incoherence > critere_incoherence Then
            remarque(1) = Feuil26.Range("dig_simple_erreur_volume").Cells(1, 1)
        End If
        volume_total = volume_total_f
    End If

    'calcul de la proportion de MV dégradés par type de boues en vue de la gestion du changement de taux de réduction global
    For I = 1 To nb_type_boues
        MV_reduits(I) = MV_reduits(I) / (inlet_MV - outlet_MV)
    Next I

    'calcul de la réduction des MV
    reduction_MV = (inlet_MV - outlet_MV) / inlet_MV
    outlet_MV_MES = outlet_MV / outlet_MS
    outlet_concentration = outlet_MS / outlet_Q


    'calcul du biogaz produit
    biogaz_Q_Nm3j = ratio_biogaz_MV * (inlet_MV - outlet_MV)
    biogaz_Q_Nm3j_nominal = biogaz_Q_Nm3j

    'calcul de la DCO du biogaz et ajustement de la DCO des boues digérées par type de boues
    DCO_biogaz = (biogaz_teneur_CH4 * biogaz_Q_Nm3j) / ratio_CH4_DCO
    'on reclacule pour chaque type de boues la DCO des boues digérées
    For I = 1 To nb_type_boues
        boues_flux(boues_digerees, I, repere_ratio_DCO_MES) = boues_flux(boues_digerees, I, repere_ratio_DCO_MES) - MV_reduits(I) * DCO_biogaz
    Next I


    'pollution soluble
    NK_soluble_in_mgL = NK_soluble_in_mgL / outlet_Q * 1000
    Pt_soluble_in_mgL = Pt_soluble_in_mgL / outlet_Q * 1000


    'calcul de la pollution solubilisée lors de la digestion
    For I = 1 To nb_type_boues
        If boues_flux(boues_digerees, I, repere_MES) <> 0 Then
            'MS
            boues_pollution_soluble(boues_digerees, I, repere_mgL_MS_soluble) = (outlet_MS - outlet_MES) / outlet_Q * 1000
            'MV
            boues_pollution_soluble(boues_digerees, I, repere_mgL_MV_soluble) = outlet_MV_MES * boues_pollution_soluble(boues_digerees, I, repere_mgL_MS_soluble)
            'DCO
            dissous = ratio_DCOsol_outlet * boues_flux(boues_digerees, I, repere_ratio_DCO_MES) * correction_T
            boues_pollution_soluble(boues_digerees, I, repere_mgL_DCO) = dissous / boues_flux(boues_digerees, I, repere_Q) * 1000
            boues_flux(boues_digerees, I, repere_ratio_DCO_MES) = (boues_flux(boues_digerees, I, repere_ratio_DCO_MES) - dissous) / boues_flux(boues_digerees, I, repere_MES)
            'DBO
            boues_flux(boues_digerees, I, repere_ratio_DBO_MES) = 0    'hypothèse
            'NK
            dissous = ratio_NK_MVdeg * MV_reduits(I) * (inlet_MV - outlet_MV) * boues_flux(boues_inlet, I, repere_ratio_NK_MES) / boues_flux(boues_inlet, I, repere_MV_MES)
            boues_pollution_soluble(boues_digerees, I, repere_mgL_NK) = NK_soluble_in_mgL + dissous / boues_flux(boues_digerees, I, repere_Q) * 1000
            boues_flux(boues_digerees, I, repere_ratio_NK_MES) = (boues_flux(boues_digerees, I, repere_ratio_NK_MES) - dissous) / boues_flux(boues_digerees, I, repere_MES)
            'Pt
            dissous = boues_flux(boues_digerees, I, repere_ratio_Pt_MES) * ratio_P_Pinlet * correction_T
            boues_pollution_soluble(boues_digerees, I, repere_mgL_Pt) = Pt_soluble_in_mgL + dissous / boues_flux(boues_digerees, I, repere_Q) * 1000
            boues_flux(boues_digerees, I, repere_ratio_Pt_MES) = (boues_flux(boues_digerees, I, repere_ratio_Pt_MES) - dissous) / boues_flux(boues_digerees, I, repere_MES)
        End If
    Next I





    'besoins thermiques pour gestion d'énergie
    boucle_reel = False
    Call bilan_thermique(besoins_thermiques, boucle_reel)
    biogaz(Q_Nm3j) = biogaz_Q_Nm3j
    biogaz(teneur_CH4) = biogaz_teneur_CH4


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "26.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    For Ligne = 1 To nb_inlet
        Feuil26.Range("dig_simple_inlet_ratio").Cells(Ligne, Colonne) = inlet_ratio_admis(Ligne)
    Next Ligne

    Ligne = 1

    Feuil26.Range("dig_simple_inlet_MES").Cells(Ligne, Colonne) = inlet_MS
    Feuil26.Range("dig_simple_inlet_Q").Cells(Ligne, Colonne) = inlet_Q
    Feuil26.Range("dig_simple_inlet_concentration").Cells(Ligne, Colonne) = inlet_concentration
    Feuil26.Range("dig_simple_inlet_MV_MES").Cells(Ligne, Colonne) = inlet_MV_MES
    Feuil26.Range("dig_simple_temperature_digesteur").Cells(Ligne, Colonne) = T_digesteur
    Feuil26.Range("dig_simple_inlet_temperature").Cells(Ligne, Colonne) = T_inlet
    Feuil26.Range("dig_simple_temperature_entree_ech").Cells(Ligne, Colonne) = temperature_entree_ech
    Feuil26.Range("dig_simple_temperature_sortie_ech").Cells(Ligne, Colonne) = temperature_sortie_ech
    Feuil26.Range("dig_simple_temps_sejour_design").Cells(Ligne, Colonne) = SRT_design
    Feuil26.Range("dig_simple_volume_total").Cells(Ligne, Colonne) = volume_total
    Feuil26.Range("dig_simple_temps_sejour_reel").Cells(Ligne, Colonne) = SRT_reel

    If choix_agitation = agitation_biogaz Then
        Feuil26.Range("dig_simple_nb_digesteur").Cells(Ligne, Colonne) = nb_digesteur
        Feuil26.Range("dig_simple_hauteur_digesteur").Cells(Ligne, Colonne) = hauteur_digesteur
    ElseIf choix_agitation = agitation_recirculation Then
        Feuil26.Range("dig_simple_agitation_taux_recirculation").Cells(Ligne, Colonne) = tx_recirculation_agitation
    End If

    Feuil26.Range("dig_simple_outlet_concentration").Cells(Ligne, Colonne) = outlet_concentration
    Feuil26.Range("dig_simple_outlet_MES").Cells(Ligne, Colonne) = outlet_MS
    Feuil26.Range("dig_simple_outlet_Q").Cells(Ligne, Colonne) = outlet_Q
    Feuil26.Range("dig_simple_outlet_MV_MES").Cells(Ligne, Colonne) = outlet_MV_MES
    Feuil26.Range("dig_simple_reduction_MV").Cells(Ligne, Colonne) = reduction_MV
    Feuil26.Range("dig_simple_extraction_pompe_nb").Cells(Ligne, Colonne) = extraction_pompe_nb
    Feuil26.Range("dig_simple_extraction_P_refoulement").Cells(Ligne, Colonne) = extraction_P_refoulement
    Feuil26.Range("dig_simple_extraction_tps_fonctionnement").Cells(Ligne, Colonne) = extraction_tps_fonctionnement
    Feuil26.Range("dig_simple_extraction_Q_unitaire").Cells(Ligne, Colonne) = extraction_Q_unitaire
    Feuil26.Range("dig_simple_extraction_pompe_rdt").Cells(Ligne, Colonne) = extraction_pompe_rdt
    Feuil26.Range("dig_simple_biogaz_teneur_CH4").Cells(Ligne, Colonne) = biogaz_teneur_CH4
    Feuil26.Range("dig_simple_ratio_biogaz_MV").Cells(Ligne, Colonne) = ratio_biogaz_MV
    Feuil26.Range("dig_simple_biogaz_Q_Nm3j").Cells(Ligne, Colonne) = biogaz_Q_Nm3j

    Feuil26.Range("dig_simple_energie_maintien_dig_kWhj").Cells(Ligne, Colonne) = energie_maintien_dig_kWhj
    '    Feuil26.Range("dig_simple_energie_eau90").Cells(ligne, colonne) = energie_eau90
    Feuil26.Range("dig_simple_electricite").Cells(Ligne, Colonne) = electricite


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "26.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    For Ligne = 1 To nb_inlet
        Feuil26.Range("dig_simple_inlet_ratio").Cells(Ligne, Colonne) = ""
    Next Ligne

    Ligne = 1

    Feuil26.Range("dig_simple_inlet_MES").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_inlet_Q").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_inlet_concentration").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_inlet_MV_MES").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_temperature_digesteur").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_inlet_temperature").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_temperature_entree_ech").Cells(Ligne, Colonne).ClearContents
    Feuil26.Range("dig_simple_temperature_sortie_ech").Cells(Ligne, Colonne).ClearContents
    Feuil26.Range("dig_simple_temps_sejour_design").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_volume_total").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_temps_sejour_reel").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_nb_digesteur").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_hauteur_digesteur").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_agitation_taux_recirculation").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_outlet_concentration").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_outlet_MES").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_outlet_Q").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_outlet_MV_MES").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_reduction_MV").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_extraction_pompe_nb").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_extraction_P_refoulement").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_extraction_tps_fonctionnement").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_extraction_Q_unitaire").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_extraction_pompe_rdt").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_biogaz_teneur_CH4").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_ratio_biogaz_MV").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_biogaz_Q_Nm3j").Cells(Ligne, Colonne) = ""
    Feuil26.Range("dig_simple_energie_maintien_dig_kWhj").Cells(Ligne, Colonne) = ""
    '    Feuil26.Range("dig_simple_energie_eau90").Cells(ligne, colonne) = ""
    Feuil26.Range("dig_simple_electricite").Cells(Ligne, Colonne) = ""


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "26.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub fonctionnement_reel(ByRef boues_flux, ByVal choix_eaux_sales_1aire, ByVal boues_I_detail, ByVal boues_II_detail, ByVal boues_III_detail, ByVal graisses_particulaire_ratio_interne, ByRef DCO_biogaz, ByRef boues_pollution_soluble, ByRef besoins_thermiques, ByRef biogaz)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim boucle_reel As Boolean

    Dim intermediaire_MES As Double
    Dim intermediaire_Q As Double
    Dim intermediaire_MV As Double
    Dim intermediaire_reduction_MV As Double
    Dim inlet_MV As Double
    Dim outlet_MV As Double
    Dim correction_SRT As Double
    Dim ratio_SRT As Double

    Dim intermediaire_MS As Double

    'gestion des boues de biostyr hors traitement séparé
    Dim coeff_reduction_MV As Double
    Dim ratio_boues_bio As Double

    'gestion des ratios de MV réduits par type de boues
    Dim MV_reduits(nb_type_boues) As Double
    Dim delta_MV_reduits As Double

    'gestion de la compartimentation de la pollution
    Dim dissous As Double
    Dim intermediaire_DCO As Double
    Dim intermediaire_NK As Double
    Dim intermediaire_Pt As Double

    'densité du biogaz
    Dim biogaz_densite As Double

    'co-fermentation
    Dim intermediaire_reduction_MV_soluble As Double
    Dim intermediaire_rdt_MV_coferments As Double

    'soluble passant à travers
    Dim NK_soluble_in_mgL As Double
    Dim Pt_soluble_in_mgL As Double

    'correction par rapport au temps de séjour de référence
    ratio_SRT = SRT_reel / SRT_reference
    'If ratio_SRT < ratio_SRT_lim1 Then
    '    correction_SRT = a_ratio_SRT_lim1 * ratio_SRT
    'ElseIf ratio_SRT < ratio_SRT_lim2 Then
    '    correction_SRT = a_ratio_SRT_lim2 * Log(ratio_SRT) + b_ratio_SRT_lim2
    'Else
    '    correction_SRT = ratio_SRT_correction_max
    'End If
    correction_SRT = (a_correction_SRT_gain * ratio_SRT ^ a_correction_SRT_n) / (ratio_SRT ^ a_correction_SRT_n + a_correction_SRT_gain - 1)


    'densité du biogaz
    biogaz_densite = 0
    For I = 0 To 1
        biogaz_densite = biogaz_densite + a_densite_biogaz(I) * biogaz_teneur_CH4 ^ I
    Next I
    ratio_biogaz_MV = 1 / biogaz_densite


    inlet_MES = 0
    inlet_MS = 0
    inlet_MV = 0
    inlet_Q = 0
    outlet_MES = 0
    outlet_MS = 0
    outlet_Q = 0
    outlet_MV = 0
    For I = 1 To nb_type_boues
        MV_reduits(I) = 0
    Next I


    For I = 1 To nb_inlet

        If choix_inlet(I) <> inlet_aucun Then

            If inlet_etape_anterieure(I) = boues_inlet Then

                'variable d'entrée
                intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MES)
                If intermediaire_MES <> 0 Then
                    intermediaire_MV = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MV_MES)
                    'on récupère la matière particulaire entrante pour chaque type de boues (en kg/j), on calculera le ratio en %/MES dans une boucle en aval
                    intermediaire_DCO = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DCO_MES)
                    intermediaire_NK = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_NK_MES)
                    intermediaire_Pt = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_Pt_MES)
                    intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_Q)

                    'pollution soluble
                    NK_soluble_in_mgL = NK_soluble_in_mgL + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_NK) * intermediaire_Q / 1000
                    Pt_soluble_in_mgL = Pt_soluble_in_mgL + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_Pt) * intermediaire_Q / 1000

                    inlet_MES = inlet_MES + intermediaire_MES
                    intermediaire_MS = intermediaire_MES
                    inlet_MS = inlet_MS + intermediaire_MS
                    inlet_MV = inlet_MV + intermediaire_MV
                    inlet_Q = inlet_Q + intermediaire_Q

                    intermediaire_reduction_MV_soluble = 0

                    'REDUCTION DES MATIERES VOLATILES
                    If inlet_type_boues(I) = boues_I And choix_eaux_sales_1aire = True Then
                        ratio_boues_bio = 0
                        coeff_reduction_MV = 0
                        For k = 1 To nb_eaux_sales_max
                            ratio_boues_bio = ratio_boues_bio + boues_I_detail(k, repere_ratio_MV)
                            coeff_reduction_MV = coeff_reduction_MV + boues_I_detail(k, repere_ratio_MV) * reduction_MV_reference(boues_I_detail(k, repere_origine))
                        Next k
                        coeff_reduction_MV = coeff_reduction_MV + (1 - ratio_boues_bio) * reduction_MV_reference(boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_origine))
                        coeff_reduction_MV = coeff_reduction_MV * correction_T * correction_SRT
                        intermediaire_reduction_MV = intermediaire_MV * coeff_reduction_MV
                    ElseIf inlet_type_boues(I) = boues_II And choix_biostyr = True Then
                        coeff_reduction_MV = 0
                        For k = 1 To nb_eaux_sales_max
                            coeff_reduction_MV = coeff_reduction_MV + boues_II_detail(k, repere_ratio_MV) * reduction_MV_reference(boues_II_detail(k, repere_origine))
                        Next k
                        coeff_reduction_MV = coeff_reduction_MV * correction_T * correction_SRT
                        intermediaire_reduction_MV = intermediaire_MV * coeff_reduction_MV
                    ElseIf inlet_type_boues(I) = boues_III And choix_biostyr_N_III = True Then
                        coeff_reduction_MV = 0
                        For k = 1 To nb_eaux_sales_max
                            coeff_reduction_MV = coeff_reduction_MV + boues_III_detail(k, repere_ratio_MV) * reduction_MV_reference(boues_III_detail(k, repere_origine))
                        Next k
                        coeff_reduction_MV = coeff_reduction_MV * correction_T * correction_SRT
                        intermediaire_reduction_MV = intermediaire_MV * coeff_reduction_MV
                    ElseIf inlet_type_boues(I) = graisses And choix_coferments = True Then
                        'graisse interne
                        intermediaire_reduction_MV = reduction_MV_reference(boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_origine)) * correction_T * correction_SRT * intermediaire_MV * graisses_particulaire_ratio_interne(repere_graisse_particulaire_ratioMV)
                        'coferments
                        intermediaire_rdt_MV_coferments = coferments_BMP_facteur_securite * (coferments_BMP_Nm3CH4_tMVapplique / 1000) / (ratio_biogaz_MV * biogaz_teneur_CH4)
                        If intermediaire_rdt_MV_coferments > 1 Then
                            intermediaire_rdt_MV_coferments = 1
                        End If
                        intermediaire_reduction_MV_soluble = intermediaire_rdt_MV_coferments * boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_MV_soluble) * intermediaire_Q / 1000
                        intermediaire_reduction_MV = intermediaire_reduction_MV + intermediaire_reduction_MV_soluble + intermediaire_rdt_MV_coferments * intermediaire_MV * (1 - graisses_particulaire_ratio_interne(repere_graisse_particulaire_ratioMV))
                        'on rajoute la partie soluble pour les MS, les MV, la DCO, le N et le P
                        intermediaire_MS = intermediaire_MS + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_MS_soluble) * intermediaire_Q / 1000
                        inlet_MS = inlet_MS + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_MS_soluble) * intermediaire_Q / 1000
                        intermediaire_MV = intermediaire_MV + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_MV_soluble) * intermediaire_Q / 1000
                        inlet_MV = inlet_MV + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_MV_soluble) * intermediaire_Q / 1000
                        intermediaire_DCO = intermediaire_DCO + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_DCO) * intermediaire_Q / 1000
                        intermediaire_NK = intermediaire_NK + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_NK) * intermediaire_Q / 1000
                        intermediaire_Pt = intermediaire_Pt + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_Pt) * intermediaire_Q / 1000
                    Else
                        intermediaire_reduction_MV = reduction_MV_reference(boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_origine)) * correction_T * correction_SRT * intermediaire_MV
                    End If
                    MV_reduits(inlet_type_boues(I)) = MV_reduits(inlet_type_boues(I)) + intermediaire_reduction_MV
                    intermediaire_MV = intermediaire_MV - intermediaire_reduction_MV
                    intermediaire_MES = intermediaire_MES - (intermediaire_reduction_MV - intermediaire_reduction_MV_soluble)
                    intermediaire_MS = intermediaire_MS - intermediaire_reduction_MV
                    intermediaire_Q = intermediaire_Q - intermediaire_reduction_MV / 1000

                    'variables de sortie
                    outlet_MES = outlet_MES + intermediaire_MES
                    outlet_MS = outlet_MS + intermediaire_MS
                    outlet_MV = outlet_MV + intermediaire_MV
                    outlet_Q = outlet_Q + intermediaire_Q

                    'je calcule les paramètres des boues digérées
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_MV_MES) = (intermediaire_MV + boues_flux(boues_digerees, inlet_type_boues(I), repere_MV_MES) * boues_flux(boues_digerees, inlet_type_boues(I), repere_MES)) / (intermediaire_MES + boues_flux(boues_digerees, inlet_type_boues(I), repere_MES))
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_MES) = boues_flux(boues_digerees, inlet_type_boues(I), repere_MES) + intermediaire_MES
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_ratio_DCO_MES) = boues_flux(boues_digerees, inlet_type_boues(I), repere_ratio_DCO_MES) + intermediaire_DCO
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_ratio_NK_MES) = boues_flux(boues_digerees, inlet_type_boues(I), repere_ratio_NK_MES) + intermediaire_NK
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_ratio_Pt_MES) = boues_flux(boues_digerees, inlet_type_boues(I), repere_ratio_Pt_MES) + intermediaire_Pt
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_Q) = boues_flux(boues_digerees, inlet_type_boues(I), repere_Q) + intermediaire_Q

                    'pour la vérification des flux
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_flux_in) = boues_flux(boues_digerees, inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_digerees, inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_digerees, inlet_type_boues(I), repere_flux_in)

                    'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                    boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                End If

            Else

                For j = 1 To nb_type_boues

                    'variables d'entrée
                    intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_MES)
                    If intermediaire_MES <> 0 Then

                        intermediaire_MV = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), j, repere_MV_MES)
                        'on récupère la matière particulaire entrante pour chaque type de boues (en kg/j), on calculera le ratio en %/MES dans une boucle en aval
                        intermediaire_DCO = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DCO_MES)
                        intermediaire_NK = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), j, repere_ratio_NK_MES)
                        intermediaire_Pt = intermediaire_MES * boues_flux(inlet_etape_anterieure(I), j, repere_ratio_Pt_MES)
                        intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_Q)


                        'pollution soluble
                        NK_soluble_in_mgL = NK_soluble_in_mgL + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_NK) * intermediaire_Q / 1000
                        Pt_soluble_in_mgL = Pt_soluble_in_mgL + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_Pt) * intermediaire_Q / 1000


                        inlet_MES = inlet_MES + intermediaire_MES
                        intermediaire_MS = intermediaire_MES
                        inlet_MS = inlet_MS + intermediaire_MS
                        inlet_MV = inlet_MV + intermediaire_MV
                        inlet_Q = inlet_Q + intermediaire_Q

                        intermediaire_reduction_MV_soluble = 0

                        'REDUCTION DES MATIERES VOLATILES
                        If j = boues_I And choix_eaux_sales_1aire = True Then
                            ratio_boues_bio = 0
                            coeff_reduction_MV = 0
                            For k = 1 To nb_eaux_sales_max
                                ratio_boues_bio = ratio_boues_bio + boues_I_detail(k, repere_ratio_MV)
                                coeff_reduction_MV = coeff_reduction_MV + boues_I_detail(k, repere_ratio_MV) * reduction_MV_reference(boues_I_detail(k, repere_origine))
                            Next k
                            coeff_reduction_MV = coeff_reduction_MV + (1 - ratio_boues_bio) * reduction_MV_reference(boues_flux(inlet_etape_anterieure(I), j, repere_origine))
                            coeff_reduction_MV = coeff_reduction_MV * correction_T * correction_SRT
                            intermediaire_reduction_MV = intermediaire_MV * coeff_reduction_MV
                        ElseIf j = boues_II And choix_biostyr = True Then
                            coeff_reduction_MV = 0
                            For k = 1 To nb_eaux_sales_max
                                coeff_reduction_MV = coeff_reduction_MV + boues_II_detail(k, repere_ratio_MV) * reduction_MV_reference(boues_II_detail(k, repere_origine))
                            Next k
                            coeff_reduction_MV = coeff_reduction_MV * correction_T * correction_SRT
                            intermediaire_reduction_MV = intermediaire_MV * coeff_reduction_MV
                        ElseIf j = boues_III And choix_biostyr_N_III = True Then
                            coeff_reduction_MV = 0
                            For k = 1 To nb_eaux_sales_max
                                coeff_reduction_MV = coeff_reduction_MV + boues_III_detail(k, repere_ratio_MV) * reduction_MV_reference(boues_III_detail(k, repere_origine))
                            Next k
                            coeff_reduction_MV = coeff_reduction_MV * correction_T * correction_SRT
                            intermediaire_reduction_MV = intermediaire_MV * coeff_reduction_MV
                        ElseIf j = graisses And choix_coferments = True Then
                            'graisse interne
                            intermediaire_reduction_MV = reduction_MV_reference(boues_flux(inlet_etape_anterieure(I), j, repere_origine)) * correction_T * correction_SRT * intermediaire_MV * graisses_particulaire_ratio_interne(repere_graisse_particulaire_ratioMV)
                            'coferments
                            intermediaire_rdt_MV_coferments = coferments_BMP_facteur_securite * (coferments_BMP_Nm3CH4_tMVapplique / 1000) / (ratio_biogaz_MV * biogaz_teneur_CH4)
                            If intermediaire_rdt_MV_coferments > 1 Then
                                intermediaire_rdt_MV_coferments = 1
                            End If
                            intermediaire_reduction_MV_soluble = intermediaire_rdt_MV_coferments * boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MV_soluble) * intermediaire_Q / 1000
                            intermediaire_reduction_MV = intermediaire_reduction_MV + intermediaire_reduction_MV_soluble + intermediaire_rdt_MV_coferments * intermediaire_MV * (1 - graisses_particulaire_ratio_interne(repere_graisse_particulaire_ratioMV))
                            'on rajoute la partie soluble pour les MV, la DCO, le N et le P
                            intermediaire_MS = intermediaire_MS + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MS_soluble) * intermediaire_Q / 1000
                            inlet_MS = inlet_MS + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MS_soluble) * intermediaire_Q / 1000
                            intermediaire_MV = intermediaire_MV + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MV_soluble) * intermediaire_Q / 1000
                            inlet_MV = inlet_MV + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MV_soluble) * intermediaire_Q / 1000
                            intermediaire_DCO = intermediaire_DCO + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_DCO) * intermediaire_Q / 1000
                            intermediaire_NK = intermediaire_NK + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_NK) * intermediaire_Q / 1000
                            intermediaire_Pt = intermediaire_Pt + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_Pt) * intermediaire_Q / 1000
                        Else
                            intermediaire_reduction_MV = reduction_MV_reference(boues_flux(inlet_etape_anterieure(I), j, repere_origine)) * correction_T * correction_SRT * intermediaire_MV
                        End If
                        MV_reduits(j) = MV_reduits(j) + intermediaire_reduction_MV
                        intermediaire_MV = intermediaire_MV - intermediaire_reduction_MV
                        intermediaire_MES = intermediaire_MES - (intermediaire_reduction_MV - intermediaire_reduction_MV_soluble)
                        intermediaire_MS = intermediaire_MS - intermediaire_reduction_MV
                        intermediaire_Q = intermediaire_Q - intermediaire_reduction_MV / 1000

                        'variables de sortie
                        outlet_MES = outlet_MES + intermediaire_MES
                        outlet_MS = outlet_MS + intermediaire_MS
                        outlet_MV = outlet_MV + intermediaire_MV
                        outlet_Q = outlet_Q + intermediaire_Q

                        'je calcule les paramètres des boues digérées
                        boues_flux(boues_digerees, j, repere_MV_MES) = (intermediaire_MV + boues_flux(boues_digerees, j, repere_MV_MES) * boues_flux(boues_digerees, j, repere_MES)) / (intermediaire_MES + boues_flux(boues_digerees, j, repere_MES))
                        boues_flux(boues_digerees, j, repere_MES) = boues_flux(boues_digerees, j, repere_MES) + intermediaire_MES
                        boues_flux(boues_digerees, j, repere_ratio_DCO_MES) = boues_flux(boues_digerees, j, repere_ratio_DCO_MES) + intermediaire_DCO
                        boues_flux(boues_digerees, j, repere_ratio_NK_MES) = boues_flux(boues_digerees, j, repere_ratio_NK_MES) + intermediaire_NK
                        boues_flux(boues_digerees, j, repere_ratio_Pt_MES) = boues_flux(boues_digerees, j, repere_ratio_Pt_MES) + intermediaire_Pt
                        boues_flux(boues_digerees, j, repere_Q) = boues_flux(boues_digerees, j, repere_Q) + intermediaire_Q
                        boues_flux(boues_digerees, j, repere_flux_in) = boues_flux(boues_digerees, j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                        boues_flux(boues_digerees, j, repere_verif_flux) = boues_flux(boues_digerees, j, repere_flux_in)

                        'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                        boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)

                    End If
                Next j

            End If

        End If

    Next I

    'calcul de la concentration en entrée
    inlet_concentration = inlet_MS / inlet_Q
    inlet_MV_MES = inlet_MV / inlet_MS

    'calcul du temps de séjour réel (j)
    SRT_reel = volume_total / inlet_Q

    'calcul de la proportion de MV dégradés par type de boues en vue de la gestion du changement de taux de réduction global
    For I = 1 To nb_type_boues
        MV_reduits(I) = MV_reduits(I) / (inlet_MV - outlet_MV)
    Next I

    'calcul de la réduction des MV
    reduction_MV = (inlet_MV - outlet_MV) / inlet_MV

    'possible modification de la réduction des MV
    If reduction_MV_force = True Then
        delta_MV_reduits = inlet_MV - outlet_MV   'en kg/j
        reduction_MV = reduction_MV_f
        delta_MV_reduits = reduction_MV * inlet_MV - delta_MV_reduits    'positif si rendement meilleur, négatif sinon
        outlet_MV = outlet_MV - delta_MV_reduits
        outlet_MES = outlet_MES - delta_MV_reduits
        outlet_MS = outlet_MS - delta_MV_reduits
        outlet_Q = outlet_Q - delta_MV_reduits / 1000
        'on recalcule les valeurs des boues en sortie de digestion en respectant la part de réduction par type de boues
        For I = 1 To nb_type_boues
            intermediaire_MV = boues_flux(boues_digerees, I, repere_MV_MES) * boues_flux(boues_digerees, I, repere_MES)
            If intermediaire_MV <> 0 Then
                boues_flux(boues_digerees, I, repere_MES) = boues_flux(boues_digerees, I, repere_MES) - MV_reduits(I) * delta_MV_reduits
                boues_flux(boues_digerees, I, repere_MV_MES) = (intermediaire_MV - MV_reduits(I) * delta_MV_reduits) / boues_flux(boues_digerees, I, repere_MES)
                boues_flux(boues_digerees, I, repere_Q) = boues_flux(boues_digerees, I, repere_Q) - MV_reduits(I) * delta_MV_reduits / 1000
            End If
        Next I
        outlet_MV_MES = outlet_MV / outlet_MS
        outlet_concentration = outlet_MS / outlet_Q
    Else
        outlet_MV_MES = outlet_MV / outlet_MS
        outlet_concentration = outlet_MS / outlet_Q
    End If

    'calcul du biogaz produit
    biogaz_Q_Nm3j = ratio_biogaz_MV * (inlet_MV - outlet_MV)

    'calcul de la DCO du biogaz et ajustement de la DCO des boues digérées par type de boues
    DCO_biogaz = (biogaz_teneur_CH4 * biogaz_Q_Nm3j) / ratio_CH4_DCO
    'on reclacule pour chaque type de boues la DCO des boues digérées
    For I = 1 To nb_type_boues
        boues_flux(boues_digerees, I, repere_ratio_DCO_MES) = boues_flux(boues_digerees, I, repere_ratio_DCO_MES) - MV_reduits(I) * DCO_biogaz
    Next I



    'pollution soluble
    NK_soluble_in_mgL = NK_soluble_in_mgL / outlet_Q * 1000
    Pt_soluble_in_mgL = Pt_soluble_in_mgL / outlet_Q * 1000


    'calcul de la pollution solubilisée lors de la digestion
    For I = 1 To nb_type_boues
        If boues_flux(boues_digerees, I, repere_MES) <> 0 Then
            'MS
            boues_pollution_soluble(boues_digerees, I, repere_mgL_MS_soluble) = (outlet_MS - outlet_MES) / outlet_Q * 1000
            'MV
            boues_pollution_soluble(boues_digerees, I, repere_mgL_MV_soluble) = outlet_MV_MES * boues_pollution_soluble(boues_digerees, I, repere_mgL_MS_soluble)
            'DCO
            dissous = ratio_DCOsol_outlet * boues_flux(boues_digerees, I, repere_ratio_DCO_MES) * correction_T
            boues_pollution_soluble(boues_digerees, I, repere_mgL_DCO) = dissous / boues_flux(boues_digerees, I, repere_Q) * 1000
            boues_flux(boues_digerees, I, repere_ratio_DCO_MES) = (boues_flux(boues_digerees, I, repere_ratio_DCO_MES) - dissous) / boues_flux(boues_digerees, I, repere_MES)
            'DBO
            boues_flux(boues_digerees, I, repere_ratio_DBO_MES) = 0    'hypothèse
            'NK
            dissous = ratio_NK_MVdeg * MV_reduits(I) * (inlet_MV - outlet_MV) * boues_flux(boues_inlet, I, repere_ratio_NK_MES) / boues_flux(boues_inlet, I, repere_MV_MES)
            boues_pollution_soluble(boues_digerees, I, repere_mgL_NK) = NK_soluble_in_mgL + dissous / boues_flux(boues_digerees, I, repere_Q) * 1000
            boues_flux(boues_digerees, I, repere_ratio_NK_MES) = (boues_flux(boues_digerees, I, repere_ratio_NK_MES) - dissous) / boues_flux(boues_digerees, I, repere_MES)
            'Pt
            dissous = boues_flux(boues_digerees, I, repere_ratio_Pt_MES) * ratio_P_Pinlet * correction_T
            boues_pollution_soluble(boues_digerees, I, repere_mgL_Pt) = Pt_soluble_in_mgL + dissous / boues_flux(boues_digerees, I, repere_Q) * 1000
            boues_flux(boues_digerees, I, repere_ratio_Pt_MES) = (boues_flux(boues_digerees, I, repere_ratio_Pt_MES) - dissous) / boues_flux(boues_digerees, I, repere_MES)
        End If
    Next I


    'besoins thermiques pour gestion d'énergie
    boucle_reel = True
    Call bilan_thermique(besoins_thermiques, boucle_reel)
    biogaz(Q_Nm3j) = biogaz_Q_Nm3j
    biogaz(teneur_CH4) = biogaz_teneur_CH4


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "26.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub



Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    'températures pour la digestion
    If T_digesteur_force = True Then
        T_digesteur = T_digesteur_f
    Else
        T_digesteur = T_reference
    End If

    'correctif de température de digestion
    If T_digesteur < 0 Then
        correction_T = 0
    ElseIf T_digesteur < 35 Then
        correction_T = T_digesteur / 35
    ElseIf T_digesteur < 38 Then
        correction_T = 1
    ElseIf T_digesteur < 45 Then
        correction_T = -T_digesteur / 14 + (3 + 5 / 7)
    ElseIf T_digesteur < 55 Then
        correction_T = 5 * T_digesteur / 91 - 1.975
    ElseIf T_digesteur < 58 Then
        correction_T = 1.05
    ElseIf T_digesteur < 80 Then
        correction_T = -T_digesteur / 21 + (3 + 9 / 11)
    Else
        correction_T = 0
    End If

    'température des boues en entrée du digesteur   en °C
    If T_inlet_force = True Then
        T_inlet = T_inlet_f
    Else
        T_inlet = pd_TemperatureEauConditionsReelles_°C
    End If

    'température des boues à l'entrée de l'échangeur
    If choix_chauffage_boues = CHAUFFAGE_BOUES_RECIRCULATION Then
        If temperature_entree_ech_force = True Then
            temperature_entree_ech = temperature_entree_ech_f
        End If
    End If

    'température des boues à la sortie de l'échangeur
    If temperature_sortie_ech_force = True Then
        temperature_sortie_ech = temperature_sortie_ech_f
    End If

    'temps de séjour design
    If SRT_design_force = True Then
        SRT_design = SRT_design_f
    Else
        SRT_design = SRT_reference
    End If

    'agitation du/des digesteur(s)
    If choix_agitation = agitation_biogaz Then
        'nombre de digesteurs
        If nb_digesteur_force = True Then
            nb_digesteur = nb_digesteur_f
        Else
            nb_digesteur = 1
        End If
        'hauteur digesteur
        If hauteur_digesteur_force = True Then
            hauteur_digesteur = hauteur_digesteur_f
        Else
            hauteur_digesteur = 10
        End If
    ElseIf choix_agitation = agitation_recirculation Then
        'taux de recirculation des boues pour agitation    voir pour taux de recirculation référence
        If tx_recirculation_agitation_force = True Then
            tx_recirculation_agitation = tx_recirculation_agitation_f
        Else
            tx_recirculation_agitation = 5    'ici c'est un taux par rapport au volume total de digesteur, le vrai taux sera recalculé dans conso élec
        End If
    End If

    'extraction des boues digérées
    If extraction_pompe_nb_force = True Then
        extraction_pompe_nb = extraction_pompe_nb_f
    Else
        extraction_pompe_nb = 2
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement = extraction_P_refoulement_f
    Else
        extraction_P_refoulement = 50     'mCE
    End If
    If extraction_tps_fonctionnement_force = True Then
        extraction_tps_fonctionnement = extraction_tps_fonctionnement_f
    Else
        extraction_tps_fonctionnement = hyp_tps_fct_pompe_extraction    'h/j par pompe
    End If

    'biogaz %CH4
    If biogaz_teneur_CH4_force = True Then
        biogaz_teneur_CH4 = biogaz_teneur_CH4_f
    Else
        biogaz_teneur_CH4 = 0.63
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "26.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    'réduction des MV référence par type de boues   'modifié par DCA le 31/05/11 d'après une comparaison NDNc+PDN avec dim-digesteur
    'primaire
    reduction_MV_reference(I_simple) = 0.638    '0.61
    reduction_MV_reference(I_reactif) = 0.614  '0.57
    'secondaire
    reduction_MV_reference(II_forte) = 0.546    '0.52
    reduction_MV_reference(II_moyenne) = 0.501    '0.4
    reduction_MV_reference(II_faible_EB) = 0.378    '0.34
    reduction_MV_reference(II_faible_ED) = 0.32   '0.23
    reduction_MV_reference(II_prolongee_EB) = 0.358   ' 0.34
    reduction_MV_reference(II_prolongee_ED) = 0.3  ' 0.23
    reduction_MV_reference(II_MBR) = 0.3    '0.2
    reduction_MV_reference(II_MBBR) = 0.358    '0.25
    reduction_MV_reference(II_HybAS) = 0.358    '0.3
    reduction_MV_reference(II_biostyr_C) = 0.48  '0.45
    reduction_MV_reference(II_biostyr_N) = 0.48  '0.35
    reduction_MV_reference(II_biostyr_NDN) = 0.48  '0.35
    reduction_MV_reference(II_biostyr_PDN) = 0.48  '0.25
    'tertiaire
    reduction_MV_reference(III_decantation) = 0.157    '0.1
    reduction_MV_reference(III_biostyr_N) = 0.157    '0.1
    reduction_MV_reference(III_biostyr_PDN) = 0.157    '0.1
    'graisses
    reduction_MV_reference(codigestion_graisses) = 0.629    '0.6

    'temps de séjour référence par type de digestion    en jours
    If choix_type = mesophile Then
        SRT_reference = 20
    ElseIf choix_type = thermophile Then
        SRT_reference = 12
    End If
    SRT_reel = SRT_reference    'on l'initialise pour la première itération dans le calcul en conditions réelles

    'température référence par type de digestion   en °C
    If choix_type = mesophile Then
        T_reference = 36
    ElseIf choix_type = thermophile Then
        T_reference = 56
    End If

    'correctif par rapport au temps de séjour référence
    'ratio_SRT_lim1 = 0.75
    '    a_ratio_SRT_lim1 = 1.26667
    'ratio_SRT_lim2 = 7
    '    a_ratio_SRT_lim2 = 0.1738
    '    b_ratio_SRT_lim2 = 1
    'ratio_SRT_correction_max = 1.34
    'relation de juillet 2011 d'Eric GUIBELIN    modifié par DCA
    a_correction_SRT_gain = 1.25
    a_correction_SRT_n = 0.7

    'correctif par rapport à la température
    'se reporter à la valeur guide de température

    'calcul de la densité du biogaz en fonction de la fraction de CH4
    a_densite_biogaz(1) = (16 - 44) / 22.4
    a_densite_biogaz(0) = 44 / 22.4

    '0.35 Nm3 de CH4 produit par kg de DCO méthanisée
    ratio_CH4_DCO = 0.35

    'hypothèses sur la solubilisation de la pollution Pt, DCO, NK
    ratio_NK_MVdeg = 0.9  '%NK solubilisé par MV dégradés   source prédimboo
    ratio_P_Pinlet = 0.16    'part du phosphore particulaire solubilisé  source prédimboo
    ratio_DCOsol_outlet = 0.05    'part de DCO soluble en sortie

    'hypothèses sur les pertes du digesteur (en fonction du climat et de l'isolation du digesteur) en kWh/(m3.j)
    digesteur_ratio_pertes_thermiques(climat_froid, isolation_oui) = 0.6
    digesteur_ratio_pertes_thermiques(climat_froid, isolation_non) = 3.72
    digesteur_ratio_pertes_thermiques(climat_tempere, isolation_oui) = 0.36
    digesteur_ratio_pertes_thermiques(climat_tempere, isolation_non) = 2.2
    digesteur_ratio_pertes_thermiques(climat_chaud, isolation_oui) = 0.24
    digesteur_ratio_pertes_thermiques(climat_chaud, isolation_non) = 1.44

    'efficacité de l'échangeur pour chauffer les boues
    efficacite_echangeur = 0.95

    'masse volumique des boues (kg/m3)
    masse_volumique_boues = 1000

    'capacité calorifique des boues   en kWh/(m3.°C)
    Cp_boues = 1.163

    'delta température dans l'échangeur des boues
    delta_T_boues = 3    ' 3°C environ pour la différence entre la T en entrée et en sortie de l'échangeur côté boues

    hyp_tps_fct_pompe_extraction = 14

    'température limite entre mésophile et thermophile
    T_limite_BT_HT = 50   '°C


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "26.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer

    'Boues en entrée de digestion
    For I = 1 To nb_inlet
        choix_inlet(I) = Feuil26.Range("dig_simple_choix_inlet").Cells(I, pi_FiliereConsideree)
        If choix_inlet(I) <> inlet_aucun Then
            inlet_ratio_admis(I) = Feuil26.Range("dig_simple_inlet_ratio").Cells(I, pi_FiliereConsideree)
            If choix_inlet(I) <= inlet_graisses Then
                inlet_etape_anterieure(I) = boues_inlet
                inlet_type_boues(I) = choix_inlet(I)
            Else
                inlet_etape_anterieure(I) = boues_epaissies(choix_inlet(I) - inlet_graisses)
                'ici on prendra directement les valeurs du flux antérieur
            End If
        End If
    Next I

    'type de digestion
    choix_type = Feuil26.Range("dig_simple_choix_type").Cells(1, pi_FiliereConsideree)

    'chauffage des boues
    choix_chauffage_boues = Feuil26.Range("dig_simple_choix_chauffage_boues").Cells(1, pi_FiliereConsideree)

    'agitation
    choix_agitation = Feuil26.Range("dig_simple_choix_agitation").Cells(1, pi_FiliereConsideree)

    'stockage_biogaz
    choix_biogaz_stock = Feuil26.Range("dig_simple_choix_biogaz_stock").Cells(1, pi_FiliereConsideree)

    'isolation
    choix_isolation = Feuil26.Range("dig_simple_choix_isolation").Cells(1, pi_FiliereConsideree)


    'valeurs coferments
    coferments_BMP_Nm3CH4_tMVapplique = Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_BMP_Nm3CH4_tMVapplique, pi_FiliereConsideree)
    coferments_BMP_facteur_securite = Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_BMP_securite, pi_FiliereConsideree)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "26.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    T_digesteur_force = Feuil26.Range("dig_simple_temperature_digesteur").Cells(1, pi_FiliereConsideree)
    T_inlet_force = Feuil26.Range("dig_simple_inlet_temperature").Cells(1, pi_FiliereConsideree)
    'If choix_chauffage_boues = CHAUFFAGE_BOUES_RECIRCULATION Then
    temperature_entree_ech_force = Feuil26.Range("dig_simple_temperature_entree_ech").Cells(1, pi_FiliereConsideree)
    'End If
    temperature_sortie_ech_force = Feuil26.Range("dig_simple_temperature_sortie_ech").Cells(1, pi_FiliereConsideree)
    SRT_design_force = Feuil26.Range("dig_simple_temps_sejour_design").Cells(1, pi_FiliereConsideree)
    volume_total_force = Feuil26.Range("dig_simple_volume_total").Cells(1, pi_FiliereConsideree)

    If choix_agitation = agitation_biogaz Then
        nb_digesteur_force = Feuil26.Range("dig_simple_nb_digesteur").Cells(1, pi_FiliereConsideree)
        hauteur_digesteur_force = Feuil26.Range("dig_simple_hauteur_digesteur").Cells(1, pi_FiliereConsideree)
    ElseIf choix_agitation = agitation_recirculation Then
        tx_recirculation_agitation_force = Feuil26.Range("dig_simple_agitation_taux_recirculation").Cells(1, pi_FiliereConsideree)
    End If

    reduction_MV_force = Feuil26.Range("dig_simple_reduction_MV").Cells(1, pi_FiliereConsideree)
    outlet_concentration_force = Feuil26.Range("dig_simple_outlet_concentration").Cells(1, pi_FiliereConsideree)
    outlet_MES_force = Feuil26.Range("dig_simple_outlet_MES").Cells(1, pi_FiliereConsideree)
    extraction_pompe_nb_force = Feuil26.Range("dig_simple_extraction_pompe_nb").Cells(1, pi_FiliereConsideree)
    extraction_P_refoulement_force = Feuil26.Range("dig_simple_extraction_P_refoulement").Cells(1, pi_FiliereConsideree)
    extraction_tps_fonctionnement_force = Feuil26.Range("dig_simple_extraction_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    extraction_Q_unitaire_force = Feuil26.Range("dig_simple_extraction_Q_unitaire").Cells(1, pi_FiliereConsideree)
    extraction_pompe_rdt_force = Feuil26.Range("dig_simple_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)
    biogaz_teneur_CH4_force = Feuil26.Range("dig_simple_biogaz_teneur_CH4").Cells(1, pi_FiliereConsideree)
    biogaz_Q_Nm3j_force = Feuil26.Range("dig_simple_biogaz_Q_Nm3j").Cells(1, pi_FiliereConsideree)

    energie_maintien_dig_kWhj_force = Feuil26.Range("dig_simple_energie_maintien_dig_kWhj").Cells(1, pi_FiliereConsideree)
    'energie_eau90_force = Feuil26.Range("dig_simple_energie_eau90").Cells(1, pi_FiliereConsideree)

    If T_digesteur_force = True Then
        T_digesteur_f = Feuil26.Range("dig_simple_temperature_digesteur").Cells(1, pi_FiliereConsideree)
    End If
    If T_inlet_force = True Then
        T_inlet_f = Feuil26.Range("dig_simple_inlet_temperature").Cells(1, pi_FiliereConsideree)
    End If
    'If choix_chauffage_boues = CHAUFFAGE_BOUES_RECIRCULATION Then
    If temperature_entree_ech_force = True Then
        temperature_entree_ech_f = Feuil26.Range("dig_simple_temperature_entree_ech").Cells(1, pi_FiliereConsideree)
    End If
    'End If
    If temperature_sortie_ech_force = True Then
        temperature_sortie_ech_f = Feuil26.Range("dig_simple_temperature_sortie_ech").Cells(1, pi_FiliereConsideree)
    End If
    If SRT_design_force = True Then
        SRT_design_f = Feuil26.Range("dig_simple_temps_sejour_design").Cells(1, pi_FiliereConsideree)
    End If
    If volume_total_force = True Then
        volume_total_f = Feuil26.Range("dig_simple_volume_total").Cells(1, pi_FiliereConsideree)
    End If
    If nb_digesteur_force = True Then
        nb_digesteur_f = Feuil26.Range("dig_simple_nb_digesteur").Cells(1, pi_FiliereConsideree)
    End If
    If hauteur_digesteur_force = True Then
        hauteur_digesteur_f = Feuil26.Range("dig_simple_hauteur_digesteur").Cells(1, pi_FiliereConsideree)
    End If
    If tx_recirculation_agitation_force = True Then
        tx_recirculation_agitation_f = Feuil26.Range("dig_simple_agitation_taux_recirculation").Cells(1, pi_FiliereConsideree)
    End If
    If reduction_MV_force = True Then
        reduction_MV_f = Feuil26.Range("dig_simple_reduction_MV").Cells(1, pi_FiliereConsideree)
    End If
    If outlet_concentration_force = True Then
        outlet_concentration_f = Feuil26.Range("dig_simple_outlet_concentration").Cells(1, pi_FiliereConsideree)
    End If
    If outlet_MES_force = True Then
        outlet_MES_f = Feuil26.Range("dig_simple_outlet_MES").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_pompe_nb_force = True Then
        extraction_pompe_nb_f = Feuil26.Range("dig_simple_extraction_pompe_nb").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement_f = Feuil26.Range("dig_simple_extraction_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_tps_fonctionnement_force = True Then
        extraction_tps_fonctionnement_f = Feuil26.Range("dig_simple_extraction_tps_fonctionnement").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_Q_unitaire_force = True Then
        extraction_Q_unitaire_f = Feuil26.Range("dig_simple_extraction_Q_unitaire").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt_f = Feuil26.Range("dig_simple_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    If biogaz_teneur_CH4_force = True Then
        biogaz_teneur_CH4_f = Feuil26.Range("dig_simple_biogaz_teneur_CH4").Cells(1, pi_FiliereConsideree)
    End If
    If biogaz_Q_Nm3j_force = True Then
        biogaz_Q_Nm3j_f = Feuil26.Range("dig_simple_biogaz_Q_Nm3j").Cells(1, pi_FiliereConsideree)
    End If
    If energie_maintien_dig_kWhj_force = True Then
        energie_maintien_dig_kWhj_f = Feuil26.Range("dig_simple_energie_maintien_dig_kWhj").Cells(1, pi_FiliereConsideree)
    End If
    '    If energie_eau90_force = True Then
    '        energie_eau90_f = Feuil26.Range("dig_simple_energie_eau90").Cells(1, pi_FiliereConsideree)
    '    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "26.2"

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







