Attribute VB_Name = "z_Epaississement"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'constantes pour les choix de type d'épaississeur
Private Const gravitaire = 1
Private Const centrifuge = 2
Private Const table_egouttage = 3
Private Const flottation = 4

'constantes pour le choix d'une bâche amont des boues
Private Const bache_oui = 1
Private Const bache_non = 2

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

Private Const nb_inlet = 2

Private Const critere_incoherence = 0.1

'variables de choix
Dim choix_type_epaississeur As Integer
Dim choix_inlet(nb_inlet) As Integer
Dim choix_bache_alimentation As Integer
Public choix_polymere_type As Integer

'variables à échanger avec l'interface
Dim inlet_ratio_admis(nb_inlet) As Double
Dim inlet_MES As Double   'on ne l'affiche plus mais utile pour taux de capture
Dim inlet_MS As Double  'on l'affiche
Dim inlet_Q As Double
Dim inlet_concentration As Double
Dim alimentation_pompe_nb As Integer
Dim alimentation_P_refoulement As Double
Dim alimentation_tps_fonctionnement As Double
Dim alimentation_Q_unitaire As Double
Dim alimentation_pompe_rdt As Double
Dim capture_MES As Double
Public polymere_flux As Double
Dim outlet_concentration As Double
Dim outlet_MES As Double
Dim outlet_MS As Double
Dim outlet_Q As Double
Dim extraction_pompe_nb As Integer
Dim extraction_P_refoulement As Double
Dim extraction_tps_fonctionnement As Double
Dim extraction_Q_unitaire As Double
Dim extraction_pompe_rdt As Double
Public electricite As Double

Public electricite_fixe As Double

'variables à échanger avec l'interface
Dim alimentation_pompe_nb_force As Boolean
Dim alimentation_P_refoulement_force As Boolean
Dim alimentation_tps_fonctionnement_force As Boolean
Dim alimentation_Q_unitaire_force As Boolean
Dim alimentation_pompe_rdt_force As Boolean
Dim capture_MES_force As Boolean
Dim dose_polymere_force As Boolean
Dim outlet_concentration_force As Boolean
Dim extraction_pompe_nb_force As Boolean
Dim extraction_P_refoulement_force As Boolean
Dim extraction_tps_fonctionnement_force As Boolean
Dim extraction_Q_unitaire_force As Boolean
Dim extraction_pompe_rdt_force As Boolean

'variables à échanger avec l'interface
Dim alimentation_pompe_nb_f As Integer
Dim alimentation_P_refoulement_f As Double
Dim alimentation_tps_fonctionnement_f As Double
Dim alimentation_Q_unitaire_f As Double
Dim alimentation_pompe_rdt_f As Double
Dim capture_MES_f As Double
Dim dose_polymere_f As Double
Dim outlet_concentration_f As Double
Dim extraction_pompe_nb_f As Integer
Dim extraction_P_refoulement_f As Double
Dim extraction_tps_fonctionnement_f As Double
Dim extraction_Q_unitaire_f As Double
Dim extraction_pompe_rdt_f As Double

'variables intermédiaires
Dim inlet_etape_anterieure(nb_inlet) As Integer
Dim inlet_type_boues(nb_inlet) As Integer

'hypothèses
Dim dose_polymere As Double
Dim polymere_Q As Double
Dim polymere_concentration As Double
Dim outlet_concentration_reference(nb_qualites_boues) As Double
'Dim retour_concentration(retour_caracteristique_nb) As double
Dim md_ConcentrationBouesEpaissiesAvantBiothelys_gL As Double
Dim md_ConcentrationBouesEpaissiesAvantExelys_gL As Double
Dim hyp_tps_fct_pompe_alim_extraction As Double
Dim md_RatioTauxRecirculationEauBlancheSurConcentrationBouesEntreeFlottateur_L_g As Double
Dim md_RatioDebitAirSurDebitEauBlancheFlottateur As Double
Dim md_PressionRefoulementPompeEauBlanche_mCE As Double
Dim md_PressionAirCompresseFlottateur_mCE As Double
Dim md_ConsommationSpecifiquePompeEauBlanche_Wh_m3mCE As Double
Dim md_ConsommationSpecifiqueCompresseurAir_Wh_Nm3mCE As Double

'erreurs
Dim remarque(4) As String
Dim stockage_numero_epaississement As Integer

Dim Q_traite_nominal As Double
Dim Q_traite_reel As Double
Sub calcul_consommation_electrique()

If version_Developpement = False Then On Error GoTo erreur_non_geree
Dim NumeroErreurNonGeree As String


Dim rendement_pompe As Double
Dim puissance_pompe As Double
Dim rendement_moteur As Double
Dim ratio_conso_spec As Double
Dim ratio_conso_polymere As Double
Dim rate_puissance_fixe As Double
Dim rate_puissance_debit As Double
Dim TauxRecirculationEauBlancheFlottateur As Double
Dim DebitEauBlanche_m3j As Double
Dim DebitAirFlottateur_Nm3d As Double

Dim incoherence As Double

Dim electricite_alimentation As Double
Dim electricite_extraction As Double
Dim electricite_specifique As Double
Dim electricite_polymere As Double



'ALIMENTATION DES BOUES
If choix_bache_alimentation = bache_oui Then
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
            If alimentation_tps_fonctionnement > hyp_tps_fct_pompe_alim_extraction Then
                alimentation_tps_fonctionnement = hyp_tps_fct_pompe_alim_extraction
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
    incoherence = (alimentation_Q_unitaire * alimentation_pompe_nb * alimentation_tps_fonctionnement - inlet_Q) / inlet_Q
    If Abs(incoherence) > critere_incoherence Then
        remarque(2) = Feuil25.Range("epaississement_erreur_pompage_alimentation").Cells(1, 1)
    End If
    If alimentation_pompe_rdt_force = True Then
        alimentation_pompe_rdt = alimentation_pompe_rdt_f
    Else
        rendement_pompe = 0.4    'hypothèse
        puissance_pompe = alimentation_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * alimentation_P_refoulement / rendement_pompe
        If puissance_pompe < 60 Then
            rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
        Else
            rendement_moteur = 0.961
        End If
        alimentation_pompe_rdt = rendement_pompe * rendement_moteur
    End If
    electricite_alimentation = alimentation_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * alimentation_P_refoulement / alimentation_pompe_rdt * alimentation_pompe_nb * alimentation_tps_fonctionnement
End If



'CONSOMMATION SPECIFIQUE PROCEDE
If choix_type_epaississeur = gravitaire Then
    rate_puissance_fixe = 0.1    'Wh/m3 par rapport au nominal
    rate_puissance_debit = 1    'Wh/m3 par rapport au débit réel
    electricite_specifique = 1 / 1000 * (rate_puissance_fixe * Q_traite_nominal + rate_puissance_debit * Q_traite_reel)
ElseIf choix_type_epaississeur = centrifuge Then
    ratio_conso_spec = 0.9  'kWh/m3  Eric Guibelin  (10 mai 2011)
    electricite_specifique = ratio_conso_spec * Q_traite_reel
ElseIf choix_type_epaississeur = table_egouttage Then
    ratio_conso_spec = 4  'kWh/tMS  Eric Guibelin et Pavel Chudoba entre 3 et 5
    electricite_specifique = ratio_conso_spec * outlet_MES / 1000
ElseIf choix_type_epaississeur = flottation Then
    '************** Début modification DCA 05/11/2013 *******************
    'source: Predimboo et E. GUIBELIN
    'Calcul du taux de recirculation et du débit d'eau blanche
    TauxRecirculationEauBlancheFlottateur = md_RatioTauxRecirculationEauBlancheSurConcentrationBouesEntreeFlottateur_L_g * inlet_concentration
    DebitEauBlanche_m3j = TauxRecirculationEauBlancheFlottateur * inlet_Q
    DebitAirFlottateur_Nm3d = md_RatioDebitAirSurDebitEauBlancheFlottateur * DebitEauBlanche_m3j
    electricite_specifique = DebitEauBlanche_m3j * md_PressionRefoulementPompeEauBlanche_mCE * md_ConsommationSpecifiquePompeEauBlanche_Wh_m3mCE / 1000 + DebitAirFlottateur_Nm3d * md_PressionAirCompresseFlottateur_mCE * md_ConsommationSpecifiqueCompresseurAir_Wh_Nm3mCE / 1000
'    ratio_conso_spec = (inlet_concentration - 0.8) / (7 - 0.8) * (90 - 500) + 500    'kWh/tMS  90 pour Neuville
'    electricite_specifique = ratio_conso_spec * outlet_MES / 1000
    '************** Fin modification DCA 05/11/2013 *********************
End If


'CONSOMMATION POUR POLYMERE (préparation+injection)
If choix_type_epaississeur = gravitaire Then
    electricite_polymere = 0
ElseIf choix_type_epaississeur = centrifuge Then
    ratio_conso_polymere = 0.76    'kWh/kg polymère
    electricite_polymere = ratio_conso_polymere * polymere_flux
ElseIf choix_type_epaississeur = table_egouttage Then
    ratio_conso_polymere = 0.76    'kWh/kg polymère
    electricite_polymere = ratio_conso_polymere * polymere_flux
ElseIf choix_type_epaississeur = flottation Then
    ratio_conso_polymere = 0.76    'kWh/kg polymère
    electricite_polymere = ratio_conso_polymere * polymere_flux
End If


'EXTRACTION DES BOUES
'extraction_Q_unitaire = outlet_Q / (extraction_pompe_nb * extraction_tps_fonctionnement)
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
        If extraction_tps_fonctionnement > hyp_tps_fct_pompe_alim_extraction Then
            extraction_tps_fonctionnement = hyp_tps_fct_pompe_alim_extraction
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
    remarque(3) = Feuil25.Range("epaississement_erreur_pompage").Cells(1, 1)
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



electricite = electricite_alimentation + electricite_extraction + electricite_specifique + electricite_polymere
If choix_type_epaississeur = gravitaire Then
    electricite_fixe = electricite_specifique
Else
    electricite_fixe = 0
End If


Exit Sub
'*********************************************************************
erreur_non_geree:
NumeroErreurNonGeree = "25.7"

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 3  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil25.Range("epaississement_erreur_nom") & " " & stockage_numero_epaississement

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "25.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub dimensionnement(ByRef boues_flux, ByRef boues_pollution_soluble, ByVal numero_epaississeur, ByRef retour_flux, ByRef retour_flux_soluble, ByVal boues_I_detail, ByVal boues_II_detail, ByVal boues_III_detail, ByVal epaississement_before_thelys)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer
    Dim Q_retour As Double
    Dim k As Integer


    Dim intermediaire_MS As Double
    Dim intermediaire_MV As Double

    Dim intermediaire_MES As Double
    Dim intermediaire_NK As Double
    Dim intermediaire_Pt As Double
    Dim intermediaire_DCO As Double
    Dim intermediaire_DBO As Double
    Dim intermediaire_Q As Double
    Dim inlet_DCO As Double
    Dim outlet_DCO As Double
    Dim inlet_DBO As Double
    Dim outlet_DBO As Double
    Dim inlet_NK As Double
    Dim outlet_NK As Double
    Dim inlet_Pt As Double
    Dim outlet_Pt As Double
    Dim ratio_DCO As Double
    Dim intermediaire_origine As Integer
    Dim ratio_boues_bio As Double

    Dim soluble_kg(nb_repere_mgL) As Double



    'on vérifie le type d'épaississement
    If epaississement_before_thelys(numero_epaississeur) > 0 And choix_type_epaississeur <> centrifuge Then
        MsgBox Feuil25.Range("epaississement_erreur_nom") & " " & numero_epaississeur & ": " & Feuil25.Range("epaississement_erreur_type_before_thelys").Cells(1, 1), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If



    inlet_MES = 0
    inlet_MS = 0
    inlet_Q = 0
    inlet_DCO = 0
    inlet_DBO = 0
    inlet_NK = 0
    inlet_Pt = 0

    outlet_MES = 0
    outlet_Q = 0
    outlet_DCO = 0
    outlet_DBO = 0
    outlet_NK = 0
    outlet_Pt = 0

    For I = 1 To nb_repere_mgL
        soluble_kg(I) = 0
    Next I


    For I = 1 To nb_inlet

        If choix_inlet(I) <> inlet_aucun Then

            If inlet_etape_anterieure(I) = boues_inlet Then

                'variable d'entrée
                intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MES)
                intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_Q)
                intermediaire_DCO = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DCO_MES) * intermediaire_MES
                intermediaire_DBO = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DBO_MES) * intermediaire_MES
                intermediaire_NK = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_NK_MES) * intermediaire_MES
                intermediaire_Pt = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_Pt_MES) * intermediaire_MES

                inlet_MES = inlet_MES + intermediaire_MES
                inlet_DCO = inlet_DCO + intermediaire_DCO
                inlet_DBO = inlet_DBO + intermediaire_DBO
                inlet_NK = inlet_NK + intermediaire_NK
                inlet_Pt = inlet_Pt + intermediaire_Pt
                inlet_Q = inlet_Q + intermediaire_Q

                'ENTREE SOLUBLE
                'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                For k = 1 To nb_repere_mgL
                    soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), k) * intermediaire_Q / 1000
                Next k

                'SEPARATION (on applique un taux de capture et on recalcule la concentration)
                intermediaire_MES = intermediaire_MES * capture_MES
                intermediaire_DCO = capture_MES * intermediaire_DCO    'on considère qu'à ce moment-là, toute la DCO est particulaire donc on peut obtenir ce qui est piégé
                intermediaire_DBO = capture_MES * intermediaire_DBO
                intermediaire_NK = capture_MES * intermediaire_NK
                intermediaire_Pt = capture_MES * intermediaire_Pt
                intermediaire_Q = intermediaire_MES / outlet_concentration

                'variables de sortie
                outlet_MES = outlet_MES + intermediaire_MES
                outlet_DCO = outlet_DCO + intermediaire_DCO
                outlet_DBO = outlet_DBO + intermediaire_DBO
                outlet_NK = outlet_NK + intermediaire_NK
                outlet_Pt = outlet_Pt + intermediaire_Pt
                outlet_Q = outlet_Q + intermediaire_Q

                'je calcule les paramètres des boues épaissies
                boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_MES) = boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_MES) + intermediaire_MES
                boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_Q) = boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_Q) + intermediaire_Q

                'pour la vérification des flux
                boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_flux_in) = boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_flux_in)

                'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)

            Else

                For j = 1 To nb_type_boues

                    'variables d'entrée
                    intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_MES)
                    intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_Q)
                    intermediaire_DCO = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DCO_MES) * intermediaire_MES
                    intermediaire_DBO = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DBO_MES) * intermediaire_MES
                    intermediaire_NK = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_NK_MES) * intermediaire_MES
                    intermediaire_Pt = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_Pt_MES) * intermediaire_MES

                    inlet_MES = inlet_MES + intermediaire_MES
                    inlet_Q = inlet_Q + intermediaire_Q
                    inlet_DCO = inlet_DCO + intermediaire_DCO
                    inlet_DBO = inlet_DBO + intermediaire_DBO
                    inlet_NK = inlet_NK + intermediaire_NK
                    inlet_Pt = inlet_Pt + intermediaire_Pt

                    'ENTREE SOLUBLE
                    'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                    For k = 1 To nb_repere_mgL
                        soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), j, k) * intermediaire_Q / 1000
                    Next k

                    'SEPARATION
                    intermediaire_MES = intermediaire_MES * capture_MES
                    intermediaire_DCO = capture_MES * intermediaire_DCO
                    intermediaire_DBO = capture_MES * intermediaire_DBO
                    intermediaire_NK = capture_MES * intermediaire_NK
                    intermediaire_Pt = capture_MES * intermediaire_Pt
                    intermediaire_Q = intermediaire_MES / outlet_concentration

                    'variables de sortie
                    outlet_MES = outlet_MES + intermediaire_MES
                    outlet_DCO = outlet_DCO + intermediaire_DCO
                    outlet_DBO = outlet_DBO + intermediaire_DBO
                    outlet_NK = outlet_NK + intermediaire_NK
                    outlet_Pt = outlet_Pt + intermediaire_Pt
                    outlet_Q = outlet_Q + intermediaire_Q

                    'je calcule les paramètres des boues épaissies
                    boues_flux(boues_epaissies(numero_epaississeur), j, repere_MES) = boues_flux(boues_epaissies(numero_epaississeur), j, repere_MES) + intermediaire_MES
                    boues_flux(boues_epaissies(numero_epaississeur), j, repere_Q) = boues_flux(boues_epaissies(numero_epaississeur), j, repere_Q) + intermediaire_Q
                    boues_flux(boues_epaissies(numero_epaississeur), j, repere_flux_in) = boues_flux(boues_epaissies(numero_epaississeur), j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_epaissies(numero_epaississeur), j, repere_verif_flux) = boues_flux(boues_epaissies(numero_epaississeur), j, repere_flux_in)

                    'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                    boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)

                Next j

            End If

        End If

    Next I


    inlet_MS = soluble_kg(repere_mgL_MS_soluble) + inlet_MES

    If inlet_MS = 0 Or inlet_Q = 0 Then
        MsgBox Feuil5.Range("erreur_choix_epaississement_inlet") & " " & numero_epaississeur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If



    inlet_concentration = inlet_MS / inlet_Q   'inlet_MES / inlet_Q



    If choix_type_epaississeur <> gravitaire Then
        polymere_flux = dose_polymere / 1000 * inlet_MS
        polymere_Q = polymere_flux / polymere_concentration
    End If

    Q_traite_nominal = inlet_Q + polymere_Q


    'CALCUL DE LA CONCENTRATION EN SORTIE AVEC UNE TELLE REPARTITION DU TYPE DE BOUES
    outlet_concentration = 0
    For I = 1 To nb_type_boues
        If I = boues_I Then
            ratio_boues_bio = 0
            For j = 1 To nb_eaux_sales_max
                ratio_boues_bio = ratio_boues_bio + boues_I_detail(j, repere_ratio_MES)
                intermediaire_origine = boues_I_detail(j, repere_origine)
                intermediaire_MES = boues_I_detail(j, repere_ratio_MES) * boues_flux(boues_epaissies(numero_epaississeur), I, repere_MES)
                outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
            Next j
            intermediaire_MES = boues_flux(boues_epaissies(numero_epaississeur), I, repere_MES) * (1 - ratio_boues_bio)
            intermediaire_origine = boues_flux(boues_epaissies(numero_epaississeur), I, repere_origine)
            outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
        ElseIf I = boues_II And choix_biostyr = True Then
            For j = 1 To nb_eaux_sales_max
                ratio_boues_bio = boues_II_detail(j, repere_ratio_MES)
                intermediaire_origine = boues_II_detail(j, repere_origine)
                intermediaire_MES = boues_II_detail(j, repere_ratio_MES) * boues_flux(boues_epaissies(numero_epaississeur), I, repere_MES)
                outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
            Next j
        ElseIf I = boues_III And choix_biostyr_N_III = True Then
            For j = 1 To nb_eaux_sales_max
                ratio_boues_bio = boues_III_detail(j, repere_ratio_MES)
                intermediaire_origine = boues_III_detail(j, repere_origine)
                intermediaire_MES = boues_III_detail(j, repere_ratio_MES) * boues_flux(boues_epaissies(numero_epaississeur), I, repere_MES)
                outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
            Next j
        Else
            intermediaire_MES = boues_flux(boues_epaissies(numero_epaississeur), I, repere_MES)
            intermediaire_origine = boues_flux(boues_epaissies(numero_epaississeur), I, repere_origine)
            outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
        End If
    Next I
    If outlet_concentration_force = True Then
        outlet_concentration = outlet_concentration_f
    ElseIf epaississement_before_thelys(numero_epaississeur) = pi_CONFIGURATION_BIOTHELYS Then
        outlet_concentration = md_ConcentrationBouesEpaissiesAvantBiothelys_gL
        If outlet_concentration < inlet_concentration Then
            outlet_concentration = inlet_concentration * 1.01
            'outlet_Q = inlet_Q * 0.99
        End If
    ElseIf epaississement_before_thelys(numero_epaississeur) = pi_CONFIGURATION_EXELYS Then
        outlet_concentration = md_ConcentrationBouesEpaissiesAvantExelys_gL
        If outlet_concentration < inlet_concentration Then
            outlet_concentration = inlet_concentration * 1.01
            'outlet_Q = inlet_Q * 0.99
        End If
    Else
        If outlet_concentration < inlet_concentration Then
            outlet_concentration = inlet_concentration * 1.01
            'outlet_Q = inlet_Q * 0.99
        End If
    End If


    'ESTIMATION DES RETOURS EN TETE DE PRIMAIRE
    Q_retour = inlet_Q + polymere_Q - outlet_Q

    retour_flux(repere_ret_Q) = retour_flux(repere_ret_Q) + Q_retour
    retour_flux(repere_ret_MES) = retour_flux(repere_ret_MES) + inlet_MES - outlet_MES
    retour_flux(repere_ret_DCO) = retour_flux(repere_ret_DCO) + inlet_DCO - outlet_DCO
    retour_flux(repere_ret_DBO) = retour_flux(repere_ret_DBO) + inlet_DBO - outlet_DBO
    retour_flux(repere_ret_NK) = retour_flux(repere_ret_NK) + inlet_NK - outlet_NK
    retour_flux(repere_ret_Pt) = retour_flux(repere_ret_Pt) + inlet_Pt - outlet_Pt



    'ajout de la pollution soluble aux retours de tête
    intermediaire_DCO = soluble_kg(repere_mgL_DCO) / (1 + outlet_Q / Q_retour)
    retour_flux(repere_ret_DCO) = retour_flux(repere_ret_DCO) + intermediaire_DCO
    'retour_flux(repere_ret_DBO) = retour_flux(repere_ret_DBO) + ratio_soluble_DBO_DCO * intermediaire_DCO
    retour_flux_soluble(repere_ret_DCO) = retour_flux_soluble(repere_ret_DCO) + intermediaire_DCO
    'retour_flux_soluble(repere_ret_DBO) = retour_flux_soluble(repere_ret_DBO) + ratio_soluble_DBO_DCO * intermediaire_DCO
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_epaissies(numero_epaississeur), I, repere_mgL_DCO) = (soluble_kg(repere_mgL_DCO) - intermediaire_DCO) / outlet_Q * 1000
    Next I
    intermediaire_NK = soluble_kg(repere_mgL_NK) / (1 + outlet_Q / Q_retour)
    retour_flux(repere_ret_NK) = retour_flux(repere_ret_NK) + intermediaire_NK
    'retour_flux(repere_ret_NH4) = retour_flux(repere_ret_NH4) + ratio_soluble_NH4_NK * intermediaire_NK
    retour_flux_soluble(repere_ret_NK) = retour_flux_soluble(repere_ret_NK) + intermediaire_NK
    'retour_flux_soluble(repere_ret_NH4) = retour_flux_soluble(repere_ret_NH4) + ratio_soluble_NH4_NK * intermediaire_NK
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_epaissies(numero_epaississeur), I, repere_mgL_NK) = (soluble_kg(repere_mgL_NK) - intermediaire_NK) / outlet_Q * 1000
    Next I
    intermediaire_Pt = soluble_kg(repere_mgL_Pt) / (1 + outlet_Q / Q_retour)
    retour_flux(repere_ret_Pt) = retour_flux(repere_ret_Pt) + intermediaire_Pt
    retour_flux_soluble(repere_ret_Pt) = retour_flux_soluble(repere_ret_Pt) + intermediaire_Pt
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_epaissies(numero_epaississeur), I, repere_mgL_Pt) = (soluble_kg(repere_mgL_Pt) - intermediaire_Pt) / outlet_Q * 1000
    Next I
    intermediaire_MS = soluble_kg(repere_mgL_MS_soluble) / (1 + outlet_Q / Q_retour)
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_epaissies(numero_epaississeur), I, repere_mgL_MS_soluble) = (soluble_kg(repere_mgL_MS_soluble) - intermediaire_MS) / outlet_Q * 1000
    Next I
    intermediaire_MV = soluble_kg(repere_mgL_MS_soluble) / (1 + outlet_Q / Q_retour)
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_epaissies(numero_epaississeur), I, repere_mgL_MV_soluble) = (soluble_kg(repere_mgL_MV_soluble) - intermediaire_MV) / outlet_Q * 1000
    Next I

    outlet_MS = outlet_MES + (soluble_kg(repere_mgL_MS_soluble) - intermediaire_MS)





    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "25.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub
Sub ecriture_resultats(ByVal numero_epaississeur)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer


    Ligne = numero_epaississeur
    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


    Feuil25.Range("epaississement_inlet1_ratio_admis").Cells(Ligne, Colonne) = inlet_ratio_admis(1)
    Feuil25.Range("epaississement_inlet2_ratio_admis").Cells(Ligne, Colonne) = inlet_ratio_admis(2)
    Feuil25.Range("epaississement_inlet_MES").Cells(Ligne, Colonne) = inlet_MS   'ATTENTION, CE SONT DES MS maintenant
    Feuil25.Range("epaississement_inlet_Q").Cells(Ligne, Colonne) = inlet_Q
    Feuil25.Range("epaississement_inlet_concentration").Cells(Ligne, Colonne) = inlet_concentration
    If choix_bache_alimentation = bache_oui Then
        Feuil25.Range("epaississement_alimentation_pompe_nb").Cells(Ligne, Colonne) = alimentation_pompe_nb
        Feuil25.Range("epaississement_alimentation_P_refoulement").Cells(Ligne, Colonne) = alimentation_P_refoulement
        Feuil25.Range("epaississement_alimentation_tps_fonctionnement").Cells(Ligne, Colonne) = alimentation_tps_fonctionnement
        Feuil25.Range("epaississement_alimentation_Q_unitaire").Cells(Ligne, Colonne) = alimentation_Q_unitaire
        Feuil25.Range("epaississement_alimentation_pompe_rdt").Cells(Ligne, Colonne) = alimentation_pompe_rdt
    End If
    Feuil25.Range("epaississement_capture_MES").Cells(Ligne, Colonne) = capture_MES
    Feuil25.Range("epaississement_polymere_flux").Cells(Ligne, Colonne) = dose_polymere
    Feuil25.Range("epaississement_outlet_concentration").Cells(Ligne, Colonne) = outlet_concentration
    Feuil25.Range("epaississement_outlet_MES").Cells(Ligne, Colonne) = outlet_MS   'ATTENTION, CE SONT DES MS maintenant
    Feuil25.Range("epaississement_outlet_Q").Cells(Ligne, Colonne) = outlet_Q
    Feuil25.Range("epaississement_extraction_pompe_nb").Cells(Ligne, Colonne) = extraction_pompe_nb
    Feuil25.Range("epaississement_extraction_P_refoulement").Cells(Ligne, Colonne) = extraction_P_refoulement
    Feuil25.Range("epaississement_extraction_tps_fonctionnement").Cells(Ligne, Colonne) = extraction_tps_fonctionnement
    Feuil25.Range("epaississement_extraction_Q_unitaire").Cells(Ligne, Colonne) = extraction_Q_unitaire
    Feuil25.Range("epaississement_extraction_pompe_rdt").Cells(Ligne, Colonne) = extraction_pompe_rdt
    Feuil25.Range("epaississement_electricite").Cells(Ligne, Colonne) = electricite


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "25.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Sub effacer_resultats(ByVal numero_epaississeur)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer


    Ligne = numero_epaississeur
    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


    Feuil25.Range("epaississement_inlet1_ratio_admis").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_inlet2_ratio_admis").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_inlet_MES").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_inlet_Q").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_inlet_concentration").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_alimentation_pompe_nb").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_alimentation_P_refoulement").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_alimentation_tps_fonctionnement").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_alimentation_Q_unitaire").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_alimentation_pompe_rdt").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_capture_MES").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_polymere_flux").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_outlet_concentration").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_outlet_MES").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_outlet_Q").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_extraction_pompe_nb").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_extraction_P_refoulement").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_extraction_tps_fonctionnement").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_extraction_Q_unitaire").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_extraction_pompe_rdt").Cells(Ligne, Colonne) = ""
    Feuil25.Range("epaississement_electricite").Cells(Ligne, Colonne) = ""


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "25.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub



Sub fonctionnement_reel(ByRef boues_flux, ByRef boues_pollution_soluble, ByVal numero_epaississeur, ByRef retour_flux, ByRef retour_flux_soluble, ByVal boues_I_detail, ByVal boues_II_detail, ByVal boues_III_detail, ByVal epaississement_before_thelys)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer
    Dim Q_retour As Double
    Dim k As Integer


    Dim intermediaire_MS As Double
    Dim intermediaire_MV As Double

    Dim intermediaire_MES As Double
    Dim intermediaire_NK As Double
    Dim intermediaire_Pt As Double
    Dim intermediaire_DCO As Double
    Dim intermediaire_DBO As Double
    Dim intermediaire_Q As Double
    Dim inlet_DCO As Double
    Dim outlet_DCO As Double
    Dim inlet_DBO As Double
    Dim outlet_DBO As Double
    Dim inlet_NK As Double
    Dim outlet_NK As Double
    Dim inlet_Pt As Double
    Dim outlet_Pt As Double
    Dim ratio_DCO As Double
    Dim intermediaire_origine As Integer
    Dim ratio_boues_bio As Double

    Dim soluble_kg(nb_repere_mgL) As Double


    inlet_MES = 0
    inlet_MS = 0
    inlet_Q = 0
    inlet_DCO = 0
    inlet_DBO = 0
    inlet_NK = 0
    inlet_Pt = 0

    outlet_MES = 0
    outlet_Q = 0
    outlet_DCO = 0
    outlet_DBO = 0
    outlet_NK = 0
    outlet_Pt = 0

    For I = 1 To nb_repere_mgL
        soluble_kg(I) = 0
    Next I


    For I = 1 To nb_inlet

        If choix_inlet(I) <> inlet_aucun Then

            If inlet_etape_anterieure(I) = boues_inlet Then

                'variable d'entrée
                intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MES)
                intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_Q)
                intermediaire_DCO = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DCO_MES) * intermediaire_MES
                intermediaire_DBO = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_DBO_MES) * intermediaire_MES
                intermediaire_NK = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_NK_MES) * intermediaire_MES
                intermediaire_Pt = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_ratio_Pt_MES) * intermediaire_MES

                inlet_MES = inlet_MES + intermediaire_MES
                inlet_DCO = inlet_DCO + intermediaire_DCO
                inlet_DBO = inlet_DBO + intermediaire_DBO
                inlet_NK = inlet_NK + intermediaire_NK
                inlet_Pt = inlet_Pt + intermediaire_Pt
                inlet_Q = inlet_Q + intermediaire_Q

                'ENTREE SOLUBLE
                'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                For k = 1 To nb_repere_mgL
                    soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), k) * intermediaire_Q / 1000
                Next k

                'SEPARATION (on applique un taux de capture et on recalcule la concentration)
                intermediaire_MES = intermediaire_MES * capture_MES
                intermediaire_DCO = capture_MES * intermediaire_DCO    'on considère qu'à ce moment-là, toute la DCO est particulaire donc on peut obtenir ce qui est piégé
                intermediaire_DBO = capture_MES * intermediaire_DBO
                intermediaire_NK = capture_MES * intermediaire_NK
                intermediaire_Pt = capture_MES * intermediaire_Pt
                intermediaire_Q = intermediaire_MES / outlet_concentration

                'variables de sortie
                outlet_MES = outlet_MES + intermediaire_MES
                outlet_DCO = outlet_DCO + intermediaire_DCO
                outlet_DBO = outlet_DBO + intermediaire_DBO
                outlet_NK = outlet_NK + intermediaire_NK
                outlet_Pt = outlet_Pt + intermediaire_Pt
                outlet_Q = outlet_Q + intermediaire_Q

                'je calcule les paramètres des boues épaissies
                boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_MES) = boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_MES) + intermediaire_MES
                boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_Q) = boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_Q) + intermediaire_Q

                'pour la vérification des flux
                boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_flux_in) = boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_epaissies(numero_epaississeur), inlet_type_boues(I), repere_flux_in)

                'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)

            Else

                For j = 1 To nb_type_boues

                    'variables d'entrée
                    intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_MES)
                    intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), j, repere_Q)
                    intermediaire_DCO = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DCO_MES) * intermediaire_MES
                    intermediaire_DBO = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_DBO_MES) * intermediaire_MES
                    intermediaire_NK = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_NK_MES) * intermediaire_MES
                    intermediaire_Pt = boues_flux(inlet_etape_anterieure(I), j, repere_ratio_Pt_MES) * intermediaire_MES

                    inlet_MES = inlet_MES + intermediaire_MES
                    inlet_Q = inlet_Q + intermediaire_Q
                    inlet_DCO = inlet_DCO + intermediaire_DCO
                    inlet_DBO = inlet_DBO + intermediaire_DBO
                    inlet_NK = inlet_NK + intermediaire_NK
                    inlet_Pt = inlet_Pt + intermediaire_Pt

                    'ENTREE SOLUBLE
                    'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                    For k = 1 To nb_repere_mgL
                        soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), j, k) * intermediaire_Q / 1000
                    Next k

                    'SEPARATION
                    intermediaire_MES = intermediaire_MES * capture_MES
                    intermediaire_DCO = capture_MES * intermediaire_DCO
                    intermediaire_DBO = capture_MES * intermediaire_DBO
                    intermediaire_NK = capture_MES * intermediaire_NK
                    intermediaire_Pt = capture_MES * intermediaire_Pt
                    intermediaire_Q = intermediaire_MES / outlet_concentration

                    'variables de sortie
                    outlet_MES = outlet_MES + intermediaire_MES
                    outlet_DCO = outlet_DCO + intermediaire_DCO
                    outlet_DBO = outlet_DBO + intermediaire_DBO
                    outlet_NK = outlet_NK + intermediaire_NK
                    outlet_Pt = outlet_Pt + intermediaire_Pt
                    outlet_Q = outlet_Q + intermediaire_Q

                    'je calcule les paramètres des boues épaissies
                    boues_flux(boues_epaissies(numero_epaississeur), j, repere_MES) = boues_flux(boues_epaissies(numero_epaississeur), j, repere_MES) + intermediaire_MES
                    boues_flux(boues_epaissies(numero_epaississeur), j, repere_Q) = boues_flux(boues_epaissies(numero_epaississeur), j, repere_Q) + intermediaire_Q
                    boues_flux(boues_epaissies(numero_epaississeur), j, repere_flux_in) = boues_flux(boues_epaissies(numero_epaississeur), j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_epaissies(numero_epaississeur), j, repere_verif_flux) = boues_flux(boues_epaissies(numero_epaississeur), j, repere_flux_in)

                    'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                    boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)

                Next j

            End If

        End If

    Next I


    inlet_MS = soluble_kg(repere_mgL_MS_soluble) + inlet_MES


    inlet_concentration = inlet_MS / inlet_Q   'inlet_MES / inlet_Q

    If choix_type_epaississeur <> gravitaire Then
        polymere_flux = dose_polymere / 1000 * inlet_MS
        polymere_Q = polymere_flux / polymere_concentration
    End If

    Q_traite_reel = inlet_Q + polymere_Q


    'CALCUL DE LA CONCENTRATION EN SORTIE AVEC UNE TELLE REPARTITION DU TYPE DE BOUES
    outlet_concentration = 0
    For I = 1 To nb_type_boues
        If I = boues_I Then
            ratio_boues_bio = 0
            For j = 1 To nb_eaux_sales_max
                ratio_boues_bio = ratio_boues_bio + boues_I_detail(j, repere_ratio_MES)
                intermediaire_origine = boues_I_detail(j, repere_origine)
                intermediaire_MES = boues_I_detail(j, repere_ratio_MES) * boues_flux(boues_epaissies(numero_epaississeur), I, repere_MES)
                outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
            Next j
            intermediaire_MES = boues_flux(boues_epaissies(numero_epaississeur), I, repere_MES) * (1 - ratio_boues_bio)
            intermediaire_origine = boues_flux(boues_epaissies(numero_epaississeur), I, repere_origine)
            outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
        ElseIf I = boues_II And choix_biostyr = True Then
            For j = 1 To nb_eaux_sales_max
                ratio_boues_bio = boues_II_detail(j, repere_ratio_MES)
                intermediaire_origine = boues_II_detail(j, repere_origine)
                intermediaire_MES = boues_II_detail(j, repere_ratio_MES) * boues_flux(boues_epaissies(numero_epaississeur), I, repere_MES)
                outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
            Next j
        ElseIf I = boues_III And choix_biostyr_N_III = True Then
            For j = 1 To nb_eaux_sales_max
                ratio_boues_bio = boues_III_detail(j, repere_ratio_MES)
                intermediaire_origine = boues_III_detail(j, repere_origine)
                intermediaire_MES = boues_III_detail(j, repere_ratio_MES) * boues_flux(boues_epaissies(numero_epaississeur), I, repere_MES)
                outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
            Next j
        Else
            intermediaire_MES = boues_flux(boues_epaissies(numero_epaississeur), I, repere_MES)
            intermediaire_origine = boues_flux(boues_epaissies(numero_epaississeur), I, repere_origine)
            outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
        End If
    Next I
    If outlet_concentration_force = True Then
        outlet_concentration = outlet_concentration_f
    ElseIf epaississement_before_thelys(numero_epaississeur) = pi_CONFIGURATION_BIOTHELYS Then
        outlet_concentration = md_ConcentrationBouesEpaissiesAvantBiothelys_gL
        If outlet_concentration < inlet_concentration Then
            outlet_concentration = inlet_concentration * 1.01
            'outlet_Q = inlet_Q * 0.99
        End If
    ElseIf epaississement_before_thelys(numero_epaississeur) = pi_CONFIGURATION_EXELYS Then
        outlet_concentration = md_ConcentrationBouesEpaissiesAvantExelys_gL
        If outlet_concentration < inlet_concentration Then
            outlet_concentration = inlet_concentration * 1.01
            'outlet_Q = inlet_Q * 0.99
        End If
    Else
        If outlet_concentration < inlet_concentration Then
            outlet_concentration = inlet_concentration * 1.01
            'outlet_Q = inlet_Q * 1.01
        End If
    End If



    'ESTIMATION DES RETOURS EN TETE DE PRIMAIRE
    Q_retour = inlet_Q + polymere_Q - outlet_Q

    retour_flux(repere_ret_Q) = retour_flux(repere_ret_Q) + Q_retour
    retour_flux(repere_ret_MES) = retour_flux(repere_ret_MES) + inlet_MES - outlet_MES
    retour_flux(repere_ret_DCO) = retour_flux(repere_ret_DCO) + inlet_DCO - outlet_DCO
    retour_flux(repere_ret_DBO) = retour_flux(repere_ret_DBO) + inlet_DBO - outlet_DBO
    retour_flux(repere_ret_NK) = retour_flux(repere_ret_NK) + inlet_NK - outlet_NK
    retour_flux(repere_ret_Pt) = retour_flux(repere_ret_Pt) + inlet_Pt - outlet_Pt


    'ajout de la pollution soluble aux retours de tête
    intermediaire_DCO = soluble_kg(repere_mgL_DCO) / (1 + outlet_Q / Q_retour)
    retour_flux(repere_ret_DCO) = retour_flux(repere_ret_DCO) + intermediaire_DCO
    'retour_flux(repere_ret_DBO) = retour_flux(repere_ret_DBO) + ratio_soluble_DBO_DCO * intermediaire_DCO
    retour_flux_soluble(repere_ret_DCO) = retour_flux_soluble(repere_ret_DCO) + intermediaire_DCO
    'retour_flux_soluble(repere_ret_DBO) = retour_flux_soluble(repere_ret_DBO) + ratio_soluble_DBO_DCO * intermediaire_DCO
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_epaissies(numero_epaississeur), I, repere_mgL_DCO) = (soluble_kg(repere_mgL_DCO) - intermediaire_DCO) / outlet_Q * 1000
    Next I
    intermediaire_NK = soluble_kg(repere_mgL_NK) / (1 + outlet_Q / Q_retour)
    retour_flux(repere_ret_NK) = retour_flux(repere_ret_NK) + intermediaire_NK
    'retour_flux(repere_ret_NH4) = retour_flux(repere_ret_NH4) + ratio_soluble_NH4_NK * intermediaire_NK
    retour_flux_soluble(repere_ret_NK) = retour_flux_soluble(repere_ret_NK) + intermediaire_NK
    'retour_flux_soluble(repere_ret_NH4) = retour_flux_soluble(repere_ret_NH4) + ratio_soluble_NH4_NK * intermediaire_NK
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_epaissies(numero_epaississeur), I, repere_mgL_NK) = (soluble_kg(repere_mgL_NK) - intermediaire_NK) / outlet_Q * 1000
    Next I
    intermediaire_Pt = soluble_kg(repere_mgL_Pt) / (1 + outlet_Q / Q_retour)
    retour_flux(repere_ret_Pt) = retour_flux(repere_ret_Pt) + intermediaire_Pt
    retour_flux_soluble(repere_ret_Pt) = retour_flux_soluble(repere_ret_Pt) + intermediaire_Pt
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_epaissies(numero_epaississeur), I, repere_mgL_Pt) = (soluble_kg(repere_mgL_Pt) - intermediaire_Pt) / outlet_Q * 1000
    Next I
    intermediaire_MS = soluble_kg(repere_mgL_MS_soluble) / (1 + outlet_Q / Q_retour)
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_epaissies(numero_epaississeur), I, repere_mgL_MS_soluble) = (soluble_kg(repere_mgL_MS_soluble) - intermediaire_MS) / outlet_Q * 1000
    Next I
    intermediaire_MV = soluble_kg(repere_mgL_MS_soluble) / (1 + outlet_Q / Q_retour)
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_epaissies(numero_epaississeur), I, repere_mgL_MV_soluble) = (soluble_kg(repere_mgL_MV_soluble) - intermediaire_MV) / outlet_Q * 1000
    Next I

    outlet_MS = outlet_MES + (soluble_kg(repere_mgL_MS_soluble) - intermediaire_MS)




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "25.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub
Public Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    'alimentation des boues
    If alimentation_pompe_nb_force = True Then
        alimentation_pompe_nb = alimentation_pompe_nb_f
    Else
        alimentation_pompe_nb = 1
    End If
    If alimentation_P_refoulement_force = True Then
        alimentation_P_refoulement = alimentation_P_refoulement_f
    Else
        alimentation_P_refoulement = 10
    End If
    If alimentation_tps_fonctionnement_force = True Then
        alimentation_tps_fonctionnement = alimentation_tps_fonctionnement_f
    Else
        alimentation_tps_fonctionnement = hyp_tps_fct_pompe_alim_extraction
    End If


    'Taux de capture des MES par défaut
    If capture_MES_force = True Then
        capture_MES = capture_MES_f
    Else
        If choix_type_epaississeur = gravitaire Then
            capture_MES = 0.95
        ElseIf choix_type_epaississeur = centrifuge Then
            capture_MES = 0.97   'si polymere
        ElseIf choix_type_epaississeur = table_egouttage Then
            capture_MES = 0.9
        ElseIf choix_type_epaississeur = flottation Then
            capture_MES = 0.98    'si polymere A VERIFIER
        End If
    End If

    'Concentration des boues obtenues en sortie  'g/L
    'ici on initialise simplement à une valeur pas trop délirante
    If choix_type_epaississeur = gravitaire Then
        outlet_concentration = 50
    ElseIf choix_type_epaississeur = centrifuge Then
        outlet_concentration = 40
    ElseIf choix_type_epaississeur = table_egouttage Then
        outlet_concentration = 70
    ElseIf choix_type_epaississeur = flottation Then
        outlet_concentration = 40
    End If

    'extraction des boues
    If extraction_pompe_nb_force = True Then
        extraction_pompe_nb = extraction_pompe_nb_f
    Else
        extraction_pompe_nb = 1
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement = extraction_P_refoulement_f
    Else
        extraction_P_refoulement = 10
    End If
    If extraction_tps_fonctionnement_force = True Then
        extraction_tps_fonctionnement = extraction_tps_fonctionnement_f
    Else
        extraction_tps_fonctionnement = hyp_tps_fct_pompe_alim_extraction
    End If



    'polymère ajouté, le dosage est en kg de m.a. par tonne de MS
    If dose_polymere_force = True Then
        dose_polymere = dose_polymere_f
    Else
        If choix_type_epaississeur = centrifuge Then
            dose_polymere = 2   'kg/tMS   Eric GUIBELIN  (11 mai 2011)
        ElseIf choix_type_epaississeur = table_egouttage Then
            dose_polymere = 5   'kg/tMS   Eric GUIBELIN  (11 mai 2011)
        ElseIf choix_type_epaississeur = flottation Then
            dose_polymere = 2   'kg/tMS
        End If
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "25.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub hypotheses()

If version_Developpement = False Then On Error GoTo erreur_non_geree
Dim NumeroErreurNonGeree As String



'polymère ajouté, le dosage est en kg de m.a. par tonne de MS
If choix_type_epaississeur = centrifuge Then
    polymere_concentration = 1.5  'g/L polymère injecté
ElseIf choix_type_epaississeur = table_egouttage Then
    polymere_concentration = 1.5  'g/L polymère injecté
ElseIf choix_type_epaississeur = flottation Then
    polymere_concentration = 1.5  'g/L polymère injecté
End If

'Concentration attendue avant un biothelys ou un exelys   DCA 28/08/13 vu avec Julien CHAUZY, on a moins pour le Biothelys car on n'arrive pas à mélanger la boues et la vapeur aussi facilement en Batch qu'en continu
md_ConcentrationBouesEpaissiesAvantBiothelys_gL = 160
md_ConcentrationBouesEpaissiesAvantExelys_gL = 230

'concentration espérée en sortie   en g/L
If choix_type_epaississeur = gravitaire Then
    'primaire
    outlet_concentration_reference(I_simple) = 60
    outlet_concentration_reference(I_reactif) = 53
    'secondaire
    outlet_concentration_reference(II_forte) = 45
    outlet_concentration_reference(II_moyenne) = 28
    outlet_concentration_reference(II_faible_EB) = 20
    outlet_concentration_reference(II_faible_ED) = 20
    outlet_concentration_reference(II_prolongee_EB) = 25
    outlet_concentration_reference(II_prolongee_ED) = 25
    outlet_concentration_reference(II_MBR) = 20
    outlet_concentration_reference(II_MBBR) = 20
    outlet_concentration_reference(II_HybAS) = 25
    outlet_concentration_reference(II_biostyr_C) = 55
    outlet_concentration_reference(II_biostyr_N) = 45
    outlet_concentration_reference(II_biostyr_NDN) = 40
    outlet_concentration_reference(II_biostyr_PDN) = 35
    'tertiaire
    outlet_concentration_reference(III_decantation) = 19
    outlet_concentration_reference(III_biostyr_N) = 19
    outlet_concentration_reference(III_biostyr_PDN) = 19
    'graisses
    outlet_concentration_reference(codigestion_graisses) = 5
ElseIf choix_type_epaississeur = centrifuge Then
    'primaire
    outlet_concentration_reference(I_simple) = 71
    outlet_concentration_reference(I_reactif) = 67
    'secondaire
    outlet_concentration_reference(II_forte) = 62
    outlet_concentration_reference(II_moyenne) = 57
    outlet_concentration_reference(II_faible_EB) = 40
    outlet_concentration_reference(II_faible_ED) = 40
    outlet_concentration_reference(II_prolongee_EB) = 50
    outlet_concentration_reference(II_prolongee_ED) = 50
    outlet_concentration_reference(II_MBR) = 39
    outlet_concentration_reference(II_MBBR) = 50
    outlet_concentration_reference(II_HybAS) = 50
    outlet_concentration_reference(II_biostyr_C) = 62
    outlet_concentration_reference(II_biostyr_N) = 57
    outlet_concentration_reference(II_biostyr_NDN) = 52
    outlet_concentration_reference(II_biostyr_PDN) = 48
    'tertiaire
    outlet_concentration_reference(III_decantation) = 34
    outlet_concentration_reference(III_biostyr_N) = 34
    outlet_concentration_reference(III_biostyr_PDN) = 34
    'graisses
    outlet_concentration_reference(codigestion_graisses) = 29
ElseIf choix_type_epaississeur = table_egouttage Then
    'primaire
    outlet_concentration_reference(I_simple) = 75
    outlet_concentration_reference(I_reactif) = 70
    'secondaire
    outlet_concentration_reference(II_forte) = 65
    outlet_concentration_reference(II_moyenne) = 60
    outlet_concentration_reference(II_faible_EB) = 45
    outlet_concentration_reference(II_faible_ED) = 45
    outlet_concentration_reference(II_prolongee_EB) = 53
    outlet_concentration_reference(II_prolongee_ED) = 53
    outlet_concentration_reference(II_MBR) = 41
    outlet_concentration_reference(II_MBBR) = 45
    outlet_concentration_reference(II_HybAS) = 45
    outlet_concentration_reference(II_biostyr_C) = 65
    outlet_concentration_reference(II_biostyr_N) = 60
    outlet_concentration_reference(II_biostyr_NDN) = 55
    outlet_concentration_reference(II_biostyr_PDN) = 50
    'tertiaire
    outlet_concentration_reference(III_decantation) = 36
    outlet_concentration_reference(III_biostyr_N) = 36
    outlet_concentration_reference(III_biostyr_PDN) = 36
    'graisses
    outlet_concentration_reference(codigestion_graisses) = 30
ElseIf choix_type_epaississeur = flottation Then
    'primaire
    outlet_concentration_reference(I_simple) = 55
    outlet_concentration_reference(I_reactif) = 50
    'secondaire
    outlet_concentration_reference(II_forte) = 53
    outlet_concentration_reference(II_moyenne) = 47
    outlet_concentration_reference(II_faible_EB) = 45
    outlet_concentration_reference(II_faible_ED) = 45
    outlet_concentration_reference(II_prolongee_EB) = 45
    outlet_concentration_reference(II_prolongee_ED) = 45
    outlet_concentration_reference(II_MBR) = 38
    outlet_concentration_reference(II_MBBR) = 45
    outlet_concentration_reference(II_HybAS) = 45
    outlet_concentration_reference(II_biostyr_C) = 53
    outlet_concentration_reference(II_biostyr_N) = 48
    outlet_concentration_reference(II_biostyr_NDN) = 48
    outlet_concentration_reference(II_biostyr_PDN) = 45
    'tertiaire
    outlet_concentration_reference(III_decantation) = 30
    outlet_concentration_reference(III_biostyr_N) = 30
    outlet_concentration_reference(III_biostyr_PDN) = 30
    'graisses
    outlet_concentration_reference(codigestion_graisses) = 60
End If

hyp_tps_fct_pompe_alim_extraction = 5

'Ratio taux de recirculation eau blanche par rapport à la concentration des boues en entrée
md_RatioTauxRecirculationEauBlancheSurConcentrationBouesEntreeFlottateur_L_g = 0.4   'Predimboo
md_RatioDebitAirSurDebitEauBlancheFlottateur = 0.14    'calculé pour 5 bars et 80% de saturation de l'eau blanche
md_PressionRefoulementPompeEauBlanche_mCE = 50   '5 bars
md_PressionAirCompresseFlottateur_mCE = md_PressionRefoulementPompeEauBlanche_mCE
md_ConsommationSpecifiquePompeEauBlanche_Wh_m3mCE = 2.725 / 0.6   'Pompe centrifuge
md_ConsommationSpecifiqueCompresseurAir_Wh_Nm3mCE = 2   'Brochure compresseur air Kaeser


Exit Sub
'*********************************************************************
erreur_non_geree:
NumeroErreurNonGeree = "25.3"

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True


End Sub
Sub lecture_choix(ByVal numero_epaississeur)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer

    stockage_numero_epaississement = numero_epaississeur

    choix_type_epaississeur = Feuil25.Range("epaississement_choix_type").Cells(numero_epaississeur, pi_FiliereConsideree)
    choix_bache_alimentation = Feuil25.Range("epaississement_choix_bache_alimentation").Cells(numero_epaississeur, pi_FiliereConsideree)

    choix_polymere_type = Feuil25.Range("epaississement_choix_polymere_type").Cells(numero_epaississeur, pi_FiliereConsideree)

    'données entrée 1
    choix_inlet(1) = Feuil25.Range("epaississement_choix_inlet1").Cells(numero_epaississeur, pi_FiliereConsideree)
    inlet_ratio_admis(1) = Feuil25.Range("epaississement_inlet1_ratio_admis").Cells(numero_epaississeur, pi_FiliereConsideree)

    'données entrée 2
    choix_inlet(2) = Feuil25.Range("epaississement_choix_inlet2").Cells(numero_epaississeur, pi_FiliereConsideree)
    inlet_ratio_admis(2) = Feuil25.Range("epaississement_inlet2_ratio_admis").Cells(numero_epaississeur, pi_FiliereConsideree)


    'on retranscrit le choix réalisé dans la synthèse pour l'adapter à notre matrice
    For I = 1 To nb_inlet
        If choix_inlet(I) <> inlet_aucun Then
            If choix_inlet(I) <= inlet_graisses Then
                inlet_etape_anterieure(I) = boues_inlet
                inlet_type_boues(I) = choix_inlet(I)
            Else
                inlet_etape_anterieure(I) = boues_epaissies(choix_inlet(I) - inlet_graisses)
                'ici on prendra directement les valeurs du flux antérieur
            End If
        End If
    Next I


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "25.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub



Sub lecture_valeurs_forcees(ByVal numero_epaississeur)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    If choix_bache_alimentation = bache_oui Then
        alimentation_pompe_nb_force = Feuil25.Range("epaississement_alimentation_pompe_nb").Cells(numero_epaississeur, pi_FiliereConsideree)
        alimentation_P_refoulement_force = Feuil25.Range("epaississement_alimentation_P_refoulement").Cells(numero_epaississeur, pi_FiliereConsideree)
        alimentation_tps_fonctionnement_force = Feuil25.Range("epaississement_alimentation_tps_fonctionnement").Cells(numero_epaississeur, pi_FiliereConsideree)
        alimentation_Q_unitaire_force = Feuil25.Range("epaississement_alimentation_Q_unitaire").Cells(numero_epaississeur, pi_FiliereConsideree)
        alimentation_pompe_rdt_force = Feuil25.Range("epaississement_alimentation_pompe_rdt").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If
    capture_MES_force = Feuil25.Range("epaississement_capture_MES").Cells(numero_epaississeur, pi_FiliereConsideree)
    dose_polymere_force = Feuil25.Range("epaississement_polymere_flux").Cells(numero_epaississeur, pi_FiliereConsideree)
    outlet_concentration_force = Feuil25.Range("epaississement_outlet_concentration").Cells(numero_epaississeur, pi_FiliereConsideree)
    extraction_pompe_nb_force = Feuil25.Range("epaississement_extraction_pompe_nb").Cells(numero_epaississeur, pi_FiliereConsideree)
    extraction_P_refoulement_force = Feuil25.Range("epaississement_extraction_P_refoulement").Cells(numero_epaississeur, pi_FiliereConsideree)
    extraction_tps_fonctionnement_force = Feuil25.Range("epaississement_extraction_tps_fonctionnement").Cells(numero_epaississeur, pi_FiliereConsideree)
    extraction_Q_unitaire_force = Feuil25.Range("epaississement_extraction_Q_unitaire").Cells(numero_epaississeur, pi_FiliereConsideree)
    extraction_pompe_rdt_force = Feuil25.Range("epaississement_extraction_pompe_rdt").Cells(numero_epaississeur, pi_FiliereConsideree)


    If alimentation_pompe_nb_force = True Then
        alimentation_pompe_nb_f = Feuil25.Range("epaississement_alimentation_pompe_nb").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If
    If alimentation_P_refoulement_force = True Then
        alimentation_P_refoulement_f = Feuil25.Range("epaississement_alimentation_P_refoulement").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If
    If alimentation_tps_fonctionnement_force = True Then
        alimentation_tps_fonctionnement_f = Feuil25.Range("epaississement_alimentation_tps_fonctionnement").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If
    If alimentation_Q_unitaire_force = True Then
        alimentation_Q_unitaire_f = Feuil25.Range("epaississement_alimentation_Q_unitaire").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If
    If alimentation_pompe_rdt_force = True Then
        alimentation_pompe_rdt_f = Feuil25.Range("epaississement_alimentation_pompe_rdt").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If
    If capture_MES_force = True Then
        capture_MES_f = Feuil25.Range("epaississement_capture_MES").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If
    If dose_polymere_force = True Then
        dose_polymere_f = Feuil25.Range("epaississement_polymere_flux").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If
    If outlet_concentration_force = True Then
        outlet_concentration_f = Feuil25.Range("epaississement_outlet_concentration").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If
    If extraction_pompe_nb_force = True Then
        extraction_pompe_nb_f = Feuil25.Range("epaississement_extraction_pompe_nb").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement_f = Feuil25.Range("epaississement_extraction_P_refoulement").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If
    If extraction_tps_fonctionnement_force = True Then
        extraction_tps_fonctionnement_f = Feuil25.Range("epaississement_extraction_tps_fonctionnement").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If
    If extraction_Q_unitaire_force = True Then
        extraction_Q_unitaire_f = Feuil25.Range("epaississement_extraction_Q_unitaire").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt_f = Feuil25.Range("epaississement_extraction_pompe_rdt").Cells(numero_epaississeur, pi_FiliereConsideree)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "25.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Sub Class_Terminate()

    If pn_EffacementResultatsEnCours = False Then
        Call affich_remarque
    End If

End Sub





