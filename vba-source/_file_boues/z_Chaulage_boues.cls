Attribute VB_Name = "z_Chaulage_boues"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Const nb_inlet = 3

'constantes pour les choix de Siccité intermédiaire après déshydratation
Private Const hygiene = 1  'en vue d'hygiéniser les boues
Private Const siccite = 2  'en vue d'augmenter la siccité
Private Const siccite_hygiene = 3   'augmenter la siccité + hygiénisation
Dim choix_objectif As Integer

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

'variables de choix
'Dim choix_performance As Integer
Dim choix_inlet(nb_inlet) As Integer
Private Const conditionnement_chaux_eteinte = 1
Dim choix_conditionnement_chaux As Integer

'entrée
Dim inlet_concentration As Double
Public chaux_eteinte_flux As Double
Public chaux_vive_flux As Double
Dim inlet_MES As Double
Dim inlet_Q As Double
Dim capture_MS As Double
Dim inlet_MV_MES As Double

'guide defaut
Dim conc_req As Double
Dim taux_chaux As Double

'hypothese
Dim ratio_elec As Double
Dim intermediaire_concentration_reference(nb_qualites_boues) As Double
Dim conso_specifique_type As Double

'calcul
Dim outlet_MES As Double
Dim outlet_Q As Double
Dim outlet_concentration As Double
Dim outlet_MV_MES As Double
Dim conso_spec As Double
Dim intermediaire_concentration As Double

'variables à échanger avec l'interface
Dim inlet_ratio_admis(nb_inlet) As Double
Public electricite As Double

'déclaration des variables relatives aux valeurs renseignées (cases bleues)
'entrée
Dim taux_chaux_force As Boolean
Dim conc_req_force As Boolean

'guide defaut
Dim taux_chaux_f As Double
Dim taux_chaux_c As Double
Dim taux_chaux_hygiene As Double
Dim MM_CaO As Double
Dim MM_CaOH2 As Double
Dim conc_req_f As Double
Dim conc_req_c As Double

'variables intermédiaires
Dim inlet_etape_anterieure(nb_inlet) As Integer
Dim inlet_type_boues(nb_inlet) As Double

'erreurs
Dim stockage_numero_chaulage As Integer
Dim remarque(2) As String
Dim correction_concentration_boues_digerees As Double
Dim stockage_inlet_MES As Double
Dim stockage_boues_pdtes As Double

Sub lecture_choix(ByVal numero_chaulage)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer

    stockage_numero_chaulage = numero_chaulage

    'Boues en entrée de séchage


    choix_inlet(1) = Feuil42.Range("chaulage_choix_inlet1").Cells(numero_chaulage, pi_FiliereConsideree)
    inlet_ratio_admis(1) = Feuil42.Range("chaulage_inlet1_ratio_admis").Cells(numero_chaulage, pi_FiliereConsideree)

    choix_inlet(2) = Feuil42.Range("chaulage_choix_inlet2").Cells(numero_chaulage, pi_FiliereConsideree)
    inlet_ratio_admis(2) = Feuil42.Range("chaulage_inlet2_ratio_admis").Cells(numero_chaulage, pi_FiliereConsideree)

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
    choix_objectif = Feuil42.Range("chaulage_choix_objectif").Cells(numero_chaulage, pi_FiliereConsideree)
    choix_conditionnement_chaux = Feuil42.Range("chaulage_choix_chaux").Cells(numero_chaulage, pi_FiliereConsideree)


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "42.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'ratios de consommation électrique et thermique
    ratio_elec = 1   'kWh/tBH   modifié le 10/03/2014 par DCA d'après données d'E. GUIBELIN (mail avec données de Bruxelles): valeur valable pour un malaxeur; pour une gavopompe, le mélange serait moins bon et la consommation électrique plus faible. Avant valeur Eolia 35 kWh/tMS
    conc_req_c = 300    'g/L
    taux_chaux_hygiene = 0.3    '%
    MM_CaOH2 = 74.08   'g/mol
    MM_CaO = 56.08    'g/mol

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "42.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Sub lecture_valeurs_forcees(ByVal numero_chaulage)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim column As Integer

    column = pi_FiliereConsideree

    taux_chaux_force = Feuil42.Range("chaulage_taux_chaux").Cells(numero_chaulage, column)
    conc_req_force = Feuil42.Range("chaulage_conc_req").Cells(numero_chaulage, column)

    If taux_chaux_force = True Then
        taux_chaux_f = Feuil42.Range("chaulage_taux_chaux").Cells(numero_chaulage, column)
    End If
    If conc_req_force = True Then
        conc_req_f = Feuil42.Range("chaulage_conc_req").Cells(numero_chaulage, column)
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "42.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub dimensionnement(ByRef boues_flux, ByVal numero_chaulage, ByRef retour_flux, ByVal boues_I_detail, ByVal boues_II_detail, ByVal boues_III_detail, ByRef boues_pollution_soluble)             ' (ByRef boues_MS, ByRef boues_Q, ByRef boues_MV_MES)

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
    Dim inlet_DCO As Double
    Dim outlet_DCO As Double
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
    'gestion des ratios de MV réduits par type de boues
    Dim MV_reduits(nb_type_boues) As Double
    Dim delta_MV_reduits As Double


    stockage_inlet_MES = inlet_MES
    stockage_boues_pdtes = outlet_MES - inlet_MES

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

                    inlet_MES = inlet_MES + intermediaire_MES
                    inlet_Q = inlet_Q + intermediaire_Q
                    inlet_DCO = inlet_DCO + intermediaire_DCO
                    inlet_NK = inlet_NK + intermediaire_NK
                    inlet_Pt = inlet_Pt + intermediaire_Pt
                    inlet_MV = inlet_MV + intermediaire_MV

                    If stockage_inlet_MES <> 0 Then
                        intermediaire_MES = intermediaire_MES + intermediaire_MES / stockage_inlet_MES * stockage_boues_pdtes
                    End If

                    'variables de sortie
                    outlet_MES = outlet_MES + intermediaire_MES
                    outlet_DCO = outlet_DCO + intermediaire_DCO
                    outlet_NK = outlet_NK + intermediaire_NK
                    outlet_Pt = outlet_Pt + intermediaire_Pt
                    outlet_Q = outlet_Q + intermediaire_Q

                    'je calcule les paramètres des boues chaulees
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MES) = boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MES) + intermediaire_MES
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_Q) = boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_Q) + intermediaire_Q
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MV_MES) = intermediaire_MV / boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MES)

                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_flux_in) = boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_flux_in)

                    'j'affecte les valeurs de pollution soluble
                    For k = 1 To nb_repere_mgL
                        boues_pollution_soluble(boues_chaulees(numero_chaulage), inlet_type_boues(I), k) = boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), k)
                    Next k

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

                        inlet_MES = inlet_MES + intermediaire_MES
                        inlet_Q = inlet_Q + intermediaire_Q
                        inlet_DCO = inlet_DCO + intermediaire_DCO
                        inlet_NK = inlet_NK + intermediaire_NK
                        inlet_Pt = inlet_Pt + intermediaire_Pt
                        inlet_MV = inlet_MV + intermediaire_MV

                        If stockage_inlet_MES <> 0 Then
                            intermediaire_MES = intermediaire_MES + intermediaire_MES / stockage_inlet_MES * stockage_boues_pdtes
                        End If

                        'variables de sortie
                        outlet_MES = outlet_MES + intermediaire_MES
                        outlet_DCO = outlet_DCO + intermediaire_DCO
                        outlet_NK = outlet_NK + intermediaire_NK
                        outlet_Pt = outlet_Pt + intermediaire_Pt
                        outlet_Q = outlet_Q + intermediaire_Q

                        'je calcule les paramètres des boues chaulees
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_MES) = boues_flux(boues_chaulees(numero_chaulage), j, repere_MES) + intermediaire_MES
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_Q) = boues_flux(boues_chaulees(numero_chaulage), j, repere_Q) + intermediaire_Q
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_MV_MES) = intermediaire_MV / boues_flux(boues_chaulees(numero_chaulage), j, repere_MES)
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_chaulees(numero_chaulage), j, repere_MES)
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_chaulees(numero_chaulage), j, repere_MES)
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_chaulees(numero_chaulage), j, repere_MES)

                        boues_flux(boues_chaulees(numero_chaulage), j, repere_flux_in) = boues_flux(boues_chaulees(numero_chaulage), j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_verif_flux) = boues_flux(boues_chaulees(numero_chaulage), j, repere_flux_in)

                        'j'affecte les valeurs de pollution soluble
                        For k = 1 To nb_repere_mgL
                            boues_pollution_soluble(boues_chaulees(numero_chaulage), j, k) = boues_pollution_soluble(inlet_etape_anterieure(I), j, k)
                        Next k

                        'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                        boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                    End If
                Next j
            End If
        End If
    Next I

    If inlet_MES = 0 Or inlet_Q = 0 Then
        MsgBox Feuil5.Range("erreur_choix_chaulage_inlet") & " " & numero_chaulage, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If

    inlet_concentration = inlet_MES / inlet_Q

    'BOUES EN ENTREE
    inlet_MV_MES = inlet_MV / inlet_MES

    conc_req = conc_req_c
    If choix_objectif = hygiene Then
        taux_chaux = taux_chaux_hygiene
    ElseIf choix_objectif = siccite Then
        If inlet_concentration > 300 Then
            taux_chaux = 0
        Else
            If choix_conditionnement_chaux = conditionnement_chaux_eteinte Then
                taux_chaux = (1 - conc_req / inlet_concentration) / ((conc_req / 1000 - 1) * (MM_CaOH2 / MM_CaO))
            Else
                taux_chaux = (1 - conc_req / inlet_concentration) / (conc_req / 1000 - (MM_CaOH2 / MM_CaO))
            End If
        End If
    Else
        If choix_conditionnement_chaux = conditionnement_chaux_eteinte Then
            taux_chaux = (1 - conc_req / inlet_concentration) / ((conc_req / 1000 - 1) * (MM_CaOH2 / MM_CaO))
        Else
            taux_chaux = (1 - conc_req / inlet_concentration) / (conc_req / 1000 - (MM_CaOH2 / MM_CaO))
        End If
        If taux_chaux > taux_chaux_hygiene Then
            taux_chaux = taux_chaux
        Else
            taux_chaux = taux_chaux_hygiene
        End If
    End If

    If choix_conditionnement_chaux = conditionnement_chaux_eteinte Then
        chaux_eteinte_flux = taux_chaux * inlet_MES * MM_CaOH2 / MM_CaO
        outlet_concentration = 1000 * ((chaux_eteinte_flux / inlet_MES) + 1) / (1000 / inlet_concentration + chaux_eteinte_flux / inlet_MES)
        outlet_MES = (inlet_MES * outlet_concentration / 1000) * (1000 / inlet_concentration + chaux_eteinte_flux / inlet_MES)
        outlet_Q = outlet_MES / outlet_concentration
    Else
        chaux_vive_flux = taux_chaux * inlet_MES
        outlet_concentration = 1000 * (((MM_CaOH2 / MM_CaO) * chaux_vive_flux / inlet_MES) + 1) / (1000 / inlet_concentration + chaux_vive_flux / inlet_MES)
        outlet_MES = (inlet_MES * outlet_concentration / 1000) * ((1000 / inlet_concentration) + (chaux_vive_flux / inlet_MES))
        outlet_Q = outlet_MES / outlet_concentration
    End If

    outlet_MV_MES = inlet_MV_MES * inlet_MES / outlet_MES


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "42.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub fonctionnement_reel(ByRef boues_flux, ByVal numero_chaulage, ByVal boues_I_detail, ByVal boues_II_detail, ByVal boues_III_detail, ByRef retour_flux, ByRef boues_pollution_soluble)                         '(ByRef boues_MS, ByRef boues_Q, ByRef boues_MV_MES)

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
    Dim outlet_Q As Double
    Dim inlet_DCO As Double
    Dim outlet_DCO As Double
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
    'gestion des ratios de MV réduits par type de boues
    Dim MV_reduits(nb_type_boues) As Double
    Dim delta_MV_reduits As Double


    stockage_inlet_MES = inlet_MES
    stockage_boues_pdtes = outlet_MES - inlet_MES

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

                    inlet_MES = inlet_MES + intermediaire_MES
                    inlet_Q = inlet_Q + intermediaire_Q
                    inlet_DCO = inlet_DCO + intermediaire_DCO
                    inlet_NK = inlet_NK + intermediaire_NK
                    inlet_Pt = inlet_Pt + intermediaire_Pt
                    inlet_MV = inlet_MV + intermediaire_MV

                    If stockage_inlet_MES <> 0 Then
                        intermediaire_MES = intermediaire_MES + intermediaire_MES / stockage_inlet_MES * stockage_boues_pdtes
                    End If

                    'variables de sortie
                    outlet_MES = outlet_MES + intermediaire_MES
                    outlet_DCO = outlet_DCO + intermediaire_DCO
                    outlet_NK = outlet_NK + intermediaire_NK
                    outlet_Pt = outlet_Pt + intermediaire_Pt
                    outlet_Q = outlet_Q + intermediaire_Q

                    'je calcule les paramètres des boues chaulees
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MES) = boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MES) + intermediaire_MES
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_Q) = boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_Q) + intermediaire_Q
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MV_MES) = intermediaire_MV / boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_MES)

                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_flux_in) = boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_chaulees(numero_chaulage), inlet_type_boues(I), repere_flux_in)

                    'j'affecte les valeurs de pollution soluble
                    For k = 1 To nb_repere_mgL
                        boues_pollution_soluble(boues_chaulees(numero_chaulage), inlet_type_boues(I), k) = boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), k)
                    Next k

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

                        inlet_MES = inlet_MES + intermediaire_MES
                        inlet_Q = inlet_Q + intermediaire_Q
                        inlet_DCO = inlet_DCO + intermediaire_DCO
                        inlet_NK = inlet_NK + intermediaire_NK
                        inlet_Pt = inlet_Pt + intermediaire_Pt
                        inlet_MV = inlet_MV + intermediaire_MV

                        If stockage_inlet_MES <> 0 Then
                            intermediaire_MES = intermediaire_MES + intermediaire_MES / stockage_inlet_MES * stockage_boues_pdtes
                        End If

                        'variables de sortie
                        outlet_MES = outlet_MES + intermediaire_MES
                        outlet_DCO = outlet_DCO + intermediaire_DCO
                        outlet_NK = outlet_NK + intermediaire_NK
                        outlet_Pt = outlet_Pt + intermediaire_Pt
                        outlet_Q = outlet_Q + intermediaire_Q

                        'je calcule les paramètres des boues séchees
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_MES) = boues_flux(boues_chaulees(numero_chaulage), j, repere_MES) + intermediaire_MES
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_Q) = boues_flux(boues_chaulees(numero_chaulage), j, repere_Q) + intermediaire_Q
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_MV_MES) = intermediaire_MV / boues_flux(boues_chaulees(numero_chaulage), j, repere_MES)
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_chaulees(numero_chaulage), j, repere_MES)
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_chaulees(numero_chaulage), j, repere_MES)
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_chaulees(numero_chaulage), j, repere_MES)

                        boues_flux(boues_chaulees(numero_chaulage), j, repere_flux_in) = boues_flux(boues_chaulees(numero_chaulage), j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                        boues_flux(boues_chaulees(numero_chaulage), j, repere_verif_flux) = boues_flux(boues_chaulees(numero_chaulage), j, repere_flux_in)

                        'j'affecte les valeurs de pollution soluble
                        For k = 1 To nb_repere_mgL
                            boues_pollution_soluble(boues_chaulees(numero_chaulage), j, k) = boues_pollution_soluble(inlet_etape_anterieure(I), j, k)
                        Next k

                        'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                        boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                    End If
                Next j
            End If
        End If
    Next I
    inlet_concentration = inlet_MES / inlet_Q

    'BOUES EN ENTREE
    inlet_MV_MES = inlet_MV / inlet_MES

    If conc_req_force = True Then
        conc_req = conc_req_f
    Else
        conc_req = conc_req_c
    End If

    If taux_chaux_force = True Then
        taux_chaux = taux_chaux_f
    Else
        If choix_objectif = hygiene Then
            taux_chaux = taux_chaux_hygiene
        ElseIf choix_objectif = siccite Then
            If inlet_concentration > 300 Then
                taux_chaux = 0
            Else
                If choix_conditionnement_chaux = conditionnement_chaux_eteinte Then
                    taux_chaux = (1 - conc_req / inlet_concentration) / ((conc_req / 1000 - 1) * (MM_CaOH2 / MM_CaO))
                Else
                    taux_chaux = (1 - conc_req / inlet_concentration) / (conc_req / 1000 - (MM_CaOH2 / MM_CaO))
                End If
            End If
        Else
            If choix_conditionnement_chaux = conditionnement_chaux_eteinte Then
                taux_chaux = (1 - conc_req / inlet_concentration) / ((conc_req / 1000 - 1) * (MM_CaOH2 / MM_CaO))
            Else
                taux_chaux = (1 - conc_req / inlet_concentration) / (conc_req / 1000 - (MM_CaOH2 / MM_CaO))
            End If
            If taux_chaux > taux_chaux_hygiene Then
                taux_chaux = taux_chaux
            Else
                taux_chaux = taux_chaux_hygiene
            End If
        End If
    End If

    If choix_conditionnement_chaux = conditionnement_chaux_eteinte Then
        chaux_eteinte_flux = taux_chaux * inlet_MES * MM_CaOH2 / MM_CaO
        outlet_concentration = 1000 * ((chaux_eteinte_flux / inlet_MES) + 1) / (1000 / inlet_concentration + chaux_eteinte_flux / inlet_MES)
        outlet_MES = (inlet_MES * outlet_concentration / 1000) * (1000 / inlet_concentration + chaux_eteinte_flux / inlet_MES)
        outlet_Q = outlet_MES / outlet_concentration
    Else
        chaux_vive_flux = taux_chaux * inlet_MES
        outlet_concentration = 1000 * (((MM_CaOH2 / MM_CaO) * chaux_vive_flux / inlet_MES) + 1) / (1000 / inlet_concentration + chaux_vive_flux / inlet_MES)
        outlet_MES = (inlet_MES * outlet_concentration / 1000) * ((1000 / inlet_concentration) + (chaux_vive_flux / inlet_MES))
        outlet_Q = outlet_MES / outlet_concentration
    End If

    outlet_MV_MES = inlet_MV_MES * inlet_MES / outlet_MES


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "42.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub calcul_consommation_electrique()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim electricite_specifique As Double

    If taux_chaux > 0 Then
        electricite_specifique = inlet_Q * ratio_elec
    Else
        electricite_specifique = 0
    End If

    electricite = electricite_specifique

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "42.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub ecriture_resultats(ByVal numero_chaulage)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Ligne = numero_chaulage

    Feuil42.Range("chaulage_inlet1_ratio_admis").Cells(Ligne, Colonne) = inlet_ratio_admis(1)
    Feuil42.Range("chaulage_inlet2_ratio_admis").Cells(Ligne, Colonne) = inlet_ratio_admis(2)
    Feuil42.Range("chaulage_inlet_MES").Cells(Ligne, Colonne) = inlet_MES
    Feuil42.Range("chaulage_inlet_Q").Cells(Ligne, Colonne) = inlet_Q
    Feuil42.Range("chaulage_inlet_concentration").Cells(Ligne, Colonne) = inlet_concentration
    Feuil42.Range("chaulage_inlet_MV_MES").Cells(Ligne, Colonne) = inlet_MV_MES
    Feuil42.Range("chaulage_conc_req").Cells(Ligne, Colonne) = conc_req
    Feuil42.Range("chaulage_taux_chaux").Cells(Ligne, Colonne) = taux_chaux
    Feuil42.Range("chaulage_quant_chaux").Cells(Ligne, Colonne) = chaux_eteinte_flux + chaux_vive_flux
    Feuil42.Range("chaulage_outlet_concentration").Cells(Ligne, Colonne) = outlet_concentration
    Feuil42.Range("chaulage_outlet_MES").Cells(Ligne, Colonne) = outlet_MES
    Feuil42.Range("chaulage_outlet_Q").Cells(Ligne, Colonne) = outlet_Q
    Feuil42.Range("chaulage_outlet_MV_MES").Cells(Ligne, Colonne) = outlet_MV_MES
    Feuil42.Range("chaulage_electricite").Cells(Ligne, Colonne) = electricite

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "42.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub effacer_resultats(ByVal numero_chaulage)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Ligne = numero_chaulage

    Feuil42.Range("chaulage_inlet1_ratio_admis").Cells(Ligne, Colonne) = ""
    Feuil42.Range("chaulage_inlet2_ratio_admis").Cells(Ligne, Colonne) = ""
    Feuil42.Range("chaulage_inlet_MES").Cells(Ligne, Colonne) = ""
    Feuil42.Range("chaulage_inlet_Q").Cells(Ligne, Colonne) = ""
    Feuil42.Range("chaulage_inlet_concentration").Cells(Ligne, Colonne) = ""
    Feuil42.Range("chaulage_inlet_MV_MES").Cells(Ligne, Colonne) = ""
    Feuil42.Range("chaulage_conc_req").Cells(Ligne, Colonne) = ""
    Feuil42.Range("chaulage_taux_chaux").Cells(Ligne, Colonne) = ""
    Feuil42.Range("chaulage_quant_chaux").Cells(Ligne, Colonne) = ""
    Feuil42.Range("chaulage_outlet_MV_MES").Cells(Ligne, Colonne) = ""
    Feuil42.Range("chaulage_electricite").Cells(Ligne, Colonne) = ""
    Feuil42.Range("chaulage_outlet_concentration").Cells(Ligne, Colonne) = ""
    Feuil42.Range("chaulage_outlet_MES").Cells(Ligne, Colonne) = ""
    Feuil42.Range("chaulage_outlet_Q").Cells(Ligne, Colonne) = ""

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "42.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub
Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 0  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil42.Range("chaulage_erreur_nom") & " " & stockage_numero_chaulage

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "42.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub Class_Terminate()



    If pn_EffacementResultatsEnCours = False Then
        Call affich_remarque
    End If



End Sub


