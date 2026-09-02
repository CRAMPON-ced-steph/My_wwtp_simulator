Attribute VB_Name = "z_Deshydratation"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'*********************************************************************************************
'Ce module fait les bilans matières sur le procédé déshydratation
'Ce module calcule les consommations électriques, de réactifs du procédé de déshydratation
'*********************************************************************************************

'constantes pour les choix de type de déshydratation
Private Const centrifuge As Integer = 1
Private Const filtre_plateaux As Integer = 2
Private Const filtre_bandes As Integer = 3

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

Private Const nb_inlet = 2

Private Const critere_incoherence = 0.1

Private Const conditionnement_polymere_FeCl3 = 1
Private Const conditionnement_chaux_FeCl3 = 2
Private Const conditionnement_pas = 3

Private choix_conditionnement As Integer
Private Const conditionnement_chaux_eteinte = 1
Private choix_conditionnement_chaux As Integer

'variables de choix
Private choix_type_deshydratation As Integer
Private choix_inlet(nb_inlet) As Integer
Public choix_polymere_type As Integer

'variables à échanger avec l'interface
Private inlet_ratio_admis(nb_inlet) As Double
Private inlet_MES As Double
Private inlet_MS As Double
Private inlet_Q As Double
Private inlet_concentration As Double
Private capture_MES As Double
Public polymere_flux As Double
Public chaux_eteinte_flux As Double
Public chaux_vive_flux As Double
Public FeCl3_flux As Double
Private chaux_flux As Double

Private outlet_concentration As Double
Private outlet_MES As Double
Public outlet_MS As Double
Public outlet_Q As Double
Private extraction_pompe_nb As Integer
Private extraction_P_refoulement As Double
Private extraction_tps_fonctionnement As Double
Private extraction_Q_unitaire As Double
Private extraction_pompe_rdt As Double
Public electricite As Double

Private dose_chaux As Double
Private dose_FeCl3 As Double
Private dose_polymere As Double

'variables à échanger avec l'interface
Private capture_MES_force As Boolean
Private outlet_concentration_force As Boolean
Private extraction_pompe_nb_force As Boolean
Private extraction_P_refoulement_force As Boolean
Private extraction_tps_fonctionnement_force As Boolean
Private extraction_Q_unitaire_force As Boolean
Private extraction_pompe_rdt_force As Boolean
Private dose_polymere_force As Boolean
Private dose_chaux_force As Boolean
Private dose_FeCl3_force As Boolean

'variables à échanger avec l'interface
Private capture_MES_f As Double
Private outlet_concentration_f As Double
Private extraction_pompe_nb_f As Integer
Private extraction_P_refoulement_f As Double
Private extraction_tps_fonctionnement_f As Double
Private extraction_Q_unitaire_f As Double
Private extraction_pompe_rdt_f As Double
Private dose_polymere_f As Double
Private dose_chaux_f As Double
Private dose_FeCl3_f As Double

'variables intermédiaires
Private inlet_etape_anterieure(nb_inlet) As Integer
Private inlet_type_boues(nb_inlet) As Integer

'hypothèses
Private polymere_Q As Double
Private chaux_Q As Double
Private FeCl3_Q As Double
Private dose_chaux_kgCaO_kgMV As Double
Private hyp_tps_fct_pompe_extraction As Double
Public boues_masse_volumique_kgm3 As Double
Private polymere_concentration As Double
Private outlet_concentration_reference(nb_qualites_boues) As Double
Private correction_concentration_boues_digerees As Double
Private hyp_taux_polymere_kgma_TMS(nb_qualites_boues) As Double
Private hyp_taux_polymere_kgma_TMS_boues_digerees(nb_qualites_boues) As Double
Private ratio_soluble_NH4_NK As Double
Private ratio_soluble_DBO_DCO As Double

'erreurs
Private remarque(3) As String
Private stockage_numero_deshydratation As Integer

'variable pour savoir si l'on a du technosable en entrée pour attribuer les valeurs guides
Public inlet_technosable As Boolean
Private outlet_concentration_technosable As Double
Public Sub lecture_choix(ByVal numero_deshydratation)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String
    Dim I As Integer

    stockage_numero_deshydratation = numero_deshydratation
    choix_type_deshydratation = Feuil28.Range("deshydratation_choix_type").Cells(numero_deshydratation, pi_FiliereConsideree)

    If choix_type_deshydratation = filtre_plateaux Then
        choix_conditionnement = Feuil28.Range("deshydratation_filtre_plateaux_type").Cells(numero_deshydratation, pi_FiliereConsideree)
        If choix_conditionnement = conditionnement_chaux_FeCl3 Then
            choix_conditionnement_chaux = Feuil28.Range("deshydratation_choix_chaux").Cells(numero_deshydratation, pi_FiliereConsideree)
        End If
    Else
        choix_conditionnement = conditionnement_pas
    End If

    'données entrée 1
    choix_inlet(1) = Feuil28.Range("deshydratation_choix_inlet1").Cells(numero_deshydratation, pi_FiliereConsideree)
    inlet_ratio_admis(1) = Feuil28.Range("deshydratation_inlet1_ratio_admis").Cells(numero_deshydratation, pi_FiliereConsideree)

    'données entrée 2
    choix_inlet(2) = Feuil28.Range("deshydratation_choix_inlet2").Cells(numero_deshydratation, pi_FiliereConsideree)
    inlet_ratio_admis(2) = Feuil28.Range("deshydratation_inlet2_ratio_admis").Cells(numero_deshydratation, pi_FiliereConsideree)


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

    'Si les boues en entrée proviennent d'Athos, on ne peut avoir qu'un filtre à plateaux
    inlet_technosable = False
    For I = 1 To nb_inlet
        If inlet_etape_anterieure(I) = boues_athos Then
            inlet_technosable = True
            If choix_type_deshydratation <> filtre_plateaux Or choix_conditionnement <> conditionnement_pas Then
                MsgBox Feuil28.Range("deshydratation_erreur_choix_type"), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub
            End If
        End If
    Next I

    choix_polymere_type = Feuil28.Range("deshydratation_choix_polymere_type").Cells(numero_deshydratation, pi_FiliereConsideree)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "28.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub
Public Sub lecture_valeurs_forcees(ByVal numero_deshydratation)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    capture_MES_force = Feuil28.Range("deshydratation_capture_MES").Cells(numero_deshydratation, pi_FiliereConsideree)
    outlet_concentration_force = Feuil28.Range("deshydratation_outlet_concentration").Cells(numero_deshydratation, pi_FiliereConsideree)
    extraction_pompe_nb_force = Feuil28.Range("deshydratation_extraction_pompe_nb").Cells(numero_deshydratation, pi_FiliereConsideree)
    extraction_P_refoulement_force = Feuil28.Range("deshydratation_extraction_P_refoulement").Cells(numero_deshydratation, pi_FiliereConsideree)
    extraction_tps_fonctionnement_force = Feuil28.Range("deshydratation_extraction_tps_fonctionnement").Cells(numero_deshydratation, pi_FiliereConsideree)
    extraction_Q_unitaire_force = Feuil28.Range("deshydratation_extraction_Q_unitaire").Cells(numero_deshydratation, pi_FiliereConsideree)
    extraction_pompe_rdt_force = Feuil28.Range("deshydratation_extraction_pompe_rdt").Cells(numero_deshydratation, pi_FiliereConsideree)
    If choix_type_deshydratation = filtre_plateaux Then
        If choix_conditionnement = conditionnement_polymere_FeCl3 Then
            dose_chaux_force = False
            dose_FeCl3_force = Feuil28.Range("deshydratation_FeCl3").Cells(numero_deshydratation, pi_FiliereConsideree)
            dose_polymere_force = Feuil28.Range("deshydratation_polymere_flux").Cells(numero_deshydratation, pi_FiliereConsideree)
        ElseIf choix_conditionnement = conditionnement_chaux_FeCl3 Then
            dose_chaux_force = Feuil28.Range("deshydratation_chaux").Cells(numero_deshydratation, pi_FiliereConsideree)
            dose_FeCl3_force = Feuil28.Range("deshydratation_FeCl3").Cells(numero_deshydratation, pi_FiliereConsideree)
            dose_polymere_force = False
        End If
    Else
        dose_polymere_force = Feuil28.Range("deshydratation_polymere_flux").Cells(numero_deshydratation, pi_FiliereConsideree)
        dose_chaux_force = False
        dose_FeCl3_force = False
    End If


    If capture_MES_force = True Then
        capture_MES_f = Feuil28.Range("deshydratation_capture_MES").Cells(numero_deshydratation, pi_FiliereConsideree)
    End If
    If dose_polymere_force = True Then
        dose_polymere_f = Feuil28.Range("deshydratation_polymere_flux").Cells(numero_deshydratation, pi_FiliereConsideree)
    End If
    If dose_chaux_force = True Then
        dose_chaux_f = Feuil28.Range("deshydratation_chaux").Cells(numero_deshydratation, pi_FiliereConsideree)
    End If
    If dose_FeCl3_force = True Then
        dose_FeCl3_f = Feuil28.Range("deshydratation_FeCl3").Cells(numero_deshydratation, pi_FiliereConsideree)
    End If
    If outlet_concentration_force = True Then
        outlet_concentration_f = Feuil28.Range("deshydratation_outlet_concentration").Cells(numero_deshydratation, pi_FiliereConsideree) * boues_masse_volumique_kgm3
    End If
    If extraction_pompe_nb_force = True Then
        extraction_pompe_nb_f = Feuil28.Range("deshydratation_extraction_pompe_nb").Cells(numero_deshydratation, pi_FiliereConsideree)
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement_f = Feuil28.Range("deshydratation_extraction_P_refoulement").Cells(numero_deshydratation, pi_FiliereConsideree)
    End If
    If extraction_tps_fonctionnement_force = True Then
        extraction_tps_fonctionnement_f = Feuil28.Range("deshydratation_extraction_tps_fonctionnement").Cells(numero_deshydratation, pi_FiliereConsideree)
    End If
    If extraction_Q_unitaire_force = True Then
        extraction_Q_unitaire_f = Feuil28.Range("deshydratation_extraction_Q_unitaire").Cells(numero_deshydratation, pi_FiliereConsideree)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt_f = Feuil28.Range("deshydratation_extraction_pompe_rdt").Cells(numero_deshydratation, pi_FiliereConsideree)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "28.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Public Sub ecriture_resultats(ByVal numero_deshydratation)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer


    Ligne = numero_deshydratation
    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


    Feuil28.Range("deshydratation_inlet1_ratio_admis").Cells(Ligne, Colonne) = inlet_ratio_admis(1)
    Feuil28.Range("deshydratation_inlet2_ratio_admis").Cells(Ligne, Colonne) = inlet_ratio_admis(2)
    Feuil28.Range("deshydratation_inlet_MES").Cells(Ligne, Colonne) = inlet_MS
    Feuil28.Range("deshydratation_inlet_Q").Cells(Ligne, Colonne) = inlet_Q
    Feuil28.Range("deshydratation_inlet_concentration").Cells(Ligne, Colonne) = inlet_concentration
    Feuil28.Range("deshydratation_capture_MES").Cells(Ligne, Colonne) = capture_MES
    Feuil28.Range("deshydratation_polymere_flux").Cells(Ligne, Colonne) = dose_polymere
    Feuil28.Range("deshydratation_outlet_concentration").Cells(Ligne, Colonne) = outlet_concentration / boues_masse_volumique_kgm3
    Feuil28.Range("deshydratation_outlet_MES").Cells(Ligne, Colonne) = outlet_MS
    Feuil28.Range("deshydratation_outlet_Q").Cells(Ligne, Colonne) = outlet_Q
    If choix_type_deshydratation = filtre_plateaux Then
        Feuil28.Range("deshydratation_extraction_pompe_nb").Cells(Ligne, Colonne) = "'-"
        Feuil28.Range("deshydratation_extraction_P_refoulement").Cells(Ligne, Colonne) = "'-"
        Feuil28.Range("deshydratation_extraction_tps_fonctionnement").Cells(Ligne, Colonne) = "'-"
        Feuil28.Range("deshydratation_extraction_Q_unitaire").Cells(Ligne, Colonne) = "'-"
        Feuil28.Range("deshydratation_extraction_pompe_rdt").Cells(Ligne, Colonne) = "'-"
    Else
        Feuil28.Range("deshydratation_extraction_pompe_nb").Cells(Ligne, Colonne) = extraction_pompe_nb
        Feuil28.Range("deshydratation_extraction_P_refoulement").Cells(Ligne, Colonne) = extraction_P_refoulement
        Feuil28.Range("deshydratation_extraction_tps_fonctionnement").Cells(Ligne, Colonne) = extraction_tps_fonctionnement
        Feuil28.Range("deshydratation_extraction_Q_unitaire").Cells(Ligne, Colonne) = extraction_Q_unitaire
        Feuil28.Range("deshydratation_extraction_pompe_rdt").Cells(Ligne, Colonne) = extraction_pompe_rdt
    End If
    Feuil28.Range("deshydratation_electricite").Cells(Ligne, Colonne) = electricite

    Feuil28.Range("deshydratation_chaux").Cells(Ligne, Colonne) = dose_chaux
    Feuil28.Range("deshydratation_fecl3").Cells(Ligne, Colonne) = dose_FeCl3


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "28.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Public Sub effacer_resultats(ByVal numero_deshydratation)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Ligne As Integer
    Dim Colonne As Integer


    Ligne = numero_deshydratation
    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


    Feuil28.Range("deshydratation_inlet1_ratio_admis").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_inlet2_ratio_admis").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_inlet_MES").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_inlet_Q").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_inlet_concentration").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_capture_MES").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_polymere_flux").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_outlet_concentration").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_outlet_MES").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_outlet_Q").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_extraction_pompe_nb").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_extraction_P_refoulement").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_extraction_tps_fonctionnement").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_extraction_Q_unitaire").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_extraction_pompe_rdt").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_electricite").Cells(Ligne, Colonne) = ""

    Feuil28.Range("deshydratation_chaux").Cells(Ligne, Colonne) = ""
    Feuil28.Range("deshydratation_fecl3").Cells(Ligne, Colonne) = ""

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "28.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Public Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'concentration espérée en sortie   en g/L
    If choix_type_deshydratation = centrifuge Then
        'primaire
        outlet_concentration_reference(I_simple) = 326
        outlet_concentration_reference(I_reactif) = 308
        'secondaire
        outlet_concentration_reference(II_forte) = 282
        outlet_concentration_reference(II_moyenne) = 220
        outlet_concentration_reference(II_faible_EB) = 200
        outlet_concentration_reference(II_faible_ED) = 195
        outlet_concentration_reference(II_prolongee_EB) = 200
        outlet_concentration_reference(II_prolongee_ED) = 195
        outlet_concentration_reference(II_MBR) = 180
        outlet_concentration_reference(II_MBBR) = 190
        outlet_concentration_reference(II_HybAS) = 190
        outlet_concentration_reference(II_biostyr_C) = 246
        outlet_concentration_reference(II_biostyr_N) = 229
        outlet_concentration_reference(II_biostyr_NDN) = 202
        outlet_concentration_reference(II_biostyr_PDN) = 185
        'tertiaire
        outlet_concentration_reference(III_decantation) = 185
        outlet_concentration_reference(III_biostyr_N) = 185
        outlet_concentration_reference(III_biostyr_PDN) = 185
        'graisses
        outlet_concentration_reference(codigestion_graisses) = 150
    ElseIf choix_type_deshydratation = filtre_plateaux Then
        'primaire
        outlet_concentration_reference(I_simple) = 420
        outlet_concentration_reference(I_reactif) = 400
        'secondaire
        outlet_concentration_reference(II_forte) = 360
        outlet_concentration_reference(II_moyenne) = 290
        outlet_concentration_reference(II_faible_EB) = 260
        outlet_concentration_reference(II_faible_ED) = 260
        outlet_concentration_reference(II_prolongee_EB) = 260
        outlet_concentration_reference(II_prolongee_ED) = 260
        outlet_concentration_reference(II_MBR) = 240
        outlet_concentration_reference(II_MBBR) = 250
        outlet_concentration_reference(II_HybAS) = 250
        outlet_concentration_reference(II_biostyr_C) = 320
        outlet_concentration_reference(II_biostyr_N) = 300
        outlet_concentration_reference(II_biostyr_NDN) = 270
        outlet_concentration_reference(II_biostyr_PDN) = 240
        'tertiaire
        outlet_concentration_reference(III_decantation) = 240
        outlet_concentration_reference(III_biostyr_N) = 240
        outlet_concentration_reference(III_biostyr_PDN) = 240
        'graisses
        outlet_concentration_reference(codigestion_graisses) = 200
        'technosable
        outlet_concentration_technosable = 600   'en fait 60% siccité, la concentration est différente car la masse volumique est plus élevée mais on écrit comme cela car le calcul est plus juste
    ElseIf choix_type_deshydratation = filtre_bandes Then
        'primaire
        outlet_concentration_reference(I_simple) = 278
        outlet_concentration_reference(I_reactif) = 263
        'secondaire
        outlet_concentration_reference(II_forte) = 240
        outlet_concentration_reference(II_moyenne) = 188
        outlet_concentration_reference(II_faible_EB) = 170
        outlet_concentration_reference(II_faible_ED) = 170
        outlet_concentration_reference(II_prolongee_EB) = 170
        outlet_concentration_reference(II_prolongee_ED) = 170
        outlet_concentration_reference(II_MBR) = 154
        outlet_concentration_reference(II_MBBR) = 170
        outlet_concentration_reference(II_HybAS) = 170
        outlet_concentration_reference(II_biostyr_C) = 210
        outlet_concentration_reference(II_biostyr_N) = 195
        outlet_concentration_reference(II_biostyr_NDN) = 173
        outlet_concentration_reference(II_biostyr_PDN) = 158
        'tertiaire
        outlet_concentration_reference(III_decantation) = 158
        outlet_concentration_reference(III_biostyr_N) = 158
        outlet_concentration_reference(III_biostyr_PDN) = 158
        'graisses
        outlet_concentration_reference(codigestion_graisses) = 128
    End If

    'taux de polymère attendu en kg/T par type de boues (sans digestion préalable)
    If choix_type_deshydratation = centrifuge Then
        'primaire
        hyp_taux_polymere_kgma_TMS(I_simple) = 5
        hyp_taux_polymere_kgma_TMS(I_reactif) = 6
        'secondaire
        hyp_taux_polymere_kgma_TMS(II_forte) = 7
        hyp_taux_polymere_kgma_TMS(II_moyenne) = 8
        hyp_taux_polymere_kgma_TMS(II_faible_EB) = 12
        hyp_taux_polymere_kgma_TMS(II_faible_ED) = 11.5
        hyp_taux_polymere_kgma_TMS(II_prolongee_EB) = 10.5
        hyp_taux_polymere_kgma_TMS(II_prolongee_ED) = 12
        hyp_taux_polymere_kgma_TMS(II_MBR) = 12
        hyp_taux_polymere_kgma_TMS(II_MBBR) = 10.5
        hyp_taux_polymere_kgma_TMS(II_HybAS) = 11.5
        hyp_taux_polymere_kgma_TMS(II_biostyr_C) = 8.5
        hyp_taux_polymere_kgma_TMS(II_biostyr_N) = 9.5
        hyp_taux_polymere_kgma_TMS(II_biostyr_NDN) = 10.5
        hyp_taux_polymere_kgma_TMS(II_biostyr_PDN) = 11.5
        'tertiaire
        hyp_taux_polymere_kgma_TMS(III_decantation) = 8.5
        hyp_taux_polymere_kgma_TMS(III_biostyr_N) = 9.5
        hyp_taux_polymere_kgma_TMS(III_biostyr_PDN) = 11.5
        'graisses
        hyp_taux_polymere_kgma_TMS(codigestion_graisses) = 16.5
    ElseIf choix_type_deshydratation = filtre_plateaux Then
        'primaire
        hyp_taux_polymere_kgma_TMS(I_simple) = 3.6
        hyp_taux_polymere_kgma_TMS(I_reactif) = 3.8
        'secondaire
        hyp_taux_polymere_kgma_TMS(II_forte) = 4.3
        hyp_taux_polymere_kgma_TMS(II_moyenne) = 7.3
        hyp_taux_polymere_kgma_TMS(II_faible_EB) = 7.9
        hyp_taux_polymere_kgma_TMS(II_faible_ED) = 8.1
        hyp_taux_polymere_kgma_TMS(II_prolongee_EB) = 7.9
        hyp_taux_polymere_kgma_TMS(II_prolongee_ED) = 8.1
        hyp_taux_polymere_kgma_TMS(II_MBR) = 7.3
        hyp_taux_polymere_kgma_TMS(II_MBBR) = 7.5
        hyp_taux_polymere_kgma_TMS(II_HybAS) = 7.9
        hyp_taux_polymere_kgma_TMS(II_biostyr_C) = 5
        hyp_taux_polymere_kgma_TMS(II_biostyr_N) = 5.5
        hyp_taux_polymere_kgma_TMS(II_biostyr_NDN) = 6.4
        hyp_taux_polymere_kgma_TMS(II_biostyr_PDN) = 7.1
        'tertiaire
        hyp_taux_polymere_kgma_TMS(III_decantation) = 7.1
        hyp_taux_polymere_kgma_TMS(III_biostyr_N) = 5.5
        hyp_taux_polymere_kgma_TMS(III_biostyr_PDN) = 7.1
        'graisses
        hyp_taux_polymere_kgma_TMS(codigestion_graisses) = 15
    ElseIf choix_type_deshydratation = filtre_bandes Then
        'primaire
        hyp_taux_polymere_kgma_TMS(I_simple) = 3
        hyp_taux_polymere_kgma_TMS(I_reactif) = 3.9
        'secondaire
        hyp_taux_polymere_kgma_TMS(II_forte) = 6.5
        hyp_taux_polymere_kgma_TMS(II_moyenne) = 7.5
        hyp_taux_polymere_kgma_TMS(II_faible_EB) = 11.5
        hyp_taux_polymere_kgma_TMS(II_faible_ED) = 11.5
        hyp_taux_polymere_kgma_TMS(II_prolongee_EB) = 10
        hyp_taux_polymere_kgma_TMS(II_prolongee_ED) = 11.5
        hyp_taux_polymere_kgma_TMS(II_MBR) = 8
        hyp_taux_polymere_kgma_TMS(II_MBBR) = 10.5
        hyp_taux_polymere_kgma_TMS(II_HybAS) = 11.5
        hyp_taux_polymere_kgma_TMS(II_biostyr_C) = 8
        hyp_taux_polymere_kgma_TMS(II_biostyr_N) = 9
        hyp_taux_polymere_kgma_TMS(II_biostyr_NDN) = 10
        hyp_taux_polymere_kgma_TMS(II_biostyr_PDN) = 11.5
        'tertiaire
        hyp_taux_polymere_kgma_TMS(III_decantation) = 8
        hyp_taux_polymere_kgma_TMS(III_biostyr_N) = 9
        hyp_taux_polymere_kgma_TMS(III_biostyr_PDN) = 11.5
        'graisses
        hyp_taux_polymere_kgma_TMS(codigestion_graisses) = 16
    End If

    'taux de polymère attendu en kg/T par type de boues (après digestion)
    If choix_type_deshydratation = centrifuge Then
        'primaire
        hyp_taux_polymere_kgma_TMS_boues_digerees(I_simple) = 6.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(I_reactif) = 6.5
        'secondaire
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_forte) = 9.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_moyenne) = 9.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_faible_EB) = 9.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_faible_ED) = 9.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_prolongee_EB) = 9.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_prolongee_ED) = 9.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_MBR) = 9.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_MBBR) = 9.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_HybAS) = 9.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_biostyr_C) = 9.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_biostyr_N) = 9.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_biostyr_NDN) = 9.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_biostyr_PDN) = 9.5
        'tertiaire
        hyp_taux_polymere_kgma_TMS_boues_digerees(III_decantation) = 8.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(III_biostyr_N) = 9.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(III_biostyr_PDN) = 9.5
        'graisses
        hyp_taux_polymere_kgma_TMS_boues_digerees(codigestion_graisses) = 16.5
    ElseIf choix_type_deshydratation = filtre_plateaux Then
        'primaire
        hyp_taux_polymere_kgma_TMS_boues_digerees(I_simple) = 5.7
        hyp_taux_polymere_kgma_TMS_boues_digerees(I_reactif) = 5.7
        'secondaire
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_forte) = 8.4
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_moyenne) = 8.4
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_faible_EB) = 8.4
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_faible_ED) = 8.4
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_prolongee_EB) = 8.4
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_prolongee_ED) = 8.4
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_MBR) = 8.4
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_MBBR) = 8.4
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_HybAS) = 8.4
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_biostyr_C) = 8.4
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_biostyr_N) = 8.4
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_biostyr_NDN) = 8.4
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_biostyr_PDN) = 8.4
        'tertiaire
        hyp_taux_polymere_kgma_TMS_boues_digerees(III_decantation) = 7.1
        hyp_taux_polymere_kgma_TMS_boues_digerees(III_biostyr_N) = 5.5
        hyp_taux_polymere_kgma_TMS_boues_digerees(III_biostyr_PDN) = 7.1
        'graisses
        hyp_taux_polymere_kgma_TMS_boues_digerees(codigestion_graisses) = 15
    ElseIf choix_type_deshydratation = filtre_bandes Then
        'primaire
        hyp_taux_polymere_kgma_TMS_boues_digerees(I_simple) = 6
        hyp_taux_polymere_kgma_TMS_boues_digerees(I_reactif) = 6
        'secondaire
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_forte) = 9
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_moyenne) = 9
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_faible_EB) = 9
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_faible_ED) = 9
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_prolongee_EB) = 9
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_prolongee_ED) = 9
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_MBR) = 9
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_MBBR) = 9
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_HybAS) = 9
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_biostyr_C) = 9
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_biostyr_N) = 9
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_biostyr_NDN) = 9
        hyp_taux_polymere_kgma_TMS_boues_digerees(II_biostyr_PDN) = 9
        'tertiaire
        hyp_taux_polymere_kgma_TMS_boues_digerees(III_decantation) = 8
        hyp_taux_polymere_kgma_TMS_boues_digerees(III_biostyr_N) = 9
        hyp_taux_polymere_kgma_TMS_boues_digerees(III_biostyr_PDN) = 9
        'graisses
        hyp_taux_polymere_kgma_TMS_boues_digerees(codigestion_graisses) = 16
    End If

    'correction siccité des boues digérées (on ajoute environ 20 g/L)
    correction_concentration_boues_digerees = 20

    'part de NH4 dans les NK soluble (pour connaître le NH4 dans les retours
    ratio_soluble_NH4_NK = 0.95

    'part de DBO dans la DCO soluble
    If choix_dig_simple = True Then
        ratio_soluble_DBO_DCO = 0.3 * 0.8  'cf Christophe AMIEL
    ElseIf pi_ChoixDigestion = True Then
        ratio_soluble_DBO_DCO = 0.2 * 0.8  'cf Christophe AMIEL
    End If

    'polymère ajouté, le dosage est en kg de m.a. par tonne de MS
    If choix_type_deshydratation = centrifuge Then
        polymere_concentration = 1.5  'g/L polymère injecté
    ElseIf choix_type_deshydratation = filtre_plateaux Then
        polymere_concentration = 1.5  'g/L polymère injecté
    ElseIf choix_type_deshydratation = filtre_bandes Then
        polymere_concentration = 1.5  'g/L polymère injecté
    End If

    'masse volumique des boues en kg/m³
    boues_masse_volumique_kgm3 = 1000

    hyp_tps_fct_pompe_extraction = 5


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "28.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub
Public Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String



    'Taux de capture des MES par défaut
    If capture_MES_force = True Then
        capture_MES = capture_MES_f
    Else
        If choix_type_deshydratation = centrifuge Then
            capture_MES = 0.95
        ElseIf choix_type_deshydratation = filtre_plateaux Then
            If inlet_technosable = True Then
                capture_MES = 0.995
            Else
                capture_MES = 0.98
            End If
        ElseIf choix_type_deshydratation = filtre_bandes Then
            capture_MES = 0.96
        End If
    End If

    'Concentration des boues obtenues en sortie  'g/L
    'ici on initialise simplement à une valeur pas trop délirante
    If choix_type_deshydratation = centrifuge Then
        outlet_concentration = 270
    ElseIf choix_type_deshydratation = filtre_plateaux Then
        If inlet_technosable = True Then
            outlet_concentration = 600
        Else
            outlet_concentration = 400
        End If
    ElseIf choix_type_deshydratation = filtre_bandes Then
        outlet_concentration = 230
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
        extraction_tps_fonctionnement = hyp_tps_fct_pompe_extraction
    End If

    'ajout de réactifs
    If choix_type_deshydratation = filtre_plateaux Then
        If choix_conditionnement = conditionnement_polymere_FeCl3 Then    'polymere_fecl3
            dose_chaux_kgCaO_kgMV = 0
            '            If dose_polymere_force = True Then
            '                dose_polymere = dose_polymere_f
            '            Else
            '                dose_polymere = 9
            '            End If
            If dose_FeCl3_force = True Then
                dose_FeCl3 = dose_FeCl3_f
            Else
                dose_FeCl3 = 50
            End If
        ElseIf choix_conditionnement = conditionnement_chaux_FeCl3 Then    'chaux_FeCl3
            If dose_chaux_force = True Then
                dose_chaux = dose_chaux_f
            Else
                dose_chaux_kgCaO_kgMV = 190 * 56.08 / 74.08    'kgCaO / tMV   EOLIA: 190 kg Ca(OH)2 / t MV
            End If
            '            dose_polymere = 0
            If dose_FeCl3_force = True Then
                dose_FeCl3 = dose_FeCl3_f
            Else
                dose_FeCl3 = 50
            End If
        Else                   'pas de conditionnement
            dose_chaux_kgCaO_kgMV = 0
            '            dose_polymere = 0
            dose_FeCl3 = 0
        End If
    Else
        'polymère ajouté, le dosage est en kg de m.a. par tonne de MS
        '        If dose_polymere_force = True Then
        '            dose_polymere = dose_polymere_f
        '        Else
        '            If choix_type_deshydratation = centrifuge Then
        '                dose_polymere = 5
        '            ElseIf choix_type_deshydratation = filtre_bandes Then
        '                dose_polymere = 5
        '            End If
        '        End If
        dose_chaux_kgCaO_kgMV = 0
        dose_FeCl3 = 0
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "28.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub

Public Sub dimensionnement(ByRef boues_flux() As Double, ByVal numero_deshydratation As Integer, ByRef retour_flux() As Double, ByRef retour_digestion() As Double, ByRef retour_flux_soluble() As Double, ByRef retour_digestion_soluble() As Double, ByRef boues_pollution_soluble() As Double, ByRef boues_I_detail() As Double, ByRef boues_II_detail() As Double, ByRef boues_III_detail() As Double, ByRef retour_athos() As Double)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Q_retour As Double

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
    Dim intermediaire_MV As Double
    Dim intermediaire_reduction_MV As Double
    Dim inlet_MV As Double
    Dim outlet_MV As Double

    Dim intermediaire_MS As Double

    Dim soluble_kg(nb_repere_mgL) As Double
    Dim ratio_boues_digerees As Double
    Dim ratio_boues_bio As Double

    inlet_MV = 0
    inlet_MES = 0
    inlet_Q = 0
    inlet_DCO = 0
    inlet_DBO = 0
    inlet_NK = 0
    inlet_Pt = 0

    outlet_MES = 0
    outlet_MV = 0
    outlet_Q = 0
    outlet_DCO = 0
    inlet_DBO = 0
    outlet_NK = 0
    outlet_Pt = 0

    For I = 1 To nb_repere_mgL
        soluble_kg(I) = 0
    Next I

    dose_polymere = 0

    For I = 1 To nb_inlet

        If choix_inlet(I) <> inlet_aucun Then

            If inlet_etape_anterieure(I) = boues_inlet Then

                'variable d'entrée
                intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MES)
                If intermediaire_MES <> 0 Then
                    intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_Q)
                    intermediaire_MV = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MV_MES) * intermediaire_MES
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
                    inlet_MV = inlet_MV + intermediaire_MV

                    'ENTREE SOLUBLE
                    'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                    For k = 1 To nb_repere_mgL
                        soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), k) * intermediaire_Q / 1000
                    Next k

                    'calcul du polymère
                    dose_polymere = dose_polymere + hyp_taux_polymere_kgma_TMS(boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_origine)) * (intermediaire_MES + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_MS_soluble) * intermediaire_Q / 1000)

                    'SEPARATION (on applique un taux de capture et on recalcule la concentration)
                    intermediaire_MES = intermediaire_MES * capture_MES
                    intermediaire_MV = capture_MES * intermediaire_MV
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

                    'je calcule les paramètres des boues déshydratées
                    intermediaire_MV = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MV_MES) * boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) + intermediaire_MV
                    intermediaire_DCO = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_DCO_MES) * boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) + intermediaire_DCO
                    intermediaire_DBO = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_DBO_MES) * boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) + intermediaire_DBO
                    intermediaire_NK = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_NK_MES) * boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) + intermediaire_NK
                    intermediaire_Pt = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_Pt_MES) * boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) + intermediaire_Pt
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) + intermediaire_MES
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_Q) = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_Q) + intermediaire_Q
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MV_MES) = intermediaire_MV / boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_DBO_MES) = intermediaire_DBO / boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES)

                    'pour la vérification des flux
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_flux_in) = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_flux_in)

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

                        inlet_MES = inlet_MES + intermediaire_MES
                        inlet_Q = inlet_Q + intermediaire_Q
                        inlet_DCO = inlet_DCO + intermediaire_DCO
                        inlet_DBO = inlet_DBO + intermediaire_DBO
                        inlet_NK = inlet_NK + intermediaire_NK
                        inlet_Pt = inlet_Pt + intermediaire_Pt
                        inlet_MV = inlet_MV + intermediaire_MV

                        'ENTREE SOLUBLE
                        'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                        For k = 1 To nb_repere_mgL
                            soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), j, k) * intermediaire_Q / 1000
                        Next k

                        'calcul du polymère
                        If inlet_etape_anterieure(I) = boues_digerees Then
                            dose_polymere = dose_polymere + hyp_taux_polymere_kgma_TMS_boues_digerees(boues_flux(inlet_etape_anterieure(I), j, repere_origine)) * (intermediaire_MES + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MS_soluble) * intermediaire_Q / 1000)
                        Else
                            dose_polymere = dose_polymere + hyp_taux_polymere_kgma_TMS(boues_flux(inlet_etape_anterieure(I), j, repere_origine)) * (intermediaire_MES + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MS_soluble) * intermediaire_Q / 1000)
                        End If

                        If inlet_etape_anterieure(I) = boues_digerees Then
                            ratio_boues_digerees = ratio_boues_digerees + intermediaire_MES
                        End If

                        'SEPARATION
                        intermediaire_MES = intermediaire_MES * capture_MES
                        intermediaire_MV = capture_MES * intermediaire_MV
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
                        intermediaire_MV = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MV_MES) * boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) + intermediaire_MV
                        intermediaire_DCO = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_DCO_MES) * boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) + intermediaire_DCO
                        intermediaire_DBO = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_DBO_MES) * boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) + intermediaire_DBO
                        intermediaire_NK = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_NK_MES) * boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) + intermediaire_NK
                        intermediaire_Pt = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_Pt_MES) * boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) + intermediaire_Pt
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) + intermediaire_MES
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_Q) = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_Q) + intermediaire_Q
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MV_MES) = intermediaire_MV / boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES)
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES)
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_DBO_MES) = intermediaire_DBO / boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES)
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES)
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES)

                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_flux_in) = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_verif_flux) = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_flux_in)

                        'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                        boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                    End If

                Next j

            End If

        End If

    Next I

    If inlet_MES = 0 Or inlet_Q = 0 Then
        MsgBox Feuil5.Range("erreur_choix_deshydratation_inlet") & " " & numero_deshydratation, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If


    inlet_MS = inlet_MES + soluble_kg(repere_mgL_MS_soluble)
    If choix_type_deshydratation = centrifuge Or choix_type_deshydratation = filtre_bandes Or (choix_type_deshydratation = filtre_plateaux And choix_conditionnement = conditionnement_polymere_FeCl3) Then
        dose_polymere = dose_polymere / inlet_MS
    Else
        dose_polymere = 0
    End If

    inlet_concentration = inlet_MS / inlet_Q
    polymere_flux = dose_polymere / 1000 * inlet_MS


    If choix_conditionnement_chaux = conditionnement_chaux_eteinte Then
        If dose_chaux_force = False Then
            dose_chaux = dose_chaux_kgCaO_kgMV * 74.08 / 56.08 * inlet_MV / inlet_MES  'en kg CaO/ tMS
        End If
        chaux_eteinte_flux = dose_chaux / 1000 * inlet_MES
        chaux_flux = chaux_eteinte_flux
    Else
        If dose_chaux_force = False Then
            dose_chaux = dose_chaux_kgCaO_kgMV * inlet_MV / inlet_MES
        End If
        chaux_vive_flux = dose_chaux / 1000 * inlet_MES
        chaux_flux = chaux_vive_flux
    End If

    FeCl3_flux = dose_FeCl3 / 1000 * inlet_MES

    polymere_Q = polymere_flux / polymere_concentration


    'CALCUL DE LA CONCENTRATION EN SORTIE AVEC UNE TELLE REPARTITION DU TYPE DE BOUES
    If inlet_technosable = True Then
        outlet_concentration = outlet_concentration_technosable
    Else
        ratio_boues_digerees = ratio_boues_digerees / inlet_MES
        correction_concentration_boues_digerees = correction_concentration_boues_digerees * ratio_boues_digerees
        outlet_concentration = 0
        For I = 1 To nb_type_boues
            If I = boues_I Then
                ratio_boues_bio = 0
                For j = 1 To nb_eaux_sales_max
                    ratio_boues_bio = ratio_boues_bio + boues_I_detail(j, repere_ratio_MES)
                    intermediaire_origine = boues_I_detail(j, repere_origine)
                    intermediaire_MES = boues_I_detail(j, repere_ratio_MES) * boues_flux(boues_deshydratees(numero_deshydratation), I, repere_MES)
                    outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
                Next j
                intermediaire_MES = boues_flux(boues_deshydratees(numero_deshydratation), I, repere_MES) * (1 - ratio_boues_bio)
                intermediaire_origine = boues_flux(boues_deshydratees(numero_deshydratation), I, repere_origine)
                outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
            ElseIf I = boues_II And choix_biostyr = True Then
                For j = 1 To nb_eaux_sales_max
                    ratio_boues_bio = boues_II_detail(j, repere_ratio_MES)
                    intermediaire_origine = boues_II_detail(j, repere_origine)
                    intermediaire_MES = boues_II_detail(j, repere_ratio_MES) * boues_flux(boues_deshydratees(numero_deshydratation), I, repere_MES)
                    outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
                Next j
            ElseIf I = boues_III And choix_biostyr_N_III = True Then
                For j = 1 To nb_eaux_sales_max
                    ratio_boues_bio = boues_III_detail(j, repere_ratio_MES)
                    intermediaire_origine = boues_III_detail(j, repere_origine)
                    intermediaire_MES = boues_III_detail(j, repere_ratio_MES) * boues_flux(boues_deshydratees(numero_deshydratation), I, repere_MES)
                    outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
                Next j
            Else
                intermediaire_MES = boues_flux(boues_deshydratees(numero_deshydratation), I, repere_MES)
                intermediaire_origine = boues_flux(boues_deshydratees(numero_deshydratation), I, repere_origine)
                outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
            End If
        Next I
        outlet_concentration = outlet_concentration + correction_concentration_boues_digerees
    End If

    'ESTIMATION DES RETOURS EN TETE DE PRIMAIRE
    Q_retour = inlet_Q + polymere_Q - outlet_Q

    If inlet_technosable = True Then
        retour_athos(repere_ret_Q) = retour_athos(repere_ret_Q) + Q_retour
        retour_athos(repere_ret_MES) = retour_athos(repere_ret_MES) + inlet_MES - outlet_MES
    End If

    retour_flux(repere_ret_Q) = retour_flux(repere_ret_Q) + Q_retour
    retour_flux(repere_ret_MES) = retour_flux(repere_ret_MES) + inlet_MES - outlet_MES
    retour_flux(repere_ret_DCO) = retour_flux(repere_ret_DCO) + inlet_DCO - outlet_DCO
    retour_flux(repere_ret_DBO) = retour_flux(repere_ret_DBO) + inlet_DBO - outlet_DBO
    retour_flux(repere_ret_NK) = retour_flux(repere_ret_NK) + inlet_NK - outlet_NK
    retour_flux(repere_ret_Pt) = retour_flux(repere_ret_Pt) + inlet_Pt - outlet_Pt
    If ratio_boues_digerees > 0 Then
        retour_digestion(repere_ret_Q) = retour_digestion(repere_ret_Q) + Q_retour
        retour_digestion(repere_ret_MES) = retour_digestion(repere_ret_MES) + inlet_MES - outlet_MES
        retour_digestion(repere_ret_DCO) = retour_digestion(repere_ret_DCO) + inlet_DCO - outlet_DCO
        retour_digestion(repere_ret_DBO) = retour_digestion(repere_ret_DBO) + inlet_DBO - outlet_DBO
        retour_digestion(repere_ret_NK) = retour_digestion(repere_ret_NK) + inlet_NK - outlet_NK
        retour_digestion(repere_ret_Pt) = retour_digestion(repere_ret_Pt) + inlet_Pt - outlet_Pt
    End If

    'ajout de la pollution soluble aux retours de tête
    intermediaire_DCO = soluble_kg(repere_mgL_DCO) / (1 + outlet_Q / Q_retour)
    retour_flux(repere_ret_DCO) = retour_flux(repere_ret_DCO) + intermediaire_DCO
    retour_flux(repere_ret_DBO) = retour_flux(repere_ret_DBO) + ratio_soluble_DBO_DCO * intermediaire_DCO
    retour_flux_soluble(repere_ret_DCO) = retour_flux_soluble(repere_ret_DCO) + intermediaire_DCO
    retour_flux_soluble(repere_ret_DBO) = retour_flux_soluble(repere_ret_DBO) + ratio_soluble_DBO_DCO * intermediaire_DCO
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_deshydratees(numero_deshydratation), I, repere_mgL_DCO) = (soluble_kg(repere_mgL_DCO) - intermediaire_DCO) / outlet_Q * 1000
    Next I
    intermediaire_NK = soluble_kg(repere_mgL_NK) / (1 + outlet_Q / Q_retour)
    retour_flux(repere_ret_NK) = retour_flux(repere_ret_NK) + intermediaire_NK
    retour_flux(repere_ret_NH4) = retour_flux(repere_ret_NH4) + ratio_soluble_NH4_NK * intermediaire_NK
    retour_flux_soluble(repere_ret_NK) = retour_flux_soluble(repere_ret_NK) + intermediaire_NK
    retour_flux_soluble(repere_ret_NH4) = retour_flux_soluble(repere_ret_NH4) + ratio_soluble_NH4_NK * intermediaire_NK
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_deshydratees(numero_deshydratation), I, repere_mgL_NK) = (soluble_kg(repere_mgL_NK) - intermediaire_NK) / outlet_Q * 1000
    Next I
    intermediaire_Pt = soluble_kg(repere_mgL_Pt) / (1 + outlet_Q / Q_retour)
    retour_flux(repere_ret_Pt) = retour_flux(repere_ret_Pt) + intermediaire_Pt
    retour_flux_soluble(repere_ret_Pt) = retour_flux_soluble(repere_ret_Pt) + intermediaire_Pt
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_deshydratees(numero_deshydratation), I, repere_mgL_Pt) = (soluble_kg(repere_mgL_Pt) - intermediaire_Pt) / outlet_Q * 1000
    Next I
    intermediaire_MS = soluble_kg(repere_mgL_MS_soluble) / (1 + outlet_Q / Q_retour)
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_deshydratees(numero_deshydratation), I, repere_mgL_MS_soluble) = (soluble_kg(repere_mgL_MS_soluble) - intermediaire_MS) / outlet_Q * 1000
    Next I
    outlet_MS = outlet_MES + (soluble_kg(repere_mgL_MS_soluble) - intermediaire_MS)
    intermediaire_MV = soluble_kg(repere_mgL_MS_soluble) / (1 + outlet_Q / Q_retour)
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_deshydratees(numero_deshydratation), I, repere_mgL_MV_soluble) = (soluble_kg(repere_mgL_MV_soluble) - intermediaire_MV) / outlet_Q * 1000
    Next I


    If ratio_boues_digerees > 0 Then
        retour_digestion(repere_ret_DCO) = retour_digestion(repere_ret_DCO) + intermediaire_DCO
        retour_digestion(repere_ret_DBO) = retour_digestion(repere_ret_DBO) + ratio_soluble_DBO_DCO * intermediaire_DCO
        retour_digestion(repere_ret_NK) = retour_digestion(repere_ret_NK) + intermediaire_NK
        retour_digestion(repere_ret_NH4) = retour_digestion(repere_ret_NH4) + ratio_soluble_NH4_NK * intermediaire_NK
        retour_digestion(repere_ret_Pt) = retour_digestion(repere_ret_Pt) + intermediaire_Pt
        retour_digestion_soluble(repere_ret_DCO) = retour_digestion_soluble(repere_ret_DCO) + intermediaire_DCO
        retour_digestion_soluble(repere_ret_DBO) = retour_digestion_soluble(repere_ret_DBO) + ratio_soluble_DBO_DCO * intermediaire_DCO
        retour_digestion_soluble(repere_ret_NK) = retour_digestion_soluble(repere_ret_NK) + intermediaire_NK
        retour_digestion_soluble(repere_ret_NH4) = retour_digestion_soluble(repere_ret_NH4) + ratio_soluble_NH4_NK * intermediaire_NK
        retour_digestion_soluble(repere_ret_Pt) = retour_digestion_soluble(repere_ret_Pt) + intermediaire_Pt
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "28.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Public Sub fonctionnement_reel(ByRef boues_flux, ByVal numero_deshydratation, ByRef retour_flux, ByRef retour_digestion, ByRef retour_flux_soluble, ByRef retour_digestion_soluble, ByRef boues_pollution_soluble, ByVal boues_I_detail, ByVal boues_II_detail, ByVal boues_III_detail, ByRef retour_athos() As Double)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim Q_retour As Double

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
    Dim intermediaire_MV As Double
    Dim intermediaire_reduction_MV As Double
    Dim inlet_MV As Double
    Dim outlet_MV As Double
    Dim soluble_kg(nb_repere_mgL) As Double
    Dim ratio_boues_digerees As Double
    Dim ratio_boues_bio As Double

    Dim intermediaire_MS As Double


    inlet_MV = 0
    inlet_MES = 0
    inlet_Q = 0
    inlet_DCO = 0
    inlet_DBO = 0
    inlet_NK = 0
    inlet_Pt = 0

    outlet_MES = 0
    outlet_MV = 0
    outlet_Q = 0
    outlet_DCO = 0
    outlet_DBO = 0
    outlet_NK = 0
    outlet_Pt = 0

    For I = 1 To nb_repere_mgL
        soluble_kg(I) = 0
    Next I

    dose_polymere = 0

    For I = 1 To nb_inlet

        If choix_inlet(I) <> inlet_aucun Then

            If inlet_etape_anterieure(I) = boues_inlet Then

                'variable d'entrée
                intermediaire_MES = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MES)
                If intermediaire_MES <> 0 Then
                    intermediaire_Q = inlet_ratio_admis(I) * boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_Q)
                    intermediaire_MV = boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_MV_MES) * intermediaire_MES
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
                    inlet_MV = inlet_MV + intermediaire_MV

                    'ENTREE SOLUBLE
                    'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                    For k = 1 To nb_repere_mgL
                        soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), k) * intermediaire_Q / 1000
                    Next k

                    'calcul du polymère
                    dose_polymere = dose_polymere + hyp_taux_polymere_kgma_TMS(boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_origine)) * (intermediaire_MES + boues_pollution_soluble(inlet_etape_anterieure(I), inlet_type_boues(I), repere_mgL_MS_soluble) * intermediaire_Q / 1000)

                    'SEPARATION (on applique un taux de capture et on recalcule la concentration)
                    intermediaire_MES = intermediaire_MES * capture_MES
                    intermediaire_MV = capture_MES * intermediaire_MV
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

                    'je calcule les paramètres des boues déshydratées
                    intermediaire_MV = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MV_MES) * boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) + intermediaire_MV
                    intermediaire_DCO = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_DCO_MES) * boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) + intermediaire_DCO
                    intermediaire_DBO = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_DBO_MES) * boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) + intermediaire_DBO
                    intermediaire_NK = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_NK_MES) * boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) + intermediaire_NK
                    intermediaire_Pt = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_Pt_MES) * boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) + intermediaire_Pt
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES) + intermediaire_MES
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_Q) = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_Q) + intermediaire_Q
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MV_MES) = intermediaire_MV / boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_DBO_MES) = intermediaire_DBO / boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES)
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_MES)

                    'pour la vérification des flux
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_flux_in) = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_flux_in) + boues_flux(inlet_etape_anterieure(I), inlet_type_boues(I), repere_flux_in) * inlet_ratio_admis(I)
                    boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_verif_flux) = boues_flux(boues_deshydratees(numero_deshydratation), inlet_type_boues(I), repere_flux_in)

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

                        inlet_MES = inlet_MES + intermediaire_MES
                        inlet_Q = inlet_Q + intermediaire_Q
                        inlet_DCO = inlet_DCO + intermediaire_DCO
                        inlet_DBO = inlet_DBO + intermediaire_DBO
                        inlet_NK = inlet_NK + intermediaire_NK
                        inlet_Pt = inlet_Pt + intermediaire_Pt
                        inlet_MV = inlet_MV + intermediaire_MV
                        'ENTREE SOLUBLE
                        'boues_pollution_soluble(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                        For k = 1 To nb_repere_mgL
                            soluble_kg(k) = soluble_kg(k) + boues_pollution_soluble(inlet_etape_anterieure(I), j, k) * intermediaire_Q / 1000
                        Next k

                        'calcul du polymère
                        If inlet_etape_anterieure(I) = boues_digerees Then
                            dose_polymere = dose_polymere + hyp_taux_polymere_kgma_TMS_boues_digerees(boues_flux(inlet_etape_anterieure(I), j, repere_origine)) * (intermediaire_MES + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MS_soluble) * intermediaire_Q / 1000)
                        Else
                            dose_polymere = dose_polymere + hyp_taux_polymere_kgma_TMS(boues_flux(inlet_etape_anterieure(I), j, repere_origine)) * (intermediaire_MES + boues_pollution_soluble(inlet_etape_anterieure(I), j, repere_mgL_MS_soluble) * intermediaire_Q / 1000)
                        End If

                        If inlet_etape_anterieure(I) = boues_digerees Then
                            ratio_boues_digerees = ratio_boues_digerees + intermediaire_MES
                        End If

                        'SEPARATION
                        intermediaire_MES = intermediaire_MES * capture_MES
                        intermediaire_MV = capture_MES * intermediaire_MV
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
                        intermediaire_MV = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MV_MES) * boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) + intermediaire_MV
                        intermediaire_DCO = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_DCO_MES) * boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) + intermediaire_DCO
                        intermediaire_DBO = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_DBO_MES) * boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) + intermediaire_DBO
                        intermediaire_NK = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_NK_MES) * boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) + intermediaire_NK
                        intermediaire_Pt = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_Pt_MES) * boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) + intermediaire_Pt
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES) + intermediaire_MES
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_Q) = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_Q) + intermediaire_Q
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MV_MES) = intermediaire_MV / boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES)
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_DCO_MES) = intermediaire_DCO / boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES)
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_DBO_MES) = intermediaire_DBO / boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES)
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_NK_MES) = intermediaire_NK / boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES)
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_ratio_Pt_MES) = intermediaire_Pt / boues_flux(boues_deshydratees(numero_deshydratation), j, repere_MES)

                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_flux_in) = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_flux_in) + boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                        boues_flux(boues_deshydratees(numero_deshydratation), j, repere_verif_flux) = boues_flux(boues_deshydratees(numero_deshydratation), j, repere_flux_in)

                        'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                        boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) = boues_flux(inlet_etape_anterieure(I), j, repere_verif_flux) - boues_flux(inlet_etape_anterieure(I), j, repere_flux_in) * inlet_ratio_admis(I)
                    End If

                Next j

            End If

        End If

    Next I


    inlet_MS = inlet_MES + soluble_kg(repere_mgL_MS_soluble)
    If dose_polymere_force = True Then
        dose_polymere = dose_polymere_f
    Else
        If choix_type_deshydratation = centrifuge Or choix_type_deshydratation = filtre_bandes Or (choix_type_deshydratation = filtre_plateaux And choix_conditionnement = conditionnement_polymere_FeCl3) Then
            dose_polymere = dose_polymere / inlet_MS
        Else
            dose_polymere = 0
        End If
    End If

    inlet_concentration = inlet_MS / inlet_Q
    polymere_flux = dose_polymere / 1000 * inlet_MS



    If choix_conditionnement_chaux = conditionnement_chaux_eteinte Then
        If choix_conditionnement_chaux = conditionnement_chaux_eteinte Then
            If dose_chaux_force = False Then
                dose_chaux = dose_chaux_kgCaO_kgMV * 74.08 / 56.08 * inlet_MV / inlet_MES    'en kg CaO/ tMS
            End If
            chaux_eteinte_flux = dose_chaux / 1000 * inlet_MES
            chaux_flux = chaux_eteinte_flux
        Else
            If dose_chaux_force = False Then
                dose_chaux = dose_chaux_kgCaO_kgMV * inlet_MV / inlet_MES
            End If
            chaux_vive_flux = dose_chaux / 1000 * inlet_MES
            chaux_flux = chaux_vive_flux
        End If
    End If

    FeCl3_flux = dose_FeCl3 / 1000 * inlet_MES

    polymere_Q = polymere_flux / polymere_concentration

    'CALCUL DE LA CONCENTRATION EN SORTIE AVEC UNE TELLE REPARTITION DU TYPE DE BOUES
    If outlet_concentration_force = True Then
        outlet_concentration = outlet_concentration_f
    Else
        If inlet_technosable = True Then
            outlet_concentration = outlet_concentration_technosable
        Else
            ratio_boues_digerees = ratio_boues_digerees / inlet_MES
            correction_concentration_boues_digerees = correction_concentration_boues_digerees * ratio_boues_digerees
            outlet_concentration = 0
            For I = 1 To nb_type_boues
                If I = boues_I Then
                    ratio_boues_bio = 0
                    For j = 1 To nb_eaux_sales_max
                        ratio_boues_bio = ratio_boues_bio + boues_I_detail(j, repere_ratio_MES)
                        intermediaire_origine = boues_I_detail(j, repere_origine)
                        intermediaire_MES = boues_I_detail(j, repere_ratio_MES) * boues_flux(boues_deshydratees(numero_deshydratation), I, repere_MES)
                        outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
                    Next j
                    intermediaire_MES = boues_flux(boues_deshydratees(numero_deshydratation), I, repere_MES) * (1 - ratio_boues_bio)
                    intermediaire_origine = boues_flux(boues_deshydratees(numero_deshydratation), I, repere_origine)
                    outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
                ElseIf I = boues_II And choix_biostyr = True Then
                    For j = 1 To nb_eaux_sales_max
                        ratio_boues_bio = boues_II_detail(j, repere_ratio_MES)
                        intermediaire_origine = boues_II_detail(j, repere_origine)
                        intermediaire_MES = boues_II_detail(j, repere_ratio_MES) * boues_flux(boues_deshydratees(numero_deshydratation), I, repere_MES)
                        outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
                    Next j
                ElseIf I = boues_III And choix_biostyr_N_III = True Then
                    For j = 1 To nb_eaux_sales_max
                        ratio_boues_bio = boues_III_detail(j, repere_ratio_MES)
                        intermediaire_origine = boues_III_detail(j, repere_origine)
                        intermediaire_MES = boues_III_detail(j, repere_ratio_MES) * boues_flux(boues_deshydratees(numero_deshydratation), I, repere_MES)
                        outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
                    Next j
                Else
                    intermediaire_MES = boues_flux(boues_deshydratees(numero_deshydratation), I, repere_MES)
                    intermediaire_origine = boues_flux(boues_deshydratees(numero_deshydratation), I, repere_origine)
                    outlet_concentration = outlet_concentration + intermediaire_MES / outlet_MES * outlet_concentration_reference(intermediaire_origine)
                End If
            Next I
            outlet_concentration = outlet_concentration + correction_concentration_boues_digerees
        End If
    End If


    'ESTIMATION DES RETOURS EN TETE DE PRIMAIRE
    Q_retour = inlet_Q + polymere_Q - outlet_Q

    If inlet_technosable = True Then
        retour_athos(repere_ret_Q) = retour_athos(repere_ret_Q) + Q_retour
        retour_athos(repere_ret_MES) = retour_athos(repere_ret_MES) + inlet_MES - outlet_MES
    End If

    retour_flux(repere_ret_Q) = retour_flux(repere_ret_Q) + Q_retour
    retour_flux(repere_ret_MES) = retour_flux(repere_ret_MES) + inlet_MES - outlet_MES
    retour_flux(repere_ret_DCO) = retour_flux(repere_ret_DCO) + inlet_DCO - outlet_DCO
    retour_flux(repere_ret_DBO) = retour_flux(repere_ret_DBO) + inlet_DBO - outlet_DBO
    retour_flux(repere_ret_NK) = retour_flux(repere_ret_NK) + inlet_NK - outlet_NK
    retour_flux(repere_ret_Pt) = retour_flux(repere_ret_Pt) + inlet_Pt - outlet_Pt
    If ratio_boues_digerees > 0 Then
        retour_digestion(repere_ret_Q) = retour_digestion(repere_ret_Q) + Q_retour
        retour_digestion(repere_ret_MES) = retour_digestion(repere_ret_MES) + inlet_MES - outlet_MES
        retour_digestion(repere_ret_DCO) = retour_digestion(repere_ret_DCO) + inlet_DCO - outlet_DCO
        retour_digestion(repere_ret_DBO) = retour_digestion(repere_ret_DBO) + inlet_DBO - outlet_DBO
        retour_digestion(repere_ret_NK) = retour_digestion(repere_ret_NK) + inlet_NK - outlet_NK
        retour_digestion(repere_ret_Pt) = retour_digestion(repere_ret_Pt) + inlet_Pt - outlet_Pt
    End If

    'ajout de la pollution soluble aux retours de tête
    intermediaire_DCO = soluble_kg(repere_mgL_DCO) / (1 + outlet_Q / Q_retour)
    retour_flux(repere_ret_DCO) = retour_flux(repere_ret_DCO) + intermediaire_DCO
    retour_flux(repere_ret_DBO) = retour_flux(repere_ret_DBO) + ratio_soluble_DBO_DCO * intermediaire_DCO
    retour_flux_soluble(repere_ret_DCO) = retour_flux_soluble(repere_ret_DCO) + intermediaire_DCO
    retour_flux_soluble(repere_ret_DBO) = retour_flux_soluble(repere_ret_DBO) + ratio_soluble_DBO_DCO * intermediaire_DCO
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_deshydratees(numero_deshydratation), I, repere_mgL_DCO) = (soluble_kg(repere_mgL_DCO) - intermediaire_DCO) / outlet_Q * 1000
    Next I
    intermediaire_NK = soluble_kg(repere_mgL_NK) / (1 + outlet_Q / Q_retour)
    retour_flux(repere_ret_NK) = retour_flux(repere_ret_NK) + intermediaire_NK
    retour_flux(repere_ret_NH4) = retour_flux(repere_ret_NH4) + ratio_soluble_NH4_NK * intermediaire_NK
    retour_flux_soluble(repere_ret_NK) = retour_flux_soluble(repere_ret_NK) + intermediaire_NK
    retour_flux_soluble(repere_ret_NH4) = retour_flux_soluble(repere_ret_NH4) + ratio_soluble_NH4_NK * intermediaire_NK
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_deshydratees(numero_deshydratation), I, repere_mgL_NK) = (soluble_kg(repere_mgL_NK) - intermediaire_NK) / outlet_Q * 1000
    Next I
    intermediaire_Pt = soluble_kg(repere_mgL_Pt) / (1 + outlet_Q / Q_retour)
    retour_flux(repere_ret_Pt) = retour_flux(repere_ret_Pt) + intermediaire_Pt
    retour_flux_soluble(repere_ret_Pt) = retour_flux_soluble(repere_ret_Pt) + intermediaire_Pt
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_deshydratees(numero_deshydratation), I, repere_mgL_Pt) = (soluble_kg(repere_mgL_Pt) - intermediaire_Pt) / outlet_Q * 1000
    Next I
    intermediaire_MS = soluble_kg(repere_mgL_MS_soluble) / (1 + outlet_Q / Q_retour)
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_deshydratees(numero_deshydratation), I, repere_mgL_MS_soluble) = (soluble_kg(repere_mgL_MS_soluble) - intermediaire_MS) / outlet_Q * 1000
    Next I
    outlet_MS = outlet_MES + (soluble_kg(repere_mgL_MS_soluble) - intermediaire_MS)
    intermediaire_MV = soluble_kg(repere_mgL_MS_soluble) / (1 + outlet_Q / Q_retour)
    For I = 1 To nb_type_boues
        boues_pollution_soluble(boues_deshydratees(numero_deshydratation), I, repere_mgL_MV_soluble) = (soluble_kg(repere_mgL_MV_soluble) - intermediaire_MV) / outlet_Q * 1000
    Next I

    If ratio_boues_digerees > 0 Then
        retour_digestion(repere_ret_DCO) = retour_digestion(repere_ret_DCO) + intermediaire_DCO
        retour_digestion(repere_ret_DBO) = retour_digestion(repere_ret_DBO) + ratio_soluble_DBO_DCO * intermediaire_DCO
        retour_digestion(repere_ret_NK) = retour_digestion(repere_ret_NK) + intermediaire_NK
        retour_digestion(repere_ret_NH4) = retour_digestion(repere_ret_NH4) + ratio_soluble_NH4_NK * intermediaire_NK
        retour_digestion(repere_ret_Pt) = retour_digestion(repere_ret_Pt) + intermediaire_Pt
        retour_digestion_soluble(repere_ret_DCO) = retour_digestion_soluble(repere_ret_DCO) + intermediaire_DCO
        retour_digestion_soluble(repere_ret_DBO) = retour_digestion_soluble(repere_ret_DBO) + ratio_soluble_DBO_DCO * intermediaire_DCO
        retour_digestion_soluble(repere_ret_NK) = retour_digestion_soluble(repere_ret_NK) + intermediaire_NK
        retour_digestion_soluble(repere_ret_NH4) = retour_digestion_soluble(repere_ret_NH4) + ratio_soluble_NH4_NK * intermediaire_NK
        retour_digestion_soluble(repere_ret_Pt) = retour_digestion_soluble(repere_ret_Pt) + intermediaire_Pt
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "28.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Public Sub calcul_consommation_electrique()

    '*********************************************************************************************
    'Cette procédure calcule la consommation électrique totale du procédé déshydratation des boues
    '   - l'équipement de séparation
    '   - l'injection de réactifs
    '   - le pompage d'extraction des boues
    'Paramètres d'entrée
    '   -
    '*********************************************************************************************

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim rendement_pompe As Double
    Dim puissance_pompe As Double
    Dim rendement_moteur As Double
    Dim ratio_conso_spec As Double
    Dim ratio_conso_polymere As Double
    Dim rate_puissance_fixe As Double
    Dim rate_puissance_debit As Double

    Dim incoherence As Double

    Dim electricite_extraction As Double
    Dim electricite_specifique As Double
    Dim electricite_polymere As Double

    Dim concentration_x1 As Double
    Dim concentration_x2 As Double
    Dim ratio_conso_x1 As Double
    Dim ratio_conso_x2 As Double

    Dim ratio_boues_digerees As Double

    'CONSOMMATION SPECIFIQUE PROCEDE
    If choix_type_deshydratation = centrifuge Then
        'ratio_conso_spec = 1.8  'kWh/m3  Eric Guibelin en moyenne
        'modifié le 11 mai 2011 --> dépend de la siccité de sortie, en général siccité comprise entre 15% et 35% et ratio conso compris entre 1 et 2 kWh/m3 --> relation linéaire entre ces valeurs
        concentration_x1 = 150  'g/L
        ratio_conso_x1 = 1  'kWh/m3
        concentration_x2 = 350  'g/L
        ratio_conso_x2 = 2.5  'kWh/m3
        ratio_conso_spec = ratio_conso_x1 + (outlet_concentration - concentration_x1) / (concentration_x2 - concentration_x1) * (ratio_conso_x2 - ratio_conso_x1)
        electricite_specifique = ratio_conso_spec * (inlet_Q + polymere_Q)
        'CONSOMMATION POUR POLYMERE (préparation+injection)
        ratio_conso_polymere = 0.76    'kWh/kg polymère
        electricite_polymere = ratio_conso_polymere * polymere_flux
    ElseIf choix_type_deshydratation = filtre_plateaux Then
        ratio_conso_spec = 50   '35kWh/tMS Eric GUIBELIN   65kWh/tMS Neuville (16 mars 2012)
        electricite_specifique = ratio_conso_spec * inlet_MES / 1000
        electricite_polymere = 0
    ElseIf choix_type_deshydratation = filtre_bandes Then
        ratio_conso_spec = 30  'kWh/tMS  Eric Guibelin  (11 mai 2011)
        electricite_specifique = ratio_conso_spec * inlet_MES / 1000
        'CONSOMMATION POUR POLYMERE (préparation+injection)
        ratio_conso_polymere = 0.76    'kWh/kg polymère
        electricite_polymere = ratio_conso_polymere * polymere_flux
    End If




    'EXTRACTION DES BOUES
    If choix_type_deshydratation = filtre_plateaux Then
        electricite_extraction = 0
    Else
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
            remarque(2) = Feuil28.Range("deshydratation_erreur_pompage").Cells(1, 1)
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
    End If

    electricite = electricite_extraction + electricite_specifique + electricite_polymere

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "28.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub


Private Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 3  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil28.Range("deshydratation_erreur_nom") & " " & stockage_numero_deshydratation

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "28.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub Class_Initialize()

    'debut modif MSM 12/12/12
    If pn_EffacementResultatsEnCours = False Then
        Call hypotheses
    End If
    'fin modif MSM 12/12/12

End Sub

Private Sub Class_Terminate()

    If pn_EffacementResultatsEnCours = False Then
        Call affich_remarque
    End If

End Sub




