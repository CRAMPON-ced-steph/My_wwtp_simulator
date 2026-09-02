Attribute VB_Name = "z_Empreinte_CO2"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit


Dim nb_compartiments As Integer    'tous sont définis en initialisation du module de classe
Dim gaz_naturel As Integer
Dim fioul As Integer
Dim electricite As Integer
Dim transport As Integer

Private Const nb_type_transport = 3
Private Const camion_15T = 1
Private Const camion_20T = 2
Private Const semi_remorque_25T = 3

Private Const France = 1
Private Const UE = 2
Private Const Monde = 3
Private Const autre_pays = 4

Private Const used_on_site = 1
Private Const revente = 2

'variable sur les choix de l'utilisateur
Dim choix_pays As Integer
Dim choix_transport_reactifs(reactifs_nb) As Integer
Dim choix_valorisation As Integer

'variables échangées avec la synthèse
Dim FE_electricite As Double
Dim FE_reactifs(reactifs_nb) As Double
Dim FE_transport_reactifs(reactifs_nb) As Double
Dim transport_reactifs_distance(reactifs_nb) As Double
Dim FE_gaz_naturel As Double
Dim FE_fioul As Double
Dim volume_boues As Double
Dim md_QuantiteRefusDegrillage_Tj As Double
Dim md_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm As Double
Dim md_DistanceTransportRefusDegrillage_km As Double
Dim md_IncertitudeFacteurEmissionTransportRefusDegrillage As Double
Dim graisses_Tj As Double
Dim FE_transport_boues As Double
Dim FE_transport_graisses As Double
Dim transport_boues_distance As Double
Dim transport_graisses_distance As Double
Dim cendres_Tj As Double
Dim FE_transport_cendres As Double
Dim transport_cendres_distance As Double
Dim REFIB_Tj As Double
Dim FE_transport_REFIB As Double
Dim transport_REFIB_distance As Double
Dim electricite_verte_consommee As Double
Dim electricite_verte_vendue As Double
Dim emissions As Double
Dim emissions_incertitude As Double
Dim evitees As Double
Dim reduites As Double
Dim evitees_reduites_incertitude As Double
Dim ratio_EH As Double
Dim ratio_Q As Double
Dim ratio_CNP As Double
Dim ratio_fourchette_EH As String
Dim ratio_fourchette_Q As String
Dim ratio_fourchette_CNP As String
Dim emissions_repartition() As Double
Dim zone(4) As Double   'graphique demi-camembert

Dim StruviteProduite_kgj As Double

'variables montrant les valeurs forcées
Dim FE_electricite_force As Boolean
Dim FE_reactifs_force(reactifs_nb) As Boolean
Dim FE_gaz_naturel_force As Boolean
Dim FE_fioul_force As Boolean
Dim FE_transport_reactifs_force(reactifs_nb) As Boolean
Dim volume_boues_force As Boolean
Dim graisses_Tj_force As Boolean
Dim mn_QuantiteRefusDegrillage_Tj_force As Boolean
Dim mn_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm_force As Boolean
Dim FE_transport_boues_force As Boolean
Dim FE_transport_graisses_force As Boolean
Dim cendres_Tj_force As Boolean
Dim FE_transport_cendres_force As Boolean
Dim REFIB_Tj_force As Boolean
Dim FE_transport_REFIB_force As Boolean
Dim electricite_verte_consommee_force As Boolean
Dim electricite_verte_vendue_force As Boolean

'variables forcées
Dim FE_electricite_f As Double
Dim FE_reactifs_f(reactifs_nb) As Double
Dim FE_gaz_naturel_f As Double
Dim FE_fioul_f As Double
Dim FE_transport_reactifs_f(reactifs_nb) As Double
Dim volume_boues_f As Double
Dim graisses_Tj_f As Double
Dim md_QuantiteRefusDegrillage_Tj_f As Double
Dim md_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm_f As Double
Dim FE_transport_boues_f As Double
Dim FE_transport_graisses_f As Double
Dim cendres_Tj_f As Double
Dim FE_transport_cendres_f As Double
Dim REFIB_Tj_f As Double
Dim FE_transport_REFIB_f As Double
Dim electricite_verte_consommee_f As Double
Dim electricite_verte_vendue_f As Double

'hypothèses
Dim FE_electricite_incertitude As Double
Dim hyp_FE_reactifs(reactifs_nb) As Double
Dim FE_reactifs_incertitude(reactifs_nb) As Double
Dim FE_transport(nb_type_transport) As Double
Dim FE_transport_incertitude(nb_type_transport) As Double
Dim FE_transport_boues_incertitude As Double
Dim FE_transport_graisses_incertitude As Double
Dim FE_transport_cendres_incertitude As Double
Dim FE_transport_REFIB_incertitude As Double
Dim densite_boues As Double
Dim densite_graisses As Double
Dim emissions_CO2_EH_an_max As Double
Dim hyp_FE_gaz_naturel As Double
Dim hyp_FE_fioul As Double
Dim FE_gaz_naturel_incertitude As Double
Dim FE_fioul_incertitude As Double
Dim FeCl3_pureté As Double
Dim NaOCl_pureté As Double


Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer
    Dim I As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil33.Range("CO2_FE_electricite").Cells(1, Colonne) = FE_electricite
    Feuil33.Range("CO2_FE_gaz_naturel").Cells(1, Colonne) = FE_gaz_naturel
    Feuil33.Range("CO2_FE_fioul").Cells(1, Colonne) = FE_fioul


    Feuil33.Range("CO2_electricite_verte_consommee").Cells(1, Colonne) = electricite_verte_consommee
    Feuil33.Range("CO2_electricite_verte_vendue").Cells(1, Colonne) = electricite_verte_vendue


    For I = 1 To reactifs_nb
        Feuil33.Range("CO2_FE_reactifs").Cells(I, Colonne) = FE_reactifs(I)
        Feuil33.Range("CO2_FE_transport_reactifs").Cells(I, Colonne) = FE_transport_reactifs(I)
        Feuil33.Range("CO2_transport_reactifs_distance").Cells(I, Colonne) = transport_reactifs_distance(I)
    Next I

    'TRANSPORT DE SOUS-PRODUITS
    Feuil33.Range("CO2_volume_boues").Cells(1, Colonne) = volume_boues
    Feuil33.Range("CO2_FE_transport_boues").Cells(1, Colonne) = FE_transport_boues
    Feuil33.Range("CO2_distance_boues").Cells(1, Colonne) = transport_boues_distance
    Feuil33.Range("CO2_QuantiteRefusDegrillage_Tj").Cells(1, Colonne) = md_QuantiteRefusDegrillage_Tj
    Feuil33.Range("CO2_FE_transport_RefusDegrillage").Cells(1, Colonne) = md_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm
    Feuil33.Range("CO2_distance_RefusDegrillage").Cells(1, Colonne) = md_DistanceTransportRefusDegrillage_km
    Feuil33.Range("CO2_graisses_Tj").Cells(1, Colonne) = graisses_Tj
    Feuil33.Range("CO2_FE_transport_graisses").Cells(1, Colonne) = FE_transport_graisses
    Feuil33.Range("CO2_transport_graisses_distance").Cells(1, Colonne) = transport_graisses_distance
    Feuil33.Range("CO2_cendres_Tj").Cells(1, Colonne) = cendres_Tj
    Feuil33.Range("CO2_FE_transport_cendres").Cells(1, Colonne) = FE_transport_cendres
    Feuil33.Range("CO2_distance_cendres").Cells(1, Colonne) = transport_cendres_distance
    Feuil33.Range("CO2_REFIB_Tj").Cells(1, Colonne) = REFIB_Tj
    Feuil33.Range("CO2_FE_transport_REFIB").Cells(1, Colonne) = FE_transport_REFIB
    Feuil33.Range("CO2_distance_REFIB").Cells(1, Colonne) = transport_REFIB_distance
    'graphique boues + cendres + REFIB
    Feuil37.Range("graph_boues_evacuees_T_an").Cells(1, pi_FiliereConsideree) = (volume_boues + graisses_Tj + cendres_Tj + REFIB_Tj) * pd_NOMBRE_JOUR_PAR_AN


    Feuil33.Range("CO2_emissions").Cells(1, Colonne) = emissions
    Feuil33.Range("CO2_emission_incertitude").Cells(1, Colonne) = emissions_incertitude

    Feuil33.Range("CO2_evitees").Cells(1, Colonne) = evitees
    Feuil33.Range("CO2_reduites").Cells(1, Colonne) = reduites
    Feuil33.Range("CO2_evitees_reduites_incertitude").Cells(1, Colonne) = evitees_reduites_incertitude

    If choix_primaire = False And choix_secondaire = False And choix_tertiaire = False Then
        Feuil33.Range("CO2_ratio_EH").Cells(1, Colonne) = ratio_EH
        Feuil33.Range("CO2_ratio_Q").Cells(1, Colonne) = "-"
        Feuil33.Range("CO2_ratio_CNP").Cells(1, Colonne) = "-"
    Else
        Feuil33.Range("CO2_ratio_EH").Cells(1, Colonne) = ratio_EH
        Feuil33.Range("CO2_ratio_Q").Cells(1, Colonne) = ratio_Q
        Feuil33.Range("CO2_ratio_CNP").Cells(1, Colonne) = ratio_CNP
    End If

    Feuil33.Range("CO2_ratio_EH").Cells(1, pi_FiliereConsideree) = ratio_fourchette_EH
    Feuil33.Range("CO2_ratio_Q").Cells(1, pi_FiliereConsideree) = ratio_fourchette_Q
    Feuil33.Range("CO2_ratio_CNP").Cells(1, pi_FiliereConsideree) = ratio_fourchette_CNP

    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
        'comparatif
        Feuil37.Range("graph_CO2_global").Cells(1, pi_FiliereConsideree) = emissions - reduites
        Feuil37.Range("graph_CO2_global").Cells(2, pi_FiliereConsideree) = evitees
        'répartition
        Colonne = 1
        For I = 1 To reactifs_nb
            Feuil37.Range("graph_CO2_step_" & pi_FiliereConsideree).Cells(I, Colonne) = emissions_repartition(I)
        Next I
        Feuil37.Range("graph_CO2_step_" & pi_FiliereConsideree).Cells(gaz_naturel, Colonne) = emissions_repartition(gaz_naturel)
        Feuil37.Range("graph_CO2_step_" & pi_FiliereConsideree).Cells(fioul, Colonne) = emissions_repartition(fioul)
        Feuil37.Range("graph_CO2_step_" & pi_FiliereConsideree).Cells(electricite, Colonne) = emissions_repartition(electricite)
        Feuil37.Range("graph_CO2_step_" & pi_FiliereConsideree).Cells(transport, Colonne) = emissions_repartition(transport)
        'autosuffisance
        Colonne = 1
        For I = 1 To 4
            Feuil37.Range("graph_zone_step_" & pi_FiliereConsideree).Cells(I, Colonne) = zone(I)
        Next I
        Call ecriture_resultats_OPEX
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "33.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub ecriture_resultats_OPEX()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer
    Dim I As Integer

    Colonne = 3 * pi_CALCUL_VALEUR_GUIDE + pi_FiliereConsideree

    Feuil55.Range("OPEX_boues_evacuees_Tj").Cells(1, Colonne) = volume_boues
    Feuil55.Range("OPEX_EvacuationRefusDegrillage_Tj").Cells(1, Colonne) = md_QuantiteRefusDegrillage_Tj
    Feuil55.Range("OPEX_graisses_Tj").Cells(1, Colonne) = graisses_Tj
    Feuil55.Range("OPEX_cendres_evacuees_Tj").Cells(1, Colonne) = cendres_Tj
    Feuil55.Range("OPEX_REFIB_evacuees_Tj").Cells(1, Colonne) = REFIB_Tj
    Feuil55.Range("OPEX_struvite_vendue_kgj").Cells(1, Colonne) = StruviteProduite_kgj


    Feuil55.Range("OPEX_electricite_pdte_vendue").Cells(1, Colonne) = electricite_verte_vendue
    Feuil55.Range("OPEX_electricite_pdte_conso_site").Cells(1, Colonne) = electricite_verte_consommee




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "33.11"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats_OPEX()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer
    Dim I As Integer

    Colonne = 3 * pi_CALCUL_VALEUR_GUIDE + pi_FiliereConsideree

    Feuil55.Range("OPEX_boues_evacuees_Tj").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_EvacuationRefusDegrillage_Tj").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_graisses_Tj").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_cendres_evacuees_Tj").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_REFIB_evacuees_Tj").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_struvite_vendue_kgj").Cells(1, Colonne) = ""     'StruviteProduite_kgj


    Feuil55.Range("OPEX_electricite_pdte_vendue").Cells(1, Colonne) = ""
    Feuil55.Range("OPEX_electricite_pdte_conso_site").Cells(1, Colonne) = ""


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "33.12"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim Colonne As Integer
    Dim I As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil33.Range("CO2_FE_electricite").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_FE_gaz_naturel").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_FE_fioul").Cells(1, Colonne) = ""

    Feuil33.Range("CO2_electricite_verte_consommee").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_electricite_verte_vendue").Cells(1, Colonne) = ""

    For I = 1 To reactifs_nb
        Feuil33.Range("CO2_FE_reactifs").Cells(I, Colonne) = ""
        Feuil33.Range("CO2_FE_transport_reactifs").Cells(I, Colonne) = ""
        Feuil33.Range("CO2_transport_reactifs_distance").Cells(I, Colonne) = ""
    Next I

    Feuil33.Range("CO2_volume_boues").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_FE_transport_boues").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_distance_boues").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_QuantiteRefusDegrillage_Tj").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_FE_transport_RefusDegrillage").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_distance_RefusDegrillage").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_graisses_Tj").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_FE_transport_graisses").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_transport_graisses_distance").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_cendres_Tj").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_FE_transport_cendres").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_distance_cendres").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_REFIB_Tj").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_FE_transport_REFIB").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_distance_REFIB").Cells(1, Colonne) = ""

    Feuil33.Range("CO2_emissions").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_emission_incertitude").Cells(1, Colonne) = ""

    Feuil33.Range("CO2_evitees").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_reduites").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_evitees_reduites_incertitude").Cells(1, Colonne) = ""

    Feuil33.Range("CO2_ratio_EH").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_ratio_Q").Cells(1, Colonne) = ""
    Feuil33.Range("CO2_ratio_CNP").Cells(1, Colonne) = ""

    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_GUIDE Then
        Feuil33.Range("CO2_ratio_EH").Cells(1, pi_FiliereConsideree) = ""
        Feuil33.Range("CO2_ratio_Q").Cells(1, pi_FiliereConsideree) = ""
        Feuil33.Range("CO2_ratio_CNP").Cells(1, pi_FiliereConsideree) = ""
    End If

    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
        Call effacer_resultats_OPEX
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "33.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer


    choix_pays = Feuil33.Range("CO2_choix_pays").Cells(1, pi_FiliereConsideree)
    If choix_pays = autre_pays Then
        FE_electricite = Feuil33.Range("CO2_FE_electricite").Cells(1, pi_FiliereConsideree)
    End If

    For I = 1 To reactifs_nb
        choix_transport_reactifs(I) = Feuil33.Range("CO2_choix_transport_reactifs").Cells(I, pi_FiliereConsideree)
        transport_reactifs_distance(I) = Feuil33.Range("CO2_transport_reactifs_distance").Cells(I, pi_FiliereConsideree)
    Next I

    transport_boues_distance = Feuil33.Range("CO2_distance_boues").Cells(1, pi_FiliereConsideree)
    md_DistanceTransportRefusDegrillage_km = Feuil33.Range("CO2_distance_RefusDegrillage").Cells(1, pi_FiliereConsideree)
    transport_graisses_distance = Feuil33.Range("CO2_transport_graisses_distance").Cells(1, pi_FiliereConsideree)
    transport_cendres_distance = Feuil33.Range("CO2_distance_cendres").Cells(1, pi_FiliereConsideree)
    transport_REFIB_distance = Feuil33.Range("CO2_distance_REFIB").Cells(1, pi_FiliereConsideree)

    choix_valorisation = Feuil33.Range("CO2_choix_valorisation").Cells(1, pi_FiliereConsideree)



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "33.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'FACTEURS D'EMISSIONS ET INCERTITUDES RELATIVES

    Dim I As Integer

    'Electricite
    If FE_electricite_force = True Then
        FE_electricite = FE_electricite_f
    Else
        If choix_pays = France Then
            'ADEME
            FE_electricite = 0.091
            FE_electricite_incertitude = 0.15
        ElseIf choix_pays = UE Then
            'ADEME (incertitude non connue donc par défaut 50%
            FE_electricite = 0.306
            FE_electricite_incertitude = 0.5
        ElseIf choix_pays = Monde Then
            'ADEME (incertitude non connue donc par défaut 50%
            FE_electricite = 0.505
            FE_electricite_incertitude = 0.5
        End If
    End If

    'Gaz naturel
    If FE_gaz_naturel_force = True Then
        FE_gaz_naturel = FE_gaz_naturel_f
    Else
        FE_gaz_naturel = hyp_FE_gaz_naturel
    End If

    'fioul
    If FE_fioul_force = True Then
        FE_fioul = FE_fioul_f
    Else
        FE_fioul = hyp_FE_fioul
    End If

    'Réactifs
    For I = 1 To reactifs_nb
        If FE_reactifs_force(I) = True Then
            FE_reactifs(I) = FE_reactifs_f(I)
        Else
            FE_reactifs(I) = hyp_FE_reactifs(I)
        End If
        If FE_transport_reactifs_force(I) = True Then
            FE_transport_reactifs(I) = FE_transport_reactifs_f(I)
        Else
            FE_transport_reactifs(I) = FE_transport(choix_transport_reactifs(I))
        End If
    Next I

    'évacuation boues
    If FE_transport_boues_force = True Then
        FE_transport_boues = FE_transport_boues_f
    Else
        FE_transport_boues = 0.129
        FE_transport_boues_incertitude = 0.1
    End If

    'évacuation refus de dégrillage
    If mn_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm_force = True Then
        md_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm = md_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm_f
    Else
        md_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm = 0.129
        md_IncertitudeFacteurEmissionTransportRefusDegrillage = 0.1
    End If

    'évacuation des graisses
    If FE_transport_graisses_force = True Then
        FE_transport_graisses = FE_transport_graisses_f
    Else
        FE_transport_graisses = 0.129
        FE_transport_graisses_incertitude = 0.1
    End If

    'évacuation cendres
    If FE_transport_cendres_force = True Then
        FE_transport_cendres = FE_transport_cendres_f
    Else
        FE_transport_cendres = 0.129
        FE_transport_cendres_incertitude = 0.1
    End If

    'évacuation REFIB
    If FE_transport_REFIB_force = True Then
        FE_transport_REFIB = FE_transport_REFIB_f
    Else
        FE_transport_REFIB = 0.129
        FE_transport_REFIB_incertitude = 0.1
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "33.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True




End Sub

Sub hypothese()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim EF_methanol_production As Double
    Dim EF_methanol_production_incertitude As Double
    Dim EF_methanol_conso As Double
    Dim EF_methanol_conso_incertitude As Double


    'Facteurs d'émissions pour le transport

    'source ADEME  kgCO2/(T.km)
    FE_transport(camion_15T) = 0.238
    FE_transport_incertitude(camion_15T) = 0.1
    FE_transport(camion_20T) = 0.223
    FE_transport_incertitude(camion_20T) = 0.1
    FE_transport(semi_remorque_25T) = 0.119
    FE_transport_incertitude(semi_remorque_25T) = 0.1

    'Facteurs d'émission pour le gaz naturel   kgCO2/kWhPCI   source: ADEME avec amont (amont+combustion)  MAJ 08/11/2012
    hyp_FE_gaz_naturel = 0.03 + 0.18      '0.234
    FE_gaz_naturel_incertitude = 0.05

    'Facteurs d'émission pour le fioul domestique   kgCO2/kWhPCI   source: ADEME avec amont (amont+combustion)  MAJ 08/11/2012
    hyp_FE_fioul = 0.06 + 0.27
    FE_fioul_incertitude = 0.05

    'Facteurs d'émissions pour le polymere kgCO2/T   source: SNF/ASTEE v5.2010.2    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_polymere_eau_poudre_anionique) = 810
    FE_reactifs_incertitude(reactifs_repere_polymere_eau_poudre_anionique) = 0.3

    'Facteurs d'émissions pour le FeCl3 (pur)  kgCO2/T  source: Eco-Invent    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_FeCl3_eau) = 800
    FE_reactifs_incertitude(reactifs_repere_FeCl3_eau) = 0.25
    FeCl3_pureté = 1

    'Facteurs d'émissions pour le méthanol  kgCO2/T    MAJ 08/11/2012
    EF_methanol_production = 521    'ADEME BC v7
    EF_methanol_production_incertitude = 0.5
    EF_methanol_conso = 1375    'Carbone 4
    EF_methanol_conso_incertitude = 0
    hyp_FE_reactifs(reactifs_repere_methanol) = EF_methanol_production + EF_methanol_conso
    FE_reactifs_incertitude(reactifs_repere_methanol) = (EF_methanol_production * EF_methanol_production_incertitude + EF_methanol_conso * EF_methanol_conso_incertitude) / hyp_FE_reactifs(reactifs_repere_methanol)

    'Facteurs d'émissions pour la chaux éteinte kgCO2/T   source: LHOIST v5.2010.2    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_chaux_eteinte) = 846
    FE_reactifs_incertitude(reactifs_repere_chaux_eteinte) = 0.3

    'Facteurs d'émissions pour la chaux vive kgCO2/T   source: LHOIST v5.2010.2/ASTEE    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_chaux_vive) = 1110
    FE_reactifs_incertitude(reactifs_repere_chaux_vive) = 0.3

    'Facteurs d'émissions pour le H2SO4 pur kgCO2/T   source: Ademe BC v7    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_H2SO4) = 148
    FE_reactifs_incertitude(reactifs_repere_H2SO4) = 0.5

    'Facteurs d'émissions pour le NaOH pur kgCO2/T   source: Ademe BC v7    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_NaOH) = 1174
    FE_reactifs_incertitude(reactifs_repere_NaOH) = 0.5

    'Facteurs d'émissions pour le NaOCl pur kgCO2/T   source: Arkema    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_NaOCl_desodo) = 2884
    FE_reactifs_incertitude(reactifs_repere_NaOCl_desodo) = 0.5
    NaOCl_pureté = 1    '0.152 * 50 / 48   à 50°Cl (soit 15%)

    'Facteurs d'émissions pour le NaHSO3  kgCO2/T   source: ASTEE 5.2010.2    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_NaHSO3) = 420
    FE_reactifs_incertitude(reactifs_repere_NaHSO3) = 0.1

    'Facteurs d'émissions pour le NaHCO3 pur  kgCO2/T   source: Carbone 4    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_NaHCO3) = 1166
    FE_reactifs_incertitude(reactifs_repere_NaHCO3) = 0.1

    'Facteurs d'émissions pour le Ca(NO3)2 à 79%  kgCO2/T   source: YARA valeur recommandée par guide de l'ASTEE 05.2010.2    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_Ca_2NO3) = 640
    FE_reactifs_incertitude(reactifs_repere_Ca_2NO3) = 0.3

    'Facteurs d'émissions pour l'O2 liquide  kgCO2/T   source: Eco-Invent v2.2    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_O2_liquide) = 408
    FE_reactifs_incertitude(reactifs_repere_O2_liquide) = 0.27

    'Facteurs d'émissions pour le H2O2   kgCO2/T   source: Eco-Invent v2.2    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_H2O2) = 1145
    FE_reactifs_incertitude(reactifs_repere_H2O2) = 0.16

    'Facteurs d'émissions pour le charbon actif  kgCO2/T   source: ASTEE 5.2010.2    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_charbon_actif) = 7000
    FE_reactifs_incertitude(reactifs_repere_charbon_actif) = 0.1

    'Facteurs d'émissions pour le CuSO4   kgCO2/T   source: PAS D'INFO POUR L'INSTANT
    hyp_FE_reactifs(reactifs_repere_CuSO4) = 0
    FE_reactifs_incertitude(reactifs_repere_H2O2) = 0.5

    'Facteurs d'émissions pour le MgCl2   kgCO2/T   source: PAS D'INFO POUR L'INSTANT
    hyp_FE_reactifs(reactifs_repere_MgCl2) = 0
    FE_reactifs_incertitude(reactifs_repere_H2O2) = 0.5

    'Facteurs d'émissions pour le Cl2 gazeux pur  kgCO2/T   source: Carbone 4    MAJ 08/11/2012
    hyp_FE_reactifs(reactifs_repere_Cl2_gazeux) = 947.5
    FE_reactifs_incertitude(reactifs_repere_Cl2_gazeux) = 0.25

    'Facteurs d'émissions pour l'ammoniac pur (NH3)  kgCO2/T   source: Ec-Eau 2.8    Bilan produit de l'ADEME
    hyp_FE_reactifs(reactifs_repere_Ammoniaque) = 2110
    FE_reactifs_incertitude(reactifs_repere_Ammoniaque) = 0.5

    'Facteurs d'émissions pour l'urée pure  kgCO2/T   source: Ec-Eau 2.8    Bilan produit de l'ADEME
    hyp_FE_reactifs(reactifs_repere_Uree) = 1230
    FE_reactifs_incertitude(reactifs_repere_Uree) = 0.5
    
    'densité boues et graisses
    densite_boues = 1
    densite_graisses = 1

    'émissions références de CO2    kgCO2/(EH.an)
    emissions_CO2_EH_an_max = 75   'kgCO2/voyageur pour vol aller simple Paris-Toulouse source: AIR FRANCE


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "33.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub



Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer


    electricite_verte_consommee_force = Feuil33.Range("CO2_electricite_verte_consommee").Cells(1, pi_FiliereConsideree)
    If electricite_verte_consommee_force = True Then
        electricite_verte_consommee_f = Feuil33.Range("CO2_electricite_verte_consommee").Cells(1, pi_FiliereConsideree)
    End If
    electricite_verte_vendue_force = Feuil33.Range("CO2_electricite_verte_vendue").Cells(1, pi_FiliereConsideree)
    If electricite_verte_vendue_force = True Then
        electricite_verte_vendue_f = Feuil33.Range("CO2_electricite_verte_vendue").Cells(1, pi_FiliereConsideree)
    End If

    FE_electricite_force = Feuil33.Range("CO2_FE_electricite").Cells(1, pi_FiliereConsideree)
    If FE_electricite_force = True Then
        FE_electricite_f = Feuil33.Range("CO2_FE_electricite").Cells(1, pi_FiliereConsideree)
    End If

    For I = 1 To reactifs_nb
        FE_reactifs_force(I) = Feuil33.Range("CO2_FE_reactifs").Cells(I, pi_FiliereConsideree)
        If FE_reactifs_force(I) = True Then
            FE_reactifs_f(I) = Feuil33.Range("CO2_FE_reactifs").Cells(I, pi_FiliereConsideree)
        End If
        FE_transport_reactifs_force(I) = Feuil33.Range("CO2_FE_transport_reactifs").Cells(I, pi_FiliereConsideree)
        If FE_transport_reactifs_force(I) = True Then
            FE_transport_reactifs_f(I) = Feuil33.Range("CO2_FE_transport_reactifs").Cells(I, pi_FiliereConsideree)
        End If
    Next I

    FE_gaz_naturel_force = Feuil33.Range("CO2_FE_gaz_naturel").Cells(1, pi_FiliereConsideree)
    If FE_gaz_naturel_force = True Then
        FE_gaz_naturel_f = Feuil33.Range("CO2_FE_gaz_naturel").Cells(1, pi_FiliereConsideree)
    End If
    FE_fioul_force = Feuil33.Range("CO2_FE_fioul").Cells(1, pi_FiliereConsideree)
    If FE_fioul_force = True Then
        FE_fioul_f = Feuil33.Range("CO2_FE_fioul").Cells(1, pi_FiliereConsideree)
    End If

    volume_boues_force = Feuil33.Range("CO2_volume_boues").Cells(1, pi_FiliereConsideree)
    If volume_boues_force = True Then
        volume_boues_f = Feuil33.Range("CO2_volume_boues").Cells(1, pi_FiliereConsideree)
    End If
    FE_transport_boues_force = Feuil33.Range("CO2_FE_transport_boues").Cells(1, pi_FiliereConsideree)
    If FE_transport_boues_force = True Then
        FE_transport_boues_f = Feuil33.Range("CO2_FE_transport_boues").Cells(1, pi_FiliereConsideree)
    End If
    mn_QuantiteRefusDegrillage_Tj_force = Feuil33.Range("CO2_QuantiteRefusDegrillage_Tj").Cells(1, pi_FiliereConsideree)
    If mn_QuantiteRefusDegrillage_Tj_force = True Then
        md_QuantiteRefusDegrillage_Tj_f = Feuil33.Range("CO2_QuantiteRefusDegrillage_Tj").Cells(1, pi_FiliereConsideree)
    End If
    mn_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm_force = Feuil33.Range("CO2_FE_transport_RefusDegrillage").Cells(1, pi_FiliereConsideree)
    If mn_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm_force = True Then
        md_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm_f = Feuil33.Range("CO2_FE_transport_RefusDegrillage").Cells(1, pi_FiliereConsideree)
    End If
    graisses_Tj_force = Feuil33.Range("CO2_graisses_Tj").Cells(1, pi_FiliereConsideree)
    If graisses_Tj_force = True Then
        graisses_Tj_f = Feuil33.Range("CO2_graisses_Tj").Cells(1, pi_FiliereConsideree)
    End If
    FE_transport_graisses_force = Feuil33.Range("CO2_FE_transport_graisses").Cells(1, pi_FiliereConsideree)
    If FE_transport_graisses_force = True Then
        FE_transport_graisses_f = Feuil33.Range("CO2_FE_transport_graisses").Cells(1, pi_FiliereConsideree)
    End If
    cendres_Tj_force = Feuil33.Range("CO2_cendres_Tj").Cells(1, pi_FiliereConsideree)
    If cendres_Tj_force = True Then
        cendres_Tj_f = Feuil33.Range("CO2_cendres_Tj").Cells(1, pi_FiliereConsideree)
    End If
    FE_transport_cendres_force = Feuil33.Range("CO2_FE_transport_cendres").Cells(1, pi_FiliereConsideree)
    If FE_transport_cendres_force = True Then
        FE_transport_cendres_f = Feuil33.Range("CO2_FE_transport_cendres").Cells(1, pi_FiliereConsideree)
    End If
    REFIB_Tj_force = Feuil33.Range("CO2_REFIB_Tj").Cells(1, pi_FiliereConsideree)
    If REFIB_Tj_force = True Then
        REFIB_Tj_f = Feuil33.Range("CO2_REFIB_Tj").Cells(1, pi_FiliereConsideree)
    End If
    FE_transport_REFIB_force = Feuil33.Range("CO2_FE_transport_REFIB").Cells(1, pi_FiliereConsideree)
    If FE_transport_REFIB_force = True Then
        FE_transport_REFIB_f = Feuil33.Range("CO2_FE_transport_REFIB").Cells(1, pi_FiliereConsideree)
    End If



    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "33.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub

Private Sub Class_Initialize()


    'on définit les constantes à utiliser
    nb_compartiments = reactifs_nb + combustibles_nb + CO2_autres_nb
    gaz_naturel = reactifs_nb + combustible_repere_gaz_naturel
    fioul = reactifs_nb + combustible_repere_fioul
    electricite = reactifs_nb + combustibles_nb + CO2_autres_repere_electricite
    transport = reactifs_nb + combustibles_nb + CO2_autres_repere_transport

    ReDim emissions_repartition(nb_compartiments) As Double

    If pn_EffacementResultatsEnCours = False Then
        Call hypothese
        Call lecture_choix
        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
            Call lecture_valeurs_forcees
        End If
        Call attribution_valeur_par_defaut
    End If


End Sub

Sub calcul(ByVal electricite_consommee, ByVal electricite_verte, ByVal polymere_flux, ByVal chaux_eteinte_flux, ByVal chaux_vive_flux, ByVal FeCl3_flux, ByVal methanol_flux, ByVal H2SO4pur_flux, ByVal NaOHpur_flux, ByVal NaOClpur_flux, ByVal NaHSO3pur_flux, ByVal C_elimine, ByVal N_elimine, ByVal P_elimine, ByVal boues_evacuees_Q, ByVal graisses_evacuees_Q, ByVal gaz_naturel_kWhPCIj, ByVal fioul_kWhPCIj, ByVal NaHCO3_flux, ByVal Ca_2NO3_pur_flux As Double, ByVal O2_liquide_flux, ByVal H2O2pur_flux, ByVal charbon_actif_flux, ByVal Cl2_gazeux_flux_kgj, ByVal cendres_evacuees_Tj, ByVal REFIB_evacues_Tj, ByVal Struvite_kgj As Double, ByVal QuantiteRefusDegrillage_kgj As Double, ByVal ConsommationAmmoniaquePur_kgj As Double, ByVal ConsommationUreePur_kgj As Double)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim reactifs_flux(reactifs_nb) As Double
    Dim I As Integer
    Dim j As Integer

    'CALCUL DE L'EMPREINTE CO2
    emissions = 0
    emissions_incertitude = 0
    For I = 1 To nb_compartiments
        emissions_repartition(I) = 0
    Next I

    'pour les OPEX
    StruviteProduite_kgj = Struvite_kgj

    'flux de réactifs en T/an
    reactifs_flux(reactifs_repere_polymere_eau_poudre_anionique) = polymere_flux / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_FeCl3_eau) = FeCl3_flux / FeCl3_pureté / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_methanol) = methanol_flux / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_chaux_eteinte) = chaux_eteinte_flux / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_chaux_vive) = chaux_vive_flux / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_H2SO4) = H2SO4pur_flux / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_NaOH) = NaOHpur_flux / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_NaHCO3) = NaHCO3_flux / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_NaOCl_desodo) = NaOClpur_flux / NaOCl_pureté / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_NaHSO3) = NaHSO3pur_flux / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_Ca_2NO3) = Ca_2NO3_pur_flux / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_O2_liquide) = O2_liquide_flux / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_H2O2) = H2O2pur_flux / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_charbon_actif) = charbon_actif_flux / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_Cl2_gazeux) = Cl2_gazeux_flux_kgj / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_Ammoniaque) = ConsommationAmmoniaquePur_kgj / 1000 * pd_NOMBRE_JOUR_PAR_AN
    reactifs_flux(reactifs_repere_Uree) = ConsommationUreePur_kgj / 1000 * pd_NOMBRE_JOUR_PAR_AN


    'calcul de l'empreinte CO2 des réactifs (production+transport)
    For I = 1 To reactifs_nb
        j = choix_transport_reactifs(I)
        emissions = emissions + reactifs_flux(I) * FE_reactifs(I) / 1000
        emissions_incertitude = emissions_incertitude + FE_reactifs_incertitude(I) * reactifs_flux(I) * FE_reactifs(I) / 1000
        emissions = emissions + reactifs_flux(I) * FE_transport_reactifs(I) * transport_reactifs_distance(I) / 1000
        emissions_incertitude = emissions_incertitude + FE_transport_incertitude(j) * reactifs_flux(I) * FE_transport_reactifs(I) * transport_reactifs_distance(I) / 1000
        emissions_repartition(I) = emissions_repartition(I) + reactifs_flux(I) * FE_reactifs(I) / 1000
        emissions_repartition(transport) = emissions_repartition(transport) + reactifs_flux(I) * FE_transport_reactifs(I) * transport_reactifs_distance(I) / 1000
    Next I

    'empreinte CO2 du transport des boues
    'DCA 01/07/11: on ne change pas la notation mais on reporte en tonnes dans la synthèse
    If volume_boues_force = True Then
        volume_boues = volume_boues_f
    Else
        volume_boues = boues_evacuees_Q
    End If
    emissions = emissions + volume_boues * pd_NOMBRE_JOUR_PAR_AN * densite_boues * FE_transport_boues * transport_boues_distance / 1000
    emissions_incertitude = emissions_incertitude + FE_transport_boues_incertitude * volume_boues * pd_NOMBRE_JOUR_PAR_AN * densite_boues * FE_transport_boues * transport_boues_distance / 1000
    emissions_repartition(transport) = emissions_repartition(transport) + volume_boues * pd_NOMBRE_JOUR_PAR_AN * densite_boues * FE_transport_boues * transport_boues_distance / 1000
    'transport des graisses
    If graisses_Tj_force = True Then
        graisses_Tj = graisses_Tj_f
    Else
        graisses_Tj = graisses_evacuees_Q * densite_graisses
    End If
    emissions = emissions + graisses_Tj * pd_NOMBRE_JOUR_PAR_AN * FE_transport_graisses * transport_graisses_distance / 1000
    emissions_incertitude = emissions_incertitude + FE_transport_graisses_incertitude * graisses_Tj * pd_NOMBRE_JOUR_PAR_AN * FE_transport_graisses * transport_graisses_distance / 1000
    emissions_repartition(transport) = emissions_repartition(transport) + graisses_Tj * pd_NOMBRE_JOUR_PAR_AN * FE_transport_graisses * transport_graisses_distance / 1000
    'transport des refus de dégrillage
    If mn_QuantiteRefusDegrillage_Tj_force = True Then
        md_QuantiteRefusDegrillage_Tj = md_QuantiteRefusDegrillage_Tj_f
    Else
        md_QuantiteRefusDegrillage_Tj = QuantiteRefusDegrillage_kgj / 1000
    End If
    emissions = emissions + md_QuantiteRefusDegrillage_Tj * pd_NOMBRE_JOUR_PAR_AN * md_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm * md_DistanceTransportRefusDegrillage_km / 1000
    emissions_incertitude = emissions_incertitude + md_IncertitudeFacteurEmissionTransportRefusDegrillage * md_QuantiteRefusDegrillage_Tj * pd_NOMBRE_JOUR_PAR_AN * md_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm * md_DistanceTransportRefusDegrillage_km / 1000
    emissions_repartition(transport) = emissions_repartition(transport) + md_QuantiteRefusDegrillage_Tj * pd_NOMBRE_JOUR_PAR_AN * md_FacteurEmissionTransportRefusDegrillage_kgCO2_Tkm * md_DistanceTransportRefusDegrillage_km / 1000
    'transport des cendres
    If cendres_Tj_force = True Then
        cendres_Tj = cendres_Tj_f
    Else
        cendres_Tj = cendres_evacuees_Tj
    End If
    emissions = emissions + cendres_Tj * pd_NOMBRE_JOUR_PAR_AN * FE_transport_cendres * transport_cendres_distance / 1000
    emissions_incertitude = emissions_incertitude + FE_transport_cendres_incertitude * cendres_Tj * pd_NOMBRE_JOUR_PAR_AN * FE_transport_cendres * transport_cendres_distance / 1000
    emissions_repartition(transport) = emissions_repartition(transport) + cendres_Tj * pd_NOMBRE_JOUR_PAR_AN * FE_transport_cendres * transport_cendres_distance / 1000
    'transport des REFIB
    If REFIB_Tj_force = True Then
        REFIB_Tj = REFIB_Tj_f
    Else
        REFIB_Tj = REFIB_evacues_Tj
    End If
    emissions = emissions + REFIB_Tj * pd_NOMBRE_JOUR_PAR_AN * FE_transport_REFIB * transport_REFIB_distance / 1000
    emissions_incertitude = emissions_incertitude + FE_transport_REFIB_incertitude * REFIB_Tj * pd_NOMBRE_JOUR_PAR_AN * FE_transport_REFIB * transport_REFIB_distance / 1000
    emissions_repartition(transport) = emissions_repartition(transport) + REFIB_Tj * pd_NOMBRE_JOUR_PAR_AN * FE_transport_REFIB * transport_REFIB_distance / 1000



    'électricité
    emissions = emissions + FE_electricite * electricite_consommee * pd_NOMBRE_JOUR_PAR_AN / 1000
    emissions_incertitude = emissions_incertitude + FE_electricite_incertitude * FE_electricite * electricite_consommee * pd_NOMBRE_JOUR_PAR_AN / 1000
    emissions_repartition(electricite) = emissions_repartition(electricite) + FE_electricite * electricite_consommee * pd_NOMBRE_JOUR_PAR_AN / 1000

    'gaz naturel
    emissions = emissions + FE_gaz_naturel * gaz_naturel_kWhPCIj * pd_NOMBRE_JOUR_PAR_AN / 1000
    emissions_incertitude = emissions_incertitude + FE_gaz_naturel_incertitude * FE_gaz_naturel * gaz_naturel_kWhPCIj * pd_NOMBRE_JOUR_PAR_AN / 1000
    emissions_repartition(gaz_naturel) = emissions_repartition(gaz_naturel) + FE_gaz_naturel * gaz_naturel_kWhPCIj * pd_NOMBRE_JOUR_PAR_AN / 1000

    'fioul
    emissions = emissions + FE_fioul * fioul_kWhPCIj * pd_NOMBRE_JOUR_PAR_AN / 1000
    emissions_incertitude = emissions_incertitude + FE_fioul_incertitude * FE_fioul * fioul_kWhPCIj * pd_NOMBRE_JOUR_PAR_AN / 1000
    emissions_repartition(fioul) = emissions_repartition(fioul) + FE_fioul * fioul_kWhPCIj * pd_NOMBRE_JOUR_PAR_AN / 1000

    'émissions évitées/réduites
    If electricite_verte_consommee_force = True Then
        electricite_verte_consommee = electricite_verte_consommee_f
        electricite_verte_vendue = electricite_verte - electricite_verte_consommee
    ElseIf electricite_verte_vendue_force = True Then
        electricite_verte_vendue = electricite_verte_vendue_f
        electricite_verte_consommee = electricite_verte - electricite_verte_vendue
    Else
        If choix_valorisation = used_on_site Then
            If electricite_verte > electricite_consommee Then
                electricite_verte_consommee = electricite_consommee
                electricite_verte_vendue = electricite_verte - electricite_consommee
            Else
                electricite_verte_consommee = electricite_verte
                electricite_verte_vendue = 0
            End If
        Else
            electricite_verte_consommee = 0
            electricite_verte_vendue = electricite_verte
        End If
    End If
    reduites = FE_electricite * electricite_verte_consommee * pd_NOMBRE_JOUR_PAR_AN / 1000
    If reduites <> 0 Then
        evitees_reduites_incertitude = FE_electricite_incertitude
    End If
    emissions_repartition(electricite) = emissions_repartition(electricite) - reduites
    evitees = FE_electricite * electricite_verte_vendue * pd_NOMBRE_JOUR_PAR_AN / 1000
    If evitees <> 0 Then
        evitees_reduites_incertitude = FE_electricite_incertitude
    End If

    'on calcule l'incertitude en %
    If emissions <> 0 Then
        emissions_incertitude = emissions_incertitude / emissions
        For I = 1 To nb_compartiments
            emissions_repartition(I) = emissions_repartition(I) / (emissions - reduites)
            If emissions_repartition(I) < 10 ^ (-4) Then
                emissions_repartition(I) = 0
            End If
        Next I
    End If


    'CALCUL DES RATIOS TYPES

    'ratio par équivalent-habitant et par jour
    If choix_primaire = False And choix_secondaire = False And choix_tertiaire = False Then
        ratio_EH = ((emissions - reduites) * 10 ^ 3) / (pd_CapaciteSTEP_EH * pd_PourcentageChargeReelleDBO)
        ratio_Q = 0
        ratio_CNP = 0
    Else
        ratio_EH = ((emissions - reduites) * 10 ^ 3) / (pd_CapaciteSTEP_EH * pd_PourcentageChargeReelleDBO)
        ratio_Q = ((emissions - reduites) / pd_NOMBRE_JOUR_PAR_AN * 10 ^ 6) / (Q_nominal * pd_PourcentageChargeReelleDebitVolumique)
        ratio_CNP = ((emissions - reduites) / pd_NOMBRE_JOUR_PAR_AN * 10 ^ 6) / (C_elimine + N_elimine + P_elimine)
    End If

    'fourchettes de ratios
    ratio_fourchette_Q = "0 - 2000"
    ratio_fourchette_CNP = "0 - 2000"

    'calcul des valeurs pour demi-camembert
    zone(1) = ((emissions - evitees) * 10 ^ 3) / (pd_CapaciteSTEP_EH * pd_PourcentageChargeReelleDBO)
    If zone(1) < 0 Then
        zone(1) = 0
    End If
    zone(2) = ratio_EH - zone(1)
    If zone(2) < 0 Then
        zone(2) = 0
    End If
    zone(3) = emissions_CO2_EH_an_max - (zone(1) + zone(2))
    If zone(3) < 0 Then
        zone(3) = 0
    End If
    zone(4) = 2 * emissions_CO2_EH_an_max - (zone(1) + zone(2) + zone(3))


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "33.12"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub


Private Sub Class_Terminate()


    If pn_EffacementResultatsEnCours = False Then
        Call ecriture_resultats
    End If


End Sub








