Option Explicit

Public pn_ErreurNonGeree As Boolean

Public flag_calcul_retour_investissement As Boolean

'constantes utilisées dans les modules
'étapes de calcul (initialisation ou calcul)
Public Const pi_CALCUL_VALEUR_GUIDE = 1
Public Const pi_CALCUL_VALEUR_EFFECTIVE = 2
Public Const pi_NOMBRE_ITERATION_RETOURS = 15

'constantes usuelles
Public Const pd_NOMBRE_HEURE_PAR_JOUR As Double = 24
Public Const pd_NOMBRE_SECONDE_PAR_HEURE As Double = 3600
Public Const pd_NOMBRE_MINUTE_PAR_HEURE As Double = 60
Public Const pd_ACCELERATION_PESANTEUR_m_s2 As Double = 9.81
Public Const pd_NOMBRE_JOUR_PAR_AN As Double = 365
Public Const pd_NOMBRE_JOUR_PAR_SEMAINE As Double = 7
Public Const pd_NOMBRE_MOIS_PAR_AN As Double = 12
Public Const pd_CONVERSION_kJ_PAR_kcal As Double = 4.185
Public Const pd_CONVERSION_K_°C As Double = 273.15
Public Const pd_CONVERSION_L_PAR_m3 As Double = 1000
Public Const pd_REJET_DBO_PAR_HABITANT_ET_PAR_JOUR As Double = 0.06  'kg DBO / (EH.j)
Public Const pd_NOMBRE_PI As Double = 3.14159265
Public Const pd_PCI_GAZ_NATUREL_kWh_Nm3 As Double = 10
Public Const pd_PCI_FIOUL_kWh_L As Double = 9.96     'Delphine NAWAWI
Public Const pd_PCI_CH4_kWh_Nm3 As Double = 9.9
Public Const pd_QUANTITE_OXYGENE_DANS_AIR_kgO2_Nm3 As Double = 0.3

'Masses molaires
Public Const pd_MASSE_MOLAIRE_H_kg_mol = 1 / 1000
Public Const pd_MASSE_MOLAIRE_C_kg_mol = 12 / 1000
Public Const pd_MASSE_MOLAIRE_N_kg_mol = 14 / 1000
Public Const pd_MASSE_MOLAIRE_O_kg_mol = 16 / 1000
Public Const pd_MASSE_MOLAIRE_Na_kg_mol = 23 / 1000
Public Const pd_MASSE_MOLAIRE_S_kg_mol = 32 / 1000
Public Const pd_MASSE_MOLAIRE_Cl_kg_mol = 35.5 / 1000
Public Const pd_MASSE_MOLAIRE_Mg_kg_mol = 24.3 / 1000
Public Const pd_MASSE_MOLAIRE_P_kg_mol = 31 / 1000
Public Const pd_MASSE_MOLAIRE_Ca_kg_mol = 40 / 1000
Public Const pd_MASSE_MOLAIRE_Fe_kg_mol = 55.8 / 1000

'constantes pour le choix Biothélys / Exelys
Public Const pi_CONFIGURATION_BIOTHELYS As Integer = 1
Public Const pi_CONFIGURATION_EXELYS As Integer = 2

'types d'eau
Public type_eau_nominal As Integer
Public type_eau_reel As Integer
Public Const eau_diluee = 1
Public Const eau_standard = 2
Public Const eau_concentree = 3

'constantes pour le calcul de la file boues
Public Const nb_etape_file_boues = 23    'nombre d'étapes de traitement sur la file boues
Public Const boues_inlet = 1
Public boues_epaissies(5) As Integer    'VERIFIER LES CONSTANTES EN CAS DE CHANGEMENT (prgm ppal)
Public Const boues_digerees = 7
Public Const boues_athos = 8
Public boues_deshydratees(3) As Integer
Public boues_chaulees(3) As Integer
Public boues_sechees(2) As Integer
Public boues_sechees_bioco(2) As Integer
Public boues_sechees_inos(2) As Integer
Public Const boues_incinerees = 21
Public Const boues_graisses_biolix = 22
Public Const boues_evacuees = 23

Public Const nb_type_boues = 6   'nombre de types de boues en entrée de la filière
Public Const boues_I = 1
Public Const boues_II = 2
Public Const boues_III = 3
Public Const boues_externes_1 = 4
Public Const boues_externes_2 = 5
Public Const graisses = 6

Public Const nb_qualites_boues = 19
'PRIMAIRE
Public Const I_simple = 1
Public Const I_reactif = 2
'SECONDAIRE
Public Const II_forte = 3
Public Const II_moyenne = 4
Public Const II_faible_EB = 5
Public Const II_faible_ED = 6
Public Const II_prolongee_EB = 7
Public Const II_prolongee_ED = 8
Public Const II_MBR = 9
Public Const II_MBBR = 10
Public Const II_HybAS = 11
Public Const II_biostyr_C = 12
Public Const II_biostyr_N = 13
Public Const II_biostyr_NDN = 14
Public Const II_biostyr_PDN = 15
'TERTIAIRE
Public Const III_decantation = 16
Public Const III_biostyr_N = 17
Public Const III_biostyr_PDN = 18
'GRAISSES
Public Const codigestion_graisses = 19

Public Const nb_parametres_boues = 10   'process_origine,MES,Q,MV/MES,verif_flux,flux_in,DCO/MES,DBO/MES,NK/MES,Pt/MES
Public Const repere_origine = 1     'nombre de process dans la file eau max par étape de traitement
Public Const repere_MES = 2
Public Const repere_Q = 3
Public Const repere_MV_MES = 4     'en réalité MV/MS
Public Const repere_verif_flux = 5
Public Const repere_flux_in = 6
Public Const repere_ratio_DCO_MES = 7
Public Const repere_ratio_DBO_MES = 8
Public Const repere_ratio_NK_MES = 9
Public Const repere_ratio_Pt_MES = 10

'paramètres pour ajuster la siccité pour avoir des boues autocombustibles dans l'incinération
Public nb_boues As Integer
Public incineration_repere_boues_inlet(2) As Integer
Public incineration_siccite_ajustable(2) As Boolean
Public incineration_siccite_autocombustibilite(2) As Double

'repère pour les boues I (cas du retour des eaux sales vers primaire)
Public Const nb_parametres_boues_I_detail = 3
'idem boues pour le 1  (origine des boues)
Public Const repere_ratio_MES = 2
Public Const repere_ratio_MV = 3
'repère pour les pollutions solubles des boues digérées
Public Const nb_repere_mgL = 5
Public Const repere_mgL_DCO = 1
Public Const repere_mgL_NK = 2
Public Const repere_mgL_Pt = 3
Public Const repere_mgL_MS_soluble = 4
Public Const repere_mgL_MV_soluble = 5
'repère pour les paramètres des graisses internes   (MS=MES hypothèse acceptable selon Eric GUIBELIN)
Public Const repere_graisse_MS = 1
Public Const repere_graisse_Q = 2
Public Const repere_graisse_MV_MES = 3
Public Const repere_graisse_DCO_MES = 4
Public Const repere_graisse_DBO_MES = 5
Public Const repere_graisse_NK_MES = 6
Public Const repere_graisse_Pt_MES = 7
'repère co-ferments pour lecture et écriture dans Excel
'public const repere_coferment_choix = 1
'MSM 15/11/12 coferment
Public Const repere_coferment_Q_m3j = 2
Public Const repere_coferment_MS_gL = 3
Public Const repere_coferment_ratio_MES_MS = 4
Public Const repere_coferment_ratio_MV_MS = 5
Public Const repere_coferment_ratio_DCO_MV = 6
Public Const repere_coferment_ratio_NK_MV = 7
Public Const repere_coferment_ratio_Pt_MV = 8
Public Const repere_coferment_BMP_Nm3CH4_tMVapplique = 9
Public Const repere_coferment_BMP_securite = 10

'repère pour les détails des graisses (part des co-ferments particulaires par rapport aux graisses internes)
Public Const nb_parametres_graisses_particulaire_detail = 2
Public Const repere_graisse_particulaire_ratioMES = 1
Public Const repere_graisse_particulaire_ratioMV = 2

'constantes pour les caractéristques des retours
Public Const retour_caracteristique_nb = 8
Public Const repere_ret_Q = 1
Public Const repere_ret_DCO = 2
Public Const repere_ret_DBO = 3
Public Const repere_ret_MES = 4
Public Const repere_ret_NK = 5
Public Const repere_ret_NH4 = 6
Public Const repere_ret_NO3 = 7
Public Const repere_ret_Pt = 8
'constantes pour la gestion des eaux sales
'repères pour le fux d'eau sale considéré  (taille matrices gérées par nb_eaux_sales_max)
Public Const nb_eaux_sales_max = 2
Public Const repere_ES_biostyr = 1    'secondaire ou tertiaire (hors post-dénitrifiaction)
Public Const repere_ES_biostyr_PDN = 2    'post-DN secondaire ou tertiaire
'paramètres eaux sales
Public Const nb_parametres_eaux_sales = 10
Public Const repere_ES_origine = 1
Public Const repere_ES_Q = 2
Public Const repere_ES_MES = 3
Public Const repere_ES_MV_MES = 4
Public Const repere_ES_DCO = 5
Public Const repere_ES_DBO = 6
Public Const repere_ES_NH4 = 7
Public Const repere_ES_NK = 8
Public Const repere_ES_NO3 = 9
Public Const repere_ES_Pt = 10
'constantes pour le choix de devenir des eaux sales
Public Const nb_devenir_ES = 3
Public Const devenir_ES_primaire = 1   'eaux sales retournées en tête du traitement primaire (prohibé sur le nitrifiant car NO3 trop élevé --> problème de décantation)
Public Const devenir_ES_traitement_separe_amont = 2   'eaux sales traitées puis retournées en tête de biostyr
Public Const devenir_ES_traitement_separe_aval = 3   'eaux sales traitées puis rejetées au milieu ou envoyées après le biostyr (pas de norme sur l'azote)

'gestion énergie thermique process
Public Const nb_niveaux_energie = 6
Public Const energie_eau_BT = 1
Public Const energie_eau_HT = 2
Public Const energie_vapeur_sat1 = 3
Public Const energie_vapeur_sat2 = 4
Public Const energie_biogaz = 5    ' besoins direct biogaz à partir du sécheur + pyrofluid
Public Const energie_combustible = 6    'besoins direct combustible à partir du sécheur + pyrofluid

Public Const nb_vapeur_max = 2
Public Const vapeur_sat1 = 1
Public Const vapeur_sat2 = 2
'constantes biogaz
Public Const nb_biogaz_caracteristiques = 2
Public Const Q_Nm3j = 1
Public Const teneur_CH4 = 2

Public choix_continuer As Boolean

'variables sur le repérage de la station exécutée et de l'étape de calcul
Public pi_FiliereConsideree As Integer
Public pi_EtapeCalculConsideree As Integer
Public nb_step_comparaison As Integer
Public iteration As Integer

'variables concernant les pompes
Public Const pompe_file_eau = 1
Public Const pompe_file_boues = 2    'de type SEEPEX

'variable commentaires sur les calculs
Public ps_CommentairesCalculsGlobaux As String
Public pn_EffacementResultatsEnCours As Boolean
Public calcul_effectue(3, 2) As Boolean   '(pi_FiliereConsideree,etape de calcul)

'variables du cahier des charges
Public pd_CapaciteSTEP_EH As Long
Public Q_nominal As Double
Public DCO_nominal As Double
Public DBO_nominal As Double
Public MES_nominal As Double
Public NK_nominal As Double
Public NH4_nominal As Double
Public Pt_nominal As Double
Public SH_nominal As Double
Public pd_TemperatureEauDimensionnement_°C As Double

Public vidange_Q_nominal As Double
Public vidange_DCO_nominal As Double
Public vidange_DBO_nominal As Double
Public vidange_MES_nominal As Double
Public vidange_NK_nominal As Double
Public vidange_NH4_nominal As Double
Public vidange_Pt_nominal As Double

Public DCO_garantie As Double
Public DBO_garantie As Double
Public MES_garantie As Double
Public NGL_garantie As Double
Public NK_garantie As Double
Public Pt_garantie As Double

'variables des conditions d'exploitation
Public pd_PourcentageChargeReelleDebitVolumique As Double
Public pd_PourcentageChargeReelleDCO As Double
Public pd_PourcentageChargeReelleDBO As Double
Public pd_PourcentageChargeReelleMES As Double
Public pd_PourcentageChargeReelleNK As Double
Public pd_PourcentageChargeReelleNH4 As Double
Public pd_PourcentageChargeReellePt As Double
Public pd_PourcentageChargeReelleSH As Double

Public vidange_Q_reel As Double
Public vidange_DCO_reel As Double
Public vidange_DBO_reel As Double
Public vidange_MES_reel As Double
Public vidange_NK_reel As Double
Public vidange_NH4_reel As Double
Public vidange_Pt_reel As Double

Public pd_CoefficientPointeHydrauliqueTempsSec As Double
Public pd_CoefficientPointeHydrauliqueTempsPluie As Double
Public pd_TemperatureEauConditionsReelles_°C As Double
Public T_air_aspire As Double
Public humidite_air As Double
Public altitude As Double

'A modifier !! MSM 26/10/12 : faire un test en Fct de la temp (BD climat ou forcée) puis déterminer le type de climat.
Public choix_climat As Integer    ' DEFINI EN FONCTION DE LA TEMPERATURE DE L'AIR DANS lecture_donnees_generales
Public Const climat_type_nb = 3
Public Const climat_froid = 1     ' -5°C en moyenne
Public Const climat_tempere = 2     ' 15°C en moyenne
Public Const climat_chaud = 3     ' 25°C en moyenne

Public pi_ChoixQualiteRejet As Integer
Public Const pi_QUALITE_REJET_TRAITEMENT_CARBONE = 1     'traitement carbone seul
Public Const pi_QUALITE_REJET_ZONE_SENSIBLE = 2     ' zone sensible
Public Const pi_QUALITE_REJET_ZONE_TRES_SENSIBLE = 3     'zone très sensible

'variables de choix de traitement
Public choix_traitement_sulfures_preventif As Boolean
Public choix_traitement_sulfures_curatif As Boolean
Public choix_relevement As Boolean
Public choix_pretraitement As Boolean
Public choix_degrillage As Boolean
Public choix_dessablage As Boolean
Public choix_primaire As Boolean
Public choix_decanteur_simple As Boolean
Public choix_decanteur_reactif As Boolean
Public choix_secondaire As Boolean
Public choix_BA_forte As Boolean
Public choix_BA_moyenne As Boolean
Public choix_BA_faible As Boolean
Public choix_BA_prolongee As Boolean
Public choix_HybAS As Boolean
Public choix_MBBR As Boolean
Public choix_MBR As Boolean
Public choix_biostyr As Boolean
Public choix_biostyr_PDN As Boolean
Public choix_tertiaire As Boolean
Public choix_biostyr_N_III As Boolean
Public choix_biostyr_PDN_III As Boolean
Public choix_decanteur_III As Boolean
Public choix_discfilter As Boolean
Public choix_filtrasable As Boolean
Public choix_desinfection As Boolean
Public choix_UV As Boolean
Public choix_chloration As Boolean
Public choix_decanteur_ES As Boolean
Public choix_biolix As Boolean
Public choix_epaississement As Boolean
Public pi_ChoixDigestion As Boolean
Public choix_dig_simple As Boolean
Public choix_biothelys As Boolean
Public choix_exelys_DLD As Boolean
Public choix_athos As Boolean
Public choix_deshydratation As Boolean
Public choix_chaulage As Boolean
Public choix_sechage As Boolean
Public choix_secheur_autre As Boolean
Public choix_secheur_bioco As Boolean
Public choix_secheur_inos As Boolean
Public choix_incineration As Boolean
Public choix_traitement_retours As Boolean
Public choix_MAP_retours As Boolean
Public choix_ANITA_Mox As Boolean
Public choix_ANITA_Shunt As Boolean
Public choix_stripping_N As Boolean
Public choix_desodo_chimique As Boolean
Public choix_desodo_bio As Boolean
Public choix_utilites_pompage_retours As Boolean
Public choix_utilites_eau_service As Boolean
Public choix_utilites_eclairage As Boolean
Public choix_utilites_pertes_enligne As Boolean
Public choix_utilites_chauffage_clim As Boolean
Public choix_utilites_bat_administration As Boolean
Public choix_utilites_bat_exploitation As Boolean
Public choix_utilites_bat_electrique As Boolean
Public choix_prod_alt_chaleur As Boolean
Public choix_PAC_eau_traitee As Boolean
Public choix_prod_alt_electricite As Boolean
Public choix_solaire_photovoltaique As Boolean
Public choix_turbine_hydraulique As Boolean
Public choix_electricite_autre_production As Boolean
Public choix_coferments As Boolean

'identifiant station BDClimat
Public idStation As Integer

'constantes pour repérer les réactifs, les combustibles, les autres postes pour le CO2   !!!!!!!!!!!! ATTENTION, c'est différent pour les OPEX !!!!!!!!!!!!!!
Public Const reactifs_nb = 28
Public Const reactifs_repere_polymere_eau_poudre_anionique = 1
Public Const reactifs_repere_polymere_eau_poudre_cationique = 2
Public Const reactifs_repere_polymere_eau_liquide_anionique = 3
Public Const reactifs_repere_polymere_eau_liquide_cationique = 4
Public Const reactifs_repere_polymere_boues_poudre_anionique = 5
Public Const reactifs_repere_polymere_boues_poudre_cationique = 6
Public Const reactifs_repere_polymere_boues_liquide_anionique = 7
Public Const reactifs_repere_polymere_boues_liquide_cationique = 8
Public Const reactifs_repere_FeCl3_eau = 9
Public Const reactifs_repere_FeCl3_boues = 10
Public Const reactifs_repere_methanol = 11
Public Const reactifs_repere_chaux_eteinte = 12
Public Const reactifs_repere_chaux_vive = 13
Public Const reactifs_repere_H2SO4 = 14
Public Const reactifs_repere_NaOH = 15
Public Const reactifs_repere_NaOCl_eau = 16
Public Const reactifs_repere_NaOCl_desodo = 17
Public Const reactifs_repere_NaHSO3 = 18
Public Const reactifs_repere_NaHCO3 = 19
Public Const reactifs_repere_Ca_2NO3 = 20
Public Const reactifs_repere_O2_liquide = 21
Public Const reactifs_repere_H2O2 = 22
Public Const reactifs_repere_charbon_actif = 23
Public Const reactifs_repere_CuSO4 = 24
Public Const reactifs_repere_MgCl2 = 25
Public Const reactifs_repere_Cl2_gazeux = 26
Public Const reactifs_repere_Ammoniaque = 27
Public Const reactifs_repere_Uree = 28
Public Const combustibles_nb = 2
Public Const combustible_repere_gaz_naturel = 1
Public Const combustible_repere_fioul = 2
Public Const CO2_autres_nb = 2
Public Const CO2_autres_repere_electricite = 1
Public Const CO2_autres_repere_transport = 2

Public Const OPEX_autres = 3
Public Const OPEX_autres_electricite = 1
Public Const OPEX_autres_transport_boues = 2
Public Const OPEX_autres_OPEX_additionnel = 3

'constantes pour le détail des consommations électriques
Public Const electricite_postes_nb = 19
Public Const electricite_postes_repere_relevement = 1
Public Const electricite_postes_repere_pretraitement = 2
Public Const electricite_postes_repere_primaire = 3
Public Const electricite_postes_repere_secondaire = 4
Public Const electricite_postes_repere_tertiaire = 5
Public Const electricite_postes_repere_desinfection = 6
Public Const electricite_postes_repere_eaux_sales = 7
Public Const electricite_postes_eau_nb = 7
Public Const electricite_postes_repere_biolix = 8
Public Const electricite_postes_repere_epaississement = 9
Public Const electricite_postes_repere_digestion = 10
Public Const electricite_postes_repere_athos = 11
Public Const electricite_postes_repere_deshydratation = 12
Public Const electricite_postes_repere_chaulage = 13
Public Const electricite_postes_repere_sechage = 14
Public Const electricite_postes_repere_incineration = 15
Public Const electricite_postes_repere_trait_retours = 16
Public Const electricite_postes_boues_nb = 16
Public Const electricite_postes_repere_desodorisation = 17
Public Const electricite_postes_repere_utilites = 18
Public Const electricite_postes_repere_autres = 19

'constantes pour le type de polymère
Public Const polymere_nb_type = 4
Public Const polymere_poudre_anion = 1
Public Const polymere_poudre_cation = 2
Public Const polymere_liquide_anion = 3
Public Const polymere_liquide_cation = 4

Sub afficher_com()

    Dim fin_texte_com As String
    Dim titre_com As String

    choix_langue = Feuil5.Range("choix_langue")

    If choix_langue = choix_langue_FR Then
        fin_texte_com = "FIN DES CALCULS"
        titre_com = "COMMENTAIRES SUR LES CALCULS"
    ElseIf choix_langue = choix_langue_EN Then
        fin_texte_com = "END OF CALCULATIONS"
        titre_com = "COMMENTS ON CALCULATIONS"
    ElseIf choix_langue = choix_langue_DE Then
        fin_texte_com = "DIE BERECHNUNGEN SIND FERTIG"
        titre_com = "BEMERKUNGEN ÜBER DIE KALKULATIONEN"
    Else
        fin_texte_com = "FIN DE LOS CALCULOS"
        titre_com = "COMENTARIOS SOBRE LOS CALCULOS"
    End If

    If ps_CommentairesCalculsGlobaux = "" Then
        ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil6.Range("commentaire_aucun").Cells(1, 1) + vbCrLf
    End If

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf + fin_texte_com + vbCrLf + vbCrLf

    Commentaires.Caption = titre_com
    Commentaires.affichage_commentaires.Text = ps_CommentairesCalculsGlobaux

    If flag_calcul_retour_investissement = True Then
        Commentaires.commentaires_titre.Caption = Feuil55.Range("Investissement_commentaire_titre")
    Else
        Commentaires.commentaires_titre.Caption = Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
    End If

    Commentaires.Show


End Sub

Sub ecrire_resultats_step(ByVal Q, ByVal DCO, ByVal DBO, ByVal MES, ByVal NK, ByVal NH4, ByVal NO3, ByVal Pt, ByVal boues_flux, ByVal graisses_internes, ByVal retour_flux, ByVal boues_pollution_soluble, ByVal choix_eaux_sales, ByRef eaux_sales, ByRef eaux_sales_discfilter, eaux_sales_filtrasable, ByVal QuantiteRefusDegrillage_kgj As Double)

    Dim Colonne As Integer
    Dim I As Integer

    Dim intermediaire_Q_graisse_coferment As Double
    Dim intermediaire_Q_coferment As Double
    Dim intermediaire_MS_coferments_kgj As Double
    Dim intermediaire_MV_MS_coferments As Double

    Dim ES_primaire_Q As Double
    Dim ES_primaire_DCO As Double
    Dim ES_primaire_DBO As Double
    Dim ES_primaire_MES As Double
    Dim ES_primaire_NK As Double
    Dim ES_primaire_NH4 As Double
    Dim ES_primaire_NO3 As Double
    Dim ES_primaire_Pt As Double


    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


    'écriture des retours en tête de primaire
    For I = 1 To retour_caracteristique_nb
        Feuil6.Range("retours_flux").Cells(I, Colonne) = retour_flux(I)
    Next I
    'écriture des concentration dans "Valeurs_générales"  modif DCA 18/11/12
    'Feuil6.Range("retours_concentration").Cells(repere_ret_Q, colonne) = retour_flux(repere_ret_Q)
    If retour_flux(repere_ret_Q) > 0 Then
        Feuil6.Range("retours_concentration").Cells(repere_ret_DCO, Colonne) = retour_flux(repere_ret_DCO) / retour_flux(repere_ret_Q) * 1000
        Feuil6.Range("retours_concentration").Cells(repere_ret_DBO, Colonne) = retour_flux(repere_ret_DBO) / retour_flux(repere_ret_Q) * 1000
        Feuil6.Range("retours_concentration").Cells(repere_ret_MES, Colonne) = retour_flux(repere_ret_MES) / retour_flux(repere_ret_Q) * 1000
        Feuil6.Range("retours_concentration").Cells(repere_ret_NK, Colonne) = retour_flux(repere_ret_NK) / retour_flux(repere_ret_Q) * 1000
        Feuil6.Range("retours_concentration").Cells(repere_ret_NH4, Colonne) = retour_flux(repere_ret_NH4) / retour_flux(repere_ret_Q) * 1000
        Feuil6.Range("retours_concentration").Cells(repere_ret_NO3, Colonne) = retour_flux(repere_ret_NO3) / retour_flux(repere_ret_Q) * 1000
        Feuil6.Range("retours_concentration").Cells(repere_ret_Pt, Colonne) = retour_flux(repere_ret_Pt) / retour_flux(repere_ret_Q) * 1000
    Else
        Feuil6.Range("retours_concentration").Cells(repere_ret_DCO, Colonne) = "'-"
        Feuil6.Range("retours_concentration").Cells(repere_ret_DBO, Colonne) = "'-"
        Feuil6.Range("retours_concentration").Cells(repere_ret_MES, Colonne) = "'-"
        Feuil6.Range("retours_concentration").Cells(repere_ret_NK, Colonne) = "'-"
        Feuil6.Range("retours_concentration").Cells(repere_ret_NH4, Colonne) = "'-"
        Feuil6.Range("retours_concentration").Cells(repere_ret_NO3, Colonne) = "'-"
        Feuil6.Range("retours_concentration").Cells(repere_ret_Pt, Colonne) = "'-"
    End If

    'début modif MSM 15/11/12
    ES_primaire_Q = eaux_sales_discfilter(repere_ES_Q) + eaux_sales_filtrasable(repere_ES_Q)
    ES_primaire_DCO = eaux_sales_discfilter(repere_ES_DCO) + eaux_sales_filtrasable(repere_ES_DCO)
    ES_primaire_DBO = eaux_sales_discfilter(repere_ES_DBO) + eaux_sales_filtrasable(repere_ES_DBO)
    ES_primaire_MES = eaux_sales_discfilter(repere_ES_MES) + eaux_sales_filtrasable(repere_ES_MES)
    ES_primaire_NK = eaux_sales_discfilter(repere_ES_NK) + eaux_sales_filtrasable(repere_ES_NK)
    ES_primaire_NH4 = eaux_sales_discfilter(repere_ES_NH4) + eaux_sales_filtrasable(repere_ES_NH4)
    ES_primaire_NO3 = eaux_sales_discfilter(repere_ES_NO3) + eaux_sales_filtrasable(repere_ES_NO3)
    ES_primaire_Pt = eaux_sales_discfilter(repere_ES_Pt) + eaux_sales_filtrasable(repere_ES_Pt)
    For I = 1 To nb_eaux_sales_max
        If choix_eaux_sales(I) = devenir_ES_primaire Then
            'gestion des flux
            ES_primaire_Q = ES_primaire_Q + eaux_sales(I, repere_ES_Q)
            ES_primaire_DCO = ES_primaire_DCO + eaux_sales(I, repere_ES_DCO)
            ES_primaire_DBO = ES_primaire_DBO + eaux_sales(I, repere_ES_DBO)
            ES_primaire_MES = ES_primaire_MES + eaux_sales(I, repere_ES_MES)
            ES_primaire_NK = ES_primaire_NK + eaux_sales(I, repere_ES_NK)
            ES_primaire_NH4 = ES_primaire_NH4 + eaux_sales(I, repere_ES_NH4)
            ES_primaire_NO3 = ES_primaire_NO3 + eaux_sales(I, repere_ES_NO3)
            ES_primaire_Pt = ES_primaire_Pt + eaux_sales(I, repere_ES_Pt)
        End If
    Next I
    'écriture des concentration dans "Valeurs_générales" ATTENTION C'EST FAUX, IL NE FAUT PAS MELANGER EAUX SALES ET RETOURS !!!!!!!!!!!!!!!!!!!
    'COMME DISCUTE LE 15, IL FAUT CREER UN AUTRE BLOC POUR LES EAUX SALES QUI REVIENNENT EN TETE DE PRIMAIRE ET QUI SONT DIFFERENTES DES RETOURS!!!
    'on peut appeler la plage de cellule "eaux_sales_tete_primaire_flux"
    'NE PAS OUBLIER LA SUB "effacer_resultats_step"
    'on utilise les repere_ret_XXX et non repere_ES_XXX car on veut afficher les mêmes infos que les retours et que les ES contiennent d'autres paramètres
    Feuil6.Range("eaux_sales_tete_primaire_flux").Cells(repere_ret_Q, Colonne) = ES_primaire_Q
    Feuil6.Range("eaux_sales_tete_primaire_flux").Cells(repere_ret_DCO, Colonne) = ES_primaire_DCO       '/ ES_primaire_Q * 1000
    Feuil6.Range("eaux_sales_tete_primaire_flux").Cells(repere_ret_DBO, Colonne) = ES_primaire_DBO       '/ ES_primaire_Q * 1000
    Feuil6.Range("eaux_sales_tete_primaire_flux").Cells(repere_ret_MES, Colonne) = ES_primaire_MES       '/ ES_primaire_Q * 1000
    Feuil6.Range("eaux_sales_tete_primaire_flux").Cells(repere_ret_NK, Colonne) = ES_primaire_NK       '/ ES_primaire_Q * 1000
    Feuil6.Range("eaux_sales_tete_primaire_flux").Cells(repere_ret_NH4, Colonne) = ES_primaire_NH4       '/ ES_primaire_Q * 1000
    Feuil6.Range("eaux_sales_tete_primaire_flux").Cells(repere_ret_NO3, Colonne) = ES_primaire_NO3       '/ ES_primaire_Q * 1000
    Feuil6.Range("eaux_sales_tete_primaire_flux").Cells(repere_ret_Pt, Colonne) = ES_primaire_Pt       '/ ES_primaire_Q * 1000
    'fin modif MSM 15/11/12

    'écriture des concentrations de sortie
    If Q > 0 Then
        Feuil6.Range("eau_epuree_DCO").Cells(1, Colonne) = DCO / Q * 1000
        Feuil6.Range("eau_epuree_DBO").Cells(1, Colonne) = DBO / Q * 1000
        Feuil6.Range("eau_epuree_MES").Cells(1, Colonne) = MES / Q * 1000
        Feuil6.Range("eau_epuree_NGL").Cells(1, Colonne) = (NK + NO3) / Q * 1000
        Feuil6.Range("eau_epuree_NK").Cells(1, Colonne) = NK / Q * 1000
        Feuil6.Range("eau_epuree_NH4").Cells(1, Colonne) = NH4 / Q * 1000
        Feuil6.Range("eau_epuree_NO3").Cells(1, Colonne) = NO3 / Q * 1000
        Feuil6.Range("eau_epuree_Pt").Cells(1, Colonne) = Pt / Q * 1000
    Else
        Feuil6.Range("eau_epuree_DCO").Cells(1, Colonne) = "'-"
        Feuil6.Range("eau_epuree_DBO").Cells(1, Colonne) = "'-"
        Feuil6.Range("eau_epuree_MES").Cells(1, Colonne) = "'-"
        Feuil6.Range("eau_epuree_NGL").Cells(1, Colonne) = "'-"
        Feuil6.Range("eau_epuree_NK").Cells(1, Colonne) = "'-"
        Feuil6.Range("eau_epuree_NH4").Cells(1, Colonne) = "'-"
        Feuil6.Range("eau_epuree_NO3").Cells(1, Colonne) = "'-"
        Feuil6.Range("eau_epuree_Pt").Cells(1, Colonne) = "'-"
    End If

    'Refus de dégrillage
    Feuil6.Range("QuantiteRefusDegrillage_kgj").Cells(1, Colonne) = QuantiteRefusDegrillage_kgj
    
    'écriture des boues en entrée de la filière de traitement
    For I = repere_MES To repere_MV_MES
        Feuil6.Range("boues_I_caracteristiques").Cells(I, Colonne) = boues_flux(boues_inlet, boues_I, I)
    Next I
    For I = repere_MES To repere_MV_MES
        Feuil6.Range("boues_II_caracteristiques").Cells(I, Colonne) = boues_flux(boues_inlet, boues_II, I)
    Next I
    For I = repere_MES To repere_MV_MES
        Feuil6.Range("boues_III_caracteristiques").Cells(I, Colonne) = boues_flux(boues_inlet, boues_III, I)
    Next I
    For I = repere_MES To repere_MV_MES
        Feuil6.Range("boues_externes_1_caracteristiques").Cells(I, Colonne) = boues_flux(boues_inlet, boues_externes_1, I)
    Next I
    For I = repere_MES To repere_MV_MES
        Feuil6.Range("boues_externes_2_caracteristiques").Cells(I, Colonne) = boues_flux(boues_inlet, boues_externes_2, I)
    Next I
    'graisses internes
    Feuil6.Range("graisses_internes_caracteristiques").Cells(repere_MES, Colonne) = graisses_internes(repere_graisse_MS)
    Feuil6.Range("graisses_internes_caracteristiques").Cells(repere_Q, Colonne) = graisses_internes(repere_graisse_Q)
    Feuil6.Range("graisses_internes_caracteristiques").Cells(repere_MV_MES, Colonne) = graisses_internes(repere_graisse_MV_MES)
    Feuil6.Range("graisses_internes_caracteristiques").Cells(repere_graisse_DCO_MES + 1, Colonne) = graisses_internes(repere_graisse_MS) * graisses_internes(repere_graisse_DCO_MES)
    'graisses externes
    If choix_coferments = True Then
        intermediaire_Q_graisse_coferment = boues_flux(boues_inlet, graisses, repere_Q)
        intermediaire_Q_coferment = intermediaire_Q_graisse_coferment - graisses_internes(repere_graisse_Q)
        intermediaire_MS_coferments_kgj = boues_flux(boues_inlet, graisses, repere_MES) + boues_flux(boues_inlet, graisses, repere_Q) * boues_pollution_soluble(boues_inlet, graisses, repere_mgL_MS_soluble) / 1000 - graisses_internes(repere_graisse_MS)
        Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_Q_m3j, Colonne) = intermediaire_Q_coferment
        Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_MS_gL, Colonne) = intermediaire_MS_coferments_kgj / intermediaire_Q_coferment
        Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_ratio_MES_MS, Colonne) = (boues_flux(boues_inlet, graisses, repere_MES) - graisses_internes(repere_graisse_MS)) / intermediaire_MS_coferments_kgj
        intermediaire_MV_MS_coferments = (boues_flux(boues_inlet, graisses, repere_MES) * boues_flux(boues_inlet, graisses, repere_MV_MES) + boues_pollution_soluble(boues_inlet, graisses, repere_mgL_MV_soluble) / 1000 * intermediaire_Q_graisse_coferment - graisses_internes(repere_graisse_MS) * graisses_internes(repere_graisse_MV_MES)) / intermediaire_MS_coferments_kgj
        Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_ratio_MV_MS, Colonne) = intermediaire_MV_MS_coferments
        Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_ratio_DCO_MV, Colonne) = (boues_flux(boues_inlet, graisses, repere_MES) * boues_flux(boues_inlet, graisses, repere_ratio_DCO_MES) + boues_pollution_soluble(boues_inlet, graisses, repere_mgL_DCO) / 1000 * intermediaire_Q_graisse_coferment - graisses_internes(repere_graisse_MS) * graisses_internes(repere_graisse_DCO_MES)) / (intermediaire_MS_coferments_kgj * intermediaire_MV_MS_coferments)
        Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_ratio_NK_MV, Colonne) = (boues_flux(boues_inlet, graisses, repere_MES) * boues_flux(boues_inlet, graisses, repere_ratio_NK_MES) + boues_pollution_soluble(boues_inlet, graisses, repere_mgL_NK) / 1000 * intermediaire_Q_graisse_coferment - graisses_internes(repere_graisse_MS) * graisses_internes(repere_graisse_NK_MES)) / (intermediaire_MS_coferments_kgj * intermediaire_MV_MS_coferments)
        Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_ratio_Pt_MV, Colonne) = (boues_flux(boues_inlet, graisses, repere_MES) * boues_flux(boues_inlet, graisses, repere_ratio_Pt_MES) + boues_pollution_soluble(boues_inlet, graisses, repere_mgL_Pt) / 1000 * intermediaire_Q_graisse_coferment - graisses_internes(repere_graisse_MS) * graisses_internes(repere_graisse_Pt_MES)) / (intermediaire_MS_coferments_kgj * intermediaire_MV_MS_coferments)
        Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_BMP_Nm3CH4_tMVapplique, Colonne) = Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_BMP_Nm3CH4_tMVapplique, pi_FiliereConsideree)
        Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_BMP_securite, Colonne) = Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_BMP_securite, pi_FiliereConsideree)
    End If



End Sub

Sub effacer_resultats_step()

    Dim Colonne As Integer
    Dim I As Integer


    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


    'retours de tête
    For I = 1 To retour_caracteristique_nb
        Feuil6.Range("retours_flux").Cells(I, Colonne) = ""
        Feuil6.Range("retours_concentration").Cells(I, Colonne) = ""
    Next I

    'eaux sales
    'ce n'est pas une erreur, c'est parce que les eaux sales contiennent plus d'information que ce que l'on veut et nous on souhaite afficher les mêmes choses que pour les retours sauf les concentrations
    For I = 1 To retour_caracteristique_nb
        Feuil6.Range("eaux_sales_tete_primaire_flux").Cells(I, Colonne) = ""
    Next I

    'concentrations de sortie
    Feuil6.Range("eau_epuree_DCO").Cells(1, Colonne) = ""
    Feuil6.Range("eau_epuree_DBO").Cells(1, Colonne) = ""
    Feuil6.Range("eau_epuree_MES").Cells(1, Colonne) = ""
    Feuil6.Range("eau_epuree_NGL").Cells(1, Colonne) = ""
    Feuil6.Range("eau_epuree_NK").Cells(1, Colonne) = ""
    Feuil6.Range("eau_epuree_NH4").Cells(1, Colonne) = ""
    Feuil6.Range("eau_epuree_NO3").Cells(1, Colonne) = ""
    Feuil6.Range("eau_epuree_Pt").Cells(1, Colonne) = ""

    'écriture des boues en entrée de la filière de traitement
    For I = repere_MES To repere_MV_MES
        Feuil6.Range("boues_I_caracteristiques").Cells(I, Colonne) = ""
    Next I
    For I = repere_MES To repere_MV_MES
        Feuil6.Range("boues_II_caracteristiques").Cells(I, Colonne) = ""
    Next I
    For I = repere_MES To repere_MV_MES
        Feuil6.Range("boues_III_caracteristiques").Cells(I, Colonne) = ""
    Next I
    For I = repere_MES To repere_MV_MES
        Feuil6.Range("boues_externes_1_caracteristiques").Cells(I, Colonne) = ""
    Next I
    For I = repere_MES To repere_MV_MES
        Feuil6.Range("boues_externes_2_caracteristiques").Cells(I, Colonne) = ""
    Next I
    For I = repere_coferment_Q_m3j To repere_coferment_BMP_securite
        Feuil6.Range("graisses_externes_caracteristiques").Cells(I, Colonne) = ""
    Next I
    For I = repere_MES To repere_MV_MES + 1
        Feuil6.Range("graisses_internes_caracteristiques").Cells(I, Colonne) = ""
    Next I

    'feuille experts
    Feuil34.Range("expert_DCO_eaux_usees").Cells(pi_FiliereConsideree, 1) = ""
    Feuil34.Range("expert_DCO_eau_traitee").Cells(pi_FiliereConsideree, 1) = ""
    Feuil34.Range("expert_DCO_conso_bio").Cells(pi_FiliereConsideree, 1) = ""
    Feuil34.Range("expert_DCO_methanol").Cells(pi_FiliereConsideree, 1) = ""
    Feuil34.Range("expert_DCO_boues_extraites").Cells(pi_FiliereConsideree, 1) = ""
    Feuil34.Range("expert_DCO_boues_externes").Cells(pi_FiliereConsideree, 1) = ""
    Feuil34.Range("expert_DCO_traitement_retours").Cells(pi_FiliereConsideree, 1) = ""
    Feuil34.Range("expert_DCO_methanol_Shunt").Cells(pi_FiliereConsideree, 1) = ""
    Feuil34.Range("expert_DCO_retours").Cells(pi_FiliereConsideree, 1) = ""
    Feuil34.Range("expert_DCO_biogaz").Cells(pi_FiliereConsideree, 1) = ""

    If version_Developpement = False Then On Error GoTo exception_2
    'Feuil34.Range("DCO_oxydation").Cells(pi_FiliereConsideree, 1) = ""
exception_2:
    If version_Developpement = False Then On Error GoTo 0    'invalide le gestionnaire d'erreur
    Feuil34.Range("expert_DCO_boues_evacuees").Cells(pi_FiliereConsideree, 1) = ""
    Feuil34.Range("expert_boues_evacuees_DCO_MV").Cells(pi_FiliereConsideree, 1) = ""
    Feuil34.Range("expert_DCO_OVH_incineration").Cells(pi_FiliereConsideree, 1) = ""
    Feuil34.Range("expert_DCO_boues_evacuees").Cells(pi_FiliereConsideree, 1) = ""


End Sub

Sub expert(ByVal DCO, ByVal DCO_conso_bio, ByVal boues_flux, ByVal retour_flux, ByVal DCO_biogaz, ByVal DCO_methanol, ByVal graisses_internes, ByVal boues_pollution_soluble, ByVal DCO_traitement_retours, ByVal DCO_methanol_Shunt, ByVal boues_Shunt, ByVal DCO_oxydation)

    Dim I As Integer
    Dim j As Integer
    Dim DCO_somme As Double
    Dim rapport As Double
    Dim intermediaire_MV As Double


    'BILAN SUR LA DCO

    'entrée eau
    Feuil34.Range("expert_DCO_eaux_usees").Cells(pi_FiliereConsideree, 1) = DCO_nominal * pd_PourcentageChargeReelleDCO

    'sortie eau
    Feuil34.Range("expert_DCO_eau_traitee").Cells(pi_FiliereConsideree, 1) = DCO

    'DCO consommée par les bactéries dans le bio
    Feuil34.Range("expert_DCO_conso_bio").Cells(pi_FiliereConsideree, 1) = DCO_conso_bio

    'DCO apportée par méthanol
    Feuil34.Range("expert_DCO_methanol").Cells(pi_FiliereConsideree, 1) = DCO_methanol

    'boues_extraites + graisses internes
    DCO_somme = 0
    For I = 1 To 3
        DCO_somme = DCO_somme + boues_flux(boues_inlet, I, repere_ratio_DCO_MES) * boues_flux(boues_inlet, I, repere_MES)
    Next I
    DCO_somme = DCO_somme + graisses_internes(repere_graisse_DCO_MES) * graisses_internes(repere_graisse_MS) - boues_Shunt(repere_graisse_DCO_MES) * boues_Shunt(repere_graisse_MS)
    Feuil34.Range("expert_DCO_boues_extraites").Cells(pi_FiliereConsideree, 1) = DCO_somme

    'boues_externes + graisses externes
    DCO_somme = 0
    For I = boues_externes_1 To nb_type_boues
        DCO_somme = DCO_somme + boues_flux(boues_inlet, I, repere_ratio_DCO_MES) * boues_flux(boues_inlet, I, repere_MES)
    Next I
    'on retranche les graisses internes mélangées aux externes
    DCO_somme = DCO_somme + boues_pollution_soluble(boues_inlet, graisses, repere_mgL_DCO) / 1000 * boues_flux(boues_inlet, graisses, repere_Q) - graisses_internes(repere_graisse_DCO_MES) * graisses_internes(repere_graisse_MS)
    Feuil34.Range("expert_DCO_boues_externes").Cells(pi_FiliereConsideree, 1) = DCO_somme

    'retours en tête
    Feuil34.Range("expert_DCO_retours").Cells(pi_FiliereConsideree, 1) = retour_flux(repere_ret_DCO)

    'traitement séparé des retours
    Feuil34.Range("expert_DCO_traitement_retours").Cells(pi_FiliereConsideree, 1) = DCO_traitement_retours
    Feuil34.Range("expert_DCO_methanol_Shunt").Cells(pi_FiliereConsideree, 1) = DCO_methanol_Shunt

    'DCO méthanisée dans le digesteur
    Feuil34.Range("expert_DCO_biogaz").Cells(pi_FiliereConsideree, 1) = DCO_biogaz

    'CO2 émis par les combustion (Athos et Pyrofluid)
    Feuil34.Range("expert_DCO_OVH_incineration").Cells(pi_FiliereConsideree, 1) = DCO_oxydation

    'DCO des boues évacuées
    DCO_somme = 0
    intermediaire_MV = 0
    For I = 1 To nb_etape_file_boues - 1
        For j = 1 To nb_type_boues
            If boues_flux(I, j, repere_verif_flux) <> 0 Then
                rapport = boues_flux(I, j, repere_verif_flux) / boues_flux(I, j, repere_flux_in)
                DCO_somme = DCO_somme + rapport * boues_flux(I, j, repere_ratio_DCO_MES) * boues_flux(I, j, repere_MES)
                DCO_somme = DCO_somme + rapport * boues_flux(I, j, repere_Q) * boues_pollution_soluble(I, j, repere_mgL_DCO) / 1000
                intermediaire_MV = intermediaire_MV + rapport * boues_flux(I, j, repere_MV_MES) * boues_flux(I, j, repere_MES)
            End If
        Next j
    Next I
    Feuil34.Range("expert_DCO_boues_evacuees").Cells(pi_FiliereConsideree, 1) = DCO_somme
    If intermediaire_MV <> 0 Then
        Feuil34.Range("expert_boues_evacuees_DCO_MV").Cells(pi_FiliereConsideree, 1) = DCO_somme / intermediaire_MV
    End If


End Sub

Public Sub gestion_erreur_non_geree(ByVal NumeroErreurNonGeree)

    'gestion d'une erreur non gérée

    Dim Titre As String
    Dim mensaje As String


    'MSM 05/12/12
    Sheets("Results0").Visible = xlSheetVeryHidden
    Dim feuille As Worksheet
    For Each feuille In ThisWorkbook.Worksheets
        If feuille.CodeName = "Feuil_Results11" Or feuille.CodeName = "Feuil_Results12" Or feuille.CodeName = "Feuil_Results13" Then
            feuille.Visible = xlSheetVeryHidden
        End If
    Next feuille


    If pn_ErreurNonGeree = False Then
        mensaje = Feuil6.Range("erreur_non_geree_1") & " " & NumeroErreurNonGeree
        mensaje = mensaje + vbCrLf + Feuil6.Range("erreur_non_geree_2")
        Titre = Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        MsgBox mensaje, Buttons:=vbCritical, Title:=Titre
    End If

    pn_ErreurNonGeree = True


End Sub

Sub hypotheses_boues(ByRef boues_ratio_NK_MV, ByRef boues_ratio_Pt_MES, ByRef boues_ratio_DCO_MV, ByRef boues_ratio_DBO_MV)

    'RATIO SUR LES POLLUTIONS DES BOUES (cf TABLEUR DE DIMENSIONNEMENT DE DIGESTION)

    'ces ratios sont utilisés pour les traitements biologiques au niveau de l'azote (une partie sort sous forme gazeuse et on ne sait pas la quantifier)
    'pour le reste, tout sera fait par bilan matière et on ne les utilisera que pour les boues entrées manuellement

    'primaire simple
    boues_ratio_NK_MV(I_simple) = 0.059
    boues_ratio_Pt_MES(I_simple) = 0.009
    boues_ratio_DCO_MV(I_simple) = 1.71
    boues_ratio_DBO_MV(I_simple) = 0.7

    'primaire reactif
    boues_ratio_NK_MV(I_reactif) = 0.059
    boues_ratio_Pt_MES(I_reactif) = 0.009
    boues_ratio_DCO_MV(I_reactif) = 1.71
    boues_ratio_DBO_MV(I_reactif) = 0.7

    'forte charge
    boues_ratio_NK_MV(II_forte) = 0.065
    boues_ratio_Pt_MES(II_forte) = 0.014
    boues_ratio_DCO_MV(II_forte) = 1.71
    boues_ratio_DBO_MV(II_forte) = 0.9

    'moyenne charge
    boues_ratio_NK_MV(II_moyenne) = 0.08
    boues_ratio_Pt_MES(II_moyenne) = 0.02
    boues_ratio_DCO_MV(II_moyenne) = 1.6
    boues_ratio_DBO_MV(II_moyenne) = 0.75

    'faible charge eau brute
    boues_ratio_NK_MV(II_faible_EB) = 0.095
    boues_ratio_Pt_MES(II_faible_EB) = 0.02
    boues_ratio_DCO_MV(II_faible_EB) = 1.5
    boues_ratio_DBO_MV(II_faible_EB) = 0.5

    'faible charge eau décantée
    boues_ratio_NK_MV(II_faible_ED) = 0.095
    boues_ratio_Pt_MES(II_faible_ED) = 0.02
    boues_ratio_DCO_MV(II_faible_ED) = 1.45
    boues_ratio_DBO_MV(II_faible_ED) = 0.5

    'aération prolongée eau brute
    boues_ratio_NK_MV(II_prolongee_EB) = 0.095
    boues_ratio_Pt_MES(II_prolongee_EB) = 0.02
    boues_ratio_DCO_MV(II_prolongee_EB) = 1.45
    boues_ratio_DBO_MV(II_prolongee_EB) = 0.5

    'aération prolongée eau brute
    boues_ratio_NK_MV(II_prolongee_ED) = 0.095
    boues_ratio_Pt_MES(II_prolongee_ED) = 0.02
    boues_ratio_DCO_MV(II_prolongee_ED) = 1.45
    boues_ratio_DBO_MV(II_prolongee_ED) = 0.5

    'MBR    A REVOIR
    boues_ratio_NK_MV(II_MBR) = 0.095
    boues_ratio_Pt_MES(II_MBR) = 0.02
    boues_ratio_DCO_MV(II_MBR) = 1.5
    boues_ratio_DBO_MV(II_MBR) = 0.5

    'MBBR   A REVOIR
    boues_ratio_NK_MV(II_MBBR) = 0.095
    boues_ratio_Pt_MES(II_MBBR) = 0.02
    boues_ratio_DCO_MV(II_MBBR) = 1.5
    boues_ratio_DBO_MV(II_MBBR) = 0.5

    'HybAS    A REVOIR
    boues_ratio_NK_MV(II_HybAS) = 0.095
    boues_ratio_Pt_MES(II_HybAS) = 0.02
    boues_ratio_DCO_MV(II_HybAS) = 1.5
    boues_ratio_DBO_MV(II_HybAS) = 0.5

    'biostyr_C
    boues_ratio_NK_MV(II_biostyr_C) = 0.01
    boues_ratio_Pt_MES(II_biostyr_C) = 0.025
    boues_ratio_DCO_MV(II_biostyr_C) = 1.6
    boues_ratio_DBO_MV(II_biostyr_C) = 0.9

    'biostyr_N
    boues_ratio_NK_MV(II_biostyr_N) = 0.01
    boues_ratio_Pt_MES(II_biostyr_N) = 0.025
    boues_ratio_DCO_MV(II_biostyr_N) = 1.55
    boues_ratio_DBO_MV(II_biostyr_N) = 0.9

    'biostyr_NDN
    boues_ratio_NK_MV(II_biostyr_NDN) = 0.01
    boues_ratio_Pt_MES(II_biostyr_NDN) = 0.025
    boues_ratio_DCO_MV(II_biostyr_NDN) = 1.5
    boues_ratio_DBO_MV(II_biostyr_NDN) = 0.9

    'biostyr_PDN
    boues_ratio_NK_MV(II_biostyr_PDN) = 0.01
    boues_ratio_Pt_MES(II_biostyr_PDN) = 0.015
    boues_ratio_DCO_MV(II_biostyr_PDN) = 1.45
    boues_ratio_DBO_MV(II_biostyr_PDN) = 0.9

    'décantation III
    boues_ratio_NK_MV(III_decantation) = 0.041
    boues_ratio_Pt_MES(III_decantation) = 0.04
    boues_ratio_DCO_MV(III_decantation) = 1.45
    boues_ratio_DBO_MV(III_decantation) = 0.45

    'biostyr_N III
    boues_ratio_NK_MV(III_biostyr_N) = 0.041
    boues_ratio_Pt_MES(III_biostyr_N) = 0.04
    boues_ratio_DCO_MV(III_biostyr_N) = 1.45
    boues_ratio_DBO_MV(III_biostyr_N) = 0.45

    'biostyr_PDN III
    boues_ratio_NK_MV(III_biostyr_PDN) = 0.041
    boues_ratio_Pt_MES(III_biostyr_PDN) = 0.04
    boues_ratio_DCO_MV(III_biostyr_PDN) = 1.45
    boues_ratio_DBO_MV(III_biostyr_PDN) = 0.45

    'graisses
    boues_ratio_NK_MV(codigestion_graisses) = 0.002
    boues_ratio_Pt_MES(codigestion_graisses) = 0.002
    boues_ratio_DCO_MV(codigestion_graisses) = 2.8
    boues_ratio_DBO_MV(codigestion_graisses) = 2.5


End Sub

Sub prgm_principal()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'POUR LE PINCH
    PINCH_ORDONNER_FLUX_PROCESS = True
    PINCH_ORDONNER_FLUX_UTILITES = True
    PINCH_DANS_OCEAN = True

    'variables de choix du nombre de traitements(épaississement, relèvements)
    Dim relevement_choix_nb As Integer

    'variables internes au programme principal (nombres entiers)
    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    'Dim message_erreur As String
    'Dim verification_configuration As Integer
    'Dim treatment_message As String
    Dim warning As Boolean

    Dim interm_etape As Double

    'itération sur les eaux sales dans le cas d'un traitement séparé avec retour en amont du biostyr
    Dim nb_iteration_ES As Integer
    Dim iteration_ES As Integer

    'variables des flux de matières à transiter (file eau)
    Dim Q As Double
    Dim DCO As Double
    Dim DBO As Double
    Dim MES As Double
    Dim NK As Double
    Dim NH4 As Double
    Dim NO3 As Double
    Dim Pt As Double
    Dim Sh As Double

    Dim Q_ES_traitement_separe As Double

    'variables pour traitement des sulfures
    Dim HS_strippe_kgj As Double
    Dim HS_traite_kgj As Double
    Dim HS_O2_liquide_kgj As Double
    Dim HS_FeCl3pur_kgj As Double
    Dim HS_Ca_2NO3_pur_flux As Double
    Dim HS_H2O2pur_flux As Double

    'variables permettant le stockage des flux lorque l'on a des itérations intermédiaires (eaux sales par exemple)
    Dim stockage_Q As Double
    Dim stockage_DCO As Double
    Dim stockage_DBO As Double
    Dim stockage_MES As Double
    Dim stockage_NK As Double
    Dim stockage_NH4 As Double
    Dim stockage_NO3 As Double
    Dim stockage_Pt As Double
    Dim stockage_DCO_conso_bio As Double

    'variables de caractérisation des retours
    Dim retour_flux(1 To retour_caracteristique_nb) As Double
    Dim retour_digestion(1 To retour_caracteristique_nb) As Double
    Dim retour_athos(retour_caracteristique_nb) As Double
    Dim retour_flux_soluble(1 To retour_caracteristique_nb) As Double
    Dim retour_digestion_soluble(1 To retour_caracteristique_nb) As Double
    Dim retour_athos_soluble(1 To retour_caracteristique_nb) As Double

    Dim Q_retour As Double

    'variables de caractérisation des eaux sales
    Dim choix_eaux_sales_1aire As Boolean
    Dim choix_eaux_sales_2aire_amont As Boolean
    Dim choix_eaux_sales_2aire_aval As Boolean
    Dim choix_eaux_sales_3aire_amont As Boolean
    Dim choix_eaux_sales_3aire_aval As Boolean
    Dim Q_eaux_sales As Double
    Dim MES_eaux_sales As Double
    Dim eaux_sales(1 To nb_eaux_sales_max, 1 To nb_parametres_eaux_sales) As Double
    Dim eaux_sales_discfilter(1 To nb_parametres_eaux_sales) As Double
    Dim eaux_sales_filtrasable(1 To nb_parametres_eaux_sales) As Double
    Dim choix_eaux_sales(nb_eaux_sales_max) As Integer
    Dim boues_I_detail(1 To nb_eaux_sales_max, 1 To nb_parametres_boues_I_detail) As Double
    Dim boues_II_detail(1 To nb_eaux_sales_max, 1 To nb_parametres_boues_I_detail) As Double
    Dim boues_III_detail(1 To nb_eaux_sales_max, 1 To nb_parametres_boues_I_detail) As Double
    Dim boues_evacuees_Q As Double
    Dim graisses_evacuees_Q As Double

    'Bilan sur la DCO
    Dim DCO_conso_bio As Double
    Dim DCO_biogaz As Double
    Dim DCO_oxydation As Double
    Dim DCO_apportee_methanol As Double
    Dim DCO_traitement_retours As Double
    Dim DCO_methanol_Shunt As Double

    'variables sur les boues à véhiculer   obj: conserver l'information pour la digestion (MES,Q,MV/MES et part des boues_extraites)
    Const nb_max_epaississeur = 5
    Const nb_max_deshydratation = 3
    Const nb_max_sechage = 2
    Const nb_max_sechage_bioco = 2
    Const nb_max_sechage_inos = 2
    Const nb_max_chaulage = 3

    Dim TableauRecapitulatifFluxBoues(1 To nb_etape_file_boues, 1 To nb_type_boues, 1 To nb_parametres_boues) As Double

    For I = 1 To nb_max_epaississeur
        boues_epaissies(I) = boues_inlet + I
    Next I
    For I = 1 To nb_max_deshydratation
        boues_deshydratees(I) = boues_athos + I
    Next I
    For I = 1 To nb_max_chaulage
        boues_chaulees(I) = boues_deshydratees(nb_max_deshydratation) + I
    Next I
    For I = 1 To nb_max_sechage_inos
        boues_sechees_inos(I) = boues_chaulees(nb_max_chaulage) + I
    Next I
    For I = 1 To nb_max_sechage_bioco
        boues_sechees_bioco(I) = boues_sechees_inos(nb_max_sechage_inos) + I
    Next I
    For I = 1 To nb_max_sechage
        boues_sechees(I) = boues_sechees_bioco(nb_max_sechage_bioco) + I
    Next I
    Dim epaississement_choix_nb As Integer
    Dim epaississement_before_thelys(nb_max_epaississeur) As Integer
    For I = 1 To nb_max_epaississeur
        epaississement_before_thelys(I) = 0
    Next I

    Dim deshydratation_choix_nb As Integer
    Dim chaulage_choix_nb As Integer
    Dim sechage_choix_nb As Integer
    Dim sechage_bioco_choix_nb As Integer
    Dim sechage_inos_choix_nb As Integer
    Dim boues_ratio_NK_MV(nb_qualites_boues) As Double
    Dim boues_ratio_Pt_MES(nb_qualites_boues) As Double
    Dim boues_ratio_DCO_MV(nb_qualites_boues) As Double
    Dim boues_ratio_DBO_MV(nb_qualites_boues) As Double
    Dim boues_pollution_soluble(1 To nb_etape_file_boues, 1 To nb_type_boues, 1 To nb_repere_mgL) As Double
    Dim graisses_internes(1 To repere_graisse_Pt_MES) As Double
    Dim graisses_particulaire_ratio_interne(1 To nb_parametres_graisses_particulaire_detail) As Double
    Dim QuantiteRefusDegrillage_kgj As Double
    
    'boues du Shunt
    Dim boues_Shunt(repere_graisse_Pt_MES) As Double      'même paramètres que graisses

    'variables sur électricité et réactifs
    Dim electricite_fixe(electricite_postes_nb) As Double    'représente les consommations fixes des procédés (électricité)
    Dim electricite_variable(electricite_postes_nb) As Double    'représente les consommations variables des procédés (électricité)

    Dim electricite_consommee As Double
    Dim electricite_II_aeration As Double
    Dim electricite_III_aeration As Double
    Dim electricite_utilites As Double

    'modif MSM 25/10/12
    Dim electricite_verte As Double
    Dim electricite_verte_biogaz As Double
    Dim electricite_produite_solaire_photovoltaique As Double
    Dim electricite_produite_turbine_hydraulique As Double
    Dim electricite_produite_autre As Double

    Dim cendres_evacuees_Tj As Double
    Dim REFIB_evacues_Tj As Double
    Dim cendres_evacuees_tMSj As Double
    Dim REFIB_evacues_tMSj As Double

    Dim polymere_flux As Double
    Dim polymere_eau_pur_kgj(1 To polymere_nb_type) As Double
    Dim polymere_boues_pur_kgj(1 To polymere_nb_type) As Double
    Dim chaux_eteinte_flux As Double
    Dim chaux_vive_flux As Double
    Dim FeCl3_flux As Double
    Dim FeCl3_eau_pur_kgj As Double
    Dim FeCl3_boues_pur_kgj As Double
    Dim methanol_flux As Double
    Dim H2SO4pur_flux As Double
    Dim NaOHpur_flux As Double
    Dim NaHCO3_flux As Double
    Dim NaOClpur_flux As Double
    Dim NaOClpur_eau_flux As Double
    Dim NaOClpur_desodo_flux As Double
    Dim NaHSO3pur_flux As Double
    Dim O2_liquide_flux As Double
    Dim charbon_actif_flux As Double
    Dim Cl2_gazeux_flux_kgj As Double
    Dim Ca_2NO3_pur_flux As Double
    Dim H2O2pur_flux As Double
    Dim CuSO4_flux As Double
    Dim MgCl2_flux As Double
    Dim ConsommationAmmoniaquePur_kgj As Double
    Dim ConsommationUreePur_kgj As Double
    
    Dim Struvite_kgj As Double

    Dim C_elimine As Double
    Dim DBO_elimine As Double
    Dim N_elimine As Double
    Dim P_elimine As Double
    Dim P_precipite As Double
    Dim gaz_naturel_kWhPCIj As Double
    Dim fioul_kWhPCIj As Double

    'variables sur besoins thermiques
    Dim besoins_thermiques(nb_niveaux_energie) As Double
    Dim disponibilites_thermiques(nb_niveaux_energie) As Double
    Dim pression_vapeur(nb_vapeur_max) As Double
    Dim nb_vapeur As Integer
    Dim biogaz(nb_biogaz_caracteristiques) As Double
    Dim type_valorisation_biogaz As Integer
    Dim valorisation_chaudiere As Integer
    Dim ratio_biogaz_valorise As Double

    'Déclaration module de calcul des KPI car on peut avoir besoin de récupérer des valeurs au fur et à mesure
    Dim GestionCalculKPI As New CLS_GeneralCalculKPI

    ps_CommentairesCalculsGlobaux = ""
    pn_ErreurNonGeree = False
    pn_EffacementResultatsEnCours = False

    'On regarde si le PINCH est activé --> on le fait au début car le calcul de certains modules en dépend (Séchage thermique autre).
    choix_analyse_pinch = Feuil57.Range("choix_analyse_pinch").Cells(1, 1)
    
    'données pour l'analyse du pincement thermique
    ReDim flux_thermique_process(1 To flux_thermique_process_nb_caracteristiques, 1 To 1)
    ReDim utilites_flux_energie(1 To flux_thermique_utilites_nb_caracteristiques, 1 To 1)
    nb_flux_utilites = 0

    'lecture des choix de procédés
    Call lecture_choix_process
    If pn_EffacementResultatsEnCours = True Then
        GoTo calcul_interrompu
    End If

    'lecture des données générales
    Call lecture_donnees_generales

    'hypothèses sur des ratios NK/MES et Pt/MES des différents types de boues
    Call hypotheses_boues(boues_ratio_NK_MV, boues_ratio_Pt_MES, boues_ratio_DCO_MV, boues_ratio_DBO_MV)

    'initialisation variables électricité et réactifs
    electricite_consommee = 0
    electricite_verte = 0
    electricite_II_aeration = 0
    electricite_III_aeration = 0
    electricite_utilites = 0

    cendres_evacuees_Tj = 0
    REFIB_evacues_Tj = 0
    cendres_evacuees_tMSj = 0
    REFIB_evacues_tMSj = 0
    Struvite_kgj = 0
    
    For I = 1 To electricite_postes_nb
        electricite_fixe(I) = 0
        electricite_variable(I) = 0
    Next I

    nb_iteration_ES = 1  'on initialise le nombre d'itération pour les eaux sales à 1



    '''''''''''''''''
    'DIMENSIONNEMENT'
    '''''''''''''''''
    'on définit le module de classe des réactifs
    Dim bilan_reactifs As New Gestion_reactifs
    Call bilan_reactifs.initialisation_procede
    
    'traitement des sulfures
    'If choix_traitement_sulfures_curatif = True Then
    Dim traitement_HS As New Traitement_sulfures
    traitement_HS.pd_PureteNitrateDeCalcium = bilan_reactifs.pd_PureteNitrateDeCalcium
    Call traitement_HS.dimensionnement(SH_nominal)
    HS_O2_liquide_kgj = traitement_HS.O2_liquide_kgj
    HS_FeCl3pur_kgj = traitement_HS.FeCl3pur_kgj
    HS_Ca_2NO3_pur_flux = traitement_HS.Ca_2NO3_kgj
    HS_H2O2pur_flux = traitement_HS.H2O2_kgj
    If pn_EffacementResultatsEnCours = True Then
        GoTo calcul_interrompu
    End If
    'End If

    'PRETRAITEMENTS (hors boucles de calcul car indépendant)
    If choix_pretraitement = True Then

        'Dégrilllage
        If choix_degrillage = True Then
            Dim degrilleur As New CLS_EauDegrillage
            Call degrilleur.initialisation_procede
            Call degrilleur.dimensionnement
            If pn_EffacementResultatsEnCours = True Then
                GoTo calcul_interrompu
            End If
        End If

        'Dessablage
        If choix_dessablage = True Then
            Dim dessableur_deshuileur As New dessablage_deshuilage
            Call dessableur_deshuileur.calcul(graisses_internes, boues_ratio_DCO_MV, boues_ratio_NK_MV, boues_ratio_Pt_MES, pd_CapaciteSTEP_EH, vidange_MES_nominal)
            If pn_EffacementResultatsEnCours = True Then
                GoTo calcul_interrompu
            End If
        End If
    End If

    For iteration = 1 To pi_NOMBRE_ITERATION_RETOURS

        'REINITIALISATION DES VARIABLES D'ENTREE
        Q = Q_nominal + vidange_Q_nominal + retour_flux(repere_ret_Q) - graisses_internes(repere_graisse_Q)
        DCO = DCO_nominal + vidange_DCO_nominal + retour_flux(repere_ret_DCO) - graisses_internes(repere_graisse_DCO_MES) * graisses_internes(repere_graisse_MS)
        DBO = DBO_nominal + vidange_DBO_nominal + retour_flux(repere_ret_DBO) - graisses_internes(repere_graisse_DBO_MES) * graisses_internes(repere_graisse_MS)
        MES = MES_nominal + vidange_MES_nominal + retour_flux(repere_ret_MES) - graisses_internes(repere_graisse_MS)
        NK = NK_nominal + vidange_NK_nominal + retour_flux(repere_ret_NK) - graisses_internes(repere_graisse_NK_MES) * graisses_internes(repere_graisse_MS)
        NH4 = NH4_nominal + vidange_NH4_nominal + retour_flux(repere_ret_NH4)
        NO3 = retour_flux(repere_ret_NO3)
        Pt = Pt_nominal + vidange_Pt_nominal + retour_flux(repere_ret_Pt) - graisses_internes(repere_graisse_Pt_MES) * graisses_internes(repere_graisse_MS)
        Sh = SH_nominal - traitement_HS.SH_strippe_kgj - traitement_HS.SH_traite_kgj

        If choix_discfilter = True Then   'on réinjecte les eaux sales du discfilter
            Q = Q + eaux_sales_discfilter(repere_ES_Q)
            DCO = DCO + eaux_sales_discfilter(repere_ES_DCO)
            DBO = DBO + eaux_sales_discfilter(repere_ES_DBO)
            MES = MES + eaux_sales_discfilter(repere_ES_MES)
            NK = NK + eaux_sales_discfilter(repere_ES_NK)
            NH4 = NH4 + eaux_sales_discfilter(repere_ES_NH4)
            NO3 = NO3 + eaux_sales_discfilter(repere_ES_NO3)
            Pt = Pt + eaux_sales_discfilter(repere_ES_Pt)
        End If

        If choix_filtrasable = True Then   'on réinjecte les eaux sales du filtre à sable
            Q = Q + eaux_sales_filtrasable(repere_ES_Q)
            DCO = DCO + eaux_sales_filtrasable(repere_ES_DCO)
            DBO = DBO + eaux_sales_filtrasable(repere_ES_DBO)
            MES = MES + eaux_sales_filtrasable(repere_ES_MES)
            NK = NK + eaux_sales_filtrasable(repere_ES_NK)
            NH4 = NH4 + eaux_sales_filtrasable(repere_ES_NH4)
            NO3 = NO3 + eaux_sales_filtrasable(repere_ES_NO3)
            Pt = Pt + eaux_sales_filtrasable(repere_ES_Pt)
        End If

        Q_retour = retour_flux(repere_ret_Q)

        'REINITIALISATION DES FLUX DE BOUES
        For I = 1 To nb_etape_file_boues
            For j = 1 To nb_type_boues
                For k = 1 To nb_parametres_boues
                    TableauRecapitulatifFluxBoues(I, j, k) = 0
                Next k
            Next j
        Next I

        DCO_conso_bio = 0
        DCO_apportee_methanol = 0
        DCO_oxydation = 0
        polymere_flux = 0
        ReDim flux_thermique_process(1 To flux_thermique_process_nb_caracteristiques, 1 To 1)
        For I = 1 To polymere_nb_type
            polymere_eau_pur_kgj(I) = 0
            polymere_boues_pur_kgj(I) = 0
        Next I

        chaux_eteinte_flux = 0
        chaux_vive_flux = 0
        FeCl3_flux = HS_FeCl3pur_kgj
        FeCl3_eau_pur_kgj = HS_FeCl3pur_kgj
        FeCl3_boues_pur_kgj = 0
        methanol_flux = 0
        H2SO4pur_flux = 0
        NaOHpur_flux = 0
        NaHCO3_flux = 0
        NaOClpur_flux = 0
        NaOClpur_eau_flux = 0
        NaOClpur_desodo_flux = 0
        NaHSO3pur_flux = 0
        electricite_verte = 0
        O2_liquide_flux = HS_O2_liquide_kgj
        Ca_2NO3_pur_flux = HS_Ca_2NO3_pur_flux
        H2O2pur_flux = HS_H2O2pur_flux
        charbon_actif_flux = 0
        Cl2_gazeux_flux_kgj = 0
        nb_vapeur = 0
        CuSO4_flux = 0
        MgCl2_flux = 0
        ConsommationAmmoniaquePur_kgj = 0
        ConsommationUreePur_kgj = 0

        'REINITIALISATION DES ENERGIES
        For I = 1 To nb_niveaux_energie
            besoins_thermiques(I) = 0
            disponibilites_thermiques(I) = 0
        Next I


        ''''''''''
        'FILE EAU'
        ''''''''''

        'TRAITEMENT PRIMAIRE
        'eaux sales retournées en tête de primaire sans traitement séparé
        For I = 1 To nb_eaux_sales_max
            If choix_eaux_sales(I) = devenir_ES_primaire Then
                'flag pour la digestion
                choix_eaux_sales_1aire = True
                'gestion des flux
                Q = Q + eaux_sales(I, repere_ES_Q)
                DCO = DCO + eaux_sales(I, repere_ES_DCO)
                DBO = DBO + eaux_sales(I, repere_ES_DBO)
                MES = MES + eaux_sales(I, repere_ES_MES)
                NK = NK + eaux_sales(I, repere_ES_NK)
                NH4 = NH4 + eaux_sales(I, repere_ES_NH4)
                NO3 = NO3 + eaux_sales(I, repere_ES_NO3)
                Pt = Pt + eaux_sales(I, repere_ES_Pt)
            End If
        Next I

        'PRIMAIRE
        If choix_primaire = True Then
            C_elimine = DCO
            DBO_elimine = DBO
            N_elimine = NK
            P_elimine = Pt
            P_precipite = 0
            If choix_decanteur_simple = True Then
                If iteration = 1 Then
                    Dim decanteur_simple As New Decantation_simple
                End If
                Call decanteur_simple.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Q_retour, eaux_sales, choix_eaux_sales, boues_I_detail, eaux_sales_discfilter)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
                P_precipite = 0
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_I, repere_origine) = I_simple
                Next I
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MES) = decanteur_simple.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_Q) = decanteur_simple.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MV_MES) = decanteur_simple.MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_flux_in) = 1
            End If

            'Décanteur Reactif
            If choix_decanteur_reactif = True Then
                If iteration = 1 Then
                    Dim decanteur_reactif As New Decantation_reactif
                End If
                Call decanteur_reactif.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Q_retour, eaux_sales, choix_eaux_sales, boues_I_detail, eaux_sales_discfilter)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
                P_precipite = decanteur_reactif.P_precipite
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_I, repere_origine) = I_reactif
                Next I
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MES) = decanteur_reactif.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_Q) = decanteur_reactif.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MV_MES) = decanteur_reactif.MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_flux_in) = 1
            End If
            'calcul des ratios de pollution dans les boues qui sont constantes jusqu'à la digestion
            C_elimine = C_elimine - DCO
            DBO_elimine = DBO_elimine - DBO
            N_elimine = N_elimine - NK
            P_elimine = P_elimine - Pt - P_precipite
            For I = 1 To boues_epaissies(nb_max_epaississeur)
                TableauRecapitulatifFluxBoues(I, boues_I, repere_MV_MES) = TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MV_MES)
                TableauRecapitulatifFluxBoues(I, boues_I, repere_ratio_DCO_MES) = C_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MES)
                TableauRecapitulatifFluxBoues(I, boues_I, repere_ratio_DBO_MES) = DBO_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MES)
                TableauRecapitulatifFluxBoues(I, boues_I, repere_ratio_NK_MES) = N_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MES)
                TableauRecapitulatifFluxBoues(I, boues_I, repere_ratio_Pt_MES) = P_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MES)
            Next I
        End If

        'TRAITEMENT SECONDAIRE
        If choix_secondaire = True Then
            'On ajoute les boues de biolix
            If choix_biolix = True Then
                Q = Q + TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_Q)
                DCO = DCO + TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_ratio_DCO_MES) * TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_MES)
                DBO = DBO + TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_ratio_DBO_MES) * TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_MES)
                MES = MES + TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_MES)
                NK = NK + TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_ratio_NK_MES) * TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_MES)
                Pt = Pt + TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_ratio_Pt_MES) * TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_MES)
            End If
            'On stocke les valeurs de pollution en entrée
            C_elimine = DCO
            'N_elimine = NK
            P_elimine = Pt
            P_precipite = 0
            If choix_BA_forte = True Then
                If iteration = 1 Then
                    Dim BA_forte As New BA_forte_charge
                End If
                Call BA_forte.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_conso_bio)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
                P_precipite = 0
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_forte
                Next I
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = BA_forte.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = BA_forte.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = BA_forte.nominal_MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
            End If
            If choix_BA_moyenne = True Then
                If iteration = 1 Then
                    Dim BA_moyenne As New BA_moyenne_charge
                End If
                Call BA_moyenne.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_conso_bio)
                P_precipite = 0
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_moyenne
                Next I
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = BA_moyenne.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = BA_moyenne.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = BA_moyenne.nominal_MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
            End If
            If choix_BA_faible = True Then
                If iteration = 1 Then
                    Dim BA_faible As New BA_faible_charge
                End If
                Call BA_faible.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_apportee_methanol, DCO_conso_bio)
                P_precipite = BA_faible.P_precipite
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = BA_faible.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = BA_faible.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = BA_faible.nominal_MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
                If choix_primaire = False Then
                    For I = 1 To nb_etape_file_boues
                        TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_faible_EB
                    Next I
                Else
                    For I = 1 To nb_etape_file_boues
                        TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_faible_ED
                    Next I
                End If
            End If
            If choix_BA_prolongee = True Then
                If iteration = 1 Then
                    Dim BA_prolongee As New BA_aeration_prolongee
                End If
                Call BA_prolongee.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_apportee_methanol, DCO_conso_bio)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
                P_precipite = BA_prolongee.P_precipite
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = BA_prolongee.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = BA_prolongee.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = BA_prolongee.nominal_MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
                If choix_primaire = False Then
                    For I = 1 To nb_etape_file_boues
                        TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_prolongee_EB
                    Next I
                Else
                    For I = 1 To nb_etape_file_boues
                        TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_prolongee_ED
                    Next I
                End If
            End If
            If choix_HybAS = True Then
                If iteration = 1 Then
                    Dim traitement_hybas As New HybAS
                End If
                Call traitement_hybas.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_apportee_methanol, DCO_conso_bio)
                P_precipite = traitement_hybas.deltaP_precipitation
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = traitement_hybas.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = traitement_hybas.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = traitement_hybas.nominal_MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_HybAS
                Next I
            End If
            If choix_MBBR = True Then
                If iteration = 1 Then
                    Dim traitement_MBBR As New MBBR
                End If
                Call traitement_MBBR.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_apportee_methanol, DCO_conso_bio)
                P_precipite = 0
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = traitement_MBBR.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = traitement_MBBR.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = traitement_MBBR.boues_MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_MBBR
                Next I
            End If
            If choix_biostyr = True Then
                'stockage des flux arrivant sur le(s) biostyr(s)
                stockage_Q = Q
                stockage_DCO = DCO
                stockage_DBO = DBO
                stockage_MES = MES
                stockage_NK = NK
                stockage_NH4 = NH4
                stockage_NO3 = NO3
                stockage_Pt = Pt
                stockage_DCO_conso_bio = DCO_conso_bio
                For iteration_ES = 1 To nb_iteration_ES
                    'on réutilise les valeurs des flux arrivants sur le biostyr stockés dans les variables
                    Q = stockage_Q
                    DCO = stockage_DCO
                    DBO = stockage_DBO
                    MES = stockage_MES
                    NK = stockage_NK
                    NH4 = stockage_NH4
                    NO3 = stockage_NO3
                    Pt = stockage_Pt
                    DCO_conso_bio = stockage_DCO_conso_bio
                    'eaux sales retournées en tête du biostyr après traitement spécifique
                    For I = 1 To nb_eaux_sales_max
                        If choix_eaux_sales(I) = devenir_ES_traitement_separe_amont Then
                            'flag pour la digestion
                            choix_eaux_sales_2aire_amont = True
                            'gestion des flux
                            Q = Q + eaux_sales(I, repere_ES_Q)
                            DCO = DCO + eaux_sales(I, repere_ES_DCO)
                            DBO = DBO + eaux_sales(I, repere_ES_DBO)
                            MES = MES + eaux_sales(I, repere_ES_MES)
                            NK = NK + eaux_sales(I, repere_ES_NK)
                            NH4 = NH4 + eaux_sales(I, repere_ES_NH4)
                            NO3 = NO3 + eaux_sales(I, repere_ES_NO3)
                            Pt = Pt + eaux_sales(I, repere_ES_Pt)
                        End If
                    Next I
                    If iteration = 1 And iteration_ES = 1 Then
                        Dim biofiltre As New Biostyr
                    End If
                    Call biofiltre.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_conso_bio, eaux_sales, choix_eaux_sales, iteration_ES, boues_ratio_DBO_MV)
                    If pn_EffacementResultatsEnCours = True Then
                        GoTo calcul_interrompu
                    End If
                    'calcul des NK et du Pt présents dans les eaux sales
                    eaux_sales(repere_ES_biostyr, repere_ES_DBO) = eaux_sales(repere_ES_biostyr, repere_ES_DBO) + boues_ratio_DBO_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    eaux_sales(repere_ES_biostyr, repere_ES_NK) = eaux_sales(repere_ES_biostyr, repere_ES_NK) + boues_ratio_NK_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    eaux_sales(repere_ES_biostyr, repere_ES_Pt) = eaux_sales(repere_ES_biostyr, repere_ES_Pt) + boues_ratio_Pt_MES(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    'calcul des charges en NK et Pt
                    NK = NK + boues_ratio_NK_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) * MES
                    Pt = Pt - boues_ratio_Pt_MES(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    'correction pour le phosphore avec la partie soluble (on considère que ce qui reste est soluble et donc répartie de manière proportionnelle --> concentration égale)
                    eaux_sales(repere_ES_biostyr, repere_ES_Pt) = eaux_sales(repere_ES_biostyr, repere_ES_Pt) + Pt * eaux_sales(repere_ES_biostyr, repere_ES_Q) / (Q + eaux_sales(repere_ES_biostyr, repere_ES_Q))
                    Pt = Pt * Q / (Q + eaux_sales(repere_ES_biostyr, repere_ES_Q))
                    If choix_biostyr_PDN = True Then
                        If iteration = 1 And iteration_ES = 1 Then
                            Dim biofiltre_PDN As New Biostyr_PDN
                        End If
                        Call biofiltre_PDN.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, DCO_conso_bio, eaux_sales, choix_eaux_sales, iteration_ES, DCO_apportee_methanol)
                        If pn_EffacementResultatsEnCours = True Then
                            GoTo calcul_interrompu
                        End If
                    End If
                    If choix_decanteur_ES = True Then
                        If iteration = 1 And iteration_ES = 1 Then
                            Dim decanteur_ES As New Decantation_eaux_sales
                        End If
                        Call decanteur_ES.dimensionnement(eaux_sales, choix_eaux_sales, boues_II_detail, C_elimine, DBO_elimine, N_elimine, P_elimine)
                        If pn_EffacementResultatsEnCours = True Then
                            GoTo calcul_interrompu
                        End If
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = decanteur_ES.boues_MES
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = decanteur_ES.boues_Q
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = decanteur_ES.MV_MES
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
                        'GESTION DES BOUES ISSUES DES EAUX SALES (compositions gérées de manière similaire aux boues primaires)
                        For I = 1 To boues_epaissies(nb_max_epaississeur)
                            TableauRecapitulatifFluxBoues(I, boues_II, repere_MV_MES) = TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES)
                            TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_DCO_MES) = C_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                            TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_DBO_MES) = DBO_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                            TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_NK_MES) = N_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                            TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_Pt_MES) = P_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                        Next I
                    End If
                Next iteration_ES
                'on itère car les eaux sales reviennent en amont du biostyr
                If choix_eaux_sales_2aire_amont = True Then
                    nb_iteration_ES = pi_NOMBRE_ITERATION_RETOURS
                End If
                For I = 1 To nb_eaux_sales_max
                    If choix_eaux_sales(I) = devenir_ES_traitement_separe_aval Then
                        'flag pour la digestion
                        choix_eaux_sales_2aire_aval = True
                        'gestion des flux
                        Q = Q + eaux_sales(I, repere_ES_Q)
                        DCO = DCO + eaux_sales(I, repere_ES_DCO)
                        DBO = DBO + eaux_sales(I, repere_ES_DBO)
                        MES = MES + eaux_sales(I, repere_ES_MES)
                        NK = NK + eaux_sales(I, repere_ES_NK)
                        NH4 = NH4 + eaux_sales(I, repere_ES_NH4)
                        NO3 = NO3 + eaux_sales(I, repere_ES_NO3)
                        Pt = Pt + eaux_sales(I, repere_ES_Pt)
                    End If
                Next I
            End If
            If choix_biostyr = False Then  'si on n'a pas de biostyr
                'calcul des ratios de pollution dans les boues
                C_elimine = C_elimine + DCO_apportee_methanol - DCO
                DCO_conso_bio = C_elimine
                'N_elimine = N_elimine - NK
                P_elimine = P_elimine - Pt - P_precipite
                For I = 1 To boues_epaissies(nb_max_epaississeur)
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_MV_MES) = TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES)
                    '*************** MODIFIE LE 26/06/2012 PAR DCA ******************************
                    '*************** suite à des écarts au niveau de la DCO dans les boues, on fixe le ratio DCO/MES et on recalcule DCO_conso_bio **********
                    'TableauRecapitulatifFluxBoues(i, boues_II, repere_ratio_DCO_MES) = (C_elimine - DCO_conso_bio) / TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_DCO_MES) = boues_ratio_DCO_MV(TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_origine)) * TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES)
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_DBO_MES) = boues_ratio_DBO_MV(TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_origine)) * TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES)
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_NK_MES) = boues_ratio_NK_MV(TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_origine)) * TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES)
                    N_elimine = TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_ratio_NK_MES) * TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_Pt_MES) = P_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                Next I
                DCO_conso_bio = DCO_conso_bio - TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_ratio_DCO_MES) * TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                If choix_BA_faible = True Or choix_BA_prolongee = True Then
                    'faible charge ou aération prolongée: NK = NH4 + NK particulaire
                    NK = NK + TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_ratio_NK_MES) * MES
                ElseIf choix_BA_forte = True Or choix_BA_moyenne = True Then
                    'on soustrait la partie qui est assimilée et qui part dans les boues
                    NK = NK - N_elimine
                End If
            End If
        End If

        'TRAITEMENT TERTIAIRE
        If choix_tertiaire = True Then
            C_elimine = DCO
            DBO_elimine = DBO
            N_elimine = NK
            P_elimine = Pt
            If choix_biostyr_N_III = True Then
                'stockage des flux arrivant sur le(s) biostyr(s)
                stockage_Q = Q
                stockage_DCO = DCO
                stockage_DBO = DBO
                stockage_MES = MES
                stockage_NK = NK
                stockage_NH4 = NH4
                stockage_NO3 = NO3
                stockage_Pt = Pt
                stockage_DCO_conso_bio = DCO_conso_bio
                For iteration_ES = 1 To nb_iteration_ES
                    'on réutilise les valeurs des flux arrivants sur le biostyr stockés dans les variables
                    Q = stockage_Q
                    DCO = stockage_DCO
                    DBO = stockage_DBO
                    MES = stockage_MES
                    NK = stockage_NK
                    NH4 = stockage_NH4
                    NO3 = stockage_NO3
                    Pt = stockage_Pt
                    DCO_conso_bio = stockage_DCO_conso_bio
                    'eaux sales retournées en tête du biostyr après traitement spécifique
                    For I = 1 To nb_eaux_sales_max
                        If choix_eaux_sales(I) = devenir_ES_traitement_separe_amont Then
                            'flag pour la digestion
                            choix_eaux_sales_3aire_amont = True
                            'gestion des flux
                            Q = Q + eaux_sales(I, repere_ES_Q)
                            DCO = DCO + eaux_sales(I, repere_ES_DCO)
                            DBO = DBO + eaux_sales(I, repere_ES_DBO)
                            MES = MES + eaux_sales(I, repere_ES_MES)
                            NK = NK + eaux_sales(I, repere_ES_NK)
                            NH4 = NH4 + eaux_sales(I, repere_ES_NH4)
                            NO3 = NO3 + eaux_sales(I, repere_ES_NO3)
                            Pt = Pt + eaux_sales(I, repere_ES_Pt)
                        End If
                    Next I
                    If iteration = 1 And iteration_ES = 1 Then
                        Dim biofiltre_III As New Biostyr_N_III
                    End If
                    Call biofiltre_III.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_conso_bio, eaux_sales, choix_eaux_sales, iteration_ES)
                    If pn_EffacementResultatsEnCours = True Then
                        GoTo calcul_interrompu
                    End If
                    'calcul des NK et du Pt présents dans les eaux sales
                    eaux_sales(repere_ES_biostyr, repere_ES_DBO) = eaux_sales(repere_ES_biostyr, repere_ES_DBO) + boues_ratio_DBO_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    eaux_sales(repere_ES_biostyr, repere_ES_NK) = eaux_sales(repere_ES_biostyr, repere_ES_NK) + boues_ratio_NK_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    eaux_sales(repere_ES_biostyr, repere_ES_Pt) = eaux_sales(repere_ES_biostyr, repere_ES_Pt) + boues_ratio_Pt_MES(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    'calcul des charges en NK et Pt
                    NK = NK + boues_ratio_NK_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) * MES
                    Pt = Pt - boues_ratio_Pt_MES(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    'correction pour le phosphore avec la partie soluble (on considère que ce qui reste est soluble et donc répartie de manière proportionnelle --> concentration égale)
                    eaux_sales(repere_ES_biostyr, repere_ES_Pt) = eaux_sales(repere_ES_biostyr, repere_ES_Pt) + Pt * eaux_sales(repere_ES_biostyr, repere_ES_Q) / (Q + eaux_sales(repere_ES_biostyr, repere_ES_Q))
                    Pt = Pt * Q / (Q + eaux_sales(repere_ES_biostyr, repere_ES_Q))
                    If choix_biostyr_PDN_III = True Then
                        If iteration = 1 And iteration_ES = 1 Then
                            Dim biofiltre_PDN_III As New Biostyr_PDN_III
                        End If
                        Call biofiltre_PDN_III.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, DCO_conso_bio, eaux_sales, choix_eaux_sales, iteration_ES, DCO_apportee_methanol)
                        If pn_EffacementResultatsEnCours = True Then
                            GoTo calcul_interrompu
                        End If
                    End If
                    If choix_decanteur_ES = True Then
                        If iteration = 1 And iteration_ES = 1 Then
                            Dim decanteur_ES_III As New Decantation_eaux_sales
                        End If
                        Call decanteur_ES_III.dimensionnement(eaux_sales, choix_eaux_sales, boues_III_detail, C_elimine, DBO_elimine, N_elimine, P_elimine)
                        If pn_EffacementResultatsEnCours = True Then
                            GoTo calcul_interrompu
                        End If
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES) = decanteur_ES_III.boues_MES
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_Q) = decanteur_ES_III.boues_Q
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MV_MES) = decanteur_ES_III.MV_MES
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_verif_flux) = 1
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_flux_in) = 1
                        For I = 1 To boues_epaissies(nb_max_epaississeur)
                            TableauRecapitulatifFluxBoues(I, boues_III, repere_MV_MES) = TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MV_MES)
                            TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_DCO_MES) = C_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                            TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_DBO_MES) = DBO_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                            TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_NK_MES) = N_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                            TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_Pt_MES) = P_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                        Next I
                    End If
                Next iteration_ES
                'on itère car les eaux sales reviennent en amont du biostyr
                If choix_eaux_sales_3aire_amont = True Then
                    nb_iteration_ES = pi_NOMBRE_ITERATION_RETOURS
                End If
                For I = 1 To nb_eaux_sales_max
                    If choix_eaux_sales(I) = devenir_ES_traitement_separe_aval Then
                        'flag pour la digestion
                        choix_eaux_sales_3aire_aval = True
                        'gestion des flux
                        Q = Q + eaux_sales(I, repere_ES_Q)
                        DCO = DCO + eaux_sales(I, repere_ES_DCO)
                        DBO = DBO + eaux_sales(I, repere_ES_DBO)
                        MES = MES + eaux_sales(I, repere_ES_MES)
                        NK = NK + eaux_sales(I, repere_ES_NK)
                        NH4 = NH4 + eaux_sales(I, repere_ES_NH4)
                        NO3 = NO3 + eaux_sales(I, repere_ES_NO3)
                        Pt = Pt + eaux_sales(I, repere_ES_Pt)
                    End If
                Next I
            End If

            If choix_decanteur_III = True Then
                If iteration = 1 Then
                    Dim decanteur_III As New Decantation_III
                End If
                Call decanteur_III.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
                P_precipite = decanteur_III.P_precipite
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_III, repere_origine) = III_decantation
                Next I
                TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES) = decanteur_III.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_Q) = decanteur_III.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MV_MES) = decanteur_III.MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_flux_in) = 1
                'calcul des ratios de pollution dans les boues qui sont constantes jusqu'à la digestion
                C_elimine = C_elimine - DCO
                DBO_elimine = DBO_elimine - DBO
                N_elimine = N_elimine - NK
                P_elimine = P_elimine - Pt - P_precipite
                For I = 1 To boues_epaissies(nb_max_epaississeur)
                    TableauRecapitulatifFluxBoues(I, boues_III, repere_MV_MES) = TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MV_MES)
                    TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_DCO_MES) = C_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                    TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_DBO_MES) = DBO_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                    TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_NK_MES) = N_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                    TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_Pt_MES) = P_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                Next I
            End If
            'Discfilter
            If choix_discfilter = True Then
                If iteration = 1 Then
                    Dim Discfilter As New Discfilter_III
                End If
                Call Discfilter.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, eaux_sales_discfilter)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
            End If
            'Filtration tertiaire sur sable
            If choix_filtrasable = True Then
                If iteration = 1 Then
                    Dim Filtre_sable As New Filtration_sable
                End If
                Call Filtre_sable.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, eaux_sales_filtrasable)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
            End If
        End If

        'Désinfection
        If choix_desinfection = True Then
            'UV
            If choix_UV = True Then
                If iteration = 1 Then
                    Dim Traitement_UV As New Desinfection_UV
                End If
                Call Traitement_UV.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
            End If
            'Chloration
            If choix_chloration = True Then
                If iteration = 1 Then
                    Dim chloration As New Desinfection_Cl
                End If
                Call chloration.dimensionnement(Q, DCO, DBO, MES, NK, NH4, NO3, Pt)
                Cl2_gazeux_flux_kgj = Cl2_gazeux_flux_kgj + chloration.Cl2_gazeux_pur_kgj
                NaOClpur_flux = NaOClpur_flux + chloration.Cl2_HClO_pur_kgj
                NaOClpur_eau_flux = NaOClpur_eau_flux + chloration.Cl2_HClO_pur_kgj
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
            End If
        End If


        ''''''''''''
        'FILE BOUES'
        ''''''''''''

        'RECAPITULATIF DES BOUES EN ENTREE
        Call recapitulatif_boues(TableauRecapitulatifFluxBoues, boues_pollution_soluble, graisses_particulaire_ratio_interne, boues_ratio_NK_MV, boues_ratio_Pt_MES, boues_ratio_DCO_MV, boues_ratio_DBO_MV, graisses_internes, nb_max_epaississeur, boues_Shunt, choix_eaux_sales_1aire, boues_I_detail)

        For I = 1 To retour_caracteristique_nb
            retour_flux(I) = 0
            retour_digestion(I) = 0
            retour_athos(I) = 0
            retour_flux_soluble(I) = 0
            retour_digestion_soluble(I) = 0
            retour_athos_soluble(I) = 0
        Next I
        
        
        'BIOLIX
        If choix_biolix = True Then
            If iteration = 1 Then
                Dim biolix As New Biolix_graisses
                Call biolix.initialisation_procede
            End If
            Call biolix.dimensionnement(TableauRecapitulatifFluxBoues, boues_pollution_soluble)
            If pn_EffacementResultatsEnCours = True Then
                GoTo calcul_interrompu
            End If
        End If


        'EPAISSISSEMENT
        If choix_epaississement = True Then
            If iteration = 1 Then
                epaississement_choix_nb = Feuil25.Range("epaississement_choix_nb").Cells(1, pi_FiliereConsideree)
                For I = 1 To epaississement_choix_nb
                    Dim epaississeur(nb_max_epaississeur) As New Epaississement
                    Call epaississeur(I).lecture_choix(I)
                    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
                        Call epaississeur(I).lecture_valeurs_forcees(I)
                    End If
                    Call epaississeur(I).hypotheses
                    Call epaississeur(I).attribution_valeur_par_defaut
                Next I
            End If
            For I = 1 To epaississement_choix_nb
                Call epaississeur(I).dimensionnement(TableauRecapitulatifFluxBoues, boues_pollution_soluble, I, retour_flux, retour_flux_soluble, boues_I_detail, boues_II_detail, boues_III_detail, epaississement_before_thelys)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
            Next I
        End If

        'DIGESTION
        If pi_ChoixDigestion = True Then
            If choix_dig_simple = True Then
                If iteration = 1 Then
                    Dim digesteur_simple As New Digestion_simple
                End If
                Call digesteur_simple.dimensionnement(TableauRecapitulatifFluxBoues, choix_eaux_sales_1aire, boues_I_detail, boues_II_detail, boues_III_detail, graisses_particulaire_ratio_interne, DCO_biogaz, boues_pollution_soluble, besoins_thermiques, biogaz)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
            End If
            If choix_biothelys = True Then
                If iteration = 1 Then
                    Dim digesteur_biothelys As New Biothelys
                    Call digesteur_biothelys.initialisation_procede
                End If
                Call digesteur_biothelys.dimensionnement(TableauRecapitulatifFluxBoues, boues_pollution_soluble, graisses_particulaire_ratio_interne, choix_eaux_sales_1aire, boues_I_detail, boues_II_detail, boues_III_detail, DCO_biogaz, besoins_thermiques, disponibilites_thermiques, biogaz, pression_vapeur, nb_vapeur, epaississement_before_thelys)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
            End If
            If choix_exelys_DLD = True Then
                If iteration = 1 Then
                    Dim digesteur_DLD As New Exelys_DLD
                    Call digesteur_DLD.initialisation_procede
                End If
                Call digesteur_DLD.dimensionnement(TableauRecapitulatifFluxBoues, graisses_particulaire_ratio_interne, choix_eaux_sales_1aire, boues_I_detail, boues_II_detail, boues_III_detail, DCO_biogaz, boues_pollution_soluble, retour_flux, retour_digestion, retour_flux_soluble, retour_digestion_soluble, besoins_thermiques, disponibilites_thermiques, biogaz, pression_vapeur, nb_vapeur, epaississement_before_thelys)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
            End If
        End If

        'ATHOS
        If choix_athos = True Then
            If iteration = 1 Then
                Dim OVH_athos As New CLS_BouesAthos
                Call OVH_athos.initialisation_procede
            End If
            Call OVH_athos.dimensionnement(TableauRecapitulatifFluxBoues, retour_flux, retour_digestion, retour_athos, retour_flux_soluble, retour_digestion_soluble, retour_athos_soluble, boues_pollution_soluble, DCO_oxydation)
            besoins_thermiques(energie_biogaz) = besoins_thermiques(energie_biogaz) + OVH_athos.pd_TraitementFumeesRTOConsommationBiogaz_kWhPCIj
            besoins_thermiques(energie_combustible) = besoins_thermiques(energie_combustible) + OVH_athos.pd_TraitementFumeesRTOConsommationCombustibleExterne_kWhPCIj
            If pn_EffacementResultatsEnCours = True Then
                GoTo calcul_interrompu
            End If
        End If

        'DESHYDRATATION
        If choix_deshydratation = True Then
            If iteration = 1 Then
                deshydratation_choix_nb = Feuil28.Range("deshydratation_choix_nb").Cells(1, pi_FiliereConsideree)
                For I = 1 To deshydratation_choix_nb
                    Dim deshydrat(nb_max_deshydratation) As New Deshydratation
                    Call deshydrat(I).lecture_choix(I)
                    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
                        Call deshydrat(I).lecture_valeurs_forcees(I)
                    End If
                    Call deshydrat(I).hypotheses
                    Call deshydrat(I).attribution_valeur_par_defaut
                Next I
            End If
            For I = 1 To deshydratation_choix_nb
                Call deshydrat(I).dimensionnement(TableauRecapitulatifFluxBoues, I, retour_flux, retour_digestion, retour_flux_soluble, retour_digestion_soluble, boues_pollution_soluble, boues_I_detail, boues_II_detail, boues_III_detail, retour_athos)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
            Next I
        End If

        'CHAULAGE
        If choix_chaulage = True Then
            If iteration = 1 Then
                chaulage_choix_nb = Feuil42.Range("chaulage_choix_nb").Cells(1, pi_FiliereConsideree)
                For I = 1 To chaulage_choix_nb
                    Dim chaulage(nb_max_chaulage) As New Chaulage_boues
                    Call chaulage(I).lecture_choix(I)
                    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
                        Call chaulage(I).lecture_valeurs_forcees(I)
                    End If
                    Call chaulage(I).hypotheses
                Next I
            End If
            For I = 1 To chaulage_choix_nb
                Call chaulage(I).dimensionnement(TableauRecapitulatifFluxBoues, I, retour_flux, boues_I_detail, boues_II_detail, boues_III_detail, boues_pollution_soluble)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
            Next I
        End If

        'SECHAGE DES BOUES
        If choix_sechage = True Then

            'Secheur Inos
            If choix_secheur_inos = True Then
                If iteration = 1 Then
                    sechage_inos_choix_nb = Feuil38.Range("sechage_inos_choix_nb").Cells(1, pi_FiliereConsideree)
                    For I = 1 To sechage_inos_choix_nb
                        Dim secheur_inos(nb_max_sechage_inos) As New sechage_inos
                        Call secheur_inos(I).lecture_choix(I)
                        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
                            Call secheur_inos(I).lecture_valeurs_forcees(I)
                        End If
                        Call secheur_inos(I).hypotheses(I)
                        Call secheur_inos(I).attribution_valeur_par_defaut
                    Next I
                End If
                For I = 1 To sechage_inos_choix_nb
                    Call secheur_inos(I).dimensionnement(TableauRecapitulatifFluxBoues, I, retour_flux, boues_I_detail, boues_II_detail, boues_III_detail, boues_pollution_soluble, besoins_thermiques, disponibilites_thermiques)
                    If pn_EffacementResultatsEnCours = True Then
                        GoTo calcul_interrompu
                    End If
                Next I
            End If

            'Secheur Bioco
            If choix_secheur_bioco = True Then
                If iteration = 1 Then
                    sechage_bioco_choix_nb = Feuil35.Range("sechage_bioco_choix_nb").Cells(1, pi_FiliereConsideree)
                    For I = 1 To sechage_bioco_choix_nb
                        Dim secheur_bioco(nb_max_sechage_bioco) As New Sechage_bioco
                        Call secheur_bioco(I).lecture_choix(I)
                        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
                            Call secheur_bioco(I).lecture_valeurs_forcees(I)
                        End If
                        Call secheur_bioco(I).hypotheses
                        Call secheur_bioco(I).attribution_valeur_par_defaut
                    Next I
                End If
                For I = 1 To sechage_bioco_choix_nb
                    Call secheur_bioco(I).dimensionnement(TableauRecapitulatifFluxBoues, I, retour_flux, boues_pollution_soluble, besoins_thermiques, disponibilites_thermiques)
                    If pn_EffacementResultatsEnCours = True Then
                        GoTo calcul_interrompu
                    End If
                Next I
            End If

            'Secheur Autre
            If choix_secheur_autre = True Then
                If iteration = 1 Then
                    sechage_choix_nb = Feuil31.Range("sechage_autre_choix_nb").Cells(1, pi_FiliereConsideree)
                    For I = 1 To sechage_choix_nb
                        Dim secheur_autre(nb_max_sechage) As New Sechage_thermique
                        Call secheur_autre(I).lecture_choix(I)
                        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
                            Call secheur_autre(I).lecture_valeurs_forcees(I)
                        End If
                        Call secheur_autre(I).hypotheses
                        Call secheur_autre(I).attribution_valeur_par_defaut
                    Next I
                End If
                For I = 1 To sechage_choix_nb
                    Call secheur_autre(I).dimensionnement(TableauRecapitulatifFluxBoues, I, retour_flux, boues_pollution_soluble, besoins_thermiques, disponibilites_thermiques)
                    If pn_EffacementResultatsEnCours = True Then
                        GoTo calcul_interrompu
                    End If
                Next I
            End If
        End If

        'INCINERATION
        If choix_incineration = True Then
            If iteration = 1 Then
                Dim incineration As New Incineration_boues
            End If
            Call incineration.dimensionnement(TableauRecapitulatifFluxBoues, boues_pollution_soluble, besoins_thermiques, disponibilites_thermiques, biogaz(teneur_CH4), DCO_oxydation)
            If pn_EffacementResultatsEnCours = True Then
                GoTo calcul_interrompu
            End If
        End If

        'TRAITEMENT SEPARE DES RETOURS DE TETE
        If choix_traitement_retours = True Then
            If choix_MAP_retours = True Then
                If iteration = 1 Then
                    Dim traitement_MAP_retours As New CLS_RetoursMAP
                End If
                Call traitement_MAP_retours.dimensionnement(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble)
            End If
            If choix_ANITA_Mox = True Then
                If iteration = 1 Then
                    Dim traitement_Mox As New ANITA_Mox
                End If
                Call traitement_Mox.dimensionnement(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble)
            End If
            If choix_ANITA_Shunt = True Then
                If iteration = 1 Then
                    Dim traitement_Shunt As New ANITA_Shunt
                End If
                Call traitement_Shunt.dimensionnement(retour_flux, retour_digestion, retour_flux_soluble, retour_digestion_soluble, retour_athos, retour_athos_soluble, boues_Shunt)
            End If
        End If

        'VERIFICATION FLUX DE BOUES
        If iteration = 1 Then
            Call verification_boues(TableauRecapitulatifFluxBoues, warning, cendres_evacuees_tMSj, REFIB_evacues_tMSj, boues_evacuees_Q, graisses_evacuees_Q)
            If warning = True Then Exit Sub
        End If



        'GESTION DE L'ENERGIE THERMIQUE
        If iteration = 1 Then
            If choix_analyse_pinch = True Then
                Dim energie_thermique_PINCH As New Gestion_energie_thermique_PINCH
            Else
                Dim energie_thermique As New Gestion_energie_thermique
            End If
        End If
        If choix_analyse_pinch = True Then
            'Call energie_thermique_PINCH.dimensionnement(biogaz, electricite_verte, type_valorisation_biogaz, valorisation_chaudiere, ratio_biogaz_valorise)
        Else
            Call energie_thermique.dimensionnement(besoins_thermiques, pression_vapeur, biogaz, disponibilites_thermiques, electricite_verte, type_valorisation_biogaz, valorisation_chaudiere, ratio_biogaz_valorise)
        End If
        If pn_EffacementResultatsEnCours = True Then
            GoTo calcul_interrompu
        End If

    Next iteration


    'GoTo bypass_reel '************************************************************************************************
    '###########################################################################################################################################################################################

    ''''''''''''''''''''
    'CONDITIONS REELLES'
    ''''''''''''''''''''

    For I = 1 To nb_eaux_sales_max
        For j = 1 To nb_parametres_eaux_sales
            eaux_sales(I, j) = 0
        Next j
    Next I
    For j = 1 To nb_parametres_eaux_sales
        eaux_sales_discfilter(j) = 0
        eaux_sales_filtrasable(j) = 0
    Next j

    'traitement des sulfures
    'If choix_traitement_sulfures_curatif = True Then
    Call traitement_HS.fonctionnement_reel(SH_nominal, pd_PourcentageChargeReelleSH)
    HS_strippe_kgj = traitement_HS.SH_strippe_kgj
    HS_traite_kgj = traitement_HS.SH_traite_kgj
    HS_O2_liquide_kgj = traitement_HS.O2_liquide_kgj
    HS_FeCl3pur_kgj = traitement_HS.FeCl3pur_kgj
    HS_Ca_2NO3_pur_flux = traitement_HS.Ca_2NO3_kgj
    HS_H2O2pur_flux = traitement_HS.H2O2_kgj
    Call traitement_HS.calcul_consommation_electrique
    Set traitement_HS = Nothing    'on le détruit à la fin
    'End If

    'relèvements (hors boucles de calcul car indépendant)
    If choix_relevement = True Then
        relevement_choix_nb = Feuil7.Range("relevement_choix_nb").Cells(1, pi_FiliereConsideree)
        For I = 1 To relevement_choix_nb
            'on créé un objet de classe "relevement"
            Dim relevage As New relevement
            'on l'appelle
            Call relevage.calcul(I)
            electricite_consommee = electricite_consommee + relevage.electricite
            electricite_fixe(electricite_postes_repere_relevement) = electricite_fixe(electricite_postes_repere_relevement)
            electricite_variable(electricite_postes_repere_relevement) = electricite_variable(electricite_postes_repere_relevement) + relevage.electricite
            Set relevage = Nothing
        Next I
    End If

    'prétraitements (hors boucles de calcul car indépendant)
    If choix_pretraitement = True Then
        If choix_degrillage = True Then
            Call degrilleur.fonctionnement_reel
            QuantiteRefusDegrillage_kgj = degrilleur.pd_QuantiteRefusDegrillage_kgj
            Call degrilleur.calcul_consommation_electrique
            electricite_consommee = electricite_consommee + degrilleur.pd_ConsommationElectrique_kWhj
            electricite_fixe(electricite_postes_repere_pretraitement) = electricite_fixe(electricite_postes_repere_pretraitement)
            electricite_variable(electricite_postes_repere_pretraitement) = electricite_variable(electricite_postes_repere_pretraitement) + degrilleur.pd_ConsommationElectrique_kWhj
            Set degrilleur = Nothing
        End If
        If choix_dessablage = True Then
            Call dessableur_deshuileur.calcul(graisses_internes, boues_ratio_DCO_MV, boues_ratio_NK_MV, boues_ratio_Pt_MES, pd_PourcentageChargeReelleDBO * pd_CapaciteSTEP_EH, vidange_MES_reel)
            Call dessableur_deshuileur.calcul_consommation_electrique
            electricite_consommee = electricite_consommee + dessableur_deshuileur.electricite
            electricite_fixe(electricite_postes_repere_pretraitement) = electricite_fixe(electricite_postes_repere_pretraitement) + dessableur_deshuileur.electricite
            electricite_variable(electricite_postes_repere_pretraitement) = electricite_variable(electricite_postes_repere_pretraitement)
            Set dessableur_deshuileur = Nothing
        End If
    End If

    For iteration = 1 To pi_NOMBRE_ITERATION_RETOURS

        'REINITIALISATION DES VARIABLES D'ENTREE
        Q = Q_nominal * pd_PourcentageChargeReelleDebitVolumique + vidange_Q_reel + retour_flux(repere_ret_Q) - graisses_internes(repere_graisse_Q)
        DCO = DCO_nominal * pd_PourcentageChargeReelleDCO + vidange_DCO_reel + retour_flux(repere_ret_DCO) - graisses_internes(repere_graisse_DCO_MES) * graisses_internes(repere_graisse_MS)
        DBO = DBO_nominal * pd_PourcentageChargeReelleDBO + vidange_DBO_reel + retour_flux(repere_ret_DBO) - graisses_internes(repere_graisse_DBO_MES) * graisses_internes(repere_graisse_MS)
        MES = MES_nominal * pd_PourcentageChargeReelleMES + vidange_MES_reel + retour_flux(repere_ret_MES) - graisses_internes(repere_graisse_MS)
        NK = NK_nominal * pd_PourcentageChargeReelleNK + vidange_NK_reel + retour_flux(repere_ret_NK) - graisses_internes(repere_graisse_NK_MES) * graisses_internes(repere_graisse_MS)
        NH4 = NH4_nominal * pd_PourcentageChargeReelleNH4 + vidange_NH4_reel + retour_flux(repere_ret_NH4)
        NO3 = retour_flux(repere_ret_NO3)
        Pt = Pt_nominal * pd_PourcentageChargeReellePt + vidange_Pt_reel + retour_flux(repere_ret_Pt) - graisses_internes(repere_graisse_Pt_MES) * graisses_internes(repere_graisse_MS)
        Sh = SH_nominal * pd_PourcentageChargeReelleSH - HS_strippe_kgj - HS_traite_kgj

        If choix_discfilter = True Then   'on réinjecte les eaux sales du discfilter
            Q = Q + eaux_sales_discfilter(repere_ES_Q)
            DCO = DCO + eaux_sales_discfilter(repere_ES_DCO)
            DBO = DBO + eaux_sales_discfilter(repere_ES_DBO)
            MES = MES + eaux_sales_discfilter(repere_ES_MES)
            NK = NK + eaux_sales_discfilter(repere_ES_NK)
            NH4 = NH4 + eaux_sales_discfilter(repere_ES_NH4)
            NO3 = NO3 + eaux_sales_discfilter(repere_ES_NO3)
            Pt = Pt + eaux_sales_discfilter(repere_ES_Pt)
        End If

        If choix_filtrasable = True Then   'on réinjecte les eaux sales du filtre à sable
            Q = Q + eaux_sales_filtrasable(repere_ES_Q)
            DCO = DCO + eaux_sales_filtrasable(repere_ES_DCO)
            DBO = DBO + eaux_sales_filtrasable(repere_ES_DBO)
            MES = MES + eaux_sales_filtrasable(repere_ES_MES)
            NK = NK + eaux_sales_filtrasable(repere_ES_NK)
            NH4 = NH4 + eaux_sales_filtrasable(repere_ES_NH4)
            NO3 = NO3 + eaux_sales_filtrasable(repere_ES_NO3)
            Pt = Pt + eaux_sales_filtrasable(repere_ES_Pt)
        End If

        Q_retour = retour_flux(repere_ret_Q)

        'REINITIALISATION DES FLUX DE BOUES
        For I = 1 To nb_etape_file_boues
            For j = 1 To nb_type_boues
                For k = 1 To nb_parametres_boues
                    TableauRecapitulatifFluxBoues(I, j, k) = 0
                Next k
            Next j
        Next I
        ReDim flux_thermique_process(1 To flux_thermique_process_nb_caracteristiques, 1 To 1)
        ReDim utilites_flux_energie(1 To flux_thermique_utilites_nb_caracteristiques, 1 To 1)
        nb_flux_utilites = 0

        DCO_conso_bio = 0
        DCO_apportee_methanol = 0
        DCO_oxydation = 0
        polymere_flux = 0
        For I = 1 To polymere_nb_type
            polymere_eau_pur_kgj(I) = 0
            polymere_boues_pur_kgj(I) = 0
        Next I
        chaux_eteinte_flux = 0
        chaux_vive_flux = 0
        FeCl3_flux = HS_FeCl3pur_kgj
        FeCl3_eau_pur_kgj = HS_FeCl3pur_kgj
        FeCl3_boues_pur_kgj = 0
        methanol_flux = 0
        H2SO4pur_flux = 0
        NaOHpur_flux = 0
        NaHCO3_flux = 0
        NaOClpur_flux = 0
        NaOClpur_eau_flux = 0
        NaOClpur_desodo_flux = 0
        NaHSO3pur_flux = 0
        electricite_verte = 0
        O2_liquide_flux = HS_O2_liquide_kgj
        Ca_2NO3_pur_flux = HS_Ca_2NO3_pur_flux
        H2O2pur_flux = HS_H2O2pur_flux
        charbon_actif_flux = 0
        Cl2_gazeux_flux_kgj = 0
        CuSO4_flux = 0
        MgCl2_flux = 0
        ConsommationAmmoniaquePur_kgj = 0
        ConsommationUreePur_kgj = 0
        
        nb_vapeur = 0
        'REINITIALISATION DES ENERGIES
        For I = 1 To nb_niveaux_energie
            besoins_thermiques(I) = 0
            disponibilites_thermiques(I) = 0
        Next I


        ''''''''''
        'FILE EAU'
        ''''''''''

        'TRAITEMENT PRIMAIRE
        'eaux sales retournées en tête de primaire sans traitement séparé
        For I = 1 To nb_eaux_sales_max
            If choix_eaux_sales(I) = devenir_ES_primaire Then
                'flag pour la digestion
                choix_eaux_sales_1aire = True
                'gestion des flux
                Q = Q + eaux_sales(I, repere_ES_Q)
                DCO = DCO + eaux_sales(I, repere_ES_DCO)
                DBO = DBO + eaux_sales(I, repere_ES_DBO)
                MES = MES + eaux_sales(I, repere_ES_MES)
                NK = NK + eaux_sales(I, repere_ES_NK)
                NH4 = NH4 + eaux_sales(I, repere_ES_NH4)
                NO3 = NO3 + eaux_sales(I, repere_ES_NO3)
                Pt = Pt + eaux_sales(I, repere_ES_Pt)
            End If
        Next I
        If choix_primaire = True Then
            C_elimine = DCO
            DBO_elimine = DBO
            N_elimine = NK
            P_elimine = Pt
            P_precipite = 0
            If choix_decanteur_simple = True Then
                Call decanteur_simple.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Q_retour, eaux_sales, choix_eaux_sales, boues_I_detail, eaux_sales_discfilter)
                P_precipite = 0
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_I, repere_origine) = I_simple
                Next I
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MES) = decanteur_simple.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_Q) = decanteur_simple.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MV_MES) = decanteur_simple.MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_flux_in) = 1
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call decanteur_simple.calcul_consommation_electrique
                    electricite_consommee = electricite_consommee + decanteur_simple.electricite
                    electricite_fixe(electricite_postes_repere_primaire) = electricite_fixe(electricite_postes_repere_primaire) + decanteur_simple.electricite_fixe
                    electricite_variable(electricite_postes_repere_primaire) = electricite_variable(electricite_postes_repere_primaire) + decanteur_simple.electricite - decanteur_simple.electricite_fixe
                    Set decanteur_simple = Nothing
                End If
            End If
            If choix_decanteur_reactif = True Then
                Call decanteur_reactif.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Q_retour, eaux_sales, choix_eaux_sales, boues_I_detail, eaux_sales_discfilter)
                P_precipite = decanteur_reactif.P_precipite
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_I, repere_origine) = I_reactif
                Next I
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MES) = decanteur_reactif.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_Q) = decanteur_reactif.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MV_MES) = decanteur_reactif.MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_flux_in) = 1
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call decanteur_reactif.calcul_consommation_electrique
                    electricite_consommee = electricite_consommee + decanteur_reactif.electricite
                    electricite_fixe(electricite_postes_repere_primaire) = electricite_fixe(electricite_postes_repere_primaire) + decanteur_reactif.electricite_fixe
                    electricite_variable(electricite_postes_repere_primaire) = electricite_variable(electricite_postes_repere_primaire) + decanteur_reactif.electricite - decanteur_reactif.electricite_fixe
                    polymere_flux = polymere_flux + decanteur_reactif.dosage_polymere * Q / 1000
                    polymere_eau_pur_kgj(decanteur_reactif.choix_polymere_type) = polymere_eau_pur_kgj(decanteur_reactif.choix_polymere_type) + decanteur_reactif.dosage_polymere * Q / 1000
                    FeCl3_flux = FeCl3_flux + decanteur_reactif.dosage_FeCl3 * Q / 1000
                    FeCl3_eau_pur_kgj = FeCl3_eau_pur_kgj + decanteur_reactif.dosage_FeCl3 * Q / 1000
                    Set decanteur_reactif = Nothing
                End If
            End If
            'calcul des ratios de pollution dans les boues qui sont constantes jusqu'à la digestion
            C_elimine = C_elimine - DCO
            DBO_elimine = DBO_elimine - DBO
            N_elimine = N_elimine - NK
            P_elimine = P_elimine - Pt - P_precipite
            For I = 1 To boues_epaissies(nb_max_epaississeur)
                TableauRecapitulatifFluxBoues(I, boues_I, repere_MV_MES) = TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MV_MES)
                TableauRecapitulatifFluxBoues(I, boues_I, repere_ratio_DCO_MES) = C_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MES)
                TableauRecapitulatifFluxBoues(I, boues_I, repere_ratio_DBO_MES) = DBO_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MES)
                TableauRecapitulatifFluxBoues(I, boues_I, repere_ratio_NK_MES) = N_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MES)
                TableauRecapitulatifFluxBoues(I, boues_I, repere_ratio_Pt_MES) = P_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_I, repere_MES)
            Next I
        End If

        'TRAITEMENT SECONDAIRE
        If choix_secondaire = True Then
            'On ajoute les boues de biolix
            If choix_biolix = True Then
                Q = Q + TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_Q)
                DCO = DCO + TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_ratio_DCO_MES) * TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_MES)
                DBO = DBO + TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_ratio_DBO_MES) * TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_MES)
                MES = MES + TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_MES)
                NK = NK + TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_ratio_NK_MES) * TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_MES)
                Pt = Pt + TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_ratio_Pt_MES) * TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_MES)
            End If
            'On stocke les valeurs de pollution en entrée
            C_elimine = DCO
            'N_elimine = NK
            P_elimine = Pt
            P_precipite = 0
            If choix_BA_forte = True Then
                Call BA_forte.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_conso_bio)
                P_precipite = 0
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_forte
                Next I
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = BA_forte.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = BA_forte.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = BA_forte.reel_MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call BA_forte.calcul_consommation_electrique(electricite_II_aeration)
                    electricite_consommee = electricite_consommee + BA_forte.electricite
                    electricite_fixe(electricite_postes_repere_secondaire) = electricite_fixe(electricite_postes_repere_secondaire) + BA_forte.electricite_fixe
                    electricite_variable(electricite_postes_repere_secondaire) = electricite_variable(electricite_postes_repere_secondaire) + BA_forte.electricite - BA_forte.electricite_fixe
                    Set BA_forte = Nothing
                End If
            End If
            If choix_BA_moyenne = True Then
                Call BA_moyenne.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_conso_bio)
                P_precipite = 0
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_moyenne
                Next I
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = BA_moyenne.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = BA_moyenne.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = BA_moyenne.reel_MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call BA_moyenne.calcul_consommation_electrique(electricite_II_aeration)
                    electricite_consommee = electricite_consommee + BA_moyenne.electricite
                    electricite_fixe(electricite_postes_repere_secondaire) = electricite_fixe(electricite_postes_repere_secondaire) + BA_moyenne.electricite_fixe
                    electricite_variable(electricite_postes_repere_secondaire) = electricite_variable(electricite_postes_repere_secondaire) + BA_moyenne.electricite - BA_moyenne.electricite_fixe
                    Set BA_moyenne = Nothing
                End If
            End If
            If choix_BA_faible = True Then
                Call BA_faible.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_apportee_methanol, DCO_conso_bio)
                P_precipite = BA_faible.P_precipite
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = BA_faible.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = BA_faible.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = BA_faible.reel_MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
                If choix_primaire = False Then
                    For I = 1 To nb_etape_file_boues
                        TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_faible_EB
                    Next I
                Else
                    For I = 1 To nb_etape_file_boues
                        TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_faible_ED
                    Next I
                End If
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call BA_faible.calcul_consommation_electrique(electricite_II_aeration)
                    electricite_consommee = electricite_consommee + BA_faible.electricite
                    electricite_fixe(electricite_postes_repere_secondaire) = electricite_fixe(electricite_postes_repere_secondaire) + BA_faible.electricite_fixe
                    electricite_variable(electricite_postes_repere_secondaire) = electricite_variable(electricite_postes_repere_secondaire) + BA_faible.electricite - BA_faible.electricite_fixe
                    FeCl3_flux = FeCl3_flux + BA_faible.FeCl3_pur
                    FeCl3_eau_pur_kgj = FeCl3_eau_pur_kgj + BA_faible.FeCl3_pur
                    methanol_flux = methanol_flux + BA_faible.methanol_pur
                    Set BA_faible = Nothing
                End If
            End If
            If choix_BA_prolongee = True Then
                Call BA_prolongee.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_apportee_methanol, DCO_conso_bio)
                P_precipite = BA_prolongee.P_precipite
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = BA_prolongee.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = BA_prolongee.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = BA_prolongee.reel_MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
                If choix_primaire = False Then
                    For I = 1 To nb_etape_file_boues
                        TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_prolongee_EB
                    Next I
                Else
                    For I = 1 To nb_etape_file_boues
                        TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_prolongee_ED
                    Next I
                End If
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call BA_prolongee.calcul_consommation_electrique(electricite_II_aeration)
                    electricite_consommee = electricite_consommee + BA_prolongee.electricite
                    electricite_fixe(electricite_postes_repere_secondaire) = electricite_fixe(electricite_postes_repere_secondaire) + BA_prolongee.electricite_fixe
                    electricite_variable(electricite_postes_repere_secondaire) = electricite_variable(electricite_postes_repere_secondaire) + BA_prolongee.electricite - BA_prolongee.electricite_fixe
                    FeCl3_flux = FeCl3_flux + BA_prolongee.FeCl3_pur
                    FeCl3_eau_pur_kgj = FeCl3_eau_pur_kgj + BA_prolongee.FeCl3_pur
                    methanol_flux = methanol_flux + BA_prolongee.methanol_pur
                    Set BA_prolongee = Nothing
                End If
            End If
            If choix_HybAS = True Then
                Call traitement_hybas.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_apportee_methanol, DCO_conso_bio)
                P_precipite = traitement_hybas.deltaP_precipitation
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = traitement_hybas.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = traitement_hybas.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = traitement_hybas.reel_MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_HybAS
                Next I
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call traitement_hybas.calcul_consommation_electrique(electricite_II_aeration)
                    electricite_consommee = electricite_consommee + traitement_hybas.electricite
                    electricite_fixe(electricite_postes_repere_secondaire) = electricite_fixe(electricite_postes_repere_secondaire) + traitement_hybas.electricite_fixe
                    electricite_variable(electricite_postes_repere_secondaire) = electricite_variable(electricite_postes_repere_secondaire) + traitement_hybas.electricite - traitement_hybas.electricite_fixe
                    FeCl3_flux = FeCl3_flux + traitement_hybas.anaerobie_FeCl3_flux
                    FeCl3_eau_pur_kgj = FeCl3_eau_pur_kgj + traitement_hybas.anaerobie_FeCl3_flux
                    methanol_flux = methanol_flux + traitement_hybas.post_denit_methanol_flux
                    Set traitement_hybas = Nothing
                End If
            End If
            If choix_MBBR = True Then
                Call traitement_MBBR.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_apportee_methanol, DCO_conso_bio)
                P_precipite = 0
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = traitement_MBBR.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = traitement_MBBR.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = traitement_MBBR.boues_MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_origine) = II_MBBR
                Next I
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call traitement_MBBR.calcul_consommation_electrique(electricite_II_aeration)
                    electricite_consommee = electricite_consommee + traitement_MBBR.electricite
                    electricite_fixe(electricite_postes_repere_secondaire) = electricite_fixe(electricite_postes_repere_secondaire) + traitement_MBBR.electricite_fixe
                    electricite_variable(electricite_postes_repere_secondaire) = electricite_variable(electricite_postes_repere_secondaire) + traitement_MBBR.electricite - traitement_MBBR.electricite_fixe
                    'FeCl3_flux = FeCl3_flux + traitement_MBBR.anaerobie_FeCl3_flux
                    'FeCl3_eau_pur_kgj = FeCl3_eau_pur_kgj + traitement_MBBR.anaerobie_FeCl3_flux
                    methanol_flux = methanol_flux + traitement_MBBR.postdenit_carbone_apporte_flux
                    Set traitement_MBBR = Nothing
                End If
            End If
            If choix_biostyr = True Then
                'stockage des flux arrivant sur le(s) biostyr(s)
                stockage_Q = Q
                stockage_DCO = DCO
                stockage_DBO = DBO
                stockage_MES = MES
                stockage_NK = NK
                stockage_NH4 = NH4
                stockage_NO3 = NO3
                stockage_Pt = Pt
                stockage_DCO_conso_bio = DCO_conso_bio
                For iteration_ES = 1 To nb_iteration_ES
                    'on réutilise les valeurs des flux arrivants sur le biostyr stockés dans les variables
                    Q = stockage_Q
                    DCO = stockage_DCO
                    DBO = stockage_DBO
                    MES = stockage_MES
                    NK = stockage_NK
                    NH4 = stockage_NH4
                    NO3 = stockage_NO3
                    Pt = stockage_Pt
                    DCO_conso_bio = stockage_DCO_conso_bio
                    'eaux sales retournées en tête du biostyr après traitement spécifique
                    Q_ES_traitement_separe = 0
                    For I = 1 To nb_eaux_sales_max
                        If choix_eaux_sales(I) = devenir_ES_traitement_separe_amont Then
                            'flag pour la digestion
                            choix_eaux_sales_2aire_amont = True
                            'gestion des flux
                            Q = Q + eaux_sales(I, repere_ES_Q)
                            Q_ES_traitement_separe = Q_ES_traitement_separe + eaux_sales(I, repere_ES_Q)
                            DCO = DCO + eaux_sales(I, repere_ES_DCO)
                            DBO = DBO + eaux_sales(I, repere_ES_DBO)
                            MES = MES + eaux_sales(I, repere_ES_MES)
                            NK = NK + eaux_sales(I, repere_ES_NK)
                            NH4 = NH4 + eaux_sales(I, repere_ES_NH4)
                            NO3 = NO3 + eaux_sales(I, repere_ES_NO3)
                            Pt = Pt + eaux_sales(I, repere_ES_Pt)
                        ElseIf choix_eaux_sales(I) = devenir_ES_traitement_separe_aval Then
                            Q_ES_traitement_separe = Q_ES_traitement_separe + eaux_sales(I, repere_ES_Q)
                        End If
                    Next I
                    Call biofiltre.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_conso_bio, eaux_sales, choix_eaux_sales, iteration_ES, boues_ratio_DBO_MV)
                    'calcul des NK et du Pt présents dans les eaux sales
                    eaux_sales(repere_ES_biostyr, repere_ES_DBO) = eaux_sales(repere_ES_biostyr, repere_ES_DBO) + boues_ratio_DBO_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    eaux_sales(repere_ES_biostyr, repere_ES_NK) = eaux_sales(repere_ES_biostyr, repere_ES_NK) + boues_ratio_NK_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    eaux_sales(repere_ES_biostyr, repere_ES_Pt) = eaux_sales(repere_ES_biostyr, repere_ES_Pt) + boues_ratio_Pt_MES(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    'calcul des charges en NK et Pt
                    NK = NK + boues_ratio_NK_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) * MES
                    Pt = Pt - boues_ratio_Pt_MES(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    'correction pour le phosphore avec la partie soluble (on considère que ce qui reste est soluble et donc répartie de manière proportionnelle --> concentration égale)
                    eaux_sales(repere_ES_biostyr, repere_ES_Pt) = eaux_sales(repere_ES_biostyr, repere_ES_Pt) + Pt * eaux_sales(repere_ES_biostyr, repere_ES_Q) / (Q + eaux_sales(repere_ES_biostyr, repere_ES_Q))
                    Pt = Pt * Q / (Q + eaux_sales(repere_ES_biostyr, repere_ES_Q))
                    If choix_biostyr_PDN = True Then
                        Call biofiltre_PDN.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, DCO_conso_bio, eaux_sales, choix_eaux_sales, iteration_ES, DCO_apportee_methanol)
                    End If
                    If choix_decanteur_ES = True Then
                        Call decanteur_ES.fonctionnement_reel(eaux_sales, choix_eaux_sales, boues_II_detail, C_elimine, DBO_elimine, N_elimine, P_elimine)
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES) = decanteur_ES.boues_MES
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_Q) = decanteur_ES.boues_Q
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES) = decanteur_ES.MV_MES
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_verif_flux) = 1
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_flux_in) = 1
                        'GESTION DES BOUES ISSUES DES EAUX SALES (compositions gérées de manière similaire aux boues primaires)
                        For I = 1 To boues_epaissies(nb_max_epaississeur)
                            TableauRecapitulatifFluxBoues(I, boues_II, repere_MV_MES) = TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES)
                            TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_DCO_MES) = C_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                            TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_DBO_MES) = DBO_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                            TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_NK_MES) = N_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                            TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_Pt_MES) = P_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                        Next I
                        If iteration = pi_NOMBRE_ITERATION_RETOURS And iteration_ES = nb_iteration_ES Then
                            Call decanteur_ES.calcul_consommation_electrique
                            electricite_consommee = electricite_consommee + decanteur_ES.electricite
                            electricite_fixe(electricite_postes_repere_eaux_sales) = electricite_fixe(electricite_postes_repere_eaux_sales) + decanteur_ES.electricite_fixe
                            electricite_variable(electricite_postes_repere_eaux_sales) = electricite_variable(electricite_postes_repere_eaux_sales) + decanteur_ES.electricite - decanteur_ES.electricite_fixe
                            polymere_flux = polymere_flux + decanteur_ES.dosage_polymere * Q_ES_traitement_separe / 1000
                            polymere_eau_pur_kgj(decanteur_ES.choix_polymere_type) = polymere_eau_pur_kgj(decanteur_ES.choix_polymere_type) + decanteur_ES.dosage_polymere * Q_ES_traitement_separe / 1000
                            FeCl3_flux = FeCl3_flux + decanteur_ES.dosage_FeCl3 * Q_ES_traitement_separe / 1000
                            FeCl3_eau_pur_kgj = FeCl3_eau_pur_kgj + decanteur_ES.dosage_FeCl3 * Q_ES_traitement_separe / 1000
                            Set decanteur_ES = Nothing
                        End If
                    End If
                Next iteration_ES
                If choix_eaux_sales_2aire_amont = True Then
                    nb_iteration_ES = pi_NOMBRE_ITERATION_RETOURS
                End If
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call biofiltre.calcul_consommation_electrique(electricite_II_aeration)
                    electricite_consommee = electricite_consommee + biofiltre.electricite
                    electricite_fixe(electricite_postes_repere_secondaire) = electricite_fixe(electricite_postes_repere_secondaire)
                    electricite_variable(electricite_postes_repere_secondaire) = electricite_variable(electricite_postes_repere_secondaire) + biofiltre.electricite
                    Set biofiltre = Nothing
                    If choix_biostyr_PDN = True Then
                        Call biofiltre_PDN.calcul_consommation_electrique
                        electricite_consommee = electricite_consommee + biofiltre_PDN.electricite
                        electricite_fixe(electricite_postes_repere_secondaire) = electricite_fixe(electricite_postes_repere_secondaire)
                        electricite_variable(electricite_postes_repere_secondaire) = electricite_variable(electricite_postes_repere_secondaire) + biofiltre_PDN.electricite
                        methanol_flux = methanol_flux + biofiltre_PDN.methanol_flux
                        Set biofiltre_PDN = Nothing
                    End If
                End If
                'eaux sales renvoyées en aval du biostyr après traitement spécifique
                '                Q_ES_traitement_separe = 0
                For I = 1 To nb_eaux_sales_max
                    If choix_eaux_sales(I) = devenir_ES_traitement_separe_aval Then
                        Q = Q + eaux_sales(I, repere_ES_Q)
                        '                        Q_ES_traitement_separe = Q_ES_traitement_separe + eaux_sales(i, repere_ES_Q)
                        DCO = DCO + eaux_sales(I, repere_ES_DCO)
                        DBO = DBO + eaux_sales(I, repere_ES_DBO)
                        MES = MES + eaux_sales(I, repere_ES_MES)
                        NK = NK + eaux_sales(I, repere_ES_NK)
                        NH4 = NH4 + eaux_sales(I, repere_ES_NH4)
                        NO3 = NO3 + eaux_sales(I, repere_ES_NO3)
                        Pt = Pt + eaux_sales(I, repere_ES_Pt)
                    End If
                Next I
            End If
            If choix_biostyr = False Then  'si on n'a pas de biostyr
                'calcul des ratios de pollution dans les boues
                C_elimine = C_elimine + DCO_apportee_methanol - DCO
                DCO_conso_bio = C_elimine
                'N_elimine = N_elimine - NK
                P_elimine = P_elimine - Pt - P_precipite
                For I = 1 To boues_epaissies(nb_max_epaississeur)
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_MV_MES) = TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES)
                    '*************** MODIFIE LE 26/06/2012 PAR DCA ******************************
                    '*************** suite à des écarts au niveau de la DCO dans les boues, on fixe le ratio DCO/MES et on recalcule DCO_conso_bio **********
                    'TableauRecapitulatifFluxBoues(i, boues_II, repere_ratio_DCO_MES) = (C_elimine - DCO_conso_bio) / TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_DCO_MES) = boues_ratio_DCO_MV(TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_origine)) * TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES)
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_DBO_MES) = boues_ratio_DBO_MV(TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_origine)) * TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES)
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_NK_MES) = boues_ratio_NK_MV(TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_origine)) * TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MV_MES)
                    N_elimine = TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_ratio_NK_MES) * TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                    TableauRecapitulatifFluxBoues(I, boues_II, repere_ratio_Pt_MES) = P_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                Next I
                DCO_conso_bio = DCO_conso_bio - TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_ratio_DCO_MES) * TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_MES)
                If choix_BA_faible = True Or choix_BA_prolongee = True Then
                    'faible charge ou aération prolongée: NK = NH4 + NK particulaire
                    NK = NK + TableauRecapitulatifFluxBoues(boues_inlet, boues_II, repere_ratio_NK_MES) * MES
                ElseIf choix_BA_forte = True Or choix_BA_moyenne = True Then
                    'on soustrait la partie qui est assimilée et qui part dans les boues
                    NK = NK - N_elimine
                End If
            End If
        End If

        'TRAITEMENT TERTIAIRE
        If choix_tertiaire = True Then
            C_elimine = DCO
            DBO_elimine = DBO
            N_elimine = NK
            P_elimine = Pt
            P_precipite = 0
            If choix_biostyr_N_III = True Then
                'stockage des flux arrivant sur le(s) biostyr(s)
                stockage_Q = Q
                stockage_DCO = DCO
                stockage_DBO = DBO
                stockage_MES = MES
                stockage_NK = NK
                stockage_NH4 = NH4
                stockage_NO3 = NO3
                stockage_Pt = Pt
                stockage_DCO_conso_bio = DCO_conso_bio
                For iteration_ES = 1 To nb_iteration_ES
                    'on réutilise les valeurs des flux arrivants sur le biostyr stockés dans les variables
                    Q = stockage_Q
                    DCO = stockage_DCO
                    DBO = stockage_DBO
                    MES = stockage_MES
                    NK = stockage_NK
                    NH4 = stockage_NH4
                    NO3 = stockage_NO3
                    Pt = stockage_Pt
                    DCO_conso_bio = stockage_DCO_conso_bio
                    'eaux sales retournées en tête du biostyr après traitement spécifique
                    Q_ES_traitement_separe = 0
                    For I = 1 To nb_eaux_sales_max
                        If choix_eaux_sales(I) = devenir_ES_traitement_separe_amont Then
                            'flag pour la digestion
                            choix_eaux_sales_3aire_amont = True
                            'gestion des flux
                            Q = Q + eaux_sales(I, repere_ES_Q)
                            Q_ES_traitement_separe = Q_ES_traitement_separe + eaux_sales(I, repere_ES_Q)
                            DCO = DCO + eaux_sales(I, repere_ES_DCO)
                            DBO = DBO + eaux_sales(I, repere_ES_DBO)
                            MES = MES + eaux_sales(I, repere_ES_MES)
                            NK = NK + eaux_sales(I, repere_ES_NK)
                            NH4 = NH4 + eaux_sales(I, repere_ES_NH4)
                            NO3 = NO3 + eaux_sales(I, repere_ES_NO3)
                            Pt = Pt + eaux_sales(I, repere_ES_Pt)
                        ElseIf choix_eaux_sales(I) = devenir_ES_traitement_separe_aval Then
                            Q_ES_traitement_separe = Q_ES_traitement_separe + eaux_sales(I, repere_ES_Q)
                        End If
                    Next I
                    Call biofiltre_III.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, Sh, DCO_conso_bio, eaux_sales, choix_eaux_sales, iteration_ES)
                    'calcul des NK et du Pt présents dans les eaux sales
                    eaux_sales(repere_ES_biostyr, repere_ES_DBO) = eaux_sales(repere_ES_biostyr, repere_ES_DBO) + boues_ratio_DBO_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    eaux_sales(repere_ES_biostyr, repere_ES_NK) = eaux_sales(repere_ES_biostyr, repere_ES_NK) + boues_ratio_NK_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    eaux_sales(repere_ES_biostyr, repere_ES_Pt) = eaux_sales(repere_ES_biostyr, repere_ES_Pt) + boues_ratio_Pt_MES(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    'calcul des charges en NK et Pt
                    NK = NK + boues_ratio_NK_MV(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MV_MES) * MES
                    Pt = Pt - boues_ratio_Pt_MES(eaux_sales(repere_ES_biostyr, repere_ES_origine)) * eaux_sales(repere_ES_biostyr, repere_ES_MES)
                    'correction pour le phosphore avec la partie soluble (on considère que ce qui reste est soluble et donc répartie de manière proportionnelle --> concentration égale)
                    eaux_sales(repere_ES_biostyr, repere_ES_Pt) = eaux_sales(repere_ES_biostyr, repere_ES_Pt) + Pt * eaux_sales(repere_ES_biostyr, repere_ES_Q) / (Q + eaux_sales(repere_ES_biostyr, repere_ES_Q))
                    Pt = Pt * Q / (Q + eaux_sales(repere_ES_biostyr, repere_ES_Q))
                    If choix_biostyr_PDN_III = True Then
                        Call biofiltre_PDN_III.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, DCO_conso_bio, eaux_sales, choix_eaux_sales, iteration_ES, DCO_apportee_methanol)
                    End If
                    If choix_decanteur_ES = True Then
                        Call decanteur_ES_III.fonctionnement_reel(eaux_sales, choix_eaux_sales, boues_III_detail, C_elimine, DBO_elimine, N_elimine, P_elimine)
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES) = decanteur_ES_III.boues_MES
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_Q) = decanteur_ES_III.boues_Q
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MV_MES) = decanteur_ES_III.MV_MES
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_verif_flux) = 1
                        TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_flux_in) = 1
                        For I = 1 To boues_epaissies(nb_max_epaississeur)
                            TableauRecapitulatifFluxBoues(I, boues_III, repere_MV_MES) = TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MV_MES)
                            TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_DCO_MES) = C_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                            TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_DBO_MES) = DBO_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                            TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_NK_MES) = N_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                            TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_Pt_MES) = P_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                        Next I
                        If iteration = pi_NOMBRE_ITERATION_RETOURS And iteration_ES = nb_iteration_ES Then
                            Call decanteur_ES_III.calcul_consommation_electrique
                            electricite_consommee = electricite_consommee + decanteur_ES_III.electricite
                            electricite_fixe(electricite_postes_repere_eaux_sales) = electricite_fixe(electricite_postes_repere_eaux_sales) + decanteur_ES_III.electricite_fixe
                            electricite_variable(electricite_postes_repere_eaux_sales) = electricite_variable(electricite_postes_repere_eaux_sales) + decanteur_ES_III.electricite - decanteur_ES_III.electricite_fixe
                            polymere_flux = polymere_flux + decanteur_ES_III.dosage_polymere * Q_ES_traitement_separe / 1000
                            polymere_eau_pur_kgj(decanteur_ES_III.choix_polymere_type) = polymere_eau_pur_kgj(decanteur_ES_III.choix_polymere_type) + decanteur_ES_III.dosage_polymere * Q_ES_traitement_separe / 1000
                            FeCl3_flux = FeCl3_flux + decanteur_ES_III.dosage_FeCl3 * Q_ES_traitement_separe / 1000
                            FeCl3_eau_pur_kgj = FeCl3_eau_pur_kgj + decanteur_ES_III.dosage_FeCl3 * Q_ES_traitement_separe / 1000
                            Set decanteur_ES_III = Nothing
                        End If
                    End If
                Next iteration_ES
                If choix_eaux_sales_3aire_amont = True Then
                    nb_iteration_ES = pi_NOMBRE_ITERATION_RETOURS
                End If
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call biofiltre_III.calcul_consommation_electrique(electricite_III_aeration)
                    electricite_consommee = electricite_consommee + biofiltre_III.electricite
                    electricite_fixe(electricite_postes_repere_tertiaire) = electricite_fixe(electricite_postes_repere_tertiaire)
                    electricite_variable(electricite_postes_repere_tertiaire) = electricite_variable(electricite_postes_repere_tertiaire) + biofiltre_III.electricite
                    Set biofiltre_III = Nothing
                    If choix_biostyr_PDN_III = True Then
                        Call biofiltre_PDN_III.calcul_consommation_electrique
                        electricite_consommee = electricite_consommee + biofiltre_PDN_III.electricite
                        electricite_fixe(electricite_postes_repere_tertiaire) = electricite_fixe(electricite_postes_repere_tertiaire)
                        electricite_variable(electricite_postes_repere_tertiaire) = electricite_variable(electricite_postes_repere_tertiaire) + biofiltre_PDN_III.electricite
                        methanol_flux = methanol_flux + biofiltre_PDN_III.methanol_flux
                        Set biofiltre_PDN_III = Nothing
                    End If
                End If
                'eaux sales renvoyées en aval du biostyr après traitement spécifique
                '                Q_ES_traitement_separe = 0
                For I = 1 To nb_eaux_sales_max
                    If choix_eaux_sales(I) = devenir_ES_traitement_separe_aval Then
                        Q = Q + eaux_sales(I, repere_ES_Q)
                        '                        Q_ES_traitement_separe = Q_ES_traitement_separe + eaux_sales(i, repere_ES_Q)
                        DCO = DCO + eaux_sales(I, repere_ES_DCO)
                        DBO = DBO + eaux_sales(I, repere_ES_DBO)
                        MES = MES + eaux_sales(I, repere_ES_MES)
                        NK = NK + eaux_sales(I, repere_ES_NK)
                        NH4 = NH4 + eaux_sales(I, repere_ES_NH4)
                        NO3 = NO3 + eaux_sales(I, repere_ES_NO3)
                        Pt = Pt + eaux_sales(I, repere_ES_Pt)
                    End If
                Next I
            End If
            If choix_decanteur_III = True Then
                Call decanteur_III.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt)
                P_precipite = decanteur_III.P_precipite
                For I = 1 To nb_etape_file_boues
                    TableauRecapitulatifFluxBoues(I, boues_III, repere_origine) = III_decantation
                Next I
                TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES) = decanteur_III.boues_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_Q) = decanteur_III.boues_Q
                TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MV_MES) = decanteur_III.MV_MES
                TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_verif_flux) = 1
                TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_flux_in) = 1
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call decanteur_III.calcul_consommation_electrique
                    electricite_consommee = electricite_consommee + decanteur_III.electricite
                    electricite_fixe(electricite_postes_repere_tertiaire) = electricite_fixe(electricite_postes_repere_tertiaire) + decanteur_III.electricite_fixe
                    electricite_variable(electricite_postes_repere_tertiaire) = electricite_variable(electricite_postes_repere_tertiaire) + decanteur_III.electricite - decanteur_III.electricite_fixe
                    polymere_flux = polymere_flux + decanteur_III.dosage_polymere * Q / 1000
                    polymere_eau_pur_kgj(decanteur_III.choix_polymere_type) = polymere_eau_pur_kgj(decanteur_III.choix_polymere_type) + decanteur_III.dosage_polymere * Q / 1000
                    FeCl3_flux = FeCl3_flux + decanteur_III.dosage_FeCl3 * Q / 1000
                    FeCl3_eau_pur_kgj = FeCl3_eau_pur_kgj + decanteur_III.dosage_FeCl3 * Q / 1000
                    Set decanteur_III = Nothing
                End If
                'calcul des ratios de pollution dans les boues qui sont constantes jusqu'à la digestion
                C_elimine = C_elimine - DCO
                N_elimine = N_elimine - NK
                P_elimine = P_elimine - Pt - P_precipite
                For I = 1 To boues_epaissies(nb_max_epaississeur)
                    TableauRecapitulatifFluxBoues(I, boues_III, repere_MV_MES) = TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MV_MES)
                    TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_DCO_MES) = C_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                    TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_DBO_MES) = DBO_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                    TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_NK_MES) = N_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                    TableauRecapitulatifFluxBoues(I, boues_III, repere_ratio_Pt_MES) = P_elimine / TableauRecapitulatifFluxBoues(boues_inlet, boues_III, repere_MES)
                Next I
            End If
            'Discfilter
            If choix_discfilter = True Then
                Call Discfilter.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, eaux_sales_discfilter)
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call Discfilter.calcul_consommation_electrique
                    polymere_flux = polymere_flux + Discfilter.polymere_flux
                    polymere_eau_pur_kgj(Discfilter.choix_polymere_type) = polymere_eau_pur_kgj(Discfilter.choix_polymere_type) + Discfilter.polymere_flux
                    FeCl3_flux = FeCl3_flux + Discfilter.FeCl3_flux
                    FeCl3_eau_pur_kgj = FeCl3_eau_pur_kgj + Discfilter.FeCl3_flux
                    electricite_consommee = electricite_consommee + Discfilter.electricite
                    electricite_fixe(electricite_postes_repere_tertiaire) = electricite_fixe(electricite_postes_repere_tertiaire)
                    electricite_variable(electricite_postes_repere_tertiaire) = electricite_variable(electricite_postes_repere_tertiaire) + Discfilter.electricite
                    Set Discfilter = Nothing
                End If
            End If
            'Filtration tertiaire sur sable
            If choix_filtrasable = True Then
                Call Filtre_sable.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, eaux_sales_filtrasable)
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call Filtre_sable.calcul_consommation_electrique
                    electricite_consommee = electricite_consommee + Filtre_sable.electricite
                    electricite_fixe(electricite_postes_repere_tertiaire) = electricite_fixe(electricite_postes_repere_tertiaire)
                    electricite_variable(electricite_postes_repere_tertiaire) = electricite_variable(electricite_postes_repere_tertiaire) + Filtre_sable.electricite
                    Set Filtre_sable = Nothing
                End If
            End If
        End If

        'Désinfection
        If choix_desinfection = True Then
            'UV
            If choix_UV = True Then
                Call Traitement_UV.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt)
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call Traitement_UV.calcul_consommation_electrique
                    electricite_consommee = electricite_consommee + Traitement_UV.electricite
                    electricite_fixe(electricite_postes_repere_desinfection) = electricite_fixe(electricite_postes_repere_desinfection)
                    electricite_variable(electricite_postes_repere_desinfection) = electricite_variable(electricite_postes_repere_desinfection) + Traitement_UV.electricite
                    Set Traitement_UV = Nothing
                End If
            End If
            'Chloration
            If choix_chloration = True Then
                Call chloration.fonctionnement_reel(Q, DCO, DBO, MES, NK, NH4, NO3, Pt)
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Cl2_gazeux_flux_kgj = Cl2_gazeux_flux_kgj + chloration.Cl2_gazeux_pur_kgj
                    NaOClpur_flux = NaOClpur_flux + chloration.Cl2_HClO_pur_kgj
                    NaOClpur_eau_flux = NaOClpur_eau_flux + chloration.Cl2_HClO_pur_kgj
                    Call chloration.calcul_consommation_electrique
                    'electricite_desinfection = electricite_desinfection + chloration.electricite
                    'electricite_consommee = electricite_consommee + chloration.electricite
                    'electricite_fixe(electricite_postes_repere_desinfection) = electricite_fixe(electricite_postes_repere_desinfection)
                    'electricite_variable(electricite_postes_repere_desinfection) = electricite_variable(electricite_postes_repere_desinfection) + chloration.electricite
                    Set chloration = Nothing
                End If
            End If
        End If

        'On récupère des valeurs pour le calcul des KPI
        GestionCalculKPI.pd_EauTraiteeFluxDBO_kgj = DBO
        GestionCalculKPI.pd_EauTraiteeFluxDCO_kgj = DCO
        GestionCalculKPI.pd_EauTraiteeFluxMES_kgj = MES
        GestionCalculKPI.pd_EauTraiteeFluxN_kgj = NK
        GestionCalculKPI.pd_EauTraiteeFluxP_kgj = Pt

        ''''''''''''
        'FILE BOUES'
        ''''''''''''

        'RECAPITULATIF DES BOUES EN ENTREE
        Call recapitulatif_boues(TableauRecapitulatifFluxBoues, boues_pollution_soluble, graisses_particulaire_ratio_interne, boues_ratio_NK_MV, boues_ratio_Pt_MES, boues_ratio_DCO_MV, boues_ratio_DBO_MV, graisses_internes, nb_max_epaississeur, boues_Shunt, choix_eaux_sales_1aire, boues_I_detail)

        For I = 1 To retour_caracteristique_nb
            retour_flux(I) = 0
            retour_digestion(I) = 0
            retour_athos(I) = 0
            retour_flux_soluble(I) = 0
            retour_digestion_soluble(I) = 0
            retour_athos_soluble(I) = 0
        Next I
        
        
        'BIOLIX
        If choix_biolix = True Then
            Call biolix.fonctionnement_reel(TableauRecapitulatifFluxBoues, boues_pollution_soluble)
            If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                Call biolix.calcul_consommation_electrique
                electricite_consommee = electricite_consommee + biolix.pd_ConsommationElectrique_kWhj
                electricite_fixe(electricite_postes_repere_biolix) = electricite_fixe(electricite_postes_repere_biolix) + biolix.pd_ConsommationElectriqueFixe_kWhj
                electricite_variable(electricite_postes_repere_biolix) = electricite_variable(electricite_postes_repere_biolix) + biolix.pd_ConsommationElectrique_kWhj - biolix.pd_ConsommationElectriqueFixe_kWhj
                Set biolix = Nothing
            End If
        End If


        'EPAISSISSEMENT
        If choix_epaississement = True Then
            For I = 1 To epaississement_choix_nb
                Call epaississeur(I).fonctionnement_reel(TableauRecapitulatifFluxBoues, boues_pollution_soluble, I, retour_flux, retour_flux_soluble, boues_I_detail, boues_II_detail, boues_III_detail, epaississement_before_thelys)
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call epaississeur(I).calcul_consommation_electrique
                    electricite_consommee = electricite_consommee + epaississeur(I).electricite
                    electricite_fixe(electricite_postes_repere_epaississement) = electricite_fixe(electricite_postes_repere_epaississement) + epaississeur(I).electricite_fixe
                    electricite_variable(electricite_postes_repere_epaississement) = electricite_variable(electricite_postes_repere_epaississement) + epaississeur(I).electricite - epaississeur(I).electricite_fixe
                    polymere_flux = polymere_flux + epaississeur(I).polymere_flux
                    polymere_boues_pur_kgj(epaississeur(I).choix_polymere_type) = polymere_boues_pur_kgj(epaississeur(I).choix_polymere_type) + epaississeur(I).polymere_flux
                    Call epaississeur(I).ecriture_resultats(I)
                    Set epaississeur(I) = Nothing
                End If
            Next I
        End If

        'DIGESTION
        If pi_ChoixDigestion = True Then
            GestionCalculKPI.pd_BouesADigerer_kgMVj = 0
            GestionCalculKPI.pd_BiogazProduit_Nm3j = 0
            If choix_dig_simple = True Then
                Call digesteur_simple.fonctionnement_reel(TableauRecapitulatifFluxBoues, choix_eaux_sales_1aire, boues_I_detail, boues_II_detail, boues_III_detail, graisses_particulaire_ratio_interne, DCO_biogaz, boues_pollution_soluble, besoins_thermiques, biogaz)
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call digesteur_simple.calcul_consommation_electrique(type_valorisation_biogaz, valorisation_chaudiere, ratio_biogaz_valorise)
                    Call digesteur_simple.transfert_flux_energie_vers_pinch
                    GestionCalculKPI.pd_BouesADigerer_kgMVj = GestionCalculKPI.pd_BouesADigerer_kgMVj + digesteur_simple.inlet_MV_MES * digesteur_simple.inlet_MES
                    GestionCalculKPI.pd_BiogazProduit_Nm3j = GestionCalculKPI.pd_BiogazProduit_Nm3j + digesteur_simple.biogaz_Q_Nm3j
                    'electricite_verte = digesteur_simple.energie_elec
                    electricite_consommee = electricite_consommee + digesteur_simple.electricite
                    electricite_fixe(electricite_postes_repere_digestion) = electricite_fixe(electricite_postes_repere_digestion) + digesteur_simple.electricite_fixe
                    electricite_variable(electricite_postes_repere_digestion) = electricite_variable(electricite_postes_repere_digestion) + digesteur_simple.electricite - digesteur_simple.electricite_fixe
                    Set digesteur_simple = Nothing
                End If
            End If
            If choix_biothelys = True Then
                Call digesteur_biothelys.fonctionnement_reel(TableauRecapitulatifFluxBoues, graisses_particulaire_ratio_interne, choix_eaux_sales_1aire, boues_I_detail, boues_II_detail, boues_III_detail, DCO_biogaz, boues_pollution_soluble, besoins_thermiques, disponibilites_thermiques, biogaz, pression_vapeur, nb_vapeur, epaississement_before_thelys)
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call digesteur_biothelys.calcul_consommation_electrique(type_valorisation_biogaz, valorisation_chaudiere, ratio_biogaz_valorise)
                    Call digesteur_biothelys.transfert_flux_energie_vers_pinch
                    GestionCalculKPI.pd_BouesADigerer_kgMVj = GestionCalculKPI.pd_BouesADigerer_kgMVj + digesteur_biothelys.pd_MatieresVolatilesBouesEntrees_kgMVj
                    GestionCalculKPI.pd_BiogazProduit_Nm3j = GestionCalculKPI.pd_BiogazProduit_Nm3j + digesteur_biothelys.pd_BiogazProduit_Nm3j
                    'electricite_verte = digesteur_biothelys.energie_elec
                    electricite_consommee = electricite_consommee + digesteur_biothelys.pd_ConsommationElectrique_kWhj
                    electricite_fixe(electricite_postes_repere_digestion) = electricite_fixe(electricite_postes_repere_digestion) + digesteur_biothelys.pd_ConsommationElectriqueFixe_kWhj
                    electricite_variable(electricite_postes_repere_digestion) = electricite_variable(electricite_postes_repere_digestion) + digesteur_biothelys.pd_ConsommationElectrique_kWhj - digesteur_biothelys.pd_ConsommationElectriqueFixe_kWhj
                    Set digesteur_biothelys = Nothing
                End If
            End If
            If choix_exelys_DLD = True Then
                Call digesteur_DLD.fonctionnement_reel(TableauRecapitulatifFluxBoues, graisses_particulaire_ratio_interne, choix_eaux_sales_1aire, boues_I_detail, boues_II_detail, boues_III_detail, DCO_biogaz, boues_pollution_soluble, retour_flux, retour_digestion, retour_flux_soluble, retour_digestion_soluble, besoins_thermiques, disponibilites_thermiques, biogaz, pression_vapeur, nb_vapeur, epaississement_before_thelys)
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call digesteur_DLD.calcul_consommation_electrique(type_valorisation_biogaz, valorisation_chaudiere, ratio_biogaz_valorise)
                    Call digesteur_DLD.transfert_flux_energie_vers_pinch
                    GestionCalculKPI.pd_BouesADigerer_kgMVj = GestionCalculKPI.pd_BouesADigerer_kgMVj + digesteur_DLD.pd_MatieresVolatilesBouesEntrees_kgMVj
                    GestionCalculKPI.pd_BiogazProduit_Nm3j = GestionCalculKPI.pd_BiogazProduit_Nm3j + digesteur_DLD.pd_BiogazProduit_Nm3j
                    'electricite_verte = digesteur_DLD.energie_elec
                    electricite_consommee = electricite_consommee + digesteur_DLD.pd_ConsommationElectrique_kWhj
                    electricite_fixe(electricite_postes_repere_digestion) = electricite_fixe(electricite_postes_repere_digestion) + digesteur_DLD.pd_ConsommationElectriqueFixe_kWhj
                    electricite_variable(electricite_postes_repere_digestion) = electricite_variable(electricite_postes_repere_digestion) + digesteur_DLD.pd_ConsommationElectrique_kWhj - digesteur_DLD.pd_ConsommationElectriqueFixe_kWhj
                    polymere_flux = polymere_flux + digesteur_DLD.pd_CentrifugeuseConsommationPolymerePur_kgj
                    polymere_boues_pur_kgj(digesteur_DLD.pi_ChoixTypePolymereCentrifugeuse) = polymere_boues_pur_kgj(digesteur_DLD.pi_ChoixTypePolymereCentrifugeuse) + digesteur_DLD.pd_CentrifugeuseConsommationPolymerePur_kgj
                    Set digesteur_DLD = Nothing
                End If
            End If
        End If

        'Athos
        If choix_athos = True Then
            Call OVH_athos.fonctionnement_reel(TableauRecapitulatifFluxBoues, retour_flux, retour_digestion, retour_athos, retour_flux_soluble, retour_digestion_soluble, retour_athos_soluble, boues_pollution_soluble, DCO_oxydation)
            besoins_thermiques(energie_biogaz) = besoins_thermiques(energie_biogaz) + OVH_athos.pd_TraitementFumeesRTOConsommationBiogaz_kWhPCIj
            besoins_thermiques(energie_combustible) = besoins_thermiques(energie_combustible) + OVH_athos.pd_TraitementFumeesRTOConsommationCombustibleExterne_kWhPCIj
            If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                Call OVH_athos.calcul_consommation_electrique
                Call OVH_athos.transfert_flux_energie_vers_pinch
                electricite_consommee = electricite_consommee + OVH_athos.pd_ConsommationElectrique_kWhj
                electricite_fixe(electricite_postes_repere_athos) = electricite_fixe(electricite_postes_repere_athos)
                electricite_variable(electricite_postes_repere_athos) = electricite_variable(electricite_postes_repere_athos) + OVH_athos.pd_ConsommationElectrique_kWhj
                O2_liquide_flux = O2_liquide_flux + OVH_athos.pd_ConsommationOxygeneLiquidePur_kgj
                CuSO4_flux = CuSO4_flux + OVH_athos.pd_ConsommationCuSO4Pur_kgj
                polymere_flux = polymere_flux + OVH_athos.pd_EpaississementConsommationPolymerePur_kgj
                polymere_boues_pur_kgj(OVH_athos.pi_ChoixTypePolymereEpaississement) = polymere_boues_pur_kgj(OVH_athos.pi_ChoixTypePolymereEpaississement) + OVH_athos.pd_EpaississementConsommationPolymerePur_kgj
                Set OVH_athos = Nothing
            End If
        End If

        'DESHYDRATATION
        If choix_deshydratation = True Then
            GestionCalculKPI.pd_ConsommationPolymereDeshydratation_kgMA_j = 0
            GestionCalculKPI.pd_BouesDeshydrateesMS_kgMSj = 0
            GestionCalculKPI.pd_BouesHumidesDeshydratees_kgj = 0
            For I = 1 To deshydratation_choix_nb
                Call deshydrat(I).fonctionnement_reel(TableauRecapitulatifFluxBoues, I, retour_flux, retour_digestion, retour_flux_soluble, retour_digestion_soluble, boues_pollution_soluble, boues_I_detail, boues_II_detail, boues_III_detail, retour_athos)
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call deshydrat(I).calcul_consommation_electrique
                    electricite_consommee = electricite_consommee + deshydrat(I).electricite
                    electricite_fixe(electricite_postes_repere_deshydratation) = electricite_fixe(electricite_postes_repere_deshydratation)
                    electricite_variable(electricite_postes_repere_deshydratation) = electricite_variable(electricite_postes_repere_deshydratation) + deshydrat(I).electricite
                    If deshydrat(I).inlet_technosable = False Then
                        GestionCalculKPI.pd_BouesDeshydrateesMS_kgMSj = GestionCalculKPI.pd_BouesDeshydrateesMS_kgMSj + deshydrat(I).outlet_MS
                        GestionCalculKPI.pd_BouesHumidesDeshydratees_kgj = GestionCalculKPI.pd_BouesHumidesDeshydratees_kgj + (deshydrat(I).outlet_Q * deshydrat(I).boues_masse_volumique_kgm3)
                        GestionCalculKPI.pd_ConsommationPolymereDeshydratation_kgMA_j = GestionCalculKPI.pd_ConsommationPolymereDeshydratation_kgMA_j + deshydrat(I).polymere_flux
                    End If
                    polymere_flux = polymere_flux + deshydrat(I).polymere_flux
                    polymere_boues_pur_kgj(deshydrat(I).choix_polymere_type) = polymere_boues_pur_kgj(deshydrat(I).choix_polymere_type) + deshydrat(I).polymere_flux
                    chaux_eteinte_flux = chaux_eteinte_flux + deshydrat(I).chaux_eteinte_flux
                    chaux_vive_flux = chaux_vive_flux + deshydrat(I).chaux_vive_flux
                    FeCl3_flux = FeCl3_flux + deshydrat(I).FeCl3_flux
                    FeCl3_boues_pur_kgj = FeCl3_boues_pur_kgj + deshydrat(I).FeCl3_flux
                    Call deshydrat(I).ecriture_resultats(I)
                    Set deshydrat(I) = Nothing
                End If
            Next I
        End If

        'CHAULAGE
        If choix_chaulage = True Then
            For I = 1 To chaulage_choix_nb
                Call chaulage(I).fonctionnement_reel(TableauRecapitulatifFluxBoues, I, boues_I_detail, boues_II_detail, boues_III_detail, retour_flux, boues_pollution_soluble)
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call chaulage(I).calcul_consommation_electrique
                    electricite_consommee = electricite_consommee + chaulage(I).electricite
                    electricite_fixe(electricite_postes_repere_chaulage) = electricite_fixe(electricite_postes_repere_chaulage)
                    electricite_variable(electricite_postes_repere_chaulage) = electricite_variable(electricite_postes_repere_chaulage) + chaulage(I).electricite
                    chaux_eteinte_flux = chaux_eteinte_flux + chaulage(I).chaux_eteinte_flux
                    chaux_vive_flux = chaux_vive_flux + chaulage(I).chaux_vive_flux
                    Call chaulage(I).ecriture_resultats(I)
                    Set chaulage(I) = Nothing
                End If
            Next I
        End If

        'SECHAGE DES BOUES
        If choix_sechage = True Then
            GestionCalculKPI.pd_BouesSecheesQuantiteMS_kgMSj = 0
            If choix_secheur_inos = True Then
                For I = 1 To sechage_inos_choix_nb
                    Call secheur_inos(I).fonctionnement_reel(TableauRecapitulatifFluxBoues, I, boues_I_detail, boues_II_detail, boues_III_detail, retour_flux, boues_pollution_soluble, besoins_thermiques, disponibilites_thermiques)
                    If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                        Call secheur_inos(I).calcul_consommation_electrique
                        Call secheur_inos(I).transfert_flux_energie_vers_pinch(I)
                        'Récupération de données pour les KPI
                        GestionCalculKPI.pd_BouesSecheesQuantiteMS_kgMSj = GestionCalculKPI.pd_BouesSecheesQuantiteMS_kgMSj + secheur_inos(I).inlet_MES
                        electricite_consommee = electricite_consommee + secheur_inos(I).electricite
                        electricite_fixe(electricite_postes_repere_sechage) = electricite_fixe(electricite_postes_repere_sechage)
                        electricite_variable(electricite_postes_repere_sechage) = electricite_variable(electricite_postes_repere_sechage) + secheur_inos(I).electricite
                        polymere_flux = polymere_flux + secheur_inos(I).polymere_flux
                        polymere_boues_pur_kgj(secheur_inos(I).choix_polymere_type) = polymere_boues_pur_kgj(secheur_inos(I).choix_polymere_type) + secheur_inos(I).polymere_flux
                        FeCl3_flux = FeCl3_flux + secheur_inos(I).FeCl3_flux
                        FeCl3_boues_pur_kgj = FeCl3_boues_pur_kgj + secheur_inos(I).FeCl3_flux
                        chaux_eteinte_flux = chaux_eteinte_flux + secheur_inos(I).chaux_eteinte_flux
                        Call secheur_inos(I).ecriture_resultats(I)
                        Set secheur_inos(I) = Nothing
                    End If
                Next I
            End If
            If choix_secheur_bioco = True Then
                For I = 1 To sechage_bioco_choix_nb
                    Call secheur_bioco(I).fonctionnement_reel(TableauRecapitulatifFluxBoues, I, retour_flux, boues_pollution_soluble, besoins_thermiques, disponibilites_thermiques)
                    If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                        Call secheur_bioco(I).calcul_consommation_electrique
                        Call secheur_bioco(I).transfert_flux_energie_vers_pinch(I)
                        GestionCalculKPI.pd_BouesSecheesQuantiteMS_kgMSj = GestionCalculKPI.pd_BouesSecheesQuantiteMS_kgMSj + secheur_bioco(I).inlet_MES
                        electricite_consommee = electricite_consommee + secheur_bioco(I).electricite
                        electricite_fixe(electricite_postes_repere_sechage) = electricite_fixe(electricite_postes_repere_sechage)
                        electricite_variable(electricite_postes_repere_sechage) = electricite_variable(electricite_postes_repere_sechage) + secheur_bioco(I).electricite
                        Call secheur_bioco(I).ecriture_resultats(I)
                        Set secheur_bioco(I) = Nothing
                    End If
                Next I
            End If
            If choix_secheur_autre = True Then
                For I = 1 To sechage_choix_nb
                    Call secheur_autre(I).fonctionnement_reel(TableauRecapitulatifFluxBoues, I, retour_flux, boues_pollution_soluble, besoins_thermiques, disponibilites_thermiques)
                    If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                        Call secheur_autre(I).calcul_consommation_electrique
                        Call secheur_autre(I).transfert_flux_energie_vers_pinch(I)
                        GestionCalculKPI.pd_BouesSecheesQuantiteMS_kgMSj = GestionCalculKPI.pd_BouesSecheesQuantiteMS_kgMSj + secheur_autre(I).inlet_MES
                        electricite_consommee = electricite_consommee + secheur_autre(I).electricite
                        electricite_fixe(electricite_postes_repere_sechage) = electricite_fixe(electricite_postes_repere_sechage)
                        electricite_variable(electricite_postes_repere_sechage) = electricite_variable(electricite_postes_repere_sechage) + secheur_autre(I).electricite
                        Call secheur_autre(I).ecriture_resultats(I)
                        Set secheur_autre(I) = Nothing
                    End If
                Next I
            End If
        End If

        'Incinération
        If choix_incineration = True Then
            Call incineration.fonctionnement_reel(TableauRecapitulatifFluxBoues, boues_pollution_soluble, besoins_thermiques, disponibilites_thermiques, biogaz(teneur_CH4), DCO_oxydation)
            If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                Call incineration.calcul_consommation_electrique
                Call incineration.transfert_flux_energie_vers_pinch
                GestionCalculKPI.pd_BouesIncinereesNonSecheesQuantiteMS_kgMSj = incineration.pd_EntreeBouesNonSechees_kgMSj
                electricite_consommee = electricite_consommee + incineration.electricite
                electricite_fixe(electricite_postes_repere_incineration) = electricite_fixe(electricite_postes_repere_incineration)
                electricite_variable(electricite_postes_repere_incineration) = electricite_variable(electricite_postes_repere_incineration) + incineration.electricite
                ConsommationAmmoniaquePur_kgj = ConsommationAmmoniaquePur_kgj + incineration.md_TraitementSNCRConsommationNH3Pur_kgj
                ConsommationUreePur_kgj = ConsommationUreePur_kgj + incineration.md_TraitementSNCRConsommationUreePure_kgj
                NaOHpur_flux = NaOHpur_flux + incineration.NaOHpur_kgj
                NaHCO3_flux = NaHCO3_flux + incineration.NaHCO3pur_kgj
                charbon_actif_flux = charbon_actif_flux + incineration.charbon_actif_kgj
                cendres_evacuees_Tj = incineration.cendres_kgj / 1000
                REFIB_evacues_Tj = incineration.REFIB_kgj / 1000
                cendres_evacuees_tMSj = cendres_evacuees_Tj * incineration.siccite_cendres
                REFIB_evacues_tMSj = REFIB_evacues_Tj * incineration.siccite_REFIB
                Call incineration.ecriture_resultats
                Set incineration = Nothing
            End If
        End If

        'TRAITEMENT SEPARE DES RETOURS DE TETE
        If choix_traitement_retours = True Then
            DCO_traitement_retours = retour_flux(repere_ret_DCO)
            If choix_MAP_retours = True Then
                Call traitement_MAP_retours.fonctionnement_reel(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble)
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call traitement_MAP_retours.calcul_consommation_electrique
                    electricite_consommee = electricite_consommee + traitement_MAP_retours.pd_electricite_kWhj
                    NaOHpur_flux = NaOHpur_flux + traitement_MAP_retours.pd_NaOHpur_kgj
                    MgCl2_flux = MgCl2_flux + traitement_MAP_retours.pd_MgCl2pur_kgj
                    Struvite_kgj = traitement_MAP_retours.pd_StruviteProduite_kgj
                    electricite_fixe(electricite_postes_repere_trait_retours) = electricite_fixe(electricite_postes_repere_trait_retours)
                    electricite_variable(electricite_postes_repere_trait_retours) = electricite_variable(electricite_postes_repere_trait_retours) + traitement_MAP_retours.pd_electricite_kWhj
                    Set traitement_MAP_retours = Nothing
                End If
            End If
            If choix_ANITA_Mox = True Then
                Call traitement_Mox.fonctionnement_reel(retour_flux, retour_flux_soluble, retour_digestion, retour_digestion_soluble, retour_athos, retour_athos_soluble)
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call traitement_Mox.calcul_consommation_electrique
                    electricite_consommee = electricite_consommee + traitement_Mox.electricite
                    electricite_fixe(electricite_postes_repere_trait_retours) = electricite_fixe(electricite_postes_repere_trait_retours)
                    electricite_variable(electricite_postes_repere_trait_retours) = electricite_variable(electricite_postes_repere_trait_retours) + traitement_Mox.electricite
                    Set traitement_Mox = Nothing
                End If
            End If
            If choix_ANITA_Shunt = True Then
                Call traitement_Shunt.fonctionnement_reel(retour_flux, retour_digestion, retour_flux_soluble, retour_digestion_soluble, retour_athos, retour_athos_soluble, boues_Shunt, DCO_methanol_Shunt)
                If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                    Call traitement_Shunt.calcul_consommation_electrique
                    methanol_flux = methanol_flux + traitement_Shunt.methanol_flux
                    electricite_consommee = electricite_consommee + traitement_Shunt.electricite
                    electricite_fixe(electricite_postes_repere_trait_retours) = electricite_fixe(electricite_postes_repere_trait_retours)
                    electricite_variable(electricite_postes_repere_trait_retours) = electricite_variable(electricite_postes_repere_trait_retours) + traitement_Shunt.electricite
                    Set traitement_Shunt = Nothing
                End If
            End If
            DCO_traitement_retours = DCO_traitement_retours + DCO_methanol_Shunt - retour_flux(repere_ret_DCO) - boues_Shunt(repere_graisse_DCO_MES) * boues_Shunt(repere_graisse_MS)
        End If


        'RETOURS EN TETE
        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
            If iteration = 1 Then
                Dim retours As New Retours_gestion
            End If
            Call retours.forcer_flux(retour_flux)
            If iteration = pi_NOMBRE_ITERATION_RETOURS Then
                Set retours = Nothing
            End If
        End If

        'début modif MSM 21/11/12
        'Chauffage & Clim
        If iteration = pi_NOMBRE_ITERATION_RETOURS Then
            'début modif MSM 03/12/12 : désodo dernière itération pour obtenir débit d'air
            'DESODORISATION
            'désodorisation chimique
            GestionCalculKPI.pd_ConsommationElectriqueTraitementOdeurs_kWhj = 0
            GestionCalculKPI.pd_DebitAirDesodorise_Nm3h = 0
            If choix_desodo_chimique = True Then
                Dim desodo_chim As New Desodorisation_chimique
                'on calcule le process
                Call desodo_chim.calcul(HS_strippe_kgj)
                If pn_EffacementResultatsEnCours = True Then
                    GoTo calcul_interrompu
                End If
                Call desodo_chim.calcul_consommation_electrique
                'on récupère les consos élecs
                GestionCalculKPI.pd_ConsommationElectriqueTraitementOdeurs_kWhj = GestionCalculKPI.pd_ConsommationElectriqueTraitementOdeurs_kWhj + desodo_chim.electricite
                GestionCalculKPI.pd_DebitAirDesodorise_Nm3h = GestionCalculKPI.pd_DebitAirDesodorise_Nm3h + desodo_chim.air_vicie_Q_Nm3h
                electricite_consommee = electricite_consommee + desodo_chim.electricite
                electricite_fixe(electricite_postes_repere_desodorisation) = electricite_fixe(electricite_postes_repere_desodorisation) + desodo_chim.electricite
                electricite_variable(electricite_postes_repere_desodorisation) = electricite_variable(electricite_postes_repere_desodorisation)
                H2SO4pur_flux = H2SO4pur_flux + desodo_chim.acide_H2SO4_pur
                NaOHpur_flux = NaOHpur_flux + desodo_chim.oxydobase_pH9_NaOH_pur + desodo_chim.oxydobase_pH11_NaOH_pur
                NaOClpur_desodo_flux = NaOClpur_desodo_flux + desodo_chim.oxydobase_Cl2_pur
                NaOClpur_flux = NaOClpur_flux + desodo_chim.oxydobase_Cl2_pur
                NaHSO3pur_flux = NaHSO3pur_flux + desodo_chim.neutre_NaHSO3_pur
                'on récupère les réactifs
                Set desodo_chim = Nothing
            End If
            'désodorisation biologique
            If choix_desodo_bio = True Then
                Dim desodo_bio As New Desodorisation_biologique
                'on calcule le process
                Call desodo_bio.calcul
                Call desodo_bio.calcul_consommation_electrique
                'on récupère les consos élecs
                GestionCalculKPI.pd_ConsommationElectriqueTraitementOdeurs_kWhj = GestionCalculKPI.pd_ConsommationElectriqueTraitementOdeurs_kWhj + desodo_bio.electricite
                GestionCalculKPI.pd_DebitAirDesodorise_Nm3h = GestionCalculKPI.pd_DebitAirDesodorise_Nm3h + desodo_bio.air_vicie_Q_Nm3h
                electricite_consommee = electricite_consommee + desodo_bio.electricite
                electricite_fixe(electricite_postes_repere_desodorisation) = electricite_fixe(electricite_postes_repere_desodorisation) + desodo_bio.electricite
                electricite_variable(electricite_postes_repere_desodorisation) = electricite_variable(electricite_postes_repere_desodorisation)
                'on récupère les réactifs
                Set desodo_bio = Nothing
            End If
            'fin modif MSM 03/12/12

            If choix_utilites_chauffage_clim = True Then

                Dim divers_utilites_HVAC As New Utilites
                electricite_utilites = 0

                'Bâtiments administratifs
                If choix_utilites_bat_administration = True Then
                    Dim HVAC_admin As New HVAC_bat_admin
                    Call HVAC_admin.calcul
                    Call HVAC_admin.transfert_flux_energie_vers_pinch
                    electricite_consommee = electricite_consommee + HVAC_admin.consoElecTotale
                    electricite_utilites = HVAC_admin.consoElecTotale
                    electricite_fixe(electricite_postes_repere_utilites) = electricite_fixe(electricite_postes_repere_utilites) + electricite_utilites
                    Call HVAC_admin.ecriture_resultats
                    Set HVAC_admin = Nothing
                    '            electricite_produite_solaire_photovoltaique = solaire_photovoltaique.electricite_produite_kwh_j
                    '            electricite_verte = electricite_verte + electricite_produite_solaire_photovoltaique
                End If

                'Bâtiments d'exploitation
                If choix_utilites_bat_exploitation = True Then
                    Dim HVAC_exploit As New HVAC_bat_exploit
                    Call HVAC_exploit.calcul
                    Call HVAC_exploit.transfert_flux_energie_vers_pinch
                    electricite_consommee = electricite_consommee + HVAC_exploit.consoElecTotale
                    electricite_utilites = HVAC_exploit.consoElecTotale
                    electricite_fixe(electricite_postes_repere_utilites) = electricite_fixe(electricite_postes_repere_utilites) + electricite_utilites
                    Call HVAC_exploit.ecriture_resultats
                    Set HVAC_exploit = Nothing
                End If

                'Bâtiments elec
                If choix_utilites_bat_electrique = True Then
                    Dim HVAC_elec As New HVAC_bat_elec
                    Call HVAC_elec.calcul
                    electricite_consommee = electricite_consommee + HVAC_elec.consoElecTotale
                    electricite_utilites = HVAC_elec.consoElecTotale
                    electricite_fixe(electricite_postes_repere_utilites) = electricite_fixe(electricite_postes_repere_utilites) + electricite_utilites
                    Call HVAC_elec.ecriture_resultats
                    Set HVAC_elec = Nothing
                End If

                Set divers_utilites_HVAC = Nothing

            End If
        End If
        'fin modif MSM 21/11/12

        'GESTION DE L'ENERGIE THERMIQUE
        If iteration = pi_NOMBRE_ITERATION_RETOURS Then
            If choix_analyse_pinch = True Then
                energie_thermique_PINCH.besoins_thermiques_biogaz_incineration = besoins_thermiques(energie_biogaz)
                energie_thermique_PINCH.besoins_thermiques_combustible_incineration = besoins_thermiques(energie_combustible)
                Call energie_thermique_PINCH.fonctionnement_reel(biogaz, electricite_verte, type_valorisation_biogaz, valorisation_chaudiere, ratio_biogaz_valorise)
                GestionCalculKPI.pd_ConsommationEnergieThermiqueExterneSechage_kWhj = energie_thermique_PINCH.pd_BesoinsCombustibleExterneHorsIncineration_kWhj
                GestionCalculKPI.pd_ConsommationEnergieThermiqueExterneIncineration_kWhj = energie_thermique_PINCH.pd_BesoinsCombustibleExterneIncineration_kWhj
                gaz_naturel_kWhPCIj = energie_thermique_PINCH.gaz_naturel_kWhPCIj
                fioul_kWhPCIj = energie_thermique_PINCH.fioul_kWhPCIj
                electricite_verte_biogaz = energie_thermique_PINCH.electricite_biogaz
                Set energie_thermique_PINCH = Nothing
            Else
                Call energie_thermique.fonctionnement_reel(besoins_thermiques, pression_vapeur, biogaz, disponibilites_thermiques, electricite_verte, type_valorisation_biogaz, valorisation_chaudiere, ratio_biogaz_valorise)
                'On récupère les consommations de combustible externe pour séchage et incinération
                'Celle du séchage vaut 0 car la valeur pd_BesoinsCombustibleExterneIncineration_kWhj contient les 2 informations
                    GestionCalculKPI.pd_ConsommationEnergieThermiqueExterneIncineration_kWhj = energie_thermique.pd_BesoinsCombustibleExterneIncineration_kWhj
                    GestionCalculKPI.pd_ConsommationEnergieThermiqueExterneSechage_kWhj = energie_thermique.pd_BesoinsCombustibleExterneIncineration_kWhj
                gaz_naturel_kWhPCIj = energie_thermique.gaz_naturel_kWhPCIj
                fioul_kWhPCIj = energie_thermique.fioul_kWhPCIj
                electricite_verte_biogaz = energie_thermique.electricite_biogaz
                Set energie_thermique = Nothing
            End If
            If pn_EffacementResultatsEnCours = True Then
                GoTo calcul_interrompu
            End If
        End If


    Next iteration

    'bypass_reel:


    'début modif MSM 03/12/12 : désodo --> dernière itération pour obtenir débit d'air
    '    'DESODORISATION
    '    'désodorisation chimique
    '    If choix_desodo_chimique = True Then
    '        Dim desodo_chim As New Desodorisation_chimique
    '        'on calcule le process
    '        Call desodo_chim.calcul(HS_strippe_kgj)
    '        Call desodo_chim.calcul_consommation_electrique
    '        'on récupère les consos élecs
    '        electricite_consommee = electricite_consommee + desodo_chim.electricite
    '        electricite_fixe(electricite_postes_repere_desodorisation) = electricite_fixe(electricite_postes_repere_desodorisation) + desodo_chim.electricite
    '        electricite_variable(electricite_postes_repere_desodorisation) = electricite_variable(electricite_postes_repere_desodorisation)
    '        H2SO4pur_flux = H2SO4pur_flux + desodo_chim.acide_H2SO4_pur
    '        NaOHpur_flux = NaOHpur_flux + desodo_chim.oxydobase_pH9_NaOH_pur + desodo_chim.oxydobase_pH11_NaOH_pur
    '        NaOClpur_desodo_flux = NaOClpur_desodo_flux + desodo_chim.oxydobase_Cl2_pur
    '        NaOClpur_flux = NaOClpur_flux + desodo_chim.oxydobase_Cl2_pur
    '        NaHSO3pur_flux = NaHSO3pur_flux + desodo_chim.neutre_NaHSO3_pur
    '        'on récupère les réactifs
    '        Set desodo_chim = Nothing
    '    End If
    '    'désodorisation biologique
    '    If choix_desodo_bio = True Then
    '        Dim desodo_bio As New Desodorisation_biologique
    '        'on calcule le process
    '        Call desodo_bio.calcul
    '        Call desodo_bio.calcul_consommation_electrique
    '        'on récupère les consos élecs
    '        electricite_consommee = electricite_consommee + desodo_bio.electricite
    '        electricite_fixe(electricite_postes_repere_desodorisation) = electricite_fixe(electricite_postes_repere_desodorisation) + desodo_bio.electricite
    '        electricite_variable(electricite_postes_repere_desodorisation) = electricite_variable(electricite_postes_repere_desodorisation)
    '        'on récupère les réactifs
    '        Set desodo_bio = Nothing
    '    End If
    'fin modif MSM 03/12/12 : désodo --> dernière itération pour obtenir débit d'air

    'on détruit le traitement de HS
    Set traitement_HS = Nothing

    'CONSOMMATIONS DES UTILITES
    Dim divers_utilites As New Utilites
    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
        Call divers_utilites.lecture_valeurs_forcees
    End If
    If choix_utilites_pompage_retours = True Then
        electricite_utilites = 0
        Call divers_utilites.calcul_retour(electricite_consommee, TableauRecapitulatifFluxBoues, retour_flux, electricite_utilites)
        Call divers_utilites.ecriture_resultats_retour
        electricite_fixe(electricite_postes_repere_utilites) = electricite_fixe(electricite_postes_repere_utilites)
        electricite_variable(electricite_postes_repere_utilites) = electricite_variable(electricite_postes_repere_utilites) + electricite_utilites
    End If
    If choix_utilites_eau_service = True Then
        electricite_utilites = 0
        Call divers_utilites.calcul_eau_service(electricite_consommee, TableauRecapitulatifFluxBoues, retour_flux, electricite_utilites)
        Call divers_utilites.ecriture_resultats_eau_service
        electricite_fixe(electricite_postes_repere_utilites) = electricite_fixe(electricite_postes_repere_utilites) + electricite_utilites
        electricite_variable(electricite_postes_repere_utilites) = electricite_variable(electricite_postes_repere_utilites)
    End If
    If choix_utilites_eclairage = True Then
        electricite_utilites = 0
        Call divers_utilites.calcul_eclairage(electricite_consommee, electricite_utilites)
        Call divers_utilites.ecriture_resultats_eclairage
        electricite_fixe(electricite_postes_repere_utilites) = electricite_fixe(electricite_postes_repere_utilites) + electricite_utilites
        electricite_variable(electricite_postes_repere_utilites) = electricite_variable(electricite_postes_repere_utilites)
    End If
    If choix_utilites_pertes_enligne = True Then
        electricite_utilites = 0
        Call divers_utilites.calcul_pertes(electricite_consommee, electricite_utilites)
        Call divers_utilites.ecriture_resultats_pertes
        electricite_fixe(electricite_postes_repere_utilites) = electricite_fixe(electricite_postes_repere_utilites) + electricite_utilites
        electricite_variable(electricite_postes_repere_utilites) = electricite_variable(electricite_postes_repere_utilites)
    End If
    Set divers_utilites = Nothing

    'ENERGIES ALTERNATIVES
    If choix_prod_alt_electricite = True Then

        'Photovoltaique
        If choix_solaire_photovoltaique = True Then
            Dim solaire_photovoltaique As New Photovoltaique
            Call solaire_photovoltaique.calcul
            Call solaire_photovoltaique.ecriture_resultats
            electricite_produite_solaire_photovoltaique = solaire_photovoltaique.electricite_produite_kwh_j
            electricite_verte = electricite_verte + electricite_produite_solaire_photovoltaique
        End If

        'Turbine hydraulique
        If choix_turbine_hydraulique = True Then
            Dim turbine_hydrau As New Turbine_hydraulique
            Call turbine_hydrau.calcul
            Call turbine_hydrau.ecriture_resultats
            electricite_produite_turbine_hydraulique = turbine_hydrau.electricite_produite_kwh_j
            electricite_verte = electricite_verte + electricite_produite_turbine_hydraulique
        End If

        'Autre production
        If choix_electricite_autre_production = True Then
            electricite_produite_autre = Feuil6.Range("electricite_produite_autre").Cells(1, pi_FiliereConsideree)
            electricite_verte = electricite_verte + electricite_produite_autre
        End If

    End If




    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
        Call expert(DCO, DCO_conso_bio, TableauRecapitulatifFluxBoues, retour_flux, DCO_biogaz, DCO_apportee_methanol, graisses_internes, boues_pollution_soluble, DCO_traitement_retours, DCO_methanol_Shunt, boues_Shunt, DCO_oxydation)
    End If

    Call ecrire_resultats_step(Q, DCO, DBO, MES, NK, NH4, NO3, Pt, TableauRecapitulatifFluxBoues, graisses_internes, retour_flux, boues_pollution_soluble, choix_eaux_sales, eaux_sales, eaux_sales_discfilter, eaux_sales_filtrasable, QuantiteRefusDegrillage_kgj)


    'CONSOMMATIONS ET PRODUCTIONS ELECTRIQUES
    Dim consoProdElec As New Bilan_electrique
    Call consoProdElec.ecrire_resultats(DCO, DBO, electricite_consommee, electricite_fixe, electricite_variable, electricite_II_aeration, electricite_III_aeration, electricite_verte, electricite_verte_biogaz, electricite_produite_solaire_photovoltaique, electricite_produite_turbine_hydraulique, electricite_produite_autre)
    'On récupère des données pour le calcul des KPI
    GestionCalculKPI.pd_ConsommationElectriqueTotale_kWhj = electricite_consommee
    GestionCalculKPI.pd_ConsommationElectriqueAeration_kWhj = electricite_II_aeration + electricite_III_aeration
    GestionCalculKPI.pd_ProductionElectriqueBiogaz_kWhj = electricite_verte_biogaz
    Set consoProdElec = Nothing

    'RECAPITULATIF REACTIFS
'    Dim bilan_reactifs As New Gestion_reactifs
    Call bilan_reactifs.calcul(polymere_eau_pur_kgj, polymere_boues_pur_kgj, chaux_eteinte_flux, chaux_vive_flux, FeCl3_eau_pur_kgj, FeCl3_boues_pur_kgj, methanol_flux, H2SO4pur_flux, NaOHpur_flux, NaOClpur_eau_flux, NaOClpur_desodo_flux, NaHSO3pur_flux, NaHCO3_flux, O2_liquide_flux, charbon_actif_flux, Cl2_gazeux_flux_kgj, Ca_2NO3_pur_flux, H2O2pur_flux, CuSO4_flux, MgCl2_flux, ConsommationAmmoniaquePur_kgj, ConsommationUreePur_kgj)
    'On récupère des données pour le calcul des KPI
    GestionCalculKPI.pd_ConsommationCoagulantFileEau_kgFe_j = FeCl3_eau_pur_kgj * pd_MASSE_MOLAIRE_Fe_kg_mol / (pd_MASSE_MOLAIRE_Fe_kg_mol + 3 * pd_MASSE_MOLAIRE_Cl_kg_mol)
    GestionCalculKPI.pd_ConsommationMethanolDenitrification_kgj = methanol_flux
    Set bilan_reactifs = Nothing

    'EMPREINTE CO2
    Call verification_boues(TableauRecapitulatifFluxBoues, warning, cendres_evacuees_tMSj, REFIB_evacues_tMSj, boues_evacuees_Q, graisses_evacuees_Q)

    Dim CO2 As New Empreinte_CO2
    C_elimine = DCO_nominal * pd_PourcentageChargeReelleDCO - DCO
    N_elimine = NK_nominal * pd_PourcentageChargeReelleNK - (NK + NO3)
    P_elimine = Pt_nominal * pd_PourcentageChargeReellePt - Pt
    Call CO2.calcul(electricite_consommee, electricite_verte, polymere_flux, chaux_eteinte_flux, chaux_vive_flux, FeCl3_flux, methanol_flux, H2SO4pur_flux, NaOHpur_flux, NaOClpur_flux, NaHSO3pur_flux, C_elimine, N_elimine, P_elimine, boues_evacuees_Q, graisses_evacuees_Q, gaz_naturel_kWhPCIj, fioul_kWhPCIj, NaHCO3_flux, Ca_2NO3_pur_flux, O2_liquide_flux, H2O2pur_flux, charbon_actif_flux, Cl2_gazeux_flux_kgj, cendres_evacuees_Tj, REFIB_evacues_Tj, Struvite_kgj, QuantiteRefusDegrillage_kgj, ConsommationAmmoniaquePur_kgj, ConsommationUreePur_kgj)
    Set CO2 = Nothing

    'CALCUL DES KPI
    Call GestionCalculKPI.calcul_et_ecriture_resultats
    Set GestionCalculKPI = Nothing


    'Exit Sub
calcul_interrompu:

    If pn_EffacementResultatsEnCours = False Then
        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
            Call afficher_com
        End If
        calcul_effectue(pi_FiliereConsideree, pi_EtapeCalculConsideree) = True
    Else
        'effacer les résultats déjà écrits
        Feuil6.Range("calcul_effectue").Cells(pi_FiliereConsideree, pi_EtapeCalculConsideree) = True
        flag_duplication(pi_FiliereConsideree) = True
        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
            Feuil6.Range("calcul_effectue").Cells(pi_FiliereConsideree, pi_CALCUL_VALEUR_GUIDE) = False
        End If
        interm_etape = pi_EtapeCalculConsideree
        Call Feuil2.suppression_resultats
        pn_EffacementResultatsEnCours = True   'la valeur est remise à False dans la subroutine appelée ci-dessus
        pi_EtapeCalculConsideree = interm_etape
        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
            Feuil6.Range("calcul_effectue").Cells(pi_FiliereConsideree, pi_CALCUL_VALEUR_GUIDE) = True
        End If
        'on dit que le calcul n'a pas été effectué
        calcul_effectue(pi_FiliereConsideree, pi_EtapeCalculConsideree) = False
        flag_duplication(pi_FiliereConsideree) = False
    End If

    Exit Sub

    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "0.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True



End Sub





Sub lecture_donnees_generales()


Dim message_erreur As String

    'choix généraux
    pi_ChoixQualiteRejet = Feuil6.Range("choix_qualite_rejet").Cells(1, pi_FiliereConsideree)

    'cahier des charges
    pd_CapaciteSTEP_EH = Feuil6.Range("Eq_hab").Cells(1, 6 + pi_FiliereConsideree)
    Q_nominal = Feuil6.Range("Q_nominal").Cells(1, 6 + pi_FiliereConsideree)
    If Q_nominal / pd_CapaciteSTEP_EH < 0.175 Then
        type_eau_nominal = eau_concentree
    ElseIf Q_nominal / pd_CapaciteSTEP_EH < 0.325 Then
        type_eau_nominal = eau_standard
    Else
        type_eau_nominal = eau_diluee
    End If
    DCO_nominal = Feuil6.Range("DCO_nominal").Cells(1, 6 + pi_FiliereConsideree)
    DBO_nominal = Feuil6.Range("DBO_nominal").Cells(1, 6 + pi_FiliereConsideree)
        If F_verification_DBO_DCO(DBO_nominal, DCO_nominal) = False Then
            message_erreur = Feuil6.Range("erreur_eau_entree_nominal").Cells(1, 1) & " " & Feuil6.Range("erreur_DCO_DBO").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub 'met fin aux calculs
        End If
    MES_nominal = Feuil6.Range("MES_nominal").Cells(1, 6 + pi_FiliereConsideree)
    NK_nominal = Feuil6.Range("NK_nominal").Cells(1, 6 + pi_FiliereConsideree)
    NH4_nominal = Feuil6.Range("NH4_nominal").Cells(1, 6 + pi_FiliereConsideree)
        
        If F_verification_NK_NH4(NK_nominal, NH4_nominal) = False Then
            message_erreur = Feuil6.Range("erreur_eau_entree_nominal").Cells(1, 1) & " " & Feuil6.Range("erreur_NH4_NK").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub 'met fin aux calculs
        End If
       
    Pt_nominal = Feuil6.Range("Pt_nominal").Cells(1, 6 + pi_FiliereConsideree)
    SH_nominal = Feuil6.Range("HS_nominal_mgL").Cells(1, 6 + pi_FiliereConsideree) * Q_nominal / 1000
    
    vidange_Q_nominal = Feuil6.Range("vidange_Q_nominal").Cells(1, 6 + pi_FiliereConsideree)
    vidange_DCO_nominal = Feuil6.Range("vidange_DCO_mgL_nominal").Cells(1, 6 + pi_FiliereConsideree) * vidange_Q_nominal / 1000
    vidange_DBO_nominal = Feuil6.Range("vidange_DBO_mgL_nominal").Cells(1, 6 + pi_FiliereConsideree) * vidange_Q_nominal / 1000
            
            If F_verification_DBO_DCO(vidange_DBO_nominal, vidange_DCO_nominal) = False Then
            message_erreur = Feuil6.Range("erreur_eau_entree_vidange").Cells(1, 1) & " " & Feuil6.Range("erreur_DCO_DBO").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub 'met fin aux calculs
            End If
    
    vidange_MES_nominal = Feuil6.Range("vidange_MES_mgL_nominal").Cells(1, 6 + pi_FiliereConsideree) * vidange_Q_nominal / 1000
    vidange_NK_nominal = Feuil6.Range("vidange_NK_mgL_nominal").Cells(1, 6 + pi_FiliereConsideree) * vidange_Q_nominal / 1000
    vidange_NH4_nominal = Feuil6.Range("vidange_NH4_mgL_nominal").Cells(1, 6 + pi_FiliereConsideree) * vidange_Q_nominal / 1000
           
           If F_verification_NK_NH4(vidange_NK_nominal, vidange_NH4_nominal) = False Then
            message_erreur = Feuil6.Range("erreur_eau_entree_vidange").Cells(1, 1) & " " & Feuil6.Range("erreur_NH4_NK").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub 'met fin aux calculs
           End If
        
    vidange_Pt_nominal = Feuil6.Range("vidange_Pt_mgL_nominal").Cells(1, 6 + pi_FiliereConsideree) * vidange_Q_nominal / 1000

    pd_TemperatureEauDimensionnement_°C = Feuil6.Range("T_eau_design").Cells(1, 6 + pi_FiliereConsideree)

    DCO_garantie = Feuil6.Range("DCO_garantie").Cells(1, 6 + pi_FiliereConsideree)
    DBO_garantie = Feuil6.Range("DBO_garantie").Cells(1, 6 + pi_FiliereConsideree)
    
        If F_verification_DBO_DCO(DBO_garantie, DCO_garantie) = False Then
            message_erreur = Feuil6.Range("erreur_eau_sortie").Cells(1, 1) & " " & Feuil6.Range("erreur_DCO_DBO").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub 'met fin aux calculs
        End If
        
    MES_garantie = Feuil6.Range("MES_garantie").Cells(1, 6 + pi_FiliereConsideree)
    If Feuil6.Range("NGL_garantie").Cells(1, 6 + pi_FiliereConsideree) <> "" Then
        NGL_garantie = Feuil6.Range("NGL_garantie").Cells(1, 6 + pi_FiliereConsideree)
    Else
        NGL_garantie = NK_nominal / Q_nominal * 1000
    End If
    If Feuil6.Range("NK_garantie").Cells(1, 6 + pi_FiliereConsideree) <> "" Then
        NK_garantie = Feuil6.Range("NK_garantie").Cells(1, 6 + pi_FiliereConsideree)
    Else
        NK_garantie = NK_nominal / Q_nominal * 1000
    End If
    If Feuil6.Range("Pt_garantie").Cells(1, 6 + pi_FiliereConsideree) <> "" Then
        Pt_garantie = Feuil6.Range("Pt_garantie").Cells(1, 6 + pi_FiliereConsideree)
    Else
        Pt_garantie = Pt_nominal / Q_nominal * 1000
    End If
    
               If F_verification_NK_NH4(NGL_garantie, NK_garantie) = False Then
            message_erreur = Feuil6.Range("erreur_eau_sortie").Cells(1, 1) & " " & Feuil6.Range("erreur_NK_NGL").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub 'met fin aux calculs
           End If

    'conditions d'exploitation
    pd_PourcentageChargeReelleDebitVolumique = Feuil6.Range("NC_Q").Cells(1, 6 + pi_FiliereConsideree)
    pd_PourcentageChargeReelleDCO = Feuil6.Range("NC_DCO").Cells(1, 6 + pi_FiliereConsideree)
    pd_PourcentageChargeReelleDBO = Feuil6.Range("NC_DBO").Cells(1, 6 + pi_FiliereConsideree)
    
        If F_verification_DBO_DCO(pd_PourcentageChargeReelleDBO * DBO_nominal, pd_PourcentageChargeReelleDCO * DCO_nominal) = False Then
                message_erreur = Feuil6.Range("erreu_eau_entree_pourcentage_reel").Cells(1, 1) & " " & Feuil6.Range("erreur_DCO_DBO").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub 'met fin aux calculs
        End If
    
    pd_PourcentageChargeReelleMES = Feuil6.Range("NC_MES").Cells(1, 6 + pi_FiliereConsideree)
    pd_PourcentageChargeReelleNK = Feuil6.Range("NC_NK").Cells(1, 6 + pi_FiliereConsideree)
    pd_PourcentageChargeReelleNH4 = Feuil6.Range("NC_NH4").Cells(1, 6 + pi_FiliereConsideree)
    
        If F_verification_NK_NH4(pd_PourcentageChargeReelleNK * NK_nominal, pd_PourcentageChargeReelleNH4 * NH4_nominal) = False Then
                message_erreur = Feuil6.Range("erreu_eau_entree_pourcentage_reel").Cells(1, 1) & " " & Feuil6.Range("erreur_NH4_NK").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub 'met fin aux calculs
        End If
    
    pd_PourcentageChargeReellePt = Feuil6.Range("NC_Pt").Cells(1, 6 + pi_FiliereConsideree)
    pd_PourcentageChargeReelleSH = (Feuil6.Range("HS_nominal_mgL").Cells(1, 6 + pi_FiliereConsideree) * Q_nominal * pd_PourcentageChargeReelleDebitVolumique / 1000) / SH_nominal

    vidange_Q_reel = Feuil6.Range("vidange_Q_reel").Cells(1, 6 + pi_FiliereConsideree)
    vidange_DCO_reel = Feuil6.Range("vidange_DCO_mgL_reel").Cells(1, 6 + pi_FiliereConsideree) * vidange_Q_reel / 1000
    vidange_DBO_reel = Feuil6.Range("vidange_DBO_mgL_reel").Cells(1, 6 + pi_FiliereConsideree) * vidange_Q_reel / 1000
    
            If F_verification_DBO_DCO(vidange_DBO_reel, vidange_DCO_reel) = False Then
                message_erreur = Feuil6.Range("erreur_eau_entree_vidange_reel").Cells(1, 1) & " " & Feuil6.Range("erreur_DCO_DBO").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub 'met fin aux calculs
            End If
    
    vidange_MES_reel = Feuil6.Range("vidange_MES_mgL_reel").Cells(1, 6 + pi_FiliereConsideree) * vidange_Q_reel / 1000
    vidange_NK_reel = Feuil6.Range("vidange_NK_mgL_reel").Cells(1, 6 + pi_FiliereConsideree) * vidange_Q_reel / 1000
    vidange_NH4_reel = Feuil6.Range("vidange_NH4_mgL_reel").Cells(1, 6 + pi_FiliereConsideree) * vidange_Q_reel / 1000
    
            If F_verification_NK_NH4(vidange_NK_reel, vidange_NH4_reel) = False Then
                message_erreur = Feuil6.Range("erreur_eau_entree_vidange_reel").Cells(1, 1) & " " & Feuil6.Range("erreur_NH4_NK").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub 'met fin aux calculs
            End If
    
    
    vidange_Pt_reel = Feuil6.Range("vidange_Pt_mgL_reel").Cells(1, 6 + pi_FiliereConsideree) * vidange_Q_reel / 1000

    pd_CoefficientPointeHydrauliqueTempsSec = Feuil6.Range("pointe_TS").Cells(1, 6 + pi_FiliereConsideree)
    pd_CoefficientPointeHydrauliqueTempsPluie = Feuil6.Range("pointe_TP").Cells(1, 6 + pi_FiliereConsideree)
        If pd_CoefficientPointeHydrauliqueTempsPluie < pd_CoefficientPointeHydrauliqueTempsSec Then
            message_erreur = Feuil6.Range("erreur_eau_pointe_hydraulique").Cells(1, 1) & " : " & Feuil6.Range("erreur_CPH").Cells(1, 1)
            MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
            pn_EffacementResultatsEnCours = True
            pn_ErreurNonGeree = True
            Exit Sub 'met fin aux calculs
        End If
    
    pd_TemperatureEauConditionsReelles_°C = Feuil6.Range("T_eau_exploit").Cells(1, 6 + pi_FiliereConsideree)
    T_air_aspire = Feuil6.Range("T_air_aspire").Cells(1, 6 + pi_FiliereConsideree)
    If T_air_aspire < 7 Then
        choix_climat = climat_froid
    ElseIf T_air_aspire < 20 Then
        choix_climat = climat_tempere
    Else
        choix_climat = climat_chaud
    End If
    humidite_air = Feuil6.Range("humidite_air").Cells(1, 6 + pi_FiliereConsideree)
        If F_verification_humidite_air(humidite_air) = False Then
                message_erreur = Feuil6.Range("erreur_humidite").Cells(1, 1) & " : " & Feuil6.Range("erreur_negatif_100").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub 'met fin aux calculs
        End If
      
    altitude = Feuil6.Range("altitude").Cells(1, 6 + pi_FiliereConsideree)
        If F_verification_altitude(altitude) = False Then
                message_erreur = Feuil6.Range("erreur_altitude").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub 'met fin aux calculs
        End If


    If pd_PourcentageChargeReelleDebitVolumique * Q_nominal / (pd_CapaciteSTEP_EH * pd_PourcentageChargeReelleDBO) < 0.175 Then
        type_eau_reel = eau_concentree
    ElseIf pd_PourcentageChargeReelleDebitVolumique * Q_nominal / (pd_CapaciteSTEP_EH * pd_PourcentageChargeReelleDBO) < 0.325 Then
        type_eau_reel = eau_standard
    Else
        type_eau_reel = eau_diluee
    End If




End Sub
Sub lecture_choix_process()

    Dim verification_configuration As Integer
    Dim message_erreur As String
    Dim verif_nb_biofiltre As Integer

    Dim choix_ES_biostyr As Integer
    Dim choix_ES_biostyr_PDN As Integer

    Dim count_traitement As Double

    verif_nb_biofiltre = 0
    count_traitement = 0

    'lecture des choix de traitement
    choix_traitement_sulfures_preventif = Feuil5.Range("choix_trait_sulfures_preventif").Cells(1, pi_FiliereConsideree)
    If choix_traitement_sulfures_preventif = True Then
        count_traitement = count_traitement + 1
    End If
    choix_traitement_sulfures_curatif = Feuil5.Range("choix_trait_sulfures").Cells(1, pi_FiliereConsideree)
    If choix_traitement_sulfures_curatif = True Then
        count_traitement = count_traitement + 1
    End If

    choix_relevement = Feuil5.Range("choix_relevement").Cells(1, pi_FiliereConsideree)
    If choix_relevement = True Then
        count_traitement = count_traitement + 1
    End If

    choix_pretraitement = Feuil5.Range("choix_pretraitement").Cells(1, pi_FiliereConsideree)
    If choix_pretraitement = True Then
        count_traitement = count_traitement + 1
        verification_configuration = 0
        choix_degrillage = Feuil5.Range("choix_degrillage").Cells(1, pi_FiliereConsideree)
        If choix_degrillage = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_dessablage = Feuil5.Range("choix_dessablage").Cells(1, pi_FiliereConsideree)
        If choix_dessablage = True Then
            verification_configuration = verification_configuration + 1
        End If
        If pn_EffacementResultatsEnCours = False Then
            If verification_configuration = 0 Then
                message_erreur = Feuil5.Range("erreur_choix_pretraitement").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub    'met fin aux calculs
                'ElseIf verification_configuration > 1 Then
                'treatment_message = Feuil2.Range("synthese_primaire") & ": "
                'Call erreur_configuration_process(treatment_message)
                'If choix_continuer = False Then
                'pn_EffacementResultatsEnCours = True
                'Exit Sub
                'End If
            End If
        End If
    Else
        choix_degrillage = False
        choix_dessablage = False
    End If

    choix_primaire = Feuil5.Range("choix_primaire").Cells(1, pi_FiliereConsideree)
    If choix_primaire = True Then
        verification_configuration = 0
        count_traitement = count_traitement + 1
        choix_decanteur_simple = Feuil5.Range("choix_decanteur_simple").Cells(1, pi_FiliereConsideree)
        If choix_decanteur_simple = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_decanteur_reactif = Feuil5.Range("choix_decanteur_reactif").Cells(1, pi_FiliereConsideree)
        If choix_decanteur_reactif = True Then
            verification_configuration = verification_configuration + 1
        End If
        If pn_EffacementResultatsEnCours = False Then
            If verification_configuration = 0 Or verification_configuration > 1 Then
                message_erreur = Feuil5.Range("erreur_choix_primaire").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub    'met fin aux calculs
                'ElseIf verification_configuration > 1 Then
                'treatment_message = Feuil2.Range("synthese_primaire") & ": "
                'Call erreur_configuration_process(treatment_message)
                'If choix_continuer = False Then
                'pn_EffacementResultatsEnCours = True
                'Exit Sub
                'End If
            End If
        End If
    Else
        choix_decanteur_simple = False
        choix_decanteur_reactif = False
    End If

    choix_secondaire = Feuil5.Range("choix_secondaire").Cells(1, pi_FiliereConsideree)
    If choix_secondaire = True Then
        verification_configuration = 0
        count_traitement = count_traitement + 1
        choix_BA_forte = Feuil5.Range("choix_BA_forte").Cells(1, pi_FiliereConsideree)
        If choix_BA_forte = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_BA_moyenne = Feuil5.Range("choix_BA_moyenne").Cells(1, pi_FiliereConsideree)
        If choix_BA_moyenne = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_BA_faible = Feuil5.Range("choix_BA_faible").Cells(1, pi_FiliereConsideree)
        If choix_BA_faible = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_BA_prolongee = Feuil5.Range("choix_BA_prolongee").Cells(1, pi_FiliereConsideree)
        If choix_BA_prolongee = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_HybAS = Feuil5.Range("choix_HybAS").Cells(1, pi_FiliereConsideree)
        If choix_HybAS = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_MBR = Feuil5.Range("choix_MBR").Cells(1, pi_FiliereConsideree)
        If choix_MBR = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_MBBR = Feuil5.Range("choix_MBBR").Cells(1, pi_FiliereConsideree)
        If choix_MBBR = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_biostyr = Feuil5.Range("choix_biostyr").Cells(1, pi_FiliereConsideree)
        If choix_biostyr = True Then
            verification_configuration = verification_configuration + 1
            verif_nb_biofiltre = verif_nb_biofiltre + 1
            choix_biostyr_PDN = Feuil5.Range("choix_biostyr_PDN").Cells(1, pi_FiliereConsideree)
            If choix_biostyr_PDN = True Then
                choix_ES_biostyr = Feuil14.Range("Biostyr_choix_devenir_ES").Cells(1, pi_FiliereConsideree)
                choix_ES_biostyr_PDN = Feuil18.Range("biostyr_PDN_choix_devenir_ES").Cells(1, pi_FiliereConsideree)
                If choix_ES_biostyr > 1 Then
                    If choix_ES_biostyr <> choix_ES_biostyr_PDN And choix_ES_biostyr_PDN > 1 Then
                        message_erreur = Feuil5.Range("erreur_choix_secondaire").Cells(1, 1)
                        MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                        pn_EffacementResultatsEnCours = True
                        pn_ErreurNonGeree = True
                        Exit Sub    'met fin aux calculs
                    End If
                End If
            End If
        Else
            choix_biostyr_PDN = Feuil5.Range("choix_biostyr_PDN").Cells(1, pi_FiliereConsideree)
            If choix_biostyr_PDN = True And pn_EffacementResultatsEnCours = False Then
                message_erreur = Feuil5.Range("erreur_choix_secondaire").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub    'met fin aux calculs
            End If
        End If
        If pn_EffacementResultatsEnCours = False Then
            If verification_configuration = 0 Or verification_configuration > 1 Then
                message_erreur = Feuil5.Range("erreur_choix_secondaire").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub    'met fin aux calculs
                'ElseIf verification_configuration > 1 Then
                'treatment_message = Feuil2.Range("synthese_secondaire") & ": "
                'Call erreur_configuration_process(treatment_message)
                'If choix_continuer = False Then
                'pn_EffacementResultatsEnCours = True
                'Exit Sub
                'End If
            End If
        End If
    Else
        choix_BA_forte = False
        choix_BA_moyenne = False
        choix_BA_faible = False
        choix_BA_prolongee = False
        choix_HybAS = False
        choix_MBBR = False
        choix_MBR = False
        choix_biostyr = False
        choix_biostyr_PDN = False
    End If

    choix_tertiaire = Feuil5.Range("choix_tertiaire").Cells(1, pi_FiliereConsideree)
    If choix_tertiaire = True Then
        verification_configuration = 0
        count_traitement = count_traitement + 1
        choix_biostyr_N_III = Feuil5.Range("choix_biostyr_N_III").Cells(1, pi_FiliereConsideree)
        If choix_biostyr_N_III = True Then
            verification_configuration = verification_configuration + 1
            verif_nb_biofiltre = verif_nb_biofiltre + 1
            choix_biostyr_PDN_III = Feuil5.Range("choix_biostyr_PDN_III").Cells(1, pi_FiliereConsideree)
            If choix_biostyr_PDN_III = True Then
                choix_ES_biostyr = Feuil21.Range("Biostyr_N3_choix_devenir_ES").Cells(1, pi_FiliereConsideree)
                choix_ES_biostyr_PDN = Feuil22.Range("biostyr_PDN3_choix_devenir_ES").Cells(1, pi_FiliereConsideree)
                If choix_ES_biostyr > 1 Then
                    If choix_ES_biostyr <> choix_ES_biostyr_PDN And choix_ES_biostyr_PDN > 1 Then
                        message_erreur = Feuil5.Range("erreur_choix_tertiaire").Cells(1, 1)
                        MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                        pn_EffacementResultatsEnCours = True
                        pn_ErreurNonGeree = True
                        Exit Sub    'met fin aux calculs
                    End If
                End If
            End If
        Else
            choix_biostyr_PDN_III = Feuil5.Range("choix_biostyr_PDN_III").Cells(1, pi_FiliereConsideree)
            If choix_biostyr_PDN_III = True And pn_EffacementResultatsEnCours = False Then
                message_erreur = Feuil5.Range("erreur_choix_tertiaire").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub    'met fin aux calculs
            End If
        End If
        choix_decanteur_III = Feuil5.Range("choix_decanteur_III").Cells(1, pi_FiliereConsideree)
        If choix_decanteur_III = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_discfilter = Feuil5.Range("choix_discfilter").Cells(1, pi_FiliereConsideree)
        If choix_discfilter = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_filtrasable = Feuil5.Range("choix_filtrasable").Cells(1, pi_FiliereConsideree)
        If choix_filtrasable = True Then
            verification_configuration = verification_configuration + 1
        End If
        If pn_EffacementResultatsEnCours = False Then
            If (verification_configuration = 0 Or verification_configuration > 1) Then
                message_erreur = Feuil5.Range("erreur_choix_tertiaire").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub    'met fin aux calculs
                'ElseIf verification_configuration > 1 Then
                'treatment_message = Feuil2.Range("synthese_secondaire") & ": "
                'Call erreur_configuration_process(treatment_message)
                'If choix_continuer = False Then
                'pn_EffacementResultatsEnCours = True
                'Exit Sub
                'End If
            End If
        End If
    Else
        choix_biostyr_N_III = False
        choix_biostyr_PDN_III = False
        choix_decanteur_III = False
        choix_discfilter = False
        choix_filtrasable = False
    End If

    choix_desinfection = Feuil5.Range("choix_desinfection").Cells(1, pi_FiliereConsideree)
    If choix_desinfection = True Then
        verification_configuration = 0
        count_traitement = count_traitement + 1
        choix_UV = Feuil5.Range("choix_UV").Cells(1, pi_FiliereConsideree)
        If choix_UV = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_chloration = Feuil5.Range("choix_chloration").Cells(1, pi_FiliereConsideree)
        If choix_chloration = True Then
            verification_configuration = verification_configuration + 1
        End If
        If pn_EffacementResultatsEnCours = False Then
            If (verification_configuration = 0 Or verification_configuration > 1) Then
                message_erreur = Feuil5.Range("erreur_choix_desinfection").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub    'met fin aux calculs
                'ElseIf verification_configuration > 1 Then
                'treatment_message = Feuil2.Range("synthese_secondaire") & ": "
                'Call erreur_configuration_process(treatment_message)
                'If choix_continuer = False Then
                'pn_EffacementResultatsEnCours = True
                'Exit Sub
                'End If
            End If
        End If
    Else
        choix_UV = False
        choix_chloration = False
    End If

    choix_decanteur_ES = Feuil5.Range("choix_decanteur_ES").Cells(1, pi_FiliereConsideree)

    choix_biolix = Feuil5.Range("choix_biolix").Cells(1, pi_FiliereConsideree)
    If choix_biolix = True Then
        count_traitement = count_traitement + 1
    End If
    
    choix_epaississement = Feuil5.Range("choix_epaississement").Cells(1, pi_FiliereConsideree)
    If choix_epaississement = True Then
        count_traitement = count_traitement + 1
    End If
    pi_ChoixDigestion = Feuil5.Range("choix_digestion").Cells(1, pi_FiliereConsideree)
    If pi_ChoixDigestion = True Then
        verification_configuration = 0
        count_traitement = count_traitement + 1
        choix_dig_simple = Feuil5.Range("choix_dig_simple").Cells(1, pi_FiliereConsideree)
        If choix_dig_simple = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_biothelys = Feuil5.Range("choix_biothelys").Cells(1, pi_FiliereConsideree)
        If choix_biothelys = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_exelys_DLD = Feuil5.Range("choix_exelys_DLD").Cells(1, pi_FiliereConsideree)
        If choix_exelys_DLD = True Then
            verification_configuration = verification_configuration + 1
        End If
        If pn_EffacementResultatsEnCours = False Then
            If verification_configuration = 0 Or verification_configuration > 1 Then
                message_erreur = Feuil5.Range("erreur_choix_digestion").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub    'met fin aux calculs
            End If
        End If
    Else
        choix_dig_simple = False
        choix_biothelys = False
        choix_exelys_DLD = False
    End If

    choix_athos = Feuil5.Range("choix_athos").Cells(1, pi_FiliereConsideree)
    If choix_athos = True Then
        count_traitement = count_traitement + 1
    End If

    choix_deshydratation = Feuil5.Range("choix_deshydratation").Cells(1, pi_FiliereConsideree)
    If choix_deshydratation = True Then
        count_traitement = count_traitement + 1
    End If

    choix_chaulage = Feuil5.Range("choix_chaulage").Cells(1, pi_FiliereConsideree)
    If choix_chaulage = True Then
        count_traitement = count_traitement + 1
    End If

    choix_sechage = Feuil5.Range("choix_sechage").Cells(1, pi_FiliereConsideree)
    If choix_sechage = True Then
        verification_configuration = 0
        count_traitement = count_traitement + 1
        choix_secheur_autre = Feuil5.Range("choix_secheur_autre").Cells(1, pi_FiliereConsideree)
        If choix_secheur_autre = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_secheur_inos = Feuil5.Range("choix_secheur_inos").Cells(1, pi_FiliereConsideree)
        If choix_secheur_inos = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_secheur_bioco = Feuil5.Range("choix_secheur_bioco").Cells(1, pi_FiliereConsideree)
        If choix_secheur_bioco = True Then
            verification_configuration = verification_configuration + 1
        End If
        If pn_EffacementResultatsEnCours = False Then
            If verification_configuration = 0 Then
                message_erreur = Feuil5.Range("erreur_choix_sechage").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub    'met fin aux calculs
            End If
        End If
    Else
        choix_secheur_autre = False
        choix_secheur_inos = False
        choix_secheur_bioco = False
    End If

    choix_incineration = Feuil5.Range("choix_incineration").Cells(1, pi_FiliereConsideree)
    If choix_incineration = True Then
        count_traitement = count_traitement + 1
    End If

    choix_traitement_retours = Feuil5.Range("choix_traitement_retours").Cells(1, pi_FiliereConsideree)
    If choix_traitement_retours = True Then
        verification_configuration = 0
        count_traitement = count_traitement + 1
        choix_MAP_retours = Feuil5.Range("choix_struvite").Cells(1, pi_FiliereConsideree)
        'On ne vérifie pas la configuration pour le MAP car il peut être combiné avec les autres traitements
        choix_ANITA_Mox = Feuil5.Range("choix_ANITA_Mox").Cells(1, pi_FiliereConsideree)
        If choix_ANITA_Mox = True Then
            verification_configuration = verification_configuration + 1
        End If
        choix_ANITA_Shunt = Feuil5.Range("choix_ANITA_Shunt").Cells(1, pi_FiliereConsideree)
        If choix_ANITA_Shunt = True Then
            verification_configuration = verification_configuration + 1
        End If
        If pn_EffacementResultatsEnCours = False Then
            If (verification_configuration = 0 And choix_MAP_retours = False) Or verification_configuration > 1 Or pi_ChoixDigestion = False Then
                message_erreur = Feuil5.Range("erreur_choix_traitement_retours").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub    'met fin aux calculs
            End If
        End If
    Else
        choix_MAP_retours = False
        choix_ANITA_Mox = False
        choix_ANITA_Shunt = False
    End If

    'choix désodorisation chimique
    choix_desodo_chimique = Feuil5.Range("choix_desodo_chimique").Cells(1, pi_FiliereConsideree)
    'choix désodorisation biologique
    choix_desodo_bio = Feuil5.Range("choix_desodo_bio").Cells(1, pi_FiliereConsideree)

    'UTILITES
    choix_utilites_eclairage = Feuil5.Range("choix_utilites_eclairage").Cells(1, pi_FiliereConsideree)
    choix_utilites_pertes_enligne = Feuil5.Range("choix_utilites_pertes_enligne").Cells(1, pi_FiliereConsideree)
    choix_utilites_eau_service = Feuil5.Range("choix_utilites_eau_service").Cells(1, pi_FiliereConsideree)
    choix_utilites_pompage_retours = Feuil5.Range("choix_utilites_pompage_retours").Cells(1, pi_FiliereConsideree)
    choix_utilites_chauffage_clim = Feuil5.Range("choix_utilites_chauffage_clim").Cells(1, pi_FiliereConsideree)
    If choix_utilites_chauffage_clim = True Then
        choix_utilites_bat_administration = Feuil5.Range("choix_utilites_bat_administration").Cells(1, pi_FiliereConsideree)
        choix_utilites_bat_exploitation = Feuil5.Range("choix_utilites_bat_exploitation").Cells(1, pi_FiliereConsideree)
        choix_utilites_bat_electrique = Feuil5.Range("choix_utilites_bat_electrique").Cells(1, pi_FiliereConsideree)
    End If

    'ENERGIES ALTERNATIVES
    '## Production alternative de chaleur
    choix_prod_alt_chaleur = Feuil5.Range("choix_prod_alt_chaleur").Cells(1, pi_FiliereConsideree)
    If choix_prod_alt_chaleur = True Then
        choix_PAC_eau_traitee = Feuil5.Range("choix_PAC_eau_traitee").Cells(1, pi_FiliereConsideree)
    End If

    '## Production alternative d'électricité
    choix_prod_alt_electricite = Feuil5.Range("choix_prod_alt_electricite").Cells(1, pi_FiliereConsideree)
    If choix_prod_alt_electricite = True Then
        choix_solaire_photovoltaique = Feuil5.Range("choix_solaire_photovoltaique").Cells(1, pi_FiliereConsideree)
        choix_turbine_hydraulique = Feuil5.Range("choix_turbine_hydraulique").Cells(1, pi_FiliereConsideree)
        choix_electricite_autre_production = Feuil5.Range("choix_electricite_autre_production").Cells(1, pi_FiliereConsideree)
        'Message d'erreur
        If pn_EffacementResultatsEnCours = False Then
            If choix_solaire_photovoltaique = False And choix_turbine_hydraulique = False And choix_electricite_autre_production = False Then
                message_erreur = Feuil5.Range("erreur_choix_energies_alternatives").Cells(1, 1)
                MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                pn_EffacementResultatsEnCours = True
                pn_ErreurNonGeree = True
                Exit Sub 'met fin aux calculs
            End If
        End If
    End If



    If pn_EffacementResultatsEnCours = False Then
        If count_traitement = 0 Then
            message_erreur = Feuil5.Range("erreur_count_traitement").Cells(1, 1)
            MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
            pn_EffacementResultatsEnCours = True
            pn_ErreurNonGeree = True
            Exit Sub    'met fin aux calculs
        End If
    End If

    'configuration des biostyrs
    If pn_EffacementResultatsEnCours = False Then
        If verif_nb_biofiltre = 2 Then
            message_erreur = Feuil5.Range("erreur_choix_biostyr").Cells(1, 1)
            MsgBox message_erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
            pn_EffacementResultatsEnCours = True
            pn_ErreurNonGeree = True
            Exit Sub    'met fin aux calculs
        End If
    End If


    'on lit le choix pour les graisses extérieures
    choix_coferments = Feuil5.Range("choix_graisses_externes").Cells(1, pi_FiliereConsideree)


End Sub



Sub erreur_configuration_process(ByRef treatment_message)


    treatment_message = treatment_message & Feuil6.Range("erreur_configuration")

    Erreur_configuration.message_erreur_conf.Caption = treatment_message

    Erreur_configuration.Show

    'après la fermeture du userform
    Erreur_configuration.message_erreur_conf.Caption = ""


End Sub

Sub recapitulatif_boues(ByRef boues_flux, ByRef boues_pollution_soluble, ByRef graisses_particulaire_ratio_interne, ByVal boues_ratio_NK_MV, ByVal boues_ratio_Pt_MES, ByVal boues_ratio_DCO_MV, ByVal boues_ratio_DBO_MV, ByVal graisses_internes, ByVal nb_max_epaississeur, ByVal boues_Shunt, ByVal choix_eaux_sales_1aire, ByVal boues_I_detail)

    'Ici, on vient récupérer les données de boues externes mais ausii des boues I,II et III pour les filières complexes

    'constantes
    Const choix_non = 0

    Dim I As Integer
    Dim k As Integer

    'Déclaration des variables locales
    Dim choix_boues_I As Integer
    Dim choix_boues_II As Integer
    Dim choix_boues_III As Integer
    Dim choix_boues_externes_1 As Integer
    Dim choix_boues_externes_2 As Integer
    'Dim choix_graisses As Integer
    Dim nb_parametres_echanges As Integer
    Dim intermediaire_MES As Double
    Dim ratio_boues_bio As Double

    'flux volumique de boues m3/j intermediaire
    Dim intermediaire_Q_m3j As Double

    'graphique répartition des boues
    Dim graph_boues_repartition(nb_type_boues + 1) As Double    '+1 pour séparer graisses et coferments
    nb_parametres_echanges = repere_MV_MES

    'LECTURE DES CHOIX
    choix_boues_I = Feuil5.Range("choix_boues_I").Cells(1, pi_FiliereConsideree)
    choix_boues_II = Feuil5.Range("choix_boues_II").Cells(1, pi_FiliereConsideree)
    choix_boues_III = Feuil5.Range("choix_boues_III").Cells(1, pi_FiliereConsideree)
    choix_boues_externes_1 = Feuil5.Range("choix_boues_externes_1").Cells(1, pi_FiliereConsideree)
    choix_boues_externes_2 = Feuil5.Range("choix_boues_externes_2").Cells(1, pi_FiliereConsideree)
    'choix_graisses = Feuil5.Range("choix_graisses_externes").Cells(1, pi_FiliereConsideree)

    'ANALYSE DES CHOIX
    'Boues I
    If choix_boues_I <> choix_non Then
        boues_flux(boues_inlet, boues_I, repere_origine) = choix_boues_I
        For I = 2 To nb_parametres_echanges
            boues_flux(boues_inlet, boues_I, I) = Feuil6.Range("boues_I_caracteristiques").Cells(I, pi_FiliereConsideree)
        Next I
        boues_flux(boues_inlet, boues_I, repere_ratio_DCO_MES) = boues_ratio_DCO_MV(choix_boues_I) * boues_flux(boues_inlet, boues_I, repere_MV_MES)
        boues_flux(boues_inlet, boues_I, repere_ratio_DBO_MES) = boues_ratio_DBO_MV(choix_boues_I) * boues_flux(boues_inlet, boues_I, repere_MV_MES)
        boues_flux(boues_inlet, boues_I, repere_ratio_NK_MES) = boues_ratio_NK_MV(choix_boues_I) * boues_flux(boues_inlet, boues_I, repere_MV_MES)
        boues_flux(boues_inlet, boues_I, repere_ratio_Pt_MES) = boues_ratio_Pt_MES(choix_boues_I)
        boues_flux(boues_inlet, boues_I, repere_verif_flux) = 1
        boues_flux(boues_inlet, boues_I, repere_flux_in) = 1
        For I = boues_epaissies(1) To boues_epaissies(nb_max_epaississeur)
            boues_flux(I, boues_I, repere_MV_MES) = boues_flux(boues_inlet, boues_I, repere_MV_MES)
            boues_flux(I, boues_I, repere_ratio_DCO_MES) = boues_flux(boues_inlet, boues_I, repere_ratio_DCO_MES)
            boues_flux(I, boues_I, repere_ratio_DBO_MES) = boues_flux(boues_inlet, boues_I, repere_ratio_DBO_MES)
            boues_flux(I, boues_I, repere_ratio_NK_MES) = boues_flux(boues_inlet, boues_I, repere_ratio_NK_MES)
            boues_flux(I, boues_I, repere_ratio_Pt_MES) = boues_flux(boues_inlet, boues_I, repere_ratio_Pt_MES)
        Next I
        For I = boues_epaissies(1) To nb_etape_file_boues
            boues_flux(I, boues_I, repere_origine) = boues_flux(boues_inlet, boues_I, repere_origine)
        Next I
    End If

    'Boues II
    If choix_boues_II <> choix_non Then
        boues_flux(boues_inlet, boues_II, repere_origine) = choix_boues_II
        For I = 2 To nb_parametres_echanges
            boues_flux(boues_inlet, boues_II, I) = Feuil6.Range("boues_II_caracteristiques").Cells(I, pi_FiliereConsideree)
        Next I
        boues_flux(boues_inlet, boues_II, repere_ratio_DCO_MES) = boues_ratio_DCO_MV(choix_boues_II) * boues_flux(boues_inlet, boues_II, repere_MV_MES)
        boues_flux(boues_inlet, boues_II, repere_ratio_DBO_MES) = boues_ratio_DBO_MV(choix_boues_II) * boues_flux(boues_inlet, boues_II, repere_MV_MES)
        boues_flux(boues_inlet, boues_II, repere_ratio_NK_MES) = boues_ratio_NK_MV(choix_boues_II) * boues_flux(boues_inlet, boues_II, repere_MV_MES)
        boues_flux(boues_inlet, boues_II, repere_ratio_Pt_MES) = boues_ratio_Pt_MES(choix_boues_II)
        boues_flux(boues_inlet, boues_II, repere_verif_flux) = 1
        boues_flux(boues_inlet, boues_II, repere_flux_in) = 1
        For I = boues_epaissies(1) To boues_epaissies(nb_max_epaississeur)
            boues_flux(I, boues_II, repere_MV_MES) = boues_flux(boues_inlet, boues_II, repere_MV_MES)
            boues_flux(I, boues_II, repere_ratio_DCO_MES) = boues_flux(boues_inlet, boues_II, repere_ratio_DCO_MES)
            boues_flux(I, boues_II, repere_ratio_DBO_MES) = boues_flux(boues_inlet, boues_II, repere_ratio_DBO_MES)
            boues_flux(I, boues_II, repere_ratio_NK_MES) = boues_flux(boues_inlet, boues_II, repere_ratio_NK_MES)
            boues_flux(I, boues_II, repere_ratio_Pt_MES) = boues_flux(boues_inlet, boues_II, repere_ratio_Pt_MES)
        Next I
        For I = boues_epaissies(1) To nb_etape_file_boues
            boues_flux(I, boues_II, repere_origine) = boues_flux(boues_inlet, boues_II, repere_origine)
        Next I
    End If

    'Boues III
    If choix_boues_III <> choix_non Then
        boues_flux(boues_inlet, boues_III, repere_origine) = choix_boues_III
        For I = 2 To nb_parametres_echanges
            boues_flux(boues_inlet, boues_III, I) = Feuil6.Range("boues_III_caracteristiques").Cells(I, pi_FiliereConsideree)
        Next I
        boues_flux(boues_inlet, boues_III, repere_ratio_DCO_MES) = boues_ratio_DCO_MV(choix_boues_III) * boues_flux(boues_inlet, boues_III, repere_MV_MES)
        boues_flux(boues_inlet, boues_III, repere_ratio_DBO_MES) = boues_ratio_DBO_MV(choix_boues_III) * boues_flux(boues_inlet, boues_III, repere_MV_MES)
        boues_flux(boues_inlet, boues_III, repere_ratio_NK_MES) = boues_ratio_NK_MV(choix_boues_III) * boues_flux(boues_inlet, boues_III, repere_MV_MES)
        boues_flux(boues_inlet, boues_III, repere_ratio_Pt_MES) = boues_ratio_Pt_MES(choix_boues_III)
        boues_flux(boues_inlet, boues_III, repere_verif_flux) = 1
        boues_flux(boues_inlet, boues_III, repere_flux_in) = 1
        For I = boues_epaissies(1) To boues_epaissies(nb_max_epaississeur)
            boues_flux(I, boues_III, repere_MV_MES) = boues_flux(boues_inlet, boues_III, repere_MV_MES)
            boues_flux(I, boues_III, repere_ratio_DCO_MES) = boues_flux(boues_inlet, boues_III, repere_ratio_DCO_MES)
            boues_flux(I, boues_III, repere_ratio_DBO_MES) = boues_flux(boues_inlet, boues_III, repere_ratio_DBO_MES)
            boues_flux(I, boues_III, repere_ratio_NK_MES) = boues_flux(boues_inlet, boues_III, repere_ratio_NK_MES)
            boues_flux(I, boues_III, repere_ratio_Pt_MES) = boues_flux(boues_inlet, boues_III, repere_ratio_Pt_MES)
        Next I
        For I = boues_epaissies(1) To nb_etape_file_boues
            boues_flux(I, boues_III, repere_origine) = boues_flux(boues_inlet, boues_III, repere_origine)
        Next I
    End If

    'Boues externes 1
    If choix_boues_externes_1 <> choix_non Then
        boues_flux(boues_inlet, boues_externes_1, repere_origine) = choix_boues_externes_1
        For I = 2 To nb_parametres_echanges
            boues_flux(boues_inlet, boues_externes_1, I) = Feuil6.Range("boues_externes_1_caracteristiques").Cells(I, pi_FiliereConsideree)
        Next I
        boues_flux(boues_inlet, boues_externes_1, repere_ratio_DCO_MES) = boues_ratio_DCO_MV(choix_boues_externes_1) * boues_flux(boues_inlet, boues_externes_1, repere_MV_MES)
        boues_flux(boues_inlet, boues_externes_1, repere_ratio_DBO_MES) = boues_ratio_DBO_MV(choix_boues_externes_1) * boues_flux(boues_inlet, boues_externes_1, repere_MV_MES)
        boues_flux(boues_inlet, boues_externes_1, repere_ratio_NK_MES) = boues_ratio_NK_MV(choix_boues_externes_1) * boues_flux(boues_inlet, boues_externes_1, repere_MV_MES)
        boues_flux(boues_inlet, boues_externes_1, repere_ratio_Pt_MES) = boues_ratio_Pt_MES(choix_boues_externes_1)
        boues_flux(boues_inlet, boues_externes_1, repere_verif_flux) = 1
        boues_flux(boues_inlet, boues_externes_1, repere_flux_in) = 1
        For I = boues_epaissies(1) To boues_epaissies(nb_max_epaississeur)
            boues_flux(I, boues_externes_1, repere_MV_MES) = boues_flux(boues_inlet, boues_externes_1, repere_MV_MES)
            boues_flux(I, boues_externes_1, repere_ratio_DCO_MES) = boues_flux(boues_inlet, boues_externes_1, repere_ratio_DCO_MES)
            boues_flux(I, boues_externes_1, repere_ratio_DBO_MES) = boues_flux(boues_inlet, boues_externes_1, repere_ratio_DBO_MES)
            boues_flux(I, boues_externes_1, repere_ratio_NK_MES) = boues_flux(boues_inlet, boues_externes_1, repere_ratio_NK_MES)
            boues_flux(I, boues_externes_1, repere_ratio_Pt_MES) = boues_flux(boues_inlet, boues_externes_1, repere_ratio_Pt_MES)
        Next I
        For I = boues_epaissies(1) To nb_etape_file_boues
            boues_flux(I, boues_externes_1, repere_origine) = boues_flux(boues_inlet, boues_externes_1, repere_origine)
        Next I
    End If

    'Boues externes 2
    If choix_boues_externes_2 <> choix_non Then
        boues_flux(boues_inlet, boues_externes_2, repere_origine) = choix_boues_externes_2
        For I = 2 To nb_parametres_echanges
            boues_flux(boues_inlet, boues_externes_2, I) = Feuil6.Range("boues_externes_2_caracteristiques").Cells(I, pi_FiliereConsideree)
        Next I
        boues_flux(boues_inlet, boues_externes_2, repere_ratio_DCO_MES) = boues_ratio_DCO_MV(choix_boues_externes_2) * boues_flux(boues_inlet, boues_externes_2, repere_MV_MES)
        boues_flux(boues_inlet, boues_externes_2, repere_ratio_DBO_MES) = boues_ratio_DBO_MV(choix_boues_externes_2) * boues_flux(boues_inlet, boues_externes_2, repere_MV_MES)
        boues_flux(boues_inlet, boues_externes_2, repere_ratio_NK_MES) = boues_ratio_NK_MV(choix_boues_externes_2) * boues_flux(boues_inlet, boues_externes_2, repere_MV_MES)
        boues_flux(boues_inlet, boues_externes_2, repere_ratio_Pt_MES) = boues_ratio_Pt_MES(choix_boues_externes_2)
        boues_flux(boues_inlet, boues_externes_2, repere_verif_flux) = 1
        boues_flux(boues_inlet, boues_externes_2, repere_flux_in) = 1
        For I = boues_epaissies(1) To boues_epaissies(nb_max_epaississeur)
            boues_flux(I, boues_externes_2, repere_MV_MES) = boues_flux(boues_inlet, boues_externes_2, repere_MV_MES)
            boues_flux(I, boues_externes_2, repere_ratio_DCO_MES) = boues_flux(boues_inlet, boues_externes_2, repere_ratio_DCO_MES)
            boues_flux(I, boues_externes_2, repere_ratio_DBO_MES) = boues_flux(boues_inlet, boues_externes_2, repere_ratio_DBO_MES)
            boues_flux(I, boues_externes_2, repere_ratio_NK_MES) = boues_flux(boues_inlet, boues_externes_2, repere_ratio_NK_MES)
            boues_flux(I, boues_externes_2, repere_ratio_Pt_MES) = boues_flux(boues_inlet, boues_externes_2, repere_ratio_Pt_MES)
        Next I
        For I = boues_epaissies(1) To nb_etape_file_boues
            boues_flux(I, boues_externes_2, repere_origine) = boues_flux(boues_inlet, boues_externes_2, repere_origine)
        Next I
    End If

    'Coferments
    If choix_coferments <> choix_non Then
        'on lit les co-ferments
        boues_flux(boues_inlet, graisses, repere_Q) = Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_Q_m3j, pi_FiliereConsideree)
        'boues_flux(boues_inlet, graisses, repere_MS) = Feuil6.Range("graisses_externes_caracteristiques").Cells(co_ferment_m3j, pi_FiliereConsideree) * Feuil6.Range("graisses_externes_caracteristiques").Cells(co_ferment_gMS_L, pi_FiliereConsideree)
        boues_flux(boues_inlet, graisses, repere_MES) = Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_Q_m3j, pi_FiliereConsideree) * Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_MS_gL, pi_FiliereConsideree) * Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_ratio_MES_MS, pi_FiliereConsideree)
        'hypothèse: MV/MES = MVtot/MS = MVsol/MSsol  ici MV/MES ne concerne que le particulaire
        'hypothèse identique pour le reste
        boues_flux(boues_inlet, graisses, repere_MV_MES) = Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_ratio_MV_MS, pi_FiliereConsideree)
        boues_flux(boues_inlet, graisses, repere_ratio_DCO_MES) = Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_ratio_DCO_MV, pi_FiliereConsideree) * boues_flux(boues_inlet, graisses, repere_MV_MES)
        boues_flux(boues_inlet, graisses, repere_ratio_DBO_MES) = boues_ratio_DBO_MV(codigestion_graisses) * boues_flux(boues_inlet, graisses, repere_MV_MES)
        If boues_flux(boues_inlet, graisses, repere_ratio_DBO_MES) > boues_flux(boues_inlet, graisses, repere_ratio_DCO_MES) Then
            boues_flux(boues_inlet, graisses, repere_ratio_DBO_MES) = 0.9 * boues_flux(boues_inlet, graisses, repere_ratio_DCO_MES)  'hypothèse: ratio DBO/DCO des graisses=0.9 (utilisé si DCO/MES<DBO/MES)
        End If
        boues_flux(boues_inlet, graisses, repere_ratio_NK_MES) = Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_ratio_NK_MV, pi_FiliereConsideree) * boues_flux(boues_inlet, graisses, repere_MV_MES)
        boues_flux(boues_inlet, graisses, repere_ratio_Pt_MES) = Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_ratio_Pt_MV, pi_FiliereConsideree) * boues_flux(boues_inlet, graisses, repere_MV_MES)
        boues_flux(boues_inlet, graisses, repere_verif_flux) = 1
        boues_flux(boues_inlet, graisses, repere_flux_in) = 1
        'on remplit pour le soluble (dans un premier temps en kg/j puis on divisera par le débit total)
        boues_pollution_soluble(boues_inlet, graisses, repere_mgL_MS_soluble) = Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_Q_m3j, pi_FiliereConsideree) * Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_MS_gL, pi_FiliereConsideree) - boues_flux(boues_inlet, graisses, repere_MES)
        boues_pollution_soluble(boues_inlet, graisses, repere_mgL_DCO) = boues_pollution_soluble(boues_inlet, graisses, repere_mgL_MS_soluble) * boues_flux(boues_inlet, graisses, repere_ratio_DCO_MES)
        boues_pollution_soluble(boues_inlet, graisses, repere_mgL_NK) = boues_pollution_soluble(boues_inlet, graisses, repere_mgL_MS_soluble) * boues_flux(boues_inlet, graisses, repere_ratio_NK_MES)
        boues_pollution_soluble(boues_inlet, graisses, repere_mgL_Pt) = boues_pollution_soluble(boues_inlet, graisses, repere_mgL_MS_soluble) * boues_flux(boues_inlet, graisses, repere_ratio_Pt_MES)
        boues_pollution_soluble(boues_inlet, graisses, repere_mgL_MV_soluble) = boues_pollution_soluble(boues_inlet, graisses, repere_mgL_MS_soluble) * boues_flux(boues_inlet, graisses, repere_MV_MES)
    End If

    If choix_coferments <> choix_non Or choix_dessablage = True Then
        'calcul des graisses totales (internes + externes)
        boues_flux(boues_inlet, graisses, repere_origine) = codigestion_graisses
        boues_flux(boues_inlet, graisses, repere_Q) = boues_flux(boues_inlet, graisses, repere_Q) + graisses_internes(repere_graisse_Q)
        intermediaire_MES = boues_flux(boues_inlet, graisses, repere_MES) + graisses_internes(repere_graisse_MS)
        boues_flux(boues_inlet, graisses, repere_MV_MES) = (boues_flux(boues_inlet, graisses, repere_MV_MES) * boues_flux(boues_inlet, graisses, repere_MES) + graisses_internes(repere_graisse_MS) * graisses_internes(repere_graisse_MV_MES)) / intermediaire_MES
        boues_flux(boues_inlet, graisses, repere_ratio_DCO_MES) = (boues_flux(boues_inlet, graisses, repere_ratio_DCO_MES) * boues_flux(boues_inlet, graisses, repere_MES) + graisses_internes(repere_graisse_MS) * graisses_internes(repere_graisse_DCO_MES)) / intermediaire_MES
        boues_flux(boues_inlet, graisses, repere_ratio_DBO_MES) = (boues_flux(boues_inlet, graisses, repere_ratio_DBO_MES) * boues_flux(boues_inlet, graisses, repere_MES) + graisses_internes(repere_graisse_MS) * graisses_internes(repere_graisse_DBO_MES)) / intermediaire_MES
        boues_flux(boues_inlet, graisses, repere_ratio_NK_MES) = (boues_flux(boues_inlet, graisses, repere_ratio_NK_MES) * boues_flux(boues_inlet, graisses, repere_MES) + graisses_internes(repere_graisse_MS) * graisses_internes(repere_graisse_NK_MES)) / intermediaire_MES
        boues_flux(boues_inlet, graisses, repere_ratio_Pt_MES) = (boues_flux(boues_inlet, graisses, repere_ratio_Pt_MES) * boues_flux(boues_inlet, graisses, repere_MES) + graisses_internes(repere_graisse_MS) * graisses_internes(repere_graisse_Pt_MES)) / intermediaire_MES
        boues_flux(boues_inlet, graisses, repere_MES) = intermediaire_MES
        boues_flux(boues_inlet, graisses, repere_verif_flux) = 1
        boues_flux(boues_inlet, graisses, repere_flux_in) = 1
        graisses_particulaire_ratio_interne(repere_graisse_particulaire_ratioMES) = graisses_internes(repere_graisse_MS) / boues_flux(boues_inlet, graisses, repere_MES)
        graisses_particulaire_ratio_interne(repere_graisse_particulaire_ratioMV) = (graisses_internes(repere_graisse_MS) * graisses_internes(repere_graisse_MV_MES)) / (boues_flux(boues_inlet, graisses, repere_MES) * boues_flux(boues_inlet, graisses, repere_MV_MES))
        For I = 1 To nb_repere_mgL
            boues_pollution_soluble(boues_inlet, graisses, I) = boues_pollution_soluble(boues_inlet, graisses, I) / boues_flux(boues_inlet, graisses, repere_Q) * 1000
            'For k = boues_epaissies(1) To boues_epaissies(nb_max_epaississeur)
            '    boues_pollution_soluble(k, graisses, i) = boues_pollution_soluble(boues_inlet, graisses, i)
            'Next k
        Next I
    End If

    For I = boues_epaissies(1) To boues_epaissies(nb_max_epaississeur)
        boues_flux(I, graisses, repere_MV_MES) = boues_flux(boues_inlet, graisses, repere_MV_MES)
        boues_flux(I, graisses, repere_ratio_DCO_MES) = boues_flux(boues_inlet, graisses, repere_ratio_DCO_MES)
        boues_flux(I, graisses, repere_ratio_DBO_MES) = boues_flux(boues_inlet, graisses, repere_ratio_DBO_MES)
        boues_flux(I, graisses, repere_ratio_NK_MES) = boues_flux(boues_inlet, graisses, repere_ratio_NK_MES)
        boues_flux(I, graisses, repere_ratio_Pt_MES) = boues_flux(boues_inlet, graisses, repere_ratio_Pt_MES)
    Next I
    For I = boues_epaissies(1) To nb_etape_file_boues
        boues_flux(I, graisses, repere_origine) = boues_flux(boues_inlet, graisses, repere_origine)
    Next I

    'ON AJOUTE LES BOUES DU SHUNT AVEC LES BOUES 2aire OU 1aire SI PAS DE 2aire
    If choix_ANITA_Shunt = True Then
        If boues_flux(boues_inlet, boues_II, repere_MES) > 0 Then
            boues_flux(boues_inlet, boues_II, repere_Q) = boues_flux(boues_inlet, boues_II, repere_Q) + boues_Shunt(repere_graisse_Q)
            intermediaire_MES = boues_flux(boues_inlet, boues_II, repere_MES) + boues_Shunt(repere_graisse_MS)
            boues_flux(boues_inlet, boues_II, repere_ratio_DCO_MES) = (boues_flux(boues_inlet, boues_II, repere_ratio_DCO_MES) * boues_flux(boues_inlet, boues_II, repere_MES) + boues_Shunt(repere_graisse_MS) * boues_Shunt(repere_graisse_DCO_MES)) / intermediaire_MES
            boues_flux(boues_inlet, boues_II, repere_ratio_DBO_MES) = (boues_flux(boues_inlet, boues_II, repere_ratio_DBO_MES) * boues_flux(boues_inlet, boues_II, repere_MES) + boues_Shunt(repere_graisse_MS) * boues_Shunt(repere_graisse_DBO_MES)) / intermediaire_MES
            boues_flux(boues_inlet, boues_II, repere_ratio_NK_MES) = (boues_flux(boues_inlet, boues_II, repere_ratio_NK_MES) * boues_flux(boues_inlet, boues_II, repere_MES) + boues_Shunt(repere_graisse_MS) * boues_Shunt(repere_graisse_NK_MES)) / intermediaire_MES
            boues_flux(boues_inlet, boues_II, repere_ratio_Pt_MES) = (boues_flux(boues_inlet, boues_II, repere_ratio_Pt_MES) * boues_flux(boues_inlet, boues_II, repere_MES) + boues_Shunt(repere_graisse_MS) * boues_Shunt(repere_graisse_Pt_MES)) / intermediaire_MES
            boues_flux(boues_inlet, boues_II, repere_MES) = intermediaire_MES
            boues_flux(boues_inlet, boues_II, repere_verif_flux) = 1
            boues_flux(boues_inlet, boues_II, repere_flux_in) = 1
            For I = boues_epaissies(1) To boues_epaissies(nb_max_epaississeur)
                boues_flux(I, boues_II, repere_MV_MES) = boues_flux(boues_inlet, boues_II, repere_MV_MES)
                boues_flux(I, boues_II, repere_ratio_DCO_MES) = boues_flux(boues_inlet, boues_II, repere_ratio_DCO_MES)
                boues_flux(I, boues_II, repere_ratio_DBO_MES) = boues_flux(boues_inlet, boues_II, repere_ratio_DBO_MES)
                boues_flux(I, boues_II, repere_ratio_NK_MES) = boues_flux(boues_inlet, boues_II, repere_ratio_NK_MES)
                boues_flux(I, boues_II, repere_ratio_Pt_MES) = boues_flux(boues_inlet, boues_II, repere_ratio_Pt_MES)
            Next I
        ElseIf boues_flux(boues_inlet, boues_I, repere_MES) > 0 Then
            boues_flux(boues_inlet, boues_I, repere_Q) = boues_flux(boues_inlet, boues_I, repere_Q) + boues_Shunt(repere_graisse_Q)
            intermediaire_MES = boues_flux(boues_inlet, boues_I, repere_MES) + boues_Shunt(repere_graisse_MS)
            boues_flux(boues_inlet, boues_I, repere_ratio_DCO_MES) = (boues_flux(boues_inlet, boues_I, repere_ratio_DCO_MES) * boues_flux(boues_inlet, boues_I, repere_MES) + boues_Shunt(repere_graisse_MS) * boues_Shunt(repere_graisse_DCO_MES)) / intermediaire_MES
            boues_flux(boues_inlet, boues_I, repere_ratio_DBO_MES) = (boues_flux(boues_inlet, boues_I, repere_ratio_DBO_MES) * boues_flux(boues_inlet, boues_I, repere_MES) + boues_Shunt(repere_graisse_MS) * boues_Shunt(repere_graisse_DBO_MES)) / intermediaire_MES
            boues_flux(boues_inlet, boues_I, repere_ratio_NK_MES) = (boues_flux(boues_inlet, boues_I, repere_ratio_NK_MES) * boues_flux(boues_inlet, boues_I, repere_MES) + boues_Shunt(repere_graisse_MS) * boues_Shunt(repere_graisse_NK_MES)) / intermediaire_MES
            boues_flux(boues_inlet, boues_I, repere_ratio_Pt_MES) = (boues_flux(boues_inlet, boues_I, repere_ratio_Pt_MES) * boues_flux(boues_inlet, boues_I, repere_MES) + boues_Shunt(repere_graisse_MS) * boues_Shunt(repere_graisse_Pt_MES)) / intermediaire_MES
            boues_flux(boues_inlet, boues_I, repere_MES) = intermediaire_MES
            boues_flux(boues_inlet, boues_I, repere_verif_flux) = 1
            boues_flux(boues_inlet, boues_I, repere_flux_in) = 1
            For I = boues_epaissies(1) To boues_epaissies(nb_max_epaississeur)
                boues_flux(I, boues_I, repere_MV_MES) = boues_flux(boues_inlet, boues_I, repere_MV_MES)
                boues_flux(I, boues_I, repere_ratio_DCO_MES) = boues_flux(boues_inlet, boues_I, repere_ratio_DCO_MES)
                boues_flux(I, boues_I, repere_ratio_DBO_MES) = boues_flux(boues_inlet, boues_I, repere_ratio_DBO_MES)
                boues_flux(I, boues_I, repere_ratio_NK_MES) = boues_flux(boues_inlet, boues_I, repere_ratio_NK_MES)
                boues_flux(I, boues_I, repere_ratio_Pt_MES) = boues_flux(boues_inlet, boues_I, repere_ratio_Pt_MES)
            Next I
        End If
    End If

    'On fait les calculs pour la partie soluble des co-ferments
    'intermediaire_Q_m3j = 0
    'For i = 1 To nb_type_boues
    '    intermediaire_Q_m3j = intermediaire_Q_m3j + boues_flux(boues_inlet, i, repere_Q)
    'Next i






    'Récapitulatif des boues extraites de la file eau   POUR LES GRAPHES
    intermediaire_MES = 0
    For I = 1 To nb_type_boues
        graph_boues_repartition(I) = 0
    Next I
    For I = 1 To nb_type_boues
        If I = boues_I And choix_eaux_sales_1aire = True Then
            ratio_boues_bio = 0
            For k = 1 To nb_eaux_sales_max
                ratio_boues_bio = ratio_boues_bio + boues_I_detail(k, repere_ratio_MES)
                If boues_I_detail(k, repere_origine) >= II_forte And boues_I_detail(k, repere_origine) < III_decantation Then
                    graph_boues_repartition(boues_II) = graph_boues_repartition(boues_II) + boues_flux(boues_inlet, I, repere_MES) * boues_I_detail(k, repere_ratio_MES)
                    intermediaire_MES = intermediaire_MES + boues_flux(boues_inlet, I, repere_MES) * boues_I_detail(k, repere_ratio_MES)
                ElseIf boues_I_detail(k, repere_origine) >= III_decantation And boues_I_detail(k, repere_origine) < codigestion_graisses Then
                    graph_boues_repartition(boues_III) = graph_boues_repartition(boues_III) + boues_flux(boues_inlet, I, repere_MES) * boues_I_detail(k, repere_ratio_MES)
                    intermediaire_MES = intermediaire_MES + boues_flux(boues_inlet, I, repere_MES) * boues_I_detail(k, repere_ratio_MES)
                End If
            Next k
            graph_boues_repartition(I) = graph_boues_repartition(I) + boues_flux(boues_inlet, I, repere_MES) * (1 - ratio_boues_bio)
            intermediaire_MES = intermediaire_MES + boues_flux(boues_inlet, I, repere_MES) * (1 - ratio_boues_bio)
        ElseIf I = graisses Then
            intermediaire_MES = intermediaire_MES + boues_flux(boues_inlet, I, repere_MES)
            'Internes
            graph_boues_repartition(I) = graph_boues_repartition(I) + graisses_particulaire_ratio_interne(repere_graisse_particulaire_ratioMES) * boues_flux(boues_inlet, I, repere_MES)
            'coferments
            graph_boues_repartition(I + 1) = graph_boues_repartition(I + 1) + (1 - graisses_particulaire_ratio_interne(repere_graisse_particulaire_ratioMES)) * boues_flux(boues_inlet, I, repere_MES)
            graph_boues_repartition(I + 1) = graph_boues_repartition(I + 1) + boues_pollution_soluble(boues_inlet, graisses, repere_mgL_MS_soluble) * boues_flux(boues_inlet, I, repere_Q) / 1000
            intermediaire_MES = intermediaire_MES + boues_pollution_soluble(boues_inlet, graisses, repere_mgL_MS_soluble) * boues_flux(boues_inlet, I, repere_Q) / 1000
        Else
            graph_boues_repartition(I) = graph_boues_repartition(I) + boues_flux(boues_inlet, I, repere_MES)
            intermediaire_MES = intermediaire_MES + boues_flux(boues_inlet, I, repere_MES)
        End If
    Next I

    For I = 1 To nb_type_boues + 1
        If intermediaire_MES <> 0 Then
            Feuil37.Range("graph_boues_step_" & pi_FiliereConsideree).Cells(I, 1) = graph_boues_repartition(I) / intermediaire_MES
        Else
            Feuil37.Range("graph_boues_step_" & pi_FiliereConsideree).Cells(I, 1) = 0
        End If
    Next I


End Sub

Sub verification_boues(ByRef boues_flux, ByRef warning, ByVal cendres_tMSj, ByVal REFIB_tMSj, ByRef boues_evacuees_Q, ByRef graisses_evacuees_Q)

    'Cette routine a pour but de vérifier que l'on ne traite pas plusieurs fois un même type de boues et aussi de déterminer les boues à évacuer


    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim verif As Double
    Dim erreur As String
    Dim rapport As Double

    Dim somme As Double

    Dim boues_totales_in_MS As Double
    Dim boues_totales_out_MS As Double

    boues_evacuees_Q = 0
    graisses_evacuees_Q = 0
    For I = 1 To nb_type_boues
        If boues_flux(boues_inlet, I, repere_flux_in) = 1 Then
            verif = 0
            For j = 1 To nb_etape_file_boues - 1
                verif = verif + boues_flux(j, I, repere_verif_flux)
                If boues_flux(j, I, repere_verif_flux) > 0 Then
                    rapport = boues_flux(j, I, repere_verif_flux) / boues_flux(j, I, repere_flux_in)
                    If I = graisses Then
                        somme = 0
                        For k = 1 To nb_type_boues
                            If k <> I Then
                                somme = somme + boues_flux(j, k, repere_Q)
                            End If
                        Next k
                        If somme > 0 And j <> boues_inlet Then  'graisses mélangées à de la boue
                            boues_evacuees_Q = boues_evacuees_Q + rapport * boues_flux(j, I, repere_Q)
                        Else   'graisses seules
                            graisses_evacuees_Q = graisses_evacuees_Q + rapport * boues_flux(j, I, repere_Q)
                        End If
                    Else
                        boues_evacuees_Q = boues_evacuees_Q + rapport * boues_flux(j, I, repere_Q)
                    End If
                    boues_flux(boues_evacuees, I, repere_Q) = boues_flux(boues_evacuees, I, repere_Q) + rapport * boues_flux(j, I, repere_Q)
                    boues_flux(boues_evacuees, I, repere_MES) = boues_flux(boues_evacuees, I, repere_MES) + rapport * boues_flux(j, I, repere_MES)
                ElseIf boues_flux(j, I, repere_verif_flux) < 0 Then
                    erreur = Feuil6.Range("erreur_boues_repartition").Cells(j, 1)
                    MsgBox erreur, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
                    pn_EffacementResultatsEnCours = True
                    warning = True
                    pn_ErreurNonGeree = True
                    Exit Sub
                End If
            Next j
        End If
    Next I

    'Graphiques sur les boues
    boues_totales_in_MS = 0
    boues_totales_out_MS = (cendres_tMSj + REFIB_tMSj) * 1000
    For I = 1 To nb_type_boues
        boues_totales_in_MS = boues_totales_in_MS + boues_flux(boues_inlet, I, repere_MES)
        boues_totales_out_MS = boues_totales_out_MS + boues_flux(boues_evacuees, I, repere_MES)
    Next I
    If boues_totales_in_MS < boues_totales_out_MS Then
        If boues_totales_in_MS <> 0 Then
            Feuil37.Range("graph_boues_reduction_" & pi_FiliereConsideree).Cells(1, 1) = 0
            Feuil37.Range("graph_boues_reduction_" & pi_FiliereConsideree).Cells(2, 1) = (boues_totales_out_MS) / (boues_totales_in_MS)
            Feuil37.Range("graph_boues_reduction_" & pi_FiliereConsideree).Cells(3, 1) = 2 - (boues_totales_out_MS) / (boues_totales_in_MS)
        Else
            Feuil37.Range("graph_boues_reduction_" & pi_FiliereConsideree).Cells(1, 1) = 0
            Feuil37.Range("graph_boues_reduction_" & pi_FiliereConsideree).Cells(2, 1) = 0
            Feuil37.Range("graph_boues_reduction_" & pi_FiliereConsideree).Cells(3, 1) = 1
        End If
    Else
        If boues_totales_in_MS <> 0 Then
            Feuil37.Range("graph_boues_reduction_" & pi_FiliereConsideree).Cells(1, 1) = (boues_totales_in_MS - boues_totales_out_MS) / (boues_totales_in_MS)
            Feuil37.Range("graph_boues_reduction_" & pi_FiliereConsideree).Cells(2, 1) = (boues_totales_out_MS) / (boues_totales_in_MS)
            Feuil37.Range("graph_boues_reduction_" & pi_FiliereConsideree).Cells(3, 1) = 1
        Else
            Feuil37.Range("graph_boues_reduction_" & pi_FiliereConsideree).Cells(1, 1) = 0
            Feuil37.Range("graph_boues_reduction_" & pi_FiliereConsideree).Cells(2, 1) = 0
            Feuil37.Range("graph_boues_reduction_" & pi_FiliereConsideree).Cells(3, 1) = 1
        End If
    End If


End Sub


Private Function F_verification_DBO_DCO(ByVal DBO As Double, ByVal DCO As Double) As Boolean

If DBO > DCO Then
    F_verification_DBO_DCO = False
Else
    F_verification_DBO_DCO = True
End If


End Function

Private Function F_verification_NK_NH4(ByVal NK As Double, ByVal NH4 As Double) As Boolean

If NH4 > NK Then
    F_verification_NK_NH4 = False
Else
    F_verification_NK_NH4 = True
End If

End Function




Private Function F_verification_humidite_air(ByVal humidite_air As Double) As Boolean

If humidite_air > 1 Or humidite_air < 0 Then
    F_verification_humidite_air = False
Else
    F_verification_humidite_air = True
End If

End Function


Private Function F_verification_altitude(ByVal altitude As Double) As Boolean

If altitude < 0 Then
    F_verification_altitude = False
Else
    F_verification_altitude = True
End If

End Function

