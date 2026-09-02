Attribute VB_Name = "z_CLS_BouesAthos"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'******************************************************************************************************************************************************
'ATHOS - Oxydation par voie Humide
'   Ce procédé comporte:
'       - une entrée de boues qui peut être un cocktail de deux boues de provenance différente
'       - une sortie de technosable (boues minéralisées)
'       - une sortie éventuelle d'eau au niveau de l'épaississeur (uniquement si l'épaississement des boues est nécessaire ou que l'utilisateur le souhaite)
'       - une sortie d'eau au niveau du décanteur (épaississement du technosable)
'   Ce module calcule la consommation électrique journalière du procédé Athos.
'   Ce module réalise les bilans matières sur le procédé et calcule les différents flux en sortie.
'   Il résoud les bilans d'énergie sur le procédé et extrait les informations nécessaire au fonctionnement du PINCH.
'   Il calcule par dichotomie la concentration des boues nécessaire en entrée pour atteindre une température du réacteur paramétrée par l'administrateur.
'   Il effectue un pré-dimensionnement du réacteur d'oxydation.
'   Il estime les consommations de produits chimiques et combustible:
'       - polymère,
'       - oxygène liquide,
'       - sulfate de cuivre CuSO4,
'       - gaz naturel ou biogaz en kWh(PCI)/j.
'******************************************************************************************************************************************************

Private Const mi_NOMBRE_FLUX_ENTREE_BOUES As Integer = 2

Private Const md_CRITERE_CONVERGENCE_SICCITE_BOUES_EPAISSIES As Double = 0.00000001
Private Const md_CRITERE_CONVERGENCE_BILAN_MATIERE As Double = 0.001

'constantes utilisées pour repérer la position des flux intermédiaires dans le tableau contenant la composition des boues/technosable/oxygène
Private Const mi_NOMBRE_FLUX_INTERMEDIAIRES As Integer = 11
Private Const mi_FLUX_ENTREE_BOUES As Integer = 0
Private Const mi_FLUX_EPAISSISSEMENT_MELANGE As Integer = 1
Private Const mi_FLUX_MELANGE_ECHANGEUR1 As Integer = 2
Private Const mi_FLUX_ECHANGEUR1_REACTEUR As Integer = 3     'on suppose que l'on injecte l'O2 et les boues directement dans le réacteur (ce qui en terme de calculs ne change pas grand chose
Private Const mi_FLUX_ALIMENTATION_OXYGENE As Integer = 4
Private Const mi_FLUX_SORTIE_FUMEES As Integer = 5
Private Const mi_FLUX_REACTEUR_ECHANGEUR2 As Integer = 6
Private Const mi_FLUX_ECHANGEUR2_ECHANGEUR3 As Integer = 7
Private Const mi_FLUX_ECHANGEUR3_DECANTEUR As Integer = 8
Private Const mi_FLUX_DECANTEUR_TECHNOSABLE As Integer = 9
Private Const mi_FLUX_RECIRCULATION_TECHNOSABLE As Integer = 10
Private Const mi_FLUX_SORTIE_TECHNOSABLE As Integer = 11

'constantes pour les paramètres relatifs aux flux intermédiaires
Private Const mi_NOMBRE_CARACTERISTIQUES_FLUX As Integer = 6
Private Const mi_CARACTERISTIQUE_MS_kgj As Integer = 1
Private Const mi_CARACTERISTIQUE_SICCITE As Integer = 2
Private Const mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj As Integer = 3
Private Const mi_CARACTERISTIQUE_RATIO_MV_MES As Integer = 4
Private Const mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C As Integer = 5
Private Const mi_CARACTERISTIQUE_TEMPERATURE_°C As Integer = 6

'valeur prise par le choix d'un flux d'entrée de boues qui n'a pas été utilisé
Private Const mi_FLUX_ENTREE_BOUES_NON_UTILISE As Integer = 0

'Constantes pour le choix du type de combustible utilisé dans le traitement des fumées RTO
Private Const mi_TRAITEMENT_FUMEES_RTO_AVEC_BIOGAZ As Integer = 0
Private Const mi_TRAITEMENT_FUMEES_RTO_AVEC_COMBUSTIBLE_EXTERNE As Integer = 1

'variables contenant les choix relatifs au procédé
Private mi_ChoixEntreesBoues(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Integer
Public pi_ChoixTypePolymereEpaississement As Integer
Private mi_ChoixTypeCombustibleTraitementFumeesRTO As Integer

'variables contenant les valeurs des paramètres qui sont affichés à l'utilisateur en fin de calcul
Private md_EntreeBouesPourcentageFluxAdmis(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Double
Private md_EntreeBouesMelangeMES_kgj As Double
Private md_EntreeBouesMelangeDebit_m3j As Double
Private md_EntreeBouesMelangeRatioMVMES As Double
Private md_EntreeBouesMelangeConcentrationMES_gL As Double
Private md_EntreeBouesMelangeTemperature_°C As Double
Private md_EpaississementTauxCaptureMES As Double
Private md_EpaississementRatioConsommationPolymerePur_kg_tMS As Double
Public pd_EpaississementConsommationPolymerePur_kgj As Double
Private md_BouesEpaissiesConcentrationMES_gL As Double
Private md_BouesEpaissiesMES_kgj As Double
Private md_BouesEpaissiesDebit_m3j As Double
Private md_ReacteurPressionAbsolue_bars As Double
Private md_ReacteurTemperature_°C As Double
Private md_ReacteurVolume_m3 As Double
Public pd_ConsommationOxygeneLiquidePur_kgj As Double
Public pd_ConsommationCuSO4Pur_kgj As Double
Private md_TraitementFumeesRTOConsommationCombustible_kWhPCIj As Double
Private md_DecanteurTauxCaptureMES As Double
Private md_DecanteurConcentrationMES_gL As Double
Private md_SortieTechnosableMES_kgj As Double
Private md_SortieTechnosableDebit_m3j As Double
Private md_SortieTechnosableRatioMVMES As Double
Public pd_ConsommationElectrique_kWhj As Double

'autres variables échangées avec d'autres modules de calcul
Public pd_TraitementFumeesRTOConsommationBiogaz_kWhPCIj As Double
Public pd_TraitementFumeesRTOConsommationCombustibleExterne_kWhPCIj As Double

'variables booléennes permettant de repérer si la valeur a été forcée par l'utilisateur (TRUE si renseignée, FALSE sinon)
Private mn_EntreeBouesMelangeTemperature_°C_force As Boolean
Private mn_EpaississementTauxCaptureMES_force As Boolean
Private mn_EpaississementRatioConsommationPolymerePur_kg_tMS_force As Boolean
Private mn_BouesEpaissiesConcentrationMES_gL_force As Boolean
Private mn_ReacteurVolume_m3_force As Boolean
Private mn_ConsommationOxygeneLiquidePur_kgj_force As Boolean
Private mn_ConsommationCuSO4Pur_kgj_force As Boolean
Private mn_TraitementFumeesRTOConsommationCombustible_kWhPCIj_force As Boolean
Private mn_DecanteurTauxCaptureMES_force As Boolean
Private mn_DecanteurConcentrationMES_gL_force As Boolean

'variables contenant les valeurs forcées par l'utilisateur
Private md_EntreeBouesMelangeTemperature_°C_f As Double
Private md_EpaississementTauxCaptureMES_f As Double
Private md_EpaississementRatioConsommationPolymerePur_kg_tMS_f As Double
Private md_BouesEpaissiesConcentrationMES_gL_f As Double
Private md_ReacteurVolume_m3_f As Double
Private md_ConsommationOxygeneLiquidePur_kgj_f As Double
Private md_ConsommationCuSO4Pur_kgj_f As Double
Private md_TraitementFumeesRTOConsommationCombustible_kWhPCIj_f As Double
Private md_DecanteurTauxCaptureMES_f As Double
Private md_DecanteurConcentrationMES_gL_f As Double

'variables intermédiaires de calcul
Private mi_ProvenanceBouesEntree(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Integer
Private mi_TypeBouesEntree(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Integer
Private md_FluxEauEvaporee_kgj As Double
Private md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX) As Double
Private mn_EpaississementBouesEntrees As Boolean    'repère si un épaississement a lieu ou non

'hypothèses utilisées dans les calculs
Private md_MasseVolumiqueBouesFroides_kg_m3 As Double
Private md_MasseVolumiqueTechnosable_kg_m3 As Double
Private md_SicciteMinimaleBouesEpaissies As Double
Private md_SicciteMaximaleBouesEpaissies As Double
Private md_TauxRecirculationTechnosable As Double
Private md_BouesOxydeesRatioMVparticulaireMVtotal As Double
Private md_ReacteurTauxReductionMV As Double
Private md_ReacteurTauxReductionDCO As Double
Private md_TemperatureTechnosableSortie_°C As Double
Private md_ReacteurRatioConsommationOxygene_kgO2_kgDCO As Double
Private md_MasseMolaireDioxydeCarbone_g_mol As Double
Private md_MasseMolaireOxygene_g_mol As Double
Private md_MasseMolaireEau_g_mol As Double
Private md_HypCapaciteCalorifiqueBouesHumidesFroides_kWh_kg°C As Double
Private md_HypCapaciteCalorifiqueBouesHumides110°C_kWh_kg°C As Double
Private md_HypCapaciteCalorifiqueBouesHumides150°C_kWh_kg°C As Double
Private md_HypCapaciteCalorifiqueBouesHumides250°C_kWh_kg°C As Double
Private md_HypCapaciteCalorifiqueOxygene15°C_kWh_kg°C As Double
Private md_HypCapaciteCalorifiqueOxygene250°C_kWh_kg°C As Double
Private md_HypCapaciteCalorifiqueDioxydeCarbone250°C_kWh_kg°C As Double
Private md_HypCapaciteCalorifiqueFumees_kWh_kg°C As Double
Private md_EfficaciteEchangeursChaleur As Double
Private md_PouvoirCalorifiqueInferieurMV_kWh_kg As Double
Private md_EnergieEchangeeEntreEch1etEch2_kWhj As Double
Private md_HypTemperatureOxygeneEntreeReacteur_°C As Double
Private md_ReacteurTemperatureCible_°C As Double
Private md_EnthalpieTotaleVapeur250°C_kWh_kg As Double
Private md_ReacteurTempsSejour_h As Double
Private md_HypEcartTemperatureBouesEchangeur2_°C As Double
Private md_MasseVolumiqueBoues250°C_kg_m3 As Double
Private md_RatioConsommationElectrique_kWh_m3BouesEpaissies As Double
Private md_MasseVolumiqueEau_kg_m3 As Double
Private md_RetoursDigestionRatioNH4surNKsoluble As Double
Private md_RetoursDigestionRatioDBOsolubleDCOsoluble As Double
Private md_RetoursDecanteurEtFiltrePlateauxRatioDCOtotaleDCOentree As Double
Private md_RetoursDecanteurEtFiltrePlateauxRatioDCOsolubleDCOtotale As Double
Private md_RetoursDecanteurEtFiltrePlateauxRatioNKtotalNKentree As Double
Private md_RetoursDecanteurEtFiltrePlateauxRatioNKsolubleNKtotal As Double
Private md_RetoursDecanteurEtFiltrePlateauxAvecDigestionRatioNH4surNKsoluble As Double
Private md_RetoursDecanteurEtFiltrePlateauxSansDigestionRatioNH4surNKsoluble As Double
Private md_RetoursDecanteurEtFiltrePlateauxAvecDigestionRatioPtotalPinlet As Double
Private md_RetoursDecanteurEtFiltrePlateauxSansDigestionRatioPtotalPinlet As Double
Private md_RetoursDecanteurEtFiltrePlateauxRatioPsolublePtotal As Double
Private md_ConcentrationPolymereInjecteEpaississement_gL As Double
Private md_HypEpaississementTauxCaptureMES As Double
Private md_ConsommationSpecifiqueCombustibleTraitementFumeesRTO_kWhPCI_tMVBouesEpaissies As Double

'Commentaires sur les calculs
Private Const mi_NOMBRE_COMMENTAIRES_CALCULS As Integer = 0
Private ms_CommentairesCalculs() As String

Private Sub attribution_Cp_Temperature()


'***********************************************************************************************************************************
'OBJECTIF
'   Cette procédure sert à attribuer les valeurs de capacité calorifique et de température
'   pour chaque flux intermédiaire interne au procédé

'VARIABLES D'ENTREE:
'   md_EntreeBouesMelangeTemperature_°C: température du mélange de boues en entrée
'   md_HypCapaciteCalorifiqueBouesHumidesFroides_kWh_kg°C: capacité calorifique des boues "froides" (entre 0 et 60°C)
'   md_HypCapaciteCalorifiqueBouesHumides110°C_kWh_kg°C: capacité calorifique des boues à 110°C
'   md_HypCapaciteCalorifiqueBouesHumides150°C_kWh_kg°C: capacité calorifique des boues à 150°C
'   md_HypCapaciteCalorifiqueBouesHumides250°C_kWh_kg°C: capacité calorifique des boues à 250°C
'   md_HypTemperatureOxygeneEntreeReacteur_°C: température de l'oxygène injecté dans le réacteur
'   md_HypCapaciteCalorifiqueOxygene15°C_kWh_kg°C: capacité calorifique de l'oxygène à 15°C
'   md_HypCapaciteCalorifiqueFumees_kWh_kg°C: capacité calorifique des fumées
'   md_TemperatureTechnosableSortie_°C: température du technosable en sortie

'VARIABLES DE SORTIE:
'   md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX): renseignement pour chaque flux
'       des caractéristiques de coordonnée mi_CARACTERISTIQUE_TEMPERATURE_°C et mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C
'***********************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree
    
'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "53.20"


'mi_FLUX_ENTREE_BOUES
md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_TEMPERATURE_°C) = md_EntreeBouesMelangeTemperature_°C
md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = md_HypCapaciteCalorifiqueBouesHumidesFroides_kWh_kg°C    'md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_SICCITE) * (hyp_Cp_boues_humides_kWh_kg°C - hyp_Cp_eau_kWh_kg°C) + hyp_Cp_eau_kWh_kg°C

'mi_FLUX_EPAISSISSEMENT_MELANGE
md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_TEMPERATURE_°C)
md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C)

'mi_FLUX_MELANGE_ECHANGEUR1
'défini plus tard dans la résolution des bilans matière et énergie

'mi_FLUX_ECHANGEUR1_REACTEUR
md_FluxIntermediaires(mi_FLUX_ECHANGEUR1_REACTEUR, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = md_HypCapaciteCalorifiqueBouesHumides150°C_kWh_kg°C
'la température sera calculé dans les bilans d'énergie

'mi_FLUX_ALIMENTATION_OXYGENE
md_FluxIntermediaires(mi_FLUX_ALIMENTATION_OXYGENE, mi_CARACTERISTIQUE_TEMPERATURE_°C) = md_HypTemperatureOxygeneEntreeReacteur_°C
md_FluxIntermediaires(mi_FLUX_ALIMENTATION_OXYGENE, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = md_HypCapaciteCalorifiqueOxygene15°C_kWh_kg°C

'mi_FLUX_SORTIE_FUMEES
md_FluxIntermediaires(mi_FLUX_SORTIE_FUMEES, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = md_HypCapaciteCalorifiqueFumees_kWh_kg°C  'calculée dans hypothèses à partir du rdt DCO
'la température sera calculée plus tard

'mi_FLUX_REACTEUR_ECHANGEUR2
md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = md_HypCapaciteCalorifiqueBouesHumides250°C_kWh_kg°C
'la température sera calculée plus tard

'mi_FLUX_ECHANGEUR2_ECHANGEUR3
md_FluxIntermediaires(mi_FLUX_ECHANGEUR2_ECHANGEUR3, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = md_HypCapaciteCalorifiqueBouesHumides110°C_kWh_kg°C
'la température sera calculée plus tard

'mi_FLUX_ECHANGEUR3_DECANTEUR
md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = md_HypCapaciteCalorifiqueBouesHumidesFroides_kWh_kg°C
md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_TEMPERATURE_°C) = md_TemperatureTechnosableSortie_°C

'mi_FLUX_DECANTEUR_TECHNOSABLE
md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C)
md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_TEMPERATURE_°C)

'mi_FLUX_RECIRCULATION_TECHNOSABLE
md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C)
md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_TEMPERATURE_°C)

'mi_FLUX_SORTIE_TECHNOSABLE
md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C)
md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_TEMPERATURE_°C)


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub


Private Sub calcul_boues_sortie_procede(ByVal BouesEpaissiesDCO_kgj As Double, ByVal BouesEpaissiesNK_kgj As Double, ByVal BouesEpaissiesP_kgj As Double, ByRef RatioBouesMinerales() As Double, ByRef FluxRetoursTotauxAthos() As Double, ByRef TableauRecapitulatifFluxBoues() As Double, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double)


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure permet de calculer la composition des boues en sortie du procédé.

'VARIABLES D'ENTREE:
'   BouesEpaissiesDCO_kgj: flux de DCO des boues épaissies en kg/j
'   BouesEpaissiesNK_kgj: flux de NK des boues épaissies en kg/j
'   BouesEpaissiesP_kgj: flux de phosphore des boues épaissies en kg/j
'   RatioBouesMinerales(): part de boues minérales contenu dans chaque type de boues en entrée
'   FluxRetoursTotauxAthos(): caractéristiques des retours en tête issus du décanteur
'   md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj): boues humides épaissies
'   md_MasseVolumiqueBouesFroides_kg_m3: masse volumique des boues "froides" (entre 0 et 60°C)

'VARIABLES DE SORTIE:
'   md_SortieTechnosableMES_kgj: Flux de MES du technosable en sortie du procédé
'   md_SortieTechnosableDebit_m3j: Débit de technosable en sortie du procédé
'   md_SortieTechnosableRatioMVMES: Part de matières volatiles dans les technosable en sortie de procédé
'   TableauRecapitulatifFluxBoues(): Récapitulatif des compositions de boues entre chaque procédé
'   TableauRecapitulatifPollutionSolubleDansFluxBoues(): Récapitulatif de la pollution soluble contenue dans les boues
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "53.23"
I = 0
j = 0


'sortie process
md_SortieTechnosableMES_kgj = md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj)
md_SortieTechnosableDebit_m3j = md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) / md_MasseVolumiqueBouesFroides_kg_m3
md_SortieTechnosableRatioMVMES = md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES)

'On ajoute la pollution soluble dans les boues
For j = 1 To nb_type_boues
    TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_athos, j, repere_mgL_DCO) = BouesEpaissiesDCO_kgj / md_SortieTechnosableDebit_m3j * 1000
    TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_athos, j, repere_mgL_NK) = BouesEpaissiesNK_kgj / md_SortieTechnosableDebit_m3j * 1000
    TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_athos, j, repere_mgL_Pt) = BouesEpaissiesP_kgj / md_SortieTechnosableDebit_m3j * 1000
Next j

'On remplit le tableau TableauRecapitulatifFluxBoues (on fait des proportions par rapport aux boues minérales en entrée
For j = 1 To nb_type_boues
    TableauRecapitulatifFluxBoues(boues_athos, j, repere_MES) = RatioBouesMinerales(j) * md_SortieTechnosableMES_kgj
    TableauRecapitulatifFluxBoues(boues_athos, j, repere_MV_MES) = md_SortieTechnosableRatioMVMES
    TableauRecapitulatifFluxBoues(boues_athos, j, repere_Q) = RatioBouesMinerales(j) * md_SortieTechnosableDebit_m3j
    TableauRecapitulatifFluxBoues(boues_athos, j, repere_ratio_DCO_MES) = 0
    TableauRecapitulatifFluxBoues(boues_athos, j, repere_ratio_DBO_MES) = 0
    TableauRecapitulatifFluxBoues(boues_athos, j, repere_ratio_NK_MES) = 0
    TableauRecapitulatifFluxBoues(boues_athos, j, repere_ratio_Pt_MES) = (BouesEpaissiesP_kgj - FluxRetoursTotauxAthos(repere_ret_Pt)) / md_SortieTechnosableMES_kgj
    For I = 1 To nb_repere_mgL
        TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_athos, j, I) = 0
    Next I
Next j


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Private Sub calcul_cocktail_boues_entrees(ByRef TableauRecapitulatifFluxBoues() As Double, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double, ByRef FluxPollutionSoluble_kgj() As Double, ByRef RatioBouesMinerales() As Double, ByRef EntreeBouesMelangeDCO_kgj As Double, ByRef EntreeBouesMelangeDBO_kgj As Double, ByRef EntreeBouesMelangeNK_kgj As Double, ByRef EntreeBouesMelangeP_kgj As Double, ByRef EntreeBouesRatioBouesDigerees As Double)


'************************************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure permet de mélanger les différents types de boues en entrée

'VARIABLES D'ENTREE:
'   TableauRecapitulatifFluxBoues(): contient toutes les informations sur les boues transitant entre les procédés
'   TableauRecapitulatifPollutionSolubleDansFluxBoues(): contient toutes les informations sur la pollution soluble contenue dans les boues transitant entre les procédés
'
'VARIABLES DE SORTIE:
'   FluxPollutionSoluble_kgj(): contient les flux de pollution soluble admis par le procédé
'   RatioBouesMinerales(): part de boues minérales en entrée de procédé pour chaque type de boues
'   EntreeBouesMelangeDCO_kgj: flux de DCO admis sur le procédé
'   EntreeBouesMelangeDBO_kgj: flux de DBO admis sur le procédé
'   EntreeBouesMelangeNK_kgj: flux de NK admis sur le procédé
'   EntreeBouesMelangeP_kgj: flux de P admis sur le procédé
'   md_EntreeBouesMelangeMES_kgj: flux de MES du mélange de boues
'   md_EntreeBouesMelangeDebit_m3j: débit volumique du mélange de boues
'   md_EntreeBouesMelangeRatioMVMES: rapport MV/MES des boues en entrée
'   EntreeBouesRatioBouesDigerees: part de boues qui proviennent de digestion
'   md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, 1 to mi_NOMBRE_CARACTERISTIQUES_FLUX): caractéristiques du flux d'entrée relatives au bilan masse
'************************************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer
Dim k As Integer
Dim IntermediaireMV_kgj As Double
Dim IntermediaireMES_kgj As Double
Dim IntermediaireNK_kgj As Double
Dim IntermediaireP_kgj As Double
Dim IntermediaireDCO_kgj As Double
Dim IntermediaireDBO_kgj As Double
Dim IntermediaireDebitVolumique_m3j As Double
Dim EntreeBouesMelangeMV_kgj As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "53.24"
I = 0
j = 0
k = 0
IntermediaireMV_kgj = 0
IntermediaireMES_kgj = 0
IntermediaireNK_kgj = 0
IntermediaireP_kgj = 0
IntermediaireDCO_kgj = 0
IntermediaireDBO_kgj = 0
IntermediaireDebitVolumique_m3j = 0
EntreeBouesMelangeMV_kgj = 0

'On initialise les variables déclarées au niveau du module
md_EntreeBouesMelangeMES_kgj = 0
md_EntreeBouesMelangeDebit_m3j = 0

For I = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    If mi_ChoixEntreesBoues(I) <> mi_FLUX_ENTREE_BOUES_NON_UTILISE Then
        If mi_ProvenanceBouesEntree(I) = boues_inlet Then
            'variable d'entrée
            IntermediaireMES_kgj = md_EntreeBouesPourcentageFluxAdmis(I) * TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_MES)
            IntermediaireMV_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_MV_MES) * IntermediaireMES_kgj
            IntermediaireDebitVolumique_m3j = md_EntreeBouesPourcentageFluxAdmis(I) * TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_Q)
            IntermediaireDCO_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_ratio_DCO_MES) * IntermediaireMES_kgj
            IntermediaireDBO_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_ratio_DBO_MES) * IntermediaireMES_kgj
            IntermediaireNK_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_ratio_NK_MES) * IntermediaireMES_kgj
            IntermediaireP_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_ratio_Pt_MES) * IntermediaireMES_kgj
            'on calcul l'entrée
            md_EntreeBouesMelangeMES_kgj = md_EntreeBouesMelangeMES_kgj + IntermediaireMES_kgj
            EntreeBouesMelangeMV_kgj = EntreeBouesMelangeMV_kgj + IntermediaireMV_kgj
            EntreeBouesMelangeDCO_kgj = EntreeBouesMelangeDCO_kgj + IntermediaireDCO_kgj
            EntreeBouesMelangeDBO_kgj = EntreeBouesMelangeDBO_kgj + IntermediaireDBO_kgj
            EntreeBouesMelangeNK_kgj = EntreeBouesMelangeNK_kgj + IntermediaireNK_kgj
            EntreeBouesMelangeP_kgj = EntreeBouesMelangeP_kgj + IntermediaireP_kgj
            md_EntreeBouesMelangeDebit_m3j = md_EntreeBouesMelangeDebit_m3j + IntermediaireDebitVolumique_m3j
            RatioBouesMinerales(mi_TypeBouesEntree(I)) = RatioBouesMinerales(mi_TypeBouesEntree(I)) + IntermediaireMES_kgj - IntermediaireMV_kgj

            'ENTREE SOLUBLE
            'TableauRecapitulatifPollutionSolubleDansFluxBoues(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
            For k = 1 To nb_repere_mgL
                FluxPollutionSoluble_kgj(k) = FluxPollutionSoluble_kgj(k) + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), k) * IntermediaireDebitVolumique_m3j / 1000
            Next k

            'pour la vérification des flux
            TableauRecapitulatifFluxBoues(boues_athos, mi_TypeBouesEntree(I), repere_flux_in) = TableauRecapitulatifFluxBoues(boues_athos, mi_TypeBouesEntree(I), repere_flux_in) + TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_flux_in) * md_EntreeBouesPourcentageFluxAdmis(I)
            TableauRecapitulatifFluxBoues(boues_athos, mi_TypeBouesEntree(I), repere_verif_flux) = TableauRecapitulatifFluxBoues(boues_athos, mi_TypeBouesEntree(I), repere_flux_in)
            'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
            TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_verif_flux) = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_verif_flux) - TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_flux_in) * md_EntreeBouesPourcentageFluxAdmis(I)
        Else
            For j = 1 To nb_type_boues
                'variables d'entrée
                IntermediaireMES_kgj = md_EntreeBouesPourcentageFluxAdmis(I) * TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_MES)
                IntermediaireMV_kgj = IntermediaireMES_kgj * TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_MV_MES)
                IntermediaireDebitVolumique_m3j = md_EntreeBouesPourcentageFluxAdmis(I) * TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_Q)
                IntermediaireDCO_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_ratio_DCO_MES) * IntermediaireMES_kgj
                IntermediaireDBO_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_ratio_DBO_MES) * IntermediaireMES_kgj
                IntermediaireNK_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_ratio_NK_MES) * IntermediaireMES_kgj
                IntermediaireP_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_ratio_Pt_MES) * IntermediaireMES_kgj
                'on calcule l'entrée
                md_EntreeBouesMelangeMES_kgj = md_EntreeBouesMelangeMES_kgj + IntermediaireMES_kgj
                EntreeBouesMelangeMV_kgj = EntreeBouesMelangeMV_kgj + IntermediaireMV_kgj
                md_EntreeBouesMelangeDebit_m3j = md_EntreeBouesMelangeDebit_m3j + IntermediaireDebitVolumique_m3j
                EntreeBouesMelangeDCO_kgj = EntreeBouesMelangeDCO_kgj + IntermediaireDCO_kgj
                EntreeBouesMelangeDBO_kgj = EntreeBouesMelangeDBO_kgj + IntermediaireDBO_kgj
                EntreeBouesMelangeNK_kgj = EntreeBouesMelangeNK_kgj + IntermediaireNK_kgj
                EntreeBouesMelangeP_kgj = EntreeBouesMelangeP_kgj + IntermediaireP_kgj
                RatioBouesMinerales(j) = RatioBouesMinerales(j) + IntermediaireMES_kgj - IntermediaireMV_kgj

                'ENTREE SOLUBLE
                'TableauRecapitulatifPollutionSolubleDansFluxBoues(nb_etape_file_boues, nb_type_boues, nb_repere_mgL)
                For k = 1 To nb_repere_mgL
                    FluxPollutionSoluble_kgj(k) = FluxPollutionSoluble_kgj(k) + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, k) * IntermediaireDebitVolumique_m3j / 1000
                Next k
                
                'On regarde si les boues ont été digérées
                If mi_ProvenanceBouesEntree(I) = boues_digerees Then
                    EntreeBouesRatioBouesDigerees = EntreeBouesRatioBouesDigerees + IntermediaireMES_kgj
                End If

                'je calcule les paramètres des boues épaissies
                TableauRecapitulatifFluxBoues(boues_athos, j, repere_flux_in) = TableauRecapitulatifFluxBoues(boues_athos, j, repere_flux_in) + TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_flux_in) * md_EntreeBouesPourcentageFluxAdmis(I)
                TableauRecapitulatifFluxBoues(boues_athos, j, repere_verif_flux) = TableauRecapitulatifFluxBoues(boues_athos, j, repere_flux_in)

                'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_verif_flux) = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_verif_flux) - TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_flux_in) * md_EntreeBouesPourcentageFluxAdmis(I)
            Next j
        End If
    End If
Next I

If md_EntreeBouesMelangeMES_kgj = 0 Or md_EntreeBouesMelangeDebit_m3j = 0 Then
    MsgBox Feuil5.Range("erreur_choix_athos_inlet"), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
    pn_EffacementResultatsEnCours = True
    pn_ErreurNonGeree = True
    Exit Sub
End If

'On calcule le ratio MV/MES et la concentration des boues en entrée
md_EntreeBouesMelangeConcentrationMES_gL = md_EntreeBouesMelangeMES_kgj / md_EntreeBouesMelangeDebit_m3j
md_EntreeBouesMelangeRatioMVMES = EntreeBouesMelangeMV_kgj / md_EntreeBouesMelangeMES_kgj
EntreeBouesRatioBouesDigerees = EntreeBouesRatioBouesDigerees / md_EntreeBouesMelangeMES_kgj

'On calcule la répartition des matières minérales par type de boues
For j = 1 To nb_type_boues
    RatioBouesMinerales(j) = RatioBouesMinerales(j) / (md_EntreeBouesMelangeMES_kgj - EntreeBouesMelangeMV_kgj)
Next j

'On renseigne le tableau contenant les données des flux internes au procédé
md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_MS_kgj) = md_EntreeBouesMelangeMES_kgj
md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) = md_EntreeBouesMelangeDebit_m3j * md_MasseVolumiqueBouesFroides_kg_m3
md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_SICCITE) = md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_MS_kgj) / md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_RATIO_MV_MES) = md_EntreeBouesMelangeRatioMVMES


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Private Sub calcul_composition_retours(ByVal EntreeBouesMelangeDCO_kgj As Double, ByVal EntreeBouesMelangeDBO_kgj As Double, ByVal EntreeBouesMelangeNK_kgj As Double, ByVal EntreeBouesMelangeP_kgj As Double, ByVal EntreeBouesRatioBouesDigerees As Double, ByRef FluxRetoursTotaux() As Double, ByRef FluxRetoursSoluble() As Double, ByRef FluxPollutionSoluble_kgj() As Double, ByRef BouesEpaissiesFluxPollutionSoluble_kgj() As Double, ByRef FluxRetoursTotauxDigestion() As Double, ByRef FluxRetoursSolubleDigestion() As Double, ByRef FluxRetoursTotauxAthos() As Double, ByRef FluxRetoursSolubleAthos() As Double)


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure permet d'estimer la composition des retours issus du procédé (sortie eau) et de les ajouter aux retours des autres procédés.

'VARIABLES D'ENTREE:
'   EntreeBouesMelangeDCO_kgj: flux de DCO des boues d'entrée
'   EntreeBouesMelangeDBO_kgj: flux de DBO des boues d'entrée
'   EntreeBouesMelangeNK_kgj: flux de NK des boues d'entrée
'   EntreeBouesMelangeP_kgj: flux de P des boues d'entrée
'   EntreeBouesRatioBouesDigerees: part de boues digérées dans les boues d'entrée
'   FluxPollutionSoluble_kgj(): pollution soluble contenue dans les boues en entrée
'   BouesEpaissiesFluxPollutionSoluble_kgj():pollution soluble contenue dans les boues épaissies
'   md_EpaississementTauxCaptureMES: taux de capture de l'épaississeur
'   pd_EpaississementConsommationPolymerePur_kgj: consommation de polymère pur
'   md_ConcentrationPolymereInjecteEpaississement_gL: concentration du polymère injecté (hypothèse de calcul)
'   md_RetoursDecanteurEtFiltrePlateauxRatioDCOtotaleDCOentree: ratio DCO retours / DCO entrée  (hypothèse de calcul)
'   md_RetoursDigestionRatioDBOsolubleDCOsoluble: part de DBO soluble / DCO soluble (hypothèse de calcul)
'   md_RetoursDigestionRatioNH4surNKsoluble: part de NH4 / NK soluble (hypothèse de calcul)

'VARIABLES DE SORTIE:
'   FluxRetoursTotaux(): somme de tous les retours de la file boues
'   FluxRetoursTotauxDigestion(): somme de tous les retours d'étapes de séparation situés juste en aval d'une étape de digestion
'   FluxRetoursSolubleDigestion(): somme de tous les retours solubles d'étapes de séparation situés juste en aval d'une étape de digestion
'   FluxRetoursTotauxAthos(): somme de tous les retours d'étapes de séparation situés en aval du réacteur d'Athos
'   FluxRetoursSolubleAthos(): somme de tous les retours d'étapes de séparation situés en aval du réacteur d'Athos
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim BouesEpaissiesDCO_kgj As Double
Dim BouesEpaissiesDBO_kgj As Double
Dim BouesEpaissiesNK_kgj As Double
Dim BouesEpaissiesP_kgj As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "53.25"
BouesEpaissiesDCO_kgj = 0
BouesEpaissiesDBO_kgj = 0
BouesEpaissiesNK_kgj = 0
BouesEpaissiesP_kgj = 0


'**************************************** RETOURS EPAISSISSEUR ********************************************************************
'on récupère les retours de l'épaississeur
If mn_EpaississementBouesEntrees = True Then
    'on commence par le particulaire
    FluxRetoursTotaux(repere_ret_Q) = FluxRetoursTotaux(repere_ret_Q) + (md_EntreeBouesMelangeDebit_m3j + pd_EpaississementConsommationPolymerePur_kgj / md_ConcentrationPolymereInjecteEpaississement_gL - md_BouesEpaissiesDebit_m3j)
    FluxRetoursTotaux(repere_ret_MES) = FluxRetoursTotaux(repere_ret_MES) + md_EntreeBouesMelangeMES_kgj * (1 - md_EpaississementTauxCaptureMES)
    FluxRetoursTotaux(repere_ret_DCO) = FluxRetoursTotaux(repere_ret_DCO) + EntreeBouesMelangeDCO_kgj * (1 - md_EpaississementTauxCaptureMES)
    FluxRetoursTotaux(repere_ret_DBO) = FluxRetoursTotaux(repere_ret_DBO) + EntreeBouesMelangeDBO_kgj * (1 - md_EpaississementTauxCaptureMES)
    FluxRetoursTotaux(repere_ret_NK) = FluxRetoursTotaux(repere_ret_NK) + EntreeBouesMelangeNK_kgj * (1 - md_EpaississementTauxCaptureMES)
    FluxRetoursTotaux(repere_ret_Pt) = FluxRetoursTotaux(repere_ret_Pt) + EntreeBouesMelangeP_kgj * (1 - md_EpaississementTauxCaptureMES)
    'on finit par le soluble
    FluxRetoursSoluble(repere_ret_DCO) = FluxRetoursSoluble(repere_ret_DCO) + FluxPollutionSoluble_kgj(repere_mgL_DCO) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO)
    FluxRetoursTotaux(repere_ret_DCO) = FluxRetoursTotaux(repere_ret_DCO) + FluxPollutionSoluble_kgj(repere_mgL_DCO) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO)
    FluxRetoursSoluble(repere_ret_DBO) = FluxRetoursSoluble(repere_ret_DBO) + md_RetoursDigestionRatioDBOsolubleDCOsoluble * (FluxPollutionSoluble_kgj(repere_mgL_DCO) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO))
    FluxRetoursTotaux(repere_ret_DBO) = FluxRetoursTotaux(repere_ret_DBO) + md_RetoursDigestionRatioDBOsolubleDCOsoluble * (FluxPollutionSoluble_kgj(repere_mgL_DCO) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO))
    FluxRetoursSoluble(repere_ret_NK) = FluxRetoursSoluble(repere_ret_NK) + FluxPollutionSoluble_kgj(repere_mgL_NK) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_NK)
    FluxRetoursTotaux(repere_ret_NK) = FluxRetoursTotaux(repere_ret_NK) + FluxPollutionSoluble_kgj(repere_mgL_NK) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_NK)
    FluxRetoursSoluble(repere_ret_NH4) = FluxRetoursSoluble(repere_ret_NH4) + md_RetoursDigestionRatioNH4surNKsoluble * (FluxPollutionSoluble_kgj(repere_mgL_NK) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_NK))
    FluxRetoursTotaux(repere_ret_NH4) = FluxRetoursTotaux(repere_ret_NH4) + md_RetoursDigestionRatioNH4surNKsoluble * (FluxPollutionSoluble_kgj(repere_mgL_NK) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_NK))
    FluxRetoursSoluble(repere_ret_Pt) = FluxRetoursSoluble(repere_ret_Pt) + FluxPollutionSoluble_kgj(repere_mgL_Pt) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_Pt)
    FluxRetoursTotaux(repere_ret_Pt) = FluxRetoursTotaux(repere_ret_Pt) + FluxPollutionSoluble_kgj(repere_mgL_Pt) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_Pt)
    'on récupère les retours de l'épaississeur dans ceux de digestion
    If EntreeBouesRatioBouesDigerees > 0 Then
        'on commence par le particulaire
        FluxRetoursTotauxDigestion(repere_ret_Q) = FluxRetoursTotauxDigestion(repere_ret_Q) + (md_EntreeBouesMelangeDebit_m3j + pd_EpaississementConsommationPolymerePur_kgj / md_ConcentrationPolymereInjecteEpaississement_gL - md_BouesEpaissiesDebit_m3j)
        FluxRetoursTotauxDigestion(repere_ret_MES) = FluxRetoursTotauxDigestion(repere_ret_MES) + md_EntreeBouesMelangeMES_kgj * (1 - md_EpaississementTauxCaptureMES)
        FluxRetoursTotauxDigestion(repere_ret_DCO) = FluxRetoursTotauxDigestion(repere_ret_DCO) + EntreeBouesMelangeDCO_kgj * (1 - md_EpaississementTauxCaptureMES)
        FluxRetoursTotauxDigestion(repere_ret_DBO) = FluxRetoursTotauxDigestion(repere_ret_DBO) + EntreeBouesMelangeDBO_kgj * (1 - md_EpaississementTauxCaptureMES)
        FluxRetoursTotauxDigestion(repere_ret_NK) = FluxRetoursTotauxDigestion(repere_ret_NK) + EntreeBouesMelangeNK_kgj * (1 - md_EpaississementTauxCaptureMES)
        FluxRetoursTotauxDigestion(repere_ret_Pt) = FluxRetoursTotauxDigestion(repere_ret_Pt) + EntreeBouesMelangeP_kgj * (1 - md_EpaississementTauxCaptureMES)
        'on finit par le soluble
        FluxRetoursSolubleDigestion(repere_ret_DCO) = FluxRetoursSolubleDigestion(repere_ret_DCO) + FluxPollutionSoluble_kgj(repere_mgL_DCO) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO)
        FluxRetoursTotauxDigestion(repere_ret_DCO) = FluxRetoursTotauxDigestion(repere_ret_DCO) + FluxPollutionSoluble_kgj(repere_mgL_DCO) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO)
        FluxRetoursSolubleDigestion(repere_ret_DBO) = FluxRetoursSolubleDigestion(repere_ret_DBO) + md_RetoursDigestionRatioDBOsolubleDCOsoluble * (FluxPollutionSoluble_kgj(repere_mgL_DCO) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO))
        FluxRetoursTotauxDigestion(repere_ret_DBO) = FluxRetoursTotauxDigestion(repere_ret_DBO) + md_RetoursDigestionRatioDBOsolubleDCOsoluble * (FluxPollutionSoluble_kgj(repere_mgL_DCO) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO))
        FluxRetoursSolubleDigestion(repere_ret_NK) = FluxRetoursSolubleDigestion(repere_ret_NK) + FluxPollutionSoluble_kgj(repere_mgL_NK) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_NK)
        FluxRetoursTotauxDigestion(repere_ret_NK) = FluxRetoursTotauxDigestion(repere_ret_NK) + FluxPollutionSoluble_kgj(repere_mgL_NK) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_NK)
        FluxRetoursSolubleDigestion(repere_ret_NH4) = FluxRetoursSolubleDigestion(repere_ret_NH4) + md_RetoursDigestionRatioNH4surNKsoluble * (FluxPollutionSoluble_kgj(repere_mgL_NK) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_NK))
        FluxRetoursTotauxDigestion(repere_ret_NH4) = FluxRetoursTotauxDigestion(repere_ret_NH4) + md_RetoursDigestionRatioNH4surNKsoluble * (FluxPollutionSoluble_kgj(repere_mgL_NK) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_NK))
        FluxRetoursSolubleDigestion(repere_ret_Pt) = FluxRetoursSolubleDigestion(repere_ret_Pt) + FluxPollutionSoluble_kgj(repere_mgL_Pt) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_Pt)
        FluxRetoursTotauxDigestion(repere_ret_Pt) = FluxRetoursTotauxDigestion(repere_ret_Pt) + FluxPollutionSoluble_kgj(repere_mgL_Pt) - BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_Pt)
    End If
End If


'******************************** POLLUTION APRES EPAISSISSEMENT ************************************************
'Voici les flux qui rentrent ensuite dans la boucle d'Athos (tout sera solubilisé) et sera réparti en concentration dans le décanteur
BouesEpaissiesDCO_kgj = EntreeBouesMelangeDCO_kgj * md_EpaississementTauxCaptureMES + BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO)
BouesEpaissiesDBO_kgj = EntreeBouesMelangeDBO_kgj * md_EpaississementTauxCaptureMES + md_RetoursDigestionRatioDBOsolubleDCOsoluble * BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO)
BouesEpaissiesNK_kgj = EntreeBouesMelangeNK_kgj * md_EpaississementTauxCaptureMES + BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_NK)
BouesEpaissiesP_kgj = EntreeBouesMelangeP_kgj * md_EpaississementTauxCaptureMES + BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_Pt)


'********************************* RETOURS DU DECANTEUR ***************************************
FluxRetoursTotauxAthos(repere_ret_Q) = (md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) - md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)) / md_MasseVolumiqueBouesFroides_kg_m3
    FluxRetoursTotaux(repere_ret_Q) = FluxRetoursTotaux(repere_ret_Q) + FluxRetoursTotauxAthos(repere_ret_Q)
FluxRetoursTotauxAthos(repere_ret_DCO) = md_RetoursDecanteurEtFiltrePlateauxRatioDCOtotaleDCOentree * BouesEpaissiesDCO_kgj
    FluxRetoursSolubleAthos(repere_ret_DCO) = FluxRetoursTotauxAthos(repere_ret_DCO) * md_RetoursDecanteurEtFiltrePlateauxRatioDCOsolubleDCOtotale
    FluxRetoursTotaux(repere_ret_DCO) = FluxRetoursTotaux(repere_ret_DCO) + FluxRetoursTotauxAthos(repere_ret_DCO)
    FluxRetoursSoluble(repere_ret_DCO) = FluxRetoursSoluble(repere_ret_DCO) + FluxRetoursSolubleAthos(repere_ret_DCO)
FluxRetoursTotauxAthos(repere_ret_DBO) = FluxRetoursTotauxAthos(repere_ret_DCO)
    FluxRetoursSolubleAthos(repere_ret_DBO) = FluxRetoursSolubleAthos(repere_ret_DCO)
    FluxRetoursTotaux(repere_ret_DBO) = FluxRetoursTotaux(repere_ret_DBO) + FluxRetoursTotauxAthos(repere_ret_DBO)
    FluxRetoursSoluble(repere_ret_DBO) = FluxRetoursSoluble(repere_ret_DBO) + FluxRetoursSolubleAthos(repere_ret_DBO)
FluxRetoursTotauxAthos(repere_ret_NK) = md_RetoursDecanteurEtFiltrePlateauxRatioNKtotalNKentree * BouesEpaissiesNK_kgj
    FluxRetoursSolubleAthos(repere_ret_NK) = FluxRetoursTotauxAthos(repere_ret_NK) * md_RetoursDecanteurEtFiltrePlateauxRatioNKsolubleNKtotal
    If EntreeBouesRatioBouesDigerees > 0 Then
        FluxRetoursSolubleAthos(repere_ret_NH4) = FluxRetoursSolubleAthos(repere_ret_NK) * md_RetoursDecanteurEtFiltrePlateauxAvecDigestionRatioNH4surNKsoluble
    Else
        FluxRetoursSolubleAthos(repere_ret_NH4) = FluxRetoursSolubleAthos(repere_ret_NK) * md_RetoursDecanteurEtFiltrePlateauxSansDigestionRatioNH4surNKsoluble
    End If
    FluxRetoursTotauxAthos(repere_ret_NH4) = FluxRetoursSolubleAthos(repere_ret_NH4)
    FluxRetoursTotaux(repere_ret_NK) = FluxRetoursTotaux(repere_ret_NK) + FluxRetoursTotauxAthos(repere_ret_NK)
    FluxRetoursSoluble(repere_ret_NK) = FluxRetoursSoluble(repere_ret_NK) + FluxRetoursSolubleAthos(repere_ret_NK)
    FluxRetoursTotaux(repere_ret_NH4) = FluxRetoursTotaux(repere_ret_NH4) + FluxRetoursSolubleAthos(repere_ret_NH4)
    FluxRetoursSoluble(repere_ret_NH4) = FluxRetoursSoluble(repere_ret_NH4) + FluxRetoursSolubleAthos(repere_ret_NH4)
FluxRetoursTotauxAthos(repere_ret_Pt) = md_RetoursDecanteurEtFiltrePlateauxAvecDigestionRatioPtotalPinlet * BouesEpaissiesP_kgj
    FluxRetoursSolubleAthos(repere_ret_Pt) = md_RetoursDecanteurEtFiltrePlateauxRatioPsolublePtotal * BouesEpaissiesP_kgj
    FluxRetoursTotaux(repere_ret_Pt) = FluxRetoursTotaux(repere_ret_Pt) + FluxRetoursTotauxAthos(repere_ret_Pt)
    FluxRetoursSoluble(repere_ret_Pt) = FluxRetoursSoluble(repere_ret_Pt) + FluxRetoursSolubleAthos(repere_ret_Pt)


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Private Sub calcul_siccite_boues_epaissies(ByVal EntreeBouesMelangeDCO_kgj As Double, ByRef FluxPollutionSoluble_kgj() As Double, ByRef BouesEpaissiesFluxPollutionSoluble_kgj() As Double, ByRef BouesEpaissiesDCO_kgj As Double, ByRef BouesEpaissiesDCOparticulaire_kgj As Double)


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure calcule la siccité à avoir après l'épaississement des boues pour que la température dans le réacteur soit égale
'   ou le plus proche possible de la valeur souhaitée "md_ReacteurTemperatureCible_°C".
'   Si la concentration des boues admises n'est pas trop élevée (<"md_SicciteMaximaleBouesEpaissies"), une dichotomie est réalisée.
'   Lors de cette dichotomie, les procédures suivantes sont appelées:
'       - "calcul_epaississement": pour simuler l'étape d'épaississement
'       - "calcul_bilan_masse": pour résoudre le bilan massique sur le procédé
'       - "calcul_bilan_energie": pour résoudre les bilans d'énergie et estimer entre autres la température du réacteur

'VARIABLES D'ENTREE:
'   EntreeBouesMelangeDCO_kgj: DCO des boues d'entrée
'   FluxPollutionSoluble_kgj(): pollution soluble des boues en entrée
'   md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX):
'       certaines données (flux d'entrée, et d'autres caractéristiques telles que les capacités calorifiques voire certaines températures)
'   mn_BouesEpaissiesConcentrationMES_gL_force: repère si l'utilisateur a forcé la concentration des boues épaissies
'       md_BouesEpaissiesConcentrationMES_gL_f: valeur renseignée par l'utilisateur
'   mn_EpaississementTauxCaptureMES_force: repère si l'utilisateur a forcé le taux de capture de l'épaississement
'       md_EpaississementTauxCaptureMES_f: valeur renseignée par l'utilisateur
'   md_MasseVolumiqueBouesFroides_kg_m3: masse volumique des boues à basse température (0°C - 60°C) (hypothèse de calcul fixée)
'   md_ReacteurTemperatureCible_°C: température de réacteur souhaitée (hypothèse de calcul fixée)
'   md_SicciteMinimaleBouesEpaissies: valeur minimale de concentration des boues épaissies
'   md_SicciteMaximaleBouesEpaissies: valeur maximale de concentration des boues épaissies
'   md_ReacteurTemperature_°C: température de fonctionnement du réacteur

'VARIABLES DE SORTIE:
'   BouesEpaissiesFluxPollutionSoluble_kgj(): pollution soluble des boues après épaississement
'   BouesEpaissiesDCO_kgj: flux de DCO des boues épaissies
'   BouesEpaissiesDCOparticulaire_kgj: flux de DCO particulaire des boues épaissies
'   md_BouesEpaissiesConcentrationMES_gL: concentration des boues en sortie d'épaississement
'   md_BouesEpaissiesMES_kgj: débit massique de boues épaissies
'   md_BouesEpaissiesDebit_m3j: débit volumique de boues épaissies
'   mn_EpaississementBouesEntrees: repère si un épaississement des boues d'entrée a lieu ou non
'   md_EpaississementTauxCaptureMES: taux de capture des MES au niveau de l'épaississement
'   md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX): la majorité des données
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim CritereConvergenceDichotomieSiccite As Double
Dim DichotomieSicciteConsideree As Double
Dim DichotomieSicciteMaximale As Double
Dim DichotomieSicciteMinimale As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "53.28"
I = 0
CritereConvergenceDichotomieSiccite = 0
DichotomieSicciteConsideree = 0
DichotomieSicciteMaximale = 0
DichotomieSicciteMinimale = 0

'on fait la boucle pour recalculer la concentration de l'épaississement
If mn_BouesEpaissiesConcentrationMES_gL_force = True Then
    md_BouesEpaissiesConcentrationMES_gL = md_BouesEpaissiesConcentrationMES_gL_f
    'on a un épaississement
    mn_EpaississementBouesEntrees = True
    'On calcule la siccité
    md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_SICCITE) = md_BouesEpaissiesConcentrationMES_gL / md_MasseVolumiqueBouesFroides_kg_m3
    'on calcule l'épaississement
    Call calcul_epaississement(BouesEpaissiesDCOparticulaire_kgj, EntreeBouesMelangeDCO_kgj)
    'on calcule la pollution soluble après l'épaississement
    For I = 1 To nb_repere_mgL
        BouesEpaissiesFluxPollutionSoluble_kgj(I) = FluxPollutionSoluble_kgj(I) * (md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) / md_MasseVolumiqueBouesFroides_kg_m3) / md_EntreeBouesMelangeDebit_m3j
    Next I
    BouesEpaissiesDCO_kgj = BouesEpaissiesDCOparticulaire_kgj + BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO)
    'on calcule le bilan masse
    Call calcul_bilan_masse(BouesEpaissiesDCO_kgj)
    'on calcule le bilan d'énergie
    Call calcul_bilan_energie
Else
    'on cherche la siccité à atteindre par une dichotomie dans le cas où l'on peut souhaiter augmenter la siccité
    If md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_SICCITE) > md_SicciteMaximaleBouesEpaissies Then
        'épaississement inutile
        mn_EpaississementBouesEntrees = False
        'on affecte les valeurs pour inhiber l'effet de l'épaississement (siccité=entrée et capture de 100%)
        md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_SICCITE) = md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_SICCITE)
        md_EpaississementTauxCaptureMES = 1
        Call calcul_epaississement(BouesEpaissiesDCOparticulaire_kgj, EntreeBouesMelangeDCO_kgj)
        'on calcule la pollution soluble après l'épaississement
        For I = 1 To nb_repere_mgL
            BouesEpaissiesFluxPollutionSoluble_kgj(I) = FluxPollutionSoluble_kgj(I) * (md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) / md_MasseVolumiqueBouesFroides_kg_m3) / md_EntreeBouesMelangeDebit_m3j
        Next I
        BouesEpaissiesDCO_kgj = BouesEpaissiesDCOparticulaire_kgj + BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO)
        'on calcule le bilan masse
        Call calcul_bilan_masse(BouesEpaissiesDCO_kgj)
        'on calcule le bilan d'énergie
        Call calcul_bilan_energie
    Else
        'on définit les bornes de la dichotomie
        If md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_SICCITE) > md_SicciteMinimaleBouesEpaissies Then
            DichotomieSicciteMinimale = md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_SICCITE)
        Else
            DichotomieSicciteMinimale = md_SicciteMinimaleBouesEpaissies
        End If
        'ON TESTE UNE 1ERE FOIS A LA SICCITE MINIMUM POUR VOIR SI DEJA AUTOTHERMIQUE
        'épaississement inutile
        mn_EpaississementBouesEntrees = False
        'on affecte les valeurs pour inhiber l'effet de l'épaississement (siccité=entrée et capture de 100%)
        md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_SICCITE) = md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_SICCITE)
        md_EpaississementTauxCaptureMES = 1
        Call calcul_epaississement(BouesEpaissiesDCOparticulaire_kgj, EntreeBouesMelangeDCO_kgj)
        'on calcule la pollution soluble après l'épaississement
        For I = 1 To nb_repere_mgL
            BouesEpaissiesFluxPollutionSoluble_kgj(I) = FluxPollutionSoluble_kgj(I) * (md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) / md_MasseVolumiqueBouesFroides_kg_m3) / md_EntreeBouesMelangeDebit_m3j
        Next I
        BouesEpaissiesDCO_kgj = BouesEpaissiesDCOparticulaire_kgj + BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO)
        'on calcule le bilan masse
        Call calcul_bilan_masse(BouesEpaissiesDCO_kgj)
        'on calcule le bilan d'énergie
        Call calcul_bilan_energie
        If md_ReacteurTemperature_°C < md_ReacteurTemperatureCible_°C Then   'dans le cas où l'on n'est pas autothermique
            'épaississement potentiellement nécessaire donc on en mets un
            mn_EpaississementBouesEntrees = True
            'taux de capture MES épaississement
            If mn_EpaississementTauxCaptureMES_force = True Then
                md_EpaississementTauxCaptureMES = md_EpaississementTauxCaptureMES_f
            Else
                md_EpaississementTauxCaptureMES = md_HypEpaississementTauxCaptureMES
            End If
            'on définit la borne max
            DichotomieSicciteMaximale = md_SicciteMaximaleBouesEpaissies
            DichotomieSicciteConsideree = (DichotomieSicciteMinimale + DichotomieSicciteMaximale) / 2
            CritereConvergenceDichotomieSiccite = md_CRITERE_CONVERGENCE_SICCITE_BOUES_EPAISSIES + 1
            'on attaque la boucle de dichotomie
            Do While CritereConvergenceDichotomieSiccite > md_CRITERE_CONVERGENCE_SICCITE_BOUES_EPAISSIES
                'on utilise la valeur de DichotomieSicciteConsideree
                md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_SICCITE) = DichotomieSicciteConsideree
                'on calcule la qualité des boues épaissies
                Call calcul_epaississement(BouesEpaissiesDCOparticulaire_kgj, EntreeBouesMelangeDCO_kgj)
                'on calcule la pollution soluble après l'épaississement
                For I = 1 To nb_repere_mgL
                    BouesEpaissiesFluxPollutionSoluble_kgj(I) = FluxPollutionSoluble_kgj(I) * (md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) / md_MasseVolumiqueBouesFroides_kg_m3) / md_EntreeBouesMelangeDebit_m3j
                Next I
                BouesEpaissiesDCO_kgj = BouesEpaissiesDCOparticulaire_kgj + BouesEpaissiesFluxPollutionSoluble_kgj(repere_mgL_DCO)
                'on calcule le bilan masse
                Call calcul_bilan_masse(BouesEpaissiesDCO_kgj)
                'on calcule le bilan d'énergie
                Call calcul_bilan_energie    'Call calcul_BE_2   '
                'on redéfinit les bornes
                If md_ReacteurTemperature_°C < md_ReacteurTemperatureCible_°C Then
                    DichotomieSicciteMinimale = DichotomieSicciteConsideree
                ElseIf md_ReacteurTemperature_°C > md_ReacteurTemperatureCible_°C Then
                    DichotomieSicciteMaximale = DichotomieSicciteConsideree
                Else
                    DichotomieSicciteMinimale = DichotomieSicciteConsideree
                    DichotomieSicciteMaximale = DichotomieSicciteConsideree
                End If
                DichotomieSicciteConsideree = (DichotomieSicciteMinimale + DichotomieSicciteMaximale) / 2
                CritereConvergenceDichotomieSiccite = (DichotomieSicciteMaximale - DichotomieSicciteMinimale) / DichotomieSicciteConsideree
            Loop
            'on utilise la dernière valeur de DichotomieSicciteConsideree (médiane du dernier intervalle de recherche)
            md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_SICCITE) = DichotomieSicciteConsideree
        End If
    End If
End If

'on récupère les valeurs en sortie d'épaississement
md_BouesEpaissiesConcentrationMES_gL = md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_SICCITE) * md_MasseVolumiqueBouesFroides_kg_m3
md_BouesEpaissiesMES_kgj = md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_MS_kgj)
md_BouesEpaissiesDebit_m3j = md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) / md_MasseVolumiqueBouesFroides_kg_m3


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Public Sub initialisation_procede()


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure a pour but d'initialiser le procédé, i.e. de lire les entrées de l'utilisateur et d'initialiser
'   toutes les variables et hypothèses qui sont utilisées dans les calculs mais qui ne sont pas elles-même calculées.
'   Elle lance les subroutines suivantes: "lecture_choix", "lecture_valeurs_forcees", "attribution_hypotheses", "attribution_valeur_par_defaut"

'VARIABLES D'ENTREE:

'VARIABLES DE SORTIE:
'******************************************************************************************************************************************************

'On lit les choix de l'utilisateur concernant le procédé
Call lecture_choix
'Pour le calcul des valeurs effectives, on lit les valeurs forcées par l'utilisateur
If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
    Call lecture_valeurs_forcees
End If
'On définit l'ensemble des variables techniques qui sont utilisées dans les calculs mais non modifiables par l'utilisateur
Call attribution_hypotheses
'On définit les variables modifiables par l'utilisateur qui sont utilisées dans les calculs mais non calculées
Call attribution_valeur_par_defaut


End Sub

Private Sub initialisation_variables_module()


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure initialise toutes les variables déclarées au niveau du module

'VARIABLES D'ENTREE:

'VARIABLES DE SORTIE:
'   toutes les variables déclarées au niveau du module de classe
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "53.0"
I = 0
j = 0

'variables contenant les choix relatifs au procédé
For I = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    mi_ChoixEntreesBoues(I) = 0
Next I
pi_ChoixTypePolymereEpaississement = 0
mi_ChoixTypeCombustibleTraitementFumeesRTO = 0

'variables contenant les valeurs des paramètres qui sont affichés à l'utilisateur en fin de calcul
For I = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    md_EntreeBouesPourcentageFluxAdmis(I) = 0
Next I
md_EntreeBouesMelangeMES_kgj = 0
md_EntreeBouesMelangeDebit_m3j = 0
md_EntreeBouesMelangeRatioMVMES = 0
md_EntreeBouesMelangeConcentrationMES_gL = 0
md_EntreeBouesMelangeTemperature_°C = 0
md_EpaississementTauxCaptureMES = 0
md_EpaississementRatioConsommationPolymerePur_kg_tMS = 0
pd_EpaississementConsommationPolymerePur_kgj = 0
md_BouesEpaissiesConcentrationMES_gL = 0
md_BouesEpaissiesMES_kgj = 0
md_BouesEpaissiesDebit_m3j = 0
md_ReacteurPressionAbsolue_bars = 0
md_ReacteurTemperature_°C = 0
md_ReacteurVolume_m3 = 0
pd_ConsommationOxygeneLiquidePur_kgj = 0
pd_ConsommationCuSO4Pur_kgj = 0
md_TraitementFumeesRTOConsommationCombustible_kWhPCIj = 0
md_DecanteurTauxCaptureMES = 0
md_DecanteurConcentrationMES_gL = 0
md_SortieTechnosableMES_kgj = 0
md_SortieTechnosableDebit_m3j = 0
md_SortieTechnosableRatioMVMES = 0
pd_ConsommationElectrique_kWhj = 0

'autres variables échangées avec d'autres modules de calcul
pd_TraitementFumeesRTOConsommationBiogaz_kWhPCIj = 0
pd_TraitementFumeesRTOConsommationCombustibleExterne_kWhPCIj = 0

'variables booléennes permettant de repérer si la valeur a été forcée par l'utilisateur (TRUE si renseignée, FALSE sinon)
mn_EntreeBouesMelangeTemperature_°C_force = False
mn_EpaississementTauxCaptureMES_force = False
mn_EpaississementRatioConsommationPolymerePur_kg_tMS_force = False
mn_BouesEpaissiesConcentrationMES_gL_force = False
mn_ReacteurVolume_m3_force = False
mn_ConsommationOxygeneLiquidePur_kgj_force = False
mn_ConsommationCuSO4Pur_kgj_force = False
mn_TraitementFumeesRTOConsommationCombustible_kWhPCIj_force = False
mn_DecanteurTauxCaptureMES_force = False
mn_DecanteurConcentrationMES_gL_force = False

'variables contenant les valeurs forcées par l'utilisateur
md_EntreeBouesMelangeTemperature_°C_f = 0
md_EpaississementTauxCaptureMES_f = 0
md_EpaississementRatioConsommationPolymerePur_kg_tMS_f = 0
md_BouesEpaissiesConcentrationMES_gL_f = 0
md_ReacteurVolume_m3_f = 0
md_ConsommationOxygeneLiquidePur_kgj_f = 0
md_ConsommationCuSO4Pur_kgj_f = 0
md_TraitementFumeesRTOConsommationCombustible_kWhPCIj_f = 0
md_DecanteurTauxCaptureMES_f = 0
md_DecanteurConcentrationMES_gL_f = 0

'variables intermédiaires de calcul
For I = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    mi_ProvenanceBouesEntree(I) = 0
    mi_TypeBouesEntree(I) = 0
Next I
md_FluxEauEvaporee_kgj = 0
For I = 0 To mi_NOMBRE_FLUX_INTERMEDIAIRES
    For j = 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX
        md_FluxIntermediaires(I, j) = 0
    Next j
Next I
mn_EpaississementBouesEntrees = False   'repère si un épaississement a lieu ou non

'hypothèses utilisées dans les calculs
md_MasseVolumiqueBouesFroides_kg_m3 = 0
md_MasseVolumiqueTechnosable_kg_m3 = 0
md_SicciteMinimaleBouesEpaissies = 0
md_SicciteMaximaleBouesEpaissies = 0
md_TauxRecirculationTechnosable = 0
md_BouesOxydeesRatioMVparticulaireMVtotal = 0
md_ReacteurTauxReductionMV = 0
md_ReacteurTauxReductionDCO = 0
md_TemperatureTechnosableSortie_°C = 0
md_ReacteurRatioConsommationOxygene_kgO2_kgDCO = 0
md_MasseMolaireDioxydeCarbone_g_mol = 0
md_MasseMolaireOxygene_g_mol = 0
md_MasseMolaireEau_g_mol = 0
md_HypCapaciteCalorifiqueBouesHumidesFroides_kWh_kg°C = 0
md_HypCapaciteCalorifiqueBouesHumides110°C_kWh_kg°C = 0
md_HypCapaciteCalorifiqueBouesHumides150°C_kWh_kg°C = 0
md_HypCapaciteCalorifiqueBouesHumides250°C_kWh_kg°C = 0
md_HypCapaciteCalorifiqueOxygene15°C_kWh_kg°C = 0
md_HypCapaciteCalorifiqueOxygene250°C_kWh_kg°C = 0
md_HypCapaciteCalorifiqueDioxydeCarbone250°C_kWh_kg°C = 0
md_HypCapaciteCalorifiqueFumees_kWh_kg°C = 0
md_EfficaciteEchangeursChaleur = 0
md_PouvoirCalorifiqueInferieurMV_kWh_kg = 0
md_EnergieEchangeeEntreEch1etEch2_kWhj = 0
md_HypTemperatureOxygeneEntreeReacteur_°C = 0
md_ReacteurTemperatureCible_°C = 0
md_EnthalpieTotaleVapeur250°C_kWh_kg = 0
md_ReacteurTempsSejour_h = 0
md_HypEcartTemperatureBouesEchangeur2_°C = 0
md_MasseVolumiqueBoues250°C_kg_m3 = 0
md_RatioConsommationElectrique_kWh_m3BouesEpaissies = 0
md_MasseVolumiqueEau_kg_m3 = 0
md_RetoursDigestionRatioNH4surNKsoluble = 0
md_RetoursDigestionRatioDBOsolubleDCOsoluble = 0
md_RetoursDecanteurEtFiltrePlateauxRatioDCOtotaleDCOentree = 0
md_RetoursDecanteurEtFiltrePlateauxRatioDCOsolubleDCOtotale = 0
md_RetoursDecanteurEtFiltrePlateauxRatioNKtotalNKentree = 0
md_RetoursDecanteurEtFiltrePlateauxRatioNKsolubleNKtotal = 0
md_RetoursDecanteurEtFiltrePlateauxAvecDigestionRatioNH4surNKsoluble = 0
md_RetoursDecanteurEtFiltrePlateauxSansDigestionRatioNH4surNKsoluble = 0
md_RetoursDecanteurEtFiltrePlateauxAvecDigestionRatioPtotalPinlet = 0
md_RetoursDecanteurEtFiltrePlateauxSansDigestionRatioPtotalPinlet = 0
md_RetoursDecanteurEtFiltrePlateauxRatioPsolublePtotal = 0
md_ConcentrationPolymereInjecteEpaississement_gL = 0
md_HypEpaississementTauxCaptureMES = 0
md_ConsommationSpecifiqueCombustibleTraitementFumeesRTO_kWhPCI_tMVBouesEpaissies = 0

'erreurs
If mi_NOMBRE_COMMENTAIRES_CALCULS > 0 Then
    ReDim ms_CommentairesCalculs(1 To mi_NOMBRE_COMMENTAIRES_CALCULS)
    For I = 1 To mi_NOMBRE_COMMENTAIRES_CALCULS
        ms_CommentairesCalculs(I) = ""
    Next I
Else
    ReDim ms_CommentairesCalculs(1 To 1)
    ms_CommentairesCalculs(1) = ""
End If

Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Private Sub lecture_choix()


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure lit les choix réalisés par l'utilisateur concernant le procédé dans l'onglet "Synthèse".
'   Elle lit également les variables renseignées dans l'onglet "Synthèse" (variables auxquelles il est impossible d'attribuer une valeur par défaut)

'VARIABLES D'ENTREE:
'   entrées de l'utilisateur dans l'onglet "Synthèse" sous forme de choix dans des listes de validation (~95% des entrées) ou de valeurs saisies.
'       Les valeurs (entiers) associées aux choix se trouvent dans l'onglet "Feuil53" nommé "Athos" où sont également reportées les valeurs saisies
'       et où la subroutine va lire les données d'entrée obligatoires utilisateur dans des plages de cellules nommées dans Excel

'VARIABLES DE SORTIE:
'   mi_ChoixEntreesBoues(1 To mi_NOMBRE_FLUX_ENTREE_BOUES): provenance des boues alimentant le procédé
'   md_EntreeBouesPourcentageFluxAdmis(1 To mi_NOMBRE_FLUX_ENTREE_BOUES): pourcentage admis du flux choisi pour alimenter le procédé
'   mi_ProvenanceBouesEntree(1 To mi_NOMBRE_FLUX_ENTREE_BOUES): similaire à mi_ChoixEntreesBoues(1 To mi_NOMBRE_FLUX_ENTREE_BOUES)
'       mais avec un décalage dans les indices car il considère toutes les provenances possible de la file eau comme une seule (boues brutes)
'   mi_TypeBouesEntree(1 To mi_NOMBRE_FLUX_ENTREE_BOUES): si la boue est brute (provenance directe de la file de traitement de l'eau)
'       elle donne le type de traitement de la file eau d'où provient la boue
'   pi_ChoixTypePolymereEpaississement: type de polymère utilisé
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "53.1"
I = 0


'Boues en entrée d'ATHOS
For I = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    mi_ChoixEntreesBoues(I) = Feuil53.Range("athos_choix_inlet").Cells(I, pi_FiliereConsideree)
    If mi_ChoixEntreesBoues(I) <> mi_FLUX_ENTREE_BOUES_NON_UTILISE Then
        md_EntreeBouesPourcentageFluxAdmis(I) = Feuil53.Range("athos_inlet_ratio").Cells(I, pi_FiliereConsideree)
        If mi_ChoixEntreesBoues(I) <= graisses Then
            mi_ProvenanceBouesEntree(I) = boues_inlet
            mi_TypeBouesEntree(I) = mi_ChoixEntreesBoues(I)
        Else
            mi_ProvenanceBouesEntree(I) = (mi_ChoixEntreesBoues(I) - graisses) + 1
            'ici on prendra directement les valeurs du flux antérieur
        End If
    End If
Next I

'Type de polymère utilisé pour l'épaississement
 pi_ChoixTypePolymereEpaississement = Feuil53.Range("athos_choix_epaississement_polymere_type").Cells(1, pi_FiliereConsideree)

'Type de combustible utilisé pour l'épaississement
mi_ChoixTypeCombustibleTraitementFumeesRTO = Feuil53.Range("athos_choix_combustible_traitement_RTO").Cells(1, pi_FiliereConsideree)


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Private Sub attribution_hypotheses()


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure attribue des valeurs à des variables utilisées dans les calculs qui ne sont ni affichées ni modifiables par l'utilisateur

'VARIABLES D'ENTREE:
'   Certaines hypothèses peuvent éventuellement dépendre de paramètres renseignés par l'utilisateur (au niveau du procédé, dans un autre procédé, ou
'   dans les données générales de la filière simulée

'VARIABLES DE SORTIE:
'   Pour chaque variable "NomVariable" de type x non affichée et non modifiable par l'utilisateur, on a une variable de sortie: mx_NomVariable
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "53.3"


'masses molaires   en g/mol
md_MasseMolaireDioxydeCarbone_g_mol = 1000 * (pd_MASSE_MOLAIRE_C_kg_mol + 2 * pd_MASSE_MOLAIRE_O_kg_mol)
md_MasseMolaireOxygene_g_mol = 1000 * (2 * pd_MASSE_MOLAIRE_O_kg_mol)
md_MasseMolaireEau_g_mol = 1000 * (2 * pd_MASSE_MOLAIRE_H_kg_mol + pd_MASSE_MOLAIRE_O_kg_mol)

'masse volumique des boues en entrée, du technosable
md_MasseVolumiqueBouesFroides_kg_m3 = 1000  'kg/m3
md_MasseVolumiqueBoues250°C_kg_m3 = 800    'kg/m3
md_MasseVolumiqueTechnosable_kg_m3 = 1000  'kg/m3

'Concentration du polymère injecté
md_ConcentrationPolymereInjecteEpaississement_gL = 1.5  'g/L polymère injecté

'Taux de capture attendu sur l'épaississement
md_HypEpaississementTauxCaptureMES = 0.9

'concentrations faibles dans Athos donc Cp_boues = Cp_eau à la même température et à la même pression (le plus rigoureux serait d'utiliser les enthalpies massiques car il y a des variations de T)
md_HypCapaciteCalorifiqueBouesHumidesFroides_kWh_kg°C = 0.07507 / md_MasseMolaireEau_g_mol * 1000 / pd_NOMBRE_SECONDE_PAR_HEURE   'Perry eau liquide 5Mpa,300K->0.07507kJ/(mol.K)
md_HypCapaciteCalorifiqueBouesHumides110°C_kWh_kg°C = 0.0762 / md_MasseMolaireEau_g_mol * 1000 / pd_NOMBRE_SECONDE_PAR_HEURE   'Perry eau liquide 5Mpa,400K->0.076438kJ/(mol.K)
md_HypCapaciteCalorifiqueBouesHumides150°C_kWh_kg°C = 0.078 / md_MasseMolaireEau_g_mol * 1000 / pd_NOMBRE_SECONDE_PAR_HEURE   'Perry eau liquide 5Mpa,400K->0.076438kJ/(mol.K)
md_HypCapaciteCalorifiqueBouesHumides250°C_kWh_kg°C = 0.0872 / md_MasseMolaireEau_g_mol * 1000 / pd_NOMBRE_SECONDE_PAR_HEURE     'Perry eau liquide 5Mpa,500K->0.083643kJ/(mol.K),537.09K->0.090740kJ/(mol.K)

'Cp pour l'oxygène
md_HypCapaciteCalorifiqueOxygene15°C_kWh_kg°C = 0.032003 / md_MasseMolaireOxygene_g_mol * 1000 / pd_NOMBRE_SECONDE_PAR_HEURE    'Perry O2 gaz à 300K et 5MPa  Cp = 0.032003 kJ/(mol.K)
md_HypCapaciteCalorifiqueOxygene250°C_kWh_kg°C = 0.031815 / md_MasseMolaireOxygene_g_mol * 1000 / pd_NOMBRE_SECONDE_PAR_HEURE    'Perry O2 gaz à 500K et 5MPa  Cp = 0.031815 kJ/(mol.K)

'Cp pour le CO2
md_HypCapaciteCalorifiqueDioxydeCarbone250°C_kWh_kg°C = 0.0486 / md_MasseMolaireDioxydeCarbone_g_mol * 1000 / pd_NOMBRE_SECONDE_PAR_HEURE    'Perry CO2 gaz à 450K et 5MPa  Cp = 0.047478 kJ/(mol.K); à 650K et 5 MPa  Cp = 0.049898 kJ/(mol.K)

'enthalpie totale massique de l'humidité des fumées  (ref liquide à 0°C)
md_EnthalpieTotaleVapeur250°C_kWh_kg = 669 * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE    'chaleur totale de la vapeur présente dans les fumées (en kWh/ kg)

'PCI de la combustion des MV
md_PouvoirCalorifiqueInferieurMV_kWh_kg = 5500 * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE   '5412 / 861    'valeur Julien CHAUZY (voir s'il faut faire varier suivant le type de boues)

'siccités limites pour les boues épaissies
md_SicciteMinimaleBouesEpaissies = 0.04
md_SicciteMaximaleBouesEpaissies = 0.1

'taux de recirculation (en volume et non em MS comme le logiciel Athos)
md_TauxRecirculationTechnosable = 0.5    'par rapport à l'entrée

'ratio de MV particulaire sur MV totale boues oxydées
md_BouesOxydeesRatioMVparticulaireMVtotal = 0.5

'réduction des MV et de la DCO
md_ReacteurTauxReductionMV = 0.9    'voir quelle valeur attribuer (sur cas de comparaison avec logiciel Athos, proche de 95%)
md_ReacteurTauxReductionDCO = 0.8    '0.8  valeur Julien CHAUZY
md_HypCapaciteCalorifiqueFumees_kWh_kg°C = (md_ReacteurTauxReductionDCO * md_HypCapaciteCalorifiqueOxygene250°C_kWh_kg°C + (1 - md_ReacteurTauxReductionDCO) * md_MasseMolaireDioxydeCarbone_g_mol / md_MasseMolaireOxygene_g_mol * md_HypCapaciteCalorifiqueDioxydeCarbone250°C_kWh_kg°C) / (md_ReacteurTauxReductionDCO + (1 - md_ReacteurTauxReductionDCO) * md_MasseMolaireDioxydeCarbone_g_mol / md_MasseMolaireOxygene_g_mol)

'températures en différents points
md_ReacteurTemperatureCible_°C = 250
md_HypEcartTemperatureBouesEchangeur2_°C = 140    'valeur Julien CHAUZY différence de température entre réacteur et sortie échangeur froid oú se fait la récupération d'énergie
md_TemperatureTechnosableSortie_°C = 40
md_HypTemperatureOxygeneEntreeReacteur_°C = 15

'efficacité des échangeurs (globale sur l'ensemble)
md_EfficaciteEchangeursChaleur = 0.9

'consommation de réactifs en fonction de la DCO des boues en entrée
md_ReacteurRatioConsommationOxygene_kgO2_kgDCO = 1

'temps de séjour pour le design  1h selon Julien CHAUZY
md_ReacteurTempsSejour_h = 1 / (7500 / 8760)  'on divise par (7500/8760) car Athos ne fonctionne que 7500h/an, le reste étant de la maintenance et cela influe sur le design

'consommation électrique pour Athos
md_RatioConsommationElectrique_kWh_m3BouesEpaissies = 20   'source: Julien CHAUZY par rapport aux boues sortie épaississeur et avant recirculation

'masse volumique de l'eau
md_MasseVolumiqueEau_kg_m3 = 1000

'part de NH4 dans le NK soluble pour les retours de digestion qui partent sur l'épaississeur
md_RetoursDigestionRatioNH4surNKsoluble = 0.95

'part de DBO dans la DCO soluble
If choix_dig_simple = True Then
    md_RetoursDigestionRatioDBOsolubleDCOsoluble = 0.3 * 0.8  'cf Christophe AMIEL
ElseIf pi_ChoixDigestion = True Then
    md_RetoursDigestionRatioDBOsolubleDCOsoluble = 0.2 * 0.8  'cf Christophe AMIEL
End If

'Valeurs pour les retours (en % de l'entrée boues sur l'ensemble décanteur et filtre à plateaux)
'Dans un premier temps, on simplifie en sommant tous les retours sur le décanteur
md_RetoursDecanteurEtFiltrePlateauxRatioDCOtotaleDCOentree = 1 - md_ReacteurTauxReductionDCO
md_RetoursDecanteurEtFiltrePlateauxRatioDCOsolubleDCOtotale = 0.95
md_RetoursDecanteurEtFiltrePlateauxRatioNKtotalNKentree = 0.92
md_RetoursDecanteurEtFiltrePlateauxRatioNKsolubleNKtotal = 0.92
md_RetoursDecanteurEtFiltrePlateauxAvecDigestionRatioNH4surNKsoluble = 0.85
md_RetoursDecanteurEtFiltrePlateauxSansDigestionRatioNH4surNKsoluble = 0.6
md_RetoursDecanteurEtFiltrePlateauxAvecDigestionRatioPtotalPinlet = 0.08
md_RetoursDecanteurEtFiltrePlateauxSansDigestionRatioPtotalPinlet = 0.12
md_RetoursDecanteurEtFiltrePlateauxRatioPsolublePtotal = 0.06

'Consommation spécifique de combustible traitement RTO
md_ConsommationSpecifiqueCombustibleTraitementFumeesRTO_kWhPCI_tMVBouesEpaissies = (64 * 10) / 0.8   'vu avec J. Chauzy --> Bruxelles 0.8 tMV/h avant recirculation technosable et 64 Nm3/h de Gaz Naturel en RTO (PCI=10kWh/Nm3)


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Private Sub attribution_valeur_par_defaut()


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure attribue des valeurs par défaut aux variables affichées à l'utilisateur lorsque celles-ci ne sont pas calculées mais fixées au départ.

'VARIABLES D'ENTREE:
'   Pour chaque paramètre "NomVariable" de type noté "x" selon la nomenclature, on retrouve deux variables associées:
'   mn_NomVariable_force: TRUE si l'utilisateur a forcé une valeur, FALSE le cas contraire
'   mx_NomVariable_f: contient la valeur renseignée par l'utilisateur

'VARIABLES DE SORTIE:
'   mx_NomVariable: contient la valeur attribuée (donnée utilisateur ou valeur attribuée par défaut)
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "53.4"


'taux de capture MES épaississement
If mn_EpaississementTauxCaptureMES_force = True Then
    md_EpaississementTauxCaptureMES = md_EpaississementTauxCaptureMES_f
Else
    md_EpaississementTauxCaptureMES = 0.9
End If

'taux de polymère ajouté
If mn_EpaississementRatioConsommationPolymerePur_kg_tMS_force = True Then
    md_EpaississementRatioConsommationPolymerePur_kg_tMS = md_EpaississementRatioConsommationPolymerePur_kg_tMS_f
Else
    md_EpaississementRatioConsommationPolymerePur_kg_tMS = 6
End If

'pression de fonctionnement du réacteur
md_ReacteurPressionAbsolue_bars = 55

'concentration en sortie  g/L
If mn_DecanteurConcentrationMES_gL_force = True Then
    md_DecanteurConcentrationMES_gL = md_DecanteurConcentrationMES_gL_f
Else
    md_DecanteurConcentrationMES_gL = 100
End If

'taux de capture MES décanteur
If mn_DecanteurTauxCaptureMES_force = True Then
    md_DecanteurTauxCaptureMES = md_DecanteurTauxCaptureMES_f
Else
    md_DecanteurTauxCaptureMES = 0.95
End If

'CuSO4
If mn_ConsommationCuSO4Pur_kgj_force = True Then
    pd_ConsommationCuSO4Pur_kgj = md_ConsommationCuSO4Pur_kgj_f
Else
    pd_ConsommationCuSO4Pur_kgj = 0
End If


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub


Private Sub lecture_valeurs_forcees()


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure repère et lit les paramètres relatifs au procédé qui ont été forcés par l'utilisateur

'VARIABLES D'ENTREE:
'   pi_FiliereConsideree: index de la filière calculée
'   entrées de l'utilisateur dans l'onglet "Calculs" lorsqu'elles sont renseignées (paramètres d'entrée facultatifs).
'       Les valeurs sont reportées dans l'onglet "Feuil53" nommée "Athos" dont les noms de plages de cellules sont définis dans Excel

'VARIABLES DE SORTIE:
'   Pour une variable "NomVariable" de type x échangée avec l'utilisateur
'   mn_NomVariable_force: prend la valeur TRUE si le paramètre a été forcé i.e. valeur différente de "0" ou "Faux" dans un booléen
'       (il n'y a pas de valeur égale à "0" car elles ont été remplacées par 10^(-21) dans excel)
'   mx_NomVariable_f: prend la valeur renseignée par l'utilisateur
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim Ligne As Integer
Dim Colonne As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "53.2"
Ligne = 1
Colonne = pi_FiliereConsideree



'On repère par des booléens les entrées de l'utilisateur qui ont été forcées
'si la valeur de la cellule vaut 0 (valeur non renseignée) --> Booléen = False ; sinon Booléen = True
'si l'utilisateur force une valeur à 0, la valeur transmise est 10^(-21) donc le Booléen est à True
mn_EntreeBouesMelangeTemperature_°C_force = Feuil53.Range("athos_inlet_temp").Cells(Ligne, Colonne)
mn_EpaississementTauxCaptureMES_force = Feuil53.Range("athos_capture_MS").Cells(Ligne, Colonne)
mn_EpaississementRatioConsommationPolymerePur_kg_tMS_force = Feuil53.Range("athos_conso_polymere").Cells(Ligne, Colonne)
mn_BouesEpaissiesConcentrationMES_gL_force = Feuil53.Range("athos_ep_siccite").Cells(Ligne, Colonne)
mn_ReacteurVolume_m3_force = Feuil53.Range("athos_reac_volume").Cells(Ligne, Colonne)
mn_ConsommationOxygeneLiquidePur_kgj_force = Feuil53.Range("athos_conso_O2").Cells(Ligne, Colonne)
mn_ConsommationCuSO4Pur_kgj_force = Feuil53.Range("athos_conso_CuSO4").Cells(Ligne, Colonne)
mn_TraitementFumeesRTOConsommationCombustible_kWhPCIj_force = Feuil53.Range("athos_conso_combustible_traitement_RTO").Cells(Ligne, Colonne)
mn_DecanteurTauxCaptureMES_force = Feuil53.Range("athos_outlet_capture_ms").Cells(Ligne, Colonne)
mn_DecanteurConcentrationMES_gL_force = Feuil53.Range("athos_outlet_concentration").Cells(Ligne, Colonne)

'Pour chaque paramètre forcé, on attribue la valeur renseignée
If mn_EntreeBouesMelangeTemperature_°C_force = True Then
    md_EntreeBouesMelangeTemperature_°C_f = Feuil53.Range("athos_inlet_temp").Cells(Ligne, Colonne)
End If
If mn_EpaississementTauxCaptureMES_force = True Then
    md_EpaississementTauxCaptureMES_f = Feuil53.Range("athos_capture_MS").Cells(Ligne, Colonne)
End If
If mn_EpaississementRatioConsommationPolymerePur_kg_tMS_force = True Then
    md_EpaississementRatioConsommationPolymerePur_kg_tMS_f = Feuil53.Range("athos_conso_polymere").Cells(Ligne, Colonne)
End If
If mn_BouesEpaissiesConcentrationMES_gL_force = True Then
    md_BouesEpaissiesConcentrationMES_gL_f = Feuil53.Range("athos_ep_siccite").Cells(Ligne, Colonne)
End If
If mn_ReacteurVolume_m3_force = True Then
    md_ReacteurVolume_m3_f = Feuil53.Range("athos_reac_volume").Cells(Ligne, Colonne)
End If
If mn_ConsommationOxygeneLiquidePur_kgj_force = True Then
    md_ConsommationOxygeneLiquidePur_kgj_f = Feuil53.Range("athos_conso_O2").Cells(Ligne, Colonne)
End If
If mn_ConsommationCuSO4Pur_kgj_force = True Then
    md_ConsommationCuSO4Pur_kgj_f = Feuil53.Range("athos_conso_CuSO4").Cells(Ligne, Colonne)
End If
If mn_TraitementFumeesRTOConsommationCombustible_kWhPCIj_force = True Then
    md_TraitementFumeesRTOConsommationCombustible_kWhPCIj_f = Feuil53.Range("athos_conso_combustible_traitement_RTO").Cells(Ligne, Colonne)
End If
If mn_DecanteurTauxCaptureMES_force = True Then
    md_DecanteurTauxCaptureMES_f = Feuil53.Range("athos_outlet_capture_ms").Cells(Ligne, Colonne)
End If
If mn_DecanteurConcentrationMES_gL_force = True Then
    md_DecanteurConcentrationMES_gL_f = Feuil53.Range("athos_outlet_concentration").Cells(Ligne, Colonne)
End If


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True


End Sub


Public Sub dimensionnement(ByRef TableauRecapitulatifFluxBoues() As Double, ByRef FluxRetoursTotaux() As Double, ByRef FluxRetoursTotauxDigestion() As Double, ByRef FluxRetoursTotauxAthos() As Double, ByRef FluxRetoursSoluble() As Double, ByRef FluxRetoursSolubleDigestion() As Double, ByRef FluxRetoursSolubleAthos() As Double, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double, ByRef FluxDCOOxydeeAthosEtPyrofluid_kgj As Double)


'*********************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure permet de dimensionner le réacteur d'oxydation et d'effectuer les bilans matière et énergie lorsque l'usine reçoit sa charge nominale.
'   Pour ce faire, elle utilise les procédures suivantes:
'       - calcul_cocktail_boues_entrees: pour calculer les caractéristiques du mélange de boues en entrée
'       - attribution_Cp_Temperature: pour attibuer les caractéristiques thermiques des différents flux internes
'       - calcul_siccite_boues_epaissies: pour résoudre les bilans masses et énergies de manière à atteindre une température définie dans le réacteur
'       - calcul_composition_retours: pour calculer les jus retournées en tête d'usine ou de traitement séparé des retours (sortie eau)
'       - calcul_boues_sortie_procede: pour calculer les caractéristiques des boues en sortie du procédé (ici technosable)

'VARIABLES D'ENTREE:
'   TableauRecapitulatifFluxBoues(): contient les caractéristiques des boues alimentant le procédé
'   TableauRecapitulatifPollutionSolubleDansFluxBoues(): contient la polution soluble des boues alimentant le procédé
'   pd_TemperatureEauDimensionnement_°C: température des eaux usées définie dans le cahier des charges
'   mn_EntreeBouesMelangeTemperature_°C_force: repère si la température des boues en entrée a été forcé par l'utilisateur
'       md_EntreeBouesMelangeTemperature_°C_f: valeur renseignée par l'utilisateur
'   mn_ReacteurVolume_m3_force: repère si le volume du réacteur a été forcé par l'utilisateur
'       md_ReacteurVolume_m3_f: contient la valeur renseignée par l'utilisateur
'   md_ConsommationSpecifiqueCombustibleTraitementFumeesRTO_kWhPCI_tMVBouesEpaissies: consommation spécifique en combustible pour le traitement des fumées
'   mi_ChoixTypeCombustibleTraitementFumeesRTO: choix du type de combustible utilisé pour le traitement RTO

'VARIABLES DE SORTIE:
'   TableauRecapitulatifFluxBoues(): contient les caractéristiques des boues en sortie du procédé
'   TableauRecapitulatifPollutionSolubleDansFluxBoues(): contient la polution soluble des boues en sortie du procédé
'   FluxRetoursTotaux(): contient l'ensemble des flux de retours de la file de traitement des boues
'   FluxRetoursTotauxDigestion(): contient l'ensemble des flux de retours de digestion
'   FluxRetoursTotauxAthos(): contient l'ensemble des flux de retours de'Athos
'   FluxRetoursSoluble(): contient l'ensemble des flux de retours solubles de la file de traitement des boues
'   FluxRetoursSolubleDigestion(): contient l'ensemble des flux de retours solubles de digestion
'   FluxRetoursSolubleAthos(): contient l'ensemble des flux de retours solubles d'Athos
'   FluxDCOOxydeeAthosEtPyrofluid_kgj: DCO oxydée dans le réacteur (pour réaliser un bilan DCO sur l'usine)
'   md_ReacteurVolume_m3: volume de réacteur
'   pd_ConsommationOxygeneLiquidePur_kgj: consommation d'oxygène liquide
'   pd_EpaississementConsommationPolymerePur_kgj: consommation de polymère
'   md_TraitementFumeesRTOConsommationCombustible_kWhPCIj: consommation de combustible pour le traitement des fumées
'   pd_TraitementFumeesRTOConsommationBiogaz_kWhPCIj: Consommation de biogaz pour traitement RTO
'   pd_TraitementFumeesRTOConsommationCombustibleExterne_kWhPCIj: Consommation de biogaz pour traitement RTO
'   md_EntreeBouesMelangeTemperature_°C: température des boues en entrée du procédé
'*********************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer
Dim k As Integer
Dim EntreeBouesMelangeDCO_kgj As Double
Dim EntreeBouesMelangeDBO_kgj As Double
Dim EntreeBouesMelangeNK_kgj As Double
Dim EntreeBouesMelangeP_kgj As Double
Dim FluxPollutionSoluble_kgj(1 To nb_repere_mgL) As Double
Dim RatioBouesMinerales(1 To nb_type_boues) As Double
Dim BouesEpaissiesDCOparticulaire_kgj As Double
Dim BouesEpaissiesDCO_kgj As Double
Dim BouesEpaissiesDBO_kgj As Double
Dim BouesEpaissiesNK_kgj As Double
Dim BouesEpaissiesP_kgj As Double
Dim BouesEpaissiesFluxPollutionSoluble_kgj(nb_repere_mgL) As Double
Dim EntreeBouesRatioBouesDigerees As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "53.5"
I = 0
j = 0
k = 0
EntreeBouesMelangeDCO_kgj = 0
EntreeBouesMelangeDBO_kgj = 0
EntreeBouesMelangeNK_kgj = 0
EntreeBouesMelangeP_kgj = 0
BouesEpaissiesDCOparticulaire_kgj = 0
BouesEpaissiesDCO_kgj = 0
BouesEpaissiesDBO_kgj = 0
BouesEpaissiesNK_kgj = 0
BouesEpaissiesP_kgj = 0
EntreeBouesRatioBouesDigerees = 0
For j = 1 To nb_repere_mgL
    FluxPollutionSoluble_kgj(j) = 0
    BouesEpaissiesFluxPollutionSoluble_kgj(j) = 0
Next j
For j = 1 To nb_type_boues
    RatioBouesMinerales(j) = 0
Next j


'On calcule le mélange de boues en entrée
Call calcul_cocktail_boues_entrees(TableauRecapitulatifFluxBoues, TableauRecapitulatifPollutionSolubleDansFluxBoues, FluxPollutionSoluble_kgj, RatioBouesMinerales, EntreeBouesMelangeDCO_kgj, EntreeBouesMelangeDBO_kgj, EntreeBouesMelangeNK_kgj, EntreeBouesMelangeP_kgj, EntreeBouesRatioBouesDigerees)

'On initialise la valeur de DCO en sortie d'épaississement à celle des boues en entrée
BouesEpaissiesDCOparticulaire_kgj = EntreeBouesMelangeDCO_kgj

'Si elle n'a pas été forcée, on attribue la valeur de température (se fait ici car potentiellement différente entre le dimensionnement et la simulation)
If mn_EntreeBouesMelangeTemperature_°C_force = True Then
    md_EntreeBouesMelangeTemperature_°C = md_EntreeBouesMelangeTemperature_°C_f
Else
    md_EntreeBouesMelangeTemperature_°C = pd_TemperatureEauDimensionnement_°C
End If

'On attribue pour les flux intermédiaires les températures et capacités calorifiques
Call attribution_Cp_Temperature

'on appelle la procédure pour estimer la siccité à avoir après l'épaississement
Call calcul_siccite_boues_epaissies(EntreeBouesMelangeDCO_kgj, FluxPollutionSoluble_kgj, BouesEpaissiesFluxPollutionSoluble_kgj, BouesEpaissiesDCO_kgj, BouesEpaissiesDCOparticulaire_kgj)

'on calcule le volume de réacteur d'oxydation
If mn_ReacteurVolume_m3_force = True Then
    md_ReacteurVolume_m3 = md_ReacteurVolume_m3_f
Else
    md_ReacteurVolume_m3 = md_ReacteurTempsSejour_h / pd_NOMBRE_HEURE_PAR_JOUR * md_FluxIntermediaires(mi_FLUX_ECHANGEUR1_REACTEUR, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) / md_MasseVolumiqueBoues250°C_kg_m3
End If

'on calcule les consommables
pd_ConsommationOxygeneLiquidePur_kgj = md_FluxIntermediaires(mi_FLUX_ALIMENTATION_OXYGENE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
pd_EpaississementConsommationPolymerePur_kgj = md_EpaississementRatioConsommationPolymerePur_kg_tMS * md_EntreeBouesMelangeMES_kgj / 1000
md_TraitementFumeesRTOConsommationCombustible_kWhPCIj = md_ConsommationSpecifiqueCombustibleTraitementFumeesRTO_kWhPCI_tMVBouesEpaissies * (md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_MS_kgj) / 1000)
If mi_ChoixTypeCombustibleTraitementFumeesRTO = mi_TRAITEMENT_FUMEES_RTO_AVEC_BIOGAZ Then
    pd_TraitementFumeesRTOConsommationBiogaz_kWhPCIj = md_TraitementFumeesRTOConsommationCombustible_kWhPCIj
    pd_TraitementFumeesRTOConsommationCombustibleExterne_kWhPCIj = 0
ElseIf mi_ChoixTypeCombustibleTraitementFumeesRTO = mi_TRAITEMENT_FUMEES_RTO_AVEC_COMBUSTIBLE_EXTERNE Then
    pd_TraitementFumeesRTOConsommationBiogaz_kWhPCIj = 0
    pd_TraitementFumeesRTOConsommationCombustibleExterne_kWhPCIj = md_TraitementFumeesRTOConsommationCombustible_kWhPCIj
End If

'bilan DCO
FluxDCOOxydeeAthosEtPyrofluid_kgj = md_ReacteurTauxReductionDCO * BouesEpaissiesDCO_kgj

'calcul des retours (sortie eau)
Call calcul_composition_retours(EntreeBouesMelangeDCO_kgj, EntreeBouesMelangeDBO_kgj, EntreeBouesMelangeNK_kgj, EntreeBouesMelangeP_kgj, EntreeBouesRatioBouesDigerees, FluxRetoursTotaux, FluxRetoursSoluble, FluxPollutionSoluble_kgj, BouesEpaissiesFluxPollutionSoluble_kgj, FluxRetoursTotauxDigestion, FluxRetoursSolubleDigestion, FluxRetoursTotauxAthos, FluxRetoursSolubleAthos)

'calcul des boues (technosable) en sortie
Call calcul_boues_sortie_procede(BouesEpaissiesDCO_kgj, BouesEpaissiesNK_kgj, BouesEpaissiesP_kgj, RatioBouesMinerales, FluxRetoursTotauxAthos, TableauRecapitulatifFluxBoues, TableauRecapitulatifPollutionSolubleDansFluxBoues)


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Private Sub calcul_bilan_masse(ByVal BouesEpaissiesDCO_kgj As Double)

'******************************************************************************************************************************************************
'OBJECTIF
'   Cette procédure permet la résolution des bilans massiques du procédé par un système itératif.
'   On initialise les caractéristiques du flux de technosable en sortie "mi_FLUX_SORTIE_TECHNOSABLE"
'   et de la recirculation "mi_FLUX_RECIRCULATION_TECHNOSABLE" puis on recalcule l'ensemble des flux jusqu'à convergence des caractéristiques
'   du flux de recirculation.

'VARIABLES D'ENTREE:
'   BouesEpaissiesDCO_kgj: DCO contenue dans les boues épaissies
'   md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX):
'       certaines caractéristiques des flux sont utilisés comme données d'entrées
'   md_DecanteurConcentrationMES_gL: concentration en MES du technosable en sortie de décanteur
'   md_MasseVolumiqueTechnosable_kg_m3: masse volumique du technosable
'   md_DecanteurTauxCaptureMES: taux de capture du décanteur
'   md_ReacteurTauxReductionMV: taux de réduction des MV dans le réacteur
'   md_BouesOxydeesRatioMVparticulaireMVtotal: ratio MVparticulaires/MVtotales des boues oxydées (en sortie de réacteur)
'   md_TauxRecirculationTechnosable: taux de recirculation du technosable défini par rapport aux boues épaissies et en kg de matière humide
'   md_ReacteurRatioConsommationOxygene_kgO2_kgDCO: consommation spécifique en oxygène pour oxyder la DCO des boues
'   md_ReacteurTauxReductionDCO: pourcentage de la DCO des boues qui est oxydée

'VARIABLES DE SORTIE:
'   md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX): calcul des caractéristiques
'       non relatives à l'énergie
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree
    
'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim EcartRelatifMS As Double
Dim EcartRelatifBH As Double
Dim EcartRelatifMV As Double
Dim ConvergenceAtteinte As Boolean

'Initialisation des variables locales
NumeroErreurNonGeree = "53.22"
EcartRelatifMS = 0
EcartRelatifBH = 0
EcartRelatifMV = 0
ConvergenceAtteinte = False


'on initialise les caractéristiques du technosable de sortie et de la recirculation
'on calcule la sortie théorique
md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_SICCITE) = md_DecanteurConcentrationMES_gL / md_MasseVolumiqueTechnosable_kg_m3
'on calcule la matière minérale que l'on ne touche pas   kg/j
md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj) = md_DecanteurTauxCaptureMES * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_MS_kgj) * (1 - md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_RATIO_MV_MES))
'on calcule la MV particulaire en sortie   kg/j
md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES) = md_DecanteurTauxCaptureMES * md_BouesOxydeesRatioMVparticulaireMVtotal * (1 - md_ReacteurTauxReductionMV) * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_MS_kgj)
'on calcule la MS totale   en kg/j
md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj) = md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj) + md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES)
'on en déduit le ratio MV/MES
md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES) / md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj)
md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) = md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj) / md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_SICCITE)
'On estime la recirculation
md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) = md_TauxRecirculationTechnosable * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_SICCITE) = md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_SICCITE)
md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj) = md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_SICCITE) * md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES)

ConvergenceAtteinte = False

Do Until ConvergenceAtteinte = True

    'criteres pour évaluer convergence
    EcartRelatifMS = md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj)
    EcartRelatifBH = md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
    EcartRelatifMV = md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj)

    'on mélange les boues épaissies avec la recirculation
    Call calcul_melange_flux(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_FLUX_MELANGE_ECHANGEUR1)

    'on propage les flux massiques et le Cp (qui en théorie doit changer/ voir si on le recalcule en fonction de la convergence antérieure) après l'échangeur  (le calcul de la température se fera plus tard)
    Call propagation_flux_massiques(mi_FLUX_MELANGE_ECHANGEUR1, mi_FLUX_ECHANGEUR1_REACTEUR)

    'O2 utilisé
    md_FluxIntermediaires(mi_FLUX_ALIMENTATION_OXYGENE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) = md_ReacteurRatioConsommationOxygene_kgO2_kgDCO * BouesEpaissiesDCO_kgj

    'fumées
    md_FluxIntermediaires(mi_FLUX_SORTIE_FUMEES, mi_CARACTERISTIQUE_MS_kgj) = md_ReacteurTauxReductionMV * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_MS_kgj) + md_FluxIntermediaires(mi_FLUX_ALIMENTATION_OXYGENE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
    md_FluxIntermediaires(mi_FLUX_SORTIE_FUMEES, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) = md_FluxIntermediaires(mi_FLUX_SORTIE_FUMEES, mi_CARACTERISTIQUE_MS_kgj) * (1 + humidite_air_gH2O_gAS(md_ReacteurTauxReductionDCO * md_MasseMolaireDioxydeCarbone_g_mol + (1 - md_ReacteurTauxReductionDCO) * md_MasseMolaireOxygene_g_mol, md_ReacteurPressionAbsolue_bars, md_ReacteurTemperatureCible_°C, 1))    'on ajoute l'humidité de l'air à 100%
    md_FluxEauEvaporee_kgj = md_FluxIntermediaires(mi_FLUX_SORTIE_FUMEES, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) - md_FluxIntermediaires(mi_FLUX_SORTIE_FUMEES, mi_CARACTERISTIQUE_MS_kgj)

    'sortie réacteur
    md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR1_REACTEUR, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) + md_FluxIntermediaires(mi_FLUX_ALIMENTATION_OXYGENE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) - md_FluxIntermediaires(mi_FLUX_SORTIE_FUMEES, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
    md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_MS_kgj) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR1_REACTEUR, mi_CARACTERISTIQUE_MS_kgj) - md_ReacteurTauxReductionMV * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_MS_kgj)
    md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR1_REACTEUR, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_ECHANGEUR1_REACTEUR, mi_CARACTERISTIQUE_MS_kgj) - md_ReacteurTauxReductionMV * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_MS_kgj)
    md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_RATIO_MV_MES) / md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_MS_kgj)
    md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_SICCITE) = md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_MS_kgj) / md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)

    'on propage les flux massiques et le Cp (qui en théorie doit changer/ voir si on le recalcule en fonction de la convergence antérieure) après l'échangeur  (le calcul de la température se fera plus tard)
    Call propagation_flux_massiques(mi_FLUX_REACTEUR_ECHANGEUR2, mi_FLUX_ECHANGEUR2_ECHANGEUR3)
    Call propagation_flux_massiques(mi_FLUX_ECHANGEUR2_ECHANGEUR3, mi_FLUX_ECHANGEUR3_DECANTEUR)

    'on calcule le décanteur (sortie Athos)
    md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj) = md_DecanteurTauxCaptureMES * md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_MS_kgj) * (md_BouesOxydeesRatioMVparticulaireMVtotal * md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_RATIO_MV_MES) + (1 - md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_RATIO_MV_MES)))
    md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_SICCITE) = md_DecanteurConcentrationMES_gL / md_MasseVolumiqueTechnosable_kg_m3
    md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) = md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj) / md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_SICCITE)
    md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES) = md_DecanteurTauxCaptureMES * md_BouesOxydeesRatioMVparticulaireMVtotal * md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_MS_kgj)
    md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES) / md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj)
    md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C)
    md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_TEMPERATURE_°C)

    'on calcule la recirculation
    md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) = md_TauxRecirculationTechnosable * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
    md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_SICCITE) = md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_SICCITE)
    md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj) = md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_SICCITE) * md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
    md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES)

    'criteres pour évaluer convergence
    EcartRelatifMS = (EcartRelatifMS - md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj)) ^ 2 / md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj)
    EcartRelatifBH = (EcartRelatifBH - md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)) ^ 2 / md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
    EcartRelatifMV = (EcartRelatifMV - md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj)) ^ 2 / (md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj))

    If (EcartRelatifMS + EcartRelatifBH + EcartRelatifMV) < md_CRITERE_CONVERGENCE_BILAN_MATIERE Then
        ConvergenceAtteinte = True
    End If

Loop

md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_SICCITE) = md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_SICCITE)
md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj) = md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj) - md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_MS_kgj)
md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) = md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) - md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
md_FluxIntermediaires(mi_FLUX_SORTIE_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_DECANTEUR_TECHNOSABLE, mi_CARACTERISTIQUE_RATIO_MV_MES)



Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True


End Sub

Private Sub propagation_flux_massiques(ByVal IndexFluxReference As Integer, ByVal IndexFluxCible As Integer)


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure permet de propager les données relatives aux bilans matières d'un flux à un autre lorsque celles-ci ne sont pas modifiées
'   par exemple lors d'un passage dans un échangeur de chaleur

'VARIABLES D'ENTREE:
'   IndexFluxReference: repère le flux de départ
'   IndexFluxCible: repère le flux qui reprend les caractéristiques du flux de référence
'   md_FluxIntermediaires(IndexFluxReference, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX): contient les caractéristques du flux de référence

'VARIABLES DE SORTIE:
'   md_FluxIntermediaires(IndexFluxCible, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX): contient les caractéristiques du flux cible
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "53.29"
I = 0


For I = mi_CARACTERISTIQUE_MS_kgj To mi_CARACTERISTIQUE_RATIO_MV_MES
    md_FluxIntermediaires(IndexFluxCible, I) = md_FluxIntermediaires(IndexFluxReference, I)
Next I


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Private Sub calcul_melange_flux(ByVal IndexFluxAMelanger1 As Integer, ByVal IndexFluxAMelanger2 As Integer, ByVal IndexFluxResultant As Integer)


'******************************************************************************************************************************************************
'OBJECTIF
'   Cette procédure permet d'effectuer le mélange de deux flux "IndexFluxAMelanger1" et "IndexFluxAMelanger2"
'   et d'estimer les caractéristiques du flux résultant "IndexFluxResultant".
'   Dans la pratique elle sert à mélanger les flux suivants: "mi_FLUX_EPAISSISSEMENT_MELANGE" et "mi_FLUX_RECIRCULATION_TECHNOSABLE" pour obtenir
'   le flux "mi_FLUX_MELANGE_ECHANGEUR1".

'VARIABLES D'ENTREE:
'   IndexFluxAMelanger1: index du flux à mélanger 1 (position dans la matrice "md_FluxIntermediaires")
'   IndexFluxAMelanger2: index du flux à mélanger 2 (position dans la matrice "md_FluxIntermediaires")
'   IndexFluxResultant: index du flux résultant du mélange (position dans la matrice "md_FluxIntermediaires")
'   md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX): contient les flux à mélanger

'VARIABLES DE SORTIE:
'   md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX): contient le flux résultant
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "53.27"

md_FluxIntermediaires(IndexFluxResultant, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) = md_FluxIntermediaires(IndexFluxAMelanger1, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) + md_FluxIntermediaires(IndexFluxAMelanger2, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
md_FluxIntermediaires(IndexFluxResultant, mi_CARACTERISTIQUE_MS_kgj) = md_FluxIntermediaires(IndexFluxAMelanger1, mi_CARACTERISTIQUE_MS_kgj) + md_FluxIntermediaires(IndexFluxAMelanger2, mi_CARACTERISTIQUE_MS_kgj)
md_FluxIntermediaires(IndexFluxResultant, mi_CARACTERISTIQUE_SICCITE) = md_FluxIntermediaires(IndexFluxResultant, mi_CARACTERISTIQUE_MS_kgj) / md_FluxIntermediaires(IndexFluxResultant, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
'MV/MES du mélange
md_FluxIntermediaires(IndexFluxResultant, mi_CARACTERISTIQUE_RATIO_MV_MES) = (md_FluxIntermediaires(IndexFluxAMelanger1, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(IndexFluxAMelanger1, mi_CARACTERISTIQUE_MS_kgj) + md_FluxIntermediaires(IndexFluxAMelanger2, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(IndexFluxAMelanger2, mi_CARACTERISTIQUE_MS_kgj))
md_FluxIntermediaires(IndexFluxResultant, mi_CARACTERISTIQUE_RATIO_MV_MES) = md_FluxIntermediaires(IndexFluxResultant, mi_CARACTERISTIQUE_RATIO_MV_MES) / md_FluxIntermediaires(IndexFluxResultant, mi_CARACTERISTIQUE_MS_kgj)
'Cp du mélange
md_FluxIntermediaires(IndexFluxResultant, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = (md_FluxIntermediaires(IndexFluxAMelanger1, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) * md_FluxIntermediaires(IndexFluxAMelanger1, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) + md_FluxIntermediaires(IndexFluxAMelanger2, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) * md_FluxIntermediaires(IndexFluxAMelanger2, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj))
md_FluxIntermediaires(IndexFluxResultant, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) = md_FluxIntermediaires(IndexFluxResultant, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) / md_FluxIntermediaires(IndexFluxResultant, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)



Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True


End Sub

Private Sub calcul_epaississement(ByRef BouesEpaissiesDCOparticulaire_kgj As Double, ByVal EntreeBouesMelangeDCO_kgj As Double)


'******************************************************************************************************************************************************
'OBJECTIF
'   Cette procédure permet de faire un bilan matière relatif aux boues au niveau de l'épaississement

'VARIABLES D'ENTREE:
'   EntreeBouesMelangeDCO_kgj: flux total de DCO des boues en entrées du procédé
'   md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX): certaines caractéristiques des flux sont utilisés
'   md_EpaississementTauxCaptureMES: taux de capture de l'épaississement

'VARIABLES DE SORTIE:
'   BouesEpaissiesDCOparticulaire_kgj: flux total de DCO des boues en sortie de l'épaississement
'   md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX): calcul de données relatives aux bilans masse
'       pour le flux de sortie de l'épaississement "mi_FLUX_EPAISSISSEMENT_MELANGE"
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "53.26"

'calcul des MES
md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_MS_kgj) = md_EpaississementTauxCaptureMES * md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_MS_kgj)
'calcul des boues humides
md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) = md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_MS_kgj) / md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_SICCITE)
'calcul du ratio MV/MES
md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_BOUES, mi_CARACTERISTIQUE_RATIO_MV_MES)

BouesEpaissiesDCOparticulaire_kgj = EntreeBouesMelangeDCO_kgj * md_EpaississementTauxCaptureMES


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True


End Sub

Private Sub calcul_bilan_energie()


'******************************************************************************************************************************************************
'OBJECTIF
'   Cette procédure permet la résolution des bilans d'énergie du procédé et notamment le calcul de la température du réacteur.
'   Il remplit une matrice "MatriceGauss" contenant un système de 5 équations (bilans d'énergie en différents endroits) et
'   de 5 inconnues (températures à différents endroits, énergie échangée) puis fait appel à la procédure publique "solve_systeme_lineaire_Gauss"
'   pour résoudre le système. Les résultats obtenus sont alors réaffectés aux variables de sortie définies au niveau du module

'VARIABLES D'ENTREE:
'   md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX): certaines caractéristiques des flux sont utilisés
'   md_EfficaciteEchangeursChaleur: efficacité des échangeurs de chaleur
'   md_PouvoirCalorifiqueInferieurMV_kWh_kg: pouvoir calorifique inférieur de la combustion des matières volatiles
'   md_ReacteurTauxReductionMV: taux de réduction des matières volatiles dans le réacteur
'   md_HypEcartTemperatureBouesEchangeur2_°C: écart de température souhaité entre les boues d'entrée et de sortie de l'échangeur 2

'VARIABLES DE SORTIE:
'   md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX): renseignement pour chaque flux
'       des caractéristiques de coordonnée mi_CARACTERISTIQUE_TEMPERATURE_°C et mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C
'   md_ReacteurTemperature_°C: température du réacteur
'   md_EnergieEchangeeEntreEch1etEch2_kWhj: Energie échangée entre les échangeurs 1 et 2
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des constantes
'taille du système d'équations
Const DIMENSION_SYSTEME_EQUATION = 5
'Bilans d'énergies (BE) considérés
Const BILAN_ENERGIE_MELANGE_BOUES_EPAISSIES_AVEC_RECIRCULATION = 1
Const BILAN_ENERGIE_ECHANGEUR1 = 2
Const BILAN_ENERGIE_REACTEUR = 3
Const BILAN_ENERGIE_ECHANGEUR2 = 4
Const BILAN_ENERGIE_ECART_TEMPERATURE_BOUES_ECHANGEUR2 = 5
'Variables inconnues
Const INCONNUE_TEMPERATURE_ENTREE_BOUES_ECHANGEUR1 = 1
Const INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR1 = 2
Const INCONNUE_TEMPERATURE_REACTEUR = 3
Const INCONNUE_ENERGIE_ECHANGEE_ENTRE_ECH1_ET_ECH2_kWhj = 4
Const INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR2 = 5
Const RESULTATS_RESOLUTION = 6
    
'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim MatriceGauss(1 To DIMENSION_SYSTEME_EQUATION, 1 To DIMENSION_SYSTEME_EQUATION + 1) As Double

'initialisation des variables locales
NumeroErreurNonGeree = "53.21"

'mélange boues épaissies et boues recirculées
MatriceGauss(BILAN_ENERGIE_MELANGE_BOUES_EPAISSIES_AVEC_RECIRCULATION, INCONNUE_TEMPERATURE_ENTREE_BOUES_ECHANGEUR1) = md_FluxIntermediaires(mi_FLUX_MELANGE_ECHANGEUR1, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) * md_FluxIntermediaires(mi_FLUX_MELANGE_ECHANGEUR1, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C)
MatriceGauss(BILAN_ENERGIE_MELANGE_BOUES_EPAISSIES_AVEC_RECIRCULATION, INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR1) = 0
MatriceGauss(BILAN_ENERGIE_MELANGE_BOUES_EPAISSIES_AVEC_RECIRCULATION, INCONNUE_TEMPERATURE_REACTEUR) = 0
MatriceGauss(BILAN_ENERGIE_MELANGE_BOUES_EPAISSIES_AVEC_RECIRCULATION, INCONNUE_ENERGIE_ECHANGEE_ENTRE_ECH1_ET_ECH2_kWhj) = 0
MatriceGauss(BILAN_ENERGIE_MELANGE_BOUES_EPAISSIES_AVEC_RECIRCULATION, INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR2) = 0
MatriceGauss(BILAN_ENERGIE_MELANGE_BOUES_EPAISSIES_AVEC_RECIRCULATION, RESULTATS_RESOLUTION) = md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) * md_EntreeBouesMelangeTemperature_°C + md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) * md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) * md_FluxIntermediaires(mi_FLUX_RECIRCULATION_TECHNOSABLE, mi_CARACTERISTIQUE_TEMPERATURE_°C)

'bilan échangeur de pré-chauffage des boues
MatriceGauss(BILAN_ENERGIE_ECHANGEUR1, INCONNUE_TEMPERATURE_ENTREE_BOUES_ECHANGEUR1) = md_FluxIntermediaires(mi_FLUX_MELANGE_ECHANGEUR1, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) * md_FluxIntermediaires(mi_FLUX_MELANGE_ECHANGEUR1, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C)
MatriceGauss(BILAN_ENERGIE_ECHANGEUR1, INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR1) = -md_FluxIntermediaires(mi_FLUX_ECHANGEUR1_REACTEUR, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) * md_FluxIntermediaires(mi_FLUX_ECHANGEUR1_REACTEUR, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C)
MatriceGauss(BILAN_ENERGIE_ECHANGEUR1, INCONNUE_TEMPERATURE_REACTEUR) = 0
MatriceGauss(BILAN_ENERGIE_ECHANGEUR1, INCONNUE_ENERGIE_ECHANGEE_ENTRE_ECH1_ET_ECH2_kWhj) = md_EfficaciteEchangeursChaleur
MatriceGauss(BILAN_ENERGIE_ECHANGEUR1, INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR2) = 0
MatriceGauss(BILAN_ENERGIE_ECHANGEUR1, RESULTATS_RESOLUTION) = 0

'bilan d'énergie sur le réacteur
MatriceGauss(BILAN_ENERGIE_REACTEUR, INCONNUE_TEMPERATURE_ENTREE_BOUES_ECHANGEUR1) = 0
MatriceGauss(BILAN_ENERGIE_REACTEUR, INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR1) = -md_FluxIntermediaires(mi_FLUX_ECHANGEUR1_REACTEUR, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) * md_FluxIntermediaires(mi_FLUX_ECHANGEUR1_REACTEUR, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C)
MatriceGauss(BILAN_ENERGIE_REACTEUR, INCONNUE_TEMPERATURE_REACTEUR) = (md_FluxIntermediaires(mi_FLUX_SORTIE_FUMEES, mi_CARACTERISTIQUE_MS_kgj) * md_FluxIntermediaires(mi_FLUX_SORTIE_FUMEES, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) + md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) * md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C))
MatriceGauss(BILAN_ENERGIE_REACTEUR, INCONNUE_ENERGIE_ECHANGEE_ENTRE_ECH1_ET_ECH2_kWhj) = 0
MatriceGauss(BILAN_ENERGIE_ECHANGEUR1, INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR2) = 0
MatriceGauss(BILAN_ENERGIE_REACTEUR, RESULTATS_RESOLUTION) = md_PouvoirCalorifiqueInferieurMV_kWh_kg * md_ReacteurTauxReductionMV * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_MS_kgj) - md_EnthalpieTotaleVapeur250°C_kWh_kg * md_FluxEauEvaporee_kgj + md_FluxIntermediaires(mi_FLUX_ALIMENTATION_OXYGENE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) * md_FluxIntermediaires(mi_FLUX_ALIMENTATION_OXYGENE, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) * md_FluxIntermediaires(mi_FLUX_ALIMENTATION_OXYGENE, mi_CARACTERISTIQUE_TEMPERATURE_°C)

'bilan d'énergie sur l'échangeur de refroidissement
MatriceGauss(BILAN_ENERGIE_ECHANGEUR2, INCONNUE_TEMPERATURE_ENTREE_BOUES_ECHANGEUR1) = 0
MatriceGauss(BILAN_ENERGIE_ECHANGEUR2, INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR1) = 0
MatriceGauss(BILAN_ENERGIE_ECHANGEUR2, INCONNUE_TEMPERATURE_REACTEUR) = md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) * md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C)
MatriceGauss(BILAN_ENERGIE_ECHANGEUR2, INCONNUE_ENERGIE_ECHANGEE_ENTRE_ECH1_ET_ECH2_kWhj) = -1
MatriceGauss(BILAN_ENERGIE_ECHANGEUR2, INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR2) = -md_FluxIntermediaires(mi_FLUX_ECHANGEUR2_ECHANGEUR3, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) * md_FluxIntermediaires(mi_FLUX_ECHANGEUR2_ECHANGEUR3, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C)
MatriceGauss(BILAN_ENERGIE_ECHANGEUR2, RESULTATS_RESOLUTION) = 0

'delta de température sur l'échangeur de refroidissement
MatriceGauss(BILAN_ENERGIE_ECART_TEMPERATURE_BOUES_ECHANGEUR2, INCONNUE_TEMPERATURE_ENTREE_BOUES_ECHANGEUR1) = 0
MatriceGauss(BILAN_ENERGIE_ECART_TEMPERATURE_BOUES_ECHANGEUR2, INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR1) = 0
MatriceGauss(BILAN_ENERGIE_ECART_TEMPERATURE_BOUES_ECHANGEUR2, INCONNUE_TEMPERATURE_REACTEUR) = 1
MatriceGauss(BILAN_ENERGIE_ECART_TEMPERATURE_BOUES_ECHANGEUR2, INCONNUE_ENERGIE_ECHANGEE_ENTRE_ECH1_ET_ECH2_kWhj) = 0
MatriceGauss(BILAN_ENERGIE_ECART_TEMPERATURE_BOUES_ECHANGEUR2, INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR2) = -1
MatriceGauss(BILAN_ENERGIE_ECART_TEMPERATURE_BOUES_ECHANGEUR2, RESULTATS_RESOLUTION) = md_HypEcartTemperatureBouesEchangeur2_°C


'on fait appel au solveur de système linéaire
Call solve_systeme_lineaire_Gauss(DIMENSION_SYSTEME_EQUATION, MatriceGauss)

md_FluxIntermediaires(mi_FLUX_MELANGE_ECHANGEUR1, mi_CARACTERISTIQUE_TEMPERATURE_°C) = MatriceGauss(INCONNUE_TEMPERATURE_ENTREE_BOUES_ECHANGEUR1, RESULTATS_RESOLUTION)
md_FluxIntermediaires(mi_FLUX_ECHANGEUR1_REACTEUR, mi_CARACTERISTIQUE_TEMPERATURE_°C) = MatriceGauss(INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR1, RESULTATS_RESOLUTION)
md_ReacteurTemperature_°C = MatriceGauss(INCONNUE_TEMPERATURE_REACTEUR, RESULTATS_RESOLUTION)
md_EnergieEchangeeEntreEch1etEch2_kWhj = MatriceGauss(INCONNUE_ENERGIE_ECHANGEE_ENTRE_ECH1_ET_ECH2_kWhj, RESULTATS_RESOLUTION)
md_FluxIntermediaires(mi_FLUX_ECHANGEUR2_ECHANGEUR3, mi_CARACTERISTIQUE_TEMPERATURE_°C) = MatriceGauss(INCONNUE_TEMPERATURE_SORTIE_BOUES_ECHANGEUR2, RESULTATS_RESOLUTION)

md_FluxIntermediaires(mi_FLUX_SORTIE_FUMEES, mi_CARACTERISTIQUE_TEMPERATURE_°C) = md_ReacteurTemperature_°C
md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_TEMPERATURE_°C) = md_ReacteurTemperature_°C


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True


End Sub



Public Sub fonctionnement_reel(ByRef TableauRecapitulatifFluxBoues() As Double, ByRef FluxRetoursTotaux() As Double, ByRef FluxRetoursTotauxDigestion() As Double, ByRef FluxRetoursTotauxAthos() As Double, ByRef FluxRetoursSoluble() As Double, ByRef FluxRetoursSolubleDigestion() As Double, ByRef FluxRetoursSolubleAthos() As Double, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double, ByRef FluxDCOOxydeeAthosEtPyrofluid_kgj As Double)


'*********************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure permet d'effectuer les bilans matière et énergie en conditions d'exploitation réelles.
'   Pour ce faire, elle utilise les procédures suivantes:
'       - calcul_cocktail_boues_entrees: pour calculer les caractéristiques du mélange de boues en entrée
'       - attribution_Cp_Temperature: pour attibuer les caractéristiques thermiques des différents flux internes
'       - calcul_siccite_boues_epaissies: pour résoudre les bilans masses et énergies de manière à atteindre une température définie dans le réacteur
'       - calcul_composition_retours: pour calculer les jus retournées en tête d'usine ou de traitement séparé des retours (sortie eau)
'       - calcul_boues_sortie_procede: pour calculer les caractéristiques des boues en sortie du procédé (ici technosable)

'VARIABLES D'ENTREE:
'   TableauRecapitulatifFluxBoues(): contient les caractéristiques des boues alimentant le procédé
'   TableauRecapitulatifPollutionSolubleDansFluxBoues(): contient la polution soluble des boues alimentant le procédé
'   pd_TemperatureEauConditionsReelles_°C: température des eaux usées réelle
'   mn_EntreeBouesMelangeTemperature_°C_force: repère si la température des boues en entrée a été forcé par l'utilisateur
'       md_EntreeBouesMelangeTemperature_°C_f: valeur renseignée par l'utilisateur
'   mn_ConsommationOxygeneLiquidePur_kgj_force: repère si la consommation d'oxygène a été forcée par l'utilisateur
'       md_ConsommationOxygeneLiquidePur_kgj_f: valeur renseignée par l'utilisateur
'   mn_TraitementFumeesRTOConsommationCombustible_kWhPCIj_force: repère si la consommation en combustible du traitement des fumées a été forcée par l'utilisateur
'        md_TraitementFumeesRTOConsommationCombustible_kWhPCIj_f: valeur renseignée par l'utilisateur
'   md_ConsommationSpecifiqueCombustibleTraitementFumeesRTO_kWhPCI_tMVBouesEpaissies: consommation spécifique en combustible pour le traitement des fumées
'   mi_ChoixTypeCombustibleTraitementFumeesRTO: choix du type de combustible utilisé pour le traitement RTO

'VARIABLES DE SORTIE:
'   TableauRecapitulatifFluxBoues(): contient les caractéristiques des boues en sortie du procédé
'   TableauRecapitulatifPollutionSolubleDansFluxBoues(): contient la polution soluble des boues en sortie du procédé
'   FluxRetoursTotaux(): contient l'ensemble des flux de retours de la file de traitement des boues
'   FluxRetoursTotauxDigestion(): contient l'ensemble des flux de retours de digestion
'   FluxRetoursTotauxAthos(): contient l'ensemble des flux de retours de'Athos
'   FluxRetoursSoluble(): contient l'ensemble des flux de retours solubles de la file de traitement des boues
'   FluxRetoursSolubleDigestion(): contient l'ensemble des flux de retours solubles de digestion
'   FluxRetoursSolubleAthos(): contient l'ensemble des flux de retours solubles d'Athos
'   FluxDCOOxydeeAthosEtPyrofluid_kgj: DCO oxydée dans le réacteur (pour réaliser un bilan DCO sur l'usine)
'   pd_ConsommationOxygeneLiquidePur_kgj: consommation d'oxygène liquide
'   pd_EpaississementConsommationPolymerePur_kgj: consommation de polymère
'   md_TraitementFumeesRTOConsommationCombustible_kWhPCIj: consommation de combustible pour le traitement des fumées
'   pd_TraitementFumeesRTOConsommationBiogaz_kWhPCIj: Consommation de biogaz pour traitement RTO
'   pd_TraitementFumeesRTOConsommationCombustibleExterne_kWhPCIj: Consommation de biogaz pour traitement RTO
'   md_EntreeBouesMelangeTemperature_°C: température des boues en entrée du procédé
'*********************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer
Dim k As Integer
Dim EntreeBouesMelangeDCO_kgj As Double
Dim EntreeBouesMelangeDBO_kgj As Double
Dim EntreeBouesMelangeNK_kgj As Double
Dim EntreeBouesMelangeP_kgj As Double
Dim FluxPollutionSoluble_kgj(1 To nb_repere_mgL) As Double
Dim RatioBouesMinerales(1 To nb_type_boues) As Double
Dim BouesEpaissiesDCOparticulaire_kgj As Double
Dim BouesEpaissiesDCO_kgj As Double
Dim BouesEpaissiesDBO_kgj As Double
Dim BouesEpaissiesNK_kgj As Double
Dim BouesEpaissiesP_kgj As Double
Dim BouesEpaissiesFluxPollutionSoluble_kgj(nb_repere_mgL) As Double
Dim EntreeBouesRatioBouesDigerees As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "53.6"
I = 0
j = 0
k = 0
EntreeBouesMelangeDCO_kgj = 0
EntreeBouesMelangeDBO_kgj = 0
EntreeBouesMelangeNK_kgj = 0
EntreeBouesMelangeP_kgj = 0
BouesEpaissiesDCOparticulaire_kgj = 0
BouesEpaissiesDCO_kgj = 0
BouesEpaissiesDBO_kgj = 0
BouesEpaissiesNK_kgj = 0
BouesEpaissiesP_kgj = 0
EntreeBouesRatioBouesDigerees = 0
For j = 1 To nb_repere_mgL
    FluxPollutionSoluble_kgj(j) = 0
    BouesEpaissiesFluxPollutionSoluble_kgj(j) = 0
Next j
For j = 1 To nb_type_boues
    RatioBouesMinerales(j) = 0
Next j


'On calcule le mélange de boues en entrée
Call calcul_cocktail_boues_entrees(TableauRecapitulatifFluxBoues, TableauRecapitulatifPollutionSolubleDansFluxBoues, FluxPollutionSoluble_kgj, RatioBouesMinerales, EntreeBouesMelangeDCO_kgj, EntreeBouesMelangeDBO_kgj, EntreeBouesMelangeNK_kgj, EntreeBouesMelangeP_kgj, EntreeBouesRatioBouesDigerees)

'On initialise la valeur de DCO en sortie d'épaississement à celle des boues en entrée
BouesEpaissiesDCOparticulaire_kgj = EntreeBouesMelangeDCO_kgj

'Si elle n'a pas été forcée, on attribue la valeur de température (se fait ici car potentiellement différente entre le dimensionnement et la simulation)
If mn_EntreeBouesMelangeTemperature_°C_force = True Then
    md_EntreeBouesMelangeTemperature_°C = md_EntreeBouesMelangeTemperature_°C_f
Else
    md_EntreeBouesMelangeTemperature_°C = pd_TemperatureEauConditionsReelles_°C
End If

'On attribue pour les flux intermédiaires les températures et capacités calorifiques
Call attribution_Cp_Temperature

'on appelle la procédure pour estimer la siccité à avoir après l'épaississement
Call calcul_siccite_boues_epaissies(EntreeBouesMelangeDCO_kgj, FluxPollutionSoluble_kgj, BouesEpaissiesFluxPollutionSoluble_kgj, BouesEpaissiesDCO_kgj, BouesEpaissiesDCOparticulaire_kgj)

'on calcule les consommables
If mn_ConsommationOxygeneLiquidePur_kgj_force = True Then
    pd_ConsommationOxygeneLiquidePur_kgj = md_ConsommationOxygeneLiquidePur_kgj_f
Else
    pd_ConsommationOxygeneLiquidePur_kgj = md_FluxIntermediaires(mi_FLUX_ALIMENTATION_OXYGENE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)
End If
pd_EpaississementConsommationPolymerePur_kgj = md_EpaississementRatioConsommationPolymerePur_kg_tMS * md_EntreeBouesMelangeMES_kgj / 1000
If mn_TraitementFumeesRTOConsommationCombustible_kWhPCIj_force = True Then
    md_TraitementFumeesRTOConsommationCombustible_kWhPCIj = md_TraitementFumeesRTOConsommationCombustible_kWhPCIj_f
Else
    md_TraitementFumeesRTOConsommationCombustible_kWhPCIj = md_ConsommationSpecifiqueCombustibleTraitementFumeesRTO_kWhPCI_tMVBouesEpaissies * (md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_MS_kgj) / 1000)
End If
If mi_ChoixTypeCombustibleTraitementFumeesRTO = mi_TRAITEMENT_FUMEES_RTO_AVEC_BIOGAZ Then
    pd_TraitementFumeesRTOConsommationBiogaz_kWhPCIj = md_TraitementFumeesRTOConsommationCombustible_kWhPCIj
    pd_TraitementFumeesRTOConsommationCombustibleExterne_kWhPCIj = 0
ElseIf mi_ChoixTypeCombustibleTraitementFumeesRTO = mi_TRAITEMENT_FUMEES_RTO_AVEC_COMBUSTIBLE_EXTERNE Then
    pd_TraitementFumeesRTOConsommationBiogaz_kWhPCIj = 0
    pd_TraitementFumeesRTOConsommationCombustibleExterne_kWhPCIj = md_TraitementFumeesRTOConsommationCombustible_kWhPCIj
End If

'bilan DCO
FluxDCOOxydeeAthosEtPyrofluid_kgj = md_ReacteurTauxReductionDCO * BouesEpaissiesDCO_kgj

'calcul des retours (sortie eau)
Call calcul_composition_retours(EntreeBouesMelangeDCO_kgj, EntreeBouesMelangeDBO_kgj, EntreeBouesMelangeNK_kgj, EntreeBouesMelangeP_kgj, EntreeBouesRatioBouesDigerees, FluxRetoursTotaux, FluxRetoursSoluble, FluxPollutionSoluble_kgj, BouesEpaissiesFluxPollutionSoluble_kgj, FluxRetoursTotauxDigestion, FluxRetoursSolubleDigestion, FluxRetoursTotauxAthos, FluxRetoursSolubleAthos)

'calcul des boues (technosable) en sortie
Call calcul_boues_sortie_procede(BouesEpaissiesDCO_kgj, BouesEpaissiesNK_kgj, BouesEpaissiesP_kgj, RatioBouesMinerales, FluxRetoursTotauxAthos, TableauRecapitulatifFluxBoues, TableauRecapitulatifPollutionSolubleDansFluxBoues)


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Public Sub transfert_flux_energie_vers_pinch()


'******************************************************************************************************************************************************
'OBJECTIF
'   Cette procédure récapitule l'ensemble des flux du procédé à chauffer ou à refroidir
'   et appelle la procédure publique "copie_dans_flux_thermique_process" pour stocker les données dans une matrice

'VARIABLES D'ENTREE:
'   IndexFluxAMelanger2: index du flux à mélanger 2 (position dans la matrice "md_FluxIntermediaires")
'   IndexFluxResultant: index du flux résultant du mélange (position dans la matrice "md_FluxIntermediaires")
'   md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX): contient les flux à mélanger
'   les noms des flux de matière à chauffer ou refroidir qui se trouvent dans l'onglet Excel "Feuil5" nommé "Gestion_choix" dans des plages de cellule nommées dans Excel

'VARIABLES DE SORTIE:
'   md_FluxIntermediaires(0 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To mi_NOMBRE_CARACTERISTIQUES_FLUX): contient le flux résultant
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim NomFlux As String
Dim TemperatureEntree As Double
Dim TemperatureSortie As Double
Dim MCp_kWh_j°C As Double
Dim DebitMassique_kgj As Double
Dim CapaciteCalorifique_kWh_kg°C As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "53.11"
NomFlux = ""
TemperatureEntree = 0
TemperatureSortie = 0
MCp_kWh_j°C = 0
DebitMassique_kgj = 0
CapaciteCalorifique_kWh_kg°C = 0


'ATHOS - PRECHAUFFAGE BOUES
'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
NomFlux = Feuil5.Range("nom_flux_athos1").Cells(1, 1).Value
TemperatureEntree = md_FluxIntermediaires(mi_FLUX_MELANGE_ECHANGEUR1, mi_CARACTERISTIQUE_TEMPERATURE_°C)
TemperatureSortie = md_FluxIntermediaires(mi_FLUX_ECHANGEUR1_REACTEUR, mi_CARACTERISTIQUE_TEMPERATURE_°C)
DebitMassique_kgj = md_FluxIntermediaires(mi_FLUX_MELANGE_ECHANGEUR1, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)    ' débit en kg/j
CapaciteCalorifique_kWh_kg°C = (md_FluxIntermediaires(mi_FLUX_MELANGE_ECHANGEUR1, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) + md_FluxIntermediaires(mi_FLUX_ECHANGEUR1_REACTEUR, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C)) / 2     ' 'kWh/(kg.°C)
MCp_kWh_j°C = DebitMassique_kgj * CapaciteCalorifique_kWh_kg°C    'kW/°C

'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
Call copie_dans_flux_thermique_process(NomFlux, TemperatureEntree, TemperatureSortie, MCp_kWh_j°C)


'ATHOS - REFROIDISSEMENT BOUES
'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
NomFlux = Feuil5.Range("nom_flux_athos2").Cells(1, 1).Value
TemperatureEntree = md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_TEMPERATURE_°C)
TemperatureSortie = md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_TEMPERATURE_°C)
DebitMassique_kgj = md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj)    ' débit en kg/s
CapaciteCalorifique_kWh_kg°C = (md_FluxIntermediaires(mi_FLUX_REACTEUR_ECHANGEUR2, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C) + md_FluxIntermediaires(mi_FLUX_ECHANGEUR3_DECANTEUR, mi_CARACTERISTIQUE_CAPACITE_CALORIFIQUE_kWh_kg°C)) / 2     ' 'kJ/(kg.°C)
MCp_kWh_j°C = DebitMassique_kgj * CapaciteCalorifique_kWh_kg°C    'kW/°C

'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
Call copie_dans_flux_thermique_process(NomFlux, TemperatureEntree, TemperatureSortie, MCp_kWh_j°C)


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Public Sub calcul_consommation_electrique()


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure permet d'estimer la consommation électrique du procédé.

'VARIABLES D'ENTREE:
'   md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj): boues humides épaissies
'   md_RatioConsommationElectrique_kWh_m3BouesEpaissies: consommation électrique spécifique en fonction du flux volumique de boues humides épaissies
'   md_MasseVolumiqueBouesFroides_kg_m3: masse volumique des boues "froides" (entre 0 et 60°C)

'VARIABLES DE SORTIE:
'   pd_ConsommationElectrique_kWhj: consommation électrique journalière du procédé
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "53.7"
    
    
'pd_ConsommationElectrique_kWhj = md_EntreeBouesMelangeDebit_m3j / 1000 * ratio_elec
pd_ConsommationElectrique_kWhj = md_RatioConsommationElectrique_kWh_m3BouesEpaissies * md_FluxIntermediaires(mi_FLUX_EPAISSISSEMENT_MELANGE, mi_CARACTERISTIQUE_BOUES_HUMIDES_kgj) / md_MasseVolumiqueBouesFroides_kg_m3


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Private Sub ecriture_resultats()


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure permet simplement d'afficher les résultats des calculs dans Excel.

'VARIABLES D'ENTREE:
'   pi_EtapeCalculConsideree: étape de calcul (calcul des valeurs guides ou calcul avec prise en compte des valeurs forcées)
'   pi_FiliereConsideree: filière calculée
'   toutes les variables d'entrée facultatives utilisateur dont la valeur est affichée à l'utilisateur dans l'onglet "Calculs".

'VARIABLES DE SORTIE:
'   l'affichage dans Excel des variables d'entrée facultatives utilisateur et des variables de sortie utilisateur
'       qui se fait dans l'onglet "Feuil53" nommé "Athos" dans des plages de cellules dont les noms sont définis dans Excel
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim Ligne As Integer
Dim Colonne As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "53.9"
Ligne = 1
Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


For Ligne = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    Feuil53.Range("athos_inlet_ratio").Cells(Ligne, Colonne) = md_EntreeBouesPourcentageFluxAdmis(Ligne)
Next Ligne

Ligne = 1

Feuil53.Range("athos_inlet_MES").Cells(Ligne, Colonne) = md_EntreeBouesMelangeMES_kgj
Feuil53.Range("athos_inlet_Q").Cells(Ligne, Colonne) = md_EntreeBouesMelangeDebit_m3j
Feuil53.Range("athos_inlet_MV_MES").Cells(Ligne, Colonne) = md_EntreeBouesMelangeRatioMVMES
Feuil53.Range("athos_inlet_siccite").Cells(Ligne, Colonne) = md_EntreeBouesMelangeConcentrationMES_gL
Feuil53.Range("athos_inlet_temp").Cells(Ligne, Colonne) = md_EntreeBouesMelangeTemperature_°C
If mn_EpaississementBouesEntrees = True Then
    Feuil53.Range("athos_capture_MS").Cells(Ligne, Colonne) = md_EpaississementTauxCaptureMES
    Feuil53.Range("athos_conso_polymere").Cells(Ligne, Colonne) = md_EpaississementRatioConsommationPolymerePur_kg_tMS
    Feuil53.Range("athos_ep_siccite").Cells(Ligne, Colonne) = md_BouesEpaissiesConcentrationMES_gL
    Feuil53.Range("athos_ep_MES").Cells(Ligne, Colonne) = md_BouesEpaissiesMES_kgj
    Feuil53.Range("athos_ep_Q").Cells(Ligne, Colonne) = md_BouesEpaissiesDebit_m3j
Else
    Feuil53.Range("athos_capture_MS").Cells(Ligne, Colonne) = Feuil53.Range("athos_no_thickening").Cells(1, 1)
    Feuil53.Range("athos_conso_polymere").Cells(Ligne, Colonne) = Feuil53.Range("athos_no_thickening").Cells(1, 1)
    Feuil53.Range("athos_ep_siccite").Cells(Ligne, Colonne) = Feuil53.Range("athos_no_thickening").Cells(1, 1)
    Feuil53.Range("athos_ep_MES").Cells(Ligne, Colonne) = Feuil53.Range("athos_no_thickening").Cells(1, 1)
    Feuil53.Range("athos_ep_Q").Cells(Ligne, Colonne) = Feuil53.Range("athos_no_thickening").Cells(1, 1)
End If
Feuil53.Range("athos_reac_pression").Cells(Ligne, Colonne) = md_ReacteurPressionAbsolue_bars
Feuil53.Range("athos_reac_temp").Cells(Ligne, Colonne) = md_ReacteurTemperature_°C
Feuil53.Range("athos_reac_volume").Cells(Ligne, Colonne) = md_ReacteurVolume_m3
Feuil53.Range("athos_conso_O2").Cells(Ligne, Colonne) = pd_ConsommationOxygeneLiquidePur_kgj
Feuil53.Range("athos_conso_CuSO4").Cells(Ligne, Colonne) = pd_ConsommationCuSO4Pur_kgj
Feuil53.Range("athos_conso_combustible_traitement_RTO").Cells(Ligne, Colonne) = md_TraitementFumeesRTOConsommationCombustible_kWhPCIj
Feuil53.Range("athos_outlet_capture_ms").Cells(Ligne, Colonne) = md_DecanteurTauxCaptureMES
Feuil53.Range("athos_outlet_concentration").Cells(Ligne, Colonne) = md_DecanteurConcentrationMES_gL
Feuil53.Range("athos_outlet_MES").Cells(Ligne, Colonne) = md_SortieTechnosableMES_kgj
Feuil53.Range("athos_outlet_Q").Cells(Ligne, Colonne) = md_SortieTechnosableDebit_m3j
Feuil53.Range("athos_outlet_MV_MES").Cells(Ligne, Colonne) = md_SortieTechnosableRatioMVMES
Feuil53.Range("athos_electricite").Cells(Ligne, Colonne) = pd_ConsommationElectrique_kWhj



Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Public Sub effacer_resultats()


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure permet simplement d'effacer les résultats des calculs dans Excel (en écrivant "" - pas de .Delete).

'VARIABLES D'ENTREE:
'   pi_EtapeCalculConsideree: étape de calcul (calcul des valeurs guides ou calcul avec prise en compte des valeurs forcées)
'   pi_FiliereConsideree: filière calculée

'VARIABLES DE SORTIE:
'   l'affichage dans Excel de la valeur "" pour les variables d'entrée facultatives utilisateur et les variables de sortie utilisateur
'       qui se fait dans l'onglet "Feuil53" nommé "Athos" dans des plages de cellules dont les noms sont définis dans Excel
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim Ligne As Integer
Dim Colonne As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "53.10"
Ligne = 1
Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


For Ligne = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    Feuil53.Range("athos_inlet_ratio").Cells(Ligne, Colonne) = ""
Next Ligne

Ligne = 1

Feuil53.Range("athos_inlet_MES").Cells(Ligne, Colonne) = ""    'md_EntreeBouesMelangeMES_kgj
Feuil53.Range("athos_inlet_Q").Cells(Ligne, Colonne) = ""    'md_EntreeBouesMelangeDebit_m3j
Feuil53.Range("athos_inlet_MV_MES").Cells(Ligne, Colonne) = ""    'md_EntreeBouesMelangeRatioMVMES
Feuil53.Range("athos_inlet_siccite").Cells(Ligne, Colonne) = ""    'md_EntreeBouesMelangeConcentrationMES_gL
Feuil53.Range("athos_inlet_temp").Cells(Ligne, Colonne) = ""    'md_EntreeBouesMelangeTemperature_°C
Feuil53.Range("athos_capture_MS").Cells(Ligne, Colonne) = ""    'md_EpaississementTauxCaptureMES
Feuil53.Range("athos_conso_polymere").Cells(Ligne, Colonne) = ""    'md_EpaississementRatioConsommationPolymerePur_kg_tMS
Feuil53.Range("athos_ep_siccite").Cells(Ligne, Colonne) = ""    'md_BouesEpaissiesConcentrationMES_gL
Feuil53.Range("athos_ep_MES").Cells(Ligne, Colonne) = ""    'md_BouesEpaissiesMES_kgj
Feuil53.Range("athos_ep_Q").Cells(Ligne, Colonne) = ""    'md_BouesEpaissiesDebit_m3j
Feuil53.Range("athos_reac_pression").Cells(Ligne, Colonne) = ""    'md_ReacteurPressionAbsolue_bars
Feuil53.Range("athos_reac_temp").Cells(Ligne, Colonne) = ""    'md_ReacteurTemperature_°C
Feuil53.Range("athos_reac_volume").Cells(Ligne, Colonne) = ""    'md_ReacteurVolume_m3
Feuil53.Range("athos_conso_O2").Cells(Ligne, Colonne) = ""    'pd_ConsommationOxygeneLiquidePur_kgj
Feuil53.Range("athos_conso_CuSO4").Cells(Ligne, Colonne) = ""    'pd_ConsommationCuSO4Pur_kgj
Feuil53.Range("athos_conso_combustible_traitement_RTO").Cells(Ligne, Colonne) = ""    'md_TraitementFumeesRTOConsommationCombustible_kWhPCIj
Feuil53.Range("athos_outlet_capture_ms").Cells(Ligne, Colonne) = ""    'md_DecanteurTauxCaptureMES
Feuil53.Range("athos_outlet_concentration").Cells(Ligne, Colonne) = ""    'md_DecanteurConcentrationMES_gL
Feuil53.Range("athos_outlet_MES").Cells(Ligne, Colonne) = ""    'md_SortieTechnosableMES_kgj
Feuil53.Range("athos_outlet_Q").Cells(Ligne, Colonne) = ""    'md_SortieTechnosableDebit_m3j
Feuil53.Range("athos_outlet_MV_MES").Cells(Ligne, Colonne) = ""    'md_SortieTechnosableRatioMVMES
Feuil53.Range("athos_electricite").Cells(Ligne, Colonne) = ""    'pd_ConsommationElectrique_kWhj



Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True


End Sub
Private Sub incrementation_commentaires_calculs()

'***********************************************************************************************************************************
'OBJECTIF
'   Cette procédure sert à incrémenter la variable "ps_CommentairesCalculsGlobaux"
'   contenant l'ensemble des commentaires sur les calculs avec ceux concernant le procédé associé au module de classe.

'VARIABLES D'ENTREE:
'   ms_CommentairesCalculs(1 To mi_NOMBRE_COMMENTAIRES_CALCULS):vecteur contenant les commentaires sur le procédé en question
'   ps_CommentairesCalculsGlobaux: contient le cumul de tous les commentaires précédents

'VARIABLES DE SORTIE:
'   ps_CommentairesCalculsGlobaux: contient le cumul de tous les commentaires (précédents & procédé associé à ce module de classe
'***********************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "53.8"
I = 0

'Ajout du nom du procédé
ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil53.Range("athos_erreur_nom")

'Ajout des différents commentaires
If mi_NOMBRE_COMMENTAIRES_CALCULS > 0 Then
    For I = 1 To mi_NOMBRE_COMMENTAIRES_CALCULS
        If ms_CommentairesCalculs(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + ms_CommentairesCalculs(I)
        End If
    Next I
End If

'Saut de ligne dans la variable contenant l'ensemble des commentaires
ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True


End Sub

Private Sub Class_Initialize()


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure se lance automatiquement à l'instanciation du module de classe
'   Elle lance la subroutine "initialisation_variables_module" qui initialise les variables déclarées au niveau module

'VARIABLES D'ENTREE:

'VARIABLES DE SORTIE:
'******************************************************************************************************************************************************


'Initialisation des variables de niveau module
Call initialisation_variables_module


End Sub

Private Sub Class_Terminate()

If pn_EffacementResultatsEnCours = False Then
    Call ecriture_resultats
    Call incrementation_commentaires_calculs
End If

End Sub











