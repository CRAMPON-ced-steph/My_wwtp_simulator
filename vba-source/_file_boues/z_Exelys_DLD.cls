Attribute VB_Name = "z_Exelys_DLD"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'Constantes pour repérer la position d'entrée des boues
Private Const mi_POSITION_ENTREE_BOUES_DIGESTEUR1 As Integer = 1
Private Const mi_POSITION_ENTREE_BOUES_NON_DIGEREES As Integer = 2

'constantes pour le choix d'agitation
Private Const mi_AGITATION_MECANIQUE As Integer = 1
Private Const mi_AGITATION_RECIRCULATION As Integer = 2
Private Const mi_AGITATION_BIOGAZ As Integer = 3

'constantes pour le choix du chauffage des boues
Private Const mi_DIGESTEUR1_CHAUFFAGE_BOUES_AMONT As Integer = 1
Private Const mi_DIGESTEUR1_CHAUFFAGE_BOUES_RECIRCULATION As Integer = 2

'constantes pour l'isolation
Private Const mi_DIGESTION_ISOLEE As Integer = 1
Private Const mi_DIGESTION_NON_ISOLEE As Integer = 2

'constantes pour le stockage de biogaz
Private Const mi_STOCKAGE_BIOGAZ As Integer = 1
Private Const mi_PAS_DE_STOCKAGE_BIOGAZ As Integer = 2

'constantes nécessaires pour l'entrée
Private Const mi_FLUX_ENTREE_BOUES_NON_UTILISE As Integer = 0
Private Const mi_FLUX_ENTREE_BOUES_I As Integer = 1
Private Const mi_FLUX_ENTREE_BOUES_II As Integer = 2
Private Const mi_FLUX_ENTREE_BOUES_III As Integer = 3
Private Const mi_FLUX_ENTREE_BOUES_EXTERIEURES_1 As Integer = 4
Private Const mi_FLUX_ENTREE_BOUES_EXTERIEURES_2 As Integer = 5
Private Const mi_FLUX_ENTREE_GRAISSES As Integer = 6
Private Const mi_FLUX_ENTREE_EPAISSISSEUR_1 As Integer = 7
Private Const mi_FLUX_ENTREE_EPAISSISSEUR_2 As Integer = 8
Private Const mi_FLUX_ENTREE_EPAISSISSEUR_3 As Integer = 9
Private Const mi_FLUX_ENTREE_EPAISSISSEUR_4 As Integer = 10

Private Const mi_NOMBRE_FLUX_ENTREE_BOUES As Integer = 3

Private Const mi_CRITERE_INCOHERENCE As Double = 0.1

'paramètres des boues intermediaires
Private Const mi_NOMBRE_PARAMETRES_BOUES As Integer = 13
Private Const mi_PARAMETRE_BOUES_MS_kgj As Integer = 1
Private Const mi_PARAMETRE_BOUES_Q_m3j As Integer = 2
Private Const mi_PARAMETRE_BOUES_CONCENTRATION_gL As Integer = 3
Private Const mi_PARAMETRE_BOUES_RATIO_MV_MES As Integer = 4
Private Const mi_PARAMETRE_BOUES_TEMPERATURE_°C As Integer = 5
Private Const mi_PARAMETRE_BOUES_RATIO_DCO_MES As Integer = 6
Private Const mi_PARAMETRE_BOUES_RATIO_DBO_MES As Integer = 7
Private Const mi_PARAMETRE_BOUES_RATIO_NK_MES As Integer = 8
Private Const mi_PARAMETRE_BOUES_RATIO_P_MES As Integer = 9
Private Const mi_PARAMETRE_BOUES_REDUCTION_MV As Integer = 10
Private Const mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION As Integer = 11
Private Const mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT As Integer = 12
Private Const mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG As Integer = 13

'paramètres des digesteurs
Private Const mi_NOMBRE_PARAMETRE_DIGESTION As Integer = 7
Private Const mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j As Integer = 1
Private Const mi_PARAMETRE_DIGESTION_TEMPERATURE_°C As Integer = 2
Private Const mi_PARAMETRE_DIGESTION_VOLUME_m3 As Integer = 3
Private Const mi_PARAMETRE_DIGESTION_HAUTEUR_m As Integer = 4
Private Const mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j As Integer = 5
Private Const mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_REEL_j As Integer = 6
Private Const mi_PARAMETRE_DIGESTION_REDUCTION_MV As Integer = 7

'pompage des boues
Private Const mi_NOMBRE_POMPAGE As Integer = 3
Private Const mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE As Integer = 1
Private Const mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR As Integer = 2
Private Const mi_POMPAGE_SORTIE_DIGESTEUR2 As Integer = 3

'paramètres des pompes
Private Const mi_NOMBRE_PARAMETRES_POMPAGE As Integer = 5
Private Const mi_PARAMETRE_POMPAGE_NOMBRE_POMPES As Integer = 1
Private Const mi_PARAMETRE_POMPAGE_HMT_mCE As Integer = 2
Private Const mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj As Integer = 3
Private Const mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h As Integer = 4
Private Const mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES As Integer = 5

'flux intermédiaires
Private Const mi_MELANGE_BOUES As Integer = nb_type_boues + 1
Private Const mi_NOMBRE_FLUX_INTERMEDIAIRES As Integer = 10
Private Const mi_FLUX_ENTREE_DIGESTEUR1 As Integer = 1
Private Const mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE As Integer = 2
Private Const mi_FLUX_CENTRIFUGEUSE_MELANGEUR As Integer = 3
Private Const mi_FLUX_BOUES_NON_DIGEREES As Integer = 4
Private Const mi_FLUX_MELANGEUR_LYSE As Integer = 5
Private Const mi_FLUX_LYSE_ECHANGEUR90 As Integer = 6
Private Const mi_FLUX_ECHANGEUR90_ECHANGEUR55 As Integer = 7
Private Const mi_FLUX_ECHANGEUR55_DILUTION As Integer = 8
Private Const mi_FLUX_DILUTION_DIGESTEUR2 As Integer = 9
Private Const mi_FLUX_SORTIE_DIGESTEUR2 As Integer = 10

'paramètres centrif
Private Const mi_NOMBRE_PARAMETRE_CENTRIFUGEUSE As Integer = 2
Private Const mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES As Integer = 1
Private Const mi_PARAMETRE_CENTRIFUGEUSE_CONSOMMATION_POLYMERE_kgMA_tMS As Integer = 2

'variables de choix
Private mi_ProvenanceBouesEntree(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Integer
Private mi_TypeBouesEntree(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Integer
Private mi_ChoixEntreesBoues(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Integer
Private mi_ChoixPositionEntreeBoues(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Integer
Private mi_ChoixTypeAgitationDigesteur1 As Integer
Private mi_ChoixTypeAgitationDigesteur2 As Integer
Public pi_ChoixTypePolymereCentrifugeuse As Integer
Private mi_ChoixStockageBiogaz As Integer
Private mi_ChoixIsolationDigestion As Integer
Private mi_ChoixTypeChauffageDigesteur1 As Integer

'variables à échanger avec l'utilisateur
Private md_EntreeBouesPourcentageFluxAdmis(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Double
Private md_FluxIntermediaires(1 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To nb_type_boues + 1, 1 To mi_NOMBRE_PARAMETRES_BOUES) As Double
Private md_CaracteristiquesDigesteur1(1 To mi_NOMBRE_PARAMETRE_DIGESTION) As Double
Private md_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C As Double
Private md_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C As Double
Private md_CaracteristiquesDigesteur2(1 To mi_NOMBRE_PARAMETRE_DIGESTION) As Double
Private md_CaracteristiquePompageBoues(1 To mi_NOMBRE_POMPAGE, 1 To mi_NOMBRE_PARAMETRES_POMPAGE) As Double
Private md_CaracteristiqueCentrifugeuse(1 To mi_NOMBRE_PARAMETRE_CENTRIFUGEUSE) As Double
Public pd_CentrifugeuseConsommationPolymerePur_kgj As Double
Private md_DebitVapeurLyse_kgj As Double
Private md_PressionAbsolueVapeurSaturanteLyse_bar As Double
Private md_TemperatureInitialeEauPourProductionVapeur_°C As Double
Private md_TemperaturePrechauffageEauPourProductionVapeur_°C As Double
Private md_DebitEauDilution_m3j As Double
Private md_TemperatureEauDilution_°C As Double
Private md_TauxReductionGlobalMV As Double
Private md_RatioCH4DansBiogaz As Double
Private md_ProductionSpecifiqueBiogaz_Nm3_kgMV As Double
Public pd_BiogazProduit_Nm3j As Double
Private md_EnergieChauffageDigesteur1_kWhj As Double
Private md_BesoinTotalEnergieProductionVapeur_kWhj As Double
Private md_EnergieRefroidissementBouesLysees_kWhj As Double
Private md_EnergieDisponibleEau90°C_kWhj As Double
Public pd_ConsommationElectrique_kWhj As Double
Public pd_ConsommationElectriqueFixe_kWhj As Double

'Valeur pour le calcul des KPI
Public pd_MatieresVolatilesBouesEntrees_kgMVj As Double

'valeurs de production de biogaz pour les co-ferments
Private md_PotentielProductionMethaneCoferments_Nm3_tMVAppliques As Double
Private md_FacteurSecuritePotentielProductionMethaneCoferments As Double

'valeurs forcées dans le tableur
Private mn_TemperatureBouesEntree_°C_force As Boolean
Private mn_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C_force As Boolean
Private mn_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C_force As Boolean
Private mn_ConcentrationBouesCentrifugeuseMelangeur_gL_force As Boolean
Private mn_BouesCentrifugeuseMelangeurRatioMVSurMES_force As Boolean
Private mn_TemperatureBouesNonDigerees_°C_force As Boolean
'Private mn_TemperatureBouesMelangeurLyse_°C_force As Boolean
Private mn_TemperatureBouesLyseEchangeur90_°C_force As Boolean
Private mn_ConcentrationBouesDilutionDigesteur2_gL_force As Boolean
Private mn_CaracteristiquesDigesteur1_force(1 To mi_NOMBRE_PARAMETRE_DIGESTION) As Boolean
Private mn_CaracteristiquesDigesteur2_force(1 To mi_NOMBRE_PARAMETRE_DIGESTION) As Boolean
Private mn_CaracteristiquePompageBoues_force(1 To mi_NOMBRE_POMPAGE, mi_NOMBRE_PARAMETRES_POMPAGE) As Boolean
Private mn_CaracteristiqueCentrifugeuse_force(1 To mi_NOMBRE_PARAMETRE_CENTRIFUGEUSE) As Boolean
Private mn_TemperatureEauDilution_°C_force As Boolean
Private mn_DebitVapeurLyse_kgj_force As Boolean
Private mn_PressionAbsolueVapeurSaturanteLyse_bar_force As Boolean
Private mn_TemperatureInitialeEauPourProductionVapeur_°C_force As Boolean
Private mn_TemperaturePrechauffageEauPourProductionVapeur_°C_force As Boolean
Private mn_RatioCH4DansBiogaz_force As Boolean
Private mn_BiogazProduit_Nm3j_force As Boolean
Private mn_EnergieChauffageDigesteur1_kWhj_force As Boolean
Private mn_BesoinTotalEnergieProductionVapeur_kWhj_force As Boolean
Private mn_EnergieRefroidissementBouesLysees_kWhj_force As Boolean
Private mn_EnergieDisponibleEau90°C_kWhj_force As Boolean

'valeurs forcées dans le tableur
Private md_TemperatureBouesEntree_°C_f As Double
Private md_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C_f As Double
Private md_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C_f As Double
Private md_ConcentrationBouesCentrifugeuseMelangeur_gL_f As Double
Private md_BouesCentrifugeuseMelangeurRatioMVSurMES_f As Double
Private md_TemperatureBouesNonDigerees_°C_f As Double
'Private md_TemperatureBouesMelangeurLyse_°C_f As Double
Private md_TemperatureBouesLyseEchangeur90_°C_f As Double
Private md_ConcentrationBouesDilutionDigesteur2_gL_f As Double
Private md_CaracteristiquesDigesteur1_f(1 To mi_NOMBRE_PARAMETRE_DIGESTION) As Double
Private md_CaracteristiquesDigesteur2_f(1 To mi_NOMBRE_PARAMETRE_DIGESTION) As Double
Private md_CaracteristiquePompageBoues_f(1 To mi_NOMBRE_POMPAGE, mi_NOMBRE_PARAMETRES_POMPAGE) As Double
Private md_CaracteristiqueCentrifugeuse_f(1 To mi_NOMBRE_PARAMETRE_CENTRIFUGEUSE) As Double
Private md_DebitVapeurLyse_kgj_f As Double
Private md_TemperatureEauDilution_°C_f As Double
Private md_PressionAbsolueVapeurSaturanteLyse_bar_f As Double
Private md_TemperatureInitialeEauPourProductionVapeur_°C_f As Double
Private md_TemperaturePrechauffageEauPourProductionVapeur_°C_f As Double
Private md_RatioCH4DansBiogaz_f As Double
Private md_BiogazProduit_Nm3j_f As Double
Private md_EnergieChauffageDigesteur1_kWhj_f As Double
Private md_BesoinTotalEnergieProductionVapeur_kWhj_f As Double
Private md_EnergieRefroidissementBouesLysees_kWhj_f As Double
Private md_EnergieDisponibleEau90°C_kWhj_f As Double

'hypothèses
Private md_ReductionMatieresVolatilesBouesNonLysees(1 To nb_qualites_boues) As Double
Private md_ReductionMatieresVolatilesBouesLysees(1 To nb_qualites_boues) As Double
Private md_ReductionMatieresVolatilesBouesDigereesLysees As Double
Private md_TemperatureReferenceDigesteur1_°C As Double
Private md_TemperatureReferenceDigesteur2_°C As Double
Private md_ReferenceTempsSejourBouesDigesteur1_j As Double
Private md_ReferenceTempsSejourBouesDigesteur2_j As Double
Private md_CoefficientsCalculDensiteBiogaz(0 To 1) As Double
Private md_ProductionSpecifiqueMethane_Nm3_kgDCO As Double
Private md_RatioNKSolubleBouesDigereesSurMVDegrades As Double
Private md_RatioPSolubleBouesDigereesSurPBouesEntree As Double
Private md_RatioDCOSolubleSurDCOTotaleBouesDigereesDigesteur1 As Double
Private md_RatioDCOSolubleSurDCOTotaleBouesDigereesDigesteur2 As Double
Private md_RatioMVParticulairesMVTotauxBouesDigereesDigesteur1 As Double
Private md_PerteTemperatureBouesDigereesDigesteur1AvantLyse_°C As Double
Private md_PertesThermiquesDigestion_kWh_m3j(1 To climat_type_nb, 1 To mi_DIGESTION_NON_ISOLEE) As Double
Private md_EfficaciteEchangeur As Double
Private md_CapaciteCalorifiqueBoues_kWh_m3°C As Double
Private md_ConcentrationPolymereCentrifugeuse_gL As Double
Private md_EnthalpieTotaleVapeurSaturee_kcal_kg As Double
Private md_CapaciteCalorifiqueBoues_kcal_kg°C As Double
Private md_RatioPertesVapeur As Double
Private md_CapaciteCalorifiqueEau_kcal_kg°C As Double
Private md_RatioNH4SurNKSolubleBouesDigerees As Double
Private md_RatioDBOSolubleSurDCOSolubleBouesDigerees As Double
Private md_MasseVolumiqueBoues_kg_m3 As Double
Private md_CoefficientsCalculTemperatureSaturationAPression(1 To 5) As Double
Private md_CoefficientsCalculEnthalpieVaporisationATemperature(0 To 1) As Double
Private md_TemperatureSaturation_°C As Double
Private md_EnthalpieVaporisationATemperature_kcal_kg As Double
Private md_TemperatureBouesEchangeur90Echangeur55_°C As Double
Private md_TemperatureLimiteEntreMesophileEtThermophile_°C As Double
Private md_EcartMinimalTemperatureEchangeur_°C As Double
Private md_CoefficientACalculFacteurCorrectionTempsSejourBouesDigesteur As Double
Private md_CoefficientBCalculFacteurCorrectionTempsSejourBouesDigesteur As Double
Private md_RatioMVSolublesMVTotauxBouesDigereesReference(1 To nb_qualites_boues) As Double
Private md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(1 To nb_qualites_boues) As Double
Private md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(1 To nb_qualites_boues) As Double
Private md_EcartTemperatureBouesEchangeurDigesteur_°C As Double
Private md_PinchChauffageEauPourVapeur_kWhj As Double
Private md_PinchVaporisation_kWhj As Double
Private md_ConcentrationBouesDilutionDigesteur2Souhaitee_gL As Double
Private md_EnergieChauffageBouesFraichesDigesteur1SansPertes_kWhj As Double
Private md_ConsommationElectriqueSpecifiqueAgitationMecaniqueDigestion_W_m3 As Double
Private md_TempsFonctionnementAgitation_hj As Double
Private md_TempsFonctionnementRecirculationBoues_hj As Double
Private md_PressionRefoulementRecirculationBoues_mCE As Double
Private md_MasseVolumiqueEau_kg_m3 As Double

'stockage de données
Private md_BiogazProduitNominal_Nm3j As Double
Private md_DebitPolymere_m3j As Double

'erreurs
Private Const mi_NOMBRE_COMMENTAIRES_CALCULS As Integer = 6
Private ms_CommentairesCalculs() As String

Private Sub forcage_rendement_MV(ByRef CaracteristiquesDigestion() As Double, ByRef CaracteristiquesDigestion_f() As Double, ByVal RepereFluxBouesSortieDigesteur As Integer)

'************************************************************************************************************
'OBJECTIF:
'   Cette procédure gère le renseignement de la réduction de MV par l'utilisateur
'A COMPLETER
'************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim DeltaReductionMV_kgj As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "27.33"
I = 0
DeltaReductionMV_kgj = 0


'On calcule l'écart de réduction des MV
DeltaReductionMV_kgj = (CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_REDUCTION_MV) / CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_REDUCTION_MV) - 1) * md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV)
CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_REDUCTION_MV) = CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_REDUCTION_MV)

'On recalcule les valeurs pour le mélange de boues
md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) - DeltaReductionMV_kgj / md_MasseVolumiqueBoues_kg_m3
md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) - DeltaReductionMV_kgj
md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) + DeltaReductionMV_kgj
md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) - DeltaReductionMV_kgj
md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)

'On recalcule les valeurs pour chaque type de boues
For I = 1 To nb_type_boues
    If md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) <> 0 Then
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, I, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, I, mi_PARAMETRE_BOUES_MS_kgj) - md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, I, mi_PARAMETRE_BOUES_REDUCTION_MV) * DeltaReductionMV_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) - md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, I, mi_PARAMETRE_BOUES_REDUCTION_MV) * DeltaReductionMV_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, I, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, I, mi_PARAMETRE_BOUES_Q_m3j) - md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, I, mi_PARAMETRE_BOUES_REDUCTION_MV) * DeltaReductionMV_kgj / 1000
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, I, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, I, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, I, mi_PARAMETRE_BOUES_Q_m3j)
    End If
Next I


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Public Sub dimensionnement(ByRef TableauRecapitulatifFluxBoues() As Double, ByRef PartGraissesInternesSurTotalCosubstrat() As Double, ByVal ChoixEauxSalesPrimaire As Boolean, ByRef DetailBouesI() As Double, ByRef DetailBouesII() As Double, ByRef DetailBouesIII() As Double, ByRef DCO_biogaz, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double, ByRef FluxRetoursTotaux() As Double, ByRef FluxRetoursTotauxDigestion() As Double, ByRef FluxRetoursSoluble() As Double, ByRef FluxRetoursSolubleDigestion() As Double, ByRef TableauBesoinsThermiques() As Double, ByRef TableauDisponibilitesThermiques() As Double, ByRef CaracteristiquesBiogaz() As Double, ByRef PressionVapeur_bar() As Double, ByRef NombreReseauxVapeur As Integer, ByRef LyseAvalEpaississement() As Integer)


'*********************************************************************************************************************************************************
'OBJECTIF:

'VARIABLES D'ENTREE:
'   TableauRecapitulatifFluxBoues(): contient les caractéristiques des boues alimentant le procédé
'   TableauRecapitulatifPollutionSolubleDansFluxBoues(): contient la polution soluble des boues alimentant le procédé

'VARIABLES DE SORTIE:
'   TableauRecapitulatifFluxBoues(): contient les caractéristiques des boues en sortie du procédé
'   TableauRecapitulatifPollutionSolubleDansFluxBoues(): contient la polution soluble des boues en sortie du procédé
'A COMPLETER
'*********************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer
Dim k As Integer
Dim ReperageBoucleReel As Boolean
Dim IncoherenceCalcul As Double
Dim BiogazProduitDigesteur1_Nm3j As Double
Dim FluxDCODansBiogazDigesteur1_kgj As Double
Dim BiogazProduitDigesteur2_Nm3j As Double
Dim FluxDCODansBiogazDigesteur2_kgj As Double
Dim RatioTempsSejourSurReference As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "27.5"
I = 0
j = 0
k = 0
ReperageBoucleReel = False
IncoherenceCalcul = 0
BiogazProduitDigesteur1_Nm3j = 0
FluxDCODansBiogazDigesteur1_kgj = 0
BiogazProduitDigesteur2_Nm3j = 0
FluxDCODansBiogazDigesteur2_kgj = 0
RatioTempsSejourSurReference = 0


'RE-INITIALISATION DES FLUX
For I = 1 To mi_NOMBRE_FLUX_INTERMEDIAIRES
    For j = 1 To nb_type_boues + 1
        For k = 1 To mi_NOMBRE_PARAMETRES_BOUES
            If k <> mi_PARAMETRE_BOUES_TEMPERATURE_°C And k <> mi_PARAMETRE_BOUES_CONCENTRATION_gL Then
                md_FluxIntermediaires(I, j, k) = 0
            End If
        Next k
    Next j
Next I

'température des boues en entrée de digesteur 1
If mn_TemperatureBouesEntree_°C_force = True Then
    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureBouesEntree_°C_f
Else
    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = pd_TemperatureEauDimensionnement_°C
End If
'température des boues non digérées
If mn_TemperatureBouesNonDigerees_°C_force = True Then
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureBouesNonDigerees_°C_f
Else
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = pd_TemperatureEauDimensionnement_°C
End If
'température de l'eau de dilution
If mn_TemperatureEauDilution_°C_force = True Then
    md_TemperatureEauDilution_°C = md_TemperatureEauDilution_°C_f
Else
    md_TemperatureEauDilution_°C = pd_TemperatureEauDimensionnement_°C
End If

'CALCUL DES BOUES EN ENTREE (VERS LYSE ET BYPASS)
Call calcul_cocktail_boues_entrees(TableauRecapitulatifFluxBoues, TableauRecapitulatifPollutionSolubleDansFluxBoues, ChoixEauxSalesPrimaire, DetailBouesI, DetailBouesII, DetailBouesIII, PartGraissesInternesSurTotalCosubstrat, LyseAvalEpaississement)

'DIGESTEUR 1
Call calcul_bilan_masse_digestion(md_CaracteristiquesDigesteur1, md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j), md_ReferenceTempsSejourBouesDigesteur1_j, mi_FLUX_ENTREE_DIGESTEUR1, mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, RatioTempsSejourSurReference)

'calcul du volume du digesteur 1
md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_VOLUME_m3) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j)
If mn_CaracteristiquesDigesteur1_force(mi_PARAMETRE_DIGESTION_VOLUME_m3) = True Then
    IncoherenceCalcul = (md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_VOLUME_m3) - md_CaracteristiquesDigesteur1_f(mi_PARAMETRE_DIGESTION_VOLUME_m3)) / md_CaracteristiquesDigesteur1_f(mi_PARAMETRE_DIGESTION_VOLUME_m3)
    If IncoherenceCalcul > mi_CRITERE_INCOHERENCE Then
        ms_CommentairesCalculs(1) = Feuil27.Range("dig_DLD_erreur_volume1").Cells(1, 1)
    End If
    md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_VOLUME_m3) = md_CaracteristiquesDigesteur1_f(mi_PARAMETRE_DIGESTION_VOLUME_m3)
End If

'calcul de la part de MV soluble en sortie
Call calcul_MV_soluble_boues_digerees(RatioTempsSejourSurReference)

'calcul du biogaz produit dans digesteur 1
Call calcul_biogaz(BiogazProduitDigesteur1_Nm3j, FluxDCODansBiogazDigesteur1_kgj, mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE)

'calcul de la pollution solubilisée lors de la digestion
Call calcul_solubilisation_pendant_digestion(TableauRecapitulatifPollutionSolubleDansFluxBoues, md_CaracteristiquesDigesteur1, mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, md_RatioDCOSolubleSurDCOTotaleBouesDigereesDigesteur1)

'On attribue les nouveaux potentiels de dégradation des MV
For I = 1 To nb_type_boues
    md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_ReductionMatieresVolatilesBouesDigereesLysees
Next I

'CENTRIFUGEUSE
Call calcul_bilan_masse_centrifugeuse(TableauRecapitulatifPollutionSolubleDansFluxBoues, FluxRetoursTotaux, FluxRetoursSoluble, FluxRetoursTotauxDigestion, FluxRetoursSolubleDigestion)

'ENTREE DES BOUES NON DIGEREES
Call calcul_bilan_masse_melangeur(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_FLUX_BOUES_NON_DIGEREES, mi_FLUX_MELANGEUR_LYSE)

'HYDROLYSE THERMIQUE
Call calcul_bilan_masse_lyse(ReperageBoucleReel, TableauRecapitulatifPollutionSolubleDansFluxBoues)

'ECHANGEURS --> SEULEMENT LORS DES BILANS THERMIQUES ET ELECTRIQUES

'DILUTION AVEC DE L'EAU
Call calcul_dilution_boues_avant_digestion(ReperageBoucleReel, TableauRecapitulatifPollutionSolubleDansFluxBoues)

'DIGESTEUR 2
Call calcul_bilan_masse_digestion(md_CaracteristiquesDigesteur2, md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j), md_ReferenceTempsSejourBouesDigesteur2_j, mi_FLUX_DILUTION_DIGESTEUR2, mi_FLUX_SORTIE_DIGESTEUR2, RatioTempsSejourSurReference)

'calcul du volume du digesteur 2
md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_VOLUME_m3) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j)
If mn_CaracteristiquesDigesteur2_force(mi_PARAMETRE_DIGESTION_VOLUME_m3) = True Then
    IncoherenceCalcul = (md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_VOLUME_m3) - md_CaracteristiquesDigesteur2_f(mi_PARAMETRE_DIGESTION_VOLUME_m3)) / md_CaracteristiquesDigesteur2_f(mi_PARAMETRE_DIGESTION_VOLUME_m3)
    If IncoherenceCalcul > mi_CRITERE_INCOHERENCE Then
        ms_CommentairesCalculs(2) = Feuil27.Range("dig_DLD_erreur_volume2").Cells(1, 1)
    End If
    md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_VOLUME_m3) = md_CaracteristiquesDigesteur2_f(mi_PARAMETRE_DIGESTION_VOLUME_m3)
End If

'biogaz produit dans digesteur 2
Call calcul_biogaz(BiogazProduitDigesteur2_Nm3j, FluxDCODansBiogazDigesteur2_kgj, mi_FLUX_SORTIE_DIGESTEUR2)
pd_BiogazProduit_Nm3j = BiogazProduitDigesteur1_Nm3j + BiogazProduitDigesteur2_Nm3j

'calcul de la pollution solubilisée lors de la digestion
Call calcul_solubilisation_pendant_digestion(TableauRecapitulatifPollutionSolubleDansFluxBoues, md_CaracteristiquesDigesteur2, mi_FLUX_SORTIE_DIGESTEUR2, md_RatioDCOSolubleSurDCOTotaleBouesDigereesDigesteur2)

'calcul de la réduction globale (ici sont aussi inclus les MV retournés en tête)
md_TauxReductionGlobalMV = (md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) * md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) - md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) + md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) - md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj))
md_TauxReductionGlobalMV = md_TauxReductionGlobalMV / (md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) * md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) * md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES))

'GESTION ENERGIE
Call calcul_bilan_thermique(TableauBesoinsThermiques, TableauDisponibilitesThermiques, PressionVapeur_bar, NombreReseauxVapeur, ReperageBoucleReel)
CaracteristiquesBiogaz(Q_Nm3j) = pd_BiogazProduit_Nm3j
CaracteristiquesBiogaz(teneur_CH4) = md_RatioCH4DansBiogaz

'RECAPITULATIF SORTIE
For I = 1 To nb_type_boues
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_MS_kgj)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_Q) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_Q_m3j)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_MV_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_MV_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_DCO_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_DBO_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_NK_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_NK_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_Pt_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_P_MES)
Next I



Exit Sub
'*********************************************************************
erreur_non_geree:
NumeroErreurNonGeree = "27.5"

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Private Sub calcul_solubilisation_pendant_digestion(ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double, ByRef CaracteristiquesDigestion() As Double, ByVal RepereFluxBouesDigerees As Integer, ByVal RatioDCOSolubleSurDCOTotaleBouesDigerees As Double)

'************************************************************************************************************
'OBJECTIF:
'   Cette procédure effectue le bilan masse lors du mélange des boues lysées avec les boues fraîches
'   (si nécessaire)
'************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim FluxSolubilise_kgj As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "27.28"
I = 0
FluxSolubilise_kgj = 0

'calcul de la pollution solubilisée lors de la digestion
For I = 1 To nb_type_boues
    If md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_MS_kgj) <> 0 Then
        'DCO
        FluxSolubilise_kgj = RatioDCOSolubleSurDCOTotaleBouesDigerees * md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) * F_facteur_correction_temperature_digesteur(CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C))
        TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_DCO) = FluxSolubilise_kgj / md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_Q_m3j) * 1000
        md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = (md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) - FluxSolubilise_kgj) / md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_MS_kgj)
        'DBO
        md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = 0    'hypothèse
        'NK
        FluxSolubilise_kgj = md_RatioNKSolubleBouesDigereesSurMVDegrades * md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_REDUCTION_MV) * md_FluxIntermediaires(RepereFluxBouesDigerees, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) * md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_NK_MES) / md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_MV_MES)
        TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_NK) = FluxSolubilise_kgj / md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_Q_m3j) * 1000
        md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_NK_MES) = (md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_NK_MES) - FluxSolubilise_kgj) / md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_MS_kgj)
        'Pt
        FluxSolubilise_kgj = md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_P_MES) * md_RatioPSolubleBouesDigereesSurPBouesEntree * F_facteur_correction_temperature_digesteur(CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C))
        TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_Pt) = FluxSolubilise_kgj / md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_Q_m3j) * 1000
        md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_P_MES) = (md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_P_MES) - FluxSolubilise_kgj) / md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_MS_kgj)
        'On recalcule bien les MV
        md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) / md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_MS_kgj)
    End If
Next I

md_FluxIntermediaires(RepereFluxBouesDigerees, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(RepereFluxBouesDigerees, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) / md_FluxIntermediaires(RepereFluxBouesDigerees, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj)


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Private Sub calcul_cocktail_boues_entrees(ByRef TableauRecapitulatifFluxBoues() As Double, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double, ByVal ChoixEauxSalesPrimaire As Boolean, ByRef DetailBouesI() As Double, ByRef DetailBouesII() As Double, ByRef DetailBouesIII() As Double, ByRef PartGraissesInternesSurTotalCosubstrat() As Double, ByRef LyseAvalEpaississement() As Integer)


'************************************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure permet de mélanger les différents types de boues en entrée de lyse et en entrée de mélangeur si certaines by-passent la lyse.
'   Tout est exprimé en flux dans un premier temps lors de la grande boucle puis on reconstituera les ratios qui vont bien à la fin
'************************************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer
Dim k As Integer
Dim BouesEntreeNonDigereesMV_kgj As Double
Dim BouesEntreeNonDigereesMES_kgj As Double
Dim BouesEntreeNonDigereesMS_kgj As Double
Dim BouesEntreeNonDigereesNK_kgj As Double
Dim BouesEntreeNonDigereesP_kgj As Double
Dim BouesEntreeNonDigereesDCO_kgj As Double
Dim BouesEntreeNonDigereesDBO_kgj As Double
Dim BouesEntreeNonDigereesDebitVolumique_m3j As Double
Dim BouesEntreeNonDigereesReductionMV_kgj As Double
Dim BouesEntreeDigesteur1MV_kgj As Double
Dim BouesEntreeDigesteur1MES_kgj As Double
Dim BouesEntreeDigesteur1MS_kgj As Double
Dim BouesEntreeDigesteur1NK_kgj As Double
Dim BouesEntreeDigesteur1P_kgj As Double
Dim BouesEntreeDigesteur1DCO_kgj As Double
Dim BouesEntreeDigesteur1DBO_kgj As Double
Dim BouesEntreeDigesteur1DebitVolumique_m3j As Double
Dim BouesEntreeDigesteur1ReductionMV_kgj As Double
Dim IntermediaireMV_kgj As Double
Dim IntermediaireMES_kgj As Double
Dim IntermediaireMS_kgj As Double
Dim IntermediaireNK_kgj As Double
Dim IntermediaireP_kgj As Double
Dim IntermediaireDCO_kgj As Double
Dim IntermediaireDBO_kgj As Double
Dim IntermediaireDebitVolumique_m3j As Double
Dim IntermediaireReductionMV_kgj As Double
Dim PartBouesBiologiques As Double
Dim CoefficientReductionMV As Double
Dim CofermentsCoefficientReductionMV As Double
Dim IntermediaireReductionMVSoluble_kgj As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "27.26"
I = 0
j = 0
k = 0
BouesEntreeNonDigereesMV_kgj = 0
BouesEntreeNonDigereesMES_kgj = 0
BouesEntreeNonDigereesMS_kgj = 0
BouesEntreeNonDigereesNK_kgj = 0
BouesEntreeNonDigereesP_kgj = 0
BouesEntreeNonDigereesDCO_kgj = 0
BouesEntreeNonDigereesDBO_kgj = 0
BouesEntreeNonDigereesDebitVolumique_m3j = 0
BouesEntreeNonDigereesReductionMV_kgj = 0
BouesEntreeDigesteur1MV_kgj = 0
BouesEntreeDigesteur1MES_kgj = 0
BouesEntreeDigesteur1MS_kgj = 0
BouesEntreeDigesteur1NK_kgj = 0
BouesEntreeDigesteur1P_kgj = 0
BouesEntreeDigesteur1DCO_kgj = 0
BouesEntreeDigesteur1DBO_kgj = 0
BouesEntreeDigesteur1DebitVolumique_m3j = 0
BouesEntreeDigesteur1ReductionMV_kgj = 0
IntermediaireMV_kgj = 0
IntermediaireMES_kgj = 0
IntermediaireMS_kgj = 0
IntermediaireNK_kgj = 0
IntermediaireP_kgj = 0
IntermediaireDCO_kgj = 0
IntermediaireDBO_kgj = 0
IntermediaireDebitVolumique_m3j = 0
IntermediaireReductionMV_kgj = 0
PartBouesBiologiques = 0
CoefficientReductionMV = 0
CofermentsCoefficientReductionMV = 0
IntermediaireReductionMVSoluble_kgj = 0

'On itère sur les différents flux en entrée
For I = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    If mi_ChoixEntreesBoues(I) <> mi_FLUX_ENTREE_BOUES_NON_UTILISE Then
        If mi_ProvenanceBouesEntree(I) = boues_inlet Then
            'variable d'entrée
            IntermediaireMES_kgj = md_EntreeBouesPourcentageFluxAdmis(I) * TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_MES)
            If IntermediaireMES_kgj <> 0 Then
                IntermediaireMS_kgj = IntermediaireMES_kgj
                IntermediaireMV_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_MV_MES) * IntermediaireMES_kgj
                IntermediaireDebitVolumique_m3j = md_EntreeBouesPourcentageFluxAdmis(I) * TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_Q)
                IntermediaireDCO_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_ratio_DCO_MES) * IntermediaireMES_kgj
                IntermediaireDBO_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_ratio_DBO_MES) * IntermediaireMES_kgj
                IntermediaireNK_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_ratio_NK_MES) * IntermediaireMES_kgj
                IntermediaireP_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_ratio_Pt_MES) * IntermediaireMES_kgj
                
                If mi_ChoixPositionEntreeBoues(I) = mi_POSITION_ENTREE_BOUES_NON_DIGEREES Then
                    'On calcule la réduction potentielle des MV
                    If mi_TypeBouesEntree(I) = boues_I And ChoixEauxSalesPrimaire = True Then
                        PartBouesBiologiques = 0
                        CoefficientReductionMV = 0
                        For k = 1 To nb_eaux_sales_max
                            If DetailBouesI(k, repere_origine) > 0 Then
                                PartBouesBiologiques = PartBouesBiologiques + DetailBouesI(k, repere_ratio_MV)
                                CoefficientReductionMV = CoefficientReductionMV + DetailBouesI(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesLysees(DetailBouesI(k, repere_origine))
                            End If
                        Next k
                        CoefficientReductionMV = CoefficientReductionMV + (1 - PartBouesBiologiques) * md_ReductionMatieresVolatilesBouesLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine))
                        IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                    ElseIf mi_TypeBouesEntree(I) = boues_II And choix_biostyr = True Then
                        CoefficientReductionMV = 0
                        For k = 1 To nb_eaux_sales_max
                            If DetailBouesII(k, repere_origine) > 0 Then
                                CoefficientReductionMV = CoefficientReductionMV + DetailBouesII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesLysees(DetailBouesII(k, repere_origine))
                            End If
                        Next k
                        IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                    ElseIf mi_TypeBouesEntree(I) = boues_III And choix_biostyr_N_III = True Then
                        CoefficientReductionMV = 0
                        For k = 1 To nb_eaux_sales_max
                            If DetailBouesIII(k, repere_origine) > 0 Then
                                CoefficientReductionMV = CoefficientReductionMV + DetailBouesIII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesLysees(DetailBouesIII(k, repere_origine))
                            End If
                        Next k
                        IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                    ElseIf mi_TypeBouesEntree(I) = graisses And choix_coferments = True Then
                        'graisse interne
                        IntermediaireReductionMV_kgj = md_ReductionMatieresVolatilesBouesLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine)) * IntermediaireMV_kgj * PartGraissesInternesSurTotalCosubstrat(repere_graisse_particulaire_ratioMV)
                        'coferments
                        CofermentsCoefficientReductionMV = md_FacteurSecuritePotentielProductionMethaneCoferments * (md_PotentielProductionMethaneCoferments_Nm3_tMVAppliques / 1000) / (md_ProductionSpecifiqueBiogaz_Nm3_kgMV * md_RatioCH4DansBiogaz)
                        If CofermentsCoefficientReductionMV > 1 Then
                            CofermentsCoefficientReductionMV = 1
                        End If
                        IntermediaireReductionMVSoluble_kgj = CofermentsCoefficientReductionMV * TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireReductionMV_kgj = IntermediaireReductionMV_kgj + IntermediaireReductionMVSoluble_kgj + CofermentsCoefficientReductionMV * IntermediaireMV_kgj * (1 - PartGraissesInternesSurTotalCosubstrat(repere_graisse_particulaire_ratioMV))
                        'on rajoute la partie soluble pour les MS, les MV, la DCO, le N et le P
                        IntermediaireMS_kgj = IntermediaireMS_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_MS_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        'BouesEntreeNonDigereesMS_kgj = BouesEntreeNonDigereesMS_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), mi_TypeBouesEntree(i), repere_mgL_MS_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireMV_kgj = IntermediaireMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        'BouesEntreeNonDigereesMV_kgj = BouesEntreeNonDigereesMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), mi_TypeBouesEntree(i), repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireDCO_kgj = IntermediaireDCO_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_DCO) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireNK_kgj = IntermediaireNK_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_NK) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireP_kgj = IntermediaireP_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_Pt) * IntermediaireDebitVolumique_m3j / 1000
                    Else
                        IntermediaireReductionMV_kgj = md_ReductionMatieresVolatilesBouesLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine)) * IntermediaireMV_kgj
                    End If
                    'on calcul l'entrée globale de la lyse
                    BouesEntreeNonDigereesMS_kgj = BouesEntreeNonDigereesMS_kgj + IntermediaireMS_kgj
                    BouesEntreeNonDigereesMES_kgj = BouesEntreeNonDigereesMES_kgj + IntermediaireMES_kgj
                    BouesEntreeNonDigereesMV_kgj = BouesEntreeNonDigereesMV_kgj + IntermediaireMV_kgj
                    BouesEntreeNonDigereesReductionMV_kgj = BouesEntreeNonDigereesReductionMV_kgj + IntermediaireReductionMV_kgj
                    BouesEntreeNonDigereesDCO_kgj = BouesEntreeNonDigereesDCO_kgj + IntermediaireDCO_kgj
                    BouesEntreeNonDigereesDBO_kgj = BouesEntreeNonDigereesDBO_kgj + IntermediaireDBO_kgj
                    BouesEntreeNonDigereesNK_kgj = BouesEntreeNonDigereesNK_kgj + IntermediaireNK_kgj
                    BouesEntreeNonDigereesP_kgj = BouesEntreeNonDigereesP_kgj + IntermediaireP_kgj
                    BouesEntreeNonDigereesDebitVolumique_m3j = BouesEntreeNonDigereesDebitVolumique_m3j + IntermediaireDebitVolumique_m3j
                    'On remplit aussi le tableau avec le détail
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_MV_MES) = IntermediaireMV_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_MV_MES)
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_REDUCTION_MV) = IntermediaireReductionMV_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_REDUCTION_MV)
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_MS_kgj) + IntermediaireMS_kgj
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DCO_MES) = IntermediaireDCO_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DCO_MES)
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DBO_MES) = IntermediaireDBO_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DBO_MES)
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_NK_MES) = IntermediaireNK_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_NK_MES)
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_P_MES) = IntermediaireP_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_P_MES)
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_Q_m3j) = IntermediaireDebitVolumique_m3j + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_Q_m3j)
                ElseIf mi_ChoixPositionEntreeBoues(I) = mi_POSITION_ENTREE_BOUES_DIGESTEUR1 Then
                    'On calcule la réduction potentielle des MV
                    If mi_TypeBouesEntree(I) = boues_I And ChoixEauxSalesPrimaire = True Then
                        PartBouesBiologiques = 0
                        CoefficientReductionMV = 0
                        For k = 1 To nb_eaux_sales_max
                            If DetailBouesI(k, repere_origine) > 0 Then
                                PartBouesBiologiques = PartBouesBiologiques + DetailBouesI(k, repere_ratio_MV)
                                CoefficientReductionMV = CoefficientReductionMV + DetailBouesI(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesNonLysees(DetailBouesI(k, repere_origine))
                                md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) + DetailBouesI(k, repere_ratio_MV) * md_RatioMVSolublesMVTotauxBouesDigereesReference(DetailBouesI(k, repere_origine))
                                md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) + DetailBouesI(k, repere_ratio_MV) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(DetailBouesI(k, repere_origine))
                                md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) + DetailBouesI(k, repere_ratio_MV) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(DetailBouesI(k, repere_origine))
                            End If
                        Next k
                        CoefficientReductionMV = CoefficientReductionMV + (1 - PartBouesBiologiques) * md_ReductionMatieresVolatilesBouesNonLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine))
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) + (1 - PartBouesBiologiques) * md_RatioMVSolublesMVTotauxBouesDigereesReference(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine))
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) + (1 - PartBouesBiologiques) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine))
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) + (1 - PartBouesBiologiques) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine))
                        IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                    ElseIf mi_TypeBouesEntree(I) = boues_II And choix_biostyr = True Then
                        CoefficientReductionMV = 0
                        For k = 1 To nb_eaux_sales_max
                            If DetailBouesII(k, repere_origine) > 0 Then
                                CoefficientReductionMV = CoefficientReductionMV + DetailBouesII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesNonLysees(DetailBouesII(k, repere_origine))
                                md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) + DetailBouesII(k, repere_ratio_MV) * md_RatioMVSolublesMVTotauxBouesDigereesReference(DetailBouesII(k, repere_origine))
                                md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) + DetailBouesII(k, repere_ratio_MV) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(DetailBouesII(k, repere_origine))
                                md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) + DetailBouesII(k, repere_ratio_MV) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(DetailBouesII(k, repere_origine))
                            End If
                        Next k
                        IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                    ElseIf mi_TypeBouesEntree(I) = boues_III And choix_biostyr_N_III = True Then
                        CoefficientReductionMV = 0
                        For k = 1 To nb_eaux_sales_max
                            If DetailBouesIII(k, repere_origine) > 0 Then
                                CoefficientReductionMV = CoefficientReductionMV + DetailBouesIII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesNonLysees(DetailBouesIII(k, repere_origine))
                                md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) + DetailBouesIII(k, repere_ratio_MV) * md_RatioMVSolublesMVTotauxBouesDigereesReference(DetailBouesIII(k, repere_origine))
                                md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) + DetailBouesIII(k, repere_ratio_MV) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(DetailBouesIII(k, repere_origine))
                                md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) + DetailBouesIII(k, repere_ratio_MV) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(DetailBouesIII(k, repere_origine))
                            End If
                        Next k
                        IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                    ElseIf mi_TypeBouesEntree(I) = graisses And choix_coferments = True Then
                        'graisse interne
                        IntermediaireReductionMV_kgj = md_ReductionMatieresVolatilesBouesNonLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine)) * IntermediaireMV_kgj * PartGraissesInternesSurTotalCosubstrat(repere_graisse_particulaire_ratioMV)
                        'coferments
                        CofermentsCoefficientReductionMV = md_FacteurSecuritePotentielProductionMethaneCoferments * (md_PotentielProductionMethaneCoferments_Nm3_tMVAppliques / 1000) / (md_ProductionSpecifiqueBiogaz_Nm3_kgMV * md_RatioCH4DansBiogaz)
                        If CofermentsCoefficientReductionMV > 1 Then
                            CofermentsCoefficientReductionMV = 1
                        End If
                        IntermediaireReductionMVSoluble_kgj = CofermentsCoefficientReductionMV * TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireReductionMV_kgj = IntermediaireReductionMV_kgj + IntermediaireReductionMVSoluble_kgj + CofermentsCoefficientReductionMV * IntermediaireMV_kgj * (1 - PartGraissesInternesSurTotalCosubstrat(repere_graisse_particulaire_ratioMV))
                        'on rajoute la partie soluble pour les MS, les MV, la DCO, le N et le P
                        IntermediaireMS_kgj = IntermediaireMS_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_MS_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        'BouesEntreeDigesteur1MS_kgj = BouesEntreeDigesteur1MS_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), mi_TypeBouesEntree(i), repere_mgL_MS_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireMV_kgj = IntermediaireMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        'BouesEntreeDigesteur1MV_kgj = BouesEntreeDigesteur1MV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), mi_TypeBouesEntree(i), repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireDCO_kgj = IntermediaireDCO_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_DCO) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireNK_kgj = IntermediaireNK_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_NK) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireP_kgj = IntermediaireP_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_Pt) * IntermediaireDebitVolumique_m3j / 1000
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) + md_RatioMVSolublesMVTotauxBouesDigereesReference(codigestion_graisses)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) + md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(codigestion_graisses)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) + md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(codigestion_graisses)
                    Else
                        IntermediaireReductionMV_kgj = md_ReductionMatieresVolatilesBouesNonLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine)) * IntermediaireMV_kgj
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) + md_RatioMVSolublesMVTotauxBouesDigereesReference(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine))
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) + md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine))
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) + md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine))
                    End If
                    'on calcul l'entrée des boues non lysées qui vont vers le mélangeur
                    BouesEntreeDigesteur1MES_kgj = BouesEntreeDigesteur1MES_kgj + IntermediaireMES_kgj
                    BouesEntreeDigesteur1MS_kgj = BouesEntreeDigesteur1MS_kgj + IntermediaireMS_kgj
                    BouesEntreeDigesteur1MV_kgj = BouesEntreeDigesteur1MV_kgj + IntermediaireMV_kgj
                    BouesEntreeDigesteur1ReductionMV_kgj = BouesEntreeDigesteur1ReductionMV_kgj + IntermediaireReductionMV_kgj
                    BouesEntreeDigesteur1DCO_kgj = BouesEntreeDigesteur1DCO_kgj + IntermediaireDCO_kgj
                    BouesEntreeDigesteur1DBO_kgj = BouesEntreeDigesteur1DBO_kgj + IntermediaireDBO_kgj
                    BouesEntreeDigesteur1NK_kgj = BouesEntreeDigesteur1NK_kgj + IntermediaireNK_kgj
                    BouesEntreeDigesteur1P_kgj = BouesEntreeDigesteur1P_kgj + IntermediaireP_kgj
                    BouesEntreeDigesteur1DebitVolumique_m3j = BouesEntreeDigesteur1DebitVolumique_m3j + IntermediaireDebitVolumique_m3j
                    'On remplit aussi le tableau avec le détail
                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_MV_MES) = IntermediaireMV_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_MV_MES)
                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_REDUCTION_MV) = IntermediaireReductionMV_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_REDUCTION_MV)
                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_MS_kgj) + IntermediaireMS_kgj
                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DCO_MES) = IntermediaireDCO_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DCO_MES)
                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DBO_MES) = IntermediaireDBO_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DBO_MES)
                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_NK_MES) = IntermediaireNK_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_NK_MES)
                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_P_MES) = IntermediaireP_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_P_MES)
                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_Q_m3j) = IntermediaireDebitVolumique_m3j + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_Q_m3j)
                End If
                'pour la vérification des flux
                TableauRecapitulatifFluxBoues(boues_digerees, mi_TypeBouesEntree(I), repere_flux_in) = TableauRecapitulatifFluxBoues(boues_digerees, mi_TypeBouesEntree(I), repere_flux_in) + TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_flux_in) * md_EntreeBouesPourcentageFluxAdmis(I)
                TableauRecapitulatifFluxBoues(boues_digerees, mi_TypeBouesEntree(I), repere_verif_flux) = TableauRecapitulatifFluxBoues(boues_digerees, mi_TypeBouesEntree(I), repere_flux_in)
                'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_verif_flux) = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_verif_flux) - TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_flux_in) * md_EntreeBouesPourcentageFluxAdmis(I)
            End If
            
        Else
            If mi_ChoixPositionEntreeBoues(I) = mi_POSITION_ENTREE_BOUES_NON_DIGEREES Then
                LyseAvalEpaississement(mi_ProvenanceBouesEntree(I) - boues_inlet) = pi_CONFIGURATION_EXELYS
            Else
                LyseAvalEpaississement(mi_ProvenanceBouesEntree(I) - boues_inlet) = 0
            End If
            For j = 1 To nb_type_boues
                'variable d'entrée
                IntermediaireMES_kgj = md_EntreeBouesPourcentageFluxAdmis(I) * TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_MES)
                If IntermediaireMES_kgj <> 0 Then
                    IntermediaireMS_kgj = IntermediaireMES_kgj
                    IntermediaireMV_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_MV_MES) * IntermediaireMES_kgj
                    IntermediaireDebitVolumique_m3j = md_EntreeBouesPourcentageFluxAdmis(I) * TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_Q)
                    IntermediaireDCO_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_ratio_DCO_MES) * IntermediaireMES_kgj
                    IntermediaireDBO_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_ratio_DBO_MES) * IntermediaireMES_kgj
                    IntermediaireNK_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_ratio_NK_MES) * IntermediaireMES_kgj
                    IntermediaireP_kgj = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_ratio_Pt_MES) * IntermediaireMES_kgj
                    
                    If mi_ChoixPositionEntreeBoues(I) = mi_POSITION_ENTREE_BOUES_NON_DIGEREES Then
                        'On calcule la réduction potentielle des MV
                        If j = boues_I And ChoixEauxSalesPrimaire = True Then
                            PartBouesBiologiques = 0
                            CoefficientReductionMV = 0
                            For k = 1 To nb_eaux_sales_max
                                If DetailBouesI(k, repere_origine) > 0 Then
                                    PartBouesBiologiques = PartBouesBiologiques + DetailBouesI(k, repere_ratio_MV)
                                    CoefficientReductionMV = CoefficientReductionMV + DetailBouesI(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesLysees(DetailBouesI(k, repere_origine))
                                End If
                            Next k
                            CoefficientReductionMV = CoefficientReductionMV + (1 - PartBouesBiologiques) * md_ReductionMatieresVolatilesBouesLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine))
                            IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                        ElseIf j = boues_II And choix_biostyr = True Then
                            CoefficientReductionMV = 0
                            For k = 1 To nb_eaux_sales_max
                                If DetailBouesII(k, repere_origine) > 0 Then
                                    CoefficientReductionMV = CoefficientReductionMV + DetailBouesII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesLysees(DetailBouesII(k, repere_origine))
                                End If
                            Next k
                            IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                        ElseIf j = boues_III And choix_biostyr_N_III = True Then
                            CoefficientReductionMV = 0
                            For k = 1 To nb_eaux_sales_max
                                If DetailBouesIII(k, repere_origine) > 0 Then
                                    CoefficientReductionMV = CoefficientReductionMV + DetailBouesIII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesLysees(DetailBouesIII(k, repere_origine))
                                End If
                            Next k
                            IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                        ElseIf j = graisses And choix_coferments = True Then
                            'graisse interne
                            IntermediaireReductionMV_kgj = md_ReductionMatieresVolatilesBouesLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine)) * IntermediaireMV_kgj * PartGraissesInternesSurTotalCosubstrat(repere_graisse_particulaire_ratioMV)
                            'coferments
                            CofermentsCoefficientReductionMV = md_FacteurSecuritePotentielProductionMethaneCoferments * (md_PotentielProductionMethaneCoferments_Nm3_tMVAppliques / 1000) / (md_ProductionSpecifiqueBiogaz_Nm3_kgMV * md_RatioCH4DansBiogaz)
                            If CofermentsCoefficientReductionMV > 1 Then
                                CofermentsCoefficientReductionMV = 1
                            End If
                            IntermediaireReductionMVSoluble_kgj = CofermentsCoefficientReductionMV * TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireReductionMV_kgj = IntermediaireReductionMV_kgj + IntermediaireReductionMVSoluble_kgj + CofermentsCoefficientReductionMV * IntermediaireMV_kgj * (1 - PartGraissesInternesSurTotalCosubstrat(repere_graisse_particulaire_ratioMV))
                            'on rajoute la partie soluble pour les MS, les MV, la DCO, le N et le P
                            IntermediaireMS_kgj = IntermediaireMS_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_MS_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            'BouesEntreeNonDigereesMS_kgj = BouesEntreeNonDigereesMS_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), j, repere_mgL_MS_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireMV_kgj = IntermediaireMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            'BouesEntreeNonDigereesMV_kgj = BouesEntreeNonDigereesMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), j, repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireDCO_kgj = IntermediaireDCO_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_DCO) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireNK_kgj = IntermediaireNK_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_NK) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireP_kgj = IntermediaireP_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_Pt) * IntermediaireDebitVolumique_m3j / 1000
                        Else
                            IntermediaireReductionMV_kgj = md_ReductionMatieresVolatilesBouesLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine)) * IntermediaireMV_kgj
                        End If
                        'on calcul l'entrée globale de la lyse
                        BouesEntreeNonDigereesMS_kgj = BouesEntreeNonDigereesMS_kgj + IntermediaireMS_kgj
                        BouesEntreeNonDigereesMES_kgj = BouesEntreeNonDigereesMES_kgj + IntermediaireMES_kgj
                        BouesEntreeNonDigereesMV_kgj = BouesEntreeNonDigereesMV_kgj + IntermediaireMV_kgj
                        BouesEntreeNonDigereesReductionMV_kgj = BouesEntreeNonDigereesReductionMV_kgj + IntermediaireReductionMV_kgj
                        BouesEntreeNonDigereesDCO_kgj = BouesEntreeNonDigereesDCO_kgj + IntermediaireDCO_kgj
                        BouesEntreeNonDigereesDBO_kgj = BouesEntreeNonDigereesDBO_kgj + IntermediaireDBO_kgj
                        BouesEntreeNonDigereesNK_kgj = BouesEntreeNonDigereesNK_kgj + IntermediaireNK_kgj
                        BouesEntreeNonDigereesP_kgj = BouesEntreeNonDigereesP_kgj + IntermediaireP_kgj
                        BouesEntreeNonDigereesDebitVolumique_m3j = BouesEntreeNonDigereesDebitVolumique_m3j + IntermediaireDebitVolumique_m3j
                        'On remplit aussi le tableau avec le détail
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = IntermediaireMV_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = IntermediaireReductionMV_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_REDUCTION_MV)
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_MS_kgj) + IntermediaireMS_kgj
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = IntermediaireDCO_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = IntermediaireDBO_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = IntermediaireNK_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_NK_MES)
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = IntermediaireP_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_P_MES)
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_Q_m3j) = IntermediaireDebitVolumique_m3j + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_Q_m3j)
                    ElseIf mi_ChoixPositionEntreeBoues(I) = mi_POSITION_ENTREE_BOUES_DIGESTEUR1 Then
                        'On calcule la réduction potentielle des MV
                        If j = boues_I And ChoixEauxSalesPrimaire = True Then
                            PartBouesBiologiques = 0
                            CoefficientReductionMV = 0
                            For k = 1 To nb_eaux_sales_max
                                If DetailBouesI(k, repere_origine) > 0 Then
                                    PartBouesBiologiques = PartBouesBiologiques + DetailBouesI(k, repere_ratio_MV)
                                    CoefficientReductionMV = CoefficientReductionMV + DetailBouesI(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesNonLysees(DetailBouesI(k, repere_origine))
                                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) + DetailBouesI(k, repere_ratio_MV) * md_RatioMVSolublesMVTotauxBouesDigereesReference(DetailBouesI(k, repere_origine))
                                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) + DetailBouesI(k, repere_ratio_MV) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(DetailBouesI(k, repere_origine))
                                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) + DetailBouesI(k, repere_ratio_MV) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(DetailBouesI(k, repere_origine))
                                End If
                            Next k
                            CoefficientReductionMV = CoefficientReductionMV + (1 - PartBouesBiologiques) * md_ReductionMatieresVolatilesBouesNonLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine))
                            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) + (1 - PartBouesBiologiques) * md_RatioMVSolublesMVTotauxBouesDigereesReference(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine))
                            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) + (1 - PartBouesBiologiques) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine))
                            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) + (1 - PartBouesBiologiques) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine))
                            IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                        ElseIf j = boues_II And choix_biostyr = True Then
                            CoefficientReductionMV = 0
                            For k = 1 To nb_eaux_sales_max
                                If DetailBouesII(k, repere_origine) > 0 Then
                                    CoefficientReductionMV = CoefficientReductionMV + DetailBouesII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesNonLysees(DetailBouesII(k, repere_origine))
                                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) + DetailBouesII(k, repere_ratio_MV) * md_RatioMVSolublesMVTotauxBouesDigereesReference(DetailBouesII(k, repere_origine))
                                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) + DetailBouesII(k, repere_ratio_MV) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(DetailBouesII(k, repere_origine))
                                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) + DetailBouesII(k, repere_ratio_MV) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(DetailBouesII(k, repere_origine))
                                End If
                            Next k
                            IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                        ElseIf j = boues_III And choix_biostyr_N_III = True Then
                            CoefficientReductionMV = 0
                            For k = 1 To nb_eaux_sales_max
                                If DetailBouesIII(k, repere_origine) > 0 Then
                                    CoefficientReductionMV = CoefficientReductionMV + DetailBouesIII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesNonLysees(DetailBouesIII(k, repere_origine))
                                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) + DetailBouesIII(k, repere_ratio_MV) * md_RatioMVSolublesMVTotauxBouesDigereesReference(DetailBouesIII(k, repere_origine))
                                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) + DetailBouesIII(k, repere_ratio_MV) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(DetailBouesIII(k, repere_origine))
                                    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) + DetailBouesIII(k, repere_ratio_MV) * md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(DetailBouesIII(k, repere_origine))
                                End If
                            Next k
                            IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                        ElseIf j = graisses And choix_coferments = True Then
                            'graisse interne
                            IntermediaireReductionMV_kgj = md_ReductionMatieresVolatilesBouesNonLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine)) * IntermediaireMV_kgj * PartGraissesInternesSurTotalCosubstrat(repere_graisse_particulaire_ratioMV)
                            'coferments
                            CofermentsCoefficientReductionMV = md_FacteurSecuritePotentielProductionMethaneCoferments * (md_PotentielProductionMethaneCoferments_Nm3_tMVAppliques / 1000) / (md_ProductionSpecifiqueBiogaz_Nm3_kgMV * md_RatioCH4DansBiogaz)
                            If CofermentsCoefficientReductionMV > 1 Then
                                CofermentsCoefficientReductionMV = 1
                            End If
                            IntermediaireReductionMVSoluble_kgj = CofermentsCoefficientReductionMV * TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireReductionMV_kgj = IntermediaireReductionMV_kgj + IntermediaireReductionMVSoluble_kgj + CofermentsCoefficientReductionMV * IntermediaireMV_kgj * (1 - PartGraissesInternesSurTotalCosubstrat(repere_graisse_particulaire_ratioMV))
                            'on rajoute la partie soluble pour les MS, les MV, la DCO, le N et le P
                            IntermediaireMS_kgj = IntermediaireMS_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_MS_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            'BouesEntreeDigesteur1MS_kgj = BouesEntreeDigesteur1MS_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), j, repere_mgL_MS_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireMV_kgj = IntermediaireMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            'BouesEntreeDigesteur1MV_kgj = BouesEntreeDigesteur1MV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), j, repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireDCO_kgj = IntermediaireDCO_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_DCO) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireNK_kgj = IntermediaireNK_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_NK) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireP_kgj = IntermediaireP_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_Pt) * IntermediaireDebitVolumique_m3j / 1000
                            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) + md_RatioMVSolublesMVTotauxBouesDigereesReference(codigestion_graisses)
                            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) + md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(codigestion_graisses)
                            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) + md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(codigestion_graisses)
                        Else
                            IntermediaireReductionMV_kgj = md_ReductionMatieresVolatilesBouesNonLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine)) * IntermediaireMV_kgj
                            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) + md_RatioMVSolublesMVTotauxBouesDigereesReference(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine))
                            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) + md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine))
                            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) + md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine))
                        End If
                        'on calcul l'entrée des boues non lysées qui vont vers le mélangeur
                        BouesEntreeDigesteur1MS_kgj = BouesEntreeDigesteur1MS_kgj + IntermediaireMS_kgj
                        BouesEntreeDigesteur1MES_kgj = BouesEntreeDigesteur1MES_kgj + IntermediaireMES_kgj
                        BouesEntreeDigesteur1MV_kgj = BouesEntreeDigesteur1MV_kgj + IntermediaireMV_kgj
                        BouesEntreeDigesteur1ReductionMV_kgj = BouesEntreeDigesteur1ReductionMV_kgj + IntermediaireReductionMV_kgj
                        BouesEntreeDigesteur1DCO_kgj = BouesEntreeDigesteur1DCO_kgj + IntermediaireDCO_kgj
                        BouesEntreeDigesteur1DBO_kgj = BouesEntreeDigesteur1DBO_kgj + IntermediaireDBO_kgj
                        BouesEntreeDigesteur1NK_kgj = BouesEntreeDigesteur1NK_kgj + IntermediaireNK_kgj
                        BouesEntreeDigesteur1P_kgj = BouesEntreeDigesteur1P_kgj + IntermediaireP_kgj
                        BouesEntreeDigesteur1DebitVolumique_m3j = BouesEntreeDigesteur1DebitVolumique_m3j + IntermediaireDebitVolumique_m3j
                        'On remplit aussi le tableau avec le détail
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = IntermediaireMV_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = IntermediaireReductionMV_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_REDUCTION_MV)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_MS_kgj) + IntermediaireMS_kgj
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = IntermediaireDCO_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = IntermediaireDBO_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = IntermediaireNK_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_NK_MES)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = IntermediaireP_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_P_MES)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_Q_m3j) = IntermediaireDebitVolumique_m3j + md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_Q_m3j)
                    End If
                    'pour la vérification des flux
                    TableauRecapitulatifFluxBoues(boues_digerees, j, repere_flux_in) = TableauRecapitulatifFluxBoues(boues_digerees, j, repere_flux_in) + TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_flux_in) * md_EntreeBouesPourcentageFluxAdmis(I)
                    TableauRecapitulatifFluxBoues(boues_digerees, j, repere_verif_flux) = TableauRecapitulatifFluxBoues(boues_digerees, j, repere_flux_in)
                    'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                    TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_verif_flux) = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_verif_flux) - TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_flux_in) * md_EntreeBouesPourcentageFluxAdmis(I)
                End If
            Next j
        End If
    End If
Next I

If (BouesEntreeDigesteur1MES_kgj = 0 Or BouesEntreeDigesteur1DebitVolumique_m3j = 0) Then
    MsgBox Feuil5.Range("erreur_choix_exelys_DLD_inlet"), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
    pn_EffacementResultatsEnCours = True
    pn_ErreurNonGeree = True
    Exit Sub
End If


'On complète la ligne sur la somme des boues puis on recalcule tous les ratios
'pour les boues qui vont être directment lysées
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) = BouesEntreeNonDigereesMS_kgj
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) = BouesEntreeNonDigereesMV_kgj
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = BouesEntreeNonDigereesDCO_kgj
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = BouesEntreeNonDigereesDBO_kgj
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_NK_MES) = BouesEntreeNonDigereesNK_kgj
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_P_MES) = BouesEntreeNonDigereesP_kgj
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) = BouesEntreeNonDigereesDebitVolumique_m3j
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) = BouesEntreeNonDigereesReductionMV_kgj
    For j = 1 To nb_type_boues + 1
        If md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_MS_kgj) > 0 Then
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_REDUCTION_MV) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_P_MES) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, j, mi_PARAMETRE_BOUES_Q_m3j)
        End If
    Next j

'pour les boues qui vont être digérées avant la lyse
    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) = BouesEntreeDigesteur1MS_kgj
    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) = BouesEntreeDigesteur1MV_kgj
    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = BouesEntreeDigesteur1DCO_kgj
    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = BouesEntreeDigesteur1DBO_kgj
    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_NK_MES) = BouesEntreeDigesteur1NK_kgj
    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_P_MES) = BouesEntreeDigesteur1P_kgj
    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) = BouesEntreeDigesteur1DebitVolumique_m3j
    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) = BouesEntreeDigesteur1ReductionMV_kgj
    For j = 1 To nb_type_boues + 1
        If md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_MS_kgj) > 0 Then
            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_REDUCTION_MV) / md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) / md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) / md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) / md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) / md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_RATIO_P_MES) / md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, j, mi_PARAMETRE_BOUES_Q_m3j)
        End If
    Next j

'Valeur récupérée pour le calcul des KPI
pd_MatieresVolatilesBouesEntrees_kgMVj = BouesEntreeDigesteur1MV_kgj + BouesEntreeNonDigereesMV_kgj


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub
Private Sub calcul_bilan_masse_digestion(ByRef CaracteristiquesDigestion() As Double, ByVal TempsSejourBouesDigesteur_j As Double, ByVal ReferenceTempsSejourBoues_j As Double, ByVal RepereFluxBouesEntreeDigesteur As Integer, ByVal RepereFluxBouesSortieDigesteur As Integer, ByRef RatioTempsSejourSurReference As Double)

'************************************************************************************************************
'OBJECTIF:
'   Cette procédure effectue le bilan masse lors de l'étape de digestion des boues
'************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer
Dim FacteurCorrectionTempsSejourBouesDigesteur As Double
Dim IntermediaireMV_kgj As Double
Dim IntermediaireMS_kgj As Double
Dim IntermediaireNK_kgj As Double
Dim IntermediaireP_kgj As Double
Dim IntermediaireDCO_kgj As Double
Dim IntermediaireDBO_kgj As Double
Dim IntermediaireDebitVolumique_m3j As Double
Dim IntermediaireReductionMV_kgj As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "27.21"
I = 0
j = 0
FacteurCorrectionTempsSejourBouesDigesteur = 0
IntermediaireMV_kgj = 0
IntermediaireMS_kgj = 0
IntermediaireNK_kgj = 0
IntermediaireP_kgj = 0
IntermediaireDCO_kgj = 0
IntermediaireDBO_kgj = 0
IntermediaireDebitVolumique_m3j = 0
IntermediaireReductionMV_kgj = 0

'On calcule le facteur de correction par rapport au temps de séjour
RatioTempsSejourSurReference = TempsSejourBouesDigesteur_j / ReferenceTempsSejourBoues_j
FacteurCorrectionTempsSejourBouesDigesteur = (md_CoefficientACalculFacteurCorrectionTempsSejourBouesDigesteur * RatioTempsSejourSurReference ^ md_CoefficientBCalculFacteurCorrectionTempsSejourBouesDigesteur) / (RatioTempsSejourSurReference ^ md_CoefficientBCalculFacteurCorrectionTempsSejourBouesDigesteur + md_CoefficientACalculFacteurCorrectionTempsSejourBouesDigesteur - 1)

'On calcule le flux en sortie de digestion
For j = 1 To nb_type_boues
    IntermediaireMS_kgj = md_FluxIntermediaires(RepereFluxBouesEntreeDigesteur, j, mi_PARAMETRE_BOUES_MS_kgj)
    If IntermediaireMS_kgj <> 0 Then
        'On récupère les flux en entrée de digestion et calcule la réduction des MV
        IntermediaireMV_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(RepereFluxBouesEntreeDigesteur, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
        IntermediaireDCO_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(RepereFluxBouesEntreeDigesteur, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
        IntermediaireDBO_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(RepereFluxBouesEntreeDigesteur, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
        IntermediaireNK_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(RepereFluxBouesEntreeDigesteur, j, mi_PARAMETRE_BOUES_RATIO_NK_MES)
        IntermediaireP_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(RepereFluxBouesEntreeDigesteur, j, mi_PARAMETRE_BOUES_RATIO_P_MES)
        IntermediaireDebitVolumique_m3j = md_FluxIntermediaires(RepereFluxBouesEntreeDigesteur, j, mi_PARAMETRE_BOUES_Q_m3j)
        IntermediaireReductionMV_kgj = md_FluxIntermediaires(RepereFluxBouesEntreeDigesteur, j, mi_PARAMETRE_BOUES_REDUCTION_MV) * IntermediaireMV_kgj * FacteurCorrectionTempsSejourBouesDigesteur * F_facteur_correction_temperature_digesteur(CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C))
        'On calcule les flux en sortie de digestion (il sont potentiellement modifiés plus tard suivant les entrées de l'utilisateur)
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, j, mi_PARAMETRE_BOUES_MS_kgj) = IntermediaireMS_kgj - IntermediaireReductionMV_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, j, mi_PARAMETRE_BOUES_Q_m3j) = IntermediaireDebitVolumique_m3j - IntermediaireReductionMV_kgj / md_MasseVolumiqueBoues_kg_m3
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, j, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, j, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, j, mi_PARAMETRE_BOUES_Q_m3j)
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = IntermediaireMV_kgj - IntermediaireReductionMV_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = IntermediaireReductionMV_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = IntermediaireDCO_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = IntermediaireDBO_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = IntermediaireNK_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = IntermediaireP_kgj
        'On ajoute également pour le global
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) + IntermediaireMS_kgj - IntermediaireReductionMV_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) + IntermediaireDebitVolumique_m3j - IntermediaireReductionMV_kgj / md_MasseVolumiqueBoues_kg_m3
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) + IntermediaireMV_kgj - IntermediaireReductionMV_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) + IntermediaireReductionMV_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) + IntermediaireDCO_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DBO_MES) + IntermediaireDBO_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_NK_MES) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_NK_MES) + IntermediaireNK_kgj
        md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_P_MES) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_P_MES) + IntermediaireP_kgj
    End If
Next j

'On calcule la part de MV dégradés pour chaque type de boues
For j = 1 To nb_type_boues
    md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, j, mi_PARAMETRE_BOUES_REDUCTION_MV) / md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV)
Next j

'calcul de la réduction des MV
CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_REDUCTION_MV) = md_FluxIntermediaires(RepereFluxBouesSortieDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) / (md_FluxIntermediaires(RepereFluxBouesEntreeDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) * md_FluxIntermediaires(RepereFluxBouesEntreeDigesteur, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj))

'Estimation de la température jusqu'au mélangeur
For I = mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE To mi_FLUX_CENTRIFUGEUSE_MELANGEUR
    md_FluxIntermediaires(I, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) - md_PerteTemperatureBouesDigereesDigesteur1AvantLyse_°C
Next I


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Private Sub calcul_bilan_masse_centrifugeuse(ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double, ByRef FluxRetoursTotaux() As Double, ByRef FluxRetoursSoluble() As Double, ByRef FluxRetoursTotauxDigestion() As Double, ByRef FluxRetoursSolubleDigestion() As Double)

'************************************************************************************************************
'OBJECTIF:
'   Cette procédure effectue le bilan masse lors de l'étape de déshydratation des boues digérées
'************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer
Dim BouesEntreeMV_kgj As Double
Dim BouesEntreeMS_kgj As Double
Dim BouesEntreeNK_kgj As Double
Dim BouesEntreeP_kgj As Double
Dim BouesEntreeDCO_kgj As Double
Dim BouesEntreeDBO_kgj As Double
Dim BouesEntreeDebitVolumique_m3j As Double
Dim BouesSortieMV_kgj As Double
Dim BouesSortieMS_kgj As Double
Dim BouesSortieNK_kgj As Double
Dim BouesSortieP_kgj As Double
Dim BouesSortieDCO_kgj As Double
Dim BouesSortieDBO_kgj As Double
Dim BouesSortieDebitVolumique_m3j As Double
Dim IntermediaireMV_kgj As Double
Dim IntermediaireMS_kgj As Double
Dim IntermediaireNK_kgj As Double
Dim IntermediaireP_kgj As Double
Dim IntermediaireDCO_kgj As Double
Dim IntermediaireDBO_kgj As Double
Dim IntermediaireDebitVolumique_m3j As Double
Dim IntermediaireReductionMV_kgj As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "27.21"
I = 0
j = 0
BouesEntreeMV_kgj = 0
BouesEntreeMS_kgj = 0
BouesEntreeNK_kgj = 0
BouesEntreeP_kgj = 0
BouesEntreeDCO_kgj = 0
BouesEntreeDBO_kgj = 0
BouesEntreeDebitVolumique_m3j = 0
BouesSortieMV_kgj = 0
BouesSortieMS_kgj = 0
BouesSortieNK_kgj = 0
BouesSortieP_kgj = 0
BouesSortieDCO_kgj = 0
BouesSortieDBO_kgj = 0
BouesSortieDebitVolumique_m3j = 0
IntermediaireMV_kgj = 0
IntermediaireMS_kgj = 0
IntermediaireNK_kgj = 0
IntermediaireP_kgj = 0
IntermediaireDCO_kgj = 0
IntermediaireDBO_kgj = 0
IntermediaireDebitVolumique_m3j = 0
IntermediaireReductionMV_kgj = 0

'conso de polymère
pd_CentrifugeuseConsommationPolymerePur_kgj = md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_CONSOMMATION_POLYMERE_kgMA_tMS) / 1000 * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj)
md_DebitPolymere_m3j = pd_CentrifugeuseConsommationPolymerePur_kgj / md_ConcentrationPolymereCentrifugeuse_gL

'calcul de la pollution soluble
For I = 1 To nb_type_boues
    For j = 1 To nb_repere_mgL
        TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, j) = TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, j) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) / (md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) + md_DebitPolymere_m3j)
    Next j
Next I

'On calcule le flux en sortie de centrifugeuse
BouesEntreeDebitVolumique_m3j = md_DebitPolymere_m3j
For j = 1 To nb_type_boues
    IntermediaireMS_kgj = md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, j, mi_PARAMETRE_BOUES_MS_kgj)
    If IntermediaireMS_kgj <> 0 Then
        'On récupère les flux en entrée de centrifugeuse
        IntermediaireMV_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
        IntermediaireDCO_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
        IntermediaireDBO_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
        IntermediaireNK_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, j, mi_PARAMETRE_BOUES_RATIO_NK_MES)
        IntermediaireP_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, j, mi_PARAMETRE_BOUES_RATIO_P_MES)
        IntermediaireDebitVolumique_m3j = md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, j, mi_PARAMETRE_BOUES_Q_m3j)
        BouesEntreeMV_kgj = BouesEntreeMV_kgj + IntermediaireMV_kgj
        BouesEntreeMS_kgj = BouesEntreeMS_kgj + IntermediaireMS_kgj
        BouesEntreeNK_kgj = BouesEntreeNK_kgj + IntermediaireNK_kgj
        BouesEntreeP_kgj = BouesEntreeP_kgj + IntermediaireP_kgj
        BouesEntreeDCO_kgj = BouesEntreeDCO_kgj + IntermediaireDCO_kgj
        BouesEntreeDBO_kgj = BouesEntreeDBO_kgj + IntermediaireDBO_kgj
        BouesEntreeDebitVolumique_m3j = BouesEntreeDebitVolumique_m3j
        
        'On calcule les flux en sortie de digestion (il sont potentiellement modifiés plus tard suivant les entrées de l'utilisateur)
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_MS_kgj) = (IntermediaireMS_kgj - IntermediaireMV_kgj) * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES) + IntermediaireMV_kgj * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES) * md_RatioMVParticulairesMVTotauxBouesDigereesDigesteur1
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = IntermediaireMV_kgj * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES) * md_RatioMVParticulairesMVTotauxBouesDigereesDigesteur1
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = IntermediaireDCO_kgj * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES)
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = IntermediaireDBO_kgj * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES)
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = IntermediaireNK_kgj * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES)
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = IntermediaireP_kgj * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES)
        BouesSortieMV_kgj = BouesSortieMV_kgj + IntermediaireMV_kgj * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES) * md_RatioMVParticulairesMVTotauxBouesDigereesDigesteur1
        BouesSortieMS_kgj = BouesSortieMS_kgj + (IntermediaireMS_kgj - IntermediaireMV_kgj) * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES) + IntermediaireMV_kgj * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES) * md_RatioMVParticulairesMVTotauxBouesDigereesDigesteur1
        BouesSortieNK_kgj = BouesSortieNK_kgj + IntermediaireNK_kgj * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES)
        BouesSortieP_kgj = BouesSortieP_kgj + IntermediaireP_kgj * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES)
        BouesSortieDCO_kgj = BouesSortieDCO_kgj + IntermediaireDCO_kgj * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES)
        BouesSortieDBO_kgj = BouesSortieDBO_kgj + IntermediaireDBO_kgj * md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES)
        BouesSortieDebitVolumique_m3j = BouesSortieDebitVolumique_m3j + md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
        
        'On ajoute également pour le global
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) + md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_MS_kgj)
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) + md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_Q_m3j)
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) + md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) + md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DBO_MES) + md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_NK_MES) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_NK_MES) + md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_NK_MES)
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_P_MES) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_P_MES) + md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_P_MES)
    End If
Next j

'On recalcule les ratios en fonction des MES
For j = 1 To nb_type_boues + 1
    IntermediaireMS_kgj = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_MS_kgj)
    If IntermediaireMS_kgj <> 0 Then
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) / IntermediaireMS_kgj
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) / IntermediaireMS_kgj
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) / IntermediaireMS_kgj
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) / IntermediaireMS_kgj
        md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_RATIO_P_MES) / IntermediaireMS_kgj
    End If
Next j

'On propage le potentiel de dégradation des MV
For j = 1 To nb_type_boues
    md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, j, mi_PARAMETRE_BOUES_REDUCTION_MV)
Next j

'calcul des retours en tête d'usine
FluxRetoursTotaux(repere_ret_Q) = FluxRetoursTotaux(repere_ret_Q) + BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j
FluxRetoursTotaux(repere_ret_MES) = FluxRetoursTotaux(repere_ret_MES) + BouesEntreeMS_kgj - BouesSortieMS_kgj
FluxRetoursTotaux(repere_ret_DCO) = FluxRetoursTotaux(repere_ret_DCO) + BouesEntreeDCO_kgj - BouesSortieDCO_kgj
FluxRetoursTotaux(repere_ret_DBO) = FluxRetoursTotaux(repere_ret_DBO) + BouesEntreeDBO_kgj - BouesSortieDBO_kgj
FluxRetoursTotaux(repere_ret_NK) = FluxRetoursTotaux(repere_ret_NK) + BouesEntreeNK_kgj - BouesSortieNK_kgj
FluxRetoursTotaux(repere_ret_Pt) = FluxRetoursTotaux(repere_ret_Pt) + BouesEntreeP_kgj - BouesSortieP_kgj
'vecteur spécifique aux retours de digestion (traitement séparé)
FluxRetoursTotauxDigestion(repere_ret_Q) = FluxRetoursTotauxDigestion(repere_ret_Q) + BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j
FluxRetoursTotauxDigestion(repere_ret_MES) = FluxRetoursTotauxDigestion(repere_ret_MES) + BouesEntreeMS_kgj - BouesSortieMS_kgj
FluxRetoursTotauxDigestion(repere_ret_DCO) = FluxRetoursTotauxDigestion(repere_ret_DCO) + BouesEntreeDCO_kgj - BouesSortieDCO_kgj
FluxRetoursTotauxDigestion(repere_ret_DBO) = FluxRetoursTotauxDigestion(repere_ret_DBO) + BouesEntreeDBO_kgj - BouesSortieDBO_kgj
FluxRetoursTotauxDigestion(repere_ret_NK) = FluxRetoursTotauxDigestion(repere_ret_NK) + BouesEntreeNK_kgj - BouesSortieNK_kgj
FluxRetoursTotauxDigestion(repere_ret_Pt) = FluxRetoursTotauxDigestion(repere_ret_Pt) + BouesEntreeP_kgj - BouesSortieP_kgj
For I = 1 To nb_type_boues
    FluxRetoursTotaux(repere_ret_DCO) = FluxRetoursTotaux(repere_ret_DCO) + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_DCO) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursTotaux(repere_ret_DBO) = FluxRetoursTotaux(repere_ret_DBO) + md_RatioDBOSolubleSurDCOSolubleBouesDigerees * TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_DCO) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursTotaux(repere_ret_NK) = FluxRetoursTotaux(repere_ret_NK) + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_NK) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursTotaux(repere_ret_NH4) = FluxRetoursTotaux(repere_ret_NH4) + md_RatioNH4SurNKSolubleBouesDigerees * TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_NK) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursTotaux(repere_ret_Pt) = FluxRetoursTotaux(repere_ret_Pt) + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_Pt) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursSoluble(repere_ret_DCO) = FluxRetoursSoluble(repere_ret_DCO) + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_DCO) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursSoluble(repere_ret_DBO) = FluxRetoursSoluble(repere_ret_DBO) + md_RatioDBOSolubleSurDCOSolubleBouesDigerees * TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_DCO) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursSoluble(repere_ret_NK) = FluxRetoursSoluble(repere_ret_NK) + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_NK) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursSoluble(repere_ret_NH4) = FluxRetoursSoluble(repere_ret_NH4) + md_RatioNH4SurNKSolubleBouesDigerees * TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_NK) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursSoluble(repere_ret_Pt) = FluxRetoursSoluble(repere_ret_Pt) + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_Pt) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    'vecteur spécifique aux retours de digestion (traitement séparé)
    FluxRetoursTotauxDigestion(repere_ret_DCO) = FluxRetoursTotauxDigestion(repere_ret_DCO) + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_DCO) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursTotauxDigestion(repere_ret_DBO) = FluxRetoursTotauxDigestion(repere_ret_DBO) + md_RatioDBOSolubleSurDCOSolubleBouesDigerees * TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_DCO) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursTotauxDigestion(repere_ret_NK) = FluxRetoursTotauxDigestion(repere_ret_NK) + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_NK) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursTotauxDigestion(repere_ret_NH4) = FluxRetoursTotauxDigestion(repere_ret_NH4) + md_RatioNH4SurNKSolubleBouesDigerees * TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_NK) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursTotauxDigestion(repere_ret_Pt) = FluxRetoursTotauxDigestion(repere_ret_Pt) + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_Pt) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursSolubleDigestion(repere_ret_DCO) = FluxRetoursSolubleDigestion(repere_ret_DCO) + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_DCO) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursSolubleDigestion(repere_ret_DBO) = FluxRetoursSolubleDigestion(repere_ret_DBO) + md_RatioDBOSolubleSurDCOSolubleBouesDigerees * TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_DCO) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursSolubleDigestion(repere_ret_NK) = FluxRetoursSolubleDigestion(repere_ret_NK) + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_NK) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursSolubleDigestion(repere_ret_NH4) = FluxRetoursSolubleDigestion(repere_ret_NH4) + md_RatioNH4SurNKSolubleBouesDigerees * TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_NK) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
    FluxRetoursSolubleDigestion(repere_ret_Pt) = FluxRetoursSolubleDigestion(repere_ret_Pt) + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_Pt) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * (BouesEntreeDebitVolumique_m3j - BouesSortieDebitVolumique_m3j) / 1000
Next I

'calcul des MV/MES (fuite des MV solubilisés et non méthanisés dans les retours)
For I = 1 To nb_type_boues
    md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) + md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) * (1 - md_RatioMVParticulairesMVTotauxBouesDigereesDigesteur1) * BouesSortieDebitVolumique_m3j / BouesEntreeDebitVolumique_m3j
Next I
md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) + md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) * (1 - md_RatioMVParticulairesMVTotauxBouesDigereesDigesteur1) * BouesSortieDebitVolumique_m3j / BouesEntreeDebitVolumique_m3j



Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Public Sub fonctionnement_reel(ByRef TableauRecapitulatifFluxBoues() As Double, ByRef PartGraissesInternesSurTotalCosubstrat() As Double, ByVal ChoixEauxSalesPrimaire As Boolean, ByRef DetailBouesI() As Double, ByRef DetailBouesII() As Double, ByRef DetailBouesIII() As Double, ByRef DCO_biogaz, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double, ByRef FluxRetoursTotaux() As Double, ByRef FluxRetoursTotauxDigestion() As Double, ByRef FluxRetoursSoluble() As Double, ByRef FluxRetoursSolubleDigestion() As Double, ByRef TableauBesoinsThermiques() As Double, ByRef TableauDisponibilitesThermiques() As Double, ByRef CaracteristiquesBiogaz() As Double, ByRef PressionVapeur_bar() As Double, ByRef NombreReseauxVapeur As Integer, ByRef LyseAvalEpaississement() As Integer)


'*********************************************************************************************************************************************************
'OBJECTIF:

'VARIABLES D'ENTREE:
'   TableauRecapitulatifFluxBoues(): contient les caractéristiques des boues alimentant le procédé
'   TableauRecapitulatifPollutionSolubleDansFluxBoues(): contient la polution soluble des boues alimentant le procédé

'VARIABLES DE SORTIE:
'   TableauRecapitulatifFluxBoues(): contient les caractéristiques des boues en sortie du procédé
'   TableauRecapitulatifPollutionSolubleDansFluxBoues(): contient la polution soluble des boues en sortie du procédé
'A COMPLETER
'*********************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer
Dim k As Integer
Dim ReperageBoucleReel As Boolean
Dim RatioMVSolubleSurMVTotauxBouesDigerees As Double
Dim CorrectionSolubilisationMVDigestionTempsSejour As Double
Dim BiogazProduitDigesteur1_Nm3j As Double
Dim FluxDCODansBiogazDigesteur1_kgj As Double
Dim FluxDCODansBiogazDigesteur2_kgj As Double
Dim BiogazProduitDigesteur2_Nm3j As Double
Dim RatioTempsSejourSurReference As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "27.6"
I = 0
j = 0
k = 0
ReperageBoucleReel = True
RatioMVSolubleSurMVTotauxBouesDigerees = 0
CorrectionSolubilisationMVDigestionTempsSejour = 0
BiogazProduitDigesteur1_Nm3j = 0
FluxDCODansBiogazDigesteur1_kgj = 0
FluxDCODansBiogazDigesteur2_kgj = 0
BiogazProduitDigesteur2_Nm3j = 0
RatioTempsSejourSurReference = 0


'RE-INITIALISATION DES FLUX
For I = 1 To mi_NOMBRE_FLUX_INTERMEDIAIRES
    For j = 1 To nb_type_boues + 1
        For k = 1 To mi_NOMBRE_PARAMETRES_BOUES
            If k <> mi_PARAMETRE_BOUES_TEMPERATURE_°C And k <> mi_PARAMETRE_BOUES_CONCENTRATION_gL Then
                md_FluxIntermediaires(I, j, k) = 0
            End If
        Next k
    Next j
Next I

'température des boues en entrée de digesteur 1
If mn_TemperatureBouesEntree_°C_force = True Then
    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureBouesEntree_°C_f
Else
    md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = pd_TemperatureEauConditionsReelles_°C
End If
'température des boues non digérées
If mn_TemperatureBouesNonDigerees_°C_force = True Then
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureBouesNonDigerees_°C_f
Else
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = pd_TemperatureEauConditionsReelles_°C
End If
'température de l'eau de dilution
If mn_TemperatureEauDilution_°C_force = True Then
    md_TemperatureEauDilution_°C = md_TemperatureEauDilution_°C_f
Else
    md_TemperatureEauDilution_°C = pd_TemperatureEauConditionsReelles_°C
End If

'CALCUL DES BOUES EN ENTREE (VERS LYSE ET BYPASS)
Call calcul_cocktail_boues_entrees(TableauRecapitulatifFluxBoues, TableauRecapitulatifPollutionSolubleDansFluxBoues, ChoixEauxSalesPrimaire, DetailBouesI, DetailBouesII, DetailBouesIII, PartGraissesInternesSurTotalCosubstrat, LyseAvalEpaississement)

'DIGESTEUR 1
md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_REEL_j) = md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_VOLUME_m3) / md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
Call calcul_bilan_masse_digestion(md_CaracteristiquesDigesteur1, md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_REEL_j), md_ReferenceTempsSejourBouesDigesteur1_j, mi_FLUX_ENTREE_DIGESTEUR1, mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, RatioTempsSejourSurReference)

'calcul de la part de MV soluble en sortie
Call calcul_MV_soluble_boues_digerees(RatioTempsSejourSurReference)

'possible modification de la réduction des MV
If mn_CaracteristiquesDigesteur1_force(mi_PARAMETRE_DIGESTION_REDUCTION_MV) = True Then
    Call forcage_rendement_MV(md_CaracteristiquesDigesteur1, md_CaracteristiquesDigesteur1_f, mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE)
End If

'calcul du biogaz produit dans digesteur 1
Call calcul_biogaz(BiogazProduitDigesteur1_Nm3j, FluxDCODansBiogazDigesteur1_kgj, mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE)

'calcul de la pollution solubilisée lors de la digestion
Call calcul_solubilisation_pendant_digestion(TableauRecapitulatifPollutionSolubleDansFluxBoues, md_CaracteristiquesDigesteur1, mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, md_RatioDCOSolubleSurDCOTotaleBouesDigereesDigesteur1)

'On attribue les nouveaux potentiels de dégradation des MV
For I = 1 To nb_type_boues
    md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_ReductionMatieresVolatilesBouesDigereesLysees
Next I

'CENTRIFUGEUSE
Call calcul_bilan_masse_centrifugeuse(TableauRecapitulatifPollutionSolubleDansFluxBoues, FluxRetoursTotaux, FluxRetoursSoluble, FluxRetoursTotauxDigestion, FluxRetoursSolubleDigestion)

'ENTREE DES BOUES NON DIGEREES
Call calcul_bilan_masse_melangeur(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_FLUX_BOUES_NON_DIGEREES, mi_FLUX_MELANGEUR_LYSE)

'HYDROLYSE THERMIQUE
Call calcul_bilan_masse_lyse(ReperageBoucleReel, TableauRecapitulatifPollutionSolubleDansFluxBoues)

'ECHANGEURS --> SEULEMENT LORS DES BILANS THERMIQUES ET ELECTRIQUES

'DILUTION AVEC DE L'EAU
Call calcul_dilution_boues_avant_digestion(ReperageBoucleReel, TableauRecapitulatifPollutionSolubleDansFluxBoues)

'DIGESTEUR 2
md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_REEL_j) = md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_VOLUME_m3) / md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
Call calcul_bilan_masse_digestion(md_CaracteristiquesDigesteur2, md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_REEL_j), md_ReferenceTempsSejourBouesDigesteur2_j, mi_FLUX_DILUTION_DIGESTEUR2, mi_FLUX_SORTIE_DIGESTEUR2, RatioTempsSejourSurReference)

'possible modification de la réduction des MV
If mn_CaracteristiquesDigesteur2_force(mi_PARAMETRE_DIGESTION_REDUCTION_MV) = True Then
    Call forcage_rendement_MV(md_CaracteristiquesDigesteur2, md_CaracteristiquesDigesteur2_f, mi_FLUX_SORTIE_DIGESTEUR2)
End If

'biogaz produit dans digesteur 2
Call calcul_biogaz(BiogazProduitDigesteur2_Nm3j, FluxDCODansBiogazDigesteur2_kgj, mi_FLUX_SORTIE_DIGESTEUR2)
pd_BiogazProduit_Nm3j = BiogazProduitDigesteur1_Nm3j + BiogazProduitDigesteur2_Nm3j

'calcul de la pollution solubilisée lors de la digestion
Call calcul_solubilisation_pendant_digestion(TableauRecapitulatifPollutionSolubleDansFluxBoues, md_CaracteristiquesDigesteur2, mi_FLUX_SORTIE_DIGESTEUR2, md_RatioDCOSolubleSurDCOTotaleBouesDigereesDigesteur2)

'calcul de la réduction globale (ici sont aussi inclus les MV retournés en tête)
md_TauxReductionGlobalMV = (md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) * md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) - md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) * md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) + md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) - md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj))
md_TauxReductionGlobalMV = md_TauxReductionGlobalMV / (md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) * md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) + md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) * md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES))

'GESTION ENERGIE
Call calcul_bilan_thermique(TableauBesoinsThermiques, TableauDisponibilitesThermiques, PressionVapeur_bar, NombreReseauxVapeur, ReperageBoucleReel)
CaracteristiquesBiogaz(Q_Nm3j) = pd_BiogazProduit_Nm3j
CaracteristiquesBiogaz(teneur_CH4) = md_RatioCH4DansBiogaz


'RECAPITULATIF SORTIE

For I = 1 To nb_type_boues
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_MS_kgj)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_Q) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_Q_m3j)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_MV_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_MV_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_DCO_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_DBO_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_NK_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_NK_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_Pt_MES) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_P_MES)
Next I



Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Private Sub calcul_bilan_masse_melangeur(ByVal RepereFluxBouesAMelanger1 As Integer, ByVal RepereFluxBouesAMelanger2 As Integer, ByVal RepereFluxBouesMelangeResultant As Integer)

'************************************************************************************************************
'OBJECTIF:
'   Cette procédure effectue le bilan masse lors du mélange des boues lysées avec les boues fraîches
'   (si nécessaire)
'************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "27.23"
I = 0

'Mélange des deux flux
For I = 1 To nb_type_boues + 1
    md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_MS_kgj) + md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_MS_kgj)
    If md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_MS_kgj) > 0 Then
        md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_Q_m3j) + md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_Q_m3j)
        md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_Q_m3j)
        md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) = (md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_MS_kgj) + md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_MS_kgj)) / md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_MS_kgj)
        md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = (md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_MS_kgj) + md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_MS_kgj)) / md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_MS_kgj)
        md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = (md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_RATIO_DBO_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_MS_kgj) + md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_RATIO_DBO_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_MS_kgj)) / md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_MS_kgj)
        md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_RATIO_NK_MES) = (md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_RATIO_NK_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_MS_kgj) + md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_RATIO_NK_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_MS_kgj)) / md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_MS_kgj)
        md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_RATIO_P_MES) = (md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_RATIO_P_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_MS_kgj) + md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_RATIO_P_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_MS_kgj)) / md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_MS_kgj)
        md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = (md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_Q_m3j) * md_CapaciteCalorifiqueBoues_kWh_m3°C * md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_TEMPERATURE_°C) + md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_Q_m3j) * md_CapaciteCalorifiqueBoues_kWh_m3°C * md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_TEMPERATURE_°C)) / (md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_Q_m3j) * md_CapaciteCalorifiqueBoues_kWh_m3°C)
        md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_REDUCTION_MV) = (md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_REDUCTION_MV) * md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_MS_kgj) + md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_REDUCTION_MV) * md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_MS_kgj))
            md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(RepereFluxBouesMelangeResultant, I, mi_PARAMETRE_BOUES_REDUCTION_MV) / (md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger1, I, mi_PARAMETRE_BOUES_MS_kgj) + md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) * md_FluxIntermediaires(RepereFluxBouesAMelanger2, I, mi_PARAMETRE_BOUES_MS_kgj))
    End If
Next I



Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Private Sub calcul_bilan_masse_lyse(ByVal ReperageBoucleReel As Boolean, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double)

'************************************************************************************************************
'OBJECTIF:
'   Cette procédure calcule la quantité de vapeur utilisée pour la lyse thermique
'************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "27.22"
I = 0
j = 0


'Toutes les valeurs concernant la pollution ne changent pas jusqu'au mélangeur
For I = mi_FLUX_LYSE_ECHANGEUR90 To mi_FLUX_ECHANGEUR55_DILUTION
    For j = 1 To nb_type_boues + 1
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_LYSE, j, mi_PARAMETRE_BOUES_MS_kgj)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_LYSE, j, mi_PARAMETRE_BOUES_REDUCTION_MV)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_LYSE, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_LYSE, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_LYSE, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_LYSE, j, mi_PARAMETRE_BOUES_RATIO_NK_MES)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_LYSE, j, mi_PARAMETRE_BOUES_RATIO_P_MES)
    Next j
Next I

'calcul du débit de vapeur en kg/j, du débit volumique et de la concentration des boues en sortie de la lyse
If mn_DebitVapeurLyse_kgj_force = True And ReperageBoucleReel = True Then
    md_DebitVapeurLyse_kgj = md_DebitVapeurLyse_kgj_f
Else
    md_DebitVapeurLyse_kgj = md_FluxIntermediaires(mi_FLUX_MELANGEUR_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * (md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR90, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) - md_FluxIntermediaires(mi_FLUX_MELANGEUR_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)) / ((1 - md_RatioPertesVapeur) * md_EnthalpieTotaleVapeurSaturee_kcal_kg - md_CapaciteCalorifiqueBoues_kcal_kg°C * md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR90, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C))
End If
For I = mi_FLUX_LYSE_ECHANGEUR90 To mi_FLUX_ECHANGEUR55_DILUTION
    md_FluxIntermediaires(I, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) + md_DebitVapeurLyse_kgj / md_MasseVolumiqueBoues_kg_m3
    md_FluxIntermediaires(I, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(I, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(I, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
    For j = 1 To nb_type_boues
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(I, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
    Next j
Next I
For I = 1 To nb_type_boues
    For j = 1 To repere_mgL_Pt
        If md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR90, I, mi_PARAMETRE_BOUES_Q_m3j) > 0 Then
            TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, j) = TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, j) * md_FluxIntermediaires(mi_FLUX_MELANGEUR_LYSE, I, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR90, I, mi_PARAMETRE_BOUES_Q_m3j)
        End If
    Next j
Next I


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Private Sub calcul_dilution_boues_avant_digestion(ByVal ReperageBoucleReel As Boolean, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double)

'************************************************************************************************************
'OBJECTIF:
'   Cette procédure effectue le bilan masse lors de la dilution en amont du digesteur
'   Elle calcule:
'       - la concentration atteinte en sortie si le débit d'eau de dilution a été renseigné par l'utilisateur
'       - le débit d'eau de dilution en fonction d'une concentration définie par défaut sinon
'************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "27.27"
I = 0
j = 0

'Calcul du débit d'eau de dilution
If mn_ConcentrationBouesDilutionDigesteur2_gL_force = True And ReperageBoucleReel = True Then
    If md_ConcentrationBouesDilutionDigesteur2_gL_f > md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL) Then
        md_DebitEauDilution_m3j = 0
    Else
        md_DebitEauDilution_m3j = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_ConcentrationBouesDilutionDigesteur2_gL_f - md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
    End If
Else
    If md_ConcentrationBouesDilutionDigesteur2Souhaitee_gL > md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL) Then
        md_DebitEauDilution_m3j = 0
    Else
        md_DebitEauDilution_m3j = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_ConcentrationBouesDilutionDigesteur2Souhaitee_gL - md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
    End If
End If

'Bilan matière
md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj)
md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) + md_DebitEauDilution_m3j
md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)

For I = 1 To nb_type_boues + 1
    If I <= nb_type_boues Then
        md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, I, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, I, mi_PARAMETRE_BOUES_MS_kgj)
        md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, I, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
        md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, I, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, I, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, I, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
        For j = 1 To repere_mgL_Pt
            TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, j) = TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, j) * md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR90, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) / md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
        Next j
    End If
    md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, I, mi_PARAMETRE_BOUES_RATIO_MV_MES)
    'md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, i, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = 0
    md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
    md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, I, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
    md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_NK_MES) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, I, mi_PARAMETRE_BOUES_RATIO_NK_MES)
    md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, I, mi_PARAMETRE_BOUES_RATIO_P_MES) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, I, mi_PARAMETRE_BOUES_RATIO_P_MES)
    md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, I, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, I, mi_PARAMETRE_BOUES_REDUCTION_MV)
Next I


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Private Sub calcul_biogaz(ByRef BiogazProduit_Nm3j As Double, ByRef FluxDCODansBiogaz_kgj As Double, ByVal RepereFluxBouesDigerees As Integer)

'************************************************************************************************************
'OBJECTIF:
'   Cette procédure calcule la production de biogaz
'************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim DensiteBiogaz_kg_Nm3 As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "27.25"
I = 0
DensiteBiogaz_kg_Nm3 = 0

'calcul du biogaz produit dans digesteur
DensiteBiogaz_kg_Nm3 = 0
For I = 0 To 1
    DensiteBiogaz_kg_Nm3 = DensiteBiogaz_kg_Nm3 + md_CoefficientsCalculDensiteBiogaz(I) * md_RatioCH4DansBiogaz ^ I
Next I
md_ProductionSpecifiqueBiogaz_Nm3_kgMV = 1 / DensiteBiogaz_kg_Nm3
BiogazProduit_Nm3j = md_ProductionSpecifiqueBiogaz_Nm3_kgMV * md_FluxIntermediaires(RepereFluxBouesDigerees, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV)

'calcul de la DCO du biogaz et ajustement de la DCO des boues digérées
FluxDCODansBiogaz_kgj = (md_RatioCH4DansBiogaz * BiogazProduit_Nm3j) / md_ProductionSpecifiqueMethane_Nm3_kgDCO
md_FluxIntermediaires(RepereFluxBouesDigerees, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(RepereFluxBouesDigerees, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) - FluxDCODansBiogaz_kgj

'on reclacule pour chaque type de boues la DCO des boues digérées
For I = 1 To nb_type_boues
    md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) - md_FluxIntermediaires(RepereFluxBouesDigerees, I, mi_PARAMETRE_BOUES_REDUCTION_MV) * FluxDCODansBiogaz_kgj
Next I



Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub
Private Sub calcul_MV_soluble_boues_digerees(ByVal RatioTempsSejourSurReference As Double)

'************************************************************************************************************
'OBJECTIF:
'   Cette procédure calcule la production de biogaz
'************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim RatioMVSolubleSurMVTotauxBouesDigerees As Double
Dim CorrectionSolubilisationMVDigestionTempsSejour As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "27.25"
I = 0
RatioMVSolubleSurMVTotauxBouesDigerees = 0
CorrectionSolubilisationMVDigestionTempsSejour = 0

'calcul de la part de MV soluble des boues digérées
RatioMVSolubleSurMVTotauxBouesDigerees = 0
For I = 1 To nb_type_boues
    If RatioTempsSejourSurReference < 1 Then
        CorrectionSolubilisationMVDigestionTempsSejour = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, I, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_COURT) * (RatioTempsSejourSurReference - 1) + 1
    Else
        CorrectionSolubilisationMVDigestionTempsSejour = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, I, mi_PARAMETRE_BOUES_CORRECTION_MV_SOLUBLE_TEMPS_SEJOUR_LONG) * (RatioTempsSejourSurReference - 1) + 1
    End If
    RatioMVSolubleSurMVTotauxBouesDigerees = RatioMVSolubleSurMVTotauxBouesDigerees + md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, I, mi_PARAMETRE_BOUES_REDUCTION_MV) * md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, I, mi_PARAMETRE_BOUES_PART_MV_SOLUBLE_APRES_DIGESTION) * CorrectionSolubilisationMVDigestionTempsSejour
Next I
md_RatioMVParticulairesMVTotauxBouesDigereesDigesteur1 = 1 - RatioMVSolubleSurMVTotauxBouesDigerees


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Public Sub calcul_consommation_electrique(ByVal EquipementValorisationBiogaz As Integer, ByVal ValorisationBiogazChaudiere As Integer, ByVal PartBiogazValorise As Double)


'******************************************************************************************************************************************************
'OBJECTIF:

'VARIABLES D'ENTREE:

'VARIABLES DE SORTIE:
'A COMPLETER
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim NombrePompes As Double
Dim PressionRefoulement_mCE As Double
Dim TempsFonctionnementChaquePompe_hj As Double
Dim DebitUnitairePompe_m3h As Double
Dim RendementGlobalPompe As Double
Dim DebitTotal_m3j As Double
Dim HMTAgitationBiogaz_mCE As Double
Dim PertesDeCharge_mCE As Double
Dim VitesseBrassageBiogaz_Nmh As Double
Dim DebitBrassageBiogaz_Nm3j As Double
Dim ConsommationElectriqueSpecifiqueSurpresseurBiogaz_Wh_Nm3mCE As Double
Dim ConsommationElectriqueAgitation_kWhj As Double
Dim DebitRecirculationAgitation_m3h As Double
Dim PuissancePompe_kW As Double
Dim RendementPompe As Double
Dim ConsommationElectriquePompages_kWhj As Double
Dim TempsSejourGazometre_h As Double
Dim VolumeGazometre_m3 As Double
Dim DebitSoufflageBiogaz_Nm3h As Double
Dim ConsommationElectriqueGazometre_kWhj As Double
Dim ConsommationElectriqueSurpresseurBiogaz_kWhj As Double
Dim IncoherenceCalcul As Double
Dim ConsommationElectriqueRecirculationEchangeur_kWhj As Double
Dim RatioConsommationElectriqueSpecifiqueDLD_kWh_tMS As Double
Dim ConsommationElectriqueSpecifiqueDLD_kWhj As Double
Dim RatioConsommationElectriqueSpecifiqueCentrifugeuse_kWh_m3 As Double
Dim ConsommationElectriqueCentrifugeuse_kWhj As Double
Dim RatioConsommationElectriqueSpecifiquePolymere_kWh_m3 As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "27.7"
NombrePompes = 0
PressionRefoulement_mCE = 0
TempsFonctionnementChaquePompe_hj = 0
DebitUnitairePompe_m3h = 0
RendementGlobalPompe = 0
DebitTotal_m3j = 0
HMTAgitationBiogaz_mCE = 0
PertesDeCharge_mCE = 0
VitesseBrassageBiogaz_Nmh = 0
DebitBrassageBiogaz_Nm3j = 0
ConsommationElectriqueSpecifiqueSurpresseurBiogaz_Wh_Nm3mCE = 0
ConsommationElectriqueAgitation_kWhj = 0
DebitRecirculationAgitation_m3h = 0
PuissancePompe_kW = 0
RendementPompe = 0
ConsommationElectriquePompages_kWhj = 0
TempsSejourGazometre_h = 0
VolumeGazometre_m3 = 0
DebitSoufflageBiogaz_Nm3h = 0
ConsommationElectriqueGazometre_kWhj = 0
ConsommationElectriqueSurpresseurBiogaz_kWhj = 0
IncoherenceCalcul = 0
ConsommationElectriqueRecirculationEchangeur_kWhj = 0
RatioConsommationElectriqueSpecifiqueDLD_kWh_tMS = 0
ConsommationElectriqueSpecifiqueDLD_kWhj = 0
RatioConsommationElectriqueSpecifiqueCentrifugeuse_kWh_m3 = 0
ConsommationElectriqueCentrifugeuse_kWhj = 0
RatioConsommationElectriqueSpecifiquePolymere_kWh_m3 = 0


pd_ConsommationElectrique_kWhj = 0


'AGITATION DES DIGESTEURS
If mi_ChoixTypeAgitationDigesteur1 = mi_AGITATION_MECANIQUE Then
    ConsommationElectriqueAgitation_kWhj = md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_VOLUME_m3) * md_ConsommationElectriqueSpecifiqueAgitationMecaniqueDigestion_W_m3 / 1000 * md_TempsFonctionnementAgitation_hj
ElseIf mi_ChoixTypeAgitationDigesteur1 = mi_AGITATION_RECIRCULATION Then
    PertesDeCharge_mCE = 10    'mCE
    DebitRecirculationAgitation_m3h = md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) * md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_VOLUME_m3) / md_TempsFonctionnementAgitation_hj
    RendementPompe = 0.4    'hypothèse
    PuissancePompe_kW = DebitRecirculationAgitation_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * PertesDeCharge_mCE / RendementPompe
    ConsommationElectriqueAgitation_kWhj = PuissancePompe_kW / F_rendement_moteur_pompe(PuissancePompe_kW) * md_TempsFonctionnementAgitation_hj
Else  'if choix_agitatio=mi_AGITATION_BIOGAZ
    VitesseBrassageBiogaz_Nmh = 1.25    'Nm/h   P.JOUAFFRE--> 1 ; Prisma--> entre 1.5 et 2
    PertesDeCharge_mCE = 2    'mCE
    DebitBrassageBiogaz_Nm3j = VitesseBrassageBiogaz_Nmh * md_TempsFonctionnementAgitation_hj * md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_VOLUME_m3) / md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_HAUTEUR_m)
    HMTAgitationBiogaz_mCE = md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_HAUTEUR_m) + PertesDeCharge_mCE
    'EN ATTENTE POUR METTRE LES RENDEMENTS DES SURPRESSEURS CHOISIS
    ConsommationElectriqueSpecifiqueSurpresseurBiogaz_Wh_Nm3mCE = 4    'Wh/(Nm3.mCE)
    ConsommationElectriqueAgitation_kWhj = ConsommationElectriqueSpecifiqueSurpresseurBiogaz_Wh_Nm3mCE * DebitBrassageBiogaz_Nm3j * HMTAgitationBiogaz_mCE / 1000
End If
If mi_ChoixTypeAgitationDigesteur2 = mi_AGITATION_MECANIQUE Then
    ConsommationElectriqueAgitation_kWhj = ConsommationElectriqueAgitation_kWhj + md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_VOLUME_m3) * md_ConsommationElectriqueSpecifiqueAgitationMecaniqueDigestion_W_m3 / 1000 * md_TempsFonctionnementAgitation_hj
ElseIf mi_ChoixTypeAgitationDigesteur2 = mi_AGITATION_RECIRCULATION Then
    PertesDeCharge_mCE = 10    'mCE
    DebitRecirculationAgitation_m3h = md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) * md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_VOLUME_m3) / md_TempsFonctionnementAgitation_hj
    RendementPompe = 0.4    'hypothèse
    PuissancePompe_kW = DebitRecirculationAgitation_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * PertesDeCharge_mCE / RendementPompe
    ConsommationElectriqueAgitation_kWhj = ConsommationElectriqueAgitation_kWhj + PuissancePompe_kW / F_rendement_moteur_pompe(PuissancePompe_kW) * md_TempsFonctionnementAgitation_hj
Else  'if choix_agitatio=mi_AGITATION_BIOGAZ
    VitesseBrassageBiogaz_Nmh = 1.25    'Nm/h   P.JOUAFFRE--> 1 ; Prisma--> entre 1.5 et 2
    PertesDeCharge_mCE = 2    'mCE
    DebitBrassageBiogaz_Nm3j = VitesseBrassageBiogaz_Nmh * md_TempsFonctionnementAgitation_hj * md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_VOLUME_m3) / md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_HAUTEUR_m)
    HMTAgitationBiogaz_mCE = md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_HAUTEUR_m) + PertesDeCharge_mCE
    'EN ATTENTE POUR METTRE LES RENDEMENTS DES SURPRESSEURS CHOISIS
    ConsommationElectriqueSpecifiqueSurpresseurBiogaz_Wh_Nm3mCE = 4    'Wh/(Nm3.mCE)
    ConsommationElectriqueAgitation_kWhj = ConsommationElectriqueAgitation_kWhj + ConsommationElectriqueSpecifiqueSurpresseurBiogaz_Wh_Nm3mCE * DebitBrassageBiogaz_Nm3j * HMTAgitationBiogaz_mCE / 1000
End If
pd_ConsommationElectrique_kWhj = pd_ConsommationElectrique_kWhj + ConsommationElectriqueAgitation_kWhj


'POMPAGE DES BOUES VERS ECHANGEUR (DIGESTEUR 1)
If mi_ChoixTypeChauffageDigesteur1 = mi_DIGESTEUR1_CHAUFFAGE_BOUES_RECIRCULATION Then
    DebitUnitairePompe_m3h = md_EnergieChauffageDigesteur1_kWhj / (md_CapaciteCalorifiqueBoues_kWh_m3°C * md_EcartTemperatureBouesEchangeurDigesteur_°C) / md_TempsFonctionnementRecirculationBoues_hj
    RendementPompe = 0.4    'hypothèse
    PuissancePompe_kW = DebitUnitairePompe_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * md_PressionRefoulementRecirculationBoues_mCE / RendementPompe
    ConsommationElectriqueRecirculationEchangeur_kWhj = PuissancePompe_kW / F_rendement_moteur_pompe(PuissancePompe_kW) * md_TempsFonctionnementRecirculationBoues_hj
    pd_ConsommationElectrique_kWhj = pd_ConsommationElectrique_kWhj + ConsommationElectriqueRecirculationEchangeur_kWhj
End If

'DIFFERENTS POSTES DE POMPAGE  (on garde les mêmes termes que dans les autres fiches pour modification rapide si nécessaire --> usage de variables intermédiaires)
'digesteur 1 --> centrif
DebitTotal_m3j = md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = True Then
    md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)
    If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = True Then
        md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj)
        If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = False Then
            If Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) < (DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) Then
                md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) + 1
            Else
                md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)))
            End If
        End If
    Else
        md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))
        If md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) > pd_NOMBRE_HEURE_PAR_JOUR Then
            md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = pd_NOMBRE_HEURE_PAR_JOUR
            If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = False Then
                If Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) < (DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) Then
                    md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) + 1
                Else
                    md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)))
                End If
                md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))
            End If
        End If
    End If
Else
    md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj))
End If
NombrePompes = md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES)
PressionRefoulement_mCE = md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_HMT_mCE)
TempsFonctionnementChaquePompe_hj = md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj)
DebitUnitairePompe_m3h = md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)
IncoherenceCalcul = (DebitUnitairePompe_m3h * NombrePompes * TempsFonctionnementChaquePompe_hj - DebitTotal_m3j) / DebitTotal_m3j
If Abs(IncoherenceCalcul) > mi_CRITERE_INCOHERENCE Then
    ms_CommentairesCalculs(4) = Feuil27.Range("dig_DLD_erreur_pompage_digesteur1_centrif").Cells(1, 1)
End If
If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = True Then
    md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES)
    RendementGlobalPompe = md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES)
Else
    RendementPompe = 0.4    'hypothèse
    PuissancePompe_kW = DebitUnitairePompe_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * PressionRefoulement_mCE / RendementPompe
    RendementGlobalPompe = RendementPompe * F_rendement_moteur_pompe(PuissancePompe_kW)
    md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = RendementGlobalPompe
End If
ConsommationElectriquePompages_kWhj = DebitUnitairePompe_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * PressionRefoulement_mCE / RendementGlobalPompe * NombrePompes * TempsFonctionnementChaquePompe_hj
'centrif --> lyse
DebitTotal_m3j = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = True Then
    md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)
    If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = True Then
        md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj)
        If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = False Then
            If Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) < (DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) Then
                md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) + 1
            Else
                md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)))
            End If
        End If
    Else
        md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))
        If md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) > pd_NOMBRE_HEURE_PAR_JOUR Then
            md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = pd_NOMBRE_HEURE_PAR_JOUR
            If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = False Then
                If Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) < (DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) Then
                    md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) + 1
                Else
                    md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)))
                End If
                md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))
            End If
        End If
    End If
Else
    md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj))
End If
NombrePompes = md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES)
PressionRefoulement_mCE = md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_HMT_mCE)
TempsFonctionnementChaquePompe_hj = md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj)
DebitUnitairePompe_m3h = md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)
IncoherenceCalcul = (DebitUnitairePompe_m3h * NombrePompes * TempsFonctionnementChaquePompe_hj - DebitTotal_m3j) / DebitTotal_m3j
If Abs(IncoherenceCalcul) > mi_CRITERE_INCOHERENCE Then
    ms_CommentairesCalculs(5) = Feuil27.Range("dig_DLD_erreur_pompage_centrif_lyse").Cells(1, 1)
End If
If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = True Then
    md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES)
    RendementGlobalPompe = md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES)
Else
    RendementPompe = 0.4    'hypothèse
    PuissancePompe_kW = DebitUnitairePompe_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * PressionRefoulement_mCE / RendementPompe
    RendementGlobalPompe = RendementPompe * F_rendement_moteur_pompe(PuissancePompe_kW)
    md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = RendementGlobalPompe
End If
ConsommationElectriquePompages_kWhj = ConsommationElectriquePompages_kWhj + DebitUnitairePompe_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * PressionRefoulement_mCE / RendementGlobalPompe * NombrePompes * TempsFonctionnementChaquePompe_hj
'sortie du digesteur 2
DebitTotal_m3j = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = True Then
    md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)
    If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = True Then
        md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj)
        If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = False Then
            If Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) < (DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) Then
                md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) + 1
            Else
                md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)))
            End If
        End If
    Else
        md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))
        If md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) > pd_NOMBRE_HEURE_PAR_JOUR Then
            md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = pd_NOMBRE_HEURE_PAR_JOUR
            If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = False Then
                If Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) < (DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) Then
                    md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) + 1
                Else
                    md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)))
                End If
                md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))
            End If
        End If
    End If
Else
    md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj))
End If
NombrePompes = md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES)
PressionRefoulement_mCE = md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_HMT_mCE)
TempsFonctionnementChaquePompe_hj = md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj)
DebitUnitairePompe_m3h = md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)
IncoherenceCalcul = (DebitUnitairePompe_m3h * NombrePompes * TempsFonctionnementChaquePompe_hj - DebitTotal_m3j) / DebitTotal_m3j
If Abs(IncoherenceCalcul) > mi_CRITERE_INCOHERENCE Then
    ms_CommentairesCalculs(6) = Feuil27.Range("dig_DLD_erreur_pompage_outlet").Cells(1, 1)
End If
If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = True Then
    md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES)
    RendementGlobalPompe = md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES)
Else
    RendementPompe = 0.4    'hypothèse
    PuissancePompe_kW = DebitUnitairePompe_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * PressionRefoulement_mCE / RendementPompe
    RendementGlobalPompe = RendementPompe * F_rendement_moteur_pompe(PuissancePompe_kW)
    md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = RendementGlobalPompe
End If
ConsommationElectriquePompages_kWhj = ConsommationElectriquePompages_kWhj + DebitUnitairePompe_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * PressionRefoulement_mCE / RendementGlobalPompe * NombrePompes * TempsFonctionnementChaquePompe_hj
pd_ConsommationElectrique_kWhj = pd_ConsommationElectrique_kWhj + ConsommationElectriquePompages_kWhj

'STOCKAGE biogaz
If mi_ChoixStockageBiogaz = mi_STOCKAGE_BIOGAZ Then
    If EquipementValorisationBiogaz = ValorisationBiogazChaudiere Then
        TempsSejourGazometre_h = 3    '3h de temps de séjour pour chaudière
    Else
        TempsSejourGazometre_h = 8    '8h pour autre type de valorisation
    End If
    VolumeGazometre_m3 = md_BiogazProduitNominal_Nm3j / pd_NOMBRE_HEURE_PAR_JOUR * TempsSejourGazometre_h
    If VolumeGazometre_m3 < 1225 Then  'm3
        DebitSoufflageBiogaz_Nm3h = 300  'Nm3/h
    Else
        DebitSoufflageBiogaz_Nm3h = 500   'Nm3/h
    End If
    PertesDeCharge_mCE = 0.25    'mCE
    ConsommationElectriqueSpecifiqueSurpresseurBiogaz_Wh_Nm3mCE = 4    'Wh/(Nm3.mCE)
    ConsommationElectriqueGazometre_kWhj = ConsommationElectriqueSpecifiqueSurpresseurBiogaz_Wh_Nm3mCE * DebitSoufflageBiogaz_Nm3h * PertesDeCharge_mCE / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
    pd_ConsommationElectrique_kWhj = pd_ConsommationElectrique_kWhj + ConsommationElectriqueGazometre_kWhj
End If


'ALIMENTATION CHAUDIERE,MICROTURBINE,COGENERATION
DebitSoufflageBiogaz_Nm3h = pd_BiogazProduit_Nm3j * PartBiogazValorise / pd_NOMBRE_HEURE_PAR_JOUR
If EquipementValorisationBiogaz = ValorisationBiogazChaudiere Then
    PertesDeCharge_mCE = 1    '1 mCE si chaudière
Else
    PertesDeCharge_mCE = 3    '3 mCE si chaudière
End If
ConsommationElectriqueSpecifiqueSurpresseurBiogaz_Wh_Nm3mCE = 4    'Wh/(Nm3.mCE)
ConsommationElectriqueSurpresseurBiogaz_kWhj = ConsommationElectriqueSpecifiqueSurpresseurBiogaz_Wh_Nm3mCE * DebitSoufflageBiogaz_Nm3h * PertesDeCharge_mCE / 1000 * pd_NOMBRE_HEURE_PAR_JOUR
pd_ConsommationElectrique_kWhj = pd_ConsommationElectrique_kWhj + ConsommationElectriqueSurpresseurBiogaz_kWhj

'CONSOMMATION SPECIFIQUE EXELYS DLD
'pompe de contrepression (après refroidissement)
'pompe eau industrielle (dilution)
RatioConsommationElectriqueSpecifiqueDLD_kWh_tMS = 9 + 13    'en kWh/tMS
ConsommationElectriqueSpecifiqueDLD_kWhj = RatioConsommationElectriqueSpecifiqueDLD_kWh_tMS * md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / 1000
pd_ConsommationElectrique_kWhj = pd_ConsommationElectrique_kWhj + ConsommationElectriqueSpecifiqueDLD_kWhj

'CONSOMMATION DE LA CENTRIFUGEUSE (centrif + polymère)
RatioConsommationElectriqueSpecifiqueCentrifugeuse_kWh_m3 = 1.8    'kWh/m3 de boues entrantes   (centrif)
ConsommationElectriqueCentrifugeuse_kWhj = RatioConsommationElectriqueSpecifiqueCentrifugeuse_kWh_m3 * (md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) + md_DebitPolymere_m3j)
RatioConsommationElectriqueSpecifiquePolymere_kWh_m3 = 0.76    'kWh/kg de polymère
ConsommationElectriqueCentrifugeuse_kWhj = ConsommationElectriqueCentrifugeuse_kWhj + RatioConsommationElectriqueSpecifiquePolymere_kWh_m3 * pd_CentrifugeuseConsommationPolymerePur_kgj
pd_ConsommationElectrique_kWhj = pd_ConsommationElectrique_kWhj + ConsommationElectriqueCentrifugeuse_kWhj

pd_ConsommationElectriqueFixe_kWhj = ConsommationElectriqueAgitation_kWhj + ConsommationElectriqueRecirculationEchangeur_kWhj

Exit Sub
'*********************************************************************
erreur_non_geree:
NumeroErreurNonGeree = "27.7"

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
NumeroErreurNonGeree = "27.8"
I = 0

'Ajout du nom du procédé
ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil27.Range("dig_DLD_erreur_nom")

'Ajout des différents commentaires
For I = 1 To mi_NOMBRE_COMMENTAIRES_CALCULS
    If ms_CommentairesCalculs(I) <> "" Then
        ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + ms_CommentairesCalculs(I)
    End If
Next I

ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf

Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True


End Sub

Private Sub calcul_bilan_thermique(ByRef TableauBesoinsThermiques() As Double, ByRef TableauDisponibilitesThermiques() As Double, ByRef PressionVapeur_bar() As Double, ByRef NombreReseauxVapeur As Integer, ByVal RepereCasExploitation As Boolean)


'******************************************************************************************************************************************************
'OBJECTIF

'VARIABLES D'ENTREE:

'VARIABLES DE SORTIE:
'A COMPLETER
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim DebitVolumiqueBouesEntreeDigesteur_m3j As Double
Dim BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj As Double
Dim BesoinThermiqueVaporisationEauPrechauffee_kWhj As Double
Dim EnergieRecupereeEchangeur55_kWhj As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "27.24"
DebitVolumiqueBouesEntreeDigesteur_m3j = 0
BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj = 0
BesoinThermiqueVaporisationEauPrechauffee_kWhj = 0
EnergieRecupereeEchangeur55_kWhj = 0


'CALCUL DES BESOINS THERMIQUES DU DIGESTEUR 1
If mn_EnergieChauffageDigesteur1_kWhj_force = True And RepereCasExploitation = True Then
    md_EnergieChauffageDigesteur1_kWhj = md_EnergieChauffageDigesteur1_kWhj_f
Else
    'pertes_thermiques
    md_EnergieChauffageDigesteur1_kWhj = md_PertesThermiquesDigestion_kWh_m3j(choix_climat, mi_ChoixIsolationDigestion) * md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_VOLUME_m3) / md_EfficaciteEchangeur
    'chauffage des boues
    DebitVolumiqueBouesEntreeDigesteur_m3j = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
    md_EnergieChauffageDigesteur1_kWhj = md_EnergieChauffageDigesteur1_kWhj + DebitVolumiqueBouesEntreeDigesteur_m3j * md_CapaciteCalorifiqueBoues_kWh_m3°C * (md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) - md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)) / md_EfficaciteEchangeur
    md_EnergieChauffageBouesFraichesDigesteur1SansPertes_kWhj = DebitVolumiqueBouesEntreeDigesteur_m3j * md_CapaciteCalorifiqueBoues_kWh_m3°C * (md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) - md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)) / md_EfficaciteEchangeur
End If

If md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) < md_TemperatureLimiteEntreMesophileEtThermophile_°C Then
    TableauBesoinsThermiques(energie_eau_BT) = TableauBesoinsThermiques(energie_eau_BT) + md_EnergieChauffageDigesteur1_kWhj
Else
    TableauBesoinsThermiques(energie_eau_HT) = TableauBesoinsThermiques(energie_eau_HT) + md_EnergieChauffageDigesteur1_kWhj
End If



'CALCUL DES BESOINS DE LA LYSE
'préchauffage de l'eau pour la vapeur et température inter échangeurs
BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj = md_DebitVapeurLyse_kgj * md_CapaciteCalorifiqueEau_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE * (md_TemperaturePrechauffageEauPourProductionVapeur_°C - md_TemperatureInitialeEauPourProductionVapeur_°C) / md_EfficaciteEchangeur
md_TemperatureBouesEchangeur90Echangeur55_°C = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR90, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) - BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj * pd_NOMBRE_SECONDE_PAR_HEURE / (md_CapaciteCalorifiqueBoues_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal * md_FluxIntermediaires(mi_FLUX_ECHANGEUR90_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3)
'vaporisation de l'eau préchauffée
If NombreReseauxVapeur < nb_vapeur_max Then
    NombreReseauxVapeur = NombreReseauxVapeur + 1
    PressionVapeur_bar(NombreReseauxVapeur) = md_PressionAbsolueVapeurSaturanteLyse_bar
    BesoinThermiqueVaporisationEauPrechauffee_kWhj = (md_DebitVapeurLyse_kgj * (md_EnthalpieTotaleVapeurSaturee_kcal_kg - md_CapaciteCalorifiqueEau_kcal_kg°C * md_TemperaturePrechauffageEauPourProductionVapeur_°C) * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE) / md_EfficaciteEchangeur
Else
    GoTo erreur_non_geree
End If
'On détermine les besoins globaux
md_BesoinTotalEnergieProductionVapeur_kWhj = BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj + BesoinThermiqueVaporisationEauPrechauffee_kWhj
md_PinchChauffageEauPourVapeur_kWhj = md_DebitVapeurLyse_kgj * md_CapaciteCalorifiqueEau_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE * (md_TemperatureSaturation_°C - md_TemperatureInitialeEauPourProductionVapeur_°C) / md_EfficaciteEchangeur
md_PinchVaporisation_kWhj = (md_DebitVapeurLyse_kgj * (md_EnthalpieTotaleVapeurSaturee_kcal_kg - md_CapaciteCalorifiqueEau_kcal_kg°C * md_TemperatureSaturation_°C) * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE) / md_EfficaciteEchangeur
If mn_BesoinTotalEnergieProductionVapeur_kWhj_force = True And RepereCasExploitation = True Then
    BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj = BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj * (md_BesoinTotalEnergieProductionVapeur_kWhj_f / md_BesoinTotalEnergieProductionVapeur_kWhj)
    BesoinThermiqueVaporisationEauPrechauffee_kWhj = BesoinThermiqueVaporisationEauPrechauffee_kWhj * (md_BesoinTotalEnergieProductionVapeur_kWhj_f / md_BesoinTotalEnergieProductionVapeur_kWhj)
    md_PinchChauffageEauPourVapeur_kWhj = md_PinchChauffageEauPourVapeur_kWhj * (md_BesoinTotalEnergieProductionVapeur_kWhj_f / md_BesoinTotalEnergieProductionVapeur_kWhj)
    md_PinchVaporisation_kWhj = md_PinchVaporisation_kWhj * (md_BesoinTotalEnergieProductionVapeur_kWhj_f / md_BesoinTotalEnergieProductionVapeur_kWhj)
    md_BesoinTotalEnergieProductionVapeur_kWhj = md_BesoinTotalEnergieProductionVapeur_kWhj_f
End If


TableauBesoinsThermiques(energie_eau_HT) = TableauBesoinsThermiques(energie_eau_HT) + BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj
TableauBesoinsThermiques(energie_eau_HT + NombreReseauxVapeur) = TableauBesoinsThermiques(energie_eau_HT + NombreReseauxVapeur) + BesoinThermiqueVaporisationEauPrechauffee_kWhj


'RECUPERATION D'ENERGIE SUR LES BOUES

If md_TemperatureBouesEchangeur90Echangeur55_°C < md_TemperaturePrechauffageEauPourProductionVapeur_°C + md_EcartMinimalTemperatureEchangeur_°C Then
    'on diminue la valeur de préchauffage de manière aléatoire (5°C)
    md_TemperaturePrechauffageEauPourProductionVapeur_°C = md_TemperaturePrechauffageEauPourProductionVapeur_°C - md_EcartMinimalTemperatureEchangeur_°C
    If iteration = pi_NOMBRE_ITERATION_RETOURS Then
        ms_CommentairesCalculs(3) = Feuil27.Range("dig_DLD_erreur_energie").Cells(1, 1)
    End If
End If

'CALCUL DES TEMPERATURES DES BOUES EN DIFFERENTS POINTS ET DE L'ENERGIE RECUPEREE
'température des boues en entrée de l'échangeur à 90°C --> 165°C température d'hydrolyse
'température des boues entre la dilution et la digestion 2
md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) + (md_PertesThermiquesDigestion_kWh_m3j(choix_climat, mi_ChoixIsolationDigestion) * md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_VOLUME_m3)) / (md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE)
'température avant la dilution
md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = (md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) - md_DebitEauDilution_m3j * md_MasseVolumiqueEau_kg_m3 * md_CapaciteCalorifiqueEau_kcal_kg°C * md_TemperatureEauDilution_°C) / (md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C)
'température des boues entre les deux échangeurs (on modifie la valeur hypothétique de 95°C si cela ne nous permet pas d'entrer dans de bonnes conditions dans le digesteur 2)
md_FluxIntermediaires(mi_FLUX_ECHANGEUR90_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureBouesEchangeur90Echangeur55_°C
If md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) > md_FluxIntermediaires(mi_FLUX_ECHANGEUR90_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) Then
    md_FluxIntermediaires(mi_FLUX_ECHANGEUR90_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
End If

'énergie récupérée sur l'échangeur à 55°C
EnergieRecupereeEchangeur55_kWhj = md_FluxIntermediaires(mi_FLUX_ECHANGEUR90_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE * (md_FluxIntermediaires(mi_FLUX_ECHANGEUR90_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) - md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C))

md_EnergieRefroidissementBouesLysees_kWhj = BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj + EnergieRecupereeEchangeur55_kWhj

If mn_EnergieRefroidissementBouesLysees_kWhj_force = True And RepereCasExploitation = True Then
    EnergieRecupereeEchangeur55_kWhj = EnergieRecupereeEchangeur55_kWhj + md_EnergieRefroidissementBouesLysees_kWhj_f - md_EnergieRefroidissementBouesLysees_kWhj
    md_EnergieRefroidissementBouesLysees_kWhj = md_EnergieRefroidissementBouesLysees_kWhj_f
End If

If EnergieRecupereeEchangeur55_kWhj >= 0 Then
    TableauDisponibilitesThermiques(energie_eau_BT) = TableauDisponibilitesThermiques(energie_eau_BT) + EnergieRecupereeEchangeur55_kWhj
    TableauDisponibilitesThermiques(energie_eau_HT) = TableauDisponibilitesThermiques(energie_eau_HT) + BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj
Else
    TableauDisponibilitesThermiques(energie_eau_BT) = TableauDisponibilitesThermiques(energie_eau_BT) + 0
    TableauDisponibilitesThermiques(energie_eau_HT) = TableauDisponibilitesThermiques(energie_eau_HT) + BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj + EnergieRecupereeEchangeur55_kWhj
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
'       Les valeurs (entiers) associées aux choix se trouvent dans l'onglet "Feuil27" nommé "Exelys_DLD" où sont également reportées les valeurs saisies
'       et où la subroutine va lire les données d'entrée obligatoires utilisateur dans des plages de cellules nommées dans Excel

'VARIABLES DE SORTIE:
'   mi_ChoixEntreesBoues(1 To mi_NOMBRE_FLUX_ENTREE_BOUES): provenance des boues alimentant le procédé
'   md_EntreeBouesPourcentageFluxAdmis(1 To mi_NOMBRE_FLUX_ENTREE_BOUES): pourcentage admis du flux choisi pour alimenter le procédé
'   mi_ProvenanceBouesEntree(1 To mi_NOMBRE_FLUX_ENTREE_BOUES): similaire à mi_ChoixEntreesBoues(1 To mi_NOMBRE_FLUX_ENTREE_BOUES)
'       mais avec un décalage dans les indices car il considère toutes les provenances possible de la file eau comme une seule (boues brutes)
'   mi_TypeBouesEntree(1 To mi_NOMBRE_FLUX_ENTREE_BOUES): si la boue est brute (provenance directe de la file de traitement de l'eau)
'       elle donne le type de traitement de la file eau d'où provient la boue
'A COMPLETER
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "27.1"
I = 0


'Boues en entrée de digestion
For I = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    mi_ChoixEntreesBoues(I) = Feuil27.Range("dig_DLD_choix_inlet").Cells(I, pi_FiliereConsideree)
    If mi_ChoixEntreesBoues(I) <> mi_FLUX_ENTREE_BOUES_NON_UTILISE Then
        md_EntreeBouesPourcentageFluxAdmis(I) = Feuil27.Range("dig_DLD_inlet_ratio").Cells(I, pi_FiliereConsideree)
        If mi_ChoixEntreesBoues(I) <= mi_FLUX_ENTREE_GRAISSES Then
            mi_ProvenanceBouesEntree(I) = boues_inlet
            mi_TypeBouesEntree(I) = mi_ChoixEntreesBoues(I)
        Else
            mi_ProvenanceBouesEntree(I) = boues_epaissies(mi_ChoixEntreesBoues(I) - mi_FLUX_ENTREE_GRAISSES)
            'ici on prendra directement les valeurs du flux antérieur
        End If
        mi_ChoixPositionEntreeBoues(I) = Feuil27.Range("digestionDLD_ChoixEntreesBoues").Cells(I, pi_FiliereConsideree)
    End If
Next I


'agitation
mi_ChoixTypeAgitationDigesteur1 = Feuil27.Range("dig_DLD_choix_agitation").Cells(1, pi_FiliereConsideree)
mi_ChoixTypeAgitationDigesteur2 = Feuil27.Range("digestionDLD_ChoixAgitationDigesteur2").Cells(1, pi_FiliereConsideree)

'chauffage des boues
mi_ChoixTypeChauffageDigesteur1 = Feuil27.Range("dig_DLD_choix_chauffage_boues_dig1").Cells(1, pi_FiliereConsideree)

'stockage_biogaz
mi_ChoixStockageBiogaz = Feuil27.Range("dig_DLD_choix_biogaz_stock").Cells(1, pi_FiliereConsideree)

'isolation
mi_ChoixIsolationDigestion = Feuil27.Range("dig_DLD_choix_isolation").Cells(1, pi_FiliereConsideree)

'choix polymère
pi_ChoixTypePolymereCentrifugeuse = Feuil27.Range("dig_DLD_choix_polymere_type").Cells(1, pi_FiliereConsideree)

'valeurs coferments
md_PotentielProductionMethaneCoferments_Nm3_tMVAppliques = Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_BMP_Nm3CH4_tMVapplique, pi_FiliereConsideree)
md_FacteurSecuritePotentielProductionMethaneCoferments = Feuil6.Range("graisses_externes_caracteristiques").Cells(repere_coferment_BMP_securite, pi_FiliereConsideree)


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
'       Les valeurs sont reportées dans l'onglet "Feuil27" nommée "Exelys_DLD" dont les noms de plages de cellules sont définis dans Excel

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
Dim I As Integer
Dim Colonne As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "27.2"
I = 1
Colonne = pi_FiliereConsideree


'température entrée digesteur
mn_TemperatureBouesEntree_°C_force = Feuil27.Range("dig_DLD_inlet_boues").Cells(mi_PARAMETRE_BOUES_TEMPERATURE_°C, Colonne)
mn_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C_force = Feuil27.Range("dig_DLD_temperature_entree_ech").Cells(1, Colonne)
mn_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C_force = Feuil27.Range("dig_DLD_temperature_sortie_ech").Cells(1, Colonne)
'digesteur 1
For I = 1 To mi_NOMBRE_PARAMETRE_DIGESTION
    mn_CaracteristiquesDigesteur1_force(I) = Feuil27.Range("dig_DLD_digesteur1").Cells(I, Colonne)
    If mn_CaracteristiquesDigesteur1_force(I) = True Then
        md_CaracteristiquesDigesteur1_f(I) = Feuil27.Range("dig_DLD_digesteur1").Cells(I, Colonne)
    End If
Next I
'digesteur 2
For I = 1 To mi_NOMBRE_PARAMETRE_DIGESTION
    mn_CaracteristiquesDigesteur2_force(I) = Feuil27.Range("dig_DLD_digesteur2").Cells(I, Colonne)
    If mn_CaracteristiquesDigesteur2_force(I) = True Then
        md_CaracteristiquesDigesteur2_f(I) = Feuil27.Range("dig_DLD_digesteur2").Cells(I, Colonne)
    End If
Next I
mn_ConcentrationBouesCentrifugeuseMelangeur_gL_force = Feuil27.Range("dig_DLD_centrif_lyse_boues").Cells(mi_PARAMETRE_BOUES_CONCENTRATION_gL, Colonne)
mn_BouesCentrifugeuseMelangeurRatioMVSurMES_force = Feuil27.Range("dig_DLD_centrif_lyse_boues").Cells(mi_PARAMETRE_BOUES_RATIO_MV_MES, Colonne)
mn_TemperatureBouesNonDigerees_°C_force = Feuil27.Range("digestionDLD_EntreeBouesNonDigerees").Cells(mi_PARAMETRE_BOUES_TEMPERATURE_°C, Colonne)
'mn_TemperatureBouesMelangeurLyse_°C_force = Feuil27.Range("digestionDLD_BouesEntreeLyse").Cells(mi_PARAMETRE_BOUES_TEMPERATURE_°C, Colonne)
mn_TemperatureBouesLyseEchangeur90_°C_force = Feuil27.Range("digestionDLD_BouesLyseEchangeur90").Cells(mi_PARAMETRE_BOUES_TEMPERATURE_°C, Colonne)
mn_ConcentrationBouesDilutionDigesteur2_gL_force = Feuil27.Range("dig_DLD_dilution_digesteur2_boues_concentration").Cells(1, Colonne)
mn_DebitVapeurLyse_kgj_force = Feuil27.Range("dig_DLD_lyse_vapeur").Cells(1, Colonne)
mn_PressionAbsolueVapeurSaturanteLyse_bar_force = Feuil27.Range("dig_DLD_lyse_vapeur_pression").Cells(1, Colonne)
mn_TemperatureInitialeEauPourProductionVapeur_°C_force = Feuil27.Range("dig_DLD_vapeur_Teau_ini").Cells(1, Colonne)
mn_TemperaturePrechauffageEauPourProductionVapeur_°C_force = Feuil27.Range("digestionDLD_TemperaturePrechauffageEauPourVapeur").Cells(1, Colonne)
mn_TemperatureEauDilution_°C_force = Feuil27.Range("digestionDLD_TemperatureEauDilution_°C").Cells(1, Colonne)
'pompage digesteur 1 vers centrif
For I = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    mn_CaracteristiquePompageBoues_force(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, I) = Feuil27.Range("dig_DLD_digesteur1_centrif_pompage").Cells(I, Colonne)
    If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, I) = True Then
        md_CaracteristiquePompageBoues_f(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, I) = Feuil27.Range("dig_DLD_digesteur1_centrif_pompage").Cells(I, Colonne)
    End If
Next I
'pompage centrif vers lyse
For I = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    mn_CaracteristiquePompageBoues_force(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, I) = Feuil27.Range("dig_DLD_centrif_lyse_pompage").Cells(I, Colonne)
    If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, I) = True Then
        md_CaracteristiquePompageBoues_f(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, I) = Feuil27.Range("dig_DLD_centrif_lyse_pompage").Cells(I, Colonne)
    End If
Next I
'pompage sortie
For I = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    mn_CaracteristiquePompageBoues_force(mi_POMPAGE_SORTIE_DIGESTEUR2, I) = Feuil27.Range("dig_DLD_outlet_pompage").Cells(I, Colonne)
    If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_SORTIE_DIGESTEUR2, I) = True Then
        md_CaracteristiquePompageBoues_f(mi_POMPAGE_SORTIE_DIGESTEUR2, I) = Feuil27.Range("dig_DLD_outlet_pompage").Cells(I, Colonne)
    End If
Next I
'centrif
For I = 1 To mi_NOMBRE_PARAMETRE_CENTRIFUGEUSE
    mn_CaracteristiqueCentrifugeuse_force(I) = Feuil27.Range("dig_DLD_centrif").Cells(I, Colonne)
    If mn_CaracteristiqueCentrifugeuse_force(I) = True Then
        md_CaracteristiqueCentrifugeuse_f(I) = Feuil27.Range("dig_DLD_centrif").Cells(I, Colonne)
    End If
Next I
mn_RatioCH4DansBiogaz_force = Feuil27.Range("dig_DLD_biogaz_teneur_CH4").Cells(1, Colonne)
mn_BiogazProduit_Nm3j_force = Feuil27.Range("dig_DLD_biogaz_Q_Nm3j").Cells(1, Colonne)
mn_EnergieChauffageDigesteur1_kWhj_force = Feuil27.Range("dig_DLD_energie_eau55").Cells(1, Colonne)
mn_BesoinTotalEnergieProductionVapeur_kWhj_force = Feuil27.Range("dig_DLD_energie_vapeur_kWhj").Cells(1, Colonne)
mn_EnergieRefroidissementBouesLysees_kWhj_force = Feuil27.Range("dig_DLD_energie_refroidissement_boues_kWhj").Cells(1, Colonne)
mn_EnergieDisponibleEau90°C_kWhj_force = Feuil27.Range("dig_DLD_dispo_eau90").Cells(1, Colonne)

If mn_TemperatureBouesEntree_°C_force = True Then
    md_TemperatureBouesEntree_°C_f = Feuil27.Range("dig_DLD_inlet_boues").Cells(mi_PARAMETRE_BOUES_TEMPERATURE_°C, Colonne)
End If
If mn_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C_force = True Then
    md_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C_f = Feuil27.Range("dig_DLD_temperature_entree_ech").Cells(1, Colonne)
End If
If mn_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C_force = True Then
    md_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C_f = Feuil27.Range("dig_DLD_temperature_sortie_ech").Cells(1, Colonne)
End If
If mn_ConcentrationBouesCentrifugeuseMelangeur_gL_force = True Then
    md_ConcentrationBouesCentrifugeuseMelangeur_gL_f = Feuil27.Range("dig_DLD_centrif_lyse_boues").Cells(mi_PARAMETRE_BOUES_CONCENTRATION_gL, Colonne)
End If
If mn_BouesCentrifugeuseMelangeurRatioMVSurMES_force = True Then
    md_BouesCentrifugeuseMelangeurRatioMVSurMES_f = Feuil27.Range("dig_DLD_centrif_lyse_boues").Cells(mi_PARAMETRE_BOUES_RATIO_MV_MES, Colonne)
End If
If mn_TemperatureBouesNonDigerees_°C_force = True Then
    md_TemperatureBouesNonDigerees_°C_f = Feuil27.Range("digestionDLD_EntreeBouesNonDigerees").Cells(mi_PARAMETRE_BOUES_TEMPERATURE_°C, Colonne)
End If
'If mn_TemperatureBouesMelangeurLyse_°C_force = True Then
'    md_TemperatureBouesMelangeurLyse_°C_f = Feuil27.Range("digestionDLD_BouesEntreeLyse").Cells(mi_PARAMETRE_BOUES_TEMPERATURE_°C, Colonne)
'End If
If mn_TemperatureBouesLyseEchangeur90_°C_force = True Then
    md_TemperatureBouesLyseEchangeur90_°C_f = Feuil27.Range("digestionDLD_BouesLyseEchangeur90").Cells(mi_PARAMETRE_BOUES_TEMPERATURE_°C, Colonne)
End If
If mn_ConcentrationBouesDilutionDigesteur2_gL_force = True Then
    md_ConcentrationBouesDilutionDigesteur2_gL_f = Feuil27.Range("dig_DLD_dilution_digesteur2_boues_concentration").Cells(1, Colonne)
End If
If mn_DebitVapeurLyse_kgj_force = True Then
    md_DebitVapeurLyse_kgj_f = Feuil27.Range("dig_DLD_lyse_vapeur").Cells(1, Colonne)
End If
If mn_PressionAbsolueVapeurSaturanteLyse_bar_force = True Then
    md_PressionAbsolueVapeurSaturanteLyse_bar_f = Feuil27.Range("dig_DLD_lyse_vapeur_pression").Cells(1, Colonne)
End If
If mn_TemperatureInitialeEauPourProductionVapeur_°C_force = True Then
    md_TemperatureInitialeEauPourProductionVapeur_°C_f = Feuil27.Range("dig_DLD_vapeur_Teau_ini").Cells(1, Colonne)
End If
If mn_TemperaturePrechauffageEauPourProductionVapeur_°C_force = True Then
    md_TemperaturePrechauffageEauPourProductionVapeur_°C_f = Feuil27.Range("digestionDLD_TemperaturePrechauffageEauPourVapeur").Cells(1, Colonne)
End If
If mn_TemperatureEauDilution_°C_force = True Then
    md_TemperatureEauDilution_°C_f = Feuil27.Range("digestionDLD_TemperatureEauDilution_°C").Cells(1, Colonne)
End If
If mn_RatioCH4DansBiogaz_force = True Then
    md_RatioCH4DansBiogaz_f = Feuil27.Range("dig_DLD_biogaz_teneur_CH4").Cells(1, Colonne)
End If
If mn_BiogazProduit_Nm3j_force = True Then
    md_BiogazProduit_Nm3j_f = Feuil27.Range("dig_DLD_biogaz_Q_Nm3j").Cells(1, Colonne)
End If
If mn_EnergieChauffageDigesteur1_kWhj_force = True Then
    md_EnergieChauffageDigesteur1_kWhj_f = Feuil27.Range("dig_DLD_energie_eau55").Cells(1, Colonne)
End If
If mn_BesoinTotalEnergieProductionVapeur_kWhj_force = True Then
    md_BesoinTotalEnergieProductionVapeur_kWhj_f = Feuil27.Range("dig_DLD_energie_vapeur_kWhj").Cells(1, Colonne)
End If
If mn_EnergieRefroidissementBouesLysees_kWhj_force = True Then
    md_EnergieRefroidissementBouesLysees_kWhj_f = Feuil27.Range("dig_DLD_energie_refroidissement_boues_kWhj").Cells(1, Colonne)
End If
If mn_EnergieDisponibleEau90°C_kWhj_force = True Then
    md_EnergieDisponibleEau90°C_kWhj_f = Feuil27.Range("dig_DLD_dispo_eau90").Cells(1, Colonne)
End If


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
'       qui se fait dans l'onglet "Feui27" nommé "Exelys_DLD" dans des plages de cellules dont les noms sont définis dans Excel
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim Ligne As Integer
Dim Colonne As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "27.9"
Ligne = 1
Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree



For Ligne = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    Feuil27.Range("dig_DLD_inlet_ratio").Cells(Ligne, Colonne) = md_EntreeBouesPourcentageFluxAdmis(Ligne)
Next Ligne

For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("dig_DLD_inlet_boues").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, Ligne)
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("dig_DLD_digesteur1_centrif_boues").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_DIGESTEUR1_CENTRIFUGEUSE, mi_MELANGE_BOUES, Ligne)
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("dig_DLD_centrif_lyse_boues").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, Ligne)
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("digestionDLD_EntreeBouesNonDigerees").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_DIGEREES, mi_MELANGE_BOUES, Ligne)
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("digestionDLD_BouesEntreeLyse").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_LYSE, mi_MELANGE_BOUES, Ligne)
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("digestionDLD_BouesLyseEchangeur90").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR90, mi_MELANGE_BOUES, Ligne)
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("dig_DLD_outlet_boues").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_SORTIE_DIGESTEUR2, mi_MELANGE_BOUES, Ligne)
Next Ligne
For Ligne = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    Feuil27.Range("dig_DLD_digesteur1_centrif_pompage").Cells(Ligne, Colonne) = md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, Ligne)
Next Ligne
For Ligne = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    Feuil27.Range("dig_DLD_centrif_lyse_pompage").Cells(Ligne, Colonne) = md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, Ligne)
Next Ligne
For Ligne = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    Feuil27.Range("dig_DLD_outlet_pompage").Cells(Ligne, Colonne) = md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, Ligne)
Next Ligne
'digesteurs 1 et 2
For Ligne = 1 To mi_NOMBRE_PARAMETRE_DIGESTION
    Feuil27.Range("dig_DLD_digesteur1").Cells(Ligne, Colonne) = md_CaracteristiquesDigesteur1(Ligne)
    Feuil27.Range("dig_DLD_digesteur2").Cells(Ligne, Colonne) = md_CaracteristiquesDigesteur2(Ligne)
Next Ligne
For Ligne = 1 To mi_NOMBRE_PARAMETRE_CENTRIFUGEUSE
    Feuil27.Range("dig_DLD_centrif").Cells(Ligne, Colonne) = md_CaracteristiqueCentrifugeuse(Ligne)
Next Ligne


Ligne = 1

Feuil27.Range("dig_DLD_lyse_vapeur").Cells(Ligne, Colonne) = md_DebitVapeurLyse_kgj
Feuil27.Range("dig_DLD_lyse_vapeur_pression").Cells(Ligne, Colonne) = md_PressionAbsolueVapeurSaturanteLyse_bar
Feuil27.Range("dig_DLD_vapeur_Teau_ini").Cells(1, Colonne) = md_TemperatureInitialeEauPourProductionVapeur_°C
Feuil27.Range("digestionDLD_TemperaturePrechauffageEauPourVapeur").Cells(1, Colonne) = md_TemperaturePrechauffageEauPourProductionVapeur_°C
Feuil27.Range("dig_DLD_dilution_digesteur2_boues_concentration").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR2, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
Feuil27.Range("digestionDLD_DebitEauDilution_m3j").Cells(Ligne, Colonne) = md_DebitEauDilution_m3j
Feuil27.Range("digestionDLD_TemperatureEauDilution_°C").Cells(Ligne, Colonne) = md_TemperatureEauDilution_°C
Feuil27.Range("dig_DLD_reduction_MV_global").Cells(Ligne, Colonne) = md_TauxReductionGlobalMV
Feuil27.Range("dig_DLD_biogaz_teneur_CH4").Cells(Ligne, Colonne) = md_RatioCH4DansBiogaz
Feuil27.Range("dig_DLD_ratio_biogaz_MV").Cells(Ligne, Colonne) = md_ProductionSpecifiqueBiogaz_Nm3_kgMV
Feuil27.Range("dig_DLD_biogaz_Q_Nm3j").Cells(Ligne, Colonne) = pd_BiogazProduit_Nm3j
Feuil27.Range("dig_DLD_energie_eau55").Cells(Ligne, Colonne) = md_EnergieChauffageDigesteur1_kWhj
Feuil27.Range("dig_DLD_energie_vapeur_kWhj").Cells(Ligne, Colonne) = md_BesoinTotalEnergieProductionVapeur_kWhj
Feuil27.Range("dig_DLD_energie_refroidissement_boues_kWhj").Cells(Ligne, Colonne) = md_EnergieRefroidissementBouesLysees_kWhj
Feuil27.Range("dig_DLD_dispo_eau90").Cells(Ligne, Colonne) = md_EnergieDisponibleEau90°C_kWhj
Feuil27.Range("dig_DLD_electricite").Cells(Ligne, Colonne) = pd_ConsommationElectrique_kWhj
Feuil27.Range("dig_DLD_temperature_entree_ech").Cells(Ligne, Colonne) = md_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C
Feuil27.Range("dig_DLD_temperature_sortie_ech").Cells(Ligne, Colonne) = md_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C


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
'       qui se fait dans l'onglet "Feuil27" nommé "Exelys_DLD" dans des plages de cellules dont les noms sont définis dans Excel
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim Ligne As Integer
Dim Colonne As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "27.10"
Ligne = 1
Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


For Ligne = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    Feuil27.Range("dig_DLD_inlet_ratio").Cells(Ligne, Colonne) = ""
Next Ligne

For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("dig_DLD_inlet_boues").Cells(Ligne, Colonne) = ""
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("dig_DLD_digesteur1_centrif_boues").Cells(Ligne, Colonne) = ""
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("dig_DLD_centrif_lyse_boues").Cells(Ligne, Colonne) = ""
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("digestionDLD_EntreeBouesNonDigerees").Cells(Ligne, Colonne) = ""
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("digestionDLD_BouesEntreeLyse").Cells(Ligne, Colonne) = ""
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("digestionDLD_BouesLyseEchangeur90").Cells(Ligne, Colonne) = ""
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil27.Range("dig_DLD_outlet_boues").Cells(Ligne, Colonne) = ""
Next Ligne
For Ligne = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    Feuil27.Range("dig_DLD_digesteur1_centrif_pompage").Cells(Ligne, Colonne) = ""
Next Ligne
For Ligne = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    Feuil27.Range("dig_DLD_centrif_lyse_pompage").Cells(Ligne, Colonne) = ""
Next Ligne
For Ligne = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    Feuil27.Range("dig_DLD_outlet_pompage").Cells(Ligne, Colonne) = ""
Next Ligne
'digesteurs 1 et 2
For Ligne = 1 To mi_NOMBRE_PARAMETRE_DIGESTION
    Feuil27.Range("dig_DLD_digesteur1").Cells(Ligne, Colonne) = ""
    Feuil27.Range("dig_DLD_digesteur2").Cells(Ligne, Colonne) = ""
Next Ligne
For Ligne = 1 To mi_NOMBRE_PARAMETRE_CENTRIFUGEUSE
    Feuil27.Range("dig_DLD_centrif").Cells(Ligne, Colonne) = ""
Next Ligne


Ligne = 1

Feuil27.Range("dig_DLD_dilution_digesteur2_boues_concentration").Cells(Ligne, Colonne) = ""
Feuil27.Range("dig_DLD_lyse_vapeur").Cells(Ligne, Colonne) = ""
Feuil27.Range("dig_DLD_lyse_vapeur_pression").Cells(Ligne, Colonne) = ""
Feuil27.Range("dig_DLD_vapeur_Teau_ini").Cells(1, Colonne) = ""
Feuil27.Range("digestionDLD_TemperaturePrechauffageEauPourVapeur").Cells(1, Colonne) = ""
Feuil27.Range("digestionDLD_DebitEauDilution_m3j").Cells(Ligne, Colonne) = ""
Feuil27.Range("digestionDLD_TemperatureEauDilution_°C").Cells(Ligne, Colonne) = ""
Feuil27.Range("dig_DLD_reduction_MV_global").Cells(Ligne, Colonne) = ""
Feuil27.Range("dig_DLD_biogaz_teneur_CH4").Cells(Ligne, Colonne) = ""
Feuil27.Range("dig_DLD_ratio_biogaz_MV").Cells(Ligne, Colonne) = ""
Feuil27.Range("dig_DLD_biogaz_Q_Nm3j").Cells(Ligne, Colonne) = ""
Feuil27.Range("dig_DLD_energie_eau55").Cells(Ligne, Colonne) = ""
Feuil27.Range("dig_DLD_energie_vapeur_kWhj").Cells(Ligne, Colonne) = ""
Feuil27.Range("dig_DLD_energie_refroidissement_boues_kWhj").Cells(Ligne, Colonne) = ""
Feuil27.Range("dig_DLD_dispo_eau90").Cells(Ligne, Colonne) = ""
Feuil27.Range("dig_DLD_electricite").Cells(Ligne, Colonne) = ""
Feuil27.Range("dig_DLD_temperature_entree_ech").Cells(Ligne, Colonne).ClearContents
Feuil27.Range("dig_DLD_temperature_sortie_ech").Cells(Ligne, Colonne).ClearContents


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
NumeroErreurNonGeree = "27.3"


'réduction des MV référence par type de boues   'modifié par DCA le 31/05/11 d'après une comparaison NDNc+PDN avec dim-digesteur
'primaire
md_ReductionMatieresVolatilesBouesNonLysees(I_simple) = 0.638    '0.61
md_ReductionMatieresVolatilesBouesNonLysees(I_reactif) = 0.614  '0.57
'secondaire
md_ReductionMatieresVolatilesBouesNonLysees(II_forte) = 0.546    '0.52
md_ReductionMatieresVolatilesBouesNonLysees(II_moyenne) = 0.501    '0.4
md_ReductionMatieresVolatilesBouesNonLysees(II_faible_EB) = 0.378    '0.34
md_ReductionMatieresVolatilesBouesNonLysees(II_faible_ED) = 0.32   '0.23
md_ReductionMatieresVolatilesBouesNonLysees(II_prolongee_EB) = 0.358   ' 0.34
md_ReductionMatieresVolatilesBouesNonLysees(II_prolongee_ED) = 0.3  ' 0.23
md_ReductionMatieresVolatilesBouesNonLysees(II_MBR) = 0.3    '0.2
md_ReductionMatieresVolatilesBouesNonLysees(II_MBBR) = 0.358    '0.25
md_ReductionMatieresVolatilesBouesNonLysees(II_HybAS) = 0.358    '0.3
md_ReductionMatieresVolatilesBouesNonLysees(II_biostyr_C) = 0.48  '0.45
md_ReductionMatieresVolatilesBouesNonLysees(II_biostyr_N) = 0.48  '0.35
md_ReductionMatieresVolatilesBouesNonLysees(II_biostyr_NDN) = 0.48  '0.35
md_ReductionMatieresVolatilesBouesNonLysees(II_biostyr_PDN) = 0.48  '0.25
'tertiaire
md_ReductionMatieresVolatilesBouesNonLysees(III_decantation) = 0.157    '0.1
md_ReductionMatieresVolatilesBouesNonLysees(III_biostyr_N) = 0.157    '0.1
md_ReductionMatieresVolatilesBouesNonLysees(III_biostyr_PDN) = 0.157    '0.1
'graisses
md_ReductionMatieresVolatilesBouesNonLysees(codigestion_graisses) = 0.629    '0.6

'réduction des MV référence par type de boues (boues préalablement lysées)   'modifié par DCA le 31/05/11 d'après une comparaison NDNc+PDN avec dim-digesteur
'primaire
md_ReductionMatieresVolatilesBouesLysees(I_simple) = 0.673
md_ReductionMatieresVolatilesBouesLysees(I_reactif) = 0.678
'secondaire
md_ReductionMatieresVolatilesBouesLysees(II_forte) = 0.553
md_ReductionMatieresVolatilesBouesLysees(II_moyenne) = 0.533
md_ReductionMatieresVolatilesBouesLysees(II_faible_EB) = 0.472
md_ReductionMatieresVolatilesBouesLysees(II_faible_ED) = 0.419
md_ReductionMatieresVolatilesBouesLysees(II_prolongee_EB) = 0.451
md_ReductionMatieresVolatilesBouesLysees(II_prolongee_ED) = 0.394
md_ReductionMatieresVolatilesBouesLysees(II_MBR) = 0.394
md_ReductionMatieresVolatilesBouesLysees(II_MBBR) = 0.451
md_ReductionMatieresVolatilesBouesLysees(II_HybAS) = 0.451
md_ReductionMatieresVolatilesBouesLysees(II_biostyr_C) = 0.493
md_ReductionMatieresVolatilesBouesLysees(II_biostyr_N) = 0.493
md_ReductionMatieresVolatilesBouesLysees(II_biostyr_NDN) = 0.493
md_ReductionMatieresVolatilesBouesLysees(II_biostyr_PDN) = 0.493
'tertiaire
md_ReductionMatieresVolatilesBouesLysees(III_decantation) = 0.161
md_ReductionMatieresVolatilesBouesLysees(III_biostyr_N) = 0.161
md_ReductionMatieresVolatilesBouesLysees(III_biostyr_PDN) = 0.161
'graisses
md_ReductionMatieresVolatilesBouesLysees(codigestion_graisses) = 0.765

'réduction dans le deuxième digesteur de boues digérées puis lysées
md_ReductionMatieresVolatilesBouesDigereesLysees = 0.3656   'à 15 jours 35% de réduction des MV donc recalculé à 20 jours suivant la relation de correction du SRT

'temps de séjour référence par type de digestion    en jours    'on est obligé de regarder si la température a été forcée ou non pour savoir si on est en thermophile
md_TemperatureLimiteEntreMesophileEtThermophile_°C = 50
If mn_CaracteristiquesDigesteur1_force(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = True Then
    If md_CaracteristiquesDigesteur1_f(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) <= md_TemperatureLimiteEntreMesophileEtThermophile_°C Then
        md_ReferenceTempsSejourBouesDigesteur1_j = 20
    Else
        md_ReferenceTempsSejourBouesDigesteur1_j = 12
    End If
Else
    md_ReferenceTempsSejourBouesDigesteur1_j = 20
End If
If mn_CaracteristiquesDigesteur2_force(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = True Then
    If md_CaracteristiquesDigesteur2_f(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) <= md_TemperatureLimiteEntreMesophileEtThermophile_°C Then
        md_ReferenceTempsSejourBouesDigesteur2_j = 15
    Else
        md_ReferenceTempsSejourBouesDigesteur2_j = 15
    End If
Else
    md_ReferenceTempsSejourBouesDigesteur2_j = 15
End If

'température référence par type de digestion   en °C
md_TemperatureReferenceDigesteur1_°C = 35
md_TemperatureReferenceDigesteur2_°C = 38

'On attribue une production spécifique de biogaz par défaut qui sera ensuite calculée pour ne pas avoir de division par 0 lors des calculs de cocktail des boues en présence de coferments (1ère itération)
md_ProductionSpecifiqueBiogaz_Nm3_kgMV = 1

'relation de juillet 2011 d'Eric GUIBELIN    modifié par DCA
md_CoefficientACalculFacteurCorrectionTempsSejourBouesDigesteur = 1.25
md_CoefficientBCalculFacteurCorrectionTempsSejourBouesDigesteur = 0.7

'calcul de la densité du biogaz en fonction de la fraction de CH4
md_CoefficientsCalculDensiteBiogaz(1) = (16 - 44) / 22.4
md_CoefficientsCalculDensiteBiogaz(0) = 44 / 22.4

'0.35 Nm3 de CH4 produit par kg de DCO méthanisée
md_ProductionSpecifiqueMethane_Nm3_kgDCO = 0.35

'hypothèses sur la solubilisation de la pollution Pt, DCO, NK
md_RatioNKSolubleBouesDigereesSurMVDegrades = 0.9  '%NK solubilisé par MV dégradés   source prédimboo
md_RatioPSolubleBouesDigereesSurPBouesEntree = 0.16    'part du phosphore particulaire solubilisé  source prédimboo
md_RatioDCOSolubleSurDCOTotaleBouesDigereesDigesteur1 = 0.05    'part de DCO soluble en sortie
md_RatioDCOSolubleSurDCOTotaleBouesDigereesDigesteur2 = 0.2    'part de DCO soluble en sortie

'calcul de la part MV soluble en fonction du temps de séjour (référence 20j mésophile, on considère que ça équivaut à 12j en thermophile)
'primaire
md_RatioMVSolublesMVTotauxBouesDigereesReference(I_simple) = 6.49 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(I_reactif) = 5.99 / 100
'secondaire
md_RatioMVSolublesMVTotauxBouesDigereesReference(II_forte) = 4.09 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(II_moyenne) = 3.51 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(II_faible_EB) = 1.5 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(II_faible_ED) = 1.5 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(II_prolongee_EB) = 1.5 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(II_prolongee_ED) = 1.5 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(II_MBR) = 1.5 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(II_MBBR) = 1.5 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(II_HybAS) = 1.5 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(II_biostyr_C) = 2.49 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(II_biostyr_N) = 2.49 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(II_biostyr_NDN) = 2.49 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(II_biostyr_PDN) = 2.49 / 100
'tertiaire
md_RatioMVSolublesMVTotauxBouesDigereesReference(III_decantation) = 0.64 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(III_biostyr_N) = 0.64 / 100
md_RatioMVSolublesMVTotauxBouesDigereesReference(III_biostyr_PDN) = 0.64 / 100
'graisses
md_RatioMVSolublesMVTotauxBouesDigereesReference(codigestion_graisses) = 7.26 / 100

'coefficient directeur de la relation concernant les MV soluble en fonction du temps de séjour réduit lorsqu'il est inférieur à 1
'primaire
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(I_simple) = -3.3846
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(I_reactif) = -3.5276
'secondaire
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(II_forte) = -3.9355
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(II_moyenne) = -4.0652
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(II_faible_EB) = -4.62
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(II_faible_ED) = -4.62
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(II_prolongee_EB) = -4.62
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(II_prolongee_ED) = -4.62
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(II_MBR) = -4.62
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(II_MBBR) = -4.62
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(II_HybAS) = -4.62
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(II_biostyr_C) = -4.3336
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(II_biostyr_N) = -4.3336
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(II_biostyr_NDN) = -4.3336
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(II_biostyr_PDN) = -4.3336
'tertiaire
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(III_decantation) = -5.1835
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(III_biostyr_N) = -5.1835
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(III_biostyr_PDN) = -5.1835
'graisses
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(codigestion_graisses) = -2.5583
'coefficient directeur de la relation concernant les MV soluble en fonction du temps de séjour réduit lorsqu'il est supérieur à 1
'primaire
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(I_simple) = -1.0293
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(I_reactif) = -1.0551
'secondaire
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(II_forte) = -1.0905
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(II_moyenne) = -1.0997
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(II_faible_EB) = -1.14
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(II_faible_ED) = -1.14
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(II_prolongee_EB) = -1.14
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(II_prolongee_ED) = -1.14
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(II_MBR) = -1.14
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(II_MBBR) = -1.14
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(II_HybAS) = -1.14
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(II_biostyr_C) = -1.1165
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(II_biostyr_N) = -1.1165
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(II_biostyr_NDN) = -1.1165
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(II_biostyr_PDN) = -1.1165
'tertiaire
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(III_decantation) = -1.2188
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(III_biostyr_N) = -1.2188
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(III_biostyr_PDN) = -1.2188
'graisses
md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(codigestion_graisses) = -0.6997

'concentration du polymère injecté
md_ConcentrationPolymereCentrifugeuse_gL = 1.5   'g/L

'hypothèses sur les températures
md_PerteTemperatureBouesDigereesDigesteur1AvantLyse_°C = 10  '°C pertes thermiques en degrés entre digesteur 1 et entrée de l'hydrolyse thermique
md_RatioPertesVapeur = 0.05    '5% de pertes thermiques sur la vapeur de la lyse
md_MasseVolumiqueEau_kg_m3 = 1000

'hypothèses sur les pertes du digesteur (en fonction du climat et de l'isolation du digesteur) en kWh/(m3.j)
md_PertesThermiquesDigestion_kWh_m3j(climat_froid, mi_DIGESTION_ISOLEE) = 0.6
md_PertesThermiquesDigestion_kWh_m3j(climat_froid, mi_DIGESTION_NON_ISOLEE) = 3.72
md_PertesThermiquesDigestion_kWh_m3j(climat_tempere, mi_DIGESTION_ISOLEE) = 0.36
md_PertesThermiquesDigestion_kWh_m3j(climat_tempere, mi_DIGESTION_NON_ISOLEE) = 2.2
md_PertesThermiquesDigestion_kWh_m3j(climat_chaud, mi_DIGESTION_ISOLEE) = 0.24
md_PertesThermiquesDigestion_kWh_m3j(climat_chaud, mi_DIGESTION_NON_ISOLEE) = 1.44

'efficacité de l'échangeur pour chauffer les boues
md_EfficaciteEchangeur = 0.95

'capacité calorifique des boues   en kWh/(m3.°C)
md_CapaciteCalorifiqueBoues_kWh_m3°C = 1.163
md_CapaciteCalorifiqueBoues_kcal_kg°C = 1
md_CapaciteCalorifiqueEau_kcal_kg°C = 1

'part de NH4 dans les NK soluble (pour connaître le NH4 dans les retours
md_RatioNH4SurNKSolubleBouesDigerees = 0.95

'part de DBO dans la DCO soluble
md_RatioDBOSolubleSurDCOSolubleBouesDigerees = 0.2 * 0.8  'cf Christophe AMIEL

'masse volumique des boues (kg/m3)
md_MasseVolumiqueBoues_kg_m3 = 1000

'coefficients nécessaires au calcul de la température de vapeur saturée à P   (en °C)
'    md_TemperatureSaturation_°C(P) = a1/(a2 - ln(a3*P)) + a4 - a5   avec P = pression absolue en bar
md_CoefficientsCalculTemperatureSaturationAPression(1) = 3816.44
md_CoefficientsCalculTemperatureSaturationAPression(2) = 18.3
md_CoefficientsCalculTemperatureSaturationAPression(3) = 750
md_CoefficientsCalculTemperatureSaturationAPression(4) = 46.13
md_CoefficientsCalculTemperatureSaturationAPression(5) = 273.15

'coefficients pour calcul de l'enthalpie totale de la vapeur saturée à T   (en kcal/kg)  Hvap_sat(T) = a0 + a1*T
md_CoefficientsCalculEnthalpieVaporisationATemperature(0) = 616.87
md_CoefficientsCalculEnthalpieVaporisationATemperature(1) = -0.7542

'température minimale d'écart sur l'échangeur de préchauffage de la vapeur
md_EcartMinimalTemperatureEchangeur_°C = 5    '°C
md_EcartTemperatureBouesEchangeurDigesteur_°C = 3    ' 3°C environ pour la différence entre la T en entrée et en sortie de l'échangeur côté boues

'Concentration souhaitée des boues diluées
md_ConcentrationBouesDilutionDigesteur2Souhaitee_gL = 100

'Hypothèses concernant l'agitation du digesteur
md_ConsommationElectriqueSpecifiqueAgitationMecaniqueDigestion_W_m3 = 7    'W/m3   E. GUIBELIN    modifié le 23/04/2013 par DCA (2W/m3 avant)
md_TempsFonctionnementAgitation_hj = 24

'Hypothèse concernant la recirculation des boues vers l'échangeur
md_TempsFonctionnementRecirculationBoues_hj = 24    'h/j
md_PressionRefoulementRecirculationBoues_mCE = 10   'mCE



Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True


End Sub

Private Function F_rendement_moteur_pompe(ByVal PuissancePompeP2_kW As Double) As Double

'************************************************************************************************************
'OBJECTIF:
'   Cette procédure calcule le rendement moteur d'une pompe
'************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "44.30"

'calcul de la puissance de pompe
If PuissancePompeP2_kW < 60 Then
    F_rendement_moteur_pompe = (5.7195 * Log(PuissancePompeP2_kW) + 72.682) / 100
Else
    F_rendement_moteur_pompe = 0.961
End If


Exit Function
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Function

Private Function F_facteur_correction_temperature_digesteur(ByVal TemperatureDigesteur_°C As Double) As Double

'************************************************************************************************************
'OBJECTIF:
'   Cette procédure calcule le facteur de correction du rendement de dégradation des MV lié à la température.
'************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "27.29"

'correctif de température de digestion (digesteur)
If TemperatureDigesteur_°C < 0 Then
    F_facteur_correction_temperature_digesteur = 0
ElseIf TemperatureDigesteur_°C < 35 Then
    F_facteur_correction_temperature_digesteur = TemperatureDigesteur_°C / 35
ElseIf TemperatureDigesteur_°C < 38 Then
    F_facteur_correction_temperature_digesteur = 1
ElseIf TemperatureDigesteur_°C < 45 Then
    F_facteur_correction_temperature_digesteur = -TemperatureDigesteur_°C / 14 + (3 + 5 / 7)
ElseIf TemperatureDigesteur_°C < 55 Then
    F_facteur_correction_temperature_digesteur = 5 * TemperatureDigesteur_°C / 91 - 1.975
ElseIf TemperatureDigesteur_°C < 58 Then
    F_facteur_correction_temperature_digesteur = 1.05
ElseIf TemperatureDigesteur_°C < 80 Then
    F_facteur_correction_temperature_digesteur = -TemperatureDigesteur_°C / 21 + (3 + 9 / 11)
Else
    F_facteur_correction_temperature_digesteur = 0
End If


Exit Function
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Function

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
Dim I As Integer
Dim j As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "27.4"
I = 0
j = 0

'températures pour la digestion
If mn_CaracteristiquesDigesteur1_force(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = True Then
    md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = md_CaracteristiquesDigesteur1_f(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C)
Else
    md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = md_TemperatureReferenceDigesteur1_°C
End If
If mn_CaracteristiquesDigesteur2_force(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = True Then
    If md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) > md_TemperatureLimiteEntreMesophileEtThermophile_°C Then
        MsgBox Feuil27.Range("digestionDLD_erreur_TemperatureDigestion") & " " & md_TemperatureLimiteEntreMesophileEtThermophile_°C & " °C", Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    Else
        md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = md_CaracteristiquesDigesteur2_f(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C)
    End If
Else
    md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = md_TemperatureReferenceDigesteur2_°C
End If

'SRT en j pour le design et initialisation du SRT réel
If mn_CaracteristiquesDigesteur1_force(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j) = True Then
    md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j) = md_CaracteristiquesDigesteur1_f(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j)
Else
    md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j) = 12
End If
If mn_CaracteristiquesDigesteur2_force(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j) = True Then
    md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j) = md_CaracteristiquesDigesteur2_f(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j)
Else
    md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j) = 15
End If

'paramètres des digesteurs intervenant dans l'agitation
If mi_ChoixTypeAgitationDigesteur1 = mi_AGITATION_BIOGAZ Then
    If mn_CaracteristiquesDigesteur1_force(mi_PARAMETRE_DIGESTION_HAUTEUR_m) = True Then
        md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_HAUTEUR_m) = md_CaracteristiquesDigesteur1_f(mi_PARAMETRE_DIGESTION_HAUTEUR_m)
    Else
        md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_HAUTEUR_m) = 10
    End If
ElseIf mi_ChoixTypeAgitationDigesteur1 = mi_AGITATION_RECIRCULATION Then
    If mn_CaracteristiquesDigesteur1_force(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) = True Then
        md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) = md_CaracteristiquesDigesteur1_f(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j)
    Else
        md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) = 5  'ici c'est un taux par rapport au volume total de digesteur, le vrai taux sera recalculé dans conso élec
    End If
End If
If mi_ChoixTypeAgitationDigesteur2 = mi_AGITATION_BIOGAZ Then
    If mn_CaracteristiquesDigesteur2_force(mi_PARAMETRE_DIGESTION_HAUTEUR_m) = True Then
        md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_HAUTEUR_m) = md_CaracteristiquesDigesteur2_f(mi_PARAMETRE_DIGESTION_HAUTEUR_m)
    Else
        md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_HAUTEUR_m) = 10
    End If
ElseIf mi_ChoixTypeAgitationDigesteur2 = mi_AGITATION_RECIRCULATION Then
    If mn_CaracteristiquesDigesteur2_force(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) = True Then
        md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) = md_CaracteristiquesDigesteur2_f(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j)
    Else
        md_CaracteristiquesDigesteur2(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) = 5  'ici c'est un taux par rapport au volume total de digesteur, le vrai taux sera recalculé dans conso élec
    End If
End If

'paramètres de la centrif (taux de capture en MES en %, taux polymère en kg/tMS et concentration des boues en g/L)
If mn_CaracteristiqueCentrifugeuse_force(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES) Then
    md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES) = md_CaracteristiqueCentrifugeuse_f(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES)
Else
    md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_TAUX_CAPTURE_MES) = 0.97
End If
If mn_CaracteristiqueCentrifugeuse_force(mi_PARAMETRE_CENTRIFUGEUSE_CONSOMMATION_POLYMERE_kgMA_tMS) Then
    md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_CONSOMMATION_POLYMERE_kgMA_tMS) = md_CaracteristiqueCentrifugeuse_f(mi_PARAMETRE_CENTRIFUGEUSE_CONSOMMATION_POLYMERE_kgMA_tMS)
Else
    md_CaracteristiqueCentrifugeuse(mi_PARAMETRE_CENTRIFUGEUSE_CONSOMMATION_POLYMERE_kgMA_tMS) = 10
End If
If mn_ConcentrationBouesCentrifugeuseMelangeur_gL_force = True Then
    md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_ConcentrationBouesCentrifugeuseMelangeur_gL_f
Else
    md_FluxIntermediaires(mi_FLUX_CENTRIFUGEUSE_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = 230
End If

'pression de la vapeur à utiliser   (en bars)
If mn_PressionAbsolueVapeurSaturanteLyse_bar_force = True Then
    md_PressionAbsolueVapeurSaturanteLyse_bar = md_PressionAbsolueVapeurSaturanteLyse_bar_f
Else
    md_PressionAbsolueVapeurSaturanteLyse_bar = 15
End If
'calcul de la température de saturation de la vapeur à P      md_TemperatureSaturation_°C(P) = a1/(a2 - ln(a3*P)) + a4 - a5
md_TemperatureSaturation_°C = md_CoefficientsCalculTemperatureSaturationAPression(1) / (md_CoefficientsCalculTemperatureSaturationAPression(2) - Log(md_CoefficientsCalculTemperatureSaturationAPression(3) * md_PressionAbsolueVapeurSaturanteLyse_bar)) + md_CoefficientsCalculTemperatureSaturationAPression(4) - md_CoefficientsCalculTemperatureSaturationAPression(5)
'calcul de l'enthalpie de vaporisation de l'eau à T
md_EnthalpieVaporisationATemperature_kcal_kg = 0
For I = 0 To 1
    md_EnthalpieVaporisationATemperature_kcal_kg = md_EnthalpieVaporisationATemperature_kcal_kg + md_CoefficientsCalculEnthalpieVaporisationATemperature(I) * md_TemperatureSaturation_°C ^ I
Next I
'calcul de l'enthalpie totale de la vapeur à T
md_EnthalpieTotaleVapeurSaturee_kcal_kg = md_CapaciteCalorifiqueEau_kcal_kg°C * (md_TemperatureSaturation_°C - 0) + md_EnthalpieVaporisationATemperature_kcal_kg

'Température de l'eau utilisée pour produire la vapeur
If mn_TemperatureInitialeEauPourProductionVapeur_°C_force = True Then
    md_TemperatureInitialeEauPourProductionVapeur_°C = md_TemperatureInitialeEauPourProductionVapeur_°C_f
Else
    md_TemperatureInitialeEauPourProductionVapeur_°C = 15
End If

'Température de préchauffage de l'eau utilisée pour faire la vapeur
If mn_TemperaturePrechauffageEauPourProductionVapeur_°C_force = True Then
    md_TemperaturePrechauffageEauPourProductionVapeur_°C = md_TemperaturePrechauffageEauPourProductionVapeur_°C_f
Else
    md_TemperaturePrechauffageEauPourProductionVapeur_°C = 95    '°C 'MSM 18/01/13
End If

'Température en entrée du premier échangeur de refroidissement
If mn_TemperatureBouesLyseEchangeur90_°C_force = True Then
    md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR90, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureBouesLyseEchangeur90_°C_f   '°C
Else
    md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR90, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = 165    '°C
End If

'convoyage/pompage des boues
md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = 2
md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_HMT_mCE) = 50   'mCE
md_CaracteristiquePompageBoues(mi_POMPAGE_DIGESTEUR1_CENTRIFUGEUSE, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = 24   'h/j
md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = 1
md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_HMT_mCE) = 120   'mCE
md_CaracteristiquePompageBoues(mi_POMPAGE_CENTRIFUGEUSE_MELANGEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = 24   'h/j
md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = 2
md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_HMT_mCE) = 50   'mCE
md_CaracteristiquePompageBoues(mi_POMPAGE_SORTIE_DIGESTEUR2, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = 24   'h/j
For I = 1 To mi_NOMBRE_POMPAGE
    For j = 1 To mi_NOMBRE_PARAMETRES_POMPAGE - 1
        If mn_CaracteristiquePompageBoues_force(I, j) = True Then
            md_CaracteristiquePompageBoues(I, j) = md_CaracteristiquePompageBoues_f(I, j)
        End If
    Next j
Next I

'biogaz %CH4
If mn_RatioCH4DansBiogaz_force = True Then
    md_RatioCH4DansBiogaz = md_RatioCH4DansBiogaz_f
Else
    md_RatioCH4DansBiogaz = 0.63
End If


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

'VARIABLES DE SORTIE:
'A COMPLETER
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim NomFlux As String
Dim TemperatureEntreeEchangeur_°C As Double
Dim TemperatureSortieEchangeur_°C As Double
Dim MCp_kWh_j°C As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "27.11"
NomFlux = ""
TemperatureEntreeEchangeur_°C = 0
TemperatureSortieEchangeur_°C = 0
MCp_kWh_j°C = 0


'DIGESTEUR 1 - Chauffage des boues

'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
'    NomFlux = "DIGESTEUR 1 - Chauffage des boues"
NomFlux = Feuil5.Range("nom_flux_digesteur1").Cells(1, 1).Value
If mi_ChoixTypeChauffageDigesteur1 = mi_DIGESTEUR1_CHAUFFAGE_BOUES_AMONT Then
    md_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C = md_FluxIntermediaires(mi_FLUX_ENTREE_DIGESTEUR1, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
ElseIf mi_ChoixTypeChauffageDigesteur1 = mi_DIGESTEUR1_CHAUFFAGE_BOUES_RECIRCULATION Then
    If mn_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C_force = True Then
        md_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C = md_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C_f
    Else
        md_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C = md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C)
    End If
End If
TemperatureEntreeEchangeur_°C = md_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C
If mn_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C_force = True Then
    md_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C = md_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C_f
Else
    If mi_ChoixTypeChauffageDigesteur1 = mi_DIGESTEUR1_CHAUFFAGE_BOUES_AMONT Then
        md_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C = md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) + (md_EnergieChauffageDigesteur1_kWhj - md_EnergieChauffageBouesFraichesDigesteur1SansPertes_kWhj) / md_EnergieChauffageBouesFraichesDigesteur1SansPertes_kWhj * (md_CaracteristiquesDigesteur1(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) - md_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C)
    ElseIf mi_ChoixTypeChauffageDigesteur1 = mi_DIGESTEUR1_CHAUFFAGE_BOUES_RECIRCULATION Then
        md_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C = md_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C + md_EcartTemperatureBouesEchangeurDigesteur_°C
    End If
End If
TemperatureSortieEchangeur_°C = md_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C

'MCp_kWh_j°C
MCp_kWh_j°C = md_EnergieChauffageDigesteur1_kWhj / (TemperatureSortieEchangeur_°C - TemperatureEntreeEchangeur_°C)

'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
Call copie_dans_flux_thermique_process(NomFlux, TemperatureEntreeEchangeur_°C, TemperatureSortieEchangeur_°C, MCp_kWh_j°C)



'LYSE - Chauffage de l'eau sous pression

'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
'    NomFlux = "LYSE - Chauffage de l'eau pour produire de la vapeur"
NomFlux = Feuil5.Range("nom_flux_digesteur2").Cells(1, 1).Value
TemperatureEntreeEchangeur_°C = md_TemperatureInitialeEauPourProductionVapeur_°C
TemperatureSortieEchangeur_°C = md_TemperatureSaturation_°C
MCp_kWh_j°C = md_PinchChauffageEauPourVapeur_kWhj / (TemperatureSortieEchangeur_°C - TemperatureEntreeEchangeur_°C)

'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
Call copie_dans_flux_thermique_process(NomFlux, TemperatureEntreeEchangeur_°C, TemperatureSortieEchangeur_°C, MCp_kWh_j°C)



'LYSE - Vaporisation

'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
'    NomFlux = "LYSE - Vaporisation de l'eau"
NomFlux = Feuil5.Range("nom_flux_digesteur3").Cells(1, 1).Value
TemperatureEntreeEchangeur_°C = md_TemperatureSaturation_°C
TemperatureSortieEchangeur_°C = md_TemperatureSaturation_°C + 0.01
MCp_kWh_j°C = md_PinchVaporisation_kWhj / (TemperatureSortieEchangeur_°C - TemperatureEntreeEchangeur_°C)

'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
Call copie_dans_flux_thermique_process(NomFlux, TemperatureEntreeEchangeur_°C, TemperatureSortieEchangeur_°C, MCp_kWh_j°C)



'DLD - Refroidissement des boues avant dilution

'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
'    NomFlux = "Exelys DLD - Refroidissement des boues avant dilution"
NomFlux = Feuil5.Range("nom_flux_digesteur4").Cells(1, 1).Value
TemperatureEntreeEchangeur_°C = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR90, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
TemperatureSortieEchangeur_°C = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
MCp_kWh_j°C = md_EnergieRefroidissementBouesLysees_kWhj / (TemperatureEntreeEchangeur_°C - TemperatureSortieEchangeur_°C)

'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
Call copie_dans_flux_thermique_process(NomFlux, TemperatureEntreeEchangeur_°C, TemperatureSortieEchangeur_°C, MCp_kWh_j°C)



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

Private Sub initialisation_variables_module()


'******************************************************************************************************************************************************
'OBJECTIF:
'   Cette procédure initialise toutes les variables déclarées au niveau du module

'VARIABLES D'ENTREE:

'VARIABLES DE SORTIE:
'   toutes les variables déclarées au niveau du module de classe
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer
Dim k As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "27.0"
I = 0
j = 0
k = 0

'variables contenant les choix relatifs au procédé
For I = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    mi_ProvenanceBouesEntree(I) = 0
    mi_TypeBouesEntree(I) = 0
    mi_ChoixEntreesBoues(I) = 0
    mi_ChoixPositionEntreeBoues(I) = 0
Next I
mi_ChoixTypeAgitationDigesteur1 = 0
mi_ChoixTypeAgitationDigesteur2 = 0
pi_ChoixTypePolymereCentrifugeuse = 0
mi_ChoixStockageBiogaz = 0
mi_ChoixIsolationDigestion = 0
mi_ChoixTypeChauffageDigesteur1 = 0

'variables contenant les valeurs des paramètres qui sont affichés à l'utilisateur en fin de calcul
For I = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    md_EntreeBouesPourcentageFluxAdmis(I) = 0
Next I
For I = 1 To mi_NOMBRE_FLUX_INTERMEDIAIRES
    For j = 1 To nb_type_boues + 1
        For k = 1 To mi_NOMBRE_PARAMETRES_BOUES
            md_FluxIntermediaires(I, j, k) = 0
        Next k
    Next j
Next I
For I = 1 To mi_NOMBRE_PARAMETRE_DIGESTION
    md_CaracteristiquesDigesteur1(I) = 0
    md_CaracteristiquesDigesteur2(I) = 0
Next I
For I = 1 To mi_NOMBRE_POMPAGE
    For j = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
        md_CaracteristiquePompageBoues(I, j) = 0
    Next j
Next I
md_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C = 0
md_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C = 0
For I = 1 To mi_NOMBRE_PARAMETRE_CENTRIFUGEUSE
    md_CaracteristiqueCentrifugeuse(I) = 0
Next I
pd_CentrifugeuseConsommationPolymerePur_kgj = 0
md_DebitVapeurLyse_kgj = 0
md_TemperatureInitialeEauPourProductionVapeur_°C = 0
md_PressionAbsolueVapeurSaturanteLyse_bar = 0
md_TemperaturePrechauffageEauPourProductionVapeur_°C = 0
md_DebitEauDilution_m3j = 0
md_TemperatureEauDilution_°C = 0
md_TauxReductionGlobalMV = 0
md_RatioCH4DansBiogaz = 0
md_ProductionSpecifiqueBiogaz_Nm3_kgMV = 0
pd_BiogazProduit_Nm3j = 0
md_EnergieChauffageDigesteur1_kWhj = 0
md_BesoinTotalEnergieProductionVapeur_kWhj = 0
md_EnergieRefroidissementBouesLysees_kWhj = 0
md_EnergieDisponibleEau90°C_kWhj = 0
pd_ConsommationElectrique_kWhj = 0
pd_ConsommationElectriqueFixe_kWhj = 0

'valeurs de production de biogaz pour les co-ferments
md_PotentielProductionMethaneCoferments_Nm3_tMVAppliques = 0
md_FacteurSecuritePotentielProductionMethaneCoferments = 0

'variables booléennes permettant de repérer si la valeur a été forcée par l'utilisateur (TRUE si renseignée, FALSE sinon)
mn_TemperatureBouesEntree_°C_force = False
mn_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C_force = False
mn_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C_force = False
mn_ConcentrationBouesCentrifugeuseMelangeur_gL_force = False
mn_BouesCentrifugeuseMelangeurRatioMVSurMES_force = False
mn_TemperatureBouesNonDigerees_°C_force = False
'mn_TemperatureBouesMelangeurLyse_°C_force = False
mn_ConcentrationBouesDilutionDigesteur2_gL_force = False
For I = 1 To mi_NOMBRE_PARAMETRE_CENTRIFUGEUSE
    mn_CaracteristiqueCentrifugeuse_force(I) = False
Next I
For I = 1 To mi_NOMBRE_PARAMETRE_DIGESTION
    mn_CaracteristiquesDigesteur1_force(I) = False
    mn_CaracteristiquesDigesteur2_force(I) = False
Next I
For I = 1 To mi_NOMBRE_POMPAGE
    For j = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
        mn_CaracteristiquePompageBoues_force(I, j) = False
    Next j
Next I
mn_DebitVapeurLyse_kgj_force = False
mn_TemperatureInitialeEauPourProductionVapeur_°C_force = False
mn_PressionAbsolueVapeurSaturanteLyse_bar_force = False
mn_TemperaturePrechauffageEauPourProductionVapeur_°C_force = False
mn_TemperatureEauDilution_°C_force = False
mn_RatioCH4DansBiogaz_force = False
mn_BiogazProduit_Nm3j_force = False
mn_BesoinTotalEnergieProductionVapeur_kWhj_force = False
mn_EnergieRefroidissementBouesLysees_kWhj_force = False
mn_EnergieDisponibleEau90°C_kWhj_force = False
mn_EnergieChauffageDigesteur1_kWhj_force = False

'variables contenant les valeurs forcées par l'utilisateur
md_TemperatureBouesEntree_°C_f = 0
md_ChauffageDigesteur1TemperatureBouesEntreeEchangeur_°C_f = 0
md_ChauffageDigesteur1TemperatureBouesSortieEchangeur_°C_f = 0
md_ConcentrationBouesCentrifugeuseMelangeur_gL_f = 0
md_BouesCentrifugeuseMelangeurRatioMVSurMES_f = 0
md_TemperatureBouesNonDigerees_°C_f = 0
'md_TemperatureBouesMelangeurLyse_°C_f = 0
md_TemperatureBouesLyseEchangeur90_°C_f = 0
md_ConcentrationBouesDilutionDigesteur2_gL_f = 0
For I = 1 To mi_NOMBRE_PARAMETRE_CENTRIFUGEUSE
md_CaracteristiqueCentrifugeuse_f(I) = 0
Next I
For I = 1 To mi_NOMBRE_PARAMETRE_DIGESTION
    md_CaracteristiquesDigesteur1_f(I) = 0
    md_CaracteristiquesDigesteur2_f(I) = 0
Next I
For I = 1 To mi_NOMBRE_POMPAGE
    For j = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
        md_CaracteristiquePompageBoues_f(I, j) = 0
    Next j
Next I
md_DebitVapeurLyse_kgj_f = 0
md_TemperatureInitialeEauPourProductionVapeur_°C_f = 0
md_PressionAbsolueVapeurSaturanteLyse_bar_f = 0
md_TemperaturePrechauffageEauPourProductionVapeur_°C_f = 0
md_TemperatureEauDilution_°C_f = 0
md_RatioCH4DansBiogaz_f = 0
md_BiogazProduit_Nm3j_f = 0
md_BesoinTotalEnergieProductionVapeur_kWhj_f = 0
md_EnergieRefroidissementBouesLysees_kWhj_f = 0
md_EnergieDisponibleEau90°C_kWhj_f = 0
md_EnergieChauffageDigesteur1_kWhj_f = 0

'variables intermédiaires de calcul
md_PinchChauffageEauPourVapeur_kWhj = 0
md_PinchVaporisation_kWhj = 0
md_BiogazProduitNominal_Nm3j = 0
md_DebitPolymere_m3j = 0

'hypothèses utilisées dans les calculs
For I = 1 To nb_qualites_boues
    md_ReductionMatieresVolatilesBouesLysees(I) = 0
    md_ReductionMatieresVolatilesBouesNonLysees(I) = 0
Next I
md_ReductionMatieresVolatilesBouesDigereesLysees = 0
md_TemperatureReferenceDigesteur1_°C = 0
md_TemperatureReferenceDigesteur2_°C = 0
md_ReferenceTempsSejourBouesDigesteur1_j = 0
md_ReferenceTempsSejourBouesDigesteur2_j = 0
For I = 0 To 1
    md_CoefficientsCalculDensiteBiogaz(I) = 0
Next I
md_ProductionSpecifiqueMethane_Nm3_kgDCO = 0
md_RatioNKSolubleBouesDigereesSurMVDegrades = 0
md_RatioPSolubleBouesDigereesSurPBouesEntree = 0
md_RatioDCOSolubleSurDCOTotaleBouesDigereesDigesteur1 = 0
md_RatioDCOSolubleSurDCOTotaleBouesDigereesDigesteur2 = 0
md_RatioMVParticulairesMVTotauxBouesDigereesDigesteur1 = 0
md_PerteTemperatureBouesDigereesDigesteur1AvantLyse_°C = 0
For I = 1 To climat_type_nb
    For j = 1 To mi_DIGESTION_NON_ISOLEE
        md_PertesThermiquesDigestion_kWh_m3j(I, j) = 0
    Next j
Next I
md_EfficaciteEchangeur = 0
md_CapaciteCalorifiqueBoues_kWh_m3°C = 0
md_ConcentrationPolymereCentrifugeuse_gL = 0
md_EnthalpieTotaleVapeurSaturee_kcal_kg = 0
md_CapaciteCalorifiqueBoues_kcal_kg°C = 0
md_RatioPertesVapeur = 0
md_CapaciteCalorifiqueEau_kcal_kg°C = 0
md_MasseVolumiqueBoues_kg_m3 = 0
md_MasseVolumiqueEau_kg_m3 = 0
md_RatioNH4SurNKSolubleBouesDigerees = 0
md_RatioDBOSolubleSurDCOSolubleBouesDigerees = 0
For I = 1 To 5
    md_CoefficientsCalculTemperatureSaturationAPression(I) = 0
Next I
For I = 0 To 1
    md_CoefficientsCalculEnthalpieVaporisationATemperature(I) = 0
Next I
md_TemperatureSaturation_°C = 0
md_EnthalpieVaporisationATemperature_kcal_kg = 0
md_TemperatureBouesEchangeur90Echangeur55_°C = 0
md_TemperatureLimiteEntreMesophileEtThermophile_°C = 0
md_EcartMinimalTemperatureEchangeur_°C = 0
md_CoefficientACalculFacteurCorrectionTempsSejourBouesDigesteur = 0
md_CoefficientBCalculFacteurCorrectionTempsSejourBouesDigesteur = 0
md_ConsommationElectriqueSpecifiqueAgitationMecaniqueDigestion_W_m3 = 0
md_TempsFonctionnementAgitation_hj = 0
For I = 1 To nb_qualites_boues
    md_RatioMVSolublesMVTotauxBouesDigereesReference(I) = 0
    md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourCourt(I) = 0
    md_CoefficientsCalculCorrectionSolubilisationMVDigestionTempsSejourLong(I) = 0
Next I
md_EcartTemperatureBouesEchangeurDigesteur_°C = 0
md_ConcentrationBouesDilutionDigesteur2Souhaitee_gL = 0
md_EnergieChauffageBouesFraichesDigesteur1SansPertes_kWhj = 0
md_ConsommationElectriqueSpecifiqueAgitationMecaniqueDigestion_W_m3 = 0
md_TempsFonctionnementRecirculationBoues_hj = 0
md_PressionRefoulementRecirculationBoues_mCE = 0

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

Private Sub Class_Terminate()

    If pn_EffacementResultatsEnCours = False Then
        Call ecriture_resultats
        Call incrementation_commentaires_calculs
    End If

End Sub








