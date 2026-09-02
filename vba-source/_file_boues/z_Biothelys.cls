Attribute VB_Name = "z_Biothelys"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'Constantes pour repérer la position d'entrée des boues
Private Const mi_POSITION_ENTREE_BOUES_LYSE As Integer = 1
Private Const mi_POSITION_ENTREE_BOUES_MELANGEUR As Integer = 2

'constantes pour le choix d'agitation
Private Const mi_AGITATION_MECANIQUE As Integer = 1
Private Const mi_AGITATION_RECIRCULATION As Integer = 2
Private Const mi_AGITATION_BIOGAZ As Integer = 3

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
Private Const mi_NOMBRE_PARAMETRES_BOUES As Integer = 10
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
Private Const mi_NOMBRE_POMPAGE As Integer = 2
Private Const mi_POMPAGE_LYSE_DIGESTEUR As Integer = 1
Private Const mi_POMPAGE_BOUES_DIGEREES As Integer = 2

'paramètres des pompes
Private Const mi_NOMBRE_PARAMETRES_POMPAGE As Integer = 5
Private Const mi_PARAMETRE_POMPAGE_NOMBRE_POMPES As Integer = 1
Private Const mi_PARAMETRE_POMPAGE_HMT_mCE As Integer = 2
Private Const mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj As Integer = 3
Private Const mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h As Integer = 4
Private Const mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES As Integer = 5

'flux intermédiaires
Private Const mi_MELANGE_BOUES As Integer = nb_type_boues + 1
Private Const mi_NOMBRE_FLUX_INTERMEDIAIRES As Integer = 7
Private Const mi_FLUX_ENTREE_LYSE As Integer = 1
Private Const mi_FLUX_LYSE_ECHANGEUR55 As Integer = 2
Private Const mi_FLUX_ECHANGEUR55_MELANGEUR As Integer = 3
Private Const mi_FLUX_BOUES_NON_LYSEES As Integer = 4
Private Const mi_FLUX_MELANGEUR_DILUTION As Integer = 5
Private Const mi_FLUX_DILUTION_DIGESTEUR As Integer = 6
Private Const mi_FLUX_BOUES_DIGEREES As Integer = 7

'variables de choix
Private mi_ProvenanceBouesEntree(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Integer
Private mi_TypeBouesEntree(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Integer
Private mi_ChoixConfiguration As Integer
Private mi_ChoixEntreesBoues(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Integer
Private mi_ChoixPositionEntreeBoues(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Integer
Private mi_ChoixTypeAgitationDigesteur As Integer
Private mi_ChoixStockageBiogaz As Integer
Private mi_ChoixIsolationDigestion As Integer

'variables à échanger avec l'utilisateur
Private md_EntreeBouesPourcentageFluxAdmis(1 To mi_NOMBRE_FLUX_ENTREE_BOUES) As Double
Private md_FluxIntermediaires(1 To mi_NOMBRE_FLUX_INTERMEDIAIRES, 1 To nb_type_boues + 1, 1 To mi_NOMBRE_PARAMETRES_BOUES) As Double
Private md_CaracteristiquesDigestion(1 To mi_NOMBRE_PARAMETRE_DIGESTION) As Double
Private md_CaracteristiquePompageBoues(1 To mi_NOMBRE_POMPAGE, 1 To mi_NOMBRE_PARAMETRES_POMPAGE) As Double
Private md_DebitVapeurLyse_kgj As Double
Private md_TemperatureInitialeEauPourProductionVapeur_°C As Double
Private md_PressionAbsolueVapeurSaturanteLyse_bar As Double
Private md_TemperaturePrechauffageEauPourProductionVapeur_°C As Double
Private md_BouesLysees_kgMSj As Double
Private md_BouesLysees_m3j As Double
Private md_BouesLysees_gL As Double
Private md_DebitEauDilution_m3j As Double
Private md_TemperatureEauDilution_°C As Double
Private md_TemperatureEauDilutionApresChauffage_°C As Double
Private md_RatioCH4DansBiogaz As Double
Private md_ProductionSpecifiqueBiogaz_Nm3_kgMV As Double
Public pd_BiogazProduit_Nm3j As Double
Private md_BesoinTotalEnergieProductionVapeur_kWhj As Double
Private md_EnergieChauffageEauDilution_kWhj As Double
Private md_EnergieChauffageDigesteur_kWhj As Double
Private md_BesoinsThermiquesEnergieEau90°C_kWhj As Double
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
Private mn_TemperatureBouesEntreeLyse_°C_force As Boolean
Private mn_CaracteristiquesDigestion_force(1 To mi_NOMBRE_PARAMETRE_DIGESTION) As Boolean
Private mn_CaracteristiquePompageBoues_force(1 To mi_NOMBRE_POMPAGE, 1 To mi_NOMBRE_PARAMETRES_POMPAGE) As Boolean
Private mn_DebitVapeurLyse_kgj_force As Boolean
Private mn_TemperatureInitialeEauPourProductionVapeur_°C_force As Boolean
Private mn_PressionAbsolueVapeurSaturanteLyse_bar_force As Boolean
Private mn_TemperaturePrechauffageEauPourProductionVapeur_°C_force As Boolean
Private mn_TemperatureSortieLyse_force As Boolean
Private mn_TemperatureBouesNonLysees_°C_force As Boolean
Private mn_DebitEauDilution_m3j_force As Boolean
Private mn_TemperatureEauDilution_°C_force As Boolean
Private mn_RatioCH4DansBiogaz_force As Boolean
Private mn_BiogazProduit_Nm3j_force As Boolean
Private mn_BesoinTotalEnergieProductionVapeur_kWhj_force As Boolean
Private mn_EnergieChauffageEauDilution_kWhj_force As Boolean
Private mn_EnergieChauffageDigesteur_kWhj_force As Boolean
Private mn_BesoinsThermiquesEnergieEau90°C_kWhj_force As Boolean
Private mn_EnergieRefroidissementBouesLysees_kWhj_force As Boolean
Private mn_EnergieDisponibleEau90°C_kWhj_force As Boolean

'valeurs forcées dans le tableur
Private md_TemperatureBouesEntreeLyse_°C_f As Double
Private md_CaracteristiquesDigestion_f(1 To mi_NOMBRE_PARAMETRE_DIGESTION) As Double
Private md_CaracteristiquePompageBoues_f(1 To mi_NOMBRE_POMPAGE, 1 To mi_NOMBRE_PARAMETRES_POMPAGE) As Double
Private md_DebitVapeurLyse_kgj_f As Double
Private md_TemperatureInitialeEauPourProductionVapeur_°C_f As Double
Private md_PressionAbsolueVapeurSaturanteLyse_bar_f As Double
Private md_TemperaturePrechauffageEauPourProductionVapeur_°C_f As Double
Private md_TemperatureSortieLyse_f As Double
Private md_TemperatureBouesNonLysees_°C_f As Double
Private md_DebitEauDilution_m3j_f As Double
Private md_TemperatureEauDilution_°C_f As Double
Private md_RatioCH4DansBiogaz_f As Double
Private md_BiogazProduit_Nm3j_f As Double
Private md_BesoinTotalEnergieProductionVapeur_kWhj_f As Double
Private md_EnergieChauffageEauDilution_kWhj_f As Double
Private md_EnergieChauffageDigesteur_kWhj_f As Double
Private md_BesoinsThermiquesEnergieEau90°C_kWhj_f As Double
Private md_EnergieRefroidissementBouesLysees_kWhj_f As Double
Private md_EnergieDisponibleEau90°C_kWhj_f As Double

'hypothèses
Private md_ReductionMatieresVolatilesBouesLysees(1 To nb_qualites_boues) As Double
Private md_ReductionMatieresVolatilesBouesNonLysees(1 To nb_qualites_boues) As Double
Private md_TemperatureDigestionReference_°C As Double
Private md_TemperatureDigestionMinimale_°C As Double
Private md_TemperatureDigestionMaximale_°C As Double
Private md_TemperatureEauDilutionMaximale_°C As Double
Private md_ReferenceTempsSejourBouesDigesteur_j As Double
Private md_CoefficientsCalculDensiteBiogaz(0 To 1) As Double
Private md_ProductionSpecifiqueMethane_Nm3_kgDCO As Double
Private md_RatioNKSolubleBouesDigereesSurMVDegrades As Double
Private md_RatioPSolubleBouesDigereesSurPBouesEntree As Double
Private md_RatioDCOSolubleSurDCOTotaleBouesDigerees As Double
Private md_PertesThermiquesDigestion_kWh_m3j(1 To climat_type_nb, 1 To mi_DIGESTION_NON_ISOLEE) As Double
Private md_EfficaciteEchangeur As Double
Private md_CapaciteCalorifiqueBoues_kWh_m3°C As Double
Private md_EnthalpieTotaleVapeurSaturee_kcal_kg As Double
Private md_CapaciteCalorifiqueBoues_kcal_kg°C As Double
Private md_RatioPertesVapeur As Double
Private md_CapaciteCalorifiqueEau_kcal_kg°C As Double
Private md_MasseVolumiqueBoues_kg_m3 As Double
Private md_MasseVolumiqueEau_kg_m3 As Double
Private md_CoefficientsCalculTemperatureSaturationAPression(1 To 5) As Double
Private md_CoefficientsCalculEnthalpieVaporisationATemperature(0 To 1) As Double
Private md_TemperatureSaturation_°C As Double
Private md_EnthalpieVaporisationATemperature_kcal_kg As Double
Private md_TemperatureLyse_°C As Double
Private md_CoefficientAConsommationSpecifiqueVapeur_kWh_m3Boues°C As Double
Private md_CoefficientBConsommationSpecifiqueVapeur_kWh_m3Boues As Double
Private md_ConcentrationBouesEntreeDigestionSouhaitee_gL As Double
Private md_CoefficientACalculFacteurCorrectionTempsSejourBouesDigesteur As Double
Private md_CoefficientBCalculFacteurCorrectionTempsSejourBouesDigesteur As Double
Private md_TemperatureMaximaleBouesEntreeBiothelysValiditeCalculs_°C As Double
Private md_ConsommationElectriqueSpecifiqueAgitationMecaniqueDigestion_W_m3 As Double
Private md_TempsFonctionnementAgitation_hj As Double
Private md_TempsFonctionnementRecirculationBoues_hj As Double
Private md_PressionRefoulementRecirculationBoues_mCE As Double
Private md_EcartTemperatureBouesEchangeurDigesteur_°C As Double

'bilan thermique
Private md_PinchChauffageEauPourVapeur_kWhj As Double
Private md_PinchVaporisation_kWhj As Double

'stockage de données
Private md_BiogazProduitNominal_Nm3j As Double

'erreurs
Private Const mi_NOMBRE_COMMENTAIRES_CALCULS As Integer = 3
Private ms_CommentairesCalculs() As String
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
Dim k As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "44.0"
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
mi_ChoixConfiguration = 0
mi_ChoixTypeAgitationDigesteur = 0
mi_ChoixStockageBiogaz = 0
mi_ChoixIsolationDigestion = 0

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
    md_CaracteristiquesDigestion(I) = 0
Next I
For I = 1 To mi_NOMBRE_POMPAGE
    For j = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
        md_CaracteristiquePompageBoues(I, j) = 0
    Next j
Next I
md_DebitVapeurLyse_kgj = 0
md_TemperatureInitialeEauPourProductionVapeur_°C = 0
md_PressionAbsolueVapeurSaturanteLyse_bar = 0
md_TemperaturePrechauffageEauPourProductionVapeur_°C = 0
md_BouesLysees_kgMSj = 0
md_BouesLysees_m3j = 0
md_BouesLysees_gL = 0
md_DebitEauDilution_m3j = 0
md_TemperatureEauDilution_°C = 0
md_TemperatureEauDilutionApresChauffage_°C = 0
md_RatioCH4DansBiogaz = 0
md_ProductionSpecifiqueBiogaz_Nm3_kgMV = 0
pd_BiogazProduit_Nm3j = 0
md_BesoinTotalEnergieProductionVapeur_kWhj = 0
md_BesoinsThermiquesEnergieEau90°C_kWhj = 0
md_EnergieRefroidissementBouesLysees_kWhj = 0
md_EnergieDisponibleEau90°C_kWhj = 0
pd_ConsommationElectrique_kWhj = 0
pd_ConsommationElectriqueFixe_kWhj = 0

'valeurs de production de biogaz pour les co-ferments
md_PotentielProductionMethaneCoferments_Nm3_tMVAppliques = 0
md_FacteurSecuritePotentielProductionMethaneCoferments = 0

'variables booléennes permettant de repérer si la valeur a été forcée par l'utilisateur (TRUE si renseignée, FALSE sinon)
mn_TemperatureBouesEntreeLyse_°C_force = False
For I = 1 To mi_NOMBRE_PARAMETRE_DIGESTION
    mn_CaracteristiquesDigestion_force(I) = False
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
mn_TemperatureSortieLyse_force = False
mn_TemperatureBouesNonLysees_°C_force = False
mn_DebitEauDilution_m3j_force = False
mn_TemperatureEauDilution_°C_force = False
mn_RatioCH4DansBiogaz_force = False
mn_BiogazProduit_Nm3j_force = False
mn_BesoinTotalEnergieProductionVapeur_kWhj_force = False
mn_BesoinsThermiquesEnergieEau90°C_kWhj_force = False
mn_EnergieRefroidissementBouesLysees_kWhj_force = False
mn_EnergieDisponibleEau90°C_kWhj_force = False

'variables contenant les valeurs forcées par l'utilisateur
md_TemperatureBouesEntreeLyse_°C_f = 0
For I = 1 To mi_NOMBRE_PARAMETRE_DIGESTION
    md_CaracteristiquesDigestion_f(I) = 0
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
md_TemperatureSortieLyse_f = 0
md_TemperatureBouesNonLysees_°C_f = 0
md_DebitEauDilution_m3j_f = 0
md_TemperatureEauDilution_°C_f = 0
md_RatioCH4DansBiogaz_f = 0
md_BiogazProduit_Nm3j_f = 0
md_BesoinTotalEnergieProductionVapeur_kWhj_f = 0
md_BesoinsThermiquesEnergieEau90°C_kWhj_f = 0
md_EnergieRefroidissementBouesLysees_kWhj_f = 0
md_EnergieDisponibleEau90°C_kWhj_f = 0

'variables intermédiaires de calcul
md_PinchChauffageEauPourVapeur_kWhj = 0
md_PinchVaporisation_kWhj = 0
md_BiogazProduitNominal_Nm3j = 0

'hypothèses utilisées dans les calculs
For I = 1 To nb_qualites_boues
    md_ReductionMatieresVolatilesBouesLysees(I) = 0
    md_ReductionMatieresVolatilesBouesNonLysees(I) = 0
Next I
md_TemperatureDigestionReference_°C = 0
md_TemperatureDigestionMinimale_°C = 0
md_TemperatureDigestionMaximale_°C = 0
md_TemperatureEauDilutionMaximale_°C = 0
md_ReferenceTempsSejourBouesDigesteur_j = 0
For I = 0 To 1
    md_CoefficientsCalculDensiteBiogaz(I) = 0
Next I
md_ProductionSpecifiqueMethane_Nm3_kgDCO = 0
md_RatioNKSolubleBouesDigereesSurMVDegrades = 0
md_RatioPSolubleBouesDigereesSurPBouesEntree = 0
md_RatioDCOSolubleSurDCOTotaleBouesDigerees = 0
For I = 1 To climat_type_nb
    For j = 1 To mi_DIGESTION_NON_ISOLEE
        md_PertesThermiquesDigestion_kWh_m3j(I, j) = 0
    Next j
Next I
md_EfficaciteEchangeur = 0
md_CapaciteCalorifiqueBoues_kWh_m3°C = 0
md_EnthalpieTotaleVapeurSaturee_kcal_kg = 0
md_CapaciteCalorifiqueBoues_kcal_kg°C = 0
md_RatioPertesVapeur = 0
md_CapaciteCalorifiqueEau_kcal_kg°C = 0
md_MasseVolumiqueBoues_kg_m3 = 0
md_MasseVolumiqueEau_kg_m3 = 0
For I = 1 To 5
    md_CoefficientsCalculTemperatureSaturationAPression(I) = 0
Next I
For I = 0 To 1
    md_CoefficientsCalculEnthalpieVaporisationATemperature(I) = 0
Next I
md_TemperatureSaturation_°C = 0
md_EnthalpieVaporisationATemperature_kcal_kg = 0
md_TemperatureLyse_°C = 0
md_CoefficientAConsommationSpecifiqueVapeur_kWh_m3Boues°C = 0
md_CoefficientBConsommationSpecifiqueVapeur_kWh_m3Boues = 0
md_ConcentrationBouesEntreeDigestionSouhaitee_gL = 0
md_CoefficientACalculFacteurCorrectionTempsSejourBouesDigesteur = 0
md_CoefficientBCalculFacteurCorrectionTempsSejourBouesDigesteur = 0
md_TemperatureMaximaleBouesEntreeBiothelysValiditeCalculs_°C = 0
md_ConsommationElectriqueSpecifiqueAgitationMecaniqueDigestion_W_m3 = 0
md_TempsFonctionnementAgitation_hj = 0
md_TempsFonctionnementRecirculationBoues_hj = 0
md_PressionRefoulementRecirculationBoues_mCE = 0
md_EcartTemperatureBouesEchangeurDigesteur_°C = 0

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


Public Sub dimensionnement(ByRef TableauRecapitulatifFluxBoues() As Double, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double, ByRef PartGraissesInternesSurTotalCosubstrat() As Double, ByVal ChoixEauxSalesPrimaire As Boolean, ByRef DetailBouesI() As Double, ByRef DetailBouesII() As Double, ByRef DetailBouesIII() As Double, ByRef FluxDCODansBiogaz_kgj As Double, ByRef TableauBesoinsThermiques() As Double, ByRef TableauDisponibilitesThermiques() As Double, ByRef CaracteristiquesBiogaz() As Double, ByRef PressionVapeur_bar() As Double, ByRef NombreReseauxVapeur As Integer, ByRef LyseAvalEpaississement() As Integer)


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

'Initialisation des variables locales
NumeroErreurNonGeree = "44.5"
I = 0
j = 0
k = 0
ReperageBoucleReel = False
IncoherenceCalcul = 0



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

'CALCUL DES BOUES EN ENTREE (VERS LYSE ET BYPASS)
Call calcul_cocktail_boues_entrees(TableauRecapitulatifFluxBoues, TableauRecapitulatifPollutionSolubleDansFluxBoues, ChoixEauxSalesPrimaire, DetailBouesI, DetailBouesII, DetailBouesIII, PartGraissesInternesSurTotalCosubstrat, LyseAvalEpaississement)

'TEMPERATURE DES BOUES EN ENTREE DE L'HYDROLYSE
If mn_TemperatureBouesEntreeLyse_°C_force = True Then
    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureBouesEntreeLyse_°C_f
    If md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) > md_TemperatureMaximaleBouesEntreeBiothelysValiditeCalculs_°C And mi_ChoixConfiguration = pi_CONFIGURATION_BIOTHELYS Then
        MsgBox Feuil44.Range("erreur_temperature_boues_entree lyse") & " " & md_TemperatureMaximaleBouesEntreeBiothelysValiditeCalculs_°C & " °C", Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If
Else
    If mi_ChoixConfiguration = pi_CONFIGURATION_BIOTHELYS Then
        If pd_TemperatureEauDimensionnement_°C > md_TemperatureMaximaleBouesEntreeBiothelysValiditeCalculs_°C Then
            md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureMaximaleBouesEntreeBiothelysValiditeCalculs_°C
        Else
            md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = pd_TemperatureEauDimensionnement_°C
        End If
    Else
        md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = pd_TemperatureEauDimensionnement_°C
    End If
End If
If mn_TemperatureBouesNonLysees_°C_force = True Then
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureBouesNonLysees_°C_f
Else
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = pd_TemperatureEauDimensionnement_°C
End If

'Température initiale de l'eau de dilution
If mn_TemperatureEauDilution_°C_force = True Then
    md_TemperatureEauDilution_°C = md_TemperatureEauDilution_°C_f
Else
    md_TemperatureEauDilution_°C = pd_TemperatureEauDimensionnement_°C
End If


'HYDROLYSE THERMIQUE DES BOUES
Call calcul_bilan_masse_lyse(ReperageBoucleReel)

'MELANGE AVEC LES BOUES NON HYDROLYSEES
Call calcul_bilan_masse_melangeur(mi_FLUX_ECHANGEUR55_MELANGEUR, mi_FLUX_BOUES_NON_LYSEES, mi_FLUX_MELANGEUR_DILUTION)
    
'DILUTION DES BOUES AVANT DIGESTION
Call calcul_dilution_boues_avant_digestion(ReperageBoucleReel)

'DIGESTEUR
Call calcul_bilan_masse_digestion(md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j))

'calcul du volume du digesteur
md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j)
If mn_CaracteristiquesDigestion_force(mi_PARAMETRE_DIGESTION_VOLUME_m3) = True Then
    IncoherenceCalcul = (md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3) - md_CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_VOLUME_m3)) / md_CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_VOLUME_m3)
    If IncoherenceCalcul > mi_CRITERE_INCOHERENCE Then
        ms_CommentairesCalculs(1) = Feuil44.Range("dig_biothelys_erreur_volume").Cells(1, 1)
    End If
    md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3) = md_CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_VOLUME_m3)
End If

'On calcule le biogaz
Call calcul_biogaz(FluxDCODansBiogaz_kgj)
md_BiogazProduitNominal_Nm3j = pd_BiogazProduit_Nm3j
CaracteristiquesBiogaz(Q_Nm3j) = pd_BiogazProduit_Nm3j
CaracteristiquesBiogaz(teneur_CH4) = md_RatioCH4DansBiogaz

'calcul de la pollution solubilisée
Call calcul_solubilisation_pendant_digestion(TableauRecapitulatifPollutionSolubleDansFluxBoues)

'GESTION ENERGIE
Call calcul_bilan_thermique(TableauBesoinsThermiques, TableauDisponibilitesThermiques, PressionVapeur_bar, NombreReseauxVapeur, ReperageBoucleReel)

'RECAPITULATIF SORTIE
For I = 1 To nb_type_boues
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_MS_kgj)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_Q) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_Q_m3j)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_MV_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_MV_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_DCO_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_DBO_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_NK_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_NK_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_Pt_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_P_MES)
Next I



Exit Sub
'*********************************************************************
erreur_non_geree:
NumeroErreurNonGeree = "44.5"

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
Dim BouesEntreeLyseMV_kgj As Double
Dim BouesEntreeLyseMES_kgj As Double
Dim BouesEntreeLyseMS_kgj As Double
Dim BouesEntreeLyseNK_kgj As Double
Dim BouesEntreeLyseP_kgj As Double
Dim BouesEntreeLyseDCO_kgj As Double
Dim BouesEntreeLyseDBO_kgj As Double
Dim BouesEntreeLyseDebitVolumique_m3j As Double
Dim BouesEntreeLyseReductionMV_kgj As Double
Dim BouesNonLyseesMV_kgj As Double
Dim BouesNonLyseesMES_kgj As Double
Dim BouesNonLyseesMS_kgj As Double
Dim BouesNonLyseesNK_kgj As Double
Dim BouesNonLyseesP_kgj As Double
Dim BouesNonLyseesDCO_kgj As Double
Dim BouesNonLyseesDBO_kgj As Double
Dim BouesNonLyseesDebitVolumique_m3j As Double
Dim BouesNonLyseesReductionMV_kgj As Double
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
NumeroErreurNonGeree = "44.26"
I = 0
j = 0
k = 0
BouesEntreeLyseMV_kgj = 0
BouesEntreeLyseMES_kgj = 0
BouesEntreeLyseMS_kgj = 0
BouesEntreeLyseNK_kgj = 0
BouesEntreeLyseP_kgj = 0
BouesEntreeLyseDCO_kgj = 0
BouesEntreeLyseDBO_kgj = 0
BouesEntreeLyseDebitVolumique_m3j = 0
BouesEntreeLyseReductionMV_kgj = 0
BouesNonLyseesMV_kgj = 0
BouesNonLyseesMES_kgj = 0
BouesNonLyseesMS_kgj = 0
BouesNonLyseesNK_kgj = 0
BouesNonLyseesP_kgj = 0
BouesNonLyseesDCO_kgj = 0
BouesNonLyseesDBO_kgj = 0
BouesNonLyseesDebitVolumique_m3j = 0
BouesNonLyseesReductionMV_kgj = 0
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
                
                If mi_ChoixPositionEntreeBoues(I) = mi_POSITION_ENTREE_BOUES_LYSE Then
                    'On calcule la réduction potentielle des MV
                    If mi_TypeBouesEntree(I) = boues_I And ChoixEauxSalesPrimaire = True Then
                        PartBouesBiologiques = 0
                        CoefficientReductionMV = 0
                        For k = 1 To nb_eaux_sales_max
                            PartBouesBiologiques = PartBouesBiologiques + DetailBouesI(k, repere_ratio_MV)
                            CoefficientReductionMV = CoefficientReductionMV + DetailBouesI(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesLysees(DetailBouesI(k, repere_origine))
                        Next k
                        CoefficientReductionMV = CoefficientReductionMV + (1 - PartBouesBiologiques) * md_ReductionMatieresVolatilesBouesLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine))
                        CoefficientReductionMV = CoefficientReductionMV
                        IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                    ElseIf mi_TypeBouesEntree(I) = boues_II And choix_biostyr = True Then
                        CoefficientReductionMV = 0
                        For k = 1 To nb_eaux_sales_max
                            CoefficientReductionMV = CoefficientReductionMV + DetailBouesII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesLysees(DetailBouesII(k, repere_origine))
                        Next k
                        CoefficientReductionMV = CoefficientReductionMV
                        IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                    ElseIf mi_TypeBouesEntree(I) = boues_III And choix_biostyr_N_III = True Then
                        CoefficientReductionMV = 0
                        For k = 1 To nb_eaux_sales_max
                            CoefficientReductionMV = CoefficientReductionMV + DetailBouesIII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesLysees(DetailBouesIII(k, repere_origine))
                        Next k
                        CoefficientReductionMV = CoefficientReductionMV
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
                        'BouesEntreeLyseMS_kgj = BouesEntreeLyseMS_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), mi_TypeBouesEntree(i), repere_mgL_MS_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireMV_kgj = IntermediaireMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        'BouesEntreeLyseMV_kgj = BouesEntreeLyseMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), mi_TypeBouesEntree(i), repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireDCO_kgj = IntermediaireDCO_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_DCO) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireNK_kgj = IntermediaireNK_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_NK) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireP_kgj = IntermediaireP_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_Pt) * IntermediaireDebitVolumique_m3j / 1000
                    Else
                        IntermediaireReductionMV_kgj = md_ReductionMatieresVolatilesBouesLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine)) * IntermediaireMV_kgj
                    End If
                    'on calcul l'entrée globale de la lyse
                    BouesEntreeLyseMS_kgj = BouesEntreeLyseMS_kgj + IntermediaireMS_kgj
                    BouesEntreeLyseMES_kgj = BouesEntreeLyseMES_kgj + IntermediaireMES_kgj
                    BouesEntreeLyseMV_kgj = BouesEntreeLyseMV_kgj + IntermediaireMV_kgj
                    BouesEntreeLyseReductionMV_kgj = BouesEntreeLyseReductionMV_kgj + IntermediaireReductionMV_kgj
                    BouesEntreeLyseDCO_kgj = BouesEntreeLyseDCO_kgj + IntermediaireDCO_kgj
                    BouesEntreeLyseDBO_kgj = BouesEntreeLyseDBO_kgj + IntermediaireDBO_kgj
                    BouesEntreeLyseNK_kgj = BouesEntreeLyseNK_kgj + IntermediaireNK_kgj
                    BouesEntreeLyseP_kgj = BouesEntreeLyseP_kgj + IntermediaireP_kgj
                    BouesEntreeLyseDebitVolumique_m3j = BouesEntreeLyseDebitVolumique_m3j + IntermediaireDebitVolumique_m3j
                    'On remplit aussi le tableau avec le détail
                    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_MV_MES) = IntermediaireMV_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_MV_MES)
                    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_REDUCTION_MV) = IntermediaireReductionMV_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_REDUCTION_MV)
                    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_MS_kgj) + IntermediaireMS_kgj
                    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DCO_MES) = IntermediaireDCO_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DCO_MES)
                    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DBO_MES) = IntermediaireDBO_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DBO_MES)
                    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_NK_MES) = IntermediaireNK_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_NK_MES)
                    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_P_MES) = IntermediaireP_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_P_MES)
                    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_Q_m3j) = IntermediaireDebitVolumique_m3j + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_Q_m3j)
                ElseIf mi_ChoixPositionEntreeBoues(I) = mi_POSITION_ENTREE_BOUES_MELANGEUR Then
                    'On calcule la réduction potentielle des MV
                    If mi_TypeBouesEntree(I) = boues_I And ChoixEauxSalesPrimaire = True Then
                        PartBouesBiologiques = 0
                        CoefficientReductionMV = 0
                        For k = 1 To nb_eaux_sales_max
                            PartBouesBiologiques = PartBouesBiologiques + DetailBouesI(k, repere_ratio_MV)
                            CoefficientReductionMV = CoefficientReductionMV + DetailBouesI(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesNonLysees(DetailBouesI(k, repere_origine))
                        Next k
                        CoefficientReductionMV = CoefficientReductionMV + (1 - PartBouesBiologiques) * md_ReductionMatieresVolatilesBouesNonLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine))
                        CoefficientReductionMV = CoefficientReductionMV
                        IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                    ElseIf mi_TypeBouesEntree(I) = boues_II And choix_biostyr = True Then
                        CoefficientReductionMV = 0
                        For k = 1 To nb_eaux_sales_max
                            CoefficientReductionMV = CoefficientReductionMV + DetailBouesII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesNonLysees(DetailBouesII(k, repere_origine))
                        Next k
                        CoefficientReductionMV = CoefficientReductionMV
                        IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                    ElseIf mi_TypeBouesEntree(I) = boues_III And choix_biostyr_N_III = True Then
                        CoefficientReductionMV = 0
                        For k = 1 To nb_eaux_sales_max
                            CoefficientReductionMV = CoefficientReductionMV + DetailBouesIII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesNonLysees(DetailBouesIII(k, repere_origine))
                        Next k
                        CoefficientReductionMV = CoefficientReductionMV
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
                        'BouesNonLyseesMS_kgj = BouesNonLyseesMS_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), mi_TypeBouesEntree(i), repere_mgL_MS_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireMV_kgj = IntermediaireMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        'BouesNonLyseesMV_kgj = BouesNonLyseesMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), mi_TypeBouesEntree(i), repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireDCO_kgj = IntermediaireDCO_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_DCO) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireNK_kgj = IntermediaireNK_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_NK) * IntermediaireDebitVolumique_m3j / 1000
                        IntermediaireP_kgj = IntermediaireP_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_mgL_Pt) * IntermediaireDebitVolumique_m3j / 1000
                    Else
                        IntermediaireReductionMV_kgj = md_ReductionMatieresVolatilesBouesNonLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_origine)) * IntermediaireMV_kgj
                    End If
                    'on calcul l'entrée des boues non lysées qui vont vers le mélangeur
                    BouesNonLyseesMES_kgj = BouesNonLyseesMES_kgj + IntermediaireMES_kgj
                    BouesNonLyseesMS_kgj = BouesNonLyseesMS_kgj + IntermediaireMS_kgj
                    BouesNonLyseesMV_kgj = BouesNonLyseesMV_kgj + IntermediaireMV_kgj
                    BouesNonLyseesReductionMV_kgj = BouesNonLyseesReductionMV_kgj + IntermediaireReductionMV_kgj
                    BouesNonLyseesDCO_kgj = BouesNonLyseesDCO_kgj + IntermediaireDCO_kgj
                    BouesNonLyseesDBO_kgj = BouesNonLyseesDBO_kgj + IntermediaireDBO_kgj
                    BouesNonLyseesNK_kgj = BouesNonLyseesNK_kgj + IntermediaireNK_kgj
                    BouesNonLyseesP_kgj = BouesNonLyseesP_kgj + IntermediaireP_kgj
                    BouesNonLyseesDebitVolumique_m3j = BouesNonLyseesDebitVolumique_m3j + IntermediaireDebitVolumique_m3j
                    'On remplit aussi le tableau avec le détail
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_MV_MES) = IntermediaireMV_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_MV_MES)
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_REDUCTION_MV) = IntermediaireReductionMV_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_REDUCTION_MV)
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_MS_kgj) + IntermediaireMS_kgj
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DCO_MES) = IntermediaireDCO_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DCO_MES)
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DBO_MES) = IntermediaireDBO_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_DBO_MES)
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_NK_MES) = IntermediaireNK_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_NK_MES)
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_P_MES) = IntermediaireP_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_RATIO_P_MES)
                    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_Q_m3j) = IntermediaireDebitVolumique_m3j + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_TypeBouesEntree(I), mi_PARAMETRE_BOUES_Q_m3j)
                End If
                'pour la vérification des flux
                TableauRecapitulatifFluxBoues(boues_digerees, mi_TypeBouesEntree(I), repere_flux_in) = TableauRecapitulatifFluxBoues(boues_digerees, mi_TypeBouesEntree(I), repere_flux_in) + TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_flux_in) * md_EntreeBouesPourcentageFluxAdmis(I)
                TableauRecapitulatifFluxBoues(boues_digerees, mi_TypeBouesEntree(I), repere_verif_flux) = TableauRecapitulatifFluxBoues(boues_digerees, mi_TypeBouesEntree(I), repere_flux_in)
                'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
                TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_verif_flux) = TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_verif_flux) - TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), mi_TypeBouesEntree(I), repere_flux_in) * md_EntreeBouesPourcentageFluxAdmis(I)
            End If
            
        Else
            If mi_ChoixPositionEntreeBoues(I) = mi_POSITION_ENTREE_BOUES_LYSE Then
                LyseAvalEpaississement(mi_ProvenanceBouesEntree(I) - boues_inlet) = mi_ChoixConfiguration
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
                    
                    If mi_ChoixPositionEntreeBoues(I) = mi_POSITION_ENTREE_BOUES_LYSE Then
                        'On calcule la réduction potentielle des MV
                        If j = boues_I And ChoixEauxSalesPrimaire = True Then
                            PartBouesBiologiques = 0
                            CoefficientReductionMV = 0
                            For k = 1 To nb_eaux_sales_max
                                PartBouesBiologiques = PartBouesBiologiques + DetailBouesI(k, repere_ratio_MV)
                                CoefficientReductionMV = CoefficientReductionMV + DetailBouesI(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesLysees(DetailBouesI(k, repere_origine))
                            Next k
                            CoefficientReductionMV = CoefficientReductionMV + (1 - PartBouesBiologiques) * md_ReductionMatieresVolatilesBouesLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine))
                            CoefficientReductionMV = CoefficientReductionMV
                            IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                        ElseIf j = boues_II And choix_biostyr = True Then
                            CoefficientReductionMV = 0
                            For k = 1 To nb_eaux_sales_max
                                CoefficientReductionMV = CoefficientReductionMV + DetailBouesII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesLysees(DetailBouesII(k, repere_origine))
                            Next k
                            CoefficientReductionMV = CoefficientReductionMV
                            IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                        ElseIf j = boues_III And choix_biostyr_N_III = True Then
                            CoefficientReductionMV = 0
                            For k = 1 To nb_eaux_sales_max
                                CoefficientReductionMV = CoefficientReductionMV + DetailBouesIII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesLysees(DetailBouesIII(k, repere_origine))
                            Next k
                            CoefficientReductionMV = CoefficientReductionMV
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
                            'BouesEntreeLyseMS_kgj = BouesEntreeLyseMS_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), j, repere_mgL_MS_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireMV_kgj = IntermediaireMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            'BouesEntreeLyseMV_kgj = BouesEntreeLyseMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), j, repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireDCO_kgj = IntermediaireDCO_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_DCO) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireNK_kgj = IntermediaireNK_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_NK) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireP_kgj = IntermediaireP_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_Pt) * IntermediaireDebitVolumique_m3j / 1000
                        Else
                            IntermediaireReductionMV_kgj = md_ReductionMatieresVolatilesBouesLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine)) * IntermediaireMV_kgj
                        End If
                        'on calcul l'entrée globale de la lyse
                        BouesEntreeLyseMS_kgj = BouesEntreeLyseMS_kgj + IntermediaireMS_kgj
                        BouesEntreeLyseMES_kgj = BouesEntreeLyseMES_kgj + IntermediaireMES_kgj
                        BouesEntreeLyseMV_kgj = BouesEntreeLyseMV_kgj + IntermediaireMV_kgj
                        BouesEntreeLyseReductionMV_kgj = BouesEntreeLyseReductionMV_kgj + IntermediaireReductionMV_kgj
                        BouesEntreeLyseDCO_kgj = BouesEntreeLyseDCO_kgj + IntermediaireDCO_kgj
                        BouesEntreeLyseDBO_kgj = BouesEntreeLyseDBO_kgj + IntermediaireDBO_kgj
                        BouesEntreeLyseNK_kgj = BouesEntreeLyseNK_kgj + IntermediaireNK_kgj
                        BouesEntreeLyseP_kgj = BouesEntreeLyseP_kgj + IntermediaireP_kgj
                        BouesEntreeLyseDebitVolumique_m3j = BouesEntreeLyseDebitVolumique_m3j + IntermediaireDebitVolumique_m3j
                        'On remplit aussi le tableau avec le détail
                        md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = IntermediaireMV_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = IntermediaireReductionMV_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_REDUCTION_MV)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_MS_kgj) + IntermediaireMS_kgj
                        md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = IntermediaireDCO_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = IntermediaireDBO_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = IntermediaireNK_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_NK_MES)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = IntermediaireP_kgj + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_P_MES)
                        md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_Q_m3j) = IntermediaireDebitVolumique_m3j + md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_Q_m3j)
                    ElseIf mi_ChoixPositionEntreeBoues(I) = mi_POSITION_ENTREE_BOUES_MELANGEUR Then
                        'On calcule la réduction potentielle des MV
                        If j = boues_I And ChoixEauxSalesPrimaire = True Then
                            PartBouesBiologiques = 0
                            CoefficientReductionMV = 0
                            For k = 1 To nb_eaux_sales_max
                                PartBouesBiologiques = PartBouesBiologiques + DetailBouesI(k, repere_ratio_MV)
                                CoefficientReductionMV = CoefficientReductionMV + DetailBouesI(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesNonLysees(DetailBouesI(k, repere_origine))
                            Next k
                            CoefficientReductionMV = CoefficientReductionMV + (1 - PartBouesBiologiques) * md_ReductionMatieresVolatilesBouesNonLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine))
                            CoefficientReductionMV = CoefficientReductionMV
                            IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                        ElseIf j = boues_II And choix_biostyr = True Then
                            CoefficientReductionMV = 0
                            For k = 1 To nb_eaux_sales_max
                                CoefficientReductionMV = CoefficientReductionMV + DetailBouesII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesNonLysees(DetailBouesII(k, repere_origine))
                            Next k
                            CoefficientReductionMV = CoefficientReductionMV
                            IntermediaireReductionMV_kgj = IntermediaireMV_kgj * CoefficientReductionMV
                        ElseIf j = boues_III And choix_biostyr_N_III = True Then
                            CoefficientReductionMV = 0
                            For k = 1 To nb_eaux_sales_max
                                CoefficientReductionMV = CoefficientReductionMV + DetailBouesIII(k, repere_ratio_MV) * md_ReductionMatieresVolatilesBouesNonLysees(DetailBouesIII(k, repere_origine))
                            Next k
                            CoefficientReductionMV = CoefficientReductionMV
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
                            'BouesNonLyseesMS_kgj = BouesNonLyseesMS_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), j, repere_mgL_MS_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireMV_kgj = IntermediaireMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            'BouesNonLyseesMV_kgj = BouesNonLyseesMV_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(i), j, repere_mgL_MV_soluble) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireDCO_kgj = IntermediaireDCO_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_DCO) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireNK_kgj = IntermediaireNK_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_NK) * IntermediaireDebitVolumique_m3j / 1000
                            IntermediaireP_kgj = IntermediaireP_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_mgL_Pt) * IntermediaireDebitVolumique_m3j / 1000
                        Else
                            IntermediaireReductionMV_kgj = md_ReductionMatieresVolatilesBouesNonLysees(TableauRecapitulatifFluxBoues(mi_ProvenanceBouesEntree(I), j, repere_origine)) * IntermediaireMV_kgj
                        End If
                        'on calcul l'entrée des boues non lysées qui vont vers le mélangeur
                        BouesNonLyseesMS_kgj = BouesNonLyseesMS_kgj + IntermediaireMS_kgj
                        BouesNonLyseesMES_kgj = BouesNonLyseesMES_kgj + IntermediaireMES_kgj
                        BouesNonLyseesMV_kgj = BouesNonLyseesMV_kgj + IntermediaireMV_kgj
                        BouesNonLyseesReductionMV_kgj = BouesNonLyseesReductionMV_kgj + IntermediaireReductionMV_kgj
                        BouesNonLyseesDCO_kgj = BouesNonLyseesDCO_kgj + IntermediaireDCO_kgj
                        BouesNonLyseesDBO_kgj = BouesNonLyseesDBO_kgj + IntermediaireDBO_kgj
                        BouesNonLyseesNK_kgj = BouesNonLyseesNK_kgj + IntermediaireNK_kgj
                        BouesNonLyseesP_kgj = BouesNonLyseesP_kgj + IntermediaireP_kgj
                        BouesNonLyseesDebitVolumique_m3j = BouesNonLyseesDebitVolumique_m3j + IntermediaireDebitVolumique_m3j
                        'On remplit aussi le tableau avec le détail
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = IntermediaireMV_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = IntermediaireReductionMV_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_REDUCTION_MV)
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_MS_kgj) + IntermediaireMS_kgj
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = IntermediaireDCO_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = IntermediaireDBO_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = IntermediaireNK_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_NK_MES)
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = IntermediaireP_kgj + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_P_MES)
                        md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_Q_m3j) = IntermediaireDebitVolumique_m3j + md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_Q_m3j)
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

If BouesEntreeLyseMES_kgj = 0 Or BouesEntreeLyseDebitVolumique_m3j = 0 Then
    MsgBox Feuil5.Range("erreur_choix_biothelys_inlet"), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
    pn_EffacementResultatsEnCours = True
    pn_ErreurNonGeree = True
    Exit Sub
End If


'On complète la ligne sur la somme des boues puis on recalcule tous les ratios
'pour les boues qui vont être lysées
    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) = BouesEntreeLyseMS_kgj
    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) = BouesEntreeLyseMV_kgj
    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = BouesEntreeLyseDCO_kgj
    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = BouesEntreeLyseDBO_kgj
    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_NK_MES) = BouesEntreeLyseNK_kgj
    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_P_MES) = BouesEntreeLyseP_kgj
    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) = BouesEntreeLyseDebitVolumique_m3j
    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) = BouesEntreeLyseReductionMV_kgj
    For j = 1 To nb_type_boues + 1
        If md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_MS_kgj) > 0 Then
            md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_REDUCTION_MV) / md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
            md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) / md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) / md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) / md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) / md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_P_MES) / md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_Q_m3j)
        End If
    Next j

'pour les boues qui vont être digérées sans être lysées
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) = BouesNonLyseesMS_kgj
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) = BouesNonLyseesMV_kgj
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = BouesNonLyseesDCO_kgj
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = BouesNonLyseesDBO_kgj
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_NK_MES) = BouesNonLyseesNK_kgj
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_P_MES) = BouesNonLyseesP_kgj
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) = BouesNonLyseesDebitVolumique_m3j
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) = BouesNonLyseesReductionMV_kgj
    For j = 1 To nb_type_boues + 1
        If md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_MS_kgj) > 0 Then
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_REDUCTION_MV) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_RATIO_P_MES) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_MS_kgj)
            md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, j, mi_PARAMETRE_BOUES_Q_m3j)
        End If
    Next j

'Calcul des KPI
pd_MatieresVolatilesBouesEntrees_kgMVj = BouesEntreeLyseMV_kgj + BouesNonLyseesMV_kgj


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Public Sub fonctionnement_reel(ByRef TableauRecapitulatifFluxBoues() As Double, ByRef PartGraissesInternesSurTotalCosubstrat() As Double, ByVal ChoixEauxSalesPrimaire As Boolean, ByRef DetailBouesI() As Double, ByRef DetailBouesII() As Double, ByRef DetailBouesIII() As Double, ByRef FluxDCODansBiogaz_kgj As Double, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double, ByRef TableauBesoinsThermiques() As Double, ByRef TableauDisponibilitesThermiques() As Double, ByRef CaracteristiquesBiogaz() As Double, ByRef PressionVapeur_bar() As Double, ByRef NombreReseauxVapeur As Integer, ByRef LyseAvalEpaississement() As Integer)


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

'Initialisation des variables locales
NumeroErreurNonGeree = "44.6"
I = 0
j = 0
k = 0
ReperageBoucleReel = True
IncoherenceCalcul = 0



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

'CALCUL DES BOUES EN ENTREE (VERS LYSE ET BYPASS)
Call calcul_cocktail_boues_entrees(TableauRecapitulatifFluxBoues, TableauRecapitulatifPollutionSolubleDansFluxBoues, ChoixEauxSalesPrimaire, DetailBouesI, DetailBouesII, DetailBouesIII, PartGraissesInternesSurTotalCosubstrat, LyseAvalEpaississement)

'Attribution de la température des boues en entrée de lyse
If mn_TemperatureBouesEntreeLyse_°C_force = True Then
    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureBouesEntreeLyse_°C_f
    If md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) > md_TemperatureMaximaleBouesEntreeBiothelysValiditeCalculs_°C And mi_ChoixConfiguration = pi_CONFIGURATION_BIOTHELYS Then
        MsgBox Feuil44.Range("erreur_temperature_boues_entree lyse") & " " & md_TemperatureMaximaleBouesEntreeBiothelysValiditeCalculs_°C & " °C", Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    End If
Else
    If mi_ChoixConfiguration = pi_CONFIGURATION_BIOTHELYS Then
        If pd_TemperatureEauConditionsReelles_°C > md_TemperatureMaximaleBouesEntreeBiothelysValiditeCalculs_°C Then
            md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureMaximaleBouesEntreeBiothelysValiditeCalculs_°C
        Else
            md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = pd_TemperatureEauConditionsReelles_°C
        End If
    Else
        md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = pd_TemperatureEauConditionsReelles_°C
    End If
End If
If mn_TemperatureBouesNonLysees_°C_force = True Then
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureBouesNonLysees_°C_f
Else
    md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = pd_TemperatureEauConditionsReelles_°C
End If

'Température initiale de l'eau de dilution
If mn_TemperatureEauDilution_°C_force = True Then
    md_TemperatureEauDilution_°C = md_TemperatureEauDilution_°C_f
Else
    md_TemperatureEauDilution_°C = pd_TemperatureEauConditionsReelles_°C
End If

'HYDROLYSE THERMIQUE DES BOUES
Call calcul_bilan_masse_lyse(ReperageBoucleReel)

'MELANGE AVEC LES BOUES NON HYDROLYSEES
Call calcul_bilan_masse_melangeur(mi_FLUX_ECHANGEUR55_MELANGEUR, mi_FLUX_BOUES_NON_LYSEES, mi_FLUX_MELANGEUR_DILUTION)
    
'DILUTION DES BOUES AVANT DIGESTION
Call calcul_dilution_boues_avant_digestion(ReperageBoucleReel)

'DIGESTEUR
md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_REEL_j) = md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3) / md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
Call calcul_bilan_masse_digestion(md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_REEL_j))

'possible modification de la réduction des MV
If mn_CaracteristiquesDigestion_force(mi_PARAMETRE_DIGESTION_REDUCTION_MV) = True Then
    Call forcage_rendement_MV
End If

'On calcule le biogaz
Call calcul_biogaz(FluxDCODansBiogaz_kgj)
CaracteristiquesBiogaz(Q_Nm3j) = pd_BiogazProduit_Nm3j
CaracteristiquesBiogaz(teneur_CH4) = md_RatioCH4DansBiogaz

'calcul de la pollution solubilisée
Call calcul_solubilisation_pendant_digestion(TableauRecapitulatifPollutionSolubleDansFluxBoues)

'GESTION ENERGIE
Call calcul_bilan_thermique(TableauBesoinsThermiques, TableauDisponibilitesThermiques, PressionVapeur_bar, NombreReseauxVapeur, ReperageBoucleReel)

'RECAPITULATIF SORTIE
For I = 1 To nb_type_boues
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_MS_kgj)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_Q) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_Q_m3j)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_MV_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_MV_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_DCO_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_DBO_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_NK_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_NK_MES)
    TableauRecapitulatifFluxBoues(boues_digerees, I, repere_ratio_Pt_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_P_MES)
Next I



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

'Initialisation des variables locales
NumeroErreurNonGeree = "44.7"
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



pd_ConsommationElectrique_kWhj = 0


'AGITATION DES DIGESTEUR
If mi_ChoixTypeAgitationDigesteur = mi_AGITATION_MECANIQUE Then
    ConsommationElectriqueAgitation_kWhj = md_ConsommationElectriqueSpecifiqueAgitationMecaniqueDigestion_W_m3 * md_TempsFonctionnementAgitation_hj * md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3) / 1000
ElseIf mi_ChoixTypeAgitationDigesteur = mi_AGITATION_RECIRCULATION Then
    'digesteur
    PertesDeCharge_mCE = 10    'mCE
    If mn_CaracteristiquesDigestion_force(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) = True Then
        md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) = md_CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j)
    End If
    DebitRecirculationAgitation_m3h = md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) * md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3) / md_TempsFonctionnementAgitation_hj
    RendementPompe = 0.4    'hypothèse
    PuissancePompe_kW = DebitRecirculationAgitation_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * PertesDeCharge_mCE / RendementPompe
    ConsommationElectriqueAgitation_kWhj = PuissancePompe_kW / F_rendement_moteur_pompe(PuissancePompe_kW) * md_TempsFonctionnementAgitation_hj
Else  'if choix_agitatio=mi_AGITATION_BIOGAZ
    VitesseBrassageBiogaz_Nmh = 1.25    'Nm/h   P.JOUAFFRE--> 1 ; Prisma--> entre 1.5 et 2
    PertesDeCharge_mCE = 2    'mCE
    'digesteur
    DebitBrassageBiogaz_Nm3j = VitesseBrassageBiogaz_Nmh * md_TempsFonctionnementAgitation_hj * md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3) / md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_HAUTEUR_m)
    HMTAgitationBiogaz_mCE = md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_HAUTEUR_m) + PertesDeCharge_mCE
    'EN ATTENTE POUR METTRE LES RENDEMENTS DES SURPRESSEURS CHOISIS
    ConsommationElectriqueSpecifiqueSurpresseurBiogaz_Wh_Nm3mCE = 4    'Wh/(Nm3.mCE)
    ConsommationElectriqueAgitation_kWhj = ConsommationElectriqueSpecifiqueSurpresseurBiogaz_Wh_Nm3mCE * DebitBrassageBiogaz_Nm3j * HMTAgitationBiogaz_mCE / 1000
End If
pd_ConsommationElectrique_kWhj = pd_ConsommationElectrique_kWhj + ConsommationElectriqueAgitation_kWhj




'DIFFERENTS POSTES DE POMPAGE  (on garde les mêmes termes que dans les autres fiches pour modification rapide si nécessaire --> usage de variables intermédiaires)
'lyse --> digesteur
DebitTotal_m3j = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = True Then
    md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)
    If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = True Then
        md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj)
        If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = False Then
            If Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) < (DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) Then
                md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) + 1
            Else
                md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)))
            End If
        End If
    Else
        md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))
        If md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) > pd_NOMBRE_HEURE_PAR_JOUR Then
            md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = pd_NOMBRE_HEURE_PAR_JOUR
            If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = False Then
                If Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) < (DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) Then
                    md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) + 1
                Else
                    md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)))
                End If
                md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))
            End If
        End If
    End If
Else
    md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj))
End If
NombrePompes = md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES)
PressionRefoulement_mCE = md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_HMT_mCE)
TempsFonctionnementChaquePompe_hj = md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj)
DebitUnitairePompe_m3h = md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)
IncoherenceCalcul = (DebitUnitairePompe_m3h * NombrePompes * TempsFonctionnementChaquePompe_hj - DebitTotal_m3j) / DebitTotal_m3j
If Abs(IncoherenceCalcul) > mi_CRITERE_INCOHERENCE Then
    ms_CommentairesCalculs(2) = Feuil44.Range("dig_biothelys_erreur_pompage_lyse_digesteur").Cells(1, 1)
End If
If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = True Then
    md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES)
    RendementGlobalPompe = md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES)
Else
    RendementPompe = 0.4    'hypothèse
    PuissancePompe_kW = DebitUnitairePompe_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * PressionRefoulement_mCE / RendementPompe
    RendementGlobalPompe = RendementPompe * F_rendement_moteur_pompe(PuissancePompe_kW)
    md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = RendementGlobalPompe
End If
ConsommationElectriquePompages_kWhj = DebitUnitairePompe_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * PressionRefoulement_mCE / RendementGlobalPompe * NombrePompes * TempsFonctionnementChaquePompe_hj
'Sortie du digesteur
DebitTotal_m3j = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = True Then
    md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)
    If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = True Then
        md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj)
        If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = False Then
            If Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) < (DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) Then
                md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) + 1
            Else
                md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)))
            End If
        End If
    Else
        md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))
        If md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) > pd_NOMBRE_HEURE_PAR_JOUR Then
            md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = pd_NOMBRE_HEURE_PAR_JOUR
            If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = False Then
                If Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) < (DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) Then
                    md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))) + 1
                Else
                    md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = Int(DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) * md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)))
                End If
                md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h))
            End If
        End If
    End If
Else
    md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h) = DebitTotal_m3j / (md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) * md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj))
End If
NombrePompes = md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES)
PressionRefoulement_mCE = md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_HMT_mCE)
TempsFonctionnementChaquePompe_hj = md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj)
DebitUnitairePompe_m3h = md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_DEBIT_UNITAIRE_POMPE_m3h)
IncoherenceCalcul = (DebitUnitairePompe_m3h * NombrePompes * TempsFonctionnementChaquePompe_hj - DebitTotal_m3j) / DebitTotal_m3j
If Abs(IncoherenceCalcul) > mi_CRITERE_INCOHERENCE Then
    ms_CommentairesCalculs(3) = Feuil44.Range("dig_biothelys_erreur_pompage_outlet").Cells(1, 1)
End If
If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = True Then
    md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = md_CaracteristiquePompageBoues_f(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES)
    RendementGlobalPompe = md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES)
Else
    RendementPompe = 0.4    'hypothèse
    PuissancePompe_kW = DebitUnitairePompe_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * PressionRefoulement_mCE / RendementPompe
    RendementGlobalPompe = RendementPompe * F_rendement_moteur_pompe(PuissancePompe_kW)
    md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_RENDEMENT_POMPES) = RendementGlobalPompe
End If
ConsommationElectriquePompages_kWhj = ConsommationElectriquePompages_kWhj + DebitUnitairePompe_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * PressionRefoulement_mCE / RendementGlobalPompe * NombrePompes * TempsFonctionnementChaquePompe_hj
'Pompage recirculation boues vers échangeur (si nécessaire)
If md_EnergieChauffageDigesteur_kWhj > 0 Then
    md_PressionRefoulementRecirculationBoues_mCE = 10    'mCE
    DebitUnitairePompe_m3h = md_EnergieChauffageDigesteur_kWhj / (md_CapaciteCalorifiqueBoues_kWh_m3°C * md_EcartTemperatureBouesEchangeurDigesteur_°C) / md_TempsFonctionnementRecirculationBoues_hj
    RendementPompe = 0.4    'hypothèse
    PuissancePompe_kW = DebitUnitairePompe_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * md_PressionRefoulementRecirculationBoues_mCE / RendementPompe
    ConsommationElectriquePompages_kWhj = ConsommationElectriquePompages_kWhj + PuissancePompe_kW / F_rendement_moteur_pompe(PuissancePompe_kW) * md_TempsFonctionnementRecirculationBoues_hj
End If
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


pd_ConsommationElectriqueFixe_kWhj = ConsommationElectriqueAgitation_kWhj


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
NumeroErreurNonGeree = "44.8"
I = 0

'Ajout du nom du procédé
ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil44.Range("dig_biothelys_erreur_nom")

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
NumeroErreurNonGeree = "44.11"
NomFlux = ""
TemperatureEntreeEchangeur_°C = 0
TemperatureSortieEchangeur_°C = 0
MCp_kWh_j°C = 0



'LYSE - Chauffage de l'eau sous pression

'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
'NomFlux = "LD - Chauffage de l'eau pour produire de la vapeur"
NomFlux = Feuil5.Range("nom_flux_biothelys1").Cells(1, 1).Value
TemperatureEntreeEchangeur_°C = md_TemperatureInitialeEauPourProductionVapeur_°C
TemperatureSortieEchangeur_°C = md_TemperatureSaturation_°C
MCp_kWh_j°C = md_PinchChauffageEauPourVapeur_kWhj / (TemperatureSortieEchangeur_°C - TemperatureEntreeEchangeur_°C)
'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
Call copie_dans_flux_thermique_process(NomFlux, TemperatureEntreeEchangeur_°C, TemperatureSortieEchangeur_°C, MCp_kWh_j°C)


'LYSE - Vaporisation

'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
'    NomFlux = "LD - Vaporisation de l'eau"
NomFlux = Feuil5.Range("nom_flux_biothelys2").Cells(1, 1).Value
TemperatureEntreeEchangeur_°C = md_TemperatureSaturation_°C
TemperatureSortieEchangeur_°C = md_TemperatureSaturation_°C + 0.01
MCp_kWh_j°C = md_PinchVaporisation_kWhj / (TemperatureSortieEchangeur_°C - TemperatureEntreeEchangeur_°C)
'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
Call copie_dans_flux_thermique_process(NomFlux, TemperatureEntreeEchangeur_°C, TemperatureSortieEchangeur_°C, MCp_kWh_j°C)


'BIOTHELYS - Refroidissement des boues avant dilution

'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
'    NomFlux = "LD - Refroidissement des boues avant dilution"
If md_EnergieRefroidissementBouesLysees_kWhj > 0 Then
    NomFlux = Feuil5.Range("nom_flux_biothelys3").Cells(1, 1).Value
    TemperatureEntreeEchangeur_°C = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
    TemperatureSortieEchangeur_°C = md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
    MCp_kWh_j°C = md_EnergieRefroidissementBouesLysees_kWhj / (TemperatureEntreeEchangeur_°C - TemperatureSortieEchangeur_°C)
    'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
    Call copie_dans_flux_thermique_process(NomFlux, TemperatureEntreeEchangeur_°C, TemperatureSortieEchangeur_°C, MCp_kWh_j°C)
Else
    If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
        MsgBox Feuil44.Range("digestionLD_erreur_BesoinChauffageBouesApresLyse"), vbExclamation, Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        If md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) < md_TemperatureDigestionMinimale_°C Or md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) > md_TemperatureDigestionMaximale_°C Then
            MsgBox Feuil44.Range("digestionLD_erreur_TemperatureDigestion") & ": " & md_TemperatureDigestionMinimale_°C & " - " & md_TemperatureDigestionMaximale_°C & " °C", Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        End If
    End If
End If


'BIOTHELYS / EXELYS LD - Chauffage de l'eau de dilution

'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
'    NomFlux = "LD - Chauffage eau de dilution"
If md_EnergieChauffageEauDilution_kWhj > 0 Then
    NomFlux = Feuil5.Range("nom_flux_biothelys4").Cells(1, 1).Value
    TemperatureEntreeEchangeur_°C = md_TemperatureEauDilution_°C
    MCp_kWh_j°C = (md_DebitEauDilution_m3j * md_MasseVolumiqueEau_kg_m3) * (md_CapaciteCalorifiqueEau_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE)
    TemperatureSortieEchangeur_°C = TemperatureEntreeEchangeur_°C + md_EnergieChauffageEauDilution_kWhj / MCp_kWh_j°C
    'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
    Call copie_dans_flux_thermique_process(NomFlux, TemperatureEntreeEchangeur_°C, TemperatureSortieEchangeur_°C, MCp_kWh_j°C)
End If


'BIOTHELYS / EXELYS LD - Chauffage du digesteur

'récupération des caractéristique du flux sur le procédé -> propre à chaque procédé
'    NomFlux = "LD - Chauffage du digesteur"
If md_EnergieChauffageDigesteur_kWhj > 0 Then
    NomFlux = Feuil5.Range("nom_flux_biothelys5").Cells(1, 1).Value
    TemperatureEntreeEchangeur_°C = md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C)
    TemperatureSortieEchangeur_°C = md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) + md_EcartTemperatureBouesEchangeurDigesteur_°C
    MCp_kWh_j°C = md_EnergieChauffageDigesteur_kWhj / (TemperatureSortieEchangeur_°C - TemperatureEntreeEchangeur_°C)
    'remplissage des caractéristique du nouveau flux -> commun aux différents procédés
    Call copie_dans_flux_thermique_process(NomFlux, TemperatureEntreeEchangeur_°C, TemperatureSortieEchangeur_°C, MCp_kWh_j°C)
End If


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
Dim BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj As Double
Dim BesoinThermiqueVaporisationEauPrechauffee_kWhj As Double
Dim DebitVolumiqueBouesEntreeDigesteur_m3j As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "44.24"
BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj = 0
BesoinThermiqueVaporisationEauPrechauffee_kWhj = 0
DebitVolumiqueBouesEntreeDigesteur_m3j = 0

'On initialise la valeur de température de l'eau de dilution après chauffage
md_TemperatureEauDilutionApresChauffage_°C = 0

'On réattribue la température du digesteur
If mn_CaracteristiquesDigestion_force(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = False Then
    md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = md_TemperatureDigestionReference_°C
End If

'préchauffage de l'eau pour la vapeur et température inter échangeurs
BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj = md_DebitVapeurLyse_kgj * md_CapaciteCalorifiqueEau_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE * (md_TemperaturePrechauffageEauPourProductionVapeur_°C - md_TemperatureInitialeEauPourProductionVapeur_°C) / md_EfficaciteEchangeur

If choix_analyse_pinch = False Then
    If mn_BesoinsThermiquesEnergieEau90°C_kWhj_force = True Then
        md_BesoinsThermiquesEnergieEau90°C_kWhj = md_BesoinsThermiquesEnergieEau90°C_kWhj_f
    Else
        md_BesoinsThermiquesEnergieEau90°C_kWhj = BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj
        TableauBesoinsThermiques(energie_eau_HT) = TableauBesoinsThermiques(energie_eau_HT) + BesoinThermiquePrechauffageEauPourProductionVapeur_kWhj
    End If
End If

'vaporisation de l'eau préchauffée
If NombreReseauxVapeur < nb_vapeur_max Then
    NombreReseauxVapeur = NombreReseauxVapeur + 1
    PressionVapeur_bar(NombreReseauxVapeur) = md_PressionAbsolueVapeurSaturanteLyse_bar
    BesoinThermiqueVaporisationEauPrechauffee_kWhj = (md_DebitVapeurLyse_kgj * (md_EnthalpieTotaleVapeurSaturee_kcal_kg - md_CapaciteCalorifiqueEau_kcal_kg°C * md_TemperaturePrechauffageEauPourProductionVapeur_°C) * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE) / md_EfficaciteEchangeur
    TableauBesoinsThermiques(energie_eau_HT + NombreReseauxVapeur) = BesoinThermiqueVaporisationEauPrechauffee_kWhj
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


'CALCUL DES TEMPERATURES DES BOUES ET DE L'ENERGIE RECUPEREE SUR L'ECHANGEUR
If mn_EnergieRefroidissementBouesLysees_kWhj_force = True And RepereCasExploitation = True Then
    md_EnergieRefroidissementBouesLysees_kWhj = md_EnergieRefroidissementBouesLysees_kWhj_f
    'Calcul de la température des boues en sortie d'échangeur
    md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) - md_EnergieRefroidissementBouesLysees_kWhj / (md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE)
    'Calcul de la température après mélange
    md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = F_temperature_melange_°C(md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3, md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C), md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3, md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C), md_CapaciteCalorifiqueBoues_kcal_kg°C)
    'On regarde si l'on peut/doit chauffer l'eau de dilution
    If md_DebitEauDilution_m3j > 0 Then
        If mn_EnergieChauffageEauDilution_kWhj_force = True Then
            md_EnergieChauffageEauDilution_kWhj = md_EnergieChauffageEauDilution_kWhj_f
            md_TemperatureEauDilutionApresChauffage_°C = md_TemperatureEauDilution_°C + md_EnergieChauffageEauDilution_kWhj / ((md_DebitEauDilution_m3j * md_MasseVolumiqueEau_kg_m3) * (md_CapaciteCalorifiqueEau_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE))
        Else
            'On calcule la température nécessaire de l'eau de dilution
            md_TemperatureEauDilutionApresChauffage_°C = F_temperature_entree_flux1_avant_melange(md_DebitEauDilution_m3j * md_MasseVolumiqueEau_kg_m3, md_CapaciteCalorifiqueEau_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3, md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C), md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C))
            'On vérifie que la température n'est pas trop élevée ou trop basse
            If md_TemperatureEauDilutionApresChauffage_°C > md_TemperatureEauDilutionMaximale_°C Then
                md_TemperatureEauDilutionApresChauffage_°C = md_TemperatureEauDilution_°C
            ElseIf md_TemperatureEauDilutionApresChauffage_°C < md_TemperatureEauDilution_°C Then
                md_TemperatureEauDilutionApresChauffage_°C = md_TemperatureEauDilution_°C
            End If
            md_EnergieChauffageEauDilution_kWhj = (md_DebitEauDilution_m3j * md_MasseVolumiqueEau_kg_m3) * (md_CapaciteCalorifiqueEau_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE) * (md_TemperatureEauDilutionApresChauffage_°C - md_TemperatureEauDilution_°C)
        End If
        'Température des boues en entrée du digesteur
        md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = F_temperature_melange_°C(md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3, md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C), md_DebitEauDilution_m3j * md_MasseVolumiqueEau_kg_m3, md_CapaciteCalorifiqueEau_kcal_kg°C, md_TemperatureEauDilutionApresChauffage_°C, md_CapaciteCalorifiqueBoues_kcal_kg°C)
    Else
        'Pas de dilution --> température des boues ne change pas
        md_EnergieChauffageEauDilution_kWhj = 0
        md_TemperatureEauDilutionApresChauffage_°C = md_TemperatureEauDilution_°C
        md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
    End If
    'On regarde si l'on doit chauffer le digesteur (par recirculation externe de boues)
    If mn_EnergieChauffageDigesteur_kWhj_force = True Then
        md_EnergieChauffageDigesteur_kWhj = md_EnergieChauffageDigesteur_kWhj_f
        If mn_CaracteristiquesDigestion_force(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = False Then
            md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) + (md_EnergieChauffageDigesteur_kWhj - md_PertesThermiquesDigestion_kWh_m3j(choix_climat, mi_ChoixIsolationDigestion) * md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3)) / (md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE)
        End If
    Else
        'On calcule le besoin de chauffage
        md_EnergieChauffageDigesteur_kWhj = md_PertesThermiquesDigestion_kWh_m3j(choix_climat, mi_ChoixIsolationDigestion) * md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3) - (md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE) * (md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) - md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C))
        'Si besoin négatif (pas de besoin), on recalcule la température du digesteur si elle n'est pas forcée
        If md_EnergieChauffageDigesteur_kWhj < 0 Then
            md_EnergieChauffageDigesteur_kWhj = 0
            If mn_CaracteristiquesDigestion_force(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = False Then
                md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) + (md_EnergieChauffageDigesteur_kWhj - md_PertesThermiquesDigestion_kWh_m3j(choix_climat, mi_ChoixIsolationDigestion) * md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3)) / (md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE)
            End If
        End If
    End If
Else
    'Calcul de la température requise des boues en entrée de digestion (après dilution)
    md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) + md_PertesThermiquesDigestion_kWh_m3j(choix_climat, mi_ChoixIsolationDigestion) * md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3) / (md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE)
    'Calcul de la température requise des boues avant la dilution
    md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = F_temperature_entree_flux1_avant_melange(md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3, md_CapaciteCalorifiqueBoues_kcal_kg°C, md_DebitEauDilution_m3j * md_MasseVolumiqueEau_kg_m3, md_CapaciteCalorifiqueEau_kcal_kg°C, md_TemperatureEauDilution_°C, md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C))
    'Calcul de la température requise des boues avant mélange avec les boues non hydrolysées
    md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = F_temperature_entree_flux1_avant_melange(md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3, md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3, md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C), md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C))
    If md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) <= md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) Then
        'Calcul de l'énergie récupérée dans l'échangeur de refroidissement
        md_EnergieRefroidissementBouesLysees_kWhj = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE * (md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) - md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C))
        md_EnergieChauffageEauDilution_kWhj = 0
        md_EnergieChauffageEauDilution_kWhj = 0
        md_EnergieChauffageDigesteur_kWhj = 0
    Else
        'Récupération d'énergie impossible, nécessité de chauffer l'eau de dilution ou le digesteur. On recalcule les températures dans l'autre sens
        'Pas de récupération d'énergie
        md_EnergieRefroidissementBouesLysees_kWhj = 0
        md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
        'Température après mélange des boues
        md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = F_temperature_melange_°C(md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3, md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_ECHANGEUR55_MELANGEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C), md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3, md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C), md_CapaciteCalorifiqueBoues_kcal_kg°C)
        'On regarde si l'on peut chauffer l'eau de dilution
        If md_DebitEauDilution_m3j > 0 Then
            If mn_EnergieChauffageEauDilution_kWhj_force = True And RepereCasExploitation = True Then
                md_EnergieChauffageEauDilution_kWhj = md_EnergieChauffageEauDilution_kWhj_f
                md_TemperatureEauDilutionApresChauffage_°C = md_TemperatureEauDilution_°C + md_EnergieChauffageEauDilution_kWhj / ((md_DebitEauDilution_m3j * md_MasseVolumiqueEau_kg_m3) * (md_CapaciteCalorifiqueEau_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE))
            Else
                'On calcule la température nécessaire de l'eau de dilution
                md_TemperatureEauDilutionApresChauffage_°C = F_temperature_entree_flux1_avant_melange(md_DebitEauDilution_m3j * md_MasseVolumiqueEau_kg_m3, md_CapaciteCalorifiqueEau_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3, md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C), md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C))
                'On vérifie que la température n'est pas trop élevée
                If md_TemperatureEauDilutionApresChauffage_°C > md_TemperatureEauDilutionMaximale_°C Then
                    md_TemperatureEauDilutionApresChauffage_°C = md_TemperatureEauDilution_°C
                ElseIf md_TemperatureEauDilutionApresChauffage_°C < md_TemperatureEauDilution_°C Then
                    md_TemperatureEauDilutionApresChauffage_°C = md_TemperatureEauDilution_°C
                End If
                md_EnergieChauffageEauDilution_kWhj = (md_DebitEauDilution_m3j * md_MasseVolumiqueEau_kg_m3) * (md_CapaciteCalorifiqueEau_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE) * (md_TemperatureEauDilutionApresChauffage_°C - md_TemperatureEauDilution_°C)
            End If
            'Température des boues en entrée du digesteur
            md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = F_temperature_melange_°C(md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3, md_CapaciteCalorifiqueBoues_kcal_kg°C, md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C), md_DebitEauDilution_m3j * md_MasseVolumiqueEau_kg_m3, md_CapaciteCalorifiqueEau_kcal_kg°C, md_TemperatureEauDilutionApresChauffage_°C, md_CapaciteCalorifiqueBoues_kcal_kg°C)
        Else
            'Pas de dilution --> température des boues ne change pas
            md_EnergieChauffageEauDilution_kWhj = 0
            md_TemperatureEauDilutionApresChauffage_°C = md_TemperatureEauDilution_°C
            md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
        End If
        'On regarde si l'on doit chauffer le digesteur (par recirculation externe de boues)
        If mn_EnergieChauffageDigesteur_kWhj_force = True And RepereCasExploitation = True Then
            md_EnergieChauffageDigesteur_kWhj = md_EnergieChauffageDigesteur_kWhj_f
            If mn_CaracteristiquesDigestion_force(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = False Then
                md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) + (md_EnergieChauffageDigesteur_kWhj - md_PertesThermiquesDigestion_kWh_m3j(choix_climat, mi_ChoixIsolationDigestion) * md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3)) / (md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE)
            End If
        Else
            'On calcule le besoin de chauffage
            md_EnergieChauffageDigesteur_kWhj = md_PertesThermiquesDigestion_kWh_m3j(choix_climat, mi_ChoixIsolationDigestion) * md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3) - (md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE) * (md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) - md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C))
            'Si besoin négatif (pas de besoin), on recalcule la température du digesteur si elle n'est pas forcée
            If md_EnergieChauffageDigesteur_kWhj < 0 Then
                md_EnergieChauffageDigesteur_kWhj = 0
                If mn_CaracteristiquesDigestion_force(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = False Then
                    md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) + (md_EnergieChauffageDigesteur_kWhj - md_PertesThermiquesDigestion_kWh_m3j(choix_climat, mi_ChoixIsolationDigestion) * md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_VOLUME_m3)) / (md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * pd_CONVERSION_kJ_PAR_kcal / pd_NOMBRE_SECONDE_PAR_HEURE)
                End If
            End If
        End If
    End If
End If
TableauDisponibilitesThermiques(energie_eau_HT) = TableauDisponibilitesThermiques(energie_eau_HT) + md_EnergieRefroidissementBouesLysees_kWhj
TableauBesoinsThermiques(energie_eau_BT) = TableauBesoinsThermiques(energie_eau_BT) + md_EnergieChauffageDigesteur_kWhj
TableauBesoinsThermiques(energie_eau_HT) = TableauBesoinsThermiques(energie_eau_HT) + md_EnergieChauffageEauDilution_kWhj


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
'       Les valeurs (entiers) associées aux choix se trouvent dans l'onglet "Feuil44" nommé "Biothelys_Exelys_LD" où sont également reportées les valeurs saisies
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
NumeroErreurNonGeree = "44.1"
I = 0


mi_ChoixConfiguration = Feuil44.Range("digestionLD_ChoixConfiguration").Cells(1, pi_FiliereConsideree)

'Boues en entrée de digestion
For I = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    mi_ChoixEntreesBoues(I) = Feuil44.Range("dig_biothelys_choix_inlet").Cells(I, pi_FiliereConsideree)
    If mi_ChoixEntreesBoues(I) <> mi_FLUX_ENTREE_BOUES_NON_UTILISE Then
        md_EntreeBouesPourcentageFluxAdmis(I) = Feuil44.Range("dig_biothelys_inlet_ratio").Cells(I, pi_FiliereConsideree)
        If mi_ChoixEntreesBoues(I) <= mi_FLUX_ENTREE_GRAISSES Then
            mi_ProvenanceBouesEntree(I) = boues_inlet
            mi_TypeBouesEntree(I) = mi_ChoixEntreesBoues(I)
        Else
            mi_ProvenanceBouesEntree(I) = boues_epaissies(mi_ChoixEntreesBoues(I) - mi_FLUX_ENTREE_GRAISSES)
            'ici on prendra directement les valeurs du flux antérieur
        End If
        mi_ChoixPositionEntreeBoues(I) = Feuil44.Range("digestionLD_ChoixEntreesBoues").Cells(I, pi_FiliereConsideree)
    End If
Next I

'agitation
mi_ChoixTypeAgitationDigesteur = Feuil44.Range("dig_biothelys_choix_agitation").Cells(1, pi_FiliereConsideree)

'stockage_biogaz
mi_ChoixStockageBiogaz = Feuil44.Range("dig_biothelys_choix_biogaz_stock").Cells(1, pi_FiliereConsideree)

'isolation
mi_ChoixIsolationDigestion = Feuil44.Range("dig_biothelys_choix_isolation").Cells(1, pi_FiliereConsideree)


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
'       Les valeurs sont reportées dans l'onglet "Feuil44" nommée "Biothelys_Exelys_LD" dont les noms de plages de cellules sont définis dans Excel

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
NumeroErreurNonGeree = "44.2"
I = 1
Colonne = pi_FiliereConsideree


'température entrée digesteur
mn_TemperatureBouesEntreeLyse_°C_force = Feuil44.Range("dig_biothelys_inlet_boues").Cells(mi_PARAMETRE_BOUES_TEMPERATURE_°C, Colonne)
'digesteur
For I = 1 To mi_NOMBRE_PARAMETRE_DIGESTION
    mn_CaracteristiquesDigestion_force(I) = Feuil44.Range("dig_biothelys_digesteur").Cells(I, Colonne)
    If mn_CaracteristiquesDigestion_force(I) = True Then
        md_CaracteristiquesDigestion_f(I) = Feuil44.Range("dig_biothelys_digesteur").Cells(I, Colonne)
    End If
Next I
mn_DebitVapeurLyse_kgj_force = Feuil44.Range("dig_biothelys_lyse_vapeur").Cells(1, Colonne)
mn_PressionAbsolueVapeurSaturanteLyse_bar_force = Feuil44.Range("dig_biothelys_lyse_vapeur_pression").Cells(1, Colonne)
mn_TemperatureInitialeEauPourProductionVapeur_°C_force = Feuil44.Range("dig_biothelys_vapeur_Teau_ini").Cells(1, Colonne)
mn_TemperaturePrechauffageEauPourProductionVapeur_°C_force = Feuil44.Range("dig_biothelys_vapeur_T_prechauff_eau").Cells(1, Colonne)
mn_TemperatureSortieLyse_force = Feuil44.Range("digestionLD_BouesLysees_°C").Cells(1, Colonne)
mn_TemperatureBouesNonLysees_°C_force = Feuil44.Range("digestionLD_BouesNonLysees_°C").Cells(1, Colonne)
mn_DebitEauDilution_m3j_force = Feuil44.Range("digestionLD_DebitEauDilution_m3j").Cells(1, Colonne)
mn_TemperatureEauDilution_°C_force = Feuil44.Range("digestionLD_TemperatureEauDilution_°C").Cells(1, Colonne)

'pompage digesteur
For I = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    mn_CaracteristiquePompageBoues_force(mi_POMPAGE_LYSE_DIGESTEUR, I) = Feuil44.Range("dig_biothelys_lyse_digesteur_pompage").Cells(I, Colonne)
    If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_LYSE_DIGESTEUR, I) = True Then
        md_CaracteristiquePompageBoues_f(mi_POMPAGE_LYSE_DIGESTEUR, I) = Feuil44.Range("dig_biothelys_lyse_digesteur_pompage").Cells(I, Colonne)
    End If
Next I
'pompage sortie
For I = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    mn_CaracteristiquePompageBoues_force(mi_POMPAGE_BOUES_DIGEREES, I) = Feuil44.Range("dig_biothelys_outlet_pompage").Cells(I, Colonne)
    If mn_CaracteristiquePompageBoues_force(mi_POMPAGE_BOUES_DIGEREES, I) = True Then
        md_CaracteristiquePompageBoues_f(mi_POMPAGE_BOUES_DIGEREES, I) = Feuil44.Range("dig_biothelys_outlet_pompage").Cells(I, Colonne)
    End If
Next I
mn_RatioCH4DansBiogaz_force = Feuil44.Range("dig_biothelys_biogaz_teneur_CH4").Cells(1, Colonne)
mn_BiogazProduit_Nm3j_force = Feuil44.Range("dig_biothelys_biogaz_Q_Nm3j").Cells(1, Colonne)
mn_BesoinTotalEnergieProductionVapeur_kWhj_force = Feuil44.Range("dig_biothelys_energie_vapeur_kWhj").Cells(1, Colonne)
If choix_analyse_pinch = False Then
    mn_BesoinsThermiquesEnergieEau90°C_kWhj_force = Feuil44.Range("dig_biothelys_energie_eau90").Cells(1, Colonne)
Else
    mn_EnergieChauffageEauDilution_kWhj_force = Feuil44.Range("dig_biothelys_energie_eau90").Cells(1, Colonne)
End If
mn_EnergieChauffageDigesteur_kWhj_force = Feuil44.Range("digestionLD_EnergieChauffageDigesteur_kWhj").Cells(1, Colonne)
mn_EnergieRefroidissementBouesLysees_kWhj_force = Feuil44.Range("dig_biothelys_energie_refroidissement_boues_kWhj").Cells(1, Colonne)
mn_EnergieDisponibleEau90°C_kWhj_force = Feuil44.Range("dig_biothelys_dispo_eau90").Cells(1, Colonne)


If mn_TemperatureBouesEntreeLyse_°C_force = True Then
    md_TemperatureBouesEntreeLyse_°C_f = Feuil44.Range("dig_biothelys_inlet_boues").Cells(mi_PARAMETRE_BOUES_TEMPERATURE_°C, Colonne)
End If
If mn_DebitVapeurLyse_kgj_force = True Then
    md_DebitVapeurLyse_kgj_f = Feuil44.Range("dig_biothelys_lyse_vapeur").Cells(1, Colonne)
End If
If mn_PressionAbsolueVapeurSaturanteLyse_bar_force = True Then
    md_PressionAbsolueVapeurSaturanteLyse_bar_f = Feuil44.Range("dig_biothelys_lyse_vapeur_pression").Cells(1, Colonne)
End If
If mn_TemperatureInitialeEauPourProductionVapeur_°C_force = True Then
    md_TemperatureInitialeEauPourProductionVapeur_°C_f = Feuil44.Range("dig_biothelys_vapeur_Teau_ini").Cells(1, Colonne)
End If
If mn_TemperaturePrechauffageEauPourProductionVapeur_°C_force = True Then
    md_TemperaturePrechauffageEauPourProductionVapeur_°C_f = Feuil44.Range("dig_biothelys_vapeur_T_prechauff_eau").Cells(1, Colonne)
End If
If mn_TemperatureSortieLyse_force = True Then
    md_TemperatureSortieLyse_f = Feuil44.Range("digestionLD_BouesLysees_°C").Cells(1, Colonne)
End If
If mn_TemperatureBouesNonLysees_°C_force = True Then
    md_TemperatureBouesNonLysees_°C_f = Feuil44.Range("digestionLD_BouesNonLysees_°C").Cells(1, Colonne)
End If
If mn_DebitEauDilution_m3j_force = True Then
    md_DebitEauDilution_m3j_f = Feuil44.Range("digestionLD_DebitEauDilution_m3j").Cells(1, Colonne)
End If
If mn_TemperatureEauDilution_°C_force = True Then
    md_TemperatureEauDilution_°C_f = Feuil44.Range("digestionLD_TemperatureEauDilution_°C").Cells(1, Colonne)
End If
If mn_RatioCH4DansBiogaz_force = True Then
    md_RatioCH4DansBiogaz_f = Feuil44.Range("dig_biothelys_biogaz_teneur_CH4").Cells(1, Colonne)
End If
If mn_BiogazProduit_Nm3j_force = True Then
    md_BiogazProduit_Nm3j_f = Feuil44.Range("dig_biothelys_biogaz_Q_Nm3j").Cells(1, Colonne)
End If
If mn_BesoinTotalEnergieProductionVapeur_kWhj_force = True Then
    md_BesoinTotalEnergieProductionVapeur_kWhj_f = Feuil44.Range("dig_biothelys_energie_vapeur_kWhj").Cells(1, Colonne)
End If
If mn_BesoinsThermiquesEnergieEau90°C_kWhj_force = True Then
    md_BesoinsThermiquesEnergieEau90°C_kWhj_f = Feuil44.Range("dig_biothelys_energie_eau90").Cells(1, Colonne)
End If
If mn_EnergieChauffageEauDilution_kWhj_force = True Then
    md_EnergieChauffageEauDilution_kWhj_f = Feuil44.Range("dig_biothelys_energie_eau90").Cells(1, Colonne)
End If
If mn_EnergieRefroidissementBouesLysees_kWhj_force = True Then
    md_EnergieRefroidissementBouesLysees_kWhj_f = Feuil44.Range("dig_biothelys_energie_refroidissement_boues_kWhj").Cells(1, Colonne)
End If
If mn_EnergieDisponibleEau90°C_kWhj_force = True Then
    md_EnergieDisponibleEau90°C_kWhj_f = Feuil44.Range("dig_biothelys_dispo_eau90").Cells(1, Colonne)
End If
If mn_EnergieChauffageDigesteur_kWhj_force = True Then
    md_EnergieChauffageDigesteur_kWhj_f = Feuil44.Range("digestionLD_EnergieChauffageDigesteur_kWhj").Cells(1, Colonne)
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
'       qui se fait dans l'onglet "Feuil44" nommé "Biothelys_Exelys_LD" dans des plages de cellules dont les noms sont définis dans Excel
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim Ligne As Integer
Dim Colonne As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "44.9"
Ligne = 1
Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


For Ligne = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    Feuil44.Range("dig_biothelys_inlet_ratio").Cells(Ligne, Colonne) = md_EntreeBouesPourcentageFluxAdmis(Ligne)
Next Ligne

For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil44.Range("dig_biothelys_inlet_boues").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, Ligne)
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_RATIO_MV_MES
    Feuil44.Range("dig_biothelys_outlet_boues").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, Ligne)
Next Ligne
For Ligne = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    Feuil44.Range("dig_biothelys_lyse_digesteur_pompage").Cells(Ligne, Colonne) = md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, Ligne)
Next Ligne
For Ligne = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    Feuil44.Range("dig_biothelys_outlet_pompage").Cells(Ligne, Colonne) = md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, Ligne)
Next Ligne
'digesteur
For Ligne = 1 To mi_NOMBRE_PARAMETRE_DIGESTION
    Feuil44.Range("dig_biothelys_digesteur").Cells(Ligne, Colonne) = md_CaracteristiquesDigestion(Ligne)
Next Ligne


Ligne = 1

Feuil44.Range("dig_biothelys_lyse_vapeur").Cells(Ligne, Colonne) = md_DebitVapeurLyse_kgj
Feuil44.Range("dig_biothelys_lyse_vapeur_pression").Cells(Ligne, Colonne) = md_PressionAbsolueVapeurSaturanteLyse_bar
Feuil44.Range("dig_biothelys_vapeur_Teau_ini").Cells(1, Colonne) = md_TemperatureInitialeEauPourProductionVapeur_°C
Feuil44.Range("dig_biothelys_vapeur_T_prechauff_eau").Cells(1, Colonne) = md_TemperaturePrechauffageEauPourProductionVapeur_°C
Feuil44.Range("digestionLD_BouesLysees_kgMSj").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj)
Feuil44.Range("digestionLD_BouesLysees_m3j").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
Feuil44.Range("digestionLD_BouesLysees_gL").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
Feuil44.Range("digestionLD_BouesLysees_RatioMVMES").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES)
Feuil44.Range("digestionLD_BouesLysees_°C").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
Feuil44.Range("digestionLD_BouesNonLysees_kgMSj").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj)
Feuil44.Range("digestionLD_BouesNonLysees_m3j").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
Feuil44.Range("digestionLD_BouesNonLysees_gL").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
Feuil44.Range("digestionLD_BouesNonLysees_RatioMVMES").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES)
Feuil44.Range("digestionLD_BouesNonLysees_°C").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
Feuil44.Range("digestionLD_DebitEauDilution_m3j").Cells(Ligne, Colonne) = md_DebitEauDilution_m3j
Feuil44.Range("digestionLD_TemperatureEauDilution_°C").Cells(Ligne, Colonne) = md_TemperatureEauDilution_°C
Feuil44.Range("digestionLD_TemperatureEauDilutionApresChauffage_°C").Cells(Ligne, Colonne) = md_TemperatureEauDilutionApresChauffage_°C
Feuil44.Range("dig_biothelys_BouesEntreeDigestion_kgMSj").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj)
Feuil44.Range("dig_biothelys_BouesEntreeDigestion_m3j").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
Feuil44.Range("dig_biothelys_BouesEntreeDigestion_gL").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
Feuil44.Range("dig_biothelys_md_BouesEntreeDigestionRatioMVSurMS").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES)
Feuil44.Range("dig_biothelys_md_BouesEntreeDigestionTemperature_°C").Cells(Ligne, Colonne) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
Feuil44.Range("dig_biothelys_biogaz_teneur_CH4").Cells(Ligne, Colonne) = md_RatioCH4DansBiogaz
Feuil44.Range("dig_biothelys_ratio_biogaz_MV").Cells(Ligne, Colonne) = md_ProductionSpecifiqueBiogaz_Nm3_kgMV
Feuil44.Range("dig_biothelys_biogaz_Q_Nm3j").Cells(Ligne, Colonne) = pd_BiogazProduit_Nm3j
If choix_analyse_pinch = False Then
    Feuil44.Range("dig_biothelys_energie_eau90").Cells(Ligne, Colonne) = md_BesoinsThermiquesEnergieEau90°C_kWhj
    Feuil44.Range("dig_biothelys_energie_vapeur_kWhj").Cells(Ligne, Colonne) = md_BesoinTotalEnergieProductionVapeur_kWhj - md_BesoinsThermiquesEnergieEau90°C_kWhj
Else
    Feuil44.Range("dig_biothelys_energie_eau90").Cells(Ligne, Colonne) = md_EnergieChauffageEauDilution_kWhj
    Feuil44.Range("dig_biothelys_energie_vapeur_kWhj").Cells(Ligne, Colonne) = md_BesoinTotalEnergieProductionVapeur_kWhj
End If
Feuil44.Range("digestionLD_EnergieChauffageDigesteur_kWhj").Cells(Ligne, Colonne) = md_EnergieChauffageDigesteur_kWhj
Feuil44.Range("dig_biothelys_energie_refroidissement_boues_kWhj").Cells(Ligne, Colonne) = md_EnergieRefroidissementBouesLysees_kWhj
Feuil44.Range("dig_biothelys_dispo_eau90").Cells(Ligne, Colonne) = md_EnergieDisponibleEau90°C_kWhj
Feuil44.Range("dig_biothelys_electricite").Cells(Ligne, Colonne) = pd_ConsommationElectrique_kWhj


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Private Sub calcul_dilution_boues_avant_digestion(ByVal ReperageBoucleReel As Boolean)

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

'Initialisation des variables locales
NumeroErreurNonGeree = "44.27"
I = 0

'Calcul du débit d'eau de dilution
If mn_DebitEauDilution_m3j_force = True And ReperageBoucleReel = True Then
    md_DebitEauDilution_m3j = md_DebitEauDilution_m3j_f
Else
    If md_ConcentrationBouesEntreeDigestionSouhaitee_gL > md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL) Then
        md_DebitEauDilution_m3j = 0
    Else
        md_DebitEauDilution_m3j = md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_ConcentrationBouesEntreeDigestionSouhaitee_gL - md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
    End If
End If

'Bilan matière
md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj)
md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) + md_DebitEauDilution_m3j
md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)

For I = 1 To nb_type_boues + 1
    If I <= nb_type_boues Then
        md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, I, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, I, mi_PARAMETRE_BOUES_MS_kgj)
        md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, I, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
        md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, I, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, I, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, I, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
    End If
    md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, I, mi_PARAMETRE_BOUES_RATIO_MV_MES)
    'md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, i, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = 0
    md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
    md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, I, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, I, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
    md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, I, mi_PARAMETRE_BOUES_RATIO_NK_MES) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, I, mi_PARAMETRE_BOUES_RATIO_NK_MES)
    md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, I, mi_PARAMETRE_BOUES_RATIO_P_MES) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, I, mi_PARAMETRE_BOUES_RATIO_P_MES)
    md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, I, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(mi_FLUX_MELANGEUR_DILUTION, I, mi_PARAMETRE_BOUES_REDUCTION_MV)
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
NumeroErreurNonGeree = "44.23"
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
        'md_FluxIntermediaires(RepereFluxBouesMelangeResultant, i, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = 0
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

Private Sub calcul_biogaz(ByRef FluxDCODansBiogaz_kgj As Double)

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
NumeroErreurNonGeree = "44.25"
I = 0
DensiteBiogaz_kg_Nm3 = 0

'calcul du biogaz produit dans digesteur
DensiteBiogaz_kg_Nm3 = 0
For I = 0 To 1
    DensiteBiogaz_kg_Nm3 = DensiteBiogaz_kg_Nm3 + md_CoefficientsCalculDensiteBiogaz(I) * md_RatioCH4DansBiogaz ^ I
Next I
md_ProductionSpecifiqueBiogaz_Nm3_kgMV = 1 / DensiteBiogaz_kg_Nm3
pd_BiogazProduit_Nm3j = md_ProductionSpecifiqueBiogaz_Nm3_kgMV * md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV)

'calcul de la DCO du biogaz et ajustement de la DCO des boues digérées
FluxDCODansBiogaz_kgj = (md_RatioCH4DansBiogaz * pd_BiogazProduit_Nm3j) / md_ProductionSpecifiqueMethane_Nm3_kgDCO
md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) - FluxDCODansBiogaz_kgj

'on reclacule pour chaque type de boues la DCO des boues digérées
For I = 1 To nb_type_boues
    md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) - md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_REDUCTION_MV) * FluxDCODansBiogaz_kgj
Next I



Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub
Private Sub calcul_solubilisation_pendant_digestion(ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double)

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
NumeroErreurNonGeree = "44.28"
I = 0
FluxSolubilise_kgj = 0

'calcul de la pollution solubilisée lors de la digestion
For I = 1 To nb_type_boues
    If md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_MS_kgj) <> 0 Then
        'DCO
        FluxSolubilise_kgj = md_RatioDCOSolubleSurDCOTotaleBouesDigerees * md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) * F_facteur_correction_temperature_digesteur(md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C))
        TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_DCO) = FluxSolubilise_kgj / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_Q_m3j) * 1000
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = (md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_DCO_MES) - FluxSolubilise_kgj) / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_MS_kgj)
        'DBO
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = 0    'hypothèse
        'NK
        FluxSolubilise_kgj = md_RatioNKSolubleBouesDigereesSurMVDegrades * md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_REDUCTION_MV) * md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) * md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_NK_MES) / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_MV_MES)
        TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_NK) = FluxSolubilise_kgj / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_Q_m3j) * 1000
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_NK_MES) = (md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_NK_MES) - FluxSolubilise_kgj) / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_MS_kgj)
        'Pt
        FluxSolubilise_kgj = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_P_MES) * md_RatioPSolubleBouesDigereesSurPBouesEntree * F_facteur_correction_temperature_digesteur(md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C))
        TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_digerees, I, repere_mgL_Pt) = FluxSolubilise_kgj / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_Q_m3j) * 1000
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_P_MES) = (md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_P_MES) - FluxSolubilise_kgj) / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_MS_kgj)
        'On recalcule bien les MV
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_MS_kgj)
    End If
Next I

md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj)


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Private Sub calcul_bilan_masse_lyse(ByVal ReperageBoucleReel As Boolean)

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
Dim ConsommationSpecifiqueVapeur_kWh_m3Boues As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "44.22"
I = 0
j = 0
ConsommationSpecifiqueVapeur_kWh_m3Boues = 0


'Toutes les valeurs concernant la pollution ne changent pas jusqu'au mélangeur
For I = mi_FLUX_LYSE_ECHANGEUR55 To mi_FLUX_ECHANGEUR55_MELANGEUR
    For j = 1 To nb_type_boues + 1
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_MS_kgj)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_REDUCTION_MV)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_NK_MES)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, j, mi_PARAMETRE_BOUES_RATIO_P_MES)
    Next j
Next I

'calcul du débit de vapeur en kg/j, du débit volumique et de la concentration des boues en sortie de la lyse
If mn_DebitVapeurLyse_kgj_force = True And ReperageBoucleReel = True Then
    md_DebitVapeurLyse_kgj = md_DebitVapeurLyse_kgj_f
Else
    If mi_ChoixConfiguration = pi_CONFIGURATION_BIOTHELYS Then
        ConsommationSpecifiqueVapeur_kWh_m3Boues = (md_CoefficientAConsommationSpecifiqueVapeur_kWh_m3Boues°C * md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) + md_CoefficientBConsommationSpecifiqueVapeur_kWh_m3Boues)
        md_DebitVapeurLyse_kgj = (ConsommationSpecifiqueVapeur_kWh_m3Boues * md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) / md_EnthalpieTotaleVapeurSaturee_kcal_kg) * (pd_NOMBRE_SECONDE_PAR_HEURE / pd_CONVERSION_kJ_PAR_kcal)
    ElseIf mi_ChoixConfiguration = pi_CONFIGURATION_EXELYS Then
        md_DebitVapeurLyse_kgj = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) * md_MasseVolumiqueBoues_kg_m3 * md_CapaciteCalorifiqueBoues_kcal_kg°C * (md_TemperatureLyse_°C - md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)) / ((1 - md_RatioPertesVapeur) * md_EnthalpieTotaleVapeurSaturee_kcal_kg - md_CapaciteCalorifiqueBoues_kcal_kg°C * md_TemperatureLyse_°C)
    End If
End If
For I = mi_FLUX_LYSE_ECHANGEUR55 To mi_FLUX_ECHANGEUR55_MELANGEUR
    md_FluxIntermediaires(I, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) + md_DebitVapeurLyse_kgj / md_MasseVolumiqueBoues_kg_m3
    md_FluxIntermediaires(I, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(I, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(I, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
    For j = 1 To nb_type_boues
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(I, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
        md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(I, j, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
    Next j
Next I

'On attribue la valeur aux variables qui seront affichées à l'utilisateur
md_BouesLysees_kgMSj = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj)
md_BouesLysees_m3j = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
md_BouesLysees_gL = md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)

Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Private Function F_temperature_entree_flux1_avant_melange(ByVal DebitFlux1_kgjOUm3j As Double, ByVal CapaciteCalorifiqueFlux1_UniteEnergieChoisie_kgjOUm3j As Double, ByVal DebitFlux2_kgjOUm3j As Double, ByVal CapaciteCalorifiqueFlux2_UniteEnergieChoisie_kgjOUm3j As Double, ByVal TemperatureFlux2_UniteTemperatureChoisie As Double, ByVal CapaciteCalorifiqueFluxMelange_UniteEnergieChoisie_kgjOUm3j As Double, ByVal TemperatureFluxMelange_UniteTemperatureChoisie As Double) As Double

'************************************************************************************************************
'OBJECTIF:
'   Cette procédure calcule la quantité de vapeur utilisée pour la lyse thermique
'************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "44.31"

'On calcule la température
F_temperature_entree_flux1_avant_melange = ((DebitFlux1_kgjOUm3j + DebitFlux2_kgjOUm3j) * CapaciteCalorifiqueFluxMelange_UniteEnergieChoisie_kgjOUm3j * TemperatureFluxMelange_UniteTemperatureChoisie - DebitFlux2_kgjOUm3j * CapaciteCalorifiqueFlux2_UniteEnergieChoisie_kgjOUm3j * TemperatureFlux2_UniteTemperatureChoisie) / (DebitFlux1_kgjOUm3j * CapaciteCalorifiqueFlux1_UniteEnergieChoisie_kgjOUm3j)


Exit Function
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Function
Private Function F_temperature_melange_°C(ByVal DebitFlux1_kgjOUm3j As Double, ByVal CapaciteCalorifiqueFlux1_UniteEnergieChoisie_kgjOUm3j As Double, ByVal TemperatureFlux1_UniteTemperatureChoisie As Double, ByVal DebitFlux2_kgjOUm3j As Double, ByVal CapaciteCalorifiqueFlux2_UniteEnergieChoisie_kgjOUm3j As Double, ByVal TemperatureFlux2_UniteTemperatureChoisie As Double, ByVal CapaciteCalorifiqueFluxMelange_UniteEnergieChoisie_kgjOUm3j As Double) As Double

'************************************************************************************************************
'OBJECTIF:
'   Cette procédure calcule la température d'un mélange de flux
'************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "44.32"

'On calcule la température
F_temperature_melange_°C = (DebitFlux1_kgjOUm3j * CapaciteCalorifiqueFlux1_UniteEnergieChoisie_kgjOUm3j * TemperatureFlux1_UniteTemperatureChoisie + DebitFlux2_kgjOUm3j * CapaciteCalorifiqueFlux2_UniteEnergieChoisie_kgjOUm3j * TemperatureFlux2_UniteTemperatureChoisie) / ((DebitFlux1_kgjOUm3j + DebitFlux2_kgjOUm3j) * CapaciteCalorifiqueFluxMelange_UniteEnergieChoisie_kgjOUm3j)


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
NumeroErreurNonGeree = "44.29"

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

Private Sub calcul_bilan_masse_digestion(ByVal TempsSejourBouesDigesteur_j As Double)

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
Dim RatioTempsSejourSurReference As Double
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
NumeroErreurNonGeree = "44.21"
I = 0
j = 0
RatioTempsSejourSurReference = 0
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
RatioTempsSejourSurReference = TempsSejourBouesDigesteur_j / md_ReferenceTempsSejourBouesDigesteur_j
FacteurCorrectionTempsSejourBouesDigesteur = (md_CoefficientACalculFacteurCorrectionTempsSejourBouesDigesteur * RatioTempsSejourSurReference ^ md_CoefficientBCalculFacteurCorrectionTempsSejourBouesDigesteur) / (RatioTempsSejourSurReference ^ md_CoefficientBCalculFacteurCorrectionTempsSejourBouesDigesteur + md_CoefficientACalculFacteurCorrectionTempsSejourBouesDigesteur - 1)

'On calcule le flux en sortie de digestion
For j = 1 To nb_type_boues
    IntermediaireMS_kgj = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, j, mi_PARAMETRE_BOUES_MS_kgj)
    If IntermediaireMS_kgj <> 0 Then
        'On récupère les flux en entrée de digestion et calcule la réduction des MV
        IntermediaireMV_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, j, mi_PARAMETRE_BOUES_RATIO_MV_MES)
        IntermediaireDCO_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES)
        IntermediaireDBO_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES)
        IntermediaireNK_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, j, mi_PARAMETRE_BOUES_RATIO_NK_MES)
        IntermediaireP_kgj = IntermediaireMS_kgj * md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, j, mi_PARAMETRE_BOUES_RATIO_P_MES)
        IntermediaireDebitVolumique_m3j = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, j, mi_PARAMETRE_BOUES_Q_m3j)
        IntermediaireReductionMV_kgj = md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, j, mi_PARAMETRE_BOUES_REDUCTION_MV) * IntermediaireMV_kgj * FacteurCorrectionTempsSejourBouesDigesteur * F_facteur_correction_temperature_digesteur(md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C))
        'On calcule les flux en sortie de digestion (il sont potentiellement modifiés plus tard suivant les entrées de l'utilisateur)
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, j, mi_PARAMETRE_BOUES_MS_kgj) = IntermediaireMS_kgj - IntermediaireReductionMV_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, j, mi_PARAMETRE_BOUES_Q_m3j) = IntermediaireDebitVolumique_m3j - IntermediaireReductionMV_kgj / md_MasseVolumiqueBoues_kg_m3
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, j, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, j, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, j, mi_PARAMETRE_BOUES_Q_m3j)
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_MV_MES) = IntermediaireMV_kgj - IntermediaireReductionMV_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = IntermediaireReductionMV_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = IntermediaireDCO_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = IntermediaireDBO_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_NK_MES) = IntermediaireNK_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, j, mi_PARAMETRE_BOUES_RATIO_P_MES) = IntermediaireP_kgj
        'On ajoute également pour le global
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) + IntermediaireMS_kgj - IntermediaireReductionMV_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) + IntermediaireDebitVolumique_m3j - IntermediaireReductionMV_kgj / md_MasseVolumiqueBoues_kg_m3
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) + IntermediaireMV_kgj - IntermediaireReductionMV_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) + IntermediaireReductionMV_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DCO_MES) + IntermediaireDCO_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DBO_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_DBO_MES) + IntermediaireDBO_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_NK_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_NK_MES) + IntermediaireNK_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_P_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_P_MES) + IntermediaireP_kgj
    End If
Next j

'On calcule la part de MV dégradés pour chaque type de boues
For j = 1 To nb_type_boues
    md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, j, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, j, mi_PARAMETRE_BOUES_REDUCTION_MV) / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV)
Next j

'calcul de la réduction des MV
md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_REDUCTION_MV) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) / (md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) * md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj))

Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Private Sub forcage_rendement_MV()

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
NumeroErreurNonGeree = "44.33"
I = 0
DeltaReductionMV_kgj = 0


'On calcule l'écart de réduction des MV
DeltaReductionMV_kgj = (md_CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_REDUCTION_MV) / md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_REDUCTION_MV) - 1) * md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV)

'On recalcule les valeurs pour le mélange de boues
md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j) - DeltaReductionMV_kgj / md_MasseVolumiqueBoues_kg_m3
md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) - DeltaReductionMV_kgj
md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_REDUCTION_MV) + DeltaReductionMV_kgj
md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES) - DeltaReductionMV_kgj
md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)

'On recalcule les valeurs pour chaque type de boues
For I = 1 To nb_type_boues
    If md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) <> 0 Then
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_MS_kgj) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_MS_kgj) - md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_REDUCTION_MV) * DeltaReductionMV_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_RATIO_MV_MES) - md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_REDUCTION_MV) * DeltaReductionMV_kgj
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_Q_m3j) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_Q_m3j) - md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_REDUCTION_MV) * DeltaReductionMV_kgj / md_MasseVolumiqueBoues_kg_m3
        md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_CONCENTRATION_gL) = md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_MS_kgj) / md_FluxIntermediaires(mi_FLUX_BOUES_DIGEREES, I, mi_PARAMETRE_BOUES_Q_m3j)
    End If
Next I

md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_REDUCTION_MV) = md_CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_REDUCTION_MV)

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
'       qui se fait dans l'onglet "Feuil44" nommé "Biothelys_Exelys_LD" dans des plages de cellules dont les noms sont définis dans Excel
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim Ligne As Integer
Dim Colonne As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "44.10"
Ligne = 1
Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree


For Ligne = 1 To mi_NOMBRE_FLUX_ENTREE_BOUES
    Feuil44.Range("dig_biothelys_inlet_ratio").Cells(Ligne, Colonne) = ""
Next Ligne

For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil44.Range("dig_biothelys_inlet_boues").Cells(Ligne, Colonne) = ""
Next Ligne
For Ligne = 1 To mi_PARAMETRE_BOUES_TEMPERATURE_°C
    Feuil44.Range("dig_biothelys_outlet_boues").Cells(Ligne, Colonne) = ""
Next Ligne
For Ligne = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    Feuil44.Range("dig_biothelys_lyse_digesteur_pompage").Cells(Ligne, Colonne) = ""
Next Ligne
For Ligne = 1 To mi_NOMBRE_PARAMETRES_POMPAGE
    Feuil44.Range("dig_biothelys_outlet_pompage").Cells(Ligne, Colonne) = ""
Next Ligne
'digesteur
For Ligne = 1 To mi_NOMBRE_PARAMETRE_DIGESTION
    Feuil44.Range("dig_biothelys_digesteur").Cells(Ligne, Colonne) = ""
Next Ligne


Ligne = 1

Feuil44.Range("dig_biothelys_lyse_vapeur").Cells(Ligne, Colonne) = ""
Feuil44.Range("dig_biothelys_lyse_vapeur_pression").Cells(Ligne, Colonne) = ""
Feuil44.Range("dig_biothelys_vapeur_Teau_ini").Cells(1, Colonne) = ""
Feuil44.Range("dig_biothelys_vapeur_T_prechauff_eau").Cells(1, Colonne) = ""      'md_TemperaturePrechauffageEauPourProductionVapeur_°C

Feuil44.Range("digestionLD_BouesLysees_kgMSj").Cells(Ligne, Colonne) = ""       'md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj)
Feuil44.Range("digestionLD_BouesLysees_m3j").Cells(Ligne, Colonne) = ""       'md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
Feuil44.Range("digestionLD_BouesLysees_gL").Cells(Ligne, Colonne) = ""       'md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
Feuil44.Range("digestionLD_BouesLysees_RatioMVMES").Cells(Ligne, Colonne) = ""       'md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES)
Feuil44.Range("digestionLD_BouesLysees_°C").Cells(Ligne, Colonne) = ""       'md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
Feuil44.Range("digestionLD_BouesNonLysees_kgMSj").Cells(Ligne, Colonne) = ""       'md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_MS_kgj)
Feuil44.Range("digestionLD_BouesNonLysees_m3j").Cells(Ligne, Colonne) = ""       'md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_Q_m3j)
Feuil44.Range("digestionLD_BouesNonLysees_gL").Cells(Ligne, Colonne) = ""       'md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_CONCENTRATION_gL)
Feuil44.Range("digestionLD_BouesNonLysees_RatioMVMES").Cells(Ligne, Colonne) = ""       'md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_RATIO_MV_MES)
Feuil44.Range("digestionLD_BouesNonLysees_°C").Cells(Ligne, Colonne) = ""       'md_FluxIntermediaires(mi_FLUX_BOUES_NON_LYSEES, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
Feuil44.Range("digestionLD_DebitEauDilution_m3j").Cells(Ligne, Colonne) = ""       'md_DebitEauDilution_m3j
Feuil44.Range("digestionLD_TemperatureEauDilution_°C").Cells(Ligne, Colonne) = ""       'md_TemperatureEauDilution_°C
Feuil44.Range("digestionLD_TemperatureEauDilutionApresChauffage_°C").Cells(Ligne, Colonne) = ""       'md_TemperatureEauDilutionApresChauffage_°C
Feuil44.Range("dig_biothelys_BouesEntreeDigestion_kgMSj").Cells(Ligne, Colonne) = ""      'md_BouesEntreeDigestion_kgMSj
Feuil44.Range("dig_biothelys_BouesEntreeDigestion_m3j").Cells(Ligne, Colonne) = ""      'md_BouesEntreeDigestion_m3j
Feuil44.Range("dig_biothelys_BouesEntreeDigestion_gL").Cells(Ligne, Colonne) = ""      'md_BouesEntreeDigestion_gL
Feuil44.Range("dig_biothelys_md_BouesEntreeDigestionRatioMVSurMS").Cells(Ligne, Colonne) = ""     'md_BouesEntreeDigestionRatioMVSurMS
Feuil44.Range("dig_biothelys_md_BouesEntreeDigestionTemperature_°C").Cells(Ligne, Colonne) = ""     'md_FluxIntermediaires(mi_FLUX_DILUTION_DIGESTEUR, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C)
Feuil44.Range("dig_biothelys_biogaz_teneur_CH4").Cells(Ligne, Colonne) = ""
Feuil44.Range("dig_biothelys_ratio_biogaz_MV").Cells(Ligne, Colonne) = ""
Feuil44.Range("dig_biothelys_biogaz_Q_Nm3j").Cells(Ligne, Colonne) = ""
Feuil44.Range("dig_biothelys_energie_vapeur_kWhj").Cells(Ligne, Colonne) = ""
Feuil44.Range("dig_biothelys_energie_eau90").Cells(Ligne, Colonne) = ""
Feuil44.Range("digestionLD_EnergieChauffageDigesteur_kWhj").Cells(Ligne, Colonne) = ""
Feuil44.Range("dig_biothelys_energie_refroidissement_boues_kWhj").Cells(Ligne, Colonne) = ""
Feuil44.Range("dig_biothelys_dispo_eau90").Cells(Ligne, Colonne) = ""
Feuil44.Range("dig_biothelys_electricite").Cells(Ligne, Colonne) = ""


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
NumeroErreurNonGeree = "44.3"


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

'réduction des MV référence par type de boues (boues non lysées)   'modifié par DCA le 31/05/11 d'après une comparaison NDNc+PDN avec dim-digesteur
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


'température référence de digestion   en °C
md_TemperatureDigestionReference_°C = 38
md_TemperatureDigestionMinimale_°C = 35
md_TemperatureDigestionMaximale_°C = 42

md_EcartTemperatureBouesEchangeurDigesteur_°C = 3

'On attribue une production spécifique de biogaz par défaut qui sera ensuite calculée pour ne pas avoir de division par 0 lors des calculs de cocktail des boues en présence de coferments (1ère itération)
md_ProductionSpecifiqueBiogaz_Nm3_kgMV = 1

'température maximale de l'eau de dilution
md_TemperatureEauDilutionMaximale_°C = 90

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
md_RatioDCOSolubleSurDCOTotaleBouesDigerees = 0.08    'part de DCO soluble en sortie

'hypothèses sur les températures
md_TemperatureLyse_°C = 165 'DCA 28/08/13 d'après Julien CHAUZY, la température dans le réacteur est inférieure à la température de la vapeur saturée de façon à maintenir le tout liquide
md_RatioPertesVapeur = 0.05    '5% de pertes thermiques sur la vapeur de la lyse

'Concentration souhaitée des boues en entrée du digesteur
md_ConcentrationBouesEntreeDigestionSouhaitee_gL = 100  'DCA 28/08/13 vu avec Julien CHAUZY

'hypothèses sur les pertes du digesteur (en fonction du climat et de l'isolation du digesteur) en kWh/(m3.j)
md_PertesThermiquesDigestion_kWh_m3j(climat_froid, mi_DIGESTION_ISOLEE) = 0.6
md_PertesThermiquesDigestion_kWh_m3j(climat_froid, mi_DIGESTION_NON_ISOLEE) = 3.72
md_PertesThermiquesDigestion_kWh_m3j(climat_tempere, mi_DIGESTION_ISOLEE) = 0.36
md_PertesThermiquesDigestion_kWh_m3j(climat_tempere, mi_DIGESTION_NON_ISOLEE) = 2.2
md_PertesThermiquesDigestion_kWh_m3j(climat_chaud, mi_DIGESTION_ISOLEE) = 0.24
md_PertesThermiquesDigestion_kWh_m3j(climat_chaud, mi_DIGESTION_NON_ISOLEE) = 1.44

'efficacité de l'échangeur pour chauffer les boues
md_EfficaciteEchangeur = 0.9

'capacité calorifique des boues   en kWh/(m3.°C)
md_CapaciteCalorifiqueBoues_kWh_m3°C = 1.163
md_CapaciteCalorifiqueBoues_kcal_kg°C = 1
md_CapaciteCalorifiqueEau_kcal_kg°C = 1
md_TemperatureMaximaleBouesEntreeBiothelysValiditeCalculs_°C = 35   '°C  DCA 29/08/13 vu avec Julien CHAUZY, le bilan thermique tel que calculé dans OCEAN n'est pas valable au delà de cette température car on ne récupère pas autant de vapeur flash --> ratio de consommation de vapeur augmente

'pour calculer le vapeur à l'entrée en kWh/m3
md_CoefficientAConsommationSpecifiqueVapeur_kWh_m3Boues°C = -1.6202
md_CoefficientBConsommationSpecifiqueVapeur_kWh_m3Boues = 193.88

'masse volumique des boues (kg/m3)
md_MasseVolumiqueBoues_kg_m3 = 1000
md_MasseVolumiqueEau_kg_m3 = 1000

'coefficients nécessaires au calcul de la température de vapeur saturée à P   (en °C)
'    md_TemperatureSaturation_°C(P) = a1/(a2 - ln(a3*P)) + a4 - a5   avec P = pression absolue en bar
md_CoefficientsCalculTemperatureSaturationAPression(1) = 3816.44       '3816.27
md_CoefficientsCalculTemperatureSaturationAPression(2) = 18.3
md_CoefficientsCalculTemperatureSaturationAPression(3) = 750
md_CoefficientsCalculTemperatureSaturationAPression(4) = 46.13
md_CoefficientsCalculTemperatureSaturationAPression(5) = 273.15      '443.15

'coefficients pour calcul de l'enthalpie totale de la vapeur saturée à T   (en kcal/kg)  Hvap_sat(T) = a0 + a1*T
md_CoefficientsCalculEnthalpieVaporisationATemperature(0) = 616.87
md_CoefficientsCalculEnthalpieVaporisationATemperature(1) = -0.7542

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
NumeroErreurNonGeree = "44.4"
I = 0
j = 0

'températures pour la digestion
If mn_CaracteristiquesDigestion_force(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = True Then
    If md_CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) < md_TemperatureDigestionMinimale_°C Or md_CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) > md_TemperatureDigestionMaximale_°C Then
        MsgBox Feuil44.Range("digestionLD_erreur_TemperatureDigestion") & ": " & md_TemperatureDigestionMinimale_°C & " - " & md_TemperatureDigestionMaximale_°C & " °C", Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
        pn_EffacementResultatsEnCours = True
        pn_ErreurNonGeree = True
        Exit Sub
    Else
        md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = md_CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C)
    End If
Else
    md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPERATURE_°C) = md_TemperatureDigestionReference_°C
End If

md_ReferenceTempsSejourBouesDigesteur_j = 15

'température des boues en entrée de digesteur
If mn_TemperatureBouesEntreeLyse_°C_force = True Then
    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureBouesEntreeLyse_°C_f
Else
    md_FluxIntermediaires(mi_FLUX_ENTREE_LYSE, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = pd_TemperatureEauConditionsReelles_°C
End If

'SRT en j pour le design et initialisation du SRT réel
If mn_CaracteristiquesDigestion_force(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j) = True Then
    md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j) = md_CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j)
Else
    md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_DESIGN_j) = 15
End If
md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_TEMPS_SEJOUR_REEL_j) = 15

'paramètres des digesteurs intervenant dans l'agitation
If mi_ChoixTypeAgitationDigesteur = mi_AGITATION_BIOGAZ Then
    If mn_CaracteristiquesDigestion_force(mi_PARAMETRE_DIGESTION_HAUTEUR_m) = True Then
        md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_HAUTEUR_m) = md_CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_HAUTEUR_m)
    Else
        md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_HAUTEUR_m) = 10
    End If
ElseIf mi_ChoixTypeAgitationDigesteur = mi_AGITATION_RECIRCULATION Then
    If mn_CaracteristiquesDigestion_force(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) = True Then
        md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) = md_CaracteristiquesDigestion_f(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j)
    Else
        md_CaracteristiquesDigestion(mi_PARAMETRE_DIGESTION_AGITATION_TAUX_RECIRCULATION_VDigesteur_j) = 5  'ici c'est un taux par rapport au volume total de digesteur, le vrai taux sera recalculé dans conso élec
    End If
End If

'température de l'eau utilisée pour produire la vapeur
If mn_TemperatureInitialeEauPourProductionVapeur_°C_force = True Then
    md_TemperatureInitialeEauPourProductionVapeur_°C = md_TemperatureInitialeEauPourProductionVapeur_°C_f
Else
    md_TemperatureInitialeEauPourProductionVapeur_°C = 15
End If
If mn_TemperaturePrechauffageEauPourProductionVapeur_°C_force = True Then
    md_TemperaturePrechauffageEauPourProductionVapeur_°C = md_TemperaturePrechauffageEauPourProductionVapeur_°C_f
Else
    md_TemperaturePrechauffageEauPourProductionVapeur_°C = 85
End If

'pression de la vapeur à utiliser   (en bars)
If mn_PressionAbsolueVapeurSaturanteLyse_bar_force = True Then
    md_PressionAbsolueVapeurSaturanteLyse_bar = md_PressionAbsolueVapeurSaturanteLyse_bar_f
Else
    md_PressionAbsolueVapeurSaturanteLyse_bar = 12.5
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

'Température en sortie de lyse
If mn_TemperatureSortieLyse_force = True Then
    md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureSortieLyse_f
Else
    If mi_ChoixConfiguration = pi_CONFIGURATION_BIOTHELYS Then   'DCA vu avec Julien CHAUZY le 28/08/13
        md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = 120    '°C
    ElseIf mi_ChoixConfiguration = pi_CONFIGURATION_EXELYS Then
        md_FluxIntermediaires(mi_FLUX_LYSE_ECHANGEUR55, mi_MELANGE_BOUES, mi_PARAMETRE_BOUES_TEMPERATURE_°C) = md_TemperatureLyse_°C
    End If
End If

'convoyage/pompage des boues
md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = 2
md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_HMT_mCE) = 50   'mCE
md_CaracteristiquePompageBoues(mi_POMPAGE_LYSE_DIGESTEUR, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = 24   'h/j
md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_NOMBRE_POMPES) = 2
md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_HMT_mCE) = 50   'mCE
md_CaracteristiquePompageBoues(mi_POMPAGE_BOUES_DIGEREES, mi_PARAMETRE_POMPAGE_TEMPS_FONCTIONNEMENT_hj) = 24   'h/j
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

Private Sub Class_Terminate()

    If pn_EffacementResultatsEnCours = False Then
        Call ecriture_resultats
        Call incrementation_commentaires_calculs
    End If

End Sub









