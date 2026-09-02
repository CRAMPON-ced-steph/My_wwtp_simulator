Attribute VB_Name = "z_Biolix_graisses"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'constantes pour le choix du type de surpresseur
Private Const mi_NOMBRE_TECHNOLOGIES_SURPRESSEUR As Integer = 3
Private Const mi_SURPRESSEUR_TRILOBES As Integer = 1
Private Const mi_SURPRESSEUR_VIS As Integer = 2
Private Const mi_SURPRESSEUR_CENTRIFUGE As Integer = 3

'constantes pour le choix de la cuve d'homogénéisation
Private Const mi_AVEC_AGITATION_CUVE As Integer = 1

Private Const mi_CRITERE_INCOHERENCE As Double = 0.1

'choix de l'onglet synthèse
Private mi_ChoixTypeSurpresseur As Integer
Private mi_ChoixAgitationCuveHomogeneisation As Double
Private mi_ChoixAgitationBassinBiologique As Double
Private md_EntreeGraissesPourcentageFluxAdmis As Double

'variables échangées avec l'utilisateur
Private md_FluxDCOGraissesEntreeDimensionnement_kgj As Double
Private md_FluxDCOGraissesEntreeExploitation_kgj As Double
Private md_RendementEliminationDCO As Double
Private md_ConsommationAzote_kgNj As Double
Private md_ConsommationPhosphore_kgPj As Double
Private md_HauteurBassinAeration_m As Double
Private md_BesoinOxygene_kgO2j As Double
Private md_DebitAirProcess_Nm3j As Double
Private md_RendementTransfertOxygeneMembraneEauClaire_1m As Double
Private md_PressionRefoulementSurpresseur_mCE As Double
Private md_VolumeCuveHomogeneisation_m3 As Double
Private md_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3 As Double
Private md_VolumeBassinBiologique_m3 As Double
Private md_PuissanceSpecifiqueAgitationBassinBiologique_W_m3 As Double
Private md_TempsFonctionnementAgitationBassinBiologique_hj As Double
Private md_BouesSortieConcentrationMES_gL As Double
Private md_BouesSortieFluxMES_kgj As Double
Private md_BouesSortieDebitVolumique_m3j As Double
Private md_BouesSortieRatioMVSurMES As Double
Private md_PartBouesBiolixVersTraitementSecondaireEau As Double
Private mi_PompageBouesBiolixNombrePompes As Integer
Private md_PompageBouesBiolixPressionRefoulement_mCE As Double
Private md_PompageBouesBiolixTempsFonctionnement_hj As Double
Private md_PompageBouesBiolixDebitUnitaire_m3h As Double
Private md_PompageBouesBiolixRendementGlobalPompes As Double
Public pd_ConsommationElectrique_kWhj As Double

Public pd_ConsommationElectriqueFixe_kWhj As Double

'Variables forcées
Private mn_FluxDCOGraissesEntreeDimensionnement_kgj_force As Boolean
Private mn_RendementEliminationDCO_force As Boolean
Private mn_ConsommationPhosphore_kgPj_force As Boolean
Private mn_ConsommationAzote_kgNj_force As Boolean
Private mn_HauteurBassinAeration_m_force As Boolean
Private mn_BesoinOxygene_kgO2j_force As Boolean
Private mn_DebitAirProcess_Nm3j_force As Boolean
Private mn_RendementTransfertOxygeneMembraneEauClaire_1m_force As Boolean
Private mn_PressionRefoulementSurpresseur_mCE_force As Boolean
Private mn_VolumeCuveHomogeneisation_m3_force As Boolean
Private mn_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3_force As Boolean
Private mn_VolumeBassinBiologique_m3_force As Boolean
Private mn_PuissanceSpecifiqueAgitationBassinBiologique_W_m3_force As Boolean
Private mn_TempsFonctionnementAgitationBassinBiologique_hj_force As Boolean
Private mn_BouesSortieConcentrationMES_gL_force As Boolean
Private mn_BouesSortieFluxMES_kgj_force As Boolean
Private mn_BouesSortieRatioMVSurMES_force As Boolean
Private mn_PartBouesBiolixVersTraitementSecondaireEau_force As Boolean
Private mn_PompageBouesBiolixNombrePompes_force As Boolean
Private mn_PompageBouesBiolixPressionRefoulement_mCE_force As Boolean
Private mn_PompageBouesBiolixTempsFonctionnement_hj_force As Boolean
Private mn_PompageBouesBiolixDebitUnitaire_m3h_force As Boolean
Private mn_PompageBouesBiolixRendementGlobalPompes_force As Boolean

'Variables forcées
Private md_FluxDCOGraissesEntreeDimensionnement_kgj_f As Double
Private md_RendementEliminationDCO_f As Double
Private md_ConsommationPhosphore_kgPj_f As Double
Private md_ConsommationAzote_kgNj_f As Double
Private md_HauteurBassinAeration_m_f As Double
Private md_BesoinOxygene_kgO2j_f As Double
Private md_DebitAirProcess_Nm3j_f As Double
Private md_RendementTransfertOxygeneMembraneEauClaire_1m_f As Double
Private md_PressionRefoulementSurpresseur_mCE_f As Double
Private md_VolumeCuveHomogeneisation_m3_f As Double
Private md_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3_f As Double
Private md_VolumeBassinBiologique_m3_f As Double
Private md_PuissanceSpecifiqueAgitationBassinBiologique_W_m3_f As Double
Private md_TempsFonctionnementAgitationBassinBiologique_hj_f As Double
Private md_BouesSortieConcentrationMES_gL_f As Double
Private md_BouesSortieFluxMES_kgj_f As Double
Private md_BouesSortieRatioMVSurMES_f As Double
Private md_PartBouesBiolixVersTraitementSecondaireEau_f As Double
Private mi_PompageBouesBiolixNombrePompes_f As Integer
Private md_PompageBouesBiolixPressionRefoulement_mCE_f As Double
Private md_PompageBouesBiolixTempsFonctionnement_hj_f As Double
Private md_PompageBouesBiolixDebitUnitaire_m3h_f As Double
Private md_PompageBouesBiolixRendementGlobalPompes_f As Double

'hypothèses
Private md_BesoinSpecifiqueO2EliminationDCO_kgO2_kgDCO As Double
Private md_AerationFacteurK As Double
Private md_TempsFonctionnementAgitationCuveHomogeneisation_hj As Double
Private md_TempsPassageHydrauliqueCuveHomogeneisation_j As Double
Private md_TempsPassageHydrauliqueBassinBiologique_j As Double
Private md_BesoinSpecifiquePhophoreBacteries_kgPj As Double
Private md_BesoinSpecifiqueAzoteBacteries_kgNj As Double
Private md_HauteurImmersionDiffuseur_m As Double
Private md_ConsommationElectriqueSpecifiqueSurpresseurs_Wh_Nm3mCE As Double
Private md_ProductionSpecifiqueBoues_kgMVS_kgDCOeliminee As Double
Private md_EntreeDilueeConcentrationDCO_gL As Double

'erreurs
Private Const mi_NOMBRE_COMMENTAIRES_CALCULS As Integer = 1
Private ms_CommentairesCalculs() As String
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
NumeroErreurNonGeree = "47.8"
I = 0

'Ajout du nom du procédé
ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil47.Range("biolix_erreur_nom")

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

Private Sub calcul_cocktail_boues_entrees(ByRef TableauRecapitulatifFluxBoues() As Double, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double, ByRef EntreeGraissesMES_kgj As Double, ByRef EntreeGraissesMV_kgj As Double, ByRef EntreeGraissesDebitVolumique_m3j As Double, ByRef EntreeGraissesDCO_kgj As Double, ByRef EntreeGraissesDBO_kgj As Double, ByRef EntreeGraissesNK_kgj As Double, ByRef EntreeGraissesPt_kgj As Double)


'************************************************************************************************************************************************************************
'OBJECTIF:

'VARIABLES D'ENTREE:
'
'VARIABLES DE SORTIE:
'************************************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "47.22"
I = 0

'Qualité des graisses admises dans le biolix (particulaire)
EntreeGraissesMES_kgj = TableauRecapitulatifFluxBoues(boues_inlet, graisses, repere_MES) * md_EntreeGraissesPourcentageFluxAdmis
EntreeGraissesMV_kgj = TableauRecapitulatifFluxBoues(boues_inlet, graisses, repere_MV_MES) * EntreeGraissesMES_kgj
EntreeGraissesDebitVolumique_m3j = TableauRecapitulatifFluxBoues(boues_inlet, graisses, repere_Q) * md_EntreeGraissesPourcentageFluxAdmis
EntreeGraissesDCO_kgj = TableauRecapitulatifFluxBoues(boues_inlet, graisses, repere_ratio_DCO_MES) * EntreeGraissesMES_kgj
EntreeGraissesDBO_kgj = TableauRecapitulatifFluxBoues(boues_inlet, graisses, repere_ratio_DBO_MES) * EntreeGraissesMES_kgj
EntreeGraissesNK_kgj = TableauRecapitulatifFluxBoues(boues_inlet, graisses, repere_ratio_NK_MES) * EntreeGraissesMES_kgj
EntreeGraissesPt_kgj = TableauRecapitulatifFluxBoues(boues_inlet, graisses, repere_ratio_Pt_MES) * EntreeGraissesMES_kgj

'Qualité des graisses admises dans le biolix (soluble)
EntreeGraissesDCO_kgj = EntreeGraissesDCO_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_inlet, graisses, repere_mgL_DCO) * EntreeGraissesDebitVolumique_m3j / 1000
EntreeGraissesNK_kgj = EntreeGraissesNK_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_inlet, graisses, repere_mgL_NK) * EntreeGraissesDebitVolumique_m3j / 1000
EntreeGraissesPt_kgj = EntreeGraissesPt_kgj + TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_inlet, graisses, repere_mgL_Pt) * EntreeGraissesDebitVolumique_m3j / 1000
For I = 1 To nb_repere_mgL
    TableauRecapitulatifPollutionSolubleDansFluxBoues(boues_graisses_biolix, graisses, I) = 0
Next I

'je calcule les paramètres des boues épaissies
TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_flux_in) = TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_flux_in) + TableauRecapitulatifFluxBoues(boues_inlet, graisses, repere_flux_in) * md_EntreeGraissesPourcentageFluxAdmis
TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_verif_flux) = TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_flux_in)

'je retranche la part de flux d'entrée utilisée (objectif: vérif des bilans matières)
TableauRecapitulatifFluxBoues(boues_inlet, graisses, repere_verif_flux) = TableauRecapitulatifFluxBoues(boues_inlet, graisses, repere_verif_flux) - TableauRecapitulatifFluxBoues(boues_inlet, graisses, repere_flux_in) * md_EntreeGraissesPourcentageFluxAdmis

'On regarde si le flux n'est pas nul
If EntreeGraissesMES_kgj = 0 Or EntreeGraissesDebitVolumique_m3j = 0 Then
    MsgBox Feuil5.Range("erreur_choix_biolix_inlet"), Title:=Feuil2.Range("synthese_nom_step" & pi_FiliereConsideree)
    pn_EffacementResultatsEnCours = True
    pn_ErreurNonGeree = True
    Exit Sub
End If


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub


Private Sub calcul_aeration_et_nutriments(ByVal ReperageBoucleReel As Boolean, ByVal EntreeGraissesDCO_kgj As Double)


'************************************************************************************************************************************************************************
'OBJECTIF:

'VARIABLES D'ENTREE:
'
'VARIABLES DE SORTIE:
'************************************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "47.20"

'Calcul des besoins en O2 et du débit d'air nécessaire
If ReperageBoucleReel = True And mn_BesoinOxygene_kgO2j_force = True Then
    md_BesoinOxygene_kgO2j = md_BesoinOxygene_kgO2j_f
Else
    md_BesoinOxygene_kgO2j = md_RendementEliminationDCO * md_BesoinSpecifiqueO2EliminationDCO_kgO2_kgDCO * EntreeGraissesDCO_kgj
End If
If ReperageBoucleReel = True And mn_DebitAirProcess_Nm3j_force = True Then
    md_DebitAirProcess_Nm3j = md_DebitAirProcess_Nm3j_f
Else
    md_DebitAirProcess_Nm3j = md_BesoinOxygene_kgO2j / (md_AerationFacteurK * md_RendementTransfertOxygeneMembraneEauClaire_1m / 100 * (md_HauteurBassinAeration_m - md_HauteurImmersionDiffuseur_m) * pd_QUANTITE_OXYGENE_DANS_AIR_kgO2_Nm3)
End If

'Calcul de la consommation en nutriments
If ReperageBoucleReel = True And mn_ConsommationPhosphore_kgPj_force = True Then
    md_ConsommationPhosphore_kgPj = md_ConsommationPhosphore_kgPj_f
Else
    md_ConsommationPhosphore_kgPj = md_BesoinSpecifiquePhophoreBacteries_kgPj * EntreeGraissesDCO_kgj
End If
If mn_ConsommationAzote_kgNj_force = True Then
    md_ConsommationAzote_kgNj = md_ConsommationAzote_kgNj_f
Else
    md_ConsommationAzote_kgNj = md_BesoinSpecifiqueAzoteBacteries_kgNj * EntreeGraissesDCO_kgj
End If


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True

End Sub

Private Sub calcul_boues_produites(ByVal ReperageBoucleReel As Boolean, ByVal EntreeGraissesDCO_kgj As Double, ByVal EntreeGraissesDBO_kgj As Double, ByVal EntreeGraissesNK_kgj As Double, ByVal EntreeGraissesPt_kgj As Double, ByRef TableauRecapitulatifFluxBoues() As Double)


'************************************************************************************************************************************************************************
'OBJECTIF:

'VARIABLES D'ENTREE:
'
'VARIABLES DE SORTIE:
'************************************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "47.21"

'On calcule la production de boues
If mn_BouesSortieFluxMES_kgj_force = True And ReperageBoucleReel = True Then
    md_BouesSortieFluxMES_kgj = md_BouesSortieFluxMES_kgj_f
Else
    md_BouesSortieFluxMES_kgj = (md_ProductionSpecifiqueBoues_kgMVS_kgDCOeliminee * (md_RendementEliminationDCO * EntreeGraissesDCO_kgj)) / md_BouesSortieRatioMVSurMES
End If

'Concentration des boues en sortie
If mn_BouesSortieConcentrationMES_gL_force = True Then
    md_BouesSortieConcentrationMES_gL = md_BouesSortieConcentrationMES_gL_f
    md_BouesSortieDebitVolumique_m3j = md_BouesSortieFluxMES_kgj / md_BouesSortieConcentrationMES_gL
Else
    md_BouesSortieConcentrationMES_gL = md_BouesSortieFluxMES_kgj / md_BouesSortieDebitVolumique_m3j
End If

'On remplit le tableau des boues
TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_Q) = md_BouesSortieDebitVolumique_m3j
TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_MES) = md_BouesSortieFluxMES_kgj
TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_MV_MES) = md_BouesSortieFluxMES_kgj * md_BouesSortieRatioMVSurMES
TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_ratio_DCO_MES) = ((1 - md_RendementEliminationDCO) * EntreeGraissesDCO_kgj) / md_BouesSortieFluxMES_kgj
TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_ratio_DBO_MES) = ((1 - md_RendementEliminationDCO) * EntreeGraissesDBO_kgj) / md_BouesSortieFluxMES_kgj
TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_ratio_NK_MES) = (EntreeGraissesNK_kgj + md_ConsommationAzote_kgNj) / md_BouesSortieFluxMES_kgj
TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_ratio_Pt_MES) = (EntreeGraissesPt_kgj + md_ConsommationPhosphore_kgPj) / md_BouesSortieFluxMES_kgj

'On gère l'utilisation des boues de biolix
md_PartBouesBiolixVersTraitementSecondaireEau = 1
If choix_secondaire = False Then
    
End If
TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_verif_flux) = TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_verif_flux) - TableauRecapitulatifFluxBoues(boues_graisses_biolix, graisses, repere_flux_in) * md_PartBouesBiolixVersTraitementSecondaireEau


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
Dim j As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "47.3"
j = 0

'Concentration DCO en entrée
md_EntreeDilueeConcentrationDCO_gL = 75   'NP 3NP60 p21/40

'Besoins spécifiques en O2 pour l'élimination biologique de la DCO
md_BesoinSpecifiqueO2EliminationDCO_kgO2_kgDCO = 1    'kg O2/kg DCO

'Facteur K pour le transfert d'oxygène
md_AerationFacteurK = (0.3 + 0.45) / 2     'Moyenne valeurs ci-après -->   Note procédé 3NP60 p.23/40 = 0.3 / Tableur de dim = 0,45

'Temps de fonctionnement de l'agitation de la cuve d'homogénéisation
md_TempsFonctionnementAgitationCuveHomogeneisation_hj = 24    'h/j

'Temps de passage (hydraulique) des graisses dans la cuve d'homogénéisation
md_TempsPassageHydrauliqueCuveHomogeneisation_j = 1       'jour   Note procédé 3NP60 p.19/40

'Temps de passage (hydraulique) des graisses dans le bassin biologique
md_TempsPassageHydrauliqueBassinBiologique_j = 20        'jour   Note procédé 3NP60 p.21/40

'Besoins en nutriments pour les bactéries
md_BesoinSpecifiquePhophoreBacteries_kgPj = 0.005   'g P/g DCO
md_BesoinSpecifiqueAzoteBacteries_kgNj = 0.025     'g N/g DCO

'Hauteur d'immersion des diffuseurs
md_HauteurImmersionDiffuseur_m = 0.25    'm

'Consommation électrique spécifique des surpresseurs      en Wh/(Nm3.mCE)
If mi_ChoixTypeSurpresseur = mi_SURPRESSEUR_TRILOBES Then
    md_ConsommationElectriqueSpecifiqueSurpresseurs_Wh_Nm3mCE = 4.5
ElseIf mi_ChoixTypeSurpresseur = mi_SURPRESSEUR_VIS Then
    md_ConsommationElectriqueSpecifiqueSurpresseurs_Wh_Nm3mCE = 3.8
ElseIf mi_ChoixTypeSurpresseur = mi_SURPRESSEUR_CENTRIFUGE Then
    md_ConsommationElectriqueSpecifiqueSurpresseurs_Wh_Nm3mCE = 3.5
End If

'Production de boues
md_ProductionSpecifiqueBoues_kgMVS_kgDCOeliminee = 0.3


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
NumeroErreurNonGeree = "47.4"


'Rendement d'élimination de la DCO
If mn_RendementEliminationDCO_force = True Then
    md_RendementEliminationDCO = md_RendementEliminationDCO_f
Else
    md_RendementEliminationDCO = 0.7   '70%   Note procédé 3NP60 p.15/40
End If

'Pourcentage de MV/MES des boues produites
If mn_BouesSortieRatioMVSurMES_force = True Then
    md_BouesSortieRatioMVSurMES = md_BouesSortieRatioMVSurMES_f
Else
    md_BouesSortieRatioMVSurMES = 0.76 'fndae 24 page 32
End If

'Hauteur d'eau dans le bassin d'aération
If mn_HauteurBassinAeration_m_force = True Then
    md_HauteurBassinAeration_m = md_HauteurBassinAeration_m_f
Else
    md_HauteurBassinAeration_m = 6  'm
End If

'Pression de refoulement des surpresseurs
If mn_PressionRefoulementSurpresseur_mCE_force = True Then
    md_PressionRefoulementSurpresseur_mCE = md_PressionRefoulementSurpresseur_mCE_f
Else
    md_PressionRefoulementSurpresseur_mCE = md_HauteurBassinAeration_m + 1 'supposition 100 mbar de plus que la hauteur d'eau dans les bassins
End If

'Rendement de transfert de l'oxygène en eau claire
If mn_RendementTransfertOxygeneMembraneEauClaire_1m_force = True Then
    md_RendementTransfertOxygeneMembraneEauClaire_1m = md_RendementTransfertOxygeneMembraneEauClaire_1m_f
Else
    md_RendementTransfertOxygeneMembraneEauClaire_1m = (5.25 + 4) / 2     '%/m   Moyenne des valeurs ci-après -->   Note procédé 3NP60 p.23/40: 5.25%/m / Tableur de dim = 4 %/m
End If

'Durée de fonctionnement de l'agitation du bassin biologique
If mn_TempsFonctionnementAgitationBassinBiologique_hj_force = True Then
    md_TempsFonctionnementAgitationBassinBiologique_hj = md_TempsFonctionnementAgitationBassinBiologique_hj_f
Else
    md_TempsFonctionnementAgitationBassinBiologique_hj = 24  'h/j   A priori, l'air est suffisant pour le brassage sauf si le biolix est sous-chargé
End If

'Puissance spécifique d'agitation de la cuve d'homogénéisation
If mn_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3_force = True Then
    md_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3 = md_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3_f
Else
    md_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3 = 25   'W/m3    entre 20 et 30 selon Note procédé 3NP60 p.19/40: 20 à 30 W/m3
End If

'Puissance spécifique d'agitationj du bassin biologique
If mn_PuissanceSpecifiqueAgitationBassinBiologique_W_m3_force = True Then
    md_PuissanceSpecifiqueAgitationBassinBiologique_W_m3 = md_PuissanceSpecifiqueAgitationBassinBiologique_W_m3_f
Else
    md_PuissanceSpecifiqueAgitationBassinBiologique_W_m3 = 25   'W/m3    entre 20 et 30 selon Note procédé 3NP60 p.19/40: 20 à 30 W/m3
End If

'Données sur le pompage des boues de biolix
If mn_PompageBouesBiolixNombrePompes_force = True Then
    mi_PompageBouesBiolixNombrePompes = mi_PompageBouesBiolixNombrePompes_f
Else
    mi_PompageBouesBiolixNombrePompes = 2
End If
If mn_PompageBouesBiolixPressionRefoulement_mCE_force = True Then
    md_PompageBouesBiolixPressionRefoulement_mCE = md_PompageBouesBiolixPressionRefoulement_mCE_f
Else
    md_PompageBouesBiolixPressionRefoulement_mCE = 5
End If
If mn_PompageBouesBiolixTempsFonctionnement_hj_force = True Then
    md_PompageBouesBiolixTempsFonctionnement_hj = md_PompageBouesBiolixTempsFonctionnement_hj_f
Else
    md_PompageBouesBiolixTempsFonctionnement_hj = 6
End If


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True


End Sub




Public Sub dimensionnement(ByRef TableauRecapitulatifFluxBoues() As Double, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double)


'*********************************************************************************************************************************************************
'OBJECTIF:

'VARIABLES D'ENTREE:

'VARIABLES DE SORTIE:
'*********************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim EntreeGraissesMES_kgj As Double
Dim EntreeGraissesMV_kgj As Double
Dim EntreeGraissesDCO_kgj As Double
Dim EntreeGraissesDBO_kgj As Double
Dim EntreeGraissesNK_kgj As Double
Dim EntreeGraissesPt_kgj As Double
Dim EntreeGraissesDebitVolumique_m3j As Double
Dim ReperageBoucleReel As Boolean

'Initialisation des variables locales
NumeroErreurNonGeree = "47.5"
EntreeGraissesMES_kgj = 0
EntreeGraissesMV_kgj = 0
EntreeGraissesDCO_kgj = 0
EntreeGraissesDBO_kgj = 0
EntreeGraissesNK_kgj = 0
EntreeGraissesPt_kgj = 0
EntreeGraissesDebitVolumique_m3j = 0
ReperageBoucleReel = False

'On calcule les graisses en entrée
Call calcul_cocktail_boues_entrees(TableauRecapitulatifFluxBoues, TableauRecapitulatifPollutionSolubleDansFluxBoues, EntreeGraissesMES_kgj, EntreeGraissesMV_kgj, EntreeGraissesDebitVolumique_m3j, EntreeGraissesDCO_kgj, EntreeGraissesDBO_kgj, EntreeGraissesNK_kgj, EntreeGraissesPt_kgj)

'On regarde si la valeur de DCO utilisée pour le dimensionnement a été forcée et on calcule le débit volumique de graisses après dilution et donc des boues de sortie
If mn_FluxDCOGraissesEntreeDimensionnement_kgj_force = True Then
    md_BouesSortieDebitVolumique_m3j = md_FluxDCOGraissesEntreeDimensionnement_kgj_f / md_EntreeDilueeConcentrationDCO_gL
    md_FluxDCOGraissesEntreeDimensionnement_kgj = md_FluxDCOGraissesEntreeDimensionnement_kgj_f
    EntreeGraissesDCO_kgj = md_FluxDCOGraissesEntreeDimensionnement_kgj
Else
    md_FluxDCOGraissesEntreeDimensionnement_kgj = EntreeGraissesDCO_kgj
    md_BouesSortieDebitVolumique_m3j = md_FluxDCOGraissesEntreeDimensionnement_kgj / md_EntreeDilueeConcentrationDCO_gL
End If
'On vérifie la cohérence des débits d'entrée et de sortie
If md_BouesSortieDebitVolumique_m3j < EntreeGraissesDebitVolumique_m3j Then
    md_BouesSortieDebitVolumique_m3j = EntreeGraissesDebitVolumique_m3j
End If

'Dimensionnement des cuves (homogénéisation et bassin biologique)
If mi_ChoixAgitationCuveHomogeneisation = mi_AVEC_AGITATION_CUVE Then
    If mn_VolumeCuveHomogeneisation_m3_force = True Then
        md_VolumeCuveHomogeneisation_m3 = md_VolumeCuveHomogeneisation_m3_f
    Else
        md_VolumeCuveHomogeneisation_m3 = md_BouesSortieDebitVolumique_m3j * md_TempsPassageHydrauliqueCuveHomogeneisation_j
    End If
End If
If mi_ChoixAgitationBassinBiologique = mi_AVEC_AGITATION_CUVE Then
    If mn_VolumeBassinBiologique_m3_force = True Then
        md_VolumeBassinBiologique_m3 = md_VolumeBassinBiologique_m3_f
    Else
        md_VolumeBassinBiologique_m3 = md_BouesSortieDebitVolumique_m3j * md_TempsPassageHydrauliqueBassinBiologique_j
    End If
End If

'Calcul de l'aération et de la consommation de nutriments
Call calcul_aeration_et_nutriments(ReperageBoucleReel, EntreeGraissesDCO_kgj)

'Calcul des boues en sortie
Call calcul_boues_produites(ReperageBoucleReel, EntreeGraissesDCO_kgj, EntreeGraissesDBO_kgj, EntreeGraissesNK_kgj, EntreeGraissesPt_kgj, TableauRecapitulatifFluxBoues)


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

'VARIABLES DE SORTIE:
'   pd_ConsommationElectrique_kWhj: consommation électrique journalière du procédé
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim ConsommationElectriqueAeration_kWhj As Double
Dim ConsommationElectriqueAgitationCuveHomogeneisation_kWhj As Double
Dim ConsommationElectriqueAgitationBassinBiologique_kWhj As Double
Dim PuissancePompe_kW As Double
Dim IncoherenceCalcul As Double
Dim RendementPompe As Double
Dim RendementGlobalPompe As Double
Dim ConsommationElectriquePompageBoues_kWhj As Double

'Initialisation des variables locales
NumeroErreurNonGeree = "47.7"
ConsommationElectriqueAeration_kWhj = 0
ConsommationElectriqueAgitationCuveHomogeneisation_kWhj = 0
ConsommationElectriqueAgitationBassinBiologique_kWhj = 0
PuissancePompe_kW = 0
IncoherenceCalcul = 0
RendementPompe = 0
RendementGlobalPompe = 0
ConsommationElectriquePompageBoues_kWhj = 0

'consommation électrique de l'aération
ConsommationElectriqueAeration_kWhj = md_DebitAirProcess_Nm3j * md_PressionRefoulementSurpresseur_mCE * md_ConsommationElectriqueSpecifiqueSurpresseurs_Wh_Nm3mCE / 1000

'Consommation électrique pour l'agitation des cuves
If mi_ChoixAgitationCuveHomogeneisation = mi_AVEC_AGITATION_CUVE Then
    ConsommationElectriqueAgitationCuveHomogeneisation_kWhj = md_TempsFonctionnementAgitationCuveHomogeneisation_hj * md_VolumeCuveHomogeneisation_m3 * md_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3 / 1000
End If

If mi_ChoixAgitationBassinBiologique = mi_AVEC_AGITATION_CUVE Then
    ConsommationElectriqueAgitationBassinBiologique_kWhj = md_TempsFonctionnementAgitationBassinBiologique_hj * md_VolumeBassinBiologique_m3 * md_PuissanceSpecifiqueAgitationBassinBiologique_W_m3 / 1000
End If

'Consommation électrique du pompage des boues
If mn_PompageBouesBiolixDebitUnitaire_m3h_force = True Then
    md_PompageBouesBiolixDebitUnitaire_m3h = md_PompageBouesBiolixDebitUnitaire_m3h_f
    If mn_PompageBouesBiolixTempsFonctionnement_hj_force = True Then
        md_PompageBouesBiolixTempsFonctionnement_hj = md_PompageBouesBiolixTempsFonctionnement_hj_f
        If mn_PompageBouesBiolixNombrePompes_force = False Then
            If Int(md_BouesSortieDebitVolumique_m3j / (md_PompageBouesBiolixTempsFonctionnement_hj * md_PompageBouesBiolixDebitUnitaire_m3h)) < (md_BouesSortieDebitVolumique_m3j / (md_PompageBouesBiolixTempsFonctionnement_hj * md_PompageBouesBiolixDebitUnitaire_m3h)) Then
                mi_PompageBouesBiolixNombrePompes = Int(md_BouesSortieDebitVolumique_m3j / (md_PompageBouesBiolixTempsFonctionnement_hj * md_PompageBouesBiolixDebitUnitaire_m3h)) + 1
            Else
                mi_PompageBouesBiolixNombrePompes = Int(md_BouesSortieDebitVolumique_m3j / (md_PompageBouesBiolixTempsFonctionnement_hj * md_PompageBouesBiolixDebitUnitaire_m3h))
            End If
        End If
    Else
        md_PompageBouesBiolixTempsFonctionnement_hj = md_BouesSortieDebitVolumique_m3j / (mi_PompageBouesBiolixNombrePompes * md_PompageBouesBiolixDebitUnitaire_m3h)
        If md_PompageBouesBiolixTempsFonctionnement_hj > pd_NOMBRE_HEURE_PAR_JOUR Then
            md_PompageBouesBiolixTempsFonctionnement_hj = pd_NOMBRE_HEURE_PAR_JOUR
            If mn_PompageBouesBiolixNombrePompes_force = False Then
                If Int(md_BouesSortieDebitVolumique_m3j / (md_PompageBouesBiolixTempsFonctionnement_hj * md_PompageBouesBiolixDebitUnitaire_m3h)) < (md_BouesSortieDebitVolumique_m3j / (md_PompageBouesBiolixTempsFonctionnement_hj * md_PompageBouesBiolixDebitUnitaire_m3h)) Then
                    mi_PompageBouesBiolixNombrePompes = Int(md_BouesSortieDebitVolumique_m3j / (md_PompageBouesBiolixTempsFonctionnement_hj * md_PompageBouesBiolixDebitUnitaire_m3h)) + 1
                Else
                    mi_PompageBouesBiolixNombrePompes = Int(md_BouesSortieDebitVolumique_m3j / (md_PompageBouesBiolixTempsFonctionnement_hj * md_PompageBouesBiolixDebitUnitaire_m3h))
                End If
                md_PompageBouesBiolixTempsFonctionnement_hj = md_BouesSortieDebitVolumique_m3j / (mi_PompageBouesBiolixNombrePompes * md_PompageBouesBiolixDebitUnitaire_m3h)
            End If
        End If
    End If
Else
    md_PompageBouesBiolixDebitUnitaire_m3h = md_BouesSortieDebitVolumique_m3j / (mi_PompageBouesBiolixNombrePompes * md_PompageBouesBiolixTempsFonctionnement_hj)
End If
IncoherenceCalcul = (md_PompageBouesBiolixDebitUnitaire_m3h * mi_PompageBouesBiolixNombrePompes * md_PompageBouesBiolixTempsFonctionnement_hj - md_BouesSortieDebitVolumique_m3j) / md_BouesSortieDebitVolumique_m3j
If Abs(IncoherenceCalcul) > mi_CRITERE_INCOHERENCE Then
    ms_CommentairesCalculs(1) = Feuil47.Range("biolix_erreur_pompage_boues").Cells(1, 1)
End If
If mn_PompageBouesBiolixRendementGlobalPompes_force = True Then
    md_PompageBouesBiolixRendementGlobalPompes = md_PompageBouesBiolixRendementGlobalPompes_f
    RendementGlobalPompe = md_PompageBouesBiolixRendementGlobalPompes
Else
    RendementPompe = 0.4    'hypothèse
    PuissancePompe_kW = md_PompageBouesBiolixDebitUnitaire_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * md_PompageBouesBiolixPressionRefoulement_mCE / RendementPompe
    RendementGlobalPompe = RendementPompe * F_rendement_moteur_pompe(PuissancePompe_kW)
    md_PompageBouesBiolixRendementGlobalPompes = RendementGlobalPompe
End If
ConsommationElectriquePompageBoues_kWhj = md_PompageBouesBiolixDebitUnitaire_m3h / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * md_PompageBouesBiolixPressionRefoulement_mCE / RendementGlobalPompe * mi_PompageBouesBiolixNombrePompes * md_PompageBouesBiolixTempsFonctionnement_hj

'Consommation électrique du procédé et calcul de la part fixe
pd_ConsommationElectrique_kWhj = ConsommationElectriqueAeration_kWhj + ConsommationElectriqueAgitationCuveHomogeneisation_kWhj + ConsommationElectriqueAgitationBassinBiologique_kWhj + ConsommationElectriquePompageBoues_kWhj
pd_ConsommationElectriqueFixe_kWhj = ConsommationElectriqueAgitationCuveHomogeneisation_kWhj + ConsommationElectriqueAgitationBassinBiologique_kWhj


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
NumeroErreurNonGeree = "47.23"

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
'******************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String

'Initialisation des variables locales
NumeroErreurNonGeree = "47.1"


'Graisses en entrée de biolix
md_EntreeGraissesPourcentageFluxAdmis = Feuil47.Range("biolix_inlet_ratio").Cells(1, pi_FiliereConsideree)

'lecture des choix internes
mi_ChoixTypeSurpresseur = Feuil47.Range("choix_biolix_surpresseur").Cells(1, pi_FiliereConsideree)
mi_ChoixAgitationCuveHomogeneisation = Feuil47.Range("choix_biolix_cuve_homogen").Cells(1, pi_FiliereConsideree)
mi_ChoixAgitationBassinBiologique = Feuil47.Range("choix_biolix_bassin_bio").Cells(1, pi_FiliereConsideree)



Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True




End Sub

Sub lecture_valeurs_forcees()


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
NumeroErreurNonGeree = "47.2"
Ligne = 1
Colonne = pi_FiliereConsideree

mn_FluxDCOGraissesEntreeDimensionnement_kgj_force = Feuil47.Range("biolix_DCO_nominal").Cells(Ligne, Colonne)
mn_RendementEliminationDCO_force = Feuil47.Range("biolix_rdt_dco").Cells(Ligne, Colonne)
mn_ConsommationPhosphore_kgPj_force = Feuil47.Range("biolix_dosage_phosphore").Cells(Ligne, Colonne)
mn_ConsommationAzote_kgNj_force = Feuil47.Range("biolix_dosage_azote").Cells(Ligne, Colonne)
mn_HauteurBassinAeration_m_force = Feuil47.Range("biolix_air_hauteur_bassin").Cells(Ligne, Colonne)
mn_BesoinOxygene_kgO2j_force = Feuil47.Range("biolix_air_besoinO2").Cells(Ligne, Colonne)
mn_DebitAirProcess_Nm3j_force = Feuil47.Range("biolix_air_debit").Cells(Ligne, Colonne)
mn_RendementTransfertOxygeneMembraneEauClaire_1m_force = Feuil47.Range("biolix_rdt_oxy").Cells(Ligne, Colonne)
mn_PressionRefoulementSurpresseur_mCE_force = Feuil47.Range("biolix_air_pression_refoulement").Cells(Ligne, Colonne)
mn_VolumeCuveHomogeneisation_m3_force = Feuil47.Range("biolix_vol_cuve").Cells(Ligne, Colonne)
mn_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3_force = Feuil47.Range("biolix_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3").Cells(Ligne, Colonne)
mn_VolumeBassinBiologique_m3_force = Feuil47.Range("biolix_vol_bassin_bio").Cells(Ligne, Colonne)
mn_PuissanceSpecifiqueAgitationBassinBiologique_W_m3_force = Feuil47.Range("biolix_PuissanceSpecifiqueAgitationBassinBiologique_W_m3").Cells(Ligne, Colonne)
mn_TempsFonctionnementAgitationBassinBiologique_hj_force = Feuil47.Range("biolix_bassin_bio_duree_fonc").Cells(Ligne, Colonne)
mn_BouesSortieConcentrationMES_gL_force = Feuil47.Range("biolix_BouesSortieConcentrationMES_gL").Cells(Ligne, Colonne)
mn_BouesSortieFluxMES_kgj_force = Feuil47.Range("biolix_BouesSortieFluxMES_kgj").Cells(Ligne, Colonne)
mn_BouesSortieRatioMVSurMES_force = Feuil47.Range("biolix_BouesSortieRatioMVSurMES").Cells(Ligne, Colonne)
mn_PartBouesBiolixVersTraitementSecondaireEau_force = Feuil47.Range("biolix_PartBouesBiolixVersTraitementSecondaireEau").Cells(Ligne, Colonne)
mn_PompageBouesBiolixNombrePompes_force = Feuil47.Range("biolix_PompageBouesBiolixNombrePompes").Cells(Ligne, Colonne)
mn_PompageBouesBiolixPressionRefoulement_mCE_force = Feuil47.Range("biolix_PompageBouesBiolixPressionRefoulement_mCE").Cells(Ligne, Colonne)
mn_PompageBouesBiolixTempsFonctionnement_hj_force = Feuil47.Range("biolix_PompageBouesBiolixTempsFonctionnement_hj").Cells(Ligne, Colonne)
mn_PompageBouesBiolixDebitUnitaire_m3h_force = Feuil47.Range("biolix_PompageBouesBiolixDebitUnitaire_m3h").Cells(Ligne, Colonne)
mn_PompageBouesBiolixRendementGlobalPompes_force = Feuil47.Range("biolix_PompageBouesBiolixRendementGlobalPompes").Cells(Ligne, Colonne)


If mn_FluxDCOGraissesEntreeDimensionnement_kgj_force = True Then
    md_FluxDCOGraissesEntreeDimensionnement_kgj_f = Feuil47.Range("biolix_DCO_nominal").Cells(Ligne, Colonne)
End If
If mn_RendementEliminationDCO_force = True Then
    md_RendementEliminationDCO_f = Feuil47.Range("biolix_rdt_dco").Cells(Ligne, Colonne)
End If
If mn_ConsommationPhosphore_kgPj_force = True Then
    md_ConsommationPhosphore_kgPj_f = Feuil47.Range("biolix_dosage_phosphore").Cells(Ligne, Colonne)
End If
If mn_ConsommationAzote_kgNj_force = True Then
    md_ConsommationAzote_kgNj_f = Feuil47.Range("biolix_dosage_azote").Cells(Ligne, Colonne)
End If
If mn_HauteurBassinAeration_m_force = True Then
    md_HauteurBassinAeration_m_f = Feuil47.Range("biolix_air_hauteur_bassin").Cells(Ligne, Colonne)
End If
If mn_BesoinOxygene_kgO2j_force = True Then
    md_BesoinOxygene_kgO2j_f = Feuil47.Range("biolix_air_besoinO2").Cells(Ligne, Colonne)
End If
If mn_DebitAirProcess_Nm3j_force = True Then
    md_DebitAirProcess_Nm3j_f = Feuil47.Range("biolix_air_debit").Cells(Ligne, Colonne)
End If
If mn_RendementTransfertOxygeneMembraneEauClaire_1m_force = True Then
    md_RendementTransfertOxygeneMembraneEauClaire_1m_f = Feuil47.Range("biolix_rdt_oxy").Cells(Ligne, Colonne)
End If
If mn_PressionRefoulementSurpresseur_mCE_force = True Then
    md_PressionRefoulementSurpresseur_mCE_f = Feuil47.Range("biolix_air_pression_refoulement").Cells(Ligne, Colonne)
End If
If mn_VolumeCuveHomogeneisation_m3_force = True Then
    md_VolumeCuveHomogeneisation_m3_f = Feuil47.Range("biolix_vol_cuve").Cells(Ligne, Colonne)
End If
If mn_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3_force = True Then
    md_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3_f = Feuil47.Range("biolix_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3").Cells(Ligne, Colonne)
End If
If mn_VolumeBassinBiologique_m3_force = True Then
    md_VolumeBassinBiologique_m3_f = Feuil47.Range("biolix_vol_bassin_bio").Cells(Ligne, Colonne)
End If
If mn_PuissanceSpecifiqueAgitationBassinBiologique_W_m3_force = True Then
    md_PuissanceSpecifiqueAgitationBassinBiologique_W_m3_f = Feuil47.Range("biolix_PuissanceSpecifiqueAgitationBassinBiologique_W_m3").Cells(Ligne, Colonne)
End If
If mn_TempsFonctionnementAgitationBassinBiologique_hj_force = True Then
    md_TempsFonctionnementAgitationBassinBiologique_hj_f = Feuil47.Range("biolix_bassin_bio_duree_fonc").Cells(Ligne, Colonne)
End If
If mn_BouesSortieConcentrationMES_gL_force = True Then
    md_BouesSortieConcentrationMES_gL_f = Feuil47.Range("biolix_BouesSortieConcentrationMES_gL").Cells(Ligne, Colonne)
End If
If mn_BouesSortieFluxMES_kgj_force = True Then
    md_BouesSortieFluxMES_kgj_f = Feuil47.Range("biolix_BouesSortieFluxMES_kgj").Cells(Ligne, Colonne)
End If
If mn_BouesSortieRatioMVSurMES_force = True Then
    md_BouesSortieRatioMVSurMES_f = Feuil47.Range("biolix_BouesSortieRatioMVSurMES").Cells(Ligne, Colonne)
End If
If mn_PartBouesBiolixVersTraitementSecondaireEau_force = True Then
    md_PartBouesBiolixVersTraitementSecondaireEau_f = Feuil47.Range("biolix_PartBouesBiolixVersTraitementSecondaireEau").Cells(Ligne, Colonne)
End If
If mn_PompageBouesBiolixNombrePompes_force = True Then
    mi_PompageBouesBiolixNombrePompes_f = Feuil47.Range("biolix_PompageBouesBiolixNombrePompes").Cells(Ligne, Colonne)
End If
If mn_PompageBouesBiolixPressionRefoulement_mCE_force = True Then
    md_PompageBouesBiolixPressionRefoulement_mCE_f = Feuil47.Range("biolix_PompageBouesBiolixPressionRefoulement_mCE").Cells(Ligne, Colonne)
End If
If mn_PompageBouesBiolixTempsFonctionnement_hj_force = True Then
    md_PompageBouesBiolixTempsFonctionnement_hj_f = Feuil47.Range("biolix_PompageBouesBiolixTempsFonctionnement_hj").Cells(Ligne, Colonne)
End If
If mn_PompageBouesBiolixDebitUnitaire_m3h_force = True Then
    md_PompageBouesBiolixDebitUnitaire_m3h_f = Feuil47.Range("biolix_PompageBouesBiolixDebitUnitaire_m3h").Cells(Ligne, Colonne)
End If
If mn_PompageBouesBiolixRendementGlobalPompes_force = True Then
    md_PompageBouesBiolixRendementGlobalPompes_f = Feuil47.Range("biolix_PompageBouesBiolixRendementGlobalPompes").Cells(Ligne, Colonne)
End If


Exit Sub
'*********************************************************************
erreur_non_geree:

Call gestion_erreur_non_geree(NumeroErreurNonGeree)
pn_EffacementResultatsEnCours = True



End Sub

Public Sub fonctionnement_reel(ByRef TableauRecapitulatifFluxBoues() As Double, ByRef TableauRecapitulatifPollutionSolubleDansFluxBoues() As Double)


'*********************************************************************************************************************************************************
'OBJECTIF:

'VARIABLES D'ENTREE:

'VARIABLES DE SORTIE:
'*********************************************************************************************************************************************************

'Gestion des erreurs non gérées
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim EntreeGraissesMES_kgj As Double
Dim EntreeGraissesMV_kgj As Double
Dim EntreeGraissesDCO_kgj As Double
Dim EntreeGraissesDBO_kgj As Double
Dim EntreeGraissesNK_kgj As Double
Dim EntreeGraissesPt_kgj As Double
Dim EntreeGraissesDebitVolumique_m3j As Double
Dim ReperageBoucleReel As Boolean

'Initialisation des variables locales
NumeroErreurNonGeree = "47.5"
EntreeGraissesMES_kgj = 0
EntreeGraissesMV_kgj = 0
EntreeGraissesDCO_kgj = 0
EntreeGraissesDBO_kgj = 0
EntreeGraissesNK_kgj = 0
EntreeGraissesPt_kgj = 0
EntreeGraissesDebitVolumique_m3j = 0
ReperageBoucleReel = True

'On calcule les graisses en entrée
Call calcul_cocktail_boues_entrees(TableauRecapitulatifFluxBoues, TableauRecapitulatifPollutionSolubleDansFluxBoues, EntreeGraissesMES_kgj, EntreeGraissesMV_kgj, EntreeGraissesDebitVolumique_m3j, EntreeGraissesDCO_kgj, EntreeGraissesDBO_kgj, EntreeGraissesNK_kgj, EntreeGraissesPt_kgj)

md_FluxDCOGraissesEntreeExploitation_kgj = EntreeGraissesDCO_kgj
md_BouesSortieDebitVolumique_m3j = EntreeGraissesDCO_kgj / md_EntreeDilueeConcentrationDCO_gL

'Calcul de l'aération et de la consommation de nutriments
Call calcul_aeration_et_nutriments(ReperageBoucleReel, EntreeGraissesDCO_kgj)

'Calcul des boues en sortie
Call calcul_boues_produites(ReperageBoucleReel, EntreeGraissesDCO_kgj, EntreeGraissesDBO_kgj, EntreeGraissesNK_kgj, EntreeGraissesPt_kgj, TableauRecapitulatifFluxBoues)


Exit Sub
'*********************************************************************
erreur_non_geree:
NumeroErreurNonGeree = "47.6"

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
Dim Colonne As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "47.9"
Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

Feuil47.Range("biolix_inlet_ratio").Cells(1, Colonne) = md_EntreeGraissesPourcentageFluxAdmis
Feuil47.Range("biolix_DCO_nominal").Cells(1, Colonne) = md_FluxDCOGraissesEntreeDimensionnement_kgj
Feuil47.Range("biolix_DCO_reel").Cells(1, Colonne) = md_FluxDCOGraissesEntreeExploitation_kgj
Feuil47.Range("biolix_rdt_dco").Cells(1, Colonne) = md_RendementEliminationDCO
Feuil47.Range("biolix_dosage_phosphore").Cells(1, Colonne) = md_ConsommationPhosphore_kgPj
Feuil47.Range("biolix_dosage_azote").Cells(1, Colonne) = md_ConsommationAzote_kgNj

Feuil47.Range("biolix_air_hauteur_bassin").Cells(1, Colonne) = md_HauteurBassinAeration_m
Feuil47.Range("biolix_air_besoinO2").Cells(1, Colonne) = md_BesoinOxygene_kgO2j
Feuil47.Range("biolix_air_debit").Cells(1, Colonne) = md_DebitAirProcess_Nm3j
Feuil47.Range("biolix_rdt_oxy").Cells(1, Colonne) = md_RendementTransfertOxygeneMembraneEauClaire_1m
Feuil47.Range("biolix_air_pression_refoulement").Cells(1, Colonne) = md_PressionRefoulementSurpresseur_mCE

Feuil47.Range("biolix_vol_cuve").Cells(1, Colonne) = md_VolumeCuveHomogeneisation_m3
Feuil47.Range("biolix_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3").Cells(1, Colonne) = md_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3
Feuil47.Range("biolix_vol_bassin_bio").Cells(1, Colonne) = md_VolumeBassinBiologique_m3
Feuil47.Range("biolix_PuissanceSpecifiqueAgitationBassinBiologique_W_m3").Cells(1, Colonne) = md_PuissanceSpecifiqueAgitationBassinBiologique_W_m3
Feuil47.Range("biolix_bassin_bio_duree_fonc").Cells(1, Colonne) = md_TempsFonctionnementAgitationBassinBiologique_hj

Feuil47.Range("biolix_BouesSortieConcentrationMES_gL").Cells(1, Colonne) = md_BouesSortieConcentrationMES_gL
Feuil47.Range("biolix_BouesSortieFluxMES_kgj").Cells(1, Colonne) = md_BouesSortieFluxMES_kgj
Feuil47.Range("biolix_BouesSortieDebitVolumique_m3j").Cells(1, Colonne) = md_BouesSortieDebitVolumique_m3j
Feuil47.Range("biolix_BouesSortieRatioMVSurMES").Cells(1, Colonne) = md_BouesSortieRatioMVSurMES
Feuil47.Range("biolix_PartBouesBiolixVersTraitementSecondaireEau").Cells(1, Colonne) = md_PartBouesBiolixVersTraitementSecondaireEau
Feuil47.Range("biolix_PompageBouesBiolixNombrePompes").Cells(1, Colonne) = mi_PompageBouesBiolixNombrePompes
Feuil47.Range("biolix_PompageBouesBiolixPressionRefoulement_mCE").Cells(1, Colonne) = md_PompageBouesBiolixPressionRefoulement_mCE
Feuil47.Range("biolix_PompageBouesBiolixTempsFonctionnement_hj").Cells(1, Colonne) = md_PompageBouesBiolixTempsFonctionnement_hj
Feuil47.Range("biolix_PompageBouesBiolixDebitUnitaire_m3h").Cells(1, Colonne) = md_PompageBouesBiolixDebitUnitaire_m3h
Feuil47.Range("biolix_PompageBouesBiolixRendementGlobalPompes").Cells(1, Colonne) = md_PompageBouesBiolixRendementGlobalPompes


Feuil47.Range("biolix_electricite").Cells(1, Colonne) = pd_ConsommationElectrique_kWhj




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
Dim Colonne As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "47.10"
Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

Feuil47.Range("biolix_inlet_ratio").Cells(1, Colonne) = ""
Feuil47.Range("biolix_DCO_nominal").Cells(1, Colonne) = ""
Feuil47.Range("biolix_DCO_reel").Cells(1, Colonne) = ""
Feuil47.Range("biolix_rdt_dco").Cells(1, Colonne) = ""
Feuil47.Range("biolix_dosage_phosphore").Cells(1, Colonne) = ""
Feuil47.Range("biolix_dosage_azote").Cells(1, Colonne) = ""

Feuil47.Range("biolix_air_hauteur_bassin").Cells(1, Colonne) = ""
Feuil47.Range("biolix_air_besoinO2").Cells(1, Colonne) = ""
Feuil47.Range("biolix_air_debit").Cells(1, Colonne) = ""
Feuil47.Range("biolix_rdt_oxy").Cells(1, Colonne) = ""
Feuil47.Range("biolix_air_pression_refoulement").Cells(1, Colonne) = ""

Feuil47.Range("biolix_vol_cuve").Cells(1, Colonne) = ""
Feuil47.Range("biolix_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3").Cells(1, Colonne) = ""
Feuil47.Range("biolix_vol_bassin_bio").Cells(1, Colonne) = ""
Feuil47.Range("biolix_PuissanceSpecifiqueAgitationBassinBiologique_W_m3").Cells(1, Colonne) = ""

Feuil47.Range("biolix_bassin_bio_duree_fonc").Cells(1, Colonne) = ""

Feuil47.Range("biolix_BouesSortieConcentrationMES_gL").Cells(1, Colonne) = ""     'md_BouesSortieConcentrationMES_gL
Feuil47.Range("biolix_BouesSortieFluxMES_kgj").Cells(1, Colonne) = ""     'md_BouesSortieFluxMES_kgj
Feuil47.Range("biolix_BouesSortieDebitVolumique_m3j").Cells(1, Colonne) = ""     'md_BouesSortieDebitVolumique_m3j
Feuil47.Range("biolix_BouesSortieRatioMVSurMES").Cells(1, Colonne) = ""     'md_BouesSortieRatioMVSurMES
Feuil47.Range("biolix_PartBouesBiolixVersTraitementSecondaireEau").Cells(1, Colonne) = ""     'md_PartBouesBiolixVersTraitementSecondaireEau
Feuil47.Range("biolix_PompageBouesBiolixNombrePompes").Cells(1, Colonne) = ""     'mi_PompageBouesBiolixNombrePompes
Feuil47.Range("biolix_PompageBouesBiolixPressionRefoulement_mCE").Cells(1, Colonne) = ""     'md_PompageBouesBiolixPressionRefoulement_mCE
Feuil47.Range("biolix_PompageBouesBiolixTempsFonctionnement_hj").Cells(1, Colonne) = ""     'md_PompageBouesBiolixTempsFonctionnement_hj
Feuil47.Range("biolix_PompageBouesBiolixDebitUnitaire_m3h").Cells(1, Colonne) = ""     'md_PompageBouesBiolixDebitUnitaire_m3h
Feuil47.Range("biolix_PompageBouesBiolixRendementGlobalPompes").Cells(1, Colonne) = ""     'md_PompageBouesBiolixRendementGlobalPompes

Feuil47.Range("biolix_electricite").Cells(1, Colonne) = ""




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
If version_Developpement = False Then On Error GoTo erreur_non_geree

'Déclaration des variables locales
Dim NumeroErreurNonGeree As String
Dim I As Integer
Dim j As Integer

'Initialisation des variables locales
NumeroErreurNonGeree = "47.0"
I = 0
j = 0

'variables contenant les choix relatifs au procédé
mi_ChoixTypeSurpresseur = 0
mi_ChoixAgitationCuveHomogeneisation = 0
mi_ChoixAgitationBassinBiologique = 0
md_EntreeGraissesPourcentageFluxAdmis = 0

'variables contenant les valeurs des paramètres qui sont affichés à l'utilisateur en fin de calcul
md_FluxDCOGraissesEntreeDimensionnement_kgj = 0
md_FluxDCOGraissesEntreeExploitation_kgj = 0
md_RendementEliminationDCO = 0
md_ConsommationAzote_kgNj = 0
md_ConsommationPhosphore_kgPj = 0
md_HauteurBassinAeration_m = 0
md_BesoinOxygene_kgO2j = 0
md_DebitAirProcess_Nm3j = 0
md_RendementTransfertOxygeneMembraneEauClaire_1m = 0
md_PressionRefoulementSurpresseur_mCE = 0
md_VolumeCuveHomogeneisation_m3 = 0
md_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3 = 0
md_VolumeBassinBiologique_m3 = 0
md_PuissanceSpecifiqueAgitationBassinBiologique_W_m3 = 0
md_TempsFonctionnementAgitationBassinBiologique_hj = 0
md_BouesSortieConcentrationMES_gL = 0
md_BouesSortieFluxMES_kgj = 0
md_BouesSortieDebitVolumique_m3j = 0
md_BouesSortieRatioMVSurMES = 0
md_PartBouesBiolixVersTraitementSecondaireEau = 0
mi_PompageBouesBiolixNombrePompes = 0
md_PompageBouesBiolixPressionRefoulement_mCE = 0
md_PompageBouesBiolixTempsFonctionnement_hj = 0
md_PompageBouesBiolixDebitUnitaire_m3h = 0
md_PompageBouesBiolixRendementGlobalPompes = 0
pd_ConsommationElectrique_kWhj = 0
pd_ConsommationElectriqueFixe_kWhj = 0

'variables booléennes permettant de repérer si la valeur a été forcée par l'utilisateur (TRUE si renseignée, FALSE sinon)
mn_FluxDCOGraissesEntreeDimensionnement_kgj_force = False
mn_RendementEliminationDCO_force = False
mn_ConsommationPhosphore_kgPj_force = False
mn_ConsommationAzote_kgNj_force = False
mn_HauteurBassinAeration_m_force = False
mn_BesoinOxygene_kgO2j_force = False
mn_DebitAirProcess_Nm3j_force = False
mn_RendementTransfertOxygeneMembraneEauClaire_1m_force = False
mn_PressionRefoulementSurpresseur_mCE_force = False
mn_VolumeCuveHomogeneisation_m3_force = False
mn_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3_force = False
mn_VolumeBassinBiologique_m3_force = False
mn_PuissanceSpecifiqueAgitationBassinBiologique_W_m3_force = False
mn_TempsFonctionnementAgitationBassinBiologique_hj_force = False
mn_BouesSortieConcentrationMES_gL_force = False
mn_BouesSortieFluxMES_kgj_force = False
mn_BouesSortieRatioMVSurMES_force = False
mn_PartBouesBiolixVersTraitementSecondaireEau_force = False
mn_PompageBouesBiolixNombrePompes_force = False
mn_PompageBouesBiolixPressionRefoulement_mCE_force = False
mn_PompageBouesBiolixTempsFonctionnement_hj_force = False
mn_PompageBouesBiolixDebitUnitaire_m3h_force = False
mn_PompageBouesBiolixRendementGlobalPompes_force = False

'variables contenant les valeurs forcées par l'utilisateur
md_FluxDCOGraissesEntreeDimensionnement_kgj_f = 0
md_RendementEliminationDCO_f = 0
md_ConsommationPhosphore_kgPj_f = 0
md_ConsommationAzote_kgNj_f = 0
md_HauteurBassinAeration_m_f = 0
md_BesoinOxygene_kgO2j_f = 0
md_DebitAirProcess_Nm3j_f = 0
md_RendementTransfertOxygeneMembraneEauClaire_1m_f = 0
md_PressionRefoulementSurpresseur_mCE_f = 0
md_VolumeCuveHomogeneisation_m3_f = 0
md_PuissanceSpecifiqueAgitationCuveHomogeneisation_W_m3_f = 0
md_VolumeBassinBiologique_m3_f = 0
md_PuissanceSpecifiqueAgitationBassinBiologique_W_m3_f = 0
md_TempsFonctionnementAgitationBassinBiologique_hj_f = 0
md_BouesSortieConcentrationMES_gL_f = 0
md_BouesSortieFluxMES_kgj_f = 0
md_BouesSortieRatioMVSurMES_f = 0
md_PartBouesBiolixVersTraitementSecondaireEau_f = 0
mi_PompageBouesBiolixNombrePompes_f = 0
md_PompageBouesBiolixPressionRefoulement_mCE_f = 0
md_PompageBouesBiolixTempsFonctionnement_hj_f = 0
md_PompageBouesBiolixDebitUnitaire_m3h_f = 0
md_PompageBouesBiolixRendementGlobalPompes_f = 0

'hypothèses utilisées dans les calculs
md_BesoinSpecifiqueO2EliminationDCO_kgO2_kgDCO = 0
md_AerationFacteurK = 0
md_TempsFonctionnementAgitationCuveHomogeneisation_hj = 0
md_TempsPassageHydrauliqueCuveHomogeneisation_j = 0
md_TempsPassageHydrauliqueBassinBiologique_j = 0
md_BesoinSpecifiquePhophoreBacteries_kgPj = 0
md_BesoinSpecifiqueAzoteBacteries_kgNj = 0
md_HauteurImmersionDiffuseur_m = 0
md_ConsommationElectriqueSpecifiqueSurpresseurs_Wh_Nm3mCE = 0
md_ProductionSpecifiqueBoues_kgMVS_kgDCOeliminee = 0
md_EntreeDilueeConcentrationDCO_gL = 0

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
