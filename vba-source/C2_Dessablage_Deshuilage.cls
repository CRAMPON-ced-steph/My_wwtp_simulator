Public Mat_volume
Public Mat_nb_ouvrage
Public Mat_Q_air_spec
Public Mat_tf
Public Mat_Prefoulement
Public Mat_conso_spec
Public Mat_conc_graisse
Public Mat_Graisse_debit
Public Mat_MVMES
Public Mat_Graisse_DCO_aff
Public Mat_elec_desdes
Public Mat_Graisse_interne


Sub DessablageDeshuilage(I, Mat_vol_ouvrage, Mat_nbouvrage, Mat_Qair, Mat_tspfct, Mat_Pref, Mat_csp, _
Mat_MES, Mat_vol_jour, Mat_MV_MES, Mat_DCO_graisse, Mat_consoelec, Mat_Graisse_Q, Mat_Graisse_DCO)

NOMBRE_JOUR_PAR_AN = 365
NOMBRE_MINUTE_PAR_HEURE = 60
NOMBRE_HEURE_PAR_JOUR = 24

'hypothèses
graisse_MS_EH = 0.83 / NOMBRE_JOUR_PAR_AN 'ratio MS par équivalent-habitant et par an
ratio_graisse_DBO_DCO = 0.9 'ratio de DBO dans DCO pour les graisses
design_tps_sejour_min = 20 'temps de séjour pour le design basé sur le débit nominal moyen
PuissanceRacleurParOuvrage_kW = 0.65   'similaire valeurs d'OTV pour râcleurs'Puissance de râcleur par ouvrage (kW)  entre 0.3 et 1 selon DWA

'''''''''''''''valeurs défauts'
'''''''''''''''''''''''''''
'''''''''''''''''''''''''''
'Nombre d'ouvrages
If Mat_nbouvrage(1, 2) = "" Then NombreOuvrages = 2
If Mat_nbouvrage(1, 2) <> "" Then NombreOuvrages = Mat_nbouvrage(1, 2)
NombreOuvrages_def = 2

'Débit d'air spécifique selon DWA A216 entre 0.5 et 1.3 (Nm3/h)/m3 dessableur
If Mat_Qair(1, 2) = "" Then DebitAirSpecifique_Nm3h_m3Dessableur = 0.5
If Mat_Qair(1, 2) <> "" Then DebitAirSpecifique_Nm3h_m3Dessableur = Mat_Qair(1, 2)
DebitAirSpecifique_Nm3h_m3Dessableur_def = 0.5

'Temps de fonctionnement de l'aération en h/j
If Mat_tspfct(1, 2) = "" Then TempsFonctionnementAeration_hj = 24
If Mat_tspfct(1, 2) <> "" Then TempsFonctionnementAeration_hj = Mat_tspfct(1, 2)
TempsFonctionnementAeration_hj_def = 24

'Pression de refoulement de l'air en mCE
If Mat_Pref(1, 2) = "" Then PressionRefoulementSurpresseurs_mCE = 4
If Mat_Pref(1, 2) <> "" Then PressionRefoulementSurpresseurs_mCE = Mat_Pref(1, 2)
PressionRefoulementSurpresseurs_mCE_def = 4


'Consommation électrique spécifique des surpresseurs
If Mat_csp(1, 2) = "" Then ConsommationElectriqueSpecifiqueSurpresseurs_Wh_Nm3mCE = 4.5 'surpresseurs à lobes par défaut
If Mat_csp(1, 2) <> "" Then ConsommationElectriqueSpecifiqueSurpresseurs_Wh_Nm3mCE = Mat_csp(1, 2)
ConsommationElectriqueSpecifiqueSurpresseurs_Wh_Nm3mCE_def = 4.5


'ratio_MV_MES des graisses
If Mat_MV_MES(1, 2) = "" Then graisse_MV_MES = 0.9
If Mat_MV_MES(1, 2) <> "" Then graisse_MV_MES = Mat_MV_MES(1, 2)
graisse_MV_MES_def = 0.9


'concentration des graisses
If Mat_MES(1, 2) = "" Then graisse_concentration = 92 'g/L sans concentrateur
If Mat_MES(1, 2) <> "" Then graisse_concentration = Mat_MES(1, 2) 'graisse_MV_MES = Mat_MES(1, 2)
graisse_concentration_def = 92

'concentration des graisses
If Mat_vol_ouvrage(1, 2) = "" Then VolumeTotalBassins_m3 = (design_tps_sejour_min / (NOMBRE_MINUTE_PAR_HEURE * NOMBRE_HEURE_PAR_JOUR)) * Q_nominal
If Mat_vol_ouvrage(1, 2) <> "" Then VolumeTotalBassins_m3 = Mat_vol_ouvrage(1, 2)
VolumeTotalBassins_m3_def = (design_tps_sejour_min / (NOMBRE_MINUTE_PAR_HEURE * NOMBRE_HEURE_PAR_JOUR)) * Q_nominal

'calcul sur les graisses

equiv_hab = Range("CelVF_Eq_hab").Value
graisse_MS = graisse_MS_EH * equiv_hab + 0.5 * vidange_MES
If Mat_Graisse_Q(1, 2) <> "" Then graisse_Q = Mat_Graisse_Q(1, 2)
If Mat_Graisse_Q(1, 2) = "" Then graisse_Q = graisse_MS / graisse_concentration
graisse_MS = graisse_concentration * graisse_Q
graisse_Q_def = graisse_MS / graisse_concentration

graisse_MS_def = graisse_concentration_def * graisse_Q_def
    
If Mat_Graisse_DCO(1, 2) <> "" Then graisse_DCO = Mat_Graisse_DCO(1, 2)
If Mat_Graisse_DCO(1, 2) = "" Then graisse_DCO = ratio("codigestion_graisses", "NK_MV") * graisse_MV_MES * graisse_MS
    graisse_DCO_def = ratio("codigestion_graisses", "NK_MV") * graisse_MV_MES * graisse_MS
    

graisse_DBO = graisse_DCO * ratio_graisse_DBO_DCO


'report des valeurs dans vecteur appropprié
graisse_MS = graisse_MS
graisse_Q = graisse_Q
graisse_MV_MES = graisse_MV_MES
graisse_DCO_MES = graisse_DCO / graisse_MS
graisse_DBO_MES = graisse_DBO / graisse_MS
graisse_NK_MES = ratio("codigestion_graisses", "NK_MV") * graisse_MV_MES
graisse_Pt_MES = ratio("codigestion_graisses", "Pt_MES")

Mat_Graisse_interne = WorksheetFunction.Transpose(Array(graisse_MS, graisse_Q, graisse_MV_MES, graisse_DCO_MES, graisse_DBO_MES, graisse_NK_MES, graisse_Pt_MES))


' calcul_consommation_electrique()

electricite_racleur_def = PuissanceRacleurParOuvrage_kW * NombreOuvrages_def * NOMBRE_HEURE_PAR_JOUR 'RACLEUR
electricite_aeration_def = (ConsommationElectriqueSpecifiqueSurpresseurs_Wh_Nm3mCE_def / 1000) * (DebitAirSpecifique_Nm3h_m3Dessableur_def * VolumeTotalBassins_m3_def * TempsFonctionnementAeration_hj_def) * PressionRefoulementSurpresseurs_mCE_def 'Aeration
electricite_def = electricite_racleur_def + electricite_aeration_def 'CONSOMMATION GLOBALE D'ELECTRICITE

electricite_racleur = PuissanceRacleurParOuvrage_kW * NombreOuvrages * NOMBRE_HEURE_PAR_JOUR 'RACLEUR
electricite_aeration = (ConsommationElectriqueSpecifiqueSurpresseurs_Wh_Nm3mCE / 1000) * (DebitAirSpecifique_Nm3h_m3Dessableur * VolumeTotalBassins_m3 * TempsFonctionnementAeration_hj) * PressionRefoulementSurpresseurs_mCE 'Aeration
electricite = electricite_racleur + electricite_aeration 'CONSOMMATION GLOBALE D'ELECTRICITE


Mat_volume = Array(VolumeTotalBassins_m3_def, Mat_vol_ouvrage(1, 2), VolumeTotalBassins_m3)
Mat_nb_ouvrage = Array(NombreOuvrages_def, Mat_nbouvrage(1, 2), NombreOuvrages)
Mat_Q_air_spec = Array(DebitAirSpecifique_Nm3h_m3Dessableur_def, Mat_Qair(1, 2), DebitAirSpecifique_Nm3h_m3Dessableur)
Mat_tf = Array(TempsFonctionnementAeration_hj_def, Mat_tspfct(1, 2), TempsFonctionnementAeration_hj)
Mat_Prefoulement = Array(PressionRefoulementSurpresseurs_mCE_def, Mat_Pref(1, 2), PressionRefoulementSurpresseurs_mCE)
Mat_conso_spec = Array(ConsommationElectriqueSpecifiqueSurpresseurs_Wh_Nm3mCE_def, Mat_csp(1, 2), ConsommationElectriqueSpecifiqueSurpresseurs_Wh_Nm3mCE)
Mat_conc_graisse = Array(graisse_concentration_def, Mat_MES(1, 2), graisse_concentration)
Mat_Graisse_debit = Array(graisse_Q_def, Mat_Graisse_Q(1, 2), graisse_Q)
Mat_MVMES = Array(graisse_MV_MES_def, Mat_MV_MES(1, 2), graisse_MV_MES)
Mat_Graisse_DCO_aff = Array(graisse_DCO_def, Mat_Graisse_DCO(1, 2), graisse_DCO)
Mat_elec_desdes = Array(electricite_def, "", electricite)








End Sub


