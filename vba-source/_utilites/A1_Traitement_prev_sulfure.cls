Attribute VB_Name = "A1_Traitement_prev_sulfure"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'Option Explicit
Public Mat_Ca2NO3
Public Mat_H_reseau
Public Mat_PTSR_result As Variant
Public Ca_2NO3_kgj As Double
Public Ca_2NO3_pure_kgj As Double
Public pd_PureteNitrateDeCalcium As Double

Sub choix_traitement_sulfures_preventif(Mat_charges_reelle, Mat_TPSR)

Q_nominal = Mat_charges_reelle(1, 1)

'hyp_purete_reactif_commercial(reactifs_repere_Ca_2NO3) = 0.45    'recherche internet "Caractéristiques Nutriox" --> page trouvée sur le site de Yara
pd_PureteNitrateDeCalcium = 0.45


''calcul du rate de stripping
'a_rate_stripping_T(0) = -0.046
'a_rate_stripping_T(1) = 0.0069

'masse volumique
md_MasseVolumiqueNitrateDeCalciumCommercial_kg_m3 = 1420  'source: Yara (recherche "Caractéristiques Nutriox" sur Internet)

preventif_CaNO3_dose_mL_m3h_force = Mat_TPSR(1, 2)
reseau_tps_retention_h_force = Mat_TPSR(2, 2)


'dose préventif Ca(NO3)2
If preventif_CaNO3_dose_mL_m3h_force <> "" Then preventif_CaNO3_dose_mL_m3h = preventif_CaNO3_dose_mL_m3h_force
If preventif_CaNO3_dose_mL_m3h_force = "" Then preventif_CaNO3_dose_mL_m3h = 15
            
'temps de rétention réseau
If reseau_tps_retention_h_force <> "" Then reseau_tps_retention_h = reseau_tps_retention_h_force
If reseau_tps_retention_h_force = "" Then reseau_tps_retention_h = 10

'on calcule la consommation de réactifs
'preventif_CaNO3_Lj = preventif_CaNO3_dose_mL_m3h / 1000 * reseau_tps_retention_h * Q_nominal
Ca_2NO3_kgj = (preventif_CaNO3_dose_mL_m3h / 1000000 * reseau_tps_retention_h * Q_nominal) * md_MasseVolumiqueNitrateDeCalciumCommercial_kg_m3 * pd_PureteNitrateDeCalcium


Ca_2NO3_pure_kgj = Ca_2NO3_kgj / pd_PureteNitrateDeCalcium



'    'on calcule alors le stripping
'    If rate_stripping_force = True Then
'        rate_stripping = rate_stripping_f
'    Else
'        rate_stripping = 0
'        For i = 0 To 1
'            rate_stripping = rate_stripping + a_rate_stripping_T(i) * pd_TemperatureEauDimensionnement_°C ^ i
'        Next i
'    End If
'
'    SH_strippe_kgj = rate_stripping * (SH_nominal - SH_traite_kgj)
'
'    If choix_traitement_sulfures_preventif = True Then
'        Feuil51.Range("trait_sulfures_preventif_CaNO3_dose_mL_m3h").Cells(1, Colonne) = preventif_CaNO3_dose_mL_m3h
'        Feuil51.Range("trait_sulfures_reseau_tps_retention_h").Cells(1, Colonne) = reseau_tps_retention_h
'        'Feuil51.Range("trait_sulfures_preventif_CaNO3_Lj").Cells(1, colonne) = preventif_CaNO3_Lj
'    End If
'
'
'    If choix_desodo_chimique = True Then
'        Feuil29.Range("desodo_chim_stripping_H2S").Cells(1, Colonne) = rate_stripping
'    End If
'
'Dim Mat_Ca2NO3 As Variant
'Dim Mat_H_reseau As Variant
'Dim Mat_PTSR_result(2, 3)

'Mat_PTSR_result = Array(15, preventif_CaNO3_dose_mL_m3h, preventif_CaNO3_dose_mL_m3h, 10, reseau_tps_retention_h, reseau_tps_retention_h)
'
'Mat_PTSR_result(1, 1) = 1
'Mat_PTSR_result(1, 2) = 1
'Mat_PTSR_result(1, 3) = 1
'
'Mat_PTSR_result(2, 1) = 1
'Mat_PTSR_result(2, 2) = 1
'Mat_PTSR_result(2, 3) = 1


Mat_Ca2NO3 = Array(15, preventif_CaNO3_dose_mL_m3h, preventif_CaNO3_dose_mL_m3h)
Mat_H_reseau = Array(10, reseau_tps_retention_h, reseau_tps_retention_h)


End Sub










