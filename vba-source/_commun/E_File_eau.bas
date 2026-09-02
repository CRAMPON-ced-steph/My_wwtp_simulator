Sub traitement_preventif_sulfure(I, Mat_charges_reelle, Mat_TPSR)
Set wf = WorksheetFunction
Dim traitement_prev_S As New A1_Traitement_prev_sulfure
Call traitement_prev_S.choix_traitement_sulfures_preventif(Mat_charges_reelle, Mat_TPSR)
Conso_Ca_2NO3_kgj = traitement_prev_S.Ca_2NO3_kgj
Conso_Ca_2NO3_pure_kgj = traitement_prev_S.Ca_2NO3_pure_kgj
PureteNitrateDeCalcium = traitement_prev_S.pd_PureteNitrateDeCalcium
Range("Mat_TPSR_cano3_" & I) = traitement_prev_S.Mat_Ca2NO3
Range("Mat_TPSR_tr_" & I) = traitement_prev_S.Mat_H_reseau
Range("Mat_TPSR_reactif_Vp_" & I) = wf.Transpose(Array(Conso_Ca_2NO3_kgj, PureteNitrateDeCalcium, Conso_Ca_2NO3_pure_kgj))
Mat_TPSR_reactif_Vr = Range("Mat_TPSR_reactif_Vr_" & I)
If Mat_TPSR_reactif_Vr(2, 1) = "" Then Range("Mat_TPSR_reactif_Ve_" & I) = Range("Mat_TPSR_reactif_Vp_" & I)
If Mat_TPSR_reactif_Vr(2, 1) <> "" Then Range("Mat_TPSR_reactif_Ve_" & I) = wf.Transpose(Array(Conso_Ca_2NO3_kgj, Mat_TPSR_reactif_Vr(2, 1), Conso_Ca_2NO3_pure_kgj / Mat_TPSR_reactif_Vr(2, 1) * PureteNitrateDeCalcium))
End Sub
Sub traitement_curatif_sulfure(I, Mat_charges_reelle, Mat_TPSR, Mat_PPPR)
Set wf = WorksheetFunction
HS_nominal_mgL = Mat_PPPR(3, 3)
Q_nominal = Mat_charges_reelle(1, 1)
Dim traitement_curatif_S As New A2_Traitement_curatif_sulfure
Call traitement_curatif_S.choix_traitement_sulfures_curatif(I, Q_nominal, HS_nominal_mgL)
Range("Mat_TCS_conc_HS_" & I) = traitement_curatif_S.Mat_TCS_conc_HS
Range("Mat_TCS_L_conduite_" & I) = traitement_curatif_S.Mat_TCS_L_conduite
Range("Mat_TCS_rdtO2_" & I) = traitement_curatif_S.Mat_rdt02
Range("Mat_TCS_reactif_" & I) = traitement_curatif_S.Mat_TCS_reactif
End Sub
Sub relevement(I, nombre_pompe_relevement, Mat_charges_reelle)
Q_nominal = Mat_charges_reelle(1, 1)
    For j = 1 To nombre_pompe_relevement
    Select Case j
    Case "1"
        choix_type_de_pompe1 = Range("btn_relevement1_type_" & I).Text
        choix_type_regulation_pompe1 = Range("btn_relevement1_VV_" & I).Text
        Mat_1_Vp = Range("Mat_relevement1_Vp_" & I)
        Mat_1_Vr = Range("Mat_relevement1_Vr_" & I)
        'Mat_1_Ve = Range("Mat_relevement1_Ve_" & i)
        Dim pompe_relevement As New B1_Relevement
        Call pompe_relevement.Pompes_relevement(I, Q_nominal, choix_type_de_pompe1, choix_type_regulation_pompe1, Mat_1_Vp, Mat_1_Vr)
        Range("Mat_relevement1_Vp_" & I) = pompe_relevement.Mat_result_relev_Vp
        Range("Mat_relevement1_Ve_" & I) = pompe_relevement.Mat_result_relev_Ve
    Case "2"
        choix_type_de_pompe2 = Range("btn_relevement2_type_" & I).Text
        choix_type_regulation_pompe2 = Range("btn_relevement2_VV_" & I).Text
        Mat_2_Vp = Range("Mat_relevement2_Vp_" & I)
        Mat_2_Vr = Range("Mat_relevement2_Vr_" & I)
        'Mat_2_Ve = Range("Mat_relevement2_Ve_" & i)
        Dim pompe_relevement2 As New B1_Relevement
        Call pompe_relevement2.Pompes_relevement(I, Q_nominal, choix_type_de_pompe1, choix_type_regulation_pompe1, Mat_1_Vp, Mat_1_Vr)
        Range("Mat_relevement2_Vp_" & I) = pompe_relevement2.Mat_result_relev_Vp
        Range("Mat_relevement2_Ve_" & I) = pompe_relevement2.Mat_result_relev_Ve
    Case "3"
        choix_type_de_pompe3 = Range("btn_relevement3_type_" & I).Text
        choix_type_regulation_pompe3 = Range("btn_relevement3_VV_" & I).Text
        Mat_3_Vp = Range("Mat_relevement3_Vp_" & I)
        Mat_3_Vr = Range("Mat_relevement3_Vr_" & I)
        'Mat_3_Ve = Range("Mat_relevement3_Ve_" & i)
        Dim pompe_relevement3 As New B1_Relevement
        Call pompe_relevement3.Pompes_relevement(I, Q_nominal, choix_type_de_pompe1, choix_type_regulation_pompe1, Mat_1_Vp, Mat_1_Vr)
        Range("Mat_relevement3_Vp_" & I) = pompe_relevement3.Mat_result_relev_Vp
        Range("Mat_relevement3_Ve_" & I) = pompe_relevement3.Mat_result_relev_Ve
    Case "4"
        choix_type_de_pompe4 = Range("btn_relevement4_type_" & I).Text
        choix_type_regulation_pompe4 = Range("btn_relevement4_VV_" & I).Text
        Mat_4_Vp = Range("Mat_relevement4_Vp_" & I)
        Mat_4_Vr = Range("Mat_relevement4_Vr_" & I)
        'Mat_4_Ve = Range("Mat_relevement4_Ve_" & i)
        Dim pompe_relevement4 As New B1_Relevement
        Call pompe_relevement4.Pompes_relevement(I, Q_nominal, choix_type_de_pompe1, choix_type_regulation_pompe1, Mat_1_Vp, Mat_1_Vr)
        Range("Mat_relevement4_Vp_" & I) = pompe_relevement4.Mat_result_relev_Vp
        Range("Mat_relevement4_Ve_" & I) = pompe_relevement4.Mat_result_relev_Ve
    End Select
    Next j
End Sub
Sub degrillage(I)
Mat_ecart = Range("Mat_degrillage_ecart_" & I)
Mat_quantite = Range("Mat_degrillage_quantite_" & I)
Mat_csp = Range("Mat_degrillage_CSP_" & I)
Mat_consoelec = Range("Mat_degrillage_consoelec_" & I)
CoefficientPointeHydrauliqueTempsPluie = Range("Mat_pointe_hydraulique_" & I).Cells(2, 3)
Dim calcul_degrillage As New C1_Degrillage
Call calcul_degrillage.Degrillage_calcul(I, Mat_ecart, Mat_quantite, Mat_csp, Mat_consoelec, CoefficientPointeHydrauliqueTempsPluie)
Range("Mat_degrillage_ecart_" & I) = calcul_degrillage.Mat_ecart_result
Range("Mat_degrillage_quantite_" & I) = calcul_degrillage.Mat_quantite_result
Range("Mat_degrillage_CSP_" & I) = calcul_degrillage.Mat_csp_result
Range("Mat_degrillage_consoelec_" & I) = calcul_degrillage.Mat_consoelec_result
End Sub
Sub DessablageEtDeshuilage(I)
Mat_vol_ouvrage = Range("Mat_desdes_volume_" & I)
Mat_nbouvrage = Range("Mat_desdes_nbouvrage_" & I)
Mat_Qair = Range("Mat_desdes_Qair_" & I)
Mat_tspfct = Range("Mat_desdes_tspfct_" & I)
Mat_Pref = Range("Mat_desdes_Pref_" & I)
Mat_csp = Range("Mat_desdes_csp_" & I)
Mat_MES = Range("Mat_desdes_MES_" & I) 'graisse
Mat_vol_jour = Range("Mat_desdes_vol_jour_" & I)
Mat_MV_MES = Range("Mat_desdes_MV_MES_" & I)
Mat_DCO_graisse = Range("Mat_desdes_DCO_graisse_" & I)
Mat_consoelec = Range("Mat_desdes_consoelec_" & I)
Mat_Graisse_Q = Range("Mat_desdes_vol_jour_" & I)
Mat_Graisse_DCO = Range("Mat_desdes_DCO_graisse_" & I)
Dim calcul_desdes As New C2_Dessablage_Deshuilage
Call calcul_desdes.DessablageDeshuilage(I, Mat_vol_ouvrage, Mat_nbouvrage, Mat_Qair, Mat_tspfct, Mat_Pref, Mat_csp, _
Mat_MES, Mat_vol_jour, Mat_MV_MES, Mat_DCO_graisse, Mat_consoelec, Mat_Graisse_Q, Mat_Graisse_DCO)
Range("Mat_desdes_volume_" & I) = calcul_desdes.Mat_volume
Range("Mat_desdes_nbouvrage_" & I) = calcul_desdes.Mat_nb_ouvrage
Range("Mat_desdes_Qair_" & I) = calcul_desdes.Mat_Q_air_spec
Range("Mat_desdes_tspfct_" & I) = calcul_desdes.Mat_tf
Range("Mat_desdes_Pref_" & I) = calcul_desdes.Mat_Prefoulement
Range("Mat_desdes_csp_" & I) = calcul_desdes.Mat_conso_spec
Range("Mat_desdes_MV_MES_" & I) = calcul_desdes.Mat_MVMES
Range("Mat_desdes_vol_jour_" & I) = calcul_desdes.Mat_Graisse_debit
Range("Mat_desdes_DCO_graisse_" & I) = calcul_desdes.Mat_Graisse_DCO_aff
Range("Mat_desdes_consoelec_" & I) = calcul_desdes.Mat_elec_desdes
Range("Mat_desdes_MES_" & I) = calcul_desdes.Mat_conc_graisse
Range("Mat_graisse_int_" & I) = calcul_desdes.Mat_Graisse_interne
End Sub
Sub Decantation_Primaire(I)

Set wf = WorksheetFunction

Q_poste_pourcent = Range("btn_I_simple_bypass_" & I)
Decantation_lamellaire = Range("btn_I_simple_lamelle_" & I)


Mat_dec_I_Vp = Range("Mat_dec_I_Vp_" & I)
Mat_dec_I_Vr = Range("Mat_dec_I_Vr_" & I)
Mat_dec_I_Ve = Range("Mat_dec_I_Ve_" & I)

Dim Decanteur_I As New D1_decanteur_simple
Call Decanteur_I.Decanteur_I_simple(I, Q_poste_pourcent, Decantation_lamellaire, Mat_dec_I_Vp, Mat_dec_I_Vr, Mat_dec_I_Ve)




End Sub












