Sub Decanteur_I_simple(I, Q_poste_pourcent, Decantation_lamellaire, Mat_dec_I_Vp, Mat_dec_I_Vr, Mat_dec_I_Ve)


'hypothèses


'Voici l'ensemble des paramètres fixés par défaut
MV_MES_reference = 0.68 'si réactif, on réduit les MV/MES suivant la dose de FeCl3
hyp_rdt_P_NoDiscfilter = 0.11 'hypothèse de rendement sur P sans retour d'eaux sales de discfilter
NOMBRE_HEURE_PAR_JOUR = 24



'valeurs par défauts



'On retrouve ici l'ensemble des valeurs pour lesquelles on attribue une valeur guide par défaut
    'by-pass
    If bypass_force = True Then
        bypass = bypass_f
    Else
        bypass = 0
    End If

'Nombre d'ouvrages
If Mat_dec_I_Vr(2, 1) = "" Then Mat_dec_I_Ve(2, 1) = 1
If Mat_dec_I_Vr(2, 1) <> "" Then Mat_dec_I_Ve(2, 1) = Mat_dec_I_Vr(2, 1)
Mat_dec_I_Vp(2, 1) = 1 'valeur par defaut
nombre_ouvrage = Mat_dec_I_Ve(2, 1)


'vitesse au mirroir limite
Select Case Decantation_lamellaire
    Case "oui": V_mirroir_limite = 13  'm/h Aquademy et CAM
    Case "non": V_mirroir_limite = 5  'm/h Aquademy et CAM
    'case "oui" and choix_biostyr = True : V_mirroir_limite = 8  'm/h Aquademy et CAM
    'case "non" and choix_biostyr = True : V_mirroir_limite = 2  'm/h Aquademy et CAM
End Select

If Mat_dec_I_Vr(7, 1) = "" Then Mat_dec_I_Ve(7, 1) = V_mirroir_limite
If Mat_dec_I_Vr(7, 1) <> "" Then Mat_dec_I_Ve(7, 1) = Mat_dec_I_Vr(7, 1)
Mat_dec_I_Vp(7, 1) = V_mirroir_limite 'valeur par defaut

'hauteur du décanteur
If Mat_dec_I_Vr(6, 1) = "" Then Mat_dec_I_Ve(6, 1) = 1
If Mat_dec_I_Vr(6, 1) <> "" Then Mat_dec_I_Ve(6, 1) = Mat_dec_I_Vr(6, 1)
Mat_dec_I_Vp(6, 1) = 3.5 'valeur par defaut
    
Select Case Decantation_lamellaire

    Case "oui"
    
        'rdt_MES
        If Mat_dec_I_Vr(10, 1) = "" Then Mat_dec_I_Ve(10, 1) = 0.6
        If Mat_dec_I_Vr(10, 1) <> "" Then Mat_dec_I_Ve(10, 1) = Mat_dec_I_Vr(10, 1)
        Mat_dec_I_Vp(10, 1) = 0.6 'valeur par defaut
        rdt_MES = Mat_dec_I_Ve(10, 1)
        
        'rdt_DCO
        If Mat_dec_I_Vr(11, 1) = "" Then Mat_dec_I_Ve(11, 1) = 0.3
        If Mat_dec_I_Vr(11, 1) <> "" Then Mat_dec_I_Ve(11, 1) = Mat_dec_I_Vr(11, 1)
        Mat_dec_I_Vp(11, 1) = 0.3 'valeur par defaut
        
        'rdt_DBO
        If Mat_dec_I_Vr(12, 1) = "" Then Mat_dec_I_Ve(12, 1) = 0.28
        If Mat_dec_I_Vr(12, 1) <> "" Then Mat_dec_I_Ve(12, 1) = Mat_dec_I_Vr(12, 1)
        Mat_dec_I_Vp(12, 1) = 0.28 'valeur par defaut
        
        'oues_concentration
        If Mat_dec_I_Vr(17, 1) = "" Then Mat_dec_I_Ve(17, 1) = 20
        If Mat_dec_I_Vr(17, 1) <> "" Then Mat_dec_I_Ve(17, 1) = Mat_dec_I_Vr(17, 1)
        Mat_dec_I_Vp(17, 1) = 20 'valeur par defaut
        
    Case "non"
    
        'rdt_MES
        If Mat_dec_I_Vr(10, 1) = "" Then Mat_dec_I_Ve(10, 1) = 0.5
        If Mat_dec_I_Vr(10, 1) <> "" Then Mat_dec_I_Ve(10, 1) = Mat_dec_I_Vr(10, 1)
        Mat_dec_I_Vp(10, 1) = 0.5 'valeur par defaut
        
        'rdt_DCO
        If Mat_dec_I_Vr(11, 1) = "" Then Mat_dec_I_Ve(11, 1) = 0.28
        If Mat_dec_I_Vr(11, 1) <> "" Then Mat_dec_I_Ve(11, 1) = Mat_dec_I_Vr(11, 1)
        Mat_dec_I_Vp(11, 1) = 0.28 'valeur par defaut
        
        'rdt_DBO
        If Mat_dec_I_Vr(12, 1) = "" Then Mat_dec_I_Ve(12, 1) = 0.25
        If Mat_dec_I_Vr(12, 1) <> "" Then Mat_dec_I_Ve(12, 1) = Mat_dec_I_Vr(12, 1)
        Mat_dec_I_Vp(12, 1) = 0.25 'valeur par defaut
        
        'oues_concentration
        If Mat_dec_I_Vr(17, 1) = "" Then Mat_dec_I_Ve(17, 1) = 10
        If Mat_dec_I_Vr(17, 1) <> "" Then Mat_dec_I_Ve(17, 1) = Mat_dec_I_Vr(17, 1)
        Mat_dec_I_Vp(17, 1) = 10 'valeur par defaut
    
End Select

'rdt_NK
If Mat_dec_I_Vr(13, 1) = "" Then Mat_dec_I_Ve(13, 1) = 0.09
If Mat_dec_I_Vr(13, 1) <> "" Then Mat_dec_I_Ve(13, 1) = Mat_dec_I_Vr(13, 1)
Mat_dec_I_Vp(13, 1) = 0.09 'valeur par defaut

'rdt_P
If Mat_dec_I_Vr(14, 1) = "" Then Mat_dec_I_Ve(14, 1) = hyp_rdt_P_NoDiscfilter
If Mat_dec_I_Vr(14, 1) <> "" Then Mat_dec_I_Ve(14, 1) = Mat_dec_I_Vr(14, 1)
Mat_dec_I_Vp(14, 1) = hyp_rdt_P_NoDiscfilter 'valeur par defaut

'MV_MES
If Mat_dec_I_Vr(15, 1) = "" Then Mat_dec_I_Ve(15, 1) = MV_MES_reference
If Mat_dec_I_Vr(15, 1) <> "" Then Mat_dec_I_Ve(15, 1) = Mat_dec_I_Vr(15, 1)
Mat_dec_I_Vp(15, 1) = MV_MES_reference 'valeur par defaut


 'extraction des boues
If Mat_dec_I_Vr(21, 1) = "" Then Mat_dec_I_Ve(21, 1) = 2
If Mat_dec_I_Vr(21, 1) <> "" Then Mat_dec_I_Ve(21, 1) = Mat_dec_I_Vr(21, 1)
Mat_dec_I_Vp(21, 1) = 2

'extraction_P_refoulement
If Mat_dec_I_Vr(22, 1) = "" Then Mat_dec_I_Ve(22, 1) = 15
If Mat_dec_I_Vr(22, 1) <> "" Then Mat_dec_I_Ve(22, 1) = Mat_dec_I_Vr(22, 1)
Mat_dec_I_Vp(22, 1) = 15

   
'extraction_tps_fct_pompe
If Mat_dec_I_Vr(23, 1) = "" Then Mat_dec_I_Ve(23, 1) = NOMBRE_HEURE_PAR_JOUR
If Mat_dec_I_Vr(23, 1) <> "" Then Mat_dec_I_Ve(23, 1) = Mat_dec_I_Vr(23, 1)
Mat_dec_I_Vp(23, 1) = NOMBRE_HEURE_PAR_JOUR




'dimensionnement


'détermination du débit d'eaux sales au global

'moi je vais dégrouper les eaux pour plus de clareté

'au final cette partie doit pouvoir être extraite du mocule de classe
ES_discfilter = Range("Mat_ES_Discfilter_" & I)
ES_filtresable = Range("Mat_ES_filtresable_" & I)
ES_ep = Range("Mat_ES_ep_" & I)
ES_deshydrat = Range("Mat_ES_deshydrat_" & I)
ES_desdig = Range("Mat_ES_desdig_" & I)


    'détermination du débit d'eaux sales au global
    Q_eaux_sales = 0
    MES_eaux_sales = 0
    inlet_MV = 0
    For I = 1 To nb_eaux_sales_max
        If choix_eaux_sales(I) = devenir_ES_primaire Then
            'débit d'eaux sales
            Q_eaux_sales = Q_eaux_sales + eaux_sales(I, repere_ES_Q)
            'MES des eaux sales (pour calcul du ratio MV/MES
            MES_eaux_sales = MES_eaux_sales + eaux_sales(I, repere_ES_MES)
            'report de l'origine des boues
            boues_I_detail(I, repere_origine) = eaux_sales(I, repere_ES_origine)
            'calcul des MV de chaque type d'eaux sales
            intermediaire_MV(I) = eaux_sales(I, repere_ES_MV_MES) * eaux_sales(I, repere_ES_MES)
            'calcul des MES de chaque type d'eaux sales
            intermediaire_MES(I) = eaux_sales(I, repere_ES_MES)
            inlet_MV = inlet_MV + intermediaire_MV(I)
        End If
    Next I
    'on ajoute les MV de l'eau brute (que l'on obtiendrait en théorie sans les eaux sales
    intermediaire_MV(repere_MV_primaire) = MV_MES_reference * (MES - MES_eaux_sales)
    inlet_MV = inlet_MV + intermediaire_MV(repere_MV_primaire)
    'on calcule la provenance des MV par type de boues (ratio par rapport aux MV totaux)
    For I = repere_MV_ES1 To repere_MV_ES2
        boues_I_detail(I, repere_ratio_MV) = intermediaire_MV(I) / inlet_MV
        boues_I_detail(I, repere_ratio_MES) = intermediaire_MES(I) / MES
    Next I
    'on calcule le ratio de MV/MES
    If MV_MES_force = True Then
        MV_MES = MV_MES_f
    Else
        MV_MES = inlet_MV / MES
    End If

'charges à considérer pour dimensionnement
Select Case Range("btn_I_simple_bypass_" & I)
    Case "sur 100% du débit"
            Q_traite_nominal = Q
            DCO_traite = DCO
            DBO_traite = DBO
            MES_traite = MES
            NK_traite = NK
            NH4_traite = NH4
            NO3_traite = NO3
            Pt_traite = Pt
            Q_retour_traite = Q_retour       'de même le Q doit être toatal
            Q_eaux_sales_traite = Q_eaux_sales
    Case "sur le % non by-passé"
            Q_traite_nominal = (1 - bypass) * Q
            DCO_traite = (1 - bypass) * DCO
            DBO_traite = (1 - bypass) * DBO
            MES_traite = (1 - bypass) * MES
            NK_traite = (1 - bypass) * NK
            NH4_traite = (1 - bypass) * NH4
            NO3_traite = (1 - bypass) * NO3
            Pt_traite = (1 - bypass) * Pt
            Q_retour_traite = (1 - bypass) * Q_retour
            Q_eaux_sales_traite = (1 - bypass) * Q_eaux_sales
End Select

'calcul de la surface au mirroir
'Mat_dec_I_Vp, Mat_dec_I_Vr, Mat_dec_I_Ve
If Mat_dec_I_Vr(4, 1) = "" Then S_mirroir = (CoefficientPointeHydrauliqueTempsPluie * (Q_traite_nominal - Q_retour_traite - Q_eaux_sales_traite) + Q_retour_traite + Q_eaux_sales_traite) / (NOMBRE_HEURE_PAR_JOUR * V_mirroir_limite)
If Mat_dec_I_Vr(4, 1) <> "" Then S_mirroir = Mat_dec_I_Vr(4, 1)
Mat_dec_I_Ve(4, 1) = S_mirroir

'calcul de la vitesse au mirroir limite recalculée
 V_mirroir_limite_recalculee = (CoefficientPointeHydrauliqueTempsPluie * (Q_traite_nominal - Q_retour_traite - Q_eaux_sales_traite) + Q_retour_traite + Q_eaux_sales_traite) / (NOMBRE_HEURE_PAR_JOUR * S_mirroir)

'calcul des boues à extraire
boues_MES = MES_traite * rdt_MES
boues_Q = boues_MES / boues_concentration


    'correction du rendement sur P
    If rdt_P_force = False Then
        rdt_P = (hyp_rdt_P_NoDiscfilter * (Pt_traite - ES_discfilter(repere_ES_Pt) * Pt_traite / Pt) + (ES_discfilter(repere_ES_Pt) * Pt_traite / Pt)) / Pt_traite
    End If


'calcul des débit et charges en sortie eau
Q = Q - boues_Q
DCO = DCO - DCO_traite * rdt_DCO
DBO = DBO - DBO_traite * rdt_DBO
MES = MES - MES_traite * rdt_MES
NK = NK - NK_traite * rdt_NK
NH4 = NH4 * Q / (Q + boues_Q)
NO3 = NO3 * Q / (Q + boues_Q)
Pt = Pt - Pt_traite * rdt_P





'CONSO ELEC



    'RACLEUR
    'rate_puissance_fixe = 0.1 'Wh/m3 par rapport au nominal
    'rate_puissance_debit = 1 'Wh/m3 par rapport au débit réel
    'Nantes: racleur/herse: 0.38 Wh/m3
    'Angers racleur/herse: 0.88 Wh/m3
    'electricite_racleur = 1 / 1000 * (rate_puissance_fixe * Q_traite_nominal + rate_puissance_debit * Q_traite_reel)
    'DCA le 15/07/11 données Pascal JOUAFFRE  de 12 à 32m diamètre --> 0.55 kW installés, de 33 à 52m --> 0.75 kW
    
    
dec_diametre_limite = 32 'm vu avec P. JOUAFFRE (07/11)
hyp_dec_I_Pw_racleur_sup_Dlim = 0.75 'kW vu avec P. JOUAFFRE (07/11)
hyp_dec_I_Pw_racleur_inf_Dlim = 0.55
    
If S_mirroir / NombreOuvrages < 3.14159 * (dec_diametre_limite / 2) ^ 2 Then
        electricite_racleur = mi_NombreOuvrages * hyp_dec_I_Pw_racleur_inf_Dlim * NOMBRE_HEURE_PAR_JOUR
Else
        electricite_racleur = mi_NombreOuvrages * hyp_dec_I_Pw_racleur_sup_Dlim * NOMBRE_HEURE_PAR_JOUR
End If


















End Sub
















Sub calcul_consommation_electrique()



    'EXTRACTION DES BOUES    ici, le rendement_moteur correspond au rendement machine (mécanique)
    
    'ON LIT LES VALEURS IMPLEMENTEES
    

    'nombre de pompe
    If Mat_dec_I_Vr(21, 1) = "" Then Mat_dec_I_Ve(21, 1) = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) + 1
    If Mat_dec_I_Vr(21, 1) <> "" Then Mat_dec_I_Ve(21, 1) = Mat_dec_I_Vr(21, 1)
    
    extraction_nb_pompe = Mat_dec_I_Ve(21, 1)
    If Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) < (boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) Then extraction_nb_pompe = extraction_nb_pompe + 1

    'durée de fonctionnement
    If Mat_dec_I_Vr(23, 1) = "" Then Mat_dec_I_Ve(23, 1) = boues_Q / (extraction_nb_pompe * extraction_Q_unitaire)
    If Mat_dec_I_Vr(23, 1) <> "" Then Mat_dec_I_Ve(23, 1) = Mat_dec_I_Vr(23, 1)
    extraction_tps_fct_pompe = Mat_dec_I_Ve(23, 1)
    
    If extraction_tps_fct_pompe > NOMBRE_HEURE_PAR_JOUR Then extraction_tps_fct_pompe = NOMBRE_HEURE_PAR_JOUR
    

            
     'extraction_débit pompe unitaire
    If Mat_dec_I_Vr(24, 1) = "" Then Mat_dec_I_Ve(24, 1) = boues_Q / (extraction_nb_pompe * extraction_tps_fct_pompe)
    If Mat_dec_I_Vr(24, 1) <> "" Then Mat_dec_I_Ve(24, 1) = Mat_dec_I_Vr(24, 1)
    extraction_Q_unitaire = Mat_dec_I_Ve(24, 1)

 
 
 
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt = extraction_pompe_rdt_f
    Else
        rendement_pompe = 0.7    'hypothèse
        puissance_pompe = extraction_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * extraction_P_refoulement / rendement_pompe
        If puissance_pompe < 60 Then
            rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
        Else
            rendement_moteur = 0.961
        End If
        extraction_pompe_rdt = rendement_pompe * rendement_moteur
    End If
    electricite_extraction = extraction_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * extraction_P_refoulement / extraction_pompe_rdt * extraction_nb_pompe * extraction_tps_fct_pompe


    'CONSOMMATION GLOBALE D'ELECTRICITE
    electricite = electricite_racleur + electricite_extraction
    electricite_fixe = electricite_racleur

    Exit Sub
    '*********************************************************************

End Sub


  'extraction_débit pompe unitaire
    If Mat_dec_I_Vr(24, 1) = "" Then Mat_dec_I_Ve(24, 1) = boues_Q / (extraction_nb_pompe * extraction_tps_fct_pompe)
    If Mat_dec_I_Vr(24, 1) <> "" Then Mat_dec_I_Ve(24, 1) = Mat_dec_I_Vr(24, 1)
    extraction_Q_unitaire = Mat_dec_I_Ve(24, 1)
    'durée de fonctionnement
    If Mat_dec_I_Vr(23, 1) = "" Then Mat_dec_I_Ve(23, 1) = boues_Q / (extraction_nb_pompe * extraction_Q_unitaire)
    If Mat_dec_I_Vr(23, 1) <> "" Then Mat_dec_I_Ve(23, 1) = Mat_dec_I_Vr(23, 1)
    extraction_tps_fct_pompe = Mat_dec_I_Ve(23, 1)
    If extraction_tps_fct_pompe > NOMBRE_HEURE_PAR_JOUR Then extraction_tps_fct_pompe = NOMBRE_HEURE_PAR_JOUR
    
    
    If Mat_dec_I_Vr(21, 1) = "" Then Mat_dec_I_Ve(21, 1) = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) + 1
    If Mat_dec_I_Vr(21, 1) <> "" Then Mat_dec_I_Ve(21, 1) = Mat_dec_I_Vr(21, 1)





    'EXTRACTION DES BOUES    ici, le rendement_moteur correspond au rendement machine (mécanique)
    If Mat_dec_I_Vr(24, 1) <> "" Then
        extraction_Q_unitaire = Mat_dec_I_Ve(24, 1)
        
        
        If Mat_dec_I_Vr(23, 1) <> "" Then
            'on ne fait rien
            If Mat_dec_I_Vr(21, 1) = "" Then
            
                If Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) < (boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) Then
                    extraction_nb_pompe = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) + 1
                Else
                    extraction_nb_pompe = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire))
                End If
            End If
        Else
            extraction_tps_fct_pompe = boues_Q / (extraction_nb_pompe * extraction_Q_unitaire)
            If extraction_tps_fct_pompe > pd_NOMBRE_HEURE_PAR_JOUR Then
                extraction_tps_fct_pompe = pd_NOMBRE_HEURE_PAR_JOUR
                If extraction_nb_pompe_force = False Then
                    If Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) < (boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) Then
                        extraction_nb_pompe = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire)) + 1
                    Else
                        extraction_nb_pompe = Int(boues_Q / (extraction_tps_fct_pompe * extraction_Q_unitaire))
                    End If
                    extraction_tps_fct_pompe = boues_Q / (extraction_nb_pompe * extraction_Q_unitaire)
                End If
            End If
        End If
    Else
        extraction_Q_unitaire = boues_Q / (extraction_nb_pompe * extraction_tps_fct_pompe)
    End If
    incoherence = (extraction_Q_unitaire * extraction_nb_pompe * extraction_tps_fct_pompe - boues_Q) / boues_Q
    If Abs(incoherence) > critere_incoherence Then
        remarque(1) = Feuil9.Range("dec_simple_erreur_pompe_extraction").Cells(1, 1)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt = extraction_pompe_rdt_f
    Else
        rendement_pompe = 0.7    'hypothèse
        puissance_pompe = extraction_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * extraction_P_refoulement / rendement_pompe
        If puissance_pompe < 60 Then
            rendement_moteur = (5.7195 * Log(puissance_pompe) + 72.682) / 100
        Else
            rendement_moteur = 0.961
        End If
        extraction_pompe_rdt = rendement_pompe * rendement_moteur
    End If
    electricite_extraction = extraction_Q_unitaire / pd_NOMBRE_SECONDE_PAR_HEURE * pd_ACCELERATION_PESANTEUR_m_s2 * extraction_P_refoulement / extraction_pompe_rdt * extraction_nb_pompe * extraction_tps_fct_pompe


    'CONSOMMATION GLOBALE D'ELECTRICITE
    electricite = electricite_racleur + electricite_extraction
    electricite_fixe = electricite_racleur




