Sub Boue_Activee_Forte_Charge()

'hypothèses

T_reference = 12 '°C Température de référence pour la cinétique biologique (°C)
G_reference = 1 'Age de boues à T_reference (j)
correctif_T = 1.072 'Facteur correctif température pour cinétique

'coefficients relation charge massique (kgDBO/kgMVS/j)  1/Cm=a1*G_eq+a0
a1_Cm = 0.8778
a0_Cm = -0.0631

'coefficients relation rendement DBO fonction de Cm_equivalente rdt_DBO=a2*Cm^2+a1*Cm+a0
FtoM_lim_kgDBO_kgMESj = 1.9 'kgDBO/(kgMES.j)
If choix_primaire = True Then Mat_coeff = Array(0.9549, -0.0966, -0.4705, 0.1798)
If choix_primaire = False Then Mat_coeff = Array(0.9854, -0.2053, -0.2472, 0.1104)
                
   
rdt_DCO = 0.6 'hypothèse de rendement sur la DCO
charge_radier_max = Array(5, 7.5, 9) 'Charge radier maximale kg/(m2.h) racleur_racle To racleur_Kruger VU C AMIEL

'hypothèses pour le calcul des besoins en O2
'source ATV relation avec l'âge de boues et la température
'rate_O2_DBOresp = a0 + (a1 * G * corrT ^ (t - Tref)) / (1 + a2 * G * corrT ^ (t - Tref))
    
a0_O2_DBO_respiration = 0.56
a1_O2_DBO_respiration = 0.15
a2_O2_DBO_respiration = 0.17
correctif_T_ATV = 1.072
T_ref_ATV = 15

    'facteurs alpha0 (celui de fines bulles est calculé)
    
Select Case Range("btn_BAforte_diffusion_" & I)
    
    Case "diffuseur moyennes bulles": O2_alpha0 = 0.75
    
    hauteur_bassin = 6
    
    
    Case "brosses d'aération": O2_alpha0 = 0.75
    aerateur_surface_ASB = 1.7 'aerateur_brosse 'aérateur de surface Apport spécifique brut (eau claire)  (kWh/kO2_transféré)   source:fndae 31
    hauteur_bassin = 3
    
    
    Case "turbines lentes": O2_alpha0 = 0.75
    aerateur_surface_ASB = 1.65 'trubine lente 'aérateur de surface Apport spécifique brut (eau claire)  (kWh/kO2_transféré)   source:fndae 31
    hauteur_bassin = 3
    
    Case "turbines rapides": O2_alpha0 = 0.75
    aerateur_surface_ASB = 1.2 'trubine rapide 'aérateur de surface Apport spécifique brut (eau claire)  (kWh/kO2_transféré)   source:fndae 31
    hauteur_bassin = 3
    
    
    Case "diffuseur fines bulles": O2_alpha0 = 0.51
        'relation calcul alpha fines bulles    ao+a_MVS*MVS+a_G*G_eq
        a_alpha_MVS = -0.062
        a_alpha_G = 0.019
        O2_facteur_alfa_maxi = 1
        O2_facteur_alfa_mini = 0.05
    
    
    hauteur_bassin = 6
    
    
    
End Select

'correction alpha en fonction des MES  alfa=alfa*(a0+a1*MES_reel)/(a0+a1*MES_lim)
a0_correctif_alfa_MES = 1.0007
a1_correctif_alfa_MES = -0.035 'L/g
MES_reference_alfa = 5 'g/L

'facteur beta
O2_facteur_beta = 0.95  'compris entre 0,95 et 0,99

'calcul de la pression atmosphérique  Patm=P0*((a0-a1*altitude)^a2)/100
P0 = 101325
a0_Patm = 1
a1_Patm = -0.0000225577
a2_Patm = 5.25588

'calcul de la surpression en diffusion de bulles  delta_P=a0/a1*P0/100*(hauteur_bassin-hauteur_diffuseur)
a0_delta_P = 0.35
a1_delta_P = 10.33
hauteur_diffuseur = 0.25 'm

'calcul O2 saturation  O2_saturation(T,P0)= somme(ai*Teau^i)
a_O2_saturation = Array(14.628, -0.4223, 0.0102, -0.0002, 0.000001)

'correction température pour facteur K
correctif_T_K = 1.024
T_ref_K = 20

ratio_kgO2_Nm3air = 0.3 'O2 présent dans l'air (kgO2/Nm3air)

'rendement de dissolution en eau claire (%/m)
O2_rdt_dissolution(aerateur_fine_bulle) = Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(3, 1)
O2_rdt_dissolution(aerateur_moyenne_bulle) = Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(2, 1)

'consommations spécifiques par type de surpresseurs      en Wh/(Nm3.mCE)
Select Case Range("btn_BAforte_surpresseur_" & I)
    Case "tri-lobes type roots": surpresseur_conso_Wh_Nm3mCE = 4.5
    Case "surpresseurs à vis": surpresseur_conso_Wh_Nm3mCE = 3.8
    Case "centrifuge type turbo": surpresseur_conso_Wh_Nm3mCE = 3.5
End Select

    
'pression max pour roots
hyp_Pmax_conseillee = 8 'mce

'impact du type de régulation
Select Case Range("btn_BAforte_regulation_" & I)

    Case "régulation sur horloge"
        If TemperatureEauConditionsReelles_°C < 14 Then ConcentrationOxygeneDissous_mgL = 4
        If TemperatureEauConditionsReelles_°C >= 14 Then ConcentrationOxygeneDissous_mgL = 4
    
    Case "régulation sans variateur de fréquence"
        If TemperatureEauConditionsReelles_°C < 14 Then ConcentrationOxygeneDissous_mgL = 2
        If TemperatureEauConditionsReelles_°C >= 14 Then ConcentrationOxygeneDissous_mgL = 1.5
    
    Case "classique - avec variateur de fréquence"
        If TemperatureEauConditionsReelles_°C < 14 Then ConcentrationOxygeneDissous_mgL = 1.5
        If TemperatureEauConditionsReelles_°C >= 14 Then ConcentrationOxygeneDissous_mgL = 1
    
    Case "avancé"
        If TemperatureEauConditionsReelles_°C < 14 Then ConcentrationOxygeneDissous_mgL = 1.5
        If TemperatureEauConditionsReelles_°C >= 14 Then ConcentrationOxygeneDissous_mgL = 1

End Select


'clarificateur
rate_puissance_fixe = 0.1    'Wh/m3 par rapport au nominal
rate_puissance_debit = 1    'Wh/m3 par rapport au débit réel

ratio_P_synthese = 0.01 'kgP/kgDBOéliminée 'P consommé dans synthèse






End Sub



Sub dimensionnement(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef Sh, ByRef DCO_conso_bio)

    'DIMENSIONNEMENT DES BASSINS

    'Age de boues equivalent à T_reference
    G_equivalent = nominal_G * correctif_T ^ (pd_TemperatureEauDimensionnement_°C - T_reference)
    'Charge massique équivalente
    Cm_equivalente = 1 / (a1_Cm * G_equivalent + a0_Cm)
    'Charge massique à T_design
    Cm = Cm_equivalente * correctif_T ^ (pd_TemperatureEauDimensionnement_°C - T_reference)
    'Masse de boues dans les bassins kgMVS
    boues_bassins_MVS = DBO / Cm

    'Volume de bassins m3
    volume_bassins = boues_bassins_MVS / (nominal_MV_MES * nominal_MES_concentration_bassin)
    If volume_bassins_force = True Then
        volume_bassins = volume_bassins_f
        'calcul de la Cm_equivalente avec ce volume de bassins
        boues_bassins_MVS = volume_bassins * nominal_MV_MES * nominal_MES_concentration_bassin
        Cm = DBO / boues_bassins_MVS
        Cm_equivalente = Cm / correctif_T ^ (pd_TemperatureEauDimensionnement_°C - T_reference)
    End If

    'surface de radier du clarificateur (m2)
    clarif_surface = (Q_nominal * pd_CoefficientPointeHydrauliqueTempsPluie + (Q - Q_nominal)) / (pd_NOMBRE_HEURE_PAR_JOUR * clarif_vitesse_max)
    If clarif_surface_force = True Then
        clarif_surface = clarif_surface_f
    End If


    'FONCTIONNEMENT NOMINAL

    'rendement d'élimination de la DBO
    'rdt_DBO = a2_rdt_DBO * Cm_equivalente ^ 2 + a1_rdt_DBO * Cm_equivalente + a0_rdt_DBO
    If (Cm_equivalente / nominal_MV_MES) > FtoM_lim_kgDBO_kgMESj Then
        rdt_DBO = 0
        For I = 0 To 3
            rdt_DBO = rdt_DBO + a_rdt_DBO(I) * (FtoM_lim_kgDBO_kgMESj) ^ I
        Next I
    Else
        rdt_DBO = 0
        For I = 0 To 3
            rdt_DBO = rdt_DBO + a_rdt_DBO(I) * (Cm_equivalente / nominal_MV_MES) ^ I
        Next I
    End If

    'Production de boues
    boues_produites = boues_bassins_MVS / (nominal_MV_MES * nominal_G)

    'ratio O2 respiration et élimination DBO5
    facteur_T_O2 = correctif_T_ATV ^ (pd_TemperatureEauConditionsReelles_°C - T_ref_ATV)
    ratio_O2_DBO_respiration = a0_O2_DBO_respiration + (a1_O2_DBO_respiration * reel_G * facteur_T_O2) / (1 + a2_O2_DBO_respiration * reel_G * facteur_T_O2)
    'calcul des besoins en O2
    O2_besoin = ratio_O2_DBO_respiration * rdt_DBO * DBO
    O2_besoin = O2_besoin + besoins_O2_HS(Sh)
    DCO_conso_bio = O2_besoin

    'vitesse hydraulique max (recalculée) clarificateur
    clarif_vitesse_max_recalc = (Q_nominal * pd_CoefficientPointeHydrauliqueTempsPluie + (Q - Q_nominal)) / (pd_NOMBRE_HEURE_PAR_JOUR * clarif_surface)
    'charge au radier clarificateur
    clarif_charge_radier_nominal = nominal_MES_concentration_bassin * Q * (1 + recirculation_taux) / (pd_NOMBRE_HEURE_PAR_JOUR * clarif_surface)
    If clarif_charge_radier_nominal > charge_radier_max(choix_racleur) Then
        remarque(1) = Feuil11.Range("BA_forte_erreur_charge_radier").Cells(1, 1)
    Else
        remarque(1) = ""
    End If


    'CALCUL DES SORTIES

    'Eau traitée
    Q = Q - boues_Q
    MES = sortie_MES * Q / 1000
    DCO = DCO * (1 - rdt_DCO)
    Pt = Pt - ratio_P_synthese * DBO * rdt_DBO
    DBO = DBO * (1 - rdt_DBO)
    NK = NK
    NH4 = NH4

    'Boues
    boues_MES = boues_produites - MES
    boues_concentration = nominal_MES_concentration_bassin * (1 + recirculation_taux) / recirculation_taux
    boues_Q = boues_MES / boues_concentration

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "11.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub fonctionnement_reel(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef Sh, ByRef DCO_conso_bio)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer

    'variables internes
    Dim G_equivalent As Double
    Dim Cm_equivalente As Double
    Dim Cm As Double
    Dim boues_bassins_MVS As Double
    Dim rdt_DBO As Double
    Dim boues_produites As Double
    Dim facteur_T_O2 As Double
    Dim ratio_O2_DBO_respiration As Double

    stockage_Q = Q


    'FONCTIONNEMENT REEL

    'Age de boues equivalent à T_reference
    G_equivalent = reel_G * correctif_T ^ (pd_TemperatureEauConditionsReelles_°C - T_reference)
    'Charge massique équivalente
    Cm_equivalente = 1 / (a1_Cm * G_equivalent + a0_Cm)
    'Charge massique à T_design
    Cm = Cm_equivalente * correctif_T ^ (pd_TemperatureEauConditionsReelles_°C - T_reference)
    'Masse de boues dans les bassins kgMVS
    boues_bassins_MVS = DBO / Cm
    'Calcul de la concentration en MES dans les bassins
    reel_MES_concentration_bassin = boues_bassins_MVS / (reel_MV_MES * volume_bassins)
    If reel_MES_concentration_bassin_force = True Then
        reel_MES_concentration_bassin = reel_MES_concentration_bassin_f
        'calcul de la Cm équivalente avec cette concentration
        boues_bassins_MVS = volume_bassins * reel_MV_MES * reel_MES_concentration_bassin
        Cm = DBO / boues_bassins_MVS
        Cm_equivalente = Cm / correctif_T ^ (pd_TemperatureEauConditionsReelles_°C - T_reference)
    End If

    stockage_MVS_bassins = boues_bassins_MVS

    'TRAITEMENT REEL

    'rendement d'élimination de la DBO
    'rdt_DBO = a2_rdt_DBO * Cm_equivalente ^ 2 + a1_rdt_DBO * Cm_equivalente + a0_rdt_DBO
    If (Cm_equivalente / reel_MV_MES) > FtoM_lim_kgDBO_kgMESj Then
        rdt_DBO = 0
        For I = 0 To 3
            rdt_DBO = rdt_DBO + a_rdt_DBO(I) * (FtoM_lim_kgDBO_kgMESj) ^ I
        Next I
    Else
        rdt_DBO = 0
        For I = 0 To 3
            rdt_DBO = rdt_DBO + a_rdt_DBO(I) * (Cm_equivalente / reel_MV_MES) ^ I
        Next I
    End If
    sortie_DBO = DBO * (1 - rdt_DBO) / Q * 1000
    If sortie_DBO_force = True Then
        sortie_DBO = sortie_DBO_f
    End If

    'Production de boues
    boues_produites = boues_bassins_MVS / (reel_MV_MES * reel_G)

    'ratio O2 respiration et élimination DBO5
    facteur_T_O2 = correctif_T_ATV ^ (pd_TemperatureEauConditionsReelles_°C - T_ref_ATV)
    ratio_O2_DBO_respiration = a0_O2_DBO_respiration + (a1_O2_DBO_respiration * reel_G * facteur_T_O2) / (1 + a2_O2_DBO_respiration * reel_G * facteur_T_O2)
    'calcul des besoins en O2
    O2_besoin = ratio_O2_DBO_respiration * (DBO - sortie_DBO * stockage_Q / 1000)
    besoins_O2_respiration = ((a1_O2_DBO_respiration * reel_G * facteur_T_O2) / (1 + a2_O2_DBO_respiration * reel_G * facteur_T_O2)) * (DBO - sortie_DBO * stockage_Q / 1000)
    O2_besoin = O2_besoin + besoins_O2_HS(Sh)
    If O2_besoin_force = True Then
        O2_besoin = O2_besoin_f
    End If
    'DCO_conso_bio = O2_besoin

    'CALCUL DES SORTIES

    'Eau traitée
    Q = Q - boues_Q
    MES = sortie_MES * Q / 1000
    DCO = DCO * (1 - rdt_DCO)
    Pt = Pt - ratio_P_synthese * (DBO - sortie_DBO * Q / 1000)
    DBO = sortie_DBO * Q / 1000
    NK = NK
    NH4 = NH4

    'Boues
    boues_MES = boues_produites - MES
    If boues_MES_force = True Then
        boues_MES = boues_MES_f
    End If
    boues_concentration = reel_MES_concentration_bassin * (1 + recirculation_taux) / recirculation_taux
    If boues_concentration_force = True Then
        boues_concentration = boues_concentration_f
    End If
    boues_Q = boues_MES / boues_concentration


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "11.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub calcul_consommation_electrique(ByRef electricite_II_aeration)
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    Dim I As Integer

    Dim clarif_diametre_limite As Double

    'variables pour l'aération
    Dim gamme_P As Integer
    Dim gamme_Q As Integer
    Dim facteur_T_O2 As Double
    Dim ratio_O2_DBO_respiration As Double
    Dim Patm As Double
    Dim delta_P As Double
    Dim O2_saturation_20_P As Double
    Dim O2_saturation_T_P As Double
    Dim O2_facteur_K As Double
    Dim electricite_racleur As Double
    Dim electricite_recirculation As Double
    Dim electricite_extraction As Double



    'AERATION

    'calcul facteur alpha
    If choix_aerateur = aerateur_fine_bulle Then
        'source: Oxygen transfer in AS (Henkel,Cornel and Wagner)
        O2_facteur_alfa = a0_alpha + (a_alpha_MVS * (stockage_MVS_bassins / volume_bassins)) + (a_alpha_G * (reel_G * correctif_T ^ (pd_TemperatureEauConditionsReelles_°C - T_reference)))
        If O2_facteur_alfa > O2_facteur_alfa_maxi Then
            O2_facteur_alfa = O2_facteur_alfa_maxi
        ElseIf O2_facteur_alfa < O2_facteur_alfa_mini Then
            O2_facteur_alfa = O2_facteur_alfa_mini
        End If
    Else
        O2_facteur_alfa = O2_alpha0(choix_aerateur)
        If reel_MES_concentration_bassin > MES_reference_alfa Then
            O2_facteur_alfa = O2_facteur_alfa * (a0_correctif_alfa_MES + a1_correctif_alfa_MES * reel_MES_concentration_bassin) / (a0_correctif_alfa_MES + a1_correctif_alfa_MES * MES_reference_alfa)
        End If
    End If
    If O2_facteur_alfa_force = True Then
        O2_facteur_alfa = O2_facteur_alfa_f
    End If

    'calcul O2 saturation à (T,P=Patm+delta_P) et (20,P0)
    Patm = P0 * ((a0_Patm + a1_Patm * altitude) ^ a2_Patm) / 100    'en hPa
    If choix_aerateur = aerateur_fine_bulle Or choix_aerateur = aerateur_moyenne_bulle Then
        delta_P = a0_delta_P / a1_delta_P * P0 / 100 * (hauteur_bassin - hauteur_diffuseur)
    Else
        delta_P = 0
    End If
    O2_saturation_20_P = 0
    O2_saturation_T_P = 0
    For I = 0 To 4
        O2_saturation_20_P = O2_saturation_20_P + a_O2_saturation(I) * 20 ^ I
        O2_saturation_T_P = O2_saturation_T_P + a_O2_saturation(I) * pd_TemperatureEauConditionsReelles_°C ^ I
    Next I
    O2_saturation_T_P = O2_saturation_T_P * (Patm + delta_P) / (P0 / 100)
    O2_saturation_20_P = O2_saturation_20_P * (Patm + delta_P) / (P0 / 100)

    'calcul facteur K
    O2_facteur_K = O2_facteur_alfa * (O2_facteur_beta * O2_saturation_T_P - O2_dissous) / O2_saturation_20_P
    O2_facteur_K = O2_facteur_K * correctif_T_K ^ (pd_TemperatureEauConditionsReelles_°C - T_ref_K)

    'calcul consommations sur l'aération
    If choix_aerateur = aerateur_fine_bulle Or choix_aerateur = aerateur_moyenne_bulle Then
        'débit d'air
        O2_rdt_transfert = O2_rdt_dissolution(choix_aerateur)
        If O2_rdt_transfert_force = True Then
            O2_rdt_transfert = O2_rdt_transfert_f
        End If
        air_Q_Nm3j = O2_besoin / (O2_facteur_K * O2_rdt_transfert / 100 * (hauteur_bassin - hauteur_diffuseur) * ratio_kgO2_Nm3air)
        If air_Q_Nm3j_force = True Then
            air_Q_Nm3j = air_Q_Nm3j_f
        End If
        If choix_surpresseur = surpresseur_roots And air_P_refoulement > hyp_Pmax_conseillee Then
            remarque(2) = Feuil11.Range("BA_forte_erreur_choix_surpresseur").Cells(1, 1)
        End If
        If surpresseur_conso_spec_force = True Then
            surpresseur_conso_spec = surpresseur_conso_spec_f
            'gamme de pression
            If air_P_refoulement <= Px1 Then
                gamme_P = surpresseur_gamme_P_0_Px1
            ElseIf air_P_refoulement <= Px2 Then
                gamme_P = surpresseur_gamme_P_Px1_Px2
            Else
                gamme_P = surpresseur_gamme_P_Px2_infini
            End If
        Else
            'gamme de pression
            If air_P_refoulement <= Px1 Then
                gamme_P = surpresseur_gamme_P_0_Px1
            ElseIf air_P_refoulement <= Px2 Then
                gamme_P = surpresseur_gamme_P_Px1_Px2
            Else
                gamme_P = surpresseur_gamme_P_Px2_infini
            End If
            'gamme de débit   (pour le nominal)
            gamme_Q = surpresseur_gamme_Q_0_infini
'            surpresseur_conso_spec = surpresseur_conso_Wh_Nm3mCE(choix_surpresseur, gamme_Q, gamme_P) * (1 + impact_regulation(choix_regulation))
            surpresseur_conso_spec = surpresseur_conso_Wh_Nm3mCE(choix_surpresseur, gamme_Q, gamme_P)
        End If
        'calcul de la consommation électrique de l'aération
        electricite_aeration = air_Q_Nm3j * air_P_refoulement * surpresseur_conso_spec / 1000
    Else
        'calcul de la consommation électrique de l'aération
        electricite_aeration = O2_besoin / (ASB_eau_claire * O2_facteur_K)
    End If
    electricite_II_aeration = electricite_II_aeration + electricite_aeration

    'CLARIFICATEUR
    'electricite_racleur = rate_puissance_fixe * pd_NOMBRE_HEURE_PAR_JOUR + rate_puissance_debit * stockage_Q / 1000
    'DCA le 15/07/11 données Pascal JOUAFFRE  de 12 à 32m diamètre --> 0.55 kW installés, de 33 à 52m --> 0.75 kW
    clarif_diametre_limite = 32
    If clarif_surface / mi_NombreClarificateurs < pd_NOMBRE_PI * (clarif_diametre_limite / 2) ^ 2 Then
        electricite_racleur = mi_NombreClarificateurs * 0.55 * pd_NOMBRE_HEURE_PAR_JOUR
    Else
        electricite_racleur = mi_NombreClarificateurs * 0.75 * pd_NOMBRE_HEURE_PAR_JOUR
    End If


    'POMPES DE RECIRCULATION ET EXTRACTION
    electricite_recirculation = ratio_elec_recirculation * recirculation_taux * stockage_Q * recirculation_P_refoulement
    electricite_extraction = ratio_elec_extraction * boues_Q * extraction_P_refoulement


    'CONSOMMATION GLOBALE D'ELECTRICITE
    electricite = electricite_aeration + electricite_racleur + electricite_recirculation + electricite_extraction
    electricite_fixe = besoins_O2_respiration / O2_besoin * electricite_aeration + electricite_racleur

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "11.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub





Private Sub attribution_valeur_par_defaut()

'''''''''''''''''''''''''''''
'VALEURS PAR DEFAUTS


'les matrices renseignées
mat_BAFC_nominal_Vp = Range("mat_BAFC_nominal_Vp_" & I)
mat_BAFC_nominal_Vr = Range("mat_BAFC_nominal_Vr_" & I)
mat_BAFC_nominal_Ve = Range("mat_BAFC_nominal_Ve_" & I)

'Age de boues design (j)
mat_BAFC_nominal_Vp(1, 1) = G_reference * correctif_T ^ (T_reference - TemperatureEauDimensionnement_°C)
If mat_BAFC_nominal_Vr(1, 1) = "" Then nominal_G = mat_BAFC_nominal_Vp(1, 1)
If mat_BAFC_nominal_Vr(1, 1) <> "" Then nominal_G = mat_BAFC_nominal_Vr(1, 1)
mat_BAFC_nominal_Ve(1, 1) = nominal_G

'Concentration MES design (dans les bassins) g/L
If TemperatureEauDimensionnement_°C < 12 Then mat_BAFC_nominal_Vp(2, 1) = 3 'nominal_MES_concentration_bassin
If TemperatureEauDimensionnement_°C < 18 Then mat_BAFC_nominal_Vp(2, 1) = 2.5
If TemperatureEauDimensionnement_°C >= 18 Then mat_BAFC_nominal_Vp(2, 1) = 1.5

If mat_BAFC_nominal_Vr(2, 1) = "" Then nominal_MES_concentration_bassin = mat_BAFC_nominal_Vp(2, 1)
If mat_BAFC_nominal_Vr(2, 1) <> "" Then nominal_MES_concentration_bassin = mat_BAFC_nominal_Vr(2, 1)
mat_BAFC_nominal_Ve(2, 1) = nominal_MES_concentration_bassin


'MV/MES des boues design
If TemperatureEauDimensionnement_°C < 12 Then mat_BAFC_nominal_Vp(3, 1) = 0.7 'MV/MES des boues design
If TemperatureEauDimensionnement_°C < 18 Then mat_BAFC_nominal_Vp(3, 1) = 0.72
If TemperatureEauDimensionnement_°C >= 18 Then mat_BAFC_nominal_Vp(3, 1) = 0.75

If mat_BAFC_nominal_Vr(3, 1) = "" Then nominal_MV_MES = mat_BAFC_nominal_Vp(3, 1)
If mat_BAFC_nominal_Vr(3, 1) <> "" Then nominal_MV_MES = mat_BAFC_nominal_Vr(3, 1)
mat_BAFC_nominal_Ve(3, 1) = nominal_MV_MES


'les matrices renseignées
mat_BAFC_reel_Vp = Range("mat_BAFC_reel_Vp_" & I)
mat_BAFC_reel_Vr = Range("mat_BAFC_reel_Vr_" & I)
mat_BAFC_reel_Ve = Range("mat_BAFC_reel_Ve_" & I)
 
 
 
 
    
'Age de boues réel (j)
G_equivalent_nominal_G = nominal_G / correctif_T ^ (T_reference - TemperatureEauDimensionnement_°C)
reel_G = G_equivalent_nominal_G * correctif_T ^ (T_reference - TemperatureEauConditionsReelles_°C)

mat_BAFC_reel_Vp(1, 1) = reel_G
If mat_BAFC_reel_Vr(1, 1) = "" Then reel_G = mat_BAFC_reel_Vp(1, 1)
If mat_BAFC_reel_Vr(1, 1) <> "" Then reel_G = mat_BAFC_reel_Vr(1, 1)
mat_BAFC_reel_Ve(1, 1) = reel_G


 'MV/MES des boues réel
If TemperatureEauDimensionnement_°C < 12 Then mat_BAFC_reel_Vp(2, 1) = 0.7 'MV/MES des boues design
If TemperatureEauDimensionnement_°C < 18 Then mat_BAFC_reel_Vp(2, 1) = 0.72
If TemperatureEauDimensionnement_°C >= 18 Then mat_BAFC_reel_Vp(2, 1) = 0.75

If mat_BAFC_reel_Vr(2, 1) = "" Then reel_MV_MES = mat_BAFC_reel_Vp(2, 1)
If mat_BAFC_reel_Vr(2, 1) <> "" Then reel_MV_MES = mat_BAFC_reel_Vr(2, 1)
mat_BAFC_reel_Ve(2, 1) = reel_MV_MES


'Concentration MES rrel (dans les bassins) g/L
If TemperatureEauDimensionnement_°C < 12 Then mat_BAFC_reel_Vp(3, 1) = 3 'reel_MES_concentration_bassin
If TemperatureEauDimensionnement_°C < 18 Then mat_BAFC_reel_Vp(3, 1) = 2.5
If TemperatureEauDimensionnement_°C >= 18 Then mat_BAFC_reel_Vp(3, 1) = 1.5

If mat_BAFC_reel_Vr(3, 1) = "" Then reel_MES_concentration_bassin = mat_BAFC_reel_Vp(3, 1)
If mat_BAFC_reel_Vr(3, 1) <> "" Then reel_MES_concentration_bassin = mat_BAFC_reel_Vr(3, 1)
mat_BAFC_reel_Ve(3, 1) = reel_MES_concentration_bassin

'Concentration moyenne en O2 dissous (mg/L)
mat_BAFC_reel_Vp(4, 1) = ConcentrationOxygeneDissous_mgL
If mat_BAFC_reel_Vr(4, 1) = "" Then O2_dissous = mat_BAFC_reel_Vp(4, 1)
If mat_BAFC_reel_Vr(4, 1) <> "" Then O2_dissous = mat_BAFC_reel_Vr(4, 1)
mat_BAFC_reel_Ve(4, 1) = O2_dissous



'Concentration en DBO5

'à compléter par la suite lors du calcul



'type de surpresseur





    'Durée depuis le dernier nettoyage des diffuseurs
    If diffuseur_encrassement_force = True Then
        diffuseur_encrassement = diffuseur_encrassement_f
    Else
        diffuseur_encrassement = 0
    End If

    'Pression de refoulement des surpresseurs (mCE)
    If air_P_refoulement_force = True Then
        air_P_refoulement = air_P_refoulement_f
    Else
        air_P_refoulement = hauteur_bassin + 2 + 0.25 * diffuseur_encrassement
    End If





    'ASB
    If ASB_eau_claire_force = True Then
        ASB_eau_claire = ASB_eau_claire_f
    Else
        ASB_eau_claire = ASB(choix_aerateur)
    End If

    'Taux de recirculation des boues du clarif
    If recirculation_taux_force = True Then
        recirculation_taux = recirculation_taux_f
    Else
        recirculation_taux = 1
    End If

    'Pression des pompes de recirculation des boues du clarif (mCE)
    If recirculation_P_refoulement_force = True Then
        recirculation_P_refoulement = recirculation_P_refoulement_f
    Else
        recirculation_P_refoulement = 5
    End If

    'recirculation des boues du clarif
    If recirculation_pompe_rdt_force = True Then
        recirculation_pompe_rdt = recirculation_pompe_rdt_f
    Else
        rdt_pompe = 0.7
        rdt_moteur_pompe = 0.88
        recirculation_pompe_rdt = rdt_pompe * rdt_moteur_pompe
    End If
    ratio_elec_recirculation = pd_ACCELERATION_PESANTEUR_m_s2 / (pd_NOMBRE_SECONDE_PAR_HEURE * recirculation_pompe_rdt)    ' kWh/(m3*mCE)

    'Nombre de clarificateurs
    If mn_NombreClarificateurs_force = True Then
        mi_NombreClarificateurs = mi_NombreClarificateurs_f
    Else
        mi_NombreClarificateurs = 1
    End If
    
    'Indice de Mohlman (mL/g)
    If boues_indice_Mohlman_force = True Then
        boues_indice_Mohlman = boues_indice_Mohlman_f
    Else
        If choix_primaire = True Then
            boues_indice_Mohlman = 250
        Else
            boues_indice_Mohlman = 80
        End If
    End If

    'Sortie MES eau traitée (mg/L)
    If sortie_MES_force = True Then
        sortie_MES = sortie_MES_f
    Else
        sortie_MES = 50
    End If

    'Hauteur du clarificateur (m)
    If clarif_hauteur_force = True Then
        clarif_hauteur = clarif_hauteur_f
    Else
        clarif_hauteur = 4
    End If

    'Vitesse hydraulique maximum (m/h)
    If clarif_vitesse_max_force = True Then
        clarif_vitesse_max = clarif_vitesse_max_f
    Else
        volume_corrige = boues_indice_Mohlman * nominal_MES_concentration_bassin
        clarif_vitesse_max = (100 * clarif_hauteur * (sortie_MES / 3.15) ^ 0.5) / ((1 + recirculation_taux) * volume_corrige)
    End If

    'Pression de refoulement de l'extraction (mCE)
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement = extraction_P_refoulement_f
    Else
        extraction_P_refoulement = 5
    End If

    'extraction des boues du clarif
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt = extraction_pompe_rdt_f
    Else
        rdt_pompe = 0.7
        rdt_moteur_pompe = 0.88
        extraction_pompe_rdt = rdt_pompe * rdt_moteur_pompe
    End If
    ratio_elec_extraction = pd_ACCELERATION_PESANTEUR_m_s2 / (pd_NOMBRE_SECONDE_PAR_HEURE * extraction_pompe_rdt)    ' kWh/(m3*mCE)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "11.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub







