Option Explicit

'constantes pour le choix du type d'aérateur
Private Const nb_type_aerateur = 5  'nombre de type d'aérateurs
Private Const aerateur_fine_bulle = 1
Private Const aerateur_moyenne_bulle = 2
Private Const aerateur_brosse = 3
Private Const aerateur_turbine_lente = 4
Private Const aerateur_turbine_rapide = 5

'constantes pour le choix du type de régulation d'air
Private Const nb_type_regulation = 4  'nombre de type d'aérateurs
Private Const regulation_horloge = 1
Private Const regulation_sans_variation = 2
Private Const regulation_avec_variation = 3
Private Const regulation_avance = 4

'constantes pour le choix du type de surpresseur
Private Const nb_type_surpresseur = 3
Private Const surpresseur_roots = 1
Private Const surpresseur_vis = 2
Private Const surpresseur_turbo = 3
Private Const nb_surpresseur_gamme_Q = 1
Private Const surpresseur_gamme_Q_0_infini = 1
Private Const nb_surpresseur_gamme_P = 3
Private Const surpresseur_gamme_P_0_Px1 = 1
Private Const Px1 = 6    'mCE
Private Const surpresseur_gamme_P_Px1_Px2 = 2
Private Const Px2 = 8    'mCE
Private Const surpresseur_gamme_P_Px2_infini = 3

'constantes pour le choix du type de râcleur
Private Const nb_type_racleur = 3  'nombre de type de râcleurs
Private Const racleur_racle = 1
Private Const racleur_racle_suce = 2
Private Const racleur_Kruger = 3

Private Const critere_incoherence = 0.1

'variables de choix internes
Dim choix_aerateur As Integer
Dim choix_regulation As Integer
Dim choix_surpresseur As Integer
Dim choix_racleur As Integer

'variables échangées avec Excel
Dim nominal_G As Double
Dim nominal_MES_concentration_bassin As Double
Public nominal_MV_MES As Double
Dim volume_bassins As Double
Dim reel_G As Double
Public reel_MV_MES As Double
Dim reel_MES_concentration_bassin As Double
Dim O2_dissous As Double
Dim sortie_DBO As Double
Dim hauteur_bassin As Double
Dim O2_besoin As Double
Dim O2_rdt_transfert As Double
Dim O2_facteur_alfa As Double
Dim air_Q_Nm3j As Double
Dim diffuseur_encrassement As Double
Dim air_P_refoulement As Double
Dim surpresseur_conso_spec As Double
Dim ASB_eau_claire As Double
Dim recirculation_taux As Double
Dim recirculation_P_refoulement As Double
Dim recirculation_pompe_rdt As Double
Dim mi_NombreClarificateurs As Integer
Dim boues_indice_Mohlman As Double
Dim sortie_MES As Double
Dim clarif_hauteur As Double
Dim clarif_vitesse_max As Double
Dim clarif_surface As Double
Dim clarif_vitesse_max_recalc As Double
Dim clarif_charge_radier_nominal As Double
Dim boues_concentration As Double
Public boues_MES As Double
Public boues_Q As Double
Dim extraction_P_refoulement As Double
Dim extraction_pompe_rdt As Double
Public electricite As Double
Dim ratio_P_synthese As Double
Dim electricite_aeration As Double

Public electricite_fixe As Double
Dim besoins_O2_respiration As Double

'variables montrant si les valeurs ont été forcées
Dim nominal_G_force As Boolean
Dim nominal_MES_concentration_bassin_force As Boolean
Dim nominal_MV_MES_force As Boolean
Dim volume_bassins_force As Boolean
Dim reel_G_force As Boolean
Dim reel_MV_MES_force As Boolean
Dim reel_MES_concentration_bassin_force As Boolean
Dim O2_dissous_force As Boolean
Dim sortie_DBO_force As Boolean
Dim hauteur_bassin_force As Boolean
Dim O2_besoin_force As Boolean
Dim O2_rdt_transfert_force As Boolean
Dim O2_facteur_alfa_force As Boolean
Dim air_Q_Nm3j_force As Boolean
Dim diffuseur_encrassement_force As Boolean
Dim air_P_refoulement_force As Boolean
Dim surpresseur_conso_spec_force As Boolean
Dim ASB_eau_claire_force As Boolean
Dim recirculation_taux_force As Boolean
Dim recirculation_P_refoulement_force As Boolean
Dim recirculation_pompe_rdt_force As Boolean
Dim mn_NombreClarificateurs_force As Boolean
Dim boues_indice_Mohlman_force As Boolean
Dim sortie_MES_force As Boolean
Dim clarif_hauteur_force As Boolean
Dim clarif_vitesse_max_force As Boolean
Dim clarif_surface_force As Boolean
Dim boues_concentration_force As Boolean
Dim boues_MES_force As Boolean
Dim boues_Q_force As Boolean
Dim extraction_P_refoulement_force As Boolean
Dim extraction_pompe_rdt_force As Boolean

'valeurs forcées par l'utilisateur
Dim nominal_G_f As Double
Dim nominal_MES_concentration_bassin_f As Double
Dim nominal_MV_MES_f As Double
Dim volume_bassins_f As Double
Dim reel_G_f As Double
Dim reel_MV_MES_f As Double
Dim reel_MES_concentration_bassin_f As Double
Dim O2_dissous_f As Double
Dim sortie_DBO_f As Double
Dim hauteur_bassin_f As Double
Dim O2_besoin_f As Double
Dim O2_rdt_transfert_f As Double
Dim O2_facteur_alfa_f As Double
Dim air_Q_Nm3j_f As Double
Dim diffuseur_encrassement_f As Double
Dim air_P_refoulement_f As Double
Dim surpresseur_conso_spec_f As Double
Dim ASB_eau_claire_f As Double
Dim recirculation_taux_f As Double
Dim recirculation_P_refoulement_f As Double
Dim recirculation_pompe_rdt_f As Double
Dim mi_NombreClarificateurs_f As Integer
Dim boues_indice_Mohlman_f As Double
Dim sortie_MES_f As Double
Dim clarif_hauteur_f As Double
Dim clarif_vitesse_max_f As Double
Dim clarif_surface_f As Double
Dim boues_concentration_f As Double
Dim boues_MES_f As Double
Dim boues_Q_f As Double
Dim extraction_P_refoulement_f As Double
Dim extraction_pompe_rdt_f As Double

'variables des hypothèses
Dim T_reference As Double
Dim G_reference As Double
Dim correctif_T As Double
Dim a1_Cm As Double
Dim a0_Cm As Double
Dim a_rdt_DBO(3) As Double
Dim FtoM_lim_kgDBO_kgMESj As Double
Dim rdt_DCO As Double
Dim charge_radier_max(nb_type_racleur) As Double
Dim a0_O2_DBO_respiration As Double
Dim a1_O2_DBO_respiration As Double
Dim a2_O2_DBO_respiration As Double
Dim T_ref_ATV As Double
Dim correctif_T_ATV As Double
Dim O2_alpha0(nb_type_aerateur) As Double
Dim a0_alpha As Double
Dim a_alpha_MVS As Double
Dim a_alpha_G As Double
Dim O2_facteur_alfa_maxi As Double
Dim O2_facteur_alfa_mini As Double
Dim a0_correctif_alfa_MES As Double
Dim a1_correctif_alfa_MES As Double
Dim MES_reference_alfa As Double
Dim O2_facteur_beta As Double
Dim P0 As Double
Dim a0_Patm As Double
Dim a1_Patm As Double
Dim a2_Patm As Double
Dim a0_delta_P As Double
Dim a1_delta_P As Double
Dim hauteur_diffuseur As Double
Dim a_O2_saturation(4) As Double
Dim correctif_T_K As Double
Dim T_ref_K As Double
Dim ratio_kgO2_Nm3air As Double
Dim O2_rdt_dissolution(nb_type_aerateur) As Double
Dim ASB(nb_type_aerateur) As Double
'Dim impact_regulation(nb_type_regulation) As Double
Dim md_ConcentrationOxygeneDissous_mgL(1 To nb_type_regulation) As Double
Dim rate_puissance_fixe As Double
Dim rate_puissance_debit As Double
Dim ratio_elec_recirculation As Double
Dim ratio_elec_extraction As Double
Dim surpresseur_conso_Wh_Nm3mCE(nb_type_surpresseur, nb_surpresseur_gamme_Q, nb_surpresseur_gamme_P) As Double

'âge de boues équivalents limites pour moyenne charge
Dim hyp_Geq_mini As Double
Dim hyp_Geq_maxi As Double

'stockage de valeurs réelles
Dim stockage_MVS_bassins As Double
Dim stockage_Q As Double

Dim remarque(3) As String


Sub dimensionnement(ByRef Q, ByRef DCO, ByRef DBO, ByRef MES, ByRef NK, ByRef NH4, ByRef NO3, ByRef Pt, ByRef Sh, ByRef DCO_conso_bio)

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
    'DCO_conso_bio = O2_besoin


    'vitesse hydraulique max (recalculée) clarificateur
    clarif_vitesse_max_recalc = (Q_nominal * pd_CoefficientPointeHydrauliqueTempsPluie + (Q - Q_nominal)) / (pd_NOMBRE_HEURE_PAR_JOUR * clarif_surface)
    'charge au radier clarificateur
    clarif_charge_radier_nominal = nominal_MES_concentration_bassin * Q * (1 + recirculation_taux) / (pd_NOMBRE_HEURE_PAR_JOUR * clarif_surface)
    If clarif_charge_radier_nominal > charge_radier_max(choix_racleur) Then
        remarque(1) = Feuil12.Range("BA_moyenne_erreur_charge_radier").Cells(1, 1)
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
    NumeroErreurNonGeree = "12.5"

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
    NumeroErreurNonGeree = "12.6"

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
    Dim surpresseur_puissance As Double
    Dim electricite_racleur As Double
    Dim electricite_recirculation As Double
    Dim electricite_extraction As Double


    'AERATION

    'calcul facteur alpha
    If choix_aerateur = aerateur_fine_bulle Then
        'source: Oxygen transfer in AS (Henkel,Cornel and Wagner)
        O2_facteur_alfa = a0_alpha + a_alpha_MVS * (stockage_MVS_bassins / volume_bassins) + a_alpha_G * (reel_G * correctif_T ^ (pd_TemperatureEauConditionsReelles_°C - T_reference))
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
        If surpresseur_conso_spec_force = True Then
            surpresseur_conso_spec = surpresseur_conso_spec_f
            'gamme de pression
            If air_P_refoulement <= Px1 Then
                gamme_P = surpresseur_gamme_P_0_Px1
            ElseIf air_P_refoulement <= Px2 Then
                gamme_P = surpresseur_gamme_P_Px1_Px2
            Else
                gamme_P = surpresseur_gamme_P_Px2_infini
                If choix_surpresseur = surpresseur_roots Then
                    remarque(2) = Feuil12.Range("BA_moyenne_erreur_choix_surpresseur").Cells(1, 1)
                End If
            End If
        Else
            'gamme de pression
            If air_P_refoulement <= Px1 Then
                gamme_P = surpresseur_gamme_P_0_Px1
            ElseIf air_P_refoulement <= Px2 Then
                gamme_P = surpresseur_gamme_P_Px1_Px2
            Else
                gamme_P = surpresseur_gamme_P_Px2_infini
                If choix_surpresseur = surpresseur_roots Then
                    remarque(2) = Feuil12.Range("BA_moyenne_erreur_choix_surpresseur").Cells(1, 1)
                End If
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
    NumeroErreurNonGeree = "12.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub lecture_choix()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'lecture des choix internes
    choix_aerateur = Feuil12.Range("BA_moyenne_choix_aerateur").Cells(1, pi_FiliereConsideree)
    choix_regulation = Feuil12.Range("BA_moyenne_choix_regulation").Cells(1, pi_FiliereConsideree)
    choix_racleur = Feuil12.Range("BA_moyenne_choix_racleur").Cells(1, pi_FiliereConsideree)
    choix_surpresseur = Feuil12.Range("BA_moyenne_choix_surpresseur").Cells(1, pi_FiliereConsideree)
    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "12.1"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub lecture_valeurs_forcees()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'lecture du choix de forcer des valeurs
    nominal_G_force = Feuil12.Range("BA_moyenne_nominal_age_boue").Cells(1, pi_FiliereConsideree)
    nominal_MES_concentration_bassin_force = Feuil12.Range("BA_moyenne_nominal_MES_concentration_bassin").Cells(1, pi_FiliereConsideree)
    nominal_MV_MES_force = Feuil12.Range("BA_moyenne_nominal_MV_MES").Cells(1, pi_FiliereConsideree)
    volume_bassins_force = Feuil12.Range("BA_moyenne_volume_bassins").Cells(1, pi_FiliereConsideree)
    reel_G_force = Feuil12.Range("BA_moyenne_reel_age_boue").Cells(1, pi_FiliereConsideree)
    reel_MV_MES_force = Feuil12.Range("BA_moyenne_reel_MV_MES").Cells(1, pi_FiliereConsideree)
    reel_MES_concentration_bassin_force = Feuil12.Range("BA_moyenne_reel_MES_concentration_bassin").Cells(1, pi_FiliereConsideree)
    O2_dissous_force = Feuil12.Range("BA_moyenne_O2_dissous").Cells(1, pi_FiliereConsideree)
    sortie_DBO_force = Feuil12.Range("BA_moyenne_sortie_DBO").Cells(1, pi_FiliereConsideree)
    hauteur_bassin_force = Feuil12.Range("BA_moyenne_hauteur_bassin").Cells(1, pi_FiliereConsideree)
    O2_besoin_force = Feuil12.Range("BA_moyenne_O2_besoin").Cells(1, pi_FiliereConsideree)
    O2_facteur_alfa_force = Feuil12.Range("BA_moyenne_O2_facteur_alpha").Cells(1, pi_FiliereConsideree)
    If choix_aerateur = aerateur_fine_bulle Or choix_aerateur = aerateur_moyenne_bulle Then
        O2_rdt_transfert_force = Feuil12.Range("BA_moyenne_O2_rdt_transfert").Cells(1, pi_FiliereConsideree)
        air_Q_Nm3j_force = Feuil12.Range("BA_moyenne_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
        diffuseur_encrassement_force = Feuil12.Range("BA_moyenne_diffuseur_encrassement").Cells(1, pi_FiliereConsideree)
        air_P_refoulement_force = Feuil12.Range("BA_moyenne_air_P_refoulement").Cells(1, pi_FiliereConsideree)
        surpresseur_conso_spec_force = Feuil12.Range("BA_moyenne_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
    Else
        ASB_eau_claire_force = Feuil12.Range("BA_moyenne_ASB_eau_claire").Cells(1, pi_FiliereConsideree)
    End If
    recirculation_taux_force = Feuil12.Range("BA_moyenne_recirculation_taux").Cells(1, pi_FiliereConsideree)
    recirculation_P_refoulement_force = Feuil12.Range("BA_moyenne_recirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
    recirculation_pompe_rdt_force = Feuil12.Range("BA_moyenne_recirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
    mn_NombreClarificateurs_force = Feuil12.Range("BA_moyenne_NombreClarificateurs").Cells(1, pi_FiliereConsideree)
    boues_indice_Mohlman_force = Feuil12.Range("BA_moyenne_boues_indice_Mohlman").Cells(1, pi_FiliereConsideree)
    sortie_MES_force = Feuil12.Range("BA_moyenne_sortie_MES").Cells(1, pi_FiliereConsideree)
    clarif_hauteur_force = Feuil12.Range("BA_moyenne_clarif_hauteur").Cells(1, pi_FiliereConsideree)
    clarif_vitesse_max_force = Feuil12.Range("BA_moyenne_clarif_vitesse_max").Cells(1, pi_FiliereConsideree)
    clarif_surface_force = Feuil12.Range("BA_moyenne_clarif_surface").Cells(1, pi_FiliereConsideree)
    boues_concentration_force = Feuil12.Range("BA_moyenne_boues_concentration").Cells(1, pi_FiliereConsideree)
    boues_MES_force = Feuil12.Range("BA_moyenne_boues_MES").Cells(1, pi_FiliereConsideree)
    boues_Q_force = False
    extraction_P_refoulement_force = Feuil12.Range("BA_moyenne_extraction_P_refoulement").Cells(1, pi_FiliereConsideree)
    extraction_pompe_rdt_force = Feuil12.Range("BA_moyenne_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)

    'lecture des valeurs forcées
    If nominal_G_force = True Then
        nominal_G_f = Feuil12.Range("BA_moyenne_nominal_age_boue").Cells(1, pi_FiliereConsideree)
    End If
    If nominal_MES_concentration_bassin_force = True Then
        nominal_MES_concentration_bassin_f = Feuil12.Range("BA_moyenne_nominal_MES_concentration_bassin").Cells(1, pi_FiliereConsideree)
    End If
    If nominal_MV_MES_force = True Then
        nominal_MV_MES_f = Feuil12.Range("BA_moyenne_nominal_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If volume_bassins_force = True Then
        volume_bassins_f = Feuil12.Range("BA_moyenne_volume_bassins").Cells(1, pi_FiliereConsideree)
    End If
    If reel_G_force = True Then
        reel_G_f = Feuil12.Range("BA_moyenne_reel_age_boue").Cells(1, pi_FiliereConsideree)
    End If
    If reel_MV_MES_force = True Then
        reel_MV_MES_f = Feuil12.Range("BA_moyenne_reel_MV_MES").Cells(1, pi_FiliereConsideree)
    End If
    If reel_MES_concentration_bassin_force = True Then
        reel_MES_concentration_bassin_f = Feuil12.Range("BA_moyenne_reel_MES_concentration_bassin").Cells(1, pi_FiliereConsideree)
    End If
    If O2_dissous_force = True Then
        O2_dissous_f = Feuil12.Range("BA_moyenne_O2_dissous").Cells(1, pi_FiliereConsideree)
    End If
    If sortie_DBO_force = True Then
        sortie_DBO_f = Feuil12.Range("BA_moyenne_sortie_DBO").Cells(1, pi_FiliereConsideree)
    End If
    If hauteur_bassin_force = True Then
        hauteur_bassin_f = Feuil12.Range("BA_moyenne_hauteur_bassin").Cells(1, pi_FiliereConsideree)
    End If
    If O2_besoin_force = True Then
        O2_besoin_f = Feuil12.Range("BA_moyenne_O2_besoin").Cells(1, pi_FiliereConsideree)
    End If
    If O2_facteur_alfa_force = True Then
        O2_facteur_alfa_f = Feuil12.Range("BA_moyenne_O2_facteur_alpha").Cells(1, pi_FiliereConsideree)
    End If
    If choix_aerateur = aerateur_fine_bulle Or choix_aerateur = aerateur_moyenne_bulle Then
        If O2_rdt_transfert_force = True Then
            O2_rdt_transfert_f = Feuil12.Range("BA_moyenne_O2_rdt_transfert").Cells(1, pi_FiliereConsideree)
        End If
        If air_Q_Nm3j_force = True Then
            air_Q_Nm3j_f = Feuil12.Range("BA_moyenne_air_Q_Nm3j").Cells(1, pi_FiliereConsideree)
        End If
        If diffuseur_encrassement_force = True Then
            diffuseur_encrassement_f = Feuil12.Range("BA_moyenne_diffuseur_encrassement").Cells(1, pi_FiliereConsideree)
        End If
        If air_P_refoulement_force = True Then
            air_P_refoulement_f = Feuil12.Range("BA_moyenne_air_P_refoulement").Cells(1, pi_FiliereConsideree)
        End If
        If surpresseur_conso_spec_force = True Then
            surpresseur_conso_spec_f = Feuil12.Range("BA_moyenne_surpresseur_conso_spec").Cells(1, pi_FiliereConsideree)
        End If
    Else
        If ASB_eau_claire_force = True Then
            ASB_eau_claire_f = Feuil12.Range("BA_moyenne_ASB_eau_claire").Cells(1, pi_FiliereConsideree)
        End If
    End If
    If recirculation_taux_force = True Then
        recirculation_taux_f = Feuil12.Range("BA_moyenne_recirculation_taux").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_P_refoulement_force = True Then
        recirculation_P_refoulement_f = Feuil12.Range("BA_moyenne_recirculation_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If recirculation_pompe_rdt_force = True Then
        recirculation_pompe_rdt_f = Feuil12.Range("BA_moyenne_recirculation_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    If mn_NombreClarificateurs_force = True Then
        mi_NombreClarificateurs_f = Feuil12.Range("BA_moyenne_NombreClarificateurs").Cells(1, pi_FiliereConsideree)
    End If
    If boues_indice_Mohlman_force = True Then
        boues_indice_Mohlman_f = Feuil12.Range("BA_moyenne_boues_indice_Mohlman").Cells(1, pi_FiliereConsideree)
    End If
    If sortie_MES_force = True Then
        sortie_MES_f = Feuil12.Range("BA_moyenne_sortie_MES").Cells(1, pi_FiliereConsideree)
    End If
    If clarif_hauteur_force = True Then
        clarif_hauteur_f = Feuil12.Range("BA_moyenne_clarif_hauteur").Cells(1, pi_FiliereConsideree)
    End If
    If clarif_vitesse_max_force = True Then
        clarif_vitesse_max_f = Feuil12.Range("BA_moyenne_clarif_vitesse_max").Cells(1, pi_FiliereConsideree)
    End If
    If clarif_surface_force = True Then
        clarif_surface_f = Feuil12.Range("BA_moyenne_clarif_surface").Cells(1, pi_FiliereConsideree)
    End If
    If boues_concentration_force = True Then
        boues_concentration_f = Feuil12.Range("BA_moyenne_boues_concentration").Cells(1, pi_FiliereConsideree)
    End If
    If boues_MES_force = True Then
        boues_MES_f = Feuil12.Range("BA_moyenne_boues_MES").Cells(1, pi_FiliereConsideree)
    End If
    If boues_Q_force = True Then
        boues_MES_f = Feuil12.Range("BA_moyenne_boues_Q").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_P_refoulement_force = True Then
        extraction_P_refoulement_f = Feuil12.Range("BA_moyenne_extraction_P_refoulement").Cells(1, pi_FiliereConsideree)
    End If
    If extraction_pompe_rdt_force = True Then
        extraction_pompe_rdt_f = Feuil12.Range("BA_moyenne_extraction_pompe_rdt").Cells(1, pi_FiliereConsideree)
    End If
    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "12.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub hypotheses()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'Voici l'ensemble des hypothèses sur le process

    Dim I As Integer
    Dim j As Integer



    'Température de référence pour la cinétique biologique (°C)
    T_reference = Feuil48.Range("hyp_II_BA_bio_equiv_Tref").Cells(1, 1)    ' 12

    'Age de boues à T_reference (j)
    G_reference = 4

    'Facteur correctif température pour cinétique
    correctif_T = Feuil48.Range("hyp_II_BA_bio_equiv_correctionT").Cells(1, 1)    '1.07

    'coefficients relation charge massique (kgDBO/kgMVS/j)  1/Cm=a1*G_eq+a0
    a1_Cm = Feuil48.Range("hyp_II_BA_charge_massique_a1").Cells(1, 1)    '0.4959
    a0_Cm = Feuil48.Range("hyp_II_BA_charge_massique_a0").Cells(1, 1)    '1.3908

    'coefficients relation rendement DBO fonction de Cm_equivalente rdt_DBO=a2*Cm^2+a1*Cm+a0
    FtoM_lim_kgDBO_kgMESj = Feuil48.Range("hyp_II_BA_rdtDBO_FtoM_max").Cells(1, 1)      '1.9
    If choix_primaire = True Then
        For I = 1 To 4
            a_rdt_DBO(I - 1) = Feuil48.Range("hyp_II_BA_rdtDBO_coef_ED").Cells(I, 1)
        Next I
    Else
        For I = 1 To 4
            a_rdt_DBO(I - 1) = Feuil48.Range("hyp_II_BA_rdtDBO_coef_EB").Cells(I, 1)
        Next I
    End If

    'hypothèse de rendement sur la DCO
    rdt_DCO = 0.9

    'Charge radier maximale kg/(m2.h)
    For I = racleur_racle To racleur_Kruger
        charge_radier_max(I) = Feuil48.Range("hyp_II_BA_charge_radier_max").Cells(I, 1)
    Next I

    'hypothèses pour le calcul des besoins en O2
    a0_O2_DBO_respiration = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_a0").Cells(1, 1)   ' 0.56
    a1_O2_DBO_respiration = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_a1").Cells(1, 1)  ' 0.15
    a2_O2_DBO_respiration = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_a2").Cells(1, 1)   '0.17
    correctif_T_ATV = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_correctionT").Cells(1, 1)   '1.072
    T_ref_ATV = Feuil48.Range("hyp_II_BA_besoinO2_DBO_resp_Tref").Cells(1, 1)   '15

    'facteurs alpha0 (celui de fines bulles est calculé)
    O2_alpha0(aerateur_moyenne_bulle) = 0.8
    O2_alpha0(aerateur_brosse) = 0.8
    O2_alpha0(aerateur_turbine_lente) = 0.8
    O2_alpha0(aerateur_turbine_rapide) = 0.8

    'correction alpha en fonction des MES  alfa=alfa*(a0+a1*MES_reel)/(a0+a1*MES_lim)
    a0_correctif_alfa_MES = Feuil48.Range("hyp_II_BA_alpha_correction_MES_a0").Cells(1, 1)    ' 1.0007
    a1_correctif_alfa_MES = Feuil48.Range("hyp_II_BA_alpha_correction_MES_a1").Cells(1, 1)    ' -0.035
    MES_reference_alfa = Feuil48.Range("hyp_II_BA_alpha_correction_MES_Cref").Cells(1, 1)    ' 5

    'relation calcul alpha fines bulles    ao+a_MVS*MVS+a_G*G_eq
    a0_alpha = Feuil48.Range("hyp_II_BA_alpha_finebulle_a0").Cells(1, 1)    ' 0.51
    a_alpha_MVS = Feuil48.Range("hyp_II_BA_alpha_finebulle_aMVS").Cells(1, 1)    '-0.062
    a_alpha_G = Feuil48.Range("hyp_II_BA_alpha_finebulle_aGeq").Cells(1, 1)    ' 0.019
    O2_facteur_alfa_maxi = 1
    O2_facteur_alfa_mini = 0.05

    'facteur beta
    O2_facteur_beta = Feuil48.Range("hyp_calcul_facteurK_facteur_Beta").Cells(1, 1)    ' 0.95

    'calcul de la pression atmosphérique  Patm=P0*((a0-a1*altitude)^a2)/100
    P0 = Feuil48.Range("hyp_calcul_Patm_coef").Cells(1, 1)   ' 101325
    a0_Patm = Feuil48.Range("hyp_calcul_Patm_coef").Cells(2, 1)    '1
    a1_Patm = Feuil48.Range("hyp_calcul_Patm_coef").Cells(3, 1)    '-0.0000225577
    a2_Patm = Feuil48.Range("hyp_calcul_Patm_coef").Cells(4, 1)    '5.25588

    'calcul de la surpression en diffusion de bulles  delta_P=a0/a1*P0/100*(hauteur_bassin-hauteur_diffuseur)
    a0_delta_P = Feuil48.Range("hyp_insufflation_deltaP_a0").Cells(1, 1)  '0.35
    a1_delta_P = Feuil48.Range("hyp_insufflation_deltaP_a1").Cells(1, 1)   '10.33
    hauteur_diffuseur = Feuil48.Range("hyp_insufflation_bulles_hauteur_diffuseur").Cells(1, 1)   '0.25 'm

    'calcul O2 saturation  O2_saturation(T,P0)= somme(ai*Teau^i)
    For I = 1 To 5
        a_O2_saturation(I - 1) = Feuil48.Range("hyp_calcul_O2sat_T_P0_coef_a").Cells(I, 1)
    Next I

    'correction température pour facteur K
    correctif_T_K = Feuil48.Range("hyp_transfertO2_facteurK_correction_T").Cells(1, 1)    ' 1.024
    T_ref_K = Feuil48.Range("hyp_transfertO2_facteurK_Tref").Cells(1, 1)  '20

    'O2 présent dans l'air (kgO2/Nm3air)
    ratio_kgO2_Nm3air = Feuil48.Range("hyp_ratio_kgO2_Nm3air").Cells(1, 1)  ' 0.3

    'rendement de dissolution en eau claire (%/m)
    O2_rdt_dissolution(aerateur_fine_bulle) = Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(3, 1)
    O2_rdt_dissolution(aerateur_moyenne_bulle) = Feuil48.Range("hyp_II_BA_rdt_dissolution_O2_eau_claire").Cells(2, 1)

    'Apport spécifique brut (eau claire)  (kWh/kO2_transféré)   source:fndae 31
    ASB(aerateur_brosse) = Feuil48.Range("hyp_II_BA_aerateur_surface_ASB").Cells(1, 1)   '1.7
    ASB(aerateur_turbine_lente) = Feuil48.Range("hyp_II_BA_aerateur_surface_ASB").Cells(2, 1)   '1.65
    ASB(aerateur_turbine_rapide) = Feuil48.Range("hyp_II_BA_aerateur_surface_ASB").Cells(3, 1)   '1.2

    'consommations spécifiques par type de surpresseurs      en Wh/(Nm3.mCE)
    For I = 1 To nb_surpresseur_gamme_Q
        For j = 1 To nb_surpresseur_gamme_P
            surpresseur_conso_Wh_Nm3mCE(surpresseur_roots, I, j) = 4.5
            surpresseur_conso_Wh_Nm3mCE(surpresseur_vis, I, j) = 3.8
            surpresseur_conso_Wh_Nm3mCE(surpresseur_turbo, I, j) = 3.5
        Next j
    Next I

    'impact du type de régulation
'    impact_regulation(regulation_horloge) = 0.4
'    impact_regulation(regulation_sans_variation) = 0.25
'    impact_regulation(regulation_avec_variation) = 0.1
'    impact_regulation(regulation_avance) = 0
    If pd_TemperatureEauConditionsReelles_°C < 14 Then
        md_ConcentrationOxygeneDissous_mgL(regulation_horloge) = 4
        md_ConcentrationOxygeneDissous_mgL(regulation_sans_variation) = 2
        md_ConcentrationOxygeneDissous_mgL(regulation_avec_variation) = 1.5
        md_ConcentrationOxygeneDissous_mgL(regulation_avance) = 1.5
    Else
        md_ConcentrationOxygeneDissous_mgL(regulation_horloge) = 4
        md_ConcentrationOxygeneDissous_mgL(regulation_sans_variation) = 1.5
        md_ConcentrationOxygeneDissous_mgL(regulation_avec_variation) = 1
        md_ConcentrationOxygeneDissous_mgL(regulation_avance) = 1
    End If

    'clarificateur
    rate_puissance_fixe = 0.1    'Wh/m3 par rapport au nominal
    rate_puissance_debit = 1    'Wh/m3 par rapport au débit réel

    'P consommé dans synthèse
    ratio_P_synthese = Feuil48.Range("hyp_II_assimilation_P").Cells(1, 1)  ' 0.01

    'âges de boues équivalents limites
    hyp_Geq_mini = 3
    hyp_Geq_maxi = 7




    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "12.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub attribution_valeur_par_defaut()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On retrouve ici l'ensemble des valeurs pour lesquelles on attribue une valeur guide par défaut
    Dim volume_corrige As Double
    Dim G_equivalent_nominal_G As Double
    Dim rdt_pompe As Double
    Dim rdt_moteur_pompe As Double


    'Age de boues design (j)
    If nominal_G_force = True Then
        nominal_G = nominal_G_f
    Else
        nominal_G = G_reference * correctif_T ^ (T_reference - pd_TemperatureEauDimensionnement_°C)
    End If

    'Concentration MES design (dans les bassins) g/L
    If nominal_MES_concentration_bassin_force = True Then
        nominal_MES_concentration_bassin = nominal_MES_concentration_bassin_f
    Else
        If choix_primaire = True Then
            If pd_TemperatureEauDimensionnement_°C < 12 Then
                nominal_MES_concentration_bassin = 4
            ElseIf pd_TemperatureEauDimensionnement_°C < 18 Then
                nominal_MES_concentration_bassin = 3.5
            Else
                nominal_MES_concentration_bassin = 3
            End If
        Else
            nominal_MES_concentration_bassin = 5
        End If
    End If

    'MV/MES des boues design
    If nominal_MV_MES_force = True Then
        nominal_MV_MES = nominal_MV_MES_f
    Else
        If choix_primaire = True Then
            If pd_TemperatureEauDimensionnement_°C < 16 Then
                nominal_MV_MES = 0.8
            Else
                nominal_MV_MES = 0.85
            End If
        Else
            nominal_MV_MES = 0.74
        End If
    End If

    'Age de boues réel (j)
    If reel_G_force = True Then
        reel_G = reel_G_f
    Else
        G_equivalent_nominal_G = nominal_G / correctif_T ^ (T_reference - pd_TemperatureEauDimensionnement_°C)
        reel_G = G_equivalent_nominal_G * correctif_T ^ (T_reference - pd_TemperatureEauConditionsReelles_°C)
    End If

    'MV/MES des boues réel
    If reel_MV_MES_force = True Then
        reel_MV_MES = reel_MV_MES_f
    Else
        If choix_primaire = True Then
            If pd_TemperatureEauConditionsReelles_°C < 16 Then
                reel_MV_MES = 0.8
            Else
                reel_MV_MES = 0.85
            End If
        Else
            reel_MV_MES = 0.74
        End If
    End If

    'Concentration moyenne en O2 dissous (mg/L)
    If O2_dissous_force = True Then
        O2_dissous = O2_dissous_f
    Else
        O2_dissous = md_ConcentrationOxygeneDissous_mgL(choix_regulation)
    End If

    'Hauteur moyenne des bassins
    If hauteur_bassin_force = True Then
        hauteur_bassin = hauteur_bassin_f
    Else
        If choix_aerateur = aerateur_fine_bulle Or choix_aerateur = aerateur_moyenne_bulle Then
            hauteur_bassin = 6
        Else
            hauteur_bassin = 3
        End If
    End If

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
            boues_indice_Mohlman = 185
        Else
            boues_indice_Mohlman = 160
        End If
    End If

    'Sortie MES eau traitée (mg/L)
    If sortie_MES_force = True Then
        sortie_MES = sortie_MES_f
    Else
        sortie_MES = 25
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
    NumeroErreurNonGeree = "12.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub ecriture_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil12.Range("BA_moyenne_nominal_age_boue").Cells(1, Colonne) = nominal_G
    Feuil12.Range("BA_moyenne_nominal_MES_concentration_bassin").Cells(1, Colonne) = nominal_MES_concentration_bassin
    Feuil12.Range("BA_moyenne_nominal_MV_MES").Cells(1, Colonne) = nominal_MV_MES
    Feuil12.Range("BA_moyenne_volume_bassins").Cells(1, Colonne) = volume_bassins
    Feuil12.Range("BA_moyenne_reel_age_boue").Cells(1, Colonne) = reel_G
    Feuil12.Range("BA_moyenne_reel_MV_MES").Cells(1, Colonne) = reel_MV_MES
    Feuil12.Range("BA_moyenne_reel_MES_concentration_bassin").Cells(1, Colonne) = reel_MES_concentration_bassin
    Feuil12.Range("BA_moyenne_O2_dissous").Cells(1, Colonne) = O2_dissous
    Feuil12.Range("BA_moyenne_sortie_DBO").Cells(1, Colonne) = sortie_DBO
    Feuil12.Range("BA_moyenne_hauteur_bassin").Cells(1, Colonne) = hauteur_bassin
    Feuil12.Range("BA_moyenne_O2_besoin").Cells(1, Colonne) = O2_besoin
    Feuil12.Range("BA_moyenne_O2_facteur_alpha").Cells(1, Colonne) = O2_facteur_alfa
    If choix_aerateur = aerateur_fine_bulle Or choix_aerateur = aerateur_moyenne_bulle Then
        Feuil12.Range("BA_moyenne_O2_rdt_transfert").Cells(1, Colonne) = O2_rdt_transfert
        Feuil12.Range("BA_moyenne_air_Q_Nm3j").Cells(1, Colonne) = air_Q_Nm3j
        Feuil12.Range("BA_moyenne_diffuseur_encrassement").Cells(1, Colonne) = diffuseur_encrassement
        Feuil12.Range("BA_moyenne_air_P_refoulement").Cells(1, Colonne) = air_P_refoulement
        Feuil12.Range("BA_moyenne_surpresseur_conso_spec").Cells(1, Colonne) = surpresseur_conso_spec
    Else
        Feuil12.Range("BA_moyenne_ASB_eau_claire").Cells(1, Colonne) = ASB_eau_claire
    End If
    Feuil12.Range("BA_moyenne_recirculation_taux").Cells(1, Colonne) = recirculation_taux
    Feuil12.Range("BA_moyenne_recirculation_P_refoulement").Cells(1, Colonne) = recirculation_P_refoulement
    Feuil12.Range("BA_moyenne_recirculation_pompe_rdt").Cells(1, Colonne) = recirculation_pompe_rdt
    Feuil12.Range("BA_moyenne_NombreClarificateurs").Cells(1, Colonne) = mi_NombreClarificateurs
    Feuil12.Range("BA_moyenne_boues_indice_Mohlman").Cells(1, Colonne) = boues_indice_Mohlman
    Feuil12.Range("BA_moyenne_sortie_MES").Cells(1, Colonne) = sortie_MES
    Feuil12.Range("BA_moyenne_clarif_hauteur").Cells(1, Colonne) = clarif_hauteur
    Feuil12.Range("BA_moyenne_clarif_vitesse_max").Cells(1, Colonne) = clarif_vitesse_max
    Feuil12.Range("BA_moyenne_clarif_surface").Cells(1, Colonne) = clarif_surface
    Feuil12.Range("BA_moyenne_clarif_vitesse_max_recalculee").Cells(1, Colonne) = clarif_vitesse_max_recalc
    Feuil12.Range("BA_moyenne_clarif_charge_radier_nominal").Cells(1, Colonne) = clarif_charge_radier_nominal
    Feuil12.Range("BA_moyenne_boues_concentration").Cells(1, Colonne) = boues_concentration
    Feuil12.Range("BA_moyenne_boues_MES").Cells(1, Colonne) = boues_MES
    Feuil12.Range("BA_moyenne_boues_Q").Cells(1, Colonne) = boues_Q
    Feuil12.Range("BA_moyenne_extraction_P_refoulement").Cells(1, Colonne) = extraction_P_refoulement
    Feuil12.Range("BA_moyenne_extraction_pompe_rdt").Cells(1, Colonne) = extraction_pompe_rdt
    Feuil12.Range("BA_moyenne_electricite").Cells(1, Colonne) = electricite
    Feuil12.Range("BA_moyenne_electricite_aeration").Cells(1, Colonne) = electricite_aeration

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "12.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Sub effacer_resultats()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String


    'On efface les valeurs sur Excel
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil12.Range("BA_moyenne_nominal_age_boue").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_nominal_MES_concentration_bassin").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_nominal_MV_MES").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_volume_bassins").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_reel_age_boue").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_reel_MV_MES").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_reel_MES_concentration_bassin").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_O2_dissous").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_sortie_DBO").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_hauteur_bassin").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_O2_besoin").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_O2_facteur_alpha").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_O2_rdt_transfert").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_air_Q_Nm3j").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_diffuseur_encrassement").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_air_P_refoulement").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_surpresseur_conso_spec").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_ASB_eau_claire").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_recirculation_taux").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_recirculation_P_refoulement").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_recirculation_pompe_rdt").Cells(1, Colonne) = ""      'recirculation_pompe_rdt
    Feuil12.Range("BA_moyenne_NombreClarificateurs").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_boues_indice_Mohlman").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_sortie_MES").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_clarif_hauteur").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_clarif_vitesse_max").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_clarif_surface").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_clarif_vitesse_max_recalculee").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_clarif_charge_radier_nominal").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_boues_concentration").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_boues_MES").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_boues_Q").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_extraction_P_refoulement").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_extraction_pompe_rdt").Cells(1, Colonne) = ""     'extraction_pompe_rdt
    Feuil12.Range("BA_moyenne_electricite").Cells(1, Colonne) = ""
    Feuil12.Range("BA_moyenne_electricite_aeration").Cells(1, Colonne) = ""         'electricite_aeration

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "12.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub


Sub affich_remarque()
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim nb_remarque As Integer
    Dim I As Integer

    nb_remarque = 2  'penser à modifier la capacité du vecteur "message_erreur"

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil12.Range("BA_moyenne_erreur_nom")

    For I = 1 To nb_remarque
        If remarque(I) <> "" Then
            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(I)
        End If
    Next I

    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf
    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "12.8"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True


End Sub

Private Sub Class_Initialize()

    If pn_EffacementResultatsEnCours = False Then
        Call lecture_choix
        If pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
            Call lecture_valeurs_forcees
        End If
        Call hypotheses
        Call attribution_valeur_par_defaut
    End If



End Sub

Private Sub Class_Terminate()

    If pn_EffacementResultatsEnCours = False Then
        'Call calcul_consommation_electrique
        Call ecriture_resultats
        Call affich_remarque
    End If

End Sub







