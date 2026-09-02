Attribute VB_Name = "Photovoltaique"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'## INDICES ET GESTION DES CHOIX
'variables de choix
Dim identifiant_station_meteo As Integer
Dim ligne_station_meteo As Integer
Dim choix_technologie As Integer

'constantes pour le choix de la technologie
Dim VECT_RENDEMENT_MODULE(1 To 5) As Double
Dim VECT_NOCT(1 To 5) As Double
Dim VECT_COEF_TEMP(1 To 5) As Double
Private Const TECHNOLOGIE_SI_MC = 1
Private Const TECHNOLOGIE_SI_PC = 2
Private Const TECHNOLOGIE_SI_A = 3
Private Const TECHNOLOGIE_CDTE = 4
Private Const TECHNOLOGIE_CIS = 5

'## CONSTANTES
Const pd_NOMBRE_PI = 3.14159265358979
Const Gs = 0.082    'MJ.m^-2.min-1 : constante solaire

'## PARAMETRES
'paramètres issus de la base de données
Dim lat As Double    'latitude du lieu
Dim G_irr_moy(1 To 12) As Double    'moyenne mensuelle l'irradiation solaire journalière en kWh.m^-2.j^-1
Dim T_air_moy(1 To 12) As Double  'moyenne mensuelle de la température journalière de l'air en °C

'paramètres d'entrée feuille "Synthèse"
Dim surfaceInstallee As Double    'surface totale [m^2]
Dim orientation As Double    ' orientation de modules PV[°]
Dim inclinaison As Double    ' inclinaison des modules par rapport à l'horizontal [°]

'rendements de l'installation
Dim rendementModule As Double
Dim NOCT As Double
Dim coefTemp As Double
Dim pertesSysteme As Double
Dim T_ref As Double

'boolean état forcé
Dim rendementModule_force As Boolean
Dim NOCT_force As Boolean
Dim coefTemp_force As Boolean
Dim pertesSysteme_force As Boolean
Dim orientation_force As Boolean
Dim inclinaison_force As Boolean

'valeurs forcées
Dim rendementModule_f As Double
Dim NOCT_f As Double
Dim coefTemp_f As Double
Dim pertesSysteme_f As Double
Dim orientation_f As Double
Dim inclinaison_f As Double

'## ELEMENTS CALCULES
'tableaux des moyennes mensuelles des irradiations journalières (inclinées et orientées)
Dim S_irr_moy_t(1 To 12) As Double
Dim D_irr_moy_t(1 To 12) As Double
Dim R_irr_moy_t(1 To 12) As Double
Dim G_irr_moy_t(1 To 12) As Double
Dim G_irr_kwh_m2_an As Double
Dim G_irr_moy_kwh_m2_j As Double

'production electricité
Dim E_p(1 To 12) As Double
Dim E_p_kwh_an As Double
Public electricite_produite_kwh_j As Double

'FONCTIONS USUELLES
Function Arccos(ByVal x As Double) As Double
    'retourne un angle en rad
    If version_Developpement = False Then On Error GoTo erreur_non_geree

    Arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)

    Exit Function
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("69.1")
    pn_EffacementResultatsEnCours = True

End Function
Function Arcsin(ByVal x As Double) As Double
    'retourne un angle en rad
    If version_Developpement = False Then On Error GoTo erreur_non_geree

    Arcsin = Atn(x / Sqr(-x * x + 1))

    Exit Function
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("69.2")
    pn_EffacementResultatsEnCours = True

End Function

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

Sub lecture_choix()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'choix station météo
    identifiant_station_meteo = Feuil5.Range("choixIdStationBDClimat").Cells(1, pi_FiliereConsideree)
    ligne_station_meteo = identifiant_station_meteo + 1    ' +1 pour correspondre à la ligne

    'surface
    surfaceInstallee = Feuil69.Range("PV_surfaceInstallee").Cells(1, pi_FiliereConsideree)

    'choix_technologie
    choix_technologie = Feuil69.Range("PV_choix_technologie").Cells(1, pi_FiliereConsideree)


    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("69.3")
    pn_EffacementResultatsEnCours = True

End Sub

Sub lecture_valeurs_forcees()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'lecture du choix de forcer des valeurs
    rendementModule_force = Feuil69.Range("PV_rendementModule").Cells(1, pi_FiliereConsideree)
    NOCT_force = Feuil69.Range("PV_NOCT").Cells(1, pi_FiliereConsideree)
    coefTemp_force = Feuil69.Range("PV_coefTemp").Cells(1, pi_FiliereConsideree)
    pertesSysteme_force = Feuil69.Range("PV_pertesSysteme").Cells(1, pi_FiliereConsideree)
    orientation_force = Feuil69.Range("PV_orientation").Cells(1, pi_FiliereConsideree)
    inclinaison_force = Feuil69.Range("PV_inclinaison").Cells(1, pi_FiliereConsideree)

    'lecture des valeurs forcées
    If rendementModule_force = True Then
        rendementModule_f = Feuil69.Range("PV_rendementModule").Cells(1, pi_FiliereConsideree)
    End If
    If NOCT_force = True Then
        NOCT_f = Feuil69.Range("PV_NOCT").Cells(1, pi_FiliereConsideree)
    End If
    If coefTemp_force = True Then
        coefTemp_f = Feuil69.Range("PV_coefTemp").Cells(1, pi_FiliereConsideree)
    End If
    If pertesSysteme_force = True Then
        pertesSysteme_f = Feuil69.Range("PV_pertesSysteme").Cells(1, pi_FiliereConsideree)
    End If
    If orientation_force = True Then
        orientation_f = Feuil69.Range("PV_orientation").Cells(1, pi_FiliereConsideree)
    End If
    If inclinaison_force = True Then
        inclinaison_f = Feuil69.Range("PV_inclinaison").Cells(1, pi_FiliereConsideree)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("69.4")
    pn_EffacementResultatsEnCours = True

End Sub

Sub hypotheses()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    Dim mois As Integer

    '## RENDEMENTS DE L'INSTALLATION
    'rendement du module en fonction du type de technologie [%]
    VECT_RENDEMENT_MODULE(TECHNOLOGIE_SI_MC) = 13
    VECT_RENDEMENT_MODULE(TECHNOLOGIE_SI_PC) = 11
    VECT_RENDEMENT_MODULE(TECHNOLOGIE_SI_A) = 5
    VECT_RENDEMENT_MODULE(TECHNOLOGIE_CDTE) = 7
    VECT_RENDEMENT_MODULE(TECHNOLOGIE_CIS) = 7.5

    'NOCT du module en fonction du type de technologie [°C] (température nominale de fonctionnement des cellules)
    VECT_NOCT(TECHNOLOGIE_SI_MC) = 45
    VECT_NOCT(TECHNOLOGIE_SI_PC) = 45
    VECT_NOCT(TECHNOLOGIE_SI_A) = 50
    VECT_NOCT(TECHNOLOGIE_CDTE) = 46
    VECT_NOCT(TECHNOLOGIE_CIS) = 47

    'coefficient de température du module en fonction du type de technologie [%/°C]
    VECT_COEF_TEMP(TECHNOLOGIE_SI_MC) = 0.4
    VECT_COEF_TEMP(TECHNOLOGIE_SI_PC) = 0.4
    VECT_COEF_TEMP(TECHNOLOGIE_SI_A) = 0.11
    VECT_COEF_TEMP(TECHNOLOGIE_CDTE) = 0.24
    VECT_COEF_TEMP(TECHNOLOGIE_CIS) = 0.46

    '## BASE DE DONNEES
    lat = Feuil68.Range("Lat").Cells(ligne_station_meteo, 1).Value    ' * pd_NOMBRE_PI / 180
    For mois = 1 To 12
        G_irr_moy(mois) = Feuil68.Range("Solar_Jan").Cells(ligne_station_meteo, mois).Value
        T_air_moy(mois) = Feuil68.Range("AirT_Jan").Cells(ligne_station_meteo, mois).Value
    Next

    'divers
    T_ref = 25    '[°C]

    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("69.5")
    pn_EffacementResultatsEnCours = True

End Sub

Private Sub attribution_valeur_par_defaut()

    If version_Developpement = False Then On Error GoTo erreur_non_geree


    'orientation
    If orientation_force = True Then
        If orientation_f > 180 Then
            orientation = 180
        ElseIf orientation_f < -180 Then
            orientation = -180
        Else
            orientation = orientation_f
        End If
    Else
        If lat >= 0 Then    'emisphère Nord -> orientation Sud
            orientation = 0    'Sud
        ElseIf lat < 0 Then    'emisphère Sud-> orientation Nord
            orientation = 180    'Nord
        End If
    End If

    'inclinaison
    If inclinaison_force = True Then
        If inclinaison_f > 90 Then
            inclinaison = 90
        ElseIf inclinaison_f < 0 Then
            inclinaison = 0
        Else
            inclinaison = inclinaison_f
        End If
    Else
        If Abs(lat) < 7.5 Then
            inclinaison = 0
        ElseIf Abs(lat) < 22.5 Then
            inclinaison = 15
        ElseIf Abs(lat) < 37.5 Then
            inclinaison = 30
        ElseIf Abs(lat) < 52.5 Then
            inclinaison = 45
        Else    ' >= 52.5°
            inclinaison = 60
        End If
    End If

    'rendementModule
    If rendementModule_force = True Then
        rendementModule = rendementModule_f
    Else
        rendementModule = VECT_RENDEMENT_MODULE(choix_technologie)
    End If

    'NOCT
    If NOCT_force Then
        NOCT = NOCT_f
    Else
        NOCT = VECT_NOCT(choix_technologie)
    End If

    'coefTemp
    If coefTemp_force = True Then
        coefTemp = coefTemp_f
    Else
        coefTemp = VECT_COEF_TEMP(choix_technologie)
    End If

    'pertesSysteme
    If pertesSysteme_force = True Then
        pertesSysteme = pertesSysteme_f
    Else
        pertesSysteme = 12    '[%] onduleur, transformateur etc.
    End If

    Exit Sub

    '    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("69.6")
    pn_EffacementResultatsEnCours = True

End Sub

Sub calcul()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    '### DECLARATION DES VARIABLES

    Dim mois As Integer    'indice pour les itérations correspondant au numpéro du mois (1 à 12)
    Dim nb_jrs_mois As Variant
    'PARAMETRES JOURNALIERS
    Dim num_jrs_mois As Variant
    Dim num_jrs_mois_cumul As Variant
    Dim j As Integer    ' indice pour les itérations correspondant au jour de l'année (1 à 365)
    Dim d_invTS    ' inverse de la distance Terre-Soleil
    Dim decl As Double    ' déclinaison solaire
    Dim angle_c As Double    'angle de lever/coucher du soleil
    Dim EA_irr As Double    'rayonnement extra atmosphérique
    Dim d_invTS_moy(1 To 12) As Double
    Dim decl_moy(1 To 12) As Double
    Dim angle_c_moy(1 To 12) As Double
    Dim EA_irr_moy(1 To 12) As Double    'moyenne mensuelle de l'irradiation extra-atmosphérique journalière
    Dim ind_clarte_moy(1 To 12) As Double    'moyenne mensuelle de l'indince de clarté
    Dim D_irr_moy(1 To 12) As Double    'moyenne mensuelle de l'irradiation diffuse journalière
    Dim S_irr_moy(1 To 12) As Double    'moyenne mensuelle de l'irradiation directe journalière
    Dim duree_j_moy(1 To 12) As Double    'moyenne mensuelle de la durée d'une journée [heure]
    'PARAMETRES HORAIRES
    'paramètres
    Dim num_jour_moyen As Variant
    Dim heure As Double    ' indice pour les itérations correspondant à l'heure pendant la période d'ensoleillement d'une journée moyenne
    Dim iterations As Single    ' indice pour les itérations correspondant au nombre de découpage horaire pendant la durée d'ensoleillement d'une journée moyenne
    Dim h_lever As Double    'heure de lever du soleil (temps solaire vrai)
    Dim h_coucher As Double    'heure de coucher du soleil (temps solaire vrai)
    Dim angle_h As Double   'angle horaire  [rad]
    Dim hauteur_S As Double    'angle correspondant à la hauteur solaire [rad]
    Dim azimut As Double    'angle azimut, projection horizontale de la position du soleil par rapport au sud [rad]
    Dim angle_i As Double    'angle d'incidence des rayons sur un plan incliné et orienté [rad]
    Dim albedo As Double    'albedo
    Dim Rb As Double                'coefficient pour la projection
    Dim pas As Double    ' pas des itérations pour la répartition horaire
    'avant transposition
    Dim G_ecl As Double             'éclairement global [kW/m2]
    Dim D_ecl As Double              'éclairement diffu [kW/m2]
    Dim S_ecl As Double             'éclairement direct [kW/m2]
    Dim S_ecl_moy(1 To 12) As Double    'moyenne mensuelle de l'éclairement direct journalier [kW/m2/j]
    Dim D_ecl_moy(1 To 12) As Double    'moyenne mensuelle de l'éclairement diffus journalier [kW/m2/j]
    Dim G_ecl_moy(1 To 12) As Double    'moyenne mensuelle de l'éclairement global journalier [kW/m2/j]
    Dim R_ecl_moy(1 To 12) As Double    'moyenne mensuelle de l'éclairement réfléchi journalier [kW/m2/j]
    'après transposition
    Dim S_ecl_t As Double    '[kWh/m2]
    Dim R_ecl_t As Double    '[kWh/m2]
    Dim D_ecl_t As Double    '[kWh/m2]
    Dim S_ecl_moy_t(1 To 12) As Double    'moyenne mensuelle de l'éclairement direct transposé journalier [kWh/m2/j]
    Dim D_ecl_moy_t(1 To 12) As Double    'moyenne mensuelle de l'éclairement diffus transposé journalier [kWh/m2/j]
    Dim G_ecl_moy_t(1 To 12) As Double    'moyenne mensuelle de l'éclairement global transposé journalier [kWh/m2/j]
    Dim R_ecl_moy_t(1 To 12) As Double    'moyenne mensuelle de l'éclairement réfléchi transposé journalier [kWh/m2/j]
    'APPLICATION DES RENDEMENTS DE L'INSTALLATION
    Dim rendement_moyen
    Dim T_module As Double



    '### INITIALISATION DES VARIABLES

    'variables internes
    nb_jrs_mois = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)    ' nombre de jours par mois pour chaque mois (total : 365)
    num_jrs_mois_cumul = Array(31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334, 365)    ' numéro du dernier jour de chaque mois (1 à 365)
    num_jour_moyen = Array(17, 47, 75, 105, 135, 162, 198, 228, 258, 288, 318, 344)    'Said Al-Hallaj - Kristofer Kiszynski - Hybrid Hydrogen Systems - Stationary and Transportation Applications (edition Springer-Verlag London Limited 2011)
    pas = 0.1  ' pas des itérations pour la répartition horaire

    'variables publiques dans le module
    lat = lat * pd_NOMBRE_PI / 180   '[rad]
    orientation = orientation * pd_NOMBRE_PI / 180    '[rad]
    inclinaison = inclinaison * pd_NOMBRE_PI / 180    '[rad]



    '### PROCEDURE

    '## CALCUL DES PARAMETRES JOURNALIERS
    mois = 1
    j = 1
    Do
        If mois <= 12 Then
            If j <= num_jrs_mois_cumul(mois - 1) Then
                'inverse de la distance terre-soleil au jour j
                d_invTS = 1 + 0.033 * Cos((2 * pd_NOMBRE_PI * j) / pd_NOMBRE_JOUR_PAR_AN)
                d_invTS_moy(mois) = d_invTS_moy(mois) + d_invTS

                'angle de déclinaison solaire au jour j
                decl = 0.409 * Sin(((2 * pd_NOMBRE_PI * j) / pd_NOMBRE_JOUR_PAR_AN) - 1.39)    '[rad]
                decl_moy(mois) = decl_moy(mois) + decl

                'angle de lever/coucher du soleil au jour j
                Dim x As Double
                x = -Tan(lat) * Tan(decl)
                angle_c = Arccos(x)
                angle_c_moy(mois) = angle_c_moy(mois) + angle_c

                'irradiation extra-atmosphérique au jour j
                EA_irr = 24 * 60 * Gs * d_invTS * (1 / pd_NOMBRE_PI) _
                         * ((angle_c * Sin(lat) * Sin(decl)) + (Cos(lat) _
                                                                * Cos(decl) * Sin(angle_c))) * (1 / 3.6)    '[kWh.m^-2.j^-1](*(1/3.6) pour passer de MJ.m^-2.j^-1 en kWh.m^-2.j^-1
                EA_irr_moy(mois) = EA_irr_moy(mois) + EA_irr

                j = j + 1
            Else

                'calcul des moyennes mensuelles
                d_invTS_moy(mois) = d_invTS_moy(mois) / nb_jrs_mois(mois - 1)
                decl_moy(mois) = decl_moy(mois) / nb_jrs_mois(mois - 1)
                angle_c_moy(mois) = angle_c_moy(mois) / nb_jrs_mois(mois - 1)
                EA_irr_moy(mois) = EA_irr_moy(mois) / nb_jrs_mois(mois - 1)

                'moyenne mensuelle de l'indince de clarté journalier
                ind_clarte_moy(mois) = G_irr_moy(mois) / EA_irr_moy(mois)    '[sans dimensions]

                'moyenne mensuelle de l'irradiation diffuse journalière
                D_irr_moy(mois) = G_irr_moy(mois) * (0.775 + 0.00606 * (angle_c_moy(mois) - (pd_NOMBRE_PI / 2)) _
                                                     - (0.505 + 0.00455 * (angle_c_moy(mois) - (pd_NOMBRE_PI / 2)) * Cos(115 * ind_clarte_moy(mois) - 103)))

                'moyenne mensuelle de l'irradiation directe journalière
                S_irr_moy(mois) = G_irr_moy(mois) - D_irr_moy(mois)

                'moyenne mensuelle de la durée d'une journée [h]
                duree_j_moy(mois) = angle_c_moy(mois) * 2 * (180 / pd_NOMBRE_PI) / 15

                mois = mois + 1
            End If
        Else
            Exit Do
        End If
    Loop

    '## CALCUL DES PARAMETRES HORAIRES
    For mois = 1 To 12

        'inverse de la distance terre-soleil au jour moyen du mois
        d_invTS = 1 + 0.033 * Cos((2 * pd_NOMBRE_PI * num_jour_moyen(mois - 1)) / pd_NOMBRE_JOUR_PAR_AN)

        'angle de déclinaison solaire au jour moyen du mois
        decl = 0.409 * Sin(((2 * pd_NOMBRE_PI * num_jour_moyen(mois - 1)) / pd_NOMBRE_JOUR_PAR_AN) - 1.39)  '[rad]

        'angle de lever/coucher du soleil au jour moyen du mois
        angle_c = Arccos(-Tan(lat) * Tan(decl))     '[rad]

        'heure du lever et du coucher du soleil (temps solaire vrai)
        h_lever = 12 - (angle_c * 180 / (pd_NOMBRE_PI * 15))
        h_coucher = 12 + (angle_c * 180 / (pd_NOMBRE_PI * 15))

        'détermination de l'albédo
        If T_air_moy(mois) < -5 Then
            albedo = 0.7
        ElseIf T_air_moy(mois) > 0 Then
            albedo = 0.2
        Else
            albedo = -T_air_moy(mois) + 0.2
        End If

        'répartition horaire entre l'heure de lever et l'heure de coucher du soleil
        iterations = 0
        For heure = Format(h_lever, "#0.0") To Format(h_coucher, "#0.0") Step pas

            'angle horaire à la date heure = h_lever + heure (en heure) [rad]
            angle_h = 15 * (pd_NOMBRE_PI / 180) * (heure - 12)

            'angle correspondant à la hauteur solaire [rad]
            hauteur_S = Arcsin(Sin(lat) * Sin(decl) + Cos(lat) * Cos(decl) * Cos(angle_h))

            'angle azimut, projection horizontale de la position du soleil par rapport au sud [rad]
            azimut = Arcsin((Cos(decl) * Sin(angle_h)) / Cos(hauteur_S))

            'angle d'incidence des rayons sur un plan incliné et orienté [rad]
            angle_i = Arccos(Sin(inclinaison) * Cos(hauteur_S) * Cos(azimut - orientation) + Cos(inclinaison) * Sin(hauteur_S))

            'coefficient pour la projection
            Rb = Cos(angle_i) / Sin(hauteur_S)

            'CALCUL DES RAYONNEMENTS
            'eclairement global
            Dim a As Double
            Dim b As Double
            a = 0.409 + 0.502 * Sin(angle_c - (pd_NOMBRE_PI / 3))
            b = 0.661 - 0.477 * Sin(angle_c - (pd_NOMBRE_PI / 3))

            G_ecl = G_irr_moy(mois) * (pd_NOMBRE_PI / 24) * (a + b * Cos(angle_h)) * (Cos(angle_h) - Cos(angle_c)) / (Sin(angle_c) - angle_c * Cos(angle_c))
            G_ecl_moy(mois) = G_ecl_moy(mois) + G_ecl

            'eclairement diffus
            D_ecl = D_irr_moy(mois) * (pd_NOMBRE_PI / 24) * (Cos(angle_h) - Cos(angle_c)) / (Sin(angle_c) - angle_c * Cos(angle_c))
            D_ecl_moy(mois) = D_ecl_moy(mois) + D_ecl
            D_ecl_t = D_ecl * (1 + Cos(inclinaison)) / 2
            D_ecl_moy_t(mois) = D_ecl_moy_t(mois) + D_ecl_t

            'eclairement direct
            S_ecl = G_ecl - D_ecl
            S_ecl_moy(mois) = S_ecl_moy(mois) + S_ecl

            S_ecl_t = S_ecl * Rb
            S_ecl_moy_t(mois) = S_ecl_moy_t(mois) + S_ecl_t

            'eclairement réfléchi
            R_ecl_t = G_ecl * albedo * (1 - Cos(inclinaison)) / 2
            R_ecl_moy_t(mois) = R_ecl_moy_t(mois) + R_ecl_t

            'moyenne mensuelle de l'éclairement global transposé journalier
            G_ecl_moy_t(mois) = S_ecl_moy_t(mois) + D_ecl_moy_t(mois) + R_ecl_moy_t(mois)

            iterations = iterations + 1
        Next heure

        'calcul des moyennes
        D_ecl_moy(mois) = D_ecl_moy(mois) / iterations
        S_ecl_moy(mois) = S_ecl_moy(mois) / iterations
        G_ecl_moy(mois) = G_ecl_moy(mois) / iterations

        S_ecl_moy_t(mois) = S_ecl_moy_t(mois) / iterations
        D_ecl_moy_t(mois) = D_ecl_moy_t(mois) / iterations
        R_ecl_moy_t(mois) = R_ecl_moy_t(mois) / iterations
        G_ecl_moy_t(mois) = G_ecl_moy_t(mois) / iterations

        'CALCULS DES IRRADIATIONS MENSUELLES [kWh/m2/mois]
        S_irr_moy_t(mois) = S_ecl_moy_t(mois) * duree_j_moy(mois) * nb_jrs_mois(mois - 1)
        D_irr_moy_t(mois) = D_ecl_moy_t(mois) * duree_j_moy(mois) * nb_jrs_mois(mois - 1)
        R_irr_moy_t(mois) = R_ecl_moy_t(mois) * duree_j_moy(mois) * nb_jrs_mois(mois - 1)

        If G_ecl_moy_t(mois) <= 0 Then    ' dans certaines orientation du module, la valeur peut être négative, donc consiérée comme nulle.
            G_irr_moy_t(mois) = 0
        Else
            G_irr_moy_t(mois) = G_ecl_moy_t(mois) * duree_j_moy(mois) * nb_jrs_mois(mois - 1)
        End If

        G_irr_kwh_m2_an = G_irr_kwh_m2_an + G_irr_moy_t(mois)    '[kWh/m2/an]

        'APPLICATION DES RENDEMENTS DE L'INSTALLATION
        'température des modules
        T_module = T_air_moy(mois) + (NOCT - 20) * (G_ecl_moy_t(mois) * 1000 / 800)    ' modèle P. Nolay
        'rendement moyen
        rendement_moyen = (rendementModule / 100) * (1 - (coefTemp / 100) * (T_module - T_ref))

        'CALCUL DE LA PRODUCTION D'ELECTRICITE
        E_p(mois) = surfaceInstallee * G_irr_moy_t(mois) * rendement_moyen * (1 - (pertesSysteme) / 100)    '[kWh/mois]
        E_p_kwh_an = E_p_kwh_an + E_p(mois)    '[kWh/an]

    Next mois

    'Calcul des moyennes journalières
    G_irr_moy_kwh_m2_j = G_irr_kwh_m2_an / pd_NOMBRE_JOUR_PAR_AN
    electricite_produite_kwh_j = E_p_kwh_an / pd_NOMBRE_JOUR_PAR_AN

    Exit Sub
    '    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("69.7")
    pn_EffacementResultatsEnCours = True

End Sub

Sub ecriture_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil69.Range("PV_rendementModule").Cells(1, Colonne) = rendementModule
    Feuil69.Range("PV_NOCT").Cells(1, Colonne) = NOCT
    Feuil69.Range("PV_coefTemp").Cells(1, Colonne) = coefTemp
    Feuil69.Range("PV_pertesSysteme").Cells(1, Colonne) = pertesSysteme
    Feuil69.Range("PV_surfaceInstallee").Cells(1, Colonne) = surfaceInstallee
    Feuil69.Range("PV_orientation").Cells(1, Colonne) = orientation * (180 / pd_NOMBRE_PI)
    Feuil69.Range("PV_inclinaison").Cells(1, Colonne) = inclinaison * (180 / pd_NOMBRE_PI)
    Feuil69.Range("PV_solaire_proj_janv").Cells(1, Colonne) = G_irr_moy_t(1)
    Feuil69.Range("PV_solaire_proj_fevr").Cells(1, Colonne) = G_irr_moy_t(2)
    Feuil69.Range("PV_solaire_proj_mars").Cells(1, Colonne) = G_irr_moy_t(3)
    Feuil69.Range("PV_solaire_proj_avri").Cells(1, Colonne) = G_irr_moy_t(4)
    Feuil69.Range("PV_solaire_proj_mai").Cells(1, Colonne) = G_irr_moy_t(5)
    Feuil69.Range("PV_solaire_proj_juin").Cells(1, Colonne) = G_irr_moy_t(6)
    Feuil69.Range("PV_solaire_proj_juil").Cells(1, Colonne) = G_irr_moy_t(7)
    Feuil69.Range("PV_solaire_proj_aout").Cells(1, Colonne) = G_irr_moy_t(8)
    Feuil69.Range("PV_solaire_proj_sept").Cells(1, Colonne) = G_irr_moy_t(9)
    Feuil69.Range("PV_solaire_proj_oct").Cells(1, Colonne) = G_irr_moy_t(10)
    Feuil69.Range("PV_solaire_proj_nove").Cells(1, Colonne) = G_irr_moy_t(11)
    Feuil69.Range("PV_solaire_proj_dece").Cells(1, Colonne) = G_irr_moy_t(12)
    Feuil69.Range("PV_solaire_proj_kwh_m2_an").Cells(1, Colonne) = G_irr_kwh_m2_an
    Feuil69.Range("PV_solaire_proj_kwh_m2_j").Cells(1, Colonne) = G_irr_moy_kwh_m2_j
    Feuil69.Range("PV_elec_prod_janv").Cells(1, Colonne) = E_p(1)
    Feuil69.Range("PV_elec_prod_fevr").Cells(1, Colonne) = E_p(2)
    Feuil69.Range("PV_elec_prod_mars").Cells(1, Colonne) = E_p(3)
    Feuil69.Range("PV_elec_prod_avri").Cells(1, Colonne) = E_p(4)
    Feuil69.Range("PV_elec_prod_mai").Cells(1, Colonne) = E_p(5)
    Feuil69.Range("PV_elec_prod_juin").Cells(1, Colonne) = E_p(6)
    Feuil69.Range("PV_elec_prod_juil").Cells(1, Colonne) = E_p(7)
    Feuil69.Range("PV_elec_prod_aout").Cells(1, Colonne) = E_p(8)
    Feuil69.Range("PV_elec_prod_sept").Cells(1, Colonne) = E_p(9)
    Feuil69.Range("PV_elec_prod_oct").Cells(1, Colonne) = E_p(10)
    Feuil69.Range("PV_elec_prod_nove").Cells(1, Colonne) = E_p(11)
    Feuil69.Range("PV_elec_prod_dece").Cells(1, Colonne) = E_p(12)
    Feuil69.Range("PV_elec_prod_kwh_an").Cells(1, Colonne) = E_p_kwh_an
    Feuil69.Range("PV_elec_prod_kwh_j").Cells(1, Colonne) = electricite_produite_kwh_j


    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("69.8")
    pn_EffacementResultatsEnCours = True

End Sub

Sub effacer_resultats()

    If version_Developpement = False Then On Error GoTo erreur_non_geree

    'On reporte sur Excel les valeurs calculées
    Dim Colonne As Integer

    Colonne = 3 * pi_EtapeCalculConsideree + pi_FiliereConsideree

    Feuil69.Range("PV_rendementModule").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_NOCT").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_coefTemp").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_pertesSysteme").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_surfaceInstallee").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_orientation").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_inclinaison").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_janv").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_fevr").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_mars").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_avri").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_mai").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_juin").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_juil").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_aout").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_sept").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_oct").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_nove").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_dece").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_kwh_m2_an").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_solaire_proj_kwh_m2_j").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_janv").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_fevr").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_mars").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_avri").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_mai").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_juin").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_juil").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_aout").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_sept").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_oct").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_nove").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_dece").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_kwh_an").Cells(1, Colonne).ClearContents
    Feuil69.Range("PV_elec_prod_kwh_j").Cells(1, Colonne).ClearContents


    Exit Sub
    '    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("69.9")
    pn_EffacementResultatsEnCours = True

End Sub

Sub affich_remarque()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    '
    '
    '    Dim nb_remarque As Integer
    '    Dim i As Integer
    '
    '    nb_remarque = 1  'penser à modifier la capacité du vecteur "message_erreur"
    '
    '    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + Feuil9.Range("dec_simple_erreur_nom")
    '
    '    For i = 1 To nb_remarque
    '        If remarque(i) <> "" Then
    '            ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + remarque(i)
    '        End If
    '    Next i
    '
    '    ps_CommentairesCalculsGlobaux = ps_CommentairesCalculsGlobaux + vbCrLf + vbCrLf
    '
    Exit Sub
    '*********************************************************************
erreur_non_geree:

    Call gestion_erreur_non_geree("69.10")
    pn_EffacementResultatsEnCours = True

End Sub

Private Sub Class_Terminate()

    If pn_EffacementResultatsEnCours = False Then
        Call ecriture_resultats
        Call affich_remarque
    End If

End Sub



