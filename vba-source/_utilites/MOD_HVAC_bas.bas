Attribute VB_Name = "MOD_HVAC"
Option Explicit

Public Function interpolation_lineaire(param_12() As Double) As Variant
    ' entrée : 12 valeurs mensuelles à interpoler (considérées comme étant valable au 15 de chaque mois)
    Dim jour As Integer
    Dim mois As Integer
    Dim a As Double    ' coeff directeur
    Dim b As Double    ' ordonnée à l'origine
    Dim jours As Variant    ' somme cumulée des jours
    Dim periode As Integer
    Dim param_13(0 To 13) As Double
    Dim coeff(1 To 13, 1 To 2) As Double
    Dim parametres(1 To 365) As Double


    param_13(0) = param_12(12)
    param_13(13) = param_12(1)
    For mois = 1 To 12
        param_13(mois) = param_12(mois)
    Next mois


    jours = Array(-16, 15, 46, 74, 105, 135, 166, 196, 227, 258, 288, 319, 349, 380)    ' (0 à 13) hypothèse : tous les 15 du mois

    For periode = 1 To 13

        a = (param_13(periode) - param_13(periode - 1)) / (jours(periode) - jours(periode - 1))
        b = param_13(periode) - a * jours(periode)
        coeff(periode, 1) = a
        coeff(periode, 2) = b

    Next periode

    periode = 1
    For jour = 1 To 365
        If jour <= jours(periode - 1) Then
            parametres(jour) = coeff(periode - 1, 1) * jour + coeff(periode - 1, 2)
        Else
            periode = periode + 1
            parametres(jour) = coeff(periode - 1, 1) * jour + coeff(periode - 1, 2)
        End If
    Next jour

    interpolation_lineaire = parametres

End Function
Public Function caracteristiquesAir(Tair As Double, Patm As Double, humRel As Double) As Variant

    'Tair : température de l'air (°C)
    'Patm : pression atmosphérique (kPa)
    'humRel : humidité relative (%)

    '## DECLARATION DES VARIABLES
    Dim PeauSat As Double    ' pression de vapeur saturante (kPa)
    Dim PvapEau As Double    ' pression partielle de la vapeur d'eau (kPa)
    Dim humAbs As Double    ' humidité absolue (kg eau / kg air)
    Dim hAir As Double    ' enthalpie de l'air (kJ/kg)
    Dim resultats(1 To 4) As Double    'vecteur contenant l'ensemble des résultats calculés

    '## CALCULS
    '# Détermination de la pression de vapeur saturante (kPa)

    'PeauSat = Exp(23.1964 - 3816.44 / ((Tair + 273.15) - 46.13)) / 1000    'kPa , équation simplifiée

    'Avec les formules de Hyland and Wexler 1983b :
    Dim C1 As Double, C2 As Double, C3 As Double, C4 As Double, C5 As Double, C6 As Double, C7 As Double, C8 As Double, C9 As Double, C10 As Double, C11 As Double, C12 As Double, C13 As Double
    C1 = -5674.5359
    C2 = -0.51523058
    C3 = -0.009677843
    C4 = 0.00000062215701
    C5 = 2.0747825E-09
    C6 = -9.484024E-13
    C7 = 4.1635019
    C8 = -5800.2206
    C9 = -5.516256
    C10 = -0.048640239
    C11 = 0.000041764768
    C12 = -0.000000014452093
    C13 = 6.5459673

    If Tair < 0 Then    ' Tair de -100°C à 0°C
        Tair = Tair + 273.15    '°C -> °K
        PeauSat = Exp(C1 / Tair + C2 + C3 * Tair + C4 * Tair ^ 2 + C5 * Tair ^ 3 + C6 * Tair ^ 4 + C7 * Log(Tair))    ' (kPa)
    Else    '   Tair de 0°C à 200°C
        Tair = Tair + 273.15
        PeauSat = Exp(C8 / Tair + C9 + C10 * Tair + C11 * Tair ^ 2 + C12 * Tair ^ 3 + C13 * Log(Tair))    '(kPa)
    End If
    Tair = Tair - 273.15    '°K -> °C

    '# Détermination de la pression partielle de la vapeur d'eau (kPa)
    PvapEau = humRel / 100 * PeauSat    'kPa

    '# Détermination de l'humidité absolue (kg eau / kg air)
    humAbs = 0.62198 * PvapEau / (Patm - PvapEau)    'kg eau / kg air
    humAbs = humAbs * 1000    'g eau / kg air

    '# Détermination de l'enthalpie de l'air (kJ/kg)
    hAir = 1.006 * Tair + humAbs * (2501 + 1.805 * Tair)    'kJ/kg

    'écriture des résultats
    resultats(1) = PeauSat
    resultats(2) = PvapEau
    resultats(3) = humAbs
    resultats(4) = hAir

    caracteristiquesAir = resultats

End Function
Public Function degresJours(Chaud_ou_Froid As String, T_base_DJ As Double, T_air_mois() As Double, precision_Mois_ou_Jour As String) As Variant

    'on constitue des profils de température sur l'année en fonction de la température de référence
    
    Dim DJ_chaud() As Double
    Dim DJ_froid() As Double
    Dim DJchaudTot As Double
    Dim DJfroidTot As Double
    Dim nbJoursClim As Double
    Dim nbJoursChauff As Double
    Dim mois As Integer
    Dim jour As Integer
    Dim nb_jrs_mois As Variant
    Dim param_T_air(0 To 13) As Double
    Dim T_air_jour As Variant
    Dim DJ_param(1 To 2) As Double


    '## INITIALISATION DES VARIABLES
    nb_jrs_mois = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)    ' nombre de jours par mois pour chaque mois (total : 365)
    DJchaudTot = 0
    DJfroidTot = 0
    nbJoursChauff = 0
    nbJoursClim = 0

    If precision_Mois_ou_Jour = "Mois" Then    ' Calcul des DJ à partir des moyennes mensuelles (12 valeurs)

        ReDim DJ_chaud(1 To 12)
        ReDim DJ_froid(1 To 12)

        'Calcul des DJ chaud
        If Chaud_ou_Froid = "Chaud" Then
            For mois = 1 To 12
                If T_air_mois(mois) <= T_base_DJ Then
                    DJ_chaud(mois) = (T_base_DJ - T_air_mois(mois)) * nb_jrs_mois(mois - 1)
                    nbJoursChauff = nbJoursChauff + nb_jrs_mois(mois - 1)
                Else
                    DJ_chaud(mois) = 0
                End If
                DJchaudTot = DJchaudTot + DJ_chaud(mois)
            Next mois
            DJ_param(1) = DJchaudTot
            DJ_param(2) = nbJoursChauff
        End If

        'Calcul des DJ froid
        If Chaud_ou_Froid = "Froid" Then
            For mois = 1 To 12
                If T_air_mois(mois) >= T_base_DJ Then
                    DJ_froid(mois) = (T_air_mois(mois) - T_base_DJ) * nb_jrs_mois(mois - 1)
                    nbJoursClim = nbJoursClim + nb_jrs_mois(mois - 1)
                Else
                    DJ_froid(mois) = 0
                End If
                DJfroidTot = DJfroidTot + DJ_froid(mois)
            Next mois
            DJ_param(1) = DJfroidTot
            DJ_param(2) = nbJoursClim
        End If

    End If

    If precision_Mois_ou_Jour = "Jour" Then    'Calcul des DJ à partir des moyennes journalières (365 valeurs)

        '        'création du vecteur contenant les moyennes mensuelles
        '        param_T_air(0) = T_air_mois(12)
        '        param_T_air(13) = T_air_mois(1)
        '        For mois = 1 To 12
        '            param_T_air(mois) = T_air_mois(mois)
        '        Next mois

        'Interpolation des valeurs considérées au jour 15 de chaque mois pour chaque jour de l'année
        T_air_jour = interpolation_lineaire(T_air_mois)

        'Calcul des DJ chaud
        If Chaud_ou_Froid = "Chaud" Then
            ReDim DJ_chaud(1 To 365)
            For jour = 1 To 365
                If T_air_jour(jour) <= T_base_DJ Then
                    DJ_chaud(jour) = T_base_DJ - T_air_jour(jour)
                    nbJoursChauff = nbJoursChauff + 1
                Else
                    DJ_chaud(jour) = 0
                End If
                DJchaudTot = DJchaudTot + DJ_chaud(jour)
            Next jour
            DJ_param(1) = DJchaudTot
            DJ_param(2) = nbJoursChauff
        End If

        'Calcul des DJ froids
        If Chaud_ou_Froid = "Froid" Then
            ReDim DJ_froid(1 To 365)
            For jour = 1 To 365
                If T_air_jour(jour) >= T_base_DJ Then
                    DJ_froid(jour) = T_air_jour(jour) - T_base_DJ
                    nbJoursClim = nbJoursClim + 1
                Else
                    DJ_froid(jour) = 0
                End If
                DJfroidTot = DJfroidTot + DJ_froid(jour)
            Next jour
            DJ_param(1) = DJfroidTot
            DJ_param(2) = nbJoursClim
        End If

    End If

    degresJours = DJ_param

End Function
Public Function grammesJours(Ete_ou_Hiver As String, T_ambiante As Double, Patm_ambiante As Double, HR_base_GJ As Double, T_air_mois() As Double, Patm_mois() As Double, HR_mois() As Double, precision_Mois_ou_Jour As String) As Variant

    Dim GJ_hiver() As Double
    Dim GJ_ete() As Double
    Dim GJeteTot As Double
    Dim GJhiverTot As Double
    Dim nbJoursDeshumEte As Double
    Dim nbJoursHumidHiver As Double
    Dim mois As Integer
    Dim jour As Integer
    Dim nb_jrs_mois As Variant
    Dim param_HA(0 To 13) As Double
    Dim HA_jour_g_kg As Variant
    Dim GJ_param(1 To 2) As Double
    Dim HA_mois_g_kg(1 To 12) As Double
    Dim HA_base_GJ_g_kg As Double
    Dim air As Variant

    '## INITIALISATION DES VARIABLES
    nb_jrs_mois = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)    ' nombre de jours par mois pour chaque mois (total : 365)
    GJeteTot = 0
    GJhiverTot = 0
    nbJoursHumidHiver = 0
    nbJoursDeshumEte = 0

    'Conversion HR en hum absolue
    air = caracteristiquesAir(T_ambiante, Patm_ambiante, HR_base_GJ)
    HA_base_GJ_g_kg = air(3)
    For mois = 1 To 12
        air = caracteristiquesAir(T_air_mois(mois), Patm_mois(mois), HR_mois(mois))
        HA_mois_g_kg(mois) = air(3)
    Next mois

    If precision_Mois_ou_Jour = "Mois" Then    ' Calcul des DJ à partir des moyennes mensuelles (12 valeurs)

        ReDim GJ_hiver(1 To 12)
        ReDim GJ_ete(1 To 12)

        'Calcul des GJ hiver
        If Ete_ou_Hiver = "Hiver" Then
            For mois = 1 To 12
                If HA_mois_g_kg(mois) <= HA_base_GJ_g_kg Then
                    GJ_hiver(mois) = (HA_base_GJ_g_kg - HA_mois_g_kg(mois)) * nb_jrs_mois(mois - 1)
                    nbJoursHumidHiver = nbJoursHumidHiver + nb_jrs_mois(mois - 1)
                Else
                    GJ_hiver(mois) = 0
                End If
                GJeteTot = GJeteTot + GJ_hiver(mois)
            Next mois
            GJ_param(1) = GJeteTot
            GJ_param(2) = nbJoursHumidHiver
        End If

        'Calcul des GJ été
        If Ete_ou_Hiver = "Ete" Then
            For mois = 1 To 12
                If HA_mois_g_kg(mois) >= HA_base_GJ_g_kg Then
                    GJ_ete(mois) = Abs(HA_base_GJ_g_kg - HA_mois_g_kg(mois)) * nb_jrs_mois(mois - 1)
                    nbJoursDeshumEte = nbJoursDeshumEte + nb_jrs_mois(mois - 1)
                Else
                    GJ_ete(mois) = 0
                End If
                GJhiverTot = GJhiverTot + GJ_ete(mois)
            Next mois
            GJ_param(1) = GJhiverTot
            GJ_param(2) = nbJoursDeshumEte
        End If

    End If

    If precision_Mois_ou_Jour = "Jour" Then    'Calcul des DJ à partir des moyennes journalières (365 valeurs)

        '        'création du vecteur contenant les moyennes mensuelles
        '        param_HA(0) = HA_mois_g_kg(12)
        '        param_HA(13) = HA_mois_g_kg(1)
        '        For mois = 1 To 12
        '            param_HA(mois) = HA_mois_g_kg(mois)
        '        Next mois

        'Interpolation des valeurs considérées au jour 15 de chaque mois pour chaque jour de l'année
        HA_jour_g_kg = interpolation_lineaire(HA_mois_g_kg)

        'Calcul des GJ hiver
        If Ete_ou_Hiver = "Hiver" Then
            ReDim GJ_hiver(1 To 365)
            For jour = 1 To 365
                If HA_jour_g_kg(jour) <= HA_base_GJ_g_kg Then
                    GJ_hiver(jour) = HA_base_GJ_g_kg - HA_jour_g_kg(jour)
                    nbJoursHumidHiver = nbJoursHumidHiver + 1
                Else
                    GJ_hiver(jour) = 0
                End If
                GJeteTot = GJeteTot + GJ_hiver(jour)
            Next jour
            GJ_param(1) = GJeteTot
            GJ_param(2) = nbJoursHumidHiver
        End If

        'Calcul des GJ été
        If Ete_ou_Hiver = "Ete" Then
            ReDim GJ_ete(1 To 365)
            For jour = 1 To 365
                If HA_jour_g_kg(jour) >= HA_base_GJ_g_kg Then
                    GJ_ete(jour) = Abs(HA_base_GJ_g_kg - HA_jour_g_kg(jour))
                    nbJoursDeshumEte = nbJoursDeshumEte + 1
                Else
                    GJ_ete(jour) = 0
                End If
                GJhiverTot = GJhiverTot + GJ_ete(jour)
            Next jour
            GJ_param(1) = GJhiverTot
            GJ_param(2) = nbJoursDeshumEte
        End If

    End If

    grammesJours = GJ_param

End Function
Public Function moy_6max_6min(param_12() As Double) As Variant
    Dim tri_ok As Boolean
    Dim mois As Integer
    Dim param_decroissant(1 To 12) As Double
    Dim tampon As Double
    Dim vect_resultats(1 To 2) As Double

    '    param_decroissant = param_12

    'tri décroissant
    Do
        tri_ok = True
        For mois = 1 To 11
            If param_12(mois) < param_12(mois + 1) Then
                tampon = param_12(mois)
                param_12(mois) = param_12(mois + 1)
                param_12(mois + 1) = tampon
                tri_ok = False
            End If
        Next
    Loop While tri_ok = False

    'calcul des moyennes
    For mois = 1 To 6
        vect_resultats(1) = vect_resultats(1) + param_12(mois)
    Next
    For mois = 7 To 12
        vect_resultats(2) = vect_resultats(2) + param_12(mois)
    Next
    vect_resultats(1) = vect_resultats(1) / 6
    vect_resultats(2) = vect_resultats(2) / 6

    moy_6max_6min = vect_resultats

End Function

Public Function moy_6ete_6hiver(T_mois12() As Double, HR_mois12() As Double, Patm_mois12() As Double) As Variant
    Dim tri_ok As Boolean
    Dim mois As Integer
    Dim parametres(1 To 12, 1 To 4) As Double
    Dim tampon(1 To 4) As Double
    Dim tab_resultats(1 To 3, 1 To 2) As Double

    For mois = 1 To 12
        parametres(mois, 1) = mois
        parametres(mois, 2) = T_mois12(mois)
        parametres(mois, 3) = HR_mois12(mois)
        parametres(mois, 4) = Patm_mois12(mois)
    Next mois

    '    param_decroissant = param_12

    'tri décroissant
    Do
        tri_ok = True
        For mois = 1 To 11
            If parametres(mois, 2) < parametres(mois + 1, 2) Then
                tampon(1) = parametres(mois, 1)
                tampon(2) = parametres(mois, 2)
                tampon(3) = parametres(mois, 3)
                tampon(4) = parametres(mois, 4)

                parametres(mois, 1) = parametres(mois + 1, 1)
                parametres(mois, 2) = parametres(mois + 1, 2)
                parametres(mois, 3) = parametres(mois + 1, 3)
                parametres(mois, 4) = parametres(mois + 1, 4)

                parametres(mois + 1, 1) = tampon(1)
                parametres(mois + 1, 2) = tampon(2)
                parametres(mois + 1, 3) = tampon(3)
                parametres(mois + 1, 4) = tampon(4)

                tri_ok = False
            End If
        Next
    Loop While tri_ok = False

    'calcul des moyennes
    For mois = 1 To 6
        tab_resultats(1, 1) = tab_resultats(1, 1) + parametres(mois, 2)
        tab_resultats(2, 1) = tab_resultats(2, 1) + parametres(mois, 3)
        tab_resultats(3, 1) = tab_resultats(3, 1) + parametres(mois, 4)
    Next
    For mois = 7 To 12
        tab_resultats(1, 2) = tab_resultats(1, 2) + parametres(mois, 2)
        tab_resultats(2, 2) = tab_resultats(2, 2) + parametres(mois, 3)
        tab_resultats(3, 2) = tab_resultats(3, 2) + parametres(mois, 4)
    Next
    tab_resultats(1, 1) = tab_resultats(1, 1) / 6
    tab_resultats(2, 1) = tab_resultats(2, 1) / 6
    tab_resultats(3, 1) = tab_resultats(3, 1) / 6
    tab_resultats(1, 2) = tab_resultats(1, 2) / 6
    tab_resultats(2, 2) = tab_resultats(2, 2) / 6
    tab_resultats(3, 2) = tab_resultats(3, 2) / 6

    moy_6ete_6hiver = tab_resultats

End Function

Function chargeChauffage()
    '                Dim airInterieur As Variant
    '                Dim airExterieur As Variant
    '                Dim TAirExt As Double
    '                Dim TAirInt As Double
    '                Dim humRelExt As Double
    '                Dim humRelInt As Double
    '                Dim hAirExt As Double
    '                Dim hAirInt As Double
    '                Dim Patm As Double
    '                Dim UAbat As Double    'kW/K
    '                Dim PuissanceDesignChauffage As Double    'kW
    '
    '                '## PROCEDURE
    '                TAirExt = -3.8
    '                TAirInt = 19
    '                humRelExt = 77
    '                humRelInt = 50
    '                Patm = 100
    '                UAbat = 2.25
    '
    '                airExterieur = caracteristiquesAir(tempDesignEte, Patm, humRelExt)
    '                airInterieur = caracteristiquesAir(TAirInt, Patm, humRelInt)
    '
    '                hAirExt = airExterieur(4)
    '                hAirInt = airInterieur(4)
    '
    '                PuissanceDesignChauffage = UAbat * (TAirInt - TAirExt) + Qair * (hAirInt - hAirExt)
End Function
