Attribute VB_Name = "A2_Traitement_curatif_sulfure"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
'Option Explicit
Public Mat_TCS_conc_HS
Public Mat_TCS_L_conduite
Public Mat_rdt02
Public Mat_TCS_reactif

Sub choix_traitement_sulfures_curatif(I, Q_nominal, HS_nominal_mgL)


choix_reactif_utilise = Range("btn_sulfures_curatif_reactif_" & I).Text
choix_oxygenateur = Range("btn_sulfures_curatif_O2_oxygenateur_" & I).Text

L_conduite = Range("Mat_TCS_L_conduite_" & I)
rdtO2 = Range("Mat_TCS_rdtO2_" & I)
dosage_reactif_pur = Range("Mat_TCS_reactif_" & I)
curatif_outlet = Range("Mat_TCS_conc_HS_" & I)

'on vient lire les valeurs imposées en sortie
If curatif_outlet(1, 2) <> "" Then curatif_outlet_HS = curatif_outlet(1, 2)
If curatif_outlet(1, 2) = "" Then curatif_outlet_HS = 5 'mg/L


'on calcule les S éliminés
SH_nominal = HS_nominal_mgL * Q_nominal / 1000
HS_removed = SH_nominal - curatif_outlet_HS * Q_nominal / 1000


'on définit les hypothèses en fonction des options pour déterminer la consommation de réactif
Select Case choix_reactif_utilise

    Case "injection d'oxygène pur" 'O2 oxy pure

        curatif_reactif_dose_pur_g_gHS = 2 'hyp_dose_O2_g_gS   'gO2/gS-SHdose théorique d'O2   stoechiométrie
     
            Select Case choix_oxygenateur
                Case "oxygénateur"
                curatif_O2_rdt_transfert = 0.99
                
                Case "injection directe dans le réseau"
                If L_conduite(1, 2) = "" Then curatif_O2_longueur_conduite = 900 'm
                If L_conduite(1, 2) <> "" Then curatif_O2_longueur_conduite = L_conduite(1, 2)
            
                If curatif_O2_longueur_conduite < 100 Then curatif_O2_rdt_transfert = 0.5
                If curatif_O2_longueur_conduite > 900 Then curatif_O2_rdt_transfert = 0.9
                If curatif_O2_longueur_conduite >= 100 And curatif_O2_longueur_conduite <= 900 Then curatif_O2_rdt_transfert = 0.5 + (curatif_O2_longueur_conduite - 100) / (900 - 100) * (0.9 - 0.5)              '50% si conduite de 100m    90% sinon
            End Select
        
        If rdtO2(1, 2) = "" Then curatif_O2_rdt_transfert = curatif_O2_rdt_transfert
        If rdtO2(1, 2) <> "" Then curatif_O2_rdt_transfert = rdtO2(1, 2)
         
        curatif_reactif_dose_pur_g_gHS = curatif_reactif_dose_pur_g_gHS / curatif_O2_rdt_transfert
        curatif_reactif_pur_kgj = curatif_reactif_dose_pur_g_gHS * HS_removed
         
        'curatif_reactif_commercial_kgj = curatif_reactif_pur_kgj
        O2_liquide_kgj = curatif_reactif_pur_kgj
        H2O2_kgj = 0
        FeCl3pur_kgj = 0
        
        Mat_TCS_L_conduite = Array(900, curatif_O2_longueur_conduite, curatif_O2_longueur_conduite)
        Mat_rdt02 = Array(0.99, curatif_O2_rdt_transfert, curatif_O2_rdt_transfert)
        
    Case "peroxyde d'hydrogène" 'H2O2 peroxyde
    curatif_reactif_dose_pur_g_gHS = 1.5    'gH2O2/gS-SH   stoechiométrie=34/32   HS- + H2O2 + H+ --> 1/8*S8 + 2*H2O


    curatif_reactif_pur_kgj = curatif_reactif_dose_pur_g_gHS * HS_removed
    'curatif_reactif_commercial_kgj = curatif_reactif_pur_kgj / curatif_reactif_commercial_purete
    H2O2_kgj = curatif_reactif_pur_kgj
    O2_liquide_kgj = 0
    FeCl3pur_kgj = 0

    Case "chlorure ferrique" 'FECL3 clhrlo ferrique
    curatif_reactif_dose_pur_g_gHS = 3.42     'gFeCl3/gS-SH    2*FeCl3 + 3*HS- --> 2*FeS + S + 6*Cl- + 3*H+

    curatif_reactif_pur_kgj = curatif_reactif_dose_pur_g_gHS * HS_removed
    'curatif_reactif_commercial_kgj = curatif_reactif_pur_kgj / curatif_reactif_commercial_purete
    FeCl3pur_kgj = curatif_reactif_pur_kgj
    O2_liquide_kgj = 0
    H2O2_kgj = 0




End Select


Mat_TCS_conc_HS = Array(5, curatif_outlet_HS, curatif_outlet_HS)
Mat_TCS_reactif = Array(FeCl3pur_kgj, O2_liquide_kgj, H2O2_kgj)

'
'
'
'       If curatif_outlet_HS >= (SH_nominal / Q_nominal * 1000) Then
'            'pas besoin de traitement
'            curatif_outlet_HS = SH_nominal / Q_nominal * 1000
'            If curatif_outlet_HS_force = True And (curatif_outlet_HS_f <> curatif_outlet_HS) Then
'                'message pour dire qu'impossible
'            End If
'        Else
'
'
'        End If
    
    
   ' SH_traite_kgj = SH_nominal - curatif_outlet_HS * Q_nominal / 1000

End Sub

