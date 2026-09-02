Attribute VB_Name = "PINCH"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
'   __      __       _ _____ _            _        v1.1
'   \ \    / /      (_)  __ (_)          | |
'    \ \  / /__ _ __ _| |__) | _ __   ___| |__
'     \ \/ / _ \ '__| |  ___/ | '_ \ / __| '_ \
'      \  /  __/ |  | | |   | | | | | (__| | | |
'       \/ \___|_|  |_|_|   |_|_| |_|\___|_| |_|

'   Algorithme d'analyse d'intégration thermique "VeriPinch"
'   Veolia Environnement Recherche et Innovation - Clément Rebillard - clement.rebillard@veolia.com
'   Début du développement :  07/11/2011
'   Version actuelle :        26/10/2012

'Description générale
'   L'analyse Pinch (Intégration Energétique) a pour objectif de calculer la récupération énergétique entre
'   les flux du Procédé, au moyen d'échangeurs de chaleur, et de réaliser la couverture des besoins
'   énergétiques restants grâce à des Utilités.

'Le Pinch est décomposé en 3 grandes composantes :
'   I. Construction des Courbes Composites (CC)
'           Calcul du point de Pincement, de la récupération énergétique (kW) maximale, des besoins
'           énergétiques à couvrir par les Utilités ;
'   II. Construction du réseau d'échangeur de chaleur (HEN)
'           Création des échanges de chaleur entre les flux de procédé. L'utilisateur peut forcer un
'           échange entre 2 flux ou au contraire l'interdire, pour représenter par exemple une contrainte
'           existant sur le site (éloignement des flux, viscosité ...) ;
'   III. Couverture des besoins par les Utilités
'           Couverture des besoins restants à partir d'une liste d'utilités définie par l'utilisateur.

'ATTENTION
'- pour optimiser l'analyse pinch, il faut que les flux de procédé soient triés : flux chauds en premiers, triés par température d'entrée décroissante ; puis les flux froids, triés par température de sortie décroissante.
'- pour les flux de procédé présentant un changement d'état, il est nécessaire d'avoir un faible DT (0.1 K par exemple) entre les 2 flux (Tin doit être différent de Tout !!!). Le MCp est alors calculé à partir de ce faible DT et de la puissance.
'- quelques bugs peuvent demeurer
'- quelques indications sont repérées par le commentaire "ATTENTION" (en majuscules) : à vérifier

'I. Tracé des Courbes Composites
'I.0. --------------------------------   Initialisations générales   -----------------------------------
'I.1. ------------------------------------   Remises à Zero   ------------------------------------------
'I.2. -------------------   Rangement des flux selon leur type : Chaud-Froid   -------------------------
'I.3. ----------------   Traitement des courants Chauds/Froids (tri, doublons etc.)   ------------------
'I.4. --------------------   Calcul du décalage nécessaire pour vérifier le Pinch   --------------------
'I.5. -----------------------   Calcul du besoin Chaud et Froid, Pinch et DTMin   ----------------------
'I.6. ------------------------------   Tracé des Courbes Composites   ----------------------------------
'I.7. ---------------------------------   Construction du HEN (vide)  ----------------------------------
'I.8. ---------------------------------------   Finalization   -----------------------------------------
'I.9. -------------------------------------------   Fin   ----------------------------------------------

'II. Construction du HEN
'II.0. ---------------------------------   Initialisation générale   ------------------------------------
'II.1. ---------------------------   Lecture des données (si nécessaire)   ------------------------------
'II.2. ------------------------------   Remise à Zero du HEN existant    --------------------------------
'II.3. ----------   Création des variables contenant les HEX à interdire et à supprimer   ---------------
'II.4. --------------------------------    Ajout des HEX manuellement    --------------------------------
'II.5. -------------------------------    Ajout des HEX automatiquement    ------------------------------
'II.6. --------------------------------------    Finalization    ----------------------------------------
'II.7. --------------------------------------------   Fin   ---------------------------------------------

'III. Gestion des Utilités
'III.0. -----------------------------------   Initialisations générales   --------------------------------
'III.1. -----------------------------   Lecture des données (si nécessaire)   ----------------------------
'III.2. -----------------------------------------   Remises à Zero   -------------------------------------
'III.3. ---------------------------   Renseignement de la liste des Utilités   ---------------------------
'III.4. --------------------------------   Liste des besoins restants   ----------------------------------
'III.5. -----------------------------------   Couverture des besoins   -----------------------------------
'III.6. -----------------------------------------   Nettoyage   ------------------------------------------
'III.7. ---------------------------------------   Affinage HEN   -----------------------------------------
'III.8. ----------------------------------------   Export HEN   ------------------------------------------
'III.9. ---------------------------------------   Finalization   -----------------------------------------
'III.10. ------------------------------------------   Fin.   ---------------------------------------------

'IV. Finalisation générale et export
'IV.1. --------------------------------------   Affinage du HEN   ---------------------------------------
'IV.2. ------------------------   Finalisation générales des besoins restants   -------------------------
'IV.3. ------------------   Affichage des grandeurs numériques sur la feuille Tampon   ------------------
'IV.4. -------------------   Nettoyage de la grille de flux en prévision de l'export   ------------------
'IV.5. -----------------------------------   Export graphique du HEN   ----------------------------------
'IV.6. --------------------------------------------   Fin.   --------------------------------------------




'GetCols
Dim colpinch As Integer         'indique la colonne du Pinch dans le tableau de flux data_flux_hen()
Dim colreste As Integer         'indique la "dernière" (en fait "Dh") colonne du tableau de flux data_flux_hen()

'Interdiction de HEX
Dim nbhexforbidden As Integer           'nombre de HEX à interdire
Dim data_hex_forbidden() As String      'caractéristiques des HEX à interdire
Dim nbheximposed As Integer             'nombre de HEX à imposer

'Construction du HEN
Dim ok As Boolean               'indique si le HEX est ok (validé après VerifierData)
Dim tinch As Double             'température d'entrée du flux chaud
Dim toutch As Double            'température de sortie du flux chaud
Dim tinfr As Double             'température d'entrée du flux froid
Dim toutfr As Double            'température de sortie du flux froid
Dim puissance As Double         'puissance échangée dans un HEX
Dim colplacehexfr As Integer    'colonne du flux chaud où on place le HEX
Dim colplacehexch As Integer    'colonne du flux froid où on place le HEX
Dim nbhex As Integer            'nombre de HEX construits
Dim nbhex_stocke As Integer
Dim surftot As Double           'Surface totale d'échange
Dim erreur As String            'erreur rencontrée lors de la construction du hen et/ou de l'ajout des utilités
Dim point_pinch As Double       'température du point de pincement
Dim pinch_mesure As Double      'valeur de pinch souhaité par l'utilisateur
Dim pinch_target As Double      'valeur de pinch souhaité par l'utilisateur
Dim ligprint As Integer             'ligne d'écriture des résultats dans feuille "Tampon")
Dim nbflux_final As Integer    'nombre de flux total (nbflux + nbutilities_updated)
Dim data_flux_utilities() As String         'contient les tous flux d'utilité, avec le meme format que data_flux_hen
Dim data_flux_utilities_updated() As String         'ne contient que les utilités auxquelles on a recours (sera affiché sous data_flux_hen)
Dim nbutilities As Integer      'nombre d'utilités disponibles
Dim nbutilities_updated As Integer  'nombre d'utilités utilisées
Dim data_utilities() As String
Dim CoefU As Double                 'coefficient U lors du calcul de la surface d'échange (creerHEX)

Dim data_flux_hen() As String       'grille de flux pour construction du HEN
Dim data_flux_all() As String       'grille de tous les flux pour affichage du HEN
Dim data_flux_all_final() As String       'grille de tous les flux pour affichage du HEN, après affinage
'| Type de flux | Nom | Tmax | (rien) | t_pinch_gch | t_pinch_gch | (rien) | tmin | Nom | Dh | Mcp | puissance thermique à gauche | puissance à droite | reste à gauche | reste à droite | match |
Dim data_flux_hen_size_empty As Integer    'valeur du HEN vide
Dim data_flux_hen_size As Integer   'nombre de colonnes de la grille de flux (étendu à chaque ajout de HEX)

Dim data_hen() As String            'récapitulatif des HEX ajoutés
'| HEX # | Puissance | Flux Chaud | Flux Froid | Tin Ch | Tout Ch | TinFr | Tout Fr | erreur |
Dim data_hen_size As Integer        'nombre de colonnes


'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'****************************************************************
'Ajouté par DCA
'****************************************************************
'marqueur pour les codenames
Dim marqueur_etape As Integer
'pour pointer sur les feuilles par leur code_name
Dim CodeNames As New Collection
'besoins restants
Public besoins_HOT As Double
Public besoins_COLD As Double
'Feuille active pour la resélectionner après
Dim feuille_active As Worksheet
'Variable permettant de conserver le nombre d'échangeurs entre flux de procédés et le nombre de flux
Dim nb_echangeurs_process As Integer
Dim nbflux As Integer               'nombre de lfux
'écritures de Clément
Private Const repere_traits_Clement As String = "------"
Dim data_flux_hen_stocke() As String       'grille de flux pour construction du HEN
Dim data_flux_hen_size_stockee As Integer
'paramètres des graphiques
Private Const hauteur_points = 300 '250
Private Const largeur_points = 500
Dim graph_CC_sans_decalage_top_points As Double
Dim graph_CC_sans_decalage_left_points As Double
Dim graph_CC_avec_decalage_top_points As Double
Dim graph_CC_avec_decalage_left_points As Double
'échelle sur les abcisses et ordonnées des CC
Private Const occupation_max_courbes = 0.85    'pour augmenter la borne max lorsque les courbes s'en approchent trop
Dim graph_abscisse_max As Double
Dim graph_ordonnee_max As Double
Dim graph_abscisse_min As Double
Dim graph_ordonnee_min As Double
'Position dans le tableau contenant les données sur des flux process
Private Const flux_process_nb_parametres = 5
Private Const flux_process_nom = 1
Private Const flux_process_T_in_°C = 2
Private Const flux_process_T_out_°C = 3
Private Const flux_process_MCp_kW_°C = 4
Private Const flux_process_Pw_kW = 5
'Position dans le tableau contenant les données sur des flux d'utilités
Private Const flux_utilites_nb_parametres = 7   'paramètres 6 et 7 existent mais aucune idée de ce qu'ils représentent
Private Const flux_utilites_nom = 1
Private Const flux_utilites_type = 2
Private Const flux_utilites_T_in_°C = 3
Private Const flux_utilites_MCp_kW_°C = 4
Private Const flux_utilites_Pw_kW = 5
'Représentation du réseau d'échangeurs
Private Const taille_case = 3  'nb de cellules d'une case (un échange)
Dim echangeurs() As Variant
Dim nb_echangeurs As Integer
Private Const echangeur_nb_caracteristique = 9
Private Const repere_echangeur_FR_nom = 1
Private Const repere_echangeur_FR_Tin = 2
Private Const repere_echangeur_FR_Tout = 3
Private Const repere_echangeur_CH_nom = 4
Private Const repere_echangeur_CH_Tin = 5
Private Const repere_echangeur_CH_Tout = 6
Private Const repere_echangeur_puissance = 7
Private Const repere_echangeur_surface = 8
Private Const repere_echangeur_force = 9
Dim puissances_non_echangees() As Variant
Dim nb_puissance_restants As Integer
Private Const P_non_echangee_nb_caracteristiques = 4
Private Const P_non_echangee_nom_flux = 1
Private Const P_non_echangee_puisssance = 2
Private Const P_non_echangee_T_chaude = 3
Private Const P_non_echangee_T_froide = 4
Private Const colonne_pinch = -9999
Dim T_pinch_reseau_DCA As Double
'Gestion des flèches
Private Const conversion_largeur_cellule_point = 1   ' 141 / 17 * 14.25 / 19
Private Const decalage_fleche_point = 10
Private Const epaisseur_fleche_point = 45
Private Const largeur_zone_texte = 300
Private Const hauteur_zone_texte = 50
Private Const tableau_largeur_colonne_fr = 50
Private Const tableau_largeur_colonne_ch = 15
Private Const tableau_hauteur_ligne = 18
'rangement des flux process
Private Const process_flux_chauds_premier = True
Private Const utilites_flux_chauds_premier = True
'Paramètres pour le tracé des CC dans OCEAN
Dim graph_CC_sans_decalage_largeur_points As Double
Dim graph_CC_sans_decalage_hauteur_points As Double
Dim graph_CC_avec_decalage_largeur_points As Double
Dim graph_CC_avec_decalage_hauteur_points As Double
Private Const DCA_CC_TaillePoints As Integer = 7
Private Const DCA_CC_ColorIndexBesoinsChauds As Integer = 3
Private Const DCA_CC_ColorIndexBesoinsFroids As Integer = 32
Private Const DCA_CC_Legende_TaillePolice As Integer = 14
Private Const DCA_CC_Titre_TaillePolice As Integer = 14
Private Const DCA_CC_Etiquette_TaillePolice As Integer = 12
'Paramètres pour construire le réseau DCA
Private Const Top_debut_reseau As Double = 200
Private Const Left_debut_reseau As Double = 300
Private Const Ecart_colonne_reseau As Double = 20
Private Const Ecart_ligne_reseau As Double = 15
Private Const Largeur_colonne As Double = 80
Private Const Largeur_espace_supplementaire_etiquette_flux As Double = 20
Private Const Largeur_etiquette_flux As Double = 150
Private Const Hauteur_ligne As Double = 45
Private Const Hauteur_marge_rectangle As Double = 5
Private Const Largeur_marge_rectangle As Double = 5
Private Const Reseau_DCA_Taille_Police As Double = 11
Dim Reseau_DCA_Hauteur_Maxi As Double
Dim Reseau_DCA_Largeur_Maxi As Double
Dim RecuperationEnergetiquePossible As Boolean
Dim separateur_decimal As String
Dim mn_ChangementLangue As Boolean
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Public Sub ecrire_reseau_DCA(ByRef Matrice() As Integer, ByRef flux_chaud() As Boolean)

    Dim i_flux As Integer
    Dim i_flux_2 As Integer
    Dim i_echangeur As Integer
    Dim i_echangeur_2 As Integer
    Dim numero_echangeur As Integer
    Dim correspondance_trouvee As Boolean
    Dim echangeur_represente() As Boolean
    Dim Ltotal As Single            'longueur totale de la grille de flux au format graphique
    Dim Htotal As Single            'hauteur totale de la grille de flux au format graphique
    Dim hactuel As Single           'longueur du graphique créé (sera ajustée)
    Dim lactuel As Single           'hauteur du graphique créé (sera ajustée)
    Dim graphique As ChartObject   'modif DCA
    Dim nb_flux_process As Integer
    Dim ZoomFeuille As Double

    nb_flux_process = UBound(flux_thermique_process, 2)

    ReDim echangeur_represente(UBound(Matrice, 2))
      
    Reseau_DCA_Largeur_Maxi = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_reseau_echangeurs_fin_" & marqueur_etape & pi_FiliereConsideree).left
        Reseau_DCA_Largeur_Maxi = Reseau_DCA_Largeur_Maxi - CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_reseau_echangeurs_debut_" & marqueur_etape & pi_FiliereConsideree).left
    Reseau_DCA_Hauteur_Maxi = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_reseau_echangeurs_fin_" & marqueur_etape & pi_FiliereConsideree).top
        Reseau_DCA_Hauteur_Maxi = Reseau_DCA_Hauteur_Maxi - CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_reseau_echangeurs_debut_" & marqueur_etape & pi_FiliereConsideree).top

    Tabelle9.DrawingObjects.Delete
    Tabelle9.Activate

    Tabelle9.Shapes.AddShape(msoShapeRectangle, Left_debut_reseau - Ecart_colonne_reseau - Largeur_espace_supplementaire_etiquette_flux - Largeur_etiquette_flux - Largeur_marge_rectangle, Top_debut_reseau - Ecart_ligne_reseau - Hauteur_ligne - Hauteur_marge_rectangle, Ecart_colonne_reseau + 2 * (Largeur_espace_supplementaire_etiquette_flux + Largeur_etiquette_flux) + UBound(Matrice, 2) * (Ecart_colonne_reseau + Largeur_colonne) + 2 * Largeur_marge_rectangle, UBound(Matrice, 1) * (Ecart_ligne_reseau + Hauteur_ligne) + Hauteur_ligne + 2 * Hauteur_marge_rectangle).Select
    With Selection
        With .ShapeRange.fill
            .Solid
            .ForeColor.SchemeColor = 9
            .Transparency = 0
        End With
        With .ShapeRange.Line
            .Weight = 1
            .DashStyle = msoLineSolid
            .Style = msoLineSingle
            .Transparency = 0
            .Visible = msoTrue
            .ForeColor.SchemeColor = 0
            .BackColor.RGB = RGB(255, 255, 255)
            .Visible = msoTrue
        End With
        'A compléter avec les trucs de Clément
    End With

    'on itère sur les flux de procédés et d'utilités
    For i_flux = 1 To UBound(Matrice, 1)
        Call representation_flux_DCA(i_flux, UBound(Matrice, 2), flux_chaud)
    Next i_flux
    For i_flux = 1 To UBound(Matrice, 1)
        For i_echangeur = 1 To UBound(Matrice, 2)
            If Matrice(i_flux, i_echangeur) > 0 Then
                numero_echangeur = Matrice(i_flux, i_echangeur)
                If echangeur_represente(numero_echangeur) = False Then
                    'On recherche le binome
                    i_flux_2 = 1
                    correspondance_trouvee = False
                    Do
                        If i_flux_2 <> i_flux Then
                            i_echangeur_2 = 1
                            Do While i_echangeur_2 <= UBound(Matrice, 2) And correspondance_trouvee = False
                                If Matrice(i_flux_2, i_echangeur_2) = numero_echangeur Then
                                    correspondance_trouvee = True
                                Else
                                    i_echangeur_2 = i_echangeur_2 + 1
                                End If
                            Loop
                        End If
                        i_flux_2 = i_flux_2 + 1
                    Loop While i_flux_2 <= UBound(Matrice, 1) And correspondance_trouvee = False
                    i_flux_2 = i_flux_2 - 1
                    'On ajoute l'échangeur
                    Call representation_echangeur_DCA(i_flux, flux_chaud(), i_echangeur, numero_echangeur, i_flux_2, i_echangeur_2, Matrice, echangeur_represente)
                    'échangeur représenté
                    echangeur_represente(numero_echangeur) = True
                End If
            ElseIf Matrice(i_flux, i_echangeur) = colonne_pinch Then
                If i_flux = 1 Then
                    'On trace la ligne pour repérer le PINCH
                    Tabelle9.Shapes.AddLine(Left_debut_reseau + (i_echangeur - 1) * (Ecart_colonne_reseau + Largeur_colonne) + Largeur_colonne / 2, Top_debut_reseau, Left_debut_reseau + (i_echangeur - 1) * (Ecart_colonne_reseau + Largeur_colonne) + Largeur_colonne / 2, Top_debut_reseau + (nb_flux_process - 1) * (Ecart_ligne_reseau + Hauteur_ligne) + Hauteur_ligne).Select    'en changeant "+ 0.1", on règle la fin du repère
                    Selection.ShapeRange.Line.DashStyle = msoLineRoundDot
                    Selection.ShapeRange.Line.ForeColor.RGB = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePinch_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.color
                    Selection.ShapeRange.Line.Weight = 3
                    'On ajoute une textbox avec PINCH écrit dedans
                    Tabelle9.Shapes.AddTextbox(msoTextOrientationHorizontal, Left_debut_reseau + (i_echangeur - 1) * (Ecart_colonne_reseau + Largeur_colonne), Top_debut_reseau - Ecart_ligne_reseau - Hauteur_ligne, Largeur_colonne, Hauteur_ligne).Select
                    With Selection
                        .Characters.Text = "PINCH" & Chr(10) & "T=" & Round(point_pinch, 1) & "°C" & Chr(10) & ChrW(8710) & "T=" & Round(pinch_mesure, 1) & "°C"
                        .AutoSize = False
                        .VerticalAlignment = xlCenter
                        .HorizontalAlignment = xlCenter
                        .Font.Name = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePinch_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.Name            '"Arial"
                        .Font.FontStyle = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePinch_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.FontStyle
                        .Font.size = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePinch_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.size
                        .Font.ColorIndex = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePinch_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.ColorIndex
                        With .ShapeRange.fill
                            .Solid
                            .ForeColor.SchemeColor = 9
                            .Transparency = 0
                        End With
                        With .ShapeRange.Line
                            .Weight = 1
                            .DashStyle = msoLineSolid
                            .Style = msoLineSingle
                            .Transparency = 0
                            .Visible = msoTrue
                            .ForeColor.RGB = RGB(255, 255, 255)
                            .BackColor.RGB = RGB(255, 255, 255)
                            .Visible = msoTrue
                        End With
                        'A compléter avec les trucs de Clément
                    End With
                End If
            ElseIf Matrice(i_flux, i_echangeur) < 0 Then
                'Ce seront les puissances non-échangées qui resteront
                Call representation_puissance_non_echangee_DCA(i_flux, flux_chaud, i_echangeur, Abs(Matrice(i_flux, i_echangeur)), Matrice, echangeur_represente)
            End If
        Next i_echangeur
    Next i_flux

    Tabelle9.Shapes.SelectAll
    Selection.ShapeRange.Group.Name = "Group99"    'nom arbitraire, ne sert que temporairement
    Ltotal = ActiveSheet.Shapes("Group99").Width
    Htotal = ActiveSheet.Shapes("Group99").Height

    'Ajout d'un graphique
    'Tabelle9.Activate
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Activate
    Set graphique = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects("HENGraphique")
'    graphique.Select
    graphique.Visible = True

    If graphique.Chart.SeriesCollection.Count > 0 Then  'Suppression des séries parasites si présentes
        Do While graphique.Chart.SeriesCollection.Count > 0
            graphique.Chart.SeriesCollection(1).Delete
        Loop
    End If
    '    If pi_EtapeCalculConsideree = 1 Then   'différenciation des calculs "guide" des calculs "normaux"
    '        graphique.Chart.Location xlLocationAsObject, Name:="PINCH " & pi_FiliereConsideree & " guide"
    '    Else
    '        graphique.Chart.Location xlLocationAsObject, Name:="PINCH " & pi_FiliereConsideree
    '    End If
'    graphique.Chart.Parent.Name = "HENGraphique"
    'Ajustement des dimensions du graphique
    hactuel = ActiveSheet.Shapes("HENGraphique").Height
    lactuel = ActiveSheet.Shapes("HENGraphique").Width
    If Ltotal > Reseau_DCA_Largeur_Maxi Then
        Ltotal = Reseau_DCA_Largeur_Maxi
    End If
    If Htotal > Reseau_DCA_Hauteur_Maxi Then
        Htotal = Reseau_DCA_Hauteur_Maxi
    End If
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Shapes("HENGraphique").ScaleWidth (Ltotal / lactuel), msoFalse, msoScaleFromTopLeft
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Shapes("HENGraphique").ScaleHeight (Htotal / hactuel), msoFalse, msoScaleFromTopLeft

    'MODIF CRE 16/10/2012   Affichage dans PINCH X
    ActiveSheet.Shapes("HENGraphique").top = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_reseau_echangeurs_debut_" & marqueur_etape & pi_FiliereConsideree).top
    ActiveSheet.Shapes("HENGraphique").left = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_reseau_echangeurs_debut_" & marqueur_etape & pi_FiliereConsideree).left

    'Insertion des dessins dans le graphique et export sous forme d'image
    Tabelle9.Activate
    Tabelle9.Shapes("Group99").Select
    Selection.Copy
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Activate
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Shapes("HENGraphique").Select
    graphique.Chart.Paste
    
    'graphique.Chart.Export FileName:="VeriPinchHEN  (pi_FiliereConsideree " & pi_FiliereConsideree & ").jpg", FilterName:="JPEG"    'on peut exporter le graphique sous forme d'image   'MODIF CRE 20/09/2012   Export du graphique sous forme d'image rendu inactif
    If Ltotal < Reseau_DCA_Largeur_Maxi Then
        CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Shapes("HENGraphique").ScaleWidth (Reseau_DCA_Largeur_Maxi / Ltotal), msoFalse, msoScaleFromTopLeft
    End If
    If Htotal < Reseau_DCA_Hauteur_Maxi Then
        CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Shapes("HENGraphique").ScaleHeight (Reseau_DCA_Hauteur_Maxi / Htotal), msoFalse, msoScaleFromTopLeft
    End If
    
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects("HENGraphique").Chart.Shapes("Group99").ScaleWidth 0.1, msoFalse, msoScaleFromTopLeft
        Ltotal = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects("HENGraphique").Chart.Shapes("Group99").Width
        CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects("HENGraphique").Chart.Shapes("Group99").ScaleWidth (Reseau_DCA_Largeur_Maxi / Ltotal), msoFalse, msoScaleFromTopLeft
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects("HENGraphique").Chart.Shapes("Group99").ScaleHeight 0.1, msoFalse, msoScaleFromTopLeft
        Htotal = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects("HENGraphique").Chart.Shapes("Group99").Height
        CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects("HENGraphique").Chart.Shapes("Group99").ScaleHeight (Reseau_DCA_Hauteur_Maxi / Htotal), msoFalse, msoScaleFromTopLeft
    

    'Finalisation
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("A1").Select
    Tabelle9.Activate
    Tabelle9.Shapes("Group99").Select
    Selection.Delete

    'Modification du zoom de la feuille pour que l'affichage soit OK
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Activate
    ZoomFeuille = ActiveWindow.Zoom
    ActiveWindow.Zoom = ZoomFeuille * 2
    ActiveWindow.Zoom = ZoomFeuille


End Sub

Private Sub ecrire_VeriPinch_position_I8(ByVal nbflux_fr As Integer, ByRef deltat_fr() As Double, ByVal hot_need As Double, ByVal nbflux_ch As Integer, ByRef deltat_ch() As Double, ByVal cold_need As Double, ByVal recup_max As Double)

Dim I As Integer
Dim j As Integer


'Sauvegarde du HEN
If PINCH_DANS_OCEAN = False Then
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 10, 2).Value = "HEN vide"
    For I = 0 To nbflux
        For j = 1 To data_flux_hen_size
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 11 + I, 2 + j).Value = Replace(data_flux_hen(I, j), separateur_decimal, ".")
        Next
    Next
End If

'MODIF CRE 10/10/2012   Ajout du code (duplication) pour affichage sur la page Tampon également
If PINCH_DANS_OCEAN = True Then
    'Besoins chauds
    If nbflux_fr > 0 Then
        CodeNames("Feuil63").Range("PINCH_besoins_chauds_totaux_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = Replace(deltat_fr(UBound(deltat_fr()), 2), separateur_decimal, ".")
    Else
        CodeNames("Feuil63").Range("PINCH_besoins_chauds_totaux_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = 0
    End If
    CodeNames("Feuil63").Range("PINCH_besoins_chauds_non_couverts_mini_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = Replace(hot_need, separateur_decimal, ".")
    'Besoins froids
    If nbflux_ch > 0 Then
        CodeNames("Feuil63").Range("PINCH_besoins_froids_totaux_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = Replace(deltat_ch(1, 2), separateur_decimal, ".")
    Else
        CodeNames("Feuil63").Range("PINCH_besoins_froids_totaux_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = 0
    End If
    CodeNames("Feuil63").Range("PINCH_besoins_froids_non_couverts_mini_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = Replace(cold_need, separateur_decimal, ".")
Else
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(5, 13).Value = "Point de Pinch (°C)"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(6, 13).Value = "Pincement effectif (K)"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(7, 13).Value = "Pincement souhaité (K)"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(8, 13).Value = "Besoins énergétiques"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(9, 13).Value = "Chauds (kW)"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(10, 13).Value = "Froids (kW)"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(11, 13).Value = "Récupération énergétique (kW)"
    If point_pinch <> -1000 Then
        If pinch_mesure <> 1000 Then    'MODIF CRE 08/09/2012   Correction de l'affichage quand il n'y a pas de récupération énergétique
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(5, 14).Value = Replace(point_pinch, separateur_decimal, ".")
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(6, 14).Value = Replace(pinch_mesure, separateur_decimal, ".")
        Else
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(5, 14).Value = ""
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(6, 14).Value = ""
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(5, 15).Value = "Les données de votre problème ne permettent pas de récupération énergétique avec la valeur de pincement fixé."
        End If
    Else
        If nbflux_ch = 0 Then
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(5, 15).Value = "Votre problème ne comporte aucun flux chaud, la récupération énergétique est nulle et le point de pinch inexistant."
        ElseIf nbflux_fr = 0 Then
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(5, 15).Value = "Votre problème ne comporte aucun flux froid, la récupération énergétique est nulle et le point de pinch inexistant."
        Else
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(5, 15).Value = "Message x01 - Contacter la Direction Technique"
        End If
    End If
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(7, 14).Value = Replace(pinch_target, separateur_decimal, ".")
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(8, 14).Value = "avant récupération"
    If nbflux_ch > 0 Then
        CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(10, 14).Value = Replace(deltat_ch(1, 2), separateur_decimal, ".")
    Else
        CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(10, 14).Value = ""
    End If
    If nbflux_fr > 0 Then
        CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(9, 14).Value = Replace(deltat_fr(UBound(deltat_fr()), 2), separateur_decimal, ".")
    Else
        CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(9, 14).Value = ""
    End If
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(8, 15).Value = "après récupération théorique"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(9, 15).Value = Replace(hot_need, separateur_decimal, ".")
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(10, 15).Value = Replace(cold_need, separateur_decimal, ".")
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(11, 15).Value = Replace(recup_max, separateur_decimal, ".")

    'MODIF CRE  16/10/2012  Affichage dans PINCH X
    
    If nbflux_fr > 0 Then
        CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(7, 4).Value = Replace(deltat_fr(UBound(deltat_fr()), 2), separateur_decimal, ".")
    Else
        CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(7, 4).Value = 0
    End If
    If nbflux_ch > 0 Then
        CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(19, 4).Value = Replace(deltat_ch(1, 2), separateur_decimal, ".")
    Else
        CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(19, 4).Value = 0
    End If
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(8, 4).Value = Replace(hot_need, separateur_decimal, ".")
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(20, 4).Value = Replace(cold_need, separateur_decimal, ".")
End If

End Sub

Private Sub ecrire_VeriPinch_position_II4()

Dim I As Integer
Dim j As Integer

If PINCH_DANS_OCEAN = False Then
    ligprint = ligprint + 10  'écriture des paramètres globaux
    ligprint = ligprint + nbflux + 4  'parce qu'on a déja écrit le HEN vierge
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint, 2).Value = "HEN avec HEN manuels"
    Call GetCols(data_flux_hen())
    For I = 0 To nbflux
        For j = 1 To data_flux_hen_size
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 1 + I, 2 + j).Value = Replace(data_flux_hen(I, j), separateur_decimal, ".")
        Next
    Next
End If

End Sub



Private Sub ecrire_VeriPinch_position_II6(ByVal nbflux_ch As Integer, ByVal nbflux_fr As Integer, ByVal restegch_ch As Double, ByVal restedte_ch As Double, ByVal restegch_fr As Double, ByVal restedte_fr As Double)

Dim I As Integer
Dim j As Integer


'Ecriture de la grille de flux avec le HEN sur Excel
If PINCH_DANS_OCEAN = True Then
    CodeNames("Feuil63").Range("PINCH_pincement_consigne_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = Replace(pinch_target, separateur_decimal, ".")
    If point_pinch <> -1000 Then
        If pinch_mesure <> 1000 Then    'MODIF CRE 08/09/2012   Correction de l'affichage quand il n'y a pas de récupération énergétique
            RecuperationEnergetiquePossible = True
            CodeNames("Feuil63").Range("PINCH_pincement_effectif_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = Replace(pinch_mesure, separateur_decimal, ".")
        Else
            RecuperationEnergetiquePossible = False
            CodeNames("Feuil63").Range("PINCH_pincement_effectif_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = "-"
            CodeNames("Feuil63").Range("PINCH_pincement_effectif_message_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Formula = "= msgInfo_pasRecup"
        End If
    Else
        RecuperationEnergetiquePossible = False
        If nbflux_ch = 0 Then
            CodeNames("Feuil63").Range("PINCH_pincement_effectif_message_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Formula = "= msgInfo_pasFluxCh"
        ElseIf nbflux_fr = 0 Then
            CodeNames("Feuil63").Range("PINCH_pincement_effectif_message_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Formula = "= msgInfo_pasFluxFr"
        Else
            CodeNames("Feuil63").Range("PINCH_pincement_effectif_message_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Formula = "Error (nbflux_ch ; nbflux_fr)"
        End If
    End If
    'CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(33, 4).Value = Replace(surftot, separateur_decimal, ".") 'MODIF CRE 06/11/12 cette valeur sera écrite plus tard (tout à la fin
    'CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(34, 4).Value = Replace(CoefU, separateur_decimal, ".")
Else
    ligprint = 4
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint, 2).Value = "Paramètres généraux"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 1, 3).Value = "nb flux"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 1, 4).Value = nbflux
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 2, 3).Value = "surf tot"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 2, 4).Value = Replace(surftot, separateur_decimal, ".")
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 2, 4).NumberFormat = "# ##0"" m²"""
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 3, 3).Value = "nb hex"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 3, 4).Value = nbhex
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 4, 4).Value = "gch"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 4, 5).Value = "dte"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 5, 3).Value = "besoin Ch"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 5, 4).Value = Replace(restegch_ch, separateur_decimal, ".")
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 5, 5).Value = Replace(restedte_ch, separateur_decimal, ".")
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 6, 3).Value = "besoin Fr"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 6, 4).Value = Replace(restegch_fr, separateur_decimal, ".")
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 6, 5).Value = Replace(restedte_fr, separateur_decimal, ".")
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 7, 3).Value = "data_flux_hen_size"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 7, 4).Value = data_flux_hen_size
    ligprint = ligprint + 10
    'MODIF CRE 16/10/2012   Affichage dans PINCH X
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(31, 4).Value = Replace(pinch_target, separateur_decimal, ".")
    If point_pinch <> -1000 Then
        If pinch_mesure <> 1000 Then    'MODIF CRE 08/09/2012   Correction de l'affichage quand il n'y a pas de récupération énergétique
            CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(32, 4).Value = Replace(pinch_mesure, separateur_decimal, ".")
        Else
            CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(32, 4).Value = "-"
            CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(32, 6).Value = Feuil5.Range("msgInfo_pasRecup").Value
        End If
    Else
        If nbflux_ch = 0 Then
            CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(32, 6).Value = Feuil5.Range("msgInfo_pasFluxCh").Value
        ElseIf nbflux_fr = 0 Then
            CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(32, 6).Value = Feuil5.Range("msgInfo_pasFluxFr").Value
        Else
            CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(32, 6).Value = "Error (nbflux_ch ; nbflux_fr)"
        End If
    End If
    'CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(33, 4).Value = Replace(surftot, separateur_decimal, ".") 'MODIF CRE 06/11/12 cette valeur sera écrite plus tard (tout à la fin
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(34, 4).Value = Replace(CoefU, separateur_decimal, ".")
End If



'décalage car données déja écrite auparavant (paramètres généraux du HEN)   'MODIF CRE 16/10/2012    on décale 2 fois parce qu'on a également écrit le HEN avec HEX manuels seulement
ligprint = ligprint + nbflux + 4
ligprint = ligprint + nbflux + 4

'Ecriture du HEN final
If PINCH_DANS_OCEAN = False Then
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint, 2).Value = "HEN sans utilités"
    Call GetCols(data_flux_hen())
    For I = 0 To nbflux
        For j = 1 To data_flux_hen_size
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 1 + I, 2 + j).Value = Replace(data_flux_hen(I, j), separateur_decimal, ".")
        Next
    Next
    ligprint = ligprint + nbflux + 4
End If



If PINCH_DANS_OCEAN = True Then
    CodeNames("Feuil63").Range("PINCH_besoins_chauds_non_couverts_apres_recuperation_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = Replace(besoins_HOT, separateur_decimal, ".")
    CodeNames("Feuil63").Range("PINCH_besoins_froids_non_couverts_apres_recuperation_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = Replace(besoins_COLD, separateur_decimal, ".")
Else
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(8, 16).Value = "après récupération réelle"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(9, 16).Value = Replace(besoins_HOT, separateur_decimal, ".")
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(10, 16).Value = Replace(besoins_COLD, separateur_decimal, ".")
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(11, 16).Value = CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(11, 15).Value - (besoins_HOT - CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(9, 16).Value)
    'MODIF CRE 16/10/2012   Affichage dans PINCH X
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(9, 4).Value = Replace(besoins_HOT, separateur_decimal, ".")
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(21, 4).Value = Replace(besoins_COLD, separateur_decimal, ".")
End If

End Sub


Private Sub ecrire_VeriPinch_position_IV3_1(ByVal jmax As Integer)

Dim I As Integer
Dim j As Integer

If PINCH_DANS_OCEAN = False Then
    'Affichage des grilles de flux Process
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint, 2).Value = "HEN avec utilités"
    For I = 0 To nbflux
        For j = 1 To UBound(data_flux_hen(), 2)
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 1 + I, 2 + j).Value = Replace(data_flux_hen(I, j), separateur_decimal, ".")
        Next
    Next
    ligprint = ligprint + I
    If nbutilities_updated > 0 Then
        'Affichage des grilles de flux Utilités (updated seulement)
        For I = 1 To nbutilities_updated
            For j = 1 To UBound(data_flux_hen(), 2)
                CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + I, 2 + j).Value = Replace(data_flux_utilities_updated(I, j), separateur_decimal, ".")
            Next
        Next
        ligprint = ligprint + I
    Else
        ligprint = ligprint + 1
    End If
    ligprint = ligprint + 2

    'Affichage de la grille globale après affinage
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint, 2).Value = "HEN après affinage"
    For I = 0 To nbflux_final
        For j = 1 To jmax
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 1 + I, 2 + j).Value = Replace(data_flux_all_final(I, j), separateur_decimal, ".")
        Next
    Next
    ligprint = ligprint + I
    ligprint = ligprint + 3

    'Ecriture des caractéristiques des HEX, pour rappel
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint, 2).Value = "Caractéristiques des HEX"
    For I = 1 To 12
        For j = 1 To data_hen_size
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + I, 2 + j).Value = Replace(data_hen(I, j), separateur_decimal, ".")
        Next
    Next
    ligprint = ligprint + 15

    If nbutilities_updated > 0 Then
        'Affichage des informations globales sur les utilités
        CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint, 2).Value = "Récap des informations des Utilités"
        CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 1, 3).Value = "Nom"
        CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 2, 3).Value = "tipe"
        CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 3, 3).Value = "tin"
        CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 4, 3).Value = "mcp"
        CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 5, 3).Value = "p totale"
        CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 6, 3).Value = "tout"
        CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + 7, 3).Value = "p restante"
        For I = 1 To 7
            For j = 1 To nbutilities
                CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(ligprint + I, 3 + j).Value = Replace(data_utilities(I, j), separateur_decimal, ".")
            Next
        Next
        ligprint = ligprint + 9
    End If
End If

End Sub

Private Sub ecrire_VeriPinch_position_IV3_2(ByVal besoins_hot_final As Double, ByVal besoins_cold_final As Double)

If PINCH_DANS_OCEAN = True Then
    CodeNames("Feuil63").Range("PINCH_besoins_chauds_non_couverts_apres_utilites_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = besoins_hot_final
    CodeNames("Feuil63").Range("PINCH_besoins_froids_non_couverts_apres_utilites_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = besoins_cold_final
Else
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(8, 17).Value = "restant à couvrir après recours aux utilités"
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(9, 17).Value = besoins_hot_final
    CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(10, 17).Value = besoins_cold_final
    'MODIF CRE 10/10/2012   Affichage dans PINCH X
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(10, 4).Value = besoins_hot_final
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(22, 4).Value = besoins_cold_final
End If

End Sub


Private Sub ecriture_formule_pour_changement_langue_flux_et_utilites_dans_plage_cellules(ByVal NomPlage As String)
'DCA 22/05/2014
'Recherche des noms des plages de cellules contenant les noms des flux procédés et utilités

'Déclarations
Dim Colonne As Integer
Dim LigneMatrice As Integer
Dim LigneNomsPlagesFlux As Integer
Dim ContenuCellule As String
Dim NomPlageFlux As String
Dim NomFluxTrouve As Boolean

'Initialisations
Colonne = 0
LigneMatrice = 0
LigneNomsPlagesFlux = 0
ContenuCellule = ""
NomPlageFlux = ""
NomFluxTrouve = False

'On se positionne en fonction de la filière calculée au niveau de la première ligne3.
Colonne = 1
LigneMatrice = 1
Do While Feuil63.Range(NomPlage).Cells(LigneMatrice, Colonne) <> ""
    Do While Feuil63.Range(NomPlage).Cells(LigneMatrice, Colonne) <> ""
        ContenuCellule = Feuil63.Range(NomPlage).Cells(LigneMatrice, Colonne)
        LigneNomsPlagesFlux = 1
        NomFluxTrouve = False
        Do While Feuil63.Range("PINCH_NomsPlagesFluxProcede").Cells(LigneNomsPlagesFlux, 1) <> "" And NomFluxTrouve = False
            NomPlageFlux = Feuil63.Range("PINCH_NomsPlagesFluxProcede").Cells(LigneNomsPlagesFlux, 1)
            If ContenuCellule = Feuil5.Range(NomPlageFlux).Cells(1, 1) Then
                Feuil63.Range(NomPlage).Cells(LigneMatrice, Colonne).Formula = "= " & NomPlageFlux
                NomFluxTrouve = True
            End If
            LigneNomsPlagesFlux = LigneNomsPlagesFlux + 1
        Loop
        LigneNomsPlagesFlux = 1
        Do While Feuil63.Range("PINCH_NomsPlagesFluxUtilites").Cells(LigneNomsPlagesFlux, 1) <> "" And NomFluxTrouve = False
            NomPlageFlux = Feuil63.Range("PINCH_NomsPlagesFluxUtilites").Cells(LigneNomsPlagesFlux, 1)
            If ContenuCellule = Feuil41.Range(NomPlageFlux).Cells(1, 1) Then
                Feuil63.Range(NomPlage).Cells(LigneMatrice, Colonne).Formula = "= " & NomPlageFlux
                NomFluxTrouve = True
            End If
            LigneNomsPlagesFlux = LigneNomsPlagesFlux + 1
        Loop
        LigneMatrice = LigneMatrice + 1
    Loop
    LigneMatrice = 1
    Colonne = Colonne + 1
Loop


End Sub

Private Sub effacer_VeriPinch_position_III2(ByVal state As Integer)

If PINCH_DANS_OCEAN = False Then
    If PINCH_DANS_OCEAN = False Then
        'Suppression des Utilités précédemment ajoutées
        If state = 3 Then
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Range("B" & nbflux * 3 + 26 & ":IV100").ClearContents
        End If
    End If
End If

End Sub

Private Sub lecture_VeriPinch_position_III1(ByVal state As Integer)

Dim I As Integer
Dim j As Integer

If PINCH_DANS_OCEAN = False Then
    'Lecture des données du HEN vierge depuis Excel si on n'a lancé le code qu'à partir de l'ajout des Utilités
    If state = 3 Then
        nbflux = CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(5, 4).Value
        nbhex = CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(7, 4).Value
        data_flux_hen_size = CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(11, 4)
        ReDim data_flux_hen(0 To nbflux, 1 To data_flux_hen_size)
        For I = 0 To nbflux
            For j = 1 To data_flux_hen_size
                data_flux_hen(I, j) = CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(I + 23 + nbflux * 2, j + 2).Value
            Next
        Next
        ligprint = 26 + nbflux * 3
        'ligprint = nbflux * 2 + 35    'mise à jour des ligprint ssi on a changé de state
    End If
End If


End Sub



Sub representation_echangeur_DCA(ByVal i_flux As Integer, ByRef flux_chaud() As Boolean, ByVal i_echangeur As Integer, ByVal numero_echangeur As Integer, ByVal i_flux_2 As Integer, ByVal i_echangeur_2 As Integer, ByRef Matrice() As Integer, ByRef echangeur_represente() As Boolean)

    Dim ecriture_temperature As Boolean

    'On génère la textbox de l'échangeur
    Tabelle9.Shapes.AddTextbox(msoTextOrientationHorizontal, Left_debut_reseau + (i_echangeur - 1) * (Ecart_colonne_reseau + Largeur_colonne), Top_debut_reseau - Ecart_ligne_reseau - Hauteur_ligne, Largeur_colonne, Hauteur_ligne).Select
    With Selection
        .Characters.Text = Feuil63.Range("PINCH_intitule_echangeur").Cells(1, 1) & Chr(10) & "n° " & numero_echangeur '& vbCrLf & Round(echangeurs(repere_echangeur_surface, numero_echangeur), 1) & " m²"
        .AutoSize = False
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Name = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceEchangeur_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.Name   '"Arial"
        .Font.FontStyle = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceEchangeur_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.FontStyle
        .Font.size = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceEchangeur_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.size
        .Font.ColorIndex = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceEchangeur_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.ColorIndex
        With .ShapeRange.fill
            .Solid
            If echangeurs(repere_echangeur_force, numero_echangeur) = True Then
                .ForeColor.RGB = RGB(204, 255, 255)
            Else
                .ForeColor.RGB = RGB(255, 255, 0)
            End If
            .Transparency = 0
        End With
        With .ShapeRange.Line
            .BackColor.RGB = RGB(255, 255, 255)
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = 1
            .DashStyle = msoLineSolid
            .Style = msoLineSingle
            .Transparency = 0
            .Visible = msoTrue
            .Visible = msoTrue
        End With
        'A compléter avec les trucs de Clément
    End With
    If i_echangeur_2 <> i_echangeur Then   'cas de croisements de flux
        Tabelle9.Shapes.AddTextbox(msoTextOrientationHorizontal, Left_debut_reseau + (i_echangeur_2 - 1) * (Ecart_colonne_reseau + Largeur_colonne), Top_debut_reseau - Ecart_ligne_reseau - Hauteur_ligne, Largeur_colonne, Hauteur_ligne).Select
        With Selection
            .Characters.Text = Feuil63.Range("PINCH_intitule_echangeur").Cells(1, 1) & " " & numero_echangeur '& vbCrLf & Round(echangeurs(repere_echangeur_surface, numero_echangeur), 1) & " m²"
            .AutoSize = False
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
            .Font.Name = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceEchangeur_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.Name   '"Arial"
            .Font.FontStyle = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceEchangeur_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.FontStyle
            .Font.size = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceEchangeur_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.size
            .Font.ColorIndex = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceEchangeur_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.ColorIndex
            With .ShapeRange.fill
                .Solid
                .ForeColor.SchemeColor = 9
                .Transparency = 0
            End With
            With .ShapeRange.Line
                .Weight = 1
                .DashStyle = msoLineSolid
                .Style = msoLineSingle
                .Transparency = 0
                .Visible = msoTrue
                .ForeColor.SchemeColor = 0
                .BackColor.RGB = RGB(255, 255, 255)
                .Visible = msoTrue
            End With
            'A compléter avec les trucs de Clément
        End With
    End If

    'On génère l'ellipse du flux 1
    Tabelle9.Shapes.AddShape(msoShapeOval, Left_debut_reseau + (i_echangeur - 1) * (Ecart_colonne_reseau + Largeur_colonne), Top_debut_reseau + (i_flux - 1) * (Ecart_ligne_reseau + Hauteur_ligne), Largeur_colonne, Hauteur_ligne).Select
    With Selection
        .Characters.Text = Round(echangeurs(repere_echangeur_puissance, numero_echangeur), 1) & " kW"
        .AutoSize = False
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Name = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePuissance_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.Name
        .Font.FontStyle = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePuissance_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.FontStyle
        .Font.size = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePuissance_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.size
        .Font.ColorIndex = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePuissance_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.ColorIndex
        With .ShapeRange.fill
            .Solid
            .ForeColor.SchemeColor = 9
            .Transparency = 0
        End With
        With .ShapeRange.Line
            .Weight = 1
            .DashStyle = msoLineSolid
            .Style = msoLineSingle
            .Transparency = 0
            .Visible = msoTrue
            .ForeColor.SchemeColor = 0
            .BackColor.RGB = RGB(255, 255, 255)
            .Visible = msoTrue
        End With
        'A compléter avec les trucs de Clément
    End With
    'affichage température gauche
    If i_echangeur > 1 Then
        If Matrice(i_flux, i_echangeur - 1) > 0 Then
            ecriture_temperature = False
        Else
            ecriture_temperature = True
        End If
    Else
        ecriture_temperature = True
    End If
    If ecriture_temperature = True Then
        Tabelle9.Shapes.AddTextbox(msoTextOrientationHorizontal, Left_debut_reseau + (i_echangeur - 1) * (Ecart_colonne_reseau + Largeur_colonne) - Ecart_colonne_reseau - Largeur_colonne / 2, Top_debut_reseau + (i_flux - 1) * (Ecart_ligne_reseau + Hauteur_ligne) - Ecart_ligne_reseau, Largeur_colonne + Ecart_colonne_reseau, Hauteur_ligne / 2 + Ecart_ligne_reseau).Select
        With Selection
            If flux_chaud(i_flux) = True Then
                .Characters.Text = Round(echangeurs(repere_echangeur_CH_Tin, numero_echangeur), 1) & " °C"
            Else
                .Characters.Text = Round(echangeurs(repere_echangeur_FR_Tout, numero_echangeur), 1) & " °C"
            End If
            .AutoSize = False
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Name = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.Name            '"Arial"
            .Font.FontStyle = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.FontStyle
            .Font.size = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.size
            .Font.ColorIndex = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.ColorIndex
            With .ShapeRange.fill
                .Solid
                .ForeColor.SchemeColor = 9
                .Transparency = 1
            End With
            With .ShapeRange.Line
                '                    .Weight = 1
                '                    .DashStyle = msoLineSolid
                '                    .Style = msoLineSingle
                '                    .Transparency = 0
                .Visible = msoFalse
                '                    .ForeColor.SchemeColor = 0
                '                    .BackColor.RGB = RGB(255, 255, 255)
                '                    .Visible = msoTrue
            End With
            'A compléter avec les trucs de Clément
        End With
    End If
    'affichage température droite
    If i_echangeur < UBound(Matrice, 2) Then
        If Matrice(i_flux, i_echangeur + 1) > 0 Then
            If echangeur_represente(Matrice(i_flux, i_echangeur + 1)) = True Then
                ecriture_temperature = False
            Else
                ecriture_temperature = True
            End If
        Else
            ecriture_temperature = True
        End If
    Else
        ecriture_temperature = True
    End If
    If ecriture_temperature = True Then
        Tabelle9.Shapes.AddTextbox(msoTextOrientationHorizontal, Left_debut_reseau + i_echangeur * (Ecart_colonne_reseau + Largeur_colonne) - Ecart_colonne_reseau - Largeur_colonne / 2, Top_debut_reseau + (i_flux - 1) * (Ecart_ligne_reseau + Hauteur_ligne) - Ecart_ligne_reseau, Largeur_colonne + Ecart_colonne_reseau, Hauteur_ligne / 2 + Ecart_ligne_reseau).Select
        With Selection
            If flux_chaud(i_flux) = True Then
                .Characters.Text = Round(echangeurs(repere_echangeur_CH_Tout, numero_echangeur), 1) & " °C"
            Else
                .Characters.Text = Round(echangeurs(repere_echangeur_FR_Tin, numero_echangeur), 1) & " °C"
            End If
            .AutoSize = False
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Name = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.Name            '"Arial"
            .Font.FontStyle = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.FontStyle
            .Font.size = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.size
            .Font.ColorIndex = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.ColorIndex
            With .ShapeRange.fill
                .Solid
                .ForeColor.SchemeColor = 9
                .Transparency = 1
            End With
            With .ShapeRange.Line
                '                    .Weight = 1
                '                    .DashStyle = msoLineSolid
                '                    .Style = msoLineSingle
                '                    .Transparency = 0
                .Visible = msoFalse
                '                    .ForeColor.SchemeColor = 0
                '                    .BackColor.RGB = RGB(255, 255, 255)
                '                    .Visible = msoTrue
            End With
            'A compléter avec les trucs de Clément
        End With
    End If

    'On génère l'ellipse du flux 2
    Tabelle9.Shapes.AddShape(msoShapeOval, Left_debut_reseau + (i_echangeur_2 - 1) * (Ecart_colonne_reseau + Largeur_colonne), Top_debut_reseau + (i_flux_2 - 1) * (Ecart_ligne_reseau + Hauteur_ligne), Largeur_colonne, Hauteur_ligne).Select
    With Selection
        .Characters.Text = Round(echangeurs(repere_echangeur_puissance, numero_echangeur), 1) & " kW"
        .AutoSize = False
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Name = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePuissance_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.Name
        .Font.FontStyle = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePuissance_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.FontStyle
        .Font.size = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePuissance_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.size
        .Font.ColorIndex = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePuissance_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.ColorIndex
        With .ShapeRange.fill
            .Solid
            .ForeColor.SchemeColor = 9
            .Transparency = 0
        End With
        With .ShapeRange.Line
            .Weight = 1
            .DashStyle = msoLineSolid
            .Style = msoLineSingle
            .Transparency = 0
            .Visible = msoTrue
            .ForeColor.SchemeColor = 0
            .BackColor.RGB = RGB(255, 255, 255)
            .Visible = msoTrue
        End With
        'A compléter avec les trucs de Clément
    End With
    'affichage température gauche
    If i_echangeur_2 > 1 Then
        If Matrice(i_flux_2, i_echangeur_2 - 1) > 0 Then
            If echangeur_represente(Matrice(i_flux_2, i_echangeur_2 - 1)) = True Then
                ecriture_temperature = False
            Else
                ecriture_temperature = True
            End If
        Else
            ecriture_temperature = True
        End If
    Else
        ecriture_temperature = True
    End If
    If ecriture_temperature = True Then
        Tabelle9.Shapes.AddTextbox(msoTextOrientationHorizontal, Left_debut_reseau + (i_echangeur_2 - 1) * (Ecart_colonne_reseau + Largeur_colonne) - Ecart_colonne_reseau - Largeur_colonne / 2, Top_debut_reseau + (i_flux_2 - 1) * (Ecart_ligne_reseau + Hauteur_ligne) - Ecart_ligne_reseau, Largeur_colonne + Ecart_colonne_reseau, Hauteur_ligne / 2 + Ecart_ligne_reseau).Select
        With Selection
            If flux_chaud(i_flux_2) = True Then
                .Characters.Text = Round(echangeurs(repere_echangeur_CH_Tin, numero_echangeur), 1) & " °C"
            Else
                .Characters.Text = Round(echangeurs(repere_echangeur_FR_Tout, numero_echangeur), 1) & " °C"
            End If
            .AutoSize = False
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Name = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.Name            '"Arial"
            .Font.FontStyle = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.FontStyle
            .Font.size = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.size
            .Font.ColorIndex = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.ColorIndex
            With .ShapeRange.fill
                .Solid
                .ForeColor.SchemeColor = 9
                .Transparency = 1
            End With
            With .ShapeRange.Line
                '                    .Weight = 1
                '                    .DashStyle = msoLineSolid
                '                    .Style = msoLineSingle
                '                    .Transparency = 0
                .Visible = msoFalse
                '                    .ForeColor.SchemeColor = 0
                '                    .BackColor.RGB = RGB(255, 255, 255)
                '                    .Visible = msoTrue
            End With
            'A compléter avec les trucs de Clément
        End With
    End If
    'affichage température droite
    If i_echangeur_2 < UBound(Matrice, 2) Then
        If Matrice(i_flux_2, i_echangeur_2 + 1) > 0 Then
            If echangeur_represente(Matrice(i_flux_2, i_echangeur_2 + 1)) = True Then
                ecriture_temperature = False
            Else
                ecriture_temperature = True
            End If
        Else
            ecriture_temperature = True
        End If
    Else
        ecriture_temperature = True
    End If
    If ecriture_temperature = True Then
        Tabelle9.Shapes.AddTextbox(msoTextOrientationHorizontal, Left_debut_reseau + i_echangeur_2 * (Ecart_colonne_reseau + Largeur_colonne) - Ecart_colonne_reseau - Largeur_colonne / 2, Top_debut_reseau + (i_flux_2 - 1) * (Ecart_ligne_reseau + Hauteur_ligne) - Ecart_ligne_reseau, Largeur_colonne + Ecart_colonne_reseau, Hauteur_ligne / 2 + Ecart_ligne_reseau).Select
        With Selection
            If flux_chaud(i_flux_2) = True Then
                .Characters.Text = Round(echangeurs(repere_echangeur_CH_Tout, numero_echangeur), 1) & " °C"
            Else
                .Characters.Text = Round(echangeurs(repere_echangeur_FR_Tin, numero_echangeur), 1) & " °C"
            End If
            .AutoSize = False
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Name = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.Name            '"Arial"
            .Font.FontStyle = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.FontStyle
            .Font.size = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.size
            .Font.ColorIndex = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.ColorIndex
            With .ShapeRange.fill
                .Solid
                .ForeColor.SchemeColor = 9
                .Transparency = 1
            End With
            With .ShapeRange.Line
                '                    .Weight = 1
                '                    .DashStyle = msoLineSolid
                '                    .Style = msoLineSingle
                '                    .Transparency = 0
                .Visible = msoFalse
                '                    .ForeColor.SchemeColor = 0
                '                    .BackColor.RGB = RGB(255, 255, 255)
                '                    .Visible = msoTrue
            End With
            'A compléter avec les trucs de Clément
        End With
    End If

    'On relie les deux ellipses
    Tabelle9.Shapes.AddLine(Left_debut_reseau + (i_echangeur - 1) * (Ecart_colonne_reseau + Largeur_colonne) + Largeur_colonne / 2, Top_debut_reseau + (i_flux - 1) * (Ecart_ligne_reseau + Hauteur_ligne) + Hauteur_ligne, Left_debut_reseau + (i_echangeur_2 - 1) * (Ecart_colonne_reseau + Largeur_colonne) + Largeur_colonne / 2, Top_debut_reseau + (i_flux_2 - 1) * (Ecart_ligne_reseau + Hauteur_ligne)).Select
    Selection.ShapeRange.Line.ForeColor.RGB = RGB(0, 0, 0)


End Sub

Sub representation_puissance_non_echangee_DCA(ByVal i_flux As Integer, ByRef flux_chaud() As Boolean, ByVal i_echangeur As Integer, ByVal numero_puissance_non_echangee As Integer, ByRef Matrice() As Integer, ByRef echangeur_represente() As Boolean)

    Dim ecriture_temperature As Boolean

    'On génère la textbox de l'échangeur
    '    Tabelle9.Shapes.AddTextbox(msoTextOrientationHorizontal, Left_debut_reseau + (i_echangeur - 1) * (Ecart_colonne_reseau + Largeur_colonne), Top_debut_reseau - Ecart_ligne_reseau - Hauteur_ligne, Largeur_colonne, Hauteur_ligne).Select
    '    With Selection
    '        .Characters.Text = "# " & numero_echangeur & vbCrLf & Round(echangeurs(repere_echangeur_surface, numero_echangeur), 1) & " m²"
    '        .AutoSize = False
    '        .VerticalAlignment = xlCenter
    '        .HorizontalAlignment = xlCenter
    '        .Font.Name = "Arial"
    '        .Font.FontStyle = "Normal"
    '        .Font.size =Reseau_DCA_Taille_Police
    '        With .ShapeRange.fill
    '            .Solid
    '            .ForeColor.SchemeColor = 9
    '            .Transparency = 0
    '        End With
    '        With .ShapeRange.Line
    '            .Weight = 1
    '            .DashStyle = msoLineSolid
    '            .Style = msoLineSingle
    '            .Transparency = 0
    '            .Visible = msoTrue
    '            .ForeColor.SchemeColor = 0
    '            .BackColor.RGB = RGB(255, 255, 255)
    '            .Visible = msoTrue
    '        End With
    '        'A compléter avec les trucs de Clément
    '    End With

    'On génère l'ellipse du flux 1
    Tabelle9.Shapes.AddShape(msoShapeOval, Left_debut_reseau + (i_echangeur - 1) * (Ecart_colonne_reseau + Largeur_colonne), Top_debut_reseau + (i_flux - 1) * (Ecart_ligne_reseau + Hauteur_ligne), Largeur_colonne, Hauteur_ligne).Select
    With Selection
        .Characters.Text = Round(puissances_non_echangees(P_non_echangee_puisssance, numero_puissance_non_echangee), 1) & " kW"
        .AutoSize = False
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Name = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePuissance_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.Name
        .Font.FontStyle = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePuissance_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.FontStyle
        .Font.size = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePuissance_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.size
        .Font.ColorIndex = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PolicePuissance_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.ColorIndex
        With .ShapeRange.fill
            .Solid
            .ForeColor.SchemeColor = 9
            .Transparency = 0
        End With
        With .ShapeRange.Line
            .Weight = 3
            .DashStyle = msoLineSolid
            .Style = msoLineSingle
            .Transparency = 0
            .Visible = msoTrue
            .BackColor.RGB = RGB(255, 255, 255)
            If i_flux <= UBound(flux_thermique_process, 2) Then
                If flux_chaud(i_flux) = False Then
                    .ForeColor.RGB = RGB(255, 0, 0)
                Else
                    .ForeColor.RGB = RGB(0, 0, 255)
                End If
            Else
                .ForeColor.RGB = RGB(192, 192, 192)
            End If
            .Visible = msoTrue
        End With
        'A compléter avec les trucs de Clément
    End With
    'affichage température gauche
    If i_echangeur > 1 Then
        If Matrice(i_flux, i_echangeur - 1) > 0 Then
            ecriture_temperature = False
        Else
            ecriture_temperature = True
        End If
    Else
        ecriture_temperature = True
    End If
    If ecriture_temperature = True Then
        Tabelle9.Shapes.AddTextbox(msoTextOrientationHorizontal, Left_debut_reseau + (i_echangeur - 1) * (Ecart_colonne_reseau + Largeur_colonne) - Ecart_colonne_reseau - Largeur_colonne / 2, Top_debut_reseau + (i_flux - 1) * (Ecart_ligne_reseau + Hauteur_ligne) - Ecart_ligne_reseau, Largeur_colonne + Ecart_colonne_reseau, Hauteur_ligne / 2 + Ecart_ligne_reseau).Select
        With Selection
            .Characters.Text = Round(puissances_non_echangees(P_non_echangee_T_chaude, numero_puissance_non_echangee), 1) & " °C"
            .AutoSize = False
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Name = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.Name            '"Arial"
            .Font.FontStyle = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.FontStyle
            .Font.size = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.size
            .Font.ColorIndex = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.ColorIndex
            With .ShapeRange.fill
                .Solid
                .ForeColor.SchemeColor = 9
                .Transparency = 1
            End With
            With .ShapeRange.Line
                '                    .Weight = 1
                '                    .DashStyle = msoLineSolid
                '                    .Style = msoLineSingle
                '                    .Transparency = 0
                .Visible = msoFalse
                '                    .ForeColor.SchemeColor = 0
                '                    .BackColor.RGB = RGB(255, 255, 255)
                '                    .Visible = msoTrue
            End With
            'A compléter avec les trucs de Clément
        End With
    End If
    'affichage température droite
    If i_echangeur < UBound(Matrice, 2) Then
        If Matrice(i_flux, i_echangeur + 1) > 0 Then
            If echangeur_represente(Matrice(i_flux, i_echangeur + 1)) = True Then
                ecriture_temperature = False
            Else
                ecriture_temperature = True
            End If
        Else
            ecriture_temperature = True
        End If
    Else
        ecriture_temperature = True
    End If
    If ecriture_temperature = True Then
        Tabelle9.Shapes.AddTextbox(msoTextOrientationHorizontal, Left_debut_reseau + i_echangeur * (Ecart_colonne_reseau + Largeur_colonne) - Ecart_colonne_reseau - Largeur_colonne / 2, Top_debut_reseau + (i_flux - 1) * (Ecart_ligne_reseau + Hauteur_ligne) - Ecart_ligne_reseau, Largeur_colonne + Ecart_colonne_reseau, Hauteur_ligne / 2 + Ecart_ligne_reseau).Select
        With Selection
            .Characters.Text = Round(puissances_non_echangees(P_non_echangee_T_froide, numero_puissance_non_echangee), 1) & " °C"
            .AutoSize = False
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Name = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.Name            '"Arial"
            .Font.FontStyle = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.FontStyle
            .Font.size = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.size
            .Font.ColorIndex = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceTemperature_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.ColorIndex
            With .ShapeRange.fill
                .Solid
                .ForeColor.SchemeColor = 9
                .Transparency = 1
            End With
            With .ShapeRange.Line
                '                    .Weight = 1
                '                    .DashStyle = msoLineSolid
                '                    .Style = msoLineSingle
                '                    .Transparency = 0
                .Visible = msoFalse
                '                    .ForeColor.SchemeColor = 0
                '                    .BackColor.RGB = RGB(255, 255, 255)
                '                    .Visible = msoTrue
            End With
            'A compléter avec les trucs de Clément
        End With
    End If




End Sub

Sub representation_flux_DCA(ByVal i_flux As Integer, ByVal nb_colonnes As Integer, ByRef flux_chaud() As Boolean)

    Dim i_utilites As Integer
    Dim nom_flux As String
    Dim nb_flux_process As Double

    nb_flux_process = UBound(flux_thermique_process, 2)
    If i_flux <= nb_flux_process Then
        'flux procédés
        nom_flux = flux_thermique_process(repere_flux_thermique_process_nom, i_flux)
    Else
        'flux utilités
        i_utilites = i_flux - nb_flux_process
        nom_flux = utilites_flux_energie(repere_flux_thermique_utilites_nom, i_utilites)
    End If

    'On passe à la représentation
    'étiquette du flux
    If flux_chaud(i_flux) = True Then
        Tabelle9.Shapes.AddTextbox(msoTextOrientationHorizontal, Left_debut_reseau - Ecart_colonne_reseau - Largeur_espace_supplementaire_etiquette_flux - Largeur_etiquette_flux, Top_debut_reseau + (i_flux - 1) * (Ecart_ligne_reseau + Hauteur_ligne), Largeur_etiquette_flux, Hauteur_ligne).Select
    Else
        Tabelle9.Shapes.AddTextbox(msoTextOrientationHorizontal, Left_debut_reseau + nb_colonnes * (Ecart_colonne_reseau + Largeur_colonne) + Largeur_espace_supplementaire_etiquette_flux, Top_debut_reseau + (i_flux - 1) * (Ecart_ligne_reseau + Hauteur_ligne), Largeur_etiquette_flux, Hauteur_ligne).Select
    End If
    With Selection
        .Characters.Text = nom_flux
        .AutoSize = False
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceFlux_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.Name            '"Arial"
        .Font.FontStyle = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceFlux_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.FontStyle
        .Font.size = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceFlux_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.size
        If flux_chaud(i_flux) = True Then
            .Font.color = RGB(255, 0, 0)
        Else
            .Font.color = RGB(0, 0, 255)
        End If
'        .Font.ColorIndex = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("ReseauDCA_PoliceFlux_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Font.ColorIndex
        With .ShapeRange.fill
            .Solid
            .ForeColor.SchemeColor = 9
            .Transparency = 0
        End With
        With .ShapeRange.Line
            .Weight = 1
            .DashStyle = msoLineSolid
            .Style = msoLineSingle
            .Transparency = 0
            .Visible = msoTrue
'            If flux_chaud(i_flux) = True Then
'                .ForeColor.SchemeColor = 10
'            Else
'                .ForeColor.SchemeColor = 12
'            End If
            .ForeColor.RGB = RGB(255, 255, 255)
            .BackColor.RGB = RGB(255, 255, 255)
            .Visible = msoTrue
        End With
        'A compléter avec les trucs de Clément
    End With
    'flèche
    If flux_chaud(i_flux) = True Then
        ActiveSheet.Shapes.AddLine(Left_debut_reseau - Ecart_colonne_reseau - Largeur_espace_supplementaire_etiquette_flux, Top_debut_reseau + (i_flux - 1) * (Ecart_ligne_reseau + Hauteur_ligne) + Hauteur_ligne / 2, Left_debut_reseau + nb_colonnes * (Ecart_colonne_reseau + Largeur_colonne) + Largeur_espace_supplementaire_etiquette_flux + Largeur_etiquette_flux, Top_debut_reseau + (i_flux - 1) * (Ecart_ligne_reseau + Hauteur_ligne) + Hauteur_ligne / 2).Select
        Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
        Selection.ShapeRange.Line.ForeColor.SchemeColor = 10
    Else
        ActiveSheet.Shapes.AddLine(Left_debut_reseau - Ecart_colonne_reseau - Largeur_espace_supplementaire_etiquette_flux - Largeur_etiquette_flux, Top_debut_reseau + (i_flux - 1) * (Ecart_ligne_reseau + Hauteur_ligne) + Hauteur_ligne / 2, Left_debut_reseau + nb_colonnes * (Ecart_colonne_reseau + Largeur_colonne) + Largeur_espace_supplementaire_etiquette_flux, Top_debut_reseau + (i_flux - 1) * (Ecart_ligne_reseau + Hauteur_ligne) + Hauteur_ligne / 2).Select
        Selection.ShapeRange.Line.BeginArrowheadStyle = msoArrowheadTriangle
        Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
    End If
    Selection.ShapeRange.Line.Weight = 1.5
    If i_flux <= nb_flux_process Then
        'flux procédés
        'Selection.ShapeRange.Line.DashStyle = msoLine
    Else
        'flux utilités
        Selection.ShapeRange.Line.DashStyle = msoLineLongDashDot
    End If
    

End Sub

Public Sub representation_reseau_DCA_lecture_donnees()
'Ajout DCA 24/05/2014 pour traduction du réseau d'échangeur en cas de changement de langue

Dim Ligne As Long
Dim LigneTab As Integer
Dim Colonne As Integer
Dim forme As Shape

'Changement de langue en cours
mn_ChangementLangue = True

'Lecture de data_hen_size et data_hen()
Ligne = 1
Colonne = 1
Do While Feuil63.Range("PINCH_RepresentationDCAMatrice_" & marqueur_etape & pi_FiliereConsideree).Cells(Ligne, Colonne) <> ""
    Colonne = Colonne + 1
Loop
data_hen_size = Colonne - 1
If data_hen_size > 0 Then
    ReDim data_hen(1 To 12, 1 To data_hen_size)
    For Ligne = 1 To 12
        For Colonne = 1 To data_hen_size
            data_hen(Ligne, Colonne) = Feuil63.Range("PINCH_RepresentationDCAMatrice_" & marqueur_etape & pi_FiliereConsideree).Cells(Ligne, Colonne)
            'data_hen(Ligne, Colonne) = Right(data_hen(Ligne, Colonne), Len(data_hen(Ligne, Colonne)) - 1)
        Next Colonne
    Next Ligne
Else
    pn_EffacementResultatsEnCours = True
    Exit Sub
End If

'Lecture de puissances_non_echangees() et nb_puissance_restants
Ligne = 1
Colonne = 1
Do While Feuil63.Range("PINCH_RepresentationDCAVecteurFluxChauds_" & marqueur_etape & pi_FiliereConsideree).Cells(Ligne, Colonne) <> ""
    Colonne = Colonne + 1
Loop
nb_puissance_restants = Colonne - 1
If nb_puissance_restants > 0 Then
    ReDim puissances_non_echangees(1 To P_non_echangee_nb_caracteristiques, 1 To nb_puissance_restants)
    For Ligne = 1 To P_non_echangee_nb_caracteristiques
        For Colonne = 1 To nb_puissance_restants
            puissances_non_echangees(Ligne, Colonne) = Feuil63.Range("PINCH_RepresentationDCAVecteurFluxChauds_" & marqueur_etape & pi_FiliereConsideree).Cells(Ligne, Colonne)
        Next Colonne
    Next Ligne
Else
    ReDim puissances_non_echangees(1 To P_non_echangee_nb_caracteristiques, 1 To 1)
End If

'Nombre d'échangeurs imposés
colonneStation = 1 + (pi_FiliereConsideree - 1) * 9
LigneTab = 4
nbheximposed = 0
Ligne = 0
Do While Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab + Ligne, colonneStation) <> ""
    If Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab + Ligne, colonneStation) = True Then
        nbheximposed = nbheximposed + 1
    End If
    Ligne = Ligne + 1
Loop

If Feuil63.Range("PINCH_pincement_consigne_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = "" Then
    RecuperationEnergetiquePossible = False
    pinch_mesure = 0
    point_pinch = 0
    T_pinch_reseau_DCA = 0
Else
    RecuperationEnergetiquePossible = True
    pinch_mesure = CodeNames("Feuil63").Range("PINCH_pincement_effectif_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1)
    point_pinch = CodeNames("Feuil63").Range("PINCH_pincement_point_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1)
    T_pinch_reseau_DCA = point_pinch + 0.01
End If

'Lecture des autres données
Call lecture_flux_process
Call lecture_flux_utilites

'On efface les résultats
For Each forme In CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects("HENGraphique").Chart.Shapes
    forme.Delete
Next forme


End Sub

Private Sub representationDCA_recuperation_puissances_non_echangees()

    Dim I As Integer
    Dim j As Integer
    
    nb_puissance_restants = 0
    For I = 1 To UBound(data_flux_all_final, 1)
        If data_flux_all_final(I, 1) <> "" Then
            For j = 3 To UBound(data_flux_all_final, 2) - 6
                If left(data_flux_all_final(I, j), 4) = "Hot_" Or left(data_flux_all_final(I, j), 4) = "Col_" Then
                    nb_puissance_restants = nb_puissance_restants + 1
                    ReDim Preserve puissances_non_echangees(1 To P_non_echangee_nb_caracteristiques, 1 To nb_puissance_restants)
                    puissances_non_echangees(P_non_echangee_puisssance, nb_puissance_restants) = CDbl(Right(data_flux_all_final(I, j), Len(data_flux_all_final(I, j)) - 4))    'puissance
                    puissances_non_echangees(P_non_echangee_T_chaude, nb_puissance_restants) = CDbl(data_flux_all_final(I, j - 1))    'température gauche
                    puissances_non_echangees(P_non_echangee_T_froide, nb_puissance_restants) = CDbl(data_flux_all_final(I, j + 1))    'température droite
                    puissances_non_echangees(P_non_echangee_nom_flux, nb_puissance_restants) = data_flux_all_final(I, 2)    'nom du flux
                ElseIf left(data_flux_all_final(I, j), 2) = "R_" Then
                    nb_puissance_restants = nb_puissance_restants + 1
                    ReDim Preserve puissances_non_echangees(1 To P_non_echangee_nb_caracteristiques, 1 To nb_puissance_restants)
                    puissances_non_echangees(P_non_echangee_puisssance, nb_puissance_restants) = CDbl(Right(data_flux_all_final(I, j), Len(data_flux_all_final(I, j)) - 2))  'puissance
                    puissances_non_echangees(P_non_echangee_T_chaude, nb_puissance_restants) = CDbl(data_flux_all_final(I, j - 1))    'température gauche
                    If data_flux_all_final(I, j + 1) = repere_traits_Clement Then
                        'METTRE MESSAGE ERREUR OU FAIRE QUELQUE CHOSE / DANS CE CAS LA TEMPERATURE ETAIT A LA POSITION j+3
                    Else
                        puissances_non_echangees(P_non_echangee_T_froide, nb_puissance_restants) = CDbl(data_flux_all_final(I, j + 1))    'température droite
                    End If
                    puissances_non_echangees(P_non_echangee_nom_flux, nb_puissance_restants) = data_flux_all_final(I, 2)    'nom du flux
                End If
            Next j
        End If
    Next I
'Ajout DCA 24/05/2014 pour actualisation en cas de changement de langue
    'On supprimme les éventuels restes et sauvegarde la matrice dans excel (pour changement de langue)
    I = 1
    j = 1
    Do While Feuil63.Range("PINCH_RepresentationDCAVecteurFluxChauds_" & marqueur_etape & pi_FiliereConsideree).Cells(I, j) <> ""
        Do While Feuil63.Range("PINCH_RepresentationDCAVecteurFluxChauds_" & marqueur_etape & pi_FiliereConsideree).Cells(I, j) <> ""
            Feuil63.Range("PINCH_RepresentationDCAVecteurFluxChauds_" & marqueur_etape & pi_FiliereConsideree).Cells(I, j).ClearContents
            I = I + 1
        Loop
        I = 1
        j = j + 1
    Loop
    If nb_puissance_restants > 0 Then
        For I = 1 To P_non_echangee_nb_caracteristiques
            For j = 1 To nb_puissance_restants
                Feuil63.Range("PINCH_RepresentationDCAVecteurFluxChauds_" & marqueur_etape & pi_FiliereConsideree).Cells(I, j) = puissances_non_echangees(I, j)
            Next j
        Next I
        'on remplace les noms de flux par des formules
        Call ecriture_formule_pour_changement_langue_flux_et_utilites_dans_plage_cellules("PINCH_RepresentationDCAVecteurFluxChauds_" & marqueur_etape & pi_FiliereConsideree)
    End If
    

End Sub


Public Sub VeriPinch(state As Integer)
    'Procédure générale qui réalise l'analyse Pinch

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String
    Dim position_erreur As String

    'state      état d'avancement du Pinch. 1 = CC+HEN+Utilities ; 2 = HEN+Utilities ; 3 = Utilities
    'step    indice de la filière calculée

    Dim data_flux_ch() As String        'données des flux chauds
    Dim data_flux_fr() As String        'données des flux froids
    '| Nom | Tin | Tout | MCp | Enthalpie |

    Dim gestion_ch() As Double          'données des points de la CC chaude
    Dim gestion_fr() As Double          'données des points de la CC froide
    '| température du point | DH correspondant | DH cumulé |

    Dim deltat_ch() As Double           'données des points de la CC chaude, avec calcul du DT_mesure
    Dim deltat_fr() As Double           'données des points de la CC froide, avec calcul du DT_mesure
    '| température du point | enthalpie du point | DT mesuré au point |

    Dim deltat_ch_raw() As Double       'données des points de la CC chaude avant décalage
    Dim deltat_fr_raw() As Double       'données des points de la CC froide avant décalage

    'Variables itératives
    Dim I As Integer
    Dim j As Double
    Dim k As Integer
    Dim m As Integer
    Dim N As Integer

    'Divers
    Dim colonneStation As Integer   'colonne pour la lecture
    Dim nbflux_ch As Integer        'nombre de flux chauds (calculé en lisant data_flux)
    Dim nbflux_fr As Integer        'nombre de flux froids (calculé en lisant data_flux)
    Dim LigneTab As Integer         'ligne de lecture des informations dans la feuille "pincement thermique"

    'Pour le décalage de la CC froide dans la détermination du point de Pinch
    Dim Tchgtetat As Double         'température de changement d'état si rencontré
    Dim TchgtetatPrev As Double     'si on a trouvé un chgt d'état, on vérifie la température du changt d'état suivant
    Dim chgtetat As Boolean         'binaire déterminant si on a trouvé un chgt d'état
    Dim deltaenthalpie As Double    'deltaenthalpie calculé à chaque point des CC (utile pour le DH cumulé)
    Dim minimal As Double           'pincement calculé (DTmin)
    Dim pinch_ch As Double          'détermine la CC sur laquelle le DTmin est calculé
    Dim pinch_fr As Double          'détermine la CC sur laquelle le DTmin est calculé
    Dim h_fr As Double              'enthalpie du point à partir duquel on a trouvé le DTmin
    Dim h_ch As Double              'enthalpie du point à partir duquel on a trouvé le DTmin
    Dim iteration As Integer        'nombre d'itérations
    Dim pas As Double               'pas de décalage de la CC (initialisé au début puis affiné pour accélérer le traitement)
    Dim precision As Double         'précision pour le calcul de pinch_mesure (1 > precision > 0)
    Dim ajustement As Double        'valeur de l'ajustement pour coler exactement au Pinch voulu

    'Après le décalage de la CC froide, calcul des principales grandeurs du problème
    Dim hot_need As Double          'besoin chaud restant
    Dim cold_need As Double         'besoin froid restant
    Dim recup_max As Double         'récupération énergétique maximale

    'Coordonnées des points de la CC (chaines de caractère)
    Dim name_fr As String
    Dim name_ch As String
    Dim col_fr As Integer
    Dim col_ch As Integer
    Dim cc_x_ch As String
    Dim cc_y_ch As String
    Dim cc_x_fr As String
    Dim cc_y_fr As String

    'Affichage du besoin Chaud et besoin Froid sur les CC
    Dim max_t_max As Double         'température maximale trouvée sur les CC
    Dim min_t_min As Double         'temp minimale trouvée sur les CC
    Dim decal As Double             'décalage des Besoins sur le graphique des CC
    Dim idlegend As Integer         'numéro de série pour affichage graphique des besoins Ch et Fr sur les courbes
    Dim idseries As Integer         'numéro de légende pour affichage graphique des besoins Ch et Fr sur les courbes
    Dim low_x_ch As String          'bornes min et max des besoins pour affichage sur CC
    Dim high_x_ch As String
    Dim low_y_ch As String
    Dim high_y_ch As String
    Dim low_x_fr As String
    Dim high_x_fr As String
    Dim low_y_fr As String
    Dim high_y_fr As String

    'Construction du HEN
    Dim tipe As String              'tipe de flux : Ch ou Fr
    Dim t_pinch_gch As Double           'température "supérieure" du pincement
    Dim t_pinch_dte As Double           'température "inférieure" du pincement
    Dim tmax As Double              'tmax du flux
    Dim tmin As Double              'tmin du flux
    Dim reste_dte As Double           'puissance restante à affecter au flux en dessous du pinch
    Dim reste_gch As Double           'puissance restante à affecter au flux au dessus du pinch
    Dim nom_flux As String              'nom du flux au dessus du pinch

    'Quand on ajoute les HEX à la main
    Dim data_hex_imposed() As String        'caractéristiques des HEX à imposer
    Dim colres1 As Integer          'indique le coté (4 ou 5) du HEX construit. besoin de passer une valeur sous forme de variable apparamment
    Dim colres2 As Integer          'indique le coté (4 ou 5) du HEX construit

    'Quand on ajoute les HEX automatiquement
    Dim last As Double              'dernier flux de la boucle imbriquée de type opposé mais ne respectant pas le mcp
    Dim ligch As Integer            'ligne du flux chaud concerné par un HEX
    Dim ligfr As Integer            'ligne du flux froid concerné par un HEX

    'Finalisation - calcul des besoins restants à couvrir par les utilités
    Dim tprev As Double
    Dim tnext As Double
    Dim label As String
    Dim restegch_ch As Double           'besoins chauds restants à gauche du pinch
    Dim restedte_ch As Double           'besoins chauds restants à droite du pinch
    Dim restegch_fr As Double           'besoins froids restants à gauche du pinch
    Dim restedte_fr As Double           'besoins froids restants à droite du pinch

    'Couverture des besoins restants
    Dim nbneeds_gch_ch As Integer       'nombre de besoins chauds à gauche du pinch (T supérieure à Tpinch)
    Dim nbneeds_gch_fr As Integer       'nombre de besoins froids à gauche du pinch (T supérieure à Tpinch)
    Dim nbneeds_dte_ch As Integer       'nombre de besoins chauds à droite du pinch (T inférieure à Tpinch)
    Dim nbneeds_dte_fr As Integer       'nombre de besoins froids à droite du pinch (T inférieure à Tpinch)
    Dim data_needs_gch_ch() As String   'Variables contenant les besoins. "_2" : transposées pour un affichage plus ergonomique
    Dim data_needs_gch_fr() As String
    Dim data_needs_dte_ch() As String
    Dim data_needs_dte_fr() As String
    Dim data_needs_gch_ch_2() As String
    Dim data_needs_gch_fr_2() As String
    Dim data_needs_dte_ch_2() As String
    Dim data_needs_dte_fr_2() As String
    Dim data_needs_ch() As String
    Dim data_needs_fr() As String
    Dim nbneeds As Integer          'nombre total de besoins à couvrir
    Dim par As Integer              'nombre d'informations stockées dans data_needs_xxx_xx (utile pour le tri)
    Dim a1 As Integer   'indices utilisés pour le passage d'argument dans certaines subroutines (un sombre problème de ou )
    Dim a2 As Integer
    Dim b1 As Integer
    Dim b2 As Integer
    Dim temp As Integer

    'Affinage du HEN
    Dim trait As String         'pour du traitement de chaine de caractère
    Dim trait2 As String        'pour du traitement de chaine de caractère
    Dim diagonal_max As Integer    'nombre maximal d'échangeurs diagonals acceptés (si <2, risque de planter VeriPinch)
    Dim diff As Integer         'nombre d'échangeurs diagonals repérés
    Dim affbool As Boolean      'affinage oui/non
    Dim colmax As Integer       'colonne d'écriture la plus éloignée des cotés du HEN
    Dim vide As Boolean         'indicateur de colonne vide
    Dim vide2 As Boolean        'idem : on teste 2 colonnes successivement
    Dim sec As Boolean
    Dim jmax As Integer

    
    'MODIF CRE 05/09/2012   Désactivation de la mise à jour de l'affichage à l'écran pendant l'éxécution du code
    Application.ScreenUpdating = False
    separateur_decimal = Application.International(xlDecimalSeparator)

    If state <= 1 And state <> 0 Then
        'I. Tracé des Courbes Composites
        'I.0. --------------------------------   Initialisations générales   -----------------------------------
        'I.1. ------------------------------------   Remises à Zero   ------------------------------------------
        'I.2. -------------------   Rangement des flux selon leur type : Chaud-Froid   -------------------------
        'I.3. ----------------   Traitement des courants Chauds/Froids (tri, doublons etc.)   ------------------
        'I.4. --------------------   Calcul du décalage nécessaire pour vérifier le Pinch   --------------------
        'I.5. -----------------------   Calcul du besoin Chaud et Froid, Pinch et DTMin   ----------------------
        'I.6. ------------------------------   Tracé des Courbes Composites   ----------------------------------
        'I.7. ---------------------------------   Construction du HEN (vide)  ----------------------------------
        'I.8. ---------------------------------------   Finalization   -----------------------------------------
        'I.9. -------------------------------------------   Fin   ----------------------------------------------



        'I.0. --------------------------------   Initialisations générales   -----------------------------------

        position_erreur = "  -I.0"

        If PINCH_DANS_OCEAN = False Then
            Call lecture_flux_process
        End If

        If PINCH_ORDONNER_FLUX_PROCESS = True Then
            Call ordonner_flux_process
        End If

        'CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).Value = "Debut du calcul à " & Hour(Time) & "h " & Minute(Time) & "min " & Second(Time) & "s"
        precision = 0
        pas = 10
        If PINCH_DANS_OCEAN = True Then
            col_ch = DCA_CC_ColorIndexBesoinsChauds
            col_fr = DCA_CC_ColorIndexBesoinsFroids
        Else
            col_ch = 7
            col_fr = 23
        End If
        ligprint = 4
        ok = True
        erreur = ""
        CoefU = 1.5

        data_flux_hen_size_empty = 18           'taille du hen vierge
        data_flux_hen_size = data_flux_hen_size_empty

        besoins_HOT = 0
        besoins_COLD = 0

        'On récupère les données sur les flux de procédés
        'Redimensionnement et remplissage de la variable data_flux
        nbflux = 0
        ReDim data_flux(1 To flux_process_nb_parametres, 1 To 1)
        For I = 1 To UBound(flux_thermique_process, 2)
            If flux_thermique_process(repere_flux_thermique_process_etat, I) = True Then
                nbflux = nbflux + 1
                ReDim Preserve data_flux(1 To flux_process_nb_parametres, 1 To nbflux)
                data_flux(flux_process_nom, nbflux) = CStr(flux_thermique_process(repere_flux_thermique_process_nom, I))
                data_flux(flux_process_T_in_°C, nbflux) = CStr(flux_thermique_process(repere_flux_thermique_process_T_in_°C, I))
                data_flux(flux_process_T_out_°C, nbflux) = CStr(flux_thermique_process(repere_flux_thermique_process_T_out_°C, I))
                data_flux(flux_process_MCp_kW_°C, nbflux) = CStr(flux_thermique_process(repere_flux_thermique_process_MCp_kW_°C, I))
                'Puissance du flux calculée
                data_flux(flux_process_Pw_kW, nbflux) = CStr(CDbl(data_flux(flux_process_MCp_kW_°C, nbflux)) * (CDbl(data_flux(flux_process_T_out_°C, nbflux)) - CDbl(data_flux(flux_process_T_in_°C, nbflux))))
            End If
        Next I

        'Valeur de pincement minimal
        colonneStation = pi_FiliereConsideree
        LigneTab = 4
        pinch_target = Feuil63.Range("DonneesPincementPinch").Cells(LigneTab, colonneStation)



        'I.1. ------------------------------------   Remises à Zero   ------------------------------------------
        position_erreur = "  -I.1"

        Call effacer_resultats



        'I.2. -------------------   Rangement des flux selon leur type : Chaud-Froid   -------------------------
        position_erreur = "  -I.2"

        'Détermination du nombre de flux chauds et froids séparément
        For I = 1 To nbflux
            If CDbl(data_flux(flux_process_Pw_kW, I)) < 0 Then
                nbflux_ch = nbflux_ch + 1
            Else
                nbflux_fr = nbflux_fr + 1
            End If
        Next

        'Remplissage des tableaux de flux chauds
        If nbflux_ch > 0 Then
            ReDim data_flux_ch(1 To nbflux_ch, 1 To flux_process_nb_parametres)
            j = 1
            For I = 1 To nbflux
                If CDbl(data_flux(flux_process_Pw_kW, I)) < 0 Then
                    For m = 1 To flux_process_nb_parametres
                        data_flux_ch(j, m) = data_flux(m, I)
                    Next
                    j = j + 1
                End If
            Next
        End If

        'Remplissage des tableaux de flux chauds
        If nbflux_fr > 0 Then
            ReDim data_flux_fr(1 To nbflux_fr, 1 To flux_process_nb_parametres)
            j = 1
            For I = 1 To nbflux
                If CDbl(data_flux(flux_process_Pw_kW, I)) > 0 Then
                    For m = 1 To flux_process_nb_parametres
                        data_flux_fr(j, m) = data_flux(m, I)
                    Next
                    j = j + 1
                End If
            Next
        End If



        'I.3. ----------------   Traitement des courants Chauds/Froids (tri, doublons etc.)   ------------------
        position_erreur = "  -I.3"

        If nbflux_ch > 0 Then
            'Remplissage des températures des flux chauds
            ReDim gestion_ch(1 To nbflux_ch * 2, 1 To 3)
            j = 1
            For I = 1 To nbflux_ch
                gestion_ch(j, 1) = CDbl(data_flux_ch(I, flux_process_T_in_°C))
                gestion_ch(j + 1, 1) = CDbl(data_flux_ch(I, flux_process_T_out_°C))
                j = j + 2
            Next

            'Tri des températures des flux chauds
            Call Sort_Array(gestion_ch(), "desc", 1)

            'Gestion des doublons pour les flux chauds : suppression des doublon de température dans gestion(), sauf s'il s'agit d'un changement d'état
            Tchgtetat = 0
            TchgtetatPrev = 0
            chgtetat = False
            I = 1
            Do
                If gestion_ch(I, 1) = gestion_ch(I + 1, 1) Then
                    For j = 1 To nbflux_ch
                        If CDbl(data_flux_ch(j, flux_process_T_in_°C)) = gestion_ch(I, 1) Then
                            If data_flux_ch(j, flux_process_T_in_°C) = data_flux_ch(j, flux_process_T_out_°C) Then
                                chgtetat = True
                                Tchgtetat = CDbl(data_flux_ch(j, flux_process_T_in_°C))
                                j = nbflux_ch
                            End If
                        End If
                    Next
                    If (chgtetat = False Or TchgtetatPrev = Tchgtetat) Then
                        Call Delete_row(gestion_ch(), I)
                        I = I - 1
                    End If
                End If
                TchgtetatPrev = Tchgtetat
                I = I + 1
            Loop Until I = UBound(gestion_ch())

            'Calcul du DeltaEnthalpie par tranche de température : on somme le DeltaEnthalpie des flux sur chaque tranche de température rencontrée sur la CC des flux chauds
            For I = 1 To UBound(gestion_ch()) - 1
                deltaenthalpie = 0
                If gestion_ch(I, 1) <> gestion_ch(I + 1, 1) Then    'Pas de changement d'état dans le cas général
                    For j = 1 To nbflux_ch
                        'Si la tranche de température considérée est comprise dans l'intervalle de températures du flux, on ajoute la part de l'enthalpie du flux sur la tranche de température considérée
                        If gestion_ch(I, 1) <= CDbl(data_flux_ch(j, flux_process_T_in_°C)) And gestion_ch(I + 1, 1) >= CDbl(data_flux_ch(j, flux_process_T_out_°C)) Then
                            deltaenthalpie = deltaenthalpie + (gestion_ch(I, 1) - gestion_ch(I + 1, 1)) / (CDbl(data_flux_ch(j, flux_process_T_in_°C)) - CDbl(data_flux_ch(j, flux_process_T_out_°C))) * CDbl(data_flux_ch(j, flux_process_Pw_kW))
                        End If
                    Next
                Else
                    For j = 1 To nbflux_ch
                        If gestion_ch(I, 1) = CDbl(data_flux_ch(j, flux_process_T_in_°C)) And gestion_ch(I + 1, 1) = CDbl(data_flux_ch(j, flux_process_T_out_°C)) Then
                            deltaenthalpie = deltaenthalpie + CDbl(data_flux_ch(j, flux_process_Pw_kW))
                        End If
                    Next
                End If
                gestion_ch(I, 2) = deltaenthalpie
            Next
            gestion_ch(I, 2) = 0

            'On somme les Deltaenthalpie en remontant les températures
            For I = UBound(gestion_ch()) - 1 To 1 Step -1
                gestion_ch(I, 3) = gestion_ch(I + 1, 3) + Abs(gestion_ch(I, 2))
            Next
        End If

        If nbflux_fr > 0 Then
            'Remplissage des températures des flux froids
            ReDim gestion_fr(1 To nbflux_fr * 2, 1 To 3)
            j = 1
            For I = 1 To nbflux_fr
                gestion_fr(j, 1) = CDbl(data_flux_fr(I, flux_process_T_in_°C))
                gestion_fr(j + 1, 1) = CDbl(data_flux_fr(I, flux_process_T_out_°C))
                j = j + 2
            Next

            'Tri des températures des flux froids
            Call Sort_Array(gestion_fr(), "asc", 1)

            'Gestion des doublons pour les flux froids : suppression des doublon de température dans gestion(), sauf s'il s'agit d'un changement d'état
            Tchgtetat = 0
            TchgtetatPrev = 0
            chgtetat = False
            I = 1
            Do
                If gestion_fr(I, 1) = gestion_fr(I + 1, 1) Then
                    For j = 1 To nbflux_fr
                        If CDbl(data_flux_fr(j, flux_process_T_in_°C)) = gestion_fr(I, 1) Then
                            If data_flux_fr(j, flux_process_T_in_°C) = data_flux_fr(j, flux_process_T_out_°C) Then
                                chgtetat = True
                                Tchgtetat = CDbl(data_flux_fr(j, flux_process_T_in_°C))
                                j = nbflux_fr
                            End If
                        End If
                    Next
                    If (chgtetat = False Or TchgtetatPrev = Tchgtetat) Then
                        Call Delete_row(gestion_fr(), I)
                        I = I - 1
                    End If
                End If
                TchgtetatPrev = Tchgtetat
                I = I + 1
            Loop Until I = UBound(gestion_fr())

            'Calcul du DeltaEnthalpie par tranche de température : on somme le DeltaEnthalpie des flux sur chaque tranche de température rencontrée sur la CC des flux froids
            For I = 1 To UBound(gestion_fr()) - 1
                deltaenthalpie = 0
                If gestion_fr(I, 1) <> gestion_fr(I + 1, 1) Then    'Pas de changement d'état dans le cas général
                    For j = 1 To nbflux_fr
                        'Si la tranche de température considérée est comprise dans l'intervalle de températures du flux, on ajoute la part de l'enthalpie du flux sur la tranche de température considérée
                        If gestion_fr(I, 1) >= CDbl(data_flux_fr(j, flux_process_T_in_°C)) And gestion_fr(I + 1, 1) <= CDbl(data_flux_fr(j, flux_process_T_out_°C)) Then
                            deltaenthalpie = deltaenthalpie + (gestion_fr(I + 1, 1) - gestion_fr(I, 1)) / (CDbl(data_flux_fr(j, flux_process_T_out_°C)) - CDbl(data_flux_fr(j, flux_process_T_in_°C))) * CDbl(data_flux_fr(j, flux_process_Pw_kW))
                        End If
                    Next
                Else
                    For j = 1 To nbflux_fr
                        If gestion_fr(I, 1) = CDbl(data_flux_fr(j, flux_process_T_in_°C)) And gestion_fr(I + 1, 1) = CDbl(data_flux_fr(j, flux_process_T_out_°C)) Then
                            deltaenthalpie = deltaenthalpie + CDbl(data_flux_fr(j, flux_process_Pw_kW))
                        End If
                    Next
                End If
                gestion_fr(I, 2) = deltaenthalpie
            Next

            'On somme les Deltaenthalpie en descendant les températures
            For I = 2 To UBound(gestion_fr())
                gestion_fr(I, 3) = gestion_fr(I - 1, 3) + gestion_fr(I - 1, 2)
            Next
        End If



        'I.4. --------------------   Calcul du décalage nécessaire pour vérifier le Pinch   --------------------
        position_erreur = "  -I.4"

        'Déclaration des variables
        If nbflux_ch > 0 Then
            ReDim deltat_ch(1 To UBound(gestion_ch()), 1 To 3)
            For I = 1 To UBound(gestion_ch())
                deltat_ch(I, 1) = gestion_ch(I, 1)
                deltat_ch(I, 2) = gestion_ch(I, 3)
            Next
            ReDim deltat_ch_raw(1 To UBound(gestion_ch()), 1 To 3)
            deltat_ch_raw = deltat_ch   'sauvegarde avant décalage
        End If
        If nbflux_fr > 0 Then
            ReDim deltat_fr(1 To UBound(gestion_fr()), 1 To 3)
            For I = 1 To UBound(gestion_fr())
                deltat_fr(I, 1) = gestion_fr(I, 1)
                deltat_fr(I, 2) = gestion_fr(I, 3)
            Next
            ReDim deltat_fr_raw(1 To UBound(gestion_fr()), 1 To 3)
            deltat_fr_raw = deltat_fr   'sauvegarde avant décalage
        End If

        'Calcul des DTmin à chaque point : on boucle sur les points de la CC chaude, puis boucle imbriquée mais inverse sur les points de la CC froide. ensuite, boucle sur les points de la CC froide avec une boucle imbriquée (mais inverse) sur les points de la CC chaude.
        iteration = 0
        point_pinch = -1000    'MODIF CRE 25/09/2012   initialisation à -1000 pour éviter les conflits où une CC a un point à 0°C
        If nbflux_ch > 0 And nbflux_fr > 0 Then
            pas = deltat_ch(1, 2) / 10000    'on ajuste le pas de décalage de la CC froide en fonction des données du problème (peut accélérer le traitement)
            Do
                If iteration > 0 Then    'décalage de la CC froide seulement après la première itération
                    For I = 1 To UBound(deltat_fr())
                        deltat_fr(I, 2) = deltat_fr(I, 2) + pas
                    Next
                End If
                If iteration = 1 Then    'après la première itération, on entre en dur le pinch infini sur le premier point de la CC Chaude
                    deltat_ch(UBound(deltat_ch()), 3) = 1000
                End If
                For I = 1 To UBound(deltat_ch()) - 1
                    For j = UBound(deltat_fr()) To 2 Step -1
                        If deltat_ch(I, 2) > deltat_fr(j, 2) Then    'le point Ch est en dehors de la CC froide : pinch infini et on passe au point chaud suivant (test en début, parce que ça peut arriver dès le départ)
                            deltat_ch(I, 3) = 1000
                            j = 2
                        ElseIf deltat_ch(I, 2) = deltat_fr(j, 2) Then    'le point Ch est à la verticale (en enthalpie) d'un point Froid : on sort
                            If deltat_ch(I, 2) <> deltat_ch(I + 1, 2) Then    'point chaud
                                If deltat_fr(j, 2) <> deltat_fr(j - 1, 2) Then    'point froid (cas 0)
                                    deltat_ch(I, 3) = deltat_ch(I, 1) - deltat_fr(j, 1)
                                Else    'saut froid (cas 1)
                                    If I = 1 Then    'c'est le premier point ch rencontré : on calcule le DTmin avec le froid du bas (cas 1a)
                                        deltat_ch(I, 3) = deltat_ch(I, 1) - deltat_fr(j - 1, 1)
                                    Else    'si c'est pas le premier (dernier y compris), on calcule avec le point froid du dessus (cas 1b)
                                        deltat_ch(I, 3) = deltat_ch(I, 1) - deltat_fr(j, 1)
                                    End If
                                End If
                            Else    'saut chaud
                                If deltat_fr(j, 2) <> deltat_fr(j + 1, 2) Then    'point froid (cas 3)
                                    deltat_ch(I, 3) = deltat_ch(I, 1) - deltat_fr(j, 1)    'ici, on est sur la CC ou sur l'avant-dernier point. vérification du dernier point dans la boucle des points froids
                                    deltat_ch(I + 1, 3) = deltat_ch(I + 1, 1) - deltat_fr(j, 1)    'on en profite pour calculer le Pinch au point ch bas et on itère sur i pour passer le saut
                                    I = I + 1
                                Else    'saut froid (cas 4)
                                    deltat_ch(I, 3) = deltat_ch(I, 1) - deltat_fr(j, 1)
                                    deltat_ch(I + 1, 3) = deltat_ch(I + 1, 1) - deltat_fr(j - 1, 1)    'on calcule aussi le pincement pour le bas des sauts et on itère au point Ch suivant (i=i+1)
                                    I = I + 1
                                End If
                            End If
                            j = 2
                        ElseIf deltat_ch(I, 2) < deltat_fr(j, 2) And deltat_ch(I, 2) > deltat_fr(j - 1, 2) Then    'point Ch compris entre 2 points Froids : on calcule le pinch et on sort de la boucle froide
                            deltat_ch(I, 3) = deltat_ch(I, 1) - ((deltat_ch(I, 2) - deltat_fr(j - 1, 2)) / (deltat_fr(j, 2) - deltat_fr(j - 1, 2)) * (deltat_fr(j, 1) - deltat_fr(j - 1, 1)) + deltat_fr(j - 1, 1))
                            j = 2
                        ElseIf deltat_ch(I, 2) < deltat_fr(j, 2) Then    'ne pas sortir de la boucle des Ch pour trouver le point froid suivant
                            deltat_ch(I, 3) = 1000
                        End If
                    Next
                Next
                For I = 1 To UBound(deltat_fr()) - 1
                    For j = UBound(deltat_ch()) To 2 Step -1
                        If deltat_fr(I, 2) < deltat_ch(j, 2) Then    'le point Ch est en dehors de la CC froide : pinch infini et on passe au point chaud suivant (test en début, parce que ça peut arriver dès le départ)
                            deltat_fr(I, 3) = 1000
                            j = 2
                        ElseIf deltat_fr(I, 2) = deltat_ch(j, 2) Then    'le point Ch est à la verticale (en enthalpie) d'un point Froid : on sort
                            If deltat_fr(I, 2) <> deltat_fr(I + 1, 2) Then    'point chaud
                                If deltat_ch(j, 2) <> deltat_ch(j - 1, 2) Then    'point froid (cas 0)
                                    deltat_fr(I, 3) = deltat_ch(j, 1) - deltat_fr(I, 1)
                                Else    'saut froid (cas 1)
                                    If I = 1 Then    'c'est le premier point ch rencontré : on calcule le DTmin avec le froid du bas (cas 1a)
                                        deltat_fr(I, 3) = deltat_ch(j - 1, 1) - deltat_fr(I, 1)
                                    Else    'si c'est pas le premier (dernier y compris), on calcule avec le point froid du dessus (cas 1b)
                                        deltat_fr(I, 3) = deltat_ch(j, 1) - deltat_fr(I, 1)
                                    End If
                                End If
                            Else    'saut chaud
                                If deltat_ch(j, 2) <> deltat_ch(j + 1, 2) Then    'point froid (cas 3)
                                    deltat_fr(I, 3) = deltat_ch(j, 1) - deltat_fr(I, 1)    'ici, on est sur la CC ou sur l'avant-dernier point. vérification du dernier point dans la boucle des points froids
                                    deltat_fr(I + 1, 3) = deltat_ch(j, 1) - deltat_fr(I + 1, 1)    'on en profite pour calculer le Pinch au point ch bas et on itère sur i pour passer le saut
                                    I = I + 1
                                Else    'saut froid (cas 4)
                                    deltat_fr(I, 3) = deltat_ch(j, 1) - deltat_fr(I, 1)
                                    deltat_fr(I + 1, 3) = deltat_ch(j - 1, 1) - deltat_fr(I + 1, 1)    'on calcule aussi le pincement pour le bas des sauts et on itère au point Ch suivant (i=i+1)
                                    I = I + 1
                                End If
                            End If
                            j = 2
                        ElseIf deltat_fr(I, 2) > deltat_ch(j, 2) And deltat_fr(I, 2) < deltat_ch(j - 1, 2) Then    'point Ch compris entre 2 points Froids : on calcule le pinch et on sort de la boucle froide
                            deltat_fr(I, 3) = ((deltat_fr(I, 2) - deltat_ch(j - 1, 2)) / (deltat_ch(j, 2) - deltat_ch(j - 1, 2)) * (deltat_ch(j, 1) - deltat_ch(j - 1, 1)) + deltat_ch(j - 1, 1)) - deltat_fr(I, 1)
                            j = 2
                        ElseIf deltat_fr(I, 2) > deltat_ch(j, 2) Then    'ne pas sortir de la boucle des Ch pour trouver le point froid suivant
                            deltat_fr(I, 3) = 1000
                        End If
                    Next
                Next
                'Gestion du dernier point de la CC froide (le dernier point de la CC chaude est toujours sur l'axe des ordonnées)
                last = UBound(deltat_fr())
                For I = 1 To UBound(deltat_ch())
                    If deltat_fr(last, 2) > deltat_ch(I, 2) Then    'le point Fr est en dehors de la CC chaude : pinch infini
                        deltat_fr(last, 3) = 1000
                        I = UBound(deltat_ch())
                    ElseIf deltat_fr(last, 2) = deltat_ch(I, 2) Then    'le point Ch est à la verticale (en enthalpie) d'un point Froid : on sort
                        If deltat_ch(I, 2) <> deltat_ch(I + 1, 2) Then    'point chaud (cas n)
                            deltat_fr(last, 3) = deltat_ch(I, 1) - deltat_fr(last, 1)
                        Else    'saut chaud: on calcule le DTmin avec le bras du bas (cas n+1)
                            deltat_fr(last, 3) = deltat_ch(I + 1, 1) - deltat_fr(last, 1)
                        End If
                        I = UBound(deltat_ch())
                    ElseIf deltat_fr(last, 2) < deltat_ch(I, 2) And deltat_fr(last, 2) > deltat_ch(I + 1, 2) Then    'point Ch compris entre 2 points Froids : on calcule le pinch et on sort de la boucle froide
                        deltat_fr(last, 3) = ((deltat_fr(last, 2) - deltat_ch(I + 1, 2)) / (deltat_ch(I, 2) - deltat_ch(I + 1, 2)) * (deltat_ch(I, 1) - deltat_ch(I + 1, 1)) + deltat_ch(I + 1, 1)) - deltat_fr(last, 1)
                        I = UBound(deltat_ch())
                    ElseIf deltat_fr(last, 2) < deltat_ch(I, 2) Then    'ne pas sortir de la boucle des Ch pour trouver le point froid suivant
                        deltat_fr(last, 3) = 1000
                    End If
                Next
                'Gestion du premier point de la CC chaude (seulement pendant la première itération, au cas où le pinch serait déja validé)
                If iteration = 0 Then
                    deltat_ch(UBound(deltat_ch()), 3) = deltat_fr(1, 3)
                End If

                Call FindPinch(deltat_ch(), deltat_fr(), pinch_ch, pinch_fr, minimal, h_ch, h_fr)
                iteration = iteration + 1
            Loop While (minimal < pinch_target / (1 + precision))

            pinch_mesure = minimal

            'Ajustement au pinch exact, seulement si on n'est pas trop loin du pinch_target
            ajustement = 0
            If iteration > 1 And pinch_mesure < pinch_target * (1.1) Then
                If pinch_fr <> -1000 Then    'pinch trouvé sur un point froid
                    For I = 1 To UBound(deltat_ch()) - 1
                        If deltat_ch(I, 2) > h_fr And h_fr > deltat_ch(I + 1, 2) Then
                            ajustement = ((pinch_fr + pinch_target) - deltat_ch(I + 1, 1)) / (deltat_ch(I, 1) - deltat_ch(I + 1, 1)) * (deltat_ch(I, 2) - deltat_ch(I + 1, 2)) + deltat_ch(I + 1, 2) - h_fr
                            I = UBound(deltat_ch()) - 1
                        End If
                    Next
                Else    'pinch trouvé sur un point chaud
                    For I = 1 To UBound(deltat_fr()) - 1
                        If deltat_fr(I, 2) < h_ch And h_ch < deltat_fr(I + 1, 2) Then
                            ajustement = -(deltat_fr(I + 1, 2) - h_ch - (deltat_fr(I + 1, 1) - (pinch_ch - pinch_target)) / (deltat_fr(I + 1, 1) - deltat_fr(I, 1)) * (deltat_fr(I + 1, 2) - deltat_fr(I, 2)))
                            I = UBound(deltat_fr()) - 1
                        End If
                    Next
                End If
            End If
            'ATTENTION : l'ajustement calculé peut être source d'erreurs. commenter la ligne qui suit pour activer l'ajustement.
            'ajustement = 0
            If ajustement <> 0 Then
                For I = 1 To UBound(deltat_fr())
                    deltat_fr(I, 2) = deltat_fr(I, 2) + ajustement
                Next
                pinch_mesure = pinch_target
            End If


        End If  'fin du si sur flux chauds et flux froids

        If nbflux_ch = 0 Or nbflux_fr = 0 Then
            pinch_mesure = 10
        End If


        'I.5. -----------------------   Calcul du besoin Chaud et Froid, Pinch et DTMin   ----------------------
        position_erreur = "  -I.5"

        If nbflux_ch > 0 And nbflux_fr > 0 Then
            If pinch_ch > pinch_fr Then
                point_pinch = pinch_ch - pinch_mesure / 2
            Else
                point_pinch = pinch_fr + pinch_mesure / 2
            End If
            hot_need = Application.Max(deltat_fr(UBound(deltat_fr()), 2) - deltat_ch(1, 2), 0)
            cold_need = Application.Max(deltat_ch(1, 2) - deltat_fr(UBound(deltat_fr()), 2), 0) + Application.Max(deltat_fr(1, 2) - deltat_ch(UBound(deltat_ch()), 2), 0)
            recup_max = deltat_ch(1, 2) - cold_need
        ElseIf nbflux_ch > 0 Then
            cold_need = deltat_ch(1, 2)
        ElseIf nbflux_fr > 0 Then
            hot_need = deltat_fr(UBound(deltat_fr()), 2)
        End If

        'on cherche les valeurs min et max des abscisses
        'MODIF CRE 05/09/2012   J'ai changé ces variables en Double, ça plantait pour des grandes valeurs de puissance (en l'occurence 34000 kW)
        graph_abscisse_min = 0   'on le fixe d'entrée car il ne changera pas
        graph_ordonnee_min = 9999
        graph_abscisse_max = -9999
        graph_ordonnee_max = -9999
        If nbflux_ch > 0 Then
            'ordonnée mini
            If graph_ordonnee_min > deltat_ch(UBound(deltat_ch, 1), 1) Then
                graph_ordonnee_min = deltat_ch(UBound(deltat_ch, 1), 1)
            End If
            If graph_ordonnee_min > deltat_ch_raw(UBound(deltat_ch_raw, 1), 1) Then
                graph_ordonnee_min = deltat_ch_raw(UBound(deltat_ch_raw, 1), 1)
            End If
            'abscisse max
            If graph_abscisse_max < deltat_ch(1, 2) Then
                graph_abscisse_max = deltat_ch(1, 2)
            End If
            If graph_abscisse_max < deltat_ch_raw(1, 2) Then
                graph_abscisse_max = deltat_ch_raw(1, 2)
            End If
            'ordonnée maxi
            If graph_ordonnee_max < deltat_ch(1, 1) Then
                graph_ordonnee_max = deltat_ch(1, 1)
            End If
            If graph_ordonnee_max < deltat_ch_raw(1, 1) Then
                graph_ordonnee_max = deltat_ch_raw(1, 1)
            End If
        End If
        If nbflux_fr > 0 Then
            'ordonnée mini
            If graph_ordonnee_min > deltat_fr(1, 1) Then
                graph_ordonnee_min = deltat_fr(1, 1)
            End If
            If graph_ordonnee_min > deltat_fr_raw(1, 1) Then
                graph_ordonnee_min = deltat_fr_raw(1, 1)
            End If
            'abscisse max
            If graph_abscisse_max < deltat_fr(UBound(deltat_fr, 1), 2) Then
                graph_abscisse_max = deltat_fr(UBound(deltat_fr, 1), 2)
            End If
            If graph_abscisse_max < deltat_fr_raw(UBound(deltat_fr_raw, 1), 2) Then
                graph_abscisse_max = deltat_fr_raw(UBound(deltat_fr_raw, 1), 2)
            End If
            'ordonnée maxi
            If graph_ordonnee_max < deltat_fr(UBound(deltat_fr, 1), 1) Then
                graph_ordonnee_max = deltat_fr(UBound(deltat_fr, 1), 1)
            End If
            If graph_ordonnee_max < deltat_fr_raw(UBound(deltat_fr_raw, 1), 1) Then
                graph_ordonnee_max = deltat_fr_raw(UBound(deltat_fr_raw, 1), 1)
            End If
        End If

        'finalisation
        If graph_ordonnee_min > 0 Then  'on place le minimum de température à 0 s'il est positif car plus lisible
            graph_ordonnee_min = 0
        End If

        graph_abscisse_min = arrondi_inf_borne_echelle(graph_abscisse_min)
        graph_abscisse_max = arrondi_sup_borne_echelle(graph_abscisse_max)
        graph_ordonnee_max = arrondi_sup_borne_echelle(graph_ordonnee_max)
        graph_ordonnee_min = arrondi_inf_borne_echelle(graph_ordonnee_min)



        'I.6. ------------------------------   Tracé des Courbes Composites   ----------------------------------
        position_erreur = "  -I.6"

        'MODIF CRE 16/10/2012   Affichage dans PINCH X
        If PINCH_DANS_OCEAN = True Then
            graph_CC_sans_decalage_top_points = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_CC_avant_decalage_" & marqueur_etape & pi_FiliereConsideree).top
                graph_CC_sans_decalage_left_points = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_CC_avant_decalage_" & marqueur_etape & pi_FiliereConsideree).left
                graph_CC_sans_decalage_hauteur_points = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_CC_avant_decalage_fin_" & marqueur_etape & pi_FiliereConsideree).top - CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_CC_avant_decalage_" & marqueur_etape & pi_FiliereConsideree).top
                graph_CC_sans_decalage_largeur_points = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_CC_avant_decalage_fin_" & marqueur_etape & pi_FiliereConsideree).left - CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_CC_avant_decalage_" & marqueur_etape & pi_FiliereConsideree).left
            graph_CC_avec_decalage_top_points = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_CC_apres_decalage_" & marqueur_etape & pi_FiliereConsideree).top
                graph_CC_avec_decalage_left_points = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_CC_apres_decalage_" & marqueur_etape & pi_FiliereConsideree).left
                graph_CC_avec_decalage_hauteur_points = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_CC_apres_decalage_fin_" & marqueur_etape & pi_FiliereConsideree).top - CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_CC_apres_decalage_" & marqueur_etape & pi_FiliereConsideree).top
                graph_CC_avec_decalage_largeur_points = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_CC_apres_decalage_fin_" & marqueur_etape & pi_FiliereConsideree).left - CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_CC_apres_decalage_" & marqueur_etape & pi_FiliereConsideree).left
        Else
            graph_CC_sans_decalage_top_points = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(42, 3).top
                graph_CC_sans_decalage_left_points = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(42, 3).left
                graph_CC_sans_decalage_hauteur_points = 250
                graph_CC_sans_decalage_largeur_points = 500
            graph_CC_avec_decalage_top_points = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(42, 12).top
                graph_CC_avec_decalage_left_points = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(42, 12).left
                graph_CC_avec_decalage_hauteur_points = 250
                graph_CC_avec_decalage_largeur_points = 500
        End If
    
        'Tracé des Courbes Composites avant décalage
        If nbflux_ch > 0 Then
            cc_x_ch = "={"
            cc_y_ch = "={"
            For I = 1 To UBound(deltat_ch())
                cc_x_ch = cc_x_ch & Round(deltat_ch_raw(I, 2), 1) & ";"
                cc_y_ch = cc_y_ch & Round(deltat_ch_raw(I, 1), 1) & ";"
            Next
            cc_x_ch = left(cc_x_ch, Len(cc_x_ch) - 1) & "}"
            cc_y_ch = left(cc_y_ch, Len(cc_y_ch) - 1) & "}"
            cc_x_ch = Replace(cc_x_ch, separateur_decimal, ".")
            cc_y_ch = Replace(cc_y_ch, separateur_decimal, ".")
            name_ch = CodeNames("Feuil5").Range("l_fluxChaud")
        End If
        If nbflux_fr > 0 Then
            cc_x_fr = "={"
            cc_y_fr = "={"
            For I = 1 To UBound(deltat_fr())
                cc_x_fr = cc_x_fr & Round(deltat_fr_raw(I, 2), 1) & ";"
                cc_y_fr = cc_y_fr & Round(deltat_fr_raw(I, 1), 1) & ";"
            Next
            cc_x_fr = left(cc_x_fr, Len(cc_x_fr) - 1) & "}"
            cc_y_fr = left(cc_y_fr, Len(cc_y_fr) - 1) & "}"
            cc_x_fr = Replace(cc_x_fr, separateur_decimal, ".")
            cc_y_fr = Replace(cc_y_fr, separateur_decimal, ".")
            name_fr = CodeNames("Feuil5").Range("l_fluxFroid")
        End If
        Call ChartDisplay("CC2", graph_CC_sans_decalage_top_points, graph_CC_sans_decalage_left_points, True, cc_x_fr, cc_y_fr, name_fr, col_fr, cc_x_ch, cc_y_ch, name_ch, col_ch)

        'Tracé des Courbes Composites après décalage
        If nbflux_ch > 0 Then
            cc_x_ch = "={"
            cc_y_ch = "={"
            For I = 1 To UBound(deltat_ch())
                cc_x_ch = cc_x_ch & Round(deltat_ch(I, 2), 1) & ";"
                cc_y_ch = cc_y_ch & Round(deltat_ch(I, 1), 1) & ";"
            Next
            cc_x_ch = left(cc_x_ch, Len(cc_x_ch) - 1) & "}"
            cc_y_ch = left(cc_y_ch, Len(cc_y_ch) - 1) & "}"
            cc_x_ch = Replace(cc_x_ch, separateur_decimal, ".")
            cc_y_ch = Replace(cc_y_ch, separateur_decimal, ".")
            name_ch = CodeNames("Feuil5").Range("l_fluxChauds")
        End If
        If nbflux_fr > 0 Then
            cc_x_fr = "={"
            cc_y_fr = "={"
            For I = 1 To UBound(deltat_fr())
                cc_x_fr = cc_x_fr & Round(deltat_fr(I, 2), 1) & ";"
                cc_y_fr = cc_y_fr & Round(deltat_fr(I, 1), 1) & ";"
            Next
            cc_x_fr = left(cc_x_fr, Len(cc_x_fr) - 1) & "}"
            cc_y_fr = left(cc_y_fr, Len(cc_y_fr) - 1) & "}"
            cc_x_fr = Replace(cc_x_fr, separateur_decimal, ".")
            cc_y_fr = Replace(cc_y_fr, separateur_decimal, ".")
            name_fr = CodeNames("Feuil5").Range("l_fluxFroids")
        End If
        Call ChartDisplay("CC", graph_CC_avec_decalage_top_points, graph_CC_avec_decalage_left_points, False, cc_x_fr, cc_y_fr, name_fr, col_fr, cc_x_ch, cc_y_ch, name_ch, col_ch)

        'ajout des besoins chauds et froids sur les CC
        If nbflux_ch > 0 And nbflux_fr > 0 Then
            max_t_max = Application.Max(deltat_ch(1, 1), deltat_fr(UBound(deltat_fr()), 1))
            min_t_min = Application.Min(deltat_ch(UBound(deltat_ch()), 1), deltat_fr(1, 1))
            idlegend = 2
            low_x_ch = deltat_ch(1, 2)
            high_x_ch = deltat_fr(UBound(deltat_fr()), 2)
            low_y_ch = deltat_fr(UBound(deltat_fr()), 1)
            high_y_ch = deltat_ch(1, 1)
            low_x_fr = 0
            high_x_fr = deltat_fr(1, 2)
            low_y_fr = deltat_ch(UBound(deltat_ch()), 1)
            high_y_fr = deltat_fr(1, 1)
        ElseIf nbflux_ch > 0 Then
            max_t_max = deltat_ch(1, 1)
            min_t_min = deltat_ch(UBound(deltat_ch()), 1)
            idlegend = 1
            low_x_fr = deltat_ch(1, 2)
            high_x_fr = 0
            high_y_fr = min_t_min
            low_y_fr = max_t_max
        ElseIf nbflux_fr > 0 Then
            max_t_max = deltat_fr(UBound(deltat_fr()), 1)
            min_t_min = deltat_fr(1, 1)
            idlegend = 1
            low_x_ch = 0
            high_x_ch = deltat_fr(UBound(deltat_fr()), 2)
            high_y_ch = min_t_min
            low_y_ch = max_t_max
        End If
        decal = Application.Max(max_t_max, min_t_min) / 16
        ActiveSheet.Shapes("CC").Select
        idseries = idlegend

        If hot_need > 0 Then
            If Abs(low_x_ch - high_x_ch) > 0 Then
                ActiveChart.SeriesCollection.NewSeries
                idseries = idseries + 1
                idlegend = idlegend + 1
                With ActiveChart.SeriesCollection(idseries)
                    .XValues = Replace("={" & low_x_ch & ";" & low_x_ch & "}", separateur_decimal, ".")
                    .Values = Replace("={" & high_y_ch & ";" & max_t_max + decal & "}", separateur_decimal, ".")
                    .Border.ColorIndex = 16
                    .MarkerStyle = xlNone
                    .Border.LineStyle = xlDot
                End With
                ActiveChart.Legend.LegendEntries(idlegend).Delete
                idlegend = idlegend - 1
                ActiveChart.SeriesCollection.NewSeries
                idseries = idseries + 1
                idlegend = idlegend + 1
                With ActiveChart.SeriesCollection(idseries)
                    .XValues = Replace("={" & low_x_ch & ";" & high_x_ch & "}", separateur_decimal, ".")
                    .Values = Replace("={" & max_t_max + decal & ";" & max_t_max + decal & "}", separateur_decimal, ".")
                    .Border.ColorIndex = 3
                    .Border.Weight = xlMedium
                    .MarkerStyle = xlNone
                    .Name = CodeNames("Feuil5").Range("l_besoinsChauds")
                End With
                ActiveChart.SeriesCollection.NewSeries
                idseries = idseries + 1
                idlegend = idlegend + 1
                With ActiveChart.SeriesCollection(idseries)
                    .XValues = Replace("={" & high_x_ch & ";" & high_x_ch & "}", separateur_decimal, ".")
                    .Values = Replace("={" & max_t_max + decal & ";" & low_y_ch & "}", separateur_decimal, ".")
                    .Border.ColorIndex = 16
                    .MarkerStyle = xlNone
                    .Border.LineStyle = xlDot
                End With
                ActiveChart.Legend.LegendEntries(idlegend).Delete
                idlegend = idlegend - 1
            End If
            If Abs(Abs(low_x_ch - high_x_ch) - hot_need) > 0.01 Then   'MODIF CRE 01/10/2012   Ajout de ce code pour gérer les cas où la CC froide englobe la CC chaude. Mais à priori on n'aura jamais besoin de lancer ce code car les 2 courbes sont collées à gauche
                ActiveChart.SeriesCollection.NewSeries
                idseries = idseries + 1
                idlegend = idlegend + 1
                With ActiveChart.SeriesCollection(idseries)
                    .XValues = Replace("={" & deltat_ch(UBound(deltat_ch()), 2) & ";" & deltat_ch(UBound(deltat_ch()), 2) & "}", separateur_decimal, ".")
                    .Values = Replace("={" & deltat_ch(UBound(deltat_ch()), 1) & ";" & deltat_ch(UBound(deltat_ch()), 1) + decal & "}", separateur_decimal, ".")
                    .Border.ColorIndex = 16
                    .MarkerStyle = xlNone
                    .Border.LineStyle = xlDot
                End With
                ActiveChart.Legend.LegendEntries(idlegend).Delete
                idlegend = idlegend - 1
                ActiveChart.SeriesCollection.NewSeries
                idseries = idseries + 1
                idlegend = idlegend + 1
                With ActiveChart.SeriesCollection(idseries)
                    .XValues = Replace("={" & deltat_ch(UBound(deltat_ch()), 2) & ";" & deltat_fr(1, 2) & "}", separateur_decimal, ".")
                    .Values = Replace("={" & deltat_ch(UBound(deltat_ch()), 1) + decal & ";" & deltat_ch(UBound(deltat_ch()), 1) + decal & "}", separateur_decimal, ".")
                    .Border.ColorIndex = 3
                    .Border.Weight = xlMedium
                    .MarkerStyle = xlNone
                    .Name = CodeNames("Feuil5").Range("l_besoinsChauds")
                End With
                If (low_x_ch - high_x_ch) > 0 Then
                    ActiveChart.Legend.LegendEntries(idlegend).Delete
                    idlegend = idlegend - 1
                End If
                ActiveChart.SeriesCollection.NewSeries
                idseries = idseries + 1
                idlegend = idlegend + 1
                With ActiveChart.SeriesCollection(idseries)
                    .XValues = Replace("={" & deltat_fr(1, 2) & ";" & deltat_fr(1, 2) & "}", separateur_decimal, ".")
                    .Values = Replace("={" & deltat_ch(UBound(deltat_ch()), 1) + decal & ";" & deltat_fr(1, 1) & "}", separateur_decimal, ".")
                    .Border.ColorIndex = 16
                    .MarkerStyle = xlNone
                    .Border.LineStyle = xlDot
                End With
                ActiveChart.Legend.LegendEntries(idlegend).Delete
                idlegend = idlegend - 1
            End If
        End If
        If cold_need > 0 Then
            If Abs(high_x_fr - low_x_fr) > 0 Then
                ActiveChart.SeriesCollection.NewSeries
                idseries = idseries + 1
                idlegend = idlegend + 1
                With ActiveChart.SeriesCollection(idseries)
                    .XValues = Replace("={" & low_x_fr & ";" & low_x_fr & "}", separateur_decimal, ".")
                    .Values = Replace("={" & low_y_fr & ";" & min_t_min - decal & "}", separateur_decimal, ".")
                    .Border.ColorIndex = 16
                    .MarkerStyle = xlNone
                    .Border.LineStyle = xlDot
                End With
                ActiveChart.Legend.LegendEntries(idlegend).Delete
                idlegend = idlegend - 1
                ActiveChart.SeriesCollection.NewSeries
                idseries = idseries + 1
                idlegend = idlegend + 1
                With ActiveChart.SeriesCollection(idseries)
                    .XValues = Replace("={" & low_x_fr & ";" & high_x_fr & "}", separateur_decimal, ".")
                    .Values = Replace("={" & min_t_min - decal & ";" & min_t_min - decal & "}", separateur_decimal, ".")
                    .Border.ColorIndex = 32
                    .Border.Weight = xlMedium
                    .MarkerStyle = xlNone
                    .Name = CodeNames("Feuil5").Range("l_besoinsFroids")
                End With
                ActiveChart.SeriesCollection.NewSeries
                idseries = idseries + 1
                idlegend = idlegend + 1
                With ActiveChart.SeriesCollection(idseries)
                    .XValues = Replace("={" & high_x_fr & ";" & high_x_fr & "}", separateur_decimal, ".")
                    .Values = Replace("={" & min_t_min - decal & ";" & high_y_fr & "}", separateur_decimal, ".")
                    .Border.ColorIndex = 16
                    .MarkerStyle = xlNone
                    .Border.LineStyle = xlDot
                End With
                ActiveChart.Legend.LegendEntries(idlegend).Delete
                idlegend = idlegend - 1
            End If
            If Abs(Abs(high_x_fr - low_x_fr) - cold_need) > 0.01 Then    'MODIF CRE 01/10/2012   Affichage des besoins froids à droite ET à gauche du pinch, dans le cas d'un recouvrement total de la CC des flux froids
                ActiveChart.SeriesCollection.NewSeries
                idseries = idseries + 1
                idlegend = idlegend + 1
                With ActiveChart.SeriesCollection(idseries)
                    .XValues = Replace("={" & deltat_fr(UBound(deltat_fr()), 2) & ";" & deltat_fr(UBound(deltat_fr()), 2) & "}", separateur_decimal, ".")
                    .Values = Replace("={" & deltat_fr(UBound(deltat_fr()), 1) & ";" & deltat_fr(UBound(deltat_fr()), 1) - decal & "}", separateur_decimal, ".")
                    .Border.ColorIndex = 16
                    .MarkerStyle = xlNone
                    .Border.LineStyle = xlDot
                End With
                ActiveChart.Legend.LegendEntries(idlegend).Delete
                idlegend = idlegend - 1
                ActiveChart.SeriesCollection.NewSeries
                idseries = idseries + 1
                idlegend = idlegend + 1
                With ActiveChart.SeriesCollection(idseries)
                    .XValues = Replace("={" & deltat_fr(UBound(deltat_fr()), 2) & ";" & deltat_ch(1, 2) & "}", separateur_decimal, ".")
                    .Values = Replace("={" & deltat_fr(UBound(deltat_fr()), 1) - decal & ";" & deltat_fr(UBound(deltat_fr()), 1) - decal & "}", separateur_decimal, ".")
                    .Border.ColorIndex = 32
                    .Border.Weight = xlMedium
                    .MarkerStyle = xlNone
                    .Name = CodeNames("Feuil5").Range("l_besoinsFroids")
                End With
                If (high_x_fr - low_x_fr) > 0 Then
                    ActiveChart.Legend.LegendEntries(idlegend).Delete
                    idlegend = idlegend - 1
                End If
                ActiveChart.SeriesCollection.NewSeries
                idseries = idseries + 1
                idlegend = idlegend + 1
                With ActiveChart.SeriesCollection(idseries)
                    .XValues = Replace("={" & deltat_ch(1, 2) & ";" & deltat_ch(1, 2) & "}", separateur_decimal, ".")
                    .Values = Replace("={" & deltat_fr(UBound(deltat_fr()), 1) - decal & ";" & deltat_ch(1, 1) & "}", separateur_decimal, ".")
                    .Border.ColorIndex = 16
                    .MarkerStyle = xlNone
                    .Border.LineStyle = xlDot
                End With
                ActiveChart.Legend.LegendEntries(idlegend).Delete
                idlegend = idlegend - 1
            End If
        End If
        Range("A1").Select



        'I.7. ---------------------------------   Construction du HEN (vide)  ----------------------------------
        position_erreur = "  -I.7"

        'Initialisation du HEN
        ReDim data_flux_hen(0 To nbflux, 1 To data_flux_hen_size)

        'Renseignement des noms des colonnes pour le tableau de flux du HEN
        data_flux_hen(0, 1) = "Type"            'type de flux : Ch ou Fr
        data_flux_hen(0, 2) = "Nom"             'nom du flux (pour au dessus du pinch)
        data_flux_hen(0, 3) = "tmax"            'tmax du flux
        data_flux_hen(0, 5) = "t_pinch_gch"     't_pinch_gch (température du flux au dessus du pinch)
        data_flux_hen(0, 6) = "t_pinch_cen"     't_pinch_cen (température du flux au pinch)
        data_flux_hen(0, 7) = "t_pinch_dte"     't_pinch_dte (température du flux au dessous du pinch)
        data_flux_hen(0, 9) = "tmin"            'tmin du flux
        data_flux_hen(0, 10) = "Nom"             'nom du flux (pour sous le pinch)
        data_flux_hen(0, 11) = "Dh"             'dh
        data_flux_hen(0, 12) = "MCp"            'mcp
        data_flux_hen(0, 13) = "p_gch"          'p gch
        data_flux_hen(0, 14) = "p_dte"          'p dte
        data_flux_hen(0, 15) = "reste_gch"      'reste gch
        data_flux_hen(0, 16) = "reste_dte"      'reste dte
        data_flux_hen(0, 17) = "match"          'match
        data_flux_hen(0, 18) = "order"          'order

        j = 1
        For I = 1 To nbflux
            'calcul des valeurs du HEN
            If data_flux(flux_process_Pw_kW, I) < 0 Then   'flux chaud
                tipe = "Ch"
                tmax = data_flux(flux_process_T_in_°C, I)
                tmin = data_flux(flux_process_T_out_°C, I)
                If point_pinch <> -1000 Then    'MODIF CRE 25/09/2012   ajout de ce test (et initialisation de point_pinch à -1000 pour éviter d'avoir un pinch quand on a un point à 0°C
                    t_pinch_gch = point_pinch + pinch_mesure / 2
                    t_pinch_dte = t_pinch_gch
                Else
                    t_pinch_gch = tmin
                End If
            Else
                tipe = "Fr"
                tmax = data_flux(flux_process_T_out_°C, I)
                tmin = data_flux(flux_process_T_in_°C, I)
                If point_pinch <> -1000 Then    'MODIF CRE 25/09/2012   ajout de ce test (et initialisation de point_pinch à -1000 pour éviter d'avoir un pinch quand on a un point à 0°C
                    t_pinch_gch = point_pinch - pinch_mesure / 2
                    t_pinch_dte = t_pinch_gch
                Else
                    t_pinch_gch = tmin
                End If
            End If
            nom_flux = data_flux(flux_process_nom, I)
            If (tmax <= t_pinch_dte) Then
                t_pinch_dte = tmax
                t_pinch_gch = -1000    'Valeur fantome car impossible de vider une variable Double. Sera supprimée par la suite
                tmax = -1000
                reste_gch = -1000
                If (data_flux(flux_process_T_in_°C, I) = data_flux(flux_process_T_out_°C, I)) Then
                    reste_dte = Abs(data_flux(flux_process_Pw_kW, I))
                Else
                    reste_dte = Abs((t_pinch_dte - tmin) / (t_pinch_dte - tmin) * data_flux(flux_process_Pw_kW, I))
                End If
            ElseIf (tmin >= t_pinch_gch) Then
                t_pinch_gch = tmin
                t_pinch_dte = -1000
                tmin = -1000
                reste_dte = -1000
                If (data_flux(flux_process_T_in_°C, I) = data_flux(flux_process_T_out_°C, I)) Then
                    reste_gch = Abs(data_flux(flux_process_Pw_kW, I))
                Else
                    reste_gch = Abs((tmax - t_pinch_gch) / (tmax - t_pinch_gch) * data_flux(flux_process_Pw_kW, I))
                End If
            Else
                reste_dte = Abs((t_pinch_dte - tmin) / (tmax - tmin) * data_flux(flux_process_Pw_kW, I))
                reste_gch = Abs((tmax - t_pinch_gch) / (tmax - tmin) * data_flux(flux_process_Pw_kW, I))
            End If

            'Ecriture des données dans une variable
            data_flux_hen(j, 1) = tipe                     'type de flux : Ch ou Fr
            data_flux_hen(j, 2) = nom_flux                 'nom du flux
            data_flux_hen(j, 3) = tmax                     'tmax du flux
            data_flux_hen(j, 5) = t_pinch_gch              't_pinch_gch (température du flux au dessus du pinch)
            data_flux_hen(j, 6) = t_pinch_gch              't_pinch_gch (température du flux au centre du pinch)
            data_flux_hen(j, 7) = t_pinch_dte              't_pinch_dte (température du flux au dessous du pinch)
            data_flux_hen(j, 9) = tmin                     'tmin du flux
            data_flux_hen(j, 10) = nom_flux                 'nom du flux (répétés aussi à la fin de la grille de flux)
            data_flux_hen(j, 11) = -CDbl(data_flux(5, I))  'dh (cf. données d'entrée)
            data_flux_hen(j, 12) = CDbl(data_flux(4, I))   'mcp (cf. données d'entrée)
            data_flux_hen(j, 13) = reste_gch               'p gch (puissance échangée par le flux à gauche du pinch)
            data_flux_hen(j, 14) = reste_dte               'p dte (puissance échangée par le flux à droite du pinch)
            data_flux_hen(j, 15) = reste_gch               'reste gch (puissance restante à échanger par le flux à gauche du pinch)
            data_flux_hen(j, 16) = reste_dte               'reste dte (puissance restante à échanger par le flux à gauche du pinch)
            data_flux_hen(j, 17) = 0                       'match
            data_flux_hen(j, 18) = j                       'order
            j = j + 1
        Next

        'Suppression des valeurs fantomes
        For I = 1 To nbflux
            For j = 1 To data_flux_hen_size
                If data_flux_hen(I, j) = "-1000" Then
                    data_flux_hen(I, j) = ""
                End If
            Next
        Next



        'I.8. ---------------------------------------   Finalization   -----------------------------------------
        position_erreur = "  -I.8"

        Call ecrire_VeriPinch_position_I8(nbflux_fr, deltat_fr(), hot_need, nbflux_ch, deltat_ch(), cold_need, recup_max)
        

        'I.9. -------------------------------------------   Fin   ----------------------------------------------
        position_erreur = "  -I.9"
        'MsgBox ("Tracé des Courbes Composites par VeriPinch terminé")

    End If



    If state <= 2 And state <> 0 Then
        'II. Construction du HEN
        'II.0. ---------------------------------   Initialisation générale   ------------------------------------
        'II.1. ---------------------------   Lecture des données (si nécessaire)   ------------------------------
        'II.2. ------------------------------   Remise à Zero du HEN existant    --------------------------------
        'II.3. ----------   Création des variables contenant les HEX à interdire et à supprimer   ---------------
        'II.4. --------------------------------    Ajout des HEX manuellement    --------------------------------
        'II.5. -------------------------------    Ajout des HEX automatiquement    ------------------------------
        'II.6. --------------------------------------    Finalization    ----------------------------------------
        'II.7. --------------------------------------------   Fin   ---------------------------------------------



        'II.0. ---------------------------------   Initialisation générale   ------------------------------------
        position_erreur = "  -II.0"

        nbhex = 0
        erreur = ""
        surftot = 0
        LigneTab = 4
        ok = True
        erreur = ""

        'Initialisation du tableau récapitulatif des données des échangeurs
        data_hen_size = 1
        ReDim data_hen(1 To 12, 1 To data_hen_size)
        data_hen(1, 1) = "HEX #"
        data_hen(2, 1) = "Puissance (kW)"
        data_hen(3, 1) = "Flux chaud"
        data_hen(4, 1) = "Flux Froid"
        data_hen(5, 1) = "Tin Ch (°C)"
        data_hen(6, 1) = "Tout Ch (°C)"
        data_hen(7, 1) = "Tin Fr (°C)"
        data_hen(8, 1) = "Tout Fr (°C)"
        data_hen(9, 1) = "Surface (m²)"
        data_hen(10, 1) = "erreur"
        data_hen(11, 1) = "colplacehexch"
        data_hen(12, 1) = "colplacehexfr"



        'II.1. -----------------------------------   Lecture des données  ---------------------------------------
        position_erreur = "  -II.1"

        'Lecture des données du HEN vierge depuis Excel si on n'a lancé le code qu'à partir du HEN
        If pinch_mesure > pinch_target Then    'MODIF CRE 21/09/2012   On utilisera la valeur du pinch souhaité par la suite, pas la valeur calculée dans le cas où les courbes n'ont pas été décalées
            'pinch_mesure = pinch_target
        End If



        'II.3. ----------   Création des variables contenant les HEX à interdire et à supprimer   ---------------
        position_erreur = "  -II.3"

        'Création de la variable des HEX à imposer
        colonneStation = 1 + (pi_FiliereConsideree - 1) * 9
        nbheximposed = 0
        I = 0
        Do While Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab + I, colonneStation) <> ""
            If Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab + I, colonneStation) = True Then
                nbheximposed = nbheximposed + 1
            End If
            I = I + 1
        Loop

        'Redimensionnement et remplissage de la variable
        If nbheximposed > 0 Then
            ReDim data_hex_imposed(1 To 8, 0 To nbheximposed)
            data_hex_imposed(1, 0) = "Nom Ch"
            data_hex_imposed(2, 0) = "Nom Fr"
            data_hex_imposed(3, 0) = "Tin Ch"
            data_hex_imposed(4, 0) = "Tout Ch"
            data_hex_imposed(5, 0) = "Tin Fr"
            data_hex_imposed(6, 0) = "Tout Fr"
            data_hex_imposed(7, 0) = "Puissance"
            data_hex_imposed(8, 0) = "Erreur"
            I = 0
            j = 1
            Do While Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab + I, colonneStation) <> ""
                If Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab + I, colonneStation) = True Then
                    For k = 1 To 7
                        data_hex_imposed(k, j) = Feuil63.Range("DonneesPincementImpositionEchange").Cells(LigneTab + I, colonneStation + k)
                        If data_hex_imposed(k, j) = "" Then
                            data_hex_imposed(k, j) = -1000    'valeur fantome utilisées dans VerifierData
                        End If
                    Next
                    j = j + 1
                End If
                I = I + 1
            Loop
        End If

        'Création de la variable des HEX à interdire
        colonneStation = 1 + (pi_FiliereConsideree - 1) * 3
        nbhexforbidden = 0
        I = 0
        Do While Feuil63.Range("DonneesPincementInterdictionEchange").Cells(LigneTab + I, colonneStation) <> ""
            If Feuil63.Range("DonneesPincementInterdictionEchange").Cells(LigneTab + I, colonneStation) = True Then
                nbhexforbidden = nbhexforbidden + 1
            End If
            I = I + 1
        Loop

        'Redimensionnement et remplissage de la variable
        If nbhexforbidden > 0 Then
            ReDim Preserve data_hex_forbidden(1 To 2, 0 To nbhexforbidden)
            data_hex_forbidden(1, 0) = "Nom Ch"
            data_hex_forbidden(2, 0) = "Nom Fr"

            I = 0
            j = 1
            Do While Feuil63.Range("DonneesPincementInterdictionEchange").Cells(LigneTab + I, colonneStation) <> ""
                If Feuil63.Range("DonneesPincementInterdictionEchange").Cells(LigneTab + I, colonneStation) = True Then
                    For k = 1 To 2
                        data_hex_forbidden(k, j) = Feuil63.Range("DonneesPincementInterdictionEchange").Cells(LigneTab + I, colonneStation + k)
                    Next
                    j = j + 1
                End If
                I = I + 1
            Loop
        End If



        'II.4. --------------------------------    Ajout des HEX manuellement    --------------------------------
        position_erreur = "  -II.4"

        Call GetCols(data_flux_hen())
        If nbheximposed > 0 Then
            For I = 1 To nbheximposed

                'Paramètres généraux du HEX
                ligfr = 0
                ligch = 0
                puissance = data_hex_imposed(7, I)
                tinch = data_hex_imposed(3, I)
                toutch = data_hex_imposed(4, I)
                tinfr = data_hex_imposed(5, I)
                toutfr = data_hex_imposed(6, I)
                For j = 1 To nbflux
                    If data_flux_hen(j, 1) = "Ch" And data_flux_hen(j, 2) = data_hex_imposed(1, I) Then
                        ligch = j
                    ElseIf data_flux_hen(j, 1) = "Fr" And data_flux_hen(j, 2) = data_hex_imposed(2, I) Then
                        ligfr = j
                    End If
                    If ligfr <> 0 And ligch <> 0 Then
                        j = nbflux
                    End If
                Next
                If ligch = 0 Or ligfr = 0 Then
                    ok = False
                    erreur = erreur & "Flux non trouvés"
                End If

                'Bornes du flux chaud
                If toutch < (point_pinch + pinch_mesure / 2) Then
                    'MODIF CRE 17/01/2013   Pour traiter les HEX à cheval sur le pinch
                    If tinch < (point_pinch + pinch_mesure / 2) Then
                        colres1 = 5
                    Else
                        colres1 = 45
                    End If
                Else
                    colres1 = 4
                End If
                If tinfr >= (point_pinch - pinch_mesure / 2) Then
                    colres2 = 4
                Else
                    'MODIF CRE 17/01/2013   Pour traiter les HEX à cheval sur le pinch
                    If toutfr >= (point_pinch - pinch_mesure / 2) Then
                        colres2 = 45
                    Else
                        colres2 = 5
                    End If
                End If
                
                'Vérification et création
                If ok = True Then
                    Call VerifierFluxSuppr(data_hex_imposed(1, I), data_hex_imposed(2, I))
                End If
                If ok = True Then
                    Call CheckData(data_flux_hen(), ligch, colres1, data_flux_hen(), ligfr, colres2, data_flux_hen_size, pinch_target, 1)
                End If
                If ok = True Then
                    nbhex = nbhex + 1
                    Call CreerHEX2(data_flux_hen(), ligch, colres1, colplacehexch, data_flux_hen(), ligfr, colres2, colplacehexfr, 1)
                End If

                ok = True
                erreur = ""
            Next
        End If

        'Ecriture du HEN avec manuels seulement 'MODIF CRE 16/10/2012   Ajout de cette écriture pour valider le boulot
        Call ecrire_VeriPinch_position_II4


        'II.5. -------------------------------    Ajout des HEX automatiquement    ------------------------------
        position_erreur = "  -II.5"

        Call GetCols(data_flux_hen())
        'Call Sort_Array_Custom(data_flux_hen(), "asc", UBound(data_flux_hen(), 2), colpinch + 2)    'Tri sur t_pinch_dte
        Call HENBuilding(data_flux_hen(), nbflux, 5, data_flux_hen(), nbflux, 5, pinch_target, data_flux_hen_size)
        'Call Sort_Array_Custom(data_flux_hen(), "asc", UBound(data_flux_hen(), 2), colreste + 7)  'retour au tri initial
        Call HENBuilding(data_flux_hen(), nbflux, 4, data_flux_hen(), nbflux, 4, pinch_target, data_flux_hen_size)



        'II.6. --------------------------------------    Finalization    ----------------------------------------
        position_erreur = "  -II.6"

        'Calcul et ajout des besoins restants
        Call GetCols(data_flux_hen())
        restegch_ch = 0
        restedte_ch = 0
        restegch_fr = 0
        restedte_fr = 0

        For I = 1 To nbflux
            If data_flux_hen(I, colreste + 4) <> "" Then    'flux présent à gauche avec un besoin restant
                If CDbl(data_flux_hen(I, colreste + 4)) > 0.1 Then
                    j = 1
                    tprev = data_flux_hen(I, 3)
                    Do While j < colpinch - 3
                        If data_flux_hen(I, 3 + j) = "" Then
                            k = 1
                            data_flux_hen(I, 3 + j) = repere_traits_Clement
                            Do While data_flux_hen(I, 3 + j + k) = ""
                                data_flux_hen(I, 3 + j + k) = repere_traits_Clement
                                k = k + 1
                            Loop
                            tnext = data_flux_hen(I, 3 + j + k)
                            puissance = (tprev - tnext) * data_flux_hen(I, colreste + 1)
                            If data_flux_hen(I, 1) = "Ch" Then
                                restegch_fr = restegch_fr + puissance
                                label = "Col_"
                            Else
                                restegch_ch = restegch_ch + puissance
                                label = "Hot_"
                            End If
                            data_flux_hen(I, 3 + j) = label & puissance
                            data_flux_hen(I, 3 + j + 1) = tnext
                            j = j + k
                        ElseIf left(data_flux_hen(I, 3 + j), 2) <> "P_" And data_flux_hen(I, 3 + j) <> repere_traits_Clement Then
                            tprev = data_flux_hen(I, 3 + j)
                        End If
                        j = j + 1
                    Loop
                End If
            End If
            If data_flux_hen(I, colreste + 5) <> "" Then    'flux présent à droite avec un besoin restant
                If CDbl(data_flux_hen(I, colreste + 5)) > 0.1 Then
                    j = 1
                    tprev = data_flux_hen(I, colpinch + 2)
                    Do While j < (colreste - 2) - (colpinch + 2)
                        If data_flux_hen(I, colpinch + 2 + j) = "" Then
                            k = 1
                            data_flux_hen(I, colpinch + 2 + j) = repere_traits_Clement
                            Do While data_flux_hen(I, colpinch + 2 + j + k) = ""
                                data_flux_hen(I, colpinch + 2 + j + k) = repere_traits_Clement
                                k = k + 1
                            Loop
                            tnext = data_flux_hen(I, colpinch + 2 + j + k)
                            puissance = (tprev - tnext) * data_flux_hen(I, colreste + 1)
                            If data_flux_hen(I, 1) = "Ch" Then
                                restedte_fr = restedte_fr + puissance
                                label = "Col_"
                            Else
                                restedte_ch = restedte_ch + puissance
                                label = "Hot_"
                            End If
                            data_flux_hen(I, colpinch + 2 + j) = label & puissance
                            data_flux_hen(I, colpinch + 2 + j + 1) = tnext
                            j = j + k
                            If data_flux_hen(I, colreste + 5) < 0.000001 Then
                                j = (colreste - 2) - (colpinch + 2)
                            End If
                        ElseIf left(data_flux_hen(I, colpinch + 2 + j), 2) <> "P_" And data_flux_hen(I, colpinch + 2 + j) <> repere_traits_Clement Then
                            tprev = data_flux_hen(I, colpinch + 2 + j)
                        End If
                        j = j + 1
                    Loop
                End If
            End If
        Next

        'Légère adaptation des températures 'MODIF CRE 26/09/2012   Nettoyage léger du HEN pour éviter la présence de certains flux à droite ou gauche du pinch si tmin ou tmax = tpinch
        For I = 1 To nbflux
            If data_flux_hen(I, 3) <> "" Then
                If data_flux_hen(I, 3) = data_flux_hen(I, colpinch) Then
                    For j = 3 To colpinch - 1
                        data_flux_hen(I, j) = ""
                    Next
                End If
            End If
            If data_flux_hen(I, colpinch + 2) <> "" Then
                If data_flux_hen(I, colpinch + 2) = data_flux_hen(I, colreste - 2) Then
                    For j = colpinch + 2 To colreste - 2
                        data_flux_hen(I, j) = ""
                    Next
                End If
            End If
        Next

        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '**********************************************************************************************
        'Ajout DCA pour récupérer les besoins restants
        '**********************************************************************************************
        'Réinitialisation du tableau avec les échanges avec les utilités
        besoins_HOT = restegch_ch + restedte_ch
        besoins_COLD = restegch_fr + restedte_fr
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


        'Pour conserver en mémoire les échangeurs de flux procédés
        nb_echangeurs_process = data_hen_size
        nbhex_stocke = nbhex
        ReDim data_flux_hen_stocke(0 To nbflux, 1 To data_flux_hen_size)
        data_flux_hen_size_stockee = data_flux_hen_size
        For I = 0 To nbflux
            For j = 1 To data_flux_hen_size
                data_flux_hen_stocke(I, j) = data_flux_hen(I, j)
            Next j
        Next I

        'MODIF CRE 10/10/2012   Ajout pour l'affichage final des résultats
        Call ecrire_VeriPinch_position_II6(nbflux_ch, nbflux_fr, restegch_ch, restedte_ch, restegch_fr, restedte_fr)
        


        'II.7. --------------------------------------------   Fin   ---------------------------------------------
        position_erreur = "  -II.7"
        'Msgbox ("HEX Construction Done.")

    End If



    If state <= 3 And state <> 0 Then
        'III. Gestion des Utilités
        'III.0. -----------------------------------   Initialisations générales   --------------------------------
        'III.1. -----------------------------   Lecture des données (si nécessaire)   ----------------------------
        'III.2. -----------------------------------------   Remises à Zero   -------------------------------------
        'III.3. ---------------------------   Renseignement de la liste des Utilités   ---------------------------
        'III.4. --------------------------------   Liste des besoins restants   ----------------------------------
        'III.5. -----------------------------------   Couverture des besoins   -----------------------------------
        'III.6. -----------------------------------------   Nettoyage   ------------------------------------------
        'III.7. ---------------------------------------   Affinage HEN   -----------------------------------------
        'III.8. ----------------------------------------   Export HEN   ------------------------------------------
        'III.9. ---------------------------------------   Finalization   -----------------------------------------
        'III.10. ------------------------------------------   Fin.   ---------------------------------------------



        'III.0. -----------------------------------   Initialisations générales   --------------------------------
        position_erreur = "  -III.0"

        If PINCH_DANS_OCEAN = False Then
            Call lecture_flux_utilites
        End If

        If PINCH_ORDONNER_FLUX_UTILITES = True Then
            Call ordonner_flux_utilites
        End If

        LigneTab = 4
        nbneeds = 0
        ok = True
        erreur = ""

        par = 10
        nbneeds_gch_ch = 0
        nbneeds_gch_fr = 0
        nbneeds_dte_ch = 0
        nbneeds_dte_fr = 0
        ReDim data_needs_gch_ch(1 To par, 0 To 1) As String
        ReDim data_needs_gch_fr(1 To par, 0 To 1) As String
        ReDim data_needs_dte_ch(1 To par, 0 To 1) As String
        ReDim data_needs_dte_fr(1 To par, 0 To 1) As String

        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        '**********************************************************************************************
        'Ajout DCA à cause du problème sur les échangeurs avec utilités qui s'accumulent quand on lance VeriPinch(3) plusieurs fois
        '**********************************************************************************************
        'Réinitialisation du tableau avec les échanges avec les utilités
        data_hen_size = nb_echangeurs_process
        nbhex = nbhex_stocke
        ReDim Preserve data_hen(1 To 12, 1 To data_hen_size)
        ReDim data_flux_hen(0 To nbflux, 1 To data_flux_hen_size_stockee)
        For I = 0 To nbflux
            For j = 1 To data_flux_hen_size_stockee
                data_flux_hen(I, j) = data_flux_hen_stocke(I, j)
            Next j
        Next I
        '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



        'III.1. -----------------------------   Lecture des données (si nécessaire)   ----------------------------
        position_erreur = "  -III.1"

        Call lecture_VeriPinch_position_III1(state)


        'III.2. -----------------------------------------   Remises à Zero   -------------------------------------
        position_erreur = "  -III.2"

        Call effacer_VeriPinch_position_III2(state)


        'III.3. ---------------------------   Renseignement de la liste des Utilités   ---------------------------
        position_erreur = "  -III.3"

        'Détermination du nb d'utilités
        nbutilities_updated = 0
        nbutilities = 0
        ReDim data_utilities(1 To 7, 1 To 1)
        For I = 1 To UBound(utilites_flux_energie, 2)
            If utilites_flux_energie(repere_flux_thermique_utilites_etat, I) = True Then
                nbutilities = nbutilities + 1
                ReDim Preserve data_utilities(1 To flux_utilites_nb_parametres, 1 To nbutilities)
                data_utilities(flux_utilites_nom, nbutilities) = utilites_flux_energie(repere_flux_thermique_utilites_nom, I)
                data_utilities(flux_utilites_type, nbutilities) = "U" & utilites_flux_energie(repere_flux_thermique_utilites_type, I)
                data_utilities(flux_utilites_T_in_°C, nbutilities) = utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, I)
                data_utilities(flux_utilites_MCp_kW_°C, nbutilities) = utilites_flux_energie(repere_flux_thermique_utilites_MCp_kW_°C, I)
                data_utilities(flux_utilites_Pw_kW, nbutilities) = utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, I)
                data_utilities(6, nbutilities) = data_utilities(flux_utilites_T_in_°C, nbutilities)
                data_utilities(7, nbutilities) = data_utilities(flux_utilites_Pw_kW, nbutilities)
            End If
        Next I

        'Création des flux d'utilités
        If nbutilities > 0 Then
            ReDim data_flux_utilities(1 To nbutilities, 1 To data_flux_hen_size)

            Call GetCols(data_flux_hen())
            For I = 1 To nbutilities
                data_flux_utilities(I, 1) = data_utilities(2, I)                'tipe
                data_flux_utilities(I, 2) = data_utilities(1, I)                'nom
                data_flux_utilities(I, colreste - 1) = data_utilities(1, I)     'nom
                data_flux_utilities(I, colreste) = data_utilities(5, I)         'puissance totale
                data_flux_utilities(I, colreste + 1) = data_utilities(4, I)     'mcp
                If data_utilities(2, I) = "UCh" Then
                    data_flux_utilities(I, 3) = data_utilities(3, I)
                    data_flux_utilities(I, colreste + 4) = data_utilities(5, I)
                    data_flux_utilities(I, colreste - 2) = (data_utilities(3, I) - data_utilities(5, I) / data_utilities(4, I))
                Else
                    data_flux_utilities(I, colreste - 2) = data_utilities(3, I)
                    data_flux_utilities(I, colreste + 5) = data_utilities(5, I)
                    data_flux_utilities(I, 3) = data_utilities(3, I) + data_utilities(5, I) / data_utilities(4, I)
                End If
                data_flux_utilities(I, colreste + 2) = data_flux_utilities(I, colreste + 4)    'pgch
                data_flux_utilities(I, colreste + 3) = data_flux_utilities(I, colreste + 5)    'pdte

                data_flux_utilities(I, colreste + 7) = nbflux + I               'order
            Next
'For i = LBound(data_flux_utilities(), 1) To UBound(data_flux_utilities(), 1)
'    For j = LBound(data_flux_utilities(), 2) To UBound(data_flux_utilities(), 2)
'        Feuil63.Cells(50 + UBound(data_flux_hen(), 1) + i, j).Value = data_flux_utilities(i, j)
'    Next j
'Next i
'
        End If



        'III.4. --------------------------------   Liste des besoins restants   ----------------------------------
        position_erreur = "  -III.4"
        Call GetCols(data_flux_hen())
        besoins_HOT = 0
        besoins_COLD = 0
        For I = 1 To nbflux
            For j = 4 To colreste - 3

                If left(data_flux_hen(I, j), 4) = "Hot_" Or left(data_flux_hen(I, j), 4) = "Col_" Then

                    nbneeds = nbneeds + 1
                    If left(data_flux_hen(I, j), 4) = "Hot_" Then
                        puissance = CDbl(Replace(Right(data_flux_hen(I, j), Len(data_flux_hen(I, j)) - 4), ".", separateur_decimal))
                        tipe = "Ch"
                        If j > colpinch Then
                            nbneeds_dte_ch = nbneeds_dte_ch + 1
                        Else
                            nbneeds_gch_ch = nbneeds_gch_ch + 1
                        End If
                        besoins_HOT = besoins_HOT + puissance
                    Else
                        puissance = CDbl(Replace(Right(data_flux_hen(I, j), Len(data_flux_hen(I, j)) - 4), ".", separateur_decimal))
                        tipe = "Fr"
                        If j > colpinch Then
                            nbneeds_dte_fr = nbneeds_dte_fr + 1
                        Else
                            nbneeds_gch_fr = nbneeds_gch_fr + 1
                        End If
                        besoins_COLD = besoins_COLD + puissance
                    End If

                    'Sauvegarde d'un résumé des besoins dans une variable
                    If tipe = "Ch" Then
                        If j > colpinch Then
                            ReDim Preserve data_needs_dte_ch(1 To par, 0 To nbneeds_dte_ch)
                            data_needs_dte_ch(1, nbneeds_dte_ch) = data_flux_hen(I, 2)
                            data_needs_dte_ch(2, nbneeds_dte_ch) = data_flux_hen(I, 1)
                            data_needs_dte_ch(3, nbneeds_dte_ch) = data_flux_hen(I, j + 1)
                            data_needs_dte_ch(4, nbneeds_dte_ch) = puissance
                            data_needs_dte_ch(5, nbneeds_dte_ch) = data_flux_hen(I, j - 1)
                            data_needs_dte_ch(6, nbneeds_dte_ch) = data_flux_hen(I, colreste + 1)
                            data_needs_dte_ch(7, nbneeds_dte_ch) = puissance
                            data_needs_dte_ch(8, nbneeds_dte_ch) = ""
                            data_needs_dte_ch(9, nbneeds_dte_ch) = I
                            data_needs_dte_ch(10, nbneeds_dte_ch) = j

                        Else
                            ReDim Preserve data_needs_gch_ch(1 To par, 0 To nbneeds_gch_ch)
                            data_needs_gch_ch(1, nbneeds_gch_ch) = data_flux_hen(I, 2)
                            data_needs_gch_ch(2, nbneeds_gch_ch) = data_flux_hen(I, 1)
                            data_needs_gch_ch(3, nbneeds_gch_ch) = data_flux_hen(I, j + 1)
                            data_needs_gch_ch(4, nbneeds_gch_ch) = puissance
                            data_needs_gch_ch(5, nbneeds_gch_ch) = data_flux_hen(I, j - 1)
                            data_needs_gch_ch(6, nbneeds_gch_ch) = data_flux_hen(I, colreste + 1)
                            data_needs_gch_ch(7, nbneeds_gch_ch) = puissance
                            data_needs_gch_ch(8, nbneeds_gch_ch) = ""
                            data_needs_gch_ch(9, nbneeds_gch_ch) = I
                            data_needs_gch_ch(10, nbneeds_gch_ch) = j
                        End If
                    Else
                        If j > colpinch Then
                            ReDim Preserve data_needs_dte_fr(1 To par, 0 To nbneeds_dte_fr)
                            data_needs_dte_fr(1, nbneeds_dte_fr) = data_flux_hen(I, 2)
                            data_needs_dte_fr(2, nbneeds_dte_fr) = data_flux_hen(I, 1)
                            data_needs_dte_fr(3, nbneeds_dte_fr) = data_flux_hen(I, j - 1)
                            data_needs_dte_fr(4, nbneeds_dte_fr) = puissance
                            data_needs_dte_fr(5, nbneeds_dte_fr) = data_flux_hen(I, j + 1)
                            data_needs_dte_fr(6, nbneeds_dte_fr) = data_flux_hen(I, colreste + 1)
                            data_needs_dte_fr(7, nbneeds_dte_fr) = puissance
                            data_needs_dte_fr(8, nbneeds_dte_fr) = ""
                            data_needs_dte_fr(9, nbneeds_dte_fr) = I
                            data_needs_dte_fr(10, nbneeds_dte_fr) = j
                        Else
                            ReDim Preserve data_needs_gch_fr(1 To par, 0 To nbneeds_gch_fr)
                            data_needs_gch_fr(1, nbneeds_gch_fr) = data_flux_hen(I, 2)
                            data_needs_gch_fr(2, nbneeds_gch_fr) = data_flux_hen(I, 1)
                            data_needs_gch_fr(3, nbneeds_gch_fr) = data_flux_hen(I, j - 1)
                            data_needs_gch_fr(4, nbneeds_gch_fr) = puissance
                            data_needs_gch_fr(5, nbneeds_gch_fr) = data_flux_hen(I, j + 1)
                            data_needs_gch_fr(6, nbneeds_gch_fr) = data_flux_hen(I, colreste + 1)
                            data_needs_gch_fr(7, nbneeds_gch_fr) = puissance
                            data_needs_gch_fr(8, nbneeds_gch_fr) = ""
                            data_needs_gch_fr(9, nbneeds_gch_fr) = I
                            data_needs_gch_fr(10, nbneeds_gch_fr) = j
                        End If
                    End If

                    'Suppression des informations du besoin identifié sur la grille de flux
                    data_flux_hen(I, j) = ""
                    k = j
                    Do While data_flux_hen(I, k - 2) = repere_traits_Clement
                        data_flux_hen(I, k + 2) = ""
                        k = k + 1
                    Loop
                    If k <> j Then
                        data_flux_hen(I, j - 1) = ""
                    End If
                    k = j
                    Do While data_flux_hen(I, k + 2) = repere_traits_Clement
                        data_flux_hen(I, k + 2) = ""
                        k = k + 1
                    Loop
                    If k <> j Then
                        data_flux_hen(I, j + 1) = ""
                    End If

                End If
            Next j
        Next I

        data_needs_dte_ch(1, 0) = "nom"
        data_needs_dte_ch(2, 0) = "tipe de flux"
        data_needs_dte_ch(3, 0) = "tin"
        data_needs_dte_ch(4, 0) = "puissance"
        data_needs_dte_ch(5, 0) = "tout"
        data_needs_dte_ch(6, 0) = "mcp"
        data_needs_dte_ch(7, 0) = "p restante"
        data_needs_dte_ch(8, 0) = "erreur"

        data_needs_dte_fr(1, 0) = "nom"
        data_needs_dte_fr(2, 0) = "tipe de flux"
        data_needs_dte_fr(3, 0) = "tin"
        data_needs_dte_fr(4, 0) = "puissance"
        data_needs_dte_fr(5, 0) = "tout"
        data_needs_dte_fr(6, 0) = "mcp"
        data_needs_dte_fr(7, 0) = "p restante"
        data_needs_dte_fr(8, 0) = "erreur"

        data_needs_gch_fr(1, 0) = "nom"
        data_needs_gch_fr(2, 0) = "tipe de flux"
        data_needs_gch_fr(3, 0) = "tin"
        data_needs_gch_fr(4, 0) = "puissance"
        data_needs_gch_fr(5, 0) = "tout"
        data_needs_gch_fr(6, 0) = "mcp"
        data_needs_gch_fr(7, 0) = "p restante"
        data_needs_gch_fr(8, 0) = "erreur"

        data_needs_gch_ch(1, 0) = "nom"
        data_needs_gch_ch(2, 0) = "tipe de flux"
        data_needs_gch_ch(3, 0) = "tin"
        data_needs_gch_ch(4, 0) = "puissance"
        data_needs_gch_ch(5, 0) = "tout"
        data_needs_gch_ch(6, 0) = "mcp"
        data_needs_gch_ch(7, 0) = "p restante"
        data_needs_gch_ch(8, 0) = "erreur"

        'Transposition des tableaux de besoins
        ReDim data_needs_gch_ch_2(0 To nbneeds_gch_ch, 1 To par)
        ReDim data_needs_gch_fr_2(0 To nbneeds_gch_fr, 1 To par)
        ReDim data_needs_dte_ch_2(0 To nbneeds_dte_ch, 1 To par)
        ReDim data_needs_dte_fr_2(0 To nbneeds_dte_fr, 1 To par)
        For I = 1 To par
            For j = 0 To nbneeds_gch_ch
                data_needs_gch_ch_2(j, I) = data_needs_gch_ch(I, j)
            Next
        Next
        For I = 1 To par
            For j = 0 To nbneeds_gch_fr
                data_needs_gch_fr_2(j, I) = data_needs_gch_fr(I, j)
            Next
        Next
        For I = 1 To par
            For j = 0 To nbneeds_dte_ch
                data_needs_dte_ch_2(j, I) = data_needs_dte_ch(I, j)
            Next
        Next
        For I = 1 To par
            For j = 0 To nbneeds_dte_fr
                data_needs_dte_fr_2(j, I) = data_needs_dte_fr(I, j)
            Next
        Next



        'III.5. -----------------------------------   Couverture des besoins   -----------------------------------
        position_erreur = "  -III.5-ch"

        'Couverture de tous les besoins chauds
        besoins_HOT = 0
        If nbneeds_gch_ch > 0 Or nbneeds_dte_ch > 0 Then
            ReDim data_needs_ch(0 To nbneeds_gch_ch + nbneeds_dte_ch, 1 To par)
            I = 0
            For I = 1 To nbneeds_gch_ch
                For j = 1 To par
                    data_needs_ch(I, j) = data_needs_gch_ch_2(I, j)
                Next
            Next
            temp = I - 1
            For I = 1 To nbneeds_dte_ch
                For j = 1 To par
                    data_needs_ch(temp + I, j) = data_needs_dte_ch_2(I, j)
                Next
            Next

            Call Sort_Array_Custom(data_needs_ch, "dsc", par, 5)
            For I = 1 To nbneeds_gch_ch + nbneeds_dte_ch
                iteration = 1
                erreur = ""
                puissance = data_needs_ch(I, 4)
                besoins_HOT = besoins_HOT + puissance
                Do While data_needs_ch(I, 4) > 0.001 And iteration < 10
                    For j = 1 To nbutilities
'For k = LBound(data_flux_hen(), 1) To UBound(data_flux_hen(), 1)
'    For m = LBound(data_flux_hen(), 2) To UBound(data_flux_hen(), 2)
'        Feuil63.Cells(50 + k, m).Value = data_flux_hen(k, m)
'    Next m
'Next k
'For k = LBound(data_flux_utilities(), 1) To UBound(data_flux_utilities(), 1)
'    For m = LBound(data_flux_utilities(), 2) To UBound(data_flux_utilities(), 2)
'        Feuil63.Cells(50 + UBound(data_flux_hen(), 1) + k, m).Value = data_flux_utilities(k, m)
'    Next m
'Next k

                        If data_flux_utilities(j, 1) = "UCh" Then

                            If data_utilities(7, j) <> "" Then
                                If data_utilities(7, j) > 0.001 Then

                                    puissance = data_needs_ch(I, 4)    'MODIF CRE 25/09/2012 ajout de cette mise à jour qui manquait
                                    If puissance > data_utilities(7, j) Or puissance < 0.000001 Then
                                        'ok = False 'MODIF CRE 24/09/2012   si la puissance disponible dans l'utilité est inférieure au besoin, on découpe le besoin
                                        k = 1
                                    End If

                                    If ok = True Then
                                        tinfr = CDbl(data_needs_ch(I, 3))
                                        toutfr = CDbl(data_needs_ch(I, 5))
                                        tinch = data_utilities(6, j)
                                        toutch = data_utilities(6, j) - puissance / data_utilities(4, j)
                                        't_pinch_dte = point_pinch - pinch_mesure / 2   'MODIF CRE 25/09/2012   Si on prends la valeur de pinch_mesure, et que les courbes n'ont pas été décalées, on n'aura pas la bonne valeur du pinch
                                        If data_flux_hen(data_needs_ch(I, 9), colpinch) <> "" Then
                                            t_pinch_dte = data_flux_hen(data_needs_ch(I, 9), colpinch)
                                        Else
                                            t_pinch_dte = data_flux_hen(data_needs_ch(I, 9), colpinch + 2)
                                        End If
                                        If tinfr < t_pinch_dte - 0.000001 Then
                                            b1 = 5
                                        Else
                                            b1 = 4
                                        End If
                                        b2 = 10
                                        a1 = data_needs_ch(I, 9)
                                        a2 = j
                                    End If

                                    If ok = True Then
                                        If data_flux_utilities(j, 1) = "UFr" Then
                                            Call VerifierFluxSuppr(data_flux_hen(a1, 2), data_flux_utilities(a2, 2))
                                        Else
                                            Call VerifierFluxSuppr(data_flux_utilities(a2, 2), data_flux_hen(a1, 2))
                                        End If
                                    End If
                                    If ok = True Then
                                        If data_flux_utilities(j, 1) = "UFr" Then
                                            Call CheckData(data_flux_hen(), a1, b1, data_flux_utilities(), a2, b2, UBound(data_flux_hen(), 2), pinch_target, 3)
                                        Else
                                            Call CheckData(data_flux_utilities(), a2, b2, data_flux_hen(), a1, b1, UBound(data_flux_hen(), 2), pinch_target, 3)
                                        End If
                                    End If
'For k = LBound(data_flux_hen(), 1) To UBound(data_flux_hen(), 1)
'    For m = LBound(data_flux_hen(), 2) To UBound(data_flux_hen(), 2)
'        Feuil63.Cells(60 + k, m).Value = data_flux_hen(k, m)
'    Next m
'Next k
'For k = LBound(data_flux_utilities(), 1) To UBound(data_flux_utilities(), 1)
'    For m = LBound(data_flux_utilities(), 2) To UBound(data_flux_utilities(), 2)
'        Feuil63.Cells(60 + UBound(data_flux_hen(), 1) + k, m).Value = data_flux_utilities(k, m)
'    Next m
'Next k
                                    If ok = True Then
                                        nbhex = nbhex + 1
                                        If data_flux_utilities(j, 1) = "UFr" Then
                                            Call CreerHEX2(data_flux_hen(), a1, b1, colplacehexfr, data_flux_utilities(), a2, b2, colplacehexch, 3)
                                        Else
                                            Call CreerHEX2(data_flux_utilities(), a2, b2, colplacehexch, data_flux_hen(), a1, b1, colplacehexfr, 3)
                                        End If
                                    End If
'For k = LBound(data_flux_hen(), 1) To UBound(data_flux_hen(), 1)
'    For m = LBound(data_flux_hen(), 2) To UBound(data_flux_hen(), 2)
'        Feuil63.Cells(70 + k, m).Value = data_flux_hen(k, m)
'    Next m
'Next k
'For k = LBound(data_flux_utilities(), 1) To UBound(data_flux_utilities(), 1)
'    For m = LBound(data_flux_utilities(), 2) To UBound(data_flux_utilities(), 2)
'        Feuil63.Cells(70 + UBound(data_flux_hen(), 1) + k, m).Value = data_flux_utilities(k, m)
'    Next m
'Next k
                                    If ok = True Then
                                        data_needs_ch(I, 4) = data_needs_ch(I, 4) - puissance
                                        data_utilities(6, j) = data_utilities(6, j) - puissance / data_utilities(4, j)
                                        If data_needs_ch(I, 4) > 0.000001 Then
                                            If data_needs_ch(I, 3) = CStr(tinfr) Then
                                                data_needs_ch(I, 3) = toutfr
                                            ElseIf data_needs_ch(I, 5) = CStr(toutfr) Then
                                                data_needs_ch(I, 5) = tinfr
                                            End If
                                        End If
                                        data_utilities(7, j) = data_utilities(7, j) - puissance
                                        besoins_HOT = besoins_HOT - puissance

                                    End If
                                    ok = True
                                    erreur = ""
                                End If
                            End If
                        End If
                        If data_needs_ch(I, 4) < 0.001 Then
                            j = nbutilities
                        End If
                    Next
                    iteration = iteration + 1
                Loop
                If iteration = 10 Then
                    erreur = erreur & "Utility not found. "
                End If
                data_needs_ch(I, 8) = erreur
            Next
        End If

'For i = LBound(data_flux_hen(), 1) To UBound(data_flux_hen(), 1)
'    For j = LBound(data_flux_hen(), 2) To UBound(data_flux_hen(), 2)
'        Feuil63.Cells(50 + i, j).Value = data_flux_hen(i, j)
'    Next j
'Next i
'If nbutilities > 0 Then
'    For i = LBound(data_flux_utilities(), 1) To UBound(data_flux_utilities(), 1)
'        For j = LBound(data_flux_utilities(), 2) To UBound(data_flux_utilities(), 2)
'            Feuil63.Cells(50 + UBound(data_flux_hen(), 1) + i, j).Value = data_flux_utilities(i, j)
'        Next j
'    Next i
'End If

        position_erreur = "  -III.5-fr"
        'Couverture de tous les besoins froids
        besoins_COLD = 0
        If nbneeds_gch_fr > 0 Or nbneeds_dte_fr > 0 Then
            ReDim data_needs_fr(0 To nbneeds_gch_fr + nbneeds_dte_fr, 1 To par)
            I = 0
            For I = 1 To nbneeds_dte_fr
                For j = 1 To par
                    data_needs_fr(I, j) = data_needs_dte_fr_2(I, j)
                Next
            Next
            temp = I - 1
            For I = 1 To nbneeds_gch_fr
                For j = 1 To par
                    data_needs_fr(temp + I, j) = data_needs_gch_fr_2(I, j)
                Next
            Next
            Call Sort_Array_Custom(data_needs_fr, "asc", par, 5)
            For I = 1 To nbneeds_gch_fr + nbneeds_dte_fr

                iteration = 1
                puissance = data_needs_fr(I, 4)
                erreur = ""
                besoins_COLD = besoins_COLD + puissance

                Do While data_needs_fr(I, 4) > 0.001 And iteration < 10
                    For j = 1 To nbutilities

                        If data_flux_utilities(j, 1) = "UFr" Then

                            If data_utilities(7, j) <> "" Then
                                If data_utilities(7, j) > 0.001 Then

                                    puissance = data_needs_fr(I, 4)    'MODIF CRE 25/09/2012 ajout de cette mise à jour qui manquait
                                    If puissance > data_utilities(7, j) Or puissance < 0.000001 Then
                                        'ok = False 'MODIF CRE 24/09/2012   si la puissance disponible dans l'utilité est inférieure au besoin, on découpe le besoin
                                        k = 1
                                    End If

                                    If ok = True Then
                                        tinch = CDbl(data_needs_fr(I, 3))
                                        toutch = CDbl(data_needs_fr(I, 5))
                                        tinfr = data_utilities(6, j)
                                        toutfr = data_utilities(6, j) + puissance / data_utilities(4, j)
                                        't_pinch_dte = point_pinch + pinch_mesure / 2   'MODIF CRE 25/09/2012   Si on prends la valeur de pinch_mesure, et que les courbes n'ont pas été décalées, on n'aura pas la bonne valeur du pinch
                                        If data_flux_hen(data_needs_fr(I, 9), colpinch) <> "" Then
                                            t_pinch_dte = data_flux_hen(data_needs_fr(I, 9), colpinch)
                                        Else
                                            t_pinch_dte = data_flux_hen(data_needs_fr(I, 9), colpinch + 2)
                                        End If

                                        If (tinch - 0.000001 <= t_pinch_dte) Then
                                            b1 = 5
                                        Else
                                            b1 = 4
                                        End If
                                        b2 = 10
                                        a1 = data_needs_fr(I, 9)
                                        a2 = j
                                    End If

                                    If ok = True Then
                                        If data_flux_utilities(j, 1) = "UFr" Then
                                            'Call VerifierFluxSuppr(data_flux_hen(i, 2), data_flux_utilities(j, 2))
                                        Else
                                            'Call VerifierFluxSuppr(data_flux_utilities(j, 2), data_flux_hen(i, 2))
                                        End If
                                    End If
                                    If ok = True Then
                                        If data_flux_utilities(j, 1) = "UFr" Then
                                            Call CheckData(data_flux_hen(), a1, b1, data_flux_utilities(), a2, b2, UBound(data_flux_hen(), 2), pinch_target, 3)
                                        Else
                                            Call CheckData(data_flux_utilities(), a2, b2, data_flux_hen(), a1, b1, UBound(data_flux_hen(), 2), pinch_target, 3)
                                        End If
                                    End If
                                    If ok = True Then
                                        nbhex = nbhex + 1
                                        If data_flux_utilities(j, 1) = "UFr" Then
                                            Call CreerHEX2(data_flux_hen(), a1, b1, colplacehexch, data_flux_utilities(), a2, b2, colplacehexfr, 3)
                                        Else
                                            Call CreerHEX2(data_flux_utilities(), a2, b2, colplacehexfr, data_flux_hen(), a1, b1, colplacehexch, 3)
                                        End If
                                    End If
                                    If ok = True Then
                                        data_needs_fr(I, 4) = data_needs_fr(I, 4) - puissance
                                        data_utilities(6, j) = data_utilities(6, j) + puissance / data_utilities(4, j)
                                        If data_needs_fr(I, 4) > 0.000001 Then
                                            If data_needs_fr(I, 3) = tinch Then
                                                data_needs_fr(I, 3) = toutch
                                            ElseIf data_needs_fr(I, 5) = toutch Then
                                                data_needs_fr(I, 5) = tinch
                                            End If
                                        End If
                                        data_utilities(7, j) = data_utilities(7, j) - puissance
                                        besoins_COLD = besoins_COLD - puissance
                                    End If
                                    ok = True
                                    erreur = ""
                                End If
                            End If
                        End If
                        If data_needs_fr(I, 4) < 0.000001 Then
                            j = nbutilities
                        End If
                    Next
                    iteration = iteration + 1
                Loop
                If iteration = 10 Then
                    erreur = erreur & "Utility not found. "
                End If
                data_needs_fr(I, 8) = erreur
            Next
        End If



        'III.6. ----------------------------------------   Nettoyage   -------------------------------------------
        position_erreur = "  -III.6"

        'Création d'un array contenant la liste des flux d'utilités
        For I = 1 To nbutilities
            nbutilities_updated = nbutilities_updated + 1
        Next

        'Copie dans le nouveau tableau ("_updated")
        If nbutilities_updated > 0 Then
            ReDim data_flux_utilities_updated(1 To nbutilities_updated, 1 To UBound(data_flux_hen(), 2))
            k = 0
            For I = 1 To nbutilities
                k = k + 1
                For j = 1 To UBound(data_flux_hen(), 2)
                    data_flux_utilities_updated(k, j) = data_flux_utilities(I, j)
                Next
            Next
        End If

        'Réajout des besoins
        Call GetCols(data_flux_hen())
        For I = 1 To nbflux
            If data_flux_hen(I, 3) <> "" Then
                For j = 4 To colpinch - 1
                    If data_flux_hen(I, j - 1) <> "" Then
                        If data_flux_hen(I, j) = "" Then
                            If data_flux_hen(I, 1) = "Ch" Then
                                label = "Col_"
                            Else
                                label = "Hot_"
                            End If
                            k = 1
                            Do While data_flux_hen(I, j + k) = "" And j + k <> colpinch
                                data_flux_hen(I, j + k) = repere_traits_Clement
                                k = k + 1
                            Loop
                            puissance = Abs(data_flux_hen(I, j - 1) - data_flux_hen(I, j + k)) * data_flux_hen(I, colreste + 1)
                            If puissance > 0.001 Or Abs(data_flux_hen(I, j - 1) - data_flux_hen(I, j + k)) > 0.001 Then   'MODIF CRE 21/09/2012   Si puissance supérieure à 1 kW, ou si écart de température > 1 °C, on affiche le besoin
                                data_flux_hen(I, j) = label & puissance
                                data_flux_hen(I, colreste + 4) = data_flux_hen(I, colreste + 4) - puissance
                                j = j + k
                            Else
                                data_flux_hen(I, j) = repere_traits_Clement
                            End If
                        End If
                    Else
                        j = colpinch - 1
                    End If
                Next
            End If
            If data_flux_hen(I, colpinch + 2) <> "" Then
                For j = colpinch + 3 To colreste - 3
                    If data_flux_hen(I, j - 1) <> "" Then
                        If data_flux_hen(I, j) = "" Then
                            If data_flux_hen(I, 1) = "Ch" Then
                                label = "Col_"
                            Else
                                label = "Hot_"
                            End If
                            k = 1
                            Do While data_flux_hen(I, j + k) = "" And j + k <> colreste - 2
                                data_flux_hen(I, j + k) = repere_traits_Clement
                                k = k + 1
                            Loop
                            puissance = Abs(data_flux_hen(I, j - 1) - data_flux_hen(I, j + k)) * data_flux_hen(I, colreste + 1)
                            If puissance > 0.001 Or Abs(data_flux_hen(I, j - 1) - data_flux_hen(I, j + k)) > 0.001 Then    'MODIF CRE 21/09/2012   Si puissance supérieure à 1 kW, ou si écart de température > 1 °C, on affiche le besoin
                                data_flux_hen(I, j) = label & puissance
                                data_flux_hen(I, colreste + 5) = data_flux_hen(I, colreste + 5) - puissance
                                j = j + k
                            Else
                                data_flux_hen(I, j) = repere_traits_Clement
                            End If
                        End If
                    Else
                        j = colreste - 3
                    End If
                Next
            End If
        Next

        'Nettoyage du tableau des flux d'utilités
'For i = LBound(data_flux_hen(), 1) To UBound(data_flux_hen(), 1)
'    For j = LBound(data_flux_hen(), 2) To UBound(data_flux_hen(), 2)
'        Feuil63.Cells(50 + i, j).Value = data_flux_hen(i, j)
'    Next j
'Next i
'If nbutilities_updated > 0 Then
'    For i = LBound(data_flux_utilities(), 1) To UBound(data_flux_utilities(), 1)
'        For j = LBound(data_flux_utilities(), 2) To UBound(data_flux_utilities(), 2)
'            Feuil63.Cells(50 + UBound(data_flux_hen(), 1) + i, j).Value = data_flux_utilities(i, j)
'        Next j
'    Next i
'End If
        Dim p As Integer
        For I = 1 To nbutilities_updated
            If data_flux_utilities(I, 1) = "UCh" Then
                For j = 3 To colreste - 2
                    If data_flux_utilities(I, j) = "" Then
                        k = 1
                        Do While data_flux_utilities(I, j + k) = ""
                            k = k + 1
                        Loop
                        If k + j <> colreste - 2 Then
                            data_flux_utilities(I, j - 1) = data_flux_utilities(I, j + k)
                            p = 0
                            Do While data_flux_utilities(I, j + p) = ""
                                data_flux_utilities(I, j + p) = repere_traits_Clement
                                p = p + 1
                            Loop
                        End If
                        j = j + k
                    End If
                Next
            Else
                For j = UBound(data_flux_utilities(), 2) To 1 Step -1

                Next
            End If
        Next I



        'III.10. -------------------------------------------   Fin   ---------------------------------------------
        position_erreur = "  -III.10-end"
        'MsgBox ("Ajout des Utilités terminé")


    End If



    'ATTENTION : nécessite d'avir lancer VeriPinch(1) auparavant !!
    If state = 0 Then
        'IV. Finalisation générale et export
        'IV.1. --------------------------------------   Affinage du HEN   ---------------------------------------
        'IV.2. ------------------------   Finalisation générales des besoins restants   -------------------------
        'IV.3. ------------------   Affichage des grandeurs numériques sur la feuille Tampon   ------------------
        'IV.4. -------------------   Nettoyage de la grille de flux en prévision de l'export   ------------------
        'IV.5. -----------------------------------   Export graphique du HEN   ----------------------------------
        'IV.6. --------------------------------------------   Fin.   --------------------------------------------


        'IV.1. --------------------------------------   Affinage du HEN   ---------------------------------------
        position_erreur = "  -IV.1"
        'Préparation du HEN
        ReDim data_flux_all(0 To nbflux + nbutilities_updated, 1 To UBound(data_flux_hen(), 2))
        For I = 0 To nbflux
            For j = 1 To UBound(data_flux_hen(), 2)
                data_flux_all(I, j) = data_flux_hen(I, j)
            Next
        Next
        For I = 1 To nbutilities_updated
            For j = 1 To UBound(data_flux_hen(), 2)
                data_flux_all(nbflux + I, j) = data_flux_utilities(I, j)
            Next j
        Next I
        nbflux_final = nbflux + nbutilities_updated
        ReDim data_flux_all_final(0 To nbflux + nbutilities_updated, 1 To UBound(data_flux_hen(), 2))
        For I = 0 To nbflux_final
            data_flux_all_final(I, 1) = data_flux_all(I, 1)                        'type
            data_flux_all_final(I, 2) = data_flux_all(I, 2)                        'nom
            data_flux_all_final(I, 3) = data_flux_all(I, 3)                        'tmax
            data_flux_all_final(I, 4) = data_flux_all(I, 4)                        'première colonne

            data_flux_all_final(I, colpinch) = data_flux_all(I, colpinch)         'colpinch
            data_flux_all_final(I, colpinch + 1) = data_flux_all(I, colpinch + 1)    'colpinch
            data_flux_all_final(I, colpinch + 2) = data_flux_all(I, colpinch + 2)    'colpinch

            For j = colreste - 3 To colreste + 5                                    'la fin
                data_flux_all_final(I, j) = data_flux_all(I, j)
            Next j
        Next I

        'Faut-il affiner le HEN ? Si N échangeurs diagonals > 2, oui.
        diagonal_max = 1
        affbool = False
        diff = 0
        For I = 1 To nbflux_final
            For j = 3 To colreste - 4
                If left(data_flux_all(I, j), 2) = "P_" Then
                    k = CDbl(Right(data_flux_all(I, j), 2))
                    If left(data_flux_all(k, j), 2) = "P_" Then
                        trait = left(data_flux_all(I, j), Len(data_flux_all(I, j)) - 3)
                        trait = Right(trait, Len(trait) - 2)
                        trait2 = left(data_flux_all(k, j), Len(data_flux_all(k, j)) - 3)
                        trait2 = Right(trait2, Len(trait2) - 2)
                        If (CDbl(Replace(trait, ".", separateur_decimal)) - CDbl(Replace(trait2, ".", separateur_decimal))) > 0.000001 Then
                            If k > nbflux Or I > nbflux Then
                                diff = diff + 1
                            End If
                        End If
                    Else
                        If k > nbflux Or I > nbflux Then
                            diff = diff + 1
                        End If
                    End If
                End If
            Next j
        Next I
        If diff / 2 >= diagonal_max Then    'MODIF CRE 19/09/2012   Besoin de définir cette variable dans l'interface ?
            affbool = True
        End If
        affbool = False    'MODIF CRE 25/09/2012   ATTENTION : en forçant affbool à False, on peut empecher l'affinage

        'Affinage (si besoin)
        If affbool = True Then
            erreur = ""

            'Annulation des informations dans la dernière colonne   'MODIF CRE 25/09/2012   Sinon l'affinage ne fonctionne pas. La dernière colonne était remplie avec l'ordre des flux
            For j = 1 To nbflux_final
                data_flux_all(j, UBound(data_flux_all(), 2)) = ""
            Next

            'Recherche des échangeurs sur Utilité Chaude. Si trouvé, on le décale à gauche, et on décale tous les HEX à gauche des flux concernés
            ok = True
            temp = 4
            For I = nbflux + 1 To nbflux_final
                If Right(data_flux_all(I, 1), 2) = "Ch" And data_flux_all(I, UBound(data_flux_all(), 2)) = "" Then    'pour toute utilité chaude
                    Call SearchAndDestroy(data_flux_all(), data_flux_all_final(), colpinch, UBound(data_flux_all_final(), 2), I, "aval", 4, temp, ok, True, I)
                    'Annulation des informations dans la dernière colonne   'MODIF CRE 25/09/2012   Sinon l'affinage ne fonctionne pas. La dernière colonne était remplie avec l'ordre des flux
                    For j = 1 To nbflux_final
                        data_flux_all_final(j, UBound(data_flux_all_final(), 2)) = ""
                    Next
                    'Mise à jour de la colonne d'écriture   'MODIF CRE 26/09/2012   Mise à jour de la colonne d'écriture, pour le passage à l'utilité suivante
                    For j = colpinch - 1 To 3 Step -1
                        For k = 1 To nbflux_final
                            If data_flux_all_final(k, j) <> "" Then
                                temp = j
                                k = nbflux_final
                                j = 3
                            End If
                        Next
                    Next
                End If
            Next

            If ok = True Then
                'Ajout des tirets
                colmax = 0
                For I = 1 To nbflux_final
                    For j = colpinch - 1 To 4 Step -1
                        If data_flux_all_final(I, j) <> "" And data_flux_all_final(I, j) <> repere_traits_Clement Then
                            'Si on a trouvé des flux plus loin, on agrandit la famille
                            If j > colmax Then
                                colmax = j
                                k = 0
                                Do While colmax - k > 4
                                    If data_flux_all_final(I, colmax - k) = "" Then
                                        data_flux_all_final(I, colmax - k) = repere_traits_Clement
                                    End If
                                    k = k + 1
                                Loop
                            End If
                            j = 4
                        End If
                    Next
                Next
                'Ajout des tirets manquants à gauche et à droite, mais toujours à gauche de colmax
                For I = 1 To nbflux_final
                    For j = colmax To 4 Step -1
                        If data_flux_all_final(I, j) = "" Then
                            data_flux_all_final(I, j) = repere_traits_Clement
                        End If
                    Next
                Next

                'Ensuite, recopie de tous les autres échangeurs à gauche du pinch
                'MODIF CRE 08/09/2012   Une fois qu'on a bouclé à gauche du pinch, et qu'on a affiné le HEN, on recopie tout brutalement
                If colmax > 4 Then
                    colmax = colmax + 2
                Else
                    colmax = colmax + 1
                End If
                For j = 5 To colreste - 2
                    vide = True
                    vide2 = True
                    I = 1
                    Do While I <= nbflux_final
                        If (data_flux_all(I, j) = repere_traits_Clement Or data_flux_all(I, j) = "000000" Or data_flux_all(I, j) = "") Then
                            I = I + 1
                        Else
                            vide = False
                            I = nbflux_final + 1
                        End If
                    Loop
                    If vide = True Then
                        k = 1
                        Do While k <= nbflux_final
                            If (data_flux_all(k, j + 1) = repere_traits_Clement Or data_flux_all(k, j + 1) = "000000" Or data_flux_all(k, j + 1) = "") Then
                                k = k + 1
                            Else
                                k = nbflux_final + 1
                                vide2 = False
                            End If
                        Loop
                    End If

                    If vide = True And vide2 = True Then
                        j = j + 3
                    Else
                        For I = 0 To nbflux_final
                            For k = 0 To 3
                                data_flux_all_final(I, colmax + k) = data_flux_all(I, j + k)
                            Next
                        Next
                        colmax = colmax + 4
                        j = j + 3
                        vide = True
                        vide2 = True
                    End If
                Next

                'Même traitement à droite du pinch, avec les utilités froides # A FINIR
                For I = nbflux + 1 To nbflux_final
                    If Right(data_flux_all(I, 1), 2) = "Fr" Then    'pour toute utilité froide
                        Call SearchAndDestroy(data_flux_all(), data_flux_all_final(), colpinch, UBound(data_flux_hen(), 2), I, "aval", 4, 4, ok, False)
                    End If
                Next

                For I = 0 To nbflux_final
                    For j = colpinch + 3 To colreste - 3
                        data_flux_all_final(I, j) = data_flux_all(I, j)
                    Next
                Next

                'Nettoyage final
                For I = 1 To nbflux_final
                    'Traitement des flux de process
                    If I <= nbflux Then
                        For j = 4 To colreste - 4
                            'Gestion des tirets
                            If data_flux_all_final(I, j) = "" Then
                                data_flux_all_final(I, j) = repere_traits_Clement
                            End If
                        Next
                        'Traitement des Utilités
                    Else
                        If data_flux_all_final(I, 1) = "UCh" Then
                            For j = colreste - 4 To 4 Step -1
                                If data_flux_all_final(I, j) <> "" Then
                                    If data_flux_all_final(I, j) = repere_traits_Clement Or data_flux_all_final(I, j) = "000000" Then
                                        data_flux_all_final(I, j) = ""
                                    Else
                                        j = 4
                                    End If
                                End If
                            Next
                        Else
                            For j = 4 To colreste - 4
                                If data_flux_all_final(I, j) <> "" Then
                                    If data_flux_all_final(I, j) = repere_traits_Clement Or data_flux_all_final(I, j) = "000000" Then
                                        data_flux_all_final(I, j) = ""
                                    Else
                                        j = colreste - 4
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next

            Else
                'ok = false : boucle. on reconstruit le HEN avec data_flux_hen et data_flux_utilities
                For I = 0 To nbflux
                    For j = 1 To UBound(data_flux_hen(), 2)
                        data_flux_all_final(I, j) = data_flux_hen(I, j)
                    Next
                Next
                For I = 1 To nbutilities_updated
                    For j = 1 To UBound(data_flux_hen(), 2)
                        data_flux_all_final(I + nbflux, j) = data_flux_utilities(I, j)
                    Next
                Next
                erreur = "Pas d'affinage. Raison : " & erreur
            End If
        Else
            'affbool = false : pas besoin d'affiner. on reconstruit le HEN avec data_flux_hen et data_flux_utilities
            For I = 0 To nbflux
                For j = 1 To UBound(data_flux_hen(), 2)
                    data_flux_all_final(I, j) = data_flux_hen(I, j)
                Next
            Next
            For I = 1 To nbutilities_updated
                For j = 1 To UBound(data_flux_hen(), 2)
                    data_flux_all_final(I + nbflux, j) = data_flux_utilities(I, j)
                Next
            Next
            erreur = "Pas besoin d'affiner"
        End If

        If erreur = "" Then
            erreur = "Pas d'erreur"
        End If
        If PINCH_DANS_OCEAN = False Then
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(4, 13).Value = "Message lors de l'affinage du HEN"
            CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Cells(4, 14).Value = erreur
        End If



        'IV.2. ------------------------   Finalisation générales des besoins restants   -------------------------
        position_erreur = "  -IV.2"

        'MODIF CRE 10/10/2012   Mise à jour des informations de colonne pour les HEX
        'MODIF CRE 24/10/2012   Inutile ici, sera renseigné dans IV.4
        '        For i = 1 To nbflux_final
        '            For j = 3 To colreste - 2
        '                If left(data_flux_all_final(i, j), 2) = "P_" Then
        '                    k = 2
        '                    Do While Right(data_hen(1, k), Len(data_hen(1, k)) - 3) <> Right(data_flux_all_final(0, j), Len(data_flux_all_final(0, j)) - 3)
        '                        k = k + 1
        '                    Loop
        '                    If data_flux_all_final(i, 1) = "Fr" Or data_flux_all_final(i, 1) = "UFr" Then
        '                        data_hen(12, k) = j
        '                    Else
        '                        data_hen(11, k) = j
        '                    End If
        '                End If
        '            Next
        '        Next

        'Insertion de colonnes pour les besoins restants
        j = 4
        jmax = UBound(data_flux_all_final(), 2)
        Do While j <> jmax
            sec = False
            For I = 1 To nbflux_final
                If left(data_flux_all_final(I, j), 4) = "Hot_" Or left(data_flux_all_final(I, j), 4) = "Col_" Or left(data_flux_all_final(I, j), 2) = "R_" Then
                    If data_flux_all_final(I, j + 1) = repere_traits_Clement Then
                        'si on n'a rien à coté de ce besoin
                        'If data_flux_all_final(0, j + 2) = "" Then 'MODIF CRE 11/10/2012   En supprimant cette condition, on vérifie si le besoin est entre 2 ovales dans tous les cas
                        'et si il y a quelque chose en tête, on teste si le besoin est entre 2 ovales
                        ok = True
                        k = 1
                        For k = 1 To nbflux_final
                            If (data_flux_all_final(k, j + 2) <> repere_traits_Clement And data_flux_all_final(k, j + 2) <> "" And left(data_flux_all_final(k, j + 2), 4) <> "Col_" And left(data_flux_all_final(k, j + 2), 4) <> "Hot_") Or (data_flux_all_final(k, j) <> repere_traits_Clement And data_flux_all_final(k, j) <> "" And left(data_flux_all_final(k, j), 4) <> "Col_" And left(data_flux_all_final(k, j), 4) <> "Hot_") Then
                                If k = I Then   'MODIF CRE 17/01/2013   Ajout de ce test sinon on perdait une information (cas trouvé avec le cas d'étude Pinch quand on force un échangeur entre 60 et 30°C sur le flux 1, on perdait la présence du HEX quand on arrivait à k = nbflux_final)
                                    For m = k + 1 To nbflux_final
                                        If (data_flux_all_final(m, j + 2) <> repere_traits_Clement And data_flux_all_final(m, j + 2) <> "" And left(data_flux_all_final(m, j + 2), 4) <> "Col_" And left(data_flux_all_final(m, j + 2), 4) <> "Hot_") Or (data_flux_all_final(m, j) <> repere_traits_Clement And data_flux_all_final(m, j) <> "" And left(data_flux_all_final(m, j), 4) <> "Col_" And left(data_flux_all_final(m, j), 4) <> "Hot_") Then
                                            ok = False
                                            m = nbflux_final
                                        End If
                                    Next
                                Else
                                    ok = False
                                End If
                                k = nbflux_final
                            End If
                        Next
                        If ok = True Then
                            'on va décaler le besoin de 2 crans, et ajouter les températures
                            For m = 1 To nbflux_final
                                If left(data_flux_all_final(m, j), 4) = "Hot_" Or left(data_flux_all_final(m, j), 4) = "Col_" Then
                                    data_flux_all_final(m, j + 2) = data_flux_all_final(m, j)
                                    data_flux_all_final(m, j) = repere_traits_Clement
                                    k = 0
                                    Do While data_flux_all_final(m, j - k) = repere_traits_Clement
                                        k = k + 1
                                    Loop
                                    data_flux_all_final(m, j + 1) = data_flux_all_final(m, j - k)
                                    k = 0
                                    Do While data_flux_all_final(m, j + 3 + k) = repere_traits_Clement
                                        k = k + 1
                                    Loop
                                    data_flux_all_final(m, j + 3) = data_flux_all_final(m, j + 3 + k)
                                End If
                            Next
                            I = nbflux_final
                            j = j + 2
                        Else
                            sec = True
                        End If
                        'Else
                        '    sec = True
                        'End If
                    Else
                        sec = True
                    End If

                    If sec = True Then
                        'si on a quelque chose, on décale
                        jmax = jmax + 4
                        ReDim Preserve data_flux_all_final(0 To nbflux_final, 1 To jmax)
                        For k = 0 To nbflux_final
                            For m = jmax - 4 To j Step -1
                                data_flux_all_final(k, m + 4) = data_flux_all_final(k, m)
                                data_flux_all_final(k, m) = ""
                            Next
                            If left(data_flux_all_final(k, m + 5), 4) = "Hot_" Or left(data_flux_all_final(k, m + 5), 4) = "Col_" Or left(data_flux_all_final(k, m + 5), 2) = "R_" Then
                                data_flux_all_final(k, m + 1) = repere_traits_Clement
                                data_flux_all_final(k, m + 2) = data_flux_all_final(k, m)
                                data_flux_all_final(k, m + 3) = data_flux_all_final(k, m + 5)
                                N = 0
                                Do While data_flux_all_final(k, m + 6 + N) = repere_traits_Clement
                                    N = N + 1
                                Loop
                                data_flux_all_final(k, m + 4) = data_flux_all_final(k, m + 6 + N)
                                data_flux_all_final(k, m + 5) = repere_traits_Clement
                                If data_flux_all_final(k, m - 2) = repere_traits_Clement Then
                                    data_flux_all_final(k, m) = repere_traits_Clement
                                End If
                            ElseIf k <> 0 Then
                                data_flux_all_final(k, m + 1) = data_flux_all_final(k, m + 5)
                                data_flux_all_final(k, m + 2) = data_flux_all_final(k, m + 5)
                                data_flux_all_final(k, m + 3) = data_flux_all_final(k, m + 5)
                                data_flux_all_final(k, m + 4) = data_flux_all_final(k, m + 5)
                            End If
                        Next
                        I = nbflux_final
                        j = j + 4
                        'et mise à jour des valeurs de colplacehexch et fr dans data_hen()  'MODIF CRE 02/10/2012   Mise à jour des indices des colonnes
                        'MODIF CRE 24/10/2012   Inutile ici, sera renseigné dans IV.4
                        'For k = 2 To UBound(data_hen(), 2)
                        '    If j - 2 < CInt(data_hen(11, k)) Then
                        '        data_hen(11, k) = data_hen(11, k) + 4
                        '    End If
                        '    If j - 2 < CInt(data_hen(12, k)) Then
                        '        data_hen(12, k) = data_hen(12, k) + 4
                        '    End If
                        'Next
                    End If
                End If
            Next
            j = j + 1
        Loop

        'Ajout des puissances retantes sur les Utilités chaudes MODIF CRE 19/10/2012    mise à jour de cette partie
        Dim dec As Boolean      'décalage : true/false
        Dim inser As Integer    'colonne d'insertion
        Call GetCols(data_flux_all_final())
        'détection d'un éventuel besoin d'insérer une colonne
        If data_flux_all_final(0, colreste - 5) = "t_pinch_cen" Then
            inser = colreste - 7
        Else
            inser = colreste - 3
        End If
        For I = nbflux + 1 To nbflux_final
            If data_flux_all_final(I, colreste + 4) <> "" Then
                If data_flux_all_final(I, colreste + 4) > 0.001 Then
                    If data_flux_all_final(I, inser - 2) <> repere_traits_Clement And data_flux_all_final(I, inser - 2) <> "" Then
                        dec = True
                        I = nbflux_final
                    Else
                        If I < nbflux_final Then
                            For j = I + 1 To nbflux_final
                                If left(data_flux_all_final(j, inser - 2), 2) = "P_" Then
                                    If CInt(Right(data_flux_all_final(j, inser - 2), 2)) < I Then
                                        dec = True
                                        j = nbflux_final
                                        I = nbflux_final
                                    End If
                                End If
                            Next j
                        End If
                    End If
                End If
            End If
        Next I
        'insertion de colonnes si besoin
        If dec = True Then
            dec = False
            data_flux_hen_size = UBound(data_flux_all_final(), 2) + 4
            ReDim Preserve data_flux_all_final(0 To nbflux_final, 1 To data_flux_hen_size)
            colreste = colreste + 4
            For k = 0 To nbflux_final
                For m = data_flux_hen_size - 4 To inser Step -1
                    data_flux_all_final(k, m + 4) = data_flux_all_final(k, m)
                    data_flux_all_final(k, m) = ""
                Next
            Next
        Else
            inser = inser - 4
        End If
        'ajout des puissances restantes
        For I = nbflux + 1 To nbflux_final
            If data_flux_all_final(I, colreste + 4) <> "" Then
                If data_flux_all_final(I, colreste + 4) > 0.001 Then
                    data_flux_all_final(I, inser + 2) = "R_" & data_flux_all_final(I, colreste + 4)
                    j = 1
                    Do While data_flux_all_final(I, inser + 2 - j) = ""
                        data_flux_all_final(I, inser + 2 - j) = repere_traits_Clement
                        j = j + 1
                    Loop
                    data_flux_all_final(I, inser + 1) = data_flux_all_final(I, inser + 2 - j)
                    j = 1
                    Do While data_flux_all_final(I, inser + 2 + j) = ""
                        data_flux_all_final(I, inser + 2 + j) = repere_traits_Clement
                        j = j + 1
                    Loop
                    data_flux_all_final(I, inser + 3) = data_flux_all_final(I, inser + 2 + j)
                End If
            End If
        Next I

        'Ajout des puissances retantes sur les Utilités froides MODIF CRE 19/10/2012    mise à jour de cette partie
        Call GetCols(data_flux_all_final())
        'détection d'un éventuel besoin d'insérer une colonne
        If data_flux_all_final(0, 6) = "t_pinch_cen" Then
            inser = 8
        Else
            inser = 4
        End If
        For I = nbflux + 1 To nbflux_final
            If data_flux_all_final(I, colreste + 5) <> "" Then
                If data_flux_all_final(I, colreste + 5) > 0.001 Then
                    If data_flux_all_final(I, inser + 2) <> repere_traits_Clement And data_flux_all_final(I, inser + 2) <> "" Then
                        dec = True
                        I = nbflux_final
                    Else
                        If I < nbflux_final Then
                            For j = I + 1 To nbflux_final
                                If left(data_flux_all_final(j, inser + 2), 2) = "P_" Then
                                    If CInt(Right(data_flux_all_final(j, inser + 2), 2)) < I Then
                                        dec = True
                                        j = nbflux_final
                                        I = nbflux_final
                                    End If
                                End If
                            Next j
                        End If
                    End If
                End If
            End If
        Next I
        'insertion de colonnes si besoin
        If dec = True Then
            dec = False
            data_flux_hen_size = UBound(data_flux_all_final(), 2) + 4
            ReDim Preserve data_flux_all_final(0 To nbflux_final, 1 To data_flux_hen_size)
            colreste = colreste + 4
            For k = 0 To nbflux_final
                For m = data_flux_hen_size - 4 To inser Step -1
                    data_flux_all_final(k, m + 4) = data_flux_all_final(k, m)
                    data_flux_all_final(k, m) = ""
                Next
            Next
            'inser = inser - 4  'MODIF CRE 06/11/12 Pas besoin de ça ici, ça plante sinon
        End If
        'ajout des puissances restantes
        For I = nbflux + 1 To nbflux_final
            If data_flux_all_final(I, colreste + 5) <> "" Then
                If data_flux_all_final(I, colreste + 5) > 0.001 Then
                    data_flux_all_final(I, inser + 2) = "R_" & data_flux_all_final(I, colreste + 5)
                    j = 1
                    Do While data_flux_all_final(I, inser + 2 - j) = ""
                        data_flux_all_final(I, inser + 2 - j) = repere_traits_Clement
                        j = j + 1
                    Loop
                    data_flux_all_final(I, inser + 1) = data_flux_all_final(I, inser + 2 - j)
                    j = 1
                    Do While data_flux_all_final(I, inser + 2 + j) = ""
                        data_flux_all_final(I, inser + 2 + j) = repere_traits_Clement
                        j = j + 1
                    Loop
                    data_flux_all_final(I, inser + 3) = data_flux_all_final(I, inser + 2 + j)
                End If
            End If
        Next I



        'IV.3. ------------------   Affichage des grandeurs numériques sur la feuille Tampon   ------------------
        position_erreur = "  -IV.3"
        data_flux_hen_size = UBound(data_flux_all_final(), 2)

        Call ecrire_VeriPinch_position_IV3_1(jmax)



        'MODIF CRE 05/09/2012   calcul des besoins chauds et froids non couverts par les Utilités
        Dim besoins_cold_final As Double
        Dim besoins_hot_final As Double

        Call GetCols(data_flux_all_final())
        For I = 1 To nbflux_final
            For j = 3 To colreste - 2
                If left(data_flux_all_final(I, j), 4) = "Col_" Then
                    besoins_cold_final = besoins_cold_final + Right(data_flux_all_final(I, j), Len(data_flux_all_final(I, j)) - 4)
                ElseIf left(data_flux_all_final(I, j), 4) = "Hot_" Then
                    besoins_hot_final = besoins_hot_final + Right(data_flux_all_final(I, j), Len(data_flux_all_final(I, j)) - 4)
                End If
            Next j
        Next I

        'MODIF CRE 10/10/2012   Ajout pour affichage des informations sur tampon
        Call ecrire_VeriPinch_position_IV3_2(besoins_hot_final, besoins_cold_final)
        
        'MODIF CRE 16/10/2012   Affichage dans PINCH X - efface les 0% pour éviter de surcharger le camembert
        'MODIF CRE 17/10/2012   Commenté, à affiner
        For I = 7 To 23
            'If CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(i, 13).Value = 0 Then
            '    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(i, 13).ClearContents
            'End If
        Next

        'Recalcul du nb de flux ch puis fr
        Call GetCols(data_flux_all_final())
        For I = 1 To UBound(data_flux_all_final())
            If data_flux_all_final(I, 1) = "Ch" Then
                nbflux_ch = nbflux_ch + 1
            ElseIf data_flux_all_final(I, 1) = "Fr" Then
                nbflux_fr = nbflux_fr + 1
            End If
        Next I

        'MODIF CRE 10/10/2012   Dernière écriture de résultats sur la page Tampon
        Call LastDisplay(data_flux_all_final(), pi_FiliereConsideree, nbflux, colpinch, colreste)



        'IV.4. -------------------   Nettoyage de la grille de flux en prévision de l'export   ------------------
        position_erreur = "  -IV.4"
        'Nettoyage, ssi on n'a que des flux chauds ou froids (pinch seul à gauche ou à droite)
        'MODIF CRE 17/10/2012   A affiner
        If nbflux_ch = 0 Or nbflux_fr = 0 Then   '*** Modif DCA 08/01/13 -> comprend pas la condition comentée à droite *** 'Or CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(6, 4).Value = "" Then  'MODIF CRE 02/10/2012   Ajout de cette condition. Comme ça on voit quand meme le pinch, quand il existe, même s'il est à une des extrémités
            If data_flux_all_final(0, 5) = "t_pinch_gch" Then   'pinch tout à gauche
                For I = 0 To UBound(data_flux_all_final())
                    For j = 4 To UBound(data_flux_all_final(), 2) - 4
                        data_flux_all_final(I, j) = data_flux_all_final(I, j + 4)
                    Next
                    If data_flux_all_final(I, 4) = "" Then
                        data_flux_all_final(I, 4) = data_flux_all_final(I, 4 + 2)
                        data_flux_all_final(I, 5) = repere_traits_Clement
                    End If
                Next
                ReDim Preserve data_flux_all_final(0 To UBound(data_flux_all_final()), 1 To colreste + 8 - 5)
                colreste = colreste - 4
                colpinch = colreste - 2
            ElseIf data_flux_all_final(0, colreste - 4) = "t_pinch_dte" Then    'pinch tout à droite
                For I = 0 To UBound(data_flux_all_final())
                    For j = colpinch - 1 To UBound(data_flux_all_final(), 2) - 4
                        data_flux_all_final(I, j) = data_flux_all_final(I, j + 4)
                    Next
                    If data_flux_all_final(I, colpinch) = "" Then
                        data_flux_all_final(I, colpinch) = data_flux_all_final(I, colpinch - 2)
                        data_flux_all_final(I, colpinch - 1) = repere_traits_Clement
                    End If
                Next
                ReDim Preserve data_flux_all_final(0 To UBound(data_flux_all_final()), 1 To colreste + 8 - 5)
                colreste = colreste - 4
                colpinch = colreste - 2
            End If
        End If

        'Verification des colonnes des HEX  'MODIF CRE 22/10/2012   Modification pour mise à jour et affinage de ce bout de code
        Dim tin_flux As Double
        For I = 1 To UBound(data_flux_all_final())
            For j = 1 To UBound(data_flux_all_final(), 2)
                If left(data_flux_all_final(I, j), 2) = "P_" Then
                    label = left(data_flux_all_final(I, j), Len(data_flux_all_final(I, j)) - 3)
                    label = Right(label, Len(label) - 2)
                    If left(data_flux_all_final(I, 1), 2) = "Ch" Or left(data_flux_all_final(I, 1), 3) = "UCh" Then
                        m = 3
                        N = 5
                        tin_flux = data_flux_all_final(I, j - 1)
                    Else
                        m = 4
                        N = 7
                        tin_flux = data_flux_all_final(I, j + 1)
                    End If
                    nom_flux = data_flux_all_final(I, 2)
                    For k = 2 To UBound(data_hen(), 2)
                        If Abs(CDbl(data_hen(2, k)) - CDbl(label)) < 0.001 And data_hen(m, k) = nom_flux And Abs(CDbl(data_hen(N, k)) - tin_flux) < 0.01 Then
                            If left(data_flux_all_final(I, 1), 2) = "Ch" Or left(data_flux_all_final(I, 1), 3) = "UCh" Then
                                data_hen(11, k) = CStr(j)
                            Else
                                data_hen(12, k) = CStr(j)
                            End If
                            k = UBound(data_hen(), 2)
                        End If
                    Next
                End If
            Next
        Next

        If PINCH_DANS_OCEAN = False Then
            'Affichage de la valeur de surface totale d'échange 'MODIF CRE 06/11/12 On écrit plutot la surface ici, après toutes les créations de HEX
            CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(33, 4).Value = Replace(surftot, separateur_decimal, ".")
        End If


        'IV.5. -----------------------------------   Export graphique du HEN   ----------------------------------
        If PINCH_DANS_OCEAN = False Then
            Call DisplayHEN(pi_FiliereConsideree, data_flux_all_final(), UBound(data_flux_all_final(), 2))
        Else
            If RecuperationEnergetiquePossible = True Then
                'On lit la température de PINCH
                T_pinch_reseau_DCA = point_pinch + 0.01  'Delta à cause du repère des échangeurs au dessus et au dessous du pinch
                CodeNames("Feuil63").Range("PINCH_pincement_point_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1) = point_pinch
            End If
'            'On doit lire les flux qu'il reste à couvrir ou les utilités que l'on n'a pas utilisées en intégralité
'            For i = LBound(data_flux_all_final(), 1) To UBound(data_flux_all_final(), 1)
'                For j = LBound(data_flux_all_final(), 2) To UBound(data_flux_all_final(), 2)
'                    Feuil63.Cells(50 + i, j).Value = data_flux_all_final(i, j)
'                Next j
'            Next i
            Call representationDCA_recuperation_puissances_non_echangees
        End If


        'IV.6. --------------------------------------------   Fin.   --------------------------------------------
        'Msgbox("Fin de l'affichage des informations du pinch")

    End If

    'MODIF CRE 05/09/2012   Résactivation de la mise à jour de l'affichage à l'écran après l'éxécution du code
    If PINCH_DANS_OCEAN = False Then
        Application.ScreenUpdating = True
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.1" & position_erreur

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub



Public Sub LastDisplay(t() As String, ste As Integer, nbfl As Integer, colp As Integer, colr As Integer)
    'Affichage final des informations sur la feuille Tampon (en mode admin) et sur la feuille PINCH X
    'Note: pour éviter toute confusion, les noms des variables passées en paramètre ont été tronquées (à part pour t()).
    'MODIF CRE 10/10/2012   Export de cette partie du code dans une subroutine : VeriPinch a dépassé la capacité acceptée par VBA (procédure trop grande)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String
    Dim position_erreur As String

    't      grille des flux
    'ste    numéro de la filière concernée  (pi_FiliereConsideree)
    'nbfl   nombre de flux de Process (nbflux)
    'colp   indice de la colonne du pincement (colpinch)
    'colr   indice de la colonne des puissances totales (colreste)

    Dim ligprin As Integer
    Dim I As Integer
    Dim j As Integer

    position_erreur = ""

    'MODIF CRE 10/10/2012   besoin de traiter la grille de flux pour les calculs qui suivent
    ligprin = 4 + 7 + 3 * (nbfl + 2) + 1 + (UBound(t()) - nbfl) + UBound(t()) + 2 + 7 * 3 + 12 + 4
    For I = 1 To UBound(t())
        If t(I, colr + 4) = "" Then
            t(I, colr + 4) = "0"
        End If
        If t(I, colr + 5) = "" Then
            t(I, colr + 5) = "0"
        End If
        For j = 3 To colp
            If left(t(I, j), 4) = "Col_" Or left(t(I, j), 4) = "Hot_" Then
                t(I, colr + 4) = CDbl(t(I, colr + 4)) + CDbl(Right(t(I, j), Len(t(I, j)) - 4))
            End If
        Next
        For j = colp + 2 To colr - 2
            If left(t(I, j), 4) = "Col_" Or left(t(I, j), 4) = "Hot_" Then
                t(I, colr + 5) = CDbl(t(I, colr + 5)) + CDbl(Right(t(I, j), Len(t(I, j)) - 4))
            End If
        Next
    Next

    If PINCH_DANS_OCEAN = False Then
        'MODIF CRE 10/10/2012   variable des flux process
        CodeNames("Feuil_Tampon" & ste).Cells(ligprin, 2).Value = "Recap des flux process"
        For I = 1 To nbfl
            CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I, 3).Value = t(I, 2)    'nom
            If t(I, 1) = "Fr" Then
                If t(I, colr - 2) <> "" Then
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I, 4).Value = Replace(t(I, colr - 2), separateur_decimal, ".")    'tin
                Else
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I, 4).Value = Replace(t(I, colp + 1), separateur_decimal, ".")    'tin
                End If
                If t(I, 3) <> "" Then
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I, 5).Value = Replace(t(I, 3), separateur_decimal, ".")    'tout
                Else
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I, 5).Value = Replace(t(I, colp + 1), separateur_decimal, ".")    'tout
                End If
            Else
                If t(I, 3) <> "" Then
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I, 4).Value = Replace(t(I, 3), separateur_decimal, ".")    'tin
                Else
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I, 4).Value = Replace(t(I, colp + 1), separateur_decimal, ".")    'tin
                End If
                If t(I, colr - 2) <> "" Then
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I, 5).Value = Replace(t(I, colr - 2), separateur_decimal, ".")    'tout
                Else
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I, 5).Value = Replace(t(I, colp + 1), separateur_decimal, ".")    'tout
                End If
            End If
            CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I, 6).Value = Replace(Abs(t(I, colr)), separateur_decimal, ".")    'ptotale
            CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I, 7).Value = Replace((Abs(t(I, colr)) - (Abs(t(I, colr + 4)) + Abs(t(I, colr + 5)))) / Abs(t(I, colr)), separateur_decimal, ".")    '% couvert
        Next

        'MODIF CRE 10/10/2012   variable des flux utilités
        CodeNames("Feuil_Tampon" & ste).Cells(ligprin, 10).Value = "Recap des flux utilités"
        For I = nbfl + 1 To UBound(t())
            CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I - nbfl, 11).Value = t(I, 2)    'nom
            If t(I, 1) = "UFr" Then
                CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I - nbfl, 12).Value = Replace(t(I, colreste - 2), separateur_decimal, ".")    'tin
                j = 1
                Do While t(I, colreste - 2 - j) <> "" And colreste - 2 - j >= 3
                    j = j + 1
                Loop
                If j <> 1 Then
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I - nbfl, 13).Value = Replace(t(I, colreste - 2 - j + 1), separateur_decimal, ".")    'tout
                Else
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I - nbfl, 13).Value = Replace(t(I, 3), separateur_decimal, ".")    'tout
                End If

                If t(I, colr - 2) <> "" Then
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I - nbfl, 12).Value = Replace(t(I, colr - 2), separateur_decimal, ".")    'tin
                Else
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I - nbfl, 12).Value = Replace(t(I, colp + 1), separateur_decimal, ".")    'tin
                End If
                If t(I, 3) <> "" Then
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I - nbfl, 13).Value = Replace(t(I, 3), separateur_decimal, ".")    'tout
                Else
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I - nbfl, 13).Value = Replace(t(I, colp + 1), separateur_decimal, ".")    'tout
                End If
            Else
                CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I - nbfl, 12).Value = Replace(t(I, 3), separateur_decimal, ".")    'tin
                j = 1
                Do While t(I, 3 + j) <> "" And 3 + j <= colreste - 2
                    j = j + 1
                Loop
                If j <> 1 Then
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I - nbfl, 13).Value = Replace(t(I, 3 + j - 1), separateur_decimal, ".")    'tout
                Else
                    CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I - nbfl, 13).Value = Replace(t(I, colreste - 2), separateur_decimal, ".")    'tout
                End If
            End If
            CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I - nbfl, 14).Value = Replace(Abs(t(I, colr)), separateur_decimal, ".")    'ptotale
            CodeNames("Feuil_Tampon" & ste).Cells(ligprin + I - nbfl, 15).Value = Replace((Abs(t(I, colr)) - (Abs(t(I, colr + 4)) + Abs(t(I, colr + 5)))) / Abs(t(I, colr)), separateur_decimal, ".")    '% couvert
        Next
    End If

    'Affichage de la liste des interdictions dans la feuille PINCH X
    If nbhexforbidden > 0 Then
        For I = 1 To UBound(data_hex_forbidden()) - 1
            If PINCH_DANS_OCEAN = True Then
                CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_InterdictionEchanges_" & marqueur_etape & pi_FiliereConsideree).Cells(1, I) = I
'                CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_InterdictionEchanges_" & marqueur_etape & pi_FiliereConsideree).Cells(2, i) = data_hex_forbidden(1, i)
                Call ecriture_nom_flux_interdiction_echange(data_hex_forbidden(1, I), CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Name, "PINCH_InterdictionEchanges_" & marqueur_etape & pi_FiliereConsideree, 2, I, Feuil63.Name, "PINCH_NomsPlagesFluxProcede")
'                CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_InterdictionEchanges_" & marqueur_etape & pi_FiliereConsideree).Cells(3, i) = data_hex_forbidden(2, i)
                Call ecriture_nom_flux_interdiction_echange(data_hex_forbidden(2, I), CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Name, "PINCH_InterdictionEchanges_" & marqueur_etape & pi_FiliereConsideree, 3, I, Feuil63.Name, "PINCH_NomsPlagesFluxProcede")
            Else
                CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(65, 3 + I) = data_hex_forbidden(1, I)
                CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(66, 3 + I) = data_hex_forbidden(2, I)
            End If
        Next
    End If


    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = ".113 " & position_erreur

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub


Public Sub SearchAndDestroy(t() As String, t2() As String, size As Integer, bigsize As Integer, I As Integer, sens As String, col As Integer, col2 As Integer, ok As Boolean, nouv As Boolean, Optional util As Integer)
    'Cette procédure cherche les échangeurs présents sur la ligne "i", dans le sens "sens", à partir de la colonne "col". Elle écrit cet échangeur dans le nouveau tableau de flux à partir de la colonne "col2"

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String
    Dim position_erreur As String

    't      tableau de flux qu'on veut affiner
    't2     tableau de flux après affinage
    'size   taille du tableau de flux (permet de déterminer les bornes des boucles)
    'i      ligne du flux sur lequel on cherche des échangeurs
    'sens   sens de recherche des échangeurs
    'col    colonne à partir de laquelle on cherche
    'col2   colonne où on écrit (indispensable ?)
    'nouv   indique si c'est une nouvelle utilité ou pas

    Dim j As Integer
    Dim k As Integer
    Dim m As Integer
    Dim cv As Integer
    Dim cv2 As Integer
    Dim kb As Integer
    Dim p As Integer
    Dim indi As Integer
    Dim indi2 As Integer
    Dim recop As Integer
    Dim liga As Integer

    Dim deb As Integer      'colonne de début de la recherche
    Dim fin As Integer      'colonne de fin
    Dim ste As Integer      'sens de la recherche
    Dim kmax As Integer


    If ok = True Then
        If sens = "aval" Then
            deb = col
            fin = size - 5
            ste = 1
        ElseIf sens = "amont" Then
            deb = col
            fin = 4
            ste = -1
        End If

        If I > nbflux Then
            If sens = "aval" Then
                j = deb + 1
                Do While t(I, j) <> ""
                    j = j + 1
                Loop
                fin = j
            End If
        End If

        position_erreur = "initialisation"

        'Affichage du HEN lors de son affinage
        Tabelle9.Activate
        Range("A1:DA100").ClearContents
        For cv = 0 To UBound(t())
            For cv2 = 1 To size
                Tabelle9.Cells(4 + cv, 2 + cv2).Value = Replace(data_flux_all_final(cv, cv2), separateur_decimal, ".")
            Next
        Next

        'Recherche des échangeurs présents sur cette ligne (i) à partir de cette colonne (deb) dans ce sens (ste)
        For j = deb To fin Step ste
            If ok = False Then
                j = fin
            Else
                If util <> 0 Then
                    If I = util Then
                        For k = 1 To UBound(t())
                            t2(k, UBound(t(), 2)) = 0
                        Next
                    End If
                End If

                If t(I, j) = "" Then
                    Exit Sub
                Else
                    If t(I, j) <> repere_traits_Clement And t(I, j) <> "000000" And left(t(I, j), 4) <> "Hot_" And left(t(I, j), 4) <> "Col_" Then

                        'Décalage si amont
                        If sens = "amont" Then
                            p = 0
                            Do While t(I, j - p) <> repere_traits_Clement And j - p > 4
                                p = p + 1
                            Loop
                            p = p + 1
                            For m = 0 To UBound(t())
                                k = colpinch - 1
                                If col2 = 4 Then
                                    kmax = 5
                                Else
                                    kmax = col2
                                End If
                                Do While k >= kmax
                                    If t2(m, k) <> "" Then
                                        t2(m, k + p) = t2(m, k)
                                        t2(m, k) = ""
                                    End If
                                    k = k - 1
                                Loop
                            Next
                        End If

                        'Recopie
                        If sens = "amont" Then
                            indi = j - 2
                        Else
                            indi = j
                        End If

                        If nouv = True Then    'MODIF CRE 26/09/2012   Si on est passé à une nouvelle utilité (nouv = true), alors on recopie directement à col2
                            recop = col2
                        Else
                            If sens = "aval" Then
                                'vers l'aval, on prends la dernière colonne libre
                                k = 0
                                Do While t2(I, fin - k) = ""
                                    k = k + 1
                                Loop
                                recop = fin - k
                            Else
                                'vers l'amont, on écrit en colonne 4 (insertion directement au début)
                                recop = col2
                            End If
                        End If
                        For k = 0 To 3
                            t2(I, recop + 1 + k) = t(I, indi + k)
                        Next
                        position_erreur = "recopie 1"

                        'Recopie de l'autre flux
                        liga = CInt(Right(t2(I, recop + 2), 2))
                        If sens = "amont" Then
                            indi = j - 2
                        Else
                            indi = j
                        End If
                        If (left(t(liga, indi + 1), Len(t(liga, indi + 1)) - 3) = left(t(I, indi + 1), Len(t(I, indi + 1)) - 3)) Then
                            'Si l'autre ligne est juste au dessus, facile
                            If sens = "amont" Then
                                indi2 = j - 2
                            Else
                                indi2 = j
                            End If
                            For k = 0 To 3
                                t2(liga, recop + 1 + k) = t(liga, indi2 + k)
                                t(liga, indi2 + k) = "000000"
                            Next
                            If liga <= nbflux Then
                                t2(0, recop + 2) = t(0, indi + 1)
                            Else
                                t2(0, recop + 2) = t(0, indi2 + 1)
                            End If
                            kb = j - 1
                        Else
                            'Sinon, recherche
                            For k = 3 To size
                                If left(t(liga, k), Len(t(liga, k)) - 2) = left(t2(I, recop + 2), Len(t2(I, recop + 2)) - 2) Then

                                    'Copie des 4 valeurs trouvées
                                    For m = 1 To 4
                                        t2(liga, recop + m) = t(liga, k - 2 + m)
                                        t(liga, k - 2 + m) = "000000"
                                    Next
                                    If liga <= nbflux Then
                                        t2(0, recop + 2) = t(0, k)
                                    Else
                                        t2(0, recop + 2) = t(0, indi + 1)
                                    End If
                                    kb = k
                                    k = size
                                End If
                            Next
                        End If
                        If t2(liga, bigsize) = "" Or t2(liga, bigsize) = "0" Then
                            t2(liga, bigsize) = 1
                        Else
                            ok = False
                            erreur = "boucle"
                        End If
                        position_erreur = "recopie 2"

                        'Rechercher les HEX en amont de cette ligne
                        If ok = True Then
                            nouv = False
                            Call SearchAndDestroy(t(), t2(), size, bigsize, liga, "amont", kb, recop, ok, False)
                        Else
                            Exit Sub
                            j = fin
                        End If

                        'finalisation
                        If ok = True Then
                            If sens = "amont" Then
                                indi = j - 2
                            Else
                                indi = j
                            End If
                            For k = 0 To 3
                                t(I, indi + k) = "000000"   'indique que l'échangeur a été déplacé
                            Next
                            t(I, UBound(t(), 2)) = 1
                        End If
                    End If
                End If
            End If
        Next
        position_erreur = "recherche"

        'Affichage du HEN lors de son affinage
        '        Tabelle9.Activate
        '        Range("A1:DA100").ClearContents
        '        For cv = 0 To UBound(t())
        '            For cv2 = 1 To size
        '                Tabelle9.Cells(4 + cv, 2 + cv2).Value = Replace(data_flux_all_final(cv, cv2), separateur_decimal, ".")
        '            Next
        '        Next
    End If
    

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "64" & position_erreur

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub


Public Sub DisplayHEN(pi_FiliereConsideree As Integer, t() As String, size As Integer)
    'Affichage de la grille de flux sous forme graphique. Cette procédure est appelée dans la partie IV du pinch, en toute fin d'éxécution.

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String
    Dim position_erreur As String

    'step   numéro de la cilière concernée
    't      grille des flux (tous les flux : Process et Utilités)
    'size   taille de la grille de flux

    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim m As Integer
    Dim p As Integer
    Dim N As Integer

    Dim x As Integer                'coordonnée horizontale pour l'ajout d'éléments graphiques
    Dim y As Integer                'coordonnée verticale pour l'ajout d'éléments graphiques
    Dim Lligne As Integer           'longueur d'un trait horizontal (1 trait pour 1 flux)
    Dim Hligne As Integer           'hauteur des éléments posés sur une ligne (ovales, des boites, des températures ...)
    Dim sepligne As Integer         'espace entre 2 traits = hauteur d'une ligne + 20 (20: valeur arbitraire)
    Dim LargCel As Integer          'largeur d'un élément unitaire lu dans la grille de flux t()
    Dim posX As Integer             'position en x (horizontale) d'un élément
    Dim posX2 As Integer            'position de l'autre ovale, pour la création d'un lien entre 2 ovales
    Dim temp As String              'pour traitement de chaine de caractère quand lecture d'une puissance
    Dim nom_echangeur As String     'nom (utile pour tester)
    Dim surf As Double              'surface de l'échangeur
    Dim nbflux_caches As Integer    'nombre de flux cachés (au cas où il y a des flux non affichés, les liens entre ovales seront bien calibrés)
    Dim colonne_hex_1 As Integer    'colonne du hex sur le flux 1 de l'échangeur
    Dim colonne_hex_2 As Integer    'colonne du hex sur le flux 2 de l'échangeur
    Dim Ltotal As Single            'longueur totale de la grille de flux au format graphique
    Dim Htotal As Single            'hauteur totale de la grille de flux au format graphique
    Dim hactuel As Single           'longueur du graphique créé (sera ajustée)
    Dim lactuel As Single           'hauteur du graphique créé (sera ajustée)

    Dim graphique As ChartObject   'modif DCA



    'MODIF CRE 26/10/2012   On va créer les éléments graphiques dans la page "Results0". A toi de voir comment tu veux l'appeler (Tabelle9 ou CodeNames(...) ou ...).

    'Initialisation des variables globales
    position_erreur = "exportHEN"
    x = 50
    y = 50
    LargCel = 40
    Lligne = (size - 6 + 4 - 2) * LargCel / 2
    Hligne = 20
    sepligne = Hligne + 20

    'Suppression éventuelle de formes restantes
    Tabelle9.Activate
    Tabelle9.DrawingObjects.Select
    Selection.Delete

    'Nettoyage du tableau de flux (avec un Replace)
    For I = 0 To UBound(t())
        For j = 1 To size
            t(I, j) = Replace(t(I, j), ".", separateur_decimal)
        Next
    Next

    'Création d'un repère pour le Pinch 'MODIF CRE 17/01/2013   Déplacé avant pour que le repère Pinch soit en dessous des puissances, pour les HEx en travers
    Call GetCols(t())
    posX = x + LargCel  'initialisation de posX juste après la case avec le nom du flux
    For j = 3 To size - 7
        posX = posX + LargCel / 2   'calcul de posX à chaque itération sur les colonnes
        'If t(0, j) = "t_pinch_cen" Then
        If j = colpinch + 1 Then
            'Affichage du repère dès qu'on a trouvé la colonne du pinch
            ActiveSheet.Shapes.AddLine(posX + LargCel / 2, 10, posX + LargCel / 2, 10 + (nbflux + 0.4) * (sepligne)).Select    'en changeant "+ 0.1", on règle la fin du repère
            Selection.ShapeRange.Line.ForeColor.SchemeColor = 10
            Selection.ShapeRange.Line.Weight = 3

            'Affichage d'une boite tout en haut avec le mot "pincement" MODIF CRE 17/10/2012    ajout de ce code
            ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, posX - 0.1 * LargCel, 0, LargCel * 1.2, Hligne).Select
            Selection.HorizontalAlignment = xlCenter
            Selection.Characters.Text = "  Pincement  "
            Call DisplayHENElement("center", 2, "black", True)

            j = size - 7    'pour sortir de la boucle rapidement
        End If
    Next

    'Création des dessins
    For I = 1 To UBound(t())
        posX = x
        
        If t(I, 1) <> "" Then
            If Abs(CDbl(t(I, colreste))) > 1 Then    'MODIF CRE 01/10/2012   Affichage d'un flux ssi P > 1 kW (valeur à sortir dans l'interface ???)

                'Création d'une fleche dont la longueur est calculée plus haut : Lligne
                ActiveSheet.Shapes.AddLine(x, y, x + Lligne, y).Select

                'Création d'une boite avec le nom de chaque flux
                If t(I, 1) = "Ch" Or t(I, 1) = "UCh" Then
                    'Si le flux est chaud, on place la boite tout à gauche et elle fait 2*LargCel de large
                    Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
                    Selection.ShapeRange.Line.ForeColor.SchemeColor = 10
                    Selection.ShapeRange.Line.Weight = 1.25
                    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y - Hligne, LargCel * 2, Hligne).Select
                Else
                    'Si c'est un flux froid, on place la boite tout à gauche : x (initial) + Lligne - Largcel*2 (la boite fait 2*largCel de large)
                    Selection.ShapeRange.Line.BeginArrowheadStyle = msoArrowheadTriangle
                    Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
                    Selection.ShapeRange.Line.Weight = 1.25
                    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, x + Lligne - LargCel * 2, y - Hligne, LargCel * 2, Hligne).Select
                End If
                'Affichage du nom du flux
                Selection.Characters.Text = t(I, 2)
                Call DisplayHENElement("left", 1, "black", False)

                'La position x courante est maintenant décalée de LargCel (ajout du nom du flux)
                posX = posX + LargCel


                'itération sur les colonnes pour placer les échangeurs
                For j = 3 To size - 6

                    'A chaque itération, on décale la position actuelle de Largcel/2 (la grille de flux est architecturée avec une température, des tirets, une température, une puissance. pour éviter d'avoir un rendu graphique trop large, on ne décale que de largcel/2.
                    posX = posX + LargCel / 2

                    'Si on rencontre une puissance (échangée ou restante)
                    If left(t(I, j), 2) = "P_" Or left(t(I, j), 4) = "Hot_" Or left(t(I, j), 4) = "Col_" Or left(t(I, j), 2) = "R_" Then

                        'Création du nom de l'échangeur (si c'est un échangeur)
                        If left(t(I, j), 2) = "P_" Then
                            'MODIF CRE 17/01/2013   J'ai sorti ces lignes de codes du test if suivant pour que ça continue de bien fonctionner
                            nom_echangeur = t(0, j)
                            If left(t(0, j), 3) = "HAX" Then    'MODIF CRE 01/10/2012   Les échangeurs manuels sont créés avec le suffixe HAX_, dans la grille de flux t(). On change ici pour HEX_, pour l'affichage seulement
                                nom_echangeur = "HEX" & Right(nom_echangeur, Len(nom_echangeur) - 3)
                            End If
                                
                            If j <> colpinch + 1 Then 'affichage du nom de l'HEX une seule fois seulement   'MODIF CRE 17/01/2013   Modification du test : on n'affiche pas HEX_n seulement si présent sur colpinch+1 (échangeurs personnalisés à travers le pinch). Est-ce un problème que tous les noms soient affichés ?
                                ActiveSheet.Shapes.AddShape(msoShapeRectangle, posX, 0, LargCel, Hligne).Select
                                'nom_echangeur = t(0, j)
                                If nom_echangeur = "" Then
                                    nom_echangeur = "erreur"
                                Else
                                    'If left(t(0, j), 3) = "HAX" Then    'MODIF CRE 01/10/2012   Les échangeurs manuels sont créés avec le suffixe HAX_, dans la grille de flux t(). On change ici pour HEX_, pour l'affichage seulement
                                    '    nom_echangeur = "HEX" & Right(nom_echangeur, Len(nom_echangeur) - 3)
                                    'End If
                                    For k = 2 To UBound(data_hen(), 2)
                                        If data_hen(1, k) = nom_echangeur Then
                                            surf = Round(CDbl(data_hen(9, k)), 1)
                                            k = UBound(data_hen(), 2)
                                        End If
                                    Next
                                End If
                                Selection.Characters.Text = "#" & Right(nom_echangeur, Len(nom_echangeur) - 4) & Chr(10) & surf & " m²"
                                Selection.HorizontalAlignment = xlCenter
                                Selection.AutoSize = False
                                
                                'Affichage du nom
                                If left(t(0, j), 3) = "HAX" Then    'MODIF CRE 01/10/2012  Affichage des HEX manuels en gras
                                    Call DisplayHENElement("center", 2, "black", True)
                                Else
                                    Call DisplayHENElement("center", 1, "black", True)
                                End If
                            Else
                                'x = 1
                            End If
                        End If

                        'Création de la puissance (échangée ou restante) dans un ovale
                        ActiveSheet.Shapes.AddShape(msoShapeOval, posX, y - Hligne / 2, LargCel, Hligne).Select
                        'Nettoyage de la puissance (on enlève "P_" et le numéro de l'autre flux à la fin de P_xxx_xx)
                        If left(t(I, j), 2) = "P_" Then
                            temp = Right(t(I, j), Len(t(I, j)) - 2)
                            temp = left(temp, Len(temp) - 3)
                        ElseIf left(t(I, j), 2) = "R_" Then
                            temp = Right(t(I, j), Len(t(I, j)) - 2)
                        Else
                            temp = Right(t(I, j), Len(t(I, j)) - 4)
                        End If
                        temp = Round(CDbl(temp), 1)    'Arrondi de la puissance à une décimale
                        Selection.Characters.Text = temp & " kW"    'Ajout de l'unité
                        'Affichage de la puissance
                        If left(t(I, j), 2) = "P_" Then
                            Call DisplayHENElement("center", 1, "black", True)
                        ElseIf left(t(I, j), 4) = "Col_" Then
                            Call DisplayHENElement("center", 3, "blue", True)
                        ElseIf left(t(I, j), 4) = "Hot_" Then
                            Call DisplayHENElement("center", 3, "red", True)
                        Else
                            Call DisplayHENElement("center", 3, "gray", True)
                        End If

                        'Création de la liaison entre les ovales
                        'MODIF CRE 02/10/2012   Modification pour repérer plus efficacement les colonnes des HEX, vu qu'ils ont un nom unique
                        'MODIF CRE 24/10/2012   On ajoute "0" en tete, pour colonne_hex_2 et colonne_hex_1, pour éviter toute erreur. Au pire, ça fera un réseau d'échangeurs hideux, mais ça plantera pas.
                        If left(t(I, j), 2) = "P_" Then
                            For m = 2 To UBound(data_hen(), 2)
                                If data_hen(1, m) = nom_echangeur Then
                                    If t(I, 1) = "Ch" Then
                                        colonne_hex_2 = CInt(0 & data_hen(12, m))
                                        colonne_hex_1 = CInt(0 & data_hen(11, m))
                                    Else
                                        colonne_hex_2 = CInt(0 & data_hen(11, m))
                                        colonne_hex_1 = CInt(0 & data_hen(12, m))
                                    End If
                                    m = UBound(data_hen(), 2)
                                End If
                            Next
                            'la position de l'ovale 2 (et donc la fin du lien entre les ovales) est calculée à partir de posX (position de l'ovale déja tracé) et de la différence des colonnes de l'HEX (data_hen(11, ) et data_hen(12, )
                            posX2 = posX + LargCel / 2 * (colonne_hex_2 - colonne_hex_1)

                            k = CDbl(Right(t(I, j), 2))
                            nbflux_caches = 0
                            If I < k Then
                                'si des flux (situés entre les 2 flux concernés par la liaison en cours de création) ont été cachés (car puissance trop faible), la liaison entre les 2 ovales doit être adaptée
                                For p = I + 1 To k - 1
                                    If Abs(t(p, colreste)) < 0.001 Then
                                        nbflux_caches = nbflux_caches + 1
                                    End If
                                Next
                                'Affichage de la liaison
                                ActiveSheet.Shapes.AddLine(posX + LargCel / 2, y + Hligne / 2, posX2 + LargCel / 2, y - Hligne / 2 + (k - I - nbflux_caches) * sepligne).Select
                                Selection.ShapeRange.Line.ForeColor.RGB = RGB(0, 0, 0)
                            End If
                        End If

                        'Création de la température à gauche de la puissance
                        If t(I, j - 3) <> repere_traits_Clement And t(I, j - 3) <> "" And left(t(I, j - 3), 4) <> "Hot_" And left(t(I, j - 3), 4) <> "Col_" Then
                            'si 2 cases avant on a une température, on l'affiche à certaines conditions :
                            If (j = colpinch + 5 Or j = 6) And left(t(I, colpinch + 1), 2) <> "P_" Then 'MODIF CRE 17/01/2013   Ajout de ce test sur t(i,colpinch+1) pour gérer aussi les HEX personnalisés à travers du pinch
                                'si on est aux bornes (après tpinch ou tmax), on affiche directement la température à gauche (tpinch ou tmax)
                                ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, posX - LargCel, y - Hligne, LargCel, Hligne).Select
                                Selection.HorizontalAlignment = xlRight
                                Selection.Characters.Text = Round(t(I, j - 1), 1) & " °C"
                                Call DisplayHENElement("right", 0, "black", False)
                            Else
                                'on n'est pas aux bornes
                                If Abs(CDbl(t(I, j - 1)) - CDbl(t(I, j - 3))) > 0.000001 Then
                                    'si la température située 2 colonnes avant est différente, on affiche la température à gauche comme normal (en fait, ce cas de figure n'est jamais rencontré)
                                    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, posX - LargCel / 2, y - Hligne, LargCel, Hligne).Select
                                    Selection.HorizontalAlignment = xlRight
                                    Selection.Characters.Text = Round(t(I, j - 1), 1) & " °C"
                                    Call DisplayHENElement("right", 0, "black", False)
                                Else
                                    'la température située 2 cases avant est la meme : on n'affiche rien, seule la température à droite sera affichée, et sera centrée entre les puissances
                                End If
                            End If
                        Else
                            'si 2 cases avant on a pas de température (on a des tirets normalement) : affichage de la température à gauche, alignée à droite
                            ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, posX - LargCel, y - Hligne, LargCel, Hligne).Select
                            Selection.HorizontalAlignment = xlRight
                            Selection.Characters.Text = Round(t(I, j - 1), 1) & " °C"
                            Call DisplayHENElement("right", 0, "black", False)
                        End If

                        'Création de la température à droite de la puissance
                        If t(I, j + 3) <> repere_traits_Clement And t(I, j + 3) <> "" Then
                            'si on a une température 2 cases après, on l'affiche à certaines conditions :
                            If (Abs(CDbl(t(I, j + 1)) - CDbl(t(I, j + 3)))) < 0.000001 Then
                                'si c'est la meme température
                                If j + 1 = colreste - 4 Or j + 1 = colpinch - 2 Then
                                    'si on est aux bornes : tmin ou avant pinch, affichage de température à droite à coté de la Puissance
                                    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, posX + LargCel, y - Hligne, LargCel, Hligne).Select
                                    Selection.HorizontalAlignment = xlLeft
                                    Selection.Characters.Text = Round(t(I, j + 1), 1) & " °C"
                                    Call DisplayHENElement("left", 0, "black", False)
                                Else
                                    'pas aux bornes : affichage de la température au centre des puissances
                                    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, posX + LargCel, y - Hligne, LargCel, Hligne).Select
                                    Selection.HorizontalAlignment = xlCenter
                                    Selection.Characters.Text = Round(t(I, j + 1), 1) & " °C"
                                    Call DisplayHENElement("center", 0, "black", False)
                                End If
                            End If
                        Else
                            'on a pas de température : affichage de la température à droite juste à coté de Puissance, et alignée à gauche
                            ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, posX + LargCel, y - Hligne, LargCel, Hligne).Select
                            Selection.HorizontalAlignment = xlLeft
                            Selection.Characters.Text = Round(t(I, j + 1), 1) & " °C"
                            Call DisplayHENElement("left", 0, "black", False)
                        End If

                    End If

                Next
            Else
                'fin de If Abs(CDbl(t(i, colreste))) > 1 Then => si la puissance de ce flux est trop faible, on n'a pas affiché ce flux : il faut réduire le décalage vertical
                y = y - sepligne
            End If
        End If

        'décalage vertical à chaque ajout de flux
        y = y + sepligne
    Next

    'Sélection de tous les éléments graphiques créés
    ActiveSheet.Shapes.SelectAll
    Selection.ShapeRange.Group.Name = "Group99"    'nom arbitraire, ne sert que temporairement
    Ltotal = ActiveSheet.Shapes("Group99").Width
    Htotal = ActiveSheet.Shapes("Group99").Height

    'Ajout d'un graphique
    'Tabelle9.Activate
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Activate
    Set graphique = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects("HENGraphique")
'    Set graphique = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects.Add(15, 20, 20, 98)

    If graphique.Chart.SeriesCollection.Count > 0 Then  'Suppression des séries parasites si présentes
        Do While graphique.Chart.SeriesCollection.Count > 0
            graphique.Chart.SeriesCollection(1).Delete
        Loop
    End If
    '    If pi_EtapeCalculConsideree = 1 Then   'différenciation des calculs "guide" des calculs "normaux"
    '        graphique.Chart.Location xlLocationAsObject, Name:="PINCH " & pi_FiliereConsideree & " guide"
    '    Else
    '        graphique.Chart.Location xlLocationAsObject, Name:="PINCH " & pi_FiliereConsideree
    '    End If
'    graphique.Chart.Parent.Name = "HENGraphique"
    'Ajustement des dimensions du graphique
    hactuel = ActiveSheet.Shapes("HENGraphique").Height
    lactuel = ActiveSheet.Shapes("HENGraphique").Width
    ActiveSheet.Shapes("HENGraphique").ScaleWidth Ltotal / lactuel, msoFalse, msoScaleFromTopLeft
    ActiveSheet.Shapes("HENGraphique").ScaleHeight Htotal / hactuel * 1.05, msoFalse, msoScaleFromTopLeft   'MODIF CRE 17/01/2013   Ajout du facteur 1.05 pour du confort visuel
    
    'MODIF CRE 16/10/2012   Affichage dans PINCH X
    ActiveSheet.Shapes("HENGraphique").top = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(75, 3).top
    ActiveSheet.Shapes("HENGraphique").left = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(75, 3).left

    'Insertion des dessins dans le graphique et export sous forme d'image
    Tabelle9.Activate
    ActiveSheet.Shapes("Group99").Select
    Selection.Copy
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Activate
    ActiveSheet.Shapes("HENGraphique").Select
    graphique.Chart.Paste
    'graphique.Chart.Export FileName:="VeriPinchHEN  (pi_FiliereConsideree " & pi_FiliereConsideree & ").jpg", FilterName:="JPEG"    'on peut exporter le graphique sous forme d'image   'MODIF CRE 20/09/2012   Export du graphique sous forme d'image rendu inactif

    'Finalisation
    ActiveSheet.Range("A1").Select
    Tabelle9.Activate
    ActiveSheet.Shapes("Group99").Select
    Selection.Delete
    
    Exit Sub
    '*********************************************************************
erreur_non_geree:

    NumeroErreurNonGeree = "65" & position_erreur

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

'fin modif MSM 03/12/12


Public Sub DisplayHENElement(align, thickness, color, fill)
    'Mise en forme d'un élément graphique (alignement du texte, couleur et épaisseur de la bordure ...)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String
    Dim position_erreur As String

    'align          alignement : "center", "left", "right"
    'thickness      épaisseur du trait (integer)
    'color          couleur : "blue", "red", "gray"
    'fill           remplissage : True (remplissage blanc) / False (pas de remplissage)

    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Normal"
        .size = 8
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With

    With Selection
        .VerticalAlignment = xlCenter
        .orientation = xlHorizontal
        .AutoSize = False
    End With

    With Selection.ShapeRange.fill
        .Solid
        .ForeColor.SchemeColor = 9
        .Transparency = 0
    End With

    With Selection.ShapeRange.Line
        .Weight = thickness
        .DashStyle = msoLineSolid
        .Style = msoLineSingle
        .Transparency = 0
        .Visible = msoTrue
        .ForeColor.SchemeColor = 8
        .BackColor.RGB = RGB(255, 255, 255)
    End With

    'Bordure
    If thickness = 0 Then
        Selection.ShapeRange.Line.Visible = msoFalse
    Else
        Selection.ShapeRange.Line.Visible = msoTrue
        Selection.ShapeRange.Line.Weight = thickness
    End If

    'Alignement horizontal
    If align = "center" Then
        Selection.HorizontalAlignment = xlCenter
    ElseIf align = "left" Then
        Selection.HorizontalAlignment = xlLeft
    Else
        Selection.HorizontalAlignment = xlRight
    End If

    'Couleur du trait
    If color = "black" Then
        Selection.ShapeRange.Line.ForeColor.SchemeColor = 0
    ElseIf color = "red" Then
        Selection.ShapeRange.Line.ForeColor.SchemeColor = 10
    ElseIf color = "blue" Then
        Selection.ShapeRange.Line.ForeColor.SchemeColor = 12
    ElseIf color = "gray" Then
        Selection.ShapeRange.Line.ForeColor.RGB = RGB(128, 128, 128)
    End If

    'Remplissage
    If fill = True Then
        Selection.ShapeRange.fill.Visible = msoTrue
    Else
        Selection.ShapeRange.fill.Visible = msoFalse
    End If

    'MODIF CRE 10/10/2012   Test de la version d'Excel
    If Application.Version = "14.0" Then
        Selection.ShapeRange.TextFrame2.WordWrap = msoFalse    'MODIF MSM 01/10/2012   Besoin de commenter car ça plante chez Martin apparemment
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "66" & position_erreur

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Private Function arrondi_sup_borne_echelle(ByVal parametre As Double) As Double
    'Ajouté par DCA   ne marche que pour des valeurs positives

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    Dim entier_arrondi As Long

    arrondi_sup_borne_echelle = parametre

    If arrondi_sup_borne_echelle > 0 Then
        If arrondi_sup_borne_echelle >= 1 Then
            I = 0
            Do Until arrondi_sup_borne_echelle >= 10 ^ I And arrondi_sup_borne_echelle < 10 ^ (I + 1)
                I = I + 1
            Loop
            entier_arrondi = arrondi_sup_borne_echelle / (10 ^ I)
            If entier_arrondi < arrondi_sup_borne_echelle / (10 ^ I) Then
                entier_arrondi = entier_arrondi + 1
            End If
            If arrondi_sup_borne_echelle < occupation_max_courbes * entier_arrondi * (10 ^ I) Then
                arrondi_sup_borne_echelle = entier_arrondi * (10 ^ I)
            Else
                arrondi_sup_borne_echelle = entier_arrondi * (10 ^ I) / occupation_max_courbes
            End If
        Else
            arrondi_sup_borne_echelle = 1
        End If
    End If




    Exit Function
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.16"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Function

Private Function arrondi_inf_borne_echelle(ByVal parametre As Double) As Double
    'Ajouté par DCA   ne marche que pour des valeurs positives

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    Dim entier_arrondi As Long

    arrondi_inf_borne_echelle = parametre

    If arrondi_inf_borne_echelle >= 1 Then
        I = 0
        Do Until arrondi_inf_borne_echelle >= 10 ^ I And arrondi_inf_borne_echelle < 10 ^ (I + 1)
            I = I + 1
        Loop
        entier_arrondi = arrondi_inf_borne_echelle / (10 ^ I)
        If entier_arrondi > arrondi_inf_borne_echelle / (10 ^ I) Then
            entier_arrondi = entier_arrondi - 1
        End If
        arrondi_inf_borne_echelle = entier_arrondi * (10 ^ I)
    Else
        arrondi_inf_borne_echelle = 0
    End If





    Exit Function
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.17"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Function

Private Sub ChartDisplay(ByVal nom As String, top As Double, left As Double, ByRef DCA_CC_avant_translation As Boolean, Optional x1 As String, Optional y1 As String, Optional name1 As String, Optional col1 As Integer, Optional x2 As String, Optional y2 As String, Optional name2 As String, Optional col2 As Integer)
    'Mise en forme d'un graphique pour les Courbes Composites (ajout des séries, des titres, positionnement ...)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'nom        nom du graphique à créer
    'titre      titre affiché
    'top        position verticale
    'left       position horizontale
    'step       filière considérée
    'x1         coordonnées X des points
    'y1         coordonnées Y des points
    'nom1       nom de la série

    Dim graphique As ChartObject
    

    Dim nbseries As Integer
    nbseries = 0

    'Ajout du graphique
'    Charts.Add
    Set graphique = CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects(nom)
'    graphique.Activate
    graphique.Visible = True

    'Suppression des séries parasites si présentes 'modifié par DCA
    Do While graphique.Chart.SeriesCollection.Count > 0
        graphique.Chart.SeriesCollection(1).Delete
    Loop

    graphique.Chart.ChartType = xlXYScatterLines
    graphique.Chart.Location Where:=xlLocationAsObject, Name:=CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Name
    graphique.Chart.Parent.Name = nom
    graphique.Chart.HasLegend = False
    nbseries = 0

    'Affichage de la première série
    If x1 <> "" Then
        graphique.Chart.SeriesCollection.NewSeries
        nbseries = nbseries + 1
        With graphique.Chart.SeriesCollection(nbseries)
            .XValues = x1
            .Values = y1
            .Name = name1
            .Border.ColorIndex = col1
            .MarkerBackgroundColorIndex = col1
            .MarkerForegroundColorIndex = col1
            .MarkerStyle = 2
            If PINCH_DANS_OCEAN = True Then
                .MarkerSize = DCA_CC_TaillePoints
            Else
                .MarkerSize = 4
            End If
            '.Format.Line.Weight = 2
        End With
    End If

    'Affichage de la deuxième série
    If x2 <> "" Then
        graphique.Chart.SeriesCollection.NewSeries
        nbseries = nbseries + 1
        With graphique.Chart.SeriesCollection(nbseries)
            .XValues = x2
            .Values = y2
            .Name = name2
            .Border.ColorIndex = col2
            .MarkerBackgroundColorIndex = col2
            .MarkerForegroundColorIndex = col2
            .MarkerStyle = 2
            If PINCH_DANS_OCEAN = True Then
                .MarkerSize = DCA_CC_TaillePoints
            Else
                .MarkerSize = 4
            End If
            '.Format.Line.Weight = 2
        End With
    End If

    'Ajout du titre
    If PINCH_DANS_OCEAN = False Then
        graphique.Chart.HasTitle = True
        With graphique.Chart.ChartTitle
            '.Characters.Text = titre
            .AutoScaleFont = False
            .Font.FontStyle = "Gras"
            .Font.size = 11
            .AutoScaleFont = False
        End With
    Else
        If x1 <> "" Or x2 <> "" Then
            graphique.Chart.HasTitle = True
            With graphique.Chart.ChartTitle
                If DCA_CC_avant_translation = True Then
                    .Characters.Text = Feuil63.Range("PINCH_CC_avant_translation_titre").Cells(1, 1)
                Else
                    .Characters.Text = Feuil63.Range("PINCH_CC_apres_translation_titre").Cells(1, 1)
                End If
                .AutoScaleFont = False
                .Font.FontStyle = "Gras"
                .Font.size = DCA_CC_Titre_TaillePolice
                .AutoScaleFont = False
            End With
        End If
    End If

    'Mise en forme du reste
    If x1 <> "" Or x2 <> "" Then
        graphique.Chart.PlotArea.Interior.ColorIndex = 2
        graphique.Chart.Axes(xlCategory).HasMajorGridlines = True
        graphique.Chart.Axes(xlCategory).MajorGridlines.Border.LineStyle = xlGray50
        graphique.Chart.Axes(xlCategory).TickLabels.AutoScaleFont = False
        graphique.Chart.Axes(xlCategory).MinimumScale = graph_abscisse_min
        graphique.Chart.Axes(xlCategory).MaximumScale = graph_abscisse_max
        graphique.Chart.Axes(xlValue).HasMajorGridlines = True
        graphique.Chart.Axes(xlValue).MajorGridlines.Border.LineStyle = xlGray50
        graphique.Chart.Axes(xlValue).TickLabels.AutoScaleFont = False
        graphique.Chart.Axes(xlValue).MinimumScale = graph_ordonnee_min
        graphique.Chart.Axes(xlValue).MaximumScale = graph_ordonnee_max
        graphique.Chart.Axes(xlCategory, xlPrimary).HasTitle = True
        graphique.Chart.Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = CodeNames("Feuil5").Range("l_deltaEnthalpie")
        graphique.Chart.Axes(xlCategory, xlPrimary).AxisTitle.AutoScaleFont = False
        graphique.Chart.Axes(xlValue, xlPrimary).HasTitle = True
        graphique.Chart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = CodeNames("Feuil5").Range("l_valeurPincement")
        graphique.Chart.Axes(xlValue, xlPrimary).AxisTitle.AutoScaleFont = False
        If PINCH_DANS_OCEAN = True Then
            graphique.Chart.Axes(xlCategory, xlPrimary).AxisTitle.Font.size = DCA_CC_Legende_TaillePolice
                graphique.Chart.Axes(xlCategory, xlPrimary).TickLabels.Font.size = DCA_CC_Legende_TaillePolice
            graphique.Chart.Axes(xlValue, xlPrimary).AxisTitle.Font.size = DCA_CC_Legende_TaillePolice
                graphique.Chart.Axes(xlValue, xlPrimary).TickLabels.Font.size = DCA_CC_Legende_TaillePolice
            graphique.Chart.HasLegend = True
            graphique.Chart.Legend.Position = xlLegendPositionBottom
            graphique.Chart.Legend.Font.size = DCA_CC_Etiquette_TaillePolice
        Else
            graphique.Chart.Legend.Position = xlTop
        End If
        graphique.Chart.Legend.AutoScaleFont = False
    End If
    
    'Positionnement du graphique.chart
    ActiveSheet.Shapes(nom).top = top
    ActiveSheet.Shapes(nom).left = left
    If DCA_CC_avant_translation = True Then
        ActiveSheet.Shapes(nom).Width = graph_CC_sans_decalage_largeur_points
        ActiveSheet.Shapes(nom).Height = graph_CC_sans_decalage_hauteur_points
    Else
        ActiveSheet.Shapes(nom).Width = graph_CC_avec_decalage_largeur_points
        ActiveSheet.Shapes(nom).Height = graph_CC_avec_decalage_hauteur_points
    End If
    
    CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Activate
    Range("A1").Select

    

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.13"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub



Public Sub Sort_Array(t() As Double, sens As String, cols As Integer)
    'Tri d'un tableau de flux (utilisé pour trier selon les températures) - http://faq.vb.free.fr/index.php?question=151#BubbleSort

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    't      tableau de variable à trier
    'sens   sens du tri ("asc" ou "dsc")
    'cols   nombre de colonnes concernées par le tri (nombre de colonnes de l'Array)

    Dim I As Double
    Dim j As Double
    Dim k As Double
    Dim tmp As Double   'valeur tampon
    Dim TEST As Double  'valeur test indicateur de sortie de boucle

    k = UBound(t())
    Do
        TEST = 0
        For I = 1 To k - 1
            If sens = "asc" Then
                If t(I, 1) > t(I + 1, 1) Then
                    TEST = I
                    For j = 1 To cols
                        tmp = t(I, j)
                        t(I, j) = t(I + 1, j)
                        t(I + 1, j) = tmp
                    Next
                End If
            Else
                If t(I, 1) < t(I + 1, 1) Then
                    TEST = I
                    For j = 1 To cols
                        tmp = t(I, j)
                        t(I, j) = t(I + 1, j)
                        t(I + 1, j) = tmp
                    Next
                End If
            End If
        Next I
        k = TEST    ' increase lower bound, improve performances
    Loop While TEST

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.14"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Public Sub Delete_row(t() As Double, delrow As Integer)
    'Suppression d'une ligne dans un tableau (utilisé dans la gestion des doublons dans la construction des CC)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    't          tableau de variables concerné
    'delrow     ligne à supprimer

    Dim t_tmp() As Double
    Dim I As Integer

    ReDim t_tmp(1 To UBound(t()) - 1, 1 To 3)
    For I = delrow To UBound(t()) - 1    'à partir du rang à supprimer, décalage des données vers le haut
        t(I, 1) = t(I + 1, 1)
    Next
    For I = 1 To UBound(t()) - 1    'stockage dans un tableau tampon
        t_tmp(I, 1) = t(I, 1)
    Next
    ReDim t(1 To UBound(t()) - 1, 1 To 3)
    t = t_tmp    'réatribution après redim

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.9"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Public Sub FindPinch(ta() As Double, tb() As Double, pinch_ch As Double, pinch_fr As Double, pinch_mesure As Double, h_ch As Double, h_fr As Double)
    'Trouve le pincement minimum dans un tableau de flux. On parcourt les flux chauds, puis les flux froids.

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'ta             tableau de variables contenant les flux chauds
    'tb             tableau de variables contenant les flux froids
    'pinch_ch       valeur du pinch trouvée sur les flux chauds
    'pinch_fr       valeur du pinch trouvée sur les flux froids
    'pinch_mesure   valeur du pinch minimal rencontré
    'h_ch           enthalpie du point de pinch minimal rencontré (flux chaud)
    'h_fr           enthalpie du point de pinch minimal rencontré (flux froid)

    Dim I As Integer

    pinch_mesure = 1000    'initialisation du DT_min à 1000
    For I = 1 To UBound(ta())
        If ta(I, 3) <= pinch_mesure Then
            pinch_mesure = ta(I, 3)
            pinch_ch = ta(I, 1)
            pinch_fr = -1000    'valeur fausse introduite à cet endroit pour déterminer sur quelle CC le pinch a été trouvé
            h_ch = ta(I, 2)
        End If
    Next
    For I = 1 To UBound(tb())
        If tb(I, 3) <= pinch_mesure Then
            pinch_mesure = tb(I, 3)
            pinch_fr = tb(I, 1)
            pinch_ch = -1000    'valeur fausse introduite à cet endroit pour déterminer sur quelle CC le pinch a été trouvé
            h_fr = tb(I, 2)
        End If
    Next

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.3"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Public Sub VerifierFluxSuppr(nomch As String, nomfr As String)
    'Lors de l'ajout d'un HEX, vérification avec la liste des HEX interdits

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'nomch      nom du flux chaud à tester
    'nomfr      nom du flux froid à tester

    Dim I As Integer
    Dim nomchsuppr As String        'nom du flux ch à interdire
    Dim nomfrsuppr As String        'nom du flux fr à interdire

    If nbhexforbidden > 0 Then
        For I = 1 To nbhexforbidden
            nomchsuppr = data_hex_forbidden(1, I)
            nomfrsuppr = data_hex_forbidden(2, I)
            If nomchsuppr = nomch And nomfrsuppr = nomfr Then
                ok = False
                erreur = erreur & " - hex supprimé"
                I = nbhexforbidden
            End If
        Next
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.2"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Public Sub GetCols(t() As String)
    'Récupère l'indice des colonnes caractéristiques du problème (à appeler à chaque ajout de HEX)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    I = 1
    Do While t(0, I) <> "t_pinch_gch"
        I = I + 1
    Loop
    colpinch = I
    Do While t(0, I) <> "Dh"
        I = I + 1
    Loop
    colreste = I

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.7"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Public Sub InsertCol(t() As String, colinsert As String, size As Integer)
    'Insère 4 colonnes au point indiqué pour insérer un HEX (avec ses températures)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    't          tableau des flux où on insère les colonnes
    'colinsert  indice de la colonne où insérer
    'size       nombre de colonnes du tableau

    Dim I As Integer
    Dim j As Integer
    Dim t_tmp() As String           'grille des flux temporaire
    Dim x As Integer                'position de la colonne à insérer

    ReDim Preserve t(LBound(t) To UBound(t), 1 To size + 4)
    ReDim t_tmp(LBound(t) To UBound(t), 1 To size + 4)

    x = Asc(colinsert) - 64
    For I = LBound(t) To UBound(t)
        For j = 1 To size
            If j < x Then
                t_tmp(I, j) = t(I, j)
            Else
                t_tmp(I, j + 4) = t(I, j)
                If I > 0 And j = x Then
                    If t(I, j - 1) = repere_traits_Clement Or t(I, j) = repere_traits_Clement Then
                        t_tmp(I, j) = repere_traits_Clement
                        t_tmp(I, j + 1) = repere_traits_Clement
                        t_tmp(I, j + 2) = repere_traits_Clement
                        t_tmp(I, j + 3) = repere_traits_Clement
                    End If
                End If
            End If
        Next
    Next
    t = t_tmp
    size = size + 4

    'Mise à jour des indices des colonnes   'MODIF CRE 02/10/2012   Ajout de cette mise à jour quand on insert des colonnes
    'MODIF CRE 24/10/2012   Inutile ici, sera renseigné dans IV.4
    '    For i = 2 To UBound(data_hen(), 2)
    '        If CInt(data_hen(11, i)) > x Then
    '            data_hen(11, i) = data_hen(11, i) + 4
    '        End If
    '        If CInt(data_hen(12, i)) > x Then
    '            data_hen(12, i) = data_hen(12, i) + 4
    '        End If
    '    Next

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.4"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Public Sub HENBuilding(source() As String, nbfluxsource As Integer, cot1 As Integer, Target() As String, nbfluxtarget As Integer, cot2 As Integer, pinch_mesure As Double, data_flux_hen_size As Integer)
    'Routine qui construit les HEX automatiquement entre 2 liste de flux (on balaie les flux de "source" à la recherche d'un vide dans la suite des températures et d'une puissance restante non nulle, puis on balaie les flux "target" pour trouver un match. Si on l'a trouvé, on crée le HEX
    'En fait, target et source sont identiques: c'est la grille des flux Process

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'source             grille des flux "source"
    'nbfluxsource       nombre de flux "source"
    'cot1               coté du flux "source" (4 à gauche, 5 à droite du pinch)
    'target             grille des flux "target"
    'nbfluxtarget       nombre de flux "target"
    'cot2               coté du flux "target" (4 à gauche, 5 à droite du pinch)
    'pinch_mesure       valeur du pinch
    'data_flux_hen_size taille des grilles de flux "target" et "source"

    Dim I As Integer
    Dim j As Integer
    Dim iter As Integer
    Dim cotch As Integer
    Dim cotfr As Integer
    Dim pot1 As Integer
    Dim pot2 As Integer

    ok = True
    erreur = ""

    pot1 = cot1
    pot2 = cot2
    If cot1 = 10 Then
        pot1 = cot2
    End If
    If cot2 = 10 Then
        pot2 = cot1
    End If

    For I = 1 To nbfluxsource
        'Si puissance restante à couvrir non nulle
        If source(I, colreste + pot1) <> "" Then
            If CDbl(source(I, colreste + pot1)) > 0.1 Then

                'Tant que la puissance est non nulle et qu'on n'a pas trop itéré
                iter = 0    'MODIF CRE 21/09/2012   oubli néfaste de cette réinitialisation qui empéchait de traiter TOUS les flux
                Do While source(I, colreste + pot1) > 0.1 And iter < 10

                    'Balayage de tous les flux de la target avec vérification de la règle des MCp
                    For j = 1 To nbfluxtarget

                        'si leur tipe est différent
                        If source(I, 1) <> Target(j, 1) Then
                            'si puissance restante non nulle
                            If Target(j, colreste + pot2) <> "" Then
                                If CDbl(Target(j, colreste + pot2)) > 0.1 Then

                                    'Si toutes les conditions sont là (tipe =!, puissance)
                                    If ok = True Then
                                        If source(I, 1) = "Ch" Then
                                            cotch = cot1
                                            cotfr = cot2
                                        Else
                                            cotfr = cot1
                                            cotch = cot2
                                        End If
                                    End If

                                    If ok = True Then
                                        If source(I, 1) = "Ch" Then
                                            Call VerifierFluxSuppr(source(I, 2), Target(j, 2))
                                        Else
                                            Call VerifierFluxSuppr(Target(j, 2), source(I, 2))
                                        End If
                                    End If
                                    If ok = True Then
                                        If source(I, 1) = "Ch" Then
                                            Call CheckData(source(), I, cotch, Target(), j, cotfr, data_flux_hen_size, pinch_mesure, 2)
                                        Else
                                            Call CheckData(Target(), j, cotch, source(), I, cotfr, data_flux_hen_size, pinch_mesure, 2)
                                        End If
                                    End If
                                    If ok = True Then
                                        nbhex = nbhex + 1
                                        If source(I, 1) = "Ch" Then
                                            Call CreerHEX2(source(), I, pot1, colplacehexch, Target(), j, pot2, colplacehexfr, 2)
                                        Else
                                            Call CreerHEX2(Target(), j, pot2, colplacehexch, source(), I, pot1, colplacehexfr, 2)
                                        End If
                                    End If

                                    ok = True
                                    erreur = ""
                                End If
                            End If
                        End If

                        'Sortie de la 2eme boucle si la puissance est déja couverte
                        If source(I, colreste + pot1) < 0.1 Then
                            j = nbfluxtarget
                        End If

                    Next
                    iter = iter + 1
                Loop

            End If
        End If
    Next

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.6"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Public Sub CheckData(chaud() As String, ligch As Integer, cot1 As Integer, froid() As String, ligfr As Integer, cot2 As Integer, data_flux_hen_size As Integer, pinch_mesure As Double, forced As Integer)
    'Vérification des données avant la création d'un échangeur : températures des flux concernés par l'échangeur (dans les bornes), de la puissance échangée (positive), d'un éventuel pincement ...
    'Et calcul des colonnes d'insertion puis insertion des colonnes

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String
    Dim position_erreur As String

    'chaud                  grille des flux qui contient le flux chaud concerné
    'ligch                  numéro du flux chaud
    'cot1                   coté du flux chaud (4 à gauche, 5 à droite du pinch)
    'froid                  grille des flux qui contient le flux froid concerné
    'ligfr                  numéro du flux froid
    'cot2                   coté du flux froid (4 à gauche, 5 à droite du pinch)
    'data_flux_hen_size     taille de la grille de flux (pour l'insertion)
    'pinch_mesure           valeur du pinch (pour le calcul du pincement)
    'forced                 échangeur forcé ou pas : 1 (échangeur Process/Process manuel), 2 (échangeur Process/Process automatique), 3 (échangeur Process/Utilité)

    Dim j As Integer
    Dim k As Integer
    Dim colmaxch As Integer
    Dim colminch As Integer
    Dim colmaxfr As Integer
    Dim colminfr As Integer
    Dim pch As Double
    Dim pfr As Double
    Dim mcpch As Double
    Dim mcpfr As Double
    Dim deb1 As Integer
    Dim deb2 As Integer
    Dim tic As Double
    Dim toc As Double
    Dim tif As Double
    Dim tof As Double
    Dim deb As Integer
    Dim fin As Integer
    Dim puissance_2 As Double
    Dim seuilPuissance As Double    'seuil de puissance pour la création d'un HEX
    Dim colonneStation As Integer   'colonne pour la lecture
    Dim LigneTab As Integer         'ligne pour la lecture
    Dim jsave As Integer

    'position de l'erreur : forced (1, 2 ou 3), + les lignes des flux chaud et froid
    position_erreur = forced & "_" & ligch & "_" & ligfr

    'Sauvegarde des bornes si déja calculées
    If forced = 1 Or forced = 3 Then
        tic = tinch
        toc = toutch
        tif = tinfr
        tof = toutfr
    End If
    tinch = 0
    toutch = 0
    tinfr = 0
    toutfr = 0
    colmaxch = -1
    colminch = -1
    colmaxfr = -1
    colminfr = -1

    'Initialisation des bornes extrèmes (deb1 et deb2) en fonction du coté (cot1 et cot2), uniquement pour les HEX automatiques
    If forced = 2 Then
        If cot1 = 4 Then
            deb1 = 3
        Else
            deb1 = colpinch + 2
        End If
        If cot2 = 4 Then
            deb2 = 3
        Else
            deb2 = colpinch + 2
        End If
    End If

    'MODIF CRE 10/10/2012   Ajout de ce code pour l'ajout à la main des HEX. Plante sans ça
    'Forced=1 : pour l'ajout de HEX à la main
    If forced = 1 Then

        'Retour des valeurs sauvegardées
        tinch = tic
        toutch = toc
        tinfr = tif
        toutfr = tof

        'Calcul des bornes extrèmes
        If cot1 = 5 Then
            deb = colpinch + 2
            fin = colreste - 2
        ElseIf cot1 = 4 Then
            deb = 3
            fin = colpinch
        'Else
        '    deb = 3
        '    fin = colreste - 2
        End If

        'Calcul des intervalles possibles
        If cot1 = 45 Then
            'Si l'échangeur traverse le pinch   'MODIF CRE 17/01/2013   gestion des HEX à travers
            colmaxch = colpinch
            colminch = colpinch + 2
        Else
            'sinon, on cherche l'intervalle dispo
            If fin - deb = 2 Then
                colmaxch = deb
                colminch = fin
            Else
                jsave = 0
                k = 0
                For j = deb To fin
                    If chaud(ligch, j) = "" Then
                        If chaud(ligch, j - 1) >= tinch Then
                            jsave = j
                        Else
                            k = j - 1
                            j = fin
                        End If
                    End If
                Next
                If k = 0 Then
                    k = jsave - 1
                End If
                colmaxch = k
    
                For j = k + 1 To fin
                    If chaud(ligch, j) <> "" Then
                        k = j
                        j = fin
                    End If
                Next
                colminch = k
            End If
        End If

        'Calcul des bornes extrèmes
        If cot2 = 5 Then
            deb = colpinch + 2
            fin = colreste - 2
        ElseIf cot2 = 4 Then
            deb = 3
            fin = colpinch
        'Else
        '    deb = 3
        '    fin = colreste - 2
        End If

        'Calcul des intervalles possibles
        If cot2 = 45 Then
            'Si l'échangeur traverse le pinch   'MODIF CRE 17/01/2013   gestion des HEX à travers
            colmaxfr = colpinch
            colminfr = colpinch + 2
        Else
            If fin - deb = 2 Then
                colmaxfr = deb
                colminfr = fin
            Else
                jsave = 0
                k = 0
                For j = deb To fin
                    If froid(ligfr, j) = "" Then
                        If froid(ligfr, j - 1) > tinch Then
                            jsave = j
                        Else
                            k = j - 1
                            j = fin
                        End If
                    End If
                Next
                If k = 0 Then
                    k = jsave - 1
                End If
                colmaxfr = k
    
                For j = k + 1 To fin
                    If froid(ligfr, j) <> "" Then
                        k = j
                    End If
                Next
                colminfr = k
            End If
        End If
    End If


    'Forced=2 : pour la création de HEX automatiquement. Pour chaque flux, recherche du premier créneau disponible et test sur la création du HEX.
    If forced = 2 Then

        'Détermination des températures et des colonnes du flux chaud
        j = 1
        Do While chaud(ligch, deb1 + j) <> "" And deb1 + j < data_flux_hen_size - 4
            j = j + 1
        Loop
        If (cot1 = 4 And deb1 + j - 1 < colpinch) Or (cot1 = 5 And deb1 + j - 1 < colreste) Then
            tinch = chaud(ligch, deb1 + j - 1)
            colmaxch = deb1 + j - 1
            Do While chaud(ligch, deb1 + j) = ""
                j = j + 1
            Loop
            If (cot1 = 4 And deb1 + j - 1 < colpinch) Or (cot1 = 5 And deb1 + j - 1 < colreste) Then
                toutch = chaud(ligch, deb1 + j)
                colminch = deb1 + j
            Else
                ok = False
            End If
        Else
            ok = False
        End If

        'Détermination des températures et des colonnes du flux froid
        j = 1
        Do While froid(ligfr, deb2 + j) <> "" And j + deb2 < data_flux_hen_size - 4
            j = j + 1
        Loop
        If (cot2 = 4 And deb2 + j < colpinch) Or (cot2 = 5 And deb2 + j < colreste) Then
            toutfr = froid(ligfr, deb2 + j - 1)
            colmaxfr = deb2 + j - 1
            Do While froid(ligfr, deb2 + j) = ""
                j = j + 1
            Loop
            If (cot2 = 4 And deb2 + j - 1 < colpinch) Or (cot2 = 5 And deb2 + j - 1 < colreste) Then
                tinfr = froid(ligfr, deb2 + j)
                colminfr = deb2 + j
            Else
                ok = False
            End If
        Else
            ok = False
        End If

    End If  'fin Forced = 2

    'Forced = 3 : Ajout des HEX manuels et Couverture des besoins
    If forced = 3 Then

        'Retour des valeurs sauvegardées
        tinch = tic
        toutch = toc
        tinfr = tif
        toutfr = tof

        'Calcul des bornes extrèmes
        If cot1 = 5 Then
            deb = colpinch + 2
            fin = colreste - 2
        ElseIf cot1 = 4 Then
            deb = 3
            fin = colpinch
        Else
            deb = 3
            fin = colreste - 2
        End If

        'Calcul de tinch
        If cot1 = 10 Then
            'Utilité : on itère jusqu'à trouver un créneau de libre 'MODIF CRE 02/10/2012   Mise à jour de l'identification de colmax et colmin
            j = 1
            Do While chaud(ligch, 2 + j) <> ""
                j = j + 1
            Loop
            j = j - 1
            colmaxch = 2 + j
            tinch = chaud(ligch, colmaxch)
        Else
            'Process : itération jusqu'à trouver un créneau libre dans l'intervalle de températures souhaitées
            j = deb
            k = j
            For j = deb To fin
                If chaud(ligch, j) = "" Then
                    If (Abs(chaud(ligch, j - 1) - tinch) < 0.0001) Then
                        k = j - 1
                        j = fin
                    Else
                        k = j - 1
                        tinch = chaud(ligch, j - 1)
                        j = fin
                    End If
                End If
            Next
            colmaxch = k
        End If

        'Calcul de toutch
        If cot1 = 10 Then
            'Utilité : on prends la borne extreme
            k = colreste - 2
        Else
            'Process : itération à partir de colmaxch jusqu'à trouver une autre température
            For j = k + 1 To fin
                If chaud(ligch, j) <> "" Then
                    If (Abs(chaud(ligch, j) - toutch) < 0.0001) Then
                        k = j
                        j = fin
                    Else
                        k = j
                        toutch = chaud(ligch, j)
                        j = fin
                    End If
                End If
            Next
        End If
        colminch = k

        'Calcul des bornes extrèmes
        If cot2 = 5 Then
            deb = colpinch + 2
            fin = colreste - 2
        ElseIf cot2 = 4 Then
            deb = 3
            fin = colpinch
        Else
            deb = 3
            fin = colreste - 2
        End If

        'Calcul de toutfr
        If cot2 = 10 Then
            'Utilité : on prends la borne max du flux
            k = 3
        Else
            'Process : itération jusqu'à trouver un espace libre
            j = deb
            k = j
            For j = deb To fin
                If froid(ligfr, j) = "" Then
                    If (Abs(froid(ligfr, j - 1) - toutfr) < 0.0001) Then
                        k = j - 1
                        j = fin
                    Else
                        k = j - 1
                        toutfr = froid(ligfr, j - 1)
                        j = fin
                    End If
                End If
            Next
        End If
        colmaxfr = k

        'Calcul de tinfr
        If cot2 = 10 Then
            'Utilité : on itère jusqu'à trouver un créneau de libre 'MODIF CRE 02/10/2012   Mise à jour de l'identification de colmax et colmin
            j = 1
            Do While froid(ligfr, colreste - 1 - j) <> ""
                j = j + 1
            Loop
            j = j - 1
            colminfr = colreste - 1 - j
            tinfr = froid(ligfr, colminfr)
        Else
            'Process : itération à partir de colmaxfr pour trouver la fin du vide
            For j = k + 1 To fin
                If froid(ligfr, j) <> "" Then
                    If (Abs(froid(ligfr, j) - tinfr) < 0.0001) Then
                        k = j
                        j = fin
                    Else
                        k = j
                        tinfr = froid(ligfr, j)
                        j = fin
                    End If
                End If
            Next
            colminfr = k
        End If

    End If  'fin du Forced = 3

    'Verification des températures  'MODIF CRE 25/09/2012   Verification des bornes min et max des températures des utilités !
    If cot1 = 10 Then
        If toutch < chaud(ligch, colreste - 2) Then
            toutch = chaud(ligch, colreste - 2)
        End If
    ElseIf cot2 = 10 Then
        If toutfr > froid(ligfr, 3) Then
            toutfr = froid(ligfr, 3)
        End If
    End If

    'Calcul de la puissance échangée
    If colmaxch <> -1 And colminch <> -1 Then
        pch = (tinch - toutch) * chaud(ligch, colreste + 1)
    Else
        pch = 10000000
    End If
    If colminfr <> -1 And colmaxfr <> -1 Then
        pfr = (toutfr - tinfr) * froid(ligfr, colreste + 1)
    Else
        pfr = 10000000
    End If
    puissance = Application.Min(pch, pfr)
    If (Abs(puissance - pch) > 0.000001) Then
        toutch = tinch - puissance / chaud(ligch, colreste + 1)
    ElseIf (Abs(puissance - pfr) > 0.000001) And pfr <> 10000000 Then
        If cot1 = 4 Then
            If chaud(ligch, colreste + 1) > froid(ligfr, colreste + 1) Then
                tinfr = toutfr - puissance / froid(ligfr, colreste + 1)
            Else
                toutfr = tinfr + puissance / froid(ligfr, colreste + 1)
            End If
        Else
            If chaud(ligch, colreste + 1) > froid(ligfr, colreste + 1) Then
                tinfr = toutfr - puissance / froid(ligfr, colreste + 1)
            Else
                toutfr = tinfr + puissance / froid(ligfr, colreste + 1)
            End If
        End If
    ElseIf pfr = 10000000 Then
        toutfr = tinfr + puissance / froid(ligfr, colreste + 1)
    End If

    'Vérification du Pinch et recalcul des températures sous le croisement (théorème de thalès)
    mcpch = chaud(ligch, colreste + 1)
    mcpfr = froid(ligfr, colreste + 1)
    puissance_2 = 0
    If (tinch - tinfr) < pinch_mesure Then
        erreur = erreur & " - échange infaisable"
        ok = False
    Else
        If ((tinch - toutfr) < pinch_mesure * 0.9) Then    'MODIF CRE 21/09/2012   légère modification du test. Permet de construire un échangeur, même si les températures ne sont pas exactement bonnes. Evite d'avoir des besoins restants très faibles, et permet de créer des HEX plus facilement. Attention à la précision quand meme : 0.1 ou 0.001 ou 0.0000001 ?
            'If LBound(chaud()) = 0 And LBound(froid()) = 0 Then    'MODIF CRE 24/09/2012   suppression de ce test pour permettre le calcul aussi avec les Utilités
            erreur = erreur & " -Croisement 'bas'"
            If cot1 = 4 And mcpfr < mcpch Then  'MODIF CRE 11/10/2012   Si les pentes sont bonnes, on fixe Tinch et on recalcule l'échange
                toutfr = tinch - pinch_mesure
                puissance_2 = (toutfr - tinfr) * mcpfr
                toutch = tinch - puissance_2 / mcpch
                puissance = puissance_2
            Else
                puissance_2 = (((tinfr * mcpfr - toutch * mcpch) / (mcpfr - mcpch)) - tinfr) * mcpfr * (1 - pinch_mesure / (toutch - tinfr))
                If puissance_2 < puissance And puissance_2 > 0 Then
                    puissance = puissance_2
                    tinch = toutch + puissance / mcpch
                    toutfr = tinfr + puissance / mcpfr
                Else
                    ok = False
                End If
            End If
            'Else
            '    ok = False
            'End If
        ElseIf ((toutch - tinfr) < pinch_mesure * 0.9) Then    'MODIF CRE 21/09/2012   légère modification du test. Permet de construire un échangeur, même si les températures ne sont pas exactement bonnes. Evite d'avoir des besoins restants très faibles, et permet de créer des HEX plus facilement. Attention à la précision quand meme : 0.1 ou 0.001 ou 0.0000001 ?
            'If LBound(chaud()) = 0 And LBound(froid()) = 0 Then    'MODIF CRE 24/09/2012   suppression de ce test pour permettre le calcul aussi avec les Utilités
            erreur = erreur & " -Croisement 'haut'"
            puissance_2 = (toutfr - ((tinfr * mcpfr - toutch * mcpch) / (mcpfr - mcpch))) * mcpfr * (1 - pinch_mesure / (tinch - toutfr))
            If puissance_2 < puissance And puissance_2 > 0 Then
                puissance = puissance_2
                toutch = tinch - puissance / mcpch
                tinfr = toutfr - puissance / mcpfr
            Else
                ok = False
            End If
            'Else
            '    ok = False
            'End If
        End If
    End If

    'Diverses erreurs
    If (Abs(tinch - toutch) < 0.000001) Then
        erreur = erreur & " - erreur de flux FR"
    End If
    If (Abs(tinfr - toutfr) < 0.000001) Then
        erreur = erreur & " - erreur de flux CH"
    End If

    'Vérification du seuil de puissance pour un échange
    colonneStation = 1 + (pi_FiliereConsideree - 1)
    LigneTab = INDICE_PREMIERE_LIGNE
    seuilPuissance = Feuil63.Range("DonneesPincementSeuilPuissance").Cells(LigneTab, colonneStation).Value
    If (puissance < seuilPuissance) Then    'MODIF CRE 21/09/2010   puissance minimale pour un échangeur (5 kW). initialisation de la variable à déporter dans l'interface ?
        erreur = erreur & " - pb puissance"
        ok = False
    End If

    'Trouver la réelle colonne d'insertion du HEX
    colplacehexch = 0
    colplacehexfr = 0
    If ok = True Then
        Call GetCols(data_flux_hen())

        'Process seul
        If left(chaud(ligch, 1), 1) <> "U" And left(froid(ligfr, 1), 1) <> "U" Then
            
            If cot1 = 45 Or cot2 = 45 Then  'MODIF CRE 17/01/2013   HEX à travers, on fixe la colonne de ce flux et on cherche rapidement pour l'autre
                If cot1 = 45 Then
                    colplacehexch = colpinch + 1
                    If colmaxfr < colpinch Then
                        colplacehexfr = colminfr - 1
                    Else
                        colplacehexfr = colmaxfr + 1
                    End If
                Else
                    colplacehexfr = colpinch + 1
                    If colmaxch < colpinch Then
                        colplacehexch = colminch - 1
                    Else
                        colplacehexch = colmaxch + 1
                    End If
                End If
            Else
                If colmaxch <= colmaxfr And colminfr <= colminch Then    'chaud englobe froid
                    If colminch < colpinch Then
                        'CRE 04/09/2012 : parfois, il calcule colminch = colmaxch = 3. Dans ce cas là, il prenait 3-1=2 comme colonne, et ça plantait tout. Maintenant c'est résolu. Cette correction est à effectuer pour les autres cas à priori. C'est difficile de tester toutes les configurations, mais il y a encore des bugs cachés :(
                        If colminfr = 3 Then
                            colplacehexfr = colminfr + 1
                        Else
                            colplacehexfr = colminfr - 1
                        End If
                        'CRE 04/09/2012 end
                    Else
                        colplacehexfr = colmaxfr + 1
                    End If
                    colplacehexch = colplacehexfr
                ElseIf colmaxfr <= colmaxch And colminch <= colminfr Then  'froid englobe chaud
                    If colminfr < colpinch Then
                        colplacehexfr = colminch - 1
                    Else
                        colplacehexfr = colmaxch + 1
                    End If
                    colplacehexch = colplacehexfr
                ElseIf colmaxch <= colmaxfr And colmaxfr <= colminch And colminch <= colminfr Then    'recouvrement par la gauche
                    If colminfr < colpinch Then
                        colplacehexfr = colminch - 1
                    Else
                        colplacehexfr = colmaxfr + 1
                    End If
                    colplacehexch = colplacehexfr
                ElseIf colmaxfr <= colmaxch And colmaxch <= colminfr And colminfr <= colminch Then    'recouvrement par la droite
                    If colminch < colpinch Then
                        colplacehexfr = colminfr - 1
                    Else
                        colplacehexfr = colmaxch + 1
                    End If
                    colplacehexch = colplacehexfr
                Else    'pas de recouvrement
                    If colmaxch < colpinch Then
                        colplacehexch = colminch - 1
                    Else
                        colplacehexch = colmaxch + 1
                    End If
                    If colmaxfr < colpinch Then
                        colplacehexfr = colminfr - 1
                    Else
                        colplacehexfr = colmaxfr + 1
                    End If
                End If
            End If
            
            'Process et utilités
        Else
            If left(chaud(ligch, 1), 1) <> "U" Then    'utilité froide
                If colminfr < colmaxch Then    'froid avant le procédé chaud
                    If colmaxch < colpinch Then
                        colplacehexch = colminch - 1
                    Else
                        colplacehexch = colmaxch + 1
                    End If
                    colplacehexfr = colminfr - 1
                ElseIf colmaxch <= colminfr And colminfr <= colminch Then   'froid englobé dans chaud
                    If colmaxch < colpinch Then
                        colplacehexch = colminch - 1
                    Else
                        colplacehexch = colmaxch + 1
                    End If
                    colplacehexch = Application.Min(colminch - 1, colminfr - 1)
                    colplacehexfr = colplacehexch
                ElseIf colminch < colminfr Then    'froid après chaud
                    colplacehexch = colminch - 1
                    colplacehexfr = colplacehexch
                End If
            Else
                If colmaxch < colmaxfr Then    'chaud avant le procédé froid
                    colplacehexfr = colmaxfr + 1
                    colplacehexch = colplacehexfr
                ElseIf colmaxfr <= colmaxch And colmaxch <= colminfr Then   'chaud englobé dans froid
                    If colmaxfr < colpinch Then
                        colplacehexfr = colminfr - 1
                    Else
                        colplacehexfr = colmaxfr + 1
                    End If
                    colplacehexfr = Application.Max(colmaxfr + 1, colmaxch + 1)
                    colplacehexch = colplacehexfr
                ElseIf colminfr < colmaxch Then    'chaud après froid
                    If colmaxfr < colpinch Then
                        colplacehexfr = colminfr - 1
                    Else
                        colplacehexfr = colmaxfr + 1
                    End If
                    colplacehexch = colmaxch + 1
                End If
            End If
        End If

    End If
    If colplacehexch = -1 Or colplacehexfr = -1 Then
        ok = False
    End If

    'Insertion des colonnes
    If ok = True Then
        If LBound(chaud()) = 0 And LBound(froid()) = 0 Then
            'MODIF CRE 17/01/2013   Gestion des HEX à travers, seulement pour les HEX process-process
            If cot1 <> 45 Then
                Call InsertCol(chaud(), Chr(64 + colplacehexch + 1), data_flux_hen_size)
            Else
                colplacehexch = colplacehexch - 2
            End If
            Call GetCols(data_flux_hen())
            If colplacehexch <> colplacehexfr Then
                If colplacehexfr > colplacehexch And cot1 <> 45 Then
                    colplacehexfr = colplacehexfr + 4
                ElseIf colplacehexch > colplacehexfr Then
                    colplacehexch = colplacehexch + 4
                End If
                If cot2 <> 45 Then
                    Call InsertCol(froid(), Chr(64 + colplacehexfr + 1), data_flux_hen_size)
                Else
                    colplacehexfr = colplacehexfr - 2
                End If
            End If
        ElseIf LBound(chaud()) = 0 Then
            Call InsertCol(chaud(), Chr(64 + colplacehexch + 1), data_flux_hen_size)
            Call GetCols(data_flux_hen())
            If colplacehexch < colplacehexfr Then
                Call InsertCol(froid(), Chr(64 + colplacehexch + 1), data_flux_hen_size - 4)
                colplacehexfr = colplacehexfr + 4
            Else
                Call InsertCol(froid(), Chr(64 + colplacehexfr + 1), data_flux_hen_size - 4)
            End If
        ElseIf LBound(froid()) = 0 Then
            Call InsertCol(froid(), Chr(64 + colplacehexfr + 1), data_flux_hen_size)
            Call GetCols(data_flux_hen())
            If colplacehexfr < colplacehexch Then
                Call InsertCol(chaud(), Chr(64 + colplacehexfr + 1), data_flux_hen_size - 4)
                colplacehexch = colplacehexch + 4
            Else
                Call InsertCol(chaud(), Chr(64 + colplacehexch + 1), data_flux_hen_size - 4)
            End If
        End If
        Call GetCols(data_flux_hen())
    End If

    'Mise à jour des flux d'utilité si on a du ajuster la température
    If (forced = 2 Or forced = 3) And (LBound(chaud()) <> 0 Or LBound(froid()) <> 0) Then
        If LBound(chaud()) = 0 Then    'MODIF CRE 02/10/2012   Mise à jour de Tinch pour les utilités chaudes pour éviter de laisser un blanc à gauche quand on a du modifier la température d'entrée du flux pour l'échangeur précédent
            If tinfr - tif < -0.001 Then
                k = 1
            End If
        Else
            If (tic - tinch) < -0.001 And colmaxch = 3 Then
                chaud(ligch, colmaxch) = tinch
            End If
        End If
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.12 " & position_erreur

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Public Sub CreerHEX2(chaud() As String, ligch As Integer, cotch As Integer, colplacehexch As Integer, froid() As String, ligfr As Integer, cotfr As Integer, colplacehexfr As Integer, forced As Integer)
    'Ajout du HEX dans la grille de flux : insertion des colonnes, affichage des paramètres du HEX (températures, puissance)

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String
    Dim position_erreur As String

    'chaud                  grille des flux qui contient le flux chaud concerné
    'ligch                  numéro du flux chaud
    'cotch                  coté du flux chaud (4 à gauche, 5 à droite du pinch)
    'colplacehexch          colonne du HEX pour le flux chaud
    'froid                  grille des flux qui contient le flux froid concerné
    'ligfr                  numéro du flux froid
    'cotfr                  coté du flux froid (4 à gauche, 5 à droite du pinch)
    'colplacehexfr          colonne du HEX pour le flux froid
    'forced                 échangeur forcé ou pas : 1 (échangeur Process/Process manuel), 2 (échangeur Process/Process automatique), 3 (échangeur Process/Utilité)

    Dim I As Integer
    Dim j As Integer
    Dim dtlm As Double
    Dim surf As Double
    Dim ok2 As Boolean

    Dim lignech As Integer
    Dim lignefr As Integer
    Dim ligchb As String
    Dim ligfrb As String
    Dim label As String
    
    Dim pch1_temp As Double 'Calcul des puissances échangées pour les HEX en travers du pinch
    Dim pch2_temp As Double
    Dim pfr1_temp As Double
    Dim pfr2_temp As Double
    
    'position de l'erreur : forced (1, 2 ou 3), + les lignes des flux chaud et froid
    position_erreur = forced & "_" & ligch & "_" & ligfr

    'si le flux chaud est une utilité, on force son coté à celui du flux froid. Et inversement
    If cotch = 10 Then
        cotch = cotfr
    ElseIf cotfr = 10 Then
        cotfr = cotch
    End If
    If cotch <> cotfr Then
        erreur = erreur & "Attention, croisement du pinch"  'ce n'est pas une erreur pour la création du HEX, mais un message à l'utilisateur
    End If

    'Températures, puissance
    Call GetCols(data_flux_hen())
    If LBound(chaud()) = 0 And LBound(froid()) = 0 Then
        lignech = chaud(ligch, colreste + 7)
        lignefr = froid(ligfr, colreste + 7)
    ElseIf LBound(chaud()) = 0 And LBound(froid()) <> 0 Then
        lignech = chaud(ligch, colreste + 7)
        lignefr = froid(ligfr, colreste + 7)
    ElseIf LBound(chaud()) <> 0 And LBound(froid()) = 0 Then
        lignech = chaud(ligch, colreste + 7)
        lignefr = froid(ligfr, colreste + 7)
    End If

    'Affinage du numéro de l'autre flux concerné
    If lignech < 10 Then
        ligchb = "0" & lignech
    Else
        ligchb = CStr(lignech)
    End If
    If lignefr < 10 Then
        ligfrb = "0" & lignefr
    Else
        ligfrb = CStr(lignefr)
    End If
    
    'MODIF CRE 17/01/2013 Calcul des puissances échangées de part et d'autre du pinch pour les HEX en travers
    If cotch = 45 Then
        pch1_temp = (tinch - CDbl(chaud(ligch, colpinch))) * CDbl(chaud(ligch, colreste + 1))
        pch2_temp = (CDbl(chaud(ligch, colpinch)) - toutch) * CDbl(chaud(ligch, colreste + 1))
    End If
    If cotfr = 45 Then
        pfr1_temp = (toutfr - CDbl(froid(ligfr, colpinch))) * CDbl(froid(ligfr, colreste + 1))
        pfr2_temp = (CDbl(froid(ligfr, colpinch)) - tinfr) * CDbl(froid(ligfr, colreste + 1))
    End If
    
    'Ecriture des paramètres (puissance, température)
    chaud(ligch, colplacehexch + 2) = "P_" & Round(puissance, 6) & "_" & ligfrb
    froid(ligfr, colplacehexfr + 2) = "P_" & Round(puissance, 6) & "_" & ligchb
    chaud(ligch, colplacehexch + 1) = tinch
    chaud(ligch, colplacehexch + 3) = toutch
    froid(ligfr, colplacehexfr + 3) = tinfr
    froid(ligfr, colplacehexfr + 1) = toutfr

    'Mise à jour des puissances restantes (colreste+4 ou colreste+5)
    If cotch <> 45 Then
        If chaud(ligch, colreste + cotch) <> "" Then
            chaud(ligch, colreste + cotch) = CStr(CDbl(chaud(ligch, colreste + cotch)) - puissance)
        ElseIf cotch = 4 Then
            chaud(ligch, colreste + cotch + 1) = CStr(CDbl(chaud(ligch, colreste + cotch + 1)) - puissance)
        Else
            chaud(ligch, colreste + cotch - 1) = CStr(CDbl(chaud(ligch, colreste + cotch - 1)) - puissance)
        End If
    Else
        chaud(ligch, colreste + 4) = CStr(CDbl(chaud(ligch, colreste + 4)) - pch1_temp)
        chaud(ligch, colreste + 5) = CStr(CDbl(chaud(ligch, colreste + 5)) - pch2_temp)
    End If
    If cotfr <> 45 Then
        If froid(ligfr, colreste + cotfr) <> "" Then
            froid(ligfr, colreste + cotfr) = CStr(CDbl(froid(ligfr, colreste + cotfr)) - puissance)
        ElseIf cotch = 4 Then
            froid(ligfr, colreste + cotfr + 1) = CStr(CDbl(froid(ligfr, colreste + cotfr + 1)) - puissance)
        Else
            froid(ligfr, colreste + cotfr - 1) = CStr(CDbl(froid(ligfr, colreste + cotfr - 1)) - puissance)
        End If
    Else
        froid(ligfr, colreste + 4) = CStr(CDbl(froid(ligfr, colreste + 4)) - pfr1_temp)
        froid(ligfr, colreste + 5) = CStr(CDbl(froid(ligfr, colreste + 5)) - pfr2_temp)
    End If

    'En-tête (numéro de HEN)
    If nbhex <= nbheximposed Then
        label = "HAX_"
    Else
        label = "HEX_"
    End If
    If LBound(chaud()) = 0 And LBound(froid()) = 0 Then
        chaud(0, colplacehexch + 2) = label & nbhex
        froid(0, colplacehexfr + 2) = label & nbhex
    ElseIf LBound(chaud()) = 0 Then
        chaud(0, colplacehexch + 2) = "UFr_" & nbhex
    Else
        froid(0, colplacehexfr + 2) = "UCh_" & nbhex
    End If

    'Ajout des tirets juste à coté des températures
    I = 0
    Do While chaud(ligch, colplacehexch + I) = ""
        I = I - 1
    Loop
    If I = 0 Then
        ok2 = False
    Else
        If Abs(chaud(ligch, colplacehexch + I) - chaud(ligch, colplacehexch + 1)) < 0.000001 Then
            j = 0
            Do While chaud(ligch, colplacehexch + j) = ""
                chaud(ligch, colplacehexch + j) = repere_traits_Clement
                j = j - 1
            Loop
        End If
    End If
    If left(chaud(ligch, 1), 1) <> "U" Then
        I = 0
        Do While chaud(ligch, colplacehexch + 4 + I) = ""
            I = I + 1
        Loop
        If I = 0 Then
            ok2 = False
        Else
            If Abs(chaud(ligch, colplacehexch + 4 + I) - chaud(ligch, colplacehexch + 3)) < 0.000001 Then
                j = 0
                Do While chaud(ligch, colplacehexch + 4 + j) = ""
                    chaud(ligch, colplacehexch + 4 + j) = repere_traits_Clement
                    j = j + 1
                Loop
            End If
        End If
    End If

    If left(froid(ligfr, 1), 1) <> "U" Then
        I = 0
        Do While froid(ligfr, colplacehexfr + I) = ""
            I = I - 1
        Loop
        If I = 0 Then
            ok2 = False
        Else
            If Abs(froid(ligfr, colplacehexfr + I) - froid(ligfr, colplacehexfr + 1)) < 0.000001 Then
                j = 0
                Do While froid(ligfr, colplacehexfr + j) = ""
                    froid(ligfr, colplacehexfr + j) = repere_traits_Clement
                    j = j - 1
                Loop
            End If
        End If
    End If
    I = 0
    Do While froid(ligfr, colplacehexfr + 4 + I) = ""
        I = I + 1
    Loop
    If I = 0 Then
        ok2 = False
    Else
        If Abs(froid(ligfr, colplacehexfr + 4 + I) - froid(ligfr, colplacehexfr + 3)) < 0.000001 Then
            j = 0
            Do While froid(ligfr, colplacehexfr + 4 + j) = ""
                froid(ligfr, colplacehexfr + 4 + j) = repere_traits_Clement
                j = j + 1
            Loop
        End If
    End If

    'Calcul des surfaces d'échange
    If (tinch - toutfr) = (toutch - tinfr) Then
        dtlm = tinch - toutfr
    Else
        dtlm = (Abs(tinch - toutfr) - Abs(toutch - tinfr)) / Log(Abs(tinch - toutfr) / Abs(toutch - tinfr))
    End If
    surf = puissance * 1000 / CoefU / dtlm
    surftot = surftot + surf

    'Stockage des informations des HEX dans une variable
    data_hen_size = data_hen_size + 1
    ReDim Preserve data_hen(1 To 12, 1 To data_hen_size)
    If LBound(chaud()) = 0 And LBound(froid()) = 0 Then
        data_hen(1, data_hen_size) = "HEX_" & nbhex
    ElseIf LBound(chaud()) = 0 Then
        data_hen(1, data_hen_size) = "UFr_" & nbhex
    ElseIf LBound(froid()) = 0 Then
        data_hen(1, data_hen_size) = "UCh_" & nbhex
    End If
    data_hen(2, data_hen_size) = puissance
    data_hen(3, data_hen_size) = chaud(ligch, 2)
    data_hen(4, data_hen_size) = froid(ligfr, 2)
    data_hen(5, data_hen_size) = tinch
    data_hen(6, data_hen_size) = toutch
    data_hen(7, data_hen_size) = tinfr
    data_hen(8, data_hen_size) = toutfr
    data_hen(9, data_hen_size) = surf
    data_hen(10, data_hen_size) = erreur
    'MODIF CRE 24/10/2012 Inutile ici, sera renseigné dans IV.4
    'data_hen(11, data_hen_size) = colplacehexch + 2
    'data_hen(12, data_hen_size) = colplacehexfr + 2

    'MODIF CRE 08/09/2012   Pour vérifier la création des HEX en direct
    '    Tabelle9.Activate
    '    Range("A1:IV100").ClearContents
    '    For i = LBound(froid(), 1) To UBound(froid(), 1)
    '        For j = 1 To UBound(froid(), 2)
    '            Cells(2 + i, 1 + j).Value = Replace(froid(i, j), separateur_decimal, ".")
    '        Next
    '    Next
    '    For i = LBound(chaud(), 1) To UBound(chaud(), 1)
    '        For j = 1 To UBound(chaud(), 2)
    '            Cells(2 + i + UBound(froid()), 1 + j).Value = Replace(chaud(i, j), separateur_decimal, ".")
    '        Next
    '    Next
    '    j = 1

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.11 " & position_erreur

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub ecrire_flux_process()

    'Ajouté par DCA   ne marche que pour des valeurs positives

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    Dim j As Integer
    Dim Colonne As Integer
    Dim LigneTab As Integer

    'on reporte les flux
    I = 1
    LigneTab = 3
    Colonne = flux_thermique_process_nb_caracteristiques * (pi_FiliereConsideree - 1)
    Do While Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab + I, Colonne + repere_flux_thermique_process_etat) <> ""     'Or i <= UBound(flux_thermique_process, 2)
        If I <= UBound(flux_thermique_process, 2) Then
            For j = repere_flux_thermique_process_etat To flux_thermique_process_nb_caracteristiques
                Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab + I, Colonne + j) = flux_thermique_process(j, I)
            Next j
        Else
            Feuil63.Range(Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab + I, Colonne + repere_flux_thermique_process_etat), Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab + I, Colonne + flux_thermique_process_nb_caracteristiques)).ClearContents
        End If
        I = I + 1
    Loop

    Call ecriture_formule_pour_changement_langue_noms_flux(pi_FiliereConsideree)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.24"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Public Sub Sort_Array_Custom(t() As String, sens As String, param As Integer, col As Double)
    'Tri customisé d'un tableau de flux

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    't      tableau à trier
    'sens   ASC ou DSC
    'param  nombre de colonnes du tableau
    'col    critère de tri

    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim tmp As String
    Dim iterat As Integer       'nombre d'itérations
    Dim size As Integer
    Dim temp() As String
    ReDim temp(1 To param)
    Dim cont As Boolean

    size = UBound(t())
    iterat = 0
    If size <= 1 Then iterat = 1000
    cont = False

    'Gestion des cases vides
    For I = 1 To UBound(t())
        If t(I, col) = "" Then
            For k = I To UBound(t())
                If t(k, col) <> "" Then
                    cont = True
                    k = UBound(t())
                End If
            Next
            If cont = True Then
                cont = False
                For k = 1 To param
                    temp(k) = t(I, k)
                    For j = I + 1 To UBound(t())
                        t(j - 1, k) = t(j, k)
                    Next
                    t(UBound(t()), k) = temp(k)
                Next
                I = I - 1
            Else
                size = I - 1
                I = UBound(t())
            End If
        End If
    Next
    If size <= 1 Then
        iterat = 200
    End If

    Do While iterat < 100
        For I = 1 To size - 1
            If sens = "dsc" Then
                If CDbl(t(I + 1, col)) > CDbl(t(I, col)) Then
                    For j = 1 To param
                        tmp = t(I + 1, j)
                        t(I + 1, j) = t(I, j)
                        t(I, j) = tmp
                    Next
                    I = I - 1
                End If
            Else
                If CDbl(t(I + 1, col)) < CDbl(t(I, col)) Then
                    For j = 1 To param
                        tmp = t(I + 1, j)
                        t(I + 1, j) = t(I, j)
                        t(I, j) = tmp
                    Next
                    I = I - 1
                End If
            End If
            iterat = iterat + 1
        Next
    Loop

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.10"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Private Sub Class_Initialize()

    Set feuille_active = ActiveSheet

    If PINCH_DANS_OCEAN = True Then
        marqueur_etape = pi_EtapeCalculConsideree
    Else
        marqueur_etape = 2
    End If

    mn_ChangementLangue = False
    
    Call InitCodeNames

End Sub


Private Sub InitCodeNames()

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    'DCA: on crée une collection avec l'ensemble des feuilles pour pointer avec le codename / appeler en initialisation du module de classe

    Dim F As Worksheet

    For Each F In ThisWorkbook.Worksheets
        CodeNames.Add F, F.CodeName
    Next

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.5"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Private Sub Class_Terminate()

    Dim legraph As ChartObject

    If pn_EffacementResultatsEnCours = False Then
        If PINCH_DANS_OCEAN = True Then
            If mn_ChangementLangue = False Then
                Call representation_reseau_DCA_sauvegarde_donnees(data_hen)
            End If
            Call representation_reseau_DCA
            For Each legraph In CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects
        '        If legraph.Name = "CC" Or legraph.Name = "CC2" Or legraph.Name = "HENGraphique" Then
        '            legraph.Delete
        '        End If
                legraph.Visible = True
            Next legraph
        End If
        PINCH_ORDONNER_FLUX_UTILITES = True
        Call ordonner_flux_utilites
        If PINCH_ORDONNER_FLUX_PROCESS = False And pi_EtapeCalculConsideree = pi_CALCUL_VALEUR_EFFECTIVE Then
            Call ecrire_flux_process
        End If
    Else
        Call effacer_resultats
    End If
    feuille_active.Activate


End Sub

Public Sub effacer_resultats()

    'Ajouté par DCA
    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String
    
    Dim I As Integer
    Dim legraph As ChartObject      'utile pour supprimer les graphes existants sur la page
    Dim forme As Shape
    

    For Each forme In CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects("HENGraphique").Chart.Shapes
        forme.Delete
    Next forme
    For Each legraph In CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).ChartObjects
'        If legraph.Name = "HENGraphique" Then
'            legraph.Chart.Delete
'        End If
        legraph.Visible = False
    Next legraph
    If PINCH_DANS_OCEAN = True Then
        'Besoins chauds
        CodeNames("Feuil63").Range("PINCH_besoins_chauds_totaux_" & marqueur_etape & pi_FiliereConsideree).ClearContents
        CodeNames("Feuil63").Range("PINCH_besoins_chauds_non_couverts_mini_" & marqueur_etape & pi_FiliereConsideree).ClearContents
        CodeNames("Feuil63").Range("PINCH_besoins_chauds_non_couverts_apres_recuperation_" & marqueur_etape & pi_FiliereConsideree).ClearContents
        CodeNames("Feuil63").Range("PINCH_besoins_chauds_non_couverts_apres_utilites_" & marqueur_etape & pi_FiliereConsideree).ClearContents
        'Besoins froids
        CodeNames("Feuil63").Range("PINCH_besoins_froids_totaux_" & marqueur_etape & pi_FiliereConsideree).ClearContents
        CodeNames("Feuil63").Range("PINCH_besoins_froids_non_couverts_mini_" & marqueur_etape & pi_FiliereConsideree).ClearContents
        CodeNames("Feuil63").Range("PINCH_besoins_froids_non_couverts_apres_recuperation_" & marqueur_etape & pi_FiliereConsideree).ClearContents
        CodeNames("Feuil63").Range("PINCH_besoins_froids_non_couverts_apres_utilites_" & marqueur_etape & pi_FiliereConsideree).ClearContents
        'Critères du pincement
        CodeNames("Feuil63").Range("PINCH_pincement_consigne_" & marqueur_etape & pi_FiliereConsideree).ClearContents
        CodeNames("Feuil63").Range("PINCH_pincement_effectif_" & marqueur_etape & pi_FiliereConsideree).ClearContents
            CodeNames("Feuil63").Range("PINCH_pincement_effectif_message_" & marqueur_etape & pi_FiliereConsideree).Cells(1, 1).ClearContents
        CodeNames("Feuil63").Range("PINCH_pincement_point_" & marqueur_etape & pi_FiliereConsideree).ClearContents
        'Flux interdits
        I = 1
        Do While CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_InterdictionEchanges_" & marqueur_etape & pi_FiliereConsideree).Cells(1, I) <> ""
            CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_InterdictionEchanges_" & marqueur_etape & pi_FiliereConsideree).Cells(1, I).ClearContents
            CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_InterdictionEchanges_" & marqueur_etape & pi_FiliereConsideree).Cells(2, I).ClearContents
            CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("PINCH_InterdictionEchanges_" & marqueur_etape & pi_FiliereConsideree).Cells(3, I).ClearContents
            I = I + 1
        Loop
    Else
        CodeNames("Feuil_Tampon" & pi_FiliereConsideree).Range("B3:IV200").ClearContents
        CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("D7:D34").ClearContents
        CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("F31:F34").ClearContents   'MODIF CRE 26/10/2012   ajout de cette ligne
        CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Range("D65:Z66").ClearContents
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.15"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub ordonner_flux_process()

    'Ajouté par DCA   ne marche que pour des valeurs positives
    If version_Developpement = False Then On Error GoTo erreur_non_geree

    Dim NumeroErreurNonGeree As String
    Dim I As Integer
    Dim j As Integer
    Dim Colonne As Integer
    Dim LigneTab As Integer
    Dim nb_flux_process As Integer
    Dim flux_deja_trie() As Boolean
    Dim flux_tries As Boolean
    Dim nb_flux_froids As Integer
    Dim nb_flux_chauds As Integer
    Dim flux_froids() As Variant
    Dim flux_chauds() As Variant
    Dim position_chaud As Integer
    Dim position_froid As Integer
    Dim T_in As Double
    Dim T_out As Double

    nb_flux_process = UBound(flux_thermique_process, 2)

    If nb_flux_process > 0 Then

        ReDim flux_deja_trie(1 To nb_flux_process)
        ReDim flux_chauds(1 To flux_thermique_process_nb_caracteristiques, 1 To 1)
        ReDim flux_froids(1 To flux_thermique_process_nb_caracteristiques, 1 To 1)
        nb_flux_froids = 0
        nb_flux_chauds = 0
        flux_tries = False

        Do Until flux_tries = True
            T_in = -999999
            T_out = -999999
            position_chaud = 0
            position_froid = 0
            For I = 1 To nb_flux_process
                If flux_deja_trie(I) = False Then
                    If flux_thermique_process(repere_flux_thermique_process_T_in_°C, I) < flux_thermique_process(repere_flux_thermique_process_T_out_°C, I) Then
                        'flux froid que l'on réchauffe
                        If flux_thermique_process(repere_flux_thermique_process_T_out_°C, I) > T_out Then
                            position_froid = I
                            T_out = flux_thermique_process(repere_flux_thermique_process_T_out_°C, I)
                        ElseIf flux_thermique_process(repere_flux_thermique_process_T_out_°C, I) = T_out Then
                            If flux_thermique_process(repere_flux_thermique_process_T_in_°C, I) > flux_thermique_process(repere_flux_thermique_process_T_in_°C, position_froid) Then
                                position_froid = I
                                T_out = flux_thermique_process(repere_flux_thermique_process_T_out_°C, I)
                            ElseIf flux_thermique_process(repere_flux_thermique_process_T_in_°C, I) = flux_thermique_process(repere_flux_thermique_process_T_in_°C, position_froid) Then
                                If flux_thermique_process(repere_flux_thermique_process_MCp_kW_°C, I) > flux_thermique_process(repere_flux_thermique_process_MCp_kW_°C, position_froid) Then
                                    position_froid = I
                                    T_out = flux_thermique_process(repere_flux_thermique_process_T_out_°C, I)
                                End If
                            End If
                        End If
                    Else
                        'flux chaud que l'on refroidit
                        If flux_thermique_process(repere_flux_thermique_process_T_in_°C, I) > T_in Then
                            position_chaud = I
                            T_in = flux_thermique_process(repere_flux_thermique_process_T_in_°C, I)
                        ElseIf flux_thermique_process(repere_flux_thermique_process_T_in_°C, I) = T_in Then
                            If flux_thermique_process(repere_flux_thermique_process_T_out_°C, I) > flux_thermique_process(repere_flux_thermique_process_T_out_°C, position_chaud) Then
                                position_chaud = I
                                T_in = flux_thermique_process(repere_flux_thermique_process_T_in_°C, I)
                            ElseIf flux_thermique_process(repere_flux_thermique_process_T_out_°C, I) = flux_thermique_process(repere_flux_thermique_process_T_out_°C, position_chaud) Then
                                If flux_thermique_process(repere_flux_thermique_process_MCp_kW_°C, I) < flux_thermique_process(repere_flux_thermique_process_MCp_kW_°C, position_chaud) Then
                                    position_chaud = I
                                    T_in = flux_thermique_process(repere_flux_thermique_process_T_out_°C, I)
                                End If
                            End If
                        End If
                    End If
                End If
            Next I
            If position_chaud = 0 And position_froid = 0 Then
                flux_tries = True
            Else
                'chaud
                If position_chaud > 0 Then
                    nb_flux_chauds = nb_flux_chauds + 1
                    ReDim Preserve flux_chauds(1 To flux_thermique_process_nb_caracteristiques, 1 To nb_flux_chauds)
                    For j = 1 To flux_thermique_process_nb_caracteristiques
                        flux_chauds(j, nb_flux_chauds) = flux_thermique_process(j, position_chaud)
                    Next j
                    flux_deja_trie(position_chaud) = True
                End If
                'froid
                If position_froid > 0 Then
                    nb_flux_froids = nb_flux_froids + 1
                    ReDim Preserve flux_froids(1 To flux_thermique_process_nb_caracteristiques, 1 To nb_flux_froids)
                    For j = 1 To flux_thermique_process_nb_caracteristiques
                        flux_froids(j, nb_flux_froids) = flux_thermique_process(j, position_froid)
                    Next j
                    flux_deja_trie(position_froid) = True
                End If
            End If
        Loop
    End If

    If process_flux_chauds_premier = True Then
        'chauds
        For I = 1 To nb_flux_chauds
            For j = 1 To flux_thermique_process_nb_caracteristiques
                flux_thermique_process(j, I) = flux_chauds(j, I)
            Next j
        Next I
        'froids
        For I = 1 To nb_flux_froids
            For j = 1 To flux_thermique_process_nb_caracteristiques
                flux_thermique_process(j, nb_flux_chauds + I) = flux_froids(j, I)
            Next j
        Next I
    Else
        'froids
        For I = 1 To nb_flux_froids
            For j = 1 To flux_thermique_process_nb_caracteristiques
                flux_thermique_process(j, I) = flux_froids(j, I)
            Next j
        Next I
        'chauds
        For I = 1 To nb_flux_chauds
            For j = 1 To flux_thermique_process_nb_caracteristiques
                flux_thermique_process(j, nb_flux_froids + I) = flux_chauds(j, I)
            Next j
        Next I
    End If

    If REPORT_FLUX_ORDONNES = True Then
        'on reporte les flux
        I = 1
        LigneTab = 3
        Colonne = flux_thermique_process_nb_caracteristiques * (pi_FiliereConsideree - 1)
        Do While Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab + I, Colonne + repere_flux_thermique_process_etat) <> "" Or I <= UBound(flux_thermique_process, 2)
            If I <= UBound(flux_thermique_process, 2) Then
                For j = repere_flux_thermique_process_etat To flux_thermique_process_nb_caracteristiques
                    Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab + I, Colonne + j) = flux_thermique_process(j, I)
                Next j
            Else
                Feuil63.Range(Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab + I, Colonne + repere_flux_thermique_process_etat), Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab + I, Colonne + flux_thermique_process_nb_caracteristiques)).ClearContents
            End If
            I = I + 1
        Loop
        Call ecriture_formule_pour_changement_langue_noms_flux(pi_FiliereConsideree)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.21"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Public Sub ecriture_formule_pour_changement_langue_noms_flux(ByVal Filiere As Integer)
'DCA 22/05/2014
'Recherche des noms des plages de cellules contenant les noms des flux procédés

'Déclarations
Dim LigneTab As Integer
Dim Colonne As Integer
Dim LigneFluxProcess As Integer
Dim LigneNomsPlagesFlux As Integer
Dim NomFlux As String
Dim NomPlageFlux As String
Dim NomFluxTrouve As Boolean

'Initialisations
LigneTab = 0
Colonne = 0
LigneFluxProcess = 0
LigneNomsPlagesFlux = 0
NomFlux = ""
NomPlageFlux = ""
NomFluxTrouve = False

'On se positionne en fonction de la filière calculée au niveau de la première ligne3.
LigneTab = 3
Colonne = flux_thermique_process_nb_caracteristiques * (Filiere - 1)

'Ajout pour éviter bug incompréhensible lors de l'import/export
If pi_FiliereConsideree <> Filiere Then
    pi_FiliereConsideree = Filiere
End If

LigneFluxProcess = 1
Do While Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab + LigneFluxProcess, Colonne + repere_flux_thermique_process_etat) <> ""
    NomFlux = Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab + LigneFluxProcess, Colonne + repere_flux_thermique_process_nom)
    LigneNomsPlagesFlux = 1
    NomFluxTrouve = False
    Do While Feuil63.Range("PINCH_NomsPlagesFluxProcede").Cells(LigneNomsPlagesFlux, 1) <> "" And NomFluxTrouve = False
        NomPlageFlux = Feuil63.Range("PINCH_NomsPlagesFluxProcede").Cells(LigneNomsPlagesFlux, 1)
        If NomFlux = Feuil5.Range(NomPlageFlux).Cells(1, 1) Then
            Feuil63.Range("DonneesPincementFluxProcedes").Cells(LigneTab + LigneFluxProcess, Colonne + repere_flux_thermique_process_nom).Formula = "= " & NomPlageFlux
            NomFluxTrouve = True
        End If
        LigneNomsPlagesFlux = LigneNomsPlagesFlux + 1
    Loop
    LigneFluxProcess = LigneFluxProcess + 1
Loop


End Sub
Public Sub ecriture_formule_pour_changement_langue_noms_utilites(ByVal Filiere As Integer)
'DCA 22/05/2014
'Recherche des noms des plages de cellules contenant les noms des flux procédés

'Déclarations
Dim LigneTab As Integer
Dim Colonne As Integer
Dim LigneFluxUtilites As Integer
Dim LigneNomsPlagesFlux As Integer
Dim NomFlux As String
Dim NomPlageFlux As String
Dim NomFluxTrouve As Boolean

'Initialisations
LigneTab = 0
Colonne = 0
LigneFluxUtilites = 0
LigneNomsPlagesFlux = 0
NomFlux = ""
NomPlageFlux = ""
NomFluxTrouve = False

'On se positionne en fonction de la filière calculée au niveau de la première ligne3.
LigneTab = 3
Colonne = flux_thermique_utilites_nb_caracteristiques * (Filiere - 1)

'Ajout pour éviter bug incompréhensible lors de l'import/export
If pi_FiliereConsideree <> Filiere Then
    pi_FiliereConsideree = Filiere
End If

LigneFluxUtilites = 1
Do While Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab + LigneFluxUtilites, Colonne + repere_flux_thermique_utilites_etat) <> ""
    NomFlux = Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab + LigneFluxUtilites, Colonne + repere_flux_thermique_utilites_nom)
    LigneNomsPlagesFlux = 1
    NomFluxTrouve = False
    Do While Feuil63.Range("PINCH_NomsPlagesFluxUtilites").Cells(LigneNomsPlagesFlux, 1) <> "" And NomFluxTrouve = False
        NomPlageFlux = Feuil63.Range("PINCH_NomsPlagesFluxUtilites").Cells(LigneNomsPlagesFlux, 1)
        If NomFlux = Feuil41.Range(NomPlageFlux).Cells(1, 1) Then
            Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab + LigneFluxUtilites, Colonne + repere_flux_thermique_utilites_nom).Formula = "= " & NomPlageFlux
            NomFluxTrouve = True
        End If
        LigneNomsPlagesFlux = LigneNomsPlagesFlux + 1
    Loop
    LigneFluxUtilites = LigneFluxUtilites + 1
Loop


End Sub

Sub ordonner_flux_utilites()

    'Ajouté par DCA   ne marche que pour des valeurs positives
    If version_Developpement = False Then On Error GoTo erreur_non_geree

    Dim NumeroErreurNonGeree As String
    Dim I As Integer
    Dim j As Integer
    Dim Colonne As Integer
    Dim LigneTab As Integer
    Dim flux_deja_trie() As Boolean
    Dim flux_tries As Boolean
    Dim nb_flux_froids As Integer
    Dim nb_flux_chauds As Integer
    Dim flux_froids() As Variant
    Dim flux_chauds() As Variant
    Dim position_chaud As Integer
    Dim position_froid As Integer
    Dim T_in_ch As Double
    Dim T_in_fr As Double

    If nb_flux_utilites > 0 Then
        ReDim flux_deja_trie(1 To nb_flux_utilites)
        ReDim flux_chauds(1 To flux_thermique_utilites_nb_caracteristiques, 1 To 1)
        ReDim flux_froids(1 To flux_thermique_utilites_nb_caracteristiques, 1 To 1)
        nb_flux_froids = 0
        nb_flux_chauds = 0
        flux_tries = False
        Do Until flux_tries = True
            T_in_ch = 999999
            T_in_fr = -999999
            position_chaud = 0
            position_froid = 0
            For I = 1 To nb_flux_utilites
                If flux_deja_trie(I) = False Then
                    If utilites_flux_energie(repere_flux_thermique_utilites_type, I) = "Fr" Then
                        'utilités froides
                        If utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, I) > T_in_fr Then
                            position_froid = I
                            T_in_fr = utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, I)
                        ElseIf utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, I) = T_in_fr Then
                            If utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, I) < utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, position_froid) Then
                                position_froid = I
                                T_in_fr = utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, I)
                            ElseIf utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, I) = utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, position_froid) Then
                                If utilites_flux_energie(repere_flux_thermique_utilites_MCp_kW_°C, I) < utilites_flux_energie(repere_flux_thermique_utilites_MCp_kW_°C, position_froid) Then
                                    position_froid = I
                                    T_in_fr = utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, I)
                                End If
                            End If
                        End If
                    Else
                        'utilités chaudes
                        If utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, I) < T_in_ch Then
                            position_chaud = I
                            T_in_ch = utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, I)
                        ElseIf utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, I) = T_in_ch Then
                            If utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, I) < utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, position_chaud) Then
                                position_chaud = I
                                T_in_ch = utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, I)
                            ElseIf utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, I) = utilites_flux_energie(repere_flux_thermique_utilites_Pw_kW, position_chaud) Then
                                If utilites_flux_energie(repere_flux_thermique_utilites_MCp_kW_°C, I) < utilites_flux_energie(repere_flux_thermique_utilites_MCp_kW_°C, position_chaud) Then
                                    position_chaud = I
                                    T_in_ch = utilites_flux_energie(repere_flux_thermique_utilites_T_in_°C, I)
                                End If
                            End If
                        End If
                    End If
                End If
            Next I
            If position_chaud = 0 And position_froid = 0 Then
                flux_tries = True
            Else
                'chaud
                If position_chaud > 0 Then
                    nb_flux_chauds = nb_flux_chauds + 1
                    ReDim Preserve flux_chauds(1 To flux_thermique_utilites_nb_caracteristiques, 1 To nb_flux_chauds)
                    For j = 1 To flux_thermique_utilites_nb_caracteristiques
                        flux_chauds(j, nb_flux_chauds) = utilites_flux_energie(j, position_chaud)
                    Next j
                    flux_deja_trie(position_chaud) = True
                End If
                'froid
                If position_froid > 0 Then
                    nb_flux_froids = nb_flux_froids + 1
                    ReDim Preserve flux_froids(1 To flux_thermique_utilites_nb_caracteristiques, 1 To nb_flux_froids)
                    For j = 1 To flux_thermique_utilites_nb_caracteristiques
                        flux_froids(j, nb_flux_froids) = utilites_flux_energie(j, position_froid)
                    Next j
                    flux_deja_trie(position_froid) = True
                End If
            End If
        Loop
    End If

    If utilites_flux_chauds_premier = True Then
        'chauds
        For I = 1 To nb_flux_chauds
            For j = 1 To flux_thermique_utilites_nb_caracteristiques
                utilites_flux_energie(j, I) = flux_chauds(j, I)
            Next j
        Next I
        'froids
        For I = 1 To nb_flux_froids
            For j = 1 To flux_thermique_utilites_nb_caracteristiques
                utilites_flux_energie(j, nb_flux_chauds + I) = flux_froids(j, I)
            Next j
        Next I
    Else
        'froids
        For I = 1 To nb_flux_froids
            For j = 1 To flux_thermique_utilites_nb_caracteristiques
                utilites_flux_energie(j, I) = flux_froids(j, I)
            Next j
        Next I
        'chauds
        For I = 1 To nb_flux_chauds
            For j = 1 To flux_thermique_utilites_nb_caracteristiques
                utilites_flux_energie(j, nb_flux_froids + I) = flux_chauds(j, I)
            Next j
        Next I
    End If

    If REPORT_FLUX_ORDONNES = True Then
        'On reporte les utilités dans l'onglet d'où le PINCH les lit
        LigneTab = 3
        Colonne = flux_thermique_utilites_nb_caracteristiques * (pi_FiliereConsideree - 1)
        For I = 1 To nb_flux_utilites
            For j = 1 To flux_thermique_utilites_nb_caracteristiques
                Feuil63.Range("DonneesPincementFluxUtilites").Cells(LigneTab + I, Colonne + j) = utilites_flux_energie(j, I)
            Next j
        Next I
    End If
    Call ecriture_formule_pour_changement_langue_noms_utilites(pi_FiliereConsideree)

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.22"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub

Sub representation_reseau_DCA()

    'Ajouté par DCA   transfère les données pour afficher le réseau d'échangeurs

    If version_Developpement = False Then On Error GoTo erreur_non_geree
    Dim NumeroErreurNonGeree As String

    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim Colonne As Integer

    Dim flux_chaud() As Boolean
    Dim position_trouvee As Boolean
    Dim Matrice() As Integer
    Dim matrice_gauche() As Integer
    Dim matrice_droite() As Integer
    Dim nb_echangeur_gauche As Integer
    Dim nb_echangeur_droite As Integer
    Dim nb_flux_process As Integer
    Dim intermediaire_flux() As Integer
    Dim nb_ech As Integer
    nb_echangeurs = data_hen_size - 1

    ReDim flux_chaud(1 To UBound(flux_thermique_process, 2) + nb_flux_utilites)

    If nb_echangeurs > 0 Or nb_puissance_restants > 0 Then
        nb_flux_process = UBound(flux_thermique_process, 2)
        ReDim Matrice(1 To nb_flux_process + nb_flux_utilites, 1 To 1)
        If nb_echangeurs > 0 Then
            ReDim echangeurs(1 To echangeur_nb_caracteristique, 1 To nb_echangeurs)
            For I = 1 To nb_echangeurs
                If I <= nbheximposed Then
                    echangeurs(repere_echangeur_force, I) = True
                Else
                    echangeurs(repere_echangeur_force, I) = False
                End If
                echangeurs(repere_echangeur_FR_nom, I) = data_hen(4, I + 1)
                echangeurs(repere_echangeur_FR_Tin, I) = CDbl(data_hen(7, I + 1))
                echangeurs(repere_echangeur_FR_Tout, I) = CDbl(data_hen(8, I + 1))
                echangeurs(repere_echangeur_CH_nom, I) = data_hen(3, I + 1)
                echangeurs(repere_echangeur_CH_Tin, I) = CDbl(data_hen(5, I + 1))
                echangeurs(repere_echangeur_CH_Tout, I) = CDbl(data_hen(6, I + 1))
                echangeurs(repere_echangeur_puissance, I) = CDbl(data_hen(2, I + 1))
                echangeurs(repere_echangeur_surface, I) = CDbl(data_hen(9, I + 1))
                '        CodeNames("Feuil_Results" & marqueur_etape & pi_FiliereConsideree).Cells(ligne + i, 1).Value = echangeurs(repere_echangeur_FR_nom, i)
            Next I
        Else
            ReDim echangeurs(1 To echangeur_nb_caracteristique, 1 To 1)
        End If
        'on esssaie de créer le tableau qui donne les priorités
        'On parcourt les flux process et classe les échangeurs
        For I = 1 To nb_flux_process
            If flux_thermique_process(repere_flux_thermique_process_T_in_°C, I) < flux_thermique_process(repere_flux_thermique_process_T_out_°C, I) Then
                'Flux froid
                flux_chaud(I) = False
                ReDim intermediaire_flux(1 To 1)
                nb_ech = 0
                For j = 1 To nb_echangeurs
                    If echangeurs(repere_echangeur_FR_nom, j) = flux_thermique_process(repere_flux_thermique_process_nom, I) Then
                        nb_ech = nb_ech + 1
                        If nb_ech = 1 Then
                            intermediaire_flux(1) = j
                        Else
                            ReDim Preserve intermediaire_flux(1 To nb_ech)
                            k = 0
                            position_trouvee = False
                            Do While k < (nb_ech - 1) And position_trouvee = False
                                k = k + 1
                                If echangeurs(repere_echangeur_FR_Tout, j) > echangeurs(repere_echangeur_FR_Tout, intermediaire_flux(k)) Then
                                    position_trouvee = True
                                End If
                            Loop
                            If position_trouvee = True Then
                                For l = 1 To nb_ech - k
                                    If (nb_ech - l) > 0 Then
                                        intermediaire_flux(nb_ech - l + 1) = intermediaire_flux(nb_ech - l)
                                    End If
                                Next l
                                intermediaire_flux(k) = j
                            Else
                                intermediaire_flux(nb_ech) = j
                            End If
                        End If
                    End If
                Next j
                'On fait les puissances restantes
                If nb_puissance_restants > 0 Then
                    For j = 1 To nb_puissance_restants
                        If puissances_non_echangees(P_non_echangee_nom_flux, j) = flux_thermique_process(repere_flux_thermique_process_nom, I) Then
                            nb_ech = nb_ech + 1
                            If nb_ech = 1 Then
                                intermediaire_flux(1) = -j    'on attribue le numéro en négatif
                            Else
                                ReDim Preserve intermediaire_flux(1 To nb_ech)
                                k = 0
                                position_trouvee = False
                                Do While k < (nb_ech - 1) And position_trouvee = False
                                    k = k + 1
                                    If intermediaire_flux(k) > 0 Then
                                        If puissances_non_echangees(P_non_echangee_T_chaude, j) > echangeurs(repere_echangeur_FR_Tout, intermediaire_flux(k)) Then
                                            position_trouvee = True
                                        End If
                                    ElseIf intermediaire_flux(k) < 0 Then
                                        If puissances_non_echangees(P_non_echangee_T_chaude, j) > puissances_non_echangees(P_non_echangee_T_chaude, -intermediaire_flux(k)) Then
                                            position_trouvee = True
                                        End If
                                    End If
                                Loop
                                If position_trouvee = True Then
                                    For l = 1 To nb_ech - k
                                        If (nb_ech - l) > 0 Then
                                            intermediaire_flux(nb_ech - l + 1) = intermediaire_flux(nb_ech - l)
                                        End If
                                    Next l
                                    intermediaire_flux(k) = -j    'on attribue le numéro en négatif
                                Else
                                    intermediaire_flux(nb_ech) = -j    'on attribue le numéro en négatif
                                End If
                            End If
                        End If
                    Next j
                End If
                'On ajoute la limite du PINCH
                nb_ech = nb_ech + 1
                If nb_ech = 1 Then
                    intermediaire_flux(1) = colonne_pinch    'on attribue le numéro en négatif
                Else
                    ReDim Preserve intermediaire_flux(1 To nb_ech)
                    If RecuperationEnergetiquePossible = True Then
                        k = 0
                        position_trouvee = False
                        Do While k < (nb_ech - 1) And position_trouvee = False
                            k = k + 1
                            If intermediaire_flux(k) > 0 Then
                                If (T_pinch_reseau_DCA - pinch_mesure / 2) >= echangeurs(repere_echangeur_FR_Tout, intermediaire_flux(k)) Then
                                    position_trouvee = True
                                End If
                            ElseIf intermediaire_flux(k) < 0 Then
                                If (T_pinch_reseau_DCA - pinch_mesure / 2) >= puissances_non_echangees(P_non_echangee_T_chaude, -intermediaire_flux(k)) Then
                                    position_trouvee = True
                                End If
                            End If
                        Loop
                    Else
                        position_trouvee = True
                        k = 1
                    End If
                    If position_trouvee = True Then
                        For l = 1 To nb_ech - k
                            If (nb_ech - l) > 0 Then
                                intermediaire_flux(nb_ech - l + 1) = intermediaire_flux(nb_ech - l)
                            End If
                        Next l
                        intermediaire_flux(k) = colonne_pinch    'on attribue le numéro en négatif
                    Else
                        intermediaire_flux(nb_ech) = colonne_pinch    'on attribue le numéro en négatif
                    End If
                End If
            Else
                'Flux chaud
                flux_chaud(I) = True
                ReDim intermediaire_flux(1 To 1)
                nb_ech = 0
                For j = 1 To nb_echangeurs
                    If echangeurs(repere_echangeur_CH_nom, j) = flux_thermique_process(repere_flux_thermique_process_nom, I) Then
                        nb_ech = nb_ech + 1
                        If nb_ech = 1 Then
                            intermediaire_flux(1) = j
                        Else
                            ReDim Preserve intermediaire_flux(1 To nb_ech)
                            k = 0
                            position_trouvee = False
                            Do While k < (nb_ech - 1) And position_trouvee = False
                                k = k + 1
                                If echangeurs(repere_echangeur_CH_Tin, j) > echangeurs(repere_echangeur_CH_Tin, intermediaire_flux(k)) Then
                                    position_trouvee = True
                                End If
                            Loop
                            If position_trouvee = True Then
                                For l = 1 To nb_ech - k
                                    If (nb_ech - l) > 0 Then
                                        intermediaire_flux(nb_ech - l + 1) = intermediaire_flux(nb_ech - l)
                                    End If
                                Next l
                                intermediaire_flux(k) = j
                            Else
                                intermediaire_flux(nb_ech) = j
                            End If
                        End If
                    End If
                Next j
                'On fait les puissances restantes
                If nb_puissance_restants > 0 Then
                    For j = 1 To nb_puissance_restants
                        If puissances_non_echangees(P_non_echangee_nom_flux, j) = flux_thermique_process(repere_flux_thermique_process_nom, I) Then
                            nb_ech = nb_ech + 1
                            If nb_ech = 1 Then
                                intermediaire_flux(1) = -j    'on attribue le numéro en négatif
                            Else
                                ReDim Preserve intermediaire_flux(1 To nb_ech)
                                k = 0
                                position_trouvee = False
                                Do While k < (nb_ech - 1) And position_trouvee = False
                                    k = k + 1
                                    If intermediaire_flux(k) > 0 Then
                                        If puissances_non_echangees(P_non_echangee_T_chaude, j) > echangeurs(repere_echangeur_CH_Tin, intermediaire_flux(k)) Then
                                            position_trouvee = True
                                        End If
                                    ElseIf intermediaire_flux(k) < 0 Then
                                        If puissances_non_echangees(P_non_echangee_T_chaude, j) > puissances_non_echangees(P_non_echangee_T_chaude, -intermediaire_flux(k)) Then
                                            position_trouvee = True
                                        End If
                                    End If
                                Loop
                                If position_trouvee = True Then
                                    For l = 1 To nb_ech - k
                                        If (nb_ech - l) > 0 Then
                                            intermediaire_flux(nb_ech - l + 1) = intermediaire_flux(nb_ech - l)
                                        End If
                                    Next l
                                    intermediaire_flux(k) = -j    'on attribue le numéro en négatif
                                Else
                                    intermediaire_flux(nb_ech) = -j    'on attribue le numéro en négatif
                                End If
                            End If
                        End If
                    Next j
                End If
                'On ajoute la limite du PINCH
                nb_ech = nb_ech + 1
                If nb_ech = 1 Then
                    intermediaire_flux(1) = colonne_pinch    'on attribue le numéro en négatif
                Else
                    ReDim Preserve intermediaire_flux(1 To nb_ech)
                    If RecuperationEnergetiquePossible = True Then
                        k = 0
                        position_trouvee = False
                        Do While k < (nb_ech - 1) And position_trouvee = False
                            k = k + 1
                            If intermediaire_flux(k) > 0 Then
                                If (T_pinch_reseau_DCA + pinch_mesure / 2) >= echangeurs(repere_echangeur_CH_Tin, intermediaire_flux(k)) Then
                                    position_trouvee = True
                                End If
                            ElseIf intermediaire_flux(k) < 0 Then
                                If (T_pinch_reseau_DCA + pinch_mesure / 2) >= puissances_non_echangees(P_non_echangee_T_chaude, -intermediaire_flux(k)) Then
                                    position_trouvee = True
                                End If
                            End If
                        Loop
                    Else
                        k = 1
                        position_trouvee = True
                    End If
                    If position_trouvee = True Then
                        For l = 1 To nb_ech - k
                            If (nb_ech - l) > 0 Then
                                intermediaire_flux(nb_ech - l + 1) = intermediaire_flux(nb_ech - l)
                            End If
                        Next l
                        intermediaire_flux(k) = colonne_pinch    'on attribue le numéro en négatif
                    Else
                        intermediaire_flux(nb_ech) = colonne_pinch    'on attribue le numéro en négatif
                    End If
                End If
            End If
            If nb_ech > UBound(Matrice, 2) Then
                ReDim Preserve Matrice(1 To nb_flux_process + nb_flux_utilites, 1 To nb_ech)
            End If
            For j = 1 To nb_ech
                Matrice(I, j) = intermediaire_flux(j)
            Next j
        Next I
        'On parcourt les utilités
        For I = 1 To nb_flux_utilites
            If utilites_flux_energie(repere_flux_thermique_utilites_type, I) = "Fr" Then
                'Flux froid
                flux_chaud(nb_flux_process + I) = False
                ReDim intermediaire_flux(1 To 1)
                nb_ech = 0
                For j = 1 To nb_echangeurs
                    If echangeurs(repere_echangeur_FR_nom, j) = utilites_flux_energie(repere_flux_thermique_utilites_nom, I) Then
                        nb_ech = nb_ech + 1
                        If nb_ech = 1 Then
                            intermediaire_flux(1) = j
                        Else
                            ReDim Preserve intermediaire_flux(1 To nb_ech)
                            k = 0
                            position_trouvee = False
                            Do While k < (nb_ech - 1) And position_trouvee = False
                                k = k + 1
                                If echangeurs(repere_echangeur_FR_Tout, j) > echangeurs(repere_echangeur_FR_Tout, intermediaire_flux(k)) Then
                                    position_trouvee = True
                                End If
                            Loop
                            If position_trouvee = True Then
                                For l = 1 To nb_ech - k
                                    If (nb_ech - l) > 0 Then
                                        intermediaire_flux(nb_ech - l + 1) = intermediaire_flux(nb_ech - l)
                                    End If
                                Next l
                                intermediaire_flux(k) = j
                            Else
                                intermediaire_flux(nb_ech) = j
                            End If
                        End If
                    End If
                Next j
                'On fait les puissances restantes
                If nb_puissance_restants > 0 Then
                    For j = 1 To nb_puissance_restants
                        If puissances_non_echangees(P_non_echangee_nom_flux, j) = utilites_flux_energie(repere_flux_thermique_utilites_nom, I) Then
                            nb_ech = nb_ech + 1
                            If nb_ech = 1 Then
                                intermediaire_flux(1) = -j    'on attribue le numéro en négatif
                            Else
                                ReDim Preserve intermediaire_flux(1 To nb_ech)
                                k = 0
                                position_trouvee = False
                                Do While k < (nb_ech - 1) And position_trouvee = False
                                    k = k + 1
                                    If intermediaire_flux(k) > 0 Then
                                        If puissances_non_echangees(P_non_echangee_T_chaude, j) > echangeurs(repere_echangeur_FR_Tout, intermediaire_flux(k)) Then
                                            position_trouvee = True
                                        End If
                                    ElseIf intermediaire_flux(k) < 0 Then
                                        If puissances_non_echangees(P_non_echangee_T_chaude, j) > puissances_non_echangees(P_non_echangee_T_chaude, -intermediaire_flux(k)) Then
                                            position_trouvee = True
                                        End If
                                    End If
                                Loop
                                If position_trouvee = True Then
                                    For l = 1 To nb_ech - k
                                        If (nb_ech - l) > 0 Then
                                            intermediaire_flux(nb_ech - l + 1) = intermediaire_flux(nb_ech - l)
                                        End If
                                    Next l
                                    intermediaire_flux(k) = -j    'on attribue le numéro en négatif
                                Else
                                    intermediaire_flux(nb_ech) = -j    'on attribue le numéro en négatif
                                End If
                            End If
                        End If
                    Next j
                End If
                'On ajoute la limite du PINCH
                nb_ech = nb_ech + 1
                If nb_ech = 1 Then
                    intermediaire_flux(1) = colonne_pinch    'on attribue le numéro en négatif
                Else
                    ReDim Preserve intermediaire_flux(1 To nb_ech)
                    If RecuperationEnergetiquePossible = True Then
                        k = 0
                        position_trouvee = False
                        Do While k < (nb_ech - 1) And position_trouvee = False
                            k = k + 1
                            If intermediaire_flux(k) > 0 Then
                                If (T_pinch_reseau_DCA + pinch_mesure / 2) >= echangeurs(repere_echangeur_CH_Tin, intermediaire_flux(k)) Then
                                    position_trouvee = True
                                End If
                            ElseIf intermediaire_flux(k) < 0 Then
    '                            If (T_pinch_reseau_DCA - pinch_mesure / 2) >= puissances_non_echangees(P_non_echangee_T_chaude, -intermediaire_flux(k)) Then
    '                                position_trouvee = True
    '                            End If
                            End If
                        Loop
                    Else
                        k = 1
                        position_trouvee = True
                    End If
                    If position_trouvee = True Then
                        For l = 1 To nb_ech - k
                            If (nb_ech - l) > 0 Then
                                intermediaire_flux(nb_ech - l + 1) = intermediaire_flux(nb_ech - l)
                            End If
                        Next l
                        intermediaire_flux(k) = colonne_pinch    'on attribue le numéro en négatif
                    Else
                        intermediaire_flux(nb_ech) = colonne_pinch    'on attribue le numéro en négatif
                    End If
                End If
            Else
                'Flux chaud
                flux_chaud(nb_flux_process + I) = True
                ReDim intermediaire_flux(1 To 1)
                nb_ech = 0
                For j = 1 To nb_echangeurs
                    If echangeurs(repere_echangeur_CH_nom, j) = utilites_flux_energie(repere_flux_thermique_utilites_nom, I) Then
                        nb_ech = nb_ech + 1
                        If nb_ech = 1 Then
                            intermediaire_flux(1) = j
                        Else
                            ReDim Preserve intermediaire_flux(1 To nb_ech)
                            k = 0
                            position_trouvee = False
                            Do While k < (nb_ech - 1) And position_trouvee = False
                                k = k + 1
                                If echangeurs(repere_echangeur_CH_Tin, j) > echangeurs(repere_echangeur_CH_Tin, intermediaire_flux(k)) Then
                                    position_trouvee = True
                                End If
                            Loop
                            If position_trouvee = True Then
                                For l = 1 To nb_ech - k
                                    If (nb_ech - l) > 0 Then
                                        intermediaire_flux(nb_ech - l + 1) = intermediaire_flux(nb_ech - l)
                                    End If
                                Next l
                                intermediaire_flux(k) = j
                            Else
                                intermediaire_flux(nb_ech) = j
                            End If
                        End If
                    End If
                Next j
                'On fait les puissances restantes
                If nb_puissance_restants > 0 Then
                    For j = 1 To nb_puissance_restants
                        If puissances_non_echangees(P_non_echangee_nom_flux, j) = utilites_flux_energie(repere_flux_thermique_utilites_nom, I) Then
                            nb_ech = nb_ech + 1
                            If nb_ech = 1 Then
                                intermediaire_flux(1) = -j    'on attribue le numéro en négatif
                            Else
                                ReDim Preserve intermediaire_flux(1 To nb_ech)
                                k = 0
                                position_trouvee = False
                                Do While k < (nb_ech - 1) And position_trouvee = False
                                    k = k + 1
                                    If intermediaire_flux(k) > 0 Then
                                        If puissances_non_echangees(P_non_echangee_T_chaude, j) > echangeurs(repere_echangeur_CH_Tin, intermediaire_flux(k)) Then
                                            position_trouvee = True
                                        End If
                                    ElseIf intermediaire_flux(k) < 0 Then
                                        If puissances_non_echangees(P_non_echangee_T_chaude, j) > puissances_non_echangees(P_non_echangee_T_chaude, -intermediaire_flux(k)) Then
                                            position_trouvee = True
                                        End If
                                    End If
                                Loop
                                If position_trouvee = True Then
                                    For l = 1 To nb_ech - k
                                        If (nb_ech - l) > 0 Then
                                            intermediaire_flux(nb_ech - l + 1) = intermediaire_flux(nb_ech - l)
                                        End If
                                    Next l
                                    intermediaire_flux(k) = -j    'on attribue le numéro en négatif
                                Else
                                    intermediaire_flux(nb_ech) = -j    'on attribue le numéro en négatif
                                End If
                            End If
                        End If
                    Next j
                End If
                'On ajoute la limite du PINCH
                nb_ech = nb_ech + 1
                If nb_ech = 1 Then
                    intermediaire_flux(1) = colonne_pinch    'on attribue le numéro en négatif
                Else
                    ReDim Preserve intermediaire_flux(1 To nb_ech)
                    If RecuperationEnergetiquePossible = True Then
                        k = 0
                        position_trouvee = False
                        Do While k < (nb_ech - 1) And position_trouvee = False
                            k = k + 1
                            If intermediaire_flux(k) > 0 Then
                                If (T_pinch_reseau_DCA - pinch_mesure / 2) >= echangeurs(repere_echangeur_FR_Tout, intermediaire_flux(k)) Then
                                    position_trouvee = True
                                End If
                            ElseIf intermediaire_flux(k) < 0 Then
    '                            If (T_pinch_reseau_DCA + pinch_mesure / 2) >= puissances_non_echangees(P_non_echangee_T_chaude, -intermediaire_flux(k)) Then
                                    position_trouvee = True
    '                            End If
                            End If
                        Loop
                    Else
                        k = 1
                        position_trouvee = True
                    End If
                    If position_trouvee = True Then
                        For l = 1 To nb_ech - k
                            If (nb_ech - l) > 0 Then
                                intermediaire_flux(nb_ech - l + 1) = intermediaire_flux(nb_ech - l)
                            End If
                        Next l
                        intermediaire_flux(k) = colonne_pinch    'on attribue le numéro en négatif
                    Else
                        intermediaire_flux(nb_ech) = colonne_pinch    'on attribue le numéro en négatif
                    End If
                End If
            End If
            If nb_ech > UBound(Matrice, 2) Then
                ReDim Preserve Matrice(1 To nb_flux_process + nb_flux_utilites, 1 To nb_ech)
            End If
            For j = 1 To nb_ech
                Matrice(nb_flux_process + I, j) = intermediaire_flux(j)
            Next j
        Next I

        'on ajoute une colonne à la fin de matrice avec des 0 (pour savoir quand s'arrête la recherche à droite
        ReDim Preserve Matrice(1 To nb_flux_process + nb_flux_utilites, 1 To UBound(Matrice, 2) + 1)
        For I = 1 To UBound(Matrice, 1)
            Matrice(I, UBound(Matrice, 2)) = 0
        Next I
        
        'On créé la matrice gauche et la matrice droite
        ReDim matrice_gauche(1 To UBound(Matrice, 1), 1 To 1)
        ReDim matrice_droite(1 To UBound(Matrice, 1), 1 To 1)
        For I = 1 To UBound(Matrice, 1)
            Colonne = 1
            'on fait la boucle pour le nombre d'échangeurs à gauche
            nb_echangeur_gauche = 0
            Do Until Matrice(I, Colonne) = 0 Or Matrice(I, Colonne) = colonne_pinch
                nb_echangeur_gauche = nb_echangeur_gauche + 1
                Colonne = Colonne + 1
            Loop
            If nb_echangeur_gauche > 0 Then
                If nb_echangeur_gauche > UBound(matrice_gauche, 2) Then
                    ReDim Preserve matrice_gauche(1 To UBound(Matrice, 1), 1 To nb_echangeur_gauche)
                    For j = 1 To UBound(Matrice, 1)
                        matrice_gauche(j, nb_echangeur_gauche) = 0
                    Next j
                End If
                For j = 1 To nb_echangeur_gauche
                    matrice_gauche(I, j) = Matrice(I, j)
                Next j
            Else
                matrice_gauche(I, 1) = 0
            End If
            Colonne = Colonne + 1
            'on fait la boucle pour le nombre d'échangeurs à droite
            nb_echangeur_droite = 0
            Do Until Matrice(I, Colonne) = 0 Or Matrice(I, Colonne) = colonne_pinch
                nb_echangeur_droite = nb_echangeur_droite + 1
                Colonne = Colonne + 1
            Loop
            If nb_echangeur_droite > 0 Then
                If nb_echangeur_droite > UBound(matrice_droite, 2) Then
                    ReDim Preserve matrice_droite(1 To UBound(Matrice, 1), 1 To nb_echangeur_droite)
                    For j = 1 To UBound(Matrice, 1)
                        matrice_droite(j, nb_echangeur_droite) = 0
                    Next j
                End If
                For j = 1 To nb_echangeur_droite
                    matrice_droite(I, j) = Matrice(I, nb_echangeur_gauche + 1 + j)
                Next j
            Else
                matrice_droite(I, 1) = 0
            End If
        Next I

        'On ordonne les échangeurs des deux matrices gauche et droite si échangeurs il y a
        ReDim Matrice(1 To nb_flux_process + nb_flux_utilites, 1 To 1)
        If UBound(matrice_gauche, 2) = 1 Then
            nb_echangeur_gauche = 0
            For I = 1 To UBound(matrice_gauche, 1)
                nb_echangeur_gauche = nb_echangeur_gauche + Abs(matrice_gauche(I, 1))
            Next I
        Else
            nb_echangeur_gauche = 1
        End If
        If nb_echangeur_gauche > 0 Then
            Call ordonner_echangeur(matrice_gauche, flux_chaud)
            ReDim Preserve Matrice(1 To nb_flux_process + nb_flux_utilites, 1 To UBound(matrice_gauche, 2) + 1)
            For I = 1 To nb_flux_process + nb_flux_utilites
                For j = 1 To UBound(matrice_gauche, 2)
                    Matrice(I, j) = matrice_gauche(I, j)
                Next j
            Next I
        End If
        If RecuperationEnergetiquePossible = True Then
            If UBound(Matrice, 2) > 1 Then
                For I = 1 To nb_flux_process + nb_flux_utilites
                    Matrice(I, UBound(matrice_gauche, 2) + 1) = colonne_pinch
                Next I
            Else
                For I = 1 To nb_flux_process + nb_flux_utilites
                    Matrice(I, 1) = colonne_pinch
                Next I
            End If
        End If
        If UBound(matrice_droite, 2) = 1 Then
            nb_echangeur_droite = 0
            For I = 1 To UBound(matrice_droite, 1)
                nb_echangeur_droite = nb_echangeur_droite + Abs(matrice_droite(I, 1))
            Next I
        Else
            nb_echangeur_droite = 1
        End If
        If nb_echangeur_droite > 0 Then
            Call ordonner_echangeur(matrice_droite, flux_chaud)
            If RecuperationEnergetiquePossible = True Then
                ReDim Preserve Matrice(1 To nb_flux_process + nb_flux_utilites, 1 To UBound(matrice_gauche, 2) + 1 + UBound(matrice_droite, 2))
                For I = 1 To nb_flux_process + nb_flux_utilites
                    For j = 1 To UBound(matrice_droite, 2)
                        Matrice(I, UBound(matrice_gauche, 2) + 1 + j) = matrice_droite(I, j)
                    Next j
                Next I
            Else
                ReDim Preserve Matrice(1 To nb_flux_process + nb_flux_utilites, 1 To UBound(matrice_droite, 2))
                For I = 1 To nb_flux_process + nb_flux_utilites
                    For j = 1 To UBound(matrice_droite, 2)
                        Matrice(I, j) = matrice_droite(I, j)
                    Next j
                Next I
            End If
        End If


'                'On écrit pour visualiser
'                For i = 1 To 20
'                    For j = 1 To 20
'                        Tabelle9.Cells(i, j) = ""
'                    Next j
'                Next i
'                For i = 1 To UBound(matrice, 1)
'                    For j = 1 To UBound(matrice, 2)
'                        If matrice(i, j) <> 0 Then
'                            Tabelle9.Cells(i, j) = matrice(i, j)
'                        End If
'                    Next j
'                Next i

        'On sauvegarde les tableaux flux_chaud et Matrice dans Excel
'        Call representation_reseau_DCA_sauvegarde_donnees(Matrice, flux_chaud)
        'On peut enfin se lancer dans la représentation graphique
        Call ecrire_reseau_DCA(Matrice, flux_chaud)
    End If

    Exit Sub
    '*********************************************************************
erreur_non_geree:
    NumeroErreurNonGeree = "63.20"

    Call gestion_erreur_non_geree(NumeroErreurNonGeree)
    pn_EffacementResultatsEnCours = True

End Sub


Private Sub representation_reseau_DCA_sauvegarde_donnees(ByRef MatriceDataHen() As String)    ', ByRef flux_chaud() As Boolean)
'Ajout DCA 24/05/2014 pour actualisation en cas de changement de langue

Dim Ligne As Long
Dim Colonne As Long

'On efface les données précédentes
Ligne = 1
Colonne = 1
Do While Feuil63.Range("PINCH_RepresentationDCAMatrice_" & marqueur_etape & pi_FiliereConsideree).Cells(Ligne, Colonne) <> ""
    Do While Feuil63.Range("PINCH_RepresentationDCAMatrice_" & marqueur_etape & pi_FiliereConsideree).Cells(Ligne, Colonne) <> ""
        Feuil63.Range("PINCH_RepresentationDCAMatrice_" & marqueur_etape & pi_FiliereConsideree).Cells(Ligne, Colonne).ClearContents
        Ligne = Ligne + 1
    Loop
    Ligne = 1
    Colonne = Colonne + 1
Loop

'On écrit les données à sauvegarder
For Colonne = LBound(MatriceDataHen, 2) To UBound(MatriceDataHen, 2)
    For Ligne = LBound(MatriceDataHen, 1) To UBound(MatriceDataHen, 1)
        Feuil63.Range("PINCH_RepresentationDCAMatrice_" & marqueur_etape & pi_FiliereConsideree).Cells(Ligne, Colonne) = "'" & MatriceDataHen(Ligne, Colonne)
    Next Ligne
Next Colonne

'on remplace les noms de flux par des formules
Call ecriture_formule_pour_changement_langue_flux_et_utilites_dans_plage_cellules("PINCH_RepresentationDCAMatrice_" & marqueur_etape & pi_FiliereConsideree)

End Sub

Private Sub ordonner_echangeur(ByRef Matrice() As Integer, ByRef flux_chaud() As Boolean)


    Dim dimension_ligne As Integer
    Dim dimension_colonne_ini As Integer

    Dim nb_echangeurs As Integer
    Dim nb_puissance_restante_ordonnee As Integer
    Dim colonne_choisie As Integer
    Dim ligne_choisie As Integer

    Dim numero_echangeur As Integer

    Dim repere_colonne_pinch As Boolean

    Dim echangeur_ordonne() As Boolean
    ReDim echangeur_ordonne(1 To 1)
    echangeur_ordonne(1) = False
    Dim puissance_restante_ordonnee() As Boolean
    ReDim puissance_restante_ordonnee(1 To 1)
    puissance_restante_ordonnee(1) = False
    Dim I As Integer
    Dim j As Integer

    Dim Ligne As Integer
    Dim Colonne As Integer

    Dim ligne_binome As Integer
    Dim colonne_binome As Integer

    Dim nb_colonnes_ajoutees As Integer
    Dim nb_total_colonnes_ajoutees As Integer

    Dim nb_echangeur_gauche As Integer
    Dim nb_echangeur_droite As Integer


    'On ajoute une colonne vierge à gauche et une à droite
    ReDim Preserve Matrice(1 To UBound(Matrice, 1), 1 To UBound(Matrice, 2) + 2)
    For I = 1 To UBound(Matrice, 1)
        For j = 1 To UBound(Matrice, 2) - 2
            Matrice(I, UBound(Matrice, 2) - j) = Matrice(I, UBound(Matrice, 2) - 1 - j)
        Next j
        Matrice(I, 1) = 0
        Matrice(I, UBound(Matrice, 2)) = 0
    Next I

    dimension_ligne = UBound(Matrice, 1)
    dimension_colonne_ini = UBound(Matrice, 2) - 1

    'On passe en revue les colonnes de droite à gauche

    colonne_choisie = dimension_colonne_ini

debut:
    Do While colonne_choisie > 1
        'On passe en revue les lignes de haut en bas et on recherche une valeur non nulle
        ligne_choisie = 1
        repere_colonne_pinch = False
        Do While Matrice(ligne_choisie, colonne_choisie) = 0 And ligne_choisie < dimension_ligne
            ligne_choisie = ligne_choisie + 1
            If Matrice(ligne_choisie, colonne_choisie) = colonne_pinch Then
                repere_colonne_pinch = True
            End If
        Loop
        'On récupère le numéro de l'échangeur ainsi que l'indice d'échangeur maximal rencontré jusque là
        numero_echangeur = Matrice(ligne_choisie, colonne_choisie)
        nb_echangeurs = UBound(echangeur_ordonne, 1)
        nb_puissance_restante_ordonnee = UBound(puissance_restante_ordonnee, 1)
        'Si le numéro est supérieur, on agrandit le vecteur contenant l'information des échangeurs triés
        If numero_echangeur > nb_echangeurs Then
            ReDim Preserve echangeur_ordonne(1 To numero_echangeur)
            For I = nb_echangeurs + 1 To numero_echangeur
                echangeur_ordonne(I) = False
            Next I
        ElseIf (numero_echangeur = 0 And repere_colonne_pinch = True) Or numero_echangeur = colonne_pinch Then
            colonne_choisie = colonne_choisie - 1
            GoTo debut
        ElseIf numero_echangeur = 0 Then
            'On supprime la colonne
            For Ligne = 1 To ligne_choisie
                For Colonne = colonne_choisie + 1 To UBound(Matrice, 2)
                    Matrice(Ligne, Colonne - 1) = Matrice(Ligne, Colonne)
                Next Colonne
            Next Ligne
            ReDim Preserve Matrice(1 To dimension_ligne, 1 To UBound(Matrice, 2) - 1)
            colonne_choisie = colonne_choisie - 1
            GoTo debut
        ElseIf numero_echangeur < 0 And Abs(numero_echangeur) > nb_puissance_restante_ordonnee Then
            ReDim Preserve puissance_restante_ordonnee(1 To Abs(numero_echangeur))
            For I = nb_puissance_restante_ordonnee + 1 To Abs(numero_echangeur)
                puissance_restante_ordonnee(I) = False
            Next I
        End If
        If numero_echangeur > 0 Then
            'On regarde si l'échangeur a déjà été trié
            If echangeur_ordonne(numero_echangeur) = False Then
                'On recherche oú se situe le binôme
                Ligne = 0
                Do
                    Ligne = Ligne + 1
                    'Il ne peut pas être sur la même ligne
                    If Ligne = ligne_choisie Then
                        Ligne = Ligne + 1
                    End If
                    Colonne = 0
                    Do
                        Colonne = Colonne + 1
                    Loop While Matrice(Ligne, Colonne) <> numero_echangeur And Colonne < colonne_choisie
                Loop While Matrice(Ligne, Colonne) <> numero_echangeur And Ligne < dimension_ligne
                If Matrice(Ligne, Colonne) = numero_echangeur Then
                    ligne_binome = Ligne
                    colonne_binome = Colonne
                    'On regarde le nombre d'échangeurs à droite du binôme pour savoir combien de colonnes ajouter
                    nb_colonnes_ajoutees = 0
                    nb_total_colonnes_ajoutees = 0
                    If colonne_binome < colonne_choisie Then
                        For Colonne = colonne_binome + 1 To colonne_choisie
                            If Matrice(ligne_binome, Colonne) <> 0 Then
                                nb_colonnes_ajoutees = nb_colonnes_ajoutees + 1
                            End If
                        Next Colonne
                        If nb_colonnes_ajoutees > 0 Then
                            'On agrandit le tableau "matrice"
                            ReDim Preserve Matrice(1 To dimension_ligne, 1 To UBound(Matrice, 2) + nb_colonnes_ajoutees)
                            For Ligne = 1 To dimension_ligne
                                'On décale tout ce qui se trouvait à droite
                                If colonne_choisie < UBound(Matrice, 2) - nb_colonnes_ajoutees Then
                                    For I = -UBound(Matrice, 2) To -(colonne_choisie + 1 + nb_colonnes_ajoutees)
                                        Colonne = Abs(I)
                                        Matrice(Ligne, Colonne) = Matrice(Ligne, Colonne - nb_colonnes_ajoutees)
                                    Next I
                                End If
                                For I = 1 To nb_colonnes_ajoutees
                                    If Ligne = ligne_binome Then
                                        Matrice(Ligne, colonne_choisie + I) = Matrice(Ligne, colonne_binome + I)
                                        Matrice(Ligne, colonne_binome + I) = 0
                                        'Attention, il se peut que l'on enlève la seule valeur de la colonne (colonne_binome+i) ce qu'il veut dire qu'il faudra tout vérifier à la fin
                                    Else
                                        Matrice(Ligne, colonne_choisie + I) = 0
                                    End If
                                Next I
                            Next Ligne
                        End If
                    End If
                    nb_total_colonnes_ajoutees = nb_total_colonnes_ajoutees + nb_colonnes_ajoutees
                    'On regarde s'il faut ajouter une colonne pour décaler le point choisi
                    I = 0    'nombre de valeurs non nulles dans la colonne de l'échangeur choisi
                    For Ligne = 1 To dimension_ligne
                        If Matrice(Ligne, colonne_choisie) <> 0 And Matrice(Ligne, colonne_choisie) <> numero_echangeur Then
                            I = I + 1
                        End If
                    Next Ligne
                    If I > 0 Then
                        nb_colonnes_ajoutees = 1
                        ReDim Preserve Matrice(1 To dimension_ligne, 1 To UBound(Matrice, 2) + nb_colonnes_ajoutees)
                        For Ligne = 1 To dimension_ligne
                            'On décale tout ce qui se trouvait à droite
                            If colonne_choisie < UBound(Matrice, 2) - nb_colonnes_ajoutees Then
                                For I = -UBound(Matrice, 2) To -(colonne_choisie + 1 + nb_colonnes_ajoutees)
                                    Colonne = Abs(I)
                                    Matrice(Ligne, Colonne) = Matrice(Ligne, Colonne - nb_colonnes_ajoutees)
                                Next I
                            End If
                            For I = 1 To nb_colonnes_ajoutees
                                If Ligne = ligne_binome Then
                                    Matrice(Ligne, colonne_choisie + I) = Matrice(Ligne, colonne_binome)
                                    Matrice(Ligne, colonne_binome) = 0
                                    'Attention, il se peut que l'on enlève la seule valeur de la colonne (colonne_binome+i) ce qu'il veut dire qu'il faudra tout vérifier à la fin
                                ElseIf Ligne = ligne_choisie Then
                                    Matrice(Ligne, colonne_choisie + I) = Matrice(Ligne, colonne_choisie)
                                    Matrice(Ligne, colonne_choisie) = 0
                                Else
                                    Matrice(Ligne, colonne_choisie + I) = 0
                                End If
                            Next I
                        Next Ligne
                    Else
                        nb_colonnes_ajoutees = 0
                        If colonne_binome <> colonne_choisie Then
                            Matrice(ligne_binome, colonne_choisie) = Matrice(ligne_binome, colonne_binome)
                            Matrice(ligne_binome, colonne_binome) = 0
                        End If
                    End If
                    nb_total_colonnes_ajoutees = nb_total_colonnes_ajoutees + nb_colonnes_ajoutees
                    'On passe à la suite
                    echangeur_ordonne(numero_echangeur) = True
                    colonne_choisie = colonne_choisie + nb_total_colonnes_ajoutees   'UBound(matrice, 2)
                Else
                    'pas de binôme
                    I = 0    'nombre de valeurs non nulles dans la colonne de l'échangeur choisi
                    For Ligne = 1 To dimension_ligne
                        If Matrice(Ligne, colonne_choisie) <> 0 And Matrice(Ligne, colonne_choisie) <> numero_echangeur Then
                            I = I + 1
                        End If
                    Next Ligne
                    If I > 0 Then
                        nb_colonnes_ajoutees = 1
                        ReDim Preserve Matrice(1 To dimension_ligne, 1 To UBound(Matrice, 2) + nb_colonnes_ajoutees)
                        For Ligne = 1 To dimension_ligne
                            'On décale tout ce qui se trouvait à droite
                            If colonne_choisie < UBound(Matrice, 2) - nb_colonnes_ajoutees Then
                                For I = -UBound(Matrice, 2) To -(colonne_choisie + 1 + nb_colonnes_ajoutees)
                                    Colonne = Abs(I)
                                    Matrice(Ligne, Colonne) = Matrice(Ligne, Colonne - nb_colonnes_ajoutees)
                                Next I
                            End If
                            For I = 1 To nb_colonnes_ajoutees
                                If Ligne = ligne_choisie Then
                                    Matrice(Ligne, colonne_choisie + I) = Matrice(Ligne, colonne_choisie)
                                    Matrice(Ligne, colonne_choisie) = 0
                                Else
                                    Matrice(Ligne, colonne_choisie + I) = 0
                                End If
                            Next I
                        Next Ligne
                    Else
                        nb_colonnes_ajoutees = 0
                    End If
                    nb_total_colonnes_ajoutees = nb_colonnes_ajoutees
                    'On passe à la suite
                    echangeur_ordonne(numero_echangeur) = True
                    colonne_choisie = colonne_choisie + nb_total_colonnes_ajoutees   'UBound(matrice, 2)
                End If
            Else   's'il a déjà été trié, on passe à la colonne suivante
                colonne_choisie = colonne_choisie - 1
            End If
        ElseIf numero_echangeur < 0 And numero_echangeur <> colonne_pinch Then
            'On regarde si la puissance restante a déjà été trié
            If puissance_restante_ordonnee(Abs(numero_echangeur)) = False Then
                'On regarde s'il y a des contraintes à gauche ou à droite
                nb_echangeur_gauche = 0
                For j = 1 To colonne_choisie - 1
                    nb_echangeur_gauche = nb_echangeur_gauche + Abs(Matrice(ligne_choisie, j))
                Next j
                nb_echangeur_droite = 0
                For j = colonne_choisie + 1 To UBound(Matrice, 2)
                    nb_echangeur_droite = nb_echangeur_droite + Abs(Matrice(ligne_choisie, j))
                Next j
                'On ordonne le flux
                If nb_echangeur_gauche = 0 Or nb_echangeur_droite = 0 Then
                    If flux_chaud(ligne_choisie) = True Then
                        'flux chaud
                        If nb_echangeur_droite = 0 Then
                            Matrice(ligne_choisie, UBound(Matrice, 2)) = Matrice(ligne_choisie, colonne_choisie)
                            Matrice(ligne_choisie, colonne_choisie) = 0
                        Else
                            Matrice(ligne_choisie, 1) = Matrice(ligne_choisie, colonne_choisie)
                            Matrice(ligne_choisie, colonne_choisie) = 0
                        End If
                    Else
                        'flux froid
                        If nb_echangeur_gauche = 0 Then
                            Matrice(ligne_choisie, 1) = Matrice(ligne_choisie, colonne_choisie)
                            Matrice(ligne_choisie, colonne_choisie) = 0
                        Else
                            Matrice(ligne_choisie, UBound(Matrice, 2)) = Matrice(ligne_choisie, colonne_choisie)
                            Matrice(ligne_choisie, colonne_choisie) = 0
                        End If
                    End If
                Else
                    I = 0    'nombre de valeurs non nulles dans la colonne de l'échangeur choisi
                    For Ligne = 1 To dimension_ligne
                        If Matrice(Ligne, colonne_choisie) <> 0 And Matrice(Ligne, colonne_choisie) <> numero_echangeur Then
                            I = I + 1
                        End If
                    Next Ligne
                    If I > 0 Then
                        nb_colonnes_ajoutees = 1
                        ReDim Preserve Matrice(1 To dimension_ligne, 1 To UBound(Matrice, 2) + nb_colonnes_ajoutees)
                        For Ligne = 1 To dimension_ligne
                            'On décale tout ce qui se trouvait à droite
                            If colonne_choisie < UBound(Matrice, 2) - nb_colonnes_ajoutees Then
                                For I = -UBound(Matrice, 2) To -(colonne_choisie + 1 + nb_colonnes_ajoutees)
                                    Colonne = Abs(I)
                                    Matrice(Ligne, Colonne) = Matrice(Ligne, Colonne - nb_colonnes_ajoutees)
                                Next I
                            End If
                            For I = 1 To nb_colonnes_ajoutees
                                If Ligne = ligne_choisie Then
                                    Matrice(Ligne, colonne_choisie + I) = Matrice(Ligne, colonne_choisie)
                                    Matrice(Ligne, colonne_choisie) = 0
                                Else
                                    Matrice(Ligne, colonne_choisie + I) = 0
                                End If
                            Next I
                        Next Ligne
                    Else
                        nb_colonnes_ajoutees = 0
                    End If
                    nb_total_colonnes_ajoutees = nb_colonnes_ajoutees
                    'On passe à la suite
                    puissance_restante_ordonnee(Abs(numero_echangeur)) = True
                    colonne_choisie = colonne_choisie + nb_total_colonnes_ajoutees    'UBound(matrice, 2)
                End If
            Else   's'il a déjà été trié, on passe à la colonne suivante
                colonne_choisie = colonne_choisie - 1
            End If
        End If
    Loop

    'On supprime la colonne finale si elle est vide (pas de puissance restante à droite
    j = 0
    For I = 1 To UBound(Matrice, 1)
        j = j + Abs(Matrice(I, UBound(Matrice, 2)))
    Next I
    If j = 0 Then
        ReDim Preserve Matrice(1 To UBound(Matrice, 1), 1 To UBound(Matrice, 2) - 1)
    End If

    'On supprime la colonne initiale si elle est vide
    j = 0
    For I = 1 To UBound(Matrice, 1)
        j = j + Abs(Matrice(I, 1))
    Next I
    If j = 0 Then
        For Ligne = 1 To UBound(Matrice, 1)
            For Colonne = 1 To UBound(Matrice, 2) - 1
                Matrice(Ligne, Colonne) = Matrice(Ligne, Colonne + 1)
            Next Colonne
        Next Ligne
        ReDim Preserve Matrice(1 To UBound(Matrice, 1), 1 To UBound(Matrice, 2) - 1)
    End If


End Sub






